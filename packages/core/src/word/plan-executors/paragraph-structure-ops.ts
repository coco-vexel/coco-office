/**
 * 段落结构 executor —— #29(insertParagraphBefore/After / deleteParagraph)。
 *
 * 设计要点:
 *   - insertParagraph:先查 anchor,再分配新 paraId(generateParaId),克隆或新建 pPr,
 *     追加 runs,DOM insertBefore / appendChild;最后更新 result.insertedParagraphs
 *   - deleteParagraph:先查段,检测 bookmarkStart / commentRangeStart 等活锚点;
 *     无锚点直接 removeChild;有锚点按 onAnchorConflict 抛错
 *   - trackChanges=true:两者都抛 UNSUPPORTED_OOXML_FEATURE(#31 负责包装)
 *   - 新 paraId 用 generateParaId(seen) —— seen 集合由当前 paraIdx.keys() 构建,
 *     保证 package 全局唯一;apply 后 kernel.applyPlan 的 finally 块重建索引
 */
import { getPart, isXmlPart, listParts, markDirty } from '../../pack/index.js';
import { generateParaId } from '../para-id.js';
import { W_NS, W14_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { findStoryFor, listStories } from '../stories.js';
import { WordError } from '../word-error.js';
import { createCrossParagraphBookmark, resolveAppendBookmarkName } from './bookmark-ops.js';
import { cleanupEmptyCommentsPart } from './comment-ops.js';
import type { ExecutorContext } from './executor-context.js';
import { appendTextAsRuns, createW, firstDirectChildW } from './_shared.js';
import {
  allocateRevisionId,
  ensureParaMarkDelete,
  ensureParaMarkInsert,
  wrapRunsAsDelete,
  wrapRunsAsInsert,
  type RevisionAttrs,
} from './revision-wrappers.js';

const COMMENTS_PART_PATH = 'word/comments.xml';

// ─────────── insertParagraphBefore / insertParagraphAfter ───────────

export function executeInsertParagraph(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'insertParagraphBefore' | 'insertParagraphAfter' }>,
  position: 'before' | 'after',
): void {
  // 1. 查 anchor
  const anchor = ctx.indices.paraIdx.get(op.anchorParaId.toUpperCase());
  if (!anchor) {
    throw new WordError(
      'PARA_NOT_FOUND',
      `paragraph with paraId='${op.anchorParaId}' not found (op: ${op.type})`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: op.type },
    );
  }

  const doc = anchor.ownerDocument;
  if (!doc)
    throw new WordError('PACKAGE_CORRUPT', 'anchor paragraph has no ownerDocument', ctx.docId);

  // 2. 分配新 paraId(从当前 paraIdx key 集合作为 seen)
  const seen = new Set(ctx.indices.paraIdx.keys());
  const newId = generateParaId(seen);

  // 3. 创建 <w:p> 并设 w14:paraId
  const newPara = doc.createElementNS(W_NS, 'w:p');
  newPara.setAttributeNS(W14_NS, 'w14:paraId', newId);

  // 4. 生成 pPr
  //    styleSource 优先级最高:存在时克隆 styleSource 段的 pPr(覆盖 style 语义)
  const styleSource = op.spec.styleSource;
  if (styleSource !== undefined) {
    // 查 styleSource 段
    const sourcePara = ctx.indices.paraIdx.get(styleSource.toUpperCase());
    if (!sourcePara) {
      // 健壮性降级（绝不崩）：styleSource 段在产物里不存在（典型成因：跨文档 / 重建后
      // 该样式源段已被删 / 文本框内段被当 styleSource 但不在 body 重建落点）。此前直接
      // 抛 PARA_NOT_FOUND 会回滚整个 plan，连累整篇 merge 崩（mam-text-box 文本框文档
      // 实测：styleSource paraId 每次不同、merge 按 paraId 找代表段 miss → CRASH）。
      // 降级：回退到 anchor 段的 pPr（anchor 是真实存在的相邻段，是最自然的样式源兜底），
      // 等价于 styleSource===undefined 的「继承 anchor」分支；anchor 无 pPr 则不加（用文档默认）。
      // 绝不崩、不丢段（新段照常插入，只是样式从精确 styleSource 退化为相邻段样式）。
      console.warn(
        `[word] insertParagraph: styleSource paragraph paraId='${styleSource}' not found; ` +
          `degrade to inherit anchor pPr (op: ${op.type}, docId: ${ctx.docId})`,
      );
      const anchorPPr = firstDirectChildW(anchor, 'pPr');
      if (anchorPPr) {
        newPara.appendChild(anchorPPr.cloneNode(true));
      }
    } else {
      // 克隆 styleSource 段的 pPr;源段无 pPr 则退化不加(不报错)
      const sourcePPr = firstDirectChildW(sourcePara, 'pPr');
      if (sourcePPr) {
        newPara.appendChild(sourcePPr.cloneNode(true));
      }
    }
  } else {
    const style = op.spec.style;
    if (style === undefined || style === 'inherit') {
      // 继承 anchor 的 pPr(深拷贝),anchor 无 pPr 时不加
      const anchorPPr = firstDirectChildW(anchor, 'pPr');
      if (anchorPPr) {
        newPara.appendChild(anchorPPr.cloneNode(true));
      }
    } else {
      // 指定 styleId:新建 pPr + pStyle
      const pPr = createW(doc, 'pPr');
      const pStyle = createW(doc, 'pStyle');
      // w:val 属性在 W_NS 下(和 <w:pStyle> 同命名空间)
      pStyle.setAttributeNS(W_NS, 'w:val', style);
      pPr.appendChild(pStyle);
      newPara.appendChild(pPr);
    }
  }

  // 4b. outlineLevel：写入/覆盖 <w:outlineLvl w:val="N"/> 到 pPr
  if (op.spec.outlineLevel !== undefined) {
    const lvl = op.spec.outlineLevel;
    if (lvl < 0 || lvl > 8 || !Number.isInteger(lvl)) {
      throw new WordError(
        'INVALID_ARGUMENT',
        `insertParagraph: outlineLevel must be an integer 0–8, got ${lvl}`,
        ctx.docId,
        { outlineLevel: lvl, opName: op.type },
      );
    }
    // 确保 pPr 存在（styleSource 无 pPr 时需新建）
    let pPr = firstDirectChildW(newPara, 'pPr');
    if (!pPr) {
      pPr = createW(doc, 'pPr');
      newPara.insertBefore(pPr, newPara.firstChild);
    }
    // 写入或替换 <w:outlineLvl>
    const existing = firstDirectChildW(pPr, 'outlineLvl');
    if (existing) {
      existing.setAttributeNS(W_NS, 'w:val', String(lvl));
    } else {
      const outlineLvlEl = createW(doc, 'outlineLvl');
      outlineLvlEl.setAttributeNS(W_NS, 'w:val', String(lvl));
      // CT_PPr schema 顺序：outlineLvl 必须在 rPr/sectPr/pPrChange 之前。
      // 策略：在 pPr 子元素中找第一个「尾部元素」（rPr/sectPr/pPrChange/divId/cnfStyle），
      // 插在其之前；这些都不存在时才 appendChild。
      // 这样 outlineLvl 一定在 rPr 等之前，符合 schema，Word 能识别大纲级别。
      const TAIL_LOCALS = new Set(['rPr', 'sectPr', 'pPrChange', 'divId', 'cnfStyle']);
      let refNode: Element | null = null;
      for (let i = 0; i < pPr.childNodes.length; i++) {
        const c = pPr.childNodes[i];
        if (
          c &&
          c.nodeType === 1 &&
          (c as Element).namespaceURI === W_NS &&
          TAIL_LOCALS.has((c as Element).localName)
        ) {
          refNode = c as Element;
          break;
        }
      }
      pPr.insertBefore(outlineLvlEl, refNode); // refNode=null → insertBefore(el, null) = appendChild
    }
  }

  // 4c. pageBreakBefore：写入 <w:pageBreakBefore/>（CT_OnOff，存在即 true）到 pPr
  //     CT_PPr schema 顺序：pStyle → keepNext → keepLines → pageBreakBefore → framePr →
  //     widowControl → numPr → … → outlineLvl → rPr → sectPr → pPrChange
  //     即 pageBreakBefore 在 pStyle/keepNext/keepLines 之后、framePr/widowControl/numPr 之前。
  //     策略：找 pPr 中第一个「后置于 pageBreakBefore 的元素」，插在其之前；否则 appendChild。
  if (op.spec.pageBreakBefore === true) {
    // 确保 pPr 存在
    let pPr = firstDirectChildW(newPara, 'pPr');
    if (!pPr) {
      pPr = createW(doc, 'pPr');
      newPara.insertBefore(pPr, newPara.firstChild);
    }
    // 不重复写入（styleSource 克隆过来的 pPr 可能已含 pageBreakBefore）
    if (!firstDirectChildW(pPr, 'pageBreakBefore')) {
      const pbbEl = createW(doc, 'pageBreakBefore');
      // CT_PPr 中 pageBreakBefore 之后的元素（作为插入参考点）
      const AFTER_LOCALS = new Set([
        'framePr',
        'widowControl',
        'suppressLineNumbers',
        'pBdr',
        'shd',
        'tabs',
        'suppressAutoHyphens',
        'kinsoku',
        'wordWrap',
        'overflowPunct',
        'topLinePunct',
        'autoSpaceDE',
        'autoSpaceDN',
        'bidi',
        'adjustRightInd',
        'snapToGrid',
        'spacing',
        'ind',
        'contextualSpacing',
        'mirrorIndents',
        'suppressOverlap',
        'jc',
        'textDirection',
        'textAlignment',
        'textboxTightWrap',
        'outlineLvl',
        'divId',
        'cnfStyle',
        'rPr',
        'sectPr',
        'pPrChange',
        'numPr',
      ]);
      let refNode: Element | null = null;
      for (let i = 0; i < pPr.childNodes.length; i++) {
        const c = pPr.childNodes[i];
        if (
          c &&
          c.nodeType === 1 &&
          (c as Element).namespaceURI === W_NS &&
          AFTER_LOCALS.has((c as Element).localName)
        ) {
          refNode = c as Element;
          break;
        }
      }
      pPr.insertBefore(pbbEl, refNode); // refNode=null → appendChild
    }
  }

  // 5. 生成内容 runs
  appendTextAsRuns(newPara, op.spec.text);

  // 5.1 可选 run-level rPr 继承:克隆源段首个直属 <w:r> 的 <w:rPr>,应用到新段所有内容 run。
  //     styleSource 存在时以 styleSource 段为源(优先级高于 runFormat);
  //     否则若 runFormat='inherit' 以 anchor 为源(#12 兼容)。
  //     用于「簇展开」/全文重组场景——样式写在直接格式(rPr)而非命名样式上。
  //     源段无直属 run / 首 run 无 rPr → 退化为裸 run 不报错。
  //     注意顺序:在 color 之前(color 覆盖 rPr 里已存在的 color)。
  if (styleSource !== undefined) {
    // styleSource 路径:step 4 查 sourcePara。存在则以它为 rPr 源；
    // 不存在（step 4 已降级回退到 anchor pPr）时以 anchor 为 rPr 源，保持与
    // pPr 降级一致——绝不对 undefined 调 applyInheritedRPrToDirectRuns（会崩）。
    const sourcePara = ctx.indices.paraIdx.get(styleSource.toUpperCase());
    applyInheritedRPrToDirectRuns(newPara, sourcePara ?? anchor);
  } else if (op.spec.runFormat === 'inherit') {
    applyInheritedRPrToDirectRuns(newPara, anchor);
  }

  // 5.2 可选 run-level 颜色:为新段所有直属 <w:r> 加 <w:rPr><w:color w:val="..."/></w:rPr>。
  //     双语对照译文场景(word-translate skill)用此区分原段和译文段。
  if (op.spec.color) {
    applyColorToDirectRuns(newPara, op.spec.color);
  }

  // 5a. trackChanges=true(#31):把新段的**内容 runs**包成 <w:ins>,且段 pPr
  //     下加段落 mark 插入标记(w:pPr/w:rPr/w:ins)
  if (ctx.options.trackChanges) {
    const attrs: RevisionAttrs = { author: ctx.options.author, date: ctx.options.date };
    // 取出所有直属 <w:r>(appendTextAsRuns 生成的就是一个 <w:r>),包成 ins
    const newRuns: Element[] = [];
    for (let i = 0; i < newPara.childNodes.length; i++) {
      const c = newPara.childNodes[i];
      if (
        c &&
        c.nodeType === 1 &&
        (c as Element).namespaceURI === W_NS &&
        (c as Element).localName === 'r'
      ) {
        newRuns.push(c as Element);
      }
    }
    if (newRuns.length > 0) {
      const insId = allocateRevisionId(ctx.pkg);
      const insEl = wrapRunsAsInsert(newRuns, insId, attrs);
      // 删原 runs,替换为 ins
      for (const r of newRuns) newPara.removeChild(r);
      newPara.appendChild(insEl);
    }
    // 段落 mark 插入标记
    const markId = allocateRevisionId(ctx.pkg);
    ensureParaMarkInsert(newPara, markId, attrs);
  }

  // 6. 插入到 DOM
  const parent = anchor.parentNode;
  if (!parent) {
    throw new WordError('PACKAGE_CORRUPT', 'anchor paragraph has no parentNode', ctx.docId);
  }
  if (position === 'before') {
    parent.insertBefore(newPara, anchor);
  } else {
    // after:插到 anchor 的下一个兄弟之前(即 anchor 之后)
    parent.insertBefore(newPara, anchor.nextSibling);
  }

  // 7. 更新 ApplyResult
  ctx.result.insertedParagraphs.push({
    opIndex: ctx.opIndex,
    anchorParaId: op.anchorParaId,
    position,
    newParaId: newId,
  });

  // 8. markDirty(必须先于 bookmark 创建,createCrossParagraphBookmark 内部也会再
  //    markDirty 一次,markDirty 幂等)
  markStoryPartDirty(ctx, anchor, op.type);

  // 9. 可选:同事务建跨段书签(#176)
  // 仅 insertParagraphAfter 提供;insertParagraphBefore 显式拒绝 —— 静默忽略会
  // 让调用方误以为书签建好了,后续查询 null 难定位。
  if (op.spec.bookmark) {
    if (position !== 'after') {
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        'spec.bookmark is only supported on insertParagraphAfter (#176)',
        ctx.docId,
        { opType: op.type, anchorParaId: op.anchorParaId },
      );
    }
    const resolvedName = resolveAppendBookmarkName(ctx, op.spec.bookmark.prefix, op.anchorParaId);
    createCrossParagraphBookmark(ctx, resolvedName, anchor, newPara, op.anchorParaId, newId);
  }
}

/**
 * 克隆 `sourcePara` 段首个直属 `<w:r>` 的 `<w:rPr>`,深拷贝后插到 `paragraph` 内
 * 所有直属 `<w:r>` 的首位。
 *
 * 调用方传 anchor 段(runFormat:'inherit' 的 #12 场景)或 styleSource 段(全文重组场景)。
 * sourcePara 无直属 run 或首 run 无 rPr → 静默退化,不修改新段 run(裸 run 行为)。
 * 调用时机:在 appendTextAsRuns 之后、applyColorToDirectRuns 之前,让 color 能覆盖继承值。
 */
function applyInheritedRPrToDirectRuns(paragraph: Element, sourcePara: Element): void {
  // 找 sourcePara 首个直属 <w:r>
  const anchorRun = firstDirectChildW(sourcePara, 'r');
  if (!anchorRun) return;

  // 取首 run 的 <w:rPr>
  const srcRPr = firstDirectChildW(anchorRun, 'rPr');
  if (!srcRPr) return;

  // 应用深拷贝到新段所有直属 <w:r>
  for (let i = 0; i < paragraph.childNodes.length; i++) {
    const c = paragraph.childNodes[i];
    if (
      !c ||
      c.nodeType !== 1 ||
      (c as Element).namespaceURI !== W_NS ||
      (c as Element).localName !== 'r'
    )
      continue;
    const run = c as Element;
    // 插到 run 首位(rPr 必须是 <w:r> 的第一个子元素)
    run.insertBefore(srcRPr.cloneNode(true), run.firstChild);
  }
}

/**
 * 把 6 位十六进制颜色(无 `#`)写到段内所有直属 `<w:r>` 的 `<w:rPr><w:color w:val="..."/></w:rPr>`。
 * 已存在 rPr 时只补/替 color;rPr 不存在则新建并放在 run 首位。
 */
function applyColorToDirectRuns(paragraph: Element, color: string): void {
  const doc = paragraph.ownerDocument;
  if (!doc) return;
  for (let i = 0; i < paragraph.childNodes.length; i++) {
    const c = paragraph.childNodes[i];
    if (
      !c ||
      c.nodeType !== 1 ||
      (c as Element).namespaceURI !== W_NS ||
      (c as Element).localName !== 'r'
    )
      continue;
    const run = c as Element;
    let rPr = firstDirectChildW(run, 'rPr');
    if (!rPr) {
      rPr = createW(doc, 'rPr');
      run.insertBefore(rPr, run.firstChild);
    }
    let colorEl = firstDirectChildW(rPr, 'color');
    if (!colorEl) {
      colorEl = createW(doc, 'color');
      rPr.appendChild(colorEl);
    }
    colorEl.setAttributeNS(W_NS, 'w:val', color);
  }
}

// ─────────── insertParagraphAtBodyStart ───────────

/**
 * 无锚点在 `<w:body>` 起始插一个空段（带 w14:paraId），新 paraId 写入
 * `ctx.result.insertedBodyStartParaIds`。
 *
 * 用途（健壮性）：body 完全无段落（极简/特殊文档：正文段被骨架删光、或源 body 仅
 *   header/footer/文本框内容）时，下游重建需要一个真实落点锚段。此前 typeset 在该情形
 *   直接 throw（"skeleton body 无任何段落，无法定位 body 重建锚点"）→ 整篇崩。本 op 提供
 *   合成锚点能力，让重建走既有 insertParagraphAfter(anchor) 路径，绝不崩、不丢内容。
 *
 * 插入位置：body 的第一个子元素之前（body 的 <w:sectPr> 恒为最后一个子元素，prepend 不会
 *   破坏 sectPr 在尾部的 schema 约束；body 为空时直接成为唯一子元素）。
 */
export function executeInsertParagraphAtBodyStart(
  ctx: ExecutorContext,
  _op: Extract<Op, { type: 'insertParagraphAtBodyStart' }>,
): void {
  const [bodyStory] = listStories(ctx.pkg, { kinds: ['body'] });
  if (!bodyStory) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      'document has no <w:body> story (op: insertParagraphAtBodyStart)',
      ctx.docId,
    );
  }
  const body = bodyStory.root;
  const doc = body.ownerDocument;
  if (!doc) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      'body element has no ownerDocument (op: insertParagraphAtBodyStart)',
      ctx.docId,
    );
  }

  const seen = new Set(ctx.indices.paraIdx.keys());
  const newId = generateParaId(seen);
  const newPara = doc.createElementNS(W_NS, 'w:p');
  newPara.setAttributeNS(W14_NS, 'w14:paraId', newId);

  // 插到 body 第一个子元素之前（body 为空时 firstChild=null → appendChild 语义）。
  body.insertBefore(newPara, body.firstChild);

  ctx.result.insertedBodyStartParaIds.push(newId);

  const part = getPart(ctx.pkg, bodyStory.partPath);
  if (part) markDirty(part);
}

// ─────────── deleteParagraph ───────────

export function executeDeleteParagraph(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'deleteParagraph' }>,
): void {
  // 1. 查段
  const paragraph = ctx.indices.paraIdx.get(op.paraId.toUpperCase());
  if (!paragraph) {
    throw new WordError(
      'PARA_NOT_FOUND',
      `paragraph with paraId='${op.paraId}' not found (op: deleteParagraph)`,
      ctx.docId,
      { paraId: op.paraId, opName: 'deleteParagraph' },
    );
  }

  // 2. 检测段内活锚点
  const affectedBookmarks = collectBookmarkNames(paragraph);
  const affectedComments = collectCommentIds(paragraph);
  const hasAnchors = affectedBookmarks.length > 0 || affectedComments.length > 0;

  if (hasAnchors) {
    if (op.onAnchorConflict === 'drop') {
      if (ctx.options.trackChanges) {
        // drop + trackChanges(#31):修订包装 + 连带删外部锚点主体的组合语义复杂,留后续
        throw new WordError(
          'UNSUPPORTED_OOXML_FEATURE',
          "deleteParagraph with onAnchorConflict='drop' + trackChanges=true not yet implemented",
          ctx.docId,
          { opType: 'deleteParagraph', paraId: op.paraId },
        );
      }
      // drop 模式(#30):删除段内所有活锚点,然后删段
      // - 书签:移除 bookmarkStart + bookmarkEnd,更新 bookmarkIdx
      // - 批注:移除三锚点(commentRangeStart / commentRangeEnd / commentReference)
      //         + comments.xml 里的 <w:comment> 主体
      for (const name of affectedBookmarks) {
        dropBookmarkByName(ctx, name, paragraph);
      }
      for (const id of affectedComments) {
        dropCommentById(ctx, id, paragraph);
      }
    } else {
      // 默认 'error'
      throw new WordError(
        'ANCHOR_CONFLICT',
        `paragraph '${op.paraId}' contains active anchors (bookmarks: ${affectedBookmarks.join(', ') || 'none'}, comments: ${affectedComments.join(', ') || 'none'})`,
        ctx.docId,
        { paraId: op.paraId, affectedBookmarks, affectedComments, opType: 'deleteParagraph' },
      );
    }
  }

  // 3. 无活锚点
  if (ctx.options.trackChanges) {
    // trackChanges=true(#31):不 removeChild;内容 runs 包 <w:del>(w:t→w:delText),
    // 段 pPr 加段落 mark 删除标记
    executeDeleteParagraphTC(ctx, paragraph);
    return;
  }

  // 非 trackChanges:直接删除
  const parent = paragraph.parentNode;
  if (!parent) {
    throw new WordError('PACKAGE_CORRUPT', 'paragraph has no parentNode', ctx.docId);
  }
  markStoryPartDirty(ctx, paragraph, 'deleteParagraph');
  parent.removeChild(paragraph);
}

/**
 * trackChanges=true 下的段落删除:
 *   - 不从 DOM 移除段
 *   - 收集直属 `<w:r>` 内容 runs,整组包成 `<w:del>`(内部 `<w:t>` 转 `<w:delText>`)
 *   - 替换原 runs 为 del
 *   - 段 pPr 下加段落 mark 删除标记(`<w:rPr><w:del/>`)
 */
function executeDeleteParagraphTC(ctx: ExecutorContext, paragraph: Element): void {
  const attrs: RevisionAttrs = { author: ctx.options.author, date: ctx.options.date };

  // 收集直属 <w:r>
  const runs: Element[] = [];
  for (let i = 0; i < paragraph.childNodes.length; i++) {
    const c = paragraph.childNodes[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === 'r'
    ) {
      runs.push(c as Element);
    }
  }

  if (runs.length > 0) {
    const delId = allocateRevisionId(ctx.pkg);
    const delEl = wrapRunsAsDelete(runs, delId, attrs);
    // 在第一个 run 前插 del,再删掉原 runs
    paragraph.insertBefore(delEl, runs[0]!);
    for (const r of runs) paragraph.removeChild(r);
  }

  // 段落 mark 删除标记
  const markId = allocateRevisionId(ctx.pkg);
  ensureParaMarkDelete(paragraph, markId, attrs);

  markStoryPartDirty(ctx, paragraph, 'deleteParagraph');
}

// ─────────── 内部:锚点检测 ───────────

/** 收集段内所有 bookmarkStart 的 w:name(字符串数组)*/
function collectBookmarkNames(paragraph: Element): string[] {
  const names: string[] = [];
  const starts = paragraph.getElementsByTagNameNS(W_NS, 'bookmarkStart');
  for (let i = 0; i < starts.length; i++) {
    const el = starts[i];
    if (!el) continue;
    const name = el.getAttributeNS(W_NS, 'name') ?? el.getAttribute('w:name') ?? '';
    names.push(name);
  }
  return names;
}

/** 收集段内所有 commentRangeStart / commentReference 的 w:id(数字数组,去重)*/
function collectCommentIds(paragraph: Element): number[] {
  const ids = new Set<number>();
  const rangeStarts = paragraph.getElementsByTagNameNS(W_NS, 'commentRangeStart');
  for (let i = 0; i < rangeStarts.length; i++) {
    const el = rangeStarts[i];
    if (!el) continue;
    const id = parseInt(el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '', 10);
    if (!isNaN(id)) ids.add(id);
  }
  const refs = paragraph.getElementsByTagNameNS(W_NS, 'commentReference');
  for (let i = 0; i < refs.length; i++) {
    const el = refs[i];
    if (!el) continue;
    const id = parseInt(el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '', 10);
    if (!isNaN(id)) ids.add(id);
  }
  return Array.from(ids);
}

// ─────────── 内部:drop 模式 —— 书签/批注清理 (#30) ───────────

/**
 * drop 模式下删除书签:移除 bookmarkStart + bookmarkEnd,从 bookmarkIdx 删条目。
 * 简化路径:直接操作 bookmarkIdx 查到的 entry(index 在 applyPlan finally 会重建)。
 */
function dropBookmarkByName(ctx: ExecutorContext, name: string, _paragraph: Element): void {
  const entry = ctx.indices.bookmarkIdx.get(name);
  if (!entry) return; // 已不在索引里(可能索引时重名被去重),直接跳过

  const { start, end, partPath } = entry;

  // 移除 bookmarkStart
  if (start.parentNode) start.parentNode.removeChild(start);

  // 移除 bookmarkEnd(可能为 null 或不在同一段)
  if (end && end.parentNode) end.parentNode.removeChild(end);

  // 更新索引(finally 会重建,此处先同步删掉防止 drop 多书签时二次访问)
  ctx.indices.bookmarkIdx.delete(name);

  // markDirty
  const part = getPart(ctx.pkg, partPath);
  if (part) markDirty(part);
}

/**
 * drop 模式下删除批注:移除三锚点(commentRangeStart / commentRangeEnd / commentReference)
 * + comments.xml 里的 <w:comment> 主体。
 * 从 commentIdx 删条目;markDirty body part + comments.xml。
 */
function dropCommentById(ctx: ExecutorContext, id: number, paragraph: Element): void {
  const entry = ctx.indices.commentIdx.get(id);

  // 若 commentIdx 里没有(可能首次创建还未重建索引),
  // 则手动从段内找并删除锚点元素
  if (!entry) {
    // 直接扫段删锚点 DOM 节点
    removeCommentAnchorsFromParagraph(ctx, paragraph, id);
    return;
  }

  const { start, end, ref } = entry.anchors;

  // body part 路径:从存在的锚点反查
  const anchorEl = start ?? end ?? ref;
  let bodyPartPath: string | null = null;
  if (anchorEl) {
    const story = findStoryFor(ctx.pkg, anchorEl);
    if (story) bodyPartPath = story.partPath;
  }

  if (start && start.parentNode) start.parentNode.removeChild(start);
  if (end && end.parentNode) end.parentNode.removeChild(end);
  if (ref) {
    // ref 可能包在一个 <w:r> 里;若移除 ref 后 run 变空,整体删 run
    const parentRun = ref.parentNode;
    if (ref.parentNode) ref.parentNode.removeChild(ref);
    if (
      parentRun &&
      parentRun.nodeType === 1 &&
      (parentRun as Element).namespaceURI === W_NS &&
      (parentRun as Element).localName === 'r'
    ) {
      if (isRunEffectivelyEmpty(parentRun as Element)) {
        if (parentRun.parentNode) parentRun.parentNode.removeChild(parentRun);
      }
    }
  }

  // 删 <w:comment> 主体
  if (entry.comment && entry.comment.parentNode) {
    entry.comment.parentNode.removeChild(entry.comment);
    const commentsPart = getPart(ctx.pkg, COMMENTS_PART_PATH);
    if (commentsPart) markDirty(commentsPart);
  }

  // markDirty body part
  if (bodyPartPath) {
    const bodyPart = getPart(ctx.pkg, bodyPartPath);
    if (bodyPart) markDirty(bodyPart);
  }

  // 更新索引
  ctx.indices.commentIdx.delete(id);

  // 若 comments.xml 已空，移除 part + 清 CT/rels（与 executeDeleteComment 对称）
  cleanupEmptyCommentsPart(ctx);
}

/**
 * 当 commentIdx 里没有目标批注时,直接从 paragraph DOM 内扫描并删除所有
 * 属于该 commentId 的锚点元素。
 */
function removeCommentAnchorsFromParagraph(
  ctx: ExecutorContext,
  paragraph: Element,
  commentId: number,
): void {
  const idStr = String(commentId);
  // 扫段内全部 commentRangeStart / commentRangeEnd / commentReference
  const localNames = ['commentRangeStart', 'commentRangeEnd', 'commentReference'];
  const toRemove: Element[] = [];
  for (const local of localNames) {
    const list = paragraph.getElementsByTagNameNS(W_NS, local);
    for (let i = 0; i < list.length; i++) {
      const el = list[i];
      if (!el) continue;
      const rawId = el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '';
      if (rawId === idStr) toRemove.push(el);
    }
  }
  for (const el of toRemove) {
    const parentRun = el.localName === 'commentReference' ? el.parentNode : null;
    if (el.parentNode) el.parentNode.removeChild(el);
    if (
      parentRun &&
      parentRun.nodeType === 1 &&
      (parentRun as Element).namespaceURI === W_NS &&
      (parentRun as Element).localName === 'r' &&
      isRunEffectivelyEmpty(parentRun as Element)
    ) {
      if (parentRun.parentNode) parentRun.parentNode.removeChild(parentRun);
    }
  }

  // 同时删 comments.xml 里的主体
  for (const part of listParts(ctx.pkg)) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    if (!part.contentType.endsWith('.comments+xml')) continue;
    const comments = part.xmlDom.getElementsByTagNameNS(W_NS, 'comment');
    const victims: Element[] = [];
    for (let i = 0; i < comments.length; i++) {
      const c = comments[i];
      if (!c) continue;
      const rawId = c.getAttributeNS(W_NS, 'id') ?? c.getAttribute('w:id') ?? '';
      if (rawId === idStr) victims.push(c);
    }
    for (const v of victims) {
      if (v.parentNode) v.parentNode.removeChild(v);
    }
    if (victims.length > 0) markDirty(part);
  }
}

/** 判断 <w:r> 是否已无实质性内容(只剩 rPr 或彻底为空) */
function isRunEffectivelyEmpty(run: Element): boolean {
  const children = run.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (!c || c.nodeType !== 1) continue;
    if ((c as Element).localName === 'rPr') continue;
    return false;
  }
  return true;
}

// ─────────── 内部:markDirty ───────────

function markStoryPartDirty(ctx: ExecutorContext, paragraph: Element, opName: string): void {
  const story = findStoryFor(ctx.pkg, paragraph);
  if (!story) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `paragraph element not found in any story (op: ${opName})`,
      ctx.docId,
      { opName },
    );
  }
  const part = getPart(ctx.pkg, story.partPath);
  if (!part) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `story part '${story.partPath}' not found in pkg (op: ${opName})`,
      ctx.docId,
      { opName, partPath: story.partPath },
    );
  }
  markDirty(part);
}
