/**
 * 段落修改 executor —— #28(replaceParagraph / replaceTextInParagraph)。
 *
 * 设计要点:
 *   - 段枚举走 paraIdx,O(1) 定位到 `<w:p>`。
 *   - 保留策略:`<w:pPr>`(样式 / 编号 / 段落属性)整体不动;
 *     本 PR **直属 marker**(段内的 bookmarkStart/End / commentRange* 等)在
 *     `replaceParagraph` 里一并清掉 —— ANCHOR_CONFLICT 检测留给 #30 补保护层
 *   - trackChanges=true(#31):replaceParagraph 把旧 runs 包成 `<w:del>`,
 *     新 runs 包成 `<w:ins>`;replaceTextInParagraph 把匹配区间包成 `<w:del>`,
 *     replacement 包成 `<w:ins>`
 *   - 新 run 使用既有命名空间前缀(从 ownerDocument 继承),用 `createElementNS`
 *     确保 URI 对齐;`<w:t>` 一律挂 `xml:space="preserve"` 保留首尾空格
 *
 * **复杂 run 限制**(沿用 `splitRunAt`):`replaceTextInParagraph` 跨 run 匹配时,
 * 匹配区间涉及到的 run 必须是"单 `<w:t>` 简单 run";遇 `<w:br>` / `<w:tab>` /
 * 多 `<w:t>` 等 → 抛 `UNSUPPORTED_OOXML_FEATURE`。
 */
import { getPart, markDirty } from '../../pack/index.js';
import { findStoryFor } from '../stories.js';
import { W_NS, W14_NS, WPS_NS, VML_NS, XML_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { getParagraphRuns, getRunText, splitRunAt } from '../runs.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';
import {
  appendSegmentsToRun,
  appendTextAsRuns,
  collectParagraphAnchors,
  collectRangeAnchors,
  createW,
  firstDirectChildW,
  pickFirstRunRPr,
} from './_shared.js';
import {
  allocateRevisionId,
  wrapRunsAsInsert,
  wrapRunsAsDelete,
  type RevisionAttrs,
} from './revision-wrappers.js';

/** 段落中"marker"(非内容)直属子元素,在 replaceParagraph 里**保留**的集合 */
const PARAGRAPH_HEADER_MARKERS = new Set<string>(['pPr']);

/** 段落中被 replaceParagraph **移除**的内容元素集合(任何能承载显示字节的 wrapper)*/
const PARAGRAPH_CONTENT_ELEMENTS = new Set<string>([
  'r',
  'hyperlink',
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'smartTag',
  'sdt',
  'fldSimple',
  'customXml',
]);

/** 段内 marker 的命名(bookmarkStart / commentRangeStart 等)—— replaceParagraph 当前一并清掉,ANCHOR_CONFLICT 由 #30 保护 */
const PARAGRAPH_ANCHOR_MARKERS = new Set<string>([
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'commentReference',
  'permStart',
  'permEnd',
  'proofErr',
]);

/**
 * 判断 `<w:r>` 是否含图形后代。
 * 检测：`<w:drawing>`、`<w:pict>`、`<wps:wsp>`（Word 2010+ shape）、
 *       `<v:line>` / `<v:shape>` / `<v:rect>`（VML）、`<w:object>`（OLE）。
 * 规则：在 `run` 所有后代中查找任意一个满足条件的元素即视为图形 run。
 */
function isGraphicRun(run: Element): boolean {
  // 直接子元素级别的 <w:drawing> / <w:pict> / <w:object>（W_NS）
  for (let i = 0; i < run.childNodes.length; i++) {
    const c = run.childNodes[i];
    if (!c || c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI === W_NS) {
      const ln = el.localName;
      if (ln === 'drawing' || ln === 'pict' || ln === 'object') return true;
    }
  }
  // 深层后代：wps:wsp（在 w:drawing 内部的 mc:AlternateContent 分支中），
  // v:line / v:shape / v:rect（VML，在 w:pict 内或 mc:Fallback 中）
  // 用 getElementsByTagNameNS 比手写递归简洁且在 slimdom 里有效
  if (run.getElementsByTagNameNS(WPS_NS, 'wsp').length > 0) return true;
  if (run.getElementsByTagNameNS(VML_NS, 'line').length > 0) return true;
  if (run.getElementsByTagNameNS(VML_NS, 'shape').length > 0) return true;
  if (run.getElementsByTagNameNS(VML_NS, 'rect').length > 0) return true;
  return false;
}

// ─────────── replaceParagraph ───────────

export function executeReplaceParagraph(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'replaceParagraph' }>,
): void {
  const paragraph = requireParagraph(ctx, op.paraId, 'replaceParagraph');

  if (ctx.options.trackChanges) {
    // trackChanges=true(#31):旧内容(含 bookmark/comment marker)整体包成 <w:del>,
    // 新 runs 包成 <w:ins>;锚点跟随被删内容进入修订包装,语义自洽:
    // 用户接受 → 文本+锚点一起删;拒绝 → 一起恢复(与 Word/WPS 原生 trackChanges 行为对齐)
    executeReplaceParagraphTC(ctx, paragraph, op);
  } else {
    // 非 trackChanges：ANCHOR_CONFLICT 检测(#30)
    // keepInlineGraphics=true 时绕过锚点冲突检测：该模式会就地保留锚点 marker，
    // 不会产生孤儿，不需要拦截。
    // 不带 keepInlineGraphics 时，直接覆盖会让锚点孤立、且无法回滚，拒掉。
    if (!op.keepInlineGraphics) {
      const anchors = collectParagraphAnchors(paragraph);
      if (anchors.bookmarks.length > 0 || anchors.comments.length > 0) {
        if (op.onAnchorConflict === 'drop') {
          // drop：跳过本次替换，保留原段落内容不变
          return;
        }
        throw new WordError(
          'ANCHOR_CONFLICT',
          `paragraph modification conflicts with active anchors (paraId=${op.paraId}); ` +
            `enable trackChanges to wrap the deletion as a revision`,
          ctx.docId,
          {
            paraId: op.paraId,
            affectedBookmarks: anchors.bookmarks,
            affectedComments: anchors.comments,
            opType: 'replaceParagraph',
          },
        );
      }
    }
    // 删旧内容前先抓首个 run 的 rPr（task-core-replacepara-rpr：透传字号/字体）
    const inheritedRPr = pickFirstRunRPr(paragraph);
    if (op.keepInlineGraphics) {
      // keepInlineGraphics=true：只删纯文字 run，图形 run 与锚点 marker 保留原位。
      // 新文字 run 插在第一个图形 run 之前，保持"文字在前、线条在后"原序。
      // 注意：要在 removeContentChildren 之前找到第一个图形 run 的引用，
      // 因为删除后文字 run 引用会脱离 DOM，不能作为 insertBefore 锚点。
      const firstGraphicRunAnchor = findFirstGraphicRunAnchor(paragraph);
      removeContentChildren(paragraph, true);
      const doc = paragraph.ownerDocument;
      if (!doc) throw new WordError('PACKAGE_CORRUPT', 'paragraph has no ownerDocument', ctx.docId);
      appendTextAsRunsAt(paragraph, op.text, inheritedRPr, firstGraphicRunAnchor);
    } else {
      removeContentChildren(paragraph);
      appendTextAsRuns(paragraph, op.text, inheritedRPr);
    }
  }

  // markDirty 对应 part
  markStoryPartDirty(ctx, paragraph, 'replaceParagraph');
}

/**
 * trackChanges=true 模式下的 replaceParagraph:
 *   - 旧内容(runs + bookmark/comment marker + 已存在的 ins/del wrapper)整体包成 `<w:del>`
 *     (w:t → w:delText;锚点 marker 直接作为 `<w:del>` 子节点保留,与 Word/WPS 行为一致)
 *   - 新 runs 整体包成 `<w:ins>`
 *   - 顺序:pPr → [图形 run(若 keepInlineGraphics)]… → del → ins → [图形 run(若在文字后)]
 *
 * keepInlineGraphics=true 时：
 *   - 图形 run 不进入 `<w:del>`（保留在段落原位，不产生 del 修订）
 *   - 只有纯文字的内容子（及非图形 run）被收入旧内容数组，包成 del
 *   - `<w:ins>` 插在图形 run 之前（首个图形 run 前），保持"文字在前、线条在后"原序
 */
function executeReplaceParagraphTC(
  ctx: ExecutorContext,
  paragraph: Element,
  op: Extract<Op, { type: 'replaceParagraph' }>,
): void {
  const attrs: RevisionAttrs = { author: ctx.options.author, date: ctx.options.date };

  // 0. 删旧内容前先抓首个 run 的 rPr（task-core-replacepara-rpr：透传字号/字体）
  //    必须在 removeContentChildren 之前，否则 run 被清掉后取不到
  const inheritedRPr = pickFirstRunRPr(paragraph);

  if (op.keepInlineGraphics) {
    // ── keepInlineGraphics=true 分支 ──
    // 1a. 分离：图形 run 保留原位，纯文字内容子收入待删列表
    const { textChildren, graphicRuns } = separateGraphicAndTextChildren(paragraph);

    // 2a. 只删纯文字内容子（图形 run 保留）
    for (const c of textChildren) {
      paragraph.removeChild(c);
    }

    // 3a. 若有旧文字内容，包成 <w:del>
    //     插入锚点：首个图形 run 之前（保持文字区在图形区之前）
    const firstGraphicRun = graphicRuns[0] ?? null;
    if (textChildren.length > 0) {
      const delId = allocateRevisionId(ctx.pkg);
      const delEl = wrapRunsAsDelete(textChildren, delId, attrs);
      paragraph.insertBefore(delEl, firstGraphicRun);
    }

    // 4a. 构建新 run 并包成 <w:ins>，同样插在首个图形 run 之前
    const doc = paragraph.ownerDocument;
    if (!doc) throw new WordError('PACKAGE_CORRUPT', 'paragraph has no ownerDocument', ctx.docId);
    const newRun = createW(doc, 'r');
    if (inheritedRPr) newRun.insertBefore(inheritedRPr, newRun.firstChild);
    appendSegmentsToRun(doc, newRun, op.text);
    const insId = allocateRevisionId(ctx.pkg);
    const insEl = wrapRunsAsInsert([newRun], insId, attrs);
    paragraph.insertBefore(insEl, firstGraphicRun);
  } else {
    // ── 默认分支（现有行为） ──
    // 1. 收集现有所有内容子(在清之前)—— 不仅是 <w:r>,还包括 bookmarkStart/End、
    //    commentRangeStart/End、commentReference、已存在的 <w:ins>/<w:del> wrapper 等。
    //    pPr 不包含在内(由 PARAGRAPH_HEADER_MARKERS 排除)。
    const existingChildren = collectContentChildren(paragraph);

    // 2. 清掉所有内容子(保留 pPr)
    removeContentChildren(paragraph);

    // 3. 若有旧内容,整体包成 <w:del>(wrapRunsAsDelete 会 clone 并把 <w:t> 转 <w:delText>,
    //    锚点 marker 内部没 <w:t>,会被原样保留)
    if (existingChildren.length > 0) {
      const delId = allocateRevisionId(ctx.pkg);
      const delEl = wrapRunsAsDelete(existingChildren, delId, attrs);
      paragraph.appendChild(delEl);
    }

    // 4. 构建新 run 并包成 <w:ins>
    const doc = paragraph.ownerDocument;
    if (!doc) throw new WordError('PACKAGE_CORRUPT', 'paragraph has no ownerDocument', ctx.docId);
    const newRun = createW(doc, 'r');
    if (inheritedRPr) newRun.insertBefore(inheritedRPr, newRun.firstChild);
    appendSegmentsToRun(doc, newRun, op.text);
    const insId = allocateRevisionId(ctx.pkg);
    const insEl = wrapRunsAsInsert([newRun], insId, attrs);
    paragraph.appendChild(insEl);
  }
}

// ─────────── replaceTextInParagraph ───────────

export function executeReplaceTextInParagraph(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'replaceTextInParagraph' }>,
): void {
  if (op.search.length === 0) {
    throw new WordError(
      'TEXT_NOT_FOUND_IN_PARA',
      `replaceTextInParagraph: 'search' cannot be empty`,
      ctx.docId,
      { paraId: op.paraId, search: op.search },
    );
  }

  const paragraph = requireParagraph(ctx, op.paraId, 'replaceTextInParagraph');
  const runs = getParagraphRuns(paragraph);

  // 1. 构建段全文 + run 偏移映射(runIndex → [runStartOffset, runEndOffset))
  //    注意:getParagraphRuns 只看直属 <w:r>,跨 hyperlink / ins/del 包装的 run
  //    不被枚举 —— 本 PR 接受(复杂结构抛错)。
  const fullText = runs.map((r) => getRunText(r)).join('');

  // 2. 在全文查找所有匹配(first / all 决定取几个),按**倒序**处理 —— 先处理后段
  //    的匹配不影响前段的 run 偏移
  const matches = findMatches(fullText, op.search, op.scope);
  if (matches.length === 0) {
    throw new WordError(
      'TEXT_NOT_FOUND_IN_PARA',
      `text '${op.search}' not found in paragraph ${op.paraId}`,
      ctx.docId,
      { paraId: op.paraId, search: op.search, scope: op.scope },
    );
  }

  // 3. 逐匹配替换(倒序)
  //    每次替换后段内 runs 变化(新增 / 删除),下一个匹配的 run 定位需要基于最新
  //    状态。最稳的做法:**每次替换后重扫 runs / fullText**。倒序处理确保前面
  //    的匹配偏移仍然有效(因为我们先改后面的)。
  const sortedMatches = matches.slice().sort((a, b) => b.start - a.start);
  for (const m of sortedMatches) {
    replaceOneMatch(ctx, paragraph, m, op);
  }

  markStoryPartDirty(ctx, paragraph, 'replaceTextInParagraph');
}

export function executeFormatTextInParagraph(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'formatTextInParagraph' }>,
): void {
  if (op.search.length === 0) {
    throw new WordError(
      'TEXT_NOT_FOUND_IN_PARA',
      `formatTextInParagraph: 'search' cannot be empty`,
      ctx.docId,
      { paraId: op.paraId, search: op.search },
    );
  }
  validateTextFormat(ctx, op);
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'formatTextInParagraph does not support trackChanges yet',
      ctx.docId,
      { paraId: op.paraId, opType: 'formatTextInParagraph' },
    );
  }

  const paragraph = requireParagraph(ctx, op.paraId, 'formatTextInParagraph');
  const runs = getParagraphRuns(paragraph);
  const fullText = runs.map((r) => getRunText(r)).join('');
  const matches = findMatches(fullText, op.search, op.scope);
  if (matches.length === 0) {
    throw new WordError(
      'TEXT_NOT_FOUND_IN_PARA',
      `text '${op.search}' not found in paragraph ${op.paraId}`,
      ctx.docId,
      { paraId: op.paraId, search: op.search, scope: op.scope },
    );
  }

  const sortedMatches = matches.slice().sort((a, b) => b.start - a.start);
  for (const m of sortedMatches) {
    formatOneMatch(ctx, paragraph, m, op);
  }

  markStoryPartDirty(ctx, paragraph, 'formatTextInParagraph');
}

export function executeFormatParagraph(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'formatParagraph' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'formatParagraph does not support trackChanges yet',
      ctx.docId,
      { paraId: op.paraId, opType: 'formatParagraph' },
    );
  }
  validateParagraphFormat(ctx, op);
  const paragraph = requireParagraph(ctx, op.paraId, 'formatParagraph');
  const pPr = ensureParagraphProperties(paragraph);

  if (op.format.alignment !== undefined) {
    removeDirectChildrenW(pPr, 'jc');
    const jc = createW(pPr.ownerDocument!, 'jc');
    jc.setAttributeNS(W_NS, 'w:val', op.format.alignment);
    pPr.insertBefore(jc, findPPrInsertRef(pPr, JC_AFTER_LOCALS));
  }

  if (op.format.spacing !== undefined) {
    const spacing = upsertDirectChildW(pPr, 'spacing', SPACING_AFTER_LOCALS);
    if (op.format.spacing.before !== undefined) {
      spacing.setAttributeNS(W_NS, 'w:before', String(op.format.spacing.before));
    }
    if (op.format.spacing.after !== undefined) {
      spacing.setAttributeNS(W_NS, 'w:after', String(op.format.spacing.after));
    }
    if (op.format.spacing.line !== undefined) {
      spacing.setAttributeNS(W_NS, 'w:line', String(op.format.spacing.line));
    }
    if (op.format.spacing.lineRule !== undefined) {
      spacing.setAttributeNS(W_NS, 'w:lineRule', op.format.spacing.lineRule);
    }
  }

  if (op.format.indent !== undefined) {
    const ind = upsertDirectChildW(pPr, 'ind', IND_AFTER_LOCALS);
    if (op.format.indent.left !== undefined) {
      ind.setAttributeNS(W_NS, 'w:left', String(op.format.indent.left));
    }
    if (op.format.indent.right !== undefined) {
      ind.setAttributeNS(W_NS, 'w:right', String(op.format.indent.right));
    }
    if (op.format.indent.firstLine !== undefined) {
      if (op.format.indent.firstLine < 0) {
        ind.removeAttributeNS(W_NS, 'firstLine');
        ind.removeAttribute('w:firstLine');
        ind.setAttributeNS(W_NS, 'w:hanging', String(Math.abs(op.format.indent.firstLine)));
      } else {
        ind.removeAttributeNS(W_NS, 'hanging');
        ind.removeAttribute('w:hanging');
        ind.setAttributeNS(W_NS, 'w:firstLine', String(op.format.indent.firstLine));
      }
    }
  }

  if (op.format.outlineLevel !== undefined) {
    const existing = firstDirectChildW(pPr, 'outlineLvl');
    if (existing) {
      existing.setAttributeNS(W_NS, 'w:val', String(op.format.outlineLevel));
    } else {
      const outlineLvl = createW(pPr.ownerDocument!, 'outlineLvl');
      outlineLvl.setAttributeNS(W_NS, 'w:val', String(op.format.outlineLevel));
      pPr.insertBefore(outlineLvl, findPPrInsertRef(pPr, OUTLINE_AFTER_LOCALS));
    }
  }

  if (op.format.pageBreakBefore !== undefined) {
    removeDirectChildrenW(pPr, 'pageBreakBefore');
    if (op.format.pageBreakBefore) {
      const pageBreakBefore = createW(pPr.ownerDocument!, 'pageBreakBefore');
      pPr.insertBefore(pageBreakBefore, findPPrInsertRef(pPr, PAGE_BREAK_BEFORE_AFTER_LOCALS));
    }
  }

  markStoryPartDirty(ctx, paragraph, 'formatParagraph');
}

// ─────────── 内部:查段 / 去内容 / 生成 run ───────────

function requireParagraph(ctx: ExecutorContext, paraId: string, opName: string): Element {
  const el = ctx.indices.paraIdx.get(paraId.toUpperCase());
  if (!el) {
    throw new WordError(
      'PARA_NOT_FOUND',
      `paragraph with paraId='${paraId}' not found (op: ${opName})`,
      ctx.docId,
      { paraId, opName },
    );
  }
  return el;
}

/**
 * 从 `<w:p>` 移除所有内容 / marker 子(保留 pPr)。
 * pPr 通常位置在最前，这里**按 localName 判断**，不依赖位置。
 *
 * @param keepInlineGraphics 若为 true，对含图形后代的 `<w:r>` 跳过（保留原位），
 *   同时保留锚点 marker（bookmarkStart/End、commentRangeStart/End、commentReference）。
 *   只删纯文字 run 和非 run 的非锚点内容子。默认 false（现有行为：全删）。
 */
function removeContentChildren(paragraph: Element, keepInlineGraphics = false): void {
  const children = Array.from(paragraph.childNodes);
  for (const child of children) {
    if (child.nodeType !== 1) {
      // 非元素(文本节点等):统一移除
      paragraph.removeChild(child);
      continue;
    }
    const el = child as Element;
    if (el.namespaceURI === W_NS && PARAGRAPH_HEADER_MARKERS.has(el.localName ?? '')) {
      continue; // pPr 保留
    }
    if (keepInlineGraphics) {
      // keepInlineGraphics=true：含图形后代的 <w:r> 跳过不删
      if (el.namespaceURI === W_NS && el.localName === 'r' && isGraphicRun(el)) {
        continue;
      }
      // keepInlineGraphics=true：直属锚点 marker（bookmarkStart/End、commentRangeStart/End）保留
      if (el.namespaceURI === W_NS && PARAGRAPH_ANCHOR_MARKERS.has(el.localName ?? '')) {
        continue;
      }
      // keepInlineGraphics=true：包含 commentReference 的 <w:r>（commentReference 通常在 run 内）保留
      if (
        el.namespaceURI === W_NS &&
        el.localName === 'r' &&
        el.getElementsByTagNameNS(W_NS, 'commentReference').length > 0
      ) {
        continue;
      }
    }
    // 其它都删:内容 wrappers（纯文字 run / hyperlink / ins / del 等）+ 任何非 W_NS 元素
    paragraph.removeChild(child);
  }
  // 断言:keepInlineGraphics=false 时只剩 pPr(或空);keepInlineGraphics=true 时剩 pPr + 图形 run + 锚点 marker。
  // 调用方负责后续追加。
  void PARAGRAPH_CONTENT_ELEMENTS; // 仅用于文档意图,运行时不做 assert
}

/**
 * 收集段落直属子中所有"内容子节点"——除 pPr 之外的所有元素子。
 * 用于 trackChanges 模式下记录旧内容(之后整体包成 `<w:del>`):
 *   - `<w:r>` 内容 run
 *   - `<w:bookmarkStart>` / `<w:bookmarkEnd>` 等书签 marker
 *   - `<w:commentRangeStart>` / `<w:commentRangeEnd>` / `<w:commentReference>` 批注 marker
 *   - 已存在的 `<w:ins>` / `<w:del>` wrapper(覆盖性替换时一并裹进新 del)
 *   - 任何非 W_NS 元素(SDT 等)
 * 与 `removeContentChildren` 的"删谁"集合保持一致(都跳过 pPr)。
 */
function collectContentChildren(paragraph: Element): Element[] {
  const out: Element[] = [];
  const children = paragraph.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (!c || c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI === W_NS && PARAGRAPH_HEADER_MARKERS.has(el.localName ?? '')) {
      continue; // pPr 跳过
    }
    out.push(el);
  }
  return out;
}

/**
 * keepInlineGraphics=true TC 分支：分离段内图形 run 和纯文字内容子。
 * 返回:
 *   - textChildren: 所有非图形的内容子（会被包进 del）
 *   - graphicRuns: 所有图形 run（保留原位，不进 del）
 */
function separateGraphicAndTextChildren(paragraph: Element): {
  textChildren: Element[];
  graphicRuns: Element[];
} {
  const textChildren: Element[] = [];
  const graphicRuns: Element[] = [];
  const children = paragraph.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (!c || c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI === W_NS && PARAGRAPH_HEADER_MARKERS.has(el.localName ?? '')) {
      continue; // pPr 跳过
    }
    if (el.namespaceURI === W_NS && el.localName === 'r' && isGraphicRun(el)) {
      graphicRuns.push(el);
    } else {
      textChildren.push(el);
    }
  }
  return { textChildren, graphicRuns };
}

/**
 * 找到段内第一个图形 run 的 DOM 引用，作为 insertBefore 的锚点。
 * 用于 keepInlineGraphics=true 非 TC 模式：新文字 run 插在第一个图形 run 之前，
 * 保持"文字在前、线条在后"原序。
 *
 * 必须在 removeContentChildren 之前调用，此时图形 run 仍在 DOM 中。
 * 若无图形 run（段内无图形），返回 null（新 run appendchild 到末尾即可）。
 */
function findFirstGraphicRunAnchor(paragraph: Element): Element | null {
  const children = paragraph.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (!c || c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI === W_NS && PARAGRAPH_HEADER_MARKERS.has(el.localName ?? '')) {
      continue; // pPr 跳过
    }
    if (el.namespaceURI === W_NS && el.localName === 'r' && isGraphicRun(el)) {
      return el; // 找到第一个图形 run
    }
  }
  return null; // 无图形 run（纯文字段）
}

/**
 * 把文本追加为 run(s)，插在 `graphicAnchor`（第一个图形 run）之前。
 * `graphicAnchor` 为 null 时退化为 appendChild（无图形 run，纯文字段）。
 * 用于 keepInlineGraphics=true 模式，保持"新文字 run 在图形 run 之前"的顺序。
 *
 * 实现：临时把所有图形 run 从段内移走，appendTextAsRuns 追加文字 run 到末尾，
 * 再把图形 run 放回。这样即使 appendTextAsRuns 生成多个 run（\n → br）也正确。
 */
function appendTextAsRunsAt(
  paragraph: Element,
  text: string,
  inheritedRPr: Element | undefined,
  graphicAnchor: Element | null,
): void {
  if (graphicAnchor === null) {
    // 无图形 run（纯文字段）：直接追加（等同于原有行为）
    appendTextAsRuns(paragraph, text, inheritedRPr);
    return;
  }

  // 收集 graphicAnchor 及其后的所有图形 run（已保留在 DOM 中的，需要暂时移走）
  const graphicRuns: Element[] = [];
  let cur: Node | null = graphicAnchor;
  while (cur) {
    if (cur.nodeType === 1) {
      const el = cur as Element;
      if (el.namespaceURI === W_NS && el.localName === 'r' && isGraphicRun(el)) {
        graphicRuns.push(el);
      }
    }
    cur = cur.nextSibling;
  }

  // 临时把图形 run 移走（detach），appendTextAsRuns 追加文字，再 re-attach
  for (const gr of graphicRuns) {
    paragraph.removeChild(gr);
  }
  appendTextAsRuns(paragraph, text, inheritedRPr);
  for (const gr of graphicRuns) {
    paragraph.appendChild(gr);
  }
}

// appendTextAsRuns / createW / firstDirectChildW 已提出到 _shared.ts(#29 重构)

function markStoryPartDirty(ctx: ExecutorContext, paragraph: Element, opName: string): void {
  const story = findStoryFor(ctx.pkg, paragraph);
  if (!story) {
    // 理论上不该发生:段在 paraIdx 但不在任何 story —— 说明装载期状态异常
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

const PAGE_BREAK_BEFORE_AFTER_LOCALS = new Set([
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

const SPACING_AFTER_LOCALS = new Set([
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
]);

const IND_AFTER_LOCALS = new Set([
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
]);

const JC_AFTER_LOCALS = new Set([
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange',
]);

const OUTLINE_AFTER_LOCALS = new Set(['rPr', 'sectPr', 'pPrChange', 'divId', 'cnfStyle']);

function validateParagraphFormat(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'formatParagraph' }>,
): void {
  if (
    op.format.alignment === undefined &&
    op.format.outlineLevel === undefined &&
    op.format.pageBreakBefore === undefined &&
    op.format.indent === undefined &&
    op.format.spacing === undefined
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatParagraph requires at least one format field',
      ctx.docId,
      { paraId: op.paraId, opType: 'formatParagraph' },
    );
  }
  if (
    op.format.alignment !== undefined &&
    !['left', 'center', 'right', 'both', 'distribute'].includes(op.format.alignment)
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatParagraph alignment must be left, center, right, both, or distribute',
      ctx.docId,
      { paraId: op.paraId, alignment: op.format.alignment, opType: 'formatParagraph' },
    );
  }
  if (
    op.format.outlineLevel !== undefined &&
    (!Number.isInteger(op.format.outlineLevel) ||
      op.format.outlineLevel < 0 ||
      op.format.outlineLevel > 8)
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatParagraph outlineLevel must be an integer from 0 to 8',
      ctx.docId,
      { paraId: op.paraId, outlineLevel: op.format.outlineLevel, opType: 'formatParagraph' },
    );
  }
  if (op.format.indent !== undefined) {
    validateOptionalTwips(ctx, op.paraId, 'indent.left', op.format.indent.left);
    validateOptionalTwips(ctx, op.paraId, 'indent.right', op.format.indent.right);
    validateOptionalTwips(ctx, op.paraId, 'indent.firstLine', op.format.indent.firstLine);
  }
  if (op.format.spacing !== undefined) {
    validateOptionalTwips(ctx, op.paraId, 'spacing.before', op.format.spacing.before);
    validateOptionalTwips(ctx, op.paraId, 'spacing.after', op.format.spacing.after);
    validateOptionalTwips(ctx, op.paraId, 'spacing.line', op.format.spacing.line);
    if (
      op.format.spacing.lineRule !== undefined &&
      !['auto', 'exact', 'atLeast'].includes(op.format.spacing.lineRule)
    ) {
      throw new WordError(
        'INVALID_ARGUMENT',
        'formatParagraph spacing.lineRule must be auto, exact, or atLeast',
        ctx.docId,
        { paraId: op.paraId, lineRule: op.format.spacing.lineRule, opType: 'formatParagraph' },
      );
    }
  }
}

function validateOptionalTwips(
  ctx: ExecutorContext,
  paraId: string,
  name: string,
  value: number | undefined,
): void {
  if (value === undefined) return;
  if (!Number.isInteger(value)) {
    throw new WordError(
      'INVALID_ARGUMENT',
      `formatParagraph ${name} must be an integer twip value`,
      ctx.docId,
      { paraId, value, opType: 'formatParagraph', field: name },
    );
  }
}

function ensureParagraphProperties(paragraph: Element): Element {
  const existing = firstDirectChildW(paragraph, 'pPr');
  if (existing) return existing;
  const pPr = createW(paragraph.ownerDocument!, 'pPr');
  paragraph.insertBefore(pPr, paragraph.firstChild);
  return pPr;
}

function removeDirectChildrenW(parent: Element, localName: string): void {
  const toRemove: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName === localName
    ) {
      toRemove.push(child as Element);
    }
  }
  for (const child of toRemove) {
    parent.removeChild(child);
  }
}

function upsertDirectChildW(parent: Element, localName: string, afterLocals: Set<string>): Element {
  const existing = firstDirectChildW(parent, localName);
  if (existing) return existing;
  const child = createW(parent.ownerDocument!, localName);
  parent.insertBefore(child, findPPrInsertRef(parent, afterLocals));
  return child;
}

function findPPrInsertRef(pPr: Element, afterLocals: Set<string>): Element | null {
  for (let i = 0; i < pPr.childNodes.length; i++) {
    const child = pPr.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      afterLocals.has((child as Element).localName)
    ) {
      return child as Element;
    }
  }
  return null;
}

// ─────────── replaceTextInParagraph:匹配 + 替换 ───────────

interface Match {
  start: number;
  end: number;
}

function findMatches(text: string, search: string, scope: 'first' | 'all'): Match[] {
  const out: Match[] = [];
  let from = 0;
  while (from <= text.length) {
    const idx = text.indexOf(search, from);
    if (idx < 0) break;
    out.push({ start: idx, end: idx + search.length });
    if (scope === 'first') break;
    from = idx + search.length;
  }
  return out;
}

/** 替换单个匹配。流程:
 *  1. 重扫当前 runs(因上一次替换可能已变)
 *  2. 找匹配 start 所在 run + run 内偏移 → splitRunAt 切出 startInnerRun
 *  3. 找匹配 end 所在 run + 偏移 → splitRunAt 切出 endInnerRun
 *  4. 收集 [startInnerRun, endInnerRun] 之间所有 run(同级兄弟)
 *  5. 删除这些 run,插入用首 run rPr 深拷贝的新 run(内容来自 replacement)
 */
function replaceOneMatch(
  ctx: ExecutorContext,
  paragraph: Element,
  match: Match,
  op: Extract<Op, { type: 'replaceTextInParagraph' }>,
): void {
  const runs = getParagraphRuns(paragraph);
  // 构造 (runIndex, runText, cumulativeStart)
  const runSpans: Array<{ run: Element; start: number; end: number }> = [];
  let cursor = 0;
  for (const r of runs) {
    const len = getRunText(r).length;
    runSpans.push({ run: r, start: cursor, end: cursor + len });
    cursor += len;
  }

  // 找 match.start 落在的 run:span 的 [start, end)
  const startSpan = runSpans.find((s) => match.start >= s.start && match.start < s.end);
  // match.end 可能恰好等于某个 span 的 end,允许 match.end === s.end
  const endSpan = runSpans.find((s) => match.end > s.start && match.end <= s.end);
  if (!startSpan || !endSpan) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `replaceTextInParagraph: match spans runs not reachable via direct <w:r> enumeration ` +
        `(possibly wrapped in hyperlink/ins/del). paraId=${op.paraId}`,
      ctx.docId,
      { paraId: op.paraId, opType: 'replaceTextInParagraph' },
    );
  }

  // 在 startSpan 的 offset = match.start - startSpan.start 切;在 endSpan 的
  // offset = match.end - endSpan.start 切。注意顺序:先切 end 再切 start,避免
  // start 切分改变 endSpan 的 element 引用(若两者同 run,之后需重找)。
  const sameRun = startSpan.run === endSpan.run;

  // A host roundtrip may coalesce text separated by `<w:tab/>` into one run with
  // multiple `<w:t>` children. If the match itself stays inside one text child,
  // replace that child in place: this preserves tabs, drawings, and the run format
  // without asking splitRunAt to split the surrounding complex run.
  if (sameRun && !ctx.options.trackChanges && !/[\r\n\t]/.test(op.replacement)) {
    const startOffset = match.start - startSpan.start;
    const endOffset = match.end - startSpan.start;
    let textOffset = 0;
    const children = startSpan.run.childNodes;
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (
        !child ||
        child.nodeType !== 1 ||
        (child as Element).namespaceURI !== W_NS ||
        (child as Element).localName !== 't'
      ) {
        continue;
      }
      const t = child as Element;
      const value = t.textContent ?? '';
      const textEnd = textOffset + value.length;
      if (startOffset >= textOffset && endOffset <= textEnd) {
        const localStart = startOffset - textOffset;
        const localEnd = endOffset - textOffset;
        t.textContent = value.slice(0, localStart) + op.replacement + value.slice(localEnd);
        t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
        return;
      }
      textOffset = textEnd;
    }
  }

  // 切 end 点(若同 run,start 的后半经过切后成为"待删 run"的 matchRun)
  const endOffset = match.end - endSpan.start;
  let startRunForMatch: Element;
  let endRunForMatch: Element;

  try {
    if (sameRun) {
      // 同 run 场景:先在 endOffset 切,得到 [原 run=leftpart+matchpart, 新 run=rightpart]
      // 然后在原 run 内再切 startOffset,得到 [左半, 匹配部分]
      const { after: rightAfterEnd } = splitRunAt(startSpan.run, endOffset);
      // 此时 startSpan.run 的文本为 [span.start, match.end)
      // 再在原 run 的 match.start 偏移处切
      const startOffsetInRun = match.start - startSpan.start;
      const { after: matchRun } = splitRunAt(startSpan.run, startOffsetInRun);
      // matchRun 是仅含 match 的 run;若没切出(startOffset=0),则 startSpan.run 本身就是 matchRun
      startRunForMatch = matchRun ?? startSpan.run;
      endRunForMatch = startRunForMatch;
      // rightAfterEnd 是 end 之后的 run(可能是 null,如果 match.end == span.end)
      void rightAfterEnd;
    } else {
      // 跨 run:分别切 startSpan(在 startOffset 处,保留左半,切出右半=匹配的首 run)
      // 和 endSpan(在 endOffset 处,保留右半,左半=匹配的末 run)
      const startOffsetInRun = match.start - startSpan.start;
      const { after: afterStart } = splitRunAt(startSpan.run, startOffsetInRun);
      startRunForMatch = afterStart ?? startSpan.run;

      // endSpan 可能是一个没被上面切过的 run;若是,offset 就是 match.end - endSpan.start
      // 若同一次调用 startSpan === endSpan 已由 sameRun 分支处理
      const { before: leftOfEnd } = splitRunAt(endSpan.run, endOffset);
      endRunForMatch = leftOfEnd;
    }
  } catch (e) {
    // splitRunAt 抛错(如复杂 run:多 <w:t> / <w:br>)→ 规范化为 UNSUPPORTED_OOXML_FEATURE
    const msg = e instanceof Error ? e.message : String(e);
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `replaceTextInParagraph: run structure too complex to split (${msg}). paraId=${op.paraId}`,
      ctx.docId,
      { paraId: op.paraId, opType: 'replaceTextInParagraph' },
    );
  }

  // 5. 收集 [startRunForMatch ... endRunForMatch] 之间所有 sibling run(都在 paragraph 直属层)
  const toRemove: Element[] = [];
  {
    let cur: Node | null = startRunForMatch;
    while (cur) {
      if (cur.nodeType === 1) toRemove.push(cur as Element);
      if (cur === endRunForMatch) break;
      cur = cur.nextSibling;
    }
    if (toRemove.length === 0 || toRemove[toRemove.length - 1] !== endRunForMatch) {
      // 末节点没在兄弟链里遇到 —— 理论不可能,保险抛
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        `replaceTextInParagraph: could not walk from start to end run. paraId=${op.paraId}`,
        ctx.docId,
        { paraId: op.paraId, opType: 'replaceTextInParagraph' },
      );
    }
  }

  // 5a. ANCHOR_CONFLICT 检测(#30):仅在非 trackChanges 模式拦截。
  //     trackChanges=true 时,锚点 marker 会随 toRemove 一起被包进 <w:del>(wrapRunsAsDelete
  //     不碰非 <w:t> 子,保留原结构),与 Word/WPS 原生 trackChanges 行为一致——
  //     用户接受 → 文本+锚点一起删;拒绝 → 一起恢复。
  if (!ctx.options.trackChanges) {
    const rangeAnchors = collectRangeAnchors(startRunForMatch, endRunForMatch);
    if (rangeAnchors.bookmarks.length > 0 || rangeAnchors.comments.length > 0) {
      throw new WordError(
        'ANCHOR_CONFLICT',
        `formatTextInParagraph conflicts with active anchors (paraId=${op.paraId})`,
        ctx.docId,
        {
          paraId: op.paraId,
          affectedBookmarks: rangeAnchors.bookmarks,
          affectedComments: rangeAnchors.comments,
          opType: 'replaceTextInParagraph',
        },
      );
    }
  }

  // 6. 用首 run 的 rPr 深拷贝 + replacement 组装新 run。
  //    replacement 为空表示纯删除，不应额外产生一个空 <w:ins> 修订。
  const hasReplacement = op.replacement.length > 0;
  const newRun = hasReplacement
    ? buildReplacementRun(paragraph, startRunForMatch, op.replacement)
    : null;

  const parent = paragraph;
  const insertionPoint = toRemove[0]!;

  if (ctx.options.trackChanges) {
    // trackChanges=true:匹配区间包成 <w:del>,新 run 包成 <w:ins>
    const attrs: RevisionAttrs = { author: ctx.options.author, date: ctx.options.date };
    const delId = allocateRevisionId(ctx.pkg);
    const delEl = wrapRunsAsDelete(toRemove, delId, attrs);

    // 先在匹配区间位置插入 del；有 replacement 时再插入 ins。
    parent.insertBefore(delEl, insertionPoint);
    if (newRun) {
      const insId = allocateRevisionId(ctx.pkg);
      const insEl = wrapRunsAsInsert([newRun], insId, attrs);
      parent.insertBefore(insEl, insertionPoint);
    }
    // 删掉原始匹配 runs
    for (const r of toRemove) {
      parent.removeChild(r);
    }
  } else {
    // 7. 插入到 toRemove[0] 之前,再把 toRemove 全删
    if (newRun) parent.insertBefore(newRun, insertionPoint);
    for (const r of toRemove) {
      parent.removeChild(r);
    }
  }
}

function formatOneMatch(
  ctx: ExecutorContext,
  paragraph: Element,
  match: Match,
  op: Extract<Op, { type: 'formatTextInParagraph' }>,
): void {
  const runs = getParagraphRuns(paragraph);
  const runSpans: Array<{ run: Element; start: number; end: number }> = [];
  let cursor = 0;
  for (const r of runs) {
    const len = getRunText(r).length;
    runSpans.push({ run: r, start: cursor, end: cursor + len });
    cursor += len;
  }

  const startSpan = runSpans.find((s) => match.start >= s.start && match.start < s.end);
  const endSpan = runSpans.find((s) => match.end > s.start && match.end <= s.end);
  if (!startSpan || !endSpan) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `formatTextInParagraph: match spans runs not reachable via direct <w:r> enumeration ` +
        `(possibly wrapped in hyperlink/ins/del). paraId=${op.paraId}`,
      ctx.docId,
      { paraId: op.paraId, opType: 'formatTextInParagraph' },
    );
  }

  const sameRun = startSpan.run === endSpan.run;
  const endOffset = match.end - endSpan.start;
  let startRunForMatch: Element;
  let endRunForMatch: Element;

  try {
    if (sameRun) {
      const { after: rightAfterEnd } = splitRunAt(startSpan.run, endOffset);
      const startOffsetInRun = match.start - startSpan.start;
      const { after: matchRun } = splitRunAt(startSpan.run, startOffsetInRun);
      startRunForMatch = matchRun ?? startSpan.run;
      endRunForMatch = startRunForMatch;
      void rightAfterEnd;
    } else {
      const startOffsetInRun = match.start - startSpan.start;
      const { after: afterStart } = splitRunAt(startSpan.run, startOffsetInRun);
      startRunForMatch = afterStart ?? startSpan.run;

      const { before: leftOfEnd } = splitRunAt(endSpan.run, endOffset);
      endRunForMatch = leftOfEnd;
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `formatTextInParagraph: run structure too complex to split (${msg}). paraId=${op.paraId}`,
      ctx.docId,
      { paraId: op.paraId, opType: 'formatTextInParagraph' },
    );
  }

  const matchedRuns: Element[] = [];
  {
    let cur: Node | null = startRunForMatch;
    while (cur) {
      if (cur.nodeType === 1) matchedRuns.push(cur as Element);
      if (cur === endRunForMatch) break;
      cur = cur.nextSibling;
    }
    if (matchedRuns.length === 0 || matchedRuns[matchedRuns.length - 1] !== endRunForMatch) {
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        `formatTextInParagraph: could not walk from start to end run. paraId=${op.paraId}`,
        ctx.docId,
        { paraId: op.paraId, opType: 'formatTextInParagraph' },
      );
    }
  }

  if (!ctx.options.trackChanges) {
    const rangeAnchors = collectRangeAnchors(startRunForMatch, endRunForMatch);
    if (rangeAnchors.bookmarks.length > 0 || rangeAnchors.comments.length > 0) {
      throw new WordError(
        'ANCHOR_CONFLICT',
        `paragraph modification conflicts with active anchors (paraId=${op.paraId}); ` +
          `enable trackChanges to wrap the deletion as a revision`,
        ctx.docId,
        {
          paraId: op.paraId,
          affectedBookmarks: rangeAnchors.bookmarks,
          affectedComments: rangeAnchors.comments,
          opType: 'formatTextInParagraph',
        },
      );
    }
  }

  for (const run of matchedRuns) {
    applyRunFormat(run, op.format);
  }
}

function buildReplacementRun(paragraph: Element, seedRun: Element, replacement: string): Element {
  const doc = paragraph.ownerDocument!;
  const run = createW(doc, 'r');

  // 继承 rPr(深拷贝,避免共享引用)
  const seedRPr = firstDirectChildW(seedRun, 'rPr');
  if (seedRPr) {
    run.appendChild(seedRPr.cloneNode(true));
  }

  appendSegmentsToRun(doc, run, replacement);
  return run;
}

// firstDirectChildW / appendTextAsRuns / createW 已提出到 _shared.ts(#29 重构)
function validateTextFormat(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'formatTextInParagraph' }>,
): void {
  if (
    op.format.bold === undefined &&
    op.format.color === undefined &&
    op.format.italic === undefined &&
    op.format.underline === undefined &&
    op.format.highlight === undefined &&
    op.format.fontSize === undefined &&
    op.format.fontFamily === undefined
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatTextInParagraph requires at least one format field',
      ctx.docId,
      { paraId: op.paraId, opType: 'formatTextInParagraph' },
    );
  }
  if (
    op.format.color !== undefined &&
    op.format.color !== 'auto' &&
    !/^[0-9A-Fa-f]{6}$/.test(op.format.color)
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatTextInParagraph color must be "auto" or a 6-digit RRGGBB value',
      ctx.docId,
      { paraId: op.paraId, color: op.format.color, opType: 'formatTextInParagraph' },
    );
  }
  if (op.format.highlight !== undefined && !/^[A-Za-z]+$/.test(op.format.highlight)) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatTextInParagraph highlight must be a Word highlight color name',
      ctx.docId,
      { paraId: op.paraId, highlight: op.format.highlight, opType: 'formatTextInParagraph' },
    );
  }
  if (
    op.format.fontSize !== undefined &&
    (!Number.isFinite(op.format.fontSize) || op.format.fontSize <= 0)
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatTextInParagraph fontSize must be a positive point value',
      ctx.docId,
      { paraId: op.paraId, fontSize: op.format.fontSize, opType: 'formatTextInParagraph' },
    );
  }
  if (op.format.fontFamily !== undefined && op.format.fontFamily.trim() === '') {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatTextInParagraph fontFamily cannot be empty',
      ctx.docId,
      { paraId: op.paraId, opType: 'formatTextInParagraph' },
    );
  }
}

function applyRunFormat(
  run: Element,
  format: {
    bold?: boolean;
    color?: string;
    italic?: boolean;
    underline?: boolean;
    highlight?: string;
    fontSize?: number;
    fontFamily?: string;
  },
): void {
  const rPr = ensureRunProperties(run);
  if (format.bold !== undefined) {
    replaceRunProperty(rPr, 'b', (el) => {
      if (!format.bold) el.setAttributeNS(W_NS, 'w:val', '0');
    });
  }
  if (format.italic !== undefined) {
    replaceRunProperty(rPr, 'i', (el) => {
      if (!format.italic) el.setAttributeNS(W_NS, 'w:val', '0');
    });
  }
  if (format.underline !== undefined) {
    replaceRunProperty(rPr, 'u', (el) => {
      el.setAttributeNS(W_NS, 'w:val', format.underline ? 'single' : 'none');
    });
  }
  if (format.color !== undefined) {
    removeRunProperties(rPr, W14_NS, 'textFill');
    replaceRunProperty(rPr, 'color', (el) => {
      el.setAttributeNS(W_NS, 'w:val', format.color!);
    });
  }
  if (format.highlight !== undefined) {
    replaceRunProperty(rPr, 'highlight', (el) => {
      el.setAttributeNS(W_NS, 'w:val', format.highlight!);
    });
  }
  if (format.fontSize !== undefined) {
    replaceRunProperty(rPr, 'sz', (el) => {
      el.setAttributeNS(W_NS, 'w:val', String(Math.round(format.fontSize! * 2)));
    });
  }
  if (format.fontFamily !== undefined) {
    replaceRunProperty(rPr, 'rFonts', (el) => {
      const family = format.fontFamily!;
      el.setAttributeNS(W_NS, 'w:ascii', family);
      el.setAttributeNS(W_NS, 'w:hAnsi', family);
      el.setAttributeNS(W_NS, 'w:eastAsia', family);
      el.setAttributeNS(W_NS, 'w:cs', family);
    });
  }
}

function ensureRunProperties(run: Element): Element {
  const existing = firstDirectChildW(run, 'rPr');
  if (existing) return existing;
  const rPr = createW(run.ownerDocument!, 'rPr');
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}

function replaceRunProperty(
  rPr: Element,
  localName: 'b' | 'i' | 'u' | 'color' | 'highlight' | 'sz' | 'rFonts',
  configure: (el: Element) => void,
): void {
  removeRunProperties(rPr, W_NS, localName);
  const el = createW(rPr.ownerDocument!, localName);
  configure(el);
  rPr.appendChild(el);
}

function removeRunProperties(rPr: Element, namespaceURI: string, localName: string): void {
  const existing: Element[] = [];
  for (let i = 0; i < rPr.childNodes.length; i++) {
    const child = rPr.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === namespaceURI &&
      (child as Element).localName === localName
    ) {
      existing.push(child as Element);
    }
  }
  for (const child of existing) {
    rPr.removeChild(child);
  }
}
