/**
 * insertTableFrom executor —— task-20260616-table-content-driven。
 *
 * 把一张**来自另一文档**的 `<w:tbl>`（已序列化为 XML 字符串）verbatim 搬进当前输出
 * 文档，插到 anchorParaId 段之后。与 insertImageAfter 同构（图按字节跨文档搬、表按
 * OOXML 子树跨文档搬），是"数据表 = 内容块"（§13.9）的底座 op。
 *
 * 为什么是「序列化 XML + 预解析图字节」而非「直接传 Element」：
 *   - op 必须跨端安全（worker 模式下 op 经结构化克隆传输，绝不能带 foreign DOM 节点）。
 *   - 调用方（typeset，inline 模式持有两个 doc）在源文档侧把 `<w:tbl>` 子树
 *     `serializeToWellFormedString` 成字符串、把每个格内图按原 r:embed 反查出原始字节，
 *     打包成自包含 payload 传进来。executor 只吃 string + Uint8Array。
 *
 * 写链：
 *   1. anchor 定位（paraIdx → PARA_NOT_FOUND）。
 *   2. parseXmlFragment(tableXml) → importNode 到 host doc（命名空间随子树带入，已验证）。
 *   3. 校验根是 <w:tbl>（否则 INVALID_ARGUMENT，绝不把异物塞进文档）。
 *   4. 重映射格内图：遍历表内每个 <a:blip @r:embed>，按原 rId 找 op.images 里对应字节，
 *      在 host doc 新建 media part + image rel + content-type，把 embed 改写成新 rId。
 *      原 rId 在 op.images 里找不到 → 该 blip 的 embed 清掉 + warn（绝不留悬空引用致文档损坏；
 *      文字/结构仍 verbatim 保真，只这一张格内图降级丢失，调用方侧应保证 images 齐全）。
 *   5. 剥离表内锚点（书签/批注 5 类），重分配每个 <w:p> 的 paraId（全局唯一）。
 *   6. 样式归一（可选）：op.styleTblPrXml 提供时，用它整体替换搬入表的 <w:tblPr>
 *      （tblStyle / 边框 / tblLook / 宽度 → 归一到模板数据表样式，覆盖内容表自带样式）。
 *   7. 插到 anchor 之后，markDirty。
 *
 * 内容零丢红线：表的行/列/值/合并/嵌套表/格内图全部 verbatim 来自 op.tableXml（源原表），
 *   executor 不重建、不裁剪、不对位——只搬运 + 重映射引用 + 套样式。
 */
import { parseXmlFragment } from 'slimdom';
import { createBinaryPart, getPart, isXmlPart, markDirty, setPart } from '../../pack/index.js';
import { generateParaId } from '../para-id.js';
import { A_NS, R_NS, W_NS, W14_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { findStoryFor } from '../stories.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';
import {
  addImageRelForHost,
  allocateMediaPath,
  CONTENT_TYPE_TO_EXT,
  ensureImageContentType,
  relsPathForHost,
} from './insert-image-ops.js';

/** 单张格内图的预解析字节（按源文档原始 r:embed rId 索引）。 */
export interface InsertTableImage {
  /** 源表 <a:blip @r:embed> 的原始 rId（host 内会被改写成新 rId）。 */
  originalRId: string;
  /** 原始图片字节（来自源文档 media part）。 */
  bytes: Uint8Array;
  /** 图片 content-type（image/png 等）。 */
  contentType: string;
}

export function executeInsertTableFrom(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'insertTableFrom' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'insertTableFrom with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'insertTableFrom' },
    );
  }

  // 1. anchor 定位
  const anchor = ctx.indices.paraIdx.get(op.anchorParaId.toUpperCase());
  if (!anchor) {
    throw new WordError(
      'PARA_NOT_FOUND',
      `paragraph with paraId='${op.anchorParaId}' not found (op: insertTableFrom)`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: 'insertTableFrom' },
    );
  }
  const parent = anchor.parentNode;
  if (!parent) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `anchor paragraph '${op.anchorParaId}' has no parentNode`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: 'insertTableFrom' },
    );
  }

  const story = findStoryFor(ctx.pkg, anchor);
  if (!story) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `anchor paragraph '${op.anchorParaId}' not in any story`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: 'insertTableFrom' },
    );
  }
  const hostPartPath = story.partPath;

  // host doc（用于 importNode + createElementNS）
  const hostPart = getPart(ctx.pkg, hostPartPath);
  if (!hostPart || !isXmlPart(hostPart) || !hostPart.xmlDom) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `host part '${hostPartPath}' missing or not XML`,
      ctx.docId,
      { opName: 'insertTableFrom' },
    );
  }
  const hostDoc = hostPart.xmlDom;

  // 2. 解析 + 导入表子树
  let imported: Element;
  try {
    const frag = parseXmlFragment(op.tableXml);
    // 取首个元素子节点（应是 <w:tbl>）
    let root: Element | null = null;
    for (let i = 0; i < frag.childNodes.length; i++) {
      const c = frag.childNodes[i] as Element;
      if (c && c.nodeType === 1) {
        root = c;
        break;
      }
    }
    if (!root) {
      throw new WordError(
        'INVALID_ARGUMENT',
        'insertTableFrom: tableXml has no element root',
        ctx.docId,
        { opName: 'insertTableFrom' },
      );
    }
    imported = hostDoc.importNode(root, true) as unknown as Element;
  } catch (e) {
    if (e instanceof WordError) throw e;
    throw new WordError(
      'INVALID_ARGUMENT',
      `insertTableFrom: failed to parse/import tableXml: ${(e as Error).message}`,
      ctx.docId,
      { opName: 'insertTableFrom' },
    );
  }

  // 3. 校验根是 <w:tbl>
  if (imported.namespaceURI !== W_NS || imported.localName !== 'tbl') {
    throw new WordError(
      'INVALID_ARGUMENT',
      `insertTableFrom: root element is '${imported.localName}', expected 'w:tbl'`,
      ctx.docId,
      { opName: 'insertTableFrom' },
    );
  }

  // 4. 重映射格内图 blip r:embed（跨文档：原 rId → host 新建 media + rel → 新 rId）
  const relsPath = relsPathForHost(hostPartPath);
  const imageByRId = new Map<string, InsertTableImage>();
  for (const img of op.images ?? []) imageByRId.set(img.originalRId, img);
  // 同一原 rId 在表里可能被多个 blip 引用，但语义上只该建一次 media；这里缓存原→新映射。
  const remapped = new Map<string, string>();
  let mediaCreated = false;

  const blips = imported.getElementsByTagNameNS(A_NS, 'blip');
  const blipSnapshot: Element[] = [];
  for (let i = 0; i < blips.length; i++) {
    const b = blips[i];
    if (b) blipSnapshot.push(b);
  }
  for (const blip of blipSnapshot) {
    const origRId = blip.getAttributeNS(R_NS, 'embed');
    if (!origRId) continue;
    let newRId = remapped.get(origRId);
    if (!newRId) {
      const img = imageByRId.get(origRId);
      if (!img) {
        // 调用方未提供该图字节 → 清掉悬空引用（绝不留无效 rId 致文档打不开），warn 记录降级。
        console.warn(
          `[word] insertTableFrom: cell image rId='${origRId}' not in op.images; ` +
            `dropping blip embed (degrade, docId: ${ctx.docId})`,
        );
        blip.removeAttributeNS(R_NS, 'embed');
        continue;
      }
      const ext = CONTENT_TYPE_TO_EXT[img.contentType];
      if (!ext) {
        console.warn(
          `[word] insertTableFrom: unsupported cell image contentType '${img.contentType}' ` +
            `(rId='${origRId}'); dropping blip embed (docId: ${ctx.docId})`,
        );
        blip.removeAttributeNS(R_NS, 'embed');
        continue;
      }
      const mediaPath = allocateMediaPath(ctx.pkg, ext);
      const mediaPart = createBinaryPart({
        path: mediaPath,
        contentType: img.contentType,
        bytes: img.bytes,
      });
      setPart(ctx.pkg, mediaPart);
      markDirty(mediaPart);
      newRId = addImageRelForHost(ctx.pkg, relsPath, mediaPath);
      ensureImageContentType(ctx.pkg, ext, img.contentType);
      remapped.set(origRId, newRId);
      mediaCreated = true;
    }
    blip.setAttributeNS(R_NS, 'r:embed', newRId);
  }

  // 5. 剥离锚点 + 重分配 paraId
  stripTableAnchors(imported);
  reassignClonedParaIds(ctx, imported);

  // 6. 样式归一（可选）：用模板数据表的 <w:tblPr> 整体替换搬入表的 <w:tblPr>，
  //    把 tblStyle / 边框 / tblLook / 宽度归一到模板（覆盖内容表自带样式）。
  if (op.styleTblPrXml) {
    overrideTblPr(ctx, hostDoc, imported, op.styleTblPrXml);
  }

  // 7. 插到 anchor 之后
  parent.insertBefore(imported, anchor.nextSibling);

  // 同步 paraIdx + tblIdx（保证同 plan 内后续 op 可查到新段/新表）。
  syncIndicesAfterInsert(ctx, imported);

  // 8. markDirty
  markDirty(hostPart);
  if (mediaCreated) {
    const relsPart = getPart(ctx.pkg, relsPath);
    if (relsPart) markDirty(relsPart);
  }
}

// ─────────── helpers ───────────

/**
 * 用模板 <w:tblPr> 替换搬入表的 <w:tblPr>（样式归一）。
 *
 * `<w:tblPr>` 持有 tblStyle（样式引用）/ tblBorders / tblLook / tblW（宽度）/ jc 等表级
 * 格式——整体替换即把内容表的表级外观换成模板数据表的外观。单元格字体随 tblStyle 引用走，
 * 故表级替换 + cell 自带 tcPr（保结构）已能让产物外观贴近模板。
 *
 * styleTblPrXml 解析失败 / 非 <w:tblPr> → 跳过（不崩，宁可保留内容表原样式也不丢内容/损坏）。
 */
function overrideTblPr(
  ctx: ExecutorContext,
  hostDoc: Document,
  tbl: Element,
  styleTblPrXml: string,
): void {
  let newTblPr: Element | null = null;
  try {
    const frag = parseXmlFragment(styleTblPrXml);
    for (let i = 0; i < frag.childNodes.length; i++) {
      const c = frag.childNodes[i] as Element;
      if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tblPr') {
        newTblPr = hostDoc.importNode(c, true) as unknown as Element;
        break;
      }
    }
  } catch (e) {
    console.warn(
      `[word] insertTableFrom: parse styleTblPrXml failed, keep original tblPr: ${(e as Error).message}`,
    );
    return;
  }
  if (!newTblPr) return;

  // 找搬入表的原 <w:tblPr>（必是 tbl 首个直属子；无则插到首位）。
  let oldTblPr: Element | null = null;
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const c = tbl.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tblPr') {
      oldTblPr = c;
      break;
    }
  }
  if (oldTblPr) {
    tbl.replaceChild(newTblPr, oldTblPr);
  } else {
    tbl.insertBefore(newTblPr, tbl.firstChild);
  }
  void ctx;
}

/**
 * 同步 paraIdx + tblIdx：把搬入表内所有 <w:p> 与表自身登记进索引，
 * 保证同一 plan 内后续 op（如再对这张表填格/查段）能命中。
 *
 * 注：applyPlan finally 块会 buildIndices 全量重建，这里只为同 plan 内后续 op 的可见性。
 */
function syncIndicesAfterInsert(ctx: ExecutorContext, tbl: Element): void {
  // <w:p> → paraIdx
  const paras = tbl.getElementsByTagNameNS(W_NS, 'p');
  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    if (!p) continue;
    const pid = p.getAttributeNS(W14_NS, 'paraId');
    if (pid) ctx.indices.paraIdx.set(pid.toUpperCase(), p);
  }
  // 表自身 → tblIdx（分配下一个 tbl-NNNN，仅同 plan 内可见；finally 全量重建会规整）。
  let maxN = 0;
  for (const id of ctx.indices.tblIdx.keys()) {
    const m = /^tbl-(\d+)$/.exec(id);
    if (m) {
      const n = parseInt(m[1]!, 10);
      if (n > maxN) maxN = n;
    }
  }
  ctx.indices.tblIdx.set(`tbl-${String(maxN + 1).padStart(4, '0')}`, tbl);
}

/**
 * 重分配子树内所有 `<w:p>` 的 paraId（全局唯一）。与 table-ops 的同名逻辑同构，
 * 但跨文档搬入的表 paraId 可能与 host 既有撞，故必须重分配。
 */
function reassignClonedParaIds(ctx: ExecutorContext, node: Element): void {
  const seen = new Set(ctx.indices.paraIdx.keys());
  const paragraphs = node.getElementsByTagNameNS(W_NS, 'p');
  const snapshot: Element[] = [];
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (p) snapshot.push(p);
  }
  for (const p of snapshot) {
    const newId = generateParaId(seen);
    p.setAttributeNS(W14_NS, 'w14:paraId', newId);
  }
}

/**
 * 剥离表子树内所有锚点（书签/批注 5 类）。跨文档搬入的表里若带源文档的书签/批注引用，
 * 在 host 里是悬空引用（host 无对应 comments.xml 条目 / bookmarkIdx）→ 必须剥离，
 * 否则产物含无效锚点。与 table-ops.stripAnchorsFromClone 同构。
 */
function stripTableAnchors(tbl: Element): void {
  const anchorLocals = [
    'bookmarkStart',
    'bookmarkEnd',
    'commentRangeStart',
    'commentRangeEnd',
    'commentReference',
  ];
  for (const local of anchorLocals) {
    const list = tbl.getElementsByTagNameNS(W_NS, local);
    const toRemove: Element[] = [];
    for (let i = 0; i < list.length; i++) {
      const el = list[i];
      if (el) toRemove.push(el);
    }
    for (const el of toRemove) {
      if (el.parentNode) el.parentNode.removeChild(el);
    }
  }
}
