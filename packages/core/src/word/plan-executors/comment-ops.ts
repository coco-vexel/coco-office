/**
 * 批注 executor —— #35(addComment / deleteComment)。
 *
 * 设计要点:
 *   - 跨 parts 事务:addComment 同时变更 body part + comments.xml +
 *     [Content_Types].xml + word/_rels/document.xml.rels(首次加批注时后两者)。
 *     所有 part 变更在同一 runTransaction 范围内,由 kernel.collectAffectedPartPaths
 *     预声明,外层回滚时一同还原。
 *   - commentId 分配:扫描 pkg 内所有 <w:comment> 的 w:id + 所有 commentRangeStart
 *     的 w:id,取 max + 1(起点 0)。全 pkg 唯一,不依赖装载期 commentIdx(运行时
 *     追加场景下 commentIdx 未同步)。
 *   - scope='all' TextRange:本 PR 抛 UNSUPPORTED_OOXML_FEATURE,留 follow-up。
 *   - deleteComment:三锚点按需移除(null 跳过);ref 所在 wrapper <w:r> 若变空则
 *     也整体删除(避免悬空空 run)。
 *   - comments.xml 删完最后一条时，移除空 part + 清 CT Override + 清 rels Relationship
    （ensureCommentsPart 的对称反向操作；task-comment-part-cleanup）。
 */
import { parseXmlDocument } from 'slimdom';
import {
  getPart,
  createXmlPart,
  setPart,
  markDirty,
  isXmlPart,
  deletePart,
} from '../../pack/index.js';
import { findStoryFor } from '../stories.js';
import { W_NS, W14_NS, XML_NS } from '../namespaces.js';
import type { Op, Range } from '../plan.js';
import { getParagraphRuns, getRunText, splitRunAt } from '../runs.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';

// ─────────── 常量 ───────────

const COMMENTS_PART_PATH = 'word/comments.xml';
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const CONTENT_TYPES_PATH = '[Content_Types].xml';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DOC_RELS_PATH = 'word/_rels/document.xml.rels';

// ─────────── addComment ───────────

export function executeAddComment(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'addComment' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'addComment with trackChanges=true not yet implemented (see #31)',
      ctx.docId,
      { opType: 'addComment' },
    );
  }

  // TextRange scope='all' 暂不支持(留 follow-up)
  if (isTextRange(op.range) && op.range.scope === 'all') {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      "addComment: scope='all' is not yet implemented for TextRange (pending follow-up issue)",
      ctx.docId,
      { opType: 'addComment', scope: 'all' },
    );
  }

  // 1. 查目标段落
  const paragraph = requireParagraph(ctx, op.range.paraId, 'addComment');

  // 2-7. 走共享路径(挂段级锚点 + 写主体 + dirty)
  const commentId = attachCommentToParagraph(ctx, paragraph, op.text, {
    range: op.range,
    author: op.author,
    date: op.date,
  });

  // 8. 记录 ApplyResult
  ctx.result.createdComments.push({ opIndex: ctx.opIndex, commentId });
}

/**
 * 给指定段落挂批注 —— addComment / addTableComment 共享核心。
 *
 * 调用方约定:
 *   - 已自行校验 paragraph 存在;trackChanges 守护也由调用方负责(addTableComment
 *     语义和 addComment 相同,本函数本身不再二次拦截)
 *   - 不写 ApplyResult,由调用方决定要不要往 createdComments 里 push(addTableComment
 *     的 commentId 不返回给调用方,只用于内部锚点串联)
 *   - range 缺省时按 ParagraphRange 处理(把整段挂作锚点,符合 addTableComment 的
 *     "挂在表格首段"语义)
 *
 * 返回新分配的 commentId,供调用方按需上报。
 */
export function attachCommentToParagraph(
  ctx: ExecutorContext,
  paragraph: Element,
  text: string,
  opts: {
    range?: Range;
    author?: string;
    /** ISO 8601 字符串 */
    date?: string;
  } = {},
): number {
  const paraId = paragraph.getAttributeNS(W14_NS, 'paraId') ?? '';
  const range: Range = opts.range ?? { paraId };

  const commentId = allocateCommentId(ctx);

  const author = opts.author ?? ctx.options.author ?? '';
  const dateStr = opts.date ?? ctx.options.date.toISOString();
  const initials = author.length > 0 ? author[0]! : '';

  const wasNew = ensureCommentsPart(ctx);

  const commentsPart = getPart(ctx.pkg, COMMENTS_PART_PATH)!;
  if (!isXmlPart(commentsPart) || !commentsPart.xmlDom) {
    throw new WordError('PACKAGE_CORRUPT', 'comments.xml part is not an XML part', ctx.docId);
  }
  const commentsDoc = commentsPart.xmlDom;

  // body 侧锚点:复用 addComment 的逻辑入口
  insertBodyAnchors(ctx, paragraph, { type: 'addComment', range, text }, commentId);

  insertCommentBody(commentsDoc, commentId, author, dateStr, initials, text);

  markBodyPartDirty(ctx, paragraph);
  markDirty(commentsPart);
  void wasNew;

  return commentId;
}

// ─────────── deleteComment ───────────

export function executeDeleteComment(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'deleteComment' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'deleteComment with trackChanges=true not yet implemented (see #31)',
      ctx.docId,
      { opType: 'deleteComment' },
    );
  }

  // 1. 查 commentIdx
  const entry = ctx.indices.commentIdx.get(op.commentId);
  if (!entry) {
    throw new WordError(
      'COMMENT_NOT_FOUND',
      `comment with id=${op.commentId} not found`,
      ctx.docId,
      { commentId: op.commentId },
    );
  }

  // 2. 移除三锚点
  const { start, end, ref } = entry.anchors;

  // 确定 body part(用于 markDirty);从存在的锚点反查
  const anchorEl = start ?? end ?? ref;
  let bodyPartPath: string | null = null;
  if (anchorEl) {
    const story = findStoryFor(ctx.pkg, anchorEl);
    if (story) bodyPartPath = story.partPath;
  }

  if (start) removeFromParent(start);
  if (end) removeFromParent(end);
  if (ref) {
    // ref 可能包在一个 <w:r> 里;若移除 ref 后 run 只剩 rPr(或变空),删掉整个 run
    const parentRun = ref.parentNode;
    removeFromParent(ref);
    if (
      parentRun &&
      parentRun.nodeType === 1 &&
      (parentRun as Element).namespaceURI === W_NS &&
      (parentRun as Element).localName === 'r'
    ) {
      // 检查 run 是否还有非 rPr 的有意义子元素
      if (isRunEffectivelyEmpty(parentRun as Element)) {
        removeFromParent(parentRun as Element);
      }
    }
  }

  // 3. 移除 <w:comment> 主体
  if (entry.comment) {
    removeFromParent(entry.comment);
    const commentsPart = getPart(ctx.pkg, COMMENTS_PART_PATH);
    if (commentsPart) markDirty(commentsPart);
  }

  // 4. markDirty body part
  if (bodyPartPath) {
    const bodyPart = getPart(ctx.pkg, bodyPartPath);
    if (bodyPart) markDirty(bodyPart);
  }

  // 5. 若 comments.xml 已无任何 <w:comment>，移除空 part + 清理 CT/rels 悬空引用
  cleanupEmptyCommentsPart(ctx);
}

// ─────────── 内部:comments part 清理（ensureCommentsPart 的对称反向）───────────

/**
 * 若 comments.xml 还剩 0 个 `<w:comment>`，执行反向清理：
 *   1. deletePart(comments.xml)
 *   2. 删 [Content_Types].xml 里 PartName="/word/comments.xml" 的 Override
 *   3. 删 word/_rels/document.xml.rels 里 Type=…/comments 的 Relationship
 *
 * 幂等：part 已不存在 / Override 已无 / rel 已无 → 各自跳过，不报错。
 *
 * 导出供 paragraph-structure-ops / table-ops 的 dropCommentById 调用（drop 路径同样
 * 需要在删光最后一条批注后清 part/CT/rels）。
 */
export function cleanupEmptyCommentsPart(ctx: ExecutorContext): void {
  const cp = getPart(ctx.pkg, COMMENTS_PART_PATH);
  if (!cp || !isXmlPart(cp) || !cp.xmlDom) return;
  // 还有 <w:comment> 就不清
  const remaining = cp.xmlDom.getElementsByTagNameNS(W_NS, 'comment').length;
  if (remaining > 0) return;
  // 空了：移除 part + 反向清 CT / rels
  deletePart(ctx.pkg, COMMENTS_PART_PATH);
  removeContentTypesOverride(ctx);
  removeDocRel(ctx);
}

function removeContentTypesOverride(ctx: ExecutorContext): void {
  const ctPart = getPart(ctx.pkg, CONTENT_TYPES_PATH);
  if (!ctPart || !isXmlPart(ctPart) || !ctPart.xmlDom) return;
  const overrides = ctPart.xmlDom.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
  // 快照后删（live collection）
  const victims: Element[] = [];
  for (let i = 0; i < overrides.length; i++) {
    const el = overrides[i];
    if (el && el.getAttribute('PartName') === '/word/comments.xml') victims.push(el);
  }
  for (const el of victims) removeFromParent(el);
  if (victims.length > 0) markDirty(ctPart);
}

function removeDocRel(ctx: ExecutorContext): void {
  const relsPart = getPart(ctx.pkg, DOC_RELS_PATH);
  if (!relsPart || !isXmlPart(relsPart) || !relsPart.xmlDom) return;
  const rels = relsPart.xmlDom.getElementsByTagNameNS(RELS_NS, 'Relationship');
  const victims: Element[] = [];
  for (let i = 0; i < rels.length; i++) {
    const el = rels[i];
    if (el && el.getAttribute('Type') === COMMENTS_RELATIONSHIP_TYPE) victims.push(el);
  }
  for (const el of victims) removeFromParent(el);
  if (victims.length > 0) markDirty(relsPart);
}

// ─────────── 内部:commentId 分配 ───────────

/**
 * 扫整个 pkg 所有 XML part 的 <w:comment> w:id + <w:commentRangeStart> w:id,取 max+1。
 * 基于全 pkg 扫描,而非 commentIdx(运行时追加场景下 commentIdx 仍是装载时快照)。
 */
function allocateCommentId(ctx: ExecutorContext): number {
  let max = -1;
  for (const part of ctx.pkg.parts.values()) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    const comments = part.xmlDom.getElementsByTagNameNS(W_NS, 'comment');
    for (let i = 0; i < comments.length; i++) {
      const v = numAttr(comments[i]!, W_NS, 'id');
      if (v !== null && v > max) max = v;
    }
    const starts = part.xmlDom.getElementsByTagNameNS(W_NS, 'commentRangeStart');
    for (let i = 0; i < starts.length; i++) {
      const v = numAttr(starts[i]!, W_NS, 'id');
      if (v !== null && v > max) max = v;
    }
  }
  return max + 1;
}

// ─────────── 内部:确保 comments.xml 存在,首次时更新 CT / rels ───────────

/**
 * 若 comments.xml 不存在,创建最小结构并更新 Content_Types + rels。
 * 返回 true 表示本次首次创建(调用方据此决定是否 markDirty CT/rels)。
 */
function ensureCommentsPart(ctx: ExecutorContext): boolean {
  if (ctx.pkg.parts.has(COMMENTS_PART_PATH)) return false;

  // 创建 comments.xml
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` + `<w:comments xmlns:w="${W_NS}"/>`;
  const dom = parseXmlDocument(xml) as unknown as Document;
  const newPart = createXmlPart({
    path: COMMENTS_PART_PATH,
    contentType: COMMENTS_CONTENT_TYPE,
    xmlDom: dom,
    // 不传 originalBytes → dirty=true(新建 part 需要序列化)
  });
  setPart(ctx.pkg, newPart);

  // 更新 Content_Types.xml
  updateContentTypes(ctx);

  // 更新 word/_rels/document.xml.rels
  updateDocRels(ctx);

  return true;
}

function updateContentTypes(ctx: ExecutorContext): void {
  const ctPart = getPart(ctx.pkg, CONTENT_TYPES_PATH);
  if (!ctPart || !isXmlPart(ctPart) || !ctPart.xmlDom) return;

  const doc = ctPart.xmlDom;
  // 检查是否已有 Override for comments.xml
  const overrides = doc.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
  for (let i = 0; i < overrides.length; i++) {
    const el = overrides[i]!;
    if (el.getAttribute('PartName') === '/word/comments.xml') return; // 已存在
  }

  // 追加 Override
  const root = doc.documentElement;
  const override = doc.createElementNS(CONTENT_TYPES_NS, 'Override');
  override.setAttribute('PartName', '/word/comments.xml');
  override.setAttribute('ContentType', COMMENTS_CONTENT_TYPE);
  root.appendChild(override);
  markDirty(ctPart);
}

function updateDocRels(ctx: ExecutorContext): void {
  const relsPart = getPart(ctx.pkg, DOC_RELS_PATH);
  if (!relsPart || !isXmlPart(relsPart) || !relsPart.xmlDom) return;

  const doc = relsPart.xmlDom;
  // 检查是否已有 comments rel
  const rels = doc.getElementsByTagNameNS(RELS_NS, 'Relationship');
  for (let i = 0; i < rels.length; i++) {
    if (rels[i]!.getAttribute('Type') === COMMENTS_RELATIONSHIP_TYPE) return; // 已存在
  }

  // 分配新 rId:扫所有 Id="rId\d+" 取 max+1
  let maxRId = 0;
  for (let i = 0; i < rels.length; i++) {
    const id = rels[i]!.getAttribute('Id') ?? '';
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = parseInt(m[1]!, 10);
      if (n > maxRId) maxRId = n;
    }
  }
  const newRId = `rId${maxRId + 1}`;

  const root = doc.documentElement;
  const rel = doc.createElementNS(RELS_NS, 'Relationship');
  rel.setAttribute('Id', newRId);
  rel.setAttribute('Type', COMMENTS_RELATIONSHIP_TYPE);
  rel.setAttribute('Target', 'comments.xml');
  root.appendChild(rel);
  markDirty(relsPart);
}

// ─────────── 内部:body 锚点插入 ───────────

function insertBodyAnchors(
  ctx: ExecutorContext,
  paragraph: Element,
  op: Extract<Op, { type: 'addComment' }>,
  commentId: number,
): void {
  const doc = paragraph.ownerDocument!;
  const range = op.range;

  if (isTextRange(range)) {
    insertTextRangeAnchors(ctx, paragraph, doc, range, commentId);
  } else {
    insertParagraphRangeAnchors(paragraph, doc, commentId);
  }
}

/**
 * ParagraphRange:段起始(紧跟 pPr 之后)插 commentRangeStart,
 * 段末插 commentRangeEnd + <w:r><w:commentReference></w:r>。
 */
function insertParagraphRangeAnchors(paragraph: Element, doc: Document, commentId: number): void {
  const rangeStart = createCommentAnchor(doc, 'commentRangeStart', commentId);
  const rangeEnd = createCommentAnchor(doc, 'commentRangeEnd', commentId);
  const refRun = createCommentRefRun(doc, commentId);

  // rangeStart:紧跟 pPr 之后;若无 pPr 则插到 firstChild
  const pPr = firstDirectChildW(paragraph, 'pPr');
  if (pPr && pPr.nextSibling) {
    paragraph.insertBefore(rangeStart, pPr.nextSibling);
  } else if (pPr) {
    paragraph.appendChild(rangeStart);
  } else {
    paragraph.insertBefore(rangeStart, paragraph.firstChild);
  }

  // rangeEnd + refRun:追加到段末
  paragraph.appendChild(rangeEnd);
  paragraph.appendChild(refRun);
}

/**
 * TextRange:找段内第一个命中的文本区间,在 run 边界插入锚点。
 * scope='all' 已在调用方提前拒绝。
 */
function insertTextRangeAnchors(
  ctx: ExecutorContext,
  paragraph: Element,
  doc: Document,
  range: { paraId: string; text: string; scope?: 'first' | 'all' },
  commentId: number,
): void {
  const runs = getParagraphRuns(paragraph);
  const fullText = runs.map((r) => getRunText(r)).join('');
  const searchText = range.text;

  if (searchText.length === 0) {
    throw new WordError(
      'TEXT_NOT_FOUND_IN_PARA',
      `addComment: 'range.text' cannot be empty`,
      ctx.docId,
      { paraId: range.paraId },
    );
  }

  const matchStart = fullText.indexOf(searchText);
  if (matchStart < 0) {
    throw new WordError(
      'TEXT_NOT_FOUND_IN_PARA',
      `addComment: text '${searchText}' not found in paragraph ${range.paraId}`,
      ctx.docId,
      { paraId: range.paraId, search: searchText },
    );
  }
  const matchEnd = matchStart + searchText.length;

  // 构造 (run, runStart, runEnd) 映射
  const runSpans: Array<{ run: Element; start: number; end: number }> = [];
  let cursor = 0;
  for (const r of runs) {
    const len = getRunText(r).length;
    runSpans.push({ run: r, start: cursor, end: cursor + len });
    cursor += len;
  }

  const startSpan = runSpans.find((s) => matchStart >= s.start && matchStart < s.end);
  const endSpan = runSpans.find((s) => matchEnd > s.start && matchEnd <= s.end);

  if (!startSpan || !endSpan) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `addComment: text match spans runs not reachable via direct <w:r> enumeration. paraId=${range.paraId}`,
      ctx.docId,
      { paraId: range.paraId, opType: 'addComment' },
    );
  }

  // 切 start run:在 matchStart - startSpan.start 处切
  const startOffsetInRun = matchStart - startSpan.start;
  let firstMatchRun: Element;
  try {
    if (startOffsetInRun === 0) {
      firstMatchRun = startSpan.run;
    } else {
      const { after } = splitRunAt(startSpan.run, startOffsetInRun);
      firstMatchRun = after ?? startSpan.run;
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `addComment: run too complex to split at start (${msg}). paraId=${range.paraId}`,
      ctx.docId,
      { paraId: range.paraId, opType: 'addComment' },
    );
  }

  // 切 end run:在 matchEnd - endSpan.start 处切(需重查,因为 startSpan.run 可能已变)
  // 如果 start 和 end 在同一 run 且 start 切分了,endSpan.run 引用还是旧的 run
  // 需要判断:若 startSpan.run === endSpan.run 且 startOffsetInRun > 0,endSpan 已被切成 firstMatchRun
  let lastMatchRun: Element;
  try {
    const endRunCurrent =
      startSpan.run === endSpan.run && startOffsetInRun > 0 ? firstMatchRun : endSpan.run;
    // matchEnd 在 endRunCurrent 内的偏移:
    //   若同 run 且 start>0 切了:endRunCurrent 从 matchStart 开始,所以 matchEnd 相对其起点
    //     = matchEnd - matchStart
    //   否则:matchEnd - endSpan.start
    const endOffsetInRun =
      startSpan.run === endSpan.run && startOffsetInRun > 0
        ? matchEnd - matchStart
        : matchEnd - endSpan.start;

    if (endOffsetInRun >= getRunText(endRunCurrent).length) {
      // 恰好到 run 末尾,不需要切
      lastMatchRun = endRunCurrent;
    } else {
      const { before } = splitRunAt(endRunCurrent, endOffsetInRun);
      lastMatchRun = before;
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `addComment: run too complex to split at end (${msg}). paraId=${range.paraId}`,
      ctx.docId,
      { paraId: range.paraId, opType: 'addComment' },
    );
  }

  // 插入锚点:rangeStart 在 firstMatchRun 之前,rangeEnd + refRun 在 lastMatchRun 之后
  const rangeStart = createCommentAnchor(doc, 'commentRangeStart', commentId);
  const rangeEnd = createCommentAnchor(doc, 'commentRangeEnd', commentId);
  const refRun = createCommentRefRun(doc, commentId);

  paragraph.insertBefore(rangeStart, firstMatchRun);

  const afterLastRun = lastMatchRun.nextSibling;
  if (afterLastRun) {
    paragraph.insertBefore(rangeEnd, afterLastRun);
    paragraph.insertBefore(refRun, afterLastRun);
  } else {
    paragraph.appendChild(rangeEnd);
    paragraph.appendChild(refRun);
  }
}

// ─────────── 内部:comment 主体插入 ───────────

function insertCommentBody(
  commentsDoc: Document,
  commentId: number,
  author: string,
  date: string,
  initials: string,
  text: string,
): void {
  const comment = commentsDoc.createElementNS(W_NS, 'w:comment');
  comment.setAttributeNS(W_NS, 'w:id', String(commentId));
  comment.setAttributeNS(W_NS, 'w:author', author);
  comment.setAttributeNS(W_NS, 'w:date', date);
  if (initials) comment.setAttributeNS(W_NS, 'w:initials', initials);

  // 批注正文:按 \n 拆段,\t → <w:tab/>,其余 → <w:t>
  // 简洁策略:单段落 + \n 映射 <w:br>(与 paragraph-ops 一致)
  const p = commentsDoc.createElementNS(W_NS, 'w:p');
  const r = commentsDoc.createElementNS(W_NS, 'w:r');
  appendCommentTextSegments(commentsDoc, r, text);
  p.appendChild(r);
  comment.appendChild(p);

  // 追加到 <w:comments> 根
  const root = commentsDoc.documentElement;
  root.appendChild(comment);
}

function appendCommentTextSegments(doc: Document, run: Element, text: string): void {
  if (text.length === 0) {
    run.appendChild(createT(doc, ''));
    return;
  }
  let buf = '';
  const flush = (): void => {
    if (buf.length > 0) {
      run.appendChild(createT(doc, buf));
      buf = '';
    }
  };
  for (const ch of text) {
    if (ch === '\n') {
      flush();
      run.appendChild(createW(doc, 'br'));
    } else if (ch === '\t') {
      flush();
      run.appendChild(createW(doc, 'tab'));
    } else {
      buf += ch;
    }
  }
  flush();
}

// ─────────── 内部:DOM 工具 ───────────

function createW(doc: Document, localName: string): Element {
  return doc.createElementNS(W_NS, `w:${localName}`);
}

function createT(doc: Document, text: string): Element {
  const t = createW(doc, 't');
  t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  t.textContent = text;
  return t;
}

function createCommentAnchor(doc: Document, localName: string, commentId: number): Element {
  const el = createW(doc, localName);
  el.setAttributeNS(W_NS, 'w:id', String(commentId));
  return el;
}

function createCommentRefRun(doc: Document, commentId: number): Element {
  const run = createW(doc, 'r');
  const ref = createW(doc, 'commentReference');
  ref.setAttributeNS(W_NS, 'w:id', String(commentId));
  run.appendChild(ref);
  return run;
}

function firstDirectChildW(el: Element, localName: string): Element | null {
  const children = el.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === localName
    ) {
      return c as Element;
    }
  }
  return null;
}

function removeFromParent(el: Element | Node): void {
  if (el.parentNode) el.parentNode.removeChild(el);
}

/**
 * 判断 <w:r> 元素是否已无实质性内容(只剩 rPr 或彻底为空)。
 * 用于 deleteComment 后清理空 wrapper run。
 */
function isRunEffectivelyEmpty(run: Element): boolean {
  const children = run.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (!c || c.nodeType !== 1) continue;
    const localName = (c as Element).localName;
    // rPr 是运行属性,不算"内容"
    if (localName === 'rPr') continue;
    return false;
  }
  return true;
}

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

function markBodyPartDirty(ctx: ExecutorContext, paragraph: Element): void {
  const story = findStoryFor(ctx.pkg, paragraph);
  if (!story) return;
  const part = getPart(ctx.pkg, story.partPath);
  if (part) markDirty(part);
}

function numAttr(el: Element, ns: string, local: string): number | null {
  const v = el.getAttributeNS(ns, local);
  if (v === null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

/** Range 类型判断:有 text 字段 → TextRange */
function isTextRange(
  range: Range,
): range is { paraId: string; text: string; scope?: 'first' | 'all' } {
  return 'text' in range;
}
