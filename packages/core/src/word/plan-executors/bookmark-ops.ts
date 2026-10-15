/**
 * 书签 executor —— #34(createBookmark / deleteBookmark)。
 *
 * 设计要点:
 *   - 段枚举走 paraIdx(O(1));bookmarkIdx 做重名检查(O(1))
 *   - w:id 分配:扫全包所有 part 的 <w:bookmarkStart> 取 max(w:id)+1,
 *     跨 part 保证 id 唯一(W_NS 枚举,禁用无谓词 XPath)
 *   - ParagraphRange:bookmarkStart 置于 pPr 之后、Run 之前;bookmarkEnd 置于段末尾
 *   - TextRange:走 splitRunAt 切 run 后插 start/end(复杂 run → UNSUPPORTED_OOXML_FEATURE)
 *   - trackChanges=true:本 PR 拒绝(#31 再做)
 *   - 校验:name 正则 `^[A-Za-z_][A-Za-z0-9_]{0,39}$`;重名 → BOOKMARK_NAME_EXISTS
 *   - deleteBookmark:移除 start + end element,markDirty 对应 part
 *   - 索引同步:executor 内同步更新 bookmarkIdx(add/delete)
 */
import { getPart, isXmlPart, listParts, markDirty } from '../../pack/index.js';
import { W_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { getParagraphRuns, getRunText, splitRunAt } from '../runs.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';

// 书签名合法正则:首字符为字母或下划线,后跟最多 39 个字母/数字/下划线 → 总长 1..40
const BOOKMARK_NAME_RE = /^[A-Za-z_][A-Za-z0-9_]{0,39}$/;

// ─────────── createBookmark ───────────

export function executeCreateBookmark(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'createBookmark' }>,
): void {
  // 1. trackChanges 拒绝(#31)
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'createBookmark with trackChanges=true not yet implemented (see #31)',
      ctx.docId,
      { opType: 'createBookmark', name: op.name },
    );
  }

  // 2. 校验 name
  validateBookmarkName(ctx, op.name);

  // 3. 重名检查
  if (ctx.indices.bookmarkIdx.has(op.name)) {
    throw new WordError(
      'BOOKMARK_NAME_EXISTS',
      `bookmark name '${op.name}' already exists`,
      ctx.docId,
      { name: op.name },
    );
  }

  // 4. 查 paraId 段
  const range = op.range;
  const paraId = range.paraId.toUpperCase();
  const paragraph = ctx.indices.paraIdx.get(paraId);
  if (!paragraph) {
    throw new WordError(
      'PARA_NOT_FOUND',
      `paragraph with paraId='${op.range.paraId}' not found (op: createBookmark)`,
      ctx.docId,
      { paraId: op.range.paraId, opName: 'createBookmark' },
    );
  }

  // 5. 分配 w:id(跨全 pkg 所有 part)
  const newId = allocateBookmarkId(ctx);

  // 6. 分 ParagraphRange / TextRange 插入
  const doc = paragraph.ownerDocument;
  if (!doc) {
    throw new WordError('PACKAGE_CORRUPT', 'paragraph element has no ownerDocument', ctx.docId);
  }

  let startEl: Element;
  let endEl: Element;

  if (isTextRange(range)) {
    // TextRange:splitRunAt 模式
    const result = insertTextRangeBookmark(ctx, doc, paragraph, op.name, newId, range.text);
    startEl = result.start;
    endEl = result.end;
  } else {
    // ParagraphRange:包住整段内容
    startEl = createBookmarkStart(doc, op.name, newId);
    endEl = createBookmarkEnd(doc, newId);
    insertParagraphRangeBookmark(paragraph, startEl, endEl);
  }

  // 7. 更新 ApplyResult
  // createBookmark 一定落在单段内(ParagraphRange 包整段、TextRange 走 splitRunAt 也
  // 不跨段,见本文件 §insertTextRangeBookmark),所以 from === to。
  ctx.result.createdBookmarks.push({
    opIndex: ctx.opIndex,
    name: op.name,
    range: { fromParaId: op.range.paraId, toParaId: op.range.paraId },
  });

  // 8. 更新索引
  const partPath = findPartPathForElement(ctx, paragraph);
  ctx.indices.bookmarkIdx.set(op.name, { start: startEl, end: endEl, partPath });

  // 9. markDirty
  markPartDirtyByPath(ctx, partPath, 'createBookmark');
}

// ─────────── deleteBookmark ───────────

export function executeDeleteBookmark(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'deleteBookmark' }>,
): void {
  // 1. trackChanges 拒绝(#31)
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'deleteBookmark with trackChanges=true not yet implemented (see #31)',
      ctx.docId,
      { opType: 'deleteBookmark', name: op.name },
    );
  }

  // 2. 查 bookmarkIdx
  const entry = ctx.indices.bookmarkIdx.get(op.name);
  if (!entry) {
    throw new WordError('BOOKMARK_NOT_FOUND', `bookmark '${op.name}' not found`, ctx.docId, {
      name: op.name,
    });
  }

  // 3. 移除 bookmarkStart
  const startParent = entry.start.parentNode;
  if (startParent) {
    startParent.removeChild(entry.start);
  }

  // 4. 移除 bookmarkEnd(若存在)
  if (entry.end) {
    const endParent = entry.end.parentNode;
    if (endParent) {
      endParent.removeChild(entry.end);
    }
  }

  // 5. 更新索引
  ctx.indices.bookmarkIdx.delete(op.name);

  // 6. markDirty 对应 part
  markPartDirtyByPath(ctx, entry.partPath, 'deleteBookmark');
}

// ─────────── 内部:校验 ───────────

function validateBookmarkName(ctx: ExecutorContext, name: string): void {
  if (name.length === 0) {
    throw new WordError('BOOKMARK_NAME_INVALID', 'bookmark name cannot be empty', ctx.docId, {
      name,
    });
  }
  if (name.length > 40) {
    throw new WordError(
      'BOOKMARK_NAME_INVALID',
      `bookmark name is too long (${name.length} chars, max 40)`,
      ctx.docId,
      { name, length: name.length },
    );
  }
  if (!BOOKMARK_NAME_RE.test(name)) {
    throw new WordError(
      'BOOKMARK_NAME_INVALID',
      `bookmark name '${name}' contains invalid characters or starts with an invalid character (must match /^[A-Za-z_][A-Za-z0-9_]{0,39}$/)`,
      ctx.docId,
      { name },
    );
  }
}

// ─────────── 内部:id 分配 ───────────

/** 扫全 pkg 所有 part 的 bookmarkStart w:id,返回 max+1(起点 0) */
function allocateBookmarkId(ctx: ExecutorContext): number {
  let maxId = -1;
  for (const part of listParts(ctx.pkg)) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    const starts = part.xmlDom.getElementsByTagNameNS(W_NS, 'bookmarkStart');
    for (let i = 0; i < starts.length; i++) {
      const el = starts[i]!;
      const idStr = el.getAttributeNS(W_NS, 'id');
      if (idStr === null) continue;
      const n = parseInt(idStr, 10);
      if (!isNaN(n) && n > maxId) maxId = n;
    }
  }
  return maxId + 1;
}

// ─────────── 内部:元素创建 ───────────

function createBookmarkStart(doc: Document, name: string, id: number): Element {
  const el = doc.createElementNS(W_NS, 'w:bookmarkStart');
  el.setAttributeNS(W_NS, 'w:id', String(id));
  el.setAttributeNS(W_NS, 'w:name', name);
  return el;
}

function createBookmarkEnd(doc: Document, id: number): Element {
  const el = doc.createElementNS(W_NS, 'w:bookmarkEnd');
  el.setAttributeNS(W_NS, 'w:id', String(id));
  return el;
}

// ─────────── 内部:ParagraphRange 插入 ───────────

/**
 * ParagraphRange 插入:
 *   - bookmarkStart 放在 pPr 之后,第一个内容节点之前
 *   - bookmarkEnd 放在段的最后一个子节点之后
 */
function insertParagraphRangeBookmark(paragraph: Element, startEl: Element, endEl: Element): void {
  // 找到 pPr 之后的第一个内容节点作为插入点
  const insertBefore = findContentInsertPoint(paragraph);
  if (insertBefore) {
    paragraph.insertBefore(startEl, insertBefore);
  } else {
    // 段落完全为空(无 pPr 也无内容):追加到末尾
    paragraph.appendChild(startEl);
  }
  // end 追加到最后
  paragraph.appendChild(endEl);
}

/**
 * 找到第一个"内容"插入点:pPr 之后、第一个非 pPr 元素之前。
 * 返回 null 表示段落除 pPr 外无其他子节点。
 */
function findContentInsertPoint(paragraph: Element): Node | null {
  const children = paragraph.childNodes;
  let pastPPr = false;
  for (let i = 0; i < children.length; i++) {
    const child = children[i]!;
    if (child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI === W_NS && el.localName === 'pPr') {
      pastPPr = true;
      continue;
    }
    // 第一个非 pPr 元素
    if (pastPPr || i === 0) {
      return child;
    }
  }
  return null;
}

// ─────────── 内部:TextRange 插入 ───────────

function isTextRange(range: {
  paraId: string;
  text?: string;
}): range is { paraId: string; text: string; scope?: string } {
  return typeof (range as { text?: string }).text === 'string';
}

interface TextRangeBookmarkResult {
  start: Element;
  end: Element;
}

/**
 * TextRange 书签插入:
 *   1. 拼段内所有 <w:r> 的 <w:t> 文本
 *   2. indexOf 找第一处 match
 *   3. splitRunAt 切出 matchRun 范围
 *   4. 在 matchRun[0] 之前插 bookmarkStart,在 matchRun[last] 之后插 bookmarkEnd
 */
function insertTextRangeBookmark(
  ctx: ExecutorContext,
  doc: Document,
  paragraph: Element,
  name: string,
  id: number,
  searchText: string,
): TextRangeBookmarkResult {
  const runs = getParagraphRuns(paragraph);

  // 构建全文 + run 偏移映射
  const runSpans: Array<{ run: Element; start: number; end: number }> = [];
  let cursor = 0;
  for (const r of runs) {
    const len = getRunText(r).length;
    runSpans.push({ run: r, start: cursor, end: cursor + len });
    cursor += len;
  }
  const fullText = runSpans.map((s) => getRunText(s.run)).join('');

  // indexOf 找第一处
  const matchStart = fullText.indexOf(searchText);
  if (matchStart < 0) {
    throw new WordError(
      'TEXT_NOT_FOUND_IN_PARA',
      `text '${searchText}' not found in paragraph (op: createBookmark)`,
      ctx.docId,
      { name, text: searchText },
    );
  }
  const matchEnd = matchStart + searchText.length;

  // 找 match.start 落在的 run
  const startSpan = runSpans.find((s) => matchStart >= s.start && matchStart < s.end);
  // match.end 可能恰好等于某个 span 的 end
  const endSpan = runSpans.find((s) => matchEnd > s.start && matchEnd <= s.end);

  if (!startSpan || !endSpan) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `createBookmark TextRange: match spans runs not reachable via direct <w:r> enumeration`,
      ctx.docId,
      { name, text: searchText, opType: 'createBookmark' },
    );
  }

  const sameRun = startSpan.run === endSpan.run;

  let startRunForMark: Element;
  let endRunForMark: Element;

  try {
    if (sameRun) {
      const endOffset = matchEnd - endSpan.start;
      // 先切 end
      splitRunAt(startSpan.run, endOffset);
      // 再切 start(此时 startSpan.run 文本为 [span.start, matchEnd))
      const startOffsetInRun = matchStart - startSpan.start;
      const { after: matchRun } = splitRunAt(startSpan.run, startOffsetInRun);
      startRunForMark = matchRun ?? startSpan.run;
      endRunForMark = startRunForMark;
    } else {
      // 跨 run:切 startSpan(取右半),切 endSpan(取左半)
      const startOffsetInRun = matchStart - startSpan.start;
      const { after: afterStart } = splitRunAt(startSpan.run, startOffsetInRun);
      startRunForMark = afterStart ?? startSpan.run;

      const endOffset = matchEnd - endSpan.start;
      const { before: leftOfEnd } = splitRunAt(endSpan.run, endOffset);
      endRunForMark = leftOfEnd;
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `createBookmark TextRange: run structure too complex to split (${msg})`,
      ctx.docId,
      { name, text: searchText, opType: 'createBookmark' },
    );
  }

  // 插入 bookmarkStart 在 startRunForMark 之前
  const startEl = createBookmarkStart(doc, name, id);
  const endEl = createBookmarkEnd(doc, id);

  paragraph.insertBefore(startEl, startRunForMark);

  // 插入 bookmarkEnd 在 endRunForMark 之后
  const afterEndRun = endRunForMark.nextSibling;
  if (afterEndRun) {
    paragraph.insertBefore(endEl, afterEndRun);
  } else {
    paragraph.appendChild(endEl);
  }

  return { start: startEl, end: endEl };
}

// ─────────── 内部:part 定位 / markDirty ───────────

/** 从 paraIdx 或 Element 的 ownerDocument 反查 partPath */
function findPartPathForElement(ctx: ExecutorContext, el: Element): string {
  // 遍历 pkg.parts,找 xmlDom 包含此 element 的 part
  for (const part of listParts(ctx.pkg)) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    if (
      part.xmlDom.documentElement &&
      part.xmlDom.documentElement.contains(el as unknown as Node)
    ) {
      return part.path;
    }
  }
  throw new WordError('PACKAGE_CORRUPT', 'element not found in any part', ctx.docId);
}

function markPartDirtyByPath(ctx: ExecutorContext, partPath: string, opName: string): void {
  const part = getPart(ctx.pkg, partPath);
  if (!part) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `part '${partPath}' not found in pkg (op: ${opName})`,
      ctx.docId,
      { partPath, opName },
    );
  }
  markDirty(part);
}

// ─────────── 跨段书签创建(供 paragraph-structure-ops 复用,#176)───────────

/** 命名规则前缀的合法字符:首字符 letter/`_`,后续 letter/digit/`_`(无长度上限) */
const BOOKMARK_PREFIX_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

/** 重名时追加后缀的最大尝试次数(防御性上限,正常场景 1~2 次足够)*/
const NAME_COLLISION_MAX_TRIES = 1000;

/**
 * 解析"插入段同时建跨段书签"的规则化命名。
 *
 * 公式:`${prefix}_${anchorParaId}`,与已有书签同名时追加 `_2`/`_3`/... 后缀。
 *
 * 校验:
 * - prefix 匹配 `BOOKMARK_PREFIX_RE`,否则 `BOOKMARK_NAME_INVALID`
 * - 最终名长度 ≤ 40,否则 `BOOKMARK_NAME_INVALID`(与 createBookmark 一致)
 * - 1000 次后仍冲突抛 `BOOKMARK_NAME_EXISTS`(实际不可能触发,防止极端场景死循环)
 */
export function resolveAppendBookmarkName(
  ctx: ExecutorContext,
  prefix: string,
  anchorParaId: string,
): string {
  if (!BOOKMARK_PREFIX_RE.test(prefix)) {
    throw new WordError(
      'BOOKMARK_NAME_INVALID',
      `bookmark prefix '${prefix}' must match /^[A-Za-z_][A-Za-z0-9_]*$/`,
      ctx.docId,
      { prefix },
    );
  }

  const base = `${prefix}_${anchorParaId}`;

  // 第一次尝试:基底名
  let candidate = base;
  let suffix = 2;
  while (ctx.indices.bookmarkIdx.has(candidate)) {
    if (suffix > NAME_COLLISION_MAX_TRIES) {
      throw new WordError(
        'BOOKMARK_NAME_EXISTS',
        `bookmark name resolution gave up after ${NAME_COLLISION_MAX_TRIES} attempts (base='${base}')`,
        ctx.docId,
        { prefix, anchorParaId, base },
      );
    }
    candidate = `${base}_${suffix}`;
    suffix += 1;
  }

  // 长度校验放在解析之后 —— 含后缀的最终名才是要写进 OOXML 的串
  if (candidate.length > 40) {
    throw new WordError(
      'BOOKMARK_NAME_INVALID',
      `resolved bookmark name '${candidate}' exceeds 40-char OOXML limit`,
      ctx.docId,
      { prefix, anchorParaId, resolvedName: candidate, length: candidate.length },
    );
  }

  return candidate;
}

/**
 * 在 anchor 段头与 newPara 段尾之间建立跨段书签(#176)。
 *
 * 调用前置条件:
 * - anchor 与 newPara 在同一 part 内(insertParagraphAfter 把 newPara 作为 anchor
 *   下一兄弟插入,因此天然成立)
 * - trackChanges=false(本路径不支持修订包装,与 createBookmark 一致)
 *
 * 副作用:
 * - 在 anchor 段 pPr 之后插 `<w:bookmarkStart>`
 * - 在 newPara 末尾插 `<w:bookmarkEnd>`
 * - 更新 `ctx.indices.bookmarkIdx`
 * - push 一条 `{ opIndex, name, range }` 到 `ctx.result.createdBookmarks`
 * - markDirty 对应 part(若调用方已 markDirty,重复调用是幂等的)
 */
export function createCrossParagraphBookmark(
  ctx: ExecutorContext,
  name: string,
  anchor: Element,
  newPara: Element,
  fromParaId: string,
  toParaId: string,
): void {
  // trackChanges 守卫(insertParagraphAfter 自身在 trackChanges 下也能跑;此处单独
  // 拒绝跨段书签,与 createBookmark 现状一致;LLM/skill 见到 UNSUPPORTED 错误后
  // 自行降级即可)。
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'insertParagraphAfter with bookmark + trackChanges=true not yet implemented',
      ctx.docId,
      { name, opType: 'insertParagraphAfter' },
    );
  }

  validateBookmarkName(ctx, name);

  // 重名兜底:resolveAppendBookmarkName 已避免,但若调用路径变了仍报错
  if (ctx.indices.bookmarkIdx.has(name)) {
    throw new WordError(
      'BOOKMARK_NAME_EXISTS',
      `bookmark name '${name}' already exists`,
      ctx.docId,
      { name },
    );
  }

  const doc = anchor.ownerDocument;
  if (!doc) {
    throw new WordError('PACKAGE_CORRUPT', 'anchor paragraph has no ownerDocument', ctx.docId);
  }

  const id = allocateBookmarkId(ctx);

  // bookmarkStart 放在 anchor 段 pPr 之后、第一个内容节点之前(若全空就 append)
  const startEl = createBookmarkStart(doc, name, id);
  const insertBefore = findContentInsertPoint(anchor);
  if (insertBefore) {
    anchor.insertBefore(startEl, insertBefore);
  } else {
    anchor.appendChild(startEl);
  }

  // bookmarkEnd 放在 newPara 末尾
  const endEl = createBookmarkEnd(doc, id);
  newPara.appendChild(endEl);

  // 索引同步
  const partPath = findPartPathForElement(ctx, anchor);
  ctx.indices.bookmarkIdx.set(name, { start: startEl, end: endEl, partPath });

  // ApplyResult
  ctx.result.createdBookmarks.push({
    opIndex: ctx.opIndex,
    name,
    range: { fromParaId, toParaId },
  });

  // anchor 与 newPara 同 part,structure-op 的 markDirty 已经覆盖,这里 markDirty
  // 是幂等的(markDirty 重复调用安全)。显式调用一次,避免依赖调用顺序。
  markPartDirtyByPath(ctx, partPath, 'insertParagraphAfter:bookmark');
}
