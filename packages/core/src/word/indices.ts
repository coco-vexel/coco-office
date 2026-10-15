/**
 * Word 装载期三索引(docs/03 §13.1 / §13.6)。
 *
 * - paraIdx:paraId(大写 8 hex)→ <w:p> Element
 * - bookmarkIdx:name → { start: <w:bookmarkStart>; end: <w:bookmarkEnd>; partPath }
 * - commentIdx:id(number)→ { comment?: <w:comment>; anchors: { start?, end?, ref? } }
 *
 * 同时完成装载期书签校验:**空名抛错**;**重名容错**(保留第一个、忽略后续,绝不崩 ——
 * 真实文档常见重复的 `_GoBack` 等编辑痕迹书签)。未来由 #15 WordError 统一,先用 Error + code。
 *
 * 性能目标:10 万段落 < 100ms(Chrome/Node;WPS WebView 宽松到 < 300ms)。由 `ensureParaIds`
 * 同样走 getElementsByTagNameNS 保证,这里复用同一枚举路径。
 */
import type { Package, Part } from '../pack/index.js';
import { isXmlPart } from '../pack/index.js';
import { W_NS, W14_NS } from './namespaces.js';
import { buildStylesIndex, type StyleInfo } from './style-queries.js';
import { buildTableIndex } from './list-tables.js';
import { buildDrawingIndex } from './list-drawings.js';

/** 书签校验错误 code(暂不挂 WordError;#15 会统一到 WordErrorCode) */
export type BookmarkLoadErrorCode =
  // 【已弃用 / 不再抛出】重名书签现为容错(保留第一个、忽略后续),buildIndices 不再产生此 code。
  // 保留在 union 中仅作向后兼容,避免破坏可能引用此类型的消费方。
  'BOOKMARK_NAME_DUPLICATE_ON_LOAD' | 'BOOKMARK_NAME_EMPTY_ON_LOAD';

/** 装载期抛出的书签错误 */
export class BookmarkLoadError extends Error {
  readonly code: BookmarkLoadErrorCode;
  readonly details: Record<string, unknown>;
  constructor(code: BookmarkLoadErrorCode, message: string, details: Record<string, unknown> = {}) {
    super(message);
    this.name = 'BookmarkLoadError';
    this.code = code;
    this.details = details;
  }
}

export interface BookmarkEntry {
  /** `<w:bookmarkStart>` 元素 */
  start: Element;
  /** `<w:bookmarkEnd>` 元素;若装载数据里缺失可能为 null */
  end: Element | null;
  /** 所在 part 的 path(如 'word/document.xml') */
  partPath: string;
}

export interface CommentEntry {
  /** `<w:comment>` 元素(来自 comments.xml);若无 comments.xml 可能为 null */
  comment: Element | null;
  /** 锚点三件套(在 document.xml / headers / footers 等 body-side part 上) */
  anchors: {
    start: Element | null;
    end: Element | null;
    ref: Element | null;
  };
}

export interface PackageIndices {
  /** paraId → <w:p> */
  paraIdx: Map<string, Element>;
  /** bookmark name → entry(含 start/end/partPath) */
  bookmarkIdx: Map<string, BookmarkEntry>;
  /** comment id → entry(comment 主体 + 三锚点) */
  commentIdx: Map<number, CommentEntry>;
  /**
   * styleId → StyleInfo（来自 styles.xml）。
   * styles.xml 不存在时为空 Map。
   * 消费方按 styleId 做 O(1) 查找，沿 basedOn 链判断语义角色时须自行防环。
   */
  stylesIdx: Map<string, StyleInfo>;
  /**
   * tableId（'tbl-NNNN'）→ <w:tbl> Element（#185）。
   * 装载期按文档顺序分配，包含所有 part 的所有表格（含嵌套）。
   * 供 listTables / getTableById / plan executor（deleteTable）使用。
   */
  tblIdx: Map<string, Element>;
  /**
   * drawingId（'dr-NNNN'）→ <w:drawing> Element（#189）。
   * 装载期按文档顺序分配，包含所有 Word 内容 part 的所有 drawing。
   * 供 listDrawings / getDrawingById / 占位图替换（#190）使用。
   *
   * 注意：嵌套在 group / canvas 内的子 drawing 也会被收入（保持枚举层完整暴露），
   * 上层（#191 / #192）按需用 hasDrawingAncestor 自行决定合并策略。
   */
  drawingIdx: Map<string, Element>;
}

// re-export StyleInfo 供其他 word 内部模块从 indices 一站式引入
export type { StyleInfo };

/** 判断 part 是否 Word 主要内容 part(按常见 contentType 前缀匹配) */
function isWordContentPart(part: Part): boolean {
  const ct = part.contentType;
  return (
    ct.includes('wordprocessingml') ||
    ct.endsWith('.main+xml') ||
    ct.endsWith('.header+xml') ||
    ct.endsWith('.footer+xml') ||
    ct.endsWith('.footnotes+xml') ||
    ct.endsWith('.endnotes+xml') ||
    ct.endsWith('.comments+xml')
  );
}

/** 判断 part 是否 comments 存储 part(commentIdx 的 comment 主体来自这里) */
function isCommentsPart(part: Part): boolean {
  return part.contentType.endsWith('.comments+xml');
}

/** 把 HTMLCollection-ish 枚举结果快照成数组(避免 live 行为) */
function snapshot<T extends Element>(col: { length: number } & Record<number, T>): T[] {
  const out: T[] = [];
  for (let i = 0; i < col.length; i++) {
    const el = col[i];
    if (el) out.push(el);
  }
  return out;
}

/**
 * 扫 Package 构建三索引。
 *
 * - **装载期书签校验**:遇空名抛 `BookmarkLoadError`(code = BOOKMARK_NAME_EMPTY_ON_LOAD);
 *   遇重名**容错**(保留第一个、忽略后续,不抛错 —— `_GoBack` 等编辑痕迹书签可安全忽略)
 * - 孤儿 `bookmarkEnd`(无匹配 start)忽略
 * - 缺 `bookmarkEnd` 的书签仍入索引,`entry.end` 为 null
 * - 不抛错的其它不一致(comment 无 <w:comment> 主体、anchors 不全)进索引,由调用方决定是否告警
 */
export function buildIndices(pkg: Package): PackageIndices {
  const paraIdx = new Map<string, Element>();
  const bookmarkIdx = new Map<string, BookmarkEntry>();
  const commentIdx = new Map<number, CommentEntry>();

  // bookmarkStart 临时:id -> { start, partPath, name }(用于和后续 bookmarkEnd 配对)
  const pendingStarts = new Map<string, { start: Element; partPath: string; name: string }>();

  for (const part of pkg.parts.values()) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    if (!isWordContentPart(part)) continue;

    // paraIdx
    const paragraphs = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'p'));
    for (const p of paragraphs) {
      const id = p.getAttributeNS(W14_NS, 'paraId');
      if (id) paraIdx.set(id.toUpperCase(), p);
    }

    // bookmarkIdx — 先收集 start,后续匹配 end
    const starts = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'bookmarkStart'));
    for (const start of starts) {
      const bId = start.getAttributeNS(W_NS, 'id');
      const bName = start.getAttributeNS(W_NS, 'name');
      if (bName === null) {
        // 缺 name 属性 —— 当作空名
        throw new BookmarkLoadError(
          'BOOKMARK_NAME_EMPTY_ON_LOAD',
          `bookmarkStart with w:id="${bId ?? ''}" has no w:name`,
          { partPath: part.path, bookmarkId: bId ?? null },
        );
      }
      if (bName.length === 0) {
        throw new BookmarkLoadError(
          'BOOKMARK_NAME_EMPTY_ON_LOAD',
          `bookmarkStart with w:id="${bId ?? ''}" has empty w:name`,
          { partPath: part.path, bookmarkId: bId ?? null },
        );
      }
      if (bookmarkIdx.has(bName)) {
        // 容错：重名书签**保留第一个、忽略后续**，不抛错（绝不因装载期数据瑕疵崩溃）。
        //
        // 真实文档（多次编辑 / 多工具处理过）经常出现重复书签名 —— 最典型的是 Word 自动
        // 插入的 `_GoBack`（标记上次编辑位置），也可能是模板拼接残留的其它重名书签。
        // OOXML 实践中 Word 自己也容忍重复书签名；bookmarkIdx 是 name→entry 的 Map，按 name
        // 唯一本就是其数据结构约束，保留首个即可满足所有下游（反查 / 重名检查 / 删除）。
        //
        // 注意：这里只跳过往 bookmarkIdx 登记，但**不会**把后续重名的 bookmarkStart 当作
        // pendingStart 处理，因此其配套的 bookmarkEnd 不会错配到第一个 entry 上（end 仍只认
        // 第一个 start 的 id 作用域）。元素本身保留在 DOM 里，保存时原样写回，内容不丢失。
        continue;
      }
      // 先登记:end 稍后补
      bookmarkIdx.set(bName, { start, end: null, partPath: part.path });
      if (bId !== null) {
        pendingStarts.set(`${part.path}::${bId}`, { start, partPath: part.path, name: bName });
      }
    }

    // 匹配 bookmarkEnd(按 w:id,作用域限定到 part 内)
    const ends = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'bookmarkEnd'));
    for (const end of ends) {
      const eId = end.getAttributeNS(W_NS, 'id');
      if (eId === null) continue;
      const matched = pendingStarts.get(`${part.path}::${eId}`);
      if (matched) {
        const entry = bookmarkIdx.get(matched.name);
        if (entry) entry.end = end;
      }
    }

    // commentIdx — 锚点(document.xml / headers / footers 等 body-side part)
    const anchorStarts = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'commentRangeStart'));
    const anchorEnds = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'commentRangeEnd'));
    const anchorRefs = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'commentReference'));
    for (const anchor of anchorStarts) {
      const id = numAttr(anchor, W_NS, 'id');
      if (id === null) continue;
      getOrCreateCommentEntry(commentIdx, id).anchors.start = anchor;
    }
    for (const anchor of anchorEnds) {
      const id = numAttr(anchor, W_NS, 'id');
      if (id === null) continue;
      getOrCreateCommentEntry(commentIdx, id).anchors.end = anchor;
    }
    for (const anchor of anchorRefs) {
      const id = numAttr(anchor, W_NS, 'id');
      if (id === null) continue;
      getOrCreateCommentEntry(commentIdx, id).anchors.ref = anchor;
    }

    // commentIdx — 主体(comments.xml)
    if (isCommentsPart(part)) {
      const comments = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'comment'));
      for (const c of comments) {
        const id = numAttr(c, W_NS, 'id');
        if (id === null) continue;
        getOrCreateCommentEntry(commentIdx, id).comment = c;
      }
    }
  }

  const stylesIdx = buildStylesIndex(pkg);
  const tblIdx = buildTableIndex(pkg.parts.values());
  const drawingIdx = buildDrawingIndex(pkg.parts.values());

  return { paraIdx, bookmarkIdx, commentIdx, stylesIdx, tblIdx, drawingIdx };
}

function numAttr(el: Element, ns: string, local: string): number | null {
  const v = el.getAttributeNS(ns, local);
  if (v === null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function getOrCreateCommentEntry(idx: Map<number, CommentEntry>, id: number): CommentEntry {
  let e = idx.get(id);
  if (!e) {
    e = { comment: null, anchors: { start: null, end: null, ref: null } };
    idx.set(id, e);
  }
  return e;
}
