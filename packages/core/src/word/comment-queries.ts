/**
 * Comment 只读查询(docs/03.1 §1.5 / §2,#26)。
 *
 * 批注由两部分组成:
 *   - **主体**(`<w:comment>` in comments.xml):id / author / date / 批注正文
 *   - **锚点**(`<w:commentRangeStart>` / `<w:commentRangeEnd>` / `<w:commentReference>`
 *     in body / headers / footers 等):标识被批注的范围
 *
 * `PackageIndices.commentIdx` 已在装载期把两边按 `w:id`(数字)串联。这里做:
 *   - 读主体拿 author / date / text
 *   - 用锚点 start / end 判 rangeKind(复用 [range-kind.ts](./range-kind.ts) 的
 *     `classifyRange`)
 *   - 由 start 元素反查 story(`findStoryFor`)
 *
 * 跟 bookmark 的字段差异:
 *   - BookmarkInfo 的 `text` 字段 → CommentInfo 的 `anchorText` 字段(语义更明确)
 *   - 没有 `hidden` 概念
 */
import type { CommentEntry, PackageIndices } from './indices.js';
import { W_NS } from './namespaces.js';
import { classifyRange } from './range-kind.js';

export type CommentId = number;

export interface CommentInfo {
  id: CommentId;
  /** w:author,缺则 '' */
  author: string;
  /** w:date(ISO 8601);缺则 '' */
  date: string;
  /** 批注正文:`<w:comment>` 内所有 `<w:t>` 拼接 */
  text: string;
  rangeKind: 'paragraph' | 'text' | 'complex';
  /** rangeKind = 'paragraph' | 'text' 时给(批注锚定的段落 paraId) */
  paraId?: string;
  /** rangeKind = 'text' 时给(start/end 之间的正文片段) */
  anchorText?: string;
}

// ─────────── 公共入口 ───────────

/**
 * 枚举所有批注。
 *
 * 顺序按 `commentIdx` 的迭代序(Map 插入序;装载期按 part 遍历 + w:id 升序扫到)。
 *
 * 主体缺失(`entry.comment === null`)或锚点无 start/end 的批注 **被跳过** ——
 * 调用方看到的永远是"主体 + 至少一个锚点"都有的条目。
 */
export function listComments(indices: PackageIndices): CommentInfo[] {
  const out: CommentInfo[] = [];
  for (const [id, entry] of indices.commentIdx) {
    const info = buildCommentInfo(id, entry);
    if (info) out.push(info);
  }
  return out;
}

/** 按 id 取单条批注。不存在或主体缺失返回 null */
export function getComment(indices: PackageIndices, id: CommentId): CommentInfo | null {
  const entry = indices.commentIdx.get(id);
  if (!entry) return null;
  return buildCommentInfo(id, entry);
}

// ─────────── 内部 ───────────

function buildCommentInfo(id: CommentId, entry: CommentEntry): CommentInfo | null {
  const comment = entry.comment;
  // 主体缺失:批注数据异常(只有锚点没有 <w:comment>)—— 此时 author/date/text 都
  // 无法给,返回 null 比"空壳占位"更诚实。装载期对此情形不抛错(buildIndices 允许
  // entry.comment 为 null),只有查询时才显式忽略。
  if (!comment) return null;

  const author = comment.getAttributeNS(W_NS, 'author') ?? '';
  const date = comment.getAttributeNS(W_NS, 'date') ?? '';
  const text = collectDescendantTText(comment);

  // rangeKind:基于锚点 start / end。两边都缺则退化 'complex'。
  // `classifyRange.text` 在 comment 语境下含义 = anchorText。
  const {
    kind,
    paraId,
    text: anchorText,
  } = classifyRange(entry.anchors.start ?? null, entry.anchors.end ?? null);

  const info: CommentInfo = {
    id,
    author,
    date,
    text,
    rangeKind: kind,
  };
  if (paraId !== undefined) info.paraId = paraId;
  if (anchorText !== undefined) info.anchorText = anchorText;
  return info;
}

/**
 * 拉 `<w:comment>` 内全部 `<w:t>` 的纯文本,按文档序拼接。
 *
 * `<w:comment>` 可能包含多个 `<w:p>` → `<w:r>` → `<w:t>` 层级,也可能混有
 * `<w:hyperlink>` 等包装 —— 一次性 `getElementsByTagNameNS(W_NS, 't')` 吃全部后代。
 * 段落分隔不在文本里(不加 `\n`),与 Bookmark 的 text 字段口径一致。
 */
function collectDescendantTText(el: Element): string {
  const ts = el.getElementsByTagNameNS(W_NS, 't');
  let out = '';
  for (let i = 0; i < ts.length; i++) {
    out += ts[i]!.textContent ?? '';
  }
  return out;
}
