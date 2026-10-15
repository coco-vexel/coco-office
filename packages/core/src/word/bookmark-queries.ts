/**
 * Bookmark 只读查询(docs/03.1 §1.5 / §2,#25)。
 *
 * 基于 `PackageIndices.bookmarkIdx`(#13 建好)做 O(1) 反查,没有任何 fontoxpath
 * 调用。rangeKind 判定走 DOM 遍历(孩子顺序 + 同 `<w:p>` 判断)。
 *
 * 三类 rangeKind(docs/03.1 §1.5):
 *   - `'paragraph'` —— start / end 在**同一 `<w:p>`** 中,且 `<w:bookmarkStart>`
 *     之前没有任何内容节点、`<w:bookmarkEnd>` 之后也没有 —— 即书签覆盖整段的"内容区"
 *   - `'text'` —— 同一 `<w:p>`,但 start / end 不在段首/段尾(书签仅覆盖段内文本片段)
 *   - `'complex'` —— 跨段落、缺 end、start / end parent 不同、或 parent 不是 `<w:p>`
 *     (表 cell、body 直子等),core 不简单建模,只标 kind
 *
 * `hidden`:name 以 `_` 开头(如 `_Toc123` / `_GoBack` / `_Hlk12345`)。`listBookmarks`
 * 默认过滤;`getBookmark(name)` 不过滤(调用方已知 name)。
 */
import type { Package } from '../pack/index.js';
import type { BookmarkEntry, PackageIndices } from './indices.js';
import { classifyRange } from './range-kind.js';
import { findStoryFor, type StoryId } from './stories.js';

export type BookmarkName = string;

export interface BookmarkInfo {
  name: BookmarkName;
  story: StoryId;
  rangeKind: 'paragraph' | 'text' | 'complex';
  /** rangeKind = 'paragraph' | 'text' 时给(书签所属的段落 paraId,大写 8 hex) */
  paraId?: string;
  /** rangeKind = 'text' 时给(start / end 之间所有 `<w:t>` 的纯文本拼接) */
  text?: string;
  /** name 以 `_` 开头时 true;否则缺字段 */
  hidden?: boolean;
}

export interface ListBookmarksOptions {
  /** 默认 false。true 时返回所有书签,包括 `_` 开头的隐藏书签 */
  includeHidden?: boolean;
  /**
   * 名字前缀过滤(大小写敏感)。提供时只返回 `name.startsWith(prefix)` 的书签。
   *
   * 配合 `InsertParagraphSpec.bookmark.prefix` 的规则化命名(`{prefix}_{anchorParaId}` /
   * 冲突时 `{prefix}_{anchorParaId}_2` 等)使用,可批量回查"某语义标签下的所有书签"。
   * `_` 开头的隐藏书签若未设置 `includeHidden=true` 仍被过滤(prefix 与 hidden 是
   * 独立轴)。
   */
  prefix?: string;
}

// ─────────── 公共入口 ───────────

export function listBookmarks(
  pkg: Package,
  indices: PackageIndices,
  opts: ListBookmarksOptions = {},
): BookmarkInfo[] {
  const includeHidden = opts.includeHidden === true;
  const prefix = opts.prefix;
  const out: BookmarkInfo[] = [];
  for (const [name, entry] of indices.bookmarkIdx) {
    if (prefix !== undefined && !name.startsWith(prefix)) continue;
    const hidden = isHiddenName(name);
    if (hidden && !includeHidden) continue;
    const info = buildBookmarkInfo(pkg, name, entry);
    if (info) out.push(info);
  }
  return out;
}

export function getBookmark(
  pkg: Package,
  indices: PackageIndices,
  name: BookmarkName,
): BookmarkInfo | null {
  const entry = indices.bookmarkIdx.get(name);
  if (!entry) return null;
  return buildBookmarkInfo(pkg, name, entry);
}

// ─────────── 内部 ───────────

function isHiddenName(name: string): boolean {
  return name.startsWith('_');
}

function buildBookmarkInfo(pkg: Package, name: string, entry: BookmarkEntry): BookmarkInfo | null {
  // story 由 start 元素反查(bookmark 横跨 part 的情况极少见;若 start 不在任一 story
  // 子树里,入索引时本应不会发生 —— 此处保险返回 null)。
  const story = findStoryFor(pkg, entry.start);
  if (!story) return null;

  const { kind, paraId, text } = classifyRange(entry.start, entry.end);

  const info: BookmarkInfo = {
    name,
    story: story.id,
    rangeKind: kind,
  };
  if (paraId !== undefined) info.paraId = paraId;
  if (text !== undefined) info.text = text;
  if (isHiddenName(name)) info.hidden = true;
  return info;
}
