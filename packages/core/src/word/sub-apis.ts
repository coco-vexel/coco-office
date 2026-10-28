/**
 * Core-Word 读 API 的 namespace 子对象(#92,docs/03.1 §2)。
 *
 * 把 Phase B 原本平铺在 `WordDocument` 上的 7 个读方法按实体聚合为 4 个子 API:
 *
 *   | 旧(平铺) | 新(namespace) |
 *   |---|---|
 *   | `doc.listParagraphs(opts)` | `doc.paragraphs.list(opts)` |
 *   | `doc.getParagraph(id)` | `doc.paragraphs.get(id)` |
 *   | `doc.listBookmarks(opts)` | `doc.bookmarks.list(opts)` |
 *   | `doc.getBookmark(name)` | `doc.bookmarks.get(name)` |
 *   | `doc.listComments()` | `doc.comments.list()` |
 *   | `doc.getComment(id)` | `doc.comments.get(id)` |
 *   | `doc.listRevisions(opts)` | `doc.revisions.list(opts)` |
 *
 * 动机:Phase F 起属性层(`format.*`)+ sections / tables / drawings / styles /
 * hyperlinks / fields 若继续平铺会让主类表面爆炸。namespace 聚合后新实体扩展只
 * 需往 `WordDocument` 加一个 `readonly` 字段(见 docs/03.1 §2 的 "未来扩展")。
 *
 * 命名规范(docs/03.1 §2.1):
 * - 子 API 方法用动词短命名 `list` / `get`,不重复实体名
 * - 过滤/分页选项保留顶层类型名(`ListParagraphsOptions` 不简化)
 * - 返回值类型名不变
 *
 * RPC method name 同步迁移:`paragraphs.list` / `paragraphs.get` / `bookmarks.list`
 * / `bookmarks.get` / `comments.list` / `comments.get` / `revisions.list`。协议
 * 版本不 bump —— 这是 method name rename,非协议结构变更。
 */
import type { BookmarkInfo, BookmarkName, ListBookmarksOptions } from './bookmark-queries.js';
import type { CommentId, CommentInfo } from './comment-queries.js';
import type { ContextSearchOptions, ContextSearchResult } from './context-search.js';
import type { ListParagraphsOptions, ListResult, ParagraphInfo } from './list-paragraphs.js';
import type {
  ListTablesOptions,
  RowHeight,
  TableCellBlockInfo,
  TableCellInfo,
  TableInfo,
  TableVerbatim,
} from './list-tables.js';
import type { DrawingInfo, ListDrawingsOptions, TextBoxBlockInfo } from './list-drawings.js';
import type {
  ApplyResult,
  Op,
  ParaId,
  ParagraphDirectFormat,
  Range,
  TableCellDirectFormat,
} from './plan.js';
import type { ListRevisionsOptions, RevisionInfo } from './revision-queries.js';
import type { SectionInfo } from './section-queries.js';
import type { TocFieldRange } from './toc-queries.js';
import type { ListStylesOptions, StyleInfo } from './style-queries.js';
import type { Transport } from './transport.js';
import { W_NS } from './namespaces.js';

// ─────────── 接口(docs/03.1 §2)───────────

export interface ParagraphApi {
  /**
   * Lists paragraphs and returns live paragraph handles.
   */
  list(opts?: ListParagraphsOptions): Promise<ParagraphList>;
  /**
   * Gets one paragraph by `paraId`.
   *
   * Returns `null` when the paragraph no longer exists.
   */
  get(paraId: string): Promise<Paragraph | null>;
}

export interface ParagraphList {
  paragraphs: Paragraph[];
  pagination: ListResult['pagination'];
}

export interface LiveObjectWriteOptions {
  /**
   * Enables tracked-change markup for the write operation when supported.
   */
  trackChanges?: boolean;
  /**
   * Author metadata used by comments and tracked changes.
   */
  author?: string;
  /**
   * Date metadata used by comments and tracked changes.
   */
  date?: Date;
}

export interface ContextApi {
  /**
   * Searches local document content and returns actionable context matches.
   */
  search(opts: ContextSearchOptions): Promise<ContextSearchResult>;
}

export interface RevisionBulkOptions extends LiveObjectWriteOptions {
  /**
   * Bulk revision writes currently target all supported revisions.
   */
  scope?: 'all';
}

export interface ParagraphDeleteOptions extends LiveObjectWriteOptions {
  /**
   * Controls what happens when deleting the paragraph would affect anchors such
   * as bookmarks or comments.
   */
  onAnchorConflict?: 'error' | 'drop';
}

export interface ParagraphInsertOptions extends LiveObjectWriteOptions {
  /**
   * Controls the new paragraph's paragraph/run style source.
   */
  style?: 'inherit' | string;
  /**
   * Controls whether the new paragraph run inherits the anchor run format.
   */
  runFormat?: 'inherit' | 'default';
}

export interface ParagraphReplaceOptions extends LiveObjectWriteOptions {
  /**
   * Controls what happens when replacing a paragraph would affect anchors such
   * as bookmarks or comments.
   */
  onAnchorConflict?: 'error' | 'drop';
  /**
   * Preserves inline drawings/objects while replacing text runs.
   */
  keepInlineGraphics?: boolean;
}

export interface ParagraphInsertImageSpec {
  /** Image bytes. The SDK accepts browser-safe Uint8Array data. */
  bytes: Uint8Array;
  /** Image MIME type, for example `image/png` or `image/jpeg`. */
  contentType: string;
  /** Rendered width in EMU. 1 inch = 914400 EMU. */
  widthEmu: number;
  /** Rendered height in EMU. 1 inch = 914400 EMU. */
  heightEmu: number;
}

export interface ParagraphInsertTableSpec {
  /**
   * Text matrix used to create a simple Word table. Missing cells in ragged
   * rows are created as empty cells.
   */
  rows: string[][];
  /** Column widths in twips. Defaults to 2400 twips for each column. */
  columnWidths?: number[];
  /** Optional Word table style id, written as `w:tblStyle`. */
  tableStyleId?: string;
}

/**
 * Live handle for one document paragraph.
 *
 * Snapshot fields such as `paraId` and `story` come from the read that created
 * the handle. Methods read or write the current document state.
 */
export interface Paragraph extends ParagraphInfo {
  /** Stable paragraph identifier. */
  readonly paraId: ParaId;
  /** Alias of `paraId`, kept for object-style APIs. */
  readonly id: ParaId;
  /** Document story that contains this paragraph. */
  readonly story: ParagraphInfo['story'];
  /**
   * Reads the current metadata for this paragraph.
   */
  getInfo(
    opts?: Pick<ListParagraphsOptions, 'includeFormat' | 'textFormat' | 'includeTableInfo'>,
  ): Promise<ParagraphInfo | null>;
  /**
   * Reads the current plain text of this paragraph.
   */
  getText(): Promise<string | null>;
  /**
   * Replaces this paragraph with plain text.
   */
  replaceWith(text: string, opts?: ParagraphReplaceOptions): Promise<ApplyResult>;
  /**
   * Sets horizontal paragraph alignment.
   */
  setAlignment(
    value: NonNullable<ParagraphDirectFormat['alignment']>,
    opts?: LiveObjectWriteOptions,
  ): Promise<ApplyResult>;
  /**
   * Sets paragraph outline level. Values 0-8 correspond to heading levels 1-9.
   */
  setOutlineLevel(value: number, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets whether this paragraph starts on a new page.
   */
  setPageBreakBefore(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets paragraph indentation in twips.
   */
  setIndent(
    value: NonNullable<ParagraphDirectFormat['indent']>,
    opts?: LiveObjectWriteOptions,
  ): Promise<ApplyResult>;
  /**
   * Sets paragraph spacing in twips.
   */
  setSpacing(
    value: NonNullable<ParagraphDirectFormat['spacing']>,
    opts?: LiveObjectWriteOptions,
  ): Promise<ApplyResult>;
  /**
   * Inserts a new paragraph before this paragraph.
   */
  insertBefore(text: string, opts?: ParagraphInsertOptions): Promise<ApplyResult>;
  /**
   * Inserts a new paragraph after this paragraph.
   */
  insertAfter(text: string, opts?: ParagraphInsertOptions): Promise<ApplyResult>;
  /**
   * Inserts a new inline image paragraph after this paragraph.
   */
  insertImageAfter(
    spec: ParagraphInsertImageSpec,
    opts?: LiveObjectWriteOptions,
  ): Promise<ApplyResult>;
  /**
   * Inserts a simple text table after this paragraph.
   *
   * For high-fidelity table copy, use `table.getVerbatim()` with
   * `doc.plan().insertTableFrom(...)`.
   */
  insertTableAfter(
    spec: ParagraphInsertTableSpec,
    opts?: LiveObjectWriteOptions,
  ): Promise<ApplyResult>;
  /**
   * Deletes this paragraph from the document.
   */
  delete(opts?: ParagraphDeleteOptions): Promise<ApplyResult>;
  /**
   * Adds a comment to this whole paragraph.
   */
  addComment(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Creates a bookmark around this whole paragraph.
   */
  createBookmark(name: BookmarkName, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Creates a text range handle by searching inside this paragraph.
   *
   * The returned range is lazy: it resolves the search when a range method is
   * executed.
   */
  findText(search: string, opts?: { scope?: 'first' | 'all' }): TextRange;
}

/**
 * Live handle for text matched inside one paragraph.
 */
export interface TextRange {
  /**
   * Replaces the matched text with plain text.
   */
  replaceWith(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Deletes the matched text.
   */
  delete(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets bold formatting on the matched text.
   */
  setBold(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets font color on the matched text.
   */
  setColor(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets italic formatting on the matched text.
   */
  setItalic(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets single underline on the matched text.
   */
  setUnderline(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets Word highlight color on the matched text.
   */
  setHighlight(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets font size in points on the matched text.
   */
  setFontSize(value: number, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets font family on the matched text.
   */
  setFontFamily(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Adds a comment to the matched text.
   */
  addComment(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Creates a bookmark around the matched text.
   */
  createBookmark(name: BookmarkName, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

export interface BookmarkApi {
  /**
   * Lists bookmarks and returns live bookmark handles.
   */
  list(opts?: ListBookmarksOptions): Promise<Bookmark[]>;
  /**
   * Gets one bookmark by name.
   *
   * Returns `null` when the bookmark no longer exists.
   */
  get(name: BookmarkName): Promise<Bookmark | null>;
}

/**
 * Live handle for one document bookmark.
 */
export interface Bookmark extends BookmarkInfo {
  /** Alias of `name`, kept for object-style APIs. */
  readonly id: BookmarkName;
  /**
   * Reads current metadata for this bookmark.
   */
  getInfo(): Promise<BookmarkInfo | null>;
  /**
   * Reads the current text covered by this bookmark when it can be expressed as
   * a paragraph or text range.
   */
  getText(): Promise<string | null>;
  /**
   * Returns the paragraph that contains this bookmark when available.
   */
  getParagraph(): Promise<Paragraph | null>;
  /**
   * Deletes this bookmark anchor pair.
   */
  delete(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

export interface CommentApi {
  /**
   * Lists comments and returns live comment handles.
   */
  list(): Promise<Comment[]>;
  /**
   * Gets one comment by id.
   *
   * Returns `null` when the comment no longer exists.
   */
  get(id: CommentId): Promise<Comment | null>;
}

/**
 * Live handle for one document comment.
 */
export interface Comment extends CommentInfo {
  /**
   * Reads current metadata for this comment.
   */
  getInfo(): Promise<CommentInfo | null>;
  /**
   * Reads the current comment body text from `comments.xml`.
   */
  getText(): Promise<string | null>;
  /**
   * Reads the current document text covered by this comment anchor.
   */
  getAnchorText(): Promise<string | null>;
  /**
   * Returns the paragraph that contains this comment anchor when available.
   */
  getParagraph(): Promise<Paragraph | null>;
  /**
   * Deletes this comment body and its anchors.
   */
  delete(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

export interface RevisionApi {
  /**
   * Lists revisions and returns live revision handles.
   */
  list(opts?: ListRevisionsOptions): Promise<Revision[]>;
  /**
   * Gets one revision by `revId`.
   *
   * Returns `null` when the revision no longer exists.
   */
  get(revId: RevisionInfo['revId'], opts?: ListRevisionsOptions): Promise<Revision | null>;
  /**
   * Accepts all supported revisions.
   */
  accept(opts?: RevisionBulkOptions): Promise<ApplyResult>;
  /**
   * Rejects all supported revisions.
   */
  reject(opts?: RevisionBulkOptions): Promise<ApplyResult>;
}

/**
 * Live handle for one document revision.
 *
 * `revision.accept/reject` operate on this single revision id. Use
 * `doc.revisions.accept/reject` for all supported revisions.
 */
export interface Revision extends RevisionInfo {
  /** Stable alias of `revId`, kept for object-style APIs. */
  readonly id: RevisionInfo['revId'];
  /**
   * Reads current metadata for this revision.
   */
  getInfo(opts?: ListRevisionsOptions): Promise<RevisionInfo | null>;
  /**
   * Reads the revision text when available.
   */
  getText(): Promise<string | null>;
  /**
   * Returns the paragraph that contains this revision when available.
   */
  getParagraph(): Promise<Paragraph | null>;
  /**
   * Accepts this revision.
   */
  accept(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Rejects this revision.
   */
  reject(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

export interface StylesApi {
  /**
   * 列出 styles.xml 中的所有样式（或按 type 过滤）。
   * styles.xml 不存在时返回空数组。
   */
  list(opts?: ListStylesOptions): Promise<StyleInfo[]>;
  /**
   * 按 styleId 精确查找，O(1)。
   * styleId 大小写敏感。不存在时返回 null。
   */
  get(styleId: string): Promise<StyleInfo | null>;
}

export interface TablesApi {
  /**
   * 枚举文档表格（#185）。
   * - scope: 'body'（默认）只返回主文档 body 中的表格
   * - includeNested: false（默认）跳过嵌套在其他表格内的表格
   */
  /**
   * Lists tables in the document and returns live table handles.
   *
   * The returned `Table` values contain a snapshot of the table metadata at
   * list time, and also expose methods that can read fresh state or write back
   * to the same table.
   */
  list(opts?: ListTablesOptions): Promise<Table[]>;
  /**
   * 按 tableId 精确取单个表格，O(1)。不存在返回 null。
   */
  /**
   * Gets one table by `tableId`.
   *
   * Returns `null` when the table no longer exists.
   */
  get(tableId: string): Promise<Table | null>;
  /**
   * 按 tableId verbatim 取一张表的 OOXML 子树 + 格内图字节（自包含 payload）。
   * 供跨文档搬表 `plan.insertTableFrom(...)`（"数据表 = 内容块" §13.9）。不存在返回 null。
   */
  /**
   * Exports the raw OOXML subtree and related media for a table.
   *
   * This is an advanced API for high-fidelity table copying. Normal table
   * reads and writes should use `list`, `get`, and the methods on `Table`.
   */
  getVerbatim(tableId: string): Promise<TableVerbatim | null>;
}

export interface TableDeleteOptions extends LiveObjectWriteOptions {
  /**
   * Controls what happens when deleting the table would affect anchors such as
   * bookmarks or comments.
   */
  onAnchorConflict?: 'error' | 'drop';
}

export interface TableCellWriteOptions extends LiveObjectWriteOptions {
  /**
   * Controls what happens when replacing cell text would affect anchors such
   * as bookmarks or comments.
   */
  onAnchorConflict?: 'error' | 'drop';
}

export interface TableRowDeleteOptions extends LiveObjectWriteOptions {
  /**
   * Controls what happens when deleting the row would affect anchors such as
   * bookmarks or comments.
   */
  onAnchorConflict?: 'error' | 'drop';
}

/**
 * Live handle for one document table.
 *
 * A `Table` is both a metadata snapshot and an operation handle:
 * snapshot fields such as `tableId`, `rowCount`, `colCount`, and `cells` come
 * from the read that created the handle, while methods such as `getInfo` and
 * `replaceCellText` read or write the current document state.
 */
export interface Table extends TableInfo {
  /** Stable alias of `tableId`, kept for object-style APIs. */
  readonly id: string;
  /**
   * Creates a live handle for one cell, addressed by zero-based row and column.
   */
  cell(row: number, col: number): TableCell;
  /**
   * Creates a live handle for one row, addressed by zero-based row index.
   */
  row(row: number): TableRow;
  /**
   * Creates a live handle for one column, addressed by zero-based column index.
   */
  column(col: number): TableColumn;
  /**
   * Reads the current metadata for this table.
   *
   * Use this after writes when you need fresh row, column, width, height, or
   * cell information.
   */
  getInfo(opts?: ListTablesOptions): Promise<TableInfo | null>;
  /**
   * Exports this table as raw OOXML plus related media.
   *
   * This is intended for high-fidelity copy/move flows, not routine editing.
   */
  getVerbatim(): Promise<TableVerbatim | null>;
  /**
   * Reads the current plain text of this table.
   *
   * Cell text is joined using tabs between columns and line breaks between
   * rows.
   */
  getText(opts?: ListTablesOptions): Promise<string | null>;
  /**
   * Reads the current table cell text as a two-dimensional array.
   *
   * Returns `null` when the table no longer exists.
   */
  getCells(opts?: ListTablesOptions): Promise<string[][] | null>;
  /** Adds a comment to the whole table range. */
  addComment(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /** Deletes this table from the document. */
  delete(opts?: TableDeleteOptions): Promise<ApplyResult>;
  /**
   * Replaces the text of one cell, addressed by zero-based row and column.
   */
  replaceCellText(
    row: number,
    col: number,
    text: string,
    opts?: TableCellWriteOptions,
  ): Promise<ApplyResult>;
  /** Deletes one row, addressed by zero-based row index. */
  deleteRow(row: number, opts?: TableRowDeleteOptions): Promise<ApplyResult>;
  /**
   * Inserts one row.
   *
   * `templateRow` is the zero-based source row used for row structure.
   * `position` is the zero-based insertion position, or `'end'`.
   */
  insertRow(
    spec: { templateRow: number; position: number | 'end'; cellTexts?: string[] },
    opts?: LiveObjectWriteOptions,
  ): Promise<ApplyResult>;
  /**
   * Sets table grid column widths in twips.
   *
   * The array length should match the table column count.
   */
  setColumnWidths(widths: number[], opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets row heights.
   *
   * Pass `null` for a row that should keep its current height.
   */
  setRowHeights(heights: (RowHeight | null)[], opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

/**
 * Live handle for one table cell.
 */
export interface TableCell {
  /** Parent table id. */
  readonly tableId: string;
  /** Zero-based row index. */
  readonly row: number;
  /** Zero-based column index. */
  readonly col: number;
  /**
   * Reads the current cell metadata.
   */
  getInfo(opts?: ListTablesOptions): Promise<TableCellInfo | null>;
  /**
   * Reads direct content blocks in this cell, preserving cell content order.
   */
  getBlocks(opts?: ListTablesOptions): Promise<TableCellBlock[] | null>;
  /**
   * Reads live paragraph handles for direct paragraphs in this cell.
   */
  getParagraphs(opts?: ListTablesOptions): Promise<Paragraph[]>;
  /**
   * Reads the current plain text of this cell.
   */
  getText(opts?: ListTablesOptions): Promise<string | null>;
  /**
   * Creates a text range handle by searching inside this cell's paragraphs.
   */
  findText(search: string, opts?: { scope?: 'first' | 'all' }): CellTextRange;
  /**
   * Replaces this cell with plain text.
   */
  replaceWith(text: string, opts?: TableCellWriteOptions): Promise<ApplyResult>;
  /**
   * Inserts a new paragraph before the first direct content block in this cell.
   */
  insertParagraphAtStart(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Inserts a new paragraph after the last direct content block in this cell.
   */
  insertParagraphAtEnd(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets direct cell background shading fill color.
   */
  setShading(fill: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets direct cell vertical alignment.
   */
  setVerticalAlignment(
    value: NonNullable<TableCellDirectFormat['verticalAlignment']>,
    opts?: LiveObjectWriteOptions,
  ): Promise<ApplyResult>;
  /**
   * Sets direct cell width in twips.
   */
  setWidth(width: number, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

export type TableCellBlock =
  | {
      type: 'paragraph';
      paragraph: Paragraph;
      paraId: string;
      text: string;
      index: number;
    }
  | {
      type: 'drawing';
      drawing: Drawing | null;
      drawingId: string;
      paraId?: string;
      index: number;
    }
  | {
      type: 'table';
      table: Table | null;
      tableId: string;
      index: number;
    };

/**
 * Live handle for text matched inside a table cell.
 */
export interface CellTextRange {
  replaceWith(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  delete(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  setBold(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  setColor(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  setItalic(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  setUnderline(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  setHighlight(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  setFontSize(value: number, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  setFontFamily(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  addComment(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  createBookmark(name: BookmarkName, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

/**
 * Live handle for one table row.
 */
export interface TableRow {
  /** Parent table id. */
  readonly tableId: string;
  /** Zero-based row index. */
  readonly index: number;
  /**
   * Reads the current cell texts for this row.
   */
  getCells(opts?: ListTablesOptions): Promise<string[] | null>;
  /**
   * Reads the current plain text of this row.
   *
   * Cell text is joined using tabs.
   */
  getText(opts?: ListTablesOptions): Promise<string | null>;
  /**
   * Deletes this row.
   */
  delete(opts?: TableRowDeleteOptions): Promise<ApplyResult>;
  /**
   * Inserts a row before this row, using this row as the structural template.
   */
  insertBefore(cellTexts?: string[], opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Inserts a row after this row, using this row as the structural template.
   */
  insertAfter(cellTexts?: string[], opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Sets the row height. Pass `null` to remove the explicit row height.
   */
  setHeight(height: RowHeight | null, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

/**
 * Live handle for one table column.
 */
export interface TableColumn {
  /** Parent table id. */
  readonly tableId: string;
  /** Zero-based column index. */
  readonly index: number;
  /**
   * Reads the current cell texts for this column.
   */
  getCells(opts?: ListTablesOptions): Promise<string[] | null>;
  /**
   * Reads the current plain text of this column.
   *
   * Cell text is joined using line breaks.
   */
  getText(opts?: ListTablesOptions): Promise<string | null>;
  /**
   * Sets this column width in twips.
   */
  setWidth(width: number, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

export interface SectionsApi {
  /**
   * 列出主 body 的所有节（#211）。节按 doc-order 排序。
   * 无 body 或文档无 sectPr 时返回空数组。
   */
  list(): Promise<SectionInfo[]>;
}

export interface TocFieldsApi {
  /**
   * 枚举主 body 中的所有 TOC 域（#197）。doc-order 排序。
   * 无 TOC 域时返回空数组。
   */
  list(): Promise<TocFieldRange[]>;
}

export interface DrawingsApi {
  /**
   * 枚举文档 drawing 对象（#189）。
   * - scope: 'body'（默认）只返回主文档 body 内的 drawing
   * - scope: 'all' 包含 headers / footers / footnotes / endnotes / comments
   *
   * 本 API 只暴露形态信号（kind / extent / wrap / position / isInHeaderFooter 等），
   * 不下 A / B / C 业务分类结论 —— 那是 #191 / #192 的事。
   */
  list(opts?: ListDrawingsOptions): Promise<Drawing[]>;
  /**
   * Lists picture drawings only.
   *
   * Prefer this over filtering `list()` when the task is explicitly about
   * images.
   */
  listPictures(opts?: ListDrawingsOptions): Promise<PictureDrawing[]>;
  /**
   * Lists textbox drawings only.
   *
   * Prefer this when the task is explicitly about text boxes.
   */
  listTextBoxes(opts?: ListDrawingsOptions): Promise<TextBoxDrawing[]>;
  /**
   * Lists plain shape drawings only.
   *
   * This does not include pictures, text boxes, charts, SmartArt, groups, or
   * canvases.
   */
  listShapes(opts?: ListDrawingsOptions): Promise<ShapeDrawing[]>;
  /**
   * 按 drawingId 精确取单个 drawing，O(1)。不存在返回 null。
   */
  get(drawingId: string): Promise<Drawing | null>;
}

export interface DrawingImageBytes {
  bytes: Uint8Array;
  contentType: string;
  extentEMU: { cx: number; cy: number };
}

/**
 * Live handle for one drawing object.
 */
export interface Drawing extends DrawingInfo {
  /** Stable alias of `drawingId`, kept for object-style APIs. */
  readonly id: string;
  /**
   * Reads current drawing metadata.
   */
  getInfo(): Promise<DrawingInfo | null>;
  /**
   * Reads original image bytes when this drawing is a picture.
   */
  getImageBytes(): Promise<DrawingImageBytes | null>;
  /**
   * Deletes this drawing.
   */
  delete(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Replaces a picture drawing with the shared placeholder image.
   */
  replaceWithPlaceholder(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Replaces text inside a shape/textbox drawing.
   */
  replaceShapeInnerText(placeholder: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

// ─────────── 工厂:持有 transport 引用 + dispose 检查 ───────────

/** `WordDocument.assertNotDisposed` 的桥接签名;避免 sub-apis.ts 反向依赖 WordDocument */
/**
 * Live handle for a picture drawing.
 */
export interface PictureDrawing extends Drawing {
  readonly kind: 'picture';
  /**
   * Reads original image bytes for this picture drawing.
   */
  getImageBytes(): Promise<DrawingImageBytes | null>;
  /**
   * Replaces this picture with the shared placeholder image while preserving layout.
   */
  replaceWithPlaceholder(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
  /**
   * Deletes this picture drawing.
   */
  delete(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

/**
 * Live handle for a textbox drawing.
 */
export interface TextBoxDrawing extends Drawing {
  readonly kind: 'textBox';
  /**
   * Reads the text currently inside this textbox.
   *
   * Returns `null` if the drawing no longer exists or is no longer a textbox.
   */
  getText(): Promise<string | null>;
  /**
   * Lists live paragraph handles inside this textbox.
   */
  getParagraphs(): Promise<Paragraph[]>;
  /**
   * Reads direct content blocks inside this textbox, preserving textbox content order.
   */
  getBlocks(): Promise<TextBoxBlock[]>;
  /**
   * Replaces all textbox content with a single paragraph containing `text`.
   */
  replaceText(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

export type TextBoxBlock =
  | {
      type: 'paragraph';
      paragraph: Paragraph;
      paraId: string;
      text: string;
      index: number;
    }
  | {
      type: 'drawing';
      drawing: Drawing | null;
      drawingId: string;
      paraId?: string;
      index: number;
    }
  | {
      type: 'table';
      table: Table | null;
      tableId: string;
      index: number;
    };

/**
 * Live handle for a plain shape drawing.
 */
export interface ShapeDrawing extends Drawing {
  readonly kind: 'shape';
  /**
   * Reads current plain-shape metadata such as border/fill and line hints.
   */
  getInfo(): Promise<(DrawingInfo & { kind: 'shape' }) | null>;
  /**
   * Deletes this plain shape drawing.
   */
  delete(opts?: LiveObjectWriteOptions): Promise<ApplyResult>;
}

export type AssertNotDisposed = () => void;

function liveObjectApplyOptions(opts: LiveObjectWriteOptions = {}): {
  trackChanges: boolean;
  author?: string;
  date?: string;
} {
  return {
    trackChanges: opts.trackChanges ?? false,
    ...(opts.author !== undefined ? { author: opts.author } : {}),
    ...(opts.date !== undefined ? { date: opts.date.toISOString() } : {}),
  };
}

export function createParagraphApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): ParagraphApi {
  const createObject = (info: ParagraphInfo) =>
    new LiveParagraph(transport, assertNotDisposed, info);

  return {
    async list(opts) {
      assertNotDisposed();
      const result = await transport.call<ListResult>('paragraphs.list', [opts]);
      return {
        paragraphs: result.paragraphs.map(createObject),
        pagination: result.pagination,
      };
    },
    async get(paraId) {
      assertNotDisposed();
      const info = await transport.call<ParagraphInfo | null>('paragraphs.get', [paraId]);
      return info ? createObject(info) : null;
    },
  };
}

class LiveParagraph implements Paragraph {
  readonly id: ParaId;
  readonly paraId!: ParaId;
  readonly story!: ParagraphInfo['story'];
  readonly text!: string;
  readonly truncated?: boolean;
  readonly styleId?: string;
  readonly hasRevisions?: boolean;
  readonly hasComments?: boolean;
  readonly hasBookmarks?: boolean;
  readonly hasMath?: boolean;
  readonly format?: ParagraphInfo['format'];
  readonly table?: ParagraphInfo['table'];

  constructor(
    private readonly transport: Transport,
    private readonly assertNotDisposed: AssertNotDisposed,
    infoOrParaId: ParagraphInfo | ParaId,
    story?: ParagraphInfo['story'],
  ) {
    const info =
      typeof infoOrParaId === 'string'
        ? ({ paraId: infoOrParaId, story: story!, text: '' } satisfies ParagraphInfo)
        : infoOrParaId;
    Object.assign(this, info);
    this.id = info.paraId;
  }

  async getInfo(
    opts?: Pick<ListParagraphsOptions, 'includeFormat' | 'textFormat' | 'includeTableInfo'>,
  ): Promise<ParagraphInfo | null> {
    this.assertNotDisposed();
    if (opts?.includeFormat || opts?.textFormat) {
      const result = await this.transport.call<ListResult>('paragraphs.list', [
        {
          scope: [this.story],
          pagination: { mode: 'all' },
          ...opts,
        } satisfies ListParagraphsOptions,
      ]);
      return result.paragraphs.find((paragraph) => paragraph.paraId === this.paraId) ?? null;
    }
    return this.transport.call<ParagraphInfo | null>('paragraphs.get', [this.paraId]);
  }

  async getText(): Promise<string | null> {
    const info = await this.getInfo();
    return info?.text ?? null;
  }

  replaceWith(text: string, opts: ParagraphReplaceOptions = {}): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'replaceParagraph',
          paraId: this.paraId,
          text,
          onAnchorConflict: opts.onAnchorConflict ?? 'error',
          ...(opts.keepInlineGraphics !== undefined
            ? { keepInlineGraphics: opts.keepInlineGraphics }
            : {}),
        },
      ],
      opts,
    );
  }

  setAlignment(
    value: NonNullable<ParagraphDirectFormat['alignment']>,
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.applyParagraphFormat({ alignment: value }, opts);
  }

  setOutlineLevel(value: number, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyParagraphFormat({ outlineLevel: value }, opts);
  }

  setPageBreakBefore(value: boolean, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyParagraphFormat({ pageBreakBefore: value }, opts);
  }

  setIndent(
    value: NonNullable<ParagraphDirectFormat['indent']>,
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.applyParagraphFormat({ indent: value }, opts);
  }

  setSpacing(
    value: NonNullable<ParagraphDirectFormat['spacing']>,
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.applyParagraphFormat({ spacing: value }, opts);
  }

  insertBefore(text: string, opts: ParagraphInsertOptions = {}): Promise<ApplyResult> {
    return this.insertParagraph('insertParagraphBefore', text, opts);
  }

  insertAfter(text: string, opts: ParagraphInsertOptions = {}): Promise<ApplyResult> {
    return this.insertParagraph('insertParagraphAfter', text, opts);
  }

  insertImageAfter(
    spec: ParagraphInsertImageSpec,
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'insertImageAfter',
          anchorParaId: this.paraId,
          bytes: spec.bytes,
          contentType: spec.contentType,
          widthEmu: spec.widthEmu,
          heightEmu: spec.heightEmu,
        },
      ],
      opts,
    );
  }

  insertTableAfter(
    spec: ParagraphInsertTableSpec,
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'insertTableFrom',
          anchorParaId: this.paraId,
          tableXml: buildSimpleTableXml(spec),
        },
      ],
      opts,
    );
  }

  delete(opts: ParagraphDeleteOptions = {}): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'deleteParagraph',
          paraId: this.paraId,
          onAnchorConflict: opts.onAnchorConflict ?? 'error',
        },
      ],
      opts,
    );
  }

  addComment(text: string, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply([{ type: 'addComment', range: this.paragraphRange(), text }], opts);
  }

  createBookmark(name: BookmarkName, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply([{ type: 'createBookmark', name, range: this.paragraphRange() }], opts);
  }

  findText(search: string, opts: { scope?: 'first' | 'all' } = {}): TextRange {
    this.assertNotDisposed();
    return new LiveTextRange(
      this.transport,
      this.assertNotDisposed,
      this.paraId,
      search,
      opts.scope ?? 'first',
    );
  }

  private paragraphRange(): Range {
    return { paraId: this.paraId };
  }

  private applyParagraphFormat(
    format: ParagraphDirectFormat,
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply([{ type: 'formatParagraph', paraId: this.paraId, format }], opts);
  }

  private insertParagraph(
    type: 'insertParagraphBefore' | 'insertParagraphAfter',
    text: string,
    opts: ParagraphInsertOptions = {},
  ): Promise<ApplyResult> {
    const spec: Extract<Op, { type: typeof type }>['spec'] = { text };
    if (opts.style !== undefined) spec.style = opts.style;
    if (opts.runFormat !== undefined) spec.runFormat = opts.runFormat;
    return this.apply([{ type, anchorParaId: this.paraId, spec }], opts);
  }

  private apply(ops: Op[], opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    this.assertNotDisposed();
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops,
        options: liveObjectApplyOptions(opts),
      },
    ]);
  }
}

export function createContextApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): ContextApi {
  return {
    search(opts) {
      assertNotDisposed();
      return transport.call<ContextSearchResult>('context.search', [opts]);
    },
  };
}

function buildSimpleTableXml(spec: ParagraphInsertTableSpec): string {
  const rows = spec.rows;
  if (rows.length === 0) {
    throw new Error('insertTableAfter requires at least one row');
  }

  const colCount = Math.max(...rows.map((row) => row.length));
  if (colCount === 0) {
    throw new Error('insertTableAfter requires at least one column');
  }

  const widths = Array.from({ length: colCount }, (_, index) =>
    normalizeTwips(spec.columnWidths?.[index] ?? 2400),
  );
  const grid = widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('');
  const style = spec.tableStyleId
    ? `<w:tblStyle w:val="${escapeXmlAttribute(spec.tableStyleId)}"/>`
    : '';
  const body = rows
    .map((row) => {
      const cells = widths
        .map((width, index) => buildSimpleTableCell(row[index] ?? '', width))
        .join('');
      return `<w:tr>${cells}</w:tr>`;
    })
    .join('');

  return `<w:tbl xmlns:w="${W_NS}"><w:tblPr>${style}<w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

function buildSimpleTableCell(text: string, width: number): string {
  const paragraph =
    text === ''
      ? '<w:p/>'
      : `<w:p><w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r></w:p>`;
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${paragraph}</w:tc>`;
}

function normalizeTwips(value: number): number {
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error('table column widths must be positive numbers');
  }
  return Math.round(value);
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

class LiveTextRange implements TextRange {
  constructor(
    private readonly transport: Transport,
    private readonly assertNotDisposed: AssertNotDisposed,
    private readonly paraId: ParaId,
    private readonly search: string,
    private readonly scope: 'first' | 'all',
  ) {}

  replaceWith(text: string, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'replaceTextInParagraph',
          paraId: this.paraId,
          search: this.search,
          replacement: text,
          scope: this.scope,
        },
      ],
      opts,
    );
  }

  delete(opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.replaceWith('', opts);
  }

  setBold(value: boolean, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyTextFormat({ bold: value }, opts);
  }

  setColor(value: string, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyTextFormat({ color: value }, opts);
  }

  setItalic(value: boolean, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyTextFormat({ italic: value }, opts);
  }

  setUnderline(value: boolean, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyTextFormat({ underline: value }, opts);
  }

  setHighlight(value: string, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyTextFormat({ highlight: value }, opts);
  }

  setFontSize(value: number, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyTextFormat({ fontSize: value }, opts);
  }

  setFontFamily(value: string, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyTextFormat({ fontFamily: value }, opts);
  }

  private applyTextFormat(
    format: Extract<Op, { type: 'formatTextInParagraph' }>['format'],
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'formatTextInParagraph',
          paraId: this.paraId,
          search: this.search,
          format,
          scope: this.scope,
        },
      ],
      opts,
    );
  }

  addComment(text: string, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply([{ type: 'addComment', range: this.toRangeDescriptor(), text }], opts);
  }

  createBookmark(name: BookmarkName, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply([{ type: 'createBookmark', name, range: this.toRangeDescriptor() }], opts);
  }

  private toRangeDescriptor(): Range {
    return { paraId: this.paraId, text: this.search, scope: this.scope };
  }

  private apply(ops: Op[], opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    this.assertNotDisposed();
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops,
        options: liveObjectApplyOptions(opts),
      },
    ]);
  }
}

export function createBookmarkApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): BookmarkApi {
  const createObject = (info: BookmarkInfo) => new LiveBookmark(transport, assertNotDisposed, info);

  return {
    async list(opts) {
      assertNotDisposed();
      const result = await transport.call<BookmarkInfo[]>('bookmarks.list', [opts]);
      return result.map(createObject);
    },
    async get(name) {
      assertNotDisposed();
      const info = await transport.call<BookmarkInfo | null>('bookmarks.get', [name]);
      return info ? createObject(info) : null;
    },
  };
}

class LiveBookmark implements Bookmark {
  readonly id: BookmarkName;
  readonly name!: BookmarkName;
  readonly story!: BookmarkInfo['story'];
  readonly rangeKind!: BookmarkInfo['rangeKind'];
  readonly paraId?: string;
  readonly text?: string;
  readonly hidden?: boolean;

  constructor(
    private readonly transport: Transport,
    private readonly assertNotDisposed: AssertNotDisposed,
    info: BookmarkInfo,
  ) {
    Object.assign(this, info);
    this.id = info.name;
  }

  getInfo(): Promise<BookmarkInfo | null> {
    this.assertNotDisposed();
    return this.transport.call<BookmarkInfo | null>('bookmarks.get', [this.name]);
  }

  async getText(): Promise<string | null> {
    const info = await this.getInfo();
    if (!info) return null;
    if (info.rangeKind === 'text') return info.text ?? '';
    if (!info.paraId) return null;
    const paragraph = await this.getParagraph();
    return paragraph?.getText() ?? null;
  }

  async getParagraph(): Promise<Paragraph | null> {
    const info = await this.getInfo();
    if (!info?.paraId) return null;
    const paragraph = await this.transport.call<ParagraphInfo | null>('paragraphs.get', [
      info.paraId,
    ]);
    return paragraph
      ? new LiveParagraph(this.transport, this.assertNotDisposed, paragraph.paraId, paragraph.story)
      : null;
  }

  delete(opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    this.assertNotDisposed();
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops: [{ type: 'deleteBookmark', name: this.name }],
        options: liveObjectApplyOptions(opts),
      },
    ]);
  }
}

export function createCommentApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): CommentApi {
  const createObject = (info: CommentInfo) => new LiveComment(transport, assertNotDisposed, info);

  return {
    async list() {
      assertNotDisposed();
      const result = await transport.call<CommentInfo[]>('comments.list', []);
      return result.map(createObject);
    },
    async get(id) {
      assertNotDisposed();
      const info = await transport.call<CommentInfo | null>('comments.get', [id]);
      return info ? createObject(info) : null;
    },
  };
}

class LiveComment implements Comment {
  readonly id!: CommentId;
  readonly author!: string;
  readonly date!: string;
  readonly text!: string;
  readonly rangeKind!: CommentInfo['rangeKind'];
  readonly paraId?: string;
  readonly anchorText?: string;

  constructor(
    private readonly transport: Transport,
    private readonly assertNotDisposed: AssertNotDisposed,
    info: CommentInfo,
  ) {
    Object.assign(this, info);
  }

  getInfo(): Promise<CommentInfo | null> {
    this.assertNotDisposed();
    return this.transport.call<CommentInfo | null>('comments.get', [this.id]);
  }

  async getText(): Promise<string | null> {
    return (await this.getInfo())?.text ?? null;
  }

  async getAnchorText(): Promise<string | null> {
    const info = await this.getInfo();
    if (!info) return null;
    if (info.rangeKind === 'text') return info.anchorText ?? '';
    if (!info.paraId) return null;
    const paragraph = await this.getParagraph();
    return paragraph?.getText() ?? null;
  }

  async getParagraph(): Promise<Paragraph | null> {
    const info = await this.getInfo();
    if (!info?.paraId) return null;
    const paragraph = await this.transport.call<ParagraphInfo | null>('paragraphs.get', [
      info.paraId,
    ]);
    return paragraph
      ? new LiveParagraph(this.transport, this.assertNotDisposed, paragraph.paraId, paragraph.story)
      : null;
  }

  delete(opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    this.assertNotDisposed();
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops: [{ type: 'deleteComment', commentId: this.id }],
        options: liveObjectApplyOptions(opts),
      },
    ]);
  }
}

export function createRevisionApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): RevisionApi {
  const createObject = (info: RevisionInfo) => new LiveRevision(transport, assertNotDisposed, info);

  return {
    async list(opts) {
      assertNotDisposed();
      const result = await transport.call<RevisionInfo[]>('revisions.list', [opts]);
      return result.map(createObject);
    },
    async get(revId, opts) {
      assertNotDisposed();
      const info = await transport.call<RevisionInfo | null>('revisions.get', [revId, opts]);
      return info ? createObject(info) : null;
    },
    accept(opts = {}) {
      assertNotDisposed();
      return applyRevisionBulk(transport, 'acceptRevisions', opts);
    },
    reject(opts = {}) {
      assertNotDisposed();
      return applyRevisionBulk(transport, 'rejectRevisions', opts);
    },
  };
}

function applyRevisionBulk(
  transport: Transport,
  type: 'acceptRevisions' | 'rejectRevisions',
  opts: RevisionBulkOptions = {},
): Promise<ApplyResult> {
  const op: Op = { type, scope: { scope: opts.scope ?? 'all' } };
  return transport.call<ApplyResult>('applyPlan', [
    {
      ops: [op],
      options: liveObjectApplyOptions(opts),
    },
  ]);
}

class LiveRevision implements Revision {
  readonly id: RevisionInfo['revId'];
  readonly revId!: RevisionInfo['revId'];
  readonly kind!: RevisionInfo['kind'];
  readonly ooxmlType!: RevisionInfo['ooxmlType'];
  readonly author!: string;
  readonly date!: string;
  readonly story!: RevisionInfo['story'];
  readonly paraId?: string;
  readonly text?: string;

  constructor(
    private readonly transport: Transport,
    private readonly assertNotDisposed: AssertNotDisposed,
    info: RevisionInfo,
  ) {
    Object.assign(this, info);
    this.id = info.revId;
  }

  getInfo(opts?: ListRevisionsOptions): Promise<RevisionInfo | null> {
    this.assertNotDisposed();
    return this.transport.call<RevisionInfo | null>('revisions.get', [this.revId, opts]);
  }

  async getText(): Promise<string | null> {
    return (await this.getInfo({ scope: 'all' }))?.text ?? null;
  }

  async getParagraph(): Promise<Paragraph | null> {
    const info = await this.getInfo({ scope: 'all' });
    if (!info?.paraId) return null;
    const paragraph = await this.transport.call<ParagraphInfo | null>('paragraphs.get', [
      info.paraId,
    ]);
    return paragraph
      ? new LiveParagraph(this.transport, this.assertNotDisposed, paragraph.paraId, paragraph.story)
      : null;
  }

  accept(opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyRevision('acceptRevisions', opts);
  }

  reject(opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.applyRevision('rejectRevisions', opts);
  }

  private applyRevision(
    type: 'acceptRevisions' | 'rejectRevisions',
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    this.assertNotDisposed();
    const op: Op = { type, scope: { scope: 'revision', revId: this.revId } };
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops: [op],
        options: liveObjectApplyOptions(opts),
      },
    ]);
  }
}

export function createStylesApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): StylesApi {
  return {
    async list(opts) {
      assertNotDisposed();
      return transport.call<StyleInfo[]>('styles.list', [opts]);
    },
    async get(styleId) {
      assertNotDisposed();
      return transport.call<StyleInfo | null>('styles.get', [styleId]);
    },
  };
}

export function createSectionsApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): SectionsApi {
  return {
    async list() {
      assertNotDisposed();
      return transport.call<SectionInfo[]>('sections.list', []);
    },
  };
}

export function createTablesApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): TablesApi {
  const createObject = (info: TableInfo) => new LiveTable(transport, assertNotDisposed, info);

  return {
    async list(opts) {
      assertNotDisposed();
      const result = await transport.call<TableInfo[]>('tables.list', [opts]);
      return result.map(createObject);
    },
    async get(tableId) {
      assertNotDisposed();
      const info = await transport.call<TableInfo | null>('tables.get', [tableId]);
      return info ? createObject(info) : null;
    },
    async getVerbatim(tableId) {
      assertNotDisposed();
      return transport.call<TableVerbatim | null>('tables.getVerbatim', [tableId]);
    },
  };
}

class LiveTable implements Table {
  readonly id: string;
  readonly tableId!: string;
  readonly partPath!: string;
  readonly index!: number;
  readonly isNested!: boolean;
  readonly rowCount!: number;
  readonly colCount!: number;
  readonly hasMerge!: boolean;
  readonly tableStyleId?: string;
  readonly hasImages!: boolean;
  readonly hasNestedTables!: boolean;
  readonly prevParaId?: string;
  readonly nextParaId?: string;
  readonly columnWidths!: number[];
  readonly rowHeights!: (RowHeight | null)[];
  readonly cells?: string[][];
  readonly cellGrid?: TableInfo['cellGrid'];

  constructor(
    private readonly transport: Transport,
    private readonly assertNotDisposed: AssertNotDisposed,
    info: TableInfo,
  ) {
    Object.assign(this, info);
    this.id = info.tableId;
  }

  cell(row: number, col: number): TableCell {
    this.assertNotDisposed();
    return new LiveTableCell(this.transport, this.assertNotDisposed, this, row, col);
  }

  row(row: number): TableRow {
    this.assertNotDisposed();
    return new LiveTableRow(this, row);
  }

  column(col: number): TableColumn {
    this.assertNotDisposed();
    return new LiveTableColumn(this, col);
  }

  async getInfo(opts?: ListTablesOptions): Promise<TableInfo | null> {
    this.assertNotDisposed();
    if (!opts) {
      return this.transport.call<TableInfo | null>('tables.get', [this.tableId]);
    }
    const result = await this.transport.call<TableInfo[]>('tables.list', [
      {
        scope: 'all',
        includeNested: true,
        ...opts,
      } satisfies ListTablesOptions,
    ]);
    return result.find((table) => table.tableId === this.tableId) ?? null;
  }

  getVerbatim(): Promise<TableVerbatim | null> {
    this.assertNotDisposed();
    return this.transport.call<TableVerbatim | null>('tables.getVerbatim', [this.tableId]);
  }

  async getText(opts?: ListTablesOptions): Promise<string | null> {
    const cells = await this.getCells(opts);
    return cells?.map((row) => row.join('\t')).join('\n') ?? null;
  }

  async getCells(opts?: ListTablesOptions): Promise<string[][] | null> {
    const info = await this.getInfo({
      scope: 'all',
      includeNested: true,
      ...opts,
      includeCellTexts: true,
    });
    return info?.cells ?? null;
  }

  addComment(text: string, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply([{ type: 'addTableComment', tableId: this.tableId, text }], opts);
  }

  delete(opts: TableDeleteOptions = {}): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'deleteTable',
          tableId: this.tableId,
          onAnchorConflict: opts.onAnchorConflict ?? 'error',
        },
      ],
      opts,
    );
  }

  replaceCellText(
    row: number,
    col: number,
    text: string,
    opts: TableCellWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'replaceTableCellText',
          tableId: this.tableId,
          row,
          col,
          text,
          onAnchorConflict: opts.onAnchorConflict ?? 'error',
        },
      ],
      opts,
    );
  }

  deleteRow(row: number, opts: TableRowDeleteOptions = {}): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'deleteTableRow',
          tableId: this.tableId,
          row,
          onAnchorConflict: opts.onAnchorConflict ?? 'error',
        },
      ],
      opts,
    );
  }

  insertRow(
    spec: { templateRow: number; position: number | 'end'; cellTexts?: string[] },
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'insertTableRow',
          tableId: this.tableId,
          templateRow: spec.templateRow,
          position: spec.position,
          cellTexts: spec.cellTexts,
        },
      ],
      opts,
    );
  }

  setColumnWidths(widths: number[], opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply([{ type: 'setTableColumnWidths', tableId: this.tableId, widths }], opts);
  }

  setRowHeights(
    heights: (RowHeight | null)[],
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply([{ type: 'setTableRowHeights', tableId: this.tableId, heights }], opts);
  }

  private apply(ops: Op[], opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    this.assertNotDisposed();
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops,
        options: liveObjectApplyOptions(opts),
      },
    ]);
  }
}

class LiveTableCell implements TableCell {
  readonly tableId: string;

  constructor(
    private readonly transport: Transport,
    private readonly assertNotDisposed: AssertNotDisposed,
    private readonly table: Table,
    readonly row: number,
    readonly col: number,
  ) {
    this.tableId = table.tableId;
  }

  async getInfo(opts?: ListTablesOptions): Promise<TableCellInfo | null> {
    const info = await this.table.getInfo({
      scope: 'all',
      includeNested: true,
      ...opts,
      includeCellTexts: true,
    });
    return info?.cellGrid?.[this.row]?.[this.col] ?? null;
  }

  async getBlocks(opts?: ListTablesOptions): Promise<TableCellBlock[] | null> {
    const info = await this.getInfo(opts);
    if (!info) return null;
    const blocks: TableCellBlock[] = [];
    for (let index = 0; index < info.blocks.length; index++) {
      const block = info.blocks[index]!;
      const mapped = await this.toLiveBlock(block, index);
      if (mapped) blocks.push(mapped);
    }
    return blocks;
  }

  async getParagraphs(opts?: ListTablesOptions): Promise<Paragraph[]> {
    const info = await this.getInfo(opts);
    if (!info) return [];
    const paragraphs = await Promise.all(
      info.paragraphIds.map((paraId) =>
        this.transport.call<ParagraphInfo | null>('paragraphs.get', [paraId]),
      ),
    );
    return paragraphs
      .filter((paragraph): paragraph is ParagraphInfo => paragraph !== null)
      .map(
        (paragraph) =>
          new LiveParagraph(
            this.transport,
            this.assertNotDisposed,
            paragraph.paraId,
            paragraph.story,
          ),
      );
  }

  async getText(opts?: ListTablesOptions): Promise<string | null> {
    return (await this.getInfo(opts))?.text ?? null;
  }

  findText(search: string, opts: { scope?: 'first' | 'all' } = {}): CellTextRange {
    this.assertNotDisposed();
    return new LiveCellTextRange(this, search, opts.scope ?? 'first');
  }

  replaceWith(text: string, opts?: TableCellWriteOptions): Promise<ApplyResult> {
    return this.table.replaceCellText(this.row, this.col, text, opts);
  }

  insertParagraphAtStart(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.insertParagraph(text, 'start', opts);
  }

  insertParagraphAtEnd(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.insertParagraph(text, 'end', opts);
  }

  setShading(fill: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.applyCellFormat({ shading: { fill } }, opts);
  }

  setVerticalAlignment(
    value: NonNullable<TableCellDirectFormat['verticalAlignment']>,
    opts?: LiveObjectWriteOptions,
  ): Promise<ApplyResult> {
    return this.applyCellFormat({ verticalAlignment: value }, opts);
  }

  setWidth(width: number, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.applyCellFormat({ width }, opts);
  }

  private async toLiveBlock(
    block: TableCellBlockInfo,
    index: number,
  ): Promise<TableCellBlock | null> {
    if (block.type === 'drawing') {
      const info = await this.transport.call<DrawingInfo | null>('drawings.get', [block.drawingId]);
      const drawing = info ? new LiveDrawing(this.transport, this.assertNotDisposed, info) : null;
      return 'paraId' in block && block.paraId
        ? { type: 'drawing', drawing, drawingId: block.drawingId, paraId: block.paraId, index }
        : { type: 'drawing', drawing, drawingId: block.drawingId, index };
    }
    if (block.type === 'table') {
      const info = await this.transport.call<TableInfo | null>('tables.get', [block.tableId]);
      return {
        type: 'table',
        tableId: block.tableId,
        table: info ? new LiveTable(this.transport, this.assertNotDisposed, info) : null,
        index,
      };
    }

    const info = await this.transport.call<ParagraphInfo | null>('paragraphs.get', [block.paraId]);
    if (!info) return null;
    return {
      type: 'paragraph',
      paragraph: new LiveParagraph(this.transport, this.assertNotDisposed, info.paraId, info.story),
      paraId: block.paraId,
      text: block.text,
      index,
    };
  }

  private applyCellFormat(
    format: TableCellDirectFormat,
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'formatTableCell',
          tableId: this.tableId,
          row: this.row,
          col: this.col,
          format,
        },
      ],
      opts,
    );
  }

  private insertParagraph(
    text: string,
    position: 'start' | 'end',
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply(
      [
        {
          type: 'insertTableCellParagraph',
          tableId: this.tableId,
          row: this.row,
          col: this.col,
          text,
          position,
        },
      ],
      opts,
    );
  }

  private apply(ops: Op[], opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    this.assertNotDisposed();
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops,
        options: liveObjectApplyOptions(opts),
      },
    ]);
  }
}

class LiveCellTextRange implements CellTextRange {
  constructor(
    private readonly cell: TableCell,
    private readonly search: string,
    private readonly scope: 'first' | 'all',
  ) {}

  async replaceWith(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.replaceWith(text, opts));
  }

  async delete(opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.replaceWith('', opts);
  }

  async setBold(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.setBold(value, opts));
  }

  async setColor(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.setColor(value, opts));
  }

  async setItalic(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.setItalic(value, opts));
  }

  async setUnderline(value: boolean, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.setUnderline(value, opts));
  }

  async setHighlight(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.setHighlight(value, opts));
  }

  async setFontSize(value: number, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.setFontSize(value, opts));
  }

  async setFontFamily(value: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.setFontFamily(value, opts));
  }

  async addComment(text: string, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.addComment(text, opts));
  }

  async createBookmark(name: BookmarkName, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.apply((range) => range.createBookmark(name, opts));
  }

  private async apply(op: (range: TextRange) => Promise<ApplyResult>): Promise<ApplyResult> {
    const paragraphs = await this.cell.getParagraphs();
    const matched: Paragraph[] = [];
    for (const paragraph of paragraphs) {
      const text = await paragraph.getText();
      if (text?.includes(this.search)) {
        matched.push(paragraph);
        if (this.scope === 'first') break;
      }
    }
    if (matched.length === 0) {
      throw new Error(
        `cell.findText: text '${this.search}' not found in table '${this.cell.tableId}' cell (${this.cell.row}, ${this.cell.col})`,
      );
    }

    let last: ApplyResult | undefined;
    for (const paragraph of matched) {
      last = await op(paragraph.findText(this.search, { scope: this.scope }));
    }
    return last!;
  }
}

class LiveTableRow implements TableRow {
  readonly tableId: string;

  constructor(
    private readonly table: Table,
    readonly index: number,
  ) {
    this.tableId = table.tableId;
  }

  async getCells(opts?: ListTablesOptions): Promise<string[] | null> {
    const cells = await this.table.getCells(opts);
    return cells?.[this.index] ?? null;
  }

  async getText(opts?: ListTablesOptions): Promise<string | null> {
    const cells = await this.getCells(opts);
    return cells?.join('\t') ?? null;
  }

  delete(opts?: TableRowDeleteOptions): Promise<ApplyResult> {
    return this.table.deleteRow(this.index, opts);
  }

  insertBefore(cellTexts?: string[], opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.table.insertRow({ templateRow: this.index, position: this.index, cellTexts }, opts);
  }

  insertAfter(cellTexts?: string[], opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    return this.table.insertRow(
      { templateRow: this.index, position: this.index + 1, cellTexts },
      opts,
    );
  }

  async setHeight(height: RowHeight | null, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    const info = await this.table.getInfo();
    const heights = (info?.rowHeights ?? this.table.rowHeights).slice();
    heights[this.index] = height;
    return this.table.setRowHeights(heights, opts);
  }
}

class LiveTableColumn implements TableColumn {
  readonly tableId: string;

  constructor(
    private readonly table: Table,
    readonly index: number,
  ) {
    this.tableId = table.tableId;
  }

  async getCells(opts?: ListTablesOptions): Promise<string[] | null> {
    const cells = await this.table.getCells(opts);
    if (!cells) return null;
    return cells.map((row) => row[this.index] ?? '');
  }

  async getText(opts?: ListTablesOptions): Promise<string | null> {
    const cells = await this.getCells(opts);
    return cells?.join('\n') ?? null;
  }

  async setWidth(width: number, opts?: LiveObjectWriteOptions): Promise<ApplyResult> {
    const info = await this.table.getInfo();
    const widths = (info?.columnWidths ?? this.table.columnWidths).slice();
    widths[this.index] = width;
    return this.table.setColumnWidths(widths, opts);
  }
}

export function createDrawingsApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): DrawingsApi {
  const createObject = (info: DrawingInfo) => new LiveDrawing(transport, assertNotDisposed, info);
  const listAll = async (opts?: ListDrawingsOptions): Promise<Drawing[]> => {
    assertNotDisposed();
    const drawings = await transport.call<DrawingInfo[]>('drawings.list', [opts]);
    return drawings.map(createObject);
  };

  return {
    list: listAll,
    async listPictures(opts) {
      const drawings = await listAll(opts);
      return drawings.filter((drawing): drawing is PictureDrawing => drawing.kind === 'picture');
    },
    async listTextBoxes(opts) {
      const drawings = await listAll(opts);
      return drawings.filter((drawing): drawing is TextBoxDrawing => drawing.kind === 'textBox');
    },
    async listShapes(opts) {
      const drawings = await listAll(opts);
      return drawings.filter((drawing): drawing is ShapeDrawing => drawing.kind === 'shape');
    },
    async get(drawingId) {
      assertNotDisposed();
      const info = await transport.call<DrawingInfo | null>('drawings.get', [drawingId]);
      return info ? createObject(info) : null;
    },
  };
}

class LiveDrawing implements Drawing {
  readonly id: string;
  readonly drawingId!: string;
  readonly partPath!: string;
  readonly index!: number;
  readonly anchoringMode!: DrawingInfo['anchoringMode'];
  readonly kind!: DrawingInfo['kind'];
  readonly extentEMU!: DrawingInfo['extentEMU'];
  readonly isInHeaderFooter!: boolean;
  readonly hasPicReference!: boolean;

  constructor(
    private readonly transport: Transport,
    private readonly assertNotDisposed: AssertNotDisposed,
    info: DrawingInfo,
  ) {
    Object.assign(this, info);
    this.id = info.drawingId;
  }

  getInfo(): Promise<DrawingInfo | null> {
    this.assertNotDisposed();
    return this.transport.call<DrawingInfo | null>('drawings.get', [this.drawingId]);
  }

  getImageBytes(): Promise<DrawingImageBytes | null> {
    this.assertNotDisposed();
    return this.transport.call<DrawingImageBytes | null>('drawings.getImageBytes', [
      this.drawingId,
    ]);
  }

  delete(opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply([{ type: 'deleteDrawing', drawingId: this.drawingId }], opts);
  }

  replaceWithPlaceholder(opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply([{ type: 'replaceDrawingWithPlaceholder', drawingId: this.drawingId }], opts);
  }

  replaceShapeInnerText(
    placeholder: string,
    opts: LiveObjectWriteOptions = {},
  ): Promise<ApplyResult> {
    return this.apply(
      [{ type: 'replaceShapeInnerText', drawingId: this.drawingId, placeholder }],
      opts,
    );
  }

  async getText(): Promise<string | null> {
    const info = await this.getInfo();
    if (!info || info.kind !== 'textBox') return null;
    return info.innerText ?? '';
  }

  async getParagraphs(): Promise<Paragraph[]> {
    const info = await this.getInfo();
    if (!info || info.kind !== 'textBox') return [];
    const paragraphs: Paragraph[] = [];
    for (const paraId of info.innerParagraphIds ?? []) {
      const paragraphInfo = await this.transport.call<ParagraphInfo | null>('paragraphs.get', [
        paraId,
      ]);
      if (paragraphInfo) {
        paragraphs.push(
          new LiveParagraph(
            this.transport,
            this.assertNotDisposed,
            paragraphInfo.paraId,
            paragraphInfo.story,
          ),
        );
      }
    }
    return paragraphs;
  }

  async getBlocks(): Promise<TextBoxBlock[]> {
    const info = await this.getInfo();
    if (!info || info.kind !== 'textBox') return [];
    const blocks: TextBoxBlock[] = [];
    for (let index = 0; index < (info.innerBlocks ?? []).length; index++) {
      const block = info.innerBlocks![index]!;
      const mapped = await this.toLiveTextBoxBlock(block, index);
      if (mapped) blocks.push(mapped);
    }
    return blocks;
  }

  replaceText(text: string, opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    return this.apply(
      [{ type: 'replaceShapeInnerText', drawingId: this.drawingId, placeholder: text }],
      opts,
    );
  }

  private async toLiveTextBoxBlock(
    block: TextBoxBlockInfo,
    index: number,
  ): Promise<TextBoxBlock | null> {
    if (block.type === 'drawing') {
      const info = await this.transport.call<DrawingInfo | null>('drawings.get', [block.drawingId]);
      const drawing = info ? new LiveDrawing(this.transport, this.assertNotDisposed, info) : null;
      return 'paraId' in block && block.paraId
        ? { type: 'drawing', drawing, drawingId: block.drawingId, paraId: block.paraId, index }
        : { type: 'drawing', drawing, drawingId: block.drawingId, index };
    }

    if (block.type === 'table') {
      const info = await this.transport.call<TableInfo | null>('tables.get', [block.tableId]);
      return {
        type: 'table',
        tableId: block.tableId,
        table: info ? new LiveTable(this.transport, this.assertNotDisposed, info) : null,
        index,
      };
    }

    const info = await this.transport.call<ParagraphInfo | null>('paragraphs.get', [block.paraId]);
    if (!info) return null;
    return {
      type: 'paragraph',
      paragraph: new LiveParagraph(this.transport, this.assertNotDisposed, info.paraId, info.story),
      paraId: block.paraId,
      text: block.text,
      index,
    };
  }

  private apply(ops: Op[], opts: LiveObjectWriteOptions = {}): Promise<ApplyResult> {
    this.assertNotDisposed();
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops,
        options: liveObjectApplyOptions(opts),
      },
    ]);
  }
}

export function createTocFieldsApi(
  transport: Transport,
  assertNotDisposed: AssertNotDisposed,
): TocFieldsApi {
  return {
    async list() {
      assertNotDisposed();
      return transport.call<TocFieldRange[]>('tocFields.list', []);
    },
  };
}
