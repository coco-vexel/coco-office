/**
 * WordPlan —— 变更 builder + apply 机制(docs/03.1 §3,#27)。
 *
 * 职责(本 PR):
 *   - 11 个链式 builder 方法 `return this`,**不返回 op 句柄**;每个方法只把 op
 *     push 进内部队列,错误集中到 `apply()` 抛(契约 §4.2)
 *   - `apply()` 状态机:
 *       - 二次 apply → `PLAN_ALREADY_APPLIED`
 *       - 空 plan → console.warn 不抛(契约默认语义;严格模式留未来)
 *       - 非空 plan → transport RPC 到 kernel 的 `applyPlan`
 *   - 多 plan 并存独立,不共享状态(每次 `doc.plan()` 新实例)
 *
 * **不含**任何 op 的真正执行 —— executor 一 PR 一张,#28 起陆续落地:
 *   - #28: replaceParagraph / replaceTextInParagraph
 *   - #29: insertParagraphBefore/After / deleteParagraph
 *   - #34: createBookmark / deleteBookmark
 *   - #35: addComment / deleteComment
 *   - #31/#32: acceptRevisions / rejectRevisions
 *
 * 当前 kernel 侧 `applyPlan` 对任何非空 ops 直接抛 `UNSUPPORTED_OOXML_FEATURE`。
 */
import type { StoryId } from './stories.js';
import type { Transport } from './transport.js';
import type { WordErrorCode } from './word-error.js';
import type { SdtSpec } from './sdt-builders.js';
import type { RowHeight } from './list-tables.js';
import { WordError } from './word-error.js';

// ─────────── 类型别名(docs/03.1 §1.1 / §1.3 / §1.7)───────────

export type ParaId = string;
export type BookmarkName = string;
export type CommentId = number;
export type RevisionId = string;

// ─────────── Range(docs/03.1 §1.2)───────────

/** 段落级定位 */
export interface ParagraphRange {
  paraId: ParaId;
}

/** 段落内关键词定位 */
export interface TextRange {
  paraId: ParaId;
  text: string;
  /** 默认 'first' */
  scope?: 'first' | 'all';
}

export type Range = ParagraphRange | TextRange;

// ─────────── RevisionScope(docs/03.1 §1.3)───────────

export type RevisionScope =
  | { scope: 'all' }
  | { scope: 'paragraph'; paraId: ParaId }
  | { scope: 'revision'; revId: RevisionId };

export interface ParagraphDirectFormat {
  alignment?: 'left' | 'center' | 'right' | 'both' | 'distribute';
  outlineLevel?: number;
  pageBreakBefore?: boolean;
  indent?: {
    left?: number;
    right?: number;
    /**
     * Positive value writes first-line indent; negative value writes hanging indent.
     */
    firstLine?: number;
  };
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: 'auto' | 'exact' | 'atLeast';
  };
}

export interface TableCellDirectFormat {
  shading?: {
    /** 6-digit hex fill color without '#', or 'auto'. */
    fill?: string;
    /** 6-digit hex foreground color without '#', or 'auto'. */
    color?: string;
    /** OOXML shading pattern. Defaults to 'clear' when only fill is provided. */
    val?: string;
  };
  verticalAlignment?: 'top' | 'center' | 'bottom' | 'both';
  /** Cell width in twips, written as tcW type='dxa'. */
  width?: number;
}

// ─────────── 插入段落 spec(docs/03.1 §1.7)───────────

export interface InsertParagraphSpec {
  text: string;
  /** 'inherit'(默认)继承 anchor 的 `<w:pPr>`;字符串为 styleId */
  style?: 'inherit' | string;
  /**
   * 控制新段内容 run 的 `<w:rPr>` 来源。
   *
   * - `'default'`(默认)：裸 run,不附加任何 rPr,行为与未传时完全一致。
   * - `'inherit'`：克隆 anchor 段**首个直属 `<w:r>`** 的 `<w:rPr>`(深拷贝 cloneNode(true))
   *   并应用到新段的内容 run。常用于「簇展开」场景:样式写在直接格式(rPr)而非命名样式上。
   *   anchor 无直属 run 或首 run 无 rPr 时退化为 `'default'`(不报错)。
   *   与 `color` 共存时:先继承 rPr,再由 color 覆盖/追加 `<w:color>`(color 优先级更高)。
   *   与 `style` 正交:style 管 pPr,runFormat 管 rPr,可独立组合。
   */
  runFormat?: 'inherit' | 'default';
  /**
   * 整段文字的字体颜色,6 位十六进制无 `#`(如 `'0070C0'` 标准蓝)。
   * 留空则继承样式默认色。
   * 颜色作为 run-level `<w:rPr><w:color w:val="..."/></w:rPr>` 写入新段唯一 run。
   */
  color?: string;
  /**
   * 从指定段（任意 `paraId`）克隆 pPr + 首 run rPr 到新段（全文重组场景）。
   *
   * - **优先级最高**：同时传 `style` / `runFormat` 时以 `styleSource` 为准（pPr + rPr 都来自 styleSource）。
   * - pPr：深拷贝 `<w:pPr>`；styleSource 段无 pPr 则新段也不加 pPr（退化，不报错）。
   * - rPr：克隆 styleSource 段**首个直属 `<w:r>`** 的 `<w:rPr>`；无直属 run 或首 run 无 rPr 则退化裸 run（不报错）。
   * - styleSource 段**不存在**时抛 `PARA_NOT_FOUND`。
   * - styleSource 段自身只读，不被修改。
   * - 与 `color` 共存：先克隆 rPr，再 color 覆盖/追加 `<w:color>`（color 优先，同 #12）。
   * - `styleSource === anchor` 时退化为 `style:'inherit' + runFormat:'inherit'` 效果（#12 兼容）。
   */
  styleSource?: ParaId;
  /**
   * 可选:插入段的同时**原子地**建一个跨段书签,range 涵盖 anchor 段 → 新插段。
   *
   * 调用方只声明命名规则(`prefix`),不传具体名 —— 系统按 `{prefix}_{anchorParaId}`
   * 生成,冲突自动追加 `_2`/`_3` 后缀;实际名通过 `ApplyResult.createdBookmarks[i].name`
   * 回传(同一 entry 的 `range` 给 fromParaId / toParaId)。
   *
   * `prefix` 校验:必须匹配 `^[A-Za-z_][A-Za-z0-9_]*$`(允许 `_` 起头作 Word hidden
   * bookmark);最终名长度受 OOXML 40 字符上限约束,过长抛 `BOOKMARK_NAME_INVALID`。
   *
   * trackChanges=true 时本字段触发 `UNSUPPORTED_OOXML_FEATURE`(与 createBookmark
   * 的现状一致,#31 再统一处理)。
   */
  bookmark?: { prefix: string };
  /**
   * 段落大纲级别（Word outline level），写入 `<w:pPr><w:outlineLvl w:val="N"/>`。
   *
   * - 取值 **0–8**（对应 Word「1 级标题…9 级标题」；heading-1→0, heading-2→1, …）。
   * - 省略（`undefined`）：不写 outlineLvl，保持 styleSource 克隆来的 pPr 原样。
   * - 0–8：在新段 pPr 写入/覆盖 `<w:outlineLvl>`（显式值优先于 styleSource 克隆来的）。
   * - 越界（< 0 或 > 8）：抛 `INVALID_ARGUMENT`。
   */
  outlineLevel?: number;
  /**
   * 段前分页（Word「段前分页」），写入 `<w:pPr><w:pageBreakBefore/>`。
   *
   * - `true`：该段强制从新页起始（比手工 `<w:br type="page"/>` 更鲁棒：绑在段上不随内容漂移）。
   * - 省略（`undefined`）/ `false`：不写该元素（继承样式或默认）。
   */
  pageBreakBefore?: boolean;
}

// ─────────── Plan 选项(docs/03.1 §2)───────────

export interface PlanOptions {
  /** 默认 false */
  trackChanges?: boolean;
  /** 默认走 `WordDocument.setDefaultAuthor` 配置的值 */
  author?: string;
  /** 默认 `new Date()` */
  date?: Date;
}

// ─────────── Op union(docs/03.1 §5)───────────

export type Op =
  | {
      type: 'replaceParagraph';
      paraId: ParaId;
      text: string;
      onAnchorConflict: 'error' | 'drop';
      /** 保留段内含图形的 run（w:drawing / w:pict / wps:wsp / v:line / w:object）。
       *  默认 false（现有行为：整段所有 run 全删）。
       *  true 时：只删纯文字 run，图形 run 原位保留；新文字 run 插在首个被删文字 run 位置。 */
      keepInlineGraphics?: boolean;
    }
  | {
      type: 'replaceTextInParagraph';
      paraId: ParaId;
      search: string;
      replacement: string;
      scope: 'first' | 'all';
    }
  | {
      type: 'formatTextInParagraph';
      paraId: ParaId;
      search: string;
      format: {
        bold?: boolean;
        color?: string;
        italic?: boolean;
        underline?: boolean;
        highlight?: string;
        fontSize?: number;
        fontFamily?: string;
      };
      scope: 'first' | 'all';
    }
  | {
      type: 'formatParagraph';
      paraId: ParaId;
      format: ParagraphDirectFormat;
    }
  | { type: 'insertParagraphBefore'; anchorParaId: ParaId; spec: InsertParagraphSpec }
  | { type: 'insertParagraphAfter'; anchorParaId: ParaId; spec: InsertParagraphSpec }
  // 无锚点在 <w:body> 起始插一个空段，返回新 paraId（ApplyResult.insertedBodyStartParaIds）。
  // 用途：body 完全无段落（极简/特殊文档，所有正文段被骨架删光）时，下游重建需要一个真实
  // 落点锚段。区别于 insertParagraphBefore/After——后者必须有 anchor 段；本 op 不需要。
  | { type: 'insertParagraphAtBodyStart' }
  | { type: 'deleteParagraph'; paraId: ParaId; onAnchorConflict: 'error' | 'drop' }
  | { type: 'createBookmark'; name: BookmarkName; range: Range }
  | { type: 'deleteBookmark'; name: BookmarkName }
  | {
      type: 'addComment';
      range: Range;
      text: string;
      author?: string;
      /** RPC 友好:按 ISO 8601 字符串序列化;builder 端接 Date,push 前 toISOString */
      date?: string;
    }
  | { type: 'deleteComment'; commentId: CommentId }
  | { type: 'acceptRevisions'; scope: RevisionScope }
  | { type: 'rejectRevisions'; scope: RevisionScope }
  // #186 表格批注与删除。批注实质挂在表格首单元格首段上,但语义入口是 tableId,
  // 避免调用方手算 firstCellFirstParaId。
  | {
      type: 'addTableComment';
      tableId: string;
      text: string;
      author?: string;
      /** ISO 8601 字符串 */
      date?: string;
    }
  | { type: 'deleteTable'; tableId: string; onAnchorConflict: 'error' | 'drop' }
  // #188 表格域化:把指定单元格首段的 runs 替换为 inline SDT(Plain Text / Date / Rich Text)
  | {
      type: 'replaceTableCellWithSdt';
      tableId: string;
      /** 0-based 行索引(对应 TableInfo.cells 网格的 i)*/
      row: number;
      /** 0-based 列索引(对应 cells[i] 的 j) */
      col: number;
      sdt: SdtSpec;
      onAnchorConflict: 'error' | 'drop';
    }
  // #188 §7.2 行级裁剪:删除指定行(可变行数表保留 1~N 样例行,其余删除)
  | {
      type: 'deleteTableRow';
      tableId: string;
      /** 0-based 行索引(按删除时刻表格状态计) */
      row: number;
      onAnchorConflict: 'error' | 'drop';
    }
  // 填纯文字版(排版回填):保留 tcPr(底纹 shd / 边框 / 对齐 vAlign / 列宽 tcW)+ 首段 pPr + 首 run rPr
  | {
      type: 'replaceTableCellText';
      tableId: string;
      /** 0-based 行索引(对应 TableInfo.cells 网格的 i)*/
      row: number;
      /** 0-based 列索引(对应 cells[i] 的 j)*/
      col: number;
      text: string;
      onAnchorConflict: 'error' | 'drop';
    }
  | {
      type: 'formatTableCell';
      tableId: string;
      /** 0-based row index, matching TableInfo.cellGrid[row][col]. */
      row: number;
      /** 0-based column index, matching TableInfo.cellGrid[row][col]. */
      col: number;
      format: TableCellDirectFormat;
    }
  | {
      type: 'insertTableCellParagraph';
      tableId: string;
      /** 0-based row index, matching TableInfo.cellGrid[row][col]. */
      row: number;
      /** 0-based column index, matching TableInfo.cellGrid[row][col]. */
      col: number;
      /** Plain text for the new paragraph. */
      text: string;
      /** Insert before the first direct cell content block or after the last one. */
      position: 'start' | 'end';
    }
  // #190 把指定 drawing 的图片内容替换为统一占位图
  //  - 仅接受 kind='picture' 的 drawing(chart/smartArt/shape 走 deleteDrawing 或 #192)
  //  - 只改 a:blip/@r:embed,wp:inline / wp:anchor / wp:extent / 环绕 / position 一字不动
  //  - 同包内多次调用幂等:首次注入共享 word/media/placeholder.png + image rel,后续复用
  //  - 替换后若旧 media 在全包无其它引用,清理 rels 条目 + binary part(保证「原图二进制不残留」)
  | {
      type: 'replaceDrawingWithPlaceholder';
      drawingId: string;
    }
  // #190 删除 drawing 节点(B/C 类裁剪用)
  //  - 不限制 kind,任何 drawing 都可删
  //  - 删 w:drawing 本身;承载 w:r 若仅含此 drawing(无其它有效内容)一并删除
  //  - 清理孤儿 media / rels 同 replaceDrawingWithPlaceholder
  | {
      type: 'deleteDrawing';
      drawingId: string;
    }
  // #210 把 Shape / TextBox 内部文字替换为统一占位文字(#192 内容型保留依赖)
  //  - 仅接受含 w:txbxContent 的 drawing(textBox / shape-with-text)
  //  - 保留 w:drawing 整个外壳(wsp/spPr/extent/anchor/wrap/position)
  //  - 重写 w:txbxContent 内部为单段单 run 占位文字,保留首段 w:pPr/w:jc 对齐
  //  - 内部 hyperlink / 嵌套 picture 的 rels 引用按 deleteDrawing 同源清理
  | {
      type: 'replaceShapeInnerText';
      drawingId: string;
      placeholder: string;
    }
  // #14 行克隆/插入:克隆模板行插入到指定位置(支撑可变行数表回填)
  //  - templateRow: 0-based 克隆源行索引
  //  - position: insert-before 语义,传 'end' 追加到末尾
  //  - cellTexts: 可选,各列回填文字;语义见 §1.3
  //  - 克隆行内锚点一律剥离;每个 <w:p> 重分配 paraId
  | {
      type: 'insertTableRow';
      tableId: string;
      templateRow: number;
      position: number | 'end';
      cellTexts?: string[];
    }
  // #15 表格克隆/插入:把已有表深克隆一份插到指定段落之后(排版重建：模板表克隆到多处)
  //  - sourceTableId: 克隆哪张表（源表只读，不改）
  //  - anchorParaId: 克隆表插到此段之后（insert-before anchorParaId.nextSibling）
  //  - 克隆表内锚点一律剥离；每个 <w:p> 重分配 paraId（含嵌套表段落）
  | { type: 'cloneTableAfter'; anchorParaId: ParaId; sourceTableId: string }
  // task-insert-image 在指定段后插入 inline 图片新段（跨端，不碰 node fs/Buffer）
  //  - anchorParaId: 新段插到此段之后
  //  - bytes: 图片二进制（Uint8Array）
  //  - contentType: 'image/png' | 'image/jpeg' | 'image/gif' | 'image/bmp' | 'image/tiff'
  //  - widthEmu / heightEmu: wp:extent@cx/cy
  | {
      type: 'insertImageAfter';
      anchorParaId: ParaId;
      bytes: Uint8Array;
      contentType: string;
      widthEmu: number;
      heightEmu: number;
    }
  // task-20260616-table-content-driven 跨文档 verbatim 搬表（"数据表=内容块"底座）
  //  - anchorParaId: 搬入的表插到此段之后
  //  - tableXml: 源文档 <w:tbl> 子树序列化字符串（调用方在源 doc 侧 serializeToWellFormedString）
  //  - images: 格内图的预解析字节（按源 r:embed 原 rId 索引）；executor 在 host 新建 media+rel 重映射
  //  - styleTblPrXml: 可选，模板数据表 <w:tblPr> 序列化串；提供则整体替换搬入表 tblPr（样式归一）
  //  - 表内锚点（书签/批注 5 类）一律剥离；每个 <w:p> 重分配 paraId（全局唯一）
  //  - anchor 不存在 → PARA_NOT_FOUND；tableXml 非 <w:tbl> → INVALID_ARGUMENT
  //  - trackChanges=true → UNSUPPORTED_OOXML_FEATURE
  | {
      type: 'insertTableFrom';
      anchorParaId: ParaId;
      tableXml: string;
      images?: Array<{ originalRId: string; bytes: Uint8Array; contentType: string }>;
      styleTblPrXml?: string;
    }
  // task-core-table-grid 设表格 tblGrid 列宽（用于 merger 克隆模板后保真源表列宽）
  //  - tableId: 目标表格
  //  - widths: 新列宽数组（twips），长度必须等于 colCount；只改 gridCol w:w，不动 tcW
  | { type: 'setTableColumnWidths'; tableId: string; widths: number[] }
  // task-core-table-rowheight 设表格行高（用于 merger 克隆模板后保真源表行高）
  //  - tableId: 目标表格
  //  - heights: 每行行高数组（RowHeight | null），长度必须等于 rowCount
  //    RowHeight → 写/替换 trPr/trHeight；null → 删除 trHeight
  | { type: 'setTableRowHeights'; tableId: string; heights: (RowHeight | null)[] }
  // task-core-update-fields 设 settings.xml 的 updateFields（让含 TOC 域的文档打开时自动更新域）
  //  - enabled=true → 写/更新 <w:updateFields w:val="true"/>
  //  - enabled=false → 写 <w:updateFields w:val="false"/>
  | { type: 'setUpdateFields'; enabled: boolean };

// ─────────── ApplyResult(docs/03.1 §1.8)───────────

export interface ApplyResult {
  /** 插入段落产生的新 paraId,按 op 顺序 */
  insertedParagraphs: Array<{
    opIndex: number;
    anchorParaId: ParaId;
    position: 'before' | 'after';
    newParaId: ParaId;
  }>;
  /**
   * `insertParagraphAtBodyStart` 在 body 起始合成的空段 paraId，按 op 顺序。
   * 调用方据此拿到锚点 paraId 做后续重建（与 insertedParagraphs 分开，语义清晰）。
   */
  insertedBodyStartParaIds: ParaId[];
  /** 新建批注分配的 commentId,按 op 顺序 */
  createdComments: Array<{
    opIndex: number;
    commentId: CommentId;
  }>;
  /**
   * 新建书签清单。
   *
   * - 显式 `createBookmark` op 产生的:`name` 是调用方传入的;`range` 给段落 ID
   *   (ParagraphRange 时 fromParaId === toParaId,TextRange 同段所以也相等)
   * - `insertParagraphAfter` 通过 `spec.bookmark` 触发的:`name` 是按规则解析后的实际名
   *   (规则冲突会有 `_2/_3` 后缀);`range.fromParaId` = anchor,`range.toParaId` = newParaId
   */
  createdBookmarks: Array<{
    opIndex: number;
    name: BookmarkName;
    range: { fromParaId: ParaId; toParaId: ParaId };
  }>;
  acceptedRevisions: number;
  rejectedRevisions: number;
  /**
   * 实际替换的表格单元格数 (#188 / #204)。
   *
   * `replaceTableCellWithSdt` op 在 `onAnchorConflict='drop'` 模式下遇锚点会
   * **静默跳过**, 该计数只统计真正完成替换的单元格数, 调用方据此报告准确域化数。
   */
  replacedTableCells: number;
  /** Diagnostic timing, present only when the kernel collects apply timing. */
  debugTiming?: {
    ops: number;
    affectedParts: number;
    collectAffectedPartsMs?: number;
    collectStoriesMs?: number;
    collectStoryLookups?: number;
    collectStoryLookupMs?: number;
    collectByType?: Array<{
      type: Op['type'];
      count: number;
      ms: number;
      storyLookups: number;
      storyLookupMs: number;
    }>;
    transactionMs: number;
    indicesMs: number;
    totalMs: number;
    kernelMs?: number;
    byType: Array<{ type: Op['type']; count: number; ms: number }>;
  };
}

export interface PlanValidationIssue {
  code: WordErrorCode;
  message: string;
  details?: Record<string, unknown>;
}

export interface PlanValidationItem {
  opIndex: number;
  opType: Op['type'];
  status: 'valid' | 'invalid' | 'unvalidated';
  issue?: PlanValidationIssue;
}

export interface PlanValidationResult {
  valid: boolean;
  items: PlanValidationItem[];
}

/** 空结果常量(空 plan / 内部初始化用)*/
export const EMPTY_APPLY_RESULT: ApplyResult = Object.freeze({
  insertedParagraphs: [],
  insertedBodyStartParaIds: [],
  createdComments: [],
  createdBookmarks: [],
  acceptedRevisions: 0,
  rejectedRevisions: 0,
  replacedTableCells: 0,
}) as ApplyResult;

// ─────────── WordPlan 接口 ───────────

export interface WordPlan {
  replaceParagraph(
    paraId: ParaId,
    text: string,
    opts?: { onAnchorConflict?: 'error' | 'drop'; keepInlineGraphics?: boolean },
  ): this;
  replaceTextInParagraph(
    paraId: ParaId,
    search: string,
    replacement: string,
    opts?: { scope?: 'first' | 'all' },
  ): this;
  formatTextInParagraph(
    paraId: ParaId,
    search: string,
    format: {
      bold?: boolean;
      color?: string;
      italic?: boolean;
      underline?: boolean;
      highlight?: string;
      fontSize?: number;
      fontFamily?: string;
    },
    opts?: { scope?: 'first' | 'all' },
  ): this;
  formatParagraph(paraId: ParaId, format: ParagraphDirectFormat): this;

  insertParagraphBefore(anchorParaId: ParaId, spec: InsertParagraphSpec): this;
  insertParagraphAfter(anchorParaId: ParaId, spec: InsertParagraphSpec): this;

  /**
   * 无锚点在 `<w:body>` 起始插一个空段，新 paraId 通过
   * `ApplyResult.insertedBodyStartParaIds` 回传。
   *
   * 用途：body 完全无段落（极简/特殊文档，正文段被骨架删光）时，重建需要一个真实落点锚段。
   * 与 insertParagraphBefore/After 不同——后者必须有 anchor 段；本 op 不依赖任何已有段。
   */
  insertParagraphAtBodyStart(): this;

  deleteParagraph(paraId: ParaId, opts?: { onAnchorConflict?: 'error' | 'drop' }): this;

  acceptRevisions(scope: RevisionScope): this;
  rejectRevisions(scope: RevisionScope): this;

  createBookmark(spec: { name: BookmarkName; range: Range }): this;
  deleteBookmark(name: BookmarkName): this;

  addComment(spec: { range: Range; text: string; author?: string; date?: Date }): this;
  deleteComment(commentId: CommentId): this;

  /** Resolve and validate supported operation targets without mutating the document. */
  validate(): Promise<PlanValidationResult>;

  /**
   * 给表格挂批注(#186 表格分类与保留决策的"决策理由可追溯"载体)。
   * 实现上批注挂在表格首单元格的首段落,语义入口走 tableId。
   */
  addTableComment(spec: { tableId: string; text: string; author?: string; date?: Date }): this;
  /**
   * 删除整张表格(#186 不保留表格的裁剪入口)。
   * 表格内段落若含活锚点(书签/批注),按 onAnchorConflict 处理。
   */
  deleteTable(tableId: string, opts?: { onAnchorConflict?: 'error' | 'drop' }): this;

  /**
   * 把表格指定单元格首段的 runs 替换为 inline SDT (#188 表格域化)。
   *
   * - (row, col) 对应 `TableInfo.cells[row][col]` 网格坐标(tr 顺序 × tc 顺序)
   * - 段落 pPr / 单元格 tcPr / 边框 / 合并关系全部保留
   * - 原首段 runs 的首 rPr 自动透传到 SDT 内 placeholder run, 字体 / 颜色 / 字号不丢
   * - 单元格首段含锚点 (书签 / 批注) 时按 onAnchorConflict 处理 ('error' 抛 / 'drop' 跳过)
   * - 单元格首段不存在或单元格不存在时抛 PARA_NOT_FOUND
   */
  replaceTableCellWithSdt(spec: {
    tableId: string;
    row: number;
    col: number;
    sdt: SdtSpec;
    onAnchorConflict?: 'error' | 'drop';
  }): this;

  /**
   * 替换表格指定单元格首段的内容为单个纯文字 run（排版回填）。
   *
   * - (row, col) 对应 TableInfo.cells[row][col] 网格坐标
   * - **单元格 tcPr（底纹 shd / 边框 / 对齐 vAlign / 列宽 tcW）完全保留**（只动首段内部）
   * - 首段 pPr 保留；原首 run 的 rPr（字体/字号/颜色）透传到新文字 run
   * - 单元格首段含锚点（书签/批注）时按 onAnchorConflict 处理（'error' 抛 / 'drop' 跳过不替换）
   * - 单元格不存在 → TABLE_NOT_FOUND；首段不存在 → PACKAGE_CORRUPT
   * - trackChanges=true 暂不支持，抛 UNSUPPORTED_OOXML_FEATURE
   */
  replaceTableCellText(spec: {
    tableId: string;
    row: number;
    col: number;
    text: string;
    onAnchorConflict?: 'error' | 'drop';
  }): this;
  formatTableCell(spec: {
    tableId: string;
    row: number;
    col: number;
    format: TableCellDirectFormat;
  }): this;
  insertTableCellParagraph(spec: {
    tableId: string;
    row: number;
    col: number;
    text: string;
    position: 'start' | 'end';
  }): this;

  /**
   * 删除表格指定行 (#188 §7.2 行级裁剪)。
   *
   * - `row` 是 op 执行时的行索引；同一 plan 内连续删多行时调用方按**降序**传入，
   *   避免索引漂移（删 row 3 后原 row 4 变成新的 row 3）
   * - 行内段落含活锚点 (书签 / 批注) 时按 `onAnchorConflict` 处理 ('error' 抛 / 'drop' 跳过)
   * - 删除最后一行会让表格只剩 tblPr / tblGrid —— Word 允许 0 行表格但显示异常，
   *   调用方应在外层保证至少留一行（典型场景：可变行数表保留表头 + 1~N 样例行）
   */
  deleteTableRow(
    tableId: string,
    row: number,
    opts?: { onAnchorConflict?: 'error' | 'drop' },
  ): this;

  /**
   * 把 drawing 的图片内容替换为共享占位图 (#190)。
   *
   * - 仅接受 kind = 'picture' 的 drawing；其它 kind(chart / smartArt / shape / textBox /
   *   group / canvas / other)抛 UNSUPPORTED_OOXML_FEATURE,需走 deleteDrawing 或 #192
   * - **布局完全保持**:wp:inline / wp:anchor、wp:extent、wrap、position、anchorParaId
   *   都不动,只改 a:blip/@r:embed 指向共享 placeholder
   * - 同一文档内多次调用幂等:共用一份 word/media/placeholder.png + 单条 image rel
   * - 替换后若旧 media 在全包无其它引用,自动清理对应 rels 条目 + binary part
   *   (保证「原始图片二进制不出现在输出文档中」的验收条件)
   * - drawingId 不存在抛 DRAWING_NOT_FOUND
   * - trackChanges=true 暂不支持,抛 UNSUPPORTED_OOXML_FEATURE
   */
  replaceDrawingWithPlaceholder(drawingId: string): this;

  /**
   * 删除 drawing(B/C 类裁剪入口)。
   *
   * - 不限制 kind,任何 drawing 都可删
   * - 物理删 w:drawing;承载它的 w:r 若仅含该 drawing 一并删除
   * - 同样清理孤儿 media / rels(逻辑与 replaceDrawingWithPlaceholder 共享)
   * - drawingId 不存在抛 DRAWING_NOT_FOUND
   * - trackChanges=true 暂不支持,抛 UNSUPPORTED_OOXML_FEATURE
   */
  deleteDrawing(drawingId: string): this;

  /**
   * 把 Shape / TextBox 内部文字替换为统一占位文字 (#210,#192 内容型保留依赖)。
   *
   * - 仅接受含 `w:txbxContent` 的 drawing(textBox / shape-with-text)
   *   其它 kind(picture / chart / smartArt / group / canvas / shape-without-text / other)
   *   抛 UNSUPPORTED_OOXML_FEATURE,需走 replaceDrawingWithPlaceholder 或 deleteDrawing
   * - **外壳完全保留**:w:drawing 整个 wp:inline / wp:anchor / wsp:spPr / wp:extent /
   *   wrap / position 一字不动
   * - 内部 `w:txbxContent` 被重写为**单段单 run** placeholder 文本;原首段
   *   `w:pPr/w:jc` 对齐保留(沿用 innerTextAlign 维度)
   * - 内部 hyperlink / 嵌套 picture 的 rels 引用按 deleteDrawing 同源清理
   * - 占位文字走默认 run 样式(v1 不保留原首段 run rPr 字体 / 字号 / 颜色)
   * - drawingId 不存在抛 DRAWING_NOT_FOUND
   * - trackChanges=true 暂不支持,抛 UNSUPPORTED_OOXML_FEATURE
   */
  replaceShapeInnerText(drawingId: string, placeholder: string): this;

  /**
   * 克隆表格某一行作模板，在指定位置插入新行（可变行数表回填）。（#14）
   *
   * - `templateRow`：0-based，克隆此行的单元格结构（tcPr 含底纹 shd / 边框 /
   *   对齐 vAlign / 列宽 tcW + 首段 pPr + 首 run rPr 全保留）
   * - `position`：插入后新行所在的 0-based 行索引（insert-before 语义）；
   *   传 'end' 追加到末尾。范围 [0, rowCount]，rowCount 等价 'end'
   * - `cellTexts`：可选，新行各列文字（保留各列 rPr）。
   *   - 不传：克隆模板行原样（含模板原文字）
   *   - 传 L 个元素：列 i ∈ [0, min(L,C)) 替换为 cellTexts[i]；列 i ∈ [L,C) 清空；多余忽略
   * - 克隆行内锚点（书签/批注）一律剥离；行内每个 <w:p> 重新分配 paraId
   * - 表格不存在 / templateRow 越界 / position 越界 → TABLE_NOT_FOUND
   * - trackChanges=true 暂不支持 → UNSUPPORTED_OOXML_FEATURE
   */
  insertTableRow(spec: {
    tableId: string;
    templateRow: number;
    position: number | 'end';
    cellTexts?: string[];
  }): this;

  /**
   * 把已有表格深克隆一份，插入到指定段落之后（排版重建：模板表克隆到多处）。（#15）
   *
   * - `sourceTableId`：克隆哪张表。深克隆整个 `<w:tbl>`（tblPr/tblGrid/所有行/
   *   单元格 tcPr/pPr/rPr 全保留）。源表自身**只读，不被修改**。
   * - `anchorParaId`：克隆表插到此段之后（与 insertParagraphAfter 同语义：
   *   插到 anchor 的下一个兄弟之前）。anchor 应为 body story 直属段落。
   * - 克隆表内**锚点（书签/批注 5 类）一律剥离**；克隆表内**每个 `<w:p>`
   *   重新分配 paraId**，保 paraId 全局唯一。
   * - anchorParaId 不存在 → PARA_NOT_FOUND；sourceTableId 不存在 → TABLE_NOT_FOUND。
   * - trackChanges=true 暂不支持 → UNSUPPORTED_OOXML_FEATURE。
   * - v1 不在 ApplyResult 回传新表 tableId；apply 后用 `doc.tables.list()` 按文档顺序定位。
   */
  cloneTableAfter(spec: { anchorParaId: ParaId; sourceTableId: string }): this;

  /**
   * 在指定段落之后新插一个段，段内含一个 inline 图片（task-insert-image）。
   *
   * - `anchorParaId`：新段插到此段之后（与 insertParagraphAfter 同语义）
   * - `bytes`：图片二进制（`Uint8Array`，跨端，绝不碰 node fs/Buffer）
   * - `contentType`：'image/png' | 'image/jpeg' | 'image/gif' | 'image/bmp' | 'image/tiff'
   * - `widthEmu` / `heightEmu`：wp:extent@cx/cy（EMU 单位，1 英寸 = 914400 EMU）
   * - anchor 不存在 → `PARA_NOT_FOUND`
   * - 不支持的 contentType → `UNSUPPORTED_OOXML_FEATURE`
   * - trackChanges=true → `UNSUPPORTED_OOXML_FEATURE`
   */
  insertImageAfter(
    anchorParaId: ParaId,
    spec: {
      bytes: Uint8Array;
      contentType: string;
      widthEmu: number;
      heightEmu: number;
    },
  ): this;

  /**
   * 跨文档 verbatim 搬一张表进当前文档，插到 anchorParaId 段之后
   * （task-20260616-table-content-driven，"数据表 = 内容块" §13.9）。
   *
   * - `anchorParaId`：搬入的表插到此段之后（与 cloneTableAfter 同语义）。
   * - `tableXml`：源文档 `<w:tbl>` 子树序列化字符串。调用方（typeset，持源 doc）用
   *   `serializeToWellFormedString(tblElement)` 取得；命名空间随子树带入，无需额外声明。
   * - `images`：格内图预解析字节，按源 `<a:blip @r:embed>` 的**原始 rId** 索引；executor 在
   *   host 新建 media part + image rel + content-type，把 embed 改写成新 rId。缺某 rId → 该
   *   blip embed 清掉 + warn（结构/文字仍保真，仅该格图降级）。
   * - `styleTblPrXml`：可选，模板数据表 `<w:tblPr>` 序列化串。提供则整体替换搬入表的 tblPr
   *   （tblStyle/边框/tblLook/宽度 → 归一到模板样式，覆盖内容表自带样式）。
   * - 表内锚点（书签/批注 5 类）一律剥离；每个 `<w:p>` 重分配 paraId（全局唯一）。
   * - 行/列/值/合并/嵌套表/格内图全部 verbatim 来自 tableXml，executor 不重建/不裁剪/不对位。
   * - anchor 不存在 → `PARA_NOT_FOUND`；tableXml 非 `<w:tbl>` → `INVALID_ARGUMENT`；
   *   trackChanges=true → `UNSUPPORTED_OOXML_FEATURE`。
   * - v1 不在 ApplyResult 回传新表 tableId；apply 后用 `doc.tables.list()` 按文档顺序定位。
   */
  insertTableFrom(spec: {
    anchorParaId: ParaId;
    tableXml: string;
    images?: Array<{ originalRId: string; bytes: Uint8Array; contentType: string }>;
    styleTblPrXml?: string;
  }): this;

  /**
   * 设表格列宽：同步更新 `<w:tblGrid>` gridCol w:w **和** 每个单元格 `<w:tcW>`
   * （task-core-table-grid：merger 保真源表列宽）。
   *
   * - `tableId`：目标表格（来自 `tables.list()` 的 tableId）。
   * - `widths`：新列宽数组（twips，1 twip = 1/20 磅）。长度必须等于该表格的
   *   `colCount`（gridCol 数量）——不等则抛 `INVALID_ARGUMENT`（防止截断/越界）。
   * - 同步设每个单元格 `<w:tcW w:type="dxa">`（Word 渲染优先 tcW over gridCol）：
   *   合并单元格（gridSpan > 1）的 tcW = 跨列宽之和；vMerge 续行单元格正常处理。
   * - tableId 不存在 → `TABLE_NOT_FOUND`。
   * - trackChanges=true → `UNSUPPORTED_OOXML_FEATURE`（暂不支持）。
   */
  setTableColumnWidths(spec: { tableId: string; widths: number[] }): this;

  /**
   * 设表格行高（task-core-table-rowheight：merger 保真源表行高）。
   *
   * - `tableId`：目标表格（来自 `tables.list()` 的 tableId）。
   * - `heights`：每行行高数组，长度必须等于该表格的 `rowCount`——不等则抛 `INVALID_ARGUMENT`。
   *   - `RowHeight`：在该行 `<w:trPr>` 写入/替换 `<w:trHeight w:val="N" w:hRule="..."/>`。
   *   - `null`：删除该行的 `<w:trHeight>`（即「不指定行高」）。
   * - `heights[i]` 套到物理第 i 行 `<w:tr>`（与 `TableInfo.rowHeights` 1:1）。
   * - tableId 不存在 → `TABLE_NOT_FOUND`。
   * - trackChanges=true → `UNSUPPORTED_OOXML_FEATURE`（暂不支持）。
   */
  setTableRowHeights(spec: { tableId: string; heights: (RowHeight | null)[] }): this;

  /**
   * 在 settings.xml 写入/更新 `<w:updateFields>`，让含 TOC 域的文档打开时自动更新域。
   * （task-core-update-fields）
   *
   * - `enabled=true`：在 settings.xml 根 `<w:settings>` 写入/更新
   *   `<w:updateFields w:val="true"/>`——Word/WPS 打开文档时自动重生成所有域（TOC/PAGEREF/SEQ 等）。
   * - `enabled=false`：写 `<w:updateFields w:val="false"/>`（显式禁用，语义更清晰）。
   * - 幂等：settings.xml 已有 `<w:updateFields>` 则改 `w:val`，不重复插入元素。
   * - settings.xml 不存在 → 抛 `UNSUPPORTED_OOXML_FEATURE`（后续跟进新建 settings part）。
   */
  setUpdateFields(enabled: boolean): this;

  apply(): Promise<ApplyResult>;

  /** 只读 op 队列,用于调试 / 日志 / 测试 */
  readonly ops: ReadonlyArray<Op>;
  /** apply 是否已被调用过(成功与否都置 true,防止二次 apply)*/
  readonly applied: boolean;
  /** Plan 创建时的选项,只读 */
  readonly options: Readonly<
    Required<Pick<PlanOptions, 'trackChanges'>> & Omit<PlanOptions, 'trackChanges'>
  >;
}

// ─────────── 实现 ───────────

export interface WordPlanImplInit {
  docId: string;
  transport: Transport;
  options: PlanOptions;
}

export class WordPlanImpl implements WordPlan {
  private readonly docId: string;
  private readonly transport: Transport;
  private readonly _options: Readonly<WordPlan['options']>;
  private readonly _ops: Op[] = [];
  private _applied = false;

  constructor(init: WordPlanImplInit) {
    this.docId = init.docId;
    this.transport = init.transport;
    // normalize:trackChanges 缺省 false;author / date 保留原值(可能 undefined)
    const normalized: { trackChanges: boolean; author?: string; date?: Date } = {
      trackChanges: init.options.trackChanges ?? false,
    };
    if (init.options.author !== undefined) normalized.author = init.options.author;
    if (init.options.date !== undefined) normalized.date = init.options.date;
    this._options = Object.freeze(normalized);
  }

  // ─────────── 状态 accessor ───────────

  get ops(): ReadonlyArray<Op> {
    return this._ops;
  }

  get applied(): boolean {
    return this._applied;
  }

  get options(): Readonly<WordPlan['options']> {
    return this._options;
  }

  // ─────────── 段落修改(#28)───────────

  replaceParagraph(
    paraId: ParaId,
    text: string,
    opts?: { onAnchorConflict?: 'error' | 'drop'; keepInlineGraphics?: boolean },
  ): this {
    this.ensureOpen('replaceParagraph');
    this._ops.push({
      type: 'replaceParagraph',
      paraId,
      text,
      onAnchorConflict: opts?.onAnchorConflict ?? 'error',
      keepInlineGraphics: opts?.keepInlineGraphics,
    });
    return this;
  }

  replaceTextInParagraph(
    paraId: ParaId,
    search: string,
    replacement: string,
    opts?: { scope?: 'first' | 'all' },
  ): this {
    this.ensureOpen('replaceTextInParagraph');
    this._ops.push({
      type: 'replaceTextInParagraph',
      paraId,
      search,
      replacement,
      scope: opts?.scope ?? 'first',
    });
    return this;
  }

  formatTextInParagraph(
    paraId: ParaId,
    search: string,
    format: {
      bold?: boolean;
      color?: string;
      italic?: boolean;
      underline?: boolean;
      highlight?: string;
      fontSize?: number;
      fontFamily?: string;
    },
    opts?: { scope?: 'first' | 'all' },
  ): this {
    this.ensureOpen('formatTextInParagraph');
    this._ops.push({
      type: 'formatTextInParagraph',
      paraId,
      search,
      format,
      scope: opts?.scope ?? 'first',
    });
    return this;
  }

  formatParagraph(paraId: ParaId, format: ParagraphDirectFormat): this {
    this.ensureOpen('formatParagraph');
    this._ops.push({
      type: 'formatParagraph',
      paraId,
      format: { ...format },
    });
    return this;
  }

  // ─────────── 段落插入 / 删除(#29)───────────

  insertParagraphBefore(anchorParaId: ParaId, spec: InsertParagraphSpec): this {
    this.ensureOpen('insertParagraphBefore');
    this._ops.push({ type: 'insertParagraphBefore', anchorParaId, spec });
    return this;
  }

  insertParagraphAfter(anchorParaId: ParaId, spec: InsertParagraphSpec): this {
    this.ensureOpen('insertParagraphAfter');
    this._ops.push({ type: 'insertParagraphAfter', anchorParaId, spec });
    return this;
  }

  insertParagraphAtBodyStart(): this {
    this.ensureOpen('insertParagraphAtBodyStart');
    this._ops.push({ type: 'insertParagraphAtBodyStart' });
    return this;
  }

  deleteParagraph(paraId: ParaId, opts?: { onAnchorConflict?: 'error' | 'drop' }): this {
    this.ensureOpen('deleteParagraph');
    this._ops.push({
      type: 'deleteParagraph',
      paraId,
      onAnchorConflict: opts?.onAnchorConflict ?? 'error',
    });
    return this;
  }

  // ─────────── 修订(#31 / #32)───────────

  acceptRevisions(scope: RevisionScope): this {
    this.ensureOpen('acceptRevisions');
    this._ops.push({ type: 'acceptRevisions', scope });
    return this;
  }

  rejectRevisions(scope: RevisionScope): this {
    this.ensureOpen('rejectRevisions');
    this._ops.push({ type: 'rejectRevisions', scope });
    return this;
  }

  // ─────────── 书签(#34)───────────

  createBookmark(spec: { name: BookmarkName; range: Range }): this {
    this.ensureOpen('createBookmark');
    this._ops.push({ type: 'createBookmark', name: spec.name, range: spec.range });
    return this;
  }

  deleteBookmark(name: BookmarkName): this {
    this.ensureOpen('deleteBookmark');
    this._ops.push({ type: 'deleteBookmark', name });
    return this;
  }

  // ─────────── 批注(#35)───────────

  addComment(spec: { range: Range; text: string; author?: string; date?: Date }): this {
    this.ensureOpen('addComment');
    const op: Extract<Op, { type: 'addComment' }> = {
      type: 'addComment',
      range: spec.range,
      text: spec.text,
    };
    if (spec.author !== undefined) op.author = spec.author;
    if (spec.date !== undefined) op.date = spec.date.toISOString();
    this._ops.push(op);
    return this;
  }

  deleteComment(commentId: CommentId): this {
    this.ensureOpen('deleteComment');
    this._ops.push({ type: 'deleteComment', commentId });
    return this;
  }

  // ─────────── 表格(#186)───────────

  addTableComment(spec: { tableId: string; text: string; author?: string; date?: Date }): this {
    this.ensureOpen('addTableComment');
    const op: Extract<Op, { type: 'addTableComment' }> = {
      type: 'addTableComment',
      tableId: spec.tableId,
      text: spec.text,
    };
    if (spec.author !== undefined) op.author = spec.author;
    if (spec.date !== undefined) op.date = spec.date.toISOString();
    this._ops.push(op);
    return this;
  }

  deleteTable(tableId: string, opts?: { onAnchorConflict?: 'error' | 'drop' }): this {
    this.ensureOpen('deleteTable');
    this._ops.push({
      type: 'deleteTable',
      tableId,
      onAnchorConflict: opts?.onAnchorConflict ?? 'error',
    });
    return this;
  }

  // ─────────── 表格域化(#188)───────────

  replaceTableCellWithSdt(spec: {
    tableId: string;
    row: number;
    col: number;
    sdt: SdtSpec;
    onAnchorConflict?: 'error' | 'drop';
  }): this {
    this.ensureOpen('replaceTableCellWithSdt');
    this._ops.push({
      type: 'replaceTableCellWithSdt',
      tableId: spec.tableId,
      row: spec.row,
      col: spec.col,
      sdt: spec.sdt,
      onAnchorConflict: spec.onAnchorConflict ?? 'error',
    });
    return this;
  }

  deleteTableRow(
    tableId: string,
    row: number,
    opts?: { onAnchorConflict?: 'error' | 'drop' },
  ): this {
    this.ensureOpen('deleteTableRow');
    this._ops.push({
      type: 'deleteTableRow',
      tableId,
      row,
      onAnchorConflict: opts?.onAnchorConflict ?? 'error',
    });
    return this;
  }

  replaceTableCellText(spec: {
    tableId: string;
    row: number;
    col: number;
    text: string;
    onAnchorConflict?: 'error' | 'drop';
  }): this {
    this.ensureOpen('replaceTableCellText');
    this._ops.push({
      type: 'replaceTableCellText',
      tableId: spec.tableId,
      row: spec.row,
      col: spec.col,
      text: spec.text,
      onAnchorConflict: spec.onAnchorConflict ?? 'error',
    });
    return this;
  }

  formatTableCell(spec: {
    tableId: string;
    row: number;
    col: number;
    format: TableCellDirectFormat;
  }): this {
    this.ensureOpen('formatTableCell');
    this._ops.push({
      type: 'formatTableCell',
      tableId: spec.tableId,
      row: spec.row,
      col: spec.col,
      format: spec.format,
    });
    return this;
  }

  // ─────────── 图片/Drawing(#190)───────────

  insertTableCellParagraph(spec: {
    tableId: string;
    row: number;
    col: number;
    text: string;
    position: 'start' | 'end';
  }): this {
    this.ensureOpen('insertTableCellParagraph');
    this._ops.push({
      type: 'insertTableCellParagraph',
      tableId: spec.tableId,
      row: spec.row,
      col: spec.col,
      text: spec.text,
      position: spec.position,
    });
    return this;
  }

  replaceDrawingWithPlaceholder(drawingId: string): this {
    this.ensureOpen('replaceDrawingWithPlaceholder');
    this._ops.push({ type: 'replaceDrawingWithPlaceholder', drawingId });
    return this;
  }

  deleteDrawing(drawingId: string): this {
    this.ensureOpen('deleteDrawing');
    this._ops.push({ type: 'deleteDrawing', drawingId });
    return this;
  }

  replaceShapeInnerText(drawingId: string, placeholder: string): this {
    this.ensureOpen('replaceShapeInnerText');
    this._ops.push({ type: 'replaceShapeInnerText', drawingId, placeholder });
    return this;
  }

  // ─────────── 表格行克隆/插入(#14)───────────

  insertTableRow(spec: {
    tableId: string;
    templateRow: number;
    position: number | 'end';
    cellTexts?: string[];
  }): this {
    this.ensureOpen('insertTableRow');
    const op: Extract<Op, { type: 'insertTableRow' }> = {
      type: 'insertTableRow',
      tableId: spec.tableId,
      templateRow: spec.templateRow,
      position: spec.position,
    };
    if (spec.cellTexts !== undefined) op.cellTexts = spec.cellTexts;
    this._ops.push(op);
    return this;
  }

  // ─────────── 表格克隆/插入(#15)───────────

  cloneTableAfter(spec: { anchorParaId: ParaId; sourceTableId: string }): this {
    this.ensureOpen('cloneTableAfter');
    this._ops.push({
      type: 'cloneTableAfter',
      anchorParaId: spec.anchorParaId,
      sourceTableId: spec.sourceTableId,
    });
    return this;
  }

  // ─────────── 图片插入（task-insert-image）───────────

  insertImageAfter(
    anchorParaId: ParaId,
    spec: {
      bytes: Uint8Array;
      contentType: string;
      widthEmu: number;
      heightEmu: number;
    },
  ): this {
    this.ensureOpen('insertImageAfter');
    this._ops.push({
      type: 'insertImageAfter',
      anchorParaId,
      bytes: spec.bytes,
      contentType: spec.contentType,
      widthEmu: spec.widthEmu,
      heightEmu: spec.heightEmu,
    });
    return this;
  }

  // ─────────── 跨文档搬表（task-20260616-table-content-driven）───────────

  insertTableFrom(spec: {
    anchorParaId: ParaId;
    tableXml: string;
    images?: Array<{ originalRId: string; bytes: Uint8Array; contentType: string }>;
    styleTblPrXml?: string;
  }): this {
    this.ensureOpen('insertTableFrom');
    const op: Extract<Op, { type: 'insertTableFrom' }> = {
      type: 'insertTableFrom',
      anchorParaId: spec.anchorParaId,
      tableXml: spec.tableXml,
    };
    if (spec.images !== undefined) op.images = spec.images;
    if (spec.styleTblPrXml !== undefined) op.styleTblPrXml = spec.styleTblPrXml;
    this._ops.push(op);
    return this;
  }

  // ─────────── 表格列宽（task-core-table-grid）───────────

  setTableColumnWidths(spec: { tableId: string; widths: number[] }): this {
    this.ensureOpen('setTableColumnWidths');
    this._ops.push({
      type: 'setTableColumnWidths',
      tableId: spec.tableId,
      widths: spec.widths,
    });
    return this;
  }

  // ─────────── 表格行高（task-core-table-rowheight）───────────

  setTableRowHeights(spec: { tableId: string; heights: (RowHeight | null)[] }): this {
    this.ensureOpen('setTableRowHeights');
    this._ops.push({
      type: 'setTableRowHeights',
      tableId: spec.tableId,
      heights: spec.heights,
    });
    return this;
  }

  // ─────────── 文档域更新设置（task-core-update-fields）───────────

  setUpdateFields(enabled: boolean): this {
    this.ensureOpen('setUpdateFields');
    this._ops.push({ type: 'setUpdateFields', enabled });
    return this;
  }

  // ─────────── apply 状态机 ───────────

  async apply(): Promise<ApplyResult> {
    if (this._applied) {
      throw new WordError(
        'PLAN_ALREADY_APPLIED',
        'WordPlan.apply() called on a plan that was already applied. Create a new plan via doc.plan().',
        this.docId,
      );
    }
    // 进入 apply 立即锁定,无论成败 —— plan 是一次性对象,失败场景由调用方重建
    this._applied = true;

    if (this._ops.length === 0) {
      // 契约 §4.1:PLAN_EMPTY 默认 warn 不抛(严格模式留未来引入)
      console.warn(
        `[word-plan] apply() called on empty plan (docId=${this.docId}); returning empty ApplyResult`,
      );
      return cloneEmptyResult();
    }

    // 走 RPC 到 kernel.applyPlan;序列化参数为纯 JSON 形态(Date 已在 addComment
    // 里转 ISO 字符串,其它字段均为 primitive / plain object)。
    return this.transport.call<ApplyResult>('applyPlan', [
      {
        ops: this._ops,
        options: {
          trackChanges: this._options.trackChanges ?? false,
          ...(this._options.author !== undefined ? { author: this._options.author } : {}),
          ...(this._options.date !== undefined ? { date: this._options.date.toISOString() } : {}),
        },
      },
    ]);
  }

  async validate(): Promise<PlanValidationResult> {
    this.ensureOpen('validate');
    return this.transport.call<PlanValidationResult>('validatePlan', [
      {
        ops: this._ops,
        options: {
          trackChanges: this._options.trackChanges ?? false,
          ...(this._options.author !== undefined ? { author: this._options.author } : {}),
          ...(this._options.date !== undefined ? { date: this._options.date.toISOString() } : {}),
        },
      },
    ]);
  }

  // ─────────── 内部 ───────────

  private ensureOpen(methodName: string): void {
    if (this._applied) {
      throw new WordError(
        'PLAN_ALREADY_APPLIED',
        `WordPlan.${methodName}() called after apply(); plans are single-use`,
        this.docId,
      );
    }
  }
}

function cloneEmptyResult(): ApplyResult {
  // 不暴露冻结引用,避免调用方误把 ApplyResult 当只读产物(例如想追加字段调试)
  return {
    insertedParagraphs: [],
    insertedBodyStartParaIds: [],
    createdComments: [],
    createdBookmarks: [],
    acceptedRevisions: 0,
    rejectedRevisions: 0,
    replacedTableCells: 0,
  };
}

// ─────────── kernel 侧载荷(applyPlan RPC 参数)───────────

/**
 * `kernel.applyPlan` 接收的结构。`date` 是 ISO 8601 字符串(序列化友好),kernel
 * 端需要时用 `new Date(str)` 解回。
 */
export interface ApplyPlanParams {
  ops: Op[];
  options: {
    trackChanges: boolean;
    author?: string;
    date?: string;
  };
}

// story 引用的 placeholder;供 executor PR 未来使用时若要引 StoryId 避免循环
// import 可直接 import from this module.
export type { StoryId };
