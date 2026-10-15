/**
 * listParagraphs —— 段落查询纯工具(docs/03 §9 / docs/03.1 §1.6,#24)。
 *
 * 职责:按 scope / filter / textFormat / pagination 返回 ParagraphInfo 列表。
 *
 * 架构:**纯函数**,拿 pkg 给结果,不知道文档尺寸、不做智能默认。"默认分页"
 * 是 `{ mode: 'byCount', pageSize: 25, pageIndex: 0 }`,要全量请显式传
 * `{ pagination: { mode: 'all' } }`(2026-04-22 从早期"≤500 段自动 all"的智能
 * 默认撤回;见 docs/03 §9.2 的历史说明)。
 *
 * 实现约束:
 *   - 段落枚举走 `getStoryParagraphs`(基于 `getElementsByTagNameNS`),**不**用
 *     fontoxpath。10 万段落上 `//w:p` 慢 460× —— CI 的 `pnpm check:xpath` 会在
 *     源码里扫到此类无谓词 XPath 即失败。
 *   - `filter.contains` 是 O(N) 子串匹配,大档慢(docs 已标注)。
 *   - `hasComments` / `hasRevisions` / `hasBookmarks` 用 `el.getElementsByTagNameNS(...).length > 0`
 *     检测段内是否存在对应标记元素,不递归手遍。
 */
import type { Package } from '../pack/index.js';
import type { PackageIndices } from './indices.js';
import { buildTableIndex } from './list-tables.js';
import { MATH_NS, W14_NS, W_NS } from './namespaces.js';
import { getParagraphText, type RunTextOptions } from './runs.js';
import { findStoryFor, getStoryParagraphs, resolveScope, type StoryId } from './stories.js';

/** list scope —— 'body' / 'all' 或显式 StoryId 列表 */
export type ListParagraphsScope = 'body' | 'all' | StoryId[];

/** text 返回格式。'preview' = 首 100 字 */
export type ListParagraphsTextFormat = 'full' | 'preview';

export interface ListParagraphsFilter {
  /** 等值匹配 `<w:pPr>/<w:pStyle w:val>` */
  styleId?: string;
  /** **O(N) 子串匹配**,区分大小写。大档谨慎使用。 */
  contains?: string;
  /** 段内是否含 comment 标记(commentRangeStart / commentReference / commentRangeEnd) */
  hasComments?: boolean;
  /** 段内是否含 revision 标记(Tier 1 只看 w:ins / w:del,不含 rPrChange/pPrChange) */
  hasRevisions?: boolean;
  /**
   * 排除位于表格单元格（`<w:tc>`）内的段落。
   *
   * 表格内容由独立的 tables phase 处理，段落 phase 只关心主文本流段落。
   * 开启后会对每个段落向上遍历祖先链（深度通常 ≤5），找到 `<w:tc>` 即排除。
   */
  excludeTableCells?: boolean;
}

export type ListParagraphsPagination =
  | { mode: 'all' }
  | { mode: 'byCount'; pageSize: number; pageIndex: number }
  | { mode: 'byChar'; charPerPage: number; pageIndex: number };

/**
 * 段落直接格式属性（#182）。
 *
 * 仅包含在 `<w:pPr>` 上**直接设置**的值，不做样式继承计算。
 * 消费方（如骨架化引擎）可将这些值与 styleId 组合为聚类 key，
 * 区分同一样式但有手工格式覆盖的段落。
 *
 * 所有字段均为可选——未直接设置的属性不出现在结果中（继承自样式或默认值）。
 */
export interface ParagraphFormat {
  /**
   * 水平对齐（`w:pPr/w:jc@w:val`）。
   * `'both'` 对应 OOXML 的 "both"（两端对齐）。
   */
  alignment?: 'left' | 'center' | 'right' | 'both' | 'distribute';
  /**
   * 缩进（`w:pPr/w:ind`），单位 twips（1/20 磅）。
   * `firstLine` 正值为首行缩进，负值为悬挂缩进（来自 `w:hanging`）。
   */
  indent?: {
    left?: number;
    right?: number;
    /** 首行缩进（正）或悬挂缩进（负，源自 `w:hanging`） */
    firstLine?: number;
  };
  /**
   * 段间距（`w:pPr/w:spacing`），单位 twips（line 例外：受 lineRule 影响）。
   */
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: 'auto' | 'exact' | 'atLeast';
  };
  /**
   * 段落级默认字体大小（`w:pPr/w:rPr/w:sz@w:val ÷ 2`），单位磅。
   * 仅当段落属性中显式设置了 `w:rPr/w:sz` 时有值；继承自样式的字号不在此字段。
   */
  fontSize?: number;
  /**
   * 段落级默认粗体（`w:pPr/w:rPr/w:b`）。
   * `true`  = 显式开启；`false` = 显式关闭（`w:val="0"`）；
   * 字段不存在 = 未直接设置（继承自样式）。
   */
  bold?: boolean;
  /**
   * 大纲级别（`w:pPr/w:outlineLvl@w:val`），0–8 对应标题 1–9，9 表示正文。
   * WPS / Word 的自定义标题样式通常在此字段直接编码级别，即使 basedOn 是 Normal。
   * 字段不存在 = 未在段落级直接设置（继承自样式或默认值）。
   */
  outlineLevel?: number;
  /**
   * 段落是否直接挂有编号属性（`w:pPr/w:numPr` 存在且含 `w:numId@w:val != "0"`）。
   * 用于列表检测：有编号属性的段落通常是列表项。
   */
  hasNumbering?: boolean;
  /**
   * 段落是否强制从新页起始（`w:pPr/w:pageBreakBefore`，#211）。
   * true = 渲染该段时从新页开始（封面段的强信号之一）。
   * 字段不存在 = 未在段落级直接设置（继承自样式或默认 false）。
   */
  pageBreakBefore?: boolean;
  /**
   * 段落是否承载节属性（`w:pPr/w:sectPr` 存在，#211）。
   * true = 该段是其所属节的最后一段，节边界在此结束。
   * 与 doc.sections.list() 的 lastParaId 字段对应。
   */
  hasSectionBreak?: boolean;
  /**
   * 段落内是否含手工分页符（`w:r/w:br[@w:type="page"]`，#211）。
   * 与 pageBreakBefore 互补 —— 后者作用于段首，前者可出现在段中/段尾。
   * 字段不存在 = 段内无手工分页符。
   */
  hasManualPageBreak?: boolean;
  /**
   * 段落级默认字体颜色（`w:pPr/w:rPr/w:color@w:val`，task-20260615-redhead-fingerprint）。
   * 段标 run 颜色，作用于无独立 rPr 的 run。值为 6 位十六进制 `RRGGBB`（如 `FF0000`）或 `auto`。
   * 仅当段落属性中显式设置了 `w:pPr/w:rPr/w:color` 时有值；继承自样式的颜色不在此字段。
   *
   * 用途（骨架化 / 排版的视觉簇判据）：判定"封面式文头"等需区分色异段（如红头）。
   * 注意：行内首 run 的颜色单独由 {@link firstRunColor} 提供——很多文档把红色编码在 run 级
   * 而非段标级，两者互补使用。
   */
  color?: string;
  /**
   * 段落**首个**直属 run（`w:r/w:rPr/w:color@w:val`）的字体颜色（task-20260615-redhead-fingerprint）。
   * 值为 6 位十六进制 `RRGGBB` 或 `auto`。仅当首个直属 run 显式设了 `w:color` 时有值。
   *
   * 与 {@link color}（段标级）互补：红头/标题等视觉显著段的颜色常落在 run 级 rPr 而非段标 rPr，
   * 需读首 run 才能拿到。取**首个直属 run**（非表格/文本框下钻），与 getParagraphText 的直属 run
   * 语义对齐；段内多 run 颜色不一致时只反映首 run（满足"段首视觉身份"判据，不做整段聚合）。
   */
  firstRunColor?: string;
}

export interface ListParagraphsOptions {
  /** 默认 'body' */
  scope?: ListParagraphsScope;
  filter?: ListParagraphsFilter;
  /** 默认 'full' */
  textFormat?: ListParagraphsTextFormat;
  /** 默认 `{ mode: 'byCount', pageSize: 25, pageIndex: 0 }`(不再有"智能默认") */
  pagination?: ListParagraphsPagination;
  /**
   * 是否在结果中包含段落直接格式属性（`ParagraphFormat`）。
   * 默认 `false`。开启后每个 `ParagraphInfo.format` 会携带从 `<w:pPr>` 读取的直接设置值。
   * 仅包含直接设置的属性，不做样式继承计算。
   */
  includeFormat?: boolean;
  /**
   * 启用结构空白保真读（task-20260623-core-tab-faithful-read）。
   * `true` 时 `ParagraphInfo.text` 含 `\t`（制表符）和 `\n`（软换行），
   * 是 `appendSegmentsToRun` 的精确逆运算。
   * 默认 `false`：只拼 `<w:t>`，偏移算术安全（与历史行为完全一致）。
   */
  structuralWhitespace?: boolean;
  /**
   * 是否在结果中包含段落所在表格的坐标信息（task-20260625-para-table-info）。
   * 默认 `false`。
   *
   * 开启后，位于表格单元格内的段落会携带 `ParagraphInfo.table`（tableId / row / col）。
   * 表格外段落的 `table` 字段不出现。
   *
   * 与 `includeFormat` 是独立开关：表格归属是结构信息，不是格式。
   * 开启时会对每个段落做一次祖先链遍历（深度 ≤ 5），
   * 并在首次调用时构建 `<w:tbl> Element → tableId` 反查映射（O(表格数)）。
   */
  includeTableInfo?: boolean;
}

export interface ParagraphInfo {
  paraId: string;
  story: StoryId;
  /** textFormat='preview' 时为首 100 字 */
  text: string;
  /** preview 模式被截断时为 true;full 模式/未截断不给 */
  truncated?: boolean;
  /** `<w:pStyle w:val>` 未设则缺 */
  styleId?: string;
  hasRevisions?: boolean;
  hasComments?: boolean;
  hasBookmarks?: boolean;
  /**
   * 段落子树是否含 OMML 公式对象（`<m:oMath>`，math namespace）。
   * Word「插入→公式」生成的公式对象。纯文字写的"公式"（如 "F=ma"）不算。
   * 含则为 true；不含时不给（与 hasComments 等同风格）。
   */
  hasMath?: boolean;
  /**
   * 段落直接格式属性。仅在 `listParagraphs` 传入 `includeFormat: true` 时存在。
   * 不包含从样式继承的属性值。
   */
  format?: ParagraphFormat;
  /**
   * 段落所在表格坐标。仅当 `includeTableInfo=true` 且段落位于表格单元格内时存在。
   *
   * - `tableId`：与 `doc.tables.list()` / `getTable()` 使用的 id 完全一致（装载期
   *   `buildIndices` 分配的 `tbl-NNNN`）。
   * - `row`：所在 `<w:tr>` 在该表中的 0 基行号。
   * - `col`：所在 `<w:tc>` 在该行中的 0 基列号（物理序号）。
   *
   * 注：`col` 是物理 `<w:tc>` 序号，**不**做 gridSpan / vMerge 的逻辑列号校正——
   * 含合并单元格的表格中，物理序号与逻辑列号可能不同。v1 取物理序号已能满足
   * "区分表格内/外编号序列"的主场景；如需精确逻辑列号，可在 v2 补充。
   *
   * 嵌套表格时报**最内层**的表格。
   */
  table?: { tableId: string; row: number; col: number };
}

export interface ListResult {
  paragraphs: ParagraphInfo[];
  pagination: {
    mode: 'all' | 'byCount' | 'byChar';
    pageIndex: number;
    totalPages: number;
    totalParagraphs: number;
    hasNext: boolean;
    /**
     * 保留字段,当前实现恒不填(#24 把"≤500 段自动 all"的智能默认撤回,改为恒定
     * byCount(25,0)默认;docs/03 §9.2)。调用方不应依赖。
     */
    autoPaginated?: boolean;
  };
}

/** 默认分页:首页、25 条 */
export const DEFAULT_PAGINATION: ListParagraphsPagination = {
  mode: 'byCount',
  pageSize: 25,
  pageIndex: 0,
};

/** preview 截断字数 */
export const PREVIEW_CHAR_LIMIT = 100;

export function listParagraphs(pkg: Package, opts: ListParagraphsOptions = {}): ListResult {
  // 1. scope → stories(resolveScope 在 stories.ts,#26 起被 listRevisions 共用)
  const stories = resolveScope(pkg, opts.scope ?? 'body');

  // 2. flatten (story, paragraphEl)
  interface Entry {
    story: StoryId;
    el: Element;
  }
  const all: Entry[] = [];
  for (const story of stories) {
    for (const el of getStoryParagraphs(story)) {
      all.push({ story: story.id, el });
    }
  }

  // 3. filter(在分页前,totalParagraphs 取过滤后)
  const filtered = opts.filter ? all.filter((e) => matchesFilter(e.el, opts.filter!)) : all;

  // 4. 分页
  const pagination = opts.pagination ?? DEFAULT_PAGINATION;
  const { slice, meta } = paginate(filtered, pagination);

  // 5. 映射为 ParagraphInfo(textFormat 在此应用 preview 截断)
  const textFormat = opts.textFormat ?? 'full';
  const includeFormat = opts.includeFormat ?? false;
  const runTextOpts: RunTextOptions | undefined = opts.structuralWhitespace
    ? { structuralWhitespace: true }
    : undefined;

  // includeTableInfo=true 时构建 tblElement → tableId 反查 Map（O(表格数)，只做一次）
  // buildTableIndex 返回 Map<tableId, tblEl>，反转后可从 Element 查 tableId
  let tblReverseIdx: Map<Element, string> | undefined;
  if (opts.includeTableInfo) {
    const tblIdx = buildTableIndex(pkg.parts.values());
    tblReverseIdx = new Map<Element, string>();
    for (const [id, el] of tblIdx) {
      tblReverseIdx.set(el, id);
    }
  }

  const paragraphs = slice.map((e) =>
    buildInfo(e.story, e.el, textFormat, includeFormat, runTextOpts, tblReverseIdx),
  );

  return { paragraphs, pagination: meta };
}

// ─────────── filter ───────────

function matchesFilter(p: Element, f: ListParagraphsFilter): boolean {
  // excludeTableCells 放最前：祖先遍历 O(depth≤5)，比文本扫描快，能快速排除大量表格段落
  if (f.excludeTableCells) {
    if (isInTableCell(p)) return false;
  }
  if (f.styleId !== undefined) {
    if (getStyleId(p) !== f.styleId) return false;
  }
  if (f.hasComments !== undefined) {
    if (detectHasComments(p) !== f.hasComments) return false;
  }
  if (f.hasRevisions !== undefined) {
    if (detectHasRevisions(p) !== f.hasRevisions) return false;
  }
  // contains 放最后:O(N) 全文扫,前面的廉价过滤已能淘汰大部分段落
  if (f.contains !== undefined) {
    if (!getParagraphText(p).includes(f.contains)) return false;
  }
  return true;
}

// ─────────── 属性检测 ───────────

function getStyleId(p: Element): string | undefined {
  // <w:p><w:pPr><w:pStyle w:val="xxx"/></w:pPr>
  const pPr = firstDirectChild(p, W_NS, 'pPr');
  if (!pPr) return undefined;
  const pStyle = firstDirectChild(pPr, W_NS, 'pStyle');
  if (!pStyle) return undefined;
  const val = pStyle.getAttributeNS(W_NS, 'val');
  return val ?? undefined;
}

function firstDirectChild(el: Element, ns: string, local: string): Element | null {
  const children = el.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === ns &&
      (c as Element).localName === local
    ) {
      return c as Element;
    }
  }
  return null;
}

/**
 * 判断 `<w:p>` 是否位于表格单元格（`<w:tc>`）内。
 *
 * 向上遍历祖先链：遇到 `<w:tc>` 返回 true，遇到 `<w:body>` 或到达根节点返回 false。
 * 典型深度：body > tbl > tr > tc > p = 4 跳，代价可忽略。
 */
function isInTableCell(p: Element): boolean {
  let el: Element | null = p.parentElement;
  while (el) {
    if (el.namespaceURI === W_NS) {
      if (el.localName === 'tc') return true;
      if (el.localName === 'body') return false; // 到达 body 层，不可能再有 tc
    }
    el = el.parentElement;
  }
  return false;
}

/**
 * 一次祖先遍历，同时拿到 `<w:tc>`、`<w:tr>`、最内层 `<w:tbl>` 三个祖先元素。
 *
 * 遍历策略：向上走，遇到首个 `<w:tc>` 记录为 tc，继续向上遇到 `<w:tr>` 记录为 tr，
 * 再继续向上遇到首个 `<w:tbl>` 记录为 tbl（即最内层表格），然后停止。
 * 遇到 story 根元素（body / hdr / ftr 等）仍未找到 `<w:tc>` 则返回 null（段落不在表格内）。
 *
 * 嵌套表格时 tbl 是最内层表格（因为自下而上遇到的第一个 tbl 就是最内层的）。
 */
function getTableCellAncestors(p: Element): { tc: Element; tr: Element; tbl: Element } | null {
  let tc: Element | null = null;
  let tr: Element | null = null;
  let el: Element | null = p.parentElement;

  while (el) {
    if (el.namespaceURI === W_NS) {
      if (!tc) {
        if (el.localName === 'tc') {
          tc = el;
        } else if (
          el.localName === 'body' ||
          el.localName === 'hdr' ||
          el.localName === 'ftr' ||
          el.localName === 'footnote' ||
          el.localName === 'endnote' ||
          el.localName === 'comment'
        ) {
          // 到达 story 根，段落不在表格内
          return null;
        }
      } else if (!tr) {
        if (el.localName === 'tr') {
          tr = el;
        }
      } else {
        // tc 和 tr 都已找到，找最内层 tbl
        if (el.localName === 'tbl') {
          return { tc, tr, tbl: el };
        }
      }
    }
    el = el.parentElement;
  }
  return null;
}

/**
 * 计算 `<w:tr>` 在其父 `<w:tbl>` 的直属 `<w:tr>` 子元素序列中的 0 基行号。
 * 只计直属子元素（不误抓嵌套表的 tr）。
 */
function rowIndexOf(tr: Element, tbl: Element): number {
  let idx = 0;
  const children = tbl.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i] as Element;
    if (c.nodeType !== 1 || c.namespaceURI !== W_NS || c.localName !== 'tr') continue;
    if (c === tr) return idx;
    idx++;
  }
  return -1; // 理论上不可达
}

/**
 * 计算 `<w:tc>` 在其父 `<w:tr>` 的直属 `<w:tc>` 子元素序列中的 0 基列号（物理序号）。
 *
 * 注：此为物理列号，不做 gridSpan / vMerge 的逻辑列号校正——含合并单元格时物理序号
 * 与逻辑列号可能不同。v1 物理序号已满足"区分表格内外编号序列"的主场景。
 */
function colIndexOf(tc: Element, tr: Element): number {
  let idx = 0;
  const children = tr.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i] as Element;
    if (c.nodeType !== 1 || c.namespaceURI !== W_NS || c.localName !== 'tc') continue;
    if (c === tc) return idx;
    idx++;
  }
  return -1; // 理论上不可达
}

/** 段内是否含 revision 标记(Tier 1:只看 w:ins / w:del) */
function detectHasRevisions(p: Element): boolean {
  return (
    p.getElementsByTagNameNS(W_NS, 'ins').length > 0 ||
    p.getElementsByTagNameNS(W_NS, 'del').length > 0
  );
}

/** 段内是否含 comment 标记 */
function detectHasComments(p: Element): boolean {
  return (
    p.getElementsByTagNameNS(W_NS, 'commentRangeStart').length > 0 ||
    p.getElementsByTagNameNS(W_NS, 'commentRangeEnd').length > 0 ||
    p.getElementsByTagNameNS(W_NS, 'commentReference').length > 0
  );
}

/** 段内是否含 bookmark 标记 */
function detectHasBookmarks(p: Element): boolean {
  return (
    p.getElementsByTagNameNS(W_NS, 'bookmarkStart').length > 0 ||
    p.getElementsByTagNameNS(W_NS, 'bookmarkEnd').length > 0
  );
}

/**
 * 段落子树是否含 OMML 公式对象（`<m:oMath>`）。
 * 覆盖块级（oMathPara > oMath）和内联（段内裸 oMath、run/hyperlink 内嵌 oMath）两种情形。
 * `getElementsByTagNameNS` 递归子树，天然覆盖所有深度。
 */
function detectHasMath(p: Element): boolean {
  return p.getElementsByTagNameNS(MATH_NS, 'oMath').length > 0;
}

// ─────────── 分页 ───────────

interface PaginationOut<T> {
  slice: T[];
  meta: ListResult['pagination'];
}

function paginate<T extends { el: Element }>(
  items: T[],
  pag: ListParagraphsPagination,
): PaginationOut<T> {
  const total = items.length;

  if (pag.mode === 'all') {
    return {
      slice: items,
      meta: {
        mode: 'all',
        pageIndex: 0,
        totalPages: total === 0 ? 0 : 1,
        totalParagraphs: total,
        hasNext: false,
      },
    };
  }

  if (pag.mode === 'byCount') {
    validatePositiveInt('pageSize', pag.pageSize);
    validateNonNegativeInt('pageIndex', pag.pageIndex);
    const totalPages = total === 0 ? 0 : Math.ceil(total / pag.pageSize);
    const start = pag.pageSize * pag.pageIndex;
    const slice = items.slice(start, start + pag.pageSize);
    return {
      slice,
      meta: {
        mode: 'byCount',
        pageIndex: pag.pageIndex,
        totalPages,
        totalParagraphs: total,
        hasNext: pag.pageIndex < totalPages - 1,
      },
    };
  }

  // byChar(docs/03 §9.3):按段落文本长度打包成页
  //   1. 累加当前段 text.length 到 buffer
  //   2. buf == 0 且 cur.length > charPerPage → 单段成页(保证完整)
  //   3. buf + next.length > charPerPage → 当前页结束,next 进下一页
  //   4. buf + next.length <= charPerPage → next 追加进当前页
  validatePositiveInt('charPerPage', pag.charPerPage);
  validateNonNegativeInt('pageIndex', pag.pageIndex);
  const pages: T[][] = [];
  let cur: T[] = [];
  let buf = 0;
  for (const item of items) {
    const len = getParagraphText(item.el).length;
    if (buf === 0 && len > pag.charPerPage) {
      // 规则 2:单段成页
      pages.push([item]);
      cur = [];
      buf = 0;
      continue;
    }
    if (buf + len > pag.charPerPage) {
      // 规则 3:当前页结束
      if (cur.length > 0) pages.push(cur);
      cur = [item];
      buf = len;
    } else {
      // 规则 4:追加到当前页
      cur.push(item);
      buf += len;
    }
  }
  if (cur.length > 0) pages.push(cur);
  const totalPages = pages.length;
  const slice = pages[pag.pageIndex] ?? [];
  return {
    slice,
    meta: {
      mode: 'byChar',
      pageIndex: pag.pageIndex,
      totalPages,
      totalParagraphs: total,
      hasNext: pag.pageIndex < totalPages - 1,
    },
  };
}

function validatePositiveInt(name: string, v: number): void {
  if (!Number.isInteger(v) || v < 1) {
    throw new Error(`listParagraphs: ${name} must be a positive integer, got: ${v}`);
  }
}

function validateNonNegativeInt(name: string, v: number): void {
  if (!Number.isInteger(v) || v < 0) {
    throw new Error(`listParagraphs: ${name} must be a non-negative integer, got: ${v}`);
  }
}

// ─────────── 单段查询(#25)───────────

/**
 * 按 paraId 精确取段。走 paraIdx → O(1) 命中后反查 story。不存在或所属 story 识别
 * 不出(理论上只在 pkg 与 indices 脱钩时发生)返回 null。
 *
 * 总是返回 full 文本(无 textFormat/preview 选项,与 docs/03.1 §2 的签名一致)。
 */
export function getParagraphByParaId(
  pkg: Package,
  indices: PackageIndices,
  paraId: string,
): ParagraphInfo | null {
  const el = indices.paraIdx.get(paraId.toUpperCase());
  if (!el) return null;
  const story = findStoryFor(pkg, el);
  if (!story) return null;
  return buildParagraphInfo(story.id, el, 'full');
}

// ─────────── ParagraphInfo 构建 ───────────

/**
 * 从单个 `<w:p>` Element 构建 ParagraphInfo。
 *
 * 在 listParagraphs 内部给每个命中段落调用;**也供 `getParagraph`(#25)复用** ——
 * 它通过 paraIdx 拿到 Element 后,用 `findStoryFor` 找到 story,再调这里组装。
 */
export function buildParagraphInfo(
  story: StoryId,
  p: Element,
  textFormat: ListParagraphsTextFormat = 'full',
  includeFormat = false,
  runTextOpts?: RunTextOptions,
): ParagraphInfo {
  return buildInfo(story, p, textFormat, includeFormat, runTextOpts);
}

function buildInfo(
  story: StoryId,
  p: Element,
  textFormat: ListParagraphsTextFormat,
  includeFormat: boolean,
  runTextOpts?: RunTextOptions,
  tblReverseIdx?: Map<Element, string>,
): ParagraphInfo {
  const paraIdRaw = p.getAttributeNS(W14_NS, 'paraId') ?? '';
  // 与 indices.ts 的 paraIdx 一致,对外用大写(indices.ts:116 也做了 toUpperCase)
  const paraId = paraIdRaw.toUpperCase();

  const fullText = getParagraphText(p, runTextOpts);
  const truncated = textFormat === 'preview' && fullText.length > PREVIEW_CHAR_LIMIT;
  const text = truncated ? fullText.slice(0, PREVIEW_CHAR_LIMIT) : fullText;

  const styleId = getStyleId(p);
  const hasRevisions = detectHasRevisions(p);
  const hasComments = detectHasComments(p);
  const hasBookmarks = detectHasBookmarks(p);
  const hasMath = detectHasMath(p);

  const info: ParagraphInfo = { paraId, story, text };
  if (truncated) info.truncated = true;
  if (styleId !== undefined) info.styleId = styleId;
  if (hasRevisions) info.hasRevisions = true;
  if (hasComments) info.hasComments = true;
  if (hasBookmarks) info.hasBookmarks = true;
  if (hasMath) info.hasMath = true;

  if (includeFormat) {
    const pPr = firstDirectChild(p, W_NS, 'pPr');
    let fmt = pPr ? parseParagraphFormat(pPr) : undefined;
    // hasManualPageBreak 来源是 run 而非 pPr，单独检测后合并
    if (detectManualPageBreak(p)) {
      fmt = { ...(fmt ?? {}), hasManualPageBreak: true };
    }
    // firstRunColor 来源是首个直属 run 的 rPr（task-20260615-redhead-fingerprint），单独检测后合并
    const runColor = detectFirstRunColor(p);
    if (runColor !== undefined) {
      fmt = { ...(fmt ?? {}), firstRunColor: runColor };
    }
    if (fmt) info.format = fmt;
  }

  // includeTableInfo=true 时：一次祖先遍历拿 tc/tr/tbl，反查 tableId，计算 row/col
  if (tblReverseIdx !== undefined) {
    const ancestors = getTableCellAncestors(p);
    if (ancestors !== null) {
      const { tc, tr, tbl } = ancestors;
      const tableId = tblReverseIdx.get(tbl);
      if (tableId !== undefined) {
        const row = rowIndexOf(tr, tbl);
        const col = colIndexOf(tc, tr);
        info.table = { tableId, row, col };
      }
    }
  }

  return info;
}

// ─────────── 段落直接格式解析(#182)───────────

/**
 * 从 `<w:pPr>` Element 解析段落直接格式属性。
 *
 * 只读取 pPr 上**直接写明**的值；继承自样式或文档默认值的属性不在此处理。
 * 若 pPr 中没有任何已知格式元素，返回 undefined（调用方可跳过 format 字段赋值）。
 */
function parseParagraphFormat(pPr: Element): ParagraphFormat | undefined {
  const fmt: ParagraphFormat = {};
  let hasAny = false;

  // ── 对齐 w:jc ──
  const jc = firstDirectChild(pPr, W_NS, 'jc');
  if (jc) {
    const val = jc.getAttributeNS(W_NS, 'val');
    if (
      val === 'left' ||
      val === 'center' ||
      val === 'right' ||
      val === 'both' ||
      val === 'distribute'
    ) {
      fmt.alignment = val;
      hasAny = true;
    }
  }

  // ── 缩进 w:ind ──
  const ind = firstDirectChild(pPr, W_NS, 'ind');
  if (ind) {
    const left = numAttrTwips(ind, 'left');
    const right = numAttrTwips(ind, 'right');
    const firstLineRaw = numAttrTwips(ind, 'firstLine');
    const hangingRaw = numAttrTwips(ind, 'hanging');

    // firstLine 和 hanging 互斥，同时出现时 OOXML 规范以 firstLine 优先
    const firstLine =
      firstLineRaw !== undefined
        ? firstLineRaw
        : hangingRaw !== undefined
          ? -hangingRaw
          : undefined;

    if (left !== undefined || right !== undefined || firstLine !== undefined) {
      fmt.indent = {};
      if (left !== undefined) fmt.indent.left = left;
      if (right !== undefined) fmt.indent.right = right;
      if (firstLine !== undefined) fmt.indent.firstLine = firstLine;
      hasAny = true;
    }
  }

  // ── 段间距 w:spacing ──
  const spacing = firstDirectChild(pPr, W_NS, 'spacing');
  if (spacing) {
    const before = numAttrTwips(spacing, 'before');
    const after = numAttrTwips(spacing, 'after');
    const line = numAttrTwips(spacing, 'line');
    const lineRuleRaw = spacing.getAttributeNS(W_NS, 'lineRule');
    const lineRule =
      lineRuleRaw === 'auto' || lineRuleRaw === 'exact' || lineRuleRaw === 'atLeast'
        ? lineRuleRaw
        : undefined;

    if (
      before !== undefined ||
      after !== undefined ||
      line !== undefined ||
      lineRule !== undefined
    ) {
      fmt.spacing = {};
      if (before !== undefined) fmt.spacing.before = before;
      if (after !== undefined) fmt.spacing.after = after;
      if (line !== undefined) fmt.spacing.line = line;
      if (lineRule !== undefined) fmt.spacing.lineRule = lineRule;
      hasAny = true;
    }
  }

  // ── 段落级默认字体大小 w:pPr/w:rPr/w:sz ──
  const pRpr = firstDirectChild(pPr, W_NS, 'rPr');
  if (pRpr) {
    const sz = firstDirectChild(pRpr, W_NS, 'sz');
    if (sz) {
      const szVal = numAttrTwips(sz, 'val'); // w:sz 单位是半磅
      if (szVal !== undefined) {
        fmt.fontSize = szVal / 2; // 转为磅
        hasAny = true;
      }
    }

    // ── 段落级默认粗体 w:pPr/w:rPr/w:b ──
    const b = firstDirectChild(pRpr, W_NS, 'b');
    if (b) {
      const bVal = b.getAttributeNS(W_NS, 'val');
      // 无 val 属性或 val="1"/"true" → bold; val="0"/"false" → 显式关闭
      fmt.bold = bVal !== '0' && bVal !== 'false';
      hasAny = true;
    }

    // ── 段落级默认颜色 w:pPr/w:rPr/w:color（task-20260615-redhead-fingerprint）──
    const color = firstDirectChild(pRpr, W_NS, 'color');
    if (color) {
      const cVal = color.getAttributeNS(W_NS, 'val');
      if (cVal) {
        fmt.color = cVal;
        hasAny = true;
      }
    }
  }

  // ── 大纲级别 w:outlineLvl ──
  const outlineLvl = firstDirectChild(pPr, W_NS, 'outlineLvl');
  if (outlineLvl) {
    const val = numAttrTwips(outlineLvl, 'val');
    if (val !== undefined && val >= 0 && val <= 9) {
      fmt.outlineLevel = val;
      hasAny = true;
    }
  }

  // ── 编号属性 w:numPr ──
  const numPr = firstDirectChild(pPr, W_NS, 'numPr');
  if (numPr) {
    const numId = firstDirectChild(numPr, W_NS, 'numId');
    // numId@w:val="0" 表示移除继承来的编号，不算"有编号"
    const numIdVal = numId?.getAttributeNS(W_NS, 'val');
    if (numIdVal !== undefined && numIdVal !== '0') {
      fmt.hasNumbering = true;
      hasAny = true;
    }
  }

  // ── 强制分页 w:pageBreakBefore（#211）──
  // 存在 = true；带 val="0"/"false" = 显式关闭
  const pbb = firstDirectChild(pPr, W_NS, 'pageBreakBefore');
  if (pbb) {
    const v = pbb.getAttributeNS(W_NS, 'val');
    fmt.pageBreakBefore = v !== '0' && v !== 'false';
    hasAny = true;
  }

  // ── 节属性 w:sectPr（#211）──
  // 存在即视为该段承载节边界
  if (firstDirectChild(pPr, W_NS, 'sectPr')) {
    fmt.hasSectionBreak = true;
    hasAny = true;
  }

  return hasAny ? fmt : undefined;
}

/**
 * 检测段落内是否含手工分页符（`<w:r>/<w:br w:type="page"/>`，#211）。
 *
 * 遍历段落直属 `<w:r>` 的 `<w:br>` 子元素，找 type='page' 的即返回 true。
 * 段属性外的 ins/del 包装内 run 暂不深入（保守：v1 只看直属 run，与 getRunText 行为对齐）。
 */
function detectManualPageBreak(p: Element): boolean {
  const runs = p.getElementsByTagNameNS(W_NS, 'r');
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    if (!run || run.parentNode !== p) continue; // 仅直属 run
    const brs = run.getElementsByTagNameNS(W_NS, 'br');
    for (let j = 0; j < brs.length; j++) {
      const br = brs[j];
      if (!br) continue;
      const type = br.getAttributeNS(W_NS, 'type');
      if (type === 'page') return true;
    }
  }
  return false;
}

/**
 * 读取段落**首个直属 run**（`<w:p>/<w:r>`）的字体颜色（`w:rPr/w:color@w:val`）。
 * task-20260615-redhead-fingerprint：红头/标题等视觉显著段的颜色常在 run 级而非段标级。
 *
 * 只看首个直属 run（`run.parentNode === p`），与 getParagraphText / detectManualPageBreak 的
 * "仅直属 run"语义一致——不下钻表格/文本框/超链接内的 run。首 run 无 color 则返回 undefined。
 */
function detectFirstRunColor(p: Element): string | undefined {
  const runs = p.getElementsByTagNameNS(W_NS, 'r');
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    if (!run || run.parentNode !== p) continue; // 仅直属 run
    const rPr = firstDirectChild(run, W_NS, 'rPr');
    if (!rPr) return undefined; // 首个直属 run 无 rPr → 无 run 级颜色
    const color = firstDirectChild(rPr, W_NS, 'color');
    if (!color) return undefined;
    const val = color.getAttributeNS(W_NS, 'val');
    return val ?? undefined;
  }
  return undefined;
}

/** 读取 Element 上指定 w: 属性并解析为数字，不是有效整数则返回 undefined。 */
function numAttrTwips(el: Element, local: string): number | undefined {
  const raw = el.getAttributeNS(W_NS, local);
  if (raw === null) return undefined;
  const n = parseInt(raw, 10);
  return Number.isFinite(n) ? n : undefined;
}
