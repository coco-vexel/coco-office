/**
 * 表格枚举与结构分析（#185）
 *
 * 为骨架化 M2 阶段（表格分类/取样/域化）提供表格数据。
 *
 * 设计原则：
 * - 只走 getElementsByTagNameNS，禁用 XPath（性能约束，同 list-paragraphs.ts）
 * - tableId 由 buildIndices 在装载期分配（tbl-0001 起），存入 tblIdx
 * - 默认只枚举主文档 body（scope: 'body'），不含嵌套表格（includeNested: false）
 * - 前后邻段落（prevParaId / nextParaId）只取紧邻兄弟 <w:p>，供题注绑定（#193）
 */

import { serializeToWellFormedString } from 'slimdom';
import { getPart, isXmlPart, type Package } from '../pack/index.js';
import { A_NS, R_NS, W_NS, W14_NS } from './namespaces.js';
import type { PackageIndices } from './indices.js';
import { getParagraphText } from './runs.js';
import { findStoryFor, resolveScope, type Story } from './stories.js';

// ─────────── 公开类型 ───────────

/**
 * 行高（对照 columnWidths）。w:tr/w:trPr/w:trHeight。
 *
 * 空行塌缩正是 hRule+val 没带过去所致，两者必须一起读/写。
 */
export interface RowHeight {
  /** w:trHeight/@w:val，行高 twips（1 twip = 1/20 磅）*/
  val: number;
  /** w:trHeight/@w:hRule：auto(默认,最小高度) / atLeast(至少) / exact(精确)。
   *  缺省（无 hRule 属性）按 OOXML 默认 'auto'。*/
  hRule: 'auto' | 'atLeast' | 'exact';
}

/**
 * 单个物理 <w:tc> 的 span 信息，携带文本与合并状态。
 * cellGrid[i][j] 对应第 i 行第 j 个物理 <w:tc>。
 */
export interface TableCellInfo {
  /** 单元格纯文本（多段以 \n 连接），与 cells[i][j] 完全一致 */
  text: string;
  /** 水平合并跨列数（<w:gridSpan w:val>，无则默认 1） */
  gridSpan: number;
  /**
   * 垂直合并状态：
   *  'restart'  = 纵向合并的起始格（<w:vMerge w:val="restart"/>）
   *  'continue' = 被并入上方格的续行占位格（<w:vMerge/> 无 val 或 val="continue"）
   *  null       = 无纵向合并
   */
  vMerge: 'restart' | 'continue' | null;
  /** Direct cell width from w:tcPr/w:tcW, in twips. */
  width?: number;
  /** Direct cell shading from w:tcPr/w:shd. */
  shading?: {
    fill?: string;
    color?: string;
    val?: string;
  };
  /** Direct vertical alignment from w:tcPr/w:vAlign. */
  verticalAlignment?: 'top' | 'center' | 'bottom' | 'both';
  /**
   * 单元格首段（首个直属 `<w:p>`）的 paraId（w14:paraId）。无首段或首段缺 paraId 时 undefined。
   *
   * 用途（task-bug2-cell-image，决策点 C）：merge 重建表后要往**正确 cell** 插单元格图，
   * 而 clone 出来的表 cell 段 paraId 是 cloneTableAfter 重分配的新值——调用方拿不到。
   * 暴露 firstParaId 后，调用方对 clone 表再 `tables.list({includeCellTexts:true})` 即可拿到
   * 每个 cell 首段 paraId，用 `plan.insertImageAfter(firstParaId)` 把图插进该 cell（insertImageAfter
   * 在 anchor 的 parentNode 内插段——anchor 在 `<w:tc>` 里，新图段就落在该 cell 内）。
   *
   * 与 replaceTableCellText 无冲突：后者只动 cell **首段**的 run，插在首段之后的独立图段不受影响。
   */
  firstParaId?: string;
  /**
   * Direct paragraph ids in this cell, in cell content order.
   *
   * Nested table paragraphs are not included.
   */
  paragraphIds: string[];
  /**
   * Direct content blocks in this cell, in OOXML child-node order.
   *
   * Paragraph drawings are emitted as drawing blocks as well, so callers can
   * preserve flows such as text paragraphs, then an image, then more text.
   */
  blocks: TableCellBlockInfo[];
  /**
   * 本单元格内直属 `<w:drawing>`（图片/形状等）的 drawingId 列表（与 drawingIdx 同 ID 空间）。
   *
   * 用途（task-bug2-cell-image，决策点 A1）：表格单元格内的图此前在 cellGrid 这层就丢了
   * （cellGrid 只搬 cell 文字），导致单元格图在 structure 抽取阶段成孤儿（图字节在 word/media
   * 但 body 无 blip 引用）。收录 drawingId 后，typeset 抽表时把 cell 图当 content 带出、merge
   * 重建时用 `drawing.getImageBytes()` + `insertImageAfter(firstParaId)` 把图插回对应 cell。
   *
   * 收录范围：只取**本 cell 子树内、且不在嵌套表/嵌套 cell 内**的 drawing（嵌套表 cell 的图归属
   * 嵌套表自己的 cellGrid，不上浮污染外层）。空（无图）时为 `[]`（与 cells 同条件填充，
   * 字段恒在便于下游免判 undefined）。
   */
  imageDrawingIds: string[];
}

export type TableCellBlockInfo =
  | { type: 'paragraph'; paraId: string; text: string }
  | { type: 'drawing'; drawingId: string; paraId?: string }
  | { type: 'table'; tableId: string };

export interface TableInfo {
  /** 分配的表格 ID（装载期顺序，'tbl-0001' 起），供 plan executor 引用 */
  tableId: string;
  /** 所在 part 路径（如 'word/document.xml'）*/
  partPath: string;
  /** 在本次枚举结果中的顺序（1-based）*/
  index: number;
  /** 是否嵌套在另一个表格内（祖先链上存在 <w:tbl>）*/
  isNested: boolean;
  /** 行数（w:tr 直接子元素数量）*/
  rowCount: number;
  /** 列数（w:tblGrid/w:gridCol 数量；无 tblGrid 时为 0）*/
  colCount: number;
  /** 是否有合并单元格（gridSpan > 1 或存在 vMerge）*/
  hasMerge: boolean;
  /** 表格样式 ID（w:tblPr/w:tblStyle @w:val），未设时 undefined */
  tableStyleId?: string;
  /** 是否包含图片（w:drawing 元素）*/
  hasImages: boolean;
  /** 是否包含嵌套表格（直接子孙 <w:tbl>）*/
  hasNestedTables: boolean;
  /**
   * 紧邻前一个兄弟 <w:p> 的 paraId，用于题注绑定（#193）。
   * 仅检查紧邻兄弟：若前一个兄弟不是 <w:p>（如另一个 <w:tbl>）则为 undefined。
   */
  prevParaId?: string;
  /**
   * 紧邻后一个兄弟 <w:p> 的 paraId，用于题注绑定（#193）。
   * 仅检查紧邻兄弟。
   */
  nextParaId?: string;
  /**
   * 每列的列宽（twips，来自 `<w:tblGrid>/<w:gridCol w:w="...">`）。
   *
   * - `columnWidths[i]` = 第 i 列的 `w:w` 值（1 twip = 1/20 磅）
   * - 无 `<w:tblGrid>` 时为 `[]`（空数组）
   * - `tblLayout=fixed` 的表格列宽完全由此决定；`auto` 布局表格也记录但仅供参考
   * - 供 merger 套用源表列宽（task-core-table-grid：克隆模板时保真列宽）
   */
  columnWidths: number[];
  /**
   * 每行行高，doc 序，每行一个；该行无 <w:trHeight> 则为 null。
   *
   * - `rowHeights[i]` = 第 i 行的 w:trPr/w:trHeight val+hRule
   * - 恒填（不依赖 includeCellTexts）
   * - 只取直属 <w:tr>，不误抓嵌套表行
   * - 供 merger 套用源表行高（空行塌缩修复）
   */
  rowHeights: (RowHeight | null)[];
  /**
   * 单元格文本网格（#186）。仅当 `includeCellTexts: true` 时填充。
   *
   * - `cells[i][j]` = 第 i 行第 j 列单元格的纯文本
   * - 单元格内含多段时用 `\n` 连接（保留分段信号，供 has-multi-content 判定）
   * - 嵌套表格内的段落不展开（不污染外层网格）
   */
  cells?: string[][];
  /**
   * 单元格网格（span-aware），仅当 `includeCellTexts: true` 时填充（与 cells 同条件）。
   *
   * - `cellGrid[i][j]` 对应第 i 行第 j 个**物理** `<w:tc>`，携带 gridSpan/vMerge
   * - 与 cells 同源同序：`cells[i][j] === cellGrid[i][j].text`
   * - 嵌套表内段落不展开（与 extractCellTexts 一致）
   */
  cellGrid?: TableCellInfo[][];
}

export interface ListTablesOptions {
  /** 枚举范围（默认 'body'）*/
  scope?: 'body' | 'all';
  /** 是否包含嵌套在其他表格内的表格（默认 false，只返回顶层表格）*/
  includeNested?: boolean;
  /**
   * 是否抽取单元格文本（#186）。默认 false（保持 #185 默认行为）。
   *
   * 开启后每张 TableInfo 会带 `cells: string[][]`，用于关键词匹配/分类。
   * 性能：每张表会枚举所有 <w:tc> 及其直属 <w:p>，按需开启。
   */
  includeCellTexts?: boolean;
}

// ─────────── 内部工具 ───────────

/** 把 HTMLCollection-ish 枚举结果快照成数组（避免 live 行为）*/
function snapshot<T extends Element>(col: { length: number } & Record<number, T>): T[] {
  const out: T[] = [];
  for (let i = 0; i < col.length; i++) {
    const el = col[i];
    if (el) out.push(el);
  }
  return out;
}

/** 判断元素是否嵌套在某个祖先 <w:tbl> 内（止于 story 根元素）*/
function hasTableAncestor(el: Element): boolean {
  let p = el.parentElement;
  while (p) {
    if (p.namespaceURI === W_NS) {
      if (p.localName === 'tbl') return true;
      // story 根元素：停止向上走
      if (
        p.localName === 'body' ||
        p.localName === 'hdr' ||
        p.localName === 'ftr' ||
        p.localName === 'footnote' ||
        p.localName === 'endnote' ||
        p.localName === 'comment'
      ) {
        return false;
      }
    }
    p = p.parentElement;
  }
  return false;
}

/** 紧邻前一个兄弟 <w:p> 的 paraId（只看直接前兄弟）*/
function prevParaId(tbl: Element): string | undefined {
  const sib = tbl.previousElementSibling;
  if (sib && sib.namespaceURI === W_NS && sib.localName === 'p') {
    return sib.getAttributeNS(W14_NS, 'paraId') ?? undefined;
  }
  return undefined;
}

/** 紧邻后一个兄弟 <w:p> 的 paraId（只看直接后兄弟）*/
function nextParaId(tbl: Element): string | undefined {
  const sib = tbl.nextElementSibling;
  if (sib && sib.namespaceURI === W_NS && sib.localName === 'p') {
    return sib.getAttributeNS(W14_NS, 'paraId') ?? undefined;
  }
  return undefined;
}

/**
 * 读取 <w:tc> 的直属 <w:tcPr> 中的 gridSpan（返回整数，无则 1）。
 * 只查 tcPr 直属子元素，不用 getElementsByTagNameNS（避免误抓嵌套表）。
 */
function readGridSpan(tc: Element): number {
  for (let i = 0; i < tc.childNodes.length; i++) {
    const child = tc.childNodes[i] as Element;
    if (child.nodeType !== 1 || child.namespaceURI !== W_NS || child.localName !== 'tcPr') continue;
    // found tcPr — look for direct child gridSpan
    for (let j = 0; j < child.childNodes.length; j++) {
      const sub = child.childNodes[j] as Element;
      if (sub.nodeType !== 1 || sub.namespaceURI !== W_NS || sub.localName !== 'gridSpan') continue;
      const val = sub.getAttributeNS(W_NS, 'val');
      const parsed = val !== null ? parseInt(val, 10) : NaN;
      return !isNaN(parsed) && parsed > 1 ? parsed : 1;
    }
    break; // tcPr found but no gridSpan child
  }
  return 1;
}

/**
 * 读取 <w:tc> 的直属 <w:tcPr> 中的 vMerge 状态。
 * - 'restart'  = <w:vMerge w:val="restart"/>
 * - 'continue' = <w:vMerge/> (无 val 或 val="continue")
 * - null       = 无 vMerge 元素
 * 只查 tcPr 直属子元素，不用 getElementsByTagNameNS。
 */
function readVMerge(tc: Element): 'restart' | 'continue' | null {
  for (let i = 0; i < tc.childNodes.length; i++) {
    const child = tc.childNodes[i] as Element;
    if (child.nodeType !== 1 || child.namespaceURI !== W_NS || child.localName !== 'tcPr') continue;
    // found tcPr — look for direct child vMerge
    for (let j = 0; j < child.childNodes.length; j++) {
      const sub = child.childNodes[j] as Element;
      if (sub.nodeType !== 1 || sub.namespaceURI !== W_NS || sub.localName !== 'vMerge') continue;
      const val = sub.getAttributeNS(W_NS, 'val');
      return val === 'restart' ? 'restart' : 'continue';
    }
    break; // tcPr found but no vMerge child
  }
  return null;
}

/**
 * 读取 <w:tr> 的直属 <w:trPr> 中的 trHeight（val + hRule）。
 * - 无 trPr 或 trPr 内无 trHeight → null
 * - 只查 trPr 直属子元素，不用 getElementsByTagNameNS（避免误抓嵌套表）。
 */
function readCellWidth(tc: Element): number | undefined {
  const tcPr = firstDirectChildW(tc, 'tcPr');
  if (!tcPr) return undefined;
  const tcW = firstDirectChildW(tcPr, 'tcW');
  if (!tcW) return undefined;
  const raw = tcW.getAttributeNS(W_NS, 'w') ?? tcW.getAttribute('w:w');
  if (raw === null) return undefined;
  const parsed = parseInt(raw, 10);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function readCellShading(tc: Element): { fill?: string; color?: string; val?: string } | undefined {
  const tcPr = firstDirectChildW(tc, 'tcPr');
  if (!tcPr) return undefined;
  const shd = firstDirectChildW(tcPr, 'shd');
  if (!shd) return undefined;
  const result: { fill?: string; color?: string; val?: string } = {};
  const fill = shd.getAttributeNS(W_NS, 'fill') ?? shd.getAttribute('w:fill');
  const color = shd.getAttributeNS(W_NS, 'color') ?? shd.getAttribute('w:color');
  const val = shd.getAttributeNS(W_NS, 'val') ?? shd.getAttribute('w:val');
  if (fill !== null) result.fill = fill;
  if (color !== null) result.color = color;
  if (val !== null) result.val = val;
  return Object.keys(result).length > 0 ? result : undefined;
}

function readCellVerticalAlignment(tc: Element): 'top' | 'center' | 'bottom' | 'both' | undefined {
  const tcPr = firstDirectChildW(tc, 'tcPr');
  if (!tcPr) return undefined;
  const vAlign = firstDirectChildW(tcPr, 'vAlign');
  if (!vAlign) return undefined;
  const raw = vAlign.getAttributeNS(W_NS, 'val') ?? vAlign.getAttribute('w:val');
  if (raw === 'top' || raw === 'center' || raw === 'bottom' || raw === 'both') return raw;
  return undefined;
}

function firstDirectChildW(el: Element, localName: string): Element | null {
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName === localName
    ) {
      return child as Element;
    }
  }
  return null;
}

function readTrHeight(tr: Element): RowHeight | null {
  for (let i = 0; i < tr.childNodes.length; i++) {
    const child = tr.childNodes[i] as Element;
    if (child.nodeType !== 1 || child.namespaceURI !== W_NS || child.localName !== 'trPr') continue;
    // found trPr — look for direct child trHeight
    for (let j = 0; j < child.childNodes.length; j++) {
      const sub = child.childNodes[j] as Element;
      if (sub.nodeType !== 1 || sub.namespaceURI !== W_NS || sub.localName !== 'trHeight') continue;
      const valAttr = sub.getAttributeNS(W_NS, 'val');
      const val = valAttr !== null ? parseInt(valAttr, 10) : 0;
      const hRuleAttr = sub.getAttributeNS(W_NS, 'hRule');
      let hRule: RowHeight['hRule'] = 'auto';
      if (hRuleAttr === 'atLeast') hRule = 'atLeast';
      else if (hRuleAttr === 'exact') hRule = 'exact';
      return { val: isNaN(val) ? 0 : val, hRule };
    }
    break; // trPr found but no trHeight child
  }
  return null;
}

/**
 * 取 `<w:tc>` 首个直属 `<w:p>` 的 w14:paraId（task-bug2-cell-image，决策点 C）。
 * 无直属 `<w:p>` 或缺 paraId 时返回 undefined。
 */
function firstCellParaId(tc: Element): string | undefined {
  for (let k = 0; k < tc.childNodes.length; k++) {
    const p = tc.childNodes[k] as Element;
    if (!p || p.nodeType !== 1 || p.namespaceURI !== W_NS || p.localName !== 'p') continue;
    return p.getAttributeNS(W14_NS, 'paraId') ?? undefined;
  }
  return undefined;
}

/**
 * 收集**本 cell 子树内、不在嵌套表内**的 `<w:drawing>`，反查出 drawingId 列表
 * （task-bug2-cell-image，决策点 A1）。
 *
 * - 用 `getElementsByTagNameNS(W_NS,'drawing')` 取 cell 子树全部 drawing（含嵌套表内的），
 *   再用 `cellOf(drawing) === tc` 过滤掉嵌套表 cell 内的图（它们归属嵌套表自己的 cellGrid）。
 * - 反查 drawingId 走传入的 `drawingToId`（drawingIdx 反转），未命中（不在索引内，
 *   如 mc:Fallback 内被装载期跳过的 drawing）则跳过——与 drawings.list 枚举集保持一致。
 */
function collectCellDrawingIds(tc: Element, drawingToId: Map<Element, string>): string[] {
  const ids: string[] = [];
  const drawings = tc.getElementsByTagNameNS(W_NS, 'drawing');
  for (let i = 0; i < drawings.length; i++) {
    const dr = drawings[i];
    if (!dr) continue;
    // 只收归属本 cell（最近祖先 <w:tc> 即 tc）的 drawing，排除嵌套表 cell 内的图。
    if (nearestCellOf(dr) !== tc) continue;
    const id = drawingToId.get(dr);
    if (id !== undefined) ids.push(id);
  }
  return ids;
}

/** 向上找元素最近的祖先 `<w:tc>`（止于 story 根）；不在任何 cell 内返回 null。 */
function nearestCellOf(el: Element): Element | null {
  let p = el.parentElement;
  while (p) {
    if (p.namespaceURI === W_NS) {
      if (p.localName === 'tc') return p;
      if (
        p.localName === 'body' ||
        p.localName === 'hdr' ||
        p.localName === 'ftr' ||
        p.localName === 'footnote' ||
        p.localName === 'endnote' ||
        p.localName === 'comment'
      ) {
        return null;
      }
    }
    p = p.parentElement;
  }
  return null;
}

function collectCellParagraphIds(tc: Element): string[] {
  const ids: string[] = [];
  for (let k = 0; k < tc.childNodes.length; k++) {
    const p = tc.childNodes[k] as Element;
    if (!p || p.nodeType !== 1 || p.namespaceURI !== W_NS || p.localName !== 'p') continue;
    const paraId = p.getAttributeNS(W14_NS, 'paraId');
    if (paraId) ids.push(paraId);
  }
  return ids;
}

function directDrawingIdsInParagraph(
  p: Element,
  ownerCell: Element,
  drawingToId: Map<Element, string>,
): string[] {
  const ids: string[] = [];
  const drawings = p.getElementsByTagNameNS(W_NS, 'drawing');
  for (let i = 0; i < drawings.length; i++) {
    const dr = drawings[i];
    if (!dr || nearestCellOf(dr) !== ownerCell) continue;
    const id = drawingToId.get(dr);
    if (id !== undefined) ids.push(id);
  }
  return ids;
}

function collectCellBlocks(
  tc: Element,
  drawingToId: Map<Element, string>,
  tableToId: Map<Element, string>,
): TableCellBlockInfo[] {
  const blocks: TableCellBlockInfo[] = [];
  for (let k = 0; k < tc.childNodes.length; k++) {
    const child = tc.childNodes[k] as Element;
    if (!child || child.nodeType !== 1 || child.namespaceURI !== W_NS) continue;

    if (child.localName === 'p') {
      const paraId = child.getAttributeNS(W14_NS, 'paraId') ?? undefined;
      const text = getParagraphText(child);
      const drawingIds = directDrawingIdsInParagraph(child, tc, drawingToId);
      if (paraId && (text.length > 0 || drawingIds.length === 0)) {
        blocks.push({ type: 'paragraph', paraId, text });
      }
      for (const drawingId of drawingIds) {
        blocks.push(
          paraId ? { type: 'drawing', drawingId, paraId } : { type: 'drawing', drawingId },
        );
      }
      continue;
    }

    if (child.localName === 'tbl') {
      const tableId = tableToId.get(child);
      if (tableId) blocks.push({ type: 'table', tableId });
    }
  }
  return blocks;
}

/**
 * 抽取整张表的单元格网格（span-aware, task-core-table-span）。
 * - 只看 <w:tbl> 直属的 <w:tr>，每行枚举直属 <w:tc>
 * - 每个 <w:tc> 内取直属 <w:p>（不展开嵌套表），多段以 `\n` 连接
 * - 每格携带 gridSpan / vMerge（只读 tcPr 直属子元素，不误抓嵌套表）
 * - 每格携带 firstParaId（首段 paraId）/ imageDrawingIds（cell 内图 drawingId）
 *   （task-bug2-cell-image：决策点 C / A1）
 *
 * @param drawingToId  drawingIdx 反转（Element→drawingId），用于把 cell 内 `<w:drawing>` 反查成 ID。
 */
function extractCellGrid(
  tbl: Element,
  drawingToId: Map<Element, string>,
  tableToId: Map<Element, string>,
): TableCellInfo[][] {
  const rows: TableCellInfo[][] = [];
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const tr = tbl.childNodes[i] as Element;
    if (!tr || tr.nodeType !== 1 || tr.namespaceURI !== W_NS || tr.localName !== 'tr') continue;

    const rowCells: TableCellInfo[] = [];
    for (let j = 0; j < tr.childNodes.length; j++) {
      const tc = tr.childNodes[j] as Element;
      if (!tc || tc.nodeType !== 1 || tc.namespaceURI !== W_NS || tc.localName !== 'tc') continue;

      const paragraphs: string[] = [];
      for (let k = 0; k < tc.childNodes.length; k++) {
        const p = tc.childNodes[k] as Element;
        if (!p || p.nodeType !== 1 || p.namespaceURI !== W_NS || p.localName !== 'p') continue;
        paragraphs.push(getParagraphText(p));
      }
      const cell: TableCellInfo = {
        text: paragraphs.join('\n'),
        gridSpan: readGridSpan(tc),
        vMerge: readVMerge(tc),
        paragraphIds: collectCellParagraphIds(tc),
        blocks: collectCellBlocks(tc, drawingToId, tableToId),
        imageDrawingIds: collectCellDrawingIds(tc, drawingToId),
      };
      const fpid = firstCellParaId(tc);
      if (fpid !== undefined) cell.firstParaId = fpid;
      const width = readCellWidth(tc);
      if (width !== undefined) cell.width = width;
      const shading = readCellShading(tc);
      if (shading !== undefined) cell.shading = shading;
      const verticalAlignment = readCellVerticalAlignment(tc);
      if (verticalAlignment !== undefined) cell.verticalAlignment = verticalAlignment;
      rowCells.push(cell);
    }
    rows.push(rowCells);
  }
  return rows;
}

/** 从 <w:tbl> 元素提取结构信息。
 *
 * @param drawingToId  drawingIdx 反转（Element→drawingId），仅在 includeCellTexts 时用于
 *   cell 图 drawingId 反查；不开 includeCellTexts 时传空 Map 即可（不进 extractCellGrid）。
 */
function analyzeTable(
  tbl: Element,
  tableId: string,
  partPath: string,
  index: number,
  includeCellTexts: boolean,
  drawingToId: Map<Element, string>,
  tableToId: Map<Element, string>,
): TableInfo {
  let rowCount = 0;
  let hasMerge = false;
  const rowHeights: (RowHeight | null)[] = [];

  // 嵌套表深度检测（#203）：嵌套表在 tbl/tr/tc/tbl 路径上，不是 tbl 直接子，
  // 旧逻辑扫直接子永远扫不到。getElementsByTagNameNS 只返回后代（不含自身），
  // 命中即说明存在嵌套表。
  const hasNestedTables = tbl.getElementsByTagNameNS(W_NS, 'tbl').length > 0;

  // 一次遍历直接子节点：统计行数、检测合并单元格、读取行高
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const child = tbl.childNodes[i] as Element;
    if (child.nodeType !== 1 /* ELEMENT_NODE */) continue;
    if (child.namespaceURI !== W_NS) continue;

    if (child.localName === 'tr') {
      rowCount++;
      // 读取行高（直属 tr，不误抓嵌套表行）
      rowHeights.push(readTrHeight(child));
      // 检测合并单元格（只在未确认时扫）
      if (!hasMerge) {
        const cells = snapshot(child.getElementsByTagNameNS(W_NS, 'tc'));
        for (const tc of cells) {
          // 水平合并：gridSpan > 1
          const gridSpanEls = tc.getElementsByTagNameNS(W_NS, 'gridSpan');
          if (gridSpanEls.length > 0) {
            const val = gridSpanEls[0]!.getAttributeNS(W_NS, 'val');
            if (val && Number(val) > 1) {
              hasMerge = true;
              break;
            }
          }
          // 垂直合并：vMerge 存在即可
          if (tc.getElementsByTagNameNS(W_NS, 'vMerge').length > 0) {
            hasMerge = true;
            break;
          }
        }
      }
    }
  }

  // 列数 + 列宽：w:tblGrid/w:gridCol 数量及每列 w:w（task-core-table-grid）
  let colCount = 0;
  const columnWidths: number[] = [];
  const tblGridEls = tbl.getElementsByTagNameNS(W_NS, 'tblGrid');
  if (tblGridEls.length > 0) {
    const gridCols = tblGridEls[0]!.getElementsByTagNameNS(W_NS, 'gridCol');
    colCount = gridCols.length;
    for (let i = 0; i < gridCols.length; i++) {
      const wAttr = gridCols[i]!.getAttributeNS(W_NS, 'w');
      columnWidths.push(wAttr !== null ? Number(wAttr) : 0);
    }
  }

  // 表格样式
  let tableStyleId: string | undefined;
  const tblStyleEls = tbl.getElementsByTagNameNS(W_NS, 'tblStyle');
  if (tblStyleEls.length > 0) {
    tableStyleId = tblStyleEls[0]!.getAttributeNS(W_NS, 'val') ?? undefined;
  }

  // 图片检测：w:drawing 子孙
  const hasImages = tbl.getElementsByTagNameNS(W_NS, 'drawing').length > 0;

  const info: TableInfo = {
    tableId,
    partPath,
    index,
    isNested: hasTableAncestor(tbl),
    rowCount,
    colCount,
    hasMerge,
    hasImages,
    hasNestedTables,
    columnWidths,
    rowHeights,
  };
  if (tableStyleId !== undefined) info.tableStyleId = tableStyleId;
  const prev = prevParaId(tbl);
  if (prev !== undefined) info.prevParaId = prev;
  const next = nextParaId(tbl);
  if (next !== undefined) info.nextParaId = next;
  if (includeCellTexts) {
    const grid = extractCellGrid(tbl, drawingToId, tableToId);
    info.cellGrid = grid;
    info.cells = grid.map((row) => row.map((cell) => cell.text));
  }

  return info;
}

/**
 * 把 drawingIdx（id→Element）反转成 Element→id，供 cell 图 drawingId 反查
 * （task-bug2-cell-image, A1）。仅 includeCellTexts 时构建（避免无谓开销）。
 */
function buildDrawingReverseMap(drawingIdx: Map<string, Element>): Map<Element, string> {
  const rev = new Map<Element, string>();
  for (const [id, el] of drawingIdx) rev.set(el, id);
  return rev;
}

function buildTableReverseMap(tblIdx: Map<string, Element>): Map<Element, string> {
  const rev = new Map<Element, string>();
  for (const [id, el] of tblIdx) rev.set(el, id);
  return rev;
}

/** 找元素所属的 Story（用 contains 检测，stories 数量通常 < 20）*/
function findStoryForEl(el: Element, stories: Story[]): Story | undefined {
  for (const s of stories) {
    if (s.root.contains(el as unknown as Node)) return s;
  }
  return undefined;
}

// ─────────── 公开函数 ───────────

/**
 * 枚举文档表格，按文档顺序返回（与 tblIdx 插入顺序一致）。
 *
 * - `scope: 'body'`（默认）—— 仅主文档 <w:body>
 * - `includeNested: false`（默认）—— 跳过嵌套在其他表格内的表格
 */
export function listTables(
  pkg: Package,
  indices: PackageIndices,
  opts?: ListTablesOptions,
): TableInfo[] {
  const scope = opts?.scope ?? 'body';
  const includeNested = opts?.includeNested ?? false;
  const includeCellTexts = opts?.includeCellTexts ?? false;

  const stories = resolveScope(pkg, scope);
  const result: TableInfo[] = [];
  let idx = 1;

  // cell 图 drawingId 反查表：仅 includeCellTexts 时构建（A1）。
  const drawingToId = includeCellTexts
    ? buildDrawingReverseMap(indices.drawingIdx)
    : EMPTY_DRAWING_REV;
  const tableToId = includeCellTexts ? buildTableReverseMap(indices.tblIdx) : EMPTY_TABLE_REV;

  // 直接遍历 tblIdx（插入顺序 = 文档顺序）
  for (const [tableId, tbl] of indices.tblIdx) {
    if (!includeNested && hasTableAncestor(tbl)) continue;

    const story = findStoryForEl(tbl, stories);
    if (!story) continue; // 不在请求的 scope 内

    result.push(
      analyzeTable(tbl, tableId, story.partPath, idx++, includeCellTexts, drawingToId, tableToId),
    );
  }

  return result;
}

/** includeCellTexts=false 时复用的空反查表（不分配）。 */
const EMPTY_DRAWING_REV: Map<Element, string> = new Map();
const EMPTY_TABLE_REV: Map<Element, string> = new Map();

/**
 * 按 tableId 精确取单个表格，O(1)。不存在返回 null。
 * 入参 `includeCellTexts` 与 `listTables` 同义（默认 false）。
 */
export function getTableById(
  pkg: Package,
  indices: PackageIndices,
  tableId: string,
  opts?: { includeCellTexts?: boolean },
): TableInfo | null {
  const tbl = indices.tblIdx.get(tableId);
  if (!tbl) return null;

  // 找所属 story 以获取 partPath
  const stories = resolveScope(pkg, 'all');
  const story = findStoryForEl(tbl, stories);
  if (!story) return null;

  const includeCellTexts = opts?.includeCellTexts ?? false;
  const drawingToId = includeCellTexts
    ? buildDrawingReverseMap(indices.drawingIdx)
    : EMPTY_DRAWING_REV;
  const tableToId = includeCellTexts ? buildTableReverseMap(indices.tblIdx) : EMPTY_TABLE_REV;

  return analyzeTable(tbl, tableId, story.partPath, 0, includeCellTexts, drawingToId, tableToId);
}

// ─────────── getTableVerbatim（task-20260616-table-content-driven）───────────

/** 一张格内图的预解析字节（按源 <a:blip @r:embed> 原始 rId 索引）。 */
export interface TableVerbatimImage {
  originalRId: string;
  bytes: Uint8Array;
  contentType: string;
}

/** verbatim 取一张表的自包含 payload（供跨文档搬表 insertTableFrom）。 */
export interface TableVerbatim {
  /** `<w:tbl>` 子树序列化字符串（命名空间随子树带入）。 */
  tableXml: string;
  /** 表内所有格内图的预解析字节（按原 rId）。原 rId 解析不到字节的图不收（搬入时 warn 降级）。 */
  images: TableVerbatimImage[];
}

/**
 * 按 tableId verbatim 取一张表的 OOXML 子树 + 其格内图字节，打包成自包含 payload。
 *
 * 这是"数据表 = 内容块"（§13.9）的**源侧读取**：调用方（typeset）拿到此 payload 后用
 * `plan.insertTableFrom(...)` 把整表 verbatim 搬进输出文档（与图片 getImageBytes→insertImageAfter
 * 同构）。表的行/列/值/合并/嵌套/格内图全保真——这里只做读取 + 序列化 + 解析格内图字节，
 * **不裁剪、不重建**。
 *
 * tableId 不存在 → 返回 null。
 */
export function getTableVerbatim(
  pkg: Package,
  indices: PackageIndices,
  tableId: string,
): TableVerbatim | null {
  const tbl = indices.tblIdx.get(tableId);
  if (!tbl) return null;

  // 序列化 <w:tbl> 子树（命名空间随子树带入，已验证 importNode 往返保真）。
  const tableXml = serializeToWellFormedString(tbl as never);

  // 收集表内所有 <a:blip @r:embed>，按原 rId 解析出原始字节。
  const story = findStoryFor(pkg, tbl);
  const hostPartPath = story?.partPath;
  const images: TableVerbatimImage[] = [];
  if (hostPartPath) {
    const seen = new Set<string>();
    const blips = tbl.getElementsByTagNameNS(A_NS, 'blip');
    for (let i = 0; i < blips.length; i++) {
      const blip = blips[i];
      if (!blip) continue;
      const rId = blip.getAttributeNS(R_NS, 'embed');
      if (!rId || seen.has(rId)) continue;
      seen.add(rId);
      const resolved = resolveImageBytesByRId(pkg, hostPartPath, rId);
      if (resolved) {
        images.push({ originalRId: rId, bytes: resolved.bytes, contentType: resolved.contentType });
      }
      // 解析不到（如外链图/损坏 rels）→ 不收；搬入时该 blip embed 清掉 + warn（结构/文字保真）。
    }
  }

  return { tableXml, images };
}

/**
 * 按宿主 part + rId 解析图片字节（rels → media part）。
 * 与 insert-image-ops.getImageBytesSync 的 rels/media 解析同源，但入口是 rId（非 drawingId），
 * 供 getTableVerbatim 直接拿格内图的原始字节。
 */
function resolveImageBytesByRId(
  pkg: Package,
  hostPartPath: string,
  rId: string,
): { bytes: Uint8Array; contentType: string } | null {
  const relsPath = relsPathFor(hostPartPath);
  const relsPart = getPart(pkg, relsPath);
  if (!relsPart || !isXmlPart(relsPart) || !relsPart.xmlDom) return null;
  const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const rels = relsPart.xmlDom.getElementsByTagNameNS(RELS_NS, 'Relationship');
  let mediaPath: string | null = null;
  for (let i = 0; i < rels.length; i++) {
    const r = rels[i]!;
    if (r.getAttribute('Id') === rId) {
      mediaPath = resolveRelTarget(relsPath, r.getAttribute('Target') ?? '');
      break;
    }
  }
  if (!mediaPath) return null;
  const mediaPart = getPart(pkg, mediaPath);
  if (!mediaPart) return null;
  const bytes = (mediaPart as unknown as { bytes?: Uint8Array }).bytes;
  if (!bytes) return null;
  return { bytes, contentType: mediaPart.contentType };
}

/** part 路径 → 对应 .rels 路径（'word/document.xml' → 'word/_rels/document.xml.rels'）。 */
function relsPathFor(partPath: string): string {
  const idx = partPath.lastIndexOf('/');
  if (idx < 0) return `_rels/${partPath}.rels`;
  const dir = partPath.slice(0, idx);
  const name = partPath.slice(idx + 1);
  return `${dir}/_rels/${name}.rels`;
}

/** rels Target → 绝对 package 路径（处理 '/' 前缀、'./'、'../' 回溯）。 */
function resolveRelTarget(relsPartPath: string, target: string): string {
  let t = target.trim();
  if (t.startsWith('/')) return t.slice(1);
  while (t.startsWith('./')) t = t.slice(2);
  const idx = relsPartPath.lastIndexOf('/_rels/');
  if (idx < 0) return t;
  const hostDir = relsPartPath.slice(0, idx);
  let parts = hostDir.split('/').filter(Boolean);
  while (t.startsWith('../')) {
    parts = parts.slice(0, -1);
    t = t.slice(3);
  }
  return parts.length === 0 ? t : `${parts.join('/')}/${t}`;
}

/**
 * 装载期构建 tblIdx：枚举 Package 中所有内容 part 的 <w:tbl>，
 * 按文档顺序分配 'tbl-NNNN' ID。由 buildIndices 调用。
 *
 * 内部函数，不对外导出。
 */
export function buildTableIndex(
  parts: Iterable<{
    contentType: string;
    xmlDom?: {
      getElementsByTagNameNS(
        ns: string,
        local: string,
      ): { length: number } & Record<number, Element>;
    };
  }>,
): Map<string, Element> {
  const tblIdx = new Map<string, Element>();
  let n = 1;

  for (const part of parts) {
    const ct = part.contentType;
    // 只扫 Word 内容 part（同 buildIndices 的 isWordContentPart 逻辑）
    if (
      !ct.includes('wordprocessingml') &&
      !ct.endsWith('.main+xml') &&
      !ct.endsWith('.header+xml') &&
      !ct.endsWith('.footer+xml') &&
      !ct.endsWith('.footnotes+xml') &&
      !ct.endsWith('.endnotes+xml')
      // 不扫 comments.xml：注释内通常无表格，跳过
    ) {
      continue;
    }
    if (!part.xmlDom) continue;

    const tbls = part.xmlDom.getElementsByTagNameNS(W_NS, 'tbl');
    for (let i = 0; i < tbls.length; i++) {
      const tbl = tbls[i];
      if (!tbl) continue;
      const id = `tbl-${String(n).padStart(4, '0')}`;
      tblIdx.set(id, tbl);
      n++;
    }
  }

  return tblIdx;
}
