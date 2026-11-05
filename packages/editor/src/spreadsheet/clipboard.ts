import { captureExcelCellStyle, formatCellAddress, formatRangeAddress, shiftFormulaReferences, type ExcelCellStyleSnapshot, type ExcelCellValue } from '@coco-vexel/coco-office/excel';
import { serializeSpreadsheetClipboardPayload } from './clipboard-serialization.js';
export { clipboardToken, readSpreadsheetClipboardPayload, serializeSpreadsheetClipboardPayload, SPREADSHEET_CLIPBOARD_MIME } from './clipboard-serialization.js';
import type { SpreadsheetRange, SpreadsheetSheetRenderModel } from './types.js';
import { getRenderCell } from './render-model.js';

export interface SpreadsheetClipboardMerge {
  startRowOffset: number;
  startColumnOffset: number;
  rowCount: number;
  columnCount: number;
}

export interface SpreadsheetClipboardPayload {
  schema?: 1;
  id: string;
  sourceWorkbookId?: string;
  sourceSheet: string;
  sourceRange: SpreadsheetRange;
  text: string;
  values: ExcelCellValue[][];
  formulas: (string | null)[][];
  styleIds: (number | undefined)[][];
  merges: SpreadsheetClipboardMerge[];
  /** Portable definitions keyed by source style id; `default` represents an implicit style. */
  styles?: Record<string, ExcelCellStyleSnapshot>;
}

export function buildSpreadsheetClipboardPayload(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
  options: { sourceWorkbookId?: string } = {},
): SpreadsheetClipboardPayload {
  const text = serializeRangeToTsv(model, range);
  const styleIds = rangeStyleIdsForClipboard(model, range);
  const styles: Record<string, ExcelCellStyleSnapshot> = {};
  for (const row of styleIds) for (const id of row) {
    const key = id === undefined ? 'default' : String(id);
    styles[key] ??= captureExcelCellStyle(model.styles, id);
  }
  return {
    schema: 1,
    id: globalThis.crypto?.randomUUID?.() ?? `clipboard-${Date.now()}-${Math.random()}`,
    ...(options.sourceWorkbookId ? { sourceWorkbookId: options.sourceWorkbookId } : {}),
    sourceSheet: model.source.name,
    sourceRange: { ...range },
    text,
    values: rangeValuesForClipboard(model, range),
    formulas: Array.from({ length: range.endRow - range.startRow + 1 }, (_, r) =>
      Array.from({ length: range.endColumn - range.startColumn + 1 }, (_, c) => {
        const row = range.startRow + r; const column = range.startColumn + c;
        return model.mergeFollowerByAddress.has(formatCellAddress(row, column)) ? null : getRenderCell(model, row, column)?.formula ?? null;
      })),
    styleIds,
    styles,
    merges: mergedRangesForClipboard(model, range),
  };
}

export function serializeRangeToTsv(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
): string {
  const lines: string[] = [];
  for (let row = range.startRow; row <= range.endRow; row++) {
    const values: string[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      values.push(escapeTsvValue(clipboardCellText(model, row, column)));
    }
    lines.push(values.join('\t'));
  }
  return lines.join('\n');
}

export function parseTsv(text: string): ExcelCellValue[][] {
  if (text === '') return [[null]];
  const rows: ExcelCellValue[][] = [];
  let row: ExcelCellValue[] = []; let value = ''; let quoted = false; let endedRow = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    endedRow = false;
    if (ch === '"') {
      if (quoted && text[i + 1] === '"') { value += '"'; i++; }
      else if (quoted || value === '') quoted = !quoted;
      else value += ch;
    } else if (!quoted && (ch === '\t' || ch === '\n' || ch === '\r')) {
      row.push(coerceClipboardValue(value)); value = '';
      if (ch !== '\t') {
        rows.push(row); row = []; endedRow = true;
        if (ch === '\r' && text[i + 1] === '\n') i++;
      }
    } else value += ch;
  }
  if (quoted) throw new Error('粘贴数据包含未闭合的引号。');
  if (!endedRow) { row.push(coerceClipboardValue(value)); rows.push(row); }
  return rows;
}

export function clipboardHtml(payload: SpreadsheetClipboardPayload): string {
  const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const css = (r: number, c: number) => {
    const style = payload.styles?.[String(payload.styleIds[r]?.[c] ?? 'default')];
    if (!style) return '';
    const parts: string[] = [];
    if (style.font.name) parts.push(`font-family:'${style.font.name.replace(/['\\]/g, '\\$&')}'`);
    if (style.font.size) parts.push(`font-size:${style.font.size}pt`);
    if (style.font.bold) parts.push('font-weight:bold');
    if (style.font.italic) parts.push('font-style:italic');
    if (style.font.underline) parts.push('text-decoration:underline');
    if (style.font.color) parts.push(`color:${style.font.color.rgb}`);
    if (style.fill.foregroundColor) parts.push(`background-color:${style.fill.foregroundColor.rgb}`);
    if (style.alignment?.horizontal) parts.push(`text-align:${style.alignment.horizontal}`);
    if (style.alignment?.vertical) parts.push(`vertical-align:${style.alignment.vertical === 'center' ? 'middle' : style.alignment.vertical}`);
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const border = style.border[side];
      if (border?.style && border.style !== 'none') parts.push(`border-${side}:${border.style === 'thick' ? 3 : border.style === 'medium' ? 2 : 1}px ${border.style === 'double' ? 'double' : /dash/i.test(border.style) ? 'dashed' : /dot/i.test(border.style) ? 'dotted' : 'solid'} ${border.color?.rgb ?? '#000000'}`);
    }
    return parts.join(';');
  };
  return `<table data-coco-spreadsheet="${escape(payload.id)}" data-coco-spreadsheet-payload="${encodeURIComponent(serializeSpreadsheetClipboardPayload(payload))}">${payload.values.map((row, r) => `<tr>${row.map((value, c) => `<td style="${escape(css(r, c))}">${escape(String(value ?? ''))}</td>`).join('')}</tr>`).join('')}</table>`;
}

export function formulasForPaste(payload: SpreadsheetClipboardPayload, row: number, column: number, cut = false): (string | null)[][] {
  return payload.formulas.map((line) => line.map((formula) => formula === null || cut ? formula :
    shiftFormulaReferences(formula, row - payload.sourceRange.startRow, column - payload.sourceRange.startColumn)));
}

export function rangeAddressForMatrix(
  startRow: number,
  startColumn: number,
  values: ExcelCellValue[][],
): string {
  const rowCount = Math.max(1, values.length);
  const columnCount = Math.max(1, values.reduce((max, row) => Math.max(max, row.length), 0));
  return formatRangeAddress(startRow, startColumn, rowCount, columnCount);
}

function escapeTsvValue(value: string): string {
  return /[\t\r\n"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

function clipboardCellText(
  model: SpreadsheetSheetRenderModel,
  row: number,
  column: number,
): string {
  const address = formatCellAddress(row, column);
  // Covered cells may still contain legacy values in imported workbooks.
  // Copy only the merge master so paste/unmerge does not surface hidden stale content.
  if (model.mergeFollowerByAddress.has(address)) return '';
  return getRenderCell(model, row, column)?.displayText ?? '';
}

function rangeValuesForClipboard(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
): ExcelCellValue[][] {
  const values: ExcelCellValue[][] = [];
  for (let row = range.startRow; row <= range.endRow; row++) {
    const line: ExcelCellValue[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      const address = formatCellAddress(row, column);
      line.push(model.mergeFollowerByAddress.has(address) ? null : getRenderCell(model, row, column)?.value ?? null);
    }
    values.push(line);
  }
  return values;
}

function mergedRangesForClipboard(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
): SpreadsheetClipboardMerge[] {
  const merges: SpreadsheetClipboardMerge[] = [];
  for (const merge of model.mergeMasterByAddress.values()) {
    if (!rangeContainsRange(range, merge.range)) continue;
    merges.push({
      startRowOffset: merge.range.startRow - range.startRow,
      startColumnOffset: merge.range.startColumn - range.startColumn,
      rowCount: merge.range.endRow - merge.range.startRow + 1,
      columnCount: merge.range.endColumn - merge.range.startColumn + 1,
    });
  }
  return merges;
}

function rangeStyleIdsForClipboard(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
): (number | undefined)[][] {
  const styleIds: (number | undefined)[][] = [];
  for (let row = range.startRow; row <= range.endRow; row++) {
    const line: (number | undefined)[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      line.push(getRenderCell(model, row, column)?.styleId);
    }
    styleIds.push(line);
  }
  return styleIds;
}

function rangeContainsRange(outer: SpreadsheetRange, inner: SpreadsheetRange): boolean {
  return (
    inner.startRow >= outer.startRow &&
    inner.endRow <= outer.endRow &&
    inner.startColumn >= outer.startColumn &&
    inner.endColumn <= outer.endColumn
  );
}

function coerceClipboardValue(value: string): ExcelCellValue {
  if (value === '') return null;
  const number = Number(value);
  if (Number.isFinite(number) && value.trim() !== '') return number;
  if (value.toUpperCase() === 'TRUE') return true;
  if (value.toUpperCase() === 'FALSE') return false;
  return value;
}
