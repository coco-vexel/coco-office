import { formatCellAddress, formatRangeAddress, type ExcelCellValue } from '@coco-vexel/coco-office/excel';
import type { SpreadsheetRange, SpreadsheetSheetRenderModel } from './types.js';
import { getRenderCell } from './render-model.js';

export interface SpreadsheetClipboardMerge {
  startRowOffset: number;
  startColumnOffset: number;
  rowCount: number;
  columnCount: number;
}

export interface SpreadsheetClipboardPayload {
  text: string;
  values: ExcelCellValue[][];
  styleIds: (number | undefined)[][];
  merges: SpreadsheetClipboardMerge[];
}

export function buildSpreadsheetClipboardPayload(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
): SpreadsheetClipboardPayload {
  const text = serializeRangeToTsv(model, range);
  return {
    text,
    values: parseTsv(text),
    styleIds: rangeStyleIdsForClipboard(model, range),
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
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .filter((line, index, lines) => !(index === lines.length - 1 && line === ''))
    .map((line) => line.split('\t').map(coerceClipboardValue));
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
  return value.replace(/\t/g, ' ').replace(/\r?\n/g, ' ');
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
  if (value === 'TRUE') return true;
  if (value === 'FALSE') return false;
  return value;
}
