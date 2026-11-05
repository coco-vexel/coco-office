import { columnIndexToName, columnNameToIndex, parseRangeAddress, rewriteFormulaReferences, splitSheetQualifiedAddress } from '@coco-vexel/coco-office/excel';
import type { SpreadsheetRange } from './types.js';

export interface SpreadsheetCellMove {
  sourceSheet: string; targetSheet: string; source: SpreadsheetRange;
  rowOffset: number; columnOffset: number; formulaSheet: string; resultSheet: string;
}

/** Cut follows cell identity, including $ references; it never rebases a moved formula like a copy. */
export function rewriteMovedCellFormula(formula: string, options: SpreadsheetCellMove): string {
  const destination = translate(options.source, options.rowOffset, options.columnOffset);
  // Overlapping source cells are moved, not overwritten. Only the new destination cells are deleted.
  const overwritten = sameSheet(options.sourceSheet, options.targetSheet) ? subtract(destination, options.source) : [destination];
  return rewriteFormulaReferences(formula, (reference) => {
    const qualified = splitSheetQualifiedAddress(reference);
    const sheet = qualified.sheetName ?? options.formulaSheet;
    const parsed = parseRangeAddress(qualified.address);
    let range: SpreadsheetRange | null = {
      startRow: parsed.startRow, startColumn: parsed.startColumn,
      endRow: parsed.startRow + parsed.rowCount - 1, endColumn: parsed.startColumn + parsed.columnCount - 1,
    };
    const original = range;
    if (sameSheet(sheet, options.targetSheet)) {
      for (const deleted of overwritten) {
        range = removeOverwrittenEdge(range, deleted);
        if (!range) return { ok: false, address: '#REF!' };
      }
    }
    // A contained reference follows the move. A partially cut edge shrinks the rectangle;
    // an interior cut never turns a reference into a disconnected union.
    const moved = sameSheet(sheet, options.sourceSheet) && contains(options.source, range);
    const resultingSheet = moved ? options.targetSheet : sheet;
    if (moved) range = translate(range, options.rowOffset, options.columnOffset);
    else if (sameSheet(sheet, options.sourceSheet)) range = removeOverwrittenEdge(range, options.source)!;
    if (!moved && sameSheet(options.formulaSheet, options.resultSheet) &&
      range.startRow === original.startRow && range.endRow === original.endRow && range.startColumn === original.startColumn && range.endColumn === original.endColumn) return { ok: true, address: reference };
    const address = formatLike(qualified.address, range);
    const prefix = qualified.sheetName || resultingSheet !== options.resultSheet ? `'${resultingSheet.replace(/'/g, "''")}'!` : '';
    return { ok: true, address: `${prefix}${address}` };
  });
}

function sameSheet(left: string, right: string): boolean { return left.toLowerCase() === right.toLowerCase(); }

function contains(outer: SpreadsheetRange, inner: SpreadsheetRange): boolean {
  return outer.startRow <= inner.startRow && outer.endRow >= inner.endRow && outer.startColumn <= inner.startColumn && outer.endColumn >= inner.endColumn;
}

function translate(range: SpreadsheetRange, rows: number, columns: number): SpreadsheetRange {
  return { startRow: range.startRow + rows, endRow: range.endRow + rows, startColumn: range.startColumn + columns, endColumn: range.endColumn + columns };
}

function subtract(range: SpreadsheetRange, other: SpreadsheetRange): SpreadsheetRange[] {
  const top = Math.max(range.startRow, other.startRow); const bottom = Math.min(range.endRow, other.endRow);
  const left = Math.max(range.startColumn, other.startColumn); const right = Math.min(range.endColumn, other.endColumn);
  if (top > bottom || left > right) return [range];
  const pieces: SpreadsheetRange[] = [];
  if (range.startRow < top) pieces.push({ ...range, endRow: top - 1 });
  if (bottom < range.endRow) pieces.push({ ...range, startRow: bottom + 1 });
  if (range.startColumn < left) pieces.push({ startRow: top, endRow: bottom, startColumn: range.startColumn, endColumn: left - 1 });
  if (right < range.endColumn) pieces.push({ startRow: top, endRow: bottom, startColumn: right + 1, endColumn: range.endColumn });
  return pieces;
}

function removeOverwrittenEdge(range: SpreadsheetRange, deleted: SpreadsheetRange): SpreadsheetRange | null {
  if (contains(deleted, range)) return null;
  const result = { ...range };
  if (deleted.startColumn <= range.startColumn && deleted.endColumn >= range.endColumn) {
    if (deleted.startRow <= range.startRow && deleted.endRow >= range.startRow) result.startRow = deleted.endRow + 1;
    if (deleted.endRow >= range.endRow && deleted.startRow <= range.endRow) result.endRow = deleted.startRow - 1;
  }
  if (deleted.startRow <= range.startRow && deleted.endRow >= range.endRow) {
    if (deleted.startColumn <= range.startColumn && deleted.endColumn >= range.startColumn) result.startColumn = deleted.endColumn + 1;
    if (deleted.endColumn >= range.endColumn && deleted.startColumn <= range.endColumn) result.endColumn = deleted.startColumn - 1;
  }
  return result;
}

function formatLike(address: string, range: SpreadsheetRange): string {
  const endpoints = address.split(':').map((part) => /^(\$?)([A-Z]+)(\$?)(\d+)$/i.exec(part)!);
  const first = endpoints[0]!; const last = endpoints[1] ?? first;
  const rowReversed = Number(first[4]) > Number(last[4]);
  const colReversed = columnNameToIndex(first[2]!) > columnNameToIndex(last[2]!);
  const cell = (match: RegExpExecArray, row: number, column: number) => `${match[1]}${columnIndexToName(column)}${match[3]}${row}`;
  const start = cell(first, rowReversed ? range.endRow : range.startRow, colReversed ? range.endColumn : range.startColumn);
  if (endpoints.length === 1) return start;
  return `${start}:${cell(last, rowReversed ? range.startRow : range.endRow, colReversed ? range.startColumn : range.endColumn)}`;
}
