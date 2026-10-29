import { columnIndexToName, columnNameToIndex, splitSheetQualifiedAddress } from './address.js';

export interface ExcelReferenceRewriteOptions {
  sheetName?: string;
}

export type ExcelReferenceRewriteResult =
  | { ok: true; address: string }
  | { ok: false; address: '#REF!' };

export type ExcelFormulaReferenceRewriter = (reference: string) => ExcelReferenceRewriteResult;

interface CellRef {
  row: number;
  column: number;
  rowAbsolute?: boolean;
  columnAbsolute?: boolean;
}

interface RangeRef {
  sheetName?: string;
  start: CellRef;
  end: CellRef;
}

interface AxisDelete {
  start: number;
  count: number;
}

export function renameSheetInReference(address: string, oldName: string, newName: string): string {
  const ref = parseReference(address);
  if (ref.sheetName !== oldName) return formatReference(ref);
  return formatReference({ ...ref, sheetName: newName });
}

export function rewriteFormulaReferences(
  formula: string,
  rewrite: ExcelFormulaReferenceRewriter,
): string {
  let output = '';
  let index = 0;
  while (index < formula.length) {
    const ch = formula[index];
    if (ch === '"') {
      const literal = readFormulaStringLiteral(formula, index);
      output += literal.text;
      index = literal.end;
      continue;
    }

    const match = matchReferenceAt(formula, index);
    if (!match) {
      output += ch;
      index += 1;
      continue;
    }

    const rewritten = rewrite(match.reference);
    output += rewritten.ok ? rewritten.address : rewritten.address;
    index = match.end;
  }
  return output;
}

export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  return rewriteFormulaReferences(formula, (reference) => ({
    ok: true,
    address: renameSheetInReference(reference, oldName, newName),
  }));
}

export function rewriteFormulaForRowInsert(
  formula: string,
  rowIndex: number,
  count = 1,
  options: ExcelReferenceRewriteOptions = {},
): string {
  return rewriteFormulaReferences(formula, (reference) =>
    rewriteReferenceForRowInsert(reference, rowIndex, count, options),
  );
}

export function rewriteFormulaForRowDelete(
  formula: string,
  rowIndex: number,
  count = 1,
  options: ExcelReferenceRewriteOptions = {},
): string {
  return rewriteFormulaReferences(formula, (reference) =>
    rewriteReferenceForRowDelete(reference, rowIndex, count, options),
  );
}

export function rewriteFormulaForColumnInsert(
  formula: string,
  columnIndex: number,
  count = 1,
  options: ExcelReferenceRewriteOptions = {},
): string {
  return rewriteFormulaReferences(formula, (reference) =>
    rewriteReferenceForColumnInsert(reference, columnIndex, count, options),
  );
}

export function rewriteFormulaForColumnDelete(
  formula: string,
  columnIndex: number,
  count = 1,
  options: ExcelReferenceRewriteOptions = {},
): string {
  return rewriteFormulaReferences(formula, (reference) =>
    rewriteReferenceForColumnDelete(reference, columnIndex, count, options),
  );
}

export function rewriteFormulaForRowMove(
  formula: string,
  rowIndex: number,
  count: number,
  beforeRowIndex: number,
  options: ExcelReferenceRewriteOptions = {},
): string {
  return rewriteFormulaReferences(formula, (reference) =>
    rewriteReferenceForRowMove(reference, rowIndex, count, beforeRowIndex, options),
  );
}

export function rewriteFormulaForColumnMove(
  formula: string,
  columnIndex: number,
  count: number,
  beforeColumnIndex: number,
  options: ExcelReferenceRewriteOptions = {},
): string {
  return rewriteFormulaReferences(formula, (reference) =>
    rewriteReferenceForColumnMove(reference, columnIndex, count, beforeColumnIndex, options),
  );
}

/**
 * Rebase a formula as if its containing cell moved by the supplied offsets.
 * Relative row/column coordinates move; absolute coordinates and string literals stay unchanged.
 */
export function shiftFormulaReferences(
  formula: string,
  rowOffset: number,
  columnOffset: number,
): string {
  assertInteger(rowOffset, 'rowOffset');
  assertInteger(columnOffset, 'columnOffset');
  return rewriteFormulaReferences(formula, (reference) => shiftReferenceRelative(reference, rowOffset, columnOffset));
}

/** Rewrites relative row references when a formula's source rows are reordered. */
export function rewriteFormulaForRowSort(
  formula: string,
  rowMap: ReadonlyMap<number, number>,
  options: ExcelReferenceRewriteOptions = {},
): string {
  return rewriteFormulaReferences(formula, (reference) => {
    const ref = parseReference(reference);
    if (!isTargetSheet(ref, options.sheetName)) return { ok: true, address: formatReference(ref) };
    const rewriteCell = (cell: CellRef): CellRef => {
      const row = cell.rowAbsolute ? cell.row : rowMap.get(cell.row) ?? cell.row;
      return { ...cell, row };
    };
    return {
      ok: true,
      address: formatReference({ ...ref, start: rewriteCell(ref.start), end: rewriteCell(ref.end) }),
    };
  });
}

export function shiftReferenceRelative(
  address: string,
  rowOffset: number,
  columnOffset: number,
): ExcelReferenceRewriteResult {
  assertInteger(rowOffset, 'rowOffset');
  assertInteger(columnOffset, 'columnOffset');
  const ref = parseReference(address);
  const start = shiftCellRelative(ref.start, rowOffset, columnOffset);
  const end = shiftCellRelative(ref.end, rowOffset, columnOffset);
  if (!start || !end) return { ok: false, address: '#REF!' };
  return { ok: true, address: formatReference({ ...ref, start, end }) };
}

export function rewriteReferenceForRowInsert(
  address: string,
  rowIndex: number,
  count = 1,
  options: ExcelReferenceRewriteOptions = {},
): ExcelReferenceRewriteResult {
  assertPositiveInteger(rowIndex, 'rowIndex');
  assertPositiveInteger(count, 'count');
  const ref = parseReference(address);
  if (!isTargetSheet(ref, options.sheetName)) return { ok: true, address: formatReference(ref) };
  return {
    ok: true,
    address: formatReference({
      ...ref,
      start: shiftCellRowsForInsert(ref.start, rowIndex, count),
      end: shiftCellRowsForInsert(ref.end, rowIndex, count),
    }),
  };
}

export function rewriteReferenceForRowDelete(
  address: string,
  rowIndex: number,
  count = 1,
  options: ExcelReferenceRewriteOptions = {},
): ExcelReferenceRewriteResult {
  assertPositiveInteger(rowIndex, 'rowIndex');
  assertPositiveInteger(count, 'count');
  const ref = parseReference(address);
  if (!isTargetSheet(ref, options.sheetName)) return { ok: true, address: formatReference(ref) };
  const rows = rewriteAxisForDelete(ref.start.row, ref.end.row, { start: rowIndex, count });
  if (!rows) return { ok: false, address: '#REF!' };
  return {
    ok: true,
    address: formatReference({
      ...ref,
      start: { ...ref.start, row: rows.start },
      end: { ...ref.end, row: rows.end },
    }),
  };
}

export function rewriteReferenceForColumnInsert(
  address: string,
  columnIndex: number,
  count = 1,
  options: ExcelReferenceRewriteOptions = {},
): ExcelReferenceRewriteResult {
  assertPositiveInteger(columnIndex, 'columnIndex');
  assertPositiveInteger(count, 'count');
  const ref = parseReference(address);
  if (!isTargetSheet(ref, options.sheetName)) return { ok: true, address: formatReference(ref) };
  return {
    ok: true,
    address: formatReference({
      ...ref,
      start: shiftCellColumnsForInsert(ref.start, columnIndex, count),
      end: shiftCellColumnsForInsert(ref.end, columnIndex, count),
    }),
  };
}

export function rewriteReferenceForColumnDelete(
  address: string,
  columnIndex: number,
  count = 1,
  options: ExcelReferenceRewriteOptions = {},
): ExcelReferenceRewriteResult {
  assertPositiveInteger(columnIndex, 'columnIndex');
  assertPositiveInteger(count, 'count');
  const ref = parseReference(address);
  if (!isTargetSheet(ref, options.sheetName)) return { ok: true, address: formatReference(ref) };
  const columns = rewriteAxisForDelete(ref.start.column, ref.end.column, {
    start: columnIndex,
    count,
  });
  if (!columns) return { ok: false, address: '#REF!' };
  return {
    ok: true,
    address: formatReference({
      ...ref,
      start: { ...ref.start, column: columns.start },
      end: { ...ref.end, column: columns.end },
    }),
  };
}

export function rewriteReferenceForRowMove(
  address: string,
  rowIndex: number,
  count: number,
  beforeRowIndex: number,
  options: ExcelReferenceRewriteOptions = {},
): ExcelReferenceRewriteResult {
  assertMoveArguments(rowIndex, count, beforeRowIndex, 'row');
  const ref = parseReference(address);
  if (!isTargetSheet(ref, options.sheetName)) return { ok: true, address: formatReference(ref) };
  const rows = rewriteAxisForMove(ref.start.row, ref.end.row, rowIndex, count, beforeRowIndex);
  if (!rows) return { ok: false, address: '#REF!' };
  return {
    ok: true,
    address: formatReference({
      ...ref,
      start: { ...ref.start, row: rows.start },
      end: { ...ref.end, row: rows.end },
    }),
  };
}

export function rewriteReferenceForColumnMove(
  address: string,
  columnIndex: number,
  count: number,
  beforeColumnIndex: number,
  options: ExcelReferenceRewriteOptions = {},
): ExcelReferenceRewriteResult {
  assertMoveArguments(columnIndex, count, beforeColumnIndex, 'column');
  const ref = parseReference(address);
  if (!isTargetSheet(ref, options.sheetName)) return { ok: true, address: formatReference(ref) };
  const columns = rewriteAxisForMove(
    ref.start.column,
    ref.end.column,
    columnIndex,
    count,
    beforeColumnIndex,
  );
  if (!columns) return { ok: false, address: '#REF!' };
  return {
    ok: true,
    address: formatReference({
      ...ref,
      start: { ...ref.start, column: columns.start },
      end: { ...ref.end, column: columns.end },
    }),
  };
}

function parseReference(address: string): RangeRef {
  const qualified = splitSheetQualifiedAddress(address);
  const parts = qualified.address.split(':');
  if (parts.length > 2 || !parts[0]) {
    throw new Error(`ExcelReference: invalid reference ${address}`);
  }
  const start = parseCellRef(parts[0]);
  const end = parseCellRef(parts[1] ?? parts[0]);
  const reversed = start.row > end.row || start.column > end.column;
  return {
    sheetName: qualified.sheetName,
    start: {
      row: Math.min(start.row, end.row),
      column: Math.min(start.column, end.column),
      rowAbsolute: reversed ? false : start.rowAbsolute,
      columnAbsolute: reversed ? false : start.columnAbsolute,
    },
    end: {
      row: Math.max(start.row, end.row),
      column: Math.max(start.column, end.column),
      rowAbsolute: reversed ? false : end.rowAbsolute,
      columnAbsolute: reversed ? false : end.columnAbsolute,
    },
  };
}

function parseCellRef(address: string): CellRef {
  const normalized = address.trim();
  const match = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(normalized);
  if (!match) {
    throw new Error(`ExcelReference: invalid cell reference ${address}`);
  }
  const row = Number(match[4]);
  if (!Number.isInteger(row) || row < 1) {
    throw new Error(`ExcelReference: invalid cell reference ${address}`);
  }
  return {
    row,
    column: columnNameToIndex(match[2]!),
    columnAbsolute: match[1] === '$',
    rowAbsolute: match[3] === '$',
  };
}

function formatReference(ref: RangeRef): string {
  const first = formatCellRef(ref.start);
  const last = formatCellRef(ref.end);
  const local = first === last ? first : `${first}:${last}`;
  return ref.sheetName ? `${quoteSheetName(ref.sheetName)}!${local}` : local;
}

function formatCellRef(ref: CellRef): string {
  return `${ref.columnAbsolute ? '$' : ''}${columnIndexToName(ref.column)}${ref.rowAbsolute ? '$' : ''}${ref.row}`;
}

function quoteSheetName(sheetName: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)) return sheetName;
  return `'${sheetName.replace(/'/g, "''")}'`;
}

function shiftCellRowsForInsert(cell: CellRef, rowIndex: number, count: number): CellRef {
  return cell.row >= rowIndex ? { ...cell, row: cell.row + count } : { ...cell };
}

function shiftCellColumnsForInsert(cell: CellRef, columnIndex: number, count: number): CellRef {
  return cell.column >= columnIndex ? { ...cell, column: cell.column + count } : { ...cell };
}

function shiftCellRelative(cell: CellRef, rowOffset: number, columnOffset: number): CellRef | null {
  const row = cell.rowAbsolute ? cell.row : cell.row + rowOffset;
  const column = cell.columnAbsolute ? cell.column : cell.column + columnOffset;
  if (row < 1 || column < 1 || column > 16384) return null;
  return { ...cell, row, column };
}

function rewriteAxisForDelete(start: number, end: number, deletion: AxisDelete): { start: number; end: number } | null {
  const deleteEnd = deletion.start + deletion.count - 1;
  const survivors: number[] = [];
  for (let value = start; value <= end; value++) {
    if (value >= deletion.start && value <= deleteEnd) continue;
    survivors.push(value > deleteEnd ? value - deletion.count : value);
  }
  if (survivors.length === 0) return null;
  return {
    start: Math.min(...survivors),
    end: Math.max(...survivors),
  };
}

export function moveAxisIndex(
  value: number,
  sourceIndex: number,
  count: number,
  beforeIndex: number,
): number {
  const sourceEnd = sourceIndex + count - 1;
  if (beforeIndex >= sourceIndex && beforeIndex <= sourceEnd + 1) return value;
  if (beforeIndex < sourceIndex) {
    if (value >= sourceIndex && value <= sourceEnd) return beforeIndex + value - sourceIndex;
    if (value >= beforeIndex && value < sourceIndex) return value + count;
    return value;
  }
  if (value >= sourceIndex && value <= sourceEnd) {
    return beforeIndex - count + value - sourceIndex;
  }
  if (value > sourceEnd && value < beforeIndex) return value - count;
  return value;
}

function rewriteAxisForMove(
  start: number,
  end: number,
  sourceIndex: number,
  count: number,
  beforeIndex: number,
): { start: number; end: number } | null {
  const moved = new Set<number>();
  for (let value = start; value <= end; value += 1) {
    moved.add(moveAxisIndex(value, sourceIndex, count, beforeIndex));
  }
  const values = [...moved].sort((a, b) => a - b);
  if (values.length === 0) return null;
  const first = values[0]!;
  const last = values[values.length - 1]!;
  if (last - first + 1 !== values.length) return null;
  return { start: first, end: last };
}

function assertMoveArguments(
  sourceIndex: number,
  count: number,
  beforeIndex: number,
  axis: 'row' | 'column',
): void {
  assertPositiveInteger(sourceIndex, `${axis}Index`);
  assertPositiveInteger(count, 'count');
  assertPositiveInteger(beforeIndex, `before${axis === 'row' ? 'Row' : 'Column'}Index`);
}

function isTargetSheet(ref: RangeRef, sheetName: string | undefined): boolean {
  return sheetName === undefined || ref.sheetName === undefined || ref.sheetName === sheetName;
}

function assertPositiveInteger(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`ExcelReference: ${name} must be a positive integer`);
  }
}

function assertInteger(value: number, name: string): void {
  if (!Number.isInteger(value)) {
    throw new Error(`ExcelReference: ${name} must be an integer`);
  }
}

interface FormulaMatch {
  reference: string;
  end: number;
}

function readFormulaStringLiteral(formula: string, start: number): { text: string; end: number } {
  let index = start + 1;
  while (index < formula.length) {
    if (formula[index] !== '"') {
      index += 1;
      continue;
    }
    if (formula[index + 1] === '"') {
      index += 2;
      continue;
    }
    return { text: formula.slice(start, index + 1), end: index + 1 };
  }
  return { text: formula.slice(start), end: formula.length };
}

function matchReferenceAt(formula: string, start: number): FormulaMatch | null {
  if (!isReferenceLeftBoundary(formula[start - 1])) return null;
  const sheet = readSheetPrefix(formula, start);
  const cellStart = sheet ? sheet.end : start;
  const firstCell = readCellToken(formula, cellStart);
  if (!firstCell) return null;

  let end = firstCell.end;
  if (formula[end] === ':') {
    const secondCell = readCellToken(formula, end + 1);
    if (secondCell) end = secondCell.end;
  }
  if (!isReferenceRightBoundary(formula[end])) return null;
  if (formula[end] === '(') return null;

  return { reference: formula.slice(start, end), end };
}

function readSheetPrefix(formula: string, start: number): { end: number } | null {
  if (formula[start] === "'") {
    let index = start + 1;
    while (index < formula.length) {
      if (formula[index] !== "'") {
        index += 1;
        continue;
      }
      if (formula[index + 1] === "'") {
        index += 2;
        continue;
      }
      return formula[index + 1] === '!' ? { end: index + 2 } : null;
    }
    return null;
  }

  const match = /^[A-Za-z_][A-Za-z0-9_.]*!/.exec(formula.slice(start));
  return match ? { end: start + match[0].length } : null;
}

function readCellToken(formula: string, start: number): { end: number } | null {
  const match = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*/.exec(formula.slice(start));
  if (!match) return null;
  const token = match[0];
  const normalized = token.replace(/\$/g, '');
  const cell = /^([A-Za-z]+)([1-9][0-9]*)$/.exec(normalized);
  if (!cell) return null;
  const column = columnNameToIndex(cell[1]!);
  if (column > 16384) return null;
  return { end: start + token.length };
}

function isReferenceLeftBoundary(ch: string | undefined): boolean {
  return ch === undefined || !/[A-Za-z0-9_.$]/.test(ch);
}

function isReferenceRightBoundary(ch: string | undefined): boolean {
  return ch === undefined || !/[A-Za-z0-9_.$]/.test(ch);
}
