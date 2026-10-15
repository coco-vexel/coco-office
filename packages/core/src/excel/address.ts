export interface CellAddress {
  address: string;
  row: number;
  column: number;
}

export interface RangeAddress {
  address: string;
  sheetName?: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
}

export interface SheetQualifiedAddress {
  sheetName?: string;
  address: string;
}

export const EXCEL_MAX_ROW = 1_048_576;
export const EXCEL_MAX_COLUMN = 16_384;

export function columnNameToIndex(columnName: string): number {
  let value = 0;
  for (const ch of columnName.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) continue;
    value = value * 26 + (code - 64);
  }
  if (value <= 0) {
    throw new Error(`ExcelWorkbook: invalid column name ${columnName}`);
  }
  return value;
}

export function columnIndexToName(column: number): string {
  if (!Number.isInteger(column) || column < 1) {
    throw new Error(`ExcelWorkbook: invalid column index ${column}`);
  }
  let value = column;
  let name = '';
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

export function parseCellAddress(address: string): CellAddress {
  const qualified = splitSheetQualifiedAddress(address);
  const normalized = qualified.address.replace(/\$/g, '').trim();
  const match = /^([A-Za-z]+)(\d+)$/.exec(normalized);
  if (!match) {
    throw new Error(`ExcelWorkbook: invalid cell address ${address}`);
  }
  const column = columnNameToIndex(match[1]!);
  const row = Number(match[2]);
  if (!Number.isInteger(row) || row < 1) {
    throw new Error(`ExcelWorkbook: invalid cell address ${address}`);
  }
  return {
    address: `${columnIndexToName(column)}${row}`,
    column,
    row,
  };
}

export function formatCellAddress(row: number, column: number): string {
  if (!Number.isInteger(row) || row < 1) {
    throw new Error(`ExcelWorkbook: invalid row index ${row}`);
  }
  return `${columnIndexToName(column)}${row}`;
}

export function parseRangeAddress(address: string): RangeAddress {
  const qualified = splitSheetQualifiedAddress(address);
  const parts = qualified.address.split(':');
  if (parts.length > 2) {
    throw new Error(`ExcelWorkbook: invalid range address ${address}`);
  }
  const [startRaw, endRaw] = parts;
  if (!startRaw) {
    throw new Error(`ExcelWorkbook: invalid range address ${address}`);
  }
  const wholeColumn = parseWholeColumnRange(parts, address);
  if (wholeColumn) {
    return { ...wholeColumn, sheetName: qualified.sheetName };
  }
  const wholeRow = parseWholeRowRange(parts, address);
  if (wholeRow) {
    return { ...wholeRow, sheetName: qualified.sheetName };
  }
  const start = parseCellAddress(startRaw);
  const end = endRaw ? parseCellAddress(endRaw) : start;
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startColumn = Math.min(start.column, end.column);
  const endColumn = Math.max(start.column, end.column);
  const first = formatCellAddress(startRow, startColumn);
  const last = formatCellAddress(endRow, endColumn);
  return {
    address: first === last ? first : `${first}:${last}`,
    sheetName: qualified.sheetName,
    startRow,
    startColumn,
    rowCount: endRow - startRow + 1,
    columnCount: endColumn - startColumn + 1,
  };
}

function parseWholeColumnRange(
  parts: string[],
  originalAddress: string,
): Omit<RangeAddress, 'sheetName'> | null {
  if (parts.length !== 2) return null;
  const [startRaw, endRaw] = parts;
  const startColumn = parseColumnReference(startRaw);
  const endColumn = endRaw ? parseColumnReference(endRaw) : startColumn;
  if (startColumn === null && endColumn === null) return null;
  if (startColumn === null || endColumn === null) {
    throw new Error(`ExcelWorkbook: invalid range address ${originalAddress}`);
  }
  const firstColumn = Math.min(startColumn, endColumn);
  const lastColumn = Math.max(startColumn, endColumn);
  return {
    address: formatRangeAddress(
      1,
      firstColumn,
      EXCEL_MAX_ROW,
      lastColumn - firstColumn + 1,
    ),
    startRow: 1,
    startColumn: firstColumn,
    rowCount: EXCEL_MAX_ROW,
    columnCount: lastColumn - firstColumn + 1,
  };
}

function parseWholeRowRange(
  parts: string[],
  originalAddress: string,
): Omit<RangeAddress, 'sheetName'> | null {
  if (parts.length !== 2) return null;
  const [startRaw, endRaw] = parts;
  const startRow = parseRowReference(startRaw);
  const endRow = endRaw ? parseRowReference(endRaw) : startRow;
  if (startRow === null && endRow === null) return null;
  if (startRow === null || endRow === null) {
    throw new Error(`ExcelWorkbook: invalid range address ${originalAddress}`);
  }
  const firstRow = Math.min(startRow, endRow);
  const lastRow = Math.max(startRow, endRow);
  return {
    address: formatRangeAddress(
      firstRow,
      1,
      lastRow - firstRow + 1,
      EXCEL_MAX_COLUMN,
    ),
    startRow: firstRow,
    startColumn: 1,
    rowCount: lastRow - firstRow + 1,
    columnCount: EXCEL_MAX_COLUMN,
  };
}

function parseColumnReference(value: string | undefined): number | null {
  const normalized = value?.replace(/\$/g, '').trim();
  if (!normalized || !/^[A-Za-z]+$/.test(normalized)) return null;
  return columnNameToIndex(normalized);
}

function parseRowReference(value: string | undefined): number | null {
  const normalized = value?.replace(/\$/g, '').trim();
  if (!normalized || !/^\d+$/.test(normalized)) return null;
  const row = Number(normalized);
  if (!Number.isInteger(row) || row < 1) {
    throw new Error(`ExcelWorkbook: invalid row reference ${value}`);
  }
  return row;
}

export function splitSheetQualifiedAddress(input: string): SheetQualifiedAddress {
  const trimmed = input.trim();
  if (!trimmed) {
    throw new Error(`ExcelWorkbook: invalid address ${input}`);
  }

  let inQuote = false;
  for (let i = 0; i < trimmed.length; i++) {
    const ch = trimmed[i];
    if (ch === "'") {
      if (inQuote && trimmed[i + 1] === "'") {
        i += 1;
      } else {
        inQuote = !inQuote;
      }
      continue;
    }
    if (ch === '!' && !inQuote) {
      const rawSheet = trimmed.slice(0, i).trim();
      const address = trimmed.slice(i + 1).trim();
      if (!rawSheet || !address) {
        throw new Error(`ExcelWorkbook: invalid sheet-qualified address ${input}`);
      }
      return { sheetName: unquoteSheetName(rawSheet), address };
    }
  }

  return { address: trimmed };
}

export function formatRangeAddress(
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number,
): string {
  if (!Number.isInteger(rowCount) || rowCount < 1) {
    throw new Error(`ExcelWorkbook: invalid row count ${rowCount}`);
  }
  if (!Number.isInteger(columnCount) || columnCount < 1) {
    throw new Error(`ExcelWorkbook: invalid column count ${columnCount}`);
  }
  const first = formatCellAddress(startRow, startColumn);
  const last = formatCellAddress(startRow + rowCount - 1, startColumn + columnCount - 1);
  return first === last ? first : `${first}:${last}`;
}

function unquoteSheetName(raw: string): string {
  if (!raw.startsWith("'")) return raw;
  if (!raw.endsWith("'") || raw.length < 2) {
    throw new Error(`ExcelWorkbook: invalid quoted sheet name ${raw}`);
  }
  return raw.slice(1, -1).replace(/''/g, "'");
}
