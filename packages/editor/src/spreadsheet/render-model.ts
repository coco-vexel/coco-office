import {
  formatCellAddress,
  parseCellAddress,
  parseRangeAddress,
  type ExcelCellModel,
  type ExcelColumnModel,
  type ExcelRowModel,
  type ExcelSheetModel,
  type ExcelStylesModel,
  type SpreadsheetFormulaResult,
} from '@coco-vexel/coco-office/excel';
import {
  rowColumnKey,
  type SpreadsheetMergeInfo,
  type SpreadsheetRange,
  type SpreadsheetRenderCell,
  type SpreadsheetSheetRenderModel,
} from './types.js';

export interface BuildSpreadsheetRenderModelOptions {
  minimumRows?: number;
  minimumColumns?: number;
  formulaResults?: Map<string, SpreadsheetFormulaResult>;
}

export function buildSpreadsheetSheetRenderModel(
  sheet: ExcelSheetModel,
  styles: ExcelStylesModel,
  options: BuildSpreadsheetRenderModelOptions = {},
): SpreadsheetSheetRenderModel {
  const cellsByAddress = new Map<string, SpreadsheetRenderCell>();
  const cellsByRowColumn = new Map<string, SpreadsheetRenderCell>();
  const rowMetaByIndex = new Map<number, ExcelRowModel>();
  const columnMetaByIndex = new Map<number, ExcelColumnModel>();
  const mergeMasterByAddress = new Map<string, SpreadsheetMergeInfo>();
  const mergeFollowerByAddress = new Map<string, string>();
  let maxRow = options.minimumRows ?? 100;
  let maxColumn = options.minimumColumns ?? 26;

  for (const row of sheet.rows) {
    rowMetaByIndex.set(row.index, row);
    maxRow = Math.max(maxRow, row.index);
  }

  for (const column of sheet.columns) {
    for (let index = column.min; index <= column.max; index++) {
      columnMetaByIndex.set(index, column);
    }
    maxColumn = Math.max(maxColumn, column.max);
  }

  for (const cell of sheet.cells) {
    const renderCell = toRenderCell(cell, styles, options.formulaResults);
    cellsByAddress.set(cell.address, renderCell);
    cellsByRowColumn.set(rowColumnKey(cell.row, cell.column), renderCell);
    maxRow = Math.max(maxRow, cell.row);
    maxColumn = Math.max(maxColumn, cell.column);
  }

  // Dynamic-array spill cells are calculated results without persisted source cells.
  for (const [address, result] of options.formulaResults ?? []) {
    if (cellsByAddress.has(address) || result.error) continue;
    const parsed = parseCellAddress(address);
    const virtualCell: ExcelCellModel = {
      address,
      row: parsed.row,
      column: parsed.column,
      kind: 'formula',
      value: null,
    };
    const renderCell = toRenderCell(virtualCell, styles, options.formulaResults);
    cellsByAddress.set(address, renderCell);
    cellsByRowColumn.set(rowColumnKey(parsed.row, parsed.column), renderCell);
    maxRow = Math.max(maxRow, parsed.row);
    maxColumn = Math.max(maxColumn, parsed.column);
  }

  for (const merge of sheet.mergedRanges) {
    const range = parseRangeAddress(merge.address);
    const endRow = range.startRow + range.rowCount - 1;
    const endColumn = range.startColumn + range.columnCount - 1;
    const renderRange: SpreadsheetRange = {
      startRow: range.startRow,
      startColumn: range.startColumn,
      endRow,
      endColumn,
    };
    const masterAddress = formatCellAddress(range.startRow, range.startColumn);
    const info: SpreadsheetMergeInfo = {
      address: merge.address,
      range: renderRange,
      masterAddress,
      source: merge,
    };
    mergeMasterByAddress.set(masterAddress, info);
    maxRow = Math.max(maxRow, endRow);
    maxColumn = Math.max(maxColumn, endColumn);

    for (let row = range.startRow; row <= endRow; row++) {
      for (let column = range.startColumn; column <= endColumn; column++) {
        const address = formatCellAddress(row, column);
        if (address !== masterAddress) mergeFollowerByAddress.set(address, masterAddress);
      }
    }
  }

  return {
    sheetId: sheet.id,
    sheetName: sheet.name,
    source: sheet,
    styles,
    cellsByAddress,
    cellsByRowColumn,
    rowMetaByIndex,
    columnMetaByIndex,
    mergeMasterByAddress,
    mergeFollowerByAddress,
    maxRow,
    maxColumn,
  };
}

export function getRenderCell(
  model: SpreadsheetSheetRenderModel,
  row: number,
  column: number,
): SpreadsheetRenderCell | undefined {
  return model.cellsByRowColumn.get(rowColumnKey(row, column));
}

function toRenderCell(
  cell: ExcelCellModel,
  styles: ExcelStylesModel,
  formulaResults?: Map<string, SpreadsheetFormulaResult>,
): SpreadsheetRenderCell {
  const formulaResult = formulaResults?.get(cell.address);
  const resultValue = formulaResult && !formulaResult.error ? formulaResult.value : undefined;
  const value = Array.isArray(resultValue) ? resultValue[0]?.[0] ?? null : resultValue ?? cell.value;
  return {
    row: cell.row,
    column: cell.column,
    address: cell.address,
    kind: cell.kind,
    value,
    displayText: cellToDisplayText(cell, styles, value, formulaResult),
    formula: cell.formula,
    styleId: cell.styleId,
    source: cell,
  };
}

function cellToDisplayText(
  cell: ExcelCellModel,
  styles: ExcelStylesModel,
  value: ExcelCellModel['value'],
  formulaResult?: SpreadsheetFormulaResult,
): string {
  if (formulaResult?.error) return formulaResult.displayText;
  if (value !== null) return formatSpreadsheetDisplayValue(value, numberFormatCodeForCell(cell, styles));
  return '';
}

function numberFormatCodeForCell(cell: ExcelCellModel, styles: ExcelStylesModel): string | undefined {
  return styles.cellFormats[cell.styleId ?? 0]?.numberFormatCode;
}

export function formatSpreadsheetDisplayValue(
  value: ExcelCellModel['value'],
  formatCode: string | undefined,
): string {
  if (value === null) return '';
  if (typeof value !== 'number' || !formatCode) return String(value);
  const normalized = formatCode.toLowerCase();
  if (normalized.includes('%')) return `${formatFixedNumber(value * 100, decimalPlaces(formatCode))}%`;
  if (normalized.includes('$')) return `$${formatFixedNumber(value, decimalPlaces(formatCode), true)}`;
  if (looksLikeDateFormat(normalized)) return formatExcelSerialDate(value);
  if (formatCode.includes(',')) return formatFixedNumber(value, decimalPlaces(formatCode), true);
  if (formatCode.includes('.')) return formatFixedNumber(value, decimalPlaces(formatCode));
  return String(value);
}

function decimalPlaces(formatCode: string): number {
  const match = /\.(0+)/.exec(formatCode);
  return match?.[1]?.length ?? 0;
}

function formatFixedNumber(value: number, digits: number, grouping = false): string {
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
    useGrouping: grouping,
  }).format(value);
}

function looksLikeDateFormat(formatCode: string): boolean {
  return /(^|[^a-z])y{2,4}([^a-z]|$)/.test(formatCode) && /(^|[^a-z])d{1,2}([^a-z]|$)/.test(formatCode);
}

function formatExcelSerialDate(serial: number): string {
  if (!Number.isFinite(serial)) return String(serial);
  const epoch = Date.UTC(1899, 11, 30);
  const date = new Date(epoch + Math.round(serial) * 86_400_000);
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}
