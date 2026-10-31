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
  date1904?: boolean;
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
    const renderCell = toRenderCell(cell, styles, options.formulaResults, options.date1904);
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
    const renderCell = toRenderCell(virtualCell, styles, options.formulaResults, options.date1904);
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
  date1904 = false,
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
    displayText: cellToDisplayText(cell, styles, value, formulaResult, date1904),
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
  date1904 = false,
): string {
  if (formulaResult?.error) return formulaResult.displayText;
  if (value !== null) return formatSpreadsheetDisplayValue(value, numberFormatCodeForCell(cell, styles), date1904);
  return '';
}

function numberFormatCodeForCell(cell: ExcelCellModel, styles: ExcelStylesModel): string | undefined {
  return styles.cellFormats[cell.styleId ?? 0]?.numberFormatCode;
}

export function formatSpreadsheetDisplayValue(
  value: ExcelCellModel['value'],
  formatCode: string | undefined,
  date1904 = false,
): string {
  if (value === null) return '';
  if (typeof value !== 'number' || !formatCode) return String(value);
  const normalized = formatCode.toLowerCase();
  if (normalized.includes('%')) return `${formatFixedNumber(value * 100, decimalPlaces(formatCode))}%`;
  if (normalized.includes('$')) return `$${formatFixedNumber(value, decimalPlaces(formatCode), true)}`;
  if (looksLikeDateTimeFormat(normalized)) return formatExcelSerialDateTime(value, normalized, date1904);
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

function looksLikeDateTimeFormat(formatCode: string): boolean {
  const knownFormats = [
    'd-mmm',
    'mmm-yy',
    'h:mm am/pm',
    'h:mm:ss am/pm',
    'h:mm',
    'h:mm:ss',
    'mm:ss',
    '[h]:mm:ss',
    'mmss.0',
  ];
  if (knownFormats.includes(formatCode)) return true;
  return (
    /(^|[^a-z])y{2,4}([^a-z]|$)/.test(formatCode) &&
    /(^|[^a-z])d{1,2}([^a-z]|$)/.test(formatCode)
  );
}

function formatExcelSerialDateTime(
  serial: number,
  formatCode: string,
  date1904: boolean,
): string {
  if (!Number.isFinite(serial)) return String(serial);
  const wholeDays = Math.floor(serial);
  const dayFraction = serial - wholeDays;
  const totalSeconds = Math.round(dayFraction * 86_400) % 86_400;
  const hour = Math.floor(totalSeconds / 3_600);
  const minute = Math.floor((totalSeconds % 3_600) / 60);
  const second = totalSeconds % 60;
  const hour12 = hour % 12 || 12;
  const amPm = hour < 12 ? 'AM' : 'PM';
  const pad = (value: number) => String(value).padStart(2, '0');

  if (formatCode === 'h:mm am/pm') return `${hour12}:${pad(minute)} ${amPm}`;
  if (formatCode === 'h:mm:ss am/pm') return `${hour12}:${pad(minute)}:${pad(second)} ${amPm}`;
  if (formatCode === 'h:mm') return `${hour}:${pad(minute)}`;
  if (formatCode === 'h:mm:ss') return `${hour}:${pad(minute)}:${pad(second)}`;
  if (formatCode === 'mm:ss') return `${pad(minute)}:${pad(second)}`;
  if (formatCode === '[h]:mm:ss') return `${Math.floor(serial * 24)}:${pad(minute)}:${pad(second)}`;
  if (formatCode === 'mmss.0') return `${pad(minute)}${pad(second)}.0`;

  const phantomLeapDay = !date1904 && wholeDays === 60;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const date = new Date(epoch + wholeDays * 86_400_000);
  const year = phantomLeapDay ? 1900 : date.getUTCFullYear();
  const monthNumber = phantomLeapDay ? 2 : date.getUTCMonth() + 1;
  const dayNumber = phantomLeapDay ? 29 : date.getUTCDate();
  const month = pad(monthNumber);
  const day = pad(dayNumber);
  const shortYear = pad(year % 100);
  const monthName = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ][monthNumber - 1]!;

  if (formatCode === 'mm-dd-yy') return `${month}-${day}-${shortYear}`;
  if (formatCode === 'd-mmm-yy') return `${dayNumber}-${monthName}-${shortYear}`;
  if (formatCode === 'd-mmm') return `${dayNumber}-${monthName}`;
  if (formatCode === 'mmm-yy') return `${monthName}-${shortYear}`;
  if (formatCode === 'm/d/yy h:mm') return `${monthNumber}/${dayNumber}/${shortYear} ${hour}:${pad(minute)}`;
  return `${year}-${month}-${day}`;
}
