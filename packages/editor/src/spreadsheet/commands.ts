import type { ExcelCellValue, ExcelWorkbook } from '@coco-vexel/coco-office/excel';
import { formatCellAddress, formatRangeAddress } from '@coco-vexel/coco-office/excel';
import { cssPixelsToExcelColumnWidth } from './layout.js';

export async function setSpreadsheetCellValue(
  workbook: ExcelWorkbook,
  sheetName: string,
  address: string,
  value: ExcelCellValue,
): Promise<void> {
  await workbook.worksheets.getItem(sheetName).getCell(address).setValue(value);
}

export async function setSpreadsheetCellFormula(
  workbook: ExcelWorkbook,
  sheetName: string,
  address: string,
  formula: string | null,
): Promise<void> {
  await workbook.worksheets.getItem(sheetName).getCell(address).setFormula(formula);
}

export async function setSpreadsheetRangeValues(
  workbook: ExcelWorkbook,
  sheetName: string,
  startRow: number,
  startColumn: number,
  values: ExcelCellValue[][],
): Promise<void> {
  const rowCount = values.length;
  const columnCount = values.reduce((max, row) => Math.max(max, row.length), 0);
  const normalized = values.map((row) => {
    const next = row.slice();
    while (next.length < columnCount) next.push(null);
    return next;
  });
  await workbook.worksheets
    .getItem(sheetName)
    .getRange(formatRangeAddress(startRow, startColumn, rowCount, columnCount))
    .setValues(normalized);
}

export async function clearSpreadsheetRange(
  workbook: ExcelWorkbook,
  sheetName: string,
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number,
): Promise<void> {
  await workbook.worksheets
    .getItem(sheetName)
    .getRange(formatRangeAddress(startRow, startColumn, rowCount, columnCount))
    .clear();
}

export async function setSpreadsheetRowHeight(
  workbook: ExcelWorkbook,
  sheetName: string,
  row: number,
  height: number,
): Promise<void> {
  await workbook.worksheets
    .getItem(sheetName)
    .getRange(formatCellAddress(row, 1))
    .setRowHeight(height);
}

export async function setSpreadsheetColumnWidth(
  workbook: ExcelWorkbook,
  sheetName: string,
  column: number,
  width: number,
): Promise<void> {
  await workbook.worksheets
    .getItem(sheetName)
    .getRange(formatCellAddress(1, column))
    .setColumnWidth(cssPixelsToExcelColumnWidth(width));
}
