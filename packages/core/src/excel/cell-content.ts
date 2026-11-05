import type { ExcelCellModel, ExcelSheetModel } from './model.js';

export function normalizeExcelCellFormula(formula: string | null): string | null {
  if (formula === null || formula === '') return null;
  const expression = formula.startsWith('=') ? formula.slice(1) : formula;
  if (expression.trim() === '') throw new Error('Excel cell formula: empty expression');
  return expression;
}

/** Clear content and its OOXML caches without discarding explicit formatting. */
export function clearExcelCellContent(
  sheet: ExcelSheetModel,
  cellsByAddress: Map<string, ExcelCellModel>,
  address: string,
): void {
  const cell = cellsByAddress.get(address);
  if (!cell) return;
  if (cell.styleId === undefined) {
    cellsByAddress.delete(address);
    const index = sheet.cells.indexOf(cell);
    if (index !== -1) sheet.cells.splice(index, 1);
    return;
  }
  cell.kind = 'blank';
  cell.value = null;
  delete cell.formula;
  delete cell.rawValue;
  delete cell.dataType;
  delete cell.sharedStringIndex;
}
