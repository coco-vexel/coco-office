import type { SpreadsheetRange, SpreadsheetSheetRenderModel } from './types.js';
import { getRenderCell } from './render-model.js';

export type SpreadsheetSortDirection = 'asc' | 'desc';

export interface SpreadsheetColumnFilter {
  column: number;
  values: readonly string[];
}

export function spreadsheetFilterValues(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
  column: number,
): string[] {
  const values = new Set<string>();
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    values.add(getRenderCell(model, row, column)?.displayText ?? '');
  }
  return [...values].sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

export function spreadsheetVisibleFilterValues(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
  column: number,
): string[] {
  const visibleValues = new Set<string>();
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    if (model.rowMetaByIndex.get(row)?.hidden) continue;
    visibleValues.add(getRenderCell(model, row, column)?.displayText ?? '');
  }
  return [...visibleValues].sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

export function spreadsheetRowsForFilter(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
  column: number,
  value: string | readonly string[],
): number[] {
  const values = new Set(typeof value === 'string' ? [value] : value);
  const hiddenRows: number[] = [];
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    if (!values.has(getRenderCell(model, row, column)?.displayText ?? '')) hiddenRows.push(row);
  }
  return hiddenRows;
}

export function spreadsheetRowsForColumnFilters(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
  filters: readonly SpreadsheetColumnFilter[],
): number[] {
  const allowedByColumn = filters.map((filter) => [filter.column, new Set(filter.values)] as const);
  const hiddenRows: number[] = [];
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const matchesAll = allowedByColumn.every(([column, values]) =>
      values.has(getRenderCell(model, row, column)?.displayText ?? ''),
    );
    if (!matchesAll) hiddenRows.push(row);
  }
  return hiddenRows;
}

export function spreadsheetSortRows(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
  column: number,
  direction: SpreadsheetSortDirection,
): number[] {
  const rows = Array.from({ length: range.endRow - range.startRow }, (_, index) => range.startRow + index + 1);
  rows.sort((leftRow, rightRow) => {
    const left = getRenderCell(model, leftRow, column)?.displayText ?? '';
    const right = getRenderCell(model, rightRow, column)?.displayText ?? '';
    if (left === '' && right !== '') return 1;
    if (left !== '' && right === '') return -1;
    const comparison = left.localeCompare(right, undefined, { numeric: true, sensitivity: 'base' });
    return (direction === 'desc' ? -comparison : comparison) || leftRow - rightRow;
  });
  return rows;
}
