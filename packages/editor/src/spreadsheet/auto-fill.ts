import { shiftFormulaReferences, type ExcelCellValue } from '@coco-vexel/coco-office/excel';
import { getRenderCell } from './render-model.js';
import type { SpreadsheetRange, SpreadsheetSheetRenderModel } from './types.js';

export type SpreadsheetAutoFillDirection = 'down' | 'right';

export interface SpreadsheetAutoFillCell {
  row: number;
  column: number;
  value: ExcelCellValue;
  formula?: string;
  styleId?: number;
}

export interface SpreadsheetAutoFillPlan {
  direction: SpreadsheetAutoFillDirection;
  source: SpreadsheetRange;
  target: SpreadsheetRange;
  cells: SpreadsheetAutoFillCell[];
}

/**
 * Produces the cells outside the original selection for a down/right fill.
 * Merged ranges are deliberately excluded for now: their copy semantics need a
 * separate structural operation rather than cell-by-cell expansion.
 */
export function buildSpreadsheetAutoFillPlan(
  model: SpreadsheetSheetRenderModel,
  sourceRange: SpreadsheetRange,
  targetRange: SpreadsheetRange,
): SpreadsheetAutoFillPlan {
  const source = normalizeRange(sourceRange);
  const target = normalizeRange(targetRange);
  const direction = resolveAutoFillDirection(source, target);
  if (rangeTouchesMergedCells(source, model) || rangeTouchesMergedCells(target, model)) {
    throw new Error('Auto fill does not support merged cells');
  }

  const cells: SpreadsheetAutoFillCell[] = [];
  for (let row = target.startRow; row <= target.endRow; row++) {
    for (let column = target.startColumn; column <= target.endColumn; column++) {
      if (isInsideRange(row, column, source)) continue;
      const sourceRow = direction === 'down'
        ? source.startRow + modulo(row - source.startRow, heightOf(source))
        : row;
      const sourceColumn = direction === 'right'
        ? source.startColumn + modulo(column - source.startColumn, widthOf(source))
        : column;
      const sourceCell = getRenderCell(model, sourceRow, sourceColumn);
      const sequenceValue = autoFillSequenceValue(model, source, direction, row, column);
      const formula = sourceCell?.formula
        ? shiftFormulaReferences(sourceCell.formula, row - sourceRow, column - sourceColumn)
        : undefined;
      cells.push({
        row,
        column,
        value: formula ? null : sequenceValue ?? sourceCell?.source.value ?? null,
        formula,
        styleId: sourceCell?.styleId,
      });
    }
  }

  return { direction, source, target, cells };
}

export function supportsSpreadsheetAutoFill(
  model: SpreadsheetSheetRenderModel,
  sourceRange: SpreadsheetRange,
): boolean {
  return !rangeTouchesMergedCells(normalizeRange(sourceRange), model);
}

function resolveAutoFillDirection(
  source: SpreadsheetRange,
  target: SpreadsheetRange,
): SpreadsheetAutoFillDirection {
  const fillsDown =
    target.startRow === source.startRow &&
    target.startColumn === source.startColumn &&
    target.endColumn === source.endColumn &&
    target.endRow > source.endRow;
  if (fillsDown) return 'down';

  const fillsRight =
    target.startRow === source.startRow &&
    target.endRow === source.endRow &&
    target.startColumn === source.startColumn &&
    target.endColumn > source.endColumn;
  if (fillsRight) return 'right';
  throw new Error('Auto fill target must extend the source selection down or right');
}

function autoFillSequenceValue(
  model: SpreadsheetSheetRenderModel,
  source: SpreadsheetRange,
  direction: SpreadsheetAutoFillDirection,
  row: number,
  column: number,
): number | undefined {
  const length = direction === 'down' ? heightOf(source) : widthOf(source);
  if (length < 2) return undefined;
  const values = Array.from({ length }, (_, index) => {
    const sourceRow = direction === 'down' ? source.startRow + index : row;
    const sourceColumn = direction === 'right' ? source.startColumn + index : column;
    const cell = getRenderCell(model, sourceRow, sourceColumn);
    return cell?.formula ? undefined : cell?.source.value;
  });
  if (!values.every((value): value is number => typeof value === 'number' && Number.isFinite(value))) {
    return undefined;
  }
  const step = values[1]! - values[0]!;
  if (!values.every((value, index) => index === 0 || value === values[0]! + step * index)) return undefined;
  const position = direction === 'down' ? row - source.startRow : column - source.startColumn;
  return values[0]! + step * position;
}

function rangeTouchesMergedCells(range: SpreadsheetRange, model: SpreadsheetSheetRenderModel): boolean {
  for (const merge of model.mergeMasterByAddress.values()) {
    if (rangesIntersect(range, merge.range)) return true;
  }
  return false;
}

function normalizeRange(range: SpreadsheetRange): SpreadsheetRange {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endRow: Math.max(range.startRow, range.endRow),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

function heightOf(range: SpreadsheetRange): number {
  return range.endRow - range.startRow + 1;
}

function widthOf(range: SpreadsheetRange): number {
  return range.endColumn - range.startColumn + 1;
}

function isInsideRange(row: number, column: number, range: SpreadsheetRange): boolean {
  return row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn;
}

function rangesIntersect(a: SpreadsheetRange, b: SpreadsheetRange): boolean {
  return !(
    a.endRow < b.startRow ||
    a.startRow > b.endRow ||
    a.endColumn < b.startColumn ||
    a.startColumn > b.endColumn
  );
}

function modulo(value: number, divisor: number): number {
  return ((value % divisor) + divisor) % divisor;
}
