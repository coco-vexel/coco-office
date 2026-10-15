import { formatCellAddress } from '@coco-vexel/coco-office/excel';
import type { SpreadsheetRange, SpreadsheetSheetRenderModel } from './types.js';

export interface SpreadsheetSelectionState {
  kind: 'cell' | 'row' | 'column';
  active: { row: number; column: number };
  range: SpreadsheetRange;
}

export function createSelection(row: number, column: number): SpreadsheetSelectionState {
  return {
    kind: 'cell',
    active: { row, column },
    range: { startRow: row, startColumn: column, endRow: row, endColumn: column },
  };
}

export function createRowSelection(
  row: number,
  bounds: { columnCount: number },
): SpreadsheetSelectionState {
  return {
    kind: 'row',
    active: { row, column: 1 },
    range: { startRow: row, startColumn: 1, endRow: row, endColumn: bounds.columnCount },
  };
}

export function createRowRangeSelection(
  startRow: number,
  endRow: number,
  bounds: { columnCount: number },
): SpreadsheetSelectionState {
  const range = normalizeSelectionRange({
    startRow,
    startColumn: 1,
    endRow,
    endColumn: bounds.columnCount,
  });
  return {
    kind: 'row',
    active: { row: range.startRow, column: 1 },
    range,
  };
}

export function createColumnSelection(
  column: number,
  bounds: { rowCount: number },
): SpreadsheetSelectionState {
  return {
    kind: 'column',
    active: { row: 1, column },
    range: { startRow: 1, startColumn: column, endRow: bounds.rowCount, endColumn: column },
  };
}

export function createColumnRangeSelection(
  startColumn: number,
  endColumn: number,
  bounds: { rowCount: number },
): SpreadsheetSelectionState {
  const range = normalizeSelectionRange({
    startRow: 1,
    startColumn,
    endRow: bounds.rowCount,
    endColumn,
  });
  return {
    kind: 'column',
    active: { row: 1, column: range.startColumn },
    range,
  };
}

export function normalizeSelectionRange(range: SpreadsheetRange): SpreadsheetRange {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endRow: Math.max(range.startRow, range.endRow),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

export function extendSelection(
  selection: SpreadsheetSelectionState,
  row: number,
  column: number,
  bounds?: { rowCount: number; columnCount: number },
): SpreadsheetSelectionState {
  if (selection.kind === 'row') {
    return {
      kind: 'row',
      active: selection.active,
      range: normalizeSelectionRange({
        startRow: selection.active.row,
        startColumn: 1,
        endRow: row,
        endColumn: bounds?.columnCount ?? selection.range.endColumn,
      }),
    };
  }
  if (selection.kind === 'column') {
    return {
      kind: 'column',
      active: selection.active,
      range: normalizeSelectionRange({
        startRow: 1,
        startColumn: selection.active.column,
        endRow: bounds?.rowCount ?? selection.range.endRow,
        endColumn: column,
      }),
    };
  }
  return {
    kind: 'cell',
    active: selection.active,
    range: normalizeSelectionRange({
      startRow: selection.active.row,
      startColumn: selection.active.column,
      endRow: row,
      endColumn: column,
    }),
  };
}

export function moveSelection(
  selection: SpreadsheetSelectionState,
  deltaRow: number,
  deltaColumn: number,
  bounds: { rowCount: number; columnCount: number },
): SpreadsheetSelectionState {
  return createSelection(
    clamp(selection.active.row + deltaRow, 1, bounds.rowCount),
    clamp(selection.active.column + deltaColumn, 1, bounds.columnCount),
  );
}

/**
 * Spreadsheet selections must include a full merged block once any cell in that block is touched.
 * Keep this shared because drag selection, paste feedback, and host integrations need the same rule.
 */
export function expandSelectionForMergedRanges(
  selection: SpreadsheetSelectionState,
  model: SpreadsheetSheetRenderModel,
): SpreadsheetSelectionState {
  if (selection.kind !== 'cell') return selection;
  const active = mergeMasterPointForCell(selection.active.row, selection.active.column, model) ?? selection.active;
  return { ...selection, active, range: expandRangeForMergedRanges(selection.range, model) };
}

export function expandRangeForMergedRanges(
  range: SpreadsheetRange,
  model: SpreadsheetSheetRenderModel,
): SpreadsheetRange {
  let expanded = normalizeSelectionRange(range);
  let changed = true;
  while (changed) {
    changed = false;
    for (const merge of model.mergeMasterByAddress.values()) {
      if (!rangesIntersect(expanded, merge.range)) continue;
      const next = {
        startRow: Math.min(expanded.startRow, merge.range.startRow),
        startColumn: Math.min(expanded.startColumn, merge.range.startColumn),
        endRow: Math.max(expanded.endRow, merge.range.endRow),
        endColumn: Math.max(expanded.endColumn, merge.range.endColumn),
      };
      changed =
        next.startRow !== expanded.startRow ||
        next.startColumn !== expanded.startColumn ||
        next.endRow !== expanded.endRow ||
        next.endColumn !== expanded.endColumn;
      expanded = next;
    }
  }
  return expanded;
}

export function selectionAfterRowDeletion(
  selection: SpreadsheetSelectionState,
  deletedStartRow: number,
  deletedCount: number,
  bounds: { rowCount: number; columnCount: number },
  model: SpreadsheetSheetRenderModel,
): SpreadsheetSelectionState {
  const nextRow = clamp(positionAfterDeletion(selection.active.row, deletedStartRow, deletedCount), 1, bounds.rowCount);
  const nextColumn = clamp(selection.active.column, 1, bounds.columnCount);
  if (selection.kind === 'row') {
    return createRowSelection(nextRow, { columnCount: bounds.columnCount });
  }
  return expandSelectionForMergedRanges(createSelection(nextRow, nextColumn), model);
}

export function selectionAfterColumnDeletion(
  selection: SpreadsheetSelectionState,
  deletedStartColumn: number,
  deletedCount: number,
  bounds: { rowCount: number; columnCount: number },
  model: SpreadsheetSheetRenderModel,
): SpreadsheetSelectionState {
  const nextRow = clamp(selection.active.row, 1, bounds.rowCount);
  const nextColumn = clamp(positionAfterDeletion(selection.active.column, deletedStartColumn, deletedCount), 1, bounds.columnCount);
  if (selection.kind === 'column') {
    return createColumnSelection(nextColumn, { rowCount: bounds.rowCount });
  }
  return expandSelectionForMergedRanges(createSelection(nextRow, nextColumn), model);
}

function positionAfterDeletion(position: number, deletedStart: number, deletedCount: number): number {
  const deletedEnd = deletedStart + deletedCount - 1;
  if (position < deletedStart) return position;
  if (position > deletedEnd) return position - deletedCount;
  return deletedStart;
}

function rangesIntersect(a: SpreadsheetRange, b: SpreadsheetRange): boolean {
  return !(
    a.endRow < b.startRow ||
    a.startRow > b.endRow ||
    a.endColumn < b.startColumn ||
    a.startColumn > b.endColumn
  );
}

function mergeMasterPointForCell(
  row: number,
  column: number,
  model: SpreadsheetSheetRenderModel,
): { row: number; column: number } | null {
  const masterAddress = model.mergeFollowerByAddress.get(formatCellAddress(row, column));
  if (!masterAddress) return null;
  const merge = model.mergeMasterByAddress.get(masterAddress);
  if (!merge) return null;
  return { row: merge.range.startRow, column: merge.range.startColumn };
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}
