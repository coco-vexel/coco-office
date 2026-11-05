import { formatCellAddress } from '@coco-vexel/coco-office/excel';
import { createSelection, expandSelectionForMergedRanges, extendSelection, normalizeSelectionRange, type SpreadsheetSelectionState } from './selection.js';
import type { SpreadsheetSheetRenderModel } from './types.js';

export function navigateSpreadsheetSelection(
  selection: SpreadsheetSelectionState,
  key: string,
  options: { shift?: boolean; command?: boolean; pageRows?: number },
  model: SpreadsheetSheetRenderModel,
  bounds: { rowCount: number; columnCount: number },
): SpreadsheetSelectionState | null {
  const range = normalizeSelectionRange(selection.range);
  const active = selection.active;
  const focus = selection.focus ?? {
    row: range.endRow === active.row ? range.startRow : range.endRow,
    column: range.endColumn === active.column ? range.startColumn : range.endColumn,
  };
  const point = options.shift && key !== 'Tab' && key !== 'Enter' ? focus : active;
  const deltas: Record<string, [number, number]> = {
    ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1],
    PageUp: [-(options.pageRows ?? 20), 0], PageDown: [options.pageRows ?? 20, 0],
    Enter: [options.shift ? -1 : 1, 0], Tab: [0, options.shift ? -1 : 1],
  };
  const delta = deltas[key];
  if (!delta && key !== 'Home' && !(key === 'End' && options.command)) return null;
  let row = point.row;
  let column = point.column;
  if (key === 'Home') {
    column = 1;
    if (options.command) row = 1;
  } else if (key === 'End') {
    row = 1; column = 1;
    for (const cell of model.source.cells) { row = Math.max(row, cell.row); column = Math.max(column, cell.column); }
  } else if (options.command && key.startsWith('Arrow')) {
    const step = delta!;
    const limit = step[0] ? bounds.rowCount : bounds.columnCount;
    const position = step[0] ? row : column;
    const direction = step[0] || step[1];
    const populated = (index: number) => {
      const cell = model.cellsByAddress.get(formatCellAddress(step[0] ? index : row, step[1] ? index : column));
      return !!cell?.formula || (cell?.value !== undefined && cell.value !== null && cell.value !== '');
    };
    let next = Math.max(1, Math.min(limit, position + direction));
    if (populated(position) && populated(next)) {
      while (next + direction >= 1 && next + direction <= limit && populated(next + direction)) next += direction;
    } else {
      while (next > 1 && next < limit && !populated(next)) next += direction;
    }
    if (step[0]) row = next; else column = next;
  } else {
    const merge = [...model.mergeMasterByAddress.values()].find((item) =>
      row >= item.range.startRow && row <= item.range.endRow && column >= item.range.startColumn && column <= item.range.endColumn);
    if (!options.shift && merge && Math.abs(delta![0]) <= 1 && Math.abs(delta![1]) <= 1) {
      row = delta![0] > 0 ? merge.range.endRow + 1 : delta![0] < 0 ? merge.range.startRow - 1 : row;
      column = delta![1] > 0 ? merge.range.endColumn + 1 : delta![1] < 0 ? merge.range.startColumn - 1 : column;
    } else { row += delta![0]; column += delta![1]; }
    if ((key === 'Tab' || key === 'Enter') && (range.startRow !== range.endRow || range.startColumn !== range.endColumn) && !merge) {
      if (key === 'Tab') {
        if (column > range.endColumn) { column = range.startColumn; row++; }
        if (column < range.startColumn) { column = range.endColumn; row--; }
      } else {
        if (row > range.endRow) { row = range.startRow; column++; }
        if (row < range.startRow) { row = range.endRow; column--; }
      }
      row = row > range.endRow ? range.startRow : row < range.startRow ? range.endRow : row;
      column = column > range.endColumn ? range.startColumn : column < range.startColumn ? range.endColumn : column;
      return { kind: selection.kind, active: { row, column }, range };
    }
  }
  row = Math.max(1, Math.min(bounds.rowCount, row));
  column = Math.max(1, Math.min(bounds.columnCount, column));
  if (options.shift && key !== 'Tab' && key !== 'Enter') {
    return { ...expandSelectionForMergedRanges(extendSelection(selection, row, column, bounds), model), focus: { row, column } };
  }
  return expandSelectionForMergedRanges(createSelection(row, column), model);
}

export function selectSpreadsheetRegion(selection: SpreadsheetSelectionState, model: SpreadsheetSheetRenderModel, bounds: { rowCount: number; columnCount: number }): SpreadsheetSelectionState {
  const { row, column } = selection.active;
  const cells = [...model.cellsByAddress.values()].filter((cell) => cell.formula || (cell.value !== null && cell.value !== ''));
  const region = { startRow: row, endRow: row, startColumn: column, endColumn: column };
  let changed = true;
  while (changed) {
    changed = false;
    for (const cell of cells) {
      if (cell.row < region.startRow - 1 || cell.row > region.endRow + 1 || cell.column < region.startColumn - 1 || cell.column > region.endColumn + 1) continue;
      if (cell.row >= region.startRow && cell.row <= region.endRow && cell.column >= region.startColumn && cell.column <= region.endColumn) continue;
      region.startRow = Math.min(region.startRow, cell.row); region.endRow = Math.max(region.endRow, cell.row);
      region.startColumn = Math.min(region.startColumn, cell.column); region.endColumn = Math.max(region.endColumn, cell.column);
      changed = true;
    }
  }
  const empty = !cells.some((cell) => cell.row === row && cell.column === column);
  const current = normalizeSelectionRange(selection.range);
  const alreadySelected = Object.keys(region).every((key) => region[key as keyof typeof region] === current[key as keyof typeof current]);
  return expandSelectionForMergedRanges({ kind: 'cell', active: selection.active, range: empty || alreadySelected
    ? { startRow: 1, startColumn: 1, endRow: bounds.rowCount, endColumn: bounds.columnCount } : region }, model);
}
