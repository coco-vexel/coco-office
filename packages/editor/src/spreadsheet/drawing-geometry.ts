import type { ExcelDrawingAnchorPosition, ExcelDrawingMarkerModel } from '@coco-vexel/coco-office/excel';
import type { SpreadsheetLayout } from './layout.js';
import type { SpreadsheetPoint, SpreadsheetRect } from './types.js';

export const DRAWING_RESIZE_HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;
export type DrawingResizeHandle = typeof DRAWING_RESIZE_HANDLES[number];

export function drawingHandleRect(rect: SpreadsheetRect, handle: DrawingResizeHandle): SpreadsheetRect {
  const size = 8;
  const x = rect.x + (handle.includes('w') ? 0 : handle.includes('e') ? rect.width : rect.width / 2);
  const y = rect.y + (handle.includes('n') ? 0 : handle.includes('s') ? rect.height : rect.height / 2);
  return { x: x - size / 2, y: y - size / 2, width: size, height: size };
}

export function dragDrawingRect(
  rect: SpreadsheetRect,
  delta: SpreadsheetPoint,
  bounds: { width: number; height: number },
  handle?: DrawingResizeHandle,
): SpreadsheetRect {
  if (!handle) {
    return {
      ...rect,
      x: clamp(rect.x + delta.x, 0, bounds.width - rect.width),
      y: clamp(rect.y + delta.y, 0, bounds.height - rect.height),
    };
  }
  const minWidth = Math.min(8, rect.width);
  const minHeight = Math.min(8, rect.height);
  const right = rect.x + rect.width;
  const bottom = rect.y + rect.height;
  const left = handle.includes('w') ? clamp(rect.x + delta.x, 0, right - minWidth) : rect.x;
  const top = handle.includes('n') ? clamp(rect.y + delta.y, 0, bottom - minHeight) : rect.y;
  return {
    x: left,
    y: top,
    width: (handle.includes('e') ? clamp(right + delta.x, left + minWidth, bounds.width) : right) - left,
    height: (handle.includes('s') ? clamp(bottom + delta.y, top + minHeight, bounds.height) : bottom) - top,
  };
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

export function drawingPositionForRect(
  layout: SpreadsheetLayout,
  rect: SpreadsheetRect,
  startRect: SpreadsheetRect,
  previous: ExcelDrawingAnchorPosition,
): ExcelDrawingAnchorPosition {
  const x = layout.cssPixelsToEmu(rect.x);
  const y = layout.cssPixelsToEmu(rect.y);
  const width = layout.cssPixelsToEmu(rect.width);
  const height = layout.cssPixelsToEmu(rect.height);
  const offset = (value: number | undefined, delta: number, minimum = -Infinity) =>
    value === undefined ? undefined : Math.max(minimum, value + layout.cssPixelsToEmu(delta));
  const transform = previous.transform && {
    x: offset(previous.transform.x, rect.x - startRect.x),
    y: offset(previous.transform.y, rect.y - startRect.y),
    width: offset(previous.transform.width, rect.width - startRect.width, 1),
    height: offset(previous.transform.height, rect.height - startRect.height, 1),
  };
  if (previous.kind === 'absolute') return { kind: 'absolute', x, y, width, height, transform };
  const from = drawingMarkerAt(layout, rect.x, rect.y);
  if (previous.kind === 'oneCell') return { kind: 'oneCell', from, width, height, transform };
  const to = drawingMarkerAt(layout, rect.x + rect.width, rect.y + rect.height);
  return { kind: 'twoCell', from, to, transform };
}

export function drawingMarkerAt(layout: SpreadsheetLayout, x: number, y: number): Required<ExcelDrawingMarkerModel> {
  const column = layout.columnAt(Math.max(0, x));
  const row = layout.rowAt(Math.max(0, y));
  return {
    column: column - 1,
    row: row - 1,
    columnOffset: layout.cssPixelsToEmu(x - layout.getColumnX(column)),
    rowOffset: layout.cssPixelsToEmu(y - layout.getRowY(row)),
  };
}
