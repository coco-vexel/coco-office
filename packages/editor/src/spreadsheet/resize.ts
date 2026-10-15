import type { SpreadsheetLayout } from './layout.js';
import type { SpreadsheetViewport } from './types.js';

export const SPREADSHEET_COLUMN_RESIZE_MIN = 24;
export const SPREADSHEET_ROW_RESIZE_MIN = 12;
export const SPREADSHEET_RESIZE_HIT_TOLERANCE = 4;

export interface SpreadsheetResizeHandle {
  kind: 'column' | 'row';
  index: number;
  startSize: number;
  guidePosition: number;
  minSize: number;
}

export interface SpreadsheetResizeDragState extends SpreadsheetResizeHandle {
  startClientX: number;
  startClientY: number;
  currentSize: number;
  currentGuidePosition: number;
}

export function hitSpreadsheetResizeHandle(
  layout: SpreadsheetLayout,
  viewports: readonly SpreadsheetViewport[],
  point: { x: number; y: number },
): SpreadsheetResizeHandle | null {
  for (const viewport of viewports) {
    if (viewport.kind === 'columnHeaderFrozen' || viewport.kind === 'columnHeaderScrollable') {
      const handle = hitSpreadsheetColumnResizeHandle(layout, viewport, point);
      if (handle) return handle;
    }
    if (viewport.kind === 'rowHeaderFrozen' || viewport.kind === 'rowHeaderScrollable') {
      const handle = hitSpreadsheetRowResizeHandle(layout, viewport, point);
      if (handle) return handle;
    }
  }
  return null;
}

export function startSpreadsheetResizeDrag(
  handle: SpreadsheetResizeHandle,
  event: { clientX: number; clientY: number },
): SpreadsheetResizeDragState {
  return {
    ...handle,
    startClientX: event.clientX,
    startClientY: event.clientY,
    currentSize: handle.startSize,
    currentGuidePosition: handle.guidePosition,
  };
}

export function updateSpreadsheetResizeDrag(
  drag: SpreadsheetResizeDragState,
  event: { clientX: number; clientY: number },
): SpreadsheetResizeDragState {
  const delta = drag.kind === 'column' ? event.clientX - drag.startClientX : event.clientY - drag.startClientY;
  const currentSize = Math.max(drag.minSize, Math.round(drag.startSize + delta));
  return {
    ...drag,
    currentSize,
    currentGuidePosition: drag.guidePosition + currentSize - drag.startSize,
  };
}

function hitSpreadsheetColumnResizeHandle(
  layout: SpreadsheetLayout,
  viewport: SpreadsheetViewport,
  point: { x: number; y: number },
): SpreadsheetResizeHandle | null {
  if (!containsSpreadsheetPoint(viewport.screenRect, point, SPREADSHEET_RESIZE_HIT_TOLERANCE)) return null;
  for (let column = viewport.columns.start; column <= viewport.columns.end; column++) {
    const logical = layout.getCellLogicalRect(1, column);
    const left = viewport.screenRect.x + logical.x - viewport.originX;
    const right = left + logical.width;
    if (Math.abs(point.x - right) <= SPREADSHEET_RESIZE_HIT_TOLERANCE) {
      return {
        kind: 'column',
        index: column,
        startSize: layout.getColumnWidth(column),
        guidePosition: right,
        minSize: SPREADSHEET_COLUMN_RESIZE_MIN,
      };
    }
  }
  return null;
}

function hitSpreadsheetRowResizeHandle(
  layout: SpreadsheetLayout,
  viewport: SpreadsheetViewport,
  point: { x: number; y: number },
): SpreadsheetResizeHandle | null {
  if (!containsSpreadsheetPoint(viewport.screenRect, point, SPREADSHEET_RESIZE_HIT_TOLERANCE)) return null;
  for (let row = viewport.rows.start; row <= viewport.rows.end; row++) {
    const logical = layout.getCellLogicalRect(row, 1);
    const top = viewport.screenRect.y + logical.y - viewport.originY;
    const bottom = top + logical.height;
    if (Math.abs(point.y - bottom) <= SPREADSHEET_RESIZE_HIT_TOLERANCE) {
      return {
        kind: 'row',
        index: row,
        startSize: layout.getRowHeight(row),
        guidePosition: bottom,
        minSize: SPREADSHEET_ROW_RESIZE_MIN,
      };
    }
  }
  return null;
}

function containsSpreadsheetPoint(
  rect: { x: number; y: number; width: number; height: number },
  point: { x: number; y: number },
  tolerance = 0,
): boolean {
  return (
    point.x >= rect.x - tolerance &&
    point.y >= rect.y - tolerance &&
    point.x <= rect.x + rect.width + tolerance &&
    point.y <= rect.y + rect.height + tolerance
  );
}
