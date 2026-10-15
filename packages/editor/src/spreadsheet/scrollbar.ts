import type { SpreadsheetScrollState } from './types.js';
import type { SpreadsheetLayout } from './layout.js';
import type { SpreadsheetFreezeState, SpreadsheetSize } from './types.js';

export function applyWheelDelta(
  scroll: SpreadsheetScrollState,
  deltaX: number,
  deltaY: number,
): SpreadsheetScrollState {
  return {
    x: Math.max(0, scroll.x + deltaX),
    y: Math.max(0, scroll.y + deltaY),
  };
}

export interface SpreadsheetScrollbarGeometry {
  trackLength: number;
  thumbStart: number;
  thumbLength: number;
  maxScroll: number;
}

export interface SpreadsheetScrollbarState {
  horizontal: SpreadsheetScrollbarGeometry;
  vertical: SpreadsheetScrollbarGeometry;
}

export function computeScrollbarState(
  layout: SpreadsheetLayout,
  viewportSize: SpreadsheetSize,
  freeze: SpreadsheetFreezeState,
  scroll: SpreadsheetScrollState,
  scrollbarThickness = 12,
): SpreadsheetScrollbarState {
  const frozenWidth = layout.getColumnX(freeze.frozenColumns + 1);
  const frozenHeight = layout.getRowY(freeze.frozenRows + 1);
  const horizontalTrackLength = Math.max(
    0,
    viewportSize.width - layout.rowHeaderWidth - frozenWidth - scrollbarThickness,
  );
  const verticalTrackLength = Math.max(
    0,
    viewportSize.height - layout.columnHeaderHeight - frozenHeight - scrollbarThickness,
  );
  const horizontalViewport = horizontalTrackLength;
  const verticalViewport = verticalTrackLength;
  const horizontalContent = Math.max(0, layout.totalWidth() - frozenWidth);
  const verticalContent = Math.max(0, layout.totalHeight() - frozenHeight);

  return {
    horizontal: computeAxisScrollbar(horizontalContent, horizontalViewport, horizontalTrackLength, scroll.x),
    vertical: computeAxisScrollbar(verticalContent, verticalViewport, verticalTrackLength, scroll.y),
  };
}

export function scrollFromThumbPosition(
  geometry: SpreadsheetScrollbarGeometry,
  thumbStart: number,
): number {
  if (geometry.maxScroll <= 0) return 0;
  const maxThumbStart = Math.max(0, geometry.trackLength - geometry.thumbLength);
  if (maxThumbStart <= 0) return 0;
  const clamped = Math.max(0, Math.min(maxThumbStart, thumbStart));
  return (clamped / maxThumbStart) * geometry.maxScroll;
}

function computeAxisScrollbar(
  contentLength: number,
  viewportLength: number,
  trackLength: number,
  scrollOffset: number,
): SpreadsheetScrollbarGeometry {
  const maxScroll = Math.max(0, contentLength - viewportLength);
  if (trackLength <= 0 || contentLength <= 0 || maxScroll <= 0) {
    return { trackLength, thumbStart: 0, thumbLength: trackLength, maxScroll: 0 };
  }
  const thumbLength = Math.max(24, Math.min(trackLength, (viewportLength / contentLength) * trackLength));
  const maxThumbStart = Math.max(0, trackLength - thumbLength);
  const thumbStart = maxScroll <= 0 ? 0 : (Math.max(0, Math.min(maxScroll, scrollOffset)) / maxScroll) * maxThumbStart;
  return { trackLength, thumbStart, thumbLength, maxScroll };
}
