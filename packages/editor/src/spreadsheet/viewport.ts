import type { ExcelPaneModel, ExcelSheetModel } from '@coco-vexel/coco-office/excel';
import type {
  SpreadsheetFreezeState,
  SpreadsheetHitResult,
  SpreadsheetPoint,
  SpreadsheetRect,
  SpreadsheetScrollState,
  SpreadsheetSize,
  SpreadsheetViewport,
} from './types.js';
import type { SpreadsheetLayout } from './layout.js';

export interface SpreadsheetViewportManagerOptions {
  viewportSize: SpreadsheetSize;
  scroll?: Partial<SpreadsheetScrollState>;
  freeze?: Partial<SpreadsheetFreezeState>;
}

export class SpreadsheetViewportManager {
  readonly scroll: SpreadsheetScrollState;
  readonly freeze: SpreadsheetFreezeState;
  readonly viewportSize: SpreadsheetSize;
  readonly viewports: SpreadsheetViewport[];

  constructor(
    private readonly layout: SpreadsheetLayout,
    options: SpreadsheetViewportManagerOptions,
  ) {
    this.viewportSize = options.viewportSize;
    this.freeze = {
      frozenRows: Math.max(0, options.freeze?.frozenRows ?? 0),
      frozenColumns: Math.max(0, options.freeze?.frozenColumns ?? 0),
    };
    this.scroll = clampScroll(layout, this.freeze, options.viewportSize, {
      x: options.scroll?.x ?? 0,
      y: options.scroll?.y ?? 0,
    });
    this.viewports = computeViewports(layout, this.viewportSize, this.scroll, this.freeze);
  }

  hitTest(point: SpreadsheetPoint): SpreadsheetHitResult {
    const viewport = this.findViewport(point);
    if (!viewport) return { kind: 'none' };
    if (viewport.kind === 'corner') return { kind: 'corner' };
    const logicalX = point.x - viewport.screenRect.x + viewport.originX;
    const logicalY = point.y - viewport.screenRect.y + viewport.originY;
    if (viewport.kind === 'columnHeaderScrollable' || viewport.kind === 'columnHeaderFrozen') {
      return { kind: 'columnHeader', column: this.layout.columnAt(logicalX) };
    }
    if (viewport.kind === 'rowHeaderScrollable' || viewport.kind === 'rowHeaderFrozen') {
      return { kind: 'rowHeader', row: this.layout.rowAt(logicalY) };
    }
    return this.layout.hitTestLogical(logicalX, logicalY);
  }

  findViewport(point: SpreadsheetPoint): SpreadsheetViewport | undefined {
    for (let index = this.viewports.length - 1; index >= 0; index--) {
      const viewport = this.viewports[index];
      if (viewport && containsPoint(viewport.screenRect, point)) return viewport;
    }
    return undefined;
  }

  cellScreenRect(row: number, column: number, viewport: SpreadsheetViewport): SpreadsheetRect {
    const logicalRect = this.layout.getCellLogicalRect(row, column);
    return {
      x: viewport.screenRect.x + logicalRect.x - viewport.originX,
      y: viewport.screenRect.y + logicalRect.y - viewport.originY,
      width: logicalRect.width,
      height: logicalRect.height,
    };
  }
}

export function freezeStateFromSheet(sheet: ExcelSheetModel): SpreadsheetFreezeState {
  const pane = sheet.sheetViews[0]?.pane;
  return freezeStateFromPane(pane);
}

export function freezeStateFromPane(pane: ExcelPaneModel | undefined): SpreadsheetFreezeState {
  if (!pane || pane.state !== 'frozen') return { frozenRows: 0, frozenColumns: 0 };
  return {
    frozenRows: Math.max(0, Math.floor(pane.ySplit ?? 0)),
    frozenColumns: Math.max(0, Math.floor(pane.xSplit ?? 0)),
  };
}

export function clampScroll(
  layout: SpreadsheetLayout,
  freeze: SpreadsheetFreezeState,
  viewportSize: SpreadsheetSize,
  scroll: SpreadsheetScrollState,
): SpreadsheetScrollState {
  const frozenWidth = layout.getColumnX(freeze.frozenColumns + 1);
  const frozenHeight = layout.getRowY(freeze.frozenRows + 1);
  const mainWidth = Math.max(0, viewportSize.width - layout.rowHeaderWidth - frozenWidth);
  const mainHeight = Math.max(0, viewportSize.height - layout.columnHeaderHeight - frozenHeight);
  const maxX = Math.max(0, layout.totalWidth() - frozenWidth - mainWidth);
  const maxY = Math.max(0, layout.totalHeight() - frozenHeight - mainHeight);
  return {
    x: Math.max(0, Math.min(maxX, scroll.x)),
    y: Math.max(0, Math.min(maxY, scroll.y)),
  };
}

function computeViewports(
  layout: SpreadsheetLayout,
  viewportSize: SpreadsheetSize,
  scroll: SpreadsheetScrollState,
  freeze: SpreadsheetFreezeState,
): SpreadsheetViewport[] {
  const headerW = layout.rowHeaderWidth;
  const headerH = layout.columnHeaderHeight;
  const frozenW = layout.getColumnX(freeze.frozenColumns + 1);
  const frozenH = layout.getRowY(freeze.frozenRows + 1);
  const scrollableX = frozenW + scroll.x;
  const scrollableY = frozenH + scroll.y;
  const bodyW = Math.max(0, viewportSize.width - headerW);
  const bodyH = Math.max(0, viewportSize.height - headerH);
  const mainX = headerW + frozenW;
  const mainY = headerH + frozenH;
  const mainW = Math.max(0, viewportSize.width - mainX);
  const mainH = Math.max(0, viewportSize.height - mainY);
  const viewports: SpreadsheetViewport[] = [];

  viewports.push({
    kind: 'corner',
    screenRect: { x: 0, y: 0, width: headerW, height: headerH },
    originX: 0,
    originY: 0,
    rows: { start: 1, end: 0 },
    columns: { start: 1, end: 0 },
  });

  if (freeze.frozenColumns > 0) {
    viewports.push({
      kind: 'columnHeaderFrozen',
      screenRect: { x: headerW, y: 0, width: Math.min(frozenW, bodyW), height: headerH },
      originX: 0,
      originY: 0,
      rows: { start: 1, end: 0 },
      columns: { start: 1, end: freeze.frozenColumns },
    });
  }

  viewports.push({
    kind: 'columnHeaderScrollable',
    screenRect: {
      x: headerW + frozenW,
      y: 0,
      width: Math.max(0, viewportSize.width - headerW - frozenW),
      height: headerH,
    },
    originX: scrollableX,
    originY: 0,
    rows: { start: 1, end: 0 },
    columns: visibleColumns(layout, scrollableX, mainW, freeze.frozenColumns + 1),
  });

  if (freeze.frozenRows > 0) {
    viewports.push({
      kind: 'rowHeaderFrozen',
      screenRect: { x: 0, y: headerH, width: headerW, height: Math.min(frozenH, bodyH) },
      originX: 0,
      originY: 0,
      rows: { start: 1, end: freeze.frozenRows },
      columns: { start: 1, end: 0 },
    });
  }

  viewports.push({
    kind: 'rowHeaderScrollable',
    screenRect: {
      x: 0,
      y: headerH + frozenH,
      width: headerW,
      height: Math.max(0, viewportSize.height - headerH - frozenH),
    },
    originX: 0,
    originY: scrollableY,
    rows: visibleRows(layout, scrollableY, mainH, freeze.frozenRows + 1),
    columns: { start: 1, end: 0 },
  });

  if (freeze.frozenRows > 0 && freeze.frozenColumns > 0) {
    viewports.push({
      kind: 'frozenTopLeft',
      screenRect: { x: headerW, y: headerH, width: frozenW, height: frozenH },
      originX: 0,
      originY: 0,
      rows: { start: 1, end: freeze.frozenRows },
      columns: { start: 1, end: freeze.frozenColumns },
    });
  }

  if (freeze.frozenRows > 0) {
    viewports.push({
      kind: 'frozenTop',
      screenRect: { x: mainX, y: headerH, width: mainW, height: frozenH },
      originX: scrollableX,
      originY: 0,
      rows: { start: 1, end: freeze.frozenRows },
      columns: visibleColumns(layout, scrollableX, mainW, freeze.frozenColumns + 1),
    });
  }

  if (freeze.frozenColumns > 0) {
    viewports.push({
      kind: 'frozenLeft',
      screenRect: { x: headerW, y: mainY, width: frozenW, height: mainH },
      originX: 0,
      originY: scrollableY,
      rows: visibleRows(layout, scrollableY, mainH, freeze.frozenRows + 1),
      columns: { start: 1, end: freeze.frozenColumns },
    });
  }

  viewports.push({
    kind: 'main',
    screenRect: { x: mainX, y: mainY, width: mainW, height: mainH },
    originX: scrollableX,
    originY: scrollableY,
    rows: visibleRows(layout, scrollableY, mainH, freeze.frozenRows + 1),
    columns: visibleColumns(layout, scrollableX, mainW, freeze.frozenColumns + 1),
  });

  return viewports.filter((viewport) => viewport.screenRect.width > 0 && viewport.screenRect.height > 0);
}

function visibleRows(
  layout: SpreadsheetLayout,
  originY: number,
  height: number,
  minimumRow: number,
): { start: number; end: number } {
  const range = layout.getVisibleRange(0, originY, { width: 1, height }, minimumRow, 1);
  return { start: range.startRow, end: range.endRow };
}

function visibleColumns(
  layout: SpreadsheetLayout,
  originX: number,
  width: number,
  minimumColumn: number,
): { start: number; end: number } {
  const range = layout.getVisibleRange(originX, 0, { width, height: 1 }, 1, minimumColumn);
  return { start: range.startColumn, end: range.endColumn };
}

function containsPoint(rect: SpreadsheetRect, point: SpreadsheetPoint): boolean {
  return (
    point.x >= rect.x &&
    point.y >= rect.y &&
    point.x < rect.x + rect.width &&
    point.y < rect.y + rect.height
  );
}

