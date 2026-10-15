import { columnIndexToName, formatCellAddress, type ExcelBorderModel } from '@coco-vexel/coco-office/excel';
import { getRenderCell } from './render-model.js';
import {
  resolveSpreadsheetCellStyle,
  resolveSpreadsheetRangeBorder,
  type SpreadsheetResolvedCellStyle,
} from './style-resolver.js';
import type { SpreadsheetLayout } from './layout.js';
import type { SpreadsheetViewportManager } from './viewport.js';
import type { SpreadsheetRange, SpreadsheetRect, SpreadsheetViewport } from './types.js';

export interface SpreadsheetCanvasRendererOptions {
  baseCanvas: HTMLCanvasElement;
  headerCanvas: HTMLCanvasElement;
  overlayCanvas: HTMLCanvasElement;
}

export class SpreadsheetCanvasRenderer {
  private readonly baseCtx: CanvasRenderingContext2D;
  private readonly headerCtx: CanvasRenderingContext2D;
  private readonly overlayCtx: CanvasRenderingContext2D;
  private cssWidth = 1;
  private cssHeight = 1;
  private scaleX = 1;
  private scaleY = 1;
  private headerFiltersEnabled = true;
  private headerFilterColumns: Set<number> | null = null;
  private activeHeaderFilterColumns = new Set<number>();

  constructor(private readonly options: SpreadsheetCanvasRendererOptions) {
    this.baseCtx = get2dContext(options.baseCanvas);
    this.headerCtx = get2dContext(options.headerCanvas);
    this.overlayCtx = get2dContext(options.overlayCanvas);
  }

  resize(width: number, height: number, devicePixelRatio = globalThis.devicePixelRatio || 1): void {
    this.cssWidth = Math.max(1, width);
    this.cssHeight = Math.max(1, height);
    const ratio = Math.max(1, devicePixelRatio || 1);
    for (const canvas of [
      this.options.baseCanvas,
      this.options.headerCanvas,
      this.options.overlayCanvas,
    ]) {
      canvas.style.width = `${this.cssWidth}px`;
      canvas.style.height = `${this.cssHeight}px`;
      canvas.width = Math.max(1, Math.ceil(this.cssWidth * ratio));
      canvas.height = Math.max(1, Math.ceil(this.cssHeight * ratio));
      const ctx = get2dContext(canvas);
      this.scaleX = canvas.width / this.cssWidth;
      this.scaleY = canvas.height / this.cssHeight;
      ctx.setTransform(this.scaleX, 0, 0, this.scaleY, 0, 0);
      ctx.imageSmoothingEnabled = false;
    }
  }

  setHeaderFiltersEnabled(enabled: boolean): void {
    this.headerFiltersEnabled = enabled;
  }

  setHeaderFilterColumns(columns: readonly number[] | null): void {
    this.headerFilterColumns = columns === null ? null : new Set(columns);
  }

  setActiveHeaderFilterColumns(columns: readonly number[]): void {
    this.activeHeaderFilterColumns = new Set(columns);
  }

  render(layout: SpreadsheetLayout, viewportManager: SpreadsheetViewportManager): void {
    this.clearAll(layout, viewportManager);
    this.drawCells(layout, viewportManager);
    this.drawHeaders(layout, viewportManager);
  }

  clearOverlay(): void {
    clearCanvas(this.overlayCtx, this.options.overlayCanvas);
  }

  renderSelection(
    layout: SpreadsheetLayout,
    viewportManager: SpreadsheetViewportManager,
    range: SpreadsheetRange,
    options: { showFillHandle?: boolean } = {},
  ): void {
    this.clearOverlay();
    const ctx = this.overlayCtx;
    for (const viewport of viewportManager.viewports) {
      if (!isCellViewport(viewport)) continue;
      const intersection = intersectRange(range, {
        startRow: viewport.rows.start,
        endRow: viewport.rows.end,
        startColumn: viewport.columns.start,
        endColumn: viewport.columns.end,
      });
      if (!intersection) continue;
      const logical = layout.getRangeLogicalRect(intersection);
      const rect = logicalRectToScreen(logical, viewport);
      withClip(ctx, viewport.screenRect, () => {
        ctx.fillStyle = 'rgba(37, 99, 235, 0.08)';
        ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
        ctx.strokeStyle = '#2563EB';
        ctx.lineWidth = 2 / this.scaleX;
        ctx.setLineDash([]);
        strokeCrispRect(ctx, insetRect(rect, 1), this.scaleX, this.scaleY);
      });
    }
    if (options.showFillHandle) this.drawFillHandle(layout, viewportManager, range);
  }

  fillHandleRect(
    layout: SpreadsheetLayout,
    viewportManager: SpreadsheetViewportManager,
    range: SpreadsheetRange,
  ): SpreadsheetRect | null {
    for (const viewport of viewportManager.viewports) {
      if (!isCellViewport(viewport)) continue;
      if (
        range.endRow < viewport.rows.start ||
        range.endRow > viewport.rows.end ||
        range.endColumn < viewport.columns.start ||
        range.endColumn > viewport.columns.end
      ) {
        continue;
      }
      const cell = viewportManager.cellScreenRect(range.endRow, range.endColumn, viewport);
      const size = 7;
      return {
        x: cell.x + cell.width - size / 2,
        y: cell.y + cell.height - size / 2,
        width: size,
        height: size,
      };
    }
    return null;
  }

  renderCopiedRange(
    layout: SpreadsheetLayout,
    viewportManager: SpreadsheetViewportManager,
    range: SpreadsheetRange,
  ): void {
    this.renderDashedRange(layout, viewportManager, range, {
      stroke: '#1D4ED8',
    });
  }

  renderDashedRange(
    layout: SpreadsheetLayout,
    viewportManager: SpreadsheetViewportManager,
    range: SpreadsheetRange,
    options: { stroke: string; fill?: string },
  ): void {
    const ctx = this.overlayCtx;
    for (const viewport of viewportManager.viewports) {
      if (!isCellViewport(viewport)) continue;
      const intersection = intersectRange(range, {
        startRow: viewport.rows.start,
        endRow: viewport.rows.end,
        startColumn: viewport.columns.start,
        endColumn: viewport.columns.end,
      });
      if (!intersection) continue;
      const logical = layout.getRangeLogicalRect(intersection);
      const rect = logicalRectToScreen(logical, viewport);
      withClip(ctx, viewport.screenRect, () => {
        ctx.save();
        if (options.fill) {
          ctx.fillStyle = options.fill;
          ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
        }
        ctx.strokeStyle = options.stroke;
        ctx.lineWidth = 2 / this.scaleX;
        ctx.setLineDash([4, 3]);
        ctx.lineDashOffset = 0;
        strokeCrispRect(ctx, insetRect(rect, 2), this.scaleX, this.scaleY);
        ctx.restore();
      });
    }
  }

  private clearAll(layout: SpreadsheetLayout, viewportManager: SpreadsheetViewportManager): void {
    void layout;
    void viewportManager;
    clearCanvas(this.baseCtx, this.options.baseCanvas);
    clearCanvas(this.headerCtx, this.options.headerCanvas);
    clearCanvas(this.overlayCtx, this.options.overlayCanvas);
  }

  private drawFillHandle(
    layout: SpreadsheetLayout,
    viewportManager: SpreadsheetViewportManager,
    range: SpreadsheetRange,
  ): void {
    const rect = this.fillHandleRect(layout, viewportManager, range);
    if (!rect) return;
    const ctx = this.overlayCtx;
    ctx.save();
    ctx.fillStyle = '#2563EB';
    ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    ctx.strokeStyle = '#FFFFFF';
    ctx.lineWidth = 1 / this.scaleX;
    strokeCrispRect(ctx, insetRect(rect, 0.5), this.scaleX, this.scaleY);
    ctx.restore();
  }

  private drawCells(layout: SpreadsheetLayout, viewportManager: SpreadsheetViewportManager): void {
    const ctx = this.baseCtx;
    ctx.save();
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, this.cssWidth, this.cssHeight);
    ctx.restore();

    for (const viewport of viewportManager.viewports) {
      if (!isCellViewport(viewport)) continue;
      withClip(ctx, viewport.screenRect, () => {
        for (let row = viewport.rows.start; row <= viewport.rows.end; row++) {
          for (let column = viewport.columns.start; column <= viewport.columns.end; column++) {
            const address = formatCellAddress(row, column);
            const mergeMaster = layout.model.mergeFollowerByAddress.get(address);
            if (mergeMaster) continue;
            this.drawCell(ctx, layout, viewportManager, viewport, row, column, {
              drawBorder: false,
              reserveTrailingSpace: this.headerFiltersEnabled && row === 1,
            });
            if (
              this.headerFiltersEnabled &&
              row === 1 &&
              (this.headerFilterColumns === null || this.headerFilterColumns.has(column))
            ) {
              const logical = layout.getCellLogicalRect(row, column);
              const rect = {
                x: viewport.screenRect.x + logical.x - viewport.originX,
                y: viewport.screenRect.y + logical.y - viewport.originY,
                width: logical.width,
                height: logical.height,
              };
              drawFilterIcon(
                ctx,
                rect,
                this.scaleX,
                this.scaleY,
                this.activeHeaderFilterColumns.has(column),
              );
            }
          }
        }
        for (let row = viewport.rows.start; row <= viewport.rows.end; row++) {
          for (let column = viewport.columns.start; column <= viewport.columns.end; column++) {
            const address = formatCellAddress(row, column);
            const mergeMaster = layout.model.mergeFollowerByAddress.get(address);
            if (mergeMaster) continue;
            this.drawCell(ctx, layout, viewportManager, viewport, row, column, {
              drawContent: false,
            });
          }
        }
      });
    }
  }

  private drawCell(
    ctx: CanvasRenderingContext2D,
    layout: SpreadsheetLayout,
    viewportManager: SpreadsheetViewportManager,
    viewport: SpreadsheetViewport,
    row: number,
    column: number,
    options: { drawBorder?: boolean; drawContent?: boolean; reserveTrailingSpace?: boolean } = {},
  ): void {
    const cell = getRenderCell(layout.model, row, column);
    const merge = cell ? layout.model.mergeMasterByAddress.get(cell.address) : undefined;
    const rect = merge
      ? logicalRectToScreen(layout.getRangeLogicalRect(merge.range), viewport)
      : viewportManager.cellScreenRect(row, column, viewport);
    const style = resolveSpreadsheetCellStyle(layout.model.styles, cell?.styleId, cell);
    const border = merge ? resolveSpreadsheetRangeBorder(layout.model, merge.range) : style.border;
    if (options.drawContent !== false) {
      ctx.fillStyle = style.fillColor;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.strokeStyle = '#E5E7EB';
      ctx.lineWidth = 1 / this.scaleX;
      strokeCrispRect(ctx, rect, this.scaleX, this.scaleY);
    }
    if (options.drawBorder !== false) drawCellBorder(ctx, rect, border, this.scaleX, this.scaleY);
    if (options.drawContent === false || !cell?.displayText) return;

    ctx.save();
    ctx.font = style.font;
    ctx.fillStyle = style.fontColor;
    ctx.textAlign = style.horizontalAlign;
    ctx.textBaseline = 'middle';
    const padding = 6;
    const trailingSpace = options.reserveTrailingSpace && rect.width >= 28 ? 22 : 0;
    const contentRect = {
      x: rect.x,
      y: rect.y,
      width: Math.max(1, rect.width - trailingSpace),
      height: rect.height,
    };
    const textX =
      style.horizontalAlign === 'right'
        ? contentRect.x + contentRect.width - padding
        : style.horizontalAlign === 'center'
          ? contentRect.x + contentRect.width / 2
          : contentRect.x + padding;
    withClip(ctx, contentRect, () => {
      const lines = style.wrapText || cell.displayText.includes('\n')
        ? wrapCanvasText(ctx, cell.displayText, Math.max(1, contentRect.width - padding * 2))
        : [cell.displayText];
      const lineHeight = style.fontSize * 1.2;
      const blockHeight = lines.length * lineHeight;
      const firstTextY = textBlockStartY(rect, style, padding, blockHeight, lineHeight);
      for (let index = 0; index < lines.length; index++) {
        const line = lines[index]!;
        const lineY = firstTextY + index * lineHeight;
        ctx.fillText(line, textX, lineY);
        if (style.underline) drawTextUnderline(ctx, line, textX, lineY, style, this.scaleX, this.scaleY);
      }
    });
    ctx.restore();
  }

  private drawHeaders(layout: SpreadsheetLayout, viewportManager: SpreadsheetViewportManager): void {
    const ctx = this.headerCtx;
    for (const viewport of viewportManager.viewports) {
      withClip(ctx, viewport.screenRect, () => {
        if (viewport.kind === 'corner') {
          drawHeaderRect(ctx, viewport.screenRect, '', this.scaleX, this.scaleY);
        } else if (
          viewport.kind === 'columnHeaderFrozen' ||
          viewport.kind === 'columnHeaderScrollable'
        ) {
          for (let column = viewport.columns.start; column <= viewport.columns.end; column++) {
            const logical = layout.getCellLogicalRect(1, column);
            const rect = {
              x: viewport.screenRect.x + logical.x - viewport.originX,
              y: viewport.screenRect.y,
              width: logical.width,
              height: viewport.screenRect.height,
            };
            drawHeaderRect(ctx, rect, columnIndexToName(column), this.scaleX, this.scaleY);
          }
        } else if (
          viewport.kind === 'rowHeaderFrozen' ||
          viewport.kind === 'rowHeaderScrollable'
        ) {
          for (let row = viewport.rows.start; row <= viewport.rows.end; row++) {
            const logical = layout.getCellLogicalRect(row, 1);
            const rect = {
              x: viewport.screenRect.x,
              y: viewport.screenRect.y + logical.y - viewport.originY,
              width: viewport.screenRect.width,
              height: logical.height,
            };
            drawHeaderRect(ctx, rect, String(row), this.scaleX, this.scaleY);
          }
        }
      });
    }
  }
}

function drawFilterIcon(
  ctx: CanvasRenderingContext2D,
  rect: SpreadsheetRect,
  scaleX: number,
  scaleY: number,
  active = false,
): void {
  if (rect.width < 28) return;
  const centerX = rect.x + rect.width - 10;
  const centerY = rect.y + rect.height / 2;
  ctx.save();
  ctx.strokeStyle = active ? '#2563EB' : '#64748B';
  ctx.lineWidth = 1 / scaleX;
  ctx.lineCap = 'round';
  ctx.beginPath();
  ctx.moveTo(centerX - 4, centerY - 3);
  ctx.lineTo(centerX + 4, centerY - 3);
  ctx.lineTo(centerX + 1, centerY + 1);
  ctx.lineTo(centerX + 1, centerY + 4);
  ctx.lineTo(centerX - 1, centerY + 3);
  ctx.lineTo(centerX - 1, centerY + 1);
  ctx.closePath();
  ctx.stroke();
  ctx.restore();
  void scaleY;
}

function intersectRange(a: SpreadsheetRange, b: SpreadsheetRange): SpreadsheetRange | null {
  const startRow = Math.max(a.startRow, b.startRow);
  const startColumn = Math.max(a.startColumn, b.startColumn);
  const endRow = Math.min(a.endRow, b.endRow);
  const endColumn = Math.min(a.endColumn, b.endColumn);
  if (startRow > endRow || startColumn > endColumn) return null;
  return { startRow, startColumn, endRow, endColumn };
}

function get2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('SpreadsheetCanvasRenderer: 2D canvas context is unavailable');
  return ctx;
}

function clearCanvas(ctx: CanvasRenderingContext2D, canvas: HTMLCanvasElement): void {
  const width = Number.parseFloat(canvas.style.width) || canvas.width;
  const height = Number.parseFloat(canvas.style.height) || canvas.height;
  ctx.clearRect(0, 0, width, height);
}

function withClip(ctx: CanvasRenderingContext2D, rect: SpreadsheetRect, draw: () => void): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();
  draw();
  ctx.restore();
}

function isCellViewport(viewport: SpreadsheetViewport): boolean {
  return (
    viewport.kind === 'main' ||
    viewport.kind === 'frozenTop' ||
    viewport.kind === 'frozenLeft' ||
    viewport.kind === 'frozenTopLeft'
  );
}

function logicalRectToScreen(rect: SpreadsheetRect, viewport: SpreadsheetViewport): SpreadsheetRect {
  return {
    x: viewport.screenRect.x + rect.x - viewport.originX,
    y: viewport.screenRect.y + rect.y - viewport.originY,
    width: rect.width,
    height: rect.height,
  };
}

function drawHeaderRect(
  ctx: CanvasRenderingContext2D,
  rect: SpreadsheetRect,
  text: string,
  scaleX: number,
  scaleY: number,
): void {
  ctx.fillStyle = '#F3F4F6';
  ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
  ctx.strokeStyle = '#D1D5DB';
  ctx.lineWidth = 1 / scaleX;
  strokeCrispRect(ctx, rect, scaleX, scaleY);
  if (!text) return;
  ctx.font = '12px Arial';
  ctx.fillStyle = '#374151';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, rect.x + rect.width / 2, rect.y + rect.height / 2);
}

function drawCellBorder(
  ctx: CanvasRenderingContext2D,
  rect: SpreadsheetRect,
  border: ExcelBorderModel,
  scaleX: number,
  scaleY: number,
): void {
  if (!border.left && !border.right && !border.top && !border.bottom) return;
  ctx.save();
  ctx.lineWidth = 1 / scaleX;
  ctx.setLineDash([]);
  if (border.top) {
    ctx.strokeStyle = border.top.color?.rgb ?? '#111827';
    drawCrispLine(ctx, rect.x, rect.y, rect.x + rect.width, rect.y, scaleX, scaleY);
  }
  if (border.right) {
    ctx.strokeStyle = border.right.color?.rgb ?? '#111827';
    drawCrispLine(ctx, rect.x + rect.width, rect.y, rect.x + rect.width, rect.y + rect.height, scaleX, scaleY);
  }
  if (border.bottom) {
    ctx.strokeStyle = border.bottom.color?.rgb ?? '#111827';
    drawCrispLine(ctx, rect.x, rect.y + rect.height, rect.x + rect.width, rect.y + rect.height, scaleX, scaleY);
  }
  if (border.left) {
    ctx.strokeStyle = border.left.color?.rgb ?? '#111827';
    drawCrispLine(ctx, rect.x, rect.y, rect.x, rect.y + rect.height, scaleX, scaleY);
  }
  ctx.restore();
}

function textBlockStartY(
  rect: SpreadsheetRect,
  style: SpreadsheetResolvedCellStyle,
  padding: number,
  blockHeight: number,
  lineHeight: number,
): number {
  if (style.verticalAlign === 'top') return rect.y + padding + lineHeight / 2;
  if (style.verticalAlign === 'bottom') return rect.y + rect.height - padding - blockHeight + lineHeight / 2;
  return rect.y + (rect.height - blockHeight) / 2 + lineHeight / 2;
}

function wrapCanvasText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split(/\r?\n/)) {
    if (paragraph === '') {
      lines.push('');
      continue;
    }
    let line = '';
    for (const character of paragraph) {
      const candidate = line + character;
      if (line && ctx.measureText(candidate).width > maxWidth) {
        lines.push(line);
        line = character;
      } else {
        line = candidate;
      }
    }
    if (line) lines.push(line);
  }
  return lines.length > 0 ? lines : [''];
}

function drawCrispLine(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  scaleX: number,
  scaleY: number,
): void {
  ctx.beginPath();
  ctx.moveTo(snapStrokeCoordinate(x1, scaleX), snapStrokeCoordinate(y1, scaleY));
  ctx.lineTo(snapStrokeCoordinate(x2, scaleX), snapStrokeCoordinate(y2, scaleY));
  ctx.stroke();
}

function drawTextUnderline(
  ctx: CanvasRenderingContext2D,
  text: string,
  textX: number,
  textY: number,
  style: SpreadsheetResolvedCellStyle,
  scaleX: number,
  scaleY: number,
): void {
  const metrics = ctx.measureText(text);
  const width = metrics.width;
  const startX =
    style.horizontalAlign === 'right'
      ? textX - width
      : style.horizontalAlign === 'center'
        ? textX - width / 2
        : textX;
  const underlineOffset = Math.max(2, (metrics.actualBoundingBoxDescent || style.fontSize * 0.18) + 1);
  ctx.save();
  ctx.strokeStyle = style.fontColor;
  ctx.lineWidth = 1 / scaleX;
  drawCrispLine(ctx, startX, textY + underlineOffset, startX + width, textY + underlineOffset, scaleX, scaleY);
  ctx.restore();
}

function insetRect(rect: SpreadsheetRect, inset: number): SpreadsheetRect {
  return {
    x: rect.x + inset,
    y: rect.y + inset,
    width: Math.max(0, rect.width - inset * 2),
    height: Math.max(0, rect.height - inset * 2),
  };
}

function strokeCrispRect(
  ctx: CanvasRenderingContext2D,
  rect: SpreadsheetRect,
  scaleX: number,
  scaleY: number,
): void {
  const x = snapStrokeCoordinate(rect.x, scaleX);
  const y = snapStrokeCoordinate(rect.y, scaleY);
  const right = snapStrokeCoordinate(rect.x + rect.width, scaleX);
  const bottom = snapStrokeCoordinate(rect.y + rect.height, scaleY);
  ctx.strokeRect(x, y, Math.max(0, right - x), Math.max(0, bottom - y));
}

function snapStrokeCoordinate(value: number, scale: number): number {
  return (Math.round(value * scale) + 0.5) / scale;
}
