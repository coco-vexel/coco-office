import { formatCellAddress } from '@coco-vexel/coco-office/excel';
import type {
  SpreadsheetHitResult,
  SpreadsheetLayoutOptions,
  SpreadsheetRange,
  SpreadsheetRect,
  SpreadsheetSheetRenderModel,
  SpreadsheetSize,
} from './types.js';

const DEFAULT_ROW_HEIGHT = 22;
const DEFAULT_COLUMN_WIDTH = 96;
const DEFAULT_ROW_HEADER_WIDTH = 48;
const DEFAULT_COLUMN_HEADER_HEIGHT = 24;
const EXCEL_POINTS_TO_CSS_PIXELS = 96 / 72;
// OOXML column widths are measured in zero-glyph characters, not CSS pixels.
const EXCEL_MAX_DIGIT_WIDTH_PX = 7;
const EXCEL_COLUMN_PADDING_PX = 5;

export class SpreadsheetLayout {
  readonly defaultRowHeight: number;
  readonly defaultColumnWidth: number;
  readonly rowHeaderWidth: number;
  readonly columnHeaderHeight: number;
  readonly rowCount: number;
  readonly columnCount: number;

  private readonly rowOffsets: number[];
  private readonly columnOffsets: number[];
  private readonly rowHeights: number[];
  private readonly columnWidths: number[];

  constructor(
    readonly model: SpreadsheetSheetRenderModel,
    options: SpreadsheetLayoutOptions = {},
  ) {
    this.defaultRowHeight = options.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;
    this.defaultColumnWidth = options.defaultColumnWidth ?? DEFAULT_COLUMN_WIDTH;
    this.rowHeaderWidth = options.rowHeaderWidth ?? DEFAULT_ROW_HEADER_WIDTH;
    this.columnHeaderHeight = options.columnHeaderHeight ?? DEFAULT_COLUMN_HEADER_HEIGHT;
    this.rowCount = Math.max(options.rowCount ?? 0, model.maxRow);
    this.columnCount = Math.max(options.columnCount ?? 0, model.maxColumn);
    this.rowHeights = buildRowHeights(model, this.rowCount, this.defaultRowHeight);
    this.columnWidths = buildColumnWidths(model, this.columnCount, this.defaultColumnWidth);
    this.rowOffsets = buildOffsets(this.rowHeights);
    this.columnOffsets = buildOffsets(this.columnWidths);
  }

  totalWidth(): number {
    return this.columnOffsets[this.columnOffsets.length - 1] ?? 0;
  }

  totalHeight(): number {
    return this.rowOffsets[this.rowOffsets.length - 1] ?? 0;
  }

  getColumnWidth(column: number): number {
    return this.columnWidths[column - 1] ?? this.defaultColumnWidth;
  }

  getRowHeight(row: number): number {
    return this.rowHeights[row - 1] ?? this.defaultRowHeight;
  }

  getColumnX(column: number): number {
    return this.columnOffsets[Math.max(0, column - 1)] ?? 0;
  }

  getRowY(row: number): number {
    return this.rowOffsets[Math.max(0, row - 1)] ?? 0;
  }

  getCellLogicalRect(row: number, column: number): SpreadsheetRect {
    return {
      x: this.getColumnX(column),
      y: this.getRowY(row),
      width: this.getColumnWidth(column),
      height: this.getRowHeight(row),
    };
  }

  getRangeLogicalRect(range: SpreadsheetRange): SpreadsheetRect {
    const x = this.getColumnX(range.startColumn);
    const y = this.getRowY(range.startRow);
    const endX = this.getColumnX(range.endColumn) + this.getColumnWidth(range.endColumn);
    const endY = this.getRowY(range.endRow) + this.getRowHeight(range.endRow);
    return { x, y, width: endX - x, height: endY - y };
  }

  getVisibleRange(
    originX: number,
    originY: number,
    size: SpreadsheetSize,
    minimumRow = 1,
    minimumColumn = 1,
  ): SpreadsheetRange {
    const startColumn = Math.max(minimumColumn, this.columnAt(originX));
    const startRow = Math.max(minimumRow, this.rowAt(originY));
    const endColumn = Math.max(
      startColumn,
      Math.min(this.columnCount, this.columnAt(originX + size.width) + 1),
    );
    const endRow = Math.max(
      startRow,
      Math.min(this.rowCount, this.rowAt(originY + size.height) + 1),
    );
    return { startRow, startColumn, endRow, endColumn };
  }

  rowAt(logicalY: number): number {
    return clampIndex(binarySearchOffset(this.rowOffsets, logicalY), this.rowCount);
  }

  columnAt(logicalX: number): number {
    return clampIndex(binarySearchOffset(this.columnOffsets, logicalX), this.columnCount);
  }

  hitTestLogical(logicalX: number, logicalY: number): SpreadsheetHitResult {
    const column = this.columnAt(logicalX);
    const row = this.rowAt(logicalY);
    if (row < 1 || column < 1 || row > this.rowCount || column > this.columnCount) {
      return { kind: 'none' };
    }
    return { kind: 'cell', row, column, address: formatCellAddress(row, column) };
  }
}

function buildRowHeights(
  model: SpreadsheetSheetRenderModel,
  rowCount: number,
  defaultHeight: number,
): number[] {
  const heights: number[] = [];
  for (let row = 1; row <= rowCount; row++) {
    const meta = model.rowMetaByIndex.get(row);
    // OOXML stores row heights in points; layout geometry is expressed in CSS pixels.
    heights.push(meta?.hidden ? 0 : meta?.height === undefined ? defaultHeight : meta.height * EXCEL_POINTS_TO_CSS_PIXELS);
  }
  return heights;
}

function buildColumnWidths(
  model: SpreadsheetSheetRenderModel,
  columnCount: number,
  defaultWidth: number,
): number[] {
  const widths: number[] = [];
  for (let column = 1; column <= columnCount; column++) {
    const meta = model.columnMetaByIndex.get(column);
    widths.push(meta?.hidden ? 0 : meta?.width === undefined ? defaultWidth : excelColumnWidthToCssPixels(meta.width));
  }
  return widths;
}

/** Convert an OOXML column width to the CSS-pixel geometry used by the renderer. */
export function excelColumnWidthToCssPixels(width: number): number {
  if (!Number.isFinite(width) || width <= 0) return 0;
  return Math.floor(
    ((256 * width + Math.floor(128 / EXCEL_MAX_DIGIT_WIDTH_PX)) / 256) * EXCEL_MAX_DIGIT_WIDTH_PX,
  );
}

/** Convert editor CSS pixels back to the OOXML width Core writes to xlsx. */
export function cssPixelsToExcelColumnWidth(width: number): number {
  if (!Number.isFinite(width) || width <= 0) return 0;
  return Math.floor(((width - EXCEL_COLUMN_PADDING_PX) / EXCEL_MAX_DIGIT_WIDTH_PX) * 256) / 256;
}

function buildOffsets(sizes: number[]): number[] {
  const offsets = [0];
  for (const size of sizes) offsets.push((offsets[offsets.length - 1] ?? 0) + size);
  return offsets;
}

function binarySearchOffset(offsets: number[], value: number): number {
  if (value <= 0) return 1;
  let low = 0;
  let high = offsets.length - 1;
  while (low < high) {
    const mid = Math.floor((low + high + 1) / 2);
    if ((offsets[mid] ?? 0) <= value) low = mid;
    else high = mid - 1;
  }
  return low + 1;
}

function clampIndex(index: number, max: number): number {
  return Math.max(1, Math.min(max, index));
}
