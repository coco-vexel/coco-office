import type { ExcelBorderModel, ExcelCellFormatModel, ExcelStylesModel } from '@coco-vexel/coco-office/excel';
import { getRenderCell } from './render-model.js';
import type { SpreadsheetRange, SpreadsheetRenderCell, SpreadsheetSheetRenderModel } from './types.js';

export interface SpreadsheetResolvedCellStyle {
  font: string;
  fontSize: number;
  underline: boolean;
  fontColor: string;
  fillColor: string;
  wrapText: boolean;
  border: ExcelBorderModel;
  horizontalAlign: CanvasTextAlign;
  verticalAlign: 'top' | 'middle' | 'bottom';
}

export function resolveSpreadsheetCellStyle(
  styles: ExcelStylesModel,
  styleId: number | undefined,
  cell?: SpreadsheetRenderCell,
): SpreadsheetResolvedCellStyle {
  const format = styles.cellFormats[styleId ?? 0] ?? {};
  const font = styles.fonts[format.fontId ?? 0] ?? {};
  const fontSize = font.size ?? 11;
  const fill = styles.fills[format.fillId ?? 0] ?? {};
  const border = styles.borders[format.borderId ?? 0] ?? {};
  return {
    font: canvasFont(font.bold, font.italic, fontSize, font.name ?? 'Arial'),
    fontSize,
    underline: font.underline === true,
    fontColor: font.color?.rgb ?? '#111827',
    fillColor: fill.foregroundColor?.rgb ?? '#FFFFFF',
    wrapText: format.alignment?.wrapText === true,
    border,
    horizontalAlign: resolveHorizontalAlign(format, cell),
    verticalAlign: resolveVerticalAlign(format),
  };
}

export function resolveSpreadsheetRangeBorder(
  model: SpreadsheetSheetRenderModel,
  range: SpreadsheetRange,
): ExcelBorderModel {
  const border: ExcelBorderModel = {};
  for (let column = range.startColumn; column <= range.endColumn; column++) {
    border.top ??= cellBorderSide(model, range.startRow, column, 'top');
    border.bottom ??= cellBorderSide(model, range.endRow, column, 'bottom');
  }
  for (let row = range.startRow; row <= range.endRow; row++) {
    border.left ??= cellBorderSide(model, row, range.startColumn, 'left');
    border.right ??= cellBorderSide(model, row, range.endColumn, 'right');
  }
  return border;
}

function cellBorderSide(
  model: SpreadsheetSheetRenderModel,
  row: number,
  column: number,
  side: 'left' | 'right' | 'top' | 'bottom',
): NonNullable<ExcelBorderModel['left']> | undefined {
  const cell = getRenderCell(model, row, column);
  const format = model.styles.cellFormats[cell?.styleId ?? 0] ?? {};
  return model.styles.borders[format.borderId ?? 0]?.[side];
}

function canvasFont(
  bold: boolean | undefined,
  italic: boolean | undefined,
  size: number,
  family: string,
): string {
  const parts = [];
  if (italic) parts.push('italic');
  if (bold) parts.push('600');
  parts.push(`${size}px`);
  parts.push(quoteFontFamily(family));
  return parts.join(' ');
}

function quoteFontFamily(family: string): string {
  return /\s/.test(family) ? JSON.stringify(family) : family;
}

function resolveHorizontalAlign(format: ExcelCellFormatModel, cell: SpreadsheetRenderCell | undefined): CanvasTextAlign {
  switch (format.alignment?.horizontal) {
    case 'left':
      return 'left';
    case 'center':
      return 'center';
    case 'right':
      return 'right';
  }
  if (cell?.kind === 'number') return 'right';
  if (cell?.kind === 'formula' && typeof cell.value === 'number') return 'right';
  if (cell?.kind === 'boolean' || cell?.kind === 'error') return 'center';
  return 'left';
}

function resolveVerticalAlign(format: ExcelCellFormatModel): 'top' | 'middle' | 'bottom' {
  switch (format.alignment?.vertical) {
    case 'top':
      return 'top';
    case 'center':
    case 'middle':
      return 'middle';
    case 'bottom':
      return 'bottom';
    default:
      return 'bottom';
  }
}
