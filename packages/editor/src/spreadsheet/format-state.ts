import type { ExcelBorderModel, ExcelCellFormatModel, ExcelColorModel } from '@coco-vexel/coco-office/excel';
import { getRenderCell } from './render-model.js';
import { normalizeSelectionRange, type SpreadsheetSelectionState } from './selection.js';
import type { SpreadsheetSheetRenderModel } from './types.js';

export type SpreadsheetMixedValue<T> = T | 'mixed';

export interface SpreadsheetSelectionFormatState {
  font: {
    bold: SpreadsheetMixedValue<boolean>;
    italic: SpreadsheetMixedValue<boolean>;
    underline: SpreadsheetMixedValue<boolean>;
    color: SpreadsheetMixedValue<ExcelColorModel | undefined>;
  };
  fill: {
    foregroundColor: SpreadsheetMixedValue<ExcelColorModel | undefined>;
  };
  numberFormatCode: SpreadsheetMixedValue<string | undefined>;
  alignment: {
    horizontal: SpreadsheetMixedValue<'left' | 'center' | 'right'>;
    vertical: SpreadsheetMixedValue<'top' | 'center' | 'bottom'>;
  };
  border: SpreadsheetMixedValue<ExcelBorderModel>;
}

export function summarizeSpreadsheetSelectionFormat(
  model: SpreadsheetSheetRenderModel,
  selection: SpreadsheetSelectionState,
): SpreadsheetSelectionFormatState {
  const summary = createEmptyFormatSummary();
  const range = normalizeSelectionRange(selection.range);
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      const format = cellFormatAt(model, row, column);
      const font = model.styles.fonts[format.fontId ?? 0] ?? {};
      const fill = model.styles.fills[format.fillId ?? 0] ?? {};
      const border = model.styles.borders[format.borderId ?? 0] ?? {};
      collectSummaryValue(summary.bold, font.bold === true);
      collectSummaryValue(summary.italic, font.italic === true);
      collectSummaryValue(summary.underline, font.underline === true);
      collectSummaryValue(summary.fontColor, font.color);
      collectSummaryValue(summary.fillColor, fill.foregroundColor);
      collectSummaryValue(summary.numberFormatCode, format.numberFormatCode);
      collectSummaryValue(summary.horizontal, normalizeHorizontalAlign(format));
      collectSummaryValue(summary.vertical, normalizeVerticalAlign(format));
      collectSummaryValue(summary.border, border);
    }
  }
  return {
    font: {
      bold: summaryValue(summary.bold),
      italic: summaryValue(summary.italic),
      underline: summaryValue(summary.underline),
      color: summaryValue(summary.fontColor),
    },
    fill: {
      foregroundColor: summaryValue(summary.fillColor),
    },
    numberFormatCode: summaryValue(summary.numberFormatCode),
    alignment: {
      horizontal: summaryValue(summary.horizontal),
      vertical: summaryValue(summary.vertical),
    },
    border: summaryValue(summary.border),
  };
}

interface FormatSummary<T> {
  hasValue: boolean;
  value: T | undefined;
  mixed: boolean;
}

function createEmptyFormatSummary() {
  return {
    bold: emptySummary<boolean>(),
    italic: emptySummary<boolean>(),
    underline: emptySummary<boolean>(),
    fontColor: emptySummary<ExcelColorModel | undefined>(),
    fillColor: emptySummary<ExcelColorModel | undefined>(),
    numberFormatCode: emptySummary<string | undefined>(),
    horizontal: emptySummary<'left' | 'center' | 'right'>(),
    vertical: emptySummary<'top' | 'center' | 'bottom'>(),
    border: emptySummary<ExcelBorderModel>(),
  };
}

function emptySummary<T>(): FormatSummary<T> {
  return { hasValue: false, value: undefined, mixed: false };
}

function collectSummaryValue<T>(summary: FormatSummary<T>, value: T): void {
  if (!summary.hasValue) {
    summary.hasValue = true;
    summary.value = value;
    return;
  }
  if (!formatValuesEqual(summary.value, value)) summary.mixed = true;
}

function summaryValue<T>(summary: FormatSummary<T>): SpreadsheetMixedValue<T> {
  return summary.mixed ? 'mixed' : (summary.value as T);
}

function cellFormatAt(
  model: SpreadsheetSheetRenderModel,
  row: number,
  column: number,
): ExcelCellFormatModel {
  const cell = getRenderCell(model, row, column);
  return model.styles.cellFormats[cell?.styleId ?? 0] ?? {};
}

function normalizeHorizontalAlign(format: ExcelCellFormatModel): 'left' | 'center' | 'right' {
  return format.alignment?.horizontal === 'center' || format.alignment?.horizontal === 'right'
    ? format.alignment.horizontal
    : 'left';
}

function normalizeVerticalAlign(format: ExcelCellFormatModel): 'top' | 'center' | 'bottom' {
  if (format.alignment?.vertical === 'top' || format.alignment?.vertical === 'center') {
    return format.alignment.vertical;
  }
  return 'bottom';
}

function formatValuesEqual(a: unknown, b: unknown): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}
