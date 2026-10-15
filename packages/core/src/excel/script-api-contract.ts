import type { ExcelWorkbookCore } from './excel-workbook.js';
import type {
  ExcelCell,
  ExcelRange,
  ExcelRangeAlignmentFormat,
  ExcelRangeBorderFormat,
  ExcelRangeFillFormat,
  ExcelRangeFontFormat,
  ExcelRangeFormat,
} from './range.js';
import type {
  ExcelChart,
  ExcelChartCollection,
  ExcelDrawingCollection,
  ExcelFormulaRefCollection,
  ExcelImageCollection,
  ExcelShapeCollection,
} from './objects.js';
import type { ExcelWorksheet, ExcelWorksheetCollection } from './worksheet.js';

/** Receivers exposed inside ExcelWorkbook.executeCode scripts. */
export interface ExcelScriptApiReceiverMap {
  workbook: ExcelWorkbookCore;
  'workbook.worksheets': ExcelWorksheetCollection;
  sheet: ExcelWorksheet;
  'sheet.drawings': ExcelDrawingCollection;
  'sheet.images': ExcelImageCollection;
  'sheet.shapes': ExcelShapeCollection;
  'sheet.charts': ExcelChartCollection;
  range: ExcelRange;
  'range.format': ExcelRangeFormat;
  'range.format.font': ExcelRangeFontFormat;
  'range.format.fill': ExcelRangeFillFormat;
  'range.format.border': ExcelRangeBorderFormat;
  'range.format.alignment': ExcelRangeAlignmentFormat;
  cell: ExcelCell;
  chart: ExcelChart;
  'chart.formulaRefs': ExcelFormulaRefCollection;
}

export type ExcelScriptApiReceiver = keyof ExcelScriptApiReceiverMap;

type MethodName<T> = {
  [K in keyof T]-?: T[K] extends (...args: never[]) => unknown ? K : never;
}[keyof T] &
  string;

/** Type-safe receiver-qualified method path for the synchronous Excel script runtime. */
export type ExcelScriptApiPath = {
  [R in ExcelScriptApiReceiver]: `${R}.${MethodName<ExcelScriptApiReceiverMap[R]>}`;
}[ExcelScriptApiReceiver];
