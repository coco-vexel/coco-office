import type {
  ExcelCellKind,
  ExcelCellModel,
  ExcelCellValue,
  ExcelColumnModel,
  ExcelMergedRangeModel,
  ExcelRowModel,
  ExcelSheetModel,
  ExcelStylesModel,
} from '@coco-vexel/coco-office/excel';

export interface SpreadsheetPoint {
  x: number;
  y: number;
}

export interface SpreadsheetSize {
  width: number;
  height: number;
}

export interface SpreadsheetRect extends SpreadsheetPoint, SpreadsheetSize {}

export interface SpreadsheetRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

export interface SpreadsheetRenderCell {
  row: number;
  column: number;
  address: string;
  kind: ExcelCellKind;
  value: ExcelCellValue;
  displayText: string;
  formula?: string;
  styleId?: number;
  source: ExcelCellModel;
}

export interface SpreadsheetMergeInfo {
  address: string;
  range: SpreadsheetRange;
  masterAddress: string;
  source: ExcelMergedRangeModel;
}

export interface SpreadsheetSheetRenderModel {
  sheetId: string;
  sheetName: string;
  source: ExcelSheetModel;
  styles: ExcelStylesModel;
  cellsByAddress: Map<string, SpreadsheetRenderCell>;
  cellsByRowColumn: Map<string, SpreadsheetRenderCell>;
  rowMetaByIndex: Map<number, ExcelRowModel>;
  columnMetaByIndex: Map<number, ExcelColumnModel>;
  mergeMasterByAddress: Map<string, SpreadsheetMergeInfo>;
  mergeFollowerByAddress: Map<string, string>;
  maxRow: number;
  maxColumn: number;
}

export interface SpreadsheetLayoutOptions {
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
  rowHeaderWidth?: number;
  columnHeaderHeight?: number;
  rowCount?: number;
  columnCount?: number;
}

export interface SpreadsheetHitCell {
  kind: 'cell';
  row: number;
  column: number;
  address: string;
}

export interface SpreadsheetHitColumnHeader {
  kind: 'columnHeader';
  column: number;
}

export interface SpreadsheetHitRowHeader {
  kind: 'rowHeader';
  row: number;
}

export interface SpreadsheetHitCorner {
  kind: 'corner';
}

export interface SpreadsheetHitNone {
  kind: 'none';
}

export type SpreadsheetHitResult =
  | SpreadsheetHitCell
  | SpreadsheetHitColumnHeader
  | SpreadsheetHitRowHeader
  | SpreadsheetHitCorner
  | SpreadsheetHitNone;

export type SpreadsheetViewportKind =
  | 'main'
  | 'frozenTop'
  | 'frozenLeft'
  | 'frozenTopLeft'
  | 'columnHeaderScrollable'
  | 'columnHeaderFrozen'
  | 'rowHeaderScrollable'
  | 'rowHeaderFrozen'
  | 'corner';

export interface SpreadsheetViewport {
  kind: SpreadsheetViewportKind;
  screenRect: SpreadsheetRect;
  originX: number;
  originY: number;
  rows: { start: number; end: number };
  columns: { start: number; end: number };
}

export interface SpreadsheetScrollState {
  x: number;
  y: number;
}

export interface SpreadsheetFreezeState {
  frozenRows: number;
  frozenColumns: number;
}

export interface SpreadsheetViewportState {
  viewportSize: SpreadsheetSize;
  scroll: SpreadsheetScrollState;
  freeze: SpreadsheetFreezeState;
}

export function rowColumnKey(row: number, column: number): string {
  return `${row}:${column}`;
}

