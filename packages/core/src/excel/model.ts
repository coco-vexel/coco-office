export type ExcelCellValue = string | number | boolean | null;

export type ExcelCellKind = 'blank' | 'string' | 'number' | 'boolean' | 'formula' | 'error';

export interface ExcelCellModel {
  /** A1-style address, for example `C12`. */
  address: string;
  /** 1-based row index. */
  row: number;
  /** 1-based column index. */
  column: number;
  kind: ExcelCellKind;
  value: ExcelCellValue;
  /** Raw `<v>` text, before shared-string / boolean / number conversion. */
  rawValue?: string;
  /** Formula text without the leading `=`. */
  formula?: string;
  /** OOXML cell type attribute, for example `s`, `inlineStr`, `b`, `str`, `e`. */
  dataType?: string;
  /** OOXML style index from the `s` attribute. */
  styleId?: number;
  /** Shared string table index when `dataType === "s"`. */
  sharedStringIndex?: number;
}

export interface ExcelColumnModel {
  min: number;
  max: number;
  width?: number;
  hidden?: boolean;
  styleId?: number;
  customWidth?: boolean;
}

export interface ExcelRowModel {
  index: number;
  height?: number;
  hidden?: boolean;
  styleId?: number;
  customHeight?: boolean;
}

export interface ExcelMergedRangeModel {
  address: string;
}

export interface ExcelHyperlinkModel {
  ref: string;
  relationshipId?: string;
  target?: string;
  targetMode?: string;
  location?: string;
  display?: string;
  tooltip?: string;
}

export interface ExcelTableColumnModel {
  id: number;
  name: string;
  totalsRowLabel?: string;
  totalsRowFunction?: string;
}

export interface ExcelTableFilterColumnModel {
  colId: number;
  hiddenButton?: boolean;
  showButton?: boolean;
  filters?: string[];
}

export interface ExcelTableSortConditionModel {
  ref?: string;
  descending?: boolean;
  sortBy?: string;
}

export interface ExcelTableSortStateModel {
  ref?: string;
  caseSensitive?: boolean;
  sortMethod?: string;
  conditions: ExcelTableSortConditionModel[];
}

export interface ExcelTableAutoFilterModel {
  ref?: string;
  filterColumns: ExcelTableFilterColumnModel[];
  sortState?: ExcelTableSortStateModel;
}

export interface ExcelTableStyleModel {
  name?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

export interface ExcelTableModel {
  id?: number;
  name?: string;
  displayName?: string;
  relationshipId?: string;
  path?: string;
  ref?: string;
  headerRowCount?: number;
  totalsRowCount?: number;
  totalsRowShown?: boolean;
  autoFilterRef?: string;
  autoFilter?: ExcelTableAutoFilterModel;
  columns: ExcelTableColumnModel[];
  style?: ExcelTableStyleModel;
}

export interface ExcelDataValidationModel {
  sqref?: string;
  type?: string;
  operator?: string;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  errorStyle?: string;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
  formula1?: string;
  formula2?: string;
}

export interface ExcelConditionalFormatRuleModel {
  type?: string;
  priority?: number;
  operator?: string;
  dxfId?: number;
  stopIfTrue?: boolean;
  aboveAverage?: boolean;
  percent?: boolean;
  bottom?: boolean;
  rank?: number;
  text?: string;
  timePeriod?: string;
  formulas: string[];
}

export interface ExcelConditionalFormatModel {
  sqref?: string;
  rules: ExcelConditionalFormatRuleModel[];
}

export interface ExcelCommentModel {
  ref: string;
  authorId?: number;
  author?: string;
  text: string;
  shapeId?: number;
}

export interface ExcelThreadedCommentModel {
  ref: string;
  id?: string;
  personId?: string;
  dateTime?: string;
  text: string;
}

export interface ExcelPaneModel {
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
  activePane?: string;
  state?: string;
}

export interface ExcelSelectionModel {
  pane?: string;
  activeCell?: string;
  activeCellId?: number;
  sqref?: string;
}

export interface ExcelSheetViewModel {
  workbookViewId?: number;
  view?: string;
  tabSelected?: boolean;
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  zoomScale?: number;
  zoomScaleNormal?: number;
  topLeftCell?: string;
  pane?: ExcelPaneModel;
  selections: ExcelSelectionModel[];
}

export interface ExcelDrawingMarkerModel {
  row?: number;
  rowOffset?: number;
  column?: number;
  columnOffset?: number;
}

export interface ExcelDrawingObjectModel {
  kind: 'chart' | 'image' | 'shape' | 'graphicFrame' | 'unknown';
  name?: string;
  relationshipId?: string;
  targetPath?: string;
  targetMode?: string;
}

export interface ExcelDrawingAnchorModel {
  kind: 'oneCell' | 'twoCell' | 'absolute';
  from?: ExcelDrawingMarkerModel;
  to?: ExcelDrawingMarkerModel;
  object?: ExcelDrawingObjectModel;
}

export interface ExcelDrawingModel {
  relationshipId?: string;
  path?: string;
  anchors: ExcelDrawingAnchorModel[];
}

export type ExcelChartType =
  | 'columnClustered'
  | 'columnStacked'
  | 'barClustered'
  | 'barStacked'
  | 'line'
  | 'lineMarkers'
  | 'pie'
  | 'doughnut';

export interface ExcelChartSeriesModel {
  name?: string;
  values: string;
}

export interface ExcelChartPositionModel {
  from: ExcelDrawingMarkerModel;
  to: ExcelDrawingMarkerModel;
}

export interface ExcelChartAddOptions {
  type: ExcelChartType;
  title?: string;
  categoryRange: string;
  series: ExcelChartSeriesModel[];
  position: ExcelChartPositionModel;
}

export interface ExcelChartModel {
  relationshipId?: string;
  drawingPath?: string;
  path?: string;
  type?: ExcelChartType;
  title?: string;
  categoryRange?: string;
  series?: ExcelChartSeriesModel[];
  position?: ExcelChartPositionModel;
  formulaRefs: string[];
}

export interface ExcelImageModel {
  relationshipId?: string;
  drawingPath?: string;
  path?: string;
  targetMode?: string;
  name?: string;
}

export interface ExcelShapeModel {
  drawingPath?: string;
  name?: string;
  anchor: ExcelDrawingAnchorModel;
}

export interface ExcelColorModel {
  /** Canonical RGB color in strict #RRGGBB form. */
  rgb: string;
}

export interface ExcelNumberFormatModel {
  id: number;
  code: string;
}

export interface ExcelFontModel {
  name?: string;
  size?: number;
  family?: number;
  charset?: number;
  scheme?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: ExcelColorModel;
}

export interface ExcelFillModel {
  patternType?: string;
  foregroundColor?: ExcelColorModel;
  backgroundColor?: ExcelColorModel;
}

export interface ExcelBorderSideModel {
  style?: string;
  color?: ExcelColorModel;
}

export interface ExcelBorderModel {
  left?: ExcelBorderSideModel;
  right?: ExcelBorderSideModel;
  top?: ExcelBorderSideModel;
  bottom?: ExcelBorderSideModel;
  diagonal?: ExcelBorderSideModel;
}

export interface ExcelAlignmentModel {
  horizontal?: string;
  vertical?: string;
  textRotation?: number;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
}

export interface ExcelCellFormatModel {
  numberFormatId?: number;
  numberFormatCode?: string;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  xfId?: number;
  applyNumberFormat?: boolean;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  alignment?: ExcelAlignmentModel;
}

export interface ExcelStylesModel {
  numberFormats: ExcelNumberFormatModel[];
  fonts: ExcelFontModel[];
  fills: ExcelFillModel[];
  borders: ExcelBorderModel[];
  cellFormats: ExcelCellFormatModel[];
}

export interface ExcelSheetModel {
  id: string;
  name: string;
  visibility?: string;
  relationshipId?: string;
  path?: string;
  dimension?: string;
  sheetViews: ExcelSheetViewModel[];
  columns: ExcelColumnModel[];
  rows: ExcelRowModel[];
  mergedRanges: ExcelMergedRangeModel[];
  hyperlinks: ExcelHyperlinkModel[];
  tables: ExcelTableModel[];
  dataValidations: ExcelDataValidationModel[];
  conditionalFormats: ExcelConditionalFormatModel[];
  comments: ExcelCommentModel[];
  threadedComments: ExcelThreadedCommentModel[];
  drawings: ExcelDrawingModel[];
  charts: ExcelChartModel[];
  images: ExcelImageModel[];
  cells: ExcelCellModel[];
}

export interface ExcelDefinedNameModel {
  name: string;
  text: string;
  localSheetId?: number;
}

export interface ExcelWorkbookModel {
  sheets: ExcelSheetModel[];
  sharedStrings: string[];
  definedNames: ExcelDefinedNameModel[];
  styles: ExcelStylesModel;
}
