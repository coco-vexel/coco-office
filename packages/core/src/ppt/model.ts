export type PptShapeKind =
  | 'placeholder'
  | 'textBox'
  | 'picture'
  | 'table'
  | 'chart'
  | 'smartArt'
  | 'geometricShape'
  | 'line'
  | 'group'
  | 'media'
  | 'oleObject'
  | 'unknown';

export interface PptSlideSize {
  width: number;
  height: number;
}

export interface PptShapeGeometry {
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
}

export interface PptTextBodyModel {
  text: string;
  style?: PptTextStyle;
  body?: PptTextBodyProperties;
  paragraphs?: PptTextParagraph[];
}

export interface PptTextBodyProperties {
  anchor?: 'top' | 'middle' | 'bottom';
  autoFit?: 'none' | 'shape' | 'normal';
  wrap?: 'none' | 'square';
  vertical?: 'eastAsian' | 'vertical';
  insetLeft?: number;
  insetTop?: number;
  insetRight?: number;
  insetBottom?: number;
}

export interface PptTextStyle {
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  textFill?: 'none';
  textGradient?: PptGradientFillModel;
  textOutlineColor?: string;
  textShadow?: PptShadowModel;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  baseline?: number;
  highlight?: string;
  alignment?: PptParagraphAlignment;
  lineHeight?: number;
  characterSpacing?: number;
}

export type PptParagraphAlignment = 'left' | 'center' | 'right' | 'justify' | 'distributed';

export interface PptHyperlinkModel {
  relationshipId?: string;
  target?: string;
  targetMode?: string;
  action?: string;
  tooltip?: string;
}

export interface PptHyperlinkWriteOptions {
  target: string;
  targetMode?: string;
  action?: string;
  tooltip?: string;
}

export interface PptTextRun {
  text: string;
  style?: PptTextStyle;
  hyperlink?: PptHyperlinkModel;
}

export interface PptTextRunWrite {
  text: string;
  style?: Partial<PptTextStyle>;
}

export interface PptTextParagraph {
  text: string;
  style?: Pick<PptTextStyle, 'alignment'> & {
    bullet?: string;
    bulletColor?: string;
    bulletFontFamily?: string;
    bulletSize?: number;
    lineHeight?: number;
    marginLeft?: number;
    textIndent?: number;
    spaceBefore?: number;
    spaceAfter?: number;
  };
  runs: PptTextRun[];
}

export interface PptTextParagraphWrite {
  text?: string;
  style?: Pick<PptTextStyle, 'alignment'> & {
    bullet?: string;
    bulletColor?: string;
    bulletFontFamily?: string;
    bulletSize?: number;
    lineHeight?: number;
    marginLeft?: number;
    textIndent?: number;
    spaceBefore?: number;
    spaceAfter?: number;
  };
  runs?: PptTextRunWrite[];
}

export interface PptShapeStyle {
  fillColor?: string;
  fillGradient?: PptGradientFillModel;
  opacity?: number;
  lineColor?: string;
  lineGradient?: PptGradientFillModel;
  lineOpacity?: number;
  lineWidth?: number;
  lineCap?: string;
  lineDash?: string;
  lineHeadEnd?: string;
  lineTailEnd?: string;
  shadow?: PptShadowModel;
}

export interface PptSlideBackgroundModel {
  fillColor?: string;
  fillGradient?: PptGradientFillModel;
  imagePath?: string;
}

export interface PptGradientFillModel {
  colors: string[];
  angle?: number;
  positions?: number[];
  opacities?: number[];
  flip?: 'none' | 'x' | 'y' | 'xy';
  rotWithShape?: boolean;
  scaled?: boolean;
  tileRect?: {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
  };
}

export interface PptShadowModel {
  color?: string;
  opacity?: number;
  blur?: number;
  distance?: number;
  direction?: number;
}

export interface PptCustomGeometryModel {
  paths: PptCustomGeometryPathModel[];
}

export interface PptShape3dModel {
  material?: string;
  extrusionHeight?: number;
}

export interface PptCustomGeometryPathModel {
  width?: number;
  height?: number;
  commands: PptCustomGeometryCommand[];
}

export type PptCustomGeometryCommand =
  | { kind: 'moveTo'; x: number; y: number }
  | { kind: 'lineTo'; x: number; y: number }
  | {
      kind: 'cubicBezierTo';
      x1: number;
      y1: number;
      x2: number;
      y2: number;
      x: number;
      y: number;
    }
  | {
      kind: 'arcTo';
      wR: number;
      hR: number;
      stAng: number;
      swAng: number;
    }
  | { kind: 'close' };

export interface PptPlaceholderModel {
  type?: string;
  index?: string;
  size?: string;
  orientation?: string;
}

export interface PptImageModel {
  relationshipId?: string;
  targetPath?: string;
  contentType?: string;
  filename?: string;
  crop?: PptImageCropModel;
  fillRect?: PptImageCropModel;
  opacity?: number;
}

export interface PptImageStyle {
  crop?: PptImageCropModel;
  fillRect?: PptImageCropModel;
  opacity?: number;
}

export interface PptImageCropModel {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}

export interface PptChartModel {
  relationshipId?: string;
  targetPath?: string;
  type?: string;
  grouping?: string;
  title?: string;
  seriesNames?: string[];
  data?: PptChartData;
  plotLayout?: PptChartLayoutModel;
  axes?: PptChartAxesModel;
}

export interface PptChartData {
  title?: string;
  series: PptChartSeriesData[];
}

export interface PptChartSeriesData {
  name?: string;
  type?: string;
  chartGroup?: number;
  categories?: string[];
  values: number[];
  fillColor?: string;
  fillGradient?: PptGradientFillModel;
  pointColors?: string[];
  borderColor?: string;
  borderWidth?: number;
  lineColor?: string;
  lineWidth?: number;
  marker?: PptChartMarkerModel;
  showValueLabels?: boolean;
  showCategoryNameLabels?: boolean;
  showPercentLabels?: boolean;
  dataLabelLayouts?: PptChartDataLabelLayoutModel[];
  smooth?: boolean;
}

export interface PptChartDataLabelLayoutModel {
  index: number;
  x?: number;
  y?: number;
}

export interface PptChartLayoutModel {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}

export interface PptChartAxesModel {
  category?: PptChartAxisModel;
  value?: PptChartAxisModel;
}

export interface PptChartAxisModel {
  deleted?: boolean;
  min?: number;
  max?: number;
  majorUnit?: number;
  majorGridlines?: boolean;
  lineColor?: string;
  lineWidth?: number;
  gridlineColor?: string;
  gridlineWidth?: number;
  tickLabelPosition?: string;
}

export interface PptChartMarkerModel {
  symbol?: string;
  size?: number;
  fillColor?: string;
  lineColor?: string;
  lineWidth?: number;
}

export interface PptShapeMediaModel {
  relationshipId?: string;
  targetPath?: string;
  mediaType?: 'audio' | 'video' | 'media';
  contentType?: string;
  filename?: string;
}

export interface PptOleObjectModel {
  relationshipId?: string;
  targetPath?: string;
  progId?: string;
  name?: string;
}

export interface PptSmartArtModel {
  relationshipIds?: Record<string, string>;
  targetPaths?: Record<string, string>;
}

export interface PptThemeModel {
  relationshipId?: string;
  path?: string;
  name?: string;
  colorScheme?: Record<string, string>;
  fontScheme?: {
    name?: string;
    majorLatin?: string;
    majorEa?: string;
    majorCs?: string;
    minorLatin?: string;
    minorEa?: string;
    minorCs?: string;
  };
}

export interface PptTableModel {
  values: string[][];
  rowCount: number;
  columnCount: number;
  columnWidths?: number[];
  rowHeights?: number[];
  cellStyles?: PptTableCellStyle[][];
  mergedCells?: PptTableCellMerge[];
}

export interface PptTableCellStyle {
  fillColor?: string;
  fillGradient?: PptGradientFillModel;
  borders?: PptTableCellBorders;
  body?: PptTextBodyProperties;
  textStyle?: PptTextStyle;
}

export interface PptTableCellBorders {
  left?: PptTableCellBorder;
  right?: PptTableCellBorder;
  top?: PptTableCellBorder;
  bottom?: PptTableCellBorder;
}

export interface PptTableCellBorder {
  color?: string;
  opacity?: number;
  width?: number;
  dash?: string;
  hidden?: boolean;
}

export interface PptTableCellMerge {
  row: number;
  column: number;
  rowSpan?: number;
  columnSpan?: number;
  hMerge?: boolean;
  vMerge?: boolean;
}

export interface PptShapeModel {
  id: string;
  name?: string;
  kind: PptShapeKind;
  geometry: PptShapeGeometry;
  style?: PptShapeStyle;
  text?: PptTextBodyModel;
  image?: PptImageModel;
  chart?: PptChartModel;
  media?: PptShapeMediaModel;
  oleObject?: PptOleObjectModel;
  smartArt?: PptSmartArtModel;
  table?: PptTableModel;
  placeholder?: PptPlaceholderModel;
  hyperlink?: PptHyperlinkModel;
  children?: PptShapeModel[];
  raw?: {
    slidePath?: string;
    relationshipId?: string;
    hostType?: string;
    presetGeometry?: string;
    presetAdjustments?: Record<string, number>;
    customGeometry?: PptCustomGeometryModel;
    threeD?: PptShape3dModel;
    groupTransform?: PptGroupTransformModel;
    useBackgroundFill?: boolean;
  };
}

export interface PptGroupTransformModel {
  childLeft?: number;
  childTop?: number;
  childWidth?: number;
  childHeight?: number;
}

export interface PptSlideModel {
  id: string;
  index: number;
  name?: string;
  title?: string;
  relationshipId?: string;
  path?: string;
  layoutId?: string;
  layout?: PptSlideLayoutModel;
  notes?: PptSlideNotesModel;
  comments?: PptCommentModel[];
  transition?: PptSlideTransitionModel;
  timing?: PptSlideTimingModel;
  background?: PptSlideBackgroundModel;
  shapes: PptShapeModel[];
}

export interface PptSlideLayoutModel {
  relationshipId?: string;
  path?: string;
  name?: string;
  background?: PptSlideBackgroundModel;
  shapes?: PptShapeModel[];
  master?: PptSlideMasterModel;
}

export interface PptSlideMasterModel {
  relationshipId?: string;
  path?: string;
  name?: string;
  background?: PptSlideBackgroundModel;
  shapes?: PptShapeModel[];
  textStyles?: {
    title?: PptTextStyle;
    body?: PptTextStyle;
    other?: PptTextStyle;
  };
  theme?: PptThemeModel;
}

export interface PptSlideNotesModel {
  text: string;
}

export interface PptCommentModel {
  id?: string;
  authorId?: string;
  author?: string;
  initials?: string;
  date?: string;
  text: string;
  x?: number;
  y?: number;
}

export interface PptCommentAddOptions {
  text: string;
  author?: string;
  initials?: string;
  date?: string;
  x?: number;
  y?: number;
}

export interface PptSlideTransitionModel {
  type?: string;
  speed?: string;
  duration?: number;
  advanceOnClick?: boolean;
  advanceAfterTime?: number;
  attributes?: Record<string, string>;
}

export interface PptSlideTimingModel {
  hasTiming: boolean;
  timeNodeCount: number;
  mainSequenceCount: number;
  interactiveSequenceCount: number;
  rootDuration?: string;
  rootNodeType?: string;
}

export interface PptMediaModel {
  path: string;
  contentType: string;
  byteLength: number;
}

export interface PptExportedImage extends PptMediaModel {
  bytes: Uint8Array;
}

export interface PptFontModel {
  family: string;
  path: string;
  contentType: string;
  byteLength: number;
  weight?: number;
  style?: 'normal' | 'italic';
}

export interface PptExportedFont extends PptFontModel {
  bytes: Uint8Array;
}

export interface PptPresentationModel {
  id?: string;
  title?: string;
  slideSize?: PptSlideSize;
  slides: PptSlideModel[];
  media?: PptMediaModel[];
  fonts?: PptFontModel[];
  themes?: PptThemeModel[];
}

export interface PptPresentationProperties {
  presentationId?: string;
  slideCount: number;
  shapeCount: number;
  textShapeCount: number;
  tableCount: number;
  pictureCount: number;
  mediaCount: number;
  slideSize?: PptSlideSize;
}

export interface PptSlideProperties {
  id: string;
  index: number;
  name?: string;
  title?: string;
  shapeCount: number;
  textShapeCount: number;
  tableCount: number;
  pictureCount: number;
}

export interface PptTextWriteOptions {
  preserveFirstRunStyle?: boolean;
}

export interface PptTextReplaceOptions extends PptTextWriteOptions {
  caseSensitive?: boolean;
}

export interface PptTextReplaceResult {
  count: number;
}

export interface PptFindTextOptions {
  caseSensitive?: boolean;
}

export interface PptTextMatch {
  slideId: string;
  slideIndex: number;
  shapeId: string;
  shapeName?: string;
  source: 'shapeText' | 'tableCell';
  text: string;
  start: number;
  length: number;
  row?: number;
  column?: number;
}

export type PptImageSource =
  | Uint8Array
  | ArrayBuffer
  | { bytes: Uint8Array | ArrayBuffer; mimeType?: string; filename?: string }
  | { base64: string; mimeType?: string; filename?: string };

export interface PptImageInfo {
  relationshipId?: string;
  targetPath?: string;
  contentType?: string;
  filename?: string;
  byteLength?: number;
}

export interface PptImageReplaceOptions {
  mimeType?: string;
  filename?: string;
}

export type PresentationRuntimeMode = 'auto' | 'inline' | 'worker';

export interface PresentationLoadOptions {
  presentationId?: string;
  mode?: PresentationRuntimeMode;
  onModeResolved?: (
    decision: { mode: 'inline' | 'worker'; source: string },
    presentationId: string,
  ) => void;
}
