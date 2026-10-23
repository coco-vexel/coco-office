import type {
  CommentInfo,
  ParagraphFormat,
  ParagraphInfo,
  RevisionInfo,
  StoryId,
  TableInfo,
} from '@coco-vexel/coco-office/word-view';

export type EditorSegmentKind = 'text' | 'revision' | 'field' | 'placeholder' | 'image' | 'drawing';
export type EditorSpanKind = 'comment' | 'revision';
export type EditorContentRevisionKind = Extract<RevisionInfo['kind'], 'insert' | 'delete'>;

export interface EditorTextSegment {
  kind: 'text';
  from: number;
  to: number;
  text: string;
  format?: EditorRunFormat;
  hyperlink?: EditorHyperlink;
}

export interface EditorHyperlink {
  target?: string;
  anchor?: string;
  tooltip?: string;
}

export interface EditorRevisionSegment {
  kind: 'revision';
  from: number;
  to: number;
  text: string;
  format?: EditorRunFormat;
  revisionId: string;
  revisionKind: EditorContentRevisionKind;
  author: string;
  date: string;
}

export interface EditorFieldSegment {
  kind: 'field';
  from: number;
  to: number;
  fieldKind: 'pageNumber' | 'totalPages' | 'sectionPages';
  instruction: string;
  cachedText?: string;
  format?: EditorRunFormat;
}

export interface EditorPlaceholderSegment {
  kind: 'placeholder';
  from: number;
  to: number;
  placeholderKind: 'drawing' | 'object' | 'pageBreak' | 'unknown';
  sourceTag: string;
}

export interface EditorImageSegment {
  kind: 'image';
  from: number;
  to: number;
  drawingId: string;
  src: string;
  contentType: string;
  width?: number;
  height?: number;
  anchoringMode?: 'inline' | 'anchor';
  wrap?: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behind' | 'inFront';
  position?: {
    horizontal?: {
      relativeFrom?: string;
      posOffset?: number;
      align?: string;
    };
    vertical?: {
      relativeFrom?: string;
      posOffset?: number;
      align?: string;
    };
  };
}

export interface EditorDrawingSegment {
  kind: 'drawing';
  from: number;
  to: number;
  drawingId: string;
  drawingKind: 'chart' | 'smartArt' | 'shape' | 'textBox' | 'group' | 'canvas' | 'other';
  text?: string;
  width?: number;
  height?: number;
  anchoringMode?: 'inline' | 'anchor';
  wrap?: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behind' | 'inFront';
  position?: EditorImageSegment['position'];
  hasBorder?: boolean;
  hasFill?: boolean;
  isLineOrConnector?: boolean;
  innerTextAlign?: string;
  paragraphs?: EditorParagraphView[];
}

export type EditorContentSegment =
  | EditorTextSegment
  | EditorRevisionSegment
  | EditorFieldSegment
  | EditorPlaceholderSegment
  | EditorImageSegment
  | EditorDrawingSegment;

export interface EditorTextSpan {
  kind: EditorSpanKind;
  id: string;
  paraId: string;
  from: number;
  to: number;
  text?: string;
}

export interface EditorCommentSpan extends EditorTextSpan {
  kind: 'comment';
  commentId: number;
  rangeKind: CommentInfo['rangeKind'];
  author: string;
  date: string;
  body: string;
}

export interface EditorRevisionSpan extends EditorTextSpan {
  kind: 'revision';
  revisionId: string;
  revisionKind: EditorContentRevisionKind;
  author: string;
  date: string;
}

export type EditorParagraphSpan = EditorCommentSpan | EditorRevisionSpan;

export interface EditorRunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  doubleStrike?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  characterSpacing?: number;
  verticalAlign?: 'baseline' | 'superscript' | 'subscript';
  highlight?: string;
  shading?: EditorShading;
}

export interface EditorParagraphFormat extends ParagraphFormat {
  fontFamily?: string;
  tabStops?: EditorTabStop[];
  numbering?: EditorNumberingFormat;
  characterSpacing?: number;
  highlight?: string;
  shading?: EditorShading;
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  snapToGrid?: boolean;
  borders?: EditorParagraphBorders;
  borderBottom?: EditorParagraphBorder;
}

export interface EditorTabStop {
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar';
  position: number;
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface EditorNumberingFormat {
  numId: string;
  ilvl: number;
  numFmt?: string;
  lvlText?: string;
  start?: number;
  label?: string;
}

export interface EditorShading {
  fill?: string;
  color?: string;
  val?: string;
}

export interface EditorParagraphView {
  type: 'paragraph';
  paraId: string;
  story: StoryId;
  text: string;
  styleId?: string;
  format?: EditorParagraphFormat;
  table?: ParagraphInfo['table'];
  flags: {
    hasComments: boolean;
    hasRevisions: boolean;
    hasBookmarks: boolean;
    hasMath: boolean;
  };
  segments: EditorContentSegment[];
  spans: EditorParagraphSpan[];
}

export interface EditorTableCellView {
  text: string;
  gridSpan: number;
  rowSpan?: number;
  vMerge: 'restart' | 'continue' | null;
  width?: number;
  margin?: EditorBoxEdges;
  borders?: EditorTableBorders;
  shading?: {
    fill?: string;
    color?: string;
    val?: string;
  };
  verticalAlignment?: 'top' | 'center' | 'bottom' | 'both';
  noWrap?: boolean;
  blocks: EditorBlockView[];
  paragraphs: EditorParagraphView[];
}

export interface EditorTableRowView {
  height?: {
    val: number;
    hRule: 'auto' | 'atLeast' | 'exact';
  };
  isHeader?: boolean;
  cantSplit?: boolean;
  cells: EditorTableCellView[];
}

export interface EditorTableView {
  type: 'table';
  tableId: string;
  tableStyleId?: string;
  columnWidths: number[];
  layout?: EditorTableLayout;
  tableLook?: EditorTableLook;
  borders?: EditorTableBorders;
  rows: EditorTableRowView[];
}

export interface EditorTableLayout {
  width?: number;
  widthType?: 'auto' | 'dxa' | 'pct';
  indent?: number;
  cellSpacing?: number;
  cellSpacingType?: 'auto' | 'dxa' | 'pct';
  alignment?: 'left' | 'center' | 'right';
}

export interface EditorTableLook {
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  noHorizontalBand?: boolean;
  noVerticalBand?: boolean;
  rowBandSize?: number;
  columnBandSize?: number;
}

export type EditorBlockView = EditorParagraphView | EditorTableView;

export interface EditorBoxEdges {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

export interface EditorTableBorder {
  val?: string;
  color?: string;
  size?: number;
}

export interface EditorParagraphBorder extends EditorTableBorder {
  space?: number;
}

export interface EditorParagraphBorders {
  top?: EditorParagraphBorder;
  right?: EditorParagraphBorder;
  bottom?: EditorParagraphBorder;
  left?: EditorParagraphBorder;
  between?: EditorParagraphBorder;
  bar?: EditorParagraphBorder;
}

export interface EditorTableBorders {
  top?: EditorTableBorder;
  right?: EditorTableBorder;
  bottom?: EditorTableBorder;
  left?: EditorTableBorder;
  insideH?: EditorTableBorder;
  insideV?: EditorTableBorder;
}

export interface EditorDocumentDefaults {
  run?: EditorRunFormat;
  paragraph?: EditorParagraphFormat;
}

export interface EditorPageLayout {
  width: number;
  height: number;
  headerDistance?: number;
  footerDistance?: number;
  docGrid?: {
    type?: 'default' | 'lines' | 'linesAndChars' | 'snapToChars';
    linePitch?: number;
    charSpace?: number;
  };
  margin: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
  orientation?: 'portrait' | 'landscape';
}

export interface EditorHeaderFooterRef {
  type: 'default' | 'first' | 'even';
  partName: string;
}

export interface EditorSectionView {
  sectionIndex: number;
  firstParaId: string;
  lastParaId: string;
  type?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  pageLayout?: EditorPageLayout;
  pageNumberStart?: number;
  headerRefs: EditorHeaderFooterRef[];
  footerRefs: EditorHeaderFooterRef[];
  titlePage: boolean;
}

export interface EditorHeaderFooterStoryView {
  kind: 'header' | 'footer';
  partName: string;
  blocks: EditorBlockView[];
  paragraphs: EditorParagraphView[];
  tables: EditorTableView[];
}

export interface EditorDocumentView {
  kind: 'word';
  defaults?: EditorDocumentDefaults;
  pageLayout?: EditorPageLayout;
  sections: EditorSectionView[];
  headers: EditorHeaderFooterStoryView[];
  footers: EditorHeaderFooterStoryView[];
  blocks: EditorBlockView[];
  paragraphs: EditorParagraphView[];
  tables: EditorTableView[];
  detachedComments: CommentInfo[];
  detachedRevisions: RevisionInfo[];
}
