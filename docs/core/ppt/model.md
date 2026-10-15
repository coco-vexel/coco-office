# PPT Model And Types

Use this document when a task mentions model snapshots, JSON fixtures,
properties snapshots, media records, or mutation payloads.

## Snapshot Boundary

```ts
const model = await presentation.toModel();
```

`toModel()` returns a defensive snapshot. Mutating the returned object does not
mutate the live `Presentation`.

## Skill Summary

Use model snapshots for inspection, selection, reporting, fixtures, and
writeback planning. Use live-object APIs for edits:

```ts
const model = await presentation.toModel();
const firstChart = model.slides
  .flatMap((slide) => slide.shapes.map((shape) => ({ slide, shape })))
  .find(({ shape }) => shape.kind === 'chart');

if (firstChart) {
  const shape = presentation.slides
    .getItem(firstChart.slide.id)
    .shapes.getItem(firstChart.shape.id);
  console.log(await shape.chart.getInfo());
}
```

Do not change `model.slides[0].shapes[0].text` and expect the deck to change.
Use `shape.text.setText(...)`, `shape.table.setCellText(...)`, or another
live-object method.

## Snapshot Usage Matrix

| Need                              | Use                                | Avoid                                     |
| --------------------------------- | ---------------------------------- | ----------------------------------------- |
| Count slides/shapes/tables/images | `presentation.properties()`        | Full `toModel()` when counts are enough.  |
| Select by kind/name/text metadata | `toModel()` then live handles      | Editing the snapshot object.              |
| Inspect groups/advanced objects   | `PptShapeModel.children`, metadata | Raw OOXML or unsupported edit methods.    |
| Export package bytes              | `presentation.toZip()`             | Reconstructing packages from snapshots.   |
| Create a synthetic deck in tests  | `Presentation.fromModel(model)`    | Depending on source-only raw fields.      |
| Runtime writeback planning        | `before/after` model snapshots     | Calling planner from user-operation code. |

## Core Model

Top-level:

```ts
interface PptPresentationModel {
  id?: string;
  title?: string;
  slideSize?: PptSlideSize;
  slides: PptSlideModel[];
  media?: PptMediaModel[];
  fonts?: PptFontModel[];
  themes?: PptThemeModel[];
}
```

Geometry:

```ts
interface PptShapeGeometry {
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
}
```

Slides:

```ts
interface PptSlideModel {
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
  shapes: PptShapeModel[];
}
```

```ts
interface PptSlideLayoutModel {
  relationshipId?: string;
  path?: string;
  name?: string;
  background?: PptSlideBackgroundModel;
  shapes?: PptShapeModel[];
  master?: PptSlideMasterModel;
}

interface PptSlideMasterModel {
  relationshipId?: string;
  path?: string;
  name?: string;
  background?: PptSlideBackgroundModel;
  shapes?: PptShapeModel[];
  theme?: PptThemeModel;
}

interface PptThemeModel {
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

interface PptSlideNotesModel {
  text: string;
}

interface PptCommentModel {
  id?: string;
  authorId?: string;
  author?: string;
  initials?: string;
  date?: string;
  text: string;
  x?: number;
  y?: number;
}

interface PptCommentAddOptions {
  text: string;
  author?: string;
  initials?: string;
  date?: string;
  x?: number;
  y?: number;
}

interface PptSlideTransitionModel {
  type?: string;
  speed?: string;
  duration?: number;
  advanceOnClick?: boolean;
  advanceAfterTime?: number;
  attributes?: Record<string, string>;
}

interface PptSlideTimingModel {
  hasTiming: boolean;
  timeNodeCount: number;
  mainSequenceCount: number;
  interactiveSequenceCount: number;
  rootDuration?: string;
  rootNodeType?: string;
}
```

Layout, master, theme, transition, and timing are read-only snapshots parsed
from presentation, master, and slide relationships. Notes and comments are
relationship-backed snapshots that can also be written through slide APIs.

## Load And Image Options

Use `PresentationLoadOptions` with `Presentation.fromZip`,
`Presentation.fromModel`, or `Presentation.createBlank` to choose a stable ID
and runtime mode. `onModeResolved` reports whether automatic selection chose the
inline or worker runtime.

```ts
type PresentationRuntimeMode = 'auto' | 'inline' | 'worker';

interface PresentationLoadOptions {
  presentationId?: string;
  mode?: PresentationRuntimeMode;
  onModeResolved?: (
    decision: { mode: 'inline' | 'worker'; source: string },
    presentationId: string,
  ) => void;
}
```

When replacing picture content, pass `PptImageReplaceOptions` to preserve an
explicit MIME type or filename when the source bytes do not provide them.

```ts
interface PptImageReplaceOptions {
  mimeType?: string;
  filename?: string;
}
```

Shapes:

```ts
interface PptShapeModel {
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
    groupTransform?: PptGroupTransformModel;
  };
}
```

`children` is populated for `kind: 'group'` and preserves direct group-child
order. Use `shape.group.items()` in live-object code instead of walking model
snapshots when subsequent operations are needed.

`PptShapeKind`:

```ts
type PptShapeKind =
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
```

## Styles

```ts
interface PptShapeStyle {
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

interface PptSlideBackgroundModel {
  fillColor?: string;
  fillGradient?: PptGradientFillModel;
}

interface PptGradientFillModel {
  colors: string[];
  angle?: number;
  positions?: number[];
  opacities?: number[];
}

interface PptShadowModel {
  color?: string;
  opacity?: number;
  blur?: number;
  distance?: number;
  direction?: number;
}

interface PptTextStyle {
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  textFill?: 'none';
  textOutlineColor?: string;
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

type PptParagraphAlignment = 'left' | 'center' | 'right' | 'justify' | 'distributed';

interface PptHyperlinkModel {
  relationshipId?: string;
  target?: string;
  targetMode?: string;
  action?: string;
  tooltip?: string;
}

interface PptChartModel {
  relationshipId?: string;
  targetPath?: string;
  type?: string;
  title?: string;
  seriesNames?: string[];
  data?: PptChartData;
  plotLayout?: PptChartLayoutModel;
  axes?: PptChartAxesModel;
}

interface PptChartData {
  title?: string;
  series: PptChartSeriesData[];
}

interface PptChartSeriesData {
  name?: string;
  type?: string;
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
  smooth?: boolean;
}

interface PptChartLayoutModel {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}

interface PptChartAxesModel {
  category?: PptChartAxisModel;
  value?: PptChartAxisModel;
}

interface PptChartAxisModel {
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

interface PptChartMarkerModel {
  symbol?: string;
  size?: number;
}

interface PptShapeMediaModel {
  relationshipId?: string;
  targetPath?: string;
  mediaType?: 'audio' | 'video' | 'media';
  contentType?: string;
  filename?: string;
}

interface PptOleObjectModel {
  relationshipId?: string;
  targetPath?: string;
  progId?: string;
  name?: string;
}

interface PptSmartArtModel {
  relationshipIds?: Record<string, string>;
  targetPaths?: Record<string, string>;
}

interface PptPlaceholderModel {
  type?: string;
  index?: string;
  size?: string;
  orientation?: string;
}

interface PptTextRun {
  text: string;
  style?: PptTextStyle;
  hyperlink?: PptHyperlinkModel;
}

interface PptTextRunWrite {
  text: string;
  style?: Partial<PptTextStyle>;
}

interface PptTextParagraph {
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

interface PptTextParagraphWrite {
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

interface PptTextBodyModel {
  text: string;
  style?: PptTextStyle;
  body?: PptTextBodyProperties;
  paragraphs?: PptTextParagraph[];
}

interface PptTextBodyProperties {
  anchor?: 'top' | 'middle' | 'bottom';
  autoFit?: 'none' | 'shape' | 'normal';
  wrap?: 'none' | 'square';
  insetLeft?: number;
  insetTop?: number;
  insetRight?: number;
  insetBottom?: number;
}

interface PptTableCellStyle {
  fillColor?: string;
  fillGradient?: PptGradientFillModel;
  borders?: PptTableCellBorders;
  body?: PptTextBodyProperties;
  textStyle?: PptTextStyle;
}

interface PptTableCellBorders {
  left?: PptTableCellBorder;
  right?: PptTableCellBorder;
  top?: PptTableCellBorder;
  bottom?: PptTableCellBorder;
}

interface PptTableCellBorder {
  color?: string;
  opacity?: number;
  width?: number;
  dash?: string;
  hidden?: boolean;
}
```

Colors are normalized as `#RRGGBB`.
`PptTextBodyModel.text` remains the aggregated text value. `paragraphs` is a
structure snapshot for inspection and testing. `PptTextParagraphWrite` and
`PptTextRunWrite` are accepted by structured text writers and intentionally
allow partial style objects.
`PptHyperlinkModel` is parsed from `a:hlinkClick` and slide relationships.
`PptHyperlinkWriteOptions` writes relationship-backed hyperlinks for shape-level
click targets and flattened text runs.

## Tables

```ts
interface PptTableModel {
  values: string[][];
  rowCount: number;
  columnCount: number;
  columnWidths?: number[];
  rowHeights?: number[];
  cellStyles?: PptTableCellStyle[][];
  mergedCells?: PptTableCellMerge[];
}

interface PptTableCellMerge {
  row: number;
  column: number;
  rowSpan?: number;
  columnSpan?: number;
  hMerge?: boolean;
  vMerge?: boolean;
}
```

Merged-cell metadata comes from OOXML `rowSpan`, `gridSpan`, `hMerge`, and
`vMerge`. Structural table edits reject merged-cell tables, while owner-cell
text/style edits are allowed and continuation cells are protected.

## Media And Images

```ts
interface PptMediaModel {
  path: string;
  contentType: string;
  byteLength: number;
}

interface PptExportedImage extends PptMediaModel {
  bytes: Uint8Array;
}

interface PptFontModel {
  family: string;
  path: string;
  contentType: string;
  byteLength: number;
  weight?: number;
  style?: 'normal' | 'italic';
}

interface PptExportedFont extends PptFontModel {
  bytes: Uint8Array;
}
```

`PptMediaModel` is metadata for snapshots. Use `presentation.exportImages()` to
retrieve package media bytes from `ppt/media/*`, and use `shape.image.bytes()`
when the task specifically needs bytes behind a replaceable picture shape. Use
`presentation.exportFonts()` to retrieve embedded font bytes.

## Advanced Shape Metadata

```ts
interface PptImageCropModel {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}

interface PptImageStyle {
  crop?: PptImageCropModel;
  fillRect?: PptImageCropModel;
  opacity?: number;
}

interface PptCustomGeometryModel {
  paths: PptCustomGeometryPathModel[];
}

interface PptCustomGeometryPathModel {
  width?: number;
  height?: number;
  commands: PptCustomGeometryCommand[];
}

type PptCustomGeometryCommand =
  | { kind: 'moveTo'; x: number; y: number }
  | { kind: 'lineTo'; x: number; y: number }
  | { kind: 'cubicBezierTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { kind: 'arcTo'; wR: number; hR: number; stAng: number; swAng: number }
  | { kind: 'close' };

interface PptGroupTransformModel {
  childLeft?: number;
  childTop?: number;
  childWidth?: number;
  childHeight?: number;
}
```

`PptImageCropModel`, `PptCustomGeometryModel`, `PptCustomGeometryPathModel`,
`PptCustomGeometryCommand`, and `PptGroupTransformModel` are snapshot metadata
for inspection and fidelity. SDK operation code should prefer supported
live-object APIs instead of editing these raw geometry structures directly.

## Properties Snapshots

`PptPresentationProperties`:

- `presentationId`
- `slideCount`
- `shapeCount`
- `textShapeCount`
- `tableCount`
- `pictureCount`
- `mediaCount`
- `slideSize`

`PptSlideProperties`:

- `id`
- `index`
- `name`
- `title`
- `shapeCount`
- `textShapeCount`
- `tableCount`
- `pictureCount`

## Search Types

`PptTextMatch` identifies either a shape text match or a table cell match. Use
`source` to select the correct edit API.

```ts
interface PptTextMatch {
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
```

For `source: 'tableCell'`, `row` and `column` are zero-based cell indexes.

## SDK Operation Rules

`PptShape3dModel` describes preserved shape 3D material and extrusion height.
`PptChartDataLabelLayoutModel` describes a chart label's zero-based `index` and
optional `x` / `y` manual layout coordinates. These inspection models do not imply
that every corresponding 3D or chart-label property has a write API.

1. Do not mutate snapshots returned by `toModel()` to edit a deck.
2. Use live-object APIs for edits.
3. Use model snapshots for inspection, reporting, and tests.
4. Do not depend on `raw` fields in ordinary SDK operation code.
5. Use `properties()` when counts are enough.
6. Use `shape.kind` and optional model metadata to choose the correct live-object domain.
7. Treat layout, master, theme, transition, timing, media, OLE, and SmartArt model fields as read-only unless a domain document explicitly says an edit API exists.
8. When operation code needs to edit an object discovered in a snapshot, re-open it through `presentation.slides.getItem(...).shapes.getItem(...)` before writing.
