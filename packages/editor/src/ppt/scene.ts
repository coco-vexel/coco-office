import type {
  PptPresentationModel,
  PptChartModel,
  PptCustomGeometryModel,
  PptGradientFillModel,
  PptImageCropModel,
  PptShapeGeometry,
  PptShapeKind,
  PptShapeModel,
  PptShapeStyle,
  PptTextStyle,
  PptSlideBackgroundModel,
  PptSlideModel,
  PptSlideSize,
  PptTableModel,
  PptTextBodyProperties,
  PptTextParagraph,
} from '@coco-vexel/coco-office/ppt';
import { resolvePptEffectivePlaceholderProperties } from '@coco-vexel/coco-office/ppt';

export type PptSceneNodeKind = PptShapeKind;
type PptShapeStyleWithLineGradient = PptShapeStyle & {
  lineGradient?: PptGradientFillModel;
  lineCap?: string;
  lineDash?: string;
};

export interface PptSceneRect {
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
}

export interface PptSceneNode {
  id: string;
  name?: string;
  kind: PptSceneNodeKind;
  bounds: PptSceneRect;
  text?: string;
  textStyle?: PptTextStyle;
  textParagraphs?: PptTextParagraph[];
  textBody?: PptTextBodyProperties;
  imagePath?: string;
  imageCrop?: PptImageCropModel;
  imageFillRect?: PptImageCropModel;
  imageOpacity?: number;
  chart?: PptChartModel;
  table?: PptTableModel;
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
  shadow?: PptShapeStyle['shadow'];
  presetGeometry?: string;
  presetAdjustments?: Record<string, number>;
  customGeometry?: PptCustomGeometryModel;
  threeD?: NonNullable<PptShapeModel['raw']>['threeD'];
  usesBackgroundFill?: boolean;
  backgroundFill?: PptSlideBackgroundModel;
  locked: boolean;
  editable: boolean;
  capabilities: PptSceneNodeCapabilities;
  groupContext?: PptSceneGroupContext;
  source: PptShapeModel;
  children: PptSceneNode[];
}

export interface PptSceneGroupContext {
  parentId: string;
  parentName?: string;
  parentBounds: PptSceneRect;
  parentTransform: {
    childLeft: number;
    childTop: number;
    childWidth: number;
    childHeight: number;
  };
}

export interface PptSceneNodeCapabilities {
  readonly select: boolean;
  readonly editText: boolean;
  readonly editTextStyle: boolean;
  readonly editTableCells: boolean;
  readonly editGeometry: boolean;
  readonly editShapeStyle: boolean;
  readonly reason?: string;
  readonly reasons?: Partial<Record<PptSceneNodeCapability, string>>;
}

export interface PptSlideScene {
  id: string;
  index: number;
  name?: string;
  title?: string;
  size: PptSlideSize;
  background?: PptSlideBackgroundModel;
  nodes: PptSceneNode[];
  groups: PptSceneNode[];
  source: PptSlideModel;
}

export interface PptPresentationScene {
  size: PptSlideSize;
  slides: PptSlideScene[];
}

export type PptSceneNodeCapability =
  'editText' | 'editTextStyle' | 'editTableCells' | 'editGeometry' | 'editShapeStyle';

export const DEFAULT_PPT_SLIDE_SIZE: PptSlideSize = {
  width: 9144000,
  height: 5143500,
};

const editableShapeKinds = new Set<PptShapeKind>([
  'placeholder',
  'textBox',
  'geometricShape',
  'line',
]);
const pptSceneGroupContexts = new WeakMap<PptShapeModel, PptSceneGroupContext>();

export function buildPptPresentationScene(model: PptPresentationModel): PptPresentationScene {
  const size = model.slideSize ?? DEFAULT_PPT_SLIDE_SIZE;
  return {
    size,
    slides: model.slides.map((slide) => buildPptSlideScene(slide, size)),
  };
}

export function buildPptSlideScene(
  slide: PptSlideModel,
  fallbackSize: PptSlideSize,
): PptSlideScene {
  const layoutShapes = slide.layout?.shapes ?? [];
  const inheritedShapes = [
    ...(slide.layout?.master?.shapes ?? []).map((shape) =>
      namespaceInheritedShape(shape, 'master'),
    ),
    ...layoutShapes.map((shape) => namespaceInheritedShape(shape, 'layout')),
  ].filter(shouldRenderInheritedShape);
  const shapes = [
    ...inheritedShapes,
    ...slide.shapes.map((shape) => applyPlaceholderInheritance(shape, slide.layout)),
  ];
  const renderShapes = shapes.flatMap(flattenRenderableShape);
  const background =
    slide.background ?? slide.layout?.background ?? slide.layout?.master?.background;
  const backgroundFill = background ?? inferFullSlidePictureBackground(shapes, fallbackSize);
  return {
    id: slide.id,
    index: slide.index,
    name: slide.name,
    title: slide.title,
    size: fallbackSize,
    ...(background ? { background } : {}),
    nodes: renderShapes.map((shape, index) =>
      buildPptSceneNode(shape, { slideSize: fallbackSize, index, background, backgroundFill }),
    ),
    groups: shapes.flatMap((shape) =>
      collectGroupSceneNodes(shape, { background, backgroundFill }),
    ),
    source: slide,
  };
}

function inferFullSlidePictureBackground(
  shapes: PptShapeModel[],
  slideSize: PptSlideSize,
): PptSlideBackgroundModel | undefined {
  const picture = shapes.find((shape) => {
    const bounds = normalizePptGeometry(shape.geometry);
    return (
      shape.kind === 'picture' &&
      !!shape.image?.targetPath &&
      bounds.left <= slideSize.width * 0.02 &&
      bounds.top <= slideSize.height * 0.02 &&
      bounds.width >= slideSize.width * 0.96 &&
      bounds.height >= slideSize.height * 0.96
    );
  });
  return picture?.image?.targetPath ? { imagePath: picture.image.targetPath } : undefined;
}

function namespaceInheritedShape(shape: PptShapeModel, source: 'layout' | 'master'): PptShapeModel {
  return {
    ...shape,
    id: `${source}:${shape.id}`,
    children: shape.children?.map((child) => namespaceInheritedShape(child, source)),
  };
}

function shouldRenderInheritedShape(shape: PptShapeModel): boolean {
  if (shape.placeholder) return false;
  return (
    shape.kind === 'group' ||
    shape.kind === 'picture' ||
    shape.kind === 'geometricShape' ||
    shape.kind === 'line' ||
    shape.kind === 'textBox'
  );
}

function flattenRenderableShape(shape: PptShapeModel): PptShapeModel[] {
  if (shape.kind !== 'group') return [shape];
  return (shape.children ?? []).flatMap((child) =>
    flattenRenderableShape(transformGroupChild(shape, child)),
  );
}

function collectGroupSceneNodes(
  shape: PptShapeModel,
  options: { background?: PptSlideBackgroundModel; backgroundFill?: PptSlideBackgroundModel } = {},
): PptSceneNode[] {
  if (shape.kind !== 'group') return [];
  return [
    buildPptSceneNode(shape, {
      background: options.background,
      backgroundFill: options.backgroundFill,
    }),
    ...(shape.children ?? [])
      .map((child) => transformGroupChild(shape, child))
      .flatMap((child) => collectGroupSceneNodes(child, options)),
  ];
}

function transformGroupChild(group: PptShapeModel, child: PptShapeModel): PptShapeModel {
  const groupBounds = normalizePptGeometry(group.geometry);
  const groupTransform = group.raw?.groupTransform;
  const childBounds = normalizePptGeometry(child.geometry);
  const childLeft = groupTransform?.childLeft ?? groupBounds.left;
  const childTop = groupTransform?.childTop ?? groupBounds.top;
  const childWidth = groupTransform?.childWidth || groupBounds.width || 1;
  const childHeight = groupTransform?.childHeight || groupBounds.height || 1;
  const scaleX = groupBounds.width / childWidth;
  const scaleY = groupBounds.height / childHeight;
  const scaledLeft = (childBounds.left - childLeft) * scaleX;
  const scaledTop = (childBounds.top - childTop) * scaleY;
  const scaledWidth = childBounds.width * scaleX;
  const scaledHeight = childBounds.height * scaleY;
  const unrotatedLeft = groupBounds.flipHorizontal
    ? groupBounds.left + groupBounds.width - scaledLeft - scaledWidth
    : groupBounds.left + scaledLeft;
  const unrotatedTop = groupBounds.flipVertical
    ? groupBounds.top + groupBounds.height - scaledTop - scaledHeight
    : groupBounds.top + scaledTop;
  const rotatedBounds = rotateGroupChildBounds(
    {
      left: unrotatedLeft,
      top: unrotatedTop,
      width: scaledWidth,
      height: scaledHeight,
    },
    groupBounds,
  );
  const rotation = normalizeRotation((groupBounds.rotation ?? 0) + (childBounds.rotation ?? 0));
  const transformed: PptShapeModel = {
    ...child,
    id: `${group.id}:${child.id}`,
    geometry: {
      ...child.geometry,
      left: rotatedBounds.left,
      top: rotatedBounds.top,
      width: scaledWidth,
      height: scaledHeight,
      rotation,
      flipHorizontal: Boolean(groupBounds.flipHorizontal) !== Boolean(childBounds.flipHorizontal),
      flipVertical: Boolean(groupBounds.flipVertical) !== Boolean(childBounds.flipVertical),
    },
    children: child.children,
  };
  pptSceneGroupContexts.set(transformed, {
    parentId: group.id,
    parentName: group.name,
    parentBounds: groupBounds,
    parentTransform: {
      childLeft,
      childTop,
      childWidth,
      childHeight,
    },
  });
  return transformed;
}

function rotateGroupChildBounds(child: PptSceneRect, group: PptSceneRect): PptSceneRect {
  if (!group.rotation) return child;
  const angle = (group.rotation * Math.PI) / 180;
  const groupCenterX = group.left + group.width / 2;
  const groupCenterY = group.top + group.height / 2;
  const childCenterX = child.left + child.width / 2;
  const childCenterY = child.top + child.height / 2;
  const dx = childCenterX - groupCenterX;
  const dy = childCenterY - groupCenterY;
  const rotatedCenterX = groupCenterX + dx * Math.cos(angle) - dy * Math.sin(angle);
  const rotatedCenterY = groupCenterY + dx * Math.sin(angle) + dy * Math.cos(angle);
  return {
    ...child,
    left: rotatedCenterX - child.width / 2,
    top: rotatedCenterY - child.height / 2,
  };
}

function normalizeRotation(rotation: number): number | undefined {
  const normalized = ((rotation % 360) + 360) % 360;
  return normalized === 0 ? undefined : normalized;
}

function applyPlaceholderInheritance(
  shape: PptShapeModel,
  layout: PptSlideModel['layout'],
): PptShapeModel {
  if (!shape.placeholder) return shape;
  const effective = resolvePptEffectivePlaceholderProperties(shape, layout);
  return {
    ...shape,
    placeholder: effective.placeholder,
    geometry: effective.geometry,
    style: effective.style,
    ...(shape.text
      ? {
          text: {
            ...shape.text,
            style: effective.textStyle,
            body: effective.textBody,
          },
        }
      : {}),
  };
}

export function buildPptSceneNode(
  shape: PptShapeModel,
  options: {
    slideSize?: PptSlideSize;
    index?: number;
    background?: PptSlideBackgroundModel;
    backgroundFill?: PptSlideBackgroundModel;
  } = {},
): PptSceneNode {
  const capabilities = pptSceneNodeCapabilities(shape);
  const editable =
    capabilities.editText ||
    capabilities.editTextStyle ||
    capabilities.editTableCells ||
    capabilities.editGeometry ||
    capabilities.editShapeStyle;
  const shapeStyle = shape.style as PptShapeStyleWithLineGradient | undefined;
  const usesBackgroundFill = shape.raw?.useBackgroundFill === true;
  const fillColor = usesBackgroundFill
    ? undefined
    : (shape.style?.fillColor ?? inferPptSceneFillColor(shape, options));
  const textStyle = inferPptSceneTextStyle(shape.text?.style, {
    fillColor,
    background: options.background,
  });
  const bounds = normalizePptGeometry(shape.geometry);
  return {
    id: shape.id,
    name: shape.name,
    kind: shape.kind,
    bounds: tableVisualBounds(shape, bounds),
    text: shape.text?.text ?? tablePreviewText(shape),
    textStyle,
    textParagraphs: shape.text?.paragraphs,
    textBody: shape.text?.body,
    imagePath: shape.image?.targetPath,
    imageCrop: shape.image?.crop,
    imageFillRect: shape.image?.fillRect,
    imageOpacity: shape.image?.opacity,
    chart: shape.chart,
    table: shape.table,
    fillColor,
    fillGradient: usesBackgroundFill ? undefined : shapeStyle?.fillGradient,
    opacity: shapeStyle?.opacity,
    lineColor: shapeStyle?.lineColor,
    lineGradient: shapeStyle?.lineGradient,
    lineOpacity: shapeStyle?.lineOpacity,
    lineWidth: shapeStyle?.lineWidth,
    lineCap: shapeStyle?.lineCap,
    lineDash: shapeStyle?.lineDash,
    lineHeadEnd: shapeStyle?.lineHeadEnd,
    lineTailEnd: shapeStyle?.lineTailEnd,
    shadow: shapeStyle?.shadow,
    presetGeometry: shape.raw?.presetGeometry,
    presetAdjustments: shape.raw?.presetAdjustments,
    customGeometry: shape.raw?.customGeometry,
    threeD: shape.raw?.threeD,
    usesBackgroundFill,
    ...(usesBackgroundFill && options.backgroundFill
      ? { backgroundFill: options.backgroundFill }
      : {}),
    locked: !editable,
    editable,
    capabilities,
    groupContext: groupContextOf(shape),
    source: shape,
    children: (shape.children ?? []).map((child) =>
      buildPptSceneNode(child, {
        background: options.background,
        backgroundFill: options.backgroundFill,
      }),
    ),
  };
}

function tablePreviewText(shape: PptShapeModel): string | undefined {
  if (!shape.table) return undefined;
  return shape.table.values.map((row) => row.join('  ')).join('\n');
}

function tableVisualBounds(shape: PptShapeModel, bounds: PptSceneRect): PptSceneRect {
  if (!shape.table) return bounds;
  const width = sumPositiveNumbers(shape.table.columnWidths);
  const height = sumPositiveNumbers(shape.table.rowHeights);
  if (width === undefined && height === undefined) return bounds;
  return {
    ...bounds,
    ...(shouldUseTableGridExtent(bounds.width, width) ? { width } : {}),
    ...(shouldUseTableGridExtent(bounds.height, height) ? { height } : {}),
  };
}

function sumPositiveNumbers(values: number[] | undefined): number | undefined {
  if (!values?.length) return undefined;
  const sum = values.reduce((total, value) => total + value, 0);
  return Number.isFinite(sum) && sum > 0 ? sum : undefined;
}

function shouldUseTableGridExtent(frameExtent: number, gridExtent: number | undefined): boolean {
  if (gridExtent === undefined) return false;
  if (!Number.isFinite(frameExtent) || frameExtent <= 0) return true;
  const ratio = gridExtent / frameExtent;
  return ratio > 1.15;
}

function inferPptSceneFillColor(
  shape: PptShapeModel,
  options: { slideSize?: PptSlideSize; index?: number },
): string | undefined {
  const bounds = normalizePptGeometry(shape.geometry);
  const size = options.slideSize;
  if (
    options.index === 0 &&
    shape.kind === 'geometricShape' &&
    !shape.text?.text &&
    size &&
    bounds.left <= 0 &&
    bounds.top <= 0 &&
    bounds.width >= size.width * 0.95 &&
    bounds.height >= size.height * 0.95
  ) {
    return '#1D3F80';
  }
  return undefined;
}

function inferPptSceneTextStyle(
  style: PptTextStyle | undefined,
  context: { fillColor?: string; background?: PptSlideBackgroundModel },
): PptTextStyle | undefined {
  if (style?.color) return style;
  const referenceColor = context.fillColor ?? context.background?.fillColor;
  if (!referenceColor || !isDarkColor(referenceColor)) return style;
  return {
    ...style,
    color: '#FFFFFF',
  };
}

function isDarkColor(color: string): boolean {
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) return false;
  const value = match[1]!;
  const r = Number.parseInt(value.slice(0, 2), 16);
  const g = Number.parseInt(value.slice(2, 4), 16);
  const b = Number.parseInt(value.slice(4, 6), 16);
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 < 0.45;
}

export function isPptShapeEditable(shape: PptShapeModel): boolean {
  const capabilities = pptSceneNodeCapabilities(shape);
  return (
    capabilities.editText ||
    capabilities.editTextStyle ||
    capabilities.editTableCells ||
    capabilities.editGeometry ||
    capabilities.editShapeStyle
  );
}

export function pptSceneNodeCapabilities(shape: PptShapeModel): PptSceneNodeCapabilities {
  if (isInheritedPptShape(shape)) {
    return {
      select: true,
      editText: false,
      editTextStyle: false,
      editTableCells: false,
      editGeometry: false,
      editShapeStyle: false,
      reason: 'Inherited from layout/master',
      reasons: allCapabilityReasons('继承自母版/版式，当前只支持查看'),
    };
  }
  if (isProjectedGroupChild(shape)) {
    const editableTextShape = Boolean(shape.text) && editableShapeKinds.has(shape.kind);
    const editableTable = shape.kind === 'table' && Boolean(shape.table);
    const baseReason = '组内子对象暂不支持该项直接编辑';
    return {
      select: true,
      editText: editableTextShape,
      editTextStyle: editableTextShape,
      editTableCells: editableTable,
      editGeometry: true,
      editShapeStyle: false,
      reason: 'Flattened group child',
      reasons: disabledCapabilityReasons({
        editText: editableTextShape ? undefined : textCapabilityReason(shape),
        editTextStyle: editableTextShape ? undefined : textCapabilityReason(shape),
        editTableCells: editableTable ? undefined : '当前对象不是表格',
        editShapeStyle: baseReason,
      }),
    };
  }
  if (shape.kind === 'table') {
    return {
      select: true,
      editText: false,
      editTextStyle: false,
      editTableCells: true,
      editGeometry: true,
      editShapeStyle: true,
      reasons: disabledCapabilityReasons({
        editText: '表格文字请使用单元格编辑',
        editTextStyle: '表格文字样式当前不在高频工具栏开放',
      }),
    };
  }
  if (shape.kind === 'picture') {
    return {
      select: true,
      editText: false,
      editTextStyle: false,
      editTableCells: false,
      editGeometry: true,
      editShapeStyle: true,
      reasons: disabledCapabilityReasons({
        editText: '图片对象没有可编辑文字',
        editTextStyle: '图片对象没有可编辑文字样式',
        editTableCells: '当前对象不是表格',
      }),
    };
  }
  if (shape.kind === 'chart' || shape.kind === 'smartArt' || shape.kind === 'group') {
    const reason = limitedObjectTypeReason(shape.kind);
    return {
      select: true,
      editText: false,
      editTextStyle: false,
      editTableCells: false,
      editGeometry: true,
      editShapeStyle: false,
      reason: 'Limited editor support for this object type',
      reasons: disabledCapabilityReasons({
        editText: reason.text,
        editTextStyle: reason.textStyle,
        editTableCells: '当前对象不是表格',
        editShapeStyle: reason.shapeStyle,
      }),
    };
  }
  const editableTextShape = Boolean(shape.text) && editableShapeKinds.has(shape.kind);
  if (editableShapeKinds.has(shape.kind)) {
    return {
      select: true,
      editText: editableTextShape,
      editTextStyle: editableTextShape,
      editTableCells: false,
      editGeometry: true,
      editShapeStyle: true,
      reasons: disabledCapabilityReasons({
        editText: editableTextShape ? undefined : textCapabilityReason(shape),
        editTextStyle: editableTextShape ? undefined : textCapabilityReason(shape),
        editTableCells: '当前对象不是表格',
      }),
    };
  }
  return {
    select: true,
    editText: false,
    editTextStyle: false,
    editTableCells: false,
    editGeometry: false,
    editShapeStyle: false,
    reason: 'Unsupported object type',
    reasons: allCapabilityReasons('当前对象类型暂不支持常规编辑'),
  };
}

function disabledCapabilityReasons(
  reasons: Partial<Record<PptSceneNodeCapability, string | undefined>>,
): Partial<Record<PptSceneNodeCapability, string>> {
  return Object.fromEntries(
    Object.entries(reasons).filter((entry): entry is [PptSceneNodeCapability, string] =>
      Boolean(entry[1]),
    ),
  ) as Partial<Record<PptSceneNodeCapability, string>>;
}

function allCapabilityReasons(reason: string): Partial<Record<PptSceneNodeCapability, string>> {
  return {
    editText: reason,
    editTextStyle: reason,
    editTableCells: reason,
    editGeometry: reason,
    editShapeStyle: reason,
  };
}

function textCapabilityReason(shape: PptShapeModel): string {
  if (!shape.text) return '当前对象没有可编辑文字内容';
  if (!editableShapeKinds.has(shape.kind)) return '当前对象类型暂不支持文字编辑';
  return '当前对象文字暂不支持编辑';
}

function limitedObjectTypeReason(kind: PptShapeKind): {
  text: string;
  textStyle: string;
  shapeStyle: string;
} {
  if (kind === 'chart') {
    return {
      text: '图表文字绑定在图表数据中，当前仅开放整体位置尺寸',
      textStyle: '图表内部文字样式暂未开放',
      shapeStyle: '图表形状样式暂未开放',
    };
  }
  if (kind === 'smartArt') {
    return {
      text: 'SmartArt 内部文本结构暂未开放',
      textStyle: 'SmartArt 内部文字样式暂未开放',
      shapeStyle: 'SmartArt 样式暂未开放',
    };
  }
  return {
    text: '组合对象请双击或点选组内子对象编辑文字',
    textStyle: '组合对象请选中组内文本子对象编辑文字样式',
    shapeStyle: '组合对象暂不支持整体形状样式编辑',
  };
}

function isInheritedPptShape(shape: PptShapeModel): boolean {
  return shape.id.startsWith('layout:') || shape.id.startsWith('master:');
}

function isProjectedGroupChild(shape: PptShapeModel): boolean {
  return shape.id.includes(':');
}

export function pptSceneGeometryToSourceGeometry(
  node: PptSceneNode,
  geometry: Partial<PptShapeGeometry>,
): Partial<PptShapeGeometry> {
  if (!node.groupContext) return geometry;
  const parent = node.groupContext.parentBounds;
  const transform = node.groupContext.parentTransform;
  const nextBounds = { ...node.bounds, ...geometry };
  const scaleX = parent.width / (transform.childWidth || parent.width || 1);
  const scaleY = parent.height / (transform.childHeight || parent.height || 1);
  const center = inverseRotatePoint(
    nextBounds.left + nextBounds.width / 2,
    nextBounds.top + nextBounds.height / 2,
    parent,
  );
  const unrotatedLeft = center.x - nextBounds.width / 2;
  const unrotatedTop = center.y - nextBounds.height / 2;
  const scaledLeft = parent.flipHorizontal
    ? parent.left + parent.width - unrotatedLeft - nextBounds.width
    : unrotatedLeft - parent.left;
  const scaledTop = parent.flipVertical
    ? parent.top + parent.height - unrotatedTop - nextBounds.height
    : unrotatedTop - parent.top;
  const source: Partial<PptShapeGeometry> = {};
  if (geometry.left !== undefined) source.left = scaledLeft / scaleX + transform.childLeft;
  if (geometry.top !== undefined) source.top = scaledTop / scaleY + transform.childTop;
  if (geometry.width !== undefined) source.width = nextBounds.width / scaleX;
  if (geometry.height !== undefined) source.height = nextBounds.height / scaleY;
  if (geometry.rotation !== undefined)
    source.rotation = normalizeRotation(geometry.rotation - (parent.rotation ?? 0));
  if (geometry.flipHorizontal !== undefined)
    source.flipHorizontal = Boolean(parent.flipHorizontal) !== Boolean(geometry.flipHorizontal);
  if (geometry.flipVertical !== undefined)
    source.flipVertical = Boolean(parent.flipVertical) !== Boolean(geometry.flipVertical);
  return source;
}

function groupContextOf(shape: PptShapeModel): PptSceneGroupContext | undefined {
  return pptSceneGroupContexts.get(shape);
}

function inverseRotatePoint(x: number, y: number, group: PptSceneRect): { x: number; y: number } {
  if (!group.rotation) return { x, y };
  const angle = (-group.rotation * Math.PI) / 180;
  const centerX = group.left + group.width / 2;
  const centerY = group.top + group.height / 2;
  const dx = x - centerX;
  const dy = y - centerY;
  return {
    x: centerX + dx * Math.cos(angle) - dy * Math.sin(angle),
    y: centerY + dx * Math.sin(angle) + dy * Math.cos(angle),
  };
}

export function normalizePptGeometry(geometry: PptShapeGeometry = {}): PptSceneRect {
  return {
    left: geometry.left ?? 0,
    top: geometry.top ?? 0,
    width: geometry.width ?? 0,
    height: geometry.height ?? 0,
    rotation: geometry.rotation,
    flipHorizontal: geometry.flipHorizontal,
    flipVertical: geometry.flipVertical,
  };
}

export function findPptSceneSlide(
  scene: PptPresentationScene,
  slideIdOrIndex: string | number,
): PptSlideScene | null {
  if (typeof slideIdOrIndex === 'number') {
    return scene.slides.find((slide) => slide.index === slideIdOrIndex) ?? null;
  }
  return (
    scene.slides.find((slide) => slide.id === slideIdOrIndex || slide.name === slideIdOrIndex) ??
    null
  );
}

export function findPptSceneNode(slide: PptSlideScene, shapeIdOrName: string): PptSceneNode | null {
  const stack = [...slide.nodes, ...slide.groups];
  while (stack.length > 0) {
    const node = stack.shift()!;
    if (node.id === shapeIdOrName || node.name === shapeIdOrName) return node;
    stack.unshift(...node.children);
  }
  return null;
}
