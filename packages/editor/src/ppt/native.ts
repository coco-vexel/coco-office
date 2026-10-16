import type { PptSceneNode, PptSlideScene } from './scene.js';
import { PptEditorSession, type PptEditorCommand, type PptEditorSelection } from './session.js';
import type { PptEditorView } from './view.js';

interface NativePptDragState {
  mode: 'move' | 'resize';
  slideId: string;
  shapeId: string;
  slideSize: PptSlideScene['size'];
  startClientX: number;
  startClientY: number;
  startBounds: PptSceneNode['bounds'];
  scale: number;
  element: HTMLElement;
  active: boolean;
}

interface PptSnapGuide {
  axis: 'x' | 'y';
  value: number;
  kind: 'slide' | 'shape';
  referenceShapeId?: string;
}

type PptSnapRole = 'start' | 'center' | 'end';

interface PptSnapResult {
  bounds: PptSceneNode['bounds'];
  guides: PptSnapGuide[];
  snapped: boolean;
}

const MIN_PPT_DRAG_VISIBLE_EMU = 240000;
const MIN_PPT_RESIZE_SIZE_EMU = 91440;
const PPT_KEYBOARD_NUDGE_EMU = 91440;
const PPT_DRAG_START_THRESHOLD_PX = 4;
const PPT_DRAG_SNAP_THRESHOLD_PX = 7;
const PPT_THUMB_WIDTH_PX = 128;
const PPT_THUMB_MARGIN_BOTTOM_PX = 8;
const PPT_THUMB_OVERSCAN = 2;
const PPT_THUMB_RASTER_SCALE = 2;
const PPT_THUMB_LAYOUT_SCALE = 4;

type PptTextAlignment = NonNullable<NonNullable<PptSceneNode['textStyle']>['alignment']>;
type NativePptTextStyle = NonNullable<PptSceneNode['textStyle']> & {
  textFill?: 'none';
  textGradient?: PptGradientModel;
  textOutlineColor?: string;
  textShadow?: NonNullable<PptSceneNode['shadow']>;
};
type PptCustomGeometryPath = NonNullable<PptSceneNode['customGeometry']>['paths'][number];
type PptCustomGeometryCommand = PptCustomGeometryPath['commands'][number];
type PptGradientModel = NonNullable<PptSceneNode['fillGradient']>;
interface PptThumbnailCanvasCacheEntry {
  readonly fingerprint: string;
  readonly canvas: HTMLCanvasElement;
}
type PptChartModel = NonNullable<PptSceneNode['chart']>;
type PptChartData = NonNullable<PptChartModel['data']>;
type PptChartSeriesList = PptChartData['series'];
type PptChartSeries = PptChartSeriesList[number];
type PptTableMerge = NonNullable<NonNullable<PptSceneNode['table']>['mergedCells']>[number];
type PptTableCellStyle = NonNullable<
  NonNullable<PptSceneNode['table']>['cellStyles']
>[number][number];
type PptParagraphStyle = NonNullable<PptSceneNode['textParagraphs']>[number]['style'];

export interface NativePptEditorOptions {
  styleIsolation?: 'shadow' | 'none';
  onChange?: (event: { revision: number; bytes: Uint8Array }) => void | Promise<void>;
  onObjectAction?: (event: {
    action:
      | 'blocked-text-edit'
      | 'cycle-selection'
      | 'nudge-geometry'
      | 'snap-geometry'
      | 'copy-shape'
      | 'paste-shape'
      | 'delete-shape'
      | 'duplicate-shape'
      | 'bring-shape-to-front'
      | 'send-shape-to-back';
    slideId: string;
    shapeId: string;
    kind: PptSceneNode['kind'];
    name?: string;
    reason: string;
    candidateIndex?: number;
    candidateCount?: number;
  }) => void;
  onTextEditStateChange?: (event: {
    state: 'start' | 'commit' | 'cancel';
    slideId: string;
    shapeId: string;
    text: string;
  }) => void;
  onTableCellSelectionChange?: (event: {
    slideId: string;
    shapeId: string;
    row: number;
    column: number;
    text: string;
  }) => void;
}

export interface NativePptEditor {
  session(): PptEditorSession;
  selection(): PptEditorSelection;
  dispatch(command: PptEditorCommand): Promise<void>;
  loadPresentation(bytes: Uint8Array | ArrayBuffer): Promise<void>;
  toPptx(): Promise<Uint8Array>;
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): Promise<void>;
  redo(): Promise<void>;
  destroy(): Promise<void>;
}

export function renderNativePptEditor(
  host: HTMLElement,
  view: PptEditorView,
  options: NativePptEditorOptions = {},
): NativePptEditor {
  return new NativePptEditorController(host, view, options);
}

function isFilledCircleLikeNode(node: PptSceneNode): boolean {
  if (node.kind === 'line') return false;
  if (node.kind !== 'geometricShape') return false;
  if (node.customGeometry) return false;
  if (node.presetGeometry && node.presetGeometry !== 'ellipse') return false;
  if (node.text || !node.fillColor) return false;
  const min = Math.min(node.bounds.width, node.bounds.height);
  const max = Math.max(node.bounds.width, node.bounds.height);
  if (min <= 0) return false;
  return max / min < 1.2;
}

function clampedPptDragBounds(
  current: PptSceneNode['bounds'],
  slideSize: PptSlideScene['size'],
  next: Partial<PptSceneNode['bounds']>,
): PptSceneNode['bounds'] {
  const width = Math.max(MIN_PPT_RESIZE_SIZE_EMU, next.width ?? current.width);
  const height = Math.max(MIN_PPT_RESIZE_SIZE_EMU, next.height ?? current.height);
  const visibleX = Math.min(MIN_PPT_DRAG_VISIBLE_EMU, width);
  const visibleY = Math.min(MIN_PPT_DRAG_VISIBLE_EMU, height);
  const left = clampNumber(
    next.left ?? current.left,
    -width + visibleX,
    slideSize.width - visibleX,
  );
  const top = clampNumber(next.top ?? current.top, -height + visibleY, slideSize.height - visibleY);
  return {
    ...current,
    left,
    top,
    width,
    height,
  };
}

function clampNumber(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

function pptSnapTargetsForAxis(
  slide: PptSlideScene,
  shapeId: string,
  axis: PptSnapGuide['axis'],
): Array<{
  value: number;
  kind: PptSnapGuide['kind'];
  role: PptSnapRole;
  referenceShapeId?: string;
}> {
  const slideSize = axis === 'x' ? slide.size.width : slide.size.height;
  const targets: Array<{
    value: number;
    kind: PptSnapGuide['kind'];
    role: PptSnapRole;
    referenceShapeId?: string;
  }> = [
    { value: 0, kind: 'slide', role: 'start' },
    { value: slideSize / 2, kind: 'slide', role: 'center' },
    { value: slideSize, kind: 'slide', role: 'end' },
  ];
  for (const node of slide.nodes) {
    if (!isPptSnapReferenceNode(node, shapeId)) continue;
    const start = axis === 'x' ? node.bounds.left : node.bounds.top;
    const length = axis === 'x' ? node.bounds.width : node.bounds.height;
    targets.push(
      { value: start, kind: 'shape', role: 'start', referenceShapeId: node.id },
      { value: start + length / 2, kind: 'shape', role: 'center', referenceShapeId: node.id },
      { value: start + length, kind: 'shape', role: 'end', referenceShapeId: node.id },
    );
  }
  return targets;
}

function isPptSnapReferenceNode(node: PptSceneNode, activeShapeId: string): boolean {
  return (
    node.id !== activeShapeId &&
    node.kind !== 'line' &&
    !node.locked &&
    !node.id.includes(':') &&
    node.capabilities.select &&
    node.capabilities.editGeometry &&
    node.bounds.width > MIN_PPT_RESIZE_SIZE_EMU &&
    node.bounds.height > MIN_PPT_RESIZE_SIZE_EMU
  );
}

function closestPptSnap(
  sources: Array<{ value: number; offset: number; role: PptSnapRole }>,
  targets: Array<{
    value: number;
    kind: PptSnapGuide['kind'];
    role: PptSnapRole;
    referenceShapeId?: string;
  }>,
  threshold: number,
  options: { matchRole?: boolean } = {},
): {
  source: { value: number; offset: number; role: PptSnapRole };
  target: {
    value: number;
    kind: PptSnapGuide['kind'];
    role: PptSnapRole;
    referenceShapeId?: string;
  };
  distance: number;
} | null {
  let closest: {
    source: { value: number; offset: number; role: PptSnapRole };
    target: {
      value: number;
      kind: PptSnapGuide['kind'];
      role: PptSnapRole;
      referenceShapeId?: string;
    };
    distance: number;
  } | null = null;
  for (const source of sources) {
    for (const target of targets) {
      if (options.matchRole !== false && source.role !== target.role) continue;
      const distance = Math.abs(source.value - target.value);
      if (distance > threshold) continue;
      if (!closest || distance < closest.distance) {
        closest = { source, target, distance };
      }
    }
  }
  return closest;
}

function snappedResizeLength(
  origin: number,
  currentLength: number,
  source: { offset: number },
  targetValue: number,
): number {
  const centerOffset = currentLength / 2;
  const next =
    Math.abs(source.offset - centerOffset) < 1 ? (targetValue - origin) * 2 : targetValue - origin;
  return Math.max(MIN_PPT_RESIZE_SIZE_EMU, next);
}

function isShapeCollectionEditable(node: PptSceneNode): boolean {
  return (
    !node.locked &&
    !node.id.includes(':') &&
    (node.capabilities.editGeometry ||
      node.capabilities.editText ||
      node.capabilities.editShapeStyle ||
      node.capabilities.editTableCells)
  );
}

function parentGroupDragNode(node: PptSceneNode): PptSceneNode {
  const context = node.groupContext;
  if (!context) return node;
  const groupNode: PptSceneNode = {
    ...node,
    id: context.parentId,
    name: context.parentName,
    kind: 'group',
    bounds: context.parentBounds,
    text: undefined,
    textStyle: undefined,
    textParagraphs: undefined,
    textBody: undefined,
    groupContext: undefined,
    locked: false,
    editable: true,
    capabilities: {
      select: true,
      editText: false,
      editTextStyle: false,
      editTableCells: false,
      editGeometry: true,
      editShapeStyle: false,
      reason: 'Parent group',
    },
    children: [],
  };
  return groupNode;
}

function isCircleLikeCustomPictureNode(node: PptSceneNode): boolean {
  if (node.kind !== 'picture' || !node.customGeometry || node.customGeometry.paths.length !== 1)
    return false;
  const min = Math.min(node.bounds.width, node.bounds.height);
  const max = Math.max(node.bounds.width, node.bounds.height);
  if (min <= 0 || max / min > 1.12) return false;
  const commands = node.customGeometry.paths[0]?.commands ?? [];
  const cubicCount = commands.filter(
    (command: PptCustomGeometryCommand) => command.kind === 'cubicBezierTo',
  ).length;
  return (
    cubicCount >= 4 &&
    commands.some((command: PptCustomGeometryCommand) => command.kind === 'close')
  );
}

class NativePptEditorController implements NativePptEditor {
  private readonly root: HTMLElement | ShadowRoot;
  private readonly sessionValue: PptEditorSession;
  private readonly imageResources: PptEditorView['images'];
  private readonly fontResources: PptEditorView['fonts'];
  private readonly referencedFontFamilies: string[];
  private readonly slideList = document.createElement('div');
  private readonly canvas = document.createElement('div');
  private readonly contextMenu = document.createElement('div');
  private readonly snapGuides = document.createElement('div');
  private readonly options: NativePptEditorOptions;
  private readonly handleCanvasClick = () => void this.dispatchCurrentSlideSelection();
  private readonly handleWindowMouseMove = (event: MouseEvent) => void this.updateDrag(event);
  private readonly handleWindowMouseUp = (event: MouseEvent) => void this.finishDrag(event);
  private readonly handleWindowKeyDown = (event: KeyboardEvent) => void this.handleKeyDown(event);
  private readonly handleWindowPointerDown = (event: PointerEvent) =>
    void this.dismissContextMenu(event);
  private readonly handleWindowScroll = () => this.hideContextMenu();
  private readonly handleCanvasSelectStart = (event: Event) => this.handleSelectStart(event);
  private readonly handleSlideListScroll = () => this.scheduleThumbnailRender();
  private dragState: NativePptDragState | null = null;
  private suppressClickAfterDrag = false;
  private clipboardShape: { slideId: string; shapeId: string } | null = null;
  private thumbnailFrame = 0;
  private thumbnailRasterEpoch = 0;
  private readonly thumbnailCanvasCache = new Map<string, PptThumbnailCanvasCacheEntry>();
  private readonly pendingCommands = new Set<Promise<void>>();

  constructor(host: HTMLElement, view: PptEditorView, options: NativePptEditorOptions) {
    this.options = options;
    this.root = options.styleIsolation === 'none' ? host : host.attachShadow({ mode: 'open' });
    this.sessionValue = new PptEditorSession(view);
    this.imageResources = view.images;
    this.fontResources = view.fonts;
    this.referencedFontFamilies = usedFontFamiliesInView(view);
    const fontLinksReady = ensureGoogleFontLinks(this.fontResources, this.referencedFontFamilies);
    exposeNativePptFontReady(this.scheduleFontRenderRefresh(fontLinksReady));
    this.mount();
    this.sessionValue.subscribe('change', async (event) => {
      this.thumbnailRasterEpoch++;
      this.render();
      await this.options.onChange?.({
        revision: event.revision,
        bytes: await this.sessionValue.toPptx(),
      });
    });
    this.sessionValue.subscribe('selectionChange', () => this.renderSelection());
    this.render();
  }

  session(): PptEditorSession {
    return this.sessionValue;
  }

  selection(): PptEditorSelection {
    return this.sessionValue.selection();
  }

  async dispatch(command: PptEditorCommand): Promise<void> {
    const pending = this.dispatchAndRender(command);
    this.pendingCommands.add(pending);
    try {
      await pending;
    } finally {
      this.pendingCommands.delete(pending);
    }
  }

  private async dispatchAndRender(command: PptEditorCommand): Promise<void> {
    const previousSlideId = this.currentSlide()?.id;
    await this.sessionValue.dispatch(command);
    if (command.kind === 'select-slide' || command.kind === 'select-shape') {
      if (this.currentSlide()?.id !== previousSlideId) {
        this.render();
        return;
      }
      this.renderSelection();
      return;
    }
    this.render();
  }

  async loadPresentation(bytes: Uint8Array | ArrayBuffer): Promise<void> {
    await this.sessionValue.loadPresentation(bytes);
    this.thumbnailCanvasCache.clear();
    this.thumbnailRasterEpoch++;
    this.render();
  }

  async toPptx(): Promise<Uint8Array> {
    // Export can follow blur in the same event turn, while its command is still
    // awaiting the undo snapshot. Commit active text/table edits and await them.
    this.root
      .querySelector<HTMLElement>('[contenteditable="true"]')
      ?.dispatchEvent(new FocusEvent('blur'));
    await Promise.all([...this.pendingCommands]);
    return this.sessionValue.toPptx();
  }

  canUndo(): boolean {
    return this.sessionValue.canUndo();
  }

  canRedo(): boolean {
    return this.sessionValue.canRedo();
  }

  async undo(): Promise<void> {
    await this.sessionValue.undo();
    this.render();
  }

  async redo(): Promise<void> {
    await this.sessionValue.redo();
    this.render();
  }

  async destroy(): Promise<void> {
    this.canvas.removeEventListener('click', this.handleCanvasClick);
    this.canvas.removeEventListener('selectstart', this.handleCanvasSelectStart);
    this.slideList.removeEventListener('scroll', this.handleSlideListScroll);
    if (this.thumbnailFrame) cancelAnimationFrame(this.thumbnailFrame);
    this.thumbnailRasterEpoch++;
    window.removeEventListener('mousemove', this.handleWindowMouseMove);
    window.removeEventListener('mouseup', this.handleWindowMouseUp);
    window.removeEventListener('keydown', this.handleWindowKeyDown);
    window.removeEventListener('pointerdown', this.handleWindowPointerDown);
    window.removeEventListener('scroll', this.handleWindowScroll, true);
    await this.sessionValue.destroy();
    while (this.root.firstChild) this.root.removeChild(this.root.firstChild);
  }

  private mount(): void {
    const style = document.createElement('style');
    style.textContent = `
      ${fontFaceCss(this.fontResources)}
      :host, .coco-ppt-editor { box-sizing: border-box; width: 100%; height: 100%; min-height: 360px; }
      .coco-ppt-editor { display: grid; grid-template-columns: 160px minmax(0, 1fr); background: #f4f6f8; color: #17202a; font: 13px/1.4 system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
      .coco-ppt-slides { overflow: auto; border-right: 1px solid #d7dde5; background: #edf1f5; padding: 8px; }
      .coco-ppt-thumb { width: 128px; aspect-ratio: 16 / 9; margin: 0 0 8px; border: 1px solid #c7d0dc; background: #fff; display: block; position: relative; overflow: hidden; cursor: pointer; padding: 0; }
      .coco-ppt-thumb[data-selected="true"] { border-color: #2563eb; box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.18); }
      .coco-ppt-thumb-number { position: absolute; left: 5px; top: 4px; z-index: 5; min-width: 16px; height: 16px; border-radius: 999px; background: rgba(15, 23, 42, 0.64); color: #fff; font-size: 10px; line-height: 16px; text-align: center; pointer-events: none; }
      .coco-ppt-thumb-preview { position: absolute; inset: 0; overflow: hidden; background: #fff; pointer-events: none; }
      .coco-ppt-thumb-placeholder { position: absolute; inset: 0; display: grid; place-items: center; color: #475569; font-size: 14px; pointer-events: none; }
      .coco-ppt-thumb-slide { position: relative; overflow: hidden; transform-origin: top left; pointer-events: none; }
      .coco-ppt-thumb-canvas { position: absolute; inset: 0; width: 100%; height: 100%; display: block; pointer-events: none; }
      .coco-ppt-thumb-node { position: absolute; box-sizing: border-box; overflow: hidden; border: 0 solid transparent; white-space: pre-wrap; overflow-wrap: anywhere; padding: 0; display: flex; align-items: flex-start; justify-content: flex-start; line-height: 1.2; font-family: "Calibri", "Microsoft YaHei", SimSun, "Segoe UI", Arial, sans-serif; user-select: none; pointer-events: none; }
      .coco-ppt-thumb-node > img { width: 100%; height: 100%; object-fit: fill; display: block; pointer-events: none; }
      .coco-ppt-stage { min-width: 0; overflow: auto; display: grid; place-items: center; padding: 24px; }
      .coco-ppt-canvas { position: relative; overflow: hidden; background: #fff; box-shadow: 0 10px 30px rgba(15, 23, 42, 0.16); transform-origin: center; user-select: none; }
      .coco-ppt-node { position: absolute; box-sizing: border-box; overflow: hidden; border: 0 solid transparent; white-space: pre-wrap; overflow-wrap: anywhere; padding: 0; display: flex; align-items: flex-start; justify-content: flex-start; line-height: 1.2; font-family: "Calibri", "Microsoft YaHei", SimSun, "Segoe UI", Arial, sans-serif; user-select: none; }
      .coco-ppt-node > img { width: 100%; height: 100%; object-fit: fill; display: block; pointer-events: none; }
      .coco-ppt-css-geometry { position: absolute; inset: 0; pointer-events: none; z-index: 0; box-sizing: border-box; }
      .coco-ppt-custom-geometry, .coco-ppt-preset-geometry { position: absolute; inset: 0; width: 100%; height: 100%; pointer-events: none; z-index: 0; }
      .coco-ppt-gradient-outline, .coco-ppt-solid-outline { position: absolute; inset: 0; width: 100%; height: 100%; pointer-events: none; z-index: 0; }
      .coco-ppt-line-geometry { position: absolute; inset: 0; width: 100%; height: 100%; overflow: visible; pointer-events: none; z-index: 0; }
      .coco-ppt-text-layer { position: relative; z-index: 1; width: 100%; min-height: 100%; box-sizing: border-box; white-space: inherit; overflow-wrap: inherit; display: inherit; align-items: inherit; justify-content: inherit; }
      .coco-ppt-rich-text { width: 100%; box-sizing: border-box; white-space: pre-wrap; overflow-wrap: anywhere; }
      .coco-ppt-distributed-text { width: 100%; display: flex; justify-content: space-between; align-items: baseline; }
      .coco-ppt-paragraph { margin: 0; min-height: 1em; }
      .coco-ppt-paragraph[data-bullet="true"] { padding-left: 1em; text-indent: -1em; }
      .coco-ppt-bullet { display: inline-block; width: 1em; text-indent: 0; }
      .coco-ppt-table { width: 100%; height: 100%; table-layout: fixed; border-collapse: collapse; font: 12px/1.25 "Calibri", "Microsoft YaHei", "Segoe UI", sans-serif; color: inherit; background: transparent; }
      .coco-ppt-table td { border: 1px solid rgba(148, 163, 184, 0.8); padding: 4px 7px; vertical-align: middle; text-align: left; white-space: pre-wrap; overflow-wrap: anywhere; }
      .coco-ppt-table td[contenteditable="true"] { cursor: text; outline: 2px solid #2563eb; outline-offset: -2px; }
      .coco-ppt-node[contenteditable="true"] { cursor: text; outline: 2px solid rgba(37, 99, 235, 0.3); user-select: text; }
      .coco-ppt-text-layer[contenteditable="true"] { cursor: text; outline: 2px solid rgba(37, 99, 235, 0.3); user-select: text; }
      .coco-ppt-node[data-kind="picture"] { background: transparent; }
      .coco-ppt-unsupported-image { position: absolute; inset: 0; display: grid; place-items: center; box-sizing: border-box; border: 1px dashed rgba(100, 116, 139, 0.48); background: repeating-linear-gradient(135deg, rgba(148, 163, 184, 0.12) 0 8px, rgba(248, 250, 252, 0.1) 8px 16px); color: rgba(51, 65, 85, 0.72); font: 11px/1.2 "Segoe UI", sans-serif; pointer-events: none; }
      .coco-ppt-thumb-node .coco-ppt-unsupported-image { font-size: 8px; border-width: 0.5px; }
      .coco-ppt-node[data-selected="true"] { outline: 2px solid rgba(37, 99, 235, 0.38); }
      .coco-ppt-node[data-draggable="true"] { cursor: move; }
      .coco-ppt-node[data-dragging="true"] { outline: 1px dashed rgba(15, 23, 42, 0.68); outline-offset: 2px; }
      .coco-ppt-node[data-snap-reference="true"] { outline: 1px dashed rgba(245, 158, 11, 0.9); outline-offset: 2px; box-shadow: 0 0 0 2px rgba(245, 158, 11, 0.16); }
      .coco-ppt-parent-group-outline { position: absolute; box-sizing: border-box; border: 1px dashed rgba(16, 185, 129, 0.9); background: transparent; pointer-events: none; z-index: 850; }
      .coco-ppt-parent-group-outline[data-selected="true"] { border-style: solid; border-color: rgba(5, 150, 105, 0.95); box-shadow: 0 0 0 2px rgba(16, 185, 129, 0.16); }
      .coco-ppt-parent-group-handle { position: absolute; left: -9px; top: -9px; width: 18px; height: 18px; border-radius: 3px; border: 1px solid rgba(4, 120, 87, 0.95); background: rgba(110, 231, 183, 0.95); cursor: move; pointer-events: auto; box-shadow: 0 1px 4px rgba(15, 23, 42, 0.24); }
      .coco-ppt-parent-group-edge { position: absolute; pointer-events: auto; cursor: move; }
      .coco-ppt-parent-group-edge[data-edge="top"] { left: -4px; right: -4px; top: -6px; height: 12px; }
      .coco-ppt-parent-group-edge[data-edge="right"] { top: -4px; bottom: -4px; right: -6px; width: 12px; }
      .coco-ppt-parent-group-edge[data-edge="bottom"] { left: -4px; right: -4px; bottom: -6px; height: 12px; }
      .coco-ppt-parent-group-edge[data-edge="left"] { top: -4px; bottom: -4px; left: -6px; width: 12px; }
      .coco-ppt-node[data-editing="true"] { cursor: text; outline: 2px solid #2563eb; z-index: 30; }
      .coco-ppt-canvas[data-interaction-mode="dragging"] .coco-ppt-node:not([data-dragging="true"]) { cursor: default; }
      .coco-ppt-canvas[data-interaction-mode="editing"] .coco-ppt-node:not([data-editing="true"]) { cursor: default; }
      .coco-ppt-resize-handle { position: absolute; right: -5px; bottom: -5px; width: 10px; height: 10px; border: 1px solid #fff; background: #2563eb; cursor: nwse-resize; display: none; }
      .coco-ppt-node[data-selected="true"][data-resizable="true"] > .coco-ppt-resize-handle { display: block; }
      .coco-ppt-node[data-editing="true"] > .coco-ppt-resize-handle { display: none; }
      .coco-ppt-snap-guides { position: absolute; inset: 0; z-index: 900; pointer-events: none; }
      .coco-ppt-snap-guide { position: absolute; background: rgba(37, 99, 235, 0.62); }
      .coco-ppt-snap-guide[data-kind="slide"] { background: rgba(100, 116, 139, 0.42); }
      .coco-ppt-snap-guide[data-axis="x"] { top: 0; bottom: 0; width: 1px; }
      .coco-ppt-snap-guide[data-axis="y"] { left: 0; right: 0; height: 1px; }
      .coco-ppt-context-menu { position: fixed; z-index: 1000; min-width: 148px; padding: 4px; background: #fff; border: 1px solid #cbd5e1; box-shadow: 0 12px 28px rgba(15, 23, 42, 0.18); display: none; }
      .coco-ppt-context-menu[data-open="true"] { display: block; }
      .coco-ppt-context-menu button { width: 100%; min-height: 28px; display: flex; align-items: center; justify-content: space-between; gap: 12px; border: 0; background: transparent; color: #17202a; padding: 5px 8px; font: inherit; text-align: left; cursor: pointer; }
      .coco-ppt-context-menu button:hover:not(:disabled) { background: #eaf2ff; color: #0f3f8f; }
      .coco-ppt-context-menu button:disabled { color: #94a3b8; cursor: default; }
      .coco-ppt-context-menu span { color: #64748b; font-size: 11px; }
      .coco-ppt-context-menu hr { border: 0; border-top: 1px solid #e2e8f0; margin: 4px 0; }
    `;
    const shell = document.createElement('div');
    shell.className = 'coco-ppt-editor';
    this.slideList.className = 'coco-ppt-slides';
    const stage = document.createElement('div');
    stage.className = 'coco-ppt-stage';
    this.canvas.className = 'coco-ppt-canvas';
    this.contextMenu.className = 'coco-ppt-context-menu';
    this.snapGuides.className = 'coco-ppt-snap-guides';
    stage.append(this.canvas);
    shell.append(this.slideList, stage);
    this.root.append(style, shell, this.contextMenu);
    this.canvas.addEventListener('click', this.handleCanvasClick);
    this.canvas.addEventListener('selectstart', this.handleCanvasSelectStart);
    this.slideList.addEventListener('scroll', this.handleSlideListScroll, { passive: true });
    window.addEventListener('mousemove', this.handleWindowMouseMove);
    window.addEventListener('mouseup', this.handleWindowMouseUp);
    window.addEventListener('keydown', this.handleWindowKeyDown);
    window.addEventListener('pointerdown', this.handleWindowPointerDown);
    window.addEventListener('scroll', this.handleWindowScroll, true);
  }

  private scheduleFontRenderRefresh(fontLinksReady: Promise<void>): Promise<void> {
    if (typeof document === 'undefined' || !document.fonts) return Promise.resolve();
    return fontLinksReady
      .then(() =>
        Promise.all(
          [...new Set(this.fontResources.map((font) => font.family))].map((family) =>
            document.fonts.load(`16px ${cssFontFamilyStack(family)}`),
          ),
        ),
      )
      .then(() =>
        Promise.all(
          this.referencedFontFamilies.map((family) =>
            document.fonts.load(`16px ${cssFontFamilyStack(family)}`),
          ),
        ),
      )
      .then(() => document.fonts.ready)
      .then(() => {
        this.render();
      })
      .catch(() => {});
  }

  private render(): void {
    const scene = this.sessionValue.scene();
    this.slideList.replaceChildren(
      ...scene.slides.map((slide) => {
        const thumb = document.createElement('button');
        thumb.className = 'coco-ppt-thumb';
        thumb.type = 'button';
        thumb.dataset.slideId = slide.id;
        thumb.dataset.slideIndex = String(slide.index);
        thumb.dataset.selected =
          this.sessionValue.selection()?.slideId === slide.id ? 'true' : 'false';
        const preview = document.createElement('div');
        preview.className = 'coco-ppt-thumb-preview';
        const placeholder = document.createElement('div');
        placeholder.className = 'coco-ppt-thumb-placeholder';
        const number = document.createElement('span');
        number.className = 'coco-ppt-thumb-number';
        number.textContent = String(slide.index + 1);
        thumb.append(preview, placeholder, number);
        thumb.addEventListener(
          'click',
          () => void this.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id }),
        );
        return thumb;
      }),
    );
    this.renderVisibleThumbnails(scene.slides);
    const selected = this.currentSlide();
    if (!selected) {
      this.canvas.replaceChildren();
      this.hideContextMenu();
      return;
    }
    const scale = this.computeScale(selected);
    this.canvas.style.width = `${selected.size.width * scale}px`;
    this.canvas.style.height = `${selected.size.height * scale}px`;
    this.canvas.style.background = cssBackground(selected.background, this.imageResources);
    this.canvas.replaceChildren(
      ...selected.nodes.map((node) => this.renderNode(node, scale, selected)),
      this.snapGuides,
    );
    this.hideSnapGuides();
    this.hideContextMenu();
    this.renderSelection();
  }

  private renderVisibleThumbnails(slides: PptSlideScene[]): void {
    const thumbs = Array.from(
      this.slideList.querySelectorAll<HTMLButtonElement>('.coco-ppt-thumb'),
    );
    const range = this.visibleThumbnailRange(slides.length);
    thumbs.forEach((thumb, index) => {
      const preview = thumb.querySelector<HTMLElement>('.coco-ppt-thumb-preview');
      const placeholder = thumb.querySelector<HTMLElement>('.coco-ppt-thumb-placeholder');
      if (!preview || !placeholder) return;
      if (index < range.start || index >= range.end) {
        preview.replaceChildren();
        delete thumb.dataset.previewRendered;
        delete thumb.dataset.previewMode;
        placeholder.hidden = false;
        return;
      }
      if (thumb.dataset.previewRendered === 'true') return;
      const slide = slides[index];
      if (!slide) return;
      const fingerprint = thumbnailFingerprint(slide);
      const cached = this.thumbnailCanvasCache.get(slide.id);
      if (cached?.fingerprint === fingerprint) {
        preview.replaceChildren(cloneThumbnailCanvas(cached.canvas));
        thumb.dataset.previewMode = 'canvas';
      } else {
        preview.replaceChildren(this.renderThumbnailSlide(slide));
        thumb.dataset.previewMode = 'dom';
        void this.rasterizeThumbnail(thumb, slide);
      }
      thumb.dataset.previewRendered = 'true';
      placeholder.hidden = true;
    });
  }

  private visibleThumbnailRange(count: number): { start: number; end: number } {
    const itemHeight = PPT_THUMB_WIDTH_PX * (9 / 16) + PPT_THUMB_MARGIN_BOTTOM_PX;
    const viewportHeight = this.slideList.clientHeight || itemHeight * Math.min(count, 5);
    const first = Math.max(
      0,
      Math.floor(this.slideList.scrollTop / itemHeight) - PPT_THUMB_OVERSCAN,
    );
    const visible = Math.ceil(viewportHeight / itemHeight) + PPT_THUMB_OVERSCAN * 2;
    const selection = this.sessionValue.selection();
    const selectedIndex =
      selection?.slideId !== undefined
        ? this.sessionValue.scene().slides.findIndex((slide) => slide.id === selection.slideId)
        : -1;
    const start =
      selectedIndex >= 0 ? Math.min(first, Math.max(0, selectedIndex - PPT_THUMB_OVERSCAN)) : first;
    const end =
      selectedIndex >= 0
        ? Math.max(first + visible, Math.min(count, selectedIndex + PPT_THUMB_OVERSCAN + 1))
        : first + visible;
    return { start, end: Math.min(count, end) };
  }

  private renderThumbnailSlide(slide: PptSlideScene): HTMLElement {
    const wrapper = document.createElement('div');
    wrapper.className = 'coco-ppt-thumb-slide';
    const layoutWidth = PPT_THUMB_WIDTH_PX * PPT_THUMB_LAYOUT_SCALE;
    const scale = layoutWidth / slide.size.width;
    wrapper.style.width = `${slide.size.width * scale}px`;
    wrapper.style.height = `${slide.size.height * scale}px`;
    wrapper.style.transform = `scale(${1 / PPT_THUMB_LAYOUT_SCALE})`;
    wrapper.style.background = cssBackground(slide.background, this.imageResources);
    wrapper.replaceChildren(
      ...slide.nodes.map((node) => this.renderNode(node, scale, slide, false)),
    );
    return wrapper;
  }

  private async rasterizeThumbnail(thumb: HTMLButtonElement, slide: PptSlideScene): Promise<void> {
    const epoch = this.thumbnailRasterEpoch;
    const fingerprint = thumbnailFingerprint(slide);
    await nextAnimationFrame();
    await nextAnimationFrame();
    if (epoch !== this.thumbnailRasterEpoch || !thumb.isConnected) return;
    const preview = thumb.querySelector<HTMLElement>('.coco-ppt-thumb-preview');
    const slideElement = preview?.querySelector<HTMLElement>('.coco-ppt-thumb-slide');
    if (!preview || !slideElement) return;
    const canvas = await rasterizeThumbnailElement(
      slideElement,
      PPT_THUMB_WIDTH_PX,
      PPT_THUMB_WIDTH_PX * (slide.size.height / slide.size.width),
      PPT_THUMB_RASTER_SCALE,
      this.root.querySelector('style')?.textContent ?? '',
      PPT_THUMB_WIDTH_PX * PPT_THUMB_LAYOUT_SCALE,
      PPT_THUMB_WIDTH_PX * PPT_THUMB_LAYOUT_SCALE * (slide.size.height / slide.size.width),
    );
    if (!canvas) return;
    if (epoch !== this.thumbnailRasterEpoch || !thumb.isConnected) return;
    const currentSlide = this.sessionValue.scene().slides.find((item) => item.id === slide.id);
    if (!currentSlide || thumbnailFingerprint(currentSlide) !== fingerprint) return;
    this.thumbnailCanvasCache.set(slide.id, { fingerprint, canvas });
    preview.replaceChildren(cloneThumbnailCanvas(canvas));
    thumb.dataset.previewMode = 'canvas';
  }

  private scheduleThumbnailRender(): void {
    if (this.thumbnailFrame) return;
    this.thumbnailFrame = requestAnimationFrame(() => {
      this.thumbnailFrame = 0;
      this.renderVisibleThumbnails(this.sessionValue.scene().slides);
    });
  }

  private renderNode(
    node: PptSceneNode,
    scale: number,
    slide: PptSlideScene,
    interactive = true,
  ): HTMLElement {
    const element = document.createElement('div');
    const lineNode = isLineNode(node);
    const frameNode = node.presetGeometry === 'frame';
    const lineAxis = lineNode ? lineAxisKind(node) : 'horizontal';
    const nodeWidth = Math.max(
      lineNode && lineAxis === 'vertical' ? 12 : 1,
      node.bounds.width * scale,
    );
    const nodeHeight = Math.max(
      lineNode && lineAxis === 'horizontal' ? 12 : 1,
      node.bounds.height * scale,
    );
    element.className = interactive ? 'coco-ppt-node' : 'coco-ppt-thumb-node';
    element.dataset.kind = node.kind;
    element.dataset.locked = String(node.locked);
    if (interactive) {
      element.dataset.draggable = String(node.capabilities.editGeometry);
      element.dataset.resizable = String(node.capabilities.editGeometry);
      element.dataset.shapeId = node.id;
    }
    element.style.left = `${node.bounds.left * scale - (lineNode && lineAxis === 'vertical' ? nodeWidth / 2 : 0)}px`;
    element.style.top = `${node.bounds.top * scale - (lineNode && lineAxis === 'horizontal' ? nodeHeight / 2 : 0)}px`;
    element.style.width = `${nodeWidth}px`;
    element.style.height = `${nodeHeight}px`;
    const transforms: string[] = [];
    if (node.bounds.rotation) transforms.push(`rotate(${node.bounds.rotation}deg)`);
    if (node.bounds.flipHorizontal) transforms.push('scaleX(-1)');
    if (node.bounds.flipVertical) transforms.push('scaleY(-1)');
    if (transforms.length > 0) element.style.transform = transforms.join(' ');
    const svgPresetNode = hasSvgPresetGeometry(node);
    const clipPath = cssClipPathForPresetGeometry(node.presetGeometry);
    const cssPresetTextNode = Boolean(
      clipPath && node.text && !node.customGeometry && !svgPresetNode,
    );
    if (node.usesBackgroundFill && node.backgroundFill) {
      applyBackgroundFill(element, node.backgroundFill, node, slide, scale, this.imageResources);
    } else if (
      node.customGeometry ||
      svgPresetNode ||
      cssPresetTextNode ||
      node.kind === 'table' ||
      lineNode ||
      frameNode
    ) {
      element.style.background = 'transparent';
    } else if (node.fillGradient && node.fillGradient.colors.length > 1) {
      element.style.background = cssLinearGradient(node.fillGradient, effectiveFillOpacity(node));
    } else if (node.fillColor) {
      element.style.background = cssColorWithOpacity(node.fillColor, effectiveFillOpacity(node));
    }
    if (shouldAllowTextOverflow(node)) element.style.overflow = 'visible';
    if (node.presetGeometry === 'roundRect' || node.presetGeometry === 'round1Rect')
      element.style.borderRadius = `${roundRectRadius(node) * scale}px`;
    if (node.presetGeometry === 'round2DiagRect') {
      const radius = `${roundRectRadius(node) * scale}px`;
      element.style.borderRadius = `${radius} 0 ${radius} 0`;
    }
    if (node.presetGeometry === 'round2SameRect') {
      const radius = `${roundRectRadius(node) * scale}px`;
      element.style.borderRadius = `${radius} ${radius} 0 0`;
    }
    if (
      node.presetGeometry === 'ellipse' ||
      isFilledCircleLikeNode(node) ||
      isCircleLikeCustomPictureNode(node)
    )
      element.style.borderRadius = '9999px';
    if (clipPath && !cssPresetTextNode && !svgPresetNode) element.style.clipPath = clipPath;
    if (frameNode) {
      element.style.borderStyle = 'solid';
      element.style.borderWidth = `${frameBorderWidth(node, scale)}px`;
      if (node.fillGradient && node.fillGradient.colors.length > 1) {
        element.style.borderImageSource = cssLinearGradient(node.fillGradient);
        element.style.borderImageSlice = '1';
      } else if (node.fillColor) {
        element.style.borderColor = node.fillColor;
      }
    } else if (
      !node.customGeometry &&
      !svgPresetNode &&
      !cssPresetTextNode &&
      !lineNode &&
      node.kind !== 'picture' &&
      node.lineColor
    ) {
      element.style.borderStyle = cssBorderStyleForLineDash(node.lineDash);
      element.style.borderWidth = `${Math.max(1, (node.lineWidth ?? 9525) * scale)}px`;
      element.style.borderColor = cssColorWithOpacity(node.lineColor, node.lineOpacity);
    }
    if (node.shadow) element.style.boxShadow = cssBoxShadow(node.shadow, scale);
    this.applyTextStyle(element, node, scale);
    applyTextMetricFallback(element, node, scale);
    if (cssPresetTextNode && clipPath)
      element.append(renderCssPresetGeometry(node, scale, clipPath));
    if (node.customGeometry && node.kind !== 'picture')
      element.append(renderCustomGeometry(node, scale));
    if (svgPresetNode) element.append(renderPresetGeometry(node, scale));
    if (lineNode) {
      element.append(renderLineGeometry(node, scale));
    } else if (node.kind === 'chart') {
      element.append(renderChartPreview(node));
    } else if (node.kind === 'table') {
      const table = renderTablePreview(node, scale, interactive);
      if (interactive) {
        table.addEventListener('click', (event) => {
          event.stopPropagation();
          if (this.consumeSuppressedClick()) return;
          this.notifyTableCellSelection(slide, node, event);
          void this.dispatch({
            kind: 'select-shape',
            slideIdOrIndex: slide.id,
            shapeIdOrName: node.id,
          });
        });
        table.addEventListener('dblclick', (event) => {
          const cell = tableCellFromEvent(event);
          if (!cell || !node.capabilities.editTableCells) return;
          event.stopPropagation();
          const row = Number(cell.dataset.row);
          const column = Number(cell.dataset.column);
          if (!Number.isInteger(row) || !Number.isInteger(column)) return;
          this.startTableCellEdit(cell, slide, node, row, column);
        });
      }
      element.append(table);
    } else if (node.kind === 'picture') {
      const resource = node.imagePath ? this.imageResources[node.imagePath] : undefined;
      const url = resource?.url;
      if (url && isBrowserRenderableImage(resource.contentType, node.imagePath)) {
        if (shouldClipCustomGeometryPicture(node, slide)) {
          element.append(renderCustomGeometryPicture(node, url));
        } else {
          const image = document.createElement('img');
          image.src = url;
          image.alt = node.name ?? '';
          applyImagePlacement(image, node);
          if (node.imageOpacity !== undefined)
            image.style.opacity = String(Math.max(0, Math.min(1, node.imageOpacity)));
          element.append(image);
        }
        if (node.lineColor) element.append(renderSolidOutline(node, scale));
      } else {
        element.dataset.unsupportedImage = node.imagePath ? 'true' : 'missing';
        if (resource?.contentType) element.dataset.imageContentType = resource.contentType;
        if (node.imagePath) element.dataset.imagePath = node.imagePath;
        element.append(renderUnsupportedImagePlaceholder(node, resource?.contentType));
      }
    } else {
      const text = node.text ?? '';
      if (
        node.customGeometry ||
        svgPresetNode ||
        cssPresetTextNode ||
        shouldCounterFlipText(node)
      ) {
        const textLayer = document.createElement('div');
        textLayer.className = 'coco-ppt-text-layer';
        applyTextCounterFlip(textLayer, node);
        appendTextContent(textLayer, node, scale, text);
        element.append(textLayer);
        applyTextWrapping(textLayer, node);
        applyTextAutoFit(textLayer, node);
      } else {
        appendTextContent(element, node, scale, text);
        applyTextWrapping(element, node);
        applyTextAutoFit(element, node);
      }
    }
    if (shouldRenderGradientOutline(node)) element.append(renderGradientOutline(node, scale));
    if (!interactive) return element;
    const handle = document.createElement('div');
    handle.className = 'coco-ppt-resize-handle';
    handle.addEventListener('mousedown', (event) => {
      if (!node.capabilities.editGeometry) return;
      event.preventDefault();
      event.stopPropagation();
      this.startDrag(event, 'resize', slide, node, scale, element);
    });
    element.append(handle);
    element.addEventListener('click', (event) => {
      this.hideContextMenu();
      event.stopPropagation();
      if (this.consumeSuppressedClick()) return;
      if (event.altKey) {
        const selection = this.selection();
        const cycledNode = cycledSelectableHit(
          slide,
          event,
          scale,
          selection?.kind === 'shape' ? selection.shapeId : node.id,
        );
        if (cycledNode) {
          void this.dispatch({
            kind: 'select-shape',
            slideIdOrIndex: slide.id,
            shapeIdOrName: cycledNode.node.id,
          });
          this.options.onObjectAction?.({
            action: 'cycle-selection',
            slideId: slide.id,
            shapeId: cycledNode.node.id,
            kind: cycledNode.node.kind,
            name: cycledNode.node.name,
            reason: `已切换到重叠对象 ${cycledNode.index + 1}/${cycledNode.count}`,
            candidateIndex: cycledNode.index,
            candidateCount: cycledNode.count,
          });
          return;
        }
      }
      const redirectedNode = this.pointerOperationNode(slide, node, event, scale);
      if (redirectedNode) {
        void this.dispatch({
          kind: 'select-shape',
          slideIdOrIndex: slide.id,
          shapeIdOrName: redirectedNode.id,
        });
        return;
      }
      void this.dispatch({
        kind: 'select-shape',
        slideIdOrIndex: slide.id,
        shapeIdOrName: node.id,
      });
    });
    element.addEventListener('contextmenu', (event) => {
      event.preventDefault();
      event.stopPropagation();
      const targetNode = this.pointerOperationNode(slide, node, event, scale) ?? node;
      void this.dispatch({
        kind: 'select-shape',
        slideIdOrIndex: slide.id,
        shapeIdOrName: targetNode.id,
      }).then(() => this.showContextMenu(event, slide, targetNode));
    });
    element.addEventListener('mousedown', (event) => {
      const targetNode = this.pointerOperationNode(slide, node, event, scale) ?? node;
      const targetElement = elementForShape(this.canvas, targetNode.id) ?? element;
      if (
        isEditingText(targetElement, event.target) ||
        event.button !== 0 ||
        !targetNode.capabilities.editGeometry
      )
        return;
      event.preventDefault();
      event.stopPropagation();
      this.startDrag(event, 'move', slide, targetNode, scale, targetElement);
    });
    element.addEventListener('dblclick', (event) => {
      event.stopPropagation();
      const redirectedNode = redirectedEditableTextHit(slide, node, event, scale);
      if (redirectedNode) {
        this.startTextEdit(
          elementForShape(this.canvas, redirectedNode.id) ?? element,
          slide,
          redirectedNode,
        );
        return;
      }
      if (!node.capabilities.editText) {
        void this.sessionValue
          .dispatch({ kind: 'select-shape', slideIdOrIndex: slide.id, shapeIdOrName: node.id })
          .then(() => this.renderSelection());
        this.options.onObjectAction?.({
          action: 'blocked-text-edit',
          slideId: slide.id,
          shapeId: node.id,
          kind: node.kind,
          name: node.name,
          reason: nonTextEditReason(node),
        });
        return;
      }
      this.startTextEdit(element, slide, node);
    });
    return element;
  }

  private pointerOperationNode(
    slide: PptSlideScene,
    current: PptSceneNode,
    event: MouseEvent,
    scale: number,
  ): PptSceneNode | null {
    const selection = this.selection();
    if (
      selection?.kind === 'shape' &&
      selection.slideId === slide.id &&
      selection.shapeId === current.id
    ) {
      return current;
    }
    return redirectedEditableTextHit(slide, current, event, scale);
  }

  private notifyTableCellSelection(slide: PptSlideScene, node: PptSceneNode, event: Event): void {
    const cell = tableCellFromEvent(event);
    if (!cell || !node.capabilities.editTableCells) return;
    const row = Number(cell.dataset.row);
    const column = Number(cell.dataset.column);
    if (!Number.isInteger(row) || !Number.isInteger(column)) return;
    this.options.onTableCellSelectionChange?.({
      slideId: slide.id,
      shapeId: node.id,
      row,
      column,
      text: tableCellRawText(cell, node, row, column),
    });
  }

  private applyTextStyle(element: HTMLElement, node: PptSceneNode, scale: number): void {
    const style = node.textStyle;
    if (style) {
      applyTextStyleToElement(element, inheritedNodeTextStyle(node, style), scale);
      if (style.alignment) {
        applyTextAlignment(element, style.alignment);
        if (style.alignment === 'center') element.style.justifyContent = 'center';
        if (style.alignment === 'right') element.style.justifyContent = 'flex-end';
      }
    }
    const body = node.kind === 'table' ? undefined : node.textBody;
    if (body) {
      const left = body.insetLeft ?? 91440;
      const top = body.insetTop ?? 45720;
      const right = body.insetRight ?? 91440;
      const bottom = body.insetBottom ?? 45720;
      const baselineOffset = textBaselineMetricOffset(node, scale);
      element.style.padding = `${top * scale + baselineOffset}px ${right * scale}px ${bottom * scale}px ${left * scale}px`;
      if (isVerticalTextBody(body)) {
        element.style.writingMode = 'vertical-rl';
        element.style.textOrientation = 'upright';
      }
      if (body.anchor === 'middle') element.style.alignItems = 'center';
      if (body.anchor === 'bottom') element.style.alignItems = 'flex-end';
      if (body.anchor === 'top') element.style.alignItems = 'flex-start';
    }
  }

  private startTextEdit(element: HTMLElement, slide: PptSlideScene, node: PptSceneNode): void {
    const originalText = node.text ?? '';
    const editableTarget = editableTextTarget(element);
    let finished = false;
    void this.sessionValue
      .dispatch({ kind: 'select-shape', slideIdOrIndex: slide.id, shapeIdOrName: node.id })
      .then(() => this.renderSelection());
    this.options.onTextEditStateChange?.({
      state: 'start',
      slideId: slide.id,
      shapeId: node.id,
      text: originalText,
    });
    editableTarget.contentEditable = 'true';
    editableTarget.spellcheck = false;
    editableTarget.setAttribute('role', 'textbox');
    editableTarget.setAttribute('aria-label', node.name ?? 'PPT text');
    element.dataset.editing = 'true';
    this.canvas.dataset.interactionMode = 'editing';
    editableTarget.textContent = originalText;
    editableTarget.focus();
    const selection = editableTarget.ownerDocument.getSelection();
    if (selection) {
      const range = editableTarget.ownerDocument.createRange();
      range.selectNodeContents(editableTarget);
      range.collapse(false);
      selection.removeAllRanges();
      selection.addRange(range);
    }
    const finish = (commit: boolean) => {
      if (finished) return;
      finished = true;
      editableTarget.contentEditable = 'false';
      editableTarget.removeAttribute('role');
      editableTarget.removeAttribute('aria-label');
      delete element.dataset.editing;
      delete this.canvas.dataset.interactionMode;
      editableTarget.removeEventListener('blur', handleBlur);
      editableTarget.removeEventListener('keydown', handleKeydown);
      editableTarget.removeEventListener('paste', handlePaste);
      const text = editableTarget.textContent ?? '';
      if (!commit) {
        editableTarget.textContent = originalText;
        this.options.onTextEditStateChange?.({
          state: 'cancel',
          slideId: slide.id,
          shapeId: node.id,
          text: originalText,
        });
        this.render();
        return;
      }
      this.options.onTextEditStateChange?.({
        state: 'commit',
        slideId: slide.id,
        shapeId: node.id,
        text,
      });
      if (text !== originalText) {
        void this.dispatch({
          kind: 'set-shape-text',
          slideIdOrIndex: slide.id,
          shapeIdOrName: node.id,
          text,
        });
      } else {
        this.render();
      }
    };
    const handleKeydown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        finish(false);
        return;
      }
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        finish(true);
      }
    };
    const handlePaste = (event: ClipboardEvent) => {
      const text = event.clipboardData?.getData('text/plain');
      if (text === undefined) return;
      event.preventDefault();
      insertPlainText(editableTarget, text);
    };
    const handleBlur = () => finish(true);
    editableTarget.addEventListener('blur', handleBlur);
    editableTarget.addEventListener('keydown', handleKeydown);
    editableTarget.addEventListener('paste', handlePaste);
  }

  private startTableCellEdit(
    cell: HTMLElement,
    slide: PptSlideScene,
    node: PptSceneNode,
    row: number,
    column: number,
  ): void {
    const originalText = tableCellRawText(cell, node, row, column);
    let finished = false;
    void this.sessionValue
      .dispatch({ kind: 'select-shape', slideIdOrIndex: slide.id, shapeIdOrName: node.id })
      .then(() => this.renderSelection());
    this.options.onTableCellSelectionChange?.({
      slideId: slide.id,
      shapeId: node.id,
      row,
      column,
      text: originalText,
    });
    cell.contentEditable = 'true';
    cell.spellcheck = false;
    this.canvas.dataset.interactionMode = 'editing';
    cell.setAttribute('role', 'textbox');
    cell.setAttribute('aria-label', `${node.name ?? 'PPT table'} cell ${row + 1},${column + 1}`);
    cell.textContent = originalText;
    cell.focus();
    const selection = cell.ownerDocument.getSelection();
    if (selection) {
      const range = cell.ownerDocument.createRange();
      range.selectNodeContents(cell);
      range.collapse(false);
      selection.removeAllRanges();
      selection.addRange(range);
    }
    const finish = (commit: boolean) => {
      if (finished) return;
      finished = true;
      cell.contentEditable = 'false';
      delete this.canvas.dataset.interactionMode;
      cell.removeAttribute('role');
      cell.removeAttribute('aria-label');
      cell.removeEventListener('blur', handleBlur);
      cell.removeEventListener('keydown', handleKeydown);
      cell.removeEventListener('paste', handlePaste);
      const text = cell.textContent ?? '';
      if (!commit) {
        cell.textContent = originalText;
        cell.dataset.cellText = originalText;
        return;
      }
      if (text !== originalText) {
        cell.dataset.cellText = text;
        void this.dispatch({
          kind: 'set-table-cell-text',
          slideIdOrIndex: slide.id,
          shapeIdOrName: node.id,
          row,
          column,
          text,
        });
      } else {
        this.render();
      }
    };
    const handleKeydown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        finish(false);
        return;
      }
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        finish(true);
      }
    };
    const handlePaste = (event: ClipboardEvent) => {
      const text = event.clipboardData?.getData('text/plain');
      if (text === undefined) return;
      event.preventDefault();
      insertPlainText(cell, text);
    };
    const handleBlur = () => finish(true);
    cell.addEventListener('blur', handleBlur);
    cell.addEventListener('keydown', handleKeydown);
    cell.addEventListener('paste', handlePaste);
  }

  private startDrag(
    event: MouseEvent,
    mode: NativePptDragState['mode'],
    slide: PptSlideScene,
    node: PptSceneNode,
    scale: number,
    element: HTMLElement,
  ): void {
    if (mode === 'resize') prepareTextResizePreview(element, node);
    this.dragState = {
      mode,
      slideId: slide.id,
      shapeId: node.id,
      slideSize: slide.size,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startBounds: { ...node.bounds },
      scale,
      element,
      active: false,
    };
    void this.sessionValue
      .dispatch({ kind: 'select-shape', slideIdOrIndex: slide.id, shapeIdOrName: node.id })
      .then(() => this.renderSelection());
  }

  private updateDrag(event: MouseEvent): void {
    const drag = this.dragState;
    if (!drag) return;
    const distance = Math.hypot(
      event.clientX - drag.startClientX,
      event.clientY - drag.startClientY,
    );
    if (distance <= PPT_DRAG_START_THRESHOLD_PX) {
      this.restoreDragElementBounds(drag);
      this.hideSnapGuides();
      drag.active = false;
      this.clearDragVisualState(drag.element, true);
      return;
    }
    drag.active = true;
    drag.element.dataset.dragging = 'true';
    this.canvas.dataset.interactionMode = 'dragging';
    const dx = (event.clientX - drag.startClientX) / drag.scale;
    const dy = (event.clientY - drag.startClientY) / drag.scale;
    if (drag.mode === 'move') {
      const { bounds } = event.altKey
        ? {
            bounds: clampedPptDragBounds(drag.startBounds, drag.slideSize, {
              left: drag.startBounds.left + dx,
              top: drag.startBounds.top + dy,
            }),
          }
        : this.snapDragBounds(drag, {
            left: drag.startBounds.left + dx,
            top: drag.startBounds.top + dy,
          });
      if (event.altKey) this.hideSnapGuides();
      drag.element.style.left = `${bounds.left * drag.scale}px`;
      drag.element.style.top = `${bounds.top * drag.scale}px`;
      return;
    }
    const { bounds } = event.altKey
      ? {
          bounds: clampedPptDragBounds(drag.startBounds, drag.slideSize, {
            width: drag.startBounds.width + dx,
            height: drag.startBounds.height + dy,
          }),
        }
      : this.snapResizeBounds(drag, {
          width: drag.startBounds.width + dx,
          height: drag.startBounds.height + dy,
        });
    if (event.altKey) this.hideSnapGuides();
    drag.element.style.width = `${bounds.width * drag.scale}px`;
    drag.element.style.height = `${bounds.height * drag.scale}px`;
  }

  private async finishDrag(event: MouseEvent): Promise<void> {
    const drag = this.dragState;
    if (!drag) return;
    this.dragState = null;
    const dx = (event.clientX - drag.startClientX) / drag.scale;
    const dy = (event.clientY - drag.startClientY) / drag.scale;
    const distance = Math.hypot(
      event.clientX - drag.startClientX,
      event.clientY - drag.startClientY,
    );
    if (distance <= PPT_DRAG_START_THRESHOLD_PX || !drag.active) {
      this.hideSnapGuides();
      this.clearDragVisualState(drag.element);
      return;
    }
    this.suppressClickAfterDrag =
      this.suppressClickAfterDrag || distance > PPT_DRAG_START_THRESHOLD_PX;
    const snap = event.altKey
      ? clampedPptDragBounds(
          drag.startBounds,
          drag.slideSize,
          drag.mode === 'move'
            ? {
                left: drag.startBounds.left + dx,
                top: drag.startBounds.top + dy,
              }
            : {
                width: drag.startBounds.width + dx,
                height: drag.startBounds.height + dy,
              },
        )
      : drag.mode === 'move'
        ? this.snapDragBounds(drag, {
            left: drag.startBounds.left + dx,
            top: drag.startBounds.top + dy,
          })
        : this.snapResizeBounds(drag, {
            width: drag.startBounds.width + dx,
            height: drag.startBounds.height + dy,
          });
    this.hideSnapGuides();
    this.clearDragVisualState(drag.element);
    await this.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: drag.slideId,
      shapeIdOrName: drag.shapeId,
      geometry: 'bounds' in snap ? snap.bounds : snap,
      preserveTextSizeOnResize: drag.mode === 'resize',
    });
    if ('snapped' in snap && snap.snapped) {
      const slide = this.currentSlide();
      const node = slide?.nodes.find((item) => item.id === drag.shapeId);
      if (slide && node) {
        this.options.onObjectAction?.({
          action: 'snap-geometry',
          slideId: slide.id,
          shapeId: node.id,
          kind: node.kind,
          name: node.name,
          reason: this.snapActionReason(snap, slide),
        });
      }
    }
  }

  private snapActionReason(snap: PptSnapResult, slide: PptSlideScene): string {
    const referenceNames = [
      ...new Set(
        snap.guides
          .map((guide) => guide.referenceShapeId)
          .filter((id): id is string => Boolean(id)),
      ),
    ]
      .map((shapeId) => slide.nodes.find((node) => node.id === shapeId))
      .filter((node): node is PptSceneNode => Boolean(node))
      .map((node) => node.name || node.id);
    if (referenceNames.length > 0) {
      return `已吸附到 ${referenceNames.join('、')}`;
    }
    return '已吸附到页面参考线';
  }

  private snapDragBounds(
    drag: NativePptDragState,
    next: Pick<PptSceneNode['bounds'], 'left' | 'top'>,
  ): PptSnapResult {
    const rawBounds = clampedPptDragBounds(drag.startBounds, drag.slideSize, next);
    const slide = this.sessionValue.scene().slides.find((item) => item.id === drag.slideId);
    if (!slide) {
      this.hideSnapGuides();
      return { bounds: rawBounds, guides: [], snapped: false };
    }
    const threshold = Math.min(180000, PPT_DRAG_SNAP_THRESHOLD_PX / drag.scale);
    const xSnap = closestPptSnap(
      [
        { offset: 0, value: rawBounds.left, role: 'start' },
        {
          offset: rawBounds.width / 2,
          value: rawBounds.left + rawBounds.width / 2,
          role: 'center',
        },
        { offset: rawBounds.width, value: rawBounds.left + rawBounds.width, role: 'end' },
      ],
      pptSnapTargetsForAxis(slide, drag.shapeId, 'x'),
      threshold,
    );
    const ySnap = closestPptSnap(
      [
        { offset: 0, value: rawBounds.top, role: 'start' },
        {
          offset: rawBounds.height / 2,
          value: rawBounds.top + rawBounds.height / 2,
          role: 'center',
        },
        { offset: rawBounds.height, value: rawBounds.top + rawBounds.height, role: 'end' },
      ],
      pptSnapTargetsForAxis(slide, drag.shapeId, 'y'),
      threshold,
    );
    const bounds = clampedPptDragBounds(drag.startBounds, drag.slideSize, {
      left: xSnap ? xSnap.target.value - xSnap.source.offset : rawBounds.left,
      top: ySnap ? ySnap.target.value - ySnap.source.offset : rawBounds.top,
    });
    const guides = [
      ...(xSnap
        ? [
            {
              axis: 'x' as const,
              value: xSnap.target.value,
              kind: xSnap.target.kind,
              referenceShapeId: xSnap.target.referenceShapeId,
            },
          ]
        : []),
      ...(ySnap
        ? [
            {
              axis: 'y' as const,
              value: ySnap.target.value,
              kind: ySnap.target.kind,
              referenceShapeId: ySnap.target.referenceShapeId,
            },
          ]
        : []),
    ];
    this.renderSnapGuides(guides, drag.scale);
    return { bounds, guides, snapped: guides.length > 0 };
  }

  private snapResizeBounds(
    drag: NativePptDragState,
    next: Pick<PptSceneNode['bounds'], 'width' | 'height'>,
  ): PptSnapResult {
    const rawBounds = clampedPptDragBounds(drag.startBounds, drag.slideSize, next);
    const slide = this.sessionValue.scene().slides.find((item) => item.id === drag.slideId);
    if (!slide) {
      this.hideSnapGuides();
      return { bounds: rawBounds, guides: [], snapped: false };
    }
    const threshold = Math.min(180000, PPT_DRAG_SNAP_THRESHOLD_PX / drag.scale);
    const xSnap = closestPptSnap(
      [
        {
          offset: rawBounds.width / 2,
          value: rawBounds.left + rawBounds.width / 2,
          role: 'center',
        },
        { offset: rawBounds.width, value: rawBounds.left + rawBounds.width, role: 'end' },
      ],
      pptSnapTargetsForAxis(slide, drag.shapeId, 'x'),
      threshold,
      { matchRole: false },
    );
    const ySnap = closestPptSnap(
      [
        {
          offset: rawBounds.height / 2,
          value: rawBounds.top + rawBounds.height / 2,
          role: 'center',
        },
        { offset: rawBounds.height, value: rawBounds.top + rawBounds.height, role: 'end' },
      ],
      pptSnapTargetsForAxis(slide, drag.shapeId, 'y'),
      threshold,
      { matchRole: false },
    );
    const bounds = clampedPptDragBounds(drag.startBounds, drag.slideSize, {
      width: xSnap
        ? snappedResizeLength(rawBounds.left, rawBounds.width, xSnap.source, xSnap.target.value)
        : rawBounds.width,
      height: ySnap
        ? snappedResizeLength(rawBounds.top, rawBounds.height, ySnap.source, ySnap.target.value)
        : rawBounds.height,
    });
    const guides = [
      ...(xSnap
        ? [
            {
              axis: 'x' as const,
              value: xSnap.target.value,
              kind: xSnap.target.kind,
              referenceShapeId: xSnap.target.referenceShapeId,
            },
          ]
        : []),
      ...(ySnap
        ? [
            {
              axis: 'y' as const,
              value: ySnap.target.value,
              kind: ySnap.target.kind,
              referenceShapeId: ySnap.target.referenceShapeId,
            },
          ]
        : []),
    ];
    this.renderSnapGuides(guides, drag.scale);
    return { bounds, guides, snapped: guides.length > 0 };
  }

  private renderSnapGuides(guides: PptSnapGuide[], scale: number): void {
    this.clearSnapReferenceHighlights();
    const referenceShapeIds = new Set(
      guides.map((guide) => guide.referenceShapeId).filter((id): id is string => Boolean(id)),
    );
    for (const shapeId of referenceShapeIds) {
      const node = elementForShape(this.canvas, shapeId);
      if (node) node.dataset.snapReference = 'true';
    }
    this.snapGuides.replaceChildren(
      ...guides.map((guide) => {
        const element = document.createElement('div');
        element.className = 'coco-ppt-snap-guide';
        element.dataset.axis = guide.axis;
        element.dataset.kind = guide.kind;
        if (guide.referenceShapeId) element.dataset.referenceShapeId = guide.referenceShapeId;
        if (guide.axis === 'x') {
          element.style.left = `${guide.value * scale}px`;
        } else {
          element.style.top = `${guide.value * scale}px`;
        }
        return element;
      }),
    );
  }

  private hideSnapGuides(): void {
    this.clearSnapReferenceHighlights();
    this.snapGuides.replaceChildren();
  }

  private clearDragVisualState(element: HTMLElement, force = false): void {
    delete element.dataset.dragging;
    if (force || !this.dragState) delete this.canvas.dataset.interactionMode;
  }

  private restoreDragElementBounds(drag: NativePptDragState): void {
    if (drag.mode === 'move') {
      drag.element.style.left = `${drag.startBounds.left * drag.scale}px`;
      drag.element.style.top = `${drag.startBounds.top * drag.scale}px`;
      return;
    }
    drag.element.style.width = `${drag.startBounds.width * drag.scale}px`;
    drag.element.style.height = `${drag.startBounds.height * drag.scale}px`;
  }

  private clearSnapReferenceHighlights(): void {
    this.canvas
      .querySelectorAll<HTMLElement>('.coco-ppt-node[data-snap-reference="true"]')
      .forEach((node) => {
        delete node.dataset.snapReference;
      });
  }

  private handleKeyDown(event: KeyboardEvent): void {
    if (
      event.defaultPrevented ||
      this.dragState ||
      isTypingTarget(event.target) ||
      Boolean(this.root.querySelector('[contenteditable="true"]'))
    )
      return;
    const selection = this.selection();
    if (selection?.kind !== 'shape') return;
    const slide = this.sessionValue.scene().slides.find((item) => item.id === selection.slideId);
    const node = slide?.nodes.find((item) => item.id === selection.shapeId);
    if (!slide || !node) return;
    const commandKey = event.ctrlKey || event.metaKey;
    if (commandKey && event.key.toLowerCase() === 'c') {
      event.preventDefault();
      this.copySelectedShape(slide, node);
      return;
    }
    if (commandKey && event.key.toLowerCase() === 'v') {
      event.preventDefault();
      this.pasteClipboardShape(slide, node);
      return;
    }
    if (commandKey && event.key.toLowerCase() === 'd') {
      event.preventDefault();
      this.duplicateSelectedShape(slide, node);
      return;
    }
    if (commandKey && event.shiftKey && event.key === ']') {
      event.preventDefault();
      this.bringSelectedShapeToFront(slide, node);
      return;
    }
    if (commandKey && event.shiftKey && event.key === '[') {
      event.preventDefault();
      this.sendSelectedShapeToBack(slide, node);
      return;
    }
    if (event.key === 'Delete' || event.key === 'Backspace') {
      if (!isShapeCollectionEditable(node)) return;
      event.preventDefault();
      this.deleteSelectedShape(slide, node);
      return;
    }
    const direction = arrowKeyDirection(event.key);
    if (!direction) return;
    if (!node.capabilities.editGeometry) return;
    const step = PPT_KEYBOARD_NUDGE_EMU * (event.shiftKey ? 10 : 1);
    const geometry = clampedPptDragBounds(node.bounds, slide.size, {
      left: node.bounds.left + direction.dx * step,
      top: node.bounds.top + direction.dy * step,
    });
    event.preventDefault();
    void this.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: slide.id,
      shapeIdOrName: node.id,
      geometry: { left: geometry.left, top: geometry.top },
    }).then(() => {
      this.options.onObjectAction?.({
        action: 'nudge-geometry',
        slideId: slide.id,
        shapeId: node.id,
        kind: node.kind,
        name: node.name,
        reason: `已移动对象 ${direction.label}${event.shiftKey ? ' 10步' : ' 1步'}`,
      });
    });
  }

  private handleSelectStart(event: Event): void {
    if (isTypingTarget(event.target)) return;
    if (event.target instanceof HTMLElement && event.target.closest('[contenteditable="true"]'))
      return;
    event.preventDefault();
  }

  private copySelectedShape(slide: PptSlideScene, node: PptSceneNode): void {
    if (!isShapeCollectionEditable(node)) return;
    this.clipboardShape = { slideId: slide.id, shapeId: node.id };
    this.options.onObjectAction?.({
      action: 'copy-shape',
      slideId: slide.id,
      shapeId: node.id,
      kind: node.kind,
      name: node.name,
      reason: '已复制对象',
    });
  }

  private pasteClipboardShape(slide: PptSlideScene, node: PptSceneNode): void {
    if (!this.clipboardShape) return;
    const source = this.clipboardShape;
    void this.dispatch({
      kind: 'paste-shape',
      slideIdOrIndex: source.slideId,
      shapeIdOrName: source.shapeId,
    }).then(() => {
      const selection = this.selection();
      this.options.onObjectAction?.({
        action: 'paste-shape',
        slideId: source.slideId,
        shapeId: selection?.kind === 'shape' ? selection.shapeId : node.id,
        kind: node.kind,
        name: node.name,
        reason: '已粘贴对象',
      });
    });
  }

  private duplicateSelectedShape(slide: PptSlideScene, node: PptSceneNode): void {
    if (!isShapeCollectionEditable(node)) return;
    void this.dispatch({
      kind: 'duplicate-shape',
      slideIdOrIndex: slide.id,
      shapeIdOrName: node.id,
    }).then(() => {
      const selection = this.selection();
      this.options.onObjectAction?.({
        action: 'duplicate-shape',
        slideId: slide.id,
        shapeId: selection?.kind === 'shape' ? selection.shapeId : node.id,
        kind: node.kind,
        name: node.name,
        reason: '已重复对象',
      });
    });
  }

  private bringSelectedShapeToFront(slide: PptSlideScene, node: PptSceneNode): void {
    if (!isShapeCollectionEditable(node)) return;
    void this.dispatch({
      kind: 'bring-shape-to-front',
      slideIdOrIndex: slide.id,
      shapeIdOrName: node.id,
    }).then(() => {
      this.options.onObjectAction?.({
        action: 'bring-shape-to-front',
        slideId: slide.id,
        shapeId: node.id,
        kind: node.kind,
        name: node.name,
        reason: '已置于顶层',
      });
    });
  }

  private sendSelectedShapeToBack(slide: PptSlideScene, node: PptSceneNode): void {
    if (!isShapeCollectionEditable(node)) return;
    void this.dispatch({
      kind: 'send-shape-to-back',
      slideIdOrIndex: slide.id,
      shapeIdOrName: node.id,
    }).then(() => {
      this.options.onObjectAction?.({
        action: 'send-shape-to-back',
        slideId: slide.id,
        shapeId: node.id,
        kind: node.kind,
        name: node.name,
        reason: '已置于底层',
      });
    });
  }

  private showContextMenu(event: MouseEvent, slide: PptSlideScene, node: PptSceneNode): void {
    const editable = isShapeCollectionEditable(node);
    const hasClipboard = Boolean(this.clipboardShape);
    const run = (action: () => void) => () => {
      this.hideContextMenu();
      action();
    };
    this.contextMenu.replaceChildren(
      contextMenuButton(
        '复制',
        'Ctrl+C',
        editable,
        run(() => this.copySelectedShape(slide, node)),
      ),
      contextMenuButton(
        '粘贴',
        'Ctrl+V',
        hasClipboard,
        run(() => this.pasteClipboardShape(slide, node)),
      ),
      contextMenuButton(
        '重复',
        'Ctrl+D',
        editable,
        run(() => this.duplicateSelectedShape(slide, node)),
      ),
      contextMenuDivider(),
      contextMenuButton(
        '置于顶层',
        'Ctrl+Shift+]',
        editable,
        run(() => this.bringSelectedShapeToFront(slide, node)),
      ),
      contextMenuButton(
        '置于底层',
        'Ctrl+Shift+[',
        editable,
        run(() => this.sendSelectedShapeToBack(slide, node)),
      ),
      contextMenuDivider(),
      contextMenuButton(
        '删除',
        'Del',
        editable,
        run(() => this.deleteSelectedShape(slide, node)),
      ),
    );
    const rootRect =
      this.root instanceof ShadowRoot
        ? this.root.host.getBoundingClientRect()
        : this.root.getBoundingClientRect();
    const left = clampNumber(event.clientX, rootRect.left + 4, window.innerWidth - 170);
    const top = clampNumber(event.clientY, rootRect.top + 4, window.innerHeight - 220);
    this.contextMenu.style.left = `${left}px`;
    this.contextMenu.style.top = `${top}px`;
    this.contextMenu.dataset.shapeId = node.id;
    this.contextMenu.dataset.open = 'true';
  }

  private deleteSelectedShape(slide: PptSlideScene, node: PptSceneNode): void {
    if (!isShapeCollectionEditable(node)) return;
    void this.dispatch({
      kind: 'delete-shape',
      slideIdOrIndex: slide.id,
      shapeIdOrName: node.id,
    }).then(() => {
      this.options.onObjectAction?.({
        action: 'delete-shape',
        slideId: slide.id,
        shapeId: node.id,
        kind: node.kind,
        name: node.name,
        reason: '已删除对象',
      });
    });
  }

  private async dispatchCurrentSlideSelection(): Promise<void> {
    if (this.consumeSuppressedClick()) return;
    const slide = this.currentSlide();
    if (!slide) return;
    await this.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id });
  }

  private consumeSuppressedClick(): boolean {
    if (!this.suppressClickAfterDrag) return false;
    this.suppressClickAfterDrag = false;
    return true;
  }

  private dismissContextMenu(event: PointerEvent): void {
    const target = event.composedPath()[0];
    if (target instanceof Node && this.contextMenu.contains(target)) return;
    this.hideContextMenu();
  }

  private hideContextMenu(): void {
    delete this.contextMenu.dataset.open;
    delete this.contextMenu.dataset.shapeId;
  }

  private renderSelection(): void {
    this.canvas
      .querySelectorAll<HTMLElement>('.coco-ppt-parent-group-outline')
      .forEach((node) => node.remove());
    const selection = this.sessionValue.selection();
    this.slideList.querySelectorAll<HTMLElement>('.coco-ppt-thumb').forEach((thumb, index) => {
      const slide = this.sessionValue.scene().slides[index];
      thumb.dataset.selected = selection?.slideId === slide?.id ? 'true' : 'false';
    });
    this.canvas.querySelectorAll<HTMLElement>('.coco-ppt-node').forEach((node) => {
      node.dataset.selected =
        selection?.kind === 'shape' && node.dataset.shapeId === selection.shapeId
          ? 'true'
          : 'false';
    });
    this.renderParentGroupSelection(selection);
  }

  private renderParentGroupSelection(selection: PptEditorSelection): void {
    if (selection?.kind !== 'shape') return;
    const slide = this.sessionValue.scene().slides.find((item) => item.id === selection.slideId);
    const selectedNode =
      slide?.nodes.find((item) => item.id === selection.shapeId) ??
      slide?.groups.find((item) => item.id === selection.shapeId);
    const groupNode = selectedNode?.groupContext
      ? (slide?.groups.find((item) => item.id === selectedNode.groupContext?.parentId) ??
        parentGroupDragNode(selectedNode))
      : selectedNode?.kind === 'group'
        ? selectedNode
        : null;
    if (!slide || !groupNode) return;
    const scale = this.computeScale(slide);
    const outline = document.createElement('div');
    outline.className = 'coco-ppt-parent-group-outline';
    outline.dataset.shapeId = groupNode.id;
    outline.dataset.selected = selection.shapeId === groupNode.id ? 'true' : 'false';
    outline.title = `父组：${groupNode.name || groupNode.id}`;
    outline.style.left = `${groupNode.bounds.left * scale}px`;
    outline.style.top = `${groupNode.bounds.top * scale}px`;
    outline.style.width = `${groupNode.bounds.width * scale}px`;
    outline.style.height = `${groupNode.bounds.height * scale}px`;
    const startParentGroupDrag = (event: MouseEvent) => {
      event.preventDefault();
      event.stopPropagation();
      this.startDrag(event, 'move', slide, groupNode, scale, outline);
    };
    (['top', 'right', 'bottom', 'left'] as const).forEach((edge) => {
      const edgeElement = document.createElement('div');
      edgeElement.className = 'coco-ppt-parent-group-edge';
      edgeElement.dataset.edge = edge;
      edgeElement.title = '拖动父组';
      edgeElement.addEventListener('mousedown', startParentGroupDrag);
      outline.append(edgeElement);
    });
    const handle = document.createElement('div');
    handle.className = 'coco-ppt-parent-group-handle';
    handle.title = '拖动父组';
    handle.addEventListener('mousedown', startParentGroupDrag);
    outline.append(handle);
    this.canvas.append(outline);
  }

  private currentSlide(): PptSlideScene | null {
    const scene = this.sessionValue.scene();
    const selection = this.sessionValue.selection();
    if (selection?.kind === 'slide' || selection?.kind === 'shape') {
      return (
        scene.slides.find((slide) => slide.id === selection.slideId) ?? scene.slides[0] ?? null
      );
    }
    return scene.slides[0] ?? null;
  }

  private computeScale(slide: PptSlideScene): number {
    const stage = this.canvas.parentElement;
    const availableWidth = stage ? Math.max(1, stage.clientWidth - 48) : 960;
    const availableHeight = stage ? Math.max(1, stage.clientHeight - 48) : 540;
    return Math.min(
      960 / slide.size.width,
      540 / slide.size.height,
      availableWidth / slide.size.width,
      availableHeight / slide.size.height,
    );
  }
}

function prepareTextResizePreview(element: HTMLElement, node: PptSceneNode): void {
  if (!shouldUseWrappingResizePreview(node)) return;
  const target =
    element.querySelector<HTMLElement>('.coco-ppt-rich-text, .coco-ppt-distributed-text') ??
    element.querySelector<HTMLElement>('.coco-ppt-text-layer') ??
    element;
  element.dataset.textResizePreview = 'wrap';
  target.dataset.textResizePreview = 'wrap';
  target.style.transform = '';
  target.style.transformOrigin = '';
  target.style.width = '';
  element.style.whiteSpace = 'pre-wrap';
  element.style.overflowWrap = 'anywhere';
  element
    .querySelectorAll<HTMLElement>('.coco-ppt-rich-text, .coco-ppt-paragraph')
    .forEach((item) => {
      item.style.whiteSpace = 'pre-wrap';
      item.style.overflowWrap = 'anywhere';
    });
}

function shouldUseWrappingResizePreview(node: PptSceneNode): boolean {
  return Boolean(
    node.capabilities.editGeometry &&
    isResizeWrappingTextShape(node) &&
    (node.text || node.textParagraphs?.length),
  );
}

function isResizeWrappingTextShape(node: PptSceneNode): boolean {
  return node.kind === 'placeholder' || node.kind === 'textBox' || node.kind === 'geometricShape';
}

function nextAnimationFrame(): Promise<void> {
  return new Promise((resolve) => requestAnimationFrame(() => resolve()));
}

async function rasterizeThumbnailElement(
  element: HTMLElement,
  width: number,
  height: number,
  scale: number,
  cssText: string,
  layoutWidth = width,
  layoutHeight = height,
): Promise<HTMLCanvasElement | null> {
  if (
    typeof XMLSerializer === 'undefined' ||
    typeof Blob === 'undefined' ||
    typeof URL === 'undefined' ||
    typeof URL.createObjectURL !== 'function'
  )
    return null;
  await waitForThumbnailImages(element);
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(width * scale));
  canvas.height = Math.max(1, Math.round(height * scale));
  const context = canvas.getContext('2d');
  if (!context) return null;
  const wrapper = document.createElement('div');
  wrapper.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
  wrapper.style.width = `${layoutWidth}px`;
  wrapper.style.height = `${layoutHeight}px`;
  wrapper.style.position = 'relative';
  wrapper.style.overflow = 'hidden';
  wrapper.style.background = '#fff';
  const style = document.createElement('style');
  style.textContent = cssText;
  const clone = element.cloneNode(true) as HTMLElement;
  await inlineThumbnailImages(clone);
  clone.style.position = 'relative';
  clone.style.left = '0';
  clone.style.top = '0';
  clone.style.transform = 'none';
  wrapper.append(style, clone);
  const serialized = new XMLSerializer().serializeToString(wrapper);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${canvas.width}" height="${canvas.height}" viewBox="0 0 ${layoutWidth} ${layoutHeight}"><foreignObject width="100%" height="100%">${serialized}</foreignObject></svg>`;
  const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml;charset=utf-8' }));
  try {
    const image = await loadRasterImage(url);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return canvas;
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}

function cloneThumbnailCanvas(source: HTMLCanvasElement): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.className = 'coco-ppt-thumb-canvas';
  canvas.width = source.width;
  canvas.height = source.height;
  const context = canvas.getContext('2d');
  if (context) context.drawImage(source, 0, 0);
  return canvas;
}

function thumbnailFingerprint(slide: PptSlideScene): string {
  return JSON.stringify({
    size: slide.size,
    background: slide.background,
    nodes: slide.nodes.map((node) => thumbnailNodeFingerprint(node)),
  });
}

function thumbnailNodeFingerprint(node: PptSceneNode): unknown {
  return {
    id: node.id,
    kind: node.kind,
    bounds: node.bounds,
    text: node.text,
    textParagraphs: node.textParagraphs,
    textStyle: node.textStyle,
    textBody: node.textBody,
    table: node.table,
    fillColor: node.fillColor,
    fillGradient: node.fillGradient,
    opacity: node.opacity,
    lineColor: node.lineColor,
    lineGradient: node.lineGradient,
    lineOpacity: node.lineOpacity,
    lineWidth: node.lineWidth,
    lineDash: node.lineDash,
    presetGeometry: node.presetGeometry,
    presetAdjustments: node.presetAdjustments,
    customGeometry: node.customGeometry,
    imagePath: node.imagePath,
    imageCrop: node.imageCrop,
    imageFillRect: node.imageFillRect,
    imageOpacity: node.imageOpacity,
    shadow: node.shadow,
    threeD: node.threeD,
    children: node.children.map((child) => thumbnailNodeFingerprint(child)),
  };
}

async function waitForThumbnailImages(element: HTMLElement): Promise<void> {
  const images = Array.from(element.querySelectorAll<HTMLImageElement>('img'));
  await Promise.all(
    images.map(async (image) => {
      if (image.complete && image.naturalWidth > 0) return;
      if (typeof image.decode === 'function') {
        await image.decode().catch(() => {});
        return;
      }
      await new Promise<void>((resolve) => {
        const done = () => resolve();
        image.addEventListener('load', done, { once: true });
        image.addEventListener('error', done, { once: true });
        setTimeout(done, 500);
      });
    }),
  );
}

async function inlineThumbnailImages(element: HTMLElement): Promise<void> {
  const images = Array.from(element.querySelectorAll<HTMLImageElement>('img'));
  await Promise.all(
    images.map(async (image) => {
      image.alt = '';
      const src = image.getAttribute('src');
      if (!src || src.startsWith('data:')) return;
      const dataUrl = await imageSrcToDataUrl(src).catch(() => undefined);
      if (dataUrl) image.setAttribute('src', dataUrl);
    }),
  );
}

async function imageSrcToDataUrl(src: string): Promise<string> {
  const response = await fetch(src);
  if (!response.ok) throw new Error(`Failed to inline thumbnail image: ${response.status}`);
  const blob = await response.blob();
  return blobToDataUrl(blob);
}

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') resolve(reader.result);
      else reject(new Error('Failed to read thumbnail image blob.'));
    };
    reader.onerror = () =>
      reject(reader.error ?? new Error('Failed to read thumbnail image blob.'));
    reader.readAsDataURL(blob);
  });
}

function loadRasterImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Failed to rasterize PPT thumbnail.'));
    image.src = url;
  });
}

function fontFaceCss(fonts: PptEditorView['fonts']): string {
  const faces = fonts.filter(isBrowserUsableFontResource).map((font) => {
    const family = cssString(font.family);
    const weight = font.weight ?? 400;
    const style = font.style ?? 'normal';
    return `@font-face { font-family: ${family}; src: url(${cssUrl(font.url)}); font-weight: ${weight}; font-style: ${style}; }`;
  });
  return faces.join('\n');
}

function isBrowserUsableFontResource(font: PptEditorView['fonts'][number]): boolean {
  return /font\/(ttf|otf|woff|woff2)|application\/font-(sfnt|woff|woff2)/i.test(font.contentType);
}

function ensureGoogleFontLinks(
  fonts: PptEditorView['fonts'],
  referencedFamilies: Iterable<string> = [],
): Promise<void> {
  if (typeof document === 'undefined') return Promise.resolve();
  const families = new Set(fonts.map((font) => font.family));
  for (const family of referencedFamilies) families.add(family);
  const pending: Promise<void>[] = [];
  for (const family of families) {
    const url = googleFontCssUrl(family);
    if (!url) continue;
    const id = `coco-ppt-google-font-${family.replace(/[^a-zA-Z0-9_-]/g, '_')}`;
    const existing = document.getElementById(id) as HTMLLinkElement | null;
    if (existing) {
      pending.push(waitForStylesheet(existing));
      continue;
    }
    const link = document.createElement('link');
    link.id = id;
    link.rel = 'stylesheet';
    link.href = url;
    pending.push(waitForStylesheet(link));
    document.head.append(link);
  }
  return Promise.all(pending).then(() => {});
}

function usedFontFamiliesInView(view: PptEditorView): string[] {
  const families = new Set<string>();
  const add = (family: string | undefined) => {
    if (family?.trim()) families.add(family.trim());
  };
  const visitNode = (node: PptSceneNode) => {
    add(node.textStyle?.fontFamily);
    for (const paragraph of node.textParagraphs ?? []) {
      add(paragraph.style?.bulletFontFamily);
      for (const run of paragraph.runs ?? []) add(run.style?.fontFamily);
    }
    for (const row of node.table?.cellStyles ?? []) {
      for (const cell of row) add(cell?.textStyle?.fontFamily);
    }
    for (const child of node.children) visitNode(child);
  };
  for (const slide of view.scene.slides) {
    for (const node of slide.nodes) visitNode(node);
  }
  return [...families];
}

function exposeNativePptFontReady(ready: Promise<void>): void {
  (globalThis as unknown as { __cocoPptFontRenderReady?: Promise<void> }).__cocoPptFontRenderReady =
    ready;
}

function waitForStylesheet(link: HTMLLinkElement): Promise<void> {
  if (link.sheet) return Promise.resolve();
  return new Promise((resolve) => {
    const finish = () => resolve();
    link.addEventListener('load', finish, { once: true });
    link.addEventListener('error', finish, { once: true });
    window.setTimeout(finish, 5000);
  });
}

function editableTextTarget(element: HTMLElement): HTMLElement {
  return element.querySelector<HTMLElement>('.coco-ppt-text-layer') ?? element;
}

function tableCellFromEvent(event: Event): HTMLElement | null {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return null;
  return target.closest<HTMLElement>('td[data-row][data-column]');
}

function contextMenuButton(
  label: string,
  shortcut: string,
  enabled: boolean,
  action: () => void,
): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.disabled = !enabled;
  const text = document.createElement('strong');
  text.textContent = label;
  const keys = document.createElement('span');
  keys.textContent = shortcut;
  button.append(text, keys);
  button.addEventListener('click', () => {
    if (button.disabled) return;
    action();
  });
  return button;
}

function contextMenuDivider(): HTMLHRElement {
  return document.createElement('hr');
}

function arrowKeyDirection(key: string): { dx: number; dy: number; label: string } | null {
  if (key === 'ArrowLeft') return { dx: -1, dy: 0, label: '左' };
  if (key === 'ArrowRight') return { dx: 1, dy: 0, label: '右' };
  if (key === 'ArrowUp') return { dx: 0, dy: -1, label: '上' };
  if (key === 'ArrowDown') return { dx: 0, dy: 1, label: '下' };
  return null;
}

function isTypingTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (target.isContentEditable) return true;
  const tagName = target.tagName.toLowerCase();
  return tagName === 'input' || tagName === 'textarea' || tagName === 'select';
}

function isEditingText(element: HTMLElement, target: EventTarget | null): boolean {
  if (element.isContentEditable) return true;
  if (target instanceof HTMLElement && target.isContentEditable) return true;
  return Boolean(element.querySelector<HTMLElement>('[contenteditable="true"]'));
}

function inheritedNodeTextStyle(
  node: PptSceneNode,
  style: NonNullable<PptSceneNode['textStyle']>,
): NonNullable<PptSceneNode['textStyle']> {
  const paragraphs = node.textParagraphs?.filter((paragraph) => paragraph.runs.length > 0);
  if (!paragraphs?.length || !shouldRenderRichText(node, paragraphs)) return style;
  const { textGradient, textFill, textOutlineColor, ...inherited } = style;
  void textGradient;
  void textFill;
  void textOutlineColor;
  return inherited;
}

function redirectedEditableTextHit(
  slide: PptSlideScene,
  current: PptSceneNode,
  event: MouseEvent,
  scale: number,
): PptSceneNode | null {
  if (current.capabilities.editText || isInteractiveNonTextNode(current)) return null;
  const point = slidePointFromMouseEvent(event, scale);
  if (!point) return null;
  const currentIndex = slide.nodes.findIndex((node) => node.id === current.id);
  const candidates = slide.nodes.slice(0, currentIndex < 0 ? slide.nodes.length : currentIndex);
  for (let index = candidates.length - 1; index >= 0; index--) {
    const candidate = candidates[index]!;
    if (!candidate.capabilities.editText) continue;
    if (pointInPptRect(point.x, point.y, candidate.bounds)) return candidate;
  }
  return null;
}

function cycledSelectableHit(
  slide: PptSlideScene,
  event: MouseEvent,
  scale: number,
  currentShapeId: string | undefined,
): { node: PptSceneNode; index: number; count: number } | null {
  const candidates = selectableHitsAtEvent(slide, event, scale);
  if (candidates.length < 2) return null;
  const currentIndex = Math.max(
    0,
    candidates.findIndex((candidate) => candidate.id === currentShapeId),
  );
  const index = (currentIndex + 1) % candidates.length;
  return { node: candidates[index]!, index, count: candidates.length };
}

function selectableHitsAtEvent(
  slide: PptSlideScene,
  event: MouseEvent,
  scale: number,
): PptSceneNode[] {
  const point = slidePointFromMouseEvent(event, scale);
  if (!point) return [];
  const hits: PptSceneNode[] = [];
  for (let index = slide.nodes.length - 1; index >= 0; index--) {
    const node = slide.nodes[index]!;
    if (!node.capabilities.select) continue;
    if (pointInPptRect(point.x, point.y, node.bounds)) hits.push(node);
  }
  return hits;
}

function isInteractiveNonTextNode(node: PptSceneNode): boolean {
  return node.kind === 'table' || node.kind === 'chart' || node.kind === 'smartArt';
}

function slidePointFromMouseEvent(
  event: MouseEvent,
  scale: number,
): { x: number; y: number } | null {
  const canvas =
    event.currentTarget &&
    typeof (event.currentTarget as { parentElement?: unknown }).parentElement === 'object'
      ? ((event.currentTarget as { parentElement?: HTMLElement | null }).parentElement ?? null)
      : null;
  const rect = canvas?.getBoundingClientRect();
  if (!rect) return null;
  return {
    x: (event.clientX - rect.left) / scale,
    y: (event.clientY - rect.top) / scale,
  };
}

function pointInPptRect(x: number, y: number, rect: PptSceneNode['bounds']): boolean {
  return (
    x >= rect.left && x <= rect.left + rect.width && y >= rect.top && y <= rect.top + rect.height
  );
}

function elementForShape(canvas: HTMLElement, shapeId: string): HTMLElement | null {
  for (const element of Array.from(canvas.querySelectorAll<HTMLElement>('.coco-ppt-node'))) {
    if (element.dataset.shapeId === shapeId) return element;
  }
  return null;
}

function insertPlainText(target: HTMLElement, text: string): void {
  const selection = target.ownerDocument.getSelection();
  if (!selection || selection.rangeCount === 0) {
    target.textContent = `${target.textContent ?? ''}${text}`;
    return;
  }
  const range = selection.getRangeAt(0);
  const selectedInsideTarget =
    target === range.commonAncestorContainer || target.contains(range.commonAncestorContainer);
  if (!selectedInsideTarget) {
    target.textContent = `${target.textContent ?? ''}${text}`;
    return;
  }
  range.deleteContents();
  const textNode = target.ownerDocument.createTextNode(text);
  range.insertNode(textNode);
  range.setStartAfter(textNode);
  range.setEndAfter(textNode);
  selection.removeAllRanges();
  selection.addRange(range);
}

function nonTextEditReason(node: PptSceneNode): string {
  if (node.capabilities.reasons?.editText) return node.capabilities.reasons.editText;
  if (node.capabilities.reason) return node.capabilities.reason;
  if (node.kind === 'table') return '表格请使用单元格编辑';
  if (node.kind === 'picture') return '图片对象不支持文字编辑';
  if (node.kind === 'chart') return '图表当前仅支持选择和位置编辑';
  if (node.kind === 'smartArt') return 'SmartArt 当前仅支持选择和位置编辑';
  if (node.kind === 'group') return '组合对象当前仅支持选择和位置编辑';
  return '该对象不支持文字编辑';
}

function googleFontCssUrl(family: string): string | undefined {
  const urls: Record<string, string> = {
    Cuprum:
      'https://fonts.googleapis.com/css2?family=Cuprum:ital,wght@0,400;0,700;1,400;1,700&display=swap',
    Podkova: 'https://fonts.googleapis.com/css2?family=Podkova:wght@400;700&display=swap',
    Ubuntu:
      'https://fonts.googleapis.com/css2?family=Ubuntu:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  };
  return urls[family];
}

function cssFontFamilyStack(family: string): string {
  return [cssString(family), '"Calibri"', '"Microsoft YaHei"', '"Segoe UI"', 'sans-serif'].join(
    ', ',
  );
}

function cssString(value: string): string {
  return `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function cssUrl(value: string): string {
  return `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function cssColorWithOpacity(color: string, opacity: number | undefined): string {
  if (opacity === undefined || opacity >= 1) return color;
  const alpha = Math.max(0, Math.min(1, opacity));
  const hex = color.match(/^#([0-9a-fA-F]{6})$/)?.[1];
  if (!hex) return color;
  const red = Number.parseInt(hex.slice(0, 2), 16);
  const green = Number.parseInt(hex.slice(2, 4), 16);
  const blue = Number.parseInt(hex.slice(4, 6), 16);
  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

function effectiveFillOpacity(node: PptSceneNode): number | undefined {
  if (node.opacity !== undefined) return Math.max(0, Math.min(1, node.opacity));
  return node.threeD?.material === 'translucentPowder' ? 0.52 : undefined;
}

function applyImagePlacement(image: HTMLImageElement, node: PptSceneNode): void {
  if (node.imageFillRect) {
    applyImageFillRect(image, node);
    return;
  }
  applyImageCrop(image, node);
}

function applyImageCrop(image: HTMLImageElement, node: PptSceneNode): void {
  const crop = node.imageCrop;
  if (!crop) return;
  const left = crop.left ?? 0;
  const right = crop.right ?? 0;
  const top = crop.top ?? 0;
  const bottom = crop.bottom ?? 0;
  const visibleWidth = Math.max(0.01, 1 - left - right);
  const visibleHeight = Math.max(0.01, 1 - top - bottom);
  image.style.position = 'absolute';
  image.style.left = `${(-left / visibleWidth) * 100}%`;
  image.style.top = `${(-top / visibleHeight) * 100}%`;
  image.style.width = `${(1 / visibleWidth) * 100}%`;
  image.style.height = `${(1 / visibleHeight) * 100}%`;
}

function applyImageFillRect(image: HTMLImageElement, node: PptSceneNode): void {
  const fillRect = node.imageFillRect;
  if (!fillRect) return;
  const left = fillRect.left ?? 0;
  const right = fillRect.right ?? 0;
  const top = fillRect.top ?? 0;
  const bottom = fillRect.bottom ?? 0;
  image.style.position = 'absolute';
  image.style.left = `${left * 100}%`;
  image.style.top = `${top * 100}%`;
  image.style.width = `${Math.max(0.01, 1 - left - right) * 100}%`;
  image.style.height = `${Math.max(0.01, 1 - top - bottom) * 100}%`;
}

function renderUnsupportedImagePlaceholder(
  node: PptSceneNode,
  contentType: string | undefined,
): HTMLDivElement {
  const placeholder = document.createElement('div');
  placeholder.className = 'coco-ppt-unsupported-image';
  const extension = node.imagePath?.split('.').pop()?.toUpperCase();
  placeholder.textContent = extension || contentTypeLabel(contentType) || 'IMAGE';
  placeholder.title = node.imagePath
    ? `Unsupported image: ${node.imagePath}`
    : 'Missing image resource';
  return placeholder;
}

function contentTypeLabel(contentType: string | undefined): string | undefined {
  if (!contentType) return undefined;
  if (contentType.includes('emf')) return 'EMF';
  if (contentType.includes('wmf')) return 'WMF';
  if (contentType.includes('svg')) return 'SVG';
  return 'IMAGE';
}

function appendTextContent(
  container: HTMLElement,
  node: PptSceneNode,
  scale: number,
  fallbackText: string,
): void {
  const paragraphs = node.textParagraphs?.filter((paragraph) => paragraph.runs.length > 0);
  const verticalText = isVerticalTextBody(node.textBody);
  if (!paragraphs?.length || !shouldRenderRichText(node, paragraphs)) {
    if (
      !verticalText &&
      node.textStyle?.alignment === 'distributed' &&
      Array.from(fallbackText).length > 1
    ) {
      appendDistributedFallbackText(container, fallbackText);
    } else {
      container.textContent = fallbackText;
    }
    return;
  }
  const wrapper = document.createElement('div');
  wrapper.className = 'coco-ppt-rich-text';
  paragraphs.forEach((paragraph) => {
    const paragraphElement = document.createElement('div');
    paragraphElement.className = 'coco-ppt-paragraph';
    if (paragraph.style?.alignment) applyTextAlignment(paragraphElement, paragraph.style.alignment);
    applyParagraphStyle(paragraphElement, paragraph.style, scale, { verticalText });
    const distributeCharacters =
      !verticalText && paragraph.style?.alignment === 'distributed' && !paragraph.style?.bullet;
    if (distributeCharacters) {
      paragraphElement.style.display = 'flex';
      paragraphElement.style.justifyContent = 'space-between';
      paragraphElement.style.alignItems = 'baseline';
      paragraphElement.style.width = '100%';
    }
    if (paragraph.style?.bullet) {
      paragraphElement.dataset.bullet = 'true';
      const bullet = document.createElement('span');
      bullet.className = 'coco-ppt-bullet';
      bullet.textContent = paragraph.style.bullet;
      if (paragraph.style.bulletColor) bullet.style.color = paragraph.style.bulletColor;
      applyBulletGlyphStyle(bullet, paragraph.style, scale, paragraph.text);
      paragraphElement.append(bullet);
    }
    paragraph.runs.forEach((run, runIndex) => {
      const texts = distributeCharacters ? Array.from(run.text) : [run.text];
      texts.forEach((text) => {
        const span = document.createElement('span');
        span.textContent = text;
        const runStyle = effectiveRichTextRunStyle(node, paragraph, run, runIndex);
        if (runStyle) applyTextStyleToElement(span, runStyle, scale);
        paragraphElement.append(span);
      });
    });
    wrapper.append(paragraphElement);
  });
  container.append(wrapper);
}

function isVerticalTextBody(body: PptSceneNode['textBody']): boolean {
  return body?.vertical === 'eastAsian' || body?.vertical === 'vertical';
}

function applyTextWrapping(container: HTMLElement, node: PptSceneNode): void {
  if (node.textBody?.wrap === 'none') setTextAutoFitNoWrap(container);
}

function applyTextAutoFit(container: HTMLElement, node: PptSceneNode): void {
  if (node.textBody?.autoFit !== 'normal') return;
  if (!node.text && !node.textParagraphs?.length) return;
  if (isSingleLineAutoFitText(node) && !shouldAllowSingleParagraphAutoFitWrap(node)) {
    setTextAutoFitNoWrap(container);
  }
  requestAnimationFrame(() => {
    if (container.dataset.textResizePreview === 'wrap') return;
    const availableWidth = container.clientWidth;
    const availableHeight = container.clientHeight;
    if (availableWidth <= 0 || availableHeight <= 0) return;
    const contentWidth = container.scrollWidth;
    const contentHeight = container.scrollHeight;
    const widthRatio = contentWidth > availableWidth ? availableWidth / contentWidth : 1;
    const heightRatio = contentHeight > availableHeight ? availableHeight / contentHeight : 1;
    const ratio = Math.max(0.25, Math.min(1, widthRatio, heightRatio));
    if (ratio >= 0.995) return;
    const target =
      container.querySelector<HTMLElement>('.coco-ppt-rich-text, .coco-ppt-distributed-text') ??
      container;
    if (target.dataset.textResizePreview === 'wrap') return;
    target.style.transformOrigin = textAutoFitOrigin(node);
    target.style.transform = `scale(${ratio})`;
    target.style.width = `${100 / ratio}%`;
  });
}

function shouldCounterFlipText(node: PptSceneNode): boolean {
  return Boolean(
    (node.text || node.textParagraphs?.length) &&
    (node.bounds.flipHorizontal || node.bounds.flipVertical),
  );
}

function applyTextCounterFlip(element: HTMLElement, node: PptSceneNode): void {
  const transforms: string[] = [];
  if (node.bounds.flipHorizontal) transforms.push('scaleX(-1)');
  if (node.bounds.flipVertical) transforms.push('scaleY(-1)');
  if (transforms.length === 0) return;
  element.style.transform = transforms.join(' ');
  element.style.transformOrigin = 'center';
}

function setTextAutoFitNoWrap(container: HTMLElement): void {
  container.style.whiteSpace = 'nowrap';
  container
    .querySelectorAll<HTMLElement>('.coco-ppt-rich-text, .coco-ppt-paragraph')
    .forEach((element) => {
      element.style.whiteSpace = 'nowrap';
    });
}

function applyTextMetricFallback(element: HTMLElement, node: PptSceneNode, scale: number): void {
  if (!shouldWidenLatinDisplayTitle(node, scale)) return;
  const fontPx = (node.textStyle?.fontSize ?? firstTextRunFontSize(node) ?? 28) * 12700 * scale;
  element.style.letterSpacing = `${Math.max(1, Math.min(2.4, fontPx * 0.035))}px`;
}

function textBaselineMetricOffset(node: PptSceneNode, scale: number): number {
  if (!shouldLowerCuprumDisplayTitle(node)) return 0;
  const fontSize = node.textStyle?.fontSize ?? firstTextRunFontSize(node);
  if (!fontSize) return 0;
  return fontSize * 12700 * scale * 0.18;
}

function shouldLowerCuprumDisplayTitle(node: PptSceneNode): boolean {
  if (node.kind !== 'textBox' && node.kind !== 'placeholder') return false;
  if (node.textBody?.anchor !== 'top' || node.textBody.wrap !== 'square') return false;
  if (node.fillColor || node.fillGradient || node.lineColor || node.lineGradient) return false;
  const family = node.textStyle?.fontFamily ?? firstTextRunFontFamily(node);
  if (!family || !/\bCuprum\b/i.test(family)) return false;
  const fontSize = node.textStyle?.fontSize ?? firstTextRunFontSize(node);
  if (!fontSize || fontSize < 72) return false;
  const paragraphs = node.textParagraphs?.filter((paragraph) => {
    const text = paragraph.text ?? paragraph.runs.map((run) => run.text).join('');
    return text.trim().length > 0;
  });
  if (paragraphs && paragraphs.length !== 1) return false;
  const text = (paragraphs?.[0]?.text ?? node.text ?? '').trim();
  return text.length >= 12 && !/[\u3400-\u9fff]/.test(text);
}

function shouldWidenLatinDisplayTitle(node: PptSceneNode, scale: number): boolean {
  if (node.textBody?.wrap !== 'square') return false;
  if (hasExplicitCharacterSpacing(node)) return false;
  const alignment = node.textStyle?.alignment ?? node.textParagraphs?.[0]?.style?.alignment;
  if (alignment !== 'center') return false;
  const family = node.textStyle?.fontFamily ?? firstTextRunFontFamily(node);
  if (!family || !/\bJosefin Sans\b/i.test(family)) return false;
  const text = (node.textParagraphs?.[0]?.text ?? node.text ?? '').replace(/\s+/g, ' ').trim();
  if (text.length < 24 || /[a-z\u3400-\u9fff]/.test(text)) return false;
  if (!/^[A-Z0-9][A-Z0-9\s&:;,.!?'"()/+-]*$/.test(text)) return false;
  const fontSize = node.textStyle?.fontSize ?? firstTextRunFontSize(node);
  if (!fontSize) return false;
  const lineHeight =
    node.textStyle?.lineHeight ?? node.textParagraphs?.[0]?.style?.lineHeight ?? 1.2;
  const lineHeightPx = fontSize * 12700 * scale * lineHeight;
  return node.bounds.height * scale >= lineHeightPx * 1.75;
}

function hasExplicitCharacterSpacing(node: PptSceneNode): boolean {
  return Boolean(
    node.textStyle?.characterSpacing !== undefined ||
    node.textParagraphs?.some((paragraph) =>
      paragraph.runs.some((run) => run.style?.characterSpacing !== undefined),
    ),
  );
}

function firstTextRunFontSize(node: PptSceneNode): number | undefined {
  for (const paragraph of node.textParagraphs ?? []) {
    for (const run of paragraph.runs) {
      if (run.style?.fontSize !== undefined) return run.style.fontSize;
    }
  }
  return undefined;
}

function firstTextRunFontFamily(node: PptSceneNode): string | undefined {
  for (const paragraph of node.textParagraphs ?? []) {
    for (const run of paragraph.runs) {
      if (run.style?.fontFamily) return run.style.fontFamily;
    }
  }
  return undefined;
}

function isSingleLineAutoFitText(node: PptSceneNode): boolean {
  const paragraphs = node.textParagraphs?.filter((paragraph) => paragraph.runs.length > 0);
  if (paragraphs && paragraphs.length !== 1) return false;
  if (paragraphs?.[0]?.style?.bullet) return false;
  const text = paragraphs?.[0]?.text ?? node.text ?? '';
  return text.length > 0 && !text.includes('\n') && !text.includes('\r');
}

function shouldAllowCompactLabelWrap(node: PptSceneNode): boolean {
  if (node.textBody?.wrap !== 'square') return false;
  if ((node.textStyle?.alignment ?? node.textParagraphs?.[0]?.style?.alignment) !== 'center')
    return false;
  const text = (node.textParagraphs?.[0]?.text ?? node.text ?? '').trim();
  if (!/^[\u3400-\u9fff]{4,6}$/.test(text)) return false;
  const min = Math.min(node.bounds.width, node.bounds.height);
  const max = Math.max(node.bounds.width, node.bounds.height);
  return min > 0 && max <= 1100000 && max / min < 1.25;
}

function shouldAllowSingleParagraphAutoFitWrap(node: PptSceneNode): boolean {
  if (shouldAllowCompactLabelWrap(node)) return true;
  if (node.textBody?.wrap !== 'square') return false;
  const text = (node.textParagraphs?.[0]?.text ?? node.text ?? '').trim();
  if (text.length < 18) return false;
  const cjkCharacters = text.match(/[\u3400-\u9fff]/g)?.length ?? 0;
  if (cjkCharacters >= 12) return true;
  return /[\s，。；、,.]/.test(text);
}

function textAutoFitOrigin(node: PptSceneNode): string {
  const alignment = node.textStyle?.alignment ?? node.textParagraphs?.[0]?.style?.alignment;
  if (alignment === 'right') return 'right top';
  if (alignment === 'center') return 'center top';
  return 'top left';
}

function shouldAllowTextOverflow(node: PptSceneNode): boolean {
  if (node.kind !== 'textBox') return false;
  if (!node.text && !node.textParagraphs?.length) return false;
  if (node.customGeometry || hasSvgPresetGeometry(node) || isLineNode(node)) return false;
  return !hasVisibleFill(node) && !hasVisibleLine(node);
}

function hasVisibleFill(node: PptSceneNode): boolean {
  if (node.fillGradient?.colors.length) {
    return (
      (node.fillGradient.opacities ?? []).some((opacity) => opacity > 0) ||
      !node.fillGradient.opacities
    );
  }
  return Boolean(node.fillColor && node.opacity !== 0);
}

function hasVisibleLine(node: PptSceneNode): boolean {
  if (!node.lineColor && !node.lineGradient?.colors.length) return false;
  if (node.lineOpacity === 0) return false;
  return (node.lineWidth ?? 9525) > 0;
}

function appendDistributedFallbackText(container: HTMLElement, text: string): void {
  const wrapper = document.createElement('div');
  wrapper.className = 'coco-ppt-distributed-text';
  Array.from(text).forEach((char) => {
    const span = document.createElement('span');
    span.textContent = char;
    wrapper.append(span);
  });
  container.append(wrapper);
}

function shouldRenderRichText(
  node: PptSceneNode,
  paragraphs: NonNullable<PptSceneNode['textParagraphs']>,
): boolean {
  return paragraphs.some(
    (paragraph) =>
      Boolean(paragraph.style?.bullet) ||
      hasParagraphLayoutStyle(paragraph.style) ||
      paragraph.runs.some(
        (run) =>
          Boolean(run.hyperlink) ||
          (run.style !== undefined && isMeaningfulRunStyle(run.style, node.textStyle)),
      ),
  );
}

function hasParagraphLayoutStyle(style: PptParagraphStyle): boolean {
  return Boolean(
    style?.lineHeight !== undefined ||
    style?.marginLeft !== undefined ||
    style?.textIndent !== undefined ||
    style?.spaceBefore !== undefined ||
    style?.spaceAfter !== undefined,
  );
}

function applyParagraphStyle(
  element: HTMLElement,
  style: PptParagraphStyle,
  scale: number,
  options: { verticalText?: boolean } = {},
): void {
  if (!style) return;
  if (style.lineHeight !== undefined)
    element.style.lineHeight = String(pptParagraphLineHeight(style));
  if (style.marginLeft !== undefined) {
    element.style.paddingLeft = `${style.marginLeft * scale}px`;
  }
  if (style.textIndent !== undefined) element.style.textIndent = `${style.textIndent * scale}px`;
  if (options.verticalText) {
    if (style.spaceBefore !== undefined)
      element.style.marginRight = paragraphSpacingToCss(style.spaceBefore, scale);
    if (style.spaceAfter !== undefined)
      element.style.marginLeft = paragraphSpacingToCss(style.spaceAfter, scale);
    return;
  }
  if (style.spaceBefore !== undefined)
    element.style.marginTop = paragraphSpacingToCss(style.spaceBefore, scale);
  if (style.spaceAfter !== undefined)
    element.style.marginBottom = paragraphSpacingToCss(style.spaceAfter, scale);
}

function pptParagraphLineHeight(style: NonNullable<PptParagraphStyle>): number {
  const lineHeight = style.lineHeight ?? 1.2;
  if (!style.bullet) return lineHeight;
  return lineHeight * 1.15;
}

function applyBulletGlyphStyle(
  element: HTMLElement,
  style: PptParagraphStyle,
  scale: number,
  paragraphText = '',
): void {
  if (style?.bulletSize !== undefined) {
    if (style.bulletSize > 4) {
      element.style.fontSize = `${style.bulletSize * 12700 * scale}px`;
    } else if (style.bulletSize > 0) {
      element.style.fontSize = `${style.bulletSize}em`;
    }
  }
  if (!shouldRenderCompactWingdingsSquareBullet(style, paragraphText)) return;
  element.style.width = '0.5em';
  element.style.marginRight = '0.12em';
  element.style.transform = 'scale(0.62)';
  element.style.transformOrigin = 'center';
  element.style.verticalAlign = 'middle';
}

function shouldRenderCompactWingdingsSquareBullet(
  style: PptParagraphStyle,
  paragraphText = '',
): boolean {
  if (!style) return false;
  const font = style.bulletFontFamily?.toLowerCase() ?? '';
  const compactLabelLength = Array.from(paragraphText.trim()).length;
  return (
    style.bullet === '□' &&
    style.alignment === 'center' &&
    compactLabelLength > 0 &&
    compactLabelLength <= 4 &&
    (font.includes('wingdings') || font.includes('webdings')) &&
    style.textIndent !== undefined &&
    style.textIndent < 0 &&
    style.marginLeft !== undefined &&
    style.marginLeft <= Math.abs(style.textIndent) * 1.25
  );
}

function paragraphSpacingToCss(value: number, scale: number): string {
  return value > 10 ? `${value * scale}px` : `${value}em`;
}

function isMeaningfulRunStyle(
  runStyle: NonNullable<
    NonNullable<PptSceneNode['textParagraphs']>[number]['runs'][number]['style']
  >,
  nodeStyle: PptSceneNode['textStyle'],
): boolean {
  const nativeRunStyle = runStyle as NativePptTextStyle;
  const nativeNodeStyle = nodeStyle as NativePptTextStyle | undefined;
  if (runStyle.fontSize !== undefined && runStyle.fontSize !== nodeStyle?.fontSize) return true;
  if (runStyle.fontFamily !== undefined && runStyle.fontFamily !== nodeStyle?.fontFamily)
    return true;
  if (runStyle.color !== undefined && runStyle.color !== nodeStyle?.color) return true;
  if (
    nativeRunStyle.textFill !== undefined &&
    nativeRunStyle.textFill !== nativeNodeStyle?.textFill
  )
    return true;
  if (
    nativeRunStyle.textGradient !== undefined &&
    nativeRunStyle.textGradient !== nativeNodeStyle?.textGradient
  )
    return true;
  if (
    nativeRunStyle.textOutlineColor !== undefined &&
    nativeRunStyle.textOutlineColor !== nativeNodeStyle?.textOutlineColor
  )
    return true;
  if (
    nativeRunStyle.textShadow !== undefined &&
    nativeRunStyle.textShadow !== nativeNodeStyle?.textShadow
  )
    return true;
  if (runStyle.bold !== undefined && runStyle.bold !== nodeStyle?.bold) return true;
  if (runStyle.italic !== undefined && runStyle.italic !== nodeStyle?.italic) return true;
  if (runStyle.underline !== undefined && runStyle.underline !== nodeStyle?.underline) return true;
  if (
    runStyle.characterSpacing !== undefined &&
    runStyle.characterSpacing !== nodeStyle?.characterSpacing
  )
    return true;
  return false;
}

function effectiveRichTextRunStyle(
  node: PptSceneNode,
  paragraph: NonNullable<PptSceneNode['textParagraphs']>[number],
  run: NonNullable<PptSceneNode['textParagraphs']>[number]['runs'][number],
  runIndex: number,
): NativePptTextStyle | undefined {
  if (!shouldUseNumberedTitleBodyFallback(node, paragraph, run, runIndex)) {
    return hyperlinkFallbackRunStyle(run, run.style as NativePptTextStyle | undefined);
  }
  const prefixSize = firstExplicitFontSize(paragraph.runs.slice(0, runIndex));
  return hyperlinkFallbackRunStyle(run, {
    ...(run.style as NativePptTextStyle | undefined),
    color: run.style?.color ?? '#000000',
    fontSize: run.style?.fontSize ?? Math.max(12, Math.round((prefixSize ?? 28) * 0.72)),
  });
}

function hyperlinkFallbackRunStyle(
  run: NonNullable<PptSceneNode['textParagraphs']>[number]['runs'][number],
  style: NativePptTextStyle | undefined,
): NativePptTextStyle | undefined {
  if (!run.hyperlink) return style;
  const next: NativePptTextStyle = { ...(style ?? {}) };
  if (next.textFill === undefined && next.textGradient === undefined) {
    next.color = '#0563C1';
  }
  if (next.underline === undefined) next.underline = true;
  return next;
}

function shouldUseNumberedTitleBodyFallback(
  node: PptSceneNode,
  paragraph: NonNullable<PptSceneNode['textParagraphs']>[number],
  run: NonNullable<PptSceneNode['textParagraphs']>[number]['runs'][number],
  runIndex: number,
): boolean {
  if (runIndex <= 0) return false;
  if (run.style?.color && run.style.fontSize !== undefined) return false;
  if (node.textBody?.autoFit !== 'shape') return false;
  if (!/[\u3400-\u9fff]/.test(run.text)) return false;
  const prefixText = paragraph.runs
    .slice(0, runIndex)
    .map((item) => item.text)
    .join('');
  const prefixMatch = /^\s*\d{1,2}\s*[、.．,，]\s*(.*)$/u.exec(prefixText);
  if (!prefixMatch) return false;
  const titleCharsBeforeRun = Array.from(prefixMatch[1] ?? '').filter((char) =>
    /[\u3400-\u9fff]/.test(char),
  ).length;
  if (titleCharsBeforeRun > 2) return false;
  return paragraph.runs.slice(0, runIndex).some((item) => {
    const style = item.style as NativePptTextStyle | undefined;
    return style?.color !== undefined || style?.fontSize !== undefined;
  });
}

function firstExplicitFontSize(
  runs: NonNullable<PptSceneNode['textParagraphs']>[number]['runs'],
): number | undefined {
  for (const run of runs) {
    if (run.style?.fontSize !== undefined) return run.style.fontSize;
  }
  return undefined;
}

function applyTextStyleToElement(
  element: HTMLElement,
  style: NativePptTextStyle,
  scale: number,
): void {
  if (style.fontSize) element.style.fontSize = `${Math.max(1, style.fontSize * 12700 * scale)}px`;
  if (style.fontFamily) element.style.fontFamily = cssFontFamilyStack(style.fontFamily);
  if (style.lineHeight) element.style.lineHeight = String(style.lineHeight);
  if (style.characterSpacing !== undefined)
    element.style.letterSpacing = `${style.characterSpacing * 12700 * scale}px`;
  if (style.color) element.style.color = style.color;
  if (style.textGradient?.colors.length && style.textGradient.colors.length > 1) {
    element.style.backgroundImage = cssLinearGradient(style.textGradient);
    element.style.backgroundClip = 'text';
    element.style.setProperty('-webkit-background-clip', 'text');
    element.style.color = 'transparent';
    element.style.setProperty('-webkit-text-fill-color', 'transparent');
  }
  if (style.textFill === 'none') {
    element.style.color = style.textOutlineColor ?? style.color ?? 'transparent';
    element.style.setProperty('-webkit-text-fill-color', 'transparent');
  }
  if (style.textOutlineColor)
    element.style.setProperty('-webkit-text-stroke', `1px ${style.textOutlineColor}`);
  if (style.textShadow) element.style.textShadow = cssBoxShadow(style.textShadow, scale);
  if (style.bold !== undefined) element.style.fontWeight = style.bold ? '700' : '400';
  if (style.italic !== undefined) element.style.fontStyle = style.italic ? 'italic' : 'normal';
  if (style.underline !== undefined)
    element.style.textDecorationLine = style.underline ? 'underline' : 'none';
}

function isBrowserRenderableImage(
  contentType: string | undefined,
  path: string | undefined,
): boolean {
  const normalized = contentType?.toLowerCase() ?? '';
  if (
    normalized === 'image/png' ||
    normalized === 'image/jpeg' ||
    normalized === 'image/jpg' ||
    normalized === 'image/gif' ||
    normalized === 'image/webp' ||
    normalized === 'image/svg+xml'
  ) {
    return true;
  }
  return Boolean(path && /\.(?:png|jpe?g|gif|webp|svg)$/i.test(path));
}

function applyTextAlignment(element: HTMLElement, alignment: PptTextAlignment): void {
  if (alignment === 'distributed') {
    element.style.textAlign = 'justify';
    element.style.textAlignLast = 'justify';
    return;
  }
  element.style.textAlign = alignment === 'justify' ? 'justify' : alignment;
}

function isLineNode(node: PptSceneNode): boolean {
  if (node.customGeometry) return false;
  const preset = node.presetGeometry;
  const min = Math.min(node.bounds.width, node.bounds.height);
  const max = Math.max(node.bounds.width, node.bounds.height);
  return (
    (preset === 'line' || preset === 'straightConnector1') &&
    max > 0 &&
    (min === 0 || min / max < 0.02)
  );
}

function lineAxisKind(node: PptSceneNode): 'horizontal' | 'vertical' {
  return node.bounds.height > node.bounds.width ? 'vertical' : 'horizontal';
}

function cssBackground(
  background: PptSlideScene['background'],
  imageResources: PptEditorView['images'] = {},
): string {
  if (background?.imagePath) {
    const resource = imageResources[background.imagePath];
    const url = resource?.url;
    if (url && isBrowserRenderableImage(resource.contentType, background.imagePath)) {
      return `url("${url}") center / cover no-repeat`;
    }
  }
  if (background?.fillGradient && background.fillGradient.colors.length > 1) {
    return cssLinearGradient(background.fillGradient);
  }
  return background?.fillColor ?? '#fff';
}

function applyBackgroundFill(
  element: HTMLElement,
  background: PptSlideScene['background'],
  node: PptSceneNode,
  slide: PptSlideScene,
  scale: number,
  imageResources: PptEditorView['images'],
): void {
  if (background?.imagePath) {
    const resource = imageResources[background.imagePath];
    const url = resource?.url;
    if (url && isBrowserRenderableImage(resource.contentType, background.imagePath)) {
      element.style.backgroundImage = `url("${url}")`;
      element.style.backgroundRepeat = 'no-repeat';
      element.style.backgroundSize = `${slide.size.width * scale}px ${slide.size.height * scale}px`;
      element.style.backgroundPosition = `${-node.bounds.left * scale}px ${-node.bounds.top * scale}px`;
      return;
    }
  }
  element.style.background = cssBackground(background, imageResources);
}

function roundRectRadius(node: PptSceneNode): number {
  const adjustment = node.presetAdjustments?.adj;
  const ratio = adjustment === undefined ? 0.08 : Math.max(0, Math.min(0.5, adjustment / 100000));
  return Math.min(node.bounds.width, node.bounds.height) * ratio;
}

function frameBorderWidth(node: PptSceneNode, scale: number): number {
  const rawAdjustment = node.presetAdjustments?.adj ?? node.presetAdjustments?.adj1;
  const ratio =
    rawAdjustment === undefined ? 0.125 : Math.max(0.02, Math.min(0.45, rawAdjustment / 100000));
  return Math.max(1, Math.min(node.bounds.width, node.bounds.height) * ratio * scale);
}

function shouldRenderGradientOutline(node: PptSceneNode): boolean {
  return (
    !node.customGeometry &&
    node.kind !== 'line' &&
    !node.lineColor &&
    !!node.lineGradient &&
    node.lineGradient.colors.length > 1 &&
    (node.presetGeometry === 'roundRect' ||
      node.presetGeometry === 'round1Rect' ||
      node.presetGeometry === 'ellipse')
  );
}

function renderCssPresetGeometry(
  node: PptSceneNode,
  scale: number,
  clipPath: string,
): HTMLDivElement {
  const shape = document.createElement('div');
  shape.className = 'coco-ppt-css-geometry';
  shape.style.clipPath = clipPath;
  if (node.fillGradient && node.fillGradient.colors.length > 1) {
    shape.style.background = cssLinearGradient(node.fillGradient, effectiveFillOpacity(node));
  } else if (node.fillColor) {
    shape.style.background = cssColorWithOpacity(node.fillColor, effectiveFillOpacity(node));
  }
  if (node.lineColor) {
    shape.style.borderStyle = cssBorderStyleForLineDash(node.lineDash);
    shape.style.borderWidth = `${Math.max(1, (node.lineWidth ?? 9525) * scale)}px`;
    shape.style.borderColor = cssColorWithOpacity(node.lineColor, node.lineOpacity);
  }
  return shape;
}

function cssLinearGradient(gradient: PptGradientModel, opacityMultiplier = 1): string {
  const angle = cssLinearGradientAngle(gradient);
  return `linear-gradient(${angle}deg, ${cssGradientStops(gradient, opacityMultiplier).join(', ')})`;
}

function cssLinearGradientAngle(gradient: PptGradientModel): number {
  return gradient.angle === undefined ? 160 : normalizeCssAngle(gradient.angle + 90);
}

function normalizeCssAngle(angle: number): number {
  return ((angle % 360) + 360) % 360;
}

function cssGradientStops(gradient: PptGradientModel, opacityMultiplier = 1): string[] {
  return normalizedGradientStops(gradient).map((stop) => {
    const opacity = Math.max(0, Math.min(1, stop.opacity * opacityMultiplier));
    const color = opacity < 1 ? cssColorWithAlpha(stop.color, opacity) : stop.color;
    return stop.position === undefined ? color : `${color} ${stop.position}%`;
  });
}

function normalizedGradientStops(
  gradient: PptGradientModel,
): { color: string; position: number | undefined; opacity: number }[] {
  return gradient.colors
    .map((color: string, index: number) => {
      const parsedColor = parseCssRgbaColor(color);
      const position = gradient.positions?.[index];
      const opacity = Math.max(
        0,
        Math.min(1, gradient.opacities?.[index] ?? parsedColor?.opacity ?? 1),
      );
      return {
        color: parsedColor?.color ?? color,
        position: position === undefined ? undefined : Math.max(0, Math.min(100, position)),
        opacity,
        index,
      };
    })
    .sort((a, b) => {
      if (a.position === undefined && b.position === undefined) return a.index - b.index;
      if (a.position === undefined) return 1;
      if (b.position === undefined) return -1;
      return a.position - b.position || a.index - b.index;
    })
    .map(({ color, position, opacity }) => ({ color, position, opacity }));
}

function cssColorWithAlpha(color: string, opacity: number): string {
  const hex = color.match(/^#([0-9a-fA-F]{6})$/)?.[1];
  if (!hex) return color;
  const red = Number.parseInt(hex.slice(0, 2), 16);
  const green = Number.parseInt(hex.slice(2, 4), 16);
  const blue = Number.parseInt(hex.slice(4, 6), 16);
  return `rgba(${red}, ${green}, ${blue}, ${Math.max(0, Math.min(1, opacity))})`;
}

function cssClipPathForPresetGeometry(preset: string | undefined): string | undefined {
  switch (preset) {
    case 'parallelogram':
      return 'polygon(18% 0, 100% 0, 82% 100%, 0 100%)';
    case 'trapezoid':
      return 'polygon(18% 0, 82% 0, 100% 100%, 0 100%)';
    case 'snip2DiagRect':
      return 'polygon(8% 0, 100% 0, 100% 92%, 92% 100%, 0 100%, 0 8%)';
    case 'rtTriangle':
      return 'polygon(0 0, 100% 100%, 0 100%)';
    case 'triangle':
      return 'polygon(50% 0, 100% 100%, 0 100%)';
    case 'diamond':
      return 'polygon(50% 0, 100% 50%, 50% 100%, 0 50%)';
    case 'pentagon':
      return 'polygon(0 0, 72% 0, 100% 50%, 72% 100%, 0 100%, 22% 50%)';
    case 'hexagon':
      return 'polygon(25% 0, 75% 0, 100% 50%, 75% 100%, 25% 100%, 0 50%)';
    case 'homePlate':
      return 'polygon(50% 0, 100% 25%, 100% 100%, 0 100%, 0 25%)';
    case 'leftRightArrow':
      return 'polygon(0 50%, 18% 0, 82% 0, 100% 50%, 82% 100%, 18% 100%)';
    case 'chevron':
      return 'polygon(0 0, 72% 0, 100% 50%, 72% 100%, 0 100%, 28% 50%)';
    case 'plaque':
      return 'polygon(0 12%, 3% 12%, 3% 0, 97% 0, 97% 12%, 100% 12%, 100% 88%, 97% 88%, 97% 100%, 3% 100%, 3% 88%, 0 88%)';
    case 'rightArrow':
      return 'polygon(0 0, 78% 0, 100% 50%, 78% 100%, 0 100%)';
    case 'leftArrow':
      return 'polygon(22% 0, 100% 0, 100% 100%, 22% 100%, 0 50%)';
    case 'bentArrow':
      return 'polygon(0 0, 30% 0, 30% 58%, 82% 58%, 82% 35%, 100% 70%, 82% 100%, 82% 77%, 12% 77%, 12% 18%, 0 18%)';
    default:
      return undefined;
  }
}

function cssBorderStyleForLineDash(dash: string | undefined): string {
  if (dash === 'dot' || dash === 'sysDot') return 'dotted';
  if (dash && dash !== 'solid') return 'dashed';
  return 'solid';
}

function svgStrokeDashArray(dash: string | undefined, strokeWidth: number): string | undefined {
  if (!dash || dash === 'solid' || strokeWidth <= 0) return undefined;
  const unit = Math.max(1, strokeWidth);
  const custom = /^cust:(\d+):(\d+)$/i.exec(dash);
  if (custom) {
    const dashUnits = Number(custom[1]) / 100000;
    const spaceUnits = Number(custom[2]) / 100000;
    if (dashUnits > 0 && spaceUnits >= 0) {
      return `${unit * dashUnits} ${unit * spaceUnits}`;
    }
  }
  switch (dash) {
    case 'dot':
    case 'sysDot':
      return `${unit} ${unit * 2}`;
    case 'dash':
    case 'sysDash':
      return `${unit * 4} ${unit * 2}`;
    case 'dashDot':
    case 'sysDashDot':
      return `${unit * 4} ${unit * 2} ${unit} ${unit * 2}`;
    case 'lgDash':
      return `${unit * 8} ${unit * 3}`;
    case 'lgDashDot':
      return `${unit * 8} ${unit * 3} ${unit} ${unit * 3}`;
    case 'lgDashDotDot':
      return `${unit * 8} ${unit * 3} ${unit} ${unit * 3} ${unit} ${unit * 3}`;
    default:
      return `${unit * 4} ${unit * 2}`;
  }
}

function hasSvgPresetGeometry(node: PptSceneNode): boolean {
  return (
    node.presetGeometry === 'curvedUpArrow' ||
    node.presetGeometry === 'donut' ||
    node.presetGeometry === 'arc' ||
    node.presetGeometry === 'blockArc' ||
    node.presetGeometry === 'plaque' ||
    node.presetGeometry === 'parallelogram' ||
    node.presetGeometry === 'trapezoid' ||
    node.presetGeometry === 'snip2DiagRect' ||
    node.presetGeometry === 'rtTriangle' ||
    node.presetGeometry === 'triangle' ||
    node.presetGeometry === 'diamond' ||
    node.presetGeometry === 'pentagon' ||
    node.presetGeometry === 'hexagon' ||
    node.presetGeometry === 'homePlate' ||
    node.presetGeometry === 'leftRightArrow' ||
    node.presetGeometry === 'chevron' ||
    node.presetGeometry === 'rightArrow' ||
    node.presetGeometry === 'leftArrow' ||
    node.presetGeometry === 'bentArrow'
  );
}

function renderPresetGeometry(node: PptSceneNode, scale: number): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('coco-ppt-preset-geometry');
  const width = Math.max(1, node.bounds.width * scale);
  const height = Math.max(1, node.bounds.height * scale);
  const extrusionDepth = presetExtrusionDepth(node, scale, width, height);
  if (node.shadow || extrusionDepth > 0) svg.style.overflow = 'visible';
  svg.setAttribute('viewBox', '0 0 100 100');
  svg.setAttribute('preserveAspectRatio', 'none');
  const gradientId =
    node.fillGradient && node.fillGradient.colors.length > 1
      ? `coco-ppt-preset-gradient-${node.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`
      : undefined;
  const lineGradientId =
    node.lineGradient && node.lineGradient.colors.length > 1
      ? `coco-ppt-preset-line-gradient-${node.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`
      : undefined;
  if (gradientId && node.fillGradient) appendSvgLinearGradient(svg, gradientId, node.fillGradient);
  if (lineGradientId && node.lineGradient)
    appendSvgLinearGradient(svg, lineGradientId, node.lineGradient);
  if (extrusionDepth > 0) {
    const backPath = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    backPath.classList.add('coco-ppt-preset-extrusion');
    backPath.setAttribute('d', presetGeometryPathData(node));
    if (node.presetGeometry === 'donut' || node.presetGeometry === 'blockArc')
      backPath.setAttribute('fill-rule', 'evenodd');
    backPath.setAttribute('fill', presetExtrusionFill(node));
    backPath.setAttribute(
      'transform',
      `translate(${fmtSvgNumber(extrusionDepth * 0.38)} ${fmtSvgNumber(extrusionDepth)})`,
    );
    backPath.setAttribute('stroke', 'none');
    svg.append(backPath);
  }
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', presetGeometryPathData(node));
  if (node.presetGeometry === 'donut' || node.presetGeometry === 'blockArc')
    path.setAttribute('fill-rule', 'evenodd');
  const strokeWidth =
    node.lineColor || node.lineGradient
      ? Math.max(0.5, ((node.lineWidth ?? 9525) * scale * 100) / Math.max(width, height))
      : 0;
  path.setAttribute(
    'fill',
    node.presetGeometry === 'arc'
      ? 'none'
      : gradientId
        ? `url(#${gradientId})`
        : (node.fillColor ?? 'none'),
  );
  const fillOpacity = effectiveFillOpacity(node);
  if (fillOpacity !== undefined && node.presetGeometry !== 'arc')
    path.setAttribute('fill-opacity', String(fillOpacity));
  path.setAttribute(
    'stroke',
    lineGradientId ? `url(#${lineGradientId})` : (node.lineColor ?? 'none'),
  );
  path.setAttribute('stroke-width', String(strokeWidth));
  if (node.lineOpacity !== undefined)
    path.setAttribute('stroke-opacity', String(Math.max(0, Math.min(1, node.lineOpacity))));
  const dashArray = svgStrokeDashArray(node.lineDash, strokeWidth);
  if (dashArray) path.setAttribute('stroke-dasharray', dashArray);
  path.setAttribute('stroke-linejoin', 'round');
  if (node.presetGeometry === 'arc') {
    path.setAttribute('stroke-linecap', 'round');
    appendArcArrowMarker(svg, path, node);
  }
  svg.append(path);
  return svg;
}

function presetExtrusionDepth(
  node: PptSceneNode,
  scale: number,
  width: number,
  height: number,
): number {
  const extrusionHeight = node.threeD?.extrusionHeight ?? 0;
  if (extrusionHeight <= 0 || node.presetGeometry === 'arc') return 0;
  const rawDepth = (extrusionHeight * scale * 100) / Math.max(width, height);
  return Math.max(0.8, Math.min(3, rawDepth));
}

function presetExtrusionFill(node: PptSceneNode): string {
  const fill = node.fillGradient?.colors[0] ?? node.fillColor ?? '#64748B';
  const parsed = parseCssRgbaColor(fill);
  const alpha = node.threeD?.material === 'translucentPowder' ? 0.12 : 0.18;
  return cssColorWithAlpha(darkenHexColor(parsed?.color ?? fill, 0.8), alpha);
}

function darkenHexColor(color: string, ratio: number): string {
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) return color;
  const value = match[1]!;
  const r = Number.parseInt(value.slice(0, 2), 16);
  const g = Number.parseInt(value.slice(2, 4), 16);
  const b = Number.parseInt(value.slice(4, 6), 16);
  return `#${toHexByte(r * ratio)}${toHexByte(g * ratio)}${toHexByte(b * ratio)}`;
}

function presetGeometryPathData(node: PptSceneNode): string {
  switch (node.presetGeometry) {
    case 'curvedUpArrow':
      return [
        'M 16 88',
        'C 3 61 9 28 38 13',
        'C 60 2 87 9 97 34',
        'L 82 36',
        'C 74 22 59 18 45 24',
        'C 23 33 17 61 28 82',
        'Z',
        'M 76 21',
        'L 100 38',
        'L 74 53',
        'Z',
      ].join(' ');
    case 'donut':
      return ['M 50 2', 'A 48 48 0 1 1 49.99 2', 'M 50 27', 'A 23 23 0 1 0 50.01 27'].join(' ');
    case 'blockArc':
      return blockArcPresetPathData(node);
    case 'plaque':
      return 'M 4 0 H 96 Q 96 8 100 8 V 92 Q 96 92 96 100 H 4 Q 4 92 0 92 V 8 Q 4 8 4 0 Z';
    case 'parallelogram':
      return 'M 18 0 H 100 L 82 100 H 0 Z';
    case 'trapezoid':
      return 'M 18 0 H 82 L 100 100 H 0 Z';
    case 'snip2DiagRect':
      return 'M 8 0 H 100 V 92 L 92 100 H 0 V 8 Z';
    case 'rtTriangle':
      return 'M 0 0 L 100 100 H 0 Z';
    case 'triangle':
      return 'M 50 0 L 100 100 H 0 Z';
    case 'diamond':
      return 'M 50 0 L 100 50 L 50 100 L 0 50 Z';
    case 'pentagon':
      return 'M 0 0 H 72 L 100 50 L 72 100 H 0 L 22 50 Z';
    case 'hexagon':
      return 'M 25 0 H 75 L 100 50 L 75 100 H 25 L 0 50 Z';
    case 'homePlate':
      return 'M 50 0 L 100 25 V 100 H 0 V 25 Z';
    case 'leftRightArrow':
      return 'M 0 50 L 18 0 H 82 L 100 50 L 82 100 H 18 Z';
    case 'chevron':
      return 'M 0 0 H 72 L 100 50 L 72 100 H 0 L 28 50 Z';
    case 'rightArrow':
      return 'M 0 0 H 78 L 100 50 L 78 100 H 0 Z';
    case 'leftArrow':
      return 'M 22 0 H 100 V 100 H 22 L 0 50 Z';
    case 'bentArrow':
      return 'M 0 0 H 30 V 58 H 82 V 35 L 100 70 L 82 100 V 77 H 12 V 18 H 0 Z';
    case 'arc':
      return arcPresetPathData(node);
    default:
      return '';
  }
}

function blockArcPresetPathData(node: PptSceneNode): string {
  const startDegrees = (node.presetAdjustments?.adj1 ?? 0) / 60000;
  const endAdjustment = (node.presetAdjustments?.adj2 ?? 5400000) / 60000;
  const sweepDegrees = normalizeSweepDegrees(endAdjustment - startDegrees);
  const endDegrees = startDegrees + sweepDegrees;
  const outerRadius = 48;
  const thicknessRatio = Math.max(
    0.08,
    Math.min(0.42, (node.presetAdjustments?.adj3 ?? 25000) / 100000),
  );
  const innerRadius = outerRadius * (1 - thicknessRatio);
  const outerStart = angleToRadiusPoint(startDegrees, outerRadius);
  const outerEnd = angleToRadiusPoint(endDegrees, outerRadius);
  const innerEnd = angleToRadiusPoint(endDegrees, innerRadius);
  const innerStart = angleToRadiusPoint(startDegrees, innerRadius);
  const largeArc = Math.abs(sweepDegrees) > 180 ? 1 : 0;
  const sweepFlag = sweepDegrees >= 0 ? 1 : 0;
  const reverseSweepFlag = sweepFlag ? 0 : 1;
  return [
    `M ${outerStart.x.toFixed(3)} ${outerStart.y.toFixed(3)}`,
    `A ${outerRadius} ${outerRadius} 0 ${largeArc} ${sweepFlag} ${outerEnd.x.toFixed(3)} ${outerEnd.y.toFixed(3)}`,
    `L ${innerEnd.x.toFixed(3)} ${innerEnd.y.toFixed(3)}`,
    `A ${innerRadius.toFixed(3)} ${innerRadius.toFixed(3)} 0 ${largeArc} ${reverseSweepFlag} ${innerStart.x.toFixed(3)} ${innerStart.y.toFixed(3)}`,
    'Z',
  ].join(' ');
}

function normalizeSweepDegrees(value: number): number {
  if (!Number.isFinite(value) || value === 0) return 90;
  let normalized = value % 360;
  if (normalized > 330) normalized = 330;
  if (normalized < -330) normalized = -330;
  if (Math.abs(normalized) < 8) normalized = normalized < 0 ? -8 : 8;
  return normalized;
}

function arcPresetPathData(node: PptSceneNode): string {
  const start = angleToPoint(node.presetAdjustments?.adj1 ?? 18000000);
  const end = angleToPoint(node.presetAdjustments?.adj2 ?? 0);
  const startDegrees = (node.presetAdjustments?.adj1 ?? 18000000) / 60000;
  const endDegrees = (node.presetAdjustments?.adj2 ?? 0) / 60000;
  const clockwiseSpan = (((endDegrees - startDegrees) % 360) + 360) % 360;
  const largeArc = clockwiseSpan > 180 ? 1 : 0;
  return `M ${start.x.toFixed(3)} ${start.y.toFixed(3)} A 48 48 0 ${largeArc} 1 ${end.x.toFixed(3)} ${end.y.toFixed(3)}`;
}

function angleToPoint(rawAngle: number): { x: number; y: number } {
  return angleToRadiusPoint(rawAngle / 60000, 48);
}

function angleToRadiusPoint(degrees: number, radius: number): { x: number; y: number } {
  const radians = (degrees / 180) * Math.PI;
  return {
    x: 50 + radius * Math.cos(radians),
    y: 50 + radius * Math.sin(radians),
  };
}

function appendArcArrowMarker(svg: SVGSVGElement, path: SVGPathElement, node: PptSceneNode): void {
  if (node.lineHeadEnd !== 'triangle' && node.lineTailEnd !== 'triangle') return;
  const markerId = `coco-ppt-arc-arrow-${node.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`;
  const marker = document.createElementNS('http://www.w3.org/2000/svg', 'marker');
  marker.setAttribute('id', markerId);
  marker.setAttribute('viewBox', '0 0 10 10');
  marker.setAttribute('markerWidth', '7');
  marker.setAttribute('markerHeight', '7');
  marker.setAttribute('refX', '5');
  marker.setAttribute('refY', '5');
  marker.setAttribute('orient', 'auto-start-reverse');
  const arrow = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  arrow.setAttribute('d', 'M 0 0 L 10 5 L 0 10 Z');
  arrow.setAttribute('fill', node.lineColor ?? node.fillColor ?? '#000000');
  marker.append(arrow);
  svg.append(marker);
  if (node.lineHeadEnd === 'triangle') path.setAttribute('marker-start', `url(#${markerId})`);
  if (node.lineTailEnd === 'triangle') path.setAttribute('marker-end', `url(#${markerId})`);
}

function shouldClipCustomGeometryPicture(node: PptSceneNode, slide: PptSlideScene): boolean {
  if (!node.customGeometry) return false;
  if (customGeometryIsRectangular(node.customGeometry)) return false;
  const nodeArea = Math.max(0, node.bounds.width) * Math.max(0, node.bounds.height);
  const slideArea = Math.max(1, slide.size.width * slide.size.height);
  const coversMostOfSlide = nodeArea / slideArea > 0.85;
  const nearSlideOrigin =
    Math.abs(node.bounds.left) < slide.size.width * 0.05 &&
    Math.abs(node.bounds.top) < slide.size.height * 0.05;
  return !(coversMostOfSlide && nearSlideOrigin);
}

function customGeometryIsRectangular(
  geometry: NonNullable<PptSceneNode['customGeometry']>,
): boolean {
  if (geometry.paths.length !== 1) return false;
  const path = geometry.paths[0];
  if (!path) return false;
  const width = path.width ?? 0;
  const height = path.height ?? 0;
  if (width <= 0 || height <= 0) return false;
  const points: Array<{ x: number; y: number }> = [];
  for (const command of path.commands) {
    if (command.kind === 'moveTo' || command.kind === 'lineTo') {
      points.push({ x: command.x, y: command.y });
    } else if (command.kind !== 'close') {
      return false;
    }
  }
  const unique = dedupeGeometryPoints(points);
  if (unique.length !== 4 && unique.length !== 5) return false;
  const corners = [
    { x: 0, y: 0 },
    { x: width, y: 0 },
    { x: width, y: height },
    { x: 0, y: height },
  ];
  return corners.every((corner) =>
    unique.some(
      (point) => geometryPointNear(point.x, corner.x) && geometryPointNear(point.y, corner.y),
    ),
  );
}

function dedupeGeometryPoints(
  points: Array<{ x: number; y: number }>,
): Array<{ x: number; y: number }> {
  const result: Array<{ x: number; y: number }> = [];
  for (const point of points) {
    if (
      !result.some(
        (existing) =>
          geometryPointNear(existing.x, point.x) && geometryPointNear(existing.y, point.y),
      )
    ) {
      result.push(point);
    }
  }
  return result;
}

function geometryPointNear(actual: number, expected: number): boolean {
  return Math.abs(actual - expected) <= 1;
}

function renderCustomGeometryPicture(node: PptSceneNode, url: string): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('coco-ppt-custom-geometry', 'coco-ppt-picture-clip');
  const pathWidth = Math.max(1, node.customGeometry?.paths[0]?.width ?? node.bounds.width);
  const pathHeight = Math.max(1, node.customGeometry?.paths[0]?.height ?? node.bounds.height);
  svg.setAttribute('viewBox', `0 0 ${pathWidth} ${pathHeight}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  const clipId = `coco-ppt-picture-clip-${node.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`;
  const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
  const clipPath = document.createElementNS('http://www.w3.org/2000/svg', 'clipPath');
  clipPath.setAttribute('id', clipId);
  for (const path of node.customGeometry?.paths ?? []) {
    const clipShape = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    clipShape.setAttribute('d', customGeometryPathData(path.commands));
    if (customGeometryNeedsEvenOddFill(path.commands))
      clipShape.setAttribute('fill-rule', 'evenodd');
    clipPath.append(clipShape);
  }
  defs.append(clipPath);
  svg.append(defs);
  const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
  image.setAttribute('href', url);
  image.setAttributeNS('http://www.w3.org/1999/xlink', 'href', url);
  const placement = imagePlacementRect(node, pathWidth, pathHeight);
  image.setAttribute('x', String(placement.x));
  image.setAttribute('y', String(placement.y));
  image.setAttribute('width', String(placement.width));
  image.setAttribute('height', String(placement.height));
  image.setAttribute('preserveAspectRatio', 'none');
  image.setAttribute('clip-path', `url(#${clipId})`);
  if (node.imageOpacity !== undefined)
    image.setAttribute('opacity', String(Math.max(0, Math.min(1, node.imageOpacity))));
  svg.append(image);
  return svg;
}

function imagePlacementRect(
  node: PptSceneNode,
  width: number,
  height: number,
): { x: number; y: number; width: number; height: number } {
  if (node.imageFillRect) {
    const left = node.imageFillRect.left ?? 0;
    const right = node.imageFillRect.right ?? 0;
    const top = node.imageFillRect.top ?? 0;
    const bottom = node.imageFillRect.bottom ?? 0;
    return {
      x: left * width,
      y: top * height,
      width: Math.max(0.01, 1 - left - right) * width,
      height: Math.max(0.01, 1 - top - bottom) * height,
    };
  }
  const crop = node.imageCrop;
  if (!crop) return { x: 0, y: 0, width, height };
  const left = crop.left ?? 0;
  const right = crop.right ?? 0;
  const top = crop.top ?? 0;
  const bottom = crop.bottom ?? 0;
  const visibleWidth = Math.max(0.01, 1 - left - right);
  const visibleHeight = Math.max(0.01, 1 - top - bottom);
  return {
    x: (-left / visibleWidth) * width,
    y: (-top / visibleHeight) * height,
    width: (1 / visibleWidth) * width,
    height: (1 / visibleHeight) * height,
  };
}

function renderCustomGeometry(node: PptSceneNode, scale: number): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('coco-ppt-custom-geometry');
  const width = node.customGeometry?.paths[0]?.width ?? node.bounds.width;
  const height = node.customGeometry?.paths[0]?.height ?? node.bounds.height;
  svg.setAttribute('viewBox', `0 0 ${Math.max(1, width)} ${Math.max(1, height)}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  const gradientId =
    node.fillGradient && node.fillGradient.colors.length > 1
      ? `coco-ppt-gradient-${node.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`
      : undefined;
  const lineGradientId =
    node.lineGradient && node.lineGradient.colors.length > 1
      ? `coco-ppt-custom-line-gradient-${node.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`
      : undefined;
  if (gradientId && node.fillGradient) appendSvgLinearGradient(svg, gradientId, node.fillGradient);
  if (lineGradientId && node.lineGradient)
    appendSvgLinearGradient(svg, lineGradientId, node.lineGradient);
  for (const path of node.customGeometry?.paths ?? []) {
    const strokeWidth = customGeometryStrokeWidth(node, path);
    const ellipse = customGeometryEllipse(path);
    const shape = ellipse
      ? document.createElementNS('http://www.w3.org/2000/svg', 'ellipse')
      : document.createElementNS('http://www.w3.org/2000/svg', 'path');
    if (ellipse) {
      shape.setAttribute('cx', String(ellipse.cx));
      shape.setAttribute('cy', String(ellipse.cy));
      shape.setAttribute('rx', String(ellipse.rx));
      shape.setAttribute('ry', String(ellipse.ry));
    } else {
      shape.setAttribute('d', customGeometryPathData(path.commands));
      if (customGeometryNeedsEvenOddFill(path.commands)) shape.setAttribute('fill-rule', 'evenodd');
    }
    const hasFill = !(node.kind === 'line' && node.lineDash);
    shape.setAttribute(
      'fill',
      hasFill ? (gradientId ? `url(#${gradientId})` : (node.fillColor ?? 'none')) : 'none',
    );
    const fillOpacity = effectiveFillOpacity(node);
    if (hasFill && fillOpacity !== undefined)
      shape.setAttribute('fill-opacity', String(fillOpacity));
    shape.setAttribute(
      'stroke',
      lineGradientId ? `url(#${lineGradientId})` : (node.lineColor ?? 'none'),
    );
    shape.setAttribute('stroke-width', String(strokeWidth));
    if (node.lineOpacity !== undefined)
      shape.setAttribute('stroke-opacity', String(Math.max(0, Math.min(1, node.lineOpacity))));
    const dashArray = svgStrokeDashArray(node.lineDash, strokeWidth);
    if (dashArray) shape.setAttribute('stroke-dasharray', dashArray);
    shape.setAttribute('stroke-linejoin', 'round');
    shape.setAttribute('stroke-linecap', 'round');
    svg.append(shape);
  }
  return svg;
}

function customGeometryEllipse(
  path: PptCustomGeometryPath,
): { cx: number; cy: number; rx: number; ry: number } | undefined {
  const width = path.width;
  const height = path.height;
  if (!width || !height || width <= 0 || height <= 0) return undefined;
  const commands = path.commands;
  if (
    commands.length !== 6 ||
    commands[0]?.kind !== 'moveTo' ||
    commands[1]?.kind !== 'cubicBezierTo' ||
    commands[2]?.kind !== 'cubicBezierTo' ||
    commands[3]?.kind !== 'cubicBezierTo' ||
    commands[4]?.kind !== 'cubicBezierTo' ||
    commands[5]?.kind !== 'close'
  ) {
    return undefined;
  }
  const cx = width / 2;
  const cy = height / 2;
  const endpoints = [commands[0], commands[1], commands[2], commands[3], commands[4]];
  const expected = [
    { x: width, y: cy },
    { x: cx, y: height },
    { x: 0, y: cy },
    { x: cx, y: 0 },
    { x: width, y: cy },
  ];
  const tolerance = Math.max(width, height) * 0.003;
  const endpointMatches = endpoints.every((command, index) => {
    if (command.kind !== 'moveTo' && command.kind !== 'cubicBezierTo') return false;
    return (
      Math.abs(command.x - expected[index]!.x) <= tolerance &&
      Math.abs(command.y - expected[index]!.y) <= tolerance
    );
  });
  return endpointMatches
    ? {
        cx,
        cy,
        rx: width / 2,
        ry: height / 2,
      }
    : undefined;
}

function customGeometryNeedsEvenOddFill(commands: PptCustomGeometryCommand[]): boolean {
  return commands.filter((command) => command.kind === 'moveTo').length > 1;
}

function customGeometryStrokeWidth(node: PptSceneNode, path: PptCustomGeometryPath): number {
  if (!node.lineWidth) return 0;
  const widthRatio = path.width && node.bounds.width ? path.width / node.bounds.width : undefined;
  const heightRatio =
    path.height && node.bounds.height ? path.height / node.bounds.height : undefined;
  const ratios = [widthRatio, heightRatio].filter(
    (ratio): ratio is number => ratio !== undefined && Number.isFinite(ratio) && ratio > 0,
  );
  const ratio = ratios.length > 0 ? Math.max(...ratios) : 1;
  return Math.max(0.5, node.lineWidth * ratio);
}

function renderGradientOutline(node: PptSceneNode, scale: number): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('coco-ppt-gradient-outline');
  const width = Math.max(1, node.bounds.width * scale);
  const height = Math.max(1, node.bounds.height * scale);
  const strokeWidth = Math.max(1, (node.lineWidth ?? 9525) * scale);
  const inset = strokeWidth / 2;
  const gradientId = `coco-ppt-line-gradient-${node.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`;
  svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  appendSvgLinearGradient(svg, gradientId, node.lineGradient!);
  const shape =
    node.presetGeometry === 'ellipse'
      ? document.createElementNS('http://www.w3.org/2000/svg', 'ellipse')
      : document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  if (shape.tagName === 'ellipse') {
    shape.setAttribute('cx', String(width / 2));
    shape.setAttribute('cy', String(height / 2));
    shape.setAttribute('rx', String(Math.max(0, width / 2 - inset)));
    shape.setAttribute('ry', String(Math.max(0, height / 2 - inset)));
  } else {
    shape.setAttribute('x', String(inset));
    shape.setAttribute('y', String(inset));
    shape.setAttribute('width', String(Math.max(0, width - strokeWidth)));
    shape.setAttribute('height', String(Math.max(0, height - strokeWidth)));
    shape.setAttribute('rx', String(Math.max(0, roundRectRadius(node) * scale - inset)));
    shape.setAttribute('ry', String(Math.max(0, roundRectRadius(node) * scale - inset)));
  }
  shape.setAttribute('fill', 'none');
  shape.setAttribute('stroke', `url(#${gradientId})`);
  shape.setAttribute('stroke-width', String(strokeWidth));
  if (node.lineOpacity !== undefined)
    shape.setAttribute('stroke-opacity', String(Math.max(0, Math.min(1, node.lineOpacity))));
  const dashArray = svgStrokeDashArray(node.lineDash, strokeWidth);
  if (dashArray) shape.setAttribute('stroke-dasharray', dashArray);
  svg.append(shape);
  return svg;
}

function renderSolidOutline(node: PptSceneNode, scale: number): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('coco-ppt-solid-outline');
  const width = Math.max(1, node.bounds.width * scale);
  const height = Math.max(1, node.bounds.height * scale);
  const strokeWidth = Math.max(1, (node.lineWidth ?? 9525) * scale);
  const strokeInsideBounds = node.kind !== 'picture';
  const inset = strokeInsideBounds ? strokeWidth / 2 : 0;
  if (!strokeInsideBounds) svg.style.overflow = 'visible';
  svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  const shape =
    node.presetGeometry === 'ellipse'
      ? document.createElementNS('http://www.w3.org/2000/svg', 'ellipse')
      : document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  if (shape.tagName === 'ellipse') {
    shape.setAttribute('cx', String(width / 2));
    shape.setAttribute('cy', String(height / 2));
    shape.setAttribute('rx', String(Math.max(0, width / 2 - inset)));
    shape.setAttribute('ry', String(Math.max(0, height / 2 - inset)));
  } else {
    shape.setAttribute('x', String(inset));
    shape.setAttribute('y', String(inset));
    shape.setAttribute('width', String(Math.max(0, width - inset * 2)));
    shape.setAttribute('height', String(Math.max(0, height - inset * 2)));
    const radius =
      node.presetGeometry === 'roundRect' || node.presetGeometry === 'round1Rect'
        ? Math.max(0, roundRectRadius(node) * scale - inset)
        : 0;
    shape.setAttribute('rx', String(radius));
    shape.setAttribute('ry', String(radius));
  }
  shape.setAttribute('fill', 'none');
  shape.setAttribute('stroke', node.lineColor ?? 'none');
  shape.setAttribute('stroke-width', String(strokeWidth));
  if (node.lineOpacity !== undefined)
    shape.setAttribute('stroke-opacity', String(Math.max(0, Math.min(1, node.lineOpacity))));
  const dashArray = svgStrokeDashArray(node.lineDash, strokeWidth);
  if (dashArray) shape.setAttribute('stroke-dasharray', dashArray);
  svg.append(shape);
  return svg;
}

function renderLineGeometry(node: PptSceneNode, scale: number): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('coco-ppt-line-geometry');
  const isVertical = lineAxisKind(node) === 'vertical';
  const width = Math.max(isVertical ? 12 : 1, node.bounds.width * scale);
  const height = Math.max(isVertical ? 1 : 12, node.bounds.height * scale);
  svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  const strokeWidth = Math.max(1, (node.lineWidth ?? 9525) * scale);
  const gradientId =
    node.lineGradient && node.lineGradient.colors.length > 1
      ? `coco-ppt-connector-gradient-${node.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`
      : undefined;
  if (gradientId && node.lineGradient) appendSvgLinearGradient(svg, gradientId, node.lineGradient);
  const color = gradientId
    ? `url(#${gradientId})`
    : cssColorWithOpacity(node.lineColor ?? '#000000', node.lineOpacity);
  const ovalRadius = ovalLineEndRadius(strokeWidth);
  const arrowLength = Math.max(5, strokeWidth * 3.2);
  const arrowWidth = Math.max(5, strokeWidth * 2.6);
  const headIsArrow = isArrowLineEnd(node.lineHeadEnd);
  const tailIsArrow = isArrowLineEnd(node.lineTailEnd);
  if (isVertical) {
    const x = width / 2;
    const y1 = headIsArrow ? Math.min(height, arrowLength) : 0;
    const y2 = tailIsArrow ? Math.max(0, height - arrowLength) : height;
    appendSvgLine(svg, x, y1, x, y2, color, strokeWidth, node.lineOpacity, node.lineCap);
    if (node.lineHeadEnd === 'oval') appendSvgCircle(svg, x, 0, ovalRadius, color);
    if (node.lineTailEnd === 'oval') appendSvgCircle(svg, x, height, ovalRadius, color);
    if (headIsArrow)
      appendSvgArrow(svg, x, 0, 'up', arrowLength, arrowWidth, color, node.lineOpacity);
    if (tailIsArrow)
      appendSvgArrow(svg, x, height, 'down', arrowLength, arrowWidth, color, node.lineOpacity);
  } else {
    const y = height / 2;
    const x1 = headIsArrow ? Math.min(width, arrowLength) : 0;
    const x2 = tailIsArrow ? Math.max(0, width - arrowLength) : width;
    appendSvgLine(svg, x1, y, x2, y, color, strokeWidth, node.lineOpacity, node.lineCap);
    if (node.lineHeadEnd === 'oval') appendSvgCircle(svg, 0, y, ovalRadius, color);
    if (node.lineTailEnd === 'oval') appendSvgCircle(svg, width, y, ovalRadius, color);
    if (headIsArrow)
      appendSvgArrow(svg, 0, y, 'left', arrowLength, arrowWidth, color, node.lineOpacity);
    if (tailIsArrow)
      appendSvgArrow(svg, width, y, 'right', arrowLength, arrowWidth, color, node.lineOpacity);
  }
  return svg;
}

function ovalLineEndRadius(strokeWidth: number): number {
  return Math.max(3, strokeWidth * 3);
}

function isArrowLineEnd(value: string | undefined): boolean {
  return value === 'arrow' || value === 'triangle';
}

function customGeometryPathData(commands: PptCustomGeometryCommand[]): string {
  const parts: string[] = [];
  let current: { x: number; y: number } | undefined;
  for (const command of commands) {
    if (command.kind === 'moveTo') {
      parts.push(`M ${command.x} ${command.y}`);
      current = { x: command.x, y: command.y };
    } else if (command.kind === 'lineTo') {
      parts.push(`L ${command.x} ${command.y}`);
      current = { x: command.x, y: command.y };
    } else if (command.kind === 'cubicBezierTo') {
      parts.push(
        `C ${command.x1} ${command.y1} ${command.x2} ${command.y2} ${command.x} ${command.y}`,
      );
      current = { x: command.x, y: command.y };
    } else if (command.kind === 'arcTo') {
      const arc = svgArcCommandForCustomGeometry(command, current);
      if (arc) {
        parts.push(arc.path);
        current = arc.end;
      }
    } else {
      parts.push('Z');
    }
  }
  return parts.join(' ');
}

function svgArcCommandForCustomGeometry(
  command: Extract<PptCustomGeometryCommand, { kind: 'arcTo' }>,
  current: { x: number; y: number } | undefined,
): { path: string; end: { x: number; y: number } } | undefined {
  if (!current || command.wR <= 0 || command.hR <= 0 || command.swAng === 0) return undefined;
  const startDegrees = command.stAng / 60000;
  const sweepDegrees = command.swAng / 60000;
  const endDegrees = startDegrees + sweepDegrees;
  const startVector = customGeometryArcVector(command.wR, command.hR, startDegrees);
  const center = {
    x: current.x - startVector.x,
    y: current.y - startVector.y,
  };
  const endVector = customGeometryArcVector(command.wR, command.hR, endDegrees);
  const end = {
    x: center.x + endVector.x,
    y: center.y + endVector.y,
  };
  const largeArc = Math.abs(sweepDegrees) > 180 ? 1 : 0;
  const sweepFlag = sweepDegrees >= 0 ? 1 : 0;
  return {
    path: `A ${fmtSvgNumber(command.wR)} ${fmtSvgNumber(command.hR)} 0 ${largeArc} ${sweepFlag} ${fmtSvgNumber(end.x)} ${fmtSvgNumber(end.y)}`,
    end,
  };
}

function customGeometryArcVector(
  radiusX: number,
  radiusY: number,
  degrees: number,
): {
  x: number;
  y: number;
} {
  const radians = (degrees / 180) * Math.PI;
  return {
    x: radiusX * Math.cos(radians),
    y: radiusY * Math.sin(radians),
  };
}

function fmtSvgNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(3).replace(/\.?0+$/, '');
}

function appendSvgLinearGradient(
  svg: SVGSVGElement,
  id: string,
  gradientModel: PptGradientModel | string[],
): void {
  const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
  const colors = Array.isArray(gradientModel) ? gradientModel : gradientModel.colors;
  const positions = Array.isArray(gradientModel) ? undefined : gradientModel.positions;
  const opacities = Array.isArray(gradientModel) ? undefined : gradientModel.opacities;
  const gradient = document.createElementNS('http://www.w3.org/2000/svg', 'linearGradient');
  const direction = Array.isArray(gradientModel)
    ? { x1: 0, y1: 0, x2: 100, y2: 100 }
    : svgLinearGradientDirection(cssLinearGradientAngle(gradientModel));
  gradient.setAttribute('id', id);
  gradient.setAttribute('x1', `${direction.x1}%`);
  gradient.setAttribute('y1', `${direction.y1}%`);
  gradient.setAttribute('x2', `${direction.x2}%`);
  gradient.setAttribute('y2', `${direction.y2}%`);
  colors
    .map((color: string, index: number) => {
      const offset = positions?.[index] ?? (index / Math.max(1, colors.length - 1)) * 100;
      const parsedColor = parseCssRgbaColor(color);
      const opacity = Math.max(0, Math.min(1, opacities?.[index] ?? parsedColor?.opacity ?? 1));
      return {
        color: parsedColor?.color ?? color,
        offset: Math.max(0, Math.min(100, offset)),
        opacity,
      };
    })
    .sort((a, b) => a.offset - b.offset)
    .forEach((item) => {
      const stop = document.createElementNS('http://www.w3.org/2000/svg', 'stop');
      stop.setAttribute('offset', `${item.offset}%`);
      stop.setAttribute('stop-color', item.color);
      if (item.opacity < 1) stop.setAttribute('stop-opacity', String(item.opacity));
      gradient.append(stop);
    });
  defs.append(gradient);
  svg.append(defs);
}

function svgLinearGradientDirection(angle: number): {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
} {
  const radians = (angle * Math.PI) / 180;
  const dx = Math.sin(radians);
  const dy = -Math.cos(radians);
  const extent = 50 / Math.max(Math.abs(dx), Math.abs(dy), 0.0001);
  return {
    x1: roundGradientCoordinate(50 - dx * extent),
    y1: roundGradientCoordinate(50 - dy * extent),
    x2: roundGradientCoordinate(50 + dx * extent),
    y2: roundGradientCoordinate(50 + dy * extent),
  };
}

function roundGradientCoordinate(value: number): number {
  return Math.round(Math.max(0, Math.min(100, value)) * 1000) / 1000;
}

function parseCssRgbaColor(color: string): { color: string; opacity: number } | undefined {
  const match =
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([0-9.]+)\s*)?\)$/i.exec(color);
  if (!match) return undefined;
  const r = Math.max(0, Math.min(255, Number(match[1])));
  const g = Math.max(0, Math.min(255, Number(match[2])));
  const b = Math.max(0, Math.min(255, Number(match[3])));
  const opacity = match[4] === undefined ? 1 : Math.max(0, Math.min(1, Number(match[4])));
  return {
    color: `#${toHexByte(r)}${toHexByte(g)}${toHexByte(b)}`,
    opacity,
  };
}

function toHexByte(value: number): string {
  return Math.round(value).toString(16).padStart(2, '0').toUpperCase();
}

function renderTablePreview(
  node: PptSceneNode,
  scale: number,
  interactive = true,
): HTMLTableElement {
  const table = document.createElement('table');
  table.className = 'coco-ppt-table';
  const model = node.table;
  if (!model) return table;
  const rowCount = Math.max(model.rowCount, model.values.length);
  const columnCount = Math.max(model.columnCount, ...model.values.map((row) => row.length), 1);
  const totalColumnWidth = model.columnWidths?.reduce((sum, width) => sum + width, 0) ?? 0;
  if (totalColumnWidth > 0) {
    const colgroup = document.createElement('colgroup');
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const col = document.createElement('col');
      const width = model.columnWidths?.[columnIndex];
      if (width) col.style.width = `${(width / totalColumnWidth) * 100}%`;
      colgroup.append(col);
    }
    table.append(colgroup);
  }
  const totalRowHeight = model.rowHeights?.reduce((sum, height) => sum + height, 0) ?? 0;
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    const row = document.createElement('tr');
    const height = model.rowHeights?.[rowIndex];
    row.style.height =
      totalRowHeight > 0 && height ? `${(height / totalRowHeight) * 100}%` : `${100 / rowCount}%`;
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const merge = tableCellMergeAt(model, rowIndex, columnIndex);
      if (merge?.hMerge || merge?.vMerge) continue;
      const cell = document.createElement('td');
      const style = model.cellStyles?.[rowIndex]?.[columnIndex];
      if (interactive) {
        cell.dataset.row = String(rowIndex);
        cell.dataset.column = String(columnIndex);
      }
      const text = model.values[rowIndex]?.[columnIndex] ?? '';
      cell.dataset.cellText = text;
      appendTableCellText(cell, text);
      if (merge?.columnSpan && merge.columnSpan > 1) cell.colSpan = merge.columnSpan;
      if (merge?.rowSpan && merge.rowSpan > 1) cell.rowSpan = merge.rowSpan;
      if (style?.fillGradient && style.fillGradient.colors.length > 1) {
        cell.style.background = cssLinearGradient(style.fillGradient);
      } else if (style?.fillColor && !isTransparentTableFill(style.fillColor)) {
        cell.style.background = style.fillColor;
      }
      if (style?.body) {
        const body = style.body;
        const left = body.insetLeft ?? 91440;
        const top = body.insetTop ?? 45720;
        const right = body.insetRight ?? 91440;
        const bottom = body.insetBottom ?? 45720;
        cell.style.padding = `${top * scale}px ${right * scale}px ${bottom * scale}px ${left * scale}px`;
        if (body.anchor === 'middle') cell.style.verticalAlign = 'middle';
        if (body.anchor === 'bottom') cell.style.verticalAlign = 'bottom';
        if (body.anchor === 'top') cell.style.verticalAlign = 'top';
      }
      if (style?.textStyle) applyTextStyleToElement(cell, style.textStyle, scale);
      if (style?.textStyle?.alignment) {
        applyTextAlignment(cell, style.textStyle.alignment);
      }
      applyTableCellBorders(cell, style, scale);
      row.append(cell);
    }
    table.append(row);
  }
  return table;
}

function applyTableCellBorders(
  cell: HTMLElement,
  style: PptTableCellStyle | undefined,
  scale: number,
): void {
  const borders = style?.borders;
  if (!borders) return;
  applyTableCellBorderSide(cell, 'Left', borders.left, scale);
  applyTableCellBorderSide(cell, 'Right', borders.right, scale);
  applyTableCellBorderSide(cell, 'Top', borders.top, scale);
  applyTableCellBorderSide(cell, 'Bottom', borders.bottom, scale);
}

function applyTableCellBorderSide(
  cell: HTMLElement,
  side: 'Left' | 'Right' | 'Top' | 'Bottom',
  border: NonNullable<PptTableCellStyle['borders']>['left'] | undefined,
  scale: number,
): void {
  if (!border) return;
  const styleProp = `border${side}Style` as const;
  const widthProp = `border${side}Width` as const;
  const colorProp = `border${side}Color` as const;
  if (border.hidden) {
    cell.style[styleProp] = 'none';
    cell.style[widthProp] = '0';
    return;
  }
  cell.style[styleProp] = cssBorderStyleForLineDash(border.dash);
  cell.style[widthProp] = `${Math.max(1, (border.width ?? 9525) * scale)}px`;
  cell.style[colorProp] = cssColorWithOpacity(border.color ?? '#94A3B8', border.opacity);
}

function appendTableCellText(cell: HTMLElement, text: string): void {
  const lines = text.split(/\r?\n/);
  if (lines.length <= 1) {
    cell.textContent = text;
    return;
  }
  lines.forEach((line, index) => {
    if (index > 0) cell.append(document.createElement('br'));
    const span = document.createElement('span');
    span.textContent = line;
    if (index > 0) {
      span.className = 'coco-ppt-table-cell-secondary-line';
      span.style.fontSize = '0.7em';
      span.style.fontWeight = '400';
    }
    cell.append(span);
  });
}

function tableCellRawText(
  cell: HTMLElement,
  node: PptSceneNode,
  row: number,
  column: number,
): string {
  return cell.dataset.cellText ?? node.table?.values[row]?.[column] ?? cell.textContent ?? '';
}

function isTransparentTableFill(color: string): boolean {
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) return false;
  const value = match[1]!;
  const r = Number.parseInt(value.slice(0, 2), 16);
  const g = Number.parseInt(value.slice(2, 4), 16);
  const b = Number.parseInt(value.slice(4, 6), 16);
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 > 0.9;
}

function cssBoxShadow(shadow: NonNullable<PptSceneNode['shadow']>, scale: number): string {
  const distance = (shadow.distance ?? 0) * scale;
  const direction = ((shadow.direction ?? 0) * Math.PI) / 180;
  const offsetX = Math.cos(direction) * distance;
  const offsetY = Math.sin(direction) * distance;
  const blur = (shadow.blur ?? 0) * scale;
  return `${offsetX}px ${offsetY}px ${blur}px ${shadowColor(shadow)}`;
}

function shadowColor(shadow: NonNullable<PptSceneNode['shadow']>): string {
  const color = shadow.color ?? '#000000';
  const alpha = shadow.opacity ?? 0.25;
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) return color;
  const value = match[1]!;
  const r = Number.parseInt(value.slice(0, 2), 16);
  const g = Number.parseInt(value.slice(2, 4), 16);
  const b = Number.parseInt(value.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${Math.max(0, Math.min(1, alpha))})`;
}

function tableCellMergeAt(
  table: NonNullable<PptSceneNode['table']>,
  row: number,
  column: number,
): PptTableMerge | undefined {
  return table.mergedCells?.find((merge) => merge.row === row && merge.column === column);
}

function renderChartPreview(node: PptSceneNode): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 100 60');
  svg.setAttribute('width', '100%');
  svg.setAttribute('height', '100%');
  svg.setAttribute('preserveAspectRatio', 'none');
  const data = node.chart?.data;
  if (!data?.series.length) {
    appendSvgText(svg, node.name ?? 'Chart', 4, 8, 4, '#17202a');
    return svg;
  }

  const chartType = node.chart?.type ?? '';
  const grouping = node.chart?.grouping;
  const rangeValues = chartRangeValues(data.series, chartType, grouping);
  const rawMax = Math.max(...rangeValues, 1);
  const rawMin = Math.min(...rangeValues, 0);
  const valueAxis = node.chart?.axes?.value;
  const categoryAxis = node.chart?.axes?.category;
  const min = valueAxis?.min ?? (rawMin >= 0 ? 0 : rawMin);
  const max = valueAxis?.max ?? (rawMin >= 0 ? niceChartMax(rawMax) : rawMax);
  const plot = chartPlotArea(node);
  const tickCount = chartTickCount(min, max, valueAxis?.majorUnit);
  if (chartType.includes('doughnut') || chartType.includes('pie')) {
    renderDoughnutChart(svg, data.series, plot);
    return svg;
  }
  if (chartType.includes('radar')) {
    renderRadarChart(svg, data.series, plot, valueAxis?.max ?? rawMax);
    return svg;
  }
  if (valueAxis?.majorGridlines !== false)
    renderChartGrid(svg, plot, tickCount, node, valueAxis, categoryAxis);
  if (valueAxis?.deleted !== true && valueAxis?.tickLabelPosition !== 'none')
    renderValueAxisLabels(svg, plot, min, max, tickCount);
  renderChartTitleAndLegend(svg, node, plot);

  const barSeries = data.series.filter((series) =>
    (series.type ?? node.chart?.type ?? '').includes('bar'),
  );
  const lineSeries = data.series.filter((series) => (series.type ?? '').includes('line'));
  if (barSeries.length > 0) {
    renderBarSeries(svg, barSeries, plot, min, max, node, grouping);
    renderCategoryLabels(svg, barSeries[0]?.categories ?? [], plot);
    if (barSeries.some((series) => series.showValueLabels))
      renderValueLabels(svg, barSeries, plot, min, max);
    if (lineSeries.length > 0) renderLineSeries(svg, lineSeries, plot, min, max, undefined, node);
    if (lineSeries.some((series) => series.showValueLabels))
      renderLineValueLabels(svg, lineSeries, plot, min, max);
  } else {
    renderLineSeries(
      svg,
      data.series,
      plot,
      min,
      max,
      ['#22C7F2', '#22C7F2', '#94A3B8', '#F59E0B'],
      node,
    );
    renderCategoryLabels(
      svg,
      data.series.find((series) => series.categories?.length)?.categories ?? [],
      plot,
    );
    if (data.series.some((series) => series.showValueLabels))
      renderLineValueLabels(svg, data.series, plot, min, max);
  }
  return svg;
}

function chartRangeValues(
  seriesList: PptChartSeriesList,
  chartType: string,
  grouping: string | undefined,
): number[] {
  const barSeries = seriesList.filter((series) => (series.type ?? chartType).includes('bar'));
  if (!isStackedBarGrouping(grouping) || barSeries.length <= 1) {
    return seriesList.flatMap((series) => series.values);
  }
  const columns = Math.max(...barSeries.map((series) => series.values.length), 1);
  const values: number[] = [];
  for (let valueIndex = 0; valueIndex < columns; valueIndex++) {
    let positive = 0;
    let negative = 0;
    for (const series of barSeries) {
      const value = series.values[valueIndex] ?? 0;
      if (value >= 0) positive += value;
      else negative += value;
    }
    values.push(positive, negative);
  }
  return values;
}

function chartPlotArea(node: PptSceneNode): {
  left: number;
  top: number;
  width: number;
  height: number;
} {
  const layout = node.chart?.plotLayout;
  const chartType = node.chart?.type ?? '';
  if (!layout) {
    if (shouldOverlayBarSeries(node.chart?.data?.series ?? []))
      return { left: 7, top: 8, width: 88, height: 40 };
    if (isCompactCategoryBarChart(node)) return { left: 18, top: 8, width: 72, height: 40 };
    return node.chart?.axes?.value?.deleted === true
      ? { left: 6, top: 8, width: 90, height: 40 }
      : { left: 18, top: 8, width: 76, height: 40 };
  }
  const needsValueAxisInset =
    node.chart?.axes?.value?.deleted !== true &&
    !chartType.includes('doughnut') &&
    !chartType.includes('pie') &&
    !chartType.includes('radar');
  const minLeft = needsValueAxisInset ? 14 : 4;
  const left = (layout.x ?? 0.12) * 100;
  const top = (layout.y ?? 0.13) * 60;
  const width = (layout.width ?? 0.82) * 100;
  const height = (layout.height ?? 0.67) * 60;
  const clampedLeft = Math.max(minLeft, Math.min(20, left));
  return {
    left: clampedLeft,
    top: Math.max(3, Math.min(12, top)),
    width: Math.max(40, Math.min(92, Math.min(width, 96 - clampedLeft))),
    height: Math.max(24, Math.min(50, height)),
  };
}

function isCompactCategoryBarChart(node: PptSceneNode): boolean {
  const chart = node.chart;
  const chartType = chart?.type ?? '';
  if (!chart || !chartType.includes('bar') || chart.plotLayout) return false;
  const series = chart.data?.series ?? [];
  if (series.length !== 1) return false;
  const categoryCount = Math.max(series[0]?.categories?.length ?? 0, series[0]?.values.length ?? 0);
  return categoryCount > 0 && categoryCount <= 2;
}

function niceChartMax(value: number): number {
  if (value <= 0) return 1;
  if (value <= 10) return Math.max(1, Math.ceil(value));
  const padded = value * 1.1;
  const magnitude = 10 ** Math.floor(Math.log10(padded));
  return Math.ceil(padded / magnitude) * magnitude;
}

function chartTickCount(min: number, max: number, majorUnit?: number): number {
  const range = Math.abs(max - min);
  if (majorUnit && majorUnit > 0) return Math.max(1, Math.min(8, Math.round(range / majorUnit)));
  if (Number.isInteger(range) && range > 0 && range <= 6) return range;
  return 4;
}

function renderChartTitleAndLegend(
  svg: SVGSVGElement,
  node: PptSceneNode,
  plot: { left: number; top: number; width: number },
): void {
  const title = node.chart?.title ?? node.chart?.data?.title;
  if (title) {
    appendSvgText(
      svg,
      title,
      plot.left + plot.width * 0.5,
      Math.max(4, plot.top - 3.2),
      3.1,
      '#17202A',
      'middle',
    );
  }
  const series = node.chart?.data?.series ?? [];
  const names = series
    .map((item) => item.name)
    .filter((value): value is string => Boolean(value))
    .slice(0, 4);
  if (names.length === 0) return;
  const colors = ['#C50D13', '#0EA5E9', '#F59E0B', '#A3E635'];
  let x = plot.left + plot.width * 0.78;
  const y = Math.max(4, plot.top - 2.4);
  names.forEach((name, index) => {
    const seriesItem = series.find((item) => item.name === name) ?? series[index];
    const fill = seriesItem ? chartLegendColor(svg, seriesItem, index, colors) : colors[index]!;
    if (seriesItem && (seriesItem.type ?? '').includes('line')) {
      appendSvgLine(svg, x, y - 1.1, x + 2.3, y - 1.1, fill, 0.45);
      appendSvgCircle(
        svg,
        x + 1.15,
        y - 1.1,
        0.45,
        seriesItem.marker?.fillColor ?? '#FFFFFF',
        seriesItem.marker?.lineColor ?? fill,
        0.32,
      );
      appendSvgText(svg, name, x + 2.8, y - 0.45, 2.2, '#64748B');
    } else {
      appendSvgRect(svg, x, y - 1.8, 1.4, 1.4, fill);
      appendSvgText(svg, name, x + 2, y - 0.45, 2.2, '#64748B');
    }
    x += Math.max(9, Array.from(name).length * 2.2 + 4.5);
  });
}

function chartLegendColor(
  svg: SVGSVGElement,
  series: PptChartSeries,
  index: number,
  fallbackColors: string[],
): string {
  if ((series.type ?? '').includes('line')) return series.lineColor ?? fallbackColors[index]!;
  return chartSeriesFill(svg, series, index, fallbackColors);
}

function renderChartGrid(
  svg: SVGSVGElement,
  plot: { left: number; top: number; width: number; height: number },
  tickCount: number,
  node: PptSceneNode,
  valueAxis?: NonNullable<PptChartModel['axes']>['value'],
  categoryAxis?: NonNullable<PptChartModel['axes']>['category'],
): void {
  const gridColor = valueAxis?.gridlineColor ?? '#E8EDF3';
  const gridWidth = chartAxisStrokeWidth(valueAxis?.gridlineWidth, node, 0.35);
  for (let i = 0; i <= tickCount; i++) {
    const y = plot.top + (plot.height / tickCount) * i;
    appendSvgLine(svg, plot.left, y, plot.left + plot.width, y, gridColor, gridWidth);
  }
  appendSvgLine(
    svg,
    plot.left,
    plot.top + plot.height,
    plot.left + plot.width,
    plot.top + plot.height,
    categoryAxis?.lineColor ?? '#CBD5E1',
    chartAxisStrokeWidth(categoryAxis?.lineWidth, node, 0.45),
  );
}

function renderValueAxisLabels(
  svg: SVGSVGElement,
  plot: { left: number; top: number; width: number; height: number },
  min: number,
  max: number,
  tickCount: number,
): void {
  for (let i = 0; i <= tickCount; i++) {
    const value = max - ((max - min) / tickCount) * i;
    const y = plot.top + (plot.height / tickCount) * i + 1;
    appendSvgText(svg, formatChartTick(value), plot.left - 1.8, y, 3, '#9AA6B2', 'end');
  }
}

function formatChartTick(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(1).replace(/\.0$/, '');
}

function renderLineSeries(
  svg: SVGSVGElement,
  seriesList: PptChartSeriesList,
  plot: { left: number; top: number; width: number; height: number },
  min: number,
  max: number,
  fallbackColors: string[] = ['#CBD5E1', '#94A3B8', '#F59E0B', '#A3E635'],
  node?: PptSceneNode,
): void {
  seriesList.forEach((series, index) => {
    if (series.values.length === 0) return;
    const points = series.values.map((value, pointIndex) => {
      const x =
        plot.left +
        (series.values.length === 1
          ? plot.width / 2
          : (plot.width * pointIndex) / (series.values.length - 1));
      const y = plot.top + plot.height - ((value - min) / Math.max(1, max - min)) * plot.height;
      return { x, y };
    });
    const path = chartLinePath(points, series.smooth === true);
    const color = series.lineColor ?? fallbackColors[index % fallbackColors.length]!;
    appendSvgPath(svg, path, color, chartSeriesLineWidth(series, node));
    if (series.marker?.symbol !== 'none') {
      const radius = Math.max(0.5, Math.min(1.6, (series.marker?.size ?? 5) * 0.18));
      const fill = series.marker?.fillColor ?? color;
      const stroke = series.marker?.lineColor;
      const strokeWidth = chartMarkerStrokeWidth(series, node);
      points.forEach((point) =>
        appendSvgCircle(svg, point.x, point.y, radius, fill, stroke, strokeWidth),
      );
    }
  });
}

function chartMarkerStrokeWidth(
  series: PptChartSeries,
  node: PptSceneNode | undefined,
): number | undefined {
  if (!series.marker?.lineWidth || !node?.bounds.width) return undefined;
  return Math.max(0.25, Math.min(1, (series.marker.lineWidth / node.bounds.width) * 100));
}

function chartSeriesLineWidth(series: PptChartSeries, node: PptSceneNode | undefined): number {
  if (!series.lineWidth || !node?.bounds.width) return 0.8;
  return Math.max(0.3, Math.min(1.2, (series.lineWidth / node.bounds.width) * 100));
}

function chartLinePath(points: { x: number; y: number }[], smooth: boolean): string {
  if (points.length === 0) return '';
  if (!smooth || points.length < 3) {
    return points
      .map((point, pointIndex) => `${pointIndex === 0 ? 'M' : 'L'} ${point.x} ${point.y}`)
      .join(' ');
  }
  const segments = [`M ${points[0]!.x} ${points[0]!.y}`];
  for (let i = 0; i < points.length - 1; i++) {
    const previous = points[Math.max(0, i - 1)]!;
    const current = points[i]!;
    const next = points[i + 1]!;
    const afterNext = points[Math.min(points.length - 1, i + 2)]!;
    const control1 = {
      x: current.x + (next.x - previous.x) / 6,
      y: current.y + (next.y - previous.y) / 6,
    };
    const control2 = {
      x: next.x - (afterNext.x - current.x) / 6,
      y: next.y - (afterNext.y - current.y) / 6,
    };
    segments.push(`C ${control1.x} ${control1.y} ${control2.x} ${control2.y} ${next.x} ${next.y}`);
  }
  return segments.join(' ');
}

function renderBarSeries(
  svg: SVGSVGElement,
  seriesList: PptChartSeriesList,
  plot: { left: number; top: number; width: number; height: number },
  min: number,
  max: number,
  node?: PptSceneNode,
  grouping?: string,
): void {
  if (isStackedBarGrouping(grouping) && seriesList.length > 1) {
    renderStackedBarSeries(svg, seriesList, plot, min, max, node);
    return;
  }
  if (shouldOverlayBarSeries(seriesList)) {
    renderOverlayBarSeries(svg, seriesList, plot, min, max, node);
    return;
  }
  const colors = ['#C50D13', '#0EA5E9', '#F59E0B', '#A3E635'];
  const columns = Math.max(...seriesList.map((series) => series.values.length), 1);
  const groupWidth = plot.width / columns;
  const barWidth = Math.max(1, (groupWidth * 0.48) / Math.max(seriesList.length, 1));
  seriesList.forEach((series, seriesIndex) => {
    const fill = chartSeriesFill(svg, series, seriesIndex, colors);
    series.values.forEach((value, valueIndex) => {
      const height = ((value - min) / Math.max(1, max - min)) * plot.height;
      const x =
        plot.left +
        valueIndex * groupWidth +
        (groupWidth - barWidth * seriesList.length) / 2 +
        seriesIndex * barWidth;
      const y = plot.top + plot.height - height;
      appendSvgRect(
        svg,
        x,
        y,
        barWidth,
        height,
        series.pointColors?.[valueIndex] ?? fill,
        series.borderColor,
        chartSeriesBorderWidth(series, node),
      );
    });
  });
}

function renderOverlayBarSeries(
  svg: SVGSVGElement,
  seriesList: PptChartSeriesList,
  plot: { left: number; top: number; width: number; height: number },
  min: number,
  max: number,
  node?: PptSceneNode,
): void {
  const colors = ['#C50D13', '#0EA5E9', '#F59E0B', '#A3E635'];
  const columns = Math.max(...seriesList.map((series) => series.values.length), 1);
  const groupWidth = plot.width / columns;
  const barWidth = Math.max(1, groupWidth * 0.34);
  seriesList.forEach((series, seriesIndex) => {
    const fill = chartSeriesFill(svg, series, seriesIndex, colors);
    series.values.forEach((value, valueIndex) => {
      const y = chartValueToY(value, plot, min, max);
      const baseline = chartValueToY(0, plot, min, max);
      const x = plot.left + valueIndex * groupWidth + (groupWidth - barWidth) / 2;
      appendSvgRect(
        svg,
        x,
        Math.min(y, baseline),
        barWidth,
        Math.max(0.1, Math.abs(baseline - y)),
        series.pointColors?.[valueIndex] ?? fill,
        series.borderColor,
        chartSeriesBorderWidth(series, node),
      );
    });
  });
}

function renderStackedBarSeries(
  svg: SVGSVGElement,
  seriesList: PptChartSeriesList,
  plot: { left: number; top: number; width: number; height: number },
  min: number,
  max: number,
  node?: PptSceneNode,
): void {
  const colors = ['#C50D13', '#0EA5E9', '#F59E0B', '#A3E635'];
  const columns = Math.max(...seriesList.map((series) => series.values.length), 1);
  const groupWidth = plot.width / columns;
  const barWidth = Math.max(1, groupWidth * 0.46);
  const positives = new Array<number>(columns).fill(Math.max(0, min));
  const negatives = new Array<number>(columns).fill(Math.min(0, max));
  seriesList.forEach((series, seriesIndex) => {
    const fill = chartSeriesFill(svg, series, seriesIndex, colors);
    series.values.forEach((value, valueIndex) => {
      const base = value >= 0 ? positives[valueIndex]! : negatives[valueIndex]!;
      const next = base + value;
      const y1 = chartValueToY(base, plot, min, max);
      const y2 = chartValueToY(next, plot, min, max);
      const x = plot.left + valueIndex * groupWidth + (groupWidth - barWidth) / 2;
      appendSvgRect(
        svg,
        x,
        Math.min(y1, y2),
        barWidth,
        Math.max(0.1, Math.abs(y2 - y1)),
        series.pointColors?.[valueIndex] ?? fill,
        series.borderColor,
        chartSeriesBorderWidth(series, node),
      );
      if (value >= 0) positives[valueIndex] = next;
      else negatives[valueIndex] = next;
    });
  });
}

function chartValueToY(
  value: number,
  plot: { top: number; height: number },
  min: number,
  max: number,
): number {
  return plot.top + plot.height - ((value - min) / Math.max(1, max - min)) * plot.height;
}

function isStackedBarGrouping(grouping: string | undefined): boolean {
  return grouping === 'stacked' || grouping === 'percentStacked';
}

function shouldOverlayBarSeries(seriesList: PptChartSeriesList): boolean {
  const groups = new Set(seriesList.map((series) => series.chartGroup ?? 0));
  return (
    groups.size > 1 && seriesList.every((series) => !series.type || series.type.includes('bar'))
  );
}

function chartSeriesBorderWidth(series: PptChartSeries, node?: PptSceneNode): number | undefined {
  if (!series.borderWidth) return undefined;
  return chartStrokeWidth(
    node?.bounds.width ? (series.borderWidth / node.bounds.width) * 100 : undefined,
    undefined,
  );
}

function chartAxisStrokeWidth(
  width: number | undefined,
  node: PptSceneNode,
  fallback: number,
): number {
  return (
    chartStrokeWidth(
      width !== undefined && node.bounds.width ? (width / node.bounds.width) * 100 : undefined,
      fallback,
    ) ?? fallback
  );
}

function chartStrokeWidth(width: number | undefined, fallback: number): number;
function chartStrokeWidth(
  width: number | undefined,
  fallback: number | undefined,
): number | undefined;
function chartStrokeWidth(
  width: number | undefined,
  fallback: number | undefined,
): number | undefined {
  if (width === undefined) return fallback;
  return Math.max(0.15, Math.min(1.4, width));
}

function chartSeriesFill(
  svg: SVGSVGElement,
  series: PptChartSeries,
  index: number,
  fallbackColors: string[],
): string {
  if (series.fillGradient?.colors.length && series.fillGradient.colors.length > 1) {
    const id = `coco-chart-gradient-${index}-${series.fillGradient.colors.join('-').replace(/[^a-zA-Z0-9_-]/g, '_')}`;
    appendSvgLinearGradient(svg, id, series.fillGradient);
    return `url(#${id})`;
  }
  return series.fillColor ?? fallbackColors[index % fallbackColors.length]!;
}

function chartPointFill(
  series: PptChartSeries,
  pointIndex: number,
  fallbackColors: string[],
): string {
  return (
    series.pointColors?.[pointIndex] ??
    series.fillColor ??
    fallbackColors[pointIndex % fallbackColors.length]!
  );
}

function renderCategoryLabels(
  svg: SVGSVGElement,
  categories: string[],
  plot: { left: number; top: number; width: number; height: number },
): void {
  if (categories.length === 0) return;
  const step = plot.width / categories.length;
  categories.forEach((category, index) => {
    appendSvgText(
      svg,
      category,
      plot.left + step * index + step / 2,
      plot.top + plot.height + 5.2,
      3.2,
      '#64748B',
      'middle',
    );
  });
}

function renderValueLabels(
  svg: SVGSVGElement,
  seriesList: PptChartSeriesList,
  plot: { left: number; top: number; width: number; height: number },
  min: number,
  max: number,
): void {
  const columns = Math.max(...seriesList.map((series) => series.values.length), 1);
  const groupWidth = plot.width / columns;
  const barWidth = Math.max(1, (groupWidth * 0.48) / Math.max(seriesList.length, 1));
  seriesList.forEach((series, seriesIndex) => {
    if (!series.showValueLabels) return;
    series.values.forEach((value, valueIndex) => {
      const height = ((value - min) / Math.max(1, max - min)) * plot.height;
      const x =
        plot.left +
        valueIndex * groupWidth +
        (groupWidth - barWidth * seriesList.length) / 2 +
        seriesIndex * barWidth +
        barWidth / 2;
      const y = plot.top + plot.height - height - 2;
      appendSvgText(svg, String(value), x, y, 3, '#A7AEB8', 'middle');
    });
  });
}

function renderLineValueLabels(
  svg: SVGSVGElement,
  seriesList: PptChartSeriesList,
  plot: { left: number; top: number; width: number; height: number },
  min: number,
  max: number,
): void {
  seriesList.forEach((series) => {
    if (!series.showValueLabels || series.values.length === 0) return;
    series.values.forEach((value, pointIndex) => {
      const x =
        plot.left +
        (series.values.length === 1
          ? plot.width / 2
          : (plot.width * pointIndex) / (series.values.length - 1));
      const y = plot.top + plot.height - ((value - min) / Math.max(1, max - min)) * plot.height;
      appendSvgText(
        svg,
        String(value),
        x,
        Math.max(plot.top + 2, y - 2.2),
        2.6,
        '#374151',
        'middle',
      );
    });
  });
}

function renderDoughnutChart(
  svg: SVGSVGElement,
  seriesList: PptChartSeriesList,
  plot: { left: number; top: number; width: number; height: number },
): void {
  const series = seriesList[0];
  if (!series?.values.length) return;
  const colors = ['#1D4ED8', '#E3B363', '#60A5FA', '#94A3B8', '#F59E0B', '#22C7F2'];
  const total = series.values.reduce((sum, value) => sum + Math.max(0, value), 0) || 1;
  const cx = plot.left + plot.width * 0.42;
  const cy = plot.top + plot.height * 0.5;
  const radius = Math.min(plot.width, plot.height) * 0.34;
  const innerRadius = radius * 0.58;
  let angle = -90;
  series.values.forEach((value, index) => {
    const sweep = (Math.max(0, value) / total) * 360;
    const startAngle = angle;
    const endAngle = angle + sweep;
    appendDonutSlice(
      svg,
      cx,
      cy,
      radius,
      innerRadius,
      startAngle,
      endAngle,
      chartPointFill(series, index, colors),
    );
    if (series.showCategoryNameLabels || series.showPercentLabels || series.showValueLabels) {
      const labelLines = doughnutSliceLabelLines(series, value, index, total);
      if (labelLines.length > 0) {
        const point = doughnutLabelPoint(
          series,
          index,
          cx,
          cy,
          radius + 3.6,
          plot,
          (startAngle + endAngle) / 2,
        );
        appendSvgMultilineText(
          svg,
          labelLines,
          point.x,
          point.y,
          1.9,
          '#334155',
          point.x >= cx ? 'start' : 'end',
        );
      }
    }
    angle += sweep;
  });
  appendSvgCircle(svg, cx, cy, innerRadius * 0.92, '#FFFFFF');
  const centerLabel = compactChartCenterLabel(series.name);
  if (centerLabel) appendSvgText(svg, centerLabel, cx, cy - 1.8, 1.8, '#64748B', 'middle');
  const first = series.values[0];
  if (first !== undefined)
    appendSvgText(
      svg,
      formatDoughnutCenterPercent(first, total),
      cx,
      cy + 3.4,
      3.9,
      '#1E3A8A',
      'middle',
    );
  renderLegend(svg, series.categories ?? [], colors, plot.left + plot.width * 0.72, plot.top + 9);
}

function doughnutSliceLabelLines(
  series: PptChartSeries,
  value: number,
  index: number,
  total: number,
): string[] {
  const parts: string[] = [];
  const category = series.categories?.[index];
  if (series.showCategoryNameLabels && category) parts.push(category);
  if (series.showValueLabels) parts.push(formatChartTick(value));
  if (series.showPercentLabels) parts.push(formatDoughnutCenterPercent(value, total));
  return parts;
}

function doughnutLabelPoint(
  series: PptChartSeries,
  index: number,
  cx: number,
  cy: number,
  fallbackRadius: number,
  plot: { width: number; height: number },
  fallbackAngle: number,
): { x: number; y: number } {
  void series;
  void index;
  void plot;
  return polarPoint(cx, cy, fallbackRadius, fallbackAngle);
}

function formatDoughnutCenterPercent(value: number, total: number): string {
  const ratio = value > 0 && value <= 1 && total <= 2 ? value : value / Math.max(1, total);
  const percent = ratio * 100;
  return `${Number.isInteger(percent) ? percent.toFixed(0) : percent.toFixed(1)}%`;
}

function compactChartCenterLabel(value: string | undefined): string {
  const text = (value ?? '').trim();
  if (!text) return '';
  const chars = Array.from(text);
  if (chars.length <= 6) return text;
  return `${chars.slice(0, 5).join('')}…`;
}

function renderRadarChart(
  svg: SVGSVGElement,
  seriesList: PptChartSeriesList,
  plot: { left: number; top: number; width: number; height: number },
  max: number,
): void {
  const series = seriesList[0];
  if (!series?.values.length) return;
  const count = series.values.length;
  const cx = plot.left + plot.width * 0.48;
  const cy = plot.top + plot.height * 0.5;
  const radius = Math.min(plot.width, plot.height) * 0.35;
  for (let ring = 1; ring <= 4; ring++) {
    appendSvgPath(svg, radarPolygon(count, cx, cy, (radius * ring) / 4, true), '#CBD5E1', 0.25);
  }
  for (let index = 0; index < count; index++) {
    const point = radarPoint(index, count, cx, cy, radius);
    appendSvgLine(svg, cx, cy, point.x, point.y, '#E2E8F0', 0.25);
  }
  seriesList.forEach((item, seriesIndex) => {
    if (!item.values.length) return;
    const valuePoints = item.values
      .slice(0, count)
      .map((value, index) =>
        radarPoint(index, count, cx, cy, (Math.max(0, value) / Math.max(1, max)) * radius),
      );
    const area = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    area.setAttribute(
      'd',
      valuePoints
        .map((point, index) => `${index === 0 ? 'M' : 'L'} ${point.x} ${point.y}`)
        .join(' ') + ' Z',
    );
    area.setAttribute('fill', seriesIndex === 0 ? 'rgba(37, 99, 235, 0.08)' : 'none');
    area.setAttribute('stroke', item.lineColor ?? '#2563EB');
    area.setAttribute('stroke-width', String(chartSeriesLineWidth(item, undefined)));
    svg.append(area);
    if (item.marker && item.marker.symbol !== 'none') {
      valuePoints.forEach((point) =>
        appendSvgCircle(svg, point.x, point.y, 0.9, item.lineColor ?? '#2563EB'),
      );
    }
    if (item.showValueLabels) {
      item.values.slice(0, count).forEach((value, index) => {
        const point = valuePoints[index]!;
        appendSvgText(
          svg,
          formatChartTick(value),
          point.x,
          point.y - 1.2,
          2.2,
          item.lineColor ?? '#2563EB',
          'middle',
        );
      });
    }
  });
  (series.categories ?? []).slice(0, count).forEach((category, index) => {
    const point = radarPoint(index, count, cx, cy, radius + 5);
    appendSvgText(svg, category, point.x, point.y + 1, 2.6, '#64748B', 'middle');
  });
}

function appendDonutSlice(
  svg: SVGSVGElement,
  cx: number,
  cy: number,
  outerRadius: number,
  innerRadius: number,
  startAngle: number,
  endAngle: number,
  color: string,
): void {
  const sweep = Math.max(0.01, endAngle - startAngle);
  const largeArc = sweep > 180 ? 1 : 0;
  const startOuter = polarPoint(cx, cy, outerRadius, startAngle);
  const endOuter = polarPoint(cx, cy, outerRadius, endAngle);
  const startInner = polarPoint(cx, cy, innerRadius, endAngle);
  const endInner = polarPoint(cx, cy, innerRadius, startAngle);
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute(
    'd',
    [
      `M ${startOuter.x} ${startOuter.y}`,
      `A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${endOuter.x} ${endOuter.y}`,
      `L ${startInner.x} ${startInner.y}`,
      `A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${endInner.x} ${endInner.y}`,
      'Z',
    ].join(' '),
  );
  path.setAttribute('fill', color);
  svg.append(path);
}

function renderLegend(
  svg: SVGSVGElement,
  labels: string[],
  colors: string[],
  x: number,
  y: number,
): void {
  labels.slice(0, 6).forEach((label, index) => {
    appendSvgRect(svg, x, y + index * 5, 2.4, 2.4, colors[index % colors.length]!);
    appendSvgText(svg, label, x + 3.6, y + index * 5 + 2.2, 2.7, '#334155');
  });
}

function radarPolygon(
  count: number,
  cx: number,
  cy: number,
  radius: number,
  closed: boolean,
): string {
  return (
    Array.from({ length: count }, (_, index) => {
      const point = radarPoint(index, count, cx, cy, radius);
      return `${index === 0 ? 'M' : 'L'} ${point.x} ${point.y}`;
    }).join(' ') + (closed ? ' Z' : '')
  );
}

function radarPoint(
  index: number,
  count: number,
  cx: number,
  cy: number,
  radius: number,
): { x: number; y: number } {
  return polarPoint(cx, cy, radius, -90 + (360 * index) / count);
}

function polarPoint(
  cx: number,
  cy: number,
  radius: number,
  angle: number,
): { x: number; y: number } {
  const radians = (angle * Math.PI) / 180;
  return { x: cx + Math.cos(radians) * radius, y: cy + Math.sin(radians) * radius };
}

function appendSvgLine(
  svg: SVGSVGElement,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  color: string,
  width: number,
  opacity?: number,
  cap?: string,
): void {
  const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  line.setAttribute('x1', String(x1));
  line.setAttribute('y1', String(y1));
  line.setAttribute('x2', String(x2));
  line.setAttribute('y2', String(y2));
  line.setAttribute('stroke', color);
  line.setAttribute('stroke-width', String(width));
  line.setAttribute('stroke-linecap', svgStrokeLineCap(cap));
  if (opacity !== undefined)
    line.setAttribute('stroke-opacity', String(Math.max(0, Math.min(1, opacity))));
  svg.append(line);
}

function svgStrokeLineCap(cap: string | undefined): 'butt' | 'round' | 'square' {
  if (cap === 'rnd') return 'round';
  if (cap === 'sq') return 'square';
  return 'butt';
}

function appendSvgPath(svg: SVGSVGElement, d: string, color: string, width: number): void {
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', d);
  path.setAttribute('fill', 'none');
  path.setAttribute('stroke', color);
  path.setAttribute('stroke-width', String(width));
  path.setAttribute('stroke-linejoin', 'round');
  path.setAttribute('stroke-linecap', 'round');
  svg.append(path);
}

function appendSvgCircle(
  svg: SVGSVGElement,
  cx: number,
  cy: number,
  r: number,
  color: string,
  stroke?: string,
  strokeWidth?: number,
): void {
  const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  circle.setAttribute('cx', String(cx));
  circle.setAttribute('cy', String(cy));
  circle.setAttribute('r', String(r));
  circle.setAttribute('fill', color);
  if (stroke) circle.setAttribute('stroke', stroke);
  if (strokeWidth !== undefined) circle.setAttribute('stroke-width', String(strokeWidth));
  svg.append(circle);
}

function appendSvgArrow(
  svg: SVGSVGElement,
  x: number,
  y: number,
  direction: 'left' | 'right' | 'up' | 'down',
  length: number,
  width: number,
  color: string,
  opacity?: number,
): void {
  const halfWidth = width / 2;
  const points =
    direction === 'right'
      ? [
          [x, y],
          [x - length, y - halfWidth],
          [x - length, y + halfWidth],
        ]
      : direction === 'left'
        ? [
            [x, y],
            [x + length, y - halfWidth],
            [x + length, y + halfWidth],
          ]
        : direction === 'down'
          ? [
              [x, y],
              [x - halfWidth, y - length],
              [x + halfWidth, y - length],
            ]
          : [
              [x, y],
              [x - halfWidth, y + length],
              [x + halfWidth, y + length],
            ];
  const polygon = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
  polygon.classList.add('coco-ppt-line-arrow');
  polygon.setAttribute('points', points.map((point) => point.join(',')).join(' '));
  polygon.setAttribute('fill', color);
  if (opacity !== undefined)
    polygon.setAttribute('fill-opacity', String(Math.max(0, Math.min(1, opacity))));
  svg.append(polygon);
}

function appendSvgRect(
  svg: SVGSVGElement,
  x: number,
  y: number,
  width: number,
  height: number,
  color: string,
  stroke?: string,
  strokeWidth?: number,
): void {
  const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  rect.setAttribute('x', String(x));
  rect.setAttribute('y', String(y));
  rect.setAttribute('width', String(width));
  rect.setAttribute('height', String(height));
  rect.setAttribute('fill', color);
  if (stroke && strokeWidth !== undefined) {
    rect.setAttribute('stroke', stroke);
    rect.setAttribute('stroke-width', String(strokeWidth));
  }
  svg.append(rect);
}

function appendSvgText(
  svg: SVGSVGElement,
  text: string,
  x: number,
  y: number,
  size: number,
  color: string,
  anchor: 'start' | 'middle' | 'end' = 'start',
): void {
  const node = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  node.setAttribute('x', String(x));
  node.setAttribute('y', String(y));
  node.setAttribute('font-size', String(size));
  node.setAttribute('fill', color);
  node.setAttribute('text-anchor', anchor);
  node.textContent = text;
  svg.append(node);
}

function appendSvgMultilineText(
  svg: SVGSVGElement,
  lines: string[],
  x: number,
  y: number,
  size: number,
  color: string,
  anchor: 'start' | 'middle' | 'end' = 'start',
): void {
  lines.forEach((line, index) => {
    appendSvgText(svg, line, x, y + index * size * 1.2, size, color, anchor);
  });
}
