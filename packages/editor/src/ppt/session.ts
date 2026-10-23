import {
  Presentation,
  Shape,
  ShapeCollection,
  type PptPresentationModel,
  type PptShapeGeometry,
  type PptShapeStyle,
  type PptTextStyle,
} from '@coco-vexel/coco-office/ppt';
import {
  buildPptSceneNode,
  buildPptPresentationScene,
  findPptSceneNode,
  findPptSceneSlide,
  pptSceneGeometryToSourceGeometry,
  type PptSceneNode,
  type PptPresentationScene,
} from './scene.js';
import type { PptEditorView } from './view.js';
import { EditorReadOnlyError } from '../read-only.js';

export type PptEditorSelection =
  { kind: 'slide'; slideId: string } | { kind: 'shape'; slideId: string; shapeId: string } | null;

export type PptEditorCommand =
  | { kind: 'select-slide'; slideIdOrIndex: string | number }
  | { kind: 'select-shape'; slideIdOrIndex: string | number; shapeIdOrName: string }
  | { kind: 'set-shape-text'; slideIdOrIndex: string | number; shapeIdOrName: string; text: string }
  | {
      kind: 'set-shape-text-style';
      slideIdOrIndex: string | number;
      shapeIdOrName: string;
      style: Partial<PptTextStyle>;
    }
  | {
      kind: 'set-table-cell-text';
      slideIdOrIndex: string | number;
      shapeIdOrName: string;
      row: number;
      column: number;
      text: string;
    }
  | {
      kind: 'set-shape-geometry';
      slideIdOrIndex: string | number;
      shapeIdOrName: string;
      geometry: Partial<PptShapeGeometry>;
      preserveTextSizeOnResize?: boolean;
    }
  | {
      kind: 'set-shape-style';
      slideIdOrIndex: string | number;
      shapeIdOrName: string;
      style: Partial<PptShapeStyle>;
    }
  | { kind: 'delete-shape'; slideIdOrIndex: string | number; shapeIdOrName: string }
  | {
      kind: 'duplicate-shape';
      slideIdOrIndex: string | number;
      shapeIdOrName: string;
      offset?: Partial<PptShapeGeometry>;
    }
  | {
      kind: 'paste-shape';
      slideIdOrIndex: string | number;
      shapeIdOrName: string;
      offset?: Partial<PptShapeGeometry>;
    }
  | { kind: 'bring-shape-to-front'; slideIdOrIndex: string | number; shapeIdOrName: string }
  | { kind: 'send-shape-to-back'; slideIdOrIndex: string | number; shapeIdOrName: string };

export type PptEditorRuntimeEvent = 'change' | 'selectionChange';

export interface PptEditorChangeEvent {
  readonly revision: number;
  readonly command: PptEditorCommand;
  readonly model: PptPresentationModel;
  readonly scene: PptPresentationScene;
}

export interface PptEditorSelectionChangeEvent {
  readonly selection: PptEditorSelection;
}

export interface PptEditorRuntimeEventMap {
  change: PptEditorChangeEvent;
  selectionChange: PptEditorSelectionChangeEvent;
}

export interface PptEditorSessionOptions {
  maxHistory?: number;
  readOnly?: boolean;
}

type Listener<T> = (event: T) => void;
type PptExistingShapeCommand = Exclude<PptEditorCommand, { kind: 'select-slide' | 'select-shape' }>;
type PptTextBodyOverride = NonNullable<PptSceneNode['textBody']>;
interface PptEditorHistoryEntry {
  readonly bytes: Uint8Array;
  readonly textBodyOverrides: Map<string, PptTextBodyOverride>;
}

export class PptEditorSession {
  private presentationValue: Presentation;
  private modelValue: PptPresentationModel;
  private sceneValue: PptPresentationScene;
  private selectionValue: PptEditorSelection = null;
  private revisionValue = 0;
  private readonly undoStack: PptEditorHistoryEntry[] = [];
  private readonly redoStack: PptEditorHistoryEntry[] = [];
  private readonly textBodyOverrides = new Map<string, PptTextBodyOverride>();
  private readonly maxHistory: number;
  private readOnlyValue: boolean;
  private readonly changeListeners = new Set<Listener<PptEditorChangeEvent>>();
  private readonly selectionListeners = new Set<Listener<PptEditorSelectionChangeEvent>>();

  constructor(view: PptEditorView, options: PptEditorSessionOptions = {}) {
    this.presentationValue = view.presentation;
    this.modelValue = view.model;
    this.sceneValue = view.scene;
    this.maxHistory = options.maxHistory ?? 50;
    this.readOnlyValue = options.readOnly === true;
    this.selectionValue = view.scene.slides[0]
      ? { kind: 'slide', slideId: view.scene.slides[0].id }
      : null;
  }

  presentation(): Presentation {
    this.assertWritable();
    return this.presentationValue;
  }

  model(): PptPresentationModel {
    return this.modelValue;
  }

  scene(): PptPresentationScene {
    return this.sceneValue;
  }

  selection(): PptEditorSelection {
    return this.selectionValue;
  }

  revision(): number {
    return this.revisionValue;
  }

  isReadOnly(): boolean {
    return this.readOnlyValue;
  }

  setReadOnly(readOnly: boolean): void {
    this.readOnlyValue = readOnly;
  }

  canUndo(): boolean {
    return !this.readOnlyValue && this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return !this.readOnlyValue && this.redoStack.length > 0;
  }

  subscribe<K extends PptEditorRuntimeEvent>(
    event: K,
    listener: Listener<PptEditorRuntimeEventMap[K]>,
  ): () => void {
    const listeners = event === 'change' ? this.changeListeners : this.selectionListeners;
    listeners.add(
      listener as Listener<PptEditorChangeEvent> & Listener<PptEditorSelectionChangeEvent>,
    );
    return () =>
      listeners.delete(
        listener as Listener<PptEditorChangeEvent> & Listener<PptEditorSelectionChangeEvent>,
      );
  }

  async dispatch(command: PptEditorCommand): Promise<void> {
    if (command.kind === 'select-slide' || command.kind === 'select-shape') {
      await this.applySelectionCommand(command);
      return;
    }
    this.assertWritable();

    const before = await this.presentationValue.toZip();
    const nextSelection = await this.applyEditCommand(command);
    this.pushUndo(before);
    this.redoStack.length = 0;
    await this.refreshAfterChange(command);
    if (nextSelection !== undefined) this.setSelection(nextSelection);
  }

  async undo(): Promise<void> {
    this.assertWritable();
    const entry = this.undoStack.pop();
    if (!entry) return;
    const selection = this.selectionValue;
    this.redoStack.push(this.historyEntry(await this.presentationValue.toZip()));
    this.restoreTextBodyOverrides(entry.textBodyOverrides);
    await this.replacePresentation(entry.bytes, selection);
  }

  async redo(): Promise<void> {
    this.assertWritable();
    const entry = this.redoStack.pop();
    if (!entry) return;
    const selection = this.selectionValue;
    this.undoStack.push(this.historyEntry(await this.presentationValue.toZip()));
    this.restoreTextBodyOverrides(entry.textBodyOverrides);
    await this.replacePresentation(entry.bytes, selection);
  }

  async loadPresentation(bytes: Uint8Array | ArrayBuffer): Promise<void> {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.textBodyOverrides.clear();
    await this.replacePresentation(bytes);
  }

  async toPptx(): Promise<Uint8Array> {
    return this.presentationValue.toZip();
  }

  async verifyExport(): Promise<PptPresentationModel> {
    const bytes = await this.toPptx();
    const reloaded = await Presentation.fromZip(bytes, { mode: 'inline' });
    try {
      return await reloaded.toModel();
    } finally {
      await reloaded.dispose();
    }
  }

  async destroy(): Promise<void> {
    await this.presentationValue.dispose();
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.changeListeners.clear();
    this.selectionListeners.clear();
  }

  private assertWritable(): void {
    if (this.readOnlyValue) throw new EditorReadOnlyError('ppt');
  }

  private async applySelectionCommand(
    command: Extract<PptEditorCommand, { kind: 'select-slide' | 'select-shape' }>,
  ): Promise<void> {
    const slide = findPptSceneSlide(this.sceneValue, command.slideIdOrIndex);
    if (!slide) throw new Error(`PPT slide not found: ${String(command.slideIdOrIndex)}`);
    if (command.kind === 'select-slide') {
      this.setSelection({ kind: 'slide', slideId: slide.id });
      return;
    }
    const shape = findPptSceneNode(slide, command.shapeIdOrName);
    if (!shape) throw new Error(`PPT shape not found: ${command.shapeIdOrName}`);
    this.setSelection({ kind: 'shape', slideId: slide.id, shapeId: shape.id });
  }

  private async applyEditCommand(
    command: Exclude<PptEditorCommand, { kind: 'select-slide' | 'select-shape' }>,
  ): Promise<PptEditorSelection | undefined> {
    if (
      command.kind === 'delete-shape' ||
      command.kind === 'duplicate-shape' ||
      command.kind === 'paste-shape' ||
      command.kind === 'bring-shape-to-front' ||
      command.kind === 'send-shape-to-back'
    ) {
      return this.applyShapeCollectionCommand(command);
    }
    const sceneNode = this.assertEditableSceneNodeForCommand(command);
    const sceneSlide = findPptSceneSlide(this.sceneValue, command.slideIdOrIndex);
    if (!sceneSlide) throw new Error(`PPT slide not found: ${String(command.slideIdOrIndex)}`);
    const slide = this.presentationValue.slides.getItem(command.slideIdOrIndex);
    const shape = resolvePptRuntimeShape(slide.shapes, command.shapeIdOrName);
    if (command.kind === 'set-shape-text') {
      await shape.text.setText(command.text, { preserveFirstRunStyle: true });
      return undefined;
    }
    if (command.kind === 'set-shape-text-style') {
      await shape.text.setTextStyle(command.style);
      return undefined;
    }
    if (command.kind === 'set-table-cell-text') {
      await shape.table.setCellText(command.row, command.column, command.text);
      return undefined;
    }
    if (command.kind === 'set-shape-geometry') {
      if (command.preserveTextSizeOnResize && shouldUseWrappingAfterTextResize(sceneNode)) {
        const body = {
          ...sceneNode.textBody,
          autoFit: 'none' as const,
          wrap: sceneNode.textBody?.wrap ?? ('square' as const),
        };
        await shape.text.setBodyProperties(body);
        this.textBodyOverrides.set(textBodyOverrideKey(sceneSlide.id, sceneNode.id), body);
      }
      await shape.setGeometry(pptSceneGeometryToSourceGeometry(sceneNode, command.geometry));
      return undefined;
    }
    await shape.setStyle(command.style);
    return undefined;
  }

  private assertEditableSceneNodeForCommand(command: PptExistingShapeCommand): PptSceneNode {
    const slide = findPptSceneSlide(this.sceneValue, command.slideIdOrIndex);
    if (!slide) throw new Error(`PPT slide not found: ${String(command.slideIdOrIndex)}`);
    const shape =
      findPptSceneNode(slide, command.shapeIdOrName) ??
      editableTopLevelSceneNode(slide, command.shapeIdOrName);
    if (!shape) throw new Error(`PPT shape not found: ${command.shapeIdOrName}`);
    if (
      isShapeCollectionCommand(command) &&
      (shape.locked || shape.id.includes(':') || !shape.editable)
    ) {
      throw new Error(`PPT shape collection operation is not editable: ${shape.id}`);
    }
    if (command.kind === 'set-shape-text' && !shape.capabilities.editText) {
      throw new Error(`PPT shape is not text-editable: ${shape.id}`);
    }
    if (command.kind === 'set-shape-text-style' && !shape.capabilities.editTextStyle) {
      throw new Error(`PPT shape text style is not editable: ${shape.id}`);
    }
    if (command.kind === 'set-table-cell-text' && !shape.capabilities.editTableCells) {
      throw new Error(`PPT shape table cells are not editable: ${shape.id}`);
    }
    if (command.kind === 'set-shape-geometry' && !shape.capabilities.editGeometry) {
      throw new Error(`PPT shape geometry is not editable: ${shape.id}`);
    }
    if (command.kind === 'set-shape-style' && !shape.capabilities.editShapeStyle) {
      throw new Error(`PPT shape style is not editable: ${shape.id}`);
    }
    return shape;
  }

  private async applyShapeCollectionCommand(
    command: Extract<
      PptEditorCommand,
      {
        kind:
          | 'delete-shape'
          | 'duplicate-shape'
          | 'paste-shape'
          | 'bring-shape-to-front'
          | 'send-shape-to-back';
      }
    >,
  ): Promise<PptEditorSelection> {
    this.assertEditableSceneNodeForCommand(command);
    const sceneSlide = findPptSceneSlide(this.sceneValue, command.slideIdOrIndex);
    if (!sceneSlide) throw new Error(`PPT slide not found: ${String(command.slideIdOrIndex)}`);
    const sceneNode = findPptSceneNode(sceneSlide, command.shapeIdOrName);
    if (!sceneNode) throw new Error(`PPT shape not found: ${command.shapeIdOrName}`);
    const slide = this.presentationValue.slides.getItem(command.slideIdOrIndex);
    const shape = slide.shapes.getItem(command.shapeIdOrName);
    if (command.kind === 'delete-shape') {
      await shape.delete();
      return { kind: 'slide', slideId: sceneSlide.id };
    }
    if (command.kind === 'bring-shape-to-front') {
      await shape.bringToFront();
      return { kind: 'shape', slideId: sceneSlide.id, shapeId: sceneNode.id };
    }
    if (command.kind === 'send-shape-to-back') {
      await shape.sendToBack();
      return { kind: 'shape', slideId: sceneSlide.id, shapeId: sceneNode.id };
    }

    const duplicateGeometry = {
      left:
        typeof sceneNode.bounds.left === 'number'
          ? sceneNode.bounds.left + (command.offset?.left ?? 180000)
          : sceneNode.bounds.left,
      top:
        typeof sceneNode.bounds.top === 'number'
          ? sceneNode.bounds.top + (command.offset?.top ?? 180000)
          : sceneNode.bounds.top,
      width: command.offset?.width ?? sceneNode.bounds.width,
      height: command.offset?.height ?? sceneNode.bounds.height,
    };
    const duplicated = await shape.duplicate({ geometry: duplicateGeometry });
    const duplicatedModel = await duplicated.toModel();
    return { kind: 'shape', slideId: sceneSlide.id, shapeId: duplicatedModel.id };
  }

  private async refreshAfterChange(command: PptEditorCommand): Promise<void> {
    this.modelValue = await this.presentationValue.toModel();
    this.sceneValue = buildPptPresentationScene(this.modelValue);
    this.applyTextBodyOverrides();
    this.revisionValue += 1;
    this.changeListeners.forEach((listener) =>
      listener({
        revision: this.revisionValue,
        command,
        model: this.modelValue,
        scene: this.sceneValue,
      }),
    );
  }

  private async replacePresentation(
    bytes: Uint8Array | ArrayBuffer,
    preferredSelection: PptEditorSelection = null,
  ): Promise<void> {
    const next = await Presentation.fromZip(bytes, { mode: 'inline' });
    await this.presentationValue.dispose();
    this.presentationValue = next;
    this.modelValue = await next.toModel();
    this.sceneValue = buildPptPresentationScene(this.modelValue);
    this.applyTextBodyOverrides();
    if (preferredSelection) {
      const slide = this.sceneValue.slides.find((item) => item.id === preferredSelection.slideId);
      if (preferredSelection.kind === 'slide' && slide) {
        this.setSelection({ kind: 'slide', slideId: slide.id });
        return;
      }
      if (
        preferredSelection.kind === 'shape' &&
        slide?.nodes.some((node) => node.id === preferredSelection.shapeId)
      ) {
        this.setSelection({
          kind: 'shape',
          slideId: slide.id,
          shapeId: preferredSelection.shapeId,
        });
        return;
      }
    }
    const firstSlide = this.sceneValue.slides[0];
    this.setSelection(firstSlide ? { kind: 'slide', slideId: firstSlide.id } : null);
  }

  private setSelection(selection: PptEditorSelection): void {
    this.selectionValue = selection;
    this.selectionListeners.forEach((listener) => listener({ selection }));
  }

  private pushUndo(bytes: Uint8Array): void {
    this.undoStack.push(this.historyEntry(bytes));
    if (this.undoStack.length > this.maxHistory) this.undoStack.shift();
  }

  private historyEntry(bytes: Uint8Array): PptEditorHistoryEntry {
    return {
      bytes,
      textBodyOverrides: new Map(this.textBodyOverrides),
    };
  }

  private restoreTextBodyOverrides(overrides: Map<string, PptTextBodyOverride>): void {
    this.textBodyOverrides.clear();
    overrides.forEach((body, key) => this.textBodyOverrides.set(key, { ...body }));
  }

  private applyTextBodyOverrides(): void {
    for (const [key, body] of this.textBodyOverrides) {
      const [slideId, shapeId] = key.split('\u0000');
      if (!slideId || !shapeId) continue;
      const slide = findPptSceneSlide(this.sceneValue, slideId);
      const node = slide ? findPptSceneNode(slide, shapeId) : null;
      if (!node || (!node.text && !node.textParagraphs?.length)) continue;
      node.textBody = {
        ...node.textBody,
        ...body,
      };
    }
  }
}

function textBodyOverrideKey(slideId: string, shapeId: string): string {
  return `${slideId}\u0000${shapeId}`;
}

function shouldUseWrappingAfterTextResize(node: PptSceneNode): boolean {
  return Boolean(
    node.capabilities.editGeometry &&
      isResizeWrappingTextShape(node) &&
      (node.text || node.textParagraphs?.length),
  );
}

function isResizeWrappingTextShape(node: PptSceneNode): boolean {
  return node.kind === 'placeholder' || node.kind === 'textBox' || node.kind === 'geometricShape';
}

function editableTopLevelSceneNode(
  slide: PptPresentationScene['slides'][number],
  idOrName: string,
): PptSceneNode | null {
  const source = slide.source.shapes.find(
    (shape) => shape.id === idOrName || shape.name === idOrName,
  );
  return source
    ? buildPptSceneNode(source, { slideSize: slide.size, background: slide.background })
    : null;
}

function resolvePptRuntimeShape(shapes: ShapeCollection, idOrName: string): Shape {
  const path = idOrName.split(':').filter(Boolean);
  if (path.length <= 1) return shapes.getItem(idOrName);
  let shape = shapes.getItem(path[0]!);
  for (const id of path.slice(1)) {
    shape = shape.group.getItem(id);
  }
  return shape;
}

function isShapeCollectionCommand(command: PptExistingShapeCommand): command is Extract<
  PptExistingShapeCommand,
  {
    kind:
      | 'delete-shape'
      | 'duplicate-shape'
      | 'paste-shape'
      | 'bring-shape-to-front'
      | 'send-shape-to-back';
  }
> {
  return (
    command.kind === 'delete-shape' ||
    command.kind === 'duplicate-shape' ||
    command.kind === 'paste-shape' ||
    command.kind === 'bring-shape-to-front' ||
    command.kind === 'send-shape-to-back'
  );
}
