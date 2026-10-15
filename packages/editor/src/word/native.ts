import type { EditorDocumentView, EditorRunFormat } from './model.js';
import defaultEditorStyles from '../generated/editor-styles.js';
import {
  allNativeParagraphs,
  createNativeWordDocument,
  findNativeParagraph,
  findNativeParagraphLocation,
  findNativeTableCellLocation,
  findNativeTableLocation,
  nativeParagraphText,
  previousNativeParagraphId,
  type NativeWordBlockNode,
  type NativeWordDocument,
  type NativeWordInlineNode,
  type NativeWordParagraphNode,
  type NativeWordTableCellNode,
} from './native-model.js';
import {
  nativeDeleteBackward,
  nativeDeleteForward,
  nativeDeleteImage,
  nativeDeleteRange,
  nativeDeleteTextRangeAcrossParagraphs,
  nativeClearTableCellRange,
  nativeCopyBlockRange,
  nativeClearRunFormat,
  cloneNativeWordBlocks,
  nativeInsertText,
  nativeInsertImage,
  nativeInsertPageBreak,
  nativePageBreakAt,
  nativeRangeContainsPageBreak,
  nativeEditTableStructure,
  nativeDeleteTable,
  nativeInsertTable,
  nativeReplaceRange,
  nativeReplaceBlockRangeAcrossParagraphs,
  nativeReplaceRangeWithBlocks,
  nativeReplaceTextRangeAcrossParagraphs,
  nativeReplaceTextRangeWithParagraphs,
  nativeReplaceTableCellRangeText,
  nativeReplaceTableCellRangeContent,
  nativeSetParagraphFormat,
  nativeSetParagraphList,
  nativeSetImageSize,
  nativeSetHyperlink,
  nativeAddComment,
  nativeUpdateComment,
  nativeDeleteComment,
  nativeAddRevision,
  nativeResolveRevision,
  nativeResolveImageRevision,
  nativeResolveImageFormatRevision,
  nativeAddFormatRevision,
  nativeResolveFormatRevision,
  nativeSetTableCellFormat,
  nativeSetTableCellRangeFormat,
  nativeTableCellsInRange,
  nativeMergeTableCells,
  nativeSplitTableCell,
  type NativeTableClipboardCell,
  nativeSetRunFormat,
  nativeSplitParagraph,
  nativeSyncParagraphText,
  type NativeWordChange,
} from './native-commands.js';
import {
  defineNativeWordElements,
  measureNativeDocumentReflowFromParagraph,
  paginateNativeDocument,
  rebuildNativeTableFragmentsFromParagraph,
  reflowNativeDocumentFromParagraph,
  renderNativeBlock,
  renderNativeDocument,
  renderNativeImage,
  renderNativeParagraph,
  refreshNativePageCommentBubbles,
  type NativeDocumentReflowMeasurement,
} from './native-renderer.js';
import {
  clearNativeImageSelection,
  createNativeSelectionRange,
  createNativeTextRange,
  currentNativeSelection,
  currentNativeCursor,
  imageOffset,
  nativeCursorFromPoint,
  nativeSelectionFromCursor,
  nativeSelection,
  parentNativeParagraph,
  restoreNativeCursor,
  restoreNativeTextSelection,
  restoreNativeElementSelection,
  selectNativeImage,
  type NativeCursorPoint,
  type NativeSelectedImage,
} from './native-selection.js';
import { canDispatchNativeWordCommand, getNativeWordCommandState } from './native-command-state.js';
import { searchNativeWordDocument } from './native-search.js';
import { createWordEditorChangeEvent } from './runtime-events.js';
import {
  NATIVE_BLOCK_CLIPBOARD_TYPE,
  nativeBlockClipboardText,
  parseNativeBlockClipboard,
  parseNativeClipboardHtml,
  serializeNativeBlockClipboard,
  serializeNativeClipboardHtml,
  type NativeClipboardParagraph,
} from './native-clipboard.js';
import {
  adjacentNativeListNumbering,
  createNativeListNumbering,
  nativeListType,
  nativeNumberingAtLevel,
  refreshNativeNumberingLabels,
} from './native-numbering.js';
import type {
  WordEditorCommand,
  WordEditorChangeEvent,
  WordEditorRuntime,
  WordEditorRuntimeEvent,
  WordEditorRuntimeEventMap,
  WordEditorRuntimeOptions,
  WordEditorSelection,
  WordEditorSelectionChangeEvent,
  WordEditorSearchOptions,
  WordEditorSearchState,
  WordEditorTableCellPoint,
} from './runtime.js';
import {
  embeddedRasterImageBlob,
  safeEmbeddedRasterImageUrl,
} from './native-url-safety.js';

export * from './native-model.js';
export * from './native-table-pagination.js';
export * from './native-commands.js';
export * from './native-renderer.js';
export * from './native-selection.js';
export * from './native-search.js';
export * from './native-command-state.js';
export * from './native-numbering.js';
export * from './ooxml-patch.js';

export interface NativeWordEditor extends WordEditorRuntime {}

const activeNativeWordHostByDocument = new WeakMap<Document, HTMLElement>();
const NATIVE_TABLE_CLIPBOARD_TYPE = 'application/x-coco-word-table-cells+json';
const NATIVE_IMAGE_CLIPBOARD_TYPE = 'application/x-coco-word-image';
const NATIVE_IMAGE_CLIPBOARD_PREFIX = 'coco-image:';
const nativeImageClipboardSources = new Map<string, string>();
let nextNativeImageClipboardId = 0;
let nextNativeSearchHighlightId = 0;
let nextNativeCommentId = 0;
  let nextNativeRevisionId = 0;

interface PendingBrowserRangeDelete {
  range: {
    anchor: { paraId: string; offset: number };
    focus: { paraId: string; offset: number };
  };
  beforeParagraphs: NativeWordParagraphNode[];
  changesStart: number;
}

interface BrowserTextTransaction extends PendingBrowserRangeDelete {
  source: 'browser' | 'local';
  patchMode?: 'range' | 'individual';
  skipDeferredRender?: boolean;
  forceFullRender?: boolean;
  deferLayoutReflow?: boolean;
  afterParagraphs: NativeWordParagraphNode[];
  changes: NativeWordChange[];
  beforeSelection?: WordEditorSelection;
  afterSelection?: WordEditorSelection;
}

interface BrowserBlockTransaction {
  source: 'local';
  range: PendingBrowserRangeDelete['range'];
  container: NativeWordBlockNode[];
  startIndex: number;
  beforeBlocks: NativeWordBlockNode[];
  afterBlocks: NativeWordBlockNode[];
  changesStart: number;
  changes: NativeWordChange[];
  beforeSelection?: WordEditorSelection;
  afterSelection?: WordEditorSelection;
}

interface PendingCompositionTransaction {
  paraId: string;
  beforeParagraph: NativeWordParagraphNode;
  changesStart: number;
  range: PendingBrowserRangeDelete['range'];
}

interface NativeContextMenuItem {
  label: string;
  disabled?: boolean;
  separatorBefore?: boolean;
  action(): unknown;
}

type BrowserTransaction = BrowserTextTransaction | BrowserBlockTransaction;

export function renderNativeWordEditor(
  host: HTMLElement,
  initialView: EditorDocumentView,
  options: WordEditorRuntimeOptions = {},
): NativeWordEditor {
  defineNativeWordElements();
  const outerHost = host;
  let shadowRoot: ShadowRoot | null = null;
  if (options.styleIsolation !== 'none') {
    outerHost.replaceChildren();
    shadowRoot = outerHost.shadowRoot ?? outerHost.attachShadow({ mode: 'open' });
    shadowRoot.replaceChildren();
    const style = outerHost.ownerDocument.createElement('style');
    style.textContent = `
      :host { display: block; position: relative; min-width: 0; min-height: 0; }
      *, *::before, *::after { box-sizing: border-box; }
      button, input, textarea, select { font: inherit; line-height: normal; }
      button { margin: 0; }
    ${options.styleText ?? defaultEditorStyles}`;
    shadowRoot.appendChild(style);
    const surface = outerHost.ownerDocument.createElement('div');
    surface.className = 'coco-word-editor coco-word-native-editor-surface';
    surface.setAttribute('part', 'surface');
    shadowRoot.appendChild(surface);
    outerHost.classList.add('coco-word-editor-host');
    outerHost.style.display = 'block';
    outerHost.style.position = 'relative';
    outerHost.style.minWidth = '0';
    outerHost.style.minHeight = '0';
    host = surface;
  }
  const focusHost = outerHost;
  host.classList.add('coco-word-native-editor-host');
  let sourceView = initialView;
  let doc = createNativeWordDocument(sourceView);
  let destroyed = false;
  const changes: NativeWordChange[] = [];
  let emittedChanges: NativeWordChange[] = [];
  let revision = 0;
  const paragraphIds = new Set(allNativeParagraphs(doc).map((paragraph) => paragraph.paraId));
  let nextParagraphId = Date.now() >>> 0;
  let selectedImage: NativeSelectedImage | null = null;
  let selection: WordEditorSelection | null = null;
  let lastTextSelection: Extract<WordEditorSelection, { kind: 'text-range' }> | null = null;
  let trackRevisions = options.trackRevisions === true;
  let searchState: WordEditorSearchState | null = null;
  let suppressNextEnterBeforeInput = false;
  let clearEnterSuppressionTimer: number | null = null;
  let deferredRenderTimer: number | null = null;
  let deferredReflowTimer: number | null = null;
  let nativeHistorySyncTimer: number | null = null;
  let deferredChangeTimer: number | null = null;
  let deferredSelectionTimer: number | null = null;
  let pendingBrowserRangeDelete: PendingBrowserRangeDelete | null = null;
  let pendingCompositionTransaction: PendingCompositionTransaction | null = null;
  let compositionActive = false;
  let typingLayoutBaseline: { paraId: string; measurement: NativeDocumentReflowMeasurement } | null = null;
  let compositionFinalizeTimer: number | null = null;
  let suppressedBrowserHistoryInput: 'historyUndo' | 'historyRedo' | null = null;
  let browserHistorySuppressionTimer: number | null = null;
  let activeLocalTypingTransaction: BrowserTextTransaction | null = null;
  const browserUndoStack: BrowserTransaction[] = [];
  const browserRedoStack: BrowserTransaction[] = [];
  const pendingImageReflows = new WeakSet<HTMLImageElement>();
  let textDragActive = false;
  let textDragMoved = false;
  let suppressTextClick = false;
  let preserveSelectionOnClick = false;
  let preserveSelectionForExternalInteraction = false;
  const listeners: Record<
    WordEditorRuntimeEvent,
    Set<(payload: WordEditorChangeEvent | WordEditorSelectionChangeEvent) => void>
  > = {
    change: new Set(),
    selectionChange: new Set(),
  };

  const emit = <Event extends WordEditorRuntimeEvent>(event: Event, payload: WordEditorRuntimeEventMap[Event]) => {
    for (const listener of listeners[event]) listener(payload);
  };

  const activate = () => activeNativeWordHostByDocument.set(document, host);

  const getCommandState = (current: WordEditorSelection | null) => ({
    ...getNativeWordCommandState(doc, current),
    trackRevisions,
  });

  let imageResizeHandle: HTMLButtonElement | null = null;
  let cancelImageResizeDrag: (() => void) | null = null;
  let contextMenuElement: HTMLDivElement | null = null;
  let tableDragStart: { tableId: string; point: WordEditorTableCellPoint } | null = null;
  let tableDragSelectionActive = false;
  let suppressTableSelectionClick = false;
  const closeContextMenu = () => {
    contextMenuElement?.remove();
    contextMenuElement = null;
  };
  const tableCellTarget = (target: EventTarget | null): {
    tableId: string;
    point: WordEditorTableCellPoint;
    element: HTMLTableCellElement;
  } | null => {
    if (!(target instanceof Element)) return null;
    const cell = target.closest<HTMLTableCellElement>('coco-word-native-table td[data-row-index][data-column-index]');
    const table = cell?.closest<HTMLElement>('coco-word-native-table[data-table-id]');
    const rowIndex = Number(cell?.dataset.rowIndex);
    const columnIndex = Number(cell?.dataset.columnIndex);
    if (!cell || !table || !Number.isInteger(rowIndex) || !Number.isInteger(columnIndex)) return null;
    return { tableId: table.dataset.tableId ?? '', point: { rowIndex, columnIndex }, element: cell };
  };
  const renderTableCellRangeSelection = (current: WordEditorSelection | null) => {
    for (const cell of Array.from(host.querySelectorAll<HTMLTableCellElement>('td.is-cell-range-selected'))) {
      cell.classList.remove('is-cell-range-selected', 'is-cell-range-anchor', 'is-cell-range-focus');
      cell.removeAttribute('aria-selected');
    }
    if (current?.kind !== 'table-cell-range') return;
    const table = Array.from(host.querySelectorAll<HTMLElement>('coco-word-native-table[data-table-id]'))
      .find((candidate) => candidate.dataset.tableId === current.tableId);
    if (!table) return;
    const startRow = Math.min(current.anchor.rowIndex, current.focus.rowIndex);
    const endRow = Math.max(current.anchor.rowIndex, current.focus.rowIndex);
    const startColumn = Math.min(current.anchor.columnIndex, current.focus.columnIndex);
    const endColumn = Math.max(current.anchor.columnIndex, current.focus.columnIndex);
    for (const cell of Array.from(table.querySelectorAll<HTMLTableCellElement>('td[data-row-index][data-column-index]'))) {
      const row = Number(cell.dataset.rowIndex);
      const column = Number(cell.dataset.columnIndex);
      const rowEnd = row + Math.max(1, cell.rowSpan) - 1;
      const columnEnd = column + Math.max(1, cell.colSpan) - 1;
      if (row <= endRow && rowEnd >= startRow && column <= endColumn && columnEnd >= startColumn) {
        cell.classList.add('is-cell-range-selected');
        cell.setAttribute('aria-selected', 'true');
        if (row === current.anchor.rowIndex && column === current.anchor.columnIndex) {
          cell.classList.add('is-cell-range-anchor');
        }
        if (row === current.focus.rowIndex && column === current.focus.columnIndex) {
          cell.classList.add('is-cell-range-focus');
        }
      }
    }
  };
  const searchHighlightId = ++nextNativeSearchHighlightId;
  const searchMatchHighlightName = `coco-word-search-match-${searchHighlightId}`;
  const searchActiveHighlightName = `coco-word-search-active-${searchHighlightId}`;
  const preservedSelectionHighlightName = `coco-word-preserved-selection-${searchHighlightId}`;
  const searchHighlightStyle = document.createElement('style');
  searchHighlightStyle.textContent = `
    ::highlight(${searchMatchHighlightName}) { background-color: rgb(250 204 21 / 42%); }
    ::highlight(${searchActiveHighlightName}) { background-color: rgb(249 115 22 / 68%); color: inherit; }
    ::highlight(${preservedSelectionHighlightName}) { background-color: rgb(0 120 215 / 34%); color: inherit; }
  `;
  (shadowRoot ?? document.head).appendChild(searchHighlightStyle);
  const searchHighlightRegistry = (globalThis.CSS as (typeof CSS & {
    highlights?: { set(name: string, highlight: unknown): void; delete(name: string): void };
  }) | undefined)?.highlights;
  const SearchHighlight = (globalThis as typeof globalThis & {
    Highlight?: new (...ranges: Range[]) => unknown;
  }).Highlight;
  const clearSearchHighlights = () => {
    searchHighlightRegistry?.delete(searchMatchHighlightName);
    searchHighlightRegistry?.delete(searchActiveHighlightName);
  };
  const clearPreservedSelectionHighlight = () => {
    searchHighlightRegistry?.delete(preservedSelectionHighlightName);
  };
  const renderPreservedSelectionHighlight = (force = false) => {
    clearPreservedSelectionHighlight();
    if (
      selection?.kind !== 'text-range' ||
      !force && (document.activeElement === focusHost || host.contains(document.activeElement)) ||
      !searchHighlightRegistry ||
      !SearchHighlight
    ) return;
    const ordered = orderedSelectionRange(doc, selection);
    const range = createNativeSelectionRange(host, ordered.anchor, ordered.focus);
    if (range && !range.collapsed) {
      searchHighlightRegistry.set(preservedSelectionHighlightName, new SearchHighlight(range));
    }
  };
  const renderSearchHighlights = () => {
    clearSearchHighlights();
    if (!searchState || !searchHighlightRegistry || !SearchHighlight) return;
    const regularRanges: Range[] = [];
    let activeRange: Range | null = null;
    searchState.matches.forEach((match, index) => {
      const range = createNativeTextRange(host, match.paraId, match.from, match.to);
      if (!range) return;
      if (index === searchState!.activeIndex) activeRange = range;
      else regularRanges.push(range);
    });
    if (regularRanges.length > 0) {
      searchHighlightRegistry.set(searchMatchHighlightName, new SearchHighlight(...regularRanges));
    }
    if (activeRange) {
      searchHighlightRegistry.set(searchActiveHighlightName, new SearchHighlight(activeRange));
    }
  };
  const searchStateSnapshot = (): WordEditorSearchState | null => searchState
    ? { options: { ...searchState.options }, matches: searchState.matches.map((match) => ({ ...match })), activeIndex: searchState.activeIndex }
    : null;
  const refreshActiveSearch = () => {
    if (!searchState) return;
    const matches = searchNativeWordDocument(doc, searchState.options);
    searchState = {
      options: { ...searchState.options },
      matches,
      activeIndex: matches.length === 0 ? -1 : Math.min(Math.max(0, searchState.activeIndex), matches.length - 1),
    };
    renderSearchHighlights();
    renderPreservedSelectionHighlight();
  };
  const scrollToActiveSearchMatch = () => {
    const match = searchState && searchState.activeIndex >= 0 ? searchState.matches[searchState.activeIndex] : null;
    const paragraph = match
      ? host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${match.paraId}"]`)
      : null;
    paragraph?.scrollIntoView({ block: 'center', inline: 'nearest' });
  };
  const search = (options: WordEditorSearchOptions): WordEditorSearchState => {
    const normalized = { query: options.query, matchCase: !!options.matchCase, wholeWord: !!options.wholeWord };
    const matches = searchNativeWordDocument(doc, normalized);
    searchState = { options: normalized, matches, activeIndex: matches.length > 0 ? 0 : -1 };
    renderSearchHighlights();
    scrollToActiveSearchMatch();
    return searchStateSnapshot()!;
  };
  const findNext = (direction: 'forward' | 'backward' = 'forward'): WordEditorSearchState | null => {
    if (!searchState || searchState.matches.length === 0) return searchStateSnapshot();
    const delta = direction === 'backward' ? -1 : 1;
    searchState = {
      ...searchState,
      activeIndex: (searchState.activeIndex + delta + searchState.matches.length) % searchState.matches.length,
    };
    renderSearchHighlights();
    scrollToActiveSearchMatch();
    return searchStateSnapshot();
  };
  const clearSearch = () => {
    searchState = null;
    clearSearchHighlights();
  };
  const removeImageResizeHandle = () => {
    imageResizeHandle?.remove();
    imageResizeHandle = null;
  };
  const positionImageResizeHandle = () => {
    if (!imageResizeHandle || !selectedImage) return;
    const image = host.querySelector<HTMLImageElement>(
      `coco-word-native-p[data-para-id="${selectedImage.paraId}"] img[data-drawing-id="${selectedImage.drawingId}"]`,
    );
    if (!image) {
      removeImageResizeHandle();
      return;
    }
    const hostRect = host.getBoundingClientRect();
    const imageRect = image.getBoundingClientRect();
    imageResizeHandle.style.left = `${imageRect.right - hostRect.left + host.scrollLeft - 6}px`;
    imageResizeHandle.style.top = `${imageRect.bottom - hostRect.top + host.scrollTop - 6}px`;
  };
  const showImageResizeHandle = () => {
    if (!selectedImage) {
      removeImageResizeHandle();
      return;
    }
    selectNativeImage(host, selectedImage);
    const image = host.querySelector<HTMLImageElement>(
      `coco-word-native-p[data-para-id="${selectedImage.paraId}"] img[data-drawing-id="${selectedImage.drawingId}"]`,
    );
    if (!image) return;
    if (!imageResizeHandle?.isConnected) {
      const handle = document.createElement('button');
      handle.type = 'button';
      handle.className = 'coco-word-image-resize-handle';
      handle.title = 'Resize image';
      handle.setAttribute('aria-label', 'Resize image');
      handle.contentEditable = 'false';
      handle.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
      });
      handle.addEventListener('pointerdown', (event) => {
        if (!selectedImage || event.button !== 0) return;
        activate();
        const selected = { ...selectedImage };
        const target = host.querySelector<HTMLImageElement>(
          `coco-word-native-p[data-para-id="${selected.paraId}"] img[data-drawing-id="${selected.drawingId}"]`,
        );
        if (!target) return;
        cancelImageResizeDrag?.();
        event.preventDefault();
        event.stopPropagation();
        handle.setPointerCapture(event.pointerId);
        const rect = target.getBoundingClientRect();
        const startX = event.clientX;
        const startY = event.clientY;
        const startWidth = Math.max(1, rect.width);
        const startHeight = Math.max(1, rect.height);
        const aspectRatio = startWidth / startHeight;
        let width = startWidth;
        let height = startHeight;
        const pointerMove = (moveEvent: PointerEvent) => {
          const deltaX = moveEvent.clientX - startX;
          const deltaY = moveEvent.clientY - startY;
          if (moveEvent.shiftKey) {
            width = Math.max(24, startWidth + deltaX);
            height = Math.max(24, startHeight + deltaY);
          } else {
            const dominantDelta = Math.abs(deltaX) >= Math.abs(deltaY) ? deltaX : deltaY * aspectRatio;
            width = Math.max(24, startWidth + dominantDelta);
            height = Math.max(24, width / aspectRatio);
          }
          target.style.width = `${width}px`;
          target.style.height = `${height}px`;
          positionImageResizeHandle();
        };
        const pointerUp = (upEvent: PointerEvent) => {
          cancelImageResizeDrag?.();
          cancelImageResizeDrag = null;
          if (handle.hasPointerCapture(upEvent.pointerId)) handle.releasePointerCapture(upEvent.pointerId);
          dispatch({
            kind: 'set-image-size',
            objectId: selected.drawingId,
            width: Math.round(width * 100) / 100,
            height: Math.round(height * 100) / 100,
          });
        };
        document.addEventListener('pointermove', pointerMove, true);
        document.addEventListener('pointerup', pointerUp, true);
        document.addEventListener('pointercancel', pointerUp, true);
        cancelImageResizeDrag = () => {
          document.removeEventListener('pointermove', pointerMove, true);
          document.removeEventListener('pointerup', pointerUp, true);
          document.removeEventListener('pointercancel', pointerUp, true);
        };
      });
      host.appendChild(handle);
      imageResizeHandle = handle;
    }
    positionImageResizeHandle();
  };

  const repositionSelectedImageHandle = () => positionImageResizeHandle();
  window.addEventListener('resize', repositionSelectedImageHandle);
  document.addEventListener('scroll', repositionSelectedImageHandle, true);

  const render = (cursor?: NativeCursorPoint, preservedSelection?: WordEditorSelection) => {
    if (deferredRenderTimer !== null) {
      window.clearTimeout(deferredRenderTimer);
      deferredRenderTimer = null;
    }
    if (deferredReflowTimer !== null) {
      window.clearTimeout(deferredReflowTimer);
      deferredReflowTimer = null;
    }
    refreshNativeNumberingLabels(doc);
    const rendered = renderNativeDocument(doc);
    host.replaceChildren(rendered);
    paginateNativeDocument(rendered, doc);
    observeNativeTableSizes(rendered);
    const tableSelection = preservedSelection?.kind === 'table-cell-range'
      ? preservedSelection
      : selection?.kind === 'table-cell-range'
        ? selection
        : null;
    renderTableCellRangeSelection(tableSelection);
    renderSearchHighlights();
    renderPreservedSelectionHighlight();
    if (selectedImage) showImageResizeHandle();
    if (preservedSelection?.kind === 'text-range') {
      restoreNativeTextSelection(host, preservedSelection.anchor, preservedSelection.focus);
    } else if (cursor) {
      restoreNativeCursor(host, cursor);
    }
  };

  const renderParagraph = (paraId: string, cursor?: NativeCursorPoint) => {
    const paragraph = findNativeParagraph(doc, paraId);
    if (!paragraph) {
      render(cursor);
      return;
    }
    const replacements = Array.from(host.querySelectorAll<HTMLElement>(`coco-word-native-p[data-para-id="${paraId}"]`));
    if (replacements.length === 0) {
      render(cursor);
      return;
    }
    for (const current of replacements) {
      current.replaceWith(renderNativeParagraph(paragraph));
    }
    const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document') ?? host.querySelector<HTMLElement>('coco-word-native-document');
    if (root) refreshNativePageCommentBubbles(root);
    if (selectedImage) showImageResizeHandle();
    if (cursor) restoreNativeCursor(host, cursor);
  };

  const scheduleRender = (cursor?: NativeCursorPoint, preservedSelection?: WordEditorSelection) => {
    if (deferredRenderTimer !== null) window.clearTimeout(deferredRenderTimer);
    deferredRenderTimer = window.setTimeout(() => {
      deferredRenderTimer = null;
      render(cursor, preservedSelection);
    }, 450);
  };

  const reflowFromParagraph = (
    paraId: string,
    cursor?: NativeCursorPoint,
    preservedSelection?: WordEditorSelection,
  ): boolean => {
    const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
    if (!root) return false;
    const changed = reflowNativeDocumentFromParagraph(root, doc, paraId);
    observeNativeTableSizes(root);
    for (const image of Array.from(
      root.querySelectorAll<HTMLImageElement>(`coco-word-native-p[data-para-id="${paraId}"] img`),
    )) {
      if (image.complete || pendingImageReflows.has(image)) continue;
      pendingImageReflows.add(image);
      image.addEventListener(
        'load',
        () => {
          pendingImageReflows.delete(image);
          reflowFromParagraph(paraId);
        },
        { once: true },
      );
    }
    if (preservedSelection?.kind === 'text-range') {
      restoreNativeTextSelection(host, preservedSelection.anchor, preservedSelection.focus);
    } else if (cursor) {
      restoreNativeCursor(host, cursor);
    }
    return changed;
  };

  const scheduleReflow = (
    paraId: string,
    delay = 180,
    shouldRun?: (root: HTMLElement) => boolean,
  ) => {
    if (compositionActive || pendingCompositionTransaction) return;
    if (deferredReflowTimer !== null) window.clearTimeout(deferredReflowTimer);
    deferredReflowTimer = window.setTimeout(() => {
      deferredReflowTimer = null;
      if (compositionActive || pendingCompositionTransaction) return;
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document') ?? host.querySelector<HTMLElement>('coco-word-native-document');
      if (!root || (shouldRun && !shouldRun(root))) return;
      const current = selection?.kind === 'object' ? selection : currentNativeSelection(host);
      if (current?.kind === 'text-range') {
        reflowFromParagraph(paraId, undefined, current);
      } else if (current?.kind === 'object') {
        reflowFromParagraph(paraId);
        selectedImage = { paraId: current.paraId, drawingId: current.objectId };
        showImageResizeHandle();
      } else {
        const cursor = current?.kind === 'caret' ? current : currentNativeCursor(host);
        reflowFromParagraph(paraId, cursor ?? undefined);
      }
    }, delay);
  };

  const scheduleImageLayoutReflow = (paraId: string, delay = 80) => {
    const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document') ?? host.querySelector<HTMLElement>('coco-word-native-document');
    if (!root) return;
    let waitsForImage = false;
    for (const image of Array.from(
      root.querySelectorAll<HTMLImageElement>(`coco-word-native-p[data-para-id="${paraId}"] img`),
    )) {
      if (image.complete || pendingImageReflows.has(image)) continue;
      waitsForImage = true;
      pendingImageReflows.add(image);
      image.addEventListener(
        'load',
        () => {
          pendingImageReflows.delete(image);
          scheduleReflow(paraId, 0);
        },
        { once: true },
      );
      image.addEventListener('error', () => pendingImageReflows.delete(image), { once: true });
    }
    scheduleReflow(paraId, waitsForImage ? 800 : delay);
  };

  const captureTypingLayoutBaseline = (paraId: string) => {
    if (typingLayoutBaseline?.paraId === paraId) return;
    const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document') ?? host.querySelector<HTMLElement>('coco-word-native-document');
    if (!root) return;
    typingLayoutBaseline = {
      paraId,
      measurement: measureNativeDocumentReflowFromParagraph(root, doc, paraId),
    };
  };

  const scheduleTypingReflow = (paraId: string) => {
    scheduleReflow(paraId, 700, (root) => {
      const baseline = typingLayoutBaseline;
      typingLayoutBaseline = null;
      if (!baseline || baseline.paraId !== paraId) return false;
      const current = measureNativeDocumentReflowFromParagraph(root, doc, paraId);
      return (
        current.overflowPx > Math.max(1, baseline.measurement.overflowPx + 1) ||
        (!baseline.measurement.canBackfillPrevious && current.canBackfillPrevious) ||
        (!baseline.measurement.canBackfillNext && current.canBackfillNext)
      );
    });
  };

  const observedTableSizes = new WeakMap<HTMLElement, number>();
  const tableResizeObserver = new ResizeObserver((entries) => {
    for (const entry of entries) {
      const table = entry.target;
      if (!(table instanceof HTMLElement)) continue;
      const height = Math.round(entry.contentRect.height);
      const previous = observedTableSizes.get(table);
      observedTableSizes.set(table, height);
      if (previous === undefined || Math.abs(previous - height) <= 1) continue;
      const paraId = table.querySelector<HTMLElement>('coco-word-native-p[data-para-id]')?.dataset.paraId;
      if (paraId) {
        const activeElement = document.activeElement;
        if (activeElement && table.contains(activeElement)) {
          const activeParaId = activeElement.closest<HTMLElement>('coco-word-native-p[data-para-id]')?.dataset.paraId;
          scheduleTypingReflow(activeParaId ?? paraId);
        } else {
          scheduleReflow(paraId);
        }
      }
    }
  });

  const observeNativeTableSizes = (root: HTMLElement) => {
    for (const table of Array.from(root.querySelectorAll<HTMLElement>('coco-word-native-table'))) {
      if (observedTableSizes.has(table)) continue;
      observedTableSizes.set(table, Math.round(table.getBoundingClientRect().height));
      tableResizeObserver.observe(table);
    }
  };

  const notifyChange = (mode: 'immediate' | 'deferred' = 'immediate') => {
    if (destroyed) return;
    const dispatchChange = () => {
      refreshActiveSearch();
      const event = createWordEditorChangeEvent(emittedChanges, changes, ++revision);
      emittedChanges = changes.slice();
      options.onChange?.(event);
      emit('change', event);
    };
    if (mode === 'immediate') {
      if (deferredChangeTimer !== null) {
        window.clearTimeout(deferredChangeTimer);
        deferredChangeTimer = null;
      }
      dispatchChange();
      return;
    }
    if (deferredChangeTimer !== null) window.clearTimeout(deferredChangeTimer);
    deferredChangeTimer = window.setTimeout(() => {
      deferredChangeTimer = null;
      dispatchChange();
    }, 180);
  };

  const notifySelectionChange = () => {
    if (destroyed) return;
    if (deferredSelectionTimer !== null) window.clearTimeout(deferredSelectionTimer);
    deferredSelectionTimer = window.setTimeout(() => {
      deferredSelectionTimer = null;
      const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
      const event: WordEditorSelectionChangeEvent = {
        selection: current,
        commandState: getCommandState(current),
      };
      options.onSelectionChange?.(event);
      emit('selectionChange', event);
    }, 80);
  };

  const suppressBrowserHistoryInput = (inputType: 'historyUndo' | 'historyRedo') => {
    suppressedBrowserHistoryInput = inputType;
    if (browserHistorySuppressionTimer !== null) window.clearTimeout(browserHistorySuppressionTimer);
    browserHistorySuppressionTimer = window.setTimeout(() => {
      suppressedBrowserHistoryInput = null;
      browserHistorySuppressionTimer = null;
    }, 1_000);
  };

  const removeNativeImageDom = (paraId: string, drawingId: string, cursor?: NativeCursorPoint) => {
    const images = Array.from(
      host.querySelectorAll<HTMLImageElement>(
        `coco-word-native-p[data-para-id="${paraId}"] img[data-drawing-id="${drawingId}"]`,
      ),
    );
    for (const image of images) image.remove();
    clearNativeImageSelection(host);
    if (cursor) restoreNativeCursor(host, cursor);
  };

  const record = (
    change: NativeWordChange | NativeWordChange[] | null,
    cursor?: NativeCursorPoint,
    renderOptions: { localParaId?: string } = {},
  ) => {
    if (Array.isArray(change)) changes.push(...change);
    else if (change) changes.push(change);
    if (change) browserRedoStack.length = 0;
    if (renderOptions.localParaId) {
      renderParagraph(renderOptions.localParaId, cursor);
      scheduleReflow(renderOptions.localParaId);
    } else {
      render(cursor);
    }
    notifyChange();
  };

  const canBrowserEditParagraphText = (paraId: string): boolean => {
    const paragraph = findNativeParagraph(doc, paraId);
    return !!paragraph && paragraph.content.every((child) => child.kind === 'text');
  };

  const paragraphTextLength = (paraId: string): number => {
    const paragraph = findNativeParagraph(doc, paraId);
    return paragraph ? nativeParagraphText(paragraph).length : 0;
  };

  const adjacentRevisionId = (
    paraId: string,
    offset: number,
    revisionKind: 'insert' | 'delete',
    side: 'before' | 'at' | 'after',
  ): string | null => {
    const paragraph = findNativeParagraph(doc, paraId);
    if (!paragraph) return null;
    const target = side === 'before' ? offset - 1 : side === 'after' ? offset + 1 : offset;
    if (target < 0) return null;
    let position = 0;
    for (const child of paragraph.content) {
      const length = child.kind === 'text' ? child.text.length : 1;
      if (target >= position && target < position + length) {
        const mark = child.kind === 'text'
          ? child.marks?.find((item) => item.kind === 'revision' && item.revisionKind === revisionKind)
          : undefined;
        return mark && 'revisionId' in mark ? mark.revisionId ?? null : null;
      }
      position += length;
    }
    return null;
  };

  const markBrowserTextChangeAsRevision = (
    paraId: string,
    change: NativeWordChange | null,
  ): NativeWordChange | null => {
    if (!trackRevisions || !change || change.kind !== 'text') return null;
    if ((change.op !== 'insert' && change.op !== 'replace') || change.text.length === 0) return null;
    return nativeAddRevision(doc, paraId, change.offset, change.offset + change.text.length, {
      revisionId: adjacentRevisionId(paraId, change.offset, 'insert', 'before') ?? allocateNativeRevisionId(doc),
      revisionKind: 'insert',
      author: 'User',
      date: new Date().toISOString(),
    });
  };

  const editableRangeParagraphs = (
    range: PendingBrowserRangeDelete['range'],
  ): NativeWordParagraphNode[] | null => {
    const anchorLocation = findNativeParagraphLocation(doc, range.anchor.paraId);
    const focusLocation = findNativeParagraphLocation(doc, range.focus.paraId);
    if (!anchorLocation || !focusLocation || anchorLocation.container !== focusLocation.container) return null;
    const from = Math.min(anchorLocation.index, focusLocation.index);
    const to = Math.max(anchorLocation.index, focusLocation.index);
    const paragraphs: NativeWordParagraphNode[] = [];
    for (let index = from; index <= to; index += 1) {
      const block = anchorLocation.container[index];
      if (block?.kind !== 'paragraph' || block.content.some((child) => child.kind !== 'text')) return null;
      paragraphs.push(cloneNativeParagraph(block));
    }
    return paragraphs;
  };

  const canBrowserDeleteRangeAcrossParagraphs = (
    range: PendingBrowserRangeDelete['range'],
  ): NativeWordParagraphNode[] | null => {
    const paragraphs = editableRangeParagraphs(range);
    if (!paragraphs) return null;
    return paragraphs.length > 1 ? paragraphs : null;
  };

  const captureBrowserRangeDelete = (range: PendingBrowserRangeDelete['range']): boolean => {
    const beforeParagraphs = canBrowserDeleteRangeAcrossParagraphs(range);
    if (!beforeParagraphs) return false;
    pendingBrowserRangeDelete = {
      range,
      beforeParagraphs,
      changesStart: changes.length,
    };
    return true;
  };

  const beginTableBrowserTypingTransaction = (
    range: PendingBrowserRangeDelete['range'],
  ): void => {
    const paraId = range.anchor.paraId;
    if (range.focus.paraId !== paraId) return;
    const paragraphEl = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${paraId}"]`);
    const paragraph = findNativeParagraph(doc, paraId);
    if (!paragraphEl?.closest('coco-word-native-table') || !paragraph) return;
    if (activeLocalTypingTransaction?.afterParagraphs[0]?.paraId === paraId) return;
    const beforeParagraph = cloneNativeParagraph(paragraph);
    const transaction: BrowserTextTransaction = {
      source: 'browser',
      patchMode: 'individual',
      skipDeferredRender: true,
      range,
      beforeParagraphs: [beforeParagraph],
      afterParagraphs: [beforeParagraph],
      changesStart: changes.length,
      changes: [],
    };
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = transaction;
  };

  const compositionStart = (event: CompositionEvent) => {
    compositionActive = true;
    if (compositionFinalizeTimer !== null) {
      window.clearTimeout(compositionFinalizeTimer);
      compositionFinalizeTimer = null;
    }
    if (pendingCompositionTransaction) finalizeComposition();
    if (deferredRenderTimer !== null) {
      window.clearTimeout(deferredRenderTimer);
      deferredRenderTimer = null;
    }
    if (deferredReflowTimer !== null) {
      window.clearTimeout(deferredReflowTimer);
      deferredReflowTimer = null;
    }
    if (deferredChangeTimer !== null) {
      window.clearTimeout(deferredChangeTimer);
      deferredChangeTimer = null;
    }
    if (deferredSelectionTimer !== null) {
      window.clearTimeout(deferredSelectionTimer);
      deferredSelectionTimer = null;
    }
    const currentSelection = currentNativeSelection(host);
    const targetParagraph = event.target instanceof Node ? parentNativeParagraph(event.target) : null;
    const targetParaId = targetParagraph?.dataset.paraId;
    const currentPoint = currentNativeCursor(host);
    const point = targetParaId
      ? {
          paraId: targetParaId,
          offset: currentPoint?.paraId === targetParaId
            ? currentPoint.offset
            : paragraphTextLength(targetParaId),
        }
      : currentPoint;
    const paragraph = point ? findNativeParagraph(doc, point.paraId) : null;
    if (!point || !paragraph || paragraph.content.some((child) => child.kind !== 'text')) return;
    captureTypingLayoutBaseline(point.paraId);
    activeLocalTypingTransaction = null;
    const range = currentSelection?.kind === 'text-range' &&
      currentSelection.anchor.paraId === point.paraId &&
      currentSelection.focus.paraId === point.paraId
      ? {
          anchor: { ...currentSelection.anchor },
          focus: { ...currentSelection.focus },
        }
      : { anchor: point, focus: point };
    pendingCompositionTransaction = {
      paraId: point.paraId,
      beforeParagraph: cloneNativeParagraph(paragraph),
      changesStart: changes.length,
      range,
    };
  };

  const finalizeComposition = () => {
    compositionFinalizeTimer = null;
    const pending = pendingCompositionTransaction;
    pendingCompositionTransaction = null;
    if (!pending) return;
    const paragraph = findNativeParagraph(doc, pending.paraId);
    if (!paragraph) return;
    const paragraphEl = host.querySelector<HTMLElement>(
      `coco-word-native-p[data-para-id="${pending.paraId}"]`,
    );
    if (!paragraphEl) return;
    const change = nativeSyncParagraphText(doc, pending.paraId, paragraphEl.textContent ?? '');
    if (change) changes.push(change);
    const revision = markBrowserTextChangeAsRevision(pending.paraId, change);
    if (revision) changes.push(revision);
    const afterParagraph = cloneNativeParagraph(paragraph);
    if (JSON.stringify(afterParagraph) === JSON.stringify(pending.beforeParagraph)) {
      changes.splice(pending.changesStart);
      return;
    }
    const transaction: BrowserTextTransaction = {
      source: 'browser',
      patchMode: 'individual',
      skipDeferredRender: true,
      range: pending.range,
      beforeParagraphs: [pending.beforeParagraph],
      afterParagraphs: [afterParagraph],
      changesStart: pending.changesStart,
      changes: changes.slice(pending.changesStart),
    };
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    const cursor = currentNativeCursor(host);
    if (cursor) setSelection({ kind: 'caret', ...cursor });
    scheduleTypingReflow(pending.paraId);
    if (revision) renderParagraph(pending.paraId, cursor ?? undefined);
    notifyChange('deferred');
  };

  const compositionEnd = () => {
    compositionActive = false;
    if (compositionFinalizeTimer !== null) window.clearTimeout(compositionFinalizeTimer);
    compositionFinalizeTimer = window.setTimeout(finalizeComposition, 80);
  };

  const replaceTransactionParagraphs = (
    transaction: BrowserTextTransaction,
    paragraphs: NativeWordParagraphNode[],
  ): boolean => {
    const affectedIds = new Set(
      [...transaction.beforeParagraphs, ...transaction.afterParagraphs].map((paragraph) => paragraph.paraId),
    );
    const firstExistingId = [...transaction.beforeParagraphs, ...transaction.afterParagraphs].find((paragraph) =>
      findNativeParagraph(doc, paragraph.paraId),
    )?.paraId;
    if (!firstExistingId) return false;
    const location = findNativeParagraphLocation(doc, firstExistingId);
    if (!location) return false;
    if (transaction.patchMode === 'individual') {
      for (const paragraph of paragraphs) {
        const paragraphLocation = findNativeParagraphLocation(doc, paragraph.paraId);
        if (!paragraphLocation) return false;
        Object.assign(paragraphLocation.paragraph, cloneNativeParagraph(paragraph));
      }
      return true;
    }
    const existingIds = Array.from(affectedIds).filter((paraId) => findNativeParagraph(doc, paraId));
    if (existingIds.length === 0 || paragraphs.length === 0) return false;
    const firstTarget = cloneNativeParagraph(paragraphs[0]!);
    Object.assign(location.paragraph, firstTarget);
    for (const paraId of existingIds) {
      if (paraId === firstExistingId) continue;
      findNativeParagraphLocation(doc, paraId)?.remove();
    }
    let previousId = firstTarget.paraId;
    for (const paragraph of paragraphs.slice(1)) {
      const previous = findNativeParagraphLocation(doc, previousId);
      if (!previous) return false;
      const next = cloneNativeParagraph(paragraph);
      previous.insertAfter(next);
      previousId = next.paraId;
    }
    return true;
  };

  const syncBrowserRangeDelete = (): boolean => {
    if (!pendingBrowserRangeDelete) return false;
    const pending = pendingBrowserRangeDelete;
    pendingBrowserRangeDelete = null;
    const domTexts = pending.beforeParagraphs.map((paragraph) => domParagraphText(host, paragraph.paraId));
    const allParagraphsRemainInDom = domTexts.every((text) => text !== null);
    const transactionChanges: NativeWordChange[] = [];

    if (allParagraphsRemainInDom) {
      for (let index = 0; index < pending.beforeParagraphs.length; index += 1) {
        const paragraph = pending.beforeParagraphs[index]!;
        const change = nativeSyncParagraphText(doc, paragraph.paraId, domTexts[index] ?? '');
        if (change) transactionChanges.push(change);
      }
    } else {
      transactionChanges.push(...nativeDeleteTextRangeAcrossParagraphs(doc, pending.range.anchor, pending.range.focus));
    }

    if (transactionChanges.length === 0) return false;
    changes.push(...transactionChanges);
    const afterParagraphs = pending.beforeParagraphs
      .map((paragraph) => findNativeParagraph(doc, paragraph.paraId))
      .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null);
    if (afterParagraphs.length > 0) {
      const transaction: BrowserTextTransaction = {
        ...pending,
        source: 'browser',
        afterParagraphs: afterParagraphs.map(cloneNativeParagraph),
        changes: transactionChanges,
      };
      browserUndoStack.push(transaction);
      activeLocalTypingTransaction = null;
      browserRedoStack.length = 0;
    }
    const cursor = currentNativeCursor(host);
    if (cursor) setSelection({ kind: 'caret', paraId: cursor.paraId, offset: cursor.offset });
    reflowFromParagraph(pending.beforeParagraphs[0]!.paraId, cursor ?? undefined);
    notifyChange('deferred');
    return true;
  };

  const syncBrowserHistoryUndo = (): boolean => {
    const transaction = browserUndoStack[browserUndoStack.length - 1];
    if (!transaction || isBrowserBlockTransaction(transaction) || !domParagraphsMatch(host, transaction.beforeParagraphs)) {
      return false;
    }
    if (!replaceTransactionParagraphs(transaction, transaction.beforeParagraphs)) return false;
    browserUndoStack.pop();
    browserRedoStack.push(transaction);
    changes.splice(transaction.changesStart);
    const cursor = currentNativeCursor(host);
    if (cursor) setSelection({ kind: 'caret', paraId: cursor.paraId, offset: cursor.offset });
    reflowFromParagraph(transaction.beforeParagraphs[0]!.paraId, cursor ?? undefined);
    notifyChange('deferred');
    return true;
  };

  const syncBrowserHistoryRedo = (): boolean => {
    const transaction = browserRedoStack[browserRedoStack.length - 1];
    if (!transaction || isBrowserBlockTransaction(transaction) || !domParagraphsMatch(host, transaction.afterParagraphs)) {
      return false;
    }
    if (!replaceTransactionParagraphs(transaction, transaction.afterParagraphs)) return false;
    browserRedoStack.pop();
    browserUndoStack.push(transaction);
    changes.splice(transaction.changesStart);
    changes.push(...transaction.changes);
    const cursor = currentNativeCursor(host);
    if (cursor) setSelection({ kind: 'caret', paraId: cursor.paraId, offset: cursor.offset });
    reflowFromParagraph(transaction.beforeParagraphs[0]!.paraId, cursor ?? undefined);
    notifyChange('deferred');
    return true;
  };

  const patchTransactionDom = (
    transaction: BrowserTextTransaction,
    paragraphs: NativeWordParagraphNode[],
    cursor?: NativeCursorPoint,
  ): boolean => {
    if (transaction.patchMode === 'individual') {
      let patched = false;
      for (const paragraph of paragraphs) {
        const replacements = Array.from(
          host.querySelectorAll<HTMLElement>(`coco-word-native-p[data-para-id="${paragraph.paraId}"]`),
        );
        for (const current of replacements) {
          current.replaceWith(renderNativeParagraph(paragraph));
          patched = true;
        }
      }
      if (cursor) restoreNativeCursor(host, cursor);
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) refreshNativePageCommentBubbles(root);
      return patched;
    }
    const affectedIds = new Set(
      [...transaction.beforeParagraphs, ...transaction.afterParagraphs].map((paragraph) => paragraph.paraId),
    );
    const affectedEls = Array.from(affectedIds).flatMap((paraId) =>
      Array.from(host.querySelectorAll<HTMLElement>(`coco-word-native-p[data-para-id="${paraId}"]`)),
    );
    const firstParaId = transaction.beforeParagraphs[0]?.paraId;
    const first = affectedEls.find((paragraph) => paragraph.dataset.paraId === firstParaId) ?? affectedEls[0];
    if (!first) return false;
    const rendered = paragraphs.map((paragraph) => renderNativeParagraph(paragraph));
    first.replaceWith(...rendered);
    for (const paragraph of affectedEls) {
      if (paragraph !== first && paragraph.isConnected) paragraph.remove();
    }
    if (cursor) restoreNativeCursor(host, cursor);
    const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document') ?? host.querySelector<HTMLElement>('coco-word-native-document');
    if (root) refreshNativePageCommentBubbles(root);
    return true;
  };

  const replaceBlockTransaction = (
    transaction: BrowserBlockTransaction,
    blocks: NativeWordBlockNode[],
    currentCount: number,
  ): boolean => {
    transaction.container.splice(transaction.startIndex, currentCount, ...cloneNativeWordBlocks(blocks));
    return true;
  };

  const patchBlockTransactionDom = (
    transaction: BrowserBlockTransaction,
    blocks: NativeWordBlockNode[],
    cursor?: NativeCursorPoint,
    preservedSelection?: Extract<WordEditorSelection, { kind: 'text-range' }>,
  ): boolean => {
    const affected = [...transaction.beforeBlocks, ...transaction.afterBlocks].flatMap((block) =>
      nativeBlockElements(host, block),
    );
    const firstParaId = transaction.afterBlocks[0]?.kind === 'paragraph'
      ? transaction.afterBlocks[0].paraId
      : transaction.beforeBlocks[0]?.kind === 'paragraph'
        ? transaction.beforeBlocks[0].paraId
        : null;
    const first = firstParaId
      ? host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${firstParaId}"]`)
      : affected[0];
    if (!first) return false;
    const rendered = blocks.map((block) => renderNativeBlock(block));
    first.replaceWith(...rendered);
    for (const element of affected) {
      if (element !== first && element.isConnected) element.remove();
    }
    const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
    if (root) observeNativeTableSizes(root);
    if (root) refreshNativePageCommentBubbles(root);
    if (preservedSelection) restoreNativeTextSelection(host, preservedSelection.anchor, preservedSelection.focus);
    else if (cursor) restoreNativeCursor(host, cursor);
    return true;
  };

  const selectionSpansPages = (range: PendingBrowserRangeDelete['range']): boolean => {
    const anchor = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${range.anchor.paraId}"]`);
    const focus = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${range.focus.paraId}"]`);
    return !!anchor && !!focus && anchor.closest('coco-word-native-page') !== focus.closest('coco-word-native-page');
  };

  const applyLocalRangeReplacement = (range: PendingBrowserRangeDelete['range'], text: string): boolean => {
    const beforeParagraphs = canBrowserDeleteRangeAcrossParagraphs(range);
    if (!beforeParagraphs) return false;
    const pending: PendingBrowserRangeDelete = {
      range,
      beforeParagraphs,
      changesStart: changes.length,
    };
    const transactionChanges = nativeReplaceTextRangeAcrossParagraphs(doc, range.anchor, range.focus, text);
    if (transactionChanges.length === 0) return false;
    const afterParagraphs = beforeParagraphs
      .map((paragraph) => findNativeParagraph(doc, paragraph.paraId))
      .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
      .map(cloneNativeParagraph);
    if (afterParagraphs.length === 0) return false;
    const transaction: BrowserTextTransaction = {
      ...pending,
      source: 'local',
      afterParagraphs,
      changes: transactionChanges,
    };
    const cursor = {
      paraId: range.anchor.paraId,
      offset: range.anchor.offset + text.length,
    };
    changes.push(...transactionChanges);
    browserUndoStack.push(transaction);
    activeLocalTypingTransaction = text.length > 0 ? transaction : null;
    browserRedoStack.length = 0;
    if (!patchTransactionDom(transaction, afterParagraphs, cursor)) render(cursor);
    else reflowFromParagraph(beforeParagraphs[0]!.paraId, cursor);
    setSelection({ kind: 'caret', ...cursor });
    notifyChange('deferred');
    return true;
  };

  const applyLocalBlockRangeReplacement = (
    range: PendingBrowserRangeDelete['range'],
    text: string,
  ): boolean => {
    const result = nativeReplaceBlockRangeAcrossParagraphs(doc, range.anchor, range.focus, text);
    if (!result) return false;
    const transaction: BrowserBlockTransaction = {
      source: 'local',
      range,
      container: result.container,
      startIndex: result.startIndex,
      beforeBlocks: result.beforeBlocks,
      afterBlocks: result.afterBlocks,
      changesStart: changes.length,
      changes: result.changes,
    };
    changes.push(...result.changes);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    if (!patchBlockTransactionDom(transaction, result.afterBlocks, result.cursor)) render(result.cursor);
    else {
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, result.cursor.paraId);
      reflowFromParagraph(result.cursor.paraId, result.cursor);
    }
    setSelection({ kind: 'caret', ...result.cursor });
    notifyChange('deferred');
    return true;
  };

  const applyLocalClipboardBlocks = (
    range: PendingBrowserRangeDelete['range'],
    blocks: readonly NativeWordBlockNode[],
  ): boolean => {
    const result = nativeReplaceRangeWithBlocks(doc, range.anchor, range.focus, blocks);
    if (!result) return false;
    const transaction: BrowserBlockTransaction = {
      source: 'local',
      range,
      container: result.container,
      startIndex: result.startIndex,
      beforeBlocks: result.beforeBlocks,
      afterBlocks: result.afterBlocks,
      changesStart: changes.length,
      changes: result.changes,
    };
    changes.push(...result.changes);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    for (const paragraph of allNativeParagraphs(doc)) paragraphIds.add(paragraph.paraId);
    if (!patchBlockTransactionDom(transaction, result.afterBlocks, result.cursor)) render(result.cursor);
    else {
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, result.cursor.paraId);
      reflowFromParagraph(result.cursor.paraId, result.cursor);
    }
    setSelection({ kind: 'caret', ...result.cursor });
    notifyChange('deferred');
    return true;
  };

  const applyLocalTableStructureEdit = (
    current: WordEditorSelection,
    command: Extract<
      WordEditorCommand,
      { kind: 'insert-table-row' | 'delete-table-row' | 'insert-table-column' | 'delete-table-column' }
    >,
  ): boolean => {
    const point = current.kind === 'table-cell-range'
      ? (() => {
          const table = findNativeTableLocation(doc, current.tableId)?.table;
          const cell = table ? nativeTableCellsInRange(table, current)[0]?.cell : null;
          const paraId = cell?.paragraphs[0]?.paraId;
          return paraId ? { paraId, offset: 0 } : null;
        })()
      : current.kind === 'text-range'
        ? current.anchor
        : { paraId: current.paraId, offset: current.offset };
    if (!point) return false;
    const location = findNativeTableCellLocation(doc, point.paraId);
    if (!location) return false;
    const beforeTable = cloneNativeWordBlocks([location.table])[0];
    if (!beforeTable || beforeTable.kind !== 'table') return false;
    const operation = command.kind === 'insert-table-row'
      ? { kind: 'insert-row' as const, position: command.position }
      : command.kind === 'delete-table-row'
        ? { kind: 'delete-row' as const }
        : command.kind === 'insert-table-column'
          ? { kind: 'insert-column' as const, position: command.position }
          : { kind: 'delete-column' as const };
    const result = nativeEditTableStructure(doc, point.paraId, operation, createRuntimeParaId);
    if (!result) return false;
    const afterTable = cloneNativeWordBlocks([location.table])[0];
    if (!afterTable || afterTable.kind !== 'table') return false;
    const revisionId = trackRevisions ? allocateNativeRevisionId(doc) : null;
    const revisionKind = command.kind === 'insert-table-row' || command.kind === 'insert-table-column' ? 'insert' : 'delete';
    if (revisionId) {
      location.table.revision = {
        kind: 'revision',
        id: `revision:${revisionId}`,
        revisionId,
        revisionKind,
        author: 'User',
        date: new Date().toISOString(),
      };
      afterTable.revision = { ...location.table.revision };
      if (result.change.kind === 'table') {
        result.change.revisionId = revisionId;
        result.change.revisionKind = revisionKind;
        result.change.beforeTable = beforeTable;
      }
    }
    const afterSelection: WordEditorSelection = { kind: 'caret', ...result.cursor };
    const transaction: BrowserBlockTransaction = {
      source: 'local',
      range: { anchor: point, focus: point },
      container: location.container,
      startIndex: location.tableIndex,
      beforeBlocks: [beforeTable],
      afterBlocks: [afterTable],
      changesStart: changes.length,
      changes: [result.change],
      beforeSelection: current,
      afterSelection,
    };
    changes.push(result.change);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    if (!patchBlockTransactionDom(transaction, transaction.afterBlocks, result.cursor)) render(result.cursor);
    else {
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, result.cursor.paraId);
      reflowFromParagraph(result.cursor.paraId, result.cursor);
    }
    setSelection(afterSelection);
    notifyChange('deferred');
    return true;
  };

  const applyLocalTableCellStructureEdit = (
    current: Extract<WordEditorSelection, { kind: 'table-cell-range' }>,
    command: Extract<WordEditorCommand, { kind: 'merge-table-cells' | 'split-table-cell' }>,
  ): boolean => {
    const location = findNativeTableLocation(doc, current.tableId);
    if (!location) return false;
    const beforeTable = cloneNativeWordBlocks([location.table])[0];
    if (!beforeTable || beforeTable.kind !== 'table') return false;
    const selected = nativeTableCellsInRange(location.table, current);
    const first = selected[0];
    const result = command.kind === 'merge-table-cells'
      ? nativeMergeTableCells(doc, current.tableId, current,)
      : first
        ? nativeSplitTableCell(doc, current.tableId, first.rowIndex, first.columnIndex, createRuntimeParaId)
        : null;
    if (!result) return false;
    const afterTable = cloneNativeWordBlocks([location.table])[0];
    if (!afterTable || afterTable.kind !== 'table') return false;
    const revisionId = trackRevisions ? allocateNativeRevisionId(doc) : null;
    const revisionKind = command.kind === 'merge-table-cells' ? 'insert' : 'delete';
    if (revisionId) {
      location.table.revision = {
        kind: 'revision',
        id: `revision:${revisionId}`,
        revisionId,
        revisionKind,
        author: 'User',
        date: new Date().toISOString(),
      };
      afterTable.revision = { ...location.table.revision };
      if (result.change.kind === 'table') {
        result.change.revisionId = revisionId;
        result.change.revisionKind = revisionKind;
        result.change.beforeTable = beforeTable;
      }
    }
    const transaction: BrowserBlockTransaction = {
      source: 'local',
      range: {
        anchor: { paraId: result.cursor.paraId, offset: result.cursor.offset },
        focus: { paraId: result.cursor.paraId, offset: result.cursor.offset },
      },
      container: location.container,
      startIndex: location.tableIndex,
      beforeBlocks: [beforeTable],
      afterBlocks: [afterTable],
      changesStart: changes.length,
      changes: [result.change],
      beforeSelection: current,
      afterSelection: { kind: 'caret', ...result.cursor },
    };
    changes.push(result.change);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    if (!patchBlockTransactionDom(transaction, transaction.afterBlocks, result.cursor)) render(result.cursor);
    else {
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, result.cursor.paraId);
      reflowFromParagraph(result.cursor.paraId, result.cursor);
    }
    setSelection({ kind: 'caret', ...result.cursor });
    notifyChange('deferred');
    return true;
  };

  const applyLocalTableLifecycleEdit = (
    current: WordEditorSelection,
    command: Extract<WordEditorCommand, { kind: 'insert-table' | 'delete-table' }>,
  ): boolean => {
    const transactionPoint = current.kind === 'text-range'
      ? current.anchor
      : current.kind === 'table-cell-range'
        ? (() => {
            const table = findNativeTableLocation(doc, current.tableId)?.table;
            const paragraph = table ? nativeTableCellsInRange(table, current)[0]?.cell.paragraphs[0] : null;
            return paragraph ? { paraId: paragraph.paraId, offset: 0 } : null;
          })()
        : { paraId: current.paraId, offset: current.offset };
    if (!transactionPoint) return false;
    const result = command.kind === 'insert-table'
      ? current.kind === 'caret'
        ? nativeInsertTable(
            doc,
            current.paraId,
            current.offset,
            command.rows,
            command.columns,
            createRuntimeParaId,
          )
        : null
      : (() => {
          const tableId = getCommandState(current).table?.tableId;
          return tableId ? nativeDeleteTable(doc, tableId, createRuntimeParaId) : null;
        })();
    if (!result) return false;
    const revisionId = trackRevisions ? allocateNativeRevisionId(doc) : null;
    const revisionKind = command.kind === 'insert-table' ? 'insert' : 'delete';
    if (revisionId) {
      const table = result.afterBlocks.find((block) => block.kind === 'table');
      if (table?.kind === 'table') {
        table.revision = {
          kind: 'revision',
          id: `revision:${revisionId}`,
          revisionId,
          revisionKind,
          author: 'User',
          date: new Date().toISOString(),
        };
        result.changes.forEach((change) => {
          if (change.kind === 'table') {
            change.revisionId = revisionId;
            change.revisionKind = revisionKind;
            const beforeTable = result.beforeBlocks.find((block) => block.kind === 'table');
            if (beforeTable?.kind === 'table') change.beforeTable = beforeTable;
            if (change.op === 'insert-table') {
              change.anchorParaId = transactionPoint.paraId;
              change.anchorOffset = transactionPoint.offset;
              const insertedTable = result.afterBlocks.find((block) => block.kind === 'table');
              if (insertedTable?.kind === 'table') change.tableSnapshot = insertedTable;
            }
          }
        });
      }
    }
    for (const change of result.changes) {
      if (change.kind !== 'table' || change.op !== 'insert-table') continue;
      change.anchorParaId = transactionPoint.paraId;
      change.anchorOffset = transactionPoint.offset;
      const insertedTable = result.afterBlocks.find((block) => block.kind === 'table');
      if (insertedTable?.kind === 'table') change.tableSnapshot = insertedTable;
    }
    const afterSelection: WordEditorSelection = { kind: 'caret', ...result.cursor };
    const transaction: BrowserBlockTransaction = {
      source: 'local',
      range: {
        anchor: transactionPoint,
        focus: current.kind === 'text-range' ? current.focus : transactionPoint,
      },
      container: result.container,
      startIndex: result.startIndex,
      beforeBlocks: result.beforeBlocks,
      afterBlocks: result.afterBlocks,
      changesStart: changes.length,
      changes: result.changes,
      beforeSelection: current,
      afterSelection,
    };
    changes.push(...result.changes);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    for (const paragraph of allNativeParagraphs(doc)) paragraphIds.add(paragraph.paraId);
    if (!patchBlockTransactionDom(transaction, transaction.afterBlocks, result.cursor)) render(result.cursor);
    else {
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, result.cursor.paraId);
      reflowFromParagraph(result.cursor.paraId, result.cursor);
    }
    setSelection(afterSelection);
    notifyChange('deferred');
    return true;
  };

  const applyLocalTableCellFormat = (
    current: WordEditorSelection,
    command: Extract<
      WordEditorCommand,
      { kind: 'set-table-cell-background' | 'set-table-cell-vertical-alignment' | 'set-table-cell-borders' }
    >,
  ): boolean => {
    const rangeLocation = current.kind === 'table-cell-range' ? findNativeTableLocation(doc, current.tableId) : null;
    const rangeCell = current.kind === 'table-cell-range' && rangeLocation
      ? nativeTableCellsInRange(rangeLocation.table, current)[0]?.cell
      : null;
    const point = current.kind === 'table-cell-range'
      ? (rangeCell?.paragraphs[0]?.paraId ? { paraId: rangeCell.paragraphs[0].paraId, offset: 0 } : null)
      : current.kind === 'text-range'
        ? current.anchor
        : { paraId: current.paraId, offset: current.offset };
    if (!point) return false;
    const location = rangeLocation ?? findNativeTableCellLocation(doc, point.paraId);
    if (!location) return false;
    const beforeTable = cloneNativeWordBlocks([location.table])[0];
    if (!beforeTable || beforeTable.kind !== 'table') return false;
    const patch = command.kind === 'set-table-cell-background'
      ? { backgroundColor: command.color }
      : command.kind === 'set-table-cell-vertical-alignment'
        ? { verticalAlignment: command.alignment }
        : { borderColor: command.color };
    const result = current.kind === 'table-cell-range'
      ? nativeSetTableCellRangeFormat(doc, current.tableId, current, patch)
      : nativeSetTableCellFormat(doc, point.paraId, patch);
    if (!result) return false;
    const afterTable = cloneNativeWordBlocks([location.table])[0];
    if (!afterTable || afterTable.kind !== 'table') return false;
    const transaction: BrowserBlockTransaction = {
      source: 'local',
      range: { anchor: point, focus: point },
      container: location.container,
      startIndex: location.tableIndex,
      beforeBlocks: [beforeTable],
      afterBlocks: [afterTable],
      changesStart: changes.length,
      changes: 'changes' in result ? result.changes : [result.change],
      beforeSelection: current,
      afterSelection: current,
    };
    changes.push(...transaction.changes);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    const cursor = current.kind === 'caret' ? current : undefined;
    const textSelection = current.kind === 'text-range' ? current : undefined;
    if (!patchBlockTransactionDom(transaction, transaction.afterBlocks, cursor, textSelection)) render(cursor, current);
    else {
      restorePreservedSelection(current);
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, point.paraId);
      reflowFromParagraph(point.paraId, cursor, textSelection);
      restorePreservedSelection(current);
    }
    setSelection(current);
    notifyChange('deferred');
    return true;
  };

  const applyLocalTableCellClear = (current: Extract<WordEditorSelection, { kind: 'table-cell-range' }>): boolean => {
    const location = findNativeTableLocation(doc, current.tableId);
    if (!location) return false;
    const beforeTable = cloneNativeWordBlocks([location.table])[0];
    if (!beforeTable || beforeTable.kind !== 'table') return false;
    const result = nativeClearTableCellRange(doc, current.tableId, current);
    if (!result) return false;
    const afterTable = cloneNativeWordBlocks([location.table])[0];
    if (!afterTable || afterTable.kind !== 'table') return false;
    const transaction: BrowserBlockTransaction = {
      source: 'local',
      range: { anchor: result.cursor, focus: result.cursor },
      container: location.container,
      startIndex: location.tableIndex,
      beforeBlocks: [beforeTable],
      afterBlocks: [afterTable],
      changesStart: changes.length,
      changes: result.changes,
      beforeSelection: current,
      afterSelection: current,
    };
    changes.push(...result.changes);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    if (!patchBlockTransactionDom(transaction, transaction.afterBlocks)) render(undefined, current);
    else {
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, result.cursor.paraId);
      reflowFromParagraph(result.cursor.paraId);
    }
    setSelection(current);
    notifyChange('deferred');
    return true;
  };

  const applyLocalTableCellPaste = (
    current: Extract<WordEditorSelection, { kind: 'table-cell-range' }>,
    input: {
      text?: readonly (readonly string[])[];
      content?: readonly (readonly (NativeTableClipboardCell | null)[])[];
    },
  ): boolean => {
    const location = findNativeTableLocation(doc, current.tableId);
    if (!location) return false;
    const beforeTable = cloneNativeWordBlocks([location.table])[0];
    if (!beforeTable || beforeTable.kind !== 'table') return false;
    const anchor = {
      rowIndex: Math.min(current.anchor.rowIndex, current.focus.rowIndex),
      columnIndex: Math.min(current.anchor.columnIndex, current.focus.columnIndex),
    };
    const values = input.content ?? input.text ?? [];
    const result = input.content
      ? nativeReplaceTableCellRangeContent(doc, current.tableId, anchor, input.content)
      : nativeReplaceTableCellRangeText(doc, current.tableId, anchor, input.text ?? []);
    if (!result) return false;
    const afterTable = cloneNativeWordBlocks([location.table])[0];
    if (!afterTable || afterTable.kind !== 'table') return false;
    const rowCount = values.length;
    const columnCount = Math.max(1, ...values.map((row) => row.length));
    const afterSelection: WordEditorSelection = {
      kind: 'table-cell-range',
      tableId: current.tableId,
      anchor,
      focus: {
        rowIndex: Math.min(location.table.rows.length - 1, anchor.rowIndex + rowCount - 1),
        columnIndex: anchor.columnIndex + columnCount - 1,
      },
    };
    const transaction: BrowserBlockTransaction = {
      source: 'local',
      range: { anchor: result.cursor, focus: result.cursor },
      container: location.container,
      startIndex: location.tableIndex,
      beforeBlocks: [beforeTable],
      afterBlocks: [afterTable],
      changesStart: changes.length,
      changes: result.changes,
      beforeSelection: current,
      afterSelection,
    };
    changes.push(...result.changes);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    if (!patchBlockTransactionDom(transaction, transaction.afterBlocks)) render(undefined, afterSelection);
    else {
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, result.cursor.paraId);
      reflowFromParagraph(result.cursor.paraId);
    }
    setSelection(afterSelection);
    notifyChange('deferred');
    return true;
  };

  const restorePreservedSelection = (preserved: WordEditorSelection | null) => {
    if (!preserved) return;
    if (preserved.kind === 'text-range') {
      restoreNativeTextSelection(host, preserved.anchor, preserved.focus);
      return;
    }
    if (preserved.kind === 'object') {
      selectedImage = { paraId: preserved.paraId, drawingId: preserved.objectId };
      showImageResizeHandle();
      return;
    }
    if (preserved.kind === 'table-cell-range') {
      renderTableCellRangeSelection(preserved);
      return;
    }
    restoreNativeCursor(host, preserved);
  };

  const applyBrowserUndoFallback = (): boolean => {
    const transaction = browserUndoStack[browserUndoStack.length - 1];
    if (!transaction) return false;
    if (isBrowserBlockTransaction(transaction)) {
      const preservedSelection: WordEditorSelection = transaction.beforeSelection ?? {
        kind: 'text-range',
        anchor: transaction.range.anchor,
        focus: transaction.range.focus,
      };
      const cursor = preservedSelection.kind === 'caret' ? preservedSelection : undefined;
      const textSelection = preservedSelection.kind === 'text-range' ? preservedSelection : undefined;
      if (!replaceBlockTransaction(transaction, transaction.beforeBlocks, transaction.afterBlocks.length)) return false;
      if (!patchBlockTransactionDom(transaction, transaction.beforeBlocks, cursor, textSelection)) {
        render(cursor, preservedSelection);
      }
      restorePreservedSelection(preservedSelection);
      browserUndoStack.pop();
      browserRedoStack.push(transaction);
      activeLocalTypingTransaction = null;
      changes.splice(transaction.changesStart);
      setSelection(preservedSelection);
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, transaction.range.anchor.paraId);
      const paraId = preservedSelection.kind === 'text-range'
        ? preservedSelection.anchor.paraId
        : preservedSelection.kind === 'table-cell-range'
          ? transaction.range.anchor.paraId
          : preservedSelection.paraId;
      reflowFromParagraph(paraId, cursor, textSelection);
      notifyChange('deferred');
      return true;
    }
    if (
      transaction.source === 'browser' &&
      !transaction.skipDeferredRender &&
      !domParagraphsMatch(host, transaction.afterParagraphs)
    ) return false;
    const cursor = { paraId: transaction.beforeParagraphs[0]!.paraId, offset: transaction.range.anchor.offset };
    if (!replaceTransactionParagraphs(transaction, transaction.beforeParagraphs)) return false;
    const preservedSelection = formatTransactionSelection(transaction, 'before');
    if (transaction.forceFullRender) {
      browserUndoStack.pop();
      browserRedoStack.push(transaction);
      activeLocalTypingTransaction = null;
      changes.splice(transaction.changesStart);
      const restored = preservedSelection ?? { kind: 'caret' as const, ...cursor };
      setSelection(restored);
      render(restored.kind === 'caret' ? restored : undefined, restored);
      restorePreservedSelection(restored);
      notifyChange('deferred');
      return true;
    }
    if (!patchTransactionDom(transaction, transaction.beforeParagraphs, preservedSelection ? undefined : cursor)) return false;
    restorePreservedSelection(preservedSelection);
    browserUndoStack.pop();
    browserRedoStack.push(transaction);
    activeLocalTypingTransaction = null;
    changes.splice(transaction.changesStart);
    setSelection(preservedSelection ?? { kind: 'caret', ...cursor });
    if (transactionRequiresReflow(transaction) || transaction.skipDeferredRender) {
      const paraId = transaction.beforeParagraphs[0]!.paraId;
      if (transaction.deferLayoutReflow) scheduleImageLayoutReflow(paraId);
      else {
        reflowFromParagraph(
          paraId,
          preservedSelection ? undefined : cursor,
          preservedSelection?.kind === 'text-range' ? preservedSelection : undefined,
        );
      }
      restorePreservedSelection(preservedSelection);
      if (!transaction.deferLayoutReflow) scheduleReflow(paraId);
    }
    notifyChange('deferred');
    return true;
  };

  const applyBrowserRedoFallback = (): boolean => {
    const transaction = browserRedoStack[browserRedoStack.length - 1];
    if (!transaction) return false;
    if (isBrowserBlockTransaction(transaction)) {
      const preservedSelection: WordEditorSelection = transaction.afterSelection ?? {
        kind: 'caret',
        paraId: transaction.afterBlocks[0]!.kind === 'paragraph'
          ? transaction.afterBlocks[0]!.paraId
          : transaction.range.anchor.paraId,
        offset: transaction.range.anchor.offset,
      };
      const cursor = preservedSelection.kind === 'caret' ? preservedSelection : undefined;
      const textSelection = preservedSelection.kind === 'text-range' ? preservedSelection : undefined;
      if (!replaceBlockTransaction(transaction, transaction.afterBlocks, transaction.beforeBlocks.length)) return false;
      if (!patchBlockTransactionDom(transaction, transaction.afterBlocks, cursor, textSelection)) {
        render(cursor, preservedSelection);
      }
      restorePreservedSelection(preservedSelection);
      browserRedoStack.pop();
      browserUndoStack.push(transaction);
      activeLocalTypingTransaction = null;
      changes.splice(transaction.changesStart);
      changes.push(...transaction.changes);
      setSelection(preservedSelection);
      const paraId = preservedSelection.kind === 'text-range'
        ? preservedSelection.anchor.paraId
        : preservedSelection.kind === 'table-cell-range'
          ? transaction.range.anchor.paraId
          : preservedSelection.paraId;
      const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
      if (root) rebuildNativeTableFragmentsFromParagraph(root, doc, paraId);
      reflowFromParagraph(paraId, cursor, textSelection);
      notifyChange('deferred');
      return true;
    }
    if (
      transaction.source === 'browser' &&
      !transaction.skipDeferredRender &&
      !domParagraphsMatch(host, transaction.beforeParagraphs)
    ) return false;
    const cursor = { paraId: transaction.afterParagraphs[0]!.paraId, offset: transaction.range.anchor.offset };
    if (!replaceTransactionParagraphs(transaction, transaction.afterParagraphs)) return false;
    const preservedSelection = formatTransactionSelection(transaction, 'after');
    if (transaction.forceFullRender) {
      browserRedoStack.pop();
      browserUndoStack.push(transaction);
      activeLocalTypingTransaction = null;
      changes.splice(transaction.changesStart);
      changes.push(...transaction.changes);
      const restored = preservedSelection ?? { kind: 'caret' as const, ...cursor };
      setSelection(restored);
      render(restored.kind === 'caret' ? restored : undefined, restored);
      restorePreservedSelection(restored);
      notifyChange('deferred');
      return true;
    }
    if (!patchTransactionDom(transaction, transaction.afterParagraphs, preservedSelection ? undefined : cursor)) return false;
    restorePreservedSelection(preservedSelection);
    browserRedoStack.pop();
    browserUndoStack.push(transaction);
    activeLocalTypingTransaction = null;
    changes.splice(transaction.changesStart);
    changes.push(...transaction.changes);
    setSelection(preservedSelection ?? { kind: 'caret', ...cursor });
    if (transactionRequiresReflow(transaction) || transaction.skipDeferredRender) {
      const paraId = transaction.beforeParagraphs[0]!.paraId;
      if (transaction.deferLayoutReflow) scheduleImageLayoutReflow(paraId);
      else {
        reflowFromParagraph(
          paraId,
          preservedSelection ? undefined : cursor,
          preservedSelection?.kind === 'text-range' ? preservedSelection : undefined,
        );
      }
      restorePreservedSelection(preservedSelection);
      if (!transaction.deferLayoutReflow) scheduleReflow(paraId);
    }
    notifyChange('deferred');
    return true;
  };

  const undo = (): boolean => {
    if (destroyed || browserUndoStack.length === 0) return false;
    activate();
    return applyBrowserUndoFallback();
  };

  const redo = (): boolean => {
    if (destroyed || browserRedoStack.length === 0) return false;
    activate();
    return applyBrowserRedoFallback();
  };

  const setSelection = (next: WordEditorSelection | null) => {
    if (next) activate();
    if (next?.kind !== 'object') removeImageResizeHandle();
    renderTableCellRangeSelection(next);
    if (sameSelection(selection, next)) {
      renderPreservedSelectionHighlight();
      return;
    }
    selection = next;
    renderPreservedSelectionHighlight();
    notifySelectionChange();
  };

  const recordImageDeletion = (
    beforeParagraph: NativeWordParagraphNode | null,
    change: NativeWordChange | null,
    cursor: NativeCursorPoint,
  ): boolean => {
    if (!beforeParagraph || change?.kind !== 'image') return false;
    const afterParagraph = findNativeParagraph(doc, change.paraId);
    if (!afterParagraph) return false;
    const transaction: BrowserTextTransaction = {
      source: 'local',
      patchMode: 'individual',
      deferLayoutReflow: true,
      range: { anchor: cursor, focus: cursor },
      beforeParagraphs: [beforeParagraph],
      afterParagraphs: [cloneNativeParagraph(afterParagraph)],
      changesStart: changes.length,
      changes: [change],
      beforeSelection: {
        kind: 'object',
        paraId: change.paraId,
        objectKind: 'image',
        objectId: change.drawingId,
        offset: change.offset,
      },
      afterSelection: { kind: 'caret', ...cursor },
    };
    changes.push(change);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    removeNativeImageDom(change.paraId, change.drawingId, cursor);
    scheduleImageLayoutReflow(change.paraId);
    selectedImage = null;
    setSelection({ kind: 'caret', ...cursor });
    notifyChange('deferred');
    return true;
  };

  const deleteImageWithHistory = (paraId: string, drawingId: string, cursor: NativeCursorPoint): boolean => {
    const paragraph = findNativeParagraph(doc, paraId);
    const image = paragraph?.content.find(
      (child) => child.kind === 'image' && child.drawingId === drawingId,
    );
    if (trackRevisions && paragraph && image?.kind === 'image') {
      const beforeParagraph = cloneNativeParagraph(paragraph);
      const revisionId = allocateNativeRevisionId(doc);
      image.revision = {
        kind: 'revision',
        id: `revision:${revisionId}`,
        revisionId,
        revisionKind: 'delete',
        author: 'User',
        date: new Date().toISOString(),
      };
      return recordFormatting(
        {
          kind: 'object',
          paraId,
          objectKind: 'image',
          objectId: drawingId,
          offset: cursor.offset,
        },
        [beforeParagraph],
        [{ kind: 'image', op: 'delete', paraId, offset: cursor.offset, drawingId, revisionId, revisionKind: 'delete' }],
        true,
      );
    }
    const beforeParagraph = paragraph ? cloneNativeParagraph(paragraph) : null;
    return recordImageDeletion(beforeParagraph, nativeDeleteImage(doc, paraId, drawingId), cursor);
  };

  const applyLocalPageBreakEdit = (
    point: NativeCursorPoint,
    operation: 'insert' | 'delete-backward' | 'delete-forward',
    selectionBefore: WordEditorSelection = { kind: 'caret', ...point },
    selectionAfter?: WordEditorSelection,
  ): boolean => {
    const paragraph = findNativeParagraph(doc, point.paraId);
    if (!paragraph) return false;
    const targetOffset = operation === 'delete-backward' ? point.offset - 1 : point.offset;
    if (operation !== 'insert' && !nativePageBreakAt(doc, point.paraId, targetOffset)) return false;
    const beforeParagraph = cloneNativeParagraph(paragraph);
    const change = operation === 'insert'
      ? nativeInsertPageBreak(doc, point.paraId, point.offset)
      : operation === 'delete-backward'
        ? nativeDeleteBackward(doc, point.paraId, point.offset)
        : nativeDeleteForward(doc, point.paraId, point.offset);
    const afterParagraph = findNativeParagraph(doc, point.paraId);
    if (!change || change.kind !== 'page-break' || !afterParagraph) return false;
    const mutationCursor = {
      paraId: point.paraId,
      offset: operation === 'insert'
        ? point.offset + 1
        : operation === 'delete-backward'
          ? Math.max(0, point.offset - 1)
          : point.offset,
    };
    const afterSelection = selectionAfter ?? { kind: 'caret' as const, ...mutationCursor };
    const cursor = afterSelection.kind === 'caret' ? afterSelection : mutationCursor;
    const transaction: BrowserTextTransaction = {
      source: 'local',
      patchMode: 'individual',
      forceFullRender: true,
      range: { anchor: point, focus: point },
      beforeParagraphs: [beforeParagraph],
      afterParagraphs: [cloneNativeParagraph(afterParagraph)],
      changesStart: changes.length,
      changes: [change],
      beforeSelection: selectionBefore,
      afterSelection,
    };
    changes.push(change);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    render(cursor, afterSelection);
    setSelection(afterSelection);
    notifyChange('deferred');
    return true;
  };

  const applyLocalPageBreakRangeDelete = (
    current: Extract<WordEditorSelection, { kind: 'text-range' }>,
  ): boolean => {
    if (current.anchor.paraId !== current.focus.paraId) return false;
    const range = normalizedSelectionRange(current.anchor.offset, current.focus.offset);
    if (!nativeRangeContainsPageBreak(doc, current.anchor.paraId, range.from, range.to)) return false;
    const paragraph = findNativeParagraph(doc, current.anchor.paraId);
    if (!paragraph) return false;
    const beforeParagraph = cloneNativeParagraph(paragraph);
    const rangeChanges = nativeDeleteRange(doc, current.anchor.paraId, range.from, range.to);
    const afterParagraph = findNativeParagraph(doc, current.anchor.paraId);
    if (!afterParagraph || !rangeChanges.some((change) => change.kind === 'page-break')) return false;
    const afterSelection: WordEditorSelection = {
      kind: 'caret',
      paraId: current.anchor.paraId,
      offset: range.from,
    };
    const transaction: BrowserTextTransaction = {
      source: 'local',
      patchMode: 'individual',
      forceFullRender: true,
      range: { anchor: current.anchor, focus: current.focus },
      beforeParagraphs: [beforeParagraph],
      afterParagraphs: [cloneNativeParagraph(afterParagraph)],
      changesStart: changes.length,
      changes: rangeChanges,
      beforeSelection: current,
      afterSelection,
    };
    changes.push(...rangeChanges);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    render(afterSelection, afterSelection);
    setSelection(afterSelection);
    notifyChange('deferred');
    return true;
  };

  const deletePageBreakAcrossParagraphBoundary = (
    current: Extract<WordEditorSelection, { kind: 'caret' }>,
    direction: 'backward' | 'forward',
  ): boolean => {
    const location = findNativeParagraphLocation(doc, current.paraId);
    if (!location) return false;
    if (direction === 'backward' && current.offset === 0 && location.previous) {
      const previousLength = nativeParagraphText(location.previous).length;
      if (!nativePageBreakAt(doc, location.previous.paraId, previousLength - 1)) return false;
      return applyLocalPageBreakEdit(
        { paraId: location.previous.paraId, offset: previousLength },
        'delete-backward',
        current,
        current,
      );
    }
    const nextBlock = location.container[location.index + 1];
    const nextParagraph = nextBlock?.kind === 'paragraph' ? nextBlock : null;
    if (direction === 'forward' && current.offset === nativeParagraphText(location.paragraph).length && nextParagraph) {
      if (!nativePageBreakAt(doc, nextParagraph.paraId, 0)) return false;
      return applyLocalPageBreakEdit(
        { paraId: nextParagraph.paraId, offset: 0 },
        'delete-forward',
        current,
        current,
      );
    }
    return false;
  };

  const applyLocalParagraphTextEdit = (
    range: PendingBrowserRangeDelete['range'],
    beforeSelection: WordEditorSelection,
    afterCursor: NativeCursorPoint,
    mutate: () => NativeWordChange | NativeWordChange[] | null,
  ): boolean => {
    if (range.anchor.paraId !== range.focus.paraId) return false;
    const paragraph = findNativeParagraph(doc, range.anchor.paraId);
    if (!paragraph) return false;
    const beforeParagraph = cloneNativeParagraph(paragraph);
    const result = mutate();
    const editChanges = Array.isArray(result) ? result : result ? [result] : [];
    const afterParagraph = findNativeParagraph(doc, range.anchor.paraId);
    if (editChanges.length === 0 || !afterParagraph) return false;
    const afterSelection: WordEditorSelection = { kind: 'caret', ...afterCursor };
    const transaction: BrowserTextTransaction = {
      source: 'local',
      patchMode: 'individual',
      range,
      beforeParagraphs: [beforeParagraph],
      afterParagraphs: [cloneNativeParagraph(afterParagraph)],
      changesStart: changes.length,
      changes: editChanges,
      beforeSelection,
      afterSelection,
    };
    changes.push(...editChanges);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    if (!patchTransactionDom(transaction, transaction.afterParagraphs, afterCursor)) render(afterCursor);
    else reflowFromParagraph(range.anchor.paraId, afterCursor);
    setSelection(afterSelection);
    notifyChange('deferred');
    return true;
  };

  const selectedTextRanges = (
    current: WordEditorSelection | null,
  ): Array<{ paraId: string; from: number; to: number }> => {
    if (current?.kind !== 'text-range') return [];
    const ordered = orderedSelectionRange(doc, current);
    if (ordered.anchor.paraId === ordered.focus.paraId) {
      return [{ paraId: ordered.anchor.paraId, from: ordered.anchor.offset, to: ordered.focus.offset }];
    }
    const start = findNativeParagraphLocation(doc, ordered.anchor.paraId);
    const end = findNativeParagraphLocation(doc, ordered.focus.paraId);
    if (!start || !end || start.container !== end.container) return [];
    const ranges: Array<{ paraId: string; from: number; to: number }> = [];
    for (let index = start.index; index <= end.index; index += 1) {
      const block = start.container[index];
      if (block?.kind !== 'paragraph') return [];
      ranges.push({
        paraId: block.paraId,
        from: index === start.index ? ordered.anchor.offset : 0,
        to: index === end.index ? ordered.focus.offset : nativeParagraphText(block).length,
      });
    }
    return ranges;
  };

  const selectedParagraphIds = (current: WordEditorSelection | null): string[] => {
    if (!current) return [];
    if (current.kind === 'caret' || current.kind === 'object') return [current.paraId];
    return selectedTextRanges(current).map((range) => range.paraId);
  };

  const clipboardTextForSelection = (current: WordEditorSelection | null): string | null => {
    if (current?.kind !== 'text-range') return null;
    const ranges = selectedTextRanges(current);
    if (ranges.length === 0) return null;
    return ranges
      .map((range) => nativeParagraphText(findNativeParagraph(doc, range.paraId)!).slice(range.from, range.to))
      .join('\n');
  };

  const clipboardParagraphsForSelection = (current: WordEditorSelection | null): NativeClipboardParagraph[] => {
    return selectedTextRanges(current).map((range) => {
      const paragraph = findNativeParagraph(doc, range.paraId)!;
      const text = nativeParagraphText(paragraph).slice(range.from, range.to);
      const runs: NativeClipboardParagraph['runs'] = [];
      let offset = 0;
      for (const child of paragraph.content) {
        const childText = nativeClipboardInlineText(child);
        const childEnd = offset + childText.length;
        const from = Math.max(range.from, offset);
        const to = Math.min(range.to, childEnd);
        const format = child.kind === 'text' || child.kind === 'field' ? child.format : null;
        const hyperlink = child.kind === 'text' ? child.hyperlink : undefined;
        if ((format || hyperlink) && to > from) {
          runs.push({
            from: from - range.from,
            to: to - range.from,
            format: { ...(format ?? {}) },
            ...(hyperlink ? { hyperlink: { ...hyperlink } } : {}),
          });
        }
        offset = childEnd;
      }
      const listType =
        paragraph.format?.hasNumbering && paragraph.format.numbering
          ? nativeListType(paragraph.format.numbering)
          : null;
      return {
        text,
        runs,
        ...(listType ? { listType } : {}),
      };
    });
  };

  const applyLocalPlainTextReplacement = (
    range: PendingBrowserRangeDelete['range'],
    text: string,
    richParagraphs?: NativeClipboardParagraph[],
  ): boolean => {
    if (trackRevisions && !richParagraphs && range.anchor.paraId === range.focus.paraId) {
      return dispatch({
        kind: 'replace-range',
        anchor: range.anchor,
        focus: range.focus,
        text,
      });
    }
    if (applyLocalBlockRangeReplacement(range, text)) return true;
    const beforeParagraphs = editableRangeParagraphs(range);
    if (!beforeParagraphs) return false;
    const result = nativeReplaceTextRangeWithParagraphs(doc, range.anchor, range.focus, text);
    if (!result || result.changes.length === 0) return false;
    for (const paraId of result.paragraphIds) paragraphIds.add(paraId);
    let activeList: { type: 'bullet' | 'numbered'; numbering: ReturnType<typeof createNativeListNumbering> } | null = null;
    for (let index = 0; index < (richParagraphs?.length ?? 0); index += 1) {
      const richParagraph = richParagraphs![index]!;
      const paraId = result.paragraphIds[index];
      if (!paraId) break;
      const baseOffset = index === 0 ? range.anchor.offset : 0;
      for (const run of richParagraph.runs) {
        if (Object.keys(run.format).length > 0) {
          const change = nativeSetRunFormat(doc, paraId, baseOffset + run.from, baseOffset + run.to, run.format);
          if (change) result.changes.push(change);
        }
        if (run.hyperlink) {
          const hyperlinkChange = nativeSetHyperlink(
            doc,
            paraId,
            baseOffset + run.from,
            baseOffset + run.to,
            run.hyperlink,
          );
          if (hyperlinkChange) result.changes.push(hyperlinkChange);
        }
      }
      if (richParagraph.listType) {
        if (!activeList || activeList.type !== richParagraph.listType) {
          activeList = { type: richParagraph.listType, numbering: createNativeListNumbering(doc, richParagraph.listType) };
        }
        const change = nativeSetParagraphList(doc, paraId, activeList.numbering);
        if (change) result.changes.push(change);
      } else {
        activeList = null;
      }
    }
    if (trackRevisions) {
      const revisionId = allocateNativeRevisionId(doc);
      const insertedChanges = result.changes
        .filter((change): change is Extract<NativeWordChange, { kind: 'text' }> =>
          change.kind === 'text' && (change.op === 'insert' || change.op === 'replace') && change.text.length > 0,
        )
        .map((change) => nativeAddRevision(doc, change.paraId, change.offset, change.offset + change.text.length, {
          revisionId,
          revisionKind: 'insert',
          author: 'User',
          date: new Date().toISOString(),
        }))
        .filter((change): change is NativeWordChange => change !== null);
      result.changes.push(...insertedChanges);
    }
    refreshNativeNumberingLabels(doc);
    const afterParagraphs = result.paragraphIds
      .map((paraId) => findNativeParagraph(doc, paraId))
      .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
      .map(cloneNativeParagraph);
    const transaction: BrowserTextTransaction = {
      source: 'local',
      range,
      beforeParagraphs,
      afterParagraphs,
      changesStart: changes.length,
      changes: result.changes,
    };
    changes.push(...result.changes);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    if (!patchTransactionDom(transaction, afterParagraphs, result.cursor)) render(result.cursor);
    else reflowFromParagraph(beforeParagraphs[0]!.paraId, result.cursor);
    setSelection({ kind: 'caret', ...result.cursor });
    notifyChange('deferred');
    return true;
  };

  const applyLocalImageReplacement = (
    range: PendingBrowserRangeDelete['range'],
    src: string,
    options: { width?: number; height?: number; select?: boolean } = {},
  ): boolean => {
    const beforeParagraphs = editableRangeParagraphs(range);
    if (!beforeParagraphs) return false;
    const replacement = nativeReplaceTextRangeWithParagraphs(doc, range.anchor, range.focus, '');
    if (!replacement) return false;
    for (const paraId of replacement.paragraphIds) paragraphIds.add(paraId);
    const imageChange = nativeInsertImage(doc, replacement.cursor.paraId, replacement.cursor.offset, src, options);
    if (!imageChange || imageChange.kind !== 'image') return false;
    const transactionChanges = [...replacement.changes, imageChange];
    if (trackRevisions) {
      const revisionId = allocateNativeRevisionId(doc);
      const insertedImage = findNativeParagraph(doc, imageChange.paraId)?.content.find(
        (child) => child.kind === 'image' && child.drawingId === imageChange.drawingId,
      );
      if (insertedImage?.kind === 'image') {
        insertedImage.revision = {
          kind: 'revision',
          id: `revision:${revisionId}`,
          revisionId,
          revisionKind: 'insert',
          author: 'User',
          date: new Date().toISOString(),
        };
        transactionChanges.push({
          kind: 'image',
          op: 'insert',
          paraId: imageChange.paraId,
          offset: imageChange.offset,
          drawingId: imageChange.drawingId,
          revisionId,
          revisionKind: 'insert',
        });
      }
    }
    const cursor = { paraId: replacement.cursor.paraId, offset: replacement.cursor.offset + 1 };
    const afterParagraphs = replacement.paragraphIds
      .map((paraId) => findNativeParagraph(doc, paraId))
      .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
      .map(cloneNativeParagraph);
    const transaction: BrowserTextTransaction = {
      source: 'local',
      patchMode: 'individual',
      deferLayoutReflow: true,
      range,
      beforeParagraphs,
      afterParagraphs,
      changesStart: changes.length,
      changes: transactionChanges,
      ...(options.select
        ? {
            beforeSelection:
              range.anchor.paraId === range.focus.paraId && range.anchor.offset === range.focus.offset
                ? { kind: 'caret' as const, ...range.anchor }
                : { kind: 'text-range' as const, anchor: range.anchor, focus: range.focus },
            afterSelection: {
              kind: 'object' as const,
              paraId: replacement.cursor.paraId,
              objectKind: 'image' as const,
              objectId: imageChange.drawingId,
              offset: replacement.cursor.offset,
            },
          }
        : {}),
    };
    changes.push(...transactionChanges);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    selectedImage = options.select
      ? { paraId: replacement.cursor.paraId, drawingId: imageChange.drawingId }
      : null;
    activeLocalTypingTransaction = null;
    const insertedImage = findNativeParagraph(doc, imageChange.paraId)?.content.find(
      (child) => child.kind === 'image' && child.drawingId === imageChange.drawingId,
    );
    if (
      !insertedImage ||
      insertedImage.kind !== 'image' ||
      !patchImageReplacementDom(range, insertedImage, cursor)
    ) {
      if (!patchTransactionDom(transaction, afterParagraphs, cursor)) render(cursor);
    }
    scheduleImageLayoutReflow(beforeParagraphs[0]!.paraId);
    if (selectedImage) showImageResizeHandle();
    setSelection(
      selectedImage
        ? {
            kind: 'object',
            paraId: selectedImage.paraId,
            objectKind: 'image',
            objectId: selectedImage.drawingId,
            offset: replacement.cursor.offset,
          }
        : { kind: 'caret', ...cursor },
    );
    notifyChange('deferred');
    return true;
  };

  const patchImageReplacementDom = (
    range: PendingBrowserRangeDelete['range'],
    image: Extract<NativeWordInlineNode, { kind: 'image' }>,
    cursor: NativeCursorPoint,
  ): boolean => {
    if (range.anchor.paraId !== range.focus.paraId) return false;
    const domRange = createNativeSelectionRange(host, range.anchor, range.focus);
    if (!domRange) return false;
    const paragraph = host.querySelector<HTMLElement>(
      `coco-word-native-p[data-para-id="${range.anchor.paraId}"]`,
    );
    if (!paragraph || !paragraph.contains(domRange.commonAncestorContainer)) return false;
    domRange.deleteContents();
    domRange.insertNode(renderNativeImage(image));
    restoreNativeCursor(host, cursor);
    return true;
  };

  const createRuntimeParaId = (): string => {
    let paraId = '';
    do {
      nextParagraphId = (nextParagraphId + 1) >>> 0;
      paraId = nextParagraphId.toString(16).toUpperCase().padStart(8, '0');
    } while (!paraId || paragraphIds.has(paraId));
    paragraphIds.add(paraId);
    return paraId;
  };

  const splitParagraphLocally = (point: NativeCursorPoint, nextParaId = createRuntimeParaId()): boolean => {
    const paragraph = findNativeParagraph(doc, point.paraId);
    if (!paragraph) return false;
    const beforeParagraph = cloneNativeParagraph(paragraph);
    const change = nativeSplitParagraph(doc, point.paraId, point.offset, nextParaId);
    if (!change) return false;
    const afterParagraphs = [point.paraId, nextParaId]
      .map((paraId) => findNativeParagraph(doc, paraId))
      .filter((candidate): candidate is NativeWordParagraphNode => candidate !== null)
      .map(cloneNativeParagraph);
    if (afterParagraphs.length !== 2) return false;
    const cursor = { paraId: nextParaId, offset: 0 };
    const transaction: BrowserTextTransaction = {
      source: 'local',
      skipDeferredRender: true,
      range: { anchor: point, focus: point },
      beforeParagraphs: [beforeParagraph],
      afterParagraphs,
      changesStart: changes.length,
      changes: [change],
    };
    changes.push(change);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    if (!patchTransactionDom(transaction, afterParagraphs, cursor)) render(cursor);
    else reflowFromParagraph(point.paraId, cursor);
    setSelection({ kind: 'caret', ...cursor });
    notifyChange('deferred');
    return true;
  };

  const mergeParagraphBackwardLocally = (point: NativeCursorPoint): boolean => {
    if (point.offset !== 0) return false;
    const location = findNativeParagraphLocation(doc, point.paraId);
    const previous = location?.previous;
    if (!location || !previous) return false;
    const beforeParagraphs = [cloneNativeParagraph(previous), cloneNativeParagraph(location.paragraph)];
    const cursor = { paraId: previous.paraId, offset: nativeParagraphText(previous).length };
    const changesStart = changes.length;
    const change = nativeDeleteBackward(doc, point.paraId, 0);
    const merged = findNativeParagraph(doc, previous.paraId);
    if (!change || change.kind !== 'block' || !merged) return false;
    const transaction: BrowserTextTransaction = {
      source: 'local',
      skipDeferredRender: true,
      range: { anchor: point, focus: point },
      beforeParagraphs,
      afterParagraphs: [cloneNativeParagraph(merged)],
      changesStart,
      changes: [change],
    };
    changes.push(change);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    if (!patchTransactionDom(transaction, transaction.afterParagraphs, cursor)) render(cursor);
    else reflowFromParagraph(previous.paraId, cursor);
    setSelection({ kind: 'caret', ...cursor });
    notifyChange('deferred');
    return true;
  };

  const mergeParagraphForwardLocally = (point: NativeCursorPoint): boolean => {
    const location = findNativeParagraphLocation(doc, point.paraId);
    if (!location || point.offset !== nativeParagraphText(location.paragraph).length) return false;
    const next = location.container[location.index + 1];
    if (next?.kind !== 'paragraph') return false;
    const beforeParagraphs = [cloneNativeParagraph(location.paragraph), cloneNativeParagraph(next)];
    const changesStart = changes.length;
    const change = nativeDeleteForward(doc, point.paraId, point.offset);
    const merged = findNativeParagraph(doc, point.paraId);
    if (!change || change.kind !== 'block' || !merged) return false;
    const transaction: BrowserTextTransaction = {
      source: 'local',
      skipDeferredRender: true,
      range: { anchor: point, focus: point },
      beforeParagraphs,
      afterParagraphs: [cloneNativeParagraph(merged)],
      changesStart,
      changes: [change],
    };
    changes.push(change);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    selectedImage = null;
    if (!patchTransactionDom(transaction, transaction.afterParagraphs, point)) render(point);
    else reflowFromParagraph(point.paraId, point);
    setSelection({ kind: 'caret', ...point });
    notifyChange('deferred');
    return true;
  };

  const recordFormatting = (
    current: WordEditorSelection,
    beforeParagraphs: NativeWordParagraphNode[],
    formattingChanges: NativeWordChange[],
    reflow = false,
  ): boolean => {
    if (current.kind === 'table-cell-range') return false;
    if (formattingChanges.length === 0) return false;
    if (beforeParagraphs.length === 0) return false;
    const afterParagraphs = beforeParagraphs
      .map((paragraph) => findNativeParagraph(doc, paragraph.paraId))
      .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
      .map(cloneNativeParagraph);
    const range =
      current.kind === 'text-range'
        ? orderedSelectionRange(doc, current)
        : { anchor: { paraId: current.paraId, offset: current.offset }, focus: { paraId: current.paraId, offset: current.offset } };
    const transaction: BrowserTextTransaction = {
      source: 'local',
      patchMode: 'individual',
      range,
      beforeParagraphs,
      afterParagraphs,
      changesStart: changes.length,
      changes: formattingChanges,
      beforeSelection: current,
      afterSelection: current,
    };
    changes.push(...formattingChanges);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    if (!patchTransactionDom(transaction, afterParagraphs)) render();
    const root = host.querySelector<HTMLElement>(':scope > coco-word-native-document');
    if (root) refreshNativePageCommentBubbles(root);
    if (current.kind === 'text-range') restoreNativeTextSelection(host, current.anchor, current.focus);
    else if (current.kind === 'object') {
      selectedImage = { paraId: current.paraId, drawingId: current.objectId };
      showImageResizeHandle();
    } else restoreNativeCursor(host, { paraId: current.paraId, offset: current.offset });
    if (reflow) {
      const cursor = current.kind === 'text-range' ? undefined : { paraId: current.paraId, offset: current.offset };
      reflowFromParagraph(beforeParagraphs[0]!.paraId, cursor, current);
    }
    if (current.kind === 'object' && selectedImage) showImageResizeHandle();
    setSelection(current);
    notifyChange('deferred');
    return true;
  };

  const applySearchReplacement = (
    searchMatches: ReadonlyArray<WordEditorSearchState['matches'][number]>,
    replacement: string,
  ): number => {
    if (searchMatches.length === 0) return 0;
    const paraIds = Array.from(new Set(searchMatches.map((match) => match.paraId)));
    const beforeParagraphs = paraIds
      .map((paraId) => findNativeParagraph(doc, paraId))
      .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
      .map(cloneNativeParagraph);
    if (beforeParagraphs.length === 0) return 0;
    const changesStart = changes.length;
    const replacementChanges: NativeWordChange[] = [];
    const ordered = [...searchMatches].sort((left, right) =>
      left.paraId === right.paraId ? right.from - left.from : paraIds.indexOf(right.paraId) - paraIds.indexOf(left.paraId));
    for (const match of ordered) {
      replacementChanges.push(...nativeReplaceRange(doc, match.paraId, match.from, match.to, replacement));
    }
    if (replacementChanges.length === 0) return 0;
    const afterParagraphs = beforeParagraphs
      .map((paragraph) => findNativeParagraph(doc, paragraph.paraId))
      .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
      .map(cloneNativeParagraph);
    const firstMatch = searchMatches[0]!;
    const transaction: BrowserTextTransaction = {
      source: 'local',
      patchMode: 'individual',
      range: {
        anchor: { paraId: firstMatch.paraId, offset: firstMatch.from },
        focus: { paraId: firstMatch.paraId, offset: firstMatch.to },
      },
      beforeParagraphs,
      afterParagraphs,
      changesStart,
      changes: replacementChanges,
      ...(selection ? { beforeSelection: selection, afterSelection: selection } : {}),
    };
    changes.push(...replacementChanges);
    browserUndoStack.push(transaction);
    browserRedoStack.length = 0;
    activeLocalTypingTransaction = null;
    if (!patchTransactionDom(transaction, afterParagraphs)) render(undefined, selection ?? undefined);
    else restorePreservedSelection(selection);
    reflowFromParagraph(firstMatch.paraId);
    notifyChange();
    return searchMatches.length;
  };

  const replaceCurrent = (replacement: string): boolean => {
    const match = searchState && searchState.activeIndex >= 0 ? searchState.matches[searchState.activeIndex] : null;
    return match ? applySearchReplacement([match], replacement) > 0 : false;
  };

  const replaceAll = (replacement: string): number => {
    return searchState ? applySearchReplacement(searchState.matches, replacement) : 0;
  };

  const dispatch = (command: WordEditorCommand): boolean => {
    if (destroyed) return false;
    activate();
    const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
    if (command.kind === 'toggle-track-revisions') {
      trackRevisions = !trackRevisions;
      notifySelectionChange();
      return true;
    }
    if (
      command.kind === 'accept-revision' ||
      command.kind === 'reject-revision' ||
      command.kind === 'accept-all-revisions' ||
      command.kind === 'reject-all-revisions'
    ) {
      const revisionId = command.kind === 'accept-revision' || command.kind === 'reject-revision'
        ? command.revisionId
        : undefined;
      const revisionKind = command.kind === 'accept-revision' || command.kind === 'reject-revision'
        ? command.revisionKind
        : undefined;
      const resolveAll = command.kind === 'accept-all-revisions' || command.kind === 'reject-all-revisions';
      const resolvedTableRevisionIds = new Set(
        browserUndoStack.flatMap((transaction) => isBrowserBlockTransaction(transaction)
          ? transaction.changes
            .filter((change): change is Extract<NativeWordChange, { kind: 'table' }> => change.kind === 'table' && Boolean(change.revisionAction && change.revisionId))
            .map((change) => change.revisionId!)
          : []),
      );
      const finishResolution = (handled: boolean): boolean => {
        if (handled && resolveAll) dispatch(command);
        return handled;
      };
      const tableRevisionTransactions = browserUndoStack.filter((transaction): transaction is BrowserBlockTransaction =>
        isBrowserBlockTransaction(transaction) && transaction.changes.some((change) => change.kind === 'table' && !change.revisionAction && Boolean(change.revisionId) && !resolvedTableRevisionIds.has(change.revisionId!)) && transaction.afterBlocks.some((block) => {
          if (block.kind !== 'table' || block.revision?.kind !== 'revision' || resolvedTableRevisionIds.has(block.revision.revisionId)) return false;
          const currentRevisionId = findNativeTableLocation(doc, block.tableId)?.table.revision?.revisionId;
          if (resolveAll) return Boolean(currentRevisionId);
          return currentRevisionId === block.revision.revisionId &&
            block.revision.revisionId === revisionId &&
            (!revisionKind || block.revision.revisionKind === revisionKind);
        }),
      );
      if (tableRevisionTransactions.length > 0) {
        const operation: 'accept' | 'reject' = command.kind === 'accept-revision' || command.kind === 'accept-all-revisions' ? 'accept' : 'reject';
        const tableSources = new Map<string, BrowserBlockTransaction[]>();
        for (const sourceTransaction of tableRevisionTransactions) {
          const markedTable = sourceTransaction.afterBlocks.find((block) => block.kind === 'table' && block.revision?.kind === 'revision');
          if (!markedTable || markedTable.kind !== 'table' || !markedTable.revision) continue;
          const sources = tableSources.get(markedTable.tableId) ?? [];
          sources.push(sourceTransaction);
          tableSources.set(markedTable.tableId, sources);
        }
        for (const [tableId, sources] of tableSources) {
          const sourceTransaction = sources[sources.length - 1];
          if (!sourceTransaction) continue;
          const markedTable = sourceTransaction.afterBlocks.find((block) => block.kind === 'table' && block.revision?.kind === 'revision');
          if (!markedTable || markedTable.kind !== 'table' || !markedTable.revision) continue;
          const location = findNativeTableLocation(doc, tableId);
          if (!location) continue;
          const beforeBlocks = [cloneNativeWordBlocks([location.table])[0]!];
          const resolvedTable = cloneNativeWordBlocks([location.table])[0];
          if (!resolvedTable || resolvedTable.kind !== 'table') continue;
          if (operation === 'accept') {
            resolvedTable.revision = undefined;
            location.table.revision = undefined;
          }
          const afterBlocks = operation === 'accept'
            ? [resolvedTable]
            : cloneNativeWordBlocks(sources[0]?.beforeBlocks ?? sourceTransaction.beforeBlocks);
          if (operation === 'reject') {
            location.container.splice(location.tableIndex, 1, ...cloneNativeWordBlocks(afterBlocks));
          }
          const resolvedChanges = sources.flatMap((source) => source.changes
            .filter((change): change is Extract<NativeWordChange, { kind: 'table' }> => change.kind === 'table')
            .map((change) => ({ ...change, revisionAction: operation })));
          const transaction: BrowserBlockTransaction = {
            source: 'local',
            range: sourceTransaction.range,
            container: location.container,
            startIndex: location.tableIndex,
            beforeBlocks,
            afterBlocks,
            changesStart: changes.length,
            changes: resolvedChanges,
            beforeSelection: current ?? undefined,
            afterSelection: current ?? undefined,
          };
          changes.push(...transaction.changes);
          browserUndoStack.push(transaction);
          browserRedoStack.length = 0;
          if (!patchBlockTransactionDom(transaction, afterBlocks, undefined)) render();
          else reflowFromParagraph(beforeBlocks[0]?.kind === 'paragraph' ? beforeBlocks[0].paraId : location.table.rows[0]?.cells[0]?.paragraphs[0]?.paraId ?? '', undefined);
          notifyChange('deferred');
        }
        return finishResolution(true);
      }
      const imageTargets = allNativeParagraphs(doc).flatMap((paragraph) => {
        let offset = 0;
        const result: Array<{ paraId: string; drawingId: string; offset: number }> = [];
        for (const child of paragraph.content) {
          if (child.kind === 'image' &&
            ((child.revision?.kind === 'revision' && (resolveAll || (child.revision.revisionId === revisionId && (!revisionKind || child.revision.revisionKind === revisionKind)))) ||
              (child.formatRevision && (resolveAll || (child.formatRevision.revisionId === revisionId && (!revisionKind || revisionKind === 'format')))))) {
            result.push({ paraId: paragraph.paraId, drawingId: child.drawingId, offset });
          }
          offset += child.kind === 'text' ? child.text.length : 1;
        }
        return result;
      });
      if (imageTargets.length > 0) {
        const operation = command.kind === 'accept-revision' || command.kind === 'accept-all-revisions' ? 'accept' : 'reject';
        const beforeParagraphs = Array.from(new Set(imageTargets.map((target) => target.paraId)))
          .map((paraId) => findNativeParagraph(doc, paraId))
          .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
          .map(cloneNativeParagraph);
        const imageChanges = imageTargets
          .map((target) => {
            const image = findNativeParagraph(doc, target.paraId)?.content.find(
              (child) => child.kind === 'image' && child.drawingId === target.drawingId,
            );
            return image?.kind === 'image' && image.formatRevision
              ? nativeResolveImageFormatRevision(doc, target.paraId, target.drawingId, operation)
              : nativeResolveImageRevision(doc, target.paraId, target.drawingId, operation);
          })
          .filter((change): change is NativeWordChange => change !== null);
        if (imageChanges.length === 0) return false;
        const target = imageTargets[0]!;
        return finishResolution(recordFormatting(
          { kind: 'caret', paraId: target.paraId, offset: target.offset },
          beforeParagraphs,
          imageChanges,
          true,
        ));
      }
      const formatRanges = allNativeParagraphs(doc).flatMap((paragraph) => {
        let offset = 0;
        const result: Array<{ paraId: string; from: number; to: number }> = [];
        for (const child of paragraph.content) {
          const length = child.kind === 'text' ? child.text.length : 1;
          if (child.kind === 'text' && child.formatRevision &&
            (resolveAll || (child.formatRevision.revisionId === revisionId && (!revisionKind || revisionKind === 'format')))) {
            result.push({ paraId: paragraph.paraId, from: offset, to: offset + length });
          }
          offset += length;
        }
        return result;
      });
      if (formatRanges.length > 0) {
        const operation = command.kind === 'accept-revision' || command.kind === 'accept-all-revisions' ? 'accept' : 'reject';
        const beforeParagraphs = Array.from(new Set(formatRanges.map((range) => range.paraId)))
          .map((paraId) => findNativeParagraph(doc, paraId))
          .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
          .map(cloneNativeParagraph);
        const formatChanges = formatRanges
          .map((range) => nativeResolveFormatRevision(doc, range.paraId, range.from, range.to, operation))
          .filter((change): change is NativeWordChange => change !== null);
        if (formatChanges.length === 0) return false;
        const first = formatRanges[0]!;
        return finishResolution(recordFormatting(
          { kind: 'caret', paraId: first.paraId, offset: first.from },
          beforeParagraphs,
          formatChanges,
          true,
        ));
      }
      const ranges = resolveAll
        ? allNativeParagraphs(doc).flatMap((paragraph) => {
            const result: Array<{ paraId: string; from: number; to: number }> = [];
            let offset = 0;
            for (const child of paragraph.content) {
              const length = child.kind === 'text' ? child.text.length : 1;
              if (child.kind === 'text' && child.marks?.some((mark) => mark.kind === 'revision')) {
                result.push({ paraId: paragraph.paraId, from: offset, to: offset + length });
              }
              offset += length;
            }
            return result;
          })
        : revisionId
        ? allNativeParagraphs(doc).flatMap((paragraph) => {
            const result: Array<{ paraId: string; from: number; to: number }> = [];
            let offset = 0;
            for (const child of paragraph.content) {
              const length = child.kind === 'text' ? child.text.length : 1;
              if (child.kind === 'text' && child.marks?.some((mark) => mark.kind === 'revision' && mark.revisionId === revisionId && (!revisionKind || mark.revisionKind === revisionKind))) {
                result.push({ paraId: paragraph.paraId, from: offset, to: offset + length });
              }
              offset += length;
            }
            return result;
          })
        : selectedTextRanges(current);
      if (ranges.length === 0) return false;
      if (resolveAll || revisionId) {
        ranges.sort((left, right) => left.paraId === right.paraId
          ? right.from - left.from || right.to - left.to
          : right.paraId.localeCompare(left.paraId));
      }
      const beforeParagraphs = ranges
        .map((range) => findNativeParagraph(doc, range.paraId))
        .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
        .map(cloneNativeParagraph);
      const operation = command.kind === 'accept-revision' || command.kind === 'accept-all-revisions' ? 'accept' : 'reject';
      const revisionChanges = ranges
        .map((range) => nativeResolveRevision(doc, range.paraId, range.from, range.to, operation))
        .filter((change): change is NativeWordChange => change !== null);
      if (revisionChanges.length === 0) return false;
      const preservedSelection = current ?? { kind: 'caret' as const, paraId: ranges[0]!.paraId, offset: ranges[0]!.from };
      return finishResolution(recordFormatting(preservedSelection, beforeParagraphs, revisionChanges, false));
    }
    if (command.kind === 'insert-text' && current?.kind === 'caret') {
      const cursor = { paraId: current.paraId, offset: current.offset + command.text.length };
      return applyLocalParagraphTextEdit(
        { anchor: current, focus: current },
        current,
        cursor,
        () => {
          const inserted = nativeInsertText(doc, current.paraId, current.offset, command.text);
          if (!trackRevisions || !inserted || command.text.length === 0) return inserted;
          const marked = nativeAddRevision(doc, current.paraId, current.offset, cursor.offset, {
            revisionId: adjacentRevisionId(current.paraId, current.offset, 'insert', 'before') ?? allocateNativeRevisionId(doc),
            revisionKind: 'insert',
            author: 'User',
            date: new Date().toISOString(),
          });
          return marked ? [inserted, marked] : inserted;
        },
      );
    }
    if (command.kind === 'replace-range' && command.anchor.paraId === command.focus.paraId) {
      const range = normalizedSelectionRange(command.anchor.offset, command.focus.offset);
      const beforeSelection: WordEditorSelection = {
        kind: 'text-range',
        anchor: command.anchor,
        focus: command.focus,
      };
      if (trackRevisions) {
        return applyLocalParagraphTextEdit(
          { anchor: command.anchor, focus: command.focus },
          beforeSelection,
          { paraId: command.anchor.paraId, offset: range.to + command.text.length },
          () => {
            const revisionChanges: NativeWordChange[] = [];
            const deleted = nativeAddRevision(doc, command.anchor.paraId, range.from, range.to, {
              revisionId: allocateNativeRevisionId(doc), revisionKind: 'delete', author: 'User', date: new Date().toISOString(),
            });
            if (deleted) revisionChanges.push(deleted);
            if (command.text.length > 0) {
              const inserted = nativeInsertText(doc, command.anchor.paraId, range.to, command.text);
              if (inserted) revisionChanges.push(inserted);
              const marked = nativeAddRevision(doc, command.anchor.paraId, range.to, range.to + command.text.length, {
                revisionId: allocateNativeRevisionId(doc), revisionKind: 'insert', author: 'User', date: new Date().toISOString(),
              });
              if (marked) revisionChanges.push(marked);
            }
            return revisionChanges;
          },
        );
      }
      return applyLocalParagraphTextEdit(
        { anchor: command.anchor, focus: command.focus },
        beforeSelection,
        { paraId: command.anchor.paraId, offset: range.from + command.text.length },
        () => nativeReplaceRange(doc, command.anchor.paraId, range.from, range.to, command.text),
      );
    }
    if (command.kind === 'delete-range' && trackRevisions) {
      const selectedRange: Extract<WordEditorSelection, { kind: 'text-range' }> = {
        kind: 'text-range',
        anchor: command.anchor,
        focus: command.focus,
      };
      const ranges = selectedTextRanges(selectedRange);
      if (ranges.length === 0) return false;
      const beforeParagraphs = Array.from(new Map(
        ranges
          .map((range) => findNativeParagraph(doc, range.paraId))
          .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
          .map((paragraph) => [paragraph.paraId, cloneNativeParagraph(paragraph)] as const),
      ).values());
      const revisionId = allocateNativeRevisionId(doc);
      const deleteChanges = ranges
        .map((range) => nativeAddRevision(doc, range.paraId, range.from, range.to, {
          revisionId,
          revisionKind: 'delete',
          author: 'User',
          date: new Date().toISOString(),
        }))
        .filter((change): change is NativeWordChange => change !== null);
      if (deleteChanges.length === 0) return false;
      const accepted = recordFormatting(selectedRange, beforeParagraphs, deleteChanges, false);
      if (accepted) {
        const ordered = orderedSelectionRange(doc, selectedRange);
        const cursor = { paraId: ordered.anchor.paraId, offset: ordered.anchor.offset };
        restoreNativeCursor(host, cursor);
        setSelection({ kind: 'caret', ...cursor });
      }
      return accepted;
    }
    if (command.kind === 'delete-range' && command.anchor.paraId === command.focus.paraId) {
      const selectedRange: Extract<WordEditorSelection, { kind: 'text-range' }> = {
        kind: 'text-range',
        anchor: command.anchor,
        focus: command.focus,
      };
      if (applyLocalPageBreakRangeDelete(selectedRange)) return true;
      const range = normalizedSelectionRange(command.anchor.offset, command.focus.offset);
      if (trackRevisions) {
        return applyLocalParagraphTextEdit(
          { anchor: command.anchor, focus: command.focus },
          selectedRange,
          { paraId: command.anchor.paraId, offset: range.from },
          () => nativeAddRevision(doc, command.anchor.paraId, range.from, range.to, {
            revisionId: adjacentRevisionId(command.anchor.paraId, range.from, 'delete', 'at') ?? allocateNativeRevisionId(doc), revisionKind: 'delete', author: 'User', date: new Date().toISOString(),
          }),
        );
      }
      return applyLocalParagraphTextEdit(
        { anchor: command.anchor, focus: command.focus },
        selectedRange,
        { paraId: command.anchor.paraId, offset: range.from },
        () => nativeDeleteRange(doc, command.anchor.paraId, range.from, range.to),
      );
    }
    if (command.kind === 'split-paragraph' && current?.kind === 'caret') {
      return splitParagraphLocally({ paraId: current.paraId, offset: current.offset });
    }
    if (command.kind === 'insert-page-break' && current?.kind === 'caret') {
      return applyLocalPageBreakEdit(current, 'insert');
    }
    if (command.kind === 'insert-image' && current && (current.kind === 'caret' || current.kind === 'text-range')) {
      const range = current.kind === 'text-range'
        ? orderedSelectionRange(doc, current)
        : {
            anchor: { paraId: current.paraId, offset: current.offset },
            focus: { paraId: current.paraId, offset: current.offset },
          };
      return applyLocalImageReplacement(range, command.src, {
        ...(command.width ? { width: command.width } : {}),
        ...(command.height ? { height: command.height } : {}),
        select: true,
      });
    }
    if (
      command.kind === 'set-image-size' &&
      current?.kind === 'object' &&
      current.objectKind === 'image' &&
      current.objectId === command.objectId
    ) {
      const paragraph = findNativeParagraph(doc, current.paraId);
      if (!paragraph) return false;
      const beforeParagraph = cloneNativeParagraph(paragraph);
      const beforeImage = paragraph.content.find(
        (child) => child.kind === 'image' && child.drawingId === current.objectId,
      );
      const beforeWidth = beforeImage?.kind === 'image' ? beforeImage.width : undefined;
      const beforeHeight = beforeImage?.kind === 'image' ? beforeImage.height : undefined;
      const change = nativeSetImageSize(doc, current.paraId, current.objectId, {
        width: command.width,
        height: command.height,
      });
      if (trackRevisions && change) {
        const image = paragraph.content.find(
          (child) => child.kind === 'image' && child.drawingId === current.objectId,
        );
        if (image?.kind === 'image') {
          image.formatRevision = {
            kind: 'format',
            revisionId: allocateNativeRevisionId(doc),
            beforeWidth,
            beforeHeight,
            afterWidth: command.width,
            afterHeight: command.height,
            author: 'User',
            date: new Date().toISOString(),
          };
          if (change.kind === 'image') {
            change.revisionId = image.formatRevision.revisionId;
            change.revisionKind = 'format';
          }
        }
      }
      return recordFormatting(current, [beforeParagraph], change ? [change] : [], true);
    }
    if (
      command.kind === 'set-image-alignment' &&
      current?.kind === 'object' &&
      current.objectKind === 'image' &&
      current.objectId === command.objectId
    ) {
      return dispatch({ kind: 'set-paragraph-alignment', alignment: command.alignment });
    }
    if (
      current &&
      (command.kind === 'insert-table' || command.kind === 'delete-table')
    ) {
      return applyLocalTableLifecycleEdit(current, command);
    }
    if (
      current &&
      (command.kind === 'insert-table-row' ||
        command.kind === 'delete-table-row' ||
        command.kind === 'insert-table-column' ||
        command.kind === 'delete-table-column')
    ) {
      return applyLocalTableStructureEdit(current, command);
    }
    if (
      current?.kind === 'table-cell-range' &&
      (command.kind === 'merge-table-cells' || command.kind === 'split-table-cell')
    ) {
      return applyLocalTableCellStructureEdit(current, command);
    }
    if (
      current &&
      (command.kind === 'set-table-cell-background' ||
        command.kind === 'set-table-cell-vertical-alignment' ||
        command.kind === 'set-table-cell-borders')
    ) {
      return applyLocalTableCellFormat(current, command);
    }
    if (command.kind === 'clear-table-cells' && current?.kind === 'table-cell-range') {
      return applyLocalTableCellClear(current);
    }
    if (command.kind === 'delete-backward') {
      if (current?.kind === 'table-cell-range') return applyLocalTableCellClear(current);
      if (current?.kind === 'text-range') {
        if (applyLocalPageBreakRangeDelete(current)) return true;
        if (trackRevisions) return dispatch({ kind: 'delete-range', anchor: current.anchor, focus: current.focus });
        return applyLocalPlainTextReplacement(orderedSelectionRange(doc, current), '');
      }
      if (current?.kind === 'object' && current.objectKind === 'image') {
        const cursor = {
          paraId: current.paraId,
          offset: current.offset,
        };
        return deleteImageWithHistory(current.paraId, current.objectId, cursor);
      }
      if (current?.kind === 'caret') {
        if (trackRevisions && current.offset > 0) {
          const from = current.offset - 1;
          return applyLocalParagraphTextEdit(
            { anchor: current, focus: current },
            current,
            { paraId: current.paraId, offset: from },
            () => nativeAddRevision(doc, current.paraId, from, current.offset, {
              revisionId: adjacentRevisionId(current.paraId, from, 'delete', 'after') ?? allocateNativeRevisionId(doc), revisionKind: 'delete', author: 'User', date: new Date().toISOString(),
            }),
          );
        }
        if (applyLocalPageBreakEdit(current, 'delete-backward')) return true;
        if (deletePageBreakAcrossParagraphBoundary(current, 'backward')) return true;
        if (mergeParagraphBackwardLocally({ paraId: current.paraId, offset: current.offset })) return true;
        const paragraph = findNativeParagraph(doc, current.paraId);
        const beforeParagraph = paragraph ? cloneNativeParagraph(paragraph) : null;
        const change = nativeDeleteBackward(doc, current.paraId, current.offset);
        const cursor = {
          paraId: current.offset > 0 ? current.paraId : previousNativeParagraphId(doc, current.paraId) ?? current.paraId,
          offset: Math.max(0, current.offset - 1),
        };
        if (change?.kind === 'image') {
          recordImageDeletion(beforeParagraph, change, cursor);
        } else {
          record(change, cursor, change?.kind === 'block' ? {} : { localParaId: current.paraId });
        }
        return true;
      }
    }
    if (command.kind === 'delete-forward' && current?.kind === 'table-cell-range') {
      return applyLocalTableCellClear(current);
    }
    if (command.kind === 'delete-forward' && current?.kind === 'text-range') {
      if (applyLocalPageBreakRangeDelete(current)) return true;
      if (trackRevisions) return dispatch({ kind: 'delete-range', anchor: current.anchor, focus: current.focus });
      return applyLocalPlainTextReplacement(orderedSelectionRange(doc, current), '');
    }
    if (command.kind === 'delete-forward' && current?.kind === 'caret') {
      if (trackRevisions && current.offset < paragraphTextLength(current.paraId)) {
        return applyLocalParagraphTextEdit(
          { anchor: current, focus: current },
          current,
          current,
          () => nativeAddRevision(doc, current.paraId, current.offset, current.offset + 1, {
            revisionId: adjacentRevisionId(current.paraId, current.offset, 'delete', 'at') ?? allocateNativeRevisionId(doc), revisionKind: 'delete', author: 'User', date: new Date().toISOString(),
          }),
        );
      }
      if (applyLocalPageBreakEdit(current, 'delete-forward')) return true;
      if (deletePageBreakAcrossParagraphBoundary(current, 'forward')) return true;
      if (mergeParagraphForwardLocally({ paraId: current.paraId, offset: current.offset })) return true;
      const paragraph = findNativeParagraph(doc, current.paraId);
      const beforeParagraph = paragraph ? cloneNativeParagraph(paragraph) : null;
      const change = nativeDeleteForward(doc, current.paraId, current.offset);
      const cursor = {
        paraId: current.paraId,
        offset: current.offset,
      };
      if (change?.kind === 'image') {
        recordImageDeletion(beforeParagraph, change, cursor);
      } else {
        record(change, cursor, { localParaId: current.paraId });
      }
      return true;
    }
    if (command.kind === 'delete-object' && current?.kind === 'object') {
      const cursor = {
        paraId: current.paraId,
        offset: current.offset,
      };
      return deleteImageWithHistory(current.paraId, command.objectId, cursor);
    }
    if (
      (command.kind === 'set-hyperlink' || command.kind === 'remove-hyperlink') &&
      current?.kind === 'text-range'
    ) {
      const ranges = selectedTextRanges(current);
      const beforeParagraphs = ranges
        .map((range) => findNativeParagraph(doc, range.paraId))
        .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
        .map(cloneNativeParagraph);
      const hyperlinkChanges = ranges
        .map((range) => nativeSetHyperlink(
          doc,
          range.paraId,
          range.from,
          range.to,
          command.kind === 'set-hyperlink' ? command.hyperlink : null,
        ))
        .filter((change): change is NativeWordChange => change !== null);
      return recordFormatting(current, beforeParagraphs, hyperlinkChanges, false);
    }
    if (command.kind === 'add-comment' && current?.kind === 'text-range') {
      const ranges = selectedTextRanges(current);
      if (ranges.length === 0 || !command.body.trim()) return false;
      const beforeParagraphs = ranges
        .map((range) => findNativeParagraph(doc, range.paraId))
        .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
        .map(cloneNativeParagraph);
      const commentId = allocateNativeCommentId(doc);
      const date = command.date ?? new Date().toISOString();
      const commentChanges = ranges
        .map((range) => nativeAddComment(doc, range.paraId, range.from, range.to, {
          commentId,
          author: command.author ?? 'User',
          date,
          body: command.body.trim(),
        }))
        .filter((change): change is NativeWordChange => change !== null);
      const recorded = recordFormatting(current, beforeParagraphs, commentChanges, false);
      if (recorded && commentChanges.length > 1) render();
      return recorded;
    }
    if (command.kind === 'edit-comment' || command.kind === 'delete-comment') {
      if (command.kind === 'edit-comment' && !command.body.trim()) return false;
      const beforeParagraphs = allNativeParagraphs(doc)
        .filter((paragraph) => paragraph.content.some((child) => child.kind === 'text' && child.marks?.some(
          (mark) => mark.kind === 'comment' && mark.commentId === command.commentId,
        )))
        .map(cloneNativeParagraph);
      if (beforeParagraphs.length === 0) return false;
      const commentChanges = command.kind === 'edit-comment'
        ? nativeUpdateComment(doc, command.commentId, command.body.trim())
        : nativeDeleteComment(doc, command.commentId);
      if (commentChanges.length === 0) return false;
      const historySelection = current ?? {
        kind: 'caret' as const,
        paraId: beforeParagraphs[0]!.paraId,
        offset: 0,
      };
      return recordFormatting(historySelection, beforeParagraphs, commentChanges, false);
    }
    if (command.kind === 'reply-comment') {
      const body = command.body.trim();
      if (!body) return false;
      let target: { paragraph: NativeWordParagraphNode; from: number; to: number } | null = null;
      for (const paragraph of allNativeParagraphs(doc)) {
        let position = 0;
        for (const child of paragraph.content) {
          const length = child.kind === 'text' ? child.text.length : 1;
          if (child.kind === 'text' && child.marks?.some((mark) => mark.kind === 'comment' && mark.commentId === command.commentId)) {
            target = { paragraph, from: position, to: position + length };
            break;
          }
          position += length;
        }
        if (target) break;
      }
      if (!target) return false;
      const beforeParagraph = cloneNativeParagraph(target.paragraph);
      const replyId = allocateNativeCommentId(doc);
      const change = nativeAddComment(doc, target.paragraph.paraId, target.from, target.to, {
        commentId: replyId,
        parentCommentId: command.commentId,
        preserveExistingComments: true,
        author: 'User',
        date: new Date().toISOString(),
        body,
      });
      if (!change) return false;
      render();
      return recordFormatting(current ?? { kind: 'caret', paraId: target.paragraph.paraId, offset: target.from }, [beforeParagraph], [change], false);
    }
    if (command.kind === 'add-revision' && current?.kind === 'text-range') {
      const ranges = selectedTextRanges(current);
      if (ranges.length === 0) return false;
      const beforeParagraphs = ranges
        .map((range) => findNativeParagraph(doc, range.paraId))
        .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
        .map(cloneNativeParagraph);
      const revisionId = allocateNativeRevisionId(doc);
      const date = command.date ?? new Date().toISOString();
      const revisionChanges = ranges
        .map((range) => nativeAddRevision(doc, range.paraId, range.from, range.to, {
          revisionId,
          revisionKind: command.revisionKind,
          author: command.author ?? 'User',
          date,
        }))
        .filter((change): change is NativeWordChange => change !== null);
      return recordFormatting(current, beforeParagraphs, revisionChanges, false);
    }
    if (
      command.kind === 'toggle-bold' ||
      command.kind === 'toggle-italic' ||
      command.kind === 'toggle-underline' ||
      command.kind === 'toggle-strike' ||
      command.kind === 'set-font-family' ||
      command.kind === 'set-font-size' ||
      command.kind === 'set-text-color' ||
      command.kind === 'set-highlight-color' ||
      command.kind === 'clear-text-format'
    ) {
      const ranges = selectedTextRanges(current);
      if (current?.kind !== 'text-range' || ranges.length === 0) return false;
      const beforeParagraphs = ranges
        .map((range) => findNativeParagraph(doc, range.paraId))
        .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
        .map(cloneNativeParagraph);
      const commandState = getCommandState(current);
      let format: Partial<EditorRunFormat> | null = null;
      if (command.kind === 'toggle-bold') format = { bold: commandState.bold !== true };
      if (command.kind === 'toggle-italic') format = { italic: commandState.italic !== true };
      if (command.kind === 'toggle-underline') format = { underline: commandState.underline !== true };
      if (command.kind === 'toggle-strike') format = { strike: commandState.strike !== true };
      if (command.kind === 'set-font-family' && command.fontFamily.trim()) {
        format = { fontFamily: command.fontFamily.trim() };
      }
      if (command.kind === 'set-font-size') format = { fontSize: command.fontSize };
      if (command.kind === 'set-text-color') format = { color: command.color.replace(/^#/, '').toUpperCase() };
      if (command.kind === 'set-highlight-color') {
        format = { highlight: command.color ? command.color.replace(/^#/, '').toUpperCase() : 'none' };
      }
      if (command.kind !== 'clear-text-format' && !format) return false;
      const formattingChanges = ranges
        .map((range) =>
          command.kind === 'clear-text-format'
            ? nativeClearRunFormat(doc, range.paraId, range.from, range.to)
            : nativeSetRunFormat(doc, range.paraId, range.from, range.to, format!),
        )
        .filter((change): change is NativeWordChange => change !== null);
      if (trackRevisions) {
        const revisionId = allocateNativeRevisionId(doc);
        const formatRevisionChanges = ranges
          .map((range) => nativeAddFormatRevision(
            doc,
            range.paraId,
            range.from,
            range.to,
            beforeParagraphs.find((paragraph) => paragraph.paraId === range.paraId)!,
            { revisionId, author: 'User', date: new Date().toISOString() },
          ))
          .filter((change): change is NativeWordChange => change !== null);
        formattingChanges.push(...formatRevisionChanges);
      }
      const reflow =
        command.kind === 'toggle-bold' ||
        command.kind === 'toggle-italic' ||
        command.kind === 'set-font-family' ||
        command.kind === 'set-font-size' ||
        command.kind === 'clear-text-format';
      return recordFormatting(current, beforeParagraphs, formattingChanges, reflow);
    }
    if (
      command.kind === 'set-paragraph-alignment' ||
      command.kind === 'set-paragraph-indent' ||
      command.kind === 'adjust-paragraph-left-indent' ||
      command.kind === 'set-paragraph-spacing' ||
      command.kind === 'set-paragraph-line-spacing'
    ) {
      const paraIds = Array.from(new Set(selectedParagraphIds(current)));
      if (!current || paraIds.length === 0) return false;
      const beforeParagraphs = paraIds
        .map((paraId) => findNativeParagraph(doc, paraId))
        .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null)
        .map(cloneNativeParagraph);
      const formattingChanges = paraIds
        .map((paraId) => {
          const paragraph = findNativeParagraph(doc, paraId);
          if (!paragraph) return null;
          if (command.kind === 'set-paragraph-alignment') {
            return nativeSetParagraphFormat(doc, paraId, { alignment: command.alignment });
          }
          if (command.kind === 'set-paragraph-indent') {
            return nativeSetParagraphFormat(doc, paraId, { indent: sanitizeIndent(command.indent) });
          }
          if (command.kind === 'adjust-paragraph-left-indent') {
            const left = Math.max(0, (paragraph.format?.indent?.left ?? 0) + command.delta);
            return nativeSetParagraphFormat(doc, paraId, { indent: { left } });
          }
          if (command.kind === 'set-paragraph-spacing') {
            return nativeSetParagraphFormat(doc, paraId, { spacing: sanitizeSpacing(command.spacing) });
          }
          return nativeSetParagraphFormat(doc, paraId, {
            spacing: {
              line: Math.max(0, command.line),
              lineRule: command.lineRule,
            },
          });
        })
        .filter((change): change is NativeWordChange => change !== null);
      return recordFormatting(current, beforeParagraphs, formattingChanges, command.kind !== 'set-paragraph-alignment');
    }
    if (command.kind === 'toggle-list' || command.kind === 'remove-list' || command.kind === 'adjust-list-level') {
      const paraIds = Array.from(new Set(selectedParagraphIds(current)));
      if (!current || paraIds.length === 0) return false;
      const paragraphs = paraIds
        .map((paraId) => findNativeParagraph(doc, paraId))
        .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null);
      const beforeParagraphs = paragraphs.map(cloneNativeParagraph);
      const allMatchToggleType =
        command.kind === 'toggle-list' &&
        paragraphs.every(
          (paragraph) =>
            paragraph.format?.hasNumbering && nativeListType(paragraph.format.numbering) === command.listType,
        );
      const newList =
        command.kind === 'toggle-list' && !allMatchToggleType
          ? adjacentNativeListNumbering(doc, paraIds, command.listType) ??
            createNativeListNumbering(doc, command.listType)
          : null;
      const listChanges = paragraphs
        .map((paragraph) => {
          if (command.kind === 'remove-list' || allMatchToggleType) {
            return paragraph.format?.hasNumbering ? nativeSetParagraphList(doc, paragraph.paraId, null) : null;
          }
          if (command.kind === 'toggle-list' && newList) {
            const level = paragraph.format?.numbering?.ilvl ?? 0;
            return nativeSetParagraphList(doc, paragraph.paraId, nativeNumberingAtLevel(newList, level));
          }
          if (command.kind === 'adjust-list-level') {
            const numbering = paragraph.format?.hasNumbering ? paragraph.format.numbering : undefined;
            if (!numbering) return null;
            return nativeSetParagraphList(
              doc,
              paragraph.paraId,
              nativeNumberingAtLevel(numbering, numbering.ilvl + command.delta),
            );
          }
          return null;
        })
        .filter((change): change is NativeWordChange => change !== null);
      refreshNativeNumberingLabels(doc);
      return recordFormatting(current, beforeParagraphs, listChanges, true);
    }
    return false;
  };

  interface NativeTableClipboardEntry extends NativeTableClipboardCell {
    text: string;
  }

  interface NativeTableClipboardPayload {
    version: 1;
    cells: Array<Array<NativeTableClipboardEntry | null>>;
  }

  const tableClipboardPayload = (
    current: Extract<WordEditorSelection, { kind: 'table-cell-range' }>,
  ): NativeTableClipboardPayload | null => {
    const table = findNativeTableLocation(doc, current.tableId)?.table;
    if (!table) return null;
    const startRow = Math.min(current.anchor.rowIndex, current.focus.rowIndex);
    const endRow = Math.max(current.anchor.rowIndex, current.focus.rowIndex);
    const startColumn = Math.min(current.anchor.columnIndex, current.focus.columnIndex);
    const endColumn = Math.max(current.anchor.columnIndex, current.focus.columnIndex);
    const cells: Array<Array<NativeTableClipboardEntry | null>> = Array.from(
      { length: endRow - startRow + 1 },
      () => Array.from({ length: endColumn - startColumn + 1 }, () => null),
    );
    for (const { cell, rowIndex, columnIndex } of nativeTableCellsInRange(table, current)) {
      const row = cells[rowIndex - startRow];
      if (row && columnIndex >= startColumn && columnIndex <= endColumn) {
        row[columnIndex - startColumn] = {
          text: cell.paragraphs.map((paragraph) => nativeParagraphText(paragraph)).join('\n'),
          blocks: cloneNativeWordBlocks(cell.blocks),
          format: tableClipboardCellFormat(cell),
        };
      }
    }
    return { version: 1, cells };
  };

  const tableClipboardCellFormat = (cell: NativeWordTableCellNode): NativeTableClipboardCell['format'] =>
    JSON.parse(JSON.stringify({
      ...(cell.margin ? { margin: cell.margin } : {}),
      ...(cell.borders ? { borders: cell.borders } : {}),
      ...(cell.shading ? { shading: cell.shading } : {}),
      ...(cell.verticalAlignment ? { verticalAlignment: cell.verticalAlignment } : {}),
      ...(cell.noWrap !== undefined ? { noWrap: cell.noWrap } : {}),
    })) as NativeTableClipboardCell['format'];

  const tableClipboardText = (payload: NativeTableClipboardPayload): string =>
    payload.cells
      .map((row) => row.map((value) => (value?.text ?? '').replace(/[\t\r\n]+/g, ' ')).join('\t'))
      .join('\n');

  const parseRichTableClipboard = (event: ClipboardEvent): NativeTableClipboardPayload | null => {
    const internal = event.clipboardData?.getData(NATIVE_TABLE_CLIPBOARD_TYPE) ?? '';
    if (!internal) return null;
    try {
      const parsed = JSON.parse(internal) as Partial<NativeTableClipboardPayload>;
      if (
        parsed.version === 1 &&
        Array.isArray(parsed.cells) &&
        parsed.cells.every((row) => Array.isArray(row) && row.every((value) =>
          value === null || (
            typeof value === 'object' &&
            typeof value.text === 'string' &&
            Array.isArray(value.blocks) &&
            typeof value.format === 'object'
          )))
      ) {
        return parsed as NativeTableClipboardPayload;
      }
    } catch {
      return null;
    }
    return null;
  };

  const parseTableClipboardText = (event: ClipboardEvent): string[][] => {
    const text = event.clipboardData?.getData('text/plain') ?? '';
    if (!text) return [];
    const rows = text.replace(/\r\n?/g, '\n').split('\n');
    if (rows[rows.length - 1] === '') rows.pop();
    return rows.map((row) => row.split('\t'));
  };

  const copy = (event: ClipboardEvent) => {
    if (selectedImage && event.clipboardData) {
      const imageSelection = selectedImage;
      const image = findNativeParagraph(doc, imageSelection.paraId)?.content.find(
        (child) => child.kind === 'image' && child.drawingId === imageSelection.drawingId,
      );
      if (!image || image.kind !== 'image') return;
      event.preventDefault();
      event.clipboardData.setData(NATIVE_IMAGE_CLIPBOARD_TYPE, storeNativeImageClipboardSource(image.src));
      event.clipboardData.setData('text/html', `<img src="${escapeHtmlAttribute(image.src)}">`);
      void copySelectedImageToSystemClipboard(imageSelection).catch((error: unknown) => {
        console.warn('Unable to copy the selected image', error);
      });
      return;
    }
    const current = selection?.kind === 'object' || selection?.kind === 'table-cell-range'
      ? selection
      : currentNativeSelection(host) ?? selection;
    if (current?.kind === 'table-cell-range' && event.clipboardData) {
      const payload = tableClipboardPayload(current);
      if (!payload) return;
      event.preventDefault();
      event.clipboardData.setData(NATIVE_TABLE_CLIPBOARD_TYPE, JSON.stringify(payload));
      event.clipboardData.setData('text/plain', tableClipboardText(payload));
      return;
    }
    if (current?.kind === 'object' && current.objectKind === 'image' && event.clipboardData) {
      const image = findNativeParagraph(doc, current.paraId)?.content.find(
        (child) => child.kind === 'image' && child.drawingId === current.objectId,
      );
      if (!image || image.kind !== 'image') return;
      event.preventDefault();
      event.clipboardData.setData(NATIVE_IMAGE_CLIPBOARD_TYPE, storeNativeImageClipboardSource(image.src));
      event.clipboardData.setData('text/html', `<img src="${escapeHtmlAttribute(image.src)}">`);
      void copySelectedImageToSystemClipboard({ paraId: current.paraId, drawingId: current.objectId }).catch(
        (error: unknown) => {
          console.warn('Unable to copy the selected image', error);
        },
      );
      return;
    }
    if (current?.kind === 'text-range' && event.clipboardData) {
      const range = orderedSelectionRange(doc, current);
      const blocks = nativeCopyBlockRange(doc, range.anchor, range.focus);
      if (blocks) {
        event.preventDefault();
        event.clipboardData.setData(NATIVE_BLOCK_CLIPBOARD_TYPE, serializeNativeBlockClipboard(blocks));
        event.clipboardData.setData('text/plain', nativeBlockClipboardText(blocks));
        return;
      }
    }
    const text = clipboardTextForSelection(current);
    if (text === null || !event.clipboardData) return;
    event.preventDefault();
    event.clipboardData.setData('text/plain', text);
    const richParagraphs = clipboardParagraphsForSelection(current);
    if (richParagraphs.length > 0) {
      event.clipboardData.setData('text/html', serializeNativeClipboardHtml(richParagraphs));
    }
  };

  const copySelectedImageToSystemClipboard = async (imageSelection: NativeSelectedImage): Promise<void> => {
    const image = findNativeParagraph(doc, imageSelection.paraId)?.content.find(
      (child) => child.kind === 'image' && child.drawingId === imageSelection.drawingId,
    );
    if (!image || image.kind !== 'image') return;
    const imageBlob = await clipboardPngBlob(image.src);
    await navigator.clipboard.write([new ClipboardItem({ 'image/png': imageBlob })]);
  };

  const cut = (event: ClipboardEvent) => {
    if (selectedImage && event.clipboardData) {
      const imageSelection = selectedImage;
      const image = findNativeParagraph(doc, imageSelection.paraId)?.content.find(
        (child) => child.kind === 'image' && child.drawingId === imageSelection.drawingId,
      );
      if (!image || image.kind !== 'image') return;
      event.preventDefault();
      event.clipboardData.setData(NATIVE_IMAGE_CLIPBOARD_TYPE, storeNativeImageClipboardSource(image.src));
      event.clipboardData.setData('text/html', `<img src="${escapeHtmlAttribute(image.src)}">`);
      const paragraphEl = host.querySelector<HTMLElement>(
        `coco-word-native-p[data-para-id="${imageSelection.paraId}"]`,
      );
      const imageEl = host.querySelector<HTMLImageElement>(
        `img[data-drawing-id="${imageSelection.drawingId}"]`,
      );
      if (!paragraphEl || !imageEl) return;
      deleteImageWithHistory(imageSelection.paraId, imageSelection.drawingId, {
        paraId: imageSelection.paraId,
        offset: imageOffset(paragraphEl, imageEl),
      });
      return;
    }
    const current = selection?.kind === 'object' || selection?.kind === 'table-cell-range'
      ? selection
      : currentNativeSelection(host) ?? selection;
    if (current?.kind === 'table-cell-range' && event.clipboardData) {
      const payload = tableClipboardPayload(current);
      if (!payload) return;
      event.preventDefault();
      event.clipboardData.setData(NATIVE_TABLE_CLIPBOARD_TYPE, JSON.stringify(payload));
      event.clipboardData.setData('text/plain', tableClipboardText(payload));
      applyLocalTableCellClear(current);
      return;
    }
    if (current?.kind === 'object' && current.objectKind === 'image' && event.clipboardData) {
      const image = findNativeParagraph(doc, current.paraId)?.content.find(
        (child) => child.kind === 'image' && child.drawingId === current.objectId,
      );
      if (!image || image.kind !== 'image') return;
      event.preventDefault();
      event.clipboardData.setData(NATIVE_IMAGE_CLIPBOARD_TYPE, storeNativeImageClipboardSource(image.src));
      event.clipboardData.setData('text/html', `<img src="${escapeHtmlAttribute(image.src)}">`);
      deleteImageWithHistory(current.paraId, current.objectId, { paraId: current.paraId, offset: current.offset });
      return;
    }
    if (current?.kind === 'text-range' && event.clipboardData) {
      const range = orderedSelectionRange(doc, current);
      const blocks = nativeCopyBlockRange(doc, range.anchor, range.focus);
      if (blocks) {
        event.preventDefault();
        event.clipboardData.setData(NATIVE_BLOCK_CLIPBOARD_TYPE, serializeNativeBlockClipboard(blocks));
        event.clipboardData.setData('text/plain', nativeBlockClipboardText(blocks));
        applyLocalBlockRangeReplacement(range, '');
        return;
      }
    }
    const text = clipboardTextForSelection(current);
    if (text === null || current?.kind !== 'text-range' || !event.clipboardData) return;
    event.preventDefault();
    event.clipboardData.setData('text/plain', text);
    const richParagraphs = clipboardParagraphsForSelection(current);
    if (richParagraphs.length > 0) {
      event.clipboardData.setData('text/html', serializeNativeClipboardHtml(richParagraphs));
    }
    applyLocalPlainTextReplacement(orderedSelectionRange(doc, current), '');
  };

  const paste = (event: ClipboardEvent) => {
    const tableSelection = selection?.kind === 'table-cell-range' ? selection : null;
    if (tableSelection) {
      const rich = parseRichTableClipboard(event);
      const text = rich ? [] : parseTableClipboardText(event);
      if (
        (rich && applyLocalTableCellPaste(tableSelection, { content: rich.cells })) ||
        (text.length > 0 && applyLocalTableCellPaste(tableSelection, { text }))
      ) {
        event.preventDefault();
      }
      return;
    }
    const nativeBlocks = parseNativeBlockClipboard(
      event.clipboardData?.getData(NATIVE_BLOCK_CLIPBOARD_TYPE) ?? '',
    );
    if (nativeBlocks) {
      const current = currentNativeSelection(host) ?? selection ?? nativeSelectionFromCursor(host);
      if (!current || current.kind === 'object' || current.kind === 'table-cell-range') return;
      const range = current.kind === 'text-range'
        ? orderedSelectionRange(doc, current)
        : { anchor: { paraId: current.paraId, offset: current.offset }, focus: { paraId: current.paraId, offset: current.offset } };
      if (applyLocalClipboardBlocks(range, nativeBlocks)) {
        event.preventDefault();
        return;
      }
    }
    const nativeImageSrc = resolveNativeImageClipboardSource(
      event.clipboardData?.getData(NATIVE_IMAGE_CLIPBOARD_TYPE) ?? '',
    );
    if (nativeImageSrc) {
      const current = currentNativeSelection(host) ?? selection ?? nativeSelectionFromCursor(host);
      if (!current || current.kind === 'object' || current.kind === 'table-cell-range') return;
      const range = current.kind === 'text-range'
        ? orderedSelectionRange(doc, current)
        : { anchor: { paraId: current.paraId, offset: current.offset }, focus: { paraId: current.paraId, offset: current.offset } };
      event.preventDefault();
      applyLocalImageReplacement(range, nativeImageSrc);
      return;
    }
    const image = Array.from(event.clipboardData?.files ?? []).find((file) => file.type.startsWith('image/'));
    if (image) {
      const current = currentNativeSelection(host) ?? selection ?? nativeSelectionFromCursor(host);
      if (!current || current.kind === 'object' || current.kind === 'table-cell-range') return;
      const range = current.kind === 'text-range'
        ? orderedSelectionRange(doc, current)
        : { anchor: { paraId: current.paraId, offset: current.offset }, focus: { paraId: current.paraId, offset: current.offset } };
      event.preventDefault();
      void blobToDataUrl(image).then((src) => applyLocalImageReplacement(range, src));
      return;
    }
    const html = event.clipboardData?.getData('text/html') ?? '';
    const richParagraphs = html ? parseNativeClipboardHtml(html) : [];
    const text = richParagraphs.length > 0
      ? richParagraphs.map((paragraph) => paragraph.text).join('\n')
      : event.clipboardData?.getData('text/plain');
    if (text === undefined) return;
    const current = currentNativeSelection(host) ?? selection ?? nativeSelectionFromCursor(host);
    if (!current || current.kind === 'object' || current.kind === 'table-cell-range') return;
    const range =
      current.kind === 'text-range'
        ? orderedSelectionRange(doc, current)
        : { anchor: { paraId: current.paraId, offset: current.offset }, focus: { paraId: current.paraId, offset: current.offset } };
    event.preventDefault();
    applyLocalPlainTextReplacement(range, text, richParagraphs.length > 0 ? richParagraphs : undefined);
  };

  const deleteContextSelection = (): boolean => {
    const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
    if (current?.kind === 'object' && current.objectKind === 'image') {
      return deleteImageWithHistory(current.paraId, current.objectId, { paraId: current.paraId, offset: current.offset });
    }
    if (current?.kind === 'text-range') {
      return applyLocalPlainTextReplacement(orderedSelectionRange(doc, current), '');
    }
    if (current?.kind === 'table-cell-range') return applyLocalTableCellClear(current);
    return false;
  };

  const pasteContextText = async (): Promise<void> => {
    if (typeof navigator.clipboard?.readText !== 'function') return;
    try {
      const text = await navigator.clipboard.readText();
      const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
      if (!current || current.kind === 'object') return;
      if (current.kind === 'table-cell-range') {
        const rows = text.replace(/\r\n?/g, '\n').split('\n').map((row) => row.split('\t'));
        applyLocalTableCellPaste(current, { text: rows });
        return;
      }
      const range = current.kind === 'text-range'
        ? orderedSelectionRange(doc, current)
        : { anchor: { paraId: current.paraId, offset: current.offset }, focus: { paraId: current.paraId, offset: current.offset } };
      applyLocalPlainTextReplacement(range, text);
    } catch (error) {
      console.warn('Unable to paste from the context menu', error);
    }
  };

  const openCommentDialog = (mode: 'add' | 'edit' | 'reply' = 'add', commentId?: number, initialBody = ''): void => {
    const overlay = document.createElement('div');
    overlay.className = 'coco-word-dialog-backdrop';
    const dialog = document.createElement('section');
    dialog.className = 'coco-word-dialog';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'coco-word-comment-dialog-title');
    const title = document.createElement('h2');
    title.id = 'coco-word-comment-dialog-title';
    title.textContent = mode === 'edit' ? '编辑批注' : mode === 'reply' ? '回复批注' : '添加批注';
    const textarea = document.createElement('textarea');
    textarea.className = 'coco-word-dialog-input';
    textarea.placeholder = '批注内容';
    textarea.rows = 4;
    textarea.value = initialBody;
    const actions = document.createElement('div');
    actions.className = 'coco-word-dialog-actions';
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'coco-word-dialog-button is-secondary';
    cancel.textContent = '取消';
    const confirm = document.createElement('button');
    confirm.type = 'button';
    confirm.className = 'coco-word-dialog-button is-primary';
    confirm.textContent = '确定';
    const close = () => overlay.remove();
    cancel.addEventListener('click', close);
    overlay.addEventListener('click', (event) => {
      if (event.target === overlay) close();
    });
    dialog.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') close();
      if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        confirm.click();
      }
    });
    confirm.addEventListener('click', () => {
      const body = textarea.value.trim();
      if (!body) {
        textarea.focus();
        return;
      }
      if (mode === 'edit' && commentId !== undefined) dispatch({ kind: 'edit-comment', commentId, body });
      else if (mode === 'reply' && commentId !== undefined) dispatch({ kind: 'reply-comment', commentId, body });
      else dispatch({ kind: 'add-comment', body });
      close();
    });
    actions.append(cancel, confirm);
    dialog.append(title, textarea, actions);
    overlay.appendChild(dialog);
    (shadowRoot ?? document.body).appendChild(overlay);
    textarea.focus();
  };

  const contextMenuItems = (): NativeContextMenuItem[] => {
    const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
    const state = getCommandState(current);
    const hasSelection = current?.kind === 'text-range' || current?.kind === 'object' || current?.kind === 'table-cell-range';
    const items: NativeContextMenuItem[] = [
      {
        label: '剪切',
        disabled: !hasSelection,
        action: () => document.execCommand('cut'),
      },
      { label: '复制', disabled: !hasSelection, action: () => document.execCommand('copy') },
      {
        label: '粘贴为纯文本',
        disabled: current?.kind === 'object' || typeof navigator.clipboard?.readText !== 'function',
        action: pasteContextText,
      },
      { label: '删除', disabled: !hasSelection, action: () => deleteContextSelection() },
    ];
    if (current?.kind === 'text-range') {
      items.push({
        label: '添加批注',
        separatorBefore: true,
        action: openCommentDialog,
      });
      items.push({
        label: '标记为插入修订',
        action: () => dispatch({ kind: 'add-revision', revisionKind: 'insert' }),
      });
      items.push({
        label: '标记为删除修订',
        action: () => dispatch({ kind: 'add-revision', revisionKind: 'delete' }),
      });
      items.push({
        label: state.bold === true ? '取消加粗' : '加粗',
        separatorBefore: true,
        disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'toggle-bold' }),
        action: () => dispatch({ kind: 'toggle-bold' }),
      });
      items.push({ label: '清除文本格式', action: () => dispatch({ kind: 'clear-text-format' }) });
      if (state.hyperlink !== null) {
        items.push({
          label: '移除超链接',
          disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'remove-hyperlink' }),
          action: () => dispatch({ kind: 'remove-hyperlink' }),
        });
      }
    }
    if (current?.kind === 'caret' && canDispatchNativeWordCommand(doc, current, {
      kind: 'insert-table',
      rows: 2,
      columns: 2,
    })) {
      items.push({
        label: '插入 2 × 2 表格',
        separatorBefore: true,
        action: () => dispatch({ kind: 'insert-table', rows: 2, columns: 2 }),
      });
    }
    if (current?.kind === 'object' && current.objectKind === 'image') {
      items.push({
        label: '图片左对齐',
        separatorBefore: true,
        action: () => dispatch({ kind: 'set-image-alignment', objectId: current.objectId, alignment: 'left' }),
      });
      items.push({
        label: '图片居中',
        action: () => dispatch({ kind: 'set-image-alignment', objectId: current.objectId, alignment: 'center' }),
      });
      items.push({
        label: '图片右对齐',
        action: () => dispatch({ kind: 'set-image-alignment', objectId: current.objectId, alignment: 'right' }),
      });
    }
    if (state.table) {
      const range = state.table.range ?? {
        startRow: state.table.rowIndex,
        endRow: state.table.rowIndex,
        startColumn: state.table.columnIndex,
        endColumn: state.table.columnIndex,
        cellCount: 1,
      };
      items.push({
        label: '选择整行',
        separatorBefore: true,
        action: () => setSelection({
          kind: 'table-cell-range',
          tableId: state.table!.tableId,
          anchor: { rowIndex: range.startRow, columnIndex: 0 },
          focus: { rowIndex: range.endRow, columnIndex: Math.max(0, state.table!.columnCount - 1) },
        }),
      });
      items.push({
        label: '选择整列',
        action: () => setSelection({
          kind: 'table-cell-range',
          tableId: state.table!.tableId,
          anchor: { rowIndex: 0, columnIndex: range.startColumn },
          focus: { rowIndex: Math.max(0, state.table!.rowCount - 1), columnIndex: range.endColumn },
        }),
      });
      items.push({
        label: '在上方插入行',
        disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'insert-table-row', position: 'before' }),
        action: () => dispatch({ kind: 'insert-table-row', position: 'before' }),
      });
      items.push({
        label: '在下方插入行',
        disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'insert-table-row', position: 'after' }),
        action: () => dispatch({ kind: 'insert-table-row', position: 'after' }),
      });
      items.push({
        label: '删除当前行',
        disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'delete-table-row' }),
        action: () => dispatch({ kind: 'delete-table-row' }),
      });
      items.push({
        label: '在左侧插入列',
        disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'insert-table-column', position: 'before' }),
        action: () => dispatch({ kind: 'insert-table-column', position: 'before' }),
      });
      items.push({
        label: '在右侧插入列',
        disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'insert-table-column', position: 'after' }),
        action: () => dispatch({ kind: 'insert-table-column', position: 'after' }),
      });
      items.push({
        label: '删除当前列',
        disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'delete-table-column' }),
        action: () => dispatch({ kind: 'delete-table-column' }),
      });
      items.push({
        label: '删除表格',
        separatorBefore: true,
        disabled: !canDispatchNativeWordCommand(doc, current, { kind: 'delete-table' }),
        action: () => dispatch({ kind: 'delete-table' }),
      });
      for (const [label, alignment] of [
        ['单元格顶部对齐', 'top'],
        ['单元格垂直居中', 'center'],
        ['单元格底部对齐', 'bottom'],
      ] as const) {
        items.push({
          label,
          action: () => dispatch({ kind: 'set-table-cell-vertical-alignment', alignment }),
        });
      }
    }
    items.push(
      { label: '撤销', separatorBefore: true, disabled: !browserUndoStack.length, action: () => undo() },
      { label: '重做', disabled: !browserRedoStack.length, action: () => redo() },
    );
    return items;
  };

  const selectContextTarget = (event: MouseEvent): WordEditorSelection | null => {
    const target = event.target;
    if (target instanceof HTMLImageElement) {
      const paragraph = parentNativeParagraph(target);
      const paraId = paragraph?.dataset.paraId;
      const drawingId = target.dataset.drawingId;
      if (paragraph && paraId && drawingId) {
        const next: WordEditorSelection = {
          kind: 'object',
          paraId,
          objectKind: 'image',
          objectId: drawingId,
          offset: imageOffset(paragraph, target),
        };
        selectedImage = { paraId, drawingId };
        showImageResizeHandle();
        setSelection(next);
        return next;
      }
    }
    const tableCell = tableCellTarget(target);
    if (tableCell) {
      if (selection?.kind === 'table-cell-range' && selection.tableId === tableCell.tableId) {
        const startRow = Math.min(selection.anchor.rowIndex, selection.focus.rowIndex);
        const endRow = Math.max(selection.anchor.rowIndex, selection.focus.rowIndex);
        const startColumn = Math.min(selection.anchor.columnIndex, selection.focus.columnIndex);
        const endColumn = Math.max(selection.anchor.columnIndex, selection.focus.columnIndex);
        if (
          tableCell.point.rowIndex >= startRow && tableCell.point.rowIndex <= endRow &&
          tableCell.point.columnIndex >= startColumn && tableCell.point.columnIndex <= endColumn
        ) {
          return selection;
        }
      }
      const next: WordEditorSelection = {
        kind: 'table-cell-range',
        tableId: tableCell.tableId,
        anchor: tableCell.point,
        focus: tableCell.point,
      };
      nativeSelection(host)?.removeAllRanges();
      setSelection(next);
      return next;
    }
    const point = nativeCursorFromPoint(host, event.clientX, event.clientY) ?? currentNativeCursor(host);
    if (!point) return selection;
    if (selection?.kind === 'text-range') {
      const ordered = orderedSelectionRange(doc, selection);
      const start = findNativeParagraphLocation(doc, ordered.anchor.paraId);
      const end = findNativeParagraphLocation(doc, ordered.focus.paraId);
      const target = findNativeParagraphLocation(doc, point.paraId);
      if (start && end && target && start.container === end.container && start.container === target.container) {
        const afterStart = target.index > start.index || (target.index === start.index && point.offset >= ordered.anchor.offset);
        const beforeEnd = target.index < end.index || (target.index === end.index && point.offset <= ordered.focus.offset);
        if (afterStart && beforeEnd) return selection;
      }
    }
    selectedImage = null;
    clearNativeImageSelection(host);
    restoreNativeCursor(host, point);
    const next: WordEditorSelection = { kind: 'caret', ...point };
    setSelection(next);
    return next;
  };

  const contextMenu = (event: MouseEvent) => {
    if (options.contextMenu === false) return;
    event.preventDefault();
    event.stopPropagation();
    activate();
    closeContextMenu();
    selectContextTarget(event);
    const menu = document.createElement('div');
    menu.className = 'coco-word-context-menu';
    menu.setAttribute('role', 'menu');
    menu.contentEditable = 'false';
    for (const item of contextMenuItems()) {
      if (item.separatorBefore) {
        const separator = document.createElement('div');
        separator.className = 'coco-word-context-menu-separator';
        separator.setAttribute('role', 'separator');
        menu.appendChild(separator);
      }
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'coco-word-context-menu-item';
      button.textContent = item.label;
      button.disabled = item.disabled ?? false;
      button.setAttribute('role', 'menuitem');
      button.addEventListener('mousedown', (pointerEvent) => pointerEvent.preventDefault());
      button.addEventListener('click', () => {
        const result = item.action();
        closeContextMenu();
        if (result instanceof Promise) {
          void result.catch((error: unknown) => console.warn('Context menu action failed', error));
        }
      });
      menu.appendChild(button);
    }
    (shadowRoot ?? document.body).appendChild(menu);
    contextMenuElement = menu;
    const rect = menu.getBoundingClientRect();
    menu.style.left = `${Math.max(4, Math.min(event.clientX, window.innerWidth - rect.width - 4))}px`;
    menu.style.top = `${Math.max(4, Math.min(event.clientY, window.innerHeight - rect.height - 4))}px`;
  };

  const closeContextMenuFromPointer = (event: PointerEvent) => {
    if (contextMenuElement && !contextMenuElement.contains(event.target as Node)) closeContextMenu();
  };

  const closeContextMenuFromKey = (event: KeyboardEvent) => {
    if (event.key === 'Escape') closeContextMenu();
  };

  const beforeInput = (event: InputEvent) => {
    if (event.inputType === suppressedBrowserHistoryInput) {
      event.preventDefault();
      suppressedBrowserHistoryInput = null;
      if (browserHistorySuppressionTimer !== null) window.clearTimeout(browserHistorySuppressionTimer);
      browserHistorySuppressionTimer = null;
      return;
    }
    if (event.inputType === 'historyUndo' && undo()) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (event.inputType === 'historyRedo' && redo()) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (
      event.isComposing ||
      compositionActive ||
      pendingCompositionTransaction ||
      event.inputType === 'insertCompositionText' ||
      event.inputType === 'deleteCompositionText' ||
      event.inputType === 'insertFromComposition'
    ) {
      return;
    }
    if (selectedImage && event.inputType === 'deleteContentBackward') {
      event.preventDefault();
      const point = { paraId: selectedImage.paraId, offset: 0 };
      deleteImageWithHistory(selectedImage.paraId, selectedImage.drawingId, point);
      return;
    }
    const domSelection = currentNativeSelection(host);
    const nativeSelection =
      domSelection?.kind === 'text-range'
        ? domSelection
        : selection?.kind === 'text-range'
          ? selection
          : domSelection;
    if (
      trackRevisions &&
      nativeSelection &&
      (event.inputType === 'insertText' ||
        event.inputType === 'deleteContentBackward' ||
        event.inputType === 'deleteContentForward')
    ) {
      const sameParagraph = nativeSelection.kind === 'caret' ||
        (nativeSelection.kind === 'text-range' && nativeSelection.anchor.paraId === nativeSelection.focus.paraId);
      if (sameParagraph) {
        event.preventDefault();
        if (event.inputType === 'insertText' && event.data) {
          if (nativeSelection.kind === 'caret') dispatch({ kind: 'insert-text', text: event.data });
          else dispatch({ kind: 'replace-range', anchor: nativeSelection.anchor, focus: nativeSelection.focus, text: event.data });
        } else if (event.inputType === 'deleteContentBackward') {
          if (nativeSelection.kind === 'text-range' && nativeSelection.anchor.offset !== nativeSelection.focus.offset) {
            dispatch({ kind: 'delete-range', anchor: nativeSelection.anchor, focus: nativeSelection.focus });
          } else dispatch({ kind: 'delete-backward' });
        } else if (event.inputType === 'deleteContentForward') {
          if (nativeSelection.kind === 'text-range' && nativeSelection.anchor.offset !== nativeSelection.focus.offset) {
            dispatch({ kind: 'delete-range', anchor: nativeSelection.anchor, focus: nativeSelection.focus });
          } else dispatch({ kind: 'delete-forward' });
        }
        return;
      }
      if (
        nativeSelection.kind === 'text-range' &&
        (event.inputType === 'deleteContentBackward' || event.inputType === 'deleteContentForward')
      ) {
        event.preventDefault();
        dispatch({ kind: 'delete-range', anchor: nativeSelection.anchor, focus: nativeSelection.focus });
        return;
      }
    }
    if (nativeSelection?.kind === 'text-range' && nativeSelection.anchor.paraId === nativeSelection.focus.paraId) {
      const range = normalizedSelectionRange(nativeSelection.anchor.offset, nativeSelection.focus.offset);
      if (
        canBrowserEditParagraphText(nativeSelection.anchor.paraId) &&
        ((event.inputType === 'insertText' && event.data) ||
          event.inputType === 'deleteContentBackward' ||
          event.inputType === 'deleteContentForward')
      ) {
        selectedImage = null;
        captureTypingLayoutBaseline(nativeSelection.anchor.paraId);
        beginTableBrowserTypingTransaction({
          anchor: { paraId: nativeSelection.anchor.paraId, offset: range.from },
          focus: { paraId: nativeSelection.focus.paraId, offset: range.to },
        });
        setSelection(nativeSelection);
        return;
      }
      if (event.inputType === 'insertText' && event.data) {
        event.preventDefault();
        selectedImage = null;
        setSelection({ kind: 'caret', paraId: nativeSelection.anchor.paraId, offset: range.from + event.data.length });
        record(nativeReplaceRange(doc, nativeSelection.anchor.paraId, range.from, range.to, event.data), {
          paraId: nativeSelection.anchor.paraId,
          offset: range.from + event.data.length,
        }, { localParaId: nativeSelection.anchor.paraId });
        return;
      }
      if (event.inputType === 'deleteContentBackward' || event.inputType === 'deleteContentForward') {
        event.preventDefault();
        selectedImage = null;
        if (applyLocalPageBreakRangeDelete(nativeSelection)) return;
        setSelection({ kind: 'caret', paraId: nativeSelection.anchor.paraId, offset: range.from });
        record(nativeDeleteRange(doc, nativeSelection.anchor.paraId, range.from, range.to), {
          paraId: nativeSelection.anchor.paraId,
          offset: range.from,
        }, { localParaId: nativeSelection.anchor.paraId });
        return;
      }
    }
    if (nativeSelection?.kind === 'text-range' && nativeSelection.anchor.paraId !== nativeSelection.focus.paraId) {
      const range = orderedSelectionRange(doc, nativeSelection);
      if (event.inputType === 'insertText' && event.data) {
        event.preventDefault();
        selectedImage = null;
        if (applyLocalBlockRangeReplacement(range, event.data)) return;
        applyLocalRangeReplacement(range, event.data);
        return;
      }
      if (event.inputType === 'deleteContentBackward' || event.inputType === 'deleteContentForward') {
        if (applyLocalBlockRangeReplacement(range, '')) {
          event.preventDefault();
          return;
        }
        if (!selectionSpansPages(range) && captureBrowserRangeDelete(range)) {
          selectedImage = null;
          setSelection({ kind: 'caret', paraId: range.anchor.paraId, offset: range.anchor.offset });
          return;
        }
        event.preventDefault();
        selectedImage = null;
        setSelection({ kind: 'caret', paraId: range.anchor.paraId, offset: range.anchor.offset });
        applyLocalRangeReplacement(range, '');
        return;
      }
      event.preventDefault();
      setSelection(nativeSelection);
      return;
    }

    const point = nativeSelection?.kind === 'caret' ? nativeSelection : currentNativeCursor(host);
    if (!point) return;
    selectedImage = null;
    setSelection({ kind: 'caret', paraId: point.paraId, offset: point.offset });
    if (event.inputType === 'insertText' && event.data && event.data !== '\n' && event.data !== '\r') {
      if (canBrowserEditParagraphText(point.paraId)) {
        captureTypingLayoutBaseline(point.paraId);
        beginTableBrowserTypingTransaction({ anchor: point, focus: point });
        return;
      }
    }
    if (event.inputType === 'deleteContentBackward') {
      if (point.offset > 0 && canBrowserEditParagraphText(point.paraId)) {
        captureTypingLayoutBaseline(point.paraId);
        beginTableBrowserTypingTransaction({ anchor: point, focus: point });
        return;
      }
    }
    if (event.inputType === 'deleteContentForward') {
      if (point.offset < paragraphTextLength(point.paraId) && canBrowserEditParagraphText(point.paraId)) {
        captureTypingLayoutBaseline(point.paraId);
        beginTableBrowserTypingTransaction({ anchor: point, focus: point });
        return;
      }
    }
    if (
      (event.inputType === 'insertText' && (event.data === '\n' || event.data === '\r')) ||
      event.inputType === 'insertParagraph'
    ) {
      event.preventDefault();
      if (suppressNextEnterBeforeInput) {
        suppressNextEnterBeforeInput = false;
        if (clearEnterSuppressionTimer !== null) window.clearTimeout(clearEnterSuppressionTimer);
        clearEnterSuppressionTimer = null;
        return;
      }
      splitParagraphLocally(point);
    } else if (event.inputType === 'insertText' && event.data) {
      event.preventDefault();
      record(nativeInsertText(doc, point.paraId, point.offset, event.data), {
        paraId: point.paraId,
        offset: point.offset + event.data.length,
      }, { localParaId: point.paraId });
    } else if (event.inputType === 'deleteContentBackward') {
      event.preventDefault();
      if (applyLocalPageBreakEdit(point, 'delete-backward')) return;
      if (deletePageBreakAcrossParagraphBoundary({ kind: 'caret', ...point }, 'backward')) return;
      if (mergeParagraphBackwardLocally(point)) return;
      const paragraph = findNativeParagraph(doc, point.paraId);
      const beforeParagraph = paragraph ? cloneNativeParagraph(paragraph) : null;
      const change = nativeDeleteBackward(doc, point.paraId, point.offset);
      const cursor = {
        paraId: point.offset > 0 ? point.paraId : previousNativeParagraphId(doc, point.paraId) ?? point.paraId,
        offset: Math.max(0, point.offset - 1),
      };
      if (change?.kind === 'image') {
        recordImageDeletion(beforeParagraph, change, cursor);
      } else {
        record(change, cursor, change?.kind === 'block' ? {} : { localParaId: point.paraId });
      }
    } else if (event.inputType === 'deleteContentForward') {
      event.preventDefault();
      if (applyLocalPageBreakEdit(point, 'delete-forward')) return;
      if (deletePageBreakAcrossParagraphBoundary({ kind: 'caret', ...point }, 'forward')) return;
      if (mergeParagraphForwardLocally(point)) return;
      const paragraph = findNativeParagraph(doc, point.paraId);
      const beforeParagraph = paragraph ? cloneNativeParagraph(paragraph) : null;
      const change = nativeDeleteForward(doc, point.paraId, point.offset);
      const cursor = {
        paraId: point.paraId,
        offset: point.offset,
      };
      if (change?.kind === 'image') {
        recordImageDeletion(beforeParagraph, change, cursor);
      } else {
        record(change, cursor, { localParaId: point.paraId });
      }
    }
  };

  const keyDown = (event: KeyboardEvent) => {
    if (destroyed) return;
    if (event.currentTarget === document) {
      if (event.composedPath().includes(host) || activeNativeWordHostByDocument.get(document) !== host) return;
    } else {
      activate();
    }
    if (compositionActive || event.isComposing || event.keyCode === 229) return;
    const isCopy =
      (event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === 'c';
    if (
      isCopy &&
      selectedImage &&
      typeof navigator.clipboard?.write === 'function' &&
      typeof ClipboardItem !== 'undefined'
    ) {
      const imageSelection = selectedImage;
      event.preventDefault();
      event.stopPropagation();
      void copySelectedImageToSystemClipboard(imageSelection).catch((error: unknown) => {
        console.warn('Unable to copy the selected image', error);
      });
      return;
    }
    const isUndo = (event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === 'z';
    const isRedo =
      ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 'z') ||
      ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y');
    if ((isUndo || isRedo) && suppressedBrowserHistoryInput) {
      suppressedBrowserHistoryInput = null;
      if (browserHistorySuppressionTimer !== null) window.clearTimeout(browserHistorySuppressionTimer);
      browserHistorySuppressionTimer = null;
    }
    const historyPoint = isUndo || isRedo ? currentNativeCursor(host) : null;
    const historyParaId = historyPoint?.paraId ?? (selection && 'paraId' in selection ? selection.paraId : null);
    if (isUndo && undo()) {
      suppressBrowserHistoryInput('historyUndo');
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (isRedo && redo()) {
      suppressBrowserHistoryInput('historyRedo');
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key === 'Enter') {
      const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
      if (current?.kind === 'caret' && dispatch({ kind: 'insert-page-break' })) {
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }
    if (isUndo || isRedo) {
      if (nativeHistorySyncTimer !== null) window.clearTimeout(nativeHistorySyncTimer);
      tableResizeObserver.disconnect();
      nativeHistorySyncTimer = window.setTimeout(() => {
        nativeHistorySyncTimer = null;
        const cursor = currentNativeCursor(host);
        const paraId = cursor?.paraId ?? historyParaId;
        if (!paraId) return;
        const paragraphEl = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${paraId}"]`);
        const paragraph = findNativeParagraph(doc, paraId);
        if (paragraphEl && paragraph && paragraph.content.every((child) => child.kind === 'text')) {
          const change = nativeSyncParagraphText(doc, paraId, paragraphEl.textContent ?? '');
          if (change) {
            changes.push(change);
            notifyChange('deferred');
          }
        }
        reflowFromParagraph(paraId, cursor ?? undefined);
      }, 120);
    }
    if (selection?.kind === 'table-cell-range') {
      if (event.key === 'Escape') {
        event.preventDefault();
        setSelection(null);
        return;
      }
      if (event.key === 'Enter' || event.key === 'F2') {
        const table = findNativeTableLocation(doc, selection.tableId)?.table;
        const cell = table ? nativeTableCellsInRange(table, selection)[0]?.cell : null;
        const paragraph = cell?.paragraphs[0];
        if (paragraph) {
          event.preventDefault();
          const point = { paraId: paragraph.paraId, offset: nativeParagraphText(paragraph).length };
          restoreNativeCursor(host, point);
          setSelection({ kind: 'caret', ...point });
        }
        return;
      }
      const movement = event.key === 'ArrowLeft'
        ? { row: 0, column: -1 }
        : event.key === 'ArrowRight'
          ? { row: 0, column: 1 }
          : event.key === 'ArrowUp'
            ? { row: -1, column: 0 }
            : event.key === 'ArrowDown'
              ? { row: 1, column: 0 }
              : null;
      if (movement) {
        const table = findNativeTableLocation(doc, selection.tableId)?.table;
        const columnCount = table
          ? Math.max(1, ...table.rows.map((row) => row.cells.reduce((sum, cell) => sum + Math.max(1, cell.gridSpan), 0)))
          : 1;
        const focus = {
          rowIndex: Math.max(0, Math.min((table?.rows.length ?? 1) - 1, selection.focus.rowIndex + movement.row)),
          columnIndex: Math.max(0, Math.min(columnCount - 1, selection.focus.columnIndex + movement.column)),
        };
        event.preventDefault();
        setSelection({
          kind: 'table-cell-range',
          tableId: selection.tableId,
          anchor: event.shiftKey ? selection.anchor : focus,
          focus,
        });
        return;
      }
    }
    if (
      !event.shiftKey &&
      ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'PageUp', 'PageDown'].includes(event.key)
    ) {
      activeLocalTypingTransaction = null;
    }
    if (selectedImage && (event.key === 'Backspace' || event.key === 'Delete')) {
      event.preventDefault();
      const point = { paraId: selectedImage.paraId, offset: 0 };
      deleteImageWithHistory(selectedImage.paraId, selectedImage.drawingId, point);
      return;
    }
    if (selection?.kind === 'table-cell-range' && (event.key === 'Backspace' || event.key === 'Delete')) {
      event.preventDefault();
      event.stopPropagation();
      applyLocalTableCellClear(selection);
      return;
    }
    if (event.key !== 'Enter') return;
    const point = currentNativeCursor(host);
    if (!point) return;
    event.preventDefault();
    event.stopPropagation();
    const paragraph = findNativeParagraph(doc, point.paraId);
    if (
      paragraph?.format?.hasNumbering &&
      paragraph.format.numbering &&
      nativeParagraphText(paragraph).length === 0
    ) {
      selectedImage = null;
      setSelection({ kind: 'caret', paraId: point.paraId, offset: 0 });
      dispatch({ kind: 'remove-list' });
      return;
    }
    selectedImage = null;
    suppressNextEnterBeforeInput = true;
    if (clearEnterSuppressionTimer !== null) window.clearTimeout(clearEnterSuppressionTimer);
    clearEnterSuppressionTimer = window.setTimeout(() => {
      suppressNextEnterBeforeInput = false;
      clearEnterSuppressionTimer = null;
    }, 0);
    splitParagraphLocally(point);
  };

  const documentBeforeInput = (event: InputEvent) => {
    if (
      destroyed ||
      event.composedPath().includes(host) ||
      activeNativeWordHostByDocument.get(document) !== host
    ) {
      return;
    }
    if (event.inputType === 'historyUndo' && undo()) {
      event.preventDefault();
      event.stopPropagation();
    } else if (event.inputType === 'historyRedo' && redo()) {
      event.preventDefault();
      event.stopPropagation();
    }
  };

  const setRevisionGroupExpanded = (group: HTMLElement, expanded: boolean) => {
    const list = group.querySelector<HTMLElement>(':scope > .coco-word-native-revision-group-list');
    const toggle = group.querySelector<HTMLElement>(':scope > .coco-word-native-revision-group-toggle');
    if (!list || !toggle) return;
    group.classList.toggle('is-expanded', expanded);
    list.hidden = !expanded;
    toggle.setAttribute('aria-expanded', String(expanded));
    const affordance = toggle.querySelector<HTMLElement>('.coco-word-native-revision-group-affordance');
    if (affordance) affordance.textContent = expanded ? '收起' : '展开';
  };

  const collapseOtherRevisionGroups = (except?: HTMLElement) => {
    host.querySelectorAll<HTMLElement>('.coco-word-native-revision-group.is-expanded').forEach((group) => {
      if (group !== except) setRevisionGroupExpanded(group, false);
    });
  };

  const click = (event: MouseEvent) => {
    const revisionAction = event.composedPath().find((item): item is HTMLButtonElement =>
      item instanceof HTMLButtonElement && item.matches('[data-revision-action][data-revision-id]'),
    ) ?? (event.target instanceof Element
      ? event.target.closest<HTMLButtonElement>('[data-revision-action][data-revision-id]')
      : null);
    if (revisionAction) {
      event.preventDefault();
      event.stopPropagation();
      const revisionId = revisionAction.dataset.revisionId;
      if (revisionId) {
        dispatch({
          kind: revisionAction.dataset.revisionAction === 'accept' ? 'accept-revision' : 'reject-revision',
          revisionId,
          revisionKind: revisionAction.dataset.revisionKind as 'insert' | 'delete' | 'format' | undefined,
        });
      }
      return;
    }
    const revisionGroupToggle = event.composedPath().find((item): item is HTMLButtonElement =>
      item instanceof HTMLButtonElement && item.matches('[data-revision-group-toggle]'),
    ) ?? (event.target instanceof Element
      ? event.target.closest<HTMLButtonElement>('[data-revision-group-toggle]')
      : null);
    if (revisionGroupToggle) {
      event.preventDefault();
      event.stopPropagation();
      const group = revisionGroupToggle.closest<HTMLElement>('.coco-word-native-revision-group');
      const list = group?.querySelector<HTMLElement>(':scope > .coco-word-native-revision-group-list');
      if (!group || !list) return;
      const shouldExpand = list.hidden;
      collapseOtherRevisionGroups(group);
      setRevisionGroupExpanded(group, shouldExpand);
      return;
    }
    const commentMenu = event.composedPath().find((item): item is HTMLButtonElement =>
      item instanceof HTMLButtonElement && item.matches('[data-comment-menu][data-comment-id]'),
    ) ?? (event.target instanceof Element
      ? event.target.closest<HTMLButtonElement>('[data-comment-menu][data-comment-id]')
      : null);
    if (commentMenu) {
      event.preventDefault();
      event.stopPropagation();
      const commentId = commentMenu.dataset.commentId;
      if (!commentId) return;
      const menu = commentMenu.parentElement?.querySelector<HTMLElement>(`[data-comment-menu-for="${CSS.escape(commentId)}"]`);
      if (!menu) return;
      host.querySelectorAll<HTMLElement>('.coco-word-native-comment-menu:not([hidden])').forEach((item) => {
        if (item !== menu) item.hidden = true;
      });
      menu.hidden = !menu.hidden;
      return;
    }
    const commentAction = event.composedPath().find((item): item is HTMLButtonElement =>
      item instanceof HTMLButtonElement && item.matches('[data-comment-action][data-comment-id]'),
    ) ?? (event.target instanceof Element
      ? event.target.closest<HTMLButtonElement>('[data-comment-action][data-comment-id]')
      : null);
    if (commentAction) {
      event.preventDefault();
      event.stopPropagation();
      const commentId = Number(commentAction.dataset.commentId);
      if (!Number.isFinite(commentId)) return;
      const bubble = commentAction.closest<HTMLElement>('.coco-word-native-comment-bubble');
      if (commentAction.dataset.commentAction === 'edit') {
        openCommentDialog('edit', commentId, bubble?.dataset.commentBody ?? '');
      } else if (commentAction.dataset.commentAction === 'reply') {
        openCommentDialog('reply', commentId);
      } else {
        dispatch({ kind: 'delete-comment', commentId });
      }
      return;
    }
    if (suppressTextClick) {
      suppressTextClick = false;
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (suppressTableSelectionClick) {
      suppressTableSelectionClick = false;
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (preserveSelectionOnClick) {
      event.preventDefault();
      event.stopPropagation();
      renderPreservedSelectionHighlight(true);
      queueMicrotask(() => {
        preserveSelectionOnClick = false;
        renderPreservedSelectionHighlight(true);
      });
      return;
    }
    activate();
    activeLocalTypingTransaction = null;
    const target = event.target;
    const annotationMark = target instanceof Element
      ? target.closest<HTMLElement>('.coco-word-comment[data-comment-id]')
      : null;
    const revisionMark = target instanceof Element
      ? target.closest<HTMLElement>('.coco-word-revision-insert[data-revision-id], .coco-word-revision-delete[data-revision-id], .coco-word-revision-format[data-revision-id], img[data-revision-id], coco-word-native-table[data-revision-id]')
      : null;
    const annotationBubble = target instanceof Element
      ? target.closest<HTMLElement>('.coco-word-native-comment-bubble, .coco-word-native-revision-bubble')
      : null;
    const annotationTarget = annotationBubble ?? (annotationMark
      ? host.querySelector<HTMLElement>(`.coco-word-native-comment-bubble[data-comment-id="${CSS.escape(annotationMark.dataset.commentId ?? '')}"]`)
        : revisionMark
        ? host.querySelector<HTMLElement>(`.coco-word-native-revision-bubble[data-revision-id="${CSS.escape(revisionMark.dataset.revisionId ?? '')}"][data-revision-kind="${CSS.escape(revisionMark.dataset.revisionKind ?? (revisionMark.classList.contains('coco-word-revision-delete') ? 'delete' : 'insert'))}"]`)
        : null);
    if (annotationTarget) {
      const revisionGroup = annotationTarget.closest<HTMLElement>('.coco-word-native-revision-group');
      if (revisionGroup) {
        collapseOtherRevisionGroups(revisionGroup);
        setRevisionGroupExpanded(revisionGroup, true);
      }
      const paraId = annotationTarget.dataset.paraId;
      const from = Number(annotationTarget.dataset.from);
      const to = Number(annotationTarget.dataset.to);
      if (paraId && Number.isFinite(from) && Number.isFinite(to)) {
        event.preventDefault();
        event.stopPropagation();
        host.querySelectorAll<HTMLElement>('.coco-word-native-comment-bubble.is-selected, .coco-word-native-revision-bubble.is-selected')
          .forEach((bubble) => bubble.classList.remove('is-selected'));
        annotationTarget.classList.add('is-selected');
        const nextSelection: WordEditorSelection = {
          kind: 'text-range',
          anchor: { paraId, offset: from },
          focus: { paraId, offset: Math.max(from + 1, to) },
        };
        setSelection(nextSelection);
        const sourceMark = annotationMark ?? (() => {
          const page = annotationBubble?.closest<HTMLElement>('coco-word-native-page');
          if (!page) return null;
          const kind = annotationTarget.dataset.revisionKind;
          const selector = kind
            ? `.coco-word-revision-insert[data-revision-id="${CSS.escape(annotationTarget.dataset.revisionId ?? '')}"], .coco-word-revision-delete[data-revision-id="${CSS.escape(annotationTarget.dataset.revisionId ?? '')}"], .coco-word-revision-format[data-revision-id="${CSS.escape(annotationTarget.dataset.revisionId ?? '')}"], img[data-revision-id="${CSS.escape(annotationTarget.dataset.revisionId ?? '')}"], coco-word-native-table[data-revision-id="${CSS.escape(annotationTarget.dataset.revisionId ?? '')}"]`
            : `.coco-word-comment[data-comment-id="${CSS.escape(annotationTarget.dataset.commentId ?? '')}"]`;
          return page.querySelector<HTMLElement>(selector);
        })();
        if (!sourceMark || !restoreNativeElementSelection(host, sourceMark)) {
          restoreNativeTextSelection(host, nextSelection.anchor, nextSelection.focus);
        }
        return;
      }
    }
    const tableCell = tableCellTarget(target);
    if (tableCell) {
      event.preventDefault();
      event.stopPropagation();
      selectedImage = null;
      clearNativeImageSelection(host);
      nativeSelection(host)?.removeAllRanges();
      if (
        selection?.kind === 'table-cell-range' &&
        selection.tableId === tableCell.tableId &&
        tableCell.point.rowIndex >= Math.min(selection.anchor.rowIndex, selection.focus.rowIndex) &&
        tableCell.point.rowIndex <= Math.max(selection.anchor.rowIndex, selection.focus.rowIndex) &&
        tableCell.point.columnIndex >= Math.min(selection.anchor.columnIndex, selection.focus.columnIndex) &&
        tableCell.point.columnIndex <= Math.max(selection.anchor.columnIndex, selection.focus.columnIndex)
      ) {
        return;
      }
      setSelection({
        kind: 'table-cell-range',
        tableId: tableCell.tableId,
        anchor: tableCell.point,
        focus: tableCell.point,
      });
      return;
    }
    const hyperlink = target instanceof Element ? target.closest<HTMLAnchorElement>('a.coco-word-hyperlink') : null;
    if (hyperlink) {
      event.preventDefault();
      if ((event.ctrlKey || event.metaKey) && safeExternalHyperlink(hyperlink.dataset.hyperlinkTarget)) {
        window.open(hyperlink.dataset.hyperlinkTarget, '_blank', 'noopener,noreferrer');
        return;
      }
    }
    if (!(target instanceof HTMLImageElement)) {
      const clickedParagraph = target instanceof Node ? parentNativeParagraph(target) : null;
      const currentSelection = currentNativeSelection(host);
      selectedImage = null;
      clearNativeImageSelection(host);
      // Text clicks and drag selection belong to the browser's native
      // contenteditable selection. The selectionchange listener mirrors it to
      // the document model; restoring it here moves the caret to the wrong spot.
      if (clickedParagraph) return;
      if (currentSelection?.kind === 'text-range') {
        setSelection(currentSelection);
        return;
      }
      const point = currentNativeCursor(host);
      if (point) restoreNativeCursor(host, point);
      setSelection(point ? { kind: 'caret', paraId: point.paraId, offset: point.offset } : null);
      return;
    }
    const paragraph = parentNativeParagraph(target);
    const paraId = paragraph?.dataset.paraId;
    const drawingId = target.dataset.drawingId;
    if (!paragraph || !paraId || !drawingId) return;
    paragraph.focus();
    selectedImage = { paraId, drawingId };
    clearNativeImageSelection(host);
    showImageResizeHandle();
    setSelection({ kind: 'object', paraId, objectKind: 'image', objectId: drawingId, offset: imageOffset(paragraph, target) });
  };

  const mouseDown = (event: MouseEvent) => {
    if (event.button !== 0) return;
    preserveSelectionForExternalInteraction = false;
    if (event.target instanceof Node && parentNativeParagraph(event.target)) lastTextSelection = null;
    if (event.target instanceof HTMLImageElement) return;
    const pointerTarget = event.target instanceof Element ? event.target : null;
    const clickedParagraph = pointerTarget ? parentNativeParagraph(pointerTarget) : null;
    const clickedDocument = pointerTarget?.closest('coco-word-native-page, coco-word-native-p, coco-word-native-table') !== null;
    const clickedControl = pointerTarget?.closest('button, input, textarea, select, a');
    if (!clickedParagraph && !clickedDocument && !clickedControl && selection?.kind === 'text-range') {
      preserveSelectionOnClick = true;
      renderPreservedSelectionHighlight(true);
    } else {
      preserveSelectionOnClick = false;
    }
    const target = tableCellTarget(event.target);
    if (target) preserveSelectionOnClick = false;
    if (!target) {
      tableDragStart = null;
      textDragActive = true;
      textDragMoved = false;
      return;
    }
    textDragActive = false;
    textDragMoved = false;
    let anchor = target.point;
    if (event.shiftKey) {
      if (selection?.kind === 'table-cell-range' && selection.tableId === target.tableId) {
        anchor = selection.anchor;
      } else if (selection && selection.kind !== 'table-cell-range') {
        const paraId = selection.kind === 'text-range' ? selection.anchor.paraId : selection.paraId;
        const location = findNativeTableCellLocation(doc, paraId);
        if (location?.table.tableId === target.tableId) {
          anchor = { rowIndex: location.rowIndex, columnIndex: location.columnIndex };
        }
      }
      event.preventDefault();
      nativeSelection(host)?.removeAllRanges();
      tableDragStart = { tableId: target.tableId, point: anchor };
      tableDragSelectionActive = true;
      setSelection({ kind: 'table-cell-range', tableId: target.tableId, anchor, focus: target.point });
      return;
    }
    tableDragStart = { tableId: target.tableId, point: target.point };
    tableDragSelectionActive = false;
  };

  const mouseMove = (event: MouseEvent) => {
    if (textDragActive && (event.buttons & 1) !== 0) {
      textDragMoved = true;
      return;
    }
    if (!tableDragStart || (event.buttons & 1) === 0) return;
    const target = tableCellTarget(event.target);
    if (!target || target.tableId !== tableDragStart.tableId) return;
    if (
      !tableDragSelectionActive &&
      target.point.rowIndex === tableDragStart.point.rowIndex &&
      target.point.columnIndex === tableDragStart.point.columnIndex
    ) {
      return;
    }
    event.preventDefault();
    nativeSelection(host)?.removeAllRanges();
    tableDragSelectionActive = true;
    setSelection({
      kind: 'table-cell-range',
      tableId: tableDragStart.tableId,
      anchor: tableDragStart.point,
      focus: target.point,
    });
  };

  const mouseUp = () => {
    if (textDragActive) {
      const currentSelection = currentNativeSelection(host);
      if (currentSelection?.kind === 'text-range') {
        lastTextSelection = currentSelection;
        setSelection(currentSelection);
        suppressTextClick = true;
      }
      textDragActive = false;
      textDragMoved = false;
    }
    if (tableDragSelectionActive) {
      nativeSelection(host)?.removeAllRanges();
      renderTableCellRangeSelection(selection);
      suppressTableSelectionClick = true;
      tableDragSelectionActive = false;
      tableDragStart = null;
      return;
    }
    tableDragStart = null;
    const currentSelection = currentNativeSelection(host);
    if (currentSelection?.kind === 'text-range') {
      lastTextSelection = currentSelection;
      selectedImage = null;
      clearNativeImageSelection(host);
      setSelection(currentSelection);
    }
  };
  const windowMouseUp = () => {
    if (tableDragStart) mouseUp();
    else if (textDragActive) mouseUp();
  };

  const selectionChange = () => {
    if (destroyed) return;
    if (compositionActive || pendingCompositionTransaction) return;
    if (selection?.kind === 'object' && selectedImage) return;
    if (selection?.kind === 'table-cell-range') return;
    if (suppressTextClick && lastTextSelection) {
      setSelection(lastTextSelection);
      renderPreservedSelectionHighlight(true);
      return;
    }
    if (preserveSelectionOnClick) {
      renderPreservedSelectionHighlight(true);
      return;
    }
    if (preserveSelectionForExternalInteraction) {
      renderPreservedSelectionHighlight(true);
      return;
    }
    const activeElement = document.activeElement;
    if (
      activeElement instanceof HTMLElement &&
      activeElement !== document.body &&
      activeElement !== focusHost &&
      !host.contains(activeElement)
    ) {
      renderPreservedSelectionHighlight();
      return;
    }
    clearPreservedSelectionHighlight();
    const current = currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
    if (current?.kind === 'text-range') lastTextSelection = current;
    if (current) {
      activate();
      setSelection(current);
    }
  };
  const documentFocusIn = (event: FocusEvent) => {
    const target = event.target;
    const eventPath = event.composedPath();
    const focusedDocument = eventPath.some(
      (item) => item instanceof Element && item.matches('coco-word-native-page, coco-word-native-p, coco-word-native-table'),
    );
    if (focusedDocument) {
      preserveSelectionForExternalInteraction = false;
    } else if (selection?.kind === 'text-range') {
      preserveSelectionForExternalInteraction = true;
    }
    queueMicrotask(renderPreservedSelectionHighlight);
  };
  const preserveSelectionFromExternalPointer = (event: PointerEvent) => {
    const eventPath = event.composedPath();
    const clickedDocument = eventPath.some(
      (item) => item instanceof Element && item.matches('coco-word-native-page, coco-word-native-p, coco-word-native-table'),
    );
    const clickedRevisionAction = eventPath.some(
      (item) => item instanceof HTMLButtonElement && item.matches('[data-revision-action][data-revision-id], [data-revision-group-toggle]'),
    );
    if (destroyed || event.button !== 0 || clickedDocument || clickedRevisionAction) return;
    const current = lastTextSelection ?? (selection?.kind === 'text-range' ? selection : currentNativeSelection(host));
    if (current?.kind !== 'text-range') return;
    if (selection !== current) setSelection(current);
    preserveSelectionForExternalInteraction = true;
    preserveSelectionOnClick = true;
    // Keep the document selection visible while toolbar/sidebar controls take focus.
    // A click inside a paragraph is intentionally left to the browser's native selection.
    renderPreservedSelectionHighlight(true);
    queueMicrotask(() => renderPreservedSelectionHighlight(true));
    window.setTimeout(() => renderPreservedSelectionHighlight(true), 0);
  };
  const preserveSelectionAfterExternalClick = (event: MouseEvent) => {
    preserveSelectionFromExternalPointer(event as PointerEvent);
    window.setTimeout(() => renderPreservedSelectionHighlight(true), 0);
  };
  const updateHyperlinkNavigationMode = (event: KeyboardEvent) => {
    host.classList.toggle('is-hyperlink-navigation-mode', event.ctrlKey || event.metaKey);
  };
  const clearHyperlinkNavigationMode = () => {
    host.classList.remove('is-hyperlink-navigation-mode');
  };

  const input = (event: Event) => {
    const inputEvent = event as InputEvent;
    if (inputEvent.inputType === suppressedBrowserHistoryInput) {
      suppressedBrowserHistoryInput = null;
      if (browserHistorySuppressionTimer !== null) window.clearTimeout(browserHistorySuppressionTimer);
      browserHistorySuppressionTimer = null;
      return;
    }
    if (
      inputEvent.isComposing ||
      compositionActive ||
      inputEvent.inputType === 'insertCompositionText' ||
      inputEvent.inputType === 'deleteCompositionText'
    ) {
      return;
    }
    if (pendingCompositionTransaction) {
      if (compositionFinalizeTimer !== null) window.clearTimeout(compositionFinalizeTimer);
      finalizeComposition();
      return;
    }
    if (syncBrowserRangeDelete()) return;
    if (inputEvent.inputType === 'historyUndo' && syncBrowserHistoryUndo()) return;
    if (inputEvent.inputType === 'historyRedo' && syncBrowserHistoryRedo()) return;

    const cursor = currentNativeCursor(host);
    const paragraphEl =
      parentNativeParagraph(event.target as Node) ??
      (cursor ? host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${cursor.paraId}"]`) : null);
    if (!paragraphEl || !host.contains(paragraphEl)) return;
    const paraId = paragraphEl.dataset.paraId ?? '';
    const paragraph = findNativeParagraph(doc, paraId);
    if (!paragraph || paragraph.content.some((child) => child.kind !== 'text')) return;
    const change = nativeSyncParagraphText(doc, paraId, paragraphEl.textContent ?? '');
    if (!change) {
      if (inputEvent.inputType === 'historyUndo' || inputEvent.inputType === 'historyRedo') scheduleReflow(paraId);
      return;
    }
    changes.push(change);
    const revision = markBrowserTextChangeAsRevision(paraId, change);
    if (revision) changes.push(revision);
    if (
      activeLocalTypingTransaction &&
      inputEvent.inputType !== 'historyUndo' &&
      inputEvent.inputType !== 'historyRedo' &&
      activeLocalTypingTransaction.afterParagraphs[0]?.paraId === paraId
    ) {
      activeLocalTypingTransaction.changes.push(change);
      activeLocalTypingTransaction.afterParagraphs = [cloneNativeParagraph(paragraph)];
    } else {
      activeLocalTypingTransaction = null;
    }
    browserRedoStack.length = 0;
    scheduleTypingReflow(paraId);
    if (cursor) {
      setSelection({ kind: 'caret', paraId: cursor.paraId, offset: cursor.offset });
    }
    if (revision) renderParagraph(paraId, cursor ?? undefined);
    notifyChange('deferred');
  };

  const applyPublicSelection = (next: WordEditorSelection | null): boolean => {
    if (destroyed) return false;
    if (!next) {
      const domSelection = nativeSelection(host);
      if (domSelection && (host.contains(domSelection.anchorNode) || host.contains(domSelection.focusNode))) {
        domSelection.removeAllRanges();
      }
      selectedImage = null;
      clearNativeImageSelection(host);
      setSelection(null);
      return true;
    }
    if (next.kind === 'caret') {
      if (!findNativeParagraph(doc, next.paraId)) return false;
      selectedImage = null;
      clearNativeImageSelection(host);
      restoreNativeCursor(host, next);
      setSelection(next);
      return true;
    }
    if (next.kind === 'text-range') {
      if (!findNativeParagraph(doc, next.anchor.paraId) || !findNativeParagraph(doc, next.focus.paraId)) return false;
      selectedImage = null;
      clearNativeImageSelection(host);
      restoreNativeTextSelection(host, next.anchor, next.focus);
      setSelection(next);
      return true;
    }
    if (next.kind === 'table-cell-range') {
      const table = findNativeTableLocation(doc, next.tableId)?.table;
      if (!table || nativeTableCellsInRange(table, next).length === 0) return false;
      selectedImage = null;
      clearNativeImageSelection(host);
      nativeSelection(host)?.removeAllRanges();
      setSelection(next);
      return true;
    }
    const image = host.querySelector<HTMLImageElement>(
      `coco-word-native-p[data-para-id="${next.paraId}"] img[data-drawing-id="${next.objectId}"]`,
    );
    if (!findNativeParagraph(doc, next.paraId) || !image) return false;
    selectedImage = { paraId: next.paraId, drawingId: next.objectId };
    clearNativeImageSelection(host);
    showImageResizeHandle();
    setSelection(next);
    return true;
  };

  const cancelScheduledWork = () => {
    if (deferredRenderTimer !== null) window.clearTimeout(deferredRenderTimer);
    if (deferredReflowTimer !== null) window.clearTimeout(deferredReflowTimer);
    if (nativeHistorySyncTimer !== null) window.clearTimeout(nativeHistorySyncTimer);
    if (deferredChangeTimer !== null) window.clearTimeout(deferredChangeTimer);
    if (deferredSelectionTimer !== null) window.clearTimeout(deferredSelectionTimer);
    if (clearEnterSuppressionTimer !== null) window.clearTimeout(clearEnterSuppressionTimer);
    if (compositionFinalizeTimer !== null) window.clearTimeout(compositionFinalizeTimer);
    if (browserHistorySuppressionTimer !== null) window.clearTimeout(browserHistorySuppressionTimer);
    deferredRenderTimer = null;
    deferredReflowTimer = null;
    nativeHistorySyncTimer = null;
    deferredChangeTimer = null;
    deferredSelectionTimer = null;
    clearEnterSuppressionTimer = null;
    compositionFinalizeTimer = null;
    browserHistorySuppressionTimer = null;
  };

  host.addEventListener('beforeinput', beforeInput);
  host.addEventListener('compositionstart', compositionStart);
  host.addEventListener('compositionend', compositionEnd);
  host.addEventListener('copy', copy);
  host.addEventListener('cut', cut);
  host.addEventListener('paste', paste);
  host.addEventListener('keydown', keyDown);
  host.addEventListener('mousedown', mouseDown);
  host.addEventListener('click', click);
  host.addEventListener('mouseup', mouseUp);
  host.addEventListener('input', input);
  host.addEventListener('contextmenu', contextMenu);
  document.addEventListener('keydown', keyDown, true);
  document.addEventListener('beforeinput', documentBeforeInput, true);
  document.addEventListener('selectionchange', selectionChange);
  document.addEventListener('focusin', documentFocusIn);
  document.addEventListener('pointerdown', preserveSelectionFromExternalPointer, true);
  document.addEventListener('pointerup', preserveSelectionFromExternalPointer, true);
  document.addEventListener('click', preserveSelectionAfterExternalClick, true);
  document.addEventListener('keydown', updateHyperlinkNavigationMode, true);
  document.addEventListener('keyup', updateHyperlinkNavigationMode, true);
  document.addEventListener('mousemove', mouseMove);
  document.addEventListener('pointerdown', closeContextMenuFromPointer, true);
  document.addEventListener('keydown', closeContextMenuFromKey, true);
  document.addEventListener('scroll', closeContextMenu, true);
  window.addEventListener('resize', closeContextMenu);
  window.addEventListener('mouseup', windowMouseUp);
  window.addEventListener('blur', clearHyperlinkNavigationMode);
  render();

  return {
    kind: 'native-word',
    get destroyed() {
      return destroyed;
    },
    destroy() {
      if (destroyed) return;
      destroyed = true;
      host.removeEventListener('beforeinput', beforeInput);
      host.removeEventListener('compositionstart', compositionStart);
      host.removeEventListener('compositionend', compositionEnd);
      host.removeEventListener('copy', copy);
      host.removeEventListener('cut', cut);
      host.removeEventListener('paste', paste);
      host.removeEventListener('keydown', keyDown);
      host.removeEventListener('mousedown', mouseDown);
      host.removeEventListener('click', click);
      host.removeEventListener('mouseup', mouseUp);
      host.removeEventListener('input', input);
      host.removeEventListener('contextmenu', contextMenu);
      document.removeEventListener('keydown', keyDown, true);
      document.removeEventListener('beforeinput', documentBeforeInput, true);
      document.removeEventListener('selectionchange', selectionChange);
      document.removeEventListener('focusin', documentFocusIn);
      document.removeEventListener('pointerdown', preserveSelectionFromExternalPointer, true);
      document.removeEventListener('pointerup', preserveSelectionFromExternalPointer, true);
      document.removeEventListener('click', preserveSelectionAfterExternalClick, true);
      document.removeEventListener('keydown', updateHyperlinkNavigationMode, true);
      document.removeEventListener('keyup', updateHyperlinkNavigationMode, true);
      document.removeEventListener('mousemove', mouseMove);
      document.removeEventListener('pointerdown', closeContextMenuFromPointer, true);
      document.removeEventListener('keydown', closeContextMenuFromKey, true);
      document.removeEventListener('scroll', closeContextMenu, true);
      window.removeEventListener('resize', closeContextMenu);
      window.removeEventListener('mouseup', windowMouseUp);
      window.removeEventListener('blur', clearHyperlinkNavigationMode);
      window.removeEventListener('resize', repositionSelectedImageHandle);
      document.removeEventListener('scroll', repositionSelectedImageHandle, true);
      cancelImageResizeDrag?.();
      cancelImageResizeDrag = null;
      closeContextMenu();
      clearSearch();
      clearPreservedSelectionHighlight();
      searchHighlightStyle.remove();
      cancelScheduledWork();
      tableResizeObserver.disconnect();
      if (activeNativeWordHostByDocument.get(document) === host) activeNativeWordHostByDocument.delete(document);
      listeners.change.clear();
      listeners.selectionChange.clear();
      if (shadowRoot) shadowRoot.replaceChildren();
      else host.replaceChildren();
      host.classList.remove('coco-word-native-editor-host');
      host.classList.remove('is-hyperlink-navigation-mode');
      outerHost.classList.remove('coco-word-editor-host');
    },
    loadDocument(nextView) {
      if (destroyed) throw new Error('Cannot load a document into a destroyed Word editor');
      cancelScheduledWork();
      tableResizeObserver.disconnect();
      sourceView = nextView;
      closeContextMenu();
      clearSearch();
      doc = createNativeWordDocument(sourceView);
      changes.length = 0;
      emittedChanges = [];
      paragraphIds.clear();
      for (const paragraph of allNativeParagraphs(doc)) paragraphIds.add(paragraph.paraId);
      nextParagraphId = Date.now() >>> 0;
      selectedImage = null;
      removeImageResizeHandle();
      selection = null;
      pendingBrowserRangeDelete = null;
      pendingCompositionTransaction = null;
      activeLocalTypingTransaction = null;
      compositionActive = false;
      typingLayoutBaseline = null;
      suppressedBrowserHistoryInput = null;
      browserUndoStack.length = 0;
      browserRedoStack.length = 0;
      render();
      notifySelectionChange();
    },
    getSourceView() {
      return sourceView;
    },
    getDocument() {
      return doc;
    },
    getSelection() {
      return selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
    },
    setSelection(next) {
      return applyPublicSelection(next);
    },
    focus() {
      if (destroyed) return false;
      const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
      if (current) return applyPublicSelection(current);
      const first = allNativeParagraphs(doc)[0];
      return first ? applyPublicSelection({ kind: 'caret', paraId: first.paraId, offset: 0 }) : false;
    },
    async insertImage(source, imageOptions = {}) {
      if (destroyed) return false;
      const src = typeof source === 'string' ? source : await blobToDataUrl(source);
      return dispatch({ kind: 'insert-image', src, ...imageOptions });
    },
    search,
    getSearchState() {
      return searchStateSnapshot();
    },
    findNext,
    replaceCurrent,
    replaceAll,
    clearSearch,
    getCommandState() {
      const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
      return getCommandState(current);
    },
    getChanges() {
      return changes.slice();
    },
    canUndo() {
      return !destroyed && browserUndoStack.length > 0;
    },
    canRedo() {
      return !destroyed && browserRedoStack.length > 0;
    },
    undo,
    redo,
    canDispatch(command) {
      if (destroyed) return false;
      const current = selection ?? currentNativeSelection(host) ?? nativeSelectionFromCursor(host);
      return canDispatchNativeWordCommand(doc, current, command);
    },
    dispatch,
    subscribe(event, listener) {
      if (destroyed) return () => {};
      const runtimeListener = listener as (
        payload: WordEditorChangeEvent | WordEditorSelectionChangeEvent,
      ) => void;
      listeners[event].add(runtimeListener);
      return () => listeners[event].delete(runtimeListener);
    },
  };
}

function cloneNativeParagraph(paragraph: NativeWordParagraphNode): NativeWordParagraphNode {
  return {
    ...paragraph,
    format: paragraph.format ? cloneNativeMetadata(paragraph.format) : null,
    content: paragraph.content.map(cloneNativeInline),
  };
}

function cloneNativeInline(child: NativeWordInlineNode): NativeWordInlineNode {
  if (child.kind === 'pageBreak') return { kind: 'pageBreak' };
  if (child.kind === 'text') {
    return {
      ...child,
      format: child.format ? cloneNativeMetadata(child.format) : null,
      ...(child.marks ? { marks: child.marks.map((mark) => ({ ...mark })) } : {}),
      ...(child.hyperlink ? { hyperlink: { ...child.hyperlink } } : {}),
    };
  }
  if (child.kind === 'field') {
    return { ...child, format: child.format ? cloneNativeMetadata(child.format) : null };
  }
  if (child.kind === 'image') {
    return {
      ...child,
      ...(child.position ? { position: cloneNativeMetadata(child.position) } : {}),
    };
  }
  return {
    ...child,
    ...(child.position ? { position: cloneNativeMetadata(child.position) } : {}),
    ...(child.paragraphs ? { paragraphs: child.paragraphs.map(cloneNativeParagraph) } : {}),
  };
}

function cloneNativeMetadata<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function safeExternalHyperlink(value: string | undefined): value is string {
  return !!value && /^(https?:|mailto:|tel:)/i.test(value.trim());
}

function isBrowserBlockTransaction(transaction: BrowserTransaction): transaction is BrowserBlockTransaction {
  return 'beforeBlocks' in transaction;
}

function nativeBlockElements(host: HTMLElement, block: NativeWordBlockNode): HTMLElement[] {
  const selector = block.kind === 'paragraph'
    ? `coco-word-native-p[data-para-id="${block.paraId}"]`
    : `coco-word-native-table[data-table-id="${block.tableId}"]`;
  return Array.from(host.querySelectorAll<HTMLElement>(selector));
}

function allocateNativeCommentId(doc: NativeWordDocument): number {
  let highest = nextNativeCommentId;
  for (const paragraph of allNativeParagraphs(doc)) {
    for (const child of paragraph.content) {
      if (child.kind !== 'text') continue;
      for (const mark of child.marks ?? []) {
        if (mark.kind === 'comment') highest = Math.max(highest, mark.commentId);
      }
    }
  }
  nextNativeCommentId = highest + 1;
  return nextNativeCommentId;
}

function allocateNativeRevisionId(doc: NativeWordDocument): string {
  let highest = nextNativeRevisionId;
  for (const paragraph of allNativeParagraphs(doc)) {
    for (const child of paragraph.content) {
      if (child.kind === 'text') {
        for (const mark of child.marks ?? []) {
          if (mark.kind !== 'revision') continue;
          const numeric = Number(mark.revisionId);
          if (Number.isInteger(numeric)) highest = Math.max(highest, numeric);
        }
        const formatRevision = child.formatRevision;
        const formatNumeric = Number(formatRevision?.revisionId);
        if (Number.isInteger(formatNumeric)) highest = Math.max(highest, formatNumeric);
      } else if (child.kind === 'image') {
        const numeric = Number(child.revision?.revisionId ?? child.formatRevision?.revisionId);
        if (Number.isInteger(numeric)) highest = Math.max(highest, numeric);
      }
    }
  }
  nextNativeRevisionId = highest + 1;
  return String(nextNativeRevisionId);
}

function sameSelection(left: WordEditorSelection | null, right: WordEditorSelection | null): boolean {
  if (left?.kind !== right?.kind) return false;
  if (!left || !right) return left === right;
  if (left.kind === 'caret' && right.kind === 'caret') {
    return left.paraId === right.paraId && left.offset === right.offset;
  }
  if (left.kind === 'object' && right.kind === 'object') {
    return left.paraId === right.paraId && left.objectKind === right.objectKind && left.objectId === right.objectId;
  }
  if (left.kind === 'text-range' && right.kind === 'text-range') {
    return (
      left.anchor.paraId === right.anchor.paraId &&
      left.anchor.offset === right.anchor.offset &&
      left.focus.paraId === right.focus.paraId &&
      left.focus.offset === right.focus.offset
    );
  }
  if (left.kind === 'table-cell-range' && right.kind === 'table-cell-range') {
    return (
      left.tableId === right.tableId &&
      left.anchor.rowIndex === right.anchor.rowIndex &&
      left.anchor.columnIndex === right.anchor.columnIndex &&
      left.focus.rowIndex === right.focus.rowIndex &&
      left.focus.columnIndex === right.focus.columnIndex
    );
  }
  return false;
}

function formatTransactionSelection(
  transaction: BrowserTextTransaction,
  phase: 'before' | 'after',
): WordEditorSelection | null {
  const explicit = phase === 'before' ? transaction.beforeSelection : transaction.afterSelection;
  if (explicit) return explicit;
  if (!transaction.changes.every((change) => change.kind === 'format' || change.kind === 'list')) return null;
  if (
    transaction.range.anchor.paraId === transaction.range.focus.paraId &&
    transaction.range.anchor.offset === transaction.range.focus.offset
  ) {
    return null;
  }
  return { kind: 'text-range', anchor: transaction.range.anchor, focus: transaction.range.focus };
}

function transactionRequiresReflow(transaction: BrowserTextTransaction): boolean {
  if (transaction.skipDeferredRender) return false;
  return transaction.changes.some((change) => {
    if (change.kind === 'block') return true;
    if (change.kind === 'image') return true;
    if (change.kind === 'page-break') return true;
    if (change.kind === 'list') return true;
    if (change.kind !== 'format') return false;
    if (change.op === 'paragraph') return change.format.alignment === undefined;
    return (
      change.clear === true ||
      change.format.bold !== undefined ||
      change.format.italic !== undefined ||
      change.format.fontFamily !== undefined ||
      change.format.fontSize !== undefined
    );
  });
}

function domParagraphText(host: HTMLElement, paraId: string): string | null {
  return host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${paraId}"]`)?.textContent ?? null;
}

function domParagraphsMatch(host: HTMLElement, paragraphs: NativeWordParagraphNode[]): boolean {
  return paragraphs.every((paragraph) => domParagraphText(host, paragraph.paraId) === nativeParagraphText(paragraph));
}

function normalizedSelectionRange(from: number, to: number): { from: number; to: number } {
  return {
    from: Math.max(0, Math.min(from, to)),
    to: Math.max(0, Math.max(from, to)),
  };
}

function blobToDataUrl(file: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener('load', () => resolve(String(reader.result ?? '')));
    reader.addEventListener('error', () => reject(reader.error ?? new Error('Unable to read clipboard image')));
    reader.readAsDataURL(file);
  });
}

function storeNativeImageClipboardSource(src: string): string {
  const safeSrc = safeEmbeddedRasterImageUrl(src);
  if (!safeSrc) return '';
  nextNativeImageClipboardId = (nextNativeImageClipboardId + 1) >>> 0;
  const token = `${NATIVE_IMAGE_CLIPBOARD_PREFIX}${Date.now().toString(36)}-${nextNativeImageClipboardId.toString(36)}`;
  nativeImageClipboardSources.set(token, safeSrc);
  while (nativeImageClipboardSources.size > 8) {
    const oldest = nativeImageClipboardSources.keys().next().value as string | undefined;
    if (!oldest) break;
    nativeImageClipboardSources.delete(oldest);
  }
  return token;
}

function resolveNativeImageClipboardSource(value: string): string {
  if (!value.startsWith(NATIVE_IMAGE_CLIPBOARD_PREFIX)) return value;
  return nativeImageClipboardSources.get(value) ?? '';
}

function nativeClipboardInlineText(child: NativeWordParagraphNode['content'][number]): string {
  if (child.kind === 'text') return child.text;
  if (child.kind === 'field') return child.cachedText ?? '\uFFFC';
  if (child.kind === 'drawing' && child.text) return child.text;
  return '\uFFFC';
}

async function clipboardPngBlob(src: string): Promise<Blob> {
  const source = embeddedRasterImageBlob(src);
  if (source.type === 'image/png') return source;

  const bitmap = await createImageBitmap(source);
  try {
    const canvas = document.createElement('canvas');
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Unable to create clipboard image canvas');
    context.drawImage(bitmap, 0, 0);
    return await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob(
        (blob) => (blob ? resolve(blob) : reject(new Error('Unable to encode clipboard image'))),
        'image/png',
      );
    });
  } finally {
    bitmap.close();
  }
}

function escapeHtmlAttribute(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

function sanitizeIndent(indent: { left?: number; right?: number; firstLine?: number }) {
  const result: { left?: number; right?: number; firstLine?: number } = {};
  if (Number.isFinite(indent.left)) result.left = Math.max(0, indent.left!);
  if (Number.isFinite(indent.right)) result.right = Math.max(0, indent.right!);
  if (Number.isFinite(indent.firstLine)) result.firstLine = indent.firstLine!;
  return result;
}

function sanitizeSpacing(spacing: { before?: number; after?: number }) {
  const result: { before?: number; after?: number } = {};
  if (Number.isFinite(spacing.before)) result.before = Math.max(0, spacing.before!);
  if (Number.isFinite(spacing.after)) result.after = Math.max(0, spacing.after!);
  return result;
}

function orderedSelectionRange(doc: NativeWordDocument, selection: Extract<WordEditorSelection, { kind: 'text-range' }>): {
  anchor: { paraId: string; offset: number };
  focus: { paraId: string; offset: number };
} {
  if (selection.anchor.paraId === selection.focus.paraId && selection.anchor.offset <= selection.focus.offset) {
    return { anchor: selection.anchor, focus: selection.focus };
  }
  if (selection.anchor.paraId !== selection.focus.paraId) {
    const anchorLocation = findNativeParagraphLocation(doc, selection.anchor.paraId);
    const focusLocation = findNativeParagraphLocation(doc, selection.focus.paraId);
    if (anchorLocation && focusLocation && anchorLocation.container === focusLocation.container) {
      return anchorLocation.index <= focusLocation.index
        ? { anchor: selection.anchor, focus: selection.focus }
        : { anchor: selection.focus, focus: selection.anchor };
    }
  }
  return { anchor: selection.focus, focus: selection.anchor };
}
