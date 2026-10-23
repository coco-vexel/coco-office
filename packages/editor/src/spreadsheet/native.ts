import {
  formatCellAddress,
  formatRangeAddress,
  type ExcelCellValue,
  type ExcelColorModel,
  type ExcelChartModel,
  type ExcelDrawingAnchorModel,
  type ExcelDrawingModel,
} from '@coco-vexel/coco-office/excel';
import { freezeStateFromSheet, SpreadsheetViewportManager } from './viewport.js';
import { SpreadsheetCanvasRenderer } from './canvas-renderer.js';
import { SpreadsheetEditorRuntime, type SpreadsheetUserTransactionOptions } from './runtime.js';
import { SpreadsheetLayout } from './layout.js';
import { SpreadsheetCellEditor } from './editing.js';
import {
  createColumnSelection,
  createRowSelection,
  createSelection,
  expandSelectionForMergedRanges,
  extendSelection,
  normalizeSelectionRange,
  type SpreadsheetSelectionState,
} from './selection.js';
import { buildSpreadsheetClipboardPayload, parseTsv, type SpreadsheetClipboardPayload } from './clipboard.js';
import { getRenderCell } from './render-model.js';
import { resolveSpreadsheetCellStyle, type SpreadsheetResolvedCellStyle } from './style-resolver.js';
import { spreadsheetFilterValues, type SpreadsheetColumnFilter } from './filter.js';
import { supportsSpreadsheetAutoFill } from './auto-fill.js';
import {
  hitSpreadsheetResizeHandle,
  startSpreadsheetResizeDrag,
  updateSpreadsheetResizeDrag,
  type SpreadsheetResizeDragState,
} from './resize.js';
import {
  applyWheelDelta,
  computeScrollbarState,
  scrollFromThumbPosition,
  type SpreadsheetScrollbarState,
} from './scrollbar.js';
import type {
  SpreadsheetFreezeState,
  SpreadsheetRange,
  SpreadsheetRect,
  SpreadsheetScrollState,
  SpreadsheetViewport,
} from './types.js';
import { EditorReadOnlyError } from '../read-only.js';

export interface SpreadsheetEditorChangeEvent {
  kind: 'workbookChange';
  bytes: Uint8Array;
}

export interface SpreadsheetEditorStateEvent {
  kind: 'stateChange';
  activeSheetName: string;
  selection: SpreadsheetSelectionState;
  editingCell: { row: number; column: number } | null;
}

export interface SpreadsheetEditorCellState {
  address: string;
  row: number;
  column: number;
  value: ExcelCellValue;
  displayText: string;
  formula?: string;
  styleId?: number;
  format?: {
    font: unknown;
    fill: unknown;
    border: unknown;
    alignment: unknown;
    numberFormatCode?: string;
  };
}

export interface SpreadsheetEditorOptions {
  workbookId?: string;
  /** Allows navigation, selection and copy while blocking every workbook mutation. */
  readOnly?: boolean;
  onChange?: (event: SpreadsheetEditorChangeEvent) => void | Promise<void>;
  showSheetTabs?: boolean;
  /** Isolates editor internals from host-page CSS. Shadow DOM is enabled by default. */
  styleIsolation?: 'shadow' | 'none';
  /** Header filter icons are opt-in; hosts can expose filtering from a column menu instead. */
  enableHeaderFilters?: boolean;
}

export interface SpreadsheetEditor {
  runtime(): SpreadsheetEditorRuntime;
  isReadOnly(): boolean;
  setReadOnly(readOnly: boolean): void;
  selection(): SpreadsheetSelectionState;
  cellState(address: string): SpreadsheetEditorCellState;
  freezeState(): SpreadsheetFreezeState;
  filterState(): ReadonlyMap<number, readonly string[]>;
  activeSheetName(): string;
  setActiveSheet(name: string): Promise<void>;
  addSheet(name: string): Promise<void>;
  renameSheet(sheetName: string, name: string): Promise<void>;
  copySheet(sheetName: string, newName: string): Promise<void>;
  deleteSheet(sheetName: string): Promise<void>;
  moveSheet(sheetName: string, toIndex: number): Promise<void>;
  setCellValue(address: string, value: ExcelCellValue): Promise<void>;
  setCellFormula(address: string, formula: string): Promise<void>;
  formatSelection(kind: 'bold' | 'italic' | 'underline'): Promise<void>;
  setSelectionFontColor(color: ExcelColorModel | undefined): Promise<void>;
  setSelectionFillColor(color: ExcelColorModel | undefined): Promise<void>;
  setSelectionHorizontalAlign(alignment: 'left' | 'center' | 'right'): Promise<void>;
  setSelectionVerticalAlign(alignment: 'top' | 'center' | 'bottom'): Promise<void>;
  setSelectionNumberFormat(code: string | undefined): Promise<void>;
  setSelectionBorder(mode: 'all' | 'outer' | 'inner' | 'clear'): Promise<void>;
  clearSelection(): Promise<void>;
  clearSelectionFormat(): Promise<void>;
  mergeSelection(): Promise<void>;
  unmergeSelection(): Promise<void>;
  freezeSelection(): Promise<void>;
  sortActiveColumn(direction: 'asc' | 'desc'): Promise<void>;
  filterColumn(column: number, values: readonly string[] | null): Promise<void>;
  runTransaction(
    options: SpreadsheetUserTransactionOptions,
    run: (editor: SpreadsheetEditor) => Promise<void>,
  ): Promise<void>;
  loadWorkbook(bytes: Uint8Array | ArrayBuffer): Promise<void>;
  toZip(): Promise<Uint8Array>;
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): Promise<void>;
  redo(): Promise<void>;
  subscribe(listener: (event: SpreadsheetEditorChangeEvent) => void): () => void;
  subscribeState(listener: (event: SpreadsheetEditorStateEvent) => void): () => void;
  destroy(): void;
}

/**
 * Public browser/Electron surface for the spreadsheet editor.
 *
 * The host owns its toolbar and persistence policy. This surface owns the
 * canvas stack, viewport lifecycle, scrolling, and worksheet switching so web
 * and desktop hosts do not rebuild those pieces independently.
 */
export async function renderSpreadsheetEditor(
  host: HTMLElement,
  bytes: Uint8Array | ArrayBuffer,
  options: SpreadsheetEditorOptions = {},
): Promise<SpreadsheetEditor> {
  const controller = new NativeSpreadsheetEditor(host, options);
  await controller.loadWorkbook(bytes);
  return controller;
}

class NativeSpreadsheetEditor implements SpreadsheetEditor {
  private readonly renderRoot: HTMLElement | ShadowRoot;
  private runtimeValue: SpreadsheetEditorRuntime | null = null;
  private renderer: SpreadsheetCanvasRenderer | null = null;
  private layout: SpreadsheetLayout | null = null;
  private viewport: SpreadsheetViewportManager | null = null;
  private scroll: SpreadsheetScrollState = { x: 0, y: 0 };
  private scrollbarState: SpreadsheetScrollbarState | null = null;
  private readonly listeners = new Set<(event: SpreadsheetEditorChangeEvent) => void>();
  private readonly stateListeners = new Set<(event: SpreadsheetEditorStateEvent) => void>();
  private readonly baseCanvas = document.createElement('canvas');
  private readonly headerCanvas = document.createElement('canvas');
  private readonly overlayCanvas = document.createElement('canvas');
  private readonly unsupportedObjectLayer = document.createElement('div');
  private readonly sheetTabs = document.createElement('div');
  private readonly horizontalScrollbar = document.createElement('div');
  private readonly verticalScrollbar = document.createElement('div');
  private readonly horizontalThumb = document.createElement('div');
  private readonly verticalThumb = document.createElement('div');
  private readonly cellEditorElement = document.createElement('textarea');
  private readonly resizeGuide = document.createElement('div');
  private readonly contextMenu = document.createElement('div');
  private readonly filterMenu = document.createElement('div');
  private readonly cellEditor: SpreadsheetCellEditor;
  private readonly resizeObserver: ResizeObserver;
  private disposed = false;
  private drag: { axis: 'horizontal' | 'vertical'; pointer: number; thumb: number } | null = null;
  private selectionState: SpreadsheetSelectionState = createSelection(1, 1);
  private editingCell: { row: number; column: number } | null = null;
  private editingCellStyle: SpreadsheetResolvedCellStyle | null = null;
  private clipboardPayload: SpreadsheetClipboardPayload | null = null;
  private selecting = false;
  private resizeDrag: SpreadsheetResizeDragState | null = null;
  private autoFillDrag: { source: SpreadsheetRange; preview: SpreadsheetRange | null } | null = null;
  private selectedChartIndex: number | null = null;
  private chartDrag: {
    chartIndex: number;
    pointer: { x: number; y: number };
    startRect: SpreadsheetRect;
    previewRect: SpreadsheetRect;
  } | null = null;
  private readonly columnFilters = new Map<number, string[]>();
  private filterMenuColumn: number | null = null;
  private copiedRange: ReturnType<typeof normalizeSelectionRange> | null = null;
  private formulaReferenceRanges: Array<{ range: ReturnType<typeof normalizeSelectionRange>; stroke: string }> = [];
  private lastStateKey = '';
  private transactionDepth = 0;
  private transactionChanged = false;
  private readOnlyValue: boolean;

  constructor(private readonly host: HTMLElement, private readonly options: SpreadsheetEditorOptions) {
    this.readOnlyValue = options.readOnly === true;
    host.replaceChildren();
    this.renderRoot = options.styleIsolation === 'none'
      ? host
      : (host.shadowRoot ?? host.attachShadow({ mode: 'open' }));
    this.renderRoot.replaceChildren();
    if (this.renderRoot instanceof ShadowRoot) {
      const style = document.createElement('style');
      style.textContent = `
        :host { display: block; position: relative; min-width: 0; min-height: 0; }
        *, *::before, *::after { box-sizing: border-box; }
        button, input, textarea, select { font: 13px Arial, sans-serif; line-height: normal; }
        button { margin: 0; }
      `;
      this.renderRoot.appendChild(style);
    }
    host.classList.add('coco-spreadsheet-editor-host');
    Object.assign(host.style, { position: 'relative', overflow: 'hidden' });
    for (const canvas of [this.baseCanvas, this.headerCanvas, this.overlayCanvas]) {
      Object.assign(canvas.style, { position: 'absolute', inset: '0', display: 'block' });
      this.renderRoot.appendChild(canvas);
    }
    this.unsupportedObjectLayer.className = 'coco-spreadsheet-unsupported-object-layer';
    Object.assign(this.unsupportedObjectLayer.style, {
      position: 'absolute',
      inset: '0',
      zIndex: '4',
      overflow: 'hidden',
      pointerEvents: 'none',
    });
    this.renderRoot.appendChild(this.unsupportedObjectLayer);
    this.cellEditorElement.className = 'coco-spreadsheet-cell-editor';
    Object.assign(this.cellEditorElement.style, {
      position: 'absolute',
      zIndex: '20',
      display: 'none',
      resize: 'none',
      overflow: 'hidden',
      padding: '2px 4px',
      border: '2px solid #2563eb',
      outline: 'none',
      font: '13px Arial, sans-serif',
      lineHeight: 'normal',
      whiteSpace: 'pre-wrap',
      overflowWrap: 'break-word',
    });
    this.renderRoot.appendChild(this.cellEditorElement);
    this.resizeGuide.className = 'coco-spreadsheet-resize-guide';
    Object.assign(this.resizeGuide.style, {
      position: 'absolute',
      zIndex: '25',
      display: 'none',
      pointerEvents: 'none',
      background: '#2563eb',
    });
    this.renderRoot.appendChild(this.resizeGuide);
    this.contextMenu.className = 'coco-spreadsheet-context-menu';
    Object.assign(this.contextMenu.style, {
      position: 'fixed',
      zIndex: '50',
      display: 'none',
      minWidth: '160px',
      padding: '4px',
      background: '#ffffff',
      border: '1px solid #cbd5e1',
      borderRadius: '6px',
      boxShadow: '0 8px 24px rgba(15, 23, 42, 0.16)',
    });
    this.renderRoot.appendChild(this.contextMenu);
    this.filterMenu.className = 'coco-spreadsheet-filter-menu';
    Object.assign(this.filterMenu.style, {
      position: 'fixed', zIndex: '60', display: 'none', width: '240px', maxHeight: '360px',
      overflow: 'auto', padding: '8px', background: '#ffffff', border: '1px solid #cbd5e1',
      borderRadius: '6px', boxShadow: '0 8px 24px rgba(15, 23, 42, 0.16)',
    });
    this.renderRoot.appendChild(this.filterMenu);
    this.cellEditor = new SpreadsheetCellEditor({
      textarea: this.cellEditorElement,
      onCommit: async (value) => this.commitCellEdit(value),
      onCancel: () => {
        this.editingCell = null;
        this.editingCellStyle = null;
        this.formulaReferenceRanges = [];
        this.render();
      },
    });
    this.cellEditorElement.addEventListener('input', () => {
      if (!this.editingCell) return;
      this.formulaReferenceRanges = formulaReferenceRangesFromText(this.cellEditorElement.value);
      this.render();
    });
    this.configureScrollbar(this.horizontalScrollbar, this.horizontalThumb, true);
    this.configureScrollbar(this.verticalScrollbar, this.verticalThumb, false);
    this.renderRoot.append(this.horizontalScrollbar, this.verticalScrollbar);
    this.sheetTabs.className = 'coco-spreadsheet-editor-tabs';
    if (options.showSheetTabs !== false) this.renderRoot.appendChild(this.sheetTabs);
    this.renderer = new SpreadsheetCanvasRenderer({
      baseCanvas: this.baseCanvas,
      headerCanvas: this.headerCanvas,
      overlayCanvas: this.overlayCanvas,
    });
    host.addEventListener('wheel', this.handleWheel, { passive: false });
    host.addEventListener('click', this.handleClick);
    host.addEventListener('dblclick', this.handleDoubleClick);
    host.addEventListener('pointerdown', this.handlePointerDown);
    host.addEventListener('pointermove', this.handlePointerMove);
    host.addEventListener('pointerup', this.handlePointerUp);
    host.addEventListener('pointerleave', this.handlePointerLeave);
    host.addEventListener('keydown', this.handleKeyDown);
    host.addEventListener('copy', this.handleCopy);
    host.addEventListener('paste', this.handlePaste);
    host.addEventListener('contextmenu', this.handleContextMenu);
    document.addEventListener('pointerdown', this.handleDocumentPointerDown);
    this.filterMenu.addEventListener('click', this.handleFilterMenuClick);
    host.tabIndex = 0;
    this.resizeObserver = new ResizeObserver(() => this.render());
    this.resizeObserver.observe(host);
  }

  runtime(): SpreadsheetEditorRuntime {
    if (!this.runtimeValue) throw new Error('SpreadsheetEditor: workbook is not loaded');
    return this.runtimeValue;
  }

  isReadOnly(): boolean {
    return this.readOnlyValue;
  }

  setReadOnly(readOnly: boolean): void {
    if (this.disposed || this.isReadOnly() === readOnly) return;
    this.cellEditor.cancel();
    this.resizeDrag = null;
    this.autoFillDrag = null;
    this.chartDrag = null;
    this.selecting = false;
    this.hideResizeGuide();
    this.hideContextMenu();
    this.hideFilterMenu();
    this.runtime().setReadOnly(readOnly);
    this.readOnlyValue = readOnly;
    this.host.setAttribute('aria-readonly', String(readOnly));
    this.render();
  }

  selection(): SpreadsheetSelectionState {
    return this.selectionState;
  }

  cellState(address: string): SpreadsheetEditorCellState {
    const model = this.runtime().activeSheet();
    const cell = model.cellsByAddress.get(address);
    if (!cell) {
      const match = address.match(/^([A-Z]+)(\d+)$/i);
      if (!match) throw new Error(`SpreadsheetEditor: invalid cell address ${address}`);
      const column = columnNumberFromLetters(match[1]!);
      const row = Number(match[2]);
      return { address: address.toUpperCase(), row, column, value: null, displayText: '' };
    }
    const format = model.styles.cellFormats[cell.styleId ?? 0];
    return {
      address: cell.address,
      row: cell.row,
      column: cell.column,
      value: cell.value,
      displayText: cell.displayText,
      formula: cell.formula,
      styleId: cell.styleId,
      format: format ? {
        font: model.styles.fonts[format.fontId ?? 0],
        fill: model.styles.fills[format.fillId ?? 0],
        border: model.styles.borders[format.borderId ?? 0],
        alignment: format.alignment,
        numberFormatCode: format.numberFormatCode,
      } : undefined,
    };
  }

  freezeState(): SpreadsheetFreezeState {
    return freezeStateFromSheet(this.runtime().activeSheet().source);
  }

  filterState(): ReadonlyMap<number, readonly string[]> {
    return new Map([...this.columnFilters.entries()].map(([column, values]) => [column, [...values] as readonly string[]]));
  }

  activeSheetName(): string {
    return this.runtime().activeSheetName();
  }

  async setActiveSheet(name: string): Promise<void> {
    await this.runtime().setActiveSheet(name);
    this.scroll = { x: 0, y: 0 };
    this.selectionState = createSelection(1, 1);
    this.render();
    this.renderTabs();
  }

  async addSheet(name: string): Promise<void> {
    await this.runtime().addSheet(name);
    this.scroll = { x: 0, y: 0 };
    this.selectionState = createSelection(1, 1);
    this.renderTabs();
    this.render();
    await this.emitChange();
  }

  async renameSheet(sheetName: string, name: string): Promise<void> {
    await this.runtime().renameSheet(sheetName, name);
    this.renderTabs();
    this.render();
    await this.emitChange();
  }

  async copySheet(sheetName: string, newName: string): Promise<void> {
    await this.runtime().copySheet(sheetName, newName);
    this.scroll = { x: 0, y: 0 };
    this.selectionState = createSelection(1, 1);
    this.renderTabs();
    this.render();
    await this.emitChange();
  }

  async deleteSheet(sheetName: string): Promise<void> {
    await this.runtime().deleteSheet(sheetName);
    this.scroll = { x: 0, y: 0 };
    this.selectionState = createSelection(1, 1);
    this.renderTabs();
    this.render();
    await this.emitChange();
  }

  async moveSheet(sheetName: string, toIndex: number): Promise<void> {
    await this.runtime().moveSheet(sheetName, toIndex);
    this.renderTabs();
    this.render();
    await this.emitChange();
  }

  async setCellValue(address: string, value: ExcelCellValue): Promise<void> {
    await this.runtime().setCellValue(address, value, { mergeUndo: true });
    this.render();
    await this.emitChange();
  }

  async setCellFormula(address: string, formula: string): Promise<void> {
    await this.runtime().setCellFormula(address, formula, { mergeUndo: true });
    this.render();
    await this.emitChange();
  }

  async formatSelection(kind: 'bold' | 'italic' | 'underline'): Promise<void> {
    await this.applyFormat(kind);
  }

  async setSelectionFontColor(color: ExcelColorModel | undefined): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    await this.runtime().setRangeFontColor(
      range.startRow,
      range.startColumn,
      range.endRow - range.startRow + 1,
      range.endColumn - range.startColumn + 1,
      color,
    );
    this.render();
    await this.emitChange();
  }

  async setSelectionFillColor(color: ExcelColorModel | undefined): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    await this.runtime().setRangeFillColor(
      range.startRow,
      range.startColumn,
      range.endRow - range.startRow + 1,
      range.endColumn - range.startColumn + 1,
      color,
    );
    this.render();
    await this.emitChange();
  }

  async setSelectionHorizontalAlign(alignment: 'left' | 'center' | 'right'): Promise<void> {
    await this.applyAlignment(alignment);
  }

  async setSelectionVerticalAlign(alignment: 'top' | 'center' | 'bottom'): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    await this.runtime().setRangeVerticalAlign(
      range.startRow,
      range.startColumn,
      range.endRow - range.startRow + 1,
      range.endColumn - range.startColumn + 1,
      alignment,
    );
    this.render();
    await this.emitChange();
  }

  async setSelectionNumberFormat(code: string | undefined): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    await this.runtime().setRangeNumberFormat(
      range.startRow,
      range.startColumn,
      range.endRow - range.startRow + 1,
      range.endColumn - range.startColumn + 1,
      code,
    );
    this.render();
    await this.emitChange();
  }

  async setSelectionBorder(mode: 'all' | 'outer' | 'inner' | 'clear'): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    await this.runtime().setRangeBorder(
      range.startRow,
      range.startColumn,
      range.endRow - range.startRow + 1,
      range.endColumn - range.startColumn + 1,
      mode,
    );
    this.render();
    await this.emitChange();
  }

  async loadWorkbook(bytes: Uint8Array | ArrayBuffer): Promise<void> {
    this.assertAlive();
    const previous = this.runtimeValue;
    this.runtimeValue = null;
    this.scroll = { x: 0, y: 0 };
    if (previous) await previous.dispose();
    this.runtimeValue = await SpreadsheetEditorRuntime.fromZip(bytes, {
      workbookId: this.options.workbookId,
      mode: 'inline',
      readOnly: this.readOnlyValue,
    });
    this.host.setAttribute('aria-readonly', String(this.isReadOnly()));
    this.renderTabs();
    this.render();
  }

  async toZip(): Promise<Uint8Array> {
    return this.runtime().toZip();
  }

  canUndo(): boolean { return this.runtime().canUndo(); }
  canRedo(): boolean { return this.runtime().canRedo(); }

  async undo(): Promise<void> {
    await this.runtime().undo();
    this.render();
    await this.emitChange();
  }

  async redo(): Promise<void> {
    await this.runtime().redo();
    this.render();
    await this.emitChange();
  }

  subscribe(listener: (event: SpreadsheetEditorChangeEvent) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  subscribeState(listener: (event: SpreadsheetEditorStateEvent) => void): () => void {
    this.stateListeners.add(listener);
    return () => this.stateListeners.delete(listener);
  }

  async runTransaction(
    options: SpreadsheetUserTransactionOptions,
    run: (editor: SpreadsheetEditor) => Promise<void>,
  ): Promise<void> {
    this.transactionDepth += 1;
    try {
      await this.runtime().runUserTransaction(options, () => run(this));
    } finally {
      this.transactionDepth -= 1;
      if (this.transactionDepth === 0 && this.transactionChanged) {
        this.transactionChanged = false;
        this.render();
        await this.emitChange();
      }
    }
  }

  destroy(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.resizeObserver.disconnect();
    this.host.removeEventListener('wheel', this.handleWheel);
    this.host.removeEventListener('click', this.handleClick);
    this.host.removeEventListener('dblclick', this.handleDoubleClick);
    this.host.removeEventListener('pointerdown', this.handlePointerDown);
    this.host.removeEventListener('pointermove', this.handlePointerMove);
    this.host.removeEventListener('pointerup', this.handlePointerUp);
    this.host.removeEventListener('pointerleave', this.handlePointerLeave);
    this.host.removeEventListener('keydown', this.handleKeyDown);
    this.host.removeEventListener('copy', this.handleCopy);
    this.host.removeEventListener('paste', this.handlePaste);
    this.host.removeEventListener('contextmenu', this.handleContextMenu);
    document.removeEventListener('pointerdown', this.handleDocumentPointerDown);
    this.contextMenu.remove();
    this.filterMenu.remove();
    this.unsupportedObjectLayer.remove();
    this.host.replaceChildren();
    const runtime = this.runtimeValue;
    this.runtimeValue = null;
    void runtime?.dispose();
  }

  private render(): void {
    const runtime = this.runtimeValue;
    const renderer = this.renderer;
    if (!runtime || !renderer || this.disposed) return;
    const rect = this.host.getBoundingClientRect();
    const width = Math.max(1, rect.width);
    const height = Math.max(1, rect.height - (this.options.showSheetTabs === false ? 0 : 40));
    const layout = new SpreadsheetLayout(runtime.activeSheet());
    const viewportSize = { width: Math.max(1, width - 12), height: Math.max(1, height - 12) };
    const viewport = new SpreadsheetViewportManager(layout, {
      viewportSize,
      scroll: this.scroll,
      freeze: freezeStateFromSheet(runtime.activeSheet().source),
    });
    this.layout = layout;
    this.viewport = viewport;
    this.scroll = viewport.scroll;
    this.scrollbarState = computeScrollbarState(
      layout,
      viewportSize,
      freezeStateFromSheet(runtime.activeSheet().source),
      this.scroll,
    );
    renderer.resize(width, height);
    const headerFiltersEnabled = this.options.enableHeaderFilters === true && !this.isReadOnly();
    renderer.setHeaderFiltersEnabled(headerFiltersEnabled);
    renderer.setHeaderFilterColumns(headerFiltersEnabled
      ? Array.from({ length: layout.columnCount }, (_, index) => index + 1)
      : null);
    renderer.setActiveHeaderFilterColumns(Array.from(this.columnFilters.keys()));
    renderer.render(layout, viewport);
    const selectionRange = normalizeSelectionRange(this.selectionState.range);
    renderer.renderSelection(layout, viewport, selectionRange, {
      showFillHandle: !this.isReadOnly() && supportsSpreadsheetAutoFill(runtime.activeSheet(), selectionRange),
    });
    this.updateCellEditorPosition();
    if (this.autoFillDrag?.preview) {
      renderer.renderDashedRange(layout, viewport, this.autoFillDrag.preview, {
        stroke: '#16A34A',
        fill: 'rgba(34, 197, 94, 0.08)',
      });
    }
    for (const reference of this.formulaReferenceRanges) {
      renderer.renderDashedRange(layout, viewport, reference.range, { stroke: reference.stroke });
    }
    if (this.copiedRange) renderer.renderCopiedRange(layout, viewport, this.copiedRange);
    this.renderUnsupportedChartNotices(layout, viewport);
    this.emitStateIfChanged(runtime.activeSheetName());
    this.updateScrollbars();
  }

  private renderUnsupportedChartNotices(
    layout: SpreadsheetLayout,
    viewport: SpreadsheetViewportManager,
  ): void {
    this.unsupportedObjectLayer.replaceChildren();
    const sheet = this.runtimeValue?.activeSheet().source;
    if (!sheet || sheet.charts.length === 0) return;

    for (let chartIndex = 0; chartIndex < sheet.charts.length; chartIndex += 1) {
      const chart = sheet.charts[chartIndex];
      if (!chart) continue;
      const anchor = findChartAnchor(sheet.drawings, chart);
      const logicalRect = this.chartDrag?.chartIndex === chartIndex
        ? this.chartDrag.previewRect
        : chartLogicalRect(layout, anchor);
      if (!logicalRect) continue;

      for (const fragment of chartScreenFragments(logicalRect, viewport)) {
        const notice = document.createElement('div');
        notice.className = 'coco-spreadsheet-unsupported-chart';
        Object.assign(notice.style, {
          position: 'absolute',
          left: `${fragment.x}px`,
          top: `${fragment.y}px`,
          width: `${Math.max(1, fragment.width)}px`,
          height: `${Math.max(1, fragment.height)}px`,
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          gap: '6px',
          padding: '12px',
          boxSizing: 'border-box',
          overflow: 'hidden',
          background: 'rgba(248, 250, 252, 0.94)',
          color: '#475569',
          textAlign: 'center',
          font: '13px Arial, sans-serif',
          lineHeight: '1.4',
          cursor: this.isReadOnly() ? 'default' : this.chartDrag?.chartIndex === chartIndex ? 'grabbing' : 'move',
          border: this.selectedChartIndex === chartIndex ? '2px solid #2563eb' : '1px dashed #94a3b8',
          boxShadow: this.selectedChartIndex === chartIndex ? '0 0 0 1px rgba(37, 99, 235, 0.18)' : 'none',
        });
        const title = document.createElement('strong');
        title.textContent = chart.title ? `图表：${chart.title}` : '图表';
        title.style.color = '#334155';
        const message = document.createElement('span');
        message.textContent = '当前编辑器暂不支持图表渲染';
        const suggestion = document.createElement('span');
        suggestion.textContent = '请使用 Excel 或 WPS 打开文件查看';
        notice.append(title, message, suggestion);
        this.unsupportedObjectLayer.appendChild(notice);
      }
    }
  }

  private readonly handleClick = (event: MouseEvent): void => {
    if (this.eventPathContains(event, this.contextMenu) || this.eventPathContains(event, this.filterMenu)) return;
    if (this.filterMenu.style.display === 'block') this.hideFilterMenu();
    if (this.editingCell) {
      if (this.eventPathContains(event, this.cellEditorElement)) return;
      this.cellEditor.cancel();
    }
    if (!this.viewport || !this.layout) return;
    const chartHit = this.chartAtPoint(this.eventPoint(event));
    if (chartHit) {
      this.selectedChartIndex = chartHit.chartIndex;
      this.render();
      event.preventDefault();
      return;
    }
    this.selectedChartIndex = null;
    const hit = this.viewport.hitTest(this.eventPoint(event));
    if (hit.kind !== 'cell' && hit.kind !== 'rowHeader' && hit.kind !== 'columnHeader') return;
    if (hit.kind === 'cell' && hit.row === 1 && this.isFilterIconHit(hit.column, event)) {
      if (this.isReadOnly()) return;
      this.openFilterMenu(hit.column, event.clientX, event.clientY);
      event.preventDefault();
      return;
    }
    if (hit.kind === 'cell') {
      const model = this.runtime().activeSheet();
      this.selectionState = expandSelectionForMergedRanges(
        createSelection(hit.row, hit.column),
        model,
      );
    }
    this.host.focus();
    this.render();
  };

  private readonly handleFilterMenuClick = (event: MouseEvent): void => {
    if (this.isReadOnly()) {
      this.hideFilterMenu();
      return;
    }
    const target = event.target as HTMLElement;
    const action = target.closest<HTMLButtonElement>('[data-filter-action]')?.dataset.filterAction;
    if (!action || this.filterMenuColumn === null) return;
    if (action === 'cancel') {
      this.hideFilterMenu();
      return;
    }
    const column = this.filterMenuColumn;
    if (action === 'clear') {
      this.columnFilters.delete(column);
    } else if (action === 'apply') {
      const values = Array.from(this.filterMenu.querySelectorAll<HTMLInputElement>('input[data-filter-value]:checked'))
        .map((input) => input.dataset.filterValue ?? '');
      this.columnFilters.set(column, values);
    }
    this.hideFilterMenu();
    void this.applyColumnFilters();
  };

  private readonly handleDoubleClick = (event: MouseEvent): void => {
    if (!this.viewport || !this.layout) return;
    const hit = this.viewport.hitTest(this.eventPoint(event));
    if (hit.kind !== 'cell') return;
    this.selectionState = expandSelectionForMergedRanges(createSelection(hit.row, hit.column), this.runtime().activeSheet());
    if (this.isReadOnly()) {
      this.render();
      return;
    }
    this.openCellEditor();
    event.preventDefault();
  };

  private readonly handlePointerDown = (event: PointerEvent): void => {
    if (this.editingCell) {
      if (this.eventPathContains(event, this.cellEditorElement)) return;
      this.cellEditor.cancel();
    }
    if (event.button !== 0 || !this.viewport || !this.layout) return;
    const point = this.eventPoint(event);
    const chartHit = this.chartAtPoint(point);
    if (chartHit) {
      this.selectedChartIndex = chartHit.chartIndex;
      if (this.isReadOnly()) {
        this.render();
        event.preventDefault();
        return;
      }
      this.chartDrag = {
        chartIndex: chartHit.chartIndex,
        pointer: point,
        startRect: chartHit.logicalRect,
        previewRect: chartHit.logicalRect,
      };
      this.host.setPointerCapture(event.pointerId);
      this.render();
      event.preventDefault();
      return;
    }
    this.selectedChartIndex = null;
    const resizeHandle = hitSpreadsheetResizeHandle(this.layout, this.viewport.viewports, point);
    if (resizeHandle && !this.isReadOnly()) {
      this.resizeDrag = startSpreadsheetResizeDrag(resizeHandle, event);
      this.host.setPointerCapture(event.pointerId);
      this.host.style.cursor = resizeHandle.kind === 'column' ? 'col-resize' : 'row-resize';
      this.updateResizeGuide(this.resizeDrag);
      event.preventDefault();
      return;
    }
    const selectionRange = normalizeSelectionRange(this.selectionState.range);
    const fillHandle = this.renderer?.fillHandleRect(this.layout, this.viewport, selectionRange);
    if (!this.isReadOnly() && fillHandle && containsPoint(fillHandle, point) && supportsSpreadsheetAutoFill(this.runtime().activeSheet(), selectionRange)) {
      this.autoFillDrag = { source: selectionRange, preview: null };
      this.host.setPointerCapture(event.pointerId);
      event.preventDefault();
      return;
    }
    const hit = this.viewport.hitTest(point);
    if (hit.kind !== 'cell' && hit.kind !== 'rowHeader' && hit.kind !== 'columnHeader') return;
    const model = this.runtime().activeSheet();
    if (hit.kind === 'cell') {
      this.selectionState = expandSelectionForMergedRanges(createSelection(hit.row, hit.column), model);
    } else if (hit.kind === 'rowHeader') {
      this.selectionState = createRowSelection(hit.row, { columnCount: this.layout.columnCount });
    } else {
      this.selectionState = createColumnSelection(hit.column, { rowCount: this.layout.rowCount });
    }
    this.selecting = true;
    this.host.setPointerCapture(event.pointerId);
    this.render();
    event.preventDefault();
  };

  private readonly handlePointerMove = (event: PointerEvent): void => {
    if (!this.viewport || !this.layout) return;
    if (this.chartDrag) {
      const point = this.eventPoint(event);
      const deltaX = point.x - this.chartDrag.pointer.x;
      const deltaY = point.y - this.chartDrag.pointer.y;
      this.chartDrag.previewRect = {
        ...this.chartDrag.startRect,
        x: Math.max(0, this.chartDrag.startRect.x + deltaX),
        y: Math.max(0, this.chartDrag.startRect.y + deltaY),
      };
      this.host.style.cursor = 'grabbing';
      this.render();
      event.preventDefault();
      return;
    }
    if (!this.resizeDrag && !this.autoFillDrag && !this.selecting && !this.editingCell) {
      const resizeHandle = this.isReadOnly()
        ? null
        : hitSpreadsheetResizeHandle(this.layout, this.viewport.viewports, this.eventPoint(event));
      this.host.style.cursor = resizeHandle
        ? resizeHandle.kind === 'column' ? 'col-resize' : 'row-resize'
        : '';
    }
    if (this.resizeDrag) {
      this.resizeDrag = updateSpreadsheetResizeDrag(this.resizeDrag, event);
      this.updateResizeGuide(this.resizeDrag);
      return;
    }
    if (this.autoFillDrag) {
      const hit = this.viewport.hitTest(this.eventPoint(event));
      this.autoFillDrag.preview = autoFillPreviewRangeForHit(this.autoFillDrag.source, hit);
      this.render();
      event.preventDefault();
      return;
    }
    if (!this.selecting) return;
    const hit = this.viewport.hitTest(this.eventPoint(event));
    if (hit.kind === 'none' || hit.kind === 'corner') return;
    const model = this.runtime().activeSheet();
    if (hit.kind === 'cell') {
      this.selectionState = expandSelectionForMergedRanges(
        extendSelection(this.selectionState, hit.row, hit.column, {
          rowCount: this.layout.rowCount,
          columnCount: this.layout.columnCount,
        }),
        model,
      );
    } else if (hit.kind === 'rowHeader' && this.selectionState.kind === 'row') {
      this.selectionState = extendSelection(this.selectionState, hit.row, 1, { columnCount: this.layout.columnCount, rowCount: this.layout.rowCount });
    } else if (hit.kind === 'columnHeader' && this.selectionState.kind === 'column') {
      this.selectionState = extendSelection(this.selectionState, 1, hit.column, { columnCount: this.layout.columnCount, rowCount: this.layout.rowCount });
    }
    this.render();
    event.preventDefault();
  };

  private readonly handlePointerUp = (event: PointerEvent): void => {
    if (this.chartDrag) {
      const completed = this.chartDrag;
      this.chartDrag = null;
      this.host.style.cursor = '';
      if (this.host.hasPointerCapture(event.pointerId)) this.host.releasePointerCapture(event.pointerId);
      if (completed.previewRect.x === completed.startRect.x && completed.previewRect.y === completed.startRect.y) {
        this.render();
        event.preventDefault();
        return;
      }
      const anchor = this.chartAnchorAtIndex(completed.chartIndex);
      if (anchor && this.layout) {
        const next = chartAnchorForLogicalRect(this.layout, completed.previewRect, anchor);
        void this.runtime().moveChart(completed.chartIndex, next).then(async () => {
          this.render();
          await this.emitChange();
        });
      } else {
        this.render();
      }
      event.preventDefault();
      return;
    }
    if (this.resizeDrag) {
      const completed = updateSpreadsheetResizeDrag(this.resizeDrag, event);
      this.resizeDrag = null;
      this.host.style.cursor = '';
      this.hideResizeGuide();
      if (this.host.hasPointerCapture(event.pointerId)) this.host.releasePointerCapture(event.pointerId);
      void this.commitResizeDrag(completed);
      event.preventDefault();
      return;
    }
    if (this.autoFillDrag) {
      const completed = this.autoFillDrag;
      this.autoFillDrag = null;
      if (this.host.hasPointerCapture(event.pointerId)) this.host.releasePointerCapture(event.pointerId);
      if (completed.preview) {
        void this.runtime().autoFillRange(completed.source, completed.preview).then(async () => {
          this.selectionState = { ...this.selectionState, range: completed.preview! };
          this.render();
          await this.emitChange();
        });
      } else {
        this.render();
      }
      event.preventDefault();
      return;
    }
    if (!this.selecting) return;
    this.selecting = false;
    if (this.host.hasPointerCapture(event.pointerId)) this.host.releasePointerCapture(event.pointerId);
    this.host.style.cursor = '';
  };

  private readonly handlePointerLeave = (): void => {
    if (!this.resizeDrag && !this.autoFillDrag && !this.chartDrag) this.host.style.cursor = '';
  };

  private chartAnchorAtIndex(chartIndex: number): ExcelDrawingAnchorModel | undefined {
    const sheet = this.runtimeValue?.activeSheet().source;
    const chart = sheet?.charts[chartIndex];
    return chart ? findChartAnchor(sheet.drawings, chart) : undefined;
  }

  private chartAtPoint(point: { x: number; y: number }): {
    chartIndex: number;
    logicalRect: SpreadsheetRect;
  } | null {
    const sheet = this.runtimeValue?.activeSheet().source;
    if (!sheet || !this.layout || !this.viewport) return null;
    for (let chartIndex = sheet.charts.length - 1; chartIndex >= 0; chartIndex -= 1) {
      const chart = sheet.charts[chartIndex];
      if (!chart) continue;
      const anchor = findChartAnchor(sheet.drawings, chart);
      const logicalRect = chartLogicalRect(this.layout, anchor);
      if (!logicalRect) continue;
      if (chartScreenFragments(logicalRect, this.viewport).some((fragment) => containsPoint(fragment, point))) {
        return { chartIndex, logicalRect };
      }
    }
    return null;
  }

  private updateResizeGuide(drag: SpreadsheetResizeDragState): void {
    if (drag.kind === 'column') {
      Object.assign(this.resizeGuide.style, {
        display: 'block',
        left: `${Math.round(drag.currentGuidePosition)}px`,
        top: '0',
        bottom: '0',
        width: '2px',
        height: '',
      });
    } else {
      Object.assign(this.resizeGuide.style, {
        display: 'block',
        left: '0',
        top: `${Math.round(drag.currentGuidePosition)}px`,
        right: '0',
        height: '2px',
        width: '',
      });
    }
  }

  private hideResizeGuide(): void {
    this.resizeGuide.style.display = 'none';
  }

  private readonly handleKeyDown = (event: KeyboardEvent): void => {
    if (this.editingCell) return;
    const { row, column } = this.selectionState.active;
    if (event.key === 'F2' || event.key === 'Enter') {
      if (this.isReadOnly()) return;
      event.preventDefault();
      this.openCellEditor();
      return;
    }
    const deltas: Record<string, [number, number]> = {
      ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1],
    };
    const delta = deltas[event.key];
    if (!delta) return;
    const model = this.runtime().activeSheet();
    this.selectionState = createSelection(
      Math.max(1, Math.min(model.maxRow, row + delta[0])),
      Math.max(1, Math.min(model.maxColumn, column + delta[1])),
    );
    event.preventDefault();
    this.render();
  };

  private readonly handleCopy = (event: ClipboardEvent): void => {
    const range = normalizeSelectionRange(this.selectionState.range);
    const payload = buildSpreadsheetClipboardPayload(this.runtime().activeSheet(), range);
    this.clipboardPayload = payload;
    this.copiedRange = range;
    event.clipboardData?.setData('text/plain', payload.text);
    event.preventDefault();
  };

  private readonly handlePaste = (event: ClipboardEvent): void => {
    if (this.isReadOnly()) {
      event.preventDefault();
      return;
    }
    const text = event.clipboardData?.getData('text/plain') ?? '';
    if (!text) return;
    event.preventDefault();
    const values = this.clipboardPayload?.text === text ? this.clipboardPayload.values : parseTsv(text);
    void this.runtime().pasteRange(this.selectionState.active.row, this.selectionState.active.column, values, {
      mode: 'all',
      merges: this.clipboardPayload?.text === text ? this.clipboardPayload.merges : [],
      styleIds: this.clipboardPayload?.text === text ? this.clipboardPayload.styleIds : [],
    }).then(async () => {
      const pastedSelection = expandSelectionForMergedRanges({
        kind: 'cell',
        active: this.selectionState.active,
        range: {
          startRow: this.selectionState.active.row,
          startColumn: this.selectionState.active.column,
          endRow: this.selectionState.active.row + values.length - 1,
          endColumn: this.selectionState.active.column + Math.max(1, values.reduce((max, row) => Math.max(max, row.length), 0)) - 1,
        },
      }, this.runtime().activeSheet());
      this.selectionState = pastedSelection;
      this.copiedRange = normalizeSelectionRange(pastedSelection.range);
      this.render();
      await this.emitChange();
    });
  };

  private readonly handleContextMenu = (event: MouseEvent): void => {
    if (!this.viewport || !this.layout) return;
    const chartHit = this.chartAtPoint(this.eventPoint(event));
    if (chartHit) {
      this.selectedChartIndex = chartHit.chartIndex;
      this.hideContextMenu();
      this.hideFilterMenu();
      this.render();
      event.preventDefault();
      return;
    }
    const hit = this.viewport.hitTest(this.eventPoint(event));
    if (hit.kind === 'cell' && !this.selectionContainsCell(hit.row, hit.column)) {
      this.selectionState = expandSelectionForMergedRanges(createSelection(hit.row, hit.column), this.runtime().activeSheet());
      this.render();
    }
    if (hit.kind !== 'cell' && hit.kind !== 'rowHeader' && hit.kind !== 'columnHeader') return;
    event.preventDefault();
    this.renderContextMenu(
      event.clientX,
      event.clientY,
      this.isReadOnly() ? undefined : hit.kind === 'columnHeader' ? hit.column : undefined,
    );
  };

  private readonly handleDocumentPointerDown = (event: PointerEvent): void => {
    const path = event.composedPath();
    if (path.includes(this.contextMenu) || path.includes(this.filterMenu)) return;
    this.hideContextMenu();
    this.hideFilterMenu();
  };

  private renderContextMenu(x: number, y: number, columnHeader?: number): void {
    const items: Array<{ label: string; disabled?: boolean; action: () => void | Promise<void> }> =
      this.isReadOnly()
        ? [{ label: 'Copy', action: () => this.copySelection() }]
        : [
          { label: 'Undo', disabled: !this.canUndo(), action: () => this.undo() },
          { label: 'Redo', disabled: !this.canRedo(), action: () => this.redo() },
          { label: 'Copy', action: () => this.copySelection() },
          { label: 'Paste', action: () => this.pasteFromClipboard() },
          { label: 'Clear cells', action: () => this.clearSelection() },
          { label: 'Clear format', action: () => this.clearSelectionFormat() },
          { label: 'Bold', action: () => this.applyFormat('bold') },
          { label: 'Italic', action: () => this.applyFormat('italic') },
          { label: 'Underline', action: () => this.applyFormat('underline') },
          { label: 'Align left', action: () => this.applyAlignment('left') },
          { label: 'Align center', action: () => this.applyAlignment('center') },
          { label: 'Align right', action: () => this.applyAlignment('right') },
          { label: 'Merge cells', action: () => this.mergeSelection() },
          { label: 'Unmerge cells', action: () => this.unmergeSelection() },
          { label: 'Freeze above and left', action: () => this.freezeSelection() },
        ];
    if (columnHeader !== undefined) {
      items.unshift({
        label: 'Filter column',
        action: () => this.openFilterMenu(columnHeader, x, y),
      });
      items.splice(1, 0,
        { label: 'Sort A-Z', action: () => this.sortColumn(columnHeader, 'asc') },
        { label: 'Sort Z-A', action: () => this.sortColumn(columnHeader, 'desc') },
      );
    }
    this.contextMenu.replaceChildren();
    for (const item of items) {
      const button = document.createElement('button');
      button.type = 'button';
      button.textContent = item.label;
      button.disabled = item.disabled === true;
      Object.assign(button.style, {
        display: 'block', width: '100%', padding: '6px 10px', border: '0',
        background: 'transparent', textAlign: 'left', cursor: button.disabled ? 'default' : 'pointer',
      });
      button.addEventListener('mouseenter', () => { if (!button.disabled) button.style.background = '#eff6ff'; });
      button.addEventListener('mouseleave', () => { button.style.background = 'transparent'; });
      button.addEventListener('click', () => {
        this.hideContextMenu();
        void item.action();
      });
      this.contextMenu.appendChild(button);
    }
    Object.assign(this.contextMenu.style, { left: `${x}px`, top: `${y}px`, display: 'block' });
  }

  private openFilterMenu(column: number, x: number, y: number): void {
    if (this.filterMenu.style.display === 'block' && this.filterMenuColumn === column) {
      this.hideFilterMenu();
      return;
    }
    const model = this.runtime().activeSheet();
    const range = { startRow: 1, startColumn: 1, endRow: model.maxRow, endColumn: model.maxColumn };
    const selected = new Set(this.columnFilters.get(column) ?? spreadsheetFilterValues(model, range, column));
    this.filterMenuColumn = column;
    this.filterMenu.replaceChildren();
    const title = document.createElement('strong');
    title.textContent = `Filter ${column}`;
    title.style.display = 'block';
    title.style.marginBottom = '6px';
    this.filterMenu.appendChild(title);
    for (const value of spreadsheetFilterValues(model, range, column)) {
      const label = document.createElement('label');
      label.style.display = 'block';
      label.style.padding = '3px 0';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = selected.has(value);
      input.dataset.filterValue = value;
      input.style.marginRight = '8px';
      label.append(input, value === '' ? '(blank)' : value);
      this.filterMenu.appendChild(label);
    }
    const actions = document.createElement('div');
    actions.style.display = 'flex';
    actions.style.gap = '6px';
    actions.style.marginTop = '8px';
    for (const [label, action] of [['Apply', 'apply'], ['Clear filter', 'clear'], ['Cancel', 'cancel']] as const) {
      const button = document.createElement('button');
      button.type = 'button';
      button.textContent = label;
      button.dataset.filterAction = action;
      actions.appendChild(button);
    }
    this.filterMenu.appendChild(actions);
    Object.assign(this.filterMenu.style, { left: `${x}px`, top: `${y}px`, display: 'block' });
  }

  private hideFilterMenu(): void {
    this.filterMenu.style.display = 'none';
    this.filterMenuColumn = null;
  }

  private isFilterIconHit(column: number, event: MouseEvent): boolean {
    if (!this.layout || !this.viewport) return false;
    const point = this.eventPoint(event);
    const viewport = this.viewport.findViewport(point);
    if (!viewport) return false;
    const rect = this.viewport.cellScreenRect(1, column, viewport);
    return rect.width >= 28 && point.x >= rect.x + rect.width - 22 && point.x <= rect.x + rect.width;
  }

  private async applyColumnFilters(): Promise<void> {
    const model = this.runtime().activeSheet();
    const filters: SpreadsheetColumnFilter[] = Array.from(this.columnFilters.entries())
      .filter(([, values]) => values.length > 0)
      .map(([column, values]) => ({ column, values }));
    await this.runtime().filterRowsByColumns(1, model.maxRow, filters);
    this.render();
    await this.emitChange();
  }

  private async sortColumn(column: number, direction: 'asc' | 'desc'): Promise<void> {
    const model = this.runtime().activeSheet();
    await this.runtime().sortRows(1, model.maxRow, column, direction);
    this.render();
    await this.emitChange();
  }

  async sortActiveColumn(direction: 'asc' | 'desc'): Promise<void> {
    await this.sortColumn(this.selectionState.active.column, direction);
  }

  async filterColumn(column: number, values: readonly string[] | null): Promise<void> {
    this.assertWritable();
    if (values === null) this.columnFilters.delete(column);
    else this.columnFilters.set(column, [...values]);
    await this.applyColumnFilters();
  }

  private hideContextMenu(): void {
    this.contextMenu.style.display = 'none';
  }

  private selectionContainsCell(row: number, column: number): boolean {
    const range = normalizeSelectionRange(this.selectionState.range);
    return row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn;
  }

  private async copySelection(): Promise<void> {
    const payload = buildSpreadsheetClipboardPayload(this.runtime().activeSheet(), normalizeSelectionRange(this.selectionState.range));
    this.clipboardPayload = payload;
    try { await navigator.clipboard.writeText(payload.text); } catch { /* Clipboard events still support internal copy. */ }
  }

  private async pasteFromClipboard(): Promise<void> {
    let text = '';
    try { text = await navigator.clipboard.readText(); } catch { return; }
    if (!text) return;
    const event = new ClipboardEvent('paste', { clipboardData: new DataTransfer() });
    event.clipboardData?.setData('text/plain', text);
    this.handlePaste(event);
  }

  async clearSelection(): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    await this.runtime().clearRange(range.startRow, range.startColumn, range.endRow - range.startRow + 1, range.endColumn - range.startColumn + 1);
    this.render();
    await this.emitChange();
  }

  async clearSelectionFormat(): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    await this.runtime().clearRangeFormat(range.startRow, range.startColumn, range.endRow - range.startRow + 1, range.endColumn - range.startColumn + 1);
    this.render();
    await this.emitChange();
  }

  private async applyFormat(kind: 'bold' | 'italic' | 'underline'): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    const rows = range.endRow - range.startRow + 1;
    const columns = range.endColumn - range.startColumn + 1;
    if (kind === 'bold') await this.runtime().setRangeBold(range.startRow, range.startColumn, rows, columns, true);
    if (kind === 'italic') await this.runtime().setRangeItalic(range.startRow, range.startColumn, rows, columns, true);
    if (kind === 'underline') await this.runtime().setRangeUnderline(range.startRow, range.startColumn, rows, columns, true);
    this.render();
    await this.emitChange();
  }

  private async applyAlignment(alignment: 'left' | 'center' | 'right'): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    await this.runtime().setRangeHorizontalAlign(
      range.startRow,
      range.startColumn,
      range.endRow - range.startRow + 1,
      range.endColumn - range.startColumn + 1,
      alignment,
    );
    this.render();
    await this.emitChange();
  }

  async mergeSelection(): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    if (range.startRow === range.endRow && range.startColumn === range.endColumn) return;
    await this.runtime().mergeRange(
      range.startRow,
      range.startColumn,
      range.endRow - range.startRow + 1,
      range.endColumn - range.startColumn + 1,
    );
    this.selectionState = expandSelectionForMergedRanges(this.selectionState, this.runtime().activeSheet());
    this.render();
    await this.emitChange();
  }

  async unmergeSelection(): Promise<void> {
    const range = normalizeSelectionRange(this.selectionState.range);
    const addresses = [...this.runtime().activeSheet().mergeMasterByAddress.values()]
      .filter((merge) => rangesIntersect(range, merge.range))
      .map((merge) => merge.address);
    if (addresses.length === 0) return;
    await this.runtime().unmergeRanges(addresses);
    this.render();
    await this.emitChange();
  }

  async freezeSelection(): Promise<void> {
    const active = this.selectionState.active;
    await this.runtime().setFreezePanes(Math.max(0, active.row - 1), Math.max(0, active.column - 1));
    this.render();
    await this.emitChange();
  }

  private async commitResizeDrag(drag: SpreadsheetResizeDragState): Promise<void> {
    if (drag.currentSize === drag.startSize) {
      this.render();
      return;
    }
    if (drag.kind === 'column') await this.runtime().setColumnWidth(drag.index, drag.currentSize);
    else await this.runtime().setRowHeight(drag.index, drag.currentSize * 72 / 96);
    this.render();
    await this.emitChange();
  }

  private openCellEditor(): void {
    if (this.isReadOnly()) return;
    if (!this.layout || !this.viewport) return;
    const cell = this.selectionState.active;
    const viewport = this.viewport.findViewport({
      x: this.layout.rowHeaderWidth + 1,
      y: this.layout.columnHeaderHeight + 1,
    });
    const renderCell = getRenderCell(this.runtime().activeSheet(), cell.row, cell.column);
    const cellStyle = resolveSpreadsheetCellStyle(this.runtime().activeSheet().styles, renderCell?.styleId, renderCell);
    Object.assign(this.cellEditorElement.style, {
      font: cellStyle.font,
      color: cellStyle.fontColor,
      textDecoration: cellStyle.underline ? 'underline' : 'none',
      textAlign: cellStyle.horizontalAlign,
    });
    const targetViewport = this.viewport.viewports.find((item) => item.kind === 'main') ?? viewport;
    if (!targetViewport) return;
    const editRange = normalizeSelectionRange(this.selectionState.range);
    const startRect = this.viewport.cellScreenRect(editRange.startRow, editRange.startColumn, targetViewport);
    const endRect = this.viewport.cellScreenRect(editRange.endRow, editRange.endColumn, targetViewport);
    const rect = {
      x: startRect.x,
      y: startRect.y,
      width: endRect.x + endRect.width - startRect.x,
      height: endRect.y + endRect.height - startRect.y,
    };
    this.editingCell = cell;
    this.editingCellStyle = cellStyle;
    const value = renderCell?.formula ? `=${renderCell.formula}` : renderCell?.displayText ?? '';
    this.formulaReferenceRanges = formulaReferenceRangesFromText(value);
    this.cellEditor.open(rect, value);
    this.alignCellEditor(rect, cellStyle);
    this.render();
  }

  private async commitCellEdit(value: string): Promise<void> {
    const cell = this.editingCell;
    this.editingCell = null;
    this.editingCellStyle = null;
    this.formulaReferenceRanges = [];
    if (!cell) return;
    this.assertWritable();
    const address = formatCellAddress(cell.row, cell.column);
    if (value.startsWith('=')) await this.runtime().setCellFormula(address, value.slice(1), { mergeUndo: true });
    else await this.runtime().setCellValue(address, coerceEditorValue(value), { mergeUndo: true });
    this.render();
    await this.emitChange();
  }

  private eventPoint(event: MouseEvent): { x: number; y: number } {
    const rect = this.host.getBoundingClientRect();
    return { x: event.clientX - rect.left, y: event.clientY - rect.top };
  }

  private assertWritable(): void {
    if (this.isReadOnly()) throw new EditorReadOnlyError('spreadsheet');
  }

  private eventPathContains(event: Event, node: EventTarget): boolean {
    return event.composedPath().includes(node);
  }

  private updateCellEditorPosition(): void {
    const cell = this.editingCell;
    const layout = this.layout;
    const viewportManager = this.viewport;
    if (!cell || !layout || !viewportManager) return;
    const targetViewport = viewportManager.viewports.find((viewport) =>
      viewport.kind !== 'columnHeaderFrozen' &&
      viewport.kind !== 'columnHeaderScrollable' &&
      viewport.kind !== 'rowHeaderFrozen' &&
      viewport.kind !== 'rowHeaderScrollable' &&
      cell.row >= viewport.rows.start && cell.row <= viewport.rows.end &&
      cell.column >= viewport.columns.start && cell.column <= viewport.columns.end,
    );
    if (!targetViewport) {
      this.cellEditorElement.style.display = 'none';
      return;
    }
    const editRange = normalizeSelectionRange(this.selectionState.range);
    const startRect = viewportManager.cellScreenRect(editRange.startRow, editRange.startColumn, targetViewport);
    const endRect = viewportManager.cellScreenRect(editRange.endRow, editRange.endColumn, targetViewport);
    const rect = {
      x: startRect.x,
      y: startRect.y,
      width: endRect.x + endRect.width - startRect.x,
      height: endRect.y + endRect.height - startRect.y,
    };
    this.cellEditorElement.style.left = `${rect.x}px`;
    this.cellEditorElement.style.top = `${rect.y}px`;
    this.cellEditorElement.style.width = `${Math.max(1, rect.width)}px`;
    this.cellEditorElement.style.height = `${Math.max(1, rect.height)}px`;
    this.cellEditorElement.style.display = 'block';
    if (this.editingCellStyle) this.alignCellEditor(rect, this.editingCellStyle);
  }

  private alignCellEditor(rect: { height: number }, style: SpreadsheetResolvedCellStyle): void {
    const textarea = this.cellEditorElement;
    const paddingX = 4;
    const paddingY = 2;
    textarea.style.padding = `${paddingY}px ${paddingX}px`;
    const extraHeight = Math.max(0, rect.height - textarea.scrollHeight);
    const topPadding =
      style.verticalAlign === 'bottom'
        ? paddingY + extraHeight
        : style.verticalAlign === 'middle'
          ? paddingY + extraHeight / 2
          : paddingY;
    textarea.style.paddingTop = `${topPadding}px`;
  }

  private async emitChange(): Promise<void> {
    if (this.transactionDepth > 0) {
      this.transactionChanged = true;
      return;
    }
    const event = { kind: 'workbookChange' as const, bytes: await this.toZip() };
    for (const listener of this.listeners) listener(event);
    await this.options.onChange?.(event);
  }

  private emitStateIfChanged(activeSheetName: string): void {
    const selection = this.selectionState;
    const selectionKey = `${selection.kind}:${selection.active.row}:${selection.active.column}:${selection.range.startRow}:${selection.range.startColumn}:${selection.range.endRow}:${selection.range.endColumn}:${this.editingCell?.row ?? 0}:${this.editingCell?.column ?? 0}`;
    const stateKey = `${activeSheetName}|${selectionKey}`;
    if (stateKey === this.lastStateKey) return;
    this.lastStateKey = stateKey;
    const event: SpreadsheetEditorStateEvent = {
      kind: 'stateChange',
      activeSheetName,
      selection,
      editingCell: this.editingCell,
    };
    for (const listener of this.stateListeners) listener(event);
  }

  private renderTabs(): void {
    if (this.options.showSheetTabs === false || !this.runtimeValue) return;
    this.sheetTabs.replaceChildren();
    Object.assign(this.sheetTabs.style, {
      position: 'absolute',
      left: '0',
      right: '0',
      bottom: '0',
      height: '40px',
      display: 'flex',
      alignItems: 'center',
      gap: '8px',
      padding: '0 12px',
      background: '#eef1f5',
      borderTop: '1px solid #e2e8f0',
      boxSizing: 'border-box',
      overflowX: 'auto',
      overflowY: 'hidden',
    });
    for (const sheet of this.runtimeValue.model().sheets) {
      const button = document.createElement('button');
      button.type = 'button';
      button.textContent = sheet.name;
      button.dataset.active = String(sheet.name === this.runtimeValue.activeSheetName());
      Object.assign(button.style, {
        height: '32px',
        minWidth: '64px',
        padding: '0 14px',
        border: '0',
        borderTop: sheet.name === this.runtimeValue.activeSheetName() ? '2px solid #2563eb' : '2px solid transparent',
        borderRadius: '3px 3px 0 0',
        background: sheet.name === this.runtimeValue.activeSheetName() ? '#ffffff' : 'transparent',
        color: sheet.name === this.runtimeValue.activeSheetName() ? '#1d4ed8' : '#475569',
        font: '13px Arial, sans-serif',
        fontWeight: sheet.name === this.runtimeValue.activeSheetName() ? '600' : '400',
        cursor: 'pointer',
        whiteSpace: 'nowrap',
      });
      button.addEventListener('mouseenter', () => {
        if (button.dataset.active !== 'true') button.style.background = '#e2e8f0';
      });
      button.addEventListener('mouseleave', () => {
        if (button.dataset.active !== 'true') button.style.background = 'transparent';
      });
      button.addEventListener('click', () => void this.setActiveSheet(sheet.name));
      this.sheetTabs.appendChild(button);
    }
  }

  private configureScrollbar(track: HTMLDivElement, thumb: HTMLDivElement, horizontal: boolean): void {
    track.className = 'coco-spreadsheet-editor-scrollbar';
    Object.assign(track.style, {
      position: 'absolute',
      zIndex: '10',
      background: '#f1f5f9',
      border: '1px solid #cbd5e1',
      cursor: 'pointer',
      ...(horizontal ? { left: '48px', bottom: '40px', height: '12px' } : { top: '24px', right: '0', bottom: '52px', width: '12px' }),
    });
    thumb.className = 'coco-spreadsheet-editor-scrollbar-thumb';
    Object.assign(thumb.style, { position: 'absolute', background: '#94a3b8', borderRadius: '999px' });
    track.appendChild(thumb);
    track.addEventListener('pointerdown', (event) => {
      const geometry = this.scrollbarState?.[horizontal ? 'horizontal' : 'vertical'];
      if (!geometry || geometry.maxScroll <= 0) return;
      const rect = track.getBoundingClientRect();
      const pointer = horizontal ? event.clientX - rect.left : event.clientY - rect.top;
      this.drag = { axis: horizontal ? 'horizontal' : 'vertical', pointer, thumb: geometry.thumbStart };
      track.setPointerCapture(event.pointerId);
      this.setScroll(horizontal ? scrollFromThumbPosition(geometry, pointer - geometry.thumbLength / 2) : this.scroll.x, horizontal ? this.scroll.y : scrollFromThumbPosition(geometry, pointer - geometry.thumbLength / 2));
    });
    track.addEventListener('pointermove', (event) => {
      const drag = this.drag;
      const geometry = this.scrollbarState?.[horizontal ? 'horizontal' : 'vertical'];
      if (!drag || !geometry || drag.axis !== (horizontal ? 'horizontal' : 'vertical')) return;
      const rect = track.getBoundingClientRect();
      const pointer = horizontal ? event.clientX - rect.left : event.clientY - rect.top;
      const next = scrollFromThumbPosition(geometry, drag.thumb + pointer - drag.pointer);
      this.setScroll(horizontal ? next : this.scroll.x, horizontal ? this.scroll.y : next);
    });
    track.addEventListener('pointerup', () => { this.drag = null; });
  }

  private updateScrollbars(): void {
    const state = this.scrollbarState;
    const layout = this.layout;
    const viewport = this.viewport;
    if (!state || !layout || !viewport) return;
    const horizontal = state.horizontal;
    const vertical = state.vertical;
    this.horizontalScrollbar.style.display = horizontal.maxScroll > 0 ? 'block' : 'none';
    this.verticalScrollbar.style.display = vertical.maxScroll > 0 ? 'block' : 'none';
    Object.assign(this.horizontalScrollbar.style, { left: `${layout.rowHeaderWidth + layout.getColumnX(viewport.freeze.frozenColumns + 1)}px`, width: `${horizontal.trackLength}px` });
    Object.assign(this.horizontalThumb.style, { left: `${horizontal.thumbStart}px`, width: `${horizontal.thumbLength}px`, top: '1px', bottom: '1px' });
    Object.assign(this.verticalScrollbar.style, { top: `${layout.columnHeaderHeight + layout.getRowY(viewport.freeze.frozenRows + 1)}px`, height: `${vertical.trackLength}px` });
    Object.assign(this.verticalThumb.style, { top: `${vertical.thumbStart}px`, height: `${vertical.thumbLength}px`, left: '1px', right: '1px' });
  }

  private setScroll(x: number, y: number): void {
    this.scroll = { x, y };
    this.render();
  }

  private readonly handleWheel = (event: WheelEvent): void => {
    if (!this.runtimeValue) return;
    event.preventDefault();
    const next = applyWheelDelta(this.scroll, event.deltaX, event.deltaY);
    this.setScroll(next.x, next.y);
  };

  private assertAlive(): void {
    if (this.disposed) throw new Error('SpreadsheetEditor: editor has been destroyed');
  }
}

function coerceEditorValue(value: string): ExcelCellValue {
  if (value === '') return null;
  if (value === 'TRUE') return true;
  if (value === 'FALSE') return false;
  const number = Number(value);
  return Number.isFinite(number) && value.trim() !== '' ? number : value;
}

function containsPoint(
  rect: { x: number; y: number; width: number; height: number },
  point: { x: number; y: number },
): boolean {
  return point.x >= rect.x && point.y >= rect.y && point.x <= rect.x + rect.width && point.y <= rect.y + rect.height;
}

function findChartAnchor(
  drawings: ExcelDrawingModel[],
  chart: ExcelChartModel,
): ExcelDrawingAnchorModel | undefined {
  const drawing = drawings.find((item) => item.path === chart.drawingPath);
  const anchors = drawing?.anchors ?? drawings.flatMap((item) => item.anchors);
  return anchors.find((anchor) =>
    anchor.object?.kind === 'chart' &&
    (!chart.relationshipId || anchor.object.relationshipId === chart.relationshipId),
  );
}

function chartLogicalRect(
  layout: SpreadsheetLayout,
  anchor: ExcelDrawingAnchorModel | undefined,
): SpreadsheetRect | null {
  if (!anchor || anchor.kind === 'absolute' || !anchor.from) return null;
  const startColumn = markerIndex(anchor.from.column);
  const startRow = markerIndex(anchor.from.row);
  const endColumn = anchor.to ? markerIndex(anchor.to.column) : startColumn + 5;
  const endRow = anchor.to ? markerIndex(anchor.to.row) : startRow + 12;
  const startX = layout.getColumnX(startColumn) + markerOffset(anchor.from.columnOffset);
  const startY = layout.getRowY(startRow) + markerOffset(anchor.from.rowOffset);
  const endX = layout.getColumnX(endColumn) + markerOffset(anchor.to?.columnOffset);
  const endY = layout.getRowY(endRow) + markerOffset(anchor.to?.rowOffset);
  return {
    x: startX,
    y: startY,
    width: Math.max(1, endX - startX),
    height: Math.max(1, endY - startY),
  };
}

function chartScreenFragments(
  logicalRect: SpreadsheetRect,
  viewport: SpreadsheetViewportManager,
): SpreadsheetRect[] {
  const fragments: SpreadsheetRect[] = [];
  for (const item of viewport.viewports) {
    if (!isBodyViewport(item)) continue;
    const visibleLeft = item.originX;
    const visibleTop = item.originY;
    const visibleRight = item.originX + item.screenRect.width;
    const visibleBottom = item.originY + item.screenRect.height;
    const left = Math.max(logicalRect.x, visibleLeft);
    const top = Math.max(logicalRect.y, visibleTop);
    const right = Math.min(logicalRect.x + logicalRect.width, visibleRight);
    const bottom = Math.min(logicalRect.y + logicalRect.height, visibleBottom);
    if (right <= left || bottom <= top) continue;
    fragments.push({
      x: item.screenRect.x + left - item.originX,
      y: item.screenRect.y + top - item.originY,
      width: right - left,
      height: bottom - top,
    });
  }
  return fragments;
}

function chartAnchorForLogicalRect(
  layout: SpreadsheetLayout,
  rect: SpreadsheetRect,
  original: ExcelDrawingAnchorModel,
): {
  from: { row: number; column: number; rowOffset?: number; columnOffset?: number };
  to?: { row: number; column: number; rowOffset?: number; columnOffset?: number };
} {
  const marker = (x: number, y: number) => {
    const column = layout.columnAt(Math.max(0, x));
    const row = layout.rowAt(Math.max(0, y));
    return {
      column: column - 1,
      row: row - 1,
      columnOffset: Math.round((x - layout.getColumnX(column)) * 9525),
      rowOffset: Math.round((y - layout.getRowY(row)) * 9525),
    };
  };
  const from = marker(rect.x, rect.y);
  const to = original.to ? marker(rect.x + rect.width, rect.y + rect.height) : undefined;
  return { from, to };
}

function isBodyViewport(viewport: SpreadsheetViewport): boolean {
  return viewport.kind === 'main' || viewport.kind === 'frozenTop' ||
    viewport.kind === 'frozenLeft' || viewport.kind === 'frozenTopLeft';
}

function markerIndex(value: number | undefined): number {
  return Math.max(1, Math.floor(value ?? 0) + 1);
}

function markerOffset(value: number | undefined): number {
  // OOXML drawing offsets are EMUs; 9,525 EMUs equal one CSS pixel at 96 DPI.
  return Number.isFinite(value) ? (value ?? 0) / 9525 : 0;
}

function columnNumberFromLetters(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) {
    value = value * 26 + character.charCodeAt(0) - 64;
  }
  return value;
}

function autoFillPreviewRangeForHit(
  source: SpreadsheetRange,
  hit: ReturnType<SpreadsheetViewportManager['hitTest']>,
): SpreadsheetRange | null {
  if (hit.kind !== 'cell') return null;
  const down = Math.max(0, hit.row - source.endRow);
  const right = Math.max(0, hit.column - source.endColumn);
  if (!down && !right) return null;
  if (down >= right) return { ...source, endRow: source.endRow + down };
  return { ...source, endColumn: source.endColumn + right };
}

function rangesIntersect(
  left: { startRow: number; startColumn: number; endRow: number; endColumn: number },
  right: { startRow: number; startColumn: number; endRow: number; endColumn: number },
): boolean {
  return !(
    left.endRow < right.startRow ||
    left.startRow > right.endRow ||
    left.endColumn < right.startColumn ||
    left.startColumn > right.endColumn
  );
}

function formulaReferenceRangesFromText(
  text: string,
): Array<{ range: ReturnType<typeof normalizeSelectionRange>; stroke: string }> {
  if (!text.startsWith('=')) return [];
  const references = text.match(/(?:\$?[A-Z]{1,3}\$?\d+)(?::(?:\$?[A-Z]{1,3}\$?\d+))?/g) ?? [];
  const colors = ['#DC2626', '#16A34A', '#7C3AED', '#EA580C', '#0891B2'];
  return references.flatMap((reference, index) => {
    const parts = reference.split(':');
    const start = parts[0] ?? '';
    const end = parts[1] ?? start;
    const first = formulaReferenceCell(start);
    const last = formulaReferenceCell(end);
    if (!first || !last) return [];
    return [{
      range: normalizeSelectionRange({
        startRow: first.row,
        startColumn: first.column,
        endRow: last.row,
        endColumn: last.column,
      }),
      stroke: colors[index % colors.length] ?? '#DC2626',
    }];
  });
}

function formulaReferenceCell(value: string): { row: number; column: number } | null {
  const match = value.replace(/\$/g, '').match(/^([A-Z]{1,3})(\d+)$/);
  if (!match) return null;
  let column = 0;
  for (const character of match[1] ?? '') column = column * 26 + character.charCodeAt(0) - 64;
  const row = Number(match[2] ?? '');
  return Number.isFinite(row) && row > 0 && column > 0 ? { row, column } : null;
}
