import {
  ExcelWorkbook,
  SpreadsheetFormulaEngine,
  rewriteFormulaForRowSort,
  diffExcelWorkbookModels,
  type ExcelCellValue,
  type ExcelWorkbookLoadOptions,
  type ExcelWorkbookModel,
  type ExcelBorderModel,
  type ExcelColorModel,
  type ExcelAgentTransactionGroupOptions,
  type ExcelHistoryApplyResult,
  type ExcelSheetViewModel,
  type ExcelTransactionBuilder,
  type ExcelTransactionGroupOptions,
  type ExcelTransactionGroupSummary,
  type ExcelWorkbookDiff,
  type SpreadsheetFormulaResult,
} from '@coco-vexel/coco-office/excel';
import { formatCellAddress, formatRangeAddress, parseCellAddress, parseRangeAddress } from '@coco-vexel/coco-office/excel';
import { buildSpreadsheetSheetRenderModel } from './render-model.js';
import { cssPixelsToExcelColumnWidth } from './layout.js';
import { buildSpreadsheetAutoFillPlan } from './auto-fill.js';
import { buildSpreadsheetReplacePlan } from './search.js';
import {
  spreadsheetRowsForColumnFilters,
  spreadsheetRowsForFilter,
  spreadsheetSortRows,
  type SpreadsheetColumnFilter,
  type SpreadsheetSortDirection,
} from './filter.js';
import type { SpreadsheetRange, SpreadsheetSheetRenderModel } from './types.js';
import type { SpreadsheetClipboardMerge } from './clipboard.js';
import { EditorReadOnlyError } from '../read-only.js';

export interface SpreadsheetEditorRuntimeOptions {
  readOnly?: boolean;
}

export interface SpreadsheetEditorRuntimeState {
  workbook: ExcelWorkbook;
  model: ExcelWorkbookModel;
  activeSheetName: string;
  activeSheet: SpreadsheetSheetRenderModel;
  formulaResults: Map<string, SpreadsheetFormulaResult>;
  formulaEngine: SpreadsheetFormulaEngine;
}

export interface SpreadsheetCellEditHistoryOptions {
  mergeUndo?: boolean;
  mergeWindowMs?: number;
}

export type SpreadsheetUserTransactionOptions = Omit<ExcelTransactionGroupOptions, 'source'>;
export type SpreadsheetAgentTransactionRunner = (runtime: SpreadsheetEditorRuntime) => Promise<void>;

export interface SpreadsheetAgentTransactionPreview {
  result: ExcelHistoryApplyResult | null;
  summary: ExcelTransactionGroupSummary | null;
  diff: ExcelWorkbookDiff;
  diffSummary: SpreadsheetPreviewDiffSummary;
  beforeModel: ExcelWorkbookModel;
  afterModel: ExcelWorkbookModel;
  activeSheetName: string;
}

export interface SpreadsheetPreviewDiffSummary {
  changed: boolean;
  changedSheets: number;
  sheetNames: string[];
  cellChanges: {
    total: number;
    added: number;
    modified: number;
    deleted: number;
  };
  structuralChanged: boolean;
  objectChanged: boolean;
  sampleCells: SpreadsheetPreviewCellChangeSummary[];
}

export interface SpreadsheetPreviewCellChangeSummary {
  sheetName: string;
  address: string;
  kind: 'added' | 'deleted' | 'modified';
  before: ExcelCellValue | null;
  after: ExcelCellValue | null;
}

export type SpreadsheetPasteMode = 'all' | 'values' | 'formats';

export interface SpreadsheetPasteOptions {
  mode?: SpreadsheetPasteMode;
  merges?: SpreadsheetClipboardMerge[];
  styleIds?: (number | undefined)[][];
}

export class SpreadsheetEditorRuntime {
  private activeTransaction: ExcelTransactionBuilder | null = null;
  private activeTransactionSource: 'user' | 'agent' | null = null;
  private transactionNeedsRefresh = false;
  private readOnlyValue: boolean;

  private constructor(private state: SpreadsheetEditorRuntimeState, options: SpreadsheetEditorRuntimeOptions = {}) {
    this.readOnlyValue = options.readOnly === true;
  }

  static async fromZip(
    bytes: Uint8Array | ArrayBuffer,
    options: ExcelWorkbookLoadOptions & SpreadsheetEditorRuntimeOptions = {},
  ): Promise<SpreadsheetEditorRuntime> {
    const { readOnly, ...loadOptions } = options;
    const workbook = await ExcelWorkbook.fromZip(bytes, loadOptions);
    return SpreadsheetEditorRuntime.fromWorkbook(workbook, { readOnly });
  }

  static async fromWorkbook(workbook: ExcelWorkbook, options: SpreadsheetEditorRuntimeOptions = {}): Promise<SpreadsheetEditorRuntime> {
    const model = await workbook.toModel();
    const firstSheet = model.sheets[0];
    if (!firstSheet) throw new Error('SpreadsheetEditorRuntime: workbook has no sheets');
    const formulaEngine = new SpreadsheetFormulaEngine(firstSheet, { sheets: model.sheets });
    const formulaResults = formulaEngine.calculate();
    return new SpreadsheetEditorRuntime({
      workbook,
      model,
      activeSheetName: firstSheet.name,
      formulaResults,
      formulaEngine,
      activeSheet: buildSpreadsheetSheetRenderModel(firstSheet, model.styles, { formulaResults }),
    }, options);
  }

  isReadOnly(): boolean {
    return this.readOnlyValue;
  }

  setReadOnly(readOnly: boolean): void {
    if (this.activeTransaction) throw new Error('SpreadsheetEditorRuntime: cannot change read-only state during a transaction');
    this.readOnlyValue = readOnly;
  }

  workbook(): ExcelWorkbook {
    this.assertWritable();
    return this.state.workbook;
  }

  model(): ExcelWorkbookModel {
    return this.state.model;
  }

  activeSheetName(): string {
    return this.state.activeSheetName;
  }

  activeSheet(): SpreadsheetSheetRenderModel {
    return this.state.activeSheet;
  }

  canUndo(): boolean {
    return !this.readOnlyValue && this.state.workbook.history().canUndo();
  }

  canRedo(): boolean {
    return !this.readOnlyValue && this.state.workbook.history().canRedo();
  }

  async undo(): Promise<ExcelHistoryApplyResult | null> {
    this.assertWritable();
    const result = await this.state.workbook.history().undo();
    if (!result) return null;
    await this.refresh();
    return result;
  }

  async redo(): Promise<ExcelHistoryApplyResult | null> {
    this.assertWritable();
    const result = await this.state.workbook.history().redo();
    if (!result) return null;
    await this.refresh();
    return result;
  }

  async runUserTransaction(
    options: SpreadsheetUserTransactionOptions,
    run: () => Promise<void>,
  ): Promise<ExcelHistoryApplyResult | null> {
    this.assertWritable();
    return this.runScopedTransaction(
      'user',
      (apply) => this.state.workbook.history().runGroup({ ...options, source: 'user' }, apply),
      run,
    );
  }

  async runAgentTransaction(
    options: ExcelAgentTransactionGroupOptions,
    run: SpreadsheetAgentTransactionRunner,
  ): Promise<ExcelHistoryApplyResult | null> {
    this.assertWritable();
    return this.runScopedTransaction(
      'agent',
      (apply) => this.state.workbook.history().runAgentGroup(options, apply),
      () => run(this),
    );
  }

  async previewAgentTransaction(
    options: ExcelAgentTransactionGroupOptions,
    run: SpreadsheetAgentTransactionRunner,
  ): Promise<SpreadsheetAgentTransactionPreview> {
    // Preview uses a cloned workbook so hosts can show diff/summary before committing.
    // Web and desktop agents should reuse the same runner for preview and apply.
    const beforeModel = cloneWorkbookModel(this.state.model);
    const previewWorkbook = await ExcelWorkbook.fromModel(beforeModel, { mode: 'inline' });
    const previewRuntime = await SpreadsheetEditorRuntime.fromWorkbook(previewWorkbook);
    await previewRuntime.setActiveSheet(this.state.activeSheetName);
    try {
      const result = await previewRuntime.runAgentTransaction(options, run);
      const afterModel = cloneWorkbookModel(previewRuntime.model());
      const diff = diffExcelWorkbookModels(beforeModel, afterModel);
      return {
        result,
        summary: result?.summary ?? null,
        diff,
        diffSummary: summarizeSpreadsheetPreviewDiff(diff),
        beforeModel,
        afterModel,
        activeSheetName: previewRuntime.activeSheetName(),
      };
    } finally {
      await previewRuntime.dispose();
    }
  }

  async setActiveSheet(sheetName: string): Promise<void> {
    const sheet = this.state.model.sheets.find((item) => item.name === sheetName);
    if (!sheet) throw new Error(`SpreadsheetEditorRuntime: sheet not found: ${sheetName}`);
    const formulaEngine = new SpreadsheetFormulaEngine(sheet, { sheets: this.state.model.sheets });
    const formulaResults = formulaEngine.calculate();
    this.state.activeSheetName = sheet.name;
    this.state.formulaResults = formulaResults;
    this.state.formulaEngine = formulaEngine;
    this.state.activeSheet = buildSpreadsheetSheetRenderModel(sheet, this.state.model.styles, {
      formulaResults,
    });
  }

  async addSheet(name: string): Promise<void> {
    await this.runWorksheetStructureMutation(
      (tx) => tx.addSheet(name),
      () => this.state.workbook.worksheets.add(name).then(() => undefined),
    );
    await this.setActiveSheet(name);
  }

  async renameSheet(sheetName: string, name: string): Promise<void> {
    const wasActive = sheetName === this.state.activeSheetName;
    await this.runWorksheetStructureMutation(
      (tx) => tx.renameSheet(sheetName, name),
      () => this.state.workbook.worksheets.getItem(sheetName).rename(name),
    );
    if (wasActive) await this.setActiveSheet(name);
  }

  async copySheet(sheetName: string, newName: string): Promise<void> {
    await this.runWorksheetStructureMutation(
      (tx) => tx.copySheet(sheetName, newName),
      () => this.state.workbook.worksheets.copy(sheetName, newName).then(() => undefined),
    );
    await this.setActiveSheet(newName);
  }

  async deleteSheet(sheetName: string): Promise<void> {
    const index = this.state.model.sheets.findIndex((sheet) => sheet.name === sheetName);
    if (index === -1) throw new Error(`SpreadsheetEditorRuntime: sheet not found: ${sheetName}`);
    if (this.state.model.sheets.length === 1) {
      throw new Error('SpreadsheetEditorRuntime: cannot delete the last worksheet');
    }
    const nextActive = this.state.model.sheets[index + 1] ?? this.state.model.sheets[index - 1];
    const wasActive = sheetName === this.state.activeSheetName;
    await this.runWorksheetStructureMutation(
      (tx) => tx.deleteSheet(sheetName),
      () => this.state.workbook.worksheets.deleteItem(sheetName),
    );
    if (wasActive && nextActive) await this.setActiveSheet(nextActive.name);
  }

  async moveSheet(sheetName: string, toIndex: number): Promise<void> {
    await this.runWorksheetStructureMutation(
      (tx) => tx.moveSheet(sheetName, toIndex),
      () => this.state.workbook.worksheets.move(sheetName, toIndex),
    );
  }

  async refresh(): Promise<void> {
    const model = await this.state.workbook.toModel();
    const sheet =
      model.sheets.find((item) => item.name === this.state.activeSheetName) ?? model.sheets[0];
    if (!sheet) throw new Error('SpreadsheetEditorRuntime: workbook has no sheets');
    const formulaEngine = new SpreadsheetFormulaEngine(sheet, { sheets: model.sheets });
    const formulaResults = formulaEngine.calculate();
    this.state = {
      ...this.state,
      model,
      activeSheetName: sheet.name,
      formulaResults,
      formulaEngine,
      activeSheet: buildSpreadsheetSheetRenderModel(sheet, model.styles, { formulaResults }),
    };
  }

  async setCellValue(
    address: string,
    value: ExcelCellValue,
    historyOptions: SpreadsheetCellEditHistoryOptions = {},
  ): Promise<void> {
    const targetAddress = this.editableCellAddress(address);
    await this.runUserMutation(this.cellEditHistoryGroup(`Set ${targetAddress}`, targetAddress, historyOptions), async (tx) => {
      await tx.setCellValue(this.state.activeSheetName, targetAddress, value);
    });
  }

  async setCellFormula(
    address: string,
    formula: string | null,
    historyOptions: SpreadsheetCellEditHistoryOptions = {},
  ): Promise<void> {
    const targetAddress = this.editableCellAddress(address);
    await this.runUserMutation(this.cellEditHistoryGroup(`Set ${targetAddress} formula`, targetAddress, historyOptions), async (tx) => {
      await tx.setCellFormula(this.state.activeSheetName, targetAddress, formula);
    });
  }

  async replaceAll(
    query: string,
    replacement: string,
    matchCase = false,
  ): Promise<number> {
    const plan = buildSpreadsheetReplacePlan(this.state.activeSheet, { query, replacement, matchCase });
    if (plan.length === 0) return 0;
    await this.runUserMutation({ label: `Replace ${query}` }, async (tx) => {
      for (const target of plan) {
        if (target.formula !== undefined) {
          await tx.setCellFormula(this.state.activeSheetName, target.address, target.formula);
        } else {
          await tx.setCellValue(this.state.activeSheetName, target.address, target.value ?? null);
        }
      }
    });
    return plan.length;
  }

  async filterRows(
    startRow: number,
    endRow: number,
    column: number,
    value: string | readonly string[] | null,
  ): Promise<void> {
    const range = { startRow, startColumn: column, endRow, endColumn: column };
    const hiddenRows = value === null
      ? []
      : spreadsheetRowsForFilter(this.state.activeSheet, { ...range, startColumn: 1, endColumn: this.state.activeSheet.maxColumn }, column, value);
    await this.runUserMutation({ label: value === null ? 'Clear filter' : 'Filter rows' }, async (tx) => {
      await tx.runSnapshot('Filter rows', 'spreadsheet row filter', async () => {
        await this.state.workbook.executeCode<void>(
          `
          const sheet = workbook.worksheets.getItem(args.sheetName);
          const hidden = new Set(args.hiddenRows);
          for (let index = args.startRow + 1; index <= args.endRow; index += 1) {
            let row = sheet.model.rows.find((item) => item.index === index);
            if (hidden.has(index)) {
              row ??= { index };
              row.hidden = true;
              if (!sheet.model.rows.includes(row)) sheet.model.rows.push(row);
            } else if (row) {
              delete row.hidden;
            }
          }
          sheet.model.rows.sort((left, right) => left.index - right.index);
          `,
          { sheetName: this.state.activeSheetName, startRow, endRow, hiddenRows },
        );
      });
    });
  }

  async filterRowsByColumns(
    startRow: number,
    endRow: number,
    filters: readonly SpreadsheetColumnFilter[],
  ): Promise<void> {
    const range = {
      startRow,
      startColumn: 1,
      endRow,
      endColumn: this.state.activeSheet.maxColumn,
    };
    const hiddenRows = spreadsheetRowsForColumnFilters(this.state.activeSheet, range, filters);
    await this.runUserMutation({ label: 'Filter rows' }, async (tx) => {
      await tx.runSnapshot('Filter rows', 'spreadsheet row filter', async () => {
        await this.state.workbook.executeCode<void>(
          `
          const sheet = workbook.worksheets.getItem(args.sheetName);
          const hidden = new Set(args.hiddenRows);
          for (let index = args.startRow + 1; index <= args.endRow; index += 1) {
            let row = sheet.model.rows.find((item) => item.index === index);
            if (hidden.has(index)) {
              row ??= { index };
              row.hidden = true;
              if (!sheet.model.rows.includes(row)) sheet.model.rows.push(row);
            } else if (row) {
              delete row.hidden;
            }
          }
          sheet.model.rows.sort((left, right) => left.index - right.index);
          `,
          { sheetName: this.state.activeSheetName, startRow, endRow, hiddenRows },
        );
      });
    });
  }

  async sortRows(
    startRow: number,
    endRow: number,
    column: number,
    direction: SpreadsheetSortDirection,
  ): Promise<void> {
    const range = { startRow, startColumn: 1, endRow, endColumn: this.state.activeSheet.maxColumn };
    const orderedRows = spreadsheetSortRows(this.state.activeSheet, range, column, direction);
    const rowMap = new Map(orderedRows.map((sourceRow, index) => [sourceRow, startRow + index + 1]));
    const formulaUpdates = this.state.activeSheet.source.cells
      .filter((cell) => cell.formula)
      .map((cell) => [
        cell.address,
        rewriteFormulaForRowSort(cell.formula ?? '', rowMap, { sheetName: this.state.activeSheetName }),
      ]);
    await this.runUserMutation({ label: direction === 'asc' ? 'Sort ascending' : 'Sort descending' }, async (tx) => {
      await tx.runSnapshot('Sort rows', 'spreadsheet row sort', async () => {
        await this.state.workbook.executeCode<void>(
          `
          const sheet = workbook.worksheets.getItem(args.sheetName);
          const rowMap = new Map(args.rowMap);
          const formulas = new Map(args.formulas);
          function formatCellAddress(row, column) {
            let value = column;
            let letters = '';
            while (value > 0) {
              const remainder = (value - 1) % 26;
              letters = String.fromCharCode(65 + remainder) + letters;
              value = Math.floor((value - 1) / 26);
            }
            return letters + row;
          }
          for (const cell of sheet.model.cells) {
            const nextFormula = formulas.get(cell.address);
            if (nextFormula !== undefined) cell.formula = nextFormula;
            const nextRow = rowMap.get(cell.row);
            if (nextRow === undefined) continue;
            cell.row = nextRow;
            cell.address = formatCellAddress(cell.row, cell.column);
          }
          for (const row of sheet.model.rows) {
            const nextIndex = rowMap.get(row.index);
            if (nextIndex !== undefined) row.index = nextIndex;
          }
          sheet.model.cells.sort((left, right) => left.row - right.row || left.column - right.column);
          sheet.model.rows.sort((left, right) => left.index - right.index);
          `,
          { sheetName: this.state.activeSheetName, rowMap: [...rowMap.entries()], formulas: formulaUpdates },
        );
      });
    });
  }

  async setRangeValues(
    startRow: number,
    startColumn: number,
    values: ExcelCellValue[][],
  ): Promise<void> {
    const rowCount = values.length;
    const columnCount = values.reduce((max, row) => Math.max(max, row.length), 0);
    const normalized = values.map((row) => {
      const next = row.slice();
      while (next.length < columnCount) next.push(null);
      return next;
    });
    await this.runUserMutation({ label: 'Set range values' }, async (tx) => {
      await tx.setRangeValues(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        normalized,
      );
    });
  }

  async autoFillRange(sourceRange: SpreadsheetRange, targetRange: SpreadsheetRange): Promise<void> {
    const plan = buildSpreadsheetAutoFillPlan(this.state.activeSheet, sourceRange, targetRange);
    if (plan.cells.length === 0) return;
    await this.runUserMutation({ label: 'Auto fill' }, async (tx) => {
      for (const cell of plan.cells) {
        const address = formatCellAddress(cell.row, cell.column);
        if (cell.formula) await tx.setCellFormula(this.state.activeSheetName, address, cell.formula);
        else await tx.setCellValue(this.state.activeSheetName, address, cell.value);
        await tx.setRangeStyleId(this.state.activeSheetName, address, cell.styleId);
      }
    });
  }

  async pasteRangeValues(
    startRow: number,
    startColumn: number,
    values: ExcelCellValue[][],
    merges: SpreadsheetClipboardMerge[] = [],
    styleIds: (number | undefined)[][] = [],
  ): Promise<void> {
    await this.pasteRange(startRow, startColumn, values, { mode: 'all', merges, styleIds });
  }

  async pasteRange(
    startRow: number,
    startColumn: number,
    values: ExcelCellValue[][],
    options: SpreadsheetPasteOptions = {},
  ): Promise<void> {
    const mode = options.mode ?? 'all';
    const rowCount = values.length;
    const columnCount = values.reduce((max, row) => Math.max(max, row.length), 0);
    const normalized = values.map((row) => {
      const next = row.slice();
      while (next.length < columnCount) next.push(null);
      return next;
    });
    const effectiveRowCount = Math.max(1, rowCount);
    const effectiveColumnCount = Math.max(1, columnCount);
    const targetRange = {
      startRow,
      startColumn,
      endRow: startRow + effectiveRowCount - 1,
      endColumn: startColumn + effectiveColumnCount - 1,
    };
    const existingMerges = [...this.state.activeSheet.mergeMasterByAddress.values()]
      .filter((merge) => rangesIntersect(targetRange, merge.range))
      .map((merge) => merge.address);
    await this.runUserMutation({ label: pasteHistoryLabel(mode) }, async (tx) => {
      if (mode !== 'formats') {
        for (const address of existingMerges) await tx.unmergeRange(this.state.activeSheetName, address);
        await tx.setRangeValues(
          this.state.activeSheetName,
          formatRangeAddress(startRow, startColumn, effectiveRowCount, effectiveColumnCount),
          normalized,
        );
      }
      if (mode !== 'values') {
        await applySpreadsheetPasteStyles(
          tx,
          this.state.activeSheetName,
          startRow,
          startColumn,
          effectiveColumnCount,
          options.styleIds ?? [],
        );
      }
      if (mode === 'all') await applySpreadsheetPasteMerges(tx, this.state.activeSheetName, startRow, startColumn, options.merges ?? []);
    });
  }

  async clearRange(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Clear range' }, async (tx) => {
      await tx.clearRange(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
      );
    });
  }

  async setRowHeight(row: number, height: number): Promise<void> {
    await this.runUserMutation({ label: 'Set row height' }, async (tx) => {
      await tx.setRowHeight(this.state.activeSheetName, row, height);
    });
  }

  async setColumnWidth(column: number, width: number): Promise<void> {
    await this.runUserMutation({ label: 'Set column width' }, async (tx) => {
      await tx.setColumnWidth(this.state.activeSheetName, column, cssPixelsToExcelColumnWidth(width));
    });
  }

  async setFreezePanes(frozenRows: number, frozenColumns: number): Promise<void> {
    const rows = normalizeFrozenPaneCount(frozenRows);
    const columns = normalizeFrozenPaneCount(frozenColumns);
    const views = spreadsheetSheetViewsWithFrozenPane(this.state.activeSheet.source.sheetViews, rows, columns);
    await this.runUserMutation({ label: rows || columns ? 'Freeze panes' : 'Unfreeze panes' }, async (tx) => {
      await tx.setSheetViews(this.state.activeSheetName, views);
    });
  }

  async insertRows(rowIndex: number, count = 1): Promise<void> {
    await this.runUserMutation({ label: 'Insert rows' }, async (tx) => {
      await tx.insertRows(this.state.activeSheetName, rowIndex, count);
    });
  }

  async deleteRows(rowIndex: number, count = 1): Promise<void> {
    await this.runUserMutation({ label: 'Delete rows' }, async (tx) => {
      await tx.deleteRows(this.state.activeSheetName, rowIndex, count);
    });
  }

  async insertColumns(columnIndex: number, count = 1): Promise<void> {
    await this.runUserMutation({ label: 'Insert columns' }, async (tx) => {
      await tx.insertColumns(this.state.activeSheetName, columnIndex, count);
    });
  }

  async deleteColumns(columnIndex: number, count = 1): Promise<void> {
    await this.runUserMutation({ label: 'Delete columns' }, async (tx) => {
      await tx.deleteColumns(this.state.activeSheetName, columnIndex, count);
    });
  }

  async setRangeBold(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    value: boolean | undefined,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Set bold' }, async (tx) => {
      await tx.setRangeBold(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        value,
      );
    });
  }

  async setRangeItalic(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    value: boolean | undefined,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Set italic' }, async (tx) => {
      await tx.setRangeItalic(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        value,
      );
    });
  }

  async setRangeUnderline(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    value: boolean | undefined,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Set underline' }, async (tx) => {
      await tx.setRangeUnderline(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        value,
      );
    });
  }

  async setRangeFontColor(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    color: ExcelColorModel | undefined,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Set font color' }, async (tx) => {
      await tx.setRangeFontColor(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        color,
      );
    });
  }

  async setRangeFillColor(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    color: ExcelColorModel | undefined,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Set fill color' }, async (tx) => {
      await tx.setRangeFillColor(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        color,
      );
    });
  }

  async setRangeHorizontalAlign(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    horizontal: 'left' | 'center' | 'right' | undefined,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Set horizontal align' }, async (tx) => {
      await tx.setRangeHorizontalAlign(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        horizontal,
      );
    });
  }

  async setRangeVerticalAlign(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    vertical: 'top' | 'center' | 'bottom' | undefined,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Set vertical align' }, async (tx) => {
      await tx.setRangeVerticalAlign(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        vertical,
      );
    });
  }

  async setRangeNumberFormat(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    code: string | undefined,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Set number format' }, async (tx) => {
      await tx.setRangeNumberFormat(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        code,
      );
    });
  }

  async clearRangeFormat(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
  ): Promise<void> {
    await this.runUserMutation({ label: 'Clear format' }, async (tx) => {
      await tx.clearRangeFormat(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
      );
    });
  }

  async mergeRange(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
  ): Promise<void> {
    const address = formatRangeAddress(startRow, startColumn, rowCount, columnCount);
    const masterStyleId = await this.cellStyleId(formatCellAddress(startRow, startColumn));
    await this.runUserMutation({ label: 'Merge range' }, async (tx) => {
      // Editor-level merge semantics hide follower cells, so normalize their styles first.
      await tx.setRangeStyleId(this.state.activeSheetName, address, masterStyleId);
      await tx.mergeRange(this.state.activeSheetName, address);
    });
  }

  async unmergeRanges(addresses: string[]): Promise<void> {
    if (addresses.length === 0) return;
    const merges = addresses.map((address) => {
      const merge = [...this.state.activeSheet.mergeMasterByAddress.values()].find((item) => item.address === address);
      const masterAddress = merge ? formatCellAddress(merge.range.startRow, merge.range.startColumn) : masterAddressForRange(address);
      return { address, masterAddress };
    });
    await this.runUserMutation({ label: 'Unmerge ranges' }, async (tx) => {
      for (const merge of merges) {
        // Prevent stale hidden follower styles from reappearing after unmerge.
        await tx.setRangeStyleId(this.state.activeSheetName, merge.address, await this.cellStyleId(merge.masterAddress));
        await tx.unmergeRange(this.state.activeSheetName, merge.address);
      }
    });
  }

  async setRangeBorder(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
    mode: 'all' | 'outer' | 'inner' | 'clear',
  ): Promise<void> {
    if (mode === 'clear') {
      await this.runUserMutation({ label: 'Clear border' }, async (tx) => {
        await tx.setRangeBorder(
          this.state.activeSheetName,
          formatRangeAddress(startRow, startColumn, rowCount, columnCount),
          {},
        );
      });
      return;
    }

    const side = spreadsheetDefaultBorderSide();
    if (mode === 'all') {
      await this.runUserMutation({ label: 'Set all borders' }, async (tx) => {
        await tx.setRangeBorder(
          this.state.activeSheetName,
          formatRangeAddress(startRow, startColumn, rowCount, columnCount),
          { left: side, right: side, top: side, bottom: side },
        );
      });
      return;
    }

    await this.runUserMutation({ label: `Set ${mode} borders` }, async (tx) => {
      await tx.setRangeBorder(
        this.state.activeSheetName,
        formatRangeAddress(startRow, startColumn, rowCount, columnCount),
        {},
      );
      for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
        for (let columnOffset = 0; columnOffset < columnCount; columnOffset++) {
          const border: ExcelBorderModel = {};
          if (mode === 'outer') {
            if (rowOffset === 0) border.top = side;
            if (rowOffset === rowCount - 1) border.bottom = side;
            if (columnOffset === 0) border.left = side;
            if (columnOffset === columnCount - 1) border.right = side;
          } else {
            if (columnOffset < columnCount - 1) border.right = side;
            if (rowOffset < rowCount - 1) border.bottom = side;
          }
          if (Object.keys(border).length === 0) continue;
          await tx.setRangeBorder(
            this.state.activeSheetName,
            formatCellAddress(startRow + rowOffset, startColumn + columnOffset),
            border,
          );
        }
      }
    });
  }

  async toZip(): Promise<Uint8Array> {
    return this.state.workbook.toZip();
  }

  async moveChart(
    chartIndex: number,
    position: {
      from: { row: number; column: number; rowOffset?: number; columnOffset?: number };
      to?: { row: number; column: number; rowOffset?: number; columnOffset?: number };
    },
  ): Promise<void> {
    await this.runUserMutation({ label: 'Move chart' }, async (tx) => {
      await tx.runSnapshot('Move chart', 'spreadsheet chart position', async () => {
        const sheetModel = this.state.model.sheets.find((sheet) => sheet.name === this.state.activeSheetName);
        const chart = sheetModel?.charts[chartIndex];
        if (!sheetModel || !chart) return;
        const drawingIndex = sheetModel.drawings.findIndex((drawing) => drawing.path === chart.drawingPath);
        const drawing = drawingIndex >= 0 ? sheetModel.drawings[drawingIndex] : sheetModel.drawings[0];
        if (!drawing) return;
        const anchorIndex = drawing.anchors.findIndex((anchor) =>
          anchor.object?.kind === 'chart' &&
          (!chart.relationshipId || anchor.object.relationshipId === chart.relationshipId),
        );
        if (anchorIndex < 0) return;
        const drawings = await this.state.workbook.worksheets.getItem(this.state.activeSheetName).drawings.items();
        const coreDrawing = drawings[drawingIndex >= 0 ? drawingIndex : 0];
        const anchor = coreDrawing?.anchors.getItem(anchorIndex);
        if (!anchor) return;
        anchor.from.setRow(position.from.row);
        anchor.from.setColumn(position.from.column);
        anchor.from.setRowOffset(position.from.rowOffset);
        anchor.from.setColumnOffset(position.from.columnOffset);
        if (position.to) {
          anchor.to.setRow(position.to.row);
          anchor.to.setColumn(position.to.column);
          anchor.to.setRowOffset(position.to.rowOffset);
          anchor.to.setColumnOffset(position.to.columnOffset);
        }
      });
    });
  }

  async dispose(): Promise<void> {
    await this.state.workbook.dispose();
  }

  editableCellAddress(address: string): string {
    const cell = parseCellAddress(address);
    const normalized = formatCellAddress(cell.row, cell.column);
    return this.state.activeSheet.mergeFollowerByAddress.get(normalized) ?? normalized;
  }

  private rebuildActiveSheetRenderModel(): void {
    this.state.formulaResults = this.state.formulaEngine.calculate();
    this.state.activeSheet = buildSpreadsheetSheetRenderModel(
      this.state.activeSheet.source,
      this.state.model.styles,
      { formulaResults: this.state.formulaResults },
    );
  }

  private async runUserMutation(
    options: SpreadsheetUserTransactionOptions,
    mutate: (transaction: ExcelTransactionBuilder) => Promise<void>,
  ): Promise<void> {
    this.assertWritable();
    if (this.activeTransaction) {
      await mutate(this.activeTransaction);
      this.transactionNeedsRefresh = true;
      return;
    }
    await this.runScopedTransaction(
      'user',
      (apply) => this.state.workbook.history().runGroup({ ...options, source: 'user' }, apply),
      async () => {
        if (!this.activeTransaction) throw new Error('SpreadsheetEditorRuntime: transaction not active');
        await mutate(this.activeTransaction);
        this.transactionNeedsRefresh = true;
      },
    );
  }

  private async runWorksheetStructureMutation(
    recordAgentMutation: (transaction: ExcelTransactionBuilder) => Promise<void>,
    mutate: () => Promise<void>,
  ): Promise<void> {
    this.assertWritable();
    if (this.activeTransactionSource === 'agent' && this.activeTransaction) {
      await recordAgentMutation(this.activeTransaction);
      this.transactionNeedsRefresh = true;
      // Sheet creation/rename changes the names used by following Agent steps.
      // Refresh now so those steps resolve the current workbook structure.
      await this.refresh();
      return;
    }

    // Match desktop spreadsheet behavior: sheet tab operations are not Ctrl/Cmd+Z edits.
    // Agent runs retain the structural patch so their entire action can roll back atomically.
    await mutate();
    await this.refresh();
  }

  private async cellStyleId(address: string): Promise<number | undefined> {
    return (await this.state.workbook.worksheets.getItem(this.state.activeSheetName).getRange(address).styleIds())[0]?.[0];
  }

  private async runScopedTransaction(
    source: 'user' | 'agent',
    execute: (run: (transaction: ExcelTransactionBuilder) => Promise<void>) => Promise<ExcelHistoryApplyResult | null>,
    run: () => Promise<void>,
  ): Promise<ExcelHistoryApplyResult | null> {
    this.assertWritable();
    if (this.activeTransaction) {
      await run();
      this.transactionNeedsRefresh = true;
      return null;
    }

    const previousNeedsRefresh = this.transactionNeedsRefresh;
    this.transactionNeedsRefresh = false;
    try {
      const result = await execute(async (transaction) => {
        this.activeTransaction = transaction;
        this.activeTransactionSource = source;
        try {
          await run();
        } finally {
          this.activeTransaction = null;
          this.activeTransactionSource = null;
        }
      });
      const shouldRefresh = this.transactionNeedsRefresh || result !== null;
      if (shouldRefresh) await this.refresh();
      return result;
    } finally {
      this.activeTransaction = null;
      this.activeTransactionSource = null;
      this.transactionNeedsRefresh = previousNeedsRefresh;
    }
  }

  private assertWritable(): void {
    if (this.readOnlyValue) throw new EditorReadOnlyError('spreadsheet');
  }

  private cellEditHistoryGroup(
    label: string,
    address: string,
    options: SpreadsheetCellEditHistoryOptions,
  ): SpreadsheetUserTransactionOptions {
    const base = { label };
    if (!options.mergeUndo) return base;
    return {
      ...base,
      mergeKey: `edit:${this.state.activeSheetName}!${address}`,
      mergeWindowMs: options.mergeWindowMs ?? 1000,
    };
  }

  private range(startRow: number, startColumn: number, rowCount: number, columnCount: number) {
    return this.state.workbook
      .worksheets
      .getItem(this.state.activeSheetName)
      .getRange(formatRangeAddress(startRow, startColumn, rowCount, columnCount));
  }
}

function spreadsheetDefaultBorderSide(): NonNullable<ExcelBorderModel['left']> {
  return { style: 'thin', color: { rgb: '#111827' } };
}

function normalizeFrozenPaneCount(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.floor(value));
}

function spreadsheetSheetViewsWithFrozenPane(
  sheetViews: ExcelSheetViewModel[],
  frozenRows: number,
  frozenColumns: number,
): ExcelSheetViewModel[] {
  if (!frozenRows && !frozenColumns) {
    return sheetViews.map((view, index) => ({
      ...view,
      pane: index === 0 ? undefined : view.pane ? { ...view.pane } : undefined,
      selections: view.selections.map((selection) => ({ ...selection })),
    }));
  }
  const currentView = sheetViews[0] ?? { workbookViewId: 0, selections: [] };
  const pane = {
    xSplit: frozenColumns || undefined,
    ySplit: frozenRows || undefined,
    topLeftCell: formatCellAddress(frozenRows + 1, frozenColumns + 1),
    activePane: frozenRows && frozenColumns ? 'bottomRight' : frozenRows ? 'bottomLeft' : 'topRight',
    state: 'frozen',
  };
  return [
    {
      ...currentView,
      pane,
      selections: currentView.selections.map((selection) => ({ ...selection })),
    },
    ...sheetViews.slice(1).map((view) => ({
      ...view,
      pane: view.pane ? { ...view.pane } : undefined,
      selections: view.selections.map((selection) => ({ ...selection })),
    })),
  ];
}

async function applySpreadsheetPasteStyles(
  tx: ExcelTransactionBuilder,
  sheetName: string,
  startRow: number,
  startColumn: number,
  columnCount: number,
  styleIds: (number | undefined)[][],
): Promise<void> {
  for (let rowOffset = 0; rowOffset < styleIds.length; rowOffset++) {
    const styleRow = styleIds[rowOffset];
    if (!styleRow) continue;
    for (let columnOffset = 0; columnOffset < columnCount; columnOffset++) {
      await tx.setRangeStyleId(
        sheetName,
        formatCellAddress(startRow + rowOffset, startColumn + columnOffset),
        styleRow[columnOffset],
      );
    }
  }
}

async function applySpreadsheetPasteMerges(
  tx: ExcelTransactionBuilder,
  sheetName: string,
  startRow: number,
  startColumn: number,
  merges: SpreadsheetClipboardMerge[],
): Promise<void> {
  for (const merge of merges) {
    await tx.mergeRange(
      sheetName,
      formatRangeAddress(
        startRow + merge.startRowOffset,
        startColumn + merge.startColumnOffset,
        merge.rowCount,
        merge.columnCount,
      ),
    );
  }
}

function pasteHistoryLabel(mode: SpreadsheetPasteMode): string {
  switch (mode) {
    case 'values':
      return 'Paste values';
    case 'formats':
      return 'Paste formats';
    default:
      return 'Paste range';
  }
}

function cloneWorkbookModel(model: ExcelWorkbookModel): ExcelWorkbookModel {
  return JSON.parse(JSON.stringify(model)) as ExcelWorkbookModel;
}

function summarizeSpreadsheetPreviewDiff(diff: ExcelWorkbookDiff): SpreadsheetPreviewDiffSummary {
  const cellChanges = {
    total: 0,
    added: 0,
    modified: 0,
    deleted: 0,
  };
  const sampleCells: SpreadsheetPreviewCellChangeSummary[] = [];
  for (const sheet of diff.sheets) {
    for (const cell of sheet.cells) {
      cellChanges.total += 1;
      cellChanges[cell.kind] += 1;
      if (sampleCells.length >= 20) continue;
      sampleCells.push({
        sheetName: sheet.afterName ?? sheet.beforeName ?? sheet.sheetId,
        address: cell.address,
        kind: cell.kind,
        before: cell.before?.value ?? null,
        after: cell.after?.value ?? null,
      });
    }
  }
  return {
    changed: diff.changed,
    changedSheets: diff.sheets.length,
    sheetNames: diff.sheets.map((sheet) => sheet.afterName ?? sheet.beforeName ?? sheet.sheetId),
    cellChanges,
    structuralChanged: diff.sheets.some((sheet) => sheet.structuralChanged),
    objectChanged: diff.sheets.some((sheet) => sheet.objectChanged),
    sampleCells,
  };
}

function rangesIntersect(a: SpreadsheetRange, b: SpreadsheetRange): boolean {
  return !(
    a.endRow < b.startRow ||
    a.startRow > b.endRow ||
    a.endColumn < b.startColumn ||
    a.startColumn > b.endColumn
  );
}

function masterAddressForRange(address: string): string {
  const range = parseRangeAddress(address);
  return formatCellAddress(range.startRow, range.startColumn);
}
