import { formatCellAddress, formatRangeAddress, parseRangeAddress } from './address.js';
import type {
  ExcelBorderModel,
  ExcelCellModel,
  ExcelCellValue,
  ExcelColumnModel,
  ExcelColorModel,
  ExcelDefinedNameModel,
  ExcelMergedRangeModel,
  ExcelRowModel,
  ExcelSheetModel,
  ExcelSheetViewModel,
  ExcelWorkbookModel,
} from './model.js';
import type { ExcelWorkbook } from './runtime.js';

export type ExcelTransactionSource = 'user' | 'agent' | 'import' | 'system';

export interface ExcelTransactionGroupMetadata {
  actor?: string;
  description?: string;
  agentRunId?: string;
  agentStepId?: string;
  agentActionId?: string;
  mergeKey?: string;
  [key: string]: unknown;
}

export interface ExcelTransactionGroupOptions {
  source: ExcelTransactionSource;
  label: string;
  id?: string;
  mergeKey?: string;
  mergeWindowMs?: number;
  metadata?: ExcelTransactionGroupMetadata;
}

export interface ExcelAgentTransactionGroupOptions {
  label: string;
  id?: string;
  actor?: string;
  description?: string;
  agentRunId?: string;
  agentStepId?: string;
  agentActionId?: string;
  mergeKey?: string;
  mergeWindowMs?: number;
  metadata?: ExcelTransactionGroupMetadata;
}

export interface ExcelTransactionGroup {
  id: string;
  source: ExcelTransactionSource;
  label: string;
  metadata?: ExcelTransactionGroupMetadata;
  mergeKey?: string;
  mergeWindowMs?: number;
  workbookId: string;
  transactions: ExcelTransaction[];
  affectedSheets: string[];
  affectedRanges: string[];
  createdAt: number;
  committedAt: number;
}

export interface ExcelTransaction {
  id: string;
  label: string;
  forward: ExcelPatch[];
  inverse: ExcelPatch[];
}

export type ExcelPatch =
  | ExcelRangeCellsPatch
  | ExcelRangeFormatPatch
  | ExcelMergedRangesPatch
  | ExcelSheetViewsPatch
  | ExcelWorkbookSheetsPatch
  | ExcelSheetModelPatch
  | ExcelSheetRowsInsertPatch
  | ExcelSheetRowsDeletePatch
  | ExcelSheetRowsRestorePatch
  | ExcelSheetRowsRestoreDeltaPatch
  | ExcelSheetColumnsInsertPatch
  | ExcelSheetColumnsDeletePatch
  | ExcelSheetColumnsRestorePatch
  | ExcelSheetColumnsRestoreDeltaPatch
  | ExcelRowHeightPatch
  | ExcelColumnWidthPatch
  | ExcelSnapshotPatch;

export interface ExcelRangeCellsPatch {
  kind: 'range.cells';
  sheetName: string;
  address: string;
  cells: ExcelCellModel[];
}

export interface ExcelRangeFormatPatch {
  kind: 'range.format';
  sheetName: string;
  address: string;
  formats: ExcelCellFormatSnapshot[];
}

export interface ExcelCellFormatSnapshot {
  row: number;
  column: number;
  styleId: number | undefined;
}

export interface ExcelMergedRangesPatch {
  kind: 'sheet.mergedRanges';
  sheetName: string;
  address: string;
  mergedRanges: string[];
}

export interface ExcelSheetViewsPatch {
  kind: 'sheet.views';
  sheetName: string;
  address: string;
  sheetViews: ExcelSheetViewModel[];
}

export interface ExcelWorkbookSheetsPatch {
  kind: 'workbook.sheets';
  sheetName: string;
  address: string;
  sheets: ExcelSheetModel[];
  definedNames: ExcelDefinedNameModel[];
}

export interface ExcelSheetModelPatch {
  kind: 'sheet.model';
  sheetName: string;
  address: string;
  sheet: ExcelSheetModel;
  definedNames: ExcelDefinedNameModel[];
}

export interface ExcelSheetRowsInsertPatch {
  kind: 'sheet.rows.insert';
  sheetName: string;
  address: string;
  rowIndex: number;
  count: number;
}

export interface ExcelSheetRowsDeletePatch {
  kind: 'sheet.rows.delete';
  sheetName: string;
  address: string;
  rowIndex: number;
  count: number;
}

export interface ExcelSheetRowsRestorePatch {
  kind: 'sheet.rows.restore';
  sheetName: string;
  address: string;
  rowIndex: number;
  count: number;
  sheet: ExcelSheetModel;
  definedNames: ExcelDefinedNameModel[];
}

export interface ExcelSheetRowsRestoreDeltaPatch {
  kind: 'sheet.rows.restoreDelta';
  sheetName: string;
  address: string;
  rowIndex: number;
  count: number;
  cells: ExcelCellModel[];
  rows: ExcelRowModel[];
  mergedRanges: ExcelMergedRangeModel[];
  definedNames: ExcelDefinedNameModel[];
}

export interface ExcelSheetColumnsInsertPatch {
  kind: 'sheet.columns.insert';
  sheetName: string;
  address: string;
  columnIndex: number;
  count: number;
}

export interface ExcelSheetColumnsDeletePatch {
  kind: 'sheet.columns.delete';
  sheetName: string;
  address: string;
  columnIndex: number;
  count: number;
}

export interface ExcelSheetColumnsRestorePatch {
  kind: 'sheet.columns.restore';
  sheetName: string;
  address: string;
  columnIndex: number;
  count: number;
  sheet: ExcelSheetModel;
  definedNames: ExcelDefinedNameModel[];
}

export interface ExcelSheetColumnsRestoreDeltaPatch {
  kind: 'sheet.columns.restoreDelta';
  sheetName: string;
  address: string;
  columnIndex: number;
  count: number;
  cells: ExcelCellModel[];
  columns: ExcelColumnModel[];
  mergedRanges: ExcelMergedRangeModel[];
  definedNames: ExcelDefinedNameModel[];
}

export interface ExcelRowHeightPatch {
  kind: 'row.height';
  sheetName: string;
  row: number;
  height: number | undefined;
}

export interface ExcelColumnWidthPatch {
  kind: 'column.width';
  sheetName: string;
  column: number;
  width: number | undefined;
}

export interface ExcelSnapshotPatch {
  kind: 'snapshot';
  reason: string;
  beforeModel: ExcelWorkbookModel;
  afterModel: ExcelWorkbookModel;
}

export interface ExcelHistoryApplyResult {
  group: ExcelTransactionGroup;
  affectedSheets: string[];
  affectedRanges: string[];
  summary: ExcelTransactionGroupSummary;
}

export interface ExcelPatchSummary {
  kind: ExcelPatch['kind'];
  direction: 'forward' | 'inverse';
  sheetName: string;
  address: string;
  snapshot: boolean;
  structural: boolean;
  restoreStrategy?: 'delta' | 'full';
}

export interface ExcelTransactionSummary {
  id: string;
  label: string;
  forwardPatchKinds: ExcelPatch['kind'][];
  inversePatchKinds: ExcelPatch['kind'][];
  patches: ExcelPatchSummary[];
}

export interface ExcelTransactionGroupSummary {
  id: string;
  source: ExcelTransactionSource;
  label: string;
  metadata?: ExcelTransactionGroupMetadata;
  mergeKey?: string;
  workbookId: string;
  transactionCount: number;
  patchCount: number;
  forwardPatchKinds: ExcelPatch['kind'][];
  inversePatchKinds: ExcelPatch['kind'][];
  affectedSheets: string[];
  affectedRanges: string[];
  containsSnapshot: boolean;
  containsStructuralPatch: boolean;
  restoreStrategies: ('delta' | 'full')[];
  transactions: ExcelTransactionSummary[];
  createdAt: number;
  committedAt: number;
}

export class ExcelWorkbookHistory {
  private readonly undoStack: ExcelTransactionGroup[] = [];
  private readonly redoStack: ExcelTransactionGroup[] = [];

  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly options: { limit?: number } = {},
  ) {}

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undoGroups(): ExcelTransactionGroup[] {
    return this.undoStack.map(cloneTransactionGroup);
  }

  redoGroups(): ExcelTransactionGroup[] {
    return this.redoStack.map(cloneTransactionGroup);
  }

  async runGroup(
    options: ExcelTransactionGroupOptions,
    run: (transaction: ExcelTransactionBuilder) => Promise<void>,
  ): Promise<ExcelHistoryApplyResult | null> {
    const builder = new ExcelTransactionGroupBuilder(this.workbook, options);
    await run(builder);
    const group = await builder.commit();
    if (!group) return null;
    const committedGroup = this.pushUndo(group, { allowMerge: true });
    this.redoStack.length = 0;
    return {
      group: cloneTransactionGroup(committedGroup),
      affectedSheets: [...committedGroup.affectedSheets],
      affectedRanges: [...committedGroup.affectedRanges],
      summary: summarizeExcelTransactionGroup(committedGroup),
    };
  }

  async runAgentGroup(
    options: ExcelAgentTransactionGroupOptions,
    run: (transaction: ExcelTransactionBuilder) => Promise<void>,
  ): Promise<ExcelHistoryApplyResult | null> {
    return this.runGroup(
      {
        id: options.id,
        source: 'agent',
        label: options.label,
        mergeKey: options.mergeKey,
        mergeWindowMs: options.mergeWindowMs,
        metadata: normalizeTransactionMetadata({
          ...options.metadata,
          actor: options.actor ?? options.metadata?.actor,
          description: options.description ?? options.metadata?.description,
          agentRunId: options.agentRunId ?? options.metadata?.agentRunId,
          agentStepId: options.agentStepId ?? options.metadata?.agentStepId,
          agentActionId: options.agentActionId ?? options.metadata?.agentActionId,
          mergeKey: options.mergeKey ?? options.metadata?.mergeKey,
        }),
      },
      run,
    );
  }

  async undo(): Promise<ExcelHistoryApplyResult | null> {
    const group = this.undoStack.pop();
    if (!group) return null;
    for (const transaction of [...group.transactions].reverse()) {
      for (const patch of [...transaction.inverse].reverse()) await applyExcelPatch(this.workbook, patch);
    }
    this.redoStack.push(group);
    return {
      group: cloneTransactionGroup(group),
      affectedSheets: [...group.affectedSheets],
      affectedRanges: [...group.affectedRanges],
      summary: summarizeExcelTransactionGroup(group),
    };
  }

  async redo(): Promise<ExcelHistoryApplyResult | null> {
    const group = this.redoStack.pop();
    if (!group) return null;
    for (const transaction of group.transactions) {
      for (const patch of transaction.forward) await applyExcelPatch(this.workbook, patch);
    }
    this.pushUndo(group, { allowMerge: false });
    return {
      group: cloneTransactionGroup(group),
      affectedSheets: [...group.affectedSheets],
      affectedRanges: [...group.affectedRanges],
      summary: summarizeExcelTransactionGroup(group),
    };
  }

  clear(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
  }

  private pushUndo(
    group: ExcelTransactionGroup,
    options: { allowMerge: boolean } = { allowMerge: false },
  ): ExcelTransactionGroup {
    if (options.allowMerge) {
      const previous = this.undoStack[this.undoStack.length - 1];
      if (previous && canMergeTransactionGroups(previous, group)) {
        const merged = mergeTransactionGroups(previous, group);
        this.undoStack[this.undoStack.length - 1] = merged;
        return merged;
      }
    }
    this.undoStack.push(group);
    const limit = this.options.limit ?? 100;
    while (this.undoStack.length > limit) this.undoStack.shift();
    return group;
  }
}

export interface ExcelTransactionBuilder {
  setCellValue(sheetName: string, address: string, value: ExcelCellValue): Promise<void>;
  setCellFormula(sheetName: string, address: string, formula: string | null): Promise<void>;
  setRangeValues(sheetName: string, address: string, values: ExcelCellValue[][]): Promise<void>;
  clearRange(sheetName: string, address: string): Promise<void>;
  setRangeBold(sheetName: string, address: string, value: boolean | undefined): Promise<void>;
  setRangeItalic(sheetName: string, address: string, value: boolean | undefined): Promise<void>;
  setRangeUnderline(sheetName: string, address: string, value: boolean | undefined): Promise<void>;
  setRangeFontColor(sheetName: string, address: string, color: ExcelColorModel | undefined): Promise<void>;
  setRangeFillColor(sheetName: string, address: string, color: ExcelColorModel | undefined): Promise<void>;
  setRangeHorizontalAlign(sheetName: string, address: string, horizontal: 'left' | 'center' | 'right' | undefined): Promise<void>;
  setRangeVerticalAlign(sheetName: string, address: string, vertical: 'top' | 'center' | 'bottom' | undefined): Promise<void>;
  setRangeNumberFormat(sheetName: string, address: string, code: string | undefined): Promise<void>;
  clearRangeFormat(sheetName: string, address: string): Promise<void>;
  setRangeStyleId(sheetName: string, address: string, styleId: number | undefined): Promise<void>;
  setSheetViews(sheetName: string, sheetViews: ExcelSheetViewModel[]): Promise<void>;
  addSheet(name: string): Promise<void>;
  deleteSheet(nameOrId: string): Promise<void>;
  renameSheet(nameOrId: string, name: string): Promise<void>;
  copySheet(nameOrId: string, newName: string): Promise<void>;
  moveSheet(nameOrId: string, toIndex: number): Promise<void>;
  mergeRange(sheetName: string, address: string): Promise<void>;
  unmergeRange(sheetName: string, address: string): Promise<void>;
  insertRows(sheetName: string, rowIndex: number, count?: number): Promise<void>;
  deleteRows(sheetName: string, rowIndex: number, count?: number): Promise<void>;
  insertColumns(sheetName: string, columnIndex: number, count?: number): Promise<void>;
  deleteColumns(sheetName: string, columnIndex: number, count?: number): Promise<void>;
  setRangeBorder(sheetName: string, address: string, border: ExcelBorderModel): Promise<void>;
  setRowHeight(sheetName: string, row: number, height: number | undefined): Promise<void>;
  setColumnWidth(sheetName: string, column: number, width: number | undefined): Promise<void>;
  runSnapshot(label: string, reason: string, mutate: () => Promise<void>): Promise<void>;
}

class ExcelTransactionGroupBuilder implements ExcelTransactionBuilder {
  private readonly createdAt = Date.now();
  private readonly transactions: ExcelTransaction[] = [];

  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly options: ExcelTransactionGroupOptions,
  ) {}

  async setCellValue(sheetName: string, address: string, value: ExcelCellValue): Promise<void> {
    await this.setRangeValues(sheetName, address, [[value]]);
  }

  async setCellFormula(sheetName: string, address: string, formula: string | null): Promise<void> {
    await this.recordRangeCells(sheetName, address, 'Set cell formula', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).setFormulas([[formula]]);
    });
  }

  async setRangeValues(sheetName: string, address: string, values: ExcelCellValue[][]): Promise<void> {
    await this.recordRangeCells(sheetName, address, 'Set range values', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).setValues(values);
    });
  }

  async clearRange(sheetName: string, address: string): Promise<void> {
    await this.recordRangeCells(sheetName, address, 'Clear range', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).clear();
    });
  }

  async setRangeBold(sheetName: string, address: string, value: boolean | undefined): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range bold', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.font.setBold(value);
    });
  }

  async setRangeItalic(sheetName: string, address: string, value: boolean | undefined): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range italic', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.font.setItalic(value);
    });
  }

  async setRangeUnderline(sheetName: string, address: string, value: boolean | undefined): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range underline', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.font.setUnderline(value);
    });
  }

  async setRangeFontColor(sheetName: string, address: string, color: ExcelColorModel | undefined): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range font color', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.font.setColor(color);
    });
  }

  async setRangeFillColor(sheetName: string, address: string, color: ExcelColorModel | undefined): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range fill color', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.fill.setForegroundColor(color);
    });
  }

  async setRangeHorizontalAlign(
    sheetName: string,
    address: string,
    horizontal: 'left' | 'center' | 'right' | undefined,
  ): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range horizontal align', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.alignment.setHorizontal(horizontal);
    });
  }

  async setRangeVerticalAlign(
    sheetName: string,
    address: string,
    vertical: 'top' | 'center' | 'bottom' | undefined,
  ): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range vertical align', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.alignment.setVertical(vertical);
    });
  }

  async setRangeNumberFormat(sheetName: string, address: string, code: string | undefined): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range number format', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.setNumberFormatCode(code);
    });
  }

  async clearRangeFormat(sheetName: string, address: string): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Clear range format', async () => {
      await this.workbook.executeCode<void>(
        `
        const sheet = workbook.worksheets.getItem(args.sheetName);
        const range = sheet.getRange(args.address);
        range.format.font.setBold(undefined);
        range.format.font.setItalic(undefined);
        range.format.font.setUnderline(undefined);
        range.format.font.setColor(undefined);
        range.format.fill.setForegroundColor(undefined);
        range.format.border.setValue({});
        range.format.alignment.setValue(undefined);
        range.format.setNumberFormatCode(undefined);
        `,
        { sheetName, address },
      );
    });
  }

  async setRangeStyleId(sheetName: string, address: string, styleId: number | undefined): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range style', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).setStyleId(styleId);
    });
  }

  async setSheetViews(sheetName: string, sheetViews: ExcelSheetViewModel[]): Promise<void> {
    const sheet = this.workbook.worksheets.getItem(sheetName);
    const before = await sheet.sheetViews();
    await sheet.setSheetViews(sheetViews);
    const after = await sheet.sheetViews();
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    const address = 'views';
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label: 'Set sheet views',
      forward: [{ kind: 'sheet.views', sheetName, address, sheetViews: after }],
      inverse: [{ kind: 'sheet.views', sheetName, address, sheetViews: before }],
    });
  }

  async addSheet(name: string): Promise<void> {
    await this.recordWorkbookSheets('Add worksheet', name, async () => {
      await this.workbook.worksheets.add(name);
    });
  }

  async deleteSheet(nameOrId: string): Promise<void> {
    await this.recordWorkbookSheets('Delete worksheet', nameOrId, async () => {
      await this.workbook.worksheets.deleteItem(nameOrId);
    });
  }

  async renameSheet(nameOrId: string, name: string): Promise<void> {
    await this.recordWorkbookSheets('Rename worksheet', nameOrId, async () => {
      await this.workbook.worksheets.getItem(nameOrId).rename(name);
    });
  }

  async copySheet(nameOrId: string, newName: string): Promise<void> {
    await this.recordWorkbookSheets('Copy worksheet', nameOrId, async () => {
      await this.workbook.worksheets.copy(nameOrId, newName);
    });
  }

  async moveSheet(nameOrId: string, toIndex: number): Promise<void> {
    await this.recordWorkbookSheets('Move worksheet', nameOrId, async () => {
      await this.workbook.worksheets.move(nameOrId, toIndex);
    });
  }

  async mergeRange(sheetName: string, address: string): Promise<void> {
    const beforeCells = await captureRangeCells(this.workbook, sheetName, address);
    const beforeMergedRanges = await captureMergedRanges(this.workbook, sheetName);
    await this.workbook.worksheets.getItem(sheetName).getRange(address).merge();
    const afterCells = await captureRangeCells(this.workbook, sheetName, address);
    const afterMergedRanges = await captureMergedRanges(this.workbook, sheetName);
    if (
      JSON.stringify(beforeCells) === JSON.stringify(afterCells) &&
      JSON.stringify(beforeMergedRanges) === JSON.stringify(afterMergedRanges)
    ) {
      return;
    }
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label: 'Merge range',
      forward: [
        { kind: 'range.cells', sheetName, address, cells: afterCells },
        { kind: 'sheet.mergedRanges', sheetName, address, mergedRanges: afterMergedRanges },
      ],
      inverse: [
        { kind: 'range.cells', sheetName, address, cells: beforeCells },
        { kind: 'sheet.mergedRanges', sheetName, address, mergedRanges: beforeMergedRanges },
      ],
    });
  }

  async unmergeRange(sheetName: string, address: string): Promise<void> {
    await this.recordMergedRanges(sheetName, address, 'Unmerge range', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).unmerge();
    });
  }

  async insertRows(sheetName: string, rowIndex: number, count = 1): Promise<void> {
    await this.workbook.worksheets.getItem(sheetName).insertRows(rowIndex, count);
    const address = `rows:${rowIndex}:${count}`;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label: 'Insert rows',
      forward: [{ kind: 'sheet.rows.insert', sheetName, address, rowIndex, count }],
      inverse: [{ kind: 'sheet.rows.delete', sheetName, address, rowIndex, count }],
    });
  }

  async deleteRows(sheetName: string, rowIndex: number, count = 1): Promise<void> {
    const before = await captureSheetModel(this.workbook, sheetName);
    const restoreDelta = canUseStructuralRestoreDelta(before)
      ? buildRowsRestoreDeltaPatch(sheetName, `rows:${rowIndex}:${count}`, rowIndex, count, before)
      : null;
    await this.workbook.worksheets.getItem(sheetName).deleteRows(rowIndex, count);
    const after = await captureSheetModel(this.workbook, sheetName);
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    const address = `rows:${rowIndex}:${count}`;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label: 'Delete rows',
      forward: [{ kind: 'sheet.rows.delete', sheetName, address, rowIndex, count }],
      inverse: [
        restoreDelta ?? {
          kind: 'sheet.rows.restore',
          sheetName,
          address,
          rowIndex,
          count,
          sheet: before.sheet,
          definedNames: before.definedNames,
        },
      ],
    });
  }

  async insertColumns(sheetName: string, columnIndex: number, count = 1): Promise<void> {
    await this.workbook.worksheets.getItem(sheetName).insertColumns(columnIndex, count);
    const address = `columns:${columnIndex}:${count}`;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label: 'Insert columns',
      forward: [{ kind: 'sheet.columns.insert', sheetName, address, columnIndex, count }],
      inverse: [{ kind: 'sheet.columns.delete', sheetName, address, columnIndex, count }],
    });
  }

  async deleteColumns(sheetName: string, columnIndex: number, count = 1): Promise<void> {
    const before = await captureSheetModel(this.workbook, sheetName);
    const restoreDelta = canUseStructuralRestoreDelta(before)
      ? buildColumnsRestoreDeltaPatch(sheetName, `columns:${columnIndex}:${count}`, columnIndex, count, before)
      : null;
    await this.workbook.worksheets.getItem(sheetName).deleteColumns(columnIndex, count);
    const after = await captureSheetModel(this.workbook, sheetName);
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    const address = `columns:${columnIndex}:${count}`;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label: 'Delete columns',
      forward: [{ kind: 'sheet.columns.delete', sheetName, address, columnIndex, count }],
      inverse: [
        restoreDelta ?? {
          kind: 'sheet.columns.restore',
          sheetName,
          address,
          columnIndex,
          count,
          sheet: before.sheet,
          definedNames: before.definedNames,
        },
      ],
    });
  }

  async setRangeBorder(sheetName: string, address: string, border: ExcelBorderModel): Promise<void> {
    await this.recordRangeFormat(sheetName, address, 'Set range border', async () => {
      await this.workbook.worksheets.getItem(sheetName).getRange(address).format.border.setValue(border);
    });
  }

  async setRowHeight(sheetName: string, row: number, height: number | undefined): Promise<void> {
    const address = formatCellAddress(row, 1);
    const range = this.workbook.worksheets.getItem(sheetName).getRange(address);
    const before = (await range.rowHeights())[0];
    await range.setRowHeight(height);
    const after = (await range.rowHeights())[0];
    if (before === after) return;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label: 'Set row height',
      forward: [{ kind: 'row.height', sheetName, row, height: after }],
      inverse: [{ kind: 'row.height', sheetName, row, height: before }],
    });
  }

  async setColumnWidth(sheetName: string, column: number, width: number | undefined): Promise<void> {
    const address = formatCellAddress(1, column);
    const range = this.workbook.worksheets.getItem(sheetName).getRange(address);
    const before = (await range.columnWidths())[0];
    await range.setColumnWidth(width);
    const after = (await range.columnWidths())[0];
    if (before === after) return;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label: 'Set column width',
      forward: [{ kind: 'column.width', sheetName, column, width: after }],
      inverse: [{ kind: 'column.width', sheetName, column, width: before }],
    });
  }

  async runSnapshot(label: string, reason: string, mutate: () => Promise<void>): Promise<void> {
    const before = await this.workbook.toModel();
    await mutate();
    const after = await this.workbook.toModel();
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label,
      forward: [{ kind: 'snapshot', reason, beforeModel: before, afterModel: after }],
      inverse: [{ kind: 'snapshot', reason, beforeModel: after, afterModel: before }],
    });
  }

  async commit(): Promise<ExcelTransactionGroup | null> {
    if (this.transactions.length === 0) return null;
    const affectedSheets = unique(this.transactions.flatMap((transaction) => transaction.forward.map(patchSheetName)));
    const affectedRanges = unique(this.transactions.flatMap((transaction) => transaction.forward.map(patchRangeAddress).filter(Boolean)));
    return {
      id: this.options.id ?? createExcelHistoryId('group'),
      source: this.options.source,
      label: this.options.label,
      metadata: normalizeTransactionMetadata({
        ...this.options.metadata,
        mergeKey: this.options.mergeKey ?? this.options.metadata?.mergeKey,
      }),
      mergeKey: this.options.mergeKey ?? this.options.metadata?.mergeKey,
      mergeWindowMs: this.options.mergeWindowMs,
      workbookId: this.workbook.workbookId,
      transactions: this.transactions.map(cloneTransaction),
      affectedSheets,
      affectedRanges,
      createdAt: this.createdAt,
      committedAt: Date.now(),
    };
  }

  private async recordRangeCells(
    sheetName: string,
    address: string,
    label: string,
    mutate: () => Promise<void>,
  ): Promise<void> {
    const before = await captureRangeCells(this.workbook, sheetName, address);
    await mutate();
    const after = await captureRangeCells(this.workbook, sheetName, address);
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label,
      forward: [{ kind: 'range.cells', sheetName, address, cells: after }],
      inverse: [{ kind: 'range.cells', sheetName, address, cells: before }],
    });
  }

  private async recordRangeFormat(
    sheetName: string,
    address: string,
    label: string,
    mutate: () => Promise<void>,
  ): Promise<void> {
    const before = await captureRangeFormats(this.workbook, sheetName, address);
    await mutate();
    const after = await captureRangeFormats(this.workbook, sheetName, address);
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label,
      forward: [{ kind: 'range.format', sheetName, address, formats: after }],
      inverse: [{ kind: 'range.format', sheetName, address, formats: before }],
    });
  }

  private async recordMergedRanges(
    sheetName: string,
    address: string,
    label: string,
    mutate: () => Promise<void>,
  ): Promise<void> {
    const before = await captureMergedRanges(this.workbook, sheetName);
    await mutate();
    const after = await captureMergedRanges(this.workbook, sheetName);
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label,
      forward: [{ kind: 'sheet.mergedRanges', sheetName, address, mergedRanges: after }],
      inverse: [{ kind: 'sheet.mergedRanges', sheetName, address, mergedRanges: before }],
    });
  }

  private async recordSheetModel(
    sheetName: string,
    address: string,
    label: string,
    mutate: () => Promise<void>,
  ): Promise<void> {
    const before = await captureSheetModel(this.workbook, sheetName);
    await mutate();
    const after = await captureSheetModel(this.workbook, sheetName);
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label,
      forward: [{ kind: 'sheet.model', sheetName, address, sheet: after.sheet, definedNames: after.definedNames }],
      inverse: [{ kind: 'sheet.model', sheetName, address, sheet: before.sheet, definedNames: before.definedNames }],
    });
  }

  private async recordWorkbookSheets(
    label: string,
    sheetName: string,
    mutate: () => Promise<void>,
  ): Promise<void> {
    const before = await captureWorkbookSheets(this.workbook);
    await mutate();
    const after = await captureWorkbookSheets(this.workbook);
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    const address = 'sheets';
    this.transactions.push({
      id: createExcelHistoryId('tx'),
      label,
      forward: [{ kind: 'workbook.sheets', sheetName, address, sheets: after.sheets, definedNames: after.definedNames }],
      inverse: [{ kind: 'workbook.sheets', sheetName, address, sheets: before.sheets, definedNames: before.definedNames }],
    });
  }
}

export async function applyExcelPatch(workbook: ExcelWorkbook, patch: ExcelPatch): Promise<void> {
  if (patch.kind === 'range.cells') {
    await applyRangeCellsPatch(workbook, patch);
    return;
  }
  if (patch.kind === 'range.format') {
    await applyRangeFormatPatch(workbook, patch);
    return;
  }
  if (patch.kind === 'sheet.mergedRanges') {
    await applyMergedRangesPatch(workbook, patch);
    return;
  }
  if (patch.kind === 'sheet.views') {
    await workbook.worksheets.getItem(patch.sheetName).setSheetViews(patch.sheetViews);
    return;
  }
  if (patch.kind === 'workbook.sheets') {
    await applyWorkbookSheetsPatch(workbook, patch);
    return;
  }
  if (patch.kind === 'sheet.model') {
    await applySheetModelPatch(workbook, patch);
    return;
  }
  if (patch.kind === 'sheet.rows.insert') {
    await workbook.worksheets.getItem(patch.sheetName).insertRows(patch.rowIndex, patch.count);
    return;
  }
  if (patch.kind === 'sheet.rows.delete') {
    await workbook.worksheets.getItem(patch.sheetName).deleteRows(patch.rowIndex, patch.count);
    return;
  }
  if (patch.kind === 'sheet.rows.restore') {
    await applySheetStructuralRestorePatch(workbook, patch);
    return;
  }
  if (patch.kind === 'sheet.rows.restoreDelta') {
    await applyRowsRestoreDeltaPatch(workbook, patch);
    return;
  }
  if (patch.kind === 'sheet.columns.insert') {
    await workbook.worksheets.getItem(patch.sheetName).insertColumns(patch.columnIndex, patch.count);
    return;
  }
  if (patch.kind === 'sheet.columns.delete') {
    await workbook.worksheets.getItem(patch.sheetName).deleteColumns(patch.columnIndex, patch.count);
    return;
  }
  if (patch.kind === 'sheet.columns.restore') {
    await applySheetStructuralRestorePatch(workbook, patch);
    return;
  }
  if (patch.kind === 'sheet.columns.restoreDelta') {
    await applyColumnsRestoreDeltaPatch(workbook, patch);
    return;
  }
  if (patch.kind === 'row.height') {
    await workbook.worksheets.getItem(patch.sheetName).getRange(formatCellAddress(patch.row, 1)).setRowHeight(patch.height);
    return;
  }
  if (patch.kind === 'column.width') {
    await workbook.worksheets.getItem(patch.sheetName).getRange(formatCellAddress(1, patch.column)).setColumnWidth(patch.width);
    return;
  }
  await applySnapshotPatch(workbook, patch);
}

export function summarizeExcelTransactionGroup(group: ExcelTransactionGroup): ExcelTransactionGroupSummary {
  const transactions = group.transactions.map(summarizeExcelTransaction);
  const patches = transactions.flatMap((transaction) => transaction.patches);
  const restoreStrategies = unique(
    patches
      .map((patch) => patch.restoreStrategy)
      .filter((strategy): strategy is 'delta' | 'full' => strategy !== undefined),
  ) as ('delta' | 'full')[];
  return {
    id: group.id,
    source: group.source,
    label: group.label,
    metadata: group.metadata ? cloneJson(group.metadata) : undefined,
    mergeKey: group.mergeKey,
    workbookId: group.workbookId,
    transactionCount: group.transactions.length,
    patchCount: patches.length,
    forwardPatchKinds: unique(
      group.transactions.flatMap((transaction) => transaction.forward.map((patch) => patch.kind)),
    ) as ExcelPatch['kind'][],
    inversePatchKinds: unique(
      group.transactions.flatMap((transaction) => transaction.inverse.map((patch) => patch.kind)),
    ) as ExcelPatch['kind'][],
    affectedSheets: [...group.affectedSheets],
    affectedRanges: [...group.affectedRanges],
    containsSnapshot: patches.some((patch) => patch.snapshot),
    containsStructuralPatch: patches.some((patch) => patch.structural),
    restoreStrategies,
    transactions,
    createdAt: group.createdAt,
    committedAt: group.committedAt,
  };
}

function summarizeExcelTransaction(transaction: ExcelTransaction): ExcelTransactionSummary {
  return {
    id: transaction.id,
    label: transaction.label,
    forwardPatchKinds: unique(transaction.forward.map((patch) => patch.kind)) as ExcelPatch['kind'][],
    inversePatchKinds: unique(transaction.inverse.map((patch) => patch.kind)) as ExcelPatch['kind'][],
    patches: [
      ...transaction.forward.map((patch) => summarizeExcelPatch(patch, 'forward')),
      ...transaction.inverse.map((patch) => summarizeExcelPatch(patch, 'inverse')),
    ],
  };
}

function summarizeExcelPatch(patch: ExcelPatch, direction: 'forward' | 'inverse'): ExcelPatchSummary {
  return {
    kind: patch.kind,
    direction,
    sheetName: patchSheetName(patch),
    address: patchRangeAddress(patch),
    snapshot: patch.kind === 'snapshot',
    structural: isStructuralPatch(patch),
    restoreStrategy: patchRestoreStrategy(patch),
  };
}

async function captureMergedRanges(workbook: ExcelWorkbook, sheetName: string): Promise<string[]> {
  return (await workbook.worksheets.getItem(sheetName).mergedRanges()).map((range) => range.address);
}

async function captureSheetModel(
  workbook: ExcelWorkbook,
  sheetName: string,
): Promise<{ sheet: ExcelSheetModel; definedNames: ExcelDefinedNameModel[] }> {
  const model = await workbook.toModel();
  const sheet = model.sheets.find((item) => item.name === sheetName || item.id === sheetName);
  if (!sheet) throw new Error(`ExcelWorkbookHistory: sheet not found: ${sheetName}`);
  return {
    sheet,
    definedNames: model.definedNames,
  };
}

async function captureWorkbookSheets(
  workbook: ExcelWorkbook,
): Promise<{ sheets: ExcelSheetModel[]; definedNames: ExcelDefinedNameModel[] }> {
  const model = await workbook.toModel();
  return { sheets: model.sheets, definedNames: model.definedNames };
}

async function captureRangeCells(
  workbook: ExcelWorkbook,
  sheetName: string,
  address: string,
): Promise<ExcelCellModel[]> {
  return workbook.worksheets.getItem(sheetName).getRange(address).toModel();
}

async function captureRangeFormats(
  workbook: ExcelWorkbook,
  sheetName: string,
  address: string,
): Promise<ExcelCellFormatSnapshot[]> {
  const range = parseRangeAddress(address);
  const styleIds = await workbook.worksheets.getItem(sheetName).getRange(address).styleIds();
  const formats: ExcelCellFormatSnapshot[] = [];
  for (let rowOffset = 0; rowOffset < range.rowCount; rowOffset++) {
    for (let columnOffset = 0; columnOffset < range.columnCount; columnOffset++) {
      formats.push({
        row: range.startRow + rowOffset,
        column: range.startColumn + columnOffset,
        styleId: styleIds[rowOffset]?.[columnOffset],
      });
    }
  }
  return formats;
}

async function applyRangeCellsPatch(workbook: ExcelWorkbook, patch: ExcelRangeCellsPatch): Promise<void> {
  await workbook.executeCode<void>(
    `
    const sheet = workbook.worksheets.getItem(args.sheetName);
    const range = sheet.getRange(args.address);
    range.clear();
    for (const cell of args.cells) {
      const target = sheet.getCell(cell.address);
      if (cell.formula !== undefined) target.setFormula(cell.formula);
      else target.setValue(cell.value);
      if (cell.styleId !== undefined) target.setStyleId(cell.styleId);
    }
    `,
    { sheetName: patch.sheetName, address: patch.address, cells: patch.cells },
  );
}

async function applyRangeFormatPatch(workbook: ExcelWorkbook, patch: ExcelRangeFormatPatch): Promise<void> {
  const formats = patch.formats.map((item) => ({
    address: formatCellAddress(item.row, item.column),
    styleId: item.styleId,
  }));
  await workbook.executeCode<void>(
    `
    const sheet = workbook.worksheets.getItem(args.sheetName);
    for (const item of args.formats) {
      sheet.getCell(item.address).setStyleId(item.styleId);
    }
    `,
    { sheetName: patch.sheetName, formats },
  );
}

async function applyMergedRangesPatch(workbook: ExcelWorkbook, patch: ExcelMergedRangesPatch): Promise<void> {
  await workbook.executeCode<void>(
    `
    const sheet = workbook.worksheets.getItem(args.sheetName);
    sheet.model.mergedRanges = args.mergedRanges.map((address) => ({ address }));
    `,
    { sheetName: patch.sheetName, mergedRanges: patch.mergedRanges },
  );
}

async function applySheetModelPatch(workbook: ExcelWorkbook, patch: ExcelSheetModelPatch): Promise<void> {
  await workbook.executeCode<void>(
    `
    const model = workbook.toModel();
    const index = model.sheets.findIndex((sheet) => sheet.name === args.sheetName || sheet.id === args.sheetName);
    if (index === -1) throw new Error('ExcelWorkbookHistory: sheet not found: ' + args.sheetName);
    model.sheets[index] = args.sheet;
    model.definedNames = args.definedNames;
    workbook._replaceModelForHistory(model);
    `,
    {
      sheetName: patch.sheetName,
      sheet: patch.sheet,
      definedNames: patch.definedNames,
    },
  );
}

async function applyWorkbookSheetsPatch(workbook: ExcelWorkbook, patch: ExcelWorkbookSheetsPatch): Promise<void> {
  await workbook.executeCode<void>(
    `
    const model = workbook.toModel();
    model.sheets = args.sheets;
    model.definedNames = args.definedNames;
    workbook._replaceModelForHistory(model);
    `,
    { sheets: patch.sheets, definedNames: patch.definedNames },
  );
}

async function applySheetStructuralRestorePatch(
  workbook: ExcelWorkbook,
  patch: ExcelSheetRowsRestorePatch | ExcelSheetColumnsRestorePatch,
): Promise<void> {
  await workbook.executeCode<void>(
    `
    const model = workbook.toModel();
    const index = model.sheets.findIndex((sheet) => sheet.name === args.sheetName || sheet.id === args.sheetName);
    if (index === -1) throw new Error('ExcelWorkbookHistory: sheet not found: ' + args.sheetName);
    model.sheets[index] = args.sheet;
    model.definedNames = args.definedNames;
    workbook._replaceModelForHistory(model);
    `,
    {
      sheetName: patch.sheetName,
      sheet: patch.sheet,
      definedNames: patch.definedNames,
    },
  );
}

async function applyRowsRestoreDeltaPatch(
  workbook: ExcelWorkbook,
  patch: ExcelSheetRowsRestoreDeltaPatch,
): Promise<void> {
  await workbook.executeCode<void>(
    `
    function columnIndexToName(column) {
      let value = column;
      let name = '';
      while (value > 0) {
        const rem = (value - 1) % 26;
        name = String.fromCharCode(65 + rem) + name;
        value = Math.floor((value - 1) / 26);
      }
      return name;
    }

    function formatCellAddress(row, column) {
      return columnIndexToName(column) + String(row);
    }

    function computeDimension(cells) {
      if (!cells.length) return undefined;
      let minRow = Infinity;
      let minColumn = Infinity;
      let maxRow = 0;
      let maxColumn = 0;
      for (const cell of cells) {
        minRow = Math.min(minRow, cell.row);
        minColumn = Math.min(minColumn, cell.column);
        maxRow = Math.max(maxRow, cell.row);
        maxColumn = Math.max(maxColumn, cell.column);
      }
      const first = formatCellAddress(minRow, minColumn);
      const last = formatCellAddress(maxRow, maxColumn);
      return first === last ? first : first + ':' + last;
    }

    const sheet = workbook.worksheets.getItem(args.sheetName);
    sheet.insertRows(args.rowIndex, args.count);

    const model = workbook.toModel();
    const index = model.sheets.findIndex((item) => item.name === args.sheetName || item.id === args.sheetName);
    if (index === -1) throw new Error('ExcelWorkbookHistory: sheet not found: ' + args.sheetName);
    const sheetModel = model.sheets[index];
    const endRow = args.rowIndex + args.count - 1;

    sheetModel.cells = sheetModel.cells.filter((cell) => cell.row < args.rowIndex || cell.row > endRow);
    sheetModel.cells.push(...args.cells.map((cell) => ({ ...cell })));
    sheetModel.cells.sort((a, b) => a.row - b.row || a.column - b.column);

    sheetModel.rows = sheetModel.rows.filter((row) => row.index < args.rowIndex || row.index > endRow);
    sheetModel.rows.push(...args.rows.map((row) => ({ ...row })));
    sheetModel.rows.sort((a, b) => a.index - b.index);

    sheetModel.mergedRanges = args.mergedRanges.map((range) => ({ ...range }));
    sheetModel.dimension = computeDimension(sheetModel.cells);
    model.definedNames = args.definedNames.map((name) => ({ ...name }));
    workbook._replaceModelForHistory(model);
    `,
    {
      sheetName: patch.sheetName,
      rowIndex: patch.rowIndex,
      count: patch.count,
      cells: patch.cells,
      rows: patch.rows,
      mergedRanges: patch.mergedRanges,
      definedNames: patch.definedNames,
    },
  );
}

async function applyColumnsRestoreDeltaPatch(
  workbook: ExcelWorkbook,
  patch: ExcelSheetColumnsRestoreDeltaPatch,
): Promise<void> {
  await workbook.executeCode<void>(
    `
    function columnIndexToName(column) {
      let value = column;
      let name = '';
      while (value > 0) {
        const rem = (value - 1) % 26;
        name = String.fromCharCode(65 + rem) + name;
        value = Math.floor((value - 1) / 26);
      }
      return name;
    }

    function formatCellAddress(row, column) {
      return columnIndexToName(column) + String(row);
    }

    function computeDimension(cells) {
      if (!cells.length) return undefined;
      let minRow = Infinity;
      let minColumn = Infinity;
      let maxRow = 0;
      let maxColumn = 0;
      for (const cell of cells) {
        minRow = Math.min(minRow, cell.row);
        minColumn = Math.min(minColumn, cell.column);
        maxRow = Math.max(maxRow, cell.row);
        maxColumn = Math.max(maxColumn, cell.column);
      }
      const first = formatCellAddress(minRow, minColumn);
      const last = formatCellAddress(maxRow, maxColumn);
      return first === last ? first : first + ':' + last;
    }

    const sheet = workbook.worksheets.getItem(args.sheetName);
    sheet.insertColumns(args.columnIndex, args.count);

    const model = workbook.toModel();
    const index = model.sheets.findIndex((item) => item.name === args.sheetName || item.id === args.sheetName);
    if (index === -1) throw new Error('ExcelWorkbookHistory: sheet not found: ' + args.sheetName);
    const sheetModel = model.sheets[index];
    const endColumn = args.columnIndex + args.count - 1;

    sheetModel.cells = sheetModel.cells.filter((cell) => cell.column < args.columnIndex || cell.column > endColumn);
    sheetModel.cells.push(...args.cells.map((cell) => ({ ...cell })));
    sheetModel.cells.sort((a, b) => a.row - b.row || a.column - b.column);

    sheetModel.columns = args.columns.map((column) => ({ ...column }));
    sheetModel.columns.sort((a, b) => a.min - b.min || a.max - b.max);

    sheetModel.mergedRanges = args.mergedRanges.map((range) => ({ ...range }));
    sheetModel.dimension = computeDimension(sheetModel.cells);
    model.definedNames = args.definedNames.map((name) => ({ ...name }));
    workbook._replaceModelForHistory(model);
    `,
    {
      sheetName: patch.sheetName,
      columnIndex: patch.columnIndex,
      count: patch.count,
      cells: patch.cells,
      columns: patch.columns,
      mergedRanges: patch.mergedRanges,
      definedNames: patch.definedNames,
    },
  );
}

async function applySnapshotPatch(workbook: ExcelWorkbook, patch: ExcelSnapshotPatch): Promise<void> {
  await workbook.executeCode<void>(
    `
    workbook._replaceModelForHistory(args.model);
    `,
    { model: patch.afterModel },
  );
}

function patchSheetName(patch: ExcelPatch): string {
  if (patch.kind === 'snapshot') return '';
  return patch.sheetName;
}

function patchRangeAddress(patch: ExcelPatch): string {
  if (patch.kind === 'row.height') return `${patch.sheetName}!${formatCellAddress(patch.row, 1)}`;
  if (patch.kind === 'column.width') return `${patch.sheetName}!${formatCellAddress(1, patch.column)}`;
  if (patch.kind === 'snapshot') return '';
  if (patch.kind === 'sheet.model') return `${patch.sheetName}!${patch.address}`;
  if (patch.kind === 'sheet.views') return `${patch.sheetName}!${patch.address}`;
  if (patch.kind === 'workbook.sheets') return `${patch.sheetName}!${patch.address}`;
  if (
    patch.kind === 'sheet.rows.insert' ||
    patch.kind === 'sheet.rows.delete' ||
    patch.kind === 'sheet.rows.restore' ||
    patch.kind === 'sheet.rows.restoreDelta' ||
    patch.kind === 'sheet.columns.insert' ||
    patch.kind === 'sheet.columns.delete' ||
    patch.kind === 'sheet.columns.restore' ||
    patch.kind === 'sheet.columns.restoreDelta'
  ) {
    return `${patch.sheetName}!${patch.address}`;
  }
  if (patch.kind === 'sheet.mergedRanges') return `${patch.sheetName}!${patch.address}`;
  return `${patch.sheetName}!${patch.address}`;
}

function isStructuralPatch(patch: ExcelPatch): boolean {
  return (
    patch.kind === 'sheet.model' ||
    patch.kind === 'workbook.sheets' ||
    patch.kind === 'sheet.rows.insert' ||
    patch.kind === 'sheet.rows.delete' ||
    patch.kind === 'sheet.rows.restore' ||
    patch.kind === 'sheet.rows.restoreDelta' ||
    patch.kind === 'sheet.columns.insert' ||
    patch.kind === 'sheet.columns.delete' ||
    patch.kind === 'sheet.columns.restore' ||
    patch.kind === 'sheet.columns.restoreDelta' ||
    patch.kind === 'sheet.mergedRanges' ||
    patch.kind === 'snapshot'
  );
}

function patchRestoreStrategy(patch: ExcelPatch): 'delta' | 'full' | undefined {
  if (patch.kind === 'sheet.rows.restoreDelta' || patch.kind === 'sheet.columns.restoreDelta') return 'delta';
  if (patch.kind === 'sheet.rows.restore' || patch.kind === 'sheet.columns.restore') return 'full';
  if (patch.kind === 'sheet.model' || patch.kind === 'snapshot') return 'full';
  if (patch.kind === 'workbook.sheets') return 'full';
  return undefined;
}

function canMergeTransactionGroups(previous: ExcelTransactionGroup, next: ExcelTransactionGroup): boolean {
  if (!previous.mergeKey || !next.mergeKey) return false;
  if (previous.mergeKey !== next.mergeKey) return false;
  if (previous.source !== next.source) return false;
  if (previous.workbookId !== next.workbookId) return false;
  const windowMs = next.mergeWindowMs ?? previous.mergeWindowMs ?? 0;
  if (windowMs <= 0) return false;
  return next.createdAt - previous.committedAt <= windowMs;
}

function mergeTransactionGroups(
  previous: ExcelTransactionGroup,
  next: ExcelTransactionGroup,
): ExcelTransactionGroup {
  return {
    ...previous,
    label: next.label,
    metadata: normalizeTransactionMetadata({
      ...previous.metadata,
      ...next.metadata,
      mergeKey: next.mergeKey,
    }),
    mergeKey: previous.mergeKey,
    mergeWindowMs: next.mergeWindowMs ?? previous.mergeWindowMs,
    transactions: [...previous.transactions.map(cloneTransaction), ...next.transactions.map(cloneTransaction)],
    affectedSheets: unique([...previous.affectedSheets, ...next.affectedSheets]),
    affectedRanges: unique([...previous.affectedRanges, ...next.affectedRanges]),
    committedAt: next.committedAt,
  };
}

function cloneTransactionGroup(group: ExcelTransactionGroup): ExcelTransactionGroup {
  return JSON.parse(JSON.stringify(group)) as ExcelTransactionGroup;
}

function cloneTransaction(transaction: ExcelTransaction): ExcelTransaction {
  return JSON.parse(JSON.stringify(transaction)) as ExcelTransaction;
}

function buildRowsRestoreDeltaPatch(
  sheetName: string,
  address: string,
  rowIndex: number,
  count: number,
  model: { sheet: ExcelSheetModel; definedNames: ExcelDefinedNameModel[] },
): ExcelSheetRowsRestoreDeltaPatch {
  const endRow = rowIndex + count - 1;
  return {
    kind: 'sheet.rows.restoreDelta',
    sheetName,
    address,
    rowIndex,
    count,
    cells: model.sheet.cells.filter((cell) => cell.row >= rowIndex && cell.row <= endRow).map(cloneJson),
    rows: model.sheet.rows.filter((row) => row.index >= rowIndex && row.index <= endRow).map(cloneJson),
    mergedRanges: model.sheet.mergedRanges.map(cloneJson),
    definedNames: model.definedNames.map(cloneJson),
  };
}

function buildColumnsRestoreDeltaPatch(
  sheetName: string,
  address: string,
  columnIndex: number,
  count: number,
  model: { sheet: ExcelSheetModel; definedNames: ExcelDefinedNameModel[] },
): ExcelSheetColumnsRestoreDeltaPatch {
  const endColumn = columnIndex + count - 1;
  return {
    kind: 'sheet.columns.restoreDelta',
    sheetName,
    address,
    columnIndex,
    count,
    cells: model.sheet.cells.filter((cell) => cell.column >= columnIndex && cell.column <= endColumn).map(cloneJson),
    columns: model.sheet.columns.map(cloneJson),
    mergedRanges: model.sheet.mergedRanges.map(cloneJson),
    definedNames: model.definedNames.map(cloneJson),
  };
}

function canUseStructuralRestoreDelta(model: {
  sheet: ExcelSheetModel;
  definedNames: ExcelDefinedNameModel[];
}): boolean {
  const sheet = model.sheet;
  return (
    model.definedNames.length === 0 &&
    !sheet.cells.some((cell) => cell.formula !== undefined) &&
    sheet.hyperlinks.length === 0 &&
    sheet.tables.length === 0 &&
    sheet.dataValidations.length === 0 &&
    sheet.conditionalFormats.length === 0 &&
    sheet.comments.length === 0 &&
    sheet.threadedComments.length === 0 &&
    sheet.drawings.length === 0 &&
    sheet.charts.length === 0 &&
    sheet.images.length === 0
  );
}

function cloneJson<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function normalizeTransactionMetadata(
  metadata: ExcelTransactionGroupMetadata | undefined,
): ExcelTransactionGroupMetadata | undefined {
  if (!metadata) return undefined;
  const normalized: ExcelTransactionGroupMetadata = {};
  for (const [key, value] of Object.entries(metadata)) {
    if (value !== undefined) normalized[key] = value;
  }
  return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function createExcelHistoryId(prefix: string): string {
  return `${prefix}:${Date.now()}:${Math.random().toString(16).slice(2)}`;
}

function unique(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))];
}
