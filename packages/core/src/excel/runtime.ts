import { ExcelWorkbookCore, type ExcelLoadOptions } from './excel-workbook.js';
import {
  aggregateExcelTable,
  type ExcelAggregateTableOptions,
  type ExcelAggregateTableResult,
} from './analysis.js';
import type {
  ExcelAlignmentModel,
  ExcelBorderModel,
  ExcelCellKind,
  ExcelCellModel,
  ExcelCellValue,
  ExcelChartAddOptions,
  ExcelChartModel,
  ExcelColumnModel,
  ExcelCommentModel,
  ExcelConditionalFormatModel,
  ExcelConditionalFormatRuleModel,
  ExcelDataValidationModel,
  ExcelDrawingAnchorModel,
  ExcelDrawingMarkerModel,
  ExcelDrawingObjectModel,
  ExcelDrawingModel,
  ExcelFillModel,
  ExcelFontModel,
  ExcelHyperlinkModel,
  ExcelImageModel,
  ExcelMergedRangeModel,
  ExcelRowModel,
  ExcelSheetModel,
  ExcelSheetViewModel,
  ExcelShapeModel,
  ExcelStylesModel,
  ExcelTableModel,
  ExcelTableAutoFilterModel,
  ExcelTableSortStateModel,
  ExcelThreadedCommentModel,
  ExcelWorkbookModel,
  ExcelDefinedNameModel,
} from './model.js';
import { executeExcelScript } from './runtime-script.js';
import { ExcelWorkbookHistory } from './history.js';
import {
  createNodeExcelWorkerPort,
  ExcelWorkerTransport,
  type ExcelWorkerPortLike,
} from './runtime-transport.js';
import type { ExcelWorkbookMutation } from './mutations.js';

export type ExcelRuntimeMode = 'auto' | 'worker' | 'inline';

export type ExcelModeSource =
  | 'explicit-inline'
  | 'explicit-worker'
  | 'auto/node-with-worker-config'
  | 'auto/node-no-worker-config'
  | 'auto/browser-factory'
  | 'auto/fallback';

export interface ExcelModeDecision {
  mode: 'worker' | 'inline';
  source: ExcelModeSource;
}

export type ExcelWorkerFactory = () => ExcelWorkerPortLike | Promise<ExcelWorkerPortLike>;

export interface ExcelWorkbookLoadOptions extends ExcelLoadOptions {
  mode?: ExcelRuntimeMode;
  onModeResolved?: (decision: ExcelModeDecision, workbookId: string) => void;
}

interface Defaults {
  mode: ExcelRuntimeMode;
  workerFactory?: ExcelWorkerFactory;
  workerBundleUrl?: string | URL;
}

const defaults: Defaults = {
  mode: 'auto',
};

export class ExcelWorkbook {
  readonly workbookId: string;
  readonly mode: 'worker' | 'inline';
  readonly modeSource: ExcelModeSource;
  readonly worksheets: ExcelWorksheetCollection;
  readonly names: ExcelNamedItemCollection;
  private readonly inlineWorkbook?: ExcelWorkbookCore;
  private readonly workerTransport?: ExcelWorkerTransport;
  private readonly workbookHistory: ExcelWorkbookHistory;
  private disposed = false;

  private constructor(init: {
    workbookId: string;
    mode: 'worker' | 'inline';
    modeSource: ExcelModeSource;
    inlineWorkbook?: ExcelWorkbookCore;
    workerTransport?: ExcelWorkerTransport;
  }) {
    this.workbookId = init.workbookId;
    this.mode = init.mode;
    this.modeSource = init.modeSource;
    this.inlineWorkbook = init.inlineWorkbook;
    this.workerTransport = init.workerTransport;
    this.worksheets = new ExcelWorksheetCollection(this);
    this.names = new ExcelNamedItemCollection(this);
    this.workbookHistory = new ExcelWorkbookHistory(this);
  }

  static async fromZip(
    bytes: Uint8Array | ArrayBuffer,
    opts: ExcelWorkbookLoadOptions = {},
  ): Promise<ExcelWorkbook> {
    const decision = this.detectModeDecision(opts.mode);
    if (decision.mode === 'inline') {
      const workbook = await ExcelWorkbookCore.fromZip(bytes, opts);
      opts.onModeResolved?.(decision, workbook.workbookId);
      return new ExcelWorkbook({
        workbookId: workbook.workbookId,
        mode: 'inline',
        modeSource: decision.source,
        inlineWorkbook: workbook,
      });
    }

    const transport = await this.createWorkerTransport();
    const loaded = await transport.call<{ workbookId: string }>(
      'load.zip',
      [bytes, toCoreLoadOptions(opts)],
      30_000,
    );
    opts.onModeResolved?.(decision, loaded.workbookId);
    return new ExcelWorkbook({
      workbookId: loaded.workbookId,
      mode: 'worker',
      modeSource: decision.source,
      workerTransport: transport,
    });
  }

  static async fromModel(
    model: ExcelWorkbookModel,
    opts: ExcelWorkbookLoadOptions = {},
  ): Promise<ExcelWorkbook> {
    const decision = this.detectModeDecision(opts.mode);
    if (decision.mode === 'inline') {
      const workbook = ExcelWorkbookCore.fromModel(model, opts);
      opts.onModeResolved?.(decision, workbook.workbookId);
      return new ExcelWorkbook({
        workbookId: workbook.workbookId,
        mode: 'inline',
        modeSource: decision.source,
        inlineWorkbook: workbook,
      });
    }

    const transport = await this.createWorkerTransport();
    const loaded = await transport.call<{ workbookId: string }>(
      'load.model',
      [model, toCoreLoadOptions(opts)],
      30_000,
    );
    opts.onModeResolved?.(decision, loaded.workbookId);
    return new ExcelWorkbook({
      workbookId: loaded.workbookId,
      mode: 'worker',
      modeSource: decision.source,
      workerTransport: transport,
    });
  }

  static fromPackage(
    pkg: Parameters<typeof ExcelWorkbookCore.fromPackage>[0],
    opts: ExcelWorkbookLoadOptions = {},
  ): ExcelWorkbook {
    if (opts.mode === 'worker') {
      throw new Error('ExcelWorkbook.fromPackage currently supports inline mode only; use fromZip for worker mode');
    }
    const workbook = ExcelWorkbookCore.fromPackage(pkg, opts);
    const decision: ExcelModeDecision = { mode: 'inline', source: opts.mode === 'inline' ? 'explicit-inline' : 'auto/node-no-worker-config' };
    opts.onModeResolved?.(decision, workbook.workbookId);
    return new ExcelWorkbook({
      workbookId: workbook.workbookId,
      mode: 'inline',
      modeSource: decision.source,
      inlineWorkbook: workbook,
    });
  }

  static setDefaultMode(mode: ExcelRuntimeMode): void {
    defaults.mode = mode;
  }

  static setWorkerFactory(factory: ExcelWorkerFactory | undefined): void {
    defaults.workerFactory = factory;
  }

  static setWorkerBundleUrl(url: string | URL | undefined): void {
    defaults.workerBundleUrl = url;
  }

  static detectMode(): 'worker' | 'inline' {
    return this.detectModeDecision().mode;
  }

  static detectModeDecision(mode: ExcelRuntimeMode = defaults.mode): ExcelModeDecision {
    if (mode === 'inline') return { mode: 'inline', source: 'explicit-inline' };
    if (mode === 'worker') {
      if (!this.hasWorkerConfig()) {
        throw new Error('ExcelWorkbook: worker mode requires workerFactory or workerBundleUrl');
      }
      return { mode: 'worker', source: 'explicit-worker' };
    }
    if (defaults.workerFactory) return { mode: 'worker', source: 'auto/browser-factory' };
    if (defaults.workerBundleUrl) return { mode: 'worker', source: 'auto/node-with-worker-config' };
    return { mode: 'inline', source: 'auto/node-no-worker-config' };
  }

  static _resetDefaults(): void {
    defaults.mode = 'auto';
    defaults.workerFactory = undefined;
    defaults.workerBundleUrl = undefined;
  }

  async executeCode<T = unknown>(
    code: string,
    args?: Record<string, unknown>,
    timeoutMs = 30_000,
    signal?: AbortSignal,
  ): Promise<T> {
    this.assertNotDisposed();
    if (this.inlineWorkbook) {
      return (await executeExcelScript(this.inlineWorkbook, code, args, timeoutMs, signal)) as T;
    }
    if (signal?.aborted) {
      await this.dispose();
      throw new Error('excel worker script aborted before dispatch');
    }

    let abortHandler: (() => void) | undefined;
    const abortPromise =
      signal &&
      new Promise<never>((_, reject) => {
        abortHandler = (): void => {
          signal.removeEventListener('abort', abortHandler!);
          reject(new Error('excel worker script aborted'));
        };
        signal.addEventListener('abort', abortHandler, { once: true });
      });

    try {
      const call = this.workerTransport!.call<T>('script.execute', [code, args, 0], timeoutMs);
      return abortPromise ? await Promise.race([call, abortPromise]) : await call;
    } catch (error) {
      if (isWorkerTerminalScriptError(error)) {
        await this.dispose();
      }
      throw error;
    } finally {
      if (signal && abortHandler) {
        signal.removeEventListener('abort', abortHandler);
      }
    }
  }

  async toModel(): Promise<ExcelWorkbookModel> {
    this.assertNotDisposed();
    if (this.inlineWorkbook) return this.inlineWorkbook.toModel();
    return this.workerTransport!.call<ExcelWorkbookModel>('to.model', [], 30_000);
  }

  async getMutations(): Promise<ExcelWorkbookMutation[]> {
    this.assertNotDisposed();
    if (this.inlineWorkbook) return this.inlineWorkbook.getMutations();
    return this.workerTransport!.call<ExcelWorkbookMutation[]>('mutations.list', [], 30_000);
  }

  async clearMutations(): Promise<void> {
    this.assertNotDisposed();
    if (this.inlineWorkbook) {
      this.inlineWorkbook.clearMutations();
      return;
    }
    await this.workerTransport!.call<void>('mutations.clear', [], 30_000);
  }

  history(): ExcelWorkbookHistory {
    this.assertNotDisposed();
    return this.workbookHistory;
  }

  async aggregateTable(options: ExcelAggregateTableOptions): Promise<ExcelAggregateTableResult> {
    this.assertNotDisposed();
    return aggregateExcelTable(await this.toModel(), options);
  }

  async toZip(): Promise<Uint8Array> {
    this.assertNotDisposed();
    if (this.inlineWorkbook) return this.inlineWorkbook.toZip();
    return this.workerTransport!.call<Uint8Array>('to.zip', [], 60_000);
  }

  async dispose(): Promise<void> {
    if (this.disposed) return;
    this.disposed = true;
    if (this.workerTransport) {
      await this.workerTransport.dispose();
    }
  }

  private static hasWorkerConfig(): boolean {
    return Boolean(defaults.workerFactory || defaults.workerBundleUrl);
  }

  private static async createWorkerTransport(): Promise<ExcelWorkerTransport> {
    const port = defaults.workerFactory
      ? await defaults.workerFactory()
      : await createNodeExcelWorkerPort(defaults.workerBundleUrl!);
    return new ExcelWorkerTransport(port);
  }

  private assertNotDisposed(): void {
    if (this.disposed) throw new Error('ExcelWorkbook has been disposed');
  }
}

export { ExcelWorkbook as ExcelRuntimeWorkbook };
export type ExcelRuntimeLoadOptions = ExcelWorkbookLoadOptions;

function toCoreLoadOptions(opts: ExcelWorkbookLoadOptions): ExcelLoadOptions {
  return {
    workbookId: opts.workbookId,
  };
}

export class ExcelWorksheetCollection {
  constructor(private readonly workbook: ExcelWorkbook) {}

  async items(): Promise<ExcelWorksheet[]> {
    const sheets = await this.workbook.executeCode<Array<{ id: string; name: string }>>(`
      return workbook.worksheets.items().map((sheet) => ({ id: sheet.id(), name: sheet.name() }));
    `);
    return sheets.map((sheet) => new ExcelWorksheet(this.workbook, sheet.name, sheet.id));
  }

  getItem(nameOrId: string): ExcelWorksheet {
    return new ExcelWorksheet(this.workbook, nameOrId);
  }

  async getItemOrNullObject(nameOrId: string): Promise<ExcelWorksheet | null> {
    const exists = await this.workbook.executeCode<boolean>(
      `
      return workbook.worksheets.getItemOrNullObject(args.nameOrId) !== null;
    `,
      { nameOrId },
    );
    return exists ? new ExcelWorksheet(this.workbook, nameOrId) : null;
  }

  async add(name: string): Promise<ExcelWorksheet> {
    const sheet = await this.workbook.executeCode<{ id: string; name: string }>(
      `
      const sheet = workbook.worksheets.add(args.name);
      return { id: sheet.id(), name: sheet.name() };
    `,
      { name },
    );
    return new ExcelWorksheet(this.workbook, sheet.name, sheet.id);
  }

  async deleteItem(nameOrId: string): Promise<void> {
    await this.workbook.executeCode<void>(`workbook.worksheets.deleteItem(args.nameOrId);`, { nameOrId });
  }

  async move(nameOrId: string, toIndex: number): Promise<void> {
    await this.workbook.executeCode<void>(`workbook.worksheets.move(args.nameOrId, args.toIndex);`, {
      nameOrId,
      toIndex,
    });
  }

  async copy(nameOrId: string, newName?: string): Promise<ExcelWorksheet> {
    const sheet = await this.workbook.executeCode<{ id: string; name: string }>(
      `
      const sheet = workbook.worksheets.copy(args.nameOrId, args.newName);
      return { id: sheet.id(), name: sheet.name() };
    `,
      { nameOrId, newName },
    );
    return new ExcelWorksheet(this.workbook, sheet.name, sheet.id);
  }
}

export class ExcelWorksheet {
  readonly tables: ExcelTableCollection;
  readonly hyperlinks: ExcelHyperlinkCollection;
  readonly comments: ExcelCommentCollection;
  readonly threadedComments: ExcelThreadedCommentCollection;
  readonly dataValidations: ExcelDataValidationCollection;
  readonly conditionalFormats: ExcelConditionalFormatCollection;
  readonly drawings: ExcelDrawingCollection;
  readonly charts: ExcelChartCollection;
  readonly images: ExcelImageCollection;
  readonly shapes: ExcelShapeCollection;

  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly nameOrId: string,
    private readonly knownId?: string,
  ) {
    this.tables = new ExcelTableCollection(workbook, nameOrId);
    this.hyperlinks = new ExcelHyperlinkCollection(workbook, nameOrId);
    this.comments = new ExcelCommentCollection(workbook, nameOrId);
    this.threadedComments = new ExcelThreadedCommentCollection(workbook, nameOrId);
    this.dataValidations = new ExcelDataValidationCollection(workbook, nameOrId);
    this.conditionalFormats = new ExcelConditionalFormatCollection(workbook, nameOrId);
    this.drawings = new ExcelDrawingCollection(workbook, nameOrId);
    this.charts = new ExcelChartCollection(workbook, nameOrId);
    this.images = new ExcelImageCollection(workbook, nameOrId);
    this.shapes = new ExcelShapeCollection(workbook, nameOrId);
  }

  async id(): Promise<string> {
    if (this.knownId) return this.knownId;
    return this.run<string>(`return sheet.id();`);
  }

  async name(): Promise<string> {
    return this.run<string>(`return sheet.name();`);
  }

  async rename(name: string): Promise<void> {
    await this.run<void>(`sheet.rename(args.name);`, { name });
  }

  async visibility(): Promise<string | undefined> {
    return this.run<string | undefined>(`return sheet.visibility();`);
  }

  async setVisibility(visibility: string | undefined): Promise<void> {
    await this.run<void>(`sheet.setVisibility(args.visibility);`, { visibility });
  }

  async path(): Promise<string | undefined> {
    return this.run<string | undefined>(`return sheet.path();`);
  }

  async dimension(): Promise<string | undefined> {
    return this.run<string | undefined>(`return sheet.dimension();`);
  }

  getRange(address: string): ExcelRange {
    return new ExcelRange(this.workbook, this.nameOrId, address);
  }

  getCell(address: string): ExcelCell {
    return new ExcelCell(this.workbook, this.nameOrId, address);
  }

  async getUsedRange(): Promise<ExcelRange | null> {
    const address = await this.run<string | null>(`
      const range = sheet.getUsedRange();
      return range ? range.address() : null;
    `);
    return address ? this.getRange(address) : null;
  }

  async cells(): Promise<ExcelCellModel[]> {
    return this.run<ExcelCellModel[]>(`return sheet.cells();`);
  }

  async sheetViews(): Promise<ExcelSheetViewModel[]> {
    return this.run<ExcelSheetViewModel[]>(`return sheet.sheetViews();`);
  }

  async setSheetViews(sheetViews: ExcelSheetViewModel[]): Promise<void> {
    await this.run<void>(`sheet.setSheetViews(args.sheetViews);`, { sheetViews });
  }

  async columns(): Promise<ExcelColumnModel[]> {
    return this.run<ExcelColumnModel[]>(`return sheet.columns();`);
  }

  async rows(): Promise<ExcelRowModel[]> {
    return this.run<ExcelRowModel[]>(`return sheet.rows();`);
  }

  async mergedRanges(): Promise<ExcelMergedRangeModel[]> {
    return this.run<ExcelMergedRangeModel[]>(`return sheet.mergedRanges();`);
  }

  async insertRows(rowIndex: number, count = 1): Promise<void> {
    await this.run<void>(`sheet.insertRows(args.rowIndex, args.count);`, { rowIndex, count });
  }

  async deleteRows(rowIndex: number, count = 1): Promise<void> {
    await this.run<void>(`sheet.deleteRows(args.rowIndex, args.count);`, { rowIndex, count });
  }

  async insertColumns(columnIndex: number, count = 1): Promise<void> {
    await this.run<void>(`sheet.insertColumns(args.columnIndex, args.count);`, { columnIndex, count });
  }

  async deleteColumns(columnIndex: number, count = 1): Promise<void> {
    await this.run<void>(`sheet.deleteColumns(args.columnIndex, args.count);`, { columnIndex, count });
  }

  async moveRows(rowIndex: number, count: number, beforeRowIndex: number): Promise<void> {
    await this.run<void>(`sheet.moveRows(args.rowIndex, args.count, args.beforeRowIndex);`, {
      rowIndex,
      count,
      beforeRowIndex,
    });
  }

  async moveColumns(columnIndex: number, count: number, beforeColumnIndex: number): Promise<void> {
    await this.run<void>(
      `sheet.moveColumns(args.columnIndex, args.count, args.beforeColumnIndex);`,
      { columnIndex, count, beforeColumnIndex },
    );
  }

  async toModel(): Promise<ExcelSheetModel> {
    return this.run<ExcelSheetModel>(`return sheet.toModel();`);
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.nameOrId);
      ${code}
    `,
      { ...(args ?? {}), nameOrId: this.nameOrId },
    );
  }
}

export class ExcelRange {
  readonly format: ExcelRangeFormat;

  constructor(
    protected readonly workbook: ExcelWorkbook,
    protected readonly sheetNameOrId: string,
    protected readonly rangeAddress: string,
  ) {
    this.format = new ExcelRangeFormat(this);
  }

  async address(): Promise<string> {
    return this.run<string>(`return range.address();`);
  }

  async rowIndex(): Promise<number> {
    return this.run<number>(`return range.rowIndex();`);
  }

  async columnIndex(): Promise<number> {
    return this.run<number>(`return range.columnIndex();`);
  }

  async rowCount(): Promise<number> {
    return this.run<number>(`return range.rowCount();`);
  }

  async columnCount(): Promise<number> {
    return this.run<number>(`return range.columnCount();`);
  }

  async values(): Promise<ExcelCellValue[][]> {
    return this.run<ExcelCellValue[][]>(`return range.values();`);
  }

  async formulas(): Promise<(string | null)[][]> {
    return this.run<(string | null)[][]>(`return range.formulas();`);
  }

  async text(): Promise<string[][]> {
    return this.run<string[][]>(`return range.text();`);
  }

  async valueTypes(): Promise<ExcelCellKind[][]> {
    return this.run<ExcelCellKind[][]>(`return range.valueTypes();`);
  }

  async rawValues(): Promise<(string | null)[][]> {
    return this.run<(string | null)[][]>(`return range.rawValues();`);
  }

  async styleIds(): Promise<(number | undefined)[][]> {
    return this.run<(number | undefined)[][]>(`return range.styleIds();`);
  }

  async setStyleId(styleId: number | undefined): Promise<void> {
    await this.run<void>(`range.setStyleId(args.styleId);`, { styleId });
  }

  async rowHeights(): Promise<Array<number | undefined>> {
    return this.run<Array<number | undefined>>(`return range.rowHeights();`);
  }

  async setRowHeight(height: number | undefined): Promise<void> {
    await this.run<void>(`range.setRowHeight(args.height);`, { height });
  }

  async columnWidths(): Promise<Array<number | undefined>> {
    return this.run<Array<number | undefined>>(`return range.columnWidths();`);
  }

  async setColumnWidth(width: number | undefined): Promise<void> {
    await this.run<void>(`range.setColumnWidth(args.width);`, { width });
  }

  async merge(): Promise<void> {
    await this.run<void>(`range.merge();`);
  }

  async unmerge(): Promise<void> {
    await this.run<void>(`range.unmerge();`);
  }

  async setValues(values: ExcelCellValue[][]): Promise<void> {
    await this.run<void>(`range.setValues(args.values);`, { values });
  }

  async setFormulas(formulas: (string | null)[][]): Promise<void> {
    await this.run<void>(`range.setFormulas(args.formulas);`, { formulas });
  }

  async clear(): Promise<void> {
    await this.run<void>(`range.clear();`);
  }

  getCell(rowOffset: number, columnOffset: number): ExcelCell {
    return new ExcelCell(this.workbook, this.sheetNameOrId, this.rangeAddress, rowOffset, columnOffset);
  }

  getRow(rowOffset: number): ExcelRange {
    return new ExcelDerivedRange(this.workbook, this.sheetNameOrId, this.rangeAddress, `range.getRow(args.rowOffset)`, {
      rowOffset,
    });
  }

  getColumn(columnOffset: number): ExcelRange {
    return new ExcelDerivedRange(
      this.workbook,
      this.sheetNameOrId,
      this.rangeAddress,
      `range.getColumn(args.columnOffset)`,
      { columnOffset },
    );
  }

  resize(rowCount: number, columnCount: number): ExcelRange {
    return new ExcelDerivedRange(
      this.workbook,
      this.sheetNameOrId,
      this.rangeAddress,
      `range.resize(args.rowCount, args.columnCount)`,
      { rowCount, columnCount },
    );
  }

  offset(rowOffset: number, columnOffset: number, rowCount?: number, columnCount?: number): ExcelRange {
    return new ExcelDerivedRange(
      this.workbook,
      this.sheetNameOrId,
      this.rangeAddress,
      `range.offset(args.rowOffset, args.columnOffset, args.rowCount, args.columnCount)`,
      { rowOffset, columnOffset, rowCount, columnCount },
    );
  }

  slice(rowOffset: number, columnOffset: number, rowCount: number, columnCount: number): ExcelRange {
    return new ExcelDerivedRange(
      this.workbook,
      this.sheetNameOrId,
      this.rangeAddress,
      `range.slice(args.rowOffset, args.columnOffset, args.rowCount, args.columnCount)`,
      { rowOffset, columnOffset, rowCount, columnCount },
    );
  }

  async toModel(): Promise<ExcelCellModel[]> {
    return this.run<ExcelCellModel[]>(`return range.toModel();`);
  }

  runFormat<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.run<T>(code, args);
  }

  protected run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const range = sheet.getRange(args.rangeAddress);
      ${code}
    `,
      { ...(args ?? {}), sheetNameOrId: this.sheetNameOrId, rangeAddress: this.rangeAddress },
    );
  }
}

class ExcelDerivedRange extends ExcelRange {
  constructor(
    workbook: ExcelWorkbook,
    sheetNameOrId: string,
    rangeAddress: string,
    private readonly expression: string,
    private readonly expressionArgs: Record<string, unknown>,
  ) {
    super(workbook, sheetNameOrId, rangeAddress);
  }

  protected override run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      let range = sheet.getRange(args.rangeAddress);
      range = ${this.expression};
      ${code}
    `,
      { ...(args ?? {}), ...this.expressionArgs, sheetNameOrId: this.sheetNameOrId, rangeAddress: this.rangeAddress },
    );
  }
}

export class ExcelCell {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly cellAddressOrRangeAddress: string,
    private readonly rowOffset?: number,
    private readonly columnOffset?: number,
  ) {}

  async address(): Promise<string> {
    return this.run<string>(`return cell.address();`);
  }

  async value(): Promise<ExcelCellValue> {
    return this.run<ExcelCellValue>(`return cell.value();`);
  }

  async formula(): Promise<string | null> {
    return this.run<string | null>(`return cell.formula();`);
  }

  async text(): Promise<string> {
    return this.run<string>(`return cell.text();`);
  }

  async valueType(): Promise<ExcelCellKind> {
    return this.run<ExcelCellKind>(`return cell.valueType();`);
  }

  async rawValue(): Promise<string | null> {
    return this.run<string | null>(`return cell.rawValue();`);
  }

  async styleId(): Promise<number | undefined> {
    return this.run<number | undefined>(`return cell.styleId();`);
  }

  async setStyleId(styleId: number | undefined): Promise<void> {
    await this.run<void>(`cell.setStyleId(args.styleId);`, { styleId });
  }

  async setValue(value: ExcelCellValue): Promise<void> {
    await this.run<void>(`cell.setValue(args.value);`, { value });
  }

  async setFormula(formula: string | null): Promise<void> {
    await this.run<void>(`cell.setFormula(args.formula);`, { formula });
  }

  async clear(): Promise<void> {
    await this.run<void>(`cell.clear();`);
  }

  toRange(): ExcelRange {
    return new ExcelRange(this.workbook, this.sheetNameOrId, this.cellAddressOrRangeAddress);
  }

  async toModel(): Promise<ExcelCellModel | null> {
    return this.run<ExcelCellModel | null>(`return cell.toModel();`);
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    const isOffsetCell = this.rowOffset !== undefined && this.columnOffset !== undefined;
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const cell = args.isOffsetCell
        ? sheet.getRange(args.address).getCell(args.rowOffset, args.columnOffset)
        : sheet.getCell(args.address);
      ${code}
    `,
      {
        ...(args ?? {}),
        sheetNameOrId: this.sheetNameOrId,
        address: this.cellAddressOrRangeAddress,
        isOffsetCell,
        rowOffset: this.rowOffset,
        columnOffset: this.columnOffset,
      },
    );
  }
}

export class ExcelRangeFormat {
  readonly font: ExcelRangeFontFormat;
  readonly fill: ExcelRangeFillFormat;
  readonly border: ExcelRangeBorderFormat;
  readonly alignment: ExcelRangeAlignmentFormat;

  constructor(private readonly range: ExcelRange) {
    this.font = new ExcelRangeFontFormat(range);
    this.fill = new ExcelRangeFillFormat(range);
    this.border = new ExcelRangeBorderFormat(range);
    this.alignment = new ExcelRangeAlignmentFormat(range);
  }

  numberFormatCode(): Promise<string | undefined> {
    return this.range.runFormat<string | undefined>(`return range.format.numberFormatCode();`);
  }

  setNumberFormatCode(code: string | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.setNumberFormatCode(args.code);`, { code });
  }
}

export class ExcelRangeFontFormat {
  constructor(private readonly range: ExcelRange) {}

  value(): Promise<ExcelFontModel> {
    return this.range.runFormat<ExcelFontModel>(`return range.format.font.value();`);
  }

  setValue(font: ExcelFontModel): Promise<void> {
    return this.range.runFormat<void>(`range.format.font.setValue(args.font);`, { font });
  }

  bold(): Promise<boolean | undefined> {
    return this.range.runFormat<boolean | undefined>(`return range.format.font.bold();`);
  }

  setBold(value: boolean | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.font.setBold(args.value);`, { value });
  }

  italic(): Promise<boolean | undefined> {
    return this.range.runFormat<boolean | undefined>(`return range.format.font.italic();`);
  }

  setItalic(value: boolean | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.font.setItalic(args.value);`, { value });
  }

  underline(): Promise<boolean | undefined> {
    return this.range.runFormat<boolean | undefined>(`return range.format.font.underline();`);
  }

  setUnderline(value: boolean | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.font.setUnderline(args.value);`, { value });
  }

  color(): Promise<ExcelFontModel['color']> {
    return this.range.runFormat<ExcelFontModel['color']>(`return range.format.font.color();`);
  }

  setColor(color: ExcelFontModel['color']): Promise<void> {
    return this.range.runFormat<void>(`range.format.font.setColor(args.color);`, { color });
  }

  size(): Promise<number | undefined> {
    return this.range.runFormat<number | undefined>(`return range.format.font.size();`);
  }

  setSize(size: number | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.font.setSize(args.size);`, { size });
  }

  name(): Promise<string | undefined> {
    return this.range.runFormat<string | undefined>(`return range.format.font.name();`);
  }

  setName(name: string | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.font.setName(args.name);`, { name });
  }
}

export class ExcelRangeFillFormat {
  constructor(private readonly range: ExcelRange) {}

  value(): Promise<ExcelFillModel> {
    return this.range.runFormat<ExcelFillModel>(`return range.format.fill.value();`);
  }

  setValue(fill: ExcelFillModel): Promise<void> {
    return this.range.runFormat<void>(`range.format.fill.setValue(args.fill);`, { fill });
  }

  patternType(): Promise<string | undefined> {
    return this.range.runFormat<string | undefined>(`return range.format.fill.patternType();`);
  }

  setPatternType(patternType: string | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.fill.setPatternType(args.patternType);`, { patternType });
  }

  foregroundColor(): Promise<ExcelFillModel['foregroundColor']> {
    return this.range.runFormat<ExcelFillModel['foregroundColor']>(`return range.format.fill.foregroundColor();`);
  }

  setForegroundColor(color: ExcelFillModel['foregroundColor']): Promise<void> {
    return this.range.runFormat<void>(`range.format.fill.setForegroundColor(args.color);`, { color });
  }

  backgroundColor(): Promise<ExcelFillModel['backgroundColor']> {
    return this.range.runFormat<ExcelFillModel['backgroundColor']>(`return range.format.fill.backgroundColor();`);
  }

  setBackgroundColor(color: ExcelFillModel['backgroundColor']): Promise<void> {
    return this.range.runFormat<void>(`range.format.fill.setBackgroundColor(args.color);`, { color });
  }
}

export class ExcelRangeBorderFormat {
  constructor(private readonly range: ExcelRange) {}

  value(): Promise<ExcelBorderModel> {
    return this.range.runFormat<ExcelBorderModel>(`return range.format.border.value();`);
  }

  setValue(border: ExcelBorderModel): Promise<void> {
    return this.range.runFormat<void>(`range.format.border.setValue(args.border);`, { border });
  }

  setLeft(side: ExcelBorderModel['left']): Promise<void> {
    return this.range.runFormat<void>(`range.format.border.setLeft(args.side);`, { side });
  }

  setRight(side: ExcelBorderModel['right']): Promise<void> {
    return this.range.runFormat<void>(`range.format.border.setRight(args.side);`, { side });
  }

  setTop(side: ExcelBorderModel['top']): Promise<void> {
    return this.range.runFormat<void>(`range.format.border.setTop(args.side);`, { side });
  }

  setBottom(side: ExcelBorderModel['bottom']): Promise<void> {
    return this.range.runFormat<void>(`range.format.border.setBottom(args.side);`, { side });
  }
}

export class ExcelRangeAlignmentFormat {
  constructor(private readonly range: ExcelRange) {}

  value(): Promise<ExcelAlignmentModel | undefined> {
    return this.range.runFormat<ExcelAlignmentModel | undefined>(`return range.format.alignment.value();`);
  }

  setValue(alignment: ExcelAlignmentModel | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.alignment.setValue(args.alignment);`, { alignment });
  }

  setHorizontal(horizontal: string | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.alignment.setHorizontal(args.horizontal);`, { horizontal });
  }

  setVertical(vertical: string | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.alignment.setVertical(args.vertical);`, { vertical });
  }

  setWrapText(wrapText: boolean | undefined): Promise<void> {
    return this.range.runFormat<void>(`range.format.alignment.setWrapText(args.wrapText);`, { wrapText });
  }
}

export class ExcelTableCollection {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
  ) {}

  async items(): Promise<ExcelTable[]> {
    const tables = await runSheetCollection<ExcelTableModel[]>(this.workbook, this.sheetNameOrId, 'tables', 'toModel');
    return tables.map((table) => new ExcelTable(this.workbook, this.sheetNameOrId, table.name ?? table.displayName ?? ''));
  }

  getItem(nameOrDisplayName: string): ExcelTable {
    return new ExcelTable(this.workbook, this.sheetNameOrId, nameOrDisplayName);
  }

  async getItemOrNullObject(nameOrDisplayName: string): Promise<ExcelTable | null> {
    return (await sheetCollectionItemExists(this.workbook, this.sheetNameOrId, 'tables', nameOrDisplayName))
      ? this.getItem(nameOrDisplayName)
      : null;
  }

  async add(model: ExcelTableModel): Promise<ExcelTable> {
    const added = await runSheetCollection<ExcelTableModel>(this.workbook, this.sheetNameOrId, 'tables', 'add', { model });
    return new ExcelTable(this.workbook, this.sheetNameOrId, added.name ?? added.displayName ?? '');
  }

  async delete(nameOrDisplayName: string): Promise<void> {
    await runSheetCollection<void>(this.workbook, this.sheetNameOrId, 'tables', 'delete', { key: nameOrDisplayName });
  }

  async toModel(): Promise<ExcelTableModel[]> {
    return runSheetCollection<ExcelTableModel[]>(this.workbook, this.sheetNameOrId, 'tables', 'toModel');
  }
}

export class ExcelTable {
  readonly columns: ExcelTableColumnCollection;
  readonly style: ExcelTableStyle;

  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly nameOrDisplayName: string,
  ) {
    this.columns = new ExcelTableColumnCollection(workbook, sheetNameOrId, nameOrDisplayName);
    this.style = new ExcelTableStyle(workbook, sheetNameOrId, nameOrDisplayName);
  }

  async id(): Promise<number | undefined> {
    return this.run<number | undefined>(`return item.id();`);
  }

  async name(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.name();`);
  }

  async displayName(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.displayName();`);
  }

  async setName(name: string): Promise<void> {
    await this.run<void>(`item.setName(args.name);`, { name });
  }

  async setDisplayName(displayName: string): Promise<void> {
    await this.run<void>(`item.setDisplayName(args.displayName);`, { displayName });
  }

  async rangeAddress(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.rangeAddress();`);
  }

  async setRange(address: string): Promise<void> {
    await this.run<void>(`item.setRange(args.address);`, { address });
  }

  async resize(address: string): Promise<void> {
    await this.run<void>(`item.resize(args.address);`, { address });
  }

  async autoFilterRangeAddress(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.autoFilterRangeAddress();`);
  }

  async setAutoFilterRange(address: string | undefined): Promise<void> {
    await this.run<void>(`item.setAutoFilterRange(args.address);`, { address });
  }

  async headerRowCount(): Promise<number | undefined> {
    return this.run<number | undefined>(`return item.headerRowCount();`);
  }

  async setHeaderRowCount(count: number | undefined): Promise<void> {
    await this.run<void>(`item.setHeaderRowCount(args.count);`, { count });
  }

  async totalsRowCount(): Promise<number | undefined> {
    return this.run<number | undefined>(`return item.totalsRowCount();`);
  }

  async setTotalsRowCount(count: number | undefined): Promise<void> {
    await this.run<void>(`item.setTotalsRowCount(args.count);`, { count });
  }

  async totalsRowShown(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.totalsRowShown();`);
  }

  async setTotalsRowShown(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.setTotalsRowShown(args.value);`, { value });
  }

  async autoFilter(): Promise<ExcelTableAutoFilterModel | undefined> {
    return this.run<ExcelTableAutoFilterModel | undefined>(`return item.autoFilter();`);
  }

  async setAutoFilter(model: ExcelTableAutoFilterModel | undefined): Promise<void> {
    await this.run<void>(`item.setAutoFilter(args.model);`, { model });
  }

  async sortState(): Promise<ExcelTableSortStateModel | undefined> {
    return this.run<ExcelTableSortStateModel | undefined>(`return item.sortState();`);
  }

  async setSortState(sortState: ExcelTableSortStateModel | undefined): Promise<void> {
    await this.run<void>(`item.setSortState(args.sortState);`, { sortState });
  }

  async toModel(): Promise<ExcelTableModel> {
    return this.run<ExcelTableModel>(`return item.toModel();`);
  }

  async delete(): Promise<void> {
    await this.run<void>(`item.delete();`);
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return runSheetItem<T>(this.workbook, this.sheetNameOrId, 'tables', this.nameOrDisplayName, code, args);
  }
}

export class ExcelTableColumnCollection {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly tableNameOrDisplayName: string,
  ) {}

  async items(): Promise<ExcelTableColumn[]> {
    const columns = await this.run<ExcelTableModel['columns']>(`return item.columns.toModel();`);
    return columns.map((column) => new ExcelTableColumn(this.workbook, this.sheetNameOrId, this.tableNameOrDisplayName, column.id));
  }

  getItem(nameOrId: string | number): ExcelTableColumn {
    return new ExcelTableColumn(this.workbook, this.sheetNameOrId, this.tableNameOrDisplayName, nameOrId);
  }

  async add(model: ExcelTableModel['columns'][number]): Promise<ExcelTableColumn> {
    const added = await this.run<ExcelTableModel['columns'][number]>(`return item.columns.add(args.model).toModel();`, { model });
    return this.getItem(added.id);
  }

  async delete(nameOrId: string | number): Promise<void> {
    await this.run<void>(`item.columns.delete(args.nameOrId);`, { nameOrId });
  }

  async deleteItem(nameOrId: string | number): Promise<void> {
    await this.run<void>(`item.columns.deleteItem(args.nameOrId);`, { nameOrId });
  }

  async toModel(): Promise<ExcelTableModel['columns']> {
    return this.run<ExcelTableModel['columns']>(`return item.columns.toModel();`);
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return runSheetItem<T>(this.workbook, this.sheetNameOrId, 'tables', this.tableNameOrDisplayName, code, args);
  }
}

export class ExcelTableColumn {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly tableNameOrDisplayName: string,
    private readonly nameOrId: string | number,
  ) {}

  async id(): Promise<number> {
    return this.run<number>(`return column.id();`);
  }

  async name(): Promise<string> {
    return this.run<string>(`return column.name();`);
  }

  async setName(name: string): Promise<void> {
    await this.run<void>(`column.setName(args.name);`, { name });
  }

  async totalsRowLabel(): Promise<string | undefined> {
    return this.run<string | undefined>(`return column.totalsRowLabel();`);
  }

  async setTotalsRowLabel(label: string | undefined): Promise<void> {
    await this.run<void>(`column.setTotalsRowLabel(args.label);`, { label });
  }

  async totalsRowFunction(): Promise<string | undefined> {
    return this.run<string | undefined>(`return column.totalsRowFunction();`);
  }

  async setTotalsRowFunction(fn: string | undefined): Promise<void> {
    await this.run<void>(`column.setTotalsRowFunction(args.fn);`, { fn });
  }

  async toModel(): Promise<ExcelTableModel['columns'][number]> {
    return this.run<ExcelTableModel['columns'][number]>(`return column.toModel();`);
  }

  async delete(): Promise<void> {
    await this.run<void>(`column.delete();`);
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const table = sheet.tables.getItem(args.tableNameOrDisplayName);
      const column = table.columns.getItem(args.nameOrId);
      ${code}
    `,
      { ...(args ?? {}), sheetNameOrId: this.sheetNameOrId, tableNameOrDisplayName: this.tableNameOrDisplayName, nameOrId: this.nameOrId },
    );
  }
}

export class ExcelTableStyle {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly tableNameOrDisplayName: string,
  ) {}

  async toModel(): Promise<ExcelTableModel['style']> {
    return runSheetItem<ExcelTableModel['style']>(
      this.workbook,
      this.sheetNameOrId,
      'tables',
      this.tableNameOrDisplayName,
      `return item.style.toModel();`,
    );
  }

  async name(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.style.name();`);
  }

  async setName(name: string | undefined): Promise<void> {
    await this.run<void>(`item.style.setName(args.name);`, { name });
  }

  async showFirstColumn(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.style.showFirstColumn();`);
  }

  async setShowFirstColumn(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.style.setShowFirstColumn(args.value);`, { value });
  }

  async showLastColumn(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.style.showLastColumn();`);
  }

  async setShowLastColumn(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.style.setShowLastColumn(args.value);`, { value });
  }

  async showRowStripes(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.style.showRowStripes();`);
  }

  async setShowRowStripes(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.style.setShowRowStripes(args.value);`, { value });
  }

  async showColumnStripes(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.style.showColumnStripes();`);
  }

  async setShowColumnStripes(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.style.setShowColumnStripes(args.value);`, { value });
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return runSheetItem<T>(this.workbook, this.sheetNameOrId, 'tables', this.tableNameOrDisplayName, code, args);
  }
}

abstract class SimpleSheetCollection<TModel, TItem> {
  protected constructor(
    protected readonly workbook: ExcelWorkbook,
    protected readonly sheetNameOrId: string,
    private readonly collection: string,
    private readonly wrap: (workbook: ExcelWorkbook, sheetNameOrId: string, model: TModel) => TItem,
  ) {}

  async items(): Promise<TItem[]> {
    const models = await this.toModel();
    return models.map((model) => this.wrap(this.workbook, this.sheetNameOrId, model));
  }

  getItem(key: string): TItem {
    return this.wrap(this.workbook, this.sheetNameOrId, { ref: key, sqref: key } as TModel);
  }

  async getItemOrNullObject(key: string): Promise<TItem | null> {
    return (await sheetCollectionItemExists(this.workbook, this.sheetNameOrId, this.collection, key))
      ? this.getItem(key)
      : null;
  }

  async add(model: TModel): Promise<TItem> {
    const added = await runSheetCollection<TModel>(this.workbook, this.sheetNameOrId, this.collection, 'add', { model });
    return this.wrap(this.workbook, this.sheetNameOrId, added);
  }

  async delete(key: string): Promise<void> {
    await runSheetCollection<void>(this.workbook, this.sheetNameOrId, this.collection, 'delete', { key });
  }

  async toModel(): Promise<TModel[]> {
    return runSheetCollection<TModel[]>(this.workbook, this.sheetNameOrId, this.collection, 'toModel');
  }
}

abstract class SimpleSheetItem<TModel> {
  protected abstract readonly collection: string;

  constructor(
    protected readonly workbook: ExcelWorkbook,
    protected readonly sheetNameOrId: string,
    protected key: string,
  ) {}

  async toModel(): Promise<TModel> {
    return runSheetItem<TModel>(this.workbook, this.sheetNameOrId, this.collection, this.key, `return item.toModel();`);
  }

  async delete(): Promise<void> {
    await runSheetItem<void>(this.workbook, this.sheetNameOrId, this.collection, this.key, `item.delete();`);
  }

  protected run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return runSheetItem<T>(this.workbook, this.sheetNameOrId, this.collection, this.key, code, args);
  }
}

abstract class ReadonlySheetCollection<TModel, TItem> {
  protected constructor(
    protected readonly workbook: ExcelWorkbook,
    protected readonly sheetNameOrId: string,
    private readonly collection: string,
    private readonly wrap: (workbook: ExcelWorkbook, sheetNameOrId: string, index: number) => TItem,
  ) {}

  async items(): Promise<TItem[]> {
    const models = await this.toModel();
    return models.map((_, index) => this.wrap(this.workbook, this.sheetNameOrId, index));
  }

  async toModel(): Promise<TModel[]> {
    return runSheetCollection<TModel[]>(this.workbook, this.sheetNameOrId, this.collection, 'toModel');
  }
}

abstract class IndexedSheetItem<TModel> {
  protected abstract readonly collection: string;

  constructor(
    protected readonly workbook: ExcelWorkbook,
    protected readonly sheetNameOrId: string,
    protected indexOrName: number | string,
  ) {}

  async toModel(): Promise<TModel> {
    return this.run<TModel>(`return item.toModel();`);
  }

  protected run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const collection = sheet[args.collection];
      const item = typeof args.indexOrName === 'number' ? collection.items()[args.indexOrName] : collection.getItem(args.indexOrName);
      ${code}
    `,
      { ...(args ?? {}), sheetNameOrId: this.sheetNameOrId, collection: this.collection, indexOrName: this.indexOrName },
    );
  }
}

export class ExcelHyperlinkCollection extends SimpleSheetCollection<ExcelHyperlinkModel, ExcelHyperlink> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'hyperlinks', (workbook, sheetNameOrId, model) => new ExcelHyperlink(workbook, sheetNameOrId, model.ref));
  }
}

export class ExcelHyperlink extends SimpleSheetItem<ExcelHyperlinkModel> {
  protected readonly collection = 'hyperlinks';

  async ref(): Promise<string> {
    return this.run<string>(`return item.ref();`);
  }

  async setRef(ref: string): Promise<void> {
    await this.run<void>(`item.setRef(args.ref);`, { ref });
    this.key = ref;
  }

  async target(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.target();`);
  }

  async targetMode(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.targetMode();`);
  }

  async display(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.display();`);
  }

  async setDisplay(display: string | undefined): Promise<void> {
    await this.run<void>(`item.setDisplay(args.display);`, { display });
  }

  async setLocation(location: string | undefined): Promise<void> {
    await this.run<void>(`item.setLocation(args.location);`, { location });
  }
}

export class ExcelCommentCollection extends SimpleSheetCollection<ExcelCommentModel, ExcelComment> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'comments', (workbook, sheetNameOrId, model) => new ExcelComment(workbook, sheetNameOrId, model.ref));
  }
}

export class ExcelComment extends SimpleSheetItem<ExcelCommentModel> {
  protected readonly collection = 'comments';

  async ref(): Promise<string> {
    return this.run<string>(`return item.ref();`);
  }

  async setRef(ref: string): Promise<void> {
    await this.run<void>(`item.setRef(args.ref);`, { ref });
    this.key = ref;
  }

  async text(): Promise<string> {
    return this.run<string>(`return item.text();`);
  }

  async setText(text: string): Promise<void> {
    await this.run<void>(`item.setText(args.text);`, { text });
  }
}

export class ExcelThreadedCommentCollection extends SimpleSheetCollection<ExcelThreadedCommentModel, ExcelThreadedComment> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'threadedComments', (workbook, sheetNameOrId, model) => new ExcelThreadedComment(workbook, sheetNameOrId, model.ref));
  }
}

export class ExcelThreadedComment extends SimpleSheetItem<ExcelThreadedCommentModel> {
  protected readonly collection = 'threadedComments';

  async ref(): Promise<string> {
    return this.run<string>(`return item.ref();`);
  }

  async setRef(ref: string): Promise<void> {
    await this.run<void>(`item.setRef(args.ref);`, { ref });
    this.key = ref;
  }

  async text(): Promise<string> {
    return this.run<string>(`return item.text();`);
  }

  async setText(text: string): Promise<void> {
    await this.run<void>(`item.setText(args.text);`, { text });
  }
}

export class ExcelDataValidationCollection extends SimpleSheetCollection<ExcelDataValidationModel, ExcelDataValidation> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'dataValidations', (workbook, sheetNameOrId, model) => new ExcelDataValidation(workbook, sheetNameOrId, model.sqref ?? ''));
  }
}

export class ExcelDataValidation extends SimpleSheetItem<ExcelDataValidationModel> {
  protected readonly collection = 'dataValidations';

  async rangeAddress(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.rangeAddress();`);
  }

  async setRange(address: string): Promise<void> {
    await this.run<void>(`item.setRange(args.address);`, { address });
    this.key = address;
  }

  async type(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.type();`);
  }

  async setType(type: string | undefined): Promise<void> {
    await this.run<void>(`item.setType(args.type);`, { type });
  }

  async operator(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.operator();`);
  }

  async setOperator(operator: string | undefined): Promise<void> {
    await this.run<void>(`item.setOperator(args.operator);`, { operator });
  }

  async allowBlank(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.allowBlank();`);
  }

  async setAllowBlank(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.setAllowBlank(args.value);`, { value });
  }

  async showDropDown(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.showDropDown();`);
  }

  async setShowDropDown(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.setShowDropDown(args.value);`, { value });
  }

  async showInputMessage(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.showInputMessage();`);
  }

  async setShowInputMessage(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.setShowInputMessage(args.value);`, { value });
  }

  async showErrorMessage(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return item.showErrorMessage();`);
  }

  async setShowErrorMessage(value: boolean | undefined): Promise<void> {
    await this.run<void>(`item.setShowErrorMessage(args.value);`, { value });
  }

  async errorStyle(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.errorStyle();`);
  }

  async setErrorStyle(style: string | undefined): Promise<void> {
    await this.run<void>(`item.setErrorStyle(args.style);`, { style });
  }

  async errorTitle(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.errorTitle();`);
  }

  async error(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.error();`);
  }

  async setError(title: string | undefined, message: string | undefined): Promise<void> {
    await this.run<void>(`item.setError(args.title, args.message);`, { title, message });
  }

  async promptTitle(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.promptTitle();`);
  }

  async prompt(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.prompt();`);
  }

  async setPrompt(title: string | undefined, message: string | undefined): Promise<void> {
    await this.run<void>(`item.setPrompt(args.title, args.message);`, { title, message });
  }

  async formula1(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.formula1();`);
  }

  async setFormula1(formula: string | undefined): Promise<void> {
    await this.run<void>(`item.setFormula1(args.formula);`, { formula });
  }

  async formula2(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.formula2();`);
  }

  async setFormula2(formula: string | undefined): Promise<void> {
    await this.run<void>(`item.setFormula2(args.formula);`, { formula });
  }
}

export class ExcelConditionalFormatCollection extends SimpleSheetCollection<ExcelConditionalFormatModel, ExcelConditionalFormat> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'conditionalFormats', (workbook, sheetNameOrId, model) => new ExcelConditionalFormat(workbook, sheetNameOrId, model.sqref ?? ''));
  }
}

export class ExcelConditionalFormat extends SimpleSheetItem<ExcelConditionalFormatModel> {
  protected readonly collection = 'conditionalFormats';
  readonly rules: ExcelConditionalFormatRuleCollection;

  constructor(workbook: ExcelWorkbook, sheetNameOrId: string, key: string) {
    super(workbook, sheetNameOrId, key);
    this.rules = new ExcelConditionalFormatRuleCollection(workbook, sheetNameOrId, key);
  }

  async rangeAddress(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.rangeAddress();`);
  }

  async setRange(address: string): Promise<void> {
    await this.run<void>(`item.setRange(args.address);`, { address });
    this.key = address;
    this.rules.formatKey = address;
  }
}

export class ExcelConditionalFormatRuleCollection {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    public formatKey: string,
  ) {}

  async items(): Promise<ExcelConditionalFormatRule[]> {
    const models = await this.toModel();
    return models.map((_, index) => new ExcelConditionalFormatRule(this.workbook, this.sheetNameOrId, this.formatKey, index));
  }

  getItem(index: number): ExcelConditionalFormatRule {
    return new ExcelConditionalFormatRule(this.workbook, this.sheetNameOrId, this.formatKey, index);
  }

  async getItemOrNullObject(index: number): Promise<ExcelConditionalFormatRule | null> {
    const exists = await this.workbook.executeCode<boolean>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const format = sheet.conditionalFormats.getItem(args.formatKey);
      return format.rules.getItemOrNullObject(args.index) !== null;
    `,
      { sheetNameOrId: this.sheetNameOrId, formatKey: this.formatKey, index },
    );
    return exists ? this.getItem(index) : null;
  }

  async add(model: ExcelConditionalFormatRuleModel): Promise<ExcelConditionalFormatRule> {
    const addedIndex = await this.workbook.executeCode<number>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const format = sheet.conditionalFormats.getItem(args.formatKey);
      format.rules.add(args.model);
      return format.rules.items().length - 1;
    `,
      { sheetNameOrId: this.sheetNameOrId, formatKey: this.formatKey, model },
    );
    return this.getItem(addedIndex);
  }

  async delete(index: number): Promise<void> {
    await this.workbook.executeCode<void>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const format = sheet.conditionalFormats.getItem(args.formatKey);
      format.rules.delete(args.index);
    `,
      { sheetNameOrId: this.sheetNameOrId, formatKey: this.formatKey, index },
    );
  }

  async toModel(): Promise<ExcelConditionalFormatRuleModel[]> {
    return this.workbook.executeCode<ExcelConditionalFormatRuleModel[]>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const format = sheet.conditionalFormats.getItem(args.formatKey);
      return format.rules.toModel();
    `,
      { sheetNameOrId: this.sheetNameOrId, formatKey: this.formatKey },
    );
  }
}

export class ExcelConditionalFormatRule {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly formatKey: string,
    private readonly index: number,
  ) {}

  async type(): Promise<string | undefined> {
    return this.run<string | undefined>(`return rule.type();`);
  }

  async setType(type: string | undefined): Promise<void> {
    await this.run<void>(`rule.setType(args.type);`, { type });
  }

  async priority(): Promise<number | undefined> {
    return this.run<number | undefined>(`return rule.priority();`);
  }

  async setPriority(priority: number | undefined): Promise<void> {
    await this.run<void>(`rule.setPriority(args.priority);`, { priority });
  }

  async operator(): Promise<string | undefined> {
    return this.run<string | undefined>(`return rule.operator();`);
  }

  async setOperator(operator: string | undefined): Promise<void> {
    await this.run<void>(`rule.setOperator(args.operator);`, { operator });
  }

  async stopIfTrue(): Promise<boolean | undefined> {
    return this.run<boolean | undefined>(`return rule.stopIfTrue();`);
  }

  async setStopIfTrue(value: boolean | undefined): Promise<void> {
    await this.run<void>(`rule.setStopIfTrue(args.value);`, { value });
  }

  async formulas(): Promise<string[]> {
    return this.run<string[]>(`return rule.formulas();`);
  }

  async setFormula(index: number, formula: string): Promise<void> {
    await this.run<void>(`rule.setFormula(args.index, args.formula);`, { index, formula });
  }

  async addFormula(formula: string): Promise<void> {
    await this.run<void>(`rule.addFormula(args.formula);`, { formula });
  }

  async deleteFormula(index: number): Promise<void> {
    await this.run<void>(`rule.deleteFormula(args.index);`, { index });
  }

  async toModel(): Promise<ExcelConditionalFormatRuleModel> {
    return this.run<ExcelConditionalFormatRuleModel>(`return rule.toModel();`);
  }

  async delete(): Promise<void> {
    await this.run<void>(`rule.delete();`);
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const format = sheet.conditionalFormats.getItem(args.formatKey);
      const rule = format.rules.getItem(args.ruleIndex);
      ${code}
    `,
      { ...(args ?? {}), sheetNameOrId: this.sheetNameOrId, formatKey: this.formatKey, ruleIndex: this.index },
    );
  }
}

export class ExcelDrawingCollection extends ReadonlySheetCollection<ExcelDrawingModel, ExcelDrawing> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'drawings', (workbook, sheetNameOrId, index) => new ExcelDrawing(workbook, sheetNameOrId, index));
  }
}

export class ExcelDrawing extends IndexedSheetItem<ExcelDrawingModel> {
  protected readonly collection = 'drawings';
  readonly anchors: ExcelDrawingAnchorCollection;

  constructor(workbook: ExcelWorkbook, sheetNameOrId: string, index: number) {
    super(workbook, sheetNameOrId, index);
    this.anchors = new ExcelDrawingAnchorCollection(workbook, sheetNameOrId, index);
  }

  async relationshipId(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.relationshipId();`);
  }

  async path(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.path();`);
  }
}

export class ExcelDrawingAnchorCollection {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly drawingIndex: number,
  ) {}

  async items(): Promise<ExcelDrawingAnchor[]> {
    const models = await this.toModel();
    return models.map((_, index) => new ExcelDrawingAnchor(this.workbook, this.sheetNameOrId, this.drawingIndex, index));
  }

  getItem(index: number): ExcelDrawingAnchor {
    return new ExcelDrawingAnchor(this.workbook, this.sheetNameOrId, this.drawingIndex, index);
  }

  async getItemOrNullObject(index: number): Promise<ExcelDrawingAnchor | null> {
    const exists = await this.workbook.executeCode<boolean>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const drawing = sheet.drawings.items()[args.drawingIndex];
      return Boolean(drawing?.anchors.getItemOrNullObject(args.index));
    `,
      { sheetNameOrId: this.sheetNameOrId, drawingIndex: this.drawingIndex, index },
    );
    return exists ? this.getItem(index) : null;
  }

  async toModel(): Promise<ExcelDrawingAnchorModel[]> {
    return this.workbook.executeCode<ExcelDrawingAnchorModel[]>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      return sheet.drawings.items()[args.drawingIndex].anchors.toModel();
    `,
      { sheetNameOrId: this.sheetNameOrId, drawingIndex: this.drawingIndex },
    );
  }
}

export class ExcelDrawingAnchor {
  readonly from: ExcelDrawingMarker;
  readonly to: ExcelDrawingMarker;

  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly drawingIndex: number,
    private readonly anchorIndex: number,
  ) {
    this.from = new ExcelDrawingMarker(workbook, sheetNameOrId, drawingIndex, anchorIndex, 'from');
    this.to = new ExcelDrawingMarker(workbook, sheetNameOrId, drawingIndex, anchorIndex, 'to');
  }

  async kind(): Promise<ExcelDrawingAnchorModel['kind']> {
    return this.run<ExcelDrawingAnchorModel['kind']>(`return anchor.kind();`);
  }

  async object(): Promise<ExcelDrawingObjectModel | undefined> {
    return this.run<ExcelDrawingObjectModel | undefined>(`return anchor.object();`);
  }

  async toModel(): Promise<ExcelDrawingAnchorModel> {
    return this.run<ExcelDrawingAnchorModel>(`return anchor.toModel();`);
  }

  private run<T>(code: string): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const anchor = sheet.drawings.items()[args.drawingIndex].anchors.getItem(args.anchorIndex);
      ${code}
    `,
      { sheetNameOrId: this.sheetNameOrId, drawingIndex: this.drawingIndex, anchorIndex: this.anchorIndex },
    );
  }
}

export class ExcelDrawingMarker {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly drawingIndex: number,
    private readonly anchorIndex: number,
    private readonly markerName: 'from' | 'to',
  ) {}

  async row(): Promise<number | undefined> {
    return this.run<number | undefined>(`return marker.row();`);
  }

  async setRow(row: number | undefined): Promise<void> {
    await this.run<void>(`marker.setRow(args.row);`, { row });
  }

  async column(): Promise<number | undefined> {
    return this.run<number | undefined>(`return marker.column();`);
  }

  async setColumn(column: number | undefined): Promise<void> {
    await this.run<void>(`marker.setColumn(args.column);`, { column });
  }

  async rowOffset(): Promise<number | undefined> {
    return this.run<number | undefined>(`return marker.rowOffset();`);
  }

  async setRowOffset(offset: number | undefined): Promise<void> {
    await this.run<void>(`marker.setRowOffset(args.offset);`, { offset });
  }

  async columnOffset(): Promise<number | undefined> {
    return this.run<number | undefined>(`return marker.columnOffset();`);
  }

  async setColumnOffset(offset: number | undefined): Promise<void> {
    await this.run<void>(`marker.setColumnOffset(args.offset);`, { offset });
  }

  async toModel(): Promise<ExcelDrawingMarkerModel> {
    return this.run<ExcelDrawingMarkerModel>(`return marker.toModel();`);
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const anchor = sheet.drawings.items()[args.drawingIndex].anchors.getItem(args.anchorIndex);
      const marker = anchor[args.markerName];
      ${code}
    `,
      {
        ...(args ?? {}),
        sheetNameOrId: this.sheetNameOrId,
        drawingIndex: this.drawingIndex,
        anchorIndex: this.anchorIndex,
        markerName: this.markerName,
      },
    );
  }
}

export class ExcelChartCollection extends ReadonlySheetCollection<ExcelChartModel, ExcelChart> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'charts', (workbook, sheetNameOrId, index) => new ExcelChart(workbook, sheetNameOrId, index));
  }

  async add(options: ExcelChartAddOptions): Promise<ExcelChart> {
    const index = await this.workbook.executeCode<number>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      sheet.charts.add(args.options);
      return sheet.charts.items().length - 1;
    `,
      { sheetNameOrId: this.sheetNameOrId, options },
    );
    return new ExcelChart(this.workbook, this.sheetNameOrId, index);
  }
}

export class ExcelChart extends IndexedSheetItem<ExcelChartModel> {
  protected readonly collection = 'charts';
  readonly formulaRefs: ExcelFormulaRefCollection;

  constructor(workbook: ExcelWorkbook, sheetNameOrId: string, index: number) {
    super(workbook, sheetNameOrId, index);
    this.formulaRefs = new ExcelFormulaRefCollection(workbook, sheetNameOrId, index);
  }

  async relationshipId(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.relationshipId();`);
  }

  async drawingPath(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.drawingPath();`);
  }

  async path(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.path();`);
  }

  async title(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.title();`);
  }

  async setTitle(title: string | undefined): Promise<void> {
    await this.run<void>(`item.setTitle(args.title);`, { title });
  }
}

export class ExcelFormulaRefCollection {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private readonly sheetNameOrId: string,
    private readonly chartIndex: number,
  ) {}

  async items(): Promise<string[]> {
    return this.run<string[]>(`return item.formulaRefs.items();`);
  }

  async getItem(index: number): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.formulaRefs.getItem(args.index);`, { index });
  }

  async setItem(index: number, formulaRef: string): Promise<void> {
    await this.run<void>(`item.formulaRefs.setItem(args.index, args.formulaRef);`, { index, formulaRef });
  }

  async add(formulaRef: string): Promise<void> {
    await this.run<void>(`item.formulaRefs.add(args.formulaRef);`, { formulaRef });
  }

  async delete(index: number): Promise<void> {
    await this.run<void>(`item.formulaRefs.delete(args.index);`, { index });
  }

  async toModel(): Promise<string[]> {
    return this.run<string[]>(`return item.formulaRefs.toModel();`);
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
      const item = sheet.charts.items()[args.chartIndex];
      ${code}
    `,
      { ...(args ?? {}), sheetNameOrId: this.sheetNameOrId, chartIndex: this.chartIndex },
    );
  }
}

export class ExcelImageCollection extends ReadonlySheetCollection<ExcelImageModel, ExcelImage> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'images', (workbook, sheetNameOrId, index) => new ExcelImage(workbook, sheetNameOrId, index));
  }
}

export class ExcelImage extends IndexedSheetItem<ExcelImageModel> {
  protected readonly collection = 'images';

  async relationshipId(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.relationshipId();`);
  }

  async drawingPath(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.drawingPath();`);
  }

  async path(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.path();`);
  }

  async targetMode(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.targetMode();`);
  }

  async name(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.name();`);
  }
}

export class ExcelShapeCollection extends ReadonlySheetCollection<ExcelShapeModel, ExcelShape> {
  constructor(workbook: ExcelWorkbook, sheetNameOrId: string) {
    super(workbook, sheetNameOrId, 'shapes', (workbook, sheetNameOrId, index) => new ExcelShape(workbook, sheetNameOrId, index));
  }

  getItem(name: string): ExcelShape {
    return new ExcelShape(this.workbook, this.sheetNameOrId, name);
  }
}

export class ExcelShape extends IndexedSheetItem<ExcelShapeModel> {
  protected readonly collection = 'shapes';

  async drawingPath(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.drawingPath();`);
  }

  async name(): Promise<string | undefined> {
    return this.run<string | undefined>(`return item.name();`);
  }

  async setName(name: string | undefined): Promise<void> {
    await this.run<void>(`item.setName(args.name);`, { name });
    if (name !== undefined) this.indexOrName = name;
  }
}

function runSheetCollection<T>(
  workbook: ExcelWorkbook,
  sheetNameOrId: string,
  collection: string,
  method: string,
  args?: Record<string, unknown>,
): Promise<T> {
  return workbook.executeCode<T>(
    `
    const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
    const collection = sheet[args.collection];
    if (args.method === 'add') return collection.add(args.model).toModel();
    if (args.method === 'delete') return collection.delete(args.key);
    if (args.method === 'toModel') return collection.toModel();
    throw new Error('unsupported collection method: ' + args.method);
  `,
    { ...(args ?? {}), sheetNameOrId, collection, method },
  );
}

function runSheetItem<T>(
  workbook: ExcelWorkbook,
  sheetNameOrId: string,
  collection: string,
  key: string,
  code: string,
  args?: Record<string, unknown>,
): Promise<T> {
  return workbook.executeCode<T>(
    `
    const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
    const item = sheet[args.collection].getItem(args.key);
    ${code}
  `,
    { ...(args ?? {}), sheetNameOrId, collection, key },
  );
}

function sheetCollectionItemExists(
  workbook: ExcelWorkbook,
  sheetNameOrId: string,
  collection: string,
  key: string,
): Promise<boolean> {
  return workbook.executeCode<boolean>(
    `
    const sheet = workbook.worksheets.getItem(args.sheetNameOrId);
    return sheet[args.collection].getItemOrNullObject(args.key) !== null;
  `,
    { sheetNameOrId, collection, key },
  );
}

export class ExcelNamedItemCollection {
  constructor(private readonly workbook: ExcelWorkbook) {}

  async items(): Promise<ExcelNamedItem[]> {
    const names = await this.workbook.executeCode<Array<{ name: string }>>(`
      return workbook.names.items().map((name) => ({ name: name.name() }));
    `);
    return names.map((name) => new ExcelNamedItem(this.workbook, name.name));
  }

  getItem(name: string): ExcelNamedItem {
    return new ExcelNamedItem(this.workbook, name);
  }

  async getItemOrNullObject(name: string): Promise<ExcelNamedItem | null> {
    const exists = await this.workbook.executeCode<boolean>(
      `return workbook.names.getItemOrNullObject(args.name) !== null;`,
      { name },
    );
    return exists ? new ExcelNamedItem(this.workbook, name) : null;
  }

  async add(model: ExcelDefinedNameModel): Promise<ExcelNamedItem> {
    const added = await this.workbook.executeCode<ExcelDefinedNameModel>(
      `return workbook.names.add(args.model).toModel();`,
      { model },
    );
    return new ExcelNamedItem(this.workbook, added.name);
  }

  async deleteItem(name: string): Promise<void> {
    await this.workbook.executeCode<void>(`workbook.names.deleteItem(args.name);`, { name });
  }
}

export class ExcelNamedItem {
  constructor(
    private readonly workbook: ExcelWorkbook,
    private nameValue: string,
  ) {}

  async name(): Promise<string> {
    return this.run<string>(`return item.name();`);
  }

  async text(): Promise<string> {
    return this.run<string>(`return item.text();`);
  }

  async setText(text: string): Promise<void> {
    await this.run<void>(`item.setText(args.text);`, { text });
  }

  async rename(name: string): Promise<void> {
    await this.run<void>(`item.rename(args.nextName);`, { nextName: name });
    this.nameValue = name;
  }

  async localSheetId(): Promise<number | undefined> {
    return this.run<number | undefined>(`return item.localSheetId();`);
  }

  async setLocalSheetId(localSheetId: number | undefined): Promise<void> {
    await this.run<void>(`item.setLocalSheetId(args.localSheetId);`, { localSheetId });
  }

  private run<T>(code: string, args?: Record<string, unknown>): Promise<T> {
    return this.workbook.executeCode<T>(
      `
      const item = workbook.names.getItem(args.name);
      ${code}
    `,
      { ...(args ?? {}), name: this.nameValue },
    );
  }
}

function isWorkerTerminalScriptError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error);
  return message.includes('excel worker call timeout') || message.includes('excel worker script aborted');
}
