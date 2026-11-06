import { createPackage, fromZip, toZip, type Package } from '../pack/index.js';
import type { ExcelStylesModel, ExcelWorkbookModel } from './model.js';
import { readWorkbookModel } from './xlsx-reader.js';
import { createWorkbookPackageFromModel, syncWorkbookModelToPackage } from './xlsx-writer.js';
import { collectContentOnlyCellPatches, patchWorkbookCellsInPackage } from './xlsx-cell-patcher.js';
import { ExcelNamedItemCollection } from './names.js';
import { ExcelWorksheetCollection } from './worksheet.js';
import { ExcelMutationLog, type ExcelWorkbookMutation } from './mutations.js';
import { importExcelCellStyles, type ExcelCellStyleSnapshot } from './cell-style.js';
import {
  addWorksheetLinesToPackage,
  setWorksheetPrintSetupInPackage,
  type ExcelWorksheetLineOptions,
  type ExcelWorksheetPrintSetupOptions,
} from './worksheet-layout.js';

export interface ExcelLoadOptions {
  workbookId?: string;
}

export class ExcelWorkbookCore {
  readonly workbookId: string;
  readonly worksheets: ExcelWorksheetCollection;
  readonly names: ExcelNamedItemCollection;
  private readonly pkg: Package;
  private readonly model: ExcelWorkbookModel;
  private baselineModel: ExcelWorkbookModel;
  private readonly mutationLog = new ExcelMutationLog();

  private constructor(init: { workbookId: string; pkg: Package; model: ExcelWorkbookModel }) {
    this.workbookId = init.workbookId;
    this.pkg = init.pkg;
    this.model = init.model;
    this.baselineModel = structuredCloneModel(init.model);
    this.worksheets = new ExcelWorksheetCollection(init.model.sheets, init.model, (mutation) =>
      this.mutationLog.record(mutation),
    );
    this.names = new ExcelNamedItemCollection(init.model.definedNames);
  }

  static async fromZip(bytes: Uint8Array | ArrayBuffer, opts: ExcelLoadOptions = {}): Promise<ExcelWorkbookCore> {
    const pkg = await fromZip(bytes);
    return ExcelWorkbookCore.fromPackage(pkg, opts);
  }

  static fromPackage(pkg: Package, opts: ExcelLoadOptions = {}): ExcelWorkbookCore {
    return new ExcelWorkbookCore({
      workbookId: opts.workbookId ?? `excel:${Date.now()}`,
      pkg,
      model: readWorkbookModel(pkg),
    });
  }

  static fromModel(model: ExcelWorkbookModel, opts: ExcelLoadOptions = {}): ExcelWorkbookCore {
    return new ExcelWorkbookCore({
      workbookId: opts.workbookId ?? `excel:${Date.now()}`,
      pkg: createPackage(),
      model: structuredCloneModel(model),
    });
  }

  toModel(): ExcelWorkbookModel {
    return structuredCloneModel(this.model);
  }

  styles(): ExcelStylesModel {
    return cloneStyles(this.model.styles);
  }

  importCellStyles(snapshots: readonly ExcelCellStyleSnapshot[]): number[] {
    const count = this.model.styles.cellFormats.length;
    const ids = importExcelCellStyles(this.model.styles, snapshots);
    if (count !== this.model.styles.cellFormats.length) this.mutationLog.record({ kind: 'unsafe', scope: 'workbook', reason: 'Imported portable cell styles' });
    return ids;
  }

  getMutations(): ExcelWorkbookMutation[] {
    return this.mutationLog.list();
  }

  _replaceModelForHistory(model: ExcelWorkbookModel): void {
    const next = structuredCloneModel(model);
    this.model.sharedStrings.splice(0, this.model.sharedStrings.length, ...next.sharedStrings);
    this.model.definedNames.splice(0, this.model.definedNames.length, ...next.definedNames);
    this.model.styles = next.styles;
    this.model.sheets.splice(0, this.model.sheets.length, ...next.sheets);
    Object.defineProperty(this, 'worksheets', {
      value: new ExcelWorksheetCollection(this.model.sheets, this.model, (mutation) =>
        this.mutationLog.record(mutation),
      ),
      writable: true,
      configurable: true,
    });
    Object.defineProperty(this, 'names', {
      value: new ExcelNamedItemCollection(this.model.definedNames),
      writable: true,
      configurable: true,
    });
  }

  clearMutations(): void {
    this.mutationLog.clear();
  }

  getPackage(): Package {
    return this.pkg;
  }

  addWorksheetLines(sheetNameOrId: string, lines: readonly ExcelWorksheetLineOptions[]): void {
    const sheet = this.worksheets.getItem(sheetNameOrId);
    this.materializePackage();
    const sheetPath = this.resolveWorksheetPackagePath(sheet.id());
    addWorksheetLinesToPackage(this.pkg, sheetPath, lines);
    this.mutationLog.record({
      kind: 'worksheet.addLines',
      sheetId: sheet.id(),
      sheetName: sheet.name(),
      lines: [...lines],
    });
  }

  setWorksheetPrintSetup(
    sheetNameOrId: string,
    options: ExcelWorksheetPrintSetupOptions,
  ): void {
    const sheet = this.worksheets.getItem(sheetNameOrId);
    this.materializePackage();
    const sheetPath = this.resolveWorksheetPackagePath(sheet.id());
    const localSheetId = this.worksheets.items().findIndex((candidate) => candidate.id() === sheet.id());
    setWorksheetPrintSetupInPackage(this.pkg, sheetPath, sheet.name(), localSheetId, options);
    this.mutationLog.record({
      kind: 'worksheet.setPrintSetup',
      sheetId: sheet.id(),
      sheetName: sheet.name(),
      options,
    });
  }

  async toZip(): Promise<Uint8Array> {
    if (this.pkg.parts.size === 0) {
      this.materializePackage();
      return toZip(this.pkg);
    }
    const cellPatches = collectContentOnlyCellPatches(this.baselineModel, this.model);
    if (cellPatches) {
      patchWorkbookCellsInPackage(this.pkg, cellPatches);
      const bytes = await toZip(this.pkg);
      this.baselineModel = structuredCloneModel(this.model);
      return bytes;
    }
    syncWorkbookModelToPackage(this.pkg, this.model);
    const bytes = await toZip(this.pkg);
    this.baselineModel = structuredCloneModel(this.model);
    return bytes;
  }

  private materializePackage(): void {
    if (this.pkg.parts.size > 0) return;
    const materialized = createWorkbookPackageFromModel(this.model);
    for (const [path, part] of materialized.parts) this.pkg.parts.set(path, part);
  }

  private resolveWorksheetPackagePath(sheetId: string): string {
    const index = this.model.sheets.findIndex((sheet) => sheet.id === sheetId);
    if (index < 0) throw new Error(`Worksheet not found: ${sheetId}`);
    const path = this.model.sheets[index]?.path ?? `xl/worksheets/sheet${index + 1}.xml`;
    if (!this.pkg.parts.has(path)) throw new Error(`Worksheet has no package part: ${sheetId}`);
    return path;
  }
}

function structuredCloneModel(model: ExcelWorkbookModel): ExcelWorkbookModel {
  return {
    date1904: model.date1904,
    sharedStrings: [...model.sharedStrings],
    definedNames: model.definedNames.map((name) => ({ ...name })),
    styles: cloneStyles(model.styles),
    sheets: model.sheets.map((sheet) => ({
      ...sheet,
      sheetViews: sheet.sheetViews.map((view) => {
        const result = {
          ...view,
          selections: view.selections.map((selection) => ({ ...selection })),
        };
        if (view.pane) {
          result.pane = { ...view.pane };
        }
        return result;
      }),
      columns: sheet.columns.map((column) => ({ ...column })),
      rows: sheet.rows.map((row) => ({ ...row })),
      mergedRanges: sheet.mergedRanges.map((range) => ({ ...range })),
      hyperlinks: sheet.hyperlinks.map((hyperlink) => ({ ...hyperlink })),
      tables: sheet.tables.map((table) => ({
        ...table,
        columns: table.columns.map((column) => ({ ...column })),
        autoFilter: table.autoFilter
          ? {
              ...table.autoFilter,
              filterColumns: table.autoFilter.filterColumns.map((column) => ({
                ...column,
                filters: column.filters ? [...column.filters] : undefined,
              })),
              sortState: table.autoFilter.sortState
                ? {
                    ...table.autoFilter.sortState,
                    conditions: table.autoFilter.sortState.conditions.map((condition) => ({ ...condition })),
                  }
                : undefined,
            }
          : undefined,
        style: table.style ? { ...table.style } : undefined,
      })),
      dataValidations: sheet.dataValidations.map((validation) => ({ ...validation })),
      conditionalFormats: sheet.conditionalFormats.map((format) => ({
        ...format,
        rules: format.rules.map((rule) => ({
          ...rule,
          formulas: [...rule.formulas],
        })),
      })),
      comments: sheet.comments.map((comment) => ({ ...comment })),
      threadedComments: sheet.threadedComments.map((comment) => ({ ...comment })),
      drawings: sheet.drawings.map((drawing) => ({
        ...drawing,
        anchors: drawing.anchors.map((anchor) => ({
          ...anchor,
          from: anchor.from ? { ...anchor.from } : undefined,
          to: anchor.to ? { ...anchor.to } : undefined,
          object: anchor.object ? { ...anchor.object } : undefined,
        })),
      })),
      charts: sheet.charts.map((chart) => ({
        ...chart,
        formulaRefs: [...chart.formulaRefs],
      })),
      images: sheet.images.map((image) => ({ ...image })),
      cells: sheet.cells.map((cell) => ({ ...cell })),
    })),
  };
}

function cloneStyles(styles: ExcelStylesModel): ExcelStylesModel {
  return {
    numberFormats: styles.numberFormats.map((format) => ({ ...format })),
    fonts: styles.fonts.map((font) => ({
      ...font,
      color: font.color ? { ...font.color } : undefined,
    })),
    fills: styles.fills.map((fill) => ({
      ...fill,
      foregroundColor: fill.foregroundColor ? { ...fill.foregroundColor } : undefined,
      backgroundColor: fill.backgroundColor ? { ...fill.backgroundColor } : undefined,
    })),
    borders: styles.borders.map((border) => ({
      left: border.left ? { ...border.left, color: border.left.color ? { ...border.left.color } : undefined } : undefined,
      right: border.right ? { ...border.right, color: border.right.color ? { ...border.right.color } : undefined } : undefined,
      top: border.top ? { ...border.top, color: border.top.color ? { ...border.top.color } : undefined } : undefined,
      bottom: border.bottom
        ? { ...border.bottom, color: border.bottom.color ? { ...border.bottom.color } : undefined }
        : undefined,
      diagonal: border.diagonal
        ? { ...border.diagonal, color: border.diagonal.color ? { ...border.diagonal.color } : undefined }
        : undefined,
    })),
    cellFormats: styles.cellFormats.map((format) => ({
      ...format,
      alignment: format.alignment ? { ...format.alignment } : undefined,
    })),
  };
}
