import type { ExcelCellModel, ExcelSheetModel, ExcelStylesModel, ExcelWorkbookModel } from './model.js';

export type ExcelWritebackStrategy =
  | 'no-op'
  | 'workbook-package-replace'
  | 'sheet-shadow-replace'
  | 'diff-patch'
  | 'unsupported';

export type ExcelWritebackPreference = 'auto' | 'package' | 'sheet-shadow' | 'diff';

export interface ExcelWritebackCapabilities {
  /**
   * Host can replace the currently-open workbook with a full xlsx package in place.
   * This is usually desktop/file-system territory, not Office/WPS/ONLYOFFICE plugin mode.
   */
  canReplaceWorkbookPackage?: boolean;
  /** Host can insert sheets from a generated workbook package or shadow workbook. */
  canInsertShadowSheets?: boolean;
  /** Host can copy sheet/range content from a shadow sheet into an existing sheet. */
  canCopySheetContent?: boolean;
  /** Host can delete temporary or existing sheets. */
  canDeleteSheet?: boolean;
  /** Host can rename sheets. */
  canRenameSheet?: boolean;
  /** Host can patch cell values/formulas directly through live object calls. */
  canPatchCells?: boolean;
  /** Host can patch non-cell sheet objects such as tables, drawings, comments, validations, and formats. */
  canPatchSheetObjects?: boolean;
  /** Upper bound for live cell patch operations before planner should prefer sheet/package replacement. */
  maxCellPatchCount?: number;
}

export interface ExcelWritebackPlanOptions {
  preference?: ExcelWritebackPreference;
}

export type ExcelWorkbookDiffKind = 'added' | 'deleted' | 'renamed' | 'modified';

export type ExcelCellDiffKind = 'added' | 'deleted' | 'modified';

export interface ExcelCellDiff {
  kind: ExcelCellDiffKind;
  address: string;
  before?: ExcelCellModel;
  after?: ExcelCellModel;
}

export interface ExcelSheetDiff {
  kind: ExcelWorkbookDiffKind;
  sheetId: string;
  beforeName?: string;
  afterName?: string;
  cellChanges: number;
  cells: ExcelCellDiff[];
  structuralChanged: boolean;
  objectChanged: boolean;
}

export interface ExcelWorkbookDiff {
  changed: boolean;
  sheets: ExcelSheetDiff[];
}

export type ExcelWritebackOperation =
  | { type: 'replaceWorkbookPackage' }
  | {
      type: 'insertShadowSheet';
      sheetId: string;
      sheetName: string;
      targetSheet: ExcelSheetModel;
      targetStyles: ExcelStylesModel;
    }
  | { type: 'copySheetContent'; fromSheetName: string; toSheetName: string }
  | { type: 'repairSheetObjects'; sheetName: string }
  | { type: 'deleteSheet'; sheetName: string }
  | { type: 'renameSheet'; fromName: string; toName: string }
  | { type: 'patchCells'; sheetName: string; cellCount: number; changes: ExcelCellDiff[] }
  | { type: 'patchSheetObjects'; sheetName: string };

export interface ExcelWritebackPlan {
  strategy: ExcelWritebackStrategy;
  diff: ExcelWorkbookDiff;
  operations: ExcelWritebackOperation[];
  warnings: string[];
  unsupportedReasons: string[];
}

export function diffExcelWorkbookModels(before: ExcelWorkbookModel, after: ExcelWorkbookModel): ExcelWorkbookDiff {
  const beforeById = new Map(before.sheets.map((sheet) => [sheet.id, sheet]));
  const afterById = new Map(after.sheets.map((sheet) => [sheet.id, sheet]));
  const sheets: ExcelSheetDiff[] = [];

  for (const beforeSheet of before.sheets) {
    const afterSheet = afterById.get(beforeSheet.id);
    if (!afterSheet) {
      sheets.push({
        kind: 'deleted',
        sheetId: beforeSheet.id,
        beforeName: beforeSheet.name,
        cellChanges: beforeSheet.cells.length,
        cells: beforeSheet.cells.map((cell) => ({
          kind: 'deleted',
          address: cell.address,
          before: cloneCell(cell),
        })),
        structuralChanged: true,
        objectChanged: hasSheetObjects(beforeSheet),
      });
      continue;
    }

    const cells = diffCells(beforeSheet.cells, afterSheet.cells);
    const structuralChanged = sheetStructureKey(beforeSheet) !== sheetStructureKey(afterSheet);
    const objectChanged = sheetObjectsKey(beforeSheet) !== sheetObjectsKey(afterSheet);
    if (beforeSheet.name !== afterSheet.name) {
      sheets.push({
        kind: 'renamed',
        sheetId: beforeSheet.id,
        beforeName: beforeSheet.name,
        afterName: afterSheet.name,
        cellChanges: cells.length,
        cells,
        structuralChanged,
        objectChanged,
      });
    } else if (cells.length || structuralChanged || objectChanged) {
      sheets.push({
        kind: 'modified',
        sheetId: beforeSheet.id,
        beforeName: beforeSheet.name,
        afterName: afterSheet.name,
        cellChanges: cells.length,
        cells,
        structuralChanged,
        objectChanged,
      });
    }
  }

  for (const afterSheet of after.sheets) {
    if (!beforeById.has(afterSheet.id)) {
      sheets.push({
        kind: 'added',
        sheetId: afterSheet.id,
        afterName: afterSheet.name,
        cellChanges: afterSheet.cells.length,
        cells: afterSheet.cells.map((cell) => ({
          kind: 'added',
          address: cell.address,
          after: cloneCell(cell),
        })),
        structuralChanged: true,
        objectChanged: hasSheetObjects(afterSheet),
      });
    }
  }

  return {
    changed:
      sheets.length > 0 ||
      stableJson(before.definedNames) !== stableJson(after.definedNames) ||
      stableJson(before.styles) !== stableJson(after.styles),
    sheets,
  };
}

export function planExcelWriteback(
  before: ExcelWorkbookModel,
  after: ExcelWorkbookModel,
  capabilities: ExcelWritebackCapabilities,
  options: ExcelWritebackPlanOptions = {},
): ExcelWritebackPlan {
  const diff = diffExcelWorkbookModels(before, after);
  if (!diff.changed) {
    return { strategy: 'no-op', diff, operations: [], warnings: [], unsupportedReasons: [] };
  }

  const preference = options.preference ?? 'auto';
  const packagePlan = tryPlanPackage(diff, capabilities);
  if (preference === 'package') return packagePlan;

  const shadowPlan = tryPlanSheetShadow(diff, after, capabilities);
  if (preference === 'sheet-shadow') return shadowPlan;

  const diffPlan = tryPlanDiff(diff, capabilities);
  if (preference === 'diff') return diffPlan;

  if (packagePlan.strategy !== 'unsupported') return packagePlan;
  if (shadowPlan.strategy !== 'unsupported') return shadowPlan;
  if (diffPlan.strategy !== 'unsupported') return diffPlan;

  return unsupported(diff, [
    ...diffPlan.unsupportedReasons,
    ...shadowPlan.unsupportedReasons,
    ...packagePlan.unsupportedReasons,
  ]);
}

function tryPlanPackage(diff: ExcelWorkbookDiff, capabilities: ExcelWritebackCapabilities): ExcelWritebackPlan {
  if (!capabilities.canReplaceWorkbookPackage) {
    return unsupported(diff, ['host cannot replace the active workbook package in place']);
  }
  return {
    strategy: 'workbook-package-replace',
    diff,
    operations: [{ type: 'replaceWorkbookPackage' }],
    warnings: [],
    unsupportedReasons: [],
  };
}

function tryPlanDiff(diff: ExcelWorkbookDiff, capabilities: ExcelWritebackCapabilities): ExcelWritebackPlan {
  if (!capabilities.canPatchCells) {
    return unsupported(diff, ['host cannot patch cells directly']);
  }
  const maxCellPatchCount = capabilities.maxCellPatchCount ?? Number.POSITIVE_INFINITY;
  const operations: ExcelWritebackOperation[] = [];
  const warnings: string[] = [];

  for (const sheet of diff.sheets) {
    if (sheet.kind === 'added' || sheet.kind === 'deleted') {
      return unsupported(diff, [`diff patch cannot ${sheet.kind} sheets without sheet lifecycle support`]);
    }
    if (sheet.kind === 'renamed') {
      if (!capabilities.canRenameSheet || !sheet.beforeName || !sheet.afterName) {
        return unsupported(diff, ['diff patch cannot rename sheets with current host capabilities']);
      }
      operations.push({ type: 'renameSheet', fromName: sheet.beforeName, toName: sheet.afterName });
    }
    if (sheet.cellChanges > maxCellPatchCount) {
      return unsupported(diff, [`cell patch count ${sheet.cellChanges} exceeds maxCellPatchCount ${maxCellPatchCount}`]);
    }
    if (sheet.cellChanges > 0 && sheet.afterName) {
      operations.push({
        type: 'patchCells',
        sheetName: sheet.afterName,
        cellCount: sheet.cellChanges,
        changes: sheet.cells.map(cloneCellDiff),
      });
    }
    if (sheet.structuralChanged || sheet.objectChanged) {
      if (!capabilities.canPatchSheetObjects || !sheet.afterName) {
        return unsupported(diff, ['diff patch cannot patch sheet structure or objects with current host capabilities']);
      }
      operations.push({ type: 'patchSheetObjects', sheetName: sheet.afterName });
      warnings.push('sheet object patch is adapter-defined and may require domain-specific repair logic');
    }
  }

  return { strategy: 'diff-patch', diff, operations, warnings, unsupportedReasons: [] };
}

function tryPlanSheetShadow(
  diff: ExcelWorkbookDiff,
  after: ExcelWorkbookModel,
  capabilities: ExcelWritebackCapabilities,
): ExcelWritebackPlan {
  const missing = requireCapabilities(capabilities, [
    'canInsertShadowSheets',
    'canCopySheetContent',
    'canDeleteSheet',
    'canRenameSheet',
  ]);
  if (missing.length) {
    return unsupported(diff, [`sheet shadow replace requires ${missing.join(', ')}`]);
  }

  const operations: ExcelWritebackOperation[] = [];
  const warnings: string[] = [];
  const afterById = new Map(after.sheets.map((sheet) => [sheet.id, sheet]));
  for (const sheet of diff.sheets) {
    if (sheet.kind === 'deleted') {
      if (sheet.beforeName) operations.push({ type: 'deleteSheet', sheetName: sheet.beforeName });
      continue;
    }
    if (!sheet.afterName) continue;
    const targetSheet = afterById.get(sheet.sheetId);
    if (!targetSheet) {
      return unsupported(diff, [`target sheet model not found for sheet id ${sheet.sheetId}`]);
    }
    const shadowName = `__coco_shadow_${sheet.sheetId}`;
    operations.push({
      type: 'insertShadowSheet',
      sheetId: sheet.sheetId,
      sheetName: shadowName,
      targetSheet: cloneSheet(targetSheet),
      targetStyles: cloneStyles(after.styles),
    });
    if (sheet.kind === 'added') {
      operations.push({ type: 'renameSheet', fromName: shadowName, toName: sheet.afterName });
      continue;
    }
    const targetName = sheet.beforeName ?? sheet.afterName;
    operations.push({ type: 'copySheetContent', fromSheetName: shadowName, toSheetName: targetName });
    if (sheet.objectChanged) {
      operations.push({ type: 'repairSheetObjects', sheetName: targetName });
      warnings.push(`sheet ${targetName} has object changes; adapter must repair host-specific object fidelity`);
    }
    operations.push({ type: 'deleteSheet', sheetName: shadowName });
    if (sheet.kind === 'renamed' && sheet.beforeName && sheet.afterName) {
      operations.push({ type: 'renameSheet', fromName: sheet.beforeName, toName: sheet.afterName });
    }
  }

  return { strategy: 'sheet-shadow-replace', diff, operations, warnings, unsupportedReasons: [] };
}

function requireCapabilities(
  capabilities: ExcelWritebackCapabilities,
  names: Array<keyof ExcelWritebackCapabilities>,
): string[] {
  return names.filter((name) => capabilities[name] !== true);
}

function unsupported(diff: ExcelWorkbookDiff, unsupportedReasons: string[]): ExcelWritebackPlan {
  return {
    strategy: 'unsupported',
    diff,
    operations: [],
    warnings: [],
    unsupportedReasons: [...new Set(unsupportedReasons)],
  };
}

function diffCells(before: ExcelCellModel[], after: ExcelCellModel[]): ExcelCellDiff[] {
  const beforeByAddress = new Map(before.map((cell) => [cell.address, cell]));
  const afterByAddress = new Map(after.map((cell) => [cell.address, cell]));
  const changes: ExcelCellDiff[] = [];
  for (const beforeCell of before) {
    const afterCell = afterByAddress.get(beforeCell.address);
    if (!afterCell) {
      changes.push({ kind: 'deleted', address: beforeCell.address, before: cloneCell(beforeCell) });
    } else if (stableJson(beforeCell) !== stableJson(afterCell)) {
      changes.push({
        kind: 'modified',
        address: beforeCell.address,
        before: cloneCell(beforeCell),
        after: cloneCell(afterCell),
      });
    }
  }
  for (const afterCell of after) {
    if (!beforeByAddress.has(afterCell.address)) {
      changes.push({ kind: 'added', address: afterCell.address, after: cloneCell(afterCell) });
    }
  }
  return changes.sort((left, right) => left.address.localeCompare(right.address, undefined, { numeric: true }));
}

function cloneCellDiff(diff: ExcelCellDiff): ExcelCellDiff {
  return {
    kind: diff.kind,
    address: diff.address,
    before: diff.before ? cloneCell(diff.before) : undefined,
    after: diff.after ? cloneCell(diff.after) : undefined,
  };
}

function cloneCell(cell: ExcelCellModel): ExcelCellModel {
  return { ...cell };
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

function cloneSheet(sheet: ExcelSheetModel): ExcelSheetModel {
  return {
    ...sheet,
    sheetViews: sheet.sheetViews.map((view) => ({
      ...view,
      pane: view.pane ? { ...view.pane } : undefined,
      selections: view.selections.map((selection) => ({ ...selection })),
    })),
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
      rules: format.rules.map((rule) => ({ ...rule, formulas: [...rule.formulas] })),
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
    charts: sheet.charts.map((chart) => ({ ...chart, formulaRefs: [...chart.formulaRefs] })),
    images: sheet.images.map((image) => ({ ...image })),
    cells: sheet.cells.map(cloneCell),
  };
}

function hasSheetObjects(sheet: ExcelSheetModel): boolean {
  return (
    sheet.hyperlinks.length > 0 ||
    sheet.tables.length > 0 ||
    sheet.dataValidations.length > 0 ||
    sheet.conditionalFormats.length > 0 ||
    sheet.comments.length > 0 ||
    sheet.threadedComments.length > 0 ||
    sheet.drawings.length > 0 ||
    sheet.charts.length > 0 ||
    sheet.images.length > 0
  );
}

function sheetStructureKey(sheet: ExcelSheetModel): string {
  return stableJson({
    name: sheet.name,
    visibility: sheet.visibility,
    dimension: sheet.dimension,
    sheetViews: sheet.sheetViews,
    columns: sheet.columns,
    rows: sheet.rows,
    mergedRanges: sheet.mergedRanges,
  });
}

function sheetObjectsKey(sheet: ExcelSheetModel): string {
  return stableJson({
    hyperlinks: sheet.hyperlinks,
    tables: sheet.tables,
    dataValidations: sheet.dataValidations,
    conditionalFormats: sheet.conditionalFormats,
    comments: sheet.comments,
    threadedComments: sheet.threadedComments,
    drawings: sheet.drawings,
    charts: sheet.charts,
    images: sheet.images,
  });
}

function stableJson(value: unknown): string {
  return JSON.stringify(sortForJson(value));
}

function sortForJson(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortForJson);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, entryValue]) => entryValue !== undefined)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, entryValue]) => [key, sortForJson(entryValue)]),
  );
}
