import { ExcelCell, ExcelRange } from './range.js';
import type { ExcelDrawingMarkerModel, ExcelSheetModel, ExcelSheetViewModel, ExcelWorkbookModel } from './model.js';
import type { ExcelMutationRecorder } from './mutations.js';
import {
  ExcelChartCollection,
  ExcelCommentCollection,
  ExcelConditionalFormatCollection,
  ExcelDataValidationCollection,
  ExcelDrawingCollection,
  ExcelHyperlinkCollection,
  ExcelImageCollection,
  ExcelShapeCollection,
  ExcelTableCollection,
  ExcelThreadedCommentCollection,
} from './objects.js';
import { formatCellAddress, formatRangeAddress } from './address.js';
import {
  renameSheetInFormula,
  rewriteFormulaForColumnDelete,
  rewriteFormulaForColumnInsert,
  rewriteFormulaForRowDelete,
  rewriteFormulaForRowInsert,
  rewriteReferenceForColumnDelete,
  rewriteReferenceForColumnInsert,
  rewriteReferenceForRowDelete,
  rewriteReferenceForRowInsert,
  type ExcelReferenceRewriteResult,
} from './reference-rewriter.js';

export class ExcelWorksheet {
  readonly hyperlinks: ExcelHyperlinkCollection;
  readonly tables: ExcelTableCollection;
  readonly dataValidations: ExcelDataValidationCollection;
  readonly conditionalFormats: ExcelConditionalFormatCollection;
  readonly comments: ExcelCommentCollection;
  readonly threadedComments: ExcelThreadedCommentCollection;
  readonly drawings: ExcelDrawingCollection;
  readonly charts: ExcelChartCollection;
  readonly images: ExcelImageCollection;
  readonly shapes: ExcelShapeCollection;
  private readonly model: ExcelSheetModel;
  private readonly workbook?: ExcelWorkbookModel;
  private readonly recordMutation?: ExcelMutationRecorder;

  constructor(model: ExcelSheetModel, workbook?: ExcelWorkbookModel, recordMutation?: ExcelMutationRecorder) {
    this.model = model;
    this.workbook = workbook;
    this.recordMutation = recordMutation;
    this.hyperlinks = new ExcelHyperlinkCollection(() => model.hyperlinks);
    this.tables = new ExcelTableCollection(() => model.tables);
    this.dataValidations = new ExcelDataValidationCollection(() => model.dataValidations);
    this.conditionalFormats = new ExcelConditionalFormatCollection(() => model.conditionalFormats);
    this.comments = new ExcelCommentCollection(() => model.comments);
    this.threadedComments = new ExcelThreadedCommentCollection(() => model.threadedComments);
    this.drawings = new ExcelDrawingCollection(() => model.drawings);
    this.charts = new ExcelChartCollection(
      () => model.charts,
      () => model.drawings,
    );
    this.images = new ExcelImageCollection(() => model.images);
    this.shapes = new ExcelShapeCollection(() => model.drawings);
  }

  id(): string {
    return this.model.id;
  }

  name(): string {
    return this.model.name;
  }

  rename(name: string): void {
    assertValidWorksheetName(name);
    this.assertUniqueWorksheetName(name, this.model);
    const oldName = this.model.name;
    this.model.name = name;
    this.rewriteWorkbookSheetFormulas((text) => renameSheetInFormula(text, oldName, name));
    this.rewriteDefinedNames((text) => renameSheetInFormula(text, oldName, name));
    this.recordMutation?.({
      kind: 'worksheet.rename',
      sheetId: this.model.id,
      sheetName: name,
      oldName,
      newName: name,
    });
  }

  insertRows(rowIndex: number, count = 1): void {
    assertPositiveIndex(rowIndex, 'rowIndex');
    assertPositiveCount(count, 'count');
    for (const cell of this.model.cells) {
      if (cell.row >= rowIndex) {
        cell.row += count;
        cell.address = formatCellAddress(cell.row, cell.column);
      }
      if (cell.formula) {
        cell.formula = rewriteFormulaForRowInsert(cell.formula, rowIndex, count, {
          sheetName: this.model.name,
        });
      }
    }
    this.rewriteDefinedNames((text) =>
      rewriteFormulaForRowInsert(text, rowIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteRuleFormulas((text) =>
      rewriteFormulaForRowInsert(text, rowIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteChartFormulaRefs((text) =>
      rewriteFormulaForRowInsert(text, rowIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteRangeReferences((address) =>
      rewriteReferenceForRowInsert(address, rowIndex, count, { sheetName: this.model.name }),
    );
    this.moveRowMetadataForInsert(rowIndex, count);
    this.moveDrawingAnchorsForRowInsert(rowIndex, count);
    normalizeSheetCells(this.model);
    this.recordMutation?.({
      kind: 'unsafe',
      scope: 'worksheet',
      sheetId: this.model.id,
      sheetName: this.model.name,
      reason: 'worksheet.insertRows is not replayable by first-phase mutation writeback',
    });
  }

  deleteRows(rowIndex: number, count = 1): void {
    assertPositiveIndex(rowIndex, 'rowIndex');
    assertPositiveCount(count, 'count');
    const endRow = rowIndex + count - 1;
    this.model.cells = this.model.cells.filter((cell) => {
      if (cell.row >= rowIndex && cell.row <= endRow) return false;
      if (cell.row > endRow) {
        cell.row -= count;
        cell.address = formatCellAddress(cell.row, cell.column);
      }
      if (cell.formula) {
        cell.formula = rewriteFormulaForRowDelete(cell.formula, rowIndex, count, {
          sheetName: this.model.name,
        });
      }
      return true;
    });
    this.rewriteDefinedNames((text) =>
      rewriteFormulaForRowDelete(text, rowIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteRuleFormulas((text) =>
      rewriteFormulaForRowDelete(text, rowIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteChartFormulaRefs((text) =>
      rewriteFormulaForRowDelete(text, rowIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteRangeReferences((address) =>
      rewriteReferenceForRowDelete(address, rowIndex, count, { sheetName: this.model.name }),
    );
    this.moveRowMetadataForDelete(rowIndex, count);
    this.moveDrawingAnchorsForRowDelete(rowIndex, count);
    normalizeSheetCells(this.model);
    this.recordMutation?.({
      kind: 'unsafe',
      scope: 'worksheet',
      sheetId: this.model.id,
      sheetName: this.model.name,
      reason: 'worksheet.deleteRows is not replayable by first-phase mutation writeback',
    });
  }

  insertColumns(columnIndex: number, count = 1): void {
    assertPositiveIndex(columnIndex, 'columnIndex');
    assertPositiveCount(count, 'count');
    for (const cell of this.model.cells) {
      if (cell.column >= columnIndex) {
        cell.column += count;
        cell.address = formatCellAddress(cell.row, cell.column);
      }
      if (cell.formula) {
        cell.formula = rewriteFormulaForColumnInsert(cell.formula, columnIndex, count, {
          sheetName: this.model.name,
        });
      }
    }
    this.rewriteDefinedNames((text) =>
      rewriteFormulaForColumnInsert(text, columnIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteRuleFormulas((text) =>
      rewriteFormulaForColumnInsert(text, columnIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteChartFormulaRefs((text) =>
      rewriteFormulaForColumnInsert(text, columnIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteRangeReferences((address) =>
      rewriteReferenceForColumnInsert(address, columnIndex, count, { sheetName: this.model.name }),
    );
    this.moveColumnMetadataForInsert(columnIndex, count);
    this.moveDrawingAnchorsForColumnInsert(columnIndex, count);
    normalizeSheetCells(this.model);
    this.recordMutation?.({
      kind: 'unsafe',
      scope: 'worksheet',
      sheetId: this.model.id,
      sheetName: this.model.name,
      reason: 'worksheet.insertColumns is not replayable by first-phase mutation writeback',
    });
  }

  deleteColumns(columnIndex: number, count = 1): void {
    assertPositiveIndex(columnIndex, 'columnIndex');
    assertPositiveCount(count, 'count');
    const endColumn = columnIndex + count - 1;
    this.model.cells = this.model.cells.filter((cell) => {
      if (cell.column >= columnIndex && cell.column <= endColumn) return false;
      if (cell.column > endColumn) {
        cell.column -= count;
        cell.address = formatCellAddress(cell.row, cell.column);
      }
      if (cell.formula) {
        cell.formula = rewriteFormulaForColumnDelete(cell.formula, columnIndex, count, {
          sheetName: this.model.name,
        });
      }
      return true;
    });
    this.rewriteDefinedNames((text) =>
      rewriteFormulaForColumnDelete(text, columnIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteRuleFormulas((text) =>
      rewriteFormulaForColumnDelete(text, columnIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteChartFormulaRefs((text) =>
      rewriteFormulaForColumnDelete(text, columnIndex, count, { sheetName: this.model.name }),
    );
    this.rewriteRangeReferences((address) =>
      rewriteReferenceForColumnDelete(address, columnIndex, count, { sheetName: this.model.name }),
    );
    this.moveColumnMetadataForDelete(columnIndex, count);
    this.moveDrawingAnchorsForColumnDelete(columnIndex, count);
    normalizeSheetCells(this.model);
    this.recordMutation?.({
      kind: 'unsafe',
      scope: 'worksheet',
      sheetId: this.model.id,
      sheetName: this.model.name,
      reason: 'worksheet.deleteColumns is not replayable by first-phase mutation writeback',
    });
  }

  visibility(): string | undefined {
    return this.model.visibility;
  }

  setVisibility(visibility: string | undefined): void {
    if (visibility !== undefined && !['visible', 'hidden', 'veryHidden'].includes(visibility)) {
      throw new Error(`ExcelWorksheet.setVisibility: invalid visibility: ${visibility}`);
    }
    this.model.visibility = visibility;
    this.recordMutation?.({
      kind: 'unsafe',
      scope: 'worksheet',
      sheetId: this.model.id,
      sheetName: this.model.name,
      reason: 'worksheet.setVisibility is not replayable by first-phase mutation writeback',
    });
  }

  path(): string | undefined {
    return this.model.path;
  }

  dimension(): string | undefined {
    return this.model.dimension;
  }

  getRange(address: string): ExcelRange {
    return new ExcelRange(this.model, address, this.workbook?.styles, this.recordMutation);
  }

  getCell(address: string): ExcelCell {
    return new ExcelCell(this.model, address, this.workbook?.styles, this.recordMutation);
  }

  getUsedRange(): ExcelRange | null {
    if (!this.model.dimension) return null;
    return this.getRange(this.model.dimension);
  }

  cells() {
    return this.model.cells.map((cell) => ({ ...cell }));
  }

  sheetViews() {
    return cloneSheetViews(this.model.sheetViews);
  }

  setSheetViews(sheetViews: ExcelSheetViewModel[]): void {
    this.model.sheetViews = cloneSheetViews(sheetViews);
    this.recordMutation?.({
      kind: 'unsafe',
      scope: 'worksheet',
      sheetId: this.model.id,
      sheetName: this.model.name,
      reason: 'worksheet.setSheetViews is not replayable by first-phase mutation writeback',
    });
  }

  columns() {
    return this.model.columns.map((column) => ({ ...column }));
  }

  rows() {
    return this.model.rows.map((row) => ({ ...row }));
  }

  mergedRanges() {
    return this.model.mergedRanges.map((range) => ({ ...range }));
  }

  toModel(): ExcelSheetModel {
    return {
      ...this.model,
      sheetViews: this.sheetViews(),
      columns: this.model.columns.map((column) => ({ ...column })),
      rows: this.model.rows.map((row) => ({ ...row })),
      mergedRanges: this.model.mergedRanges.map((range) => ({ ...range })),
      hyperlinks: this.hyperlinks.toModel(),
      tables: this.tables.toModel(),
      dataValidations: this.dataValidations.toModel(),
      conditionalFormats: this.conditionalFormats.toModel(),
      comments: this.comments.toModel(),
      threadedComments: this.threadedComments.toModel(),
      drawings: this.drawings.toModel(),
      charts: this.charts.toModel(),
      images: this.images.toModel(),
      cells: this.model.cells.map((cell) => ({ ...cell })),
    };
  }

  private rewriteDefinedNames(rewrite: (text: string) => string): void {
    if (!this.workbook) return;
    const localSheetIndex = this.workbook.sheets.indexOf(this.model);
    for (const name of this.workbook.definedNames) {
      if (name.localSheetId !== undefined && name.localSheetId !== localSheetIndex) continue;
      name.text = rewrite(name.text);
    }
  }

  private rewriteRuleFormulas(rewrite: (text: string) => string): void {
    for (const validation of this.model.dataValidations) {
      if (validation.formula1) validation.formula1 = rewrite(validation.formula1);
      if (validation.formula2) validation.formula2 = rewrite(validation.formula2);
    }
    for (const format of this.model.conditionalFormats) {
      for (const rule of format.rules) {
        rule.formulas = rule.formulas.map((formula) => rewrite(formula));
      }
    }
  }

  private rewriteWorkbookSheetFormulas(rewrite: (text: string) => string): void {
    const sheets = this.workbook?.sheets ?? [this.model];
    for (const sheet of sheets) {
      for (const cell of sheet.cells) {
        if (cell.formula) cell.formula = rewrite(cell.formula);
      }
      for (const validation of sheet.dataValidations) {
        if (validation.formula1) validation.formula1 = rewrite(validation.formula1);
        if (validation.formula2) validation.formula2 = rewrite(validation.formula2);
      }
      for (const format of sheet.conditionalFormats) {
        for (const rule of format.rules) {
          rule.formulas = rule.formulas.map((formula) => rewrite(formula));
        }
      }
      for (const chart of sheet.charts) {
        chart.formulaRefs = chart.formulaRefs.map((reference) => rewrite(reference));
      }
    }
  }

  private rewriteChartFormulaRefs(rewrite: (text: string) => string): void {
    for (const chart of this.model.charts) {
      chart.formulaRefs = chart.formulaRefs.map((reference) => rewrite(reference));
    }
  }

  private rewriteRangeReferences(rewrite: (address: string) => ExcelReferenceRewriteResult): void {
    this.model.mergedRanges = this.model.mergedRanges.flatMap((range) => {
      const result = rewrite(range.address);
      return result.ok ? [{ ...range, address: result.address }] : [];
    });
    this.model.hyperlinks = this.model.hyperlinks.flatMap((hyperlink) => {
      const result = rewrite(hyperlink.ref);
      return result.ok ? [{ ...hyperlink, ref: result.address }] : [];
    });
    this.model.comments = this.model.comments.flatMap((comment) => {
      const result = rewrite(comment.ref);
      return result.ok ? [{ ...comment, ref: result.address }] : [];
    });
    this.model.threadedComments = this.model.threadedComments.flatMap((comment) => {
      const result = rewrite(comment.ref);
      return result.ok ? [{ ...comment, ref: result.address }] : [];
    });
    for (const view of this.model.sheetViews) {
      if (view.pane?.topLeftCell) {
        const result = rewrite(view.pane.topLeftCell);
        view.pane.topLeftCell = result.ok ? result.address : undefined;
      }
      for (const selection of view.selections) {
        if (selection.activeCell) {
          const result = rewrite(selection.activeCell);
          selection.activeCell = result.ok ? result.address : undefined;
        }
        if (selection.sqref) {
          selection.sqref = rewriteSqref(selection.sqref, rewrite);
        }
      }
    }
    this.model.tables = this.model.tables.flatMap((table) => {
      const ref = table.ref ? rewrite(table.ref) : undefined;
      if (table.ref && !ref?.ok) return [];
      const autoFilterRef = table.autoFilterRef ? rewrite(table.autoFilterRef) : undefined;
      return [
        {
          ...table,
          ref: ref?.address,
          autoFilterRef: autoFilterRef?.ok ? autoFilterRef.address : undefined,
          autoFilter:
            table.autoFilter && autoFilterRef?.ok
              ? { ...table.autoFilter, ref: autoFilterRef.address }
              : table.autoFilter,
        },
      ];
    });
    this.model.dataValidations = this.model.dataValidations.flatMap((validation) => {
      const sqref = validation.sqref ? rewriteSqref(validation.sqref, rewrite) : undefined;
      if (validation.sqref && !sqref) return [];
      return [{ ...validation, sqref }];
    });
    this.model.conditionalFormats = this.model.conditionalFormats.flatMap((format) => {
      const sqref = format.sqref ? rewriteSqref(format.sqref, rewrite) : undefined;
      if (format.sqref && !sqref) return [];
      return [{ ...format, sqref }];
    });
  }

  private moveRowMetadataForInsert(rowIndex: number, count: number): void {
    for (const row of this.model.rows) {
      if (row.index >= rowIndex) row.index += count;
    }
  }

  private moveRowMetadataForDelete(rowIndex: number, count: number): void {
    const endRow = rowIndex + count - 1;
    this.model.rows = this.model.rows.flatMap((row) => {
      if (row.index >= rowIndex && row.index <= endRow) return [];
      if (row.index > endRow) return [{ ...row, index: row.index - count }];
      return [row];
    });
  }

  private moveColumnMetadataForInsert(columnIndex: number, count: number): void {
    this.model.columns = this.model.columns.map((column) => {
      if (column.min >= columnIndex) {
        return { ...column, min: column.min + count, max: column.max + count };
      }
      if (column.max >= columnIndex) {
        return { ...column, max: column.max + count };
      }
      return column;
    });
  }

  private moveColumnMetadataForDelete(columnIndex: number, count: number): void {
    const endColumn = columnIndex + count - 1;
    this.model.columns = this.model.columns.flatMap((column) => {
      if (column.min >= columnIndex && column.max <= endColumn) return [];
      if (column.min > endColumn) {
        return [{ ...column, min: column.min - count, max: column.max - count }];
      }
      if (column.min < columnIndex && column.max > endColumn) {
        return [{ ...column, max: column.max - count }];
      }
      if (column.min < columnIndex && column.max >= columnIndex) {
        return [{ ...column, max: columnIndex - 1 }];
      }
      if (column.min <= endColumn && column.max > endColumn) {
        return [{ ...column, min: columnIndex, max: column.max - count }];
      }
      return [column];
    });
  }

  private moveDrawingAnchorsForRowInsert(rowIndex: number, count: number): void {
    const markerIndex = rowIndex - 1;
    for (const marker of this.drawingMarkers()) {
      if (marker.row !== undefined && marker.row >= markerIndex) marker.row += count;
    }
  }

  private moveDrawingAnchorsForRowDelete(rowIndex: number, count: number): void {
    const markerIndex = rowIndex - 1;
    const endIndex = markerIndex + count - 1;
    for (const marker of this.drawingMarkers()) {
      if (marker.row === undefined) continue;
      if (marker.row > endIndex) marker.row -= count;
      else if (marker.row >= markerIndex) marker.row = markerIndex;
    }
  }

  private moveDrawingAnchorsForColumnInsert(columnIndex: number, count: number): void {
    const markerIndex = columnIndex - 1;
    for (const marker of this.drawingMarkers()) {
      if (marker.column !== undefined && marker.column >= markerIndex) marker.column += count;
    }
  }

  private moveDrawingAnchorsForColumnDelete(columnIndex: number, count: number): void {
    const markerIndex = columnIndex - 1;
    const endIndex = markerIndex + count - 1;
    for (const marker of this.drawingMarkers()) {
      if (marker.column === undefined) continue;
      if (marker.column > endIndex) marker.column -= count;
      else if (marker.column >= markerIndex) marker.column = markerIndex;
    }
  }

  private drawingMarkers(): ExcelDrawingMarkerModel[] {
    return this.model.drawings.flatMap((drawing) =>
      drawing.anchors.flatMap((anchor) => {
        const markers: ExcelDrawingMarkerModel[] = [];
        if (anchor.from) markers.push(anchor.from);
        if (anchor.to) markers.push(anchor.to);
        return markers;
      }),
    );
  }

  private assertUniqueWorksheetName(name: string, current: ExcelSheetModel): void {
    if (!this.workbook) return;
    if (this.workbook.sheets.some((sheet) => sheet !== current && sheet.name === name)) {
      throw new Error(`ExcelWorkbook: worksheet already exists: ${name}`);
    }
  }
}

function assertValidWorksheetName(name: string): void {
  if (name.length < 1) {
    throw new Error('ExcelWorksheet.rename: worksheet name must not be empty');
  }
  if (name.length > 31) {
    throw new Error('ExcelWorksheet.rename: worksheet name must be 31 characters or fewer');
  }
  if (/[:\\/?*[\]]/.test(name)) {
    throw new Error(`ExcelWorksheet.rename: invalid worksheet name: ${name}`);
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new Error(`ExcelWorksheet.rename: worksheet name must not start or end with apostrophe`);
  }
}

function assertPositiveIndex(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`ExcelWorksheet: ${name} must be a positive integer`);
  }
}

function assertPositiveCount(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`ExcelWorksheet: ${name} must be a positive integer`);
  }
}

function normalizeSheetCells(sheet: ExcelSheetModel): void {
  sheet.cells.sort((a, b) => a.row - b.row || a.column - b.column);
  sheet.dimension = computeDimension(sheet);
}

function computeDimension(sheet: ExcelSheetModel): string | undefined {
  if (sheet.cells.length === 0) return undefined;
  let minRow = Number.POSITIVE_INFINITY;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let maxColumn = 0;
  for (const cell of sheet.cells) {
    minRow = Math.min(minRow, cell.row);
    minColumn = Math.min(minColumn, cell.column);
    maxRow = Math.max(maxRow, cell.row);
    maxColumn = Math.max(maxColumn, cell.column);
  }
  return formatRangeAddress(minRow, minColumn, maxRow - minRow + 1, maxColumn - minColumn + 1);
}

function rewriteSqref(
  sqref: string,
  rewrite: (address: string) => ExcelReferenceRewriteResult,
): string | undefined {
  const refs: string[] = [];
  for (const ref of sqref.split(/\s+/)) {
    if (!ref) continue;
    const result = rewrite(ref);
    if (result.ok) refs.push(result.address);
  }
  return refs.length ? refs.join(' ') : undefined;
}

export class ExcelWorksheetCollection {
  private readonly worksheets: ExcelWorksheet[];
  private readonly sheets: ExcelSheetModel[];
  private readonly workbook?: ExcelWorkbookModel;
  private readonly recordMutation?: ExcelMutationRecorder;

  constructor(
    sheets: ExcelSheetModel[],
    workbook?: ExcelWorkbookModel,
    recordMutation?: ExcelMutationRecorder,
  ) {
    this.sheets = sheets;
    this.workbook = workbook;
    this.recordMutation = recordMutation;
    this.worksheets = sheets.map((sheet) => new ExcelWorksheet(sheet, workbook, recordMutation));
  }

  items(): ExcelWorksheet[] {
    return [...this.worksheets];
  }

  getItem(nameOrId: string): ExcelWorksheet {
    const sheet = this.getItemOrNullObject(nameOrId);
    if (!sheet) {
      throw new Error(`ExcelWorkbook: worksheet not found: ${nameOrId}`);
    }
    return sheet;
  }

  getItemOrNullObject(nameOrId: string): ExcelWorksheet | null {
    return (
      this.worksheets.find((sheet) => sheet.name() === nameOrId || sheet.id() === nameOrId) ?? null
    );
  }

  add(name: string): ExcelWorksheet {
    assertValidWorksheetName(name);
    if (this.getItemOrNullObject(name)) {
      throw new Error(`ExcelWorkbook: worksheet already exists: ${name}`);
    }
    const sheet = createBlankSheetModel(name, nextSheetId(this.sheets));
    this.sheets.push(sheet);
    const worksheet = new ExcelWorksheet(sheet, this.workbook, this.recordMutation);
    this.worksheets.push(worksheet);
    this.recordMutation?.({
      kind: 'worksheet.add',
      sheetId: sheet.id,
      sheetName: sheet.name,
    });
    return worksheet;
  }

  deleteItem(nameOrId: string): void {
    const index = this.worksheets.findIndex(
      (sheet) => sheet.name() === nameOrId || sheet.id() === nameOrId,
    );
    if (index === -1) throw new Error(`ExcelWorkbook: worksheet not found: ${nameOrId}`);
    if (this.worksheets.length === 1)
      throw new Error('ExcelWorkbook: cannot delete the last worksheet');
    const sheet = this.sheets[index]!;
    this.worksheets.splice(index, 1);
    this.sheets.splice(index, 1);
    this.recordMutation?.({
      kind: 'worksheet.delete',
      sheetId: sheet.id,
      sheetName: sheet.name,
    });
  }

  move(nameOrId: string, toIndex: number): void {
    if (!Number.isInteger(toIndex) || toIndex < 0 || toIndex >= this.worksheets.length) {
      throw new Error(`ExcelWorksheetCollection.move: toIndex out of bounds: ${toIndex}`);
    }
    const fromIndex = this.worksheets.findIndex(
      (sheet) => sheet.name() === nameOrId || sheet.id() === nameOrId,
    );
    if (fromIndex === -1) throw new Error(`ExcelWorkbook: worksheet not found: ${nameOrId}`);
    if (fromIndex === toIndex) return;
    const [worksheet] = this.worksheets.splice(fromIndex, 1);
    const [sheet] = this.sheets.splice(fromIndex, 1);
    this.worksheets.splice(toIndex, 0, worksheet!);
    this.sheets.splice(toIndex, 0, sheet!);
  }

  copy(nameOrId: string, newName?: string): ExcelWorksheet {
    const sourceIndex = this.worksheets.findIndex(
      (sheet) => sheet.name() === nameOrId || sheet.id() === nameOrId,
    );
    if (sourceIndex === -1) throw new Error(`ExcelWorkbook: worksheet not found: ${nameOrId}`);
    const source = this.sheets[sourceIndex]!;
    const name = newName ?? uniqueWorksheetName(source.name, this.sheets);
    assertValidWorksheetName(name);
    if (this.getItemOrNullObject(name)) {
      throw new Error(`ExcelWorkbook: worksheet already exists: ${name}`);
    }
    const clone = cloneSheetModel(source);
    clone.id = nextSheetId(this.sheets);
    clone.name = name;
    clone.relationshipId = undefined;
    clone.path = undefined;
    this.sheets.splice(sourceIndex + 1, 0, clone);
    const worksheet = new ExcelWorksheet(clone, this.workbook, this.recordMutation);
    this.worksheets.splice(sourceIndex + 1, 0, worksheet);
    this.recordMutation?.({
      kind: 'unsafe',
      scope: 'worksheet',
      sheetId: clone.id,
      sheetName: clone.name,
      reason: 'worksheet.copy is not replayable by first-phase mutation writeback',
    });
    return worksheet;
  }
}

function createBlankSheetModel(name: string, id: string): ExcelSheetModel {
  return {
    id,
    name,
    sheetViews: [{ workbookViewId: 0, selections: [{ activeCell: 'A1', sqref: 'A1' }] }],
    columns: [],
    rows: [],
    mergedRanges: [],
    hyperlinks: [],
    tables: [],
    dataValidations: [],
    conditionalFormats: [],
    comments: [],
    threadedComments: [],
    drawings: [],
    charts: [],
    images: [],
    cells: [],
  };
}

function cloneSheetViews(sheetViews: ExcelSheetViewModel[]): ExcelSheetViewModel[] {
  return sheetViews.map((view) => ({
    ...view,
    pane: view.pane ? { ...view.pane } : undefined,
    selections: view.selections.map((selection) => ({ ...selection })),
  }));
}

function nextSheetId(sheets: ExcelSheetModel[]): string {
  let max = 0;
  for (const sheet of sheets) {
    const value = Number(sheet.id);
    if (Number.isFinite(value)) max = Math.max(max, value);
  }
  return String(max + 1 || sheets.length + 1);
}

function uniqueWorksheetName(baseName: string, sheets: ExcelSheetModel[]): string {
  const names = new Set(sheets.map((sheet) => sheet.name));
  for (let index = 2; index < 1000; index++) {
    const name = `${baseName} (${index})`;
    if (name.length <= 31 && !names.has(name)) return name;
  }
  throw new Error(
    `ExcelWorksheetCollection.copy: unable to create unique worksheet name for ${baseName}`,
  );
}

function cloneSheetModel(sheet: ExcelSheetModel): ExcelSheetModel {
  return JSON.parse(JSON.stringify(sheet)) as ExcelSheetModel;
}
