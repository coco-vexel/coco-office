import {
  formatCellAddress,
  formatRangeAddress,
  parseRangeAddress,
  type RangeAddress,
} from './address.js';
import type {
  ExcelAlignmentModel,
  ExcelBorderModel,
  ExcelCellFormatModel,
  ExcelCellKind,
  ExcelCellModel,
  ExcelCellValue,
  ExcelColorModel,
  ExcelFillModel,
  ExcelFontModel,
  ExcelSheetModel,
  ExcelStylesModel,
} from './model.js';
import { normalizeExcelColor } from './color.js';
import type {
  ExcelMutationRecorder,
  ExcelRangeSetBorderMutation,
  ExcelRangeSetColumnWidthMutation,
  ExcelRangeSetFillForegroundColorMutation,
  ExcelRangeSetFontColorMutation,
  ExcelRangeSetRowHeightMutation,
  ExcelWorkbookMutation,
} from './mutations.js';

type ReplayableRangeFormatMutation =
  | Omit<ExcelRangeSetFontColorMutation, 'sheetId' | 'sheetName'>
  | Omit<ExcelRangeSetFillForegroundColorMutation, 'sheetId' | 'sheetName'>
  | Omit<ExcelRangeSetBorderMutation, 'sheetId' | 'sheetName'>;

export class ExcelRange {
  readonly format: ExcelRangeFormat;
  private readonly sheet: ExcelSheetModel;
  private readonly styles: ExcelStylesModel;
  private readonly ref: RangeAddress;
  private readonly cellsByAddress: Map<string, ExcelCellModel>;
  private readonly recordMutation?: ExcelMutationRecorder;

  constructor(
    sheet: ExcelSheetModel,
    address: string,
    styles: ExcelStylesModel = createEmptyStyles(),
    recordMutation?: ExcelMutationRecorder,
  ) {
    this.sheet = sheet;
    this.styles = styles;
    this.recordMutation = recordMutation;
    this.ref = parseRangeAddress(address);
    if (this.ref.sheetName && this.ref.sheetName !== sheet.name) {
      throw new Error(
        `ExcelRange: address sheet '${this.ref.sheetName}' does not match worksheet '${sheet.name}'`,
      );
    }
    this.cellsByAddress = new Map(sheet.cells.map((cell) => [cell.address, cell]));
    this.format = new ExcelRangeFormat(this);
  }

  address(): string {
    return this.ref.address;
  }

  rowIndex(): number {
    return this.ref.startRow;
  }

  columnIndex(): number {
    return this.ref.startColumn;
  }

  rowCount(): number {
    return this.ref.rowCount;
  }

  columnCount(): number {
    return this.ref.columnCount;
  }

  values(): ExcelCellValue[][] {
    return this.matrix((cell) => cell.value);
  }

  formulas(): (string | null)[][] {
    return this.matrix((cell) => cell.formula ?? null);
  }

  text(): string[][] {
    return this.matrix((cell) => cellToText(cell));
  }

  valueTypes(): ExcelCellKind[][] {
    return this.matrix((cell) => cell.kind);
  }

  rawValues(): (string | null)[][] {
    return this.matrix((cell) => cell.rawValue ?? null);
  }

  styleIds(): (number | undefined)[][] {
    return this.matrix((cell) => cell.styleId);
  }

  setStyleId(styleId: number | undefined): void {
    if (styleId !== undefined && (!Number.isInteger(styleId) || styleId < 0)) {
      throw new Error(`ExcelRange.setStyleId: styleId must be a non-negative integer`);
    }
    for (let r = 0; r < this.ref.rowCount; r++) {
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        const cell = this.ensureCell(address);
        if (styleId === undefined) delete cell.styleId;
        else cell.styleId = styleId;
      }
    }
    normalizeSheetCells(this.sheet);
    this.recordMutation?.({
      kind: 'unsafe',
      scope: 'range',
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
      reason: 'range.setStyleId is not replayable by first-phase mutation writeback',
    });
  }

  rowHeights(): Array<number | undefined> {
    const heights: Array<number | undefined> = [];
    for (let r = 0; r < this.ref.rowCount; r++) {
      const rowIndex = this.ref.startRow + r;
      heights.push(this.sheet.rows.find((row) => row.index === rowIndex)?.height);
    }
    return heights;
  }

  setRowHeight(height: number | undefined): void {
    if (height !== undefined && (!Number.isFinite(height) || height < 0)) {
      throw new Error(`ExcelRange.setRowHeight: height must be a non-negative number`);
    }
    for (let r = 0; r < this.ref.rowCount; r++) {
      const row = this.ensureRow(this.ref.startRow + r);
      if (height === undefined) {
        delete row.height;
        delete row.customHeight;
      } else {
        row.height = height;
        row.customHeight = true;
      }
    }
    this.recordMutation?.({
      kind: 'range.setRowHeight',
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
      address: this.ref.address,
      height,
    });
  }

  columnWidths(): Array<number | undefined> {
    const widths: Array<number | undefined> = [];
    for (let c = 0; c < this.ref.columnCount; c++) {
      const columnIndex = this.ref.startColumn + c;
      widths.push(
        this.sheet.columns.find((column) => column.min <= columnIndex && column.max >= columnIndex)
          ?.width,
      );
    }
    return widths;
  }

  setColumnWidth(width: number | undefined): void {
    if (width !== undefined && (!Number.isFinite(width) || width < 0)) {
      throw new Error(`ExcelRange.setColumnWidth: width must be a non-negative number`);
    }
    for (let c = 0; c < this.ref.columnCount; c++) {
      const column = this.ensureColumn(this.ref.startColumn + c);
      if (width === undefined) {
        delete column.width;
        delete column.customWidth;
      } else {
        column.width = width;
        column.customWidth = true;
      }
    }
    normalizeColumns(this.sheet);
    this.recordMutation?.({
      kind: 'range.setColumnWidth',
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
      address: this.ref.address,
      width,
    });
  }

  merge(): void {
    const masterAddress = formatCellAddress(this.ref.startRow, this.ref.startColumn);
    for (let r = 0; r < this.ref.rowCount; r++) {
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        if (address !== masterAddress) this.clearCellContent(address);
      }
    }
    normalizeSheetCells(this.sheet);
    if (!this.sheet.mergedRanges.some((range) => range.address === this.ref.address)) {
      this.sheet.mergedRanges.push({ address: this.ref.address });
    }
    this.recordMutation?.({
      kind: 'range.merge',
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
      address: this.ref.address,
    });
  }

  unmerge(): void {
    this.sheet.mergedRanges = this.sheet.mergedRanges.filter(
      (range) => range.address !== this.ref.address,
    );
    this.recordMutation?.({
      kind: 'range.unmerge',
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
      address: this.ref.address,
    });
  }

  setValues(values: ExcelCellValue[][]): void {
    this.assertMatrixShape(values, 'ExcelRange.setValues');
    for (let r = 0; r < this.ref.rowCount; r++) {
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        this.setCellValue(address, values[r]![c]!);
      }
    }
    normalizeSheetCells(this.sheet);
    this.recordMutation?.({
      kind: 'range.setValues',
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
      address: this.ref.address,
      values,
    });
  }

  setFormulas(formulas: (string | null)[][]): void {
    this.assertMatrixShape(formulas, 'ExcelRange.setFormulas');
    for (let r = 0; r < this.ref.rowCount; r++) {
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        this.setCellFormula(address, formulas[r]![c]!);
      }
    }
    normalizeSheetCells(this.sheet);
    this.recordMutation?.({
      kind: 'range.setFormulas',
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
      address: this.ref.address,
      formulas,
    });
  }

  clear(): void {
    for (let r = 0; r < this.ref.rowCount; r++) {
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        this.deleteCell(address);
      }
    }
    normalizeSheetCells(this.sheet);
    this.recordMutation?.({
      kind: 'range.clear',
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
      address: this.ref.address,
    });
  }

  getCell(rowOffset: number, columnOffset: number): ExcelCell {
    this.assertOffset(rowOffset, columnOffset);
    const row = this.ref.startRow + rowOffset;
    const column = this.ref.startColumn + columnOffset;
    return new ExcelCell(this.sheet, formatCellAddress(row, column), this.styles, this.recordMutation);
  }

  getRow(rowOffset: number): ExcelRange {
    if (!Number.isInteger(rowOffset) || rowOffset < 0 || rowOffset >= this.ref.rowCount) {
      throw new Error(`ExcelRange.getRow: row offset out of bounds: ${rowOffset}`);
    }
    const row = this.ref.startRow + rowOffset;
    const start = formatCellAddress(row, this.ref.startColumn);
    const end = formatCellAddress(row, this.ref.startColumn + this.ref.columnCount - 1);
    return new ExcelRange(this.sheet, `${start}:${end}`, this.styles, this.recordMutation);
  }

  getColumn(columnOffset: number): ExcelRange {
    if (
      !Number.isInteger(columnOffset) ||
      columnOffset < 0 ||
      columnOffset >= this.ref.columnCount
    ) {
      throw new Error(`ExcelRange.getColumn: column offset out of bounds: ${columnOffset}`);
    }
    const column = this.ref.startColumn + columnOffset;
    const start = formatCellAddress(this.ref.startRow, column);
    const end = formatCellAddress(this.ref.startRow + this.ref.rowCount - 1, column);
    return new ExcelRange(this.sheet, `${start}:${end}`, this.styles, this.recordMutation);
  }

  resize(rowCount: number, columnCount: number): ExcelRange {
    return new ExcelRange(
      this.sheet,
      formatRangeAddress(this.ref.startRow, this.ref.startColumn, rowCount, columnCount),
      this.styles,
      this.recordMutation,
    );
  }

  offset(
    rowOffset: number,
    columnOffset: number,
    rowCount = this.ref.rowCount,
    columnCount = this.ref.columnCount,
  ): ExcelRange {
    if (!Number.isInteger(rowOffset) || !Number.isInteger(columnOffset)) {
      throw new Error(`ExcelRange.offset: offsets must be integers`);
    }
    return new ExcelRange(
      this.sheet,
      formatRangeAddress(
        this.ref.startRow + rowOffset,
        this.ref.startColumn + columnOffset,
        rowCount,
        columnCount,
      ),
      this.styles,
      this.recordMutation,
    );
  }

  slice(
    rowOffset: number,
    columnOffset: number,
    rowCount: number,
    columnCount: number,
  ): ExcelRange {
    if (
      !Number.isInteger(rowOffset) ||
      !Number.isInteger(columnOffset) ||
      !Number.isInteger(rowCount) ||
      !Number.isInteger(columnCount) ||
      rowOffset < 0 ||
      columnOffset < 0 ||
      rowCount < 1 ||
      columnCount < 1 ||
      rowOffset + rowCount > this.ref.rowCount ||
      columnOffset + columnCount > this.ref.columnCount
    ) {
      throw new Error(
        `ExcelRange.slice: range out of bounds row=${rowOffset}, column=${columnOffset}, rows=${rowCount}, columns=${columnCount}`,
      );
    }
    return new ExcelRange(
      this.sheet,
      formatRangeAddress(
        this.ref.startRow + rowOffset,
        this.ref.startColumn + columnOffset,
        rowCount,
        columnCount,
      ),
      this.styles,
      this.recordMutation,
    );
  }

  toModel(): ExcelCellModel[] {
    const cells: ExcelCellModel[] = [];
    for (let r = 0; r < this.ref.rowCount; r++) {
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        const cell = this.cellsByAddress.get(address);
        if (cell) cells.push({ ...cell });
      }
    }
    return cells;
  }

  getResolvedFormat(): ExcelCellFormatModel {
    const cell = this.firstCell();
    return cloneCellFormat(this.styles.cellFormats[cell?.styleId ?? 0] ?? {});
  }

  updateFormat(
    update: (format: ExcelCellFormatModel) => ExcelCellFormatModel,
    options: { recordUnsafe?: boolean } = {},
  ): void {
    for (let r = 0; r < this.ref.rowCount; r++) {
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        const cell = this.ensureCell(address);
        const current = cloneCellFormat(this.styles.cellFormats[cell.styleId ?? 0] ?? {});
        cell.styleId = ensureCellFormat(this.styles, update(current));
      }
    }
    normalizeSheetCells(this.sheet);
    if (options.recordUnsafe ?? true) {
      this.recordMutation?.({
        kind: 'unsafe',
        scope: 'range',
        sheetId: this.sheet.id,
        sheetName: this.sheet.name,
        reason: 'range.updateFormat is not replayable by first-phase mutation writeback',
      });
    }
  }

  recordReplayableMutation(mutation: ReplayableRangeFormatMutation): void {
    this.recordMutation?.({
      ...mutation,
      sheetId: this.sheet.id,
      sheetName: this.sheet.name,
    } as ExcelWorkbookMutation);
  }

  ensureFont(font: ExcelFontModel): number {
    return ensureStyleItem(this.styles.fonts, font, cloneFont);
  }

  ensureFill(fill: ExcelFillModel): number {
    return ensureStyleItem(this.styles.fills, fill, cloneFill);
  }

  ensureBorder(border: ExcelBorderModel): number {
    return ensureStyleItem(this.styles.borders, border, cloneBorder);
  }

  ensureNumberFormat(code: string): number {
    const builtIn = builtInNumberFormatId(code);
    if (builtIn !== undefined) return builtIn;
    const existing = this.styles.numberFormats.find((format) => format.code === code);
    if (existing) return existing.id;
    const maxId = this.styles.numberFormats.reduce((max, format) => Math.max(max, format.id), 163);
    const next = { id: maxId + 1, code };
    this.styles.numberFormats.push(next);
    return next.id;
  }

  fontById(id: number | undefined): ExcelFontModel {
    return cloneFont(this.styles.fonts[id ?? 0] ?? {});
  }

  fillById(id: number | undefined): ExcelFillModel {
    return cloneFill(this.styles.fills[id ?? 0] ?? {});
  }

  borderById(id: number | undefined): ExcelBorderModel {
    return cloneBorder(this.styles.borders[id ?? 0] ?? {});
  }

  private matrix<T>(read: (cell: ExcelCellModel) => T): T[][] {
    const rows: T[][] = [];
    for (let r = 0; r < this.ref.rowCount; r++) {
      const row: T[] = [];
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        row.push(read(this.cellsByAddress.get(address) ?? blankCell(address)));
      }
      rows.push(row);
    }
    return rows;
  }

  private assertOffset(rowOffset: number, columnOffset: number): void {
    if (
      !Number.isInteger(rowOffset) ||
      !Number.isInteger(columnOffset) ||
      rowOffset < 0 ||
      columnOffset < 0 ||
      rowOffset >= this.ref.rowCount ||
      columnOffset >= this.ref.columnCount
    ) {
      throw new Error(
        `ExcelRange.getCell: offset out of bounds: row=${rowOffset}, column=${columnOffset}`,
      );
    }
  }

  private assertMatrixShape<T>(matrix: T[][], operation: string): void {
    if (matrix.length !== this.ref.rowCount) {
      throw new Error(
        `${operation}: expected ${this.ref.rowCount} rows, received ${matrix.length}`,
      );
    }
    for (let r = 0; r < matrix.length; r++) {
      if (!Array.isArray(matrix[r]) || matrix[r]!.length !== this.ref.columnCount) {
        throw new Error(
          `${operation}: expected row ${r} to contain ${this.ref.columnCount} columns`,
        );
      }
    }
  }

  private setCellValue(address: string, value: ExcelCellValue): void {
    if (value === null) {
      this.deleteCell(address);
      return;
    }
    const cell = this.ensureCell(address);
    cell.value = value;
    cell.kind =
      typeof value === 'boolean' ? 'boolean' : typeof value === 'number' ? 'number' : 'string';
    delete cell.formula;
    delete cell.rawValue;
    delete cell.dataType;
    delete cell.sharedStringIndex;
  }

  private setCellFormula(address: string, formula: string | null): void {
    if (formula === null) {
      this.deleteCell(address);
      return;
    }
    const cell = this.ensureCell(address);
    cell.kind = 'formula';
    cell.formula = formula.startsWith('=') ? formula.slice(1) : formula;
    cell.value = null;
    delete cell.rawValue;
    delete cell.dataType;
    delete cell.sharedStringIndex;
  }

  private ensureCell(address: string): ExcelCellModel {
    const existing = this.cellsByAddress.get(address);
    if (existing) return existing;
    const cell = blankCell(address);
    this.sheet.cells.push(cell);
    this.cellsByAddress.set(address, cell);
    return cell;
  }

  private deleteCell(address: string): void {
    const cell = this.cellsByAddress.get(address);
    if (!cell) return;
    this.cellsByAddress.delete(address);
    const index = this.sheet.cells.indexOf(cell);
    if (index !== -1) this.sheet.cells.splice(index, 1);
  }

  private clearCellContent(address: string): void {
    const cell = this.cellsByAddress.get(address);
    if (!cell) return;
    if (cell.styleId === undefined) {
      this.deleteCell(address);
      return;
    }
    cell.kind = 'blank';
    cell.value = null;
    delete cell.formula;
    delete cell.rawValue;
    delete cell.dataType;
    delete cell.sharedStringIndex;
  }

  private ensureRow(index: number) {
    let row = this.sheet.rows.find((item) => item.index === index);
    if (!row) {
      row = { index };
      this.sheet.rows.push(row);
      this.sheet.rows.sort((a, b) => a.index - b.index);
    }
    return row;
  }

  private ensureColumn(index: number) {
    let column = this.sheet.columns.find((item) => item.min === index && item.max === index);
    if (!column) {
      column = { min: index, max: index };
      this.sheet.columns.push(column);
    }
    return column;
  }

  private firstCell(): ExcelCellModel | undefined {
    for (let r = 0; r < this.ref.rowCount; r++) {
      for (let c = 0; c < this.ref.columnCount; c++) {
        const address = formatCellAddress(this.ref.startRow + r, this.ref.startColumn + c);
        const cell = this.cellsByAddress.get(address);
        if (cell) return cell;
      }
    }
    return undefined;
  }
}

export class ExcelCell {
  private readonly range: ExcelRange;

  constructor(
    sheet: ExcelSheetModel,
    address: string,
    styles?: ExcelStylesModel,
    recordMutation?: ExcelMutationRecorder,
  ) {
    this.range = new ExcelRange(sheet, address, styles, recordMutation);
  }

  address(): string {
    return this.range.address();
  }

  value(): ExcelCellValue {
    return this.range.values()[0]?.[0] ?? null;
  }

  formula(): string | null {
    return this.range.formulas()[0]?.[0] ?? null;
  }

  text(): string {
    return this.range.text()[0]?.[0] ?? '';
  }

  valueType(): ExcelCellKind {
    return this.range.valueTypes()[0]?.[0] ?? 'blank';
  }

  rawValue(): string | null {
    return this.range.rawValues()[0]?.[0] ?? null;
  }

  styleId(): number | undefined {
    return this.range.styleIds()[0]?.[0];
  }

  setStyleId(styleId: number | undefined): void {
    this.range.setStyleId(styleId);
  }

  setValue(value: ExcelCellValue): void {
    this.range.setValues([[value]]);
  }

  setFormula(formula: string | null): void {
    this.range.setFormulas([[formula]]);
  }

  clear(): void {
    this.range.clear();
  }

  toRange(): ExcelRange {
    return this.range;
  }

  toModel(): ExcelCellModel | null {
    return this.range.toModel()[0] ? { ...this.range.toModel()[0]! } : null;
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

  numberFormatCode(): string | undefined {
    return this.range.getResolvedFormat().numberFormatCode;
  }

  setNumberFormatCode(code: string | undefined): void {
    this.range.updateFormat((format) => {
      if (code === undefined) {
        delete format.numberFormatCode;
        delete format.numberFormatId;
        delete format.applyNumberFormat;
      } else {
        format.numberFormatCode = code;
        format.numberFormatId = this.range.ensureNumberFormat(code);
        format.applyNumberFormat = true;
      }
      return format;
    });
  }
}

export class ExcelRangeFontFormat {
  constructor(private readonly range: ExcelRange) {}

  value(): ExcelFontModel {
    return this.range.fontById(this.range.getResolvedFormat().fontId);
  }

  setValue(font: ExcelFontModel, options: { recordUnsafe?: boolean } = {}): void {
    const normalized = {
      ...font,
      color: normalizeExcelColor(font.color, 'ExcelRange.format.font.setValue'),
    };
    removeUndefined(normalized);
    this.range.updateFormat(
      (format) => {
        format.fontId = this.range.ensureFont(normalized);
        format.applyFont = true;
        return format;
      },
      options,
    );
  }

  bold(): boolean | undefined {
    return this.value().bold;
  }

  setBold(value: boolean | undefined): void {
    this.patch({ bold: value });
  }

  italic(): boolean | undefined {
    return this.value().italic;
  }

  setItalic(value: boolean | undefined): void {
    this.patch({ italic: value });
  }

  underline(): boolean | undefined {
    return this.value().underline;
  }

  setUnderline(value: boolean | undefined): void {
    this.patch({ underline: value });
  }

  color(): ExcelColorModel | undefined {
    return this.value().color;
  }

  setColor(color: ExcelColorModel | undefined): void {
    const normalized = normalizeExcelColor(color, 'ExcelRange.format.font.setColor');
    this.patch({ color: normalized }, { recordUnsafe: false });
    this.range.recordReplayableMutation({
      kind: 'range.format.font.setColor',
      address: this.range.address(),
      color: normalized,
    });
  }

  size(): number | undefined {
    return this.value().size;
  }

  setSize(size: number | undefined): void {
    this.patch({ size });
  }

  name(): string | undefined {
    return this.value().name;
  }

  setName(name: string | undefined): void {
    this.patch({ name });
  }

  private patch(patch: Partial<ExcelFontModel>, options: { recordUnsafe?: boolean } = {}): void {
    this.range.updateFormat(
      (format) => {
        const next = {
          ...this.range.fontById(format.fontId),
          ...patch,
        };
        removeUndefined(next);
        format.fontId = this.range.ensureFont(next);
        format.applyFont = true;
        return format;
      },
      options,
    );
  }
}

export class ExcelRangeFillFormat {
  constructor(private readonly range: ExcelRange) {}

  value(): ExcelFillModel {
    return this.range.fillById(this.range.getResolvedFormat().fillId);
  }

  setValue(fill: ExcelFillModel, options: { recordUnsafe?: boolean } = {}): void {
    const normalized = {
      ...fill,
      patternType: fill.patternType ?? (fill.foregroundColor ? 'solid' : undefined),
      foregroundColor: normalizeExcelColor(
        fill.foregroundColor,
        'ExcelRange.format.fill.setValue.foregroundColor',
      ),
      backgroundColor: normalizeExcelColor(
        fill.backgroundColor,
        'ExcelRange.format.fill.setValue.backgroundColor',
      ),
    };
    removeUndefined(normalized);
    this.range.updateFormat(
      (format) => {
        format.fillId = this.range.ensureFill(normalized);
        format.applyFill = true;
        return format;
      },
      options,
    );
  }

  patternType(): string | undefined {
    return this.value().patternType;
  }

  setPatternType(patternType: string | undefined): void {
    this.patch({ patternType });
  }

  foregroundColor(): ExcelColorModel | undefined {
    return this.value().foregroundColor;
  }

  setForegroundColor(color: ExcelColorModel | undefined): void {
    const normalized = normalizeExcelColor(color, 'ExcelRange.format.fill.setForegroundColor');
    this.patch(
      {
        foregroundColor: normalized,
        patternType: normalized ? 'solid' : this.patternType(),
      },
      { recordUnsafe: false },
    );
    this.range.recordReplayableMutation({
      kind: 'range.format.fill.setForegroundColor',
      address: this.range.address(),
      color: normalized,
    });
  }

  backgroundColor(): ExcelColorModel | undefined {
    return this.value().backgroundColor;
  }

  setBackgroundColor(color: ExcelColorModel | undefined): void {
    this.patch({
      backgroundColor: normalizeExcelColor(color, 'ExcelRange.format.fill.setBackgroundColor'),
    });
  }

  private patch(patch: Partial<ExcelFillModel>, options: { recordUnsafe?: boolean } = {}): void {
    const next = { ...this.value(), ...patch };
    removeUndefined(next);
    this.setValue(next, options);
  }
}

export class ExcelRangeBorderFormat {
  constructor(private readonly range: ExcelRange) {}

  value(): ExcelBorderModel {
    return this.range.borderById(this.range.getResolvedFormat().borderId);
  }

  setValue(border: ExcelBorderModel): void {
    const normalized = normalizeBorder(border, 'ExcelRange.format.border.setValue');
    this.range.updateFormat(
      (format) => {
        format.borderId = this.range.ensureBorder(normalized);
        format.applyBorder = true;
        return format;
      },
      { recordUnsafe: false },
    );
    this.range.recordReplayableMutation({
      kind: 'range.format.border.setValue',
      address: this.range.address(),
      border: normalized,
    });
  }

  setLeft(side: ExcelBorderModel['left']): void {
    this.patch({ left: side });
  }

  setRight(side: ExcelBorderModel['right']): void {
    this.patch({ right: side });
  }

  setTop(side: ExcelBorderModel['top']): void {
    this.patch({ top: side });
  }

  setBottom(side: ExcelBorderModel['bottom']): void {
    this.patch({ bottom: side });
  }

  private patch(patch: Partial<ExcelBorderModel>): void {
    const next = { ...this.value(), ...patch };
    removeUndefined(next);
    this.setValue(next);
  }
}

function normalizeBorder(border: ExcelBorderModel, context: string): ExcelBorderModel {
  const normalized: ExcelBorderModel = {};
  for (const name of ['left', 'right', 'top', 'bottom', 'diagonal'] as const) {
    const side = border[name];
    if (!side) continue;
    normalized[name] = {
      ...side,
      color: normalizeExcelColor(side.color, `${context}.${name}.color`),
    };
    removeUndefined(normalized[name]!);
  }
  return normalized;
}

export class ExcelRangeAlignmentFormat {
  constructor(private readonly range: ExcelRange) {}

  value(): ExcelAlignmentModel | undefined {
    const value = this.range.getResolvedFormat().alignment;
    return value ? { ...value } : undefined;
  }

  setValue(alignment: ExcelAlignmentModel | undefined): void {
    this.range.updateFormat((format) => {
      if (alignment === undefined) {
        delete format.alignment;
        delete format.applyAlignment;
      } else {
        format.alignment = { ...alignment };
        format.applyAlignment = true;
      }
      return format;
    });
  }

  setHorizontal(horizontal: string | undefined): void {
    this.patch({ horizontal });
  }

  setVertical(vertical: string | undefined): void {
    this.patch({ vertical });
  }

  setWrapText(wrapText: boolean | undefined): void {
    this.patch({ wrapText });
  }

  private patch(patch: Partial<ExcelAlignmentModel>): void {
    const next = { ...(this.value() ?? {}), ...patch };
    removeUndefined(next);
    this.setValue(Object.keys(next).length ? next : undefined);
  }
}

function blankCell(address: string): ExcelCellModel {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  return {
    address,
    row: match ? Number(match[2]) : 0,
    column: match ? columnNameToIndexSafe(match[1]!) : 0,
    kind: 'blank',
    value: null,
  };
}

function columnNameToIndexSafe(columnName: string): number {
  let value = 0;
  for (const ch of columnName) value = value * 26 + (ch.charCodeAt(0) - 64);
  return value;
}

function cellToText(cell: ExcelCellModel): string {
  if (cell.value === null) return '';
  return String(cell.value);
}

function normalizeSheetCells(sheet: ExcelSheetModel): void {
  sheet.cells.sort((a, b) => a.row - b.row || a.column - b.column);
  sheet.dimension = computeDimension(sheet.cells);
}

function normalizeColumns(sheet: ExcelSheetModel): void {
  sheet.columns.sort((a, b) => a.min - b.min || a.max - b.max);
}

function createEmptyStyles(): ExcelStylesModel {
  return {
    numberFormats: [],
    fonts: [{}],
    fills: [{}, { patternType: 'gray125' }],
    borders: [{}],
    cellFormats: [{}],
  };
}

function ensureCellFormat(styles: ExcelStylesModel, format: ExcelCellFormatModel): number {
  const normalized = cloneCellFormat(format);
  removeUndefined(normalized);
  return ensureStyleItem(styles.cellFormats, normalized, cloneCellFormat);
}

function ensureStyleItem<T>(items: T[], item: T, clone: (value: T) => T): number {
  const normalized = clone(item);
  removeUndefined(normalized as object);
  const key = stableStringify(normalized);
  const existing = items.findIndex((candidate) => stableStringify(candidate) === key);
  if (existing !== -1) return existing;
  items.push(normalized);
  return items.length - 1;
}

function cloneCellFormat(format: ExcelCellFormatModel): ExcelCellFormatModel {
  return {
    ...format,
    alignment: format.alignment ? { ...format.alignment } : undefined,
  };
}

function cloneFont(font: ExcelFontModel): ExcelFontModel {
  return { ...font, color: font.color ? { ...font.color } : undefined };
}

function cloneFill(fill: ExcelFillModel): ExcelFillModel {
  return {
    ...fill,
    foregroundColor: fill.foregroundColor ? { ...fill.foregroundColor } : undefined,
    backgroundColor: fill.backgroundColor ? { ...fill.backgroundColor } : undefined,
  };
}

function cloneBorder(border: ExcelBorderModel): ExcelBorderModel {
  return {
    left: border.left
      ? { ...border.left, color: border.left.color ? { ...border.left.color } : undefined }
      : undefined,
    right: border.right
      ? { ...border.right, color: border.right.color ? { ...border.right.color } : undefined }
      : undefined,
    top: border.top
      ? { ...border.top, color: border.top.color ? { ...border.top.color } : undefined }
      : undefined,
    bottom: border.bottom
      ? { ...border.bottom, color: border.bottom.color ? { ...border.bottom.color } : undefined }
      : undefined,
    diagonal: border.diagonal
      ? {
          ...border.diagonal,
          color: border.diagonal.color ? { ...border.diagonal.color } : undefined,
        }
      : undefined,
  };
}

function removeUndefined(value: object): void {
  for (const key of Object.keys(value)) {
    if ((value as Record<string, unknown>)[key] === undefined)
      delete (value as Record<string, unknown>)[key];
  }
}

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
  return `{${Object.keys(value as Record<string, unknown>)
    .sort()
    .map(
      (key) => `${JSON.stringify(key)}:${stableStringify((value as Record<string, unknown>)[key])}`,
    )
    .join(',')}}`;
}

function builtInNumberFormatId(code: string): number | undefined {
  const builtIns: Record<string, number> = {
    General: 0,
    '0': 1,
    '0.00': 2,
    '#,##0': 3,
    '#,##0.00': 4,
    '0%': 9,
    '0.00%': 10,
  };
  return builtIns[code];
}

function computeDimension(cells: ExcelCellModel[]): string | undefined {
  if (cells.length === 0) return undefined;
  let minRow = Number.POSITIVE_INFINITY;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let maxColumn = 0;
  for (const cell of cells) {
    minRow = Math.min(minRow, cell.row);
    minColumn = Math.min(minColumn, cell.column);
    maxRow = Math.max(maxRow, cell.row);
    maxColumn = Math.max(maxColumn, cell.column);
  }
  return formatRangeAddress(minRow, minColumn, maxRow - minRow + 1, maxColumn - minColumn + 1);
}
