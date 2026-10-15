import { describe, expect, test } from 'vitest';
import {
  ExcelRangeCore as ExcelRange,
  type ExcelSheetModel,
  type ExcelStylesModel,
} from '../index.js';

const sheet: ExcelSheetModel = {
  id: '1',
  name: 'Sheet1',
  dimension: 'A1:D4',
  sheetViews: [],
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
  cells: [
    { address: 'A1', row: 1, column: 1, kind: 'string', value: 'a' },
    { address: 'C1', row: 1, column: 3, kind: 'number', value: 3 },
    { address: 'B2', row: 2, column: 2, kind: 'formula', value: 4, formula: 'A2*2' },
    { address: 'D4', row: 4, column: 4, kind: 'boolean', value: true },
  ],
};

describe('ExcelRange', () => {
  test('returns rectangular matrices with blank cells filled', () => {
    const range = new ExcelRange(sheet, 'A1:C2');

    expect(range.values()).toEqual([
      ['a', null, 3],
      [null, 4, null],
    ]);
    expect(range.formulas()).toEqual([
      [null, null, null],
      [null, 'A2*2', null],
    ]);
    expect(range.valueTypes()).toEqual([
      ['string', 'blank', 'number'],
      ['blank', 'formula', 'blank'],
    ]);
    expect(range.text()).toEqual([
      ['a', '', '3'],
      ['', '4', ''],
    ]);
  });

  test('supports sheet-qualified local addresses', () => {
    expect(new ExcelRange(sheet, 'Sheet1!$C$1:$A$1').address()).toBe('A1:C1');
    expect(() => new ExcelRange(sheet, 'Other!A1')).toThrow(/does not match worksheet/);
  });

  test('supports cell, row, column, resize, offset and slice helpers', () => {
    const range = new ExcelRange(sheet, 'A1:D4');

    expect(range.getCell(1, 1).address()).toBe('B2');
    expect(range.getCell(1, 1).formula()).toBe('A2*2');
    expect(range.getRow(1).address()).toBe('A2:D2');
    expect(range.getColumn(2).address()).toBe('C1:C4');
    expect(range.resize(2, 2).address()).toBe('A1:B2');
    expect(range.offset(1, 1, 2, 2).address()).toBe('B2:C3');
    expect(range.slice(1, 1, 2, 2).values()).toEqual([
      [4, null],
      [null, null],
    ]);
  });

  test('rejects out-of-bounds offsets and invalid shapes', () => {
    const range = new ExcelRange(sheet, 'A1:D4');

    expect(() => range.getCell(4, 0)).toThrow(/out of bounds/);
    expect(() => range.getRow(-1)).toThrow(/out of bounds/);
    expect(() => range.getColumn(4)).toThrow(/out of bounds/);
    expect(() => range.resize(0, 1)).toThrow(/invalid row count/);
    expect(() => range.offset(-1, 0)).toThrow(/invalid row index/);
    expect(() => range.slice(3, 3, 2, 1)).toThrow(/out of bounds/);
  });

  test('mutates values, formulas and cleared cells in the sheet model', () => {
    const model: ExcelSheetModel = {
      ...sheet,
      cells: sheet.cells.map((cell) => ({ ...cell })),
    };

    const range = new ExcelRange(model, 'E5:F6');
    range.setValues([
      ['new', 7],
      [false, null],
    ]);

    expect(new ExcelRange(model, 'E5:F6').values()).toEqual([
      ['new', 7],
      [false, null],
    ]);
    expect(model.dimension).toBe('A1:F6');

    new ExcelRange(model, 'F5:F5').setFormulas([['=E5*2']]);
    expect(new ExcelRange(model, 'F5').formulas()).toEqual([['E5*2']]);
    expect(new ExcelRange(model, 'F5').values()).toEqual([[null]]);

    new ExcelRange(model, 'A1:F5').clear();
    expect(new ExcelRange(model, 'A1:F6').values()).toEqual([
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, false, null],
    ]);
    expect(model.dimension).toBe('E6');
  });

  test('cell mutation delegates to its one-cell range', () => {
    const model: ExcelSheetModel = {
      ...sheet,
      cells: sheet.cells.map((cell) => ({ ...cell })),
    };

    const cell = new ExcelRange(model, 'A1').getCell(0, 0);
    cell.setValue('updated');
    expect(cell.value()).toBe('updated');

    cell.setFormula('B1+C1');
    expect(cell.formula()).toBe('B1+C1');
    expect(cell.value()).toBeNull();

    cell.clear();
    expect(cell.valueType()).toBe('blank');
  });

  test('mutates basic range formatting metadata', () => {
    const model: ExcelSheetModel = {
      ...sheet,
      cells: sheet.cells.map((cell) => ({ ...cell })),
      rows: [],
      columns: [],
      mergedRanges: [],
    };

    const range = new ExcelRange(model, 'B2:C3');
    range.setStyleId(2);
    range.setRowHeight(24);
    range.setColumnWidth(18);
    range.merge();

    expect(new ExcelRange(model, 'B2:C3').styleIds()).toEqual([
      [2, 2],
      [2, 2],
    ]);
    expect(new ExcelRange(model, 'B2:C3').rowHeights()).toEqual([24, 24]);
    expect(new ExcelRange(model, 'B2:C3').columnWidths()).toEqual([18, 18]);
    expect(model.rows).toEqual([
      { index: 2, height: 24, customHeight: true },
      { index: 3, height: 24, customHeight: true },
    ]);
    expect(model.columns).toEqual([
      { min: 2, max: 2, width: 18, customWidth: true },
      { min: 3, max: 3, width: 18, customWidth: true },
    ]);
    expect(model.mergedRanges).toEqual([{ address: 'B2:C3' }]);

    range.unmerge();
    range.setStyleId(undefined);
    range.setRowHeight(undefined);
    range.setColumnWidth(undefined);

    expect(new ExcelRange(model, 'B2:C3').styleIds()).toEqual([
      [undefined, undefined],
      [undefined, undefined],
    ]);
    expect(model.rows).toEqual([{ index: 2 }, { index: 3 }]);
    expect(model.columns).toEqual([
      { min: 2, max: 2 },
      { min: 3, max: 3 },
    ]);
    expect(model.mergedRanges).toEqual([]);
  });

  test('clears covered cell content when merging a range', () => {
    const model: ExcelSheetModel = {
      ...sheet,
      cells: [
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'master' },
        { address: 'B1', row: 1, column: 2, kind: 'string', value: 'covered' },
        { address: 'A2', row: 2, column: 1, kind: 'formula', value: null, formula: 'B1*2' },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 4, styleId: 1 },
      ],
      mergedRanges: [],
    };

    const range = new ExcelRange(model, 'A1:B2');
    range.merge();
    range.unmerge();

    expect(new ExcelRange(model, 'A1:B2').values()).toEqual([
      ['master', null],
      [null, null],
    ]);
    expect(new ExcelRange(model, 'A1:B2').formulas()).toEqual([
      [null, null],
      [null, null],
    ]);
    expect(new ExcelRange(model, 'B2').styleIds()).toEqual([[1]]);
  });

  test('uses whole-column and whole-row references for dimension metadata', () => {
    const model: ExcelSheetModel = {
      ...sheet,
      cells: sheet.cells.map((cell) => ({ ...cell })),
      rows: [],
      columns: [],
    };

    const columns = new ExcelRange(model, 'A:C');
    columns.setColumnWidth(21);
    expect(columns.address()).toBe('A1:C1048576');
    expect(new ExcelRange(model, 'A1:C1').columnWidths()).toEqual([21, 21, 21]);
    expect(model.columns).toEqual([
      { min: 1, max: 1, width: 21, customWidth: true },
      { min: 2, max: 2, width: 21, customWidth: true },
      { min: 3, max: 3, width: 21, customWidth: true },
    ]);

    const rows = new ExcelRange(model, '1:2');
    rows.setRowHeight(19);
    expect(rows.address()).toBe('A1:XFD2');
    expect(new ExcelRange(model, 'A1:A2').rowHeights()).toEqual([19, 19]);
    expect(model.rows).toEqual([
      { index: 1, height: 19, customHeight: true },
      { index: 2, height: 19, customHeight: true },
    ]);
  });

  test('mutates fine-grained format through the style table', () => {
    const model: ExcelSheetModel = {
      ...sheet,
      cells: sheet.cells.map((cell) => ({ ...cell })),
    };
    const styles: ExcelStylesModel = {
      numberFormats: [],
      fonts: [{}],
      fills: [{}, { patternType: 'gray125' }],
      borders: [{}],
      cellFormats: [{}],
    };

    const range = new ExcelRange(model, 'A1:B2', styles);
    range.format.font.setBold(true);
    range.format.font.setItalic(true);
    range.format.font.setUnderline(true);
    range.format.font.setName('Calibri');
    range.format.font.setSize(12);
    range.format.font.setColor({ rgb: '#FF0000' });
    range.format.fill.setForegroundColor({ rgb: '#FFFF00' });
    range.format.border.setBottom({ style: 'thin', color: { rgb: '#000000' } });
    range.format.alignment.setHorizontal('center');
    range.format.alignment.setWrapText(true);
    range.format.setNumberFormatCode('0.00%');

    const styleIds = new ExcelRange(model, 'A1:B2', styles).styleIds();
    expect(styleIds.flat().every((id) => id === styleIds[0]?.[0])).toBe(true);
    const format = styles.cellFormats[styleIds[0]![0]!]!;
    expect(styles.fonts[format.fontId!]).toMatchObject({
      bold: true,
      italic: true,
      underline: true,
      name: 'Calibri',
      size: 12,
      color: { rgb: '#FF0000' },
    });
    expect(styles.fills[format.fillId!]).toMatchObject({
      patternType: 'solid',
      foregroundColor: { rgb: '#FFFF00' },
    });
    expect(styles.borders[format.borderId!]).toMatchObject({
      bottom: { style: 'thin', color: { rgb: '#000000' } },
    });
    expect(format).toMatchObject({
      numberFormatId: 10,
      numberFormatCode: '0.00%',
      applyNumberFormat: true,
      applyFont: true,
      applyFill: true,
      applyBorder: true,
      applyAlignment: true,
      alignment: { horizontal: 'center', wrapText: true },
    });
    expect(new ExcelRange(model, 'A1', styles).format.font.bold()).toBe(true);
    expect(new ExcelRange(model, 'A1', styles).format.numberFormatCode()).toBe('0.00%');
  });

  test('preserves per-cell font attributes when patching a range font color', () => {
    const model: ExcelSheetModel = {
      ...sheet,
      cells: sheet.cells.map((cell) => ({ ...cell })),
    };
    const styles: ExcelStylesModel = {
      numberFormats: [],
      fonts: [{}],
      fills: [{}, { patternType: 'gray125' }],
      borders: [{}],
      cellFormats: [{}],
    };

    new ExcelRange(model, 'A1', styles).format.font.setBold(true);
    new ExcelRange(model, 'B1', styles).format.font.setItalic(true);
    new ExcelRange(model, 'A1:B1', styles).format.font.setColor({ rgb: '#C00000' });

    const a1Format = styles.cellFormats[new ExcelRange(model, 'A1', styles).styleIds()[0]![0]!]!;
    const b1Format = styles.cellFormats[new ExcelRange(model, 'B1', styles).styleIds()[0]![0]!]!;
    expect(styles.fonts[a1Format.fontId!]).toMatchObject({
      bold: true,
      color: { rgb: '#C00000' },
    });
    expect(styles.fonts[b1Format.fontId!]).toMatchObject({
      italic: true,
      color: { rgb: '#C00000' },
    });
  });

  test('rejects mutation matrices with unexpected dimensions', () => {
    const range = new ExcelRange(
      { ...sheet, cells: sheet.cells.map((cell) => ({ ...cell })) },
      'A1:B2',
    );

    expect(() => range.setValues([[1, 2]])).toThrow(/expected 2 rows/);
    expect(() => range.setFormulas([['A1'], ['A2', 'B2', 'C2']])).toThrow(/expected row 0/);
  });

  test('rejects non-canonical color inputs at the format boundary', () => {
    const range = new ExcelRange(
      { ...sheet, cells: sheet.cells.map((cell) => ({ ...cell })) },
      'A1',
    );

    expect(() => range.format.font.setColor({ rgb: 'FFFF0000' })).toThrow(/#RRGGBB/);
    expect(() => range.format.fill.setForegroundColor({ rgb: 'D9D9D9' })).toThrow(/#RRGGBB/);
    expect(() =>
      range.format.border.setBottom({
        style: 'thin',
        color: { rgb: '#000000', theme: 1 } as never,
      }),
    ).toThrow(/only the rgb field/);
  });
});
