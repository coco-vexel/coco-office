import { describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { ExcelWorkbook } from '../index.js';
import type { ExcelWorkbookModel } from '../model.js';
import { ExcelRange } from '../range.js';

describe('clearing styled cell content', () => {
  test('treats an empty formula string as content clearing and rejects an empty expression before any writes', () => {
    const sheet: ExcelWorkbookModel['sheets'][number] = {
      id: '1', name: 'Sheet1', sheetViews: [], columns: [], rows: [], mergedRanges: [],
      hyperlinks: [], tables: [], dataValidations: [], conditionalFormats: [], comments: [],
      threadedComments: [], drawings: [], charts: [], images: [],
      cells: [
        { address: 'A1', row: 1, column: 1, kind: 'formula' as const, value: 1, formula: '1', styleId: 0 },
        { address: 'B1', row: 1, column: 2, kind: 'formula' as const, value: 2, formula: '2', styleId: 1 },
      ],
    };
    const range = new ExcelRange(sheet, 'A1:B1');
    expect(() => range.setFormulas([['3', '=']])).toThrow(/empty expression/);
    expect(range.formulas()).toEqual([['1', '2']]);
    range.setFormulas([['', '']]);
    expect(range.formulas()).toEqual([[null, null]]);
    expect(range.values()).toEqual([[null, null]]);
    expect(range.styleIds()).toEqual([[0, 1]]);
    expect(sheet.cells.every((cell) => cell.kind === 'blank')).toBe(true);
  });

  test.each(['patch', 'rewrite'])('keeps null, empty text, whitespace, zero and false distinct through XLSX %s', async (mode) => {
    const source: ExcelWorkbookModel = {
      sharedStrings: [], definedNames: [],
      styles: { numberFormats: [], fonts: [{}], fills: [{}], borders: [{}], cellFormats: [{}] },
      sheets: [{
        id: '1', name: 'Sheet1', sheetViews: [], columns: [], rows: [], mergedRanges: [],
        hyperlinks: [], tables: [], dataValidations: [], conditionalFormats: [], comments: [],
        threadedComments: [], drawings: [], charts: [], images: [], cells: [],
      }],
    };
    const seed = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    await seed.worksheets.getItem('Sheet1').getRange('A1:E1').setValues([['a', 'b', 'c', 'd', 'e']]);
    const initial = await seed.toZip();
    await seed.dispose();
    const workbook = await ExcelWorkbook.fromZip(initial, { mode: 'inline' });
    try {
      const range = workbook.worksheets.getItem('Sheet1').getRange('A1:E1');
      await range.setValues([[null, '', ' ', 0, false]]);
      if (mode === 'rewrite') await range.setColumnWidth(12);
      const bytes = await workbook.toZip();
      const zip = await JSZip.loadAsync(bytes);
      expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toMatch(/<t xml:space="preserve"> <\/t>/);
      const reopened = await ExcelWorkbook.fromZip(bytes, { mode: 'inline' });
      try {
        expect(await reopened.worksheets.getItem('Sheet1').getRange('A1:E1').values()).toEqual([[null, '', ' ', 0, false]]);
      } finally { await reopened.dispose(); }
      // Imported t="str" cells with an empty <v> must also stay empty text, not null.
      const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
      zip.file('xl/worksheets/sheet1.xml', sheetXml.replace(/<c\b[^>]*r="B1"[^>]*>[\s\S]*?<\/c>/, '<c r="B1" t="str"><v/></c>'));
      const imported = await ExcelWorkbook.fromZip(await zip.generateAsync({ type: 'uint8array' }), { mode: 'inline' });
      try { expect(await imported.worksheets.getItem('Sheet1').getRange('B1').values()).toEqual([['']]); }
      finally { await imported.dispose(); }
    } finally { await workbook.dispose(); }
  });

  test.each(['values', 'formulas'])('preserves formatting when clearing %s, including after XLSX writeback', async (method) => {
    const side = { style: 'thin', color: { rgb: '#2C7A69' } };
    const border = { top: side, bottom: side, left: side, right: side };
    const source: ExcelWorkbookModel = {
      sharedStrings: [], definedNames: [],
      styles: {
        numberFormats: [], fonts: [{ name: 'KaiTi', size: 26 }], fills: [{}],
        borders: [{}, border], cellFormats: [{}, { fontId: 0, borderId: 1, alignment: { horizontal: 'center' } }],
      },
      sheets: [{
        id: '1', name: 'Sheet1', sheetViews: [], columns: [], rows: [], mergedRanges: [],
        hyperlinks: [], tables: [], dataValidations: [], conditionalFormats: [], comments: [],
        threadedComments: [], drawings: [], charts: [], images: [],
        cells: [
          { address: 'A1', row: 1, column: 1, kind: 'blank', value: null, styleId: 1 },
          { address: 'B1', row: 1, column: 2, kind: 'string', value: '田', styleId: 1 },
          { address: 'C1', row: 1, column: 3, kind: 'formula', value: 1, formula: '1+0', styleId: 1 },
          { address: 'D1', row: 1, column: 4, kind: 'number', value: 42 },
        ],
      }],
    };
    const seed = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const bytes = await seed.toZip();
    await seed.dispose();
    const workbook = await ExcelWorkbook.fromZip(bytes, { mode: 'inline' });
    try {
      const range = workbook.worksheets.getItem('Sheet1').getRange('A1:D1');
      if (method === 'values') await range.setValues([[null, null, null, null]]);
      else await range.setFormulas([[null, null, null, null]]);
      expect(await range.styleIds()).toEqual([[1, 1, 1, undefined]]);
      const model = await workbook.toModel();
      expect(model.sheets[0]!.cells).toEqual([
        { address: 'A1', row: 1, column: 1, kind: 'blank', value: null, styleId: 1 },
        { address: 'B1', row: 1, column: 2, kind: 'blank', value: null, styleId: 1 },
        { address: 'C1', row: 1, column: 3, kind: 'blank', value: null, styleId: 1 },
      ]);
      const reopened = await ExcelWorkbook.fromZip(await workbook.toZip(), { mode: 'inline' });
      try {
        expect(await reopened.worksheets.getItem('Sheet1').getRange('A1:D1').styleIds()).toEqual([[1, 1, 1, undefined]]);
        expect((await reopened.toModel()).styles.borders[1]).toMatchObject(border);
        expect(await reopened.worksheets.getItem('Sheet1').getRange('A1:C1').formulas()).toEqual([[null, null, null]]);
      } finally { await reopened.dispose(); }
    } finally { await workbook.dispose(); }
  });
});
