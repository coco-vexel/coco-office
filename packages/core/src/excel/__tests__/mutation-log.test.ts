import { describe, expect, test } from 'vitest';
import { ExcelWorkbookCore, type ExcelWorkbookModel } from '../index.js';

function model(): ExcelWorkbookModel {
  return {
    sharedStrings: [],
    definedNames: [],
    styles: {
      numberFormats: [],
      fonts: [{}],
      fills: [{}],
      borders: [{}],
      cellFormats: [{}],
    },
    sheets: [
      {
        id: '1',
        name: 'Sheet1',
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
        cells: [],
      },
    ],
  };
}

describe('Excel mutation log', () => {
  test('records first-phase replayable range mutations', () => {
    const workbook = ExcelWorkbookCore.fromModel(model());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.getRange('C2:D3').setValues([
      ['张三', '13812345678'],
      ['李四', '15987654321'],
    ]);
    sheet.getRange('E2:E2').setFormulas([['=D2']]);
    sheet.getRange('F1:G1').merge();
    sheet.getRange('F1:G1').unmerge();
    sheet.getRange('H1').clear();

    expect(workbook.getMutations()).toMatchObject([
      {
        kind: 'range.setValues',
        sheetId: '1',
        sheetName: 'Sheet1',
        address: 'C2:D3',
        values: [
          ['张三', '13812345678'],
          ['李四', '15987654321'],
        ],
      },
      {
        kind: 'range.setFormulas',
        sheetName: 'Sheet1',
        address: 'E2',
        formulas: [['=D2']],
      },
      { kind: 'range.merge', sheetName: 'Sheet1', address: 'F1:G1' },
      { kind: 'range.unmerge', sheetName: 'Sheet1', address: 'F1:G1' },
      { kind: 'range.clear', sheetName: 'Sheet1', address: 'H1' },
    ]);
  });

  test('records worksheet lifecycle and unsafe fallback mutations', () => {
    const workbook = ExcelWorkbookCore.fromModel(model());
    const added = workbook.worksheets.add('Added');
    added.rename('Renamed');
    added.getRange('A1').setStyleId(1);
    workbook.worksheets.deleteItem('Renamed');

    expect(workbook.getMutations().map((item) => item.kind)).toEqual([
      'worksheet.add',
      'worksheet.rename',
      'unsafe',
      'worksheet.delete',
    ]);

    workbook.clearMutations();
    expect(workbook.getMutations()).toEqual([]);
  });

  test('records common range format mutations as replayable operations', () => {
    const workbook = ExcelWorkbookCore.fromModel(model());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.getRange('1:1').setRowHeight(20);
    sheet.getRange('A:C').setColumnWidth(14);
    sheet.getRange('I2:I6').format.font.setColor({ rgb: '#FF0000' });
    sheet.getRange('A1:B1').format.fill.setForegroundColor({ rgb: '#FFFF00' });
    sheet.getRange('A1:B6').format.border.setValue({
      left: { style: 'thin', color: { rgb: '#000000' } },
      right: { style: 'thin', color: { rgb: '#000000' } },
      top: { style: 'thin', color: { rgb: '#000000' } },
      bottom: { style: 'thin', color: { rgb: '#000000' } },
    });

    expect(workbook.getMutations()).toMatchObject([
      {
        kind: 'range.setRowHeight',
        sheetId: '1',
        sheetName: 'Sheet1',
        address: 'A1:XFD1',
        height: 20,
      },
      {
        kind: 'range.setColumnWidth',
        sheetId: '1',
        sheetName: 'Sheet1',
        address: 'A1:C1048576',
        width: 14,
      },
      {
        kind: 'range.format.font.setColor',
        sheetId: '1',
        sheetName: 'Sheet1',
        address: 'I2:I6',
        color: { rgb: '#FF0000' },
      },
      {
        kind: 'range.format.fill.setForegroundColor',
        sheetId: '1',
        sheetName: 'Sheet1',
        address: 'A1:B1',
        color: { rgb: '#FFFF00' },
      },
      {
        kind: 'range.format.border.setValue',
        sheetId: '1',
        sheetName: 'Sheet1',
        address: 'A1:B6',
        border: {
          left: { style: 'thin', color: { rgb: '#000000' } },
          right: { style: 'thin', color: { rgb: '#000000' } },
          top: { style: 'thin', color: { rgb: '#000000' } },
          bottom: { style: 'thin', color: { rgb: '#000000' } },
        },
      },
    ]);
  });
});
