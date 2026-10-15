import { describe, expect, test } from 'vitest';
import {
  ExcelWorkbookCore as ExcelWorkbook,
  type ExcelSheetModel,
  type ExcelStylesModel,
} from '../index.js';

const emptyStyles: ExcelStylesModel = {
  numberFormats: [],
  fonts: [],
  fills: [],
  borders: [],
  cellFormats: [],
};

function makeSheet(): ExcelSheetModel {
  return {
    id: '1',
    name: 'Grid',
    dimension: 'A1:C3',
    sheetViews: [
      {
        workbookViewId: 0,
        pane: { topLeftCell: 'A2' },
        selections: [{ activeCell: 'B1', sqref: 'A2:C3 B1' }],
      },
    ],
    columns: [
      { min: 1, max: 3, width: 20 },
      { min: 5, max: 5, width: 10 },
    ],
    rows: [
      { index: 2, height: 20 },
      { index: 5, height: 10 },
    ],
    mergedRanges: [{ address: 'A2:C3' }, { address: 'B1:B1' }],
    hyperlinks: [
      { ref: 'A2:C3', target: 'https://example.com/range' },
      { ref: 'B1', target: 'https://example.com/single' },
    ],
    tables: [
      {
        id: 1,
        name: 'GridTable',
        displayName: 'GridTable',
        ref: 'A2:C3',
        autoFilterRef: 'A2:C3',
        columns: [],
      },
    ],
    dataValidations: [{ sqref: 'A2:C3 B1', type: 'whole', formula1: 'A2', formula2: 'C3' }],
    conditionalFormats: [
      {
        sqref: 'A2:C3 B1',
        rules: [{ type: 'cellIs', formulas: ['A2>0'] }],
      },
    ],
    comments: [{ ref: 'A2', text: 'comment' }],
    threadedComments: [{ ref: 'B1', text: 'threaded comment' }],
    drawings: [
      {
        path: 'xl/drawings/drawing1.xml',
        anchors: [
          {
            kind: 'twoCell',
            from: { row: 1, column: 1 },
            to: { row: 3, column: 3 },
            object: { kind: 'chart', relationshipId: 'rId1' },
          },
        ],
      },
    ],
    charts: [
      {
        relationshipId: 'rId1',
        drawingPath: 'xl/drawings/drawing1.xml',
        path: 'xl/charts/chart1.xml',
        formulaRefs: ['Grid!A2:C3', 'B1'],
      },
    ],
    images: [],
    cells: [
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'a' },
      { address: 'B1', row: 1, column: 2, kind: 'string', value: 'b' },
      { address: 'A2', row: 2, column: 1, kind: 'string', value: 'c' },
      { address: 'C3', row: 3, column: 3, kind: 'formula', value: null, formula: 'A2&B1+Grid!A2' },
    ],
  };
}

function makeWorkbook() {
  return ExcelWorkbook.fromModel({
    sharedStrings: [],
    definedNames: [
      { name: 'GlobalGridRef', text: 'Grid!A2:C3' },
      { name: 'LocalGridRef', text: 'A2:C3', localSheetId: 0 },
      { name: 'OtherLocalRef', text: 'Grid!A2', localSheetId: 1 },
    ],
    styles: emptyStyles,
    sheets: [makeSheet()],
  });
}

describe('ExcelWorksheet mutations', () => {
  test('inserts and deletes rows in the cell grid model', () => {
    const sheet = makeWorkbook().worksheets.getItem('Grid');

    sheet.insertRows(2);

    expect(sheet.dimension()).toBe('A1:C4');
    expect(sheet.cells().map((cell) => cell.address)).toEqual(['A1', 'B1', 'A3', 'C4']);
    expect(sheet.getRange('A1:C4').values()).toEqual([
      ['a', 'b', null],
      [null, null, null],
      ['c', null, null],
      [null, null, null],
    ]);
    expect(sheet.getCell('C4').formula()).toBe('A3&B1+Grid!A3');
    expect(sheet.mergedRanges()).toEqual([{ address: 'A3:C4' }, { address: 'B1' }]);
    expect(sheet.hyperlinks.toModel().map((hyperlink) => hyperlink.ref)).toEqual(['A3:C4', 'B1']);
    expect(sheet.rows().map((row) => row.index)).toEqual([3, 6]);
    expect(sheet.columns().map((column) => [column.min, column.max])).toEqual([
      [1, 3],
      [5, 5],
    ]);
    expect(sheet.comments.toModel().map((comment) => comment.ref)).toEqual(['A3']);
    expect(sheet.threadedComments.toModel().map((comment) => comment.ref)).toEqual(['B1']);
    expect(sheet.sheetViews()[0]?.pane?.topLeftCell).toBe('A3');
    expect(sheet.sheetViews()[0]?.selections[0]?.activeCell).toBe('B1');
    expect(sheet.sheetViews()[0]?.selections[0]?.sqref).toBe('A3:C4 B1');
    expect(sheet.tables.toModel()[0]?.ref).toBe('A3:C4');
    expect(sheet.tables.toModel()[0]?.autoFilterRef).toBe('A3:C4');
    expect(sheet.dataValidations.toModel().map((validation) => validation.sqref)).toEqual([
      'A3:C4 B1',
    ]);
    expect(sheet.dataValidations.toModel()[0]?.formula1).toBe('A3');
    expect(sheet.dataValidations.toModel()[0]?.formula2).toBe('C4');
    expect(sheet.conditionalFormats.toModel().map((format) => format.sqref)).toEqual(['A3:C4 B1']);
    expect(sheet.conditionalFormats.toModel()[0]?.rules[0]?.formulas).toEqual(['A3>0']);
    expect(sheet.charts.toModel()[0]?.formulaRefs).toEqual(['Grid!A3:C4', 'B1']);
    expect(sheet.drawings.toModel()[0]?.anchors[0]?.from).toMatchObject({ row: 2, column: 1 });
    expect(sheet.drawings.toModel()[0]?.anchors[0]?.to).toMatchObject({ row: 4, column: 3 });

    sheet.deleteRows(1, 2);

    expect(sheet.dimension()).toBe('A1:C2');
    expect(sheet.cells().map((cell) => cell.address)).toEqual(['A1', 'C2']);
    expect(sheet.getRange('A1:C2').values()).toEqual([
      ['c', null, null],
      [null, null, null],
    ]);
    expect(sheet.getCell('C2').formula()).toBe('A1&#REF!+Grid!A1');
    expect(sheet.toModel().cells.find((cell) => cell.address === 'C2')?.formula).toBe(
      'A1&#REF!+Grid!A1',
    );
    expect(sheet.mergedRanges()).toEqual([{ address: 'A1:C2' }]);
    expect(sheet.hyperlinks.toModel().map((hyperlink) => hyperlink.ref)).toEqual(['A1:C2']);
    expect(sheet.rows().map((row) => row.index)).toEqual([1, 4]);
    expect(sheet.comments.toModel().map((comment) => comment.ref)).toEqual(['A1']);
    expect(sheet.threadedComments.toModel()).toEqual([]);
    expect(sheet.sheetViews()[0]?.pane?.topLeftCell).toBe('A1');
    expect(sheet.sheetViews()[0]?.selections[0]?.activeCell).toBeUndefined();
    expect(sheet.sheetViews()[0]?.selections[0]?.sqref).toBe('A1:C2');
    expect(sheet.tables.toModel()[0]?.ref).toBe('A1:C2');
    expect(sheet.tables.toModel()[0]?.autoFilterRef).toBe('A1:C2');
    expect(sheet.dataValidations.toModel().map((validation) => validation.sqref)).toEqual([
      'A1:C2',
    ]);
    expect(sheet.dataValidations.toModel()[0]?.formula1).toBe('A1');
    expect(sheet.dataValidations.toModel()[0]?.formula2).toBe('C2');
    expect(sheet.conditionalFormats.toModel().map((format) => format.sqref)).toEqual(['A1:C2']);
    expect(sheet.conditionalFormats.toModel()[0]?.rules[0]?.formulas).toEqual(['A1>0']);
    expect(sheet.charts.toModel()[0]?.formulaRefs).toEqual(['Grid!A1:C2', '#REF!']);
    expect(sheet.drawings.toModel()[0]?.anchors[0]?.from).toMatchObject({ row: 0, column: 1 });
    expect(sheet.drawings.toModel()[0]?.anchors[0]?.to).toMatchObject({ row: 2, column: 3 });
  });

  test('inserts and deletes columns in the cell grid model', () => {
    const sheet = makeWorkbook().worksheets.getItem('Grid');

    sheet.insertColumns(2, 2);

    expect(sheet.dimension()).toBe('A1:E3');
    expect(sheet.cells().map((cell) => cell.address)).toEqual(['A1', 'D1', 'A2', 'E3']);
    expect(sheet.getRange('A1:E1').values()).toEqual([['a', null, null, 'b', null]]);
    expect(sheet.getCell('E3').formula()).toBe('A2&D1+Grid!A2');
    expect(sheet.mergedRanges()).toEqual([{ address: 'A2:E3' }, { address: 'D1' }]);
    expect(sheet.hyperlinks.toModel().map((hyperlink) => hyperlink.ref)).toEqual(['A2:E3', 'D1']);
    expect(sheet.columns().map((column) => [column.min, column.max])).toEqual([
      [1, 5],
      [7, 7],
    ]);
    expect(sheet.comments.toModel().map((comment) => comment.ref)).toEqual(['A2']);
    expect(sheet.threadedComments.toModel().map((comment) => comment.ref)).toEqual(['D1']);
    expect(sheet.sheetViews()[0]?.selections[0]?.activeCell).toBe('D1');
    expect(sheet.sheetViews()[0]?.selections[0]?.sqref).toBe('A2:E3 D1');
    expect(sheet.tables.toModel()[0]?.ref).toBe('A2:E3');
    expect(sheet.tables.toModel()[0]?.autoFilterRef).toBe('A2:E3');
    expect(sheet.dataValidations.toModel().map((validation) => validation.sqref)).toEqual([
      'A2:E3 D1',
    ]);
    expect(sheet.dataValidations.toModel()[0]?.formula1).toBe('A2');
    expect(sheet.dataValidations.toModel()[0]?.formula2).toBe('E3');
    expect(sheet.conditionalFormats.toModel().map((format) => format.sqref)).toEqual(['A2:E3 D1']);
    expect(sheet.conditionalFormats.toModel()[0]?.rules[0]?.formulas).toEqual(['A2>0']);
    expect(sheet.charts.toModel()[0]?.formulaRefs).toEqual(['Grid!A2:E3', 'D1']);
    expect(sheet.drawings.toModel()[0]?.anchors[0]?.from).toMatchObject({ row: 1, column: 3 });
    expect(sheet.drawings.toModel()[0]?.anchors[0]?.to).toMatchObject({ row: 3, column: 5 });

    sheet.deleteColumns(1);

    expect(sheet.dimension()).toBe('C1:D3');
    expect(sheet.cells().map((cell) => cell.address)).toEqual(['C1', 'D3']);
    expect(sheet.getRange('A1:D3').values()).toEqual([
      [null, null, 'b', null],
      [null, null, null, null],
      [null, null, null, null],
    ]);
    expect(sheet.getCell('D3').formula()).toBe('#REF!&C1+#REF!');
    expect(sheet.mergedRanges()).toEqual([{ address: 'A2:D3' }, { address: 'C1' }]);
    expect(sheet.hyperlinks.toModel().map((hyperlink) => hyperlink.ref)).toEqual(['A2:D3', 'C1']);
    expect(sheet.columns().map((column) => [column.min, column.max])).toEqual([
      [1, 4],
      [6, 6],
    ]);
    expect(sheet.comments.toModel()).toEqual([]);
    expect(sheet.threadedComments.toModel().map((comment) => comment.ref)).toEqual(['C1']);
    expect(sheet.sheetViews()[0]?.selections[0]?.activeCell).toBe('C1');
    expect(sheet.sheetViews()[0]?.selections[0]?.sqref).toBe('A2:D3 C1');
    expect(sheet.tables.toModel()[0]?.ref).toBe('A2:D3');
    expect(sheet.tables.toModel()[0]?.autoFilterRef).toBe('A2:D3');
    expect(sheet.dataValidations.toModel().map((validation) => validation.sqref)).toEqual([
      'A2:D3 C1',
    ]);
    expect(sheet.dataValidations.toModel()[0]?.formula1).toBe('#REF!');
    expect(sheet.dataValidations.toModel()[0]?.formula2).toBe('D3');
    expect(sheet.conditionalFormats.toModel().map((format) => format.sqref)).toEqual(['A2:D3 C1']);
    expect(sheet.conditionalFormats.toModel()[0]?.rules[0]?.formulas).toEqual(['#REF!>0']);
    expect(sheet.charts.toModel()[0]?.formulaRefs).toEqual(['Grid!A2:D3', 'C1']);
    expect(sheet.drawings.toModel()[0]?.anchors[0]?.from).toMatchObject({ row: 1, column: 2 });
    expect(sheet.drawings.toModel()[0]?.anchors[0]?.to).toMatchObject({ row: 3, column: 4 });
  });

  test('removes merged ranges and hyperlinks when their reference is fully deleted', () => {
    const sheet = makeWorkbook().worksheets.getItem('Grid');

    sheet.deleteRows(1);

    expect(sheet.mergedRanges()).toEqual([{ address: 'A1:C2' }]);
    expect(sheet.hyperlinks.toModel().map((hyperlink) => hyperlink.ref)).toEqual(['A1:C2']);
    expect(sheet.tables.toModel()[0]?.ref).toBe('A1:C2');
    expect(sheet.tables.toModel()[0]?.autoFilterRef).toBe('A1:C2');
    expect(sheet.dataValidations.toModel().map((validation) => validation.sqref)).toEqual([
      'A1:C2',
    ]);
    expect(sheet.conditionalFormats.toModel().map((format) => format.sqref)).toEqual(['A1:C2']);
  });

  test('trims merged ranges when deleted rows or columns intersect them', () => {
    const workbook = ExcelWorkbook.fromModel({
      sharedStrings: [],
      definedNames: [],
      styles: emptyStyles,
      sheets: [
        {
          ...makeSheet(),
          cells: [
            { address: 'A2', row: 2, column: 1, kind: 'string', value: 'row merge' },
            { address: 'A6', row: 6, column: 1, kind: 'string', value: 'column merge' },
          ],
          mergedRanges: [
            { address: 'A2:C5' },
            { address: 'A6:D6' },
            { address: 'F2:G3' },
          ],
          hyperlinks: [],
          tables: [],
          dataValidations: [],
          conditionalFormats: [],
          comments: [],
          threadedComments: [],
          drawings: [],
          charts: [],
        },
      ],
    });
    const sheet = workbook.worksheets.getItem('Grid');

    sheet.deleteRows(3, 2);

    expect(sheet.mergedRanges()).toEqual([
      { address: 'A2:C3' },
      { address: 'A4:D4' },
      { address: 'F2:G2' },
    ]);
    expect(sheet.getRange('A1:D4').values()).toEqual([
      [null, null, null, null],
      ['row merge', null, null, null],
      [null, null, null, null],
      ['column merge', null, null, null],
    ]);

    sheet.deleteColumns(2, 2);

    expect(sheet.mergedRanges()).toEqual([
      { address: 'A2:A3' },
      { address: 'A4:B4' },
      { address: 'D2:E2' },
    ]);
    expect(sheet.getRange('A1:B4').values()).toEqual([
      [null, null],
      ['row merge', null],
      [null, null],
      ['column merge', null],
    ]);
  });

  test('removes validations and conditional formats when every sqref segment is deleted', () => {
    const workbook = ExcelWorkbook.fromModel({
      sharedStrings: [],
      definedNames: [],
      styles: emptyStyles,
      sheets: [
        {
          ...makeSheet(),
          cells: [],
          mergedRanges: [],
          hyperlinks: [],
          dataValidations: [{ sqref: 'B1', type: 'whole' }],
          conditionalFormats: [{ sqref: 'B1', rules: [{ formulas: ['B1>0'] }] }],
        },
      ],
    });
    const sheet = workbook.worksheets.getItem('Grid');

    sheet.deleteColumns(2);

    expect(sheet.dataValidations.toModel()).toEqual([]);
    expect(sheet.conditionalFormats.toModel()).toEqual([]);
  });

  test('renames explicit sheet references in local formulas', () => {
    const workbook = makeWorkbook();
    const sheet = workbook.worksheets.getItem('Grid');

    sheet.rename('Renamed');

    expect(sheet.name()).toBe('Renamed');
    expect(sheet.getCell('C3').formula()).toBe('A2&B1+Renamed!A2');
    expect(workbook.names.getItem('GlobalGridRef').text()).toBe('Renamed!A2:C3');
    expect(workbook.names.getItem('LocalGridRef').text()).toBe('A2:C3');
    expect(workbook.names.getItem('OtherLocalRef').text()).toBe('Grid!A2');
  });

  test('renames explicit sheet references across workbook sheets', () => {
    const workbook = ExcelWorkbook.fromModel({
      sharedStrings: [],
      definedNames: [],
      styles: emptyStyles,
      sheets: [
        makeSheet(),
        {
          ...makeSheet(),
          id: '2',
          name: 'Refs',
          dimension: 'A1:A1',
          sheetViews: [],
          columns: [],
          rows: [],
          mergedRanges: [],
          hyperlinks: [],
          tables: [],
          dataValidations: [{ sqref: 'B1', type: 'whole', formula1: 'Grid!A2' }],
          conditionalFormats: [
            {
              sqref: 'C1',
              rules: [{ type: 'cellIs', formulas: ['Grid!A2>0'] }],
            },
          ],
          comments: [],
          threadedComments: [],
          drawings: [],
          charts: [
            {
              relationshipId: 'rId2',
              drawingPath: 'xl/drawings/drawing2.xml',
              path: 'xl/charts/chart2.xml',
              formulaRefs: ['Grid!A2:C3'],
            },
          ],
          images: [],
          cells: [
            {
              address: 'A1',
              row: 1,
              column: 1,
              kind: 'number',
              value: 1,
              formula: 'Grid!A2',
            },
          ],
        },
      ],
    });

    workbook.worksheets.getItem('Grid').rename('Renamed');
    const refs = workbook.worksheets.getItem('Refs');

    expect(refs.getCell('A1').formula()).toBe('Renamed!A2');
    expect(refs.dataValidations.toModel()[0]?.formula1).toBe('Renamed!A2');
    expect(refs.conditionalFormats.toModel()[0]?.rules[0]?.formulas).toEqual(['Renamed!A2>0']);
    expect(refs.charts.toModel()[0]?.formulaRefs).toEqual(['Renamed!A2:C3']);
  });

  test('supports worksheet lifecycle and defined name mutations', async () => {
    const workbook = makeWorkbook();

    const added = workbook.worksheets.add('Added');
    added.getCell('A1').setValue('added');
    added.setVisibility('hidden');
    const copied = workbook.worksheets.copy('Grid', 'Grid Copy');
    workbook.worksheets.move('Added', 0);
    workbook.worksheets.deleteItem('Grid Copy');

    const name = workbook.names.add({ name: 'AddedRef', text: 'Added!$A$1' });
    name.setText('Added!$A$1:$A$1');
    name.rename('AddedRange');
    name.setLocalSheetId(0);

    expect(workbook.worksheets.items().map((sheet) => sheet.name())).toEqual(['Added', 'Grid']);
    expect(added.visibility()).toBe('hidden');
    expect(copied.name()).toBe('Grid Copy');
    expect(workbook.names.getItem('AddedRange').toModel()).toEqual({
      name: 'AddedRange',
      text: 'Added!$A$1:$A$1',
      localSheetId: 0,
    });

    workbook.names.deleteItem('AddedRange');
    expect(workbook.names.getItemOrNullObject('AddedRange')).toBeNull();

    const roundtrip = await ExcelWorkbook.fromZip(await workbook.toZip());
    expect(roundtrip.worksheets.items().map((sheet) => sheet.name())).toEqual(['Added', 'Grid']);
    expect(roundtrip.worksheets.getItem('Added').visibility()).toBe('hidden');
    expect(roundtrip.worksheets.getItem('Added').getCell('A1').value()).toBe('added');
  });

  test('rewrites defined names for local row and column mutations', () => {
    const workbook = makeWorkbook();
    const sheet = workbook.worksheets.getItem('Grid');

    sheet.insertRows(2);
    expect(workbook.names.getItem('GlobalGridRef').text()).toBe('Grid!A3:C4');
    expect(workbook.names.getItem('LocalGridRef').text()).toBe('A3:C4');
    expect(workbook.names.getItem('OtherLocalRef').text()).toBe('Grid!A2');

    sheet.deleteColumns(1);
    expect(workbook.names.getItem('GlobalGridRef').text()).toBe('Grid!A3:B4');
    expect(workbook.names.getItem('LocalGridRef').text()).toBe('A3:B4');
    expect(workbook.names.getItem('OtherLocalRef').text()).toBe('Grid!A2');
  });

  test('rejects invalid row and column mutation arguments', () => {
    const sheet = makeWorkbook().worksheets.getItem('Grid');

    expect(() => sheet.insertRows(0)).toThrow(/rowIndex/);
    expect(() => sheet.deleteRows(1, 0)).toThrow(/count/);
    expect(() => sheet.insertColumns(0)).toThrow(/columnIndex/);
    expect(() => sheet.deleteColumns(1, 0)).toThrow(/count/);
  });
});
