import { describe, expect, test } from 'vitest';
import { ExcelWorkbook, type ExcelWorkbookModel } from '@coco-vexel/coco-office/excel';
import { buildSpreadsheetClipboardPayload } from '../clipboard.js';
import { buildSpreadsheetSheetRenderModel } from '../render-model.js';
import { SpreadsheetEditorRuntime } from '../runtime.js';

describe('spreadsheet clipboard', () => {
  test('preserves merge ranges for internal copy and paste', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    const payload = buildSpreadsheetClipboardPayload(runtime.activeSheet(), {
      startRow: 1,
      startColumn: 1,
      endRow: 4,
      endColumn: 4,
    });

    expect(payload.merges).toEqual([
      { startRowOffset: 1, startColumnOffset: 1, rowCount: 2, columnCount: 2 },
    ]);
    expect(payload.values.slice(1, 3).map((row) => row.slice(1, 3))).toEqual([
      ['merged', null],
      [null, null],
    ]);
    expect(payload.styleIds[1]?.slice(1, 3)).toEqual([1, 2]);

    await runtime.pasteRangeValues(6, 6, payload.values, payload.merges);

    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'B2:C3',
      'G7:H8',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('G7').values()).toEqual([['merged']]);

    await workbook.worksheets.getItem('Sheet1').getRange('G7:H8').unmerge();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('G7:H8').values()).toEqual([
      ['merged', null],
      [null, null],
    ]);
  });

  test('only captures merge ranges fully contained by the copied range', () => {
    const sheet = buildSpreadsheetSheetRenderModel(model().sheets[0]!, model().styles);

    const payload = buildSpreadsheetClipboardPayload(sheet, {
      startRow: 2,
      startColumn: 2,
      endRow: 2,
      endColumn: 2,
    });

    expect(payload.merges).toEqual([]);
  });
});

function model(): ExcelWorkbookModel {
  return {
    sharedStrings: [],
    definedNames: [],
    styles: {
      numberFormats: [],
      fonts: [{}, { bold: true }, { italic: true }],
      fills: [{}],
      borders: [{}],
      cellFormats: [{}, { fontId: 1 }, { fontId: 2 }],
    },
    sheets: [
      {
        id: '1',
        name: 'Sheet1',
        sheetViews: [],
        columns: [],
        rows: [],
        mergedRanges: [{ address: 'B2:C3' }],
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
          { address: 'A1', row: 1, column: 1, kind: 'string', value: 'a1' },
          { address: 'B2', row: 2, column: 2, kind: 'string', value: 'merged', styleId: 1 },
          { address: 'C2', row: 2, column: 3, kind: 'string', value: 'stale covered value', styleId: 2 },
          { address: 'B3', row: 3, column: 2, kind: 'string', value: 'another stale value' },
          { address: 'D4', row: 4, column: 4, kind: 'number', value: 4 },
        ],
      },
    ],
  };
}
