import { describe, expect, test } from 'vitest';
import { ExcelWorkbook, type ExcelWorkbookModel } from '@coco-vexel/coco-office/excel';
import { buildSpreadsheetClipboardPayload, clipboardHtml, parseTsv, readSpreadsheetClipboardPayload, serializeRangeToTsv, serializeSpreadsheetClipboardPayload, SPREADSHEET_CLIPBOARD_MIME } from '../clipboard.js';
import { buildSpreadsheetSheetRenderModel } from '../render-model.js';
import { SpreadsheetEditorRuntime } from '../runtime.js';

describe('spreadsheet clipboard', () => {
  test.each(['custom', 'html'])('round-trips a self-contained %s payload without an editor cache', (kind) => {
    const input = model();
    input.sheets[0]!.cells[0]!.value = '田\t字\n"帖"';
    input.sheets[0]!.cells.push({ address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: "'Another '' Sheet'!A1" });
    const payload = buildSpreadsheetClipboardPayload(buildSpreadsheetSheetRenderModel(input.sheets[0]!, input.styles), { startRow: 1, endRow: 4, startColumn: 1, endColumn: 4 }, { sourceWorkbookId: 'source' });
    const data = { getData: (type: string) => kind === 'custom' && type === SPREADSHEET_CLIPBOARD_MIME ? serializeSpreadsheetClipboardPayload(payload) : kind === 'html' && type === 'text/html' ? clipboardHtml(payload) : '' };
    expect(readSpreadsheetClipboardPayload(data)).toEqual(payload);
  });

  test.each(['schema', 'style', 'matrix', 'merge'])('rejects invalid portable %s metadata', (kind) => {
    const input = model();
    const payload = buildSpreadsheetClipboardPayload(buildSpreadsheetSheetRenderModel(input.sheets[0]!, input.styles), { startRow: 1, endRow: 4, startColumn: 1, endColumn: 4 });
    const invalid = JSON.parse(serializeSpreadsheetClipboardPayload(payload));
    if (kind === 'schema') invalid.schema = 2;
    if (kind === 'style') invalid.styles = {};
    if (kind === 'matrix') invalid.values.pop();
    if (kind === 'merge') invalid.merges.push(invalid.merges[0]);
    expect(() => readSpreadsheetClipboardPayload({ getData: (type) => type === SPREADSHEET_CLIPBOARD_MIME ? JSON.stringify(invalid) : '' })).toThrow(/格式无效/);
  });

  test('parses Excel quoted TSV without losing tabs, newlines, quotes, or blank fields', () => {
    expect(parseTsv('"a\tb"\t"first\r\nsecond"\t"say ""hi"""\t\r\n')).toEqual([['a\tb', 'first\r\nsecond', 'say "hi"', null]]);
    expect(parseTsv('')).toEqual([[null]]);
    expect(parseTsv('false\tTRUE\t0')).toEqual([[false, true, 0]]);
    expect(() => parseTsv('"unterminated')).toThrow(/引号/);
  });

  test('serializes embedded cell separators as quoted TSV instead of replacing content', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    try {
      await runtime.setRangeValues(1, 1, [['a\tb', 'a\nb', 'a"b']]);
      const range = { startRow: 1, startColumn: 1, endRow: 1, endColumn: 3 };
      expect(parseTsv(serializeRangeToTsv(runtime.activeSheet(), range))).toEqual([['a\tb', 'a\nb', 'a"b']]);
    } finally { await runtime.dispose(); }
  });
  test.each([
    { name: 'single blank', values: [[null]] },
    { name: 'trailing blank row', values: [['田'], [null]] },
    { name: 'all blank rows', values: [[null], [null], [null]] },
    { name: 'empty string and falsy values', values: [[''], [' '], [0], [false], [null]] },
  ])('preserves the shape and raw values of an internal $name copy', async ({ values }) => {
    const source = model();
    source.sheets[0]!.mergedRanges = [];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    try {
      await runtime.setRangeValues(1, 1, values);
      const payload = buildSpreadsheetClipboardPayload(runtime.activeSheet(), {
        startRow: 1, startColumn: 1, endRow: values.length, endColumn: 1,
      });
      expect(payload.values).toEqual(values);
    } finally { await runtime.dispose(); }
  });

  test('keeps whitespace, zero, false and explicit blank columns in external text', () => {
    expect(parseTsv(' \t0\tFALSE\t\r\n')).toEqual([[' ', 0, false, null]]);
  });

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
