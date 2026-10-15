import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { ExcelWorkbook } from '@coco-vexel/coco-office/excel';
import { describe, expect, test } from 'vitest';
import { SpreadsheetEditorRuntime } from '../runtime.js';
import { SpreadsheetLayout } from '../layout.js';

const fixturePath = fileURLToPath(
  new URL('../../../../core/test-fixtures/excel/microsoft-rich-objects.xlsx', import.meta.url),
);

describe('spreadsheet xlsx fixture regression', () => {
  test('keeps source row metadata and renders Excel points as CSS pixels', async () => {
    const workbook = await ExcelWorkbook.fromZip(await readFile(fixturePath), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    const sheet = runtime.activeSheet();
    const layout = new SpreadsheetLayout(sheet);

    expect(sheet.rowMetaByIndex.get(11)?.height).toBe(15.75);
    expect(layout.getRowHeight(11)).toBeCloseTo(21);
    expect(sheet.cellsByAddress.get('A1')?.displayText).toBe('__COCO_EXCEL_UNSAVED_1785722092803__');
  });

  test('preserves fixture structure through xlsx roundtrip', async () => {
    const workbook = await ExcelWorkbook.fromZip(await readFile(fixturePath), { mode: 'inline' });
    const before = await workbook.toModel();
    const roundtrip = await ExcelWorkbook.fromZip(await workbook.toZip(), { mode: 'inline' });
    const after = await roundtrip.toModel();

    expect(after.sheets[0]?.rows.find((row) => row.index === 11)?.height).toBe(
      before.sheets[0]?.rows.find((row) => row.index === 11)?.height,
    );
    expect(after.sheets[0]?.mergedRanges).toEqual(before.sheets[0]?.mergedRanges);
    expect(after.sheets[0]?.cells.filter((cell) => cell.formula).length).toBe(
      before.sheets[0]?.cells.filter((cell) => cell.formula).length,
    );
  });
});
