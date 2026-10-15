import { describe, expect, test } from 'vitest';
import {
  createShadowWorkbookModel,
  createShadowWorkbookPayload,
  ExcelWorkbookCore,
  type ExcelSheetModel,
  type ExcelStylesModel,
} from '../index.js';

const styles: ExcelStylesModel = {
  numberFormats: [],
  fonts: [
    {
      name: 'Aptos',
      size: 11,
    },
  ],
  fills: [],
  borders: [],
  cellFormats: [
    {
      fontId: 0,
      applyFont: true,
    },
  ],
};

const sheet: ExcelSheetModel = {
  id: 'source-sheet',
  name: 'Target',
  dimension: 'A1:B2',
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
    { address: 'A1', row: 1, column: 1, kind: 'string', value: 'name', styleId: 0 },
    { address: 'B1', row: 1, column: 2, kind: 'string', value: 'amount' },
    { address: 'A2', row: 2, column: 1, kind: 'string', value: 'alpha' },
    { address: 'B2', row: 2, column: 2, kind: 'formula', value: 42, formula: '21*2' },
  ],
};

describe('Excel shadow workbook exporter', () => {
  test('creates a one-sheet workbook model with a temporary shadow sheet name', () => {
    const model = createShadowWorkbookModel(sheet, {
      sheetName: '__coco_shadow_1',
      styles,
      definedNames: [{ name: 'Answer', text: "'__coco_shadow_1'!$B$2" }],
    });

    expect(model.sheets).toHaveLength(1);
    expect(model.sheets[0]?.name).toBe('__coco_shadow_1');
    expect(model.sheets[0]?.relationshipId).toBeUndefined();
    expect(model.sheets[0]?.path).toBeUndefined();
    expect(model.styles.fonts[0]?.name).toBe('Aptos');
    expect(model.definedNames[0]?.name).toBe('Answer');
  });

  test('exports readable xlsx bytes and base64 for host adapters', async () => {
    const payload = await createShadowWorkbookPayload(sheet, {
      sheetName: '__coco_shadow_1',
      styles,
    });

    expect(payload.sheetName).toBe('__coco_shadow_1');
    expect(payload.bytes).toBeInstanceOf(Uint8Array);
    expect(payload.bytes.length).toBeGreaterThan(0);
    expect(payload.base64.length).toBeGreaterThan(0);

    const workbook = await ExcelWorkbookCore.fromZip(payload.bytes);
    const shadow = workbook.worksheets.getItem('__coco_shadow_1');
    expect(shadow.getCell('A1').value()).toBe('name');
    expect(shadow.getCell('B2').formula()).toBe('21*2');
    expect(workbook.styles().fonts[0]?.name).toBe('Aptos');
  });

  test('does not mutate the input sheet model', () => {
    const model = createShadowWorkbookModel(sheet, { sheetName: '__shadow' });
    model.sheets[0]!.cells[0]!.value = 'mutated';

    expect(sheet.cells[0]?.value).toBe('name');
  });
});
