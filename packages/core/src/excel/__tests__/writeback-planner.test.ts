import { describe, expect, test } from 'vitest';
import {
  diffExcelWorkbookModels,
  planExcelWriteback,
  type ExcelWorkbookModel,
} from '../index.js';

const BASE_MODEL: ExcelWorkbookModel = {
  sharedStrings: [],
  definedNames: [],
  styles: {
    numberFormats: [],
    fonts: [],
    fills: [],
    borders: [],
    cellFormats: [],
  },
  sheets: [
    {
      id: '1',
      name: 'Sheet1',
      dimension: 'A1:A1',
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
      cells: [{ address: 'A1', row: 1, column: 1, kind: 'number', value: 1 }],
    },
  ],
};

function cloneModel(model = BASE_MODEL): ExcelWorkbookModel {
  return structuredClone(model);
}

describe('Excel writeback planner', () => {
  test('returns no-op for unchanged workbook models', () => {
    const plan = planExcelWriteback(cloneModel(), cloneModel(), {});

    expect(plan.strategy).toBe('no-op');
    expect(plan.operations).toEqual([]);
    expect(plan.diff.changed).toBe(false);
  });

  test('plans bounded cell diffs as direct patches when host supports live cell patching', () => {
    const after = cloneModel();
    after.sheets[0]!.cells[0] = { address: 'A1', row: 1, column: 1, kind: 'number', value: 2 };

    const plan = planExcelWriteback(cloneModel(), after, {
      canPatchCells: true,
      maxCellPatchCount: 10,
    });

    expect(plan.strategy).toBe('diff-patch');
    expect(plan.operations).toEqual([
      {
        type: 'patchCells',
        sheetName: 'Sheet1',
        cellCount: 1,
        changes: [
          {
            kind: 'modified',
            address: 'A1',
            before: { address: 'A1', row: 1, column: 1, kind: 'number', value: 1 },
            after: { address: 'A1', row: 1, column: 1, kind: 'number', value: 2 },
          },
        ],
      },
    ]);
  });

  test('prefers sheet shadow over diff patch in auto mode when both are available', () => {
    const after = cloneModel();
    after.sheets[0]!.cells.push({ address: 'B1', row: 1, column: 2, kind: 'number', value: 2 });

    const plan = planExcelWriteback(cloneModel(), after, {
      canPatchCells: true,
      maxCellPatchCount: 10,
      canInsertShadowSheets: true,
      canCopySheetContent: true,
      canDeleteSheet: true,
      canRenameSheet: true,
    });

    expect(plan.strategy).toBe('sheet-shadow-replace');
    expect(plan.operations.map((operation) => operation.type)).toEqual([
      'insertShadowSheet',
      'copySheetContent',
      'deleteSheet',
    ]);
    expect(plan.operations[0]).toMatchObject({
      type: 'insertShadowSheet',
      sheetId: '1',
      sheetName: '__coco_shadow_1',
      targetSheet: {
        id: '1',
        name: 'Sheet1',
        cells: [
          { address: 'A1', row: 1, column: 1, kind: 'number', value: 1 },
          { address: 'B1', row: 1, column: 2, kind: 'number', value: 2 },
        ],
      },
      targetStyles: {
        numberFormats: [],
        fonts: [],
        fills: [],
        borders: [],
        cellFormats: [],
      },
    });
  });

  test('falls back to sheet shadow replacement when cell patch exceeds host threshold', () => {
    const after = cloneModel();
    after.sheets[0]!.cells.push({ address: 'B1', row: 1, column: 2, kind: 'number', value: 2 });

    const plan = planExcelWriteback(cloneModel(), after, {
      canPatchCells: true,
      maxCellPatchCount: 0,
      canInsertShadowSheets: true,
      canCopySheetContent: true,
      canDeleteSheet: true,
      canRenameSheet: true,
    });

    expect(plan.strategy).toBe('sheet-shadow-replace');
  });

  test('prefers full package replacement when explicitly requested and supported', () => {
    const after = cloneModel();
    after.definedNames.push({ name: 'Answer', text: 'Sheet1!$A$1' });

    const plan = planExcelWriteback(
      cloneModel(),
      after,
      { canReplaceWorkbookPackage: true },
      { preference: 'package' },
    );

    expect(plan.strategy).toBe('workbook-package-replace');
    expect(plan.operations).toEqual([{ type: 'replaceWorkbookPackage' }]);
  });

  test('reports unsupported reasons when host capabilities cannot express the diff', () => {
    const after = cloneModel();
    after.sheets[0]!.tables.push({
      id: 1,
      name: 'Table1',
      displayName: 'Table1',
      ref: 'A1:B2',
      columns: [
        { id: 1, name: 'Name' },
        { id: 2, name: 'Amount' },
      ],
    });

    const plan = planExcelWriteback(cloneModel(), after, { canPatchCells: true });

    expect(plan.strategy).toBe('unsupported');
    expect(plan.unsupportedReasons.length).toBeGreaterThan(0);
  });

  test('classifies sheet adds, deletes, renames, and object changes', () => {
    const before = cloneModel();
    const after = cloneModel();
    after.sheets[0]!.name = 'Renamed';
    after.sheets.push({
      ...structuredClone(after.sheets[0]!),
      id: '2',
      name: 'Added',
      cells: [],
      comments: [{ ref: 'A1', text: 'note' }],
    });

    const diff = diffExcelWorkbookModels(before, after);

    expect(diff.changed).toBe(true);
    expect(diff.sheets.map((sheet) => sheet.kind)).toEqual(['renamed', 'added']);
    expect(diff.sheets[1]?.objectChanged).toBe(true);
    expect(diff.sheets[1]?.cells).toEqual([]);
  });

  test('emits added and deleted cell payloads', () => {
    const before = cloneModel();
    before.sheets[0]!.cells.push({ address: 'B1', row: 1, column: 2, kind: 'string', value: 'remove' });
    const after = cloneModel();
    after.sheets[0]!.cells.push({ address: 'C1', row: 1, column: 3, kind: 'string', value: 'add' });

    const diff = diffExcelWorkbookModels(before, after);

    expect(diff.sheets[0]?.cells).toEqual([
      {
        kind: 'deleted',
        address: 'B1',
        before: { address: 'B1', row: 1, column: 2, kind: 'string', value: 'remove' },
      },
      {
        kind: 'added',
        address: 'C1',
        after: { address: 'C1', row: 1, column: 3, kind: 'string', value: 'add' },
      },
    ]);
  });

  test('sheet shadow target sheet payload is a defensive clone', () => {
    const after = cloneModel();
    after.sheets[0]!.cells[0] = { address: 'A1', row: 1, column: 1, kind: 'string', value: 'planned' };

    const plan = planExcelWriteback(
      cloneModel(),
      after,
      {
        canInsertShadowSheets: true,
        canCopySheetContent: true,
        canDeleteSheet: true,
        canRenameSheet: true,
      },
      { preference: 'sheet-shadow' },
    );
    const insert = plan.operations[0];
    if (insert?.type !== 'insertShadowSheet') throw new Error('expected insertShadowSheet');

    insert.targetSheet.cells[0]!.value = 'mutated';

    expect(after.sheets[0]?.cells[0]?.value).toBe('planned');
  });
});
