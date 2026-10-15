import { describe, expect, test } from 'vitest';
import { ExcelWorkbook, summarizeExcelTransactionGroup, type ExcelWorkbookModel } from '../index.js';

describe('Excel workbook history', () => {
  test('undoes and redoes a grouped user edit', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });

    await workbook.history().runGroup({ source: 'user', label: 'Edit sheet' }, async (tx) => {
      await tx.setCellValue('Sheet1', 'A1', 'changed');
      await tx.setCellValue('Sheet1', 'B1', 42);
    });

    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['changed', 42],
    ]);
    expect(workbook.history().canUndo()).toBe(true);

    await workbook.history().undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['before', null],
    ]);
    expect(workbook.history().canRedo()).toBe(true);

    await workbook.history().redo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['changed', 42],
    ]);
  });

  test('restores range formatting with undo', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });

    await workbook.history().runGroup({ source: 'user', label: 'Format header' }, async (tx) => {
      await tx.setRangeBold('Sheet1', 'A1:B1', true);
      await tx.setRangeFillColor('Sheet1', 'A1:B1', { rgb: '#FF0000' });
    });

    const after = await workbook.toModel();
    expect(after.sheets[0]!.cells.find((cell) => cell.address === 'A1')?.styleId).toBeGreaterThan(0);

    await workbook.history().undo();
    const undone = await workbook.toModel();
    expect(undone.sheets[0]!.cells.find((cell) => cell.address === 'A1')?.styleId).toBeUndefined();
    expect(undone.sheets[0]!.cells.find((cell) => cell.address === 'B1')?.styleId).toBeUndefined();
  });

  test('undoes and redoes merged ranges', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });

    await workbook.history().runGroup({ source: 'user', label: 'Merge title' }, async (tx) => {
      await tx.mergeRange('Sheet1', 'A1:B1');
    });

    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B1',
    ]);

    await workbook.history().undo();
    expect(await workbook.worksheets.getItem('Sheet1').mergedRanges()).toEqual([]);

    await workbook.history().redo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B1',
    ]);

    await workbook.history().runGroup({ source: 'user', label: 'Unmerge title' }, async (tx) => {
      await tx.unmergeRange('Sheet1', 'A1:B1');
    });
    expect(await workbook.worksheets.getItem('Sheet1').mergedRanges()).toEqual([]);

    await workbook.history().undo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B1',
    ]);
  });

  test('restores covered cell values when undoing a merge', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });

    await workbook.history().runGroup({ source: 'user', label: 'Prepare merge' }, async (tx) => {
      await tx.setCellValue('Sheet1', 'B1', 'covered');
      await tx.setCellValue('Sheet1', 'A2', 2);
      await tx.setCellFormula('Sheet1', 'B2', '=A2*2');
    });
    workbook.history().clear();

    await workbook.history().runGroup({ source: 'user', label: 'Merge block' }, async (tx) => {
      await tx.mergeRange('Sheet1', 'A1:B2');
    });

    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B2').values()).toEqual([
      ['before', null],
      [null, null],
    ]);

    await workbook.history().undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B2').values()).toEqual([
      ['before', 'covered'],
      [2, null],
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('B2').formulas()).toEqual([['A2*2']]);
    expect(await workbook.worksheets.getItem('Sheet1').mergedRanges()).toEqual([]);

    await workbook.history().redo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B2').values()).toEqual([
      ['before', null],
      [null, null],
    ]);
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B2',
    ]);
  });

  test('uses structural operation patches for row and column insertions', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    await workbook.worksheets.getItem('Sheet1').getRange('A2:B2').setValues([['left', 'right']]);

    await workbook.history().runGroup({ source: 'user', label: 'Insert row and column' }, async (tx) => {
      await tx.insertRows('Sheet1', 2, 1);
      await tx.insertColumns('Sheet1', 2, 1);
    });

    expect(workbook.history().undoGroups()[0]?.transactions.flatMap((tx) => tx.forward.map((patch) => patch.kind))).toEqual([
      'sheet.rows.insert',
      'sheet.columns.insert',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A3:C3').values()).toEqual([
      ['left', null, 'right'],
    ]);

    await workbook.history().undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A2:B2').values()).toEqual([
      ['left', 'right'],
    ]);

    await workbook.history().redo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A3:C3').values()).toEqual([
      ['left', null, 'right'],
    ]);
  });

  test('restores deleted rows and columns from structural patches', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    await workbook.worksheets.getItem('Sheet1').getRange('A1:C3').setValues([
      ['a1', 'b1', 'c1'],
      ['a2', 'b2', 'c2'],
      ['a3', 'b3', 'c3'],
    ]);
    workbook.history().clear();

    await workbook.history().runGroup({ source: 'user', label: 'Delete row and column' }, async (tx) => {
      await tx.deleteRows('Sheet1', 2, 1);
      await tx.deleteColumns('Sheet1', 2, 1);
    });

    expect(workbook.history().undoGroups()[0]?.transactions.flatMap((tx) => tx.forward.map((patch) => patch.kind))).toEqual([
      'sheet.rows.delete',
      'sheet.columns.delete',
    ]);
    expect(workbook.history().undoGroups()[0]?.transactions.flatMap((tx) => tx.inverse.map((patch) => patch.kind))).toEqual([
      'sheet.rows.restoreDelta',
      'sheet.columns.restoreDelta',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B2').values()).toEqual([
      ['a1', 'c1'],
      ['a3', 'c3'],
    ]);

    await workbook.history().undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:C3').values()).toEqual([
      ['a1', 'b1', 'c1'],
      ['a2', 'b2', 'c2'],
      ['a3', 'b3', 'c3'],
    ]);

    await workbook.history().redo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B2').values()).toEqual([
      ['a1', 'c1'],
      ['a3', 'c3'],
    ]);
  });

  test('restores merged ranges when deleting rows and columns with structural delta patches', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    await workbook.history().runGroup({ source: 'user', label: 'Prepare merged block' }, async (tx) => {
      await tx.setRangeValues('Sheet1', 'A1:D4', [
        ['a1', 'b1', 'c1', 'd1'],
        ['a2', 'b2', 'c2', 'd2'],
        ['a3', 'b3', 'c3', 'd3'],
        ['a4', 'b4', 'c4', 'd4'],
      ]);
      await tx.mergeRange('Sheet1', 'A1:C3');
      await tx.mergeRange('Sheet1', 'D1:D4');
    });
    workbook.history().clear();

    await workbook.history().runGroup({ source: 'user', label: 'Delete through merges' }, async (tx) => {
      await tx.deleteRows('Sheet1', 2, 1);
      await tx.deleteColumns('Sheet1', 2, 1);
    });

    expect(workbook.history().undoGroups()[0]?.transactions.flatMap((tx) => tx.inverse.map((patch) => patch.kind))).toEqual([
      'sheet.rows.restoreDelta',
      'sheet.columns.restoreDelta',
    ]);
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B2',
      'C1:C3',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:C3').values()).toEqual([
      ['a1', null, 'd1'],
      [null, null, null],
      ['a4', 'c4', null],
    ]);

    await workbook.history().undo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:C3',
      'D1:D4',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:D4').values()).toEqual([
      ['a1', null, null, 'd1'],
      [null, null, null, null],
      [null, null, null, null],
      ['a4', 'b4', 'c4', null],
    ]);

    await workbook.history().redo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B2',
      'C1:C3',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:C3').values()).toEqual([
      ['a1', null, 'd1'],
      [null, null, null],
      ['a4', 'c4', null],
    ]);
  });

  test('falls back to full row restore when formulas need reference recovery', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    await workbook.history().runGroup({ source: 'user', label: 'Prepare formula' }, async (tx) => {
      await tx.setCellValue('Sheet1', 'A2', 2);
      await tx.setCellFormula('Sheet1', 'C1', '=A1');
    });
    workbook.history().clear();

    await workbook.history().runGroup({ source: 'user', label: 'Delete formula row' }, async (tx) => {
      await tx.deleteRows('Sheet1', 2, 1);
    });

    expect(workbook.history().undoGroups()[0]?.transactions.flatMap((tx) => tx.forward.map((patch) => patch.kind))).toEqual([
      'sheet.rows.delete',
    ]);
    expect(workbook.history().undoGroups()[0]?.transactions.flatMap((tx) => tx.inverse.map((patch) => patch.kind))).toEqual([
      'sheet.rows.restore',
    ]);
  });

  test('falls back to full column restore when formulas need reference recovery', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    await workbook.history().runGroup({ source: 'user', label: 'Prepare formula' }, async (tx) => {
      await tx.setCellValue('Sheet1', 'B2', 2);
      await tx.setCellFormula('Sheet1', 'C1', '=A1');
    });
    workbook.history().clear();

    await workbook.history().runGroup({ source: 'user', label: 'Delete formula column' }, async (tx) => {
      await tx.deleteColumns('Sheet1', 2, 1);
    });

    expect(workbook.history().undoGroups()[0]?.transactions.flatMap((tx) => tx.forward.map((patch) => patch.kind))).toEqual([
      'sheet.columns.delete',
    ]);
    expect(workbook.history().undoGroups()[0]?.transactions.flatMap((tx) => tx.inverse.map((patch) => patch.kind))).toEqual([
      'sheet.columns.restore',
    ]);
  });

  test('records agent metadata and reverts an agent group as one unit', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });

    const result = await workbook.history().runAgentGroup(
      {
        label: 'Generate summary',
        actor: 'spreadsheet-agent',
        description: 'Writes a summary row',
        agentRunId: 'run-1',
        agentStepId: 'step-1',
        metadata: {
          custom: { nested: true },
          skipped: undefined,
        },
      },
      async (tx) => {
        await tx.setCellValue('Sheet1', 'A1', 'agent-title');
        await tx.setCellValue('Sheet1', 'B1', 99);
      },
    );
    expect(result?.summary).toMatchObject({
      source: 'agent',
      label: 'Generate summary',
      transactionCount: 2,
      patchCount: 4,
      forwardPatchKinds: ['range.cells'],
      inversePatchKinds: ['range.cells'],
      affectedSheets: ['Sheet1'],
      affectedRanges: ['Sheet1!A1', 'Sheet1!B1'],
      containsSnapshot: false,
      containsStructuralPatch: false,
      restoreStrategies: [],
      metadata: {
        actor: 'spreadsheet-agent',
        agentRunId: 'run-1',
        agentStepId: 'step-1',
      },
    });

    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['agent-title', 99],
    ]);

    const [group] = workbook.history().undoGroups();
    expect(group).toMatchObject({
      source: 'agent',
      label: 'Generate summary',
      metadata: {
        actor: 'spreadsheet-agent',
        description: 'Writes a summary row',
        agentRunId: 'run-1',
        agentStepId: 'step-1',
        custom: { nested: true },
      },
    });
    expect(group?.metadata).not.toHaveProperty('skipped');

    (group!.metadata!.custom as { nested: boolean }).nested = false;
    expect((workbook.history().undoGroups()[0]!.metadata!.custom as { nested: boolean }).nested).toBe(true);

    await workbook.history().undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['before', null],
    ]);
  });

  test('summarizes structural restore strategies', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    await workbook.worksheets.getItem('Sheet1').getRange('A1:C3').setValues([
      ['a1', 'b1', 'c1'],
      ['a2', 'b2', 'c2'],
      ['a3', 'b3', 'c3'],
    ]);
    workbook.history().clear();

    const result = await workbook.history().runGroup({ source: 'agent', label: 'Delete structure' }, async (tx) => {
      await tx.deleteRows('Sheet1', 2, 1);
      await tx.deleteColumns('Sheet1', 2, 1);
    });

    expect(result?.summary).toMatchObject({
      source: 'agent',
      label: 'Delete structure',
      transactionCount: 2,
      forwardPatchKinds: ['sheet.rows.delete', 'sheet.columns.delete'],
      inversePatchKinds: ['sheet.rows.restoreDelta', 'sheet.columns.restoreDelta'],
      containsSnapshot: false,
      containsStructuralPatch: true,
      restoreStrategies: ['delta'],
    });
    expect(result?.summary.transactions.map((transaction) => transaction.label)).toEqual([
      'Delete rows',
      'Delete columns',
    ]);

    const [group] = workbook.history().undoGroups();
    expect(group ? summarizeExcelTransactionGroup(group) : null).toMatchObject(result?.summary ?? {});
  });

  test('clears redo stack after a new transaction', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });

    await workbook.history().runGroup({ source: 'user', label: 'First edit' }, async (tx) => {
      await tx.setCellValue('Sheet1', 'A1', 'first');
    });
    await workbook.history().undo();
    expect(workbook.history().canRedo()).toBe(true);

    await workbook.history().runGroup({ source: 'user', label: 'Second edit' }, async (tx) => {
      await tx.setCellValue('Sheet1', 'A1', 'second');
    });

    expect(workbook.history().canRedo()).toBe(false);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([
      ['second'],
    ]);
  });

  test('merges adjacent transactions with the same merge key inside the window', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });

    await workbook.history().runGroup(
      { source: 'user', label: 'Type A1', mergeKey: 'edit:Sheet1!A1', mergeWindowMs: 1000 },
      async (tx) => {
        await tx.setCellValue('Sheet1', 'A1', 'a');
      },
    );
    await workbook.history().runGroup(
      { source: 'user', label: 'Type A1', mergeKey: 'edit:Sheet1!A1', mergeWindowMs: 1000 },
      async (tx) => {
        await tx.setCellValue('Sheet1', 'A1', 'ab');
      },
    );
    await workbook.history().runGroup(
      { source: 'user', label: 'Type A1', mergeKey: 'edit:Sheet1!A1', mergeWindowMs: 1000 },
      async (tx) => {
        await tx.setCellValue('Sheet1', 'A1', 'abc');
      },
    );

    const groups = workbook.history().undoGroups();
    expect(groups).toHaveLength(1);
    expect(groups[0]?.transactions).toHaveLength(3);
    expect(groups[0]?.mergeKey).toBe('edit:Sheet1!A1');
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([
      ['abc'],
    ]);

    await workbook.history().undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([
      ['before'],
    ]);

    await workbook.history().redo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([
      ['abc'],
    ]);
  });

  test('does not merge transactions with different merge keys or expired windows', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });

    await workbook.history().runGroup(
      { source: 'user', label: 'Type A1', mergeKey: 'edit:Sheet1!A1', mergeWindowMs: 1000 },
      async (tx) => {
        await tx.setCellValue('Sheet1', 'A1', 'a');
      },
    );
    await workbook.history().runGroup(
      { source: 'user', label: 'Type B1', mergeKey: 'edit:Sheet1!B1', mergeWindowMs: 1000 },
      async (tx) => {
        await tx.setCellValue('Sheet1', 'B1', 'b');
      },
    );
    await workbook.history().runGroup(
      { source: 'user', label: 'Type B1 later', mergeKey: 'edit:Sheet1!B1', mergeWindowMs: 0 },
      async (tx) => {
        await tx.setCellValue('Sheet1', 'B1', 'bb');
      },
    );

    expect(workbook.history().undoGroups()).toHaveLength(3);
  });
});

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
        cells: [
          { address: 'A1', row: 1, column: 1, kind: 'string', value: 'before' },
        ],
      },
    ],
  };
}
