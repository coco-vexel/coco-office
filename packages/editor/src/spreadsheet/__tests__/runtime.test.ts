import { describe, expect, test } from 'vitest';
import { ExcelWorkbook, type ExcelBorderModel, type ExcelWorkbookModel } from '@coco-vexel/coco-office/excel';
import { buildSpreadsheetClipboardPayload } from '../clipboard.js';
import { SpreadsheetEditorRuntime } from '../runtime.js';
import { buildSpreadsheetSheetRenderModel } from '../render-model.js';
import { resolveSpreadsheetRangeBorder } from '../style-resolver.js';

describe('spreadsheet runtime', () => {
  test('moves chart anchors through the user history stack', async () => {
    const source = model();
    const sheet = source.sheets[0]!;
    sheet.drawings = [{
      path: 'xl/drawings/drawing1.xml',
      anchors: [{
        kind: 'twoCell',
        from: { row: 1, column: 1 },
        to: { row: 8, column: 8 },
        object: { kind: 'chart', relationshipId: 'rId1' },
      }],
    }];
    sheet.charts = [{
      drawingPath: 'xl/drawings/drawing1.xml',
      relationshipId: 'rId1',
      type: 'columnClustered',
      formulaRefs: [],
    }];
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(await ExcelWorkbook.fromModel(source, { mode: 'inline' }));

    await runtime.moveChart(0, {
      from: { row: 3, column: 2, rowOffset: 10, columnOffset: 20 },
      to: { row: 10, column: 9, rowOffset: 30, columnOffset: 40 },
    });

    expect(runtime.model().sheets[0]!.drawings[0]!.anchors[0]).toMatchObject({
      from: { row: 3, column: 2, rowOffset: 10, columnOffset: 20 },
      to: { row: 10, column: 9, rowOffset: 30, columnOffset: 40 },
    });
    expect(runtime.canUndo()).toBe(true);
    await runtime.undo();
    expect(runtime.model().sheets[0]!.drawings[0]!.anchors[0]!.from).toMatchObject({ row: 1, column: 1 });
    await runtime.redo();
    expect(runtime.model().sheets[0]!.drawings[0]!.anchors[0]!.from).toMatchObject({ row: 3, column: 2 });
  });

  test('previews an agent transaction without mutating the live runtime', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    const runner = async (target: SpreadsheetEditorRuntime) => {
      await target.setRangeValues(1, 1, [['preview', 'only']]);
      await target.setRangeFillColor(1, 1, 1, 2, { rgb: '#FF0000' });
    };

    const preview = await runtime.previewAgentTransaction(
      {
        label: 'Preview agent edit',
        actor: 'spreadsheet-agent',
        agentRunId: 'preview-run',
      },
      runner,
    );

    expect(preview.summary).toMatchObject({
      source: 'agent',
      label: 'Preview agent edit',
      transactionCount: 2,
      forwardPatchKinds: ['range.cells', 'range.format'],
      affectedRanges: ['Sheet1!A1:B1'],
      containsSnapshot: false,
    });
    expect(preview.diffSummary).toEqual({
      changed: true,
      changedSheets: 1,
      sheetNames: ['Sheet1'],
      cellChanges: {
        total: 2,
        added: 2,
        modified: 0,
        deleted: 0,
      },
      structuralChanged: true,
      objectChanged: false,
      sampleCells: [
        { sheetName: 'Sheet1', address: 'A1', kind: 'added', before: null, after: 'preview' },
        { sheetName: 'Sheet1', address: 'B1', kind: 'added', before: null, after: 'only' },
      ],
    });
    expect(preview.afterModel.sheets[0]?.cells.find((cell) => cell.address === 'A1')?.value).toBe('preview');
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      [null, null],
    ]);
    expect(workbook.history().canUndo()).toBe(false);
    expect(runtime.model().sheets[0]?.cells).toEqual([]);

    await runtime.runAgentTransaction(
      {
        label: 'Preview agent edit',
        actor: 'spreadsheet-agent',
        agentRunId: 'preview-run',
      },
      runner,
    );

    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['preview', 'only'],
    ]);
    expect(workbook.history().undoGroups()).toHaveLength(1);
    expect(workbook.history().undoGroups()[0]?.source).toBe('agent');
  });

  test('groups multiple runtime operations into one agent transaction', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    const result = await runtime.runAgentTransaction(
      {
        label: 'Agent writes summary',
        actor: 'spreadsheet-agent',
        agentRunId: 'run-1',
        agentStepId: 'step-1',
      },
      async () => {
        await runtime.setRangeValues(1, 1, [['agent', 'summary']]);
        await runtime.setRangeBold(1, 1, 1, 2, true);
        await runtime.mergeRange(1, 1, 1, 2);
      },
    );

    expect(result?.group).toMatchObject({
      source: 'agent',
      label: 'Agent writes summary',
      metadata: {
        actor: 'spreadsheet-agent',
        agentRunId: 'run-1',
        agentStepId: 'step-1',
      },
    });
    expect(result?.group.transactions).toHaveLength(3);
    expect(workbook.history().undoGroups()).toHaveLength(1);
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B1',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['agent', null],
    ]);

    await runtime.undo();
    expect(workbook.history().canUndo()).toBe(false);
    expect(workbook.history().canRedo()).toBe(true);
    expect(await workbook.worksheets.getItem('Sheet1').mergedRanges()).toEqual([]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      [null, null],
    ]);

    await runtime.redo();
    expect(workbook.history().undoGroups()).toHaveLength(1);
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B1',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['agent', null],
    ]);
  });

  test('groups multiple user operations into one undo group', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    const result = await runtime.runUserTransaction(
      { label: 'Format summary' },
      async () => {
        await runtime.setRangeValues(1, 1, [['summary']]);
        await runtime.setRangeBold(1, 1, 1, 1, true);
        await runtime.setRangeFillColor(1, 1, 1, 1, { rgb: '#FFF2CC' });
      },
    );

    expect(result?.group).toMatchObject({ source: 'user', label: 'Format summary' });
    expect(result?.group.transactions).toHaveLength(3);
    expect(workbook.history().undoGroups()).toHaveLength(1);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([[null]]);
    expect(workbook.history().canRedo()).toBe(true);
  });

  test('replaces values in one undo group', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setRangeValues(1, 1, [['Alpha', 'Alpha']]);
    const replaced = await runtime.replaceAll('alpha', 'Beta');

    expect(replaced).toBe(2);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['Beta', 'Beta'],
    ]);
    expect(workbook.history().undoGroups()).toHaveLength(2);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['Alpha', 'Alpha'],
    ]);
  });

  test('filters rows through a workbook snapshot and restores them with undo', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    await runtime.setRangeValues(1, 1, [['Category'], ['Keep'], ['Also keep'], ['Hide']]);

    await runtime.filterRows(1, 4, 1, ['Keep', 'Also keep']);
    expect(runtime.activeSheet().rowMetaByIndex.get(2)?.hidden).toBeUndefined();
    expect(runtime.activeSheet().rowMetaByIndex.get(3)?.hidden).toBeUndefined();
    expect(runtime.activeSheet().rowMetaByIndex.get(4)?.hidden).toBe(true);

    await runtime.undo();
    expect(runtime.activeSheet().rowMetaByIndex.get(4)?.hidden).toBeUndefined();
  });

  test('keeps multiple column filters together', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    await runtime.setRangeValues(1, 1, [
      ['Category', 'Region'],
      ['Keep', 'East'],
      ['Keep', 'West'],
      ['Hide', 'East'],
    ]);

    await runtime.filterRowsByColumns(1, 4, [
      { column: 1, values: ['Keep'] },
      { column: 2, values: ['East'] },
    ]);
    expect(runtime.activeSheet().rowMetaByIndex.get(2)?.hidden).toBeUndefined();
    expect(runtime.activeSheet().rowMetaByIndex.get(3)?.hidden).toBe(true);
    expect(runtime.activeSheet().rowMetaByIndex.get(4)?.hidden).toBe(true);

    await runtime.filterRowsByColumns(1, 4, [{ column: 1, values: ['Keep'] }]);
    expect(runtime.activeSheet().rowMetaByIndex.get(2)?.hidden).toBeUndefined();
    expect(runtime.activeSheet().rowMetaByIndex.get(3)?.hidden).toBeUndefined();
    expect(runtime.activeSheet().rowMetaByIndex.get(4)?.hidden).toBe(true);

    await runtime.undo();
    expect(runtime.activeSheet().rowMetaByIndex.get(3)?.hidden).toBe(true);
  });

  test('sorts complete data rows through one workbook snapshot', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    await runtime.setRangeValues(1, 1, [
      ['Category', 'Amount'],
      ['B', 20],
      ['A', 10],
      ['C', 30],
    ]);

    await runtime.sortRows(1, 4, 1, 'asc');
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A2:B4').values()).toEqual([
      ['A', 10],
      ['B', 20],
      ['C', 30],
    ]);
    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A2:B4').values()).toEqual([
      ['B', 20],
      ['A', 10],
      ['C', 30],
    ]);
  });

  test('keeps blank sort values after populated rows', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    await runtime.setRangeValues(1, 1, [
      ['Category'],
      ['B'],
      [''],
      ['A'],
    ]);

    await runtime.sortRows(1, 100, 1, 'asc');
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A2:A4').values()).toEqual([
      ['A'],
      ['B'],
      [''],
    ]);
  });

  test('switches active sheets without leaking edits or formula results between sheets', async () => {
    const source = model();
    source.sheets.push({
      ...model().sheets[0]!,
      id: '2',
      name: 'Summary',
      cells: [
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 2 },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: '=A1*3' },
      ],
    });
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setCellValue('A1', 'First sheet');
    await runtime.setActiveSheet('Summary');
    expect(runtime.activeSheetName()).toBe('Summary');
    expect(runtime.activeSheet().cellsByAddress.get('B1')?.displayText).toBe('6');

    await runtime.setCellValue('A1', 5);
    expect(runtime.activeSheet().cellsByAddress.get('B1')?.displayText).toBe('15');
    await runtime.setActiveSheet('Sheet1');

    expect(runtime.activeSheetName()).toBe('Sheet1');
    expect(runtime.activeSheet().cellsByAddress.get('A1')?.value).toBe('First sheet');
    expect(await workbook.worksheets.getItem('Summary').getRange('A1:B1').values()).toEqual([[5, null]]);
    expect(await workbook.worksheets.getItem('Summary').getRange('B1').formulas()).toEqual([['=A1*3']]);
  });

  test('auto fills values, relative formulas, and styles in one undoable operation', async () => {
    const source = model();
    source.styles.fonts.push({ bold: true });
    source.styles.cellFormats.push({ fontId: 1 });
    source.sheets[0]!.cells = [
      { address: 'A1', row: 1, column: 1, kind: 'number', value: 1, styleId: 1 },
      { address: 'A2', row: 2, column: 1, kind: 'number', value: 2, styleId: 1 },
      { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'A1+10', styleId: 1 },
      { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'A2+10', styleId: 1 },
    ];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.autoFillRange(
      { startRow: 1, startColumn: 1, endRow: 2, endColumn: 2 },
      { startRow: 1, startColumn: 1, endRow: 5, endColumn: 2 },
    );

    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:A5').values()).toEqual([[1], [2], [3], [4], [5]]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('B3:B5').formulas()).toEqual([
      ['A3+10'],
      ['A4+10'],
      ['A5+10'],
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'A3', 'B5')).toEqual([1, 1]);
    expect(workbook.history().undoGroups()).toHaveLength(1);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A3:B5').values()).toEqual([
      [null, null],
      [null, null],
      [null, null],
    ]);
  });

  test('keeps direct worksheet structure operations out of the user undo stack', async () => {
    const source = model();
    source.sheets.push({ ...model().sheets[0]!, id: '2', name: 'Summary' });
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);
    await runtime.setCellValue('A1', 'User edit');
    const undoGroupsBeforeStructureChanges = workbook.history().undoGroups();

    await runtime.addSheet('Notes');
    expect(runtime.activeSheetName()).toBe('Notes');
    expect(runtime.model().sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Summary', 'Notes']);

    await runtime.renameSheet('Notes', 'Notes 2026');
    await runtime.copySheet('Notes 2026', 'Notes Copy');
    await runtime.moveSheet('Notes Copy', 0);
    expect(runtime.model().sheets.map((sheet) => sheet.name)).toEqual([
      'Notes Copy',
      'Sheet1',
      'Summary',
      'Notes 2026',
    ]);
    expect(workbook.history().undoGroups()).toEqual(undoGroupsBeforeStructureChanges);

    await runtime.deleteSheet('Notes Copy');
    expect(runtime.activeSheetName()).toBe('Sheet1');
    expect(runtime.model().sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Summary', 'Notes 2026']);

    await runtime.undo();
    expect(runtime.model().sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Summary', 'Notes 2026']);
    expect(runtime.activeSheet().cellsByAddress.get('A1')?.value).toBeUndefined();
  });

  test('records worksheet structure only inside an agent transaction', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.runAgentTransaction(
      { label: 'Agent creates worksheet', actor: 'spreadsheet-agent' },
      async (target) => {
        await target.addSheet('Agent Notes');
        await target.setCellValue('A1', 'created by agent');
      },
    );

    expect(runtime.activeSheetName()).toBe('Agent Notes');
    expect(workbook.history().undoGroups()[0]?.source).toBe('agent');
    expect(workbook.history().undoGroups()[0]?.transactions.map((transaction) => transaction.forward[0]?.kind)).toEqual([
      'workbook.sheets',
      'range.cells',
    ]);

    await runtime.undo();
    expect(runtime.model().sheets.map((sheet) => sheet.name)).toEqual(['Sheet1']);
  });

  test('replaces previous borders when switching all borders to outer border', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setRangeBorder(1, 1, 3, 3, 'all');
    expect(borderFor(await workbook.toModel(), 'B2')).toMatchObject({
      left: { style: 'thin' },
      right: { style: 'thin' },
      top: { style: 'thin' },
      bottom: { style: 'thin' },
    });

    await runtime.setRangeBorder(1, 1, 3, 3, 'outer');

    expect(borderFor(await workbook.toModel(), 'B2')).toEqual({});
    expect(borderFor(await workbook.toModel(), 'A1')).toMatchObject({
      left: { style: 'thin' },
      top: { style: 'thin' },
    });
    expect(borderFor(await workbook.toModel(), 'B1')).toEqual({
      top: { style: 'thin', color: { rgb: '#111827' } },
    });
  });

  test('sets vertical alignment through undo redo history', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setRangeVerticalAlign(1, 1, 1, 2, 'center');

    let current = await workbook.toModel();
    expect(current.styles.cellFormats[current.sheets[0]?.cells.find((cell) => cell.address === 'A1')?.styleId ?? 0]?.alignment).toMatchObject({
      vertical: 'center',
    });
    expect(current.styles.cellFormats[current.sheets[0]?.cells.find((cell) => cell.address === 'B1')?.styleId ?? 0]?.alignment).toMatchObject({
      vertical: 'center',
    });

    await runtime.undo();
    current = await workbook.toModel();
    expect(current.sheets[0]?.cells.find((cell) => cell.address === 'A1')?.styleId).toBeUndefined();
    expect(current.sheets[0]?.cells.find((cell) => cell.address === 'B1')?.styleId).toBeUndefined();

    await runtime.redo();
    current = await workbook.toModel();
    expect(current.styles.cellFormats[current.sheets[0]?.cells.find((cell) => cell.address === 'A1')?.styleId ?? 0]?.alignment).toMatchObject({
      vertical: 'center',
    });
  });

  test('exports edited cell formats and merged ranges through xlsx roundtrip', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setRangeValues(1, 1, [
      ['Header', 0.25],
      ['Merged note', 42],
      ['Styled detail', null],
    ]);
    await runtime.setRangeBold(1, 1, 1, 2, true);
    await runtime.setRangeItalic(3, 1, 1, 1, true);
    await runtime.setRangeUnderline(3, 1, 1, 1, true);
    await runtime.setRangeFontColor(1, 1, 3, 2, { rgb: '#C00000' });
    await runtime.setRangeFillColor(1, 1, 1, 2, { rgb: '#FFF2CC' });
    await runtime.setRangeNumberFormat(1, 2, 1, 1, '0.00%');
    await runtime.setRangeHorizontalAlign(1, 1, 2, 2, 'center');
    await runtime.setRangeVerticalAlign(1, 1, 2, 2, 'bottom');
    await runtime.setRangeBorder(1, 1, 2, 2, 'outer');
    await runtime.mergeRange(2, 1, 1, 2);

    const roundtrip = await ExcelWorkbook.fromZip(await runtime.toZip(), { workbookId: 'editor-roundtrip.xlsx' });
    const sheet = roundtrip.worksheets.getItem('Sheet1');
    const modelAfterRoundtrip = await roundtrip.toModel();
    const a1Style = cellFormat(modelAfterRoundtrip, 'A1');
    const b1Style = cellFormat(modelAfterRoundtrip, 'B1');
    const a3Style = cellFormat(modelAfterRoundtrip, 'A3');

    expect(await sheet.getRange('A1:B3').values()).toEqual([
      ['Header', 0.25],
      ['Merged note', null],
      ['Styled detail', null],
    ]);
    expect(await sheet.mergedRanges()).toEqual([{ address: 'A2:B2' }]);
    expect(modelAfterRoundtrip.styles.fonts[a1Style.fontId ?? 0]).toMatchObject({
      bold: true,
      color: { rgb: '#C00000' },
    });
    expect(modelAfterRoundtrip.styles.fonts[a3Style.fontId ?? 0]).toMatchObject({
      italic: true,
      underline: true,
      color: { rgb: '#C00000' },
    });
    expect(modelAfterRoundtrip.styles.fills[a1Style.fillId ?? 0]).toMatchObject({
      foregroundColor: { rgb: '#FFF2CC' },
    });
    expect(b1Style).toMatchObject({
      numberFormatCode: '0.00%',
      alignment: { horizontal: 'center', vertical: 'bottom' },
    });
    expect(modelAfterRoundtrip.styles.borders[a1Style.borderId ?? 0]).toMatchObject({
      left: { style: 'thin' },
      top: { style: 'thin' },
    });
    expect(modelAfterRoundtrip.styles.borders[b1Style.borderId ?? 0]).toMatchObject({
      right: { style: 'thin' },
      top: { style: 'thin' },
    });
  });

  test('exports cleared formats through xlsx roundtrip while preserving values', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setRangeValues(1, 1, [['Keep me']]);
    await runtime.setRangeBold(1, 1, 1, 1, true);
    await runtime.setRangeItalic(1, 1, 1, 1, true);
    await runtime.setRangeUnderline(1, 1, 1, 1, true);
    await runtime.setRangeFontColor(1, 1, 1, 1, { rgb: '#C00000' });
    await runtime.setRangeFillColor(1, 1, 1, 1, { rgb: '#FFF2CC' });
    await runtime.setRangeNumberFormat(1, 1, 1, 1, '0.00%');
    await runtime.setRangeHorizontalAlign(1, 1, 1, 1, 'center');
    await runtime.setRangeVerticalAlign(1, 1, 1, 1, 'bottom');
    await runtime.setRangeBorder(1, 1, 1, 1, 'all');

    await runtime.clearRangeFormat(1, 1, 1, 1);

    const roundtrip = await ExcelWorkbook.fromZip(await runtime.toZip(), { workbookId: 'clear-format.xlsx' });
    const modelAfterRoundtrip = await roundtrip.toModel();
    const format = cellFormat(modelAfterRoundtrip, 'A1');

    expect(await roundtrip.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([['Keep me']]);
    expect(modelAfterRoundtrip.styles.fonts[format.fontId ?? 0] ?? {}).toEqual({});
    expect(modelAfterRoundtrip.styles.fills[format.fillId ?? 0]?.foregroundColor).toBeUndefined();
    expect(modelAfterRoundtrip.styles.borders[format.borderId ?? 0] ?? {}).toEqual({});
    expect(format.numberFormatCode === undefined || format.numberFormatCode === 'General').toBe(true);
    expect(format.alignment).toBeUndefined();
  });

  test('exports pasted formats through xlsx roundtrip without changing target values', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setRangeValues(1, 1, [['Source', 0.5]]);
    await runtime.setRangeValues(3, 1, [['Target', 99]]);
    await runtime.setRangeBold(1, 1, 1, 2, true);
    await runtime.setRangeFontColor(1, 1, 1, 2, { rgb: '#C00000' });
    await runtime.setRangeFillColor(1, 1, 1, 2, { rgb: '#FFF2CC' });
    await runtime.setRangeNumberFormat(1, 2, 1, 1, '0.00%');
    await runtime.setRangeHorizontalAlign(1, 1, 1, 2, 'center');
    await runtime.setRangeVerticalAlign(1, 1, 1, 2, 'bottom');
    await runtime.setRangeBorder(1, 1, 1, 2, 'outer');

    const payload = buildSpreadsheetClipboardPayload(runtime.activeSheet(), {
      startRow: 1,
      startColumn: 1,
      endRow: 1,
      endColumn: 2,
    });
    await runtime.pasteRange(3, 1, payload.values, {
      mode: 'formats',
      styleIds: payload.styleIds,
    });

    const roundtrip = await ExcelWorkbook.fromZip(await runtime.toZip(), { workbookId: 'paste-formats.xlsx' });
    const modelAfterRoundtrip = await roundtrip.toModel();
    const a3Style = cellFormat(modelAfterRoundtrip, 'A3');
    const b3Style = cellFormat(modelAfterRoundtrip, 'B3');

    expect(await roundtrip.worksheets.getItem('Sheet1').getRange('A3:B3').values()).toEqual([['Target', 99]]);
    expect(modelAfterRoundtrip.styles.fonts[a3Style.fontId ?? 0]).toMatchObject({
      bold: true,
      color: { rgb: '#C00000' },
    });
    expect(modelAfterRoundtrip.styles.fills[a3Style.fillId ?? 0]).toMatchObject({
      foregroundColor: { rgb: '#FFF2CC' },
    });
    expect(b3Style).toMatchObject({
      numberFormatCode: '0.00%',
      alignment: { horizontal: 'center', vertical: 'bottom' },
    });
    expect(modelAfterRoundtrip.styles.borders[a3Style.borderId ?? 0]).toMatchObject({
      left: { style: 'thin' },
      top: { style: 'thin' },
      bottom: { style: 'thin' },
    });
    expect(modelAfterRoundtrip.styles.borders[b3Style.borderId ?? 0]).toMatchObject({
      right: { style: 'thin' },
      top: { style: 'thin' },
      bottom: { style: 'thin' },
    });
  });

  test('round-trips row height and column width changes through undo redo', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setColumnWidth(2, 144);
    await runtime.setRowHeight(3, 36);

    expect(runtime.activeSheet().columnMetaByIndex.get(2)?.width).toBeCloseTo(19.85546875);
    expect(runtime.activeSheet().rowMetaByIndex.get(3)?.height).toBe(36);

    await runtime.undo();
    expect(runtime.activeSheet().columnMetaByIndex.get(2)?.width).toBeCloseTo(19.85546875);
    expect(runtime.activeSheet().rowMetaByIndex.get(3)?.height).toBeUndefined();

    await runtime.undo();
    expect(runtime.activeSheet().columnMetaByIndex.get(2)?.width).toBeUndefined();
    expect(runtime.activeSheet().rowMetaByIndex.get(3)?.height).toBeUndefined();

    await runtime.redo();
    await runtime.redo();
    expect(runtime.activeSheet().columnMetaByIndex.get(2)?.width).toBeCloseTo(19.85546875);
    expect(runtime.activeSheet().rowMetaByIndex.get(3)?.height).toBe(36);
  });

  test('freezes panes through a lightweight view patch with undo redo and xlsx export', async () => {
    const workbook = await ExcelWorkbook.fromModel(model(), { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setFreezePanes(0, 0);
    expect(workbook.history().canUndo()).toBe(false);

    await runtime.setFreezePanes(3, 2);
    expect(runtime.activeSheet().source.sheetViews[0]?.pane).toEqual({
      xSplit: 2,
      ySplit: 3,
      topLeftCell: 'C4',
      activePane: 'bottomRight',
      state: 'frozen',
    });
    expect(workbook.history().undoGroups()[0]?.transactions[0]?.forward[0]?.kind).toBe('sheet.views');

    await runtime.undo();
    expect(runtime.activeSheet().source.sheetViews).toEqual([]);

    await runtime.redo();
    const roundtrip = await ExcelWorkbook.fromZip(await runtime.toZip(), { workbookId: 'frozen-panes.xlsx' });
    expect((await roundtrip.worksheets.getItem('Sheet1').sheetViews())[0]?.pane).toEqual({
      xSplit: 2,
      ySplit: 3,
      topLeftCell: 'C4',
      activePane: 'bottomRight',
      state: 'frozen',
    });

    await runtime.setFreezePanes(0, 0);
    expect(runtime.activeSheet().source.sheetViews[0]?.pane).toBeUndefined();
  });

  test('redirects merged follower cell edits to the merge master', async () => {
    const source = model();
    source.sheets[0]!.mergedRanges = [{ address: 'A1:B1' }];
    source.sheets[0]!.cells = [
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'master' },
      { address: 'B1', row: 1, column: 2, kind: 'blank', value: null },
    ];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    expect(runtime.editableCellAddress('B1')).toBe('A1');

    await runtime.setCellValue('B1', 'edited');
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['edited', null],
    ]);

    await runtime.setCellFormula('B1', '=1+1');
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').formulas()).toEqual([
      ['1+1', null],
    ]);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['edited', null],
    ]);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['master', null],
    ]);
  });

  test('keeps merged range outer border visible when edge styles are stored on followers', async () => {
    const source = model();
    source.sheets[0]!.mergedRanges = [{ address: 'A1:B1' }];
    source.sheets[0]!.cells = [
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'merged' },
      { address: 'B1', row: 1, column: 2, kind: 'blank', value: null },
    ];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.setRangeBorder(1, 1, 1, 2, 'all');
    await runtime.setRangeBorder(1, 1, 1, 2, 'outer');

    const modelAfterBorder = await workbook.toModel();
    expect(borderFor(modelAfterBorder, 'A1')).toMatchObject({
      left: { style: 'thin' },
      top: { style: 'thin' },
      bottom: { style: 'thin' },
    });
    expect(borderFor(modelAfterBorder, 'B1')).toMatchObject({
      right: { style: 'thin' },
      top: { style: 'thin' },
      bottom: { style: 'thin' },
    });
    const renderModel = buildSpreadsheetSheetRenderModel(modelAfterBorder.sheets[0]!, modelAfterBorder.styles);
    expect(resolveSpreadsheetRangeBorder(renderModel, { startRow: 1, startColumn: 1, endRow: 1, endColumn: 2 })).toMatchObject({
      left: { style: 'thin' },
      right: { style: 'thin' },
      top: { style: 'thin' },
      bottom: { style: 'thin' },
    });
  });

  test('normalizes hidden follower styles when merging and unmerging in the editor', async () => {
    const source = model();
    source.styles.fonts.push({ bold: true }, { italic: true });
    source.styles.cellFormats.push({ fontId: 1 }, { fontId: 2 });
    source.sheets[0]!.cells = [
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'master', styleId: 1 },
      { address: 'B1', row: 1, column: 2, kind: 'string', value: 'covered', styleId: 2 },
    ];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.mergeRange(1, 1, 1, 2);
    expect(styleIdsFor(await workbook.toModel(), 'A1', 'B1')).toEqual([1, 1]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').values()).toEqual([
      ['master', null],
    ]);

    await runtime.unmergeRanges(['A1:B1']);
    expect(await workbook.worksheets.getItem('Sheet1').mergedRanges()).toEqual([]);
    expect(styleIdsFor(await workbook.toModel(), 'A1', 'B1')).toEqual([1, 1]);

    await runtime.undo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B1',
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'A1', 'B1')).toEqual([1, 1]);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').mergedRanges()).toEqual([]);
    expect(styleIdsFor(await workbook.toModel(), 'A1', 'B1')).toEqual([1, 2]);
  });

  test('keeps merged copy paste as one undoable editor action', async () => {
    const source = model();
    source.sheets[0]!.mergedRanges = [{ address: 'B2:C3' }];
    source.sheets[0]!.cells = [
      { address: 'B2', row: 2, column: 2, kind: 'string', value: 'merged' },
      { address: 'C2', row: 2, column: 3, kind: 'string', value: 'hidden stale' },
      { address: 'B3', row: 3, column: 2, kind: 'string', value: 'hidden stale 2' },
      { address: 'D4', row: 4, column: 4, kind: 'number', value: 4 },
    ];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.pasteRangeValues(
      6,
      6,
      [
        ['merged', null],
        [null, null],
      ],
      [{ startRowOffset: 0, startColumnOffset: 0, rowCount: 2, columnCount: 2 }],
    );

    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'B2:C3',
      'F6:G7',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('F6:G7').values()).toEqual([
      ['merged', null],
      [null, null],
    ]);
    expect(workbook.history().undoGroups()).toHaveLength(1);

    await runtime.undo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'B2:C3',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('F6:G7').values()).toEqual([
      [null, null],
      [null, null],
    ]);

    await runtime.redo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'B2:C3',
      'F6:G7',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('F6:G7').values()).toEqual([
      ['merged', null],
      [null, null],
    ]);
  });

  test('pastes internal clipboard formatting with values and undo redo', async () => {
    const source = model();
    source.styles.fonts.push({ bold: true }, { italic: true });
    source.styles.fills.push({ foregroundColor: { rgb: '#FFFF00' } });
    source.styles.cellFormats.push({ fontId: 1, fillId: 1 }, { fontId: 2 });
    source.sheets[0]!.cells = [
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'bold fill', styleId: 1 },
      { address: 'B1', row: 1, column: 2, kind: 'string', value: 'italic', styleId: 2 },
      { address: 'D1', row: 1, column: 4, kind: 'string', value: 'target', styleId: 2 },
    ];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.pasteRangeValues(
      1,
      4,
      [['bold fill', 'italic']],
      [],
      [[1, 2]],
    );

    expect(await workbook.worksheets.getItem('Sheet1').getRange('D1:E1').values()).toEqual([
      ['bold fill', 'italic'],
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'D1', 'E1')).toEqual([1, 2]);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('D1:E1').values()).toEqual([
      ['target', null],
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'D1', 'E1')).toEqual([2, undefined]);

    await runtime.redo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('D1:E1').values()).toEqual([
      ['bold fill', 'italic'],
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'D1', 'E1')).toEqual([1, 2]);
  });

  test('supports paste special values and formats modes', async () => {
    const source = model();
    source.styles.fonts.push({ bold: true }, { italic: true });
    source.styles.fills.push({ foregroundColor: { rgb: '#FFFF00' } });
    source.styles.cellFormats.push({ fontId: 1, fillId: 1 }, { fontId: 2 });
    source.sheets[0]!.cells = [
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'source bold', styleId: 1 },
      { address: 'B1', row: 1, column: 2, kind: 'string', value: 'source italic', styleId: 2 },
      { address: 'D1', row: 1, column: 4, kind: 'string', value: 'old d', styleId: 2 },
      { address: 'E1', row: 1, column: 5, kind: 'string', value: 'old e' },
      { address: 'D3', row: 3, column: 4, kind: 'string', value: 'keep d', styleId: 2 },
      { address: 'E3', row: 3, column: 5, kind: 'string', value: 'keep e' },
    ];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.pasteRange(1, 4, [['new d', 'new e']], {
      mode: 'values',
      styleIds: [[1, 2]],
    });
    expect(await workbook.worksheets.getItem('Sheet1').getRange('D1:E1').values()).toEqual([
      ['new d', 'new e'],
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'D1', 'E1')).toEqual([2, undefined]);

    await runtime.pasteRange(3, 4, [['ignored d', 'ignored e']], {
      mode: 'formats',
      styleIds: [[1, 2]],
    });
    expect(await workbook.worksheets.getItem('Sheet1').getRange('D3:E3').values()).toEqual([
      ['keep d', 'keep e'],
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'D3', 'E3')).toEqual([1, 2]);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('D3:E3').values()).toEqual([
      ['keep d', 'keep e'],
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'D3', 'E3')).toEqual([2, undefined]);

    await runtime.undo();
    expect(await workbook.worksheets.getItem('Sheet1').getRange('D1:E1').values()).toEqual([
      ['old d', 'old e'],
    ]);
    expect(styleIdsFor(await workbook.toModel(), 'D1', 'E1')).toEqual([2, undefined]);
  });

  test('round-trips merged ranges through row and column deletion undo redo', async () => {
    const source = model();
    source.sheets[0]!.mergedRanges = [
      { address: 'A1:C3' },
      { address: 'D1:D4' },
    ];
    source.sheets[0]!.cells = [
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'wide' },
      { address: 'D1', row: 1, column: 4, kind: 'string', value: 'tall' },
      { address: 'A4', row: 4, column: 1, kind: 'string', value: 'outside' },
      { address: 'C4', row: 4, column: 3, kind: 'string', value: 'keep' },
    ];
    const workbook = await ExcelWorkbook.fromModel(source, { mode: 'inline' });
    const runtime = await SpreadsheetEditorRuntime.fromWorkbook(workbook);

    await runtime.runUserTransaction({ label: 'Delete through merges' }, async () => {
      await runtime.deleteRows(2, 1);
      await runtime.deleteColumns(2, 1);
    });

    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B2',
      'C1:C3',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:C3').values()).toEqual([
      ['wide', null, 'tall'],
      [null, null, null],
      ['outside', 'keep', null],
    ]);
    expect(workbook.history().undoGroups()).toHaveLength(1);

    await runtime.undo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:C3',
      'D1:D4',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:D4').values()).toEqual([
      ['wide', null, null, 'tall'],
      [null, null, null, null],
      [null, null, null, null],
      ['outside', null, 'keep', null],
    ]);

    await runtime.redo();
    expect((await workbook.worksheets.getItem('Sheet1').mergedRanges()).map((range) => range.address)).toEqual([
      'A1:B2',
      'C1:C3',
    ]);
    expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:C3').values()).toEqual([
      ['wide', null, 'tall'],
      [null, null, null],
      ['outside', 'keep', null],
    ]);
  });
});

function borderFor(model: ExcelWorkbookModel, address: string): ExcelBorderModel {
  const cell = model.sheets[0]?.cells.find((item) => item.address === address);
  const format = model.styles.cellFormats[cell?.styleId ?? 0] ?? {};
  return model.styles.borders[format.borderId ?? 0] ?? {};
}

function cellFormat(model: ExcelWorkbookModel, address: string): NonNullable<ExcelWorkbookModel['styles']['cellFormats'][number]> {
  const cell = model.sheets[0]?.cells.find((item) => item.address === address);
  return model.styles.cellFormats[cell?.styleId ?? 0] ?? {};
}

function styleIdsFor(model: ExcelWorkbookModel, ...addresses: string[]): Array<number | undefined> {
  return addresses.map((address) => model.sheets[0]?.cells.find((item) => item.address === address)?.styleId);
}

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
