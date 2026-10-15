import { execSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { ExcelWorkbook, ExcelWorkbookCore, type ExcelWorkbookModel } from '../index.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = resolve(__dirname, '../../../../..');
const workerBundleUrl = resolve(rootDir, 'packages/core/dist/excel-worker.js');

const MODEL: ExcelWorkbookModel = {
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
      cells: [
        {
          address: 'A1',
          row: 1,
          column: 1,
          kind: 'number',
          value: 21,
        },
      ],
    },
  ],
};

beforeAll(() => {
  execSync('corepack pnpm --filter @coco-vexel/coco-office build', { cwd: rootDir, stdio: 'pipe' });
  ExcelWorkbook.setWorkerBundleUrl(workerBundleUrl);
}, 120_000);

afterAll(() => {
  ExcelWorkbook._resetDefaults();
});

describe('ExcelWorkbook async runtime facade', () => {
  test('inline mode supports async worksheet/range/cell APIs', async () => {
    const workbook = await ExcelWorkbook.fromModel(MODEL, {
      workbookId: 'inline-book',
      mode: 'inline',
    });

    const sheet = workbook.worksheets.getItem('Sheet1');
    await sheet.getCell('B1').setFormula('A1*2');

    expect(workbook.mode).toBe('inline');
    expect(await sheet.getCell('B1').formula()).toBe('A1*2');
    const model = await workbook.toModel();
    expect(model.sheets[0]?.cells.find((cell) => cell.address === 'B1')?.formula).toBe('A1*2');
    await workbook.dispose();
  });

  test('worker mode supports same async worksheet/range/cell APIs and exports xlsx bytes', async () => {
    const decisions: string[] = [];
    const workbook = await ExcelWorkbook.fromModel(MODEL, {
      workbookId: 'worker-book',
      mode: 'worker',
      onModeResolved: (decision) => decisions.push(`${decision.mode}:${decision.source}`),
    });

    const sheet = workbook.worksheets.getItem('Sheet1');
    await sheet.getRange('B1:C1').setValues([[7, 8]]);
    await sheet.getCell('D1').setFormula('B1+C1');

    expect(workbook.mode).toBe('worker');
    expect(decisions).toEqual(['worker:explicit-worker']);
    expect(await sheet.getCell('D1').formula()).toBe('B1+C1');
    expect(
      ((await sheet.getCell('B1').value()) as number) +
        ((await sheet.getCell('C1').value()) as number),
    ).toBe(15);

    const model = await workbook.toModel();
    expect(model.sheets[0]?.cells.find((cell) => cell.address === 'D1')?.formula).toBe('B1+C1');

    const bytes = await workbook.toZip();
    expect(bytes).toBeInstanceOf(Uint8Array);
    const roundtrip = await ExcelWorkbookCore.fromZip(bytes, { workbookId: 'roundtrip' });
    expect(roundtrip.worksheets.getItem('Sheet1').getCell('B1').value()).toBe(7);
    expect(roundtrip.worksheets.getItem('Sheet1').getCell('D1').formula()).toBe('B1+C1');
    await workbook.dispose();
  }, 120_000);

  test('worker mode supports worksheet lifecycle, names, and basic formatting APIs', async () => {
    const workbook = await ExcelWorkbook.fromModel(MODEL, {
      workbookId: 'worker-lifecycle-book',
      mode: 'worker',
    });

    const added = await workbook.worksheets.add('Added');
    await added.getCell('A1').setValue('added');
    await added.setVisibility('hidden');
    await added.getRange('A1:B2').setStyleId(1);
    await added.getRange('A1:B2').setRowHeight(22);
    await added.getRange('A1:B2').setColumnWidth(16);
    await added.getRange('A1:B2').merge();
    await added.getRange('A1:B2').format.font.setBold(true);
    await added.getRange('A1:B2').format.font.setColor({ rgb: '#FF0000' });
    await added.getRange('A1:B2').format.fill.setPatternType('solid');
    await added.getRange('A1:B2').format.fill.setForegroundColor({ rgb: '#FFFF00' });
    await added
      .getRange('A1:B2')
      .format.border.setBottom({ style: 'thin', color: { rgb: '#000000' } });
    await added.getRange('A1:B2').format.alignment.setHorizontal('center');
    await added.getRange('A1:B2').format.alignment.setWrapText(true);
    await added.getRange('A1:B2').format.setNumberFormatCode('0.00%');
    await workbook.worksheets.copy('Sheet1', 'Sheet Copy');
    await workbook.worksheets.move('Added', 0);
    await workbook.worksheets.deleteItem('Sheet Copy');

    const named = await workbook.names.add({ name: 'AddedRef', text: 'Added!$A$1' });
    await named.setText('Added!$A$1:$B$2');
    await named.rename('AddedRange');
    await named.setLocalSheetId(0);

    const worksheetNames = await Promise.all(
      (await workbook.worksheets.items()).map((sheet) => sheet.name()),
    );
    expect(worksheetNames).toEqual(['Added', 'Sheet1']);
    expect(await added.visibility()).toBe('hidden');
    expect(await added.getCell('A1').styleId()).toBeGreaterThan(0);
    expect(await added.getRange('A1:B2').rowHeights()).toEqual([22, 22]);
    expect(await added.getRange('A1:B2').columnWidths()).toEqual([16, 16]);
    expect(await added.getRange('A1').format.font.bold()).toBe(true);
    expect(await added.getRange('A1').format.font.color()).toEqual({ rgb: '#FF0000' });
    expect(await added.getRange('A1').format.fill.foregroundColor()).toEqual({ rgb: '#FFFF00' });
    expect(await added.getRange('A1').format.border.value()).toMatchObject({
      bottom: { style: 'thin', color: { rgb: '#000000' } },
    });
    expect(await added.getRange('A1').format.alignment.value()).toMatchObject({
      horizontal: 'center',
      wrapText: true,
    });
    expect(await added.getRange('A1').format.numberFormatCode()).toBe('0.00%');
    expect(await workbook.names.getItem('AddedRange').text()).toBe('Added!$A$1:$B$2');
    expect(await workbook.names.getItem('AddedRange').localSheetId()).toBe(0);

    await workbook.names.deleteItem('AddedRange');
    expect(await workbook.names.getItemOrNullObject('AddedRange')).toBeNull();

    const model = await workbook.toModel();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(['Added', 'Sheet1']);
    expect(model.sheets[0]?.visibility).toBe('hidden');
    expect(model.sheets[0]?.mergedRanges).toEqual([{ address: 'A1:B2' }]);
    expect(model.styles.fonts.some((font) => font.bold && font.color?.rgb === '#FF0000')).toBe(
      true,
    );
    expect(model.styles.fills.some((fill) => fill.foregroundColor?.rgb === '#FFFF00')).toBe(true);

    const roundtrip = await ExcelWorkbookCore.fromZip(await workbook.toZip(), {
      workbookId: 'roundtrip-lifecycle',
    });
    expect(roundtrip.worksheets.items().map((sheet) => sheet.name())).toEqual(['Added', 'Sheet1']);
    expect(roundtrip.worksheets.getItem('Added').visibility()).toBe('hidden');
    expect(roundtrip.worksheets.getItem('Added').getCell('A1').value()).toBe('added');
    expect(roundtrip.worksheets.getItem('Added').mergedRanges()).toEqual([{ address: 'A1:B2' }]);
    const roundtripStyleId = roundtrip.worksheets.getItem('Added').getCell('A1').styleId();
    expect(roundtripStyleId).toBeGreaterThan(0);
    expect(roundtrip.styles().cellFormats[roundtripStyleId!]).toMatchObject({
      numberFormatCode: '0.00%',
      applyFont: true,
      applyFill: true,
      applyBorder: true,
      applyAlignment: true,
    });
    await workbook.dispose();
  }, 120_000);

  test('worker mode supports async sheet object domains without leaking core objects', async () => {
    const workbook = await ExcelWorkbook.fromModel(MODEL, {
      workbookId: 'worker-objects-book',
      mode: 'worker',
    });

    const sheet = workbook.worksheets.getItem('Sheet1');
    await sheet.getRange('A1:B2').setValues([
      ['Name', 'Amount'],
      ['alpha', 42],
    ]);

    const table = await sheet.tables.add({
      id: 1,
      name: 'Table1',
      displayName: 'Table1',
      ref: 'A1:B2',
      columns: [
        { id: 1, name: 'Name' },
        { id: 2, name: 'Amount' },
      ],
      style: { name: 'TableStyleMedium2', showRowStripes: true },
    });
    await table.setDisplayName('TableRenamed');
    await table.resize('A1:C3');
    await table.setHeaderRowCount(1);
    await table.setTotalsRowCount(1);
    await table.setTotalsRowShown(true);
    await table.setAutoFilter({
      ref: 'A1:C3',
      filterColumns: [{ colId: 1, filters: ['42'] }],
      sortState: {
        ref: 'A2:C3',
        conditions: [{ ref: 'B2:B3', descending: true }],
      },
    });
    await table.style.setShowColumnStripes(true);
    await table.columns.getItem('Amount').setTotalsRowFunction('sum');
    await table.columns.getItem('Column3').setTotalsRowLabel('total');
    await table.columns.add({ id: 4, name: 'Temp' });
    await table.columns.deleteItem('Temp');
    await sheet.comments.add({ ref: 'B2', author: 'Coco', text: 'checked' });
    await sheet.hyperlinks.add({ ref: 'A2', target: 'https://example.com', display: 'alpha' });
    const validation = await sheet.dataValidations.add({
      sqref: 'B2',
      type: 'whole',
      operator: 'between',
      allowBlank: false,
      formula1: '1',
      formula2: '100',
    });
    const format = await sheet.conditionalFormats.add({
      sqref: 'B2',
      rules: [{ type: 'cellIs', operator: 'greaterThan', priority: 1, formulas: ['10'] }],
    });

    expect(await table.displayName()).toBe('TableRenamed');
    expect(await table.rangeAddress()).toBe('A1:C3');
    expect(await table.headerRowCount()).toBe(1);
    expect(await table.totalsRowCount()).toBe(1);
    expect(await table.totalsRowShown()).toBe(true);
    expect(await table.autoFilter()).toMatchObject({
      ref: 'A1:C3',
      filterColumns: [{ colId: 1, filters: ['42'] }],
      sortState: { ref: 'A2:C3', conditions: [{ ref: 'B2:B3', descending: true }] },
    });
    expect(await table.style.showColumnStripes()).toBe(true);
    expect(await table.columns.getItem('Amount').totalsRowFunction()).toBe('sum');
    expect(await table.columns.getItem('Column3').totalsRowLabel()).toBe('total');
    expect((await table.columns.items()).length).toBe(3);
    const comment = sheet.comments.getItem('B2');
    await comment.setText('updated comment');
    expect(await comment.text()).toBe('updated comment');

    const hyperlink = sheet.hyperlinks.getItem('A2');
    await hyperlink.setDisplay('renamed alpha');
    expect(await hyperlink.display()).toBe('renamed alpha');
    expect(await hyperlink.target()).toBe('https://example.com');

    await validation.setFormula2('200');
    await validation.setAllowBlank(true);
    expect(await validation.formula2()).toBe('200');
    expect(await validation.allowBlank()).toBe(true);

    const rule = format.rules.getItem(0);
    await rule.addFormula('20');
    await rule.setPriority(2);
    expect(await rule.formulas()).toEqual(['10', '20']);
    expect(await rule.priority()).toBe(2);

    const model = await workbook.toModel();
    expect(model.sheets[0]?.tables[0]?.displayName).toBe('TableRenamed');
    expect(model.sheets[0]?.tables[0]?.ref).toBe('A1:C3');
    expect(model.sheets[0]?.tables[0]?.headerRowCount).toBe(1);
    expect(model.sheets[0]?.tables[0]?.totalsRowCount).toBe(1);
    expect(model.sheets[0]?.tables[0]?.totalsRowShown).toBe(true);
    expect(model.sheets[0]?.tables[0]?.autoFilter).toMatchObject({
      ref: 'A1:C3',
      filterColumns: [{ colId: 1, filters: ['42'] }],
    });
    expect(model.sheets[0]?.tables[0]?.style?.showColumnStripes).toBe(true);
    expect(model.sheets[0]?.tables[0]?.columns).toHaveLength(3);
    expect(model.sheets[0]?.tables[0]?.columns[1]?.totalsRowFunction).toBe('sum');
    expect(model.sheets[0]?.tables[0]?.columns[2]?.totalsRowLabel).toBe('total');
    expect(model.sheets[0]?.comments[0]?.text).toBe('updated comment');
    expect(model.sheets[0]?.hyperlinks[0]?.target).toBe('https://example.com');
    expect(model.sheets[0]?.hyperlinks[0]?.display).toBe('renamed alpha');
    expect(model.sheets[0]?.dataValidations[0]).toMatchObject({
      formula2: '200',
      allowBlank: true,
    });
    expect(model.sheets[0]?.conditionalFormats[0]?.rules[0]).toMatchObject({
      priority: 2,
      formulas: ['10', '20'],
    });

    const bytes = await workbook.toZip();
    const roundtrip = await ExcelWorkbookCore.fromZip(bytes, { workbookId: 'roundtrip-objects' });
    const roundtripSheet = roundtrip.worksheets.getItem('Sheet1');
    expect(roundtripSheet.tables.getItem('Table1').displayName()).toBe('TableRenamed');
    expect(roundtripSheet.tables.getItem('Table1').rangeAddress()).toBe('A1:C3');
    expect(roundtripSheet.tables.getItem('Table1').autoFilter()).toMatchObject({
      ref: 'A1:C3',
      filterColumns: [{ colId: 1, filters: ['42'] }],
    });
    expect(roundtripSheet.tables.getItem('Table1').style.showColumnStripes()).toBe(true);
    expect(roundtripSheet.hyperlinks.getItem('A2').display()).toBe('renamed alpha');
    expect(roundtripSheet.comments.getItem('B2').text()).toBe('updated comment');
    expect(roundtripSheet.dataValidations.getItem('B2').formula2()).toBe('200');
    expect(roundtripSheet.conditionalFormats.getItem('B2').rules.getItem(0).formulas()).toEqual([
      '10',
      '20',
    ]);

    await sheet.comments.delete('B2');
    expect(await sheet.comments.getItemOrNullObject('B2')).toBeNull();
    await workbook.dispose();
  }, 120_000);

  test('worker mode supports drawing/chart/image/shape live object APIs', async () => {
    const workbook = await ExcelWorkbook.fromModel(
      {
        ...MODEL,
        sheets: [
          {
            ...MODEL.sheets[0]!,
            drawings: [
              {
                relationshipId: 'rIdDrawing1',
                path: 'xl/drawings/drawing1.xml',
                anchors: [
                  {
                    kind: 'twoCell',
                    from: { row: 1, column: 1, rowOffset: 0, columnOffset: 0 },
                    to: { row: 5, column: 4 },
                    object: { kind: 'shape', name: 'Shape 1' },
                  },
                ],
              },
            ],
            charts: [
              {
                relationshipId: 'rIdChart1',
                drawingPath: 'xl/drawings/drawing1.xml',
                path: 'xl/charts/chart1.xml',
                title: 'Old title',
                formulaRefs: ['Sheet1!$A$1:$A$2'],
              },
            ],
            images: [
              {
                relationshipId: 'rIdImage1',
                drawingPath: 'xl/drawings/drawing1.xml',
                path: 'xl/media/image1.png',
                targetMode: 'Internal',
                name: 'Image 1',
              },
            ],
          },
        ],
      },
      { workbookId: 'worker-drawings-book', mode: 'worker' },
    );

    const sheet = workbook.worksheets.getItem('Sheet1');
    const drawing = (await sheet.drawings.items())[0]!;
    expect(await drawing.path()).toBe('xl/drawings/drawing1.xml');
    const anchor = drawing.anchors.getItem(0);
    expect(await anchor.kind()).toBe('twoCell');
    await anchor.from.setRow(2);
    expect(await anchor.from.row()).toBe(2);

    const chart = (await sheet.charts.items())[0]!;
    await chart.setTitle('New title');
    await chart.formulaRefs.setItem(0, 'Sheet1!$B$1:$B$2');
    await chart.formulaRefs.add('Sheet1!$C$1:$C$2');
    expect(await chart.title()).toBe('New title');
    expect(await chart.formulaRefs.items()).toEqual(['Sheet1!$B$1:$B$2', 'Sheet1!$C$1:$C$2']);
    const addedChart = await sheet.charts.add({
      type: 'pie',
      title: 'Generated chart',
      categoryRange: 'Sheet1!$A$1:$A$2',
      series: [{ name: 'Amount', values: 'Sheet1!$B$1:$B$2' }],
      position: { from: { row: 7, column: 1 }, to: { row: 14, column: 6 } },
    });
    expect(await addedChart.title()).toBe('Generated chart');

    const image = (await sheet.images.items())[0]!;
    expect(await image.name()).toBe('Image 1');
    expect(await image.path()).toBe('xl/media/image1.png');

    const shape = sheet.shapes.getItem('Shape 1');
    await shape.setName('Shape Renamed');
    expect(await shape.name()).toBe('Shape Renamed');

    const model = await workbook.toModel();
    expect(model.sheets[0]?.drawings[0]?.anchors[0]?.from?.row).toBe(2);
    expect(model.sheets[0]?.drawings[0]?.anchors[0]?.object?.name).toBe('Shape Renamed');
    expect(model.sheets[0]?.charts[0]?.title).toBe('New title');
    expect(model.sheets[0]?.charts[0]?.formulaRefs).toEqual([
      'Sheet1!$B$1:$B$2',
      'Sheet1!$C$1:$C$2',
    ]);
    expect(model.sheets[0]?.charts[1]?.type).toBe('pie');
    expect(model.sheets[0]?.drawings[0]?.anchors.at(-1)?.object?.kind).toBe('chart');
    await workbook.dispose();
  }, 120_000);

  test('worker mode terminates runtime on script timeout', async () => {
    const runtime = await ExcelWorkbook.fromModel(MODEL, {
      workbookId: 'timeout-book',
      mode: 'worker',
    });

    await expect(
      runtime.executeCode(
        `
        while (true) {}
      `,
        undefined,
        50,
      ),
    ).rejects.toThrow(/timeout/);

    await expect(runtime.toModel()).rejects.toThrow(/disposed/);
  }, 120_000);
});
