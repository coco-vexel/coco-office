import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, test } from 'vitest';
import { ExcelWorkbookCore as ExcelWorkbook } from '../index.js';

const fixtureDir = path.resolve(process.cwd(), 'packages/core/test-fixtures/excel');

async function readFixture(name: string): Promise<ExcelWorkbook> {
  return await ExcelWorkbook.fromZip(readFileSync(path.join(fixtureDir, name)), {
    workbookId: name,
  });
}

describe('ExcelWorkbook real xlsx fixtures', () => {
  test('reads a Microsoft Excel package exported from a live unsaved workbook', async () => {
    const workbook = await readFixture('microsoft-unsaved-marker.xlsx');
    const sheet = workbook.worksheets.getItem('Sheet1');
    const model = workbook.toModel();

    expect(workbook.workbookId).toBe('microsoft-unsaved-marker.xlsx');
    expect(workbook.worksheets.items().map((item) => item.name())).toEqual(['Sheet1']);
    expect(sheet.dimension()).toBe('A1:AL63');
    expect(sheet.sheetViews()).toEqual([
      {
        workbookViewId: 0,
        tabSelected: true,
        topLeftCell: 'Q1',
        selections: [{ activeCell: 'I23', sqref: 'I23' }],
      },
    ]);
    expect(sheet.getCell('A1').value()).toBe('__COCO_EXCEL_UNSAVED_1785722092803__');
    expect(sheet.getRange('A1:C2').values()).toEqual([
      ['__COCO_EXCEL_UNSAVED_1785722092803__', 111, 111],
      [2121, 2121, 2121],
    ]);
    expect(model.sharedStrings).toContain('__COCO_EXCEL_UNSAVED_1785722092803__');
    expect(model.sheets[0]?.cells).toHaveLength(2393);
  });

  test('reads a WPS SaveAsBinaryString package with the same workbook shape', async () => {
    const workbook = await readFixture('wps-save-as-binary-string.xlsx');
    const sheet = workbook.worksheets.getItem('Sheet1');
    const model = workbook.toModel();

    expect(workbook.workbookId).toBe('wps-save-as-binary-string.xlsx');
    expect(workbook.worksheets.items().map((item) => item.name())).toEqual(['Sheet1']);
    expect(sheet.dimension()).toBe('A1:AL63');
    expect(sheet.getCell('A1').value()).toBe('__COCO_WPS_FAST_1785726467275__');
    expect(sheet.getRange('A1:C2').values()).toEqual([
      ['__COCO_WPS_FAST_1785726467275__', 111, 111],
      [2121, 2121, 2121],
    ]);
    expect(model.sharedStrings).toContain('__COCO_WPS_FAST_1785726467275__');
    expect(model.sheets[0]?.cells).toHaveLength(2394);
  });

  test('reads a richer Microsoft probe package with multiple sheets and formulas', async () => {
    const workbook = await readFixture('microsoft-rich-objects.xlsx');
    const sheet = workbook.worksheets.getItem('COCO_22570659');
    const styles = workbook.styles();

    expect(workbook.worksheets.items().map((item) => item.name())).toEqual([
      'Sheet1',
      'COCO_22570659',
    ]);
    expect(sheet.sheetViews()).toEqual([
      {
        workbookViewId: 0,
        tabSelected: true,
        selections: [],
      },
    ]);
    expect(styles.fonts).toHaveLength(6);
    expect(styles.fills[2]).toEqual({
      patternType: 'solid',
      foregroundColor: { rgb: '#FFF2CC' },
      backgroundColor: { rgb: '#000000' },
    });
    expect(styles.cellFormats[1]).toMatchObject({
      fontId: 3,
      fillId: 2,
      applyFont: true,
      applyFill: true,
    });
    expect(styles.cellFormats[2]).toMatchObject({
      fontId: 4,
      fillId: 3,
      applyFont: true,
      applyFill: true,
    });
    expect(sheet.dimension()).toBe('A1:H8');
    expect(sheet.columns().slice(0, 3)).toEqual([
      { min: 1, max: 1, width: 29.08203125, customWidth: true },
      { min: 2, max: 2, width: 32.9140625, customWidth: true },
      { min: 3, max: 3, width: 30.08203125, customWidth: true },
    ]);
    expect(sheet.mergedRanges()).toEqual([{ address: 'A4:B4' }]);
    expect(sheet.hyperlinks.toModel()).toEqual([
      {
        ref: 'E2',
        relationshipId: 'rId1',
        target: 'https://example.com/?marker=__COCO_LINK_1785722570659__',
        targetMode: 'External',
        tooltip: '__COCO_LINK_1785722570659__',
      },
    ]);
    expect(sheet.tables.toModel()).toEqual([
      {
        id: 1,
        name: 'CocoProbeTable570659',
        displayName: 'CocoProbeTable570659',
        relationshipId: 'rId4',
        path: 'xl/tables/table1.xml',
        ref: 'A6:C8',
        totalsRowShown: false,
        autoFilterRef: 'A6:C8',
        columns: [
          { id: 1, name: 'Name' },
          { id: 2, name: 'Amount' },
          { id: 3, name: 'Marker' },
        ],
        style: {
          name: 'TableStyleMedium2',
          showFirstColumn: false,
          showLastColumn: false,
          showRowStripes: true,
          showColumnStripes: false,
        },
      },
    ]);
    expect(sheet.comments.toModel()).toHaveLength(1);
    expect(sheet.comments.toModel()[0]).toMatchObject({
      ref: 'F2',
      authorId: 0,
      shapeId: 0,
    });
    expect(sheet.comments.toModel()[0]?.text).toContain('__COCO_COMMENT_1785722570659__');
    expect(sheet.threadedComments.toModel()).toEqual([
      {
        ref: 'F2',
        id: '{5707AC8F-71D1-4B24-8572-8A50977FCEF2}',
        personId: '{228EF9CD-F571-4947-9DCB-C421F103AC48}',
        dateTime: '2026-08-03T02:02:53.14',
        text: '__COCO_COMMENT_1785722570659__',
      },
    ]);
    expect(sheet.drawings.toModel()).toHaveLength(1);
    expect(sheet.drawings.toModel()[0]).toMatchObject({
      relationshipId: 'rId2',
      path: 'xl/drawings/drawing1.xml',
      anchors: [
        {
          kind: 'twoCell',
          object: {
            kind: 'chart',
            relationshipId: 'rId1',
            targetPath: 'xl/charts/chart1.xml',
          },
        },
      ],
    });
    expect(sheet.charts.toModel()).toEqual([
      {
        relationshipId: 'rId1',
        drawingPath: 'xl/drawings/drawing1.xml',
        path: 'xl/charts/chart1.xml',
        title: '__COCO_CHART_1785722570659__',
        formulaRefs: ['COCO_22570659!$B$6', 'COCO_22570659!$A$7:$A$8', 'COCO_22570659!$B$7:$B$8'],
      },
    ]);
    expect(sheet.images.toModel()).toEqual([]);
    expect(sheet.getCell('A1').value()).toBe('__COCO_CELL_1785722570659__');
    expect(sheet.getCell('H2').formula()).toBe('"__COCO_FORMULA_1785722570659__"');
    expect(sheet.getCell('H2').value()).toBe('__COCO_FORMULA_1785722570659__');
    expect(sheet.getCell('H3').formula()).toBe('SUM(G2:G3)');
    expect(sheet.getCell('H3').value()).toBe(42);
    expect(sheet.getRange('A6:C8').values()).toEqual([
      ['Name', 'Amount', 'Marker'],
      ['alpha', 10, '__COCO_TABLE_1785722570659__'],
      ['beta', 32, 'row-2'],
    ]);
  });
});
