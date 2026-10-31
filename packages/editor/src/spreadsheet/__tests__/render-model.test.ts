import { describe, expect, test } from 'vitest';
import {
  calculateSpreadsheetSheetFormulas,
  type ExcelSheetModel,
  type ExcelStylesModel,
} from '@coco-vexel/coco-office/excel';
import { buildSpreadsheetSheetRenderModel, formatSpreadsheetDisplayValue } from '../render-model.js';
import { resolveSpreadsheetCellStyle, resolveSpreadsheetRangeBorder } from '../style-resolver.js';
import { SpreadsheetLayout } from '../layout.js';
import { SpreadsheetViewportManager, freezeStateFromSheet } from '../viewport.js';
import {
  SPREADSHEET_COLUMN_RESIZE_MIN,
  SPREADSHEET_ROW_RESIZE_MIN,
  hitSpreadsheetResizeHandle,
  startSpreadsheetResizeDrag,
  updateSpreadsheetResizeDrag,
} from '../resize.js';
import { summarizeSpreadsheetSelectionFormat } from '../format-state.js';
import { buildSpreadsheetReplacePlan, findSpreadsheetMatches, nextSpreadsheetMatch, rangeForSpreadsheetMatch } from '../search.js';
import {
  createColumnSelection,
  createColumnRangeSelection,
  createRowRangeSelection,
  createRowSelection,
  createSelection,
  expandSelectionForMergedRanges,
  extendSelection,
  selectionAfterColumnDeletion,
  selectionAfterRowDeletion,
} from '../selection.js';

describe('spreadsheet render model', () => {
  test('indexes cells, rows, columns, and merged ranges', () => {
    const renderModel = buildSpreadsheetSheetRenderModel(makeSheet(), makeStyles());

    expect(renderModel.cellsByAddress.get('A1')?.displayText).toBe('Name');
    expect(renderModel.cellsByRowColumn.get('2:2')?.displayText).toBe('42');
    expect(renderModel.rowMetaByIndex.get(2)?.height).toBe(30);
    expect(renderModel.columnMetaByIndex.get(2)?.width).toBe(17.1484375);
    expect(renderModel.mergeMasterByAddress.get('C1')?.range).toMatchObject({
      startRow: 1,
      startColumn: 3,
      endRow: 2,
      endColumn: 4,
    });
    expect(renderModel.mergeFollowerByAddress.get('D2')).toBe('C1');
  });

  test('shows a simple formula result when no cached value exists', () => {
    const sheet = makeSheet();
    sheet.cells.push(
      { address: 'C7', row: 7, column: 3, kind: 'number', value: 12171 },
      { address: 'D7', row: 7, column: 4, kind: 'number', value: 12171 },
      { address: 'E7', row: 7, column: 5, kind: 'formula', value: null, formula: 'C7+D7' },
    );
    const renderModel = buildSpreadsheetSheetRenderModel(sheet, makeStyles(), {
      formulaResults: calculateSpreadsheetSheetFormulas(sheet),
    });

    expect(renderModel.cellsByAddress.get('E7')?.displayText).toBe('24342');
  });

  test('projects dynamic array spill results into virtual cells', () => {
    const sheet = makeSheet();
    sheet.cells.push({ address: 'E1', row: 1, column: 5, kind: 'formula', value: null, formula: 'SEQUENCE(2,2,10,5)' });
    const renderModel = buildSpreadsheetSheetRenderModel(sheet, makeStyles(), {
      formulaResults: calculateSpreadsheetSheetFormulas(sheet),
    });

    expect(renderModel.cellsByAddress.get('E1')?.displayText).toBe('10');
    expect(renderModel.cellsByAddress.get('F1')?.displayText).toBe('15');
    expect(renderModel.cellsByAddress.get('E2')?.displayText).toBe('20');
    expect(renderModel.cellsByAddress.get('F2')?.displayText).toBe('25');
  });

  test('formats numeric display text from cell number formats', () => {
    const sheet = makeSheet();
    sheet.cells.push(
      { address: 'A3', row: 3, column: 1, kind: 'number', value: 1234, styleId: 1 },
      { address: 'B3', row: 3, column: 2, kind: 'number', value: 0.1234, styleId: 2 },
      { address: 'C3', row: 3, column: 3, kind: 'number', value: 45292, styleId: 3 },
      { address: 'D3', row: 3, column: 4, kind: 'number', value: 1234, styleId: 4 },
    );
    const renderModel = buildSpreadsheetSheetRenderModel(sheet, {
      ...makeStyles(),
      numberFormats: [
        { id: 164, code: '#,##0.00' },
        { id: 165, code: '0.00%' },
        { id: 166, code: 'yyyy-mm-dd' },
        { id: 167, code: '$#,##0.00' },
      ],
      cellFormats: [
        {},
        { numberFormatId: 164, numberFormatCode: '#,##0.00' },
        { numberFormatId: 165, numberFormatCode: '0.00%' },
        { numberFormatId: 166, numberFormatCode: 'yyyy-mm-dd' },
        { numberFormatId: 167, numberFormatCode: '$#,##0.00' },
      ],
    });

    expect(renderModel.cellsByAddress.get('A3')?.displayText).toBe('1,234.00');
    expect(renderModel.cellsByAddress.get('B3')?.displayText).toBe('12.34%');
    expect(renderModel.cellsByAddress.get('C3')?.displayText).toBe('2024-01-01');
    expect(renderModel.cellsByAddress.get('D3')?.displayText).toBe('$1,234.00');
  });

  test('formats supported numeric values directly', () => {
    expect(formatSpreadsheetDisplayValue(1234, '#,##0.00')).toBe('1,234.00');
    expect(formatSpreadsheetDisplayValue(0.5, '0.00%')).toBe('50.00%');
    expect(formatSpreadsheetDisplayValue(45292.75, 'yyyy-mm-dd')).toBe('2024-01-01');
    expect(formatSpreadsheetDisplayValue(60, 'yyyy-mm-dd')).toBe('1900-02-29');
    expect(formatSpreadsheetDisplayValue(1, 'yyyy-mm-dd', true)).toBe('1904-01-02');
    expect(formatSpreadsheetDisplayValue(45292.5, 'mm-dd-yy')).toBe('01-01-24');
    expect(formatSpreadsheetDisplayValue(45292.5, 'd-mmm-yy')).toBe('1-Jan-24');
    expect(formatSpreadsheetDisplayValue(45292.5, 'h:mm AM/PM')).toBe('12:00 PM');
    expect(formatSpreadsheetDisplayValue(45292.5, 'm/d/yy h:mm')).toBe('1/1/24 12:00');
  });

  test('resolves underline font style for canvas rendering', () => {
    const styles = makeStyles();
    styles.fonts.push({ underline: true, size: 12 });
    styles.cellFormats.push({ fontId: 1 });

    expect(resolveSpreadsheetCellStyle(styles, 1)).toMatchObject({
      fontSize: 12,
      underline: true,
    });
  });

  test('resolves Excel-like default cell alignment by value type', () => {
    const model = buildSpreadsheetSheetRenderModel(
      {
        ...makeSheet(),
        cells: [
          { address: 'A1', row: 1, column: 1, kind: 'string', value: 'Name' },
          { address: 'B1', row: 1, column: 2, kind: 'number', value: 42 },
          { address: 'C1', row: 1, column: 3, kind: 'boolean', value: true },
        ],
      },
      makeStyles(),
    );

    expect(resolveSpreadsheetCellStyle(model.styles, undefined, model.cellsByAddress.get('A1')).horizontalAlign).toBe(
      'left',
    );
    expect(resolveSpreadsheetCellStyle(model.styles, undefined, model.cellsByAddress.get('B1')).horizontalAlign).toBe(
      'right',
    );
    expect(resolveSpreadsheetCellStyle(model.styles, undefined, model.cellsByAddress.get('C1')).horizontalAlign).toBe(
      'center',
    );
    expect(resolveSpreadsheetCellStyle(model.styles, undefined, model.cellsByAddress.get('A1')).verticalAlign).toBe(
      'bottom',
    );
  });

  test('respects explicit horizontal and vertical alignment from cell styles', () => {
    const styles = makeStyles();
    styles.cellFormats.push({ alignment: { horizontal: 'center', vertical: 'center' } });
    const model = buildSpreadsheetSheetRenderModel(
      {
        ...makeSheet(),
        cells: [{ address: 'A1', row: 1, column: 1, kind: 'number', value: 42, styleId: 1 }],
      },
      styles,
    );

    expect(resolveSpreadsheetCellStyle(model.styles, 1, model.cellsByAddress.get('A1'))).toMatchObject({
      horizontalAlign: 'center',
      verticalAlign: 'middle',
    });
  });

  test('aggregates merged-range borders from edge cells', () => {
    const styles = makeStyles();
    styles.borders.push(
      { left: { style: 'thin', color: { rgb: '#111827' } }, top: { style: 'thin', color: { rgb: '#111827' } } },
      { right: { style: 'thin', color: { rgb: '#111827' } }, bottom: { style: 'thin', color: { rgb: '#111827' } } },
    );
    styles.cellFormats.push({ borderId: 1 }, { borderId: 2 });
    const sheet = makeSheet();
    sheet.mergedRanges = [{ address: 'A1:B1' }];
    sheet.cells = [
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'merged', styleId: 1 },
      { address: 'B1', row: 1, column: 2, kind: 'blank', value: null, styleId: 2 },
    ];
    const renderModel = buildSpreadsheetSheetRenderModel(sheet, styles);

    expect(resolveSpreadsheetRangeBorder(renderModel, { startRow: 1, startColumn: 1, endRow: 1, endColumn: 2 })).toEqual({
      left: { style: 'thin', color: { rgb: '#111827' } },
      right: { style: 'thin', color: { rgb: '#111827' } },
      top: { style: 'thin', color: { rgb: '#111827' } },
      bottom: { style: 'thin', color: { rgb: '#111827' } },
    });
  });
});

describe('spreadsheet layout', () => {
  test('resolves custom sizes and visible ranges', () => {
    const layout = new SpreadsheetLayout(buildSpreadsheetSheetRenderModel(makeSheet(), makeStyles()));

    expect(layout.getRowHeight(2)).toBe(40);
    expect(layout.getColumnWidth(2)).toBe(120);
    expect(layout.getColumnX(3)).toBe(216);
    expect(layout.getRowY(3)).toBe(62);
    expect(layout.getVisibleRange(0, 0, { width: 220, height: 60 })).toMatchObject({
      startRow: 1,
      startColumn: 1,
      endRow: 3,
      endColumn: 4,
    });
  });
});

describe('spreadsheet viewport manager', () => {
  test('uses viewports for frozen pane hit testing', () => {
    const sheet = makeSheet();
    const layout = new SpreadsheetLayout(buildSpreadsheetSheetRenderModel(sheet, makeStyles()));
    const manager = new SpreadsheetViewportManager(layout, {
      viewportSize: { width: 420, height: 240 },
      scroll: { x: 100, y: 50 },
      freeze: freezeStateFromSheet(sheet),
    });

    expect(manager.viewports.some((viewport) => viewport.kind === 'frozenTopLeft')).toBe(true);
    expect(manager.hitTest({ x: 54, y: 28 })).toMatchObject({
      kind: 'cell',
      row: 1,
      column: 1,
    });
    expect(manager.hitTest({ x: 260, y: 120 })).toMatchObject({ kind: 'cell' });
  });
});

describe('spreadsheet search', () => {
  test('finds display text, raw values, and formulas in row-major order', () => {
    const model = buildSpreadsheetSheetRenderModel(
      {
        ...makeSheet(),
        cells: [
          { address: 'A1', row: 1, column: 1, kind: 'string', value: 'Alpha' },
          { address: 'B1', row: 1, column: 2, kind: 'number', value: 42 },
          { address: 'A2', row: 2, column: 1, kind: 'formula', value: 84, formula: 'B1*2' },
          { address: 'C3', row: 3, column: 3, kind: 'string', value: 'alpha tail' },
        ],
      },
      makeStyles(),
    );

    expect(findSpreadsheetMatches(model, { query: 'alpha' }).map((match) => match.address)).toEqual(['A1', 'C3']);
    expect(findSpreadsheetMatches(model, { query: 'Alpha', matchCase: true }).map((match) => match.address)).toEqual(['A1']);
    expect(findSpreadsheetMatches(model, { query: '42' }).map((match) => match.address)).toEqual(['B1']);
    expect(findSpreadsheetMatches(model, { query: '=B1*2' }).map((match) => match.address)).toEqual(['A2']);
  });

  test('navigates search matches with wraparound', () => {
    const matches = [
      { row: 1, column: 1, address: 'A1', text: 'one' },
      { row: 2, column: 2, address: 'B2', text: 'two' },
      { row: 4, column: 1, address: 'A4', text: 'three' },
    ];

    expect(nextSpreadsheetMatch(matches, { row: 1, column: 1 }, 'forward')?.address).toBe('B2');
    expect(nextSpreadsheetMatch(matches, { row: 5, column: 1 }, 'forward')?.address).toBe('A1');
    expect(nextSpreadsheetMatch(matches, { row: 4, column: 1 }, 'backward')?.address).toBe('B2');
    expect(nextSpreadsheetMatch(matches, { row: 1, column: 1 }, 'backward')?.address).toBe('A4');
    expect(rangeForSpreadsheetMatch(matches[1]!)).toEqual({
      startRow: 2,
      startColumn: 2,
      endRow: 2,
      endColumn: 2,
    });
  });

  test('builds a replacement plan for values and formulas', () => {
    const model = buildSpreadsheetSheetRenderModel(
      {
        ...makeSheet(),
        cells: [
          { address: 'A1', row: 1, column: 1, kind: 'string', value: 'Alpha' },
          { address: 'A2', row: 2, column: 1, kind: 'formula', value: null, formula: 'A1+Alpha' },
        ],
      },
      makeStyles(),
    );

    expect(buildSpreadsheetReplacePlan(model, { query: 'alpha', replacement: 'Beta' })).toEqual([
      { address: 'A1', value: 'Beta' },
      { address: 'A2', formula: 'A1+Beta' },
    ]);
  });
});

describe('spreadsheet format state', () => {
  test('summarizes uniform selection formatting', () => {
    const styles = makeStyles();
    styles.fonts.push({ bold: true, italic: true, underline: true, color: { rgb: '#FF0000' } });
    styles.fills.push({ foregroundColor: { rgb: '#FFFF00' } });
    styles.borders.push({ top: { style: 'thin' } });
    styles.cellFormats.push({
      fontId: 1,
      fillId: 2,
      borderId: 1,
      numberFormatCode: '#,##0.00',
      alignment: { horizontal: 'center', vertical: 'center' },
    });
    const model = buildSpreadsheetSheetRenderModel(
      {
        ...makeSheet(),
        cells: [
          { address: 'A1', row: 1, column: 1, kind: 'number', value: 1, styleId: 1 },
          { address: 'A2', row: 2, column: 1, kind: 'number', value: 2, styleId: 1 },
        ],
      },
      styles,
    );

    expect(
      summarizeSpreadsheetSelectionFormat(model, {
        kind: 'cell',
        active: { row: 1, column: 1 },
        range: { startRow: 1, startColumn: 1, endRow: 2, endColumn: 1 },
      }),
    ).toMatchObject({
      font: {
        bold: true,
        italic: true,
        underline: true,
        color: { rgb: '#FF0000' },
      },
      fill: { foregroundColor: { rgb: '#FFFF00' } },
      numberFormatCode: '#,##0.00',
      alignment: { horizontal: 'center', vertical: 'center' },
      border: { top: { style: 'thin' } },
    });
  });

  test('marks mixed formatting across a selection', () => {
    const styles = makeStyles();
    styles.fonts.push({ bold: true, color: { rgb: '#FF0000' } }, { italic: true, color: { rgb: '#00AA00' } });
    styles.fills.push({ foregroundColor: { rgb: '#FFFF00' } }, { foregroundColor: { rgb: '#DDEEFF' } });
    styles.cellFormats.push(
      { fontId: 1, fillId: 2, numberFormatCode: '#,##0.00', alignment: { horizontal: 'left', vertical: 'top' } },
      { fontId: 2, fillId: 3, numberFormatCode: '0.00%', alignment: { horizontal: 'right', vertical: 'bottom' } },
    );
    const model = buildSpreadsheetSheetRenderModel(
      {
        ...makeSheet(),
        cells: [
          { address: 'A1', row: 1, column: 1, kind: 'number', value: 1, styleId: 1 },
          { address: 'B1', row: 1, column: 2, kind: 'number', value: 2, styleId: 2 },
        ],
      },
      styles,
    );

    expect(
      summarizeSpreadsheetSelectionFormat(model, {
        kind: 'cell',
        active: { row: 1, column: 1 },
        range: { startRow: 1, startColumn: 1, endRow: 1, endColumn: 2 },
      }),
    ).toMatchObject({
      font: {
        bold: 'mixed',
        italic: 'mixed',
        color: 'mixed',
      },
      fill: { foregroundColor: 'mixed' },
      numberFormatCode: 'mixed',
      alignment: { horizontal: 'mixed', vertical: 'mixed' },
    });
  });
});

describe('spreadsheet resize', () => {
  test('hits column and row resize handles in visible headers', () => {
    const sheet = makeSheet();
    const layout = new SpreadsheetLayout(buildSpreadsheetSheetRenderModel(sheet, makeStyles()));
    const manager = new SpreadsheetViewportManager(layout, {
      viewportSize: { width: 420, height: 240 },
      scroll: { x: 0, y: 0 },
      freeze: freezeStateFromSheet(sheet),
    });

    expect(hitSpreadsheetResizeHandle(layout, manager.viewports, { x: 144, y: 12 })).toMatchObject({
      kind: 'column',
      index: 1,
      startSize: 96,
      minSize: SPREADSHEET_COLUMN_RESIZE_MIN,
    });
    expect(hitSpreadsheetResizeHandle(layout, manager.viewports, { x: 24, y: 46 })).toMatchObject({
      kind: 'row',
      index: 1,
      startSize: 22,
      minSize: SPREADSHEET_ROW_RESIZE_MIN,
    });
    expect(hitSpreadsheetResizeHandle(layout, manager.viewports, { x: 80, y: 12 })).toBeNull();
  });

  test('updates resize drag size and guide position with minimum clamping', () => {
    const columnDrag = startSpreadsheetResizeDrag(
      { kind: 'column', index: 2, startSize: 96, guidePosition: 240, minSize: SPREADSHEET_COLUMN_RESIZE_MIN },
      { clientX: 240, clientY: 0 },
    );
    expect(updateSpreadsheetResizeDrag(columnDrag, { clientX: 276, clientY: 0 })).toMatchObject({
      currentSize: 132,
      currentGuidePosition: 276,
    });
    expect(updateSpreadsheetResizeDrag(columnDrag, { clientX: 120, clientY: 0 })).toMatchObject({
      currentSize: SPREADSHEET_COLUMN_RESIZE_MIN,
      currentGuidePosition: 168,
    });

    const rowDrag = startSpreadsheetResizeDrag(
      { kind: 'row', index: 3, startSize: 22, guidePosition: 90, minSize: SPREADSHEET_ROW_RESIZE_MIN },
      { clientX: 0, clientY: 90 },
    );
    expect(updateSpreadsheetResizeDrag(rowDrag, { clientX: 0, clientY: 120 })).toMatchObject({
      currentSize: 52,
      currentGuidePosition: 120,
    });
    expect(updateSpreadsheetResizeDrag(rowDrag, { clientX: 0, clientY: 40 })).toMatchObject({
      currentSize: SPREADSHEET_ROW_RESIZE_MIN,
      currentGuidePosition: 80,
    });
  });
});

describe('spreadsheet selection', () => {
  test('creates and extends row and column selections', () => {
    expect(createRowSelection(3, { columnCount: 8 })).toMatchObject({
      kind: 'row',
      active: { row: 3, column: 1 },
      range: { startRow: 3, startColumn: 1, endRow: 3, endColumn: 8 },
    });
    expect(
      extendSelection(createRowSelection(3, { columnCount: 8 }), 5, 8, {
        rowCount: 20,
        columnCount: 8,
      }),
    ).toMatchObject({
      kind: 'row',
      range: { startRow: 3, startColumn: 1, endRow: 5, endColumn: 8 },
    });
    expect(
      extendSelection(createColumnSelection(4, { rowCount: 20 }), 20, 6, {
        rowCount: 20,
        columnCount: 8,
      }),
    ).toMatchObject({
      kind: 'column',
      range: { startRow: 1, startColumn: 4, endRow: 20, endColumn: 6 },
    });
  });

  test('creates row and column range selections for inserted structure', () => {
    expect(createRowRangeSelection(5, 7, { columnCount: 8 })).toEqual({
      kind: 'row',
      active: { row: 5, column: 1 },
      range: { startRow: 5, startColumn: 1, endRow: 7, endColumn: 8 },
    });
    expect(createColumnRangeSelection(4, 2, { rowCount: 20 })).toEqual({
      kind: 'column',
      active: { row: 1, column: 2 },
      range: { startRow: 1, startColumn: 2, endRow: 20, endColumn: 4 },
    });
  });

  test('expands cell selections to include touched merged ranges', () => {
    const renderModel = buildSpreadsheetSheetRenderModel(
      {
        ...makeSheet(),
        mergedRanges: [
          { address: 'C1:D2' },
          { address: 'E2:F3' },
        ],
      },
      makeStyles(),
    );

    expect(expandSelectionForMergedRanges(createSelection(4, 5), renderModel).range).toEqual({
      startRow: 4,
      startColumn: 5,
      endRow: 4,
      endColumn: 5,
    });
    expect(
      expandSelectionForMergedRanges(
        {
          kind: 'cell',
          active: { row: 1, column: 3 },
          range: { startRow: 2, startColumn: 4, endRow: 2, endColumn: 5 },
        },
        renderModel,
      ).range,
    ).toEqual({
      startRow: 1,
      startColumn: 3,
      endRow: 3,
      endColumn: 6,
    });
    expect(expandSelectionForMergedRanges(createSelection(2, 4), renderModel)).toMatchObject({
      active: { row: 1, column: 3 },
      range: { startRow: 1, startColumn: 3, endRow: 2, endColumn: 4 },
    });
  });

  test('recomputes selections after row and column deletion', () => {
    const renderModel = buildSpreadsheetSheetRenderModel(
      {
        ...makeSheet(),
        mergedRanges: [{ address: 'C3:D4' }],
      },
      makeStyles(),
    );
    const bounds = { rowCount: 20, columnCount: 8 };

    expect(selectionAfterRowDeletion(createSelection(7, 3), 3, 2, bounds, renderModel)).toEqual({
      kind: 'cell',
      active: { row: 5, column: 3 },
      range: { startRow: 5, startColumn: 3, endRow: 5, endColumn: 3 },
    });
    expect(selectionAfterRowDeletion(createSelection(4, 3), 3, 2, bounds, renderModel)).toEqual({
      kind: 'cell',
      active: { row: 3, column: 3 },
      range: { startRow: 3, startColumn: 3, endRow: 4, endColumn: 4 },
    });
    expect(selectionAfterRowDeletion(createRowSelection(4, { columnCount: 8 }), 3, 2, bounds, renderModel)).toEqual({
      kind: 'row',
      active: { row: 3, column: 1 },
      range: { startRow: 3, startColumn: 1, endRow: 3, endColumn: 8 },
    });
    expect(selectionAfterColumnDeletion(createSelection(3, 7), 3, 2, bounds, renderModel)).toEqual({
      kind: 'cell',
      active: { row: 3, column: 5 },
      range: { startRow: 3, startColumn: 5, endRow: 3, endColumn: 5 },
    });
    expect(selectionAfterColumnDeletion(createSelection(3, 4), 3, 2, bounds, renderModel)).toEqual({
      kind: 'cell',
      active: { row: 3, column: 3 },
      range: { startRow: 3, startColumn: 3, endRow: 4, endColumn: 4 },
    });
    expect(selectionAfterColumnDeletion(createColumnSelection(4, { rowCount: 20 }), 3, 2, bounds, renderModel)).toEqual({
      kind: 'column',
      active: { row: 1, column: 3 },
      range: { startRow: 1, startColumn: 3, endRow: 20, endColumn: 3 },
    });
  });
});

function makeSheet(): ExcelSheetModel {
  return {
    id: '1',
    name: 'Sheet1',
    sheetViews: [
      {
        pane: { xSplit: 1, ySplit: 1, state: 'frozen' },
        selections: [],
      },
    ],
    // Excel stores this as character width; the layout converts it to 120 CSS px.
    columns: [{ min: 2, max: 2, width: 17.1484375, customWidth: true }],
    rows: [{ index: 2, height: 30, customHeight: true }],
    mergedRanges: [{ address: 'C1:D2' }],
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
      { address: 'A1', row: 1, column: 1, kind: 'string', value: 'Name' },
      { address: 'B2', row: 2, column: 2, kind: 'number', value: 42 },
      { address: 'C1', row: 1, column: 3, kind: 'string', value: 'Merged' },
    ],
  };
}

function makeStyles(): ExcelStylesModel {
  return {
    numberFormats: [],
    fonts: [{}],
    fills: [{}, { patternType: 'gray125' }],
    borders: [{}],
    cellFormats: [{}],
  };
}
