import { describe, expect, test } from 'vitest';
import { aggregateExcelTable, ExcelWorkbook, type ExcelWorkbookModel } from '../index.js';

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
      name: '销售订单',
      dimension: 'A1:D7',
      sheetViews: [],
      columns: [],
      rows: [],
      mergedRanges: [],
      hyperlinks: [],
      tables: [
        {
          id: 1,
          name: 'OrderTable',
          displayName: 'OrderTable',
          ref: 'A1:D7',
          headerRowCount: 1,
          columns: [
            { id: 1, name: '合同ID' },
            { id: 2, name: '签订日期' },
            { id: 3, name: '销售区域' },
            { id: 4, name: '合同金额' },
          ],
        },
      ],
      dataValidations: [],
      conditionalFormats: [],
      comments: [],
      threadedComments: [],
      drawings: [],
      charts: [],
      images: [],
      cells: [
        { address: 'A1', row: 1, column: 1, kind: 'string', value: '合同ID' },
        { address: 'B1', row: 1, column: 2, kind: 'string', value: '签订日期' },
        { address: 'C1', row: 1, column: 3, kind: 'string', value: '销售区域' },
        { address: 'D1', row: 1, column: 4, kind: 'string', value: '合同金额' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'O-001' },
        { address: 'B2', row: 2, column: 2, kind: 'string', value: '2023-03-12' },
        { address: 'C2', row: 2, column: 3, kind: 'string', value: '东部' },
        { address: 'D2', row: 2, column: 4, kind: 'number', value: 150000 },
        { address: 'A3', row: 3, column: 1, kind: 'string', value: 'O-002' },
        { address: 'B3', row: 3, column: 2, kind: 'string', value: '2023/04/08' },
        { address: 'C3', row: 3, column: 3, kind: 'string', value: '北部' },
        { address: 'D3', row: 3, column: 4, kind: 'number', value: 89000 },
        { address: 'A4', row: 4, column: 1, kind: 'string', value: 'O-003' },
        { address: 'B4', row: 4, column: 2, kind: 'string', value: '2023年5月10日' },
        { address: 'C4', row: 4, column: 3, kind: 'string', value: '南部' },
        { address: 'D4', row: 4, column: 4, kind: 'number', value: 215000 },
        { address: 'A5', row: 5, column: 1, kind: 'string', value: 'O-004' },
        { address: 'B5', row: 5, column: 2, kind: 'number', value: 45056 },
        { address: 'C5', row: 5, column: 3, kind: 'string', value: '东部' },
        { address: 'D5', row: 5, column: 4, kind: 'number', value: 92000 },
        { address: 'A6', row: 6, column: 1, kind: 'string', value: 'O-005' },
        { address: 'B6', row: 6, column: 2, kind: 'string', value: '2023-06-21' },
        { address: 'C6', row: 6, column: 3, kind: 'string', value: '西部' },
        { address: 'D6', row: 6, column: 4, kind: 'number', value: 75000 },
        { address: 'A7', row: 7, column: 1, kind: 'string', value: 'O-006' },
        { address: 'B7', row: 7, column: 2, kind: 'string', value: '2023-06-30' },
        { address: 'C7', row: 7, column: 3, kind: 'string', value: '中部' },
        { address: 'D7', row: 7, column: 4, kind: 'string', value: '132,000' },
      ],
    },
  ],
};

describe('aggregateExcelTable', () => {
  test('groups rows by month and sums numeric metrics deterministically', () => {
    const result = aggregateExcelTable(MODEL, {
      source: { sheetName: '销售订单', tableName: 'OrderTable' },
      groupBy: [{ column: '签订日期', bucket: 'month', name: '月份' }],
      metrics: [{ column: '合同金额', op: 'sum', name: '合同金额' }],
    });

    expect(result).toMatchObject({
      ok: true,
      sheetName: '销售订单',
      sourceAddress: 'A1:D7',
      headers: ['月份', '合同金额'],
      inputRowCount: 6,
      outputRowCount: 4,
    });
    expect(result.values).toEqual([
      ['月份', '合同金额'],
      ['2023-03', 150000],
      ['2023-04', 89000],
      ['2023-05', 307000],
      ['2023-06', 207000],
    ]);
  });

  test('groups by text columns and supports count and average', () => {
    const result = aggregateExcelTable(MODEL, {
      source: { sheetName: '销售订单', address: 'A1:D7' },
      groupBy: [{ column: '销售区域' }],
      metrics: [
        { column: '合同金额', op: 'sum', name: '合同金额' },
        { op: 'count', name: '订单数' },
        { column: '合同金额', op: 'average', name: '平均合同金额' },
      ],
      sort: 'none',
    });

    expect(result.values).toEqual([
      ['销售区域', '合同金额', '订单数', '平均合同金额'],
      ['东部', 242000, 2, 121000],
      ['北部', 89000, 1, 89000],
      ['南部', 215000, 1, 215000],
      ['西部', 75000, 1, 75000],
      ['中部', 132000, 1, 132000],
    ]);
  });

  test('is available through the async workbook facade', async () => {
    const workbook = await ExcelWorkbook.fromModel(MODEL, { mode: 'inline' });

    const result = await workbook.aggregateTable({
      source: { sheetName: '销售订单' },
      groupBy: [{ column: 'C', name: '区域' }],
      metrics: [{ column: 'D', op: 'max', name: '最大合同金额' }],
      sort: 'none',
    });

    expect(result.values).toEqual([
      ['区域', '最大合同金额'],
      ['东部', 150000],
      ['北部', 89000],
      ['南部', 215000],
      ['西部', 75000],
      ['中部', 132000],
    ]);
  });
});
