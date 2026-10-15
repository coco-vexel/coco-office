import {
  columnIndexToName,
  columnNameToIndex,
  formatCellAddress,
  parseRangeAddress,
} from './address.js';
import type { ExcelCellModel, ExcelCellValue, ExcelWorkbookModel } from './model.js';

export type ExcelAggregateOperation = 'sum' | 'count' | 'average' | 'min' | 'max';
export type ExcelAggregateBucket = 'none' | 'year' | 'quarter' | 'month' | 'day';

export interface ExcelAggregateSource {
  sheetName: string;
  address?: string;
  tableName?: string;
  header?: 'firstRow' | 'none';
}

export interface ExcelAggregateGroupBy {
  column: string;
  bucket?: ExcelAggregateBucket;
  name?: string;
}

export interface ExcelAggregateMetric {
  column?: string;
  op: ExcelAggregateOperation;
  name?: string;
}

export interface ExcelAggregateTableOptions {
  source: ExcelAggregateSource;
  groupBy: ExcelAggregateGroupBy[];
  metrics: ExcelAggregateMetric[];
  sort?: 'asc' | 'desc' | 'none';
}

export interface ExcelAggregateResultRow {
  key: ExcelCellValue[];
  values: ExcelCellValue[];
  rowCount: number;
}

export interface ExcelAggregateTableResult {
  ok: true;
  sheetName: string;
  sourceAddress: string;
  headers: string[];
  rows: ExcelAggregateResultRow[];
  values: ExcelCellValue[][];
  inputRowCount: number;
  outputRowCount: number;
}

interface SourceTable {
  sheetName: string;
  address: string;
  headers: string[];
  rows: Array<{
    values: ExcelCellValue[];
    cells: Array<ExcelCellModel | undefined>;
  }>;
  startColumn: number;
}

interface GroupBucketValue {
  value: ExcelCellValue;
  sortKey: string;
}

interface MetricAccumulator {
  op: ExcelAggregateOperation;
  sum: number;
  count: number;
  min: number | null;
  max: number | null;
}

export function aggregateExcelTable(
  workbook: ExcelWorkbookModel,
  options: ExcelAggregateTableOptions,
): ExcelAggregateTableResult {
  if (options.groupBy.length === 0) {
    throw new Error('aggregateExcelTable: groupBy must contain at least one column');
  }
  if (options.metrics.length === 0) {
    throw new Error('aggregateExcelTable: metrics must contain at least one metric');
  }

  const source = readSourceTable(workbook, options.source);
  const groups = new Map<
    string,
    {
      key: ExcelCellValue[];
      sortKeys: string[];
      rowCount: number;
      metrics: MetricAccumulator[];
    }
  >();

  const groupColumns = options.groupBy.map((group) => resolveColumn(source, group.column));
  const metricColumns = options.metrics.map((metric) =>
    metric.column ? resolveColumn(source, metric.column) : null,
  );

  for (const row of source.rows) {
    const groupValues = groupColumns.map((columnIndex, index) =>
      bucketValue(row.values[columnIndex] ?? null, options.groupBy[index]?.bucket ?? 'none'),
    );
    const keyText = groupValues.map((item) => stableKey(item.value)).join('\u001f');
    let group = groups.get(keyText);
    if (!group) {
      group = {
        key: groupValues.map((item) => item.value),
        sortKeys: groupValues.map((item) => item.sortKey),
        rowCount: 0,
        metrics: options.metrics.map((metric) => ({
          op: metric.op,
          sum: 0,
          count: 0,
          min: null,
          max: null,
        })),
      };
      groups.set(keyText, group);
    }
    group.rowCount += 1;

    options.metrics.forEach((metric, index) => {
      const acc = group!.metrics[index]!;
      const value =
        metric.op === 'count' && !metric.column
          ? 1
          : metricColumns[index] === null
            ? null
            : row.values[metricColumns[index]!] ?? null;
      accumulate(acc, value);
    });
  }

  const sortedGroups = Array.from(groups.values());
  if ((options.sort ?? 'asc') !== 'none') {
    const direction = options.sort === 'desc' ? -1 : 1;
    sortedGroups.sort((a, b) => direction * compareSortKeys(a.sortKeys, b.sortKeys));
  }

  const rows = sortedGroups.map((group) => ({
    key: group.key,
    values: group.metrics.map(finalizeMetric),
    rowCount: group.rowCount,
  }));
  const headers = [
    ...options.groupBy.map((group) => group.name ?? group.column),
    ...options.metrics.map((metric) => metric.name ?? defaultMetricName(metric)),
  ];

  return {
    ok: true,
    sheetName: source.sheetName,
    sourceAddress: source.address,
    headers,
    rows,
    values: [headers, ...rows.map((row) => [...row.key, ...row.values])],
    inputRowCount: source.rows.length,
    outputRowCount: rows.length,
  };
}

function readSourceTable(workbook: ExcelWorkbookModel, source: ExcelAggregateSource): SourceTable {
  const sheet = workbook.sheets.find(
    (item) => item.name === source.sheetName || item.id === source.sheetName,
  );
  if (!sheet) throw new Error(`aggregateExcelTable: worksheet not found: ${source.sheetName}`);

  const table = source.tableName
    ? sheet.tables.find(
        (item) => item.name === source.tableName || item.displayName === source.tableName,
      )
    : undefined;
  const address = source.address ?? table?.ref ?? sheet.dimension;
  if (!address) throw new Error(`aggregateExcelTable: source range is empty: ${source.sheetName}`);

  const parsed = parseRangeAddress(address);
  const byAddress = new Map(sheet.cells.map((cell) => [cell.address, cell]));
  const matrix: Array<Array<ExcelCellModel | undefined>> = [];
  for (let r = 0; r < parsed.rowCount; r += 1) {
    const row: Array<ExcelCellModel | undefined> = [];
    for (let c = 0; c < parsed.columnCount; c += 1) {
      row.push(byAddress.get(formatCellAddress(parsed.startRow + r, parsed.startColumn + c)));
    }
    matrix.push(row);
  }

  const hasHeader = source.header !== 'none';
  const headerCells = hasHeader ? matrix[0] ?? [] : [];
  const headers = Array.from({ length: parsed.columnCount }, (_, index) => {
    const value = headerCells[index]?.value;
    if (hasHeader && value !== null && value !== undefined && String(value).trim().length > 0) {
      return String(value);
    }
    return columnIndexToName(parsed.startColumn + index);
  });
  const dataRows = hasHeader ? matrix.slice(1) : matrix;

  return {
    sheetName: sheet.name,
    address: parsed.address,
    headers,
    rows: dataRows.map((cells) => ({
      cells,
      values: cells.map((cell) => cell?.value ?? null),
    })),
    startColumn: parsed.startColumn,
  };
}

function resolveColumn(source: SourceTable, column: string): number {
  const normalized = column.trim();
  const headerIndex = source.headers.findIndex((header) => header === normalized);
  if (headerIndex >= 0) return headerIndex;
  if (/^[A-Za-z]+$/.test(normalized)) {
    const absoluteIndex = columnNameToIndex(normalized.toUpperCase());
    const offset = absoluteIndex - source.startColumn;
    if (offset >= 0 && offset < source.headers.length) return offset;
  }
  throw new Error(`aggregateExcelTable: column not found: ${column}`);
}

function bucketValue(value: ExcelCellValue, bucket: ExcelAggregateBucket): GroupBucketValue {
  if (bucket === 'none') return { value, sortKey: stableKey(value) };
  const date = coerceExcelDate(value);
  if (!date) {
    return { value: value === null ? null : String(value), sortKey: stableKey(value) };
  }
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  if (bucket === 'year') return { value: String(year), sortKey: `${year}` };
  if (bucket === 'quarter') {
    const quarter = Math.floor((month - 1) / 3) + 1;
    return { value: `${year}Q${quarter}`, sortKey: `${year}-${quarter}` };
  }
  if (bucket === 'month') {
    return {
      value: `${year}-${String(month).padStart(2, '0')}`,
      sortKey: `${year}-${String(month).padStart(2, '0')}`,
    };
  }
  return {
    value: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
    sortKey: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
  };
}

function coerceExcelDate(value: ExcelCellValue): Date | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    const epoch = Date.UTC(1899, 11, 30);
    return new Date(epoch + Math.round(value * 86_400_000));
  }
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  const match = /^(\d{4})[-/年](\d{1,2})(?:[-/月](\d{1,2}))?/.exec(trimmed);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = match[3] ? Number(match[3]) : 1;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return new Date(Date.UTC(year, month - 1, day));
}

function accumulate(acc: MetricAccumulator, value: ExcelCellValue): void {
  if (acc.op === 'count') {
    if (value !== null && value !== '') acc.count += 1;
    return;
  }
  const numeric = coerceNumber(value);
  if (numeric === null) return;
  acc.sum += numeric;
  acc.count += 1;
  acc.min = acc.min === null ? numeric : Math.min(acc.min, numeric);
  acc.max = acc.max === null ? numeric : Math.max(acc.max, numeric);
}

function finalizeMetric(acc: MetricAccumulator): ExcelCellValue {
  if (acc.op === 'count') return acc.count;
  if (acc.count === 0) return null;
  if (acc.op === 'sum') return acc.sum;
  if (acc.op === 'average') return acc.sum / acc.count;
  if (acc.op === 'min') return acc.min;
  return acc.max;
}

function coerceNumber(value: ExcelCellValue): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value !== 'string') return null;
  const normalized = value.replace(/,/g, '').trim();
  if (normalized.length === 0) return null;
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function defaultMetricName(metric: ExcelAggregateMetric): string {
  return metric.column ? `${metric.op}_${metric.column}` : metric.op;
}

function stableKey(value: ExcelCellValue): string {
  if (value === null) return '';
  return `${typeof value}:${String(value)}`;
}

function compareSortKeys(left: string[], right: string[]): number {
  const count = Math.max(left.length, right.length);
  for (let index = 0; index < count; index += 1) {
    const result = (left[index] ?? '').localeCompare(right[index] ?? '', undefined, {
      numeric: true,
    });
    if (result !== 0) return result;
  }
  return 0;
}
