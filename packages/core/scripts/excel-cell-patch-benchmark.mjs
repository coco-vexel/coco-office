import { performance } from 'node:perf_hooks';
import { ExcelWorkbookCore } from '../dist/excel.js';

const rows = positiveInteger(process.env.COCO_EXCEL_BENCH_ROWS, 10_000);
const columns = positiveInteger(process.env.COCO_EXCEL_BENCH_COLUMNS, 10);
const cells = [];
for (let row = 1; row <= rows; row += 1) {
  for (let column = 1; column <= columns; column += 1) {
    cells.push({
      address: cellAddress(row, column),
      row,
      column,
      kind: 'number',
      value: row * columns + column,
    });
  }
}

let generated = ExcelWorkbookCore.fromModel({
  sharedStrings: [],
  definedNames: [],
  styles: {
    numberFormats: [],
    fonts: [{}],
    fills: [{}, { patternType: 'gray125' }],
    borders: [{}],
    cellFormats: [{}],
  },
  sheets: [
    {
      id: '1',
      name: 'Benchmark',
      dimension: `A1:${cellAddress(rows, columns)}`,
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
      cells,
    },
  ],
});
const sourceBytes = await generated.toZip();
// eslint-disable-next-line no-useless-assignment -- release the generated model before measuring retained heap.
generated = undefined;
cells.length = 0;
forceGc();
const beforeLoadHeap = process.memoryUsage().heapUsed;
const loadStarted = performance.now();
const workbook = await ExcelWorkbookCore.fromZip(sourceBytes, { workbookId: 'fidelity-benchmark' });
const loadMs = performance.now() - loadStarted;
forceGc();
const afterLoadHeap = process.memoryUsage().heapUsed;

workbook.worksheets
  .getItem('Benchmark')
  .getRange(cellAddress(rows, columns))
  .setValues([[42]]);
const exportStarted = performance.now();
const outputBytes = await workbook.toZip();
const exportMs = performance.now() - exportStarted;
forceGc();
const afterExportHeap = process.memoryUsage().heapUsed;

// eslint-disable-next-line no-console -- benchmark output is intentionally machine-readable.
console.log(
  JSON.stringify(
    {
      rows,
      columns,
      cellCount: rows * columns,
      sourceBytes: sourceBytes.length,
      outputBytes: outputBytes.length,
      loadMs: round(loadMs),
      singleCellExportMs: round(exportMs),
      loadHeapDeltaMiB: mib(afterLoadHeap - beforeLoadHeap),
      exportHeapDeltaMiB: mib(afterExportHeap - afterLoadHeap),
      heapUsedMiB: mib(afterExportHeap),
    },
    null,
    2,
  ),
);

function positiveInteger(raw, fallback) {
  const value = Number(raw ?? fallback);
  if (!Number.isInteger(value) || value < 1)
    throw new Error(`expected positive integer, received ${raw}`);
  return value;
}

function cellAddress(row, column) {
  let name = '';
  for (let value = column; value > 0; value = Math.floor((value - 1) / 26)) {
    name = String.fromCharCode(65 + ((value - 1) % 26)) + name;
  }
  return `${name}${row}`;
}

function round(value) {
  return Math.round(value * 100) / 100;
}

function mib(bytes) {
  return round(bytes / 1024 / 1024);
}

function forceGc() {
  if (typeof globalThis.gc !== 'function') {
    throw new Error('benchmark requires node --expose-gc');
  }
  globalThis.gc();
}
