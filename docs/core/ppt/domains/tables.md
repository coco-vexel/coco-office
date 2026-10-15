# PPT Tables Domain

Use this document when a task mentions PPT tables, cells, rows, columns, cell
style, or table values.

## Skill Summary

PPT tables are shape domains. First get a table shape, then use `shape.table`
and zero-based row/column indexes:

```ts
const table = slide.shapes.getItem('SummaryTable');
await table.table.setCellText(1, 2, 'Done');
await table.table.addRow('end', ['Total', '100']);
```

For unknown shapes, call `shape.table.isTable()` before table operations. Check
merged-cell metadata before structural row/column edits.

## Create Table

```ts
const tableShape = await slide.shapes.addTable(2, 3, {
  name: 'SummaryTable',
  values: [
    ['Metric', 'Q1', 'Q2'],
    ['Revenue', '10', '12'],
  ],
});
```

`values` must match `rowCount x columnCount`.

`PptTableAddOptions` extends `PptShapeAddOptions`:

| Field      | Type                        | Description                                       |
| ---------- | --------------------------- | ------------------------------------------------- |
| `name`     | `string`                    | Optional table shape name.                        |
| `geometry` | `Partial<PptShapeGeometry>` | Optional EMU geometry.                            |
| `values`   | `string[][]`                | Initial text matrix; must match table dimensions. |

## Table APIs

| API                                          | Returns                        | Notes                                                                 |
| -------------------------------------------- | ------------------------------ | --------------------------------------------------------------------- |
| `shape.table.isTable()`                      | `Promise<boolean>`             | True when shape contains a table.                                     |
| `shape.table.rowCount()`                     | `Promise<number>`              | Row count.                                                            |
| `shape.table.columnCount()`                  | `Promise<number>`              | Column count.                                                         |
| `shape.table.values()`                       | `Promise<string[][]>`          | Text matrix.                                                          |
| `shape.table.hasMergedCells()`               | `Promise<boolean>`             | True when merge metadata exists.                                      |
| `shape.table.mergedCells()`                  | `Promise<PptTableCellMerge[]>` | `gridSpan`, `rowSpan`, `hMerge`, `vMerge` snapshot.                   |
| `shape.table.setValues(values)`              | `Promise<void>`                | Matrix dimensions must match.                                         |
| `shape.table.setCellText(row, column, text)` | `Promise<void>`                | Direct cell update.                                                   |
| `shape.table.getCell(row, column)`           | `TableCell`                    | Cell handle.                                                          |
| `shape.table.addRow(index?, values?)`        | `Promise<void>`                | `index` is 0-based or `'end'`; numeric insertion indexes are clamped. |
| `shape.table.deleteRow(index)`               | `Promise<void>`                | Rejects deleting final row.                                           |
| `shape.table.addColumn(index?, values?)`     | `Promise<void>`                | `index` is 0-based or `'end'`; numeric insertion indexes are clamped. |
| `shape.table.deleteColumn(index)`            | `Promise<void>`                | Rejects deleting final column.                                        |

## Cell APIs

```ts
const cell = tableShape.table.getCell(0, 1);
await cell.setText('Updated');
await cell.setStyle({
  fillColor: '#DDEEFF',
  textStyle: { bold: true, color: '#112233', alignment: 'center' },
  body: {
    anchor: 'middle',
    wrap: 'square',
    insetLeft: 91440,
    insetRight: 91440,
    insetTop: 45720,
    insetBottom: 45720,
  },
});
```

| API                    | Returns                       | Notes                                     |
| ---------------------- | ----------------------------- | ----------------------------------------- |
| `cell.text()`          | `Promise<string>`             | Cell text.                                |
| `cell.paragraphs()`    | `Promise<PptTextParagraph[]>` | Paragraph/run structure.                  |
| `cell.runs()`          | `Promise<PptTextRun[]>`       | Flattened runs.                           |
| `cell.setParagraphs()` | `Promise<void>`               | Structured text update.                   |
| `cell.setRunStyle()`   | `Promise<void>`               | Updates one flattened run.                |
| `cell.setText(text)`   | `Promise<void>`               | Cell text update.                         |
| `cell.style()`         | `Promise<PptTableCellStyle>`  | `fillColor`, `fillGradient`, `borders`, `body`, `textStyle`. |
| `cell.setStyle(style)` | `Promise<void>`               | Partial fill/body/text style cell update. |

`PptTableCellStyle`:

| Field       | Type                    | Description                                                                          |
| ----------- | ----------------------- | ------------------------------------------------------------------------------------ |
| `fillColor` | `string`                | Cell fill color as `#RRGGBB` or `RRGGBB`.                                            |
| `fillGradient` | `PptGradientFillModel` | Cell gradient fill parsed from real PPT files.                                      |
| `borders`   | `PptTableCellBorders`   | Read-only snapshot of explicit left/right/top/bottom cell borders.                   |
| `body`      | `PptTextBodyProperties` | Cell text-body anchor, wrap, and insets in EMUs.                                     |
| `textStyle` | `PptTextStyle`          | Cell text style: font, color, emphasis, baseline, highlight, alignment, and spacing. |

Cell `fillGradient` and `borders` can appear in snapshots from real PPT files,
but the P0 public write API targets solid cell fill, cell body properties, and
cell text style. Retrieve [../style-capabilities.md](../style-capabilities.md)
for the full style matrix.

## Row And Column Examples

Append a total row:

```ts
const columns = await tableShape.table.columnCount();
await tableShape.table.addRow(
  'end',
  Array.from({ length: columns }, (_, i) => (i === 0 ? 'Total' : '')),
);
```

Insert a status column:

```ts
const rows = await tableShape.table.rowCount();
await tableShape.table.addColumn(
  'end',
  Array.from({ length: rows }, () => 'OK'),
);
```

## Current Boundary

The current table API targets ordinary tables for structural edits. Merged-cell
tables expose `hasMergedCells()` and `mergedCells()` so SDK operation code can detect
them. Direct text/style edits are allowed for merge owner cells, but continuation
cells (`hMerge` or `vMerge`) and structural operations (`setValues`, row/column
add/delete) are rejected to preserve the OOXML table structure.

## SDK Operation Rules

1. Use 0-based row and column indexes.
2. Check `isTable()` before operating on unknown shapes.
3. Keep `setValues` dimensions equal to current dimensions.
4. Use `addRow`/`addColumn` before setting larger matrices.
5. Do not edit table OOXML directly in SDK operation code.
6. Check `hasMergedCells()` before structural table edits; avoid continuation cells listed by `mergedCells()`.
7. `addRow` and `addColumn` clamp numeric insertion indexes; use current counts when exact placement matters.
