# Table APIs

This domain covers table discovery, table live objects, cells, rows, columns, and text ranges inside cells.

Use this document when the task says: table, row, column, cell, merged cell, table text, insert row, delete row, cell background, cell vertical alignment, cell width, row height, column width, nested table, table picture, or text in a cell.

## Skill Summary

Table APIs use live objects:

```ts
const tables = await doc.tables.list({ includeCellTexts: true });
const table = await doc.tables.get('tbl-0001');
const cell = table?.cell(0, 1);
```

Cells, rows, and columns are addressed with zero-based indexes. A cell is a small content container: use `cell.getParagraphs()` and `cell.getBlocks()` when you need to operate on content inside the cell.

To insert a new simple text table, use a paragraph anchor:

```ts
await paragraph.insertTableAfter({
  rows: [
    ['Name', 'Score'],
    ['Ada', '42'],
  ],
});
```

For high-fidelity table copying, use `table.getVerbatim()` and `doc.plan().insertTableFrom(...)`.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.tables.list(opts?)` | List tables as live handles. | `ListTablesOptions` | `Promise<Table[]>` |
| `doc.tables.get(tableId)` | Get one table by id. | `tableId: string` | `Promise<Table \| null>` |
| `doc.tables.getVerbatim(tableId)` | Export raw table OOXML and images. | `tableId: string` | `Promise<TableVerbatim \| null>` |

## ListTablesOptions

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `scope` | `'body' \| 'all'` | `'body'` | Read body tables only or all story tables. |
| `includeNested` | `boolean` | `false` | Include tables nested inside other tables. |
| `includeCellTexts` | `boolean` | `false` | Include `cells` and `cellGrid`; required for cell text analysis. |

## TableInfo

`Table` extends `TableInfo`, so listed tables carry these snapshot fields.

| Field | Type | Description |
| --- | --- | --- |
| `tableId` / `id` | `string` | Stable table id. |
| `partPath` | `string` | OOXML part path such as `word/document.xml`. |
| `index` | `number` | 1-based order in the list result. |
| `isNested` | `boolean` | Whether the table is nested inside another table. |
| `rowCount` | `number` | Direct row count. |
| `colCount` | `number` | Table grid column count. |
| `hasMerge` | `boolean` | Has horizontal or vertical merged cells. |
| `tableStyleId` | `string?` | Direct table style id. |
| `hasImages` | `boolean` | Table contains drawings. |
| `hasNestedTables` | `boolean` | Table contains direct nested tables. |
| `prevParaId` | `string?` | Previous sibling paragraph id, if any. |
| `nextParaId` | `string?` | Next sibling paragraph id, if any. |
| `columnWidths` | `number[]` | Grid column widths in twips. |
| `rowHeights` | `(RowHeight \| null)[]` | Row heights in document order. |
| `cells` | `string[][]?` | Cell text grid, only with `includeCellTexts:true`. |
| `cellGrid` | `TableCellInfo[][]?` | Rich cell grid, only with `includeCellTexts:true`. |

`RowHeight`:

| Field | Type | Description |
| --- | --- | --- |
| `val` | `number` | Row height in twips. |
| `hRule` | `'auto' \| 'atLeast' \| 'exact'` | Height rule. |

## Table Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `table.cell(row, col)` | Create a cell handle. | `row: number`, `col: number` | `TableCell` |
| `table.row(row)` | Create a row handle. | `row: number` | `TableRow` |
| `table.column(col)` | Create a column handle. | `col: number` | `TableColumn` |
| `table.getInfo(opts?)` | Read current table metadata. | `ListTablesOptions?` | `Promise<TableInfo \| null>` |
| `table.getVerbatim()` | Export this table OOXML and images. | none | `Promise<TableVerbatim \| null>` |
| `table.getText(opts?)` | Read table plain text. | `ListTablesOptions?` | `Promise<string \| null>` |
| `table.getCells(opts?)` | Read cell text grid. | `ListTablesOptions?` | `Promise<string[][] \| null>` |
| `table.addComment(text, opts?)` | Add a comment to the whole table. | `text: string`, `LiveObjectWriteOptions?` such as `{author:'halvee'}` | `Promise<ApplyResult>` |
| `table.delete(opts?)` | Delete this table. | `TableDeleteOptions?` | `Promise<ApplyResult>` |
| `table.replaceCellText(row, col, text, opts?)` | Replace one cell's text. | `row`, `col`, `text` | `Promise<ApplyResult>` |
| `table.deleteRow(row, opts?)` | Delete one row. | `row: number` | `Promise<ApplyResult>` |
| `table.insertRow(spec, opts?)` | Insert one row. | `{templateRow:number; position:number \| 'end'; cellTexts?:string[]}` | `Promise<ApplyResult>` |
| `table.setColumnWidths(widths, opts?)` | Set grid column widths. | `number[]` twips | `Promise<ApplyResult>` |
| `table.setRowHeights(heights, opts?)` | Set row heights. | `(RowHeight \| null)[]` | `Promise<ApplyResult>` |

`TableDeleteOptions`, `TableCellWriteOptions`, and `TableRowDeleteOptions` support:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `trackChanges` | `boolean` | `false` | Request tracked-change markup. |
| `onAnchorConflict` | `'error' \| 'drop'` | `'error'` | How to handle affected comments/bookmarks. |

## TableCellInfo

| Field | Type | Description |
| --- | --- | --- |
| `text` | `string` | Cell plain text. Multiple paragraphs join with `\n`. |
| `gridSpan` | `number` | Horizontal span. Default is `1`. |
| `vMerge` | `'restart' \| 'continue' \| null` | Vertical merge state. |
| `width` | `number?` | Direct cell width in twips. |
| `shading` | `{fill?:string; color?:string; val?:string}?` | Direct cell shading. |
| `verticalAlignment` | `'top' \| 'center' \| 'bottom' \| 'both'?` | Direct vertical alignment. |
| `firstParaId` | `string?` | First direct paragraph id in the cell. |
| `paragraphIds` | `string[]` | Direct paragraph ids in cell order. |
| `blocks` | `TableCellBlockInfo[]` | Direct content blocks in OOXML order. |
| `imageDrawingIds` | `string[]` | Direct drawing ids inside this cell. |

`TableCellBlockInfo`:

| Shape | Description |
| --- | --- |
| `{type:'paragraph'; paraId; text}` | A direct paragraph. |
| `{type:'drawing'; drawingId; paraId?}` | A drawing in a direct paragraph. |
| `{type:'table'; tableId}` | A nested direct table. |

## TableCell Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `cell.getInfo(opts?)` | Read current cell metadata. | `ListTablesOptions?` | `Promise<TableCellInfo \| null>` |
| `cell.getBlocks(opts?)` | Read live content blocks in order. | `ListTablesOptions?` | `Promise<TableCellBlock[] \| null>` |
| `cell.getParagraphs(opts?)` | Get direct paragraph handles. | `ListTablesOptions?` | `Promise<Paragraph[]>` |
| `cell.getText(opts?)` | Read current cell text. | `ListTablesOptions?` | `Promise<string \| null>` |
| `cell.findText(search, opts?)` | Create a cell text range. | `search: string`, `{scope?:'first'\|'all'}` | `CellTextRange` |
| `cell.replaceWith(text, opts?)` | Replace cell content with plain text. | `text: string` | `Promise<ApplyResult>` |
| `cell.insertParagraphAtStart(text, opts?)` | Insert a paragraph at cell start. | `text: string` | `Promise<ApplyResult>` |
| `cell.insertParagraphAtEnd(text, opts?)` | Insert a paragraph at cell end. | `text: string` | `Promise<ApplyResult>` |
| `cell.setShading(fill, opts?)` | Set direct background fill. | `fill: string` | `Promise<ApplyResult>` |
| `cell.setVerticalAlignment(value, opts?)` | Set vertical alignment. | `'top' \| 'center' \| 'bottom' \| 'both'` | `Promise<ApplyResult>` |
| `cell.setWidth(width, opts?)` | Set direct cell width. | `width: number` twips | `Promise<ApplyResult>` |

`TableCellBlock` returned by `cell.getBlocks()` contains live handles:

| Shape | Description |
| --- | --- |
| `{type:'paragraph'; paragraph; paraId; text; index}` | Direct paragraph block. |
| `{type:'drawing'; drawing; drawingId; paraId?; index}` | Drawing block. `drawing` may be `null` if it no longer exists. |
| `{type:'table'; table; tableId; index}` | Nested table block. `table` may be `null` if it no longer exists. |

## Row And Column Objects

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `row.getCells(opts?)` | Read row cell texts. | `ListTablesOptions?` | `Promise<string[] \| null>` |
| `row.getText(opts?)` | Read row text joined by tabs. | `ListTablesOptions?` | `Promise<string \| null>` |
| `row.delete(opts?)` | Delete this row. | `TableRowDeleteOptions?` | `Promise<ApplyResult>` |
| `row.insertBefore(cellTexts?, opts?)` | Insert row before this row. | `cellTexts?: string[]` | `Promise<ApplyResult>` |
| `row.insertAfter(cellTexts?, opts?)` | Insert row after this row. | `cellTexts?: string[]` | `Promise<ApplyResult>` |
| `row.setHeight(height, opts?)` | Set row height; `null` removes explicit height. | `RowHeight \| null` | `Promise<ApplyResult>` |
| `column.getCells(opts?)` | Read column cell texts. | `ListTablesOptions?` | `Promise<string[] \| null>` |
| `column.getText(opts?)` | Read column text joined by line breaks. | `ListTablesOptions?` | `Promise<string \| null>` |
| `column.setWidth(width, opts?)` | Set this column width. | `width: number` twips | `Promise<ApplyResult>` |

## TableVerbatim

`getVerbatim` is for high-fidelity copy/move flows.

| Field | Type | Description |
| --- | --- | --- |
| `tableXml` | `string` | Serialized table OOXML subtree. |
| `images` | `TableVerbatimImage[]` | Images referenced by the table. |

`TableVerbatimImage`: `{ originalRId: string; bytes: Uint8Array; contentType: string }`.

## Examples

Find a table containing a keyword:

```ts
const tables = await doc.tables.list({ includeCellTexts: true });
const target = tables.find((table) => table.cells?.some((row) => row.some((cell) => cell.includes('金额'))));
```

Update one cell:

```ts
const table = await doc.tables.get('tbl-0001');
if (table) {
  await table.cell(1, 2).replaceWith('已确认');
}
```

Insert a row after the first row:

```ts
const table = await doc.tables.get('tbl-0001');
await table?.row(0).insertAfter(['序号', '事项', '状态']);
```

Preserve cell content order:

```ts
const table = await doc.tables.get('tbl-0001');
const blocks = await table?.cell(0, 0).getBlocks();
for (const block of blocks ?? []) {
  if (block.type === 'paragraph') {
    console.log(await block.paragraph.getText());
  }
}
```

Insert a simple table after a known paragraph:

```ts
const anchor = await doc.paragraphs.get('31086A56');
await anchor?.insertTableAfter({
  rows: [
    ['Item', 'Owner'],
    ['Budget', 'Finance'],
  ],
  columnWidths: [2200, 2200],
});
```

## CodeAgent Rules

1. Use zero-based row and column indexes.
2. Use `includeCellTexts:true` when selecting a table by cell content.
3. Use `cell.getParagraphs()` or `cell.getBlocks()` for content-aware operations inside a cell.
4. Use `cell.findText()` for exact text operations inside a cell.
5. Do not treat `cells` or `cellGrid` as live objects; they are snapshot data.
6. Insert a new simple table through `paragraph.insertTableAfter(...)`; use `getVerbatim` plus `insertTableFrom` only for high-fidelity copy/move.
