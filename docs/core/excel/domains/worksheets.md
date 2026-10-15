# Worksheet APIs

This domain covers worksheet discovery, worksheet live objects, sheet metadata, row/column structural edits, and sheet-level domains.

Use this document when the task says: worksheet, sheet, rename sheet, used range, insert row, delete column, merged range, freeze pane, table on sheet, comments on sheet, drawings on sheet, or list sheets.

## Skill Summary

Worksheet APIs are async live-object APIs. `workbook.worksheets.items()` returns `ExcelWorksheet` handles, not raw sheet data. A worksheet handle exposes range/cell operations and domain collections such as `sheet.tables`, `sheet.comments`, and `sheet.drawings`.

Generated code should use worksheet domain properties directly. Do not invent parallel methods such as `sheet.tables()`.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `workbook.worksheets.items()` | List worksheets as live handles. | none | `Promise<ExcelWorksheet[]>` |
| `workbook.worksheets.getItem(nameOrId)` | Get one worksheet by name or id. | `string` | `ExcelWorksheet` |
| `workbook.worksheets.getItemOrNullObject(nameOrId)` | Get one worksheet by name or id, or null. | `string` | `Promise<ExcelWorksheet \| null>` |
| `workbook.worksheets.add(name)` | Add a blank worksheet. | `name: string` | `Promise<ExcelWorksheet>` |
| `workbook.worksheets.copy(nameOrId, newName?)` | Copy a worksheet model and insert it after the source sheet. | `string`, `string?` | `Promise<ExcelWorksheet>` |
| `workbook.worksheets.move(nameOrId, toIndex)` | Move a worksheet to a 0-based workbook index. | `string`, `number` | `Promise<void>` |
| `workbook.worksheets.deleteItem(nameOrId)` | Delete a worksheet. | `string` | `Promise<void>` |

`getItem` throws when no sheet matches. `getItemOrNullObject` returns `null`.

`deleteItem` rejects deleting the final remaining worksheet. `copy` clears package relationship id and worksheet part path on the copied sheet; `toZip()` assigns fresh package parts during serialization.

## Worksheet Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.id()` | Read sheet id. | none | `Promise<string>` |
| `sheet.name()` | Read sheet name. | none | `Promise<string>` |
| `sheet.rename(name)` | Rename the worksheet and rewrite supported local references. | `string` | `Promise<void>` |
| `sheet.visibility()` | Read visibility. | none | `Promise<string \| undefined>` |
| `sheet.setVisibility(visibility)` | Set visibility. Use `undefined` or `visible` for normal visibility, `hidden`, or `veryHidden`. | `string \| undefined` | `Promise<void>` |
| `sheet.path()` | Read source worksheet part path when package-backed. | none | `Promise<string \| undefined>` |
| `sheet.dimension()` | Read worksheet dimension. | none | `Promise<string \| undefined>` |
| `sheet.getRange(address)` | Create a range handle. | `string` such as `A1:D4` | `ExcelRange` |
| `sheet.getCell(address)` | Create a cell handle. | `string` such as `B2` | `ExcelCell` |
| `sheet.getUsedRange()` | Create a range handle for the worksheet dimension. | none | `Promise<ExcelRange \| null>` |
| `sheet.cells()` | Return cell snapshots. | none | `Promise<ExcelCellModel[]>` |
| `sheet.sheetViews()` | Return sheet-view snapshots. | none | `Promise<ExcelSheetViewModel[]>` |
| `sheet.columns()` | Return column metadata snapshots. | none | `Promise<ExcelColumnModel[]>` |
| `sheet.rows()` | Return row metadata snapshots. | none | `Promise<ExcelRowModel[]>` |
| `sheet.mergedRanges()` | Return merged-range snapshots. | none | `Promise<ExcelMergedRangeModel[]>` |
| `sheet.toModel()` | Return a defensive sheet snapshot. | none | `Promise<ExcelSheetModel>` |

## Structural Writes

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.insertRows(rowIndex, count?)` | Insert rows before `rowIndex`. | `rowIndex: number`, `count = 1` | `Promise<void>` |
| `sheet.deleteRows(rowIndex, count?)` | Delete rows starting at `rowIndex`. | `rowIndex: number`, `count = 1` | `Promise<void>` |
| `sheet.insertColumns(columnIndex, count?)` | Insert columns before `columnIndex`. | `columnIndex: number`, `count = 1` | `Promise<void>` |
| `sheet.deleteColumns(columnIndex, count?)` | Delete columns starting at `columnIndex`. | `columnIndex: number`, `count = 1` | `Promise<void>` |

Rows and columns are 1-based. Counts must be positive integers.

Structural edits update supported dependent references:

| Area | Updated |
| --- | --- |
| Cell addresses | yes |
| Cell formulas | best-effort A1 rewrite |
| Defined names | best-effort A1 rewrite |
| Data validation formulas and ranges | best-effort A1 rewrite |
| Conditional format formulas and ranges | best-effort A1 rewrite |
| Table ranges and auto-filter ranges | best-effort A1 rewrite |
| Hyperlink/comment/threaded-comment refs | yes when reference remains valid |
| Sheet-view panes/selections | yes when reference remains valid |
| Drawing anchors | moved by marker index |
| Chart formula refs | best-effort A1 rewrite |

## Sheet-Level Domains

| Property | Type | Documentation |
| --- | --- | --- |
| `sheet.tables` | `ExcelTableCollection` | [tables.md](./tables.md) |
| `sheet.hyperlinks` | `ExcelHyperlinkCollection` | [annotations.md](./annotations.md) |
| `sheet.comments` | `ExcelCommentCollection` | [annotations.md](./annotations.md) |
| `sheet.threadedComments` | `ExcelThreadedCommentCollection` | [annotations.md](./annotations.md) |
| `sheet.dataValidations` | `ExcelDataValidationCollection` | [formatting-rules.md](./formatting-rules.md) |
| `sheet.conditionalFormats` | `ExcelConditionalFormatCollection` | [formatting-rules.md](./formatting-rules.md) |
| `sheet.drawings` | `ExcelDrawingCollection` | [drawings.md](./drawings.md) |
| `sheet.charts` | `ExcelChartCollection` | [drawings.md](./drawings.md) |
| `sheet.images` | `ExcelImageCollection` | [drawings.md](./drawings.md) |
| `sheet.shapes` | `ExcelShapeCollection` | [drawings.md](./drawings.md) |

## CodeAgent Rules

1. Treat `sheet` as a worksheet live object, not as an Excel table.
2. Use `sheet.tables` only for Excel structured tables/list objects.
3. Use 1-based row and column indexes for worksheet structural APIs.
4. Use `sheet.getRange(...)` for matrix operations and `sheet.getCell(...)` for scalar operations.
5. After structural edits, call `workbook.toZip()` only if the caller needs package output.
6. Use `workbook.worksheets.add/copy/move/deleteItem` for sheet lifecycle. Do not edit `toModel().sheets` directly.
