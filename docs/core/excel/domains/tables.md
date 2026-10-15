# Table APIs

This domain covers Excel structured tables/list objects and their columns.

Use this document when the task says: Excel table, list object, structured table, table column, table range, auto filter range, or table style.

## Skill Summary

`sheet.tables` represents Excel structured tables inside a worksheet. A worksheet itself is not a table. Public table APIs are async live-object APIs and work in both inline and worker mode.

## Table Collection

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.tables.items()` | List tables as live handles. | none | `Promise<ExcelTable[]>` |
| `sheet.tables.getItem(nameOrDisplayName)` | Get a table handle by internal name or display name. | `string` | `ExcelTable` |
| `sheet.tables.getItemOrNullObject(nameOrDisplayName)` | Get a table handle or null. | `string` | `Promise<ExcelTable \| null>` |
| `sheet.tables.add(model)` | Add a table from model data. | `ExcelTableModel` | `Promise<ExcelTable>` |
| `sheet.tables.delete(nameOrDisplayName)` | Delete a table when present. | `string` | `Promise<void>` |
| `sheet.tables.toModel()` | Return table snapshots. | none | `Promise<ExcelTableModel[]>` |

## Table Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `table.id()` | Read table id. | none | `Promise<number \| undefined>` |
| `table.name()` | Read internal table name. | none | `Promise<string \| undefined>` |
| `table.displayName()` | Read visible table name. | none | `Promise<string \| undefined>` |
| `table.setName(name)` | Set internal table name. | `string` | `Promise<void>` |
| `table.setDisplayName(displayName)` | Set visible table name. | `string` | `Promise<void>` |
| `table.rangeAddress()` | Read table range. | none | `Promise<string \| undefined>` |
| `table.setRange(address)` | Set table range and auto-filter range when present. | `string` | `Promise<void>` |
| `table.resize(address)` | Set table range and adjust table column metadata count to match the new range width. | `string` | `Promise<void>` |
| `table.autoFilterRangeAddress()` | Read auto-filter range. | none | `Promise<string \| undefined>` |
| `table.setAutoFilterRange(address)` | Set or clear auto-filter range. | `string \| undefined` | `Promise<void>` |
| `table.headerRowCount()` | Read header row count. | none | `Promise<number \| undefined>` |
| `table.setHeaderRowCount(count)` | Set or clear header row count. | `number \| undefined` | `Promise<void>` |
| `table.totalsRowCount()` | Read totals row count. | none | `Promise<number \| undefined>` |
| `table.setTotalsRowCount(count)` | Set or clear totals row count. | `number \| undefined` | `Promise<void>` |
| `table.totalsRowShown()` | Read totals-row shown flag. | none | `Promise<boolean \| undefined>` |
| `table.setTotalsRowShown(value)` | Set or clear totals-row shown flag. | `boolean \| undefined` | `Promise<void>` |
| `table.autoFilter()` | Read auto-filter metadata snapshot. | none | `Promise<ExcelTableAutoFilterModel \| undefined>` |
| `table.setAutoFilter(model)` | Set or clear auto-filter metadata. | `ExcelTableAutoFilterModel \| undefined` | `Promise<void>` |
| `table.sortState()` | Read table sort-state metadata snapshot. | none | `Promise<ExcelTableSortStateModel \| undefined>` |
| `table.setSortState(sortState)` | Set or clear table sort-state metadata inside auto-filter. | `ExcelTableSortStateModel \| undefined` | `Promise<void>` |
| `table.toModel()` | Return a table snapshot. | none | `Promise<ExcelTableModel>` |
| `table.delete()` | Remove this table from the sheet. | none | `Promise<void>` |

`resize(address)` updates table metadata only. It does not insert/delete worksheet cells. When the new width is larger than `table.columns`, placeholder columns named `ColumnN` are added. When the new width is smaller, trailing table column metadata is removed.

Filter and sort APIs preserve OOXML metadata. They do not execute filtering or sorting against worksheet cell values.

## Table Columns

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `table.columns.items()` | List table columns. | none | `Promise<ExcelTableColumn[]>` |
| `table.columns.getItem(nameOrId)` | Get a column handle by name or id. | `string \| number` | `ExcelTableColumn` |
| `table.columns.add(model)` | Add a column from model data. | `ExcelTableColumnModel` | `Promise<ExcelTableColumn>` |
| `table.columns.delete(nameOrId)` | Delete a column when present. | `string \| number` | `Promise<void>` |
| `table.columns.deleteItem(nameOrId)` | Delete a column or throw when missing. | `string \| number` | `Promise<void>` |
| `table.columns.toModel()` | Return column snapshots. | none | `Promise<ExcelTableColumnModel[]>` |

`ExcelTableColumn`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `column.id()` | Read column id. | none | `Promise<number>` |
| `column.name()` | Read column name. | none | `Promise<string>` |
| `column.setName(name)` | Set column name. | `string` | `Promise<void>` |
| `column.totalsRowLabel()` | Read totals-row label. | none | `Promise<string \| undefined>` |
| `column.setTotalsRowLabel(label)` | Set or clear totals-row label. | `string \| undefined` | `Promise<void>` |
| `column.totalsRowFunction()` | Read totals-row function. | none | `Promise<string \| undefined>` |
| `column.setTotalsRowFunction(fn)` | Set or clear totals-row function. | `string \| undefined` | `Promise<void>` |
| `column.toModel()` | Return a column snapshot. | none | `Promise<ExcelTableColumnModel>` |
| `column.delete()` | Remove this column metadata from the table. | none | `Promise<void>` |

## Table Style

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `table.style.name()` | Read table style name. | none | `Promise<string \| undefined>` |
| `table.style.setName(name)` | Set or clear style name. | `string \| undefined` | `Promise<void>` |
| `table.style.showFirstColumn()` | Read first-column flag. | none | `Promise<boolean \| undefined>` |
| `table.style.setShowFirstColumn(value)` | Set or clear first-column flag. | `boolean \| undefined` | `Promise<void>` |
| `table.style.showLastColumn()` | Read last-column flag. | none | `Promise<boolean \| undefined>` |
| `table.style.setShowLastColumn(value)` | Set or clear last-column flag. | `boolean \| undefined` | `Promise<void>` |
| `table.style.showRowStripes()` | Read row-stripes flag. | none | `Promise<boolean \| undefined>` |
| `table.style.setShowRowStripes(value)` | Set or clear row-stripes flag. | `boolean \| undefined` | `Promise<void>` |
| `table.style.showColumnStripes()` | Read column-stripes flag. | none | `Promise<boolean \| undefined>` |
| `table.style.setShowColumnStripes(value)` | Set or clear column-stripes flag. | `boolean \| undefined` | `Promise<void>` |
| `table.style.toModel()` | Return a style snapshot. | none | `Promise<ExcelTableStyleModel \| undefined>` |

## CodeAgent Rules

1. Use `sheet.tables` only for Excel structured tables/list objects.
2. Use `await table.rangeAddress()` before `sheet.getRange(...)`.
3. `table.columns.delete(...)` removes table column metadata; it does not delete worksheet columns.
4. Use `sheet.deleteColumns(...)` when the task asks to delete actual worksheet columns.
5. Do not call `sheet.tables()`; `tables` is a property.
6. Use `table.resize(...)` for table metadata range changes; use worksheet/range APIs separately for cell data shape changes.
7. Treat `autoFilter` and `sortState` as metadata preservation/edit APIs, not calculation APIs.
