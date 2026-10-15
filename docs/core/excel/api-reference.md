# Core Excel API Reference

This file is the public API index. Detailed parameters, return values, model fields, host notes, and code-agent rules live in the domain documents.

Import:

```ts
import { ExcelWorkbook } from '@coco-vexel/coco-office/excel';
import { planExcelWriteback } from '@coco-vexel/coco-office/excel';
import type { ExcelWorkbookModel, ExcelSheetModel, ExcelWorkbookMutation } from '@coco-vexel/coco-office/excel';
```

`ExcelWorkbook` is the public root object. It is async by design and can run against the same core kernel inline or in a worker, matching the `WordDocument` runtime shape.

## API Contract

The Excel SDK exposes one public live-object surface. The same user code must work whether the runtime is inline or worker-backed.

| Rule | Contract |
| --- | --- |
| Public root | Use `ExcelWorkbook` from `@coco-vexel/coco-office/excel`. |
| Runtime mode | Select runtime with `mode`, `setDefaultMode`, `setWorkerFactory`, or `setWorkerBundleUrl`; do not branch business logic by mode. |
| Live handles | `getItem(...)`, `getRange(...)`, `getCell(...)`, `getRow(...)`, `getColumn(...)`, `resize(...)`, `offset(...)`, and `slice(...)` return lightweight live handles synchronously. |
| Reads and writes | Methods that read, mutate, serialize, or check existence return `Promise`. Always `await` them. |
| Snapshot boundary | `toModel()` returns a defensive snapshot. Mutating that object does not mutate the workbook. |
| Mutation boundary | `getMutations()` and `clearMutations()` expose SDK-recorded edit operations for host writeback optimization. They are adapter/runtime APIs, not normal workbook editing APIs. |
| Package boundary | `toZip()` is the package serialization boundary for host writeback and file output. |
| Low-level core | `*Core` classes are for kernel internals, tests, and advanced adapters. Generated SDK code should not use them. |
| Script escape hatch | `executeCode(...)` is available for runtime-side bulk scripts, but the default generated path should use live-object APIs. |

### Handle Versus Operation

Live handles are intentionally cheap and synchronous. They identify workbook locations or objects, but they do not fetch data by themselves:

```ts
const sheet = workbook.worksheets.getItem('Sheet1'); // sync handle
const range = sheet.getRange('A1:B2'); // sync handle
const cell = range.getCell(0, 0); // sync handle

const values = await range.values(); // async read
await cell.setValue('done'); // async write
```

If a method name starts with `getItem` and does not include `OrNullObject`, it returns a handle and may throw only when a later operation resolves it. Use `getItemOrNullObject(...)` when existence must be checked before operating.

## Primary Object Model

`ExcelWorkbook` is the root object. Domain APIs hang from `workbook` and `worksheet`:

| Domain Property | Main Types | Full Documentation |
| --- | --- | --- |
| `workbook.worksheets` | `ExcelWorksheetCollection`, `ExcelWorksheet` | [domains/worksheets.md](./domains/worksheets.md) |
| `workbook.names` | `ExcelNamedItemCollection`, `ExcelNamedItem` | [domains/workbook.md](./domains/workbook.md) |
| `worksheet.getRange(...)` | `ExcelRange` | [domains/ranges-cells.md](./domains/ranges-cells.md) |
| `worksheet.getCell(...)` | `ExcelCell` | [domains/ranges-cells.md](./domains/ranges-cells.md) |
| `worksheet.tables` | `ExcelTableCollection`, `ExcelTable`, `ExcelTableColumnCollection` | [domains/tables.md](./domains/tables.md) |
| `worksheet.hyperlinks` | `ExcelHyperlinkCollection`, `ExcelHyperlink` | [domains/annotations.md](./domains/annotations.md) |
| `worksheet.comments` | `ExcelCommentCollection`, `ExcelComment` | [domains/annotations.md](./domains/annotations.md) |
| `worksheet.threadedComments` | `ExcelThreadedCommentCollection`, `ExcelThreadedComment` | [domains/annotations.md](./domains/annotations.md) |
| `worksheet.dataValidations` | `ExcelDataValidationCollection`, `ExcelDataValidation` | [domains/formatting-rules.md](./domains/formatting-rules.md) |
| `worksheet.conditionalFormats` | `ExcelConditionalFormatCollection`, `ExcelConditionalFormat` | [domains/formatting-rules.md](./domains/formatting-rules.md) |
| `worksheet.drawings` | `ExcelDrawingCollection`, `ExcelDrawing`, `ExcelDrawingAnchor` | [domains/drawings.md](./domains/drawings.md) |
| `worksheet.charts` | `ExcelChartCollection`, `ExcelChart`, `ExcelFormulaRefCollection` | [domains/drawings.md](./domains/drawings.md) |
| `worksheet.images` | `ExcelImageCollection`, `ExcelImage` | [domains/drawings.md](./domains/drawings.md) |
| `worksheet.shapes` | `ExcelShapeCollection`, `ExcelShape` | [domains/drawings.md](./domains/drawings.md) |

## Workbook Lifecycle

| API | Documentation |
| --- | --- |
| `ExcelWorkbook.fromZip(bytes, opts?)` | [domains/workbook.md](./domains/workbook.md) |
| `ExcelWorkbook.fromPackage(pkg, opts?)` | [domains/workbook.md](./domains/workbook.md) |
| `ExcelWorkbook.fromModel(model, opts?)` | [domains/workbook.md](./domains/workbook.md), [model.md](./model.md) |
| `workbook.toModel()` | [model.md](./model.md) |
| `workbook.toZip()` | [domains/workbook.md](./domains/workbook.md), [host-integration.md](./host-integration.md) |
| `workbook.dispose()` | [runtime.md](./runtime.md) |

## Common Public APIs

### Workbook

| API | Returns | Notes |
| --- | --- | --- |
| `ExcelWorkbook.fromZip(bytes, opts?)` | `Promise<ExcelWorkbook>` | Preferred when a host can provide `.xlsx` bytes. Supports worker mode. |
| `ExcelWorkbook.fromModel(model, opts?)` | `Promise<ExcelWorkbook>` | Preferred when a host bridge already normalized sheet/model data. Supports worker mode. |
| `ExcelWorkbook.fromPackage(pkg, opts?)` | `ExcelWorkbook` | Inline only. Use only when a parsed package object is already available. |
| `workbook.worksheets.items()` | `Promise<ExcelWorksheet[]>` | Lists worksheet handles. |
| `workbook.worksheets.getItem(nameOrId)` | `ExcelWorksheet` | Returns a worksheet handle. |
| `workbook.worksheets.getItemOrNullObject(nameOrId)` | `Promise<ExcelWorksheet \| null>` | Checks worksheet existence. |
| `workbook.worksheets.add(name)` | `Promise<ExcelWorksheet>` | Adds a blank worksheet. |
| `workbook.worksheets.copy(nameOrId, newName?)` | `Promise<ExcelWorksheet>` | Copies a worksheet model and inserts it after the source. |
| `workbook.worksheets.move(nameOrId, toIndex)` | `Promise<void>` | Moves a worksheet to a 0-based workbook index. |
| `workbook.worksheets.deleteItem(nameOrId)` | `Promise<void>` | Deletes a worksheet; deleting the final sheet is rejected. |
| `workbook.names.items()` | `Promise<ExcelNamedItem[]>` | Lists defined-name handles. |
| `workbook.names.add(model)` | `Promise<ExcelNamedItem>` | Adds a defined name. |
| `workbook.names.deleteItem(name)` | `Promise<void>` | Deletes a defined name. |
| `workbook.toModel()` | `Promise<ExcelWorkbookModel>` | Returns a defensive model snapshot. |
| `workbook.aggregateTable(options)` | `Promise<ExcelAggregateTableResult>` | Deterministically groups rows and calculates `sum`, `count`, `average`, `min`, or `max`; supports date buckets such as `month`, `quarter`, `year`, and `day`. |
| `workbook.getMutations()` | `Promise<ExcelWorkbookMutation[]>` | Advanced host-runtime API. Lists SDK-recorded mutations since the last `clearMutations()` call. See [mutation-writeback.md](./mutation-writeback.md). |
| `workbook.clearMutations()` | `Promise<void>` | Advanced host-runtime API. Clears the mutation log after a host bridge has accepted the current workbook baseline. |
| `workbook.toZip()` | `Promise<Uint8Array>` | Exports an `.xlsx` package. |
| `workbook.dispose()` | `Promise<void>` | Releases worker resources when present. |

### Analysis

Use `workbook.aggregateTable(...)` when generated code or an agent must compute grouped statistics. Do not rely on an LLM to hand-calculate grouped totals.

```ts
const monthly = await workbook.aggregateTable({
  source: { sheetName: '销售订单', tableName: 'OrderTable' },
  groupBy: [{ column: '签订日期', bucket: 'month', name: '月份' }],
  metrics: [{ column: '合同金额', op: 'sum', name: '合同金额' }],
});

await workbook.worksheets
  .getItem('统计分析')
  .getRange('A1:B5')
  .setValues(monthly.values);
```

`source` may point to an Excel table (`tableName`), an explicit A1 range (`address`), or the worksheet used range. `groupBy[].column` and `metrics[].column` accept either a header value from the first row or an absolute Excel column letter such as `B`. Count metrics may omit `column` to count rows.

### Writeback Planner

| API | Returns | Notes |
| --- | --- | --- |
| `diffExcelWorkbookModels(before, after)` | `ExcelWorkbookDiff` | Classifies changed sheets and includes executable `ExcelCellDiff[]` payloads for cell additions, deletions, and modifications. See [writeback-planner.md](./writeback-planner.md). |
| `planExcelWriteback(before, after, capabilities, options?)` | `ExcelWritebackPlan` | Chooses `no-op`, `workbook-package-replace`, `sheet-shadow-replace`, `diff-patch`, or `unsupported` and returns portable operations for host adapters. Sheet-shadow operations include target sheet payloads; `auto` mode treats diff patch as fallback rather than the primary host writeback path. |
| `createShadowWorkbookModel(sheet, options?)` | `ExcelWorkbookModel` | Advanced host-adapter helper. Builds a one-sheet workbook model from a target sheet payload. |
| `createShadowWorkbookPayload(sheet, options?)` | `Promise<ExcelShadowWorkbookPayload>` | Advanced host-adapter helper. Builds a one-sheet workbook package, `.xlsx` bytes, and base64 for sheet-shadow writeback. |

### Mutation Writeback

Mutation APIs exist so a host adapter can avoid whole-workbook or whole-sheet copy when the SDK can prove a small edit is safe to replay directly in the live host.

| API / Type | Returns | Notes |
| --- | --- | --- |
| `workbook.getMutations()` | `Promise<ExcelWorkbookMutation[]>` | Lists recorded mutations such as `range.setValues`, `range.setFormulas`, `range.clear`, `range.merge`, `range.unmerge`, `worksheet.add`, `worksheet.delete`, and `worksheet.rename`. |
| `workbook.clearMutations()` | `Promise<void>` | Clears the log. Runtimes should call this after loading a baseline and after a successful host writeback. |
| `ExcelWorkbookMutation` | union type | Plain JSON mutation payload used by host runtimes. It can also contain `unsafe` records that force fallback writeback. |
| `changedSheetNamesFromMutations(mutations)` | `string[]` | Helper for diagnostics and adapter planning. |

Generated workbook-operation code should not call mutation APIs. Host runtimes may call them around a user task:

```ts
const baseline = await workbook.toModel();
await workbook.clearMutations();

await workbook.worksheets.getItem('Sheet1').getRange('C2:D7').setValues(values);

const after = await workbook.toModel();
const mutations = await workbook.getMutations();
// Runtime verifies baseline/after/mutations before host replay.
```

### Worksheet

| API | Returns | Notes |
| --- | --- | --- |
| `sheet.id()` | `Promise<string>` | Stable internal sheet id. |
| `sheet.name()` | `Promise<string>` | Display sheet name. |
| `sheet.rename(name)` | `Promise<void>` | Renames the sheet and rewrites affected references supported by core. |
| `sheet.visibility()` | `Promise<string \| undefined>` | Sheet visibility state from workbook metadata. |
| `sheet.setVisibility(visibility)` | `Promise<void>` | Sets `visible`, `hidden`, `veryHidden`, or clears visibility with `undefined`. |
| `sheet.dimension()` | `Promise<string \| undefined>` | Source sheet dimension when available. |
| `sheet.getRange(address)` | `ExcelRange` | Returns a range handle. |
| `sheet.getCell(address)` | `ExcelCell` | Returns a single-cell handle. |
| `sheet.getUsedRange()` | `Promise<ExcelRange \| null>` | Returns a used-range handle when data exists. |
| `sheet.cells()` | `Promise<ExcelCellModel[]>` | Snapshot of cell models. |
| `sheet.insertRows(rowIndex, count?)` | `Promise<void>` | Inserts rows and rewrites supported references. |
| `sheet.deleteRows(rowIndex, count?)` | `Promise<void>` | Deletes rows and rewrites supported references. |
| `sheet.insertColumns(columnIndex, count?)` | `Promise<void>` | Inserts columns and rewrites supported references. |
| `sheet.deleteColumns(columnIndex, count?)` | `Promise<void>` | Deletes columns and rewrites supported references. |
| `sheet.toModel()` | `Promise<ExcelSheetModel>` | Defensive sheet snapshot. |

### Range And Cell

Range addresses use Excel A1 notation. `sheet.getRange(...)` accepts ordinary cell ranges such as `A1:D4`, whole-column references such as `A:A` and `A:C`, and whole-row references such as `1:1` and `1:7`.

| API | Returns | Notes |
| --- | --- | --- |
| `range.address()` | `Promise<string>` | A1-style address. |
| `range.values()` | `Promise<ExcelCellValue[][]>` | Matrix of typed values. |
| `range.formulas()` | `Promise<(string \| null)[][]>` | Matrix of formulas without replacing values. |
| `range.text()` | `Promise<string[][]>` | Display text matrix. |
| `range.valueTypes()` | `Promise<ExcelCellKind[][]>` | Cell kind matrix. |
| `range.styleIds()` | `Promise<(number \| undefined)[][]>` | OOXML style id matrix. |
| `range.setValues(values)` | `Promise<void>` | Writes values over the range. Matrix size must match the range shape. |
| `range.setFormulas(formulas)` | `Promise<void>` | Writes formulas over the range. Matrix size must match the range shape. |
| `range.setStyleId(styleId)` | `Promise<void>` | Sets or clears OOXML style id on all cells in the range. |
| `range.rowHeights()` | `Promise<Array<number \| undefined>>` | Reads row heights intersecting the range. |
| `range.setRowHeight(height)` | `Promise<void>` | Sets or clears row height. |
| `range.columnWidths()` | `Promise<Array<number \| undefined>>` | Reads column widths intersecting the range. |
| `range.setColumnWidth(width)` | `Promise<void>` | Sets or clears column width. |
| `range.merge()` | `Promise<void>` | Adds this exact range to merged ranges. |
| `range.unmerge()` | `Promise<void>` | Removes this exact range from merged ranges. |
| `range.clear()` | `Promise<void>` | Clears cell data in the range. |
| `range.format.numberFormatCode()` | `Promise<string \| undefined>` | Reads first resolved number format code. |
| `range.format.setNumberFormatCode(code)` | `Promise<void>` | Sets or clears number format code. |
| `range.format.font.*` | mixed `Promise` APIs | Reads/writes font name, size, bold, italic, underline, and color. |
| `range.format.fill.*` | mixed `Promise` APIs | Reads/writes fill pattern and foreground/background colors. |
| `range.format.border.*` | mixed `Promise` APIs | Reads/writes full border or individual sides. |
| `range.format.alignment.*` | mixed `Promise` APIs | Reads/writes horizontal, vertical, and wrap text alignment. |
| `range.getCell(rowOffset, columnOffset)` | `ExcelCell` | Returns a cell handle inside the range. |
| `range.getRow(rowOffset)` | `ExcelRange` | Returns a derived row handle. |
| `range.getColumn(columnOffset)` | `ExcelRange` | Returns a derived column handle. |
| `range.resize(rowCount, columnCount)` | `ExcelRange` | Returns a resized range handle. |
| `range.offset(rowOffset, columnOffset, rowCount?, columnCount?)` | `ExcelRange` | Returns an offset range handle. |
| `range.slice(rowOffset, columnOffset, rowCount, columnCount)` | `ExcelRange` | Returns a subrange handle. |
| `cell.value()` | `Promise<ExcelCellValue>` | Single-cell typed value. |
| `cell.formula()` | `Promise<string \| null>` | Single-cell formula. |
| `cell.styleId()` | `Promise<number \| undefined>` | Single-cell OOXML style id. |
| `cell.setValue(value)` | `Promise<void>` | Writes one value. |
| `cell.setFormula(formula)` | `Promise<void>` | Writes one formula or clears formula with `null`. |
| `cell.setStyleId(styleId)` | `Promise<void>` | Sets or clears single-cell OOXML style id. |

### Sheet Object Domains

Every worksheet object domain follows the same live-object collection pattern:

| Domain | Collection APIs | Item APIs |
| --- | --- | --- |
| Tables | `sheet.tables.items()`, `getItem(name)`, `getItemOrNullObject(name)`, `add(model)` | `name()`, `displayName()`, `setDisplayName(...)`, `rangeAddress()`, `resize(...)`, `autoFilter()`, `sortState()`, `columns`, `style` |
| Hyperlinks | `sheet.hyperlinks.items()`, `getItem(ref)`, `getItemOrNullObject(ref)`, `add(model)` | `ref()`, `setRef(...)`, `target()`, `display()`, `setDisplay(...)`, `location()` |
| Comments | `sheet.comments.items()`, `getItem(ref)`, `getItemOrNullObject(ref)`, `add(model)` | `ref()`, `setRef(...)`, `text()`, `setText(...)` |
| Threaded comments | `sheet.threadedComments.items()`, `getItem(ref)`, `getItemOrNullObject(ref)`, `add(model)` | `ref()`, `setRef(...)`, `text()`, `setText(...)` |
| Data validations | `sheet.dataValidations.items()`, `getItem(range)`, `getItemOrNullObject(range)`, `add(model)` | `rangeAddress()`, `setRange(...)`, `type()`, `operator()`, formula and prompt setters |
| Conditional formats | `sheet.conditionalFormats.items()`, `getItem(range)`, `getItemOrNullObject(range)`, `add(model)` | `rangeAddress()`, `setRange(...)`, `rules` |
| Drawings | `sheet.drawings.items()`, `getItem(id)`, `getItemOrNullObject(id)` | object metadata and anchors |
| Charts | `sheet.charts.items()`, `getItem(id)`, `getItemOrNullObject(id)`, `add(options)` | title, refs, anchor, chart model; basic chart creation |
| Images | `sheet.images.items()`, `getItem(id)`, `getItemOrNullObject(id)` | media relationship and anchor metadata |
| Shapes | `sheet.shapes.items()`, `getItem(id)`, `getItemOrNullObject(id)` | shape metadata and anchor metadata |

For detailed parameter fields, use the matching domain document rather than inferring model shape from examples.

## Runtime Configuration

| API | Documentation |
| --- | --- |
| `ExcelWorkbook.setDefaultMode(mode)` | [runtime.md](./runtime.md) |
| `ExcelWorkbook.setWorkerFactory(factory)` | [runtime.md](./runtime.md) |
| `ExcelWorkbook.setWorkerBundleUrl(url)` | [runtime.md](./runtime.md) |
| `ExcelWorkbook.detectMode()` | [runtime.md](./runtime.md) |
| `workbook.executeCode(code, args?, timeoutMs?, signal?)` | [runtime.md](./runtime.md) |

## Advanced And Low-Level Exports

| Area | Examples |
| --- | --- |
| Address helpers | `parseCellAddress`, `parseRangeAddress`, `formatCellAddress`, `formatRangeAddress`, `columnNameToIndex`, `columnIndexToName` |
| Formula/reference rewrite helpers | `renameSheetInFormula`, `rewriteFormulaForRowInsert`, `rewriteFormulaForRowDelete`, `rewriteFormulaForColumnInsert`, `rewriteFormulaForColumnDelete`, `rewriteReferenceForRowInsert`, `rewriteReferenceForRowDelete`, `rewriteReferenceForColumnInsert`, `rewriteReferenceForColumnDelete` |
| Package readers | `readWorkbookModel` |
| Core-kernel classes | `ExcelWorkbookCore`, `ExcelWorksheetCore`, `ExcelRangeCore`, `ExcelCellCore`, `ExcelTableCore`, `ExcelCommentCore`, `ExcelDrawingCore`, and other `*Core` object exports |
| Runtime worker entries | `@coco-vexel/coco-office/excel-worker`, `@coco-vexel/coco-office/excel-worker-browser` |
| Model types | `ExcelWorkbookModel`, `ExcelSheetModel`, `ExcelCellModel`, object and style model interfaces |
| Mutation writeback types | `ExcelWorkbookMutation`, `ExcelRangeSetValuesMutation`, `ExcelRangeSetFormulasMutation`, `ExcelUnsafeMutation`, and related mutation interfaces |

For ordinary generated workbook-operation code, prefer live-object APIs documented in `domains/`.

## Code Generation Rules

These rules are normative for agents generating Excel SDK code:

1. Import public operation APIs from `@coco-vexel/coco-office/excel`.
2. Use `ExcelWorkbook`, `ExcelWorksheet`, `ExcelRange`, `ExcelCell`, and other non-`Core` live-object types for generated operations.
3. Do not import `ExcelWorkbookCore` or any `*Core` class unless the task explicitly asks for kernel tests, parser internals, or an adapter implementation.
4. Treat all reads, writes, serialization, and existence checks as async.
5. Do not mutate objects returned by `toModel()` and expect the workbook to change.
6. Do not assume host APIs exist in core code. Office.js, WPS JS, ONLYOFFICE, and desktop bridges must obtain source bytes/models and perform writeback outside core.
7. Prefer `range.setValues(...)` and `range.setFormulas(...)` for rectangular cell edits.
8. Prefer object-domain APIs for tables, comments, hyperlinks, validations, conditional formats, drawings, charts, images, and shapes.
9. Use `executeCode(...)` only when a bulk operation needs to run as one runtime-side transaction or when a host adapter deliberately wants a script boundary.
10. In worker mode, arguments and return values must be structured-cloneable. Never return live handles from `executeCode(...)`.
11. Do not call `getMutations()` or `clearMutations()` from generated workbook-operation code. They are for host runtimes that verify and execute writeback.

## CodeAgent Retrieval Map

| User Intent | Retrieve |
| --- | --- |
| "load xlsx", "export xlsx", "read workbook names", "styles" | `workbook.md` |
| "worker mode", "avoid UI freeze", "execute generated Excel code", "runtime mode" | `runtime.md` |
| "sheet", "rename sheet", "insert rows", "delete columns" | `worksheets.md` |
| "cell", "range", "values", "formula", "clear" | `ranges-cells.md` |
| "Excel table", "structured table", "table columns" | `tables.md` |
| "hyperlink", "comment", "threaded comment" | `annotations.md` |
| "data validation", "conditional formatting" | `formatting-rules.md` |
| "chart", "image", "shape", "drawing anchor" | `drawings.md` |
| "host bridge", "Office/WPS/ONLYOFFICE writeback", "desktop" | `host-integration.md` |
| "mutation replay", "fast writeback", "recorded SDK operations" | `mutation-writeback.md` |
| "model bridge", "JSON model", "fixtures" | `model.md` |
