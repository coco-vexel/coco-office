# Workbook APIs

This domain covers workbook loading, exporting, styles, defined names, package access, and lifecycle.

Use this document when the task says: load workbook, read `.xlsx`, create workbook from model, export `.xlsx`, workbook styles, defined names, or package writeback.

## Skill Summary

`ExcelWorkbook` is the public async root object. Core can load an existing workbook from `.xlsx` bytes or from a normalized `ExcelWorkbookModel`. Host environments are responsible for obtaining bytes/model data and for applying exported bytes/model changes back to Microsoft Excel, WPS, ONLYOFFICE, or desktop files.

Generated workbook-operation code should normally assume a ready `workbook: ExcelWorkbook` is passed in by the host/runtime.

For worker isolation, use [../runtime.md](../runtime.md). `ExcelWorkbook` itself supports inline and worker mode, matching the `WordDocument` API shape.

## Load APIs

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `ExcelWorkbook.fromZip(bytes, opts?)` | Load an `.xlsx` zip package from bytes. | `Uint8Array \| ArrayBuffer`, `ExcelWorkbookLoadOptions?` | `Promise<ExcelWorkbook>` |
| `ExcelWorkbook.fromPackage(pkg, opts?)` | Load from an already parsed package. Inline-only. | `Package`, `ExcelWorkbookLoadOptions?` | `ExcelWorkbook` |
| `ExcelWorkbook.fromModel(model, opts?)` | Load from a normalized workbook model. | `ExcelWorkbookModel`, `ExcelWorkbookLoadOptions?` | `Promise<ExcelWorkbook>` |

`ExcelWorkbookLoadOptions`:

| Field | Type | Required | Description |
| --- | --- | --- | --- |
| `workbookId` | `string` | no | Stable workbook id for diagnostics. Defaults to `excel:${Date.now()}`. |
| `mode` | `'auto' \| 'worker' \| 'inline'` | no | Runtime mode. Defaults to global default, initially `auto`. |
| `onModeResolved` | `(decision, workbookId) => void` | no | Receives runtime mode decision. |

## Workbook Properties

| Property | Type | Description |
| --- | --- | --- |
| `workbook.workbookId` | `string` | Workbook id from load options or generated default. |
| `workbook.mode` | `'worker' \| 'inline'` | Actual runtime mode. |
| `workbook.modeSource` | `ExcelModeSource` | Why the mode was chosen. |
| `workbook.worksheets` | `ExcelWorksheetCollection` | Worksheet domain. |
| `workbook.names` | `ExcelNamedItemCollection` | Defined-name domain. |

## Workbook Methods

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `workbook.toModel()` | Return a defensive normalized snapshot. | none | `Promise<ExcelWorkbookModel>` |
| `workbook.getMutations()` | Advanced host-runtime API: list SDK-recorded mutations since the last baseline clear. | none | `Promise<ExcelWorkbookMutation[]>` |
| `workbook.clearMutations()` | Advanced host-runtime API: clear recorded mutations after load or successful host writeback. | none | `Promise<void>` |
| `workbook.toZip()` | Export current workbook as `.xlsx` bytes. | none | `Promise<Uint8Array>` |
| `workbook.dispose()` | Dispose the underlying runtime/worker. | none | `Promise<void>` |

`toModel()` returns a defensive copy. Mutating the returned object does not mutate the workbook.

`getMutations()` and `clearMutations()` are not workbook editing APIs. They support fast host writeback by exposing the operations that were performed through SDK live-object methods. Host runtimes should still compare the before/after workbook models before replaying mutations. See [../mutation-writeback.md](../mutation-writeback.md).

`toZip()` behavior:

| Source | Behavior |
| --- | --- |
| `fromZip` / `fromPackage` | Synchronizes the live model back into the source package and exports bytes. |
| `fromModel` | Creates a standalone `.xlsx` package for supported standalone domains. |

Standalone package generation supports workbook metadata, worksheets, cells, formulas, tables, legacy comments, hyperlinks, data validations, conditional formats, and basic charts created through `sheet.charts.add(...)`. It does not synthesize arbitrary drawings, images, shapes, threaded comments, or styles from scratch.

## Defined Names

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `workbook.names.items()` | List defined names. | none | `Promise<ExcelNamedItem[]>` |
| `workbook.names.getItem(name)` | Get one defined name or throw. | `name: string` | `ExcelNamedItem` |
| `workbook.names.getItemOrNullObject(name)` | Get one defined name or null. | `name: string` | `Promise<ExcelNamedItem \| null>` |
| `workbook.names.add(model)` | Add a defined name. | `ExcelDefinedNameModel` | `Promise<ExcelNamedItem>` |
| `workbook.names.deleteItem(name)` | Delete a defined name. | `name: string` | `Promise<void>` |

`ExcelNamedItem`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `name.name()` | Read name. | none | `Promise<string>` |
| `name.text()` | Read name formula/text. | none | `Promise<string>` |
| `name.setText(text)` | Set name formula/text. | `string` | `Promise<void>` |
| `name.rename(name)` | Rename the defined name. The live handle tracks the new key after success. | `string` | `Promise<void>` |
| `name.localSheetId()` | Read local sheet index when scoped. | none | `Promise<number \| undefined>` |
| `name.setLocalSheetId(localSheetId)` | Set or clear local sheet scope. | `number \| undefined` | `Promise<void>` |

Defined name text is stored as workbook formula text, for example `Sheet1!$A$1:$B$2`. `localSheetId` is a 0-based sheet index following OOXML workbook semantics.

## CodeAgent Rules

1. Use `ExcelWorkbook.fromZip` when the host gives real `.xlsx` bytes.
2. Use `ExcelWorkbook.fromModel` when the host bridge already produced normalized data.
3. Do not use `ExcelWorkbookCore.getPackage()` for normal generated edits.
4. Use `toZip()` only when the caller needs `.xlsx` bytes.
5. Use `sheet.charts.add(...)` for basic chart creation; do not rely on standalone `fromModel(...).toZip()` for arbitrary drawing/image/shape synthesis.
6. Use `workbook.names` for defined-name mutations. Do not mutate `toModel().definedNames` directly.
7. Do not use `getMutations()` or `clearMutations()` in generated workbook-operation code. They are for host runtimes and adapters.
