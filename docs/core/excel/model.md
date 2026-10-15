# Excel Model Contract

This document defines the normalized model used by `ExcelWorkbook.fromModel`, `workbook.toModel`, package readers/writers, host bridges, and tests.

Use this document when the task says: model, JSON bridge, fixture, normalized workbook, OOXML parse output, host bridge data, or standalone package generation.

## Skill Summary

The model is plain data. The live-object SDK wraps this model and mutates it through APIs. `toModel()` returns defensive snapshots; direct mutation of those snapshots does not mutate the workbook.

Host integrations can use the model as a common exchange format when package bytes are unavailable or when a live host API must be sampled into structured data.

## Workbook Model

| Field | Type | Description |
| --- | --- | --- |
| `sheets` | `ExcelSheetModel[]` | Worksheets in workbook order. |
| `sharedStrings` | `string[]` | Shared string table values. |
| `definedNames` | `ExcelDefinedNameModel[]` | Workbook or sheet-scoped defined names. |
| `styles` | `ExcelStylesModel` | Parsed style tables. |

## Sheet Model

| Field | Type | Description |
| --- | --- | --- |
| `id` | `string` | Sheet id. |
| `name` | `string` | Sheet name. |
| `visibility` | `string?` | Visibility state when present. |
| `path` | `string?` | Worksheet part path when package-backed. |
| `dimension` | `string?` | Used range/dimension such as `A1:D8`. |
| `cells` | `ExcelCellModel[]` | Cell data sorted by row and column. |
| `columns` | `ExcelColumnModel[]` | Column metadata. |
| `rows` | `ExcelRowModel[]` | Row metadata. |
| `mergedRanges` | `ExcelMergedRangeModel[]` | Merge refs. |
| `sheetViews` | `ExcelSheetViewModel[]` | Panes and selections. |
| `hyperlinks` | `ExcelHyperlinkModel[]` | Hyperlinks. |
| `tables` | `ExcelTableModel[]` | Excel structured tables/list objects. |
| `dataValidations` | `ExcelDataValidationModel[]` | Data-validation blocks. |
| `conditionalFormats` | `ExcelConditionalFormatModel[]` | Conditional-format blocks. |
| `comments` | `ExcelCommentModel[]` | Legacy comments. |
| `threadedComments` | `ExcelThreadedCommentModel[]` | Threaded comments. |
| `drawings` | `ExcelDrawingModel[]` | Drawing parts and anchors. |
| `charts` | `ExcelChartModel[]` | Chart parts and formula refs. |
| `images` | `ExcelImageModel[]` | Image/media parts. |

## Cell Model

| Field | Type | Description |
| --- | --- | --- |
| `address` | `string` | A1 cell address. |
| `row` | `number` | 1-based row index. |
| `column` | `number` | 1-based column index. |
| `kind` | `ExcelCellKind` | Value kind. |
| `value` | `string \| number \| boolean \| null` | Normalized cell value. |
| `formula` | `string?` | Formula text without leading `=`. |
| `rawValue` | `string?` | Raw OOXML value. |
| `dataType` | `string?` | Raw OOXML cell type. |
| `styleId` | `number?` | Cell style index. |
| `sharedStringIndex` | `number?` | Shared string index when present. |

`ExcelCellKind`:

| Value | Description |
| --- | --- |
| `blank` | Missing or empty cell. |
| `string` | String value. |
| `number` | Number value. |
| `boolean` | Boolean value. |
| `formula` | Formula cell. |
| `error` | Error cell. |
## Object Model Groups

| Group | Main Interfaces | Documentation |
| --- | --- | --- |
| Tables | `ExcelTableModel`, `ExcelTableColumnModel`, `ExcelTableStyleModel` | [domains/tables.md](./domains/tables.md) |
| Annotations | `ExcelHyperlinkModel`, `ExcelCommentModel`, `ExcelThreadedCommentModel` | [domains/annotations.md](./domains/annotations.md) |
| Formatting rules | `ExcelDataValidationModel`, `ExcelConditionalFormatModel`, `ExcelConditionalFormatRuleModel` | [domains/formatting-rules.md](./domains/formatting-rules.md) |
| Drawings | `ExcelDrawingModel`, `ExcelDrawingAnchorModel`, `ExcelDrawingMarkerModel`, `ExcelChartModel`, `ExcelImageModel`, `ExcelShapeModel` | [domains/drawings.md](./domains/drawings.md) |
| Styles | `ExcelStylesModel`, `ExcelCellFormatModel`, `ExcelFontModel`, `ExcelFillModel`, `ExcelBorderModel`, `ExcelNumberFormatModel` | [domains/workbook.md](./domains/workbook.md) |

## Color Model

`ExcelColorModel` has one public representation:

```ts
{ rgb: '#D9D9D9' }
```

`rgb` is required and must be a leading `#` followed by exactly six hexadecimal RGB digits. SDK callers do not provide OOXML alpha, theme, indexed, tint, or auto color attributes. Imported OOXML colors are resolved to this RGB form, and exports convert it to OOXML ARGB internally.

## Standalone Writer Coverage

`ExcelWorkbook.fromModel(model).toZip()` can create a standalone package for:

| Area | Coverage |
| --- | --- |
| Workbook/sheets | Supported. |
| Cells/values/formulas | Supported. |
| Tables | Supported. |
| Hyperlinks | Supported. |
| Legacy comments | Supported. |
| Data validations | Supported. |
| Conditional formats | Supported. |
| Threaded comments | Not synthesized from scratch. |
| Styles | Not synthesized from scratch. |
| Basic charts | Supported through `sheet.charts.add(...)`; serialized as drawing and chart OOXML parts. |
| Arbitrary drawings/images/shapes | Not synthesized from scratch. |

Package-backed workbooks loaded through `fromZip` or `fromPackage` can preserve and update more OOXML parts because the original package provides relationships and media parts.

## CodeAgent Rules

1. Use the model as bridge data, not as the primary editing API when a live `ExcelWorkbook` is available.
2. Use live objects for generated operations, then call `toModel()` or `toZip()` for output.
3. Keep formulas without a leading `=`.
4. Keep worksheet row/column indexes 1-based and drawing marker indexes 0-based.
5. Do not put host-specific fields in the core model unless they are promoted into the public model contract.
