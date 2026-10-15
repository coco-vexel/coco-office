# Range And Cell APIs

This domain covers cell and range reads, writes, navigation, clearing, formulas, and value types.

Use this document when the task says: cell, range, A1, values, formulas, text, set cell, clear range, row inside range, column inside range, offset, resize, or matrix write.

## Skill Summary

`ExcelRange` and `ExcelCell` are async live-object APIs. A range points to an address on a worksheet and reads or writes the current workbook model through the selected `ExcelWorkbook` runtime. A cell is a one-cell range convenience wrapper.

Rows and columns inside range addresses use Excel A1 notation. Range offsets are 0-based relative to the range.

Supported A1 range forms include ordinary cell ranges such as `A1:D4`, whole-column references such as `A:A` and `A:C`, and whole-row references such as `1:1` and `1:7`. Whole-column references normalize to row `1` through Excel row `1048576`; whole-row references normalize to column `A` through `XFD`.

## Range Object

| API                                                              | Purpose                                                       | Params                                   | Returns                               |
| ---------------------------------------------------------------- | ------------------------------------------------------------- | ---------------------------------------- | ------------------------------------- |
| `range.address()`                                                | Read normalized range address.                                | none                                     | `Promise<string>`                     |
| `range.rowIndex()`                                               | Read first row index.                                         | none                                     | `Promise<number>`                     |
| `range.columnIndex()`                                            | Read first column index.                                      | none                                     | `Promise<number>`                     |
| `range.rowCount()`                                               | Read row count.                                               | none                                     | `Promise<number>`                     |
| `range.columnCount()`                                            | Read column count.                                            | none                                     | `Promise<number>`                     |
| `range.values()`                                                 | Read cell values as matrix.                                   | none                                     | `Promise<ExcelCellValue[][]>`         |
| `range.formulas()`                                               | Read formulas as matrix.                                      | none                                     | `Promise<(string \| null)[][]>`       |
| `range.text()`                                                   | Read display text approximation as matrix.                    | none                                     | `Promise<string[][]>`                 |
| `range.valueTypes()`                                             | Read value kinds as matrix.                                   | none                                     | `Promise<ExcelCellKind[][]>`          |
| `range.rawValues()`                                              | Read raw OOXML values as matrix.                              | none                                     | `Promise<(string \| null)[][]>`       |
| `range.styleIds()`                                               | Read OOXML style ids as matrix.                               | none                                     | `Promise<(number \| undefined)[][]>`  |
| `range.setValues(values)`                                        | Write values to every cell in the range.                      | `ExcelCellValue[][]`                     | `Promise<void>`                       |
| `range.setFormulas(formulas)`                                    | Write formulas to every cell in the range.                    | `(string \| null)[][]`                   | `Promise<void>`                       |
| `range.setStyleId(styleId)`                                      | Set or clear the OOXML style id on every cell in the range.   | `number \| undefined`                    | `Promise<void>`                       |
| `range.rowHeights()`                                             | Read row heights for rows intersecting the range.             | none                                     | `Promise<Array<number \| undefined>>` |
| `range.setRowHeight(height)`                                     | Set or clear row height for rows intersecting the range.      | `number \| undefined`                    | `Promise<void>`                       |
| `range.columnWidths()`                                           | Read column widths for columns intersecting the range.        | none                                     | `Promise<Array<number \| undefined>>` |
| `range.setColumnWidth(width)`                                    | Set or clear column width for columns intersecting the range. | `number \| undefined`                    | `Promise<void>`                       |
| `range.merge()`                                                  | Add this range to worksheet merged ranges.                    | none                                     | `Promise<void>`                       |
| `range.unmerge()`                                                | Remove this exact range from worksheet merged ranges.         | none                                     | `Promise<void>`                       |
| `range.clear()`                                                  | Clear cells in the range.                                     | none                                     | `Promise<void>`                       |
| `range.getCell(rowOffset, columnOffset)`                         | Get one cell by 0-based offset.                               | `number`, `number`                       | `ExcelCell`                           |
| `range.getRow(rowOffset)`                                        | Get one row inside the range.                                 | `number`                                 | `ExcelRange`                          |
| `range.getColumn(columnOffset)`                                  | Get one column inside the range.                              | `number`                                 | `ExcelRange`                          |
| `range.resize(rowCount, columnCount)`                            | Return a range with the same top-left cell and new size.      | `number`, `number`                       | `ExcelRange`                          |
| `range.offset(rowOffset, columnOffset, rowCount?, columnCount?)` | Return an offset range.                                       | `number`, `number`, `number?`, `number?` | `ExcelRange`                          |
| `range.slice(rowOffset, columnOffset, rowCount, columnCount)`    | Return a bounded subrange.                                    | `number`, `number`, `number`, `number`   | `ExcelRange`                          |
| `range.toModel()`                                                | Return cell snapshots inside the range.                       | none                                     | `Promise<ExcelCellModel[]>`           |

`setValues` and `setFormulas` require matrix dimensions to exactly match `range.rowCount()` and `range.columnCount()`.

Formula text is stored without a leading `=`. `setFormulas([['=A1+B1']])` and `setFormulas([['A1+B1']])` both store `A1+B1`.

`styleIds()` and `setStyleId()` operate on OOXML style indexes, not on expanded font/fill/border objects. Use workbook styles or host adapters to interpret style ids. `setRowHeight()` and `setColumnWidth()` write row/column metadata and set the corresponding custom flags when a numeric value is provided.

## Range Format

`range.format` is the fine-grained formatting live object. It mutates workbook styles and assigns the resulting style id to every cell in the range.

Colors use one canonical SDK representation everywhere: `{ rgb: "#RRGGBB" }`. The leading `#` and exactly six hexadecimal digits are required. ARGB strings such as `"FFD9D9D9"`, bare hex strings such as `"D9D9D9"`, and OOXML `theme`/`indexed` color objects are not accepted by SDK write APIs. OOXML conversion is handled internally.

| API                                                | Purpose                                                                               | Params                              | Returns                                     |
| -------------------------------------------------- | ------------------------------------------------------------------------------------- | ----------------------------------- | ------------------------------------------- |
| `range.format.numberFormatCode()`                  | Read number format code from the first resolved cell style.                           | none                                | `Promise<string \| undefined>`              |
| `range.format.setNumberFormatCode(code)`           | Set or clear number format code. Built-in codes reuse built-in OOXML ids.             | `string \| undefined`               | `Promise<void>`                             |
| `range.format.font.value()`                        | Read font model from the first resolved cell style.                                   | none                                | `Promise<ExcelFontModel>`                   |
| `range.format.font.setValue(font)`                 | Set full font model.                                                                  | `ExcelFontModel`                    | `Promise<void>`                             |
| `range.format.font.bold()`                         | Read bold.                                                                            | none                                | `Promise<boolean \| undefined>`             |
| `range.format.font.setBold(value)`                 | Set or clear bold.                                                                    | `boolean \| undefined`              | `Promise<void>`                             |
| `range.format.font.italic()`                       | Read italic.                                                                          | none                                | `Promise<boolean \| undefined>`             |
| `range.format.font.setItalic(value)`               | Set or clear italic.                                                                  | `boolean \| undefined`              | `Promise<void>`                             |
| `range.format.font.underline()`                    | Read underline.                                                                       | none                                | `Promise<boolean \| undefined>`             |
| `range.format.font.setUnderline(value)`            | Set or clear underline.                                                               | `boolean \| undefined`              | `Promise<void>`                             |
| `range.format.font.color()`                        | Read font color.                                                                      | none                                | `Promise<ExcelColorModel \| undefined>`     |
| `range.format.font.setColor(color)`                | Set or clear font color.                                                              | `ExcelColorModel \| undefined`      | `Promise<void>`                             |
| `range.format.font.name()`                         | Read font name.                                                                       | none                                | `Promise<string \| undefined>`              |
| `range.format.font.setName(name)`                  | Set or clear font name.                                                               | `string \| undefined`               | `Promise<void>`                             |
| `range.format.font.size()`                         | Read font size.                                                                       | none                                | `Promise<number \| undefined>`              |
| `range.format.font.setSize(size)`                  | Set or clear font size.                                                               | `number \| undefined`               | `Promise<void>`                             |
| `range.format.fill.value()`                        | Read fill model.                                                                      | none                                | `Promise<ExcelFillModel>`                   |
| `range.format.fill.setValue(fill)`                 | Set full fill model.                                                                  | `ExcelFillModel`                    | `Promise<void>`                             |
| `range.format.fill.setPatternType(patternType)`    | Set or clear fill pattern type.                                                       | `string \| undefined`               | `Promise<void>`                             |
| `range.format.fill.setForegroundColor(color)`      | Set or clear foreground color. Setting a color also sets the fill pattern to `solid`. | `ExcelColorModel \| undefined`      | `Promise<void>`                             |
| `range.format.fill.setBackgroundColor(color)`      | Set or clear background color.                                                        | `ExcelColorModel \| undefined`      | `Promise<void>`                             |
| `range.format.border.value()`                      | Read border model.                                                                    | none                                | `Promise<ExcelBorderModel>`                 |
| `range.format.border.setValue(border)`             | Set full border model.                                                                | `ExcelBorderModel`                  | `Promise<void>`                             |
| `range.format.border.setLeft(side)`                | Set or clear left border side.                                                        | `ExcelBorderSideModel \| undefined` | `Promise<void>`                             |
| `range.format.border.setRight(side)`               | Set or clear right border side.                                                       | `ExcelBorderSideModel \| undefined` | `Promise<void>`                             |
| `range.format.border.setTop(side)`                 | Set or clear top border side.                                                         | `ExcelBorderSideModel \| undefined` | `Promise<void>`                             |
| `range.format.border.setBottom(side)`              | Set or clear bottom border side.                                                      | `ExcelBorderSideModel \| undefined` | `Promise<void>`                             |
| `range.format.alignment.value()`                   | Read alignment model.                                                                 | none                                | `Promise<ExcelAlignmentModel \| undefined>` |
| `range.format.alignment.setValue(alignment)`       | Set or clear alignment model.                                                         | `ExcelAlignmentModel \| undefined`  | `Promise<void>`                             |
| `range.format.alignment.setHorizontal(horizontal)` | Set or clear horizontal alignment.                                                    | `string \| undefined`               | `Promise<void>`                             |
| `range.format.alignment.setVertical(vertical)`     | Set or clear vertical alignment.                                                      | `string \| undefined`               | `Promise<void>`                             |
| `range.format.alignment.setWrapText(wrapText)`     | Set or clear wrap text.                                                               | `boolean \| undefined`              | `Promise<void>`                             |

Formatting writes are style-table based. Repeated equivalent formats reuse the same style entries where possible. Reads use the first non-empty resolved style in the range, matching the live-object convention that a range can represent mixed formatting.

```ts
await range.format.font.setColor({ rgb: '#FF0000' });
await range.format.fill.setForegroundColor({ rgb: '#D9D9D9' });
await range.format.border.setValue({
  left: { style: 'thin', color: { rgb: '#000000' } },
  right: { style: 'thin', color: { rgb: '#000000' } },
  top: { style: 'thin', color: { rgb: '#000000' } },
  bottom: { style: 'thin', color: { rgb: '#000000' } },
});
```

## Cell Object

| API                        | Purpose                                    | Params                | Returns                           |
| -------------------------- | ------------------------------------------ | --------------------- | --------------------------------- |
| `cell.address()`           | Read cell address.                         | none                  | `Promise<string>`                 |
| `cell.value()`             | Read cell value.                           | none                  | `Promise<ExcelCellValue>`         |
| `cell.formula()`           | Read formula without leading `=`, or null. | none                  | `Promise<string \| null>`         |
| `cell.text()`              | Read display text approximation.           | none                  | `Promise<string>`                 |
| `cell.valueType()`         | Read cell kind.                            | none                  | `Promise<ExcelCellKind>`          |
| `cell.rawValue()`          | Read raw OOXML value.                      | none                  | `Promise<string \| null>`         |
| `cell.styleId()`           | Read OOXML style id.                       | none                  | `Promise<number \| undefined>`    |
| `cell.setValue(value)`     | Write one value.                           | `ExcelCellValue`      | `Promise<void>`                   |
| `cell.setFormula(formula)` | Write or clear one formula.                | `string \| null`      | `Promise<void>`                   |
| `cell.setStyleId(styleId)` | Set or clear OOXML style id.               | `number \| undefined` | `Promise<void>`                   |
| `cell.clear()`             | Clear this cell.                           | none                  | `Promise<void>`                   |
| `cell.toRange()`           | Return the one-cell range handle.          | none                  | `ExcelRange`                      |
| `cell.toModel()`           | Return cell snapshot or null.              | none                  | `Promise<ExcelCellModel \| null>` |

## Value Types

| Type      | Meaning                                               |
| --------- | ----------------------------------------------------- |
| `blank`   | Empty or missing cell.                                |
| `string`  | Shared string, inline string, or direct string value. |
| `number`  | Numeric value.                                        |
| `boolean` | Boolean value.                                        |
| `formula` | Formula cell.                                         |
| `error`   | Error value.                                          |

`ExcelCellValue` is `string | number | boolean | null`.

## CodeAgent Rules

1. Use ranges for bulk reads/writes; use cells for scalar reads/writes.
2. Always provide full matrix dimensions for `setValues` and `setFormulas`.
3. Use 0-based offsets only inside `range.getCell`, `getRow`, `getColumn`, `offset`, and `slice`.
4. Use 1-based A1 addresses for worksheet addresses. `sheet.getRange(...)` accepts ordinary cell ranges such as `A1:D4`, whole-column ranges such as `A:A` or `A:C`, and whole-row ranges such as `1:1` or `1:7`.
5. Do not expect formulas to include a leading `=` when reading.
6. Use `styleId` APIs only when preserving or applying known workbook style indexes.
7. `merge()` and `unmerge()` currently operate on the exact range address, not partial-overlap merge splitting.
8. Prefer `range.format` for font/fill/border/number-format/alignment changes instead of manually composing style ids.
9. Always express colors as `{ rgb: "#RRGGBB" }`; never emit ARGB, bare hex, `theme`, or `indexed` color values.
10. There is no `autofitColumns()` / `autoFitColumns()` API in Excel Core. For "fit columns to content" tasks, read the relevant `range.text()`, estimate explicit widths from display text, and call `range.setColumnWidth(width)` on each target column.
