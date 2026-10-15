# Excel mutation writeback

`mutation writeback` is the fast path for applying SDK edits back to a live spreadsheet host.

It is not a replacement for package or sheet-shadow writeback. It records high-confidence SDK mutations while code runs, lets the runtime verify that those mutations cover the actual workbook diff, and lets host adapters replay the operations directly through live APIs.

## Goals

1. Avoid full workbook or all-sheet overwrite when the agent only changed a small range.
2. Keep the public Excel SDK live-object API stable for small models.
3. Preserve existing writeback fallbacks for complex workbook structures.
4. Make writeback logs explain what was applied: operation type, sheet, range, and strategy.

## Flow

```text
Excel SDK live object API
  -> record ExcelWorkbookMutation
  -> mutate ExcelWorkbookModel
  -> runtime compares before/after model
  -> host replays mutations when safe
  -> fallback to sheet-shadow/package writeback when unsafe
```

## Mutation contract

Mutations are plain JSON values exported as `ExcelWorkbookMutation`.

First-phase replayable mutations:

- `range.setValues`
- `range.setFormulas`
- `range.clear`
- `range.merge`
- `range.unmerge`
- `worksheet.add`
- `worksheet.delete`
- `worksheet.rename`

Unsupported or broad changes must record:

```ts
{
  kind: "unsafe",
  scope: "worksheet" | "range" | "object" | "workbook",
  reason: "..."
}
```

Examples that currently force fallback:

- row/column insert or delete
- row height or column width changes
- style-table or fine-grained format changes
- worksheet copy
- complex objects unless they are explicitly covered by a later mutation type

## Runtime gate

The host must not replay mutations just because a mutation list exists. The runtime must first compare:

- model before code execution
- model after code execution
- recorded mutation list

Replay is allowed only when:

1. mutation list is non-empty
2. no `unsafe` mutation exists
3. workbook defined names did not change
4. workbook style table did not change
5. changed sheets can be explained by supported mutations
6. no changed sheet has object changes

If any condition fails, runtime passes no replay permission and the host uses the existing fallback strategy.

## Scoped fallback

When replay is not allowed, the runtime still computes a workbook diff and passes a fallback scope to the host:

- `changedSheetNames`: added, renamed, or modified sheets that exist in the generated workbook.
- `deletedSheetNames`: sheets removed from the generated workbook.

Host adapters should use this scope to avoid all-sheet overwrite:

1. If `changedSheetNames` is non-empty, only copy/replace those sheets from the generated workbook.
2. If `deletedSheetNames` is non-empty, only delete those explicit target sheets.
3. Do not delete target sheets merely because they are absent from the scoped source list.
4. If no scope is provided, keep the legacy full workbook/all-sheet fallback behavior.

This keeps complex operations safe while avoiding unnecessary copy cost for common cases such as formatting a single sheet, modifying a table-like area, or touching objects on one analysis sheet.

## Host mapping

### Microsoft Excel

Replay uses Office.js live APIs:

- `worksheet.add()`
- `worksheet.delete()`
- `worksheet.name = ...`
- `range.values = ...`
- `range.formulas = ...`
- `range.clear(Excel.ClearApplyTo.all)`
- `range.merge()`
- `range.unmerge()`

Fallback remains `insertWorksheetsFromBase64 + copyFrom + object repair`.
When replay is unsafe but a scoped fallback is available, the adapter inserts only the scoped source sheets and deletes only explicit `deletedSheetNames`.

### WPS

Replay uses WPSJS live APIs:

- `Worksheets.Add`
- `Worksheet.Delete`
- `Worksheet.Name = ...`
- `Range.Value2 = ...`
- `Range.Formula = ...`
- `Range.Clear()`
- `Range.Merge()`
- `Range.UnMerge()`

Fallback remains temp workbook open plus sheet copy.
When replay is unsafe but a scoped fallback is available, the adapter copies only the scoped source sheets and deletes only explicit `deletedSheetNames`.

## Expansion plan

Add mutation types only when all three pieces are ready:

1. SDK records the mutation.
2. Runtime can verify the mutation against before/after model state.
3. Microsoft/WPS/ONLYOFFICE host adapters have tested replay behavior or an explicit fallback.

Likely next candidates:

- range number format
- range fill/font/border patches
- row height and column width
- table add/delete
- chart add with constrained `ChartSpec`
