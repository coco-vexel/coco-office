# Excel Host Integration Notes

This document maps the Excel core SDK to host read/writeback strategies discovered during technical probes.

Use this document when the task says: Microsoft Excel add-in, WPS, ONLYOFFICE, desktop, writeback, full replacement, sheet-level replacement, package bytes, host bridge, or agent runtime.

## Core Boundary

Core is host-neutral:

```text
host source -> xlsx bytes or ExcelWorkbookModel -> ExcelWorkbook -> toModel/toZip -> host writeback
```

Core does not call Office.js, WPS JS, ONLYOFFICE plugin APIs, COM, file pickers, or download flows. Each host bridge owns source acquisition and writeback.

Use [mutation-writeback.md](./mutation-writeback.md) for the fast path where SDK-recorded operations can be safely replayed through host live APIs. Use [writeback-planner.md](./writeback-planner.md) when an adapter has both the source model and target model and needs a portable writeback intent list. Both mechanisms are part of core-level planning; executing the returned operations is still adapter responsibility.

## Host Strategy Summary

| Host | Read Source | Writeback Strategy | Current Conclusion |
| --- | --- | --- | --- |
| Microsoft Excel add-in | `Office.context.document.getFileAsync(Compressed)` can export current workbook package including unsaved state. | No discovered in-place whole-package replacement API. Use inserted shadow workbook/sheet plus sheet-level copy/repair, or diff-based live operations. | Read is package-capable. Writeback is sheet-level/diff, not whole-package replacement. |
| WPS Spreadsheet | `ActiveWorkbook.SaveAsBinaryString()` quickly returns zip-like workbook bytes including unsaved marker. `SaveCopyAs` also works but can freeze UI on larger data. | No discovered in-place whole-package replacement API. Use generated/shadow workbook plus cross-workbook sheet copy, then repair cross-workbook formulas and objects as needed. | Read is package-capable. Writeback should avoid slow full save-copy paths. |
| ONLYOFFICE Spreadsheet | Live API can read/write cell/range data. Package/download APIs can trigger desktop save/download UI and are unsafe for silent source acquisition. | Range copy is fast for values/formulas/formats. Sheet-shadow copy is the current in-place fallback for changed sheets. Comments/shapes and some objects require live-object repair. | Prefer model/live-object bridge. Package path is unsafe in desktop plugin. |
| Desktop integration | File system or automation can provide package bytes and replace files directly when safe. | Direct package replacement or application automation depending on product policy. | Easiest place for package-level source and writeback. |

## Microsoft Excel Notes

Confirmed read behavior:

- `getFileAsync(Compressed)` can export `.xlsx` package bytes.
- Exported package includes unsaved workbook state.

Confirmed writeback constraints:

- `Office.context.document.setSelectedDataAsync(..., Ooxml)` is not supported in Excel for workbook/sheet replacement.
- `workbook.insertWorksheetsFromBase64(base64)` inserts worksheets; it does not replace the existing sheet in place.
- No whole-package in-place replacement candidate was discovered on `workbook`.
- First-phase mutation replay is available for safe SDK operations: values, formulas, range clear, merge/unmerge, and worksheet add/delete/rename.

Practical bridge:

1. Read current package bytes.
2. Load into `ExcelWorkbook`.
3. Apply core edits.
4. Export package/model.
5. Prefer mutation replay when the runtime marks the mutation list safe.
6. Otherwise use host sheet-level writeback: insert/shadow target, copy used range, copy floating objects where possible, repair chart series refs, delete temporary sheet.

## WPS Notes

Confirmed read behavior:

- `ActiveWorkbook.SaveAsBinaryString()` returns zip-like workbook content quickly in probe data.
- It can include unsaved edits.
- `SaveCopyAs(temp.xlsx)` can produce a file copy, but larger probes freeze or pause WPS UI.

Confirmed writeback constraints:

- Direct package replacement opens a new workbook instead of replacing the current workbook in place.
- No silent in-place whole-package rehydration API was confirmed.
- `Range.Copy` is fast and high-fidelity for same-workbook sheet overwrite.
- Cross-workbook formulas can require repair because references may bind to the source workbook/sheet.
- First-phase mutation replay is available for safe SDK operations: values, formulas, range clear, merge/unmerge, and worksheet add/delete/rename.

Practical bridge:

1. Prefer fast source read with `SaveAsBinaryString()`.
2. Avoid `SaveCopyAs` for frequent agent loops unless user explicitly accepts UI freeze risk.
3. Prefer mutation replay when the runtime marks the mutation list safe.
4. Otherwise build a shadow workbook/sheet from core output.
5. Copy sheet content into the live workbook using WPS range/sheet copy.
6. Repair formulas that refer to the shadow workbook.
7. Validate object fidelity for charts, comments, hyperlinks, formats, validation, conditional formats, and shapes.

## ONLYOFFICE Notes

Confirmed live behavior:

- Direct live writes to cells/ranges work.
- Range copy is fast for large matrix data.
- Range copy preserves values, formulas, and basic formats.
- Object preservation is incomplete for comments/shapes in plain range copy; targeted live-object repair can restore known comments and shapes.

Confirmed package constraints:

- `GetFileToDownload` or related package/download paths may trigger save/download UI and can block the plugin flow.
- Treat package download as unsafe for silent agent operations in desktop plugin mode.
- Silent whole-workbook package replacement is not proven.
- The current in-place fallback is sheet-shadow copy: write changed SDK sheet models into temporary sheets, copy their grid ranges into the original sheets, then delete the temporary sheets.

Practical bridge:

1. Use ONLYOFFICE live API to build an `ExcelWorkbookModel` when package bytes are unavailable.
2. Apply edits through core.
3. Prefer safe mutation replay for small SDK-recorded operations when runtime verification passes.
4. Otherwise use sheet-shadow copy for changed sheets.
5. Use fast range copy for grid data and formulas.
6. Use object-specific repair for comments, shapes, charts, images, and other floating objects as support matures.

Current fidelity boundary:

- Sheet-shadow copy covers grid-layer content first: values, formulas, blanks, numbers, strings, booleans, and merged ranges.
- Charts, shapes, images, comments, tables, data validation, and conditional formatting are not guaranteed by sheet-shadow copy. They require object-specific host repair because ONLYOFFICE plain `Range.Copy(targetRange)` does not clone every object-layer feature.

## Agent Rules

1. Generated core operation code must not call host APIs.
2. Host adapters may choose package, model, or diff bridges, but the SDK API stays the same.
3. Prefer package read when available and non-blocking.
4. Prefer full workbook package writeback only in desktop/file contexts where replacement is actually available.
5. Prefer mutation replay before sheet-shadow writeback when the runtime can prove the SDK mutation list covers the actual model diff.
6. In plugin hosts, assume writeback is sheet-level or diff-level unless a proven in-place package API exists.
7. Treat host object fidelity as a separate adapter concern; do not leak host-specific repair APIs into the core SDK.
