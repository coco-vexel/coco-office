# Spreadsheet Editor API

This document defines the integration boundary for browser and desktop hosts.
The editor owns spreadsheet interaction and rendering; the host owns its
toolbar, Agent panel, persistence, routing, and application layout.

The editor is backed by `@coco-vexel/coco-office/excel`. Hosts must not create a
second workbook model or reimplement spreadsheet mutations in the UI layer.

## Initialize

```ts
import {
  renderSpreadsheetEditor,
  type SpreadsheetEditor,
} from '@coco-vexel/coco-office-editor/spreadsheet';

const editor: SpreadsheetEditor = await renderSpreadsheetEditor(host, bytes, {
  workbookId: documentId,
  // Set true for preview while preserving scroll, selection, copy and sheet tabs.
  readOnly: false,
  showSheetTabs: true,
  // Isolate editor internals from host-page CSS.
  styleIsolation: 'shadow',
  // Header filter icons are opt-in. Column-header context menus still work.
  enableHeaderFilters: false,
  onChange({ bytes }) {
    persistDraft(documentId, bytes);
  },
});
```

Use `editor.setReadOnly(true)` to switch an existing instance into preview
mode. Mutation methods and direct mutation calls through `editor.runtime()`
reject with `EditorReadOnlyError`; `setActiveSheet()`, selection, scrolling,
copy and export remain available.

The host should keep the returned instance until the document is closed, then
call `editor.destroy()`. Replacing a document uses `loadWorkbook()` on the same
instance or destroys and creates a new instance when the host container changes.

## State

`selection()` is a synchronous snapshot. `subscribeState()` is the preferred
way for custom controls to update their active state:

```ts
const unsubscribe = editor.subscribeState((event) => {
  const { selection, editingCell, activeSheetName } = event;
  updateToolbar({
    activeSheetName,
    address: selection.active,
    editing: editingCell !== null,
  });
});
```

Read-only workbook state is available without reaching into the runtime:

```ts
const cell = editor.cellState('B4');
const freeze = editor.freezeState();
const filters = editor.filterState();
```

`cellState()` returns the displayed value, formula, style id, and resolved
font/fill/border/alignment/number-format information. `freezeState()` and
`filterState()` return snapshots; callers must not mutate them as a way to
change the workbook.

## Commands

Use editor commands from custom UI instead of calling `runtime()` directly.
Commands refresh the renderer and emit the persistence callback consistently.

```ts
await editor.setCellValue('B4', 42);
await editor.setCellFormula('C4', 'B4*2');

await editor.formatSelection('bold');
await editor.setSelectionFontColor({ rgb: '#C00000' });
await editor.setSelectionFillColor({ rgb: '#FFF2CC' });
await editor.setSelectionHorizontalAlign('center');
await editor.setSelectionVerticalAlign('center');
await editor.setSelectionNumberFormat('#,##0.00');
await editor.setSelectionBorder('outer');
```

Structural and data commands follow the same rule:

```ts
await editor.mergeSelection();
await editor.freezeSelection();
await editor.sortActiveColumn('asc');
await editor.filterColumn(2, ['Open', 'Pending']);
await editor.filterColumn(2, null); // clear this column filter
```

Worksheet operations are exposed on the editor as well:

```ts
await editor.addSheet('Summary');
await editor.renameSheet('Summary', 'Report');
await editor.copySheet('Report', 'Report copy');
await editor.moveSheet('Report copy', 0);
await editor.deleteSheet('Report copy');
```

Sheet tab operations are intentionally not regular Ctrl/Cmd+Z edits, matching
desktop spreadsheet behavior. Hosts should show a confirmation UI before
deleting a worksheet.

## Transactions

When one user or Agent action performs several commands, use one transaction.
It creates one Core history group, one undo step, one final render, and one
`onChange` event:

```ts
await editor.runTransaction({ label: 'Prepare summary' }, async (tx) => {
  await tx.setCellValue('A1', 'Total');
  await tx.setCellFormula('B1', 'SUM(B2:B20)');
  await tx.formatSelection('bold');
  await tx.setSelectionBorder('outer');
});
```

Agent preview/apply remains a Core runtime concern. A host that needs preview
diffs can use `runtime().previewAgentTransaction()`; after an applied Agent
operation, the host should call or wrap changes through the editor boundary so
the canvas and persistence snapshot stay synchronized.

## Export and lifecycle

```ts
const xlsxBytes = await editor.toZip();
unsubscribe();
editor.destroy();
```

Do not store selection, scroll offsets, or open menus in the XLSX model. Those
are view state and belong to the host or editor instance. Freeze panes,
worksheet structure, cell values, formulas, and formatting are workbook state
and are exported through `toZip()`.

## Scope boundary

The editor intentionally focuses on rendering, viewing, editing, and common
spreadsheet operations. It is not intended to reproduce every Excel feature.
Charts, macros, pivot tables, and advanced collaboration are separate product
capabilities. Charts are preserved in the workbook but are not rendered by the
editor yet; when a chart anchor is available, the editor places a same-size
notice at that location and directs the user to Excel or WPS for chart viewing.
These capabilities should not leak into the editing API until their data
contracts are defined.
