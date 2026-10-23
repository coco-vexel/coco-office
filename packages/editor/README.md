# @coco-vexel/coco-office-editor

Coco Office editor runtime for browser and Electron hosts.

The package is designed as an embeddable editor layer rather than an app-specific UI. Hosts provide file loading, toolbar UI, persistence, and desktop/browser integration. The editor package owns document projection, rendering, selection mapping, editing commands, and export-oriented change records.

After the initial npm release: `npm install @coco-vexel/coco-office-editor`. Core is installed as a dependency.
For a runnable integration, build the repository then run `corepack pnpm demo`.
See [runtime and demo requirements](https://github.com/coco-vexel/coco-office/blob/main/docs/runtime-and-demos.md).

## Quick start: a runnable PowerPoint editor

Use Node.js 22 or 24 for tooling and a modern browser with DOM/Canvas2D support. The runtime works in a browser or Electron renderer, not Node SSR. No Office installation is required. ESM/CJS entrypoints and TypeScript declarations are included; browser apps need a bundler. This quick start uses Vite.

In an empty directory, run:

```bash
npm init -y
npm install @coco-vexel/coco-office-editor
npm install --save-dev vite
```

Create `index.html`:

```html
<!doctype html>
<html lang="en">
  <meta charset="UTF-8" />
  <title>Coco PowerPoint editor</title>
  <input id="file" type="file" accept=".pptx" />
  <button id="save" disabled>Export copy</button>
  <p id="status">Open a PPTX file, edit its text, then export a copy.</p>
  <div id="editor" style="height: 75vh; min-height: 500px"></div>
  <script type="module" src="/main.js"></script>
</html>
```

Create `main.js`, then run `npx vite` and open the printed HTTP URL:

```js
import {
  buildPptEditorViewFromPptx,
  renderNativePptEditor,
} from '@coco-vexel/coco-office-editor/ppt';

const input = document.querySelector('#file');
const save = document.querySelector('#save');
const host = document.querySelector('#editor');
const status = document.querySelector('#status');
let editor;
let filename;
async function run(action) {
  input.disabled = save.disabled = true;
  try {
    await action();
  } catch (error) {
    status.textContent = error.message;
  } finally {
    input.disabled = false;
    save.disabled = !editor;
  }
}
input.addEventListener('change', () =>
  run(async () => {
    const file = input.files?.[0];
    if (!file) return;
    const view = await buildPptEditorViewFromPptx(new Uint8Array(await file.arrayBuffer()));
    editor?.destroy();
    editor = undefined;
    // A new mount avoids reattaching a ShadowRoot when switching files.
    const mount = document.createElement('div');
    mount.style.height = '100%';
    host.replaceChildren(mount);
    editor = renderNativePptEditor(mount, view);
    filename = file.name.replace(/\.pptx$/i, '.edited.pptx');
    input.value = '';
    status.textContent =
      'Edit supported text/shapes, then export. Export before opening another file.';
  }),
);
save.addEventListener('click', () =>
  run(async () => {
    const bytes = await editor.toPptx();
    const url = URL.createObjectURL(new Blob([bytes]));
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    status.textContent = 'Exported a copy. Reopen it to check the result.';
  }),
);
window.addEventListener('pagehide', () => editor?.destroy());
```

This is an unstyled integration shell; the native PPT renderer supplies its editor surface. The host provides file access, save UI and session-switch policy. For React/Vue, create the editor after the mount exists and call `destroy()` on unmount. Do not share a live editor instance between DOM mounts or create it during server rendering.

## Choose an editor and export path

| Document | Public entry            | Load/render                                              | Export                                                                                                           |
| -------- | ----------------------- | -------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| Word     | `/word`, `/word/native` | `buildWordEditorViewFromDocx` + `renderNativeWordEditor` | Apply `getChanges()` to a fresh input package, then Core `toZip`; inspect unsupported/skipped/unresolved results |
| Excel    | `/spreadsheet`          | `await renderSpreadsheetEditor(host, bytes, options)`    | `await editor.toZip()`                                                                                           |
| PPT      | `/ppt` or `/ppt/native` | `buildPptEditorViewFromPptx` + `renderNativePptEditor`   | `await editor.toPptx()`                                                                                          |

All import paths above start with `@coco-vexel/coco-office-editor`. If your app directly imports Core (for example to apply Word changes), declare it explicitly with `npm install @coco-vexel/coco-office`. Examples below named `host`, `xlsxBytes` or `persistWorkbook` are integration fragments: supply your own mount, file bytes and persistence callback.

### Read-only preview

All three native editors accept `readOnly: true` at mount time and expose
`isReadOnly()` / `setReadOnly(readOnly)`. Read-only mode is enforced by the
Word runtime, spreadsheet runtime, and PPT session as well as their DOM event
handlers. Programmatic spreadsheet and PPT mutations reject with
`EditorReadOnlyError` (`code: 'EDITOR_READ_ONLY'`); Word's synchronous command
contract reports the same condition through `canDispatch() === false` and
`dispatch() === false`.

Read-only preview intentionally keeps scrolling, zoom owned by the host,
selection/copy, links, worksheet switching, and slide navigation available.
It blocks typing, IME, cut/paste/drop, formatting, structure changes,
drag/resize, undo/redo, and mutation APIs. Navigation and selection do not emit
change events or advance revisions. Spreadsheet sort/filter/freeze operations
remain blocked because they currently modify the workbook rather than a
separate view state.

PPT hosts that render their own navigator can also pass `showSlideList: false`:

```ts
const preview = renderNativePptEditor(host, view, {
  readOnly: true,
  showSlideList: false,
});
```

Word custom Shadow DOM integration needs stylesheet text via `styleText`; Vite supports `@coco-vexel/coco-office-editor/styles.css?inline`. Other bundlers must provide an equivalent CSS-string loader. The `?inline` suffix is a Vite feature, not a portable JavaScript import. Spreadsheet's default isolated renderer and the default PPT renderer manage their surfaces. Use a mount with an explicit nonzero height.

The complete three-format file/open/edit/export flow, including Word export diagnostics, is in the [public demo source](https://github.com/coco-vexel/coco-office/blob/main/packages/editor/demo/main.js). No private product dependency is needed. These editors are experimental; file round-trip success does not certify visual parity or support for every advanced object.

## Scope

Word, spreadsheet and fixed-canvas PowerPoint editors share Core document APIs. All three browser surfaces are experimental; advanced objects and layout have feature-specific limits.

```text
@coco-vexel/coco-office-editor
  word/
    XML package or EditorDocumentView
    -> NativeWordDocument
    -> custom element DOM renderer
    -> selection bridge
    -> native command layer
    -> change records for export
  spreadsheet/   Canvas2D spreadsheet runtime
  ppt/           fixed-canvas slide runtime
```

## Word Architecture

The Word editor uses a native Coco model and custom elements as the core editing surface. The legacy ProseMirror spike remains isolated for compatibility and is not part of the active implementation path.

The runtime has five layers:

1. `extract.ts` reads a `.docx` package into `EditorDocumentView`.
2. `native-model.ts` projects `EditorDocumentView` into `NativeWordDocument`.
3. `native-renderer.ts` renders the native model into Word-specific custom elements.
4. `native-selection.ts` maps DOM selection and selected objects back to native model positions.
5. `native-commands.ts` mutates the native model through explicit commands and emits change records.

`native.ts` wires those layers into `renderNativeWordEditor(host, view, options)` for browser and Electron usage.

Export back to Word follows a two-step path:

1. `ooxml-patch.ts` maps native editor change records into typed OOXML patch intents.
2. A package executor applies those patch intents to the original `.docx` XML parts.

The browser demo applies changes to a fresh copy of the input package on each export and reports unsupported, skipped or unresolved changes before offering a download.

## Public Entrypoints

Use the native Word entrypoint for the self-developed editor path:

```ts
import { renderNativeWordEditor } from '@coco-vexel/coco-office-editor/word/native';
import {
  buildWordEditorViewFromDocx,
  buildWordEditorViewFromFlatOpc,
} from '@coco-vexel/coco-office-editor/word';
import editorStyles from '@coco-vexel/coco-office-editor/styles.css?inline';
import '@coco-vexel/coco-office-editor/styles.css';

const documentView = await buildWordEditorViewFromDocx(docxBytes);
const editor = renderNativeWordEditor(host, documentView, {
  readOnly: false,
  contextMenu: true,
  // Pass the bundled editor CSS text when the host uses Shadow DOM isolation.
  // Vite hosts can load it with: import editorStyles from '../styles/editor.css?inline';
  styleIsolation: 'shadow',
  styleText: editorStyles,
  onChange(event) {
    persistIncrementalChanges(event.changes, event.revertedChanges);
  },
});

// If an agent returns Flat OPC instead of a DOCX zip, reuse the same view boundary.
editor.loadDocument(buildWordEditorViewFromFlatOpc(flatOpcXml));

editor.setSelection({ kind: 'caret', paraId: 'paragraph-id', offset: 0 });
editor.dispatch({ kind: 'insert-text', text: 'Hello' });
```

`buildWordEditorViewFromDocx(bytes)` is the browser/Electron host helper for the normal load path. It accepts `Uint8Array` or `ArrayBuffer`, so hosts do not need to assemble `fromZip`, `buildIndices`, and `buildWordEditorViewFromPackage` themselves. `buildWordEditorViewFromFlatOpc(xml)` provides the same boundary when an agent returns Flat OPC XML.

The same runtime is used in a normal browser page and an Electron renderer. File access and IPC stay in the host application; the editor itself only requires a DOM and an `EditorDocumentView`.

The native Word runtime uses Shadow DOM isolation by default. The editor surface,
search highlights, comment dialogs, and context menus are mounted in the same
ShadowRoot, so host-page rules do not leak into the editing surface. Hosts should
provide the editor stylesheet through `styleText` (for example with Vite's
`?inline` CSS import). Legacy hosts that intentionally manage styles in light DOM
can opt out with `styleIsolation: 'none'`.

### Native runtime contract

`renderNativeWordEditor` returns a UI-independent runtime. A host toolbar should read `getCommandState()`, check `canDispatch(command)`, and call `dispatch(command)` rather than editing rendered custom elements directly.

```ts
const unsubscribeSelection = editor.subscribe('selectionChange', ({ selection, commandState }) => {
  toolbar.update(selection, commandState);
});

const unsubscribeChange = editor.subscribe('change', (event) => {
  // Incremental since the previous change notification. Undo moves records to
  // revertedChanges; redo returns them to changes.
  console.log(event.revision, event.changes, event.revertedChanges, event.totalChangeCount);
});

editor.setSelection({
  kind: 'text-range',
  anchor: { paraId: 'p1', offset: 0 },
  focus: { paraId: 'p2', offset: 4 },
});
editor.dispatch({ kind: 'toggle-bold' });
editor.dispatch({
  kind: 'set-hyperlink',
  hyperlink: { target: 'https://coco.example/docs', tooltip: 'Coco docs' },
});
editor.dispatch({ kind: 'remove-hyperlink' });

if (editor.canUndo()) editor.undo();
if (editor.canRedo()) editor.redo();

// Structural table commands are enabled for regular, unmerged grids.
editor.dispatch({ kind: 'insert-table', rows: 2, columns: 3 });
const table = editor.getCommandState().table;
if (table?.structuralEditingSupported) {
  editor.dispatch({ kind: 'insert-table-row', position: 'after' });
  editor.dispatch({ kind: 'insert-table-column', position: 'before' });
}
if (table) {
  editor.dispatch({ kind: 'set-table-cell-background', color: '#FFF2CC' });
  editor.dispatch({ kind: 'set-table-cell-vertical-alignment', alignment: 'center' });
  editor.dispatch({ kind: 'set-table-cell-borders', color: '#808080' });
  editor.dispatch({ kind: 'delete-table' });
}

// Insert from a browser File/Blob or a data URL, then edit the selected image.
await editor.insertImage(file, { width: 320, height: 180 });
const image = editor.getCommandState().image;
if (image) {
  editor.dispatch({
    kind: 'set-image-size',
    objectId: image.objectId,
    width: 400,
    height: 225,
  });
  editor.dispatch({
    kind: 'set-image-alignment',
    objectId: image.objectId,
    alignment: 'center',
  });
}

// Search body paragraphs and nested table content without changing the editor DOM.
const searchState = editor.search({
  query: 'OfficeAI',
  matchCase: false,
  wholeWord: false,
});
console.log(searchState.activeIndex, searchState.matches.length);
editor.findNext('forward');
editor.replaceCurrent('Coco');
const replacementCount = editor.replaceAll('Coco');
console.log(replacementCount);
editor.clearSearch();

// Reuse the host for another file without retaining selection or undo history.
editor.loadDocument(nextDocumentView);

unsubscribeChange();
unsubscribeSelection();
editor.destroy();
```

The lifecycle and integration rules are:

- `getChanges()` returns a snapshot of currently active export-oriented changes.
- `loadDocument()` clears selection, pending changes, and undo/redo history before rendering the new view.
- `destroy()` is idempotent, removes DOM/document listeners and observers, clears the host, and makes commands return `false`.
- `destroyed` exposes lifecycle state. `loadDocument()` after destruction throws because the instance cannot be reused.
- Multiple runtimes may coexist in one document. Document-level shortcuts are routed only to the most recently active editor.
- The built-in HTML context menu is enabled by default in browsers and Electron renderers. Set `contextMenu: false` when the host supplies its own menu UI.
- `getDocument()` exposes the live native projection for inspection and export integration; hosts must not mutate it directly.
- Clicking an inline image creates an object selection exposed as `getCommandState().image`. The native resize handle preserves the aspect ratio by default; hold Shift while dragging for independent width and height changes.
- Image insertion, resize, alignment, deletion, undo, and redo are supported in the live editor model. Ordinary inline image insertion, resize, deletion, and tracked image insertion resolution are supported by the current OOXML patch executor and covered by DOCX round-trip tests. Complex floating drawings, unsupported drawing types, and layout-specific wrap semantics remain conservative and must not be treated as fully editable.

The broader Word barrel is still available:

```ts
import { extractWordDocumentView } from '@coco-vexel/coco-office-editor/word';
```

The self-developed spreadsheet editor entrypoint is backed by Coco Excel Core and is usable from browser and Electron hosts:

See [SPREADSHEET_EDITOR_API.md](https://github.com/coco-vexel/coco-office/blob/main/packages/editor/SPREADSHEET_EDITOR_API.md) for the full
custom-host integration contract and examples.

```ts
import { renderSpreadsheetEditor } from '@coco-vexel/coco-office-editor/spreadsheet';

const editor = await renderSpreadsheetEditor(host, xlsxBytes, {
  workbookId: 'browser-document.xlsx',
  readOnly: false,
  showSheetTabs: true,
  // Shadow DOM isolation is enabled by default; use 'none' only for legacy hosts.
  styleIsolation: 'shadow',
  // Optional; column menus remain available when this is omitted.
  enableHeaderFilters: false,
  onChange({ bytes }) {
    persistWorkbook(bytes);
  },
});

await editor.setActiveSheet('Sheet2');
const nextBytes = await editor.toZip();
editor.destroy();
```

`renderSpreadsheetEditor()` owns the canvas stack, viewport lifecycle, DPI-aware
resizing, scrolling, and worksheet switching. The host owns toolbar UI,
Agent panels, persistence policy, and browser/Electron integration. The lower
level `SpreadsheetEditorRuntime`, `SpreadsheetLayout`, and renderer exports
remain public for custom hosts and tests, but normal web and desktop hosts
should use the entrypoint so those responsibilities are not reimplemented in
each application.

The public entrypoint uses Shadow DOM style isolation by default. This keeps
editor controls, cell editing, menus, scrollbars, and sheet tabs from inheriting
host-page CSS. Hosts that require the legacy light-DOM behavior can opt out with
`styleIsolation: 'none'`.

The public surface also owns selection, cell editing, copy/paste, formula
reference highlighting, merged-cell selection, auto fill, row/column resizing,
filtering, sorting, freeze panes, and worksheet structure commands. Header
filter icons are opt-in with `enableHeaderFilters`; when disabled, filtering and
sorting remain available from a column-header context menu.

Formatting commands include font/fill colors, bold/italic/underline,
horizontal/vertical alignment, number formats, borders, and format clearing.

Custom hosts can subscribe to `subscribeState()` for selection, active-sheet,
and editing changes. Multi-step UI or Agent actions should use
`editor.runTransaction({ label }, async (editor) => { ... })`; the commands in
the callback share one Core history group and produce one render/export change
event after the transaction completes.

### Spreadsheet worksheets

`SpreadsheetEditorRuntime` exposes the active worksheet through `activeSheet()` and switches it with `await setActiveSheet(name)`. The runtime rebuilds the render model and formula results for the selected worksheet; cell edits always apply only to that active sheet.

Web and desktop hosts should own view-only state by sheet name. Before a tab switch, store the current selection and scroll position; after `setActiveSheet()`, restore the target sheet's saved state, defaulting to `A1` and `{ x: 0, y: 0 }`. Do not store this host state in the workbook model or undo history. The dev host follows this rule and also serializes tab switches so quick consecutive clicks cannot apply out of order.

Freeze panes are workbook-backed view settings, not host-only scroll state. Use `await runtime.setFreezePanes(frozenRows, frozenColumns)` to persist them through undo/redo and XLSX export. The dev toolbar derives the split from the active cell: selecting `C4` and choosing `Freeze panes` freezes rows `1:3` and columns `A:B`; `Freeze rows above`, `Freeze columns left`, and `Unfreeze panes` cover the common variants. Hosts should continue to keep their transient scroll offsets separate from these persisted pane settings.

Worksheet structure changes are workbook-backed as well. Use `addSheet()`, `renameSheet()`, `copySheet()`, `moveSheet()`, and `deleteSheet()` on `SpreadsheetEditor` or `SpreadsheetEditorRuntime`. Direct user sheet-tab operations deliberately do not enter the Ctrl/Cmd+Z history, matching desktop spreadsheet behavior; deletion must be protected by host UI confirmation. When the same methods run inside `runAgentTransaction()`, they create a `workbook.sheets` structural patch so the complete Agent action can roll back atomically. The dev host exposes add from the `+` sheet-tab button and rename, duplicate, move, and delete from a sheet-tab context menu. A renamed sheet retains its host-side selection and scroll state; a copied sheet opens at `A1`; deleting the active sheet selects its nearest neighbor.

### Spreadsheet auto fill

`autoFillRange(sourceRange, targetRange)` extends a cell selection down or right as one user undo group. The first version repeats ordinary values and styles, continues a consistent numeric sequence when at least two numeric source cells establish one, and rebases relative formula references through Core's `shiftFormulaReferences()` utility. `$`-anchored row/column references and formula string literals remain unchanged. The dev canvas presents the fill handle on the selection's lower-right corner and previews the target range while dragging. Merged selections intentionally do not expose a handle yet, because merged-block expansion requires structural copy semantics rather than a cell fill.

### Spreadsheet merged cells

Spreadsheet hosts should treat merged ranges as one visual editing unit:

- Use `expandSelectionForMergedRanges()` before rendering drag/paste selections. If a range touches any merged block, the selection expands to include the full block.
- Internal copy keeps merged range metadata and exports text only from the merge master. Covered cells are copied as blank even if an imported workbook still contains stale hidden values.
- Internal copy also carries cell style ids for same-workbook paste, so bold/italic/fill/border/number-format state can paste with values. Plain external TSV paste remains value-only.
- Paste first unmerges intersecting target merges, writes the rectangular payload, then reapplies copied merges. Undo records that paste as one user transaction.
- Runtime paste supports `mode: 'all' | 'values' | 'formats'`. `all` writes values, styles, and copied merges; `values` writes only values; `formats` writes only style ids.
- The dev host context menu exposes `Paste`, `Paste values`, and `Paste formats`. `Paste formats` is shown only when the current clipboard matches an internal Coco copy payload.
- The dev host supports lightweight find and jump: `Ctrl/Cmd+F` opens a sheet search bar, Enter/Shift+Enter moves through matches, and the replace field can replace all matching values/formula text as one undo group. The name box accepts both cells such as `C5` and ranges such as `B2:D8`.
- Spreadsheet filtering is disabled by default. After enabling the toolbar `Filter` toggle, the first spreadsheet filter is opened by clicking the small funnel icon inside a first-row header cell; clicking the `A/B/C` coordinate header still selects the whole column. The menu uses a scroll-limited checkbox list, supports selecting multiple values, blank values, applying the selection, and clearing the filter. It also provides A-Z and Z-A sorting, which moves complete data rows through one Core snapshot. The `Headers` toggle controls whether row 1 is excluded from sorting/filtering; disable it for data-only sheets without a header row. The legacy toolbar filter remains available for explicit multi-row selections. Non-matching data rows are hidden through one Core snapshot transaction. `filterRows()` accepts one value or a value array, and `sortRows()` accepts a direction; both operations remain exportable and undoable.
- Editor merge/unmerge normalizes hidden follower styles to the merge master's style before changing structure. This prevents stale covered-cell formatting from reappearing after unmerge.
- Row/column deletion trims or shifts merged ranges through core structural patches. Undo/redo restores the deleted structure and merged ranges without requiring a full-sheet snapshot when formulas do not need reference recovery.

Merged-cell smoke checklist for web/desktop hosts:

- Click a covered cell in a merged block; selection should move to the merge master and cover the full merged range.
- Drag a range that partially touches a merged block; selection should expand until every touched merged block is fully included.
- Right-click inside an existing range selection; the range selection should not collapse to the clicked cell.
- Insert rows/columns from the context menu; the inserted rows/columns should remain selected as a range.
- Delete rows/columns from the context menu; selection should move to the nearest valid cell/row/column and expand if it lands inside a merged block.
- Copy/paste a merged range, then unmerge the pasted block; only the merge master should keep copied content.
- Copy/paste a formatted range; values and visible formatting should land together and undo/redo as one paste.
- Paste special values and paste special formats through runtime API; values-only should not overwrite formatting, and formats-only should not overwrite values.
- Use the context menu paste commands after an internal copy; `Paste values` should ignore styles/merges, while `Paste formats` should keep existing values.
- Use `Ctrl/Cmd+F` to find visible values, raw values, or formula text; results should wrap and scroll the active match into view.
- Type an address or range into the name box and press Enter; selection should jump there and scroll into view.
- Toggle merge/unmerge and undo; history should use `sheet.mergedRanges` patches without a full sheet snapshot.
- Delete rows/columns through a merged block and undo/redo; the visible merge shape and values outside the merge should round-trip.
- Drag column and row header boundaries; the resize guide should track the pointer, clamp to minimum sizes, commit through core row/column size patches, and undo/redo cleanly.
- Select a mixed-format range; toolbar and cell settings should show dashed mixed states for B/I/U, colors, number format, and alignment instead of using only the active cell.

### Spreadsheet Agent actions

Spreadsheet agent integrations should use action objects instead of separate preview and apply implementations. An action is a reusable runner plus metadata. The same `run()` function is executed against a cloned workbook for preview and against the live workbook for apply, so web and desktop hosts do not accidentally show one result and commit another.

```ts
import {
  applySpreadsheetAgentAction,
  previewSpreadsheetAgentAction,
  type SpreadsheetAgentAction,
} from '@coco-vexel/coco-office-editor/spreadsheet';

const action: SpreadsheetAgentAction = {
  id: 'fill-summary',
  label: 'Fill summary',
  actor: 'spreadsheet-agent',
  agentRunId: runId,
  agentStepId: stepId,
  async run(runtime, context) {
    const range = context.selection ?? { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 };
    await runtime.setRangeValues(range.startRow, range.startColumn, [['Summary', 42]]);
    await runtime.setRangeBold(range.startRow, range.startColumn, 1, 2, true);
  },
};

const context = { selection: currentSelection };
const preview = await previewSpreadsheetAgentAction(runtime, action, context);
renderConfirmation(preview.transaction.summary, preview.transaction.diffSummary);

// Call this only after the user or host policy confirms the preview.
const applied = await applySpreadsheetAgentAction(runtime, action, context);
```

Host integration rules:

- Build one `SpreadsheetAgentAction` per agent-visible operation.
- Use `previewSpreadsheetAgentAction()` for confirmation UI. It runs on a cloned workbook and does not mutate the live editor or live undo stack.
- Use `applySpreadsheetAgentAction()` to commit. It records one agent transaction group, so `undo()` reverts the whole action in one step.
- Keep preview and apply on the same action and context. Do not duplicate the runner logic in UI code.
- Store `summary` and `diffSummary` for agent explanations, audit trails, and desktop/web confirmation dialogs.

The ProseMirror entrypoint remains isolated:

```ts
import { createWordEditor } from '@coco-vexel/coco-office-editor/word/prosemirror';
```

## Current Word Coverage

The detailed live-editing versus DOCX-writeback boundary is maintained in [docs/word-editor-capability-matrix.md](../../docs/word-editor-capability-matrix.md).

Implemented in the native path:

- Paragraph projection and rendering.
- Inline text and image projection.
- Table rows, cells, and nested paragraphs.
- Paragraph style rendering for alignment, spacing, indentation, line height, font size, bold, color, and borders.
- Run style rendering for bold, italic, underline, font size, and color.
- Table cell style rendering for width, shading, and vertical alignment.
- Basic editing commands for text insertion, backward delete, paragraph split/merge, image insertion/delete/resize/alignment, and browser text sync.
- Row and column insertion/deletion for regular tables, with local table rendering and undo/redo snapshots.
- Top-level table insertion at a body caret and whole-table deletion, with paragraph splitting, local block rendering, and one-step undo/redo.
- Current-cell background, vertical alignment, and four-side border formatting, including merged cells and undo/redo.
- Rectangular table-cell selection by cross-cell drag or Shift-click, with arrow-key extension, row/column selection, bulk formatting, clear, and one-step undo/redo.
- Table-cell range copy/cut/paste uses a private two-dimensional rich-content payload inside Coco, preserving paragraph/run formatting, cell presentation, and images with fresh object IDs. Other applications receive tab-separated plain text.
- Body and nested-table find/replace with case and whole-word options, CSS Highlight rendering, cyclic navigation, and one-step undo/redo for replace-all.
- External and document-anchor hyperlink parsing, native rendering, selection state, set/remove commands, rich clipboard preservation, and one-step undo/redo. External links open only through Ctrl/Cmd-click.
- A shared browser/Electron HTML context menu for history, clipboard, text, image, and table commands.
- OOXML patch planning for text, paragraph, image, comment, revision, table, header, and footer changes.
- OOXML patch application for ordinary text, formatting, inline images, comments, revisions, regular tables, headers, and footers, with conservative handling for unsupported structures.
- Coverage summaries for comparing source `EditorDocumentView` content against rendered native content.

Known gaps:

- Complex text ranges spanning unsupported Word run structures are detected and skipped by the executor.
- Complex Word layout remains partial: floating drawings, footnotes, section layout, advanced numbering, and special wrapping semantics need explicit model support.
- Complex merged-grid structural edits remain disabled when the original table structure cannot be represented safely.
- Collaborative editing and remote-operation reconciliation are not implemented.
- Find/replace does not currently include headers, footers, comments, hidden revision content, or phrases spanning paragraph boundaries.
- External http/https hyperlink changes are written back to DOCX relationships; document-anchor navigation works in the editor, while bookmark-aware anchor writeback waits for first-class bookmark projection.
- Cross-table mixed text selection is not yet an editable range.
- Structural editing of complex merged table grids remains disabled to prevent content loss; regular row/column changes and horizontal merge/split are applied to DOCX XML during export.

## Development

Run the Word editor dev page:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor dev
```

Open the native engine:

```text
http://127.0.0.1:5177/?engine=native
```

Validate the package:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor test
corepack pnpm --filter @coco-vexel/coco-office-editor build
```

## Extension Direction

The native Word runtime should stay model-first. New document features should usually be added in this order:

1. Preserve the information in `EditorDocumentView` extraction.
2. Add a typed native model representation.
3. Render the native node or decoration through custom elements.
4. Map selection or object interaction back to model positions.
5. Add command records that can later be exported to OOXML.

This keeps display fidelity, editing operations, and export semantics tied to the original Word structure instead of a generic HTML document shape.

## Standalone browser demo

From the repository root, run `pnpm build` then `pnpm demo`. The demo accepts local DOCX/XLSX/PPTX files and generates synthetic samples without external documents. `pnpm build:demo` produces static files without publishing them. Export changes before switching documents.

## License

Apache-2.0 for original code; see LICENSE and THIRD_PARTY_NOTICES.md. The Core dependency carries its own notices.
