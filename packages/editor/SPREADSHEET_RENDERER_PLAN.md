# Coco Spreadsheet Renderer Plan

This document records the first implementation plan for the self-developed spreadsheet online editor in `@coco-vexel/coco-office-editor`.

## Positioning

The spreadsheet editor is an online editing surface for Coco Excel Core. It is not a second Excel core, a second workbook model, or an agent operation framework.

```text
@coco-vexel/coco-office/excel
  Owns xlsx parsing, normalized workbook model, runtime, mutations, writeback planning, and xlsx export.

@coco-vexel/coco-office-editor/spreadsheet
  Owns rendering, scrolling, frozen panes, selection, cell editing, clipboard, shortcuts, and editor UI integration.
```

The editor package should use `spreadsheet` as the feature name because it describes the editing experience. The core package keeps `excel` because it describes the file/runtime domain.

```text
packages/editor/src/spreadsheet/
@coco-vexel/coco-office-editor/spreadsheet
```

## Product Scope

The goal is an editable online spreadsheet workspace for web and desktop hosts.

First-phase requirements:

- Render worksheets and inspect data.
- Support high-frequency spreadsheet editing operations.
- Export back to `.xlsx`.

Non-goals for the first phase:

- Full Excel feature compatibility.
- A replacement formula engine.
- Pivot table editing.
- Chart/image rendering fidelity.
- Multi-user realtime collaboration.
- A new spreadsheet domain model independent from Coco core.

## Core Data Flow

The runtime owns the workbook truth. The renderer owns only derived caches.

```text
xlsx bytes
  -> ExcelWorkbook.fromZip(bytes, { mode: "auto" })
  -> workbook.toModel()
  -> SpreadsheetRenderModel
  -> Canvas2D renderer

user edit
  -> spreadsheet editor command
  -> ExcelWorkbook live-object API
  -> workbook.toModel()
  -> update render cache
  -> repaint dirty area

export
  -> workbook.toZip()
```

Important rule:

```text
ExcelWorkbook = truth source
ExcelWorkbookModel snapshot = render source
SpreadsheetRenderModel = cache
Canvas = projection
```

The editor should not mutate `ExcelWorkbookModel` snapshots directly. UI commands should call the public `@coco-vexel/coco-office/excel` runtime APIs first, then refresh the render cache from core state.

## Proposed Module Layout

```text
packages/editor/src/spreadsheet/
  index.ts
  runtime.ts
  render-model.ts
  layout.ts
  viewport.ts
  canvas-renderer.ts
  style-resolver.ts
  selection.ts
  editing.ts
  commands.ts
  clipboard.ts
  scrollbar.ts
  painters/
    grid-painter.ts
    cell-painter.ts
    header-painter.ts
    border-painter.ts
    selection-painter.ts
```

Responsibilities:

- `runtime.ts`: binds an `ExcelWorkbook`, tracks active sheet, loads snapshots, exports bytes.
- `render-model.ts`: converts `ExcelWorkbookModel` / `ExcelSheetModel` into indexed render caches.
- `layout.ts`: resolves row/column dimensions, logical cell rects, visible ranges, and hit testing primitives.
- `viewport.ts`: owns scroll state, freeze state, viewport splitting, and coordinate transforms.
- `canvas-renderer.ts`: coordinates Canvas2D layers and painters.
- `style-resolver.ts`: maps core style tables to canvas fill, text, border, and alignment styles.
- `selection.ts`: stores active cell and logical selection ranges.
- `editing.ts`: positions the DOM cell editor and commits/cancels edits.
- `commands.ts`: maps editor commands to Coco Excel runtime APIs.
- `clipboard.ts`: handles TSV copy/paste for ranges.
- `scrollbar.ts`: handles custom scrollbars after the first wheel-based implementation.

## Rendering Technology

Use Canvas2D for the spreadsheet body. Do not start with WebGL.

Canvas2D is the right first implementation because the early complexity is in layout, text measurement, styles, merged cells, hit testing, selection, and editing alignment. WebGL would make text and Excel-style border/layout behavior more expensive to build.

Use DOM for:

- Cell editor.
- Formula bar.
- Toolbar.
- Context menu.
- Dialogs.
- Sheet tabs if convenient.

Use three stacked canvases:

```text
baseCanvas
  Cell fills, gridlines, cell text, borders.

headerCanvas
  Row headers, column headers, corner, active row/column highlights.

overlayCanvas
  Active cell, selection, fill handle, drag marquee, search highlights, agent highlights.
```

Frozen panes should not be implemented by adding many canvas elements. They should be implemented by viewport clipping and coordinate transforms within the existing layers.

## Render Model

The render model is derived from `ExcelWorkbookModel`.

Recommended first cache:

```ts
interface SpreadsheetSheetRenderModel {
  sheetId: string;
  sheetName: string;
  cellsByAddress: Map<string, SpreadsheetRenderCell>;
  cellsByRowColumn: Map<string, SpreadsheetRenderCell>;
  rowMetaByIndex: Map<number, ExcelRowModel>;
  columnMetaByIndex: Map<number, ExcelColumnModel>;
  mergeMasterByAddress: Map<string, ExcelMergedRangeModel>;
  mergeFollowerByAddress: Map<string, string>;
  styles: ExcelStylesModel;
}
```

Cell display should start simple:

- Formula cells display the current cached value from core.
- Number/string/boolean cells display `String(value)`.
- Blank cells display empty text.
- Number formatting can be added through `style-resolver.ts` after the basic renderer is stable.

## Layout Engine

Layout is the core of the renderer.

The sheet has one logical coordinate system:

```text
logicalX(col) = sum(width of columns before col)
logicalY(row) = sum(height of rows before row)
```

First implementation:

- Default row height: `22px`.
- Default column width: `96px`.
- Row header width: `48px`.
- Column header height: `24px`.
- Prefix arrays for row and column offsets.

Support in phase one:

- Custom row heights.
- Custom column widths.
- Hidden rows/columns.
- Merged cell rects.
- Zoom as a later extension point.

Potential later optimization:

- Replace prefix arrays with a Fenwick tree or segment tree when dynamic sizes and very large sheets require faster updates.

## Scrolling And Frozen Panes

Scrolling and frozen panes should share one abstraction: `ViewportManager`.

Mental model:

```text
The worksheet is one large logical plane.
The screen is split into viewports.
Each viewport has a screen rect and a logical origin.
Scroll changes only non-frozen viewport origins.
```

Even without frozen panes, the editor has these logical viewports:

```text
corner
columnHeader
rowHeader
main
```

With frozen rows/columns, cell content is split into:

```text
frozenTopLeft
frozenTop
frozenLeft
main
```

Headers are split similarly:

```text
columnHeaderFrozen
columnHeaderScrollable
rowHeaderFrozen
rowHeaderScrollable
corner
```

First phase support:

- Freeze the first N rows.
- Freeze the first M columns.
- Read basic freeze state from `sheet.sheetViews[0]?.pane`.

First phase exclusions:

- Split panes.
- Bottom/right frozen areas.
- Full Excel `activePane` behavior.
- Pixel-perfect Excel scrollbar behavior.

Selection should always be stored as a logical range, not as viewport-specific ranges.

```text
selection = A1:D10
```

When drawing, intersect the logical selection with each viewport and paint the visible fragments separately.

Hit testing should be viewport-aware:

```text
point
  -> viewport
  -> logical coordinate using viewport origin
  -> row/column
```

Cell editor positioning should use the same viewport coordinate transform. A cell in a frozen viewport stays fixed while editing; a cell in the main viewport follows scroll.

## Editing And Commands

The editor should start with high-frequency operations:

- Click to select a cell.
- Drag to select a range.
- Double-click or type to edit a cell.
- `Enter` commits and moves down.
- `Tab` commits and moves right.
- `Esc` cancels.
- Copy selected range as TSV.
- Paste TSV into a range.
- Wheel scroll.
- Switch sheet tabs.

Commands should call Coco core APIs:

```ts
await sheet.getCell("A1").setValue("hello");
await sheet.getRange("A1:B2").setValues([[1, 2], [3, 4]]);
const bytes = await workbook.toZip();
```

For the first implementation, rebuilding the current sheet render cache after an edit is acceptable. After the editing loop is correct, use core mutation records or model diffs to update only affected cells and repaint dirty regions.

## First-Phase Rendering Coverage

Implement:

- Worksheet grid.
- Sheet tabs.
- Row and column headers.
- Cell text/value display.
- Basic formulas as displayed cached values.
- Basic fonts.
- Fill color.
- Borders.
- Alignment.
- Row heights and column widths.
- Merged cells.
- Active cell and selection overlay.
- DOM cell editing.
- TSV paste.
- `.xlsx` export through core.

Defer:

- Rich text inside cells.
- Full number format fidelity.
- Formula editing assistance and dependency highlights.
- Conditional formatting rendering.
- Images, shapes, and charts.
- Pivot tables.
- Data validation UI.
- Advanced filtering UI.
- Realtime collaboration.

## Milestones

### M1: Read-Only Canvas

- Load an `ExcelWorkbook` from bytes or model.
- Render the first sheet using Canvas2D.
- Draw visible cells, headers, gridlines, and basic text.

### M2: Scroll And Freeze

- Add `ViewportManager`.
- Add wheel-based scroll state.
- Add visible range calculation.
- Support basic `xSplit` / `ySplit` frozen rows and columns.

### M3: Selection

- Add viewport-aware hit testing.
- Add active cell.
- Add drag range selection.
- Paint overlay without repainting base cells.

### M4: Cell Editing

- Add DOM text editor positioned over the active cell.
- Commit edits through `ExcelWorkbook` live-object APIs.
- Refresh current sheet render cache and repaint.

### M5: Clipboard And Export

- Copy selected range as TSV.
- Paste TSV through range APIs.
- Export through `workbook.toZip()`.

### M6: Style Fidelity

- Improve font, fill, border, alignment, row/column size, and merge rendering.
- Add basic number format display.

## Design Principles

1. Core remains the truth source.
2. The spreadsheet editor is a projection and interaction layer.
3. Render caches are disposable and rebuildable.
4. Viewport abstraction owns scrolling, freezing, hit testing, selection drawing, and editor positioning.
5. Canvas2D is the first rendering backend; WebGL is not part of the first plan.
6. First make the editing/export loop correct, then optimize with dirty regions and mutation-driven cache updates.
