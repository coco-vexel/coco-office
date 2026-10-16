# Core PPT API Reference

This file is the public API index for `@coco-vexel/coco-office/ppt`. Detailed
parameters, examples, model fields, runtime notes, and domain rules live in the
domain documents.

Import:

```ts
import { Presentation } from '@coco-vexel/coco-office/ppt';
import type {
  PptPresentationModel,
  PptSlideModel,
  PptShapeModel,
  PptMutation,
  PptWritebackPlan,
  PptHostWritebackProfile,
} from '@coco-vexel/coco-office/ppt';
```

`Presentation` is the public root object. It is async by design and is intended
to expose the same live-object API in inline and future worker modes.

## API Contract

| Rule              | Contract                                                                                                                     |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| Public root       | Use `Presentation` from `@coco-vexel/coco-office/ppt`.                                                                             |
| Runtime mode      | `mode: 'auto' \| 'inline' \| 'worker'` exists in options. Current implementation resolves to inline; worker mode is planned. |
| Live handles      | `getItem(...)`, `getItemAt(...)`, `getCell(...)`, and collection properties return lightweight handles synchronously.        |
| Reads and writes  | Methods that read, mutate, serialize, search, check existence, or return snapshots are async. Always `await` them.           |
| Snapshot boundary | `toModel()` returns a defensive snapshot. Mutating it does not mutate the deck.                                              |
| Package boundary  | `toZip()` exports `.pptx` bytes for host writeback or file output.                                                           |
| Mutation boundary | `getMutations()` and `clearMutations()` are host/runtime APIs for writeback optimization, not normal authoring APIs.         |
| Low-level core    | `*Core` exports are for tests and adapters. Ordinary SDK callers should use non-Core live objects.                           |
| Host APIs         | Office.js/WPSJS/ONLYOFFICE/Desktop APIs are adapter concerns, not core SDK operation code.                                   |

### Handle Versus Operation

Live handles are intentionally cheap and synchronous. They identify a slide,
shape, or table cell, but they do not fetch data by themselves:

```ts
const slide = presentation.slides.getItem(0); // sync handle
const shape = slide.shapes.getItem('Title'); // sync handle
const cell = tableShape.table.getCell(0, 1); // sync handle

const title = await shape.text.getText(); // async read
await shape.text.setText(`${title} - Final`); // async write
```

If a method name starts with `getItem` and does not include `OrNullObject`, it
returns a handle and the operation may fail when a later async method resolves
the target. Use `getItemOrNullObject(...)` when absence is an expected branch.

## Object Model

| Domain Property       | Main Types                               | Full Documentation                       |
| --------------------- | ---------------------------------------- | ---------------------------------------- |
| `presentation.slides` | `SlideCollection`, `Slide`               | [domains/slides.md](./domains/slides.md) |
| `slide.shapes`        | `ShapeCollection`, `Shape`               | [domains/shapes.md](./domains/shapes.md) |
| `shape.text`          | `ShapeTextApi`                           | [domains/text.md](./domains/text.md)     |
| `shape.image`         | `ShapeImageApi`                          | [domains/images.md](./domains/images.md) |
| `shape.table`         | `TableApi`, `TableCell`                  | [domains/tables.md](./domains/tables.md) |
| Search methods        | `PptTextMatch`                           | [domains/search.md](./domains/search.md) |
| Snapshot/model APIs   | `PptPresentationModel`, properties types | [model.md](./model.md)                   |

## Common Usage Patterns

### Load, Edit, Export

Use this pattern in tests, file conversion flows, and package-backed runtimes:

```ts
const presentation = await Presentation.fromZip(bytes, { presentationId: 'deck.pptx' });
const slide = presentation.slides.getItem(0);
await slide.shapes.addTextBox('Reviewed', {
  name: 'ReviewStatus',
  geometry: { left: 500000, top: 500000, width: 1800000, height: 500000 },
});
const output = await presentation.toZip();
await presentation.dispose();
```

SDK operation code can receive a ready `presentation: Presentation` from the
runtime:

```ts
export async function run(presentation: Presentation) {
  const result = await presentation.replaceText('Draft', 'Final');
  return { replacements: result.count };
}
```

### Find Before Editing

When the target slide or shape is unknown, search first and route by match
source:

```ts
for (const match of await presentation.findText('TODO')) {
  const slide = presentation.slides.getItem(match.slideId);
  const shape = slide.shapes.getItem(match.shapeId);
  if (match.source === 'shapeText') {
    await shape.text.replace('TODO', 'Done');
  } else {
    await shape.table.setCellText(match.row!, match.column!, 'Done');
  }
}
```

### Preserve Formatting

Use aggregate setters for simple replacement. They may collapse multi-run text
content. Use paragraph/run APIs when formatting matters:

```ts
const runs = await shape.text.runs();
if (runs.length > 0) {
  await shape.text.setRunTextStyle(0, { bold: true, color: '#1F4E79' });
}
```

### Work With Existing Decks Conservatively

Real PPT files can contain charts, media, OLE objects, SmartArt, groups,
placeholders, theme inheritance, and animations. The SDK identifies many of
these objects, but only common text/image/table/shape operations are editable in
core. For unknown or advanced shapes, read metadata through `toModel()` and use
supported sibling edits instead of raw OOXML edits.

## Common Public APIs

### Presentation

| API                                                    | Returns                              | Notes                                                                  |
| ------------------------------------------------------ | ------------------------------------ | ---------------------------------------------------------------------- |
| `Presentation.fromZip(bytes, opts?)`                   | `Promise<Presentation>`              | Preferred when a host can provide `.pptx` bytes.                       |
| `Presentation.fromModel(model, opts?)`                 | `Promise<Presentation>`              | Builds a presentation package from a normalized model.                 |
| `Presentation.createBlank(opts?)`                      | `Promise<Presentation>`              | Creates a one-slide blank presentation.                                |
| `presentation.toModel()`                               | `Promise<PptPresentationModel>`      | Defensive snapshot.                                                    |
| `presentation.properties()`                            | `Promise<PptPresentationProperties>` | Counts slides, shapes, text shapes, tables, pictures, package media.   |
| `presentation.findText(query, opts?)`                  | `Promise<PptTextMatch[]>`            | Searches text shapes and table cells across the deck.                  |
| `presentation.replaceText(search, replacement, opts?)` | `Promise<PptTextReplaceResult>`      | Replaces across text shapes and table cells.                           |
| `presentation.exportImages()`                          | `Promise<PptExportedImage[]>`        | Returns package media bytes from `ppt/media/*`.                        |
| `presentation.exportFonts()`                           | `Promise<PptExportedFont[]>`         | Returns package font bytes discovered from presentation relationships. |
| `presentation.themes.items()` / `getCount()`           | `Promise<...>`                       | Theme metadata, color scheme, and font scheme.                         |
| `presentation.getMutations()`                          | `Promise<PptMutation[]>`             | Host/runtime writeback API.                                            |
| `presentation.clearMutations()`                        | `Promise<void>`                      | Clears mutation log after baseline/writeback.                          |
| `presentation.toZip()`                                 | `Promise<Uint8Array>`                | Exports `.pptx`.                                                       |
| `presentation.dispose()`                               | `Promise<void>`                      | Releases runtime state.                                                |

### Slides

| API                                                  | Returns                                    | Notes                                              |
| ---------------------------------------------------- | ------------------------------------------ | -------------------------------------------------- |
| `presentation.slides.items()`                        | `Promise<Slide[]>`                         | Lists slide handles in visual order.               |
| `presentation.slides.getCount()`                     | `Promise<number>`                          | Slide count.                                       |
| `presentation.slides.getItem(indexOrId)`             | `Slide`                                    | 0-based index or slide id/name.                    |
| `presentation.slides.getItemOrNullObject(indexOrId)` | `Promise<Slide \| null>`                   | Existence check.                                   |
| `presentation.slides.add(options?)`                  | `Promise<Slide>`                           | Adds a slide, optionally with title and index.     |
| `presentation.slides.copy(indexOrId, options?)`      | `Promise<Slide>`                           | Copies a slide part and relationship part.         |
| `presentation.slides.move(indexOrId, toIndex)`       | `Promise<void>`                            | Moves to 0-based index; destination is clamped.    |
| `presentation.slides.deleteItem(indexOrId)`          | `Promise<void>`                            | Deletes slide; final slide deletion is rejected.   |
| `slide.id()` / `index()` / `name()` / `title()`      | `Promise<...>`                             | Basic metadata.                                    |
| `slide.layout()`                                     | `Promise<PptSlideLayoutModel \| null>`     | Slide layout/master relationship metadata.         |
| `slide.transition()`                                 | `Promise<PptSlideTransitionModel \| null>` | Slide transition read-only metadata.               |
| `slide.timing()`                                     | `Promise<PptSlideTimingModel \| null>`     | Animation timing read-only summary.                |
| `slide.properties()`                                 | `Promise<PptSlideProperties>`              | Per-slide counts and title.                        |
| `slide.setTitle(text, opts?)`                        | `Promise<void>`                            | Updates first text-like title shape.               |
| `slide.notes.hasText()` / `getText()`                | `Promise<...>`                             | Speaker notes text.                                |
| `slide.notes.setText(text)` / `clear()`              | `Promise<void>`                            | Writes or clears speaker notes.                    |
| `slide.comments.items()` / `getCount()`              | `Promise<...>`                             | Slide comment list.                                |
| `slide.comments.add(options)`                        | `Promise<PptCommentModel>`                 | Adds a slide comment and comment author if needed. |
| `slide.comments.delete(id)` / `clear()`              | `Promise<void>`                            | Deletes one or all slide comments.                 |
| `slide.findText(...)` / `replaceText(...)`           | `Promise<...>`                             | Slide-limited search and replace.                  |
| `slide.duplicate(options?)`                          | `Promise<Slide>`                           | Copies this slide.                                 |
| `slide.moveTo(index)`                                | `Promise<void>`                            | Moves this slide.                                  |
| `slide.delete()`                                     | `Promise<void>`                            | Deletes this slide.                                |
| `slide.toModel()`                                    | `Promise<PptSlideModel>`                   | Defensive slide snapshot.                          |

### Shapes

| API                                                      | Returns                              | Notes                                                    |
| -------------------------------------------------------- | ------------------------------------ | -------------------------------------------------------- |
| `slide.shapes.items()`                                   | `Promise<Shape[]>`                   | Lists shapes in shape-tree order, which is z-order.      |
| `slide.shapes.getCount()`                                | `Promise<number>`                    | Shape count.                                             |
| `slide.shapes.getItem(idOrName)`                         | `Shape`                              | Ambiguous duplicated names throw.                        |
| `slide.shapes.getItemAt(index)`                          | `Shape`                              | 0-based shape-tree index.                                |
| `slide.shapes.getItemOrNullObject(idOrName)`             | `Promise<Shape \| null>`             | Existence check.                                         |
| `slide.shapes.addTextBox(text, options?)`                | `Promise<Shape>`                     | Adds text box.                                           |
| `slide.shapes.addTable(rowCount, columnCount, options?)` | `Promise<Shape>`                     | Adds simple table.                                       |
| `slide.shapes.addImage(source, options?)`                | `Promise<Shape>`                     | Adds picture and media relationship.                     |
| `slide.shapes.addShape(preset, options?)`                | `Promise<Shape>`                     | Adds `rect`, `roundRect`, `ellipse`, or `line`.          |
| `slide.shapes.deleteItem(idOrName)`                      | `Promise<void>`                      | Deletes shape.                                           |
| `shape.id()` / `name()` / `kind()`                       | `Promise<...>`                       | Basic shape metadata.                                    |
| `shape.setName(name)`                                    | `Promise<void>`                      | Updates `p:cNvPr/@name`.                                 |
| `shape.geometry()` / `setGeometry(...)`                  | `Promise<...>`                       | Reads/writes `a:xfrm`.                                   |
| `shape.style()` / `setStyle(...)`                        | `Promise<...>`                       | Shape fill, line, opacity, and shadow style.             |
| `shape.hyperlink()`                                      | `Promise<PptHyperlinkModel \| null>` | Shape-level click hyperlink.                             |
| `shape.setHyperlink(options)` / `clearHyperlink()`       | `Promise<void>`                      | Writes/removes shape-level click hyperlink relationship. |
| `shape.toModel().placeholder`                            | `PptPlaceholderModel \| undefined`   | Placeholder type/index/size/orientation snapshot.        |
| `shape.setFillColor(color)`                              | `Promise<void>`                      | Convenience fill setter.                                 |
| `shape.setLineStyle(style)`                              | `Promise<void>`                      | Convenience line style setter.                           |
| `shape.bringToFront()` / `sendToBack()`                  | `Promise<void>`                      | Moves shape in `p:spTree`.                               |
| `shape.duplicate(options?)`                              | `Promise<Shape>`                     | Duplicates this shape on the same slide.                 |
| `shape.delete()`                                         | `Promise<void>`                      | Deletes this shape.                                      |
| `shape.toModel()`                                        | `Promise<PptShapeModel>`             | Defensive shape snapshot.                                |

### Charts

| API                         | Returns                          | Notes                                                              |
| --------------------------- | -------------------------------- | ------------------------------------------------------------------ |
| `shape.chart.isChart()`     | `Promise<boolean>`               | True for chart shapes with chart metadata.                         |
| `shape.chart.getInfo()`     | `Promise<PptChartModel \| null>` | Relationship, part path, type, title, and series names.            |
| `shape.chart.getData()`     | `Promise<PptChartData \| null>`  | Title, series names, categories, and values from chart XML caches. |
| `shape.chart.setData(data)` | `Promise<void>`                  | Replaces existing chart XML cache data; series count must match.   |

### Media, OLE, And SmartArt

| API                         | Returns                           | Notes                                                                 |
| --------------------------- | --------------------------------- | --------------------------------------------------------------------- |
| `shape.kind()`              | `Promise<PptShapeKind>`           | Returns `media`, `oleObject`, `smartArt`, or `unknown` when detected. |
| `shape.toModel().media`     | `PptShapeMediaModel \| undefined` | Audio/video/media relationship metadata, read-only.                   |
| `shape.toModel().oleObject` | `PptOleObjectModel \| undefined`  | OLE relationship, target path, and `progId`, read-only.               |
| `shape.toModel().smartArt`  | `PptSmartArtModel \| undefined`   | Diagram relationship ids and target paths, read-only.                 |
| `shape.toModel().raw`       | `PptShapeModel['raw']`            | Unsupported host/geometry/group metadata, read-only.                  |

### Groups

| API                                                  | Returns                   | Notes                                    |
| ---------------------------------------------------- | ------------------------- | ---------------------------------------- |
| `shape.group.isGroup()`                              | `Promise<boolean>`        | True when `shape.kind()` is `group`.     |
| `shape.group.items()`                                | `Promise<Shape[]>`        | Direct child shapes in group-tree order. |
| `shape.group.getCount()`                             | `Promise<number>`         | Direct child shape count.                |
| `shape.group.getItem(idOrName)`                      | `Shape`                   | Looks up a direct child shape.           |
| `shape.group.getItemAt(index)`                       | `Shape`                   | 0-based direct child index.              |
| `shape.group.getItemOrNullObject(idOrName)`          | `Promise<Shape \| null>`  | Direct child existence check.            |
| `PptShapeModel.children`                             | `PptShapeModel[]`         | Snapshot children for `kind: 'group'`.   |
| `presentation.findText(...)` / `slide.findText(...)` | `Promise<PptTextMatch[]>` | Includes text inside group child shapes. |

### Text

| API                                              | Returns                          | Notes                                                               |
| ------------------------------------------------ | -------------------------------- | ------------------------------------------------------------------- |
| `shape.text.hasText()`                           | `Promise<boolean>`               | True when shape has text runs.                                      |
| `shape.text.getText()`                           | `Promise<string>`                | Aggregated shape text.                                              |
| `shape.text.paragraphs()`                        | `Promise<PptTextParagraph[]>`    | Paragraph/run structure.                                            |
| `shape.text.runs()`                              | `Promise<PptTextRun[]>`          | Flattened text runs, including run hyperlinks.                      |
| `shape.text.effectiveRuns()`                     | `Promise<PptTextRun[]>`          | Runs with inherited master, layout, shape, paragraph, and run style. |
| `shape.text.body()`                              | `Promise<PptTextBodyProperties>` | Text body anchor, auto-fit, wrap, vertical mode, and insets.        |
| `shape.text.setParagraphs(paragraphs)`           | `Promise<void>`                  | Rebuilds structured paragraphs and runs.                            |
| `shape.text.setBodyProperties(body)`             | `Promise<void>`                  | Updates body properties such as auto-fit without rebuilding text.   |
| `shape.text.setTextStyle(style)`                 | `Promise<void>`                  | Preferred whole-shape text style update.                            |
| `shape.text.setRunTextStyle(index, style)`       | `Promise<void>`                  | Preferred explicit run style update.                                |
| `shape.text.setRunStyle(index, style)`           | `Promise<void>`                  | Compatibility alias for `setRunTextStyle`.                          |
| `shape.text.setRunHyperlink(index, options)`     | `Promise<void>`                  | Writes hyperlink on one flattened text run.                         |
| `shape.text.clearRunHyperlink(index)`            | `Promise<void>`                  | Removes hyperlink from one flattened text run.                      |
| `shape.text.setText(text, opts?)`                | `Promise<void>`                  | Writes text to first run and clears extra runs.                     |
| `shape.text.replace(search, replacement, opts?)` | `Promise<PptTextReplaceResult>`  | Replaces inside one shape.                                          |
| `shape.text.style()`                             | `Promise<PptTextStyle>`          | Font, color, emphasis, baseline, highlight, alignment, and spacing. |
| `shape.text.effectiveStyle()`                    | `Promise<PptTextStyle>`          | Effective style after master, layout, and direct-shape inheritance. |
| `shape.text.setStyle(style)`                     | `Promise<void>`                  | Compatibility alias for `setTextStyle`.                             |
| `shape.text.setBold(enabled)`                    | `Promise<void>`                  | Convenience setter.                                                 |
| `shape.text.setItalic(enabled)`                  | `Promise<void>`                  | Convenience setter.                                                 |
| `shape.text.setAlignment(alignment)`             | `Promise<void>`                  | `left`, `center`, `right`, `justify`, `distributed`.                |

### Images

| API                             | Returns                         | Notes                                                             |
| ------------------------------- | ------------------------------- | ----------------------------------------------------------------- |
| `shape.image.isImage()`         | `Promise<boolean>`              | True for picture shapes with media target.                        |
| `shape.image.getInfo()`         | `Promise<PptImageInfo \| null>` | Metadata and byte length.                                         |
| `shape.image.style()`           | `Promise<PptImageStyle>`        | Picture crop, fill rectangle, and opacity snapshot.               |
| `shape.image.setStyle(style)`   | `Promise<void>`                 | Partial crop/fillRect/opacity style update.                       |
| `shape.image.setCrop(crop)`     | `Promise<void>`                 | Updates picture crop ratios.                                      |
| `shape.image.setFillRect(rect)` | `Promise<void>`                 | Updates picture fill rectangle ratios.                            |
| `shape.image.setOpacity(n)`     | `Promise<void>`                 | Updates picture opacity ratio.                                    |
| `shape.image.bytes()`           | `Promise<Uint8Array>`           | Raw image bytes.                                                  |
| `shape.image.replace(source)`   | `Promise<void>`                 | Replaces media bytes in-place; preserves media path/content type. |
| `presentation.exportImages()`   | `Promise<PptExportedImage[]>`   | Global package media export.                                      |

### Tables

| API                                          | Returns                        | Notes                                                                         |
| -------------------------------------------- | ------------------------------ | ----------------------------------------------------------------------------- |
| `shape.table.isTable()`                      | `Promise<boolean>`             | True when shape is table.                                                     |
| `shape.table.rowCount()` / `columnCount()`   | `Promise<number>`              | Table dimensions.                                                             |
| `shape.table.values()`                       | `Promise<string[][]>`          | Text matrix.                                                                  |
| `shape.table.hasMergedCells()`               | `Promise<boolean>`             | True when OOXML merge attributes are present.                                 |
| `shape.table.mergedCells()`                  | `Promise<PptTableCellMerge[]>` | Merge metadata for cells with span/merge attrs.                               |
| `shape.table.setValues(values)`              | `Promise<void>`                | Matrix must match dimensions.                                                 |
| `shape.table.setCellText(row, column, text)` | `Promise<void>`                | Direct cell text setter.                                                      |
| `shape.table.getCell(row, column)`           | `TableCell`                    | Cell handle.                                                                  |
| `shape.table.addRow(index?, values?)`        | `Promise<void>`                | Inserts row; values length must equal column count; numeric index is clamped. |
| `shape.table.deleteRow(index)`               | `Promise<void>`                | Deletes row; final row deletion rejected.                                     |
| `shape.table.addColumn(index?, values?)`     | `Promise<void>`                | Inserts column; values length must equal row count; numeric index is clamped. |
| `shape.table.deleteColumn(index)`            | `Promise<void>`                | Deletes column; final column deletion rejected.                               |
| `cell.text()` / `setText(text)`              | `Promise<...>`                 | Cell text read/write.                                                         |
| `cell.paragraphs()` / `runs()`               | `Promise<...>`                 | Cell paragraph/run structure.                                                 |
| `cell.setParagraphs(paragraphs)`             | `Promise<void>`                | Rebuilds structured cell text.                                                |
| `cell.setRunStyle(index, style)`             | `Promise<void>`                | Updates one flattened cell run style.                                         |
| `cell.style()` / `setStyle(style)`           | `Promise<...>`                 | Cell fill, body, and text style.                                              |

### Writeback Planning

| API                                                        | Returns                       | Notes                                                                                                                       |
| ---------------------------------------------------------- | ----------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `diffPptPresentationModels(before, after)`                 | `PptPresentationDiff`         | Classifies presentation, slide, and shape changes between normalized snapshots.                                             |
| `planPptWriteback(before, after, capabilities, options?)`  | `PptWritebackPlan`            | Chooses `no-op`, `presentation-package-replace`, `slide-package-replace`, `mutation-replay`, or `unsupported`.              |
| `options.mutations`                                        | `PptMutation[]`               | Required when planning `mutation-replay`. Usually comes from `presentation.getMutations()`.                                 |
| `capabilities.supportedMutationTypes` / `maxMutationCount` | host capability controls      | Allows adapters to restrict replay to proven mutation kinds and operation counts.                                           |
| `createPptHostWritebackProfile(host, options?)`            | `PptHostWritebackProfile`     | Creates a conservative host capability profile for Microsoft PowerPoint, WPS, ONLYOFFICE, desktop package, or custom hosts. |
| `createPptHostWritebackProfileFromProbeReport(report)`     | `PptHostWritebackProfile`     | Derives a probed profile from a host probe JSON report.                                                                     |
| `runPptHostProbe(adapter, options?)`                       | `Promise<PptHostProbeReport>` | Runs a host probe adapter and collects plain JSON check results.                                                            |
| `getDefaultPptHostProbeChecks(host)`                       | `PptHostProbeCheckName[]`     | Returns the default checklist for live hosts or desktop package writeback.                                                  |
| `summarizePptHostProbeReport(report)`                      | `PptHostProbeSummary`         | Lists passed/failed probe checks and derived capabilities.                                                                  |
| `getPptWritebackOperationRequirements(operation)`          | capability names              | Maps a planner operation to the capabilities required to execute it.                                                        |
| `validatePptWritebackPlan(plan, profileOrCapabilities)`    | validation result             | Checks a plan against declared host capabilities and replay mutation allow-lists.                                           |

## SDK Operation Rules

These rules are normative for code that uses the public PPT SDK surface:

1. Import public operation APIs from `@coco-vexel/coco-office/ppt`.
2. Operation code should receive a ready `presentation: Presentation`; it should not initialize host bridges.
3. Use non-Core live-object classes (`Presentation`, `Slide`, `Shape`, `TableCell`) in ordinary SDK code.
4. Treat all reads/writes/search/export/existence checks as async.
5. Use 0-based indexes for slides, shapes, rows, and columns.
6. Prefer `findText(...)` before editing unknown content.
7. Prefer `shape.text.setText(...)`, `presentation.replaceText(...)`, or `slide.replaceText(...)` for text operations; do not edit OOXML directly.
8. Prefer table APIs for table edits; do not infer table XML structure in operation code.
9. Use `addImage(...)` with `Uint8Array`, `ArrayBuffer`, `{ bytes, mimeType, filename }`, or `{ base64, mimeType, filename }`.
10. Do not call Office.js/WPSJS/ONLYOFFICE APIs from core operation code.
11. Do not call `getMutations()` or `clearMutations()` from ordinary SDK operation code.
12. Use `toZip()` only when the task explicitly needs output package bytes.
13. Do not call `planPptWriteback(...)` or `diffPptPresentationModels(...)` from ordinary SDK operation code; these are runtime/writeback orchestration APIs.
14. Treat host profiles as adapter declarations. Microsoft PowerPoint, WPS, and ONLYOFFICE default profiles are intentionally `unproven` until runnable probes enable capabilities.
15. Probe reports must be plain JSON. Do not include native host objects in `details`.
16. When preserving layout matters, prefer targeted edits on existing shapes over rebuilding slides from `toModel()`.
17. When table cells may be merged, inspect `hasMergedCells()` / `mergedCells()` before structural row or column edits.
18. When adding visual objects, provide explicit `name` and `geometry` so later steps can retrieve and adjust them deterministically.

## Domain Retrieval Map

| User Intent                                                             | Retrieve                                      |
| ----------------------------------------------------------------------- | --------------------------------------------- |
| "load pptx", "export pptx", "deck properties", "image export"           | `domains/presentation.md`                     |
| "slide", "add slide", "copy slide", "move slide", "delete slide"        | `domains/slides.md`                           |
| "shape", "text box", "rectangle", "line", "z-order", "fill", "geometry" | `domains/shapes.md`                           |
| "style", "font", "color", "fill", "line", "opacity", "shadow"           | `style-capabilities.md`, then domain doc      |
| "text", "font", "bold", "italic", "alignment", "replace text"           | `domains/text.md`, `domains/search.md`        |
| "image", "picture", "extract image", "replace image"                    | `domains/images.md`                           |
| "table", "cell", "row", "column", "cell style"                          | `domains/tables.md`                           |
| "find text", "search", "replace across deck"                            | `domains/search.md`                           |
| "model", "JSON snapshot", "properties type", "fixture"                  | `model.md`                                    |
| "worker", "execute PPT code", "runtime mode"                            | `runtime.md`                                  |
| "host bridge", "PowerPoint/WPS/ONLYOFFICE writeback"                    | `host-integration.md`                         |
| "mutation replay", "fast writeback", "recorded SDK operations"          | `mutation-writeback.md`                       |
| "writeback plan", "model diff", "package replacement", "slide replay"   | `writeback-planner.md`                        |
| "host capability", "capability profile", "validate writeback operation" | `writeback-planner.md`, `host-integration.md` |

## Current Status

The SDK has focused unit tests and real `.pptx` regression tests under
`packages/core/src/ppt/__tests__/presentation.test.ts`. Real fixtures from
`<maintainer-ppt-corpus>` cover parsing, round-trip, SDK writes, and existing
content edits when the fixture contains editable text/table/image content.
