# PPT Core Style Capabilities

This document is the core-only style capability matrix for
`@coco-vexel/coco-office/ppt`. It describes SDK model/API support only. Host adapters,
Agent tools, and UI command surfaces are out of scope.

## Status Legend

| Status             | Meaning                                                                 |
| ------------------ | ----------------------------------------------------------------------- |
| `read-write`       | Public live-object API can read and mutate the style.                   |
| `structured-write` | Writable through structured paragraph/run APIs, not a dedicated setter. |
| `read-only`        | Parsed into the normalized model, but not mutated by public APIs.       |
| `planned`          | Candidate for a later core API.                                         |
| `out-of-scope`     | Intentionally not part of core authoring for now.                       |

## P0: Supported Core Writes

These are the style operations already implemented in the core SDK and covered
by package round-trip tests.

| Area       | Capability                                        | Status             | API                                                                                            |
| ---------- | ------------------------------------------------- | ------------------ | ---------------------------------------------------------------------------------------------- |
| Text       | Font size                                         | `read-write`       | `shape.text.setTextStyle({ fontSize })`, `setRunTextStyle(...)`                                |
| Text       | Font family                                       | `read-write`       | `shape.text.setTextStyle({ fontFamily })`, `setRunTextStyle(...)`                              |
| Text       | Text color                                        | `read-write`       | `shape.text.setTextStyle({ color })`, `setRunTextStyle(...)`                                   |
| Text       | Bold                                              | `read-write`       | `shape.text.setTextStyle({ bold })`, `setBold(...)`                                            |
| Text       | Italic                                            | `read-write`       | `shape.text.setTextStyle({ italic })`, `setItalic(...)`                                        |
| Text       | Underline                                         | `read-write`       | `shape.text.setTextStyle({ underline })`, `setRunTextStyle(...)`                               |
| Text       | Strike                                            | `read-write`       | `shape.text.setTextStyle({ strike })`, `setRunTextStyle(...)`                                  |
| Text       | Baseline shift                                    | `read-write`       | `shape.text.setTextStyle({ baseline })`, `setRunTextStyle(...)`                                |
| Text       | Highlight color                                   | `read-write`       | `shape.text.setTextStyle({ highlight })`, `setRunTextStyle(...)`                               |
| Text       | Line height                                       | `read-write`       | `shape.text.setTextStyle({ lineHeight })`, `setParagraphs(...)`                                |
| Text       | Character spacing                                 | `read-write`       | `shape.text.setTextStyle({ characterSpacing })`, `setRunTextStyle(...)`                        |
| Text       | Paragraph alignment                               | `read-write`       | `shape.text.setTextStyle({ alignment })`, `setAlignment(...)`                                  |
| Text       | Body auto-fit, wrap, anchor, vertical, and insets | `read-write`       | `shape.text.setBodyProperties({ autoFit, wrap, anchor, vertical, insetLeft })`                 |
| Text       | Run-level formatting                              | `read-write`       | `shape.text.setRunTextStyle(runIndex, style)`                                                  |
| Text       | Paragraph/run rebuild with formatting             | `structured-write` | `shape.text.setParagraphs(paragraphs)`                                                         |
| Paragraph  | Bullet character, color, font, size               | `structured-write` | `shape.text.setParagraphs([{ style: { bullet, bulletColor, bulletFontFamily, bulletSize } }])` |
| Paragraph  | Spacing and indents                               | `structured-write` | `shape.text.setParagraphs([{ style: { marginLeft, textIndent, spaceBefore, spaceAfter } }])`   |
| Shape      | Solid fill color                                  | `read-write`       | `shape.setFillColor(color)`, `shape.setStyle({ fillColor })`                                   |
| Shape      | Solid fill opacity                                | `read-write`       | `shape.setStyle({ opacity })`                                                                  |
| Shape      | Solid line color and line width                   | `read-write`       | `shape.setLineStyle({ lineColor, lineWidth })`                                                 |
| Shape      | Solid line opacity                                | `read-write`       | `shape.setStyle({ lineOpacity })`                                                              |
| Shape      | Simple outer shadow                               | `read-write`       | `shape.setStyle({ shadow: { color, opacity, blur, distance, direction } })`                    |
| Image      | Crop ratios                                       | `read-write`       | `shape.image.setCrop({ left, top, right, bottom })`, `shape.image.setStyle({ crop })`          |
| Image      | Fill rectangle ratios                             | `read-write`       | `shape.image.setFillRect({ left, top, right, bottom })`, `shape.image.setStyle({ fillRect })`  |
| Image      | Picture opacity                                   | `read-write`       | `shape.image.setOpacity(opacity)`, `shape.image.setStyle({ opacity })`                         |
| Table cell | Cell solid fill                                   | `read-write`       | `cell.setStyle({ fillColor })`                                                                 |
| Table cell | Cell text style                                   | `read-write`       | `cell.setStyle({ textStyle })`                                                                 |
| Table cell | Cell body anchor, wrap, and insets                | `read-write`       | `cell.setStyle({ body: { anchor, wrap, insetLeft, insetTop, insetRight, insetBottom } })`      |
| Table cell | Cell run-level formatting                         | `read-write`       | `cell.setRunStyle(runIndex, style)`                                                            |
| Table cell | Paragraph/run rebuild with formatting             | `structured-write` | `cell.setParagraphs(paragraphs)`                                                               |

`PptImageStyle` contains writable picture `crop`, `fillRect`, and `opacity`
fields.

Colors are normalized to `#RRGGBB`. Aggregate text setters can collapse
multi-run content; use `paragraphs()`, `runs()`, `setParagraphs(...)`, and
`setRunTextStyle(...)` when preserving run boundaries matters.

## P0 Read Coverage

These fields are part of the normalized model and are expected to be
discoverable from existing decks:

| Area       | Model Fields                                                                                                                                       |
| ---------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| Text       | `fontSize`, `fontFamily`, `color`, `bold`, `italic`, `underline`, `strike`, `baseline`, `highlight`, `alignment`, `lineHeight`, `characterSpacing` |
| Paragraph  | `alignment`, `bullet`, `bulletColor`, `bulletFontFamily`, `bulletSize`, spacing, and indents                                                       |
| Shape      | `fillColor`, `fillGradient`, `opacity`, `lineColor`, `lineGradient`, `lineOpacity`, `lineWidth`, `lineCap`, `lineDash`, line ends, `shadow`        |
| Slide      | `background.fillColor`, `background.fillGradient`                                                                                                  |
| Image      | `crop`, `fillRect`, `opacity`                                                                                                                      |
| Table cell | `fillColor`, `fillGradient`, explicit side `borders`, `body`, `textStyle`                                                                          |

Read coverage does not imply a public write API. When a model field is
read-only, core should preserve the source OOXML during unrelated edits.

## P1: Core Style Write Candidates

The current style-focused P1 candidates have been promoted into P0 core writes.
Future P1 style work should start from concrete fixture evidence or a product
workflow, then add a narrow mutation API with round-trip tests.

## P2: Advanced Or Read-Only Style Areas

| Area                        | Current Status      | Notes                                                              |
| --------------------------- | ------------------- | ------------------------------------------------------------------ |
| Theme color/font schemes    | `read-only`         | Theme metadata is available for inspection and inheritance hints.  |
| Master/layout style editing | `out-of-scope`      | Core edits concrete slide/package content first.                   |
| SmartArt style editing      | `out-of-scope`      | SmartArt is detected as advanced content and preserved.            |
| OLE style editing           | `out-of-scope`      | Embedded object internals are not edited by PPT core.              |
| Animation/timing styling    | `read-only`         | Timing and transition summaries are inspection APIs.               |
| Complex effects and 3D      | `planned/read-only` | Add only after concrete fixture evidence and narrow mutation APIs. |

## Testing Requirements

Every new P0 style writer should include:

1. A unit test that writes through the public live-object API.
2. A `toZip()` reload assertion that reads the updated style from the package.
3. A mutation-log assertion when the operation should participate in writeback
   planning.
4. Real PPT fixture coverage when the operation touches existing OOXML patterns
   found under the configured PPT fixture directory.
