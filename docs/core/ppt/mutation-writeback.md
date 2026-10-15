# PPT Mutation Writeback

Use this document when a task mentions recorded operations, fast host writeback,
mutation replay, or adapter planning.

## Purpose

Mutation APIs let host adapters inspect SDK-recorded operations and decide
whether a small edit can be replayed in a live host before falling back to full
`.pptx` package replacement.

Use [writeback-planner.md](./writeback-planner.md) for the formal
`diffPptPresentationModels(...)` and `planPptWriteback(...)` contract.

Ordinary SDK operation code should normally not call mutation APIs.

## APIs

```ts
const mutations = await presentation.getMutations();
await presentation.clearMutations();
```

Host runtime pattern:

```ts
const baseline = await presentation.toModel();
await presentation.clearMutations();

// Execute SDK operation code.

const after = await presentation.toModel();
const mutations = await presentation.getMutations();
// Host validates baseline/after/mutations before live-host replay.
```

Planner pattern:

```ts
const plan = planPptWriteback(
  baseline,
  after,
  { canReplayMutations: true, supportedMutationTypes: ['shape.text.set'] },
  { preference: 'mutation', mutations },
);
```

## Mutation Types

Current `PptMutation` records include:

- `shape.add`
- `shape.delete`
- `shape.name.set`
- `shape.zOrder.set`
- `shape.geometry.set`
- `shape.style.set`
- `shape.hyperlink.set`
- `shape.hyperlink.clear`
- `shape.text.set`
- `shape.text.style.set`
- `shape.text.body.set` — applies `body: Partial<PptTextBodyProperties>` to the text container.
- `shape.text.paragraphs.set`
- `shape.text.run.style.set`
- `shape.text.run.hyperlink.set`
- `shape.text.run.hyperlink.clear`
- `shape.text.replace`
- `slide.notes.set`
- `slide.notes.clear`
- `slide.comment.add`
- `slide.comment.delete`
- `slide.comments.clear`
- `shape.chart.data.set`
- `shape.image.replace`
- `shape.image.style.set`
- `table.cell.text.set`
- `table.cell.style.set`
- `table.cell.text.paragraphs.set`
- `table.cell.text.run.style.set`
- `table.values.set`
- `table.row.add`
- `table.row.delete`
- `table.column.add`
- `table.column.delete`

All mutation payloads are plain JSON. Listing a mutation here means the SDK can
record it; a host adapter may replay it only when that mutation type is present
in the adapter's `supportedMutationTypes` and the baseline validation passes.
If `supportedMutationTypes` is omitted, the planner treats the allow-list as
open, so adapters should omit it only after their own validation layer can safely
handle every mutation they pass into the planner.

## Payload Summary

Common fields:

| Field     | Meaning                                         |
| --------- | ----------------------------------------------- |
| `slideId` | SDK slide id from the baseline presentation.    |
| `shapeId` | SDK shape id within a slide, when shape-scoped. |
| `row`     | Zero-based table row index.                     |
| `column`  | Zero-based table column index.                  |

Mutation-specific payloads:

| Type                             | Extra Payload                                        |
| -------------------------------- | ---------------------------------------------------- |
| `shape.add`                      | `kind`, optional `imageRef`                          |
| `shape.delete`                   | none                                                 |
| `shape.name.set`                 | `name`                                               |
| `shape.zOrder.set`               | `position: 'front' \| 'back'`                        |
| `shape.geometry.set`             | `geometry: Partial<PptShapeGeometry>`                |
| `shape.style.set`                | `style: Partial<PptShapeStyle>`                      |
| `shape.hyperlink.set`            | `hyperlink: PptHyperlinkWriteOptions`                |
| `shape.hyperlink.clear`          | none                                                 |
| `shape.text.set`                 | `text`                                               |
| `shape.text.style.set`           | `style: Partial<PptTextStyle>`                       |
| `shape.text.paragraphs.set`      | `paragraphs: PptTextParagraphWrite[]`                |
| `shape.text.run.style.set`       | `runIndex`, `style: Partial<PptTextStyle>`           |
| `shape.text.run.hyperlink.set`   | `runIndex`, `hyperlink: PptHyperlinkWriteOptions`    |
| `shape.text.run.hyperlink.clear` | `runIndex`                                           |
| `shape.text.replace`             | `search`, `replacement`, `count`                     |
| `slide.notes.set`                | `text`                                               |
| `slide.notes.clear`              | none                                                 |
| `slide.comment.add`              | `comment: PptCommentAddOptions`                      |
| `slide.comment.delete`           | `commentId`                                          |
| `slide.comments.clear`           | none                                                 |
| `shape.chart.data.set`           | `data: PptChartData`                                 |
| `shape.image.replace`            | `imageRef`                                           |
| `shape.image.style.set`          | `style: Partial<PptImageStyle>`                      |
| `table.cell.text.set`            | `row`, `column`, `text`                              |
| `table.cell.style.set`           | `row`, `column`, `style: Partial<PptTableCellStyle>` |
| `table.cell.text.paragraphs.set` | `row`, `column`, `paragraphs`                        |
| `table.cell.text.run.style.set`  | `row`, `column`, `runIndex`, `style`                 |
| `table.values.set`               | `values: string[][]`                                 |
| `table.row.add`                  | `index`, `values: string[]`                          |
| `table.row.delete`               | `index`                                              |
| `table.column.add`               | `index`, `values: string[]`                          |
| `table.column.delete`            | `index`                                              |

`imageRef` is an SDK/package reference for bytes already written into the PPT
package. A host adapter must map it to a host-native image replacement path or
fall back to package replacement.

## Adapter Guidance

Adapters may use mutation logs for optimization only after verifying:

1. the live host document still matches the SDK baseline;
2. every mutation kind is supported by that host adapter;
3. indexes, ids, relationship ids, and media paths are still valid;
4. fallback to package replacement is available when replay is unsafe.

Probe-derived profiles expose only a conservative replay subset. Adapters can
declare additional `supportedMutationTypes` manually, but those declarations
should be backed by host-specific tests.

Mutation logs are not an audit log or collaboration log. They are scoped to the
loaded SDK instance and should be cleared after a host runtime accepts a new
baseline.

## SDK Operation Rules

1. Do not use `getMutations()` or `clearMutations()` in ordinary SDK operation code.
2. Use mutations only in host runtime/writeback code.
3. Mutation replay is an optimization, not the source of truth.
4. `toZip()` remains the safe full writeback boundary.
5. Do not infer user intent from mutation payloads; use them only to replay SDK operations that already ran.
6. Do not construct `PptMutation` objects by hand in ordinary SDK operation code.
