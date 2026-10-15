# WordPlan Advanced API

`WordPlan` is the low-level operation builder returned by `doc.plan(opts?)`.

For small-model code generation, prefer live-object domain APIs in `docs/core/domains/`. Use `WordPlan` only when a caller explicitly needs low-level queued operations, cross-object batching, or APIs not yet exposed as live-object methods.

## Create A Plan

```ts
const plan = doc.plan({
  trackChanges: false,
  author: 'Coco',
  date: new Date(),
});
```

`PlanOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `trackChanges` | `boolean` | `false` | Request tracked-change markup when supported. |
| `author` | `string` | document default author | Author for operations that create review metadata. |
| `date` | `Date` | operation time | Date for operations that create review metadata. |

## Plan State

| Property | Type | Description |
| --- | --- | --- |
| `plan.ops` | `ReadonlyArray<Op>` | Queued operations. |
| `plan.applied` | `boolean` | Whether `apply()` has been called. |
| `plan.options` | readonly normalized options | Plan creation options. |

`apply()` is one-shot. Do not append operations or call `apply()` again after a plan has been applied.

## Paragraph Operations

| API | Purpose |
| --- | --- |
| `replaceParagraph(paraId, text, opts?)` | Replace a paragraph with plain text. |
| `replaceTextInParagraph(paraId, search, replacement, opts?)` | Replace matched text inside one paragraph. |
| `formatTextInParagraph(paraId, search, format, opts?)` | Format matched text inside one paragraph. |
| `formatParagraph(paraId, format)` | Set direct paragraph formatting. |
| `insertParagraphBefore(anchorParaId, spec)` | Insert a paragraph before an anchor paragraph. |
| `insertParagraphAfter(anchorParaId, spec)` | Insert a paragraph after an anchor paragraph. |
| `insertParagraphAtBodyStart()` | Insert a paragraph at the start of the body. |
| `deleteParagraph(paraId, opts?)` | Delete a paragraph. |

Common params:

| Param | Type | Description |
| --- | --- | --- |
| `paraId` / `anchorParaId` | `string` | Paragraph id. |
| `text` | `string` | Plain text. |
| `opts.scope` | `'first' \| 'all'` | Match scope for text operations. |
| `opts.onAnchorConflict` | `'error' \| 'drop'` | How to handle affected comments/bookmarks. |

`ParagraphDirectFormat`:

| Field | Type |
| --- | --- |
| `alignment` | `'left' \| 'center' \| 'right' \| 'both' \| 'distribute'` |
| `outlineLevel` | `number` |
| `pageBreakBefore` | `boolean` |
| `indent` | `{left?:number; right?:number; firstLine?:number}` |
| `spacing` | `{before?:number; after?:number; line?:number; lineRule?:'auto'\|'exact'\|'atLeast'}` |

`formatTextInParagraph` format:

| Field | Type |
| --- | --- |
| `bold` | `boolean` |
| `color` | `string` |
| `italic` | `boolean` |
| `underline` | `boolean` |
| `highlight` | `string` |
| `fontSize` | `number` |
| `fontFamily` | `string` |

`InsertParagraphSpec`:

| Field | Type | Description |
| --- | --- | --- |
| `text` | `string` | Inserted paragraph text. |
| `style` | `'inherit' \| string` | Use inherited or explicit paragraph style. |
| `runFormat` | `'inherit' \| 'default'` | Inserted run formatting policy. |
| `color` | `string?` | Optional text color. |
| `styleSource` | `string?` | Source paragraph id for style copying. |
| `bookmark` | `{prefix:string}?` | Optional generated bookmark prefix. |
| `outlineLevel` | `number?` | Optional outline level. |
| `pageBreakBefore` | `boolean?` | Optional page-break-before setting. |

## Revision Operations

| API | Purpose |
| --- | --- |
| `acceptRevisions(scope)` | Accept revisions matching a scope. |
| `rejectRevisions(scope)` | Reject revisions matching a scope. |

`RevisionScope`:

| Shape | Description |
| --- | --- |
| `{scope:'all'}` | Target all supported revisions. |
| `{scope:'paragraph'; paraId:string}` | Target supported revisions inside one paragraph. |
| `{scope:'revision'; revId:string}` | Target one revision id. |

## Bookmark And Comment Operations

| API | Purpose |
| --- | --- |
| `createBookmark({name, range})` | Create a bookmark around a paragraph or text range descriptor. |
| `deleteBookmark(name)` | Delete bookmark anchors. |
| `addComment({range, text, author?, date?})` | Add a comment to a paragraph or text range descriptor. |
| `deleteComment(commentId)` | Delete comment body and anchors. |

Range descriptors:

| Shape | Description |
| --- | --- |
| `{paraId}` | Whole paragraph range. |
| `{paraId, text, scope?}` | Text search range inside a paragraph. |

## Table Operations

| API | Purpose |
| --- | --- |
| `addTableComment({tableId, text, author?, date?})` | Add a comment to a whole table. |
| `deleteTable(tableId, opts?)` | Delete a table. |
| `replaceTableCellWithSdt(spec)` | Replace a cell's first paragraph with an inline SDT. |
| `replaceTableCellText(spec)` | Replace a cell's text. |
| `formatTableCell(spec)` | Set direct cell formatting. |
| `insertTableCellParagraph(spec)` | Insert a paragraph at the start or end of a cell. |
| `deleteTableRow(tableId, row, opts?)` | Delete a row by zero-based index. |
| `insertTableRow(spec)` | Insert a row using a template row. |
| `cloneTableAfter({anchorParaId, sourceTableId})` | Clone an existing table after a paragraph. |
| `insertTableFrom(spec)` | Insert a serialized table subtree and optional images after a paragraph. |
| `setTableColumnWidths({tableId, widths})` | Set table grid column widths. |
| `setTableRowHeights({tableId, heights})` | Set row heights. |

`TableCellDirectFormat`:

| Field | Type |
| --- | --- |
| `shading` | `{fill?:string; color?:string; val?:string}` |
| `verticalAlignment` | `'top' \| 'center' \| 'bottom' \| 'both'` |
| `width` | `number` |

## Drawing And Image Operations

| API | Purpose |
| --- | --- |
| `replaceDrawingWithPlaceholder(drawingId)` | Replace a picture drawing with the shared placeholder image. |
| `deleteDrawing(drawingId)` | Delete a drawing. |
| `replaceShapeInnerText(drawingId, placeholder)` | Replace text inside a text-bearing drawing. |
| `insertImageAfter(anchorParaId, spec)` | Insert a new inline image paragraph after a paragraph. |

`insertImageAfter` spec:

| Field | Type | Description |
| --- | --- | --- |
| `bytes` | `Uint8Array` | Image bytes. |
| `contentType` | `string` | Supported image MIME type. |
| `widthEmu` | `number` | Width in EMU. |
| `heightEmu` | `number` | Height in EMU. |

## Settings

| API | Purpose |
| --- | --- |
| `setUpdateFields(enabled)` | Set document setting for updating fields on open. |

## Apply

| API | Purpose | Returns |
| --- | --- | --- |
| `apply()` | Execute queued operations. | `Promise<ApplyResult>` |

`ApplyResult`:

| Field | Type | Description |
| --- | --- | --- |
| `insertedParagraphs` | `{opIndex; anchorParaId; position; newParaId}[]` | Paragraphs inserted around anchors. |
| `insertedBodyStartParaIds` | `string[]` | Paragraph ids inserted at body start. |
| `createdComments` | `{opIndex; commentId}[]` | Created comment ids. |
| `createdBookmarks` | `{opIndex; name; range:{fromParaId; toParaId}}[]` | Created bookmark records. |
| `acceptedRevisions` | `number` | Count of accepted revisions. |
| `rejectedRevisions` | `number` | Count of rejected revisions. |
| `replacedTableCells` | `number` | Count of replaced table cells. |

## CodeAgent Rules

1. Do not retrieve this document for ordinary paragraph/table/text tasks unless live-object APIs are insufficient.
2. Never invent plan APIs. Use only methods listed here.
3. `WordPlan` queues operations; reads after queuing but before `apply()` do not include queued changes.
4. For explicit all/bulk revision accept/reject requests, use `doc.plan().acceptRevisions({ scope: 'all' }).apply()` / `doc.plan().rejectRevisions({ scope: 'all' }).apply()` or live-object `doc.revisions.accept()` / `doc.revisions.reject()`; do not loop over `doc.revisions.list()` unless filtering and deduping by `revId`.
5. Prefer live-object APIs when generating code for small models.
