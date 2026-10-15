# Core Word API Reference

This file is the public API index. Detailed parameters, return values, examples, and code-agent rules live in the domain documents.

Import:

```ts
import { WordDocument } from '@coco-vexel/coco-office/word';
import type { Paragraph, Table, TextRange } from '@coco-vexel/coco-office/word';
```

## Primary Object Model

`WordDocument` is the root object. Domain APIs hang from `doc`:

| Domain Property  | Main Types                                                                                  | Full Documentation                                       |
| ---------------- | ------------------------------------------------------------------------------------------- | -------------------------------------------------------- |
| `doc.paragraphs` | `ParagraphApi`, `ParagraphList`, `Paragraph`, `ParagraphInfo`, `ParagraphFormat`            | [domains/paragraphs.md](./domains/paragraphs.md)         |
| `doc.tables`     | `TablesApi`, `Table`, `TableCell`, `TableRow`, `TableColumn`, `TableInfo`, `TableCellInfo`  | [domains/tables.md](./domains/tables.md)                 |
| `doc.drawings`   | `DrawingsApi`, `Drawing`, `PictureDrawing`, `TextBoxDrawing`, `ShapeDrawing`, `DrawingInfo` | [domains/drawings.md](./domains/drawings.md)             |
| `doc.bookmarks`  | `BookmarkApi`, `Bookmark`, `BookmarkInfo`                                                   | [domains/bookmarks.md](./domains/bookmarks.md)           |
| `doc.comments`   | `CommentApi`, `Comment`, `CommentInfo`                                                      | [domains/comments.md](./domains/comments.md)             |
| `doc.revisions`  | `RevisionApi`, `Revision`, `RevisionInfo`                                                   | [domains/revisions.md](./domains/revisions.md)           |
| `doc.styles`     | `StylesApi`, `StyleInfo`                                                                    | [domains/styles.md](./domains/styles.md)                 |
| `doc.sections`   | `SectionsApi`, `SectionInfo`                                                                | [domains/sections.md](./domains/sections.md)             |
| `doc.tocFields`  | `TocFieldsApi`, `TocFieldRange`                                                             | [domains/toc-fields.md](./domains/toc-fields.md)         |
| `doc.snapshots`  | `SnapshotApi`, `RollbackOptions`                                                            | [domains/snapshots.md](./domains/snapshots.md)           |
| `doc.context`    | `ContextApi`, `ContextSearchResult`, `ContextSearchMatch`                                   | [domains/context-search.md](./domains/context-search.md) |

Text ranges are created from paragraph or cell objects:

| Creator                             | Returns         | Documentation                            |
| ----------------------------------- | --------------- | ---------------------------------------- |
| `paragraph.findText(search, opts?)` | `TextRange`     | [domains/ranges.md](./domains/ranges.md) |
| `cell.findText(search, opts?)`      | `CellTextRange` | [domains/ranges.md](./domains/ranges.md) |

## Document Lifecycle

| API                                   | Documentation                                |
| ------------------------------------- | -------------------------------------------- |
| `WordDocument.createBlank(opts)`      | [domains/document.md](./domains/document.md) |
| `WordDocument.fromZip(buf, opts)`     | [domains/document.md](./domains/document.md) |
| `WordDocument.fromFlatOpc(xml, opts)` | [domains/document.md](./domains/document.md) |
| `doc.appendParagraph(text, opts?)`    | [domains/document.md](./domains/document.md) |
| `doc.toZip()`                         | [domains/document.md](./domains/document.md) |
| `doc.toFlatOpc()`                     | [domains/document.md](./domains/document.md) |
| `doc.dispose()`                       | [domains/document.md](./domains/document.md) |
| `doc.on('change', handler, opts?)`    | [domains/document.md](./domains/document.md) |

## Advanced And Low-Level Exports

The package also exports lower-level utilities for core internals, tests, and advanced integrations:

| Area                      | Examples                                                                                                                                        |
| ------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| OOXML namespaces          | `W_NS`, `R_NS`, `W14_NS`, `A_NS`, `WP_NS`                                                                                                       |
| Paragraph/run utilities   | `ensureParaIds`, `getParagraphRuns`, `getRunText`, `getParagraphText`, `splitRunAt`                                                             |
| Read-query helpers        | `listParagraphs`, `listTables`, `listDrawings`, `listBookmarks`, `listComments`, `listRevisions`, `listSections`, `listTocFields`, `listStyles` |
| Indices and stories       | `buildIndices`, `listStories`, `getStoryParagraphs`, `resolveScope`                                                                             |
| Advanced mutation/runtime | `WordPlanImpl`, `applyXQueryUpdate`, `runTransaction`, `Kernel`, `InlineTransport`, `WorkerTransport`                                           |
| Errors                    | `WordError`, `WordErrorCode`                                                                                                                    |

For ordinary generated document-operation code, prefer the live-object APIs documented in `domains/`.

Low-level `WordPlan` is documented in [advanced/plan.md](./advanced/plan.md).

## CodeAgent Retrieval Map

| User Intent                                                           | Retrieve                                                                             |
| --------------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| "create blank document", "new document", "generate docx from scratch" | `document.md`, then the content domains needed for writing                           |
| "modify paragraph", "insert after heading", "delete paragraph"        | `paragraphs.md`                                                                      |
| "append paragraph", "write new paragraph at end"                      | `document.md`, `paragraphs.md`                                                       |
| "find keyword/context/clause/topic", "locate bookmarked content"      | `context-search.md`, then the source domain needed for follow-up writes              |
| "mark words red", "bold keyword", "replace matched text"              | `paragraphs.md`, `ranges.md`                                                         |
| "insert image/picture after paragraph"                                | `paragraphs.md`, `drawings.md`                                                       |
| "insert simple table after paragraph"                                 | `paragraphs.md`, `tables.md`                                                         |
| "operate table/cell/row/column"                                       | `tables.md`, optionally `ranges.md`                                                  |
| "comment"                                                             | `comments.md`, plus the source domain (`paragraphs.md`, `ranges.md`, or `tables.md`) |
| "bookmark"                                                            | `bookmarks.md`, plus `paragraphs.md` or `ranges.md` for creation                     |
| "tracked changes"                                                     | `revisions.md`                                                                       |
| "image/textbox/shape"                                                 | `drawings.md`                                                                        |
| "heading style/style id"                                              | `styles.md`, usually `paragraphs.md`                                                 |
| "section/header/footer boundary"                                      | `sections.md`                                                                        |
| "TOC/table of contents"                                               | `toc-fields.md`, usually `paragraphs.md`                                             |
| "rollback/checkpoint"                                                 | `snapshots.md`                                                                       |
