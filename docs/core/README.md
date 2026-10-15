# Core Word API Docs

This directory documents the public `@coco-vexel/coco-office/word` document-operation APIs.

The docs are organized by functional domain so they can be read by humans and retrieved by a code agent as small skill-like chunks.

## Recommended Reading Order

1. [domains/document.md](./domains/document.md): `WordDocument` loading, exporting, runtime mode, events, and lifecycle.
2. [domains/paragraphs.md](./domains/paragraphs.md): paragraph objects and paragraph-level operations.
3. [domains/ranges.md](./domains/ranges.md): matched text operations such as replace, delete, bold, color, comments, and bookmarks.
4. [domains/tables.md](./domains/tables.md): tables, cells, rows, columns, and cell content.
5. [domains/context-search.md](./domains/context-search.md): planned local context search for keyword/topic/bookmark-driven document operations.
6. [api-reference.md](./api-reference.md): clean domain index for the current public surface.

## Functional Domains

| Domain         | Document                                                 | Main Use                                                                        |
| -------------- | -------------------------------------------------------- | ------------------------------------------------------------------------------- |
| Document       | [domains/document.md](./domains/document.md)             | Load/export documents, runtime mode, events, lifecycle.                         |
| Paragraphs     | [domains/paragraphs.md](./domains/paragraphs.md)         | List paragraphs, read text/info, replace, insert, delete, paragraph formatting. |
| Ranges         | [domains/ranges.md](./domains/ranges.md)                 | Operate on matched text inside paragraphs or cells.                             |
| Tables         | [domains/tables.md](./domains/tables.md)                 | Table/cell/row/column objects and table content operations.                     |
| Context Search | [domains/context-search.md](./domains/context-search.md) | Locate actionable local context by keyword, heading, table text, and bookmarks. |
| Bookmarks      | [domains/bookmarks.md](./domains/bookmarks.md)           | Read/delete bookmarks and create bookmarks from paragraphs or ranges.           |
| Comments       | [domains/comments.md](./domains/comments.md)             | Read/delete comments and add comments from paragraphs/ranges/tables.            |
| Revisions      | [domains/revisions.md](./domains/revisions.md)           | Read tracked insert/delete revisions and accept/reject one revision.            |
| Drawings       | [domains/drawings.md](./domains/drawings.md)             | Pictures, text boxes, shapes, drawing metadata and supported writes.            |
| Styles         | [domains/styles.md](./domains/styles.md)                 | Read style definitions and style metadata.                                      |
| Sections       | [domains/sections.md](./domains/sections.md)             | Read section boundaries and header/footer references.                           |
| TOC Fields     | [domains/toc-fields.md](./domains/toc-fields.md)         | Read table-of-contents field ranges.                                            |
| Snapshots      | [domains/snapshots.md](./domains/snapshots.md)           | Take, list, rollback, and clear document snapshots.                             |

## Core Boundary

`@coco-vexel/coco-office/word` is an OOXML-native document operation kernel. Callers pass document bytes or Flat OPC XML into core; core parses, queries, modifies, and serializes the document.

Core does not directly handle:

- Local file picking or file path policy.
- Network access.
- UI confirmation.
- Product strategy or business classification.

## CodeAgent Guidance

These docs intentionally use a stable structure:

1. Skill Summary
2. Entry Points
3. Data shapes
4. Object methods
5. Examples
6. CodeAgent Rules

When building a code agent, retrieve only the domain documents required by the user task. For example, a task that says "mark typo words red and bold" should retrieve `paragraphs.md` and `ranges.md`, not the full API set.

Prefer live-object APIs for generated code. Treat `doc.plan()` and low-level helpers as advanced APIs unless a task explicitly asks for them.

Advanced APIs are documented separately:

- [advanced/plan.md](./advanced/plan.md): low-level `WordPlan` operation builder.

## Maintenance

- `packages/core/src/word/index.ts` is the export source of truth.
- `packages/core/src/word/sub-apis.ts` and `packages/core/src/word/word-document.ts` are the live-object API source of truth.
- Update the matching domain document whenever a public API changes.
