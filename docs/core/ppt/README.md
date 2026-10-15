# Core PPT API Docs

This directory documents the public `@coco-vexel/coco-office/ppt` presentation-operation APIs.

The docs follow the same domain-based structure as the Word and Excel SDK docs:
an entry README, a compact API index, domain-level API specifications, model
documentation, runtime notes, mutation/writeback notes, and host integration
notes. SDK callers should retrieve only the domain documents required by the
current task instead of loading the full PPT surface.

## Recommended Reading Order

1. [api-reference.md](./api-reference.md): public API index, domain retrieval map, and SDK operation rules.
2. [domains/presentation.md](./domains/presentation.md): `Presentation` loading, exporting, properties, image export, and lifecycle.
3. [domains/slides.md](./domains/slides.md): slide listing, creation, copy, move, delete, properties, and text search.
4. [domains/shapes.md](./domains/shapes.md): shape discovery, geometry, naming, z-order, creation, deletion, and style.
5. [domains/text.md](./domains/text.md): shape text reads/writes, replacement, and text style.
6. [domains/images.md](./domains/images.md): picture metadata, bytes, replacement, add image, and global image export.
7. [domains/tables.md](./domains/tables.md): table values, cells, cell style, row/column structure edits.
8. [domains/search.md](./domains/search.md): cross-presentation and slide-level text search/replace.
9. [style-capabilities.md](./style-capabilities.md): core-only style read/write matrix and P0/P1/P2 plan.
10. [model.md](./model.md): normalized model, properties snapshots, mutation payloads, and test fixtures.
11. [runtime.md](./runtime.md): inline/worker boundary and `executeCode` status.
12. [mutation-writeback.md](./mutation-writeback.md): SDK-recorded mutations for host adapters.
13. [writeback-planner.md](./writeback-planner.md): model diff and host-neutral writeback strategy planning.
14. [host-integration.md](./host-integration.md): Microsoft PowerPoint, WPS, ONLYOFFICE, and desktop bridge guidance.
15. [sdk-plan.md](./sdk-plan.md): milestone plan and completion status.
16. [sdk-review.md](./sdk-review.md): SDK boundary review, current evidence, risks, and next hardening work.

Compatibility entry: [../ppt-sdk-api.md](../ppt-sdk-api.md) points existing
SDK/API-list references to this domain-based documentation tree.

## Functional Domains

| Domain             | Document                                             | Main Use                                                                               |
| ------------------ | ---------------------------------------------------- | -------------------------------------------------------------------------------------- |
| Presentation       | [domains/presentation.md](./domains/presentation.md) | Load `.pptx`, create blank deck, export `.pptx`, read properties, export media/fonts.  |
| Slides             | [domains/slides.md](./domains/slides.md)             | List/get/add/copy/move/delete slides and slide-level text search.                      |
| Shapes             | [domains/shapes.md](./domains/shapes.md)             | List/get/create/delete shapes, groups, geometry, naming, z-order, fill and line style. |
| Text               | [domains/text.md](./domains/text.md)                 | Text read/write/replace and font style.                                                |
| Images             | [domains/images.md](./domains/images.md)             | Add images, inspect image metadata, extract/replace bytes.                             |
| Tables             | [domains/tables.md](./domains/tables.md)             | Table values, cell text/style, rows, and columns.                                      |
| Search             | [domains/search.md](./domains/search.md)             | Locate and replace text across text shapes and table cells.                            |
| Style Matrix       | [style-capabilities.md](./style-capabilities.md)     | Check which core style fields are readable, writable, planned, or out of scope.        |
| Model              | [model.md](./model.md)                               | Use snapshots, properties, media models, and mutation records.                         |
| Runtime            | [runtime.md](./runtime.md)                           | Understand async live-object semantics and worker status.                              |
| Mutation Writeback | [mutation-writeback.md](./mutation-writeback.md)     | Let host adapters inspect safe SDK operations before fallback package writeback.       |
| Writeback Planner  | [writeback-planner.md](./writeback-planner.md)       | Convert model diffs and host capabilities into portable writeback intents.             |
| Host Integration   | [host-integration.md](./host-integration.md)         | Map core package/model operations to PowerPoint/WPS/ONLYOFFICE/Desktop bridges.        |
| SDK Review         | [sdk-review.md](./sdk-review.md)                     | Review current SDK readiness, gaps, and core boundaries.                               |

## Core Boundary

`@coco-vexel/coco-office/ppt` is an OOXML-native presentation operation kernel. Callers
provide `.pptx` bytes or a normalized `PptPresentationModel`. Core parses,
queries, mutates, records operations, and serializes the presentation package.

Core does not directly handle:

- Office.js, WPS JS, ONLYOFFICE plugin APIs, COM automation, or browser task pane lifecycle.
- File dialogs, local path permissions, downloads, upload policy, or user confirmation UI.
- Host selection, cursor state, thumbnails/rendering, animations, slideshow playback, or live application focus.
- Business routing, model/provider configuration, login state, or agent orchestration policy.

## SDK Caller Guidance

1. PPT operation code should normally assume a ready `presentation: Presentation`.
2. Prefer async live-object APIs such as `presentation.slides.getItem(0).shapes.addTextBox(...)`.
3. Do not import or use `PresentationCore` or other `*Core` classes unless the task explicitly asks for kernel tests or adapter internals.
4. Do not use Office.js, WPSJS, ONLYOFFICE APIs, or raw OOXML in core operation code.
5. Treat `toModel()` as a snapshot. Mutating the returned object does not mutate the presentation.
6. Use `toZip()` only when the runtime needs package bytes for writeback or file output.
7. Use `getMutations()` only in runtime/writeback code. Ordinary SDK operation code should normally not inspect mutation logs.
8. Retrieve only the matching domain document for the task.

## Maintenance

- `packages/core/src/ppt/index.ts` is the export source of truth.
- `packages/core/src/ppt/runtime.ts` is the public async live-object API source of truth.
- `packages/core/src/ppt/presentation-core.ts` is the synchronous core-kernel source of truth for tests and advanced adapters.
- `packages/core/src/ppt/model.ts` is the normalized model source of truth.
- `packages/core/src/ppt/mutations.ts` is the mutation payload source of truth.
- `packages/core/src/ppt/__tests__/presentation.test.ts` is the current capability and real-PPT regression test matrix.
- Update the matching domain document whenever a public API changes.
