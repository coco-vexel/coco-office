# Document APIs

This domain covers creating blank documents, loading, exporting, runtime mode, change events, document-level domains, and document disposal.

Use this document when the task says: create blank document, create document object, load docx, load Flat OPC, export docx, export Flat OPC, listen to changes, worker mode, inline mode, dispose, or initialize core.

## Skill Summary

`WordDocument` is the root object. Core can create an in-memory blank Word document with `WordDocument.createBlank`. Host environments are responsible for obtaining document bytes or Flat OPC XML when loading an existing document. Core does not know how a host obtains the current document or opens editor windows; host code must provide bytes/XML to `WordDocument.fromZip` or `WordDocument.fromFlatOpc` for existing documents.

Generated document-operation code should normally assume a ready `doc: WordDocument` is passed in by the host/runtime. It should not initialize host-specific APIs.

## Load APIs

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `WordDocument.createBlank(opts)` | Create a new in-memory blank Word document. | `LoadOptions` | `Promise<WordDocument>` |
| `WordDocument.fromZip(buf, opts)` | Load a `.docx`/zip package. | `Uint8Array \| ArrayBuffer`, `LoadOptions` | `Promise<WordDocument>` |
| `WordDocument.fromFlatOpc(xml, opts)` | Load Flat OPC XML. | `xml: string`, `LoadOptions` | `Promise<WordDocument>` |

`createBlank` creates a minimal valid Word document with one empty paragraph as an editable anchor and uses the same runtime, snapshot, and diagnostics options as the load APIs. It does not open a document in WPS, Microsoft Word, OnlyOffice, or any other host; opening or replacing a host document remains a host integration responsibility.

`LoadOptions`:

| Field | Type | Required | Description |
| --- | --- | --- | --- |
| `docId` | `string` | yes | Stable document id for errors, snapshots, and diagnostics. |
| `mode` | `'auto' \| 'worker' \| 'inline'` | no | Runtime mode. Default uses global default, initially `'auto'`. |
| `onModeResolved` | `(decision, docId) => void` | no | Receives runtime mode decision. |
| `onProgress` | `(phase, pct?) => void` | no | Load progress callback. Phases include `parse`. |
| `snapshotStorage` | `SnapshotStorageAdapter` | no | Per-document snapshot storage. |
| `maxSnapshots` | `number` | no | FIFO limit for snapshots. Must be positive integer. |

`RuntimeMode`:

| Value | Description |
| --- | --- |
| `auto` | Choose worker when available; otherwise inline. |
| `worker` | Require worker transport. Throws if unavailable. |
| `inline` | Run kernel in the current JS thread. |

## Document Properties

| Property | Type | Description |
| --- | --- | --- |
| `doc.docId` | `string` | Document id from `LoadOptions`. |
| `doc.mode` | `'worker' \| 'inline'` | Actual runtime mode. |
| `doc.modeSource` | `ModeSource` | Why the mode was chosen. |
| `doc.paragraphs` | `ParagraphApi` | Paragraph domain. |
| `doc.tables` | `TablesApi` | Table domain. |
| `doc.bookmarks` | `BookmarkApi` | Bookmark domain. |
| `doc.comments` | `CommentApi` | Comment domain. |
| `doc.revisions` | `RevisionApi` | Revision domain. |
| `doc.styles` | `StylesApi` | Style domain. |
| `doc.drawings` | `DrawingsApi` | Drawing domain. |
| `doc.sections` | `SectionsApi` | Section domain. |
| `doc.tocFields` | `TocFieldsApi` | TOC field domain. |
| `doc.snapshots` | `SnapshotApi` | Snapshot domain. |

## Document-Level Writes

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.appendParagraph(text, opts?)` | Append a plain-text paragraph to the end of the main document body. | `text: string`, `ParagraphInsertOptions?` | `Promise<ApplyResult>` |

`appendParagraph` is a convenience live-object wrapper over paragraph insertion. It targets the main document body. For precise placement, list or get a paragraph and call `paragraph.insertBefore(...)` or `paragraph.insertAfter(...)`.

## Export And Dispose

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.toZip()` | Export current document as docx bytes. | none | `Promise<Uint8Array>` |
| `doc.toFlatOpc()` | Export current document as Flat OPC XML. | none | `Promise<string>` |
| `doc.dispose()` | Dispose the underlying transport. | none | `Promise<void>` |

After `dispose`, public APIs throw when used.

## Change Events

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.on('change', handler, opts?)` | Subscribe to successful document changes. | `ChangeListener`, `ChangeSubscribeOptions?` | `() => void` unsubscribe |

`ChangeEvent`:

| Field | Type | Description |
| --- | --- | --- |
| `ops` | `Op[]` | Applied operation list. |
| `result` | `ApplyResult` | Apply result. |
| `source` | `'plan' \| 'script'` | Change source. |
| `flatOpc` | `string?` | Present when subscribed with `fullDocument:'flatOpc'`. |
| `zip` | `Uint8Array?` | Present when subscribed with `fullDocument:'zip'`. |

`ChangeSubscribeOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `fullDocument` | `'none' \| 'flatOpc' \| 'zip'` | `'none'` | Include full document output in each event. This is expensive. |

## Static Runtime Configuration

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `WordDocument.setDefaultMode(mode)` | Set global default runtime mode. | `RuntimeMode` | `void` |
| `WordDocument.detectMode()` | Detect current auto mode. | none | `'worker' \| 'inline'` |
| `WordDocument.detectModeDecision()` | Detect mode and decision source. | none | `ModeDecision` |
| `WordDocument.setDefaultAuthor(author)` | Set default author for operations that need one. | `string` | `void` |
| `WordDocument.setWorkerFactory(factory)` | Configure custom worker creation. | `WorkerFactory \| undefined` | `void` |
| `WordDocument.setWorkerBundleUrl(url)` | Configure Node worker bundle URL. | `string \| URL \| undefined` | `void` |
| `WordDocument.setDefaultSnapshotStorage(adapter)` | Set global snapshot storage. | `SnapshotStorageAdapter \| undefined` | `void` |

## Advanced APIs

| API | Purpose | Guidance |
| --- | --- | --- |
| `doc.executeCode(code, args?, timeoutMs?, signal?)` | Execute JS in the document runtime. | Advanced runtime bridge. Do not use for normal generated document operations when a `WordDocument` is already available. |
| `doc.plan(opts?)` | Low-level operation builder. | Advanced surface. Prefer live-object APIs for small-model code generation and workflow composition. |

## Examples

Load from docx bytes:

```ts
const doc = await WordDocument.fromZip(bytes, {
  docId: 'example.docx',
  mode: 'auto',
});
```

Create a blank document and write the first paragraph:

```ts
const doc = await WordDocument.createBlank({
  docId: 'new-doc',
  mode: 'auto',
});

const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });
await paragraphs[0]?.replaceWith('First paragraph');
await doc.appendParagraph('Second paragraph');

const output = await doc.toZip();
```

Export after edits:

```ts
const paragraph = await doc.paragraphs.get('31086A56');
await paragraph?.replaceWith('新的段落文本');
const output = await doc.toZip();
```

Subscribe to changes:

```ts
const unsubscribe = doc.on('change', (event) => {
  console.log(event.source, event.result);
});
```

## CodeAgent Rules

1. For document-editing tasks on an existing host document, assume the host/runtime provides a ready `doc`.
2. Use `WordDocument.createBlank` only when the task explicitly needs a new in-memory document, not when it asks to open a new host editor window.
3. Do not call host APIs, browser file APIs, or environment-specific document loaders from core operation code.
4. Prefer live-object domain APIs over `doc.plan()` for generated code.
5. Use `toZip` or `toFlatOpc` only when the task requires exporting the updated document.
6. Use `dispose` when code owns the document lifecycle.
