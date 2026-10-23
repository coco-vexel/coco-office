# Revision APIs

This domain covers WordprocessingML tracked-revision discovery and accept/reject.

Use this document when the task says: revision, tracked changes, accept change, reject change, inserted text, deleted text, or review changes.

## Skill Summary

Revision APIs expose the SDK's WordprocessingML revision model: inserted/deleted content and paragraph marks, moved content, run/paragraph/section/table property history, table row/cell revisions, numbering revisions, custom XML revision ranges, and Office 2010 conflict revisions. The read and accept/reject paths use the same type registry.

## Entry Points

| API                               | Purpose                         | Params                                   | Returns                     |
| --------------------------------- | ------------------------------- | ---------------------------------------- | --------------------------- |
| `doc.revisions.list(opts?)`       | List revisions.                 | `ListRevisionsOptions`                   | `Promise<Revision[]>`       |
| `doc.revisions.get(revId, opts?)` | Get one revision.               | `revId: string`, `ListRevisionsOptions?` | `Promise<Revision \| null>` |
| `doc.revisions.accept(opts?)`     | Accept all supported revisions. | `RevisionBulkOptions?`                   | `Promise<ApplyResult>`      |
| `doc.revisions.reject(opts?)`     | Reject all supported revisions. | `RevisionBulkOptions?`                   | `Promise<ApplyResult>`      |

`ListRevisionsOptions`:

| Field   | Type                           | Default  | Description                                                    |
| ------- | ------------------------------ | -------- | -------------------------------------------------------------- |
| `scope` | `'body' \| 'all' \| StoryId[]` | `'body'` | Read body revisions, all story revisions, or explicit stories. |

`RevisionBulkOptions`:

| Field   | Type    | Default | Description                     |
| ------- | ------- | ------- | ------------------------------- |
| `scope` | `'all'` | `'all'` | Target all supported revisions. |

## RevisionInfo

`Revision` extends `RevisionInfo`.

| Field          | Type                         | Description                                                                               |
| -------------- | ---------------------------- | ----------------------------------------------------------------------------------------- |
| `revId` / `id` | `string`                     | Revision id.                                                                              |
| `kind`         | `RevisionKind`               | Semantic SDK category, such as `insert`, `paragraphProperty`, `moveFrom`, or `cellMerge`. |
| `ooxmlType`    | `SupportedRevisionType`      | Exact OOXML local name, such as `pPrChange`, `cellIns`, or `conflictIns`.                 |
| `author`       | `string`                     | Revision author.                                                                          |
| `date`         | `string`                     | Revision date.                                                                            |
| `story`        | `StoryId`                    | Story that contains the revision.                                                         |
| `paraId`       | `string?`                    | Paragraph id when the revision is inside a paragraph.                                     |
| `text`         | `string?`                    | Inserted/deleted text when available.                                                     |

Supported OOXML types are `ins`, `del`, `moveFrom`, `moveTo`, `rPrChange`, `pPrChange`, `sectPrChange`, `tblPrChange`, `trPrChange`, `tcPrChange`, `tblPrExChange`, `tblGridChange`, `numberingChange`, `cellIns`, `cellDel`, `cellMerge`, the four custom XML insert/delete/move range starts, and `w14:conflictIns`, `w14:conflictDel`, `w14:customXmlConflictInsRangeStart`, and `w14:customXmlConflictDelRangeStart`. Matching range-end elements are structural markers and are not returned as separate revisions.

Accept/reject restores historical property snapshots on reject, resolves paragraph and row marks structurally, converts deleted text back to normal text when required, and removes paired move/custom XML range markers.

## Revision Object

| API                       | Purpose                                     | Params                    | Returns                         |
| ------------------------- | ------------------------------------------- | ------------------------- | ------------------------------- |
| `revision.getInfo(opts?)` | Read current metadata.                      | `ListRevisionsOptions?`   | `Promise<RevisionInfo \| null>` |
| `revision.getText()`      | Read revision text.                         | none                      | `Promise<string \| null>`       |
| `revision.getParagraph()` | Return containing paragraph when available. | none                      | `Promise<Paragraph \| null>`    |
| `revision.accept(opts?)`  | Accept this revision.                       | `LiveObjectWriteOptions?` | `Promise<ApplyResult>`          |
| `revision.reject(opts?)`  | Reject this revision.                       | `LiveObjectWriteOptions?` | `Promise<ApplyResult>`          |

## Examples

Accept all supported revisions:

```ts
await doc.revisions.accept();
```

Reject all supported revisions:

```ts
await doc.revisions.reject();
```

Reject revisions containing a phrase:

```ts
const handled = new Set<string>();
for (const revision of await doc.revisions.list({ scope: 'all' })) {
  if (handled.has(revision.revId)) continue;
  const text = await revision.getText();
  if (text?.includes('临时')) {
    await revision.reject();
    handled.add(revision.revId);
  }
}
```

## CodeAgent Rules

1. Use `doc.revisions.accept()` / `doc.revisions.reject()` for explicit all/bulk revision accept/reject requests.
2. Use revision object methods for single-revision or filtered operations.
3. If looping over `doc.revisions.list()`, dedupe by `revId` before accepting or rejecting. Replacement-like edits may expose multiple handles that share one `revId`.
4. Do not promise support for attribute-level revisions.
5. `getText()` may return `null` for non-textual or unsupported revision shapes.
