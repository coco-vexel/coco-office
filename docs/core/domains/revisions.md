# Revision APIs

This domain covers tracked insert/delete revision discovery and revision accept/reject.

Use this document when the task says: revision, tracked changes, accept change, reject change, inserted text, deleted text, or review changes.

## Skill Summary

Revision APIs expose Tier 1 tracked changes: `<w:ins>` and `<w:del>`. Attribute-level revisions such as `w:rPrChange`, `w:pPrChange`, move ranges, table property changes, or section property changes are unsupported and may throw `UNSUPPORTED_OOXML_FEATURE`.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.revisions.list(opts?)` | List revisions. | `ListRevisionsOptions` | `Promise<Revision[]>` |
| `doc.revisions.get(revId, opts?)` | Get one revision. | `revId: string`, `ListRevisionsOptions?` | `Promise<Revision \| null>` |
| `doc.revisions.accept(opts?)` | Accept all supported revisions. | `RevisionBulkOptions?` | `Promise<ApplyResult>` |
| `doc.revisions.reject(opts?)` | Reject all supported revisions. | `RevisionBulkOptions?` | `Promise<ApplyResult>` |

`ListRevisionsOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `scope` | `'body' \| 'all' \| StoryId[]` | `'body'` | Read body revisions, all story revisions, or explicit stories. |

`RevisionBulkOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `scope` | `'all'` | `'all'` | Target all supported revisions. |

## RevisionInfo

`Revision` extends `RevisionInfo`.

| Field | Type | Description |
| --- | --- | --- |
| `revId` / `id` | `string` | Revision id. |
| `kind` | `'insert' \| 'delete'` | Inserted or deleted content. |
| `author` | `string` | Revision author. |
| `date` | `string` | Revision date. |
| `story` | `StoryId` | Story that contains the revision. |
| `paraId` | `string?` | Paragraph id when the revision is inside a paragraph. |
| `text` | `string?` | Inserted/deleted text when available. |

## Revision Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `revision.getInfo(opts?)` | Read current metadata. | `ListRevisionsOptions?` | `Promise<RevisionInfo \| null>` |
| `revision.getText()` | Read revision text. | none | `Promise<string \| null>` |
| `revision.getParagraph()` | Return containing paragraph when available. | none | `Promise<Paragraph \| null>` |
| `revision.accept(opts?)` | Accept this revision. | `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |
| `revision.reject(opts?)` | Reject this revision. | `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |

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
