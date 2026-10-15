# Comment APIs

This domain covers comment discovery, comment live objects, comment body text, anchored text, and comment deletion.

Use this document when the task says: comment, annotation, review note, add comment, delete comment, comment body, or commented text.

## Skill Summary

Comments have two parts: a comment body in `comments.xml` and an anchor in the document. `doc.comments.list()` and `doc.comments.get()` return live `Comment` handles. Adding comments is done from `Paragraph`, `TextRange`, or `Table`.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.comments.list()` | List comments. | none | `Promise<Comment[]>` |
| `doc.comments.get(id)` | Get one comment. | `id: number` | `Promise<Comment \| null>` |

## CommentInfo

`Comment` extends `CommentInfo`.

| Field | Type | Description |
| --- | --- | --- |
| `id` | `number` | Comment id. |
| `author` | `string` | Comment author, empty string if missing. |
| `date` | `string` | ISO date string, empty string if missing. |
| `text` | `string` | Comment body text. |
| `rangeKind` | `'paragraph' \| 'text' \| 'complex'` | Shape of the anchor range. |
| `paraId` | `string?` | Anchor paragraph id when available. |
| `anchorText` | `string?` | Text covered by a text anchor. |

## Comment Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `comment.getInfo()` | Read current metadata. | none | `Promise<CommentInfo \| null>` |
| `comment.getText()` | Read current comment body text. | none | `Promise<string \| null>` |
| `comment.getAnchorText()` | Read current document text covered by the anchor. | none | `Promise<string \| null>` |
| `comment.getParagraph()` | Return the paragraph containing the anchor when available. | none | `Promise<Paragraph \| null>` |
| `comment.delete(opts?)` | Delete comment body and anchors. | `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |

## Add Comments

| API | Purpose |
| --- | --- |
| `paragraph.addComment(text, opts?)` | Add a comment to a whole paragraph. |
| `paragraph.findText(search, rangeOpts?).addComment(text, opts?)` | Add a comment to matched text. |
| `cell.findText(search, rangeOpts?).addComment(text, opts?)` | Add a comment to matched text in a cell. |
| `table.addComment(text, opts?)` | Add a comment to the whole table. |

`opts` is `LiveObjectWriteOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `trackChanges` | `boolean` | `false` | Request tracked-change markup when supported. |
| `author` | `string` | document default | Comment author. |
| `date` | `Date` | operation time | Comment date metadata. |

Author example:

```ts
await paragraph.addComment('修改建议。', { author: 'halvee' });
await paragraph.findText('错别字').addComment('建议修改。', { author: 'halvee' });
await table.addComment('表格需要复核。', { author: 'halvee' });
await cell.findText('风险').addComment('这里需要说明。', { author: 'halvee' });
```

## Examples

List comments:

```ts
const comments = await doc.comments.list();
```

Delete resolved comments by body text:

```ts
for (const comment of await doc.comments.list()) {
  const text = await comment.getText();
  if (text?.includes('已处理')) {
    await comment.delete();
  }
}
```

Add a comment to matched text:

```ts
const paragraph = await doc.paragraphs.get('31086A56');
await paragraph?.findText('请确认').addComment('需要业务确认。', {
  author: 'Coco',
  date: new Date(),
});
```

## CodeAgent Rules

1. Use `comment.getText()` for comment body; use `comment.getAnchorText()` for document text under the comment.
2. Adding comments belongs to paragraph/range/table objects; do not invent `doc.comments.create`.
3. Handle `null` from `getInfo`, `getText`, `getAnchorText`, and `getParagraph`.
