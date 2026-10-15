# Bookmark APIs

This domain covers bookmark discovery, bookmark live objects, bookmark text, and bookmark deletion.

Use this document when the task says: bookmark, named range, anchor, locate by bookmark, delete bookmark, or create bookmark.

## Skill Summary

Bookmarks are named anchors in the document. `doc.bookmarks.list()` and `doc.bookmarks.get()` return live `Bookmark` handles. Creating bookmarks is done from `Paragraph` or `TextRange` objects.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.bookmarks.list(opts?)` | List bookmarks. | `ListBookmarksOptions` | `Promise<Bookmark[]>` |
| `doc.bookmarks.get(name)` | Get one bookmark by name. | `name: string` | `Promise<Bookmark \| null>` |

`ListBookmarksOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `includeHidden` | `boolean` | `false` | Include bookmarks whose name starts with `_`, such as Word internal bookmarks. |
| `prefix` | `string` | none | Return only bookmarks whose name starts with this prefix. Case-sensitive. |

## BookmarkInfo

`Bookmark` extends `BookmarkInfo`.

| Field | Type | Description |
| --- | --- | --- |
| `name` / `id` | `string` | Bookmark name. |
| `story` | `StoryId` | Story that contains the bookmark start. |
| `rangeKind` | `'paragraph' \| 'text' \| 'complex'` | Shape of the bookmark range. |
| `paraId` | `string?` | Paragraph id when the range can be represented by one paragraph. |
| `text` | `string?` | Covered text when `rangeKind:'text'`. |
| `hidden` | `boolean?` | Present when the bookmark name starts with `_`. |

`rangeKind`:

| Value | Meaning |
| --- | --- |
| `paragraph` | Bookmark covers a whole paragraph. |
| `text` | Bookmark covers a text span inside one paragraph. |
| `complex` | Cross-paragraph, missing endpoint, or unsupported shape. |

## Bookmark Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `bookmark.getInfo()` | Read current metadata. | none | `Promise<BookmarkInfo \| null>` |
| `bookmark.getText()` | Read current covered text when representable. | none | `Promise<string \| null>` |
| `bookmark.getParagraph()` | Return containing paragraph when available. | none | `Promise<Paragraph \| null>` |
| `bookmark.delete(opts?)` | Delete bookmark anchor pair. | `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |

## Create Bookmarks

| API | Purpose |
| --- | --- |
| `paragraph.createBookmark(name, opts?)` | Bookmark the whole paragraph. |
| `paragraph.findText(search).createBookmark(name, opts?)` | Bookmark matched text. |
| `cell.findText(search).createBookmark(name, opts?)` | Bookmark matched text in a cell. |

## Examples

List business bookmarks:

```ts
const bookmarks = await doc.bookmarks.list({ prefix: 'coco_' });
```

Read bookmarked text:

```ts
const bookmark = await doc.bookmarks.get('target_clause');
const text = await bookmark?.getText();
```

Create a bookmark around a paragraph:

```ts
const paragraph = await doc.paragraphs.get('31086A56');
await paragraph?.createBookmark('review_anchor');
```

Create a bookmark around matched text:

```ts
const paragraph = await doc.paragraphs.get('31086A56');
await paragraph?.findText('合同金额').createBookmark('amount_text');
```

## CodeAgent Rules

1. Use `doc.bookmarks.get(name)` when the bookmark name is known.
2. Hidden bookmarks are excluded by default; set `includeHidden:true` only when explicitly needed.
3. If `rangeKind:'complex'`, do not assume `getParagraph()` or `getText()` can return useful data.
4. Create bookmarks from `Paragraph` or `TextRange`; do not invent `doc.bookmarks.create`.
