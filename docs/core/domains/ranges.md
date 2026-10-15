# Range APIs

This domain covers text-range operations. A text range is a live handle created by searching within a paragraph or a table cell.

Use this document when the task says: selected words, matched text, typo word, keyword, replace text, delete text, mark text red, bold text, underline text, highlight text, font size, font family, comment on matched text, or bookmark matched text.

## Skill Summary

Core currently exposes text ranges through object methods:

```ts
const range = paragraph.findText('目标文字', { scope: 'first' });
const cellRange = table.cell(0, 1).findText('目标文字', { scope: 'all' });
```

The range resolves lazily when a write method runs. If no text is found, the operation rejects.

## Range Creators

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `paragraph.findText(search, opts?)` | Search inside one paragraph. | `search: string`, `{scope?:'first'\|'all'}` | `TextRange` |
| `cell.findText(search, opts?)` | Search inside direct paragraphs in one table cell. | `search: string`, `{scope?:'first'\|'all'}` | `CellTextRange` |

`scope`:

| Value | Description |
| --- | --- |
| `'first'` | Apply to the first match. This is the default and safest choice. |
| `'all'` | Apply to all matches found in the paragraph or cell. |

## TextRange Methods

`TextRange` and `CellTextRange` support the same operation set.

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `range.replaceWith(text, opts?)` | Replace matched text with plain text. | `text: string`, `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |
| `range.delete(opts?)` | Delete matched text. | `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |
| `range.setBold(value, opts?)` | Set bold. | `value: boolean` | `Promise<ApplyResult>` |
| `range.setColor(value, opts?)` | Set font color. | `value: string` | `Promise<ApplyResult>` |
| `range.setItalic(value, opts?)` | Set italic. | `value: boolean` | `Promise<ApplyResult>` |
| `range.setUnderline(value, opts?)` | Set single underline. | `value: boolean` | `Promise<ApplyResult>` |
| `range.setHighlight(value, opts?)` | Set Word highlight color. | `value: string` | `Promise<ApplyResult>` |
| `range.setFontSize(value, opts?)` | Set font size in points. | `value: number` | `Promise<ApplyResult>` |
| `range.setFontFamily(value, opts?)` | Set font family. | `value: string` | `Promise<ApplyResult>` |
| `range.addComment(text, opts?)` | Add a comment to matched text. | `text: string`, `LiveObjectWriteOptions?` such as `{author:'halvee'}` | `Promise<ApplyResult>` |
| `range.createBookmark(name, opts?)` | Create a bookmark around matched text. | `name: string` | `Promise<ApplyResult>` |

`LiveObjectWriteOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `trackChanges` | `boolean` | `false` | Request tracked-change markup when supported. |
| `author` | `string` | document default | Author metadata for comments and tracked changes. |
| `date` | `Date` | operation time | Date metadata for comments and tracked changes. |

## Color Values

| API | Value Shape | Example |
| --- | --- | --- |
| `setColor` | OOXML color string, usually `RRGGBB`. | `'FF0000'` |
| `setHighlight` | Word highlight value. | `'yellow'` |

## Examples

Mark all matches in one paragraph red and bold:

```ts
const paragraph = await doc.paragraphs.get('31086A56');
if (paragraph) {
  const range = paragraph.findText('错别字', { scope: 'all' });
  await range.setColor('FF0000');
  await range.setBold(true);
}
```

Replace text inside a table cell:

```ts
const table = await doc.tables.get('tbl-0001');
if (table) {
  await table.cell(0, 0).findText('旧值').replaceWith('新值');
}
```

Add a comment to a matched phrase:

```ts
const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });
for (const paragraph of paragraphs) {
  const text = await paragraph.getText();
  if (text?.includes('请核对')) {
    await paragraph.findText('请核对').addComment('这里需要人工确认。', {
      author: 'Coco',
    });
  }
}
```

## CodeAgent Rules

1. Do not invent coordinate-based APIs such as `plan.format.range()` or `doc.range(start,end)`.
2. To operate on text, first get a paragraph or cell object, then call `findText`.
3. If the same matched text needs multiple styles, call multiple range methods on the same search.
4. For whole paragraph formatting, use paragraph APIs. For only matched words, use range APIs.
5. `CellTextRange` searches cell paragraphs; it is not a separate mini document API.
