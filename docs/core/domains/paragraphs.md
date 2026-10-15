# Paragraph APIs

This domain covers paragraph discovery, paragraph live objects, and paragraph-level editing.

Use this document when the task says: paragraph, title, heading, body text, insert paragraph, delete paragraph, paragraph alignment, indentation, spacing, page break, comment on paragraph, bookmark paragraph, or find text inside a paragraph.

## Skill Summary

Paragraph APIs are live-object APIs. `doc.paragraphs.list()` returns `Paragraph` handles, not raw `ParagraphInfo` data. A handle has stable identity fields (`id`, `paraId`, `story`) and methods that read or write the current document state.

Prefer paragraph live-object methods for generated code. Do not generate `doc.plan().formatParagraph(...)` unless you are explicitly writing low-level plan code.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.paragraphs.list(opts?)` | List paragraphs as live handles. | `ListParagraphsOptions` | `Promise<ParagraphList>` |
| `doc.paragraphs.get(paraId)` | Get one paragraph by id. | `paraId: string` | `Promise<Paragraph \| null>` |

`ParagraphList`:

| Field | Type | Description |
| --- | --- | --- |
| `paragraphs` | `Paragraph[]` | Live paragraph handles. |
| `pagination` | `ListResult['pagination']` | Page metadata for the list operation. |

## ListParagraphsOptions

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `scope` | `'body' \| 'all' \| StoryId[]` | `'body'` | Read only main document body, all stories, or explicit stories. |
| `filter` | `ListParagraphsFilter` | none | Filter before pagination. |
| `textFormat` | `'full' \| 'preview'` | `'full'` | `preview` returns the first 100 chars and may set `truncated`. |
| `pagination` | `{mode:'all'} \| {mode:'byCount'; pageSize; pageIndex} \| {mode:'byChar'; charPerPage; pageIndex}` | `{mode:'byCount', pageSize:25, pageIndex:0}` | Controls list size. Use `{mode:'all'}` only when the task needs the whole document. |
| `includeFormat` | `boolean` | `false` | Include direct paragraph formatting in `ParagraphInfo.format`. |
| `structuralWhitespace` | `boolean` | `false` | Preserve tabs and soft line breaks in returned text. |
| `includeTableInfo` | `boolean` | `false` | Include table coordinates for paragraphs inside cells. |

`ListParagraphsFilter`:

| Field | Type | Description |
| --- | --- | --- |
| `styleId` | `string` | Keep paragraphs with this paragraph style id. |
| `contains` | `string` | Keep paragraphs whose text contains this exact substring. Case-sensitive. |
| `hasComments` | `boolean` | Keep paragraphs with or without comment markers. |
| `hasRevisions` | `boolean` | Keep paragraphs with or without insert/delete revision markers. |
| `excludeTableCells` | `boolean` | Exclude paragraphs located inside table cells. |

## ParagraphInfo

`paragraph.getInfo()` returns this data shape, or `null` if the paragraph no longer exists.

| Field | Type | Description |
| --- | --- | --- |
| `paraId` | `string` | Stable paragraph id. |
| `story` | `StoryId` | Story that contains the paragraph, such as body/header/footer/comment. |
| `text` | `string` | Plain text for the paragraph. |
| `truncated` | `boolean?` | Present when `textFormat:'preview'` truncated the text. |
| `styleId` | `string?` | Direct paragraph style id. |
| `hasRevisions` | `boolean?` | Paragraph contains tracked insert/delete markers. |
| `hasComments` | `boolean?` | Paragraph contains comment markers. |
| `hasBookmarks` | `boolean?` | Paragraph contains bookmark markers. |
| `hasMath` | `boolean?` | Paragraph contains OMML math. |
| `format` | `ParagraphFormat?` | Direct paragraph formatting, only with `includeFormat:true`. |
| `table` | `{tableId:string; row:number; col:number}?` | Cell coordinates, only with `includeTableInfo:true` and paragraph is inside a table cell. |

`ParagraphFormat` contains direct OOXML formatting only. It does not resolve inherited style values.

| Field | Type | Description |
| --- | --- | --- |
| `alignment` | `'left' \| 'center' \| 'right' \| 'both' \| 'distribute'` | Horizontal paragraph alignment. |
| `indent` | `{left?:number; right?:number; firstLine?:number}` | Indentation in twips. `firstLine < 0` means hanging indent. |
| `spacing` | `{before?:number; after?:number; line?:number; lineRule?:'auto'\|'exact'\|'atLeast'}` | Paragraph spacing in twips. |
| `fontSize` | `number` | Direct paragraph default font size in points. |
| `bold` | `boolean` | Direct paragraph default bold. |
| `outlineLevel` | `number` | `0` to `8` maps to heading levels 1 to 9. |
| `hasNumbering` | `boolean` | Paragraph has numbering properties. |
| `pageBreakBefore` | `boolean` | Paragraph starts on a new page. |
| `hasSectionBreak` | `boolean` | Paragraph contains section properties. |
| `hasManualPageBreak` | `boolean` | Paragraph contains a manual page break. |
| `color` | `string` | Direct paragraph default color, usually `RRGGBB` or `auto`. |
| `firstRunColor` | `string` | First direct run color, useful for red-header detection. |

## Paragraph Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `paragraph.getInfo(opts?)` | Read current metadata. | `Pick<ListParagraphsOptions, 'includeFormat' \| 'textFormat'>` | `Promise<ParagraphInfo \| null>` |
| `paragraph.getText()` | Read current plain text. | none | `Promise<string \| null>` |
| `paragraph.replaceWith(text, opts?)` | Replace the whole paragraph with plain text. | `text: string`, `ParagraphReplaceOptions?` | `Promise<ApplyResult>` |
| `paragraph.setAlignment(value, opts?)` | Set horizontal alignment. | `'left' \| 'center' \| 'right' \| 'both' \| 'distribute'` | `Promise<ApplyResult>` |
| `paragraph.setOutlineLevel(value, opts?)` | Set heading outline level. | `number` from `0` to `8` | `Promise<ApplyResult>` |
| `paragraph.setPageBreakBefore(value, opts?)` | Start paragraph on a new page. | `boolean` | `Promise<ApplyResult>` |
| `paragraph.setIndent(value, opts?)` | Set paragraph indentation. | `{left?:number; right?:number; firstLine?:number}` | `Promise<ApplyResult>` |
| `paragraph.setSpacing(value, opts?)` | Set paragraph spacing. | `{before?:number; after?:number; line?:number; lineRule?:'auto'\|'exact'\|'atLeast'}` | `Promise<ApplyResult>` |
| `paragraph.insertBefore(text, opts?)` | Insert a paragraph before this one. | `text: string`, `ParagraphInsertOptions?` | `Promise<ApplyResult>` |
| `paragraph.insertAfter(text, opts?)` | Insert a paragraph after this one. | `text: string`, `ParagraphInsertOptions?` | `Promise<ApplyResult>` |
| `paragraph.insertImageAfter(spec, opts?)` | Insert a new inline image paragraph after this one. | `ParagraphInsertImageSpec`, `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |
| `paragraph.insertTableAfter(spec, opts?)` | Insert a simple text table after this one. | `ParagraphInsertTableSpec`, `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |
| `paragraph.delete(opts?)` | Delete this paragraph. | `ParagraphDeleteOptions?` | `Promise<ApplyResult>` |
| `paragraph.addComment(text, opts?)` | Add a comment anchored to this paragraph. | `text: string`, `LiveObjectWriteOptions?` such as `{author:'halvee'}` | `Promise<ApplyResult>` |
| `paragraph.createBookmark(name, opts?)` | Create a bookmark around this paragraph. | `name: string` | `Promise<ApplyResult>` |
| `paragraph.findText(search, opts?)` | Create a text range handle inside this paragraph. | `search: string`, `{scope?:'first'\|'all'}` | `TextRange` |

`LiveObjectWriteOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `trackChanges` | `boolean` | `false` | Request tracked-change markup when the operation supports it. |
| `author` | `string` | document default | Author metadata for comments and tracked changes. |
| `date` | `Date` | operation time | Date metadata for comments and tracked changes. |

`ParagraphReplaceOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `trackChanges` | `boolean` | `false` | Request tracked-change markup. |
| `author` | `string` | document default | Author metadata for tracked changes. |
| `date` | `Date` | operation time | Date metadata for tracked changes. |
| `onAnchorConflict` | `'error' \| 'drop'` | `'error'` | How to handle bookmarks/comments affected by replacement. |
| `keepInlineGraphics` | `boolean` | `false` | Preserve inline drawings/objects while replacing text runs. Use this when replacing paragraph text but keeping pictures/shapes embedded in the paragraph. |

`ParagraphInsertOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `trackChanges` | `boolean` | `false` | Request tracked-change markup. |
| `author` | `string` | document default | Author metadata for tracked changes. |
| `date` | `Date` | operation time | Date metadata for tracked changes. |
| `style` | `'inherit' \| string` | implementation default | Use inherited style or explicit style id. |
| `runFormat` | `'inherit' \| 'default'` | implementation default | Control inserted run formatting. |

`ParagraphInsertImageSpec`:

| Field | Type | Description |
| --- | --- | --- |
| `bytes` | `Uint8Array` | Image bytes. |
| `contentType` | `string` | Image MIME type, such as `image/png` or `image/jpeg`. |
| `widthEmu` | `number` | Display width in EMU. 1 inch = 914400 EMU. |
| `heightEmu` | `number` | Display height in EMU. |

`ParagraphInsertTableSpec`:

| Field | Type | Description |
| --- | --- | --- |
| `rows` | `string[][]` | Text matrix for a simple table. Empty missing cells are filled in for ragged rows. |
| `columnWidths` | `number[]?` | Optional column widths in twips. Defaults to 2400 twips per column. |
| `tableStyleId` | `string?` | Optional Word table style id. |

`insertTableAfter` creates a simple text-only table. For high-fidelity table copying, use `table.getVerbatim()` plus `doc.plan().insertTableFrom(...)`.

`ParagraphDeleteOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `trackChanges` | `boolean` | `false` | Request tracked-change markup. |
| `author` | `string` | document default | Author metadata for tracked changes. |
| `date` | `Date` | operation time | Date metadata for tracked changes. |
| `onAnchorConflict` | `'error' \| 'drop'` | `'error'` | How to handle bookmarks/comments affected by deletion. |

## Examples

List all body paragraphs:

```ts
const { paragraphs } = await doc.paragraphs.list({
  pagination: { mode: 'all' },
});
```

Find a heading and insert a paragraph after it:

```ts
const { paragraphs } = await doc.paragraphs.list({
  filter: { styleId: 'Heading1' },
  pagination: { mode: 'all' },
});

for (const paragraph of paragraphs) {
  const text = await paragraph.getText();
  if (text?.includes('风险提示')) {
    await paragraph.insertAfter('请补充风险控制措施。', { style: 'inherit' });
  }
}
```

Set paragraph formatting:

```ts
const paragraph = await doc.paragraphs.get('31086A56');
if (paragraph) {
  await paragraph.setAlignment('center');
  await paragraph.setSpacing({ before: 240, after: 120 });
}
```

Replace paragraph text while preserving inline pictures:

```ts
const paragraph = await doc.paragraphs.get('31086A56');
await paragraph?.replaceWith('新的段落文本', { keepInlineGraphics: true });
```

Find text inside a paragraph and format it:

```ts
const paragraph = await doc.paragraphs.get('31086A56');
await paragraph?.findText('错别字', { scope: 'all' }).setBold(true);
await paragraph?.findText('错别字', { scope: 'all' }).setColor('FF0000');
```

Insert an image after a paragraph:

```ts
await paragraph?.insertImageAfter({
  bytes: imageBytes,
  contentType: 'image/png',
  widthEmu: 914400,
  heightEmu: 914400,
});
```

Insert a simple table after a paragraph:

```ts
await paragraph?.insertTableAfter({
  rows: [
    ['Name', 'Score'],
    ['Ada', '42'],
  ],
  columnWidths: [2400, 1200],
});
```

## CodeAgent Rules

1. For ordinary document editing, use live objects: `doc.paragraphs.list/get` then `paragraph.*`.
2. Use `{ pagination: { mode: 'all' } }` only when the user explicitly asks for whole-document work.
3. `paragraph.findText()` returns a `TextRange`; text formatting belongs to the range, not to `ParagraphInfo`.
4. `getInfo()` and `getText()` may return `null`; generated code must handle missing objects when using stored ids.
5. Do not assume inherited style values are present in `ParagraphFormat`.
6. Use `paragraph.insertImageAfter(...)` and `paragraph.insertTableAfter(...)` for simple anchored content insertion.
