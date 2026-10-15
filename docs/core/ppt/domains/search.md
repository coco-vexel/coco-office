# PPT Search Domain

Use this document when a task mentions finding text, replacing text across the
deck, locating a slide/shape by content, or editing text in table cells.

## Skill Summary

Search APIs locate text in shape text and table cells. Always route edits by
`match.source`:

```ts
for (const match of await presentation.findText('TODO')) {
  const shape = presentation.slides.getItem(match.slideId).shapes.getItem(match.shapeId);
  if (match.source === 'shapeText') await shape.text.replace('TODO', 'Done');
  if (match.source === 'tableCell')
    await shape.table.setCellText(match.row!, match.column!, 'Done');
}
```

Use `presentation.replaceText(...)` for simple global replacements. Use
`findText(...)` first when follow-up styling, table routing, or slide-specific
logic is needed.

## APIs

| API                                                    | Returns                         | Scope           |
| ------------------------------------------------------ | ------------------------------- | --------------- |
| `presentation.findText(query, opts?)`                  | `Promise<PptTextMatch[]>`       | All slides.     |
| `presentation.replaceText(search, replacement, opts?)` | `Promise<PptTextReplaceResult>` | All slides.     |
| `slide.findText(query, opts?)`                         | `Promise<PptTextMatch[]>`       | One slide.      |
| `slide.replaceText(search, replacement, opts?)`        | `Promise<PptTextReplaceResult>` | One slide.      |
| `shape.text.replace(search, replacement, opts?)`       | `Promise<PptTextReplaceResult>` | One text shape. |

`query` and `search` may be `string` or `RegExp`.

Options:

```ts
interface PptFindTextOptions {
  caseSensitive?: boolean;
}

interface PptTextReplaceOptions {
  caseSensitive?: boolean;
  preserveFirstRunStyle?: boolean;
}
```

`PptTextReplaceResult`: `{ count: number }`.

`preserveFirstRunStyle` is a compatibility option for aggregate text
replacement. For deterministic formatting, inspect `runs()` and use
`setParagraphs(...)` or `setRunTextStyle(...)`.

## Match Shape

`PptTextMatch`:

```ts
{
  slideId: string;
  slideIndex: number;
  shapeId: string;
  shapeName?: string;
  source: 'shapeText' | 'tableCell';
  text: string;
  start: number;
  length: number;
  row?: number;
  column?: number;
}
```

Tables are returned as `tableCell` matches with `row` and `column`; they are not
also returned as aggregate shape-text matches.

## Examples

Replace all instances:

```ts
const result = await presentation.replaceText('FY24', 'FY25');
console.log(result.count);
```

Find and style only matching text shapes:

```ts
for (const match of await presentation.findText('Urgent', { caseSensitive: true })) {
  if (match.source !== 'shapeText') continue;
  const shape = presentation.slides.getItem(match.slideId).shapes.getItem(match.shapeId);
  await shape.text.setTextStyle({ color: '#FF0000', bold: true });
}
```

Update matching table cells:

```ts
for (const match of await presentation.findText('TODO')) {
  if (match.source !== 'tableCell') continue;
  const table = presentation.slides.getItem(match.slideId).shapes.getItem(match.shapeId).table;
  await table.setCellText(match.row!, match.column!, 'Done');
}
```

## SDK Operation Rules

1. Use `findText` before editing unknown content.
2. Respect `source`: use shape text APIs for `shapeText`, table APIs for `tableCell`.
3. Use `presentation.replaceText` for simple global text replacement.
4. Use `caseSensitive: true` when the user asks for exact case.
5. Do not inspect raw OOXML to find text.
