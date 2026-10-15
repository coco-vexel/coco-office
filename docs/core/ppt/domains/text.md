# PPT Text Domain

Use this document when a task mentions text boxes, changing text, replacing
text inside one shape, font size/color, bold, italic, or paragraph alignment.

## Skill Summary

Text APIs operate on shape text bodies. Check `hasText()` before editing an
unknown shape:

```ts
if (await shape.text.hasText()) {
  await shape.text.replace('Draft', 'Final');
  await shape.text.setTextStyle({ bold: true, color: '#1F4E79' });
}
```

Use aggregate setters for simple text. They may collapse multi-run text content.
Use `setTextStyle(...)` for whole-shape text formatting. Use `paragraphs()`,
`runs()`, `setParagraphs(...)`, and `setRunTextStyle(...)` only when run-level
formatting matters.

## Text Handles

Every `Shape` has `shape.text`. Check `hasText()` before operating on unknown
shapes.

```ts
for (const shape of await slide.shapes.items()) {
  if (await shape.text.hasText()) {
    await shape.text.setText('Updated');
  }
}
```

## APIs

| API                                              | Returns                          | Notes                                                                              |
| ------------------------------------------------ | -------------------------------- | ---------------------------------------------------------------------------------- |
| `shape.text.hasText()`                           | `Promise<boolean>`               | True when the shape has text runs.                                                 |
| `shape.text.getText()`                           | `Promise<string>`                | Aggregated text.                                                                   |
| `shape.text.paragraphs()`                        | `Promise<PptTextParagraph[]>`    | Paragraph/run structure.                                                           |
| `shape.text.runs()`                              | `Promise<PptTextRun[]>`          | Flattened text runs, including run hyperlinks.                                     |
| `shape.text.body()`                              | `Promise<PptTextBodyProperties>` | Text body anchor, auto-fit, wrap, vertical mode, and insets.                       |
| `shape.text.setParagraphs(paragraphs)`           | `Promise<void>`                  | Rebuilds structured paragraphs and runs.                                           |
| `shape.text.setBodyProperties(body)`             | `Promise<void>`                  | Updates body properties such as auto-fit without rebuilding text.                  |
| `shape.text.setTextStyle(style)`                 | `Promise<void>`                  | Preferred whole-shape text style update.                                           |
| `shape.text.setRunTextStyle(index, style)`       | `Promise<void>`                  | Preferred explicit run style update.                                               |
| `shape.text.setRunStyle(index, style)`           | `Promise<void>`                  | Compatibility alias for `setRunTextStyle`.                                         |
| `shape.text.setText(text, opts?)`                | `Promise<void>`                  | Writes first run and clears extra text runs.                                       |
| `shape.text.replace(search, replacement, opts?)` | `Promise<PptTextReplaceResult>`  | One-shape replace.                                                                 |
| `shape.text.style()`                             | `Promise<PptTextStyle>`          | Font family, size, color, emphasis, alignment, line height, and character spacing. |
| `shape.text.setStyle(style)`                     | `Promise<void>`                  | Compatibility alias for `setTextStyle`.                                            |
| `shape.text.setBold(enabled)`                    | `Promise<void>`                  | Convenience setter.                                                                |
| `shape.text.setItalic(enabled)`                  | `Promise<void>`                  | Convenience setter.                                                                |
| `shape.text.setAlignment(alignment)`             | `Promise<void>`                  | `left`, `center`, `right`, `justify`, `distributed`.                               |

## Option Types

`PptTextWriteOptions`:

| Field                   | Type      | Description                                                                                                                                                          |
| ----------------------- | --------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `preserveFirstRunStyle` | `boolean` | Compatibility option for aggregate setters. Current core preserves existing first-run XML styling when possible; use structured run APIs when style control matters. |

`PptTextReplaceOptions` extends `PptTextWriteOptions`:

| Field           | Type      | Description                   |
| --------------- | --------- | ----------------------------- |
| `caseSensitive` | `boolean` | Match text with exact casing. |

`PptTextStyle`:

| Field              | Type                    | Description                                                       |
| ------------------ | ----------------------- | ----------------------------------------------------------------- |
| `fontSize`         | `number`                | Font size.                                                        |
| `fontFamily`       | `string`                | Latin/East Asian/complex-script typeface.                         |
| `color`            | `string`                | `#RRGGBB` or `RRGGBB`.                                            |
| `bold`             | `boolean`               | Bold text.                                                        |
| `italic`           | `boolean`               | Italic text.                                                      |
| `underline`        | `boolean`               | Single underline when true, no underline when false.              |
| `strike`           | `boolean`               | Single strike when true, no strike when false.                    |
| `baseline`         | `number`                | Baseline shift ratio; positive raises, negative lowers text.      |
| `highlight`        | `string`                | Highlight color as `#RRGGBB` or `RRGGBB`.                         |
| `alignment`        | `PptParagraphAlignment` | `left`, `center`, `right`, `justify`, `distributed`.              |
| `lineHeight`       | `number`                | Paragraph line height ratio, written as OOXML percentage spacing. |
| `characterSpacing` | `number`                | Character spacing using the SDK's normalized point-scale value.   |

`PptTextStyle` also reserves read-model fields such as `textFill` and
`textOutlineColor`. These are not part of the P0 public write surface yet;
retrieve [../style-capabilities.md](../style-capabilities.md) when deciding
whether a style operation is writable.

## Examples

Update all text boxes containing a phrase:

```ts
for (const match of await presentation.findText('Draft')) {
  const slide = presentation.slides.getItem(match.slideId);
  const shape = slide.shapes.getItem(match.shapeId);
  if (match.source === 'shapeText') {
    await shape.text.replace('Draft', 'Final');
  }
}
```

Set title styling:

```ts
const title = presentation.slides.getItem(0).shapes.getItem('Title');
await title.text.setTextStyle({
  fontSize: 28,
  color: '#1F4E79',
  bold: true,
  alignment: 'center',
});
```

Read paragraph and run structure:

```ts
const paragraphs = await shape.text.paragraphs();
for (const paragraph of paragraphs) {
  for (const run of paragraph.runs) {
    console.log(run.text, run.style);
  }
}
```

Read and write hyperlinks on text runs:

```ts
for (const run of await shape.text.runs()) {
  if (run.hyperlink?.target) {
    console.log(run.text, run.hyperlink.target);
  }
}

await shape.text.setRunHyperlink(0, {
  target: 'https://example.com',
  tooltip: 'Open link',
});
await shape.text.clearRunHyperlink(0);
```

Write structured paragraphs and runs:

```ts
await shape.text.setParagraphs([
  {
    style: {
      alignment: 'center',
      bullet: '•',
      bulletColor: '#1F4E79',
      bulletFontFamily: 'Arial',
      lineHeight: 1.25,
      marginLeft: 342900,
      textIndent: -171450,
      spaceBefore: 1270,
      spaceAfter: 2540,
    },
    runs: [
      { text: 'Hello ', style: { bold: true, color: '#112233' } },
      { text: 'Coco', style: { italic: true } },
    ],
  },
]);

await shape.text.setRunTextStyle(1, { color: '#445566', bold: false });
```

`setParagraphs` replaces the paragraph nodes inside the existing text body.
`setRunTextStyle`, `setRunHyperlink`, and `clearRunHyperlink` use the flattened
`runs()` order. `setStyle` and `setRunStyle` remain compatibility aliases, but
SDK operation code should prefer `setTextStyle` and `setRunTextStyle` because
their scope is explicit. Structured paragraph writers do not create hyperlink
runs; use `setRunHyperlink` after rebuilding text when links are needed.
Paragraph `style` supports `alignment`, `bullet`, `bulletColor`,
`bulletFontFamily`, `bulletSize`, `lineHeight`, `marginLeft`, `textIndent`,
`spaceBefore`, and `spaceAfter`.

## Search Versus Shape Text

Use `presentation.findText` / `presentation.replaceText` for cross-deck text
tasks. Use `shape.text.replace` only when the target shape is already known.
Retrieve [search.md](./search.md) for cross-deck semantics.

## SDK Operation Rules

1. Call `hasText()` before editing unknown shapes.
2. Do not assume table text should be edited through `shape.text`; use table cell APIs.
3. Use `presentation.replaceText` for global replacement.
4. Use `setTextStyle` for whole-title or whole-text-box style changes in one call.
5. Colors should be `#RRGGBB` or `RRGGBB`.
6. Use `setParagraphs(...)` for structured rewrites when run-level formatting must be preserved.
7. Use `setRunTextStyle(index, style)` only after inspecting `runs()` or when the run index is known.
8. Use `setRunHyperlink` / `clearRunHyperlink` for run links; do not inject raw `a:hlinkClick`.
