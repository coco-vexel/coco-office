# PPT Slides Domain

Use this document when a task mentions listing slides, reading titles, adding,
copying, moving, deleting, or operating within a specific slide.

## Skill Summary

Slides are addressed by zero-based index, slide id, or name. Use slide handles
to scope content operations:

```ts
const slide = presentation.slides.getItem(0);
await slide.setTitle('Q4 Review');
const matches = await slide.findText('Draft');
```

Slide layout, transition, and timing APIs are read-only inspection APIs in the
current core SDK. Use `slide.shapes` for slide content editing.

## Collection APIs

```ts
const count = await presentation.slides.getCount();
const slide = presentation.slides.getItem(0);
const maybe = await presentation.slides.getItemOrNullObject('256');
const slides = await presentation.slides.items();
```

| API                                                  | Returns                  | Notes                                                           |
| ---------------------------------------------------- | ------------------------ | --------------------------------------------------------------- |
| `presentation.slides.items()`                        | `Promise<Slide[]>`       | Visual order.                                                   |
| `presentation.slides.getCount()`                     | `Promise<number>`        | Slide count.                                                    |
| `presentation.slides.getItem(indexOrId)`             | `Slide`                  | 0-based index or id/name.                                       |
| `presentation.slides.getItemOrNullObject(indexOrId)` | `Promise<Slide \| null>` | Existence check.                                                |
| `presentation.slides.add(options?)`                  | `Promise<Slide>`         | Options include `index?: number \| 'end'` and `title?: string`. |
| `presentation.slides.copy(indexOrId, options?)`      | `Promise<Slide>`         | Copies slide XML and slide rels.                                |
| `presentation.slides.move(indexOrId, toIndex)`       | `Promise<void>`          | 0-based destination index.                                      |
| `presentation.slides.deleteItem(indexOrId)`          | `Promise<void>`          | Rejects deleting final slide.                                   |

## Slide Options

`PptSlideAddOptions`:

| Field   | Type              | Description                                                                                                               |
| ------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `index` | `number \| 'end'` | 0-based insertion index, or append with `'end'`. Out-of-range numeric indexes are clamped into the valid insertion range. |
| `title` | `string`          | Optional title text for the new slide.                                                                                    |

`PptSlideCopyOptions`:

| Field   | Type              | Description                                                                                                               |
| ------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `index` | `number \| 'end'` | 0-based insertion index, or append with `'end'`. Out-of-range numeric indexes are clamped into the valid insertion range. |

## Slide APIs

```ts
const slide = presentation.slides.getItem(0);
const id = await slide.id();
const title = await slide.title();
await slide.setTitle('Quarterly Review');
```

| API                                             | Returns                                    | Notes                                  |
| ----------------------------------------------- | ------------------------------------------ | -------------------------------------- |
| `slide.id()`                                    | `Promise<string>`                          | OOXML slide id.                        |
| `slide.index()`                                 | `Promise<number>`                          | 0-based index.                         |
| `slide.name()`                                  | `Promise<string \| undefined>`             | Optional model name.                   |
| `slide.title()`                                 | `Promise<string \| undefined>`             | Derived from text/placeholder content. |
| `slide.layout()`                                | `Promise<PptSlideLayoutModel \| null>`     | Layout/master relationship metadata.   |
| `slide.transition()`                            | `Promise<PptSlideTransitionModel \| null>` | Slide transition metadata, read-only.  |
| `slide.timing()`                                | `Promise<PptSlideTimingModel \| null>`     | Animation timing summary, read-only.   |
| `slide.properties()`                            | `Promise<PptSlideProperties>`              | Counts and metadata.                   |
| `slide.setTitle(text, opts?)`                   | `Promise<void>`                            | Uses first text shape as title target. |
| `slide.notes.hasText()`                         | `Promise<boolean>`                         | True when speaker notes text exists.   |
| `slide.notes.getText()`                         | `Promise<string>`                          | Speaker notes text.                    |
| `slide.notes.setText(text)`                     | `Promise<void>`                            | Writes speaker notes text.             |
| `slide.notes.clear()`                           | `Promise<void>`                            | Clears speaker notes text.             |
| `slide.comments.items()`                        | `Promise<PptCommentModel[]>`               | Slide comments.                        |
| `slide.comments.getCount()`                     | `Promise<number>`                          | Slide comment count.                   |
| `slide.comments.add(options)`                   | `Promise<PptCommentModel>`                 | Adds a slide comment.                  |
| `slide.comments.delete(id)`                     | `Promise<void>`                            | Deletes a slide comment by id.         |
| `slide.comments.clear()`                        | `Promise<void>`                            | Deletes all comments on this slide.    |
| `slide.findText(query, opts?)`                  | `Promise<PptTextMatch[]>`                  | Slide-scoped search.                   |
| `slide.replaceText(search, replacement, opts?)` | `Promise<PptTextReplaceResult>`            | Slide-scoped replace.                  |
| `slide.duplicate(options?)`                     | `Promise<Slide>`                           | Copies this slide.                     |
| `slide.moveTo(index)`                           | `Promise<void>`                            | Moves this slide.                      |
| `slide.delete()`                                | `Promise<void>`                            | Deletes this slide.                    |
| `slide.toModel()`                               | `Promise<PptSlideModel>`                   | Defensive snapshot.                    |

## Examples

Add a titled slide:

```ts
const slide = await presentation.slides.add({ title: 'Appendix', index: 'end' });
```

Move the last slide to the front:

```ts
const slides = await presentation.slides.items();
const last = slides[slides.length - 1];
await last.moveTo(0);
```

Read and write notes and comments:

```ts
const notes = await slide.notes.getText();
const comments = await slide.comments.items();

await slide.notes.setText('Talk track for this slide');
await slide.comments.add({
  text: 'Review this claim',
  author: 'Coco Reviewer',
  initials: 'CR',
});
```

`PptCommentAddOptions`:

| Field      | Type     | Description                           |
| ---------- | -------- | ------------------------------------- |
| `text`     | `string` | Comment body.                         |
| `author`   | `string` | Optional display author.              |
| `initials` | `string` | Optional author initials.             |
| `date`     | `string` | Optional ISO-like date string.        |
| `x`        | `number` | Optional comment anchor x coordinate. |
| `y`        | `number` | Optional comment anchor y coordinate. |

Read layout and master metadata:

```ts
const layout = await slide.layout();
console.log(layout?.name, layout?.master?.name);
```

Read transition and animation timing metadata:

```ts
const transition = await slide.transition();
const timing = await slide.timing();
console.log(transition?.type, transition?.speed, timing?.timeNodeCount);
```

Delete slides with no shapes, keeping at least one slide:

```ts
for (const slide of await presentation.slides.items()) {
  if ((await presentation.slides.getCount()) <= 1) break;
  if ((await slide.properties()).shapeCount === 0) {
    await slide.delete();
  }
}
```

## SDK Operation Rules

1. Slide indexes are 0-based.
2. Use `getItemOrNullObject` when absence is acceptable.
3. Do not delete the final slide.
4. Use `slide.shapes` for content operations; retrieve [shapes.md](./shapes.md).
5. Use slide-scoped `findText`/`replaceText` when the user names a specific slide.
6. Use `slide.notes` and `slide.comments` for speaker notes and comment operations; they create missing package parts when needed.
7. Treat `slide.layout()` as read-only metadata; layout switching is not implemented.
8. Treat `slide.transition()` and `slide.timing()` as read-only inspection APIs; animation and transition editing are not implemented.
9. Numeric insertion and move destinations are clamped into valid ranges; use `getCount()` first when exact placement matters.
