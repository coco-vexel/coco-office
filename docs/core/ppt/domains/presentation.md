# PPT Presentation Domain

Use this document when a task mentions loading/exporting a PPTX, deck-level
properties, global search/replace, media/font export, or lifecycle.

## Skill Summary

`Presentation` is the SDK root. Use it to read deck-level properties, search
across slides, export package/media/font bytes, and access slide collections:

```ts
const props = await presentation.properties();
const matches = await presentation.findText('Revenue');
const slide = presentation.slides.getItem(0);
```

SDK operation code should usually receive a ready `presentation` from the
runtime. Loading from bytes and writing `toZip()` belong to file/runtime
boundaries unless the task explicitly asks for package input or output.

## Import

```ts
import { Presentation } from '@coco-vexel/coco-office/ppt';
```

SDK operation code usually receives a ready `presentation: Presentation` from
the runtime. Do not call host APIs here.

## Loading And Export

```ts
const presentation = await Presentation.fromZip(bytes, {
  presentationId: 'deck.pptx',
  mode: 'auto',
});

const output = await presentation.toZip();
await presentation.dispose();
```

APIs:

| API                                    | Returns                         | Notes                                  |
| -------------------------------------- | ------------------------------- | -------------------------------------- |
| `Presentation.fromZip(bytes, opts?)`   | `Promise<Presentation>`         | Loads `.pptx` package bytes.           |
| `Presentation.fromModel(model, opts?)` | `Promise<Presentation>`         | Creates package from normalized model. |
| `Presentation.createBlank(opts?)`      | `Promise<Presentation>`         | One blank slide.                       |
| `presentation.toModel()`               | `Promise<PptPresentationModel>` | Defensive snapshot.                    |
| `presentation.toZip()`                 | `Promise<Uint8Array>`           | `.pptx` bytes.                         |
| `presentation.dispose()`               | `Promise<void>`                 | Releases runtime state.                |

## Themes

```ts
const themes = await presentation.themes.items();
for (const theme of themes) {
  console.log(theme.name, theme.colorScheme, theme.fontScheme);
}
```

| API                              | Returns                    | Notes                         |
| -------------------------------- | -------------------------- | ----------------------------- |
| `presentation.themes.items()`    | `Promise<PptThemeModel[]>` | Read-only theme snapshots.    |
| `presentation.themes.getCount()` | `Promise<number>`          | Number of parsed theme parts. |

Themes are parsed from presentation/master relationships. They include theme
name, part path, basic color scheme, and major/minor latin fonts.

## Properties

```ts
const props = await presentation.properties();
```

`PptPresentationProperties` includes:

- `presentationId`
- `slideCount`
- `shapeCount`
- `textShapeCount`
- `tableCount`
- `pictureCount`
- `mediaCount`: package media parts under `ppt/media/*`; this can include image,
  audio, or video parts.
- `slideSize`

## Global Search And Replace

```ts
const matches = await presentation.findText('Q3');
const result = await presentation.replaceText('Q3', 'Q4');
```

Search covers text shapes and table cells. Tables are reported as `tableCell`
matches, not duplicate aggregate shape-text matches.

## Export Package Media

```ts
const images = await presentation.exportImages();
for (const image of images) {
  console.log(image.path, image.contentType, image.byteLength, image.bytes);
}
```

`exportImages()` returns package media bytes from `ppt/media/*`. Despite the
legacy API name and `PptExportedImage` type, this is a package media export and
can include image, audio, or video parts when those files live under
`ppt/media/*`. It does not render slides, shapes, or thumbnails.

## Export Fonts

```ts
const fonts = await presentation.exportFonts();
for (const font of fonts) {
  console.log(font.family, font.path, font.contentType, font.byteLength, font.bytes);
}
```

`exportFonts()` returns embedded font bytes discovered from presentation
relationships.

## Mutation Log

```ts
const mutations = await presentation.getMutations();
await presentation.clearMutations();
```

Ordinary SDK operation code should not normally call these. Host runtimes use
them to optimize writeback before falling back to package replacement.

## SDK Operation Rules

1. Assume `presentation` is already initialized by the host/runtime.
2. Use `toZip()` only when output package bytes are needed.
3. Treat `presentation.themes` as read-only metadata.
4. Use `properties()` for counts instead of recalculating from snapshots unless the task needs raw model details.
5. Use `findText`/`replaceText` for global text tasks; retrieve [search.md](./search.md) for details.
6. Use `exportFonts()` when embedded font bytes are required.
7. Do not use host APIs or raw OOXML in SDK operation code.
