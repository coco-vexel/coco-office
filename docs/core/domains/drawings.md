# Drawing APIs

This domain covers pictures, text boxes, shapes, and other drawing objects.

Use this document when the task says: image, picture, drawing, shape, text box, floating object, inline object, watermark-like object, replace picture with placeholder, delete shape, or read text box content.

## Skill Summary

Drawing is the OOXML container-level concept. For small-model code generation, prefer the precise typed list method:

```ts
const pictures = await doc.drawings.listPictures();
const textBoxes = await doc.drawings.listTextBoxes();
const shapes = await doc.drawings.listShapes();
```

Do not treat all drawings as pictures. `getImageBytes()` returns bytes only when the drawing has an image reference.

To insert a new simple inline image, use a paragraph anchor:

```ts
await paragraph.insertImageAfter({
  bytes,
  contentType: 'image/png',
  widthEmu: 914400,
  heightEmu: 914400,
});
```

The drawing domain edits existing drawings. Creating new arbitrary shapes or text boxes is not exposed as a direct live-object API yet.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.drawings.list(opts?)` | List all drawing objects. | `ListDrawingsOptions` | `Promise<Drawing[]>` |
| `doc.drawings.listPictures(opts?)` | List picture drawings only. | `ListDrawingsOptions` | `Promise<PictureDrawing[]>` |
| `doc.drawings.listTextBoxes(opts?)` | List text box drawings only. | `ListDrawingsOptions` | `Promise<TextBoxDrawing[]>` |
| `doc.drawings.listShapes(opts?)` | List plain shape drawings only. | `ListDrawingsOptions` | `Promise<ShapeDrawing[]>` |
| `doc.drawings.get(drawingId)` | Get one drawing. | `drawingId: string` | `Promise<Drawing \| null>` |

## ListDrawingsOptions

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `scope` | `'body' \| 'all'` | `'body'` | Body only or all stories including headers, footers, footnotes, endnotes, and comments. |
| `structuralWhitespace` | `boolean` | `false` | Preserve tabs and soft line breaks in text box `innerText`. |

## DrawingInfo

`Drawing`, `PictureDrawing`, `TextBoxDrawing`, and `ShapeDrawing` extend `DrawingInfo`.

| Field | Type | Description |
| --- | --- | --- |
| `drawingId` / `id` | `string` | Stable drawing id, such as `dr-0001`. |
| `partPath` | `string` | OOXML part path. |
| `index` | `number` | 1-based order in the list result. |
| `anchoringMode` | `'inline' \| 'anchor'` | Inline or floating anchored object. |
| `kind` | `DrawingKind` | OOXML content kind. |
| `extentEMU` | `{cx:number; cy:number}` | Display size in EMUs. |
| `wrap` | `DrawingWrapType?` | Text wrapping, only for anchored drawings. |
| `position` | `DrawingPosition?` | Floating position, only for anchored drawings. |
| `anchorParaId` | `string?` | Paragraph carrying the drawing when available. |
| `runIndex` | `number?` | Direct run index in the anchor paragraph when available. |
| `isInHeaderFooter` | `boolean` | Whether the drawing is in a header/footer part. |
| `hasPicReference` | `boolean` | Whether the subtree references picture bytes. |
| `isLineOrConnector` | `boolean?` | Shape-only signal for line/connector geometry. |
| `hasBorder` | `boolean?` | Shape/textbox border signal. |
| `hasFill` | `boolean?` | Shape/textbox fill signal. |
| `hasInnerText` | `boolean?` | Textbox has non-empty text. |
| `innerText` | `string?` | Textbox inner text snapshot. |
| `innerParagraphIds` | `string[]?` | Textbox internal paragraph ids. |
| `innerBlocks` | `TextBoxBlockInfo[]?` | Textbox direct content blocks in order. |
| `innerTextAlign` | `string?` | First inner paragraph alignment. |

`DrawingKind`:

| Value | Meaning |
| --- | --- |
| `picture` | Embedded or linked picture. |
| `chart` | Native chart. |
| `smartArt` | SmartArt/diagram. |
| `shape` | Plain word-processing shape without textbox content. |
| `textBox` | Shape containing textbox content. |
| `group` | Grouped drawing. |
| `canvas` | Drawing canvas. |
| `other` | Unknown/unsupported drawing kind. |

`DrawingWrapType`: `'none' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behind' | 'inFront'`.

## Drawing Object

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `drawing.getInfo()` | Read current metadata. | none | `Promise<DrawingInfo \| null>` |
| `drawing.getImageBytes()` | Read original image bytes when available. | none | `Promise<DrawingImageBytes \| null>` |
| `drawing.delete(opts?)` | Delete this drawing. | `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |
| `drawing.replaceWithPlaceholder(opts?)` | Replace picture drawing with shared placeholder image. | `LiveObjectWriteOptions?` | `Promise<ApplyResult>` |
| `drawing.replaceShapeInnerText(placeholder, opts?)` | Replace shape/textbox inner text. | `placeholder: string` | `Promise<ApplyResult>` |

`DrawingImageBytes`:

| Field | Type | Description |
| --- | --- | --- |
| `bytes` | `Uint8Array` | Original image bytes. |
| `contentType` | `string` | MIME content type. |
| `extentEMU` | `{cx:number; cy:number}` | Display size. |

## PictureDrawing

`PictureDrawing` has `kind: 'picture'` and supports:

| API | Purpose |
| --- | --- |
| `picture.getImageBytes()` | Read original image bytes. |
| `picture.replaceWithPlaceholder(opts?)` | Replace with placeholder while preserving layout. |
| `picture.delete(opts?)` | Delete picture drawing. |

## TextBoxDrawing

`TextBoxDrawing` has `kind: 'textBox'`.

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `textBox.getText()` | Read current textbox text. | none | `Promise<string \| null>` |
| `textBox.getParagraphs()` | Get paragraph handles inside textbox. | none | `Promise<Paragraph[]>` |
| `textBox.getBlocks()` | Get direct textbox blocks in order. | none | `Promise<TextBoxBlock[]>` |
| `textBox.replaceText(text, opts?)` | Replace all textbox content with one paragraph. | `text: string` | `Promise<ApplyResult>` |

`TextBoxBlock`:

| Shape | Description |
| --- | --- |
| `{type:'paragraph'; paragraph; paraId; text; index}` | Direct textbox paragraph. |
| `{type:'drawing'; drawing; drawingId; paraId?; index}` | Direct nested drawing. |
| `{type:'table'; table; tableId; index}` | Direct nested table. |

## ShapeDrawing

`ShapeDrawing` has `kind: 'shape'`.

| API | Purpose | Returns |
| --- | --- | --- |
| `shape.getInfo()` | Read current shape metadata. | `Promise<(DrawingInfo & {kind:'shape'}) \| null>` |
| `shape.delete(opts?)` | Delete this shape. | `Promise<ApplyResult>` |

## Capability Matrix

| Capability | `picture` | `textBox` | `shape` | Other kinds |
| --- | --- | --- | --- | --- |
| List | yes | yes | yes | via `list()` |
| Read metadata | yes | yes | yes | yes |
| Read image bytes | yes | only if image-backed | no | maybe `null` |
| Delete | yes | yes | yes | yes |
| Replace with placeholder | yes | not primary | not primary | not primary |
| Read inner text | no | yes | no | no |
| Replace inner text | no | yes | via `replaceShapeInnerText` if supported | no |
| Read ordered inner blocks | no | yes | no | no |
| Create new object | via `paragraph.insertImageAfter(...)` | no direct API yet | no direct API yet | no |

## Examples

Read all pictures:

```ts
const pictures = await doc.drawings.listPictures({ scope: 'all' });
for (const picture of pictures) {
  const image = await picture.getImageBytes();
  console.log(picture.drawingId, image?.contentType);
}
```

Replace all text boxes containing a word:

```ts
const textBoxes = await doc.drawings.listTextBoxes();
for (const box of textBoxes) {
  const text = await box.getText();
  if (text?.includes('占位')) {
    await box.replaceText('新文本');
  }
}
```

Delete line shapes:

```ts
const shapes = await doc.drawings.listShapes();
for (const shape of shapes) {
  const info = await shape.getInfo();
  if (info?.isLineOrConnector) {
    await shape.delete();
  }
}
```

Operate on paragraphs inside a text box:

```ts
const boxes = await doc.drawings.listTextBoxes();
for (const box of boxes) {
  for (const paragraph of await box.getParagraphs()) {
    const text = await paragraph.getText();
    if (text?.includes('错误')) {
      await paragraph.findText('错误').setColor('FF0000');
    }
  }
}
```

## CodeAgent Rules

1. Prefer `listPictures`, `listTextBoxes`, or `listShapes` when the user asks for a specific kind.
2. Check `kind` before calling kind-specific behavior on objects returned by `list()`.
3. `getImageBytes()` may return `null`; do not assume every drawing is an image.
4. Use `textBox.getParagraphs()` or `textBox.getBlocks()` for content-aware textbox operations.
5. Do not generate APIs for charts, SmartArt, groups, or canvases beyond list/get/read metadata/delete unless documented here.
6. Insert new pictures through `paragraph.insertImageAfter(...)`; do not invent drawing creation APIs for shapes or text boxes.
