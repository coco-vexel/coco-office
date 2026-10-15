# PPT Images Domain

Use this document when a task mentions pictures, image bytes, replacing images,
adding images, or exporting all images.

## Skill Summary

Images are represented as picture shapes plus package media bytes. Add new
image shapes through `slide.shapes.addImage(...)`; replace existing picture
bytes through `shape.image.replace(...)`:

```ts
if (await shape.image.isImage()) {
  await shape.image.setCrop({ left: 0.08, right: 0.08 });
  await shape.image.setFillRect({ left: 0, right: 0 });
  await shape.image.setOpacity(0.85);
  await shape.image.replace({ bytes: logoBytes, mimeType: 'image/png' });
}
```

`presentation.exportImages()` extracts package media from `ppt/media/*`. The
method name is legacy: the result can include non-picture media parts such as
audio or video files when they are stored in that package directory. It does
not render slide thumbnails or shape screenshots.

## Add Image

```ts
await slide.shapes.addImage(
  { bytes: imageBytes, mimeType: 'image/png', filename: 'logo.png' },
  {
    name: 'Logo',
    geometry: { left: 500000, top: 500000, width: 1200000, height: 600000 },
  },
);
```

`PptImageSource` may be:

- `Uint8Array`
- `ArrayBuffer`
- `{ bytes: Uint8Array | ArrayBuffer; mimeType?: string; filename?: string }`
- `{ base64: string; mimeType?: string; filename?: string }`

`PptImageAddOptions` extends `PptShapeAddOptions`:

| Field      | Type                        | Description                     |
| ---------- | --------------------------- | ------------------------------- |
| `name`     | `string`                    | Optional picture shape name.    |
| `geometry` | `Partial<PptShapeGeometry>` | Optional EMU geometry.          |
| `mimeType` | `string`                    | Optional content type override. |
| `filename` | `string`                    | Optional media filename hint.   |

When `source` is plain bytes and `mimeType` is omitted, the SDK sniffs PNG,
JPEG, and GIF signatures. Unknown bytes fall back to PNG metadata.

## Picture Shape APIs

| API                             | Returns                         | Notes                                                              |
| ------------------------------- | ------------------------------- | ------------------------------------------------------------------ |
| `shape.image.isImage()`         | `Promise<boolean>`              | True when shape is a replaceable picture.                          |
| `shape.image.getInfo()`         | `Promise<PptImageInfo \| null>` | Relationship id, target path, content type, filename, byte length. |
| `shape.image.style()`           | `Promise<PptImageStyle>`        | Picture crop, fill rectangle, and opacity snapshot.                |
| `shape.image.setStyle(style)`   | `Promise<void>`                 | Partial crop/fillRect/opacity style update.                        |
| `shape.image.setCrop(crop)`     | `Promise<void>`                 | Updates `a:srcRect`; values are ratios from `0` to `1`.            |
| `shape.image.setFillRect(rect)` | `Promise<void>`                 | Updates `a:stretch/a:fillRect`; values are ratios from `0` to `1`. |
| `shape.image.setOpacity(n)`     | `Promise<void>`                 | Updates `a:alphaModFix`; value is a ratio from `0` to `1`.         |
| `shape.image.bytes()`           | `Promise<Uint8Array>`           | Raw media part bytes.                                              |
| `shape.image.replace(source)`   | `Promise<void>`                 | Replaces bytes in the existing media part.                         |

`PptImageStyle`:

| Field      | Type                | Description                                                 |
| ---------- | ------------------- | ----------------------------------------------------------- |
| `crop`     | `PptImageCropModel` | Optional `left`, `top`, `right`, and `bottom` crop ratios.  |
| `fillRect` | `PptImageCropModel` | Optional stretch fill rectangle ratios for picture fitting. |
| `opacity`  | `number`            | Picture opacity ratio from `0` to `1`; `1` is fully opaque. |

`replace(source)` accepts the same `PptImageSource` shapes, but replacement is
in-place: it updates bytes in the existing media part and preserves the existing
target path and content type. Use `addImage(...)` when a new content type or
filename must be reflected in the package metadata.

## Global Export

```ts
const images = await presentation.exportImages();
```

`PptExportedImage` includes:

- `path`
- `contentType`
- `byteLength`
- `bytes`

This exports package media from `ppt/media/*`, not only picture shapes. Use
`shape.image.bytes()` when the task specifically needs the bytes behind a
replaceable picture shape. It does not render slides, shapes, or thumbnails.

## Replace First Existing Picture

```ts
for (const slide of await presentation.slides.items()) {
  for (const shape of await slide.shapes.items()) {
    if (await shape.image.isImage()) {
      await shape.image.replace({ bytes: replacement, mimeType: 'image/png' });
      return;
    }
  }
}
```

## SDK Operation Rules

1. Use `addImage` for new pictures; do not manually create package rels.
2. Use `image.replace` only when `isImage()` is true.
3. Use `image.setCrop(...)`, `image.setFillRect(...)`, and
   `image.setOpacity(...)` for picture style edits; do not rebuild picture XML
   manually.
4. Use `presentation.exportImages()` for bulk extraction.
5. Do not promise slide thumbnails; rendering is outside current core.
6. Use explicit `mimeType` when adding bytes whose type cannot be reliably sniffed.
7. `image.replace(...)` preserves the original media part path and content type;
   do not rely on it to convert PNG metadata to JPEG/GIF metadata.
