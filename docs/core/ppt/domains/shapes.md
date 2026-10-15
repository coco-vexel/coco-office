# PPT Shapes Domain

Use this document when a task mentions shapes, text boxes, rectangles, lines,
geometry, fill, outline, z-order, naming, creation, or deletion.

## Skill Summary

Shapes are live handles under a slide. Use the collection for discovery and
creation, then route by domain:

```ts
const shape = slide.shapes.getItem('Title');
await shape.setGeometry({ left: 500000, top: 300000 });
await shape.text.setText('Updated Title');
```

Use `shape.text`, `shape.image`, and `shape.table` for type-specific work.
Charts, media, OLE, and SmartArt are mostly read-only metadata in core.

## Shape Collection

```ts
const slide = presentation.slides.getItem(0);
const shapes = await slide.shapes.items();
const title = slide.shapes.getItem('Title');
```

| API                                          | Returns                  | Notes                                                             |
| -------------------------------------------- | ------------------------ | ----------------------------------------------------------------- |
| `slide.shapes.items()`                       | `Promise<Shape[]>`       | Shape-tree order; later shapes are visually above earlier shapes. |
| `slide.shapes.getCount()`                    | `Promise<number>`        | Shape count.                                                      |
| `slide.shapes.getItem(idOrName)`             | `Shape`                  | Ambiguous duplicated names throw.                                 |
| `slide.shapes.getItemAt(index)`              | `Shape`                  | 0-based shape-tree index.                                         |
| `slide.shapes.getItemOrNullObject(idOrName)` | `Promise<Shape \| null>` | Existence check.                                                  |
| `slide.shapes.deleteItem(idOrName)`          | `Promise<void>`          | Deletes matching shape.                                           |

## Shape Creation

```ts
await slide.shapes.addTextBox('Hello', {
  name: 'Intro',
  geometry: { left: 600000, top: 600000, width: 3000000, height: 700000 },
});

await slide.shapes.addShape('rect', { name: 'Box' });
await slide.shapes.addShape('line', { name: 'Divider' });
```

| API                                         | Presets / Options                                                          |
| ------------------------------------------- | -------------------------------------------------------------------------- |
| `addTextBox(text, options?)`                | `name`, `geometry`                                                         |
| `addTable(rowCount, columnCount, options?)` | See [tables.md](./tables.md).                                              |
| `addImage(source, options?)`                | See [images.md](./images.md).                                              |
| `addShape(preset, options?)`                | `preset: 'rect' \| 'roundRect' \| 'ellipse' \| 'line'`; `name`, `geometry` |

Geometry values are EMUs from OOXML. Use existing geometry as a reference when
editing a real deck.

`PptShapeAddOptions` and `PptShapeDuplicateOptions`:

| Field      | Type                        | Description                              |
| ---------- | --------------------------- | ---------------------------------------- |
| `name`     | `string`                    | Optional shape name for later retrieval. |
| `geometry` | `Partial<PptShapeGeometry>` | Optional EMU geometry partial.           |

`PptShapeGeometry`:

| Field            | Type      | Description                                 |
| ---------------- | --------- | ------------------------------------------- |
| `left`           | `number`  | X position in EMUs.                         |
| `top`            | `number`  | Y position in EMUs.                         |
| `width`          | `number`  | Width in EMUs.                              |
| `height`         | `number`  | Height in EMUs.                             |
| `rotation`       | `number`  | Optional rotation from source model.        |
| `flipHorizontal` | `boolean` | Optional horizontal flip from source model. |
| `flipVertical`   | `boolean` | Optional vertical flip from source model.   |

## Shape Metadata And Geometry

| API                           | Returns                              | Notes                                                                   |
| ----------------------------- | ------------------------------------ | ----------------------------------------------------------------------- |
| `shape.id()`                  | `Promise<string>`                    | Stable within slide.                                                    |
| `shape.name()`                | `Promise<string \| undefined>`       | Display name when present.                                              |
| `shape.setName(name)`         | `Promise<void>`                      | Updates `p:cNvPr/@name`.                                                |
| `shape.kind()`                | `Promise<PptShapeKind>`              | `textBox`, `picture`, `table`, `geometricShape`, `line`, etc.           |
| `shape.geometry()`            | `Promise<PptShapeGeometry>`          | `{ left, top, width, height, rotation, flipHorizontal, flipVertical }`. |
| `shape.hyperlink()`           | `Promise<PptHyperlinkModel \| null>` | Shape-level click hyperlink.                                            |
| `shape.setHyperlink(options)` | `Promise<void>`                      | Writes `a:hlinkClick` and slide hyperlink relationship.                 |
| `shape.clearHyperlink()`      | `Promise<void>`                      | Removes shape-level hyperlink and unused relationship.                  |
| `shape.toModel().placeholder` | `PptPlaceholderModel \| undefined`   | Placeholder type/index/size/orientation snapshot.                       |
| `shape.setGeometry(geometry)` | `Promise<void>`                      | Partial update.                                                         |
| `shape.toModel()`             | `Promise<PptShapeModel>`             | Defensive snapshot.                                                     |

## Shape Style

```ts
await shape.setFillColor('#FF0000');
await shape.setLineStyle({ lineColor: '#333333', lineWidth: 25400 });
await shape.setStyle({
  fillColor: '#FFFF00',
  opacity: 0.5,
  lineColor: '#000000',
  lineOpacity: 0.75,
  shadow: {
    color: '#333333',
    opacity: 0.35,
    blur: 12700,
    distance: 25400,
    direction: 45,
  },
});
```

| API                         | Returns                  | Notes                                                                   |
| --------------------------- | ------------------------ | ----------------------------------------------------------------------- |
| `shape.style()`             | `Promise<PptShapeStyle>` | Fill, line, opacity, gradient, cap, dash, line-end, and shadow model fields. |
| `shape.setStyle(style)`     | `Promise<void>`          | Partial update.                                                         |
| `shape.setFillColor(color)` | `Promise<void>`          | `#RRGGBB` or `RRGGBB`.                                                  |
| `shape.setLineStyle(style)` | `Promise<void>`          | `lineColor`, `lineWidth`.                                               |

The P0 write surface is solid fill color/opacity plus solid line
color/opacity/width and simple outer shadow. Additional fields such as gradient,
line cap, line dash, line ends, and complex effects are parsed into `PptShapeStyle` for
inspection and preservation. Retrieve
[../style-capabilities.md](../style-capabilities.md) before generating style
write code beyond the P0 fields.

## Shape Hyperlinks

```ts
await shape.setHyperlink({
  target: 'https://example.com',
  tooltip: 'Open source',
});

await shape.clearHyperlink();
```

`targetMode` defaults to `External`. Replacing or clearing a hyperlink removes
the old slide relationship only when no other hyperlink still references it.

`PptHyperlinkWriteOptions`:

| Field        | Type     | Description                                     |
| ------------ | -------- | ----------------------------------------------- |
| `target`     | `string` | Required URL or relationship target.            |
| `targetMode` | `string` | Optional target mode; defaults to `External`.   |
| `action`     | `string` | Optional PowerPoint hyperlink action attribute. |
| `tooltip`    | `string` | Optional tooltip.                               |

## Charts

```ts
if (await shape.chart.isChart()) {
  const info = await shape.chart.getInfo();
  const data = await shape.chart.getData();
  console.log(info?.type, data?.title, data?.series);
}

await shape.chart.setData({
  title: 'Updated Sales',
  series: [
    { name: 'North', categories: ['Q1', 'Q2'], values: [10, 12] },
    { name: 'South', categories: ['Q1', 'Q2'], values: [8, 14] },
  ],
});
```

| API                         | Returns                          | Notes                                                    |
| --------------------------- | -------------------------------- | -------------------------------------------------------- |
| `shape.chart.isChart()`     | `Promise<boolean>`               | True for chart shapes with chart metadata.               |
| `shape.chart.getInfo()`     | `Promise<PptChartModel \| null>` | Relationship, chart part, type, title, and series names. |
| `shape.chart.getData()`     | `Promise<PptChartData \| null>`  | Chart XML cache title, categories, and values.           |
| `shape.chart.setData(data)` | `Promise<void>`                  | Replaces cache data for existing series only.            |

Chart data replacement is conservative. It updates chart XML caches for the
existing series count and does not edit linked workbook parts or add/remove
series.

`PptChartData`:

| Field    | Type                   | Description                       |
| -------- | ---------------------- | --------------------------------- |
| `title`  | `string`               | Optional chart title.             |
| `series` | `PptChartSeriesData[]` | Existing series data replacement. |

`PptChartSeriesData`: `{ name?: string; categories?: string[]; values: number[] }`.

## Media, OLE, And SmartArt

```ts
const model = await shape.toModel();
if (model.kind === 'media') console.log(model.media?.mediaType, model.media?.targetPath);
if (model.kind === 'oleObject') console.log(model.oleObject?.progId);
if (model.kind === 'smartArt') console.log(model.smartArt?.targetPaths);
```

These domains are read-only shape identification APIs. Unknown advanced
`graphicFrame` payloads are exposed as `kind: 'unknown'` with `raw.hostType`
when detected. Advanced and unknown shapes preserve package relationships for
analysis and writeback planning, but they do not play media, render SmartArt, or
edit embedded OLE objects.

## Z-Order And Deletion

```ts
await shape.bringToFront();
await shape.sendToBack();
const copy = await shape.duplicate({ name: 'Copy' });
await shape.delete();
```

`items()` returns shape-tree order after reload. `bringToFront` appends the shape
to the slide shape tree; `sendToBack` inserts it before the first normal shape.
`duplicate(options?)` clones the shape on the same slide with a new shape id and
optional `name` / `geometry` overrides. Picture duplicates get their own media
part so replacing the copy image does not mutate the original image.

## Groups

```ts
const group = slide.shapes.getItem('Group 1');
if (await group.group.isGroup()) {
  for (const child of await group.group.items()) {
    if (await child.text.hasText()) {
      await child.text.setRunTextStyle(0, { bold: true });
    }
  }
}
```

| API                                         | Returns                  | Notes                         |
| ------------------------------------------- | ------------------------ | ----------------------------- |
| `shape.group.isGroup()`                     | `Promise<boolean>`       | True for `kind: 'group'`.     |
| `shape.group.items()`                       | `Promise<Shape[]>`       | Direct children only.         |
| `shape.group.getCount()`                    | `Promise<number>`        | Direct child count.           |
| `shape.group.getItem(idOrName)`             | `Shape`                  | Direct child lookup.          |
| `shape.group.getItemAt(index)`              | `Shape`                  | Direct child index.           |
| `shape.group.getItemOrNullObject(idOrName)` | `Promise<Shape \| null>` | Direct child existence check. |

`presentation.findText(...)` and `slide.findText(...)` include text in group
children. Group traversal is supported for core reads and targeted child
operations; group-level regrouping and child reordering are not part of the
current core contract.

## SDK Operation Rules

1. Use names for template-friendly code only when the template controls names.
2. Use `items()` and filter by `kind()`/`name()` when names may be duplicated.
3. Use `shape.text`, `shape.image`, or `shape.table` for type-specific work.
4. Avoid raw OOXML for geometry/style edits.
5. Colors should be `#RRGGBB` or `RRGGBB`.
6. For groups, call `shape.group.items()` and operate on child shape handles.
7. Treat `media`, `oleObject`, `smartArt`, and `unknown` as read-only metadata in core.
