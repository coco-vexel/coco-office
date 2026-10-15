# Drawing APIs

This domain covers drawings, drawing anchors, charts, images, and shapes.

Use this document when the task says: chart, image, picture, drawing, shape, text box, drawing anchor, chart source formula, or floating object.

## Skill Summary

Drawing APIs expose OOXML drawing metadata as async live objects. They are intended for high-fidelity read, targeted repair, package-backed metadata updates, and basic chart creation. Standalone `fromModel(...).toZip()` can synthesize charts created through `sheet.charts.add(...)`, but does not synthesize arbitrary drawing/image/shape snapshots from scratch.

## Drawings

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.drawings.items()` | List drawing parts. | none | `Promise<ExcelDrawing[]>` |
| `sheet.drawings.toModel()` | Return drawing snapshots. | none | `Promise<ExcelDrawingModel[]>` |

`ExcelDrawing`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `drawing.path()` | Read drawing part path. | none | `Promise<string \| undefined>` |
| `drawing.relationshipId()` | Read worksheet relationship id. | none | `Promise<string \| undefined>` |
| `drawing.anchors` | Drawing anchor collection. | property | `ExcelDrawingAnchorCollection` |
| `drawing.toModel()` | Return drawing snapshot. | none | `Promise<ExcelDrawingModel>` |

`ExcelDrawingAnchorCollection`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `drawing.anchors.items()` | List anchors. | none | `Promise<ExcelDrawingAnchor[]>` |
| `drawing.anchors.getItem(index)` | Get anchor by zero-based index. | `number` | `ExcelDrawingAnchor` |
| `drawing.anchors.getItemOrNullObject(index)` | Get anchor by zero-based index, or null. | `number` | `Promise<ExcelDrawingAnchor \| null>` |
| `drawing.anchors.toModel()` | Return anchor snapshots. | none | `Promise<ExcelDrawingAnchorModel[]>` |

`ExcelDrawingAnchor`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `anchor.kind()` | Read anchor kind. | none | `Promise<string>` |
| `anchor.from` | Start marker live object. | property | `ExcelDrawingMarker` |
| `anchor.to` | End marker live object. | property | `ExcelDrawingMarker` |
| `anchor.object()` | Read object metadata. | none | `Promise<ExcelDrawingObjectModel \| undefined>` |
| `anchor.toModel()` | Return anchor snapshot. | none | `Promise<ExcelDrawingAnchorModel>` |

`ExcelDrawingMarker`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `marker.row()` | Read zero-based row marker. | none | `Promise<number \| undefined>` |
| `marker.setRow(row)` | Set zero-based row marker. | `number \| undefined` | `Promise<void>` |
| `marker.column()` | Read zero-based column marker. | none | `Promise<number \| undefined>` |
| `marker.setColumn(column)` | Set zero-based column marker. | `number \| undefined` | `Promise<void>` |
| `marker.rowOffset()` | Read row offset. | none | `Promise<number \| undefined>` |
| `marker.setRowOffset(offset)` | Set row offset. | `number \| undefined` | `Promise<void>` |
| `marker.columnOffset()` | Read column offset. | none | `Promise<number \| undefined>` |
| `marker.setColumnOffset(offset)` | Set column offset. | `number \| undefined` | `Promise<void>` |
| `marker.toModel()` | Return marker snapshot. | none | `Promise<ExcelDrawingMarkerModel>` |

## Charts

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.charts.items()` | List chart parts. | none | `Promise<ExcelChart[]>` |
| `sheet.charts.add(options)` | Create a basic chart and anchor it on the worksheet. | `ExcelChartAddOptions` | `Promise<ExcelChart>` |
| `sheet.charts.toModel()` | Return chart snapshots. | none | `Promise<ExcelChartModel[]>` |

`ExcelChartAddOptions`:

| Field | Purpose |
| --- | --- |
| `type` | One of `columnClustered`, `columnStacked`, `barClustered`, `barStacked`, `line`, `lineMarkers`, `pie`, `doughnut`. |
| `title` | Optional chart title. |
| `categoryRange` | A1 reference for category labels, such as `Summary!A2:A5`. |
| `series` | Data series. Each item has `name?: string` and `values: string`. Pie and doughnut charts require exactly one series. |
| `position` | Two-cell anchor with zero-based `from` and `to` markers. |

### Chart Creation Recipe

Use this pattern for generated chart operations:

1. Create or reuse a worksheet for chart source data, usually a summary sheet.
2. Write a compact table with category labels in one column and numeric series in adjacent columns.
3. Use absolute A1 references with the sheet name, such as `Summary!$A$2:$A$13`.
4. Call `sheet.charts.add(...)` once per chart.
5. Place charts with non-overlapping two-cell anchors. Marker `row` and `column` are zero-based.

Column or bar chart:

```ts
await sheet.charts.add({
  type: 'columnClustered',
  title: 'Monthly Contract And Collection Amount',
  categoryRange: 'Summary!$A$2:$A$13',
  series: [
    { name: 'Contract Amount', values: 'Summary!$B$2:$B$13' },
    { name: 'Collection Amount', values: 'Summary!$C$2:$C$13' },
  ],
  position: {
    from: { row: 1, column: 4 },
    to: { row: 18, column: 12 },
  },
});
```

Pie chart:

```ts
await sheet.charts.add({
  type: 'pie',
  title: 'Contract Amount By Sales Region',
  categoryRange: 'Summary!$E$2:$E$8',
  series: [{ name: 'Contract Amount', values: 'Summary!$F$2:$F$8' }],
  position: {
    from: { row: 20, column: 4 },
    to: { row: 37, column: 12 },
  },
});
```

Do not use pie or doughnut charts with more than one series. If the task asks for multiple numeric measures by category, use `columnClustered`, `barClustered`, or `line`.

`ExcelChart`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `chart.relationshipId()` | Read drawing relationship id. | none | `Promise<string \| undefined>` |
| `chart.drawingPath()` | Read owning drawing path. | none | `Promise<string \| undefined>` |
| `chart.path()` | Read chart part path. | none | `Promise<string \| undefined>` |
| `chart.title()` | Read title text. | none | `Promise<string \| undefined>` |
| `chart.setTitle(title)` | Set or clear title text. | `string \| undefined` | `Promise<void>` |
| `chart.formulaRefs` | Chart formula reference collection. | property | `ExcelFormulaRefCollection` |
| `chart.toModel()` | Return chart snapshot. | none | `Promise<ExcelChartModel>` |

`ExcelFormulaRefCollection`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `chart.formulaRefs.items()` | List chart formula refs. | none | `Promise<string[]>` |
| `chart.formulaRefs.getItem(index)` | Read one formula ref by zero-based index. | `number` | `Promise<string \| undefined>` |
| `chart.formulaRefs.setItem(index, formulaRef)` | Replace one formula ref. | `number`, `string` | `Promise<void>` |
| `chart.formulaRefs.add(formulaRef)` | Append a formula ref. | `string` | `Promise<void>` |
| `chart.formulaRefs.delete(index)` | Delete one formula ref. | `number` | `Promise<void>` |
| `chart.formulaRefs.toModel()` | Return formula-ref snapshots. | none | `Promise<string[]>` |

## Images

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.images.items()` | List images as live handles. | none | `Promise<ExcelImage[]>` |
| `sheet.images.toModel()` | Return image snapshots. | none | `Promise<ExcelImageModel[]>` |

`ExcelImage`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `image.relationshipId()` | Read drawing relationship id. | none | `Promise<string \| undefined>` |
| `image.drawingPath()` | Read owning drawing path. | none | `Promise<string \| undefined>` |
| `image.path()` | Read media path. | none | `Promise<string \| undefined>` |
| `image.targetMode()` | Read target mode. | none | `Promise<string \| undefined>` |
| `image.name()` | Read image name. | none | `Promise<string \| undefined>` |
| `image.toModel()` | Return image snapshot. | none | `Promise<ExcelImageModel>` |

## Shapes

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.shapes.items()` | List drawing shape objects. | none | `Promise<ExcelShape[]>` |
| `sheet.shapes.getItem(name)` | Get shape handle by name. | `string` | `ExcelShape` |
| `sheet.shapes.toModel()` | Return shape snapshots. | none | `Promise<ExcelShapeModel[]>` |

`ExcelShape`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `shape.drawingPath()` | Read owning drawing path. | none | `Promise<string \| undefined>` |
| `shape.name()` | Read shape name. | none | `Promise<string \| undefined>` |
| `shape.setName(name)` | Set or clear shape name. | `string \| undefined` | `Promise<void>` |
| `shape.toModel()` | Return shape snapshot. | none | `Promise<ExcelShapeModel>` |

## CodeAgent Rules

1. Drawing markers are 0-based, unlike worksheet row/column APIs.
2. Use `sheet.charts` for chart metadata and formula refs; use `sheet.drawings` for anchors.
3. Use `sheet.shapes` for editable shape names and anchor lookup.
4. Use `sheet.charts.add(...)` for basic chart creation. Do not hand-edit `toModel().charts`.
5. Before creating a chart from raw records, first compute and write a summary table. Point the chart at the summary table, not at unaggregated source rows.
6. For host writeback, floating-object fidelity may require host-specific copy/repair outside core.
