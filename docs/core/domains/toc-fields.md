# TOC Field APIs

This domain covers reading table-of-contents field ranges.

Use this document when the task says: table of contents, TOC, catalog, directory field, field instruction, or exclude TOC paragraphs.

## Skill Summary

TOC APIs are read-only. They identify TOC field ranges so other code can skip or classify paragraphs inside the table of contents.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.tocFields.list()` | List TOC field ranges in body order. | none | `Promise<TocFieldRange[]>` |

## TocFieldRange

| Field | Type | Description |
| --- | --- | --- |
| `tocIndex` | `number` | Zero-based TOC order. |
| `firstParaId` | `string` | First paragraph id covered by the TOC field. |
| `lastParaId` | `string` | Last paragraph id covered by the TOC field. |
| `instr` | `string` | Field instruction text, such as `TOC \\o "1-3" \\h \\z \\u`. |

## Examples

Skip paragraphs inside TOC ranges:

```ts
const tocFields = await doc.tocFields.list();
const paragraphs = await doc.paragraphs.list({ pagination: { mode: 'all' } });

// TOC range comparison requires paragraph order from the current list.
const indexByParaId = new Map(paragraphs.paragraphs.map((p, index) => [p.paraId, index]));
const inToc = (paraId: string) =>
  tocFields.some((toc) => {
    const start = indexByParaId.get(toc.firstParaId);
    const end = indexByParaId.get(toc.lastParaId);
    const current = indexByParaId.get(paraId);
    return start !== undefined && end !== undefined && current !== undefined && current >= start && current <= end;
  });
```

## CodeAgent Rules

1. Use TOC APIs for classification only.
2. Do not generate TOC update/delete/write APIs.
3. TOC ranges are paragraph-id ranges; if you need to compare positions, build an order map from `doc.paragraphs.list()`.
