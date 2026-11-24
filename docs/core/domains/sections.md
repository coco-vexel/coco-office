# Section APIs

This domain covers section boundaries, header and footer references, and page layout.

Use this document when the task says: section, page setup, header, footer, first page header, section break, continuous section, or cover page signal.

## Skill Summary

Use the section index to update page orientation and margins for one section at a time.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.sections.list()` | List body sections in document order. | none | `Promise<SectionInfo[]>` |
| `doc.sections.setPageLayout(sectionIndex, layout)` | Set page orientation, margins, or both. | `number`, `SectionPageLayout` | `Promise<ApplyResult>` |

## SectionPageLayout

| Field | Type | Description |
| --- | --- | --- |
| `orientation` | `'portrait' \| 'landscape'` | Sets the direction and arranges the existing paper dimensions to match. |
| `pageSize` | `{ width: number; height: number }` | Sets the paper dimensions in points. A supplied orientation arranges the dimensions accordingly. |
| `verticalAlignment` | `'top' \| 'center' \| 'both' \| 'bottom'` | Sets the vertical alignment of text on the section's pages. |
| `margins` | `{ top?: number; bottom?: number; left?: number; right?: number }` | Each side is independently writable in points (pt). Fractional points are rounded to the nearest twip when written to DOCX (20 twips per point). |

Omit a field to keep its existing value. For documents with no explicit section properties, use index `0` for the implicit section. The SDK creates section properties on write and uses the same defaults as `WordDocument.createBlank()` for missing page dimensions or a missing margin element.

## SectionInfo

| Field | Type | Description |
| --- | --- | --- |
| `sectionIndex` | `number` | Zero-based section order. |
| `firstParaId` | `string` | First paragraph id in the section. |
| `lastParaId` | `string` | Last paragraph id in the section. |
| `type` | `'nextPage' \| 'continuous' \| 'evenPage' \| 'oddPage' \| 'nextColumn'?` | Section break type. Missing means OOXML default. |
| `headerRefs` | `HeaderFooterRef[]?` | Header part references. |
| `footerRefs` | `HeaderFooterRef[]?` | Footer part references. |
| `titlePage` | `boolean?` | First-page-different setting. |

`HeaderFooterRef`:

| Field | Type | Description |
| --- | --- | --- |
| `type` | `'default' \| 'first' \| 'even'` | Header/footer variant. |
| `partName` | `string` | Target OOXML part path, such as `word/header1.xml`. |

## Examples

Find section boundaries:

```ts
const sections = await doc.sections.list();
for (const section of sections) {
  console.log(section.firstParaId, section.lastParaId);
}
```

Find sections with first-page headers:

```ts
const sections = await doc.sections.list();
const coverLike = sections.filter((section) => section.titlePage);
```

Set the first section to landscape with four margins of about 10 mm (28.35 pt each).

```ts
await doc.sections.setPageLayout(0, {
  orientation: 'landscape',
  margins: { top: 28.35, bottom: 28.35, left: 28.35, right: 28.35 },
});
```

## CodeAgent Rules

1. Use `setPageLayout` with an explicit `sectionIndex`. Section indices are zero-based.
2. Use `firstParaId` and `lastParaId` with paragraph APIs when subsequent paragraph operations are needed.
