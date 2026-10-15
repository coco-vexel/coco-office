# Section APIs

This domain covers reading section boundaries and header/footer references.

Use this document when the task says: section, page setup, header, footer, first page header, section break, continuous section, or cover page signal.

## Skill Summary

Section APIs are read-only. They describe the body section structure. Editing sections is not exposed through live-object APIs yet.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.sections.list()` | List body sections in document order. | none | `Promise<SectionInfo[]>` |

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

## CodeAgent Rules

1. Use sections for analysis and navigation only.
2. Do not generate section write APIs; they are not part of the current surface.
3. Use `firstParaId` and `lastParaId` with paragraph APIs when subsequent paragraph operations are needed.
