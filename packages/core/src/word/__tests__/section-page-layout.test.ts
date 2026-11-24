import JSZip from 'jszip';
import { parseXmlDocument } from 'slimdom';
import { expect, test } from 'vitest';
import { W_NS, WordDocument } from '../index.js';

async function loadBody(body: string): Promise<WordDocument> {
  const flatOpc = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
      <pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document></pkg:xmlData>
    </pkg:part>
  </pkg:package>`;
  return WordDocument.fromFlatOpc(flatOpc, { docId: 'page-layout', mode: 'inline' });
}

async function readDocumentXml(doc: WordDocument) {
  const zip = await JSZip.loadAsync(await doc.toZip());
  const xml = await zip.file('word/document.xml')!.async('string');
  return parseXmlDocument(xml);
}

test('sets orientation and all four margins on one section and preserves the paper size', async () => {
  const doc = await loadBody(`
    <w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>
    <w:p><w:r><w:t>Keep this text</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="1440" w:header="400" w:footer="500" w:gutter="100"/>
    </w:sectPr>`);
  const sections = await doc.sections.list();
  await doc.sections.setPageLayout(sections[1]!.sectionIndex, {
    orientation: 'landscape',
    margins: { top: (10 * 72) / 25.4, bottom: 24, left: 30, right: 36 },
  });
  await doc.sections.setPageLayout(1, { orientation: 'landscape', margins: { left: 0 } });

  const xml = await readDocumentXml(doc);
  const [first, last] = Array.from(xml.getElementsByTagNameNS(W_NS, 'pgSz'));
  expect(first!.getAttributeNS(W_NS, 'w')).toBe('11906');
  expect(first!.getAttributeNS(W_NS, 'h')).toBe('16838');
  expect(first!.hasAttributeNS(W_NS, 'orient')).toBe(false);
  expect(last!.getAttributeNS(W_NS, 'w')).toBe('16838');
  expect(last!.getAttributeNS(W_NS, 'h')).toBe('11906');
  expect(last!.getAttributeNS(W_NS, 'orient')).toBe('landscape');
  const [margins] = Array.from(xml.getElementsByTagNameNS(W_NS, 'pgMar'));
  expect(
    ['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'].map((name) =>
      margins!.getAttributeNS(W_NS, name),
    ),
  ).toEqual(['567', '480', '0', '720', '400', '500', '100']);
  expect(xml.getElementsByTagNameNS(W_NS, 't')[0]!.textContent).toBe('Keep this text');

  await doc.sections.setPageLayout(1, { orientation: 'portrait' });
  const portrait = (await readDocumentXml(doc)).getElementsByTagNameNS(W_NS, 'pgSz')[1]!;
  expect(portrait.getAttributeNS(W_NS, 'w')).toBe('11906');
  expect(portrait.getAttributeNS(W_NS, 'h')).toBe('16838');
  expect(portrait.getAttributeNS(W_NS, 'orient')).toBe('portrait');
  await doc.dispose();
});

test('writes missing section properties and inserts page settings before later section elements', async () => {
  const doc = await loadBody(`<w:p><w:pPr><w:sectPr><w:titlePg/></w:sectPr></w:pPr></w:p><w:p/>`);
  await doc.sections.setPageLayout(0, { orientation: 'landscape', margins: { bottom: 27 } });
  await doc.sections.setPageLayout(1, { margins: { top: 18 } });

  const xml = await readDocumentXml(doc);
  const [first, last] = Array.from(xml.getElementsByTagNameNS(W_NS, 'sectPr'));
  expect(Array.from(first!.children, (child) => child.localName)).toEqual([
    'pgSz',
    'pgMar',
    'titlePg',
  ]);
  expect(last!.parentElement?.localName).toBe('body');
  expect(last!.firstElementChild!.getAttributeNS(W_NS, 'top')).toBe('360');
  expect(last!.firstElementChild!.getAttributeNS(W_NS, 'right')).toBe('1440');
  await doc.dispose();
});

test('rolls back page layout changes when a later operation fails', async () => {
  const doc = await WordDocument.createBlank({ docId: 'page-layout-rollback', mode: 'inline' });
  const before = await doc.toFlatOpc();
  const plan = doc
    .plan()
    .setSectionPageLayout(0, { orientation: 'landscape', margins: { top: 18 } })
    .setSectionPageLayout(0, { margins: { right: -0.01 } });
  await expect(plan.apply()).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  expect(await doc.toFlatOpc()).toBe(before);
  await expect(doc.sections.setPageLayout(1, { orientation: 'landscape' })).rejects.toMatchObject({
    code: 'INVALID_ARGUMENT',
  });
  await doc.dispose();
});
