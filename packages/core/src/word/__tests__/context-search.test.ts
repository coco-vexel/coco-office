import { describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { W14_NS, W_NS } from '../namespaces.js';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

async function makeDoc(documentXml: string): Promise<WordDocument> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', documentXml);
  const buf = await zip.generateAsync({ type: 'uint8array' });
  return WordDocument.fromZip(buf, { docId: 'context-search', mode: 'inline' });
}

function documentXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="HEAD0001">
      <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      <w:r><w:t>合同条款</w:t></w:r>
    </w:p>
    <w:p w14:paraId="PAY00001"><w:r><w:t>付款方式为银行转账。</w:t></w:r></w:p>
    <w:p w14:paraId="BREACH01">
      <w:bookmarkStart w:id="7" w:name="_co_append_BREACH01"/>
      <w:r><w:t>违约金按每日万分之五计算。</w:t></w:r>
      <w:bookmarkEnd w:id="7"/>
    </w:p>
    <w:p w14:paraId="REV00001">
      <w:r><w:t>违约责任说明</w:t></w:r>
      <w:ins w:id="1"><w:r><w:t>需要复核</w:t></w:r></w:ins>
    </w:p>
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p w14:paraId="CELL0001"><w:r><w:t>字段</w:t></w:r></w:p></w:tc>
        <w:tc><w:p w14:paraId="CELL0002"><w:r><w:t>保证金金额 1000 元</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;
}

describe('doc.context.search', () => {
  test('finds paragraph context with heading path, flags, and bookmark metadata', async () => {
    const doc = await makeDoc(documentXml());

    const result = await doc.context.search({
      query: '违约',
      context: { before: 1, after: 1 },
      bookmarkFilter: { prefix: '_co_append' },
    });

    expect(result.matches.length).toBeGreaterThanOrEqual(2);
    const breach = result.matches.find((match) => match.target.kind === 'paragraph');
    expect(breach).toMatchObject({
      reason: 'keyword',
      target: { kind: 'paragraph', paraId: 'BREACH01' },
      headingPath: ['合同条款'],
      bookmarks: [{ name: '_co_append_BREACH01', rangeKind: 'paragraph', writable: true }],
    });
    expect(breach?.before?.[0]).toMatchObject({ paraId: 'PAY00001' });
    expect(breach?.after?.[0]).toMatchObject({ paraId: 'REV00001' });

    const revision = result.matches.find(
      (match) => match.target.kind === 'paragraph' && match.target.paraId === 'REV00001',
    );
    expect(revision?.flags).toEqual({ hasRevisions: true });

    await doc.dispose();
  });

  test('finds table cell text with table coordinates', async () => {
    const doc = await makeDoc(documentXml());

    const result = await doc.context.search({ query: '保证金', limit: 3 });

    expect(result.matches).toHaveLength(1);
    expect(result.matches[0]).toMatchObject({
      reason: 'table',
      target: { kind: 'tableCell', paraId: 'CELL0002', row: 0, col: 1 },
      text: '保证金金额 1000 元',
    });

    await doc.dispose();
  });

  test('searches covered text for explicit hidden bookmark prefix', async () => {
    const doc = await makeDoc(documentXml());

    const result = await doc.context.search({
      query: '',
      include: { bookmarks: 'coveredText' },
      bookmarkFilter: { prefix: '_co_append' },
    });

    expect(result.matches).toHaveLength(1);
    expect(result.matches[0]).toMatchObject({
      reason: 'bookmark',
      target: {
        kind: 'bookmark',
        bookmarkName: '_co_append_BREACH01',
        rangeKind: 'paragraph',
        paraId: 'BREACH01',
      },
      text: '违约金按每日万分之五计算。',
    });

    await doc.dispose();
  });
});
