/**
 * WordDocument 读 API namespace 表面测试(#92)。
 *
 * 验证 Phase B 原本平铺的 7 个读方法已迁移到 4 个 namespace 子对象,且旧平铺
 * 方法**已移除**(防止调用方回退)。
 */
import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { A_NS, R_NS, W14_NS, W_NS, WP_NS } from '../namespaces.js';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const CT_XML_WITH_PNG = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS_WITH_IMAGE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0b, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body><w:p w14:paraId="AAAAAAA1"><w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>surface</w:t></w:r></w:p></w:body>
</w:document>`;

async function makeDoc(documentXml = DOC_XML): Promise<WordDocument> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', documentXml);
  const buf = await zip.generateAsync({ type: 'uint8array' });
  return WordDocument.fromZip(buf, { docId: 'sub-apis-surface', mode: 'inline' });
}

async function makeDocWithImage(documentXml: string): Promise<WordDocument> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML_WITH_PNG);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', documentXml);
  zip.file('word/_rels/document.xml.rels', DOC_RELS_WITH_IMAGE);
  zip.file('word/media/image1.png', TINY_PNG);
  const buf = await zip.generateAsync({ type: 'uint8array' });
  return WordDocument.fromZip(buf, { docId: 'sub-apis-surface', mode: 'inline' });
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('WordDocument 读 API namespace 表面(#92)', () => {
  test('paragraphs.list / paragraphs.get 存在且为函数', async () => {
    const doc = await makeDoc();
    expect(typeof doc.paragraphs.list).toBe('function');
    expect(typeof doc.paragraphs.get).toBe('function');
    await doc.dispose();
  });

  test('bookmarks.list / bookmarks.get 存在且为函数', async () => {
    const doc = await makeDoc();
    expect(typeof doc.bookmarks.list).toBe('function');
    expect(typeof doc.bookmarks.get).toBe('function');
    await doc.dispose();
  });

  test('comments.list / comments.get 存在且为函数', async () => {
    const doc = await makeDoc();
    expect(typeof doc.comments.list).toBe('function');
    expect(typeof doc.comments.get).toBe('function');
    await doc.dispose();
  });

  test('revisions.list 存在且为函数', async () => {
    const doc = await makeDoc();
    expect(typeof doc.revisions.list).toBe('function');
    expect(typeof doc.revisions.get).toBe('function');
    expect(typeof doc.revisions.accept).toBe('function');
    expect(typeof doc.revisions.reject).toBe('function');
    await doc.dispose();
  });

  test('旧平铺方法已移除(`listParagraphs` / `getParagraph` / ...)', async () => {
    const doc = await makeDoc();
    const d = doc as unknown as Record<string, unknown>;
    expect(d.listParagraphs).toBeUndefined();
    expect(d.getParagraph).toBeUndefined();
    expect(d.listBookmarks).toBeUndefined();
    expect(d.getBookmark).toBeUndefined();
    expect(d.listComments).toBeUndefined();
    expect(d.getComment).toBeUndefined();
    expect(d.listRevisions).toBeUndefined();
    await doc.dispose();
  });

  test('已有 namespace(snapshots)未变更', async () => {
    const doc = await makeDoc();
    expect(typeof doc.snapshots.take).toBe('function');
    expect(typeof doc.snapshots.rollback).toBe('function');
    await doc.dispose();
  });

  test('tables.list / tables.get 存在且为函数（#185）', async () => {
    const doc = await makeDoc();
    expect(typeof doc.tables.list).toBe('function');
    expect(typeof doc.tables.get).toBe('function');
    await doc.dispose();
  });

  test('end-to-end:tables.list returns writable Table', async () => {
    const doc = await makeDoc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p w14:paraId="BBBBBBB1"><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p w14:paraId="BBBBBBB2"><w:r><w:t>B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`);
    const [table] = await doc.tables.list({ includeCellTexts: true });

    expect(table!.id).toBe(table!.tableId);
    expect(table!.rowCount).toBe(1);
    expect(table!.colCount).toBe(2);
    expect(await table!.getText()).toBe('A\tB');
    expect(await table!.getCells()).toEqual([['A', 'B']]);
    expect(await table!.cell(0, 1).getText()).toBe('B');
    expect(await table!.row(0).getText()).toBe('A\tB');
    expect(await table!.column(1).getText()).toBe('B');
    expect((await table!.getInfo())?.tableId).toBe(table!.tableId);
    expect((await table!.getVerbatim())?.tableXml).toContain('<w:tbl');

    await table!.cell(0, 1).replaceWith('BB');
    expect(await table!.getCells()).toEqual([['A', 'BB']]);

    await table!.row(0).insertAfter(['C', 'D']);
    expect(await table!.getCells()).toEqual([
      ['A', 'BB'],
      ['C', 'D'],
    ]);

    await table!.row(1).setHeight({ val: 360, hRule: 'atLeast' });
    expect((await table!.getInfo())?.rowHeights[1]).toEqual({ val: 360, hRule: 'atLeast' });

    await table!.column(1).setWidth(4000);
    expect((await table!.getInfo())?.columnWidths).toEqual([2000, 4000]);

    await table!.row(1).delete();
    expect(await table!.getCells()).toEqual([['A', 'BB']]);

    await doc.dispose();
  });

  test('end-to-end:table cell exposes ordered content blocks and searchable text ranges', async () => {
    const doc = await makeDocWithImage(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:p w14:paraId="CELLP001"><w:r><w:t>before</w:t></w:r></w:p>
          <w:p w14:paraId="CELLP002"><w:r>
            <w:drawing>
              <wp:inline>
                <wp:extent cx="914400" cy="685800"/>
                <wp:docPr id="1" name="image1"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic>
                      <pic:nvPicPr>
                        <pic:cNvPr id="1" name="image1"/>
                        <pic:cNvPicPr/>
                      </pic:nvPicPr>
                      <pic:blipFill>
                        <a:blip r:embed="rId1"/>
                        <a:stretch><a:fillRect/></a:stretch>
                      </pic:blipFill>
                      <pic:spPr>
                        <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="685800"/></a:xfrm>
                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                      </pic:spPr>
                    </pic:pic>
                  </a:graphicData>
                </a:graphic>
              </wp:inline>
            </w:drawing>
          </w:r></w:p>
          <w:p w14:paraId="CELLP003"><w:r><w:t>after target</w:t></w:r></w:p>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>
            <w:tr>
              <w:tc><w:p w14:paraId="NESTP001"><w:r><w:t>nested</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`);
    const [table] = await doc.tables.list({ includeCellTexts: true, includeNested: true });
    const cell = table!.cell(0, 0);
    const info = await cell.getInfo();
    expect(info?.paragraphIds).toEqual(['CELLP001', 'CELLP002', 'CELLP003']);
    expect(info?.blocks).toEqual([
      { type: 'paragraph', paraId: 'CELLP001', text: 'before' },
      { type: 'drawing', drawingId: 'dr-0001', paraId: 'CELLP002' },
      { type: 'paragraph', paraId: 'CELLP003', text: 'after target' },
      { type: 'table', tableId: 'tbl-0002' },
    ]);

    const paragraphs = await cell.getParagraphs();
    expect(await Promise.all(paragraphs.map((paragraph) => paragraph.getText()))).toEqual([
      'before',
      '',
      'after target',
    ]);

    const blocks = await cell.getBlocks();
    expect(blocks?.map((block) => block.type)).toEqual([
      'paragraph',
      'drawing',
      'paragraph',
      'table',
    ]);
    expect(blocks?.[0]?.type === 'paragraph' ? blocks[0].paragraph.id : null).toBe('CELLP001');
    expect(blocks?.[1]?.type === 'drawing' ? blocks[1].drawing?.id : null).toBe('dr-0001');
    expect(
      blocks?.[1]?.type === 'drawing' ? (await blocks[1].drawing?.getImageBytes())?.bytes : null,
    ).toEqual(TINY_PNG);
    expect(blocks?.[3]?.type === 'table' ? blocks[3].table?.tableId : null).toBe('tbl-0002');

    await cell.findText('target').setBold(true);
    const xml = await doc.toFlatOpc();
    expect(xml).toContain('<w:b/>');

    await doc.dispose();
  });

  test('drawings.list / drawings.get 存在且为函数（#189）', async () => {
    const doc = await makeDoc();
    expect(typeof doc.drawings.list).toBe('function');
    expect(typeof doc.drawings.get).toBe('function');
    await doc.dispose();
  });

  test('end-to-end: drawings.list 对无 drawing 文档返回空数组（#189）', async () => {
    const doc = await makeDoc();
    const list = await doc.drawings.list();
    expect(Array.isArray(list)).toBe(true);
    expect(list).toHaveLength(0);
    await doc.dispose();
  });

  test('tocFields.list 存在且为函数（#197）', async () => {
    const doc = await makeDoc();
    expect(typeof doc.tocFields.list).toBe('function');
    const list = await doc.tocFields.list();
    expect(Array.isArray(list)).toBe(true);
    expect(list).toHaveLength(0);
    await doc.dispose();
  });

  test('写入口(plan / toZip / toFlatOpc / dispose)保留平铺签名', async () => {
    const doc = await makeDoc();
    expect(typeof doc.plan).toBe('function');
    expect(typeof doc.toZip).toBe('function');
    expect(typeof doc.toFlatOpc).toBe('function');
    expect(typeof doc.dispose).toBe('function');
    await doc.dispose();
  });

  test('end-to-end:paragraphs.list 正常工作', async () => {
    const doc = await makeDoc();
    const r = await doc.paragraphs.list({ includeFormat: true, pagination: { mode: 'all' } });
    expect(r.paragraphs).toHaveLength(1);
    expect(r.paragraphs[0]!.paraId).toBe('AAAAAAA1');
    expect(r.paragraphs[0]!.text).toBe('surface');
    expect(r.paragraphs[0]!.format?.firstRunColor).toBe('0000FF');
    expect(await r.paragraphs[0]!.getText()).toBe('surface');
    await doc.dispose();
  });

  test('end-to-end:paragraph live object getInfo 可按需读取 format', async () => {
    const doc = await makeDoc();
    const paragraph = await doc.paragraphs.get('AAAAAAA1');
    expect((await paragraph?.getInfo())?.format).toBeUndefined();
    expect((await paragraph?.getInfo({ includeFormat: true }))?.format?.firstRunColor).toBe(
      '0000FF',
    );
    await doc.dispose();
  });

  test('end-to-end:paragraph live object getInfo 可按需读取表格坐标', async () => {
    const doc = await makeDoc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="OUTSIDE1"><w:r><w:t>outside</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr><w:tc><w:p w14:paraId="INSIDE01"><w:r><w:t>inside</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:body>
</w:document>`);
    const [table] = await doc.tables.list();
    const inside = await doc.paragraphs.get('INSIDE01');
    const outside = await doc.paragraphs.get('OUTSIDE1');

    expect((await inside?.getInfo())?.table).toBeUndefined();
    expect((await inside?.getInfo({ includeTableInfo: true }))?.table).toEqual({
      tableId: table!.tableId,
      row: 0,
      col: 0,
    });
    expect((await outside?.getInfo({ includeTableInfo: true }))?.table).toBeUndefined();

    await doc.dispose();
  });

  test('end-to-end:paragraph.findText 返回可写 TextRange', async () => {
    const doc = await makeDoc();
    const paragraph = await doc.paragraphs.get('AAAAAAA1');
    await paragraph?.findText('surface').replaceWith('range-updated');
    expect(await paragraph?.getText()).toBe('range-updated');
    await doc.dispose();
  });

  test('end-to-end:TextRange.setColor removes w14 textFill that overrides w:color', async () => {
    const doc = await makeDoc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="AAAAAAA1">
      <w:r>
        <w:rPr>
          <w14:textFill><w14:solidFill><w14:schemeClr w14:val="tx1"/></w14:solidFill></w14:textFill>
          <w:color w:val="000000"/>
        </w:rPr>
        <w:t>surface</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`);
    const paragraph = await doc.paragraphs.get('AAAAAAA1');
    await paragraph?.findText('surface').setColor('FF0000');
    const xml = await doc.toFlatOpc();
    expect(xml).toContain('w:val="FF0000"');
    expect(xml).not.toContain('textFill');
    await doc.dispose();
  });
});
