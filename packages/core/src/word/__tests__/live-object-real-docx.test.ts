import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { A_NS, PIC_NS, R_NS, W14_NS, W_NS, WP_NS, WPS_NS } from '../namespaces.js';

const DOC_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const IMG_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CT}"/>
</Types>`;

const CT_XML_WITH_JPG = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CT}"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const EMPTY_DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

const DOC_RELS_WITH_JPG = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="${IMG_REL}" Target="media/original.jpg"/>
</Relationships>`;

const ORIGINAL_JPG_BYTES = new Uint8Array([
  0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46,
]);

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

async function makeDocx(
  bodyInner: string,
  opts: {
    contentTypes?: string;
    documentRels?: string;
    media?: Record<string, Uint8Array>;
  } = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', opts.contentTypes ?? CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  if (opts.documentRels) {
    zip.file('word/_rels/document.xml.rels', opts.documentRels);
  }
  for (const [path, bytes] of Object.entries(opts.media ?? {})) {
    zip.file(path, bytes);
  }
  return zip.generateAsync({ type: 'uint8array' });
}

async function reload(doc: WordDocument, docId: string): Promise<WordDocument> {
  const out = await doc.toZip();
  await doc.dispose();
  return WordDocument.fromZip(out, { docId, mode: 'inline' });
}

async function readPart(doc: WordDocument, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await doc.toZip());
  return zip.file(path)!.async('string');
}

function tableBody(): string {
  return `<w:tbl>
    <w:tblPr/>
    <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr>
      <w:tc><w:p w14:paraId="TBL00001"><w:r><w:t>A</w:t></w:r></w:p></w:tc>
      <w:tc><w:p w14:paraId="TBL00002"><w:r><w:t>B</w:t></w:r></w:p></w:tc>
    </w:tr>
    <w:tr>
      <w:tc><w:p w14:paraId="TBL00003"><w:r><w:t>C</w:t></w:r></w:p></w:tc>
      <w:tc><w:p w14:paraId="TBL00004"><w:r><w:t>D</w:t></w:r></w:p></w:tc>
    </w:tr>
  </w:tbl>`;
}

function pictureDrawingXml(paraId: string): string {
  return `<w:p w14:paraId="${paraId}"><w:r><w:drawing>
    <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:pic="${PIC_NS}" distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="3000000" cy="2000000"/>
      <wp:docPr id="1" name="Picture 1"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="0" name="original.jpg"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;
}

function textBoxDrawingXml(): string {
  return `<w:p w14:paraId="TXT00001"><w:r><w:drawing>
    <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:wps="${WPS_NS}" distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="2000000" cy="1000000"/>
      <wp:docPr id="1" name="TextBox 1"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic>
        <a:graphicData uri="${WPS_NS}">
          <wps:wsp>
            <wps:cNvSpPr/>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="FFFFCC"/></a:solidFill>
              <a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p w14:paraId="TXT00002"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>old text</w:t></w:r></w:p>
                <w:p w14:paraId="TXT00003"><w:r><w:t>second line</w:t></w:r></w:p>
              </w:txbxContent>
            </wps:txbx>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;
}

function plainShapeDrawingXml(): string {
  return `<w:p w14:paraId="SHP00001"><w:r><w:drawing>
    <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:wps="${WPS_NS}" distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="1000000" cy="500000"/>
      <wp:docPr id="2" name="Shape 1"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic>
        <a:graphicData uri="${WPS_NS}">
          <wps:wsp>
            <wps:cNvSpPr/>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
            </wps:spPr>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('live object APIs on real docx packages', () => {
  test('table, cell, row and column writes persist after reload', async () => {
    let doc = await WordDocument.fromZip(await makeDocx(tableBody()), {
      docId: 'live-table-real',
      mode: 'inline',
    });

    let table = (await doc.tables.list({ includeCellTexts: true }))[0]!;
    await table.cell(0, 1).replaceWith('BB');
    await table.row(0).insertBefore(['P', 'Q']);
    table = (await doc.tables.list({ includeCellTexts: true }))[0]!;
    await table.row(2).insertAfter(['E', 'F']);
    table = (await doc.tables.list({ includeCellTexts: true }))[0]!;
    await table.row(1).setHeight({ val: 360, hRule: 'atLeast' });
    await table.column(1).setWidth(4200);
    await table.row(2).delete();

    doc = await reload(doc, 'live-table-real-reload');
    table = (await doc.tables.list({ includeCellTexts: true }))[0]!;
    expect(await table.getCells()).toEqual([
      ['P', 'Q'],
      ['A', 'BB'],
      ['E', 'F'],
    ]);
    expect((await table.getInfo())?.rowHeights[1]).toEqual({ val: 360, hRule: 'atLeast' });
    expect((await table.getInfo())?.columnWidths).toEqual([2000, 4200]);
    expect((await table.getVerbatim())?.tableXml).toContain('<w:tbl');
    await doc.dispose();
  });

  test('table comments and delete persist after reload', async () => {
    let doc = await WordDocument.fromZip(await makeDocx(tableBody()), {
      docId: 'live-table-comment',
      mode: 'inline',
    });

    let table = (await doc.tables.list({ includeCellTexts: true }))[0]!;
    await table.addComment('table level comment');

    doc = await reload(doc, 'live-table-comment-reload');
    expect((await doc.comments.list()).map((comment) => comment.text)).toContain(
      'table level comment',
    );

    table = (await doc.tables.list({ includeCellTexts: true }))[0]!;
    await table.delete({ onAnchorConflict: 'drop' });
    doc = await reload(doc, 'live-table-delete-reload');
    expect(await doc.tables.list()).toEqual([]);
    expect(await readPart(doc, 'word/document.xml')).not.toContain('<w:tbl');
    await doc.dispose();
  });

  test('paragraph TextRange operations persist after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(
        `<w:p w14:paraId="PARA0001"><w:r><w:t>alpha beta gamma beta</w:t></w:r></w:p>`,
      ),
      { docId: 'live-text-range', mode: 'inline' },
    );

    const paragraph = (await doc.paragraphs.get('PARA0001'))!;
    await paragraph.findText('alpha').replaceWith('ALPHA');
    await paragraph.findText('gamma').delete();
    await paragraph.findText('beta').setBold(true);
    await paragraph.findText('beta').setColor('FF0000');
    await paragraph.findText('beta').setItalic(true);
    await paragraph.findText('beta').setUnderline(true);
    await paragraph.findText('beta').setHighlight('yellow');
    await paragraph.findText('beta').setFontSize(14);
    await paragraph.findText('beta').setFontFamily('Arial');
    await paragraph.findText('beta').addComment('beta comment');
    await paragraph.findText('beta').createBookmark('betaBm');

    doc = await reload(doc, 'live-text-range-reload');
    expect(await (await doc.paragraphs.get('PARA0001'))!.getText()).toBe('ALPHA beta  beta');
    expect((await doc.comments.list()).map((comment) => comment.text)).toContain('beta comment');
    expect((await doc.bookmarks.get('betaBm'))?.rangeKind).toBe('text');

    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain('w:b');
    expect(xml).toContain('w:color w:val="FF0000"');
    expect(xml).toContain('w:i');
    expect(xml).toContain('w:u w:val="single"');
    expect(xml).toContain('w:highlight w:val="yellow"');
    expect(xml).toContain('w:sz w:val="28"');
    expect(xml).toContain('w:rFonts');
    expect(xml).toContain('w:ascii="Arial"');
    expect(xml).toContain('w:eastAsia="Arial"');
    expect(xml).toContain('w:bookmarkStart');
    expect(xml).toContain('w:name="betaBm"');
    await doc.dispose();
  });

  test('paragraph insertBefore and insertAfter persist order after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(`<w:p w14:paraId="INS00001"><w:r><w:t>middle</w:t></w:r></w:p>`),
      { docId: 'live-paragraph-insert', mode: 'inline' },
    );

    const paragraph = (await doc.paragraphs.get('INS00001'))!;
    const beforeResult = await paragraph.insertBefore('before');
    const afterResult = await paragraph.insertAfter('after');
    expect(beforeResult.insertedParagraphs).toHaveLength(1);
    expect(afterResult.insertedParagraphs).toHaveLength(1);
    expect(beforeResult.insertedParagraphs[0]!.position).toBe('before');
    expect(afterResult.insertedParagraphs[0]!.position).toBe('after');

    doc = await reload(doc, 'live-paragraph-insert-reload');
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await Promise.all(list.paragraphs.map((item) => item.getText()))).toEqual([
      'before',
      'middle',
      'after',
    ]);

    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain(beforeResult.insertedParagraphs[0]!.newParaId);
    expect(xml).toContain(afterResult.insertedParagraphs[0]!.newParaId);
    await doc.dispose();
  });

  test('paragraph direct format operations persist after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(`<w:p w14:paraId="FMT00001"><w:r><w:t>format me</w:t></w:r></w:p>`),
      { docId: 'live-paragraph-format', mode: 'inline' },
    );

    const paragraph = (await doc.paragraphs.get('FMT00001'))!;
    await paragraph.setAlignment('center');
    await paragraph.setOutlineLevel(1);
    await paragraph.setPageBreakBefore(true);
    await paragraph.setIndent({ left: 720, right: 120, firstLine: 360 });
    await paragraph.setSpacing({ before: 120, after: 240, line: 360, lineRule: 'exact' });

    doc = await reload(doc, 'live-paragraph-format-reload');
    const info = await (await doc.paragraphs.get('FMT00001'))!.getInfo({ includeFormat: true });
    expect(info?.format?.alignment).toBe('center');
    expect(info?.format?.outlineLevel).toBe(1);
    expect(info?.format?.pageBreakBefore).toBe(true);
    expect(info?.format?.indent).toEqual({ left: 720, right: 120, firstLine: 360 });
    expect(info?.format?.spacing).toEqual({
      before: 120,
      after: 240,
      line: 360,
      lineRule: 'exact',
    });

    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain('w:jc w:val="center"');
    expect(xml).toContain('w:outlineLvl w:val="1"');
    expect(xml).toContain('w:pageBreakBefore');
    expect(xml).toContain('w:ind');
    expect(xml).toContain('w:left="720"');
    expect(xml).toContain('w:right="120"');
    expect(xml).toContain('w:firstLine="360"');
    expect(xml).toContain('w:spacing');
    expect(xml).toContain('w:before="120"');
    expect(xml).toContain('w:after="240"');
    expect(xml).toContain('w:line="360"');
    expect(xml).toContain('w:lineRule="exact"');

    await (await doc.paragraphs.get('FMT00001'))!.setPageBreakBefore(false);
    doc = await reload(doc, 'live-paragraph-format-clear-reload');
    const cleared = await (await doc.paragraphs.get('FMT00001'))!.getInfo({ includeFormat: true });
    expect(cleared?.format?.pageBreakBefore).toBeUndefined();
    await doc.dispose();
  });

  test('comment live objects read body text, anchor text, paragraph and delete after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(`<w:p w14:paraId="C0AA0001"><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>`),
      { docId: 'live-comment', mode: 'inline' },
    );

    const paragraph = (await doc.paragraphs.get('C0AA0001'))!;
    await paragraph.findText('beta').addComment('comment body', {
      author: 'Reviewer',
      date: new Date('2026-07-09T09:10:11.000Z'),
    });

    doc = await reload(doc, 'live-comment-reload');
    let comments = await doc.comments.list();
    expect(comments).toHaveLength(1);
    const comment = comments[0]!;
    expect(comment.id).toBe(0);
    expect(comment.author).toBe('Reviewer');
    expect(comment.date).toContain('2026-07-09T09:10:11');
    expect(comment.rangeKind).toBe('text');
    expect(comment.anchorText).toBe('beta');
    expect(await comment.getText()).toBe('comment body');
    expect(await comment.getAnchorText()).toBe('beta');
    expect(await (await comment.getParagraph())?.getText()).toBe('alpha beta gamma');

    await comment.delete();
    doc = await reload(doc, 'live-comment-delete-reload');
    comments = await doc.comments.list();
    expect(comments).toHaveLength(0);
    expect(await doc.comments.get(0)).toBeNull();

    const zip = await JSZip.loadAsync(await doc.toZip());
    expect(zip.file('word/comments.xml')).toBeNull();
    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).not.toContain('commentRangeStart');
    expect(xml).not.toContain('commentReference');
    await doc.dispose();
  });

  test('paragraph.replaceWith can preserve inline picture drawings after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(
        `<w:p w14:paraId="PIC00001">
          <w:r><w:t>before</w:t></w:r>
          <w:r><w:drawing>
            <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:pic="${PIC_NS}" distT="0" distB="0" distL="0" distR="0">
              <wp:extent cx="3000000" cy="2000000"/>
              <wp:docPr id="1" name="Picture 1"/>
              <wp:cNvGraphicFramePr/>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic>
                    <pic:nvPicPr><pic:cNvPr id="0" name="original.jpg"/><pic:cNvPicPr/></pic:nvPicPr>
                    <pic:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing></w:r>
          <w:r><w:t>after</w:t></w:r>
        </w:p>`,
        {
          contentTypes: CT_XML_WITH_JPG,
          documentRels: DOC_RELS_WITH_JPG,
          media: { 'word/media/original.jpg': ORIGINAL_JPG_BYTES },
        },
      ),
      { docId: 'live-paragraph-replace-keep-inline-picture', mode: 'inline' },
    );

    const paragraph = (await doc.paragraphs.get('PIC00001'))!;
    await paragraph.replaceWith('replacement', { keepInlineGraphics: true });

    doc = await reload(doc, 'live-paragraph-replace-keep-inline-picture-reload');
    expect(await (await doc.paragraphs.get('PIC00001'))?.getText()).toBe('replacement');
    expect(await doc.drawings.listPictures()).toHaveLength(1);
    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain('<w:drawing>');
    expect(xml).toContain('replacement');
    expect(xml).not.toContain('before');
    expect(xml).not.toContain('after');
    await doc.dispose();
  });

  test('revision live objects read and accept or reject after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(`<w:p w14:paraId="D0AA0001">
        <w:r><w:t>before </w:t></w:r>
        <w:ins w:id="7" w:author="alice" w:date="2026-07-09T09:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
        <w:del w:id="8" w:author="bob" w:date="2026-07-09T10:00:00Z">
          <w:r><w:delText>deleted</w:delText></w:r>
        </w:del>
      </w:p>`),
      { docId: 'live-revision', mode: 'inline' },
    );

    doc = await reload(doc, 'live-revision-reload');
    const revisions = await doc.revisions.list();
    expect(revisions.map((revision) => revision.revId).sort()).toEqual(['7', '8']);

    const insert = (await doc.revisions.get('7'))!;
    expect(insert.id).toBe('7');
    expect(insert.kind).toBe('insert');
    expect(await insert.getText()).toBe('inserted');
    expect((await insert.getInfo())?.author).toBe('alice');
    expect((await insert.getParagraph())?.paraId).toBe('D0AA0001');

    let deletion = (await doc.revisions.get('8'))!;
    expect(deletion.kind).toBe('delete');
    expect(await deletion.getText()).toBe('deleted');
    expect(await doc.revisions.get('missing')).toBeNull();

    const acceptResult = await insert.accept();
    expect(acceptResult.acceptedRevisions).toBe(1);
    doc = await reload(doc, 'live-revision-accept-reload');
    expect(await doc.revisions.get('7')).toBeNull();
    expect((await doc.revisions.list()).map((revision) => revision.revId)).toEqual(['8']);
    expect(await (await doc.paragraphs.get('D0AA0001'))?.getText()).toBe('before inserted');
    let xml = await readPart(doc, 'word/document.xml');
    expect(xml).not.toContain('w:id="7"');
    expect(xml).toContain('w:id="8"');

    deletion = (await doc.revisions.get('8'))!;
    const rejectResult = await deletion.reject();
    expect(rejectResult.rejectedRevisions).toBe(1);
    doc = await reload(doc, 'live-revision-reject-reload');
    expect(await doc.revisions.list()).toHaveLength(0);
    expect(await (await doc.paragraphs.get('D0AA0001'))?.getText()).toBe('before inserteddeleted');
    xml = await readPart(doc, 'word/document.xml');
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('<w:del');
    expect(xml).toContain('inserted');
    expect(xml).toContain('deleted');
    await doc.dispose();
  });

  test('revision namespace accepts and rejects all revisions after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(`<w:p w14:paraId="D0AA0002">
        <w:r><w:t>before </w:t></w:r>
        <w:ins w:id="17" w:author="alice" w:date="2026-07-09T09:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
        <w:del w:id="18" w:author="bob" w:date="2026-07-09T10:00:00Z">
          <w:r><w:delText>deleted</w:delText></w:r>
        </w:del>
      </w:p>`),
      { docId: 'live-revisions-bulk-reject', mode: 'inline' },
    );

    const rejectResult = await doc.revisions.reject();
    expect(rejectResult.rejectedRevisions).toBe(2);
    doc = await reload(doc, 'live-revisions-bulk-reject-reload');
    expect(await doc.revisions.list()).toHaveLength(0);
    expect(await (await doc.paragraphs.get('D0AA0002'))?.getText()).toBe('before deleted');
    await doc.dispose();

    doc = await WordDocument.fromZip(
      await makeDocx(`<w:p w14:paraId="D0AA0003">
        <w:r><w:t>before </w:t></w:r>
        <w:ins w:id="19" w:author="alice" w:date="2026-07-09T09:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
        <w:del w:id="20" w:author="bob" w:date="2026-07-09T10:00:00Z">
          <w:r><w:delText>deleted</w:delText></w:r>
        </w:del>
      </w:p>`),
      { docId: 'live-revisions-bulk-accept', mode: 'inline' },
    );

    const acceptResult = await doc.revisions.accept();
    expect(acceptResult.acceptedRevisions).toBe(2);
    doc = await reload(doc, 'live-revisions-bulk-accept-reload');
    expect(await doc.revisions.list()).toHaveLength(0);
    expect(await (await doc.paragraphs.get('D0AA0003'))?.getText()).toBe('before inserted');
    await doc.dispose();
  });

  test('bookmark live objects read text, resolve paragraphs and delete after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(`<w:p w14:paraId="ABCD0001">
        <w:bookmarkStart w:id="1" w:name="paraBm"/>
        <w:r><w:t>paragraph bookmark</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>
      <w:p w14:paraId="ABCD0002">
        <w:r><w:t>before </w:t></w:r>
        <w:bookmarkStart w:id="2" w:name="textBm"/>
        <w:r><w:t>target text</w:t></w:r>
        <w:bookmarkEnd w:id="2"/>
        <w:r><w:t> after</w:t></w:r>
      </w:p>
      <w:p w14:paraId="ABCD0003">
        <w:bookmarkStart w:id="3" w:name="_HiddenBm"/>
        <w:r><w:t>hidden bookmark</w:t></w:r>
        <w:bookmarkEnd w:id="3"/>
      </w:p>`),
      { docId: 'live-bookmark', mode: 'inline' },
    );

    expect((await doc.bookmarks.list()).map((bookmark) => bookmark.name)).toEqual([
      'paraBm',
      'textBm',
    ]);
    expect(
      (await doc.bookmarks.list({ includeHidden: true })).map((bookmark) => bookmark.name),
    ).toEqual(['paraBm', 'textBm', '_HiddenBm']);

    const paraBookmark = (await doc.bookmarks.get('paraBm'))!;
    expect(paraBookmark.id).toBe('paraBm');
    expect(paraBookmark.rangeKind).toBe('paragraph');
    expect(await paraBookmark.getText()).toBe('paragraph bookmark');
    expect((await paraBookmark.getParagraph())?.paraId).toBe('ABCD0001');

    const textBookmark = (await doc.bookmarks.get('textBm'))!;
    expect(textBookmark.rangeKind).toBe('text');
    expect(await textBookmark.getText()).toBe('target text');
    expect(await (await textBookmark.getParagraph())?.getText()).toBe('before target text after');

    await textBookmark.delete();
    doc = await reload(doc, 'live-bookmark-delete-reload');
    expect(await doc.bookmarks.get('textBm')).toBeNull();
    expect((await doc.bookmarks.get('paraBm'))?.rangeKind).toBe('paragraph');
    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).not.toContain('w:name="textBm"');
    expect(xml).toContain('w:name="paraBm"');
    await doc.dispose();
  });

  test('cell TextRange operations persist after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(`<w:tbl>
        <w:tblPr/>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc><w:p w14:paraId="CELL0001"><w:r><w:t>cell alpha beta gamma</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`),
      { docId: 'live-cell-range', mode: 'inline' },
    );

    let cell = (await doc.tables.list({ includeCellTexts: true }))[0]!.cell(0, 0);
    await cell.findText('alpha').replaceWith('ALPHA');
    await cell.findText('beta').addComment('cell beta comment');
    await cell.findText('beta').setHighlight('cyan');
    await cell.findText('beta').setUnderline(true);
    await cell.findText('gamma').createBookmark('cellGammaBm');
    await cell.findText('ALPHA').delete();

    doc = await reload(doc, 'live-cell-range-reload');
    cell = (await doc.tables.list({ includeCellTexts: true }))[0]!.cell(0, 0);
    expect(await cell.getText()).toBe('cell  beta gamma');
    expect((await doc.comments.list()).map((comment) => comment.text)).toContain(
      'cell beta comment',
    );
    expect((await doc.bookmarks.get('cellGammaBm'))?.rangeKind).toBe('text');
    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain('w:highlight w:val="cyan"');
    expect(xml).toContain('w:u w:val="single"');
    await doc.dispose();
  });

  test('cell direct format operations persist after reload', async () => {
    let doc = await WordDocument.fromZip(await makeDocx(tableBody()), {
      docId: 'live-cell-format',
      mode: 'inline',
    });

    let cell = (await doc.tables.list({ includeCellTexts: true }))[0]!.cell(0, 0);
    await cell.setShading('FFF2CC');
    await cell.setVerticalAlignment('center');
    await cell.setWidth(2600);

    doc = await reload(doc, 'live-cell-format-reload');
    cell = (await doc.tables.list({ includeCellTexts: true }))[0]!.cell(0, 0);
    const info = await cell.getInfo();
    expect(info?.shading).toEqual({ fill: 'FFF2CC', val: 'clear' });
    expect(info?.verticalAlignment).toBe('center');
    expect(info?.width).toBe(2600);

    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain('w:tcW w:w="2600" w:type="dxa"');
    expect(xml).toContain('w:shd w:val="clear" w:fill="FFF2CC"');
    expect(xml).toContain('w:vAlign w:val="center"');
    await doc.dispose();
  });

  test('cell paragraph insertion preserves direct block order after reload', async () => {
    let doc = await WordDocument.fromZip(await makeDocx(tableBody()), {
      docId: 'live-cell-paragraph-insert',
      mode: 'inline',
    });

    let cell = (await doc.tables.list({ includeCellTexts: true }))[0]!.cell(0, 0);
    const startResult = await cell.insertParagraphAtStart('cell start');
    const endResult = await cell.insertParagraphAtEnd('cell end');
    expect(startResult.insertedParagraphs).toHaveLength(1);
    expect(endResult.insertedParagraphs).toHaveLength(1);

    doc = await reload(doc, 'live-cell-paragraph-insert-reload');
    cell = (await doc.tables.list({ includeCellTexts: true }))[0]!.cell(0, 0);
    expect(await cell.getText()).toBe('cell start\nA\ncell end');

    const blocks = await cell.getBlocks();
    expect(blocks?.map((block) => block.type)).toEqual(['paragraph', 'paragraph', 'paragraph']);
    expect(blocks?.map((block) => (block.type === 'paragraph' ? block.text : null))).toEqual([
      'cell start',
      'A',
      'cell end',
    ]);

    const paragraphs = await cell.getParagraphs();
    expect(await Promise.all(paragraphs.map((paragraph) => paragraph.getText()))).toEqual([
      'cell start',
      'A',
      'cell end',
    ]);

    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain(startResult.insertedParagraphs[0]!.newParaId);
    expect(xml).toContain(endResult.insertedParagraphs[0]!.newParaId);
    await doc.dispose();
  });

  test('cell paragraph handles support paragraph writes after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(`<w:tbl>
        <w:tblPr/>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:p w14:paraId="CPH00001"><w:r><w:t>first</w:t></w:r></w:p>
            <w:p w14:paraId="CPH00002"><w:r><w:t>middle</w:t></w:r></w:p>
            <w:p w14:paraId="CPH00003"><w:r><w:t>last</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>`),
      { docId: 'live-cell-paragraph-handles', mode: 'inline' },
    );

    let cell = (await doc.tables.list({ includeCellTexts: true }))[0]!.cell(0, 0);
    let paragraphs = await cell.getParagraphs();
    await paragraphs[1]!.insertBefore('before middle');
    await paragraphs[1]!.insertAfter('after middle');
    await paragraphs[0]!.setAlignment('center');
    await paragraphs[2]!.delete();

    doc = await reload(doc, 'live-cell-paragraph-handles-reload');
    cell = (await doc.tables.list({ includeCellTexts: true }))[0]!.cell(0, 0);
    expect(await cell.getText()).toBe('first\nbefore middle\nmiddle\nafter middle');

    paragraphs = await cell.getParagraphs();
    expect(await Promise.all(paragraphs.map((paragraph) => paragraph.getText()))).toEqual([
      'first',
      'before middle',
      'middle',
      'after middle',
    ]);

    const firstInfo = await paragraphs[0]!.getInfo({ includeFormat: true });
    expect(firstInfo?.format?.alignment).toBe('center');
    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain('before middle');
    expect(xml).toContain('after middle');
    expect(xml).not.toContain('>last<');
    expect(xml).toContain('w:jc w:val="center"');
    await doc.dispose();
  });

  test('cell blocks expose writable drawing and nested table handles after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(
        `<w:tbl>
          <w:tblPr/>
          <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
          <w:tr>
            <w:tc>
              <w:p w14:paraId="CBL00001"><w:r><w:t>before block</w:t></w:r></w:p>
              ${pictureDrawingXml('CBL00002')}
              <w:p w14:paraId="CBL00003"><w:r><w:t>between block</w:t></w:r></w:p>
              <w:tbl>
                <w:tblPr/>
                <w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>
                <w:tr>
                  <w:tc><w:p w14:paraId="CBLN0001"><w:r><w:t>nested cell</w:t></w:r></w:p></w:tc>
                </w:tr>
              </w:tbl>
            </w:tc>
          </w:tr>
        </w:tbl>`,
        {
          contentTypes: CT_XML_WITH_JPG,
          documentRels: DOC_RELS_WITH_JPG,
          media: { 'word/media/original.jpg': ORIGINAL_JPG_BYTES },
        },
      ),
      { docId: 'live-cell-block-handles', mode: 'inline' },
    );

    let outerTable = (await doc.tables.list({ includeCellTexts: true, includeNested: true }))[0]!;
    let cell = outerTable.cell(0, 0);
    let blocks = await cell.getBlocks();
    expect(blocks?.map((block) => block.type)).toEqual([
      'paragraph',
      'drawing',
      'paragraph',
      'table',
    ]);

    const drawingBlock = blocks?.find((block) => block.type === 'drawing');
    expect(drawingBlock?.type === 'drawing' ? drawingBlock.drawing?.kind : null).toBe('picture');
    expect(
      drawingBlock?.type === 'drawing'
        ? (await drawingBlock.drawing?.getImageBytes())?.bytes
        : null,
    ).toEqual(ORIGINAL_JPG_BYTES);
    await (drawingBlock?.type === 'drawing'
      ? drawingBlock.drawing?.replaceWithPlaceholder()
      : null);

    const nestedTableBlock = blocks?.find((block) => block.type === 'table');
    expect(
      nestedTableBlock?.type === 'table' ? await nestedTableBlock.table?.getText() : null,
    ).toBe('nested cell');
    await (nestedTableBlock?.type === 'table'
      ? nestedTableBlock.table?.cell(0, 0).replaceWith('nested updated')
      : null);

    doc = await reload(doc, 'live-cell-block-handles-reload');
    outerTable = (await doc.tables.list({ includeCellTexts: true, includeNested: true }))[0]!;
    cell = outerTable.cell(0, 0);
    blocks = await cell.getBlocks();
    expect(blocks?.map((block) => block.type)).toEqual([
      'paragraph',
      'drawing',
      'paragraph',
      'table',
    ]);

    const reloadedDrawing = blocks?.find((block) => block.type === 'drawing');
    expect(reloadedDrawing?.type === 'drawing' ? reloadedDrawing.drawing?.kind : null).toBe(
      'picture',
    );
    expect(
      reloadedDrawing?.type === 'drawing'
        ? (await reloadedDrawing.drawing?.getImageBytes())?.contentType
        : null,
    ).toBe('image/png');

    const reloadedNestedTable = blocks?.find((block) => block.type === 'table');
    expect(
      reloadedNestedTable?.type === 'table' ? await reloadedNestedTable.table?.getText() : null,
    ).toBe('nested updated');

    const zip = await JSZip.loadAsync(await doc.toZip());
    expect(zip.file('word/media/original.jpg')).toBeNull();
    expect(zip.file('word/media/placeholder.png')).not.toBeNull();
    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain('nested updated');
    expect(xml).not.toContain('nested cell');
    await doc.dispose();
  });

  test('drawing image bytes, placeholder replacement and delete persist after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(pictureDrawingXml('DRW00001'), {
        contentTypes: CT_XML_WITH_JPG,
        documentRels: DOC_RELS_WITH_JPG,
        media: { 'word/media/original.jpg': ORIGINAL_JPG_BYTES },
      }),
      { docId: 'live-drawing-picture', mode: 'inline' },
    );

    let picture = (await doc.drawings.listPictures())[0]!;
    expect(await doc.drawings.listTextBoxes()).toHaveLength(0);
    expect(await doc.drawings.listShapes()).toHaveLength(0);
    const drawing = (await doc.drawings.list())[0]!;
    expect(picture.kind).toBe('picture');
    expect((await picture.getImageBytes())?.bytes).toEqual(ORIGINAL_JPG_BYTES);
    await expect(drawing.replaceShapeInnerText('not for pictures')).rejects.toThrow(
      'replaceDrawingWithPlaceholder',
    );
    await picture.replaceWithPlaceholder();

    doc = await reload(doc, 'live-drawing-placeholder-reload');
    let zip = await JSZip.loadAsync(await doc.toZip());
    expect(zip.file('word/media/original.jpg')).toBeNull();
    expect(zip.file('word/media/placeholder.png')).not.toBeNull();
    expect(await doc.drawings.list()).toHaveLength(1);

    picture = (await doc.drawings.listPictures())[0]!;
    await picture.delete();
    doc = await reload(doc, 'live-drawing-delete-reload');
    expect(await doc.drawings.list()).toHaveLength(0);
    zip = await JSZip.loadAsync(await doc.toZip());
    expect(await zip.file('word/document.xml')!.async('string')).not.toContain('<w:drawing');
    await doc.dispose();
  });

  test('drawing text box replacement persists after reload', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(textBoxDrawingXml(), { documentRels: EMPTY_DOC_RELS }),
      { docId: 'live-drawing-textbox', mode: 'inline' },
    );

    let drawing = (await doc.drawings.list())[0]!;
    let textBox = (await doc.drawings.listTextBoxes())[0]!;
    expect(drawing.kind).toBe('textBox');
    expect(await doc.drawings.listPictures()).toHaveLength(0);
    expect(await doc.drawings.listTextBoxes()).toHaveLength(1);
    expect(await doc.drawings.listShapes()).toHaveLength(0);
    expect(await textBox.getText()).toBe('old textsecond line');
    let paragraphs = await textBox.getParagraphs();
    await expect(Promise.all(paragraphs.map((paragraph) => paragraph.getText()))).resolves.toEqual([
      'old text',
      'second line',
    ]);
    expect((await textBox.getBlocks()).map((block) => block.type)).toEqual([
      'paragraph',
      'paragraph',
    ]);
    expect((await drawing.getImageBytes())?.bytes).toBeUndefined();
    await expect(drawing.replaceWithPlaceholder()).rejects.toThrow("only supports kind='picture'");
    await textBox.replaceText('shape placeholder');
    expect(await textBox.getText()).toBe('shape placeholder');
    paragraphs = await textBox.getParagraphs();
    await expect(Promise.all(paragraphs.map((paragraph) => paragraph.getText()))).resolves.toEqual([
      'shape placeholder',
    ]);
    expect((await textBox.getBlocks()).map((block) => block.type)).toEqual(['paragraph']);

    doc = await reload(doc, 'live-drawing-textbox-reload');
    drawing = (await doc.drawings.list())[0]!;
    textBox = (await doc.drawings.listTextBoxes())[0]!;
    expect(drawing.kind).toBe('textBox');
    expect(await textBox.getText()).toBe('shape placeholder');
    await expect(
      Promise.all((await textBox.getParagraphs()).map((paragraph) => paragraph.getText())),
    ).resolves.toEqual(['shape placeholder']);
    expect((await textBox.getBlocks()).map((block) => block.type)).toEqual(['paragraph']);
    const xml = await readPart(doc, 'word/document.xml');
    expect(xml).toContain('shape placeholder');
    expect(xml).not.toContain('old text');
    expect(xml).not.toContain('second line');
    expect(xml).toContain('FFFFCC');
    expect(xml).toContain('w="6350"');
    await doc.dispose();
  });

  test('drawing typed list helpers return only matching kinds', async () => {
    let doc = await WordDocument.fromZip(
      await makeDocx(
        `${pictureDrawingXml('DTP00001')}${textBoxDrawingXml()}${plainShapeDrawingXml()}`,
        {
          contentTypes: CT_XML_WITH_JPG,
          documentRels: DOC_RELS_WITH_JPG,
          media: { 'word/media/original.jpg': ORIGINAL_JPG_BYTES },
        },
      ),
      { docId: 'live-drawing-typed-lists', mode: 'inline' },
    );

    const pictures = await doc.drawings.listPictures();
    const textBoxes = await doc.drawings.listTextBoxes();
    const shapes = await doc.drawings.listShapes();

    expect(pictures.map((drawing) => drawing.kind)).toEqual(['picture']);
    expect(textBoxes.map((drawing) => drawing.kind)).toEqual(['textBox']);
    expect(shapes.map((drawing) => drawing.kind)).toEqual(['shape']);
    expect((await pictures[0]!.getImageBytes())?.bytes).toEqual(ORIGINAL_JPG_BYTES);
    expect((await textBoxes[0]!.getInfo())?.innerText).toBe('old textsecond line');
    expect(await textBoxes[0]!.getText()).toBe('old textsecond line');
    await expect(
      Promise.all((await textBoxes[0]!.getParagraphs()).map((paragraph) => paragraph.getText())),
    ).resolves.toEqual(['old text', 'second line']);
    expect((await textBoxes[0]!.getBlocks()).map((block) => block.type)).toEqual([
      'paragraph',
      'paragraph',
    ]);
    expect((await shapes[0]!.getInfo())?.hasFill).toBe(true);
    await shapes[0]!.delete();
    doc = await reload(doc, 'live-drawing-shape-delete-reload');
    expect(await doc.drawings.listShapes()).toHaveLength(0);
    expect((await doc.drawings.list()).map((drawing) => drawing.kind)).toEqual([
      'picture',
      'textBox',
    ]);
    await doc.dispose();
  });
});
