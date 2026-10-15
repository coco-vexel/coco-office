import { describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { fromFlatOpc, toFlatOpc } from '@coco-vexel/coco-office/pack';
import { fromZip, toZip } from '@coco-vexel/coco-office/pack';
import { buildIndices, W14_NS, W_NS } from '@coco-vexel/coco-office/word-view';
import {
  applyNativeWordChangesToPackage,
  createNativeWordDocument,
  nativeDeleteBackward,
  nativeDeleteImage,
  nativeInsertText,
  nativeInsertImage,
  nativeAddComment,
  nativeMergeTableCells,
  nativeSplitTableCell,
  nativeSplitParagraph,
  nativeSetHyperlink,
  nativeSetImageSize,
  nativeSyncParagraphText,
  type NativeWordChange,
} from '../native.js';
import {
  buildWordEditorViewFromDocx,
  buildWordEditorViewFromFlatOpc,
  buildWordEditorViewFromPackage,
} from '../extract.js';

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

async function makeHeaderFooterDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', HEADER_FOOTER_CT);
  zip.file('_rels/.rels', `<Relationships xmlns="${PACKAGE_RELS_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', HEADER_FOOTER_DOC_RELS);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?>${HEADER_FOOTER_DOCUMENT}`);
  zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8"?>${HEADER_XML}`);
  zip.file('word/footer1.xml', `<?xml version="1.0" encoding="UTF-8"?>${HEADER_FOOTER_XML}`);
  zip.file('word/comments.xml', `<?xml version="1.0" encoding="UTF-8"?>${HEADER_FOOTER_COMMENTS}`);
  return zip.generateAsync({ type: 'uint8array' });
}

const FLAT_OPC = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p>
          <w:p w14:paraId="P0000002"><w:r><w:t>World</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

const HEADER_FOOTER_CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`;

const HEADER_FOOTER_DOC_RELS = `<Relationships xmlns="${PACKAGE_RELS_NS}">
  <Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;

const HEADER_FOOTER_DOCUMENT = `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="${R_NS}">
  <w:body>
    <w:p w14:paraId="HF-BODY"><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1200" w:right="900" w:bottom="1000" w:left="900"/></w:sectPr>
  </w:body>
</w:document>`;

const HEADER_FOOTER_XML = `<w:ftr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:p w14:paraId="HF-FOOTER"><w:r><w:rPr><w:i/></w:rPr><w:t>Footer </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
</w:ftr>`;

const HEADER_XML = `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:p w14:paraId="HF-HEADER"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Header</w:t></w:r></w:p>
</w:hdr>`;

const HEADER_FOOTER_COMMENTS = `<w:comments xmlns:w="${W_NS}"/>`;

const HYPERLINK_FLAT_OPC = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="${R_NS}">
        <w:body><w:p w14:paraId="LINK0001"><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r><w:r><w:t> world</w:t></w:r></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/></pkg:xmlData>
  </pkg:part>
</pkg:package>`;

const TABLE_FLAT_OPC = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
          <w:tr><w:tc><w:tcPr/><w:p w14:paraId="T0000001"><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p w14:paraId="T0000002"><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
          <w:tr><w:tc><w:tcPr/><w:p w14:paraId="T0000003"><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p w14:paraId="T0000004"><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
      </w:body></w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

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

async function makeDocxWithPicture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML_WITH_PNG);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/_rels/document.xml.rels', DOC_RELS_WITH_IMAGE);
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="${R_NS}">
  <w:body>
    <w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:p w14:paraId="P0000002">
      <w:r>
        <w:drawing>
          <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="Image 1"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
      <w:r><w:t>PictureEnd</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`,
  );
  zip.file('word/media/image1.png', TINY_PNG);
  return zip.generateAsync({ type: 'uint8array' });
}

describe('@coco-vexel/coco-office-editor native Word save roundtrip', () => {
  test('builds an editor view directly from DOCX bytes', async () => {
    const view = await buildWordEditorViewFromDocx(await makeDocxWithPicture());
    expect(view.paragraphs[0]?.text).toBe('Hello');
    expect(view.paragraphs.some((paragraph) => paragraph.segments.some((segment) => segment.kind === 'image'))).toBe(true);
  });

  test('builds an editor view directly from Flat OPC', () => {
    const view = buildWordEditorViewFromFlatOpc(FLAT_OPC);
    expect(view.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Hello', 'World']);
  });

  test('persists native text insert changes through Flat OPC reload', () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const doc = createNativeWordDocument(view);
    const change = nativeInsertText(doc, 'P0000001', 5, '!');

    const result = applyNativeWordChangesToPackage(pkg, change ? [change] : []);

    expect(result).toMatchObject({
      plan: {
        unsupported: [],
        patches: [{ kind: 'paragraph-text', op: 'insert', paraId: 'P0000001' }],
      },
      applied: [{ kind: 'paragraph-text', paraId: 'P0000001' }],
      skipped: [],
      unresolved: [],
    });

    const reloaded = fromFlatOpc(toFlatOpc(pkg));
    const reloadedView = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded));
    expect(reloadedView.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Hello!', 'World']);
  });

  test('persists native browser-sync replace and delete changes through reload', () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const doc = createNativeWordDocument(view);
    const replace = nativeSyncParagraphText(doc, 'P0000002', 'Word');
    const deletion = nativeDeleteBackward(doc, 'P0000002', 4);

    const changes = [replace, deletion].filter((change): change is NativeWordChange => change !== null);
    const result = applyNativeWordChangesToPackage(pkg, changes);

    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);
    expect(result.applied).toMatchObject([
      { kind: 'paragraph-text', op: 'delete', paraId: 'P0000002', offset: 3, text: 'l' },
      { kind: 'paragraph-text', op: 'delete', paraId: 'P0000002', offset: 3, text: 'd' },
    ]);

    const reloaded = fromFlatOpc(toFlatOpc(pkg));
    const reloadedView = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded));
    expect(reloadedView.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Hello', 'Wor']);
  });

  test('persists combined text, split, and image deletion changes through docx reload', async () => {
    const pkg = await fromZip(await makeDocxWithPicture());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const doc = createNativeWordDocument(view);
    const image = view.paragraphs[1]!.segments.find((segment) => segment.kind === 'image');
    expect(image).toMatchObject({ kind: 'image' });

    const changes = [
      nativeInsertText(doc, 'P0000001', 5, '!'),
      nativeSplitParagraph(doc, 'P0000001', 3, 'P00000AA'),
      image?.kind === 'image' ? nativeDeleteImage(doc, 'P0000002', image.drawingId) : null,
    ].filter((change): change is NativeWordChange => change !== null);

    const result = applyNativeWordChangesToPackage(pkg, changes);

    expect(result).toMatchObject({
      plan: {
        unsupported: [],
        patches: [
          { kind: 'paragraph-text', op: 'insert', paraId: 'P0000001', offset: 5, text: '!' },
          { kind: 'paragraph-split', paraId: 'P0000001', newParaId: 'P00000AA', offset: 3 },
          { kind: 'drawing-delete', paraId: 'P0000002' },
        ],
      },
      skipped: [],
      unresolved: [],
    });

    const reloaded = await fromZip(await toZip(pkg));
    const reloadedView = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded));
    expect(reloadedView.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Hel', 'lo!', 'PictureEnd']);
    expect(reloadedView.paragraphs.flatMap((paragraph) => paragraph.segments).filter((segment) => segment.kind === 'image')).toEqual([]);
  });

  test('accepts a tracked image deletion during DOCX writeback', async () => {
    const pkg = await fromZip(await makeDocxWithPicture());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const image = view.paragraphs[1]!.segments.find((segment) => segment.kind === 'image');
    expect(image).toMatchObject({ kind: 'image' });
    if (image?.kind !== 'image') return;
    const change: NativeWordChange = {
      kind: 'image', op: 'accept', paraId: 'P0000002', offset: 0, drawingId: image.drawingId,
      revisionId: 'image-delete-1', revisionKind: 'delete',
    };
    const result = applyNativeWordChangesToPackage(pkg, [change]);
    expect(result.skipped).toEqual([]);
    const reloaded = await fromZip(await toZip(pkg));
    expect(buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded)).paragraphs
      .flatMap((paragraph) => paragraph.segments).some((segment) => segment.kind === 'image')).toBe(false);
  });

  test('writes a text replacement as delete plus insert revisions', () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const deletion: NativeWordChange = {
      kind: 'revision', op: 'delete', paraId: 'P0000001', from: 0, to: 5, revisionId: '71', revisionKind: 'delete',
    };
    const insertText: NativeWordChange = {
      kind: 'text', op: 'insert', paraId: 'P0000001', offset: 5, text: 'New',
    };
    const insertion: NativeWordChange = {
      kind: 'revision', op: 'insert', paraId: 'P0000001', from: 5, to: 8, revisionId: '72', revisionKind: 'insert',
    };
    expect(applyNativeWordChangesToPackage(pkg, [deletion, insertText, insertion]).skipped).toEqual([]);
    const body = pkg.parts.get('word/document.xml')!.xmlDom!;
    expect(body.getElementsByTagNameNS(W_NS, 'delText')[0]!.textContent).toBe('Hello');
    expect(body.getElementsByTagNameNS(W_NS, 'ins')[0]!.textContent).toBe('New');

    const reject = [deletion, insertion].map((change) => ({ ...change, op: 'reject' as const }));
    expect(applyNativeWordChangesToPackage(pkg, reject).skipped).toEqual([]);
    expect(buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).paragraphs[0]!.text).toBe('Hello');
  });

  test('rejects a tracked image insertion after the image is written', async () => {
    const pkg = await fromZip(await makeDocxWithPicture());
    const src = `data:image/png;base64,${Buffer.from(TINY_PNG).toString('base64')}`;
    const inserted = nativeInsertImage(createNativeWordDocument(buildWordEditorViewFromPackage(pkg, buildIndices(pkg))), 'P0000001', 5, src, {
      width: 32,
      height: 32,
      drawingId: 'tracked-image',
    });
    expect(inserted).toMatchObject({ kind: 'image', op: 'insert' });
    if (!inserted || inserted.kind !== 'image') return;
    const revision: NativeWordChange = {
      kind: 'image', op: 'reject', paraId: 'P0000001', offset: 5, drawingId: inserted.drawingId,
      revisionId: 'image-insert-1', revisionKind: 'insert',
    };
    expect(applyNativeWordChangesToPackage(pkg, [inserted, revision]).skipped).toEqual([]);
    const reloaded = await fromZip(await toZip(pkg));
    expect(buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded)).paragraphs[0]!.segments
      .some((segment) => segment.kind === 'image')).toBe(false);
  });

  test('writes common run formatting changes to DOCX XML', () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const change: NativeWordChange = {
      kind: 'format',
      op: 'run',
      paraId: 'P0000001',
      from: 0,
      to: 5,
      format: { bold: true, underline: true, color: '#c00000', fontSize: 14, fontFamily: 'SimSun' },
    };
    expect(applyNativeWordChangesToPackage(pkg, [change]).skipped).toEqual([]);
    const body = pkg.parts.get('word/document.xml')!.xmlDom!;
    const run = body.getElementsByTagNameNS(W_NS, 'r')[0]!;
    const rPr = run.getElementsByTagNameNS(W_NS, 'rPr')[0]!;
    expect(rPr.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(1);
    expect(rPr.getElementsByTagNameNS(W_NS, 'u')[0]!.getAttribute('val')).toBe('single');
    expect(rPr.getElementsByTagNameNS(W_NS, 'color')[0]!.getAttribute('val')).toBe('C00000');
    expect(rPr.getElementsByTagNameNS(W_NS, 'sz')[0]!.getAttribute('val')).toBe('28');
    expect(rPr.getElementsByTagNameNS(W_NS, 'rFonts')[0]!.getAttribute('eastAsia')).toBe('SimSun');
  });

  test('writes footer paragraph text and formatting to its own DOCX story part', async () => {
    const pkg = await fromZip(await makeHeaderFooterDocx());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const footerParagraph = view.footers[0]?.paragraphs[0];
    const headerParagraph = view.headers[0]?.paragraphs[0];
    expect(footerParagraph).toMatchObject({ paraId: 'HF-FOOTER', text: 'Footer \uFFFC' });
    expect(headerParagraph).toMatchObject({ paraId: 'HF-HEADER', text: 'Header' });
    if (!footerParagraph || !headerParagraph) return;
    const doc = createNativeWordDocument(view);
    const insert = nativeInsertText(doc, footerParagraph.paraId, footerParagraph.text.length, '!');
    const format: NativeWordChange = {
      kind: 'format', op: 'run', paraId: headerParagraph.paraId, from: 0, to: headerParagraph.text.length,
      format: { bold: true, color: '#C00000' },
    };
    expect(applyNativeWordChangesToPackage(pkg, insert ? [insert, format] : [format]).skipped).toEqual([]);
    const reloaded = await fromZip(await toZip(pkg));
    const reloadedView = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded));
    expect(reloadedView.headers[0]?.paragraphs[0]?.text).toBe('Header');
    expect(reloadedView.footers[0]?.paragraphs[0]?.text).toBe('Footer \uFFFC!');
    const headerXml = reloaded.parts.get('word/header1.xml')!.xmlDom!;
    const footerXml = reloaded.parts.get('word/footer1.xml')!.xmlDom!;
    const bodyXml = reloaded.parts.get('word/document.xml')!.xmlDom!;
    expect(headerXml.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(1);
    expect(headerXml.getElementsByTagNameNS(W_NS, 'color')[0]!.getAttribute('val')).toBe('C00000');
    expect(footerXml.getElementsByTagNameNS(W_NS, 'fldChar')).toHaveLength(3);
    expect(footerXml.getElementsByTagNameNS(W_NS, 'instrText')).toHaveLength(1);
    expect(bodyXml.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(0);
    expect(bodyXml.documentElement?.textContent ?? '').toContain('Body');
  });

  test('writes comments and resolves revisions inside header story parts', async () => {
    const commentPkg = await fromZip(await makeHeaderFooterDocx());
    const commentView = buildWordEditorViewFromPackage(commentPkg, buildIndices(commentPkg));
    const commentDoc = createNativeWordDocument(commentView);
    const comment = nativeAddComment(commentDoc, 'HF-HEADER', 0, 6, {
      commentId: 41,
      author: 'Tester',
      date: '2026-08-24T00:00:00.000Z',
      body: 'Check header',
    });
    expect(comment).toMatchObject({ kind: 'annotation', op: 'insert', commentId: 41 });
    expect(applyNativeWordChangesToPackage(commentPkg, comment ? [comment] : []).skipped).toEqual([]);
    const reloadedCommentPkg = await fromZip(await toZip(commentPkg));
    const reloadedCommentView = buildWordEditorViewFromPackage(reloadedCommentPkg, buildIndices(reloadedCommentPkg));
    expect(reloadedCommentView.headers[0]?.paragraphs[0]?.spans).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: 'comment', commentId: 41, body: 'Check header' }),
    ]));
    expect(reloadedCommentPkg.parts.get('word/comments.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'comment')).toHaveLength(1);
    expect(reloadedCommentPkg.parts.get('word/header1.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'commentRangeStart')).toHaveLength(1);

    const revisionPkg = await fromZip(await makeHeaderFooterDocx());
    const revision: NativeWordChange = {
      kind: 'revision', op: 'insert', paraId: 'HF-HEADER', from: 0, to: 6,
      revisionId: 'header-revision-1', revisionKind: 'insert',
    };
    expect(applyNativeWordChangesToPackage(revisionPkg, [revision]).skipped).toEqual([]);
    expect(applyNativeWordChangesToPackage(revisionPkg, [{ ...revision, op: 'accept' }]).skipped).toEqual([]);
    const headerXml = revisionPkg.parts.get('word/header1.xml')!.xmlDom!;
    expect(headerXml.getElementsByTagNameNS(W_NS, 'ins')).toHaveLength(0);
    expect(headerXml.documentElement?.textContent ?? '').toContain('Header');
  });

  test('writes and resizes an image inside a header story part', async () => {
    const pkg = await fromZip(await makeHeaderFooterDocx());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const header = view.headers[0]?.paragraphs[0];
    if (!header) return;
    const doc = createNativeWordDocument(view);
    const src = `data:image/png;base64,${Buffer.from(TINY_PNG).toString('base64')}`;
    const insert = nativeInsertImage(doc, header.paraId, header.text.length, src, {
      width: 32,
      height: 32,
      drawingId: 'header-image-1',
    });
    expect(insert).toMatchObject({ kind: 'image', op: 'insert' });
    if (!insert || insert.kind !== 'image') return;
    expect(applyNativeWordChangesToPackage(pkg, [insert]).skipped).toEqual([]);
    const insertedView = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const image = insertedView.headers[0]?.paragraphs[0]?.segments.find((node) => node.kind === 'image');
    expect(image).toMatchObject({ kind: 'image' });
    if (!image || image.kind !== 'image') return;
    const resize: NativeWordChange = {
      kind: 'image', op: 'resize', paraId: header.paraId, offset: header.text.length,
      drawingId: image.drawingId, width: 64, height: 24,
    };
    expect(applyNativeWordChangesToPackage(pkg, [resize]).skipped).toEqual([]);
    const reloaded = await fromZip(await toZip(pkg));
    const reloadedImage = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded)).headers[0]?.paragraphs[0]?.segments
      .find((node) => node.kind === 'image');
    expect(reloadedImage).toMatchObject({ kind: 'image', width: 64, height: 24 });
    expect(reloaded.parts.has('word/media/image1.png')).toBe(true);
    expect(reloaded.parts.has('word/_rels/header1.xml.rels')).toBe(true);
    expect(reloaded.parts.get('[Content_Types].xml')!.xmlDom!.getElementsByTagNameNS(
      'http://schemas.openxmlformats.org/package/2006/content-types', 'Default',
    )).toHaveLength(3);
  });

  test('writes and rejects a run format revision with its before-format snapshot', () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const format: NativeWordChange = {
      kind: 'format', op: 'run', paraId: 'P0000001', from: 0, to: 5, format: { bold: true },
    };
    const revision: NativeWordChange = {
      kind: 'format', op: 'run', paraId: 'P0000001', from: 0, to: 5, format: {},
      revisionId: '91', revisionKind: 'format', beforeFormat: {},
    };
    expect(applyNativeWordChangesToPackage(pkg, [format, revision]).skipped).toEqual([]);
    const body = pkg.parts.get('word/document.xml')!.xmlDom!;
    expect(body.getElementsByTagNameNS(W_NS, 'rPrChange')).toHaveLength(1);
    const reject: NativeWordChange = { ...revision, op: 'reject' };
    expect(applyNativeWordChangesToPackage(pkg, [reject]).skipped).toEqual([]);
    expect(body.getElementsByTagNameNS(W_NS, 'rPrChange')).toHaveLength(0);
    expect(body.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(0);
  });

  test('formats and resolves a format revision nested inside an insertion', () => {
    const trackedFlatOpc = FLAT_OPC.replace(
      '<w:r><w:t>Hello</w:t></w:r>',
      '<w:r><w:t>Hello</w:t></w:r><w:ins w:id="101" w:author="User"><w:r><w:t>!</w:t></w:r></w:ins>',
    );
    const rejectPkg = fromFlatOpc(trackedFlatOpc);
    const format: NativeWordChange = {
      kind: 'format', op: 'run', paraId: 'P0000001', from: 5, to: 6, format: { bold: true },
    };
    const revision: NativeWordChange = {
      kind: 'format', op: 'run', paraId: 'P0000001', from: 5, to: 6, format: {},
      revisionId: '102', revisionKind: 'format', beforeFormat: {},
    };
    expect(applyNativeWordChangesToPackage(rejectPkg, [format, revision]).skipped).toEqual([]);
    const reject = applyNativeWordChangesToPackage(rejectPkg, [{ ...revision, op: 'reject' }]);
    expect(reject.skipped).toEqual([]);
    const rejectBody = rejectPkg.parts.get('word/document.xml')!.xmlDom!;
    expect(rejectBody.getElementsByTagNameNS(W_NS, 'rPrChange')).toHaveLength(0);
    expect(rejectBody.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(0);

    const acceptPkg = fromFlatOpc(trackedFlatOpc);
    expect(applyNativeWordChangesToPackage(acceptPkg, [format, revision]).skipped).toEqual([]);
    expect(applyNativeWordChangesToPackage(acceptPkg, [{ ...revision, op: 'accept' }]).skipped).toEqual([]);
    const acceptBody = acceptPkg.parts.get('word/document.xml')!.xmlDom!;
    expect(acceptBody.getElementsByTagNameNS(W_NS, 'rPrChange')).toHaveLength(0);
    expect(acceptBody.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(1);
  });

  test('formats and resolves a format revision nested inside a deletion', () => {
    const trackedFlatOpc = FLAT_OPC.replace(
      '<w:r><w:t>Hello</w:t></w:r>',
      '<w:del w:id="103" w:author="User"><w:r><w:rPr><w:i/></w:rPr><w:delText>Hello</w:delText></w:r></w:del>',
    );
    const format: NativeWordChange = {
      kind: 'format', op: 'run', paraId: 'P0000001', from: 0, to: 5, format: { bold: true },
    };
    const revision: NativeWordChange = {
      kind: 'format', op: 'run', paraId: 'P0000001', from: 0, to: 5, format: {},
      revisionId: '104', revisionKind: 'format', beforeFormat: { italic: true },
    };
    const rejectPkg = fromFlatOpc(trackedFlatOpc);
    expect(applyNativeWordChangesToPackage(rejectPkg, [format, revision]).skipped).toEqual([]);
    expect(applyNativeWordChangesToPackage(rejectPkg, [{ ...revision, op: 'reject' }]).skipped).toEqual([]);
    const rejectRun = rejectPkg.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'r')[0]!;
    expect(rejectRun.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(0);
    expect(rejectRun.getElementsByTagNameNS(W_NS, 'i')).toHaveLength(1);

    const acceptPkg = fromFlatOpc(trackedFlatOpc);
    expect(applyNativeWordChangesToPackage(acceptPkg, [format, revision]).skipped).toEqual([]);
    expect(applyNativeWordChangesToPackage(acceptPkg, [{ ...revision, op: 'accept' }]).skipped).toEqual([]);
    const acceptRun = acceptPkg.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'r')[0]!;
    expect(acceptRun.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(1);
    expect(acceptRun.getElementsByTagNameNS(W_NS, 'i')).toHaveLength(1);
  });

  test('persists text inserted after an inline image at the logical caret position', async () => {
    const pkg = await fromZip(await makeDocxWithPicture());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const doc = createNativeWordDocument(view);
    const change = nativeInsertText(doc, 'P0000002', 1, 'Before');

    const result = applyNativeWordChangesToPackage(pkg, change ? [change] : []);

    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);

    const reloaded = await fromZip(await toZip(pkg));
    const reloadedView = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded));
    const paragraph = reloadedView.paragraphs[1]!;
    expect(paragraph.text).toBe('\uFFFCBeforePictureEnd');
    expect(paragraph.segments.map((segment) => segment.kind)).toEqual(['image', 'text']);
    expect(paragraph.segments[1]).toMatchObject({ kind: 'text', text: 'BeforePictureEnd' });
  });

  test('persists image resize in wp:extent and DrawingML transform through DOCX reload', async () => {
    const pkg = await fromZip(await makeDocxWithPicture());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const image = view.paragraphs[1]!.segments.find((segment) => segment.kind === 'image');
    expect(image?.kind).toBe('image');
    const doc = createNativeWordDocument(view);
    const change = image?.kind === 'image'
      ? nativeSetImageSize(doc, 'P0000002', image.drawingId, { width: 192, height: 48 })
      : null;
    expect(change).toMatchObject({ kind: 'image', op: 'resize', width: 192, height: 48 });
    const result = applyNativeWordChangesToPackage(pkg, change ? [change] : []);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);

    const reloaded = await fromZip(await toZip(pkg));
    const drawing = reloaded.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'drawing')[0]!;
    const extent = drawing.getElementsByTagNameNS(WP_NS, 'extent')[0]!;
    const transformExtent = drawing.getElementsByTagNameNS(A_NS, 'ext')[0]!;
    expect(extent.getAttribute('cx')).toBe('1828800');
    expect(extent.getAttribute('cy')).toBe('457200');
    expect(transformExtent.getAttribute('cx')).toBe('1828800');
    expect(transformExtent.getAttribute('cy')).toBe('457200');
  });

  test('persists a base64 image insertion with media, relationship, and content type', async () => {
    const pkg = await fromZip(await makeDocxWithPicture());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const doc = createNativeWordDocument(view);
    const src = `data:image/png;base64,${Buffer.from(TINY_PNG).toString('base64')}`;
    const change = nativeInsertImage(doc, 'P0000001', 5, src, { width: 32, height: 32, drawingId: 'inserted-image' });
    expect(change).toMatchObject({ kind: 'image', op: 'insert', drawingId: 'inserted-image' });
    const result = applyNativeWordChangesToPackage(pkg, change ? [change] : []);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);

    const reloaded = await fromZip(await toZip(pkg));
    expect(reloaded.parts.has('word/media/image2.png')).toBe(true);
    const rels = reloaded.parts.get('word/_rels/document.xml.rels')!.xmlDom!;
    const imageRelationship = Array.from({ length: rels.getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship').length }, (_, index) => rels.getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship')[index]!)
      .find((relationship) => relationship.getAttribute('Type')?.endsWith('/image') && relationship.getAttribute('Target') === 'media/image2.png');
    expect(imageRelationship).toBeTruthy();
    const contentTypes = reloaded.parts.get('[Content_Types].xml')!.xmlDom!;
    expect(Array.from({ length: contentTypes.getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Default').length }, (_, index) => contentTypes.getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Default')[index]!).some((entry) => entry.getAttribute('Extension') === 'png')).toBe(true);
    expect(buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded)).paragraphs[0]!.segments.some((segment) => segment.kind === 'image')).toBe(true);
  });

  test('persists a body comment with comment part, relationship, and range anchors', async () => {
    const pkg = await fromZip(await makeDocxWithPicture());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const doc = createNativeWordDocument(view);
    const change = nativeAddComment(doc, 'P0000001', 0, 5, {
      commentId: 7,
      author: 'Tester',
      date: '2026-08-23T00:00:00.000Z',
      body: 'Check this text',
    });
    expect(change).toMatchObject({ kind: 'annotation', op: 'insert', commentId: 7 });
    const result = applyNativeWordChangesToPackage(pkg, change ? [change] : []);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);

    const reloaded = await fromZip(await toZip(pkg));
    const comments = reloaded.parts.get('word/comments.xml')!.xmlDom!;
    const comment = comments.getElementsByTagNameNS(W_NS, 'comment')[0]!;
    expect(comment.getAttributeNS(W_NS, 'id')).toBe('7');
    expect(comment.textContent).toContain('Check this text');
    const body = reloaded.parts.get('word/document.xml')!.xmlDom!;
    expect(body.getElementsByTagNameNS(W_NS, 'commentRangeStart')).toHaveLength(1);
    expect(body.getElementsByTagNameNS(W_NS, 'commentRangeEnd')).toHaveLength(1);
    expect(body.getElementsByTagNameNS(W_NS, 'commentReference')).toHaveLength(1);
    const reloadedParagraph = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded)).paragraphs[0]!;
    expect(reloadedParagraph.spans.some((span) => span.kind === 'comment' && span.commentId === 7)).toBe(true);

    const reply: NativeWordChange = {
      kind: 'annotation',
      op: 'insert',
      paraId: 'P0000001',
      from: 0,
      to: 5,
      commentId: 8,
      parentCommentId: 7,
      body: 'Reply to this text',
    };
    expect(applyNativeWordChangesToPackage(reloaded, [reply]).skipped).toEqual([]);
    const replied = await fromZip(await toZip(reloaded));
    const extended = replied.parts.get('word/commentsExtended.xml')!.xmlDom!;
    const commentEx = extended.getElementsByTagNameNS('http://schemas.microsoft.com/office/word/2012/wordml', 'commentEx')[0]!;
    expect(commentEx.getAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'paraId')).toBe('00000008');
    expect(commentEx.getAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'paraIdParent')).toBe('00000007');

    const update: NativeWordChange = {
      kind: 'annotation',
      op: 'update',
      paraId: 'P0000001',
      from: 0,
      to: 5,
      commentId: 7,
      body: 'Updated comment',
    };
    expect(applyNativeWordChangesToPackage(replied, [update]).skipped).toEqual([]);
    const updated = await fromZip(await toZip(replied));
    const updatedComments = updated.parts.get('word/comments.xml')!.xmlDom!;
    expect(updatedComments.getElementsByTagNameNS(W_NS, 't')[0]!.textContent).toBe('Updated comment');

    const remove: NativeWordChange = {
      kind: 'annotation',
      op: 'delete',
      paraId: 'P0000001',
      from: 0,
      to: 5,
      commentId: 7,
    };
    expect(applyNativeWordChangesToPackage(updated, [remove]).skipped).toEqual([]);
    const deleted = await fromZip(await toZip(updated));
    const deletedComments = deleted.parts.get('word/comments.xml')!.xmlDom!;
    const deletedBody = deleted.parts.get('word/document.xml')!.xmlDom!;
    const deletedExtended = deleted.parts.get('word/commentsExtended.xml')!.xmlDom!;
    expect(deletedComments.getElementsByTagNameNS(W_NS, 'comment')).toHaveLength(1);
    expect(deletedComments.getElementsByTagNameNS(W_NS, 'comment')[0]!.getAttributeNS(W_NS, 'id')).toBe('8');
    expect(deletedBody.getElementsByTagNameNS(W_NS, 'commentRangeStart')).toHaveLength(1);
    expect(deletedBody.getElementsByTagNameNS(W_NS, 'commentRangeEnd')).toHaveLength(1);
    expect(deletedBody.getElementsByTagNameNS(W_NS, 'commentReference')).toHaveLength(1);
    expect(deletedExtended.getElementsByTagNameNS('http://schemas.microsoft.com/office/word/2012/wordml', 'commentEx')).toHaveLength(0);
  });

  test('writes one comment body for anchors spanning multiple paragraphs', async () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const changes: NativeWordChange[] = [
      { kind: 'annotation', op: 'insert', paraId: 'P0000001', from: 0, to: 5, commentId: 9, body: 'Cross paragraph' },
      { kind: 'annotation', op: 'insert', paraId: 'P0000002', from: 0, to: 5, commentId: 9, body: 'Cross paragraph' },
    ];
    const result = applyNativeWordChangesToPackage(pkg, changes);
    expect(result.skipped).toEqual([]);
    const reloaded = await fromZip(await toZip(pkg));
    const comments = reloaded.parts.get('word/comments.xml')!.xmlDom!;
    const body = reloaded.parts.get('word/document.xml')!.xmlDom!;
    expect(comments.getElementsByTagNameNS(W_NS, 'comment')).toHaveLength(1);
    expect(body.getElementsByTagNameNS(W_NS, 'commentRangeStart')).toHaveLength(2);
    expect(body.getElementsByTagNameNS(W_NS, 'commentRangeEnd')).toHaveLength(2);
    expect(body.getElementsByTagNameNS(W_NS, 'commentReference')).toHaveLength(2);
  });

  test('writes and resolves a paragraph text revision through DOCX reload', async () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const insert: NativeWordChange = {
      kind: 'revision',
      op: 'insert',
      paraId: 'P0000001',
      from: 0,
      to: 5,
      revisionId: '41',
      revisionKind: 'insert',
    };
    expect(applyNativeWordChangesToPackage(pkg, [insert]).skipped).toEqual([]);
    const marked = await fromFlatOpc(toFlatOpc(pkg));
    const markedBody = marked.parts.get('word/document.xml')!.xmlDom!;
    expect(markedBody.getElementsByTagNameNS(W_NS, 'ins')).toHaveLength(1);
    expect(markedBody.getElementsByTagNameNS(W_NS, 'ins')[0]!.textContent).toContain('Hello');

    const accept: NativeWordChange = { ...insert, op: 'accept' };
    expect(applyNativeWordChangesToPackage(marked, [accept]).skipped).toEqual([]);
    const resolved = fromFlatOpc(toFlatOpc(marked));
    const resolvedBody = resolved.parts.get('word/document.xml')!.xmlDom!;
    expect(resolvedBody.getElementsByTagNameNS(W_NS, 'ins')).toHaveLength(0);
    expect(buildWordEditorViewFromPackage(resolved, buildIndices(resolved)).paragraphs[0]!.text).toBe('Hello');

    const deletePkg = fromFlatOpc(FLAT_OPC);
    const deletion: NativeWordChange = { ...insert, op: 'delete', revisionId: '42', revisionKind: 'delete' };
    expect(applyNativeWordChangesToPackage(deletePkg, [deletion]).skipped).toEqual([]);
    const deletedMarked = fromFlatOpc(toFlatOpc(deletePkg));
    const deletedMarkedBody = deletedMarked.parts.get('word/document.xml')!.xmlDom!;
    expect(deletedMarkedBody.getElementsByTagNameNS(W_NS, 'del')).toHaveLength(1);
    expect(deletedMarkedBody.getElementsByTagNameNS(W_NS, 'delText')[0]!.textContent).toBe('Hello');
    const rejectDeletion: NativeWordChange = { ...deletion, op: 'reject' };
    expect(applyNativeWordChangesToPackage(deletedMarked, [rejectDeletion]).skipped).toEqual([]);
    const restored = fromFlatOpc(toFlatOpc(deletedMarked));
    expect(buildWordEditorViewFromPackage(restored, buildIndices(restored)).paragraphs[0]!.text).toBe('Hello');
  });

  test('coalesces consecutive DOCX revision patches and resolves them in batch', () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const first: NativeWordChange = {
      kind: 'revision', op: 'insert', paraId: 'P0000001', from: 0, to: 1, revisionId: '51', revisionKind: 'insert',
    };
    const second: NativeWordChange = { ...first, from: 1, to: 2 };
    const result = applyNativeWordChangesToPackage(pkg, [first, second]);
    expect(result.skipped).toEqual([]);
    const body = pkg.parts.get('word/document.xml')!.xmlDom!;
    const insertions = body.getElementsByTagNameNS(W_NS, 'ins');
    expect(insertions).toHaveLength(1);
    expect(insertions[0]!.textContent).toBe('He');

    const acceptAll = [first, second].map((change) => ({ ...change, op: 'accept' as const }));
    expect(applyNativeWordChangesToPackage(pkg, acceptAll).skipped).toEqual([]);
    expect(body.getElementsByTagNameNS(W_NS, 'ins')).toHaveLength(0);
  });

  test('writes a cross-paragraph revision ID and resolves all matching ranges', () => {
    const pkg = fromFlatOpc(FLAT_OPC);
    const first: NativeWordChange = {
      kind: 'revision', op: 'delete', paraId: 'P0000001', from: 0, to: 5, revisionId: '61', revisionKind: 'delete',
    };
    const second: NativeWordChange = { ...first, paraId: 'P0000002' };
    expect(applyNativeWordChangesToPackage(pkg, [first, second]).skipped).toEqual([]);
    const body = pkg.parts.get('word/document.xml')!.xmlDom!;
    expect(body.getElementsByTagNameNS(W_NS, 'del')).toHaveLength(2);
    expect(body.getElementsByTagNameNS(W_NS, 'delText')).toHaveLength(2);

    const rejectAll = [first, second].map((change) => ({ ...change, op: 'reject' as const }));
    expect(applyNativeWordChangesToPackage(pkg, rejectAll).skipped).toEqual([]);
    expect(body.getElementsByTagNameNS(W_NS, 'del')).toHaveLength(0);
    expect(buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).paragraphs.map((paragraph) => paragraph.text)).toEqual(['Hello', 'World']);
  });

  test('resolves a revision spanning multiple Word runs without leaving wrappers behind', () => {
    const markedOpc = FLAT_OPC.replace(
      '<w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p>',
      '<w:p w14:paraId="P0000001"><w:ins w:id="71" w:author="Tester" w:date="2026-08-24T00:00:00Z"><w:r><w:t>He</w:t></w:r><w:r><w:t>llo</w:t></w:r></w:ins></w:p>',
    );
    const accepted = fromFlatOpc(markedOpc);
    const accept: NativeWordChange = {
      kind: 'revision',
      op: 'accept',
      paraId: 'P0000001',
      from: 0,
      to: 5,
      revisionId: '71',
      revisionKind: 'insert',
    };
    expect(applyNativeWordChangesToPackage(accepted, [accept]).skipped).toEqual([]);
    const acceptedBody = accepted.parts.get('word/document.xml')!.xmlDom!;
    expect(acceptedBody.getElementsByTagNameNS(W_NS, 'ins')).toHaveLength(0);
    expect(buildWordEditorViewFromPackage(accepted, buildIndices(accepted)).paragraphs[0]!.text).toBe('Hello');

    const rejected = fromFlatOpc(markedOpc);
    expect(applyNativeWordChangesToPackage(rejected, [{ ...accept, op: 'reject' }]).skipped).toEqual([]);
    const rejectedBody = rejected.parts.get('word/document.xml')!.xmlDom!;
    expect(rejectedBody.getElementsByTagNameNS(W_NS, 'ins')).toHaveLength(0);
    expect(buildWordEditorViewFromPackage(rejected, buildIndices(rejected)).paragraphs[0]!.text).toBe('');
  });

  test('resolves mixed insert and delete revisions across paragraphs', () => {
    const markedOpc = FLAT_OPC
      .replace(
        '<w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p>',
        '<w:p w14:paraId="P0000001"><w:r><w:t>Hi</w:t></w:r><w:ins w:id="81" w:author="Tester" w:date="2026-08-24T00:00:00Z"><w:r><w:t> there</w:t></w:r></w:ins></w:p>',
      )
      .replace(
        '<w:p w14:paraId="P0000002"><w:r><w:t>World</w:t></w:r></w:p>',
        '<w:p w14:paraId="P0000002"><w:del w:id="82" w:author="Tester" w:date="2026-08-24T00:00:00Z"><w:r><w:delText>World</w:delText></w:r></w:del></w:p>',
      );
    const changes: NativeWordChange[] = [
      { kind: 'revision', op: 'accept', paraId: 'P0000001', from: 2, to: 8, revisionId: '81', revisionKind: 'insert' },
      { kind: 'revision', op: 'reject', paraId: 'P0000002', from: 0, to: 5, revisionId: '82', revisionKind: 'delete' },
    ];
    const resolved = fromFlatOpc(markedOpc);
    expect(applyNativeWordChangesToPackage(resolved, changes).skipped).toEqual([]);
    const body = resolved.parts.get('word/document.xml')!.xmlDom!;
    expect(body.getElementsByTagNameNS(W_NS, 'ins')).toHaveLength(0);
    expect(body.getElementsByTagNameNS(W_NS, 'del')).toHaveLength(0);
    expect(buildWordEditorViewFromPackage(resolved, buildIndices(resolved)).paragraphs.map((paragraph) => paragraph.text)).toEqual(['Hi there', 'World']);

    const opposite = fromFlatOpc(markedOpc);
    expect(applyNativeWordChangesToPackage(opposite, changes.map((change) => ({
      ...change,
      op: change.op === 'accept' ? 'reject' : 'accept',
    }))).skipped).toEqual([]);
    expect(buildWordEditorViewFromPackage(opposite, buildIndices(opposite)).paragraphs.map((paragraph) => paragraph.text)).toEqual(['Hi', '']);
  });

  test('writes external hyperlink relationships and preserves selected run formatting', () => {
    const pkg = fromFlatOpc(HYPERLINK_FLAT_OPC);
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const doc = createNativeWordDocument(view);
    const change = nativeSetHyperlink(doc, 'LINK0001', 0, 5, {
      target: 'https://coco.example/docs',
      tooltip: 'Coco docs',
    });
    expect(change).not.toBeNull();

    const result = applyNativeWordChangesToPackage(pkg, change ? [change] : []);
    expect(result).toMatchObject({
      plan: { unsupported: [], patches: [{ kind: 'hyperlink', op: 'set', paraId: 'LINK0001' }] },
      applied: [{ kind: 'hyperlink', op: 'set', paraId: 'LINK0001' }],
      skipped: [],
      unresolved: [],
    });

    const reloaded = fromFlatOpc(toFlatOpc(pkg));
    const reloadedView = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded));
    expect(reloadedView.paragraphs[0]?.segments).toMatchObject([
      { kind: 'text', text: 'Hello', format: { bold: true }, hyperlink: { target: 'https://coco.example/docs', tooltip: 'Coco docs' } },
      { kind: 'text', text: ' world' },
    ]);
    const rels = reloaded.parts.get('word/_rels/document.xml.rels')!.xmlDom!;
    const relationship = rels.getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship')[0]!;
    expect(relationship.getAttribute('Target')).toBe('https://coco.example/docs');
    expect(relationship.getAttribute('TargetMode')).toBe('External');
  });

  test('removes an external hyperlink wrapper without dropping the run content', () => {
    const pkg = fromFlatOpc(HYPERLINK_FLAT_OPC);
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const doc = createNativeWordDocument(view);
    const set = nativeSetHyperlink(doc, 'LINK0001', 0, 5, { target: 'https://coco.example/docs' });
    expect(set).not.toBeNull();
    applyNativeWordChangesToPackage(pkg, set ? [set] : []);

    const reloadedAfterSet = fromFlatOpc(toFlatOpc(pkg));
    const removeDoc = createNativeWordDocument(buildWordEditorViewFromPackage(reloadedAfterSet, buildIndices(reloadedAfterSet)));
    const remove = nativeSetHyperlink(removeDoc, 'LINK0001', 0, 5, null);
    expect(remove).toMatchObject({ kind: 'hyperlink', op: 'remove' });
    const result = applyNativeWordChangesToPackage(reloadedAfterSet, remove ? [remove] : []);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);

    const reloaded = fromFlatOpc(toFlatOpc(reloadedAfterSet));
    const paragraph = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded)).paragraphs[0]!;
    expect(paragraph.text).toBe('Hello world');
    expect(paragraph.segments[0]).toMatchObject({ kind: 'text', text: 'Hello', format: { bold: true } });
    expect(paragraph.segments[0]?.hyperlink).toBeUndefined();
  });

  test('writes regular table row/column and horizontal merge/split changes through DOCX reload', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const tableId = view.tables[0]!.tableId;
    const doc = createNativeWordDocument(view);
    const merge = nativeMergeTableCells(doc, tableId, {
      anchor: { rowIndex: 0, columnIndex: 0 },
      focus: { rowIndex: 0, columnIndex: 1 },
    });
    expect(merge?.change).toMatchObject({ kind: 'table', op: 'merge-cells', endColumnIndex: 1 });
    const mergedResult = applyNativeWordChangesToPackage(pkg, merge ? [merge.change] : []);
    expect(mergedResult.skipped).toEqual([]);
    expect(mergedResult.unresolved).toEqual([]);

    let table = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'tbl')[0]!;
    let firstRow = table.getElementsByTagNameNS(W_NS, 'tr')[0]!;
    expect(firstRow.getElementsByTagNameNS(W_NS, 'tc')).toHaveLength(1);
    expect(firstRow.getElementsByTagNameNS(W_NS, 'gridSpan')[0]!.getAttributeNS(W_NS, 'val')).toBe('2');

    const splitPkg = fromFlatOpc(toFlatOpc(pkg));
    const splitView = buildWordEditorViewFromPackage(splitPkg, buildIndices(splitPkg));
    const splitDoc = createNativeWordDocument(splitView);
    const split = nativeSplitTableCell(splitDoc, tableId, 0, 0);
    expect(split?.change).toMatchObject({ kind: 'table', op: 'split-cell' });
    const splitResult = applyNativeWordChangesToPackage(splitPkg, split ? [split.change] : []);
    expect(splitResult.skipped).toEqual([]);
    expect(splitResult.unresolved).toEqual([]);
    table = fromFlatOpc(toFlatOpc(splitPkg)).parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'tbl')[0]!;
    firstRow = table.getElementsByTagNameNS(W_NS, 'tr')[0]!;
    expect(firstRow.getElementsByTagNameNS(W_NS, 'tc')).toHaveLength(2);
    expect(firstRow.getElementsByTagNameNS(W_NS, 'gridSpan')).toHaveLength(0);

    const rowColumnPkg = fromFlatOpc(TABLE_FLAT_OPC);
    const rowColumnResult = applyNativeWordChangesToPackage(rowColumnPkg, [
      { kind: 'table', op: 'insert-row', tableId, rowIndex: 1, columnIndex: 0 },
      { kind: 'table', op: 'insert-column', tableId, rowIndex: 0, columnIndex: 1 },
    ]);
    expect(rowColumnResult.skipped).toEqual([]);
    expect(rowColumnResult.unresolved).toEqual([]);
    table = fromFlatOpc(toFlatOpc(rowColumnPkg)).parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'tbl')[0]!;
    expect(table.getElementsByTagNameNS(W_NS, 'tr')).toHaveLength(3);
    expect(table.getElementsByTagNameNS(W_NS, 'gridCol')).toHaveLength(3);
    expect(table.getElementsByTagNameNS(W_NS, 'tr')[0]!.getElementsByTagNameNS(W_NS, 'tc')).toHaveLength(3);

    const deletePkg = fromFlatOpc(TABLE_FLAT_OPC);
    const deleteResult = applyNativeWordChangesToPackage(deletePkg, [
      { kind: 'table', op: 'delete-row', tableId, rowIndex: 1, columnIndex: 0 },
      { kind: 'table', op: 'delete-column', tableId, rowIndex: 0, columnIndex: 0 },
    ]);
    expect(deleteResult.skipped).toEqual([]);
    expect(deleteResult.unresolved).toEqual([]);
    table = fromFlatOpc(toFlatOpc(deletePkg)).parts.get('word/document.xml')!.xmlDom!
      .getElementsByTagNameNS(W_NS, 'tbl')[0]!;
    expect(table.getElementsByTagNameNS(W_NS, 'tr')).toHaveLength(1);
    expect(table.getElementsByTagNameNS(W_NS, 'gridCol')).toHaveLength(1);
    expect(table.getElementsByTagNameNS(W_NS, 'tr')[0]!.getElementsByTagNameNS(W_NS, 'tc')).toHaveLength(1);
  });

  test('writes tracked regular table accept/reject actions without losing the original table', () => {
    const source = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(source, buildIndices(source)).tables[0]!.tableId;
    const trackedInsert: NativeWordChange = {
      kind: 'table',
      op: 'insert-row',
      tableId,
      rowIndex: 1,
      columnIndex: 0,
      revisionId: 'table-revision-1',
      revisionKind: 'insert',
    };

    const accepted = fromFlatOpc(TABLE_FLAT_OPC);
    const acceptResult = applyNativeWordChangesToPackage(accepted, [
      trackedInsert,
      { ...trackedInsert, revisionAction: 'accept' },
    ]);
    expect(acceptResult.skipped).toEqual([]);
    expect(acceptResult.unresolved).toEqual([]);
    expect(
      fromFlatOpc(toFlatOpc(accepted)).parts.get('word/document.xml')!.xmlDom!
        .getElementsByTagNameNS(W_NS, 'tr'),
    ).toHaveLength(3);

    const rejected = fromFlatOpc(TABLE_FLAT_OPC);
    const rejectResult = applyNativeWordChangesToPackage(rejected, [
      trackedInsert,
      { ...trackedInsert, revisionAction: 'reject' },
    ]);
    expect(rejectResult.skipped).toEqual([]);
    expect(rejectResult.unresolved).toEqual([]);
    expect(
      fromFlatOpc(toFlatOpc(rejected)).parts.get('word/document.xml')!.xmlDom!
        .getElementsByTagNameNS(W_NS, 'tr'),
    ).toHaveLength(2);

    const deleted = fromFlatOpc(TABLE_FLAT_OPC);
    const trackedDelete: NativeWordChange = {
      kind: 'table',
      op: 'delete-table',
      tableId,
      rowIndex: 0,
      columnIndex: 0,
      revisionId: 'table-revision-2',
      revisionKind: 'delete',
    };
    const deleteResult = applyNativeWordChangesToPackage(deleted, [
      trackedDelete,
      { ...trackedDelete, revisionAction: 'accept' },
    ]);
    expect(deleteResult.skipped).toEqual([]);
    expect(deleteResult.unresolved).toEqual([]);
    expect(fromFlatOpc(toFlatOpc(deleted)).parts.get('word/document.xml')!.xmlDom!
      .getElementsByTagNameNS(W_NS, 'tbl')).toHaveLength(0);

    const kept = fromFlatOpc(TABLE_FLAT_OPC);
    const rejectDeleteResult = applyNativeWordChangesToPackage(kept, [
      trackedDelete,
      { ...trackedDelete, revisionAction: 'reject' },
    ]);
    expect(rejectDeleteResult.skipped).toEqual([]);
    expect(rejectDeleteResult.unresolved).toEqual([]);
    expect(fromFlatOpc(toFlatOpc(kept)).parts.get('word/document.xml')!.xmlDom!
      .getElementsByTagNameNS(W_NS, 'tbl')).toHaveLength(1);
  });

  test('resolves all tracked table structure revisions instead of only the latest operation', () => {
    const source = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(source, buildIndices(source)).tables[0]!.tableId;
    const insertRow: NativeWordChange = {
      kind: 'table',
      op: 'insert-row',
      tableId,
      rowIndex: 1,
      columnIndex: 0,
      revisionId: 'table-revision-row',
      revisionKind: 'insert',
    };
    const insertColumn: NativeWordChange = {
      kind: 'table',
      op: 'insert-column',
      tableId,
      rowIndex: 0,
      columnIndex: 1,
      revisionId: 'table-revision-column',
      revisionKind: 'insert',
    };

    const accepted = fromFlatOpc(TABLE_FLAT_OPC);
    const acceptResult = applyNativeWordChangesToPackage(accepted, [
      insertRow,
      insertColumn,
      { ...insertRow, revisionAction: 'accept' },
      { ...insertColumn, revisionAction: 'accept' },
    ]);
    expect(acceptResult.skipped).toEqual([]);
    expect(acceptResult.unresolved).toEqual([]);
    const acceptedTable = fromFlatOpc(toFlatOpc(accepted)).parts.get('word/document.xml')!.xmlDom!
      .getElementsByTagNameNS(W_NS, 'tbl')[0]!;
    expect(acceptedTable.getElementsByTagNameNS(W_NS, 'tr')).toHaveLength(3);
    expect(acceptedTable.getElementsByTagNameNS(W_NS, 'gridCol')).toHaveLength(3);

    const rejected = fromFlatOpc(TABLE_FLAT_OPC);
    const rejectResult = applyNativeWordChangesToPackage(rejected, [
      insertRow,
      insertColumn,
      { ...insertRow, revisionAction: 'reject' },
      { ...insertColumn, revisionAction: 'reject' },
    ]);
    expect(rejectResult.skipped).toEqual([]);
    expect(rejectResult.unresolved).toEqual([]);
    const rejectedTable = fromFlatOpc(toFlatOpc(rejected)).parts.get('word/document.xml')!.xmlDom!
      .getElementsByTagNameNS(W_NS, 'tbl')[0]!;
    expect(rejectedTable.getElementsByTagNameNS(W_NS, 'tr')).toHaveLength(2);
    expect(rejectedTable.getElementsByTagNameNS(W_NS, 'gridCol')).toHaveLength(2);
  });

  test('writes table cell background, alignment, and border formats', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).tables[0]!.tableId;
    const format: NativeWordChange = {
      kind: 'table',
      op: 'format-cell',
      tableId,
      rowIndex: 0,
      columnIndex: 0,
      format: { backgroundColor: '#AABBCC', verticalAlignment: 'center', borderColor: '#112233' },
    };
    const result = applyNativeWordChangesToPackage(pkg, [format]);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);
    const document = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!;
    const cell = document.getElementsByTagNameNS(W_NS, 'tc')[0]!;
    const properties = cell.getElementsByTagNameNS(W_NS, 'tcPr')[0]!;
    expect(properties.getElementsByTagNameNS(W_NS, 'shd')[0]!.getAttributeNS(W_NS, 'fill')).toBe('AABBCC');
    expect(properties.getElementsByTagNameNS(W_NS, 'vAlign')[0]!.getAttributeNS(W_NS, 'val')).toBe('center');
    expect(properties.getElementsByTagNameNS(W_NS, 'tcBorders')[0]!.getElementsByTagNameNS(W_NS, 'top')[0]!
      .getAttributeNS(W_NS, 'color')).toBe('112233');
  });

  test('writes clearing a table cell without removing cell formatting', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).tables[0]!.tableId;
    const result = applyNativeWordChangesToPackage(pkg, [{
      kind: 'table', op: 'clear-cell', tableId, rowIndex: 0, columnIndex: 0,
    }]);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);
    const document = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!;
    const cell = document.getElementsByTagNameNS(W_NS, 'tc')[0]!;
    expect(cell.textContent).toBe('');
    expect(cell.getElementsByTagNameNS(W_NS, 'p')).toHaveLength(1);
  });

  test('writes plain text table cell replacement while preserving cell properties', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).tables[0]!.tableId;
    const result = applyNativeWordChangesToPackage(pkg, [{
      kind: 'table', op: 'replace-cell', tableId, rowIndex: 0, columnIndex: 0, text: 'Updated cell',
    }]);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);
    const document = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!;
    const cell = document.getElementsByTagNameNS(W_NS, 'tc')[0]!;
    expect(cell.textContent).toBe('Updated cell');
    expect(cell.getElementsByTagNameNS(W_NS, 'tcPr')).toHaveLength(1);
  });

  test('writes rich table cell replacement with basic run formatting', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).tables[0]!.tableId;
    const paragraph = {
      kind: 'paragraph' as const,
      paraId: 'rich-cell-paragraph',
      story: 'word/document.xml',
      styleId: null,
      format: null,
      content: [{ kind: 'text' as const, text: 'Rich cell', format: { bold: true, color: 'C00000', fontSize: 14 } }],
    };
    const richResult = applyNativeWordChangesToPackage(pkg, [{
      kind: 'table', op: 'replace-cell', tableId, rowIndex: 0, columnIndex: 0,
      cellSnapshot: { blocks: [paragraph], paragraphs: [paragraph], gridSpan: 1, width: 2000 },
    }]);
    expect(richResult.skipped).toEqual([]);
    expect(richResult.unresolved).toEqual([]);
    const document = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!;
    const cell = document.getElementsByTagNameNS(W_NS, 'tc')[0]!;
    expect(cell.textContent).toBe('Rich cell');
    expect(cell.getElementsByTagNameNS(W_NS, 'b')).toHaveLength(1);
    expect(cell.getElementsByTagNameNS(W_NS, 'color')[0]!.getAttributeNS(W_NS, 'val')).toBe('C00000');
    expect(cell.getElementsByTagNameNS(W_NS, 'sz')[0]!.getAttributeNS(W_NS, 'val')).toBe('28');
  });

  test('writes an image in a rich table cell replacement through DOCX reload', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).tables[0]!.tableId;
    const result = applyNativeWordChangesToPackage(pkg, [{
      kind: 'table',
      op: 'replace-cell',
      tableId,
      rowIndex: 0,
      columnIndex: 0,
      cellSnapshot: {
        blocks: [{
          kind: 'paragraph',
          paraId: 'image-cell-paragraph',
          story: 'word/document.xml',
          styleId: null,
          format: null,
          content: [{
            kind: 'image',
            drawingId: 'image-cell-1',
            src: 'data:image/png;base64,AA==',
            width: 40,
            height: 20,
          }],
        }],
        paragraphs: [],
        gridSpan: 1,
        width: 2000,
      },
    }]);
    expect(result.applied).toHaveLength(1);
    expect(result.plan.unsupported).toEqual([]);
    const document = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!;
    expect(document.getElementsByTagNameNS(W_NS, 'drawing')).toHaveLength(1);
    expect(Array.from(pkg.parts.keys()).some((path) => /^word\/media\/image\d+\.png$/.test(path))).toBe(true);
    const relationships = pkg.parts.get('word/_rels/document.xml.rels')?.xmlDom;
    expect(relationships?.getElementsByTagNameNS(PACKAGE_RELS_NS, 'Relationship').length).toBeGreaterThan(0);
  });

  test('writes an external hyperlink in a rich table cell replacement', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).tables[0]!.tableId;
    const result = applyNativeWordChangesToPackage(pkg, [{
      kind: 'table',
      op: 'replace-cell',
      tableId,
      rowIndex: 0,
      columnIndex: 0,
      cellSnapshot: {
        blocks: [{
          kind: 'paragraph',
          paraId: 'hyperlink-cell-paragraph',
          story: 'word/document.xml',
          styleId: null,
          format: null,
          content: [{
            kind: 'text',
            text: 'Coco docs',
            format: null,
            hyperlink: { target: 'https://coco.example/docs' },
          }],
        }],
        paragraphs: [],
        gridSpan: 1,
        width: 2000,
      },
    }]);
    expect(result.applied).toHaveLength(1);
    expect(result.plan.unsupported).toEqual([]);
    const document = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!;
    expect(document.getElementsByTagNameNS(W_NS, 'hyperlink')).toHaveLength(1);
    const relationships = pkg.parts.get('word/_rels/document.xml.rels')?.xmlDom;
    const external = Array.from(relationships?.getElementsByTagNameNS(PACKAGE_RELS_NS, 'Relationship') ?? [])
      .find((relationship) => relationship.getAttribute('Target') === 'https://coco.example/docs');
    expect(external?.getAttribute('TargetMode')).toBe('External');
  });

  test('writes a single-column vertical table merge', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).tables[0]!.tableId;
    const result = applyNativeWordChangesToPackage(pkg, [{
      kind: 'table', op: 'merge-cells', tableId, rowIndex: 0, columnIndex: 0, endRowIndex: 1,
    }]);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);
    const document = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!;
    const rows = document.getElementsByTagNameNS(W_NS, 'tr');
    const topCell = rows[0]!.getElementsByTagNameNS(W_NS, 'tc')[0]!;
    const lowerCell = rows[1]!.getElementsByTagNameNS(W_NS, 'tc')[0]!;
    expect(topCell.textContent).toContain('A');
    expect(lowerCell.textContent).toBe('');
    expect(topCell.getElementsByTagNameNS(W_NS, 'vMerge')[0]!.getAttributeNS(W_NS, 'val')).toBe('restart');
    expect(lowerCell.getElementsByTagNameNS(W_NS, 'vMerge')[0]!.getAttributeNS(W_NS, 'val')).toBe('continue');
  });

  test('writes a single-column vertical table split', () => {
    const pkg = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(pkg, buildIndices(pkg)).tables[0]!.tableId;
    const result = applyNativeWordChangesToPackage(pkg, [
      { kind: 'table', op: 'merge-cells', tableId, rowIndex: 0, columnIndex: 0, endRowIndex: 1 },
      { kind: 'table', op: 'split-cell', tableId, rowIndex: 0, columnIndex: 0, endRowIndex: 1 },
    ]);
    expect(result.skipped).toEqual([]);
    expect(result.unresolved).toEqual([]);
    const document = fromFlatOpc(toFlatOpc(pkg)).parts.get('word/document.xml')!.xmlDom!;
    const rows = document.getElementsByTagNameNS(W_NS, 'tr');
    const topCell = rows[0]!.getElementsByTagNameNS(W_NS, 'tc')[0]!;
    const lowerCell = rows[1]!.getElementsByTagNameNS(W_NS, 'tc')[0]!;
    expect(topCell.getElementsByTagNameNS(W_NS, 'vMerge')).toHaveLength(0);
    expect(lowerCell.getElementsByTagNameNS(W_NS, 'vMerge')).toHaveLength(0);
  });

  test('writes tracked vertical table merge accept and reject actions', () => {
    const source = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(source, buildIndices(source)).tables[0]!.tableId;
    const trackedMerge: NativeWordChange = {
      kind: 'table',
      op: 'merge-cells',
      tableId,
      rowIndex: 0,
      columnIndex: 0,
      endRowIndex: 1,
      revisionId: 'vertical-merge-revision-1',
      revisionKind: 'insert',
    };
    const accepted = fromFlatOpc(TABLE_FLAT_OPC);
    const acceptResult = applyNativeWordChangesToPackage(accepted, [
      trackedMerge,
      { ...trackedMerge, revisionAction: 'accept' },
    ]);
    expect(acceptResult.skipped).toEqual([]);
    expect(acceptResult.unresolved).toEqual([]);
    const acceptedRows = fromFlatOpc(toFlatOpc(accepted)).parts.get('word/document.xml')!.xmlDom!
      .getElementsByTagNameNS(W_NS, 'tr');
    expect(acceptedRows[0]!.getElementsByTagNameNS(W_NS, 'vMerge')[0]!.getAttributeNS(W_NS, 'val')).toBe('restart');
    expect(acceptedRows[1]!.getElementsByTagNameNS(W_NS, 'vMerge')[0]!.getAttributeNS(W_NS, 'val')).toBe('continue');

    const rejected = fromFlatOpc(TABLE_FLAT_OPC);
    const rejectResult = applyNativeWordChangesToPackage(rejected, [
      trackedMerge,
      { ...trackedMerge, revisionAction: 'reject' },
    ]);
    expect(rejectResult.skipped).toEqual([]);
    expect(rejectResult.unresolved).toEqual([]);
    const rejectedDocument = fromFlatOpc(toFlatOpc(rejected)).parts.get('word/document.xml')!.xmlDom!;
    expect(rejectedDocument.getElementsByTagNameNS(W_NS, 'vMerge')).toHaveLength(0);
  });

  test('writes a mixed tracked vertical merge and split sequence as one final table state', () => {
    const source = fromFlatOpc(TABLE_FLAT_OPC);
    const tableId = buildWordEditorViewFromPackage(source, buildIndices(source)).tables[0]!.tableId;
    const merge: NativeWordChange = {
      kind: 'table',
      op: 'merge-cells',
      tableId,
      rowIndex: 0,
      columnIndex: 0,
      endRowIndex: 1,
      revisionId: 'vertical-merge-revision-2',
      revisionKind: 'insert',
    };
    const split: NativeWordChange = {
      ...merge,
      op: 'split-cell',
      revisionId: 'vertical-split-revision-1',
      revisionKind: 'delete',
    };

    const accepted = fromFlatOpc(TABLE_FLAT_OPC);
    const acceptResult = applyNativeWordChangesToPackage(accepted, [
      merge,
      split,
      { ...merge, revisionAction: 'accept' },
      { ...split, revisionAction: 'accept' },
    ]);
    expect(acceptResult.skipped).toEqual([]);
    expect(acceptResult.unresolved).toEqual([]);
    const acceptedDocument = fromFlatOpc(toFlatOpc(accepted)).parts.get('word/document.xml')!.xmlDom!;
    expect(acceptedDocument.getElementsByTagNameNS(W_NS, 'vMerge')).toHaveLength(0);

    const rejected = fromFlatOpc(TABLE_FLAT_OPC);
    const rejectResult = applyNativeWordChangesToPackage(rejected, [
      merge,
      split,
      { ...merge, revisionAction: 'reject' },
      { ...split, revisionAction: 'reject' },
    ]);
    expect(rejectResult.skipped).toEqual([]);
    expect(rejectResult.unresolved).toEqual([]);
    const rejectedDocument = fromFlatOpc(toFlatOpc(rejected)).parts.get('word/document.xml')!.xmlDom!;
    expect(rejectedDocument.getElementsByTagNameNS(W_NS, 'vMerge')).toHaveLength(0);
  });

  test('writes a tracked table insertion at its paragraph anchor', () => {
    const tableSnapshot = {
      kind: 'table' as const,
      tableId: 'runtime-table-1',
      columnWidths: [1800, 1800],
      rows: [{
        cells: [0, 1].map((index) => {
          const paragraph = {
            kind: 'paragraph' as const,
            paraId: `runtime-cell-${index}`,
            story: 'word/document.xml',
            styleId: null,
            format: null,
            content: [{ kind: 'text' as const, text: `Cell ${index + 1}`, format: null }],
          };
          return { blocks: [paragraph], paragraphs: [paragraph], gridSpan: 1, width: 1800 };
        }),
      }],
    };
    const insert: NativeWordChange = {
      kind: 'table',
      op: 'insert-table',
      tableId: 'runtime-table-1',
      rowIndex: 0,
      columnIndex: 0,
      revisionId: 'table-insert-revision-1',
      revisionKind: 'insert',
      anchorParaId: 'P0000001',
      anchorOffset: 0,
      tableSnapshot,
    };
    const accepted = fromFlatOpc(FLAT_OPC);
    const acceptResult = applyNativeWordChangesToPackage(accepted, [insert, { ...insert, revisionAction: 'accept' }]);
    expect(acceptResult.skipped).toEqual([]);
    expect(acceptResult.unresolved).toEqual([]);
    const acceptedDocument = fromFlatOpc(toFlatOpc(accepted)).parts.get('word/document.xml')!.xmlDom!;
    const acceptedTables = acceptedDocument.getElementsByTagNameNS(W_NS, 'tbl');
    expect(acceptedTables).toHaveLength(1);
    expect(acceptedTables[0]!.textContent).toContain('Cell 1');

    const rejected = fromFlatOpc(FLAT_OPC);
    const rejectResult = applyNativeWordChangesToPackage(rejected, [insert, { ...insert, revisionAction: 'reject' }]);
    expect(rejectResult.skipped).toEqual([]);
    expect(rejectResult.unresolved).toEqual([]);
    expect(fromFlatOpc(toFlatOpc(rejected)).parts.get('word/document.xml')!.xmlDom!
      .getElementsByTagNameNS(W_NS, 'tbl')).toHaveLength(0);
  });
});
