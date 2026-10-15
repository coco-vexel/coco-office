import assert from 'node:assert/strict';
import {
  flatOpcToOnlyOfficeWriteback,
  jsonToFlatOpc,
  withPackageLongTailParts,
} from '../src/index.mjs';

const CT = {
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
};

const doc = {
  type: 'document',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'run', content: ['Hello writeback'] },
        { type: 'endRun', content: [] },
      ],
    },
  ],
};

const hyperlinkDoc = {
  type: 'document',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'run', content: ['Before link '] },
        {
          type: 'hyperlink',
          value: 'https://example.com/coco-link',
          tooltip: 'Coco link tooltip',
          content: [{ type: 'run', content: ['Coco link'] }],
        },
        { type: 'run', content: [' after'] },
        { type: 'endRun', content: [] },
      ],
    },
  ],
};

const bookmarkDoc = {
  type: 'document',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'run', content: ['Before bookmark '] },
        { type: 'run', content: ['Bookmark anchor'] },
        { type: 'run', content: [' after'] },
        { type: 'endRun', content: [] },
      ],
    },
  ],
};

const bodyImageDoc = {
  type: 'document',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'run', content: ['Body image '] },
        {
          type: 'run',
          content: [
            {
              type: 'paraDrawing',
              extent: { cx: 914400, cy: 304800 },
              graphic: {
                type: 'image',
                blipFill: {
                  rasterImageId: 'data:image/png;base64,iVBORw0KGgo=',
                },
              },
            },
          ],
        },
        { type: 'endRun', content: [] },
      ],
    },
  ],
};

const commentedDoc = {
  type: 'document',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'run', content: ['Prefix '] },
        { type: 'commentRangeStart', id: '7' },
        { type: 'run', content: ['Commented anchor'] },
        { type: 'commentRangeEnd', id: '7' },
        { type: 'run', content: [' suffix'] },
        { type: 'endRun', content: [] },
      ],
    },
  ],
};

const headerFooterDoc = {
  type: 'document',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'run', content: ['Body with header footer'] },
        { type: 'endRun', content: [] },
      ],
    },
  ],
  sectPr: {
    headerReference: { default: { id: 'rIdHeader' } },
    footerReference: { even: { id: 'rIdFooter' } },
  },
};

function makeFlatOpc(extraParts = []) {
  return jsonToFlatOpc(doc, { extraParts });
}

function makeHyperlinkFlatOpc() {
  return jsonToFlatOpc(hyperlinkDoc);
}

function makeBookmarkFlatOpc() {
  return jsonToFlatOpc(bookmarkDoc).replace(
    '<w:r><w:t xml:space="preserve">Bookmark anchor</w:t></w:r>',
    '<w:bookmarkStart w:id="42" w:name="coco_bookmark_test"/><w:r><w:t xml:space="preserve">Bookmark anchor</w:t></w:r><w:bookmarkEnd w:id="42"/>',
  );
}

function makeBodyImageFlatOpc() {
  const longTail = withPackageLongTailParts(bodyImageDoc);
  return jsonToFlatOpc(longTail.doc, { extraParts: longTail.extraParts });
}

function makeAnchoredBodyImageFlatOpc() {
  return makeBodyImageFlatOpc()
    .replace(
      /<wp:inline\b[^>]*>/,
      '<wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">',
    )
    .replace('</wp:inline>', '</wp:anchor>');
}

function makeChartDrawingFlatOpc() {
  return makeFlatOpc([
    {
      name: '/word/charts/chart1.xml',
      contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      xml: '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>',
      rel: {
        id: 'rIdChart1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
        target: 'charts/chart1.xml',
      },
    },
  ]).replace(
    '<w:r><w:t xml:space="preserve">Hello writeback</w:t></w:r>',
    '<w:r><w:t xml:space="preserve">Chart before </w:t></w:r><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="11" name="Chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>',
  );
}

function makeTextBoxDrawingFlatOpc() {
  return makeFlatOpc().replace(
    '<w:r><w:t xml:space="preserve">Hello writeback</w:t></w:r>',
    '<w:r><w:t xml:space="preserve">Text box before </w:t></w:r><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="12" name="Text Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>Inside text box</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>',
  );
}

function makeCommentFlatOpc() {
  return jsonToFlatOpc(commentedDoc, {
    extraParts: [
      {
        name: '/word/comments.xml',
        contentType: CT.comments,
        xml: '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="7" w:author="Coco"><w:p><w:r><w:t>Review note</w:t></w:r></w:p></w:comment></w:comments>',
      },
    ],
  });
}

function makeNotesFlatOpc() {
  return jsonToFlatOpc(
    {
      type: 'document',
      content: [
        {
          type: 'paragraph',
          content: [
            {
              type: 'run',
              content: [
                'Footnote anchor',
                { type: 'footnoteRef', footnote: '101' },
                ' and endnote anchor',
                { type: 'endnoteRef', footnote: '202' },
              ],
            },
            { type: 'endRun', content: [] },
          ],
        },
      ],
    },
    {
      extraParts: [
        {
          name: '/word/footnotes.xml',
          contentType: CT.footnotes,
          xml: '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="101"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:i/><w:sz w:val="28"/><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:eastAsia="Calibri Light"/><w:color w:val="C00000"/></w:rPr><w:t>Footnote body</w:t></w:r></w:p><w:p><w:r><w:t>Footnote second paragraph</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="304800"/><wp:docPr id="9" name="Footnote Image"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="9" name="Footnote Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdFootnoteImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Footnote table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:footnote></w:footnotes>',
          rel: {
            id: 'rIdFootnotes',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
            target: 'footnotes.xml',
          },
        },
        {
          name: '/word/media/image1.png',
          contentType: 'image/png',
          binary: 'iVBORw0KGgo=',
          rel: {
            source: '/word/footnotes.xml',
            id: 'rIdFootnoteImage',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            target: 'media/image1.png',
          },
        },
        {
          name: '/word/endnotes.xml',
          contentType: CT.endnotes,
          xml: '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="202"><w:p><w:r><w:rPr><w:b/><w:sz w:val="26"/></w:rPr><w:t>Endnote body</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Endnote table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:endnote></w:endnotes>',
          rel: {
            id: 'rIdEndnotes',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
            target: 'endnotes.xml',
          },
        },
      ],
    },
  );
}

function makeHeaderFooterFlatOpc() {
  return jsonToFlatOpc(headerFooterDoc, {
    extraParts: [
      {
        name: '/word/header1.xml',
        contentType: CT.header,
        xml: '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>',
        rel: {
          id: 'rIdHeader',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
          target: 'header1.xml',
        },
      },
      {
        name: '/word/footer1.xml',
        contentType: CT.footer,
        xml: '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>',
        rel: {
          id: 'rIdFooter',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
          target: 'footer1.xml',
        },
      },
    ],
  });
}

function makeHeaderTableFlatOpc() {
  return jsonToFlatOpc(headerFooterDoc, {
    extraParts: [
      {
        name: '/word/header1.xml',
        contentType: CT.header,
        xml: '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header before table</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Header table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>',
        rel: {
          id: 'rIdHeader',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
          target: 'header1.xml',
        },
      },
    ],
  });
}

function makeHeaderImageFlatOpc() {
  return jsonToFlatOpc(headerFooterDoc, {
    extraParts: [
      {
        name: '/word/header1.xml',
        contentType: CT.header,
        xml: '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:p><w:r><w:t>Header with image</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="304800"/><wp:docPr id="5" name="Header Image"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="5" name="Header Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdHeaderImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:hdr>',
        rel: {
          id: 'rIdHeader',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
          target: 'header1.xml',
        },
      },
      {
        name: '/word/media/image1.png',
        contentType: 'image/png',
        binary: 'iVBORw0KGgo=',
        rel: {
          source: '/word/header1.xml',
          id: 'rIdHeaderImage',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          target: 'media/image1.png',
        },
      },
    ],
  });
}

function makeHeaderDateFieldFlatOpc() {
  return jsonToFlatOpc(headerFooterDoc, {
    extraParts: [
      {
        name: '/word/header1.xml',
        contentType: CT.header,
        xml: '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE \\\\@ "yyyy-MM-dd" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2026-07-09</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>',
        rel: {
          id: 'rIdHeader',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
          target: 'header1.xml',
        },
      },
    ],
  });
}

function makeMultiSectionHeaderFooterFlatOpc() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/[Content_Types].xml" pkg:contentType="application/vnd.openxmlformats-package.content-types+xml"><pkg:xmlData>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/header1.xml" ContentType="${CT.header}"/>
      <Override PartName="/word/footer1.xml" ContentType="${CT.footer}"/>
      <Override PartName="/word/header2.xml" ContentType="${CT.header}"/>
    </Types>
  </pkg:xmlData></pkg:part>
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdOfficeDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>
  </pkg:xmlData></pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
      <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
      <Relationship Id="rIdHeader2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
    </Relationships>
  </pkg:xmlData></pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body>
        <w:p><w:r><w:t>Section one</w:t></w:r><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="2160" w:left="2520" w:header="720" w:footer="840" w:gutter="120"/><w:pgNumType w:start="3" w:fmt="decimal"/><w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr></w:p>
        <w:p><w:r><w:t>Section two inherits</w:t></w:r><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:cols w:equalWidth="0"><w:col w:w="3000" w:space="360"/><w:col w:w="4200"/></w:cols></w:sectPr></w:pPr></w:p>
        <w:p><w:r><w:t>Section three</w:t></w:r></w:p>
        <w:sectPr><w:headerReference w:type="even" r:id="rIdHeader2"/><w:titlePg/><w:type w:val="oddPage"/></w:sectPr>
      </w:body>
    </w:document>
  </pkg:xmlData></pkg:part>
  <pkg:part pkg:name="/word/header1.xml" pkg:contentType="${CT.header}"><pkg:xmlData><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Section one header</w:t></w:r></w:p></w:hdr></pkg:xmlData></pkg:part>
  <pkg:part pkg:name="/word/footer1.xml" pkg:contentType="${CT.footer}"><pkg:xmlData><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Section one footer</w:t></w:r></w:p></w:ftr></pkg:xmlData></pkg:part>
  <pkg:part pkg:name="/word/header2.xml" pkg:contentType="${CT.header}"><pkg:xmlData><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Section three even header</w:t></w:r></w:p></w:hdr></pkg:xmlData></pkg:part>
</pkg:package>`;
}

let pass = 0;
const ok = (name, fn) => {
  fn();
  pass += 1;
  console.log('  PASS', name);
};

function collectTypes(value, result = []) {
  if (Array.isArray(value)) {
    for (const item of value) collectTypes(item, result);
    return result;
  }
  if (!value || typeof value !== 'object') return result;
  if (value.type) result.push(value.type);
  for (const child of Object.values(value)) collectTypes(child, result);
  return result;
}

ok('plain Flat OPC produces non-empty bodyJson', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeFlatOpc());
  assert.equal(result.bodyJson?.type, 'document');
  assert.ok(Array.isArray(result.bodyJson.content));
  assert.ok(result.bodyJson.content.length > 0);
  assert.deepEqual(result.diagnostics.unsupportedDomains, []);
  assert.equal(result.gaps.comments.present, false);
  assert.equal(result.gaps.headersFooters.count, 0);
  assert.equal(result.gaps.media.status, 'pending');
});

ok('package-only domains are diagnostics, not host-side parsing work', () => {
  const result = flatOpcToOnlyOfficeWriteback(
    makeFlatOpc([
      {
        name: '/word/comments.xml',
        contentType: CT.comments,
        xml: '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
      },
      {
        name: '/word/header1.xml',
        contentType: CT.header,
        xml: '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>',
      },
      {
        name: '/word/media/image1.png',
        contentType: 'image/png',
        binary: 'iVBORw0KGgo=',
      },
    ]),
  );
  assert.deepEqual(
    new Set(result.diagnostics.unsupportedDomains),
    new Set(['comments', 'headersFooters', 'media']),
  );
  assert.match(result.diagnostics.warnings.join('\n'), /gap domains pending rehydration/);
  assert.equal(result.gaps.comments.present, true);
  assert.equal(result.gaps.comments.count, 1);
  assert.equal(result.gaps.comments.parts[0].path, 'word/comments.xml');
  assert.equal(result.gaps.headersFooters.parts[0].contentType, CT.header);
  assert.equal(result.gaps.media.parts[0].path, 'word/media/image1.png');
  assert.equal(result.gaps.notes.present, false);
});

ok('comments gap includes adapter-owned rehydration payload', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeCommentFlatOpc());
  assert.equal(result.gaps.comments.present, true);
  assert.equal(result.gaps.comments.status, 'payload-ready');
  assert.deepEqual(result.gaps.comments.items, [
    {
      id: '7',
      author: 'Coco',
      date: '',
      text: 'Review note',
      anchorText: 'Commented anchor',
      paragraphText: 'Prefix Commented anchor suffix',
    },
  ]);
  assert.ok(!collectTypes(result.bodyJson).includes('commentRangeStart'));
  assert.ok(!collectTypes(result.bodyJson).includes('commentRangeEnd'));
  assert.ok(!collectTypes(result.bodyJson).includes('commentReference'));
});

ok('notes gap includes adapter-owned rehydration payload', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeNotesFlatOpc());
  const footnote = result.gaps.notes.items[0];
  const endnote = result.gaps.notes.items[1];
  assert.equal(result.gaps.notes.present, true);
  assert.equal(result.gaps.notes.status, 'payload-ready');
  assert.equal(result.gaps.notes.items.length, 2);
  assert.deepEqual(
    result.gaps.notes.items.map((item) => ({
      kind: item.kind,
      id: item.id,
      text: item.text,
      anchorText: item.anchorText,
      bodyType: item.bodyJson?.type,
    })),
    [
      {
        kind: 'footnote',
        id: '101',
        text: 'Footnote bodyFootnote second paragraphFootnote table cell',
        anchorText: 'Footnote anchor',
        bodyType: 'document',
      },
      {
        kind: 'endnote',
        id: '202',
        text: 'Endnote bodyEndnote table cell',
        anchorText: 'Footnote anchor and endnote anchor',
        bodyType: 'document',
      },
    ],
  );
  assert.deepEqual(
    footnote.bodyJson.content.map((block) => block.type),
    ['paragraph', 'paragraph', 'paragraph', 'table'],
  );
  assert.equal(footnote.bodyJson.content[0].pPr.jc, 'center');
  assert.equal(footnote.bodyJson.content[0].content[0].rPr.b, true);
  assert.equal(footnote.bodyJson.content[0].content[0].rPr.i, true);
  assert.equal(footnote.bodyJson.content[0].content[0].rPr.sz, 28);
  assert.equal(footnote.bodyJson.content[0].content[0].rPr.rFonts.ascii, 'Calibri Light');
  const footnoteDrawing = footnote.bodyJson.content[2].content[0].content[0];
  assert.equal(footnoteDrawing.type, 'paraDrawing');
  assert.match(footnoteDrawing.graphic.blipFill.rasterImageId, /^data:image\/png;base64,/);
  assert.equal(
    footnote.bodyJson.content[3].content[0].content[0].content.content[0].content[0].content[0],
    'Footnote table cell',
  );
  assert.deepEqual(
    endnote.bodyJson.content.map((block) => block.type),
    ['paragraph', 'table'],
  );
  assert.equal(endnote.bodyJson.content[0].content[0].rPr.b, true);
  assert.equal(
    endnote.bodyJson.content[1].content[0].content[0].content.content[0].content[0].content[0],
    'Endnote table cell',
  );
  assert.ok(!collectTypes(result.bodyJson).includes('footnoteRef'));
  assert.ok(!collectTypes(result.bodyJson).includes('endnoteRef'));
});

ok('hyperlinks round-trip through bodyJson and document relationships', () => {
  const flat = makeHyperlinkFlatOpc();
  assert.match(flat, /<w:hyperlink r:id="rIdCocoHyperlink1" w:tooltip="Coco link tooltip">/);
  assert.match(
    flat,
    /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"/,
  );
  assert.match(flat, /Target="https:\/\/example\.com\/coco-link" TargetMode="External"/);
  const result = flatOpcToOnlyOfficeWriteback(flat);
  const link = result.bodyJson.content[0].content.find((item) => item.type === 'hyperlink');
  assert.equal(link.value, 'https://example.com/coco-link');
  assert.equal(link.tooltip, 'Coco link tooltip');
  assert.equal(link.content[0].content[0], 'Coco link');
});

ok('bookmarks become adapter-owned writeback payloads', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeBookmarkFlatOpc());
  assert.equal(result.gaps.bookmarks.present, true);
  assert.equal(result.gaps.bookmarks.status, 'payload-ready');
  assert.equal(result.gaps.bookmarks.items.length, 1);
  assert.equal(result.gaps.bookmarks.items[0].id, '42');
  assert.equal(result.gaps.bookmarks.items[0].name, 'coco_bookmark_test');
  assert.equal(result.gaps.bookmarks.items[0].anchorText, 'Bookmark anchor');
  assert.equal(
    result.bodyJson.content[0].content.some(
      (item) => item.type === 'bookmarkStart' || item.type === 'bookmarkEnd',
    ),
    false,
  );
});

ok('body images are inlined into bodyJson for FromJSON replacement', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeBodyImageFlatOpc());
  const drawing = result.bodyJson.content[0].content[1].content[0];
  assert.equal(drawing.type, 'paraDrawing');
  assert.equal(drawing.docPr.name, 'Picture');
  assert.deepEqual(drawing.effectExtent, { l: 0, t: 0, r: 0, b: 0 });
  assert.equal(drawing.extent.cx, 914400);
  assert.equal(drawing.extent.cy, 304800);
  assert.match(drawing.graphic.blipFill.rasterImageId, /^data:image\/png;base64,/);
  assert.ok(!result.diagnostics.unsupportedDomains.includes('media'));
  assert.ok(!result.diagnostics.unsupportedDomains.includes('drawings'));
  assert.deepEqual(result.diagnostics.unsupportedItems, []);
  assert.equal(result.gaps.media.status, 'payload-ready');
  assert.equal(result.gaps.media.unsupportedCount, 0);
  assert.equal(result.gaps.media.items[0].status, 'inlined');
  assert.equal(result.gaps.drawings.status, 'payload-ready');
  assert.equal(result.gaps.drawings.unsupportedCount, 0);
  assert.equal(result.gaps.drawings.items[0].status, 'inlined');
  assert.equal(result.gaps.drawings.items[0].drawingType, 'inline');
});

ok('floating body images stay diagnostics-only instead of being coerced to inline FromJSON', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeAnchoredBodyImageFlatOpc());
  assert.ok(!collectTypes(result.bodyJson).includes('paraDrawing'));
  assert.ok(result.diagnostics.unsupportedDomains.includes('media'));
  assert.ok(result.diagnostics.unsupportedDomains.includes('drawings'));
  assert.deepEqual(
    result.diagnostics.unsupportedItems.map((item) => ({
      domain: item.domain,
      partPath: item.partPath,
      reason: item.reason,
      drawingType: item.drawingType,
    })),
    [
      {
        domain: 'drawings',
        partPath: 'word/document.xml',
        reason: 'floating-anchor',
        drawingType: 'anchor',
      },
      {
        domain: 'media',
        partPath: 'word/media/image1.png',
        reason: 'referenced-by-unsupported-drawing',
        drawingType: undefined,
      },
    ],
  );
  assert.match(result.diagnostics.warnings.join('\n'), /drawings:floating-anchor=1/);
  assert.equal(result.gaps.drawings.status, 'pending');
  assert.equal(result.gaps.drawings.unsupportedCount, 1);
  assert.equal(result.gaps.drawings.items[0].drawingType, 'anchor');
  assert.equal(result.gaps.drawings.items[0].unsupportedReason, 'floating-anchor');
  assert.equal(result.gaps.media.status, 'pending');
  assert.equal(result.gaps.media.unsupportedCount, 1);
  assert.equal(result.gaps.media.items[0].status, 'referenced-by-unsupported-drawing');
});

ok('chart drawings are surfaced with a chart-specific unsupported reason', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeChartDrawingFlatOpc());
  const item = result.diagnostics.unsupportedItems.find(
    (entry) => entry.domain === 'drawings' && entry.reason === 'chart-drawing',
  );
  assert.ok(item);
  assert.equal(item.partPath, 'word/document.xml');
  assert.equal(item.drawingType, 'inline');
  assert.match(item.graphicDataUri, /\/chart$/);
  assert.match(result.diagnostics.warnings.join('\n'), /drawings:chart-drawing=1/);
});

ok('text box drawings are surfaced with a text-box-specific unsupported reason', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeTextBoxDrawingFlatOpc());
  const item = result.diagnostics.unsupportedItems.find(
    (entry) => entry.domain === 'drawings' && entry.reason === 'text-box-drawing',
  );
  assert.ok(item);
  assert.equal(item.partPath, 'word/document.xml');
  assert.equal(item.drawingType, 'inline');
  assert.match(result.diagnostics.warnings.join('\n'), /drawings:text-box-drawing=1/);
});

ok('headersFooters gap includes adapter-owned text payload', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeHeaderFooterFlatOpc());
  assert.equal(result.gaps.headersFooters.present, true);
  assert.equal(result.gaps.headersFooters.status, 'payload-ready');
  assert.deepEqual(result.gaps.headersFooters.items, [
    {
      kind: 'header',
      type: 'default',
      partPath: 'word/header1.xml',
      text: 'Header text',
    },
    {
      kind: 'footer',
      type: 'even',
      partPath: 'word/footer1.xml',
      text: 'Footer text',
    },
  ]);
  assert.equal(result.gaps.headersFooters.sections.length, 1);
  assert.equal(result.gaps.headersFooters.sections[0].startBlockIndex, 0);
  assert.equal(result.gaps.headersFooters.sections[0].endBlockIndex, 0);
  assert.equal(result.gaps.headersFooters.sections[0].boundary, 'body');
  assert.equal(result.gaps.headersFooters.sections[0].headers.default.state, 'own');
  assert.equal(result.gaps.headersFooters.sections[0].headers.default.text, 'Header text');
  assert.equal(result.gaps.headersFooters.sections[0].headers.default.bodyJson.type, 'document');
  assert.equal(
    result.gaps.headersFooters.sections[0].headers.default.bodyJson.content[0].type,
    'paragraph',
  );
  assert.equal(result.gaps.headersFooters.sections[0].headers.first.state, 'empty');
  assert.equal(result.gaps.headersFooters.sections[0].footers.even.state, 'own');
  assert.equal(result.gaps.headersFooters.sections[0].footers.default.state, 'empty');
});

ok('headersFooters sections include structured bodyJson for table content', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeHeaderTableFlatOpc());
  const header = result.gaps.headersFooters.sections[0].headers.default;
  assert.equal(header.state, 'own');
  assert.equal(header.bodyJson.content.length, 2);
  assert.equal(header.bodyJson.content[0].type, 'paragraph');
  assert.equal(header.bodyJson.content[1].type, 'table');
  assert.equal(header.text, 'Header before tableHeader table cell');
});

ok('headersFooters sections inline image payloads into bodyJson', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeHeaderImageFlatOpc());
  const header = result.gaps.headersFooters.sections[0].headers.default;
  const drawing = header.bodyJson.content[0].content[1].content[0];
  assert.equal(header.state, 'own');
  assert.equal(header.text, 'Header with image');
  assert.equal(drawing.type, 'paraDrawing');
  assert.deepEqual(drawing.extent, { cx: 914400, cy: 304800 });
  assert.match(drawing.graphic.blipFill.rasterImageId, /^data:image\/png;base64,/);
});

ok('headersFooters sections preserve inherited and empty semantics', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeMultiSectionHeaderFooterFlatOpc());
  const sections = result.gaps.headersFooters.sections;
  assert.equal(sections.length, 3);
  assert.equal(sections[0].type, 'nextPage');
  assert.equal(sections[0].startBlockIndex, 0);
  assert.equal(sections[0].endBlockIndex, 0);
  assert.equal(sections[0].boundary, 'paragraph');
  assert.equal(sections[0].headers.default.state, 'own');
  assert.equal(sections[0].headers.default.text, 'Section one header');
  assert.equal(sections[0].footers.default.state, 'own');
  assert.deepEqual(sections[0].pageSize, { width: 12240, height: 15840, orient: '' });
  assert.deepEqual(sections[0].pageMargins, {
    top: 1440,
    right: 1800,
    bottom: 2160,
    left: 2520,
    header: 720,
    footer: 840,
    gutter: 120,
  });
  assert.deepEqual(sections[0].pageNumbering, {
    start: 3,
    fmt: 'decimal',
    chapStyle: '',
    chapSep: '',
  });
  assert.equal(sections[0].columns.count, 2);
  assert.equal(sections[0].columns.space, 720);
  assert.equal(sections[1].type, 'continuous');
  assert.equal(sections[1].startBlockIndex, 1);
  assert.equal(sections[1].endBlockIndex, 1);
  assert.equal(sections[1].boundary, 'paragraph');
  assert.equal(sections[1].headers.default.state, 'inherited');
  assert.equal(sections[1].headers.default.inheritedFrom, 0);
  assert.equal(sections[1].footers.default.state, 'inherited');
  assert.deepEqual(sections[1].pageSize, { width: 15840, height: 12240, orient: 'landscape' });
  assert.equal(sections[1].columns.equalWidth, false);
  assert.deepEqual(sections[1].columns.col, [
    { width: 3000, space: 360 },
    { width: 4200, space: undefined },
  ]);
  assert.equal(sections[2].type, 'oddPage');
  assert.equal(sections[2].startBlockIndex, 2);
  assert.equal(sections[2].endBlockIndex, 2);
  assert.equal(sections[2].boundary, 'body');
  assert.equal(sections[2].titlePage, true);
  assert.equal(sections[2].headers.default.state, 'inherited');
  assert.equal(sections[2].headers.even.state, 'own');
  assert.equal(sections[2].headers.even.text, 'Section three even header');
  assert.equal(sections[2].footers.even.state, 'inherited');
});

ok('headersFooters page numbering format is surfaced as unsupported diagnostics', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeMultiSectionHeaderFooterFlatOpc());
  const item = result.diagnostics.unsupportedItems.find(
    (entry) => entry.domain === 'headersFooters' && entry.reason === 'page-numbering-format',
  );
  assert.ok(item);
  assert.equal(item.index, 0);
  assert.equal(item.property, 'pageNumbering');
  assert.deepEqual(item.value, { fmt: 'decimal', chapStyle: '', chapSep: '' });
  assert.match(result.diagnostics.warnings.join('\n'), /headersFooters:page-numbering-format=1/);
});

ok('headersFooters dynamic date fields are surfaced as unsupported diagnostics', () => {
  const result = flatOpcToOnlyOfficeWriteback(makeHeaderDateFieldFlatOpc());
  const item = result.diagnostics.unsupportedItems.find(
    (entry) => entry.domain === 'headersFooters' && entry.reason === 'dynamic-date-field',
  );
  assert.ok(item);
  assert.equal(item.partPath, 'word/header1.xml');
  assert.equal(item.fieldType, 'DATE');
  assert.match(item.instruction, /^DATE/);
  assert.match(result.diagnostics.warnings.join('\n'), /headersFooters:dynamic-date-field=1/);
});

ok('failOnUnsupported turns diagnostics into an explicit failure', () => {
  assert.throws(
    () =>
      flatOpcToOnlyOfficeWriteback(
        makeFlatOpc([
          {
            name: '/word/comments.xml',
            contentType: CT.comments,
            xml: '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
          },
        ]),
        { failOnUnsupported: true },
      ),
    /unsupported gap domains: comments/,
  );
});

ok('invalid Flat OPC fails before host execution', () => {
  assert.throws(
    () => flatOpcToOnlyOfficeWriteback('not flat opc'),
    /PACKAGE_CORRUPT|fromFlatOpc|XML|Parsing document failed/i,
  );
});

console.log(`writeback adapter tests: ${pass} passed`);
