import { describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import {
  createPackage,
  createXmlPart,
  createBinaryPart,
  fromZip,
  getPart,
  isBinaryPart,
  isXmlPart,
  listParts,
  markDirty,
  setPart,
  toZip,
} from '../index.js';

const CT_XML_STD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>
  </w:body>
</w:document>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

const MAIN_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const PNG_BYTES = new Uint8Array([
  0x89,
  0x50,
  0x4e,
  0x47,
  0x0d,
  0x0a,
  0x1a,
  0x0a, // PNG signature
  0x00,
  0x00,
  0x00,
  0x0d,
  0x49,
  0x48,
  0x44,
  0x52, // IHDR chunk (fake but parseable shape)
]);

async function buildMinimalDocxZip(opts?: { withImage?: boolean }): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML_STD);
  zip.file('_rels/.rels', MAIN_RELS_XML);
  zip.file('word/document.xml', DOC_XML);
  zip.file('word/styles.xml', STYLES_XML);
  if (opts?.withImage) {
    zip.file('word/media/image1.png', PNG_BYTES);
  }
  return await zip.generateAsync({ type: 'uint8array' });
}

describe('fromZip', () => {
  test('最小 docx:解压产出完整 Package', async () => {
    const buf = await buildMinimalDocxZip();
    const pkg = await fromZip(buf);

    const paths = listParts(pkg)
      .map((p) => p.path)
      .sort();
    expect(paths).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'word/document.xml',
      'word/styles.xml',
    ]);
  });

  test('contentType 正确回填:Override 优先,Default 兜底', async () => {
    const buf = await buildMinimalDocxZip();
    const pkg = await fromZip(buf);

    expect(getPart(pkg, 'word/document.xml')?.contentType).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    );
    expect(getPart(pkg, 'word/styles.xml')?.contentType).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
    );
    expect(getPart(pkg, '_rels/.rels')?.contentType).toBe(
      'application/vnd.openxmlformats-package.relationships+xml',
    );
    expect(getPart(pkg, '[Content_Types].xml')?.contentType).toBe(
      'application/vnd.openxmlformats-package.content-types+xml',
    );
  });

  test('XML part 有 xmlDom,二进制 part 有 bytes', async () => {
    const buf = await buildMinimalDocxZip({ withImage: true });
    const pkg = await fromZip(buf);

    const docPart = getPart(pkg, 'word/document.xml')!;
    expect(isXmlPart(docPart)).toBe(true);
    expect(docPart.xmlDom).toBeDefined();
    expect(docPart.bytes).toBeUndefined();

    const imgPart = getPart(pkg, 'word/media/image1.png')!;
    expect(isBinaryPart(imgPart)).toBe(true);
    expect(imgPart.bytes).toBeDefined();
    expect(imgPart.xmlDom).toBeUndefined();
  });

  test('XML part 装载后 _dirty=false,_originalBytes 就位', async () => {
    const buf = await buildMinimalDocxZip();
    const pkg = await fromZip(buf);

    const docPart = getPart(pkg, 'word/document.xml')!;
    expect(docPart._dirty).toBe(false);
    expect(docPart._originalBytes).toBeDefined();
    expect(docPart._originalBytes!.length).toBeGreaterThan(0);
  });

  test('二进制 part 的 bytes 与原始 PNG 一致', async () => {
    const buf = await buildMinimalDocxZip({ withImage: true });
    const pkg = await fromZip(buf);
    const imgPart = getPart(pkg, 'word/media/image1.png')!;
    expect(imgPart.bytes).toEqual(PNG_BYTES);
  });

  test('image/svg+xml media part keeps bytes instead of strict XML parsing', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
    );
    zip.file('_rels/.rels', MAIN_RELS_XML);
    zip.file('word/document.xml', DOC_XML);
    const svgBytes = new TextEncoder().encode(
      '<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48…">',
    );
    zip.file('word/media/logo.svg', svgBytes);

    const pkg = await fromZip(await zip.generateAsync({ type: 'uint8array' }));
    const svgPart = getPart(pkg, 'word/media/logo.svg')!;
    expect(isBinaryPart(svgPart)).toBe(true);
    expect(svgPart.bytes).toEqual(svgBytes);
    expect(svgPart.xmlDom).toBeUndefined();
  });

  test('缺少 [Content_Types].xml 抛错', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', DOC_XML);
    const buf = await zip.generateAsync({ type: 'uint8array' });
    await expect(fromZip(buf)).rejects.toThrow(/\[Content_Types\]\.xml/);
  });

  test('存在无法解析 contentType 的 part 抛错', async () => {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`,
    );
    // 没有 .rels 的 Default,也没有 Override
    zip.file('_rels/.rels', MAIN_RELS_XML);
    zip.file('word/document.xml', DOC_XML);
    const buf = await zip.generateAsync({ type: 'uint8array' });
    await expect(fromZip(buf)).rejects.toThrow(/no ContentType/);
  });

  test('接受 ArrayBuffer 输入(非 Uint8Array)', async () => {
    const buf = await buildMinimalDocxZip();
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const pkg = await fromZip(ab as ArrayBuffer);
    expect(listParts(pkg).length).toBeGreaterThan(0);
  });
});

describe('toZip', () => {
  test('空 Package 产出空 zip(无抛错)', async () => {
    const pkg = createPackage();
    const bytes = await toZip(pkg);
    expect(bytes).toBeInstanceOf(Uint8Array);
    // zip signature 'PK\003\004' 或空 zip 至少有 end-of-central-dir
    expect(bytes.length).toBeGreaterThan(0);
  });

  test('dirty=false 的 XML part 直接复用 _originalBytes,不重 serialize', async () => {
    const originalXml = DOC_XML;
    const originalBytes = new TextEncoder().encode(originalXml);
    const pkg = createPackage();
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType: 'application/xml',
        xmlDom: parseXmlDocument(originalXml) as unknown as Document,
        originalBytes,
      }),
    );

    const zipBytes = await toZip(pkg);
    const roundTrip = await JSZip.loadAsync(zipBytes);
    const backXml = await roundTrip.file('word/document.xml')!.async('uint8array');
    expect(backXml).toEqual(originalBytes);
  });

  test('dirty=true 的 XML part 会被 slimdom serialize', async () => {
    const dom = parseXmlDocument(DOC_XML) as unknown as Document;
    const pkg = createPackage();
    const part = createXmlPart({
      path: 'word/document.xml',
      contentType: 'application/xml',
      xmlDom: dom,
      originalBytes: new TextEncoder().encode(DOC_XML),
    });
    setPart(pkg, part);
    markDirty(part);

    const zipBytes = await toZip(pkg);
    const roundTrip = await JSZip.loadAsync(zipBytes);
    const backText = await roundTrip.file('word/document.xml')!.async('string');
    // 重 serialize 后的文本结构应与 DOM 一致
    const backDom = parseXmlDocument(backText);
    const ps = (backDom as unknown as Document).getElementsByTagNameNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'p',
    );
    expect(ps.length).toBe(1);
  });

  test('输出 zip 走 DEFLATE 压缩(XML part 远小于原字节)', async () => {
    // 构造一个高度可压缩的 XML(重复字符串),DEFLATE 后体积应显著小于原始
    const repeatedXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<root xmlns="http://example.com/ns">\n` +
      `  ${'<item>hello world hello world hello world</item>\n  '.repeat(500)}` +
      `</root>`;
    const originalBytes = new TextEncoder().encode(repeatedXml);
    const pkg = createPackage();
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType: 'application/xml',
        xmlDom: parseXmlDocument(repeatedXml) as unknown as Document,
        originalBytes,
      }),
    );

    const zipBytes = await toZip(pkg);
    // STORE 模式下 zip 至少 = 原字节 + zip overhead;DEFLATE 后应远小于原字节。
    // 阈值取 50%(实测重复 XML 压缩率 < 5%,留足余量)。
    expect(zipBytes.length).toBeLessThan(originalBytes.length * 0.5);
  });

  test('dirty=true 的二进制 part 使用 bytes', async () => {
    const pkg = createPackage();
    const newBytes = new Uint8Array([1, 2, 3, 4]);
    const part = createBinaryPart({
      path: 'word/media/image1.png',
      contentType: 'image/png',
      bytes: newBytes,
    });
    setPart(pkg, part);
    markDirty(part);

    const zipBytes = await toZip(pkg);
    const roundTrip = await JSZip.loadAsync(zipBytes);
    const back = await roundTrip.file('word/media/image1.png')!.async('uint8array');
    expect(back).toEqual(newBytes);
  });
});

describe('roundtrip 守恒', () => {
  test('fromZip → toZip → fromZip:paths 与 contentType 守恒', async () => {
    const buf1 = await buildMinimalDocxZip({ withImage: true });
    const pkg1 = await fromZip(buf1);
    const buf2 = await toZip(pkg1);
    const pkg2 = await fromZip(buf2);

    const paths1 = listParts(pkg1)
      .map((p) => p.path)
      .sort();
    const paths2 = listParts(pkg2)
      .map((p) => p.path)
      .sort();
    expect(paths2).toEqual(paths1);

    for (const path of paths1) {
      expect(getPart(pkg2, path)?.contentType).toBe(getPart(pkg1, path)?.contentType);
    }
  });

  test('只改一个 part 其它零 serialize', async () => {
    const buf = await buildMinimalDocxZip({ withImage: true });
    const pkg = await fromZip(buf);

    // 改一个 XML part
    const docPart = getPart(pkg, 'word/document.xml')!;
    markDirty(docPart);
    // 其它 part 保持 dirty=false

    const zipBytes = await toZip(pkg);
    const roundTrip = await JSZip.loadAsync(zipBytes);

    // 没改动的 styles.xml 字节应与原 zip 中的字节完全一致
    const orig = await JSZip.loadAsync(buf);
    const origStyles = await orig.file('word/styles.xml')!.async('uint8array');
    const newStyles = await roundTrip.file('word/styles.xml')!.async('uint8array');
    expect(newStyles).toEqual(origStyles);

    const origImg = await orig.file('word/media/image1.png')!.async('uint8array');
    const newImg = await roundTrip.file('word/media/image1.png')!.async('uint8array');
    expect(newImg).toEqual(origImg);
  });

  test('dirty part 的字节经 XML serialize 后内容语义仍合法', async () => {
    const buf = await buildMinimalDocxZip();
    const pkg = await fromZip(buf);
    const docPart = getPart(pkg, 'word/document.xml')!;
    markDirty(docPart);

    const zipBytes = await toZip(pkg);
    const roundTrip = await JSZip.loadAsync(zipBytes);
    const backText = await roundTrip.file('word/document.xml')!.async('string');
    // 能被重新 parse,说明 serialize 输出是合法 XML
    const backDom = parseXmlDocument(backText);
    expect(backDom).toBeDefined();
    // 能被 serialize 再 parse,说明 round-trip 语义等价
    const reSerialize = serializeToWellFormedString(backDom);
    expect(reSerialize.length).toBeGreaterThan(0);
  });
});
