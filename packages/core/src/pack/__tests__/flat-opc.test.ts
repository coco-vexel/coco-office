import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import {
  createBinaryPart,
  createPackage,
  createXmlPart,
  fromFlatOpc,
  getPart,
  isBinaryPart,
  isXmlPart,
  listParts,
  markDirty,
  setPart,
  toFlatOpc,
} from '../index.js';

const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

/** 构建最小 Flat OPC fixture(单 XML part) */
function minimalFlatOpc(): string {
  return `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="${PKG_NS}">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CTYPE}">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

/** 含 XML + 二进制 part 的 Flat OPC fixture */
function flatOpcWithImage(): string {
  // 简短 PNG base64(假 PNG header + IHDR)
  const pngB64 = 'iVBORw0KGgoAAAANSUhEUg==';
  return `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="${PKG_NS}">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CTYPE}">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>doc</w:t></w:r></w:p></w:body></w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png" pkg:compression="store">
    <pkg:binaryData>${pngB64}</pkg:binaryData>
  </pkg:part>
</pkg:package>`;
}

describe('fromFlatOpc', () => {
  test('最小 Flat OPC → 解包为 1 个 XML part', () => {
    const pkg = fromFlatOpc(minimalFlatOpc());
    const parts = listParts(pkg);
    expect(parts).toHaveLength(1);
    expect(parts[0]?.path).toBe('word/document.xml');
    expect(parts[0]?.contentType).toBe(DOC_CTYPE);
    expect(isXmlPart(parts[0]!)).toBe(true);
  });

  test('pkg:name 的前导 / 被剥掉', () => {
    const pkg = fromFlatOpc(minimalFlatOpc());
    expect(getPart(pkg, 'word/document.xml')).toBeDefined();
    expect(getPart(pkg, '/word/document.xml')).toBeUndefined();
  });

  test('[Content_Types].xml 不作为 part 产生(Flat OPC 不含 CT)', () => {
    const pkg = fromFlatOpc(minimalFlatOpc());
    expect(getPart(pkg, '[Content_Types].xml')).toBeUndefined();
  });

  test('XML part 的 xmlDom 可用,文本内容保留', () => {
    const pkg = fromFlatOpc(minimalFlatOpc());
    const part = getPart(pkg, 'word/document.xml');
    expect(part?.xmlDom).toBeDefined();
    const ts = part!.xmlDom!.getElementsByTagNameNS(W_NS, 't');
    expect(ts.length).toBe(2);
    expect(ts[0]?.textContent).toBe('Hello ');
    expect(ts[1]?.textContent).toBe('World');
  });

  test('XML part 装载后 _dirty=false,_originalBytes 就位', () => {
    const pkg = fromFlatOpc(minimalFlatOpc());
    const part = getPart(pkg, 'word/document.xml')!;
    expect(part._dirty).toBe(false);
    expect(part._originalBytes).toBeDefined();
    expect(part._originalBytes!.length).toBeGreaterThan(0);
  });

  test('二进制 part:base64 解码产出正确字节', () => {
    const pkg = fromFlatOpc(flatOpcWithImage());
    const img = getPart(pkg, 'word/media/image1.png')!;
    expect(isBinaryPart(img)).toBe(true);
    // 'iVBORw0KGgoAAAANSUhEUg==' decoded bytes:PNG 签名开头 + IHDR 部分
    expect(img.bytes?.[0]).toBe(0x89); // PNG 魔数首字节
    expect(img.bytes?.[1]).toBe(0x50); // 'P'
    expect(img.bytes?.[2]).toBe(0x4e); // 'N'
    expect(img.bytes?.[3]).toBe(0x47); // 'G'
  });

  test('二进制 part 的 base64 含空白 / 换行时仍可解', () => {
    const xml = `<?xml version="1.0" standalone="yes"?>
<pkg:package xmlns:pkg="${PKG_NS}">
  <pkg:part pkg:name="/data/blob.bin" pkg:contentType="application/octet-stream">
    <pkg:binaryData>
      AAEC
      AwQF
    </pkg:binaryData>
  </pkg:part>
</pkg:package>`;
    const pkg = fromFlatOpc(xml);
    const part = getPart(pkg, 'data/blob.bin')!;
    expect(Array.from(part.bytes!)).toEqual([0, 1, 2, 3, 4, 5]);
  });

  test('缺 xmlData 的 XML contentType part 抛错', () => {
    const xml = `<?xml version="1.0" standalone="yes"?>
<pkg:package xmlns:pkg="${PKG_NS}">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CTYPE}"/>
</pkg:package>`;
    expect(() => fromFlatOpc(xml)).toThrow(/no <pkg:xmlData>/);
  });

  test('pkg:part 缺 name / contentType 抛错', () => {
    const xml = `<?xml version="1.0" standalone="yes"?>
<pkg:package xmlns:pkg="${PKG_NS}">
  <pkg:part pkg:contentType="application/xml"><pkg:xmlData><r/></pkg:xmlData></pkg:part>
</pkg:package>`;
    expect(() => fromFlatOpc(xml)).toThrow(/missing name or contentType/);
  });
});

describe('toFlatOpc', () => {
  test('空包产出顶部 PI + 空 pkg:package', () => {
    const pkg = createPackage();
    const out = toFlatOpc(pkg);
    expect(out).toMatch(/<\?mso-application progid="Word\.Document"\?>/);
    expect(out).toMatch(/<pkg:package/);
    // 无 <pkg:part>
    expect(out).not.toMatch(/<pkg:part\b/);
  });

  test('XML part 正确嵌入 xmlData,命名空间保留', () => {
    const pkg = createPackage();
    const xml = `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`;
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType: DOC_CTYPE,
        xmlDom: parseXmlDocument(xml) as unknown as Document,
      }),
    );

    const out = toFlatOpc(pkg);
    expect(out).toMatch(/<pkg:part pkg:name="\/word\/document\.xml"/);
    expect(out).toMatch(new RegExp(`pkg:contentType="${DOC_CTYPE.replace(/[.+]/g, '\\$&')}"`));
    expect(out).toMatch(/<pkg:xmlData>/);
    expect(out).toMatch(/<w:document/);
  });

  test('二进制 part 产出 pkg:binaryData 带 base64 + pkg:compression="store"', () => {
    const pkg = createPackage();
    setPart(
      pkg,
      createBinaryPart({
        path: 'word/media/image1.png',
        contentType: 'image/png',
        bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
      }),
    );

    const out = toFlatOpc(pkg);
    expect(out).toMatch(/<pkg:part[^>]*pkg:compression="store"/);
    expect(out).toMatch(/<pkg:binaryData>iVBORw==<\/pkg:binaryData>/);
  });

  test('[Content_Types].xml part 被跳过(Flat OPC 不包含)', () => {
    const pkg = createPackage();
    setPart(
      pkg,
      createXmlPart({
        path: '[Content_Types].xml',
        contentType: 'application/vnd.openxmlformats-package.content-types+xml',
        xmlDom: parseXmlDocument('<Types xmlns="http://x"/>') as unknown as Document,
      }),
    );
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType: DOC_CTYPE,
        xmlDom: parseXmlDocument(`<w:document xmlns:w="${W_NS}"/>`) as unknown as Document,
      }),
    );
    const out = toFlatOpc(pkg);
    expect(out).not.toMatch(/\[Content_Types\]\.xml/);
    expect(out).toMatch(/word\/document\.xml/);
  });

  test('path 含特殊字符(空格等)正确转义', () => {
    const pkg = createPackage();
    setPart(
      pkg,
      createBinaryPart({
        path: 'word/media/my photo.png',
        contentType: 'image/png',
        bytes: new Uint8Array([1, 2, 3]),
      }),
    );
    const out = toFlatOpc(pkg);
    expect(out).toMatch(/pkg:name="\/word\/media\/my photo\.png"/);
    // 输出应可被 slimdom 再解析
    expect(() => parseXmlDocument(out)).not.toThrow();
  });
});

describe('roundtrip 守恒', () => {
  test('fromFlatOpc → toFlatOpc → fromFlatOpc:paths / contentType 守恒', () => {
    const xml1 = flatOpcWithImage();
    const pkg1 = fromFlatOpc(xml1);
    const xml2 = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(xml2);

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

  test('roundtrip 后 XML part 文本内容守恒(<w:t> / 属性 xml:space 保留)', () => {
    const xml1 = minimalFlatOpc();
    const pkg1 = fromFlatOpc(xml1);
    const xml2 = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(xml2);

    const docPart = getPart(pkg2, 'word/document.xml')!;
    const ts = docPart.xmlDom!.getElementsByTagNameNS(W_NS, 't');
    expect(ts.length).toBe(2);
    expect(ts[0]?.textContent).toBe('Hello ');
    expect(ts[1]?.textContent).toBe('World');
    // xml:space 保留
    expect(ts[0]?.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space')).toBe('preserve');
  });

  test('roundtrip 后二进制 part 字节完全一致', () => {
    const xml1 = flatOpcWithImage();
    const pkg1 = fromFlatOpc(xml1);
    const bytes1 = getPart(pkg1, 'word/media/image1.png')!.bytes!;
    const xml2 = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(xml2);
    const bytes2 = getPart(pkg2, 'word/media/image1.png')!.bytes!;
    expect(Array.from(bytes2)).toEqual(Array.from(bytes1));
  });

  test('顶部 mso-application PI 在输出中存在', () => {
    const xml1 = minimalFlatOpc();
    const pkg = fromFlatOpc(xml1);
    const xml2 = toFlatOpc(pkg);
    expect(xml2).toMatch(/<\?mso-application progid="Word\.Document"\?>/);
  });

  test('改过的 XML part(dirty)重新 serialize 也能 roundtrip', () => {
    const pkg = fromFlatOpc(minimalFlatOpc());
    const part = getPart(pkg, 'word/document.xml')!;
    markDirty(part);
    const xml = toFlatOpc(pkg);
    const pkg2 = fromFlatOpc(xml);
    expect(getPart(pkg2, 'word/document.xml')).toBeDefined();
    const ts = getPart(pkg2, 'word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 't');
    expect(ts.length).toBe(2);
  });
});
