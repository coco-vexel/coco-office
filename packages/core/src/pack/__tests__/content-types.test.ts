import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import {
  createPackage,
  createBinaryPart,
  createXmlPart,
  ensureContentTypes,
  getPart,
  setPart,
} from '../index.js';

const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

describe('ensureContentTypes', () => {
  test('空包 → 合成一个空 <Types/>', () => {
    const pkg = createPackage();
    const changed = ensureContentTypes(pkg);
    expect(changed).toBe(true);
    const ct = getPart(pkg, '[Content_Types].xml');
    expect(ct).toBeDefined();
    expect(ct!.contentType).toBe('application/vnd.openxmlformats-package.content-types+xml');
  });

  test('已有 CT part → 不改动,返回 false', () => {
    const pkg = createPackage();
    setPart(
      pkg,
      createXmlPart({
        path: '[Content_Types].xml',
        contentType: 'application/vnd.openxmlformats-package.content-types+xml',
        xmlDom: parseXmlDocument(
          `<Types xmlns="${CT_NS}"><Override PartName="/x" ContentType="application/xml"/></Types>`,
        ) as unknown as Document,
      }),
    );
    const originalCt = getPart(pkg, '[Content_Types].xml');
    const changed = ensureContentTypes(pkg);
    expect(changed).toBe(false);
    // 仍是原对象引用
    expect(getPart(pkg, '[Content_Types].xml')).toBe(originalCt);
  });

  test('每个非 CT part 产生一条 Override', () => {
    const pkg = createPackage();
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType: DOC_CTYPE,
        xmlDom: parseXmlDocument('<r/>') as unknown as Document,
      }),
    );
    setPart(
      pkg,
      createBinaryPart({
        path: 'word/media/img.png',
        contentType: 'image/png',
        bytes: new Uint8Array([1]),
      }),
    );

    ensureContentTypes(pkg);
    const ct = getPart(pkg, '[Content_Types].xml')!;
    const overrides = ct.xmlDom!.getElementsByTagNameNS(CT_NS, 'Override');
    const pairs = Array.from({ length: overrides.length }, (_, i) => {
      const el = overrides[i]!;
      return [el.getAttribute('PartName'), el.getAttribute('ContentType')];
    });
    expect(pairs).toEqual(
      expect.arrayContaining([
        ['/word/document.xml', DOC_CTYPE],
        ['/word/media/img.png', 'image/png'],
      ]),
    );
    expect(pairs).toHaveLength(2);
  });

  test('合成后 CT part 的 _dirty=false(通过原始字节路径,零 serialize)', () => {
    const pkg = createPackage();
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType: DOC_CTYPE,
        xmlDom: parseXmlDocument('<r/>') as unknown as Document,
      }),
    );
    ensureContentTypes(pkg);
    const ct = getPart(pkg, '[Content_Types].xml')!;
    expect(ct._dirty).toBe(false);
    expect(ct._originalBytes).toBeDefined();
  });
});
