import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import {
  createBinaryPart,
  createXmlPart,
  isBinaryPart,
  isDirty,
  isXmlPart,
  markDirty,
} from '../index.js';

function makeXmlDom(): Document {
  // slimdom 的 parseXmlDocument 实现 DOM Level 3,返回的 Document 与 lib.dom.Document 结构兼容
  return parseXmlDocument('<root xml:space="preserve"><child/></root>') as unknown as Document;
}

describe('createXmlPart', () => {
  test('originalBytes 传入 → _dirty 初始为 false(装载路径)', () => {
    const xmlDom = makeXmlDom();
    const bytes = new Uint8Array([0x3c, 0x72, 0x6f, 0x6f, 0x74, 0x2f, 0x3e]); // '<root/>'
    const part = createXmlPart({
      path: 'word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      xmlDom,
      originalBytes: bytes,
    });
    expect(part.path).toBe('word/document.xml');
    expect(part.xmlDom).toBe(xmlDom);
    expect(part._originalBytes).toBe(bytes);
    expect(part._dirty).toBe(false);
  });

  test('originalBytes 省略 → _dirty 初始为 true(新建路径)', () => {
    const part = createXmlPart({
      path: 'word/new-part.xml',
      contentType: 'application/xml',
      xmlDom: makeXmlDom(),
    });
    expect(part._originalBytes).toBeUndefined();
    expect(part._dirty).toBe(true);
  });

  test('path 带前导斜杠抛错', () => {
    expect(() =>
      createXmlPart({
        path: '/word/document.xml',
        contentType: 'application/xml',
        xmlDom: makeXmlDom(),
      }),
    ).toThrow(/must not start with '\/'/);
  });

  test('path 为空抛错', () => {
    expect(() =>
      createXmlPart({
        path: '',
        contentType: 'application/xml',
        xmlDom: makeXmlDom(),
      }),
    ).toThrow(/non-empty string/);
  });

  test('contentType 为空抛错', () => {
    expect(() =>
      createXmlPart({
        path: 'word/document.xml',
        contentType: '',
        xmlDom: makeXmlDom(),
      }),
    ).toThrow(/contentType/);
  });

  test('xmlDom 缺失抛错', () => {
    expect(() =>
      createXmlPart({
        path: 'word/document.xml',
        contentType: 'application/xml',
        // @ts-expect-error 故意制造类型违规
        xmlDom: undefined,
      }),
    ).toThrow(/xmlDom required/);
  });
});

describe('createBinaryPart', () => {
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

  test('bytes 传入 → originalBytes 默认等于 bytes,_dirty 恒为 false', () => {
    const part = createBinaryPart({
      path: 'word/media/image1.png',
      contentType: 'image/png',
      bytes: png,
    });
    expect(part.path).toBe('word/media/image1.png');
    expect(part.bytes).toBe(png);
    expect(part._originalBytes).toBe(png);
    expect(part._dirty).toBe(false);
  });

  test('显式 originalBytes 不等于 bytes 时,_originalBytes 保留指定值', () => {
    const original = new Uint8Array([1, 2, 3]);
    const current = new Uint8Array([4, 5, 6]);
    const part = createBinaryPart({
      path: 'word/media/img.png',
      contentType: 'image/png',
      bytes: current,
      originalBytes: original,
    });
    expect(part.bytes).toBe(current);
    expect(part._originalBytes).toBe(original);
    expect(part._dirty).toBe(false);
  });

  test('path 带前导斜杠抛错', () => {
    expect(() =>
      createBinaryPart({
        path: '/word/media/img.png',
        contentType: 'image/png',
        bytes: png,
      }),
    ).toThrow(/must not start with/);
  });

  test('bytes 缺失抛错', () => {
    expect(() =>
      createBinaryPart({
        path: 'word/media/img.png',
        contentType: 'image/png',
        // @ts-expect-error 故意制造类型违规
        bytes: undefined,
      }),
    ).toThrow(/bytes required/);
  });
});

describe('markDirty / isDirty', () => {
  test('markDirty 把 XML part 置脏', () => {
    const part = createXmlPart({
      path: 'word/document.xml',
      contentType: 'application/xml',
      xmlDom: makeXmlDom(),
      originalBytes: new Uint8Array([1]),
    });
    expect(isDirty(part)).toBe(false);
    markDirty(part);
    expect(isDirty(part)).toBe(true);
  });

  test('markDirty 幂等(重复调用仍是 true)', () => {
    const part = createXmlPart({
      path: 'word/document.xml',
      contentType: 'application/xml',
      xmlDom: makeXmlDom(),
      originalBytes: new Uint8Array([1]),
    });
    markDirty(part);
    markDirty(part);
    expect(isDirty(part)).toBe(true);
  });

  test('二进制 part 也可被 markDirty(未来有插图替换场景)', () => {
    const part = createBinaryPart({
      path: 'word/media/img.png',
      contentType: 'image/png',
      bytes: new Uint8Array([1, 2, 3]),
    });
    expect(isDirty(part)).toBe(false);
    markDirty(part);
    expect(isDirty(part)).toBe(true);
  });
});

describe('isXmlPart / isBinaryPart', () => {
  test('XML part:isXmlPart=true,isBinaryPart=false', () => {
    const part = createXmlPart({
      path: 'word/document.xml',
      contentType: 'application/xml',
      xmlDom: makeXmlDom(),
    });
    expect(isXmlPart(part)).toBe(true);
    expect(isBinaryPart(part)).toBe(false);
  });

  test('binary part:isBinaryPart=true,isXmlPart=false', () => {
    const part = createBinaryPart({
      path: 'word/media/img.png',
      contentType: 'image/png',
      bytes: new Uint8Array([0x89, 0x50]),
    });
    expect(isBinaryPart(part)).toBe(true);
    expect(isXmlPart(part)).toBe(false);
  });
});
