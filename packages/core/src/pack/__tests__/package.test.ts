import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import {
  createBinaryPart,
  createPackage,
  createXmlPart,
  deletePart,
  getPart,
  hasPart,
  listParts,
  setPart,
} from '../index.js';

function xmlPart(path: string) {
  return createXmlPart({
    path,
    contentType: 'application/xml',
    xmlDom: parseXmlDocument('<r/>') as unknown as Document,
    originalBytes: new Uint8Array([0x3c, 0x72, 0x2f, 0x3e]),
  });
}

function binPart(path: string) {
  return createBinaryPart({
    path,
    contentType: 'image/png',
    bytes: new Uint8Array([0x89, 0x50]),
  });
}

describe('createPackage', () => {
  test('创建空包', () => {
    const pkg = createPackage();
    expect(pkg.parts).toBeInstanceOf(Map);
    expect(pkg.parts.size).toBe(0);
  });
});

describe('setPart / getPart / hasPart', () => {
  test('setPart 后 getPart 能取回,hasPart=true', () => {
    const pkg = createPackage();
    const p = xmlPart('word/document.xml');
    setPart(pkg, p);
    expect(hasPart(pkg, 'word/document.xml')).toBe(true);
    expect(getPart(pkg, 'word/document.xml')).toBe(p);
  });

  test('未设置的 path:getPart 返 undefined,hasPart=false', () => {
    const pkg = createPackage();
    expect(getPart(pkg, 'nonexistent.xml')).toBeUndefined();
    expect(hasPart(pkg, 'nonexistent.xml')).toBe(false);
  });

  test('同 path 再次 setPart 覆盖旧值', () => {
    const pkg = createPackage();
    const p1 = xmlPart('word/document.xml');
    const p2 = xmlPart('word/document.xml');
    setPart(pkg, p1);
    setPart(pkg, p2);
    expect(getPart(pkg, 'word/document.xml')).toBe(p2);
    expect(pkg.parts.size).toBe(1);
  });
});

describe('deletePart', () => {
  test('存在时删除并返 true', () => {
    const pkg = createPackage();
    setPart(pkg, xmlPart('word/document.xml'));
    expect(deletePart(pkg, 'word/document.xml')).toBe(true);
    expect(hasPart(pkg, 'word/document.xml')).toBe(false);
  });

  test('不存在时返 false', () => {
    const pkg = createPackage();
    expect(deletePart(pkg, 'nothing.xml')).toBe(false);
  });
});

describe('listParts', () => {
  test('返回所有 parts 数组,顺序 = 插入顺序', () => {
    const pkg = createPackage();
    setPart(pkg, xmlPart('word/document.xml'));
    setPart(pkg, binPart('word/media/img.png'));
    setPart(pkg, xmlPart('word/styles.xml'));
    const arr = listParts(pkg);
    expect(arr).toHaveLength(3);
    expect(arr.map((p) => p.path)).toEqual([
      'word/document.xml',
      'word/media/img.png',
      'word/styles.xml',
    ]);
  });

  test('空包返空数组', () => {
    expect(listParts(createPackage())).toEqual([]);
  });
});

describe('Package + Part 集成:dirty 状态与集合操作正交', () => {
  test('包内 XML part 初始 dirty=false,标脏后包里的同引用仍是 dirty', () => {
    const pkg = createPackage();
    const p = xmlPart('word/document.xml');
    setPart(pkg, p);
    expect(getPart(pkg, 'word/document.xml')?._dirty).toBe(false);

    p._dirty = true;

    // Map 存的是引用,不是拷贝
    expect(getPart(pkg, 'word/document.xml')?._dirty).toBe(true);
  });
});
