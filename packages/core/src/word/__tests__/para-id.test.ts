import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { ensureParaIds } from '../para-id.js';
import { W14_NS, W_NS } from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function xmlPartFromString(path: string, xml: string) {
  const dom = parseXmlDocument(xml) as unknown as Document;
  return createXmlPart({
    path,
    contentType: DOC_CTYPE,
    xmlDom: dom,
    originalBytes: new TextEncoder().encode(xml),
  });
}

function makePackageWithWordDoc(xml: string): Package {
  const pkg = createPackage();
  setPart(pkg, xmlPartFromString('word/document.xml', xml));
  return pkg;
}

function getAllParaIds(pkg: Package): string[] {
  const ids: string[] = [];
  for (const part of pkg.parts.values()) {
    if (!part.xmlDom) continue;
    const ps = part.xmlDom.getElementsByTagNameNS(W_NS, 'p');
    for (let i = 0; i < ps.length; i++) {
      const id = ps[i]?.getAttributeNS(W14_NS, 'paraId');
      if (id) ids.push(id);
    }
  }
  return ids;
}

describe('ensureParaIds', () => {
  test('单段落无 paraId → 分配一个 8 位十六进制 id', () => {
    const pkg = makePackageWithWordDoc(`<?xml version="1.0"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p><w:r><w:t>hello</w:t></w:r></w:p>
  </w:body>
</w:document>`);
    const result = ensureParaIds(pkg);
    expect(result.added).toBe(1);
    expect(result.total).toBe(1);
    expect(result.touchedParts).toEqual(['word/document.xml']);

    const ids = getAllParaIds(pkg);
    expect(ids).toHaveLength(1);
    expect(ids[0]).toMatch(/^[0-9A-F]{8}$/);
  });

  test('既有 paraId 保留,新段落各拿一个', () => {
    const pkg = makePackageWithWordDoc(`<?xml version="1.0"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="DEADBEEF"><w:r><w:t>existing</w:t></w:r></w:p>
    <w:p><w:r><w:t>new 1</w:t></w:r></w:p>
    <w:p w14:paraId="CAFEBABE"><w:r><w:t>existing 2</w:t></w:r></w:p>
    <w:p><w:r><w:t>new 2</w:t></w:r></w:p>
  </w:body>
</w:document>`);
    const result = ensureParaIds(pkg);
    expect(result.added).toBe(2);
    expect(result.total).toBe(4);

    const ids = getAllParaIds(pkg);
    expect(ids).toHaveLength(4);
    expect(ids).toContain('DEADBEEF');
    expect(ids).toContain('CAFEBABE');
    // 所有 id 都是 8 hex
    for (const id of ids) expect(id).toMatch(/^[0-9A-F]{8}$/);
    // 无重复
    expect(new Set(ids).size).toBe(4);
  });

  test('空 package(无 XML part) → added=0, total=0, touched=空', () => {
    const pkg = createPackage();
    const result = ensureParaIds(pkg);
    expect(result).toEqual({ added: 0, total: 0, touchedParts: [] });
  });

  test('所有段落都已有 paraId → 不分配,不标脏', () => {
    const xml = `<?xml version="1.0"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="AAAAAAAA"><w:r><w:t>a</w:t></w:r></w:p>
    <w:p w14:paraId="BBBBBBBB"><w:r><w:t>b</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    const pkg = makePackageWithWordDoc(xml);
    const docPart = pkg.parts.get('word/document.xml')!;
    expect(docPart._dirty).toBe(false);

    const result = ensureParaIds(pkg);
    expect(result.added).toBe(0);
    expect(result.total).toBe(2);
    expect(result.touchedParts).toEqual([]);
    expect(docPart._dirty).toBe(false);
  });

  test('发生分配时 part 被 markDirty,未发生分配的 part 不被 markDirty', () => {
    const pkg = createPackage();
    // 完全已有 paraId 的 part
    setPart(
      pkg,
      xmlPartFromString(
        'word/header1.xml',
        `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
          <w:p w14:paraId="11111111"><w:r><w:t>hdr</w:t></w:r></w:p>
        </w:hdr>`,
      ),
    );
    // 缺 paraId 的 part
    setPart(
      pkg,
      xmlPartFromString(
        'word/document.xml',
        `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
          <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
        </w:document>`,
      ),
    );

    const result = ensureParaIds(pkg);
    expect(result.added).toBe(1);
    expect(result.total).toBe(2);
    expect(result.touchedParts).toEqual(['word/document.xml']);

    expect(pkg.parts.get('word/header1.xml')!._dirty).toBe(false);
    expect(pkg.parts.get('word/document.xml')!._dirty).toBe(true);
  });

  test('paraId 全局唯一跨 body / header / footer', () => {
    const pkg = createPackage();
    setPart(
      pkg,
      xmlPartFromString(
        'word/document.xml',
        `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
          <w:body>
            <w:p><w:r><w:t>b1</w:t></w:r></w:p>
            <w:p><w:r><w:t>b2</w:t></w:r></w:p>
          </w:body>
        </w:document>`,
      ),
    );
    setPart(
      pkg,
      xmlPartFromString(
        'word/header1.xml',
        `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
          <w:p><w:r><w:t>h1</w:t></w:r></w:p>
        </w:hdr>`,
      ),
    );
    setPart(
      pkg,
      xmlPartFromString(
        'word/footer1.xml',
        `<w:ftr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
          <w:p><w:r><w:t>f1</w:t></w:r></w:p>
        </w:ftr>`,
      ),
    );

    const result = ensureParaIds(pkg);
    expect(result.added).toBe(4);
    expect(result.total).toBe(4);
    expect(new Set(result.touchedParts)).toEqual(
      new Set(['word/document.xml', 'word/header1.xml', 'word/footer1.xml']),
    );

    const ids = getAllParaIds(pkg);
    expect(ids).toHaveLength(4);
    expect(new Set(ids).size).toBe(4);
  });

  test('幂等:第二次调用不产生新分配', () => {
    const pkg = makePackageWithWordDoc(`<?xml version="1.0"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p><w:r><w:t>hello</w:t></w:r></w:p>
    <w:p><w:r><w:t>world</w:t></w:r></w:p>
  </w:body>
</w:document>`);
    const first = ensureParaIds(pkg);
    expect(first.added).toBe(2);

    const second = ensureParaIds(pkg);
    expect(second.added).toBe(0);
    expect(second.total).toBe(2);
    expect(second.touchedParts).toEqual([]);
  });

  test('大档基线:1000 段落应该在百毫秒内完成(非严格基线,保证不退化)', () => {
    const paragraphs = Array.from(
      { length: 1000 },
      (_, i) => `<w:p><w:r><w:t>p${i}</w:t></w:r></w:p>`,
    ).join('');
    const xml = `<?xml version="1.0"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${paragraphs}</w:body>
</w:document>`;
    const pkg = makePackageWithWordDoc(xml);
    const start = performance.now();
    const result = ensureParaIds(pkg);
    const elapsed = performance.now() - start;
    expect(result.added).toBe(1000);
    expect(result.total).toBe(1000);
    // 阈值 500ms(宽松容 CI 慢环境),正常应该 <50ms
    expect(elapsed).toBeLessThan(500);
  });
});
