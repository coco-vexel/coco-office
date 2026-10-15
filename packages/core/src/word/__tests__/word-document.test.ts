import { afterEach, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordError } from '../word-error.js';
import { WordDocument } from '../word-document.js';
import { W14_NS, W_NS } from '../namespaces.js';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p><w:r><w:t>hello</w:t></w:r></w:p>
    <w:p><w:r><w:t>world</w:t></w:r></w:p>
  </w:body>
</w:document>`;

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49,
  0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06,
  0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0b, 0x49, 0x44,
  0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d,
  0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42,
  0x60, 0x82,
]);

async function makeFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', DOC_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

const FLAT_OPC = `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p><w:r><w:t>flat</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

afterEach(() => {
  WordDocument._resetDefaults();
});

describe('WordDocument.createBlank', () => {
  test('创建一个含空段落锚点的空白文档', async () => {
    const doc = await WordDocument.createBlank({ docId: 'blank-doc', mode: 'inline' });

    expect(doc).toBeInstanceOf(WordDocument);
    expect(doc.docId).toBe('blank-doc');
    expect(doc.mode).toBe('inline');

    const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(paragraphs).toHaveLength(1);
    expect(await paragraphs[0]!.getText()).toBe('');
    expect(paragraphs[0]!.paraId).toMatch(/^[0-9A-F]{8}$/);
  });

  test('可直接写入内容并导出为可重开的 docx', async () => {
    const doc = await WordDocument.createBlank({ docId: 'blank-write', mode: 'inline' });
    const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });

    await paragraphs[0]!.replaceWith('hello from blank');
    const out = await doc.toZip();
    const reopened = await WordDocument.fromZip(out, {
      docId: 'blank-write-reopen',
      mode: 'inline',
    });
    const after = await reopened.paragraphs.list({ pagination: { mode: 'all' } });

    expect(after.paragraphs).toHaveLength(1);
    expect(await after.paragraphs[0]!.getText()).toBe('hello from blank');
  });

  test('appendParagraph 鍙湪姝ｆ枃鏈熬杩藉姞鏂版钀?', async () => {
    const doc = await WordDocument.createBlank({ docId: 'blank-append', mode: 'inline' });

    const result = await doc.appendParagraph('appended paragraph');
    const { paragraphs } = await doc.paragraphs.list({ scope: 'body', pagination: { mode: 'all' } });

    expect(result.insertedParagraphs).toHaveLength(1);
    expect(paragraphs.map((paragraph) => paragraph.text)).toEqual(['', 'appended paragraph']);
  });

  test('paragraph.insertImageAfter 鍙洿鎺ュ湪娈靛悗鎻掑叆鍥剧墖', async () => {
    const doc = await WordDocument.createBlank({ docId: 'blank-image', mode: 'inline' });
    const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });

    await paragraphs[0]!.replaceWith('image anchor');
    await paragraphs[0]!.insertImageAfter({
      bytes: TINY_PNG,
      contentType: 'image/png',
      widthEmu: 914400,
      heightEmu: 914400,
    });

    const pictures = await doc.drawings.listPictures({ scope: 'body' });
    expect(pictures).toHaveLength(1);
    expect(pictures[0]!.kind).toBe('picture');
  });

  test('paragraph.insertTableAfter 鍙洿鎺ュ湪娈靛悗鎻掑叆绠€鍗曡〃鏍?', async () => {
    const doc = await WordDocument.createBlank({ docId: 'blank-table', mode: 'inline' });
    const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });

    await paragraphs[0]!.replaceWith('table anchor');
    await paragraphs[0]!.insertTableAfter({
      rows: [
        ['Name', 'Score'],
        ['Ada & Bob', '42'],
      ],
      columnWidths: [1800, 1200],
    });

    const tables = await doc.tables.list({ scope: 'body', includeCellTexts: true });
    expect(tables).toHaveLength(1);
    expect(tables[0]!.rowCount).toBe(2);
    expect(tables[0]!.colCount).toBe(2);
    await expect(tables[0]!.getCells({ includeCellTexts: true })).resolves.toEqual([
      ['Name', 'Score'],
      ['Ada & Bob', '42'],
    ]);
  });
});

describe('WordDocument.fromZip', () => {
  test('正常装载:返回实例,docId / mode 正确', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'doc-A' });
    expect(doc).toBeInstanceOf(WordDocument);
    expect(doc.docId).toBe('doc-A');
    expect(doc.mode).toBe('inline');
  });

  test('装载期自动 ensureParaIds(给段落补 w14:paraId)', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'doc-A' });
    const indices = doc._getIndices();
    // DOC_XML 里两段都没有 w14:paraId,装载后都应有
    expect(indices.paraIdx.size).toBe(2);
    for (const id of indices.paraIdx.keys()) {
      expect(id).toMatch(/^[0-9A-F]{8}$/);
    }
  });

  test('装载期建索引(paraIdx / bookmarkIdx / commentIdx)', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'doc-A' });
    const { paraIdx, bookmarkIdx, commentIdx } = doc._getIndices();
    expect(paraIdx).toBeInstanceOf(Map);
    expect(bookmarkIdx).toBeInstanceOf(Map);
    expect(commentIdx).toBeInstanceOf(Map);
    expect(paraIdx.size).toBe(2);
    expect(bookmarkIdx.size).toBe(0);
    expect(commentIdx.size).toBe(0);
  });

  test('onProgress 按顺序回调 parse 与 normalize', async () => {
    const buf = await makeFixtureZip();
    const phases: string[] = [];
    await WordDocument.fromZip(buf, {
      docId: 'doc-A',
      onProgress: (p) => phases.push(p),
    });
    expect(phases).toContain('parse');
    expect(phases).toContain('normalize');
    expect(phases.indexOf('parse')).toBeLessThan(phases.indexOf('normalize'));
  });

  test('docId 缺失 → WordError(PACKAGE_UNSUPPORTED_FORMAT)', async () => {
    const buf = await makeFixtureZip();
    // @ts-expect-error 故意构造缺 docId
    await expect(WordDocument.fromZip(buf, {})).rejects.toMatchObject({
      code: 'PACKAGE_UNSUPPORTED_FORMAT',
    });
  });

  test('docId 空字符串 → 拒收', async () => {
    const buf = await makeFixtureZip();
    await expect(WordDocument.fromZip(buf, { docId: '' })).rejects.toBeInstanceOf(WordError);
  });

  test('mode=worker 抛 WORKER_CREATION_FAILED(本 PR 尚未实装 worker)', async () => {
    const buf = await makeFixtureZip();
    await expect(
      WordDocument.fromZip(buf, { docId: 'doc-A', mode: 'worker' }),
    ).rejects.toMatchObject({ code: 'WORKER_CREATION_FAILED', docId: 'doc-A' });
  });

  test('mode=inline 显式也走 inline', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'doc-A', mode: 'inline' });
    expect(doc.mode).toBe('inline');
  });

  test('mode=auto 默认走 inline(真 auto 在 #18)', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'doc-A', mode: 'auto' });
    expect(doc.mode).toBe('inline');
  });

  test('缺 [Content_Types].xml → PACKAGE_CORRUPT(docId 自动注入)', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', DOC_XML);
    const bad = await zip.generateAsync({ type: 'uint8array' });
    try {
      await WordDocument.fromZip(bad, { docId: 'doc-B' });
      throw new Error('should have thrown');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      const err = e as WordError;
      expect(err.code).toBe('PACKAGE_CORRUPT');
      expect(err.docId).toBe('doc-B');
      expect(err.details.phase).toBe('fromZip');
    }
  });

  test('重名 bookmark(含真实文档常见的重复 _GoBack)装载容错:不崩,正常返回 doc', async () => {
    // 多次编辑 / 多工具处理过的真实文档常出现重复书签名(尤其 Word 自动插入的 _GoBack)。
    // 旧逻辑遇重名抛 BOOKMARK_NAME_DUPLICATE_ON_LOAD → fromZip reject → 排版崩溃。
    // 修复后必须容错(保留第一个、忽略后续),fromZip 正常返回。
    const zip = new JSZip();
    zip.file('[Content_Types].xml', CT_XML);
    zip.file('_rels/.rels', MAIN_RELS);
    zip.file(
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}"><w:body>
        <w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p>
        <w:p><w:r><w:t>正文</w:t></w:r></w:p>
        <w:p><w:bookmarkStart w:id="1" w:name="_GoBack"/><w:bookmarkEnd w:id="1"/></w:p>
        <w:p><w:bookmarkStart w:id="2" w:name="Dup"/><w:bookmarkEnd w:id="2"/></w:p>
        <w:p><w:bookmarkStart w:id="3" w:name="Dup"/><w:bookmarkEnd w:id="3"/></w:p>
      </w:body></w:document>`,
    );
    const buf = await zip.generateAsync({ type: 'uint8array' });
    const doc = await WordDocument.fromZip(buf, { docId: 'doc-dup' });
    expect(doc.docId).toBe('doc-dup');
  });

  test('空 name bookmark → BOOKMARK_NAME_EMPTY_ON_LOAD', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', CT_XML);
    zip.file('_rels/.rels', MAIN_RELS);
    zip.file(
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}"><w:body>
        <w:p><w:bookmarkStart w:id="1" w:name=""/><w:bookmarkEnd w:id="1"/></w:p>
      </w:body></w:document>`,
    );
    const buf = await zip.generateAsync({ type: 'uint8array' });
    await expect(WordDocument.fromZip(buf, { docId: 'doc-empty' })).rejects.toMatchObject({
      code: 'BOOKMARK_NAME_EMPTY_ON_LOAD',
      docId: 'doc-empty',
    });
  });

  test('接受 ArrayBuffer 输入', async () => {
    const buf = await makeFixtureZip();
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const doc = await WordDocument.fromZip(ab as ArrayBuffer, { docId: 'doc-A' });
    expect(doc.docId).toBe('doc-A');
  });
});

describe('WordDocument.fromFlatOpc', () => {
  test('正常装载:docId / mode 正确', async () => {
    const doc = await WordDocument.fromFlatOpc(FLAT_OPC, { docId: 'doc-flat' });
    expect(doc.docId).toBe('doc-flat');
    expect(doc.mode).toBe('inline');
    expect(doc._getIndices().paraIdx.size).toBe(1);
  });

  test('无效 XML → PACKAGE_CORRUPT', async () => {
    await expect(WordDocument.fromFlatOpc('not xml', { docId: 'doc-x' })).rejects.toMatchObject({
      code: 'PACKAGE_CORRUPT',
      docId: 'doc-x',
    });
  });

  test('docId 必填', async () => {
    // @ts-expect-error 故意缺 docId
    await expect(WordDocument.fromFlatOpc(FLAT_OPC, {})).rejects.toBeInstanceOf(WordError);
  });
});

describe('WordDocument.toZip / toFlatOpc', () => {
  test('fromZip → toZip 可往返', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'd' });
    const out = await doc.toZip();
    expect(out).toBeInstanceOf(Uint8Array);
    expect(out.length).toBeGreaterThan(0);
    // 第二次装载应成功
    const doc2 = await WordDocument.fromZip(out, { docId: 'd' });
    expect(doc2._getIndices().paraIdx.size).toBe(2);
  });

  test('fromFlatOpc → toFlatOpc 字符串含 <?mso-application?>', async () => {
    const doc = await WordDocument.fromFlatOpc(FLAT_OPC, { docId: 'd' });
    const out = await doc.toFlatOpc();
    expect(typeof out).toBe('string');
    expect(out).toMatch(/<\?mso-application progid="Word\.Document"\?>/);
  });

  test('fromZip → toFlatOpc:跨形态输出', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'd' });
    const out = await doc.toFlatOpc();
    expect(out).toMatch(/pkg:package/);
  });

  test('fromFlatOpc → toZip:跨形态输出可再装载', async () => {
    const doc = await WordDocument.fromFlatOpc(FLAT_OPC, { docId: 'd' });
    const out = await doc.toZip();
    const doc2 = await WordDocument.fromZip(out, { docId: 'd' });
    expect(doc2._getIndices().paraIdx.size).toBe(1);
  });
});

describe('WordDocument.dispose', () => {
  test('dispose 后 toZip / toFlatOpc 抛错', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'd' });
    await doc.dispose();
    await expect(doc.toZip()).rejects.toBeInstanceOf(WordError);
    await expect(doc.toFlatOpc()).rejects.toBeInstanceOf(WordError);
  });
});

describe('静态 API — setDefaultMode / detectMode / setDefaultAuthor', () => {
  test('detectMode 返 inline(本 PR 恒值)', () => {
    expect(WordDocument.detectMode()).toBe('inline');
  });

  test('setDefaultMode 改全局默认;装载时走默认', async () => {
    const buf = await makeFixtureZip();
    WordDocument.setDefaultMode('inline');
    const doc = await WordDocument.fromZip(buf, { docId: 'd' });
    expect(doc.mode).toBe('inline');

    WordDocument.setDefaultMode('worker');
    await expect(WordDocument.fromZip(buf, { docId: 'd' })).rejects.toMatchObject({
      code: 'WORKER_CREATION_FAILED',
    });
  });

  test('opts.mode 覆盖全局默认', async () => {
    const buf = await makeFixtureZip();
    WordDocument.setDefaultMode('worker');
    const doc = await WordDocument.fromZip(buf, { docId: 'd', mode: 'inline' });
    expect(doc.mode).toBe('inline');
  });

  test('setDefaultAuthor + _getDefaultAuthor 能 round-trip', () => {
    WordDocument.setDefaultAuthor('Claude');
    expect(WordDocument._getDefaultAuthor()).toBe('Claude');
  });

  test('_resetDefaults 回到初始(mode=auto,author="")', () => {
    WordDocument.setDefaultMode('worker');
    WordDocument.setDefaultAuthor('X');
    WordDocument._resetDefaults();
    expect(WordDocument._getDefaultMode()).toBe('auto');
    expect(WordDocument._getDefaultAuthor()).toBe('');
  });
});

describe('内部 accessor(_getPackage / _getIndices / _getInitialBytes / _getInitialFlatOpc)', () => {
  test('fromZip 装载后 _getInitialBytes 返回初始字节', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'd' });
    const initial = doc._getInitialBytes();
    expect(initial).toBeDefined();
    expect(initial).toBeInstanceOf(Uint8Array);
    expect(initial!.length).toBe(buf.length);
    expect(doc._getInitialFlatOpc()).toBeUndefined();
  });

  test('fromFlatOpc 装载后 _getInitialFlatOpc 返回原始字符串', async () => {
    const doc = await WordDocument.fromFlatOpc(FLAT_OPC, { docId: 'd' });
    expect(doc._getInitialFlatOpc()).toBe(FLAT_OPC);
    expect(doc._getInitialBytes()).toBeUndefined();
  });

  test('_getPackage 返回归一化后的 Package', async () => {
    const doc = await WordDocument.fromZip(await makeFixtureZip(), { docId: 'd' });
    const pkg = doc._getPackage();
    expect(pkg.parts.size).toBeGreaterThan(0);
  });
});
