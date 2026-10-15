/**
 * WordPlan apply 端到端 —— insertParagraphAfter + spec.bookmark(#176)。
 *
 * 覆盖:
 *   - 基本:apply 后 ApplyResult.createdBookmarks 含 { name, range },range 跨 anchor → newPara
 *   - 同 anchor 重复 append 触发 _2 后缀
 *   - rangeKind = complex(跨段,classifyRange 行为)
 *   - toZip → reload:书签随 OOXML 持久化
 *   - bookmarks.list({ prefix }) 过滤
 *   - prefix + includeHidden 独立轴
 *   - 非法 prefix 抛 BOOKMARK_NAME_INVALID
 *   - insertParagraphBefore + bookmark 抛 UNSUPPORTED_OOXML_FEATURE
 *   - trackChanges=true + bookmark 抛 UNSUPPORTED_OOXML_FEATURE
 */
import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
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

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

async function makeFixture(bodyInner: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  return zip.generateAsync({ type: 'uint8array' });
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('insertParagraphAfter + spec.bookmark 端到端 (#176)', () => {
  test('基本:createdBookmarks 含 name+range,range 跨 anchor → newPara', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>原文</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ipa-bm', mode: 'inline' });

    const result = await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', {
        text: 'translated',
        bookmark: { prefix: '_co_append' },
      })
      .apply();

    // insertedParagraphs 的 newParaId 与 createdBookmarks.range.toParaId 一致
    expect(result.insertedParagraphs).toHaveLength(1);
    const inserted = result.insertedParagraphs[0]!;
    const newParaId = inserted.newParaId;

    expect(result.createdBookmarks).toHaveLength(1);
    const bm = result.createdBookmarks[0]!;
    expect(bm.opIndex).toBe(0);
    expect(bm.name).toBe(`_co_append_AAAAAAA1`);
    expect(bm.range).toEqual({ fromParaId: 'AAAAAAA1', toParaId: newParaId });

    await doc.dispose();
  });

  test('同 anchor 二次 append 触发 _2 后缀(跨 plan 持久化检查)', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>原文</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ipa-bm-collide', mode: 'inline' });

    // 第一次:basename
    const r1 = await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'first', bookmark: { prefix: 'tag' } })
      .apply();
    expect(r1.createdBookmarks[0]!.name).toBe('tag_AAAAAAA1');

    // 第二次:同 anchor 同 prefix 触发冲突 → _2
    const r2 = await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'second', bookmark: { prefix: 'tag' } })
      .apply();
    expect(r2.createdBookmarks[0]!.name).toBe('tag_AAAAAAA1_2');

    await doc.dispose();
  });

  test('listBookmarks 命中,rangeKind=complex(跨段)', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ipa-bm-rangekind', mode: 'inline' });

    await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'new', bookmark: { prefix: 'span' } })
      .apply();

    // 默认不 includeHidden,但 'span_' 不是 _ 起头,所以会列出
    const bms = await doc.bookmarks.list();
    const bm = bms.find((b) => b.name === 'span_AAAAAAA1');
    expect(bm).toBeDefined();
    // start 和 end 在不同段 → classifyRange 返回 'complex'
    expect(bm!.rangeKind).toBe('complex');

    await doc.dispose();
  });

  test('toZip → reload:书签持久化(锚点段+新段+bookmarkStart/End 都落盘)', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, {
      docId: 'e2e-ipa-bm-persist',
      mode: 'inline',
    });

    const result = await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', {
        text: 'persisted',
        bookmark: { prefix: 'persist' },
      })
      .apply();
    const newParaId = result.insertedParagraphs[0]!.newParaId;
    const bookmarkName = result.createdBookmarks[0]!.name;

    const out = await doc.toZip();
    await doc.dispose();

    const doc2 = await WordDocument.fromZip(out, {
      docId: 'e2e-ipa-bm-persist-reload',
      mode: 'inline',
    });
    const bm = await doc2.bookmarks.get(bookmarkName);
    expect(bm).not.toBeNull();
    expect(bm!.rangeKind).toBe('complex');

    // 新段在重新装载后仍可查
    const para = await doc2.paragraphs.get(newParaId);
    expect(await para?.getText()).toBe('persisted');

    await doc2.dispose();
  });

  test('bookmarks.list({ prefix }) 过滤命中 + 未命中', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>p1</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>p2</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA3"><w:r><w:t>p3</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-prefix', mode: 'inline' });

    await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: '甲', bookmark: { prefix: 'transA' } })
      .insertParagraphAfter('AAAAAAA2', { text: '乙', bookmark: { prefix: 'transA' } })
      .insertParagraphAfter('AAAAAAA3', { text: '丙', bookmark: { prefix: 'reviewB' } })
      .apply();

    const transAList = await doc.bookmarks.list({ prefix: 'transA' });
    expect(transAList).toHaveLength(2);
    expect(transAList.every((b) => b.name.startsWith('transA'))).toBe(true);

    const reviewList = await doc.bookmarks.list({ prefix: 'reviewB' });
    expect(reviewList).toHaveLength(1);
    expect(reviewList[0]!.name).toBe('reviewB_AAAAAAA3');

    const nonMatch = await doc.bookmarks.list({ prefix: 'nope' });
    expect(nonMatch).toHaveLength(0);

    await doc.dispose();
  });

  test('prefix 与 includeHidden 独立:_ 起头的命名默认仍被隐藏过滤', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>p1</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-hidden', mode: 'inline' });

    await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', {
        text: '隐藏标签',
        bookmark: { prefix: '_co_append' },
      })
      .apply();

    // 默认 includeHidden=false,即使前缀匹配也被滤掉(_ 起头)
    const def = await doc.bookmarks.list({ prefix: '_co_append' });
    expect(def).toHaveLength(0);

    // includeHidden=true 才能拿到
    const all = await doc.bookmarks.list({ prefix: '_co_append', includeHidden: true });
    expect(all).toHaveLength(1);
    expect(all[0]!.name).toBe('_co_append_AAAAAAA1');

    await doc.dispose();
  });

  test('非法 prefix → BOOKMARK_NAME_INVALID,事务回滚(段不被插入)', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>only</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-bad-prefix', mode: 'inline' });

    await expect(
      doc
        .plan()
        .insertParagraphAfter('AAAAAAA1', { text: 'x', bookmark: { prefix: '1bad' } })
        .apply(),
    ).rejects.toMatchObject({ code: 'BOOKMARK_NAME_INVALID' });

    // 事务回退 —— 文档段数不变
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs).toHaveLength(1);

    await doc.dispose();
  });

  test('insertParagraphBefore + bookmark → UNSUPPORTED_OOXML_FEATURE', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>only</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-before', mode: 'inline' });

    await expect(
      doc
        .plan()
        .insertParagraphBefore('AAAAAAA1', { text: 'x', bookmark: { prefix: 'tag' } })
        .apply(),
    ).rejects.toMatchObject({ code: 'UNSUPPORTED_OOXML_FEATURE' });

    await doc.dispose();
  });

  test('trackChanges=true + bookmark → UNSUPPORTED_OOXML_FEATURE', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>only</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-tc', mode: 'inline' });

    await expect(
      doc
        .plan({ trackChanges: true })
        .insertParagraphAfter('AAAAAAA1', { text: 'x', bookmark: { prefix: 'tag' } })
        .apply(),
    ).rejects.toMatchObject({ code: 'UNSUPPORTED_OOXML_FEATURE' });

    await doc.dispose();
  });

  test('apply 后 toZip 字节里含书签元素', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-zip', mode: 'inline' });

    await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'new', bookmark: { prefix: 'persist' } })
      .apply();

    const out = await doc.toZip();
    const reloaded = await JSZip.loadAsync(out);
    const docXmlContent = await reloaded.file('word/document.xml')!.async('string');
    expect(docXmlContent).toContain('persist_AAAAAAA1');
    expect(docXmlContent).toMatch(/bookmarkStart/);
    expect(docXmlContent).toMatch(/bookmarkEnd/);

    await doc.dispose();
  });

  test('无 spec.bookmark 时回归路径不变(不创建任何书签)', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>only</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-regress', mode: 'inline' });

    const result = await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'no-bookmark' })
      .apply();

    expect(result.insertedParagraphs).toHaveLength(1);
    expect(result.createdBookmarks).toHaveLength(0);

    const bms = await doc.bookmarks.list({ includeHidden: true });
    expect(bms).toHaveLength(0);

    await doc.dispose();
  });
});
