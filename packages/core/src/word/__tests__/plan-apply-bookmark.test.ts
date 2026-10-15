/**
 * WordPlan.apply 端到端测试 —— #34(书签创建 / 删除)。
 *
 * 走整条链:doc.plan() → builder → apply() → kernel.applyPlan → dispatchOp →
 * executeCreateBookmark / executeDeleteBookmark → toZip。
 *
 * 覆盖:
 *   - createBookmark ParagraphRange → apply → listBookmarks 含新书签,rangeKind='paragraph'
 *   - createBookmark TextRange → apply → listBookmarks rangeKind='text',text 字段匹配
 *   - deleteBookmark → apply → listBookmarks 不含该名
 *   - 事务回滚:创建 + 删除不存在的书签 → 第一个书签不应出现
 *   - BOOKMARK_NAME_EXISTS 抛错路径
 *
 * inline 模式;worker 端到端见 plan-apply-worker-bookmark.test.ts。
 */
import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
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

describe('WordPlan apply —— createBookmark ParagraphRange 端到端(inline)', () => {
  test('apply 后 listBookmarks 含新书签,rangeKind=paragraph,paraId 正确', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>content</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-para', mode: 'inline' });
    const result = await doc
      .plan()
      .createBookmark({ name: 'myBm', range: { paraId: 'AAAAAAA1' } })
      .apply();

    expect(result.createdBookmarks).toEqual([
      { opIndex: 0, name: 'myBm', range: { fromParaId: 'AAAAAAA1', toParaId: 'AAAAAAA1' } },
    ]);

    const bookmarks = await doc.bookmarks.list();
    const bm = bookmarks.find((b) => b.name === 'myBm');
    expect(bm).toBeDefined();
    expect(bm!.rangeKind).toBe('paragraph');
    expect(bm!.paraId?.toUpperCase()).toBe('AAAAAAA1');
    await doc.dispose();
  });

  test('apply 后 toZip 含 bookmarkStart/End 元素', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-zip', mode: 'inline' });
    await doc
      .plan()
      .createBookmark({ name: 'zipBm', range: { paraId: 'AAAAAAA1' } })
      .apply();
    const out = await doc.toZip();
    const roundtrip = await JSZip.loadAsync(out);
    const xml = await roundtrip.file('word/document.xml')!.async('string');
    expect(xml).toContain('bookmarkStart');
    expect(xml).toContain('zipBm');
    await doc.dispose();
  });
});

describe('WordPlan apply —— createBookmark TextRange 端到端(inline)', () => {
  test('apply 后 listBookmarks rangeKind=text,text 字段匹配', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>the quick brown fox</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-text', mode: 'inline' });
    await doc
      .plan()
      .createBookmark({ name: 'txtBm', range: { paraId: 'AAAAAAA1', text: 'quick' } })
      .apply();

    const bookmarks = await doc.bookmarks.list();
    const bm = bookmarks.find((b) => b.name === 'txtBm');
    expect(bm).toBeDefined();
    expect(bm!.rangeKind).toBe('text');
    expect(bm!.text).toBe('quick');
    await doc.dispose();
  });
});

describe('WordPlan apply —— deleteBookmark 端到端(inline)', () => {
  test('deleteBookmark → apply → listBookmarks 不含该名', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="toDelete"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-del', mode: 'inline' });

    // 确认书签存在
    const before = await doc.bookmarks.list();
    expect(before.find((b) => b.name === 'toDelete')).toBeDefined();

    await doc.plan().deleteBookmark('toDelete').apply();

    const after = await doc.bookmarks.list();
    expect(after.find((b) => b.name === 'toDelete')).toBeUndefined();
    await doc.dispose();
  });
});

describe('WordPlan apply —— 事务回滚(inline)', () => {
  test('创建书签 + 删除不存在的书签 → 第一个书签不应出现(回滚)', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>para</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-tx', mode: 'inline' });

    const plan = doc
      .plan()
      .createBookmark({ name: 'firstBm', range: { paraId: 'AAAAAAA1' } })
      .deleteBookmark('doesNotExist'); // 这个会抛 BOOKMARK_NOT_FOUND

    await expect(plan.apply()).rejects.toBeInstanceOf(WordError);

    // 事务回滚:firstBm 不应存在
    const bookmarks = await doc.bookmarks.list();
    expect(bookmarks.find((b) => b.name === 'firstBm')).toBeUndefined();
    await doc.dispose();
  });
});

describe('WordPlan apply —— BOOKMARK_NAME_EXISTS 错误路径(inline)', () => {
  test('createBookmark 重名 → apply 抛 BOOKMARK_NAME_EXISTS', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="existing"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-bm-exists', mode: 'inline' });
    const plan = doc.plan().createBookmark({ name: 'existing', range: { paraId: 'AAAAAAA1' } });

    try {
      await plan.apply();
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('BOOKMARK_NAME_EXISTS');
    }
    await doc.dispose();
  });
});
