/**
 * WordPlan.apply 端到端测试 —— #28(段落修改)。
 *
 * 走整条链:doc.plan() → builder → apply() → kernel.applyPlan → dispatchOp →
 * executeReplaceParagraph / executeReplaceTextInParagraph → runTransaction → toZip。
 *
 * 覆盖:
 *   - 单 op:replaceParagraph 改文本;替换后 toZip 含新文本
 *   - 单 op:replaceTextInParagraph 部分替换;首 rPr 继承
 *   - 多 op(同 plan):按序 apply,两次改动都生效
 *   - 失败事务回滚:第一个 op 改完,第二个 op 抛 PARA_NOT_FOUND → 整体回滚,
 *     第一个 op 的改动也应撤销
 *   - 索引同步:apply 后 getParagraph 返回新文本
 *   - Header / body 混合:改 body 不影响 header
 *
 * inline 模式;worker 端到端见 [plan-apply-worker.test.ts](./plan-apply-worker.test.ts)。
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

describe('WordPlan apply —— replaceParagraph 端到端(inline)', () => {
  test('replaceParagraph:apply 后 getParagraph 返回新文本', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-rp', mode: 'inline' });
    const result = await doc.plan().replaceParagraph('AAAAAAA1', 'new-text').apply();
    expect(result.insertedParagraphs).toEqual([]);

    const info = await doc.paragraphs.get('AAAAAAA1');
    expect(await info?.getText()).toBe('new-text');
    await doc.dispose();
  });

  test('replaceParagraph:toZip 产物含新文本', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old-content</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-zip', mode: 'inline' });
    await doc.plan().replaceParagraph('AAAAAAA1', 'new-content').apply();
    const out = await doc.toZip();

    // 把产物再解开读一次 document.xml 验证
    const roundtrip = await JSZip.loadAsync(out);
    const xml = await roundtrip.file('word/document.xml')!.async('string');
    expect(xml).toContain('new-content');
    expect(xml).not.toContain('old-content');
    await doc.dispose();
  });

  test('多 op 同 plan:按序生效', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>foo</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>bar</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-multi', mode: 'inline' });
    await doc
      .plan()
      .replaceParagraph('AAAAAAA1', 'new-foo')
      .replaceParagraph('AAAAAAA2', 'new-bar')
      .apply();
    expect(await (await doc.paragraphs.get('AAAAAAA1'))?.getText()).toBe('new-foo');
    expect(await (await doc.paragraphs.get('AAAAAAA2'))?.getText()).toBe('new-bar');
    await doc.dispose();
  });

  test('事务回滚:其中一个 op 抛错,之前的 op 也被撤销', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>keep</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-tx', mode: 'inline' });

    const plan = doc.plan().replaceParagraph('AAAAAAA1', 'mutated').replaceParagraph('NOPE', 'x'); // 第二个 op 会抛 PARA_NOT_FOUND

    await expect(plan.apply()).rejects.toBeInstanceOf(WordError);

    // 回滚后原文仍是 'keep'
    const info = await doc.paragraphs.get('AAAAAAA1');
    expect(await info?.getText()).toBe('keep');
    await doc.dispose();
  });

  test('索引同步:apply 后 paraId 仍可 getParagraph', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-idx', mode: 'inline' });
    await doc.plan().replaceParagraph('AAAAAAA1', 'after').apply();
    // paraIdx 应指向同一 <w:p> 元素,还能查到
    const info = await doc.paragraphs.get('AAAAAAA1');
    expect(info).not.toBeNull();
    expect(await info!.getText()).toBe('after');
    await doc.dispose();
  });
});

describe('WordPlan apply —— replaceTextInParagraph 端到端', () => {
  test('段内替换,paragraph 其它 run 保留', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>hello </w:t></w:r>
        <w:r><w:t>world</w:t></w:r>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-rt', mode: 'inline' });
    await doc.plan().replaceTextInParagraph('AAAAAAA1', 'world', 'everyone').apply();
    const info = await doc.paragraphs.get('AAAAAAA1');
    expect(await info?.getText()).toBe('hello everyone');
    await doc.dispose();
  });

  test('scope=all 多次命中', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>a a a</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-all', mode: 'inline' });
    await doc.plan().replaceTextInParagraph('AAAAAAA1', 'a', 'b', { scope: 'all' }).apply();
    expect(await (await doc.paragraphs.get('AAAAAAA1'))?.getText()).toBe('b b b');
    await doc.dispose();
  });
});
