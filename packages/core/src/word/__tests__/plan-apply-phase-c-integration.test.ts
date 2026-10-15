/**
 * Phase C 端到端集成测试(#30)。
 *
 * 覆盖 #28 / #29 / #30(drop 模式)的组合场景,以及事务原子性在混合 op 里的表现。
 * 纯 executor 测试见 `plan-executors/__tests__/*.test.ts`;本文件聚焦**跨 op 的**
 * 完整链路:doc.plan() → 多 op builder → apply → transport RPC → kernel.applyPlan
 * → dispatch → runTransaction → toZip。
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

describe('Phase C 集成 —— #28 + #29 + #30 组合', () => {
  test('多 op 同 plan:insertAfter + replaceText + deleteParagraph 按序生效', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>alpha</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>beta</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'c-multi', mode: 'inline' });
    const result = await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'new-para' })
      .replaceTextInParagraph('AAAAAAA2', 'beta', 'gamma')
      .deleteParagraph('AAAAAAA1')
      .apply();

    // insertedParagraphs 一条,来自 insertParagraphAfter
    expect(result.insertedParagraphs).toHaveLength(1);
    const { newParaId } = result.insertedParagraphs[0]!;

    // 段落序列:new-para, gamma(AAAAAAA1 被删)
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await Promise.all(list.paragraphs.map((p) => p.getText()))).toEqual([
      'new-para',
      'gamma',
    ]);
    expect(list.paragraphs[0]!.paraId).toBe(newParaId);

    await doc.dispose();
  });

  test('事务原子性:一串 op 里抛错 → 整体回滚', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>keep</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'c-tx', mode: 'inline' });
    const plan = doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'temp' })
      .replaceParagraph('NOTEXIST', 'x'); // 第二 op 抛 PARA_NOT_FOUND

    await expect(plan.apply()).rejects.toBeInstanceOf(WordError);
    // 回滚后只有原始 AAAAAAA1,段落数为 1
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs).toHaveLength(1);
    expect(await list.paragraphs[0]!.getText()).toBe('keep');
    await doc.dispose();
  });

  test('deleteParagraph drop:段内书签 + 段一并清理', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="bm-keep"/>
        <w:r><w:t>with-bm</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>
      <w:p w14:paraId="AAAAAAA2"><w:r><w:t>survivor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'c-drop-bm', mode: 'inline' });
    await doc.plan().deleteParagraph('AAAAAAA1', { onAnchorConflict: 'drop' }).apply();
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await Promise.all(list.paragraphs.map((p) => p.getText()))).toEqual(['survivor']);
    const bookmarks = await doc.bookmarks.list({ includeHidden: true });
    expect(bookmarks.map((b) => b.name)).not.toContain('bm-keep');
    await doc.dispose();
  });

  test('ANCHOR_CONFLICT default:段有书签时 replaceParagraph 抛错,事务全回滚', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="keep"/>
        <w:r><w:t>protected</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'c-conflict', mode: 'inline' });
    const plan = doc.plan().replaceParagraph('AAAAAAA1', 'force-replace');
    await expect(plan.apply()).rejects.toMatchObject({ code: 'ANCHOR_CONFLICT' });

    // 回滚后原内容 + 书签保留
    const info = await doc.paragraphs.get('AAAAAAA1');
    expect(await info?.getText()).toBe('protected');
    const bookmarks = await doc.bookmarks.list();
    expect(bookmarks.map((b) => b.name)).toContain('keep');
    await doc.dispose();
  });

  test('replaceTextInParagraph 匹配区间外有书签:不触发 ANCHOR_CONFLICT', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="start-bm"/>
        <w:r><w:t>before target after</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'c-outside', mode: 'inline' });
    // 匹配 'target',在 run 中间,bookmarkStart 在段首、End 在段末 —— 现在 collectRangeAnchors
    // 扫 [startRun, endRun] sibling 链内的锚点,恰好 bookmarkStart/End 不在这个区间里?
    // 验证:取决于 splitRunAt 后 startRun/endRun 的 sibling 关系。
    // 当前实现下,bookmarkStart / End 是段直属子,位于 <w:r> 的兄弟层;
    // 匹配 run 切分后,startRun / endRun 指向切出的中间 run;sibling 从 startRun 到 endRun
    // 只含 match 段,不包括段首 bookmarkStart。所以应该成功。
    const result = await doc.plan().replaceTextInParagraph('AAAAAAA1', 'target', 'X').apply();
    expect(result).toBeDefined();
    expect(await (await doc.paragraphs.get('AAAAAAA1'))?.getText()).toBe('before X after');
    // 书签仍在
    const bookmarks = await doc.bookmarks.list();
    expect(bookmarks.map((b) => b.name)).toContain('start-bm');
    await doc.dispose();
  });

  test('toZip → 回加载:整链变更全部持久化', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>before</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'c-zip', mode: 'inline' });
    await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'inserted' })
      .replaceTextInParagraph('AAAAAAA1', 'before', 'after')
      .apply();
    const out = await doc.toZip();

    const reloaded = await WordDocument.fromZip(out, { docId: 'c-zip-2', mode: 'inline' });
    const list = await reloaded.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await Promise.all(list.paragraphs.map((p) => p.getText()))).toEqual([
      'after',
      'inserted',
    ]);
    await doc.dispose();
    await reloaded.dispose();
  });

  test('多轮 plan:第二轮基于第一轮的结果继续操作', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>round-0</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'c-multi-plan', mode: 'inline' });

    const r1 = await doc.plan().insertParagraphAfter('AAAAAAA1', { text: 'round-1' }).apply();
    const newId1 = r1.insertedParagraphs[0]!.newParaId;

    await doc.plan().insertParagraphAfter(newId1, { text: 'round-2' }).apply();

    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await Promise.all(list.paragraphs.map((p) => p.getText()))).toEqual([
      'round-0',
      'round-1',
      'round-2',
    ]);
    await doc.dispose();
  });

  test('ANCHOR_CONFLICT 不阻塞后续 plan:整 plan 回滚,新 plan 可继续', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="b1"/>
        <w:r><w:t>locked</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'c-bounce', mode: 'inline' });
    // 第一次:conflict 回滚
    await expect(doc.plan().replaceParagraph('AAAAAAA1', 'x').apply()).rejects.toMatchObject({
      code: 'ANCHOR_CONFLICT',
    });
    // 第二次:drop 模式成功
    await doc.plan().deleteParagraph('AAAAAAA1', { onAnchorConflict: 'drop' }).apply();
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs).toHaveLength(0);
    await doc.dispose();
  });
});
