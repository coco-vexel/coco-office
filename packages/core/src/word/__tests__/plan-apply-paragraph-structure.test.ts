/**
 * WordPlan.apply 端到端测试 —— #29(段落结构操作)。
 *
 * 走整条链:doc.plan() → builder → apply() → kernel.applyPlan → dispatchOp →
 * executeInsertParagraph / executeDeleteParagraph → runTransaction → toZip。
 *
 * 覆盖:
 *   - insertParagraphAfter:apply 后 listParagraphs 返回 2 段,新 paraId 匹配 ApplyResult
 *   - insertParagraphBefore:新段位置在 anchor 之前
 *   - deleteParagraph:apply 后 getParagraph 返回 null(段被deleted);listParagraphs 少一段
 *   - 多 op:先 insertParagraphAfter 再 deleteParagraph(另一段),两者都生效
 *   - 事务回滚:insert 成功 + delete 抛 PARA_NOT_FOUND → 回滚后 paraIdx 原样
 *   - pageBreakBefore round-trip（task-core-pagebreak-before §3 #5/#5b）
 *
 * inline 模式;worker 端到端见 [plan-apply-worker-structure.test.ts](./plan-apply-worker-structure.test.ts)。
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

describe('WordPlan apply —— insertParagraphAfter 端到端(inline)', () => {
  test('apply 后 listParagraphs 返回 2 段,新 paraId 匹配 ApplyResult', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ipa', mode: 'inline' });

    const result = await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'new-paragraph' })
      .apply();

    expect(result.insertedParagraphs).toHaveLength(1);
    const entry = result.insertedParagraphs[0]!;
    expect(entry.position).toBe('after');
    expect(entry.anchorParaId).toBe('AAAAAAA1');
    expect(entry.newParaId).toMatch(/^[0-9A-F]{8}$/);

    // 新段可用 getParagraph 查到
    const newPara = await doc.paragraphs.get(entry.newParaId);
    expect(await newPara?.getText()).toBe('new-paragraph');

    // listParagraphs 返回 2 段
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs).toHaveLength(2);
    await doc.dispose();
  });

  test('insertParagraphBefore:新段在 anchor 之前', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>first</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ipb', mode: 'inline' });

    const result = await doc
      .plan()
      .insertParagraphBefore('AAAAAAA1', { text: 'before-first' })
      .apply();

    const entry = result.insertedParagraphs[0]!;
    expect(entry.position).toBe('before');

    // 全量列段,新段应在 AAAAAAA1 之前
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs[0]!.paraId).toBe(entry.newParaId);
    expect(list.paragraphs[1]!.paraId).toBe('AAAAAAA1');
    await doc.dispose();
  });

  test('toZip 产物含新段落文本和 paraId', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ipa-zip', mode: 'inline' });
    const result = await doc.plan().insertParagraphAfter('AAAAAAA1', { text: 'appended' }).apply();
    const out = await doc.toZip();

    const zip = await JSZip.loadAsync(out);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('appended');
    expect(xml).toContain(result.insertedParagraphs[0]!.newParaId);
    await doc.dispose();
  });

  test('color 选项:新段 run 有 <w:color w:val="..."/>', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ipa-color', mode: 'inline' });
    await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'translated', color: '0070C0' })
      .apply();
    const out = await doc.toZip();
    const zip = await JSZip.loadAsync(out);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('translated');
    // color val 大小写 / 单双引号在 slimdom 序列化下都一致 —— 用稍宽松的正则匹配
    expect(xml).toMatch(/<w:color\s+w:val=["']0070C0["']\s*\/>/);
    await doc.dispose();
  });

  test('color 留空:新段 run 不带 <w:color> 元素(默认色)', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ipa-no-color', mode: 'inline' });
    const result = await doc.plan().insertParagraphAfter('AAAAAAA1', { text: 'plain' }).apply();
    const out = await doc.toZip();
    const zip = await JSZip.loadAsync(out);
    const xml = await zip.file('word/document.xml')!.async('string');
    // 找到新段的 paraId 后,定位它的片段;断言新段 XML 不含 <w:color
    const newParaId = result.insertedParagraphs[0]!.newParaId;
    const m = xml.match(new RegExp(`<w:p[^>]*${newParaId}[\\s\\S]*?</w:p>`));
    expect(m).not.toBeNull();
    expect(m![0]).not.toMatch(/<w:color\b/);
    await doc.dispose();
  });
});

// ─────────── pageBreakBefore round-trip（task-core-pagebreak-before §3 #5/#5b）───────────

describe('WordPlan apply —— insertParagraph pageBreakBefore round-trip', () => {
  test('#5 round-trip zip：insert(pageBreakBefore:true) → toZip → fromZip → format.pageBreakBefore===true', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="PBBRT001"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'pbb-rt-zip', mode: 'inline' });

    const result = await doc
      .plan()
      .insertParagraphAfter('PBBRT001', { text: 'chapter-start', pageBreakBefore: true })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const roundTripXml = await (await JSZip.loadAsync(out))
      .file('word/document.xml')!
      .async('text');
    expect(roundTripXml).toContain('pageBreakBefore');

    // 重载后读回
    const doc2 = await WordDocument.fromZip(out, { docId: 'pbb-rt-zip-2', mode: 'inline' });
    const list = await doc2.paragraphs.list({ pagination: { mode: 'all' }, includeFormat: true });
    const newParaId = result.insertedParagraphs[0]!.newParaId;
    const newPara = list.paragraphs.find((p) => p.paraId === newParaId);
    expect(newPara).toBeDefined();
    expect((await newPara!.getInfo({ includeFormat: true }))?.format?.pageBreakBefore).toBe(true);
    await doc2.dispose();
  });

  test('#5b round-trip FlatOpc：insert(pageBreakBefore:true) → toFlatOpc → fromFlatOpc → format.pageBreakBefore===true', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="PBBRT002"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'pbb-rt-flat', mode: 'inline' });

    const result = await doc
      .plan()
      .insertParagraphAfter('PBBRT002', { text: 'chapter-start-flat', pageBreakBefore: true })
      .apply();

    const flatOpc = await doc.toFlatOpc();
    await doc.dispose();

    // fromFlatOpc 重载
    const doc2 = await WordDocument.fromFlatOpc(flatOpc, {
      docId: 'pbb-rt-flat-2',
      mode: 'inline',
    });
    const list = await doc2.paragraphs.list({ pagination: { mode: 'all' }, includeFormat: true });
    const newParaId = result.insertedParagraphs[0]!.newParaId;
    const newPara = list.paragraphs.find((p) => p.paraId === newParaId);
    expect(newPara).toBeDefined();
    expect((await newPara!.getInfo({ includeFormat: true }))?.format?.pageBreakBefore).toBe(true);
    await doc2.dispose();
  });
});

describe('WordPlan apply —— deleteParagraph 端到端(inline)', () => {
  test('apply 后 getParagraph 返回 null;listParagraphs 减少一段', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>del-me</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>keep</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-dp', mode: 'inline' });

    await doc.plan().deleteParagraph('AAAAAAA1').apply();

    expect(await doc.paragraphs.get('AAAAAAA1')).toBeNull();
    expect(await doc.paragraphs.get('AAAAAAA2')).not.toBeNull();

    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs).toHaveLength(1);
    await doc.dispose();
  });

  test('多 op:insert + delete 都生效', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>a</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>b</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-multi-struct', mode: 'inline' });

    const result = await doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'inserted' })
      .deleteParagraph('AAAAAAA2')
      .apply();

    expect(result.insertedParagraphs).toHaveLength(1);

    // 新段存在
    const newPara = await doc.paragraphs.get(result.insertedParagraphs[0]!.newParaId);
    expect(await newPara?.getText()).toBe('inserted');

    // AAAAAAA2 被删
    expect(await doc.paragraphs.get('AAAAAAA2')).toBeNull();

    // 共 2 段(AAAAAAA1 + 新段)
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs).toHaveLength(2);
    await doc.dispose();
  });

  test('事务回滚:insert 成功 + delete 抛 PARA_NOT_FOUND → 回滚后原样', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-rollback-struct', mode: 'inline' });

    const plan = doc
      .plan()
      .insertParagraphAfter('AAAAAAA1', { text: 'new' })
      .deleteParagraph('NONEXISTENT'); // 会抛 PARA_NOT_FOUND

    await expect(plan.apply()).rejects.toBeInstanceOf(WordError);

    // 回滚后只剩原始一段
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs).toHaveLength(1);
    expect(list.paragraphs[0]!.paraId).toBe('AAAAAAA1');
    await doc.dispose();
  });
});
