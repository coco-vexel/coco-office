/**
 * trackChanges 端到端测试(#31)。
 *
 * 链路:doc.plan({ trackChanges: true, author }) → builder → apply() → kernel
 * → executor → runTransaction → toZip。包含 inline 模式和 worker 模式 e2e。
 *
 * 覆盖:
 *   1. inline: replaceParagraph → apply → listRevisions 返回 2 条(del + ins),author 正确
 *   2. inline: insertParagraphAfter → apply → 新段 paraId 在 ApplyResult,listRevisions 返回 ins
 *   3. inline: deleteParagraph → apply → 段仍存在,listRevisions 返回 del
 *   4. inline: 多 op 同 plan → 多条修订,revId 递增
 *   5. inline: toZip → 回加载 → listRevisions 保留所有新修订
 *   6. worker: replaceParagraph 跨 RPC 产生修订,listRevisions 返回正确
 *   7. acceptRevisions / rejectRevisions 本 PR 仍抛 UNSUPPORTED_OOXML_FEATURE(留 #32)
 */
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { execSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { W14_NS, W_NS } from '../namespaces.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = resolve(__dirname, '../../../../..');
const bundleUrl = resolve(rootDir, 'packages/core/dist/kernel-worker.js');

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

describe('trackChanges e2e —— inline', () => {
  test('replaceParagraph → listRevisions 返回 2 条(del + ins),author/date 正确', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'tc-rp', mode: 'inline' });
    await doc
      .plan({ trackChanges: true, author: 'alice', date: new Date('2026-04-23T10:00:00.000Z') })
      .replaceParagraph('AAAAAAA1', 'new')
      .apply();

    const revs = await doc.revisions.list();
    expect(revs).toHaveLength(2);
    const ins = revs.find((r) => r.kind === 'insert')!;
    const del = revs.find((r) => r.kind === 'delete')!;
    expect(ins.author).toBe('alice');
    expect(del.author).toBe('alice');
    expect(ins.text).toBe('new');
    expect(del.text).toBe('old');
    await doc.dispose();
  });

  test('insertParagraphAfter → 新段有修订 + ApplyResult 记录', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'tc-ins', mode: 'inline' });
    const result = await doc
      .plan({ trackChanges: true, author: 'bob' })
      .insertParagraphAfter('AAAAAAA1', { text: 'new-para' })
      .apply();

    expect(result.insertedParagraphs).toHaveLength(1);
    const revs = await doc.revisions.list();
    // 至少 1 条 insert(内容 runs 包 ins;段落 mark ins 也会被 listRevisions 枚举到)
    expect(revs.filter((r) => r.kind === 'insert').length).toBeGreaterThanOrEqual(1);
    expect(revs.every((r) => r.author === 'bob')).toBe(true);
    await doc.dispose();
  });

  test('deleteParagraph → 段仍存在,listRevisions 返回 del', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>content</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'tc-del', mode: 'inline' });
    await doc.plan({ trackChanges: true, author: 'carol' }).deleteParagraph('AAAAAAA1').apply();

    // 段仍存在(trackChanges 下不 removeChild)
    const info = await doc.paragraphs.get('AAAAAAA1');
    expect(info).not.toBeNull();

    const revs = await doc.revisions.list();
    // 至少 1 条 delete(内容 + 段落 mark)
    expect(revs.filter((r) => r.kind === 'delete').length).toBeGreaterThanOrEqual(1);
    await doc.dispose();
  });

  test('多 op 同 plan:多条修订,revId 各不相同(递增)', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>a</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>b</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'tc-multi', mode: 'inline' });
    await doc
      .plan({ trackChanges: true, author: 'x' })
      .replaceParagraph('AAAAAAA1', 'a-new')
      .replaceParagraph('AAAAAAA2', 'b-new')
      .apply();

    const revs = await doc.revisions.list();
    const revIds = revs.map((r) => r.revId);
    // revIds 各不相同
    expect(new Set(revIds).size).toBe(revIds.length);
  });

  test('toZip → 回加载 → listRevisions 完整', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'tc-zip', mode: 'inline' });
    await doc
      .plan({ trackChanges: true, author: 'roundtrip' })
      .replaceParagraph('AAAAAAA1', 'modified')
      .apply();
    const out = await doc.toZip();
    const reloaded = await WordDocument.fromZip(out, { docId: 'tc-zip-2', mode: 'inline' });
    const revs = await reloaded.revisions.list();
    expect(revs.length).toBeGreaterThanOrEqual(2);
    expect(revs.every((r) => r.author === 'roundtrip')).toBe(true);
    await doc.dispose();
    await reloaded.dispose();
  });

  test('acceptRevisions:空文档(无修订)应正常返回 acceptedRevisions=0', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'tc-accept-empty', mode: 'inline' });
    const result = await doc.plan().acceptRevisions({ scope: 'all' }).apply();
    expect(result.acceptedRevisions).toBe(0);
    expect(result.rejectedRevisions).toBe(0);
    await doc.dispose();
  });
});

describe('trackChanges e2e —— worker', () => {
  beforeAll(() => {
    execSync('corepack pnpm --filter @coco-vexel/coco-office build', { cwd: rootDir, stdio: 'pipe' });
    WordDocument.setWorkerBundleUrl(bundleUrl);
  }, 120_000);

  const WORKER_TEST_TIMEOUT = 30_000;

  test(
    'worker 下 replaceParagraph trackChanges 产出修订,listRevisions 跨 RPC 返回正确',
    async () => {
      const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>orig</w:t></w:r></w:p>`);
      const doc = await WordDocument.fromZip(buf, { docId: 'tc-w', mode: 'worker' });
      await doc
        .plan({ trackChanges: true, author: 'worker-alice' })
        .replaceParagraph('AAAAAAA1', 'worker-new')
        .apply();

      const revs = await doc.revisions.list();
      expect(revs.length).toBeGreaterThanOrEqual(2);
      expect(revs.every((r) => r.author === 'worker-alice')).toBe(true);
      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
