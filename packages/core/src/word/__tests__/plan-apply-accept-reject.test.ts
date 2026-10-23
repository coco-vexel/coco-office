/**
 * acceptRevisions / rejectRevisions 端到端测试(#32)。
 *
 * 走完整链路:doc.plan().acceptRevisions/rejectRevisions().apply() → kernel.applyPlan
 * → dispatchOp → runTransaction → toZip。验证与只读 API(#26 listRevisions)集成正确。
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

describe('acceptRevisions / rejectRevisions e2e —— inline', () => {
  test('accept all:listRevisions 清零,ApplyResult.acceptedRevisions 正确', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="alice" w:date="2026-04-23T00:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
        <w:del w:id="2" w:author="alice" w:date="2026-04-23T00:00:00Z">
          <w:r><w:delText>deleted</w:delText></w:r>
        </w:del>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'ar-accept-all', mode: 'inline' });
    const before = await doc.revisions.list();
    expect(before).toHaveLength(2);

    const result = await doc.plan().acceptRevisions({ scope: 'all' }).apply();
    expect(result.acceptedRevisions).toBe(2);
    expect(result.rejectedRevisions).toBe(0);

    const after = await doc.revisions.list();
    expect(after).toHaveLength(0);

    // 内容正确:inserted 保留,deleted 被删
    const info = await doc.paragraphs.get('AAAAAAA1');
    expect(await info?.getText()).toBe('inserted');
    await doc.dispose();
  });

  test('reject all:ins 删内容 / del 恢复内容', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>keep </w:t></w:r>
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z">
          <w:r><w:t>inserted </w:t></w:r>
        </w:ins>
        <w:del w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z">
          <w:r><w:delText>deleted</w:delText></w:r>
        </w:del>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'ar-reject-all', mode: 'inline' });
    const result = await doc.plan().rejectRevisions({ scope: 'all' }).apply();
    expect(result.rejectedRevisions).toBe(2);

    const info = await doc.paragraphs.get('AAAAAAA1');
    // ins 被 reject,inserted 丢;del 被 reject,deleted 恢复
    expect(await info?.getText()).toBe('keep deleted');
    await doc.dispose();
  });

  test('scope=paragraph:只影响目标段,其它段修订保留', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
      </w:p>
      <w:p w14:paraId="AAAAAAA2">
        <w:ins w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>b</w:t></w:r></w:ins>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'ar-scope-para', mode: 'inline' });
    await doc.plan().acceptRevisions({ scope: 'paragraph', paraId: 'AAAAAAA1' }).apply();
    const revs = await doc.revisions.list();
    // AAAAAAA2 的修订仍在
    expect(revs).toHaveLength(1);
    expect(revs[0]!.revId).toBe('2');
    await doc.dispose();
  });

  test('scope=revision:w:id 精确匹配,其它不动', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="5" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
        <w:ins w:id="6" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>b</w:t></w:r></w:ins>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'ar-scope-rev', mode: 'inline' });
    await doc.plan().acceptRevisions({ scope: 'revision', revId: '5' }).apply();
    const revs = await doc.revisions.list();
    expect(revs).toHaveLength(1);
    expect(revs[0]!.revId).toBe('6');
    await doc.dispose();
  });

  test('段落 mark reject ins → 撤销分段并合并下一段', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>keep</w:t></w:r></w:p>
      <w:p w14:paraId="AAAAAAA2">
        <w:pPr><w:rPr><w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>dropMe</w:t></w:r>
      </w:p><w:p w14:paraId="AAAAAAA3"><w:r><w:t>tail</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'ar-para-mark', mode: 'inline' });
    await doc.plan().rejectRevisions({ scope: 'all' }).apply();

    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await Promise.all(list.paragraphs.map((p) => p.getText()))).toEqual([
      'keep',
      'dropMetail',
    ]);
    await doc.dispose();
  });

  test('事务原子性:accept 一半 op 抛错 → 全回滚', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'ar-tx', mode: 'inline' });
    const plan = doc.plan().acceptRevisions({ scope: 'all' }).replaceParagraph('NOTEXIST', 'x'); // 第二 op 抛 PARA_NOT_FOUND
    await expect(plan.apply()).rejects.toMatchObject({ code: 'PARA_NOT_FOUND' });

    // 回滚:ins 仍在
    const revs = await doc.revisions.list();
    expect(revs).toHaveLength(1);
    await doc.dispose();
  });

  test('toZip → 回加载:accept 后 listRevisions 为空', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>content</w:t></w:r></w:ins>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'ar-zip', mode: 'inline' });
    await doc.plan().acceptRevisions({ scope: 'all' }).apply();
    const out = await doc.toZip();

    const reloaded = await WordDocument.fromZip(out, { docId: 'ar-zip-2', mode: 'inline' });
    expect(await reloaded.revisions.list()).toHaveLength(0);
    expect(await (await reloaded.paragraphs.get('AAAAAAA1'))?.getText()).toBe('content');
    await doc.dispose();
    await reloaded.dispose();
  });
});

describe('acceptRevisions / rejectRevisions e2e —— worker', () => {
  beforeAll(() => {
    execSync('corepack pnpm --filter @coco-vexel/coco-office build', {
      cwd: rootDir,
      stdio: 'pipe',
    });
    WordDocument.setWorkerBundleUrl(bundleUrl);
  }, 120_000);

  const WORKER_TEST_TIMEOUT = 30_000;

  test(
    'worker 下 accept all 正常工作,listRevisions 跨 RPC 清零',
    async () => {
      const buf = await makeFixture(
        `<w:p w14:paraId="AAAAAAA1">
          <w:ins w:id="1" w:author="w" w:date="2026-04-23T00:00:00Z"><w:r><w:t>x</w:t></w:r></w:ins>
        </w:p>`,
      );
      const doc = await WordDocument.fromZip(buf, { docId: 'ar-w', mode: 'worker' });
      const result = await doc.plan().acceptRevisions({ scope: 'all' }).apply();
      expect(result.acceptedRevisions).toBe(1);
      expect(await doc.revisions.list()).toHaveLength(0);
      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
