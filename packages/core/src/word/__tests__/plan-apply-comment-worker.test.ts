/**
 * WordPlan.apply —— addComment worker 模式端到端(#35)。
 *
 * 验证 #35 executor 通过 worker RPC 正常工作:
 *   - worker 下 addComment → toZip → 回加载后 listComments 有该 id
 *   - worker 模式 deleteComment 正常删除
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

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="AAAAAAA1"><w:r><w:t>worker comment test</w:t></w:r></w:p>
  </w:body>
</w:document>`;

async function makeFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', DOC_XML);
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  return zip.generateAsync({ type: 'uint8array' });
}

beforeAll(() => {
  execSync('corepack pnpm --filter @coco-vexel/coco-office build', { cwd: rootDir, stdio: 'pipe' });
  WordDocument.setWorkerBundleUrl(bundleUrl);
}, 120_000);

afterAll(() => {
  WordDocument._resetDefaults();
});

const WORKER_TEST_TIMEOUT = 30_000;

describe('WordPlan apply —— addComment / deleteComment worker 模式 (#35)', () => {
  test(
    'worker 下 validatePlan 返回全部目标错误且不写入',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-w-validate', mode: 'worker' });
      const validation = await doc
        .plan()
        .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'ok' })
        .addComment({ range: { paraId: 'NOTFOUND' }, text: 'bad' })
        .validate();

      expect(validation.valid).toBe(false);
      expect(validation.items.map((item) => item.status)).toEqual(['valid', 'invalid']);
      expect(validation.items[1]?.issue?.code).toBe('PARA_NOT_FOUND');
      expect(await doc.comments.list()).toEqual([]);
      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'worker 下 addComment → toZip → 回加载后 listComments 有该 id',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-w-ac', mode: 'worker' });

      const result = await doc
        .plan({ author: 'worker-user' })
        .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'worker comment' })
        .apply();

      expect(result.createdComments).toHaveLength(1);
      const commentId = result.createdComments[0]!.commentId;

      const zipBytes = await doc.toZip();
      await doc.dispose();

      // 重新加载验证持久化
      const doc2 = await WordDocument.fromZip(zipBytes, {
        docId: 'e2e-w-ac-reload',
        mode: 'inline',
      });
      const comments = await doc2.comments.list();
      expect(comments.some((c) => c.id === commentId && c.text === 'worker comment')).toBe(true);
      await doc2.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'worker 下 addComment + deleteComment:最终 listComments 为空',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-w-dc', mode: 'worker' });

      // 先 add 一个
      const r1 = await doc
        .plan()
        .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'to be deleted' })
        .apply();
      const id = r1.createdComments[0]!.commentId;

      // 再 delete
      await doc.plan().deleteComment(id).apply();

      const comments = await doc.comments.list();
      expect(comments).toHaveLength(0);
      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
