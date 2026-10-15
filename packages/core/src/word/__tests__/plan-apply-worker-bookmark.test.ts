/**
 * WordPlan.apply —— worker 模式端到端 书签(#34)。
 *
 * 验证 createBookmark / deleteBookmark 通过 RPC 正常工作:
 *   - worker 模式 createBookmark → toZip → 回加载 listBookmarks 有新书签
 *   - worker 模式 deleteBookmark → listBookmarks 不含该名
 *
 * 纯函数 / inline 端到端覆盖见 bookmark-ops.test.ts / plan-apply-bookmark.test.ts。
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

beforeAll(() => {
  execSync('corepack pnpm --filter @coco-vexel/coco-office build', { cwd: rootDir, stdio: 'pipe' });
  WordDocument.setWorkerBundleUrl(bundleUrl);
}, 120_000);

afterAll(() => {
  WordDocument._resetDefaults();
});

const WORKER_TEST_TIMEOUT = 30_000;

describe('WordPlan apply —— worker 模式书签 (#34)', () => {
  test(
    'worker createBookmark → toZip → 回加载 listBookmarks 有新书签',
    async () => {
      const buf = await makeFixture(
        `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>paragraph</w:t></w:r></w:p>`,
      );
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-w-bm', mode: 'worker' });
      const result = await doc
        .plan()
        .createBookmark({ name: 'workerBm', range: { paraId: 'AAAAAAA1' } })
        .apply();
      expect(result.createdBookmarks).toEqual([
        { opIndex: 0, name: 'workerBm', range: { fromParaId: 'AAAAAAA1', toParaId: 'AAAAAAA1' } },
      ]);

      // toZip → 回加载验证书签持久化
      const out = await doc.toZip();
      const doc2 = await WordDocument.fromZip(out, { docId: 'e2e-w-bm-reload', mode: 'inline' });
      const bookmarks = await doc2.bookmarks.list();
      const bm = bookmarks.find((b) => b.name === 'workerBm');
      expect(bm).toBeDefined();
      expect(bm!.rangeKind).toBe('paragraph');

      await doc.dispose();
      await doc2.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'worker deleteBookmark → listBookmarks 不含该名',
    async () => {
      const buf = await makeFixture(
        `<w:p w14:paraId="AAAAAAA1">
          <w:bookmarkStart w:id="0" w:name="delMe"/>
          <w:r><w:t>text</w:t></w:r>
          <w:bookmarkEnd w:id="0"/>
        </w:p>`,
      );
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-w-delbm', mode: 'worker' });
      await doc.plan().deleteBookmark('delMe').apply();

      const bookmarks = await doc.bookmarks.list();
      expect(bookmarks.find((b) => b.name === 'delMe')).toBeUndefined();
      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
