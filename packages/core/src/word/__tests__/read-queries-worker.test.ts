/**
 * getParagraph / listBookmarks / getBookmark —— worker 模式端到端(#25)。
 *
 * 验证 RPC 参数透传与结果反序列化:
 *   - inline / worker 两种模式在同一 fixture 上输出等价
 *   - ParagraphInfo / BookmarkInfo 跨 RPC 边界字段完整
 *   - 未命中返回 null(而非报错)
 *
 * 纯函数完整覆盖在 [read-queries.test.ts](./read-queries.test.ts)。
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

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="AAAAAAA1">
      <w:bookmarkStart w:id="1" w:name="visible"/>
      <w:r><w:t>entire</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
    <w:p w14:paraId="AAAAAAA2">
      <w:r><w:t>pre </w:t></w:r>
      <w:bookmarkStart w:id="2" w:name="middle"/>
      <w:r><w:t>mid</w:t></w:r>
      <w:bookmarkEnd w:id="2"/>
      <w:r><w:t> post</w:t></w:r>
    </w:p>
    <w:p w14:paraId="AAAAAAA3">
      <w:bookmarkStart w:id="3" w:name="_Toc42"/>
      <w:r><w:t>toc</w:t></w:r>
      <w:bookmarkEnd w:id="3"/>
    </w:p>
  </w:body>
</w:document>`;

async function makeFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', DOC_XML);
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

describe('read queries —— worker 模式端到端 (#25)', () => {
  test(
    'getParagraph:命中返回 ParagraphInfo,未命中返回 null',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'rq-w-1', mode: 'worker' });

      const hit = await doc.paragraphs.get('AAAAAAA2');
      expect(hit).not.toBeNull();
      expect(hit!.paraId).toBe('AAAAAAA2');
      expect(await hit!.getText()).toBe('pre mid post');
      const info = await hit!.getInfo();
      expect(info?.story).toBe('body');
      expect(info?.hasBookmarks).toBe(true);

      const miss = await doc.paragraphs.get('NOTEXIST');
      expect(miss).toBeNull();

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'inline / worker 一致性:listBookmarks 默认与 includeHidden',
    async () => {
      const buf = await makeFixtureZip();
      const docI = await WordDocument.fromZip(buf, { docId: 'rq-i', mode: 'inline' });
      const docW = await WordDocument.fromZip(buf, { docId: 'rq-w', mode: 'worker' });

      const [iNames, wNames] = [
        (await docI.bookmarks.list()).map((b) => b.name),
        (await docW.bookmarks.list()).map((b) => b.name),
      ];
      expect(iNames).toEqual(['visible', 'middle']); // `_Toc42` 被默认过滤
      expect(wNames).toEqual(iNames);

      const iAll = (await docI.bookmarks.list({ includeHidden: true })).map((b) => b.name);
      const wAll = (await docW.bookmarks.list({ includeHidden: true })).map((b) => b.name);
      expect(iAll).toEqual(['visible', 'middle', '_Toc42']);
      expect(wAll).toEqual(iAll);

      await docI.dispose();
      await docW.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'getBookmark(name):hidden 仍返回,未命中 null',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'rq-w-2', mode: 'worker' });

      const toc = await doc.bookmarks.get('_Toc42');
      expect(toc).not.toBeNull();
      expect(toc!.hidden).toBe(true);
      expect(toc!.rangeKind).toBe('paragraph');
      expect(toc!.paraId).toBe('AAAAAAA3');

      const mid = await doc.bookmarks.get('middle');
      expect(mid!.rangeKind).toBe('text');
      expect(mid!.text).toBe('mid');

      expect(await doc.bookmarks.get('no-such')).toBeNull();

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
