/**
 * listParagraphs —— worker 模式端到端(#24)。
 *
 * 目的:
 *   - 验证 RPC 通道下 listParagraphs 的参数透传、结果反序列化正确
 *   - 覆盖 inline + worker 两种模式下的一致性(对同一 fixture 结果应相同)
 *   - 非法参数透过 RPC 转为 WordError(kernel 内部 try/catch 包装)
 *
 * 纯函数的完整覆盖在 [list-paragraphs.test.ts](./list-paragraphs.test.ts)。
 */
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { execSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
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

function buildDocXml(paraCount: number): string {
  const paras = Array.from({ length: paraCount }, (_, i) => {
    const pid = String(i).padStart(8, 'A').slice(-8).toUpperCase();
    return `<w:p w14:paraId="${pid}"><w:r><w:t>p-${i}</w:t></w:r></w:p>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${paras}</w:body>
</w:document>`;
}

async function makeFixtureZip(paraCount: number): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', buildDocXml(paraCount));
  return zip.generateAsync({ type: 'uint8array' });
}

beforeAll(() => {
  execSync('corepack pnpm --filter @coco-vexel/coco-office build', { cwd: rootDir, stdio: 'pipe' });
  WordDocument.setWorkerBundleUrl(bundleUrl);
}, 120_000);

afterAll(() => {
  WordDocument._resetDefaults();
});

// worker_threads 并发压力下超时放宽(和 worker-reload.test.ts 一致)
const WORKER_TEST_TIMEOUT = 30_000;

describe('listParagraphs —— worker 模式端到端 (#24)', () => {
  test(
    '默认 pagination byCount(25,0):worker 返回前 25 条 + totalPages',
    async () => {
      const buf = await makeFixtureZip(60);
      const doc = await WordDocument.fromZip(buf, { docId: 'lp-w-1', mode: 'worker' });

      const r = await doc.paragraphs.list();
      expect(r.paragraphs).toHaveLength(25);
      expect(r.pagination.mode).toBe('byCount');
      expect(r.pagination.totalParagraphs).toBe(60);
      expect(r.pagination.totalPages).toBe(3);
      expect(r.pagination.hasNext).toBe(true);
      expect(r.paragraphs[0]!.paraId).toMatch(/^[A-F0-9]{8}$/);

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'inline 与 worker 在同一 fixture 上返回等价结果',
    async () => {
      const buf = await makeFixtureZip(10);
      const docInline = await WordDocument.fromZip(buf, { docId: 'lp-i', mode: 'inline' });
      const docWorker = await WordDocument.fromZip(buf, { docId: 'lp-w', mode: 'worker' });

      const ri = await docInline.paragraphs.list({ pagination: { mode: 'all' } });
      const rw = await docWorker.paragraphs.list({ pagination: { mode: 'all' } });

      // paragraphs 数组一致(paraId / text / story 都应完全相同)
      expect(ri.paragraphs.map((p) => p.paraId)).toEqual(rw.paragraphs.map((p) => p.paraId));
      expect(await Promise.all(ri.paragraphs.map((p) => p.getText()))).toEqual(
        await Promise.all(rw.paragraphs.map((p) => p.getText())),
      );
      expect(ri.pagination).toEqual(rw.pagination);

      await docInline.dispose();
      await docWorker.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'worker 模式下非法 pageSize 透传回主线程为 WordError',
    async () => {
      const buf = await makeFixtureZip(3);
      const doc = await WordDocument.fromZip(buf, { docId: 'lp-w-err', mode: 'worker' });

      await expect(
        doc.paragraphs.list({ pagination: { mode: 'byCount', pageSize: 0, pageIndex: 0 } }),
      ).rejects.toBeInstanceOf(WordError);

      // 详细类型:code 来自 kernel 的 catch(UNSUPPORTED_OOXML_FEATURE)
      try {
        await doc.paragraphs.list({ pagination: { mode: 'byCount', pageSize: 0, pageIndex: 0 } });
      } catch (e) {
        const err = e as WordError;
        expect(err.code).toBe('UNSUPPORTED_OOXML_FEATURE');
        expect(err.docId).toBe('lp-w-err');
        expect(err.message).toMatch(/pageSize/);
      }

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
