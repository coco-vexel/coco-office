/**
 * listComments / getComment / listRevisions —— worker 模式端到端(#26)。
 *
 * 验证 3 个方法在 RPC 路径下参数 / 结果跨 Worker 边界的完整性,与 inline 模式一致。
 * 纯函数覆盖见 [comment-queries.test.ts](./comment-queries.test.ts) /
 * [revision-queries.test.ts](./revision-queries.test.ts)。
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
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="AAAAAAA1">
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>annotated</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
    </w:p>
    <w:p w14:paraId="AAAAAAA2">
      <w:ins w:id="10" w:author="alice" w:date="2026-04-22T10:00:00Z">
        <w:r><w:t>inserted</w:t></w:r>
      </w:ins>
      <w:del w:id="20" w:author="bob" w:date="2026-04-22T11:00:00Z">
        <w:r><w:delText>deleted</w:delText></w:r>
      </w:del>
    </w:p>
  </w:body>
</w:document>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W_NS}">
  <w:comment w:id="1" w:author="reviewer" w:date="2026-04-22T09:00:00Z">
    <w:p><w:r><w:t>此处存疑</w:t></w:r></w:p>
  </w:comment>
</w:comments>`;

async function makeFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  zip.file('word/document.xml', DOC_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
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

describe('comments / revisions —— worker 模式端到端 (#26)', () => {
  test(
    'inline / worker:listComments 结果等价,字段完整',
    async () => {
      const buf = await makeFixtureZip();
      const docI = await WordDocument.fromZip(buf, { docId: 'cr-i', mode: 'inline' });
      const docW = await WordDocument.fromZip(buf, { docId: 'cr-w', mode: 'worker' });

      const ci = await docI.comments.list();
      const cw = await docW.comments.list();

      expect(await Promise.all(cw.map((comment) => comment.getInfo()))).toEqual(
        await Promise.all(ci.map((comment) => comment.getInfo())),
      );
      expect(ci).toHaveLength(1);
      const c = ci[0]!;
      expect(c.id).toBe(1);
      expect(c.author).toBe('reviewer');
      expect(c.text).toBe('此处存疑');
      expect(c.rangeKind).toBe('paragraph');
      expect(c.paraId).toBe('AAAAAAA1');

      await docI.dispose();
      await docW.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'getComment RPC:命中 + 未命中 null',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'cr-w-2', mode: 'worker' });

      const hit = await doc.comments.get(1);
      expect(hit?.author).toBe('reviewer');
      expect(await doc.comments.get(999)).toBeNull();

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'listRevisions 默认 scope=body:跨 RPC 返回 ins + del,字段完整',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'cr-w-3', mode: 'worker' });

      const revs = await doc.revisions.list();
      expect(revs).toHaveLength(2);
      const ins = revs.find((r) => r.kind === 'insert')!;
      const del = revs.find((r) => r.kind === 'delete')!;
      expect(ins.revId).toBe('10');
      expect(ins.text).toBe('inserted');
      expect(ins.story).toBe('body');
      expect(ins.paraId).toBe('AAAAAAA2');
      expect(del.revId).toBe('20');
      expect(del.text).toBe('deleted');
      expect(del.author).toBe('bob');

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
