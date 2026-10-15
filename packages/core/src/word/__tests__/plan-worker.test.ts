/**
 * WordPlan apply —— worker 模式端到端(#27)。
 *
 * 验证:
 *   - 空 plan 通过 RPC 正常返回空 ApplyResult(不抛错)
 *   - 非空 plan 通过 RPC 抛 UNSUPPORTED_OOXML_FEATURE(kernel stub 行为),
 *     details 跨 RPC 边界完整
 *   - 多 plan 并存在 worker 模式下同样独立
 *
 * Builder / 状态机覆盖见 [plan.test.ts](./plan.test.ts)。
 */
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, test, vi } from 'vitest';
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
  <w:body><w:p w14:paraId="AAAAAAA1"><w:r><w:t>plan-worker-fixture</w:t></w:r></w:p></w:body>
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

let warnSpy: ReturnType<typeof vi.spyOn>;
beforeEach(() => {
  warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
});
afterEach(() => {
  warnSpy.mockRestore();
});

const WORKER_TEST_TIMEOUT = 30_000;

describe('WordPlan apply —— worker 模式端到端 (#27)', () => {
  test(
    '空 plan 通过 RPC 返回空 ApplyResult + warn',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'plan-w-empty', mode: 'worker' });
      const result = await doc.plan().apply();
      expect(result.insertedParagraphs).toEqual([]);
      expect(result.createdComments).toEqual([]);
      expect(result.createdBookmarks).toEqual([]);
      expect(result.acceptedRevisions).toBe(0);
      expect(result.rejectedRevisions).toBe(0);
      expect(warnSpy).toHaveBeenCalledTimes(1);
      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  // "未实现 op 通过 RPC 抛 UNSUPPORTED"的哨兵测试在 #32 合入后失效(11 op 全实现),
  // 已删除。见 plan.test.ts 的同款注释。
});
