/**
 * WordPlan.apply —— worker 模式端到端(#28)。
 *
 * 验证 #28 executor 通过 RPC 正常工作:
 *   - inline / worker 一致性:同 fixture + 同 plan → 同 toZip 产物
 *   - apply 后 getParagraph 通过 worker RPC 返回新文本
 *   - trackChanges=true 的 UNSUPPORTED_OOXML_FEATURE 跨 RPC 边界完整传回
 *
 * Builder / 纯 executor 覆盖见 [plan-apply-paragraph.test.ts](./plan-apply-paragraph.test.ts) /
 * [paragraph-ops.test.ts](../plan-executors/__tests__/paragraph-ops.test.ts)。
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
    <w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>
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

describe('WordPlan apply —— worker 模式端到端 (#28)', () => {
  test(
    'worker 下 replaceParagraph 生效 + getParagraph 返回新文本',
    async () => {
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-w-rp', mode: 'worker' });
      await doc.plan().replaceParagraph('AAAAAAA1', 'worker-new').apply();
      const info = await doc.paragraphs.get('AAAAAAA1');
      expect(await info?.getText()).toBe('worker-new');
      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'inline / worker 一致性:相同 plan 产出相同 document.xml',
    async () => {
      const buf = await makeFixtureZip();
      const docI = await WordDocument.fromZip(buf, { docId: 'e2e-i', mode: 'inline' });
      const docW = await WordDocument.fromZip(buf, { docId: 'e2e-w', mode: 'worker' });
      await docI.plan().replaceTextInParagraph('AAAAAAA1', 'original', 'same').apply();
      await docW.plan().replaceTextInParagraph('AAAAAAA1', 'original', 'same').apply();
      const xmlI = await (await JSZip.loadAsync(await docI.toZip()))
        .file('word/document.xml')!
        .async('string');
      const xmlW = await (await JSZip.loadAsync(await docW.toZip()))
        .file('word/document.xml')!
        .async('string');
      expect(xmlI).toBe(xmlW);
      expect(xmlI).toContain('same');
      await docI.dispose();
      await docW.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  // trackChanges=true 的 worker e2e 见 plan-apply-trackChanges.test.ts(#31 已实现包装层)
});
