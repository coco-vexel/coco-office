/**
 * WordPlan.apply —— worker 模式端到端(#29:段落结构操作)。
 *
 * 验证 insertParagraphAfter + toZip 产物在 Worker 模式下正确工作:
 *   - worker 下 insertParagraphAfter 生效 + getParagraph 返回新段文本
 *   - worker 下 deleteParagraph + toZip 产物不含被删段落的 paraId
 *
 * Builder / 纯 executor 覆盖见
 * [paragraph-structure-ops.test.ts](../plan-executors/__tests__/paragraph-structure-ops.test.ts) /
 * [plan-apply-paragraph-structure.test.ts](./plan-apply-paragraph-structure.test.ts)。
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

function makeDocXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

async function makeFixtureZip(bodyInner: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', makeDocXml(bodyInner));
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

describe('WordPlan apply —— worker 模式端到端 (#29)', () => {
  test(
    'worker 下 insertParagraphAfter 生效:getParagraph 返回新文本,toZip 含新 paraId',
    async () => {
      const buf = await makeFixtureZip(
        `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>`,
      );
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-w-ipa', mode: 'worker' });

      const result = await doc
        .plan()
        .insertParagraphAfter('AAAAAAA1', { text: 'worker-inserted' })
        .apply();

      expect(result.insertedParagraphs).toHaveLength(1);
      const newId = result.insertedParagraphs[0]!.newParaId;
      expect(newId).toMatch(/^[0-9A-F]{8}$/);

      const newPara = await doc.paragraphs.get(newId);
      expect(await newPara?.getText()).toBe('worker-inserted');

      // toZip 产物验证
      const out = await doc.toZip();
      const zip = await JSZip.loadAsync(out);
      const xml = await zip.file('word/document.xml')!.async('string');
      expect(xml).toContain('worker-inserted');
      expect(xml).toContain(newId);

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'worker 下 deleteParagraph 生效:toZip 产物不含被删 paraId',
    async () => {
      const buf = await makeFixtureZip(
        `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>del-me</w:t></w:r></w:p>
         <w:p w14:paraId="AAAAAAA2"><w:r><w:t>keep</w:t></w:r></w:p>`,
      );
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-w-dp', mode: 'worker' });

      await doc.plan().deleteParagraph('AAAAAAA1').apply();

      expect(await doc.paragraphs.get('AAAAAAA1')).toBeNull();
      expect(await doc.paragraphs.get('AAAAAAA2')).not.toBeNull();

      const out = await doc.toZip();
      const zip = await JSZip.loadAsync(out);
      const xml = await zip.file('word/document.xml')!.async('string');
      expect(xml).not.toContain('AAAAAAA1');
      expect(xml).toContain('AAAAAAA2');

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
