/**
 * Worker crash + reload —— #19 (docs/03 §4.7)。
 *
 * 验证:
 *   1. Worker 崩溃模拟(通过 `WorkerTransport._triggerCrashForTest` 触发 handleWorkerError)
 *   2. 崩溃时 in-flight RPC 以 `WordError{ code: 'WORKER_RELOADED' }` reject
 *   3. reload 后的新 RPC 成功(证明 reload 工作)
 *   4. 多次崩溃各自独立 reload
 *   5. reload 期间 queued 的 RPC 最终被 flush 执行
 *   6. Inline 模式永远不触发 WORKER_RELOADED
 */
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'vitest';
import { execSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
import { W14_NS, W_NS } from '../namespaces.js';
import type { WorkerTransport } from '../worker-transport.js';

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
    <w:p><w:r><w:t>reload-fixture</w:t></w:r></w:p>
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

afterEach(() => {
  // 不 reset workerBundleUrl,保留 beforeAll 的设置
});

describe('Worker crash + reload (#19)', () => {
  test('in-flight RPC 在崩溃时以 WORKER_RELOADED reject', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'reload-1', mode: 'worker' });

    // 发起一个 toZip RPC,但不 await —— 抓到 promise 后立即 trigger crash
    const inflight = doc.toZip();
    const transport = doc._getTransport() as WorkerTransport;
    transport._triggerCrashForTest(new Error('boom'));

    await expect(inflight).rejects.toBeInstanceOf(WordError);
    try {
      await inflight;
    } catch (e) {
      const err = e as WordError;
      expect(err.code).toBe('WORKER_RELOADED');
      expect(err.docId).toBe('reload-1');
      expect(err.details).toMatchObject({ method: 'toZip' });
      // cause 注入:包含原 err 的 message
      expect(String(err.details.cause)).toContain('boom');
    }

    await doc.dispose();
  });

  test('reload 后下一个 RPC 成功(证明 reload 工作)', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'reload-2', mode: 'worker' });

    const inflight = doc.toZip();
    const transport = doc._getTransport() as WorkerTransport;
    transport._triggerCrashForTest();

    // 先吃掉 inflight 的 WORKER_RELOADED
    await expect(inflight).rejects.toBeInstanceOf(WordError);

    // reload 完成后再发 RPC:应正常返回
    const out = await doc.toZip();
    expect(out).toBeInstanceOf(Uint8Array);
    expect(out.length).toBeGreaterThan(0);

    await doc.dispose();
  });

  test('reload 期间 queued 的 RPC 在 reload 后按序执行', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'reload-3', mode: 'worker' });

    const transport = doc._getTransport() as WorkerTransport;

    // 1. 发起 in-flight RPC
    const inflight = doc.toZip();

    // 2. 模拟 crash —— state 切到 reloading
    transport._triggerCrashForTest();

    // 3. state == reloading,新的 call 进入 queue
    expect(transport._getStateForTest()).toBe('reloading');
    const queued1 = doc.toZip();
    const queued2 = doc.toFlatOpc();

    // inflight 必定 reject
    await expect(inflight).rejects.toBeInstanceOf(WordError);

    // queued 全部应成功 resolve(reload 完成后 flush)
    const [z, x] = await Promise.all([queued1, queued2]);
    expect(z).toBeInstanceOf(Uint8Array);
    expect(typeof x).toBe('string');
    expect(x).toMatch(/reload-fixture/);
    expect(transport._getStateForTest()).toBe('idle');

    await doc.dispose();
  });

  test('多次 crash 可各自独立 reload', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'reload-multi', mode: 'worker' });
    const transport = doc._getTransport() as WorkerTransport;

    for (let i = 0; i < 3; i++) {
      const inflight = doc.toZip();
      transport._triggerCrashForTest(new Error(`crash-${i}`));
      await expect(inflight).rejects.toBeInstanceOf(WordError);
      // reload 后 RPC 成功
      const out = await doc.toZip();
      expect(out).toBeInstanceOf(Uint8Array);
      expect(transport._getStateForTest()).toBe('idle');
    }

    await doc.dispose();
  });

  test('reload 后基于初始字节,pkg 恢复成初始状态(同 bytes 可再 roundtrip)', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'reload-roundtrip', mode: 'worker' });
    const transport = doc._getTransport() as WorkerTransport;

    // 触发一次 reload
    const inflight = doc.toZip();
    transport._triggerCrashForTest();
    await expect(inflight).rejects.toBeInstanceOf(WordError);

    // reload 后 toFlatOpc 应该包含初始字节 parse 出来的 para
    const xml = await doc.toFlatOpc();
    expect(xml).toMatch(/reload-fixture/);

    await doc.dispose();
  });

  test('Inline 模式:WORKER_RELOADED 永远不触发', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'reload-inline', mode: 'inline' });
    expect(doc.mode).toBe('inline');

    // inline transport 没有 _triggerCrashForTest(WorkerTransport 独有)
    const transport = doc._getTransport();
    // 类型防御:inline transport 不应暴露 _triggerCrashForTest
    expect(
      (transport as unknown as { _triggerCrashForTest?: () => void })._triggerCrashForTest,
    ).toBeUndefined();

    // RPC 正常执行不会产出 WORKER_RELOADED
    const out = await doc.toZip();
    expect(out).toBeInstanceOf(Uint8Array);

    await doc.dispose();
  });
});
