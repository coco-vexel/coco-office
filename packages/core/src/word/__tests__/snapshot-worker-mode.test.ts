/**
 * Snapshot × Worker mode 端到端测试 —— #22。
 *
 * 验证点(对齐 docs/03.2 §4 / §6):
 *   - worker 模式下 take 成功:kernel.serializeParts 跨 RPC 返回 parts
 *   - worker 模式下 rollback 成功:kernel.rebuildFromParts 跨 RPC 被调用
 *   - rollback 的 onProgress 能穿过 Worker 边界到达主线程(RpcProgress)
 *   - FIFO 淘汰在主线程发生,与 worker 无关(配 maxSnapshots 验证)
 *
 * 测试基础设施参考 worker-mode.test.ts:beforeAll 跑一次 core 定向 build,
 * 然后 `setWorkerBundleUrl(dist/kernel-worker.js)`。
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

const DOC_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CT}"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function docXml(paragraphs: string[]): string {
  const body = paragraphs.map((t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${body}</w:body>
</w:document>`;
}

async function makeZip(paragraphs: string[]): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(paragraphs));
  return zip.generateAsync({ type: 'uint8array' });
}

beforeAll(() => {
  execSync('corepack pnpm --filter @coco-vexel/coco-office exec tsup --no-clean', {
    cwd: rootDir,
    stdio: 'pipe',
  });
  WordDocument.setWorkerBundleUrl(bundleUrl);
}, 120_000);

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('snapshots worker mode — take / rollback 端到端', () => {
  test('worker 模式 take 返回 id,list 能看到', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['w-hello']), {
      docId: 'ws-1',
      mode: 'worker',
    });
    const id = await doc.snapshots.take('w-take');
    expect(id).toBe('snap-0');
    const infos = await doc.snapshots.list();
    expect(infos.map((i) => i.id)).toEqual([id]);
    expect(infos[0]?.reason).toBe('w-take');
    await doc.dispose();
  });

  test('worker 模式 rollback 成功:跨 RPC 调 rebuildFromParts,随后 toZip 可再往返', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['base']), {
      docId: 'ws-2',
      mode: 'worker',
    });
    const id = await doc.snapshots.take();
    await doc.snapshots.rollback(id);
    // rollback 成功后,该快照被消费
    expect(await doc.snapshots.list()).toEqual([]);
    // 重建后的 pkg 仍可 toZip / toFlatOpc
    const out = await doc.toZip();
    expect(out).toBeInstanceOf(Uint8Array);
    expect(out.length).toBeGreaterThan(0);
    await doc.dispose();
  });

  test('worker 模式 rollback onProgress 穿过 Worker 边界(parse / build-indices / done)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['a', 'b']), {
      docId: 'ws-3',
      mode: 'worker',
    });
    const id = await doc.snapshots.take();

    const phases: string[] = [];
    const pcts: Array<number | undefined> = [];
    await doc.snapshots.rollback(id, {
      onProgress: (phase, pct) => {
        phases.push(phase);
        pcts.push(pct);
      },
    });

    expect(phases).toContain('parse');
    expect(phases).toContain('build-indices');
    expect(phases).toContain('done');
    // done 最后,pct=100
    expect(phases[phases.length - 1]).toBe('done');
    expect(pcts[pcts.length - 1]).toBe(100);
    // 所有 pct 非负
    for (const p of pcts) {
      if (p !== undefined) {
        expect(p).toBeGreaterThanOrEqual(0);
        expect(p).toBeLessThanOrEqual(100);
      }
    }

    await doc.dispose();
  });

  test('worker 模式 FIFO(maxSnapshots=3):超限淘汰发生在主线程', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'ws-fifo',
      mode: 'worker',
      maxSnapshots: 3,
    });
    for (let i = 0; i < 5; i++) {
      await doc.snapshots.take();
    }
    const ids = (await doc.snapshots.list()).map((i) => i.id);
    expect(ids).toEqual(['snap-2', 'snap-3', 'snap-4']);
    await doc.dispose();
  });
});
