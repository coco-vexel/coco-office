/**
 * Snapshot + Worker reload —— #23(docs/03.2 §9.1)。
 *
 * 验证:
 *   1. crash 前 take 过 N 个快照 → adapter 里有字节;crash → reload 后,take 的新 id
 *      继续在 `max(snap-<N>) + 1`,不与 adapter 里已有 id 撞车。
 *   2. in-flight take 在 crash 时被 reject 为 `WordError{ WORKER_RELOADED }`(沿用 #19)。
 *   3. crash 前 recordChange 累加的 changelog:crash 后 `changes()` 返回空数组
 *      (语义丢失,与 worker 侧 DOM 一起被 reload 的初始字节替代)。
 *   4. 已持久化到 adapter 的字节在 reload 后仍可用 rollback。
 *   5. Inline 模式:SnapshotApi 无 reload 语义,`reinitializeAfterReload` 可手调但
 *      不应被 transport 自动触发(InlineTransport 无 onReload)。
 */
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'vitest';
import { execSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { InMemorySnapshotStorage } from '../../shared/index.js';
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
    <w:p><w:r><w:t>snapshot-reload-fixture</w:t></w:r></w:p>
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
  execSync('corepack pnpm --filter @coco-vexel/coco-office exec tsup --no-clean', {
    cwd: rootDir,
    stdio: 'pipe',
  });
  WordDocument.setWorkerBundleUrl(bundleUrl);
}, 120_000);

afterAll(() => {
  WordDocument._resetDefaults();
});

afterEach(() => {
  // 保留 beforeAll 配置的 workerBundleUrl
});

// 并行 worker_threads 压力下 5s 默认超时偶发过短,放宽到 30s 防 CI flake
const WORKER_TEST_TIMEOUT = 30_000;

describe('Worker reload 后 SnapshotApi 内存态重建 (#23)', () => {
  test(
    'crash 前 3 个快照 → reload 后 take 继续 snap-3(不撞车)',
    async () => {
      const adapter = new InMemorySnapshotStorage();
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, {
        docId: 'snap-reload-1',
        mode: 'worker',
        snapshotStorage: adapter,
      });

      const s0 = await doc.snapshots.take('step-0');
      const s1 = await doc.snapshots.take('step-1');
      const s2 = await doc.snapshots.take('step-2');
      expect([s0, s1, s2]).toEqual(['snap-0', 'snap-1', 'snap-2']);
      // adapter 里字节都在
      expect((await adapter.list('snap-reload-1')).map((i) => i.id).sort()).toEqual([
        'snap-0',
        'snap-1',
        'snap-2',
      ]);

      // 触发 crash → reload
      const transport = doc._getTransport() as WorkerTransport;
      const inflight = doc.toZip();
      transport._triggerCrashForTest(new Error('boom'));
      await expect(inflight).rejects.toBeInstanceOf(WordError);

      // reload 完成后(下一个 RPC 走完即可确认),新的 take 应为 snap-3
      // 这里 take 本身会在 reload 期间 queued,flush 时 transport 已先 await 过
      // reinitializeAfterReload,nextId 此时已指向 3。
      const s3 = await doc.snapshots.take('step-3-after-reload');
      expect(s3).toBe('snap-3');

      // adapter 里 4 个都在,且 id 无冲突
      const ids = (await adapter.list('snap-reload-1')).map((i) => i.id).sort();
      expect(ids).toEqual(['snap-0', 'snap-1', 'snap-2', 'snap-3']);

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'in-flight take 在 crash 时以 WORKER_RELOADED reject',
    async () => {
      const adapter = new InMemorySnapshotStorage();
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, {
        docId: 'snap-reload-inflight',
        mode: 'worker',
        snapshotStorage: adapter,
      });

      const transport = doc._getTransport() as WorkerTransport;

      // 发起 take 但不 await —— take 内部首先 await transport.call('serializeParts'),
      // 这个 RPC 就是 in-flight 的部分。crash 应把它以 WORKER_RELOADED reject。
      const inflight = doc.snapshots.take('inflight');
      transport._triggerCrashForTest(new Error('crash during take'));

      await expect(inflight).rejects.toBeInstanceOf(WordError);
      try {
        await inflight;
      } catch (e) {
        const err = e as WordError;
        expect(err.code).toBe('WORKER_RELOADED');
        expect(err.docId).toBe('snap-reload-inflight');
        expect(err.details).toMatchObject({ method: 'serializeParts' });
      }

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'in-flight rollback 在 crash 时以 WORKER_RELOADED reject',
    async () => {
      const adapter = new InMemorySnapshotStorage();
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, {
        docId: 'snap-reload-rollback-inflight',
        mode: 'worker',
        snapshotStorage: adapter,
      });

      // 先拍一个快照,保证 rollback 有东西可回
      const snapId = await doc.snapshots.take('to-rollback');
      expect(snapId).toBe('snap-0');

      const transport = doc._getTransport() as WorkerTransport;

      // 发起 rollback 但不 await;先让 adapter.load 完成,再等下一 microtask
      // 让 transport.call('rebuildFromParts') 进入 pending 后再 crash,确保 RPC in-flight。
      const inflight = doc.snapshots.rollback(snapId);
      await Promise.resolve(); // 让 adapter.load 的 await 完成,transport.call 进入 pending 后再 crash
      transport._triggerCrashForTest(new Error('crash during rollback'));

      await expect(inflight).rejects.toBeInstanceOf(WordError);
      try {
        await inflight;
      } catch (e) {
        expect((e as WordError).code).toBe('WORKER_RELOADED');
        expect((e as WordError).details).toMatchObject({ method: 'rebuildFromParts' });
      }

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'crash 前 recordChange 的 changelog 在 reload 后丢失',
    async () => {
      const adapter = new InMemorySnapshotStorage();
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, {
        docId: 'snap-reload-changelog',
        mode: 'worker',
        snapshotStorage: adapter,
      });

      doc.snapshots.recordChange({ toolName: 'replace_paragraphs', summary: { paraId: 'A' } });
      doc.snapshots.recordChange({ toolName: 'add_comment' });
      expect(doc.snapshots.changes()).toHaveLength(2);

      // crash → reload
      const transport = doc._getTransport() as WorkerTransport;
      const inflight = doc.toZip();
      transport._triggerCrashForTest();
      await expect(inflight).rejects.toBeInstanceOf(WordError);

      // 等 reload 完成 —— 再跑一次 RPC 确保 state 回 idle,reinitializeAfterReload 已执行
      await doc.toZip();
      expect(transport._getStateForTest()).toBe('idle');

      // changelog 清零,changes() 返回空
      expect(doc.snapshots.changes()).toEqual([]);

      // 再 take 的 changesCaptured 也应是空(reload 之后没有新 recordChange)
      const id = await doc.snapshots.take();
      const info = (await doc.snapshots.list()).find((i) => i.id === id);
      expect(info?.changesCaptured).toEqual([]);

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'rollback 在 reload 后仍可用(字节保留在 adapter)',
    async () => {
      const adapter = new InMemorySnapshotStorage();
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, {
        docId: 'snap-reload-rollback',
        mode: 'worker',
        snapshotStorage: adapter,
      });

      // crash 前拍一个快照,里面有初始文本
      const snapId = await doc.snapshots.take('before-crash');
      expect(snapId).toBe('snap-0');

      // crash → reload
      const transport = doc._getTransport() as WorkerTransport;
      const inflight = doc.toZip();
      transport._triggerCrashForTest();
      await expect(inflight).rejects.toBeInstanceOf(WordError);

      // reload 完成后 adapter 里仍有 snap-0
      // 发起 rollback —— 会在 queued 中 flush;transport 已先 await reinitializeAfterReload,
      // 但 rollback 不依赖 nextId,只需要 adapter.load 能拿到 payload
      await doc.snapshots.rollback(snapId);

      // rollback 后,toFlatOpc 内容仍含初始文本(快照内容就是初始字节 parse 出的 DOM)
      const xml = await doc.toFlatOpc();
      expect(xml).toMatch(/snapshot-reload-fixture/);

      // rollback 消费语义:adapter 里的 snap-0 已被删
      const remaining = (await adapter.list('snap-reload-rollback')).map((i) => i.id);
      expect(remaining).not.toContain(snapId);

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'reload 后 nextId 收缩:只看 adapter 当前剩余的 id(rollback/delete 过的不再占用)',
    async () => {
      // 覆盖 reinitialize 的可观测收缩行为:
      //   crash 前 take 3 次(snap-0/1/2,nextId=3),然后删掉 snap-2 —— 主线程侧
      //   nextId 仍然是 3(没有"归还"语义)。crash → reload → reinitialize 读 adapter,
      //   adapter 里只剩 snap-0/1,max+1=2,nextId 收缩到 2。下一个 take 是 snap-2。
      const adapter = new InMemorySnapshotStorage();
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, {
        docId: 'snap-reload-shrink',
        mode: 'worker',
        snapshotStorage: adapter,
      });

      await doc.snapshots.take();
      await doc.snapshots.take();
      await doc.snapshots.take();
      await doc.snapshots.delete('snap-2');
      // 确认 adapter 里只剩 snap-0/1
      expect((await adapter.list('snap-reload-shrink')).map((i) => i.id).sort()).toEqual([
        'snap-0',
        'snap-1',
      ]);

      // crash → reload
      const transport = doc._getTransport() as WorkerTransport;
      const inflight = doc.toZip();
      transport._triggerCrashForTest();
      await expect(inflight).rejects.toBeInstanceOf(WordError);

      // 下一个 take 的 id 应收缩到 snap-2(max(0,1)+1)
      const next = await doc.snapshots.take('after-shrink');
      expect(next).toBe('snap-2');

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    'reload 后 nextId 重建考虑 adapter 里已有的所有 id(取 max)',
    async () => {
      const adapter = new InMemorySnapshotStorage();
      const buf = await makeFixtureZip();
      const doc = await WordDocument.fromZip(buf, {
        docId: 'snap-reload-maxid',
        mode: 'worker',
        snapshotStorage: adapter,
      });

      // 拍 5 个,然后删中间两个(主线程 nextId 此时是 5,adapter 里有 snap-0/1/3/4)
      await doc.snapshots.take();
      await doc.snapshots.take();
      await doc.snapshots.take();
      await doc.snapshots.take();
      await doc.snapshots.take();
      await doc.snapshots.delete('snap-2');
      await doc.snapshots.delete('snap-1');

      // crash → reload → 下一个 take 应是 snap-5(max(0,3,4) + 1)
      const transport = doc._getTransport() as WorkerTransport;
      const inflight = doc.toZip();
      transport._triggerCrashForTest();
      await expect(inflight).rejects.toBeInstanceOf(WordError);

      const next = await doc.snapshots.take('after-reload');
      expect(next).toBe('snap-5');

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test('Inline 模式:reinitializeAfterReload 手调不炸,且 transport 不暴露 onReload', async () => {
    const adapter = new InMemorySnapshotStorage();
    const doc = await WordDocument.fromZip(await makeFixtureZip(), {
      docId: 'snap-reload-inline',
      mode: 'inline',
      snapshotStorage: adapter,
    });
    expect(doc.mode).toBe('inline');

    // InlineTransport 不应暴露 onReload(可选 API)
    const transport = doc._getTransport();
    expect((transport as unknown as { onReload?: unknown }).onReload).toBeUndefined();

    // 手工调 reinitializeAfterReload 不应抛错(adapter 为空,nextId 回到 0)
    const id0 = await doc.snapshots.take();
    expect(id0).toBe('snap-0');
    // 删掉使 adapter 再次空
    await doc.snapshots.delete(id0);
    // 手调重建:max=-1 → nextId=0
    await doc.snapshots.reinitializeAfterReload();
    const idAfter = await doc.snapshots.take();
    expect(idAfter).toBe('snap-0');

    await doc.dispose();
  });
});
