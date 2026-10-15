/**
 * Runtime mode 决策测试 —— #18。
 *
 * 覆盖 docs/03 §4.2 行为表 + §4.3 auto 决策顺序:
 *
 *   | mode          | 可用          | 不可用                   |
 *   |---------------|--------------|--------------------------|
 *   | 'worker' 显式 | worker       | 抛 WORKER_CREATION_FAILED|
 *   | 'inline' 显式 | inline       | inline                   |
 *   | 'auto' 默认   | worker       | 静默 fallback 到 inline  |
 *
 * 本测试跑在 Node 下,主要验证 Node 路径 + workerFactory/workerBundleUrl 的组合;
 * 浏览器 Worker 的真实 ping 留 #52 浏览器示例完善。
 */
import { afterEach, describe, expect, test, vi } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
import { W14_NS, W_NS } from '../namespaces.js';
import type { WorkerPortLike } from '../worker-transport.js';

// ─────────── 测试夹具 ───────────

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
    <w:p><w:r><w:t>mode</w:t></w:r></w:p>
  </w:body>
</w:document>`;

async function makeFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', DOC_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

/** 每个 case 独立清状态,避免互相污染全局 defaults。*/
afterEach(() => {
  WordDocument._resetDefaults();
});

// ─────────── detectMode —— 反映 auto 的真实选择 ───────────

describe('detectMode() 反映 auto 的真实选择', () => {
  test('Node + 未配 bundle / factory → detectMode 返 inline', () => {
    expect(WordDocument.detectMode()).toBe('inline');
  });

  test('Node + 设 workerBundleUrl → detectMode 返 worker', () => {
    WordDocument.setWorkerBundleUrl('file:///fake/kernel-worker.mjs');
    expect(WordDocument.detectMode()).toBe('worker');
  });

  test('Node + 设 workerFactory → detectMode 返 worker', () => {
    WordDocument.setWorkerFactory(() => ({
      postMessage: () => {},
      onMessage: () => {},
      onError: () => {},
      terminate: async () => {},
    }));
    expect(WordDocument.detectMode()).toBe('worker');
  });

  test('清空 workerBundleUrl 后 detectMode 回到 inline', () => {
    WordDocument.setWorkerBundleUrl('file:///fake/kernel-worker.mjs');
    expect(WordDocument.detectMode()).toBe('worker');
    WordDocument.setWorkerBundleUrl(undefined);
    expect(WordDocument.detectMode()).toBe('inline');
  });
});

// ─────────── detectModeDecision —— 决策同时带原因 ───────────

describe('detectModeDecision() 暴露决策来源', () => {
  test('Node 主线程 + 无配置 → auto/node-no-worker-config', () => {
    expect(WordDocument.detectModeDecision()).toEqual({
      mode: 'inline',
      source: 'auto/node-no-worker-config',
    });
  });

  test('Node 主线程 + 配 bundleUrl → auto/node-with-worker-config', () => {
    WordDocument.setWorkerBundleUrl('file:///fake/kernel-worker.mjs');
    expect(WordDocument.detectModeDecision()).toEqual({
      mode: 'worker',
      source: 'auto/node-with-worker-config',
    });
  });
});

// ─────────── 装载日志 / onModeResolved 钩子 ───────────

describe('装载时模式日志 + onModeResolved 钩子', () => {
  test('未传 onModeResolved → 默认 console.info 打 mode/source/docId', async () => {
    const buf = await makeFixtureZip();
    const spy = vi.spyOn(console, 'info').mockImplementation(() => {});
    try {
      const doc = await WordDocument.fromZip(buf, { docId: 'm-log1' });
      await doc.dispose();
      expect(spy).toHaveBeenCalledWith(
        '[coco-office:word] mode=inline source=auto/node-no-worker-config docId=m-log1',
      );
    } finally {
      spy.mockRestore();
    }
  });

  test('传 onModeResolved → 默认 console.info 不再打,回调拿到决策', async () => {
    const buf = await makeFixtureZip();
    const spy = vi.spyOn(console, 'info').mockImplementation(() => {});
    const calls: Array<[string, string, string]> = [];
    try {
      const doc = await WordDocument.fromZip(buf, {
        docId: 'm-log2',
        onModeResolved: (decision, docId) => {
          calls.push([decision.mode, decision.source, docId]);
        },
      });
      await doc.dispose();
      expect(calls).toEqual([['inline', 'auto/node-no-worker-config', 'm-log2']]);
      expect(spy).not.toHaveBeenCalled();
    } finally {
      spy.mockRestore();
    }
  });

  test('onModeResolved 抛错不影响装载', async () => {
    const buf = await makeFixtureZip();
    const spy = vi.spyOn(console, 'info').mockImplementation(() => {});
    try {
      const doc = await WordDocument.fromZip(buf, {
        docId: 'm-log3',
        onModeResolved: () => {
          throw new Error('host bug');
        },
      });
      expect(doc.mode).toBe('inline');
      expect(doc.modeSource).toBe('auto/node-no-worker-config');
      await doc.dispose();
    } finally {
      spy.mockRestore();
    }
  });

  test('doc.modeSource 与决策一致(显式 inline)', async () => {
    const buf = await makeFixtureZip();
    const spy = vi.spyOn(console, 'info').mockImplementation(() => {});
    try {
      const doc = await WordDocument.fromZip(buf, { docId: 'm-log4', mode: 'inline' });
      expect(doc.mode).toBe('inline');
      expect(doc.modeSource).toBe('explicit-inline');
      await doc.dispose();
    } finally {
      spy.mockRestore();
    }
  });
});

// ─────────── 'inline' 显式 —— 始终 inline ───────────

describe("mode: 'inline' 显式", () => {
  test('Node + 无 bundle → inline', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'm-i1', mode: 'inline' });
    expect(doc.mode).toBe('inline');
    await doc.dispose();
  });

  test('Node + 配了 bundle 也走 inline(显式优先)', async () => {
    const buf = await makeFixtureZip();
    WordDocument.setWorkerBundleUrl('file:///fake/kernel-worker.mjs');
    // 即便 bundle 不存在,inline 模式不会实例化 worker,因此不会抛 WORKER_CREATION_FAILED
    const doc = await WordDocument.fromZip(buf, { docId: 'm-i2', mode: 'inline' });
    expect(doc.mode).toBe('inline');
    await doc.dispose();
  });
});

// ─────────── 'worker' 显式 —— 不可用时抛 WORKER_CREATION_FAILED ───────────

describe("mode: 'worker' 显式", () => {
  test('Node + 无 bundle / factory → 抛 WORKER_CREATION_FAILED', async () => {
    const buf = await makeFixtureZip();
    try {
      await WordDocument.fromZip(buf, { docId: 'm-w1', mode: 'worker' });
      throw new Error('should throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('WORKER_CREATION_FAILED');
      expect((e as WordError).docId).toBe('m-w1');
    }
  });

  test('Node + factory 可用 → 装载成功,mode=worker', async () => {
    // 用 InlineKernel 驱动一个假的 WorkerPortLike,走 worker RPC 路径
    const { port, dispose } = makeInMemoryWorkerPort();
    WordDocument.setWorkerFactory(() => port);
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'm-w2', mode: 'worker' });
    expect(doc.mode).toBe('worker');
    await doc.dispose();
    await dispose();
  });
});

// ─────────── 'auto' / 默认 —— 不可用时静默 fallback ───────────

describe("mode: 'auto'(默认)", () => {
  test('Node + 未配 bundle → 静默走 inline,无 warning / 不抛', async () => {
    const buf = await makeFixtureZip();
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      // 默认就是 'auto'
      const doc = await WordDocument.fromZip(buf, { docId: 'm-a1' });
      expect(doc.mode).toBe('inline');
      await doc.dispose();
      // 显式 auto
      const doc2 = await WordDocument.fromZip(buf, { docId: 'm-a1b', mode: 'auto' });
      expect(doc2.mode).toBe('inline');
      await doc2.dispose();
      expect(warnSpy).not.toHaveBeenCalled();
      expect(errorSpy).not.toHaveBeenCalled();
    } finally {
      warnSpy.mockRestore();
      errorSpy.mockRestore();
    }
  });

  test('Node + workerFactory 配了 → auto 走 worker', async () => {
    const { port, dispose } = makeInMemoryWorkerPort();
    WordDocument.setWorkerFactory(() => port);
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'm-a2', mode: 'auto' });
    expect(doc.mode).toBe('worker');
    await doc.dispose();
    await dispose();
  });

  test('setDefaultMode(auto) + 无 bundle → 走 inline(默认行为)', async () => {
    WordDocument.setDefaultMode('auto');
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'm-a3' });
    expect(doc.mode).toBe('inline');
    await doc.dispose();
  });
});

// ─────────── 决策顺序(docs/03 §4.3) ───────────

describe('auto 决策顺序', () => {
  test('1. Node 主线程 + 无配置 → inline', () => {
    // 纯 Node 场景,detectMode 直接验证第 1/2/4 条逻辑落点
    expect(WordDocument.detectMode()).toBe('inline');
  });

  test('2. Node 主线程 + 配了 bundle → worker(对应 §4.3 第 3 条在 Node 的实现)', () => {
    WordDocument.setWorkerBundleUrl('file:///fake/kernel-worker.mjs');
    expect(WordDocument.detectMode()).toBe('worker');
  });

  test('3. detectMode 不应抛错,无论配置如何', () => {
    expect(() => WordDocument.detectMode()).not.toThrow();
    WordDocument.setWorkerBundleUrl('file:///fake/kernel-worker.mjs');
    expect(() => WordDocument.detectMode()).not.toThrow();
    WordDocument.setWorkerFactory(() => ({
      postMessage: () => {},
      onMessage: () => {},
      onError: () => {},
      terminate: async () => {},
    }));
    expect(() => WordDocument.detectMode()).not.toThrow();
  });
});

// ─────────── 辅助:内存 WorkerPortLike(免 bundle 构建) ───────────

/**
 * 构造一个由同进程 Kernel 驱动的 `WorkerPortLike`,跳过真实线程边界。
 * 用于验证 worker 模式在 Node 下的装载路径,而无需构建 dist bundle。
 *
 * 实现:本地 Kernel + kernelHandlers dispatch,模拟 kernel-worker-entry 的消息回路。
 */
function makeInMemoryWorkerPort(): { port: WorkerPortLike; dispose: () => Promise<void> } {
  type Listener = (data: unknown) => void;
  const mainListeners: Listener[] = [];

  // 延迟加载 kernel/handlers,避免测试模块级的同步副作用
  let dispatch: ((raw: unknown) => Promise<void>) | null = null;

  async function ensureDispatch(): Promise<void> {
    if (dispatch) return;
    const [{ Kernel }, { kernelHandlers }, { WordError }] = await Promise.all([
      import('../kernel.js'),
      import('../kernel-handlers.js'),
      import('../word-error.js'),
    ]);
    const kernel = new Kernel();

    function serialize(e: unknown): {
      code?: string;
      message: string;
      docId?: string;
      details?: Record<string, unknown>;
      name?: string;
    } {
      if (e instanceof WordError) {
        return {
          code: e.code,
          message: e.message,
          docId: e.docId,
          details: e.details,
          name: e.name,
        };
      }
      const err = e as Error | undefined;
      return { message: err?.message ?? String(e), name: err?.name };
    }

    dispatch = async (raw: unknown): Promise<void> => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const req = raw as any;
      if (!req || req.kind !== 'request') return;
      const handler = (kernelHandlers as Record<string, unknown>)[req.method];
      if (typeof handler !== 'function') {
        broadcast({ kind: 'error', id: req.id, error: { message: `unknown: ${req.method}` } });
        return;
      }
      try {
        const result = await (
          handler as (
            ctx: { kernel: unknown; onProgress: (p: string, pct?: number) => void },
            args: unknown[],
          ) => Promise<unknown>
        )(
          {
            kernel,
            onProgress: (phase, pct) => broadcast({ kind: 'progress', id: req.id, phase, pct }),
          },
          req.args,
        );
        broadcast({ kind: 'response', id: req.id, result });
      } catch (e) {
        broadcast({ kind: 'error', id: req.id, error: serialize(e) });
      }
    };
  }

  function broadcast(msg: unknown): void {
    queueMicrotask(() => {
      for (const l of mainListeners) l(msg);
    });
  }

  const port: WorkerPortLike = {
    postMessage(data) {
      void ensureDispatch().then(() => dispatch?.(data));
    },
    onMessage(listener) {
      mainListeners.push(listener);
    },
    onError() {
      // 本测试不模拟错误通道
    },
    async terminate() {
      mainListeners.length = 0;
    },
  };

  return {
    port,
    dispose: async () => {
      mainListeners.length = 0;
    },
  };
}
