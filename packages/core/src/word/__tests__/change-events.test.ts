/**
 * `WordDocument.on('change', ...)` —— inline + worker 端到端。
 *
 * 覆盖:
 *   - inline 路径:plan apply 触发 source='plan' 事件;ops/result 完整透传
 *   - inline 路径:executeCode 内 `doc.plan().apply()` 触发 source='script' 事件
 *   - 空 plan 不触发(契约:空 plan 不算变更)
 *   - apply 抛错时不触发
 *   - fullDocument: 'flatOpc' 拉到全文 XML;'zip' 拉到字节;多订阅者共享一次拉取
 *   - 取消订阅后不再收事件
 *   - 多次连续 apply,事件按变更顺序串行交付(handler async 也保序)
 *   - worker 路径:RpcEvent 跨 worker 边界 fan-out;source 字段保留
 *
 * worker e2e 走 [plan-apply-worker.test.ts](./plan-apply-worker.test.ts) 同款 bundle 路径。
 */
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import { execSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import type { ChangeEvent } from '../change-events.js';
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

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

async function makeFixture(bodyInner: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  return zip.generateAsync({ type: 'uint8array' });
}

const ONE_PARA = `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`;

/**
 * 等到 `predicate()` 真返回 true 再继续。change 派发是微任务 + 可能 await
 * toZip(jszip 压缩可达几十 ms),固定 sleep 在并行高负载 CI 下会 flake。
 * 默认上限 1000 ms,够慢机器跑完一次 toZip 还有富余。
 */
async function waitFor(predicate: () => boolean, timeoutMs = 1000): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (predicate()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  if (!predicate()) throw new Error(`waitFor: predicate stayed false for ${timeoutMs}ms`);
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('WordDocument.on(change) —— inline 路径', () => {
  test('plan().apply() 后 listener 被调用,source=plan,ops/result 完整', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-inl-1', mode: 'inline' });

    const events: ChangeEvent[] = [];
    doc.on('change', (ev) => {
      events.push(ev);
    });

    await doc.plan().replaceParagraph('AAAAAAA1', 'changed').apply();

    // listener 同步 + dispatch 走微任务;让事件循环转一次
    await Promise.resolve();
    await Promise.resolve();

    expect(events).toHaveLength(1);
    const ev = events[0]!;
    expect(ev.source).toBe('plan');
    expect(ev.ops).toEqual([
      { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'changed', onAnchorConflict: 'error' },
    ]);
    expect(ev.result.insertedParagraphs).toEqual([]);
    expect(ev.flatOpc).toBeUndefined();
    expect(ev.zip).toBeUndefined();
    await doc.dispose();
  });

  test('executeCode 内调 doc.plan().apply() → source=script', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-inl-2', mode: 'inline' });

    const events: ChangeEvent[] = [];
    doc.on('change', (ev) => {
      events.push(ev);
    });

    // 沙箱 doc 与主线程 WordDocument 同形态(#159);plan().apply() 走 sandbox
    // transport,该 transport 显式 tag source='script'(覆盖 handler 默认 'plan')
    await doc.executeCode(
      `
      await doc.plan().replaceParagraph('AAAAAAA1', 'from-script').apply();
      return 'done';
    `,
    );

    await Promise.resolve();
    await Promise.resolve();

    expect(events).toHaveLength(1);
    expect(events[0]!.source).toBe('script');
    expect(events[0]!.ops[0]).toMatchObject({ type: 'replaceParagraph', text: 'from-script' });

    const info = await doc.paragraphs.get('AAAAAAA1');
    expect(await info?.getText()).toBe('from-script');
    await doc.dispose();
  });

  test('空 plan 不触发事件', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-inl-empty', mode: 'inline' });

    const events: ChangeEvent[] = [];
    doc.on('change', (ev) => {
      events.push(ev);
    });

    // 空 plan 走 console.warn 路径,不抛
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    await doc.plan().apply();
    warnSpy.mockRestore();

    await Promise.resolve();
    expect(events).toHaveLength(0);
    await doc.dispose();
  });

  test('apply 抛错(回滚)时不触发事件', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-inl-throw', mode: 'inline' });

    const events: ChangeEvent[] = [];
    doc.on('change', (ev) => {
      events.push(ev);
    });

    await expect(doc.plan().replaceParagraph('NOTAPARA', 'x').apply()).rejects.toThrow();

    await Promise.resolve();
    expect(events).toHaveLength(0);
    await doc.dispose();
  });

  test('fullDocument:flatOpc 携带全文 XML,zip 携带字节', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-inl-full', mode: 'inline' });

    const flatEvents: ChangeEvent[] = [];
    const zipEvents: ChangeEvent[] = [];
    doc.on(
      'change',
      (ev) => {
        flatEvents.push(ev);
      },
      { fullDocument: 'flatOpc' },
    );
    doc.on(
      'change',
      (ev) => {
        zipEvents.push(ev);
      },
      { fullDocument: 'zip' },
    );

    await doc.plan().replaceParagraph('AAAAAAA1', 'with-doc').apply();

    // toZip 走 jszip 压缩,比 toFlatOpc 慢一截;轮询等到都到齐
    await waitFor(() => flatEvents.length >= 1 && zipEvents.length >= 1);

    expect(flatEvents).toHaveLength(1);
    expect(flatEvents[0]!.flatOpc).toContain('with-doc');
    expect(flatEvents[0]!.zip).toBeUndefined();

    expect(zipEvents).toHaveLength(1);
    expect(zipEvents[0]!.zip).toBeInstanceOf(Uint8Array);
    expect(zipEvents[0]!.flatOpc).toBeUndefined();

    // 解开 zip 验证内容
    const inner = await JSZip.loadAsync(zipEvents[0]!.zip!);
    const xml = await inner.file('word/document.xml')!.async('string');
    expect(xml).toContain('with-doc');
    await doc.dispose();
  });

  test('多 flatOpc 订阅者共享一次 toFlatOpc 拉取', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-inl-share', mode: 'inline' });

    const flatSpy = vi.spyOn(doc, 'toFlatOpc');
    const a: ChangeEvent[] = [];
    const b: ChangeEvent[] = [];
    doc.on(
      'change',
      (ev) => {
        a.push(ev);
      },
      { fullDocument: 'flatOpc' },
    );
    doc.on(
      'change',
      (ev) => {
        b.push(ev);
      },
      { fullDocument: 'flatOpc' },
    );

    await doc.plan().replaceParagraph('AAAAAAA1', 'shared').apply();
    await waitFor(() => a.length >= 1 && b.length >= 1);

    expect(a).toHaveLength(1);
    expect(b).toHaveLength(1);
    expect(a[0]!.flatOpc).toContain('shared');
    expect(b[0]!.flatOpc).toContain('shared');
    expect(flatSpy).toHaveBeenCalledTimes(1);
    flatSpy.mockRestore();
    await doc.dispose();
  });

  test('取消订阅后不再收事件', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-inl-off', mode: 'inline' });

    const events: ChangeEvent[] = [];
    const off = doc.on('change', (ev) => {
      events.push(ev);
    });
    await doc.plan().replaceParagraph('AAAAAAA1', 'first').apply();
    await Promise.resolve();
    await Promise.resolve();
    expect(events).toHaveLength(1);

    off();
    await doc.plan().replaceParagraph('AAAAAAA1', 'second').apply();
    await Promise.resolve();
    await Promise.resolve();
    expect(events).toHaveLength(1);

    // 幂等
    off();
    await doc.dispose();
  });

  test('async listener 按变更顺序串行交付', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-inl-order', mode: 'inline' });

    const seen: string[] = [];
    doc.on('change', async (ev) => {
      // 第一个事件故意慢;若没串行,第二次会先 push 进去
      const text = (ev.ops[0] as { text: string }).text;
      if (text === 'first') {
        await new Promise((r) => setTimeout(r, 30));
      }
      seen.push(text);
    });

    await doc.plan().replaceParagraph('AAAAAAA1', 'first').apply();
    await doc.plan().replaceParagraph('AAAAAAA1', 'second').apply();

    // 等到两次 handler 都跑完(第一次 handler 内部 sleep 30ms,留富余)
    await waitFor(() => seen.length >= 2);
    expect(seen).toEqual(['first', 'second']);
    await doc.dispose();
  });
});

describe('WordDocument.on(change) —— worker 路径', () => {
  beforeAll(() => {
    execSync('corepack pnpm --filter @coco-vexel/coco-office build', { cwd: rootDir, stdio: 'pipe' });
    WordDocument.setWorkerBundleUrl(bundleUrl);
  }, 120_000);

  test('worker 模式下 plan apply / executeCode 都能跨 worker 边界推事件', async () => {
    const buf = await makeFixture(ONE_PARA);
    const doc = await WordDocument.fromZip(buf, { docId: 'ce-w-1', mode: 'worker' });

    const events: ChangeEvent[] = [];
    doc.on(
      'change',
      (ev) => {
        events.push(ev);
      },
      { fullDocument: 'flatOpc' },
    );

    await doc.plan().replaceParagraph('AAAAAAA1', 'plan-w').apply();
    await waitFor(() => events.length >= 1);

    await doc.executeCode(`
        await doc.plan().replaceParagraph('AAAAAAA1', 'script-w').apply();
      `);
    await waitFor(() => events.length >= 2);

    expect(events).toHaveLength(2);
    expect(events[0]!.source).toBe('plan');
    expect(events[0]!.flatOpc).toContain('plan-w');
    expect(events[1]!.source).toBe('script');
    expect(events[1]!.flatOpc).toContain('script-w');

    await doc.dispose();
  }, 30_000);
});
