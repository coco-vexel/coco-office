/**
 * Worker mode 端到端测试 —— #17。
 *
 * 步骤:
 *   1. beforeAll 跑 `pnpm turbo run build`,产出 dist/kernel-worker.js
 *   2. WordDocument.setWorkerBundleUrl(dist 路径)
 *   3. 以 mode: 'worker' 装载,验证:
 *      - 装载成功
 *      - toZip / toFlatOpc 跨 RPC 边界返回正确结果
 *      - 错误(缺 CT)正确重抛为 WordError(保留 code / docId / details)
 *      - dispose 清理 worker,不留 pending
 */
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'vitest';
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

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p><w:r><w:t>worker-load</w:t></w:r></w:p>
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
  // turbo 缓存:首次慢,后续秒过
  execSync('corepack pnpm --filter @coco-vexel/coco-office build', { cwd: rootDir, stdio: 'pipe' });
  WordDocument.setWorkerBundleUrl(bundleUrl);
}, 120_000);

afterAll(() => {
  WordDocument._resetDefaults();
});

afterEach(() => {
  // 不 reset workerBundleUrl,只重置其他默认
});

describe('WordDocument worker mode — 正常路径', () => {
  test('mode=worker:fromZip 成功,mode 反映为 worker', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-1', mode: 'worker' });
    expect(doc.mode).toBe('worker');
    expect(doc.docId).toBe('w-1');
    await doc.dispose();
  });

  test('worker 模式 toZip 跨 RPC 返回 Uint8Array,可再次 fromZip(inline)装载', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-2', mode: 'worker' });
    const out = await doc.toZip();
    expect(out).toBeInstanceOf(Uint8Array);
    expect(out.length).toBeGreaterThan(0);

    // 用 inline 再装载,验证内容正常
    const roundtrip = await WordDocument.fromZip(out, { docId: 'w-2r', mode: 'inline' });
    expect(roundtrip._getIndices().paraIdx.size).toBe(1);
    await roundtrip.dispose();
    await doc.dispose();
  });

  test('worker 模式 toFlatOpc 跨 RPC 返回字符串', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-3', mode: 'worker' });
    const xml = await doc.toFlatOpc();
    expect(typeof xml).toBe('string');
    expect(xml).toMatch(/<\?mso-application progid="Word\.Document"\?>/);
    await doc.dispose();
  });

  test('worker 模式装载 Flat OPC 字符串', async () => {
    const flat = `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>flat-worker</w:t></w:r></w:p></w:body></w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    const doc = await WordDocument.fromFlatOpc(flat, { docId: 'w-4', mode: 'worker' });
    const out = await doc.toFlatOpc();
    expect(out).toMatch(/flat-worker/);
    await doc.dispose();
  });

  test('worker 模式 executeCode 跨 RPC:script.execute 分支正常 dispatch(#76)', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-script', mode: 'worker' });

    // 基础 return
    const r1 = await doc.executeCode('return args.x + 1', { x: 10 });
    expect(r1).toBe(11);

    // 顶层 await + doc(沙箱 facade,与主线程同形态,#159)
    const r2 = await doc.executeCode(
      `await Promise.resolve();
       const list = await doc.paragraphs.list();
       return list.paragraphs.length;`,
    );
    expect(r2).toBe(1); // DOC_XML 只有一段 'worker-load'

    // 错误透传
    await expect(doc.executeCode('throw new Error("boom")')).rejects.toThrow('boom');

    await doc.dispose();
  });

  test('worker 模式 executeCode 同步死循环 timeout → terminate + reload + 后续 RPC 可用(#144)', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-timeout', mode: 'worker' });

    // 同步死循环:Promise.race 软超时无能为力,只有主线程 terminate 能救
    const start = Date.now();
    let caught: unknown;
    try {
      await doc.executeCode('while (true) {} return 1', undefined, 200);
    } catch (e) {
      caught = e;
    }
    const elapsed = Date.now() - start;

    // 应被 transport 转 WORKER_RELOADED reject
    expect(caught).toBeInstanceOf(WordError);
    expect((caught as WordError).code).toBe('WORKER_RELOADED');
    // terminate 应在 timeoutMs 附近触发,不会等到测试 timeout
    expect(elapsed).toBeLessThan(5000);

    // reload 后文档应能继续工作:listParagraphs 返回原文档(从初始字节重建)
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await Promise.all(list.paragraphs.map((p) => p.getText()))).toEqual(['worker-load']);

    // 再跑一次脚本验证新 worker 健康
    const r = await doc.executeCode('return 42');
    expect(r).toBe(42);

    await doc.dispose();
  }, 30_000);

  test('worker 模式 executeCode timeoutMs=0 → 不超时,正常返回', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-no-timeout', mode: 'worker' });

    // 慢一点的脚本,timeoutMs=0 → 不应被强杀
    const r = await doc.executeCode(
      `await new Promise(r => setTimeout(r, 50)); return 'slow-ok'`,
      undefined,
      0,
    );
    expect(r).toBe('slow-ok');

    await doc.dispose();
  });

  test('worker 模式 executeCode 同步死循环 + signal abort → terminate + reload + 后续 RPC 可用(#151)', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-abort', mode: 'worker' });

    const ac = new AbortController();
    // 100ms 后 abort,模拟用户点"停止生成"按钮
    setTimeout(() => ac.abort(), 100);

    const start = Date.now();
    let caught: unknown;
    try {
      // 同步死循环 —— 只有 worker.terminate() 能救它
      await doc.executeCode('while (true) {} return 1', undefined, 30000, ac.signal);
    } catch (e) {
      caught = e;
    }
    const elapsed = Date.now() - start;

    // 应被 transport 转 WORKER_RELOADED reject
    expect(caught).toBeInstanceOf(WordError);
    expect((caught as WordError).code).toBe('WORKER_RELOADED');
    // abort 应在 ~100ms 触发,不会等到 timeout 30s
    expect(elapsed).toBeLessThan(2000);

    // reload 后文档仍可工作
    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await Promise.all(list.paragraphs.map((p) => p.getText()))).toEqual(['worker-load']);

    // 再跑脚本验证新 worker 健康
    const r = await doc.executeCode('return 42');
    expect(r).toBe(42);

    await doc.dispose();
  }, 30_000);

  test('worker 模式 executeCode 不传 signal → 行为同 #144(timeout reload)不变', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, {
      docId: 'w-no-signal',
      mode: 'worker',
    });

    // 不传 signal,单纯跑普通脚本
    const r = await doc.executeCode('return args.x * 2', { x: 21 });
    expect(r).toBe(42);

    await doc.dispose();
  });

  test('onProgress 回调跨 Worker 边界触发', async () => {
    const buf = await makeFixtureZip();
    const phases: string[] = [];
    const doc = await WordDocument.fromZip(buf, {
      docId: 'w-5',
      mode: 'worker',
      onProgress: (p) => phases.push(p),
    });
    // 至少收到 normalize(由 kernel 内 onProgress 回调产生)
    expect(phases).toContain('normalize');
    await doc.dispose();
  });
});

describe('WordDocument worker mode — 错误传递', () => {
  test('缺 [Content_Types].xml → 跨 Worker 重抛为 WordError,code/docId 保留', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', DOC_XML);
    const bad = await zip.generateAsync({ type: 'uint8array' });
    try {
      await WordDocument.fromZip(bad, { docId: 'w-bad', mode: 'worker' });
      throw new Error('should have thrown');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      const err = e as WordError;
      expect(err.code).toBe('PACKAGE_CORRUPT');
      expect(err.docId).toBe('w-bad');
    }
  });

  test('worker 模式重名 bookmark 装载容错:不崩,正常返回 doc', async () => {
    // 与 inline 路径一致:重名书签(含真实文档常见的重复 _GoBack)不再抛错,
    // worker 装载同样容错(保留第一个、忽略后续),fromZip 正常 resolve。
    const zip = new JSZip();
    zip.file('[Content_Types].xml', CT_XML);
    zip.file('_rels/.rels', MAIN_RELS);
    zip.file(
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}"><w:body>
        <w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p>
        <w:p><w:bookmarkStart w:id="1" w:name="_GoBack"/><w:bookmarkEnd w:id="1"/></w:p>
        <w:p><w:bookmarkStart w:id="2" w:name="Dup"/><w:bookmarkEnd w:id="2"/></w:p>
        <w:p><w:bookmarkStart w:id="3" w:name="Dup"/><w:bookmarkEnd w:id="3"/></w:p>
      </w:body></w:document>`,
    );
    const buf = await zip.generateAsync({ type: 'uint8array' });
    const doc = await WordDocument.fromZip(buf, { docId: 'w-dup', mode: 'worker' });
    expect(doc.docId).toBe('w-dup');
  });
});

describe('WordDocument worker mode — 内部 accessor 保护', () => {
  test('worker 模式 _getPackage 抛 UNSUPPORTED_OOXML_FEATURE', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-acc', mode: 'worker' });
    try {
      doc._getPackage();
      throw new Error('should throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
    }
    try {
      doc._getIndices();
      throw new Error('should throw');
    } catch (e) {
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
    }
    await doc.dispose();
  });
});

describe('WordDocument worker mode — dispose', () => {
  test('dispose 后 toZip 抛错', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-disp', mode: 'worker' });
    await doc.dispose();
    await expect(doc.toZip()).rejects.toBeInstanceOf(WordError);
  });

  test('dispose 多次不抛错', async () => {
    const buf = await makeFixtureZip();
    const doc = await WordDocument.fromZip(buf, { docId: 'w-disp2', mode: 'worker' });
    await doc.dispose();
    await expect(doc.dispose()).resolves.toBeUndefined();
  });
});

describe('WordDocument worker mode — 配置', () => {
  test('canUseWorker:workerBundleUrl 设置后 detectMode 返 worker', () => {
    expect(WordDocument.detectMode()).toBe('worker');
  });

  test('workerBundleUrl 清空后 mode=worker 抛 WORKER_CREATION_FAILED', async () => {
    const buf = await makeFixtureZip();
    WordDocument.setWorkerBundleUrl(undefined);
    try {
      await WordDocument.fromZip(buf, { docId: 'w-nofactory', mode: 'worker' });
      throw new Error('should throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('WORKER_CREATION_FAILED');
    } finally {
      WordDocument.setWorkerBundleUrl(bundleUrl); // 恢复供后续测试
    }
  });
});
