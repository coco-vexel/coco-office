/**
 * kernel.applyPlan 串行化互斥单测。
 *
 * 背景:多个 sub-agent 并发跑 `wordDoc.plan().apply()` 时,共享同一份 kernel /
 * pkg / xmlDom。applyPlan 内部会做 deepClone snapshot + 同步 mutate + 索引重建,
 * 没有跨调用互斥时(修复前):
 *   - 多份 deepClone 同时占内存(50w 字 doc × N 并发 → WPS WebView OOM)
 *   - dispatch 在共享 xmlDom 上交错执行,paraIdx / commentIdx 错位
 *   - 失败回滚的 snapshot 包含其它调用者的 mutation,会一起 wipe
 *
 * 修复:Kernel 内 `_applyPlanChain` promise chain 串行化所有 applyPlan,无论
 * 来源(plan / script)。
 *
 * 本测试通过 spy `_applyPlanInner`(私有方法,运行时仍可访问)统计同时 in-flight
 * 数量,验证修复后并发上限 = 1;并端到端验证并发 N 个 apply 全部命中预期文本。
 */
import { afterAll, describe, expect, test, vi } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import type { InlineTransport } from '../transport.js';
import { W14_NS, W_NS } from '../namespaces.js';

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

function paragraph(paraId: string, text: string): string {
  return `<w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

async function makeFixtureWithParagraphs(count: number): Promise<Uint8Array> {
  const ids: string[] = [];
  const body: string[] = [];
  for (let i = 0; i < count; i++) {
    // paraId 必须是 8 位 hex(w14:paraId 规约)
    const id = (0x10000000 + i).toString(16).toUpperCase().padStart(8, '0');
    ids.push(id);
    body.push(paragraph(id, `before-${i}`));
  }
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${body.join('')}</w:body>
</w:document>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml);
  return zip.generateAsync({ type: 'uint8array' });
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('kernel.applyPlan —— 并发互斥(修复点 C)', () => {
  test('5 个并发 apply 串行化:同时 in-flight 数永远 ≤ 1', async () => {
    const N = 5;
    const buf = await makeFixtureWithParagraphs(N);
    const doc = await WordDocument.fromZip(buf, {
      docId: 'mutex-inflight',
      mode: 'inline',
    });

    // 拿到底层 kernel(InlineTransport.kernel 是 readonly public)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const kernel = (doc as unknown as { transport: InlineTransport }).transport.kernel as any;
    const innerOriginal = kernel._applyPlanInner.bind(kernel);

    let active = 0;
    let maxConcurrent = 0;
    const spy = vi
      .spyOn(kernel, '_applyPlanInner')
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .mockImplementation(async (...args: any[]) => {
        active += 1;
        if (active > maxConcurrent) maxConcurrent = active;
        try {
          // 让出一次 microtask + 一次 macrotask,放大窗口让任何漏防都暴露
          await Promise.resolve();
          await new Promise((r) => setTimeout(r, 0));
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          return await innerOriginal(...(args as [any, any]));
        } finally {
          active -= 1;
        }
      });

    // 拿到所有 paraIds(按出现顺序)
    const list = await doc.paragraphs.list({ scope: 'body' });
    const ids = list.paragraphs.map((p) => p.paraId);
    expect(ids).toHaveLength(N);

    // 并发触发 N 个 apply,每个改一个独立段
    await Promise.all(
      ids.map((paraId, i) =>
        doc.plan().replaceTextInParagraph(paraId, `before-${i}`, `AFTER-${i}`).apply(),
      ),
    );

    // 关键断言:全程同时 in-flight 不超过 1(若没互斥应是 N)
    expect(maxConcurrent).toBe(1);
    expect(spy).toHaveBeenCalledTimes(N);

    // 端到端断言:每段文本都被改对了(无丢写、无错位)
    const after = await doc.paragraphs.list({ scope: 'body' });
    for (let i = 0; i < N; i++) {
      expect(await after.paragraphs[i]!.getText()).toBe(`AFTER-${i}`);
    }

    spy.mockRestore();
    await doc.dispose();
  });

  test('一个 apply 抛错不会卡住后续 apply(链上 release 不漏)', async () => {
    const buf = await makeFixtureWithParagraphs(2);
    const doc = await WordDocument.fromZip(buf, {
      docId: 'mutex-error-release',
      mode: 'inline',
    });

    const list = await doc.paragraphs.list({ scope: 'body' });
    const [id0, id1] = list.paragraphs.map((p) => p.paraId);

    // apply1:故意写不存在的 paraId → 触发 PARA_NOT_FOUND,事务回滚 + reject
    const apply1 = doc.plan().replaceParagraph('FFFFFFFF', 'never').apply();
    // apply2:并发触发,合法 op
    const apply2 = doc.plan().replaceParagraph(id1!, 'second-ok').apply();

    await expect(apply1).rejects.toBeDefined();
    await expect(apply2).resolves.toBeDefined();

    // 第一段未被动(apply1 整体回滚),第二段写成功
    const after = await doc.paragraphs.list({ scope: 'body' });
    expect(await after.paragraphs[0]!.getText()).toBe('before-0');
    expect(await after.paragraphs[1]!.getText()).toBe('second-ok');

    // 链未死锁:再发一个 apply 仍能跑
    await doc.plan().replaceParagraph(id0!, 'third-ok').apply();
    const final = await doc.paragraphs.list({ scope: 'body' });
    expect(await final.paragraphs[0]!.getText()).toBe('third-ok');

    await doc.dispose();
  });
});
