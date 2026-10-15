/**
 * script-execute + WordDocument.executeCode 单测(#76,04.2 §六)。
 *
 * 覆盖:
 * - 基本执行:return 值穿透
 * - 顶层 await
 * - 访问 doc 的 namespace 形态 API:doc.paragraphs.list / doc.plan()(#159)
 * - args 传参
 * - 脚本抛错 → 透出 Error
 * - 语法错误 → 抛 'script syntax error: ...'
 * - timeoutMs 超时 → 抛 'script timeout after Xms'
 * - timeoutMs === 0 → 不超时
 * - 通过 InlineTransport / WordDocument.executeCode end-to-end
 */
import { describe, expect, test } from 'vitest';
import { W14_NS, W_NS } from '../namespaces.js';
import { WordDocument } from '../word-document.js';

const FLAT_OPC = `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p><w:r><w:t>hello</w:t></w:r></w:p>
          <w:p><w:r><w:t>world</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

async function makeDoc() {
  return WordDocument.fromFlatOpc(FLAT_OPC, { docId: 'd', mode: 'inline' });
}

describe('WordDocument.executeCode (inline)', () => {
  test('基本执行:return 字面量', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode('return 42');
    expect(result).toBe(42);
  });

  test('顶层 await', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode('await new Promise(r => setTimeout(r, 1)); return "done"');
    expect(result).toBe('done');
  });

  test('args 传参', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode('return args.x * 2', { x: 21 });
    expect(result).toBe(42);
  });

  test('#159 sandbox doc: paragraphs.list (namespace 形态)', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode(
      `const list = await doc.paragraphs.list(); return Promise.all(list.paragraphs.map(p => p.getText()))`,
    );
    expect(result).toEqual(['hello', 'world']);
  });

  test('#159 sandbox doc: docId 字段', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode(`return doc.docId`);
    expect(result).toBe('d');
  });

  test('#159 sandbox doc: plan().replaceParagraph().apply()', async () => {
    const doc = await makeDoc();
    const before = await doc.executeCode(
      `const list = await doc.paragraphs.list(); return Promise.all(list.paragraphs.map(async p => ({ id: p.paraId, text: await p.getText() })))`,
    );
    const firstId = (before as Array<{ id: string }>)[0]!.id;
    const result = await doc.executeCode(
      `await doc.plan().replaceParagraph(args.id, 'replaced').apply();
       const list = await doc.paragraphs.list();
       return Promise.all(list.paragraphs.map(p => p.getText()));`,
      { id: firstId },
    );
    expect(result).toEqual(['replaced', 'world']);
  });

  test('#159 sandbox doc: 不暴露 legacy kernel-direct API', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode(
      `return {
         hasListParagraphs: typeof doc.listParagraphs,
         hasApplyPlan: typeof doc.applyPlan,
         hasGetParagraph: typeof doc.getParagraph,
         hasExecuteCode: typeof doc.executeCode,
       }`,
    );
    expect(result).toEqual({
      hasListParagraphs: 'undefined',
      hasApplyPlan: 'undefined',
      hasGetParagraph: 'undefined',
      hasExecuteCode: 'undefined',
    });
  });

  test('脚本抛错 → 错误透出', async () => {
    const doc = await makeDoc();
    await expect(doc.executeCode('throw new Error("custom-fail")')).rejects.toThrow('custom-fail');
  });

  test('语法错误 → script syntax error', async () => {
    const doc = await makeDoc();
    await expect(doc.executeCode('this is not valid ((( js')).rejects.toThrow(
      /script syntax error/,
    );
  });

  test('timeoutMs 超时:Promise.race 抛 timeout', async () => {
    const doc = await makeDoc();
    await expect(
      doc.executeCode('await new Promise(r => setTimeout(r, 200)); return 1', undefined, 50),
    ).rejects.toThrow(/script timeout after 50ms/);
  });

  test('timeoutMs === 0:不超时', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode(
      'await new Promise(r => setTimeout(r, 30)); return "ok"',
      undefined,
      0,
    );
    expect(result).toBe('ok');
  });

  test('args 默认 undefined → 脚本里 args === undefined', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode('return args === undefined');
    expect(result).toBe(true);
  });

  test('返回值经 structured-clone(POJO 可,函数不可)', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode('return { a: 1, b: [2, 3] }');
    expect(result).toEqual({ a: 1, b: [2, 3] });
  });

  // ─────────── #151 abort 语义 ───────────

  test('abort: signal 已 abort → 立即抛 script aborted', async () => {
    const doc = await makeDoc();
    const ac = new AbortController();
    ac.abort();
    await expect(doc.executeCode('return 1', undefined, 30000, ac.signal)).rejects.toThrow(
      /aborted/,
    );
  });

  test('abort: 脚本运行中触发 abort → race 立即拒绝', async () => {
    const doc = await makeDoc();
    const ac = new AbortController();
    setTimeout(() => ac.abort(), 30);
    const start = Date.now();
    await expect(
      doc.executeCode(
        'await new Promise(r => setTimeout(r, 1000)); return "should not get here"',
        undefined,
        30000,
        ac.signal,
      ),
    ).rejects.toThrow(/aborted/);
    const elapsed = Date.now() - start;
    expect(elapsed).toBeLessThan(200); // 不等脚本自然完成的 1000ms
  });

  test('abort: 不传 signal → 行为同前(不影响)', async () => {
    const doc = await makeDoc();
    const result = await doc.executeCode('return 7', undefined, 30000);
    expect(result).toBe(7);
  });
});
