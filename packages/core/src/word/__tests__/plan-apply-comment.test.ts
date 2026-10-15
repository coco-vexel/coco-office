/**
 * WordPlan.apply —— addComment / deleteComment 端到端测试(#35)。
 *
 * 走整条链:doc.plan() → builder → apply() → kernel.applyPlan → dispatchOp →
 * executeAddComment / executeDeleteComment → runTransaction。
 *
 * 覆盖:
 *   - addComment ParagraphRange → listComments 含新 id,author / date / text / rangeKind / paraId 正确
 *   - 连续 3 次 addComment,id 递增(0, 1, 2)
 *   - addComment + deleteComment 同 plan:apply 结果 createdComments 有 1,最终 listComments 不含那个 id
 *   - 跨 parts 原子性:addComment + 第二 op 抛错 → 回滚后 comments.xml 不存在(若首次)
 *   - deleteComment 不存在 id → COMMENT_NOT_FOUND
 *   - toZip → 回加载 → listComments 保留批注
 *
 * inline 模式;worker 端到端见 [plan-apply-worker.test.ts](./plan-apply-worker.test.ts)。
 */
import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
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

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

async function makeFixture(
  bodyInner: string,
  opts: { docRels?: boolean } = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  // word/_rels/document.xml.rels must exist for addComment to update rels
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  void opts;
  return zip.generateAsync({ type: 'uint8array' });
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('WordPlan apply —— addComment 端到端(inline)', () => {
  test('validate 收集全部目标错误且不修改文档', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello world</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'validate-ac-1', mode: 'inline' });
    const validation = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1', text: 'hello' }, text: 'valid' })
      .addComment({ range: { paraId: 'NOTFOUND', text: 'hello' }, text: 'bad para' })
      .addComment({ range: { paraId: 'AAAAAAA1', text: 'missing' }, text: 'bad text' })
      .validate();

    expect(validation.valid).toBe(false);
    expect(validation.items.map((item) => item.status)).toEqual(['valid', 'invalid', 'invalid']);
    expect(validation.items[1]?.issue?.code).toBe('PARA_NOT_FOUND');
    expect(validation.items[2]?.issue?.code).toBe('TEXT_NOT_FOUND_IN_PARA');
    expect(await doc.comments.list()).toEqual([]);
    await doc.dispose();
  });

  test('addComment ParagraphRange → listComments 字段正确', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ac-1', mode: 'inline' });
    const result = await doc
      .plan({ author: 'alice', date: new Date('2026-04-22T10:00:00Z') })
      .addComment({
        range: { paraId: 'AAAAAAA1' },
        text: 'this is a paragraph comment',
      })
      .apply();

    expect(result.createdComments).toHaveLength(1);
    const { commentId } = result.createdComments[0]!;
    expect(commentId).toBe(0);

    const comments = await doc.comments.list();
    expect(comments).toHaveLength(1);
    const c = comments[0]!;
    expect(c.id).toBe(0);
    expect(c.author).toBe('alice');
    expect(c.text).toBe('this is a paragraph comment');
    expect(c.rangeKind).toBe('paragraph');
    expect(c.paraId).toBe('AAAAAAA1');

    await doc.dispose();
  });

  test('连续 3 次 addComment,id 递增(0, 1, 2)', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>p1</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>p2</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA3"><w:r><w:t>p3</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ac-2', mode: 'inline' });
    const result = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'c1' })
      .addComment({ range: { paraId: 'AAAAAAA2' }, text: 'c2' })
      .addComment({ range: { paraId: 'AAAAAAA3' }, text: 'c3' })
      .apply();

    expect(result.createdComments.map((e) => e.commentId)).toEqual([0, 1, 2]);
    const comments = await doc.comments.list();
    expect(comments).toHaveLength(3);

    await doc.dispose();
  });

  test('addComment + deleteComment 同 plan:createdComments 有 1,listComments 最终不含该 id', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ac-3', mode: 'inline' });

    // 先单独加一个批注以获得 id=0
    const r1 = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'first' })
      .apply();
    const firstId = r1.createdComments[0]!.commentId; // 0

    // 再 plan:add(新 id=1) + delete(id=0)
    const r2 = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'second' })
      .deleteComment(firstId)
      .apply();

    expect(r2.createdComments).toHaveLength(1);
    const secondId = r2.createdComments[0]!.commentId;

    const comments = await doc.comments.list();
    const ids = comments.map((c) => c.id);
    expect(ids).toContain(secondId);
    expect(ids).not.toContain(firstId);

    await doc.dispose();
  });

  test('跨 parts 原子性:addComment + 第二 op 抛错 → 回滚后 comments.xml 不存在(首次)', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ac-4', mode: 'inline' });

    // addComment(成功) + replaceParagraph('NOTFOUND', 'x')(失败) → 整体回滚
    const plan = doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'rollback me' })
      .replaceParagraph('NOTFOUND', 'x');

    await expect(plan.apply()).rejects.toBeInstanceOf(WordError);

    // 回滚后 listComments 不含任何批注
    const comments = await doc.comments.list();
    expect(comments).toHaveLength(0);

    await doc.dispose();
  });

  test('deleteComment 不存在 id → COMMENT_NOT_FOUND', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-dc-1', mode: 'inline' });
    const plan = doc.plan().deleteComment(9999);
    await expect(plan.apply()).rejects.toMatchObject({ code: 'COMMENT_NOT_FOUND' });
    await doc.dispose();
  });

  test('toZip → 回加载 → listComments 保留批注', async () => {
    const buf = await makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ac-roundtrip', mode: 'inline' });
    await doc
      .plan({ author: 'roundtrip-user' })
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'survived' })
      .apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    // 重新加载
    const doc2 = await WordDocument.fromZip(zipBytes, { docId: 'e2e-ac-reload', mode: 'inline' });
    const comments = await doc2.comments.list();
    expect(comments).toHaveLength(1);
    expect(comments[0]!.text).toBe('survived');
    expect(comments[0]!.author).toBe('roundtrip-user');
    await doc2.dispose();
  });
});
