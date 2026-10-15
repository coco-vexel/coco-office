/**
 * Phase E 跨 parts 事务集成测试 —— #36。
 *
 * 覆盖 #34(createBookmark / deleteBookmark)+ #35(addComment / deleteComment)的
 * 组合场景 + 事务原子性验证(三处回滚:comments.xml / Content_Types.xml / rels)。
 *
 * 场景清单:
 *   1.  createBookmark + addComment 同段 → listBookmarks + listComments 都命中,rangeKind 正确;toZip 重加载一致
 *   2.  createBookmark + addComment 不同段 → 独立且正确
 *   3.  addComment → deleteComment → comments.xml 被移除(listComments=[]，task-comment-part-cleanup)
 *   4.  plan1 addComment(创建 comments.xml) → plan2 追加 → 都命中
 *   5.  同段已有 bookmark,再 addComment → 两锚点并存
 *   6.  首次 addComment 失败回滚 → comments.xml 不存在 + CT 无新 Override + rels 无新 Relationship + body 干净
 *   7.  第二次 addComment 失败回滚 → comments.xml 只保留 1 条 + CT/rels 不变 + body 回滚
 *   8.  deleteComment + 抛错 → 批注恢复
 *   9.  addComment → toZip → 回加载 → listComments + getComment 完整
 *   10. 删光所有批注 → toZip → 回加载 → comments.xml 不存在,listComments=[]（task-comment-part-cleanup）
 *
 * inline 模式;worker 端到端见 plan-apply-phase-e-worker.test.ts。
 */
import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { parseXmlDocument } from 'slimdom';
import { WordDocument } from '../word-document.js';
import { W14_NS, W_NS } from '../namespaces.js';

// ─────────── fixture 常量 ───────────

const CT_XML_BASE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS_ONE_REL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

const COMMENTS_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

/** 最小 fixture:无 comments.xml,有 doc rels */
async function makeFixture(bodyInner: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML_BASE);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  zip.file('word/_rels/document.xml.rels', DOC_RELS_ONE_REL);
  return zip.generateAsync({ type: 'uint8array' });
}

/** 从 toZip 结果中读取指定文件内容(字符串) */
async function readZipFile(zipBytes: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(zipBytes);
  const file = zip.file(path);
  if (!file) return null;
  return file.async('string');
}

/** 判断 XML 字符串中是否含 comments.xml 的 CT Override */
function hasCtOverrideForComments(ctXml: string): boolean {
  const dom = parseXmlDocument(ctXml) as unknown as Document;
  const overrides = dom.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
  for (let i = 0; i < overrides.length; i++) {
    if (overrides[i]!.getAttribute('PartName') === '/word/comments.xml') return true;
  }
  return false;
}

/** 判断 rels XML 字符串中是否含 comments 关系 */
function hasCommentsRel(relsXml: string): boolean {
  const dom = parseXmlDocument(relsXml) as unknown as Document;
  const rels = dom.getElementsByTagNameNS(RELS_NS, 'Relationship');
  for (let i = 0; i < rels.length; i++) {
    if (rels[i]!.getAttribute('Type') === COMMENTS_RELATIONSHIP_TYPE) return true;
  }
  return false;
}

afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── 组合场景 ───────────

describe('Phase E 集成 —— 组合场景', () => {
  test('1. createBookmark + addComment 同段 → 两者命中,rangeKind 正确;toZip 重加载一致', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">conclusion paragraph</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-1', mode: 'inline' });

    const result = await doc
      .plan({ author: 'alice', date: new Date('2026-04-22T10:00:00Z') })
      .createBookmark({ name: 'conclusionBm', range: { paraId: 'AAAAAAA1' } })
      .addComment({
        range: { paraId: 'AAAAAAA1', text: 'conclusion', scope: 'first' },
        text: '此处需要补充数据支撑',
      })
      .apply();

    // ApplyResult 同时含书签和批注
    expect(result.createdBookmarks).toEqual([
      {
        opIndex: 0,
        name: 'conclusionBm',
        range: { fromParaId: 'AAAAAAA1', toParaId: 'AAAAAAA1' },
      },
    ]);
    expect(result.createdComments).toHaveLength(1);
    const commentId = result.createdComments[0]!.commentId;

    // listBookmarks 命中
    const bookmarks = await doc.bookmarks.list();
    const bm = bookmarks.find((b) => b.name === 'conclusionBm');
    expect(bm).toBeDefined();
    expect(bm!.rangeKind).toBe('paragraph');
    expect(bm!.paraId?.toUpperCase()).toBe('AAAAAAA1');

    // listComments 命中
    const comments = await doc.comments.list();
    const cm = comments.find((c) => c.id === commentId);
    expect(cm).toBeDefined();
    expect(cm!.rangeKind).toBe('text');
    expect(cm!.anchorText).toBe('conclusion');
    expect(cm!.paraId?.toUpperCase()).toBe('AAAAAAA1');

    // toZip → 重加载仍一致
    const zipBytes = await doc.toZip();
    await doc.dispose();

    const doc2 = await WordDocument.fromZip(zipBytes, { docId: 'e2e-ph-1-reload', mode: 'inline' });
    const bms2 = await doc2.bookmarks.list();
    const cms2 = await doc2.comments.list();
    expect(bms2.find((b) => b.name === 'conclusionBm')).toBeDefined();
    expect(cms2.find((c) => c.id === commentId)).toBeDefined();
    await doc2.dispose();
  });

  test('2. createBookmark + addComment 不同段 → 独立且正确', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">para one</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t xml:space="preserve">para two</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-2', mode: 'inline' });

    const result = await doc
      .plan()
      .createBookmark({ name: 'bmForP1', range: { paraId: 'AAAAAAA1' } })
      .addComment({ range: { paraId: 'AAAAAAA2' }, text: 'comment on p2' })
      .apply();

    expect(result.createdBookmarks).toHaveLength(1);
    expect(result.createdBookmarks[0]!.name).toBe('bmForP1');
    expect(result.createdComments).toHaveLength(1);

    const bm = (await doc.bookmarks.list()).find((b) => b.name === 'bmForP1');
    expect(bm!.paraId?.toUpperCase()).toBe('AAAAAAA1');

    const cm = (await doc.comments.list())[0]!;
    expect(cm.paraId?.toUpperCase()).toBe('AAAAAAA2');
    expect(cm.rangeKind).toBe('paragraph');

    await doc.dispose();
  });

  test('3. addComment → deleteComment → comments.xml 被移除,listComments=[]', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">text</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-3', mode: 'inline' });

    const r = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'to delete' })
      .apply();
    const id = r.createdComments[0]!.commentId;

    // deleteComment
    await doc.plan().deleteComment(id).apply();

    // listComments 为空
    const comments = await doc.comments.list();
    expect(comments).toHaveLength(0);

    // 新行为：删光最后一条后 comments.xml part 被移除
    // toZip 后 word/comments.xml 不存在
    const zipBytes = await doc.toZip();
    const commentsXml = await readZipFile(zipBytes, 'word/comments.xml');
    expect(commentsXml).toBeNull();

    // CT 无 Override
    const ctXml = await readZipFile(zipBytes, '[Content_Types].xml');
    expect(ctXml).not.toContain('/word/comments.xml');

    await doc.dispose();
  });

  test('4. plan1 addComment(创建 comments.xml) → plan2 追加(commentId=1) → 两条都命中', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">para</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-4', mode: 'inline' });

    const r1 = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'first comment' })
      .apply();
    expect(r1.createdComments[0]!.commentId).toBe(0);

    const r2 = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'second comment' })
      .apply();
    expect(r2.createdComments[0]!.commentId).toBe(1);

    const comments = await doc.comments.list();
    expect(comments).toHaveLength(2);
    expect(comments.map((c) => c.id).sort()).toEqual([0, 1]);

    await doc.dispose();
  });

  test('5. 同段已有 bookmark,再 addComment → 两锚点并存,各指向同段', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="existingBm"/>
        <w:r><w:t xml:space="preserve">content</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-5', mode: 'inline' });

    // 验证书签已加载
    const bmsBefore = await doc.bookmarks.list();
    expect(bmsBefore.find((b) => b.name === 'existingBm')).toBeDefined();

    // 再加批注
    await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'new comment on same para' })
      .apply();

    // 书签还在
    const bms = await doc.bookmarks.list();
    expect(bms.find((b) => b.name === 'existingBm')).toBeDefined();

    // 批注也在
    const cms = await doc.comments.list();
    expect(cms).toHaveLength(1);
    expect(cms[0]!.paraId?.toUpperCase()).toBe('AAAAAAA1');

    await doc.dispose();
  });
});

// ─────────── 跨 parts 事务原子性 ───────────

describe('Phase E 集成 —— 跨 parts 事务原子性', () => {
  test('6. 首次 addComment 失败回滚 → comments.xml 消失 + CT 无 Override + rels 无 Relationship + body 干净', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">text</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-6', mode: 'inline' });

    // addComment(OK) + insertParagraphAfter('NOTEXIST', ...)(抛 PARA_NOT_FOUND) → 整体回滚
    const plan = doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'will be rolled back' })
      .insertParagraphAfter('NOTEXIST', { text: 'ghost' });

    await expect(plan.apply()).rejects.toMatchObject({ code: 'PARA_NOT_FOUND' });

    // listComments 为空
    expect(await doc.comments.list()).toHaveLength(0);

    // toZip 验证三处回滚
    const zipBytes = await doc.toZip();

    // comments.xml 不应存在(首次创建被回滚)
    const commentsXml = await readZipFile(zipBytes, 'word/comments.xml');
    expect(commentsXml).toBeNull();

    // Content_Types 无新 Override
    const ctXml = await readZipFile(zipBytes, '[Content_Types].xml');
    expect(ctXml).not.toBeNull();
    expect(hasCtOverrideForComments(ctXml!)).toBe(false);

    // rels 无新 Relationship
    const relsXml = await readZipFile(zipBytes, 'word/_rels/document.xml.rels');
    expect(relsXml).not.toBeNull();
    expect(hasCommentsRel(relsXml!)).toBe(false);

    // body 回滚:无锚点残留
    const docXmlContent = await readZipFile(zipBytes, 'word/document.xml');
    expect(docXmlContent).not.toContain('commentRangeStart');
    expect(docXmlContent).not.toContain('commentReference');

    await doc.dispose();
  });

  test('7. 第二次 addComment 失败回滚 → comments.xml 保留 1 条,CT/rels 不变,body 无多余锚点', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">p1</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t xml:space="preserve">p2</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-7', mode: 'inline' });

    // plan1: 成功 addComment → 建立 comments.xml
    await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'stays' })
      .apply();

    // 记录 CT / rels 状态
    const zip1Bytes = await doc.toZip();
    const ct1 = await readZipFile(zip1Bytes, '[Content_Types].xml');
    const rels1 = await readZipFile(zip1Bytes, 'word/_rels/document.xml.rels');
    expect(hasCtOverrideForComments(ct1!)).toBe(true);
    expect(hasCommentsRel(rels1!)).toBe(true);

    // plan2: addComment(AAAAAAA2 OK) + insertParagraphAfter('NOTEXIST' 失败) → 回滚
    const plan2 = doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA2' }, text: 'will rollback' })
      .insertParagraphAfter('NOTEXIST', { text: 'fail' });

    await expect(plan2.apply()).rejects.toMatchObject({ code: 'PARA_NOT_FOUND' });

    // comments.xml 仍有 1 条(回滚前的状态)
    const comments = await doc.comments.list();
    expect(comments).toHaveLength(1);
    expect(comments[0]!.text).toBe('stays');

    // CT / rels 不变(addComment 第二次不会重复添加 CT/rels)
    const zip2Bytes = await doc.toZip();
    const ct2 = await readZipFile(zip2Bytes, '[Content_Types].xml');
    const rels2 = await readZipFile(zip2Bytes, 'word/_rels/document.xml.rels');
    expect(hasCtOverrideForComments(ct2!)).toBe(true);
    expect(hasCommentsRel(rels2!)).toBe(true);

    await doc.dispose();
  });

  test('8. deleteComment + 抛错回滚 → 批注恢复', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">para</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-8', mode: 'inline' });

    // 先建批注
    const r = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'to be restored' })
      .apply();
    const id = r.createdComments[0]!.commentId;

    // plan: deleteComment(OK) + insertParagraphAfter('NOTEXIST' 失败) → 回滚
    const plan = doc.plan().deleteComment(id).insertParagraphAfter('NOTEXIST', { text: 'fail' });

    await expect(plan.apply()).rejects.toMatchObject({ code: 'PARA_NOT_FOUND' });

    // 批注恢复
    const comments = await doc.comments.list();
    expect(comments).toHaveLength(1);
    expect(comments[0]!.id).toBe(id);
    expect(comments[0]!.text).toBe('to be restored');

    await doc.dispose();
  });
});

// ─────────── toZip 一致性 ───────────

describe('Phase E 集成 —— toZip 一致性', () => {
  test('9. addComment → toZip → 回加载 → getComment 返回新批注,author/date/text 完整', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-9', mode: 'inline' });
    const result = await doc
      .plan({ author: 'bob', date: new Date('2026-04-22T12:00:00Z') })
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'detailed note' })
      .apply();

    const commentId = result.createdComments[0]!.commentId;
    const zipBytes = await doc.toZip();
    await doc.dispose();

    const doc2 = await WordDocument.fromZip(zipBytes, { docId: 'e2e-ph-9-reload', mode: 'inline' });
    const cm = await doc2.comments.get(commentId);
    expect(cm).not.toBeNull();
    expect(cm!.text).toBe('detailed note');
    expect(cm!.author).toBe('bob');
    // date が ISO 形式で含まれることを確認
    expect(cm!.date).toContain('2026-04-22');
    await doc2.dispose();
  });

  test('10. 删光所有批注 → toZip → 回加载 → comments.xml 不存在,listComments=[]', async () => {
    const buf = await makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">content</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t xml:space="preserve">more</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-10', mode: 'inline' });

    // 建 2 条批注
    const r = await doc
      .plan()
      .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'c1' })
      .addComment({ range: { paraId: 'AAAAAAA2' }, text: 'c2' })
      .apply();

    const id0 = r.createdComments[0]!.commentId;
    const id1 = r.createdComments[1]!.commentId;

    // 删除全部
    await doc.plan().deleteComment(id0).deleteComment(id1).apply();

    // 当前 listComments 为空
    expect(await doc.comments.list()).toHaveLength(0);

    // toZip → 回加载
    const zipBytes = await doc.toZip();
    await doc.dispose();

    const doc2 = await WordDocument.fromZip(zipBytes, {
      docId: 'e2e-ph-10-reload',
      mode: 'inline',
    });

    // listComments=[]
    expect(await doc2.comments.list()).toHaveLength(0);

    // 新行为：comments.xml 被移除（不再残留空 part）
    const commentsXml = await readZipFile(zipBytes, 'word/comments.xml');
    expect(commentsXml).toBeNull();

    // CT 无 Override
    const ctXml = await readZipFile(zipBytes, '[Content_Types].xml');
    expect(ctXml).not.toContain('/word/comments.xml');

    await doc2.dispose();
  });
});
