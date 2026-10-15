/**
 * Phase E worker 模式端到端 —— #36。
 *
 * 验证 createBookmark + addComment 跨 RPC 边界的组合场景:
 *   1. 跨 RPC createBookmark + addComment 同 plan → 两者都命中
 *   2. worker 模式下事务回滚正确:addComment + 失败 op → comments.xml 回滚
 *   3. worker → toZip → inline 加载一致:comments.xml / Content_Types / rels 三处都存在
 *
 * inline 集成覆盖见 plan-apply-phase-e-integration.test.ts。
 */
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { execSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { parseXmlDocument } from 'slimdom';
import { WordDocument } from '../word-document.js';
import { W14_NS, W_NS } from '../namespaces.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = resolve(__dirname, '../../../../..');
const bundleUrl = resolve(rootDir, 'packages/core/dist/kernel-worker.js');

// ─────────── 常量 ───────────

const COMMENTS_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

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

async function makeFixture(bodyInner: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  return zip.generateAsync({ type: 'uint8array' });
}

async function readZipFile(zipBytes: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(zipBytes);
  const file = zip.file(path);
  if (!file) return null;
  return file.async('string');
}

function hasCtOverrideForComments(ctXml: string): boolean {
  const dom = parseXmlDocument(ctXml) as unknown as Document;
  const overrides = dom.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
  for (let i = 0; i < overrides.length; i++) {
    if (overrides[i]!.getAttribute('PartName') === '/word/comments.xml') return true;
  }
  return false;
}

function hasCommentsRel(relsXml: string): boolean {
  const dom = parseXmlDocument(relsXml) as unknown as Document;
  const rels = dom.getElementsByTagNameNS(RELS_NS, 'Relationship');
  for (let i = 0; i < rels.length; i++) {
    if (rels[i]!.getAttribute('Type') === COMMENTS_RELATIONSHIP_TYPE) return true;
  }
  return false;
}

// ─────────── beforeAll / afterAll ───────────

beforeAll(() => {
  execSync('corepack pnpm --filter @coco-vexel/coco-office build', { cwd: rootDir, stdio: 'pipe' });
  WordDocument.setWorkerBundleUrl(bundleUrl);
}, 120_000);

afterAll(() => {
  WordDocument._resetDefaults();
});

const WORKER_TEST_TIMEOUT = 30_000;

// ─────────── worker 端到端 ───────────

describe('Phase E worker 端到端 —— 跨 RPC 组合 (#36)', () => {
  test(
    '1. 跨 RPC createBookmark + addComment 同 plan → 两者都命中',
    async () => {
      const buf = await makeFixture(
        `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">worker test paragraph</w:t></w:r></w:p>`,
      );
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-w-1', mode: 'worker' });

      const result = await doc
        .plan({ author: 'worker-alice', date: new Date('2026-04-22T10:00:00Z') })
        .createBookmark({ name: 'workerBm', range: { paraId: 'AAAAAAA1' } })
        .addComment({
          range: { paraId: 'AAAAAAA1' },
          text: 'worker comment on same para',
        })
        .apply();

      // ApplyResult 包含书签和批注
      expect(result.createdBookmarks).toEqual([
        { opIndex: 0, name: 'workerBm', range: { fromParaId: 'AAAAAAA1', toParaId: 'AAAAAAA1' } },
      ]);
      expect(result.createdComments).toHaveLength(1);
      const commentId = result.createdComments[0]!.commentId;

      // listBookmarks 命中
      const bookmarks = await doc.bookmarks.list();
      const bm = bookmarks.find((b) => b.name === 'workerBm');
      expect(bm).toBeDefined();
      expect(bm!.rangeKind).toBe('paragraph');

      // listComments 命中
      const comments = await doc.comments.list();
      const cm = comments.find((c) => c.id === commentId);
      expect(cm).toBeDefined();
      expect(cm!.paraId?.toUpperCase()).toBe('AAAAAAA1');

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    '2. worker 模式下事务回滚:首次 addComment 失败 → comments.xml 不存在',
    async () => {
      const buf = await makeFixture(
        `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">text</w:t></w:r></w:p>`,
      );
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-w-2', mode: 'worker' });

      // addComment(AAAAAAA1 OK) + insertParagraphAfter('NOTEXIST' → PARA_NOT_FOUND) → 回滚
      const plan = doc
        .plan()
        .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'will rollback' })
        .insertParagraphAfter('NOTEXIST', { text: 'fail' });

      await expect(plan.apply()).rejects.toMatchObject({ code: 'PARA_NOT_FOUND' });

      // listComments 为空(回滚成功)
      const comments = await doc.comments.list();
      expect(comments).toHaveLength(0);

      // toZip 验证:comments.xml 不应存在
      const zipBytes = await doc.toZip();
      const commentsXml = await readZipFile(zipBytes, 'word/comments.xml');
      expect(commentsXml).toBeNull();

      // Content_Types 无 Override
      const ctXml = await readZipFile(zipBytes, '[Content_Types].xml');
      expect(ctXml).not.toBeNull();
      expect(hasCtOverrideForComments(ctXml!)).toBe(false);

      // rels 无 comments 关系
      const relsXml = await readZipFile(zipBytes, 'word/_rels/document.xml.rels');
      expect(relsXml).not.toBeNull();
      expect(hasCommentsRel(relsXml!)).toBe(false);

      await doc.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );

  test(
    '3. worker → toZip → inline 加载一致:addComment 后三处 part 均存在',
    async () => {
      const buf = await makeFixture(
        `<w:p w14:paraId="AAAAAAA1"><w:r><w:t xml:space="preserve">reload test</w:t></w:r></w:p>`,
      );
      const doc = await WordDocument.fromZip(buf, { docId: 'e2e-ph-w-3', mode: 'worker' });

      const result = await doc
        .plan({ author: 'carol', date: new Date('2026-04-22T14:00:00Z') })
        .addComment({ range: { paraId: 'AAAAAAA1' }, text: 'worker note' })
        .apply();

      const commentId = result.createdComments[0]!.commentId;
      const zipBytes = await doc.toZip();
      await doc.dispose();

      // inline 重加载
      const doc2 = await WordDocument.fromZip(zipBytes, {
        docId: 'e2e-ph-w-3-reload',
        mode: 'inline',
      });

      // getComment 完整
      const cm = await doc2.comments.get(commentId);
      expect(cm).not.toBeNull();
      expect(cm!.text).toBe('worker note');
      expect(cm!.author).toBe('carol');
      expect(cm!.date).toContain('2026-04-22');

      // toZip 验证三处 part 均存在
      const ctXml = await readZipFile(zipBytes, '[Content_Types].xml');
      expect(hasCtOverrideForComments(ctXml!)).toBe(true);

      const relsXml = await readZipFile(zipBytes, 'word/_rels/document.xml.rels');
      expect(hasCommentsRel(relsXml!)).toBe(true);

      const commentsXml = await readZipFile(zipBytes, 'word/comments.xml');
      expect(commentsXml).not.toBeNull();
      const dom = parseXmlDocument(commentsXml!) as unknown as Document;
      const commentEls = dom.getElementsByTagNameNS(W_NS, 'comment');
      expect(commentEls.length).toBe(1);

      await doc2.dispose();
    },
    WORKER_TEST_TIMEOUT,
  );
});
