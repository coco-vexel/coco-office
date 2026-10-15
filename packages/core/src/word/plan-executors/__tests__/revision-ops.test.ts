/**
 * revision-ops executor 纯函数测试(#32,Phase D 收尾)。
 *
 * 覆盖矩阵:
 *   - Accept / Reject × <w:ins> / <w:del> wrapper × 正文(段内)场景
 *   - Accept / Reject × 段落 mark(pPr/rPr/<w:ins>|del>)场景
 *   - 三档 scope:all / paragraph / revision
 *   - Tier 2 属性修订(rPrChange / pPrChange 等)抛 UNSUPPORTED_OOXML_FEATURE
 *   - PARA_NOT_FOUND / REVISION_NOT_FOUND
 *   - ApplyResult.acceptedRevisions / rejectedRevisions 计数
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult, Op } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeAcceptRevisions, executeRejectRevisions } from '../revision-ops.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function putXml(pkg: Package, path: string, xml: string): void {
  setPart(
    pkg,
    createXmlPart({
      path,
      contentType: DOC_CTYPE,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
      originalBytes: new TextEncoder().encode(xml),
    }),
  );
}

function makeDoc(bodyInner: string): { pkg: Package; ctx: ExecutorContext } {
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`,
  );
  const indices = buildIndices(pkg);
  const ctx: ExecutorContext = {
    pkg,
    indices,
    options: { trackChanges: false, author: '', date: new Date() },
    result: emptyApplyResult(),
    opIndex: 0,
    docId: 'test-doc',
  };
  return { pkg, ctx };
}

function emptyApplyResult(): ApplyResult {
  return {
    insertedParagraphs: [],
    insertedBodyStartParaIds: [],
    createdComments: [],
    createdBookmarks: [],
    acceptedRevisions: 0,
    rejectedRevisions: 0,
    replacedTableCells: 0,
  };
}

function paraXml(pkg: Package, paraId: string): string {
  const idx = buildIndices(pkg);
  const el = idx.paraIdx.get(paraId.toUpperCase())!;
  return serializeToWellFormedString(el as never);
}

function paraCount(pkg: Package): number {
  const idx = buildIndices(pkg);
  return idx.paraIdx.size;
}

function acceptOp(
  scope: Extract<Op, { type: 'acceptRevisions' }>['scope'],
): Extract<Op, { type: 'acceptRevisions' }> {
  return { type: 'acceptRevisions', scope };
}

function rejectOp(
  scope: Extract<Op, { type: 'rejectRevisions' }>['scope'],
): Extract<Op, { type: 'rejectRevisions' }> {
  return { type: 'rejectRevisions', scope };
}

// ─────────── Accept <w:ins>(正文)─────────────

describe('acceptRevisions —— <w:ins> wrapper', () => {
  test('剥 wrapper,保留内容(w:t 不变)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="5" w:author="alice" w:date="2026-04-23T00:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins/);
    expect(xml).toMatch(/<w:t[^>]*>inserted<\/w:t>/);
    expect(ctx.result.acceptedRevisions).toBe(1);
    expect(ctx.result.rejectedRevisions).toBe(0);
  });
});

describe('rejectRevisions —— <w:ins> wrapper', () => {
  test('剥 wrapper + 内容一起删', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>before </w:t></w:r>
        <w:ins w:id="5" w:author="x" w:date="2026-04-23T00:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
        <w:r><w:t> after</w:t></w:r>
      </w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins/);
    expect(xml).not.toMatch(/inserted/);
    expect(xml).toMatch(/before/);
    expect(xml).toMatch(/after/);
    expect(ctx.result.rejectedRevisions).toBe(1);
  });
});

// ─────────── Accept / Reject <w:del>(正文)───────────

describe('acceptRevisions —— <w:del> wrapper', () => {
  test('剥 wrapper + 删内容(内部 delText 一起删)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:del w:id="7" w:author="x" w:date="2026-04-23T00:00:00Z">
          <w:r><w:delText>deleted</w:delText></w:r>
        </w:del>
        <w:r><w:t>kept</w:t></w:r>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:del/);
    expect(xml).not.toMatch(/delText/);
    expect(xml).not.toMatch(/deleted/);
    expect(xml).toMatch(/kept/);
    expect(ctx.result.acceptedRevisions).toBe(1);
  });
});

describe('rejectRevisions —— <w:del> wrapper', () => {
  test('剥 wrapper + delText 转回 t(恢复被删内容)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:del w:id="7" w:author="x" w:date="2026-04-23T00:00:00Z">
          <w:r><w:delText xml:space="preserve">restored</w:delText></w:r>
        </w:del>
      </w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:del/);
    expect(xml).not.toMatch(/delText/);
    expect(xml).toMatch(/<w:t[^>]*xml:space="preserve"[^>]*>restored<\/w:t>/);
    expect(ctx.result.rejectedRevisions).toBe(1);
  });
});

// ─────────── 段落 mark ───────────

describe('段落 mark —— pPr/rPr/<w:ins>', () => {
  test('accept ins mark:删 mark,段保留', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:rPr><w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>content</w:t></w:r>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    expect(paraCount(pkg)).toBe(1); // 段仍在
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins/);
    expect(ctx.result.acceptedRevisions).toBe(1);
  });

  test('reject ins mark:删整段', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>keep</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2">
        <w:pPr><w:rPr><w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>toDelete</w:t></w:r>
      </w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    expect(paraCount(pkg)).toBe(1);
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA2')).toBe(false);
    expect(ctx.result.rejectedRevisions).toBe(1);
  });
});

describe('段落 mark —— pPr/rPr/<w:del>', () => {
  test('accept del mark:删整段(段合并简化为删)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:rPr><w:del w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>merged</w:t></w:r>
      </w:p>
      <w:p w14:paraId="AAAAAAA2"><w:r><w:t>next</w:t></w:r></w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    expect(paraCount(pkg)).toBe(1);
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(false);
    expect(ctx.result.acceptedRevisions).toBe(1);
  });

  test('reject del mark:删 mark,段保留', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:rPr><w:del w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>keep</w:t></w:r>
      </w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    expect(paraCount(pkg)).toBe(1);
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:del/);
    expect(ctx.result.rejectedRevisions).toBe(1);
  });
});

// ─────────── scope ───────────

describe('scope 三档', () => {
  test("scope='paragraph':只处理指定段内的修订", () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
      </w:p>
      <w:p w14:paraId="AAAAAAA2">
        <w:ins w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>b</w:t></w:r></w:ins>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'paragraph', paraId: 'AAAAAAA1' }));
    const xml1 = paraXml(pkg, 'AAAAAAA1');
    const xml2 = paraXml(pkg, 'AAAAAAA2');
    expect(xml1).not.toMatch(/<w:ins/); // 处理过
    expect(xml2).toMatch(/<w:ins/); // 未动
    expect(ctx.result.acceptedRevisions).toBe(1);
  });

  test("scope='paragraph':paraId 不存在 → PARA_NOT_FOUND", () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    try {
      executeAcceptRevisions(ctx, acceptOp({ scope: 'paragraph', paraId: 'NONE' }));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('PARA_NOT_FOUND');
    }
  });

  test("scope='revision':按 w:id 精确匹配", () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="5" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
        <w:ins w:id="6" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>b</w:t></w:r></w:ins>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'revision', revId: '5' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    // id=5 被剥,a 保留;id=6 未动
    expect(xml).toMatch(/<w:ins[^>]+w:id="6"/);
    expect((xml.match(/<w:ins/g) ?? []).length).toBe(1);
    expect(ctx.result.acceptedRevisions).toBe(1);
  });

  test("scope='revision':revId 无匹配 → REVISION_NOT_FOUND", () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="5" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
      </w:p>`,
    );
    try {
      executeAcceptRevisions(ctx, acceptOp({ scope: 'revision', revId: '99' }));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('REVISION_NOT_FOUND');
    }
  });

  test("scope='all':全文 ins/del 都被处理,多 part 也覆盖", () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
        <w:del w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:delText>b</w:delText></w:r></w:del>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins|<w:del/);
    expect(ctx.result.acceptedRevisions).toBe(2);
  });
});

// ─────────── Tier 2 属性修订拒绝 ───────────

describe('Tier 2 属性级修订', () => {
  test('扫到 rPrChange → UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r>
          <w:rPr>
            <w:rPrChange w:id="9" w:author="x" w:date="2026-04-23T00:00:00Z">
              <w:rPr/>
            </w:rPrChange>
          </w:rPr>
          <w:t>x</w:t>
        </w:r>
      </w:p>`,
    );
    try {
      executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
      expect((e as WordError).message).toMatch(/rPrChange|#33/);
    }
  });

  test('扫到 pPrChange → UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr>
          <w:pPrChange w:id="9" w:author="x" w:date="2026-04-23T00:00:00Z">
            <w:pPr/>
          </w:pPrChange>
        </w:pPr>
        <w:r><w:t>x</w:t></w:r>
      </w:p>`,
    );
    try {
      executeAcceptRevisions(ctx, acceptOp({ scope: 'paragraph', paraId: 'AAAAAAA1' }));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
    }
  });
});

// ─────────── 边界 ───────────

describe('边界', () => {
  test('空文档(无修订):成功,count=0', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    expect(ctx.result.acceptedRevisions).toBe(0);
  });

  test('同段内多个 ins + del 一次 accept 处理完', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
        <w:r><w:t>-</w:t></w:r>
        <w:ins w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>b</w:t></w:r></w:ins>
        <w:del w:id="3" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:delText>c</w:delText></w:r></w:del>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins|<w:del/);
    // a 保留,- 保留,b 保留,c 被删(accept del)
    expect(xml).toMatch(/a/);
    expect(xml).toMatch(/-/);
    expect(xml).toMatch(/b/);
    expect(xml).not.toMatch(/>c</);
    expect(ctx.result.acceptedRevisions).toBe(3);
  });
});
