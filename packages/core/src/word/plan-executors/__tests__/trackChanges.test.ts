/**
 * 段级 executor 在 trackChanges=true 下的行为测试(#31)。
 *
 * 覆盖:
 *   - replaceParagraph:旧 runs 包 del,新 runs 包 ins
 *   - replaceTextInParagraph:同 run / 跨 run,匹配区间包 del,新 run 包 ins
 *   - insertParagraphBefore/After:新段内容 runs 包 ins,段 pPr 加段落 mark 插入标记
 *   - deleteParagraph:非 drop 情况下,runs 变 del,段 pPr 加段落 mark 删除标记,段本身保留
 *   - deleteParagraph drop + trackChanges=true → UNSUPPORTED_OOXML_FEATURE
 *   - w:id / w:author / w:date 属性:ctx.options 值正确使用
 *   - 多 op 同 plan:revision id 递增(通过独立调用验证)
 *   - trackChanges + ANCHOR_CONFLICT:锚点检测不在本 PR scope(#30),本文件只覆盖无锚点路径
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult, Op } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeReplaceParagraph, executeReplaceTextInParagraph } from '../paragraph-ops.js';
import { executeDeleteParagraph, executeInsertParagraph } from '../paragraph-structure-ops.js';

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

function makeDoc(bodyInner: string): {
  pkg: Package;
  ctx: ExecutorContext;
} {
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
    options: {
      trackChanges: true,
      author: 'alice',
      date: new Date('2026-04-23T10:00:00.000Z'),
    },
    result: emptyApplyResult(),
    opIndex: 0,
    docId: 'tc-test',
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

function paraByRef(pkg: Package, paraId: string): Element {
  const idx = buildIndices(pkg);
  return idx.paraIdx.get(paraId.toUpperCase())!;
}

function countInsDel(paragraph: Element): { ins: number; del: number } {
  return {
    ins: paragraph.getElementsByTagNameNS(W_NS, 'ins').length,
    del: paragraph.getElementsByTagNameNS(W_NS, 'del').length,
  };
}

// ─────────── replaceParagraph ───────────

describe('executeReplaceParagraph + trackChanges', () => {
  test('旧 runs 包成 w:del,新 runs 包成 w:ins,pPr 保留', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r><w:t>old</w:t></w:r>
      </w:p>`,
    );
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'new' });
    const p = paraByRef(pkg, 'AAAAAAA1');
    const { ins, del } = countInsDel(p);
    expect(ins).toBe(1);
    expect(del).toBe(1);
    // del 内应有 delText
    expect(p.getElementsByTagNameNS(W_NS, 'delText').length).toBe(1);
    expect(p.getElementsByTagNameNS(W_NS, 'delText')[0]!.textContent).toBe('old');
    // ins 内应有 w:t(新内容)
    const insEl = p.getElementsByTagNameNS(W_NS, 'ins')[0]!;
    expect(insEl.getElementsByTagNameNS(W_NS, 't')[0]!.textContent).toBe('new');
    // pPr 保留
    expect(p.getElementsByTagNameNS(W_NS, 'pStyle').length).toBe(1);
  });

  test('w:id / w:author / w:date 从 ctx.options 取', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`);
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'new' });
    const p = paraByRef(pkg, 'AAAAAAA1');
    const del = p.getElementsByTagNameNS(W_NS, 'del')[0]!;
    expect(del.getAttributeNS(W_NS, 'author') ?? del.getAttribute('w:author')).toBe('alice');
    expect(del.getAttributeNS(W_NS, 'date') ?? del.getAttribute('w:date')).toBe(
      '2026-04-23T10:00:00.000Z',
    );
    // del 和 ins 各自有独立 id(递增)
    const delId = parseInt(del.getAttributeNS(W_NS, 'id') ?? del.getAttribute('w:id') ?? '', 10);
    const ins = p.getElementsByTagNameNS(W_NS, 'ins')[0]!;
    const insId = parseInt(ins.getAttributeNS(W_NS, 'id') ?? ins.getAttribute('w:id') ?? '', 10);
    expect(insId).toBeGreaterThan(delId);
  });

  test('空 text:只有 del 没有 ins?其实 appendSegmentsToRun 对空文本建 placeholder <w:t/>', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`);
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: '' });
    const p = paraByRef(pkg, 'AAAAAAA1');
    // del 存在(旧内容)
    expect(p.getElementsByTagNameNS(W_NS, 'del').length).toBe(1);
    // ins 存在,但内部 t 为空
    const ins = p.getElementsByTagNameNS(W_NS, 'ins')[0]!;
    expect(ins).toBeDefined();
  });
});

// ─────────── replaceTextInParagraph ───────────

describe('executeReplaceTextInParagraph + trackChanges', () => {
  function replaceOp(
    paraId: string,
    search: string,
    replacement: string,
    scope: 'first' | 'all' = 'first',
  ): Extract<Op, { type: 'replaceTextInParagraph' }> {
    return { type: 'replaceTextInParagraph', paraId, search, replacement, scope };
  }

  test('同 run:match 区间包 del,新 run 包 ins', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello world</w:t></w:r></w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'world', 'everyone'));
    const p = paraByRef(pkg, 'AAAAAAA1');
    const { ins, del } = countInsDel(p);
    expect(ins).toBe(1);
    expect(del).toBe(1);
    // del 内有 delText 'world'
    const delTexts = p.getElementsByTagNameNS(W_NS, 'delText');
    expect(
      Array.from(delTexts)
        .map((el) => el.textContent)
        .join(''),
    ).toBe('world');
    // ins 内有 t 'everyone'
    const insEl = p.getElementsByTagNameNS(W_NS, 'ins')[0]!;
    expect(insEl.getElementsByTagNameNS(W_NS, 't')[0]!.textContent).toBe('everyone');
    // 前后原 run 保留('hello ' / '')
    // 文档顺序:[run-'hello '] [del] [ins]
  });

  test('空 replacement:只产生删除修订,不产生空插入修订', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>answer#text</w:t></w:r></w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', '#', ''));
    const p = paraByRef(pkg, 'AAAAAAA1');
    const { ins, del } = countInsDel(p);
    expect(del).toBe(1);
    expect(ins).toBe(0);
    const delTexts = p.getElementsByTagNameNS(W_NS, 'delText');
    expect(
      Array.from(delTexts)
        .map((t) => t.textContent)
        .join(''),
    ).toBe('#');
  });

  test('跨 run:多 run 包入同一 del', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>hello wo</w:t></w:r>
        <w:r><w:t>rld</w:t></w:r>
      </w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'world', 'everyone'));
    const p = paraByRef(pkg, 'AAAAAAA1');
    const del = p.getElementsByTagNameNS(W_NS, 'del')[0]!;
    expect(del).toBeDefined();
    // del 内含 2 个 delText('wo' + 'rld')
    const delTexts = del.getElementsByTagNameNS(W_NS, 'delText');
    expect(delTexts.length).toBe(2);
    expect(
      Array.from(delTexts)
        .map((t) => t.textContent)
        .join(''),
    ).toBe('world');
  });
});

// ─────────── insertParagraph ───────────

describe('executeInsertParagraph + trackChanges', () => {
  test('insertParagraphAfter:新段内容 runs 包 ins,pPr 加段落 mark 插入标记', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'AAAAAAA1',
        spec: { text: 'new-para' },
      },
      'after',
    );
    // 找新段(非 AAAAAAA1)
    const idx = buildIndices(pkg);
    const allIds = Array.from(idx.paraIdx.keys());
    const newId = allIds.find((id) => id !== 'AAAAAAA1')!;
    const newPara = idx.paraIdx.get(newId)!;

    // 新段应有 ins(包 runs)
    expect(newPara.getElementsByTagNameNS(W_NS, 'ins').length).toBeGreaterThanOrEqual(2);
    // 至少一个 ins 在 rPr 里(段落 mark)
    const rPr = newPara.getElementsByTagNameNS(W_NS, 'rPr')[0]!;
    expect(rPr).toBeDefined();
    expect(rPr.getElementsByTagNameNS(W_NS, 'ins').length).toBe(1);
  });

  test('ApplyResult.insertedParagraphs 正常记录', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'AAAAAAA1',
        spec: { text: 'x' },
      },
      'after',
    );
    expect(ctx.result.insertedParagraphs).toHaveLength(1);
    expect(ctx.result.insertedParagraphs[0]!.anchorParaId).toBe('AAAAAAA1');
    expect(ctx.result.insertedParagraphs[0]!.position).toBe('after');
  });
});

// ─────────── deleteParagraph ───────────

describe('executeDeleteParagraph + trackChanges', () => {
  test('无锚点:runs 变 del,段 pPr 加段落 mark 删除标记,段保留', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>content</w:t></w:r></w:p>`);
    executeDeleteParagraph(ctx, {
      type: 'deleteParagraph',
      paraId: 'AAAAAAA1',
      onAnchorConflict: 'error',
    });
    // 段**仍存在**(trackChanges 不 removeChild)
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(true);

    const p = idx.paraIdx.get('AAAAAAA1')!;
    // 有 del(包 runs)
    expect(p.getElementsByTagNameNS(W_NS, 'del').length).toBeGreaterThanOrEqual(2);
    // 段落 mark 删除:rPr 内有 del
    const rPr = p.getElementsByTagNameNS(W_NS, 'rPr')[0]!;
    expect(rPr).toBeDefined();
    expect(rPr.getElementsByTagNameNS(W_NS, 'del').length).toBe(1);
  });

  test('drop + trackChanges=true → UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="bm"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    try {
      executeDeleteParagraph(ctx, {
        type: 'deleteParagraph',
        paraId: 'AAAAAAA1',
        onAnchorConflict: 'drop',
      });
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
      expect((e as WordError).message).toMatch(/drop.*trackChanges/i);
    }
  });
});
