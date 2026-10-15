/**
 * paragraph-structure-ops executor 纯函数测试(#29)。
 *
 * 覆盖:
 *   executeInsertParagraph:
 *     - before / after 位置
 *     - spec.style='inherit':anchor 有 pPr 时克隆;anchor 无 pPr 时不加
 *     - spec.style='Heading1':生成新 pPr + pStyle
 *     - spec.text 含 \n / \t:生成 <w:br/> / <w:tab/>
 *     - ApplyResult.insertedParagraphs 字段完整
 *     - PARA_NOT_FOUND(anchor 不存在)
 *     - trackChanges=true 抛 UNSUPPORTED_OOXML_FEATURE(含 #31)
 *     - newParaId 唯一(不与已有 paraIdx key 冲突)
 *     - spec.runFormat='inherit':克隆 anchor 首 run rPr 到新段 run
 *     - spec.runFormat='inherit':anchor 无直属 run → 退化为裸 run
 *     - spec.runFormat='inherit' + style='inherit':pPr 和 rPr 都继承
 *     - spec.runFormat='inherit' + color:rPr 继承 + color 覆盖
 *     - 不传 runFormat:行为与现状完全一致(回归)
 *
 *   executeDeleteParagraph:
 *     - 无锚点的段被删除
 *     - 有 bookmarkStart + onAnchorConflict='error' → ANCHOR_CONFLICT(含 affectedBookmarks)
 *     - 有 commentRangeStart + onAnchorConflict='error' → ANCHOR_CONFLICT(含 affectedComments)
 *     - onAnchorConflict='drop' + 有锚点 → UNSUPPORTED_OOXML_FEATURE(含 #34/#35)
 *     - onAnchorConflict='drop' + 无锚点 → 成功删除
 *     - PARA_NOT_FOUND
 *     - trackChanges=true → UNSUPPORTED_OOXML_FEATURE
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult, Op } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeInsertParagraph, executeDeleteParagraph } from '../paragraph-structure-ops.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function putXml(pkg: Package, path: string, xml: string, contentType = DOC_CTYPE): void {
  setPart(
    pkg,
    createXmlPart({
      path,
      contentType,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
      originalBytes: new TextEncoder().encode(xml),
    }),
  );
}

function makeDoc(bodyInner: string): {
  pkg: Package;
  indices: ReturnType<typeof buildIndices>;
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
    options: { trackChanges: false, author: '', date: new Date('2026-04-22T00:00:00Z') },
    result: emptyApplyResult(),
    opIndex: 0,
    docId: 'test-doc',
  };
  return { pkg, indices, ctx };
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

/** 序列化某 paraId 对应段落 XML */
function paraXml(pkg: Package, paraId: string): string {
  const idx = buildIndices(pkg);
  const el = idx.paraIdx.get(paraId.toUpperCase())!;
  return serializeToWellFormedString(el as never);
}

/** 取段落所有 <w:t> 拼接成纯文本 */
function paraText(pkg: Package, paraId: string): string {
  const idx = buildIndices(pkg);
  const el = idx.paraIdx.get(paraId.toUpperCase())!;
  const ts = el.getElementsByTagNameNS(W_NS, 't');
  let s = '';
  for (let i = 0; i < ts.length; i++) s += ts[i]!.textContent ?? '';
  return s;
}

/** 取 body 内所有 w14:paraId(按 DOM 顺序)*/
function bodyParaIds(pkg: Package): string[] {
  const idx = buildIndices(pkg);
  return [...idx.paraIdx.keys()];
}

// ─────────── executeInsertParagraph ───────────

describe('executeInsertParagraph', () => {
  function insertBeforeOp(
    anchorParaId: string,
    text: string,
    style?: string,
  ): Extract<Op, { type: 'insertParagraphBefore' }> {
    return {
      type: 'insertParagraphBefore',
      anchorParaId,
      spec: { text, ...(style !== undefined ? { style } : {}) },
    };
  }

  function insertAfterOp(
    anchorParaId: string,
    text: string,
    style?: string,
  ): Extract<Op, { type: 'insertParagraphAfter' }> {
    return {
      type: 'insertParagraphAfter',
      anchorParaId,
      spec: { text, ...(style !== undefined ? { style } : {}) },
    };
  }

  test('insertParagraphAfter:新段在 anchor 之后', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>first</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>second</w:t></w:r></w:p>`,
    );
    executeInsertParagraph(ctx, insertAfterOp('AAAAAAA1', 'inserted'), 'after');
    const ids = bodyParaIds(pkg);
    // AAAAAAA1 之后应是新段,然后才是 AAAAAAA2
    expect(ids.indexOf(ctx.result.insertedParagraphs[0]!.newParaId)).toBe(1);
    expect(ids.indexOf('AAAAAAA2')).toBe(2);
    expect(paraText(pkg, ctx.result.insertedParagraphs[0]!.newParaId)).toBe('inserted');
  });

  test('insertParagraphBefore:新段在 anchor 之前', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>first</w:t></w:r></w:p>`);
    executeInsertParagraph(ctx, insertBeforeOp('AAAAAAA1', 'before-first'), 'before');
    const ids = bodyParaIds(pkg);
    // 新段在 AAAAAAA1 之前
    expect(ids[0]).toBe(ctx.result.insertedParagraphs[0]!.newParaId);
    expect(ids[1]).toBe('AAAAAAA1');
    expect(paraText(pkg, ctx.result.insertedParagraphs[0]!.newParaId)).toBe('before-first');
  });

  test('spec.style=inherit(默认):anchor 有 pPr → 克隆到新段', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r><w:t>anchor</w:t></w:r>
      </w:p>`,
    );
    executeInsertParagraph(ctx, insertAfterOp('AAAAAAA1', 'new', 'inherit'), 'after');
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toMatch(/<w:pStyle\s+w:val="Heading1"/);
  });

  test('spec.style=inherit:anchor 无 pPr → 新段也无 pPr', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(ctx, insertAfterOp('AAAAAAA1', 'new', 'inherit'), 'after');
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).not.toMatch(/pPr/);
  });

  test('spec.style 为字符串:生成新 pPr + pStyle', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(ctx, insertAfterOp('AAAAAAA1', 'styled', 'MyStyle'), 'after');
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toMatch(/pPr/);
    expect(xml).toMatch(/pStyle/);
    expect(xml).toMatch(/MyStyle/);
  });

  test('spec.text 含 \\n → <w:br/>,含 \\t → <w:tab/>', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(ctx, insertAfterOp('AAAAAAA1', 'a\nb\tc'), 'after');
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toMatch(/<w:br\s*\/>/);
    expect(xml).toMatch(/<w:tab\s*\/>/);
    expect(paraText(pkg, newId)).toBe('abc');
  });

  test('ApplyResult.insertedParagraphs 字段完整', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    ctx.opIndex = 3;
    executeInsertParagraph(ctx, insertAfterOp('AAAAAAA1', 'new'), 'after');
    const entry = ctx.result.insertedParagraphs[0]!;
    expect(entry.opIndex).toBe(3);
    expect(entry.anchorParaId).toBe('AAAAAAA1');
    expect(entry.position).toBe('after');
    expect(entry.newParaId).toMatch(/^[0-9A-F]{8}$/);
  });

  test('PARA_NOT_FOUND:anchor 不存在时抛错', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(() => executeInsertParagraph(ctx, insertAfterOp('NOPE', 'new'), 'after')).toThrow(
      expect.objectContaining({ code: 'PARA_NOT_FOUND' }) as unknown as Error,
    );
  });

  // trackChanges=true 的覆盖在 trackChanges.test.ts(#31)

  test('newParaId 唯一:不与已有 paraIdx key 冲突', () => {
    // 多段测试:每次 insert 产生的 newParaId 不重复
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>a</w:t></w:r></w:p>`);
    const ids = new Set<string>();
    for (let i = 0; i < 20; i++) {
      const freshCtx = { ...ctx, indices: buildIndices(ctx.pkg), result: emptyApplyResult() };
      executeInsertParagraph(freshCtx, insertAfterOp('AAAAAAA1', 'x'), 'after');
      const newId = freshCtx.result.insertedParagraphs[0]!.newParaId;
      expect(newId).not.toBe('AAAAAAA1');
      ids.add(newId);
    }
    // 20 次生成应全部唯一(概率极低碰撞)
    expect(ids.size).toBe(20);
  });

  // ─────────── runFormat 测试组 ───────────

  test('runFormat=inherit:克隆 anchor 首 run 的 rPr(rFonts/sz/b)到新段 run', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r>
          <w:rPr>
            <w:rFonts w:eastAsia="黑体"/>
            <w:sz w:val="32"/>
            <w:b/>
          </w:rPr>
          <w:t>anchor text</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'AAAAAAA1',
        spec: { text: '新', style: 'inherit', runFormat: 'inherit' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toMatch(/rFonts/);
    expect(xml).toMatch(/黑体/);
    expect(xml).toMatch(/w:sz/);
    expect(xml).toMatch(/32/);
    expect(xml).toMatch(/<w:b\b/);
  });

  test('runFormat=inherit:anchor 无直属 run(空段)→ 退化为裸 run,不报错', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      </w:p>`,
    );
    expect(() =>
      executeInsertParagraph(
        ctx,
        {
          type: 'insertParagraphAfter',
          anchorParaId: 'AAAAAAA1',
          spec: { text: '新段', runFormat: 'inherit' },
        },
        'after',
      ),
    ).not.toThrow();
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 裸 run:没有 rPr(非 color 添加)
    expect(xml).not.toMatch(/rPr/);
    expect(paraText(pkg, newId)).toBe('新段');
  });

  test('runFormat=inherit + style=inherit:pPr 和 rPr 都继承(正交)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr>
          <w:t>标题文字</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'AAAAAAA1',
        spec: { text: '复制段', style: 'inherit', runFormat: 'inherit' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // pPr 继承:含 Heading2
    expect(xml).toMatch(/Heading2/);
    // rPr 继承:含宋体
    expect(xml).toMatch(/宋体/);
  });

  test('runFormat=inherit + color 共存:rPr 继承 + color 覆盖(color 优先)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr>
          <w:t>anchor</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'AAAAAAA1',
        spec: { text: '带色', runFormat: 'inherit', color: 'FF0000' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 继承的字体仍在
    expect(xml).toMatch(/黑体/);
    // color 被写入
    expect(xml).toMatch(/FF0000/);
    expect(xml).toMatch(/w:color/);
  });

  test('不传 runFormat:行为与现状完全一致(裸 run,无 rPr)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr>
          <w:t>anchor</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(ctx, insertAfterOp('AAAAAAA1', '新段'), 'after');
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 没有继承 rPr
    expect(xml).not.toMatch(/rPr/);
    expect(paraText(pkg, newId)).toBe('新段');
  });

  test('runFormat=default:与不传时行为完全相同(裸 run)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r>
          <w:rPr><w:sz w:val="28"/></w:rPr>
          <w:t>anchor</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'AAAAAAA1',
        spec: { text: '新段', runFormat: 'default' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).not.toMatch(/rPr/);
  });

  // ─────────── #12 follow-up 白盒测试 ───────────

  test('runFormat=inherit: anchor 首 run 有 run 但无 rPr → 退化裸 run 不报错', () => {
    // spec 行为契约第 3 条第二子条件:anchor 有直属 run 但该 run 无 rPr → 退化为裸 run
    // (现有测试只覆盖了"无直属 run"的子条件)
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r>
          <w:t>anchor no rPr run</w:t>
        </w:r>
      </w:p>`,
    );
    expect(() =>
      executeInsertParagraph(
        ctx,
        {
          type: 'insertParagraphAfter',
          anchorParaId: 'AAAAAAA1',
          spec: { text: '新段', runFormat: 'inherit' },
        },
        'after',
      ),
    ).not.toThrow();
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 退化为裸 run:没有 rPr(无 color 追加)
    expect(xml).not.toMatch(/w:rPr/);
    expect(paraText(pkg, newId)).toBe('新段');
  });

  test('runFormat=inherit + color: anchor rPr 已含 color → 被 color 参数覆盖(只有一个 w:color)', () => {
    // 测 applyColorToDirectRuns 的"覆盖"分支:
    // anchor rPr 已含 <w:color w:val="0000FF"/>，color 参数传 FF0000 → 新段 color 是 FF0000 且只有一个 <w:color>
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r>
          <w:rPr><w:color w:val="0000FF"/></w:rPr>
          <w:t>anchor</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'AAAAAAA1',
        spec: { text: '新段', runFormat: 'inherit', color: 'FF0000' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 最终 color 是 FF0000（覆盖了继承的 0000FF）
    expect(xml).toContain('FF0000');
    // 旧的 color 值不再出现
    expect(xml).not.toContain('0000FF');
    // 整个段落里只有一个 <w:color>
    expect((xml.match(/<w:color\b/g) ?? []).length).toBe(1);
  });

  // ─────────── styleSource 测试组（task-core-style-source）───────────

  test('styleSource 基础:pPr + rPr 都来自 styleSource 段，不是 anchor', () => {
    // 测试 1：anchor 是普通段（left/宋体），styleSource 是标题段（center/黑体/sz32）
    // 新段应继承 styleSource 的样式
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="ANCHOR01">
        <w:pPr><w:jc w:val="left"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr>
          <w:t>anchor</w:t>
        </w:r>
      </w:p>
      <w:p w14:paraId="SOURCE01">
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:eastAsia="黑体"/>
            <w:sz w:val="32"/>
          </w:rPr>
          <w:t>source</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'ANCHOR01',
        spec: { text: '新段', styleSource: 'SOURCE01' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // pPr 来自 styleSource（center）
    expect(xml).toMatch(/jc.*center|center.*jc/);
    // rPr 来自 styleSource（黑体/sz32）
    expect(xml).toMatch(/黑体/);
    expect(xml).toMatch(/w:sz/);
    expect(xml).toMatch(/32/);
  });

  test('styleSource 不继承 anchor（核心）:新段样式是 styleSource，不含 anchor 特征', () => {
    // 测试 2：这是设计目标的核心——全文重组时正文段不能继承上一段（节标题）的样式
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="ANCHOR02">
        <w:pPr><w:jc w:val="left"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr>
          <w:t>anchor-宋体-left</w:t>
        </w:r>
      </w:p>
      <w:p w14:paraId="SOURCE02">
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr>
          <w:t>source-黑体-center</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'ANCHOR02',
        spec: { text: '新段', styleSource: 'SOURCE02' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 含 styleSource 的特征
    expect(xml).toMatch(/黑体/);
    expect(xml).toMatch(/center/);
    // 不含 anchor 的特征（宋体 + left 不应出现在新段 pPr/rPr）
    expect(xml).not.toMatch(/宋体/);
    // jc=left 也不应出现（只有 center）
    expect((xml.match(/jc/g) ?? []).length).toBe(1);
    expect(xml).not.toMatch(/w:val="left"/);
  });

  test('styleSource 段不存在 → 降级回退到 anchor pPr（绝不崩、不丢段）', () => {
    // 测试 3（健壮性降级，原行为是抛 PARA_NOT_FOUND，现改为降级回退 anchor pPr）。
    // anchor 带可识别的 pPr（黑体 + center），styleSource 指向不存在的段时新段应继承 anchor 样式而非崩。
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="ANCHOR03"><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="黑体"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="黑体"/></w:rPr><w:t>anchor</w:t></w:r></w:p>`,
    );
    expect(() =>
      executeInsertParagraph(
        ctx,
        {
          type: 'insertParagraphAfter',
          anchorParaId: 'ANCHOR03',
          spec: { text: '新段', styleSource: 'NOPE0000' },
        },
        'after',
      ),
    ).not.toThrow();
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toMatch(/新段/); // 段内容保留，没有丢段
    expect(xml).toMatch(/center/); // 回退继承了 anchor 的 pPr（jc=center）
    expect(xml).toMatch(/黑体/); // 回退继承了 anchor 的 rPr（黑体）
  });

  test('styleSource 退化：源段无 pPr → 新段不加 pPr；源段无 run rPr → 裸 run', () => {
    // 测试 4：styleSource 段无 pPr、无 rPr 时退化不报错
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="ANCHOR04">
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr>
          <w:t>anchor</w:t>
        </w:r>
      </w:p>
      <w:p w14:paraId="EMPTY004">
        <w:r><w:t>source-no-ppr-no-rpr</w:t></w:r>
      </w:p>`,
    );
    expect(() =>
      executeInsertParagraph(
        ctx,
        {
          type: 'insertParagraphAfter',
          anchorParaId: 'ANCHOR04',
          spec: { text: '新段', styleSource: 'EMPTY004' },
        },
        'after',
      ),
    ).not.toThrow();
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 无 pPr（源段没有）
    expect(xml).not.toMatch(/pPr/);
    // 无 rPr（源段首 run 无 rPr）
    expect(xml).not.toMatch(/rPr/);
    expect(paraText(pkg, newId)).toBe('新段');
  });

  test('styleSource + color 共存：先克隆 rPr，再 color 覆盖（color 优先）', () => {
    // 测试 5：styleSource rPr 含 color=0000FF，spec color=FF0000 → 新段 color=FF0000 且唯一
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="ANCHOR05"><w:r><w:t>anchor</w:t></w:r></w:p>
      <w:p w14:paraId="SOURCE05">
        <w:r>
          <w:rPr><w:color w:val="0000FF"/></w:rPr>
          <w:t>source</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'ANCHOR05',
        spec: { text: '新段', styleSource: 'SOURCE05', color: 'FF0000' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 最终 color 是 FF0000（spec color 覆盖 styleSource 的 0000FF）
    expect(xml).toContain('FF0000');
    expect(xml).not.toContain('0000FF');
    // 只有一个 <w:color>
    expect((xml.match(/<w:color\b/g) ?? []).length).toBe(1);
  });

  test('styleSource 优先级：同时传 style + runFormat + styleSource → 以 styleSource 为准', () => {
    // 测试 6：style:'inherit' 和 runFormat:'inherit' 都被 styleSource 覆盖
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="ANCHOR06">
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr>
          <w:t>anchor</w:t>
        </w:r>
      </w:p>
      <w:p w14:paraId="SOURCE06">
        <w:pPr><w:pStyle w:val="BodyText"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr>
          <w:t>source</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'ANCHOR06',
        spec: { text: '新段', style: 'inherit', runFormat: 'inherit', styleSource: 'SOURCE06' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 来自 styleSource（BodyText + 黑体）
    expect(xml).toMatch(/BodyText/);
    expect(xml).toMatch(/黑体/);
    // 不来自 anchor（Heading1 + 宋体）
    expect(xml).not.toMatch(/Heading1/);
    expect(xml).not.toMatch(/宋体/);
  });

  test('styleSource=anchor：退化为 #12 inherit 效果（pPr+rPr 都来自 anchor）', () => {
    // 测试 7：styleSource 可以等于 anchor，这是 runFormat:'inherit'+style:'inherit' 的特例
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="ANCHOR07">
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr>
          <w:t>anchor</w:t>
        </w:r>
      </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'ANCHOR07',
        spec: { text: '新段', styleSource: 'ANCHOR07' },
      },
      'after',
    );
    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // pPr 来自 anchor（center）
    expect(xml).toMatch(/center/);
    // rPr 来自 anchor（黑体）
    expect(xml).toMatch(/黑体/);
  });

  test('全文重组场景（关键）：3 种角色交替插入，每段指定各自 styleSource，不串味', () => {
    // 测试 8：h1/h2/body 规范段在文档末尾，按交替顺序重组：h1→body→h2→body
    // anchor 滚动（每次用上一个插入的段作为 anchor）
    // 期望：每段样式正确（不继承 anchor 的样式）
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="TMPANCHR">
        <w:r><w:t>start</w:t></w:r>
      </w:p>
      <w:p w14:paraId="H1SPEC01">
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="36"/></w:rPr>
          <w:t>h1-spec</w:t>
        </w:r>
      </w:p>
      <w:p w14:paraId="H2SPEC01">
        <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="28"/></w:rPr>
          <w:t>h2-spec</w:t>
        </w:r>
      </w:p>
      <w:p w14:paraId="BDYSPEC1">
        <w:pPr><w:jc w:val="both"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr>
          <w:t>body-spec</w:t>
        </w:r>
      </w:p>`,
    );

    // 插入顺序：h1, body, h2, body（anchor 滚动）
    executeInsertParagraph(
      ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: 'TMPANCHR',
        spec: { text: '第一章', styleSource: 'H1SPEC01' },
      },
      'after',
    );
    const h1Id = ctx.result.insertedParagraphs[0]!.newParaId;
    ctx.indices = buildIndices(pkg); // 重建索引

    executeInsertParagraph(
      { ...ctx, result: emptyApplyResult() },
      {
        type: 'insertParagraphAfter',
        anchorParaId: h1Id,
        spec: { text: '正文一', styleSource: 'BDYSPEC1' },
      },
      'after',
    );
    // 注意：需要从更新后的 result 取 id
    const body1Ctx = { ...ctx, result: emptyApplyResult(), indices: buildIndices(pkg) };
    executeInsertParagraph(
      body1Ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: h1Id,
        spec: { text: '正文一', styleSource: 'BDYSPEC1' },
      },
      'after',
    );
    const body1Id = body1Ctx.result.insertedParagraphs[0]!.newParaId;
    const body1Xml = paraXml(pkg, body1Id);
    // body 段：含宋体、sz24、both，不含 Heading1/黑体/36
    expect(body1Xml).toMatch(/宋体/);
    expect(body1Xml).toMatch(/24/);
    expect(body1Xml).not.toMatch(/Heading1/);
    expect(body1Xml).not.toMatch(/36/);

    const h2Ctx = { ...ctx, result: emptyApplyResult(), indices: buildIndices(pkg) };
    executeInsertParagraph(
      h2Ctx,
      {
        type: 'insertParagraphAfter',
        anchorParaId: body1Id,
        spec: { text: '1.1节', styleSource: 'H2SPEC01' },
      },
      'after',
    );
    const h2Id = h2Ctx.result.insertedParagraphs[0]!.newParaId;
    const h2Xml = paraXml(pkg, h2Id);
    // h2 段：含 Heading2、sz28，不含 Heading1/36/宋体
    expect(h2Xml).toMatch(/Heading2/);
    expect(h2Xml).toMatch(/28/);
    expect(h2Xml).not.toMatch(/Heading1/);
    expect(h2Xml).not.toMatch(/宋体/);

    const h1Xml = paraXml(pkg, h1Id);
    // h1 段：含 Heading1、sz36/黑体，不含宋体/24/Heading2
    expect(h1Xml).toMatch(/Heading1/);
    expect(h1Xml).toMatch(/36/);
    expect(h1Xml).not.toMatch(/宋体/);
    expect(h1Xml).not.toMatch(/Heading2/);
  });
});

// ─────────── executeDeleteParagraph ───────────

describe('executeDeleteParagraph', () => {
  function deleteOp(
    paraId: string,
    onAnchorConflict: 'error' | 'drop' = 'error',
  ): Extract<Op, { type: 'deleteParagraph' }> {
    return { type: 'deleteParagraph', paraId, onAnchorConflict };
  }

  test('基本:无锚点的段被成功删除', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>delete-me</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>keep</w:t></w:r></w:p>`,
    );
    executeDeleteParagraph(ctx, deleteOp('AAAAAAA1'));
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(false);
    expect(idx.paraIdx.has('AAAAAAA2')).toBe(true);
  });

  test('bookmarkStart + onAnchorConflict=error → ANCHOR_CONFLICT(affectedBookmarks)', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="myBookmark"/>
        <w:r><w:t>content</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    try {
      executeDeleteParagraph(ctx, deleteOp('AAAAAAA1', 'error'));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('ANCHOR_CONFLICT');
      const details = (e as WordError).details as { affectedBookmarks: string[] };
      expect(details.affectedBookmarks).toContain('myBookmark');
    }
  });

  test('commentRangeStart + onAnchorConflict=error → ANCHOR_CONFLICT(affectedComments)', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:commentRangeStart w:id="42"/>
        <w:r><w:t>content</w:t></w:r>
        <w:commentRangeEnd w:id="42"/>
      </w:p>`,
    );
    try {
      executeDeleteParagraph(ctx, deleteOp('AAAAAAA1', 'error'));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('ANCHOR_CONFLICT');
      const details = (e as WordError).details as { affectedComments: number[] };
      expect(details.affectedComments).toContain(42);
    }
  });

  test('onAnchorConflict=drop + 有书签 → 书签 + 段一并删除(#30)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="bm"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    executeDeleteParagraph(ctx, deleteOp('AAAAAAA1', 'drop'));
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(false);
    expect(idx.bookmarkIdx.has('bm')).toBe(false);
  });

  test('onAnchorConflict=drop + 无锚点 → 成功删除(走普通删除路径)', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>clean</w:t></w:r></w:p>`);
    executeDeleteParagraph(ctx, deleteOp('AAAAAAA1', 'drop'));
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(false);
  });

  test('PARA_NOT_FOUND:不存在的 paraId 抛错', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(() => executeDeleteParagraph(ctx, deleteOp('NOPE'))).toThrow(
      expect.objectContaining({ code: 'PARA_NOT_FOUND' }) as unknown as Error,
    );
  });

  // trackChanges=true 的覆盖在 trackChanges.test.ts(#31)

  test('段内 commentReference 也被检测为活锚点', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>x</w:t></w:r>
        <w:r><w:commentReference w:id="7"/></w:r>
      </w:p>`,
    );
    try {
      executeDeleteParagraph(ctx, deleteOp('AAAAAAA1', 'error'));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('ANCHOR_CONFLICT');
      const details = (e as WordError).details as { affectedComments: number[] };
      expect(details.affectedComments).toContain(7);
    }
  });
});

// ─────────── outlineLevel 测试（task-para-outline-level §3 矩阵）───────────

describe('executeInsertParagraph — outlineLevel', () => {
  function insertAfterOp(
    anchorParaId: string,
    spec: Parameters<typeof executeInsertParagraph>[1]['spec'],
  ): Extract<Op, { type: 'insertParagraphAfter' }> {
    return { type: 'insertParagraphAfter', anchorParaId, spec };
  }

  function insertBeforeOp(
    anchorParaId: string,
    spec: Parameters<typeof executeInsertParagraph>[1]['spec'],
  ): Extract<Op, { type: 'insertParagraphBefore' }> {
    return { type: 'insertParagraphBefore', anchorParaId, spec };
  }

  // §3 #1: outlineLevel:0 → pPr 有 <w:outlineLvl w:val="0"/>
  test('#1 outlineLevel:0 → 新段 pPr 含 <w:outlineLvl w:val="0"/>', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="OL000001"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      insertAfterOp('OL000001', { text: 'heading1', outlineLevel: 0 }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toContain('outlineLvl');
    expect(xml).toMatch(/w:outlineLvl[^>]*w:val="0"/);
  });

  // §3 #2: outlineLevel:2 → w:val="2"
  test('#2 outlineLevel:2 → w:val="2"', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="OL000002"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      insertAfterOp('OL000002', { text: 'heading3', outlineLevel: 2 }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toMatch(/w:outlineLvl[^>]*w:val="2"/);
  });

  // §3 #3: 省略 outlineLevel → 新段 pPr 无 outlineLvl（继承 anchor 的 pPr 也不含时）
  test('#3 省略 outlineLevel → pPr 无 outlineLvl', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="OL000003"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      insertAfterOp('OL000003', { text: 'normal', style: 'inherit' }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).not.toContain('outlineLvl');
  });

  // §3 #4: styleSource 克隆的 pPr 已有 outlineLvl + 显式 outlineLevel → 被覆盖（不重复）
  test('#4 styleSource 已有 outlineLvl + 显式 outlineLevel → 覆盖为显式值，无重复元素', () => {
    // source 段有 outlineLvl w:val="1"
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="OL000004">
         <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
         <w:r><w:t>source</w:t></w:r>
       </w:p>
       <w:p w14:paraId="OL000005"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    executeInsertParagraph(
      ctx,
      insertAfterOp('OL000005', { text: 'override', styleSource: 'OL000004', outlineLevel: 3 }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 覆盖为 3
    expect(xml).toMatch(/w:outlineLvl[^>]*w:val="3"/);
    // 只有一个 outlineLvl 元素
    const count = (xml.match(/outlineLvl/g) ?? []).length;
    // two occurrences expected: opening tag + closing (or self-closing counts once)
    // actually self-closing: just one occurrence in tag name
    expect(count).toBe(1);
  });

  // §3 #5: styleSource 段无 pPr + outlineLevel → 新建 pPr 含 outlineLvl
  test('#5 styleSource 无 pPr + outlineLevel → 新建 pPr 含 outlineLvl', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="OL000006"><w:r><w:t>noPPr source</w:t></w:r></w:p>
       <w:p w14:paraId="OL000007"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    executeInsertParagraph(
      ctx,
      insertAfterOp('OL000007', { text: 'new', styleSource: 'OL000006', outlineLevel: 1 }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toContain('pPr');
    expect(xml).toMatch(/w:outlineLvl[^>]*w:val="1"/);
  });

  // §3 #6: 越界 → INVALID_ARGUMENT
  test('#6 越界 outlineLevel=-1 → INVALID_ARGUMENT', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="OL000008"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    expect(() =>
      executeInsertParagraph(
        ctx,
        insertAfterOp('OL000008', { text: 'x', outlineLevel: -1 }),
        'after',
      ),
    ).toThrow(expect.objectContaining({ code: 'INVALID_ARGUMENT' }) as unknown as Error);
  });

  test('#6b 越界 outlineLevel=9 → INVALID_ARGUMENT', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="OL000009"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    expect(() =>
      executeInsertParagraph(
        ctx,
        insertAfterOp('OL000009', { text: 'x', outlineLevel: 9 }),
        'after',
      ),
    ).toThrow(expect.objectContaining({ code: 'INVALID_ARGUMENT' }) as unknown as Error);
  });

  // §3 #7: schema 顺序 — outlineLvl 在 numPr 之後（若有）
  test('#7 schema 顺序：outlineLvl 在 pPr 内 numPr 之后（若有 numPr）', () => {
    // anchor 段 pPr 含 numPr
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="OL000010">
         <w:pPr>
           <w:pStyle w:val="ListParagraph"/>
           <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
         </w:pPr>
         <w:r><w:t>list</w:t></w:r>
       </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      insertAfterOp('OL000010', { text: 'list with level', style: 'inherit', outlineLevel: 1 }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // outlineLvl 存在
    expect(xml).toContain('outlineLvl');
    // outlineLvl 出现位置在 numPr 之后
    const numPrPos = xml.indexOf('numPr');
    const outlineLvlPos = xml.indexOf('outlineLvl');
    expect(numPrPos).toBeGreaterThanOrEqual(0);
    expect(outlineLvlPos).toBeGreaterThan(numPrPos);
  });

  // §3 #8: insertParagraphBefore 对称支持
  test('#8 insertParagraphBefore 同样支持 outlineLevel', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="OL000011"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      insertBeforeOp('OL000011', { text: 'before heading', outlineLevel: 5 }),
      'before',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toMatch(/w:outlineLvl[^>]*w:val="5"/);
  });

  // §3 #9: 不回归 — 不带 outlineLevel 时行为不变（由于 describe 共享 makeDoc，
  // 本测试只验证没有意外注入 outlineLvl）
  test('#9 不回归：不传 outlineLevel 时无 outlineLvl 注入', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="OL000012">
         <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
         <w:r><w:t>anchor</w:t></w:r>
       </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      insertAfterOp('OL000012', { text: 'no outline', style: 'inherit' }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // pStyle 继承了，但没有 outlineLvl（因为 anchor pPr 没有）
    expect(xml).toContain('Heading1');
    expect(xml).not.toContain('outlineLvl');
  });

  // 位置修复验证：outlineLvl 在 rPr 之前（merger 真实场景：pPr 有 spacing + 段落标记 rPr，无 numPr）
  test('#10 位置修复：pPr 含 spacing+rPr（无 numPr）时 outlineLvl 在 rPr 之前', () => {
    // styleSource 段的 pPr 含 spacing + 段落标记 rPr（标题编号是文字，没有 numPr）
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="OL000013">
         <w:pPr>
           <w:pStyle w:val="Heading1"/>
           <w:spacing w:before="240" w:after="0"/>
           <w:rPr><w:b/></w:rPr>
         </w:pPr>
         <w:r><w:t>第一章</w:t></w:r>
       </w:p>
       <w:p w14:paraId="OL000014"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    executeInsertParagraph(
      ctx,
      insertAfterOp('OL000014', { text: '第二章', styleSource: 'OL000013', outlineLevel: 1 }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);

    expect(xml).toContain('outlineLvl');
    // 关键断言：outlineLvl 必须在 rPr 之前（Word 才能识别大纲级别）
    const outlineLvlPos = xml.indexOf('outlineLvl');
    const rPrPos = xml.indexOf('<w:rPr');
    expect(outlineLvlPos).toBeGreaterThanOrEqual(0);
    expect(rPrPos).toBeGreaterThanOrEqual(0);
    expect(outlineLvlPos).toBeLessThan(rPrPos);
  });

  // 建议 2：非整数越界
  test('#11 非整数 outlineLevel=1.5 → INVALID_ARGUMENT', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="OL000015"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    expect(() =>
      executeInsertParagraph(
        ctx,
        insertAfterOp('OL000015', { text: 'x', outlineLevel: 1.5 }),
        'after',
      ),
    ).toThrow(expect.objectContaining({ code: 'INVALID_ARGUMENT' }) as unknown as Error);
  });
});

// ─────────── pageBreakBefore 测试（task-core-pagebreak-before §3 矩阵）───────────

describe('executeInsertParagraph — pageBreakBefore', () => {
  function insertAfterOp(
    anchorParaId: string,
    spec: Parameters<typeof executeInsertParagraph>[1]['spec'],
  ): Extract<Op, { type: 'insertParagraphAfter' }> {
    return { type: 'insertParagraphAfter', anchorParaId, spec };
  }
  function insertBeforeOp(
    anchorParaId: string,
    spec: Parameters<typeof executeInsertParagraph>[1]['spec'],
  ): Extract<Op, { type: 'insertParagraphBefore' }> {
    return { type: 'insertParagraphBefore', anchorParaId, spec };
  }

  // §3 #1: insertParagraphAfter + pageBreakBefore:true → pPr 含 <w:pageBreakBefore/>
  test('#1 insertParagraphAfter + pageBreakBefore:true → 新段 pPr 含 <w:pageBreakBefore/>', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="PBB00001"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      insertAfterOp('PBB00001', { text: 'new-page', pageBreakBefore: true }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toContain('pageBreakBefore');
    // <w:pageBreakBefore/> 是自闭合元素（无 val），名字只出现一次
    const count = (xml.match(/pageBreakBefore/g) ?? []).length;
    expect(count).toBeGreaterThanOrEqual(1);
  });

  // §3 #2: pageBreakBefore 省略/false → 无 pageBreakBefore 元素
  test('#2 pageBreakBefore 省略 → 新段 pPr 无 pageBreakBefore', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="PBB00002"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(ctx, insertAfterOp('PBB00002', { text: 'no-break' }), 'after');

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).not.toContain('pageBreakBefore');
  });

  test('#2b pageBreakBefore:false → 新段 pPr 无 pageBreakBefore', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="PBB00003"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      insertAfterOp('PBB00003', { text: 'no-break', pageBreakBefore: false }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).not.toContain('pageBreakBefore');
  });

  // §3 #3: 与 outlineLevel 共存 → pPr 同时含两者，pageBreakBefore 在 outlineLvl 之前
  test('#3 与 outlineLevel 共存 → pPr 同时含 pageBreakBefore + outlineLvl，pageBreakBefore 在前', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="PBB00004"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      insertAfterOp('PBB00004', {
        text: 'heading-new-page',
        outlineLevel: 1,
        pageBreakBefore: true,
      }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toContain('pageBreakBefore');
    expect(xml).toContain('outlineLvl');
    // pageBreakBefore 必须在 outlineLvl 之前（CT_PPr schema 顺序）
    const pbbPos = xml.indexOf('pageBreakBefore');
    const outlineLvlPos = xml.indexOf('outlineLvl');
    expect(pbbPos).toBeGreaterThanOrEqual(0);
    expect(outlineLvlPos).toBeGreaterThanOrEqual(0);
    expect(pbbPos).toBeLessThan(outlineLvlPos);
  });

  // §3 #4: 与 styleSource 共存 → styleSource 克隆基础上加 pageBreakBefore，不破坏克隆内容
  test('#4 与 styleSource 共存 → styleSource 克隆 pPr 基础上加 pageBreakBefore', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="PBB00005">
         <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
         <w:r><w:t>source</w:t></w:r>
       </w:p>
       <w:p w14:paraId="PBB00006"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    executeInsertParagraph(
      ctx,
      insertAfterOp('PBB00006', {
        text: 'styled-new-page',
        styleSource: 'PBB00005',
        pageBreakBefore: true,
      }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 克隆来的 pStyle 仍在
    expect(xml).toContain('Heading1');
    // pageBreakBefore 加上了
    expect(xml).toContain('pageBreakBefore');
  });

  // §3 #6: insertParagraphBefore 同样生效
  test('#6 insertParagraphBefore 同样生效 → 新段 pPr 含 pageBreakBefore', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="PBB00007"><w:r><w:t>anchor</w:t></w:r></w:p>`);
    executeInsertParagraph(
      ctx,
      insertBeforeOp('PBB00007', { text: 'before-page', pageBreakBefore: true }),
      'before',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    expect(xml).toContain('pageBreakBefore');
  });

  // §3 #7: 不回归 — 存量无 pageBreakBefore 的段不受影响
  test('#7 不回归：不传 pageBreakBefore 时已有段落不变', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="PBB00008">
         <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
         <w:r><w:t>existing</w:t></w:r>
       </w:p>`,
    );
    executeInsertParagraph(
      ctx,
      insertAfterOp('PBB00008', { text: 'plain', style: 'inherit' }),
      'after',
    );

    const newId = ctx.result.insertedParagraphs[0]!.newParaId;
    const xml = paraXml(pkg, newId);
    // 继承了 pStyle，没有 pageBreakBefore 注入
    expect(xml).toContain('Heading2');
    expect(xml).not.toContain('pageBreakBefore');
  });
});
