/**
 * paragraph-ops executor 纯函数测试(#28)。
 *
 * 覆盖:
 *   - replaceParagraph:基本 / pPr 保留 / \n 转 `<w:br>` / \t 转 `<w:tab>` / 空
 *     text / 多个旧 run 全清 / marker 子(bookmarkStart 等)当前被清掉
 *     (ANCHOR_CONFLICT 保护交 #30)/ PARA_NOT_FOUND
 *   - replaceTextInParagraph:同 run 替换 / 跨 run 替换 / scope=all / 多次命中 /
 *     首 run rPr 继承 / TEXT_NOT_FOUND_IN_PARA / 空 search / 复杂 run(多 <w:t>)
 *     抛 UNSUPPORTED_OOXML_FEATURE
 *   - trackChanges=true 当前抛 UNSUPPORTED_OOXML_FEATURE(#31)
 *
 * 端到端测试(plan.apply → toZip 含新文本)见 [plan-apply-e2e.test.ts](../../__tests__/plan-apply-e2e.test.ts)。
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS, WPS_NS, VML_NS } from '../../namespaces.js';
import type { ApplyResult, Op } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeReplaceParagraph, executeReplaceTextInParagraph } from '../paragraph-ops.js';

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

/** 拿段落序列化后的 XML(便于断言结构)*/
function paraXml(pkg: Package, paraId: string): string {
  const idx = buildIndices(pkg);
  const el = idx.paraIdx.get(paraId.toUpperCase())!;
  return serializeToWellFormedString(el as never);
}

/** 取段落纯文本(所有 `<w:t>` 拼接)*/
function paraText(pkg: Package, paraId: string): string {
  const idx = buildIndices(pkg);
  const el = idx.paraIdx.get(paraId.toUpperCase())!;
  const ts = el.getElementsByTagNameNS(W_NS, 't');
  let s = '';
  for (let i = 0; i < ts.length; i++) s += ts[i]!.textContent ?? '';
  return s;
}

// ─────────── replaceParagraph ───────────

describe('executeReplaceParagraph', () => {
  test('基本:替换内容,保留 pPr', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r><w:t>old</w:t></w:r>
      </w:p>`,
    );
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'new' });
    expect(paraText(pkg, 'AAAAAAA1')).toBe('new');
    // pPr 仍在
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:pStyle\s+w:val="Heading1"/);
  });

  test('\\n → <w:br/>,\\t → <w:tab/>', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`);
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'a\nb\tc' });
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:br\s*\/>/);
    expect(xml).toMatch(/<w:tab\s*\/>/);
    // t 段顺序正确(a, b, c)
    expect(paraText(pkg, 'AAAAAAA1')).toBe('abc');
  });

  test('空 text:段仍在但内容为空(有 placeholder 空 <w:t>)', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`);
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: '' });
    expect(paraText(pkg, 'AAAAAAA1')).toBe('');
    // 段仍存在(没被整体移除)
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(true);
  });

  test('多个旧 run 被全清', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>a</w:t></w:r>
        <w:r><w:t>b</w:t></w:r>
        <w:r><w:t>c</w:t></w:r>
      </w:p>`,
    );
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'x' });
    expect(paraText(pkg, 'AAAAAAA1')).toBe('x');
  });

  test('活锚点保护(#30):段内含 bookmark 时抛 ANCHOR_CONFLICT', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="b1"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    try {
      executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'new' });
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('ANCHOR_CONFLICT');
      expect((e as WordError).details?.affectedBookmarks).toEqual(['b1']);
    }
  });

  test('PARA_NOT_FOUND:不存在的 paraId 抛错', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(() =>
      executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'NONE', text: 'x' }),
    ).toThrow(expect.objectContaining({ code: 'PARA_NOT_FOUND' }) as unknown as Error);
  });

  // trackChanges=true 的覆盖由 trackChanges.test.ts 统一做(#31 已实现包装层)
});

// ─────────── replaceTextInParagraph ───────────

describe('executeReplaceTextInParagraph', () => {
  function replaceOp(
    paraId: string,
    search: string,
    replacement: string,
    scope: 'first' | 'all' = 'first',
  ): Extract<Op, { type: 'replaceTextInParagraph' }> {
    return { type: 'replaceTextInParagraph', paraId, search, replacement, scope };
  }

  test('同 run:命中一次,替换正确', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello world</w:t></w:r></w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'world', 'everyone'));
    expect(paraText(pkg, 'AAAAAAA1')).toBe('hello everyone');
  });

  test('跨 run:匹配跨 2 个 run,被正确拼替换', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>hello wo</w:t></w:r>
        <w:r><w:t>rld</w:t></w:r>
      </w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'world', 'everyone'));
    expect(paraText(pkg, 'AAAAAAA1')).toBe('hello everyone');
  });

  test('scope=all:多处命中全替换(倒序处理不影响结果)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>foo bar foo baz foo</w:t></w:r></w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'foo', 'X', 'all'));
    expect(paraText(pkg, 'AAAAAAA1')).toBe('X bar X baz X');
  });

  test('scope=first:只替换首次', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>foo foo foo</w:t></w:r></w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'foo', 'X', 'first'));
    expect(paraText(pkg, 'AAAAAAA1')).toBe('X foo foo');
  });

  test('rPr 继承:首 match run 的 <w:rPr> 深拷贝到新 run', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'bold', 'B'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:rPr>\s*<w:b\s*\/>\s*<\/w:rPr>/);
    expect(paraText(pkg, 'AAAAAAA1')).toBe('B');
  });

  test('replacement 含 \\n / \\t:生成 <w:br/> / <w:tab/>', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>XXX</w:t></w:r></w:p>`);
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'XXX', 'a\nb'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:br\s*\/>/);
    expect(paraText(pkg, 'AAAAAAA1')).toBe('ab');
  });

  test('TEXT_NOT_FOUND_IN_PARA:不命中抛错', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    try {
      executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'notfound', 'x'));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('TEXT_NOT_FOUND_IN_PARA');
    }
  });

  test('空 search 抛 TEXT_NOT_FOUND_IN_PARA(防 infinite loop)', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    try {
      executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', '', 'x'));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('TEXT_NOT_FOUND_IN_PARA');
    }
  });

  test('复杂 run(多 <w:t>)抛 UNSUPPORTED_OOXML_FEATURE', () => {
    // splitRunAt 仅支持单 <w:t> run;跨 run 匹配时要切复杂 run → 抛错
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>hel</w:t><w:t>lo</w:t></w:r>
        <w:r><w:t> world</w:t></w:r>
      </w:p>`,
    );
    try {
      executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'lo wo', 'X'));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
    }
  });

  test('复杂 run 内单个 w:t 命中：原位替换并保留相邻 tab 与其他文本节点', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>{机关}</w:t><w:tab/><w:t>{日期}印发</w:t></w:r>
      </w:p>`,
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', '{机关}', '办公室'));
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', '{日期}', '2026年5月11日'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toContain('<w:tab/>');
    expect(xml).toContain('办公室');
    expect(xml).toContain('2026年5月11日印发');
    expect(xml).not.toContain('{机关}');
    expect(xml).not.toContain('{日期}');
  });

  // trackChanges=true 的覆盖在 trackChanges.test.ts(#31)

  test('PARA_NOT_FOUND', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    try {
      executeReplaceTextInParagraph(ctx, replaceOp('NONE', 'h', 'x'));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('PARA_NOT_FOUND');
    }
  });
});

// ─────────── keepInlineGraphics ───────────

/**
 * 验收条目 1-5：keepInlineGraphics 选项。
 *
 * 覆盖:
 *   1. wps:wsp 图形 run 保留，新文字 run 正确插入（替代旧文字 run）
 *   2. 默认不带 keepInlineGraphics → 图形 run 被删（现有行为，回归）
 *   3. v:line（VML w:pict）版记横线段保留
 *   4. trackChanges=true 变体：图形 run 不进 w:del，保留原位
 *   5. 纯文字段 + keepInlineGraphics:true → 行为与不带选项一致
 */
describe('executeReplaceParagraph + keepInlineGraphics', () => {
  /** 带 w:drawing/wps:wsp 的图形 run XML 片段（模拟公文红头 shape） */
  const _DRAWING_RUN = `<w:r xmlns:w="${W_NS}" xmlns:wps="${WPS_NS}"><w:drawing><wps:wsp><wps:spPr/></wps:wsp></w:drawing></w:r>`;

  /** 带 w:pict/v:line 的 VML 图形 run XML 片段（模拟版记横线） */
  const _PICT_RUN = `<w:r xmlns:w="${W_NS}" xmlns:v="${VML_NS}"><w:pict><v:line style="position:absolute"/></w:pict></w:r>`;

  function makeDocNS(bodyInner: string): {
    pkg: Package;
    indices: ReturnType<typeof buildIndices>;
    ctx: ExecutorContext;
  } {
    const pkg = createPackage();
    // 注入所需命名空间声明
    const xml = `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:wps="${WPS_NS}" xmlns:v="${VML_NS}"><w:body>${bodyInner}</w:body></w:document>`;
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        xmlDom: parseXmlDocument(xml) as unknown as Document,
        originalBytes: new TextEncoder().encode(xml),
      }),
    );
    const indices = buildIndices(pkg);
    const ctx: ExecutorContext = {
      pkg,
      indices,
      options: { trackChanges: false, author: '', date: new Date('2026-06-22T00:00:00Z') },
      result: {
        insertedParagraphs: [],
        insertedBodyStartParaIds: [],
        createdComments: [],
        createdBookmarks: [],
        acceptedRevisions: 0,
        rejectedRevisions: 0,
        replacedTableCells: 0,
      },
      opIndex: 0,
      docId: 'test-doc-graphics',
    };
    return { pkg, indices, ctx };
  }

  function makeDocTC(bodyInner: string): {
    pkg: Package;
    ctx: ExecutorContext;
  } {
    const pkg = createPackage();
    const xml = `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:wps="${WPS_NS}" xmlns:v="${VML_NS}"><w:body>${bodyInner}</w:body></w:document>`;
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        xmlDom: parseXmlDocument(xml) as unknown as Document,
        originalBytes: new TextEncoder().encode(xml),
      }),
    );
    const indices = buildIndices(pkg);
    const ctx: ExecutorContext = {
      pkg,
      indices,
      options: { trackChanges: true, author: 'alice', date: new Date('2026-06-22T00:00:00Z') },
      result: {
        insertedParagraphs: [],
        insertedBodyStartParaIds: [],
        createdComments: [],
        createdBookmarks: [],
        acceptedRevisions: 0,
        rejectedRevisions: 0,
        replacedTableCells: 0,
      },
      opIndex: 0,
      docId: 'test-doc-tc',
    };
    return { pkg, ctx };
  }

  // ── 验收条目 1 ──
  test('条目1: wps:wsp 图形 run 保留，新文字 run 插入正确位置', () => {
    const { pkg, ctx } = makeDocNS(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>红头文字</w:t></w:r>
        <w:r><w:drawing><wps:wsp><wps:spPr/></wps:wsp></w:drawing></w:r>
      </w:p>`,
    );
    executeReplaceParagraph(ctx, {
      type: 'replaceParagraph',
      paraId: 'AAAAAAA1',
      text: '{发文机关标志}',
      onAnchorConflict: 'error',
      keepInlineGraphics: true,
    });
    const xml = paraXml(pkg, 'AAAAAAA1');
    // 新文字 run 应存在
    expect(paraText(pkg, 'AAAAAAA1')).toBe('{发文机关标志}');
    // 图形 run 未丢（wps:wsp 仍在）
    const idx = buildIndices(pkg);
    const el = idx.paraIdx.get('AAAAAAA1')!;
    expect(el.getElementsByTagNameNS(WPS_NS, 'wsp').length).toBeGreaterThan(0);
    // 文字 run 在图形 run 之前（文字先于线条）
    const runs = Array.from(el.childNodes).filter(
      (c) => c.nodeType === 1 && (c as Element).localName === 'r',
    ) as Element[];
    expect(runs.length).toBe(2); // 1 文字 run + 1 图形 run
    const firstRunHasText = runs[0]!.getElementsByTagNameNS(W_NS, 't').length > 0;
    const secondRunHasDrawing = runs[1]!.getElementsByTagNameNS(WPS_NS, 'wsp').length > 0;
    expect(firstRunHasText).toBe(true);
    expect(secondRunHasDrawing).toBe(true);
    void xml;
  });

  // ── 验收条目 2 ──
  test('条目2: 默认不带 keepInlineGraphics → 图形 run 被删（现有行为回归）', () => {
    const { pkg, ctx } = makeDocNS(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>红头文字</w:t></w:r>
        <w:r><w:drawing><wps:wsp><wps:spPr/></wps:wsp></w:drawing></w:r>
      </w:p>`,
    );
    executeReplaceParagraph(ctx, {
      type: 'replaceParagraph',
      paraId: 'AAAAAAA1',
      text: '{发文机关标志}',
      onAnchorConflict: 'error',
      // 不传 keepInlineGraphics（默认 false）
    });
    const idx = buildIndices(pkg);
    const el = idx.paraIdx.get('AAAAAAA1')!;
    // 图形 run 应被删除
    expect(el.getElementsByTagNameNS(WPS_NS, 'wsp').length).toBe(0);
    // 文字仍然替换成功
    expect(paraText(pkg, 'AAAAAAA1')).toBe('{发文机关标志}');
  });

  // ── 验收条目 3 ──
  test('条目3: v:line（VML w:pict）版记横线段同样保留', () => {
    const { pkg, ctx } = makeDocNS(
      `<w:p w14:paraId="AAAAAAA2">
        <w:r><w:t>印发机关</w:t></w:r>
        <w:r><w:pict><v:line style="position:absolute"/></w:pict></w:r>
      </w:p>`,
    );
    executeReplaceParagraph(ctx, {
      type: 'replaceParagraph',
      paraId: 'AAAAAAA2',
      text: '{印发机关}',
      onAnchorConflict: 'error',
      keepInlineGraphics: true,
    });
    const idx = buildIndices(pkg);
    const el = idx.paraIdx.get('AAAAAAA2')!;
    // 文字替换成功
    expect(paraText(pkg, 'AAAAAAA2')).toBe('{印发机关}');
    // v:line 横线仍在
    expect(el.getElementsByTagNameNS(VML_NS, 'line').length).toBeGreaterThan(0);
  });

  // ── 验收条目 4 ──
  test('条目4: trackChanges=true 变体 — 图形 run 不进 w:del，保留原位', () => {
    const { pkg, ctx } = makeDocTC(
      `<w:p w14:paraId="AAAAAAA3">
        <w:r><w:t>红头文字</w:t></w:r>
        <w:r><w:drawing><wps:wsp><wps:spPr/></wps:wsp></w:drawing></w:r>
      </w:p>`,
    );
    executeReplaceParagraph(ctx, {
      type: 'replaceParagraph',
      paraId: 'AAAAAAA3',
      text: '{发文机关标志}',
      onAnchorConflict: 'error',
      keepInlineGraphics: true,
    });
    const idx = buildIndices(pkg);
    const el = idx.paraIdx.get('AAAAAAA3')!;
    // 图形 run 仍在段落中（不在 del 内）
    expect(el.getElementsByTagNameNS(WPS_NS, 'wsp').length).toBeGreaterThan(0);
    // w:del 存在（旧文字被包进 del）
    expect(el.getElementsByTagNameNS(W_NS, 'del').length).toBeGreaterThan(0);
    // w:ins 存在（新文字被包进 ins）
    expect(el.getElementsByTagNameNS(W_NS, 'ins').length).toBeGreaterThan(0);
    // 图形 run 不在 del 内（del 元素不含 wps:wsp）
    const delEls = el.getElementsByTagNameNS(W_NS, 'del');
    for (let i = 0; i < delEls.length; i++) {
      expect(delEls[i]!.getElementsByTagNameNS(WPS_NS, 'wsp').length).toBe(0);
    }
    // ins 内含新文字
    const insEls = el.getElementsByTagNameNS(W_NS, 'ins');
    let insText = '';
    for (let i = 0; i < insEls.length; i++) {
      const ts = insEls[i]!.getElementsByTagNameNS(W_NS, 't');
      for (let j = 0; j < ts.length; j++) insText += ts[j]!.textContent ?? '';
    }
    expect(insText).toBe('{发文机关标志}');
  });

  // ── 验收条目 5 ──
  test('条目5: 纯文字段 + keepInlineGraphics:true → 行为与不带选项一致', () => {
    const { pkg: pkg1, ctx: ctx1 } = makeDocNS(
      `<w:p w14:paraId="AAAAAAA4"><w:r><w:t>old text</w:t></w:r></w:p>`,
    );
    const { pkg: pkg2, ctx: ctx2 } = makeDocNS(
      `<w:p w14:paraId="AAAAAAA4"><w:r><w:t>old text</w:t></w:r></w:p>`,
    );
    // 一个带 keepInlineGraphics，一个不带
    executeReplaceParagraph(ctx1, {
      type: 'replaceParagraph',
      paraId: 'AAAAAAA4',
      text: 'new text',
      onAnchorConflict: 'error',
      keepInlineGraphics: true,
    });
    executeReplaceParagraph(ctx2, {
      type: 'replaceParagraph',
      paraId: 'AAAAAAA4',
      text: 'new text',
      onAnchorConflict: 'error',
    });
    // 文字都替换成功，结果一致
    expect(paraText(pkg1, 'AAAAAAA4')).toBe('new text');
    expect(paraText(pkg2, 'AAAAAAA4')).toBe('new text');
    // 段落中没有多余的图形残留
    const idx1 = buildIndices(pkg1);
    const el1 = idx1.paraIdx.get('AAAAAAA4')!;
    const idx2 = buildIndices(pkg2);
    const el2 = idx2.paraIdx.get('AAAAAAA4')!;
    expect(el1.getElementsByTagNameNS(WPS_NS, 'wsp').length).toBe(0);
    expect(el2.getElementsByTagNameNS(WPS_NS, 'wsp').length).toBe(0);
  });

  // ── issue 验收标准 1：锚点安全整段替换（task-20260622-replacepara-anchor-safe）──
  test('锚点安全(anchor-safe): 段含v:line图形+文字+commentRange+bookmark，keepInlineGraphics:true 替换后文字变、图形在、锚点 marker 不孤儿', () => {
    // 构造段：pict/v:line 图形 run + 文字 run + commentRangeStart/End + bookmarkStart/End
    // 这是版记印发行的典型布局：段上同时压着浮动版记横线 + COCO 批注 + _GoBack 书签
    const { pkg, ctx } = makeDocNS(
      `<w:p w14:paraId="ANCHOR01">
        <w:r><w:pict><v:line style="position:absolute"/></w:pict></w:r>
        <w:r><w:t>江临市人民政府办公室　2026年5月16日印发</w:t></w:r>
        <w:commentRangeStart w:id="29"/>
        <w:commentRangeEnd w:id="29"/>
        <w:r><w:commentReference w:id="29"/></w:r>
        <w:bookmarkStart w:id="0" w:name="_GoBack"/>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );

    // 执行：keepInlineGraphics:true + onAnchorConflict:'drop'（真机调用方式）
    // 修复前：此调用会被 onAnchorConflict:'drop' 静默跳过，文字纹丝不动
    // 修复后：顺利执行，文字被替换
    executeReplaceParagraph(ctx, {
      type: 'replaceParagraph',
      paraId: 'ANCHOR01',
      text: '{印发机关}　{印发日期}印发',
      onAnchorConflict: 'drop',
      keepInlineGraphics: true,
    });

    const idx = buildIndices(pkg);
    const el = idx.paraIdx.get('ANCHOR01')!;

    // ① 段文字 == 新文字
    expect(paraText(pkg, 'ANCHOR01')).toBe('{印发机关}　{印发日期}印发');

    // ② 图形 run（v:line）仍在
    expect(el.getElementsByTagNameNS(VML_NS, 'line').length).toBeGreaterThan(0);

    // ③ comment range marker 仍在（不孤儿）
    expect(el.getElementsByTagNameNS(W_NS, 'commentRangeStart').length).toBe(1);
    expect(el.getElementsByTagNameNS(W_NS, 'commentRangeEnd').length).toBe(1);
    expect(el.getElementsByTagNameNS(W_NS, 'commentReference').length).toBe(1);

    // ④ bookmark marker 仍在（_GoBack 不孤儿）
    expect(el.getElementsByTagNameNS(W_NS, 'bookmarkStart').length).toBe(1);
    expect(el.getElementsByTagNameNS(W_NS, 'bookmarkEnd').length).toBe(1);

    // ⑤ 再次 buildIndices 重建（模拟重开），paragraphs.list 不崩
    const idx2 = buildIndices(pkg);
    const el2 = idx2.paraIdx.get('ANCHOR01');
    expect(el2).toBeDefined();
    // 重建后文字仍正确
    const ts2 = el2!.getElementsByTagNameNS(W_NS, 't');
    let text2 = '';
    for (let i = 0; i < ts2.length; i++) text2 += ts2[i]!.textContent ?? '';
    expect(text2).toBe('{印发机关}　{印发日期}印发');
  });

  // ── 回归：keepInlineGraphics:true + 有锚点 + onAnchorConflict:'error' 应当正常替换（不再抛 ANCHOR_CONFLICT）──
  test('回归: keepInlineGraphics:true + 有锚点 + onAnchorConflict:error → 正常替换（不抛锚点冲突）', () => {
    const { pkg, ctx } = makeDocNS(
      `<w:p w14:paraId="ANCHOR02">
        <w:bookmarkStart w:id="1" w:name="_GoBack"/>
        <w:r><w:t>原文字</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
        <w:r><w:pict><v:line style="position:absolute"/></w:pict></w:r>
      </w:p>`,
    );
    // keepInlineGraphics:true 时，即使有锚点，也不应抛 ANCHOR_CONFLICT
    expect(() =>
      executeReplaceParagraph(ctx, {
        type: 'replaceParagraph',
        paraId: 'ANCHOR02',
        text: '新文字',
        onAnchorConflict: 'error',
        keepInlineGraphics: true,
      }),
    ).not.toThrow();

    // 文字正确替换
    expect(paraText(pkg, 'ANCHOR02')).toBe('新文字');

    // 锚点 marker 保留
    const idx = buildIndices(pkg);
    const el = idx.paraIdx.get('ANCHOR02')!;
    expect(el.getElementsByTagNameNS(W_NS, 'bookmarkStart').length).toBe(1);
    expect(el.getElementsByTagNameNS(W_NS, 'bookmarkEnd').length).toBe(1);

    // 图形 run 保留
    expect(el.getElementsByTagNameNS(VML_NS, 'line').length).toBeGreaterThan(0);
  });

  // ── 回归：不带 keepInlineGraphics 的锚点冲突语义零回归 ──
  test('回归: 不带 keepInlineGraphics + 有锚点 → 仍按 onAnchorConflict 原语义处理', () => {
    // onAnchorConflict:'error' 应仍抛
    const { ctx: ctxErr } = makeDocNS(
      `<w:p w14:paraId="ANCHOR03">
        <w:bookmarkStart w:id="1" w:name="bk1"/>
        <w:r><w:t>text</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    expect(() =>
      executeReplaceParagraph(ctxErr, {
        type: 'replaceParagraph',
        paraId: 'ANCHOR03',
        text: 'new',
        onAnchorConflict: 'error',
        // 不传 keepInlineGraphics
      }),
    ).toThrow(expect.objectContaining({ code: 'ANCHOR_CONFLICT' }) as unknown as Error);

    // onAnchorConflict:'drop' 应仍静默跳过（文字不变）
    const { pkg: pkgDrop, ctx: ctxDrop } = makeDocNS(
      `<w:p w14:paraId="ANCHOR04">
        <w:bookmarkStart w:id="2" w:name="bk2"/>
        <w:r><w:t>原文不变</w:t></w:r>
        <w:bookmarkEnd w:id="2"/>
      </w:p>`,
    );
    executeReplaceParagraph(ctxDrop, {
      type: 'replaceParagraph',
      paraId: 'ANCHOR04',
      text: '新文字',
      onAnchorConflict: 'drop',
      // 不传 keepInlineGraphics
    });
    expect(paraText(pkgDrop, 'ANCHOR04')).toBe('原文不变');
  });
});
