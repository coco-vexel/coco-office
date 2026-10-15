/**
 * ANCHOR_CONFLICT 检测测试 —— #30。
 *
 * 覆盖:
 *   - replaceParagraph:无锚点 / 有 bookmark / 有 comment / 两者同时有
 *   - replaceTextInParagraph:匹配区间外锚点(不抛) / 区间内 bookmark / 区间内 comment /
 *     跨 run 匹配 + 中间穿 bookmark
 *   - deleteParagraph:无锚点 'error' / 有锚点 'error' / 有 bookmark 'drop' /
 *     有 comment 'drop' / bookmark + comment 'drop'
 *   - details 字段格式:paraId / affectedBookmarks / affectedComments / opType
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult, Op } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeReplaceParagraph, executeReplaceTextInParagraph } from '../paragraph-ops.js';
import { executeDeleteParagraph } from '../paragraph-structure-ops.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const COMMENTS_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

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

function makeDoc(
  bodyInner: string,
  commentsXmlOrOpts?: string | { commentsXml?: string; trackChanges?: boolean },
  legacyTrackChanges?: boolean,
): { pkg: Package; ctx: ExecutorContext } {
  // 兼容旧签名 makeDoc(body, commentsXml) 与新签名 makeDoc(body, { commentsXml, trackChanges })
  const opts =
    typeof commentsXmlOrOpts === 'string'
      ? { commentsXml: commentsXmlOrOpts, trackChanges: legacyTrackChanges ?? false }
      : {
          commentsXml: commentsXmlOrOpts?.commentsXml,
          trackChanges: commentsXmlOrOpts?.trackChanges ?? false,
        };
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`,
  );
  if (opts.commentsXml) {
    putXml(pkg, 'word/comments.xml', opts.commentsXml, COMMENTS_CTYPE);
  }
  const indices = buildIndices(pkg);
  const ctx: ExecutorContext = {
    pkg,
    indices,
    options: {
      trackChanges: opts.trackChanges,
      author: 'tester',
      date: new Date('2026-04-22T00:00:00Z'),
    },
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
  const el = idx.paraIdx.get(paraId.toUpperCase());
  if (!el) return '';
  return serializeToWellFormedString(el as never);
}

// ─────────── replaceParagraph ───────────

describe('executeReplaceParagraph — ANCHOR_CONFLICT', () => {
  test('段内无锚点 → 成功替换', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>`);
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'new' });
    // 段内仍在且文本被替换
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toContain('new');
  });

  test('段内有 bookmarkStart → 抛 ANCHOR_CONFLICT,affectedBookmarks 含 name', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="myBm"/>
        <w:r><w:t>text</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    try {
      executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'x' });
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      const we = e as WordError;
      expect(we.code).toBe('ANCHOR_CONFLICT');
      expect(we.details.paraId).toBe('AAAAAAA1');
      expect(we.details.affectedBookmarks).toContain('myBm');
      expect(we.details.opType).toBe('replaceParagraph');
    }
  });

  test('段内有 commentRangeStart → 抛 ANCHOR_CONFLICT,affectedComments 含 id', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:commentRangeStart w:id="5"/>
        <w:r><w:t>text</w:t></w:r>
        <w:commentRangeEnd w:id="5"/>
        <w:r><w:commentReference w:id="5"/></w:r>
      </w:p>`,
    );
    try {
      executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'x' });
      expect.unreachable();
    } catch (e) {
      const we = e as WordError;
      expect(we.code).toBe('ANCHOR_CONFLICT');
      expect(we.details.affectedComments).toContain(5);
      expect((we.details.affectedBookmarks as string[]).length).toBe(0);
    }
  });

  test('段内 bookmark + comment 都有 → details 两数组都完整', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="bk1"/>
        <w:commentRangeStart w:id="2"/>
        <w:r><w:t>text</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
        <w:commentRangeEnd w:id="2"/>
        <w:r><w:commentReference w:id="2"/></w:r>
      </w:p>`,
    );
    try {
      executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'x' });
      expect.unreachable();
    } catch (e) {
      const we = e as WordError;
      expect(we.code).toBe('ANCHOR_CONFLICT');
      expect(we.details.affectedBookmarks).toContain('bk1');
      expect(we.details.affectedComments).toContain(2);
    }
  });
});

// ─────────── replaceTextInParagraph ───────────

describe('executeReplaceTextInParagraph — ANCHOR_CONFLICT', () => {
  function replaceOp(
    paraId: string,
    search: string,
    replacement: string,
    scope: 'first' | 'all' = 'first',
  ): Extract<Op, { type: 'replaceTextInParagraph' }> {
    return { type: 'replaceTextInParagraph', paraId, search, replacement, scope };
  }

  test('匹配区间外有 bookmark(段首) → 不抛,替换成功', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="outsideBk"/>
        <w:bookmarkEnd w:id="1"/>
        <w:r><w:t>hello world</w:t></w:r>
      </w:p>`,
    );
    // bookmark 在 run 之前 — 不在匹配区间内
    // replaceTextInParagraph 只检测 run 区间内的 sibling 锚点
    // 此段内锚点在文字 run 之外 → 不抛
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'world', 'everyone'));
    const ts = buildIndices(pkg).paraIdx.get('AAAAAAA1')!.getElementsByTagNameNS(W_NS, 't');
    let text = '';
    for (let i = 0; i < ts.length; i++) text += ts[i]!.textContent ?? '';
    expect(text).toContain('everyone');
  });

  test('matchRun 区间内有 bookmarkStart → 抛 ANCHOR_CONFLICT', () => {
    // 布局:run("hello ") | bookmarkStart | run("world")
    // 搜索 "hello world" → 跨两 run → 切出 startRun=run("hello "), endRun=run("world")
    // 两者之间夹了 bookmarkStart → 抛
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>hello </w:t></w:r>
        <w:bookmarkStart w:id="1" w:name="innerBk"/>
        <w:r><w:t>world</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    try {
      executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'hello world', 'X'));
      expect.unreachable();
    } catch (e) {
      const we = e as WordError;
      expect(we.code).toBe('ANCHOR_CONFLICT');
      expect(we.details.affectedBookmarks).toContain('innerBk');
      expect(we.details.opType).toBe('replaceTextInParagraph');
    }
  });

  test('匹配区间内有 commentRangeStart → 抛 ANCHOR_CONFLICT', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>abc </w:t></w:r>
        <w:commentRangeStart w:id="3"/>
        <w:r><w:t>def</w:t></w:r>
      </w:p>`,
    );
    try {
      executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'abc def', 'X'));
      expect.unreachable();
    } catch (e) {
      const we = e as WordError;
      expect(we.code).toBe('ANCHOR_CONFLICT');
      expect(we.details.affectedComments).toContain(3);
    }
  });

  test('同 run 内匹配:区间内无锚点 → 成功(同 run 不穿 sibling)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="sideBk"/>
        <w:r><w:t>hello world</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    // 匹配 "world" 在同 run 内 —— 切出来的 startRunForMatch === endRunForMatch,
    // collectRangeAnchors 只扫从 startRun 到 endRun 的 sibling 链,同一节点之间无锚点
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'world', 'everyone'));
    const ts = buildIndices(pkg).paraIdx.get('AAAAAAA1')!.getElementsByTagNameNS(W_NS, 't');
    let text = '';
    for (let i = 0; i < ts.length; i++) text += ts[i]!.textContent ?? '';
    expect(text).toContain('everyone');
  });
});

// ─────────── deleteParagraph ───────────

describe('executeDeleteParagraph — ANCHOR_CONFLICT', () => {
  test('无锚点 + onAnchorConflict=error → 成功删段', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>bye</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>stay</w:t></w:r></w:p>`,
    );
    executeDeleteParagraph(ctx, {
      type: 'deleteParagraph',
      paraId: 'AAAAAAA1',
      onAnchorConflict: 'error',
    });
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(false);
    expect(idx.paraIdx.has('AAAAAAA2')).toBe(true);
  });

  test('有 bookmark + onAnchorConflict=error → 抛 ANCHOR_CONFLICT', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="errBk"/>
        <w:r><w:t>text</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
    );
    try {
      executeDeleteParagraph(ctx, {
        type: 'deleteParagraph',
        paraId: 'AAAAAAA1',
        onAnchorConflict: 'error',
      });
      expect.unreachable();
    } catch (e) {
      const we = e as WordError;
      expect(we.code).toBe('ANCHOR_CONFLICT');
      expect(we.details.affectedBookmarks).toContain('errBk');
      expect(we.details.paraId).toBe('AAAAAAA1');
      expect(we.details.opType).toBe('deleteParagraph');
    }
  });

  test('有 bookmark + onAnchorConflict=drop → 成功删段 + bookmark DOM 被移除', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="dropBk"/>
        <w:r><w:t>text</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>
      <w:p w14:paraId="AAAAAAA2"><w:r><w:t>keep</w:t></w:r></w:p>`,
    );
    executeDeleteParagraph(ctx, {
      type: 'deleteParagraph',
      paraId: 'AAAAAAA1',
      onAnchorConflict: 'drop',
    });
    const idx = buildIndices(pkg);
    // 段被删
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(false);
    // 书签不再存在于文档
    expect(idx.bookmarkIdx.has('dropBk')).toBe(false);
  });

  test('有 comment + onAnchorConflict=drop → 成功删段 + comment 三锚点被移除', () => {
    const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:comments xmlns:w="${W_NS}">
        <w:comment w:id="7" w:author="A" w:date="2026-04-22T00:00:00Z">
          <w:p><w:r><w:t>hello</w:t></w:r></w:p>
        </w:comment>
      </w:comments>`;
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:commentRangeStart w:id="7"/>
        <w:r><w:t>commented text</w:t></w:r>
        <w:commentRangeEnd w:id="7"/>
        <w:r><w:commentReference w:id="7"/></w:r>
      </w:p>
      <w:p w14:paraId="AAAAAAA2"><w:r><w:t>other</w:t></w:r></w:p>`,
      commentsXml,
    );
    executeDeleteParagraph(ctx, {
      type: 'deleteParagraph',
      paraId: 'AAAAAAA1',
      onAnchorConflict: 'drop',
    });
    const idx = buildIndices(pkg);
    // 段被删
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(false);
    // comment 不再存在
    expect(idx.commentIdx.has(7)).toBe(false);
  });

  test('bookmark + comment 同段 + drop → 全部清理', () => {
    const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:comments xmlns:w="${W_NS}">
        <w:comment w:id="9" w:author="A" w:date="2026-04-22T00:00:00Z">
          <w:p><w:r><w:t>note</w:t></w:r></w:p>
        </w:comment>
      </w:comments>`;
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="2" w:name="comboBook"/>
        <w:commentRangeStart w:id="9"/>
        <w:r><w:t>mixed</w:t></w:r>
        <w:bookmarkEnd w:id="2"/>
        <w:commentRangeEnd w:id="9"/>
        <w:r><w:commentReference w:id="9"/></w:r>
      </w:p>`,
      commentsXml,
    );
    executeDeleteParagraph(ctx, {
      type: 'deleteParagraph',
      paraId: 'AAAAAAA1',
      onAnchorConflict: 'drop',
    });
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(false);
    expect(idx.bookmarkIdx.has('comboBook')).toBe(false);
    expect(idx.commentIdx.has(9)).toBe(false);
  });
});

// ─────────── trackChanges=true:锚点不再拦截,跟随被删内容包进 <w:del> ───────────

describe('executeReplaceTextInParagraph + trackChanges=true — 锚点跟随包入 <w:del>', () => {
  function replaceOp(
    paraId: string,
    search: string,
    replacement: string,
  ): Extract<Op, { type: 'replaceTextInParagraph' }> {
    return { type: 'replaceTextInParagraph', paraId, search, replacement, scope: 'first' };
  }

  test('跨 commentRangeStart → 不抛,<w:del> 直系子节点含 <w:commentRangeStart>', () => {
    // 布局:run("abc ") | commentRangeStart | run("def")
    // 搜索 "abc def" → toRemove = [run, commentRangeStart, run] → 全部包入 <w:del>
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>abc </w:t></w:r>
        <w:commentRangeStart w:id="3"/>
        <w:r><w:t>def</w:t></w:r>
      </w:p>`,
      { trackChanges: true },
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'abc def', 'X'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    // <w:del> 包含 commentRangeStart
    expect(xml).toMatch(/<w:del[^>]*>[\s\S]*<w:commentRangeStart[^>]*w:id="3"[\s\S]*<\/w:del>/);
    // 旧文本转成 <w:delText>
    expect(xml).toContain('<w:delText>abc </w:delText>');
    expect(xml).toContain('<w:delText>def</w:delText>');
    // 新文本 X 在 <w:ins> 里
    expect(xml).toMatch(/<w:ins[^>]*>[\s\S]*<w:t[^>]*>X<\/w:t>[\s\S]*<\/w:ins>/);
  });

  test('跨 bookmarkStart → 不抛,<w:del> 含 <w:bookmarkStart>', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>hello </w:t></w:r>
        <w:bookmarkStart w:id="1" w:name="innerBk"/>
        <w:r><w:t>world</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
      { trackChanges: true },
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'hello world', 'Y'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:del[^>]*>[\s\S]*<w:bookmarkStart[^>]*w:name="innerBk"[\s\S]*<\/w:del>/);
    expect(xml).toContain('<w:delText>hello </w:delText>');
    expect(xml).toContain('<w:delText>world</w:delText>');
    expect(xml).toMatch(/<w:ins[^>]*>[\s\S]*<w:t[^>]*>Y<\/w:t>[\s\S]*<\/w:ins>/);
  });

  test('匹配区间外锚点(段首) + trackChanges=true → 仍然成功,锚点不动', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="outsideBk"/>
        <w:bookmarkEnd w:id="1"/>
        <w:r><w:t>hello world</w:t></w:r>
      </w:p>`,
      { trackChanges: true },
    );
    executeReplaceTextInParagraph(ctx, replaceOp('AAAAAAA1', 'world', 'everyone'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    // bookmarkStart 仍在段首,未被包进 <w:del>
    expect(xml).toMatch(/<w:bookmarkStart[^>]*w:name="outsideBk"[^>]*\/>[\s\S]*<w:del/);
    expect(xml).toMatch(/<w:ins[^>]*>[\s\S]*<w:t[^>]*>everyone<\/w:t>[\s\S]*<\/w:ins>/);
  });
});

describe('executeReplaceParagraph + trackChanges=true — 整段锚点跟随包入 <w:del>', () => {
  test('段内 bookmarkStart → 不抛,<w:del> 含 bookmarkStart', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="1" w:name="myBm"/>
        <w:r><w:t>text</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
      </w:p>`,
      { trackChanges: true },
    );
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'new' });
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:del[^>]*>[\s\S]*<w:bookmarkStart[^>]*w:name="myBm"[\s\S]*<\/w:del>/);
    expect(xml).toContain('<w:delText>text</w:delText>');
    expect(xml).toMatch(/<w:ins[^>]*>[\s\S]*<w:t[^>]*>new<\/w:t>[\s\S]*<\/w:ins>/);
  });

  test('段内 commentRangeStart/End/Reference → 不抛,三个 marker 都进 <w:del>', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:commentRangeStart w:id="5"/>
        <w:r><w:t>text</w:t></w:r>
        <w:commentRangeEnd w:id="5"/>
        <w:r><w:commentReference w:id="5"/></w:r>
      </w:p>`,
      { trackChanges: true },
    );
    executeReplaceParagraph(ctx, { type: 'replaceParagraph', paraId: 'AAAAAAA1', text: 'new' });
    const xml = paraXml(pkg, 'AAAAAAA1');
    // 三个 marker 都被包进 <w:del>(单个 <w:del> 包了所有旧内容)
    expect(xml).toMatch(/<w:del[^>]*>[\s\S]*<w:commentRangeStart[^>]*w:id="5"/);
    expect(xml).toMatch(/<w:del[^>]*>[\s\S]*<w:commentRangeEnd[^>]*w:id="5"/);
    expect(xml).toMatch(/<w:del[^>]*>[\s\S]*<w:commentReference[^>]*w:id="5"/);
    expect(xml).toContain('<w:delText>text</w:delText>');
  });
});
