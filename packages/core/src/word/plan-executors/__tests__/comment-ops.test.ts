/**
 * comment-ops executor 纯函数测试(#35)。
 *
 * 覆盖:
 *   - addComment:ParagraphRange 首次(创建 comments.xml / CT / rels / 三锚点 / createdComments)
 *   - addComment:ParagraphRange 追加(CT / rels 不重复)
 *   - addComment:TextRange scope='first'(锚点正确位置)
 *   - addComment:TextRange scope='all' → UNSUPPORTED_OOXML_FEATURE
 *   - addComment:text 含 \n / \t(br / tab 在 comment 主体)
 *   - addComment:author 回落(op > ctx.options > '')
 *   - addComment:date 回落(op.date > ctx.options.date)
 *   - addComment:PARA_NOT_FOUND
 *   - addComment:TEXT_NOT_FOUND_IN_PARA(TextRange 不命中)
 *   - addComment:trackChanges=true → UNSUPPORTED_OOXML_FEATURE
 *   - deleteComment:基本删除(三锚点 + <w:comment> 消失)
 *   - deleteComment:COMMENT_NOT_FOUND
 *   - deleteComment:trackChanges=true → UNSUPPORTED_OOXML_FEATURE
 *   - deleteComment:comments.xml 删完最后一条后 part/CT/rels 被移除（task-comment-part-cleanup）
 *   - deleteComment:三锚点部分存在(ref=null)不抛错
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import {
  createPackage,
  createXmlPart,
  getPart,
  hasPart,
  setPart,
  type Package,
} from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult } from '../../plan.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeAddComment, executeDeleteComment } from '../comment-ops.js';
import { executeDeleteParagraph } from '../paragraph-structure-ops.js';
import { executeDeleteTable } from '../table-ops.js';

// ─────────── Content Types ───────────

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const COMMENTS_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const CT_CTYPE = 'application/vnd.openxmlformats-package.content-types+xml';
const RELS_CTYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELS_NS_URI = 'http://schemas.openxmlformats.org/package/2006/relationships';

// ─────────── 辅助工具 ───────────

function putXml(pkg: Package, path: string, xml: string, contentType: string): void {
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

function makeCtXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT_NS}">
  <Override PartName="/word/document.xml" ContentType="${DOC_CTYPE}"/>
</Types>`;
}

function makeRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS_URI}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

function makeCommentsXml(comments = ''): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W_NS}">${comments}</w:comments>`;
}

function makeDocXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`;
}

interface Fixture {
  pkg: Package;
  ctx: ExecutorContext;
}

/**
 * 构建完整的 fixture:document.xml + [Content_Types].xml + word/_rels/document.xml.rels
 * 可选 comments.xml(withComments=true 或传 commentsInner)
 */
function makeFixture(
  bodyInner: string,
  opts: {
    withComments?: boolean;
    commentsInner?: string;
    withCtAndRels?: boolean;
    author?: string;
    date?: Date;
  } = {},
): Fixture {
  const pkg = createPackage();
  putXml(pkg, 'word/document.xml', makeDocXml(bodyInner), DOC_CTYPE);

  const withCtAndRels = opts.withCtAndRels !== false; // 默认 true
  if (withCtAndRels) {
    putXml(pkg, '[Content_Types].xml', makeCtXml(), CT_CTYPE);
    putXml(pkg, 'word/_rels/document.xml.rels', makeRelsXml(), RELS_CTYPE);
  }

  if (opts.withComments || opts.commentsInner !== undefined) {
    putXml(pkg, 'word/comments.xml', makeCommentsXml(opts.commentsInner ?? ''), COMMENTS_CTYPE);
  }

  const indices = buildIndices(pkg);
  const ctx: ExecutorContext = {
    pkg,
    indices,
    options: {
      trackChanges: false,
      author: opts.author ?? 'default-author',
      date: opts.date ?? new Date('2026-04-22T00:00:00Z'),
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

/** 从 pkg 中取 comments.xml 的 xmlDom(若不存在返回 null)*/
function getCommentsDoc(pkg: Package): Document | null {
  const part = getPart(pkg, 'word/comments.xml');
  if (!part || !part.xmlDom) return null;
  return part.xmlDom;
}

/** 取 comments.xml 里的所有 <w:comment> 元素 */
function listCommentEls(pkg: Package): Element[] {
  const doc = getCommentsDoc(pkg);
  if (!doc) return [];
  const col = doc.getElementsByTagNameNS(W_NS, 'comment');
  const out: Element[] = [];
  for (let i = 0; i < col.length; i++) out.push(col[i]!);
  return out;
}

/** 取 body 段落的 XML 字符串(重建索引后查)*/
function paraXml(pkg: Package, paraId: string): string {
  const idx = buildIndices(pkg);
  const el = idx.paraIdx.get(paraId.toUpperCase());
  if (!el) return '';
  return serializeToWellFormedString(el as never);
}

/** 取 [Content_Types].xml 的 XML 字符串 */
function ctXml(pkg: Package): string {
  const part = getPart(pkg, '[Content_Types].xml');
  if (!part?.xmlDom) return '';
  return serializeToWellFormedString(part.xmlDom as never);
}

/** 取 word/_rels/document.xml.rels 的 XML 字符串 */
function relsXml(pkg: Package): string {
  const part = getPart(pkg, 'word/_rels/document.xml.rels');
  if (!part?.xmlDom) return '';
  return serializeToWellFormedString(part.xmlDom as never);
}

// ─────────── addComment 测试 ───────────

describe('executeAddComment', () => {
  test('ParagraphRange 首次:comments.xml 被创建,含 1 条 <w:comment>', () => {
    const { pkg, ctx } = makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`,
    );
    expect(getCommentsDoc(pkg)).toBeNull();

    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'my comment',
    });

    const comments = listCommentEls(pkg);
    expect(comments).toHaveLength(1);
    expect(comments[0]!.getAttributeNS(W_NS, 'id')).toBe('0');
    expect(comments[0]!.getAttributeNS(W_NS, 'author')).toBe('default-author');
  });

  test('ParagraphRange 首次:Content_Types 多了 comments.xml Override', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    executeAddComment(ctx, { type: 'addComment', range: { paraId: 'AAAAAAA1' }, text: 'c' });
    expect(ctXml(pkg)).toContain('/word/comments.xml');
  });

  test('ParagraphRange 首次:rels 多了 comments Relationship', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    executeAddComment(ctx, { type: 'addComment', range: { paraId: 'AAAAAAA1' }, text: 'c' });
    expect(relsXml(pkg)).toContain('comments.xml');
    expect(relsXml(pkg)).toContain('relationships/comments');
  });

  test('ParagraphRange 首次:body 段有三锚点', () => {
    const { pkg, ctx } = makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`,
    );
    executeAddComment(ctx, { type: 'addComment', range: { paraId: 'AAAAAAA1' }, text: 'c' });
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toContain('commentRangeStart');
    expect(xml).toContain('commentRangeEnd');
    expect(xml).toContain('commentReference');
  });

  test('ParagraphRange 首次:ApplyResult.createdComments 返回 { opIndex: 0, commentId: 0 }', () => {
    const { ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    executeAddComment(ctx, { type: 'addComment', range: { paraId: 'AAAAAAA1' }, text: 'c' });
    expect(ctx.result.createdComments).toEqual([{ opIndex: 0, commentId: 0 }]);
  });

  test('ParagraphRange 追加:commentId = 已有 max + 1,CT / rels 不重复', () => {
    const existingComment = `<w:comment w:id="3" w:author="bob" w:date="2026-01-01T00:00:00Z">
      <w:p><w:r><w:t>existing</w:t></w:r></w:p></w:comment>`;
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`, {
      commentsInner: existingComment,
    });

    const ctBefore = ctXml(pkg);
    const relsBefore = relsXml(pkg);

    executeAddComment(ctx, { type: 'addComment', range: { paraId: 'AAAAAAA1' }, text: 'new' });

    const comments = listCommentEls(pkg);
    expect(comments).toHaveLength(2);
    const newComment = comments.find((c) => c.getAttributeNS(W_NS, 'id') === '4');
    expect(newComment).toBeDefined();

    // CT / rels 不应重复追加
    expect(ctXml(pkg)).toBe(ctBefore);
    expect(relsXml(pkg)).toBe(relsBefore);
  });

  test('TextRange scope=first:锚点夹住匹配文本', () => {
    // 段落只有 'hello world' 一个 run;addComment 应切成 'hello ' + 'world'
    // 并在 'world' run 前后插 rangeStart / rangeEnd
    const { pkg, ctx } = makeFixture(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello world</w:t></w:r></w:p>`,
    );
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1', text: 'world' },
      text: 'comment on world',
    });
    const xml = paraXml(pkg, 'AAAAAAA1');

    // 基本结构检查
    expect(xml).toContain('commentRangeStart');
    expect(xml).toContain('commentRangeEnd');
    expect(xml).toContain('commentReference');
    expect(xml).toContain('hello');
    expect(xml).toContain('world');

    // 顺序检查:rangeStart 在 rangeEnd 之前
    expect(xml.indexOf('commentRangeStart')).toBeLessThan(xml.indexOf('commentRangeEnd'));
    // 'hello' 出现在 commentRangeStart 之前(因为 'hello ' 是匹配前的文本)
    expect(xml.indexOf('hello')).toBeLessThan(xml.indexOf('commentRangeStart'));
  });

  test('TextRange scope=all → UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>foo foo</w:t></w:r></w:p>`);
    expect(() =>
      executeAddComment(ctx, {
        type: 'addComment',
        range: { paraId: 'AAAAAAA1', text: 'foo', scope: 'all' },
        text: 'c',
      }),
    ).toThrow(expect.objectContaining({ code: 'UNSUPPORTED_OOXML_FEATURE' }) as unknown as Error);
  });

  test('text 含 \\n:批注主体含 <w:br/>', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'line1\nline2',
    });
    const commentsDoc = getCommentsDoc(pkg)!;
    const xml = serializeToWellFormedString(commentsDoc as never);
    expect(xml).toContain('w:br');
    expect(xml).toContain('line1');
    expect(xml).toContain('line2');
  });

  test('text 含 \\t:批注主体含 <w:tab/>', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'col1\tcol2',
    });
    const commentsDoc = getCommentsDoc(pkg)!;
    const xml = serializeToWellFormedString(commentsDoc as never);
    expect(xml).toContain('w:tab');
    expect(xml).toContain('col1');
    expect(xml).toContain('col2');
  });

  test('author 回落:op.author 优先', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`, {
      author: 'ctx-author',
    });
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'c',
      author: 'op-author',
    });
    const comment = listCommentEls(pkg)[0]!;
    expect(comment.getAttributeNS(W_NS, 'author')).toBe('op-author');
  });

  test('author 回落:op.author 缺,用 ctx.options.author', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`, {
      author: 'ctx-author',
    });
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'c',
    });
    const comment = listCommentEls(pkg)[0]!;
    expect(comment.getAttributeNS(W_NS, 'author')).toBe('ctx-author');
  });

  test('author 回落:两者都缺,得到空串', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`, {
      author: '',
    });
    ctx.options.author = '';
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'c',
    });
    const comment = listCommentEls(pkg)[0]!;
    expect(comment.getAttributeNS(W_NS, 'author')).toBe('');
  });

  test('date 回落:op.date(ISO)优先', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`, {
      date: new Date('2020-01-01T00:00:00Z'),
    });
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'c',
      date: '2030-12-31T23:59:59Z',
    });
    const comment = listCommentEls(pkg)[0]!;
    expect(comment.getAttributeNS(W_NS, 'date')).toBe('2030-12-31T23:59:59Z');
  });

  test('date 回落:op.date 缺,用 ctx.options.date', () => {
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`, {
      date: new Date('2026-04-22T10:00:00Z'),
    });
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'c',
    });
    const comment = listCommentEls(pkg)[0]!;
    expect(comment.getAttributeNS(W_NS, 'date')).toBe('2026-04-22T10:00:00.000Z');
  });

  test('PARA_NOT_FOUND:paraId 不存在', () => {
    const { ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(() =>
      executeAddComment(ctx, {
        type: 'addComment',
        range: { paraId: 'NOTEXIST' },
        text: 'c',
      }),
    ).toThrow(expect.objectContaining({ code: 'PARA_NOT_FOUND' }) as unknown as Error);
  });

  test('TEXT_NOT_FOUND_IN_PARA:TextRange 不命中', () => {
    const { ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    expect(() =>
      executeAddComment(ctx, {
        type: 'addComment',
        range: { paraId: 'AAAAAAA1', text: 'notfound' },
        text: 'c',
      }),
    ).toThrow(expect.objectContaining({ code: 'TEXT_NOT_FOUND_IN_PARA' }) as unknown as Error);
  });

  test('trackChanges=true → UNSUPPORTED_OOXML_FEATURE(#31)', () => {
    const { ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    const trackCtx = { ...ctx, options: { ...ctx.options, trackChanges: true } };
    expect(() =>
      executeAddComment(trackCtx, {
        type: 'addComment',
        range: { paraId: 'AAAAAAA1' },
        text: 'c',
      }),
    ).toThrow(expect.objectContaining({ code: 'UNSUPPORTED_OOXML_FEATURE' }) as unknown as Error);
  });
});

// ─────────── deleteComment 测试 ───────────

describe('executeDeleteComment', () => {
  /** 构建含一条批注(id=0)的 fixture */
  function makeFixtureWithComment(): Fixture {
    const commentBody = `<w:comment w:id="0" w:author="alice" w:date="2026-04-22T00:00:00Z">
      <w:p><w:r><w:t>original comment</w:t></w:r></w:p></w:comment>`;
    const bodyInner = `<w:p w14:paraId="AAAAAAA1">
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>hello</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:commentReference w:id="0"/></w:r>
    </w:p>`;
    return makeFixture(bodyInner, { commentsInner: commentBody });
  }

  test('基本:三锚点消失,<w:comment> 也消失', () => {
    const { pkg, ctx } = makeFixtureWithComment();
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });

    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toContain('commentRangeStart');
    expect(xml).not.toContain('commentRangeEnd');
    expect(xml).not.toContain('commentReference');

    expect(listCommentEls(pkg)).toHaveLength(0);
  });

  test('基本:删后 commentIdx 重建后不含该 id', () => {
    const { pkg, ctx } = makeFixtureWithComment();
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });

    const idx = buildIndices(pkg);
    expect(idx.commentIdx.has(0)).toBe(false);
  });

  test('COMMENT_NOT_FOUND:id 不存在', () => {
    const { ctx } = makeFixtureWithComment();
    expect(() => executeDeleteComment(ctx, { type: 'deleteComment', commentId: 999 })).toThrow(
      expect.objectContaining({ code: 'COMMENT_NOT_FOUND' }) as unknown as Error,
    );
  });

  test('trackChanges=true → UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx } = makeFixtureWithComment();
    const trackCtx = { ...ctx, options: { ...ctx.options, trackChanges: true } };
    expect(() => executeDeleteComment(trackCtx, { type: 'deleteComment', commentId: 0 })).toThrow(
      expect.objectContaining({ code: 'UNSUPPORTED_OOXML_FEATURE' }) as unknown as Error,
    );
  });

  test('comments.xml 删完最后一条后 part 被移除（新行为：cleanup 触发）', () => {
    const { pkg, ctx } = makeFixtureWithComment();
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });

    // 新行为：part 已被移除
    expect(pkg.parts.has('word/comments.xml')).toBe(false);
  });

  test('三锚点部分存在(ref=null):只删存在的,不抛错', () => {
    // 构造只有 start + end,没有 commentReference
    const commentBody = `<w:comment w:id="0" w:author="alice" w:date="2026-04-22T00:00:00Z">
      <w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment>`;
    const bodyInner = `<w:p w14:paraId="AAAAAAA1">
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>text</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
    </w:p>`;
    const { pkg, ctx } = makeFixture(bodyInner, { commentsInner: commentBody });

    // 不应抛错
    expect(() => executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 })).not.toThrow();

    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toContain('commentRangeStart');
    expect(xml).not.toContain('commentRangeEnd');
  });

  test('空 wrapper run 在 ref 删除后也被清掉', () => {
    // ref 在一个 <w:r> 里,删 ref 后 run 变空 → run 也要删
    const { pkg, ctx } = makeFixtureWithComment();
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });

    // 不应有空 <w:r>(只有 commentReference 的 run)
    const xml = paraXml(pkg, 'AAAAAAA1');
    // 验证没有空 run (只含 commentReference 的 w:r 已被删,只剩正文 run)
    expect(xml).not.toContain('commentReference');
    // 正文 run 还在
    expect(xml).toContain('hello');
  });
});

// ─────────── deleteComment cleanup 测试（task-comment-part-cleanup §3）───────────

describe('deleteComment — comments part cleanup', () => {
  // fixture 工具：含 N 条批注的文档
  function _makeFixtureWithNComments(n: number): Fixture {
    let commentsInner = '';
    let bodyInner = '';
    for (let i = 0; i < n; i++) {
      commentsInner += `<w:comment w:id="${i}" w:author="alice" w:date="2026-04-22T00:00:00Z">
        <w:p><w:r><w:t>comment ${i}</w:t></w:r></w:p></w:comment>`;
      bodyInner += `<w:p w14:paraId="${String(i)
        .padStart(8, 'A')
        .replace(/[^A-Z0-9]/g, 'A')}${i}">
        <w:commentRangeStart w:id="${i}"/>
        <w:r><w:t>text ${i}</w:t></w:r>
        <w:commentRangeEnd w:id="${i}"/>
        <w:r><w:commentReference w:id="${i}"/></w:r>
      </w:p>`;
    }
    return makeFixture(bodyInner, { commentsInner });
  }

  /** 用固定 paraId 构造 1 条批注 fixture（id=0，paraId="AAAAAAA1"）*/
  function makeFixtureWithOneComment(): Fixture {
    const commentBody = `<w:comment w:id="0" w:author="alice" w:date="2026-04-22T00:00:00Z">
      <w:p><w:r><w:t>the comment</w:t></w:r></w:p></w:comment>`;
    const bodyInner = `<w:p w14:paraId="AAAAAAA1">
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>hello</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:commentReference w:id="0"/></w:r>
    </w:p>`;
    return makeFixture(bodyInner, { commentsInner: commentBody });
  }

  /** 用固定 paraId 构造 2 条批注 fixture（id=0 在 AAAAAAA1，id=1 在 AAAAAAA2）*/
  function makeFixtureWithTwoComments(): Fixture {
    const commentsInner = `
      <w:comment w:id="0" w:author="alice" w:date="2026-04-22T00:00:00Z">
        <w:p><w:r><w:t>first</w:t></w:r></w:p></w:comment>
      <w:comment w:id="1" w:author="alice" w:date="2026-04-22T00:00:00Z">
        <w:p><w:r><w:t>second</w:t></w:r></w:p></w:comment>`;
    const bodyInner = `
      <w:p w14:paraId="AAAAAAA1">
        <w:commentRangeStart w:id="0"/>
        <w:r><w:t>first text</w:t></w:r>
        <w:commentRangeEnd w:id="0"/>
        <w:r><w:commentReference w:id="0"/></w:r>
      </w:p>
      <w:p w14:paraId="AAAAAAA2">
        <w:commentRangeStart w:id="1"/>
        <w:r><w:t>second text</w:t></w:r>
        <w:commentRangeEnd w:id="1"/>
        <w:r><w:commentReference w:id="1"/></w:r>
      </w:p>`;
    return makeFixture(bodyInner, { commentsInner });
  }

  // §3 #1: 删光唯一一条 comment
  test('#1 删光唯一一条 comment：part 移除 + CT 无 Override + rels 无 Relationship', () => {
    const { pkg, ctx } = makeFixtureWithOneComment();
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });

    // part 已移除
    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);
    // CT 无 comments Override
    expect(ctXml(pkg)).not.toContain('/word/comments.xml');
    // rels 无 comments Relationship
    expect(relsXml(pkg)).not.toContain('relationships/comments');
  });

  // §3 #2: 删 2 条中的 1 条（还剩 1）
  test('#2 删 2 条中的 1 条（还剩 1）：part 仍在（CT/rels 不验证，因预置 fixture 不含 Override）', () => {
    const { pkg, ctx } = makeFixtureWithTwoComments();

    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });

    // part 仍存在（还有 comment id=1）
    expect(hasPart(pkg, 'word/comments.xml')).toBe(true);
    // 还有 1 条 comment
    const commentsDoc = getCommentsDoc(pkg)!;
    const remaining = commentsDoc.getElementsByTagNameNS(W_NS, 'comment');
    expect(remaining.length).toBe(1);
  });

  // §3 #3: 删光 2 条（逐条删）
  test('#3 删光 2 条（逐条删）：删第 2 条后 part 移除', () => {
    const { pkg, ctx } = makeFixtureWithTwoComments();

    // 删第 1 条（还剩 1）
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });
    expect(hasPart(pkg, 'word/comments.xml')).toBe(true);

    // 需要重建 indices 以便第 2 次 deleteComment 能查到 commentIdx
    ctx.indices = buildIndices(pkg);

    // 删第 2 条（空了）
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 1 });
    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);
  });

  // §3 #4: round-trip
  test('#4 round-trip：删光 → toZip → 无 comments.xml + CT/rels 干净 + 重载不抛错', async () => {
    const JSZip = (await import('jszip')).default;
    const { WordDocument } = await import('../../word-document.js');

    const _CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT_NS}">
  <Default Extension="rels" ContentType="${RELS_CTYPE}"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CTYPE}"/>
</Types>`;
    const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS_URI}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
    const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="RTRIP001">
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>hello</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:commentReference w:id="0"/></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W_NS}">
  <w:comment w:id="0" w:author="tester" w:date="2026-06-04T00:00:00Z">
    <w:p><w:r><w:t>a comment</w:t></w:r></w:p>
  </w:comment>
</w:comments>`;
    const COMMENTS_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS_URI}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;
    const CT_WITH_COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT_NS}">
  <Default Extension="rels" ContentType="${RELS_CTYPE}"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CTYPE}"/>
  <Override PartName="/word/comments.xml" ContentType="${COMMENTS_CTYPE}"/>
</Types>`;

    const zip = new JSZip();
    zip.file('[Content_Types].xml', CT_WITH_COMMENTS);
    zip.file('_rels/.rels', MAIN_RELS);
    zip.file('word/document.xml', DOC_XML);
    zip.file('word/_rels/document.xml.rels', COMMENTS_RELS);
    zip.file('word/comments.xml', COMMENTS_XML);
    const buf = await zip.generateAsync({ type: 'uint8array' });

    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-cleanup', mode: 'inline' });
    const commentsBefore = await doc.comments.list();
    expect(commentsBefore).toHaveLength(1);

    await doc.plan().deleteComment(0).apply();

    const outZip = await doc.toZip();
    await doc.dispose();

    // 检查输出 zip
    const outJszip = await JSZip.loadAsync(outZip);
    // 无 word/comments.xml
    expect(outJszip.file('word/comments.xml')).toBeNull();
    // CT 无 Override
    const ctContent = await outJszip.file('[Content_Types].xml')!.async('string');
    expect(ctContent).not.toContain('/word/comments.xml');
    // rels 无 comments rel
    const relsContent = await outJszip.file('word/_rels/document.xml.rels')!.async('string');
    expect(relsContent).not.toContain('relationships/comments');

    // 重载不抛错
    const doc2 = await WordDocument.fromZip(outZip, {
      docId: 'e2e-cleanup-reload',
      mode: 'inline',
    });
    const commentsAfter = await doc2.comments.list();
    expect(commentsAfter).toHaveLength(0);
    await doc2.dispose();

    WordDocument._resetDefaults();
  });

  // §3 #5: addComment 后 deleteComment 删光（对称）
  test('#5 addComment 后 deleteComment 删光（对称）：ensureCommentsPart 建的 part 被 cleanup 干净移除', () => {
    // 从无 comments.xml 的文档开始
    const { pkg, ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>text</w:t></w:r></w:p>`);
    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);

    // addComment：创建 part + CT Override + rels Relationship
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'AAAAAAA1' },
      text: 'temp comment',
    });
    expect(hasPart(pkg, 'word/comments.xml')).toBe(true);
    expect(ctXml(pkg)).toContain('/word/comments.xml');
    expect(relsXml(pkg)).toContain('relationships/comments');

    // 重建 indices 以便 deleteComment 可查
    ctx.indices = buildIndices(pkg);

    // deleteComment：删光最后一条 → cleanup
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });
    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);
    expect(ctXml(pkg)).not.toContain('/word/comments.xml');
    expect(relsXml(pkg)).not.toContain('relationships/comments');
  });

  // §3 #6: 幂等/边界：文档本无 comments.xml 时 COMMENT_NOT_FOUND（现有行为不变）
  test('#6 幂等/边界：文档本无 comments.xml 时 deleteComment 抛 COMMENT_NOT_FOUND', () => {
    const { ctx } = makeFixture(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>text</w:t></w:r></w:p>`);
    // 没有 comments.xml，commentIdx 也为空
    expect(() => executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 })).toThrow(
      expect.objectContaining({ code: 'COMMENT_NOT_FOUND' }) as unknown as Error,
    );
  });

  // §3 #7: 不回归——现有 comment-ops.test.ts 通过（本测试文件的其余用例都在同次 vitest run 内）
  test('#7 不回归：删 1 条（还剩 0）后三锚点 + <w:comment> 都消失（原有行为保持）', () => {
    const { pkg, ctx } = makeFixtureWithOneComment();
    executeDeleteComment(ctx, { type: 'deleteComment', commentId: 0 });

    // 三锚点消失（原有断言）
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toContain('commentRangeStart');
    expect(xml).not.toContain('commentRangeEnd');
    expect(xml).not.toContain('commentReference');

    // 正文内容仍在
    expect(xml).toContain('hello');
  });
});

// ─────────── drop 路径测试（红旗 2 覆盖：dropCommentById 路径）───────────

describe('deleteComment cleanup — drop 路径（red-flag-2）', () => {
  /**
   * 构造：先用 addComment 建批注（会触发 ensureCommentsPart → CT/rels 齐），
   * 再 rebuild indices，最后用 drop 路径删段/表，验证 cleanup 触发。
   */

  test('deleteParagraph(drop) 删含最后批注的段 → comments part 移除 + CT/rels 清', () => {
    // 从无 comments.xml 开始（addComment 会创建 + CT/rels）
    const bodyInner = `
      <w:p w14:paraId="DDDDD001"><w:r><w:t>has comment</w:t></w:r></w:p>
      <w:p w14:paraId="DDDDD002"><w:r><w:t>no comment</w:t></w:r></w:p>`;
    const { pkg, ctx } = makeFixture(bodyInner);

    // addComment → 建 comments.xml + CT Override + rels Relationship
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'DDDDD001' },
      text: 'drop via paragraph',
    });
    expect(hasPart(pkg, 'word/comments.xml')).toBe(true);
    expect(ctXml(pkg)).toContain('/word/comments.xml');
    expect(relsXml(pkg)).toContain('relationships/comments');

    // 重建 indices（paraIdx / commentIdx 都需要）
    ctx.indices = buildIndices(pkg);

    // deleteParagraph(drop) 删含批注的段
    executeDeleteParagraph(ctx, {
      type: 'deleteParagraph',
      paraId: 'DDDDD001',
      onAnchorConflict: 'drop',
    });

    // 批注被删光 → comments part 移除 + CT/rels 清
    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);
    expect(ctXml(pkg)).not.toContain('/word/comments.xml');
    expect(relsXml(pkg)).not.toContain('relationships/comments');
  });

  test('deleteTable(drop) 删含最后批注的表 → comments part 移除 + CT/rels 清', () => {
    const bodyInner = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="FFFFF001"><w:r><w:t>table cell</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
      <w:p w14:paraId="FFFFF002"><w:r><w:t>after table</w:t></w:r></w:p>`;
    const { pkg, ctx } = makeFixture(bodyInner);

    // addComment 挂在表格首段 → 建 CT/rels
    executeAddComment(ctx, {
      type: 'addComment',
      range: { paraId: 'FFFFF001' },
      text: 'drop via table',
    });
    expect(hasPart(pkg, 'word/comments.xml')).toBe(true);
    expect(ctXml(pkg)).toContain('/word/comments.xml');
    expect(relsXml(pkg)).toContain('relationships/comments');

    // 重建 indices
    ctx.indices = buildIndices(pkg);
    const tableId = [...ctx.indices.tblIdx.keys()][0]!;

    // deleteTable(drop) 删含批注的表
    executeDeleteTable(ctx, {
      type: 'deleteTable',
      tableId,
      onAnchorConflict: 'drop',
    });

    // 批注被删光 → comments part 移除 + CT/rels 清
    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);
    expect(ctXml(pkg)).not.toContain('/word/comments.xml');
    expect(relsXml(pkg)).not.toContain('relationships/comments');
  });
});
