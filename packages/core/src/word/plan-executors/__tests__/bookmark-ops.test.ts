/**
 * bookmark-ops executor 纯函数测试(#34)。
 *
 * 覆盖:
 *   createBookmark:
 *     - ParagraphRange 基本:段内生成 bookmarkStart/End,w:id 为 0(empty pkg)
 *     - ParagraphRange:w:id 递增,第二个书签 w:id = 1
 *     - ParagraphRange 段已有 pPr:bookmarkStart 在 pPr **之后**,bookmarkEnd 在段末
 *     - TextRange 同 run:在 <w:t> 中间找文本,切 run 后插 start/end 包住匹配
 *     - TextRange 跨 run:match 横跨 2 个 <w:r>
 *     - TextRange 文本不命中:TEXT_NOT_FOUND_IN_PARA
 *     - BOOKMARK_NAME_EXISTS:重名拒绝
 *     - BOOKMARK_NAME_INVALID:空 name
 *     - BOOKMARK_NAME_INVALID:首字符非字母/下划线
 *     - BOOKMARK_NAME_INVALID:含特殊字符
 *     - BOOKMARK_NAME_INVALID:超 40 字符
 *     - 下划线开头 _Toc42:校验通过
 *     - PARA_NOT_FOUND
 *     - trackChanges=true 拒绝(含 #31)
 *     - ApplyResult.createdBookmarks 字段完整
 *     - 复杂 run 下 TextRange 抛 UNSUPPORTED_OOXML_FEATURE
 *   deleteBookmark:
 *     - 基本:删除后段内再无 bookmarkStart/End
 *     - BOOKMARK_NOT_FOUND
 *     - 隐藏书签(_Toc42)也能删
 *     - trackChanges=true 拒绝
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult, Op } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import {
  executeCreateBookmark,
  executeDeleteBookmark,
  resolveAppendBookmarkName,
} from '../bookmark-ops.js';

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

function paraXml(pkg: Package, paraId: string): string {
  const idx = buildIndices(pkg);
  const el = idx.paraIdx.get(paraId.toUpperCase())!;
  return serializeToWellFormedString(el as never);
}

function createOp(name: string, paraId: string): Extract<Op, { type: 'createBookmark' }> {
  return { type: 'createBookmark', name, range: { paraId } };
}

function createTextOp(
  name: string,
  paraId: string,
  text: string,
): Extract<Op, { type: 'createBookmark' }> {
  return { type: 'createBookmark', name, range: { paraId, text } };
}

function deleteOp(name: string): Extract<Op, { type: 'deleteBookmark' }> {
  return { type: 'deleteBookmark', name };
}

// ─────────── createBookmark ───────────

describe('executeCreateBookmark', () => {
  test('ParagraphRange 基本:段内生成 bookmarkStart/End,w:id 为 0(empty pkg)', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    executeCreateBookmark(ctx, createOp('foo', 'AAAAAAA1'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/w:bookmarkStart/);
    expect(xml).toMatch(/w:bookmarkEnd/);
    expect(xml).toMatch(/w:id="0"/);
    expect(xml).toMatch(/w:name="foo"/);
  });

  test('ParagraphRange:w:id 递增,第二个书签 w:id = 1', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>p1</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>p2</w:t></w:r></w:p>`,
    );
    executeCreateBookmark(ctx, createOp('bm1', 'AAAAAAA1'));
    // 重建索引后插入第二个书签
    ctx.indices = buildIndices(pkg);
    executeCreateBookmark({ ...ctx, opIndex: 1 }, createOp('bm2', 'AAAAAAA2'));
    const xml2 = paraXml(pkg, 'AAAAAAA2');
    expect(xml2).toMatch(/w:id="1"/);
  });

  test('ParagraphRange 段已有 pPr:bookmarkStart 在 pPr 之后,bookmarkEnd 在段末', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r><w:t>content</w:t></w:r>
      </w:p>`,
    );
    executeCreateBookmark(ctx, createOp('sec', 'AAAAAAA1'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    // pPr 应该在 bookmarkStart 之前
    const pPrIdx = xml.indexOf('pPr');
    const bStartIdx = xml.indexOf('bookmarkStart');
    const bEndIdx = xml.indexOf('bookmarkEnd');
    expect(pPrIdx).toBeLessThan(bStartIdx);
    // bookmarkEnd 在 run 之后(末尾)
    const runIdx = xml.indexOf('w:r');
    expect(bEndIdx).toBeGreaterThan(runIdx);
  });

  test('TextRange 同 run:在 <w:t> 中间找文本,切 run 后插 start/end 包住匹配', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello world end</w:t></w:r></w:p>`,
    );
    executeCreateBookmark(ctx, createTextOp('bm', 'AAAAAAA1', 'world'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    // bookmarkStart 应该在第二个 run(包含"world")之前
    expect(xml).toMatch(/w:bookmarkStart/);
    expect(xml).toMatch(/w:bookmarkEnd/);
    // "hello " 应在 bookmarkStart 之前,"world" 应在 start/end 之间
    const bStartIdx = xml.indexOf('bookmarkStart');
    const bEndIdx = xml.indexOf('bookmarkEnd');
    const helloIdx = xml.indexOf('hello');
    const worldIdx = xml.indexOf('world');
    expect(helloIdx).toBeLessThan(bStartIdx);
    expect(worldIdx).toBeGreaterThan(bStartIdx);
    expect(worldIdx).toBeLessThan(bEndIdx);
  });

  test('TextRange 跨 run:match 横跨 2 个 <w:r>', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>hel</w:t></w:r>
        <w:r><w:t>lo world</w:t></w:r>
      </w:p>`,
    );
    // "lo" 跨 run 边界("hel" + "lo")
    executeCreateBookmark(ctx, createTextOp('cross', 'AAAAAAA1', 'hello'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/w:bookmarkStart/);
    expect(xml).toMatch(/w:bookmarkEnd/);
  });

  test('TextRange 文本不命中:TEXT_NOT_FOUND_IN_PARA', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    expect(() => executeCreateBookmark(ctx, createTextOp('bm', 'AAAAAAA1', 'notfound'))).toThrow(
      expect.objectContaining({ code: 'TEXT_NOT_FOUND_IN_PARA' }) as unknown as Error,
    );
  });

  test('BOOKMARK_NAME_EXISTS:重名拒绝', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="existing"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    expect(() => executeCreateBookmark(ctx, createOp('existing', 'AAAAAAA1'))).toThrow(
      expect.objectContaining({ code: 'BOOKMARK_NAME_EXISTS' }) as unknown as Error,
    );
  });

  test('BOOKMARK_NAME_INVALID:空 name', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    try {
      executeCreateBookmark(ctx, createOp('', 'AAAAAAA1'));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('BOOKMARK_NAME_INVALID');
      expect((e as WordError).message).toMatch(/empty/);
    }
  });

  test('BOOKMARK_NAME_INVALID:首字符为数字', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(() => executeCreateBookmark(ctx, createOp('1invalid', 'AAAAAAA1'))).toThrow(
      expect.objectContaining({ code: 'BOOKMARK_NAME_INVALID' }) as unknown as Error,
    );
  });

  test('BOOKMARK_NAME_INVALID:含特殊字符(空格)', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(() => executeCreateBookmark(ctx, createOp('bad name', 'AAAAAAA1'))).toThrow(
      expect.objectContaining({ code: 'BOOKMARK_NAME_INVALID' }) as unknown as Error,
    );
  });

  test('BOOKMARK_NAME_INVALID:超 40 字符', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    const longName = 'a'.repeat(41);
    try {
      executeCreateBookmark(ctx, createOp(longName, 'AAAAAAA1'));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('BOOKMARK_NAME_INVALID');
      expect((e as WordError).message).toMatch(/41/);
    }
  });

  test('下划线开头 _Toc42:校验通过', () => {
    const { pkg, ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    // 不应抛错
    executeCreateBookmark(ctx, createOp('_Toc42', 'AAAAAAA1'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/w:name="_Toc42"/);
  });

  test('PARA_NOT_FOUND', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(() => executeCreateBookmark(ctx, createOp('bm', 'FFFFFFFF'))).toThrow(
      expect.objectContaining({ code: 'PARA_NOT_FOUND' }) as unknown as Error,
    );
  });

  test('trackChanges=true 拒绝(含 #31)', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    const trackCtx = { ...ctx, options: { ...ctx.options, trackChanges: true } };
    try {
      executeCreateBookmark(trackCtx, createOp('bm', 'AAAAAAA1'));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
      expect((e as WordError).message).toMatch(/#31/);
    }
  });

  test('ApplyResult.createdBookmarks 字段完整', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    ctx.opIndex = 3;
    executeCreateBookmark(ctx, createOp('mybm', 'AAAAAAA1'));
    expect(ctx.result.createdBookmarks).toEqual([
      { opIndex: 3, name: 'mybm', range: { fromParaId: 'AAAAAAA1', toParaId: 'AAAAAAA1' } },
    ]);
  });

  test('复杂 run 下 TextRange 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    // 多 <w:t> 的复杂 run:splitRunAt 会抛错
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>hel</w:t><w:t>lo</w:t></w:r>
        <w:r><w:t> world</w:t></w:r>
      </w:p>`,
    );
    try {
      executeCreateBookmark(ctx, createTextOp('bm', 'AAAAAAA1', 'lo wo'));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
    }
  });
});

// ─────────── deleteBookmark ───────────

describe('executeDeleteBookmark', () => {
  test('基本:删除后段内再无 bookmarkStart/End', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="mybm"/>
        <w:r><w:t>content</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    executeDeleteBookmark(ctx, deleteOp('mybm'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/bookmarkStart|bookmarkEnd/);
  });

  test('BOOKMARK_NOT_FOUND', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(() => executeDeleteBookmark(ctx, deleteOp('nonexistent'))).toThrow(
      expect.objectContaining({ code: 'BOOKMARK_NOT_FOUND' }) as unknown as Error,
    );
  });

  test('隐藏书签(_Toc42)也能删', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="_Toc42"/>
        <w:r><w:t>toc</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    // 隐藏书签也在 bookmarkIdx 中(装载期不过滤,listBookmarks 才过滤)
    executeDeleteBookmark(ctx, deleteOp('_Toc42'));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/bookmarkStart|bookmarkEnd/);
  });

  test('trackChanges=true 拒绝', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="mybm"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    const trackCtx = { ...ctx, options: { ...ctx.options, trackChanges: true } };
    try {
      executeDeleteBookmark(trackCtx, deleteOp('mybm'));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
      expect((e as WordError).message).toMatch(/#31/);
    }
  });

  test('索引同步:deleteBookmark 后 bookmarkIdx 不含该 name', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="todel"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    expect(ctx.indices.bookmarkIdx.has('todel')).toBe(true);
    executeDeleteBookmark(ctx, deleteOp('todel'));
    expect(ctx.indices.bookmarkIdx.has('todel')).toBe(false);
  });
});

// ─────────── resolveAppendBookmarkName(#176)───────────

describe('resolveAppendBookmarkName', () => {
  test('基础公式:{prefix}_{anchorParaId}', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(resolveAppendBookmarkName(ctx, '_co_append', 'AAAAAAA1')).toBe('_co_append_AAAAAAA1');
  });

  test('同 anchor 重名 → 追加 _2 后缀', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="_co_append_AAAAAAA1"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    expect(resolveAppendBookmarkName(ctx, '_co_append', 'AAAAAAA1')).toBe('_co_append_AAAAAAA1_2');
  });

  test('多次冲突 → _2 / _3 / _4 递增', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="tag_AAAAAAA1"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
        <w:bookmarkStart w:id="1" w:name="tag_AAAAAAA1_2"/>
        <w:r><w:t>y</w:t></w:r>
        <w:bookmarkEnd w:id="1"/>
        <w:bookmarkStart w:id="2" w:name="tag_AAAAAAA1_3"/>
        <w:r><w:t>z</w:t></w:r>
        <w:bookmarkEnd w:id="2"/>
      </w:p>`,
    );
    expect(resolveAppendBookmarkName(ctx, 'tag', 'AAAAAAA1')).toBe('tag_AAAAAAA1_4');
  });

  test('不同 prefix 互不干扰', () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:bookmarkStart w:id="0" w:name="alpha_AAAAAAA1"/>
        <w:r><w:t>x</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
      </w:p>`,
    );
    // alpha_ 已占,beta_ 不冲突
    expect(resolveAppendBookmarkName(ctx, 'beta', 'AAAAAAA1')).toBe('beta_AAAAAAA1');
  });

  test('prefix 非法字符 → BOOKMARK_NAME_INVALID', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    try {
      resolveAppendBookmarkName(ctx, '1bad', 'AAAAAAA1');
      throw new Error('expected throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('BOOKMARK_NAME_INVALID');
    }
  });

  test('prefix 空串 → BOOKMARK_NAME_INVALID', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    try {
      resolveAppendBookmarkName(ctx, '', 'AAAAAAA1');
      throw new Error('expected throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('BOOKMARK_NAME_INVALID');
    }
  });

  test('解析后名长度 > 40 → BOOKMARK_NAME_INVALID', () => {
    // 30 字符 prefix + '_' + 8 字符 paraId = 39 → 边界过(刚好 ≤ 40);+1 字符则超
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    const longPrefix = 'a'.repeat(32); // 32 + 1 + 8 = 41
    try {
      resolveAppendBookmarkName(ctx, longPrefix, 'AAAAAAA1');
      throw new Error('expected throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('BOOKMARK_NAME_INVALID');
      expect((e as WordError).message).toMatch(/40-char/);
    }
  });

  test('下划线起头的 prefix 通过校验(Word hidden bookmark 约定)', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(resolveAppendBookmarkName(ctx, '_hidden', 'AAAAAAA1')).toBe('_hidden_AAAAAAA1');
  });
});
