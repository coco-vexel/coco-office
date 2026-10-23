/**
 * listParagraphs 纯函数测试(#24)。
 *
 * 覆盖:
 *   - scope:'body' / 'all' / `StoryId[]`(body / header / footer / footnote)
 *   - filter:styleId / contains / hasComments / hasRevisions(及组合)
 *   - textFormat:full / preview(首 100 字 + truncated)
 *   - pagination:默认 byCount(25,0)/ 'all' / byCount 分页 / byChar 四条规则
 *   - ParagraphInfo 的 hasBookmarks / hasRevisions / hasComments 检测
 *   - 边界:空档、非法 pageSize / pageIndex / charPerPage、scope=StoryId[] 命中与未命中
 *
 * worker 模式端到端见 [word-document.test.ts](./word-document.test.ts) /
 * [worker-mode.test.ts](./worker-mode.test.ts) 的新增用例(本 PR 同步补)。
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { DEFAULT_PAGINATION, PREVIEW_CHAR_LIMIT, listParagraphs } from '../list-paragraphs.js';
import { W14_NS, W_NS } from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const HDR_CTYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FTR_CTYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const FOOTNOTES_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml';

function putXmlPart(pkg: Package, path: string, xml: string, contentType = DOC_CTYPE): void {
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

/** 只有 body,包含 N 个段落,每段文本为 `prefix-<i>`(可选 styleId) */
function makeBodyOnly(n: number, prefix = 'p', styleId?: string): Package {
  const pkg = createPackage();
  const paras = Array.from({ length: n }, (_, i) => {
    const pid = String(i).padStart(8, 'A').slice(-8).toUpperCase();
    const pPr = styleId ? `<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` : '';
    return `<w:p w14:paraId="${pid}">${pPr}<w:r><w:t>${prefix}-${i}</w:t></w:r></w:p>`;
  }).join('');
  putXmlPart(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${paras}</w:body></w:document>`,
  );
  return pkg;
}

// ─────────── scope ───────────

describe('listParagraphs — scope', () => {
  test('默认 scope=body:只返回 body 段落', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1"><w:r><w:t>body-1</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA2"><w:r><w:t>body-2</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="BBBBBBB1"><w:r><w:t>header</w:t></w:r></w:p>
      </w:hdr>`,
      HDR_CTYPE,
    );

    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    expect(r.paragraphs.map((p) => p.paraId)).toEqual(['AAAAAAA1', 'AAAAAAA2']);
    expect(r.paragraphs.every((p) => p.story === 'body')).toBe(true);
  });

  test("scope='all':返回 body + header + footer + footnote", () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body><w:p w14:paraId="BODYBODY"><w:r><w:t>body</w:t></w:r></w:p></w:body>
      </w:document>`,
    );
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="HDRHDRHD"><w:r><w:t>h</w:t></w:r></w:p>
      </w:hdr>`,
      HDR_CTYPE,
    );
    putXmlPart(
      pkg,
      'word/footer1.xml',
      `<w:ftr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="FTRFTRFT"><w:r><w:t>f</w:t></w:r></w:p>
      </w:ftr>`,
      FTR_CTYPE,
    );
    putXmlPart(
      pkg,
      'word/footnotes.xml',
      `<w:footnotes xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:footnote w:id="1">
          <w:p w14:paraId="FN1FN1FN"><w:r><w:t>fn</w:t></w:r></w:p>
        </w:footnote>
      </w:footnotes>`,
      FOOTNOTES_CTYPE,
    );

    const r = listParagraphs(pkg, { scope: 'all', pagination: { mode: 'all' } });
    expect(r.paragraphs.map((p) => p.paraId).sort()).toEqual([
      'BODYBODY',
      'FN1FN1FN',
      'FTRFTRFT',
      'HDRHDRHD',
    ]);
    // story 值正确分布
    const stories = new Set(r.paragraphs.map((p) => p.story));
    expect(stories.has('body')).toBe(true);
    expect(stories.has('header:word/header1.xml')).toBe(true);
    expect(stories.has('footer:word/footer1.xml')).toBe(true);
    expect(stories.has('footnote:1')).toBe(true);
  });

  test('scope=StoryId[]:精确指定命中,未命中的 id 忽略', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body><w:p w14:paraId="BODYBODY"><w:r><w:t>body</w:t></w:r></w:p></w:body>
      </w:document>`,
    );
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="HDRHDRHD"><w:r><w:t>h</w:t></w:r></w:p>
      </w:hdr>`,
      HDR_CTYPE,
    );

    // 只要 header,且同时传一个不存在的 id
    const r = listParagraphs(pkg, {
      scope: ['header:word/header1.xml', 'footnote:999'],
      pagination: { mode: 'all' },
    });
    expect(r.paragraphs.map((p) => p.paraId)).toEqual(['HDRHDRHD']);
  });

  test('空档(scope=body 但无段落):返回空 + totalPages=0', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}"><w:body></w:body></w:document>`,
    );

    const r = listParagraphs(pkg);
    expect(r.paragraphs).toEqual([]);
    expect(r.pagination.totalParagraphs).toBe(0);
    expect(r.pagination.totalPages).toBe(0);
    expect(r.pagination.hasNext).toBe(false);
  });
});

// ─────────── filter ───────────

describe('listParagraphs — filter', () => {
  test('styleId:严格等值匹配', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>h1</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA2"><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>n</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA3"><w:r><w:t>no-style</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );

    const r = listParagraphs(pkg, {
      filter: { styleId: 'Heading1' },
      pagination: { mode: 'all' },
    });
    expect(r.paragraphs.map((p) => p.paraId)).toEqual(['AAAAAAA1']);
    expect(r.pagination.totalParagraphs).toBe(1); // filter 后的总数
  });

  test('contains:区分大小写的子串匹配', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello world</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA2"><w:r><w:t>Hello World</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA3"><w:r><w:t>foo</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );

    const r = listParagraphs(pkg, {
      filter: { contains: 'hello' },
      pagination: { mode: 'all' },
    });
    // 'Hello World' 不命中(大小写敏感)
    expect(r.paragraphs.map((p) => p.paraId)).toEqual(['AAAAAAA1']);
  });

  test('hasComments:命中 commentRangeStart / commentReference', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:commentRangeStart w:id="1"/>
            <w:r><w:t>commented</w:t></w:r>
            <w:commentRangeEnd w:id="1"/>
            <w:r><w:commentReference w:id="1"/></w:r>
          </w:p>
          <w:p w14:paraId="AAAAAAA2"><w:r><w:t>clean</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );

    const r = listParagraphs(pkg, {
      filter: { hasComments: true },
      pagination: { mode: 'all' },
    });
    expect(r.paragraphs.map((p) => p.paraId)).toEqual(['AAAAAAA1']);

    // 反向 filter
    const r2 = listParagraphs(pkg, {
      filter: { hasComments: false },
      pagination: { mode: 'all' },
    });
    expect(r2.paragraphs.map((p) => p.paraId)).toEqual(['AAAAAAA2']);
  });

  test('hasRevisions:命中内容和属性修订', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:ins w:id="1" w:author="x" w:date="2026-04-22T00:00:00Z">
              <w:r><w:t>inserted</w:t></w:r>
            </w:ins>
          </w:p>
          <w:p w14:paraId="AAAAAAA2">
            <w:del w:id="2" w:author="x" w:date="2026-04-22T00:00:00Z">
              <w:r><w:delText>deleted</w:delText></w:r>
            </w:del>
          </w:p>
          <w:p w14:paraId="AAAAAAA3"><w:r><w:t>clean</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA4"><w:pPr><w:pPrChange w:id="3"><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>formatted</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );

    const r = listParagraphs(pkg, {
      filter: { hasRevisions: true },
      pagination: { mode: 'all' },
    });
    expect(r.paragraphs.map((p) => p.paraId).sort()).toEqual([
      'AAAAAAA1',
      'AAAAAAA2',
      'AAAAAAA4',
    ]);
  });

  test('组合 filter:styleId + contains', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>alpha</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA2"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>beta</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA3"><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>alpha</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );

    const r = listParagraphs(pkg, {
      filter: { styleId: 'Heading1', contains: 'alpha' },
      pagination: { mode: 'all' },
    });
    expect(r.paragraphs.map((p) => p.paraId)).toEqual(['AAAAAAA1']);
  });
});

// ─────────── textFormat ───────────

describe('listParagraphs — textFormat', () => {
  test("默认 'full':完整文本", () => {
    const pkg = makeBodyOnly(1, 'p', undefined);
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    expect(r.paragraphs[0]!.text).toBe('p-0');
    expect(r.paragraphs[0]!.truncated).toBeUndefined();
  });

  test("'preview' 不截断(<= 100 字):不设 truncated", () => {
    const pkg = createPackage();
    const shortText = 'a'.repeat(50);
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1"><w:r><w:t>${shortText}</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const r = listParagraphs(pkg, { textFormat: 'preview', pagination: { mode: 'all' } });
    expect(r.paragraphs[0]!.text).toBe(shortText);
    expect(r.paragraphs[0]!.truncated).toBeUndefined();
  });

  test("'preview' 截断(> 100 字):text = 首 100 字 + truncated=true", () => {
    const pkg = createPackage();
    const longText = 'x'.repeat(PREVIEW_CHAR_LIMIT + 50);
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1"><w:r><w:t>${longText}</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const r = listParagraphs(pkg, { textFormat: 'preview', pagination: { mode: 'all' } });
    expect(r.paragraphs[0]!.text.length).toBe(PREVIEW_CHAR_LIMIT);
    expect(r.paragraphs[0]!.truncated).toBe(true);
  });
});

// ─────────── pagination ───────────

describe('listParagraphs — pagination', () => {
  test('默认 pagination = byCount(25, 0):返回前 25 条 + hasNext 按总数算', () => {
    const pkg = makeBodyOnly(60);
    const r = listParagraphs(pkg); // 不传 pagination
    expect(r.paragraphs).toHaveLength(25);
    expect(r.pagination.mode).toBe('byCount');
    expect(r.pagination.pageIndex).toBe(0);
    // meta 不回显 pageSize(契约中没有);类型上也不存在,用索引访问确认运行时 undefined
    expect((r.pagination as Record<string, unknown>).pageSize).toBeUndefined();
    expect(r.pagination.totalPages).toBe(3); // ceil(60/25)
    expect(r.pagination.totalParagraphs).toBe(60);
    expect(r.pagination.hasNext).toBe(true);
    // autoPaginated 恒不填(A 方案语义)
    expect(r.pagination.autoPaginated).toBeUndefined();
  });

  test('默认 pagination 不动态切换:3 段也是 byCount(25,0)', () => {
    const pkg = makeBodyOnly(3);
    const r = listParagraphs(pkg);
    // 返回 3 条(一页装不满)
    expect(r.paragraphs).toHaveLength(3);
    expect(r.pagination.mode).toBe('byCount');
    expect(r.pagination.totalPages).toBe(1);
    expect(r.pagination.hasNext).toBe(false);
  });

  test("'all':hasNext=false, totalPages=1(非空时)", () => {
    const pkg = makeBodyOnly(7);
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    expect(r.paragraphs).toHaveLength(7);
    expect(r.pagination.mode).toBe('all');
    expect(r.pagination.totalPages).toBe(1);
    expect(r.pagination.hasNext).toBe(false);
  });

  test('byCount 中间页:正确切片 + hasNext', () => {
    const pkg = makeBodyOnly(60);
    const r = listParagraphs(pkg, { pagination: { mode: 'byCount', pageSize: 25, pageIndex: 1 } });
    expect(r.paragraphs).toHaveLength(25);
    expect(r.paragraphs[0]!.text).toBe('p-25');
    expect(r.paragraphs[24]!.text).toBe('p-49');
    expect(r.pagination.hasNext).toBe(true);
  });

  test('byCount 最后一页:slice 剩余 + hasNext=false', () => {
    const pkg = makeBodyOnly(60);
    const r = listParagraphs(pkg, { pagination: { mode: 'byCount', pageSize: 25, pageIndex: 2 } });
    expect(r.paragraphs).toHaveLength(10); // 60 - 50
    expect(r.paragraphs[0]!.text).toBe('p-50');
    expect(r.paragraphs[9]!.text).toBe('p-59');
    expect(r.pagination.hasNext).toBe(false);
  });

  test('byCount 越界 pageIndex:返回空,hasNext=false', () => {
    const pkg = makeBodyOnly(60);
    const r = listParagraphs(pkg, { pagination: { mode: 'byCount', pageSize: 25, pageIndex: 10 } });
    expect(r.paragraphs).toEqual([]);
    expect(r.pagination.totalPages).toBe(3);
    expect(r.pagination.hasNext).toBe(false);
  });

  test('byCount 非法参数:抛错', () => {
    const pkg = makeBodyOnly(3);
    expect(() =>
      listParagraphs(pkg, { pagination: { mode: 'byCount', pageSize: 0, pageIndex: 0 } }),
    ).toThrow(/pageSize must be a positive integer/);
    expect(() =>
      listParagraphs(pkg, { pagination: { mode: 'byCount', pageSize: 5, pageIndex: -1 } }),
    ).toThrow(/pageIndex must be a non-negative integer/);
    expect(() =>
      listParagraphs(pkg, { pagination: { mode: 'byCount', pageSize: 1.5, pageIndex: 0 } }),
    ).toThrow(/pageSize/);
  });

  test('byChar 规则 1+4:同页累加直到超限', () => {
    const pkg = makePkgByCharFixture([10, 10, 10, 10]); // 4 段,各 10 字
    const r = listParagraphs(pkg, {
      pagination: { mode: 'byChar', charPerPage: 25, pageIndex: 0 },
    });
    // 规则 4:10 + 10 = 20 <= 25 → 追加;20 + 10 = 30 > 25 → 下一页
    // 所以 page 0 是前 2 段
    expect(r.paragraphs.map((p) => p.text)).toEqual(['a'.repeat(10), 'a'.repeat(10)]);
    expect(r.pagination.totalPages).toBe(2);
    expect(r.pagination.hasNext).toBe(true);
  });

  test('byChar 规则 2:空页且单段 > charPerPage → 单段成页(完整返回)', () => {
    const pkg = makePkgByCharFixture([100, 10, 10]); // 100 字段落 + 两段小
    const r = listParagraphs(pkg, {
      pagination: { mode: 'byChar', charPerPage: 30, pageIndex: 0 },
    });
    // page 0:单段 100 字(规则 2)
    expect(r.paragraphs).toHaveLength(1);
    expect(r.paragraphs[0]!.text.length).toBe(100);
    // page 1:10 + 10 = 20 <= 30 → 两段
    const r1 = listParagraphs(pkg, {
      pagination: { mode: 'byChar', charPerPage: 30, pageIndex: 1 },
    });
    expect(r1.paragraphs).toHaveLength(2);
  });

  test('byChar 规则 3:累加超限时当前段进下一页(不切段)', () => {
    const pkg = makePkgByCharFixture([15, 15]); // 两段各 15 字
    const r = listParagraphs(pkg, {
      pagination: { mode: 'byChar', charPerPage: 20, pageIndex: 0 },
    });
    // 0 + 15 = 15 <= 20 → 追加第 1 段(buf=15)
    // 15 + 15 = 30 > 20 → 结束当前页;第 2 段进下一页
    // page 0 = [第1段];page 1 = [第2段]
    expect(r.paragraphs).toHaveLength(1);
    expect(r.pagination.totalPages).toBe(2);
  });

  test('byChar 非法参数:抛错', () => {
    const pkg = makeBodyOnly(3);
    expect(() =>
      listParagraphs(pkg, { pagination: { mode: 'byChar', charPerPage: 0, pageIndex: 0 } }),
    ).toThrow(/charPerPage must be a positive integer/);
  });
});

// ─────────── ParagraphInfo 字段 ───────────

describe('listParagraphs — ParagraphInfo 字段', () => {
  test('hasBookmarks 检测:命中 bookmarkStart', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:bookmarkStart w:id="1" w:name="b1"/>
            <w:r><w:t>x</w:t></w:r>
            <w:bookmarkEnd w:id="1"/>
          </w:p>
          <w:p w14:paraId="AAAAAAA2"><w:r><w:t>clean</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    expect(r.paragraphs[0]!.hasBookmarks).toBe(true);
    expect(r.paragraphs[1]!.hasBookmarks).toBeUndefined();
  });

  test('paraId 大写化(与 indices.paraIdx 一致)', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="abc12345"><w:r><w:t>lower</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    expect(r.paragraphs[0]!.paraId).toBe('ABC12345');
  });

  test('styleId 未设时缺字段,不是 undefined 字符串', () => {
    const pkg = makeBodyOnly(1, 'p'); // 无 pStyle
    const r = listParagraphs(pkg);
    expect(Object.prototype.hasOwnProperty.call(r.paragraphs[0], 'styleId')).toBe(false);
  });
});

// ─────────── 默认值导出 ───────────

describe('listParagraphs — 默认值导出', () => {
  test('DEFAULT_PAGINATION 对齐规范', () => {
    expect(DEFAULT_PAGINATION).toEqual({ mode: 'byCount', pageSize: 25, pageIndex: 0 });
  });

  test('PREVIEW_CHAR_LIMIT = 100', () => {
    expect(PREVIEW_CHAR_LIMIT).toBe(100);
  });
});

// ─────────── helpers ───────────

/** 构造 body 只含指定字数段落的 pkg(用于 byChar 测试) */
function makePkgByCharFixture(charCounts: number[]): Package {
  const pkg = createPackage();
  const paras = charCounts
    .map((n, i) => {
      const pid = String(i).padStart(8, 'A').slice(-8).toUpperCase();
      return `<w:p w14:paraId="${pid}"><w:r><w:t>${'a'.repeat(n)}</w:t></w:r></w:p>`;
    })
    .join('');
  putXmlPart(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${paras}</w:body></w:document>`,
  );
  return pkg;
}
