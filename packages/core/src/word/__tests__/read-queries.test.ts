/**
 * 只读查询纯函数测试(#25):getParagraphByParaId / listBookmarks / getBookmark。
 *
 * 覆盖:
 *   - getParagraph:命中(大小写不敏感)、未命中、body 段 / header 段 story 反查
 *   - listBookmarks:hidden 默认过滤 / includeHidden=true 全量 / 顺序稳定
 *   - getBookmark:hidden 依然返回(不过滤)
 *   - rangeKind 判定:paragraph / text / complex(缺 end / 跨段 / parent 非 `<w:p>`)
 *   - text rangeKind 的 text 字段:跨 run / 跨 hyperlink / 跨 ins-del 的拼接正确
 *   - hidden 字段:`_` 开头为 true,其它缺字段
 *   - paraId 字段:text / paragraph 给,complex 缺
 *
 * worker 端到端见 [read-queries-worker.test.ts](./read-queries-worker.test.ts)。
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { getBookmark, listBookmarks } from '../bookmark-queries.js';
import { buildIndices } from '../indices.js';
import { getParagraphByParaId } from '../list-paragraphs.js';
import { W14_NS, W_NS } from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const HDR_CTYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';

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

// ─────────── getParagraphByParaId ───────────

describe('getParagraphByParaId', () => {
  test('命中:大小写不敏感,返回 ParagraphInfo 含 story=body', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="abcd1234"><w:r><w:t>hello</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const idx = buildIndices(pkg);

    const p1 = getParagraphByParaId(pkg, idx, 'ABCD1234');
    expect(p1).not.toBeNull();
    expect(p1!.paraId).toBe('ABCD1234');
    expect(p1!.text).toBe('hello');
    expect(p1!.story).toBe('body');

    // 小写传入效果等价
    const p2 = getParagraphByParaId(pkg, idx, 'abcd1234');
    expect(p2?.paraId).toBe('ABCD1234');
  });

  test('未命中:返回 null', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body><w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p></w:body>
      </w:document>`,
    );
    const idx = buildIndices(pkg);
    expect(getParagraphByParaId(pkg, idx, 'NOTEXIST')).toBeNull();
  });

  test('header 段 story 反查正确', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body></w:body></w:document>`,
    );
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="HHHHHHHH"><w:r><w:t>hdr</w:t></w:r></w:p>
      </w:hdr>`,
      HDR_CTYPE,
    );
    const idx = buildIndices(pkg);
    const p = getParagraphByParaId(pkg, idx, 'HHHHHHHH');
    expect(p?.story).toBe('header:word/header1.xml');
  });
});

// ─────────── listBookmarks / getBookmark ───────────

describe('listBookmarks / getBookmark', () => {
  function makeBookmarkFixture(): { pkg: Package; indices: ReturnType<typeof buildIndices> } {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <!-- paragraph 覆盖整段 -->
          <w:p w14:paraId="AAAAAAA1">
            <w:bookmarkStart w:id="1" w:name="full-para"/>
            <w:r><w:t>entire</w:t></w:r>
            <w:bookmarkEnd w:id="1"/>
          </w:p>
          <!-- text:段内片段 -->
          <w:p w14:paraId="AAAAAAA2">
            <w:r><w:t>prefix </w:t></w:r>
            <w:bookmarkStart w:id="2" w:name="middle"/>
            <w:r><w:t>middle</w:t></w:r>
            <w:bookmarkEnd w:id="2"/>
            <w:r><w:t> suffix</w:t></w:r>
          </w:p>
          <!-- complex:跨段落 -->
          <w:p w14:paraId="AAAAAAA3">
            <w:bookmarkStart w:id="3" w:name="cross"/>
            <w:r><w:t>first</w:t></w:r>
          </w:p>
          <w:p w14:paraId="AAAAAAA4">
            <w:r><w:t>second</w:t></w:r>
            <w:bookmarkEnd w:id="3"/>
          </w:p>
          <!-- 隐藏书签 -->
          <w:p w14:paraId="AAAAAAA5">
            <w:bookmarkStart w:id="4" w:name="_Toc123"/>
            <w:r><w:t>toc</w:t></w:r>
            <w:bookmarkEnd w:id="4"/>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const indices = buildIndices(pkg);
    return { pkg, indices };
  }

  test('默认过滤隐藏书签(_ 开头)', () => {
    const { pkg, indices } = makeBookmarkFixture();
    const names = listBookmarks(pkg, indices).map((b) => b.name);
    expect(names).toEqual(['full-para', 'middle', 'cross']);
  });

  test('includeHidden=true 返回全部,保持插入序', () => {
    const { pkg, indices } = makeBookmarkFixture();
    const names = listBookmarks(pkg, indices, { includeHidden: true }).map((b) => b.name);
    expect(names).toEqual(['full-para', 'middle', 'cross', '_Toc123']);
  });

  test("rangeKind='paragraph':整段覆盖,paraId 给、text 不给", () => {
    const { pkg, indices } = makeBookmarkFixture();
    const b = listBookmarks(pkg, indices).find((x) => x.name === 'full-para')!;
    expect(b.rangeKind).toBe('paragraph');
    expect(b.paraId).toBe('AAAAAAA1');
    expect(b.text).toBeUndefined();
    expect(b.story).toBe('body');
  });

  test("rangeKind='text':段内片段,text 字段为 start/end 之间的拼接", () => {
    const { pkg, indices } = makeBookmarkFixture();
    const b = listBookmarks(pkg, indices).find((x) => x.name === 'middle')!;
    expect(b.rangeKind).toBe('text');
    expect(b.paraId).toBe('AAAAAAA2');
    expect(b.text).toBe('middle');
  });

  test("rangeKind='complex':跨段落,paraId / text 都不给", () => {
    const { pkg, indices } = makeBookmarkFixture();
    const b = listBookmarks(pkg, indices).find((x) => x.name === 'cross')!;
    expect(b.rangeKind).toBe('complex');
    expect(b.paraId).toBeUndefined();
    expect(b.text).toBeUndefined();
  });

  test("rangeKind='complex':缺 bookmarkEnd", () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:bookmarkStart w:id="1" w:name="orphan"/>
            <w:r><w:t>x</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const indices = buildIndices(pkg);
    const b = getBookmark(pkg, indices, 'orphan');
    expect(b?.rangeKind).toBe('complex');
  });

  test("rangeKind='complex':parent 不是 <w:p>(表 cell 直子)", () => {
    // 极少见:w:tc 直属 bookmarkStart/End(包 cell 整体),非 w:p 子
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:tbl>
            <w:tr>
              <w:tc>
                <w:bookmarkStart w:id="1" w:name="cell-wrap"/>
                <w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>
                <w:bookmarkEnd w:id="1"/>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>`,
    );
    const indices = buildIndices(pkg);
    const b = getBookmark(pkg, indices, 'cell-wrap');
    expect(b?.rangeKind).toBe('complex');
  });

  test('text rangeKind:hyperlink 内的 w:t 也被计入', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:r><w:t>pre</w:t></w:r>
            <w:bookmarkStart w:id="1" w:name="link"/>
            <w:hyperlink>
              <w:r><w:t>clickme</w:t></w:r>
            </w:hyperlink>
            <w:bookmarkEnd w:id="1"/>
            <w:r><w:t>post</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const indices = buildIndices(pkg);
    const b = getBookmark(pkg, indices, 'link');
    expect(b?.rangeKind).toBe('text');
    expect(b?.text).toBe('clickme');
  });

  test('paragraph rangeKind:bookmark 前后只有其他 marker 节点也算整段', () => {
    // comment 标记在书签前后不影响 paragraph 判定(标记非 content)
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
            <w:commentRangeStart w:id="5"/>
            <w:bookmarkStart w:id="1" w:name="wrap"/>
            <w:r><w:t>body</w:t></w:r>
            <w:bookmarkEnd w:id="1"/>
            <w:commentRangeEnd w:id="5"/>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const indices = buildIndices(pkg);
    const b = getBookmark(pkg, indices, 'wrap');
    expect(b?.rangeKind).toBe('paragraph');
  });

  test('hidden 字段:_ 开头为 true,其它缺字段', () => {
    const { pkg, indices } = makeBookmarkFixture();
    const visible = getBookmark(pkg, indices, 'full-para')!;
    expect(Object.prototype.hasOwnProperty.call(visible, 'hidden')).toBe(false);
    const toc = getBookmark(pkg, indices, '_Toc123')!;
    expect(toc.hidden).toBe(true);
  });

  test('getBookmark 未命中:返回 null', () => {
    const { pkg, indices } = makeBookmarkFixture();
    expect(getBookmark(pkg, indices, 'nonexistent')).toBeNull();
  });
});
