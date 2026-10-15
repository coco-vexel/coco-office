import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { buildIndices, BookmarkLoadError } from '../indices.js';
import { W_NS, W14_NS } from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const HDR_CTYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const COMMENTS_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

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

describe('buildIndices — paraIdx', () => {
  test('覆盖 body / header / comments 的所有 <w:p>', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:t>body1</w:t></w:r></w:p>
          <w:p w14:paraId="BBBBBBBB"><w:r><w:t>body2</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="CCCCCCCC"><w:r><w:t>hdr</w:t></w:r></w:p>
      </w:hdr>`,
      HDR_CTYPE,
    );

    const { paraIdx } = buildIndices(pkg);
    expect(paraIdx.size).toBe(3);
    expect(paraIdx.has('AAAAAAAA')).toBe(true);
    expect(paraIdx.has('BBBBBBBB')).toBe(true);
    expect(paraIdx.has('CCCCCCCC')).toBe(true);
  });

  test('paraId 大小写不敏感:取自 attribute 的值 toUpperCase 存入', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="deadbeef"><w:r><w:t>t</w:t></w:r></w:p>
      </w:document>`,
    );
    const { paraIdx } = buildIndices(pkg);
    expect(paraIdx.has('DEADBEEF')).toBe(true);
    expect(paraIdx.has('deadbeef')).toBe(false);
  });

  test('缺 paraId 的段落不入索引', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p><w:r><w:t>no id</w:t></w:r></w:p>
        <w:p w14:paraId="11111111"><w:r><w:t>has id</w:t></w:r></w:p>
      </w:document>`,
    );
    const { paraIdx } = buildIndices(pkg);
    expect(paraIdx.size).toBe(1);
    expect(paraIdx.has('11111111')).toBe(true);
  });

  test('非 Word content part 不参与枚举(如 images)', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="AAAAAAAA"><w:r><w:t>x</w:t></w:r></w:p>
      </w:document>`,
    );
    // 伪装成外部 XML(非 wordprocessingml 命名空间)
    putXmlPart(
      pkg,
      'custom.xml',
      `<root xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="ZZZZZZZZ"><w:r><w:t>not-word-content</w:t></w:r></w:p>
      </root>`,
      'application/xml',
    );
    const { paraIdx } = buildIndices(pkg);
    expect(paraIdx.has('AAAAAAAA')).toBe(true);
    expect(paraIdx.has('ZZZZZZZZ')).toBe(false);
  });
});

describe('buildIndices — bookmarkIdx', () => {
  test('正常配对的 bookmark:name → { start, end, partPath }', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p>
            <w:bookmarkStart w:id="1" w:name="Intro"/>
            <w:r><w:t>content</w:t></w:r>
            <w:bookmarkEnd w:id="1"/>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const { bookmarkIdx } = buildIndices(pkg);
    expect(bookmarkIdx.size).toBe(1);
    const entry = bookmarkIdx.get('Intro')!;
    expect(entry.start.localName).toBe('bookmarkStart');
    expect(entry.end?.localName).toBe('bookmarkEnd');
    expect(entry.partPath).toBe('word/document.xml');
  });

  test('缺 bookmarkEnd:start 仍入索引,end=null', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body><w:p><w:bookmarkStart w:id="1" w:name="Lonely"/></w:p></w:body>
      </w:document>`,
    );
    const { bookmarkIdx } = buildIndices(pkg);
    expect(bookmarkIdx.get('Lonely')?.end).toBeNull();
  });

  test('孤儿 bookmarkEnd(无 start)被忽略,不影响索引', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body><w:p><w:bookmarkEnd w:id="99"/></w:p></w:body>
      </w:document>`,
    );
    const { bookmarkIdx } = buildIndices(pkg);
    expect(bookmarkIdx.size).toBe(0);
  });

  test('跨 part 的同名 bookmark 容错:不抛错,保留第一个(document.xml 的)', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body><w:p><w:bookmarkStart w:id="1" w:name="Shared"/><w:bookmarkEnd w:id="1"/></w:p></w:body>
      </w:document>`,
    );
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr xmlns:w="${W_NS}">
        <w:p><w:bookmarkStart w:id="1" w:name="Shared"/><w:bookmarkEnd w:id="1"/></w:p>
      </w:hdr>`,
      HDR_CTYPE,
    );

    // 重名书签**不再抛错**(真实文档常见,绝不崩);保留第一个、忽略后续。
    let bookmarkIdx!: ReturnType<typeof buildIndices>['bookmarkIdx'];
    expect(() => {
      bookmarkIdx = buildIndices(pkg).bookmarkIdx;
    }).not.toThrow();
    expect(bookmarkIdx.has('Shared')).toBe(true);
    // 保留的是第一个遇到的(parts 迭代序中 document.xml 在前)
    expect(bookmarkIdx.get('Shared')!.partPath).toBe('word/document.xml');
  });

  test('同 part 内重名 bookmark 容错:不抛错,只入 1 条', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:bookmarkStart w:id="1" w:name="Dup"/><w:bookmarkEnd w:id="1"/></w:p>
          <w:p><w:bookmarkStart w:id="2" w:name="Dup"/><w:bookmarkEnd w:id="2"/></w:p>
        </w:body>
      </w:document>`,
    );
    let bookmarkIdx!: ReturnType<typeof buildIndices>['bookmarkIdx'];
    expect(() => {
      bookmarkIdx = buildIndices(pkg).bookmarkIdx;
    }).not.toThrow();
    expect(bookmarkIdx.has('Dup')).toBe(true);
    expect([...bookmarkIdx.keys()].filter((k) => k === 'Dup')).toHaveLength(1);
  });

  test('重复 _GoBack 书签(真实文档 crash 复现):不抛错,正常入索引', () => {
    // repro：多次编辑 / 多工具处理过的真实文档常出现重复的 _GoBack，
    // 旧逻辑遇重名书签抛 BOOKMARK_NAME_DUPLICATE_ON_LOAD → 排版崩溃。
    // 修复后必须容错（保留第一个、其余忽略），其它正常书签不受影响。
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p>
          <w:p><w:r><w:t>正文段落</w:t></w:r></w:p>
          <w:p><w:bookmarkStart w:id="1" w:name="_GoBack"/><w:bookmarkEnd w:id="1"/></w:p>
          <w:p><w:bookmarkStart w:id="2" w:name="Real"/><w:bookmarkEnd w:id="2"/></w:p>
        </w:body>
      </w:document>`,
    );
    let result!: ReturnType<typeof buildIndices>;
    expect(() => {
      result = buildIndices(pkg);
    }).not.toThrow();
    expect(result.bookmarkIdx.has('_GoBack')).toBe(true);
    expect(result.bookmarkIdx.has('Real')).toBe(true);
    // _GoBack 只登记一条
    expect([...result.bookmarkIdx.keys()].filter((k) => k === '_GoBack')).toHaveLength(1);
  });

  test('空 name 抛 BOOKMARK_NAME_EMPTY_ON_LOAD', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body><w:p><w:bookmarkStart w:id="1" w:name=""/><w:bookmarkEnd w:id="1"/></w:p></w:body>
      </w:document>`,
    );
    // 空名是真正的结构损坏(非编辑痕迹),仍抛 BookmarkLoadError —— 仅重名改为容错。
    expect(() => buildIndices(pkg)).toThrowError(BookmarkLoadError);
    expect(() => buildIndices(pkg)).toThrowError(
      expect.objectContaining({ code: 'BOOKMARK_NAME_EMPTY_ON_LOAD' }),
    );
  });

  test('隐藏书签(_Toc... / _GoBack)与普通书签一视同仁入索引', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body><w:p>
          <w:bookmarkStart w:id="1" w:name="_Toc123"/><w:bookmarkEnd w:id="1"/>
          <w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/>
          <w:bookmarkStart w:id="3" w:name="Visible"/><w:bookmarkEnd w:id="3"/>
        </w:p></w:body>
      </w:document>`,
    );
    const { bookmarkIdx } = buildIndices(pkg);
    expect(bookmarkIdx.has('_Toc123')).toBe(true);
    expect(bookmarkIdx.has('_GoBack')).toBe(true);
    expect(bookmarkIdx.has('Visible')).toBe(true);
  });
});

describe('buildIndices — commentIdx', () => {
  test('锚点三件套 + 主体正确关联', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body><w:p>
          <w:commentRangeStart w:id="5"/>
          <w:r><w:t>anchor text</w:t></w:r>
          <w:commentRangeEnd w:id="5"/>
          <w:r><w:commentReference w:id="5"/></w:r>
        </w:p></w:body>
      </w:document>`,
    );
    putXmlPart(
      pkg,
      'word/comments.xml',
      `<w:comments xmlns:w="${W_NS}">
        <w:comment w:id="5" w:author="Alice" w:date="2026-04-21T10:00:00Z">
          <w:p><w:r><w:t>need review</w:t></w:r></w:p>
        </w:comment>
      </w:comments>`,
      COMMENTS_CTYPE,
    );

    const { commentIdx } = buildIndices(pkg);
    expect(commentIdx.size).toBe(1);
    const entry = commentIdx.get(5)!;
    expect(entry.comment?.localName).toBe('comment');
    expect(entry.anchors.start).not.toBeNull();
    expect(entry.anchors.end).not.toBeNull();
    expect(entry.anchors.ref).not.toBeNull();
  });

  test('缺 comments.xml:entry.comment=null,anchors 仍在', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}">
        <w:body><w:p>
          <w:commentRangeStart w:id="1"/>
          <w:r><w:t>x</w:t></w:r>
          <w:commentRangeEnd w:id="1"/>
          <w:r><w:commentReference w:id="1"/></w:r>
        </w:p></w:body>
      </w:document>`,
    );
    const { commentIdx } = buildIndices(pkg);
    expect(commentIdx.size).toBe(1);
    expect(commentIdx.get(1)?.comment).toBeNull();
    expect(commentIdx.get(1)?.anchors.start).not.toBeNull();
  });

  test('只有 comments.xml 无锚点:entry 存在,anchors 全 null', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`,
    );
    putXmlPart(
      pkg,
      'word/comments.xml',
      `<w:comments xmlns:w="${W_NS}">
        <w:comment w:id="7"><w:p><w:r><w:t>orphan</w:t></w:r></w:p></w:comment>
      </w:comments>`,
      COMMENTS_CTYPE,
    );
    const { commentIdx } = buildIndices(pkg);
    expect(commentIdx.get(7)?.comment?.localName).toBe('comment');
    expect(commentIdx.get(7)?.anchors.start).toBeNull();
  });
});

describe('buildIndices — 性能基线', () => {
  test('1000 段落 + 200 书签 + 50 批注 应 < 200ms', () => {
    const paragraphs = Array.from(
      { length: 1000 },
      (_, i) =>
        `<w:p w14:paraId="${(i + 1).toString(16).padStart(8, '0').toUpperCase()}"><w:r><w:t>p${i}</w:t></w:r></w:p>`,
    ).join('');
    const bookmarks = Array.from(
      { length: 200 },
      (_, i) =>
        `<w:p><w:bookmarkStart w:id="${i + 1000}" w:name="bm${i}"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="${i + 1000}"/></w:p>`,
    ).join('');
    const anchors = Array.from(
      { length: 50 },
      (_, i) =>
        `<w:p><w:commentRangeStart w:id="${i + 1}"/><w:r><w:t>a</w:t></w:r><w:commentRangeEnd w:id="${i + 1}"/><w:r><w:commentReference w:id="${i + 1}"/></w:r></w:p>`,
    ).join('');
    const commentsBody = Array.from(
      { length: 50 },
      (_, i) => `<w:comment w:id="${i + 1}"><w:p><w:r><w:t>c${i}</w:t></w:r></w:p></w:comment>`,
    ).join('');

    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>${paragraphs}${bookmarks}${anchors}</w:body>
      </w:document>`,
    );
    putXmlPart(
      pkg,
      'word/comments.xml',
      `<w:comments xmlns:w="${W_NS}">${commentsBody}</w:comments>`,
      COMMENTS_CTYPE,
    );

    const start = performance.now();
    const idx = buildIndices(pkg);
    const elapsed = performance.now() - start;

    expect(idx.paraIdx.size).toBe(1000);
    expect(idx.bookmarkIdx.size).toBe(200);
    expect(idx.commentIdx.size).toBe(50);
    expect(elapsed).toBeLessThan(200);
  });
});

describe('buildIndices — 集成', () => {
  test('返回值结构完整,可独立使用各个索引', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body><w:p w14:paraId="AAAA1111"/></w:body></w:document>`,
    );
    const { paraIdx, bookmarkIdx, commentIdx } = buildIndices(pkg);
    expect(paraIdx).toBeInstanceOf(Map);
    expect(bookmarkIdx).toBeInstanceOf(Map);
    expect(commentIdx).toBeInstanceOf(Map);
  });

  test('空包:三索引都是空 Map', () => {
    const { paraIdx, bookmarkIdx, commentIdx } = buildIndices(createPackage());
    expect(paraIdx.size).toBe(0);
    expect(bookmarkIdx.size).toBe(0);
    expect(commentIdx.size).toBe(0);
  });
});
