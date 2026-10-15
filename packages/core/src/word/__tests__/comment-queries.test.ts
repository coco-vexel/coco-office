/**
 * Comment 只读查询纯函数测试(#26)。
 *
 * 覆盖:
 *   - listComments / getComment 命中 + 顺序 + 未命中 null
 *   - rangeKind 三档:paragraph / text / complex(跨段 / 缺锚点)
 *   - 主体缺失(entry.comment === null)被跳过(list 里不出现)
 *   - 锚点全缺(start & end 都 null)→ rangeKind complex + 两个字段都不给
 *   - text 字段:`<w:comment>` 内多段 `<w:p>` 跨行纯文本拼接
 *   - author / date 原样从属性读回,缺则 ''
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { getComment, listComments } from '../comment-queries.js';
import { buildIndices } from '../indices.js';
import { W_NS } from '../namespaces.js';

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

/** 典型 fixture:3 条批注,覆盖三档 rangeKind */
function makeFixture() {
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body>
        <w:p w14:paraId="AAAAAAA1">
          <w:commentRangeStart w:id="1"/>
          <w:r><w:t>entire</w:t></w:r>
          <w:commentRangeEnd w:id="1"/>
          <w:r><w:commentReference w:id="1"/></w:r>
        </w:p>
        <w:p w14:paraId="AAAAAAA2">
          <w:r><w:t>pre </w:t></w:r>
          <w:commentRangeStart w:id="2"/>
          <w:r><w:t>mid</w:t></w:r>
          <w:commentRangeEnd w:id="2"/>
          <w:r><w:t> post</w:t></w:r>
          <w:r><w:commentReference w:id="2"/></w:r>
        </w:p>
        <w:p w14:paraId="AAAAAAA3">
          <w:commentRangeStart w:id="3"/>
          <w:r><w:t>first</w:t></w:r>
        </w:p>
        <w:p w14:paraId="AAAAAAA4">
          <w:r><w:t>second</w:t></w:r>
          <w:commentRangeEnd w:id="3"/>
          <w:r><w:commentReference w:id="3"/></w:r>
        </w:p>
      </w:body>
    </w:document>`,
  );
  putXml(
    pkg,
    'word/comments.xml',
    `<w:comments xmlns:w="${W_NS}">
      <w:comment w:id="1" w:author="alice" w:date="2026-04-22T10:00:00Z">
        <w:p><w:r><w:t>简评 1</w:t></w:r></w:p>
      </w:comment>
      <w:comment w:id="2" w:author="bob" w:date="2026-04-22T11:00:00Z">
        <w:p><w:r><w:t>行 A</w:t></w:r></w:p>
        <w:p><w:r><w:t>行 B</w:t></w:r></w:p>
      </w:comment>
      <w:comment w:id="3" w:author="carol" w:date="2026-04-22T12:00:00Z">
        <w:p><w:r><w:t>跨段评</w:t></w:r></w:p>
      </w:comment>
    </w:comments>`,
    COMMENTS_CTYPE,
  );
  return { pkg, indices: buildIndices(pkg) };
}

describe('listComments / getComment', () => {
  test('listComments:按 Map 插入序返回全部有效批注', () => {
    const { indices } = makeFixture();
    const list = listComments(indices);
    expect(list.map((c) => c.id)).toEqual([1, 2, 3]);
  });

  test("rangeKind='paragraph':整段覆盖,paraId 给、anchorText 不给", () => {
    const { indices } = makeFixture();
    const c = getComment(indices, 1)!;
    expect(c.rangeKind).toBe('paragraph');
    expect(c.paraId).toBe('AAAAAAA1');
    expect(c.anchorText).toBeUndefined();
    expect(c.author).toBe('alice');
    expect(c.date).toBe('2026-04-22T10:00:00Z');
    expect(c.text).toBe('简评 1');
  });

  test("rangeKind='text':段内片段,anchorText = start/end 之间文本", () => {
    const { indices } = makeFixture();
    const c = getComment(indices, 2)!;
    expect(c.rangeKind).toBe('text');
    expect(c.paraId).toBe('AAAAAAA2');
    expect(c.anchorText).toBe('mid');
  });

  test('text 字段:跨多个 <w:p> 的批注正文拼接(无换行分隔)', () => {
    const { indices } = makeFixture();
    const c = getComment(indices, 2)!;
    expect(c.text).toBe('行 A行 B');
  });

  test("rangeKind='complex':跨段落,paraId / anchorText 都不给", () => {
    const { indices } = makeFixture();
    const c = getComment(indices, 3)!;
    expect(c.rangeKind).toBe('complex');
    expect(c.paraId).toBeUndefined();
    expect(c.anchorText).toBeUndefined();
  });

  test('主体缺失(entry.comment = null):listComments 跳过,getComment 返回 null', () => {
    const pkg = createPackage();
    // 只有锚点,没有 comments.xml
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:commentRangeStart w:id="5"/>
            <w:r><w:t>x</w:t></w:r>
            <w:commentRangeEnd w:id="5"/>
            <w:r><w:commentReference w:id="5"/></w:r>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const indices = buildIndices(pkg);
    // 装载期 commentIdx 有 entry(来自锚点)但 comment 主体为 null
    expect(indices.commentIdx.has(5)).toBe(true);
    expect(listComments(indices)).toEqual([]);
    expect(getComment(indices, 5)).toBeNull();
  });

  test('getComment 未命中:返回 null', () => {
    const { indices } = makeFixture();
    expect(getComment(indices, 999)).toBeNull();
  });

  test('缺 w:author / w:date 时降级为空串', () => {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:commentRangeStart w:id="7"/>
            <w:r><w:t>x</w:t></w:r>
            <w:commentRangeEnd w:id="7"/>
            <w:r><w:commentReference w:id="7"/></w:r>
          </w:p>
        </w:body>
      </w:document>`,
    );
    putXml(
      pkg,
      'word/comments.xml',
      `<w:comments xmlns:w="${W_NS}">
        <w:comment w:id="7">
          <w:p><w:r><w:t>no meta</w:t></w:r></w:p>
        </w:comment>
      </w:comments>`,
      COMMENTS_CTYPE,
    );
    const indices = buildIndices(pkg);
    const c = getComment(indices, 7)!;
    expect(c.author).toBe('');
    expect(c.date).toBe('');
    expect(c.text).toBe('no meta');
  });
});
