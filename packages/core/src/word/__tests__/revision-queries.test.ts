/**
 * Revision 只读查询纯函数测试(#26,Tier 1)。
 *
 * 覆盖:
 *   - w:ins / w:del 基本字段:revId / kind / author / date / story / paraId
 *   - text 字段:insert 用 `<w:t>`、delete 用 `<w:delText>`;空 text 时字段缺
 *   - scope:默认 body;'all' 包含 header / footer;StoryId[] 精确命中
 *   - Tier 1 边界:`<w:rPrChange>` / `<w:pPrChange>` 不被枚举(它们不是 w:ins/w:del
 *     标签,getElementsByTagNameNS 天然过滤)
 *   - 缺 w:id 的 ins/del 被跳过(无法稳定引用)
 *   - paragraph mark 插入(`<w:pPr><w:rPr><w:ins>`)也会被枚举,text 为空时缺字段
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { W14_NS, W_NS } from '../namespaces.js';
import { listRevisions } from '../revision-queries.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const HDR_CTYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';

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

describe('listRevisions — 基本字段', () => {
  test('w:ins:revId / kind / author / date / story / paraId / text 全字段', () => {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:ins w:id="10" w:author="alice" w:date="2026-04-22T10:00:00Z">
              <w:r><w:t>inserted</w:t></w:r>
            </w:ins>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const revs = listRevisions(pkg);
    expect(revs).toHaveLength(1);
    const r = revs[0]!;
    expect(r.revId).toBe('10');
    expect(r.kind).toBe('insert');
    expect(r.author).toBe('alice');
    expect(r.date).toBe('2026-04-22T10:00:00Z');
    expect(r.story).toBe('body');
    expect(r.paraId).toBe('AAAAAAA1');
    expect(r.text).toBe('inserted');
  });

  test('w:del:text 取自 `<w:delText>` 而非 `<w:t>`', () => {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:del w:id="20" w:author="bob" w:date="2026-04-22T11:00:00Z">
              <w:r><w:delText>deleted</w:delText></w:r>
            </w:del>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const r = listRevisions(pkg)[0]!;
    expect(r.kind).toBe('delete');
    expect(r.text).toBe('deleted');
  });

  test('ins / del 混合,且一个段落内多修订,按文档序返回(ins 在前, del 在后)', () => {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins>
            <w:del w:id="2"><w:r><w:delText>y</w:delText></w:r></w:del>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const ids = listRevisions(pkg).map((r) => r.revId);
    // 实现先枚举所有 ins 再枚举所有 del —— 文档序无法严格对齐,但同 kind 内保持
    // getElementsByTagNameNS 的文档序。契约:**调用方应按 revId 自己排序**。
    expect(ids).toEqual(['1', '2']);
  });

  test('缺 w:id 的修订被跳过(无稳定引用)', () => {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:ins w:author="nobody"><w:r><w:t>no-id</w:t></w:r></w:ins>
            <w:ins w:id="5"><w:r><w:t>with-id</w:t></w:r></w:ins>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const revs = listRevisions(pkg);
    expect(revs).toHaveLength(1);
    expect(revs[0]!.revId).toBe('5');
  });

  test('text 为空时:info.text 缺字段', () => {
    // 段落 mark 插入:<w:pPr><w:rPr><w:ins> 没有 <w:t>
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:pPr>
              <w:rPr>
                <w:ins w:id="30" w:author="x" w:date="2026-01-01T00:00:00Z"/>
              </w:rPr>
            </w:pPr>
            <w:r><w:t>after</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const revs = listRevisions(pkg);
    expect(revs).toHaveLength(1);
    const r = revs[0]!;
    expect(r.revId).toBe('30');
    expect(r.kind).toBe('insert');
    expect(Object.prototype.hasOwnProperty.call(r, 'text')).toBe(false);
    // paraId 仍能通过祖先链反查到
    expect(r.paraId).toBe('AAAAAAA1');
  });
});

describe('listRevisions — Tier 2 属性修订守护(#33)', () => {
  test('扫到 <w:rPrChange> → 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:r>
              <w:rPr>
                <w:rPrChange w:id="101" w:author="x" w:date="2026-01-01T00:00:00Z">
                  <w:rPr/>
                </w:rPrChange>
              </w:rPr>
              <w:t>x</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`,
    );
    expect(() => listRevisions(pkg)).toThrow(/rPrChange/);
  });

  test('扫到 <w:pPrChange> → 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:pPr>
              <w:pPrChange w:id="100" w:author="x" w:date="2026-01-01T00:00:00Z">
                <w:pPr/>
              </w:pPrChange>
            </w:pPr>
            <w:r><w:t>x</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`,
    );
    expect(() => listRevisions(pkg)).toThrow(/pPrChange/);
  });

  test('扫到 <w:sectPrChange> → 抛', () => {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:sectPr>
            <w:sectPrChange w:id="1" w:author="x" w:date="2026-01-01T00:00:00Z">
              <w:sectPr/>
            </w:sectPrChange>
          </w:sectPr>
        </w:body>
      </w:document>`,
    );
    expect(() => listRevisions(pkg)).toThrow(/sectPrChange/);
  });
});

describe('listRevisions — scope', () => {
  function makeMultiStoryFixture(): Package {
    const pkg = createPackage();
    putXml(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1">
            <w:ins w:id="1"><w:r><w:t>body-ins</w:t></w:r></w:ins>
          </w:p>
        </w:body>
      </w:document>`,
    );
    putXml(
      pkg,
      'word/header1.xml',
      `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:p w14:paraId="HHHHHHHH">
          <w:ins w:id="2"><w:r><w:t>header-ins</w:t></w:r></w:ins>
        </w:p>
      </w:hdr>`,
      HDR_CTYPE,
    );
    return pkg;
  }

  test('默认 scope=body:仅返回 body 修订', () => {
    const pkg = makeMultiStoryFixture();
    const revs = listRevisions(pkg);
    expect(revs.map((r) => r.revId)).toEqual(['1']);
    expect(revs[0]!.story).toBe('body');
  });

  test("scope='all':含 header", () => {
    const pkg = makeMultiStoryFixture();
    const revs = listRevisions(pkg, { scope: 'all' });
    expect(revs.map((r) => r.revId).sort()).toEqual(['1', '2']);
    const hdr = revs.find((r) => r.revId === '2')!;
    expect(hdr.story).toBe('header:word/header1.xml');
  });

  test('scope=StoryId[]:精确命中 header,body 不返回', () => {
    const pkg = makeMultiStoryFixture();
    const revs = listRevisions(pkg, { scope: ['header:word/header1.xml'] });
    expect(revs.map((r) => r.revId)).toEqual(['2']);
  });
});
