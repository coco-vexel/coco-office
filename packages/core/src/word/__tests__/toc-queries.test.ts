/**
 * listTocFields 测试（#197）。
 *
 * 覆盖：
 *   - simple field TOC（w:fldSimple）
 *   - complex field TOC（begin/separate/end 三 run 跨多段）
 *   - complex field TOC 嵌套 PAGEREF / HYPERLINK 子域
 *   - 多个 TOC 域并存
 *   - 无 TOC 域的文档 → 返回空数组
 *   - 非 TOC 域（PAGE / DATE）→ 不被识别
 *   - 节边界跨越（TOC 跨节也能闭合）
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { listTocFields } from '../toc-queries.js';
import { W_NS, W14_NS } from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function putXml(pkg: Package, path: string, xml: string, ct: string): void {
  setPart(
    pkg,
    createXmlPart({
      path,
      contentType: ct,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
      originalBytes: new TextEncoder().encode(xml),
    }),
  );
}

function pid(i: number): string {
  return String(i).padStart(8, '0').toUpperCase();
}

function makeBody(bodyInner: string): Package {
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`,
    DOC_CTYPE,
  );
  return pkg;
}

// 构建 complex field TOC 的段序列
// p1: begin 段, p2...pN-1: 目录项段, pN: end 段
function complexTocParas(
  firstId: number,
  contentParaIds: number[],
  instrText: string = ' TOC \\o "1-3" \\h \\z \\u ',
): string {
  const beginPara = `<w:p w14:paraId="${pid(firstId)}">
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText>${instrText}</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  </w:p>`;

  const contentParas = contentParaIds
    .map((id) => `<w:p w14:paraId="${pid(id)}"><w:r><w:t>目录项</w:t></w:r></w:p>`)
    .join('');

  const lastId =
    contentParaIds.length > 0 ? contentParaIds[contentParaIds.length - 1]! + 1 : firstId + 1;
  const endPara = `<w:p w14:paraId="${pid(lastId)}">
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>`;

  return beginPara + contentParas + endPara;
}

describe('listTocFields — 边界', () => {
  test('无 body → []', () => {
    const pkg = createPackage();
    expect(listTocFields(pkg)).toEqual([]);
  });

  test('body 无 TOC 域 → []', () => {
    const pkg = makeBody(`<w:p w14:paraId="${pid(1)}"><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr/>`);
    expect(listTocFields(pkg)).toEqual([]);
  });
});

describe('listTocFields — simple field', () => {
  test('w:fldSimple TOC → 识别，firstParaId=lastParaId', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}">
        <w:fldSimple w:instr=' TOC \\o "1-3" '>
          <w:r><w:t>目录内容</w:t></w:r>
        </w:fldSimple>
      </w:p>`,
    );
    const out = listTocFields(pkg);
    expect(out).toHaveLength(1);
    expect(out[0]?.tocIndex).toBe(0);
    expect(out[0]?.firstParaId).toBe(pid(1));
    expect(out[0]?.lastParaId).toBe(pid(1));
    expect(out[0]?.instr).toMatch(/^TOC/i);
  });

  test('w:fldSimple 非 TOC 域（PAGE）→ 不识别', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}">
        <w:fldSimple w:instr=" PAGE ">
          <w:r><w:t>1</w:t></w:r>
        </w:fldSimple>
      </w:p>`,
    );
    expect(listTocFields(pkg)).toEqual([]);
  });
});

describe('listTocFields — complex field', () => {
  test('begin/separate/end 三段 complex TOC → 识别', () => {
    const firstId = 10;
    const contentIds = [11, 12, 13];
    const lastId = 14; // end 段
    const pkg = makeBody(complexTocParas(firstId, contentIds));
    const out = listTocFields(pkg);
    expect(out).toHaveLength(1);
    expect(out[0]?.tocIndex).toBe(0);
    expect(out[0]?.firstParaId).toBe(pid(firstId));
    expect(out[0]?.lastParaId).toBe(pid(lastId));
    expect(out[0]?.instr).toMatch(/^TOC/i);
  });

  test('非 TOC complex field（DATE）→ 不识别', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> DATE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      </w:p>
      <w:p w14:paraId="${pid(2)}">
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`,
    );
    expect(listTocFields(pkg)).toEqual([]);
  });

  test('复杂 TOC 跨多段（10 个目录项）→ firstParaId / lastParaId 正确', () => {
    const ids = Array.from({ length: 12 }, (_, i) => i + 1);
    const firstId = ids[0]!;
    const contentIds = ids.slice(1, 11);
    const endId = ids[11]!;
    const pkg = makeBody(complexTocParas(firstId, contentIds));
    const out = listTocFields(pkg);
    expect(out).toHaveLength(1);
    expect(out[0]?.firstParaId).toBe(pid(firstId));
    expect(out[0]?.lastParaId).toBe(pid(endId));
  });
});

describe('listTocFields — 嵌套子域', () => {
  test('TOC 内含 PAGEREF 嵌套子域 → 只返回外层 TOC，instr 不含子域 instrText', () => {
    // TOC 域内包含 PAGEREF 子域
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      </w:p>
      <w:p w14:paraId="${pid(2)}">
        <w:r><w:t>第一章</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> PAGEREF _Toc001 \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>5</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
      <w:p w14:paraId="${pid(3)}">
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`,
    );
    const out = listTocFields(pkg);
    expect(out).toHaveLength(1);
    expect(out[0]?.firstParaId).toBe(pid(1));
    expect(out[0]?.lastParaId).toBe(pid(3));
    // instr 不含子域 PAGEREF 的文本
    expect(out[0]?.instr).toMatch(/^TOC/i);
    expect(out[0]?.instr).not.toMatch(/PAGEREF/i);
  });

  test('TOC 内含 HYPERLINK 嵌套子域 → depth 正确，只返回外层 TOC', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> TOC \\h </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      </w:p>
      <w:p w14:paraId="${pid(2)}">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> HYPERLINK \\l "_Toc001" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>第一章</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
      <w:p w14:paraId="${pid(3)}">
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`,
    );
    const out = listTocFields(pkg);
    expect(out).toHaveLength(1);
    expect(out[0]?.instr).not.toMatch(/HYPERLINK/i);
  });
});

describe('listTocFields — 多个 TOC 域', () => {
  test('两个 TOC 域并存 → tocIndex 0 和 1', () => {
    // 第一个: simple field
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}">
        <w:fldSimple w:instr=' TOC \\o "1-2" '>
          <w:r><w:t>目录A</w:t></w:r>
        </w:fldSimple>
      </w:p>
      <w:p w14:paraId="${pid(2)}"><w:r><w:t>正文</w:t></w:r></w:p>
      ${complexTocParas(3, [4, 5])}`,
    );
    const out = listTocFields(pkg);
    expect(out).toHaveLength(2);
    expect(out[0]?.tocIndex).toBe(0);
    expect(out[0]?.firstParaId).toBe(pid(1));
    expect(out[1]?.tocIndex).toBe(1);
    expect(out[1]?.firstParaId).toBe(pid(3));
  });
});

describe('listTocFields — 节边界跨越', () => {
  test('TOC 域跨越 sectPr 分节边界 → 依然正确闭合', () => {
    // TOC begin 在第一节，end 在第二节
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      </w:p>
      <w:p w14:paraId="${pid(2)}"><w:pPr><w:sectPr/></w:pPr><w:r><w:t>项1</w:t></w:r></w:p>
      <w:p w14:paraId="${pid(3)}">
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
      <w:sectPr/>`,
    );
    const out = listTocFields(pkg);
    expect(out).toHaveLength(1);
    expect(out[0]?.firstParaId).toBe(pid(1));
    expect(out[0]?.lastParaId).toBe(pid(3));
  });
});
