/**
 * listParagraphs — hasMath 字段检测测试（task-formula-1-core）
 *
 * 覆盖 spec §4 测试矩阵：
 *   1. 块级公式段（oMathPara > oMath）→ hasMath === true
 *   2. 内联公式段（段内裸 oMath）→ hasMath === true
 *   3. 纯文字"公式"段（无 oMath）→ hasMath 不存在
 *   4. 普通正文段 → hasMath 不存在
 *   5. 不回归（不含 oMath 文档，所有段无 hasMath）
 *   6. oMath 嵌在 run 内（getElementsByTagNameNS 递归覆盖）→ hasMath === true
 *
 * Fixture 采用内联 XML 构造（parseXmlDocument + createXmlPart），
 * document.xml 根节点加 xmlns:m 声明以模拟真实 Word OMML 文档。
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { listParagraphs } from '../list-paragraphs.js';
import { MATH_NS, W14_NS, W_NS } from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

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

/**
 * 构造含 OMML 公式的 Package fixture。
 * - 段 A (MATHPARA): 块级公式 <m:oMathPara><m:oMath>…</m:oMath></m:oMathPara>
 * - 段 B (MATHINLN): 内联裸 <m:oMath>…</m:oMath>（直属 w:p）
 * - 段 C (PLAINTXT): 纯文字 "F = ma (1)"，无 oMath
 * - 段 D (NORMALPA): 普通正文段，无公式
 */
function makeOmmlFixture(): Package {
  const pkg = createPackage();
  // math namespace: http://schemas.openxmlformats.org/officeDocument/2006/math
  const xml = `<w:document
    xmlns:w="${W_NS}"
    xmlns:w14="${W14_NS}"
    xmlns:m="${MATH_NS}">
    <w:body>
      <!-- 段 A: 块级公式（oMathPara 包裹 oMath） -->
      <w:p w14:paraId="MATHPARA">
        <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
        <m:oMathPara>
          <m:oMath>
            <m:r><m:t>a + b = c</m:t></m:r>
          </m:oMath>
        </m:oMathPara>
      </w:p>
      <!-- 段 B: 内联裸 oMath（直接在 w:p 下，不经 oMathPara） -->
      <w:p w14:paraId="MATHINLN">
        <w:r><w:t>see </w:t></w:r>
        <m:oMath>
          <m:r><m:t>E = mc^2</m:t></m:r>
        </m:oMath>
        <w:r><w:t> above</w:t></w:r>
      </w:p>
      <!-- 段 C: 纯文字公式，无 oMath -->
      <w:p w14:paraId="PLAINTXT">
        <w:r><w:t>F = ma (1)</w:t></w:r>
      </w:p>
      <!-- 段 D: 普通正文段 -->
      <w:p w14:paraId="NORMALPA">
        <w:r><w:t>Hello World</w:t></w:r>
      </w:p>
    </w:body>
  </w:document>`;
  putXmlPart(pkg, 'word/document.xml', xml);
  return pkg;
}

/**
 * 构造 oMath 嵌在 run 内的 fixture（用于测试 case 6 的递归穿透）。
 * Word 有时会把 oMath 放在 w:r 子树里。
 */
function makeOmmlInRunFixture(): Package {
  const pkg = createPackage();
  const xml = `<w:document
    xmlns:w="${W_NS}"
    xmlns:w14="${W14_NS}"
    xmlns:m="${MATH_NS}">
    <w:body>
      <w:p w14:paraId="MATHRUN1">
        <w:r>
          <m:oMath>
            <m:r><m:t>x^2 + y^2</m:t></m:r>
          </m:oMath>
        </w:r>
      </w:p>
      <w:p w14:paraId="CLEANPA1">
        <w:r><w:t>no math here</w:t></w:r>
      </w:p>
    </w:body>
  </w:document>`;
  putXmlPart(pkg, 'word/document.xml', xml);
  return pkg;
}

// ─────────── 测试矩阵 ───────────

describe('listParagraphs — hasMath 字段', () => {
  test('Case 1: 块级公式段（oMathPara > oMath）→ hasMath === true', () => {
    const pkg = makeOmmlFixture();
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    const paraA = r.paragraphs.find((p) => p.paraId === 'MATHPARA');
    expect(paraA).toBeDefined();
    expect(paraA!.hasMath).toBe(true);
  });

  test('Case 2: 内联裸 oMath 段（oMath 直属 w:p）→ hasMath === true', () => {
    const pkg = makeOmmlFixture();
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    const paraB = r.paragraphs.find((p) => p.paraId === 'MATHINLN');
    expect(paraB).toBeDefined();
    expect(paraB!.hasMath).toBe(true);
  });

  test('Case 3: 纯文字"公式"段（无 oMath）→ hasMath 不存在（undefined）', () => {
    const pkg = makeOmmlFixture();
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    const paraC = r.paragraphs.find((p) => p.paraId === 'PLAINTXT');
    expect(paraC).toBeDefined();
    expect(paraC!.hasMath).toBeUndefined();
    expect(Object.prototype.hasOwnProperty.call(paraC, 'hasMath')).toBe(false);
  });

  test('Case 4: 普通正文段 → hasMath 不存在（undefined）', () => {
    const pkg = makeOmmlFixture();
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    const paraD = r.paragraphs.find((p) => p.paraId === 'NORMALPA');
    expect(paraD).toBeDefined();
    expect(paraD!.hasMath).toBeUndefined();
    expect(Object.prototype.hasOwnProperty.call(paraD, 'hasMath')).toBe(false);
  });

  test('Case 5: 不回归——不含 oMath 的文档，所有段无 hasMath 字段', () => {
    // 用现有 list-paragraphs.test.ts 风格的纯文字 pkg
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello</w:t></w:r></w:p>
          <w:p w14:paraId="AAAAAAA2"><w:r><w:t>world</w:t></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    for (const p of r.paragraphs) {
      expect(p.hasMath).toBeUndefined();
      expect(Object.prototype.hasOwnProperty.call(p, 'hasMath')).toBe(false);
    }
  });

  test('Case 6: oMath 嵌在 w:r 内（递归穿透）→ hasMath === true', () => {
    const pkg = makeOmmlInRunFixture();
    const r = listParagraphs(pkg, { pagination: { mode: 'all' } });
    const mathPara = r.paragraphs.find((p) => p.paraId === 'MATHRUN1');
    const cleanPara = r.paragraphs.find((p) => p.paraId === 'CLEANPA1');
    expect(mathPara).toBeDefined();
    expect(mathPara!.hasMath).toBe(true);
    expect(cleanPara).toBeDefined();
    expect(cleanPara!.hasMath).toBeUndefined();
  });
});

// ─────────── MATH_NS 常量导出验证 ───────────

describe('MATH_NS 常量', () => {
  test('值与 OOXML 规范一致', () => {
    expect(MATH_NS).toBe('http://schemas.openxmlformats.org/officeDocument/2006/math');
  });
});
