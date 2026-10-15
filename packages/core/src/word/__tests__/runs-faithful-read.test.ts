/**
 * 保真读测试（task-20260623-core-tab-faithful-read）。
 *
 * 验证：
 *  1. 对称 round-trip：getRunText / getParagraphText 保真读与 appendTextAsRuns 写侧逆运算对称
 *  2. 默认偏移不变：不传 opts 时行为字节级不变，tab/br 都算 0 长
 *  3. 保真模式用例：各种位置的 tab/br
 *  4. drawings.innerText 保真读
 *  5. 铁律确认：偏移算术调用方调用链不受影响
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { W_NS, W14_NS, WP_NS, WPS_NS } from '../namespaces.js';
import { getParagraphRuns, getParagraphText, getRunText } from '../runs.js';
import { appendTextAsRuns, appendSegmentsToRun } from '../plan-executors/_shared.js';
import { listDrawings } from '../list-drawings.js';
import { buildIndices } from '../indices.js';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';

// ─────────── 辅助函数 ───────────

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function parseParagraph(xml: string): Element {
  const dom = parseXmlDocument(`<root xmlns:w="${W_NS}">${xml}</root>`) as unknown as Document;
  const ps = dom.getElementsByTagNameNS(W_NS, 'p');
  const p = ps[0];
  if (!p) throw new Error('no <w:p> found in fixture');
  return p as unknown as Element;
}

/** 在同一个 DOM 里创建段落并用 appendTextAsRuns 写 text，返回段落 Element */
function makeParaWithText(text: string): Element {
  const dom = parseXmlDocument(
    `<root xmlns:w="${W_NS}" xmlns:xml="http://www.w3.org/XML/1998/namespace"><w:p/></root>`,
  ) as unknown as Document;
  const ps = dom.getElementsByTagNameNS(W_NS, 'p');
  const p = ps[0] as unknown as Element;
  appendTextAsRuns(p, text);
  return p;
}

function _putXmlPart(pkg: Package, path: string, xml: string, contentType = DOC_CTYPE): void {
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

// ─────────── 1. 默认偏移不变 ───────────

describe('默认行为（偏移安全，字节级不变）', () => {
  test('getRunText 默认忽略 <w:tab/>', () => {
    const p = parseParagraph('<w:p><w:r><w:t>甲</w:t><w:tab/><w:t>乙</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run)).toBe('甲乙');
  });

  test('getRunText 默认忽略 <w:br/>（textWrapping）', () => {
    const p = parseParagraph('<w:p><w:r><w:t>甲</w:t><w:br/><w:t>乙</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run)).toBe('甲乙');
  });

  test('getParagraphText 默认忽略 tab / br', () => {
    const p = parseParagraph('<w:p><w:r><w:t>A</w:t><w:tab/></w:r><w:r><w:t>B</w:t></w:r></w:p>');
    expect(getParagraphText(p)).toBe('AB');
  });

  test('显式 structuralWhitespace:false 与不传 opts 效果相同', () => {
    const p = parseParagraph('<w:p><w:r><w:t>X</w:t><w:tab/><w:t>Y</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: false })).toBe('XY');
    expect(getParagraphText(p, { structuralWhitespace: false })).toBe('XY');
  });

  test('tab 在段首时默认长度 = 0（偏移不飘）', () => {
    const p = parseParagraph('<w:p><w:r><w:tab/><w:t>hello</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    // 默认模式：tab 不计入长度
    expect(getRunText(run).length).toBe(5);
  });
});

// ─────────── 2. 保真读单元 ───────────

describe('保真读：getRunText structuralWhitespace:true', () => {
  test('甲<w:tab/>乙 → 甲\\t乙', () => {
    const p = parseParagraph('<w:p><w:r><w:t>甲</w:t><w:tab/><w:t>乙</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('甲\t乙');
  });

  test('甲<w:br/>乙（无 type）→ 甲\\n乙', () => {
    const p = parseParagraph('<w:p><w:r><w:t>甲</w:t><w:br/><w:t>乙</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('甲\n乙');
  });

  test('<w:br w:type="textWrapping"/> → \\n', () => {
    const p = parseParagraph(
      `<w:p><w:r><w:t>A</w:t><w:br w:type="textWrapping"/><w:t>B</w:t></w:r></w:p>`,
    );
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('A\nB');
  });

  test('<w:br w:type="page"/> 不还原（非写侧产出）', () => {
    const p = parseParagraph(`<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>`);
    const run = getParagraphRuns(p)[0]!;
    // page break 不在写侧逆运算内，忽略
    expect(getRunText(run, { structuralWhitespace: true })).toBe('AB');
  });

  test('<w:br w:type="column"/> 不还原', () => {
    const p = parseParagraph(
      `<w:p><w:r><w:t>X</w:t><w:br w:type="column"/><w:t>Y</w:t></w:r></w:p>`,
    );
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('XY');
  });

  test('tab 在段首', () => {
    const p = parseParagraph('<w:p><w:r><w:tab/><w:t>hello</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('\thello');
  });

  test('tab 在段尾', () => {
    const p = parseParagraph('<w:p><w:r><w:t>hello</w:t><w:tab/></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('hello\t');
  });

  test('多个 tab 连续', () => {
    const p = parseParagraph('<w:p><w:r><w:tab/><w:tab/><w:t>x</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('\t\tx');
  });

  test('tab 与 br 混合', () => {
    const p = parseParagraph(
      '<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p>',
    );
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('A\tB\nC');
  });

  test('纯 tab run（无 w:t）', () => {
    const p = parseParagraph('<w:p><w:r><w:tab/></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(getRunText(run, { structuralWhitespace: true })).toBe('\t');
  });
});

// ─────────── 3. getParagraphText 保真读 ───────────

describe('保真读：getParagraphText structuralWhitespace:true', () => {
  test('跨 run 的 tab（分属不同 run）', () => {
    // 两个 run，第一个末尾有 tab（实际 appendSegmentsToRun 会把 tab 独立放进 run）
    const p = parseParagraph(
      '<w:p>' +
        '<w:r><w:t>机关名称</w:t><w:tab/></w:r>' +
        '<w:r><w:t>2026-06-23</w:t></w:r>' +
        '</w:p>',
    );
    expect(getParagraphText(p, { structuralWhitespace: true })).toBe('机关名称\t2026-06-23');
  });

  test('空段落保真读也是空串', () => {
    const p = parseParagraph('<w:p/>');
    expect(getParagraphText(p, { structuralWhitespace: true })).toBe('');
  });

  test('纯文字段落保真读与默认相同', () => {
    const p = parseParagraph('<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>');
    expect(getParagraphText(p, { structuralWhitespace: true })).toBe('Hello World');
    expect(getParagraphText(p)).toBe('Hello World');
  });
});

// ─────────── 4. 对称 round-trip ───────────

describe('对称 round-trip：appendSegmentsToRun 写 → 保真读 = 原文', () => {
  test('含 \\t 的文本 round-trip', () => {
    const original = '机关名称\t2026-06-23';
    const p = makeParaWithText(original);
    const result = getParagraphText(p, { structuralWhitespace: true });
    expect(result).toBe(original);
  });

  test('含 \\n 的文本 round-trip', () => {
    const original = '第一行\n第二行';
    const p = makeParaWithText(original);
    const result = getParagraphText(p, { structuralWhitespace: true });
    expect(result).toBe(original);
  });

  test('含多个 \\t 和 \\n 的文本 round-trip', () => {
    const original = 'A\tB\nC\t\tD\n';
    const p = makeParaWithText(original);
    const result = getParagraphText(p, { structuralWhitespace: true });
    expect(result).toBe(original);
  });

  test('\\t 在开头和结尾 round-trip', () => {
    const original = '\t首段\t';
    const p = makeParaWithText(original);
    const result = getParagraphText(p, { structuralWhitespace: true });
    expect(result).toBe(original);
  });

  test('默认读（不保真）round-trip 后 tab/br 丢失（验证现状，lock 住已知行为）', () => {
    const original = '甲\t乙';
    const p = makeParaWithText(original);
    // 默认模式：tab 丢失
    const defaultResult = getParagraphText(p);
    expect(defaultResult).toBe('甲乙');
    // 保真模式：正确还原
    const faithfulResult = getParagraphText(p, { structuralWhitespace: true });
    expect(faithfulResult).toBe('甲\t乙');
  });

  test('appendSegmentsToRun 写入后再保真读逐节点验证', () => {
    // 构造 DOM + run，直接用 appendSegmentsToRun 填入
    const dom = parseXmlDocument(
      `<root xmlns:w="${W_NS}" xmlns:xml="http://www.w3.org/XML/1998/namespace"><w:r/></root>`,
    ) as unknown as Document;
    const rs = dom.getElementsByTagNameNS(W_NS, 'r');
    const run = rs[0] as unknown as Element;
    appendSegmentsToRun(dom as unknown as Document, run, 'A\tB\nC');
    // 验证 DOM 结构：应该有 w:t(A)、w:tab、w:t(B)、w:br、w:t(C)
    const children = Array.from(
      { length: run.childNodes.length },
      (_, i) => run.childNodes[i] as Element,
    );
    const locals = children.map((c) => c.localName);
    expect(locals).toEqual(['t', 'tab', 't', 'br', 't']);
    // 保真读还原
    expect(getRunText(run, { structuralWhitespace: true })).toBe('A\tB\nC');
    // 默认读丢 tab/br
    expect(getRunText(run)).toBe('ABC');
  });
});

// ─────────── 5. drawings.innerText 保真读 ───────────

describe('drawings.innerText 保真读', () => {
  const NS_DECL =
    ` xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:wp="${WP_NS}" xmlns:wps="${WPS_NS}"` +
    ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
  const _WPC_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas';
  const _WPG_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

  function makeTextBoxDoc(innerXml: string): Package {
    // TextBox drawing 包含 w:txbxContent
    const docXml = `<w:document${NS_DECL}>
  <w:body>
    <w:p w14:paraId="AAAABBBB">
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="685800"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="1" name="TextBox1"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>
                  <wps:spPr>
                    <a:prstGeom prst="rect"/>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>${innerXml}</w:txbxContent>
                  </wps:txbx>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
    const pkg = createPackage();
    setPart(
      pkg,
      createXmlPart({
        path: 'word/document.xml',
        contentType: DOC_CTYPE,
        xmlDom: parseXmlDocument(docXml) as unknown as Document,
        originalBytes: new TextEncoder().encode(docXml),
      }),
    );
    return pkg;
  }

  test('默认模式：innerText 只拼 w:t，不含 \\t', () => {
    // 文本框内：机关名\t日期
    const innerXml = `<w:p><w:r><w:t xml:space="preserve">机关名</w:t><w:tab/><w:t>日期</w:t></w:r></w:p>`;
    const pkg = makeTextBoxDoc(innerXml);
    const indices = buildIndices(pkg);
    const drawings = listDrawings(pkg, indices);
    // 默认：tab 丢失
    expect(drawings[0]?.innerText).toBe('机关名日期');
  });

  test('保真模式：innerText 含 \\t', () => {
    const innerXml = `<w:p><w:r><w:t xml:space="preserve">机关名</w:t><w:tab/><w:t>日期</w:t></w:r></w:p>`;
    const pkg = makeTextBoxDoc(innerXml);
    const indices = buildIndices(pkg);
    const drawings = listDrawings(pkg, indices, { structuralWhitespace: true });
    expect(drawings[0]?.innerText).toBe('机关名\t日期');
  });

  test('保真模式：innerText 含 \\n（软换行）', () => {
    const innerXml = `<w:p><w:r><w:t>第一行</w:t><w:br/><w:t>第二行</w:t></w:r></w:p>`;
    const pkg = makeTextBoxDoc(innerXml);
    const indices = buildIndices(pkg);
    const drawings = listDrawings(pkg, indices, { structuralWhitespace: true });
    expect(drawings[0]?.innerText).toBe('第一行\n第二行');
  });

  test('保真模式：page break 不计入 innerText', () => {
    const innerXml = `<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>`;
    const pkg = makeTextBoxDoc(innerXml);
    const indices = buildIndices(pkg);
    const drawings = listDrawings(pkg, indices, { structuralWhitespace: true });
    expect(drawings[0]?.innerText).toBe('AB');
  });
});
