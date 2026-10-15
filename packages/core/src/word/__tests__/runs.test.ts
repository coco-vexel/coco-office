import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { W_NS } from '../namespaces.js';
import { getParagraphRuns, getParagraphText, getRunText, splitRunAt } from '../runs.js';

/**
 * 从一段 w:p XML 文本构建段落 Element
 */
function parseParagraph(xml: string): Element {
  const dom = parseXmlDocument(`<root xmlns:w="${W_NS}">${xml}</root>`) as unknown as Document;
  const ps = dom.getElementsByTagNameNS(W_NS, 'p');
  const p = ps[0];
  if (!p) throw new Error('no <w:p> found in fixture');
  return p as unknown as Element;
}

function serialize(node: Element): string {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return serializeToWellFormedString(node as any);
}

describe('getParagraphRuns', () => {
  test('0 runs → 空数组', () => {
    const p = parseParagraph('<w:p><w:pPr/></w:p>');
    expect(getParagraphRuns(p)).toEqual([]);
  });

  test('单 run', () => {
    const p = parseParagraph('<w:p><w:r><w:t>hello</w:t></w:r></w:p>');
    const runs = getParagraphRuns(p);
    expect(runs).toHaveLength(1);
    expect(runs[0]?.localName).toBe('r');
  });

  test('多 run 按文档顺序返回', () => {
    const p = parseParagraph(
      '<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r><w:r><w:t>c</w:t></w:r></w:p>',
    );
    const runs = getParagraphRuns(p);
    expect(runs).toHaveLength(3);
    expect(runs.map(getRunText)).toEqual(['a', 'b', 'c']);
  });

  test('不递归:嵌套子元素里的 <w:r> 不计入', () => {
    // <w:sdt> / <w:ins> 等结构化元素里的 run 不应该作为"直接子 run"返回
    const p = parseParagraph(
      '<w:p><w:r><w:t>outer</w:t></w:r><w:ins><w:r><w:t>nested</w:t></w:r></w:ins></w:p>',
    );
    const runs = getParagraphRuns(p);
    expect(runs).toHaveLength(1);
    expect(getRunText(runs[0]!)).toBe('outer');
  });
});

describe('getRunText / getParagraphText', () => {
  test('getRunText 拼接所有 <w:t>', () => {
    const p = parseParagraph(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>foo</w:t><w:t>bar</w:t></w:r></w:p>',
    );
    const runs = getParagraphRuns(p);
    expect(getRunText(runs[0]!)).toBe('foobar');
  });

  test('getRunText 忽略 <w:br> / <w:tab>', () => {
    const p = parseParagraph(
      '<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t><w:tab/><w:t>c</w:t></w:r></w:p>',
    );
    const runs = getParagraphRuns(p);
    expect(getRunText(runs[0]!)).toBe('abc');
  });

  test('getParagraphText 拼接所有 run 文本', () => {
    const p = parseParagraph(
      '<w:p><w:r><w:t>foo</w:t></w:r><w:r><w:t> </w:t></w:r><w:r><w:t>bar</w:t></w:r></w:p>',
    );
    expect(getParagraphText(p)).toBe('foo bar');
  });

  test('空段落文本为空串', () => {
    expect(getParagraphText(parseParagraph('<w:p/>'))).toBe('');
  });
});

describe('splitRunAt', () => {
  test('offset=0 不分裂,返回 after=null', () => {
    const p = parseParagraph('<w:p><w:r><w:t>hello</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    const { before, after } = splitRunAt(run, 0);
    expect(before).toBe(run);
    expect(after).toBeNull();
    expect(getParagraphText(p)).toBe('hello');
    expect(getParagraphRuns(p)).toHaveLength(1);
  });

  test('offset=runLen 不分裂', () => {
    const p = parseParagraph('<w:p><w:r><w:t>hello</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    const { after } = splitRunAt(run, 5);
    expect(after).toBeNull();
    expect(getParagraphText(p)).toBe('hello');
    expect(getParagraphRuns(p)).toHaveLength(1);
  });

  test('中间 offset 分裂为两 run,文本保留,顺序正确', () => {
    const p = parseParagraph('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>hello</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    const { before, after } = splitRunAt(run, 3);
    expect(before).toBe(run);
    expect(after).not.toBeNull();
    const runs = getParagraphRuns(p);
    expect(runs).toHaveLength(2);
    expect(runs[0]).toBe(before);
    expect(runs[1]).toBe(after);
    expect(getRunText(runs[0]!)).toBe('hel');
    expect(getRunText(runs[1]!)).toBe('lo');
    expect(getParagraphText(p)).toBe('hello');
  });

  test('分裂后新 run 保留 <w:rPr>(粗体等格式克隆)', () => {
    const p = parseParagraph(
      '<w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>formatted</w:t></w:r></w:p>',
    );
    const run = getParagraphRuns(p)[0]!;
    const { after } = splitRunAt(run, 4);
    expect(after).not.toBeNull();
    const afterXml = serialize(after!);
    // after run 里应有 <w:rPr> 带 <w:b> 与 <w:sz>
    expect(afterXml).toMatch(/<w:rPr>/);
    expect(afterXml).toMatch(/<w:b\s*\/?>/);
    expect(afterXml).toMatch(/w:sz/);
    // 原 run 的 rPr 不应被破坏
    const beforeXml = serialize(run);
    expect(beforeXml).toMatch(/<w:b\s*\/?>/);
  });

  test('<w:t xml:space="preserve"> 的 xml:space 属性在 after 上保留', () => {
    const p = parseParagraph('<w:p><w:r><w:t xml:space="preserve">  spaced  </w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    const { after } = splitRunAt(run, 4);
    expect(after).not.toBeNull();
    expect(serialize(after!)).toMatch(/xml:space="preserve"/);
  });

  test('offset 超范围抛错', () => {
    const p = parseParagraph('<w:p><w:r><w:t>abc</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(() => splitRunAt(run, -1)).toThrow(/out of range/);
    expect(() => splitRunAt(run, 4)).toThrow(/out of range/);
  });

  test('run 含多个 <w:t> 时抛错(本期不支持)', () => {
    const p = parseParagraph('<w:p><w:r><w:t>a</w:t><w:t>b</w:t></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(() => splitRunAt(run, 1)).toThrow(/single-<w:t>/);
  });

  test('run 无 <w:t>(只有 <w:br> 等)时抛错', () => {
    const p = parseParagraph('<w:p><w:r><w:br/></w:r></w:p>');
    const run = getParagraphRuns(p)[0]!;
    expect(() => splitRunAt(run, 0)).toThrow(/no <w:t> children/);
  });

  test('非 <w:r> 元素抛错', () => {
    const p = parseParagraph('<w:p><w:pPr/></w:p>');
    expect(() => splitRunAt(p, 0)).toThrow(/not a <w:r>/);
  });

  test('多 run 段落中拆某一个 run 后,整段文本仍保持', () => {
    const p = parseParagraph(
      '<w:p><w:r><w:t>foo</w:t></w:r><w:r><w:t>bar</w:t></w:r><w:r><w:t>baz</w:t></w:r></w:p>',
    );
    const middle = getParagraphRuns(p)[1]!;
    splitRunAt(middle, 1);
    expect(getParagraphText(p)).toBe('foobarbaz');
    expect(getParagraphRuns(p)).toHaveLength(4);
  });
});
