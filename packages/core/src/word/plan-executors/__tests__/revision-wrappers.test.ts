/**
 * revision-wrappers helper 单测(#31)。
 *
 * 覆盖:
 *   - allocateRevisionId:空 pkg / 已有 ins/del / 非数字 id 忽略
 *   - wrapRunsAsInsert:属性完整、w:t 保留、run 数量可变
 *   - wrapRunsAsDelete:w:t → w:delText 转换,xml:space 属性保留
 *   - ensureParaMarkInsert / ensureParaMarkDelete:pPr/rPr 按需建,ins/del 放在 rPr 内
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { W_NS, XML_NS } from '../../namespaces.js';
import {
  allocateRevisionId,
  ensureParaMarkDelete,
  ensureParaMarkInsert,
  wrapRunsAsDelete,
  wrapRunsAsInsert,
} from '../revision-wrappers.js';

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

function makePkg(bodyInner: string): Package {
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}"><w:body>${bodyInner}</w:body></w:document>`,
  );
  return pkg;
}

function makeParagraph(pkg: Package): Element {
  const doc = pkg.parts.get('word/document.xml')!.xmlDom as unknown as Document;
  return doc.getElementsByTagNameNS(W_NS, 'p')[0]!;
}

const attrs = { author: 'alice', date: new Date('2026-04-23T00:00:00.000Z') };

// ─────────── allocateRevisionId ───────────

describe('allocateRevisionId', () => {
  test('空 pkg → 0', () => {
    const pkg = makePkg('<w:p></w:p>');
    expect(allocateRevisionId(pkg)).toBe(0);
  });

  test('已有 ins w:id=5 → 6', () => {
    const pkg = makePkg(
      `<w:p><w:ins w:id="5" w:author="x" w:date="2026-01-01T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins></w:p>`,
    );
    expect(allocateRevisionId(pkg)).toBe(6);
  });

  test('ins=5 + del=7 → 8', () => {
    const pkg = makePkg(
      `<w:p>
        <w:ins w:id="5" w:author="x" w:date="2026-01-01T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
        <w:del w:id="7" w:author="x" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>b</w:delText></w:r></w:del>
      </w:p>`,
    );
    expect(allocateRevisionId(pkg)).toBe(8);
  });

  test('非数字 id 被忽略', () => {
    const pkg = makePkg(
      `<w:p><w:ins w:id="notanumber" w:author="x" w:date="2026-01-01T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins></w:p>`,
    );
    expect(allocateRevisionId(pkg)).toBe(0);
  });
});

// ─────────── wrapRunsAsInsert ───────────

describe('wrapRunsAsInsert', () => {
  test('单 run:包成 <w:ins>,属性完整,w:t 保留', () => {
    const pkg = makePkg(`<w:p><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p>`);
    const paragraph = makeParagraph(pkg);
    const run = paragraph.getElementsByTagNameNS(W_NS, 'r')[0]!;

    const ins = wrapRunsAsInsert([run], 3, attrs);
    expect(ins.namespaceURI).toBe(W_NS);
    expect(ins.localName).toBe('ins');
    expect(ins.getAttributeNS(W_NS, 'id') ?? ins.getAttribute('w:id')).toBe('3');
    expect(ins.getAttributeNS(W_NS, 'author') ?? ins.getAttribute('w:author')).toBe('alice');
    expect(ins.getAttributeNS(W_NS, 'date') ?? ins.getAttribute('w:date')).toBe(
      '2026-04-23T00:00:00.000Z',
    );
    // 内部 run 保留 w:t(不是 delText)
    const t = ins.getElementsByTagNameNS(W_NS, 't')[0]!;
    expect(t.textContent).toBe('hello');
  });

  test('多 run:包在同一 ins,所有 run 顺序保留', () => {
    const pkg = makePkg(`<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>`);
    const paragraph = makeParagraph(pkg);
    const runs = Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'));
    const ins = wrapRunsAsInsert(runs, 0, attrs);
    expect(ins.getElementsByTagNameNS(W_NS, 'r').length).toBe(2);
    expect(ins.getElementsByTagNameNS(W_NS, 't').length).toBe(2);
  });
});

// ─────────── wrapRunsAsDelete ───────────

describe('wrapRunsAsDelete', () => {
  test('单 run:w:t 转 w:delText,xml:space 属性保留', () => {
    const pkg = makePkg(`<w:p><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p>`);
    const paragraph = makeParagraph(pkg);
    const run = paragraph.getElementsByTagNameNS(W_NS, 'r')[0]!;

    const del = wrapRunsAsDelete([run], 5, attrs);
    expect(del.localName).toBe('del');
    // del 内部不应再有 w:t
    expect(del.getElementsByTagNameNS(W_NS, 't').length).toBe(0);
    // 应有 w:delText
    const delText = del.getElementsByTagNameNS(W_NS, 'delText')[0]!;
    expect(delText).toBeDefined();
    expect(delText.textContent).toBe('hello');
    // xml:space 属性搬过来
    expect(delText.getAttributeNS(XML_NS, 'space')).toBe('preserve');
  });

  test('多 run,多 w:t → 都被转为 delText', () => {
    const pkg = makePkg(
      `<w:p>
        <w:r><w:t>a</w:t></w:r>
        <w:r><w:t>b</w:t></w:r>
      </w:p>`,
    );
    const paragraph = makeParagraph(pkg);
    const runs = Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'));
    const del = wrapRunsAsDelete(runs, 1, attrs);
    const delTexts = del.getElementsByTagNameNS(W_NS, 'delText');
    expect(delTexts.length).toBe(2);
    expect(del.getElementsByTagNameNS(W_NS, 't').length).toBe(0);
  });

  test('原 run 不变(deep clone 副本才被改)', () => {
    const pkg = makePkg(`<w:p><w:r><w:t>a</w:t></w:r></w:p>`);
    const paragraph = makeParagraph(pkg);
    const run = paragraph.getElementsByTagNameNS(W_NS, 'r')[0]!;
    wrapRunsAsDelete([run], 0, attrs);
    // 原 run 仍有 w:t
    expect(run.getElementsByTagNameNS(W_NS, 't').length).toBe(1);
  });
});

// ─────────── ensureParaMarkInsert / ensureParaMarkDelete ───────────

describe('ensureParaMarkInsert / Delete', () => {
  test('ensureParaMarkInsert:pPr 不存在 → 建 pPr + rPr + ins', () => {
    const pkg = makePkg(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`);
    const paragraph = makeParagraph(pkg);
    ensureParaMarkInsert(paragraph, 0, attrs);
    const pPr = paragraph.getElementsByTagNameNS(W_NS, 'pPr')[0]!;
    expect(pPr).toBeDefined();
    const rPr = pPr.getElementsByTagNameNS(W_NS, 'rPr')[0]!;
    expect(rPr).toBeDefined();
    const ins = rPr.getElementsByTagNameNS(W_NS, 'ins')[0]!;
    expect(ins).toBeDefined();
    expect(ins.getAttributeNS(W_NS, 'id') ?? ins.getAttribute('w:id')).toBe('0');
  });

  test('ensureParaMarkInsert:pPr 存在无 rPr → 建 rPr', () => {
    const pkg = makePkg(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    const paragraph = makeParagraph(pkg);
    ensureParaMarkInsert(paragraph, 1, attrs);
    const pPr = paragraph.getElementsByTagNameNS(W_NS, 'pPr')[0]!;
    // pPr 只有一个 pPrs,但现在多了 rPr
    expect(pPr.getElementsByTagNameNS(W_NS, 'rPr').length).toBe(1);
    expect(pPr.getElementsByTagNameNS(W_NS, 'pStyle').length).toBe(1); // pStyle 仍在
  });

  test('ensureParaMarkDelete:pPr/rPr 都已存在 → rPr 追加 del', () => {
    const pkg = makePkg(`<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`);
    const paragraph = makeParagraph(pkg);
    ensureParaMarkDelete(paragraph, 2, attrs);
    const rPr = paragraph.getElementsByTagNameNS(W_NS, 'rPr')[0]!;
    // rPr 现在有 w:b 和 w:del
    expect(rPr.getElementsByTagNameNS(W_NS, 'b').length).toBe(1);
    expect(rPr.getElementsByTagNameNS(W_NS, 'del').length).toBe(1);
  });
});
