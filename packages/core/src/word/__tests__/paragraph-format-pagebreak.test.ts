/**
 * ParagraphFormat 新增字段（#211）单测：
 *   - pageBreakBefore（w:pPr/w:pageBreakBefore）
 *   - hasManualPageBreak（w:r/w:br[@w:type="page"]）
 *   - hasSectionBreak（w:pPr/w:sectPr）
 *
 * 只验单字段触发；现有 format 字段（alignment / indent / spacing / fontSize 等）
 * 由 list-paragraphs 现有测试覆盖。
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { listParagraphs } from '../list-paragraphs.js';
import { W_NS, W14_NS } from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function pkgFromBody(bodyInner: string): Package {
  const pkg = createPackage();
  const xml = `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`;
  setPart(
    pkg,
    createXmlPart({
      path: 'word/document.xml',
      contentType: DOC_CTYPE,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
      originalBytes: new TextEncoder().encode(xml),
    }),
  );
  return pkg;
}

function pid(i: number): string {
  return String(i).padStart(8, 'A').slice(-8).toUpperCase();
}

function firstParaFormat(pkg: Package) {
  const { paragraphs } = listParagraphs(pkg, {
    pagination: { mode: 'all' },
    includeFormat: true,
  });
  return paragraphs[0]?.format;
}

describe('ParagraphFormat.pageBreakBefore (#211)', () => {
  test('w:pageBreakBefore 存在无 val → true', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.pageBreakBefore).toBe(true);
  });

  test('w:pageBreakBefore val="1" → true', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:pPr><w:pageBreakBefore w:val="1"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.pageBreakBefore).toBe(true);
  });

  test('w:pageBreakBefore val="0" → false', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.pageBreakBefore).toBe(false);
  });

  test('未设 → 字段缺失', () => {
    const pkg = pkgFromBody(`<w:p w14:paraId="${pid(1)}"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(firstParaFormat(pkg)?.pageBreakBefore).toBeUndefined();
  });
});

describe('ParagraphFormat.hasSectionBreak (#211)', () => {
  test('w:sectPr 存在 → true', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:pPr><w:sectPr/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.hasSectionBreak).toBe(true);
  });

  test('无 sectPr → 字段缺失', () => {
    const pkg = pkgFromBody(`<w:p w14:paraId="${pid(1)}"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(firstParaFormat(pkg)?.hasSectionBreak).toBeUndefined();
  });
});

describe('ParagraphFormat.hasManualPageBreak (#211)', () => {
  test('w:r/w:br[@w:type="page"] → true', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:br w:type="page"/></w:r><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.hasManualPageBreak).toBe(true);
  });

  test('w:br 无 type=page（普通换行） → false / 缺失', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:br/></w:r><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.hasManualPageBreak).toBeUndefined();
  });

  test('w:br type=column → false', () => {
    const pkg = pkgFromBody(`<w:p w14:paraId="${pid(1)}"><w:r><w:br w:type="column"/></w:r></w:p>`);
    expect(firstParaFormat(pkg)?.hasManualPageBreak).toBeUndefined();
  });
});

describe('多字段同时存在', () => {
  test('pageBreakBefore + hasManualPageBreak + hasSectionBreak 共存', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:pPr><w:pageBreakBefore/><w:sectPr/></w:pPr>` +
        `<w:r><w:br w:type="page"/></w:r><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    const fmt = firstParaFormat(pkg);
    expect(fmt?.pageBreakBefore).toBe(true);
    expect(fmt?.hasSectionBreak).toBe(true);
    expect(fmt?.hasManualPageBreak).toBe(true);
  });
});
