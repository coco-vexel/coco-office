/**
 * ParagraphFormat 颜色字段单测（task-20260615-redhead-fingerprint）：
 *   - color（w:pPr/w:rPr/w:color@val）—— 段标级颜色
 *   - firstRunColor（首个直属 w:r/w:rPr/w:color@val）—— run 级颜色
 *
 * 纯增可选字段，向后兼容：无 color 时字段缺失。
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

describe('ParagraphFormat.color（段标级，task-20260615）', () => {
  test('w:pPr/w:rPr/w:color → color = RRGGBB', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.color).toBe('FF0000');
  });

  test('w:color val="auto" → 透传 auto', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:pPr><w:rPr><w:color w:val="auto"/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.color).toBe('auto');
  });

  test('无段标 color → 字段缺失', () => {
    const pkg = pkgFromBody(`<w:p w14:paraId="${pid(1)}"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(firstParaFormat(pkg)?.color).toBeUndefined();
  });
});

describe('ParagraphFormat.firstRunColor（run 级，task-20260615）', () => {
  test('首个直属 run 的 w:rPr/w:color → firstRunColor = RRGGBB', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>红头</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.firstRunColor).toBe('FF0000');
  });

  test('首 run 无 color、第二 run 有 color → 仍缺失（只看首 run）', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r>` +
        `<w:r><w:rPr><w:color w:val="00FF00"/></w:rPr><w:t>b</w:t></w:r></w:p>`,
    );
    expect(firstParaFormat(pkg)?.firstRunColor).toBeUndefined();
  });

  test('首 run 无 rPr → firstRunColor 缺失', () => {
    const pkg = pkgFromBody(`<w:p w14:paraId="${pid(1)}"><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(firstParaFormat(pkg)?.firstRunColor).toBeUndefined();
  });
});

describe('段标 color 与首 run color 共存', () => {
  test('两者独立采集', () => {
    const pkg = pkgFromBody(
      `<w:p w14:paraId="${pid(1)}"><w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr>` +
        `<w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
    );
    const fmt = firstParaFormat(pkg);
    expect(fmt?.color).toBe('FF0000');
    expect(fmt?.firstRunColor).toBe('0000FF');
  });
});
