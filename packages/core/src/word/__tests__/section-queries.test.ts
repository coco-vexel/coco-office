/**
 * listSections 测试（#211）。
 *
 * 覆盖：
 *   - 无 body / 无 sectPr → []
 *   - 单节（body 末尾散 sectPr）
 *   - 多节（pPr/sectPr 分隔 + 末尾散 sectPr）
 *   - sectPr type 解析（nextPage / continuous / evenPage / oddPage / nextColumn）
 *   - headerRefs / footerRefs（type='default' / 'first' / 'even'，resolve 到 partName）
 *   - titlePg
 *
 * ParagraphFormat 新增字段 pageBreakBefore / hasManualPageBreak / hasSectionBreak
 * 的测试合并到本文件 list-paragraphs 已覆盖部分（这里只补 sectPr 触发的）。
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { listSections } from '../section-queries.js';
import { W_NS, W14_NS, R_NS } from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const RELS_CTYPE = 'application/vnd.openxmlformats-package.relationships+xml';

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
  return String(i).padStart(8, 'A').slice(-8).toUpperCase();
}

function makeBody(bodyInner: string): Package {
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="${R_NS}"><w:body>${bodyInner}</w:body></w:document>`,
    DOC_CTYPE,
  );
  return pkg;
}

function addRels(pkg: Package, rels: Array<{ id: string; target: string; type?: string }>): void {
  const xml = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels
    .map(
      (r) =>
        `<Relationship Id="${r.id}" Type="${r.type ?? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'}" Target="${r.target}"/>`,
    )
    .join('')}</Relationships>`;
  putXml(pkg, 'word/_rels/document.xml.rels', xml, RELS_CTYPE);
}

describe('listSections — 边界', () => {
  test('无 body → []', () => {
    const pkg = createPackage();
    expect(listSections(pkg)).toEqual([]);
  });

  test('body 但无 sectPr → []', () => {
    const pkg = makeBody(`<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r></w:p>`);
    expect(listSections(pkg)).toEqual([]);
  });
});

describe('listSections — 单节', () => {
  test('body 末尾散 sectPr → 一节', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r></w:p>` +
        `<w:p w14:paraId="${pid(2)}"><w:r><w:t>b</w:t></w:r></w:p>` +
        `<w:sectPr/>`,
    );
    const out = listSections(pkg);
    expect(out).toHaveLength(1);
    expect(out[0]?.sectionIndex).toBe(0);
    expect(out[0]?.firstParaId).toBe(pid(1));
    expect(out[0]?.lastParaId).toBe(pid(2));
  });
});

describe('listSections — 多节', () => {
  test('pPr/sectPr 分割 + body 末尾散 sectPr → 两节', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:t>cover</w:t></w:r></w:p>` +
        `<w:p w14:paraId="${pid(2)}"><w:pPr><w:sectPr/></w:pPr><w:r><w:t>cover2</w:t></w:r></w:p>` +
        `<w:p w14:paraId="${pid(3)}"><w:r><w:t>body</w:t></w:r></w:p>` +
        `<w:p w14:paraId="${pid(4)}"><w:r><w:t>body2</w:t></w:r></w:p>` +
        `<w:sectPr/>`,
    );
    const out = listSections(pkg);
    expect(out).toHaveLength(2);
    expect(out[0]?.firstParaId).toBe(pid(1));
    expect(out[0]?.lastParaId).toBe(pid(2));
    expect(out[1]?.firstParaId).toBe(pid(3));
    expect(out[1]?.lastParaId).toBe(pid(4));
  });
});

describe('listSections — sectPr type', () => {
  test.each([['nextPage'], ['continuous'], ['evenPage'], ['oddPage'], ['nextColumn']] as const)(
    'type=%s',
    (typeVal) => {
      const pkg = makeBody(
        `<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r></w:p>` +
          `<w:sectPr><w:type w:val="${typeVal}"/></w:sectPr>`,
      );
      const out = listSections(pkg);
      expect(out[0]?.type).toBe(typeVal);
    },
  );

  test('未知 type 不写入', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r></w:p>` +
        `<w:sectPr><w:type w:val="unknown"/></w:sectPr>`,
    );
    expect(listSections(pkg)[0]?.type).toBeUndefined();
  });
});

describe('listSections — headerRefs / footerRefs', () => {
  test('headerReference + footerReference resolve 到 partName', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r></w:p>` +
        `<w:sectPr>` +
        `<w:headerReference w:type="default" r:id="rId10"/>` +
        `<w:headerReference w:type="first" r:id="rId11"/>` +
        `<w:footerReference w:type="default" r:id="rId12"/>` +
        `</w:sectPr>`,
    );
    addRels(pkg, [
      { id: 'rId10', target: 'header1.xml' },
      { id: 'rId11', target: 'header2.xml' },
      {
        id: 'rId12',
        target: 'footer1.xml',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
      },
    ]);
    const out = listSections(pkg);
    expect(out[0]?.headerRefs).toEqual([
      { type: 'default', partName: 'word/header1.xml' },
      { type: 'first', partName: 'word/header2.xml' },
    ]);
    expect(out[0]?.footerRefs).toEqual([{ type: 'default', partName: 'word/footer1.xml' }]);
  });

  test('rels 不存在的 rId → 跳过该 ref', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r></w:p>` +
        `<w:sectPr><w:headerReference w:type="default" r:id="rIdMissing"/></w:sectPr>`,
    );
    addRels(pkg, []);
    expect(listSections(pkg)[0]?.headerRefs).toBeUndefined();
  });
});

describe('listSections — titlePg', () => {
  test('含 w:titlePg → titlePage=true', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r></w:p>` +
        `<w:sectPr><w:titlePg/></w:sectPr>`,
    );
    expect(listSections(pkg)[0]?.titlePage).toBe(true);
  });

  test('无 w:titlePg → 字段缺失', () => {
    const pkg = makeBody(
      `<w:p w14:paraId="${pid(1)}"><w:r><w:t>a</w:t></w:r></w:p>` + `<w:sectPr/>`,
    );
    expect(listSections(pkg)[0]?.titlePage).toBeUndefined();
  });
});
