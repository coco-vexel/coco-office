/**
 * insertTableFrom + getTableVerbatim 端到端测试（task-20260616-table-content-driven）。
 *
 * 验证"数据表 = 内容块"底座：跨文档 verbatim 搬表。
 *   1  getTableVerbatim 取回 <w:tbl> 子树 + 自包含 payload
 *   2  insertTableFrom 把源表 verbatim 搬进目标文档，内容（行/列/值/合并）零丢
 *   3  搬入表的 <w:p> paraId 与 host 既有不撞（重分配，全局唯一）
 *   4  源表内锚点（书签/批注）被剥离
 *   5  styleTblPrXml 覆盖搬入表的 tblPr（样式归一）
 *   6  多张不同结构表分别搬入 → 各归各位，零 cross-wire（真机 bug 反例）
 *   7  anchor 不存在 → PARA_NOT_FOUND；非 <w:tbl> → INVALID_ARGUMENT
 */

import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
import { W14_NS, W_NS } from '../namespaces.js';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

async function makeFixture(bodyInner: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  return zip.generateAsync({ type: 'uint8array' });
}

/** 材料表：2 列、表头 + 1 数据行。 */
function materialTable(): string {
  return `<w:tbl>
  <w:tblPr><w:tblStyle w:val="ContentStyle"/><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
  <w:tr>
    <w:tc><w:p w14:paraId="CCCCC001"><w:r><w:t>材料名称</w:t></w:r></w:p></w:tc>
    <w:tc><w:p w14:paraId="CCCCC002"><w:r><w:t>数量</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p w14:paraId="CCCCC003"><w:r><w:t>水泥</w:t></w:r></w:p></w:tc>
    <w:tc><w:p w14:paraId="CCCCC004"><w:r><w:t>100吨</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>`;
}

/** 参数表：3 列、含合并表头（gridSpan=2）+ 1 数据行。 */
function paramTable(): string {
  return `<w:tbl>
  <w:tblPr><w:tblStyle w:val="ContentStyle"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p w14:paraId="DDDDD001"><w:r><w:t>参数名称</w:t></w:r></w:p></w:tc>
    <w:tc><w:p w14:paraId="DDDDD002"><w:r><w:t>取值</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p w14:paraId="DDDDD003"><w:r><w:t>压力</w:t></w:r></w:p></w:tc>
    <w:tc><w:p w14:paraId="DDDDD004"><w:r><w:t>单位</w:t></w:r></w:p></w:tc>
    <w:tc><w:p w14:paraId="DDDDD005"><w:r><w:t>2.5MPa</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>`;
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('getTableVerbatim + insertTableFrom（数据表=内容块）', () => {
  test('1+2 getTableVerbatim 取回，insertTableFrom verbatim 搬入，内容零丢', async () => {
    const contentBuf = await makeFixture(materialTable() + '<w:p w14:paraId="CONTEN01"/>');
    const contentDoc = await WordDocument.fromZip(contentBuf, { docId: 'c1', mode: 'inline' });
    const verb = await contentDoc.tables.getVerbatim('tbl-0001');
    expect(verb).not.toBeNull();
    expect(verb!.tableXml).toContain('材料名称');
    expect(verb!.tableXml).toContain('水泥');
    expect(verb!.images).toHaveLength(0);

    // 目标文档：仅一个 anchor 段。
    const outBuf = await makeFixture('<w:p w14:paraId="OUTANCH1"><w:r><w:t>正文</w:t></w:r></w:p>');
    const outDoc = await WordDocument.fromZip(outBuf, { docId: 'o1', mode: 'inline' });
    const plan = outDoc.plan();
    plan.insertTableFrom({ anchorParaId: 'OUTANCH1', tableXml: verb!.tableXml });
    await plan.apply();

    const tables = await outDoc.tables.list({ includeCellTexts: true });
    expect(tables).toHaveLength(1);
    const cells = tables[0]!.cells!;
    expect(cells[0]).toEqual(['材料名称', '数量']);
    expect(cells[1]).toEqual(['水泥', '100吨']);
    expect(tables[0]!.rowCount).toBe(2);
    expect(tables[0]!.colCount).toBe(2);

    contentDoc.dispose();
    outDoc.dispose();
  });

  test('3 搬入表 <w:p> paraId 与 host 既有不撞（重分配）', async () => {
    const contentBuf = await makeFixture(materialTable());
    const contentDoc = await WordDocument.fromZip(contentBuf, { docId: 'c3', mode: 'inline' });
    const verb = await contentDoc.tables.getVerbatim('tbl-0001');

    // host 故意已有 paraId=CCCCC001（与源表首格段 paraId 相同）→ 搬入必须重分配避免撞。
    const outBuf = await makeFixture(
      '<w:p w14:paraId="CCCCC001"><w:r><w:t>已有同名段</w:t></w:r></w:p>',
    );
    const outDoc = await WordDocument.fromZip(outBuf, { docId: 'o3', mode: 'inline' });
    const plan = outDoc.plan();
    plan.insertTableFrom({ anchorParaId: 'CCCCC001', tableXml: verb!.tableXml });
    await plan.apply();

    const { paragraphs } = await outDoc.paragraphs.list({
      scope: 'body',
      pagination: { mode: 'all' },
    });
    const ids = paragraphs.map((p) => p.paraId.toUpperCase());
    // 不能有重复 paraId。
    expect(new Set(ids).size).toBe(ids.length);
    // 原 host 段仍在。
    expect(ids).toContain('CCCCC001');

    contentDoc.dispose();
    outDoc.dispose();
  });

  test('4 源表内书签锚点被剥离', async () => {
    const tblWithBookmark = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr><w:tc>
        <w:p w14:paraId="EEEEE001">
          <w:bookmarkStart w:id="9" w:name="srcBm"/><w:r><w:t>带书签格</w:t></w:r><w:bookmarkEnd w:id="9"/>
        </w:p>
      </w:tc></w:tr>
    </w:tbl>`;
    const contentBuf = await makeFixture(tblWithBookmark);
    const contentDoc = await WordDocument.fromZip(contentBuf, { docId: 'c4', mode: 'inline' });
    const verb = await contentDoc.tables.getVerbatim('tbl-0001');

    const outBuf = await makeFixture('<w:p w14:paraId="OUTANCH4"/>');
    const outDoc = await WordDocument.fromZip(outBuf, { docId: 'o4', mode: 'inline' });
    const plan = outDoc.plan();
    plan.insertTableFrom({ anchorParaId: 'OUTANCH4', tableXml: verb!.tableXml });
    await plan.apply();

    const zip = await outDoc.toZip();
    const jz = await JSZip.loadAsync(zip);
    const docXmlOut = await jz.file('word/document.xml')!.async('string');
    expect(docXmlOut).toContain('带书签格'); // 文字保真
    expect(docXmlOut).not.toContain('bookmarkStart'); // 锚点剥离
    expect(docXmlOut).not.toContain('srcBm');

    contentDoc.dispose();
    outDoc.dispose();
  });

  test('5 styleTblPrXml 覆盖搬入表 tblPr（样式归一）', async () => {
    const contentBuf = await makeFixture(materialTable());
    const contentDoc = await WordDocument.fromZip(contentBuf, { docId: 'c5', mode: 'inline' });
    const verb = await contentDoc.tables.getVerbatim('tbl-0001');

    const outBuf = await makeFixture('<w:p w14:paraId="OUTANCH5"/>');
    const outDoc = await WordDocument.fromZip(outBuf, { docId: 'o5', mode: 'inline' });
    const plan = outDoc.plan();
    // 模板数据表样式：tblStyle=TemplateStyle（覆盖内容表的 ContentStyle）。
    const styleTblPrXml = `<w:tblPr xmlns:w="${W_NS}"><w:tblStyle w:val="TemplateStyle"/><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>`;
    plan.insertTableFrom({ anchorParaId: 'OUTANCH5', tableXml: verb!.tableXml, styleTblPrXml });
    await plan.apply();

    const tables = await outDoc.tables.list({ includeCellTexts: true });
    expect(tables[0]!.tableStyleId).toBe('TemplateStyle'); // 归一到模板样式
    // 内容仍 verbatim 保真。
    expect(tables[0]!.cells![1]).toEqual(['水泥', '100吨']);

    contentDoc.dispose();
    outDoc.dispose();
  });

  test('6 多张不同表分别搬入 → 各归各位，零 cross-wire（真机 bug 反例）', async () => {
    // content 文档有 2 张不同表（材料表 2列 + 参数表 3列）。
    const contentBuf = await makeFixture(
      materialTable() + '<w:p w14:paraId="MID00001"/>' + paramTable(),
    );
    const contentDoc = await WordDocument.fromZip(contentBuf, { docId: 'c6', mode: 'inline' });
    const list = await contentDoc.tables.list();
    expect(list).toHaveLength(2);
    const vMaterial = await contentDoc.tables.getVerbatim(list[0]!.tableId);
    const vParam = await contentDoc.tables.getVerbatim(list[1]!.tableId);

    // 模板只有 1 张表（数量/结构都不同）——旧对位填充会丢表 + 错配。content-driven 应忽略模板表数。
    const outBuf = await makeFixture(
      '<w:p w14:paraId="OUTA0001"><w:r><w:t>正文1</w:t></w:r></w:p>',
    );
    const outDoc = await WordDocument.fromZip(outBuf, { docId: 'o6', mode: 'inline' });

    // 逆序在同一 body anchor 后搬入（doc 序 = 材料表 → 参数表）。
    //   两表都锚到 body 级段 OUTA0001（绝不锚到表内 cell 段，否则会嵌套）；
    //   逆序插（先参数后材料），各自插在 anchor 之后 → 最终顺序 = 材料、参数。
    for (const v of [vParam!, vMaterial!]) {
      const plan = outDoc.plan();
      plan.insertTableFrom({ anchorParaId: 'OUTA0001', tableXml: v.tableXml });
      await plan.apply();
    }

    const tables = await outDoc.tables.list({ includeCellTexts: true });
    expect(tables).toHaveLength(2); // 两表全在（不丢表）
    // 材料表数据是材料（不是参数）——零 cross-wire。
    expect(tables[0]!.cells![0]).toEqual(['材料名称', '数量']);
    expect(tables[0]!.cells![1]).toEqual(['水泥', '100吨']);
    // 参数表数据是参数（含合并表头）——压力/单位/取值各归各位。
    expect(tables[1]!.cells![0]).toEqual(['参数名称', '取值']);
    expect(tables[1]!.cells![1]).toEqual(['压力', '单位', '2.5MPa']);
    expect(tables[1]!.hasMerge).toBe(true); // 合并表头保真

    contentDoc.dispose();
    outDoc.dispose();
  });

  test('7 anchor 不存在 → PARA_NOT_FOUND；非 w:tbl → INVALID_ARGUMENT', async () => {
    const outBuf = await makeFixture('<w:p w14:paraId="OUTANCH7"/>');
    const outDoc = await WordDocument.fromZip(outBuf, { docId: 'o7', mode: 'inline' });

    const p1 = outDoc.plan();
    p1.insertTableFrom({
      anchorParaId: 'NOSUCH00',
      tableXml: `<w:tbl xmlns:w="${W_NS}"><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
    });
    await expect(p1.apply()).rejects.toMatchObject({ code: 'PARA_NOT_FOUND' });

    const p2 = outDoc.plan();
    p2.insertTableFrom({
      anchorParaId: 'OUTANCH7',
      tableXml: `<w:p xmlns:w="${W_NS}"><w:r><w:t>not a table</w:t></w:r></w:p>`,
    });
    await expect(p2.apply()).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });

    outDoc.dispose();
  });
});

// 引用 WordError 以满足 lint（错误码断言走 toMatchObject）。
void WordError;
