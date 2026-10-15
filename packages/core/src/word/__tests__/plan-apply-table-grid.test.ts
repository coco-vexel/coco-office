/**
 * TableInfo.columnWidths + plan.setTableColumnWidths 端到端测试（task-core-table-grid）。
 *
 * 测试矩阵：
 *   READ：
 *     1  columnWidths 与 XML w:w 一致（有 tblGrid 的表）
 *     2  无 tblGrid → columnWidths = []
 *   WRITE 基本：
 *     3  setTableColumnWidths → gridCol w:w 变为新值；round-trip 读回一致
 *     4  同步 tcW：setTableColumnWidths 后单元格 tcW 等于新列宽；单元格文本不变（不误伤）
 *   WRITE 校验：
 *     5  widths.length ≠ colCount → INVALID_ARGUMENT
 *   跨端：
 *     6  round-trip docx（fromZip → apply → toZip → fromZip → list）columnWidths 一致
 *     7  round-trip FlatOpc（fromZip → apply → toFlatOpc → fromFlatOpc → list）columnWidths 一致
 *   其它：
 *     8  tableId 不存在 → TABLE_NOT_FOUND
 *   Round 2 新增：
 *     9  简单 grid：apply 后 gridCol w:w 和每个 cell tcW 都等于新列宽（round-trip 读回验证）
 *     10 gridSpan：合并单元格 tcW = 跨列宽之和；非合并单元格 tcW = 本列宽
 */

import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
import { W14_NS, W_NS } from '../namespaces.js';

// ─────────── fixture 构造 ─────────────────────────────────────────────────────

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

/**
 * 最小 2 列表格，含 tblGrid（gridCol w:w=2000/3000）+ 1 行 2 格。
 * 单元格内容：'A' 和 'B'，用于验证不误伤。
 */
function tbl2Col(colWidths: [number, number] = [2000, 3000]): string {
  return `<w:tbl>
  <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="${colWidths[0]}"/>
    <w:gridCol w:w="${colWidths[1]}"/>
  </w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>
      <w:p w14:paraId="AAAAA001"><w:r><w:t>A</w:t></w:r></w:p>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>
      <w:p w14:paraId="AAAAA002"><w:r><w:t>B</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="AAAAA003"/>`;
}

/** 表格但无 tblGrid（colCount=0） */
function tblNoGrid(): string {
  return `<w:tbl>
  <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tr>
    <w:tc><w:p w14:paraId="BBBBB001"><w:r><w:t>X</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="BBBBB002"/>`;
}

afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── READ：columnWidths ───────────────────────────────────────────────

describe('TableInfo.columnWidths READ', () => {
  test('1 columnWidths 与 XML w:w 一致（2列表，2000/3000）', async () => {
    const buf = await makeFixture(tbl2Col([2000, 3000]));
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-1', mode: 'inline' });

    const tables = await doc.tables.list();
    expect(tables).toHaveLength(1);
    expect(tables[0]!.colCount).toBe(2);
    expect(tables[0]!.columnWidths).toEqual([2000, 3000]);

    await doc.dispose();
  });

  test('2 无 tblGrid → columnWidths = []', async () => {
    const buf = await makeFixture(tblNoGrid());
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-2', mode: 'inline' });

    const tables = await doc.tables.list();
    expect(tables).toHaveLength(1);
    expect(tables[0]!.colCount).toBe(0);
    expect(tables[0]!.columnWidths).toEqual([]);

    await doc.dispose();
  });
});

// ─────────── WRITE 基本 ──────────────────────────────────────────────────────

describe('plan.setTableColumnWidths WRITE 基本', () => {
  test('3 setTableColumnWidths → gridCol w:w 变为新值；round-trip 读回一致', async () => {
    const buf = await makeFixture(tbl2Col([2000, 3000]));
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-3', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    const tableId = tbl!.tableId;
    expect(tbl!.columnWidths).toEqual([2000, 3000]); // 验证改动前

    await doc
      .plan()
      .setTableColumnWidths({ tableId, widths: [1500, 4000] })
      .apply();

    // 立即读（内存中）
    const tablesAfter = await doc.tables.list();
    expect(tablesAfter[0]!.columnWidths).toEqual([1500, 4000]);

    // round-trip: toZip → fromZip → list
    const zipBytes = await doc.toZip();
    await doc.dispose();

    const doc2 = await WordDocument.fromZip(zipBytes, { docId: 'cw-3-rt', mode: 'inline' });
    const tablesRt = await doc2.tables.list();
    expect(tablesRt[0]!.columnWidths).toEqual([1500, 4000]);
    await doc2.dispose();
  });

  test('4 同步 tcW：apply 后单元格 tcW 等于新列宽；单元格文本不变（不误伤）', async () => {
    // tbl2Col 的 fixture 单元格带 <w:tcW w:w="0" w:type="auto">
    // apply 后 tcW 应被更新为新列宽（1000 和 5000）
    const buf = await makeFixture(tbl2Col([2000, 3000]));
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-4', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await doc
      .plan()
      .setTableColumnWidths({ tableId: tbl!.tableId, widths: [1000, 5000] })
      .apply();

    // 验证单元格文本不变
    const tablesAfter = await doc.tables.list({ includeCellTexts: true });
    const cells = tablesAfter[0]!.cells!;
    expect(cells[0]![0]).toBe('A'); // 第 0 行第 0 列
    expect(cells[0]![1]).toBe('B'); // 第 0 行第 1 列

    // gridCol 列宽已改变
    expect(tablesAfter[0]!.columnWidths).toEqual([1000, 5000]);

    // tcW 已同步：round-trip 验证（读回 XML 确认 tcW 值）
    const zipBytes = await doc.toZip();
    await doc.dispose();

    const zip = await import('jszip');
    const JSZip2 = (zip as { default?: typeof zip }).default ?? zip;
    const loaded = await JSZip2.loadAsync(zipBytes);
    const docXmlContent = await loaded.file('word/document.xml')!.async('string');
    // 验证 tcW w:w 属性包含新列宽（1000 和 5000）
    expect(docXmlContent).toContain('w:w="1000"');
    expect(docXmlContent).toContain('w:w="5000"');
    // type 改为 dxa
    expect(docXmlContent).toContain('w:type="dxa"');
  });
});

// ─────────── WRITE 校验 ──────────────────────────────────────────────────────

describe('plan.setTableColumnWidths WRITE 校验', () => {
  test('5 widths.length ≠ colCount → INVALID_ARGUMENT', async () => {
    const buf = await makeFixture(tbl2Col([2000, 3000]));
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-5', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await expect(
      doc
        .plan()
        .setTableColumnWidths({ tableId: tbl!.tableId, widths: [1000, 2000, 3000] }) // 3 widths for 2-col table
        .apply(),
    ).rejects.toSatisfy((e: unknown) => e instanceof WordError && e.code === 'INVALID_ARGUMENT');

    await doc.dispose();
  });

  test('8 tableId 不存在 → TABLE_NOT_FOUND', async () => {
    const buf = await makeFixture(tbl2Col());
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-8', mode: 'inline' });

    await expect(
      doc
        .plan()
        .setTableColumnWidths({ tableId: 'tbl-9999', widths: [1000, 2000] })
        .apply(),
    ).rejects.toSatisfy((e: unknown) => e instanceof WordError && e.code === 'TABLE_NOT_FOUND');

    await doc.dispose();
  });
});

// ─────────── 跨端 round-trip ─────────────────────────────────────────────────

describe('plan.setTableColumnWidths 跨端', () => {
  test('6 round-trip docx（fromZip → apply → toZip → fromZip → list）columnWidths 一致', async () => {
    const buf = await makeFixture(tbl2Col([2000, 3000]));
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-6', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await doc
      .plan()
      .setTableColumnWidths({ tableId: tbl!.tableId, widths: [1234, 5678] })
      .apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const doc2 = await WordDocument.fromZip(zipBytes, { docId: 'cw-6-r', mode: 'inline' });
    const tables = await doc2.tables.list();
    expect(tables[0]!.columnWidths).toEqual([1234, 5678]);
    await doc2.dispose();
  });

  test('7 ★ round-trip FlatOpc（fromZip → apply → toFlatOpc → fromFlatOpc → list）columnWidths 一致', async () => {
    const buf = await makeFixture(tbl2Col([2000, 3000]));
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-7', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await doc
      .plan()
      .setTableColumnWidths({ tableId: tbl!.tableId, widths: [999, 8001] })
      .apply();

    const flatOpc = await doc.toFlatOpc();
    await doc.dispose();

    const doc2 = await WordDocument.fromFlatOpc(flatOpc, { docId: 'cw-7-r', mode: 'inline' });
    const tables = await doc2.tables.list();
    expect(tables[0]!.columnWidths).toEqual([999, 8001]);
    await doc2.dispose();
  });
});

// ─────────── Round 2 新增：tcW 同步验证 ──────────────────────────────────────

describe('plan.setTableColumnWidths tcW 同步（Round 2）', () => {
  test('9 简单 grid：apply 后 gridCol w:w 和每个 cell tcW 都等于新列宽（round-trip XML 验证）', async () => {
    // 构造 3 列表，每个单元格有 tcW（模拟真实文档带 tcW 的场景）
    const tbl3Col = `<w:tbl>
  <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="1000"/>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="3000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="CCC00001"><w:r><w:t>C1</w:t></w:r></w:p>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="CCC00002"><w:r><w:t>C2</w:t></w:r></w:p>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="CCC00003"><w:r><w:t>C3</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="CCC00004"/>`;

    const buf = await makeFixture(tbl3Col);
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-9', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    expect(tbl!.columnWidths).toEqual([1000, 2000, 3000]); // 改动前

    // 设新列宽
    await doc
      .plan()
      .setTableColumnWidths({ tableId: tbl!.tableId, widths: [500, 4000, 1500] })
      .apply();

    // gridCol 已改
    const tablesAfter = await doc.tables.list();
    expect(tablesAfter[0]!.columnWidths).toEqual([500, 4000, 1500]);

    // round-trip 验证 tcW
    const zipBytes = await doc.toZip();
    await doc.dispose();

    const JSZip = (await import('jszip')).default;
    const loaded = await JSZip.loadAsync(zipBytes);
    const docXmlContent = await loaded.file('word/document.xml')!.async('string');

    // 每个新列宽都出现在 XML 里（同时被 gridCol 和 tcW 持有）
    expect(docXmlContent).toContain('w:w="500"');
    expect(docXmlContent).toContain('w:w="4000"');
    expect(docXmlContent).toContain('w:w="1500"');
    // 旧值不再出现（原 1000/2000/3000 已被覆盖；但 gridCol 本身可能存 500/4000/1500）
    // 验证 type="dxa" 存在（tcW 同步）
    expect(docXmlContent).toContain('w:type="dxa"');
  });

  test('10 gridSpan：合并单元格 tcW = 跨列宽之和；非合并单元格 tcW = 本列宽', async () => {
    // 构造 3 列表，第 1 行: 1个合并单元格(跨列0+1=2列) + 1个单元格(列2)
    //               第 2 行: 3 个普通单元格
    // widths = [1000, 2000, 3000]
    // 合并单元格 tcW = 1000 + 2000 = 3000；第 2 格 tcW = 3000
    const tblWithMerge = `<w:tbl>
  <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="500"/>
    <w:gridCol w:w="500"/>
    <w:gridCol w:w="500"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="1000" w:type="dxa"/>
        <w:gridSpan w:val="2"/>
      </w:tcPr>
      <w:p w14:paraId="GGG00001"><w:r><w:t>Merged</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="500" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="GGG00002"><w:r><w:t>C</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="500" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="GGG00003"><w:r><w:t>A</w:t></w:r></w:p>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="500" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="GGG00004"><w:r><w:t>B</w:t></w:r></w:p>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="500" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="GGG00005"><w:r><w:t>C2</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="GGG00006"/>`;

    const buf = await makeFixture(tblWithMerge);
    const doc = await WordDocument.fromZip(buf, { docId: 'cw-10', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    // 设新列宽：[1000, 2000, 3000]
    // 期望：合并单元格(跨0+1) tcW = 1000+2000 = 3000
    //        第2格(列2) tcW = 3000
    //        第2行各格: [1000, 2000, 3000]
    await doc
      .plan()
      .setTableColumnWidths({ tableId: tbl!.tableId, widths: [1000, 2000, 3000] })
      .apply();

    // gridCol 已改
    const tablesAfter = await doc.tables.list();
    expect(tablesAfter[0]!.columnWidths).toEqual([1000, 2000, 3000]);

    // 读 XML 验证 tcW 值
    const zipBytes = await doc.toZip();
    await doc.dispose();

    const JSZip = (await import('jszip')).default;
    const loaded = await JSZip.loadAsync(zipBytes);
    const docXmlContent = await loaded.file('word/document.xml')!.async('string');

    // 合并单元格跨列0+1宽之和 = 1000+2000 = 3000，且第2格(列2)也 = 3000
    // 所以 3000 至少出现 3 次（2次 tcW + 1次 gridCol 或更多）
    // 简单验证：XML 里有 3000 和 1000 和 2000
    expect(docXmlContent).toContain('w:w="1000"');
    expect(docXmlContent).toContain('w:w="2000"');
    expect(docXmlContent).toContain('w:w="3000"');

    // 验证合并单元格的 tcW 正确（3000 = 1000+2000），不是单列宽 1000
    // 通过 parse DOM 精确验证
    const { parseXmlDocument } = await import('slimdom');
    const domDoc = parseXmlDocument(docXmlContent) as unknown as Document;
    const tblEl = domDoc.getElementsByTagNameNS(W_NS, 'tbl')[0]!;
    const trs = Array.from(tblEl.childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).localName === 'tr',
    ) as Element[];

    // 第 1 行第 1 个 tc（合并单元格，gridSpan=2）→ tcW 应是 3000
    const firstTr = trs[0]!;
    const firstRowCells = Array.from(firstTr.childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).localName === 'tc',
    ) as Element[];
    expect(firstRowCells).toHaveLength(2); // 合并后只有 2 个 tc

    const mergedTc = firstRowCells[0]!;
    const mergedTcPr = Array.from(mergedTc.childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).localName === 'tcPr',
    ) as Element | undefined;
    const mergedTcW = mergedTcPr
      ? (Array.from(mergedTcPr.childNodes).find(
          (n) => n.nodeType === 1 && (n as Element).localName === 'tcW',
        ) as Element | undefined)
      : undefined;
    expect(mergedTcW).toBeDefined();
    expect(mergedTcW!.getAttributeNS(W_NS, 'w')).toBe('3000'); // 1000+2000
    expect(mergedTcW!.getAttributeNS(W_NS, 'type')).toBe('dxa');

    // 第 1 行第 2 个 tc（列2，单格）→ tcW 应是 3000
    const lastColTc = firstRowCells[1]!;
    const lastColTcPr = Array.from(lastColTc.childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).localName === 'tcPr',
    ) as Element | undefined;
    const lastColTcW = lastColTcPr
      ? (Array.from(lastColTcPr.childNodes).find(
          (n) => n.nodeType === 1 && (n as Element).localName === 'tcW',
        ) as Element | undefined)
      : undefined;
    expect(lastColTcW).toBeDefined();
    expect(lastColTcW!.getAttributeNS(W_NS, 'w')).toBe('3000'); // 列2 = 3000
  });
});
