/**
 * TableInfo.rowHeights + plan.setTableRowHeights 端到端测试（task-core-table-rowheight）。
 *
 * 测试矩阵（spec §3）：
 *   1  setTableRowHeights 全行 {val, hRule:'atLeast'} → 每行 trPr 有 trHeight w:val/w:hRule="atLeast"
 *   2  hRule='exact' / 'auto' 各一 → hRule 正确写入
 *   3  某行传 null → 该行无 trHeight（原有的被删）
 *   4  已有 trHeight 的行覆盖 → val/hRule 被改，不重复 trHeight 元素
 *   5  trPr 不存在的行 → 新建 trPr 含 trHeight，trPr 在 tr 首位
 *   6  heights.length ≠ 行数 → INVALID_ARGUMENT
 *   7  round-trip → set → toZip → fromZip，doc.tables.get().rowHeights 读回值一致（val+hRule）
 *   8  读：TableInfo.rowHeights → 源表读出每行 val+hRule，无 trHeight 行为 null
 *   9  不回归：现有 list-tables / setTableColumnWidths 测试不受影响
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
 * 2 行 2 列表格，行 0 有 trHeight(val=800, hRule='atLeast')，行 1 无 trHeight。
 * 每行每格带 paraId（必填，避免 tsc 报错）。
 */
function tbl2Row(): string {
  return `<w:tbl>
  <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="3000"/>
  </w:tblGrid>
  <w:tr>
    <w:trPr><w:trHeight w:val="800" w:hRule="atLeast"/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="RRRR0001"><w:r><w:t>R0C0</w:t></w:r></w:p>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="RRRR0002"><w:r><w:t>R0C1</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="RRRR0003"><w:r><w:t>R1C0</w:t></w:r></w:p>
    </w:tc>
    <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="RRRR0004"><w:r><w:t>R1C1</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="RRRR0005"/>`;
}

/**
 * 3 行表格，所有行没有 trHeight，也没有 trPr（bare rows）。
 * 用于测试 5（trPr 不存在时新建）。
 */
function tbl3RowNoPr(): string {
  return `<w:tbl>
  <w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="2000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="BARE0001"><w:r><w:t>A</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="BARE0002"><w:r><w:t>B</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="BARE0003"><w:r><w:t>C</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="BARE0004"/>`;
}

afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── 8. 读：TableInfo.rowHeights ─────────────────────────────────────

describe('TableInfo.rowHeights READ', () => {
  test('8 读：源表读出每行 val+hRule，无 trHeight 行为 null', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-8', mode: 'inline' });

    const tables = await doc.tables.list();
    expect(tables).toHaveLength(1);
    const { rowHeights } = tables[0]!;

    // 行 0: trHeight val=800 hRule=atLeast
    expect(rowHeights).toHaveLength(2);
    expect(rowHeights[0]).toEqual({ val: 800, hRule: 'atLeast' });
    // 行 1: 无 trHeight → null
    expect(rowHeights[1]).toBeNull();

    await doc.dispose();
  });

  test('8b 读：无任何 trHeight 的表 → 全部 null', async () => {
    const buf = await makeFixture(tbl3RowNoPr());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-8b', mode: 'inline' });

    const tables = await doc.tables.list();
    expect(tables[0]!.rowHeights).toEqual([null, null, null]);

    await doc.dispose();
  });
});

// ─────────── WRITE 基本 ──────────────────────────────────────────────────────

describe('plan.setTableRowHeights WRITE 基本', () => {
  test('1 全行 {val, hRule:"atLeast"} → 每行 trPr 有 trHeight val/hRule="atLeast"', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-1', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    const tableId = tbl!.tableId;

    await doc
      .plan()
      .setTableRowHeights({
        tableId,
        heights: [
          { val: 600, hRule: 'atLeast' },
          { val: 700, hRule: 'atLeast' },
        ],
      })
      .apply();

    const tablesAfter = await doc.tables.list();
    const { rowHeights } = tablesAfter[0]!;
    expect(rowHeights[0]).toEqual({ val: 600, hRule: 'atLeast' });
    expect(rowHeights[1]).toEqual({ val: 700, hRule: 'atLeast' });

    // XML 验证
    const zipBytes = await doc.toZip();
    await doc.dispose();
    const loaded = await JSZip.loadAsync(zipBytes);
    const xml = await loaded.file('word/document.xml')!.async('string');
    expect(xml).toContain('w:val="600"');
    expect(xml).toContain('w:val="700"');
    expect(xml).toContain('w:hRule="atLeast"');
  });

  test('2 hRule="exact" / "auto" 各一 → hRule 正确写入', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-2', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await doc
      .plan()
      .setTableRowHeights({
        tableId: tbl!.tableId,
        heights: [
          { val: 500, hRule: 'exact' },
          { val: 300, hRule: 'auto' },
        ],
      })
      .apply();

    const tablesAfter = await doc.tables.list();
    const { rowHeights } = tablesAfter[0]!;
    expect(rowHeights[0]).toEqual({ val: 500, hRule: 'exact' });
    expect(rowHeights[1]).toEqual({ val: 300, hRule: 'auto' });

    await doc.dispose();
  });

  test('3 某行传 null → 该行无 trHeight（原有的被删）', async () => {
    // 行 0 有 trHeight=800, 传 null → 删除
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-3', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    // 确认改前行 0 有 trHeight
    expect(tbl!.rowHeights[0]).toEqual({ val: 800, hRule: 'atLeast' });

    await doc
      .plan()
      .setTableRowHeights({
        tableId: tbl!.tableId,
        heights: [null, null],
      })
      .apply();

    const tablesAfter = await doc.tables.list();
    expect(tablesAfter[0]!.rowHeights[0]).toBeNull();
    expect(tablesAfter[0]!.rowHeights[1]).toBeNull();

    // XML 中不应再有 trHeight
    const zipBytes = await doc.toZip();
    await doc.dispose();
    const loaded = await JSZip.loadAsync(zipBytes);
    const xml = await loaded.file('word/document.xml')!.async('string');
    expect(xml).not.toContain('trHeight');
  });

  test('4 已有 trHeight 的行覆盖 → val/hRule 被改，不重复 trHeight 元素', async () => {
    // 行 0 原有 trHeight(800, atLeast)，覆盖为 (1200, exact)
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-4', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await doc
      .plan()
      .setTableRowHeights({
        tableId: tbl!.tableId,
        heights: [{ val: 1200, hRule: 'exact' }, null],
      })
      .apply();

    const tablesAfter = await doc.tables.list();
    expect(tablesAfter[0]!.rowHeights[0]).toEqual({ val: 1200, hRule: 'exact' });

    // 验证不存在重复 trHeight 元素（只出现一次）
    const zipBytes = await doc.toZip();
    await doc.dispose();
    const loaded = await JSZip.loadAsync(zipBytes);
    const xml = await loaded.file('word/document.xml')!.async('string');
    // 计数 trHeight 出现次数 = 1（行 0 覆盖，行 1 无）
    const matches = [...xml.matchAll(/trHeight/g)];
    expect(matches.length).toBe(1);
    expect(xml).toContain('w:val="1200"');
    expect(xml).toContain('w:hRule="exact"');
  });

  test('5 trPr 不存在的行 → 新建 trPr 含 trHeight，trPr 在 tr 首位', async () => {
    // tbl3RowNoPr 的行都没有 trPr
    const buf = await makeFixture(tbl3RowNoPr());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-5', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await doc
      .plan()
      .setTableRowHeights({
        tableId: tbl!.tableId,
        heights: [
          { val: 400, hRule: 'atLeast' },
          { val: 500, hRule: 'exact' },
          { val: 600, hRule: 'auto' },
        ],
      })
      .apply();

    const tablesAfter = await doc.tables.list();
    const { rowHeights } = tablesAfter[0]!;
    expect(rowHeights[0]).toEqual({ val: 400, hRule: 'atLeast' });
    expect(rowHeights[1]).toEqual({ val: 500, hRule: 'exact' });
    expect(rowHeights[2]).toEqual({ val: 600, hRule: 'auto' });

    // XML 验证: trPr 紧接 tr 开标签之后，trHeight 在 trPr 内
    const zipBytes = await doc.toZip();
    await doc.dispose();
    const loaded = await JSZip.loadAsync(zipBytes);
    const xml = await loaded.file('word/document.xml')!.async('string');

    // 检查 trPr 与 trHeight 都存在
    expect(xml).toContain('w:trPr');
    expect(xml).toContain('w:trHeight');

    // 验证 trPr 是 tr 的首个子元素（trPr 紧接 <w:tr> 后）
    const { parseXmlDocument } = await import('slimdom');
    const dom = parseXmlDocument(xml) as unknown as Document;
    const tblEl = dom.getElementsByTagNameNS(W_NS, 'tbl')[0]!;
    const trs = Array.from(tblEl.childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).localName === 'tr',
    ) as Element[];
    expect(trs).toHaveLength(3);

    for (const tr of trs) {
      const firstEl = Array.from(tr.childNodes).find((n) => n.nodeType === 1) as
        | Element
        | undefined;
      expect(firstEl).toBeDefined();
      expect(firstEl!.localName).toBe('trPr');

      // trPr 内有 trHeight
      const trHeight = Array.from(firstEl!.childNodes).find(
        (n) => n.nodeType === 1 && (n as Element).localName === 'trHeight',
      ) as Element | undefined;
      expect(trHeight).toBeDefined();
    }
  });
});

// ─────────── WRITE 校验 ──────────────────────────────────────────────────────

describe('plan.setTableRowHeights WRITE 校验', () => {
  test('6 heights.length ≠ 行数 → INVALID_ARGUMENT', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-6', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    // 2 行表格传 3 个高度
    await expect(
      doc
        .plan()
        .setTableRowHeights({
          tableId: tbl!.tableId,
          heights: [
            { val: 400, hRule: 'atLeast' },
            { val: 500, hRule: 'exact' },
            { val: 600, hRule: 'auto' },
          ],
        })
        .apply(),
    ).rejects.toSatisfy((e: unknown) => e instanceof WordError && e.code === 'INVALID_ARGUMENT');

    await doc.dispose();
  });

  test('6b tableId 不存在 → TABLE_NOT_FOUND', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-6b', mode: 'inline' });

    await expect(
      doc
        .plan()
        .setTableRowHeights({
          tableId: 'tbl-9999',
          heights: [{ val: 400, hRule: 'atLeast' }],
        })
        .apply(),
    ).rejects.toSatisfy((e: unknown) => e instanceof WordError && e.code === 'TABLE_NOT_FOUND');

    await doc.dispose();
  });
});

// ─────────── round-trip ──────────────────────────────────────────────────────

describe('plan.setTableRowHeights round-trip', () => {
  test('7 set → toZip → fromZip → rowHeights 读回值一致（val+hRule）', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-7', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await doc
      .plan()
      .setTableRowHeights({
        tableId: tbl!.tableId,
        heights: [
          { val: 960, hRule: 'exact' },
          { val: 480, hRule: 'atLeast' },
        ],
      })
      .apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const doc2 = await WordDocument.fromZip(zipBytes, { docId: 'rh-7-rt', mode: 'inline' });
    const tablesRt = await doc2.tables.list();
    expect(tablesRt[0]!.rowHeights[0]).toEqual({ val: 960, hRule: 'exact' });
    expect(tablesRt[0]!.rowHeights[1]).toEqual({ val: 480, hRule: 'atLeast' });
    await doc2.dispose();
  });

  test('7b FlatOpc round-trip → rowHeights 读回值一致', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-7b', mode: 'inline' });

    const [tbl] = await doc.tables.list();
    await doc
      .plan()
      .setTableRowHeights({
        tableId: tbl!.tableId,
        heights: [
          { val: 720, hRule: 'atLeast' },
          { val: 360, hRule: 'auto' },
        ],
      })
      .apply();

    const flatOpc = await doc.toFlatOpc();
    await doc.dispose();

    const doc2 = await WordDocument.fromFlatOpc(flatOpc, { docId: 'rh-7b-rt', mode: 'inline' });
    const tablesRt = await doc2.tables.list();
    expect(tablesRt[0]!.rowHeights[0]).toEqual({ val: 720, hRule: 'atLeast' });
    expect(tablesRt[0]!.rowHeights[1]).toEqual({ val: 360, hRule: 'auto' });
    await doc2.dispose();
  });
});

// ─────────── 9. 不回归：columnWidths 不受影响 ────────────────────────────────

describe('plan.setTableRowHeights 不回归', () => {
  test('9 setTableRowHeights 后 columnWidths 不变', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-9', mode: 'inline' });

    const [tblBefore] = await doc.tables.list();
    const originalWidths = tblBefore!.columnWidths;
    expect(originalWidths).toEqual([2000, 3000]); // fixture 初始值

    await doc
      .plan()
      .setTableRowHeights({
        tableId: tblBefore!.tableId,
        heights: [{ val: 500, hRule: 'atLeast' }, null],
      })
      .apply();

    const [tblAfter] = await doc.tables.list();
    expect(tblAfter!.columnWidths).toEqual([2000, 3000]); // 未变

    await doc.dispose();
  });

  test('9b setTableColumnWidths 后 rowHeights 不变', async () => {
    const buf = await makeFixture(tbl2Row());
    const doc = await WordDocument.fromZip(buf, { docId: 'rh-9b', mode: 'inline' });

    const [tblBefore] = await doc.tables.list();
    // 行 0 有 rowHeight(800, atLeast)，行 1 null
    expect(tblBefore!.rowHeights[0]).toEqual({ val: 800, hRule: 'atLeast' });
    expect(tblBefore!.rowHeights[1]).toBeNull();

    await doc
      .plan()
      .setTableColumnWidths({ tableId: tblBefore!.tableId, widths: [1500, 4000] })
      .apply();

    const [tblAfter] = await doc.tables.list();
    // rowHeights 未被 setTableColumnWidths 碰
    expect(tblAfter!.rowHeights[0]).toEqual({ val: 800, hRule: 'atLeast' });
    expect(tblAfter!.rowHeights[1]).toBeNull();
    // columnWidths 已变
    expect(tblAfter!.columnWidths).toEqual([1500, 4000]);

    await doc.dispose();
  });
});
