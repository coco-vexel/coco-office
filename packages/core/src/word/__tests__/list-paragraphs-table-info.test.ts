/**
 * listParagraphs — includeTableInfo 选项测试（task-20260625-para-table-info）。
 *
 * 覆盖场景：
 *   1. 含表格文档：表格内段落 includeTableInfo:true → table.tableId 与 listTables 返回的 tableId 一致；
 *      row/col 正确。
 *   2. 表格外正文段落：table 字段不出现。
 *   3. includeTableInfo:false（或不传）：所有段落都无 table 字段。
 *   4. 嵌套表格：报最内层 tableId / row / col。
 *   5. tableId 与 buildTableIndex / buildIndices 分配的 id 完全一致（反查路径验证）。
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { listParagraphs } from '../list-paragraphs.js';
import { buildIndices } from '../indices.js';
import { listTables } from '../list-tables.js';
import { W14_NS, W_NS } from '../namespaces.js';

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
 * 构建含一张 2×2 表格的文档：
 *
 *   表格外正文段 (BODYPARA)
 *   [Table 1]
 *     row0: tc(0,0) → paraId=TC00PARA | tc(0,1) → paraId=TC01PARA
 *     row1: tc(1,0) → paraId=TC10PARA | tc(1,1) → paraId=TC11PARA
 *   表格后正文段 (BODYPAR2)
 */
function makeDocWithTable(): Package {
  const pkg = createPackage();
  putXmlPart(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
      <w:body>
        <w:p w14:paraId="BODYPARA"><w:r><w:t>正文段落</w:t></w:r></w:p>
        <w:tbl>
          <w:tr>
            <w:tc><w:p w14:paraId="TC00PARA"><w:r><w:t>r0c0</w:t></w:r></w:p></w:tc>
            <w:tc><w:p w14:paraId="TC01PARA"><w:r><w:t>r0c1</w:t></w:r></w:p></w:tc>
          </w:tr>
          <w:tr>
            <w:tc><w:p w14:paraId="TC10PARA"><w:r><w:t>r1c0</w:t></w:r></w:p></w:tc>
            <w:tc><w:p w14:paraId="TC11PARA"><w:r><w:t>r1c1</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
        <w:p w14:paraId="BODYPAR2"><w:r><w:t>表格后正文</w:t></w:r></w:p>
      </w:body>
    </w:document>`,
  );
  return pkg;
}

// ─────────── 场景 1：表格内段落 tableId/row/col 正确 ───────────

describe('listParagraphs — includeTableInfo: 表格内段落', () => {
  test('table.tableId 与 listTables() 返回的 tableId 一致', () => {
    const pkg = makeDocWithTable();

    // 先通过 listTables 拿到官方 tableId
    const indices = buildIndices(pkg);
    const tables = listTables(pkg, indices, { scope: 'body' });
    expect(tables).toHaveLength(1);
    const officialTableId = tables[0]!.tableId;
    expect(officialTableId).toMatch(/^tbl-\d{4}$/);

    // 再通过 listParagraphs includeTableInfo 拿段落的 tableId
    const result = listParagraphs(pkg, {
      pagination: { mode: 'all' },
      includeTableInfo: true,
    });

    const byParaId = new Map(result.paragraphs.map((p) => [p.paraId, p]));

    // 表格内四段都应该有 table 字段，且 tableId 与 listTables 一致
    for (const pid of ['TC00PARA', 'TC01PARA', 'TC10PARA', 'TC11PARA']) {
      const para = byParaId.get(pid);
      expect(para, `段落 ${pid} 应存在`).toBeDefined();
      expect(para!.table, `段落 ${pid} 应有 table 字段`).toBeDefined();
      expect(para!.table!.tableId).toBe(officialTableId);
    }
  });

  test('row/col 序号正确（0-based 物理序号）', () => {
    const pkg = makeDocWithTable();
    const result = listParagraphs(pkg, {
      pagination: { mode: 'all' },
      includeTableInfo: true,
    });
    const byParaId = new Map(result.paragraphs.map((p) => [p.paraId, p]));

    // row 0, col 0
    expect(byParaId.get('TC00PARA')!.table).toEqual(expect.objectContaining({ row: 0, col: 0 }));
    // row 0, col 1
    expect(byParaId.get('TC01PARA')!.table).toEqual(expect.objectContaining({ row: 0, col: 1 }));
    // row 1, col 0
    expect(byParaId.get('TC10PARA')!.table).toEqual(expect.objectContaining({ row: 1, col: 0 }));
    // row 1, col 1
    expect(byParaId.get('TC11PARA')!.table).toEqual(expect.objectContaining({ row: 1, col: 1 }));
  });
});

// ─────────── 场景 2：表格外段落无 table 字段 ───────────

describe('listParagraphs — includeTableInfo: 表格外段落', () => {
  test('正文段落 table 字段不出现', () => {
    const pkg = makeDocWithTable();
    const result = listParagraphs(pkg, {
      pagination: { mode: 'all' },
      includeTableInfo: true,
    });
    const byParaId = new Map(result.paragraphs.map((p) => [p.paraId, p]));

    // 表格前后的正文段落不应有 table 字段
    expect(byParaId.get('BODYPARA')!.table).toBeUndefined();
    expect(byParaId.get('BODYPAR2')!.table).toBeUndefined();
  });
});

// ─────────── 场景 3：includeTableInfo:false 或不传 → 无 table 字段 ───────────

describe('listParagraphs — includeTableInfo: 关闭时零开销', () => {
  test('不传 includeTableInfo：所有段落无 table 字段', () => {
    const pkg = makeDocWithTable();
    const result = listParagraphs(pkg, { pagination: { mode: 'all' } });
    for (const para of result.paragraphs) {
      expect(Object.prototype.hasOwnProperty.call(para, 'table')).toBe(false);
    }
  });

  test('includeTableInfo:false：所有段落无 table 字段', () => {
    const pkg = makeDocWithTable();
    const result = listParagraphs(pkg, {
      pagination: { mode: 'all' },
      includeTableInfo: false,
    });
    for (const para of result.paragraphs) {
      expect(Object.prototype.hasOwnProperty.call(para, 'table')).toBe(false);
    }
  });
});

// ─────────── 场景 4：嵌套表格 — 报最内层 ───────────

describe('listParagraphs — includeTableInfo: 嵌套表格', () => {
  test('嵌套表格内段落报最内层 tableId', () => {
    /**
     * 外层表格（tbl-0001）：1×1
     *   cell 内嵌套内层表格（tbl-0002）：1×1
     *     内层 cell 的段落 → INNERPAR
     * 外层 cell 还有一个非嵌套段落 → OUTERPAR
     */
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:tbl>
            <w:tr>
              <w:tc>
                <w:p w14:paraId="OUTERPAR"><w:r><w:t>外层段落</w:t></w:r></w:p>
                <w:tbl>
                  <w:tr>
                    <w:tc>
                      <w:p w14:paraId="INNERPAR"><w:r><w:t>内层段落</w:t></w:r></w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>`,
    );

    // 获取官方 tableId（外层=tbl-0001，内层=tbl-0002，按 buildTableIndex 文档顺序）
    const indices = buildIndices(pkg);
    const tables = listTables(pkg, indices, { scope: 'body', includeNested: true });
    expect(tables).toHaveLength(2);
    // 外层表格先遇到，所以 index 较小
    const outerTable = tables.find((t) => !t.isNested);
    const innerTable = tables.find((t) => t.isNested);
    expect(outerTable).toBeDefined();
    expect(innerTable).toBeDefined();

    const result = listParagraphs(pkg, {
      pagination: { mode: 'all' },
      includeTableInfo: true,
    });
    const byParaId = new Map(result.paragraphs.map((p) => [p.paraId, p]));

    // OUTERPAR 在外层 tc 直接子段（不在嵌套表内），应报外层 tableId
    const outerPara = byParaId.get('OUTERPAR');
    expect(outerPara!.table).toBeDefined();
    expect(outerPara!.table!.tableId).toBe(outerTable!.tableId);
    expect(outerPara!.table!.row).toBe(0);
    expect(outerPara!.table!.col).toBe(0);

    // INNERPAR 在内层 tc 内，应报最内层 tableId（innerTable）
    const innerPara = byParaId.get('INNERPAR');
    expect(innerPara!.table).toBeDefined();
    expect(innerPara!.table!.tableId).toBe(innerTable!.tableId);
    expect(innerPara!.table!.row).toBe(0);
    expect(innerPara!.table!.col).toBe(0);
  });
});

// ─────────── 场景 5：与 includeFormat 独立，两者可同时开启 ───────────

describe('listParagraphs — includeTableInfo 与 includeFormat 独立', () => {
  test('两者同时开启：format 和 table 都有值', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
        <w:body>
          <w:tbl>
            <w:tr>
              <w:tc>
                <w:p w14:paraId="FMTPARA">
                  <w:pPr><w:jc w:val="center"/></w:pPr>
                  <w:r><w:t>格式+表格</w:t></w:r>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>`,
    );

    const result = listParagraphs(pkg, {
      pagination: { mode: 'all' },
      includeTableInfo: true,
      includeFormat: true,
    });
    const para = result.paragraphs.find((p) => p.paraId === 'FMTPARA');
    expect(para).toBeDefined();
    // 两个字段都有值
    expect(para!.table).toBeDefined();
    expect(para!.table!.tableId).toMatch(/^tbl-/);
    expect(para!.format).toBeDefined();
    expect(para!.format!.alignment).toBe('center');
  });
});
