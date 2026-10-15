/**
 * TableInfo.cellGrid（span-aware cells）读取测试（task-core-table-span）。
 *
 * 测试矩阵：
 *   1  colspan：某行含 gridSpan=3 格 → 该格 .gridSpan===3，其余 .gridSpan===1
 *   2  rowspan：上格 vMerge=restart、下格 vMerge=continue → 值正确；非合并格 .vMerge===null
 *   3  组合格：同一格 gridSpan>1 且 vMerge=restart → 两属性都正确
 *   4  简单网格：所有格 .gridSpan===1 && .vMerge===null
 *   5  向后兼容：cells 与 cellGrid 同源同序（cells[i][j] === cellGrid[i][j].text）
 *   6  开关：includeCellTexts:false → cellGrid===undefined
 *   7  嵌套表：外层格属性不被嵌套内层 tcPr 污染（gridSpan/vMerge 只读直属 tcPr）
 */

import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { W14_NS, W_NS } from '../namespaces.js';

// ─────────── fixture 基础设施 ─────────────────────────────────────────────────

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

// ─────────── fixture XML 构建器 ───────────────────────────────────────────────

/**
 * 简单 2×2 网格表（无 span）。
 * Row 0: A | B
 * Row 1: C | D
 */
function tblSimple2x2(): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:p w14:paraId="aa000001"><w:r><w:t>A</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="aa000002"><w:r><w:t>B</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:p w14:paraId="aa000003"><w:r><w:t>C</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="aa000004"><w:r><w:t>D</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="aa000099"/>`;
}

/**
 * colspan 测试表：2 行 × 3 列网格。
 * Row 0: cell spanning 3 columns (gridSpan=3) → text="Wide"
 * Row 1: 3 normal cells → texts "X" | "Y" | "Z"
 */
function tblColspan(): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:gridSpan w:val="3"/></w:tcPr>
      <w:p w14:paraId="bb000001"><w:r><w:t>Wide</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:p w14:paraId="bb000002"><w:r><w:t>X</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="bb000003"><w:r><w:t>Y</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="bb000004"><w:r><w:t>Z</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="bb000099"/>`;
}

/**
 * rowspan 测试表：3 行 × 2 列网格。
 * Row 0: vMerge=restart ("Top") | normal ("Right0")
 * Row 1: vMerge=continue ("")    | normal ("Right1")
 * Row 2: normal ("Bot")          | normal ("Right2")
 */
function tblRowspan(): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
      <w:p w14:paraId="cc000001"><w:r><w:t>Top</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="cc000002"><w:r><w:t>Right0</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:tcPr><w:vMerge/></w:tcPr>
      <w:p w14:paraId="cc000003"><w:r><w:t></w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="cc000004"><w:r><w:t>Right1</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:p w14:paraId="cc000005"><w:r><w:t>Bot</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="cc000006"><w:r><w:t>Right2</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="cc000099"/>`;
}

/**
 * 组合格测试表：2 行 × 3 列网格。
 * Row 0: combo cell (gridSpan=2 + vMerge=restart) | normal
 * Row 1: continue cell (gridSpan=2 + vMerge no val) | normal
 */
function tblCombined(): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
    <w:gridCol w:w="2000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:gridSpan w:val="2"/>
        <w:vMerge w:val="restart"/>
      </w:tcPr>
      <w:p w14:paraId="dd000001"><w:r><w:t>Combo</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="dd000002"><w:r><w:t>N0</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:gridSpan w:val="2"/>
        <w:vMerge/>
      </w:tcPr>
      <w:p w14:paraId="dd000003"><w:r><w:t></w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="dd000004"><w:r><w:t>N1</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="dd000099"/>`;
}

/**
 * 嵌套表测试：外层 1×1 表，格内嵌套一张有 gridSpan=2 的表。
 * 外层格自身无 gridSpan / vMerge，不应被嵌套表 tcPr 污染。
 */
function tblNested(): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid>
    <w:gridCol w:w="6000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:p w14:paraId="ee000001"><w:r><w:t>outer</w:t></w:r></w:p>
      <w:tbl>
        <w:tblPr/>
        <w:tblGrid>
          <w:gridCol w:w="2000"/>
          <w:gridCol w:w="2000"/>
        </w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
            <w:p w14:paraId="ee000002"><w:r><w:t>inner-wide</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="ee000099"/>`;
}

/**
 * Bug2 fixture：1 行 × 2 列表，左格纯文字，右格首段后挂一张 inline 图（<w:drawing>）。
 * 右格首段 paraId = ff000002（用于断言 firstParaId）；图段 paraId = ff000003。
 */
function tblCellImage(): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid>
    <w:gridCol w:w="3000"/>
    <w:gridCol w:w="3000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:p w14:paraId="ff000001"><w:r><w:t>L</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:p w14:paraId="ff000002"><w:r><w:t>R</w:t></w:r></w:p>
      <w:p w14:paraId="ff000003"><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="100" cy="100"/></wp:inline></w:drawing></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="ff000099"/>`;
}

/**
 * Bug2 嵌套表 fixture：外层 1×1 表，外格内含「外格图」+ 嵌套表，嵌套表 cell 内含「内格图」。
 * 验证 collectCellDrawingIds 只把**直属本 cell**（非嵌套表 cell）的图算给外格。
 */
function tblNestedCellImage(): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid>
    <w:gridCol w:w="6000"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:p w14:paraId="01000001"><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="100" cy="100"/></wp:inline></w:drawing></w:r></w:p>
      <w:tbl>
        <w:tblPr/>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:p w14:paraId="01000002"><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="50" cy="50"/></wp:inline></w:drawing></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="01000099"/>`;
}

afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── Test 1: colspan ─────────────────────────────────────────────────

describe('cellGrid colspan', () => {
  test('1 colspan: gridSpan=3 格正确；其余格 gridSpan===1', async () => {
    const buf = await makeFixture(tblColspan());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-1', mode: 'inline' });

    const tables = await doc.tables.list({ includeCellTexts: true });
    expect(tables).toHaveLength(1);

    const grid = tables[0]!.cellGrid!;
    expect(grid).toBeDefined();

    // Row 0: 1 physical cell spanning 3 columns
    expect(grid[0]).toHaveLength(1);
    expect(grid[0]![0]!.text).toBe('Wide');
    expect(grid[0]![0]!.gridSpan).toBe(3);
    expect(grid[0]![0]!.vMerge).toBeNull();

    // Row 1: 3 normal cells
    expect(grid[1]).toHaveLength(3);
    expect(grid[1]![0]!.gridSpan).toBe(1);
    expect(grid[1]![1]!.gridSpan).toBe(1);
    expect(grid[1]![2]!.gridSpan).toBe(1);
    expect(grid[1]![0]!.vMerge).toBeNull();

    await doc.dispose();
  });
});

// ─────────── Test 2: rowspan ─────────────────────────────────────────────────

describe('cellGrid rowspan', () => {
  test('2 rowspan: restart/continue/null 三态', async () => {
    const buf = await makeFixture(tblRowspan());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-2', mode: 'inline' });

    const tables = await doc.tables.list({ includeCellTexts: true });
    const grid = tables[0]!.cellGrid!;

    // Row 0 col 0: restart
    expect(grid[0]![0]!.vMerge).toBe('restart');
    expect(grid[0]![0]!.gridSpan).toBe(1);
    expect(grid[0]![0]!.text).toBe('Top');
    // Row 0 col 1: no merge
    expect(grid[0]![1]!.vMerge).toBeNull();

    // Row 1 col 0: continue (vMerge with no val)
    expect(grid[1]![0]!.vMerge).toBe('continue');
    expect(grid[1]![0]!.gridSpan).toBe(1);

    // Row 2: no merges
    expect(grid[2]![0]!.vMerge).toBeNull();
    expect(grid[2]![0]!.text).toBe('Bot');
    expect(grid[2]![1]!.vMerge).toBeNull();

    await doc.dispose();
  });
});

// ─────────── Test 3: combined cell ───────────────────────────────────────────

describe('cellGrid combined cell', () => {
  test('3 组合格: gridSpan>1 且 vMerge=restart 同格两属性都对', async () => {
    const buf = await makeFixture(tblCombined());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-3', mode: 'inline' });

    const tables = await doc.tables.list({ includeCellTexts: true });
    const grid = tables[0]!.cellGrid!;

    // Row 0 col 0: combo — gridSpan=2 + vMerge=restart
    expect(grid[0]![0]!.gridSpan).toBe(2);
    expect(grid[0]![0]!.vMerge).toBe('restart');
    expect(grid[0]![0]!.text).toBe('Combo');

    // Row 0 col 1: normal
    expect(grid[0]![1]!.gridSpan).toBe(1);
    expect(grid[0]![1]!.vMerge).toBeNull();

    // Row 1 col 0: gridSpan=2 + vMerge=continue
    expect(grid[1]![0]!.gridSpan).toBe(2);
    expect(grid[1]![0]!.vMerge).toBe('continue');

    await doc.dispose();
  });
});

// ─────────── Test 4: simple grid ─────────────────────────────────────────────

describe('cellGrid simple grid', () => {
  test('4 简单网格: 所有格 gridSpan===1 && vMerge===null', async () => {
    const buf = await makeFixture(tblSimple2x2());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-4', mode: 'inline' });

    const tables = await doc.tables.list({ includeCellTexts: true });
    const grid = tables[0]!.cellGrid!;

    expect(grid).toHaveLength(2);
    for (const row of grid) {
      expect(row).toHaveLength(2);
      for (const cell of row) {
        expect(cell.gridSpan).toBe(1);
        expect(cell.vMerge).toBeNull();
      }
    }

    await doc.dispose();
  });
});

// ─────────── Test 5: backward compatibility ───────────────────────────────────

describe('cellGrid backward compatibility', () => {
  test('5 cells 与 cellGrid 同源同序：cells[i][j] === cellGrid[i][j].text', async () => {
    const buf = await makeFixture(tblColspan());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-5', mode: 'inline' });

    const tables = await doc.tables.list({ includeCellTexts: true });
    const t = tables[0]!;
    const cells = t.cells!;
    const grid = t.cellGrid!;

    expect(cells).toBeDefined();
    expect(grid).toBeDefined();
    expect(cells.length).toBe(grid.length);
    for (let i = 0; i < cells.length; i++) {
      const row = cells[i]!;
      const gridRow = grid[i]!;
      expect(row.length).toBe(gridRow.length);
      for (let j = 0; j < row.length; j++) {
        expect(row[j]).toBe(gridRow[j]!.text);
      }
    }

    await doc.dispose();
  });

  test('5b cells 值与改前一致（不被 cellGrid 污染）', async () => {
    const buf = await makeFixture(tblRowspan());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-5b', mode: 'inline' });

    const tables = await doc.tables.list({ includeCellTexts: true });
    const cells = tables[0]!.cells!;

    // Row 0: ["Top", "Right0"]
    expect(cells[0]).toEqual(['Top', 'Right0']);
    // Row 1: ["", "Right1"]
    expect(cells[1]).toEqual(['', 'Right1']);
    // Row 2: ["Bot", "Right2"]
    expect(cells[2]).toEqual(['Bot', 'Right2']);

    await doc.dispose();
  });
});

// ─────────── Test 6: includeCellTexts switch ──────────────────────────────────

describe('cellGrid includeCellTexts switch', () => {
  test('6 includeCellTexts:false → cellGrid===undefined', async () => {
    const buf = await makeFixture(tblColspan());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-6', mode: 'inline' });

    const tables = await doc.tables.list({ includeCellTexts: false });
    expect(tables[0]!.cellGrid).toBeUndefined();
    expect(tables[0]!.cells).toBeUndefined();

    await doc.dispose();
  });

  test('6b default (no opts) → cellGrid===undefined', async () => {
    const buf = await makeFixture(tblColspan());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-6b', mode: 'inline' });

    const tables = await doc.tables.list();
    expect(tables[0]!.cellGrid).toBeUndefined();
    expect(tables[0]!.cells).toBeUndefined();

    await doc.dispose();
  });
});

// ─────────── Test 7: nested table isolation ──────────────────────────────────

describe('cellGrid nested table isolation', () => {
  test('7 外层格无 gridSpan/vMerge，不被嵌套表 tcPr 污染', async () => {
    const buf = await makeFixture(tblNested());
    const doc = await WordDocument.fromZip(buf, {
      docId: 'span-7',
      mode: 'inline',
    });

    // includeNested:false → only the outer table
    const tables = await doc.tables.list({ includeCellTexts: true });
    expect(tables).toHaveLength(1);

    const grid = tables[0]!.cellGrid!;
    // Outer table: 1 row, 1 physical cell
    expect(grid).toHaveLength(1);
    expect(grid[0]).toHaveLength(1);

    // The outer cell has no span attributes — must not inherit from nested table
    expect(grid[0]![0]!.gridSpan).toBe(1);
    expect(grid[0]![0]!.vMerge).toBeNull();

    await doc.dispose();
  });
});

// ─────────── Test 8: Bug2 cell firstParaId + imageDrawingIds ──────────────────

describe('cellGrid Bug2 fields (firstParaId + imageDrawingIds)', () => {
  test('8 每个 cell 带首段 paraId；无图 cell imageDrawingIds=[]', async () => {
    const buf = await makeFixture(tblCellImage());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-8', mode: 'inline' });

    const grid = (await doc.tables.list({ includeCellTexts: true }))[0]!.cellGrid!;
    // 左格：首段 paraId = ff000001（大写，indices 统一 upper；但 firstParaId 透传原值）
    expect(grid[0]![0]!.firstParaId?.toUpperCase()).toBe('FF000001');
    expect(grid[0]![0]!.imageDrawingIds).toEqual([]);
    // 右格：首段 paraId = ff000002（注意是首段，不是图段 ff000003）
    expect(grid[0]![1]!.firstParaId?.toUpperCase()).toBe('FF000002');

    await doc.dispose();
  });

  test('8b 含图 cell 收录该 cell 内 drawingId（与 drawings.list 同 ID 空间）', async () => {
    const buf = await makeFixture(tblCellImage());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-8b', mode: 'inline' });

    const grid = (await doc.tables.list({ includeCellTexts: true }))[0]!.cellGrid!;
    const ids = grid[0]![1]!.imageDrawingIds;
    expect(ids).toHaveLength(1);
    // 与 drawings.list 同 ID 空间，且能反查到 picture 信息。
    const drawings = await doc.drawings.list({ scope: 'body' });
    expect(drawings.map((d) => d.drawingId)).toContain(ids[0]);

    await doc.dispose();
  });

  test('8c 嵌套表 cell 内的图不上浮污染外层 cell（只算直属本 cell 的图）', async () => {
    const buf = await makeFixture(tblNestedCellImage());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-8c', mode: 'inline' });

    // includeNested:false → 只返回外层表；外格只应有「外格图」1 张，不含嵌套表内的图。
    const grid = (await doc.tables.list({ includeCellTexts: true }))[0]!.cellGrid!;
    expect(grid[0]![0]!.imageDrawingIds).toHaveLength(1);

    // 全文档共 2 张 drawing（外格图 + 内格图）；外格 imageDrawingIds 只占其一。
    const allDrawings = await doc.drawings.list({ scope: 'body' });
    expect(allDrawings.length).toBe(2);

    await doc.dispose();
  });
});

// 占位（保持原文件 describe 收尾结构对齐）
describe('cellGrid Bug2 switch off', () => {
  test('8d includeCellTexts:false → 不计算 imageDrawingIds（cellGrid undefined）', async () => {
    const buf = await makeFixture(tblCellImage());
    const doc = await WordDocument.fromZip(buf, { docId: 'span-8d', mode: 'inline' });
    expect((await doc.tables.list())[0]!.cellGrid).toBeUndefined();
    await doc.dispose();
  });
});
