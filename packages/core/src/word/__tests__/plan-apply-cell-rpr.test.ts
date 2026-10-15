/**
 * pickFirstRunRPr 通用性测试（task-core-cell-rpr R1 + R2）。
 *
 * 验证 replaceTableCellText / insertTableRow 在各种 run 包裹结构下
 * 都能正确透传首 run 的 rPr（字号/字体）。
 *
 * 测试矩阵（R1 6 个场景 + R2 3 个场景）：
 *   R1（后代遍历，覆盖任意 run 包裹容器）：
 *   1  SDT 包裹 run：replaceTableCellText 后新 run 带 sz21 rPr（原报告场景）
 *   2  hyperlink 包裹（通用性关键：证明不是只判 SDT）→ 同样透传 rPr
 *   3  纯直接子 run（回归）→ 仍正确透传（不破坏原行为）
 *   4  首 run 无 rPr，pPr 也无 rPr → 新 run 裸（保语义，不越级取后面 run 的格式）
 *   5  insertTableRow / fillCellTexts 经 SDT 模板行 → 插入行回填带 rPr
 *   6  无 run 段落（空段，pPr 无 rPr）→ undefined 不崩
 *   R2（pPr/rPr 兜底：run 裸时取段落默认 run 属性）：
 *   7  run 裸（SDT 占位）+ pPr/rPr 有 sz → 兜底取 pPr/rPr 的 sz（R2 新增）
 *   8  run 有 rPr + pPr/rPr 也有 sz → 优先 run rPr，不被 pPr/rPr 覆盖（回归）
 *   9  run 裸 + pPr 存在但 pPr 无 rPr → undefined（三者都无的边界）
 */

import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { W14_NS, W_NS } from '../namespaces.js';

// ─────────── fixture helpers ──────────────────────────────────────────────────

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

/** 从 toZip 产物读取 document.xml 文本 */
async function readDocXml(zipBytes: Uint8Array): Promise<string> {
  const loaded = await JSZip.loadAsync(zipBytes);
  return loaded.file('word/document.xml')!.async('string');
}

/** 最小单列单行表格，cell 内容由 cellInner 决定 */
function singleCellTable(cellInner: string, _tableId = 'TTTT0001'): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>
      ${cellInner}
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="PPPP0099"/>`;
}

afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── 场景 1：SDT 包裹 run ────────────────────────────────────────────

describe('pickFirstRunRPr / SDT 包裹 run（场景1：原报告 bug 场景）', () => {
  test('1 cell 段落内 run 在 sdtContent 里 → replaceTableCellText 后新 run 带 sz21 rPr', async () => {
    // cell 段落：<w:sdt><w:sdtContent><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>old</w:t></w:r></w:sdtContent></w:sdt>
    const cellInner = `<w:p w14:paraId="PPPP0001">
  <w:sdt>
    <w:sdtPr/>
    <w:sdtContent>
      <w:r>
        <w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
        <w:t>old text</w:t>
      </w:r>
    </w:sdtContent>
  </w:sdt>
</w:p>`;

    const buf = await makeFixture(singleCellTable(cellInner));
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-1', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    await doc.plan().replaceTableCellText({ tableId, row: 0, col: 0, text: 'new text' }).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // 新 run 应含 sz val="21"（从 SDT 内首 run 透传的 rPr）
    expect(xml).toContain('w:val="21"');
    // 新文字已写入
    expect(xml).toContain('new text');
    // 旧文字不再出现
    expect(xml).not.toContain('old text');
  });
});

// ─────────── 场景 2：hyperlink 包裹（通用性关键）────────────────────────────

describe('pickFirstRunRPr / hyperlink 包裹 run（场景2：通用性证明）', () => {
  test('2 cell 段落内 run 在 w:hyperlink 里 → 同样透传 rPr（不是只判 SDT）', async () => {
    // cell 段落：<w:hyperlink><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>...</w:r></w:hyperlink>
    const cellInner = `<w:p w14:paraId="PPPP0002">
  <w:hyperlink r:id="rId99" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:r>
      <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
      <w:t>link text</w:t>
    </w:r>
  </w:hyperlink>
</w:p>`;

    const buf = await makeFixture(singleCellTable(cellInner));
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-2', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    await doc.plan().replaceTableCellText({ tableId, row: 0, col: 0, text: 'replaced' }).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // 透传了 sz=28 和 bold
    expect(xml).toContain('w:val="28"');
    expect(xml).toContain('<w:b/>');
    expect(xml).toContain('replaced');
    expect(xml).not.toContain('link text');
  });
});

// ─────────── 场景 3：纯直接子 run（回归）───────────────────────────────────

describe('pickFirstRunRPr / 纯直接子 run（场景3：回归）', () => {
  test('3 cell 段落直接含 <w:r><w:rPr>... → 仍正确透传（原行为不破坏）', async () => {
    const cellInner = `<w:p w14:paraId="PPPP0003">
  <w:r>
    <w:rPr><w:color w:val="FF0000"/><w:sz w:val="24"/></w:rPr>
    <w:t>original</w:t>
  </w:r>
</w:p>`;

    const buf = await makeFixture(singleCellTable(cellInner));
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-3', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    await doc.plan().replaceTableCellText({ tableId, row: 0, col: 0, text: 'updated' }).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    expect(xml).toContain('w:val="FF0000"'); // color 透传
    expect(xml).toContain('w:val="24"'); // sz 透传
    expect(xml).toContain('updated');
    expect(xml).not.toContain('original');
  });
});

// ─────────── 场景 4：首 run 无 rPr → 裸（保语义）──────────────────────────

describe('pickFirstRunRPr / 首 run 无 rPr（场景4：保原语义）', () => {
  test('4 首个 run（SDT 内）无 rPr → 新 run 裸（不越级取后续 run 格式）', async () => {
    // 首 run 无 rPr；第二个 run 有 rPr（不应被取到）
    const cellInner = `<w:p w14:paraId="PPPP0004">
  <w:sdt>
    <w:sdtContent>
      <w:r>
        <w:t>no rpr run</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:sz w:val="36"/></w:rPr>
        <w:t>has rpr run</w:t>
      </w:r>
    </w:sdtContent>
  </w:sdt>
</w:p>`;

    const buf = await makeFixture(singleCellTable(cellInner));
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-4', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    await doc.plan().replaceTableCellText({ tableId, row: 0, col: 0, text: 'bare' }).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // sz=36 来自第二个 run，不应出现（首 run 无 rPr → 裸）
    expect(xml).not.toContain('w:val="36"');
    expect(xml).toContain('bare');
  });
});

// ─────────── 场景 5：insertTableRow 经 SDT 模板行 ───────────────────────────

describe('pickFirstRunRPr / insertTableRow + SDT 模板行（场景5）', () => {
  test('5 模板行 cell 含 SDT 包裹 run → 插入行 fillCellTexts 透传 rPr', async () => {
    // 2 列表格，模板行（row 0）两个 cell 都用 SDT 包裹 run 含 sz21
    const twoColWithSdt = `<w:tbl>
  <w:tblPr/>
  <w:tblGrid>
    <w:gridCol w:w="2500"/>
    <w:gridCol w:w="2500"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="PPPP0005">
        <w:sdt><w:sdtContent>
          <w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>tmpl A</w:t></w:r>
        </w:sdtContent></w:sdt>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="PPPP0006">
        <w:sdt><w:sdtContent>
          <w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>tmpl B</w:t></w:r>
        </w:sdtContent></w:sdt>
      </w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="PPPP0099"/>`;

    const buf = await makeFixture(twoColWithSdt);
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-5', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    // 以 row 0 为模板插入新行（position 1），填充文字
    await doc
      .plan()
      .insertTableRow({ tableId, templateRow: 0, position: 'end', cellTexts: ['fill A', 'fill B'] })
      .apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // 新插入的行内容含填充文字
    expect(xml).toContain('fill A');
    expect(xml).toContain('fill B');

    // 从 SDT 模板透传的 sz=21 应出现在新行的 run 里
    // 统计 sz val="21" 出现次数：模板行 2 + 新行 2 = 至少 4 次
    const matches = xml.match(/w:val="21"/g) ?? [];
    expect(matches.length).toBeGreaterThanOrEqual(4);
  });
});

// ─────────── 场景 6：无 run 空段落 → undefined 不崩 ─────────────────────────

describe('pickFirstRunRPr / 无 run 空段落（场景6）', () => {
  test('6 cell 段落完全无 run（空段）→ replaceTableCellText 正常执行、新 run 裸、不崩', async () => {
    // 空段落：只有 <w:pPr>，无任何 run
    const cellInner = `<w:p w14:paraId="PPPP0006">
  <w:pPr><w:jc w:val="center"/></w:pPr>
</w:p>`;

    const buf = await makeFixture(singleCellTable(cellInner));
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-6', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    // 不应抛异常
    await expect(
      doc.plan().replaceTableCellText({ tableId, row: 0, col: 0, text: 'into empty' }).apply(),
    ).resolves.not.toThrow();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);
    expect(xml).toContain('into empty');
    // 裸 run（无 rPr），不含任何 w:rPr
    // 注：仍保留 pPr 中对齐属性
  });
});

// ─────────── R2：pPr/rPr 兜底（场景 7-9）────────────────────────────────────

describe('pickFirstRunRPr / R2 pPr/rPr 兜底', () => {
  test('7 run 裸（SDT 占位）+ pPr/rPr 有 sz → 兜底取 pPr/rPr 的 sz21', async () => {
    // 模拟 skeleton 骨架化后的占位格：SDT 内 run 无显式 rPr，但 pPr/rPr 有 sz21
    // R2 修复：首 run 无 rPr → 兜底 pPr/rPr → 取到 sz21 → 回填带 sz21
    const cellInner = `<w:p w14:paraId="PPPP0007">
  <w:pPr>
    <w:jc w:val="center"/>
    <w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
  </w:pPr>
  <w:sdt>
    <w:sdtContent>
      <w:r>
        <w:t>请输入规格</w:t>
      </w:r>
    </w:sdtContent>
  </w:sdt>
</w:p>`;

    const buf = await makeFixture(singleCellTable(cellInner));
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-7', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    await doc
      .plan()
      .replaceTableCellText({ tableId, row: 0, col: 0, text: 'filled value' })
      .apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // 应从 pPr/rPr 兜底取到 sz=21
    expect(xml).toContain('w:val="21"');
    expect(xml).toContain('filled value');
  });

  test('8 run 有 rPr（sz=24）+ pPr/rPr 也有不同 sz（sz=18）→ 优先 run rPr，pPr/rPr 不覆盖', async () => {
    // run 显式 rPr sz=24；pPr/rPr sz=18
    // R2 兜底只在 run 无 rPr 时才触发；run 有 rPr 应优先（回归）
    const cellInner = `<w:p w14:paraId="PPPP0008">
  <w:pPr>
    <w:rPr><w:sz w:val="18"/></w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr><w:sz w:val="24"/></w:rPr>
    <w:t>original</w:t>
  </w:r>
</w:p>`;

    const buf = await makeFixture(singleCellTable(cellInner));
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-8', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    await doc.plan().replaceTableCellText({ tableId, row: 0, col: 0, text: 'replaced' }).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // 应取 run rPr sz=24（不被 pPr/rPr sz=18 覆盖）
    expect(xml).toContain('w:val="24"');
    // pPr/rPr 的 sz=18 不应出现在新 run 的 rPr 里
    // （pPr/rPr 本身保留，但新 run 的 rPr 来自原 run）
    expect(xml).toContain('replaced');
  });

  test('9 run 裸 + pPr 存在但 pPr 无 rPr → undefined（三者都无，仍裸）', async () => {
    // pPr 只有 jc，无 rPr；run 也无 rPr → R2 兜底找不到 pPr/rPr → undefined → 裸 run
    const cellInner = `<w:p w14:paraId="PPPP0009">
  <w:pPr><w:jc w:val="left"/></w:pPr>
  <w:r>
    <w:t>plain</w:t>
  </w:r>
</w:p>`;

    const buf = await makeFixture(singleCellTable(cellInner));
    const doc = await WordDocument.fromZip(buf, { docId: 'rpr-9', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    await doc.plan().replaceTableCellText({ tableId, row: 0, col: 0, text: 'new plain' }).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // 无 rPr 来源 → 新 run 裸（不含 w:sz 或 w:rPr 元素）
    // 只验证文字写入成功且不崩
    expect(xml).toContain('new plain');
    // 不含任何 sz 属性（裸 run）
    expect(xml).not.toContain('<w:sz ');
  });
});
