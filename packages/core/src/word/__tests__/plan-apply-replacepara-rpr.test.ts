/**
 * replaceParagraph rPr 透传测试（task-core-replacepara-rpr）。
 *
 * 验证 executeReplaceParagraph（非 trackChanges + trackChanges 两路径）在各种
 * 格式来源下都能正确透传首 run 的 rPr（字号/字体），修复全文正文字号变大的 bug。
 *
 * 测试矩阵（6 个场景）：
 *   1  基本：段落首 run 带 rPr（sz24）→ replaceParagraph 后新 run 带 sz24
 *   2  SDT/包裹 run：首 run 在 sdtContent 里 → 同样透传（后代遍历）
 *   3  pPr/rPr 兜底：首 run 裸 + pPr/rPr 有 sz → 新 run 取 pPr/rPr 的 sz
 *   4  trackChanges 路径：新 ins run 也带 rPr
 *   5  边界：段落无 run（空段）→ 新 run 裸（不崩）
 *   6  回归：pickFirstRunRPr 提到 _shared 后，replaceTableCellText 行为不变
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

async function readDocXml(zipBytes: Uint8Array): Promise<string> {
  const loaded = await JSZip.loadAsync(zipBytes);
  return loaded.file('word/document.xml')!.async('string');
}

afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── 场景 1：基本带 rPr 透传 ─────────────────────────────────────────

describe('replaceParagraph rPr / 基本带 rPr 透传（场景1）', () => {
  test('1 段落首 run 带 sz24 + 字体 → replaceParagraph 后新 run 带同 rPr', async () => {
    const paraInner = `<w:p w14:paraId="PPPP0001">
  <w:r>
    <w:rPr><w:rFonts w:ascii="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
    <w:t>原始长段落内容</w:t>
  </w:r>
</w:p>`;

    const buf = await makeFixture(paraInner);
    const doc = await WordDocument.fromZip(buf, { docId: 'rp-1', mode: 'inline' });

    await doc.plan().replaceParagraph('PPPP0001', '截断后的内容').apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // 新 run 应带 sz24（从原首 run 透传）
    expect(xml).toContain('w:val="24"');
    // 字体也透传
    expect(xml).toContain('宋体');
    // 新文字写入
    expect(xml).toContain('截断后的内容');
    // 旧文字不再出现
    expect(xml).not.toContain('原始长段落内容');
  });
});

// ─────────── 场景 2：SDT/包裹 run（后代遍历）───────────────────────────────

describe('replaceParagraph rPr / SDT 包裹 run（场景2）', () => {
  test('2 首 run 在 sdtContent 里带 sz21 → 透传到新 run（后代遍历覆盖包裹容器）', async () => {
    const paraInner = `<w:p w14:paraId="PPPP0002">
  <w:sdt>
    <w:sdtContent>
      <w:r>
        <w:rPr><w:sz w:val="21"/></w:rPr>
        <w:t>SDT 内内容</w:t>
      </w:r>
    </w:sdtContent>
  </w:sdt>
</w:p>`;

    const buf = await makeFixture(paraInner);
    const doc = await WordDocument.fromZip(buf, { docId: 'rp-2', mode: 'inline' });

    await doc.plan().replaceParagraph('PPPP0002', '新内容').apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // SDT 内首 run 的 sz21 透传到新 run
    expect(xml).toContain('w:val="21"');
    expect(xml).toContain('新内容');
  });
});

// ─────────── 场景 3：pPr/rPr 兜底 ──────────────────────────────────────────

describe('replaceParagraph rPr / pPr/rPr 兜底（场景3）', () => {
  test('3 首 run 裸（无 rPr）+ pPr/rPr 有 sz21 → 新 run 取 pPr/rPr 的 sz21', async () => {
    const paraInner = `<w:p w14:paraId="PPPP0003">
  <w:pPr>
    <w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
  </w:pPr>
  <w:r>
    <w:t>裸 run 无 rPr 的段落</w:t>
  </w:r>
</w:p>`;

    const buf = await makeFixture(paraInner);
    const doc = await WordDocument.fromZip(buf, { docId: 'rp-3', mode: 'inline' });

    await doc.plan().replaceParagraph('PPPP0003', '填充内容').apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // pPr/rPr 兜底取到 sz21
    expect(xml).toContain('w:val="21"');
    expect(xml).toContain('填充内容');
  });
});

// ─────────── 场景 4：trackChanges 路径 ─────────────────────────────────────

describe('replaceParagraph rPr / trackChanges 路径（场景4）', () => {
  test('4 trackChanges=true → 新 ins run 也带原 rPr（sz24）', async () => {
    const paraInner = `<w:p w14:paraId="PPPP0004">
  <w:r>
    <w:rPr><w:sz w:val="24"/><w:color w:val="0000FF"/></w:rPr>
    <w:t>原始内容</w:t>
  </w:r>
</w:p>`;

    const buf = await makeFixture(paraInner);
    const doc = await WordDocument.fromZip(buf, { docId: 'rp-4', mode: 'inline' });

    await doc
      .plan({ trackChanges: true, author: 'TestAuthor', date: new Date('2026-01-01T00:00:00Z') })
      .replaceParagraph('PPPP0004', '修订后内容')
      .apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // ins 元素存在（trackChanges 路径）
    expect(xml).toContain('<w:ins');
    expect(xml).toContain('修订后内容');
    // 新 run 在 ins 内，且带 sz24
    expect(xml).toContain('w:val="24"');
  });
});

// ─────────── 场景 5：边界——空段落 ───────────────────────────────────────────

describe('replaceParagraph rPr / 空段落（场景5）', () => {
  test('5 段落无 run（空段）→ replaceParagraph 正常执行、新 run 裸、不崩', async () => {
    const paraInner = `<w:p w14:paraId="PPPP0005">
  <w:pPr><w:jc w:val="center"/></w:pPr>
</w:p>`;

    const buf = await makeFixture(paraInner);
    const doc = await WordDocument.fromZip(buf, { docId: 'rp-5', mode: 'inline' });

    // 不应抛异常
    await expect(
      doc.plan().replaceParagraph('PPPP0005', '填入内容').apply(),
    ).resolves.not.toThrow();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);
    expect(xml).toContain('填入内容');
    // 空段无 rPr 来源 → 裸 run（无 sz 属性）
    expect(xml).not.toContain('<w:sz ');
  });
});

// ─────────── 场景 6：回归——replaceTableCellText 共用 helper 行为不变 ─────────

describe('replaceParagraph rPr / 回归——replaceTableCellText 提取后不变（场景6）', () => {
  test('6 replaceTableCellText 仍正确透传 rPr（pickFirstRunRPr 提到 _shared 后回归）', async () => {
    // 简单表格，cell 段落 run 带 sz20
    const tblBody = `<w:tbl>
  <w:tblPr/>
  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>
      <w:p w14:paraId="PPPP0006">
        <w:r>
          <w:rPr><w:sz w:val="20"/></w:rPr>
          <w:t>原格内容</w:t>
        </w:r>
      </w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:p w14:paraId="PPPP0007"/>`;

    const buf = await makeFixture(tblBody);
    const doc = await WordDocument.fromZip(buf, { docId: 'rp-6', mode: 'inline' });

    const tables = await doc.tables.list();
    const tableId = tables[0]!.tableId;

    await doc.plan().replaceTableCellText({ tableId, row: 0, col: 0, text: '新格内容' }).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const xml = await readDocXml(zipBytes);

    // sz20 透传（pickFirstRunRPr 共用实现）
    expect(xml).toContain('w:val="20"');
    expect(xml).toContain('新格内容');
    expect(xml).not.toContain('原格内容');
  });
});
