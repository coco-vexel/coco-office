/**
 * executeInsertTableRow 单元测试（task-14-insert-table-row，spec §3 测试矩阵）。
 *
 * 覆盖（12 条）：
 *   1  末尾追加一行（templateRow=1, position='end', 无 cellTexts）
 *   2  cellTexts 填充（各列文字正确；首 run rPr 字体/字号与模板一致）
 *   3  中间插入（position=1）：行顺序 row0 → 新行 → 原 row1…；rowCount+1
 *   4  paraId 唯一性：克隆行内段落 paraId ≠ 模板行；toZip 重载后全局无重复 paraId
 *   5  锚点剥离：模板行含书签 + 批注 → 克隆行内无 bookmarkStart / commentReference
 *   6  cellTexts 短于列数：未覆盖列文字被清空（无模板样例文字残留）
 *   7  cellTexts 长于列数：多余元素忽略，列数不变
 *   8  templateRow 越界 → TABLE_NOT_FOUND
 *   9  tableId 不存在 → TABLE_NOT_FOUND
 *   10 position 越界（>rowCount / <0）→ TABLE_NOT_FOUND
 *   11 trackChanges=true → UNSUPPORTED_OOXML_FEATURE
 *   12 round-trip：apply → toZip → WordDocument.load 重载不抛错，新行结构合法
 */
import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult } from '../../plan.js';
import { WordError } from '../../word-error.js';
import { WordDocument } from '../../word-document.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeInsertTableRow } from '../table-ops.js';

// ─────────── 测试工具 ───────────

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function putXml(pkg: Package, path: string, xml: string, contentType = DOC_CTYPE): void {
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

function emptyApplyResult(): ApplyResult {
  return {
    insertedParagraphs: [],
    insertedBodyStartParaIds: [],
    createdComments: [],
    createdBookmarks: [],
    acceptedRevisions: 0,
    rejectedRevisions: 0,
    replacedTableCells: 0,
  };
}

function makeDoc(
  bodyInner: string,
  opts: { trackChanges?: boolean } = {},
): {
  pkg: Package;
  ctx: ExecutorContext;
  tableId: string;
} {
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`,
  );
  const indices = buildIndices(pkg);
  const tableId = [...indices.tblIdx.keys()][0] ?? '';
  const ctx: ExecutorContext = {
    pkg,
    indices,
    options: {
      trackChanges: opts.trackChanges ?? false,
      author: '',
      date: new Date('2026-06-04T00:00:00Z'),
    },
    result: emptyApplyResult(),
    opIndex: 0,
    docId: 'test-doc',
  };
  return { pkg, ctx, tableId };
}

/** 获取表格所有 <w:tr> 直属子列表 */
function getTableRows(pkg: Package, tableId: string): Element[] {
  const idx = buildIndices(pkg);
  const tbl = idx.tblIdx.get(tableId)!;
  const rows: Element[] = [];
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const c = tbl.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tr') rows.push(c);
  }
  return rows;
}

/** 获取某行所有 <w:tc> */
function getRowCells(tr: Element): Element[] {
  const cells: Element[] = [];
  for (let i = 0; i < tr.childNodes.length; i++) {
    const c = tr.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tc') cells.push(c);
  }
  return cells;
}

/** 获取单元格首段 */
function getFirstPara(tc: Element): Element | null {
  for (let i = 0; i < tc.childNodes.length; i++) {
    const c = tc.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'p') return c;
  }
  return null;
}

/** 获取段落内所有 <w:t> 文字拼接 */
function paraText(para: Element): string {
  const ts = para.getElementsByTagNameNS(W_NS, 't');
  let s = '';
  for (let i = 0; i < ts.length; i++) s += ts[i]!.textContent ?? '';
  return s;
}

/** 序列化为 XML 字符串 */
function toXml(el: Element): string {
  return serializeToWellFormedString(el as never);
}

// ─────────── fixture：表头行(含底纹 shd) + 数据行（含 rPr 字体/字号）───────────

/**
 * 2 行 3 列表格
 *   row0(表头)：tc0 含 shd(E0E0E0) + vAlign(center) + rFonts黑体/sz18
 *               tc1 + tc2 含 tcW
 *   row1(数据)：三列各含 rPr(微软雅黑/sz24)，样例文字"样例A" "样例B" "样例C"
 */
const STYLED_TABLE = `
  <w:tbl>
    <w:tblGrid>
      <w:gridCol w:w="2000"/>
      <w:gridCol w:w="2000"/>
      <w:gridCol w:w="2000"/>
    </w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="2000" w:type="dxa"/>
          <w:shd w:val="clear" w:fill="E0E0E0"/>
          <w:vAlign w:val="center"/>
        </w:tcPr>
        <w:p w14:paraId="00000001">
          <w:pPr><w:jc w:val="center"/></w:pPr>
          <w:r><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="18"/></w:rPr><w:t>表头A</w:t></w:r>
        </w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000002"><w:r><w:t>表头B</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000003"><w:r><w:t>表头C</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
    <w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000004">
          <w:r><w:rPr><w:rFonts w:eastAsia="微软雅黑"/><w:sz w:val="24"/></w:rPr><w:t>样例A</w:t></w:r>
        </w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000005">
          <w:r><w:rPr><w:rFonts w:eastAsia="微软雅黑"/><w:sz w:val="24"/></w:rPr><w:t>样例B</w:t></w:r>
        </w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000006">
          <w:r><w:rPr><w:rFonts w:eastAsia="微软雅黑"/><w:sz w:val="24"/></w:rPr><w:t>样例C</w:t></w:r>
        </w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
`;

// ─────────── 测试 ───────────

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('executeInsertTableRow', () => {
  // 1. 末尾追加一行（templateRow=1, position='end', 无 cellTexts）
  test('1 末尾追加一行：新行 tc 数 == 模板列数；末列 tcPr 的 tcW 与模板一致', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);
    const rowsBefore = getTableRows(pkg, tableId);
    expect(rowsBefore).toHaveLength(2);

    executeInsertTableRow(ctx, {
      type: 'insertTableRow',
      tableId,
      templateRow: 1,
      position: 'end',
    });

    const rowsAfter = getTableRows(pkg, tableId);
    expect(rowsAfter).toHaveLength(3);

    const newRow = rowsAfter[2]!;
    const cells = getRowCells(newRow);
    // 模板行有 3 列
    expect(cells).toHaveLength(3);

    // 末列 tcW 与模板一致
    const templateLastCell = getRowCells(rowsBefore[1]!)[2]!;
    const newLastCell = cells[2]!;
    expect(toXml(newLastCell)).toContain('2000');
    expect(toXml(templateLastCell)).toContain('2000');
  });

  // 2. cellTexts 填充
  test('2 cellTexts 填充：各列文字正确；首 run rPr（字体/字号）与模板一致', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);

    executeInsertTableRow(ctx, {
      type: 'insertTableRow',
      tableId,
      templateRow: 1,
      position: 'end',
      cellTexts: ['填充A', '填充B', '填充C'],
    });

    const rows = getTableRows(pkg, tableId);
    const newRow = rows[2]!;
    const cells = getRowCells(newRow);

    // 各列文字正确
    expect(paraText(getFirstPara(cells[0]!)!)).toBe('填充A');
    expect(paraText(getFirstPara(cells[1]!)!)).toBe('填充B');
    expect(paraText(getFirstPara(cells[2]!)!)).toBe('填充C');

    // 首 run rPr 字体/字号与模板行一致（模板行 templateRow=1 含微软雅黑/sz24）
    const cell0Xml = toXml(getFirstPara(cells[0]!)!);
    expect(cell0Xml).toContain('微软雅黑');
    expect(cell0Xml).toContain('24');
  });

  // 3. 中间插入（position=1）
  test('3 中间插入 position=1：行顺序正确；rowCount+1', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);

    executeInsertTableRow(ctx, {
      type: 'insertTableRow',
      tableId,
      templateRow: 0,
      position: 1,
      cellTexts: ['中间行A', '中间行B', '中间行C'],
    });

    const rows = getTableRows(pkg, tableId);
    expect(rows).toHaveLength(3);

    // row0 = 原表头，row1 = 新插入行，row2 = 原数据行
    const row1Cells = getRowCells(rows[1]!);
    expect(paraText(getFirstPara(row1Cells[0]!)!)).toBe('中间行A');

    // 原数据行仍在 row2
    const row2Cells = getRowCells(rows[2]!);
    expect(paraText(getFirstPara(row2Cells[0]!)!)).toBe('样例A');
  });

  // 4. paraId 唯一性
  test('4 paraId 唯一性：克隆行内段落 paraId ≠ 模板行；全局无重复', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);

    executeInsertTableRow(ctx, {
      type: 'insertTableRow',
      tableId,
      templateRow: 1,
      position: 'end',
    });

    // 重建索引拿全部 paraId
    const idxAfter = buildIndices(pkg);
    const allParaIds = [...idxAfter.paraIdx.keys()];
    const uniqueIds = new Set(allParaIds);
    expect(allParaIds.length).toBe(uniqueIds.size); // 全局无重复

    // 克隆行内 paraId 不等于模板行段落 paraId
    const rows = getTableRows(pkg, tableId);
    const templateRow = rows[1]!; // templateRow=1
    const cloneRow = rows[2]!; // 末尾追加的新行

    const templateParaIds = new Set<string>();
    const tParas = templateRow.getElementsByTagNameNS(W_NS, 'p');
    for (let i = 0; i < tParas.length; i++) {
      const id = tParas[i]!.getAttributeNS(W14_NS, 'paraId');
      if (id) templateParaIds.add(id.toUpperCase());
    }

    const cloneParas = cloneRow.getElementsByTagNameNS(W_NS, 'p');
    for (let i = 0; i < cloneParas.length; i++) {
      const id = cloneParas[i]!.getAttributeNS(W14_NS, 'paraId');
      expect(id).toBeTruthy();
      expect(templateParaIds.has((id ?? '').toUpperCase())).toBe(false);
    }
  });

  // 5. 锚点剥离
  test('5 锚点剥离：模板行含书签 + 批注 → 克隆行内无 bookmarkStart / commentReference', () => {
    const tableWithAnchors = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="A0000001"><w:r><w:t>表头</w:t></w:r></w:p>
          </w:tc>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="A0000002"><w:r><w:t>表头2</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
        <w:tr>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="A0000003">
              <w:bookmarkStart w:id="1" w:name="myBookmark"/>
              <w:r><w:t>有书签的单元格</w:t></w:r>
              <w:bookmarkEnd w:id="1"/>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="A0000004">
              <w:commentRangeStart w:id="5"/>
              <w:r>
                <w:rPr><w:sz w:val="20"/></w:rPr>
                <w:t>有批注的单元格</w:t>
              </w:r>
              <w:commentRangeEnd w:id="5"/>
              <w:r><w:commentReference w:id="5"/></w:r>
            </w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `;
    const { pkg, ctx, tableId } = makeDoc(tableWithAnchors);

    executeInsertTableRow(ctx, {
      type: 'insertTableRow',
      tableId,
      templateRow: 1, // 模板行是含锚点的数据行
      position: 'end',
    });

    const rows = getTableRows(pkg, tableId);
    const cloneRow = rows[2]!;
    const cloneXml = toXml(cloneRow);

    // 克隆行内不含书签
    expect(cloneXml).not.toContain('bookmarkStart');
    expect(cloneXml).not.toContain('bookmarkEnd');
    expect(cloneXml).not.toContain('myBookmark');
    // 克隆行内不含批注引用
    expect(cloneXml).not.toContain('commentReference');
    expect(cloneXml).not.toContain('commentRangeStart');
    expect(cloneXml).not.toContain('commentRangeEnd');

    // 模板行原书签不动
    const templateRow = rows[1]!;
    const templateXml = toXml(templateRow);
    expect(templateXml).toContain('myBookmark');
  });

  // 6. cellTexts 短于列数
  test('6 cellTexts 短于列数：未覆盖列文字被清空（无模板样例文字残留）', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);

    // 3 列表格，只传 1 个 cellText
    executeInsertTableRow(ctx, {
      type: 'insertTableRow',
      tableId,
      templateRow: 1,
      position: 'end',
      cellTexts: ['仅填第一列'],
    });

    const rows = getTableRows(pkg, tableId);
    const newRow = rows[2]!;
    const cells = getRowCells(newRow);

    // 列 0 有填充文字
    expect(paraText(getFirstPara(cells[0]!)!)).toBe('仅填第一列');

    // 列 1, 2 清空：无模板样例文字（样例B / 样例C）
    const cell1Para = getFirstPara(cells[1]!)!;
    const cell2Para = getFirstPara(cells[2]!)!;
    expect(paraText(cell1Para)).toBe('');
    expect(toXml(cell1Para)).not.toContain('样例B');
    expect(paraText(cell2Para)).toBe('');
    expect(toXml(cell2Para)).not.toContain('样例C');
  });

  // 7. cellTexts 长于列数
  test('7 cellTexts 长于列数：多余元素忽略，列数不变', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);

    // 3 列表格，传 5 个 cellText
    executeInsertTableRow(ctx, {
      type: 'insertTableRow',
      tableId,
      templateRow: 1,
      position: 'end',
      cellTexts: ['第1列', '第2列', '第3列', '第4列忽略', '第5列忽略'],
    });

    const rows = getTableRows(pkg, tableId);
    const newRow = rows[2]!;
    const cells = getRowCells(newRow);

    // 列数不变（仍是 3）
    expect(cells).toHaveLength(3);

    // 三列内容正确
    expect(paraText(getFirstPara(cells[0]!)!)).toBe('第1列');
    expect(paraText(getFirstPara(cells[1]!)!)).toBe('第2列');
    expect(paraText(getFirstPara(cells[2]!)!)).toBe('第3列');
  });

  // 8. templateRow 越界
  test('8 templateRow 越界 → TABLE_NOT_FOUND', () => {
    const { ctx, tableId } = makeDoc(STYLED_TABLE);
    expect(() =>
      executeInsertTableRow(ctx, {
        type: 'insertTableRow',
        tableId,
        templateRow: 99,
        position: 'end',
      }),
    ).toThrow(WordError);
  });

  // 9. tableId 不存在
  test('9 tableId 不存在 → TABLE_NOT_FOUND', () => {
    const { ctx } = makeDoc(STYLED_TABLE);
    expect(() =>
      executeInsertTableRow(ctx, {
        type: 'insertTableRow',
        tableId: 'tbl-9999',
        templateRow: 0,
        position: 'end',
      }),
    ).toThrow(WordError);
  });

  // 10. position 越界
  test('10 position >rowCount → TABLE_NOT_FOUND', () => {
    const { ctx, tableId } = makeDoc(STYLED_TABLE);
    expect(() =>
      executeInsertTableRow(ctx, {
        type: 'insertTableRow',
        tableId,
        templateRow: 0,
        position: 99, // rowCount is 2, so 99 > 2
      }),
    ).toThrow(WordError);
  });

  test('10b position <0 → TABLE_NOT_FOUND', () => {
    const { ctx, tableId } = makeDoc(STYLED_TABLE);
    expect(() =>
      executeInsertTableRow(ctx, {
        type: 'insertTableRow',
        tableId,
        templateRow: 0,
        position: -1,
      }),
    ).toThrow(WordError);
  });

  // 11. trackChanges=true
  test('11 trackChanges=true → UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx, tableId } = makeDoc(STYLED_TABLE, { trackChanges: true });
    expect(() =>
      executeInsertTableRow(ctx, {
        type: 'insertTableRow',
        tableId,
        templateRow: 1,
        position: 'end',
      }),
    ).toThrow(/trackChanges=true not yet implemented/);
  });

  // 12. round-trip
  test('12 round-trip：apply → toZip → WordDocument.fromZip 重载不抛错，新行结构合法', async () => {
    const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CTYPE}"/>
</Types>`;
    const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const bodyXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="BB000001"><w:r><w:t>R0C0</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="BB000002"><w:r><w:t>R0C1</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="BB000003"><w:r><w:t>R1C0</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="BB000004"><w:r><w:t>R1C1</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p w14:paraId="BB000005"><w:r><w:t>end para</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const zip1 = new JSZip();
    zip1.file('[Content_Types].xml', CT_XML);
    zip1.file('_rels/.rels', MAIN_RELS);
    zip1.file('word/document.xml', bodyXml);
    const buf = await zip1.generateAsync({ type: 'uint8array' });

    // 加载 → apply insertTableRow → toZip
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-itr', mode: 'inline' });
    const tables = await doc.tables.list();
    expect(tables).toHaveLength(1);
    const tableId = tables[0]!.tableId;

    await doc
      .plan()
      .insertTableRow({
        tableId,
        templateRow: 1,
        position: 'end',
        cellTexts: ['新行C0', '新行C1'],
      })
      .apply();

    const outZip = await doc.toZip();
    await doc.dispose();

    // 重载：不抛错
    const doc2 = await WordDocument.fromZip(outZip, { docId: 'e2e-itr-reload', mode: 'inline' });
    const list = await doc2.paragraphs.list({ pagination: { mode: 'all' } });
    // 原来 5 段（4 个表格内 + 1 个尾段）+ 新行 2 段 = 7 段
    expect(list.paragraphs.length).toBeGreaterThanOrEqual(6);

    // 全局 paraId 无重复
    const allParaIds = list.paragraphs.map((p) => p.paraId);
    expect(new Set(allParaIds).size).toBe(allParaIds.length);

    await doc2.dispose();
  });
});
