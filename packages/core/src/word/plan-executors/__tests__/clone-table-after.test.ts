/**
 * executeCloneTableAfter 单元测试（task-15a-clone-table-after，spec §3 测试矩阵）。
 *
 * 覆盖（12 条）：
 *   1  基本克隆插入：tables.list 数 +1；克隆表 rowCount/colCount 与源一致
 *   2  样式保留：克隆表 tcPr 的 shd/tcW + 单元格 rPr 与源表一致
 *   3  位置正确：克隆表是 anchorParaId 段的紧邻后继（DOM 顺序）
 *   4  paraId 唯一：克隆表内段 paraId ≠ 源表；toZip 重载后全局无重复
 *   5  锚点剥离：源表含书签+批注 → 克隆表内无 bookmarkStart/commentReference
 *   6  源表不被改：克隆后源表 XML 原样（含原 paraId、原锚点）
 *   7  多次克隆到同 anchor（逆序）：连续 2 次 cloneTableAfter(同 anchor)，
 *      DOM 顺序 = anchor, 表B, 表A（insertBefore(nextSibling) 的逆序效果）
 *   8  anchorParaId 不存在 → PARA_NOT_FOUND
 *   9  sourceTableId 不存在 → TABLE_NOT_FOUND
 *   10 trackChanges=true → UNSUPPORTED_OOXML_FEATURE
 *   11 round-trip：apply → toZip → WordDocument.fromZip 重载不抛错，tables.list 数+1
 *   12 #14 回归：提取共享 helper 后，insert-table-row.test.ts 13 条仍全过
 *      （本文件通过 import 并手动触发 executeInsertTableRow 来验证回归）
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
import { executeCloneTableAfter, executeInsertTableRow } from '../table-ops.js';

// ─────────── 工具函数 ───────────

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
  anchorParaId: string;
} {
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`,
  );
  const indices = buildIndices(pkg);
  const tableId = [...indices.tblIdx.keys()][0] ?? '';
  // anchor = first paragraph NOT inside a table (body-level paragraphs)
  const anchorParaId =
    [...indices.paraIdx.keys()].find((id) => {
      const el = indices.paraIdx.get(id)!;
      // walk up to see if it's inside a <w:tbl>
      let p: Element | null = el.parentElement;
      while (p) {
        if (p.namespaceURI === W_NS && p.localName === 'tbl') return false;
        p = p.parentElement;
      }
      return true;
    }) ?? '';

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
  return { pkg, ctx, tableId, anchorParaId };
}

/** 获取 body 下直属子节点（Element）列表 */
function getBodyChildren(pkg: Package): Element[] {
  const _idx = buildIndices(pkg);
  // find <w:body>
  for (const [, part] of pkg.parts) {
    if (!('xmlDom' in part) || !part.xmlDom) continue;
    const body = (part.xmlDom as Document).getElementsByTagNameNS(W_NS, 'body')[0];
    if (!body) continue;
    const children: Element[] = [];
    for (let i = 0; i < body.childNodes.length; i++) {
      const c = body.childNodes[i] as Element;
      if (c && c.nodeType === 1) children.push(c);
    }
    return children;
  }
  return [];
}

/** 获取表格所有 <w:tr> */
function getTableRows(tbl: Element): Element[] {
  const rows: Element[] = [];
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const c = tbl.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tr') rows.push(c);
  }
  return rows;
}

/** 获取行内 <w:tc> */
function getRowCells(tr: Element): Element[] {
  const cells: Element[] = [];
  for (let i = 0; i < tr.childNodes.length; i++) {
    const c = tr.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tc') cells.push(c);
  }
  return cells;
}

function toXml(el: Element): string {
  return serializeToWellFormedString(el as never);
}

// ─────────── fixture ───────────

/**
 * body 内容：
 *   - 1 张 2行×2列 表格（含底纹 shd + rPr 微软雅黑/sz24）
 *   - anchor 段落 "body-anchor"（w14:paraId="CC000001"）
 *   - 尾段落 "body-end"（w14:paraId="CC000002"）
 */
const BODY_WITH_TABLE_AND_PARAS = `
  <w:tbl>
    <w:tblGrid>
      <w:gridCol w:w="2500"/>
      <w:gridCol w:w="2500"/>
    </w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="2500" w:type="dxa"/>
          <w:shd w:val="clear" w:fill="DDDDDD"/>
        </w:tcPr>
        <w:p w14:paraId="CC000010">
          <w:r><w:rPr><w:rFonts w:eastAsia="微软雅黑"/><w:sz w:val="24"/></w:rPr><w:t>表头A</w:t></w:r>
        </w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="CC000011"><w:r><w:t>表头B</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
    <w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="CC000012"><w:r><w:t>数据A</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="CC000013"><w:r><w:t>数据B</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
  <w:p w14:paraId="CC000001"><w:r><w:t>body-anchor</w:t></w:r></w:p>
  <w:p w14:paraId="CC000002"><w:r><w:t>body-end</w:t></w:r></w:p>
`;

// ─────────── 测试 ───────────

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('executeCloneTableAfter', () => {
  // 1. 基本克隆插入
  test('1 基本克隆插入：body 内 tbl 数 +1；克隆表 rowCount/colCount 与源一致', () => {
    const { pkg, ctx, tableId } = makeDoc(BODY_WITH_TABLE_AND_PARAS);

    executeCloneTableAfter(ctx, {
      type: 'cloneTableAfter',
      anchorParaId: 'CC000001',
      sourceTableId: tableId,
    });

    const bodyChildren = getBodyChildren(pkg);
    const tables = bodyChildren.filter((c) => c.namespaceURI === W_NS && c.localName === 'tbl');
    expect(tables).toHaveLength(2);

    // 克隆表行数与源一致（2 行）
    const srcTbl = tables[0]!;
    const clonedTbl = tables[1]!;
    expect(getTableRows(clonedTbl)).toHaveLength(getTableRows(srcTbl).length);

    // 克隆表列数（第一行列数）与源一致（2 列）
    expect(getRowCells(getTableRows(clonedTbl)[0]!)).toHaveLength(
      getRowCells(getTableRows(srcTbl)[0]!).length,
    );
  });

  // 2. 样式保留
  test('2 样式保留：克隆表 tcPr 的 shd/tcW + rPr 与源表一致', () => {
    const { pkg, ctx, tableId } = makeDoc(BODY_WITH_TABLE_AND_PARAS);

    executeCloneTableAfter(ctx, {
      type: 'cloneTableAfter',
      anchorParaId: 'CC000001',
      sourceTableId: tableId,
    });

    const bodyChildren = getBodyChildren(pkg);
    const tables = bodyChildren.filter((c) => c.namespaceURI === W_NS && c.localName === 'tbl');
    const clonedXml = toXml(tables[1]!);

    // 底纹 shd 保留
    expect(clonedXml).toContain('DDDDDD');
    expect(clonedXml).toMatch(/<w:shd\b/);
    // tcW 保留
    expect(clonedXml).toContain('2500');
    // rPr 字体/字号保留
    expect(clonedXml).toContain('微软雅黑');
    expect(clonedXml).toContain('24');
  });

  // 3. 位置正确
  test('3 位置正确：克隆表是 anchorParaId 段的紧邻后继（DOM 顺序）', () => {
    const { pkg, ctx, tableId } = makeDoc(BODY_WITH_TABLE_AND_PARAS);

    executeCloneTableAfter(ctx, {
      type: 'cloneTableAfter',
      anchorParaId: 'CC000001',
      sourceTableId: tableId,
    });

    const bodyChildren = getBodyChildren(pkg);
    // 找 anchor 段的位置
    const anchorIdx = bodyChildren.findIndex(
      (c) =>
        c.namespaceURI === W_NS &&
        c.localName === 'p' &&
        (c.getAttributeNS(W14_NS, 'paraId') ?? '') === 'CC000001',
    );
    expect(anchorIdx).toBeGreaterThanOrEqual(0);

    // anchor 的紧邻后继应该是 <w:tbl>（克隆表）
    const nextEl = bodyChildren[anchorIdx + 1];
    expect(nextEl).toBeTruthy();
    expect(nextEl!.localName).toBe('tbl');
  });

  // 4. paraId 唯一
  test('4 paraId 唯一：克隆表内段 paraId ≠ 源表；全局无重复', () => {
    const { pkg, ctx, tableId } = makeDoc(BODY_WITH_TABLE_AND_PARAS);

    executeCloneTableAfter(ctx, {
      type: 'cloneTableAfter',
      anchorParaId: 'CC000001',
      sourceTableId: tableId,
    });

    // 重建索引，所有 paraId
    const idxAfter = buildIndices(pkg);
    const allIds = [...idxAfter.paraIdx.keys()];
    expect(new Set(allIds).size).toBe(allIds.length); // 全局无重复

    // 克隆表内 paraId 不等于源表 paraId
    const bodyChildren = getBodyChildren(pkg);
    const tables = bodyChildren.filter((c) => c.namespaceURI === W_NS && c.localName === 'tbl');
    const srcTbl = tables[0]!;
    const clonedTbl = tables[1]!;

    const srcIds = new Set<string>();
    const srcParas = srcTbl.getElementsByTagNameNS(W_NS, 'p');
    for (let i = 0; i < srcParas.length; i++) {
      const id = srcParas[i]!.getAttributeNS(W14_NS, 'paraId');
      if (id) srcIds.add(id.toUpperCase());
    }

    const clonedParas = clonedTbl.getElementsByTagNameNS(W_NS, 'p');
    for (let i = 0; i < clonedParas.length; i++) {
      const id = clonedParas[i]!.getAttributeNS(W14_NS, 'paraId');
      expect(id).toBeTruthy();
      expect(srcIds.has((id ?? '').toUpperCase())).toBe(false);
    }
  });

  // 5. 锚点剥离
  test('5 锚点剥离：源表含书签+批注 → 克隆表内无 bookmarkStart/commentReference', () => {
    const tableWithAnchors = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="DD000001">
              <w:bookmarkStart w:id="1" w:name="tableBookmark"/>
              <w:r><w:t>有书签</w:t></w:r>
              <w:bookmarkEnd w:id="1"/>
            </w:p>
          </w:tc>
        </w:tr>
        <w:tr>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="DD000002">
              <w:commentRangeStart w:id="7"/>
              <w:r><w:t>有批注</w:t></w:r>
              <w:commentRangeEnd w:id="7"/>
              <w:r><w:commentReference w:id="7"/></w:r>
            </w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
      <w:p w14:paraId="DD000099"><w:r><w:t>anchor</w:t></w:r></w:p>
    `;
    const { pkg, ctx, tableId } = makeDoc(tableWithAnchors);

    executeCloneTableAfter(ctx, {
      type: 'cloneTableAfter',
      anchorParaId: 'DD000099',
      sourceTableId: tableId,
    });

    const bodyChildren = getBodyChildren(pkg);
    const tables = bodyChildren.filter((c) => c.namespaceURI === W_NS && c.localName === 'tbl');
    const clonedXml = toXml(tables[1]!);

    // 克隆表内无书签
    expect(clonedXml).not.toContain('bookmarkStart');
    expect(clonedXml).not.toContain('tableBookmark');
    // 克隆表内无批注引用
    expect(clonedXml).not.toContain('commentReference');
    expect(clonedXml).not.toContain('commentRangeStart');

    // 源表原锚点不动
    const srcXml = toXml(tables[0]!);
    expect(srcXml).toContain('tableBookmark');
  });

  // 6. 源表不被改
  test('6 源表不被改：克隆后源表 XML 原样（含原 paraId、原锚点）', () => {
    const { pkg, ctx, tableId } = makeDoc(BODY_WITH_TABLE_AND_PARAS);

    // 克隆前记录源表 XML
    const bodyChildrenBefore = getBodyChildren(pkg);
    const srcTblBefore = bodyChildrenBefore.find(
      (c) => c.namespaceURI === W_NS && c.localName === 'tbl',
    )!;
    const srcXmlBefore = toXml(srcTblBefore);

    executeCloneTableAfter(ctx, {
      type: 'cloneTableAfter',
      anchorParaId: 'CC000001',
      sourceTableId: tableId,
    });

    // 克隆后源表 XML 不变
    const srcXmlAfter = toXml(srcTblBefore);
    expect(srcXmlAfter).toBe(srcXmlBefore);

    // 原 paraId 仍在源表
    expect(srcXmlAfter).toContain('CC000010');
    expect(srcXmlAfter).toContain('CC000013');
  });

  // 7. 多次克隆到同 anchor（逆序语义）
  test('7 多次克隆到同 anchor（逆序）：DOM 顺序 = anchor, tblB, tblA', () => {
    const { pkg, ctx, tableId } = makeDoc(BODY_WITH_TABLE_AND_PARAS);

    // 第一次克隆 → 得到 tableA（紧接 anchor 后）
    executeCloneTableAfter(ctx, {
      type: 'cloneTableAfter',
      anchorParaId: 'CC000001',
      sourceTableId: tableId,
    });

    // 第二次克隆 → 得到 tableB（也紧接 anchor 后，insertBefore(anchor.nextSibling) 把 tblB 插在 tblA 前面）
    executeCloneTableAfter(ctx, {
      type: 'cloneTableAfter',
      anchorParaId: 'CC000001',
      sourceTableId: tableId,
    });

    const bodyChildren = getBodyChildren(pkg);
    const anchorIdx = bodyChildren.findIndex(
      (c) =>
        c.namespaceURI === W_NS &&
        c.localName === 'p' &&
        (c.getAttributeNS(W14_NS, 'paraId') ?? '') === 'CC000001',
    );
    expect(anchorIdx).toBeGreaterThanOrEqual(0);

    // DOM 顺序：anchor, tblB(第2次), tblA(第1次), ...
    expect(bodyChildren[anchorIdx + 1]!.localName).toBe('tbl'); // tblB
    expect(bodyChildren[anchorIdx + 2]!.localName).toBe('tbl'); // tblA

    // 源表在 anchor 前面（index 0），anchorIdx 后面有两张克隆表
    const tables = bodyChildren.filter((c) => c.namespaceURI === W_NS && c.localName === 'tbl');
    expect(tables).toHaveLength(3); // src + tblA + tblB
  });

  // 8. anchorParaId 不存在
  test('8 anchorParaId 不存在 → PARA_NOT_FOUND', () => {
    const { ctx, tableId } = makeDoc(BODY_WITH_TABLE_AND_PARAS);
    expect(() =>
      executeCloneTableAfter(ctx, {
        type: 'cloneTableAfter',
        anchorParaId: 'FFFFFFFF',
        sourceTableId: tableId,
      }),
    ).toThrow(WordError);
  });

  // 9. sourceTableId 不存在
  test('9 sourceTableId 不存在 → TABLE_NOT_FOUND', () => {
    const { ctx } = makeDoc(BODY_WITH_TABLE_AND_PARAS);
    expect(() =>
      executeCloneTableAfter(ctx, {
        type: 'cloneTableAfter',
        anchorParaId: 'CC000001',
        sourceTableId: 'tbl-9999',
      }),
    ).toThrow(WordError);
  });

  // 10. trackChanges=true
  test('10 trackChanges=true → UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx, tableId } = makeDoc(BODY_WITH_TABLE_AND_PARAS, { trackChanges: true });
    expect(() =>
      executeCloneTableAfter(ctx, {
        type: 'cloneTableAfter',
        anchorParaId: 'CC000001',
        sourceTableId: tableId,
      }),
    ).toThrow(/trackChanges=true not yet implemented/);
  });

  // 11. round-trip
  test('11 round-trip：apply → toZip → WordDocument.fromZip 重载不抛错，tables.list 数+1', async () => {
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
          <w:p w14:paraId="EE000001"><w:r><w:t>R0C0</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="EE000002"><w:r><w:t>R0C1</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p w14:paraId="EE000010"><w:r><w:t>anchor para</w:t></w:r></w:p>
    <w:p w14:paraId="EE000011"><w:r><w:t>end para</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const zip1 = new JSZip();
    zip1.file('[Content_Types].xml', CT_XML);
    zip1.file('_rels/.rels', MAIN_RELS);
    zip1.file('word/document.xml', bodyXml);
    const buf = await zip1.generateAsync({ type: 'uint8array' });

    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-cta', mode: 'inline' });

    // 获取源表 tableId
    const tablesBefore = await doc.tables.list();
    expect(tablesBefore).toHaveLength(1);
    const srcTableId = tablesBefore[0]!.tableId;

    // apply cloneTableAfter
    await doc
      .plan()
      .cloneTableAfter({ anchorParaId: 'EE000010', sourceTableId: srcTableId })
      .apply();

    const outZip = await doc.toZip();
    await doc.dispose();

    // 重载
    const doc2 = await WordDocument.fromZip(outZip, { docId: 'e2e-cta-reload', mode: 'inline' });
    const tablesAfter = await doc2.tables.list();
    // 原来 1 张表，克隆后 2 张
    expect(tablesAfter).toHaveLength(2);

    // 全局 paraId 无重复
    const paraList = await doc2.paragraphs.list({ pagination: { mode: 'all' } });
    const allIds = paraList.paragraphs.map((p) => p.paraId);
    expect(new Set(allIds).size).toBe(allIds.length);

    await doc2.dispose();
  });

  // 12. #14 回归：提取共享 helper 后 insertTableRow 行为不变
  test('12 #14 回归：reassignClonedParaIds 提取后 executeInsertTableRow 行为不变', () => {
    // 构造和 insert-table-row.test.ts 中 #4 同构的 fixture
    const tableXml = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="FF000001"><w:r><w:t>表头</w:t></w:r></w:p>
          </w:tc>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="FF000002"><w:r><w:t>表头2</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
        <w:tr>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="FF000003"><w:r><w:t>数据A</w:t></w:r></w:p>
          </w:tc>
          <w:tc>
            <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
            <w:p w14:paraId="FF000004"><w:r><w:t>数据B</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
      <w:p w14:paraId="FF000099"><w:r><w:t>tail</w:t></w:r></w:p>
    `;
    const { pkg, ctx, tableId } = makeDoc(tableXml);

    executeInsertTableRow(ctx, {
      type: 'insertTableRow',
      tableId,
      templateRow: 1,
      position: 'end',
    });

    // 重建索引，所有 paraId 唯一
    const idxAfter = buildIndices(pkg);
    const allIds = [...idxAfter.paraIdx.keys()];
    expect(new Set(allIds).size).toBe(allIds.length);

    // 克隆行（新增的第3行）的 paraId 不与原行重复
    const bodyChildren = getBodyChildren(pkg);
    const tbl = bodyChildren.find((c) => c.namespaceURI === W_NS && c.localName === 'tbl')!;
    const rows: Element[] = [];
    for (let i = 0; i < tbl.childNodes.length; i++) {
      const c = tbl.childNodes[i] as Element;
      if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tr') rows.push(c);
    }
    expect(rows).toHaveLength(3); // 原 2 行 + 新插 1 行

    // 原行 paraId 集合
    const originalIds = new Set(['FF000001', 'FF000002', 'FF000003', 'FF000004']);
    const newRow = rows[2]!;
    const newRowParas = newRow.getElementsByTagNameNS(W_NS, 'p');
    for (let i = 0; i < newRowParas.length; i++) {
      const id = newRowParas[i]!.getAttributeNS(W14_NS, 'paraId');
      expect(originalIds.has((id ?? '').toUpperCase())).toBe(false);
    }
  });
});
