/**
 * executeReplaceTableCellWithSdt 单元测试(#188)。
 *
 * 覆盖:
 *   - 基本: 替换单元格首段 runs 为 inline SDT, 保留 pPr / tcPr / 表格结构
 *   - rPr 透传: 原 run 的 rPr (颜色/字体) 拷贝到 SDT 内 placeholder run
 *   - 多段单元格: 只替换首段, 其它段保留
 *   - 锚点冲突: error 抛 ANCHOR_CONFLICT / drop 跳过
 *   - 不存在的 tableId / row / col 抛对应错
 *   - trackChanges=true 抛 UNSUPPORTED_OOXML_FEATURE
 *   - hashSdtId 幂等: 相同 (tableId, row, col, alias) 重复跑产同一 id
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeReplaceTableCellWithSdt, executeDeleteTableRow } from '../table-ops.js';

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
      date: new Date('2026-05-19T00:00:00Z'),
    },
    result: emptyApplyResult(),
    opIndex: 0,
    docId: 'test-doc',
  };
  return { pkg, ctx, tableId };
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

/** 单元格首段序列化, 便于结构断言 */
function cellFirstParaXml(pkg: Package, tableId: string, row: number, col: number): string {
  const idx = buildIndices(pkg);
  const tbl = idx.tblIdx.get(tableId)!;
  const trs = tbl.getElementsByTagNameNS(W_NS, 'tr');
  const tr = trs[row]!;
  // 取该 tr 的直属 tc (避免拿到嵌套表的 tc)
  const tcs: Element[] = [];
  for (let i = 0; i < tr.childNodes.length; i++) {
    const c = tr.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tc') tcs.push(c);
  }
  const tc = tcs[col]!;
  // 该 tc 的首个直属 <w:p>
  for (let i = 0; i < tc.childNodes.length; i++) {
    const c = tc.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'p') {
      return serializeToWellFormedString(c as never);
    }
  }
  throw new Error('no <w:p> in cell');
}

const SIMPLE_TABLE = `
  <w:tbl>
    <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000001"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>编制单位</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000002"><w:r><w:t>上海某某科技有限公司</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
    <w:tr>
      <w:tc>
        <w:p w14:paraId="00000003"><w:r><w:t>签字日期</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:p w14:paraId="00000004"><w:r><w:t>2025-12-31</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
`;

describe('executeReplaceTableCellWithSdt', () => {
  test('基本: 替换单元格首段 runs 为 SDT, 保留 pPr', () => {
    const { pkg, ctx, tableId } = makeDoc(SIMPLE_TABLE);
    executeReplaceTableCellWithSdt(ctx, {
      type: 'replaceTableCellWithSdt',
      tableId,
      row: 0,
      col: 1, // "上海某某科技有限公司" 单元格
      sdt: { kind: 'text', alias: '编制单位' },
      onAnchorConflict: 'error',
    });

    const xml = cellFirstParaXml(pkg, tableId, 0, 1);
    expect(xml).toContain('<w:sdt');
    expect(xml).toContain('<w:alias w:val="编制单位"');
    expect(xml).toContain('<w:text');
    expect(xml).toContain('请输入编制单位');
    // 原文本不应残留
    expect(xml).not.toContain('上海某某科技有限公司');
  });

  test('rPr 透传: 原 run 颜色拷贝到 SDT 内 placeholder run', () => {
    const { pkg, ctx, tableId } = makeDoc(SIMPLE_TABLE);
    // 替换 (0,0) "编制单位" 单元格 —— 该 run 有 <w:color w:val="FF0000"/>
    executeReplaceTableCellWithSdt(ctx, {
      type: 'replaceTableCellWithSdt',
      tableId,
      row: 0,
      col: 0,
      sdt: { kind: 'text', alias: '字段名' },
      onAnchorConflict: 'error',
    });

    const xml = cellFirstParaXml(pkg, tableId, 0, 0);
    // sdtPr 内一个 rPr, sdtContent 内一个 rPr
    expect((xml.match(/<w:rPr/g) ?? []).length).toBe(2);
    expect((xml.match(/<w:color w:val="FF0000"/g) ?? []).length).toBe(2);
    // pPr (段落属性 jc=center) 保留
    expect(xml).toContain('<w:jc w:val="center"');
  });

  test('Date SDT: dateFormat / lid 写入', () => {
    const { pkg, ctx, tableId } = makeDoc(SIMPLE_TABLE);
    executeReplaceTableCellWithSdt(ctx, {
      type: 'replaceTableCellWithSdt',
      tableId,
      row: 1,
      col: 1, // "2025-12-31"
      sdt: { kind: 'date', alias: '签字日期', dateFormat: 'yyyy-MM-dd', locale: 'zh-CN' },
      onAnchorConflict: 'error',
    });
    const xml = cellFirstParaXml(pkg, tableId, 1, 1);
    expect(xml).toContain('<w:date');
    expect(xml).toContain('<w:dateFormat w:val="yyyy-MM-dd"');
    expect(xml).toContain('<w:lid w:val="zh-CN"');
  });

  test('锚点冲突 error: 单元格内有书签 → 抛 ANCHOR_CONFLICT', () => {
    const tableWithBookmark = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:p w14:paraId="00000010">
              <w:bookmarkStart w:id="0" w:name="someBookmark"/>
              <w:r><w:t>data</w:t></w:r>
              <w:bookmarkEnd w:id="0"/>
            </w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `;
    const { ctx, tableId } = makeDoc(tableWithBookmark);
    expect(() =>
      executeReplaceTableCellWithSdt(ctx, {
        type: 'replaceTableCellWithSdt',
        tableId,
        row: 0,
        col: 0,
        sdt: { kind: 'text', alias: 'x' },
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('replacedTableCells 计数 (#204): drop 模式跳过的不计入', () => {
    const tableWithMixed = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr><w:tc><w:p w14:paraId="00000400"><w:r><w:t>plain</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p w14:paraId="00000401">
          <w:bookmarkStart w:id="0" w:name="anchorMark"/>
          <w:r><w:t>locked</w:t></w:r>
          <w:bookmarkEnd w:id="0"/>
        </w:p></w:tc></w:tr>
      </w:tbl>
    `;
    const { ctx, tableId } = makeDoc(tableWithMixed);
    expect(ctx.result.replacedTableCells).toBe(0);

    // (0,0) 无锚点 → 真正替换, 计数 +1
    executeReplaceTableCellWithSdt(ctx, {
      type: 'replaceTableCellWithSdt',
      tableId,
      row: 0,
      col: 0,
      sdt: { kind: 'text', alias: 'a' },
      onAnchorConflict: 'drop',
    });
    expect(ctx.result.replacedTableCells).toBe(1);

    // (1,0) 有书签 + drop → 静默跳过, 计数不变
    executeReplaceTableCellWithSdt(ctx, {
      type: 'replaceTableCellWithSdt',
      tableId,
      row: 1,
      col: 0,
      sdt: { kind: 'text', alias: 'b' },
      onAnchorConflict: 'drop',
    });
    expect(ctx.result.replacedTableCells).toBe(1);
  });

  test('锚点冲突 drop: 单元格内有书签 → 跳过, 原内容保留', () => {
    const tableWithBookmark = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:p w14:paraId="00000011">
              <w:bookmarkStart w:id="0" w:name="some"/>
              <w:r><w:t>keepme</w:t></w:r>
              <w:bookmarkEnd w:id="0"/>
            </w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `;
    const { pkg, ctx, tableId } = makeDoc(tableWithBookmark);
    executeReplaceTableCellWithSdt(ctx, {
      type: 'replaceTableCellWithSdt',
      tableId,
      row: 0,
      col: 0,
      sdt: { kind: 'text', alias: 'x' },
      onAnchorConflict: 'drop',
    });
    const xml = cellFirstParaXml(pkg, tableId, 0, 0);
    expect(xml).toContain('keepme');
    expect(xml).not.toContain('<w:sdt');
  });

  test('不存在的 tableId 抛 PARA_NOT_FOUND', () => {
    const { ctx } = makeDoc(SIMPLE_TABLE);
    expect(() =>
      executeReplaceTableCellWithSdt(ctx, {
        type: 'replaceTableCellWithSdt',
        tableId: 'tbl-9999',
        row: 0,
        col: 0,
        sdt: { kind: 'text', alias: 'x' },
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('row / col 越界抛 PARA_NOT_FOUND', () => {
    const { ctx, tableId } = makeDoc(SIMPLE_TABLE);
    expect(() =>
      executeReplaceTableCellWithSdt(ctx, {
        type: 'replaceTableCellWithSdt',
        tableId,
        row: 99,
        col: 0,
        sdt: { kind: 'text', alias: 'x' },
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
    expect(() =>
      executeReplaceTableCellWithSdt(ctx, {
        type: 'replaceTableCellWithSdt',
        tableId,
        row: 0,
        col: 99,
        sdt: { kind: 'text', alias: 'x' },
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('trackChanges=true 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx, tableId } = makeDoc(SIMPLE_TABLE, { trackChanges: true });
    expect(() =>
      executeReplaceTableCellWithSdt(ctx, {
        type: 'replaceTableCellWithSdt',
        tableId,
        row: 0,
        col: 0,
        sdt: { kind: 'text', alias: 'x' },
        onAnchorConflict: 'error',
      }),
    ).toThrowError(/trackChanges=true not yet implemented/);
  });

  test('删除最后一行后, 仍可域化剩余行', () => {
    const { pkg, ctx, tableId } = makeDoc(SIMPLE_TABLE);
    // 先删第 1 行 (1 row 数据)
    executeDeleteTableRow(ctx, {
      type: 'deleteTableRow',
      tableId,
      row: 1,
      onAnchorConflict: 'error',
    });
    // 再域化原 row 0 (现在是唯一行)
    executeReplaceTableCellWithSdt(ctx, {
      type: 'replaceTableCellWithSdt',
      tableId,
      row: 0,
      col: 1,
      sdt: { kind: 'text', alias: '编制单位' },
      onAnchorConflict: 'error',
    });
    const xml = cellFirstParaXml(pkg, tableId, 0, 1);
    expect(xml).toContain('<w:alias w:val="编制单位"');
    // 确认表格只剩一行
    const idx = buildIndices(pkg);
    const tbl = idx.tblIdx.get(tableId)!;
    const trs = tbl.getElementsByTagNameNS(W_NS, 'tr');
    expect(trs.length).toBe(1);
  });

  test('幂等: 相同 (tableId, row, col, alias) 重复跑, SDT id 不变', () => {
    const { pkg: pkg1, ctx: ctx1, tableId: t1 } = makeDoc(SIMPLE_TABLE);
    const { pkg: pkg2, ctx: ctx2, tableId: t2 } = makeDoc(SIMPLE_TABLE);
    expect(t1).toBe(t2);

    executeReplaceTableCellWithSdt(ctx1, {
      type: 'replaceTableCellWithSdt',
      tableId: t1,
      row: 0,
      col: 1,
      sdt: { kind: 'text', alias: '编制单位' },
      onAnchorConflict: 'error',
    });
    executeReplaceTableCellWithSdt(ctx2, {
      type: 'replaceTableCellWithSdt',
      tableId: t2,
      row: 0,
      col: 1,
      sdt: { kind: 'text', alias: '编制单位' },
      onAnchorConflict: 'error',
    });

    const xml1 = cellFirstParaXml(pkg1, t1, 0, 1);
    const xml2 = cellFirstParaXml(pkg2, t2, 0, 1);
    const idMatch1 = xml1.match(/<w:id w:val="(\d+)"/);
    const idMatch2 = xml2.match(/<w:id w:val="(\d+)"/);
    expect(idMatch1).not.toBeNull();
    expect(idMatch1![1]).toBe(idMatch2![1]);
  });
});

// ─────────── executeDeleteTableRow ───────────

describe('executeDeleteTableRow', () => {
  const MULTI_ROW_TABLE = `
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr><w:tc><w:p w14:paraId="00000100"><w:r><w:t>R0C0</w:t></w:r></w:p></w:tc>
            <w:tc><w:p w14:paraId="00000101"><w:r><w:t>R0C1</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p w14:paraId="00000110"><w:r><w:t>R1C0</w:t></w:r></w:p></w:tc>
            <w:tc><w:p w14:paraId="00000111"><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p w14:paraId="00000120"><w:r><w:t>R2C0</w:t></w:r></w:p></w:tc>
            <w:tc><w:p w14:paraId="00000121"><w:r><w:t>R2C1</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p w14:paraId="00000130"><w:r><w:t>R3C0</w:t></w:r></w:p></w:tc>
            <w:tc><w:p w14:paraId="00000131"><w:r><w:t>R3C1</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  `;

  function rowCount(pkg: Package, tableId: string): number {
    const idx = buildIndices(pkg);
    const tbl = idx.tblIdx.get(tableId)!;
    let n = 0;
    for (let i = 0; i < tbl.childNodes.length; i++) {
      const c = tbl.childNodes[i] as Element;
      if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tr') n++;
    }
    return n;
  }

  test('删除中间行: 行数 -1, 剩余行内容不变', () => {
    const { pkg, ctx, tableId } = makeDoc(MULTI_ROW_TABLE);
    executeDeleteTableRow(ctx, {
      type: 'deleteTableRow',
      tableId,
      row: 1,
      onAnchorConflict: 'error',
    });
    expect(rowCount(pkg, tableId)).toBe(3);
    // 验证剩余行内容: R0, R2, R3
    const idx = buildIndices(pkg);
    const tbl = idx.tblIdx.get(tableId)!;
    const txt = tbl.textContent ?? '';
    expect(txt).toContain('R0C0');
    expect(txt).not.toContain('R1C0');
    expect(txt).toContain('R2C0');
    expect(txt).toContain('R3C0');
  });

  test('降序连续删: 同 plan 内删 row 3 → row 2, 索引不漂移', () => {
    const { pkg, ctx, tableId } = makeDoc(MULTI_ROW_TABLE);
    executeDeleteTableRow(ctx, {
      type: 'deleteTableRow',
      tableId,
      row: 3,
      onAnchorConflict: 'error',
    });
    executeDeleteTableRow(ctx, {
      type: 'deleteTableRow',
      tableId,
      row: 2,
      onAnchorConflict: 'error',
    });
    expect(rowCount(pkg, tableId)).toBe(2);
    const idx = buildIndices(pkg);
    const tbl = idx.tblIdx.get(tableId)!;
    const txt = tbl.textContent ?? '';
    expect(txt).toContain('R0C0');
    expect(txt).toContain('R1C0');
    expect(txt).not.toContain('R2C0');
    expect(txt).not.toContain('R3C0');
  });

  test('row 越界抛 PARA_NOT_FOUND', () => {
    const { ctx, tableId } = makeDoc(MULTI_ROW_TABLE);
    expect(() =>
      executeDeleteTableRow(ctx, {
        type: 'deleteTableRow',
        tableId,
        row: 99,
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('锚点冲突 error: 行内含书签 → 抛 ANCHOR_CONFLICT', () => {
    const tableWithBookmark = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr><w:tc><w:p w14:paraId="00000200"><w:r><w:t>row0</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p w14:paraId="00000201">
          <w:bookmarkStart w:id="0" w:name="rowmark"/>
          <w:r><w:t>row1</w:t></w:r>
          <w:bookmarkEnd w:id="0"/>
        </w:p></w:tc></w:tr>
      </w:tbl>
    `;
    const { ctx, tableId } = makeDoc(tableWithBookmark);
    expect(() =>
      executeDeleteTableRow(ctx, {
        type: 'deleteTableRow',
        tableId,
        row: 1,
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('锚点冲突 drop: 行内含书签 → 删行 + 清理书签', () => {
    const tableWithBookmark = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr><w:tc><w:p w14:paraId="00000300"><w:r><w:t>row0</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p w14:paraId="00000301">
          <w:bookmarkStart w:id="0" w:name="rowmark"/>
          <w:r><w:t>row1</w:t></w:r>
          <w:bookmarkEnd w:id="0"/>
        </w:p></w:tc></w:tr>
      </w:tbl>
    `;
    const { pkg, ctx, tableId } = makeDoc(tableWithBookmark);
    executeDeleteTableRow(ctx, {
      type: 'deleteTableRow',
      tableId,
      row: 1,
      onAnchorConflict: 'drop',
    });
    expect(rowCount(pkg, tableId)).toBe(1);
    // 书签索引应已清理
    const idx = buildIndices(pkg);
    expect(idx.bookmarkIdx.has('rowmark')).toBe(false);
  });

  test('trackChanges=true 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx, tableId } = makeDoc(MULTI_ROW_TABLE, { trackChanges: true });
    expect(() =>
      executeDeleteTableRow(ctx, {
        type: 'deleteTableRow',
        tableId,
        row: 1,
        onAnchorConflict: 'error',
      }),
    ).toThrowError(/trackChanges=true not yet implemented/);
  });
});
