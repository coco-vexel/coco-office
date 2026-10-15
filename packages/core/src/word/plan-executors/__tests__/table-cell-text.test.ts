/**
 * executeReplaceTableCellText 单元测试(task-core-cell-text)。
 *
 * 覆盖:
 *   - 基本: 替换单元格首段 run 为纯文字 run，rPr 透传(字体/字号)
 *   - 底纹保留(核心验收点): tcPr 含 shd/vAlign，替换后原样保留
 *   - pPr 保留: 首段 jc 居中在替换后保留
 *   - 锚点冲突 error: 首段含书签 → 抛 ANCHOR_CONFLICT
 *   - 锚点冲突 drop: 首段含书签 → 跳过，原内容保留，不计数
 *   - cell 不存在 → TABLE_NOT_FOUND
 *   - trackChanges=true → UNSUPPORTED_OOXML_FEATURE
 *   - 多单元格: 连续替换不同 (row,col) 都正确
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeReplaceTableCellText } from '../table-ops.js';

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
      date: new Date('2026-06-03T00:00:00Z'),
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

/** 单元格整个 tc 序列化(用于断言 tcPr 保留) */
function cellTcXml(pkg: Package, tableId: string, row: number, col: number): string {
  const idx = buildIndices(pkg);
  const tbl = idx.tblIdx.get(tableId)!;
  const trs = tbl.getElementsByTagNameNS(W_NS, 'tr');
  const tr = trs[row]!;
  const tcs: Element[] = [];
  for (let i = 0; i < tr.childNodes.length; i++) {
    const c = tr.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tc') tcs.push(c);
  }
  return serializeToWellFormedString(tcs[col]! as never);
}

/** 单元格首段序列化 */
function cellFirstParaXml(pkg: Package, tableId: string, row: number, col: number): string {
  const idx = buildIndices(pkg);
  const tbl = idx.tblIdx.get(tableId)!;
  const trs = tbl.getElementsByTagNameNS(W_NS, 'tr');
  const tr = trs[row]!;
  const tcs: Element[] = [];
  for (let i = 0; i < tr.childNodes.length; i++) {
    const c = tr.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'tc') tcs.push(c);
  }
  const tc = tcs[col]!;
  for (let i = 0; i < tc.childNodes.length; i++) {
    const c = tc.childNodes[i] as Element;
    if (c && c.nodeType === 1 && c.namespaceURI === W_NS && c.localName === 'p') {
      return serializeToWellFormedString(c as never);
    }
  }
  throw new Error('no <w:p> in cell');
}

// ─────────── fixture：带底纹 / 对齐的表格 ───────────

/**
 * 核心 fixture：
 *   (0,0) 表头单元格：tcPr 含 shd(E0E0E0) + vAlign(center)，段落 jc=center，run 有 rFonts 黑体/sz18
 *   (0,1) 数据单元格：tcPr 含 tcW，无特殊 shd，run 无 rPr
 */
const STYLED_TABLE = `
  <w:tbl>
    <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="3000" w:type="dxa"/>
          <w:shd w:val="clear" w:fill="E0E0E0"/>
          <w:vAlign w:val="center"/>
        </w:tcPr>
        <w:p w14:paraId="00000001">
          <w:pPr><w:jc w:val="center"/></w:pPr>
          <w:r><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="18"/></w:rPr><w:t>旧文字</w:t></w:r>
        </w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000002"><w:r><w:t>数据值</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
    <w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000003"><w:r><w:t>行1列0</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
        <w:p w14:paraId="00000004"><w:r><w:t>行1列1</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
`;

describe('executeReplaceTableCellText', () => {
  test('基本: rPr 透传——黑体/sz18 出现在新 run 内', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 0,
      col: 0,
      text: '新文字',
      onAnchorConflict: 'error',
    });

    const xml = cellFirstParaXml(pkg, tableId, 0, 0);
    // 原文字已替换
    expect(xml).not.toContain('旧文字');
    expect(xml).toContain('新文字');
    // rPr 透传：黑体 + sz18 保留在新 run
    expect(xml).toContain('黑体');
    expect(xml).toContain('18');
    // 不包含 SDT
    expect(xml).not.toContain('<w:sdt');
  });

  test('★底纹保留★: tcPr 的 shd(E0E0E0) + vAlign(center) 替换后原样保留', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 0,
      col: 0,
      text: '回填内容',
      onAnchorConflict: 'error',
    });

    // 断言 tc 整体序列化（tcPr 在 tc 内，不在 p 内）
    const tcXml = cellTcXml(pkg, tableId, 0, 0);
    // 底纹 shd 完整保留
    expect(tcXml).toContain('E0E0E0');
    expect(tcXml).toMatch(/<w:shd\b/);
    // vAlign 完整保留
    expect(tcXml).toContain('center');
    expect(tcXml).toMatch(/<w:vAlign\b/);
  });

  test('pPr 保留: 首段 jc=center 替换后保留', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 0,
      col: 0,
      text: '居中文字',
      onAnchorConflict: 'error',
    });

    const xml = cellFirstParaXml(pkg, tableId, 0, 0);
    // pPr 中的 jc=center 保留
    expect(xml).toMatch(/<w:jc\s+w:val="center"/);
  });

  test('无 rPr 的单元格: 退化为裸 run（不报错）', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 0,
      col: 1, // 数据单元格，run 无 rPr
      text: '填入数据',
      onAnchorConflict: 'error',
    });

    const xml = cellFirstParaXml(pkg, tableId, 0, 1);
    expect(xml).toContain('填入数据');
    expect(xml).not.toContain('旧文字');
    // 裸 run：不含 rPr
    expect(xml).not.toMatch(/w:rPr/);
  });

  test('锚点冲突 error: 首段含书签 → 抛 ANCHOR_CONFLICT', () => {
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
      executeReplaceTableCellText(ctx, {
        type: 'replaceTableCellText',
        tableId,
        row: 0,
        col: 0,
        text: '新内容',
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('锚点冲突 drop: 首段含书签 → 跳过，原内容保留，不计数', () => {
    const tableWithBookmark = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:p w14:paraId="00000011">
              <w:bookmarkStart w:id="0" w:name="some"/>
              <w:r><w:t>保留内容</w:t></w:r>
              <w:bookmarkEnd w:id="0"/>
            </w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `;
    const { pkg, ctx, tableId } = makeDoc(tableWithBookmark);
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 0,
      col: 0,
      text: '不会写入',
      onAnchorConflict: 'drop',
    });
    const xml = cellFirstParaXml(pkg, tableId, 0, 0);
    expect(xml).toContain('保留内容');
    expect(xml).not.toContain('不会写入');
    // drop 不计数
    expect(ctx.result.replacedTableCells).toBe(0);
  });

  test('replacedTableCells 计数: drop 跳过不计，正常替换计入', () => {
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

    // (0,0) 无锚点 → 真正替换，计数 +1
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 0,
      col: 0,
      text: '填入',
      onAnchorConflict: 'drop',
    });
    expect(ctx.result.replacedTableCells).toBe(1);

    // (1,0) 有书签 + drop → 静默跳过，计数不变
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 1,
      col: 0,
      text: '跳过',
      onAnchorConflict: 'drop',
    });
    expect(ctx.result.replacedTableCells).toBe(1);
  });

  test('cell 不存在(row 越界) → TABLE_NOT_FOUND', () => {
    const { ctx, tableId } = makeDoc(STYLED_TABLE);
    expect(() =>
      executeReplaceTableCellText(ctx, {
        type: 'replaceTableCellText',
        tableId,
        row: 99,
        col: 0,
        text: 'x',
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('cell 不存在(col 越界) → TABLE_NOT_FOUND', () => {
    const { ctx, tableId } = makeDoc(STYLED_TABLE);
    expect(() =>
      executeReplaceTableCellText(ctx, {
        type: 'replaceTableCellText',
        tableId,
        row: 0,
        col: 99,
        text: 'x',
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('tableId 不存在 → TABLE_NOT_FOUND', () => {
    const { ctx } = makeDoc(STYLED_TABLE);
    expect(() =>
      executeReplaceTableCellText(ctx, {
        type: 'replaceTableCellText',
        tableId: 'tbl-9999',
        row: 0,
        col: 0,
        text: 'x',
        onAnchorConflict: 'error',
      }),
    ).toThrowError(WordError);
  });

  test('trackChanges=true → UNSUPPORTED_OOXML_FEATURE', () => {
    const { ctx, tableId } = makeDoc(STYLED_TABLE, { trackChanges: true });
    expect(() =>
      executeReplaceTableCellText(ctx, {
        type: 'replaceTableCellText',
        tableId,
        row: 0,
        col: 0,
        text: 'x',
        onAnchorConflict: 'error',
      }),
    ).toThrowError(/trackChanges=true not yet implemented/);
  });

  test('多单元格: 连续替换不同 (row,col) 都正确，互不干扰', () => {
    const { pkg, ctx, tableId } = makeDoc(STYLED_TABLE);

    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 0,
      col: 0,
      text: '第一格',
      onAnchorConflict: 'error',
    });
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 0,
      col: 1,
      text: '第二格',
      onAnchorConflict: 'error',
    });
    executeReplaceTableCellText(ctx, {
      type: 'replaceTableCellText',
      tableId,
      row: 1,
      col: 0,
      text: '第三格',
      onAnchorConflict: 'error',
    });

    expect(cellFirstParaXml(pkg, tableId, 0, 0)).toContain('第一格');
    expect(cellFirstParaXml(pkg, tableId, 0, 1)).toContain('第二格');
    expect(cellFirstParaXml(pkg, tableId, 1, 0)).toContain('第三格');
    // 未动的单元格保留原内容
    expect(cellFirstParaXml(pkg, tableId, 1, 1)).toContain('行1列1');
    // 计数 3
    expect(ctx.result.replacedTableCells).toBe(3);
  });
});
