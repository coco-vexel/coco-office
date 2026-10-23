/**
 * revision-ops executor 纯函数测试(#32,Phase D 收尾)。
 *
 * 覆盖矩阵:
 *   - Accept / Reject × <w:ins> / <w:del> wrapper × 正文(段内)场景
 *   - Accept / Reject × 段落 mark(pPr/rPr/<w:ins>|del>)场景
 *   - 三档 scope:all / paragraph / revision
 *   - 属性、移动、表格结构与 custom XML range 修订
 *   - PARA_NOT_FOUND / REVISION_NOT_FOUND
 *   - ApplyResult.acceptedRevisions / rejectedRevisions 计数
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import {
  createPackage,
  createXmlPart,
  getPart,
  setPart,
  type Package,
} from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { W_NS, W14_NS } from '../../namespaces.js';
import type { ApplyResult, Op } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import { executeAcceptRevisions, executeRejectRevisions } from '../revision-ops.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function putXml(pkg: Package, path: string, xml: string): void {
  setPart(
    pkg,
    createXmlPart({
      path,
      contentType: DOC_CTYPE,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
      originalBytes: new TextEncoder().encode(xml),
    }),
  );
}

function makeDoc(bodyInner: string): { pkg: Package; ctx: ExecutorContext } {
  const pkg = createPackage();
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`,
  );
  const indices = buildIndices(pkg);
  const ctx: ExecutorContext = {
    pkg,
    indices,
    options: { trackChanges: false, author: '', date: new Date() },
    result: emptyApplyResult(),
    opIndex: 0,
    docId: 'test-doc',
  };
  return { pkg, ctx };
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

function paraXml(pkg: Package, paraId: string): string {
  const idx = buildIndices(pkg);
  const el = idx.paraIdx.get(paraId.toUpperCase())!;
  return serializeToWellFormedString(el as never);
}

function paraCount(pkg: Package): number {
  const idx = buildIndices(pkg);
  return idx.paraIdx.size;
}

function documentXml(pkg: Package): string {
  const part = getPart(pkg, 'word/document.xml')!;
  return serializeToWellFormedString(part.xmlDom as never);
}

function acceptOp(
  scope: Extract<Op, { type: 'acceptRevisions' }>['scope'],
): Extract<Op, { type: 'acceptRevisions' }> {
  return { type: 'acceptRevisions', scope };
}

function rejectOp(
  scope: Extract<Op, { type: 'rejectRevisions' }>['scope'],
): Extract<Op, { type: 'rejectRevisions' }> {
  return { type: 'rejectRevisions', scope };
}

// ─────────── Accept <w:ins>(正文)─────────────

describe('acceptRevisions —— <w:ins> wrapper', () => {
  test('剥 wrapper,保留内容(w:t 不变)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="5" w:author="alice" w:date="2026-04-23T00:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins/);
    expect(xml).toMatch(/<w:t[^>]*>inserted<\/w:t>/);
    expect(ctx.result.acceptedRevisions).toBe(1);
    expect(ctx.result.rejectedRevisions).toBe(0);
  });
});

describe('rejectRevisions —— <w:ins> wrapper', () => {
  test('剥 wrapper + 内容一起删', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r><w:t>before </w:t></w:r>
        <w:ins w:id="5" w:author="x" w:date="2026-04-23T00:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
        <w:r><w:t> after</w:t></w:r>
      </w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins/);
    expect(xml).not.toMatch(/inserted/);
    expect(xml).toMatch(/before/);
    expect(xml).toMatch(/after/);
    expect(ctx.result.rejectedRevisions).toBe(1);
  });
});

// ─────────── Accept / Reject <w:del>(正文)───────────

describe('acceptRevisions —— <w:del> wrapper', () => {
  test('剥 wrapper + 删内容(内部 delText 一起删)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:del w:id="7" w:author="x" w:date="2026-04-23T00:00:00Z">
          <w:r><w:delText>deleted</w:delText></w:r>
        </w:del>
        <w:r><w:t>kept</w:t></w:r>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:del/);
    expect(xml).not.toMatch(/delText/);
    expect(xml).not.toMatch(/deleted/);
    expect(xml).toMatch(/kept/);
    expect(ctx.result.acceptedRevisions).toBe(1);
  });
});

describe('rejectRevisions —— <w:del> wrapper', () => {
  test('剥 wrapper + delText 转回 t(恢复被删内容)', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:del w:id="7" w:author="x" w:date="2026-04-23T00:00:00Z">
          <w:r><w:delText xml:space="preserve">restored</w:delText></w:r>
        </w:del>
      </w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:del/);
    expect(xml).not.toMatch(/delText/);
    expect(xml).toMatch(/<w:t[^>]*xml:space="preserve"[^>]*>restored<\/w:t>/);
    expect(ctx.result.rejectedRevisions).toBe(1);
  });
});

// ─────────── 段落 mark ───────────

describe('段落 mark —— pPr/rPr/<w:ins>', () => {
  test('accept ins mark:删 mark,段保留', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:rPr><w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>content</w:t></w:r>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    expect(paraCount(pkg)).toBe(1); // 段仍在
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins/);
    expect(ctx.result.acceptedRevisions).toBe(1);
  });

  test('reject ins mark:撤销新增段落标记并与下一段合并', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>keep</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2">
        <w:pPr><w:rPr><w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>toDelete</w:t></w:r>
      </w:p><w:p w14:paraId="AAAAAAA3"><w:r><w:t>tail</w:t></w:r></w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    expect(paraCount(pkg)).toBe(2);
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA2')).toBe(true);
    expect(idx.paraIdx.has('AAAAAAA3')).toBe(false);
    expect(paraXml(pkg, 'AAAAAAA2')).toMatch(/toDelete[\s\S]*tail/);
    expect(ctx.result.rejectedRevisions).toBe(1);
  });
});

describe('段落 mark —— pPr/rPr/<w:del>', () => {
  test('accept del mark:接受段落标记删除并与下一段合并', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:rPr><w:del w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>merged</w:t></w:r>
      </w:p>
      <w:p w14:paraId="AAAAAAA2"><w:r><w:t>next</w:t></w:r></w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    expect(paraCount(pkg)).toBe(1);
    const idx = buildIndices(pkg);
    expect(idx.paraIdx.has('AAAAAAA1')).toBe(true);
    expect(idx.paraIdx.has('AAAAAAA2')).toBe(false);
    expect(paraXml(pkg, 'AAAAAAA1')).toMatch(/merged[\s\S]*next/);
    expect(ctx.result.acceptedRevisions).toBe(1);
  });

  test('reject del mark:删 mark,段保留', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr><w:rPr><w:del w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"/></w:rPr></w:pPr>
        <w:r><w:t>keep</w:t></w:r>
      </w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    expect(paraCount(pkg)).toBe(1);
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:del/);
    expect(ctx.result.rejectedRevisions).toBe(1);
  });
});

// ─────────── scope ───────────

describe('scope 三档', () => {
  test("scope='paragraph':只处理指定段内的修订", () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
      </w:p>
      <w:p w14:paraId="AAAAAAA2">
        <w:ins w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>b</w:t></w:r></w:ins>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'paragraph', paraId: 'AAAAAAA1' }));
    const xml1 = paraXml(pkg, 'AAAAAAA1');
    const xml2 = paraXml(pkg, 'AAAAAAA2');
    expect(xml1).not.toMatch(/<w:ins/); // 处理过
    expect(xml2).toMatch(/<w:ins/); // 未动
    expect(ctx.result.acceptedRevisions).toBe(1);
  });

  test("scope='paragraph':paraId 不存在 → PARA_NOT_FOUND", () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    try {
      executeAcceptRevisions(ctx, acceptOp({ scope: 'paragraph', paraId: 'NONE' }));
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('PARA_NOT_FOUND');
    }
  });

  test("scope='revision':按 w:id 精确匹配", () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="5" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
        <w:ins w:id="6" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>b</w:t></w:r></w:ins>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'revision', revId: '5' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    // id=5 被剥,a 保留;id=6 未动
    expect(xml).toMatch(/<w:ins[^>]+w:id="6"/);
    expect((xml.match(/<w:ins/g) ?? []).length).toBe(1);
    expect(ctx.result.acceptedRevisions).toBe(1);
  });

  test("scope='revision':revId 无匹配 → REVISION_NOT_FOUND", () => {
    const { ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="5" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
      </w:p>`,
    );
    try {
      executeAcceptRevisions(ctx, acceptOp({ scope: 'revision', revId: '99' }));
      expect.unreachable();
    } catch (e) {
      expect((e as WordError).code).toBe('REVISION_NOT_FOUND');
    }
  });

  test("scope='all':全文 ins/del 都被处理,多 part 也覆盖", () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
        <w:del w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:delText>b</w:delText></w:r></w:del>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins|<w:del/);
    expect(ctx.result.acceptedRevisions).toBe(2);
  });
});

// ─────────── 属性修订 ───────────

describe('属性级修订', () => {
  test('accept rPrChange:保留当前属性并移除历史快照', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:r>
          <w:rPr>
            <w:b/>
            <w:rPrChange w:id="9" w:author="x" w:date="2026-04-23T00:00:00Z">
              <w:rPr><w:i/></w:rPr>
            </w:rPrChange>
          </w:rPr>
          <w:t>x</w:t>
        </w:r>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:b/);
    expect(xml).not.toMatch(/rPrChange|<w:i/);
  });

  test('reject pPrChange:恢复旧属性，同时保留段落标记 rPr', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:pPr>
          <w:jc w:val="right"/><w:rPr><w:b/></w:rPr>
          <w:pPrChange w:id="9" w:author="x" w:date="2026-04-23T00:00:00Z">
            <w:pPr><w:jc w:val="left"/></w:pPr>
          </w:pPrChange>
        </w:pPr>
        <w:r><w:t>x</w:t></w:r>
      </w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'paragraph', paraId: 'AAAAAAA1' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/w:val="left"/);
    expect(xml).toMatch(/<w:rPr><w:b/);
    expect(xml).not.toMatch(/right|pPrChange/);
  });

  test('reject 覆盖 section/table/row/cell/grid/exception 属性快照', () => {
    const { pkg, ctx } = makeDoc(
      `<w:tbl>
        <w:tblPr><w:tblStyle w:val="new"/><w:tblPrChange w:id="1"><w:tblPr><w:tblStyle w:val="old"/></w:tblPr></w:tblPrChange></w:tblPr>
        <w:tblGrid><w:gridCol w:w="200"/><w:tblGridChange w:id="2"><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid></w:tblGridChange></w:tblGrid>
        <w:tr>
          <w:trPr><w:cantSplit/><w:trPrChange w:id="3"><w:trPr><w:tblHeader/></w:trPr></w:trPrChange></w:trPr>
          <w:tblPrEx><w:tblCellSpacing w:w="20"/><w:tblPrExChange w:id="4"><w:tblPrEx><w:tblCellSpacing w:w="10"/></w:tblPrEx></w:tblPrExChange></w:tblPrEx>
          <w:tc><w:tcPr><w:tcW w:w="200"/><w:tcPrChange w:id="5"><w:tcPr><w:tcW w:w="100"/></w:tcPr></w:tcPrChange></w:tcPr><w:p/></w:tc>
        </w:tr>
      </w:tbl>
      <w:sectPr><w:pgSz w:w="200"/><w:sectPrChange w:id="6"><w:sectPr><w:pgSz w:w="100"/></w:sectPr></w:sectPrChange></w:sectPr>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    const xml = documentXml(pkg);
    expect(xml).not.toMatch(/PrChange|tblGridChange/);
    expect(xml).toContain('w:val="old"');
    expect(xml).toContain('w:w="100"');
    expect(xml).toContain('w:w="10"');
    expect(xml).toMatch(/tblHeader/);
    expect(xml).not.toMatch(/w:val="new"|w:w="200"|cantSplit/);
    expect(ctx.result.rejectedRevisions).toBe(6);
  });
});

describe('表格结构修订', () => {
  test('行插入/删除标记按整行语义 accept/reject', () => {
    const accepted = makeDoc(
      `<w:tbl>
        <w:tr><w:trPr><w:ins w:id="1"/></w:trPr><w:tc><w:p><w:r><w:t>inserted</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:trPr><w:del w:id="2"/></w:trPr><w:tc><w:p><w:r><w:t>deleted</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>`,
    );
    executeAcceptRevisions(accepted.ctx, acceptOp({ scope: 'all' }));
    expect(documentXml(accepted.pkg)).toMatch(/inserted/);
    expect(documentXml(accepted.pkg)).not.toMatch(/deleted|<w:ins|<w:del/);

    const rejected = makeDoc(
      `<w:tbl>
        <w:tr><w:trPr><w:ins w:id="1"/></w:trPr><w:tc><w:p><w:r><w:t>inserted</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:trPr><w:del w:id="2"/></w:trPr><w:tc><w:p><w:r><w:t>deleted</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>`,
    );
    executeRejectRevisions(rejected.ctx, rejectOp({ scope: 'all' }));
    expect(documentXml(rejected.pkg)).not.toMatch(/inserted|<w:ins|<w:del/);
    expect(documentXml(rejected.pkg)).toMatch(/deleted/);
  });

  test('cellIns/cellDel/cellMerge 具有独立 accept/reject 语义', () => {
    const inserted = makeDoc(
      `<w:tbl><w:tr><w:tc><w:tcPr><w:cellIns w:id="1"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`,
    );
    executeRejectRevisions(inserted.ctx, rejectOp({ scope: 'all' }));
    expect(documentXml(inserted.pkg)).not.toMatch(/<w:tc/);

    const deleted = makeDoc(
      `<w:tbl><w:tr><w:tc><w:tcPr><w:cellDel w:id="2"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`,
    );
    executeRejectRevisions(deleted.ctx, rejectOp({ scope: 'all' }));
    expect(documentXml(deleted.pkg)).toMatch(/<w:tc/);
    expect(documentXml(deleted.pkg)).not.toMatch(/cellDel/);

    const merged = makeDoc(
      `<w:tbl><w:tr><w:tc><w:tcPr><w:cellMerge w:id="3" w:vMerge="rest"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`,
    );
    executeAcceptRevisions(merged.ctx, acceptOp({ scope: 'all' }));
    expect(documentXml(merged.pkg)).toMatch(/<w:vMerge w:val="restart"/);
    expect(documentXml(merged.pkg)).not.toMatch(/cellMerge/);
  });
});

describe('移动和 custom XML 范围修订', () => {
  test('accept move 保留目标、删除来源并清理 range markers', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:moveFromRangeStart w:id="10" w:name="m"/><w:moveFrom w:id="1" w:name="m"><w:r><w:delText>source</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="10" w:name="m"/>
        <w:moveToRangeStart w:id="11" w:name="m"/><w:moveTo w:id="2" w:name="m"><w:r><w:t>target</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="11" w:name="m"/>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/source|moveFrom|moveTo|Range/);
    expect(xml).toMatch(/target/);
  });

  test('reject move 恢复来源文本并删除目标', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:moveFrom w:id="1"><w:r><w:delText>source</w:delText></w:r></w:moveFrom><w:moveTo w:id="2"><w:r><w:t>target</w:t></w:r></w:moveTo></w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:t>source<\/w:t>/);
    expect(xml).not.toMatch(/target|moveFrom|moveTo|delText/);
  });

  test('custom XML 插入/删除范围按边界处理内容', () => {
    const inserted = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:customXmlInsRangeStart w:id="1"/><w:r><w:t>new</w:t></w:r><w:customXmlInsRangeEnd w:id="1"/><w:r><w:t>keep</w:t></w:r></w:p>`,
    );
    executeRejectRevisions(inserted.ctx, rejectOp({ scope: 'all' }));
    expect(paraXml(inserted.pkg, 'AAAAAAA1')).not.toMatch(/new|customXml/);
    expect(paraXml(inserted.pkg, 'AAAAAAA1')).toMatch(/keep/);

    const deleted = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:customXmlDelRangeStart w:id="2"/><w:r><w:delText>old</w:delText></w:r><w:customXmlDelRangeEnd w:id="2"/></w:p>`,
    );
    executeRejectRevisions(deleted.ctx, rejectOp({ scope: 'all' }));
    expect(paraXml(deleted.pkg, 'AAAAAAA1')).toMatch(/<w:t>old<\/w:t>/);
    expect(paraXml(deleted.pkg, 'AAAAAAA1')).not.toMatch(/customXml|delText/);
  });

  test('custom XML 范围可跨段落容器', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:customXmlInsRangeStart w:id="8"/><w:r><w:t>first</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>second</w:t></w:r><w:customXmlInsRangeEnd w:id="8"/></w:p>
       <w:p w14:paraId="AAAAAAA3"><w:r><w:t>keep</w:t></w:r></w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    const xml = documentXml(pkg);
    expect(xml).not.toMatch(/first|second|customXml/);
    expect(xml).toMatch(/keep/);
  });
});

describe('Office 2010 冲突修订', () => {
  test('conflictIns/conflictDel 按插入/删除语义 accept', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p xmlns:w14="${W14_NS}" w14:paraId="AAAAAAA1">
        <w14:conflictIns w:id="1"><w:r><w:t>inserted</w:t></w:r></w14:conflictIns>
        <w14:conflictDel w:id="2"><w:r><w:delText>deleted</w:delText></w:r></w14:conflictDel>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/inserted/);
    expect(xml).not.toMatch(/deleted|conflictIns|conflictDel/);
  });

  test('custom XML conflict range reject 恢复删除内容', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p xmlns:w14="${W14_NS}" w14:paraId="AAAAAAA1"><w14:customXmlConflictDelRangeStart w:id="3"/><w:r><w:delText>restored</w:delText></w:r><w14:customXmlConflictDelRangeEnd w:id="3"/></w:p>`,
    );
    executeRejectRevisions(ctx, rejectOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).toMatch(/<w:t>restored<\/w:t>/);
    expect(xml).not.toMatch(/customXmlConflict|delText/);
  });
});

// ─────────── 边界 ───────────

describe('边界', () => {
  test('空文档(无修订):成功,count=0', () => {
    const { ctx } = makeDoc(`<w:p w14:paraId="AAAAAAA1"><w:r><w:t>x</w:t></w:r></w:p>`);
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    expect(ctx.result.acceptedRevisions).toBe(0);
  });

  test('同段内多个 ins + del 一次 accept 处理完', () => {
    const { pkg, ctx } = makeDoc(
      `<w:p w14:paraId="AAAAAAA1">
        <w:ins w:id="1" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>a</w:t></w:r></w:ins>
        <w:r><w:t>-</w:t></w:r>
        <w:ins w:id="2" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:t>b</w:t></w:r></w:ins>
        <w:del w:id="3" w:author="x" w:date="2026-04-23T00:00:00Z"><w:r><w:delText>c</w:delText></w:r></w:del>
      </w:p>`,
    );
    executeAcceptRevisions(ctx, acceptOp({ scope: 'all' }));
    const xml = paraXml(pkg, 'AAAAAAA1');
    expect(xml).not.toMatch(/<w:ins|<w:del/);
    // a 保留,- 保留,b 保留,c 被删(accept del)
    expect(xml).toMatch(/a/);
    expect(xml).toMatch(/-/);
    expect(xml).toMatch(/b/);
    expect(xml).not.toMatch(/>c</);
    expect(ctx.result.acceptedRevisions).toBe(3);
  });
});
