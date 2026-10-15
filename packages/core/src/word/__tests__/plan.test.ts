/**
 * WordPlan builder + apply 状态机测试(#27)。
 *
 * 覆盖:
 *   - 11 个链式 builder 方法:返回 this,push 正确的 op 到 _ops
 *   - ops 默认 / 默认 scope('first' / 'error')/ addComment Date→ISO 序列化
 *   - applied 状态:初始 false;apply 进入后立即 true;方法被 apply 后再调抛
 *   - options:trackChanges 默认 false;author 走 setDefaultAuthor / PlanOptions.author
 *   - 多 plan 独立:doc.plan() 每次新实例,ops 不串
 *   - 空 plan apply:warn + 返回空 ApplyResult(字段全到位)
 *   - 二次 apply:PLAN_ALREADY_APPLIED(同一 plan)
 *   - 非空 plan 现阶段抛 UNSUPPORTED_OOXML_FEATURE(契约 stub 行为,#28 起替换)
 *
 * Worker 端到端见 [plan-worker.test.ts](./plan-worker.test.ts)。
 */
import { afterAll, afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
import { W14_NS, W_NS } from '../namespaces.js';

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

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body><w:p w14:paraId="AAAAAAA1"><w:r><w:t>plan-fixture</w:t></w:r></w:p></w:body>
</w:document>`;

async function makeDoc(docId = 'plan-test'): Promise<WordDocument> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', DOC_XML);
  const buf = await zip.generateAsync({ type: 'uint8array' });
  return WordDocument.fromZip(buf, { docId, mode: 'inline' });
}

// Avoid real console.warn 染污测试输出;仅让测试可检查
let warnSpy: ReturnType<typeof vi.spyOn>;
beforeEach(() => {
  warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
});
afterEach(() => {
  warnSpy.mockRestore();
});
afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── builder ───────────

describe('WordPlan builder —— 链式返回 + op push', () => {
  test('每个方法都返回 this,支持完整链式', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    // 同一 plan 多次调用应收集 op
    const returned = plan
      .replaceParagraph('A', 'x')
      .replaceTextInParagraph('A', 'foo', 'bar')
      .insertParagraphBefore('A', { text: 'new' })
      .insertParagraphAfter('A', { text: 'new', style: 'Heading1' })
      .deleteParagraph('A')
      .acceptRevisions({ scope: 'all' })
      .rejectRevisions({ scope: 'paragraph', paraId: 'A' })
      .createBookmark({ name: 'b1', range: { paraId: 'A' } })
      .deleteBookmark('b1')
      .addComment({ range: { paraId: 'A' }, text: '批注' })
      .deleteComment(5);
    expect(returned).toBe(plan);
    expect(plan.ops).toHaveLength(11);
    await doc.dispose();
  });

  test('replaceTextInParagraph scope 默认 "first"', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    plan.replaceTextInParagraph('A', 's', 'r');
    const op = plan.ops[0] as Extract<
      (typeof plan.ops)[number],
      { type: 'replaceTextInParagraph' }
    >;
    expect(op.scope).toBe('first');

    plan.replaceTextInParagraph('A', 's', 'r', { scope: 'all' });
    const op2 = plan.ops[1] as Extract<
      (typeof plan.ops)[number],
      { type: 'replaceTextInParagraph' }
    >;
    expect(op2.scope).toBe('all');
    await doc.dispose();
  });

  test('formatTextInParagraph scope defaults to "first"', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    const returned = plan.formatTextInParagraph('A', 's', {
      bold: true,
      color: 'FF0000',
      italic: true,
      underline: true,
      highlight: 'yellow',
      fontSize: 12,
      fontFamily: 'Arial',
    });
    const op = plan.ops[0] as Extract<(typeof plan.ops)[number], { type: 'formatTextInParagraph' }>;
    expect(returned).toBe(plan);
    expect(op).toMatchObject({
      type: 'formatTextInParagraph',
      paraId: 'A',
      search: 's',
      format: {
        bold: true,
        color: 'FF0000',
        italic: true,
        underline: true,
        highlight: 'yellow',
        fontSize: 12,
        fontFamily: 'Arial',
      },
      scope: 'first',
    });

    plan.formatTextInParagraph('A', 's', { color: '00AA00' }, { scope: 'all' });
    const op2 = plan.ops[1] as Extract<
      (typeof plan.ops)[number],
      { type: 'formatTextInParagraph' }
    >;
    expect(op2.scope).toBe('all');
    await doc.dispose();
  });

  test('deleteParagraph onAnchorConflict 默认 "error"', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    plan.deleteParagraph('A');
    const op = plan.ops[0] as Extract<(typeof plan.ops)[number], { type: 'deleteParagraph' }>;
    expect(op.onAnchorConflict).toBe('error');
    await doc.dispose();
  });

  test('formatParagraph pushes direct paragraph format op', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    const returned = plan.formatParagraph('A', {
      alignment: 'center',
      outlineLevel: 1,
      pageBreakBefore: true,
      indent: { left: 720, right: 120, firstLine: 360 },
      spacing: { before: 120, after: 240, line: 360, lineRule: 'exact' },
    });
    const op = plan.ops[0] as Extract<(typeof plan.ops)[number], { type: 'formatParagraph' }>;
    expect(returned).toBe(plan);
    expect(op).toEqual({
      type: 'formatParagraph',
      paraId: 'A',
      format: {
        alignment: 'center',
        outlineLevel: 1,
        pageBreakBefore: true,
        indent: { left: 720, right: 120, firstLine: 360 },
        spacing: { before: 120, after: 240, line: 360, lineRule: 'exact' },
      },
    });
    await doc.dispose();
  });

  test('formatTableCell pushes direct cell format op', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    const returned = plan.formatTableCell({
      tableId: 'tbl-0001',
      row: 0,
      col: 1,
      format: {
        width: 2400,
        shading: { fill: 'FFF2CC' },
        verticalAlignment: 'center',
      },
    });
    const op = plan.ops[0] as Extract<(typeof plan.ops)[number], { type: 'formatTableCell' }>;
    expect(returned).toBe(plan);
    expect(op).toEqual({
      type: 'formatTableCell',
      tableId: 'tbl-0001',
      row: 0,
      col: 1,
      format: {
        width: 2400,
        shading: { fill: 'FFF2CC' },
        verticalAlignment: 'center',
      },
    });
    await doc.dispose();
  });

  test('insertTableCellParagraph pushes cell paragraph insert op', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    const returned = plan.insertTableCellParagraph({
      tableId: 'tbl-0001',
      row: 1,
      col: 2,
      text: 'inside cell',
      position: 'end',
    });
    const op = plan.ops[0] as Extract<
      (typeof plan.ops)[number],
      { type: 'insertTableCellParagraph' }
    >;
    expect(returned).toBe(plan);
    expect(op).toEqual({
      type: 'insertTableCellParagraph',
      tableId: 'tbl-0001',
      row: 1,
      col: 2,
      text: 'inside cell',
      position: 'end',
    });
    await doc.dispose();
  });

  test('addComment:Date → ISO 字符串,未传 date 则字段缺', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    plan.addComment({
      range: { paraId: 'A' },
      text: 't',
      author: 'alice',
      date: new Date('2026-04-22T10:00:00.000Z'),
    });
    const op1 = plan.ops[0] as Extract<(typeof plan.ops)[number], { type: 'addComment' }>;
    expect(op1.date).toBe('2026-04-22T10:00:00.000Z');
    expect(op1.author).toBe('alice');

    plan.addComment({ range: { paraId: 'A' }, text: 't2' });
    const op2 = plan.ops[1] as Extract<(typeof plan.ops)[number], { type: 'addComment' }>;
    expect(Object.prototype.hasOwnProperty.call(op2, 'date')).toBe(false);
    expect(Object.prototype.hasOwnProperty.call(op2, 'author')).toBe(false);
    await doc.dispose();
  });

  test('ops 数组是只读快照(类型上),但底层实现用同一数组 —— 不要求 immutable', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    plan.replaceParagraph('A', 'x');
    const opsRef = plan.ops;
    plan.replaceParagraph('A', 'y');
    // 同引用,追加可见(只读在类型层面,运行时不 copy)
    expect(opsRef).toHaveLength(2);
    await doc.dispose();
  });
});

// ─────────── options ───────────

describe('WordPlan options —— 默认与覆盖', () => {
  test('trackChanges 缺省 false', async () => {
    const doc = await makeDoc();
    expect(doc.plan().options.trackChanges).toBe(false);
    expect(doc.plan({ trackChanges: true }).options.trackChanges).toBe(true);
    await doc.dispose();
  });

  test('author 回落顺序:opts.author > setDefaultAuthor > ""', async () => {
    const doc = await makeDoc();
    // 未设任何 author
    expect(doc.plan().options.author).toBeUndefined();

    WordDocument.setDefaultAuthor('global-default');
    expect(doc.plan().options.author).toBe('global-default');

    expect(doc.plan({ author: 'override' }).options.author).toBe('override');

    WordDocument._resetDefaults();
    await doc.dispose();
  });

  test('options 对象是 frozen,外部不可改', async () => {
    const doc = await makeDoc();
    const plan = doc.plan({ trackChanges: true, author: 'a' });
    expect(Object.isFrozen(plan.options)).toBe(true);
    await doc.dispose();
  });
});

// ─────────── applied 状态 ───────────

describe('WordPlan applied 状态', () => {
  test('初始 false;空 plan apply 成功后 true', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    expect(plan.applied).toBe(false);
    await plan.apply();
    expect(plan.applied).toBe(true);
    await doc.dispose();
  });

  test('二次 apply → PLAN_ALREADY_APPLIED,docId 正确', async () => {
    const doc = await makeDoc('plan-twice');
    const plan = doc.plan();
    await plan.apply();
    await expect(plan.apply()).rejects.toBeInstanceOf(WordError);
    try {
      await plan.apply();
    } catch (e) {
      const err = e as WordError;
      expect(err.code).toBe('PLAN_ALREADY_APPLIED');
      expect(err.docId).toBe('plan-twice');
    }
    await doc.dispose();
  });

  test('apply 后再调 builder 方法 → PLAN_ALREADY_APPLIED(单次锁语义)', async () => {
    const doc = await makeDoc();
    const plan = doc.plan();
    await plan.apply();
    try {
      plan.replaceParagraph('A', 'x');
      expect.unreachable('builder after apply should throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('PLAN_ALREADY_APPLIED');
    }
    await doc.dispose();
  });
});

// ─────────── 多 plan 独立 ───────────

describe('WordPlan 多实例独立', () => {
  test('两次 doc.plan() 返回不同实例,ops 互不串', async () => {
    const doc = await makeDoc();
    const p1 = doc.plan();
    const p2 = doc.plan();
    p1.replaceParagraph('A', 'x');
    p2.deleteBookmark('b');
    expect(p1).not.toBe(p2);
    expect(p1.ops).toHaveLength(1);
    expect(p2.ops).toHaveLength(1);
    expect(p1.ops[0]!.type).toBe('replaceParagraph');
    expect(p2.ops[0]!.type).toBe('deleteBookmark');
    await doc.dispose();
  });
});

// ─────────── apply 行为 ───────────

describe('WordPlan apply —— 空 plan / 非空 plan', () => {
  test('空 plan:console.warn + 返回空 ApplyResult(字段齐全)', async () => {
    const doc = await makeDoc('empty-plan');
    const plan = doc.plan();
    const result = await plan.apply();

    expect(warnSpy).toHaveBeenCalledTimes(1);
    expect(warnSpy.mock.calls[0]![0]).toContain('empty plan');
    expect(warnSpy.mock.calls[0]![0]).toContain('empty-plan');

    expect(result).toEqual({
      insertedParagraphs: [],
      insertedBodyStartParaIds: [],
      createdComments: [],
      createdBookmarks: [],
      acceptedRevisions: 0,
      rejectedRevisions: 0,
      replacedTableCells: 0,
    });
    await doc.dispose();
  });

  // 曾用作"有 op 未实现"的哨兵测试(指向 replaceParagraph → createBookmark →
  // addComment → acceptRevisions);#32 合入后 11 个 op 全部有 executor,哨兵失效,
  // 测试已删除。未来新增 op 时如仍想要这类哨兵,可用假 op(type 不在 union 里)
  // 触发 dispatchOp 的 _never 分支。
});
