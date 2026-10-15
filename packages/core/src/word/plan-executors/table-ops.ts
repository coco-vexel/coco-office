/**
 * 表格 executor —— #186(addTableComment / deleteTable)。
 *
 * 设计要点:
 *   - 入口都是 `tableId`(装载期分配的 'tbl-NNNN',来源:`indices.tblIdx`),不接
 *     Element / paraId,避免调用方自己绕 DOM 反查。
 *   - addTableComment:语义"给表格挂批注",实现"挂在表格首单元格首段",复用
 *     `attachCommentToParagraph`。这条决定由 docs/07 §06 推论:批注是 #186 决策
 *     理由的可追溯载体,挂表格内段落不会污染表外正文,删表时 drop 模式还能连带
 *     清理(锚点都在表内段落上)。
 *   - deleteTable:扫表内所有段落收锚点(书签 + 批注 ID),无锚点直接 removeChild;
 *     有锚点按 onAnchorConflict='error'(默认)抛 ANCHOR_CONFLICT,'drop' 则连带
 *     清理后删表。drop 实现复用段落 deleteParagraph 的清理路径(逐段调,带 drop)。
 *   - trackChanges=true 暂不支持(同 deleteParagraph 行为,留 follow-up)。
 *   - setTableColumnWidths(task-core-table-grid)：只设 tblGrid gridCol w:w，
 *     不动单元格 tcW（本 case 表格 tblLayout=fixed 且无 tcW；若将来 auto-layout
 *     带 tcW 需同步，记 follow-up，不在本 task 范围）。
 */
import { getPart, isXmlPart, listParts, markDirty } from '../../pack/index.js';
import { findStoryFor } from '../stories.js';
import { W_NS, W14_NS, XML_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { WordError } from '../word-error.js';
import { attachCommentToParagraph, cleanupEmptyCommentsPart } from './comment-ops.js';
import { buildInlineSdt, hashSdtId } from '../sdt-builders.js';
import { appendTextAsRuns, collectParagraphAnchors, pickFirstRunRPr } from './_shared.js';
import type { ExecutorContext } from './executor-context.js';
import { generateParaId } from '../para-id.js';

const COMMENTS_PART_PATH = 'word/comments.xml';

// ─────────── addTableComment ───────────

export function executeAddTableComment(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'addTableComment' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'addTableComment with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'addTableComment', tableId: op.tableId },
    );
  }

  const tbl = requireTable(ctx, op.tableId, 'addTableComment');
  const firstPara = findFirstParagraphInTable(tbl);
  if (!firstPara) {
    // 健壮性降级（绝不崩）：OOXML 名义上要求每个 <w:tc> 至少一个 <w:p>，但真实文档
    // （poi2-Bug54771a：单元格无 <w:p>）会违反。批注是诊断辅助（题注/角色标注），
    // 无可挂载段时跳过该条批注即可——不影响文档主体内容，绝不为此回滚整个 plan 而崩。
    // 用 warn 显式记录降级（不静默吞），保持可观测。
    console.warn(
      `[word] addTableComment: table '${op.tableId}' contains no <w:p>; ` +
        `skip this comment (degrade, docId: ${ctx.docId})`,
    );
    return;
  }

  attachCommentToParagraph(ctx, firstPara, op.text, {
    author: op.author,
    date: op.date,
  });
  // 不写 ApplyResult.createdComments —— addTableComment 不像 addComment 那样
  // 要把 id 回传给调用方(调用方握的是 tableId,不需要 commentId)。
}

// ─────────── deleteTable ───────────

export function executeDeleteTable(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'deleteTable' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'deleteTable with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'deleteTable', tableId: op.tableId },
    );
  }

  const tbl = requireTable(ctx, op.tableId, 'deleteTable');

  // 收集表内所有活锚点(书签 / 批注)
  const affectedBookmarks = collectBookmarkNamesInTable(tbl);
  const affectedComments = collectCommentIdsInTable(tbl);
  const hasAnchors = affectedBookmarks.length > 0 || affectedComments.length > 0;

  if (hasAnchors) {
    if (op.onAnchorConflict === 'error') {
      throw new WordError(
        'ANCHOR_CONFLICT',
        `table '${op.tableId}' contains active anchors (bookmarks: ${affectedBookmarks.join(', ') || 'none'}, comments: ${affectedComments.join(', ') || 'none'})`,
        ctx.docId,
        {
          tableId: op.tableId,
          affectedBookmarks,
          affectedComments,
          opType: 'deleteTable',
        },
      );
    }
    // drop:连带清理
    for (const name of affectedBookmarks) {
      dropBookmarkByName(ctx, name);
    }
    for (const id of affectedComments) {
      dropCommentById(ctx, id, tbl);
    }
  }

  // markDirty 必须在 removeChild 之前(removeChild 后 findStoryFor 找不到 tbl)
  markStoryPartDirty(ctx, tbl, 'deleteTable');

  const parent = tbl.parentNode;
  if (!parent) {
    throw new WordError('PACKAGE_CORRUPT', `table '${op.tableId}' has no parentNode`, ctx.docId, {
      tableId: op.tableId,
    });
  }
  parent.removeChild(tbl);

  // 索引同步:applyPlan finally 块会 buildIndices 全量重建,这里先从 tblIdx 删,
  // 避免后续 op 在同一 plan 内复用已游离的 tblIdx 条目。
  ctx.indices.tblIdx.delete(op.tableId);
}

// ─────────── replaceTableCellWithSdt(#188)───────────

export function executeReplaceTableCellWithSdt(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'replaceTableCellWithSdt' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'replaceTableCellWithSdt with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'replaceTableCellWithSdt', tableId: op.tableId },
    );
  }

  const tbl = requireTable(ctx, op.tableId, 'replaceTableCellWithSdt');
  const tc = findCell(tbl, op.row, op.col);
  if (!tc) {
    throw new WordError(
      'TABLE_NOT_FOUND',
      `cell at (${op.row}, ${op.col}) not found in table '${op.tableId}'`,
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, opType: 'replaceTableCellWithSdt' },
    );
  }

  const firstPara = firstChildW(tc, 'p');
  if (!firstPara) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `cell at (${op.row}, ${op.col}) in table '${op.tableId}' has no <w:p>`,
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col },
    );
  }

  // 锚点冲突检测 (与 replaceParagraph 行为对齐)
  const anchors = collectParagraphAnchors(firstPara);
  if (anchors.bookmarks.length > 0 || anchors.comments.length > 0) {
    if (op.onAnchorConflict === 'drop') {
      // drop: 跳过本次替换, 保留单元格内容不变。
      // 不计入 result.replacedTableCells (#204), phase 据此报告准确域化数。
      return;
    }
    throw new WordError(
      'ANCHOR_CONFLICT',
      `table cell (${op.row}, ${op.col}) modification conflicts with active anchors ` +
        `in table '${op.tableId}'; enable trackChanges or use onAnchorConflict='drop'`,
      ctx.docId,
      {
        tableId: op.tableId,
        row: op.row,
        col: op.col,
        affectedBookmarks: anchors.bookmarks,
        affectedComments: anchors.comments,
        opType: 'replaceTableCellWithSdt',
      },
    );
  }

  // 从原首段首 run 取 rPr 透传 (字体 / 颜色 / 字号不丢)
  const inheritedRPr = pickFirstRunRPr(firstPara);

  // 替换段落内容: 保留 pPr, 移除其余子节点, 追加 SDT
  const doc = firstPara.ownerDocument;
  if (!doc) {
    throw new WordError('PACKAGE_CORRUPT', 'paragraph has no ownerDocument', ctx.docId, {
      tableId: op.tableId,
      opType: 'replaceTableCellWithSdt',
    });
  }

  const id = hashSdtId(op.tableId, op.row, op.col, op.sdt.alias);
  const sdt = buildInlineSdt(doc, op.sdt, id, inheritedRPr);

  removeNonPPrChildren(firstPara);
  firstPara.appendChild(sdt);

  markStoryPartDirty(ctx, firstPara, 'replaceTableCellWithSdt');

  // #204: 真正替换完成才计数 (drop 模式下 anchor 冲突已 return, 不会到这)
  ctx.result.replacedTableCells++;
}

// ─────────── replaceTableCellText(填纯文字,保留 tcPr/rPr)───────────

export function executeReplaceTableCellText(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'replaceTableCellText' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'replaceTableCellText with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'replaceTableCellText', tableId: op.tableId },
    );
  }

  const tbl = requireTable(ctx, op.tableId, 'replaceTableCellText');
  const tc = findCell(tbl, op.row, op.col);
  if (!tc) {
    throw new WordError(
      'TABLE_NOT_FOUND',
      `cell at (${op.row}, ${op.col}) not found in table '${op.tableId}'`,
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, opType: 'replaceTableCellText' },
    );
  }

  const firstPara = firstChildW(tc, 'p');
  if (!firstPara) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `cell at (${op.row}, ${op.col}) in table '${op.tableId}' has no <w:p>`,
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col },
    );
  }

  // 锚点冲突检测(与 replaceTableCellWithSdt 语义对齐)
  const anchors = collectParagraphAnchors(firstPara);
  if (anchors.bookmarks.length > 0 || anchors.comments.length > 0) {
    if (op.onAnchorConflict === 'drop') {
      // drop:跳过本次替换,保留单元格内容不变。不计入 result.replacedTableCells。
      return;
    }
    throw new WordError(
      'ANCHOR_CONFLICT',
      `table cell (${op.row}, ${op.col}) modification conflicts with active anchors ` +
        `in table '${op.tableId}'; enable trackChanges or use onAnchorConflict='drop'`,
      ctx.docId,
      {
        tableId: op.tableId,
        row: op.row,
        col: op.col,
        affectedBookmarks: anchors.bookmarks,
        affectedComments: anchors.comments,
        opType: 'replaceTableCellText',
      },
    );
  }

  // 从原首段首 run 取 rPr 透传(字体 / 颜色 / 字号不丢)
  const inheritedRPr = pickFirstRunRPr(firstPara);

  // 替换段落内容:保留 pPr,移除其余子节点,追加纯文字 run
  const doc = firstPara.ownerDocument;
  if (!doc) {
    throw new WordError('PACKAGE_CORRUPT', 'paragraph has no ownerDocument', ctx.docId, {
      tableId: op.tableId,
      opType: 'replaceTableCellText',
    });
  }

  // 构造 <w:r>{rPr?}<w:t xml:space="preserve">{text}</w:t></w:r>
  const run = doc.createElementNS(W_NS, 'w:r');
  if (inheritedRPr) {
    run.appendChild(inheritedRPr); // 已是 cloneNode(true),直接追加
  }
  const t = doc.createElementNS(W_NS, 'w:t');
  t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  t.textContent = op.text;
  run.appendChild(t);

  removeNonPPrChildren(firstPara);
  firstPara.appendChild(run);

  markStoryPartDirty(ctx, firstPara, 'replaceTableCellText');

  // 真正替换完成才计数(drop 模式下 anchor 冲突已 return,不会到这)
  ctx.result.replacedTableCells++;
}

// ─────────── deleteTableRow(#188 §7.2)───────────

export function executeFormatTableCell(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'formatTableCell' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'formatTableCell does not support trackChanges yet',
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, opType: 'formatTableCell' },
    );
  }

  validateTableCellFormat(ctx, op);

  const tbl = requireTable(ctx, op.tableId, 'formatTableCell');
  const tc = findCell(tbl, op.row, op.col);
  if (!tc) {
    throw new WordError(
      'TABLE_NOT_FOUND',
      `cell at (${op.row}, ${op.col}) not found in table '${op.tableId}'`,
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, opType: 'formatTableCell' },
    );
  }

  const tcPr = ensureTableCellProperties(tc);

  if (op.format.width !== undefined) {
    const tcW = upsertTcPrChild(tcPr, 'tcW', TCW_AFTER_LOCALS);
    tcW.setAttributeNS(W_NS, 'w:w', String(op.format.width));
    tcW.setAttributeNS(W_NS, 'w:type', 'dxa');
  }

  if (op.format.shading !== undefined) {
    const shd = upsertTcPrChild(tcPr, 'shd', SHD_AFTER_LOCALS);
    if (op.format.shading.val !== undefined) {
      shd.setAttributeNS(W_NS, 'w:val', op.format.shading.val);
    } else if (op.format.shading.fill !== undefined && !shd.hasAttributeNS(W_NS, 'val')) {
      shd.setAttributeNS(W_NS, 'w:val', 'clear');
    }
    if (op.format.shading.color !== undefined) {
      shd.setAttributeNS(W_NS, 'w:color', op.format.shading.color);
    }
    if (op.format.shading.fill !== undefined) {
      shd.setAttributeNS(W_NS, 'w:fill', op.format.shading.fill);
    }
  }

  if (op.format.verticalAlignment !== undefined) {
    const vAlign = upsertTcPrChild(tcPr, 'vAlign', VALIGN_AFTER_LOCALS);
    vAlign.setAttributeNS(W_NS, 'w:val', op.format.verticalAlignment);
  }

  markStoryPartDirty(ctx, tc, 'formatTableCell');
}

export function executeInsertTableCellParagraph(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'insertTableCellParagraph' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'insertTableCellParagraph does not support trackChanges yet',
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, opType: 'insertTableCellParagraph' },
    );
  }

  if (typeof op.text !== 'string') {
    throw new WordError(
      'INVALID_ARGUMENT',
      'insertTableCellParagraph text must be a string',
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, opType: 'insertTableCellParagraph' },
    );
  }
  if (op.position !== 'start' && op.position !== 'end') {
    throw new WordError(
      'INVALID_ARGUMENT',
      'insertTableCellParagraph position must be start or end',
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, position: op.position },
    );
  }

  const tbl = requireTable(ctx, op.tableId, 'insertTableCellParagraph');
  const tc = findCell(tbl, op.row, op.col);
  if (!tc) {
    throw new WordError(
      'TABLE_NOT_FOUND',
      `cell at (${op.row}, ${op.col}) not found in table '${op.tableId}'`,
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, opType: 'insertTableCellParagraph' },
    );
  }

  const doc = tc.ownerDocument;
  if (!doc) {
    throw new WordError('PACKAGE_CORRUPT', 'table cell has no ownerDocument', ctx.docId, {
      tableId: op.tableId,
      row: op.row,
      col: op.col,
      opType: 'insertTableCellParagraph',
    });
  }

  const seen = new Set(ctx.indices.paraIdx.keys());
  const newId = generateParaId(seen);
  const newPara = doc.createElementNS(W_NS, 'w:p');
  newPara.setAttributeNS(W14_NS, 'w14:paraId', newId);
  appendTextAsRuns(newPara, op.text);

  if (op.position === 'start') {
    tc.insertBefore(newPara, firstCellContentChild(tc));
  } else {
    tc.appendChild(newPara);
  }

  ctx.indices.paraIdx.set(newId.toUpperCase(), newPara);
  ctx.result.insertedParagraphs.push({
    opIndex: ctx.opIndex,
    anchorParaId: newId,
    position: op.position === 'start' ? 'before' : 'after',
    newParaId: newId,
  });
  markStoryPartDirty(ctx, tc, 'insertTableCellParagraph');
}

export function executeDeleteTableRow(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'deleteTableRow' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'deleteTableRow with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'deleteTableRow', tableId: op.tableId, row: op.row },
    );
  }

  const tbl = requireTable(ctx, op.tableId, 'deleteTableRow');
  const tr = findTableRow(tbl, op.row);
  if (!tr) {
    throw new WordError(
      'TABLE_NOT_FOUND',
      `row ${op.row} not found in table '${op.tableId}'`,
      ctx.docId,
      { tableId: op.tableId, row: op.row, opType: 'deleteTableRow' },
    );
  }

  // 锚点冲突: 行内段落含书签 / 批注时按 onAnchorConflict 处理
  const affectedBookmarks: string[] = [];
  {
    const list = tr.getElementsByTagNameNS(W_NS, 'bookmarkStart');
    for (let i = 0; i < list.length; i++) {
      const el = list[i];
      if (!el) continue;
      const name = el.getAttributeNS(W_NS, 'name') ?? el.getAttribute('w:name') ?? '';
      affectedBookmarks.push(name);
    }
  }
  const affectedComments: number[] = [];
  {
    const idSet = new Set<number>();
    for (const local of ['commentRangeStart', 'commentReference']) {
      const list = tr.getElementsByTagNameNS(W_NS, local);
      for (let i = 0; i < list.length; i++) {
        const el = list[i];
        if (!el) continue;
        const raw = el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '';
        const id = parseInt(raw, 10);
        if (!isNaN(id)) idSet.add(id);
      }
    }
    for (const id of idSet) affectedComments.push(id);
  }

  if (affectedBookmarks.length > 0 || affectedComments.length > 0) {
    if (op.onAnchorConflict === 'error') {
      throw new WordError(
        'ANCHOR_CONFLICT',
        `table row (${op.tableId} row ${op.row}) contains active anchors ` +
          `(bookmarks: ${affectedBookmarks.join(', ') || 'none'}, ` +
          `comments: ${affectedComments.join(', ') || 'none'})`,
        ctx.docId,
        {
          tableId: op.tableId,
          row: op.row,
          affectedBookmarks,
          affectedComments,
          opType: 'deleteTableRow',
        },
      );
    }
    // drop: 连带清理
    for (const name of affectedBookmarks) {
      dropBookmarkByName(ctx, name);
    }
    for (const id of affectedComments) {
      dropCommentById(ctx, id, tr);
    }
  }

  // markDirty 必须在 removeChild 之前
  markStoryPartDirty(ctx, tr, 'deleteTableRow');

  const parent = tr.parentNode;
  if (!parent) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `row of table '${op.tableId}' has no parentNode`,
      ctx.docId,
      {
        tableId: op.tableId,
        row: op.row,
      },
    );
  }
  parent.removeChild(tr);
}

// ─────────── insertTableRow(#14)───────────

export function executeInsertTableRow(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'insertTableRow' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'insertTableRow with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'insertTableRow', tableId: op.tableId },
    );
  }

  // 1. 拿表格
  const tbl = requireTable(ctx, op.tableId, 'insertTableRow');

  // 2. 找模板行
  const templateTr = findTableRow(tbl, op.templateRow);
  if (!templateTr) {
    throw new WordError(
      'TABLE_NOT_FOUND',
      `template row ${op.templateRow} not found in table '${op.tableId}'`,
      ctx.docId,
      { tableId: op.tableId, templateRow: op.templateRow, opType: 'insertTableRow' },
    );
  }

  // 3. 统计当前行数 + 校验 position
  let rowCount = 0;
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const c = tbl.childNodes[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === 'tr'
    ) {
      rowCount++;
    }
  }

  let insertPos: number;
  if (op.position === 'end') {
    insertPos = rowCount;
  } else {
    if (op.position < 0 || op.position > rowCount) {
      throw new WordError(
        'TABLE_NOT_FOUND',
        `position ${op.position} out of range [0, ${rowCount}] in table '${op.tableId}'`,
        ctx.docId,
        { tableId: op.tableId, position: op.position, rowCount, opType: 'insertTableRow' },
      );
    }
    insertPos = op.position;
  }

  // 4. 克隆模板行
  const newTr = templateTr.cloneNode(true) as Element;

  // 5. 剥离克隆行内锚点
  stripAnchorsFromClone(newTr);

  // 6. 重分配克隆行内每个 <w:p> 的 paraId
  reassignClonedParaIds(ctx, newTr);

  // 7. 填 cellTexts（若传）
  if (op.cellTexts !== undefined) {
    fillCellTexts(newTr, op.cellTexts);
  }

  // 8. 插入 DOM（markDirty 必须在 newTr 进树之后）
  if (insertPos === rowCount) {
    tbl.appendChild(newTr);
  } else {
    // 找第 insertPos 个 <w:tr>
    let trIdx = 0;
    let refNode: Element | null = null;
    for (let i = 0; i < tbl.childNodes.length; i++) {
      const c = tbl.childNodes[i];
      if (
        c &&
        c.nodeType === 1 &&
        (c as Element).namespaceURI === W_NS &&
        (c as Element).localName === 'tr'
      ) {
        if (trIdx === insertPos) {
          refNode = c as Element;
          break;
        }
        trIdx++;
      }
    }
    tbl.insertBefore(newTr, refNode);
  }

  // 9. markDirty（节点已在树内）
  markStoryPartDirty(ctx, newTr, 'insertTableRow');
}

// ─────────── cloneTableAfter(#15)───────────

export function executeCloneTableAfter(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'cloneTableAfter' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'cloneTableAfter with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'cloneTableAfter', sourceTableId: op.sourceTableId },
    );
  }

  // 2. 查 anchor 段
  const anchor = ctx.indices.paraIdx.get(op.anchorParaId.toUpperCase());
  if (!anchor) {
    throw new WordError(
      'PARA_NOT_FOUND',
      `paragraph with paraId='${op.anchorParaId}' not found (op: cloneTableAfter)`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: 'cloneTableAfter' },
    );
  }

  // 3. 查源表
  const sourceTbl = requireTable(ctx, op.sourceTableId, 'cloneTableAfter');

  // 4. 深克隆
  const cloned = sourceTbl.cloneNode(true) as Element;

  // 5. 剥离克隆表内锚点
  stripAnchorsFromClone(cloned);

  // 6. 重分配克隆表内所有 <w:p> 的 paraId
  reassignClonedParaIds(ctx, cloned);

  // 7. 插入 DOM：紧贴 anchor 之后（insertBefore(cloned, anchor.nextSibling)）
  const parent = anchor.parentNode;
  if (!parent) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `anchor paragraph '${op.anchorParaId}' has no parentNode`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: 'cloneTableAfter' },
    );
  }
  parent.insertBefore(cloned, anchor.nextSibling);

  // 8. markDirty（节点已在树内）
  markStoryPartDirty(ctx, cloned, 'cloneTableAfter');
}

/**
 * 重分配克隆子树内所有 `<w:p>` 的 paraId（全局唯一）。
 *
 * 从当前 `paraIdx.keys()` 构建 seen 集合，遍历节点内所有 `<w:p>`，
 * 用 `generateParaId(seen)` 分配新 id（generateParaId 内部自动 add，防同棵树多段冲突）。
 * 仅限 table-ops.ts 内部使用，供 insertTableRow 和 cloneTableAfter 共享。
 */
function reassignClonedParaIds(ctx: ExecutorContext, node: Element): void {
  const seen = new Set(ctx.indices.paraIdx.keys());
  const paragraphs = node.getElementsByTagNameNS(W_NS, 'p');
  // 快照成数组避免 live collection 问题
  const paraSnapshot: Element[] = [];
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (p) paraSnapshot.push(p);
  }
  for (const p of paraSnapshot) {
    const newId = generateParaId(seen);
    p.setAttributeNS(W14_NS, 'w14:paraId', newId);
  }
}

/**
 * 剥离克隆行内所有锚点元素:
 * - bookmarkStart / bookmarkEnd
 * - commentRangeStart / commentRangeEnd
 * - commentReference（连同空 run 一并清）
 */
function stripAnchorsFromClone(tr: Element): void {
  const anchorLocals = ['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd'];
  for (const local of anchorLocals) {
    const list = tr.getElementsByTagNameNS(W_NS, local);
    // 快照避免 live collection 修改问题
    const toRemove: Element[] = [];
    for (let i = 0; i < list.length; i++) {
      const el = list[i];
      if (el) toRemove.push(el);
    }
    for (const el of toRemove) {
      if (el.parentNode) el.parentNode.removeChild(el);
    }
  }

  // commentReference：连同空 run 一并清
  const refs = tr.getElementsByTagNameNS(W_NS, 'commentReference');
  const refSnapshot: Element[] = [];
  for (let i = 0; i < refs.length; i++) {
    const el = refs[i];
    if (el) refSnapshot.push(el);
  }
  for (const ref of refSnapshot) {
    const parentRun = ref.parentNode;
    if (ref.parentNode) ref.parentNode.removeChild(ref);
    // 若 run 仅剩 rPr（或完全空），一并删除
    if (
      parentRun &&
      parentRun.nodeType === 1 &&
      (parentRun as Element).namespaceURI === W_NS &&
      (parentRun as Element).localName === 'r' &&
      isRunEffectivelyEmpty(parentRun as Element)
    ) {
      if (parentRun.parentNode) parentRun.parentNode.removeChild(parentRun);
    }
  }
}

/**
 * 按 §1.3 语义填充克隆行各列文字。
 * 设列数 C，cellTexts 长度 L：
 *   - i ∈ [0, min(L,C))：替换该列单元格首段为 cellTexts[i]（保留 tcPr/pPr/首 run rPr）
 *   - i ∈ [L, C)：清空该列（removeNonPPrChildren，不添加文字 run）
 *   - i ≥ C：忽略（cellTexts 超出列数的元素）
 */
function fillCellTexts(tr: Element, cellTexts: string[]): void {
  // 收集直属 <w:tc>
  const tcs: Element[] = [];
  for (let i = 0; i < tr.childNodes.length; i++) {
    const c = tr.childNodes[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === 'tc'
    ) {
      tcs.push(c as Element);
    }
  }

  const C = tcs.length;
  const L = cellTexts.length;
  const fillCount = Math.min(L, C);

  for (let i = 0; i < C; i++) {
    const tc = tcs[i]!;
    const firstPara = firstChildW(tc, 'p');
    if (!firstPara) continue;

    if (i < fillCount) {
      // 替换首段：保留 pPr + 首 run rPr，写入文字
      const inheritedRPr = pickFirstRunRPr(firstPara);
      const doc = firstPara.ownerDocument;
      if (!doc) continue;

      const text = cellTexts[i]!;
      const run = doc.createElementNS(W_NS, 'w:r');
      if (inheritedRPr) {
        run.appendChild(inheritedRPr);
      }
      const t = doc.createElementNS(W_NS, 'w:t');
      t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      t.textContent = text;
      run.appendChild(t);

      removeNonPPrChildren(firstPara);
      firstPara.appendChild(run);
    } else {
      // 清空（i ∈ [L, C)）：仅保留 pPr，不加文字
      removeNonPPrChildren(firstPara);
    }
  }
}

// ─────────── setTableColumnWidths(task-core-table-grid)───────────

/**
 * 设表格列宽：同步更新 `<w:tblGrid>` gridCol w:w **和** 每个单元格 `<w:tcW>`。
 *
 * Round 2 修正（lead 实测）：真实文档单元格带 `<w:tcW w:type="dxa">`，Word 渲染
 * 优先走 tcW 而非 gridCol，因此只设 gridCol 不够——必须同步 tcW，列宽才生效。
 *
 * 定位逻辑：
 *   1. 按 tableId 找 `<w:tbl>`（tblIdx）。
 *   2. 找 `<w:tblGrid>`（tbl 直属子）；若不存在则在 tblPr 之后（或 tbl 首位）创建。
 *   3. 枚举 `<w:gridCol>` 子元素，数量必须与 widths.length 一致（INVALID_ARGUMENT）。
 *   4. 依次把每个 gridCol 的 `w:w` 设为对应 widths 值。
 *   5. 遍历每行（`<w:tr>`）→ 每格（`<w:tc>`），跟踪**网格列索引** col：
 *      - 读该 cell 的 `<w:gridSpan w:val>`（默认 1），span 跨几列。
 *      - 该 cell 的 tcW = widths[col] + widths[col+1] + … + widths[col+span-1]（跨列宽之和）。
 *      - 若已有 `<w:tcW>` → 更新 w:w（type 改 dxa）；若无 tcW → 补建（更稳，避免
 *        某些渲染器只读 tcW 不读 gridCol 的场景）。
 *      - col 前进 span；`<w:vMerge>` 续行单元格照常计入列索引。
 *
 * 边界说明：
 *   - 同步 tcW（含 gridSpan 跨列求和），与 gridCol 保持一致；tcW type 统一设 dxa。
 *   - trackChanges=true 暂不支持（同其他 table ops 行为）。
 */
export function executeSetTableColumnWidths(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'setTableColumnWidths' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'setTableColumnWidths with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'setTableColumnWidths', tableId: op.tableId },
    );
  }

  const tbl = requireTable(ctx, op.tableId, 'setTableColumnWidths');

  // ── Step 1: 定位或创建 tblGrid ────────────────────────────────────────────
  let tblGrid = firstChildW(tbl, 'tblGrid');
  if (!tblGrid) {
    const doc = tbl.ownerDocument;
    if (!doc) {
      throw new WordError(
        'PACKAGE_CORRUPT',
        `table '${op.tableId}' has no ownerDocument`,
        ctx.docId,
        { tableId: op.tableId, opType: 'setTableColumnWidths' },
      );
    }
    tblGrid = doc.createElementNS(W_NS, 'w:tblGrid');
    // 插入到 tblPr 之后（或 tbl 首子节点之前）
    const tblPr = firstChildW(tbl, 'tblPr');
    if (tblPr && tblPr.nextSibling) {
      tbl.insertBefore(tblGrid, tblPr.nextSibling);
    } else if (tblPr) {
      tbl.appendChild(tblGrid);
    } else {
      tbl.insertBefore(tblGrid, tbl.firstChild);
    }
  }

  // ── Step 2: 收集 gridCol 元素 ────────────────────────────────────────────
  const gridCols: Element[] = [];
  for (let i = 0; i < tblGrid.childNodes.length; i++) {
    const c = tblGrid.childNodes[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === 'gridCol'
    ) {
      gridCols.push(c as Element);
    }
  }

  // ── Step 3: 校验长度 ─────────────────────────────────────────────────────
  if (gridCols.length !== op.widths.length) {
    throw new WordError(
      'INVALID_ARGUMENT',
      `setTableColumnWidths: widths.length (${op.widths.length}) must equal colCount (${gridCols.length}) for table '${op.tableId}'`,
      ctx.docId,
      {
        tableId: op.tableId,
        opType: 'setTableColumnWidths',
        widthsLength: op.widths.length,
        colCount: gridCols.length,
      },
    );
  }

  // ── Step 4: 设置 gridCol w:w ─────────────────────────────────────────────
  for (let i = 0; i < gridCols.length; i++) {
    gridCols[i]!.setAttributeNS(W_NS, 'w:w', String(op.widths[i]));
  }

  // ── Step 5: 同步每个单元格 tcW（Round 2：Word 渲染优先 tcW over gridCol）─
  // 遍历所有 <w:tr>，在每行内按网格列索引逐 <w:tc> 处理。
  // gridSpan > 1 的合并单元格 tcW = 跨列宽之和；vMerge 续行单元格照常计入列位置。
  const doc = tbl.ownerDocument;
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const trNode = tbl.childNodes[i];
    if (
      !trNode ||
      trNode.nodeType !== 1 ||
      (trNode as Element).namespaceURI !== W_NS ||
      (trNode as Element).localName !== 'tr'
    )
      continue;
    const tr = trNode as Element;

    let col = 0; // 当前网格列索引

    for (let j = 0; j < tr.childNodes.length; j++) {
      const tcNode = tr.childNodes[j];
      if (
        !tcNode ||
        tcNode.nodeType !== 1 ||
        (tcNode as Element).namespaceURI !== W_NS ||
        (tcNode as Element).localName !== 'tc'
      )
        continue;
      const tc = tcNode as Element;

      // 读取 gridSpan（默认 1）
      const tcPr = firstChildW(tc, 'tcPr');
      let span = 1;
      if (tcPr) {
        const gridSpanEl = firstChildW(tcPr, 'gridSpan');
        if (gridSpanEl) {
          const val = gridSpanEl.getAttributeNS(W_NS, 'val');
          const parsed = val !== null ? parseInt(val, 10) : NaN;
          if (!isNaN(parsed) && parsed > 1) span = parsed;
        }
      }

      // tcW = 跨列宽之和（span=1 退化为 widths[col]）
      let tcWidth = 0;
      for (let s = 0; s < span && col + s < op.widths.length; s++) {
        tcWidth += op.widths[col + s]!;
      }

      // 更新或补建 <w:tcW>
      if (tcPr) {
        const tcWEl = firstChildW(tcPr, 'tcW');
        if (tcWEl) {
          // 已存在：直接更新 w:w 和 w:type
          tcWEl.setAttributeNS(W_NS, 'w:w', String(tcWidth));
          tcWEl.setAttributeNS(W_NS, 'w:type', 'dxa');
        } else if (doc) {
          // 无 tcW：补建，插入到 tcPr 内适当位置（gridSpan 之后或 tcPr 末尾）
          const newTcW = doc.createElementNS(W_NS, 'w:tcW');
          newTcW.setAttributeNS(W_NS, 'w:w', String(tcWidth));
          newTcW.setAttributeNS(W_NS, 'w:type', 'dxa');
          // 插到 tcPr 首位（tcW 一般是 tcPr 第一个子元素）
          tcPr.insertBefore(newTcW, tcPr.firstChild);
        }
      }
      // 若无 tcPr：不补建（表示该单元格完全继承默认属性，gridCol 已足够）

      col += span;
    }
  }

  markStoryPartDirty(ctx, tbl, 'setTableColumnWidths');
}

// ─────────── setTableRowHeights(task-core-table-rowheight)───────────

/**
 * 设表格行高：逐行写/替换/删 `<w:trHeight>`。
 *
 * 实现镜像 executeSetTableColumnWidths：
 *   1. requireTable 拿表。
 *   2. 取表直属 <w:tr> 列表。
 *   3. 长度校验：heights.length 必须 == rowCount，否则 INVALID_ARGUMENT。
 *   4. 逐行：
 *      - 确保 <w:trPr> 存在（无则在 <w:tr> 内首位新建）。
 *      - RowHeight：找 trPr 内 <w:trHeight>，有则改 w:val/w:hRule，无则新建插入。
 *      - null：删除 trPr 内 <w:trHeight>（若存在）。
 *   5. trackChanges=true → not yet implemented。
 *
 * CT_TrPr 顺序：trHeight 在
 * cnfStyle/divId/gridBefore/gridAfter/wBefore/wAfter/cantSplit/[trHeight]/
 * tblHeader/tblCellSpacing/jc/hidden 位置。
 * 实现上：若 trPr 内有 cantSplit，把 trHeight 插到 cantSplit 之后；否则追加到末尾。
 * Word 对 trPr 子元素顺序较宽容，此策略安全。
 */
export function executeSetTableRowHeights(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'setTableRowHeights' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'setTableRowHeights with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'setTableRowHeights', tableId: op.tableId },
    );
  }

  const tbl = requireTable(ctx, op.tableId, 'setTableRowHeights');

  // ── 收集直属 <w:tr> ──────────────────────────────────────────────────────
  const rows: Element[] = [];
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const c = tbl.childNodes[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === 'tr'
    ) {
      rows.push(c as Element);
    }
  }

  // ── 长度校验 ─────────────────────────────────────────────────────────────
  if (op.heights.length !== rows.length) {
    throw new WordError(
      'INVALID_ARGUMENT',
      `setTableRowHeights: heights.length (${op.heights.length}) must equal rowCount (${rows.length}) for table '${op.tableId}'`,
      ctx.docId,
      {
        tableId: op.tableId,
        opType: 'setTableRowHeights',
        heightsLength: op.heights.length,
        rowCount: rows.length,
      },
    );
  }

  const doc = tbl.ownerDocument;

  for (let i = 0; i < rows.length; i++) {
    const tr = rows[i]!;
    const height = op.heights[i]!;

    // ── 确保 <w:trPr> 存在 ───────────────────────────────────────────────
    let trPr = firstChildW(tr, 'trPr');
    if (!trPr) {
      if (!doc) {
        throw new WordError(
          'PACKAGE_CORRUPT',
          `table '${op.tableId}' row ${i} has no ownerDocument`,
          ctx.docId,
          { tableId: op.tableId, opType: 'setTableRowHeights' },
        );
      }
      trPr = doc.createElementNS(W_NS, 'w:trPr');
      // trPr 是 tr 的首个子元素
      tr.insertBefore(trPr, tr.firstChild);
    }

    if (height === null) {
      // ── null：删除 trPr 内 <w:trHeight>（若存在）──────────────────────
      const existingTrHeight = firstChildW(trPr, 'trHeight');
      if (existingTrHeight && existingTrHeight.parentNode) {
        existingTrHeight.parentNode.removeChild(existingTrHeight);
      }
    } else {
      // ── RowHeight：写/替换 <w:trHeight> ──────────────────────────────
      const existingTrHeight = firstChildW(trPr, 'trHeight');
      if (existingTrHeight) {
        // 已有 trHeight：更新属性
        existingTrHeight.setAttributeNS(W_NS, 'w:val', String(height.val));
        existingTrHeight.setAttributeNS(W_NS, 'w:hRule', height.hRule);
      } else {
        // 无 trHeight：新建并插入
        if (!doc) {
          throw new WordError(
            'PACKAGE_CORRUPT',
            `table '${op.tableId}' row ${i} has no ownerDocument`,
            ctx.docId,
            { tableId: op.tableId, opType: 'setTableRowHeights' },
          );
        }
        const newTrHeight = doc.createElementNS(W_NS, 'w:trHeight');
        newTrHeight.setAttributeNS(W_NS, 'w:val', String(height.val));
        newTrHeight.setAttributeNS(W_NS, 'w:hRule', height.hRule);
        // 按 CT_TrPr 顺序：trHeight 在 cantSplit 之后；若无 cantSplit 则追加到末尾。
        const cantSplit = firstChildW(trPr, 'cantSplit');
        if (cantSplit && cantSplit.nextSibling) {
          trPr.insertBefore(newTrHeight, cantSplit.nextSibling);
        } else if (cantSplit) {
          trPr.appendChild(newTrHeight);
        } else {
          trPr.appendChild(newTrHeight);
        }
      }
    }
  }

  markStoryPartDirty(ctx, tbl, 'setTableRowHeights');
}

// ─────────── 内部 ───────────

const TCW_AFTER_LOCALS = new Set([
  'gridSpan',
  'hMerge',
  'vMerge',
  'tcBorders',
  'shd',
  'noWrap',
  'tcMar',
  'textDirection',
  'tcFitText',
  'vAlign',
  'hideMark',
  'headers',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
]);

const SHD_AFTER_LOCALS = new Set([
  'noWrap',
  'tcMar',
  'textDirection',
  'tcFitText',
  'vAlign',
  'hideMark',
  'headers',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
]);

const VALIGN_AFTER_LOCALS = new Set([
  'hideMark',
  'headers',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
]);

function requireTable(ctx: ExecutorContext, tableId: string, opName: string): Element {
  const tbl = ctx.indices.tblIdx.get(tableId);
  if (!tbl) {
    throw new WordError(
      'TABLE_NOT_FOUND',
      `table with tableId='${tableId}' not found (op: ${opName})`,
      ctx.docId,
      { tableId, opName },
    );
  }
  return tbl;
}

/**
 * 找表格的首段落:第一行 → 第一格 → 第一段。
 * OOXML 保证每个 <w:tc> 至少含一个 <w:p>。
 */
function findFirstParagraphInTable(tbl: Element): Element | null {
  const firstTr = firstChildW(tbl, 'tr');
  if (!firstTr) return null;
  const firstTc = firstChildW(firstTr, 'tc');
  if (!firstTc) return null;
  return firstChildW(firstTc, 'p');
}

function firstChildW(el: Element, localName: string): Element | null {
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === localName
    ) {
      return c as Element;
    }
  }
  return null;
}

function firstCellContentChild(tc: Element): Element | null {
  for (let i = 0; i < tc.childNodes.length; i++) {
    const child = tc.childNodes[i];
    if (
      child &&
      child.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName !== 'tcPr'
    ) {
      return child as Element;
    }
  }
  return null;
}

/**
 * 按 row 索引定位 `<w:tr>`。row 是 tbl 下第 i 个 `<w:tr>` 直属子 (按文档顺序)。
 */
function findTableRow(tbl: Element, row: number): Element | null {
  if (row < 0) return null;
  let trIdx = 0;
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const c = tbl.childNodes[i];
    if (
      !c ||
      c.nodeType !== 1 ||
      (c as Element).namespaceURI !== W_NS ||
      (c as Element).localName !== 'tr'
    )
      continue;
    if (trIdx === row) return c as Element;
    trIdx++;
  }
  return null;
}

/**
 * 按 (row, col) 网格坐标定位 `<w:tc>`。坐标语义与 TableInfo.cells 一致:
 *   - row = tbl 下第 i 个 `<w:tr>` 直属子 (按文档顺序)
 *   - col = 该 tr 下第 j 个 `<w:tc>` 直属子
 *
 * 合并单元格 (gridSpan / vMerge='continue') 不在坐标系内做展开 —— TableInfo.cells
 * 也是按 tc 顺序填的, 调用方拿到的网格与 executor 解析的网格 1:1 对应。
 */
function findCell(tbl: Element, row: number, col: number): Element | null {
  if (row < 0 || col < 0) return null;
  let trIdx = 0;
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const c = tbl.childNodes[i];
    if (
      !c ||
      c.nodeType !== 1 ||
      (c as Element).namespaceURI !== W_NS ||
      (c as Element).localName !== 'tr'
    )
      continue;
    if (trIdx === row) {
      let tcIdx = 0;
      const tr = c as Element;
      for (let j = 0; j < tr.childNodes.length; j++) {
        const cc = tr.childNodes[j];
        if (
          !cc ||
          cc.nodeType !== 1 ||
          (cc as Element).namespaceURI !== W_NS ||
          (cc as Element).localName !== 'tc'
        )
          continue;
        if (tcIdx === col) return cc as Element;
        tcIdx++;
      }
      return null;
    }
    trIdx++;
  }
  return null;
}

/** 段落内除 `<w:pPr>` 外所有直属子节点移除 (保留段落属性, 清空内容)。 */
function ensureTableCellProperties(tc: Element): Element {
  const existing = firstChildW(tc, 'tcPr');
  if (existing) return existing;
  const doc = tc.ownerDocument!;
  const tcPr = doc.createElementNS(W_NS, 'w:tcPr');
  tc.insertBefore(tcPr, tc.firstChild);
  return tcPr;
}

function upsertTcPrChild(tcPr: Element, localName: string, afterLocals: Set<string>): Element {
  const existing = firstChildW(tcPr, localName);
  if (existing) return existing;
  const doc = tcPr.ownerDocument!;
  const child = doc.createElementNS(W_NS, `w:${localName}`);
  tcPr.insertBefore(child, findTcPrInsertRef(tcPr, afterLocals));
  return child;
}

function findTcPrInsertRef(tcPr: Element, afterLocals: Set<string>): Element | null {
  for (let i = 0; i < tcPr.childNodes.length; i++) {
    const child = tcPr.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      afterLocals.has((child as Element).localName)
    ) {
      return child as Element;
    }
  }
  return null;
}

function validateTableCellFormat(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'formatTableCell' }>,
): void {
  if (
    op.format.width === undefined &&
    op.format.shading === undefined &&
    op.format.verticalAlignment === undefined
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatTableCell requires at least one format field',
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, opType: 'formatTableCell' },
    );
  }
  if (
    op.format.width !== undefined &&
    (!Number.isInteger(op.format.width) || op.format.width <= 0)
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatTableCell width must be a positive integer twip value',
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, width: op.format.width },
    );
  }
  if (
    op.format.verticalAlignment !== undefined &&
    !['top', 'center', 'bottom', 'both'].includes(op.format.verticalAlignment)
  ) {
    throw new WordError(
      'INVALID_ARGUMENT',
      'formatTableCell verticalAlignment must be top, center, bottom, or both',
      ctx.docId,
      {
        tableId: op.tableId,
        row: op.row,
        col: op.col,
        verticalAlignment: op.format.verticalAlignment,
      },
    );
  }
  if (op.format.shading !== undefined) {
    validateColorLike(ctx, op, 'shading.fill', op.format.shading.fill);
    validateColorLike(ctx, op, 'shading.color', op.format.shading.color);
    if (
      op.format.shading.fill === undefined &&
      op.format.shading.color === undefined &&
      op.format.shading.val === undefined
    ) {
      throw new WordError(
        'INVALID_ARGUMENT',
        'formatTableCell shading requires fill, color, or val',
        ctx.docId,
        { tableId: op.tableId, row: op.row, col: op.col, opType: 'formatTableCell' },
      );
    }
  }
}

function validateColorLike(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'formatTableCell' }>,
  field: string,
  value: string | undefined,
): void {
  if (value === undefined) return;
  if (value !== 'auto' && !/^[0-9A-Fa-f]{6}$/.test(value)) {
    throw new WordError(
      'INVALID_ARGUMENT',
      `formatTableCell ${field} must be a 6-digit hex color or auto`,
      ctx.docId,
      { tableId: op.tableId, row: op.row, col: op.col, field, value },
    );
  }
}

function removeNonPPrChildren(paragraph: Element): void {
  const children = Array.from(paragraph.childNodes);
  for (const c of children) {
    if (c.nodeType !== 1) {
      paragraph.removeChild(c);
      continue;
    }
    const el = c as Element;
    if (el.namespaceURI === W_NS && el.localName === 'pPr') continue;
    paragraph.removeChild(c);
  }
}

function collectBookmarkNamesInTable(tbl: Element): string[] {
  const names: string[] = [];
  const starts = tbl.getElementsByTagNameNS(W_NS, 'bookmarkStart');
  for (let i = 0; i < starts.length; i++) {
    const el = starts[i];
    if (!el) continue;
    const name = el.getAttributeNS(W_NS, 'name') ?? el.getAttribute('w:name') ?? '';
    names.push(name);
  }
  return names;
}

function collectCommentIdsInTable(tbl: Element): number[] {
  const ids = new Set<number>();
  for (const local of ['commentRangeStart', 'commentReference']) {
    const list = tbl.getElementsByTagNameNS(W_NS, local);
    for (let i = 0; i < list.length; i++) {
      const el = list[i];
      if (!el) continue;
      const raw = el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '';
      const id = parseInt(raw, 10);
      if (!isNaN(id)) ids.add(id);
    }
  }
  return Array.from(ids);
}

/**
 * drop 书签:移除 bookmarkStart / bookmarkEnd,删 bookmarkIdx 条目。
 * 与 paragraph-structure-ops.ts dropBookmarkByName 同构,但 #186 表内书签是要
 * 整张表删的,所以单独走一份避免跨文件 export。
 */
function dropBookmarkByName(ctx: ExecutorContext, name: string): void {
  const entry = ctx.indices.bookmarkIdx.get(name);
  if (!entry) return;
  const { start, end, partPath } = entry;
  if (start.parentNode) start.parentNode.removeChild(start);
  if (end && end.parentNode) end.parentNode.removeChild(end);
  ctx.indices.bookmarkIdx.delete(name);
  const part = getPart(ctx.pkg, partPath);
  if (part) markDirty(part);
}

/**
 * drop 批注:删 commentRangeStart/End + commentReference 三锚点 + comments.xml
 * 里的 <w:comment> 主体。
 *
 * 注:整张表都将被 removeChild,严格说只需要删 comments.xml 里的主体即可;但保留
 * 锚点清理逻辑作为兜底(若某 op 在 deleteTable 之前 / 之后访问 commentIdx 可保
 * 一致)。
 */
function dropCommentById(ctx: ExecutorContext, id: number, tbl: Element): void {
  const entry = ctx.indices.commentIdx.get(id);

  if (!entry) {
    // 索引未命中:直接扫表删锚点
    removeCommentAnchorsInTable(ctx, tbl, id);
    return;
  }

  const { start, end, ref } = entry.anchors;
  const anchorEl = start ?? end ?? ref;
  let bodyPartPath: string | null = null;
  if (anchorEl) {
    const story = findStoryFor(ctx.pkg, anchorEl);
    if (story) bodyPartPath = story.partPath;
  }

  if (start && start.parentNode) start.parentNode.removeChild(start);
  if (end && end.parentNode) end.parentNode.removeChild(end);
  if (ref) {
    const parentRun = ref.parentNode;
    if (ref.parentNode) ref.parentNode.removeChild(ref);
    if (
      parentRun &&
      parentRun.nodeType === 1 &&
      (parentRun as Element).namespaceURI === W_NS &&
      (parentRun as Element).localName === 'r' &&
      isRunEffectivelyEmpty(parentRun as Element)
    ) {
      if (parentRun.parentNode) parentRun.parentNode.removeChild(parentRun);
    }
  }

  if (entry.comment && entry.comment.parentNode) {
    entry.comment.parentNode.removeChild(entry.comment);
    const commentsPart = getPart(ctx.pkg, COMMENTS_PART_PATH);
    if (commentsPart) markDirty(commentsPart);
  }

  if (bodyPartPath) {
    const bodyPart = getPart(ctx.pkg, bodyPartPath);
    if (bodyPart) markDirty(bodyPart);
  }

  ctx.indices.commentIdx.delete(id);

  // 若 comments.xml 已空，移除 part + 清 CT/rels（与 executeDeleteComment 对称）
  cleanupEmptyCommentsPart(ctx);
}

function removeCommentAnchorsInTable(ctx: ExecutorContext, tbl: Element, id: number): void {
  const idStr = String(id);
  const toRemove: Element[] = [];
  for (const local of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
    const list = tbl.getElementsByTagNameNS(W_NS, local);
    for (let i = 0; i < list.length; i++) {
      const el = list[i];
      if (!el) continue;
      const raw = el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '';
      if (raw === idStr) toRemove.push(el);
    }
  }
  for (const el of toRemove) {
    const parentRun = el.localName === 'commentReference' ? el.parentNode : null;
    if (el.parentNode) el.parentNode.removeChild(el);
    if (
      parentRun &&
      parentRun.nodeType === 1 &&
      (parentRun as Element).namespaceURI === W_NS &&
      (parentRun as Element).localName === 'r' &&
      isRunEffectivelyEmpty(parentRun as Element)
    ) {
      if (parentRun.parentNode) parentRun.parentNode.removeChild(parentRun);
    }
  }

  for (const part of listParts(ctx.pkg)) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    if (!part.contentType.endsWith('.comments+xml')) continue;
    const comments = part.xmlDom.getElementsByTagNameNS(W_NS, 'comment');
    const victims: Element[] = [];
    for (let i = 0; i < comments.length; i++) {
      const c = comments[i];
      if (!c) continue;
      const raw = c.getAttributeNS(W_NS, 'id') ?? c.getAttribute('w:id') ?? '';
      if (raw === idStr) victims.push(c);
    }
    for (const v of victims) {
      if (v.parentNode) v.parentNode.removeChild(v);
    }
    if (victims.length > 0) markDirty(part);
  }
}

function isRunEffectivelyEmpty(run: Element): boolean {
  const children = run.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (!c || c.nodeType !== 1) continue;
    if ((c as Element).localName === 'rPr') continue;
    return false;
  }
  return true;
}

function markStoryPartDirty(ctx: ExecutorContext, el: Element, opName: string): void {
  const story = findStoryFor(ctx.pkg, el);
  if (!story) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `element not found in any story (op: ${opName})`,
      ctx.docId,
      { opName },
    );
  }
  const part = getPart(ctx.pkg, story.partPath);
  if (!part) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `story part '${story.partPath}' not found (op: ${opName})`,
      ctx.docId,
      { opName, partPath: story.partPath },
    );
  }
  markDirty(part);
}
