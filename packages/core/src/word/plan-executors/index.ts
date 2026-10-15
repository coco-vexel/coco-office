/**
 * Plan executor dispatcher(#28 起)。
 *
 * WordPlan 提交的 op 队列走 `kernel.applyPlan` → `dispatchOp` → 具体 executor。
 * 每个 executor 一个文件:
 *   - paragraph-ops.ts       #28 replaceParagraph / replaceTextInParagraph
 *   - paragraph-structure-ops.ts  #29 insertParagraphBefore/After / deleteParagraph
 *   - bookmark-ops.ts        #34 createBookmark / deleteBookmark
 *   - comment-ops.ts         #35 addComment / deleteComment
 *   - revision-ops.ts        #31 trackChanges 包装 / #32 acceptRevisions / rejectRevisions
 *
 * **Executor 契约**(给后续 PR 照抄):
 *   - 签名:`function executeXxx(ctx: ExecutorContext, op): void`
 *   - 报错:throw `WordError`(对应 code:PARA_NOT_FOUND / TEXT_NOT_FOUND_IN_PARA /
 *     ANCHOR_CONFLICT / UNSUPPORTED_OOXML_FEATURE / BOOKMARK_NAME_EXISTS 等);
 *     docId 由外层 kernel.applyPlan 捕后重抛注入
 *   - 产出:mutate `ctx.result` 累积 ApplyResult 字段(insertedParagraphs 等);
 *     不用 return 值
 *   - 事务:外层已把所有 affected parts 包进 runTransaction;executor 只关心
 *     当前 op 的变更,抛错自动回滚
 *   - markDirty:executor 负责在改完 part 后 `markDirty(pkg, partPath)`
 *   - 索引同步:executor 自行维护 indices(paraIdx / bookmarkIdx / commentIdx)
 *     对简单 op(只改 `<w:p>` 内部),索引条目指向的 `<w:p>` Element 不变,
 *     不需动;涉及段落增删或书签/批注增删时必须同步
 */
import type { Op } from '../plan.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';
import {
  executeFormatParagraph,
  executeFormatTextInParagraph,
  executeReplaceParagraph,
  executeReplaceTextInParagraph,
} from './paragraph-ops.js';
import { executeCreateBookmark, executeDeleteBookmark } from './bookmark-ops.js';
import { executeAddComment, executeDeleteComment } from './comment-ops.js';
import {
  executeInsertParagraph,
  executeInsertParagraphAtBodyStart,
  executeDeleteParagraph,
} from './paragraph-structure-ops.js';
import { executeAcceptRevisions, executeRejectRevisions } from './revision-ops.js';
import {
  executeAddTableComment,
  executeDeleteTable,
  executeReplaceTableCellWithSdt,
  executeReplaceTableCellText,
  executeFormatTableCell,
  executeInsertTableCellParagraph,
  executeDeleteTableRow,
  executeInsertTableRow,
  executeCloneTableAfter,
  executeSetTableColumnWidths,
  executeSetTableRowHeights,
} from './table-ops.js';
import {
  executeReplaceDrawingWithPlaceholder,
  executeDeleteDrawing,
  executeReplaceShapeInnerText,
} from './drawing-ops.js';
import { executeInsertImageAfter } from './insert-image-ops.js';
import { executeInsertTableFrom } from './insert-table-ops.js';
import { executeSetUpdateFields } from './settings-ops.js';

export type { ExecutorContext } from './executor-context.js';

/**
 * 分发一个 op 到对应 executor。未实现的 op 类型抛 `UNSUPPORTED_OOXML_FEATURE`,
 * message 指向对应 issue。
 */
export function dispatchOp(op: Op, ctx: ExecutorContext): void {
  switch (op.type) {
    case 'replaceParagraph':
      executeReplaceParagraph(ctx, op);
      return;
    case 'replaceTextInParagraph':
      executeReplaceTextInParagraph(ctx, op);
      return;
    case 'formatTextInParagraph':
      executeFormatTextInParagraph(ctx, op);
      return;
    case 'formatParagraph':
      executeFormatParagraph(ctx, op);
      return;
    // #29
    case 'insertParagraphBefore':
      executeInsertParagraph(ctx, op, 'before');
      return;
    case 'insertParagraphAfter':
      executeInsertParagraph(ctx, op, 'after');
      return;
    case 'insertParagraphAtBodyStart':
      executeInsertParagraphAtBodyStart(ctx, op);
      return;
    case 'deleteParagraph':
      executeDeleteParagraph(ctx, op);
      return;
    // #34
    case 'createBookmark':
      executeCreateBookmark(ctx, op);
      return;
    case 'deleteBookmark':
      executeDeleteBookmark(ctx, op);
      return;
    // #35
    case 'addComment':
      executeAddComment(ctx, op);
      return;
    case 'deleteComment':
      executeDeleteComment(ctx, op);
      return;
    // #32
    case 'acceptRevisions':
      executeAcceptRevisions(ctx, op);
      return;
    case 'rejectRevisions':
      executeRejectRevisions(ctx, op);
      return;
    // #186
    case 'addTableComment':
      executeAddTableComment(ctx, op);
      return;
    case 'deleteTable':
      executeDeleteTable(ctx, op);
      return;
    // #188
    case 'replaceTableCellWithSdt':
      executeReplaceTableCellWithSdt(ctx, op);
      return;
    case 'replaceTableCellText':
      executeReplaceTableCellText(ctx, op);
      return;
    case 'formatTableCell':
      executeFormatTableCell(ctx, op);
      return;
    case 'insertTableCellParagraph':
      executeInsertTableCellParagraph(ctx, op);
      return;
    case 'deleteTableRow':
      executeDeleteTableRow(ctx, op);
      return;
    // #14
    case 'insertTableRow':
      executeInsertTableRow(ctx, op);
      return;
    // #15
    case 'cloneTableAfter':
      executeCloneTableAfter(ctx, op);
      return;
    // task-core-table-grid
    case 'setTableColumnWidths':
      executeSetTableColumnWidths(ctx, op);
      return;
    // task-core-table-rowheight
    case 'setTableRowHeights':
      executeSetTableRowHeights(ctx, op);
      return;
    // #190
    case 'replaceDrawingWithPlaceholder':
      executeReplaceDrawingWithPlaceholder(ctx, op);
      return;
    case 'deleteDrawing':
      executeDeleteDrawing(ctx, op);
      return;
    // #210
    case 'replaceShapeInnerText':
      executeReplaceShapeInnerText(ctx, op);
      return;
    // task-insert-image
    case 'insertImageAfter':
      executeInsertImageAfter(ctx, op);
      return;
    // task-20260616-table-content-driven
    case 'insertTableFrom':
      executeInsertTableFrom(ctx, op);
      return;
    // task-core-update-fields
    case 'setUpdateFields':
      executeSetUpdateFields(ctx, op);
      return;
    default: {
      // 类型穷尽兜底
      const _never: never = op;
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        `unknown plan op type: ${JSON.stringify(_never)}`,
        ctx.docId,
      );
    }
  }
}

// notImplemented helper removed in #32 —— 所有 11 个 op 都已实现,无 stub 分支
