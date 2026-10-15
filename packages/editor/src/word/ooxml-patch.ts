import { createBinaryPart, createXmlPart, getPart, markDirty, setPart, type Package, type Part } from '@coco-vexel/coco-office/pack';
import { buildIndices, W14_NS, W_NS, type PackageIndices } from '@coco-vexel/coco-office/word-view';
import type { NativeWordChange } from './native-commands.js';
import type { NativeWordParagraphNode, NativeWordTableCellNode, NativeWordTableNode } from './native-model.js';
import type { EditorRunFormat } from './model.js';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const COMMENTS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_EXTENDED_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const COMMENTS_EXTENDED_REL_TYPE = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const HYPERLINK_REL_TYPE = `${R_NS}/hyperlink`;
const RELS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';

export type WordOoxmlPatchKind =
  | 'paragraph-text'
  | 'paragraph-split'
  | 'paragraph-merge'
  | 'paragraph-format'
  | 'format-revision'
  | 'drawing-delete'
  | 'hyperlink'
  | 'table-structure'
  | 'table-insert'
  | 'table-cell-format'
  | 'table-cell-clear'
  | 'table-cell-replace'
  | 'table-cell-replace-rich'
  | 'revision'
  | 'drawing-revision'
  | 'table-delete';

export type WordOoxmlPatch =
  | {
      kind: 'paragraph-text';
      op: 'insert' | 'delete' | 'replace';
      paraId: string;
      offset: number;
      text: string;
      deletedText?: string;
    }
  | {
      kind: 'paragraph-split';
      paraId: string;
      newParaId: string;
      offset: number;
    }
  | {
      kind: 'paragraph-merge';
      paraId: string;
    }
  | {
      kind: 'paragraph-format';
      paraId: string;
      from: number;
      to: number;
      format: Partial<EditorRunFormat>;
    }
  | {
      kind: 'format-revision';
      op: 'run' | 'accept' | 'reject';
      paraId: string;
      from: number;
      to: number;
      revisionId: string;
      beforeFormat?: Partial<EditorRunFormat> | null;
    }
  | {
      kind: 'drawing-delete';
      paraId: string;
      drawingId: string;
      offset: number;
    }
  | {
      kind: 'drawing-resize';
      paraId: string;
      drawingId: string;
      offset: number;
      width: number;
      height: number;
    }
  | {
      kind: 'drawing-insert';
      paraId: string;
      drawingId: string;
      offset: number;
      src: string;
      width: number;
      height: number;
    }
  | {
      kind: 'drawing-revision';
      op: 'accept' | 'reject';
      paraId: string;
      drawingId: string;
      offset: number;
      revisionId: string;
      revisionKind: 'insert' | 'delete';
    }
  | {
      kind: 'comment';
      op: 'insert' | 'delete' | 'update';
      paraId: string;
      from: number;
      to: number;
      commentId: number;
      body?: string;
      parentCommentId?: number;
    }
  | {
      kind: 'revision';
      op: 'insert' | 'delete' | 'accept' | 'reject';
      paraId: string;
      from: number;
      to: number;
      revisionId: string;
      revisionKind: 'insert' | 'delete';
    }
  | {
      kind: 'hyperlink';
      op: 'set' | 'remove';
      paraId: string;
      from: number;
      to: number;
      hyperlink?: { target?: string; anchor?: string; tooltip?: string };
    }
  | {
      kind: 'table-structure';
      op: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column' | 'merge-cells' | 'split-cell';
      tableId: string;
      rowIndex: number;
      columnIndex: number;
      endColumnIndex?: number;
      endRowIndex?: number;
    }
  | {
      kind: 'table-cell-replace-rich';
      tableId: string;
      rowIndex: number;
      columnIndex: number;
      cell: NativeWordTableCellNode;
    }
  | {
      kind: 'table-cell-replace';
      tableId: string;
      rowIndex: number;
      columnIndex: number;
      text: string;
    }
  | {
      kind: 'table-cell-clear';
      tableId: string;
      rowIndex: number;
      columnIndex: number;
    }
  | {
      kind: 'table-cell-format';
      tableId: string;
      rowIndex: number;
      columnIndex: number;
      format: {
        backgroundColor?: string | null;
        verticalAlignment?: 'top' | 'center' | 'bottom' | 'both' | null;
        borderColor?: string | null;
      };
    }
  | {
      kind: 'table-insert';
      anchorParaId: string;
      anchorOffset: number;
      table: NativeWordTableNode;
    }
  | {
      kind: 'table-delete';
      tableId: string;
    };

export interface WordOoxmlPatchPlan {
  target: 'word/document.xml';
  patches: WordOoxmlPatch[];
  unsupported: UnsupportedNativeWordChange[];
}

export type ResolvedWordOoxmlPatchTarget =
  | {
      kind: 'paragraph';
      patch: Extract<WordOoxmlPatch, { kind: 'paragraph-text' | 'paragraph-split' | 'paragraph-merge' | 'paragraph-format' | 'format-revision' | 'hyperlink' | 'revision' }>;
      partPath: string;
      paragraph: Element;
    }
  | {
      kind: 'table';
      patch: Extract<WordOoxmlPatch, { kind: 'table-structure' | 'table-insert' | 'table-cell-format' | 'table-cell-clear' | 'table-cell-replace' | 'table-cell-replace-rich' | 'table-delete' }>;
      partPath: string;
      table: Element;
      paragraph?: Element;
    }
  | {
      kind: 'drawing';
      patch: Extract<WordOoxmlPatch, { kind: 'drawing-delete' | 'drawing-resize' | 'drawing-insert' | 'drawing-revision' }>;
      partPath: string;
      drawing: Element;
      paragraph: Element | null;
    }
  | {
      kind: 'comment';
      patch: Extract<WordOoxmlPatch, { kind: 'comment' }>;
      partPath: string;
      paragraph: Element;
    };

export interface UnresolvedWordOoxmlPatchTarget {
  patch: WordOoxmlPatch;
  reason: string;
}

export interface WordOoxmlPatchTargetResolution {
  resolved: ResolvedWordOoxmlPatchTarget[];
  unresolved: UnresolvedWordOoxmlPatchTarget[];
}

export interface SkippedWordOoxmlPatch {
  patch: WordOoxmlPatch;
  reason: string;
}

export interface WordOoxmlPatchApplyResult {
  applied: WordOoxmlPatch[];
  skipped: SkippedWordOoxmlPatch[];
  unresolved: UnresolvedWordOoxmlPatchTarget[];
}

export interface NativeWordChangeApplyResult extends WordOoxmlPatchApplyResult {
  plan: WordOoxmlPatchPlan;
}

export interface UnsupportedNativeWordChange {
  change: NativeWordChange;
  reason: string;
}

export function planWordOoxmlPatches(changes: readonly NativeWordChange[]): WordOoxmlPatchPlan {
  const patches: WordOoxmlPatch[] = [];
  const unsupported: UnsupportedNativeWordChange[] = [];
  const tableResolution = new Map<string, 'accept' | 'reject'>();
  for (const change of changes) {
    if (change.kind === 'table' && change.revisionId && change.revisionAction) {
      tableResolution.set(change.revisionId, change.revisionAction);
    }
  }

  for (const change of changes) {
    if (change.kind === 'table' && change.revisionAction) continue;
    if (change.kind === 'table' && change.revisionId && tableResolution.get(change.revisionId) === 'reject') continue;
    const patch = planWordOoxmlPatch(change, {
      allowTrackedTable: change.kind === 'table' && !!change.revisionId && tableResolution.get(change.revisionId) === 'accept',
    });
    if (patch) {
      patches.push(patch);
    } else {
      unsupported.push({
        change,
        reason: unsupportedReason(change),
      });
    }
  }

  return {
    target: 'word/document.xml',
    patches,
    unsupported,
  };
}

export function applyNativeWordChangesToPackage(
  pkg: Package,
  changes: readonly NativeWordChange[],
  indices: PackageIndices = buildIndices(pkg),
): NativeWordChangeApplyResult {
  const plan = planWordOoxmlPatches(changes);
  const applyResult = applyWordOoxmlPatches(pkg, indices, plan.patches);
  return {
    plan,
    ...applyResult,
  };
}

export function planWordOoxmlPatch(
  change: NativeWordChange,
  options: { allowTrackedTable?: boolean } = {},
): WordOoxmlPatch | null {
  if (change.kind === 'text') {
    return {
      kind: 'paragraph-text',
      op: change.op,
      paraId: change.paraId,
      offset: change.offset,
      text: change.text,
      ...(change.deletedText !== undefined ? { deletedText: change.deletedText } : {}),
    };
  }

  if (change.kind === 'format' && change.op === 'run' && change.revisionId && change.revisionKind === 'format') {
    return {
      kind: 'format-revision',
      op: 'run',
      paraId: change.paraId,
      from: change.from,
      to: change.to,
      revisionId: change.revisionId,
      ...(change.beforeFormat !== undefined ? { beforeFormat: change.beforeFormat } : {}),
    };
  }

  if (change.kind === 'format' && change.op === 'run') {
    return {
      kind: 'paragraph-format',
      paraId: change.paraId,
      from: change.from,
      to: change.to,
      format: change.format,
    };
  }

  if (change.kind === 'format' && (change.op === 'accept' || change.op === 'reject') && change.revisionId) {
    return {
      kind: 'format-revision',
      op: change.op,
      paraId: change.paraId,
      from: change.from,
      to: change.to,
      revisionId: change.revisionId,
    };
  }

  if (change.kind === 'image' && change.op === 'delete') {
    return {
      kind: 'drawing-delete',
      paraId: change.paraId,
      drawingId: change.drawingId,
      offset: change.offset,
    };
  }

  if (change.kind === 'image' && change.op === 'resize' && change.width && change.height) {
    return {
      kind: 'drawing-resize',
      paraId: change.paraId,
      drawingId: change.drawingId,
      offset: change.offset,
      width: change.width,
      height: change.height,
    };
  }

  if (change.kind === 'image' && change.op === 'insert' && change.src && change.width && change.height) {
    return {
      kind: 'drawing-insert',
      paraId: change.paraId,
      drawingId: change.drawingId,
      offset: change.offset,
      src: change.src,
      width: change.width,
      height: change.height,
    };
  }

  if (change.kind === 'image' && (change.op === 'accept' || change.op === 'reject') && change.revisionId && (change.revisionKind === 'insert' || change.revisionKind === 'delete')) {
    return {
      kind: 'drawing-revision',
      op: change.op,
      paraId: change.paraId,
      drawingId: change.drawingId,
      offset: change.offset,
      revisionId: change.revisionId,
      revisionKind: change.revisionKind,
    };
  }

  if (change.kind === 'block' && change.op === 'split' && change.newParaId && change.offset !== undefined) {
    return {
      kind: 'paragraph-split',
      paraId: change.paraId,
      newParaId: change.newParaId,
      offset: change.offset,
    };
  }

  if (change.kind === 'block' && change.op === 'merge') {
    return {
      kind: 'paragraph-merge',
      paraId: change.paraId,
    };
  }

  if (change.kind === 'hyperlink') {
    if (change.from >= change.to) return null;
    const hyperlink = change.hyperlink;
    if (change.op === 'set' && (!hyperlink?.target || !isExternalHyperlinkTarget(hyperlink.target) || hyperlink.anchor)) {
      return null;
    }
    return {
      kind: 'hyperlink',
      op: change.op,
      paraId: change.paraId,
      from: change.from,
      to: change.to,
      ...(change.op === 'set' && hyperlink ? { hyperlink } : {}),
    };
  }

  if (
    change.kind === 'table' &&
    (!change.revisionId || options.allowTrackedTable === true) &&
    change.op !== 'insert-table' &&
    change.op !== 'delete-table' &&
    change.op !== 'format-cell' &&
    change.op !== 'clear-cell' &&
    change.op !== 'replace-cell'
  ) {
    return {
      kind: 'table-structure',
      op: change.op,
      tableId: change.tableId,
      rowIndex: change.rowIndex,
      columnIndex: change.columnIndex,
      ...(change.endColumnIndex !== undefined ? { endColumnIndex: change.endColumnIndex } : {}),
      ...(change.endRowIndex !== undefined ? { endRowIndex: change.endRowIndex } : {}),
    };
  }

  if (change.kind === 'table' && change.op === 'insert-table' && change.tableSnapshot && change.anchorParaId) {
    return {
      kind: 'table-insert',
      anchorParaId: change.anchorParaId,
      anchorOffset: change.anchorOffset ?? 0,
      table: change.tableSnapshot,
    };
  }

  if (change.kind === 'table' && change.op === 'format-cell' && change.format) {
    return {
      kind: 'table-cell-format',
      tableId: change.tableId,
      rowIndex: change.rowIndex,
      columnIndex: change.columnIndex,
      format: change.format,
    };
  }

  if (change.kind === 'table' && change.op === 'clear-cell') {
    return { kind: 'table-cell-clear', tableId: change.tableId, rowIndex: change.rowIndex, columnIndex: change.columnIndex };
  }

  if (change.kind === 'table' && change.op === 'replace-cell' && change.text !== undefined) {
    return { kind: 'table-cell-replace', tableId: change.tableId, rowIndex: change.rowIndex, columnIndex: change.columnIndex, text: change.text };
  }

  if (change.kind === 'table' && change.op === 'replace-cell' && change.cellSnapshot) {
    if (tableCellSnapshotHasUnsupportedInlineContent(change.cellSnapshot)) return null;
    return { kind: 'table-cell-replace-rich', tableId: change.tableId, rowIndex: change.rowIndex, columnIndex: change.columnIndex, cell: change.cellSnapshot };
  }


  if (change.kind === 'annotation' && (change.op === 'insert' || change.op === 'delete' || change.op === 'update')) {
    return {
      kind: 'comment',
      op: change.op,
      paraId: change.paraId,
      from: change.from,
      to: change.to,
      commentId: change.commentId,
      ...(change.parentCommentId !== undefined ? { parentCommentId: change.parentCommentId } : {}),
      ...(change.body !== undefined ? { body: change.body } : {}),
    };
  }

  if (change.kind === 'revision') {
    return {
      kind: 'revision',
      op: change.op,
      paraId: change.paraId,
      from: change.from,
      to: change.to,
      revisionId: change.revisionId,
      revisionKind: change.revisionKind,
    };
  }

  if (change.kind === 'table' && change.op === 'delete-table' && options.allowTrackedTable === true) {
    return { kind: 'table-delete', tableId: change.tableId };
  }

  return null;
}

export function resolveWordOoxmlPatchTargets(
  pkg: Package,
  indices: PackageIndices,
  patches: readonly WordOoxmlPatch[],
): WordOoxmlPatchTargetResolution {
  const resolved: ResolvedWordOoxmlPatchTarget[] = [];
  const unresolved: UnresolvedWordOoxmlPatchTarget[] = [];

  for (const patch of patches) {
    if (patch.kind === 'drawing-delete' || patch.kind === 'drawing-resize' || patch.kind === 'drawing-revision') {
      let drawing = indices.drawingIdx.get(patch.drawingId);
      if (!drawing && patch.kind === 'drawing-revision') {
        const paragraph = indices.paraIdx.get(patch.paraId.toUpperCase());
        drawing = paragraph ? findDrawingAtParagraphOffset(paragraph, patch.offset) ?? undefined : undefined;
      }
      if (!drawing) {
        unresolved.push({ patch, reason: `drawing not found: ${patch.drawingId}` });
        continue;
      }
      const partPath = partPathForNode(pkg, drawing);
      if (!partPath) {
        unresolved.push({ patch, reason: `drawing part not found: ${patch.drawingId}` });
        continue;
      }
      resolved.push({
        kind: 'drawing',
        patch,
        partPath,
        drawing,
        paragraph: indices.paraIdx.get(patch.paraId.toUpperCase()) ?? null,
      });
      continue;
    }

    if (patch.kind === 'drawing-insert') {
      const paragraph = indices.paraIdx.get(patch.paraId.toUpperCase());
      if (!paragraph) {
        unresolved.push({ patch, reason: `paragraph not found: ${patch.paraId}` });
        continue;
      }
      const partPath = partPathForNode(pkg, paragraph);
      if (!partPath) {
        unresolved.push({ patch, reason: `paragraph part not found: ${patch.paraId}` });
        continue;
      }
      resolved.push({ kind: 'drawing', patch, partPath, drawing: paragraph, paragraph });
      continue;
    }

    if (patch.kind === 'comment') {
      const paragraph = indices.paraIdx.get(patch.paraId.toUpperCase());
      if (!paragraph) {
        unresolved.push({ patch, reason: `paragraph not found: ${patch.paraId}` });
        continue;
      }
      const partPath = partPathForNode(pkg, paragraph);
      if (!partPath) {
        unresolved.push({ patch, reason: `paragraph part not found: ${patch.paraId}` });
        continue;
      }
      resolved.push({ kind: 'comment', patch, partPath, paragraph });
      continue;
    }

    if (patch.kind === 'table-insert') {
      const paragraph = indices.paraIdx.get(patch.anchorParaId.toUpperCase());
      if (!paragraph) {
        unresolved.push({ patch, reason: `anchor paragraph not found: ${patch.anchorParaId}` });
        continue;
      }
      const partPath = partPathForNode(pkg, paragraph);
      if (!partPath) {
        unresolved.push({ patch, reason: `paragraph part not found: ${patch.anchorParaId}` });
        continue;
      }
      resolved.push({ kind: 'table', patch, partPath, table: paragraph, paragraph });
      continue;
    }

    if (patch.kind === 'table-structure' || patch.kind === 'table-cell-format' || patch.kind === 'table-cell-clear' || patch.kind === 'table-cell-replace' || patch.kind === 'table-cell-replace-rich' || patch.kind === 'table-delete') {
      const table = indices.tblIdx.get(patch.tableId);
      if (!table) {
        unresolved.push({ patch, reason: `table not found: ${patch.tableId}` });
        continue;
      }
      const partPath = partPathForNode(pkg, table);
      if (!partPath) {
        unresolved.push({ patch, reason: `table part not found: ${patch.tableId}` });
        continue;
      }
      resolved.push({ kind: 'table', patch, partPath, table });
      continue;
    }

    const paragraph = indices.paraIdx.get(patch.paraId.toUpperCase());
    if (!paragraph) {
      unresolved.push({ patch, reason: `paragraph not found: ${patch.paraId}` });
      continue;
    }
    const partPath = partPathForNode(pkg, paragraph);
    if (!partPath) {
      unresolved.push({ patch, reason: `paragraph part not found: ${patch.paraId}` });
      continue;
    }
    resolved.push({
      kind: 'paragraph',
      patch,
      partPath,
      paragraph,
    });
  }

  return { resolved, unresolved };
}

export function applyWordOoxmlPatches(
  pkg: Package,
  indices: PackageIndices,
  patches: readonly WordOoxmlPatch[],
): WordOoxmlPatchApplyResult {
  const applied: WordOoxmlPatch[] = [];
  const skipped: SkippedWordOoxmlPatch[] = [];
  const unresolved: UnresolvedWordOoxmlPatchTarget[] = [];
  const resolvedRevisionIds = new Set<string>();
  let currentIndices = indices;

  for (const patch of patches) {
    const resolution = resolveWordOoxmlPatchTargets(pkg, currentIndices, [patch]);
    if (resolution.unresolved.length > 0) {
      unresolved.push(...resolution.unresolved);
      continue;
    }
    const target = resolution.resolved[0];
    if (!target) continue;
    const part = partForNode(pkg, target.kind === 'paragraph' ? target.paragraph : target.kind === 'drawing' ? target.drawing : target.kind === 'table' ? target.table : target.paragraph);
    if (!part) {
      skipped.push({ patch: target.patch, reason: `part not found: ${target.partPath}` });
      continue;
    }

    if (target.kind === 'drawing') {
      if (target.patch.kind === 'drawing-insert') {
        const inserted = insertDrawingIntoParagraph(pkg, target.partPath, target.paragraph ?? target.drawing, target.patch);
        if (!inserted.ok) {
          skipped.push({ patch: target.patch, reason: inserted.reason });
          continue;
        }
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      if (target.patch.kind === 'drawing-revision') {
        const shouldRemove =
          (target.patch.revisionKind === 'insert' && target.patch.op === 'reject') ||
          (target.patch.revisionKind === 'delete' && target.patch.op === 'accept');
        if (shouldRemove) removeDrawing(target.drawing);
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      if (target.patch.kind === 'drawing-resize') {
        const resized = resizeDrawing(target.drawing, target.patch.width, target.patch.height);
        if (!resized.ok) {
          skipped.push({ patch: target.patch, reason: resized.reason });
          continue;
        }
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      removeDrawing(target.drawing);
      markDirty(part);
      applied.push(target.patch);
      currentIndices = buildIndices(pkg);
      continue;
    }

    if (target.kind === 'table') {
      if (target.patch.kind === 'table-insert') {
        const inserted = createNativeTableElement(target.table.ownerDocument, target.patch.table, pkg, target.partPath);
        target.table.parentNode?.insertBefore(inserted, target.table.nextSibling);
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      if (target.patch.kind === 'table-cell-format') {
        const formatted = applyTableCellFormatPatch(target.table, target.patch);
        if (!formatted.ok) {
          skipped.push({ patch: target.patch, reason: formatted.reason });
          continue;
        }
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      if (target.patch.kind === 'table-cell-clear') {
        const cleared = clearTableCellContent(target.table, target.patch);
        if (!cleared.ok) {
          skipped.push({ patch: target.patch, reason: cleared.reason });
          continue;
        }
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      if (target.patch.kind === 'table-cell-replace') {
        const replaced = replaceTableCellText(target.table, target.patch);
        if (!replaced.ok) {
          skipped.push({ patch: target.patch, reason: replaced.reason });
          continue;
        }
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      if (target.patch.kind === 'table-cell-replace-rich') {
        const replaced = replaceTableCellSnapshot(pkg, target.partPath, target.table, target.patch);
        if (!replaced.ok) {
          skipped.push({ patch: target.patch, reason: replaced.reason });
          continue;
        }
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      if (target.patch.kind === 'table-delete') {
        removeNode(target.table);
        markDirty(part);
        applied.push(target.patch);
        currentIndices = buildIndices(pkg);
        continue;
      }
      const appliedTable = applyTableStructurePatch(target.table, target.patch);
      if (!appliedTable.ok) {
        skipped.push({ patch: target.patch, reason: appliedTable.reason });
        continue;
      }
      markDirty(part);
      applied.push(target.patch);
      currentIndices = buildIndices(pkg);
      continue;
    }

    if (target.kind === 'comment') {
      const comment = applyCommentPatch(pkg, target.partPath, target.paragraph, target.patch);
      if (!comment.ok) {
        skipped.push({ patch: target.patch, reason: comment.reason });
        continue;
      }
      markDirty(part);
      applied.push(target.patch);
      currentIndices = buildIndices(pkg);
      continue;
    }

    if (target.patch.kind === 'revision') {
      const revision = applyRevisionPatch(target.paragraph, target.patch);
      if (!revision.ok) {
        if (
          (target.patch.op === 'accept' || target.patch.op === 'reject') &&
          revision.reason === `revision not found: ${target.patch.revisionId}` &&
          resolvedRevisionIds.has(target.patch.revisionId)
        ) {
          applied.push(target.patch);
          continue;
        }
        skipped.push({ patch: target.patch, reason: revision.reason });
        continue;
      }
      markDirty(part);
      applied.push(target.patch);
      if (target.patch.op === 'accept' || target.patch.op === 'reject') resolvedRevisionIds.add(target.patch.revisionId);
      currentIndices = buildIndices(pkg);
      continue;
    }

    if (target.patch.kind === 'paragraph-format') {
      const formatted = applyParagraphFormatPatch(target.paragraph, target.patch);
      if (!formatted.ok) {
        skipped.push({ patch: target.patch, reason: formatted.reason });
        continue;
      }
      markDirty(part);
      applied.push(target.patch);
      currentIndices = buildIndices(pkg);
      continue;
    }

    if (target.patch.kind === 'format-revision') {
      const formatted = applyFormatRevisionPatch(target.paragraph, target.patch);
      if (!formatted.ok) {
        skipped.push({ patch: target.patch, reason: formatted.reason });
        continue;
      }
      markDirty(part);
      applied.push(target.patch);
      currentIndices = buildIndices(pkg);
      continue;
    }

    if (target.patch.kind === 'paragraph-split') {
      const split = splitParagraphElement(target.paragraph, target.patch.offset, target.patch.newParaId);
      if (!split.ok) {
        skipped.push({ patch: target.patch, reason: split.reason });
        continue;
      }
      markDirty(part);
      applied.push(target.patch);
      currentIndices = buildIndices(pkg);
      continue;
    }

    if (target.patch.kind === 'paragraph-merge') {
      const merge = mergeNextParagraphElement(target.paragraph);
      if (!merge.ok) {
        skipped.push({ patch: target.patch, reason: merge.reason });
        continue;
      }
      markDirty(part);
      applied.push(target.patch);
      currentIndices = buildIndices(pkg);
      continue;
    }

    if (target.patch.kind === 'hyperlink') {
      const hyperlink = applyParagraphHyperlinkPatch(pkg, target.partPath, target.paragraph, target.patch);
      if (!hyperlink.ok) {
        skipped.push({ patch: target.patch, reason: hyperlink.reason });
        continue;
      }
      markDirty(part);
      applied.push(target.patch);
      currentIndices = buildIndices(pkg);
      continue;
    }

    const appliedText = applyParagraphTextPatch(target.paragraph, target.patch);
    if (!appliedText.ok) {
      skipped.push({ patch: target.patch, reason: appliedText.reason });
      continue;
    }
    markDirty(part);
    applied.push(target.patch);
  }

  return {
    applied,
    skipped,
    unresolved,
  };
}

function unsupportedReason(change: NativeWordChange): string {
  if (change.kind === 'block' && change.op === 'split') {
    return 'paragraph split change is missing newParaId or offset';
  }
  if (change.kind === 'hyperlink') {
    if (change.op === 'set' && change.hyperlink?.anchor) return 'hyperlink anchors are not yet written back';
    return 'only external http/https hyperlinks can be written back';
  }
  if (change.kind === 'image' && change.op === 'resize') return 'image resize is missing width or height';
  if (change.kind === 'image' && change.op === 'insert') return 'image insert requires data URL, width, and height';
  if (change.kind === 'table') {
    if (change.op === 'replace-cell' && change.cellSnapshot && tableCellSnapshotHasUnsupportedInlineContent(change.cellSnapshot)) {
      return 'table cell rich replacement does not yet write back drawings or hyperlink anchors';
    }
    if (change.revisionId || change.revisionKind) return 'table revision writeback requires an original table structure snapshot';
    return 'table structure change is not supported for DOCX writeback';
  }
  return `unsupported native change: ${change.kind}/${'op' in change ? change.op : 'unknown'}`;
}

function tableCellSnapshotHasUnsupportedInlineContent(cell: NativeWordTableCellNode): boolean {
  return cell.blocks.some((block) =>
    block.kind !== 'paragraph' || block.content.some((child) => child.kind === 'drawing' || (child.kind === 'text' && child.hyperlink !== undefined && !child.hyperlink.target)),
  );
}

function isExternalHyperlinkTarget(target: string): boolean {
  try {
    const protocol = new URL(target).protocol.toLowerCase();
    return protocol === 'http:' || protocol === 'https:';
  } catch {
    return false;
  }
}

function applyParagraphHyperlinkPatch(
  pkg: Package,
  partPath: string,
  paragraph: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'hyperlink' }>,
): { ok: true } | { ok: false; reason: string } {
  const children = directChildren(paragraph).filter((child) => !(child.namespaceURI === W_NS && child.localName === 'pPr'));
  const totalLength = children.reduce((sum, child) => sum + inlineNodeLength(child), 0);
  if (patch.from < 0 || patch.to > totalLength || patch.from >= patch.to) {
    return { ok: false, reason: `hyperlink range out of range: ${patch.from}:${patch.to}` };
  }
  if (children.some((child) => child.namespaceURI !== W_NS || !['r', 'hyperlink'].includes(child.localName ?? ''))) {
    return { ok: false, reason: 'hyperlink patch only supports text runs and existing hyperlinks' };
  }
  if (children.some((child) => inlineNodeContainsObject(child))) {
    return { ok: false, reason: 'hyperlink patch cannot cross inline objects' };
  }

  let relationshipId: string | null = null;
  let relsPart: Part | null = null;
  if (patch.op === 'set') {
    const target = patch.hyperlink?.target;
    if (!target || !isExternalHyperlinkTarget(target)) {
      return { ok: false, reason: 'only external http/https hyperlinks can be written back' };
    }
    const relationship = ensureExternalHyperlinkRelationship(pkg, partPath, target);
    relationshipId = relationship.id;
    relsPart = relationship.part;
  }

  const nextChildren: Element[] = [];
  let position = 0;
  const selectedChildren: Element[] = [];
  for (const child of children) {
    const length = inlineNodeLength(child);
    const localFrom = Math.max(0, patch.from - position);
    const localTo = Math.min(length, patch.to - position);
    const intersects = localFrom < localTo;
    if (!intersects) {
      nextChildren.push(child);
      position += length;
      continue;
    }
    if (localFrom > 0) {
      const before = sliceInlineElement(child, 0, localFrom);
      if (before) nextChildren.push(before);
    }
    const selected = sliceInlineElement(child, localFrom, localTo);
    if (selected) selectedChildren.push(...unwrapHyperlinkSlice(selected));
    if (localTo < length) {
      const after = sliceInlineElement(child, localTo, length);
      if (after) nextChildren.push(after);
    }
    position += length;
  }

  if (patch.op === 'set') {
    const doc = paragraph.ownerDocument;
    if (!doc || !relationshipId) return { ok: false, reason: 'paragraph has no owner document or relationship' };
    const wrapper = doc.createElementNS(W_NS, 'w:hyperlink');
    wrapper.setAttributeNS(R_NS, 'r:id', relationshipId);
    const tooltip = patch.hyperlink?.tooltip?.trim();
    if (tooltip) wrapper.setAttributeNS(W_NS, 'w:tooltip', tooltip);
    for (const selected of selectedChildren) wrapper.appendChild(selected);
    insertSelectedHyperlink(nextChildren, wrapper, patch);
  } else {
    insertSelectedNodes(nextChildren, selectedChildren, patch);
  }

  for (const child of children) removeNode(child);
  for (const child of nextChildren) paragraph.appendChild(child);
  if (relsPart) markDirty(relsPart);
  return { ok: true };
}

function insertSelectedHyperlink(nextChildren: Element[], wrapper: Element, patch: Extract<WordOoxmlPatch, { kind: 'hyperlink' }>): void {
  insertSelectedNodes(nextChildren, [wrapper], patch);
}

function insertSelectedNodes(nextChildren: Element[], selected: Element[], patch: Extract<WordOoxmlPatch, { kind: 'hyperlink' }>): void {
  // Rebuild once so the selected content lands at its original logical offset.
  let cursor = 0;
  let inserted = false;
  const rebuilt: Element[] = [];
  for (const child of nextChildren) {
    if (!inserted && cursor >= patch.from) {
      rebuilt.push(...selected);
      inserted = true;
    }
    rebuilt.push(child);
    cursor += inlineNodeLength(child);
  }
  if (!inserted) rebuilt.push(...selected);
  nextChildren.splice(0, nextChildren.length, ...rebuilt);
}

function sliceInlineElement(element: Element, from: number, to: number): Element | null {
  if (to <= from) return null;
  const clone = element.cloneNode(true) as Element;
  applyTextSlice(clone, from, to);
  return inlineNodeLength(clone) > 0 ? clone : null;
}

function unwrapHyperlinkSlice(element: Element): Element[] {
  if (!(element.namespaceURI === W_NS && element.localName === 'hyperlink')) return [element];
  return directChildren(element).filter((child) => child.namespaceURI === W_NS && child.localName === 'r');
}

function inlineNodeContainsObject(element: Element): boolean {
  return element.getElementsByTagNameNS(W_NS, 'drawing').length > 0 || element.getElementsByTagNameNS(W_NS, 'pict').length > 0;
}

function ensureExternalHyperlinkRelationship(pkg: Package, partPath: string, target: string): { id: string; part: Part } {
  const relsPath = relsPathFor(partPath);
  let part = pkg.parts.get(relsPath);
  if (!part?.xmlDom) {
    const sourceDocument = Array.from(pkg.parts.values()).find((candidate) => candidate.xmlDom)?.xmlDom;
    if (!sourceDocument) throw new Error('cannot create hyperlink relationships without an XML part');
    const dom = sourceDocument.implementation?.createDocument(PACKAGE_RELS_NS, 'Relationships', null) as unknown as Document;
    if (!dom) throw new Error('XML implementation cannot create relationship document');
    part = createXmlPart({ path: relsPath, contentType: RELS_CONTENT_TYPE, xmlDom: dom });
    setPart(pkg, part);
  }
  const relsDom = part.xmlDom;
  if (!relsDom) throw new Error(`relationship part has no XML DOM: ${relsPath}`);
  const relationships = relsDom.getElementsByTagNameNS(PACKAGE_RELS_NS, 'Relationship');
  for (let index = 0; index < relationships.length; index += 1) {
    const relationship = relationships[index];
    if (relationship?.getAttribute('Type') === HYPERLINK_REL_TYPE && relationship.getAttribute('Target') === target) {
      return { id: relationship.getAttribute('Id') ?? '', part };
    }
  }
  const used = new Set<string>();
  for (let index = 0; index < relationships.length; index += 1) {
    const id = relationships[index]?.getAttribute('Id');
    if (id) used.add(id);
  }
  let counter = 1;
  while (used.has(`rId${counter}`)) counter += 1;
  const relationship = relsDom.createElementNS(PACKAGE_RELS_NS, 'Relationship');
  relationship.setAttribute('Id', `rId${counter}`);
  relationship.setAttribute('Type', HYPERLINK_REL_TYPE);
  relationship.setAttribute('Target', target);
  relationship.setAttribute('TargetMode', 'External');
  relsDom.documentElement?.appendChild(relationship);
  return { id: `rId${counter}`, part };
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash >= 0 ? partPath.slice(0, slash) : '';
  const file = partPath.slice(slash + 1);
  return dir ? `${dir}/_rels/${file}.rels` : `_rels/${file}.rels`;
}

function createNativeTableElement(document: Document, table: NativeWordTableNode, pkg: Package, partPath: string): Element {
  const create = (name: string): Element => document.createElementNS(W_NS, `w:${name}`);
  const tableElement = create('tbl');
  const properties = create('tblPr');
  const width = create('tblW');
  width.setAttributeNS(W_NS, 'w:w', String(Math.max(0, Math.round(table.layout?.width ?? 0))));
  width.setAttributeNS(W_NS, 'w:type', table.layout?.widthType === 'pct' ? 'pct' : 'dxa');
  properties.appendChild(width);
  const borders = create('tblBorders');
  for (const edge of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const) {
    const border = create(edge);
    border.setAttributeNS(W_NS, 'w:val', 'single');
    border.setAttributeNS(W_NS, 'w:sz', '4');
    border.setAttributeNS(W_NS, 'w:color', table.borders?.[edge as keyof NonNullable<NativeWordTableNode['borders']>]?.color ?? 'D9D9D9');
    borders.appendChild(border);
  }
  properties.appendChild(borders);
  tableElement.appendChild(properties);
  const grid = create('tblGrid');
  for (const columnWidth of table.columnWidths) {
    const column = create('gridCol');
    column.setAttributeNS(W_NS, 'w:w', String(Math.max(1, Math.round(columnWidth))));
    grid.appendChild(column);
  }
  tableElement.appendChild(grid);
  for (const row of table.rows) {
    const rowElement = create('tr');
    for (const cell of row.cells) {
      rowElement.appendChild(createNativeTableCellElement(document, cell, pkg, partPath));
    }
    tableElement.appendChild(rowElement);
  }
  return tableElement;
}

function createNativeTableCellElement(document: Document, cell: NativeWordTableCellNode, pkg: Package, partPath: string): Element {
  const create = (name: string): Element => document.createElementNS(W_NS, `w:${name}`);
  const cellElement = create('tc');
  const properties = create('tcPr');
  const width = create('tcW');
  width.setAttributeNS(W_NS, 'w:w', String(Math.max(1, Math.round(cell.width ?? 0))));
  width.setAttributeNS(W_NS, 'w:type', 'dxa');
  properties.appendChild(width);
  if (cell.shading?.fill) {
    const shading = create('shd');
    shading.setAttributeNS(W_NS, 'w:val', cell.shading.val ?? 'clear');
    shading.setAttributeNS(W_NS, 'w:fill', cell.shading.fill.replace(/^#/, '').toUpperCase());
    properties.appendChild(shading);
  }
  if (cell.verticalAlignment) {
    const alignment = create('vAlign');
    alignment.setAttributeNS(W_NS, 'w:val', cell.verticalAlignment);
    properties.appendChild(alignment);
  }
  cellElement.appendChild(properties);
  const paragraphs = cell.blocks.filter((block): block is NativeWordParagraphNode => block.kind === 'paragraph');
  const paragraphNodes = paragraphs.length > 0 ? paragraphs : [null];
  for (const block of paragraphNodes) {
    const paragraph = create('p');
    if (block) {
      for (const child of block.content) {
        if (child.kind === 'image') {
          const image = parseImageDataUrl(child.src);
          if (!image) continue;
          const relationship = ensureImageRelationship(pkg, partPath, image.bytes, image.extension, image.contentType);
          paragraph.appendChild(createInlinePicture(document, relationship.id, child.drawingId, child.width ?? 96, child.height ?? 96));
          markDirty(relationship.part);
          markDirty(relationship.mediaPart);
          markContentTypeDirty(pkg, image.extension, image.contentType);
          continue;
        }
        if (child.kind !== 'text') continue;
        const run = create('r');
        if (child.format) {
          const runProperties = create('rPr');
          if (child.format.bold) runProperties.appendChild(create('b'));
          if (child.format.italic) runProperties.appendChild(create('i'));
          if (child.format.underline) {
            const underline = create('u');
            underline.setAttributeNS(W_NS, 'w:val', 'single');
            runProperties.appendChild(underline);
          }
          if (child.format.strike) runProperties.appendChild(create('strike'));
          if (child.format.color) {
            const color = create('color');
            color.setAttributeNS(W_NS, 'w:val', child.format.color.replace(/^#/, '').toUpperCase());
            runProperties.appendChild(color);
          }
          if (child.format.fontSize) {
            const size = create('sz');
            size.setAttributeNS(W_NS, 'w:val', String(Math.round(child.format.fontSize * 2)));
            runProperties.appendChild(size);
          }
          if (child.format.fontFamily) {
            const fonts = create('rFonts');
            for (const attribute of ['ascii', 'hAnsi', 'eastAsia']) fonts.setAttributeNS(W_NS, `w:${attribute}`, child.format.fontFamily);
            runProperties.appendChild(fonts);
          }
          if (runProperties.childNodes.length > 0) run.appendChild(runProperties);
        }
        const text = create('t');
        text.textContent = child.text;
        if (/^\s|\s$/.test(child.text)) text.setAttributeNS(XML_NS, 'xml:space', 'preserve');
         run.appendChild(text);
         if (child.hyperlink?.target) {
           const relationship = ensureExternalHyperlinkRelationship(pkg, partPath, child.hyperlink.target);
           const hyperlink = document.createElementNS(W_NS, 'w:hyperlink');
           hyperlink.setAttributeNS(R_NS, 'r:id', relationship.id);
           hyperlink.appendChild(run);
           paragraph.appendChild(hyperlink);
           markDirty(relationship.part);
         } else {
           paragraph.appendChild(run);
         }
      }
    }
    if (!paragraph.firstChild) paragraph.appendChild(create('r'));
    cellElement.appendChild(paragraph);
  }
  return cellElement;
}

function applyTableCellFormatPatch(
  table: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'table-cell-format' }>,
): { ok: true } | { ok: false; reason: string } {
  const rows = directChildren(table).filter((child) => child.namespaceURI === W_NS && child.localName === 'tr');
  const row = rows[patch.rowIndex];
  if (!row) return { ok: false, reason: `table row out of range: ${patch.rowIndex}` };
  const cells = directChildren(row).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc');
  const cell = cells[patch.columnIndex];
  if (!cell) return { ok: false, reason: `table column out of range: ${patch.columnIndex}` };
  const document = table.ownerDocument;
  let properties = directChildren(cell).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr');
  if (!properties) {
    properties = document.createElementNS(W_NS, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
  }
  if (patch.format.backgroundColor !== undefined) {
    const existing = directChildren(properties).find((child) => child.namespaceURI === W_NS && child.localName === 'shd');
    if (patch.format.backgroundColor === null) {
      if (existing) removeNode(existing);
    } else {
      const shading = existing ?? document.createElementNS(W_NS, 'w:shd');
      shading.setAttributeNS(W_NS, 'w:val', 'clear');
      shading.setAttributeNS(W_NS, 'w:fill', patch.format.backgroundColor.replace(/^#/, '').toUpperCase());
      if (!existing) properties.appendChild(shading);
    }
  }
  if (patch.format.verticalAlignment !== undefined) {
    const existing = directChildren(properties).find((child) => child.namespaceURI === W_NS && child.localName === 'vAlign');
    if (patch.format.verticalAlignment === null) {
      if (existing) removeNode(existing);
    } else {
      const alignment = existing ?? document.createElementNS(W_NS, 'w:vAlign');
      alignment.setAttributeNS(W_NS, 'w:val', patch.format.verticalAlignment);
      if (!existing) properties.appendChild(alignment);
    }
  }
  if (patch.format.borderColor !== undefined) {
    const existing = directChildren(properties).find((child) => child.namespaceURI === W_NS && child.localName === 'tcBorders');
    const borders = existing ?? document.createElementNS(W_NS, 'w:tcBorders');
    for (const edge of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const) {
      let border = directChildren(borders).find((child) => child.namespaceURI === W_NS && child.localName === edge);
      if (!border) {
        border = document.createElementNS(W_NS, `w:${edge}`);
        borders.appendChild(border);
      }
      border.setAttributeNS(W_NS, 'w:val', patch.format.borderColor === null ? 'nil' : 'single');
      if (patch.format.borderColor !== null) {
        border.setAttributeNS(W_NS, 'w:sz', '4');
        border.setAttributeNS(W_NS, 'w:color', patch.format.borderColor.replace(/^#/, '').toUpperCase());
      }
    }
    if (!existing) properties.appendChild(borders);
  }
  return { ok: true };
}

function clearTableCellContent(
  table: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'table-cell-clear' }>,
): { ok: true } | { ok: false; reason: string } {
  const rows = directChildren(table).filter((child) => child.namespaceURI === W_NS && child.localName === 'tr');
  const row = rows[patch.rowIndex];
  if (!row) return { ok: false, reason: `table row out of range: ${patch.rowIndex}` };
  const cells = directChildren(row).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc');
  const cell = cells[patch.columnIndex];
  if (!cell) return { ok: false, reason: `table column out of range: ${patch.columnIndex}` };
  const properties = directChildren(cell).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr');
  for (const child of directChildren(cell)) {
    if (child !== properties) removeNode(child);
  }
  cell.appendChild(cell.ownerDocument.createElementNS(W_NS, 'w:p'));
  return { ok: true };
}

function replaceTableCellText(
  table: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'table-cell-replace' }>,
): { ok: true } | { ok: false; reason: string } {
  const rows = directChildren(table).filter((child) => child.namespaceURI === W_NS && child.localName === 'tr');
  const row = rows[patch.rowIndex];
  if (!row) return { ok: false, reason: `table row out of range: ${patch.rowIndex}` };
  const cells = directChildren(row).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc');
  const cell = cells[patch.columnIndex];
  if (!cell) return { ok: false, reason: `table column out of range: ${patch.columnIndex}` };
  const properties = directChildren(cell).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr');
  const sourceParagraph = directChildren(cell).find((child) => child.namespaceURI === W_NS && child.localName === 'p');
  for (const child of directChildren(cell)) {
    if (child !== properties) removeNode(child);
  }
  const paragraph = cell.ownerDocument.createElementNS(W_NS, 'w:p');
  const pPr = sourceParagraph && directChildren(sourceParagraph).find((child) => child.namespaceURI === W_NS && child.localName === 'pPr');
  if (pPr) paragraph.appendChild(pPr.cloneNode(true));
  const run = cell.ownerDocument.createElementNS(W_NS, 'w:r');
  const text = cell.ownerDocument.createElementNS(W_NS, 'w:t');
  text.textContent = patch.text;
  if (/^\s|\s$/.test(patch.text)) text.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  run.appendChild(text);
  paragraph.appendChild(run);
  cell.appendChild(paragraph);
  return { ok: true };
}

function replaceTableCellSnapshot(
  pkg: Package,
  partPath: string,
  table: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'table-cell-replace-rich' }>,
): { ok: true } | { ok: false; reason: string } {
  const rows = directChildren(table).filter((child) => child.namespaceURI === W_NS && child.localName === 'tr');
  const row = rows[patch.rowIndex];
  if (!row) return { ok: false, reason: `table row out of range: ${patch.rowIndex}` };
  const cells = directChildren(row).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc');
  const target = cells[patch.columnIndex];
  if (!target) return { ok: false, reason: `table column out of range: ${patch.columnIndex}` };
  const generated = createNativeTableCellElement(table.ownerDocument, patch.cell, pkg, partPath);
  for (const child of directChildren(target)) removeNode(child);
  for (const child of directChildren(generated)) target.appendChild(child.cloneNode(true));
  return { ok: true };
}

function applyTableStructurePatch(
  table: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'table-structure' }>,
): { ok: true } | { ok: false; reason: string } {
  const rows = directChildren(table).filter((child) => child.namespaceURI === W_NS && child.localName === 'tr');
  if (rows.length === 0) return { ok: false, reason: 'table has no rows' };
  if (patch.op === 'insert-row' || patch.op === 'delete-row') {
    if (patch.rowIndex < 0 || patch.rowIndex >= rows.length + (patch.op === 'insert-row' ? 1 : 0)) {
      return { ok: false, reason: `table row out of range: ${patch.rowIndex}` };
    }
    if (patch.op === 'delete-row') {
      if (rows.length <= 1) return { ok: false, reason: 'cannot delete the only table row' };
      removeNode(rows[patch.rowIndex]!);
      return { ok: true };
    }
    const template = rows[Math.min(patch.rowIndex, rows.length - 1)]!;
    const inserted = blankTableRowElement(template);
    table.insertBefore(inserted, rows[patch.rowIndex] ?? null);
    return { ok: true };
  }

  const row = rows[patch.rowIndex];
  if (!row) return { ok: false, reason: `table row out of range: ${patch.rowIndex}` };
  const cells = directChildren(row).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc');
  if (patch.op === 'insert-column' || patch.op === 'delete-column') {
    if (cells.length === 0) return { ok: false, reason: 'table row has no cells' };
    if (patch.columnIndex < 0 || patch.columnIndex >= cells.length + (patch.op === 'insert-column' ? 1 : 0)) {
      return { ok: false, reason: `table column out of range: ${patch.columnIndex}` };
    }
    const allRows = rows.map((candidate) => directChildren(candidate).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc'));
    if (allRows.some((candidate) => candidate.length !== cells.length || candidate.some((cell) => cellGridSpan(cell) !== 1 || cellHasVerticalMerge(cell)))) {
      return { ok: false, reason: 'table column structure is not a regular non-merged grid' };
    }
    for (const rowCells of allRows) {
      if (patch.op === 'delete-column') {
        if (rowCells.length <= 1) return { ok: false, reason: 'cannot delete the only table column' };
        removeNode(rowCells[patch.columnIndex]!);
      } else {
        const template = rowCells[Math.min(patch.columnIndex, rowCells.length - 1)]!;
        const inserted = blankTableCellElement(template);
        rowCells[0]!.parentNode?.insertBefore(inserted, rowCells[patch.columnIndex] ?? null);
      }
    }
    updateTableGrid(table, patch.op === 'insert-column' ? 'insert' : 'delete', patch.columnIndex);
    return { ok: true };
  }

  if (patch.op === 'merge-cells') {
    if (patch.endRowIndex !== undefined && patch.endRowIndex > patch.rowIndex) {
      if (patch.columnIndex < 0 || patch.endRowIndex >= rows.length) {
        return { ok: false, reason: `vertical merge row out of range: ${patch.rowIndex}:${patch.endRowIndex}` };
      }
      const selected = rows.slice(patch.rowIndex, patch.endRowIndex + 1).map((candidate) =>
        directChildren(candidate).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc')[patch.columnIndex],
      );
      if (selected.some((cell) => !cell || cellGridSpan(cell) !== 1 || cellHasVerticalMerge(cell))) {
        return { ok: false, reason: 'vertical merge requires a regular non-merged column' };
      }
      setCellVerticalMerge(selected[0]!, 'restart');
      for (const cell of selected.slice(1)) {
        setCellVerticalMerge(cell!, 'continue');
        const properties = directChildren(cell!).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr');
        for (const child of directChildren(cell!)) if (child !== properties) removeNode(child);
        cell!.appendChild(cell!.ownerDocument.createElementNS(W_NS, 'w:p'));
      }
      return { ok: true };
    }
    const endColumnIndex = patch.endColumnIndex ?? patch.columnIndex + 1;
    if (patch.columnIndex < 0 || endColumnIndex >= cells.length || endColumnIndex <= patch.columnIndex) {
      return { ok: false, reason: `merge start column out of range: ${patch.columnIndex}` };
    }
    if (cells.some((cell) => cellGridSpan(cell) !== 1 || cellHasVerticalMerge(cell))) {
      return { ok: false, reason: 'cell merge requires a regular non-merged row' };
    }
    const first = cells[patch.columnIndex]!;
    for (let index = patch.columnIndex + 1; index <= endColumnIndex; index += 1) {
      const cell = cells[index]!;
      appendCellContent(first, cell);
      removeNode(cell);
    }
    setCellGridSpan(first, endColumnIndex - patch.columnIndex + 1);
    return { ok: true };
  }

  if (patch.endRowIndex !== undefined && patch.endRowIndex > patch.rowIndex) {
    if (patch.columnIndex < 0 || patch.endRowIndex >= rows.length) {
      return { ok: false, reason: `vertical split row out of range: ${patch.rowIndex}:${patch.endRowIndex}` };
    }
    const selected = rows.slice(patch.rowIndex, patch.endRowIndex + 1).map((candidate) =>
      directChildren(candidate).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc')[patch.columnIndex],
    );
    if (selected.some((candidate, index) => {
      if (!candidate || cellGridSpan(candidate) !== 1) return true;
      const merge = directChildren(candidate).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr')
        ?.getElementsByTagNameNS(W_NS, 'vMerge')[0];
      const value = merge?.getAttributeNS(W_NS, 'val');
      return index === 0 ? value !== 'restart' : value !== 'continue';
    })) {
      return { ok: false, reason: 'cell is not a vertically merged cell' };
    }
    for (const candidate of selected) removeCellVerticalMerge(candidate!);
    return { ok: true };
  }

  const splitRow = rows[patch.rowIndex]!;
  const splitCells = directChildren(splitRow).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc');
  let logicalColumn = 0;
  const cell = splitCells.find((candidate) => {
    const matches = logicalColumn === patch.columnIndex;
    logicalColumn += cellGridSpan(candidate);
    return matches;
  });
  if (!cell) return { ok: false, reason: `split column out of range: ${patch.columnIndex}` };
  if (cellHasVerticalMerge(cell) || cellGridSpan(cell) <= 1) return { ok: false, reason: 'cell is not a horizontally merged cell' };
  const span = cellGridSpan(cell);
  removeCellGridSpan(cell);
  for (let index = 1; index < span; index += 1) {
    const blank = blankTableCellElement(cell);
    splitRow.insertBefore(blank, cell.nextSibling);
  }
  return { ok: true };
}

function setCellVerticalMerge(cell: Element, value: 'restart' | 'continue'): void {
  const document = cell.ownerDocument;
  let properties = directChildren(cell).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr');
  if (!properties) {
    properties = document.createElementNS(W_NS, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
  }
  const merge = directChildren(properties).find((child) => child.namespaceURI === W_NS && child.localName === 'vMerge')
    ?? document.createElementNS(W_NS, 'w:vMerge');
  merge.setAttributeNS(W_NS, 'w:val', value);
  if (!merge.parentNode) properties.appendChild(merge);
}

function removeCellVerticalMerge(cell: Element): void {
  const properties = directChildren(cell).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr');
  const merge = properties && directChildren(properties).find((child) => child.namespaceURI === W_NS && child.localName === 'vMerge');
  if (merge) removeNode(merge);
}

function blankTableRowElement(template: Element): Element {
  const clone = template.cloneNode(true) as Element;
  const cells = directChildren(clone).filter((child) => child.namespaceURI === W_NS && child.localName === 'tc');
  for (const cell of cells) {
    const blank = blankTableCellElement(cell);
    replaceNode(cell, blank);
  }
  return clone;
}

function blankTableCellElement(template: Element): Element {
  const doc = template.ownerDocument;
  const blank = doc!.createElementNS(W_NS, 'w:tc');
  const tcPr = directChildren(template).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr');
  if (tcPr) blank.appendChild(tcPr.cloneNode(true));
  const paragraph = doc!.createElementNS(W_NS, 'w:p');
  const sourceParagraph = directChildren(template).find((child) => child.namespaceURI === W_NS && child.localName === 'p');
  const pPr = sourceParagraph && directChildren(sourceParagraph).find((child) => child.namespaceURI === W_NS && child.localName === 'pPr');
  if (pPr) paragraph.appendChild(pPr.cloneNode(true));
  paragraph.appendChild(doc!.createElementNS(W_NS, 'w:r'));
  blank.appendChild(paragraph);
  return blank;
}

function appendCellContent(target: Element, source: Element): void {
  for (const child of directChildren(source)) {
    if (child.namespaceURI === W_NS && child.localName === 'tcPr') continue;
    target.appendChild(child.cloneNode(true));
  }
}

function cellGridSpan(cell: Element): number {
  const value = directChildren(cell)
    .find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr')
    ?.getElementsByTagNameNS(W_NS, 'gridSpan')[0]
    ?.getAttributeNS(W_NS, 'val');
  const span = Number(value);
  return Number.isFinite(span) && span > 0 ? span : 1;
}

function cellHasVerticalMerge(cell: Element): boolean {
  return directChildren(cell)
    .find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr')
    ?.getElementsByTagNameNS(W_NS, 'vMerge').length
    ? true
    : false;
}

function setCellGridSpan(cell: Element, span: number): void {
  const tcPr = ensureDirectChild(cell, 'tcPr');
  let gridSpan = directChildren(tcPr).find((child) => child.namespaceURI === W_NS && child.localName === 'gridSpan');
  if (!gridSpan) {
    gridSpan = cell.ownerDocument!.createElementNS(W_NS, 'w:gridSpan');
    tcPr.appendChild(gridSpan);
  }
  gridSpan.setAttributeNS(W_NS, 'w:val', String(span));
}

function removeCellGridSpan(cell: Element): void {
  const tcPr = directChildren(cell).find((child) => child.namespaceURI === W_NS && child.localName === 'tcPr');
  const gridSpan = tcPr && directChildren(tcPr).find((child) => child.namespaceURI === W_NS && child.localName === 'gridSpan');
  if (gridSpan) removeNode(gridSpan);
}

function ensureDirectChild(element: Element, localName: string): Element {
  const existing = directChildren(element).find((child) => child.namespaceURI === W_NS && child.localName === localName);
  if (existing) return existing;
  const created = element.ownerDocument!.createElementNS(W_NS, `w:${localName}`);
  element.insertBefore(created, element.firstChild);
  return created;
}

function updateTableGrid(table: Element, operation: 'insert' | 'delete', columnIndex: number): void {
  const grid = directChildren(table).find((child) => child.namespaceURI === W_NS && child.localName === 'tblGrid');
  if (!grid) return;
  const columns = directChildren(grid).filter((child) => child.namespaceURI === W_NS && child.localName === 'gridCol');
  if (operation === 'delete') {
    if (columns.length > 1 && columns[columnIndex]) removeNode(columns[columnIndex]!);
    return;
  }
  const template = columns[Math.min(columnIndex, columns.length - 1)];
  if (!template) return;
  const inserted = template.cloneNode(true) as Element;
  grid.insertBefore(inserted, columns[columnIndex] ?? null);
}

function partPathForNode(pkg: Package, node: Node): string | null {
  return partForNode(pkg, node)?.path ?? null;
}

function partForNode(pkg: Package, node: Node): Part | null {
  const ownerDocument = node.nodeType === 9 ? node : node.ownerDocument;
  if (!ownerDocument) return null;
  for (const part of pkg.parts.values()) {
    if (part.xmlDom === ownerDocument) return part;
  }
  return null;
}

function applyParagraphTextPatch(
  paragraph: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'paragraph-text' }>,
): { ok: true } | { ok: false; reason: string } {
  if (patch.op === 'insert') {
    return insertParagraphText(paragraph, patch.offset, patch.text);
  }
  if (patch.op === 'delete') {
    return replaceParagraphTextRange(paragraph, patch.offset, patch.text.length, '');
  }
  return replaceParagraphTextRange(paragraph, patch.offset, patch.deletedText?.length ?? 0, patch.text);
}

function insertParagraphText(
  paragraph: Element,
  offset: number,
  text: string,
): { ok: true } | { ok: false; reason: string } {
  if (text.length === 0) return { ok: true };
  const units = collectParagraphOffsetUnits(paragraph);
  const totalLength = units.reduce((sum, unit) => sum + unit.length, 0);
  if (hasFieldMarkup(paragraph) && offset === totalLength) {
    // Field result runs are cached display values, not editable text. Appending
    // at the logical paragraph end must create a run after the complete field.
    appendTextRun(paragraph, text);
    return { ok: true };
  }
  if (units.every((unit) => unit.kind !== 'text')) {
    if (offset !== 0) return { ok: false, reason: `insert offset out of range: ${offset}` };
    appendTextRun(paragraph, text);
    return { ok: true };
  }
  const location = locateTextNode(units, offset, true);
  if (!location) return { ok: false, reason: `insert offset out of range: ${offset}` };
  if (location.node.localName === 'delText' && location.localOffset === (location.node.textContent?.length ?? 0)) {
    const wrapper = containingRevisionWrapper(location.node, 'del');
    const paragraphParent = wrapper?.parentNode;
    if (wrapper && paragraphParent === paragraph) {
      const doc = paragraph.ownerDocument;
      if (!doc) return { ok: false, reason: 'paragraph has no owner document' };
      const run = doc.createElementNS(W_NS, 'w:r');
      const textNode = doc.createElementNS(W_NS, 'w:t');
      textNode.textContent = text;
      preserveTextSpace(textNode);
      run.appendChild(textNode);
      paragraph.insertBefore(run, wrapper.nextSibling);
      return { ok: true };
    }
  }
  const current = location.node.textContent ?? '';
  location.node.textContent = `${current.slice(0, location.localOffset)}${text}${current.slice(location.localOffset)}`;
  preserveTextSpace(location.node);
  return { ok: true };
}

function replaceParagraphTextRange(
  paragraph: Element,
  offset: number,
  length: number,
  text: string,
): { ok: true } | { ok: false; reason: string } {
  if (length === 0) return insertParagraphText(paragraph, offset, text);
  const units = collectParagraphOffsetUnits(paragraph);
  const textNodes = units.filter((unit): unit is ParagraphTextOffsetUnit => unit.kind === 'text').map((unit) => unit.node);
  const start = locateTextNode(units, offset, false);
  const end = locateTextNode(units, offset + length, true);
  if (!start || !end) return { ok: false, reason: `text range out of range: ${offset}:${length}` };
  const current = start.node.textContent ?? '';
  if (start.node === end.node) {
    start.node.textContent = `${current.slice(0, start.localOffset)}${text}${current.slice(end.localOffset)}`;
    preserveTextSpace(start.node);
    return { ok: true };
  }

  const startIndex = textNodes.indexOf(start.node);
  const endIndex = textNodes.indexOf(end.node);
  if (startIndex < 0 || endIndex < 0 || startIndex > endIndex) {
    return { ok: false, reason: `text range out of range: ${offset}:${length}` };
  }
  const endText = end.node.textContent ?? '';
  start.node.textContent = `${current.slice(0, start.localOffset)}${text}`;
  preserveTextSpace(start.node);
  for (let i = startIndex + 1; i < endIndex; i++) {
    textNodes[i]!.textContent = '';
  }
  end.node.textContent = endText.slice(end.localOffset);
  preserveTextSpace(end.node);
  return { ok: true };
}

interface ParagraphTextOffsetUnit {
  kind: 'text';
  node: Element;
  length: number;
}

interface ParagraphObjectOffsetUnit {
  kind: 'object';
  node: Element;
  length: 1;
}

type ParagraphOffsetUnit = ParagraphTextOffsetUnit | ParagraphObjectOffsetUnit;

function collectParagraphOffsetUnits(paragraph: Element): ParagraphOffsetUnit[] {
  const out: ParagraphOffsetUnit[] = [];
  collectParagraphOffsetUnitsFromNode(paragraph, out);
  return out;
}

function hasFieldMarkup(paragraph: Element): boolean {
  return paragraph.getElementsByTagNameNS(W_NS, 'fldChar').length > 0 ||
    paragraph.getElementsByTagNameNS(W_NS, 'instrText').length > 0;
}

function collectTextNodes(paragraph: Element): Element[] {
  return collectParagraphOffsetUnits(paragraph)
    .filter((unit): unit is ParagraphTextOffsetUnit => unit.kind === 'text')
    .map((unit) => unit.node);
}

function collectParagraphOffsetUnitsFromNode(node: Node, out: ParagraphOffsetUnit[]): void {
  if (node.nodeType !== 1) return;
  const element = node as Element;
  if (element.namespaceURI === W_NS && element.localName === 't') {
    out.push({ kind: 'text', node: element, length: element.textContent?.length ?? 0 });
    return;
  }
  if (element.namespaceURI === W_NS && element.localName === 'delText') {
    out.push({ kind: 'text', node: element, length: element.textContent?.length ?? 0 });
    return;
  }
  if (
    (element.namespaceURI === W_NS && element.localName === 'drawing') ||
    (element.namespaceURI === W_NS && element.localName === 'pict')
  ) {
    out.push({ kind: 'object', node: element, length: 1 });
    return;
  }
  for (let index = 0; index < element.childNodes.length; index += 1) {
    const child = element.childNodes[index];
    if (child) collectParagraphOffsetUnitsFromNode(child, out);
  }
}

function containingRevisionWrapper(node: Node, localName: 'ins' | 'del'): Element | null {
  let current: Node | null = node.parentNode;
  while (current && current.nodeType === 1) {
    const element = current as Element;
    if (element.namespaceURI === W_NS && element.localName === localName) return element;
    if (element.namespaceURI === W_NS && element.localName === 'p') return null;
    current = current.parentNode;
  }
  return null;
}

function findDrawingAtParagraphOffset(paragraph: Element, offset: number): Element | null {
  let position = 0;
  const units = collectParagraphOffsetUnits(paragraph);
  for (const unit of units) {
    if (unit.kind === 'object' && offset >= position && offset <= position + unit.length) return unit.node;
    position += unit.length;
  }
  return null;
}

function locateTextNode(
  units: readonly ParagraphOffsetUnit[],
  offset: number,
  allowEnd: boolean,
): { node: Element; localOffset: number } | null {
  let position = 0;
  for (let index = 0; index < units.length; index += 1) {
    const unit = units[index]!;
    const end = position + unit.length;
    if (unit.kind === 'text') {
      if (offset >= position && (offset < end || (allowEnd && offset === end))) {
        return { node: unit.node, localOffset: offset - position };
      }
    } else if (offset >= position && offset <= end) {
      return nextTextUnit(units, index) ?? previousTextUnit(units, index);
    }
    position = end;
  }
  return null;
}

function nextTextUnit(units: readonly ParagraphOffsetUnit[], fromIndex: number): { node: Element; localOffset: number } | null {
  for (let index = fromIndex + 1; index < units.length; index += 1) {
    const unit = units[index]!;
    if (unit.kind === 'text') return { node: unit.node, localOffset: 0 };
  }
  return null;
}

function previousTextUnit(units: readonly ParagraphOffsetUnit[], fromIndex: number): { node: Element; localOffset: number } | null {
  for (let index = fromIndex - 1; index >= 0; index -= 1) {
    const unit = units[index]!;
    if (unit.kind === 'text') return { node: unit.node, localOffset: unit.length };
  }
  return null;
}

function appendTextRun(paragraph: Element, text: string): void {
  const doc = paragraph.ownerDocument;
  if (!doc) return;
  const run = doc.createElementNS(W_NS, 'w:r');
  const textEl = doc.createElementNS(W_NS, 'w:t');
  textEl.textContent = text;
  preserveTextSpace(textEl);
  run.appendChild(textEl);
  paragraph.appendChild(run);
}

function preserveTextSpace(textEl: Element): void {
  const text = textEl.textContent ?? '';
  if (/^\s|\s$|\s{2,}/.test(text)) {
    textEl.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  }
}

function splitParagraphElement(
  paragraph: Element,
  offset: number,
  newParaId: string,
): { ok: true } | { ok: false; reason: string } {
  const parent = paragraph.parentNode;
  const doc = paragraph.ownerDocument;
  if (!parent || !doc) return { ok: false, reason: 'paragraph has no parent document' };

  const contentChildren = directChildren(paragraph).filter((child) => !(child.namespaceURI === W_NS && child.localName === 'pPr'));
  const totalLength = contentChildren.reduce((sum, child) => sum + inlineNodeLength(child), 0);
  if (offset < 0 || offset > totalLength) return { ok: false, reason: `split offset out of range: ${offset}` };

  const afterChildren = splitParagraphContentChildren(doc, contentChildren, offset);
  const newParagraph = doc.createElementNS(W_NS, 'w:p');
  copyParagraphAttributes(paragraph, newParagraph, newParaId);
  const pPr = directChildren(paragraph).find((child) => child.namespaceURI === W_NS && child.localName === 'pPr');
  if (pPr) newParagraph.appendChild(pPr.cloneNode(true));
  for (const child of afterChildren) {
    newParagraph.appendChild(child);
  }
  parent.insertBefore(newParagraph, paragraph.nextSibling);
  return { ok: true };
}

function mergeNextParagraphElement(paragraph: Element): { ok: true } | { ok: false; reason: string } {
  const next = nextSiblingParagraph(paragraph);
  if (!next) return { ok: false, reason: 'next paragraph not found for merge' };
  for (const child of directChildren(next)) {
    if (child.namespaceURI === W_NS && child.localName === 'pPr') continue;
    paragraph.appendChild(child.cloneNode(true));
  }
  removeNode(next);
  return { ok: true };
}

function splitParagraphContentChildren(doc: Document, children: Element[], offset: number): Element[] {
  const after: Element[] = [];
  let position = 0;
  let split = false;
  for (const child of children) {
    const length = inlineNodeLength(child);
    const end = position + length;
    if (split) {
      after.push(child.cloneNode(true) as Element);
      removeNode(child);
    } else if (offset <= position) {
      after.push(child.cloneNode(true) as Element);
      removeNode(child);
    } else if (offset >= end) {
      // Keep child in the original paragraph.
    } else {
      const splitChild = splitElementTextAt(doc, child, offset - position);
      if (splitChild.after) after.push(splitChild.after);
      split = true;
    }
    if (!split && offset === end) split = true;
    position = end;
  }
  return after;
}

function splitElementTextAt(
  _doc: Document,
  element: Element,
  offset: number,
): { after: Element | null } {
  const beforeClone = element.cloneNode(true) as Element;
  const afterClone = element.cloneNode(true) as Element;
  applyTextSlice(beforeClone, 0, offset);
  applyTextSlice(afterClone, offset, inlineNodeLength(element));
  replaceNode(element, beforeClone);
  return hasTextContent(afterClone) || inlineNodeLength(afterClone) > 0 ? { after: afterClone } : { after: null };
}

function applyTextSlice(element: Element, from: number, to: number): void {
  const textNodes = collectTextNodes(element);
  let position = 0;
  for (const textNode of textNodes) {
    const text = textNode.textContent ?? '';
    const end = position + text.length;
    const localFrom = Math.max(0, from - position);
    const localTo = Math.min(text.length, to - position);
    textNode.textContent = localFrom < localTo ? text.slice(localFrom, localTo) : '';
    preserveTextSpace(textNode);
    position = end;
  }
}

function inlineNodeLength(element: Element): number {
  if (element.namespaceURI === W_NS && element.localName === 'drawing') return 1;
  let length = 0;
  const textNodes = element.getElementsByTagNameNS(W_NS, 't');
  for (let i = 0; i < textNodes.length; i++) {
    length += textNodes[i]?.textContent?.length ?? 0;
  }
  const drawings = element.getElementsByTagNameNS(W_NS, 'drawing');
  length += drawings.length;
  return length;
}

function copyParagraphAttributes(source: Element, target: Element, paraId: string): void {
  for (let i = 0; i < source.attributes.length; i++) {
    const attr = source.attributes[i];
    if (!attr) continue;
    if (attr.namespaceURI === W14_NS && attr.localName === 'paraId') continue;
    target.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
  }
  target.setAttributeNS(W14_NS, 'w14:paraId', paraId);
}

function directChildren(element: Element): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < element.childNodes.length; i++) {
    const child = element.childNodes[i];
    if (child?.nodeType === 1) out.push(child as Element);
  }
  return out;
}

function nextSiblingParagraph(paragraph: Element): Element | null {
  let current = paragraph.nextSibling;
  while (current) {
    if (current.nodeType === 1) {
      const el = current as Element;
      if (el.namespaceURI === W_NS && el.localName === 'p') return el;
      return null;
    }
    current = current.nextSibling;
  }
  return null;
}

function hasTextContent(element: Element): boolean {
  return collectTextNodes(element).some((node) => (node.textContent ?? '').length > 0);
}

function removeNode(node: Node): void {
  node.parentNode?.removeChild(node);
}

function replaceNode(node: Node, replacement: Node): void {
  node.parentNode?.replaceChild(replacement, node);
}

function removeDrawing(drawing: Element): void {
  const parent = drawing.parentNode;
  if (!parent) return;
  if (isElement(parent) && parent.namespaceURI === W_NS && parent.localName === 'r' && isDrawingOnlyRun(parent)) {
    parent.parentNode?.removeChild(parent);
    return;
  }
  parent.removeChild(drawing);
}

function resizeDrawing(drawing: Element, widthPx: number, heightPx: number): { ok: true } | { ok: false; reason: string } {
  if (!Number.isFinite(widthPx) || !Number.isFinite(heightPx) || widthPx <= 0 || heightPx <= 0) {
    return { ok: false, reason: 'image size must be positive finite pixels' };
  }
  const widthEmu = Math.round((widthPx / 96) * 914400);
  const heightEmu = Math.round((heightPx / 96) * 914400);
  const container = directChildren(drawing).find((child) => child.namespaceURI === WP_NS && (child.localName === 'inline' || child.localName === 'anchor'));
  if (!container) return { ok: false, reason: 'drawing has no wp:inline or wp:anchor container' };
  const extent = directChildren(container).find((child) => child.namespaceURI === WP_NS && child.localName === 'extent');
  if (!extent) return { ok: false, reason: 'drawing has no wp:extent' };
  extent.setAttribute('cx', String(widthEmu));
  extent.setAttribute('cy', String(heightEmu));
  const transforms = container.getElementsByTagNameNS(A_NS, 'xfrm');
  for (let index = 0; index < transforms.length; index += 1) {
    const transform = transforms[index];
    const ext = transform && directChildren(transform).find((child) => child.namespaceURI === A_NS && child.localName === 'ext');
    if (!ext) continue;
    ext.setAttribute('cx', String(widthEmu));
    ext.setAttribute('cy', String(heightEmu));
  }
  return { ok: true };
}

function insertDrawingIntoParagraph(
  pkg: Package,
  partPath: string,
  paragraph: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'drawing-insert' }>,
): { ok: true } | { ok: false; reason: string } {
  const image = parseImageDataUrl(patch.src);
  if (!image) return { ok: false, reason: 'image insert requires a base64 data:image URL' };
  const relationship = ensureImageRelationship(pkg, partPath, image.bytes, image.extension, image.contentType);
  const drawing = createInlinePicture(paragraph.ownerDocument!, relationship.id, patch.drawingId, patch.width, patch.height);
  const children = directChildren(paragraph).filter((child) => !(child.namespaceURI === W_NS && child.localName === 'pPr'));
  const totalLength = children.reduce((sum, child) => sum + inlineNodeLength(child), 0);
  if (patch.offset < 0 || patch.offset > totalLength) return { ok: false, reason: `image insert offset out of range: ${patch.offset}` };
  const next: Element[] = [];
  let position = 0;
  let inserted = false;
  for (const child of children) {
    const length = inlineNodeLength(child);
    if (!inserted && patch.offset <= position) {
      next.push(drawing);
      inserted = true;
    }
    if (!inserted && patch.offset < position + length && child.namespaceURI === W_NS && child.localName === 'r') {
      const before = sliceInlineElement(child, 0, patch.offset - position);
      const after = sliceInlineElement(child, patch.offset - position, length);
      if (before) next.push(before);
      next.push(drawing);
      inserted = true;
      if (after) next.push(after);
    } else {
      next.push(child);
    }
    position += length;
  }
  if (!inserted) next.push(drawing);
  for (const child of children) removeNode(child);
  for (const child of next) paragraph.appendChild(child);
  markDirty(relationship.part);
  markDirty(relationship.mediaPart);
  markContentTypeDirty(pkg, image.extension, image.contentType);
  return { ok: true };
}

function createInlinePicture(document: Document, relationshipId: string, drawingId: string, widthPx: number, heightPx: number): Element {
  const widthEmu = Math.round((widthPx / 96) * 914400);
  const heightEmu = Math.round((heightPx / 96) * 914400);
  const run = document.createElementNS(W_NS, 'w:r');
  const drawing = document.createElementNS(W_NS, 'w:drawing');
  const inline = document.createElementNS(WP_NS, 'wp:inline');
  const extent = document.createElementNS(WP_NS, 'wp:extent');
  extent.setAttribute('cx', String(widthEmu));
  extent.setAttribute('cy', String(heightEmu));
  inline.appendChild(extent);
  const docPr = document.createElementNS(WP_NS, 'wp:docPr');
  docPr.setAttribute('id', String(Math.abs(hashString(drawingId)) + 1));
  docPr.setAttribute('name', `Picture ${drawingId}`);
  inline.appendChild(docPr);
  inline.appendChild(document.createElementNS(WP_NS, 'wp:cNvGraphicFramePr'));
  const graphic = document.createElementNS(A_NS, 'a:graphic');
  const graphicData = document.createElementNS(A_NS, 'a:graphicData');
  graphicData.setAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
  const picture = document.createElementNS(PIC_NS, 'pic:pic');
  const nvPicPr = document.createElementNS(PIC_NS, 'pic:nvPicPr');
  const cNvPr = document.createElementNS(PIC_NS, 'pic:cNvPr');
  cNvPr.setAttribute('id', '0');
  cNvPr.setAttribute('name', `image-${drawingId}`);
  nvPicPr.appendChild(cNvPr);
  nvPicPr.appendChild(document.createElementNS(PIC_NS, 'pic:cNvPicPr'));
  picture.appendChild(nvPicPr);
  const blipFill = document.createElementNS(PIC_NS, 'pic:blipFill');
  const blip = document.createElementNS(A_NS, 'a:blip');
  blip.setAttributeNS(R_NS, 'r:embed', relationshipId);
  blipFill.appendChild(blip);
  const stretch = document.createElementNS(A_NS, 'a:stretch');
  stretch.appendChild(document.createElementNS(A_NS, 'a:fillRect'));
  blipFill.appendChild(stretch);
  picture.appendChild(blipFill);
  const spPr = document.createElementNS(PIC_NS, 'pic:spPr');
  const transform = document.createElementNS(A_NS, 'a:xfrm');
  const transformExtent = document.createElementNS(A_NS, 'a:ext');
  transformExtent.setAttribute('cx', String(widthEmu));
  transformExtent.setAttribute('cy', String(heightEmu));
  transform.appendChild(document.createElementNS(A_NS, 'a:off'));
  transform.appendChild(transformExtent);
  spPr.appendChild(transform);
  const geometry = document.createElementNS(A_NS, 'a:prstGeom');
  geometry.setAttribute('prst', 'rect');
  geometry.appendChild(document.createElementNS(A_NS, 'a:avLst'));
  spPr.appendChild(geometry);
  picture.appendChild(spPr);
  graphicData.appendChild(picture);
  graphic.appendChild(graphicData);
  inline.appendChild(graphic);
  drawing.appendChild(inline);
  run.appendChild(drawing);
  return run;
}

function parseImageDataUrl(src: string): { bytes: Uint8Array; extension: string; contentType: string } | null {
  const match = /^data:(image\/(png|jpeg|jpg|gif|bmp|webp));base64,([a-z0-9+/=\s]+)$/i.exec(src.trim());
  if (!match) return null;
  const contentType = match[1]!.toLowerCase() === 'image/jpg' ? 'image/jpeg' : match[1]!.toLowerCase();
  const extension = contentType === 'image/jpeg' ? 'jpeg' : contentType.slice('image/'.length);
  const binary = atob(match[3]!.replace(/\s/g, ''));
  return { bytes: Uint8Array.from(binary, (char) => char.charCodeAt(0)), extension, contentType };
}

function ensureImageRelationship(
  pkg: Package,
  partPath: string,
  bytes: Uint8Array,
  extension: string,
  contentType: string,
): { id: string; part: Part; mediaPart: Part } {
  const relsPath = relsPathFor(partPath);
  let part = pkg.parts.get(relsPath);
  if (!part?.xmlDom) {
    const sourceDocument = Array.from(pkg.parts.values()).find((candidate) => candidate.xmlDom)?.xmlDom;
    if (!sourceDocument) throw new Error('cannot create image relationships without an XML part');
    const dom = sourceDocument.implementation?.createDocument(PACKAGE_RELS_NS, 'Relationships', null) as unknown as Document;
    if (!dom) throw new Error('XML implementation cannot create relationship document');
    part = createXmlPart({ path: relsPath, contentType: RELS_CONTENT_TYPE, xmlDom: dom });
    setPart(pkg, part);
  }
  const relsDom = part.xmlDom;
  if (!relsDom) throw new Error(`relationship part has no XML DOM: ${relsPath}`);
  const relationships = relsDom.getElementsByTagNameNS(PACKAGE_RELS_NS, 'Relationship');
  let idNumber = 1;
  const usedIds = new Set<string>();
  for (let index = 0; index < relationships.length; index += 1) {
    const id = relationships[index]?.getAttribute('Id');
    if (id) usedIds.add(id);
  }
  while (usedIds.has(`rId${idNumber}`)) idNumber += 1;
  let mediaNumber = 1;
  while (pkg.parts.has(`word/media/image${mediaNumber}.${extension}`)) mediaNumber += 1;
  const mediaPath = `word/media/image${mediaNumber}.${extension}`;
  const mediaPart = createBinaryPart({ path: mediaPath, contentType, bytes });
  setPart(pkg, mediaPart);
  const relationship = relsDom.createElementNS(PACKAGE_RELS_NS, 'Relationship');
  relationship.setAttribute('Id', `rId${idNumber}`);
  relationship.setAttribute('Type', `${R_NS}/image`);
  relationship.setAttribute('Target', `media/image${mediaNumber}.${extension}`);
  relsDom.documentElement?.appendChild(relationship);
  return { id: `rId${idNumber}`, part, mediaPart };
}

function markContentTypeDirty(pkg: Package, extension: string, contentType: string): void {
  const part = getPart(pkg, '[Content_Types].xml');
  const dom = part?.xmlDom;
  if (!part || !dom) return;
  const defaults = dom.getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Default');
  for (let index = 0; index < defaults.length; index += 1) {
    if (defaults[index]?.getAttribute('Extension')?.toLowerCase() === extension.toLowerCase()) return;
  }
  const created = dom.createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Default');
  created.setAttribute('Extension', extension);
  created.setAttribute('ContentType', contentType);
  dom.documentElement?.appendChild(created);
  markDirty(part);
}

function hashString(value: string): number {
  let hash = 0;
  for (let index = 0; index < value.length; index += 1) hash = (hash * 31 + value.charCodeAt(index)) | 0;
  return hash;
}

function applyRevisionPatch(
  paragraph: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'revision' }>,
): { ok: true } | { ok: false; reason: string } {
  if (patch.op === 'accept' || patch.op === 'reject') {
    const localName = patch.revisionKind === 'insert' ? 'ins' : 'del';
    const wrappers = Array.from(paragraph.getElementsByTagNameNS(W_NS, localName));
    const matched = wrappers.filter((wrapper) => wrapper.getAttributeNS(W_NS, 'id') === patch.revisionId);
    if (matched.length === 0) return { ok: false, reason: `revision not found: ${patch.revisionId}` };
    for (const wrapper of matched) {
      const parent = wrapper.parentNode;
      if (!parent) continue;
      const keep = (localName === 'ins' && patch.op === 'accept') || (localName === 'del' && patch.op === 'reject');
      if (!keep) {
        parent.removeChild(wrapper);
        continue;
      }
      const children = Array.from(wrapper.childNodes);
      if (localName === 'del') {
        for (const child of children) {
          if (child.nodeType === 1) convertRevisionDeleteText(child as Element);
        }
      }
      for (const child of children) parent.insertBefore(child, wrapper);
      parent.removeChild(wrapper);
    }
    return { ok: true };
  }

  if (patch.from < 0 || patch.to <= patch.from) return { ok: false, reason: 'revision range is empty or invalid' };
  let position = 0;
  let selected: Element[] = [];
  let wrapperIndex = -1;
  const rebuilt: Node[] = [];
  for (const child of Array.from(paragraph.childNodes)) {
    if (child.nodeType !== 1 || (child as Element).namespaceURI !== W_NS) {
      rebuilt.push(child);
      continue;
    }
    const run = child as Element;
    if (run.localName === 'ins' || run.localName === 'del') {
      rebuilt.push(run);
      position += run.textContent?.length ?? 0;
      continue;
    }
    if (run.localName !== 'r') {
      rebuilt.push(run);
      continue;
    }
    const texts = Array.from(run.getElementsByTagNameNS(W_NS, 't'));
    if (texts.length !== 1) return { ok: false, reason: 'revision writeback requires runs with one text node' };
    const text = texts[0]!.textContent ?? '';
    const runStart = position;
    const runEnd = position + text.length;
    position = runEnd;
    if (runEnd <= patch.from || runStart >= patch.to) {
      rebuilt.push(run);
      continue;
    }
    const localFrom = Math.max(0, patch.from - runStart);
    const localTo = Math.min(text.length, patch.to - runStart);
    if (localFrom > 0) rebuilt.push(cloneRevisionRun(run, text.slice(0, localFrom)));
    if (localTo > localFrom) {
      if (wrapperIndex < 0) wrapperIndex = rebuilt.length;
      selected.push(cloneRevisionRun(run, text.slice(localFrom, localTo)));
    }
    if (localTo < text.length) rebuilt.push(cloneRevisionRun(run, text.slice(localTo)));
  }
  if (selected.length === 0 || patch.to > position) return { ok: false, reason: `revision range out of range: ${patch.from}:${patch.to}` };
  const doc = paragraph.ownerDocument;
  if (!doc) return { ok: false, reason: 'revision paragraph has no owner document' };
  const wrapper = doc.createElementNS(W_NS, patch.revisionKind === 'insert' ? 'w:ins' : 'w:del');
  wrapper.setAttributeNS(W_NS, 'w:id', patch.revisionId);
  wrapper.setAttributeNS(W_NS, 'w:author', 'User');
  wrapper.setAttributeNS(W_NS, 'w:date', new Date().toISOString());
  for (const run of selected) {
    if (patch.revisionKind === 'delete') convertRevisionText(run);
    wrapper.appendChild(run);
  }
  rebuilt.splice(wrapperIndex < 0 ? rebuilt.length : wrapperIndex, 0, wrapper);
  mergeAdjacentRevisionWrappers(rebuilt);
  while (paragraph.firstChild) paragraph.removeChild(paragraph.firstChild);
  for (const child of rebuilt) paragraph.appendChild(child);
  return { ok: true };
}

function applyParagraphFormatPatch(
  paragraph: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'paragraph-format' }>,
): { ok: true } | { ok: false; reason: string } {
  if (patch.from < 0 || patch.to <= patch.from) return { ok: false, reason: 'format range is empty or invalid' };
  const hasRevisionWrappers =
    paragraph.getElementsByTagNameNS(W_NS, 'ins').length > 0 ||
    paragraph.getElementsByTagNameNS(W_NS, 'del').length > 0;
  if (hasRevisionWrappers) {
    // Keep tracked-content containers intact while formatting their descendant runs.
    // Partial runs are deliberately rejected here; splitting them would require
    // rebuilding the surrounding w:ins/w:del structure and risks losing tracking.
    const runs = Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'));
    let position = 0;
    let selected = 0;
    for (const run of runs) {
      const text = wordRunText(run);
      if (text.length === 0) continue;
      const start = position;
      const end = position + text.length;
      position = end;
      if (end <= patch.from || start >= patch.to) continue;
      if (patch.from > start || patch.to < end) {
        return { ok: false, reason: 'format writeback requires whole runs inside tracked content' };
      }
      applyRunFormat(run, patch.format);
      selected += text.length;
    }
    if (selected === 0 || patch.to > position) return { ok: false, reason: `format range out of range: ${patch.from}:${patch.to}` };
    return { ok: true };
  }
  const rebuilt: Node[] = [];
  let position = 0;
  let selected = 0;
  for (const child of Array.from(paragraph.childNodes)) {
    if (child.nodeType !== 1 || (child as Element).namespaceURI !== W_NS || (child as Element).localName !== 'r') {
      rebuilt.push(child);
      continue;
    }
    const run = child as Element;
    const texts = Array.from(run.getElementsByTagNameNS(W_NS, 't'));
    if (texts.length !== 1) return { ok: false, reason: 'format writeback requires runs with one text node' };
    const text = texts[0]!.textContent ?? '';
    const start = position;
    const end = position + text.length;
    position = end;
    if (end <= patch.from || start >= patch.to) {
      rebuilt.push(run);
      continue;
    }
    const localFrom = Math.max(0, patch.from - start);
    const localTo = Math.min(text.length, patch.to - start);
    if (localFrom > 0) rebuilt.push(cloneRevisionRun(run, text.slice(0, localFrom)));
    if (localTo > localFrom) {
      const formatted = cloneRevisionRun(run, text.slice(localFrom, localTo));
      applyRunFormat(formatted, patch.format);
      rebuilt.push(formatted);
      selected += localTo - localFrom;
    }
    if (localTo < text.length) rebuilt.push(cloneRevisionRun(run, text.slice(localTo)));
  }
  if (selected === 0 || patch.to > position) return { ok: false, reason: `format range out of range: ${patch.from}:${patch.to}` };
  while (paragraph.firstChild) paragraph.removeChild(paragraph.firstChild);
  for (const child of rebuilt) paragraph.appendChild(child);
  return { ok: true };
}

function applyFormatRevisionPatch(
  paragraph: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'format-revision' }>,
): { ok: true } | { ok: false; reason: string } {
  const runs = Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'));
  if (patch.op === 'run') {
    if (patch.beforeFormat === undefined) return { ok: false, reason: 'format revision is missing beforeFormat snapshot' };
    let position = 0;
    let changed = 0;
    for (const run of runs) {
      const text = wordRunText(run);
      const start = position;
      const end = position + text.length;
      position = end;
      if (end <= patch.from || start >= patch.to) continue;
      if (patch.from > start || patch.to < end) return { ok: false, reason: 'format revision currently requires whole runs' };
      const rPr = ensureDirectRPr(run);
      removeDirectChildren(rPr, 'rPrChange');
      const change = rPr.ownerDocument!.createElementNS(W_NS, 'w:rPrChange');
      change.setAttributeNS(W_NS, 'w:id', patch.revisionId);
      change.setAttributeNS(W_NS, 'w:author', 'User');
      change.setAttributeNS(W_NS, 'w:date', new Date().toISOString());
      const oldRun = rPr.ownerDocument!.createElementNS(W_NS, 'w:r');
      applyRunFormat(oldRun, patch.beforeFormat ?? {});
      const oldRPr = directChildren(oldRun).find((child) => child.namespaceURI === W_NS && child.localName === 'rPr');
      if (oldRPr) for (const child of Array.from(oldRPr.childNodes)) change.appendChild(child);
      rPr.appendChild(change);
      changed += text.length;
    }
    if (changed === 0 || patch.to > position) return { ok: false, reason: `format revision range not found: ${patch.from}:${patch.to}` };
    return { ok: true };
  }

  let changed = 0;
  for (const run of runs) {
    const rPr = directChildren(run).find((child) => child.namespaceURI === W_NS && child.localName === 'rPr');
    const change = rPr && directChildren(rPr).find(
      (child) => child.namespaceURI === W_NS && child.localName === 'rPrChange' && child.getAttributeNS(W_NS, 'id') === patch.revisionId,
    );
    if (!change || !rPr) continue;
    if (patch.op === 'reject') {
      for (const child of Array.from(rPr.childNodes)) {
        if (child !== change) rPr.removeChild(child);
      }
      for (const child of Array.from(change.childNodes)) rPr.appendChild(child.cloneNode(true));
    }
    rPr.removeChild(change);
    changed += 1;
  }
  return changed > 0 ? { ok: true } : { ok: false, reason: `format revision not found: ${patch.revisionId}` };
}

function wordRunText(run: Element): string {
  return Array.from(run.getElementsByTagNameNS(W_NS, 't'))
    .concat(Array.from(run.getElementsByTagNameNS(W_NS, 'delText')))
    .map((node) => node.textContent ?? '')
    .join('');
}

function ensureDirectRPr(run: Element): Element {
  const existing = directChildren(run).find((child) => child.namespaceURI === W_NS && child.localName === 'rPr');
  if (existing) return existing;
  const rPr = run.ownerDocument!.createElementNS(W_NS, 'w:rPr');
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}

function applyRunFormat(run: Element, format: Partial<EditorRunFormat>): void {
  const doc = run.ownerDocument!;
  let rPr = directChildren(run).find((child) => child.namespaceURI === W_NS && child.localName === 'rPr') ?? null;
  if (!rPr) {
    rPr = doc.createElementNS(W_NS, 'w:rPr');
    run.insertBefore(rPr, run.firstChild);
  }
  const booleanProperties: Array<[keyof EditorRunFormat, string]> = [
    ['bold', 'b'], ['italic', 'i'], ['strike', 'strike'], ['doubleStrike', 'dstrike'],
    ['caps', 'caps'], ['smallCaps', 'smallCaps'],
  ];
  for (const [key, localName] of booleanProperties) {
    if (format[key] === undefined) continue;
    removeDirectChildren(rPr, localName);
    if (format[key]) rPr.appendChild(doc.createElementNS(W_NS, `w:${localName}`));
  }
  if (format.underline !== undefined) {
    removeDirectChildren(rPr, 'u');
    if (format.underline) {
      const underline = doc.createElementNS(W_NS, 'w:u');
      underline.setAttribute('val', 'single');
      rPr.appendChild(underline);
    }
  }
  if (format.color !== undefined) {
    removeDirectChildren(rPr, 'color');
    if (format.color) {
      const color = doc.createElementNS(W_NS, 'w:color');
      color.setAttribute('val', format.color.replace(/^#/, '').toUpperCase());
      rPr.appendChild(color);
    }
  }
  if (format.highlight !== undefined || format.shading !== undefined) {
    removeDirectChildren(rPr, 'shd');
    const fill = format.shading?.fill ?? format.highlight;
    if (fill) {
      const shading = doc.createElementNS(W_NS, 'w:shd');
      shading.setAttribute('val', 'clear');
      shading.setAttribute('fill', fill.replace(/^#/, '').toUpperCase());
      rPr.appendChild(shading);
    }
  }
  if (format.fontSize !== undefined) {
    removeDirectChildren(rPr, 'sz');
    const size = doc.createElementNS(W_NS, 'w:sz');
    size.setAttribute('val', String(Math.round(format.fontSize * 2)));
    rPr.appendChild(size);
  }
  if (format.fontFamily !== undefined) {
    removeDirectChildren(rPr, 'rFonts');
    const fonts = doc.createElementNS(W_NS, 'w:rFonts');
    fonts.setAttribute('ascii', format.fontFamily);
    fonts.setAttribute('hAnsi', format.fontFamily);
    fonts.setAttribute('eastAsia', format.fontFamily);
    fonts.setAttribute('cs', format.fontFamily);
    rPr.appendChild(fonts);
  }
}

function removeDirectChildren(parent: Element, localName: string): void {
  for (const child of directChildren(parent)) {
    if (child.namespaceURI === W_NS && child.localName === localName) parent.removeChild(child);
  }
}

function mergeAdjacentRevisionWrappers(nodes: Node[]): void {
  for (let index = 1; index < nodes.length; index += 1) {
    const previous = nodes[index - 1];
    const current = nodes[index];
    if (previous?.nodeType !== 1 || current?.nodeType !== 1) continue;
    const previousElement = previous as Element;
    const currentElement = current as Element;
    if (
      previousElement.namespaceURI !== W_NS || currentElement.namespaceURI !== W_NS ||
      previousElement.localName !== currentElement.localName ||
      (previousElement.localName !== 'ins' && previousElement.localName !== 'del') ||
      previousElement.getAttributeNS(W_NS, 'id') !== currentElement.getAttributeNS(W_NS, 'id')
    ) continue;
    for (const child of Array.from(currentElement.childNodes)) previousElement.appendChild(child);
    nodes.splice(index, 1);
    index -= 1;
  }
}

function cloneRevisionRun(run: Element, text: string): Element {
  const clone = run.cloneNode(true) as Element;
  const textNode = clone.getElementsByTagNameNS(W_NS, 't')[0];
  if (textNode) textNode.textContent = text;
  return clone;
}

function convertRevisionText(run: Element): void {
  const doc = run.ownerDocument!;
  const texts = Array.from(run.getElementsByTagNameNS(W_NS, 't'));
  for (const text of texts) {
    const deleted = doc.createElementNS(W_NS, 'w:delText');
    for (let index = 0; index < text.attributes.length; index += 1) {
      const attribute = text.attributes[index];
      if (attribute) deleted.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    }
    deleted.textContent = text.textContent;
    text.parentNode?.replaceChild(deleted, text);
  }
}

function convertRevisionDeleteText(root: Element): void {
  const doc = root.ownerDocument!;
  const texts = Array.from(root.getElementsByTagNameNS(W_NS, 'delText'));
  for (const text of texts) {
    const normal = doc.createElementNS(W_NS, 'w:t');
    for (let index = 0; index < text.attributes.length; index += 1) {
      const attribute = text.attributes[index];
      if (attribute) normal.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    }
    normal.textContent = text.textContent;
    text.parentNode?.replaceChild(normal, text);
  }
}

function applyCommentPatch(
  pkg: Package,
  partPath: string,
  paragraph: Element,
  patch: Extract<WordOoxmlPatch, { kind: 'comment' }>,
): { ok: true } | { ok: false; reason: string } {
  if (patch.op === 'insert') {
    if (patch.from < 0 || patch.to <= patch.from) return { ok: false, reason: 'comment range is empty or invalid' };
    const totalLength = collectParagraphOffsetUnits(paragraph).reduce((sum, unit) => sum + unit.length, 0);
    if (patch.to > totalLength) return { ok: false, reason: `comment range out of range: ${patch.from}:${patch.to}` };
    ensureCommentsPart(pkg, partPath);
    insertCommentAnchors(paragraph, patch.from, patch.to, patch.commentId);
    const date = new Date().toISOString();
    appendCommentBody(pkg, patch.commentId, patch.body ?? '', 'User', date);
    if (patch.parentCommentId !== undefined) appendExtendedCommentRelation(pkg, patch.commentId, patch.parentCommentId);
    return { ok: true };
  }

  if (patch.op === 'update') {
    if (patch.body === undefined) return { ok: false, reason: 'comment update requires body' };
    const commentsPart = pkg.parts.get('word/comments.xml');
    const commentsDom = commentsPart?.xmlDom;
    if (!commentsPart || !commentsDom) return { ok: false, reason: 'comments.xml part not found' };
    const comments = commentsDom.getElementsByTagNameNS(W_NS, 'comment');
    for (let index = 0; index < comments.length; index += 1) {
      const comment = comments[index];
      if (comment?.getAttributeNS(W_NS, 'id') !== String(patch.commentId)) continue;
      const textNodes = comment.getElementsByTagNameNS(W_NS, 't');
      if (textNodes.length === 0) {
        const paragraph = commentsDom.createElementNS(W_NS, 'w:p');
        const run = commentsDom.createElementNS(W_NS, 'w:r');
        const text = commentsDom.createElementNS(W_NS, 'w:t');
        text.textContent = patch.body;
        run.appendChild(text);
        paragraph.appendChild(run);
        comment.appendChild(paragraph);
      } else {
        textNodes[0]!.textContent = patch.body;
        for (let textIndex = textNodes.length - 1; textIndex >= 1; textIndex -= 1) removeNode(textNodes[textIndex]!);
      }
      markDirty(commentsPart);
      return { ok: true };
    }
    return { ok: false, reason: `comment not found: ${patch.commentId}` };
  }

  removeCommentAnchors(pkg, patch.commentId);
  removeExtendedCommentRelation(pkg, patch.commentId);
  const commentsPart = pkg.parts.get('word/comments.xml');
  const commentsDom = commentsPart?.xmlDom;
  if (commentsDom) {
    const comments = commentsDom.getElementsByTagNameNS(W_NS, 'comment');
    for (let index = 0; index < comments.length; index += 1) {
      const comment = comments[index];
      if (comment?.getAttributeNS(W_NS, 'id') === String(patch.commentId)) removeNode(comment);
    }
    markDirty(commentsPart!);
  }
  return { ok: true };
}

function ensureCommentsPart(pkg: Package, partPath: string): Part {
  let commentsPart = pkg.parts.get('word/comments.xml');
  if (!commentsPart?.xmlDom) {
    const sourceDocument = Array.from(pkg.parts.values()).find((part) => part.xmlDom)?.xmlDom;
    if (!sourceDocument) throw new Error('cannot create comments.xml without an XML part');
    const dom = sourceDocument.implementation?.createDocument(W_NS, 'w:comments', null) as unknown as Document;
    if (!dom) throw new Error('XML implementation cannot create comments.xml');
    commentsPart = createXmlPart({ path: 'word/comments.xml', contentType: COMMENTS_CONTENT_TYPE, xmlDom: dom });
    setPart(pkg, commentsPart);
    ensureCommentsContentType(pkg);
  }
  ensureCommentsRelationship(pkg, partPath);
  return commentsPart;
}

function ensureCommentsContentType(pkg: Package): void {
  const contentTypes = getPart(pkg, '[Content_Types].xml');
  const dom = contentTypes?.xmlDom;
  if (!contentTypes || !dom) return;
  const overrides = dom.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
  for (let index = 0; index < overrides.length; index += 1) {
    if (overrides[index]?.getAttribute('PartName') === '/word/comments.xml') return;
  }
  const override = dom.createElementNS(CONTENT_TYPES_NS, 'Override');
  override.setAttribute('PartName', '/word/comments.xml');
  override.setAttribute('ContentType', COMMENTS_CONTENT_TYPE);
  dom.documentElement?.appendChild(override);
  markDirty(contentTypes);
}

function ensureCommentsRelationship(pkg: Package, partPath: string): void {
  const relsPath = relsPathFor(partPath);
  let relsPart = pkg.parts.get(relsPath);
  if (!relsPart?.xmlDom) return;
  const dom = relsPart.xmlDom;
  const relationships = dom.getElementsByTagNameNS(PACKAGE_RELS_NS, 'Relationship');
  for (let index = 0; index < relationships.length; index += 1) {
    if (relationships[index]?.getAttribute('Type') === `${R_NS}/comments`) return;
  }
  const used = new Set<string>();
  for (let index = 0; index < relationships.length; index += 1) {
    const id = relationships[index]?.getAttribute('Id');
    if (id) used.add(id);
  }
  let idNumber = 1;
  while (used.has(`rId${idNumber}`)) idNumber += 1;
  const relationship = dom.createElementNS(PACKAGE_RELS_NS, 'Relationship');
  relationship.setAttribute('Id', `rId${idNumber}`);
  relationship.setAttribute('Type', `${R_NS}/comments`);
  relationship.setAttribute('Target', 'comments.xml');
  dom.documentElement?.appendChild(relationship);
  markDirty(relsPart);
}

function appendCommentBody(pkg: Package, commentId: number, body: string, author: string, date: string): void {
  const part = pkg.parts.get('word/comments.xml');
  const dom = part?.xmlDom;
  if (!part || !dom) return;
  const root = dom.documentElement;
  if (!root) return;
  if (Array.from(dom.getElementsByTagNameNS(W_NS, 'comment'))
    .some((comment) => comment.getAttributeNS(W_NS, 'id') === String(commentId))) return;
  const comment = dom.createElementNS(W_NS, 'w:comment');
  comment.setAttributeNS(W_NS, 'w:id', String(commentId));
  comment.setAttributeNS(W_NS, 'w:author', author);
  comment.setAttributeNS(W_NS, 'w:date', date);
  comment.setAttributeNS(W_NS, 'w:initials', author.slice(0, 1));
  const paragraph = dom.createElementNS(W_NS, 'w:p');
  paragraph.setAttributeNS(W14_NS, 'w14:paraId', commentParaId(commentId));
  const run = dom.createElementNS(W_NS, 'w:r');
  const text = dom.createElementNS(W_NS, 'w:t');
  text.textContent = body;
  preserveTextSpace(text);
  run.appendChild(text);
  paragraph.appendChild(run);
  comment.appendChild(paragraph);
  root.appendChild(comment);
  markDirty(part);
}

function appendExtendedCommentRelation(pkg: Package, commentId: number, parentCommentId: number): void {
  const commentsPart = pkg.parts.get('word/comments.xml');
  const commentsDom = commentsPart?.xmlDom;
  if (!commentsDom) return;
  const commentParaIdFor = (id: number): string | null => {
    const comments = commentsDom.getElementsByTagNameNS(W_NS, 'comment');
    for (let index = 0; index < comments.length; index += 1) {
      if (comments[index]?.getAttributeNS(W_NS, 'id') !== String(id)) continue;
      const paragraph = comments[index]?.getElementsByTagNameNS(W_NS, 'p')[0];
      return paragraph?.getAttributeNS(W14_NS, 'paraId') ?? null;
    }
    return null;
  };
  const paraId = commentParaIdFor(commentId);
  const parentParaId = commentParaIdFor(parentCommentId);
  if (!paraId || !parentParaId) return;
  const part = ensureCommentsExtendedPart(pkg, 'word/document.xml');
  const dom = part.xmlDom;
  const root = dom?.documentElement;
  if (!dom || !root) return;
  const existing = Array.from(root.getElementsByTagNameNS(W15_NS, 'commentEx'))
    .some((entry) => entry.getAttributeNS(W15_NS, 'paraId') === paraId);
  if (existing) return;
  const relation = dom.createElementNS(W15_NS, 'w15:commentEx');
  relation.setAttributeNS(W15_NS, 'w15:paraId', paraId);
  relation.setAttributeNS(W15_NS, 'w15:paraIdParent', parentParaId);
  root.appendChild(relation);
  markDirty(part);
}

function commentParaId(commentId: number): string {
  return commentId.toString(16).padStart(8, '0').slice(-8);
}

function removeExtendedCommentRelation(pkg: Package, commentId: number): void {
  const commentsPart = pkg.parts.get('word/comments.xml');
  const commentsDom = commentsPart?.xmlDom;
  const extendedPart = pkg.parts.get('word/commentsExtended.xml');
  const extendedDom = extendedPart?.xmlDom;
  if (!commentsDom || !extendedDom) return;
  const comments = commentsDom.getElementsByTagNameNS(W_NS, 'comment');
  let paraId: string | null = null;
  for (let index = 0; index < comments.length; index += 1) {
    if (comments[index]?.getAttributeNS(W_NS, 'id') !== String(commentId)) continue;
    paraId = comments[index]?.getElementsByTagNameNS(W_NS, 'p')[0]?.getAttributeNS(W14_NS, 'paraId') ?? null;
    break;
  }
  if (!paraId) return;
  const relations = Array.from(extendedDom.getElementsByTagNameNS(W15_NS, 'commentEx'));
  for (const relation of relations) {
    if (relation.getAttributeNS(W15_NS, 'paraId') === paraId || relation.getAttributeNS(W15_NS, 'paraIdParent') === paraId) {
      removeNode(relation);
      markDirty(extendedPart);
    }
  }
}

function ensureCommentsExtendedPart(pkg: Package, partPath: string): Part {
  let part = pkg.parts.get('word/commentsExtended.xml');
  if (!part?.xmlDom) {
    const sourceDocument = pkg.parts.get(partPath)?.xmlDom ?? Array.from(pkg.parts.values()).find((candidate) => candidate.xmlDom)?.xmlDom;
    if (!sourceDocument) throw new Error('cannot create commentsExtended.xml without an XML part');
    const dom = sourceDocument.implementation?.createDocument(W15_NS, 'w15:commentsEx', null) as unknown as Document;
    if (!dom) throw new Error('XML implementation cannot create commentsExtended.xml');
    part = createXmlPart({ path: 'word/commentsExtended.xml', contentType: COMMENTS_EXTENDED_CONTENT_TYPE, xmlDom: dom });
    setPart(pkg, part);
  }
  const contentTypes = getPart(pkg, '[Content_Types].xml');
  if (contentTypes?.xmlDom) {
    const overrides = contentTypes.xmlDom.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
    if (!Array.from(overrides).some((entry) => entry.getAttribute('PartName') === '/word/commentsExtended.xml')) {
      const override = contentTypes.xmlDom.createElementNS(CONTENT_TYPES_NS, 'Override');
      override.setAttribute('PartName', '/word/commentsExtended.xml');
      override.setAttribute('ContentType', COMMENTS_EXTENDED_CONTENT_TYPE);
      contentTypes.xmlDom.documentElement?.appendChild(override);
      markDirty(contentTypes);
    }
  }
  const relsPart = pkg.parts.get(relsPathFor(partPath));
  const relsDom = relsPart?.xmlDom;
  if (relsDom && !Array.from(relsDom.getElementsByTagNameNS(PACKAGE_RELS_NS, 'Relationship'))
    .some((entry) => entry.getAttribute('Type') === COMMENTS_EXTENDED_REL_TYPE)) {
    const relationships = relsDom.getElementsByTagNameNS(PACKAGE_RELS_NS, 'Relationship');
    const used = new Set(Array.from(relationships).map((entry) => entry.getAttribute('Id')).filter(Boolean));
    let idNumber = 1;
    while (used.has(`rId${idNumber}`)) idNumber += 1;
    const relationship = relsDom.createElementNS(PACKAGE_RELS_NS, 'Relationship');
    relationship.setAttribute('Id', `rId${idNumber}`);
    relationship.setAttribute('Type', COMMENTS_EXTENDED_REL_TYPE);
    relationship.setAttribute('Target', 'commentsExtended.xml');
    relsDom.documentElement?.appendChild(relationship);
    markDirty(relsPart!);
  }
  return part;
}

function insertCommentAnchors(paragraph: Element, from: number, to: number, commentId: number): void {
  const document = paragraph.ownerDocument!;
  const start = document.createElementNS(W_NS, 'w:commentRangeStart');
  start.setAttributeNS(W_NS, 'w:id', String(commentId));
  const end = document.createElementNS(W_NS, 'w:commentRangeEnd');
  end.setAttributeNS(W_NS, 'w:id', String(commentId));
  const referenceRun = document.createElementNS(W_NS, 'w:r');
  const reference = document.createElementNS(W_NS, 'w:commentReference');
  reference.setAttributeNS(W_NS, 'w:id', String(commentId));
  referenceRun.appendChild(reference);
  const markers = new Map<number, Element[]>();
  markers.set(from, [start]);
  markers.set(to, [end, referenceRun]);
  const children = directChildren(paragraph).filter((child) => !(child.namespaceURI === W_NS && child.localName === 'pPr'));
  const next: Element[] = [];
  let position = 0;
  for (const child of children) {
    const atStartMarkers = markers.get(position);
    if (atStartMarkers) next.push(...atStartMarkers);
    const length = inlineNodeLength(child);
    const boundaries = [from, to].filter((boundary) => boundary > position && boundary < position + length).sort((a, b) => a - b);
    let local = 0;
    for (const boundary of [...boundaries, position + length]) {
      const piece = sliceInlineElement(child, local, boundary - position);
      if (piece) next.push(piece);
      const absolute = boundary;
      const boundaryMarkers = markers.get(absolute);
      if (boundaryMarkers) next.push(...boundaryMarkers);
      local = boundary - position;
    }
    if (length === 0) next.push(child);
    position += length;
  }
  for (const boundary of [from, to]) {
    if (boundary === position) {
      const boundaryMarkers = markers.get(boundary);
      if (boundaryMarkers) next.push(...boundaryMarkers);
    }
  }
  for (const child of children) removeNode(child);
  for (const child of next) paragraph.appendChild(child);
}

function removeCommentAnchors(pkg: Package, commentId: number): void {
  for (const part of pkg.parts.values()) {
    const dom = part.xmlDom;
    if (!dom) continue;
    for (const localName of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
      const nodes = dom.getElementsByTagNameNS(W_NS, localName);
      const matches: Element[] = [];
      for (let index = 0; index < nodes.length; index += 1) {
        const node = nodes[index];
        if (node?.getAttributeNS(W_NS, 'id') === String(commentId)) matches.push(node);
      }
      for (const node of matches) {
        const parent = node.parentNode;
        if (localName === 'commentReference' && parent?.nodeType === 1 && (parent as Element).namespaceURI === W_NS && (parent as Element).localName === 'r') {
          removeNode(parent);
        } else {
          removeNode(node);
        }
      }
      if (matches.length > 0) markDirty(part);
    }
  }
}

function isDrawingOnlyRun(run: Element): boolean {
  for (let i = 0; i < run.childNodes.length; i++) {
    const child = run.childNodes[i];
    if (!child || child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI === W_NS && (el.localName === 'rPr' || el.localName === 'drawing')) continue;
    return false;
  }
  return true;
}

function isElement(node: Node): node is Element {
  return node.nodeType === 1;
}
