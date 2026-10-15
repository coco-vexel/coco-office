import type { MarkType, Node as ProseMirrorNode } from 'prosemirror-model';
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { EditorParagraphFormat } from './model.js';
import { wordEditorSchema, wordStructureNormalizationKey } from './prosemirror.js';

export type WordEditMark = 'strong' | 'textStyle';
export type WordBlockEditOperation = 'insert' | 'delete' | 'replace' | 'split' | 'merge' | 'move';
export type WordAnnotationEditOperation = 'insert' | 'delete' | 'update';
export type WordContentEditOperation = 'insert' | 'delete' | 'replace';
export type WordRevisionEditOperation = 'insert' | 'delete' | 'accept' | 'reject';
export type WordTableEditOperation =
  | 'insertTable'
  | 'deleteTable'
  | 'insertRow'
  | 'deleteRow'
  | 'insertColumn'
  | 'deleteColumn'
  | 'mergeCells'
  | 'splitCell'
  | 'updateCell'
  | 'updateTable';
export type WordTextEditOperation = 'insert' | 'delete' | 'replace';
export type WordParagraphAlignment = NonNullable<EditorParagraphFormat['alignment']>;

export type WordChangeStatus = 'captured' | 'planned' | 'exportable';

export interface WordChangeTarget {
  story: string;
  paraId?: string;
  tableId?: string;
  rowIndex?: number;
  columnIndex?: number;
  cellId?: string;
  annotationId?: string;
  revisionId?: string;
}

export interface WordEditorRange {
  from: number;
  to: number;
}

export interface WordParagraphRange extends WordEditorRange {
  paraId: string;
  story: string;
}

export interface WordParagraphBlockSnapshot {
  kind: 'paragraph';
  paraId: string;
  story: string;
  styleId: string | null;
  text: string;
  content: WordContent[];
}

export interface WordTableBlockSnapshot {
  kind: 'table';
  tableId: string;
  text: string;
}

export type WordBlockSnapshot = WordParagraphBlockSnapshot | WordTableBlockSnapshot;

export interface WordImageContent {
  kind: 'image';
  drawingId: string | null;
  alt: string;
  src: string;
  width: number | null;
  height: number | null;
}

export interface WordTextContent {
  kind: 'text';
  text: string;
}

export type WordContent = WordImageContent | WordTextContent;

export interface WordSelectedParagraph {
  paraId: string;
  from: number;
  to: number;
  text: string;
  styleId: string | null;
  story: string;
}

export interface WordSelectionInfo {
  from: number;
  to: number;
  empty: boolean;
  selectedText: string;
  paragraphs: WordSelectedParagraph[];
  tableCell: {
    depth: number;
    colspan: number;
    rowspan: number;
  } | null;
}

export interface WordEditChangeBase {
  from: number;
  to: number;
  paragraphs: WordSelectedParagraph[];
  target: WordChangeTarget;
  status: WordChangeStatus;
  value: unknown;
}

export interface WordTextEditChangeDraft extends WordEditChangeBase {
  kind: 'text';
  textOp: WordTextEditOperation;
  text: string;
  deletedText: string;
}

export interface WordContentEditChangeDraft extends WordEditChangeBase {
  kind: 'content';
  contentOp: WordContentEditOperation;
  content: WordContent[];
  deletedContent: WordContent[];
}

export interface WordMarkEditChangeDraft extends WordEditChangeBase {
  kind: 'mark';
  mark: WordEditMark;
}

export interface WordParagraphEditChangeDraft extends WordEditChangeBase {
  kind: 'paragraph';
  attr: 'alignment';
}

export interface WordBlockEditChangeDraft extends WordEditChangeBase {
  kind: 'block';
  blockOp: WordBlockEditOperation;
  blocks?: WordBlockSnapshot[];
  deletedBlocks?: WordBlockSnapshot[];
}

export interface WordTableEditChangeDraft extends WordEditChangeBase {
  kind: 'table';
  tableOp: WordTableEditOperation;
  tableId?: string;
}

export interface WordAnnotationEditChangeDraft extends WordEditChangeBase {
  kind: 'annotation';
  annotationOp: WordAnnotationEditOperation;
  annotationId: string;
}

export interface WordRevisionEditChangeDraft extends WordEditChangeBase {
  kind: 'revision';
  revisionOp: WordRevisionEditOperation;
  revisionId?: string;
}

export type WordEditChangeDraft =
  | WordTextEditChangeDraft
  | WordContentEditChangeDraft
  | WordMarkEditChangeDraft
  | WordParagraphEditChangeDraft
  | WordBlockEditChangeDraft
  | WordTableEditChangeDraft
  | WordAnnotationEditChangeDraft
  | WordRevisionEditChangeDraft;

export type WordEditChange = WordEditChangeDraft & {
  id: string;
  createdAt: number;
};

export interface WordChangeSet {
  changes: WordEditChange[];
}

export interface WordEditingPluginOptions {
  onChangeSet?: (changeSet: WordChangeSet) => void;
}

interface WordEditingTransactionMeta {
  changes?: WordEditChange[];
}

export const wordEditingPluginKey = new PluginKey<WordChangeSet>('coco-word-editing');

let changeSequence = 0;

export function createWordEditingPlugin(options: WordEditingPluginOptions = {}): Plugin<WordChangeSet> {
  return new Plugin<WordChangeSet>({
    key: wordEditingPluginKey,
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((tr) => tr.docChanged)) return null;
      if (transactions.every((tr) => tr.getMeta(wordStructureNormalizationKey))) return null;
      const blockChanges = blockChangesFromDocumentDiff(oldState.doc, newState.doc);
      const changes =
        blockChanges.length > 0 ? blockChanges : contentChangesFromDocumentDiff(oldState.doc, newState.doc);
      return changes.length > 0 ? withChanges(newState.tr, changes) : null;
    },
    state: {
      init: () => ({ changes: [] }),
      apply(tr, previous) {
        const meta = tr.getMeta(wordEditingPluginKey) as WordEditingTransactionMeta | undefined;
        if (!meta?.changes?.length) return previous;
        const next = { changes: [...previous.changes, ...meta.changes] };
        options.onChangeSet?.(next);
        return next;
      },
    },
  });
}

export function getWordChangeSet(state: EditorState): WordChangeSet {
  return wordEditingPluginKey.getState(state) ?? { changes: [] };
}

export function getWordSelectionInfo(state: EditorState): WordSelectionInfo {
  const { from, to, empty } = state.selection;
  const $from = state.selection.$from;
  const cellDepth = findParentDepth($from, 'table_cell');
  const tableCell =
    cellDepth === null
      ? null
      : {
          depth: cellDepth,
          colspan: Number($from.node(cellDepth).attrs.colspan ?? 1),
          rowspan: Number($from.node(cellDepth).attrs.rowspan ?? 1),
        };

  return {
    from,
    to,
    empty,
    selectedText: state.doc.textBetween(from, to, '\n'),
    paragraphs: selectedParagraphs(state, from, to),
    tableCell,
  };
}

export function toggleBold(
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
  view?: EditorView,
): boolean {
  const markType = wordEditorSchema.marks.strong;
  return applySelectionMark(state, dispatch, view, markType, {}, 'strong');
}

export function setFontSize(sizePt: number) {
  return (state: EditorState, dispatch?: (tr: Transaction) => void, view?: EditorView): boolean => {
    if (!Number.isFinite(sizePt) || sizePt <= 0) return false;
    return applySelectionMark(state, dispatch, view, wordEditorSchema.marks.textStyle, { fontSize: sizePt }, 'textStyle');
  };
}

export function setColor(color: string) {
  return (state: EditorState, dispatch?: (tr: Transaction) => void, view?: EditorView): boolean => {
    const normalized = normalizeHexColor(color);
    if (!normalized) return false;
    return applySelectionMark(state, dispatch, view, wordEditorSchema.marks.textStyle, { color: normalized }, 'textStyle');
  };
}

export function setParagraphAlign(alignment: WordParagraphAlignment) {
  return (state: EditorState, dispatch?: (tr: Transaction) => void): boolean => {
    const paragraphs = selectedParagraphs(state, state.selection.from, state.selection.to);
    if (paragraphs.length === 0) return false;
    if (!dispatch) return true;

    let tr = state.tr;
    state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
      if (node.type !== wordEditorSchema.nodes.paragraph) return true;
      const format = mergeParagraphFormat(node.attrs.format, { alignment });
      tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, format });
      return false;
    });

    tr = withChange(tr, {
      kind: 'paragraph',
      from: state.selection.from,
      to: state.selection.to,
      paragraphs,
      target: changeTargetForParagraphs(paragraphs),
      status: 'captured',
      attr: 'alignment',
      value: alignment,
    });
    dispatch(tr.scrollIntoView());
    return true;
  };
}

function applySelectionMark(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  view: EditorView | undefined,
  markType: MarkType,
  attrs: Record<string, unknown>,
  mark: WordEditMark,
): boolean {
  const { from, to, empty } = state.selection;
  if (empty) return false;
  const paragraphs = selectedParagraphs(state, from, to);
  if (paragraphs.length === 0) return false;
  const shouldRemove = mark === 'strong' && rangeHasMark(state.doc, from, to, markType);
  if (!dispatch) return true;

  let tr = state.tr;
  if (shouldRemove) {
    tr = tr.removeMark(from, to, markType);
  } else {
    tr = tr.addMark(from, to, markType.create(attrs));
  }
  tr = withChange(tr, {
    kind: 'mark',
    from,
    to,
    paragraphs,
    target: changeTargetForParagraphs(paragraphs),
    status: 'captured',
    mark,
    value: mark === 'strong' ? !shouldRemove : attrs,
  });
  dispatch(tr.scrollIntoView());
  view?.focus();
  return true;
}

function selectedParagraphs(state: EditorState, from: number, to: number): WordSelectedParagraph[] {
  const paragraphs: WordSelectedParagraph[] = [];
  const scanTo = Math.max(to, from + 1);
  state.doc.nodesBetween(from, scanTo, (node, pos) => {
    if (node.type !== wordEditorSchema.nodes.paragraph) return true;
    const contentStart = pos + 1;
    const contentEnd = contentStart + node.content.size;
    paragraphs.push({
      paraId: typeof node.attrs.paraId === 'string' ? node.attrs.paraId : '',
      from: Math.max(from, contentStart) - contentStart,
      to: Math.min(to, contentEnd) - contentStart,
      text: node.textContent,
      styleId: typeof node.attrs.styleId === 'string' ? node.attrs.styleId : null,
      story: typeof node.attrs.story === 'string' ? node.attrs.story : 'body',
    });
    return false;
  });
  return paragraphs;
}

function changeTargetForParagraphs(paragraphs: WordSelectedParagraph[]): WordChangeTarget {
  const paragraph = paragraphs[0];
  return {
    story: paragraph?.story ?? 'body',
    ...(paragraph?.paraId ? { paraId: paragraph.paraId } : {}),
  };
}

function rangeHasMark(doc: ProseMirrorNode, from: number, to: number, markType: MarkType): boolean {
  let found = false;
  doc.nodesBetween(from, to, (node) => {
    if (found || !node.isText) return !found;
    found = markType.isInSet(node.marks) !== undefined;
    return !found;
  });
  return found;
}

function withChange(
  tr: Transaction,
  change: WordEditChangeDraft,
): Transaction {
  return withChanges(tr, [change]);
}

function withChanges(
  tr: Transaction,
  changesToAdd: WordEditChangeDraft[],
): Transaction {
  const meta = (tr.getMeta(wordEditingPluginKey) as WordEditingTransactionMeta | undefined) ?? {};
  const changes = meta.changes ?? [];
  return tr.setMeta(wordEditingPluginKey, {
    ...meta,
    changes: [
      ...changes,
      ...changesToAdd.map((change) => ({
        ...change,
        id: `word-edit:${++changeSequence}`,
        createdAt: Date.now(),
      })),
    ],
  });
}

interface ParagraphContentUnit {
  key: string;
  content: WordContent;
}

interface ParagraphSnapshot {
  paraId: string;
  pos: number;
  text: string;
  content: ParagraphContentUnit[];
  styleId: string | null;
  story: string;
}

interface TopLevelBlockSnapshot {
  key: string;
  pos: number;
  size: number;
  block: WordBlockSnapshot;
}

function blockChangesFromDocumentDiff(oldDoc: ProseMirrorNode, newDoc: ProseMirrorNode): WordEditChangeDraft[] {
  const oldBlocks = topLevelBlockSnapshots(oldDoc);
  const newBlocks = topLevelBlockSnapshots(newDoc);
  const split = splitBlockChange(oldBlocks, newBlocks);
  if (split) return [split];
  const merge = mergeBlockChange(oldBlocks, newBlocks);
  if (merge) return [merge];
  const insertOrDelete = insertOrDeleteBlockChange(oldBlocks, newBlocks);
  return insertOrDelete ? [insertOrDelete] : [];
}

function splitBlockChange(
  oldBlocks: TopLevelBlockSnapshot[],
  newBlocks: TopLevelBlockSnapshot[],
): WordEditChangeDraft | null {
  if (newBlocks.length !== oldBlocks.length + 1) return null;
  for (let index = 0; index < oldBlocks.length; index += 1) {
    if (!blockSequencesEqual(oldBlocks.slice(0, index), newBlocks.slice(0, index))) continue;
    if (!blockSequencesEqual(oldBlocks.slice(index + 1), newBlocks.slice(index + 2))) continue;
    const before = oldBlocks[index]!;
    const first = newBlocks[index]!;
    const second = newBlocks[index + 1]!;
    if (
      before.block.kind !== 'paragraph' ||
      first.block.kind !== 'paragraph' ||
      second.block.kind !== 'paragraph' ||
      first.block.text + second.block.text !== before.block.text
    ) {
      continue;
    }

    return {
      kind: 'block',
      blockOp: 'split',
      from: before.pos,
      to: before.pos + before.size,
      paragraphs: blockParagraphs([first, second]),
      target: targetForBlock(before.block),
      status: 'captured',
      blocks: [first.block, second.block],
      deletedBlocks: [before.block],
      value: {
        before: before.block,
        after: [first.block, second.block],
      },
    };
  }
  return null;
}

function mergeBlockChange(
  oldBlocks: TopLevelBlockSnapshot[],
  newBlocks: TopLevelBlockSnapshot[],
): WordEditChangeDraft | null {
  if (oldBlocks.length !== newBlocks.length + 1) return null;
  for (let index = 0; index < newBlocks.length; index += 1) {
    if (!blockSequencesEqual(oldBlocks.slice(0, index), newBlocks.slice(0, index))) continue;
    if (!blockSequencesEqual(oldBlocks.slice(index + 2), newBlocks.slice(index + 1))) continue;
    const first = oldBlocks[index]!;
    const second = oldBlocks[index + 1]!;
    const after = newBlocks[index]!;
    if (
      first.block.kind !== 'paragraph' ||
      second.block.kind !== 'paragraph' ||
      after.block.kind !== 'paragraph' ||
      first.block.text + second.block.text !== after.block.text
    ) {
      continue;
    }

    return {
      kind: 'block',
      blockOp: 'merge',
      from: first.pos,
      to: second.pos + second.size,
      paragraphs: blockParagraphs([after]),
      target: targetForBlock(after.block),
      status: 'captured',
      blocks: [after.block],
      deletedBlocks: [first.block, second.block],
      value: {
        before: [first.block, second.block],
        after: after.block,
      },
    };
  }
  return null;
}

function insertOrDeleteBlockChange(
  oldBlocks: TopLevelBlockSnapshot[],
  newBlocks: TopLevelBlockSnapshot[],
): WordEditChangeDraft | null {
  let prefix = 0;
  const maxPrefix = Math.min(oldBlocks.length, newBlocks.length);
  while (prefix < maxPrefix && blockEquals(oldBlocks[prefix]!, newBlocks[prefix]!)) prefix += 1;

  let suffix = 0;
  const maxSuffix = Math.min(oldBlocks.length - prefix, newBlocks.length - prefix);
  while (
    suffix < maxSuffix &&
    blockEquals(oldBlocks[oldBlocks.length - 1 - suffix]!, newBlocks[newBlocks.length - 1 - suffix]!)
  ) {
    suffix += 1;
  }

  const deleted = oldBlocks.slice(prefix, oldBlocks.length - suffix);
  const inserted = newBlocks.slice(prefix, newBlocks.length - suffix);
  if (deleted.length === 0 && inserted.length === 0) return null;
  const first = inserted[0] ?? deleted[0];
  const last = inserted[inserted.length - 1] ?? deleted[deleted.length - 1] ?? first;
  const operation =
    inserted.length > 0 && deleted.length > 0 ? 'replace' : inserted.length > 0 ? 'insert' : 'delete';

  return {
    kind: 'block',
    blockOp: operation,
    from: first?.pos ?? 0,
    to: last ? last.pos + last.size : first?.pos ?? 0,
    paragraphs: blockParagraphs(inserted.length > 0 ? inserted : deleted),
    target: first ? targetForBlock(first.block) : { story: 'body' },
    status: 'captured',
    blocks: inserted.map((item) => item.block),
    deletedBlocks: deleted.map((item) => item.block),
    value: {
      before: deleted.map((item) => item.block),
      after: inserted.map((item) => item.block),
    },
  };
}

function contentChangesFromDocumentDiff(
  oldDoc: ProseMirrorNode,
  newDoc: ProseMirrorNode,
): WordEditChangeDraft[] {
  const oldParagraphs = paragraphSnapshots(oldDoc);
  const newParagraphs = paragraphSnapshots(newDoc);
  const oldByParaId = new Map(oldParagraphs.map((paragraph) => [paragraph.paraId, paragraph]));
  const changes: WordEditChangeDraft[] = [];

  for (const next of newParagraphs) {
    if (!next.paraId) continue;
    const previous = oldByParaId.get(next.paraId);
    if (!previous || paragraphContentEquals(previous.content, next.content)) continue;
    const diff = contentDiff(previous.content, next.content);
    if (!diff) continue;
    const from = next.pos + 1 + diff.from;
    const to = next.pos + 1 + diff.toAfter;
    const paragraph = {
      paraId: next.paraId,
      from: diff.from,
      to: diff.toAfter,
      text: next.text,
      styleId: next.styleId,
      story: next.story,
    };

    if (contentListIsTextOnly(diff.insertedContent) && contentListIsTextOnly(diff.deletedContent)) {
      changes.push({
        kind: 'text',
        from,
        to,
        paragraphs: [paragraph],
        target: changeTargetForParagraphs([paragraph]),
        status: 'captured',
        textOp: diff.operation,
        text: contentText(diff.insertedContent),
        deletedText: contentText(diff.deletedContent),
        value: {
          operation: diff.operation,
          before: previous.text,
          after: next.text,
        },
      });
      continue;
    }

    changes.push({
      kind: 'content',
      from,
      to,
      paragraphs: [paragraph],
      target: changeTargetForParagraphs([paragraph]),
      status: 'captured',
      contentOp: diff.operation,
      content: mergeTextContent(diff.insertedContent),
      deletedContent: mergeTextContent(diff.deletedContent),
      value: {
        operation: diff.operation,
        before: mergeTextContent(previous.content.map((unit) => unit.content)),
        after: mergeTextContent(next.content.map((unit) => unit.content)),
      },
    });
  }

  return changes;
}

function paragraphSnapshots(doc: ProseMirrorNode): ParagraphSnapshot[] {
  const paragraphs: ParagraphSnapshot[] = [];
  doc.descendants((node, pos) => {
    if (node.type !== wordEditorSchema.nodes.paragraph) return true;
    paragraphs.push({
      paraId: typeof node.attrs.paraId === 'string' ? node.attrs.paraId : '',
      pos,
      text: node.textContent,
      content: paragraphContentUnits(node),
      styleId: typeof node.attrs.styleId === 'string' ? node.attrs.styleId : null,
      story: typeof node.attrs.story === 'string' ? node.attrs.story : 'body',
    });
    return false;
  });
  return paragraphs;
}

function topLevelBlockSnapshots(doc: ProseMirrorNode): TopLevelBlockSnapshot[] {
  const blocks: TopLevelBlockSnapshot[] = [];
  doc.forEach((node, offset, index) => {
    if (node.type === wordEditorSchema.nodes.paragraph) {
      const paraId = typeof node.attrs.paraId === 'string' ? node.attrs.paraId : '';
      const story = typeof node.attrs.story === 'string' ? node.attrs.story : 'body';
      const block: WordParagraphBlockSnapshot = {
        kind: 'paragraph',
        paraId,
        story,
        styleId: typeof node.attrs.styleId === 'string' ? node.attrs.styleId : null,
        text: node.textContent,
        content: mergeTextContent(paragraphContentUnits(node).map((unit) => unit.content)),
      };
      blocks.push({
        key: `paragraph:${paraId || index}`,
        pos: offset,
        size: node.nodeSize,
        block,
      });
      return;
    }

    if (node.type === wordEditorSchema.nodes.table) {
      const tableId = typeof node.attrs.tableId === 'string' ? node.attrs.tableId : '';
      const block: WordTableBlockSnapshot = {
        kind: 'table',
        tableId,
        text: node.textContent,
      };
      blocks.push({
        key: `table:${tableId || index}`,
        pos: offset,
        size: node.nodeSize,
        block,
      });
    }
  });
  return blocks;
}

function blockSequencesEqual(left: TopLevelBlockSnapshot[], right: TopLevelBlockSnapshot[]): boolean {
  return left.length === right.length && left.every((block, index) => blockEquals(block, right[index]!));
}

function blockEquals(left: TopLevelBlockSnapshot, right: TopLevelBlockSnapshot): boolean {
  return left.key === right.key;
}

function blockParagraphs(blocks: TopLevelBlockSnapshot[]): WordSelectedParagraph[] {
  return blocks.flatMap((snapshot) => {
    if (snapshot.block.kind !== 'paragraph') return [];
    return [
      {
        paraId: snapshot.block.paraId,
        from: 0,
        to: snapshot.block.text.length,
        text: snapshot.block.text,
        styleId: snapshot.block.styleId,
        story: snapshot.block.story,
      },
    ];
  });
}

function targetForBlock(block: WordBlockSnapshot): WordChangeTarget {
  if (block.kind === 'table') {
    return {
      story: 'body',
      ...(block.tableId ? { tableId: block.tableId } : {}),
    };
  }

  return {
    story: block.story,
    ...(block.paraId ? { paraId: block.paraId } : {}),
  };
}

function paragraphContentUnits(paragraph: ProseMirrorNode): ParagraphContentUnit[] {
  const units: ParagraphContentUnit[] = [];
  paragraph.forEach((child) => {
    if (child.isText) {
      for (const char of child.text ?? '') {
        units.push({ key: `text:${char}`, content: { kind: 'text', text: char } });
      }
      return;
    }

    if (child.type === wordEditorSchema.nodes.image) {
      const image = imageContent(child);
      units.push({ key: `image:${image.drawingId ?? image.src}`, content: image });
      return;
    }

    units.push({
      key: `${child.type.name}:${JSON.stringify(child.attrs)}`,
      content: { kind: 'text', text: child.textContent },
    });
  });
  return units;
}

function imageContent(node: ProseMirrorNode): WordImageContent {
  return {
    kind: 'image',
    drawingId: typeof node.attrs.drawingId === 'string' ? node.attrs.drawingId : null,
    alt: typeof node.attrs.alt === 'string' ? node.attrs.alt : '',
    src: typeof node.attrs.src === 'string' ? node.attrs.src : '',
    width: typeof node.attrs.width === 'number' ? node.attrs.width : null,
    height: typeof node.attrs.height === 'number' ? node.attrs.height : null,
  };
}

function contentDiff(
  before: ParagraphContentUnit[],
  after: ParagraphContentUnit[],
): {
  operation: WordContentEditOperation;
  from: number;
  toBefore: number;
  toAfter: number;
  insertedContent: WordContent[];
  deletedContent: WordContent[];
} | null {
  if (paragraphContentEquals(before, after)) return null;

  let prefix = 0;
  const maxPrefix = Math.min(before.length, after.length);
  while (prefix < maxPrefix && before[prefix]?.key === after[prefix]?.key) prefix += 1;

  let suffix = 0;
  const maxSuffix = Math.min(before.length - prefix, after.length - prefix);
  while (
    suffix < maxSuffix &&
    before[before.length - 1 - suffix]?.key === after[after.length - 1 - suffix]?.key
  ) {
    suffix += 1;
  }

  const toBefore = before.length - suffix;
  const toAfter = after.length - suffix;
  const insertedContent = after.slice(prefix, toAfter).map((unit) => unit.content);
  const deletedContent = before.slice(prefix, toBefore).map((unit) => unit.content);
  const operation =
    insertedContent.length > 0 && deletedContent.length > 0
      ? 'replace'
      : insertedContent.length > 0
        ? 'insert'
        : 'delete';

  return {
    operation,
    from: prefix,
    toBefore,
    toAfter,
    insertedContent,
    deletedContent,
  };
}

function paragraphContentEquals(left: ParagraphContentUnit[], right: ParagraphContentUnit[]): boolean {
  return left.length === right.length && left.every((unit, index) => unit.key === right[index]?.key);
}

function contentListIsTextOnly(content: WordContent[]): boolean {
  return content.every((item) => item.kind === 'text');
}

function contentText(content: WordContent[]): string {
  return content.map((item) => (item.kind === 'text' ? item.text : '')).join('');
}

function mergeTextContent(content: WordContent[]): WordContent[] {
  const merged: WordContent[] = [];
  for (const item of content) {
    const previous = merged[merged.length - 1];
    if (item.kind === 'text' && previous?.kind === 'text') {
      previous.text += item.text;
    } else {
      merged.push({ ...item });
    }
  }
  return merged;
}

function mergeParagraphFormat(value: unknown, patch: Partial<EditorParagraphFormat>): EditorParagraphFormat {
  const base = value !== null && typeof value === 'object' ? (value as EditorParagraphFormat) : {};
  return { ...base, ...patch };
}

function normalizeHexColor(color: string): string | null {
  const trimmed = color.trim().replace(/^#/, '');
  return /^[0-9a-f]{6}$/i.test(trimmed) ? trimmed.toUpperCase() : null;
}

function findParentDepth($pos: EditorState['selection']['$from'], nodeName: string): number | null {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    if ($pos.node(depth).type.name === nodeName) return depth;
  }
  return null;
}
