import type {
  NativeWordBlockNode,
  NativeWordDocument,
  NativeWordInlineNode,
  NativeWordParagraphNode,
  NativeWordTableCellNode,
  NativeWordTableNode,
  NativeWordTableRowNode,
} from './native-model.js';

export type NativeTableRowPaginationKind =
  | 'repeat-header'
  | 'keep-together'
  | 'block-fragmentable'
  | 'text-fragmentable'
  | 'unsplittable';

export type NativeTableRowPaginationRisk =
  | 'cant-split'
  | 'row-span'
  | 'vertical-merge'
  | 'nested-table'
  | 'media'
  | 'field'
  | 'page-break'
  | 'comment'
  | 'revision'
  | 'empty-row';

export interface NativeTableRowPaginationAnalysis {
  kind: NativeTableRowPaginationKind;
  risks: NativeTableRowPaginationRisk[];
  textLength: number;
  paragraphCount: number;
}

export interface NativeTablePaginationSummary {
  tables: number;
  rows: number;
  repeatHeaderRows: number;
  keepTogetherRows: number;
  blockFragmentableRows: number;
  textFragmentableRows: number;
  unsplittableRows: number;
  risks: Record<NativeTableRowPaginationRisk, number>;
}

export interface NativeTableRowBlockSplitPlan {
  cellBlockCounts: number[];
}

export interface NativeTableRowBlockSplitResult {
  head: NativeWordTableRowNode;
  tail: NativeWordTableRowNode;
}

export interface NativeTableRowTextSplitPlan {
  cellTextOffsets: number[];
}

export function analyzeNativeTableRowPagination(
  row: NativeWordTableRowNode,
): NativeTableRowPaginationAnalysis {
  const risks = nativeTableRowPaginationRisks(row);
  const paragraphCount = row.cells.reduce((count, cell) => count + countCellParagraphs(cell), 0);
  const textLength = row.cells.reduce((count, cell) => count + nativeCellTextLength(cell), 0);

  if (row.isHeader) return { kind: 'repeat-header', risks, textLength, paragraphCount };
  if (risks.includes('cant-split')) return { kind: 'keep-together', risks, textLength, paragraphCount };
  if (risks.length > 0) return { kind: 'unsplittable', risks, textLength, paragraphCount };
  if (textLength === 0) return { kind: 'keep-together', risks: ['empty-row'], textLength, paragraphCount };
  if (paragraphCount > row.cells.length) {
    return { kind: 'block-fragmentable', risks, textLength, paragraphCount };
  }
  return { kind: 'text-fragmentable', risks, textLength, paragraphCount };
}

export function summarizeNativeTablePagination(doc: NativeWordDocument): NativeTablePaginationSummary {
  const summary: NativeTablePaginationSummary = {
    tables: 0,
    rows: 0,
    repeatHeaderRows: 0,
    keepTogetherRows: 0,
    blockFragmentableRows: 0,
    textFragmentableRows: 0,
    unsplittableRows: 0,
    risks: {
      'cant-split': 0,
      'row-span': 0,
      'vertical-merge': 0,
      'nested-table': 0,
      media: 0,
      field: 0,
      'page-break': 0,
      comment: 0,
      revision: 0,
      'empty-row': 0,
    },
  };
  summarizeNativeTablePaginationFromBlocks(doc.blocks, summary);
  for (const story of doc.headers) summarizeNativeTablePaginationFromBlocks(story.blocks, summary);
  for (const story of doc.footers) summarizeNativeTablePaginationFromBlocks(story.blocks, summary);
  return summary;
}

export function splitNativeTableRowAtBlockBoundary(
  row: NativeWordTableRowNode,
  plan: NativeTableRowBlockSplitPlan,
): NativeTableRowBlockSplitResult | null {
  const analysis = analyzeNativeTableRowPagination(row);
  if (analysis.kind !== 'block-fragmentable') return null;
  if (plan.cellBlockCounts.length !== row.cells.length) return null;

  const headCells: NativeWordTableCellNode[] = [];
  const tailCells: NativeWordTableCellNode[] = [];
  let hasHeadContent = false;
  let hasTailContent = false;
  for (let index = 0; index < row.cells.length; index += 1) {
    const cell = row.cells[index]!;
    const split = splitNativeTableCellAtBlockBoundary(cell, plan.cellBlockCounts[index]!);
    if (!split) return null;
    headCells.push(split.head);
    tailCells.push(split.tail);
    hasHeadContent ||= split.head.blocks.length > 0;
    hasTailContent ||= split.tail.blocks.length > 0;
  }
  if (!hasHeadContent || !hasTailContent) return null;

  return {
    head: { ...row, height: undefined, cells: headCells },
    tail: { ...row, height: undefined, cells: tailCells },
  };
}

export function splitNativeTableRowAtTextOffset(
  row: NativeWordTableRowNode,
  plan: NativeTableRowTextSplitPlan,
): NativeTableRowBlockSplitResult | null {
  const analysis = analyzeNativeTableRowPagination(row);
  if (analysis.kind !== 'text-fragmentable') return null;
  if (plan.cellTextOffsets.length !== row.cells.length) return null;

  const headCells: NativeWordTableCellNode[] = [];
  const tailCells: NativeWordTableCellNode[] = [];
  for (let index = 0; index < row.cells.length; index += 1) {
    const split = splitNativeTableCellAtTextOffset(row.cells[index]!, plan.cellTextOffsets[index]!);
    if (!split) return null;
    headCells.push(split.head);
    tailCells.push(split.tail);
  }
  return {
    head: { ...row, height: undefined, cells: headCells },
    tail: { ...row, height: undefined, cells: tailCells },
  };
}

function splitNativeTableCellAtTextOffset(
  cell: NativeWordTableCellNode,
  textOffset: number,
): { head: NativeWordTableCellNode; tail: NativeWordTableCellNode } | null {
  if (cell.blocks.length !== 1) return null;
  const block = cell.blocks[0];
  if (!block || block.kind !== 'paragraph') return null;
  const split = splitNativeParagraphAtTextOffset(block, textOffset);
  if (!split) return null;
  return {
    head: {
      ...cell,
      blocks: [split.head],
      paragraphs: [split.head],
    },
    tail: {
      ...cell,
      blocks: [split.tail],
      paragraphs: [split.tail],
    },
  };
}

function splitNativeParagraphAtTextOffset(
  paragraph: NativeWordParagraphNode,
  textOffset: number,
): { head: NativeWordParagraphNode; tail: NativeWordParagraphNode } | null {
  const textLength = paragraph.content.reduce(
    (count, inline) => count + (inline.kind === 'text' ? inline.text.length : 0),
    0,
  );
  if (!Number.isInteger(textOffset) || textOffset <= 0 || textOffset >= textLength) return null;
  const headContent: NativeWordInlineNode[] = [];
  const tailContent: NativeWordInlineNode[] = [];
  let consumed = 0;
  for (const inline of paragraph.content) {
    if (inline.kind !== 'text') return null;
    const start = consumed;
    const end = consumed + inline.text.length;
    consumed = end;
    if (end <= textOffset) {
      headContent.push(inline);
      continue;
    }
    if (start >= textOffset) {
      tailContent.push(inline);
      continue;
    }
    const localOffset = textOffset - start;
    headContent.push({ ...inline, text: inline.text.slice(0, localOffset) });
    tailContent.push({ ...inline, text: inline.text.slice(localOffset) });
  }
  if (headContent.length === 0 || tailContent.length === 0) return null;
  return {
    head: { ...paragraph, content: headContent },
    tail: { ...paragraph, content: tailContent },
  };
}

function splitNativeTableCellAtBlockBoundary(
  cell: NativeWordTableCellNode,
  headBlockCount: number,
): { head: NativeWordTableCellNode; tail: NativeWordTableCellNode } | null {
  if (!Number.isInteger(headBlockCount)) return null;
  if (headBlockCount < 0 || headBlockCount > cell.blocks.length) return null;
  const headBlocks = cell.blocks.slice(0, headBlockCount);
  const tailBlocks = cell.blocks.slice(headBlockCount);
  if (headBlocks.length === 0 && tailBlocks.length === 0) return null;
  if (headBlocks.length === 0 || tailBlocks.length === 0) return null;
  return {
    head: {
      ...cell,
      blocks: headBlocks,
      paragraphs: headBlocks.flatMap(nativeParagraphsFromBlock),
    },
    tail: {
      ...cell,
      blocks: tailBlocks,
      paragraphs: tailBlocks.flatMap(nativeParagraphsFromBlock),
    },
  };
}

function summarizeNativeTablePaginationFromBlocks(
  blocks: NativeWordBlockNode[],
  summary: NativeTablePaginationSummary,
): void {
  for (const block of blocks) {
    if (block.kind !== 'table') continue;
    summarizeNativeTable(block, summary);
  }
}

function summarizeNativeTable(table: NativeWordTableNode, summary: NativeTablePaginationSummary): void {
  summary.tables += 1;
  for (const row of table.rows) {
    summary.rows += 1;
    const analysis = analyzeNativeTableRowPagination(row);
    if (analysis.kind === 'repeat-header') summary.repeatHeaderRows += 1;
    else if (analysis.kind === 'keep-together') summary.keepTogetherRows += 1;
    else if (analysis.kind === 'block-fragmentable') summary.blockFragmentableRows += 1;
    else if (analysis.kind === 'text-fragmentable') summary.textFragmentableRows += 1;
    else summary.unsplittableRows += 1;
    for (const risk of analysis.risks) summary.risks[risk] += 1;
    for (const cell of row.cells) summarizeNativeTablePaginationFromBlocks(cell.blocks, summary);
  }
}

function nativeTableRowPaginationRisks(row: NativeWordTableRowNode): NativeTableRowPaginationRisk[] {
  const risks = new Set<NativeTableRowPaginationRisk>();
  if (row.cantSplit) risks.add('cant-split');
  for (const cell of row.cells) {
    if ((cell.rowSpan ?? 1) > 1) risks.add('row-span');
    if (cell.vMerge === 'restart' || cell.vMerge === 'continue') risks.add('vertical-merge');
    collectCellRisks(cell, risks);
  }
  return [...risks];
}

function collectCellRisks(cell: NativeWordTableCellNode, risks: Set<NativeTableRowPaginationRisk>): void {
  for (const block of cell.blocks) {
    if (block.kind === 'table') {
      risks.add('nested-table');
      continue;
    }
    collectParagraphRisks(block, risks);
  }
}

function collectParagraphRisks(
  paragraph: NativeWordParagraphNode,
  risks: Set<NativeTableRowPaginationRisk>,
): void {
  for (const inline of paragraph.content) {
    collectInlineRisks(inline, risks);
  }
}

function collectInlineRisks(
  inline: NativeWordInlineNode,
  risks: Set<NativeTableRowPaginationRisk>,
): void {
  if (inline.kind === 'image' || inline.kind === 'drawing') risks.add('media');
  if (inline.kind === 'field') risks.add('field');
  if (inline.kind === 'pageBreak') risks.add('page-break');
  if (inline.kind !== 'text') return;
  if (inline.marks?.some((mark) => mark.kind === 'comment')) risks.add('comment');
  if (inline.marks?.some((mark) => mark.kind === 'revision')) risks.add('revision');
}

function countCellParagraphs(cell: NativeWordTableCellNode): number {
  return cell.blocks.reduce((count, block) => {
    if (block.kind === 'paragraph') return count + 1;
    return count + block.rows.reduce((rowCount, row) => rowCount + row.cells.reduce(
      (cellCount, nestedCell) => cellCount + countCellParagraphs(nestedCell),
      0,
    ), 0);
  }, 0);
}

function nativeCellTextLength(cell: NativeWordTableCellNode): number {
  return cell.blocks.reduce((count, block) => count + nativeBlockTextLength(block), 0);
}

function nativeParagraphsFromBlock(block: NativeWordBlockNode): NativeWordParagraphNode[] {
  if (block.kind === 'paragraph') return [block];
  return block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks.flatMap(nativeParagraphsFromBlock)));
}

function nativeBlockTextLength(block: NativeWordBlockNode): number {
  if (block.kind === 'paragraph') {
    return block.content.reduce((count, inline) => count + (inline.kind === 'text' ? inline.text.length : 0), 0);
  }
  return block.rows.reduce(
    (rowCount, row) => rowCount + row.cells.reduce((cellCount, cell) => cellCount + nativeCellTextLength(cell), 0),
    0,
  );
}
