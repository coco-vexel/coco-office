import type {
  EditorBlockView,
  EditorDocumentView,
  EditorDocumentDefaults,
  EditorDrawingSegment,
  EditorFieldSegment,
  EditorPageLayout,
  EditorImageSegment,
  EditorRevisionSegment,
  EditorHeaderFooterStoryView,
  EditorHyperlink,
  EditorCommentSpan,
  EditorParagraphFormat,
  EditorParagraphView,
  EditorRunFormat,
  EditorTableBorders,
  EditorTableCellView,
  EditorTableView,
  EditorTextSegment,
} from './model.js';

export interface NativeWordTextNode {
  kind: 'text';
  text: string;
  format: EditorRunFormat | null;
  marks?: NativeWordTextMark[];
  formatRevision?: NativeWordFormatRevisionMark;
  hyperlink?: EditorHyperlink;
}

export type NativeWordTextMark = NativeWordCommentMark | NativeWordRevisionMark;

export interface NativeWordFormatRevisionMark {
  kind: 'format';
  revisionId: string;
  beforeFormat: EditorRunFormat | null;
  author?: string;
  date?: string;
}

export interface NativeWordCommentMark {
  kind: 'comment';
  id: string;
  commentId: number;
  parentCommentId?: number;
  author: string;
  date: string;
  body: string;
}

export interface NativeWordRevisionMark {
  kind: 'revision';
  id: string;
  revisionId: string;
  revisionKind: EditorRevisionSegment['revisionKind'];
  author?: string;
  date?: string;
}

export interface NativeWordImageNode {
  kind: 'image';
  drawingId: string;
  src: string;
  revision?: NativeWordRevisionMark;
  formatRevision?: {
    kind: 'format';
    revisionId: string;
    beforeWidth?: number;
    beforeHeight?: number;
    afterWidth?: number;
    afterHeight?: number;
    author?: string;
    date?: string;
  };
  width?: number;
  height?: number;
  anchoringMode?: EditorImageSegment['anchoringMode'];
  wrap?: EditorImageSegment['wrap'];
  position?: EditorImageSegment['position'];
}

export interface NativeWordDrawingNode {
  kind: 'drawing';
  drawingId: string;
  drawingKind: EditorDrawingSegment['drawingKind'];
  text?: string;
  width?: number;
  height?: number;
  anchoringMode?: EditorDrawingSegment['anchoringMode'];
  wrap?: EditorDrawingSegment['wrap'];
  position?: EditorDrawingSegment['position'];
  hasBorder?: boolean;
  hasFill?: boolean;
  isLineOrConnector?: boolean;
  innerTextAlign?: string;
  paragraphs?: NativeWordParagraphNode[];
}

export interface NativeWordPageBreakNode {
  kind: 'pageBreak';
}

export interface NativeWordFieldNode {
  kind: 'field';
  fieldKind: EditorFieldSegment['fieldKind'];
  instruction: string;
  cachedText?: string;
  format: EditorRunFormat | null;
}

export type NativeWordInlineNode =
  | NativeWordTextNode
  | NativeWordImageNode
  | NativeWordDrawingNode
  | NativeWordFieldNode
  | NativeWordPageBreakNode;

export interface NativeWordParagraphNode {
  kind: 'paragraph';
  paraId: string;
  story: string;
  styleId: string | null;
  format: EditorParagraphFormat | null;
  content: NativeWordInlineNode[];
}

export interface NativeWordTableNode {
  kind: 'table';
  tableId: string;
  revision?: NativeWordRevisionMark;
  tableStyleId?: string;
  columnWidths: number[];
  layout?: EditorTableView['layout'];
  borders?: EditorTableBorders;
  rows: NativeWordTableRowNode[];
}

export interface NativeWordTableRowNode {
  height?: {
    val: number;
    hRule: 'auto' | 'atLeast' | 'exact';
  };
  isHeader?: boolean;
  cantSplit?: boolean;
  cells: NativeWordTableCellNode[];
}

export interface NativeWordTableCellNode {
  blocks: NativeWordBlockNode[];
  paragraphs: NativeWordParagraphNode[];
  gridSpan: number;
  rowSpan?: number;
  vMerge?: EditorTableCellView['vMerge'];
  width?: number;
  margin?: EditorTableCellView['margin'];
  borders?: EditorTableCellView['borders'];
  shading?: EditorTableCellView['shading'];
  verticalAlignment?: EditorTableCellView['verticalAlignment'];
  noWrap?: EditorTableCellView['noWrap'];
}

export type NativeWordBlockNode = NativeWordParagraphNode | NativeWordTableNode;

export interface NativeWordDocument {
  kind: 'word';
  defaults?: EditorDocumentDefaults;
  pageLayout?: EditorPageLayout;
  sections: EditorDocumentView['sections'];
  headers: NativeWordHeaderFooterStory[];
  footers: NativeWordHeaderFooterStory[];
  blocks: NativeWordBlockNode[];
}

export interface NativeWordHeaderFooterStory {
  kind: EditorHeaderFooterStoryView['kind'];
  partName: string;
  blocks: NativeWordBlockNode[];
}

export interface NativeParagraphLocation {
  paragraph: NativeWordParagraphNode;
  previous: NativeWordParagraphNode | null;
  container: NativeWordBlockNode[];
  index: number;
  insertAfter(paragraph: NativeWordParagraphNode): void;
  remove(): void;
}

export interface NativeTableCellLocation {
  table: NativeWordTableNode;
  container: NativeWordBlockNode[];
  tableIndex: number;
  row: NativeWordTableRowNode;
  rowIndex: number;
  cell: NativeWordTableCellNode;
  cellIndex: number;
  columnIndex: number;
}

export interface NativeTableLocation {
  table: NativeWordTableNode;
  container: NativeWordBlockNode[];
  tableIndex: number;
}

export function createNativeWordDocument(view: EditorDocumentView): NativeWordDocument {
  return {
    kind: 'word',
    ...(view.defaults ? { defaults: view.defaults } : {}),
    ...(view.pageLayout ? { pageLayout: view.pageLayout } : {}),
    sections: view.sections,
    headers: view.headers.map(nativeHeaderFooterStoryFromEditorStory),
    footers: view.footers.map(nativeHeaderFooterStoryFromEditorStory),
    blocks: view.blocks.map((block) => nativeBlockFromEditorBlock(block)),
  };
}

export function findNativeParagraph(
  doc: NativeWordDocument,
  paraId: string,
): NativeWordParagraphNode | null {
  return findNativeParagraphLocation(doc, paraId)?.paragraph ?? null;
}

export function previousNativeParagraphId(doc: NativeWordDocument, paraId: string): string | null {
  const location = findNativeParagraphLocation(doc, paraId);
  return location?.previous?.paraId ?? null;
}

export function findNativeParagraphLocation(
  doc: NativeWordDocument,
  paraId: string,
): NativeParagraphLocation | null {
  return (
    findNativeParagraphLocationInBlocks(doc.blocks, paraId) ??
    doc.headers.reduce<NativeParagraphLocation | null>(
      (found, story) => found ?? findNativeParagraphLocationInBlocks(story.blocks, paraId),
      null,
    ) ??
    doc.footers.reduce<NativeParagraphLocation | null>(
      (found, story) => found ?? findNativeParagraphLocationInBlocks(story.blocks, paraId),
      null,
    )
  );
}

export function findNativeTableCellLocation(
  doc: NativeWordDocument,
  paraId: string,
): NativeTableCellLocation | null {
  return (
    findNativeTableCellLocationInBlocks(doc.blocks, paraId) ??
    doc.headers.reduce<NativeTableCellLocation | null>(
      (found, story) => found ?? findNativeTableCellLocationInBlocks(story.blocks, paraId),
      null,
    ) ??
    doc.footers.reduce<NativeTableCellLocation | null>(
      (found, story) => found ?? findNativeTableCellLocationInBlocks(story.blocks, paraId),
      null,
    )
  );
}

export function findNativeTableLocation(
  doc: NativeWordDocument,
  tableId: string,
): NativeTableLocation | null {
  return (
    findNativeTableLocationInBlocks(doc.blocks, tableId) ??
    doc.headers.reduce<NativeTableLocation | null>(
      (found, story) => found ?? findNativeTableLocationInBlocks(story.blocks, tableId),
      null,
    ) ??
    doc.footers.reduce<NativeTableLocation | null>(
      (found, story) => found ?? findNativeTableLocationInBlocks(story.blocks, tableId),
      null,
    )
  );
}

function findNativeTableLocationInBlocks(
  blocks: NativeWordBlockNode[],
  tableId: string,
): NativeTableLocation | null {
  for (let tableIndex = 0; tableIndex < blocks.length; tableIndex += 1) {
    const block = blocks[tableIndex];
    if (block?.kind !== 'table') continue;
    if (block.tableId === tableId) return { table: block, container: blocks, tableIndex };
    for (const row of block.rows) {
      for (const cell of row.cells) {
        const nested = findNativeTableLocationInBlocks(cell.blocks, tableId);
        if (nested) return nested;
      }
    }
  }
  return null;
}

function findNativeTableCellLocationInBlocks(
  blocks: NativeWordBlockNode[],
  paraId: string,
): NativeTableCellLocation | null {
  for (let tableIndex = 0; tableIndex < blocks.length; tableIndex += 1) {
    const table = blocks[tableIndex];
    if (table?.kind !== 'table') continue;
    for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex += 1) {
      const row = table.rows[rowIndex]!;
      let columnIndex = 0;
      for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
        const cell = row.cells[cellIndex]!;
        const nested = findNativeTableCellLocationInBlocks(cell.blocks, paraId);
        if (nested) return nested;
        if (cell.blocks.some((block) => block.kind === 'paragraph' && block.paraId === paraId)) {
          return { table, container: blocks, tableIndex, row, rowIndex, cell, cellIndex, columnIndex };
        }
        columnIndex += Math.max(1, cell.gridSpan);
      }
    }
  }
  return null;
}

function findNativeParagraphLocationInBlocks(
  blocks: NativeWordBlockNode[],
  paraId: string,
): NativeParagraphLocation | null {
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    if (block?.kind === 'paragraph' && block.paraId === paraId) {
      const previousBlock = blocks[index - 1];
      return {
        paragraph: block,
        previous: previousBlock?.kind === 'paragraph' ? previousBlock : null,
        container: blocks,
        index,
        insertAfter(paragraph) {
          blocks.splice(index + 1, 0, paragraph);
        },
        remove() {
          blocks.splice(index, 1);
        },
      };
    }
    if (block?.kind === 'table') {
      const tableLocation = findTableParagraphLocation(block, paraId);
      if (tableLocation) return tableLocation;
    }
  }
  return null;
}

export function allNativeParagraphs(doc: NativeWordDocument): NativeWordParagraphNode[] {
  return [
    ...collectNativeParagraphsFromBlocks(doc.blocks),
    ...doc.headers.flatMap((story) => collectNativeParagraphsFromBlocks(story.blocks)),
    ...doc.footers.flatMap((story) => collectNativeParagraphsFromBlocks(story.blocks)),
  ];
}

function collectNativeParagraphsFromBlocks(blocks: NativeWordBlockNode[]): NativeWordParagraphNode[] {
  return blocks.flatMap((block) => {
    if (block.kind === 'paragraph') return [block];
    return block.rows.flatMap((row) => row.cells.flatMap((cell) => collectNativeParagraphsFromBlocks(cell.blocks)));
  });
}

export function nativeParagraphText(paragraph: NativeWordParagraphNode): string {
  return paragraph.content
    .map((child) => {
      if (child.kind === 'text') return child.text;
      if (child.kind === 'field') return child.cachedText ?? '\uFFFC';
      if (child.kind === 'drawing' && child.text) return child.text;
      return '\uFFFC';
    })
    .join('');
}

export function createNativeParaId(doc: NativeWordDocument): string {
  const seen = new Set(allNativeParagraphs(doc).map((paragraph) => paragraph.paraId).filter(Boolean));
  let paraId = '';
  do {
    paraId = Math.floor(Math.random() * 0xffffffff)
      .toString(16)
      .toUpperCase()
      .padStart(8, '0');
  } while (!paraId || seen.has(paraId));
  return paraId;
}

function nativeBlockFromEditorBlock(block: EditorBlockView): NativeWordBlockNode {
  if (block.type === 'table') {
    return {
      kind: 'table',
      tableId: block.tableId,
      ...(block.tableStyleId ? { tableStyleId: block.tableStyleId } : {}),
      columnWidths: block.columnWidths,
      ...(block.layout ? { layout: block.layout } : {}),
      ...(block.borders ? { borders: block.borders } : {}),
      rows: block.rows.map((row) => ({
        ...(row.height ? { height: row.height } : {}),
        ...(row.isHeader ? { isHeader: true } : {}),
        ...(row.cantSplit ? { cantSplit: true } : {}),
        cells: row.cells.map((cell) => nativeTableCellFromEditorCell(cell)),
      })),
    };
  }
  return nativeParagraphFromEditorParagraph(block);
}

function nativeHeaderFooterStoryFromEditorStory(
  story: EditorHeaderFooterStoryView,
): NativeWordHeaderFooterStory {
  return {
    kind: story.kind,
    partName: story.partName,
    blocks: story.blocks.map((block) => nativeBlockFromEditorBlock(block)),
  };
}

function nativeTableCellFromEditorCell(cell: EditorTableCellView): NativeWordTableCellNode {
  const cellBlocks = cell.blocks ?? cell.paragraphs;
  const blocks = cellBlocks.map((block) => nativeBlockFromEditorBlock(block));
  const nativeCell: NativeWordTableCellNode = {
    gridSpan: cell.gridSpan,
    blocks,
    paragraphs: collectNativeParagraphsFromBlocks(blocks),
  };
  if (cell.rowSpan !== undefined) nativeCell.rowSpan = cell.rowSpan;
  if (cell.vMerge !== undefined) nativeCell.vMerge = cell.vMerge;
  if (cell.width !== undefined) nativeCell.width = cell.width;
  if (cell.margin !== undefined) nativeCell.margin = cell.margin;
  if (cell.borders !== undefined) nativeCell.borders = cell.borders;
  if (cell.shading !== undefined) nativeCell.shading = cell.shading;
  if (cell.verticalAlignment !== undefined) nativeCell.verticalAlignment = cell.verticalAlignment;
  if (cell.noWrap !== undefined) nativeCell.noWrap = cell.noWrap;
  return nativeCell;
}

function nativeParagraphFromEditorParagraph(paragraph: EditorParagraphView): NativeWordParagraphNode {
  const commentSpans = paragraph.spans.filter((span): span is EditorCommentSpan => span.kind === 'comment');
  return {
    kind: 'paragraph',
    paraId: paragraph.paraId,
    story: paragraph.story,
    styleId: paragraph.styleId ?? null,
    format: paragraph.format ?? null,
    content: paragraph.segments.flatMap((segment): NativeWordInlineNode[] => {
      if (segment.kind === 'text') return textNodeFromSegment(segment, commentSpans);
      if (segment.kind === 'revision') return textNodeFromRevisionSegment(segment);
      if (segment.kind === 'field') return fieldNodeFromSegment(segment);
      if (segment.kind === 'image') return imageNodeFromSegment(segment);
      if (segment.kind === 'drawing') return drawingNodeFromSegment(segment);
      if (segment.kind === 'placeholder' && segment.placeholderKind === 'pageBreak') {
        return [{ kind: 'pageBreak' }];
      }
      return [];
    }),
  };
}

function fieldNodeFromSegment(segment: EditorFieldSegment): NativeWordFieldNode[] {
  return [
    {
      kind: 'field',
      fieldKind: segment.fieldKind,
      instruction: segment.instruction,
      ...(segment.cachedText !== undefined ? { cachedText: segment.cachedText } : {}),
      format: segment.format ?? null,
    },
  ];
}

function textNodeFromRevisionSegment(segment: EditorRevisionSegment): NativeWordTextNode[] {
  if (segment.text.length === 0) return [];
  return [
    {
      kind: 'text',
      text: segment.text,
      format: segment.format ?? null,
      marks: [
        {
          kind: 'revision',
          id: `revision:${segment.revisionId}`,
          revisionId: segment.revisionId,
          revisionKind: segment.revisionKind,
          author: segment.author,
          date: segment.date,
        },
      ],
    },
  ];
}

function textNodeFromSegment(
  segment: EditorTextSegment,
  commentSpans: EditorCommentSpan[],
): NativeWordTextNode[] {
  if (segment.text.length === 0) return [];
  const nodes: NativeWordTextNode[] = [];
  let cursor = segment.from;
  for (const span of commentSpans) {
    const from = Math.max(segment.from, span.from);
    const to = Math.min(segment.to, span.to);
    if (to <= from) continue;
    if (from > cursor) {
      nodes.push({
        kind: 'text',
        text: segment.text.slice(cursor - segment.from, from - segment.from),
        format: segment.format ?? null,
        ...(segment.hyperlink ? { hyperlink: { ...segment.hyperlink } } : {}),
      });
    }
    nodes.push({
      kind: 'text',
      text: segment.text.slice(from - segment.from, to - segment.from),
      format: segment.format ?? null,
      ...(segment.hyperlink ? { hyperlink: { ...segment.hyperlink } } : {}),
      marks: [
        {
          kind: 'comment',
          id: span.id,
          commentId: span.commentId,
          author: span.author,
          date: span.date,
          body: span.body,
        },
      ],
    });
    cursor = to;
  }
  if (cursor < segment.to) {
    nodes.push({
      kind: 'text',
      text: segment.text.slice(cursor - segment.from),
      format: segment.format ?? null,
      ...(segment.hyperlink ? { hyperlink: { ...segment.hyperlink } } : {}),
    });
  }
  return nodes;
}

function imageNodeFromSegment(segment: EditorImageSegment): NativeWordImageNode[] {
  return [
    {
      kind: 'image',
      drawingId: segment.drawingId,
      src: segment.src,
      ...(segment.width !== undefined ? { width: segment.width } : {}),
      ...(segment.height !== undefined ? { height: segment.height } : {}),
      ...(segment.anchoringMode !== undefined ? { anchoringMode: segment.anchoringMode } : {}),
      ...(segment.wrap !== undefined ? { wrap: segment.wrap } : {}),
      ...(segment.position !== undefined ? { position: segment.position } : {}),
    },
  ];
}

function drawingNodeFromSegment(segment: EditorDrawingSegment): NativeWordDrawingNode[] {
  const node: NativeWordDrawingNode = {
    kind: 'drawing',
    drawingId: segment.drawingId,
    drawingKind: segment.drawingKind,
  };
  if (segment.text !== undefined) node.text = segment.text;
  if (segment.width !== undefined) node.width = segment.width;
  if (segment.height !== undefined) node.height = segment.height;
  if (segment.anchoringMode !== undefined) node.anchoringMode = segment.anchoringMode;
  if (segment.wrap !== undefined) node.wrap = segment.wrap;
  if (segment.position !== undefined) node.position = segment.position;
  if (segment.hasBorder !== undefined) node.hasBorder = segment.hasBorder;
  if (segment.hasFill !== undefined) node.hasFill = segment.hasFill;
  if (segment.isLineOrConnector !== undefined) node.isLineOrConnector = segment.isLineOrConnector;
  if (segment.innerTextAlign !== undefined) node.innerTextAlign = segment.innerTextAlign;
  if (segment.paragraphs && segment.paragraphs.length > 0) {
    node.paragraphs = segment.paragraphs.map((paragraph) => nativeParagraphFromEditorParagraph(paragraph));
  }
  return [node];
}

function findTableParagraphLocation(
  table: NativeWordTableNode,
  paraId: string,
): NativeParagraphLocation | null {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const paragraphIndex = cell.paragraphs.findIndex((paragraph) => paragraph.paraId === paraId);
      const blockIndex = cell.blocks.findIndex((block) => block.kind === 'paragraph' && block.paraId === paraId);
      if (paragraphIndex >= 0 && blockIndex >= 0) {
        return {
          paragraph: cell.paragraphs[paragraphIndex]!,
          previous: cell.paragraphs[paragraphIndex - 1] ?? null,
          container: cell.blocks,
          index: blockIndex,
          insertAfter(paragraph) {
            cell.paragraphs.splice(paragraphIndex + 1, 0, paragraph);
            cell.blocks.splice(blockIndex + 1, 0, paragraph);
          },
          remove() {
            cell.paragraphs.splice(paragraphIndex, 1);
            cell.blocks.splice(blockIndex, 1);
          },
        };
      }
    }
  }
  return null;
}
