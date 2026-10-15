import {
  createNativeParaId,
  allNativeParagraphs,
  findNativeParagraph,
  findNativeParagraphLocation,
  findNativeTableCellLocation,
  findNativeTableLocation,
  nativeParagraphText,
  type NativeWordDocument,
  type NativeWordBlockNode,
  type NativeWordInlineNode,
  type NativeWordParagraphNode,
  type NativeWordTableCellNode,
  type NativeWordTableNode,
  type NativeWordTableRowNode,
} from './native-model.js';
import type { EditorHyperlink, EditorNumberingFormat, EditorParagraphFormat, EditorRunFormat, EditorTableBorders } from './model.js';
import { nativeListIndent } from './native-numbering.js';

export type NativeWordChange =
  | {
      kind: 'text';
      op: 'insert' | 'delete' | 'replace';
      paraId: string;
      offset: number;
      text: string;
      deletedText?: string;
    }
  | {
      kind: 'block';
      op: 'split' | 'merge';
      paraId: string;
      newParaId?: string;
      offset?: number;
    }
  | {
      kind: 'structure';
      op: 'delete-range';
      startParaId: string;
      endParaId: string;
      deletedBlockIds: string[];
    }
  | {
      kind: 'structure';
      op: 'insert-range';
      paraId: string;
      insertedBlockIds: string[];
    }
  | {
      kind: 'image';
      op: 'insert' | 'delete' | 'resize' | 'accept' | 'reject';
      paraId: string;
      offset: number;
      drawingId: string;
      src?: string;
      width?: number;
      height?: number;
      revisionId?: string;
      revisionKind?: 'insert' | 'delete' | 'format';
    }
  | {
      kind: 'format';
      op: 'run' | 'accept' | 'reject';
      paraId: string;
      from: number;
      to: number;
      format: Partial<EditorRunFormat>;
      clear?: boolean;
      revisionId?: string;
      revisionKind?: 'format';
      beforeFormat?: Partial<EditorRunFormat> | null;
    }
  | {
      kind: 'format';
      op: 'paragraph';
      paraId: string;
      format: Partial<EditorParagraphFormat>;
    }
  | {
      kind: 'list';
      paraId: string;
      numbering: EditorNumberingFormat | null;
    }
  | {
      kind: 'hyperlink';
      op: 'set' | 'remove';
      paraId: string;
      from: number;
      to: number;
      hyperlink?: EditorHyperlink;
    }
  | {
      kind: 'annotation';
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
      kind: 'page-break';
      op: 'insert' | 'delete';
      paraId: string;
      offset: number;
    }
  | {
      kind: 'table';
      op: 'insert-table' | 'delete-table' | 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column' | 'merge-cells' | 'split-cell' | 'format-cell' | 'clear-cell' | 'replace-cell';
      tableId: string;
      rowIndex: number;
      columnIndex: number;
      endColumnIndex?: number;
      endRowIndex?: number;
      format?: {
        backgroundColor?: string | null;
        verticalAlignment?: NativeWordTableCellNode['verticalAlignment'] | null;
        borderColor?: string | null;
      };
      text?: string;
      cellSnapshot?: NativeWordTableCellNode;
      revisionId?: string;
      revisionKind?: 'insert' | 'delete';
      revisionAction?: 'accept' | 'reject';
      beforeTable?: NativeWordTableNode;
      anchorParaId?: string;
      anchorOffset?: number;
      tableSnapshot?: NativeWordTableNode;
    };

export interface NativePlainTextReplacement {
  changes: NativeWordChange[];
  cursor: { paraId: string; offset: number };
  paragraphIds: string[];
}

export interface NativeBlockRangeReplacement {
  changes: NativeWordChange[];
  cursor: { paraId: string; offset: number };
  container: NativeWordBlockNode[];
  startIndex: number;
  beforeBlocks: NativeWordBlockNode[];
  afterBlocks: NativeWordBlockNode[];
}

export interface NativeTableEditResult {
  change: NativeWordChange;
  cursor: { paraId: string; offset: number };
}

export interface NativeTableCellPoint {
  rowIndex: number;
  columnIndex: number;
}

export interface NativeTableCellRange {
  anchor: NativeTableCellPoint;
  focus: NativeTableCellPoint;
}

export interface NativeTableRangeEditResult {
  changes: NativeWordChange[];
  cursor: { paraId: string; offset: number };
}

export interface NativeTableClipboardCell {
  blocks: NativeWordBlockNode[];
  format: {
    margin?: NativeWordTableCellNode['margin'];
    borders?: NativeWordTableCellNode['borders'];
    shading?: NativeWordTableCellNode['shading'];
    verticalAlignment?: NativeWordTableCellNode['verticalAlignment'];
    noWrap?: NativeWordTableCellNode['noWrap'];
  };
}

export type NativeTableStructureOperation =
  | { kind: 'insert-row'; position: 'before' | 'after' }
  | { kind: 'delete-row' }
  | { kind: 'insert-column'; position: 'before' | 'after' }
  | { kind: 'delete-column' };

export interface NativeTableCellFormatPatch {
  backgroundColor?: string | null;
  verticalAlignment?: NativeWordTableCellNode['verticalAlignment'] | null;
  borderColor?: string | null;
}

export function cloneNativeWordBlocks(blocks: readonly NativeWordBlockNode[]): NativeWordBlockNode[] {
  return cloneNativeBlocks(blocks);
}

export function nativeTableSupportsStructuralEditing(table: NativeWordTableNode): boolean {
  if (table.rows.length === 0) return false;
  const columnCount = table.rows[0]?.cells.length ?? 0;
  return columnCount > 0 && table.rows.every(
    (row) =>
      row.cells.length === columnCount &&
      row.cells.every(
        (cell) =>
          cell.gridSpan === 1 &&
          (cell.rowSpan === undefined || cell.rowSpan === 1) &&
          cell.vMerge == null,
      ),
  );
}

export function nativeInsertTable(
  doc: NativeWordDocument,
  paraId: string,
  offset: number,
  rowCount: number,
  columnCount: number,
  createParaId: () => string = () => createNativeParaId(doc),
  createTableId: () => string = () => freshNativeTableId(doc),
): NativeBlockRangeReplacement | null {
  const location = findNativeParagraphLocation(doc, paraId);
  if (!location || location.container !== doc.blocks) return null;
  const rows = Math.max(1, Math.min(20, Math.trunc(rowCount)));
  const columns = Math.max(1, Math.min(20, Math.trunc(columnCount)));
  if (!Number.isFinite(rowCount) || !Number.isFinite(columnCount)) return null;

  const paragraph = location.paragraph;
  const beforeBlocks = cloneNativeBlocks([paragraph]);
  const splitOffset = Math.max(0, Math.min(offset, nativeParagraphText(paragraph).length));
  const split = splitInlineContentAt(paragraph.content, splitOffset);
  paragraph.content = split.before;
  const trailingParagraph: NativeWordParagraphNode = {
    kind: 'paragraph',
    paraId: createParaId(),
    story: paragraph.story,
    styleId: paragraph.styleId,
    format: paragraph.format ? cloneJson(paragraph.format) : null,
    content: split.after,
  };
  const tableId = createTableId();
  const availableWidth = Math.max(
    1440,
    (doc.pageLayout?.width ?? 11906) -
      (doc.pageLayout?.margin.left ?? 1440) -
      (doc.pageLayout?.margin.right ?? 1440),
  );
  const columnWidth = Math.max(360, Math.floor(availableWidth / columns));
  const table: NativeWordTableNode = {
    kind: 'table',
    tableId,
    columnWidths: Array.from({ length: columns }, () => columnWidth),
    layout: { width: columnWidth * columns, widthType: 'dxa', alignment: 'left' },
    borders: visibleCellBorders('000000'),
    rows: Array.from({ length: rows }, () => ({
      cells: Array.from({ length: columns }, () => {
        const cellParagraph: NativeWordParagraphNode = {
          kind: 'paragraph',
          paraId: createParaId(),
          story: paragraph.story,
          styleId: null,
          format: null,
          content: [{ kind: 'text', text: '', format: null }],
        };
        return {
          gridSpan: 1,
          width: columnWidth,
          borders: visibleCellBorders('000000'),
          blocks: [cellParagraph],
          paragraphs: [cellParagraph],
        };
      }),
    })),
  };
  const replacement: NativeWordBlockNode[] = [paragraph, table, trailingParagraph];
  location.container.splice(location.index, 1, ...replacement);
  return {
    changes: [{ kind: 'table', op: 'insert-table', tableId, rowIndex: 0, columnIndex: 0 }],
    cursor: { paraId: table.rows[0]!.cells[0]!.paragraphs[0]!.paraId, offset: 0 },
    container: location.container,
    startIndex: location.index,
    beforeBlocks,
    afterBlocks: cloneNativeBlocks(replacement),
  };
}

export function nativeDeleteTable(
  doc: NativeWordDocument,
  tableId: string,
  createParaId: () => string = () => createNativeParaId(doc),
): NativeBlockRangeReplacement | null {
  const location = findNativeTableLocation(doc, tableId);
  if (!location || location.container !== doc.blocks) return null;
  const firstParagraph = location.table.rows[0]?.cells[0]?.paragraphs[0];
  const paragraph: NativeWordParagraphNode = {
    kind: 'paragraph',
    paraId: createParaId(),
    story: firstParagraph?.story ?? 'word/document.xml',
    styleId: null,
    format: null,
    content: [{ kind: 'text', text: '', format: null }],
  };
  const beforeBlocks = cloneNativeBlocks([location.table]);
  location.container.splice(location.tableIndex, 1, paragraph);
  return {
    changes: [{ kind: 'table', op: 'delete-table', tableId, rowIndex: 0, columnIndex: 0 }],
    cursor: { paraId: paragraph.paraId, offset: 0 },
    container: location.container,
    startIndex: location.tableIndex,
    beforeBlocks,
    afterBlocks: cloneNativeBlocks([paragraph]),
  };
}

export function nativeEditTableStructure(
  doc: NativeWordDocument,
  paraId: string,
  operation: NativeTableStructureOperation,
  createParaId: () => string = () => createNativeParaId(doc),
): NativeTableEditResult | null {
  const location = findNativeTableCellLocation(doc, paraId);
  if (!location || !nativeTableSupportsStructuralEditing(location.table)) return null;
  const { table, rowIndex, cellIndex } = location;
  if (operation.kind === 'insert-row') {
    const insertAt = rowIndex + (operation.position === 'after' ? 1 : 0);
    const row = blankTableRow(location.row, createParaId);
    table.rows.splice(insertAt, 0, row);
    return {
      change: { kind: 'table', op: 'insert-row', tableId: table.tableId, rowIndex: insertAt, columnIndex: cellIndex },
      cursor: { paraId: firstCellParagraph(row.cells[cellIndex] ?? row.cells[0]!).paraId, offset: 0 },
    };
  }
  if (operation.kind === 'delete-row') {
    if (table.rows.length <= 1) return null;
    table.rows.splice(rowIndex, 1);
    const nextRow = table.rows[Math.min(rowIndex, table.rows.length - 1)]!;
    return {
      change: { kind: 'table', op: 'delete-row', tableId: table.tableId, rowIndex, columnIndex: cellIndex },
      cursor: { paraId: firstCellParagraph(nextRow.cells[Math.min(cellIndex, nextRow.cells.length - 1)]!).paraId, offset: 0 },
    };
  }
  if (operation.kind === 'insert-column') {
    const insertAt = cellIndex + (operation.position === 'after' ? 1 : 0);
    for (const row of table.rows) {
      const template = row.cells[Math.min(cellIndex, row.cells.length - 1)]!;
      row.cells.splice(insertAt, 0, blankTableCell(template, createParaId));
    }
    const width = table.columnWidths[Math.min(cellIndex, table.columnWidths.length - 1)] ?? 0;
    table.columnWidths.splice(insertAt, 0, width);
    const targetRow = table.rows[rowIndex]!;
    return {
      change: { kind: 'table', op: 'insert-column', tableId: table.tableId, rowIndex, columnIndex: insertAt },
      cursor: { paraId: firstCellParagraph(targetRow.cells[insertAt]!).paraId, offset: 0 },
    };
  }
  if (table.rows[0]!.cells.length <= 1) return null;
  for (const row of table.rows) row.cells.splice(cellIndex, 1);
  if (cellIndex < table.columnWidths.length) table.columnWidths.splice(cellIndex, 1);
  const targetRow = table.rows[rowIndex]!;
  const nextCell = targetRow.cells[Math.min(cellIndex, targetRow.cells.length - 1)]!;
  return {
    change: { kind: 'table', op: 'delete-column', tableId: table.tableId, rowIndex, columnIndex: cellIndex },
    cursor: { paraId: firstCellParagraph(nextCell).paraId, offset: 0 },
  };
}

export function nativeSetTableCellFormat(
  doc: NativeWordDocument,
  paraId: string,
  patch: NativeTableCellFormatPatch,
): NativeTableEditResult | null {
  const location = findNativeTableCellLocation(doc, paraId);
  if (!location) return null;
  const before = JSON.stringify({
    shading: location.cell.shading,
    verticalAlignment: location.cell.verticalAlignment,
    borders: location.cell.borders,
  });
  const changeFormat: NonNullable<Extract<NativeWordChange, { kind: 'table' }>['format']> = {};
  if (patch.backgroundColor !== undefined) {
    const color = normalizeTableColor(patch.backgroundColor);
    if (patch.backgroundColor !== null && !color) return null;
    if (color) location.cell.shading = { fill: color, val: 'clear' };
    else delete location.cell.shading;
    changeFormat.backgroundColor = color ? `#${color}` : null;
  }
  if (patch.verticalAlignment !== undefined) {
    if (patch.verticalAlignment) location.cell.verticalAlignment = patch.verticalAlignment;
    else delete location.cell.verticalAlignment;
    changeFormat.verticalAlignment = patch.verticalAlignment ?? null;
  }
  if (patch.borderColor !== undefined) {
    const color = normalizeTableColor(patch.borderColor);
    if (patch.borderColor !== null && !color) return null;
    location.cell.borders = color ? visibleCellBorders(color) : hiddenCellBorders();
    changeFormat.borderColor = color ? `#${color}` : null;
  }
  const after = JSON.stringify({
    shading: location.cell.shading,
    verticalAlignment: location.cell.verticalAlignment,
    borders: location.cell.borders,
  });
  if (before === after) return null;
  return {
    change: {
      kind: 'table',
      op: 'format-cell',
      tableId: location.table.tableId,
      rowIndex: location.rowIndex,
      columnIndex: location.columnIndex,
      format: changeFormat,
    },
    cursor: { paraId, offset: 0 },
  };
}

export function nativeTableCellsInRange(
  table: NativeWordTableNode,
  range: NativeTableCellRange,
): Array<{ cell: NativeWordTableCellNode; rowIndex: number; columnIndex: number }> {
  const startRow = Math.max(0, Math.min(range.anchor.rowIndex, range.focus.rowIndex));
  const endRow = Math.min(table.rows.length - 1, Math.max(range.anchor.rowIndex, range.focus.rowIndex));
  const startColumn = Math.max(0, Math.min(range.anchor.columnIndex, range.focus.columnIndex));
  const endColumn = Math.max(range.anchor.columnIndex, range.focus.columnIndex);
  const cells: Array<{ cell: NativeWordTableCellNode; rowIndex: number; columnIndex: number }> = [];
  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
    const row = table.rows[rowIndex];
    if (!row) continue;
    let columnIndex = 0;
    for (const cell of row.cells) {
      const span = Math.max(1, cell.gridSpan);
      const cellEnd = columnIndex + span - 1;
      if (columnIndex <= endColumn && cellEnd >= startColumn) cells.push({ cell, rowIndex, columnIndex });
      columnIndex += span;
    }
  }
  return cells;
}

export function nativeMergeTableCells(
  doc: NativeWordDocument,
  tableId: string,
  range: NativeTableCellRange,
): { change: NativeWordChange; cursor: { paraId: string; offset: number } } | null {
  const location = findNativeTableLocation(doc, tableId);
  if (!location) return null;
  const startRow = Math.min(range.anchor.rowIndex, range.focus.rowIndex);
  const endRow = Math.max(range.anchor.rowIndex, range.focus.rowIndex);
  const startColumn = Math.min(range.anchor.columnIndex, range.focus.columnIndex);
  const endColumn = Math.max(range.anchor.columnIndex, range.focus.columnIndex);
  if (startRow < 0 || startRow >= location.table.rows.length || endRow >= location.table.rows.length) return null;
  if (startRow !== endRow && startColumn !== endColumn) return null;
  const row = location.table.rows[startRow];
  if (!row) return null;
  if (startRow !== endRow) {
    const selected = location.table.rows.slice(startRow, endRow + 1).map((candidate) => candidate.cells[startColumn]);
    if (selected.some((cell) => !cell || cell.gridSpan !== 1 || cell.vMerge != null)) return null;
    const selectedCells = selected as NativeWordTableCellNode[];
    const first = selectedCells[0]!;
    const mergedBlocks = selectedCells.flatMap((cell) => cell.blocks.map(cloneJson));
    const mergedParagraphs = mergedBlocks.flatMap((block) => block.kind === 'paragraph' ? [block] : collectParagraphs([block]));
    first.blocks = mergedBlocks;
    first.paragraphs = mergedParagraphs;
    first.rowSpan = selectedCells.length;
    first.vMerge = 'restart';
    for (const cell of selectedCells.slice(1)) {
      const paragraph = cell!.paragraphs[0];
      if (paragraph) paragraph.content = [{ kind: 'text', text: '', format: null }];
      cell!.blocks = paragraph ? [paragraph] : [];
      cell!.paragraphs = paragraph ? [paragraph] : [];
      cell!.rowSpan = undefined;
      cell!.vMerge = 'continue';
    }
    return {
      change: { kind: 'table', op: 'merge-cells', tableId, rowIndex: startRow, columnIndex: startColumn, endRowIndex: endRow },
      cursor: { paraId: first.paragraphs[0]!.paraId, offset: 0 },
    };
  }
  if (startColumn === endColumn) return null;
  const cells = nativeTableCellsInRange(location.table, range)
    .filter((item) => item.rowIndex === startRow);
  if (cells.length < 2 || cells.some((item) => item.cell.gridSpan !== 1 || item.cell.vMerge != null)) return null;
  const first = cells[0]!;
  const last = cells[cells.length - 1]!;
  if (first.columnIndex !== startColumn || last.columnIndex !== endColumn || cells.length !== endColumn - startColumn + 1) return null;
  const mergedBlocks = cells.flatMap(({ cell }) => cell.blocks.map(cloneJson));
  const mergedParagraphs = mergedBlocks.flatMap((block) => block.kind === 'paragraph' ? [block] : collectParagraphs([block]));
  first.cell.blocks = mergedBlocks;
  first.cell.paragraphs = mergedParagraphs;
  first.cell.gridSpan = cells.length;
  const firstIndex = row.cells.indexOf(first.cell);
  if (firstIndex < 0) return null;
  row.cells.splice(firstIndex, cells.length, first.cell);
  return {
    change: {
      kind: 'table',
      op: 'merge-cells',
      tableId,
      rowIndex: startRow,
      columnIndex: startColumn,
      endColumnIndex: endColumn,
    },
    cursor: { paraId: first.cell.paragraphs[0]!.paraId, offset: 0 },
  };
}

export function nativeSplitTableCell(
  doc: NativeWordDocument,
  tableId: string,
  rowIndex: number,
  columnIndex: number,
  createParaId: () => string = () => createNativeParaId(doc),
): { change: NativeWordChange; cursor: { paraId: string; offset: number } } | null {
  const location = findNativeTableLocation(doc, tableId);
  const row = location?.table.rows[rowIndex];
  if (!location || !row) return null;
  let logicalColumn = 0;
  const cellIndex = row.cells.findIndex((cell) => {
    const matches = logicalColumn === columnIndex;
    logicalColumn += Math.max(1, cell.gridSpan);
    return matches;
  });
  const cell = cellIndex >= 0 ? row.cells[cellIndex] : null;
  if (!cell) return null;
  if (cell.vMerge === 'restart' && (cell.rowSpan ?? 1) > 1) {
    const endRowIndex = rowIndex + (cell.rowSpan ?? 1) - 1;
    if (endRowIndex >= location.table.rows.length) return null;
    const selected = location.table.rows.slice(rowIndex, endRowIndex + 1).map((candidate) => candidate.cells[columnIndex]);
    if (selected.length !== cell.rowSpan || selected.some((candidate, index) => {
      if (!candidate || candidate.gridSpan !== 1) return true;
      return index === 0 ? candidate.vMerge !== 'restart' : candidate.vMerge !== 'continue';
    })) return null;
    cell.rowSpan = undefined;
    cell.vMerge = undefined;
    for (const candidate of selected.slice(1)) {
      candidate!.rowSpan = undefined;
      candidate!.vMerge = undefined;
    }
    return {
      change: { kind: 'table', op: 'split-cell', tableId, rowIndex, columnIndex, endRowIndex },
      cursor: { paraId: cell.paragraphs[0]!.paraId, offset: 0 },
    };
  }
  if (cell.gridSpan <= 1 || cell.vMerge != null || cell.rowSpan && cell.rowSpan > 1) return null;
  const span = cell.gridSpan;
  cell.gridSpan = 1;
  const blanks = Array.from({ length: span - 1 }, () => blankTableCell(cell, createParaId));
  row.cells.splice(cellIndex + 1, 0, ...blanks);
  return {
    change: { kind: 'table', op: 'split-cell', tableId, rowIndex, columnIndex },
    cursor: { paraId: cell.paragraphs[0]!.paraId, offset: 0 },
  };
}

export function nativeSetTableCellRangeFormat(
  doc: NativeWordDocument,
  tableId: string,
  range: NativeTableCellRange,
  patch: NativeTableCellFormatPatch,
): NativeTableRangeEditResult | null {
  const location = findNativeTableLocation(doc, tableId);
  if (!location) return null;
  const changes: NativeWordChange[] = [];
  let cursor: { paraId: string; offset: number } | null = null;
  for (const { cell } of nativeTableCellsInRange(location.table, range)) {
    const paraId = firstCellParagraph(cell).paraId;
    cursor ??= { paraId, offset: 0 };
    const result = nativeSetTableCellFormat(doc, paraId, patch);
    if (result) changes.push(result.change);
  }
  return cursor && changes.length > 0 ? { changes, cursor } : null;
}

export function nativeClearTableCellRange(
  doc: NativeWordDocument,
  tableId: string,
  range: NativeTableCellRange,
): NativeTableRangeEditResult | null {
  const location = findNativeTableLocation(doc, tableId);
  if (!location) return null;
  const changes: NativeWordChange[] = [];
  let cursor: { paraId: string; offset: number } | null = null;
  for (const { cell, rowIndex, columnIndex } of nativeTableCellsInRange(location.table, range)) {
    const paragraph = firstCellParagraph(cell);
    cursor ??= { paraId: paragraph.paraId, offset: 0 };
    const alreadyEmpty = cell.blocks.length === 1 && paragraph.content.length === 1 &&
      paragraph.content[0]?.kind === 'text' && paragraph.content[0].text.length === 0;
    if (alreadyEmpty) continue;
    paragraph.content = [{ kind: 'text', text: '', format: null }];
    cell.blocks = [paragraph];
    cell.paragraphs = [paragraph];
    changes.push({
      kind: 'table',
      op: 'clear-cell',
      tableId,
      rowIndex,
      columnIndex,
    });
  }
  return cursor && changes.length > 0 ? { changes, cursor } : null;
}

export function nativeReplaceTableCellRangeText(
  doc: NativeWordDocument,
  tableId: string,
  anchor: NativeTableCellPoint,
  values: readonly (readonly string[])[],
): NativeTableRangeEditResult | null {
  const location = findNativeTableLocation(doc, tableId);
  if (!location || values.length === 0 || values.every((row) => row.length === 0)) return null;
  const rowCount = values.length;
  const columnCount = Math.max(0, ...values.map((row) => row.length));
  const range = {
    anchor,
    focus: {
      rowIndex: anchor.rowIndex + rowCount - 1,
      columnIndex: anchor.columnIndex + columnCount - 1,
    },
  };
  const changes: NativeWordChange[] = [];
  let cursor: { paraId: string; offset: number } | null = null;
  for (const { cell, rowIndex, columnIndex } of nativeTableCellsInRange(location.table, range)) {
    const value = values[rowIndex - anchor.rowIndex]?.[columnIndex - anchor.columnIndex];
    if (value === undefined) continue;
    const paragraph = firstCellParagraph(cell);
    cursor ??= { paraId: paragraph.paraId, offset: value.length };
    const existing = cell.paragraphs.map((candidate) => nativeParagraphText(candidate)).join('\n');
    if (existing === value && cell.blocks.length === 1) continue;
    const format = paragraph.content.find((child) => child.kind === 'text')?.format ?? null;
    paragraph.content = [{ kind: 'text', text: value, format: format ? cloneJson(format) : null }];
    cell.blocks = [paragraph];
    cell.paragraphs = [paragraph];
    changes.push({ kind: 'table', op: 'replace-cell', tableId, rowIndex, columnIndex, text: value });
  }
  return cursor && changes.length > 0 ? { changes, cursor } : null;
}

export function nativeReplaceTableCellRangeContent(
  doc: NativeWordDocument,
  tableId: string,
  anchor: NativeTableCellPoint,
  values: readonly (readonly (NativeTableClipboardCell | null)[])[],
): NativeTableRangeEditResult | null {
  const location = findNativeTableLocation(doc, tableId);
  if (!location || values.length === 0 || values.every((row) => row.length === 0)) return null;
  const columnCount = Math.max(0, ...values.map((row) => row.length));
  const range = {
    anchor,
    focus: {
      rowIndex: anchor.rowIndex + values.length - 1,
      columnIndex: anchor.columnIndex + columnCount - 1,
    },
  };
  const changes: NativeWordChange[] = [];
  let cursor: { paraId: string; offset: number } | null = null;
  for (const { cell, rowIndex, columnIndex } of nativeTableCellsInRange(location.table, range)) {
    const value = values[rowIndex - anchor.rowIndex]?.[columnIndex - anchor.columnIndex];
    if (!value) continue;
    const blocks = cloneClipboardBlocksWithFreshIds(doc, value.blocks);
    const paragraphs = collectParagraphs(blocks);
    if (paragraphs.length === 0) continue;
    cell.blocks = blocks;
    cell.paragraphs = paragraphs;
    assignOptionalCellFormat(cell, 'margin', value.format.margin);
    assignOptionalCellFormat(cell, 'borders', value.format.borders);
    assignOptionalCellFormat(cell, 'shading', value.format.shading);
    assignOptionalCellFormat(cell, 'verticalAlignment', value.format.verticalAlignment);
    assignOptionalCellFormat(cell, 'noWrap', value.format.noWrap);
    cursor ??= { paraId: paragraphs[0]!.paraId, offset: nativeParagraphText(paragraphs[0]!).length };
    changes.push({ kind: 'table', op: 'replace-cell', tableId, rowIndex, columnIndex, cellSnapshot: cloneJson(cell) });
  }
  return cursor && changes.length > 0 ? { changes, cursor } : null;
}

function assignOptionalCellFormat<Key extends 'margin' | 'borders' | 'shading' | 'verticalAlignment' | 'noWrap'>(
  cell: NativeWordTableCellNode,
  key: Key,
  value: NativeWordTableCellNode[Key],
): void {
  if (value === undefined) delete cell[key];
  else cell[key] = cloneJson(value) as NativeWordTableCellNode[Key];
}

function normalizeTableColor(color: string | null): string | null {
  if (color === null) return null;
  const normalized = color.trim().replace(/^#/, '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(normalized) ? normalized : null;
}

function visibleCellBorders(color: string): EditorTableBorders {
  const border = { val: 'single', color, size: 4 };
  return { top: { ...border }, right: { ...border }, bottom: { ...border }, left: { ...border } };
}

function hiddenCellBorders(): EditorTableBorders {
  const border = { val: 'nil' };
  return { top: { ...border }, right: { ...border }, bottom: { ...border }, left: { ...border } };
}

function blankTableRow(template: NativeWordTableRowNode, createParaId: () => string): NativeWordTableRowNode {
  const { cells, ...format } = template;
  return { ...cloneJson(format), cells: cells.map((cell) => blankTableCell(cell, createParaId)) };
}

function blankTableCell(template: NativeWordTableCellNode, createParaId: () => string): NativeWordTableCellNode {
  const { blocks: _blocks, paragraphs: _paragraphs, ...format } = template;
  const source = firstCellParagraph(template);
  const paragraph: NativeWordParagraphNode = {
    kind: 'paragraph',
    paraId: createParaId(),
    story: source.story,
    styleId: source.styleId,
    format: source.format ? cloneJson(source.format) : null,
    content: [{ kind: 'text', text: '', format: null }],
  };
  return { ...cloneJson(format), gridSpan: 1, blocks: [paragraph], paragraphs: [paragraph] };
}

function firstCellParagraph(cell: NativeWordTableCellNode): NativeWordParagraphNode {
  const paragraph = cell.paragraphs[0];
  if (!paragraph) throw new Error('Native table cell must contain at least one paragraph');
  return paragraph;
}

function cloneJson<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

export function nativeInsertText(
  doc: NativeWordDocument,
  paraId: string,
  offset: number,
  text: string,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph || text.length === 0) return null;
  paragraph.content = insertInlineText(paragraph.content, offset, text);
  return { kind: 'text', op: 'insert', paraId, offset, text };
}

export function nativeInsertImage(
  doc: NativeWordDocument,
  paraId: string,
  offset: number,
  src: string,
  options: { drawingId?: string; width?: number; height?: number } = {},
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph || !src) return null;
  const drawingId = options.drawingId ?? createNativeDrawingId(doc);
  const width = positiveImageDimension(options.width);
  const height = positiveImageDimension(options.height);
  const split = splitInlineContentAt(paragraph.content, offset);
  paragraph.content = normalizeTextNodes([
    ...split.before,
    { kind: 'image', drawingId, src, ...(width ? { width } : {}), ...(height ? { height } : {}) },
    ...split.after,
  ]);
  return {
    kind: 'image',
    op: 'insert',
    paraId,
    offset,
    drawingId,
    src,
    ...(width ? { width } : {}),
    ...(height ? { height } : {}),
  };
}

export function nativeInsertPageBreak(
  doc: NativeWordDocument,
  paraId: string,
  offset: number,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph) return null;
  const boundedOffset = Math.max(0, Math.min(offset, nativeParagraphText(paragraph).length));
  const split = splitInlineContentAt(paragraph.content, boundedOffset);
  paragraph.content = normalizeTextNodes([...split.before, { kind: 'pageBreak' }, ...split.after]);
  return { kind: 'page-break', op: 'insert', paraId, offset: boundedOffset };
}

export function nativePageBreakAt(doc: NativeWordDocument, paraId: string, offset: number): boolean {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph || offset < 0) return false;
  let position = 0;
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (offset >= position && offset < position + length) return child.kind === 'pageBreak';
    position += length;
  }
  return false;
}

export function nativeRangeContainsPageBreak(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
): boolean {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return false;
  let position = 0;
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind === 'pageBreak' && range.from < position + length && range.to > position) return true;
    position += length;
  }
  return false;
}

export function nativeSetImageSize(
  doc: NativeWordDocument,
  paraId: string,
  drawingId: string,
  size: { width: number; height: number },
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const image = paragraph?.content.find(
    (child): child is Extract<NativeWordInlineNode, { kind: 'image' }> =>
      child.kind === 'image' && child.drawingId === drawingId,
  );
  const width = positiveImageDimension(size.width);
  const height = positiveImageDimension(size.height);
  if (!paragraph || !image || !width || !height) return null;
  if (image.width === width && image.height === height) return null;
  image.width = width;
  image.height = height;
  let offset = 0;
  for (const child of paragraph.content) {
    if (child === image) break;
    offset += inlineLength(child);
  }
  return { kind: 'image', op: 'resize', paraId, offset, drawingId, width, height };
}

export function nativeDeleteBackward(
  doc: NativeWordDocument,
  paraId: string,
  offset: number,
): NativeWordChange | null {
  const location = findNativeParagraphLocation(doc, paraId);
  if (!location) return null;
  const { paragraph } = location;

  if (offset > 0) {
    return deleteInlineBackward(paragraph, offset);
  }

  const previous = location.previous;
  if (!previous) return null;
  previous.content = normalizeTextNodes([...previous.content, ...paragraph.content]);
  location.remove();
  return { kind: 'block', op: 'merge', paraId: previous.paraId };
}

export function nativeDeleteForward(
  doc: NativeWordDocument,
  paraId: string,
  offset: number,
): NativeWordChange | null {
  const location = findNativeParagraphLocation(doc, paraId);
  if (!location) return null;
  const { paragraph } = location;
  if (offset < inlineContentLength(paragraph.content)) return deleteInlineAt(paragraph, offset);
  const next = location.container[location.index + 1];
  if (next?.kind !== 'paragraph') return null;
  paragraph.content = normalizeTextNodes([...paragraph.content, ...next.content]);
  location.container.splice(location.index + 1, 1);
  return { kind: 'block', op: 'merge', paraId: paragraph.paraId };
}

export function nativeDeleteRange(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
): NativeWordChange[] {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return [];
  const result = deleteInlineRange(paragraph, range.from, range.to);
  paragraph.content = result.content;
  return result.changes;
}

export function nativeDeleteTextRangeAcrossParagraphs(
  doc: NativeWordDocument,
  anchor: { paraId: string; offset: number },
  focus: { paraId: string; offset: number },
): NativeWordChange[] {
  if (anchor.paraId === focus.paraId) return nativeDeleteRange(doc, anchor.paraId, anchor.offset, focus.offset);
  const anchorLocation = findNativeParagraphLocation(doc, anchor.paraId);
  const focusLocation = findNativeParagraphLocation(doc, focus.paraId);
  if (!anchorLocation || !focusLocation) return [];
  if (anchorLocation.container !== focusLocation.container) return [];

  const range =
    anchorLocation.index < focusLocation.index
      ? {
          start: anchorLocation,
          end: focusLocation,
          startOffset: anchor.offset,
          endOffset: focus.offset,
        }
      : {
          start: focusLocation,
          end: anchorLocation,
          startOffset: focus.offset,
          endOffset: anchor.offset,
        };

  if (
    range.start.paragraph.content.some((child) => child.kind !== 'text') ||
    range.end.paragraph.content.some((child) => child.kind !== 'text')
  ) {
    return [];
  }
  if (range.start.container.slice(range.start.index, range.end.index + 1).some((block) => block.kind !== 'paragraph')) {
    return [];
  }
  for (let index = range.start.index + 1; index < range.end.index; index += 1) {
    const block = range.start.container[index];
    if (block?.kind !== 'paragraph' || block.content.some((child) => child.kind !== 'text')) return [];
  }

  const changes: NativeWordChange[] = [];
  const startText = nativeParagraphText(range.start.paragraph);
  const endText = nativeParagraphText(range.end.paragraph);
  const deleteStartOffset = Math.max(0, Math.min(range.startOffset, startText.length));
  const deleteEndOffset = Math.max(0, Math.min(range.endOffset, endText.length));
  const mergedText = `${startText.slice(0, deleteStartOffset)}${endText.slice(deleteEndOffset)}`;

  const removedParagraphs = range.end.index - range.start.index;
  for (let index = range.end.index; index > range.start.index; index -= 1) {
    const block = range.start.container[index];
    if (block?.kind !== 'paragraph') continue;
    if (block === range.end.paragraph && deleteEndOffset < endText.length) {
      changes.push({
        kind: 'text',
        op: 'delete',
        paraId: block.paraId,
        offset: 0,
        text: endText.slice(0, deleteEndOffset),
      });
    } else {
      const text = nativeParagraphText(block);
      if (text.length > 0) {
        changes.push({ kind: 'text', op: 'delete', paraId: block.paraId, offset: 0, text });
      }
    }
    findNativeParagraphLocation(doc, block.paraId)?.remove();
  }

  if (deleteStartOffset < startText.length) {
    changes.push({
      kind: 'text',
      op: 'delete',
      paraId: range.start.paragraph.paraId,
      offset: deleteStartOffset,
      text: startText.slice(deleteStartOffset),
    });
  }
  for (let index = 0; index < removedParagraphs; index += 1) {
    changes.push({ kind: 'block', op: 'merge', paraId: range.start.paragraph.paraId });
  }

  range.start.paragraph.content = mergedText.length > 0 ? [{ kind: 'text', text: mergedText, format: inferRunFormatAt(range.start.paragraph.content, deleteStartOffset) }] : [];
  return changes;
}

export function nativeReplaceBlockRangeAcrossParagraphs(
  doc: NativeWordDocument,
  anchor: { paraId: string; offset: number },
  focus: { paraId: string; offset: number },
  text: string,
): NativeBlockRangeReplacement | null {
  const anchorLocation = findNativeParagraphLocation(doc, anchor.paraId);
  const focusLocation = findNativeParagraphLocation(doc, focus.paraId);
  if (!anchorLocation || !focusLocation || anchorLocation.container !== doc.blocks || focusLocation.container !== doc.blocks) {
    return null;
  }
  const ordered =
    anchorLocation.index < focusLocation.index ||
    (anchorLocation.index === focusLocation.index && anchor.offset <= focus.offset)
      ? { start: anchorLocation, end: focusLocation, startOffset: anchor.offset, endOffset: focus.offset }
      : { start: focusLocation, end: anchorLocation, startOffset: focus.offset, endOffset: anchor.offset };
  if (ordered.start.index === ordered.end.index) return null;
  if (
    ordered.start.paragraph.content.some((child) => child.kind !== 'text') ||
    ordered.end.paragraph.content.some((child) => child.kind !== 'text')
  ) {
    return null;
  }
  const selectedBlocks = doc.blocks.slice(ordered.start.index, ordered.end.index + 1);
  if (!selectedBlocks.slice(1, -1).some((block) => block.kind === 'table' || block.content.some((child) => child.kind !== 'text'))) {
    return null;
  }

  const beforeBlocks = cloneNativeBlocks(selectedBlocks);
  const startTextLength = nativeParagraphText(ordered.start.paragraph).length;
  const endTextLength = nativeParagraphText(ordered.end.paragraph).length;
  const startOffset = Math.max(0, Math.min(ordered.startOffset, startTextLength));
  const endOffset = Math.max(0, Math.min(ordered.endOffset, endTextLength));
  const changes = [
    ...nativeDeleteRange(doc, ordered.end.paragraph.paraId, 0, endOffset),
    ...nativeDeleteRange(doc, ordered.start.paragraph.paraId, startOffset, startTextLength),
  ];
  ordered.start.paragraph.content = normalizeTextNodes([
    ...ordered.start.paragraph.content,
    ...ordered.end.paragraph.content,
  ]);
  doc.blocks.splice(ordered.start.index + 1, ordered.end.index - ordered.start.index);
  if (text.length > 0) {
    const insertion = nativeInsertText(doc, ordered.start.paragraph.paraId, startOffset, text);
    if (insertion) changes.push(insertion);
  }
  changes.push({
    kind: 'structure',
    op: 'delete-range',
    startParaId: ordered.start.paragraph.paraId,
    endParaId: ordered.end.paragraph.paraId,
    deletedBlockIds: selectedBlocks.slice(1).map(nativeBlockId),
  });
  return {
    changes,
    cursor: { paraId: ordered.start.paragraph.paraId, offset: startOffset + text.length },
    container: doc.blocks,
    startIndex: ordered.start.index,
    beforeBlocks,
    afterBlocks: cloneNativeBlocks([ordered.start.paragraph]),
  };
}

export function nativeCopyBlockRange(
  doc: NativeWordDocument,
  anchor: { paraId: string; offset: number },
  focus: { paraId: string; offset: number },
): NativeWordBlockNode[] | null {
  const ordered = orderedTopLevelBlockRange(doc, anchor, focus);
  if (!ordered || (ordered.blocks.length < 2 && !rangeContainsComplexBlock(ordered.blocks))) return null;
  const blocks = cloneNativeBlocks(ordered.blocks);
  const start = blocks[0];
  const end = blocks[blocks.length - 1];
  if (start?.kind !== 'paragraph' || end?.kind !== 'paragraph') return null;
  start.content = splitInlineContentAt(start.content, ordered.startOffset).after;
  end.content = splitInlineContentAt(end.content, ordered.endOffset).before;
  return blocks;
}

export function nativeReplaceRangeWithBlocks(
  doc: NativeWordDocument,
  anchor: { paraId: string; offset: number },
  focus: { paraId: string; offset: number },
  clipboardBlocks: readonly NativeWordBlockNode[],
): NativeBlockRangeReplacement | null {
  if (anchor.paraId !== focus.paraId || clipboardBlocks.length === 0) return null;
  const location = findNativeParagraphLocation(doc, anchor.paraId);
  if (!location || location.container !== doc.blocks) return null;
  const firstClipboard = clipboardBlocks[0];
  const lastClipboard = clipboardBlocks[clipboardBlocks.length - 1];
  if (firstClipboard?.kind !== 'paragraph' || lastClipboard?.kind !== 'paragraph') return null;

  const beforeBlocks = cloneNativeBlocks([location.paragraph]);
  const targetLength = nativeParagraphText(location.paragraph).length;
  const from = Math.max(0, Math.min(anchor.offset, focus.offset, targetLength));
  const to = Math.max(from, Math.min(Math.max(anchor.offset, focus.offset), targetLength));
  const targetStart = splitInlineContentAt(location.paragraph.content, from).before;
  const targetEnd = splitInlineContentAt(location.paragraph.content, to).after;
  const inserted = cloneClipboardBlocksWithFreshIds(doc, clipboardBlocks);
  const first = inserted[0];
  const last = inserted[inserted.length - 1];
  if (first?.kind !== 'paragraph' || last?.kind !== 'paragraph') return null;

  const replacesEmptyParagraph = targetStart.length === 0 && targetEnd.length === 0;
  if (replacesEmptyParagraph) {
    location.paragraph.styleId = first.styleId;
    location.paragraph.format = first.format ? cloneJson(first.format) : null;
  }
  location.paragraph.content = normalizeTextNodes([...targetStart, ...first.content]);
  let cursor = {
    paraId: location.paragraph.paraId,
    offset: inlineContentLength(location.paragraph.content),
  };
  const replacement: NativeWordBlockNode[] = [location.paragraph];
  if (inserted.length === 1) {
    location.paragraph.content = normalizeTextNodes([...location.paragraph.content, ...targetEnd]);
  } else {
    last.content = normalizeTextNodes([...last.content, ...targetEnd]);
    replacement.push(...inserted.slice(1));
    cursor = { paraId: last.paraId, offset: inlineContentLength(last.content) - inlineContentLength(targetEnd) };
  }
  doc.blocks.splice(location.index, 1, ...replacement);
  const afterBlocks = cloneNativeBlocks(replacement);
  const change: NativeWordChange = {
    kind: 'structure',
    op: 'insert-range',
    paraId: location.paragraph.paraId,
    insertedBlockIds: replacement.slice(1).map(nativeBlockId),
  };
  return {
    changes: [change],
    cursor,
    container: doc.blocks,
    startIndex: location.index,
    beforeBlocks,
    afterBlocks,
  };
}

export function nativeReplaceRange(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
  text: string,
): NativeWordChange[] {
  const range = normalizedRange(from, to);
  const changes = nativeDeleteRange(doc, paraId, range.from, range.to);
  const insert = nativeInsertText(doc, paraId, range.from, text);
  return insert ? [...changes, insert] : changes;
}

export function nativeReplaceTextRangeAcrossParagraphs(
  doc: NativeWordDocument,
  anchor: { paraId: string; offset: number },
  focus: { paraId: string; offset: number },
  text: string,
): NativeWordChange[] {
  if (anchor.paraId === focus.paraId) {
    return nativeReplaceRange(doc, anchor.paraId, anchor.offset, focus.offset, text);
  }
  const anchorLocation = findNativeParagraphLocation(doc, anchor.paraId);
  const focusLocation = findNativeParagraphLocation(doc, focus.paraId);
  if (!anchorLocation || !focusLocation || anchorLocation.container !== focusLocation.container) return [];
  const start = anchorLocation.index < focusLocation.index ? anchor : focus;
  const changes = nativeDeleteTextRangeAcrossParagraphs(doc, anchor, focus);
  if (changes.length === 0 || text.length === 0) return changes;
  const insert = nativeInsertText(doc, start.paraId, start.offset, text);
  return insert ? [...changes, insert] : changes;
}

export function nativeReplaceTextRangeWithParagraphs(
  doc: NativeWordDocument,
  anchor: { paraId: string; offset: number },
  focus: { paraId: string; offset: number },
  text: string,
): NativePlainTextReplacement | null {
  const anchorLocation = findNativeParagraphLocation(doc, anchor.paraId);
  const focusLocation = findNativeParagraphLocation(doc, focus.paraId);
  if (!anchorLocation || !focusLocation || anchorLocation.container !== focusLocation.container) return null;
  const start =
    anchorLocation.index < focusLocation.index ||
    (anchorLocation.index === focusLocation.index && anchor.offset <= focus.offset)
      ? anchor
      : focus;
  const changes = nativeDeleteTextRangeAcrossParagraphs(doc, anchor, focus);
  const paragraph = findNativeParagraph(doc, start.paraId);
  if (!paragraph || paragraph.content.some((child) => child.kind !== 'text')) return null;

  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const paragraphIds = [start.paraId];
  let currentParaId = start.paraId;
  let insertOffset = Math.max(0, Math.min(start.offset, nativeParagraphText(paragraph).length));
  const firstInsert = nativeInsertText(doc, currentParaId, insertOffset, lines[0] ?? '');
  if (firstInsert) changes.push(firstInsert);

  for (let index = 1; index < lines.length; index += 1) {
    const nextParaId = createNativeParaId(doc);
    const split = nativeSplitParagraph(doc, currentParaId, insertOffset + (lines[index - 1]?.length ?? 0), nextParaId);
    if (!split) return null;
    changes.push(split);
    paragraphIds.push(nextParaId);
    currentParaId = nextParaId;
    insertOffset = 0;
    const insert = nativeInsertText(doc, currentParaId, 0, lines[index] ?? '');
    if (insert) changes.push(insert);
  }

  return {
    changes,
    cursor: {
      paraId: currentParaId,
      offset: lines.length === 1
        ? insertOffset + (lines[0]?.length ?? 0)
        : (lines[lines.length - 1]?.length ?? 0),
    },
    paragraphIds,
  };
}

export function nativeSetRunFormat(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
  format: Partial<EditorRunFormat>,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return null;
  let position = 0;
  let changed = false;
  const next: NativeWordInlineNode[] = [];
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind !== 'text' || position + length <= range.from || position >= range.to) {
      next.push(child);
      position += length;
      continue;
    }
    const localFrom = Math.max(0, range.from - position);
    const localTo = Math.min(length, range.to - position);
    if (localFrom > 0) next.push({ ...child, text: child.text.slice(0, localFrom) });
    if (localTo > localFrom) {
      next.push({
        ...child,
        text: child.text.slice(localFrom, localTo),
        format: { ...(child.format ?? {}), ...format },
      });
      changed = true;
    }
    if (localTo < length) next.push({ ...child, text: child.text.slice(localTo) });
    position += length;
  }
  if (!changed) return null;
  paragraph.content = normalizeTextNodes(next);
  return { kind: 'format', op: 'run', paraId, from: range.from, to: range.to, format };
}

export function nativeAddComment(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
  comment: { commentId: number; author: string; date: string; body: string; parentCommentId?: number; preserveExistingComments?: boolean },
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return null;
  let position = 0;
  let changed = false;
  const next: NativeWordInlineNode[] = [];
  const mark = {
    kind: 'comment' as const,
    id: `comment:${comment.commentId}`,
    commentId: comment.commentId,
    ...(comment.parentCommentId !== undefined ? { parentCommentId: comment.parentCommentId } : {}),
    author: comment.author,
    date: comment.date,
    body: comment.body,
  };
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind !== 'text' || position + length <= range.from || position >= range.to) {
      next.push(child);
      position += length;
      continue;
    }
    const localFrom = Math.max(0, range.from - position);
    const localTo = Math.min(length, range.to - position);
    if (localFrom > 0) next.push({ ...child, text: child.text.slice(0, localFrom) });
    if (localTo > localFrom) {
      const text = child.text.slice(localFrom, localTo);
      if (comment.preserveExistingComments && child.marks?.some((item) => item.kind === 'comment')) {
        next.push({ ...child, text });
        next.push({ ...child, text, marks: [mark, ...(child.marks ?? [])] });
      } else {
        const marks = (child.marks ?? []).filter((item) => item.kind !== 'comment');
        next.push({ ...child, text, marks: [...marks, mark] });
      }
      changed = true;
    }
    if (localTo < length) next.push({ ...child, text: child.text.slice(localTo) });
    position += length;
  }
  if (!changed) return null;
  paragraph.content = normalizeTextNodes(next);
  return {
    kind: 'annotation', op: 'insert', paraId, from: range.from, to: range.to,
    commentId: comment.commentId, body: comment.body, parentCommentId: comment.parentCommentId,
  };
}

export function nativeUpdateComment(
  doc: NativeWordDocument,
  commentId: number,
  body: string,
): NativeWordChange[] {
  const changes: NativeWordChange[] = [];
  for (const paragraph of allNativeParagraphs(doc)) {
    let position = 0;
    let from: number | null = null;
    let to: number | null = null;
    let changed = false;
    const content = paragraph.content.map((child) => {
      const length = inlineLength(child);
      if (child.kind !== 'text') {
        position += length;
        return child;
      }
      const comment = child.marks?.find((mark) => mark.kind === 'comment' && mark.commentId === commentId);
      if (!comment) {
        position += length;
        return child;
      }
      from = from === null ? position : from;
      to = position + length;
      changed = true;
      position += length;
      return {
        ...child,
        marks: child.marks?.map((mark) => mark.kind === 'comment' && mark.commentId === commentId
          ? { ...mark, body }
          : mark),
      };
    });
    if (!changed || from === null || to === null) continue;
    paragraph.content = normalizeTextNodes(content);
    changes.push({ kind: 'annotation', op: 'update', paraId: paragraph.paraId, from, to, commentId, body });
  }
  return changes;
}

export function nativeDeleteComment(doc: NativeWordDocument, commentId: number): NativeWordChange[] {
  const changes: NativeWordChange[] = [];
  for (const paragraph of allNativeParagraphs(doc)) {
    let position = 0;
    let from: number | null = null;
    let to: number | null = null;
    let changed = false;
    const content = paragraph.content.map((child) => {
      const length = inlineLength(child);
      if (child.kind !== 'text') {
        position += length;
        return child;
      }
      const hasComment = child.marks?.some((mark) => mark.kind === 'comment' && mark.commentId === commentId) ?? false;
      if (!hasComment) {
        position += length;
        return child;
      }
      from = from === null ? position : from;
      to = position + length;
      changed = true;
      position += length;
      const marks = child.marks?.filter((mark) => !(mark.kind === 'comment' && mark.commentId === commentId));
      return marks && marks.length > 0 ? { ...child, marks } : { ...child, marks: undefined };
    });
    if (!changed || from === null || to === null) continue;
    paragraph.content = normalizeTextNodes(content);
    changes.push({ kind: 'annotation', op: 'delete', paraId: paragraph.paraId, from, to, commentId });
  }
  return changes;
}

export function nativeAddRevision(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
  revision: { revisionId: string; revisionKind: 'insert' | 'delete'; author: string; date: string },
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return null;
  let position = 0;
  let changed = false;
  const next: NativeWordInlineNode[] = [];
  const mark = {
    kind: 'revision' as const,
    id: `revision:${revision.revisionId}`,
    revisionId: revision.revisionId,
    revisionKind: revision.revisionKind,
    author: revision.author,
    date: revision.date,
  };
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind !== 'text' || position + length <= range.from || position >= range.to) {
      next.push(child);
      position += length;
      continue;
    }
    const localFrom = Math.max(0, range.from - position);
    const localTo = Math.min(length, range.to - position);
    if (localFrom > 0) next.push({ ...child, text: child.text.slice(0, localFrom) });
    if (localTo > localFrom) {
      const existingRevision = (child.marks ?? []).find(
        (item) => item.kind === 'revision',
      );
      // A pending insertion that is deleted in the same editing session has
      // no net document change. Remove it instead of turning it into a
      // deletion revision; the surrounding transaction still preserves undo.
      if (!(revision.revisionKind === 'delete' && existingRevision?.revisionKind === 'insert')) {
        const marks = (child.marks ?? []).filter((item) => item.kind !== 'revision');
        next.push({ ...child, text: child.text.slice(localFrom, localTo), marks: [...marks, mark] });
      }
      changed = true;
    }
    if (localTo < length) next.push({ ...child, text: child.text.slice(localTo) });
    position += length;
  }
  if (!changed) return null;
  paragraph.content = normalizeTextNodes(next);
  return {
    kind: 'revision',
    op: revision.revisionKind,
    paraId,
    from: range.from,
    to: range.to,
    revisionId: revision.revisionId,
    revisionKind: revision.revisionKind,
  };
}

export function nativeResolveRevision(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
  operation: 'accept' | 'reject',
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return null;
  let position = 0;
  let changed = false;
  let revisionId = '';
  let revisionKind: 'insert' | 'delete' = 'insert';
  const next: NativeWordInlineNode[] = [];
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind !== 'text' || position + length <= range.from || position >= range.to) {
      next.push(child);
      position += length;
      continue;
    }
    const localFrom = Math.max(0, range.from - position);
    const localTo = Math.min(length, range.to - position);
    if (localFrom > 0) next.push({ ...child, text: child.text.slice(0, localFrom) });
    if (localTo > localFrom) {
      const selectedText = child.text.slice(localFrom, localTo);
      const revision = (child.marks ?? []).find((mark) => mark.kind === 'revision');
      if (!revision) {
        next.push({ ...child, text: selectedText });
      } else {
        changed = true;
        revisionId = revision.revisionId ?? revision.id.replace(/^revision:/, '');
        revisionKind = revision.revisionKind ?? 'insert';
        const keepText = operation === 'accept'
          ? revisionKind === 'insert'
          : revisionKind === 'delete';
        if (keepText) {
          next.push({
            ...child,
            text: selectedText,
            marks: (child.marks ?? []).filter((mark) => mark.kind !== 'revision'),
          });
        }
      }
    }
    if (localTo < length) next.push({ ...child, text: child.text.slice(localTo) });
    position += length;
  }
  if (!changed) return null;
  paragraph.content = normalizeTextNodes(next);
  return {
    kind: 'revision',
    op: operation,
    paraId,
    from: range.from,
    to: range.to,
    revisionId,
    revisionKind,
  };
}

export function nativeResolveFormatRevision(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
  operation: 'accept' | 'reject',
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return null;
  let position = 0;
  let changed = false;
  let revisionId = '';
  const next: NativeWordInlineNode[] = [];
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind !== 'text' || position + length <= range.from || position >= range.to) {
      next.push(child);
      position += length;
      continue;
    }
    const localFrom = Math.max(0, range.from - position);
    const localTo = Math.min(length, range.to - position);
    if (localFrom > 0) next.push({ ...child, text: child.text.slice(0, localFrom) });
    if (localTo > localFrom) {
      const mark = child.formatRevision;
      if (!mark) next.push({ ...child, text: child.text.slice(localFrom, localTo) });
      else {
        revisionId = mark.revisionId;
        changed = true;
        next.push({
          ...child,
          text: child.text.slice(localFrom, localTo),
          format: operation === 'reject' ? (mark.beforeFormat ? { ...mark.beforeFormat } : null) : child.format,
          formatRevision: undefined,
        });
      }
    }
    if (localTo < length) next.push({ ...child, text: child.text.slice(localTo) });
    position += length;
  }
  if (!changed) return null;
  paragraph.content = normalizeTextNodes(next);
  return { kind: 'format', op: operation, paraId, from: range.from, to: range.to, format: {}, revisionId, revisionKind: 'format' };
}

export function nativeSetHyperlink(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
  hyperlink: EditorHyperlink | null,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return null;
  const normalized = hyperlink ? normalizeHyperlink(hyperlink) : null;
  if (hyperlink && !normalized) return null;
  let position = 0;
  let changed = false;
  const next: NativeWordInlineNode[] = [];
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind !== 'text' || position + length <= range.from || position >= range.to) {
      next.push(child);
      position += length;
      continue;
    }
    const localFrom = Math.max(0, range.from - position);
    const localTo = Math.min(length, range.to - position);
    if (localFrom > 0) next.push({ ...child, text: child.text.slice(0, localFrom) });
    if (localTo > localFrom) {
      const selected = { ...child, text: child.text.slice(localFrom, localTo) };
      if (normalized) selected.hyperlink = { ...normalized };
      else delete selected.hyperlink;
      next.push(selected);
      if (!sameJson(child.hyperlink ?? null, normalized)) changed = true;
    }
    if (localTo < length) next.push({ ...child, text: child.text.slice(localTo) });
    position += length;
  }
  if (!changed) return null;
  paragraph.content = normalizeTextNodes(next);
  return {
    kind: 'hyperlink',
    op: normalized ? 'set' : 'remove',
    paraId,
    from: range.from,
    to: range.to,
    ...(normalized ? { hyperlink: normalized } : {}),
  };
}

export function nativeSetParagraphFormat(
  doc: NativeWordDocument,
  paraId: string,
  format: Partial<EditorParagraphFormat>,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph) return null;
  const current = paragraph.format ?? {};
  paragraph.format = { ...current, ...format };
  if (format.indent) paragraph.format.indent = { ...(current.indent ?? {}), ...format.indent };
  if (format.spacing) paragraph.format.spacing = { ...(current.spacing ?? {}), ...format.spacing };
  return { kind: 'format', op: 'paragraph', paraId, format };
}

export function nativeClearRunFormat(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return null;
  let position = 0;
  let changed = false;
  const next: NativeWordInlineNode[] = [];
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind !== 'text' || position + length <= range.from || position >= range.to) {
      next.push(child);
      position += length;
      continue;
    }
    const localFrom = Math.max(0, range.from - position);
    const localTo = Math.min(length, range.to - position);
    if (localFrom > 0) next.push({ ...child, text: child.text.slice(0, localFrom) });
    if (localTo > localFrom) {
      next.push({ ...child, text: child.text.slice(localFrom, localTo), format: null });
      if (child.format && Object.keys(child.format).length > 0) changed = true;
    }
    if (localTo < length) next.push({ ...child, text: child.text.slice(localTo) });
    position += length;
  }
  if (!changed) return null;
  paragraph.content = normalizeTextNodes(next);
  return { kind: 'format', op: 'run', paraId, from: range.from, to: range.to, format: {}, clear: true };
}

export function nativeSetParagraphList(
  doc: NativeWordDocument,
  paraId: string,
  numbering: EditorNumberingFormat | null,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph) return null;
  const current = paragraph.format ?? {};
  if (numbering) {
    paragraph.format = {
      ...current,
      hasNumbering: true,
      numbering: { ...numbering },
      indent: { ...(current.indent ?? {}), ...nativeListIndent(numbering.ilvl) },
    };
  } else {
    const next: EditorParagraphFormat = {
      ...current,
      hasNumbering: false,
      indent: { ...(current.indent ?? {}), left: 0, firstLine: 0 },
    };
    delete next.numbering;
    paragraph.format = next;
  }
  return { kind: 'list', paraId, numbering: numbering ? { ...numbering } : null };
}

export function nativeTextRangeIsBold(doc: NativeWordDocument, paraId: string, from: number, to: number): boolean {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return false;
  let position = 0;
  let hasText = false;
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    const overlaps = child.kind === 'text' && position + length > range.from && position < range.to;
    if (overlaps) {
      hasText = true;
      if (child.format?.bold !== true) return false;
    }
    position += length;
  }
  return hasText;
}

export function nativeDeleteImage(
  doc: NativeWordDocument,
  paraId: string,
  drawingId: string,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph) return null;
  let offset = 0;
  const next: NativeWordInlineNode[] = [];
  let change: NativeWordChange | null = null;
  for (const child of paragraph.content) {
    if (!change && child.kind === 'image' && child.drawingId === drawingId) {
      change = { kind: 'image', op: 'delete', paraId, offset, drawingId };
    } else {
      next.push(child);
    }
    offset += inlineLength(child);
  }
  if (!change) return null;
  paragraph.content = normalizeTextNodes(next);
  return change;
}

export function nativeAddFormatRevision(
  doc: NativeWordDocument,
  paraId: string,
  from: number,
  to: number,
  beforeParagraph: NativeWordParagraphNode,
  revision: { revisionId: string; author?: string; date?: string },
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  const range = normalizedRange(from, to);
  if (!paragraph || range.from === range.to) return null;
  let position = 0;
  let changed = false;
  const next: NativeWordInlineNode[] = [];
  const formatAt = (offset: number): EditorRunFormat | null => {
    let beforePosition = 0;
    for (const child of beforeParagraph.content) {
      const length = inlineLength(child);
      if (child.kind === 'text' && offset >= beforePosition && offset < beforePosition + length) {
        return child.format ? { ...child.format } : null;
      }
      beforePosition += length;
    }
    return null;
  };
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind !== 'text' || position + length <= range.from || position >= range.to) {
      next.push(child);
      position += length;
      continue;
    }
    const localFrom = Math.max(0, range.from - position);
    const localTo = Math.min(length, range.to - position);
    if (localFrom > 0) next.push({ ...child, text: child.text.slice(0, localFrom) });
    if (localTo > localFrom) {
      next.push({
        ...child,
        text: child.text.slice(localFrom, localTo),
        formatRevision: child.formatRevision ?? {
            kind: 'format',
            revisionId: revision.revisionId,
            beforeFormat: formatAt(position + localFrom),
            author: revision.author,
            date: revision.date,
          },
      });
      changed = true;
    }
    if (localTo < length) next.push({ ...child, text: child.text.slice(localTo) });
    position += length;
  }
  if (!changed) return null;
  paragraph.content = normalizeTextNodes(next);
  return {
    kind: 'format',
    op: 'run',
    paraId,
    from: range.from,
    to: range.to,
    format: {},
    revisionId: revision.revisionId,
    revisionKind: 'format',
    beforeFormat: formatAt(range.from),
  };
}

export function nativeResolveImageRevision(
  doc: NativeWordDocument,
  paraId: string,
  drawingId: string,
  operation: 'accept' | 'reject',
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph) return null;
  let offset = 0;
  for (let index = 0; index < paragraph.content.length; index += 1) {
    const child = paragraph.content[index];
    if (child?.kind !== 'image' || child.drawingId !== drawingId) {
      offset += child ? inlineLength(child) : 0;
      continue;
    }
    const revision = child.revision;
    if (!revision) return null;
    if ((revision.revisionKind === 'delete' && operation === 'accept') ||
      (revision.revisionKind === 'insert' && operation === 'reject')) {
      paragraph.content.splice(index, 1);
    } else {
      child.revision = undefined;
    }
    paragraph.content = normalizeTextNodes(paragraph.content);
    return {
      kind: 'image',
      op: operation,
      paraId,
      offset,
      drawingId,
      revisionId: revision.revisionId,
      revisionKind: revision.revisionKind,
    };
  }
  return null;
}

export function nativeResolveImageFormatRevision(
  doc: NativeWordDocument,
  paraId: string,
  drawingId: string,
  operation: 'accept' | 'reject',
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph) return null;
  let offset = 0;
  for (const child of paragraph.content) {
    if (child.kind !== 'image' || child.drawingId !== drawingId) {
      offset += inlineLength(child);
      continue;
    }
    const revision = child.formatRevision;
    if (!revision) return null;
    if (operation === 'reject') {
      child.width = revision.beforeWidth;
      child.height = revision.beforeHeight;
    }
    const revisionId = revision.revisionId;
    child.formatRevision = undefined;
    return {
      kind: 'image',
      op: operation,
      paraId,
      offset,
      drawingId,
      revisionId,
      revisionKind: 'format',
    };
  }
  return null;
}

export function nativeSplitParagraph(
  doc: NativeWordDocument,
  paraId: string,
  offset: number,
  newParaId = createNativeParaId(doc),
): NativeWordChange | null {
  const location = findNativeParagraphLocation(doc, paraId);
  if (!location) return null;
  const { paragraph } = location;

  const split = splitInlineContentAt(paragraph.content, offset);
  paragraph.content = split.before;
  location.insertAfter({
    kind: 'paragraph',
    paraId: newParaId,
    story: paragraph.story,
    styleId: paragraph.styleId,
    format: paragraph.format ? JSON.parse(JSON.stringify(paragraph.format)) : null,
    content: split.after,
  });
  return { kind: 'block', op: 'split', paraId, newParaId, offset };
}

export function nativeSyncParagraphText(
  doc: NativeWordDocument,
  paraId: string,
  nextText: string,
): NativeWordChange | null {
  const paragraph = findNativeParagraph(doc, paraId);
  if (!paragraph || paragraph.content.some((child) => child.kind !== 'text')) return null;

  const previousText = nativeParagraphText(paragraph);
  if (previousText === nextText) return null;

  const diff = textDiff(previousText, nextText);
  let nextContent = paragraph.content;
  if (diff.deletedText.length > 0) {
    nextContent = deleteInlineTextRange(nextContent, diff.offset, diff.deletedText.length);
  }
  if (diff.insertedText.length > 0) {
    nextContent = insertInlineText(nextContent, diff.offset, diff.insertedText, inferRunFormatAt(paragraph.content, diff.offset));
  }
  paragraph.content = nextContent;
  if (diff.deletedText.length === 0) {
    return { kind: 'text', op: 'insert', paraId, offset: diff.offset, text: diff.insertedText };
  }
  if (diff.insertedText.length === 0) {
    return { kind: 'text', op: 'delete', paraId, offset: diff.offset, text: diff.deletedText };
  }
  return {
    kind: 'text',
    op: 'replace',
    paraId,
    offset: diff.offset,
    text: diff.insertedText,
    deletedText: diff.deletedText,
  };
}

function insertInlineText(
  content: NativeWordInlineNode[],
  offset: number,
  text: string,
  fallbackFormat: EditorRunFormat | null = null,
): NativeWordInlineNode[] {
  const next: NativeWordInlineNode[] = [];
  let position = 0;
  let inserted = false;
  for (const child of content) {
    const length = inlineLength(child);
    if (!inserted && offset <= position + length) {
      if (child.kind === 'text') {
        const localOffset = Math.max(0, offset - position);
        next.push({
          ...child,
          text: `${child.text.slice(0, localOffset)}${text}${child.text.slice(localOffset)}`,
        });
      } else if (offset <= position) {
        next.push({ kind: 'text', text, format: fallbackFormat }, child);
      } else {
        next.push(child, { kind: 'text', text, format: fallbackFormat });
      }
      inserted = true;
    } else {
      next.push(child);
    }
    position += length;
  }
  if (!inserted) next.push({ kind: 'text', text, format: fallbackFormat });
  return normalizeTextNodes(next);
}

function deleteInlineTextRange(content: NativeWordInlineNode[], offset: number, length: number): NativeWordInlineNode[] {
  if (length <= 0) return content;
  const next: NativeWordInlineNode[] = [];
  let position = 0;
  const end = offset + length;
  for (const child of content) {
    const childLength = inlineLength(child);
    const childEnd = position + childLength;
    if (child.kind !== 'text' || childEnd <= offset || position >= end) {
      next.push(child);
    } else {
      const keepBefore = Math.max(0, offset - position);
      const keepAfter = Math.max(0, childEnd - end);
      const text = `${child.text.slice(0, keepBefore)}${child.text.slice(child.text.length - keepAfter)}`;
      if (text.length > 0) next.push({ ...child, text });
    }
    position = childEnd;
  }
  return normalizeTextNodes(next);
}

function deleteInlineBackward(paragraph: NativeWordParagraphNode, offset: number): NativeWordChange | null {
  return deleteInlineAt(paragraph, offset - 1);
}

function createNativeDrawingId(doc: NativeWordDocument): string {
  const used = new Set(
    allNativeParagraphs(doc).flatMap((paragraph) =>
      paragraph.content
        .filter((child) => child.kind === 'image' || child.kind === 'drawing')
        .map((child) => child.drawingId),
    ),
  );
  let suffix = 1;
  while (used.has(`coco-clipboard-${suffix}`)) suffix += 1;
  return `coco-clipboard-${suffix}`;
}

function deleteInlineAt(paragraph: NativeWordParagraphNode, deleteOffset: number): NativeWordChange | null {
  if (deleteOffset < 0) return null;
  const next: NativeWordInlineNode[] = [];
  let position = 0;
  let change: NativeWordChange | null = null;
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (!change && deleteOffset >= position && deleteOffset < position + length) {
      if (child.kind === 'image') {
        change = {
          kind: 'image',
          op: 'delete',
          paraId: paragraph.paraId,
          offset: deleteOffset,
          drawingId: child.drawingId,
        };
      } else if (child.kind === 'pageBreak') {
        change = { kind: 'page-break', op: 'delete', paraId: paragraph.paraId, offset: deleteOffset };
      } else if (child.kind === 'text') {
        const localOffset = deleteOffset - position;
        const deleted = child.text.slice(localOffset, localOffset + 1);
        const text = `${child.text.slice(0, localOffset)}${child.text.slice(localOffset + 1)}`;
        if (text.length > 0) next.push({ ...child, text });
        change = { kind: 'text', op: 'delete', paraId: paragraph.paraId, offset: deleteOffset, text: deleted };
      } else {
        next.push(child);
      }
    } else {
      next.push(child);
    }
    position += length;
  }
  if (!change) return null;
  paragraph.content = normalizeTextNodes(next);
  return change;
}

function deleteInlineRange(
  paragraph: NativeWordParagraphNode,
  from: number,
  to: number,
): {
  content: NativeWordInlineNode[];
  changes: NativeWordChange[];
} {
  const next: NativeWordInlineNode[] = [];
  const changes: NativeWordChange[] = [];
  let position = 0;

  for (const child of paragraph.content) {
    const length = inlineLength(child);
    const end = position + length;
    const overlaps = from < end && to > position;
    if (!overlaps) {
      next.push(child);
      position = end;
      continue;
    }

    if (child.kind === 'text') {
      const localFrom = Math.max(0, from - position);
      const localTo = Math.min(child.text.length, to - position);
      const deleted = child.text.slice(localFrom, localTo);
      const kept = `${child.text.slice(0, localFrom)}${child.text.slice(localTo)}`;
      if (kept.length > 0) {
        next.push({ ...child, text: kept });
      }
      if (deleted.length > 0) {
        changes.push({
          kind: 'text',
          op: 'delete',
          paraId: paragraph.paraId,
          offset: position + localFrom,
          text: deleted,
        });
      }
    } else if (child.kind === 'image') {
      changes.push({
        kind: 'image',
        op: 'delete',
        paraId: paragraph.paraId,
        offset: position,
        drawingId: child.drawingId,
      });
    } else if (child.kind === 'pageBreak') {
      changes.push({
        kind: 'page-break',
        op: 'delete',
        paraId: paragraph.paraId,
        offset: position,
      });
    } else {
      next.push(child);
    }

    position = end;
  }

  return {
    content: normalizeTextNodes(next),
    changes: changes.sort((left, right) => changeOffset(right) - changeOffset(left)),
  };
}

function splitInlineContentAt(content: NativeWordInlineNode[], offset: number): {
  before: NativeWordInlineNode[];
  after: NativeWordInlineNode[];
} {
  const before: NativeWordInlineNode[] = [];
  const after: NativeWordInlineNode[] = [];
  let position = 0;
  let split = false;
  for (const child of content) {
    const length = inlineLength(child);
    if (!split && offset <= position + length) {
      if (child.kind === 'text') {
        const localOffset = Math.max(0, offset - position);
        if (localOffset > 0) before.push({ ...child, text: child.text.slice(0, localOffset) });
        if (localOffset < child.text.length) after.push({ ...child, text: child.text.slice(localOffset) });
      } else if (offset <= position) {
        after.push(child);
      } else {
        before.push(child);
      }
      split = true;
    } else if (split) {
      after.push(child);
    } else {
      before.push(child);
    }
    position += length;
  }
  return {
    before: normalizeTextNodes(before),
    after: normalizeTextNodes(after),
  };
}

function inlineLength(child: NativeWordInlineNode): number {
  return child.kind === 'text' ? child.text.length : 1;
}

function positiveImageDimension(value: number | undefined): number | undefined {
  if (!Number.isFinite(value) || value === undefined || value <= 0) return undefined;
  return Math.max(1, Math.round(value * 100) / 100);
}

function normalizedRange(from: number, to: number): { from: number; to: number } {
  const start = Math.max(0, Math.min(from, to));
  const end = Math.max(0, Math.max(from, to));
  return { from: start, to: end };
}

function changeOffset(change: NativeWordChange): number {
  return 'offset' in change && typeof change.offset === 'number' ? change.offset : 0;
}

function normalizeTextNodes(content: NativeWordInlineNode[]): NativeWordInlineNode[] {
  const next: NativeWordInlineNode[] = [];
  for (const child of content) {
    if (child.kind === 'text') {
      if (child.text.length === 0) continue;
      const previous = next[next.length - 1];
      if (
        previous?.kind === 'text' &&
        sameRunFormat(previous.format, child.format) &&
        sameJson(previous.hyperlink ?? null, child.hyperlink ?? null) &&
        sameJson(previous.marks ?? null, child.marks ?? null)
      ) {
        previous.text += child.text;
      } else {
        next.push({ ...child });
      }
    } else {
      next.push(child);
    }
  }
  return next;
}

function cloneNativeBlocks(blocks: readonly NativeWordBlockNode[]): NativeWordBlockNode[] {
  const cloned = JSON.parse(JSON.stringify(blocks)) as NativeWordBlockNode[];
  synchronizeTableParagraphAliases(cloned);
  return cloned;
}

function nativeBlockId(block: NativeWordBlockNode): string {
  return block.kind === 'paragraph' ? block.paraId : block.tableId;
}

function orderedTopLevelBlockRange(
  doc: NativeWordDocument,
  anchor: { paraId: string; offset: number },
  focus: { paraId: string; offset: number },
): {
  blocks: NativeWordBlockNode[];
  startOffset: number;
  endOffset: number;
} | null {
  const anchorLocation = findNativeParagraphLocation(doc, anchor.paraId);
  const focusLocation = findNativeParagraphLocation(doc, focus.paraId);
  if (!anchorLocation || !focusLocation || anchorLocation.container !== doc.blocks || focusLocation.container !== doc.blocks) {
    return null;
  }
  const ordered =
    anchorLocation.index < focusLocation.index ||
    (anchorLocation.index === focusLocation.index && anchor.offset <= focus.offset)
      ? { start: anchorLocation, end: focusLocation, startOffset: anchor.offset, endOffset: focus.offset }
      : { start: focusLocation, end: anchorLocation, startOffset: focus.offset, endOffset: anchor.offset };
  if (ordered.start.index === ordered.end.index) return null;
  if (
    ordered.start.paragraph.content.some((child) => child.kind !== 'text') ||
    ordered.end.paragraph.content.some((child) => child.kind !== 'text')
  ) {
    return null;
  }
  return {
    blocks: doc.blocks.slice(ordered.start.index, ordered.end.index + 1),
    startOffset: Math.max(0, Math.min(ordered.startOffset, nativeParagraphText(ordered.start.paragraph).length)),
    endOffset: Math.max(0, Math.min(ordered.endOffset, nativeParagraphText(ordered.end.paragraph).length)),
  };
}

function rangeContainsComplexBlock(blocks: readonly NativeWordBlockNode[]): boolean {
  return blocks.slice(1, -1).some(
    (block) => block.kind === 'table' || block.content.some((child) => child.kind !== 'text'),
  );
}

function cloneClipboardBlocksWithFreshIds(
  doc: NativeWordDocument,
  blocks: readonly NativeWordBlockNode[],
): NativeWordBlockNode[] {
  const cloned = cloneNativeBlocks(blocks);
  const usedParaIds = new Set(allNativeParagraphs(doc).map((paragraph) => paragraph.paraId));
  const usedDrawingIds = new Set<string>();
  const usedTableIds = new Set<string>();
  visitNativeBlocks(doc.blocks, (block) => {
    if (block.kind === 'table') usedTableIds.add(block.tableId);
    else {
      for (const child of block.content) {
        if (child.kind === 'image' || child.kind === 'drawing') usedDrawingIds.add(child.drawingId);
      }
    }
  });
  let nextParaId = Date.now() >>> 0;
  const freshParaId = () => {
    let paraId = '';
    do {
      nextParaId = (nextParaId + 1) >>> 0;
      paraId = nextParaId.toString(16).toUpperCase().padStart(8, '0');
    } while (usedParaIds.has(paraId));
    usedParaIds.add(paraId);
    return paraId;
  };
  const freshNamedId = (used: Set<string>, prefix: string) => {
    let suffix = 1;
    while (used.has(`${prefix}-${suffix}`)) suffix += 1;
    const id = `${prefix}-${suffix}`;
    used.add(id);
    return id;
  };
  visitNativeBlocks(cloned, (block) => {
    if (block.kind === 'table') {
      block.tableId = freshNamedId(usedTableIds, 'coco-clipboard-table');
      return;
    }
    block.paraId = freshParaId();
    for (const child of block.content) {
      if (child.kind === 'image' || child.kind === 'drawing') {
        child.drawingId = freshNamedId(usedDrawingIds, 'coco-clipboard-drawing');
      }
      if (child.kind === 'drawing' && child.paragraphs) {
        for (const paragraph of child.paragraphs) paragraph.paraId = freshParaId();
      }
    }
  });
  synchronizeTableParagraphAliases(cloned);
  return cloned;
}

function freshNativeTableId(doc: NativeWordDocument): string {
  const used = new Set<string>();
  for (const blocks of [doc.blocks, ...doc.headers.map((story) => story.blocks), ...doc.footers.map((story) => story.blocks)]) {
    visitNativeBlocks(blocks, (block) => {
      if (block.kind === 'table') used.add(block.tableId);
    });
  }
  let suffix = 1;
  while (used.has(`coco-table-${suffix}`)) suffix += 1;
  return `coco-table-${suffix}`;
}

function visitNativeBlocks(
  blocks: readonly NativeWordBlockNode[],
  visitor: (block: NativeWordBlockNode) => void,
): void {
  for (const block of blocks) {
    visitor(block);
    if (block.kind !== 'table') continue;
    for (const row of block.rows) {
      for (const cell of row.cells) visitNativeBlocks(cell.blocks, visitor);
    }
  }
}

function synchronizeTableParagraphAliases(blocks: NativeWordBlockNode[]): void {
  for (const block of blocks) {
    if (block.kind !== 'table') continue;
    for (const row of block.rows) {
      for (const cell of row.cells) {
        synchronizeTableParagraphAliases(cell.blocks);
        cell.paragraphs = collectParagraphs(cell.blocks);
      }
    }
  }
}

function collectParagraphs(blocks: readonly NativeWordBlockNode[]): NativeWordParagraphNode[] {
  return blocks.flatMap((block) =>
    block.kind === 'paragraph'
      ? [block]
      : block.rows.flatMap((row) => row.cells.flatMap((cell) => collectParagraphs(cell.blocks))),
  );
}

function inlineContentLength(content: readonly NativeWordInlineNode[]): number {
  return content.reduce((length, child) => length + inlineLength(child), 0);
}

function sameRunFormat(a: EditorRunFormat | null, b: EditorRunFormat | null): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

function normalizeHyperlink(hyperlink: EditorHyperlink): EditorHyperlink | null {
  const target = hyperlink.target?.trim();
  const anchor = hyperlink.anchor?.trim().replace(/^#/, '');
  const tooltip = hyperlink.tooltip?.trim();
  if (!target && !anchor) return null;
  return {
    ...(target ? { target } : {}),
    ...(anchor ? { anchor } : {}),
    ...(tooltip ? { tooltip } : {}),
  };
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function inferRunFormatAt(content: NativeWordInlineNode[], offset: number): EditorRunFormat | null {
  let position = 0;
  let previousTextFormat: EditorRunFormat | null = null;
  for (const child of content) {
    const length = inlineLength(child);
    if (child.kind === 'text') {
      if (offset >= position && offset <= position + length) return child.format;
      previousTextFormat = child.format;
    }
    position += length;
  }
  return previousTextFormat;
}

function textDiff(previousText: string, nextText: string): {
  offset: number;
  deletedText: string;
  insertedText: string;
} {
  let start = 0;
  while (
    start < previousText.length &&
    start < nextText.length &&
    previousText.charCodeAt(start) === nextText.charCodeAt(start)
  ) {
    start += 1;
  }

  let previousEnd = previousText.length;
  let nextEnd = nextText.length;
  while (
    previousEnd > start &&
    nextEnd > start &&
    previousText.charCodeAt(previousEnd - 1) === nextText.charCodeAt(nextEnd - 1)
  ) {
    previousEnd -= 1;
    nextEnd -= 1;
  }

  return {
    offset: start,
    deletedText: previousText.slice(start, previousEnd),
    insertedText: nextText.slice(start, nextEnd),
  };
}
