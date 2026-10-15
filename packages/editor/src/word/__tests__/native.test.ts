import { describe, expect, test } from 'vitest';
import { summarizeEditorCoverage, summarizeNativeCoverage } from '../coverage.js';
import { resolveParagraphStyle, resolveRunStyle, resolveTableCellStyle } from '../style-resolver.js';
import {
  adjacentNativeListNumbering,
  createNativeWordDocument,
  createNativeListNumbering,
  findNativeParagraph,
  nativeClearRunFormat,
  nativeCopyBlockRange,
  nativeDeleteBackward,
  nativeDeleteForward,
  nativeDeleteImage,
  nativeDeleteTable,
  nativeDeleteRange,
  nativeDeleteTextRangeAcrossParagraphs,
  nativeInsertImage,
  nativeInsertPageBreak,
  nativeInsertTable,
  nativeInsertText,
  nativeEditTableStructure,
  nativeTableSupportsStructuralEditing,
  nativeReplaceRange,
  nativeReplaceBlockRangeAcrossParagraphs,
  nativeReplaceRangeWithBlocks,
  nativeReplaceTextRangeAcrossParagraphs,
  nativeReplaceTextRangeWithParagraphs,
  nativeSetParagraphFormat,
  nativeSetImageSize,
  nativeSetHyperlink,
  nativeSetTableCellFormat,
  nativeSetTableCellRangeFormat,
  nativeClearTableCellRange,
  nativeReplaceTableCellRangeText,
  nativeReplaceTableCellRangeContent,
  nativeTableCellsInRange,
  cloneNativeWordBlocks,
  nativeSetParagraphList,
  nativeSetRunFormat,
  nativeTextRangeIsBold,
  resolveNativeTableCellBorders,
  nativeDrawingWidthCss,
  nativeSyncParagraphText,
  nativePageBreakAt,
  nativeRangeContainsPageBreak,
  nativeSplitParagraph,
  nativeAnchorLayoutMetadata,
  analyzeNativeTableRowPagination,
  resolveNativeSectionPageNumberStarts,
  refreshNativeNumberingLabels,
  splitNativeTableRowAtBlockBoundary,
  splitNativeTableRowAtTextOffset,
  summarizeNativeTablePagination,
  type NativeWordDocument,
} from '../native.js';
import type { EditorDocumentView } from '../model.js';

function createView(): EditorDocumentView {
  return {
    kind: 'word',
    sections: [],
    headers: [],
    footers: [],
    blocks: [
      {
        type: 'paragraph',
        paraId: 'P0000001',
        story: 'body',
        text: 'AC',
        format: {
          alignment: 'center',
          spacing: { before: 240, after: 120, line: 288, lineRule: 'auto' },
          fontSize: 16,
          bold: true,
          color: 'C00000',
        },
        flags: {
          hasBookmarks: false,
          hasComments: false,
          hasMath: false,
          hasRevisions: false,
        },
        segments: [{ kind: 'text', from: 0, to: 2, text: 'AC' }],
        spans: [],
      },
      {
        type: 'paragraph',
        paraId: 'P0000002',
        story: 'body',
        text: 'B',
        flags: {
          hasBookmarks: false,
          hasComments: false,
          hasMath: false,
          hasRevisions: false,
        },
        segments: [
          { kind: 'text', from: 0, to: 1, text: 'B', format: { bold: true, color: '00AA00', fontSize: 14 } },
        ],
        spans: [],
      },
      {
        type: 'paragraph',
        paraId: 'P0000003',
        story: 'body',
        text: 'XY',
        flags: {
          hasBookmarks: false,
          hasComments: false,
          hasMath: false,
          hasRevisions: false,
        },
        segments: [
          { kind: 'text', from: 0, to: 1, text: 'X' },
          {
            kind: 'image',
            from: 1,
            to: 2,
            drawingId: 'drawing-1',
            src: 'data:image/png;base64,abc',
            contentType: 'image/png',
            width: 80,
            height: 40,
          },
          { kind: 'text', from: 2, to: 3, text: 'Y' },
        ],
        spans: [],
      },
      {
        type: 'table',
        tableId: 'tbl-0001',
        columnWidths: [1200],
        borders: { insideH: { val: 'single', color: 'DEE0E3', size: 4 } },
        rows: [
          {
            cells: [
              {
                text: 'MN',
                gridSpan: 1,
                width: 1440,
                margin: { top: 60, right: 120, bottom: 30, left: 120 },
                shading: { fill: 'D9EAF7' },
                verticalAlignment: 'center',
                vMerge: null,
                paragraphs: [
                  {
                    type: 'paragraph',
                    paraId: 'P0000004',
                    story: 'body',
                    text: 'MN',
                    flags: {
                      hasBookmarks: false,
                      hasComments: false,
                      hasMath: false,
                      hasRevisions: false,
                    },
                    segments: [
                      { kind: 'text', from: 0, to: 1, text: 'M' },
                      {
                        kind: 'image',
                        from: 1,
                        to: 2,
                        drawingId: 'drawing-2',
                        src: 'data:image/png;base64,def',
                        contentType: 'image/png',
                        width: 20,
                        height: 10,
                      },
                      { kind: 'text', from: 2, to: 3, text: 'N' },
                    ],
                    spans: [],
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
    paragraphs: [],
    tables: [],
    detachedComments: [],
    detachedRevisions: [],
  };
}

function paragraphTexts(doc: NativeWordDocument): string[] {
  return doc.blocks.flatMap((block) =>
    block.kind === 'paragraph'
      ? [block.content.map((child) => (child.kind === 'text' ? child.text : '[image]')).join('')]
      : [],
  );
}

function allParagraphTexts(doc: NativeWordDocument): string[] {
  return doc.blocks.flatMap((block) => {
    if (block.kind === 'paragraph') {
      return [block.content.map((child) => (child.kind === 'text' ? child.text : '[image]')).join('')];
    }
    return block.rows.flatMap((row) =>
      row.cells.flatMap((cell) =>
        cell.paragraphs.map((paragraph) =>
          paragraph.content.map((child) => (child.kind === 'text' ? child.text : '[image]')).join(''),
        ),
      ),
    );
  });
}

function nativeTextParagraph(paraId: string, text: string): NativeWordDocument['blocks'][number] {
  return {
    kind: 'paragraph',
    paraId,
    story: 'body',
    styleId: null,
    format: null,
    content: [{ kind: 'text', text, format: null }],
  };
}

function nativeTableDocument(): NativeWordDocument {
  const paragraph = (paraId: string, text: string) => nativeTextParagraph(paraId, text) as Extract<
    NativeWordDocument['blocks'][number],
    { kind: 'paragraph' }
  >;
  const cell = (paraId: string, text: string) => {
    const block = paragraph(paraId, text);
    return { gridSpan: 1, blocks: [block], paragraphs: [block] };
  };
  return {
    kind: 'word',
    sections: [],
    headers: [],
    footers: [],
    blocks: [
      {
        kind: 'table',
        tableId: 'table-edit',
        columnWidths: [1200, 1800],
        rows: [
          { cells: [cell('T0000001', 'A1'), cell('T0000002', 'B1')] },
          { cells: [cell('T0000003', 'A2'), cell('T0000004', 'B2')] },
        ],
      },
    ],
  };
}

describe('@coco-vexel/coco-office-editor native word POC', () => {
  test('sets, edits, and removes hyperlink semantics without merging adjacent plain text', () => {
    const doc = createNativeWordDocument(createView());
    const set = nativeSetHyperlink(doc, 'P0000001', 0, 1, {
      target: 'https://coco.example/docs',
      tooltip: 'Docs',
    });
    const paragraph = findNativeParagraph(doc, 'P0000001')!;

    expect(set).toMatchObject({
      kind: 'hyperlink',
      op: 'set',
      from: 0,
      to: 1,
      hyperlink: { target: 'https://coco.example/docs', tooltip: 'Docs' },
    });
    expect(paragraph.content).toMatchObject([
      { kind: 'text', text: 'A', hyperlink: { target: 'https://coco.example/docs' } },
      { kind: 'text', text: 'C' },
    ]);
    expect(paragraph.content[1]).not.toHaveProperty('hyperlink');

    nativeInsertText(doc, 'P0000001', 1, '!');
    expect(paragraph.content).toMatchObject([
      { kind: 'text', text: 'A!', hyperlink: { target: 'https://coco.example/docs' } },
      { kind: 'text', text: 'C' },
    ]);

    const removed = nativeSetHyperlink(doc, 'P0000001', 0, 2, null);
    expect(removed).toMatchObject({ kind: 'hyperlink', op: 'remove', from: 0, to: 2 });
    expect(paragraph.content).toMatchObject([{ kind: 'text', text: 'A!C' }]);
    expect(paragraph.content[0]).not.toHaveProperty('hyperlink');
  });

  test('inserts a table by splitting a body paragraph and deletes it without losing surrounding text', () => {
    const doc = createNativeWordDocument(createView());
    let nextId = 10;
    const createParaId = () => `P${String(nextId++).padStart(7, '0')}`;
    const inserted = nativeInsertTable(doc, 'P0000001', 1, 2, 3, createParaId, () => 'table-new');

    expect(inserted).toMatchObject({
      changes: [{ kind: 'table', op: 'insert-table', tableId: 'table-new' }],
      cursor: { paraId: 'P0000011', offset: 0 },
      startIndex: 0,
    });
    expect(doc.blocks.slice(0, 3).map((block) => block.kind)).toEqual(['paragraph', 'table', 'paragraph']);
    expect(doc.blocks[0]?.kind === 'paragraph' ? doc.blocks[0].content[0] : null).toMatchObject({ text: 'A' });
    expect(doc.blocks[2]?.kind === 'paragraph' ? doc.blocks[2].content[0] : null).toMatchObject({ text: 'C' });
    const table = doc.blocks[1];
    expect(table?.kind === 'table' ? [table.rows.length, table.rows[0]?.cells.length] : null).toEqual([2, 3]);

    const deleted = nativeDeleteTable(doc, 'table-new', createParaId);
    expect(deleted?.changes).toEqual([
      { kind: 'table', op: 'delete-table', tableId: 'table-new', rowIndex: 0, columnIndex: 0 },
    ]);
    expect(doc.blocks.slice(0, 3).map((block) => block.kind)).toEqual(['paragraph', 'paragraph', 'paragraph']);
    expect(doc.blocks[0]?.kind === 'paragraph' ? doc.blocks[0].content[0] : null).toMatchObject({ text: 'A' });
    expect(doc.blocks[2]?.kind === 'paragraph' ? doc.blocks[2].content[0] : null).toMatchObject({ text: 'C' });
  });

  test('inserts and deletes table rows without changing existing cell content', () => {
    const doc = nativeTableDocument();
    let nextId = 10;
    const createParaId = () => `T${String(nextId++).padStart(7, '0')}`;

    const inserted = nativeEditTableStructure(
      doc,
      'T0000001',
      { kind: 'insert-row', position: 'after' },
      createParaId,
    );
    const table = doc.blocks[0]!;

    expect(inserted).toMatchObject({
      change: { kind: 'table', op: 'insert-row', tableId: 'table-edit', rowIndex: 1, columnIndex: 0 },
      cursor: { paraId: 'T0000010', offset: 0 },
    });
    expect(
      table.kind === 'table'
        ? table.rows.map((row) =>
            row.cells.map((cell) =>
              cell.paragraphs[0]?.content.map((child) => child.kind === 'text' ? child.text : '').join(''),
            ),
          )
        : [],
    ).toEqual([
      ['A1', 'B1'],
      ['', ''],
      ['A2', 'B2'],
    ]);

    const deleted = nativeEditTableStructure(doc, 'T0000010', { kind: 'delete-row' }, createParaId);
    expect(deleted?.change).toMatchObject({ kind: 'table', op: 'delete-row', rowIndex: 1 });
    expect(table.kind === 'table' ? table.rows.length : 0).toBe(2);
    expect(findNativeParagraph(doc, 'T0000003')).not.toBeNull();
  });

  test('inserts and deletes table columns while preserving widths and content', () => {
    const doc = nativeTableDocument();
    let nextId = 20;
    const createParaId = () => `T${String(nextId++).padStart(7, '0')}`;

    const inserted = nativeEditTableStructure(
      doc,
      'T0000002',
      { kind: 'insert-column', position: 'before' },
      createParaId,
    );
    const table = doc.blocks[0]!;

    expect(inserted).toMatchObject({
      change: { kind: 'table', op: 'insert-column', columnIndex: 1 },
      cursor: { paraId: 'T0000020', offset: 0 },
    });
    expect(table.kind === 'table' ? table.columnWidths : []).toEqual([1200, 1800, 1800]);
    expect(findNativeParagraph(doc, 'T0000001')).not.toBeNull();
    expect(findNativeParagraph(doc, 'T0000004')).not.toBeNull();

    const deleted = nativeEditTableStructure(doc, 'T0000020', { kind: 'delete-column' }, createParaId);
    expect(deleted?.change).toMatchObject({ kind: 'table', op: 'delete-column', columnIndex: 1 });
    expect(table.kind === 'table' ? table.columnWidths : []).toEqual([1200, 1800]);
  });

  test('rejects structural editing for merged tables', () => {
    const doc = nativeTableDocument();
    const table = doc.blocks[0];
    if (table?.kind !== 'table') throw new Error('Expected table fixture');
    table.rows[0]!.cells[0]!.gridSpan = 2;

    expect(nativeTableSupportsStructuralEditing(table)).toBe(false);
    expect(nativeEditTableStructure(doc, 'T0000001', { kind: 'delete-column' })).toBeNull();
  });

  test('formats table cell background, alignment, and borders without changing content', () => {
    const doc = nativeTableDocument();

    const result = nativeSetTableCellFormat(doc, 'T0000001', {
      backgroundColor: '#AABBCC',
      verticalAlignment: 'bottom',
      borderColor: '#112233',
    });
    const table = doc.blocks[0];
    if (table?.kind !== 'table') throw new Error('Expected table fixture');
    const cell = table.rows[0]!.cells[0]!;

    expect(result?.change).toEqual({
      kind: 'table',
      op: 'format-cell',
      tableId: 'table-edit',
      rowIndex: 0,
      columnIndex: 0,
      format: {
        backgroundColor: '#AABBCC',
        verticalAlignment: 'bottom',
        borderColor: '#112233',
      },
    });
    expect(cell.shading).toEqual({ fill: 'AABBCC', val: 'clear' });
    expect(cell.verticalAlignment).toBe('bottom');
    expect(cell.borders?.top).toEqual({ val: 'single', color: '112233', size: 4 });
    expect(findNativeParagraph(doc, 'T0000001')?.content[0]).toMatchObject({ text: 'A1' });

    nativeSetTableCellFormat(doc, 'T0000001', {
      backgroundColor: null,
      verticalAlignment: null,
      borderColor: null,
    });
    expect(cell.shading).toBeUndefined();
    expect(cell.verticalAlignment).toBeUndefined();
    expect(cell.borders?.top).toEqual({ val: 'nil' });
  });

  test('formats, clears, and replaces a reverse table cell range as rectangular operations', () => {
    const doc = nativeTableDocument();
    const table = doc.blocks[0];
    if (table?.kind !== 'table') throw new Error('Expected table fixture');
    const range = {
      anchor: { rowIndex: 1, columnIndex: 1 },
      focus: { rowIndex: 0, columnIndex: 0 },
    };

    expect(nativeTableCellsInRange(table, range)).toHaveLength(4);
    const formatted = nativeSetTableCellRangeFormat(doc, table.tableId, range, {
      backgroundColor: '#DDEEFF',
      verticalAlignment: 'center',
    });
    expect(formatted?.changes).toHaveLength(4);
    expect(table.rows.flatMap((row) => row.cells).every((cell) => cell.shading?.fill === 'DDEEFF')).toBe(true);

    const cleared = nativeClearTableCellRange(doc, table.tableId, range);
    expect(cleared?.changes).toHaveLength(4);
    expect(allParagraphTexts(doc)).toEqual(['', '', '', '']);

    const replaced = nativeReplaceTableCellRangeText(
      doc,
      table.tableId,
      { rowIndex: 0, columnIndex: 0 },
      [['A', 'B'], ['C', 'D']],
    );
    expect(replaced?.changes).toHaveLength(4);
    expect(allParagraphTexts(doc)).toEqual(['A', 'B', 'C', 'D']);
  });

  test('pastes rich table cell content and presentation with fresh object identities', () => {
    const doc = nativeTableDocument();
    const table = doc.blocks[0];
    if (table?.kind !== 'table') throw new Error('Expected table fixture');
    const source = table.rows[0]!.cells[0]!;
    const sourceParagraph = source.paragraphs[0]!;
    sourceParagraph.content = [
      { kind: 'text', text: 'Styled', format: { bold: true, italic: true, color: 'AA0000' } },
      { kind: 'image', drawingId: 'source-image', src: 'data:image/png;base64,abc', width: 20, height: 10 },
    ];
    source.shading = { fill: 'DDEEFF', val: 'clear' };
    source.verticalAlignment = 'bottom';

    const result = nativeReplaceTableCellRangeContent(
      doc,
      table.tableId,
      { rowIndex: 1, columnIndex: 1 },
      [[{
        blocks: cloneNativeWordBlocks(source.blocks),
        format: { shading: source.shading, verticalAlignment: source.verticalAlignment },
      }]],
    );
    const target = table.rows[1]!.cells[1]!;
    const targetParagraph = target.paragraphs[0]!;
    const targetImage = targetParagraph.content.find((child) => child.kind === 'image');

    expect(result?.changes).toHaveLength(1);
    expect(targetParagraph.paraId).not.toBe(sourceParagraph.paraId);
    expect(targetParagraph.content[0]).toMatchObject({
      kind: 'text',
      text: 'Styled',
      format: { bold: true, italic: true, color: 'AA0000' },
    });
    expect(targetImage).toMatchObject({ kind: 'image', src: 'data:image/png;base64,abc' });
    expect(targetImage?.kind === 'image' ? targetImage.drawingId : null).not.toBe('source-image');
    expect(target.shading).toEqual({ fill: 'DDEEFF', val: 'clear' });
    expect(target.verticalAlignment).toBe('bottom');
  });

  test('keeps paragraph and image identity when projecting from EditorDocumentView', () => {
    const doc = createNativeWordDocument(createView());

    expect(doc.blocks[0]).toMatchObject({ kind: 'paragraph', paraId: 'P0000001' });
    expect(doc.blocks[0]).toMatchObject({
      kind: 'paragraph',
      content: [{ kind: 'text', text: 'AC', format: null }],
    });
    expect(doc.blocks[0]).toMatchObject({
      kind: 'paragraph',
      format: {
        alignment: 'center',
        fontSize: 16,
        color: 'C00000',
      },
    });
    expect(doc.blocks[2]).toMatchObject({
      kind: 'paragraph',
      paraId: 'P0000003',
      content: [
        { kind: 'text', text: 'X' },
        { kind: 'image', drawingId: 'drawing-1', width: 80, height: 40 },
        { kind: 'text', text: 'Y' },
      ],
    });
    expect(doc.blocks[1]).toMatchObject({
      kind: 'paragraph',
      content: [{ kind: 'text', text: 'B', format: { bold: true, color: '00AA00', fontSize: 14 } }],
    });
  });

  test('preserves paragraph pagination signals in the native model', () => {
    const view = createView();
    const paragraph = view.blocks[0];
    if (paragraph?.type !== 'paragraph') throw new Error('paragraph fixture missing');
    paragraph.format = {
      ...paragraph.format,
      pageBreakBefore: true,
      hasSectionBreak: true,
      keepNext: true,
      keepLines: true,
      widowControl: false,
    };

    const doc = createNativeWordDocument(view);

    expect(doc.blocks[0]).toMatchObject({
      kind: 'paragraph',
      format: {
        pageBreakBefore: true,
        hasSectionBreak: true,
        keepNext: true,
        keepLines: true,
        widowControl: false,
      },
    });
  });

  test('continues page numbering across sections unless Word explicitly restarts it', () => {
    const starts = resolveNativeSectionPageNumberStarts(
      [
        { sectionIndex: 0 },
        { sectionIndex: 1, pageNumberStart: 1 },
        { sectionIndex: 2 },
        { sectionIndex: 3 },
      ],
      [0, 1, 2, 3],
      new Map([
        [0, 5],
        [1, 35],
        [2, 2],
        [3, 86],
      ]),
    );

    expect(starts).toEqual(
      new Map([
        [0, 1],
        [1, 1],
        [2, 36],
        [3, 38],
      ]),
    );
  });

  test('resolves anchored image wrap and position metadata for native layout', () => {
    const metadata = nativeAnchorLayoutMetadata({
      wrap: 'square',
      position: {
        horizontal: { relativeFrom: 'margin', align: 'right' },
        vertical: { relativeFrom: 'paragraph', posOffset: 914400 },
      },
    });

    expect(metadata.classNames).toContain('is-float-right');
    expect(metadata.dataset).toMatchObject({
      positionHorizontalRelativeFrom: 'margin',
      positionHorizontalAlign: 'right',
      positionVerticalOffset: '914400',
    });
    expect(metadata.styleProperties['--coco-word-anchor-top']).toBe('96px');
  });

  test('resolves behind text anchors as overlay-like native layout metadata', () => {
    const metadata = nativeAnchorLayoutMetadata({
      wrap: 'behind',
      position: {
        horizontal: { relativeFrom: 'margin', posOffset: -228600 },
        vertical: { relativeFrom: 'paragraph', posOffset: -1493 },
      },
    });

    expect(metadata.classNames).toContain('is-overlay-like');
    expect(metadata.classNames).toContain('is-absolute-anchor');
    expect(metadata.classNames).toContain('is-paragraph-relative-anchor');
    expect(metadata.classNames).toContain('is-behind-text');
    expect(metadata.dataset.positionHorizontalOffset).toBe('-228600');
    expect(metadata.styleProperties['--coco-word-anchor-left']).toBe('-24px');
  });

  test('uses horizontal alignment for in-front absolute anchors', () => {
    const metadata = nativeAnchorLayoutMetadata({
      wrap: 'inFront',
      position: {
        horizontal: { relativeFrom: 'margin', align: 'right' },
        vertical: { relativeFrom: 'paragraph', posOffset: 0 },
      },
    });

    expect(metadata.classNames).toContain('is-absolute-anchor');
    expect(metadata.classNames).toContain('is-anchor-right');
    expect(metadata.classNames).toContain('is-in-front-of-text');
    expect(metadata.dataset.positionHorizontalAlign).toBe('right');
  });

  test('preserves page-relative anchor coordinates for page layer promotion', () => {
    const metadata = nativeAnchorLayoutMetadata({
      wrap: 'behind',
      position: {
        horizontal: { relativeFrom: 'page', posOffset: 914400 },
        vertical: { relativeFrom: 'page', posOffset: 457200 },
      },
    });

    expect(metadata.classNames).toContain('is-absolute-anchor');
    expect(metadata.classNames).toContain('is-page-relative-anchor');
    expect(metadata.dataset).toMatchObject({
      positionHorizontalRelativeFrom: 'page',
      positionVerticalRelativeFrom: 'page',
      positionHorizontalOffset: '914400',
      positionVerticalOffset: '457200',
    });
    expect(metadata.styleProperties['--coco-word-anchor-left']).toBe('96px');
    expect(metadata.styleProperties['--coco-word-anchor-top']).toBe('48px');
  });

  test('keeps mixed margin and paragraph anchors paragraph-relative', () => {
    const metadata = nativeAnchorLayoutMetadata({
      wrap: 'behind',
      position: {
        horizontal: { relativeFrom: 'margin', posOffset: -228600 },
        vertical: { relativeFrom: 'paragraph', posOffset: -1493 },
      },
    });

    expect(metadata.classNames).toContain('is-paragraph-relative-anchor');
    expect(metadata.classNames).not.toContain('is-page-relative-anchor');
  });

  test('keeps top and bottom anchors out of text-wrap floats', () => {
    const metadata = nativeAnchorLayoutMetadata({
      wrap: 'topAndBottom',
      position: {
        horizontal: { relativeFrom: 'margin', posOffset: -127672 },
        vertical: { relativeFrom: 'paragraph', posOffset: 917351 },
      },
    });

    expect(metadata.classNames).toContain('is-top-bottom');
    expect(metadata.classNames).not.toContain('is-float-left');
    expect(metadata.classNames).not.toContain('is-float-right');
  });

  test('constrains flow drawing widths without changing absolute anchors', () => {
    expect(nativeDrawingWidthCss({ width: 624, anchoringMode: 'inline' })).toBe(
      'min(624px, max(1px, calc(100% - 4pt)))',
    );
    expect(nativeDrawingWidthCss({ width: 624, anchoringMode: 'inline', isLineOrConnector: true })).toBe(
      'min(624px, 100%)',
    );
    expect(nativeDrawingWidthCss({ width: 624, anchoringMode: 'anchor', wrap: 'inFront' })).toBe('624px');
  });

  test('preserves comment spans as native text marks', () => {
    const view = createView();
    const paragraph = view.blocks[0];
    if (paragraph?.type !== 'paragraph') throw new Error('paragraph fixture missing');
    paragraph.text = 'ABC';
    paragraph.flags.hasComments = true;
    paragraph.segments = [{ kind: 'text', from: 0, to: 3, text: 'ABC' }];
    paragraph.spans = [
      {
        kind: 'comment',
        id: 'comment:7',
        commentId: 7,
        paraId: paragraph.paraId,
        from: 1,
        to: 2,
        text: 'B',
        rangeKind: 'text',
        author: 'reviewer',
        date: '',
        body: 'check',
      },
    ];

    const doc = createNativeWordDocument(view);

    expect(doc.blocks[0]).toMatchObject({
      kind: 'paragraph',
      content: [
        { kind: 'text', text: 'A' },
        { kind: 'text', text: 'B', marks: [{ kind: 'comment', commentId: 7 }] },
        { kind: 'text', text: 'C' },
      ],
    });
  });

  test('preserves revision segments as native text marks', () => {
    const view = createView();
    const paragraph = view.blocks[0];
    if (paragraph?.type !== 'paragraph') throw new Error('paragraph fixture missing');
    paragraph.text = 'Aoldnew';
    paragraph.segments = [
      { kind: 'text', from: 0, to: 1, text: 'A' },
      {
        kind: 'revision',
        from: 1,
        to: 4,
        text: 'old',
        revisionId: '20',
        revisionKind: 'delete',
        author: 'bob',
        date: '2026-04-22T11:00:00Z',
      },
      {
        kind: 'revision',
        from: 4,
        to: 7,
        text: 'new',
        revisionId: '10',
        revisionKind: 'insert',
        author: 'alice',
        date: '2026-04-22T10:00:00Z',
      },
    ];
    paragraph.spans = [
      {
        kind: 'revision',
        id: 'revision:20',
        revisionId: '20',
        revisionKind: 'delete',
        paraId: paragraph.paraId,
        from: 1,
        to: 4,
        text: 'old',
        author: 'bob',
        date: '2026-04-22T11:00:00Z',
      },
      {
        kind: 'revision',
        id: 'revision:10',
        revisionId: '10',
        revisionKind: 'insert',
        paraId: paragraph.paraId,
        from: 4,
        to: 7,
        text: 'new',
        author: 'alice',
        date: '2026-04-22T10:00:00Z',
      },
    ];

    const doc = createNativeWordDocument(view);

    expect(doc.blocks[0]).toMatchObject({
      kind: 'paragraph',
      content: [
        { kind: 'text', text: 'A' },
        {
          kind: 'text',
          text: 'old',
          marks: [{ kind: 'revision', revisionId: '20', revisionKind: 'delete' }],
        },
        {
          kind: 'text',
          text: 'new',
          marks: [{ kind: 'revision', revisionId: '10', revisionKind: 'insert' }],
        },
      ],
    });
  });

  test('projects text box drawings into the native model', () => {
    const view = createView();
    const paragraph = view.blocks[0];
    if (paragraph?.type !== 'paragraph') throw new Error('paragraph fixture missing');
    paragraph.text = 'Seal';
    paragraph.segments = [
      {
        kind: 'drawing',
        from: 0,
        to: 4,
        drawingId: 'dr-textbox-1',
        drawingKind: 'textBox',
        text: 'Seal',
        width: 160,
        height: 48,
        anchoringMode: 'anchor',
        wrap: 'square',
        position: { horizontal: { align: 'right' } },
        hasBorder: true,
      },
    ];

    const doc = createNativeWordDocument(view);

    expect(doc.blocks[0]).toMatchObject({
      kind: 'paragraph',
      content: [
        {
          kind: 'drawing',
          drawingId: 'dr-textbox-1',
          drawingKind: 'textBox',
          text: 'Seal',
          width: 160,
          anchoringMode: 'anchor',
          wrap: 'square',
        },
      ],
    });
  });

  test('inserts text into a paragraph without a ProseMirror document', () => {
    const doc = createNativeWordDocument(createView());
    const change = nativeInsertText(doc, 'P0000001', 1, 'B');

    expect(change).toEqual({ kind: 'text', op: 'insert', paraId: 'P0000001', offset: 1, text: 'B' });
    expect(paragraphTexts(doc)[0]).toBe('ABC');
  });

  test('keeps run format when inserting into formatted text', () => {
    const doc = createNativeWordDocument(createView());
    const change = nativeInsertText(doc, 'P0000002', 1, 'C');

    expect(change).toEqual({ kind: 'text', op: 'insert', paraId: 'P0000002', offset: 1, text: 'C' });
    expect(doc.blocks[1]).toMatchObject({
      kind: 'paragraph',
      content: [{ kind: 'text', text: 'BC', format: { bold: true, color: '00AA00', fontSize: 14 } }],
    });
  });

  test('splits and merges paragraphs in the native model', () => {
    const doc = createNativeWordDocument(createView());

    const split = nativeSplitParagraph(doc, 'P0000001', 1, 'P00000FF');
    expect(split).toEqual({
      kind: 'block',
      op: 'split',
      paraId: 'P0000001',
      newParaId: 'P00000FF',
      offset: 1,
    });
    expect(doc.blocks[1]).toMatchObject({
      kind: 'paragraph',
      format: {
        alignment: 'center',
        fontSize: 16,
        color: 'C00000',
      },
    });
    expect(paragraphTexts(doc).slice(0, 3)).toEqual(['A', 'C', 'B']);

    const merge = nativeDeleteBackward(doc, 'P00000FF', 0);
    expect(merge).toEqual({ kind: 'block', op: 'merge', paraId: 'P0000001' });
    expect(paragraphTexts(doc).slice(0, 2)).toEqual(['AC', 'B']);
  });

  test('syncs browser-mutated paragraph text back into the native model', () => {
    const doc = createNativeWordDocument(createView());

    expect(nativeSyncParagraphText(doc, 'P0000001', 'ABC')).toEqual({
      kind: 'text',
      op: 'insert',
      paraId: 'P0000001',
      offset: 1,
      text: 'B',
    });
    expect(nativeSyncParagraphText(doc, 'P0000001', 'ADC')).toEqual({
      kind: 'text',
      op: 'replace',
      paraId: 'P0000001',
      offset: 1,
      text: 'D',
      deletedText: 'B',
    });
    expect(nativeSyncParagraphText(doc, 'P0000001', 'AC')).toEqual({
      kind: 'text',
      op: 'delete',
      paraId: 'P0000001',
      offset: 1,
      text: 'D',
    });
    expect(paragraphTexts(doc)[0]).toBe('AC');
  });

  test('preserves run format when syncing browser-mutated formatted text', () => {
    const doc = createNativeWordDocument(createView());

    expect(nativeSyncParagraphText(doc, 'P0000002', 'BC')).toEqual({
      kind: 'text',
      op: 'insert',
      paraId: 'P0000002',
      offset: 1,
      text: 'C',
    });
    expect(doc.blocks[1]).toMatchObject({
      kind: 'paragraph',
      content: [{ kind: 'text', text: 'BC', format: { bold: true, color: '00AA00', fontSize: 14 } }],
    });

    expect(nativeSyncParagraphText(doc, 'P0000002', 'DC')).toEqual({
      kind: 'text',
      op: 'replace',
      paraId: 'P0000002',
      offset: 0,
      text: 'D',
      deletedText: 'B',
    });
    expect(doc.blocks[1]).toMatchObject({
      kind: 'paragraph',
      content: [{ kind: 'text', text: 'DC', format: { bold: true, color: '00AA00', fontSize: 14 } }],
    });
  });

  test('does not collapse image paragraphs during text fallback sync', () => {
    const doc = createNativeWordDocument(createView());
    const change = nativeSyncParagraphText(doc, 'P0000003', 'XZ');

    expect(change).toBeNull();
    expect(paragraphTexts(doc)[2]).toBe('X[image]Y');
  });

  test('preserves and deletes image nodes as native inline content', () => {
    const doc = createNativeWordDocument(createView());

    const split = nativeSplitParagraph(doc, 'P0000003', 2, 'P00000AA');
    expect(split).toEqual({
      kind: 'block',
      op: 'split',
      paraId: 'P0000003',
      newParaId: 'P00000AA',
      offset: 2,
    });
    expect(paragraphTexts(doc).slice(2, 4)).toEqual(['X[image]', 'Y']);

    const deleteImage = nativeDeleteBackward(doc, 'P0000003', 2);
    expect(deleteImage).toEqual({
      kind: 'image',
      op: 'delete',
      paraId: 'P0000003',
      offset: 1,
      drawingId: 'drawing-1',
    });
    expect(paragraphTexts(doc).slice(2, 4)).toEqual(['X', 'Y']);
  });

  test('inserts and deletes page breaks as one-unit inline nodes', () => {
    const doc = createNativeWordDocument(createView());

    expect(nativeInsertPageBreak(doc, 'P0000001', 1)).toEqual({
      kind: 'page-break',
      op: 'insert',
      paraId: 'P0000001',
      offset: 1,
    });
    expect(nativePageBreakAt(doc, 'P0000001', 1)).toBe(true);
    expect(nativeDeleteBackward(doc, 'P0000001', 2)).toEqual({
      kind: 'page-break',
      op: 'delete',
      paraId: 'P0000001',
      offset: 1,
    });
    expect(nativePageBreakAt(doc, 'P0000001', 1)).toBe(false);

    nativeInsertPageBreak(doc, 'P0000001', 1);
    expect(nativeDeleteForward(doc, 'P0000001', 1)).toEqual({
      kind: 'page-break',
      op: 'delete',
      paraId: 'P0000001',
      offset: 1,
    });

    nativeInsertPageBreak(doc, 'P0000001', 1);
    expect(nativeRangeContainsPageBreak(doc, 'P0000001', 0, 2)).toBe(true);
    expect(nativeDeleteRange(doc, 'P0000001', 0, 2)).toEqual([
      { kind: 'page-break', op: 'delete', paraId: 'P0000001', offset: 1 },
      { kind: 'text', op: 'delete', paraId: 'P0000001', offset: 0, text: 'A' },
    ]);
    expect(nativePageBreakAt(doc, 'P0000001', 0)).toBe(false);
  });

  test('inserts an image with dimensions and records later resizing', () => {
    const doc = createNativeWordDocument(createView());

    const inserted = nativeInsertImage(doc, 'P0000001', 1, 'data:image/png;base64,image', {
      drawingId: 'drawing-inserted',
      width: 160,
      height: 90,
    });
    const resized = nativeSetImageSize(doc, 'P0000001', 'drawing-inserted', { width: 320, height: 180 });
    const paragraph = doc.blocks[0];

    expect(inserted).toMatchObject({
      kind: 'image',
      op: 'insert',
      drawingId: 'drawing-inserted',
      width: 160,
      height: 90,
    });
    expect(resized).toEqual({
      kind: 'image',
      op: 'resize',
      paraId: 'P0000001',
      offset: 1,
      drawingId: 'drawing-inserted',
      width: 320,
      height: 180,
    });
    expect(paragraph).toMatchObject({
      kind: 'paragraph',
      content: [
        { kind: 'text', text: 'A' },
        { kind: 'image', drawingId: 'drawing-inserted', width: 320, height: 180 },
        { kind: 'text', text: 'C' },
      ],
    });
  });

  test('deletes forward from a caret before inline media', () => {
    const doc = createNativeWordDocument(createView());

    const deleteImage = nativeDeleteForward(doc, 'P0000003', 1);

    expect(deleteImage).toEqual({
      kind: 'image',
      op: 'delete',
      paraId: 'P0000003',
      offset: 1,
      drawingId: 'drawing-1',
    });
    expect(paragraphTexts(doc)[2]).toBe('XY');
  });

  test('replaces a paragraph selection without dropping image delete intent', () => {
    const doc = createNativeWordDocument(createView());

    const changes = nativeReplaceRange(doc, 'P0000003', 0, 3, 'Z');

    expect(changes).toEqual([
      { kind: 'text', op: 'delete', paraId: 'P0000003', offset: 2, text: 'Y' },
      { kind: 'image', op: 'delete', paraId: 'P0000003', offset: 1, drawingId: 'drawing-1' },
      { kind: 'text', op: 'delete', paraId: 'P0000003', offset: 0, text: 'X' },
      { kind: 'text', op: 'insert', paraId: 'P0000003', offset: 0, text: 'Z' },
    ]);
    expect(paragraphTexts(doc)[2]).toBe('Z');
  });

  test('deletes a selected text range from the end first so export offsets remain stable', () => {
    const doc = createNativeWordDocument(createView());

    const changes = nativeDeleteRange(doc, 'P0000003', 0, 3);

    expect(changes.map((change) => ('offset' in change ? change.offset : undefined))).toEqual([2, 1, 0]);
    expect(paragraphTexts(doc)[2]).toBe('');
  });

  test('deletes a cross-paragraph text selection and merges the remaining edges', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks = [
      nativeTextParagraph('P1000001', 'Alpha'),
      nativeTextParagraph('P1000002', 'Middle'),
      nativeTextParagraph('P1000003', 'Omega'),
    ];

    const changes = nativeDeleteTextRangeAcrossParagraphs(
      doc,
      { paraId: 'P1000001', offset: 2 },
      { paraId: 'P1000003', offset: 2 },
    );

    expect(doc.blocks).toHaveLength(1);
    expect(paragraphTexts(doc)).toEqual(['Alega']);
    expect(changes).toEqual([
      { kind: 'text', op: 'delete', paraId: 'P1000003', offset: 0, text: 'Om' },
      { kind: 'text', op: 'delete', paraId: 'P1000002', offset: 0, text: 'Middle' },
      { kind: 'text', op: 'delete', paraId: 'P1000001', offset: 2, text: 'pha' },
      { kind: 'block', op: 'merge', paraId: 'P1000001' },
      { kind: 'block', op: 'merge', paraId: 'P1000001' },
    ]);
  });

  test('replaces a cross-paragraph text selection and keeps the ordered start paragraph', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks = [
      nativeTextParagraph('P1000001', 'Alpha'),
      nativeTextParagraph('P1000002', 'Middle'),
      nativeTextParagraph('P1000003', 'Omega'),
    ];

    const changes = nativeReplaceTextRangeAcrossParagraphs(
      doc,
      { paraId: 'P1000003', offset: 2 },
      { paraId: 'P1000001', offset: 2 },
      'XYZ',
    );

    expect(paragraphTexts(doc)).toEqual(['AlXYZega']);
    expect(changes[changes.length - 1]).toEqual({
      kind: 'text',
      op: 'insert',
      paraId: 'P1000001',
      offset: 2,
      text: 'XYZ',
    });
  });

  test('replaces complete table and image blocks between body paragraphs atomically', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks.push(nativeTextParagraph('P0000005', 'ZQ'));

    const result = nativeReplaceBlockRangeAcrossParagraphs(
      doc,
      { paraId: 'P0000002', offset: 1 },
      { paraId: 'P0000005', offset: 1 },
      '!',
    );

    expect(result).not.toBeNull();
    expect(result?.beforeBlocks).toHaveLength(4);
    expect(result?.afterBlocks).toHaveLength(1);
    expect(result?.cursor).toEqual({ paraId: 'P0000002', offset: 2 });
    expect(doc.blocks).toHaveLength(2);
    expect(paragraphTexts(doc)).toEqual(['AC', 'B!Q']);
    expect(result?.changes[result.changes.length - 1]).toEqual({
      kind: 'structure',
      op: 'delete-range',
      startParaId: 'P0000002',
      endParaId: 'P0000005',
      deletedBlockIds: ['P0000003', 'tbl-0001', 'P0000005'],
    });
  });

  test('copies and pastes table and image blocks with fresh object identities', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks.push(nativeTextParagraph('P0000005', 'ZQ'));
    const copied = nativeCopyBlockRange(
      doc,
      { paraId: 'P0000002', offset: 0 },
      { paraId: 'P0000005', offset: 1 },
    );

    expect(copied).not.toBeNull();
    expect(copied?.map((block) => block.kind)).toEqual(['paragraph', 'paragraph', 'table', 'paragraph']);
    const result = nativeReplaceRangeWithBlocks(
      doc,
      { paraId: 'P0000001', offset: 1 },
      { paraId: 'P0000001', offset: 1 },
      copied!,
    );

    expect(result?.beforeBlocks).toHaveLength(1);
    expect(result?.afterBlocks).toHaveLength(4);
    expect(doc.blocks).toHaveLength(8);
    expect(paragraphTexts(doc).slice(0, 3)).toEqual(['AB', 'X[image]Y', 'ZC']);
    const pastedTable = doc.blocks[2];
    expect(pastedTable?.kind).toBe('table');
    if (!pastedTable || pastedTable.kind !== 'table') throw new Error('pasted table missing');
    expect(pastedTable.tableId).not.toBe('tbl-0001');
    expect(pastedTable.rows[0]?.cells[0]?.blocks[0]).toBe(pastedTable.rows[0]?.cells[0]?.paragraphs[0]);
    const ids = doc.blocks.flatMap((block) => block.kind === 'paragraph' ? [block.paraId] : [block.tableId]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  test('copies ordinary paragraphs with formatting and applies the first paragraph style to an empty target', () => {
    const doc = createNativeWordDocument(createView());
    const first = doc.blocks[0];
    const second = doc.blocks[1];
    expect(first?.kind).toBe('paragraph');
    expect(second?.kind).toBe('paragraph');
    if (first?.kind !== 'paragraph' || second?.kind !== 'paragraph') throw new Error('paragraph fixture missing');
    first.styleId = 'DocumentTitle';
    first.format = { alignment: 'center', spacing: { after: 240 } };
    first.content = [{ kind: 'text', text: 'Red title', format: { bold: true, color: 'C00000', fontSize: 24 } }];
    second.styleId = 'BodyText';
    second.format = { alignment: 'justify', indent: { firstLine: 420 } };
    second.content = [{ kind: 'text', text: 'Body copy', format: { fontSize: 12 } }];
    const target = nativeTextParagraph('P0000005', '');
    doc.blocks.push(target);

    const copied = nativeCopyBlockRange(
      doc,
      { paraId: first.paraId, offset: 0 },
      { paraId: second.paraId, offset: 'Body copy'.length },
    );
    expect(copied).not.toBeNull();

    const result = nativeReplaceRangeWithBlocks(
      doc,
      { paraId: target.paraId, offset: 0 },
      { paraId: target.paraId, offset: 0 },
      copied!,
    );

    expect(result).not.toBeNull();
    const pastedFirst = doc.blocks.find((block) => block.kind === 'paragraph' && block.paraId === target.paraId);
    expect(pastedFirst).toMatchObject({
      kind: 'paragraph',
      styleId: 'DocumentTitle',
      format: { alignment: 'center', spacing: { after: 240 } },
      content: [{ text: 'Red title', format: { bold: true, color: 'C00000', fontSize: 24 } }],
    });
    const pastedSecond = doc.blocks[doc.blocks.indexOf(pastedFirst!) + 1];
    expect(pastedSecond).toMatchObject({
      kind: 'paragraph',
      styleId: 'BodyText',
      format: { alignment: 'justify', indent: { firstLine: 420 } },
      content: [{ text: 'Body copy', format: { fontSize: 12 } }],
    });
  });

  test('pastes multiline plain text into a paragraph and preserves the suffix', () => {
    const doc = createNativeWordDocument(createView());

    const result = nativeReplaceTextRangeWithParagraphs(
      doc,
      { paraId: 'P0000001', offset: 1 },
      { paraId: 'P0000001', offset: 1 },
      'one\ntwo\nthree',
    );

    expect(result?.cursor.offset).toBe(5);
    expect(result?.paragraphIds).toHaveLength(3);
    expect(paragraphTexts(doc).slice(0, 3)).toEqual(['Aone', 'two', 'threeC']);
  });

  test('keeps the insertion offset for a single-line replacement result', () => {
    const doc = createNativeWordDocument(createView());

    const result = nativeReplaceTextRangeWithParagraphs(
      doc,
      { paraId: 'P0000001', offset: 1 },
      { paraId: 'P0000001', offset: 1 },
      '',
    );

    expect(result?.cursor).toEqual({ paraId: 'P0000001', offset: 1 });
    expect(paragraphTexts(doc)[0]).toBe('AC');
  });

  test('replaces a cross-paragraph selection with multiline plain text', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks = [nativeTextParagraph('P3000001', 'Alpha'), nativeTextParagraph('P3000002', 'Beta')];

    const result = nativeReplaceTextRangeWithParagraphs(
      doc,
      { paraId: 'P3000001', offset: 2 },
      { paraId: 'P3000002', offset: 2 },
      'X\nY',
    );

    expect(result?.paragraphIds).toHaveLength(2);
    expect(paragraphTexts(doc)).toEqual(['AlX', 'Yta']);
  });

  test('applies run formatting to only the selected text range', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks = [nativeTextParagraph('P1000001', 'Alpha')];

    const change = nativeSetRunFormat(doc, 'P1000001', 1, 4, { bold: true, fontSize: 18, color: '00AA44' });
    const paragraph = doc.blocks[0];

    expect(change).toMatchObject({ kind: 'format', op: 'run', paraId: 'P1000001', from: 1, to: 4 });
    expect(paragraph).toMatchObject({
      kind: 'paragraph',
      content: [
        { kind: 'text', text: 'A' },
        { kind: 'text', text: 'lph', format: { bold: true, fontSize: 18, color: '00AA44' } },
        { kind: 'text', text: 'a' },
      ],
    });
    expect(nativeTextRangeIsBold(doc, 'P1000001', 1, 4)).toBe(true);
    expect(nativeTextRangeIsBold(doc, 'P1000001', 0, 5)).toBe(false);
  });

  test('clears only the selected run formatting', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks = [
      {
        kind: 'paragraph',
        paraId: 'P1000001',
        story: 'body',
        styleId: null,
        format: null,
        content: [{ kind: 'text', text: 'Alpha', format: { bold: true, italic: true, fontFamily: 'Arial' } }],
      },
    ];

    const change = nativeClearRunFormat(doc, 'P1000001', 1, 4);

    expect(change).toEqual({
      kind: 'format',
      op: 'run',
      paraId: 'P1000001',
      from: 1,
      to: 4,
      format: {},
      clear: true,
    });
    expect(findNativeParagraph(doc, 'P1000001')?.content).toEqual([
      { kind: 'text', text: 'A', format: { bold: true, italic: true, fontFamily: 'Arial' } },
      { kind: 'text', text: 'lph', format: null },
      { kind: 'text', text: 'a', format: { bold: true, italic: true, fontFamily: 'Arial' } },
    ]);
  });

  test('applies paragraph alignment without changing paragraph content', () => {
    const doc = createNativeWordDocument(createView());
    const before = paragraphTexts(doc)[0];

    const change = nativeSetParagraphFormat(doc, 'P0000001', { alignment: 'center' });

    expect(change).toEqual({
      kind: 'format',
      op: 'paragraph',
      paraId: 'P0000001',
      format: { alignment: 'center' },
    });
    expect(findNativeParagraph(doc, 'P0000001')?.format?.alignment).toBe('center');
    expect(paragraphTexts(doc)[0]).toBe(before);
  });

  test('merges nested paragraph indent and spacing fields', () => {
    const doc = createNativeWordDocument(createView());
    const paragraph = findNativeParagraph(doc, 'P0000001');
    if (!paragraph) throw new Error('paragraph fixture missing');
    paragraph.format = {
      ...(paragraph.format ?? {}),
      indent: { left: 720, right: 120, firstLine: 360 },
      spacing: { before: 120, after: 240, line: 276, lineRule: 'auto' },
    };

    nativeSetParagraphFormat(doc, 'P0000001', { indent: { left: 1440 } });
    nativeSetParagraphFormat(doc, 'P0000001', { spacing: { line: 360, lineRule: 'auto' } });

    expect(paragraph.format).toMatchObject({
      indent: { left: 1440, right: 120, firstLine: 360 },
      spacing: { before: 120, after: 240, line: 360, lineRule: 'auto' },
    });
  });

  test('applies and removes a structured paragraph list', () => {
    const doc = createNativeWordDocument(createView());
    const numbering = createNativeListNumbering(doc, 'bullet');

    const applied = nativeSetParagraphList(doc, 'P0000001', numbering);
    refreshNativeNumberingLabels(doc);

    expect(applied).toMatchObject({ kind: 'list', paraId: 'P0000001', numbering: { numFmt: 'bullet' } });
    expect(findNativeParagraph(doc, 'P0000001')?.format).toMatchObject({
      hasNumbering: true,
      numbering: { numId: numbering.numId, ilvl: 0, label: '•' },
      indent: { left: 720, firstLine: -360 },
    });

    expect(nativeSetParagraphList(doc, 'P0000001', null)).toEqual({
      kind: 'list',
      paraId: 'P0000001',
      numbering: null,
    });
    expect(findNativeParagraph(doc, 'P0000001')?.format).toMatchObject({
      hasNumbering: false,
      indent: { left: 0, firstLine: 0 },
    });
    expect(findNativeParagraph(doc, 'P0000001')?.format?.numbering).toBeUndefined();
  });

  test('recalculates numbered labels after splitting a list paragraph', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks = [nativeTextParagraph('P1000001', 'Alpha'), nativeTextParagraph('P1000002', 'Beta')];
    const numbering = createNativeListNumbering(doc, 'numbered');
    nativeSetParagraphList(doc, 'P1000001', numbering);
    nativeSetParagraphList(doc, 'P1000002', numbering);
    nativeSplitParagraph(doc, 'P1000001', 2, 'P1000003');

    refreshNativeNumberingLabels(doc);

    expect(
      doc.blocks.map((block) => block.kind === 'paragraph' ? block.format?.numbering?.label : null),
    ).toEqual(['1.', '2.', '3.']);
  });

  test('continues an adjacent native numbered list across separately converted paragraphs', () => {
    const doc = createNativeWordDocument(createView());
    doc.blocks = [
      nativeTextParagraph('P2000001', 'Alpha'),
      nativeTextParagraph('P2000002', 'Beta'),
      nativeTextParagraph('P2000003', 'Gamma'),
    ];
    const firstNumbering = createNativeListNumbering(doc, 'numbered');
    nativeSetParagraphList(doc, 'P2000001', firstNumbering);
    const continued = adjacentNativeListNumbering(doc, ['P2000002', 'P2000003'], 'numbered');

    expect(continued?.numId).toBe(firstNumbering.numId);
    nativeSetParagraphList(doc, 'P2000002', continued);
    nativeSetParagraphList(doc, 'P2000003', continued);
    refreshNativeNumberingLabels(doc);

    expect(
      doc.blocks.map((block) => block.kind === 'paragraph' ? block.format?.numbering?.label : null),
    ).toEqual(['1.', '2.', '3.']);
  });

  test('deletes an explicitly selected image by drawing id', () => {
    const doc = createNativeWordDocument(createView());

    const change = nativeDeleteImage(doc, 'P0000003', 'drawing-1');

    expect(change).toEqual({
      kind: 'image',
      op: 'delete',
      paraId: 'P0000003',
      offset: 1,
      drawingId: 'drawing-1',
    });
    expect(paragraphTexts(doc)[2]).toBe('XY');
  });

  test('renders table cell paragraphs as editable native content', () => {
    const doc = createNativeWordDocument(createView());
    const table = doc.blocks[3];

    expect(table).toMatchObject({
      kind: 'table',
      rows: [
        {
          cells: [
            {
              paragraphs: [
                {
                  paraId: 'P0000004',
                  content: [
                    { kind: 'text', text: 'M' },
                    { kind: 'image', drawingId: 'drawing-2' },
                    { kind: 'text', text: 'N' },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });

    expect(nativeInsertText(doc, 'P0000004', 1, 'Q')).toEqual({
      kind: 'text',
      op: 'insert',
      paraId: 'P0000004',
      offset: 1,
      text: 'Q',
    });
    expect(allParagraphTexts(doc)[3]).toBe('MQ[image]N');

    expect(nativeDeleteImage(doc, 'P0000004', 'drawing-2')).toEqual({
      kind: 'image',
      op: 'delete',
      paraId: 'P0000004',
      offset: 2,
      drawingId: 'drawing-2',
    });
    expect(allParagraphTexts(doc)[3]).toBe('MQN');
  });

  test('keeps table paragraph aliases synchronized for list rendering', () => {
    const doc = createNativeWordDocument(createView());
    const table = doc.blocks[3];
    if (!table || table.kind !== 'table') throw new Error('table fixture missing');
    const cell = table.rows[0]!.cells[0]!;
    const block = cell.blocks[0];
    if (!block || block.kind !== 'paragraph') throw new Error('table paragraph missing');

    expect(cell.paragraphs[0]).toBe(block);

    const numbering = createNativeListNumbering(doc, 'numbered');
    nativeSetParagraphList(doc, block.paraId, numbering);
    refreshNativeNumberingLabels(doc);

    expect(block.format?.numbering).toMatchObject({ numFmt: 'decimal', label: '1.' });
    expect(cell.paragraphs[0]!.format?.numbering).toMatchObject({ numFmt: 'decimal', label: '1.' });
  });

  test('preserves vertical merge metadata for native table rendering', () => {
    const view = createView();
    const table = view.blocks[3];
    if (!table || table.type !== 'table') throw new Error('table fixture missing');
    table.rows = [
      {
        cells: [
          {
            ...table.rows[0]!.cells[0]!,
            rowSpan: 2,
            vMerge: 'restart',
          },
          {
            ...table.rows[0]!.cells[0]!,
            text: 'A',
            vMerge: null,
          },
        ],
      },
      {
        cells: [
          {
            ...table.rows[0]!.cells[0]!,
            text: '',
            paragraphs: [],
            vMerge: 'continue',
          },
          {
            ...table.rows[0]!.cells[0]!,
            text: 'B',
            vMerge: null,
          },
        ],
      },
    ];

    const doc = createNativeWordDocument(view);
    const nativeTable = doc.blocks[3];
    if (!nativeTable || nativeTable.kind !== 'table') throw new Error('native table missing');

    expect(nativeTable.rows[0]!.cells[0]!.vMerge).toBe('restart');
    expect(nativeTable.rows[1]!.cells[0]!.vMerge).toBe('continue');
  });

  test('classifies native table row pagination risk for fragment planning', () => {
    const doc = createNativeWordDocument(createView());
    const table = doc.blocks[3];
    if (!table || table.kind !== 'table') throw new Error('native table missing');

    expect(analyzeNativeTableRowPagination(table.rows[0]!).kind).toBe('unsplittable');
    expect(analyzeNativeTableRowPagination(table.rows[0]!).risks).toContain('media');

    table.rows = [
      {
        isHeader: true,
        cells: [
          {
            ...table.rows[0]!.cells[0]!,
            blocks: [nativeTextParagraph('H1', 'Header')],
          },
        ],
      },
      {
        cantSplit: true,
        cells: [
          {
            ...table.rows[0]!.cells[0]!,
            blocks: [nativeTextParagraph('P1', 'Keep this row')],
          },
        ],
      },
      {
        cells: [
          {
            ...table.rows[0]!.cells[0]!,
            blocks: [nativeTextParagraph('P2', 'Text only')],
          },
        ],
      },
      {
        cells: [
          {
            ...table.rows[0]!.cells[0]!,
            blocks: [nativeTextParagraph('P3', 'First'), nativeTextParagraph('P4', 'Second')],
          },
        ],
      },
      {
        cells: [
          {
            ...table.rows[0]!.cells[0]!,
            rowSpan: 2,
            blocks: [nativeTextParagraph('P5', 'Merged')],
          },
        ],
      },
    ];

    expect(analyzeNativeTableRowPagination(table.rows[0]!).kind).toBe('repeat-header');
    expect(analyzeNativeTableRowPagination(table.rows[1]!).kind).toBe('keep-together');
    expect(analyzeNativeTableRowPagination(table.rows[1]!).risks).toContain('cant-split');
    expect(analyzeNativeTableRowPagination(table.rows[2]!).kind).toBe('text-fragmentable');
    expect(analyzeNativeTableRowPagination(table.rows[3]!).kind).toBe('block-fragmentable');
    expect(analyzeNativeTableRowPagination(table.rows[4]!).kind).toBe('unsplittable');
    expect(analyzeNativeTableRowPagination(table.rows[4]!).risks).toContain('row-span');

    expect(summarizeNativeTablePagination(doc)).toMatchObject({
      tables: 1,
      rows: 5,
      repeatHeaderRows: 1,
      keepTogetherRows: 1,
      blockFragmentableRows: 1,
      textFragmentableRows: 1,
      unsplittableRows: 1,
    });
  });

  test('splits native table rows at block boundaries for low-risk pagination', () => {
    const doc = createNativeWordDocument(createView());
    const table = doc.blocks[3];
    if (!table || table.kind !== 'table') throw new Error('native table missing');
    const sourceCell = table.rows[0]!.cells[0]!;
    const row = {
      cells: [
        {
          ...sourceCell,
          blocks: [
            nativeTextParagraph('P1', 'One'),
            nativeTextParagraph('P2', 'Two'),
            nativeTextParagraph('P3', 'Three'),
          ],
        },
        {
          ...sourceCell,
          blocks: [nativeTextParagraph('P4', 'Alpha'), nativeTextParagraph('P5', 'Beta')],
        },
      ],
    };

    const split = splitNativeTableRowAtBlockBoundary(row, { cellBlockCounts: [2, 1] });

    expect(split?.head.cells.map((cell) => cell.blocks.length)).toEqual([2, 1]);
    expect(split?.tail.cells.map((cell) => cell.blocks.length)).toEqual([1, 1]);
    expect(split?.head.cells[0]!.paragraphs.map((paragraph) => paragraph.paraId)).toEqual(['P1', 'P2']);
    expect(split?.tail.cells[0]!.paragraphs.map((paragraph) => paragraph.paraId)).toEqual(['P3']);
    expect(split?.head.cells[0]).toMatchObject({
      gridSpan: sourceCell.gridSpan,
      margin: sourceCell.margin,
      shading: sourceCell.shading,
    });
    expect(splitNativeTableRowAtBlockBoundary({ ...row, cantSplit: true }, { cellBlockCounts: [1, 1] })).toBeNull();
    expect(splitNativeTableRowAtBlockBoundary(row, { cellBlockCounts: [0, 1] })).toBeNull();
  });

  test('splits native table rows at text offsets while preserving run formats', () => {
    const doc = createNativeWordDocument(createView());
    const table = doc.blocks[3];
    if (!table || table.kind !== 'table') throw new Error('native table missing');
    const sourceCell = table.rows[0]!.cells[0]!;
    const row = {
      cells: [
        {
          ...sourceCell,
          blocks: [
            {
              kind: 'paragraph' as const,
              paraId: 'P1',
              story: 'body',
              styleId: null,
              format: null,
              content: [
                { kind: 'text' as const, text: 'Hello', format: { bold: true } },
                { kind: 'text' as const, text: 'World', format: { italic: true } },
              ],
            },
          ],
        },
        {
          ...sourceCell,
          blocks: [nativeTextParagraph('P2', 'AlphaBeta')],
        },
      ],
    };

    const split = splitNativeTableRowAtTextOffset(row, { cellTextOffsets: [7, 5] });

    const firstHead = split?.head.cells[0]!.paragraphs[0]!;
    const firstTail = split?.tail.cells[0]!.paragraphs[0]!;
    expect(firstHead.content).toEqual([
      { kind: 'text', text: 'Hello', format: { bold: true } },
      { kind: 'text', text: 'Wo', format: { italic: true } },
    ]);
    expect(firstTail.content).toEqual([{ kind: 'text', text: 'rld', format: { italic: true } }]);
    expect(split?.head.cells[1]!.paragraphs[0]!.content).toEqual([{ kind: 'text', text: 'Alpha', format: null }]);
    expect(split?.tail.cells[1]!.paragraphs[0]!.content).toEqual([{ kind: 'text', text: 'Beta', format: null }]);
    expect(splitNativeTableRowAtTextOffset(row, { cellTextOffsets: [0, 5] })).toBeNull();
    expect(splitNativeTableRowAtTextOffset({ ...row, cantSplit: true }, { cellTextOffsets: [4, 5] })).toBeNull();
  });

  test('counts top-level and table content in editor coverage', () => {
    const view = createView();
    const topLevelParagraphs = view.blocks.filter((block) => block.type === 'paragraph');
    const table = view.blocks.find((block) => block.type === 'table');
    if (!table || table.type !== 'table') throw new Error('table fixture missing');
    view.tables = [table];
    view.paragraphs = [
      ...topLevelParagraphs,
      ...table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)),
    ];
    const doc = createNativeWordDocument(view);

    expect(summarizeEditorCoverage(view)).toEqual({
      topLevelBlocks: 4,
      topLevelParagraphs: 3,
      tableBlocks: 1,
      tableParagraphs: 1,
      allParagraphs: 4,
      topLevelImages: 1,
      tableImages: 1,
      allImages: 2,
      topLevelDrawings: 0,
      tableDrawings: 0,
      allDrawings: 0,
      pageBreaks: 0,
    });
    expect(summarizeNativeCoverage(view, doc)).toEqual({
      editor: summarizeEditorCoverage(view),
      native: {
        blocks: 4,
        paragraphs: 4,
        images: 2,
        drawings: 0,
        pageBreaks: 0,
      },
      missing: {
        paragraphs: 0,
        images: 0,
        drawings: 0,
        pageBreaks: 0,
      },
    });
  });

  test('counts native table paragraphs from rendered cell blocks', () => {
    const view = createView();
    const topLevelParagraphs = view.blocks.filter((block) => block.type === 'paragraph');
    const table = view.blocks.find((block) => block.type === 'table');
    if (!table || table.type !== 'table') throw new Error('table fixture missing');
    view.tables = [table];
    view.paragraphs = [
      ...topLevelParagraphs,
      ...table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)),
    ];
    const doc = createNativeWordDocument(view);
    const nativeTable = doc.blocks.find((block) => block.kind === 'table');
    if (!nativeTable || nativeTable.kind !== 'table') throw new Error('native table missing');

    nativeTable.rows[0]!.cells[0]!.paragraphs = [];

    expect(summarizeNativeCoverage(view, doc).missing.paragraphs).toBe(0);
    expect(summarizeNativeCoverage(view, doc).native.paragraphs).toBe(4);
  });

  test('keeps table paragraphs without stable para ids in editor coverage', () => {
    const view = createView();
    const topLevelParagraphs = view.blocks.filter((block) => block.type === 'paragraph');
    const table = view.blocks.find((block) => block.type === 'table');
    if (!table || table.type !== 'table') throw new Error('table fixture missing');
    table.rows[0]!.cells[0]!.paragraphs[0]!.paraId = '';
    view.tables = [table];
    view.paragraphs = topLevelParagraphs;

    expect(summarizeEditorCoverage(view).tableParagraphs).toBe(1);
    expect(summarizeEditorCoverage(view).allParagraphs).toBe(4);
  });

  test('resolves word view styles into css declarations', () => {
    const view = createView();
    const paragraph = view.blocks[0];
    const table = view.blocks[3];
    if (paragraph?.type !== 'paragraph' || table?.type !== 'table') throw new Error('style fixture missing');

    expect(resolveParagraphStyle(paragraph.format).declarations).toMatchObject({
      'text-align': 'center',
      'margin-top': '12pt',
      'margin-bottom': '6pt',
      'line-height': '1.2',
      'font-size': '16pt',
      'font-weight': '700',
    });
    expect(resolveParagraphStyle(paragraph.format).declarations).not.toHaveProperty('color');
    expect(
      resolveParagraphStyle({
        hasNumbering: true,
        numbering: { numId: '1', ilvl: 0, label: '1.' },
        indent: { left: 0 },
      }).declarations,
    ).not.toHaveProperty('margin-left');
    expect(
      resolveParagraphStyle({
        borders: {
          top: { val: 'double', color: '00AA00', size: 8, space: 2 },
          right: { val: 'single', color: '0000FF', size: 4, space: 1 },
          bottom: { val: 'single', color: 'FF0000', size: 12, space: 0 },
          left: { val: 'single', color: '999999', size: 6, space: 3 },
        },
      }).declarations,
    ).toMatchObject({
      'border-top': '1pt double #00AA00',
      'padding-top': '2pt',
      'border-right': '0.75pt solid #0000FF',
      'padding-right': '1pt',
      'border-bottom': '1.5pt solid #FF0000',
      'padding-bottom': '0pt',
      'border-left': '0.75pt solid #999999',
      'padding-left': '3pt',
    });
    expect(
      resolveParagraphStyle({
        borders: {
          between: { val: 'single', color: 'CCCCCC', size: 4, space: 2 },
          bar: { val: 'single', color: '333333', size: 8, space: 1 },
        },
      }).declarations,
    ).toMatchObject({
      'border-bottom': '0.75pt solid #CCCCCC',
      'padding-bottom': '2pt',
      'border-left': '1pt solid #333333',
      'padding-left': '1pt',
    });
    expect(resolveParagraphStyle({ shading: { fill: 'E2F0D9', val: 'clear' } }).declarations).toMatchObject({
      'background-color': '#E2F0D9',
    });
    expect(resolveParagraphStyle({ characterSpacing: -10 }).declarations).toMatchObject({
      'letter-spacing': '-0.5pt',
    });
    expect(resolveParagraphStyle({ spacing: { line: 360, lineRule: 'auto' }, snapToGrid: true }).declarations).toMatchObject({
      'line-height': 'calc(var(--coco-word-doc-grid-line-height, 1em) * 1.5)',
    });
    expect(resolveTableCellStyle(table.rows[0]!.cells[0]!).declarations).toMatchObject({
      width: '72pt',
      'padding-top': '3pt',
      'padding-right': '6pt',
      'padding-bottom': '1.5pt',
      'padding-left': '6pt',
      'background-color': '#D9EAF7',
      'vertical-align': 'middle',
    });
    expect(resolveRunStyle({ bold: true, color: '00AA00', fontSize: 14, fontFamily: '黑体' }).declarations).toMatchObject({
      'font-weight': '700',
      'font-size': '14pt',
      'font-family': '"黑体", SimSun, "Songti SC", "Noto Serif CJK SC", serif',
      color: '#00AA00',
    });
    expect(resolveRunStyle({ fontFamily: 'Bad\\";color:red\nNext' }).declarations).toMatchObject({
      'font-family': '"Bad\\\\\\";color:red Next", SimSun, "Songti SC", "Noto Serif CJK SC", serif',
    });
    expect(resolveRunStyle({ shading: { fill: 'D9EAF7' } }).declarations).toMatchObject({
      'background-color': '#D9EAF7',
    });
    expect(resolveRunStyle({ highlight: 'yellow', shading: { fill: 'D9EAF7' } }).declarations).toMatchObject({
      'background-color': '#ffff00',
    });
    expect(resolveRunStyle({ characterSpacing: 40 }).declarations).toMatchObject({
      'letter-spacing': '2pt',
    });
    expect(resolveRunStyle({ verticalAlign: 'superscript' }).declarations).toMatchObject({
      'vertical-align': 'super',
      'font-size': 'smaller',
    });
    expect(resolveRunStyle({ verticalAlign: 'subscript', fontSize: 9 }).declarations).toMatchObject({
      'vertical-align': 'sub',
      'font-size': '9pt',
    });
    expect(resolveRunStyle({ underline: true, strike: true }).declarations).toMatchObject({
      'text-decoration-line': 'underline line-through',
    });
    expect(resolveRunStyle({ doubleStrike: true }).declarations).toMatchObject({
      'text-decoration-line': 'line-through',
      'text-decoration-style': 'double',
    });
    expect(resolveRunStyle({ caps: true, smallCaps: true }).declarations).toMatchObject({
      'text-transform': 'uppercase',
      'font-variant-caps': 'small-caps',
    });
  });

  test('maps table cell vertical alignment to valid css values', () => {
    expect(resolveTableCellStyle({ verticalAlignment: 'top' }).declarations).toMatchObject({
      'vertical-align': 'top',
    });
    expect(resolveTableCellStyle({ verticalAlignment: 'center' }).declarations).toMatchObject({
      'vertical-align': 'middle',
    });
    expect(resolveTableCellStyle({ verticalAlignment: 'bottom' }).declarations).toMatchObject({
      'vertical-align': 'bottom',
    });
    expect(resolveTableCellStyle({ verticalAlignment: 'both' }).declarations).toMatchObject({
      'vertical-align': 'middle',
    });
  });

  test('maps table cell noWrap to nowrap css', () => {
    expect(resolveTableCellStyle({ noWrap: true }).declarations).toMatchObject({
      'white-space': 'nowrap',
    });
    expect(resolveTableCellStyle({ noWrap: false }).declarations).not.toHaveProperty('white-space');
  });

  test('resolves table inside borders by cell position before direct cell border overrides', () => {
    const cell = {
      gridSpan: 1,
      vMerge: null,
      paragraphs: [],
    };
    const table = {
      kind: 'table' as const,
      tableId: 'tbl-border',
      columnWidths: [1200, 1200],
      borders: {
        insideH: { val: 'single' as const, color: 'FF0000', size: 8 },
        insideV: { val: 'single' as const, color: '00AA00', size: 8 },
      },
      rows: [
        {
          cells: [
            {
              ...cell,
              borders: { bottom: { val: 'single' as const, color: '0000FF', size: 16 } },
            },
            cell,
          ],
        },
        { cells: [cell, cell] },
      ],
    };

    const firstCellBorders = resolveNativeTableCellBorders(table, table.rows[0]!.cells[0]!, {
      rowIndex: 0,
      rowCount: 2,
      columnIndex: 0,
      columnSpan: 1,
      columnCount: 2,
    });
    expect(firstCellBorders).toEqual({
      top: { val: 'nil' },
      right: { val: 'single', color: '00AA00', size: 8 },
      bottom: { val: 'single', color: '0000FF', size: 16 },
      left: { val: 'nil' },
    });
    expect(resolveTableCellStyle({ ...cell, borders: firstCellBorders }).declarations).toMatchObject({
      'border-top': '0',
      'border-right': '1pt solid #00AA00',
      'border-bottom': '2pt solid #0000FF',
      'border-left': '0',
    });

    expect(
      resolveNativeTableCellBorders(table, table.rows[1]!.cells[1]!, {
        rowIndex: 1,
        rowCount: 2,
        columnIndex: 1,
        columnSpan: 1,
        columnCount: 2,
      }),
    ).toEqual({
      top: { val: 'single', color: 'FF0000', size: 8 },
      right: { val: 'nil' },
      bottom: { val: 'nil' },
      left: { val: 'single', color: '00AA00', size: 8 },
    });
  });
});
