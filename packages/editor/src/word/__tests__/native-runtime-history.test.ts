// @vitest-environment happy-dom

import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { allNativeParagraphs, renderNativeWordEditor, type NativeWordEditor } from '../native.js';
import { layoutNativePageAnnotationBubbles, refreshNativePageCommentBubbles } from '../native-renderer.js';
import type { EditorDocumentView, EditorParagraphView } from '../model.js';

function paragraph(paraId: string, text: string): EditorParagraphView {
  return {
    type: 'paragraph',
    paraId,
    story: 'body',
    text,
    flags: {
      hasBookmarks: false,
      hasComments: false,
      hasMath: false,
      hasRevisions: false,
    },
    segments: text ? [{ kind: 'text', from: 0, to: text.length, text }] : [],
    spans: [],
  };
}

function viewFixture(): EditorDocumentView {
  return {
    kind: 'word',
    sections: [],
    headers: [],
    footers: [],
    blocks: [paragraph('P1', 'Alpha'), paragraph('P2', 'Beta')],
    paragraphs: [],
    tables: [],
    detachedComments: [],
    detachedRevisions: [],
  };
}

function revisionViewFixture(): EditorDocumentView {
  const revised = paragraph('P1', 'Original replacement');
  revised.flags.hasRevisions = true;
  revised.segments = [
    {
      kind: 'revision',
      from: 0,
      to: 8,
      text: 'Original',
      revisionId: 'revision-delete',
      revisionKind: 'delete',
      author: 'Reviewer',
      date: '2026-10-23T00:00:00.000Z',
    },
    {
      kind: 'revision',
      from: 8,
      to: 20,
      text: ' replacement',
      revisionId: 'revision-insert',
      revisionKind: 'insert',
      author: 'Reviewer',
      date: '2026-10-23T00:00:00.000Z',
    },
  ];
  return { ...viewFixture(), blocks: [revised] };
}

function pageBreakViewFixture(): EditorDocumentView {
  const breakParagraph = paragraph('PB', '');
  breakParagraph.segments = [{
    kind: 'placeholder',
    from: 0,
    to: 1,
    placeholderKind: 'pageBreak',
    sourceTag: 'br',
  }];
  return {
    ...viewFixture(),
    blocks: [paragraph('P1', 'Before'), breakParagraph, paragraph('P2', 'After')],
  };
}

function emptyParagraphViewFixture(): EditorDocumentView {
  return {
    ...viewFixture(),
    blocks: [paragraph('P1', 'Alpha'), paragraph('E', ''), paragraph('P2', 'Beta')],
  };
}

function formattedClipboardViewFixture(): EditorDocumentView {
  const title = paragraph('P1', 'Red title');
  title.styleId = 'DocumentTitle';
  title.format = { alignment: 'center', spacing: { after: 240 } };
  title.segments = [{
    kind: 'text',
    from: 0,
    to: title.text.length,
    text: title.text,
    format: { bold: true, color: 'C00000', fontSize: 24 },
  }];
  const body = paragraph('P2', 'Body copy');
  body.styleId = 'BodyText';
  body.format = { alignment: 'justify', indent: { firstLine: 420 } };
  body.segments = [{
    kind: 'text',
    from: 0,
    to: body.text.length,
    text: body.text,
    format: { fontSize: 12 },
  }];
  return { ...viewFixture(), blocks: [title, body, paragraph('P3', '')] };
}

function largeViewFixture(count = 600): EditorDocumentView {
  return {
    ...viewFixture(),
    blocks: Array.from({ length: count }, (_, index) =>
      paragraph(`L${index}`, `Paragraph ${index} 中文内容 for native editor acceptance.`)),
  };
}

function visibleParagraphText(paragraphNode: ReturnType<NativeWordEditor['getDocument']>['blocks'][number]): string {
  if (paragraphNode.kind !== 'paragraph') return '';
  return paragraphNode.content.map((child) => {
    if (child.kind === 'text') return child.text;
    if (child.kind === 'field') return child.cachedText ?? '\uFFFC';
    if (child.kind === 'drawing' && child.text) return child.text;
    return '';
  }).join('');
}

function commentsForParagraph(document: ReturnType<NativeWordEditor['getDocument']>, paraId: string): Array<{ commentId: number }> {
  const paragraph = allNativeParagraphs(document).find((item) => item.paraId === paraId);
  return paragraph?.content.flatMap((item) => item.kind === 'text'
    ? item.marks?.filter((mark) => mark.kind === 'comment') ?? []
    : []) ?? [];
}

function expectModelAndDomToAgree(editor: NativeWordEditor, host: HTMLElement): void {
  const paragraphs = allNativeParagraphs(editor.getDocument());
  const modelIds = new Set(paragraphs.map((item) => item.paraId));
  for (const item of paragraphs) {
    const element = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${item.paraId}"]`);
    expect(element, `DOM paragraph ${item.paraId}`).not.toBeNull();
    expect(element?.textContent ?? '', `paragraph text ${item.paraId}`).toBe(visibleParagraphText(item));
    expect(element?.querySelectorAll('img').length ?? 0, `paragraph images ${item.paraId}`).toBe(
      item.content.filter((child) => child.kind === 'image').length,
    );
    expect(element?.querySelectorAll('[data-kind="page-break"]').length ?? 0, `paragraph breaks ${item.paraId}`).toBe(
      item.content.filter((child) => child.kind === 'pageBreak').length,
    );
  }
  for (const element of Array.from(host.querySelectorAll<HTMLElement>('coco-word-native-p[data-para-id]'))) {
    expect(modelIds.has(element.dataset.paraId ?? ''), `stale DOM paragraph ${element.dataset.paraId}`).toBe(true);
  }
}

function dispatchClipboardEvent(host: HTMLElement, type: 'copy' | 'paste', data: Map<string, string>): Event {
  const event = new Event(type, { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', {
    value: {
      files: [],
      getData: (format: string) => data.get(format) ?? '',
      setData: (format: string, value: string) => {
        data.set(format, value);
      },
    },
  });
  host.dispatchEvent(event);
  return event;
}

function paragraphKinds(editor: NativeWordEditor, paraId: string): string[] {
  const block = editor.getDocument().blocks.find(
    (candidate) => candidate.kind === 'paragraph' && candidate.paraId === paraId,
  );
  return block?.kind === 'paragraph' ? block.content.map((child) => child.kind) : [];
}

function pressHistoryKey(host: HTMLElement, key: 'z' | 'y', shiftKey = false): void {
  host.dispatchEvent(new KeyboardEvent('keydown', {
    key,
    ctrlKey: true,
    shiftKey,
    bubbles: true,
    cancelable: true,
  }));
}

describe('native runtime history integration', () => {
  let host: HTMLDivElement;
  let editor: NativeWordEditor;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    editor = renderNativeWordEditor(host, viewFixture(), { styleIsolation: 'none' });
  });

  afterEach(() => {
    editor.destroy();
    host.remove();
  });

  test('keeps selection available while blocking every mutation in read-only mode', () => {
    editor.setReadOnly(true);
    const before = editor.getDocument().blocks;

    expect(editor.isReadOnly()).toBe(true);
    expect(editor.setSelection({ kind: 'text-range', anchor: { paraId: 'P1', offset: 0 }, focus: { paraId: 'P1', offset: 5 } })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-text', text: 'blocked' })).toBe(false);
    expect(editor.replaceCurrent('blocked')).toBe(false);
    expect(editor.canDispatch({ kind: 'insert-page-break' })).toBe(false);
    expect(editor.canUndo()).toBe(false);
    expect(host.querySelector('[contenteditable="true"]')).toBeNull();
    expect(editor.getDocument().blocks).toEqual(before);

    const beforeInput = new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertText', data: 'x' });
    host.dispatchEvent(beforeInput);
    expect(beforeInput.defaultPrevented).toBe(true);

    editor.setReadOnly(false);
    expect(editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 5 })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-text', text: '!' })).toBe(true);
  });

  test('round-trips a page break through undo and redo', () => {
    expect(editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-page-break' })).toBe(true);
    expect(paragraphKinds(editor, 'P1')).toEqual(['text', 'pageBreak', 'text']);
    expect(editor.canUndo()).toBe(true);

    pressHistoryKey(host, 'z');
    expect(paragraphKinds(editor, 'P1')).toEqual(['text']);
    expect(editor.canRedo()).toBe(true);

    pressHistoryKey(host, 'y');
    expect(paragraphKinds(editor, 'P1')).toEqual(['text', 'pageBreak', 'text']);
  });

  test('round-trips a mixed formatting, list, image, and page-break sequence', () => {
    const initial = structuredClone(editor.getDocument());

    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P1', offset: 5 },
    });
    expect(editor.dispatch({ kind: 'toggle-bold' })).toBe(true);
    expect(editor.dispatch({ kind: 'set-text-color', color: '#C00000' })).toBe(true);
    expect(editor.dispatch({
      kind: 'set-hyperlink',
      hyperlink: { target: 'https://coco.example/docs' },
    })).toBe(true);

    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 5 });
    expect(editor.dispatch({ kind: 'toggle-list', listType: 'numbered' })).toBe(true);
    expect(editor.dispatch({ kind: 'set-paragraph-alignment', alignment: 'center' })).toBe(true);

    editor.setSelection({ kind: 'caret', paraId: 'P2', offset: 2 });
    expect(editor.dispatch({
      kind: 'insert-image',
      src: 'data:image/png;base64,AA==',
      width: 80,
      height: 40,
    })).toBe(true);
    const image = editor.getCommandState().image;
    expect(image).not.toBeNull();
    expect(editor.dispatch({
      kind: 'set-image-size',
      objectId: image!.objectId,
      width: 160,
      height: 90,
    })).toBe(true);
    expect(editor.dispatch({
      kind: 'set-image-alignment',
      objectId: image!.objectId,
      alignment: 'center',
    })).toBe(true);
    expect(editor.dispatch({
      kind: 'delete-object',
      objectKind: 'image',
      objectId: image!.objectId,
    })).toBe(true);

    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 });
    expect(editor.dispatch({ kind: 'insert-page-break' })).toBe(true);
    const final = structuredClone(editor.getDocument());

    for (let index = 0; index < 10; index += 1) expect(editor.undo()).toBe(true);
    expect(editor.getDocument()).toEqual(initial);
    expect(editor.canUndo()).toBe(false);

    for (let index = 0; index < 10; index += 1) expect(editor.redo(), `redo step ${index + 1}`).toBe(true);
    expect(editor.getDocument()).toEqual(final);
    expect(editor.canRedo()).toBe(false);
  });

  test('records programmatic text insertion in editor history', () => {
    const initial = structuredClone(editor.getDocument());
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 });

    expect(editor.dispatch({ kind: 'insert-text', text: '!' })).toBe(true);
    expect(editor.dispatch({
      kind: 'replace-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P1', offset: 2 },
      text: 'AL',
    })).toBe(true);
    expect(editor.dispatch({
      kind: 'delete-range',
      anchor: { paraId: 'P1', offset: 2 },
      focus: { paraId: 'P1', offset: 3 },
    })).toBe(true);
    expect(editor.canUndo()).toBe(true);
    expect(editor.undo()).toBe(true);
    expect(editor.undo()).toBe(true);
    expect(editor.undo()).toBe(true);
    expect(editor.getDocument()).toEqual(initial);
  });

  test('round-trips remaining run, paragraph, hyperlink, and list commands', () => {
    const initial = structuredClone(editor.getDocument());
    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P1', offset: 5 },
    });
    const commands = [
      { kind: 'toggle-italic' as const },
      { kind: 'toggle-underline' as const },
      { kind: 'toggle-strike' as const },
      { kind: 'set-font-family' as const, fontFamily: 'Microsoft YaHei' },
      { kind: 'set-font-size' as const, fontSize: 18 },
      { kind: 'set-highlight-color' as const, color: '#FFF2CC' },
      { kind: 'clear-text-format' as const },
      { kind: 'set-hyperlink' as const, hyperlink: { target: 'https://coco.example/p0' } },
      { kind: 'remove-hyperlink' as const },
      { kind: 'set-paragraph-alignment' as const, alignment: 'right' as const },
      { kind: 'set-paragraph-indent' as const, indent: { left: 360, right: 180, firstLine: 240 } },
      { kind: 'adjust-paragraph-left-indent' as const, delta: 180 },
      { kind: 'set-paragraph-spacing' as const, spacing: { before: 120, after: 240 } },
      { kind: 'set-paragraph-line-spacing' as const, line: 360, lineRule: 'exact' as const },
      { kind: 'toggle-list' as const, listType: 'numbered' as const },
      { kind: 'adjust-list-level' as const, delta: 1 },
      { kind: 'remove-list' as const },
    ];
    for (const command of commands) expect(editor.dispatch(command), command.kind).toBe(true);
    const final = structuredClone(editor.getDocument());

    for (let index = 0; index < commands.length; index += 1) expect(editor.undo(), `undo ${index + 1}`).toBe(true);
    expect(editor.getDocument()).toEqual(initial);
    for (let index = 0; index < commands.length; index += 1) expect(editor.redo(), `redo ${index + 1}`).toBe(true);
    expect(editor.getDocument()).toEqual(final);
  });

  test('round-trips table lifecycle, structure, and cell formatting', () => {
    const initial = structuredClone(editor.getDocument());
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 });

    expect(editor.dispatch({ kind: 'insert-table', rows: 2, columns: 2 })).toBe(true);
    const insertedTable = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(insertedTable?.kind).toBe('table');
    const firstCellParagraph = insertedTable?.kind === 'table' ? insertedTable.rows[0]?.cells[0]?.paragraphs[0] : null;
    expect(firstCellParagraph).toBeTruthy();
    editor.setSelection({ kind: 'caret', paraId: firstCellParagraph!.paraId, offset: 0 });
    expect(editor.dispatch({ kind: 'insert-text', text: 'Cell value' })).toBe(true);
    editor.setSelection({
      kind: 'table-cell-range',
      tableId: insertedTable!.tableId,
      anchor: { rowIndex: 0, columnIndex: 0 },
      focus: { rowIndex: 1, columnIndex: 1 },
    });
    expect(editor.dispatch({ kind: 'set-table-cell-background', color: '#FFF2CC' })).toBe(true);
    expect(editor.dispatch({ kind: 'set-table-cell-borders', color: '#4472C4' })).toBe(true);
    expect(editor.dispatch({ kind: 'clear-table-cells' })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-table-row', position: 'after' })).toBe(true);
    expect(editor.dispatch({ kind: 'delete-table-row' })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-table-column', position: 'after' })).toBe(true);
    expect(editor.dispatch({ kind: 'delete-table-column' })).toBe(true);
    expect(editor.dispatch({ kind: 'set-table-cell-vertical-alignment', alignment: 'center' })).toBe(true);

    const expandedTable = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(expandedTable?.kind === 'table' ? expandedTable.rows.length : 0).toBe(2);
    expect(expandedTable?.kind === 'table' ? expandedTable.rows[0]?.cells.length : 0).toBe(2);
    expect(editor.dispatch({ kind: 'delete-table' })).toBe(true);
    expect(editor.getDocument().blocks.some((block) => block.kind === 'table')).toBe(false);

    for (let index = 0; index < 11; index += 1) expect(editor.undo(), `table undo ${index + 1}`).toBe(true);
    expect(editor.getDocument()).toEqual(initial);
    expect(editor.canUndo()).toBe(false);

    for (let index = 0; index < 11; index += 1) expect(editor.redo(), `table redo step ${index + 1}`).toBe(true);
    expect(editor.getDocument().blocks.some((block) => block.kind === 'table')).toBe(false);
    expect(editor.canRedo()).toBe(false);
  });

  test('allows typing into newly inserted table rows and columns', () => {
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 });
    expect(editor.dispatch({ kind: 'insert-table', rows: 2, columns: 2 })).toBe(true);
    const table = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(table?.kind).toBe('table');
    if (table?.kind !== 'table') throw new Error('table fixture missing');

    const sourceParagraph = table.rows[0]?.cells[0]?.paragraphs[0];
    expect(sourceParagraph).toBeTruthy();
    editor.setSelection({ kind: 'caret', paraId: sourceParagraph!.paraId, offset: 0 });
    expect(editor.dispatch({ kind: 'insert-table-row', position: 'after' })).toBe(true);
    const rowSelection = editor.getSelection();
    expect(rowSelection?.kind).toBe('caret');
    if (rowSelection?.kind !== 'caret') throw new Error('inserted row selection missing');
    const newRowParagraph = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${rowSelection.paraId}"]`);
    expect(newRowParagraph?.firstChild).toBeTruthy();
    expect(newRowParagraph?.contentEditable).toBe('true');
    expect(newRowParagraph).not.toBeNull();
    newRowParagraph!.dispatchEvent(new InputEvent('beforeinput', {
      inputType: 'insertText',
      data: 'new row',
      bubbles: true,
      cancelable: true,
    }));
    (newRowParagraph!.firstChild as Text).data = 'new row';
    newRowParagraph!.dispatchEvent(new InputEvent('input', {
      inputType: 'insertText',
      data: 'new row',
      bubbles: true,
    }));

    expect(editor.dispatch({ kind: 'insert-table-column', position: 'after' })).toBe(true);
    const columnSelection = editor.getSelection();
    expect(columnSelection?.kind).toBe('caret');
    if (columnSelection?.kind !== 'caret') throw new Error('inserted column selection missing');
    const newColumnParagraph = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${columnSelection.paraId}"]`);
    expect(newColumnParagraph?.firstChild).toBeTruthy();
    expect(newColumnParagraph?.contentEditable).toBe('true');
    expect(editor.dispatch({ kind: 'insert-text', text: 'new column' })).toBe(true);

    const paragraphs = allNativeParagraphs(editor.getDocument());
    expect(paragraphs.some((paragraphNode) => visibleParagraphText(paragraphNode) === 'new row')).toBe(true);
    expect(paragraphs.some((paragraphNode) => visibleParagraphText(paragraphNode) === 'new column')).toBe(true);
  });

  test('merges and splits regular table cells with undo and redo', () => {
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 });
    expect(editor.dispatch({ kind: 'insert-table', rows: 1, columns: 2 })).toBe(true);
    const table = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(table?.kind).toBe('table');
    if (table?.kind !== 'table') throw new Error('table fixture missing');
    const first = table.rows[0]!.cells[0]!.paragraphs[0]!;
    const second = table.rows[0]!.cells[1]!.paragraphs[0]!;
    editor.setSelection({ kind: 'caret', paraId: first.paraId, offset: 0 });
    expect(editor.dispatch({ kind: 'insert-text', text: 'left' })).toBe(true);
    editor.setSelection({ kind: 'caret', paraId: second.paraId, offset: 0 });
    expect(editor.dispatch({ kind: 'insert-text', text: 'right' })).toBe(true);
    editor.setSelection({
      kind: 'table-cell-range',
      tableId: table.tableId,
      anchor: { rowIndex: 0, columnIndex: 0 },
      focus: { rowIndex: 0, columnIndex: 1 },
    });
    expect(editor.canDispatch({ kind: 'merge-table-cells' })).toBe(true);
    expect(editor.dispatch({ kind: 'merge-table-cells' })).toBe(true);
    const merged = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(merged?.kind === 'table' ? merged.rows[0]!.cells.length : 0).toBe(1);
    expect(merged?.kind === 'table' ? merged.rows[0]!.cells[0]!.gridSpan : 0).toBe(2);
    expect(merged?.kind === 'table' ? merged.rows[0]!.cells[0]!.paragraphs.map(visibleParagraphText).join('') : '').toContain('left');
    expect(merged?.kind === 'table' ? merged.rows[0]!.cells[0]!.paragraphs.map(visibleParagraphText).join('') : '').toContain('right');

    editor.setSelection({
      kind: 'table-cell-range',
      tableId: table.tableId,
      anchor: { rowIndex: 0, columnIndex: 0 },
      focus: { rowIndex: 0, columnIndex: 0 },
    });
    expect(editor.canDispatch({ kind: 'split-table-cell' })).toBe(true);
    expect(editor.dispatch({ kind: 'split-table-cell' })).toBe(true);
    const split = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(split?.kind === 'table' ? split.rows[0]!.cells.length : 0).toBe(2);
    expect(split?.kind === 'table' ? split.rows[0]!.cells.every((cell) => cell.gridSpan === 1) : false).toBe(true);
    expect(editor.undo()).toBe(true);
    expect(editor.undo()).toBe(true);
    const restored = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(restored?.kind === 'table' ? restored.rows[0]!.cells.length : 0).toBe(2);
    expect(editor.redo()).toBe(true);
    expect(editor.redo()).toBe(true);
    expect(editor.getDocument().blocks.some((block) => block.kind === 'table' && block.rows[0]!.cells.length === 2)).toBe(true);
  });

  test('selects and splits a vertically merged table cell through the DOM', () => {
    editor.destroy();
    const top = paragraph('VM1', 'Merged top');
    const lower = paragraph('VM2', '');
    const right1 = paragraph('VM3', 'Right 1');
    const right2 = paragraph('VM4', 'Right 2');
    editor = renderNativeWordEditor(host, {
      ...viewFixture(),
      blocks: [
        paragraph('P1', 'Alpha'),
        {
          type: 'table',
          tableId: 'vertical-table',
          columnWidths: [1800, 1800],
          rows: [
            { cells: [
              { text: top.text, gridSpan: 1, rowSpan: 2, vMerge: 'restart', blocks: [top], paragraphs: [top] },
              { text: right1.text, gridSpan: 1, vMerge: null, blocks: [right1], paragraphs: [right1] },
            ] },
            { cells: [
              { text: lower.text, gridSpan: 1, vMerge: 'continue', blocks: [lower], paragraphs: [lower] },
              { text: right2.text, gridSpan: 1, vMerge: null, blocks: [right2], paragraphs: [right2] },
            ] },
          ],
        },
      ],
    }, { styleIsolation: 'none' });
    const cell = host.querySelector<HTMLElement>('coco-word-native-table[data-table-id="vertical-table"] td[data-row-index="0"][data-column-index="0"]');
    expect(cell).not.toBeNull();
    cell?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 10, clientY: 10 }));
    expect(editor.canDispatch({ kind: 'split-table-cell' })).toBe(true);
    expect(editor.dispatch({ kind: 'split-table-cell' })).toBe(true);
    const table = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(table?.kind === 'table' ? table.rows[0]!.cells[0]!.vMerge : null).toBeUndefined();
    expect(table?.kind === 'table' ? table.rows[1]!.cells[0]!.vMerge : null).toBeUndefined();
    expect(host.querySelector('coco-word-native-table td[rowspan="2"]')).toBeNull();
  });

  test('tracks table cell structure edits and rejects them as one revision', () => {
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 });
    expect(editor.dispatch({ kind: 'insert-table', rows: 1, columns: 2 })).toBe(true);
    const table = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(table?.kind).toBe('table');
    if (table?.kind !== 'table') throw new Error('table fixture missing');
    editor.dispatch({ kind: 'toggle-track-revisions' });
    editor.setSelection({
      kind: 'table-cell-range',
      tableId: table.tableId,
      anchor: { rowIndex: 0, columnIndex: 0 },
      focus: { rowIndex: 0, columnIndex: 1 },
    });
    expect(editor.dispatch({ kind: 'merge-table-cells' })).toBe(true);
    const merged = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(merged?.kind === 'table' ? merged.revision?.revisionKind : null).toBe('insert');
    const revisionId = merged?.kind === 'table' ? merged.revision?.revisionId : undefined;
    expect(revisionId).toBeTruthy();
    expect(editor.dispatch({ kind: 'reject-revision', revisionId })).toBe(true);
    const rejected = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(rejected?.kind === 'table' ? rejected.rows[0]!.cells.length : 0).toBe(2);
  });

  test('rejects all consecutive tracked table structure edits back to the original table', () => {
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 });
    expect(editor.dispatch({ kind: 'insert-table', rows: 2, columns: 2 })).toBe(true);
    const inserted = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(inserted?.kind).toBe('table');
    if (inserted?.kind !== 'table') throw new Error('table fixture missing');
    const firstCell = inserted.rows[0]!.cells[0]!.paragraphs[0]!;
    editor.setSelection({ kind: 'caret', paraId: firstCell.paraId, offset: 0 });
    expect(editor.dispatch({ kind: 'toggle-track-revisions' })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-table-row', position: 'after' })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-table-column', position: 'after' })).toBe(true);

    const changed = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(changed?.kind === 'table' ? changed.rows.length : 0).toBe(3);
    expect(changed?.kind === 'table' ? changed.rows[0]!.cells.length : 0).toBe(3);
    expect(editor.dispatch({ kind: 'reject-all-revisions' })).toBe(true);

    const rejected = editor.getDocument().blocks.find((block) => block.kind === 'table');
    expect(rejected?.kind === 'table' ? rejected.rows.length : 0).toBe(2);
    expect(rejected?.kind === 'table' ? rejected.rows[0]!.cells.length : 0).toBe(2);
    expect(rejected?.kind === 'table' ? rejected.revision : null).toBeUndefined();
    expect(host.querySelectorAll('[data-revision-id]').length).toBe(0);
  });

  test('edits header and footer paragraphs through the native runtime', () => {
    editor.destroy();
    const header = paragraph('H1', 'Header');
    header.story = 'header';
    const footer = paragraph('F1', 'Footer');
    footer.story = 'footer';
    editor = renderNativeWordEditor(host, {
      ...viewFixture(),
      headers: [{ kind: 'header', partName: 'word/header1.xml', blocks: [header], paragraphs: [header], tables: [] }],
      footers: [{ kind: 'footer', partName: 'word/footer1.xml', blocks: [footer], paragraphs: [footer], tables: [] }],
    }, { styleIsolation: 'none' });
    editor.setSelection({ kind: 'caret', paraId: 'H1', offset: 6 });
    expect(editor.dispatch({ kind: 'insert-text', text: '!' })).toBe(true);
    editor.setSelection({ kind: 'caret', paraId: 'F1', offset: 6 });
    expect(editor.dispatch({ kind: 'insert-text', text: '?' })).toBe(true);
    const headerNode = allNativeParagraphs(editor.getDocument()).find((item) => item.paraId === 'H1');
    const footerNode = allNativeParagraphs(editor.getDocument()).find((item) => item.paraId === 'F1');
    expect(headerNode ? visibleParagraphText(headerNode) : '').toBe('Header!');
    expect(footerNode ? visibleParagraphText(footerNode) : '').toBe('Footer?');
    expect(host.querySelector('coco-word-native-story[data-story-kind="header"] [data-para-id="H1"]')).not.toBeNull();
    expect(host.querySelector('coco-word-native-story[data-story-kind="footer"] [data-para-id="F1"]')).not.toBeNull();
  });

  test('adds comments to selected text and round-trips through native history', async () => {
    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P1', offset: 5 },
    });
    expect(editor.dispatch({ kind: 'add-comment', body: 'Review this text', author: 'Tester' })).toBe(true);
    const paragraphNode = allNativeParagraphs(editor.getDocument()).find((item) => item.paraId === 'P1');
    const markedText = paragraphNode?.content.find((item) => item.kind === 'text' && item.marks?.some((mark) => mark.kind === 'comment'));
    expect(markedText?.kind).toBe('text');
    expect(markedText?.marks?.[0]).toMatchObject({ kind: 'comment', body: 'Review this text', author: 'Tester' });
    expect(host.querySelector('.coco-word-comment[data-comment-body="Review this text"]')).not.toBeNull();

    const commentBubble = host.querySelector<HTMLElement>('.coco-word-native-comment-bubble');
    commentBubble?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(editor.getSelection()).toMatchObject({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P1', offset: 5 },
    });
    expect(window.getSelection()?.isCollapsed).toBe(false);
    await Promise.resolve();

    const interactiveBubble = document.createElement('aside');
    interactiveBubble.className = 'coco-word-native-comment-bubble';
    interactiveBubble.dataset.paraId = 'P1';
    interactiveBubble.dataset.from = '0';
    interactiveBubble.dataset.to = '5';
    interactiveBubble.dataset.commentBody = 'Review this text';
    const actions = document.createElement('div');
    const menuButton = document.createElement('button');
    menuButton.dataset.commentMenu = 'true';
    menuButton.dataset.commentId = '1';
    const menu = document.createElement('div');
    menu.dataset.commentMenuFor = '1';
    menu.hidden = true;
    const editButton = document.createElement('button');
    editButton.dataset.commentAction = 'edit';
    editButton.dataset.commentId = '1';
    menu.appendChild(editButton);
    actions.append(menuButton, menu);
    interactiveBubble.appendChild(actions);
    host.appendChild(interactiveBubble);

    menuButton.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(menu.hidden).toBe(false);
    editButton.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(document.querySelector('#coco-word-comment-dialog-title')?.textContent).toBe('编辑批注');
    document.querySelector('.coco-word-dialog-backdrop')?.remove();
    interactiveBubble.remove();

    expect(editor.undo()).toBe(true);
    expect(host.querySelector('.coco-word-comment')).toBeNull();
    expect(editor.redo()).toBe(true);
    expect(host.querySelector('.coco-word-comment[data-comment-body="Review this text"]')).not.toBeNull();
  });

  test('packs comments and revisions into one non-overlapping annotation rail', () => {
    const page = document.createElement('coco-word-native-page');
    const commentLayer = document.createElement('div');
    commentLayer.className = 'coco-word-native-comment-layer';
    const revisionLayer = document.createElement('div');
    revisionLayer.className = 'coco-word-native-revision-layer';
    page.append(commentLayer, revisionLayer);
    Object.defineProperty(page, 'getBoundingClientRect', {
      value: () => ({ width: 800, height: 160, top: 0, right: 800, bottom: 160, left: 0, x: 0, y: 0, toJSON: () => ({}) }),
    });

    const addBubble = (layer: HTMLElement, className: string, anchorTop: number, height: number) => {
      const bubble = document.createElement('aside');
      bubble.className = className;
      bubble.dataset.annotationAnchorTop = String(anchorTop);
      Object.defineProperty(bubble, 'getBoundingClientRect', {
        value: () => ({ width: 220, height, top: 0, right: 220, bottom: height, left: 0, x: 0, y: 0, toJSON: () => ({}) }),
      });
      layer.appendChild(bubble);
      return bubble;
    };
    const comment = addBubble(commentLayer, 'coco-word-native-comment-bubble', 110, 60);
    const deletion = addBubble(revisionLayer, 'coco-word-native-revision-bubble is-delete', 115, 60);
    const insertion = addBubble(revisionLayer, 'coco-word-native-revision-bubble is-insert', 120, 60);

    layoutNativePageAnnotationBubbles(page);

    expect(comment.style.getPropertyValue('--coco-word-comment-top')).toBe('0px');
    expect(deletion.style.getPropertyValue('--coco-word-revision-top')).toBe('68px');
    expect(insertion.style.getPropertyValue('--coco-word-revision-top')).toBe('136px');
    expect(page.style.getPropertyValue('--coco-word-native-annotation-overflow')).toBe('');
  });

  test('groups revisions by paragraph and expands them without changing page height', () => {
    const root = host.querySelector<HTMLElement>('coco-word-native-document')!;
    const paragraphNode = root.querySelector<HTMLElement>('coco-word-native-p[data-para-id="P1"]')!;
    paragraphNode.innerHTML = [
      '<span class="coco-word-revision-delete" data-revision-id="R1" data-revision-kind="delete" data-revision-text="old">old</span>',
      '<span class="coco-word-revision-insert" data-revision-id="R2" data-revision-kind="insert" data-revision-text="new">new</span>',
    ].join('');
    const page = document.createElement('coco-word-native-page');
    const pageBody = document.createElement('coco-word-native-page-body');
    pageBody.appendChild(paragraphNode);
    page.appendChild(pageBody);
    root.appendChild(page);
    refreshNativePageCommentBubbles(root);

    const group = host.querySelector<HTMLElement>('.coco-word-native-revision-group');
    const toggle = group?.querySelector<HTMLButtonElement>('[data-revision-group-toggle]');
    const list = group?.querySelector<HTMLElement>('.coco-word-native-revision-group-list');
    expect(group).not.toBeNull();
    expect(toggle?.textContent).toContain('本段 2 条修订');
    expect(list?.hidden).toBe(true);
    expect(page.style.getPropertyValue('--coco-word-native-annotation-overflow')).toBe('');

    toggle?.click();
    expect(group?.classList.contains('is-expanded')).toBe(true);
    expect(toggle?.getAttribute('aria-expanded')).toBe('true');
    expect(list?.hidden).toBe(false);
    expect(list?.querySelectorAll('.coco-word-native-revision-bubble')).toHaveLength(2);
    expect(page.style.getPropertyValue('--coco-word-native-annotation-overflow')).toBe('');

    toggle?.click();
    expect(list?.hidden).toBe(true);
    paragraphNode.querySelector<HTMLElement>('[data-revision-id="R1"]')?.click();
    expect(group?.classList.contains('is-expanded')).toBe(true);
    expect(list?.hidden).toBe(false);
  });

  test('edits and deletes a comment through native history', () => {
    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P1', offset: 5 },
    });
    expect(editor.dispatch({ kind: 'add-comment', body: 'Original', author: 'Tester' })).toBe(true);
    const mark = allNativeParagraphs(editor.getDocument()).flatMap((item) => item.content)
      .find((item) => item.kind === 'text' && item.marks?.some((value) => value.kind === 'comment'));
    const commentId = mark?.kind === 'text' ? mark.marks?.find((value) => value.kind === 'comment')?.commentId : undefined;
    expect(commentId).toBeTypeOf('number');

    expect(editor.dispatch({ kind: 'edit-comment', commentId: commentId!, body: 'Updated' })).toBe(true);
    expect(host.querySelector('.coco-word-comment[data-comment-body="Updated"]')).not.toBeNull();
    expect(editor.undo()).toBe(true);
    expect(host.querySelector('.coco-word-comment[data-comment-body="Original"]')).not.toBeNull();
    expect(editor.redo()).toBe(true);
    expect(host.querySelector('.coco-word-comment[data-comment-body="Updated"]')).not.toBeNull();

    expect(editor.dispatch({ kind: 'delete-comment', commentId: commentId! })).toBe(true);
    expect(host.querySelector('.coco-word-comment')).toBeNull();
    expect(editor.undo()).toBe(true);
    expect(host.querySelector('.coco-word-comment[data-comment-body="Updated"]')).not.toBeNull();
  });

  test('keeps a cross-paragraph comment as one logical annotation', () => {
    editor.destroy();
    editor = renderNativeWordEditor(host, pageBreakViewFixture(), { styleIsolation: 'none' });
    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 2 },
      focus: { paraId: 'P2', offset: 3 },
    });
    expect(editor.dispatch({ kind: 'add-comment', body: 'Across paragraphs', author: 'Tester' })).toBe(true);
    const marked = allNativeParagraphs(editor.getDocument()).flatMap((paragraph) => paragraph.content)
      .filter((item) => item.kind === 'text' && item.marks?.some((mark) => mark.kind === 'comment'));
    const commentIds = new Set(marked.flatMap((item) => item.kind === 'text'
      ? item.marks?.filter((mark) => mark.kind === 'comment').map((mark) => mark.commentId) ?? []
      : []));
    expect(commentIds.size).toBe(1);
    expect(host.querySelectorAll('.coco-word-comment')).toHaveLength(2);
    expect(editor.undo()).toBe(true);
    expect(host.querySelectorAll('.coco-word-comment')).toHaveLength(0);
    expect(editor.redo()).toBe(true);
    expect(host.querySelectorAll('.coco-word-comment')).toHaveLength(2);
  });

  test('keeps a reply linked to its parent through native history', () => {
    expect(editor.setSelection({ kind: 'text-range', anchor: { paraId: 'P1', offset: 0 }, focus: { paraId: 'P1', offset: 5 } })).toBe(true);
    expect(editor.dispatch({ kind: 'add-comment', body: 'Parent', author: 'Tester' })).toBe(true);
    const parent = allNativeParagraphs(editor.getDocument())[0]!.content
      .find((item) => item.kind === 'text' && item.marks?.some((mark) => mark.kind === 'comment'));
    const parentId = parent?.kind === 'text' ? parent.marks?.find((mark) => mark.kind === 'comment')?.commentId : undefined;
    expect(parentId).toBeTypeOf('number');
    expect(editor.dispatch({ kind: 'reply-comment', commentId: parentId!, body: 'Reply' })).toBe(true);
    const comments = allNativeParagraphs(editor.getDocument())[0]!.content
      .flatMap((item) => item.kind === 'text' ? item.marks?.filter((mark) => mark.kind === 'comment') ?? [] : []);
    const reply = comments.find((mark) => mark.body === 'Reply');
    expect(reply).toMatchObject({ parentCommentId: parentId });
    expect(editor.undo()).toBe(true);
    expect(commentsForParagraph(editor.getDocument(), 'P1')).toHaveLength(1);
    expect(editor.redo()).toBe(true);
    expect(commentsForParagraph(editor.getDocument(), 'P1')).toHaveLength(3);
    const replyId = commentsForParagraph(editor.getDocument(), 'P1').find((mark) => mark.body === 'Reply')?.commentId;
    expect(replyId).toBeTypeOf('number');
    expect(editor.dispatch({ kind: 'reply-comment', commentId: replyId!, body: 'Nested reply' })).toBe(true);
    const nested = commentsForParagraph(editor.getDocument(), 'P1').find((mark) => mark.body === 'Nested reply');
    expect(nested).toMatchObject({ parentCommentId: replyId });
    expect(editor.dispatch({ kind: 'delete-comment', commentId: parentId! })).toBe(true);
    const remainingIds = new Set(commentsForParagraph(editor.getDocument(), 'P1').map((mark) => mark.commentId));
    expect(remainingIds).toEqual(new Set([replyId, nested?.commentId]));
  });

  test('marks selected text as insert or delete revision and round-trips history', () => {
    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P1', offset: 5 },
    });
    expect(editor.dispatch({ kind: 'add-revision', revisionKind: 'insert', author: 'Tester' })).toBe(true);
    expect(host.querySelector('.coco-word-revision-insert')).not.toBeNull();
    expect(editor.undo()).toBe(true);
    expect(host.querySelector('.coco-word-revision-insert')).toBeNull();
    expect(editor.redo()).toBe(true);
    expect(host.querySelector('.coco-word-revision-insert')).not.toBeNull();
    const revisionId = host.querySelector<HTMLElement>('.coco-word-revision-insert')?.dataset.revisionId;
    expect(revisionId).toBeTruthy();

    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P2', offset: 0 },
      focus: { paraId: 'P2', offset: 4 },
    });
    expect(editor.dispatch({ kind: 'add-revision', revisionKind: 'delete' })).toBe(true);
    expect(host.querySelector('.coco-word-revision-delete')).not.toBeNull();

    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P1', offset: 5 },
    });
    expect(editor.dispatch({ kind: 'accept-revision' })).toBe(true);
    expect(host.querySelector('.coco-word-revision-insert')).toBeNull();

    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P2', offset: 0 },
      focus: { paraId: 'P2', offset: 4 },
    });
    expect(editor.dispatch({ kind: 'reject-revision' })).toBe(true);
    expect(host.querySelector('.coco-word-revision-delete')).toBeNull();
  });

  test('tracks direct insert and delete commands while revision mode is enabled', () => {
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 5 });
    expect(editor.getCommandState().trackRevisions).toBe(false);
    expect(editor.dispatch({ kind: 'toggle-track-revisions' })).toBe(true);
    expect(editor.getCommandState().trackRevisions).toBe(true);
    expect(editor.dispatch({ kind: 'insert-text', text: '!' })).toBe(true);
    expect(host.querySelector('.coco-word-revision-insert')).not.toBeNull();
    expect(editor.dispatch({ kind: 'delete-backward' })).toBe(true);
    expect(host.querySelector('.coco-word-revision-insert')).toBeNull();
    expect(host.querySelector('.coco-word-revision-delete')).toBeNull();
    expect(visibleParagraphText(editor.getDocument().blocks.find((item) => item.kind === 'paragraph' && item.paraId === 'P1')!)).toBe('Alpha');
    expect(editor.undo()).toBe(true);
    expect(host.querySelector('.coco-word-revision-insert')).not.toBeNull();
    expect(editor.redo()).toBe(true);
    expect(host.querySelector('.coco-word-revision-insert')).toBeNull();
  });

  test('cancels a tracked insertion when the inserted text is deleted', () => {
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 5 });
    expect(editor.dispatch({ kind: 'toggle-track-revisions' })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-text', text: '123' })).toBe(true);
    expect(editor.dispatch({ kind: 'delete-range', anchor: { paraId: 'P1', offset: 5 }, focus: { paraId: 'P1', offset: 8 } })).toBe(true);

    const block = editor.getDocument().blocks.find((item) => item.kind === 'paragraph' && item.paraId === 'P1');
    expect(visibleParagraphText(block!)).toBe('Alpha');
    expect(host.querySelector('.coco-word-revision-insert')).toBeNull();
    expect(host.querySelector('.coco-word-revision-delete')).toBeNull();
  });

  test('coalesces consecutive tracked typing into one revision', () => {
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 0 });
    expect(editor.dispatch({ kind: 'toggle-track-revisions' })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-text', text: '1' })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-text', text: '2' })).toBe(true);
    expect(editor.dispatch({ kind: 'insert-text', text: '3' })).toBe(true);
    const insertions = Array.from(host.querySelectorAll<HTMLElement>('.coco-word-revision-insert'));
    expect(new Set(insertions.map((item) => item.dataset.revisionId)).size).toBe(1);
    expect(insertions.map((item) => item.textContent).join('')).toBe('123');
  });

  test('coalesces consecutive tracked backward deletes into one revision', () => {
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 5 });
    expect(editor.dispatch({ kind: 'toggle-track-revisions' })).toBe(true);
    expect(editor.dispatch({ kind: 'delete-backward' })).toBe(true);
    expect(editor.dispatch({ kind: 'delete-backward' })).toBe(true);
    expect(editor.dispatch({ kind: 'delete-backward' })).toBe(true);
    const deletions = Array.from(host.querySelectorAll<HTMLElement>('.coco-word-revision-delete'));
    expect(new Set(deletions.map((item) => item.dataset.revisionId)).size).toBe(1);
  });

  test('tracks a cross-paragraph selection delete without removing the text', () => {
    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 2 },
      focus: { paraId: 'P2', offset: 2 },
    });
    expect(editor.dispatch({ kind: 'toggle-track-revisions' })).toBe(true);
    expect(editor.dispatch({
      kind: 'delete-range',
      anchor: { paraId: 'P1', offset: 2 },
      focus: { paraId: 'P2', offset: 2 },
    })).toBe(true);
    const deletions = Array.from(host.querySelectorAll<HTMLElement>('.coco-word-revision-delete'));
    expect(deletions.length).toBeGreaterThan(0);
    expect(new Set(deletions.map((item) => item.dataset.revisionId)).size).toBe(1);
    const textOf = (paraId: string) => allNativeParagraphs(editor.getDocument())
      .find((item) => item.paraId === paraId)?.content
      .filter((item) => item.kind === 'text')
      .map((item) => item.text)
      .join('');
    expect(textOf('P1')).toBe('Alpha');
    expect(textOf('P2')).toBe('Beta');
  });

  test('round-trips deleting a standalone page break from the next paragraph', () => {
    editor.destroy();
    editor = renderNativeWordEditor(host, pageBreakViewFixture(), { styleIsolation: 'none' });
    expect(paragraphKinds(editor, 'PB')).toEqual(['pageBreak']);
    editor.setSelection({ kind: 'caret', paraId: 'P2', offset: 0 });

    expect(editor.dispatch({ kind: 'delete-backward' })).toBe(true);
    expect(paragraphKinds(editor, 'PB')).toEqual([]);
    expect(editor.undo()).toBe(true);
    expect(paragraphKinds(editor, 'PB')).toEqual(['pageBreak']);
    expect(editor.redo()).toBe(true);
    expect(paragraphKinds(editor, 'PB')).toEqual([]);
  });

  test('preserves paragraph and run formatting when copying and pasting ordinary paragraphs', () => {
    editor.destroy();
    editor = renderNativeWordEditor(host, formattedClipboardViewFixture(), { styleIsolation: 'none' });
    const clipboard = new Map<string, string>();
    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 0 },
      focus: { paraId: 'P2', offset: 'Body copy'.length },
    });

    expect(dispatchClipboardEvent(host, 'copy', clipboard).defaultPrevented).toBe(true);
    expect(clipboard.has('application/x-coco-word-blocks+json')).toBe(true);

    editor.setSelection({ kind: 'caret', paraId: 'P3', offset: 0 });
    expect(dispatchClipboardEvent(host, 'paste', clipboard).defaultPrevented).toBe(true);

    const blocks = editor.getDocument().blocks;
    const targetIndex = blocks.findIndex((block) => block.kind === 'paragraph' && block.paraId === 'P3');
    expect(blocks[targetIndex]).toMatchObject({
      kind: 'paragraph',
      styleId: 'DocumentTitle',
      format: { alignment: 'center', spacing: { after: 240 } },
      content: [{ text: 'Red title', format: { bold: true, color: 'C00000', fontSize: 24 } }],
    });
    expect(blocks[targetIndex + 1]).toMatchObject({
      kind: 'paragraph',
      styleId: 'BodyText',
      format: { alignment: 'justify', indent: { firstLine: 420 } },
      content: [{ text: 'Body copy', format: { fontSize: 12 } }],
    });
    expect(editor.undo()).toBe(true);
    expect(editor.getDocument().blocks).toHaveLength(3);
    expect(editor.redo()).toBe(true);
    expect(editor.getDocument().blocks).toHaveLength(4);
  });

  test('copies a large image through a lightweight internal token and restores deletion history', () => {
    const src = `data:image/png;base64,${'A'.repeat(2_000_000)}`;
    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 2 });
    expect(editor.dispatch({ kind: 'insert-image', src, width: 120, height: 80 })).toBe(true);
    const inserted = editor.getCommandState().image;
    expect(inserted).not.toBeNull();
    const clipboard = new Map<string, string>();

    expect(dispatchClipboardEvent(host, 'copy', clipboard).defaultPrevented).toBe(true);
    const internal = clipboard.get('application/x-coco-word-image') ?? '';
    expect(internal).toMatch(/^coco-image:/);
    expect(internal.length).toBeLessThan(100);

    editor.setSelection({ kind: 'caret', paraId: 'P2', offset: 2 });
    const paragraphElement = host.querySelector('coco-word-native-p[data-para-id="P2"]');
    expect(dispatchClipboardEvent(host, 'paste', clipboard).defaultPrevented).toBe(true);
    expect(host.querySelector('coco-word-native-p[data-para-id="P2"]')).toBe(paragraphElement);
    const pastedParagraph = editor.getDocument().blocks.find(
      (block) => block.kind === 'paragraph' && block.paraId === 'P2',
    );
    expect(pastedParagraph?.kind === 'paragraph' ? pastedParagraph.content.some(
      (child) => child.kind === 'image' && child.src === src,
    ) : false).toBe(true);

    const pastedImage = pastedParagraph?.kind === 'paragraph'
      ? pastedParagraph.content.find((child) => child.kind === 'image')
      : null;
    expect(pastedImage?.kind).toBe('image');
    if (pastedImage?.kind !== 'image') throw new Error('pasted image missing');
    editor.setSelection({
      kind: 'object',
      paraId: 'P2',
      objectKind: 'image',
      objectId: pastedImage.drawingId,
      offset: 2,
    });
    expect(editor.dispatch({ kind: 'delete-object', objectKind: 'image', objectId: pastedImage.drawingId })).toBe(true);
    expect(editor.undo()).toBe(true);
    expect(editor.redo()).toBe(true);
  });

  test('keeps a large document model and DOM consistent through a mixed edit history', () => {
    editor.destroy();
    editor = renderNativeWordEditor(host, largeViewFixture(), { styleIsolation: 'none' });
    const initial = structuredClone(editor.getDocument());
    let operationCount = 0;
    const apply = (result: boolean) => {
      expect(result).toBe(true);
      operationCount += 1;
      expectModelAndDomToAgree(editor, host);
    };

    editor.setSelection({ kind: 'caret', paraId: 'L10', offset: 10 });
    apply(editor.dispatch({ kind: 'insert-text', text: '快速输入中文' }));

    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'L100', offset: 5 },
      focus: { paraId: 'L102', offset: 12 },
    });
    apply(editor.dispatch({ kind: 'delete-backward' }));

    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'L200', offset: 0 },
      focus: { paraId: 'L205', offset: 20 },
    });
    apply(editor.dispatch({ kind: 'toggle-bold' }));
    apply(editor.dispatch({ kind: 'set-highlight-color', color: '#FFF2CC' }));
    apply(editor.dispatch({ kind: 'set-paragraph-alignment', alignment: 'both' }));

    editor.setSelection({ kind: 'caret', paraId: 'L300', offset: 8 });
    apply(editor.dispatch({ kind: 'insert-image', src: 'data:image/png;base64,AA==', width: 96, height: 64 }));
    const image = editor.getCommandState().image;
    expect(image).not.toBeNull();
    apply(editor.dispatch({ kind: 'set-image-size', objectId: image!.objectId, width: 144, height: 96 }));
    apply(editor.dispatch({ kind: 'delete-object', objectKind: 'image', objectId: image!.objectId }));

    editor.setSelection({ kind: 'caret', paraId: 'L400', offset: 4 });
    apply(editor.dispatch({ kind: 'insert-page-break' }));

    editor.setSelection({ kind: 'caret', paraId: 'L500', offset: 6 });
    apply(editor.dispatch({ kind: 'insert-table', rows: 3, columns: 3 }));
    apply(editor.dispatch({ kind: 'set-table-cell-background', color: '#D9EAF7' }));
    const final = structuredClone(editor.getDocument());

    for (let index = 0; index < operationCount; index += 1) {
      expect(editor.undo(), `large undo ${index + 1}`).toBe(true);
      expectModelAndDomToAgree(editor, host);
    }
    expect(editor.getDocument()).toEqual(initial);
    for (let index = 0; index < operationCount; index += 1) {
      expect(editor.redo(), `large redo ${index + 1}`).toBe(true);
      expectModelAndDomToAgree(editor, host);
    }
    expect(editor.getDocument()).toEqual(final);
  });

  test('handles empty paragraphs, paragraph boundaries, and selected Delete symmetrically', () => {
    editor.destroy();
    editor = renderNativeWordEditor(host, emptyParagraphViewFixture(), { styleIsolation: 'none' });
    const initial = structuredClone(editor.getDocument());

    editor.setSelection({ kind: 'caret', paraId: 'E', offset: 0 });
    expect(editor.dispatch({ kind: 'delete-backward' })).toBe(true);
    expect(allNativeParagraphs(editor.getDocument()).map((item) => item.paraId)).toEqual(['P1', 'P2']);
    expect(editor.undo()).toBe(true);
    expect(editor.getDocument()).toEqual(initial);

    editor.setSelection({ kind: 'caret', paraId: 'P1', offset: 'Alpha'.length });
    expect(editor.dispatch({ kind: 'delete-forward' })).toBe(true);
    expect(allNativeParagraphs(editor.getDocument()).map((item) => item.paraId)).toEqual(['P1', 'P2']);
    expect(editor.undo()).toBe(true);
    expect(editor.getDocument()).toEqual(initial);

    editor.setSelection({
      kind: 'text-range',
      anchor: { paraId: 'P1', offset: 1 },
      focus: { paraId: 'P2', offset: 2 },
    });
    expect(editor.dispatch({ kind: 'delete-forward' })).toBe(true);
    expect(allNativeParagraphs(editor.getDocument()).map(visibleParagraphText)).toEqual(['Ata']);
    expectModelAndDomToAgree(editor, host);
    expect(editor.undo()).toBe(true);
    expect(editor.getDocument()).toEqual(initial);
  });

  test('renders the native editor inside an isolated shadow root', () => {
    editor.destroy();
    editor = renderNativeWordEditor(host, viewFixture(), {
      styleIsolation: 'shadow',
      styleText: '.coco-word-native-editor-host { color: rgb(1, 2, 3); }',
    });

    expect(host.shadowRoot).not.toBeNull();
    expect(host.shadowRoot?.querySelector('style')?.textContent).toContain('rgb(1, 2, 3)');
    expect(host.shadowRoot?.querySelector('coco-word-native-p')).not.toBeNull();
    expect(host.querySelector('coco-word-native-p')).toBeNull();

    editor.destroy();
    expect(host.shadowRoot?.childElementCount).toBe(0);
  });

  test('adds the annotation rail outside the flow document width', () => {
    editor.destroy();
    editor = renderNativeWordEditor(host, revisionViewFixture(), { styleIsolation: 'none' });

    const documentNode = host.querySelector<HTMLElement>('coco-word-native-document');
    expect(documentNode?.classList.contains('has-annotation-rail')).toBe(true);
    expect(documentNode?.style.width).toBe('calc(100% + 260px)');
    expect(documentNode?.style.minWidth).toBe('calc(100% + 260px)');
  });
});
