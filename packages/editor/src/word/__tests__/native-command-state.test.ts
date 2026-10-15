import { describe, expect, test } from 'vitest';
import { canDispatchNativeWordCommand, getNativeWordCommandState } from '../native-command-state.js';
import type { NativeWordDocument } from '../native-model.js';

function documentFixture(): NativeWordDocument {
  return {
    kind: 'word',
    defaults: { run: { fontFamily: 'Aptos', fontSize: 11, color: '000000' } },
    sections: [],
    headers: [],
    footers: [],
    blocks: [
      {
        kind: 'paragraph',
        paraId: 'p1',
        story: 'body',
        styleId: null,
        format: {
          alignment: 'left',
          indent: { left: 720, firstLine: 420 },
          spacing: { before: 120, after: 240, line: 360, lineRule: 'auto' },
        },
        content: [
          {
            kind: 'text',
            text: 'Bold',
            format: { bold: true, italic: true, underline: true, fontFamily: 'Arial', color: 'C00000', highlight: 'FFFF00' },
          },
          { kind: 'text', text: 'Plain', format: { bold: false, strike: true } },
        ],
      },
      {
        kind: 'paragraph',
        paraId: 'p2',
        story: 'body',
        styleId: null,
        format: {
          alignment: 'center',
          indent: { right: 240 },
          spacing: { after: 120, line: 480, lineRule: 'auto' },
        },
        content: [{ kind: 'text', text: 'Large', format: { fontSize: 18 } }],
      },
    ],
  };
}

describe('native command state', () => {
  test('resolves the run format at a caret', () => {
    const state = getNativeWordCommandState(documentFixture(), { kind: 'caret', paraId: 'p1', offset: 2 });
    expect(state).toMatchObject({
      selectionKind: 'caret',
      bold: true,
      italic: true,
      underline: true,
      strike: false,
      fontFamily: 'Arial',
      fontSize: 11,
      textColor: '#C00000',
      highlightColor: '#FFFF00',
      paragraphAlignment: 'left',
      leftIndent: 720,
      rightIndent: 0,
      firstLineIndent: 420,
      spacingBefore: 120,
      spacingAfter: 240,
      lineSpacing: { line: 360, lineRule: 'auto' },
      listType: null,
      listLevel: null,
    });
  });

  test('reports mixed run values inside one paragraph', () => {
    const selection = {
      kind: 'text-range' as const,
      anchor: { paraId: 'p1', offset: 0 },
      focus: { paraId: 'p1', offset: 9 },
    };
    expect(getNativeWordCommandState(documentFixture(), selection)).toMatchObject({
      bold: 'mixed',
      italic: 'mixed',
      underline: 'mixed',
      strike: 'mixed',
      fontFamily: 'mixed',
      fontSize: 11,
      textColor: 'mixed',
      highlightColor: 'mixed',
      paragraphAlignment: 'left',
    });
    expect(canDispatchNativeWordCommand(documentFixture(), selection, { kind: 'toggle-bold' })).toBe(true);
    expect(canDispatchNativeWordCommand(documentFixture(), selection, { kind: 'toggle-italic' })).toBe(true);
    expect(canDispatchNativeWordCommand(documentFixture(), selection, { kind: 'clear-text-format' })).toBe(true);
  });

  test('reports hyperlink state and gates set/remove commands to text selections', () => {
    const doc = documentFixture();
    const paragraph = doc.blocks[0];
    if (paragraph?.kind !== 'paragraph' || paragraph.content[0]?.kind !== 'text') {
      throw new Error('paragraph fixture missing');
    }
    paragraph.content[0].hyperlink = { target: 'https://coco.example/docs', tooltip: 'Docs' };
    const linked = {
      kind: 'text-range' as const,
      anchor: { paraId: 'p1', offset: 0 },
      focus: { paraId: 'p1', offset: 4 },
    };

    expect(getNativeWordCommandState(doc, linked).hyperlink).toEqual({
      target: 'https://coco.example/docs',
      tooltip: 'Docs',
    });
    expect(canDispatchNativeWordCommand(doc, linked, { kind: 'remove-hyperlink' })).toBe(true);
    expect(canDispatchNativeWordCommand(doc, linked, {
      kind: 'set-hyperlink',
      hyperlink: { anchor: 'Section_1' },
    })).toBe(true);
    expect(canDispatchNativeWordCommand(doc, { kind: 'caret', paraId: 'p1', offset: 2 }, {
      kind: 'set-hyperlink',
      hyperlink: { target: 'https://coco.example' },
    })).toBe(false);
  });

  test('reports mixed paragraph values across paragraphs', () => {
    const selection = {
      kind: 'text-range' as const,
      anchor: { paraId: 'p1', offset: 2 },
      focus: { paraId: 'p2', offset: 2 },
    };
    expect(getNativeWordCommandState(documentFixture(), selection)).toMatchObject({
      bold: 'mixed',
      fontSize: 'mixed',
      textColor: 'mixed',
      paragraphAlignment: 'mixed',
      leftIndent: 'mixed',
      rightIndent: 'mixed',
      firstLineIndent: 'mixed',
      spacingBefore: 'mixed',
      spacingAfter: 'mixed',
      lineSpacing: 'mixed',
    });
  });

  test('disables text formatting without a text range', () => {
    const selection = { kind: 'caret' as const, paraId: 'p1', offset: 2 };
    expect(canDispatchNativeWordCommand(documentFixture(), selection, { kind: 'toggle-bold' })).toBe(false);
    expect(
      canDispatchNativeWordCommand(documentFixture(), selection, {
        kind: 'set-paragraph-alignment',
        alignment: 'center',
      }),
    ).toBe(true);
    expect(canDispatchNativeWordCommand(documentFixture(), selection, {
      kind: 'insert-table',
      rows: 2,
      columns: 3,
    })).toBe(true);
    expect(canDispatchNativeWordCommand(documentFixture(), selection, {
      kind: 'insert-table',
      rows: 0,
      columns: 3,
    })).toBe(false);
  });

  test('allows page-break insertion only at a caret', () => {
    const doc = documentFixture();
    expect(canDispatchNativeWordCommand(doc, { kind: 'caret', paraId: 'p1', offset: 2 }, {
      kind: 'insert-page-break',
    })).toBe(true);
    expect(canDispatchNativeWordCommand(doc, {
      kind: 'text-range',
      anchor: { paraId: 'p1', offset: 1 },
      focus: { paraId: 'p1', offset: 2 },
    }, { kind: 'insert-page-break' })).toBe(false);
  });

  test('reports list type and mixed list levels', () => {
    const doc = documentFixture();
    const first = doc.blocks[0];
    const second = doc.blocks[1];
    if (!first || first.kind !== 'paragraph' || !second || second.kind !== 'paragraph') {
      throw new Error('paragraph fixtures missing');
    }
    first.format = {
      ...(first.format ?? {}),
      hasNumbering: true,
      numbering: { numId: 'list-1', ilvl: 0, numFmt: 'bullet', lvlText: '•', label: '•' },
    };
    second.format = {
      ...(second.format ?? {}),
      hasNumbering: true,
      numbering: { numId: 'list-1', ilvl: 1, numFmt: 'bullet', lvlText: '◦', label: '◦' },
    };
    const selection = {
      kind: 'text-range' as const,
      anchor: { paraId: 'p1', offset: 0 },
      focus: { paraId: 'p2', offset: 2 },
    };

    expect(getNativeWordCommandState(doc, selection)).toMatchObject({
      listType: 'bullet',
      listLevel: 'mixed',
    });
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'adjust-list-level', delta: 1 })).toBe(true);
  });

  test('exposes table position and gates structural commands for merged cells', () => {
    const doc = documentFixture();
    const cellParagraph = {
      kind: 'paragraph' as const,
      paraId: 'table-p1',
      story: 'body',
      styleId: null,
      format: null,
      content: [{ kind: 'text' as const, text: 'Cell', format: null }],
    };
    doc.blocks.push({
      kind: 'table',
      tableId: 'table-1',
      columnWidths: [1200],
      rows: [{ cells: [{ gridSpan: 1, vMerge: null, blocks: [cellParagraph], paragraphs: [cellParagraph] }] }],
    });
    const selection = { kind: 'caret' as const, paraId: 'table-p1', offset: 2 };

    expect(getNativeWordCommandState(doc, selection).table).toEqual({
      tableId: 'table-1',
      rowIndex: 0,
      rowCount: 1,
      columnIndex: 0,
      columnCount: 1,
      range: null,
      structuralEditingSupported: true,
      backgroundColor: null,
      verticalAlignment: null,
      borderColor: null,
      bordersVisible: false,
    });
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'insert-table-row', position: 'after' })).toBe(true);
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'delete-table-row' })).toBe(false);
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'delete-table' })).toBe(true);
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'insert-table', rows: 2, columns: 2 })).toBe(false);

    const table = doc.blocks[2];
    if (table?.kind !== 'table') throw new Error('Expected table fixture');
    table.rows[0]!.cells[0]!.gridSpan = 2;
    expect(getNativeWordCommandState(doc, selection).table?.structuralEditingSupported).toBe(false);
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'insert-table-column', position: 'after' })).toBe(false);
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'set-table-cell-background', color: '#FF0000' })).toBe(true);
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'clear-table-cells' })).toBe(false);
  });

  test('summarizes mixed formatting across a table cell range', () => {
    const doc = documentFixture();
    const paragraph = (paraId: string, text: string) => ({
      kind: 'paragraph' as const,
      paraId,
      story: 'body',
      styleId: null,
      format: null,
      content: [{ kind: 'text' as const, text, format: null }],
    });
    const cell = (paraId: string, text: string, fill?: string) => {
      const block = paragraph(paraId, text);
      return {
        gridSpan: 1,
        vMerge: null,
        ...(fill ? { shading: { fill, val: 'clear' } } : {}),
        blocks: [block],
        paragraphs: [block],
      };
    };
    doc.blocks.push({
      kind: 'table',
      tableId: 'table-range',
      columnWidths: [1200, 1200],
      rows: [
        { cells: [cell('range-p1', 'A', 'FF0000'), cell('range-p2', 'B')] },
        { cells: [cell('range-p3', 'C', 'FF0000'), cell('range-p4', 'D')] },
      ],
    });
    const selection = {
      kind: 'table-cell-range' as const,
      tableId: 'table-range',
      anchor: { rowIndex: 1, columnIndex: 1 },
      focus: { rowIndex: 0, columnIndex: 0 },
    };

    expect(getNativeWordCommandState(doc, selection).table).toMatchObject({
      tableId: 'table-range',
      rowIndex: 1,
      columnIndex: 1,
      backgroundColor: 'mixed',
      range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1, cellCount: 4 },
    });
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'clear-table-cells' })).toBe(true);
    expect(canDispatchNativeWordCommand(doc, selection, { kind: 'set-table-cell-borders', color: '#000000' })).toBe(true);
  });
});
