import { describe, expect, test } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import {
  createWordEditingPlugin,
  getWordChangeSet,
  setParagraphAlign,
  toggleBold,
} from '../editing.js';
import { createWordEditorState, wordEditorSchema } from '../prosemirror.js';
import { EditorState } from 'prosemirror-state';

function createState(text = 'ABC'): EditorState {
  const doc = wordEditorSchema.nodes.doc.create(null, [
    wordEditorSchema.nodes.paragraph.create(
      {
        paraId: 'P0000001',
        story: 'body',
        styleId: null,
        format: null,
      },
      wordEditorSchema.text(text),
    ),
  ]);

  return EditorState.create({
    schema: wordEditorSchema,
    doc,
    plugins: [createWordEditingPlugin()],
  });
}

function createStateWithImage(): EditorState {
  const doc = wordEditorSchema.nodes.doc.create(null, [
    wordEditorSchema.nodes.paragraph.create(
      {
        paraId: 'P0000001',
        story: 'body',
        styleId: null,
        format: null,
      },
      [
        wordEditorSchema.text('A'),
        wordEditorSchema.nodes.image.create({
          src: 'data:image/png;base64,abc',
          alt: 'diagram',
          width: 120,
          height: 80,
          drawingId: 'rIdImage1',
        }),
        wordEditorSchema.text('B'),
      ],
    ),
  ]);

  return EditorState.create({
    schema: wordEditorSchema,
    doc,
    plugins: [createWordEditingPlugin()],
  });
}

function createStateWithParagraphs(...paragraphs: Array<{ paraId: string; content: string }>): EditorState {
  const doc = wordEditorSchema.nodes.doc.create(
    null,
    paragraphs.map((paragraph) =>
      wordEditorSchema.nodes.paragraph.create(
        {
          paraId: paragraph.paraId,
          story: 'body',
          styleId: null,
          format: null,
        },
        paragraph.content.length > 0 ? wordEditorSchema.text(paragraph.content) : undefined,
      ),
    ),
  );

  return EditorState.create({
    schema: wordEditorSchema,
    doc,
    plugins: [createWordEditingPlugin()],
  });
}

function paragraph(paraId: string, content: string) {
  return wordEditorSchema.nodes.paragraph.create(
    {
      paraId,
      story: 'body',
      styleId: null,
      format: null,
    },
    content.length > 0 ? wordEditorSchema.text(content) : undefined,
  );
}

describe('@coco-vexel/coco-office-editor word editing commands', () => {
  test('toggles bold and records paragraph-relative selection ranges', () => {
    let state = createState();
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 2)));

    const applied = toggleBold(state, (tr) => {
      state = state.apply(tr);
    });

    expect(applied).toBe(true);
    expect(state.doc.child(0).child(0).marks.map((mark) => mark.type.name)).toEqual(['strong']);
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'mark',
        from: 1,
        to: 2,
        mark: 'strong',
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        value: true,
        paragraphs: [{ paraId: 'P0000001', from: 0, to: 1, text: 'ABC' }],
      },
    ]);
  });

  test('updates paragraph alignment without dropping existing paragraph metadata', () => {
    let state = createState();
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 1)));

    const applied = setParagraphAlign('center')(state, (tr) => {
      state = state.apply(tr);
    });

    expect(applied).toBe(true);
    expect(state.doc.child(0).attrs).toEqual({
      paraId: 'P0000001',
      story: 'body',
      styleId: null,
      format: { alignment: 'center' },
    });
    expect(getWordChangeSet(state).changes[0]).toMatchObject({
      kind: 'paragraph',
      attr: 'alignment',
      status: 'captured',
      target: { story: 'body', paraId: 'P0000001' },
      value: 'center',
      paragraphs: [{ paraId: 'P0000001' }],
    });
  });

  test('records typed text as a paragraph-relative insert change', () => {
    let state = createState('AC');
    state = state.apply(state.tr.insertText('B', 2));

    expect(state.doc.child(0).textContent).toBe('ABC');
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'text',
        textOp: 'insert',
        from: 2,
        to: 3,
        text: 'B',
        deletedText: '',
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        paragraphs: [{ paraId: 'P0000001', from: 1, to: 2, text: 'ABC' }],
      },
    ]);
  });

  test('records deleted text as a paragraph-relative delete change', () => {
    let state = createState('ABC');
    state = state.apply(state.tr.delete(2, 3));

    expect(state.doc.child(0).textContent).toBe('AC');
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'text',
        textOp: 'delete',
        from: 2,
        to: 2,
        text: '',
        deletedText: 'B',
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        paragraphs: [{ paraId: 'P0000001', from: 1, to: 1, text: 'AC' }],
      },
    ]);
  });

  test('records replacement text as a paragraph-relative replace change', () => {
    let state = createState('ABC');
    state = state.apply(state.tr.insertText('X', 2, 3));

    expect(state.doc.child(0).textContent).toBe('AXC');
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'text',
        textOp: 'replace',
        from: 2,
        to: 3,
        text: 'X',
        deletedText: 'B',
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        paragraphs: [{ paraId: 'P0000001', from: 1, to: 2, text: 'AXC' }],
      },
    ]);
  });

  test('records inline image deletion as content deletion instead of a text edit', () => {
    let state = createStateWithImage();
    state = state.apply(state.tr.delete(2, 3));

    expect(state.doc.child(0).textContent).toBe('AB');
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'content',
        contentOp: 'delete',
        from: 2,
        to: 2,
        content: [],
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        deletedContent: [
          {
            kind: 'image',
            drawingId: 'rIdImage1',
            alt: 'diagram',
            src: 'data:image/png;base64,abc',
            width: 120,
            height: 80,
          },
        ],
        paragraphs: [{ paraId: 'P0000001', from: 1, to: 1, text: 'AB' }],
      },
    ]);
  });

  test('records mixed text and image deletion as structured content deletion', () => {
    let state = createStateWithImage();
    state = state.apply(state.tr.delete(1, 3));

    expect(state.doc.child(0).textContent).toBe('B');
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'content',
        contentOp: 'delete',
        from: 1,
        to: 1,
        content: [],
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        deletedContent: [
          { kind: 'text', text: 'A' },
          {
            kind: 'image',
            drawingId: 'rIdImage1',
            alt: 'diagram',
            src: 'data:image/png;base64,abc',
            width: 120,
            height: 80,
          },
        ],
        paragraphs: [{ paraId: 'P0000001', from: 0, to: 0, text: 'B' }],
      },
    ]);
  });

  test('records inserted paragraph as a block insert change', () => {
    let state = createStateWithParagraphs({ paraId: 'P0000001', content: 'A' });
    state = state.apply(state.tr.insert(3, paragraph('P0000002', 'B')));

    expect(state.doc.childCount).toBe(2);
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'block',
        blockOp: 'insert',
        status: 'captured',
        target: { story: 'body', paraId: 'P0000002' },
        blocks: [{ kind: 'paragraph', paraId: 'P0000002', text: 'B' }],
        deletedBlocks: [],
      },
    ]);
  });

  test('records deleted paragraph as a block delete change', () => {
    let state = createStateWithParagraphs(
      { paraId: 'P0000001', content: 'A' },
      { paraId: 'P0000002', content: 'B' },
    );
    state = state.apply(state.tr.delete(0, state.doc.child(0).nodeSize));

    expect(state.doc.childCount).toBe(1);
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'block',
        blockOp: 'delete',
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        blocks: [],
        deletedBlocks: [{ kind: 'paragraph', paraId: 'P0000001', text: 'A' }],
      },
    ]);
  });

  test('records paragraph split as a block split change', () => {
    let state = createStateWithParagraphs({ paraId: 'P0000001', content: 'ABC' });
    state = state.apply(
      state.tr.split(2, 1, [
        {
          type: wordEditorSchema.nodes.paragraph,
          attrs: {
            paraId: 'P0000002',
            story: 'body',
            styleId: null,
            format: null,
          },
        },
      ]),
    );

    expect([...Array(state.doc.childCount)].map((_, index) => state.doc.child(index).textContent)).toEqual(['A', 'BC']);
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'block',
        blockOp: 'split',
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        blocks: [
          { kind: 'paragraph', paraId: 'P0000001', text: 'A' },
          { kind: 'paragraph', paraId: 'P0000002', text: 'BC' },
        ],
        deletedBlocks: [{ kind: 'paragraph', paraId: 'P0000001', text: 'ABC' }],
      },
    ]);
  });

  test('records paragraph merge as a block merge change', () => {
    let state = createStateWithParagraphs(
      { paraId: 'P0000001', content: 'A' },
      { paraId: 'P0000002', content: 'BC' },
    );
    state = state.apply(state.tr.delete(2, 4));

    expect(state.doc.childCount).toBe(1);
    expect(state.doc.child(0).textContent).toBe('ABC');
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'block',
        blockOp: 'merge',
        status: 'captured',
        target: { story: 'body', paraId: 'P0000001' },
        blocks: [{ kind: 'paragraph', paraId: 'P0000001', text: 'ABC' }],
        deletedBlocks: [
          { kind: 'paragraph', paraId: 'P0000001', text: 'A' },
          { kind: 'paragraph', paraId: 'P0000002', text: 'BC' },
        ],
      },
    ]);
  });

  test('records deleted image paragraph without dropping image content', () => {
    const doc = wordEditorSchema.nodes.doc.create(null, [
      createStateWithImage().doc.child(0),
      paragraph('P0000002', 'after'),
    ]);
    let state = EditorState.create({
      schema: wordEditorSchema,
      doc,
      plugins: [createWordEditingPlugin()],
    });
    state = state.apply(state.tr.delete(0, state.doc.child(0).nodeSize));

    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'block',
        blockOp: 'delete',
        deletedBlocks: [
          {
            kind: 'paragraph',
            paraId: 'P0000001',
            content: [
              { kind: 'text', text: 'A' },
              { kind: 'image', drawingId: 'rIdImage1' },
              { kind: 'text', text: 'B' },
            ],
          },
        ],
      },
    ]);
  });

  test('normalizes duplicate paragraph ids after a user split without recording extra changes', () => {
    const editorView = {
      blocks: [
        {
          type: 'paragraph',
          paraId: 'P0000001',
          story: 'body',
          text: 'ABC',
          segments: [{ kind: 'text', from: 0, to: 3, text: 'ABC' }],
          spans: [],
        },
      ],
      paragraphs: [
        {
          paraId: 'P0000001',
          story: 'body',
          text: 'ABC',
          segments: [{ kind: 'text', from: 0, to: 3, text: 'ABC' }],
          spans: [],
        },
      ],
      tables: [],
      comments: [],
    } as Parameters<typeof createWordEditorState>[0];
    let state = createWordEditorState(editorView, { plugins: [createWordEditingPlugin()] });

    state = state.applyTransaction(state.tr.split(2)).state;

    const firstParaId = state.doc.child(0).attrs.paraId;
    const secondParaId = state.doc.child(1).attrs.paraId;
    expect(firstParaId).toBe('P0000001');
    expect(secondParaId).toMatch(/^[0-9A-F]{8}$/);
    expect(secondParaId).not.toBe(firstParaId);
    expect(getWordChangeSet(state).changes).toMatchObject([
      {
        kind: 'block',
        blockOp: 'split',
      },
    ]);
    expect(getWordChangeSet(state).changes).toHaveLength(1);
  });
});
