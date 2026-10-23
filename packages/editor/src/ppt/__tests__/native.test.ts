/**
 * @vitest-environment happy-dom
 */
import { describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { Presentation } from '@coco-vexel/coco-office/ppt';
import { renderNativePptEditor } from '../native.js';
import { buildPptPresentationScene, type PptSceneNode, type PptSlideScene } from '../scene.js';
import { buildPptEditorViewFromModel, buildPptEditorViewFromPptx } from '../view.js';

async function buildEditorHost() {
  const host = document.createElement('div');
  document.body.append(host);
  const view = await buildPptEditorViewFromModel({
    id: 'native-test',
    slideSize: { width: 9144000, height: 5143500 },
    slides: [
      {
        id: '256',
        index: 0,
        shapes: [
          {
            id: '2',
            name: 'Title',
            kind: 'textBox',
            geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
            text: { text: 'Draft' },
          },
        ],
      },
    ],
  });
  const editor = renderNativePptEditor(host, view);
  return { editor, root: host.shadowRoot! };
}

function mouse(
  type: string,
  clientX: number,
  clientY: number,
  options: MouseEventInit = {},
): MouseEvent {
  return new MouseEvent(type, { bubbles: true, clientX, clientY, button: 0, ...options });
}

function contextMenu(clientX: number, clientY: number): MouseEvent {
  return new MouseEvent('contextmenu', {
    bubbles: true,
    cancelable: true,
    clientX,
    clientY,
    button: 2,
  });
}

function menuButton(root: ShadowRoot, label: string): HTMLButtonElement {
  const buttons = [...root.querySelectorAll<HTMLButtonElement>('.coco-ppt-context-menu button')];
  const button = buttons.find((item) => item.textContent?.includes(label));
  if (!button) throw new Error(`Missing context menu button: ${label}`);
  return button;
}

async function waitFor(assertion: () => void): Promise<void> {
  let lastError: unknown;
  for (let i = 0; i < 20; i++) {
    try {
      assertion();
      return;
    } catch (error) {
      lastError = error;
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  throw lastError;
}

function waitForNextChange(editor: ReturnType<typeof renderNativePptEditor>): Promise<void> {
  return new Promise((resolve) => {
    const unsubscribe = editor.session().subscribe('change', () => {
      unsubscribe();
      resolve();
    });
  });
}

function testThumbnailFingerprint(slide: PptSlideScene): string {
  return JSON.stringify({
    size: slide.size,
    background: slide.background,
    nodes: slide.nodes.map((node) => testThumbnailNodeFingerprint(node)),
  });
}

function testThumbnailNodeFingerprint(node: PptSceneNode): unknown {
  return {
    id: node.id,
    kind: node.kind,
    bounds: node.bounds,
    text: node.text,
    textParagraphs: node.textParagraphs,
    textStyle: node.textStyle,
    textBody: node.textBody,
    table: node.table,
    fillColor: node.fillColor,
    fillGradient: node.fillGradient,
    opacity: node.opacity,
    lineColor: node.lineColor,
    lineGradient: node.lineGradient,
    lineOpacity: node.lineOpacity,
    lineWidth: node.lineWidth,
    lineDash: node.lineDash,
    presetGeometry: node.presetGeometry,
    presetAdjustments: node.presetAdjustments,
    customGeometry: node.customGeometry,
    imagePath: node.imagePath,
    imageCrop: node.imageCrop,
    imageFillRect: node.imageFillRect,
    imageOpacity: node.imageOpacity,
    shadow: node.shadow,
    threeD: node.threeD,
    children: node.children.map((child) => testThumbnailNodeFingerprint(child)),
  };
}

describe('renderNativePptEditor', () => {
  test('uses an Office-like default font stack for PPT text', async () => {
    const { editor, root } = await buildEditorHost();
    const css = root.querySelector('style')?.textContent ?? '';

    expect(css).toContain(
      'font-family: "Calibri", "Microsoft YaHei", SimSun, "Segoe UI", Arial, sans-serif',
    );

    await editor.destroy();
  });

  test('loads known web fonts referenced by slide text even when they are not embedded', async () => {
    document.getElementById('coco-ppt-google-font-Cuprum')?.remove();
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-referenced-font-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Web font title', style: { fontFamily: 'Cuprum' } },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textStyle = { fontFamily: 'Cuprum' };
    const editor = renderNativePptEditor(host, view);

    const cuprum = document.getElementById('coco-ppt-google-font-Cuprum') as HTMLLinkElement | null;

    expect(cuprum?.getAttribute('href')).toContain('fonts.googleapis.com');
    expect(cuprum?.getAttribute('href')).toContain('Cuprum');

    await editor.destroy();
    cuprum?.remove();
  });

  test('maps PPT bold text to the standard Office CSS weight', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-bold-weight-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Bold text' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textParagraphs = [
      {
        text: 'Bold text',
        runs: [
          { text: 'Bold', style: { bold: true } },
          { text: ' text', style: { italic: true } },
        ],
      },
    ];
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-rich-text span')!;

    expect(node.style.fontWeight).toBe('700');

    await editor.destroy();
  });

  test('projects tables using grid extents when frame extents are stale', () => {
    const scene = buildPptPresentationScene({
      id: 'native-table-grid-extent-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '15',
              name: 'Grid sized table',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 1000000, height: 1000000 },
              table: {
                values: [
                  ['A', 'B', 'C'],
                  ['D', 'E', 'F'],
                ],
                rowCount: 2,
                columnCount: 3,
                columnWidths: [1200000, 1600000, 1400000],
                rowHeights: [700000, 900000],
              },
            },
          ],
        },
      ],
    });

    expect(scene.slides[0]!.nodes[0]!.bounds).toMatchObject({
      width: 1200000 + 1600000 + 1400000,
      height: 700000 + 900000,
    });
  });

  test('keeps table frame extents when grid extents are smaller', () => {
    const scene = buildPptPresentationScene({
      id: 'native-table-grid-extent-shrink-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '15',
              name: 'Frame sized table',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 5000000, height: 2400000 },
              table: {
                values: [
                  ['A', 'B'],
                  ['C', 'D'],
                ],
                rowCount: 2,
                columnCount: 2,
                columnWidths: [1200000, 1200000],
                rowHeights: [500000, 500000],
              },
            },
          ],
        },
      ],
    });

    expect(scene.slides[0]!.nodes[0]!.bounds).toMatchObject({
      width: 5000000,
      height: 2400000,
    });
  });

  test('keeps multiline table cell text editable while rendering secondary lines smaller', async () => {
    let selectedText = '';
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-table-multiline-cell-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '15',
              name: 'Multiline table',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 2200000, height: 900000 },
              table: {
                values: [['Add a main point\nElaborate on what you want to discuss.']],
                rowCount: 1,
                columnCount: 1,
              },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view, {
      onTableCellSelectionChange: (event) => {
        selectedText = event.text;
      },
    });
    const cell = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-table td',
    )!;

    expect(cell.querySelector('.coco-ppt-table-cell-secondary-line')?.textContent).toBe(
      'Elaborate on what you want to discuss.',
    );
    cell.dispatchEvent(new MouseEvent('click', { bubbles: true, composed: true }));
    cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true }));

    expect(selectedText).toBe('Add a main point\nElaborate on what you want to discuss.');
    expect(cell.textContent).toBe('Add a main point\nElaborate on what you want to discuss.');

    await editor.destroy();
  });

  test('renders explicit table cell borders from the PPT model', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-table-cell-border-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '8',
              name: 'Border Table',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 800000 },
              table: {
                values: [['A', 'B']],
                rowCount: 1,
                columnCount: 2,
                cellStyles: [
                  [
                    {
                      borders: {
                        left: { hidden: true },
                        top: { color: '#FF0000', opacity: 0.5, width: 25400, dash: 'dash' },
                      },
                    },
                    {},
                  ],
                ],
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.table!.cellStyles = [
      [
        {
          borders: {
            left: { hidden: true },
            top: { color: '#FF0000', opacity: 0.5, width: 25400, dash: 'dash' },
          },
        },
        {},
      ],
    ];
    const editor = renderNativePptEditor(host, view);
    const cell = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-table td[data-row="0"][data-column="0"]',
    )!;

    expect(cell.style.borderLeftStyle).toBe('none');
    expect(cell.style.borderTopStyle).toBe('dashed');
    expect(cell.style.borderTopColor).toBe('rgba(255, 0, 0, 0.5)');
    expect(Number.parseFloat(cell.style.borderTopWidth)).toBeGreaterThan(0);

    await editor.destroy();
  });

  test('edits shape text from the slide canvas', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    node.textContent = 'Final';
    node.dispatchEvent(new FocusEvent('blur'));

    await waitFor(() => expect(editor.session().scene().slides[0]!.nodes[0]!.text).toBe('Final'));
    const verified = await editor.session().verifyExport();
    expect(verified.slides[0]!.shapes[0]!.text?.text).toBe('Final');

    await editor.destroy();
  });

  test.each([false, true])(
    'exports pending text edits immediately (already blurred: %s)',
    async (alreadyBlurred) => {
      const { editor, root } = await buildEditorHost();
      try {
        const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;
        node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
        node.textContent = 'Exported immediately';
        if (alreadyBlurred) node.dispatchEvent(new FocusEvent('blur'));
        const reloaded = await Presentation.fromZip(await editor.toPptx());
        try {
          expect(await reloaded.slides.getItem(0).shapes.getItem('Title').text.getText()).toBe(
            'Exported immediately',
          );
        } finally {
          await reloaded.dispose();
        }
      } finally {
        await editor.destroy();
      }
    },
  );

  test('commits text edits with Enter and keeps the shape selected', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    await editor.dispatch({ kind: 'select-shape', slideIdOrIndex: 0, shapeIdOrName: 'Title' });
    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    node.textContent = 'Committed';
    node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));

    await waitFor(() =>
      expect(editor.session().scene().slides[0]!.nodes[0]!.text).toBe('Committed'),
    );
    expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: '2' });

    await editor.destroy();
  });

  test('cancels text edits with Escape without adding history', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    node.textContent = 'Cancelled';
    node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Escape' }));

    await waitFor(() => expect(node.contentEditable).not.toBe('true'));
    expect(editor.session().scene().slides[0]!.nodes[0]!.text).toBe('Draft');
    expect(editor.canUndo()).toBe(false);

    await editor.destroy();
  });

  test('emits text edit state changes for start, commit, and cancel', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-text-edit-state-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Draft' },
            },
          ],
        },
      ],
    });
    const events: string[] = [];
    const editor = renderNativePptEditor(host, view, {
      onTextEditStateChange: (event) => events.push(`${event.state}:${event.text}`),
    });
    const node = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-node')!;

    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    node.textContent = 'Committed';
    node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await waitFor(() =>
      expect(editor.session().scene().slides[0]!.nodes[0]!.text).toBe('Committed'),
    );
    const renderedNode = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-node')!;
    renderedNode.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    renderedNode.textContent = 'Cancelled';
    renderedNode.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Escape' }));

    expect(events).toEqual([
      'start:Draft',
      'commit:Committed',
      'start:Committed',
      'cancel:Committed',
    ]);

    await editor.destroy();
  });

  test('keeps text editing active on Shift+Enter', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    node.dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', shiftKey: true }),
    );

    expect(node.contentEditable).toBe('true');
    expect(editor.session().scene().slides[0]!.nodes[0]!.text).toBe('Draft');

    await editor.destroy();
  });

  test('marks active direct text editing and keeps pasted content plain', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(node.dataset.editing).toBe('true');
    const paste = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(paste, 'clipboardData', {
      value: { getData: (type: string) => (type === 'text/plain' ? '<b>Plain</b>' : '') },
    });
    node.textContent = '';
    node.dispatchEvent(paste);
    node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));

    await waitFor(() =>
      expect(editor.session().scene().slides[0]!.nodes[0]!.text).toBe('<b>Plain</b>'),
    );
    expect(root.querySelector('.coco-ppt-node')?.getAttribute('data-editing')).toBeNull();

    await editor.destroy();
  });

  test('starts text editing after a selected shape receives a no-op double-click drag sequence', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    node.dispatchEvent(mouse('click', 10, 10));
    await waitFor(() => expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: '2' }));
    node.dispatchEvent(
      new MouseEvent('mousedown', {
        bubbles: true,
        cancelable: true,
        clientX: 10,
        clientY: 10,
        button: 0,
      }),
    );
    window.dispatchEvent(mouse('mouseup', 10, 10));
    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, clientX: 10, clientY: 10 }));

    expect(node.contentEditable).toBe('true');

    await editor.destroy();
  });

  test('reports blocked text editing for non-text editable objects', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-blocked-text-edit-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '7',
              name: 'Picture',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1000000, height: 600000 },
            },
            {
              id: '8',
              name: 'Table',
              kind: 'textBox',
              geometry: { left: 1200000, top: 100000, width: 1200000, height: 800000 },
            },
          ],
        },
      ],
    });
    Object.assign(view.scene.slides[0]!.nodes[0]!, {
      kind: 'picture',
      capabilities: {
        select: true,
        editText: false,
        editTextStyle: false,
        editTableCells: false,
        editGeometry: true,
        editShapeStyle: true,
      },
    });
    Object.assign(view.scene.slides[0]!.nodes[1]!, {
      kind: 'table',
      table: {
        rowCount: 1,
        columnCount: 1,
        values: [['Cell']],
        rows: [{ height: 800000 }],
        columns: [{ width: 1200000 }],
      },
      capabilities: {
        select: true,
        editText: false,
        editTextStyle: false,
        editTableCells: true,
        editGeometry: true,
        editShapeStyle: true,
      },
    });
    const events: string[] = [];
    const editor = renderNativePptEditor(host, view, {
      onObjectAction: (event) => events.push(`${event.shapeId}:${event.reason}`),
    });
    const root = host.shadowRoot!;
    const picture = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="7"]')!;
    const table = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="8"]')!;

    picture.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    await waitFor(() => expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: '7' }));
    expect(root.querySelector('[contenteditable="true"]')).toBeNull();
    table.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    await waitFor(() => expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: '8' }));

    expect(root.querySelector('[contenteditable="true"]')).toBeNull();
    expect(events).toEqual(['7:图片对象不支持文字编辑', '8:表格请使用单元格编辑']);

    await editor.destroy();
  });

  test('renders table cells with stable row and column coordinates', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-table-cell-coordinate-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 600000 },
              table: {
                values: [
                  ['A1', 'B1'],
                  ['A2', 'B2'],
                ],
                rowCount: 2,
                columnCount: 2,
              },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const cell = host.shadowRoot!.querySelector<HTMLTableCellElement>(
      '.coco-ppt-table td[data-row="1"][data-column="1"]',
    )!;

    expect(cell).not.toBeNull();
    expect(cell.textContent).toBe('B2');

    await editor.destroy();
  });

  test('keeps table grid flush and does not inherit shape alignment into cells', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-table-cell-alignment-inheritance-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '4',
              name: 'Schedule',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 600000 },
              text: {
                text: 'Task NameCentered',
                style: { alignment: 'center' },
              },
              table: {
                values: [['Task Name', 'Centered']],
                rowCount: 1,
                columnCount: 2,
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = {
      insetLeft: 91440,
      insetTop: 45720,
      insetRight: 91440,
      insetBottom: 45720,
    };
    view.scene.slides[0]!.nodes[0]!.table!.cellStyles = [
      [{ textStyle: { fontFamily: 'SimSun' } }, { textStyle: { alignment: 'center' } }],
    ];
    const editor = renderNativePptEditor(host, view);
    try {
      const node = host.shadowRoot!.querySelector<HTMLElement>(
        '.coco-ppt-node[data-kind="table"]',
      )!;
      const leftCell = host.shadowRoot!.querySelector<HTMLTableCellElement>(
        '.coco-ppt-table td[data-row="0"][data-column="0"]',
      )!;
      const centerCell = host.shadowRoot!.querySelector<HTMLTableCellElement>(
        '.coco-ppt-table td[data-row="0"][data-column="1"]',
      )!;

      expect(node.style.padding).toBe('');
      expect(getComputedStyle(leftCell).textAlign).toBe('left');
      expect(getComputedStyle(leftCell).fontFamily).toContain('SimSun');
      expect(getComputedStyle(centerCell).textAlign).toBe('center');
    } finally {
      await editor.destroy();
    }
  });

  test('selects the table shape when clicking a rendered table cell', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-table-cell-click-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 600000 },
              table: {
                values: [['A1']],
                rowCount: 1,
                columnCount: 1,
              },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const cell = host.shadowRoot!.querySelector<HTMLTableCellElement>(
      '.coco-ppt-table td[data-row="0"][data-column="0"]',
    )!;

    cell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await waitFor(() =>
      expect(editor.selection()).toMatchObject({ kind: 'shape', slideId: '256', shapeId: '4' }),
    );

    await editor.destroy();
  });

  test('reports the clicked table cell coordinates and text', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-table-cell-selection-callback-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 600000 },
              table: {
                values: [
                  ['A1', 'B1'],
                  ['A2', 'B2'],
                ],
                rowCount: 2,
                columnCount: 2,
              },
            },
          ],
        },
      ],
    });
    const events: string[] = [];
    const editor = renderNativePptEditor(host, view, {
      onTableCellSelectionChange: (event) =>
        events.push(`${event.shapeId}:${event.row},${event.column}:${event.text}`),
    });
    const cell = host.shadowRoot!.querySelector<HTMLTableCellElement>(
      '.coco-ppt-table td[data-row="1"][data-column="1"]',
    )!;

    cell.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    await waitFor(() => expect(events).toEqual(['4:1,1:B2']));

    await editor.destroy();
  });

  test('edits table cell text directly from the slide canvas', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-table-cell-direct-edit-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 600000 },
              table: {
                values: [['A1']],
                rowCount: 1,
                columnCount: 1,
              },
            },
          ],
        },
      ],
    });
    const events: string[] = [];
    const editor = renderNativePptEditor(host, view, {
      onTableCellSelectionChange: (event) =>
        events.push(`${event.shapeId}:${event.row},${event.column}:${event.text}`),
    });
    const cell = host.shadowRoot!.querySelector<HTMLTableCellElement>(
      '.coco-ppt-table td[data-row="0"][data-column="0"]',
    )!;

    cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(cell.contentEditable).toBe('true');
    expect(events).toEqual(['4:0,0:A1']);
    cell.textContent = 'Canvas Cell';
    cell.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));

    await waitFor(() =>
      expect(editor.session().scene().slides[0]!.nodes[0]!.table?.values[0]?.[0]).toBe(
        'Canvas Cell',
      ),
    );
    const verified = await editor.session().verifyExport();
    expect(verified.slides[0]!.shapes[0]!.table?.values[0]?.[0]).toBe('Canvas Cell');

    await editor.destroy();
  });

  test('cancels direct table cell edits with Escape', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-table-cell-direct-edit-cancel-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 600000 },
              table: {
                values: [['A1']],
                rowCount: 1,
                columnCount: 1,
              },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const cell = host.shadowRoot!.querySelector<HTMLTableCellElement>(
      '.coco-ppt-table td[data-row="0"][data-column="0"]',
    )!;

    cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    cell.textContent = 'Cancelled Cell';
    cell.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Escape' }));

    await waitFor(() => expect(cell.contentEditable).not.toBe('true'));
    expect(editor.session().scene().slides[0]!.nodes[0]!.table?.values[0]?.[0]).toBe('A1');
    expect(editor.canUndo()).toBe(false);

    await editor.destroy();
  });

  test('redirects overlapped decorative hits to editable text underneath', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-overlap-hit-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: 'text',
              name: 'Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Editable underlay' },
            },
            {
              id: 'overlay',
              name: 'Decorative overlay',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
            },
          ],
        },
      ],
    });
    Object.assign(view.scene.slides[0]!.nodes[1]!, {
      kind: 'geometricShape',
      capabilities: {
        select: true,
        editText: false,
        editTextStyle: false,
        editTableCells: false,
        editGeometry: true,
        editShapeStyle: true,
      },
    });
    const editor = renderNativePptEditor(host, view);
    const root = host.shadowRoot!;
    const canvas = root.querySelector<HTMLElement>('.coco-ppt-canvas')!;
    const overlay = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="overlay"]')!;
    const text = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="text"]')!;
    const scale = Number.parseFloat(canvas.style.width) / 9144000;
    const clientX = 10;
    const clientY = 10;
    canvas.getBoundingClientRect = () =>
      ({
        left: clientX - (100000 + 1000000) * scale,
        top: clientY - (100000 + 250000) * scale,
        width: 1000,
        height: 600,
        right: 1000,
        bottom: 600,
        x: 0,
        y: 0,
        toJSON: () => ({}),
      }) as DOMRect;

    overlay.dispatchEvent(
      new MouseEvent('dblclick', {
        bubbles: true,
        clientX,
        clientY,
      }),
    );

    await waitFor(() => expect(text.contentEditable).toBe('true'));
    expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: 'text' });

    await editor.destroy();
  });

  test('uses the same overlapped target for context menu and drag operations', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-overlap-operation-target-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: 'text',
              name: 'Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Editable underlay' },
            },
            {
              id: 'overlay',
              name: 'Decorative overlay',
              kind: 'geometricShape',
              raw: { presetGeometry: 'rect' },
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
            },
          ],
        },
      ],
    });
    Object.assign(view.scene.slides[0]!.nodes[1]!, {
      capabilities: {
        select: true,
        editText: false,
        editTextStyle: false,
        editTableCells: false,
        editGeometry: true,
        editShapeStyle: true,
      },
    });
    const editor = renderNativePptEditor(host, view);
    const root = host.shadowRoot!;
    const canvas = root.querySelector<HTMLElement>('.coco-ppt-canvas')!;
    const overlay = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="overlay"]')!;
    const text = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="text"]')!;
    const scale = Number.parseFloat(canvas.style.width) / 9144000;
    const clientX = 10;
    const clientY = 10;
    canvas.getBoundingClientRect = () =>
      ({
        left: clientX - (100000 + 1000000) * scale,
        top: clientY - (100000 + 250000) * scale,
        width: 1000,
        height: 600,
        right: 1000,
        bottom: 600,
        x: 0,
        y: 0,
        toJSON: () => ({}),
      }) as DOMRect;

    overlay.dispatchEvent(contextMenu(clientX, clientY));
    await waitFor(() => {
      expect(root.querySelector<HTMLElement>('.coco-ppt-context-menu')?.dataset.shapeId).toBe(
        'text',
      );
    });
    expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: 'text' });

    const textBefore = editor.session().scene().slides[0]!.nodes[0]!.bounds;
    const overlayBefore = editor.session().scene().slides[0]!.nodes[1]!.bounds;
    overlay.dispatchEvent(mouse('mousedown', clientX, clientY, { cancelable: true }));
    window.dispatchEvent(mouse('mousemove', clientX + 30, clientY + 20));
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', clientX + 30, clientY + 20));
    await changed;

    expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).toBeGreaterThan(
      textBefore.left,
    );
    expect(editor.session().scene().slides[0]!.nodes[1]!.bounds).toMatchObject(overlayBefore);

    await editor.destroy();
  });

  test('keeps an explicitly selected overlapped object as the drag target', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-overlap-selected-target-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: 'text',
              name: 'Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Editable underlay' },
            },
            {
              id: 'overlay',
              name: 'Decorative overlay',
              kind: 'geometricShape',
              raw: { presetGeometry: 'rect' },
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
            },
          ],
        },
      ],
    });
    Object.assign(view.scene.slides[0]!.nodes[1]!, {
      capabilities: {
        select: true,
        editText: false,
        editTextStyle: false,
        editTableCells: false,
        editGeometry: true,
        editShapeStyle: true,
      },
    });
    const editor = renderNativePptEditor(host, view);
    const root = host.shadowRoot!;
    const canvas = root.querySelector<HTMLElement>('.coco-ppt-canvas')!;
    const overlay = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="overlay"]')!;
    const scale = Number.parseFloat(canvas.style.width) / 9144000;
    const clientX = 10;
    const clientY = 10;
    canvas.getBoundingClientRect = () =>
      ({
        left: clientX - (100000 + 1000000) * scale,
        top: clientY - (100000 + 250000) * scale,
        width: 1000,
        height: 600,
        right: 1000,
        bottom: 600,
        x: 0,
        y: 0,
        toJSON: () => ({}),
      }) as DOMRect;

    await editor.dispatch({ kind: 'select-shape', slideIdOrIndex: 0, shapeIdOrName: 'overlay' });
    const textBefore = editor.session().scene().slides[0]!.nodes[0]!.bounds;
    const overlayBefore = editor.session().scene().slides[0]!.nodes[1]!.bounds;
    overlay.dispatchEvent(mouse('mousedown', clientX, clientY, { cancelable: true }));
    window.dispatchEvent(mouse('mousemove', clientX + 30, clientY + 20));
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', clientX + 30, clientY + 20));
    await changed;

    expect(editor.session().scene().slides[0]!.nodes[0]!.bounds).toMatchObject(textBefore);
    expect(editor.session().scene().slides[0]!.nodes[1]!.bounds.left).toBeGreaterThan(
      overlayBefore.left,
    );

    await editor.destroy();
  });

  test('cycles overlapped selectable objects with Alt click', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-overlap-cycle-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: 'bottom',
              name: 'Bottom text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Bottom' },
            },
            {
              id: 'middle',
              name: 'Middle shape',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
            },
            {
              id: 'top',
              name: 'Top shape',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
            },
          ],
        },
      ],
    });
    for (const node of view.scene.slides[0]!.nodes.slice(1)) {
      Object.assign(node, {
        kind: 'geometricShape',
        capabilities: {
          select: true,
          editText: false,
          editTextStyle: false,
          editTableCells: false,
          editGeometry: true,
          editShapeStyle: true,
        },
      });
    }
    const events: string[] = [];
    const editor = renderNativePptEditor(host, view, {
      onObjectAction: (event) => events.push(`${event.shapeId}:${event.reason}`),
    });
    const root = host.shadowRoot!;
    const canvas = root.querySelector<HTMLElement>('.coco-ppt-canvas')!;
    const top = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="top"]')!;
    const scale = Number.parseFloat(canvas.style.width) / 9144000;
    const clientX = 10;
    const clientY = 10;
    canvas.getBoundingClientRect = () =>
      ({
        left: clientX - (100000 + 1000000) * scale,
        top: clientY - (100000 + 250000) * scale,
        width: 1000,
        height: 600,
        right: 1000,
        bottom: 600,
        x: 0,
        y: 0,
        toJSON: () => ({}),
      }) as DOMRect;

    await editor.dispatch({ kind: 'select-shape', slideIdOrIndex: '256', shapeIdOrName: 'top' });
    top.dispatchEvent(new MouseEvent('click', { bubbles: true, altKey: true, clientX, clientY }));
    await waitFor(() =>
      expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: 'middle' }),
    );
    top.dispatchEvent(new MouseEvent('click', { bubbles: true, altKey: true, clientX, clientY }));
    await waitFor(() =>
      expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: 'bottom' }),
    );

    expect(events).toEqual(['middle:已切换到重叠对象 2/3', 'bottom:已切换到重叠对象 3/3']);

    await editor.destroy();
  });

  test('renders the target slide when selecting a shape on another slide', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-cross-slide-selection-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'First',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'First' },
            },
          ],
        },
        {
          id: '257',
          index: 1,
          shapes: [
            {
              id: '3',
              name: 'Second',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Second' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);

    await editor.dispatch({ kind: 'select-shape', slideIdOrIndex: '257', shapeIdOrName: 'Second' });

    expect(host.shadowRoot!.querySelector('[data-shape-id="2"]')).toBeNull();
    expect(host.shadowRoot!.querySelector('[data-shape-id="3"]')).not.toBeNull();

    await editor.destroy();
  });

  test('renders real slide thumbnails without exposing editable shape selectors', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-slide-thumbnail-preview-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'First',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'First thumbnail' },
            },
          ],
        },
        {
          id: '257',
          index: 1,
          shapes: [
            {
              id: '3',
              name: 'Second',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Second thumbnail' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const root = host.shadowRoot!;
    const thumbs = root.querySelectorAll<HTMLButtonElement>('.coco-ppt-thumb');

    expect(thumbs).toHaveLength(2);
    expect(thumbs[0]!.querySelector('.coco-ppt-thumb-node')?.textContent).toContain(
      'First thumbnail',
    );
    expect(thumbs[0]!.querySelector('.coco-ppt-thumb-number')?.textContent).toBe('1');
    expect(thumbs[0]!.querySelector('.coco-ppt-thumb-placeholder')?.textContent).toBe('');
    expect(root.querySelector('.coco-ppt-thumb-node[data-shape-id]')).toBeNull();
    expect(root.querySelector('.coco-ppt-thumb .coco-ppt-resize-handle')).toBeNull();
    const thumbnailSlide = thumbs[0]!.querySelector<HTMLElement>('.coco-ppt-thumb-slide')!;
    const thumbnailNode = thumbs[0]!.querySelector<HTMLElement>('.coco-ppt-thumb-node')!;
    expect(thumbnailSlide.style.width).toBe('512px');
    expect(thumbnailSlide.style.transform).toBe('scale(0.25)');
    expect(Number.parseFloat(thumbnailNode.style.width)).toBeGreaterThan(100);

    thumbs[1]!.click();
    await waitFor(() =>
      expect(editor.selection()).toMatchObject({ kind: 'slide', slideId: '257' }),
    );

    await editor.destroy();
  });

  test('renders layout-inherited placeholders at their master geometry in the slide and thumbnail', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromPptx(await makeNativePlaceholderInheritancePptx());
    expect(view.scene.slides[0]!.nodes[0]!.bounds).toEqual({
      left: 457200,
      top: 274638,
      width: 8229600,
      height: 1143000,
    });
    const editor = renderNativePptEditor(host, view);
    const root = host.shadowRoot!;
    const node = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="2"]')!;
    const thumbnailNode = root.querySelector<HTMLElement>('.coco-ppt-thumb-node')!;

    expect(node.textContent).toContain('Office preview');
    expect(Number.parseFloat(thumbnailNode.style.width)).toBeGreaterThan(100);
    expect(Number.parseFloat(thumbnailNode.style.height)).toBeGreaterThan(50);
    expect(thumbnailNode.textContent).toContain('Office preview');

    await editor.destroy();
  });

  test('does not render locked decorative shape names as slide text', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-locked-decoration-name-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: '矩形: 圆角 16',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 700000 },
              style: { fillColor: '#1f3f8f' },
              locked: true,
            },
            {
              id: '3',
              name: '真实标题',
              kind: 'textBox',
              geometry: { left: 100000, top: 900000, width: 2000000, height: 700000 },
              text: { text: '真实标题' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const root = host.shadowRoot!;

    expect(root.querySelector('[data-shape-id="2"]')?.textContent).toBe('');
    expect(root.querySelector('[data-shape-id="3"]')?.textContent).toContain('真实标题');
    expect(root.textContent).not.toContain('矩形: 圆角 16');

    await editor.destroy();
  });

  test('does not apply editor hatch backgrounds to locked transparent text shapes', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-locked-transparent-text-style-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: '文本框 119',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 700000 },
              style: { fillColor: '#000000', opacity: 0, lineColor: '#000000', lineWidth: 0 },
              text: { text: '透明文本' },
              locked: true,
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>('[data-shape-id="2"]')!;

    expect(node.textContent).toContain('透明文本');
    expect(node.style.background).not.toContain('repeating-linear-gradient');
    expect(host.shadowRoot!.querySelector('style')?.textContent).not.toContain(
      'repeating-linear-gradient(45deg',
    );

    await editor.destroy();
  });

  test('virtualizes slide thumbnail rendering around the visible scroll range', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const heightDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight');
    Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-slides') ? 160 : 0;
      },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-slide-thumbnail-virtualization-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: Array.from({ length: 12 }, (_, index) => ({
        id: String(256 + index),
        index,
        shapes: [
          {
            id: String(2 + index),
            name: `Slide ${index + 1}`,
            kind: 'textBox',
            geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
            text: { text: `Thumb ${index + 1}` },
          },
        ],
      })),
    });
    const editor = renderNativePptEditor(host, view);
    try {
      const root = host.shadowRoot!;
      const list = root.querySelector<HTMLElement>('.coco-ppt-slides')!;
      const renderedBefore = root.querySelectorAll('.coco-ppt-thumb[data-preview-rendered="true"]');

      expect(renderedBefore.length).toBeGreaterThan(0);
      expect(renderedBefore.length).toBeLessThan(12);
      expect(
        root.querySelector('.coco-ppt-thumb[data-slide-index="10"][data-preview-rendered="true"]'),
      ).toBeNull();

      await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: '262' });
      list.scrollTop = 560;
      list.dispatchEvent(new Event('scroll'));
      await new Promise((resolve) => requestAnimationFrame(resolve));

      expect(
        root.querySelector('.coco-ppt-thumb[data-slide-index="0"][data-preview-rendered="true"]'),
      ).toBeNull();
      expect(
        root.querySelector('.coco-ppt-thumb[data-slide-index="10"][data-preview-rendered="true"]'),
      ).not.toBeNull();
    } finally {
      if (heightDescriptor)
        Object.defineProperty(HTMLElement.prototype, 'clientHeight', heightDescriptor);
      await editor.destroy();
    }
  });

  test('reuses cached canvas thumbnails for unchanged slides after editing one slide', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-slide-thumbnail-per-slide-cache-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'First',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'First thumbnail' },
            },
          ],
        },
        {
          id: '257',
          index: 1,
          shapes: [
            {
              id: '3',
              name: 'Second',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Second thumbnail' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const cache = (
      editor as unknown as {
        thumbnailCanvasCache: Map<string, { fingerprint: string; canvas: HTMLCanvasElement }>;
      }
    ).thumbnailCanvasCache;
    for (const slide of editor.session().scene().slides) {
      cache.set(slide.id, {
        fingerprint: testThumbnailFingerprint(slide),
        canvas: document.createElement('canvas'),
      });
    }

    await editor.dispatch({
      kind: 'set-shape-text',
      slideIdOrIndex: '256',
      shapeIdOrName: 'First',
      text: 'Changed thumbnail',
    });

    const root = host.shadowRoot!;
    const firstThumb = root.querySelector<HTMLElement>('.coco-ppt-thumb[data-slide-index="0"]')!;
    const secondThumb = root.querySelector<HTMLElement>('.coco-ppt-thumb[data-slide-index="1"]')!;

    expect(firstThumb.dataset.previewMode).toBe('dom');
    expect(firstThumb.querySelector('.coco-ppt-thumb-slide')?.textContent).toContain(
      'Changed thumbnail',
    );
    expect(secondThumb.dataset.previewMode).toBe('canvas');
    expect(secondThumb.querySelector('.coco-ppt-thumb-canvas')).not.toBeNull();
    expect(secondThumb.querySelector('.coco-ppt-thumb-slide')).toBeNull();

    await editor.destroy();
  });

  test('updates the current slide thumbnail after geometry edits while reusing unchanged canvases', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-slide-thumbnail-geometry-refresh-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'First',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Movable thumbnail' },
            },
          ],
        },
        {
          id: '257',
          index: 1,
          shapes: [
            {
              id: '3',
              name: 'Second',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Stable thumbnail' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const cache = (
      editor as unknown as {
        thumbnailCanvasCache: Map<string, { fingerprint: string; canvas: HTMLCanvasElement }>;
      }
    ).thumbnailCanvasCache;
    for (const slide of editor.session().scene().slides) {
      cache.set(slide.id, {
        fingerprint: testThumbnailFingerprint(slide),
        canvas: document.createElement('canvas'),
      });
    }

    await editor.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: '256',
      shapeIdOrName: 'First',
      geometry: { left: 400000, top: 260000 },
    });

    const root = host.shadowRoot!;
    const firstThumb = root.querySelector<HTMLElement>('.coco-ppt-thumb[data-slide-index="0"]')!;
    const secondThumb = root.querySelector<HTMLElement>('.coco-ppt-thumb[data-slide-index="1"]')!;
    const movedThumbNode = firstThumb.querySelector<HTMLElement>('.coco-ppt-thumb-node')!;

    expect(firstThumb.dataset.previewMode).toBe('dom');
    expect(movedThumbNode.textContent).toContain('Movable thumbnail');
    expect(Number.parseFloat(movedThumbNode.style.left)).toBeGreaterThan((100000 * 128) / 9144000);
    expect(Number.parseFloat(movedThumbNode.style.top)).toBeGreaterThan((100000 * 128) / 9144000);
    expect(secondThumb.dataset.previewMode).toBe('canvas');
    expect(secondThumb.querySelector('.coco-ppt-thumb-canvas')).not.toBeNull();
    expect(secondThumb.querySelector('.coco-ppt-thumb-slide')).toBeNull();

    await editor.destroy();
  });

  test('edits the text layer for text shapes with rendered geometry overlays', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-text-layer-edit-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '21',
              name: 'Banner',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 3000000, height: 600000 },
              raw: { presetGeometry: 'parallelogram' },
              text: { text: 'Draft' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.customGeometry = {
      paths: [
        {
          width: 100000,
          height: 100000,
          commands: [
            { kind: 'moveTo', x: 0, y: 0 },
            { kind: 'lineTo', x: 100000, y: 0 },
            { kind: 'lineTo', x: 92000, y: 100000 },
            { kind: 'lineTo', x: 0, y: 100000 },
            { kind: 'close' },
          ],
        },
      ],
    };
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="21"]',
    )!;
    const textLayer = node.querySelector<HTMLElement>('.coco-ppt-text-layer')!;

    expect(node.querySelector('.coco-ppt-custom-geometry')).not.toBeNull();
    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    textLayer.textContent = 'Final';
    textLayer.dispatchEvent(new FocusEvent('blur'));

    await waitFor(() => expect(editor.session().scene().slides[0]!.nodes[0]!.text).toBe('Final'));
    expect(editor.session().selection()).toMatchObject({
      kind: 'shape',
      slideId: '256',
      shapeId: '21',
    });

    await editor.destroy();
  });

  test('normalizes SVG gradient stops for custom geometry fills', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-custom-gradient-stop-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '22',
              name: 'Custom Gradient',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1000000, height: 1000000 },
              raw: { presetGeometry: 'rect' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.fillGradient = {
      colors: ['rgba(154, 211, 245, 0)', 'rgba(154, 211, 245, 0)', '#56B6EE'],
      positions: [0, 98, 50],
      opacities: [0, 0, 1],
    };
    view.scene.slides[0]!.nodes[0]!.opacity = 0.35;
    view.scene.slides[0]!.nodes[0]!.customGeometry = {
      paths: [
        {
          width: 100000,
          height: 100000,
          commands: [
            { kind: 'moveTo', x: 0, y: 100000 },
            { kind: 'cubicBezierTo', x1: 25000, y1: 0, x2: 75000, y2: 100000, x: 100000, y: 0 },
            { kind: 'lineTo', x: 100000, y: 100000 },
            { kind: 'close' },
          ],
        },
      ],
    };

    const editor = renderNativePptEditor(host, view);
    const stops = host.shadowRoot!.querySelectorAll<SVGStopElement>(
      '.coco-ppt-canvas .coco-ppt-custom-geometry stop',
    );

    expect(Array.from(stops).map((stop) => stop.getAttribute('offset'))).toEqual([
      '0%',
      '50%',
      '98%',
    ]);
    expect(stops[0]!.getAttribute('stop-color')).toBe('#9AD3F5');
    expect(stops[0]!.getAttribute('stop-opacity')).toBe('0');
    const path = host.shadowRoot!.querySelector<SVGPathElement>(
      '.coco-ppt-canvas .coco-ppt-custom-geometry path',
    )!;
    expect(path.getAttribute('fill-opacity')).toBe('0.35');

    await editor.destroy();
  });

  test('keeps run colors when node text style has a gradient default', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-rich-text-node-gradient-inheritance-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '30',
              name: 'Mixed Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2200000, height: 500000 },
              text: {
                text: '项目展示',
                style: {
                  fontSize: 24,
                  color: '#000000',
                  textGradient: { colors: ['#3878D0', '#56B6EE'] },
                },
                paragraphs: [
                  {
                    runs: [
                      { text: '项目', style: { color: '#000000' } },
                      {
                        text: '展示',
                        style: {
                          color: '#3878D0',
                          textGradient: { colors: ['#3878D0', '#56B6EE'] },
                        },
                      },
                    ],
                  },
                ],
              },
            },
          ],
        },
      ],
    });

    const editor = renderNativePptEditor(host, view);
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.textStyle = {
      fontSize: 24,
      color: '#000000',
      textGradient: { colors: ['#3878D0', '#56B6EE'] },
    };
    nodeModel.textParagraphs = [
      {
        text: '项目展示',
        runs: [
          { text: '项目', style: { color: '#000000' } },
          {
            text: '展示',
            style: {
              color: '#3878D0',
              textGradient: { colors: ['#3878D0', '#56B6EE'] },
            },
          },
        ],
      },
    ];
    editor.render();
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="30"]',
    )!;
    const spans = node.querySelectorAll<HTMLElement>('.coco-ppt-rich-text span');

    expect(node.style.getPropertyValue('-webkit-text-fill-color')).toBe('');
    expect(spans[0]!.style.color).toBe('#000000');
    expect(spans[0]!.style.getPropertyValue('-webkit-text-fill-color')).toBe('');
    expect(spans[1]!.style.getPropertyValue('-webkit-text-fill-color')).toBe('transparent');

    await editor.destroy();
  });

  test('renders multi-subpath custom geometry with even-odd fill', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-custom-evenodd-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '23',
              name: 'Cutout Logo',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1000000, height: 1000000 },
              raw: { presetGeometry: 'rect' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.fillColor = '#05E1FF';
    view.scene.slides[0]!.nodes[0]!.opacity = 0.15;
    view.scene.slides[0]!.nodes[0]!.customGeometry = {
      paths: [
        {
          width: 100,
          height: 100,
          commands: [
            { kind: 'moveTo', x: 0, y: 0 },
            { kind: 'lineTo', x: 100, y: 0 },
            { kind: 'lineTo', x: 100, y: 100 },
            { kind: 'lineTo', x: 0, y: 100 },
            { kind: 'close' },
            { kind: 'moveTo', x: 35, y: 35 },
            { kind: 'lineTo', x: 65, y: 35 },
            { kind: 'lineTo', x: 65, y: 65 },
            { kind: 'lineTo', x: 35, y: 65 },
            { kind: 'close' },
          ],
        },
      ],
    };

    const editor = renderNativePptEditor(host, view);
    const path = host.shadowRoot!.querySelector<SVGPathElement>(
      '.coco-ppt-canvas .coco-ppt-custom-geometry path',
    )!;

    expect(path.getAttribute('fill-rule')).toBe('evenodd');
    expect(path.getAttribute('fill-opacity')).toBe('0.15');

    await editor.destroy();
  });

  test('renders slide image backgrounds from editor image resources', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-slide-background-image-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          background: { imagePath: 'ppt/media/background.png' },
          shapes: [],
        },
      ],
    });
    view.scene.slides[0]!.background = { imagePath: 'ppt/media/background.png' };
    view.images['ppt/media/background.png'] = {
      url: 'blob:coco-ppt-background',
      contentType: 'image/png',
    };

    const editor = renderNativePptEditor(host, view);
    const canvas = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-canvas')!;
    const thumb = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-thumb-slide')!;

    expect(canvas.style.backgroundImage).toContain('blob:coco-ppt-background');
    expect(thumb.style.backgroundImage).toContain('blob:coco-ppt-background');

    await editor.destroy();
  });

  test('renders useBgFill shapes from inferred full-slide picture backgrounds', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const model = {
      id: 'native-use-bg-fill-shape-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          layout: {
            id: 'layout-1',
            name: 'Image Layout',
            shapes: [
              {
                id: 'layout-bg',
                name: 'Layout Background',
                kind: 'picture',
                geometry: { left: 0, top: 0, width: 9144000, height: 5143500 },
                image: {
                  relationshipId: 'rId1',
                  targetPath: 'ppt/media/background.png',
                  filename: 'background.png',
                },
              },
            ],
          },
          shapes: [
            {
              id: '7',
              name: 'Background Fill Mask',
              kind: 'geometricShape',
              geometry: { left: 914400, top: 514350, width: 1828800, height: 1028700 },
              style: { fillColor: '#5B9BD5' },
              raw: { presetGeometry: 'rect', useBackgroundFill: true },
            },
          ],
        },
      ],
    };
    const view = await buildPptEditorViewFromModel(model);
    view.scene = buildPptPresentationScene(model);
    view.images['ppt/media/background.png'] = {
      url: 'blob:coco-ppt-background',
      contentType: 'image/png',
    };

    expect(view.scene.slides[0]!.nodes.find((node) => node.id === '7')).toMatchObject({
      usesBackgroundFill: true,
      backgroundFill: { imagePath: 'ppt/media/background.png' },
    });

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="7"]',
    )!;

    expect(node.style.backgroundImage).toContain('blob:coco-ppt-background');
    expect(node.style.backgroundRepeat).toBe('no-repeat');
    expect(node.style.backgroundSize).not.toBe('');
    expect(node.style.backgroundPosition).toContain('-');

    await editor.destroy();
  });

  test('marks unsupported vector images without rendering broken image elements', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-unsupported-image-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '5',
              name: 'EMF Picture',
              kind: 'picture',
              geometry: { left: 100000, top: 100000, width: 1000000, height: 1000000 },
              image: {
                relationshipId: 'rId1',
                targetPath: 'ppt/media/image1.emf',
                filename: 'image1.emf',
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.kind = 'picture';
    view.scene.slides[0]!.nodes[0]!.imagePath = 'ppt/media/image1.emf';
    view.images['ppt/media/image1.emf'] = {
      url: 'blob:coco-ppt-emf',
      contentType: 'image/x-emf',
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="5"]',
    )!;

    expect(node.dataset.unsupportedImage).toBe('true');
    expect(node.dataset.imageContentType).toBe('image/x-emf');
    expect(node.querySelector('img')).toBeNull();
    const placeholder = node.querySelector<HTMLElement>('.coco-ppt-unsupported-image')!;
    expect(placeholder).not.toBeNull();
    expect(placeholder.textContent).toBe('EMF');

    await editor.destroy();
  });

  test('renders hyperlink runs with default link styling', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-hyperlink-run-style-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '5',
              name: 'Link Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 4000000, height: 500000 },
              text: {
                text: 'PPT模板：www.1ppt.com/moban/',
                paragraphs: [
                  {
                    text: 'PPT模板：www.1ppt.com/moban/',
                    runs: [
                      { text: 'PPT模板：' },
                      {
                        text: 'www.1ppt.com/moban/',
                        hyperlink: { target: 'https://www.1ppt.com/moban/' },
                      },
                    ],
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textParagraphs = [
      {
        text: 'PPT模板：www.1ppt.com/moban/',
        runs: [
          { text: 'PPT模板：' },
          {
            text: 'www.1ppt.com/moban/',
            hyperlink: { target: 'https://www.1ppt.com/moban/' },
          },
        ],
      },
    ];

    const editor = renderNativePptEditor(host, view);
    const spans = host.shadowRoot!.querySelectorAll<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="5"] .coco-ppt-rich-text span',
    );

    expect(spans).toHaveLength(2);
    expect(spans[0]!.style.color).toBe('');
    expect(spans[0]!.style.textDecorationLine).toBe('');
    expect(spans[1]!.style.color).toBe('#0563C1');
    expect(spans[1]!.style.textDecorationLine).toBe('underline');

    await editor.destroy();
  });

  test('renders hyperlink runs with office link color over plain run colors', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-hyperlink-run-color-override-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '5',
              name: 'Link Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 4000000, height: 500000 },
              text: {
                text: 'www.1ppt.com/moban/',
                paragraphs: [
                  {
                    text: 'www.1ppt.com/moban/',
                    runs: [
                      {
                        text: 'www.1ppt.com/moban/',
                        style: { color: '#3C3B38' },
                        hyperlink: { target: 'https://www.1ppt.com/moban/' },
                      },
                    ],
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textParagraphs = [
      {
        text: 'www.1ppt.com/moban/',
        runs: [
          {
            text: 'www.1ppt.com/moban/',
            style: { color: '#3C3B38' },
            hyperlink: { target: 'https://www.1ppt.com/moban/' },
          },
        ],
      },
    ];

    const editor = renderNativePptEditor(host, view);
    const span = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="5"] .coco-ppt-rich-text span',
    )!;

    expect(span.style.color).toBe('#0563C1');
    expect(span.style.textDecorationLine).toBe('underline');

    await editor.destroy();
  });

  test('does not start dragging while editing an inner text layer', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-text-layer-no-drag-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '22',
              name: 'Banner',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 3000000, height: 600000 },
              raw: { presetGeometry: 'parallelogram' },
              text: { text: 'Draft' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.customGeometry = {
      paths: [
        {
          width: 100000,
          height: 100000,
          commands: [
            { kind: 'moveTo', x: 0, y: 0 },
            { kind: 'lineTo', x: 100000, y: 0 },
            { kind: 'lineTo', x: 92000, y: 100000 },
            { kind: 'lineTo', x: 0, y: 100000 },
            { kind: 'close' },
          ],
        },
      ],
    };
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="22"]',
    )!;
    const canvas = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-canvas')!;
    const textLayer = node.querySelector<HTMLElement>('.coco-ppt-text-layer')!;

    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(canvas.dataset.interactionMode).toBe('editing');
    expect(node.dataset.editing).toBe('true');
    textLayer.dispatchEvent(mouse('mousedown', 10, 10));
    window.dispatchEvent(mouse('mousemove', 40, 60));
    window.dispatchEvent(mouse('mouseup', 40, 60));

    expect(editor.session().scene().slides[0]!.nodes[0]!.bounds).toMatchObject({
      left: 100000,
      top: 100000,
    });

    await editor.destroy();
  });

  test('moves shapes through geometry commands', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;
    const canvas = root.querySelector<HTMLElement>('.coco-ppt-canvas')!;

    node.dispatchEvent(mouse('mousedown', 10, 10));
    window.dispatchEvent(mouse('mousemove', 30, 40));
    expect(canvas.dataset.interactionMode).toBe('dragging');
    expect(node.dataset.dragging).toBe('true');
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', 30, 40));
    await changed;
    expect(canvas.dataset.interactionMode).toBeUndefined();
    expect(node.dataset.dragging).toBeUndefined();
    node.dispatchEvent(mouse('click', 30, 40));

    await waitFor(() => {
      const bounds = editor.session().scene().slides[0]!.nodes[0]!.bounds;
      expect(bounds.left).toBeGreaterThan(100000);
      expect(bounds.top).toBeGreaterThan(100000);
    });
    const movedBounds = editor.session().scene().slides[0]!.nodes[0]!.bounds;
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(editor.session().scene().slides[0]!.nodes[0]!.bounds).toMatchObject(movedBounds);

    await editor.destroy();
  });

  test('keeps drag preview in place while geometry commit is pending', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;
    const originalDispatch = editor.session().dispatch.bind(editor.session());
    let releaseCommit: (() => void) | null = null;
    let commitStarted = false;
    editor.session().dispatch = async (command) => {
      if (command.kind === 'set-shape-geometry') {
        commitStarted = true;
        await new Promise<void>((resolve) => {
          releaseCommit = resolve;
        });
      }
      await originalDispatch(command);
    };

    node.dispatchEvent(mouse('mousedown', 10, 10));
    window.dispatchEvent(mouse('mousemove', 40, 40));
    const preview = { left: node.style.left, top: node.style.top };
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', 40, 40));

    expect(commitStarted).toBe(true);
    expect(node.style.left).toBe(preview.left);
    expect(node.style.top).toBe(preview.top);
    expect(
      root.querySelector<HTMLElement>('.coco-ppt-canvas')!.dataset.interactionMode,
    ).toBeUndefined();

    releaseCommit?.();
    await changed;
    await waitFor(() => {
      expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).toBeGreaterThan(100000);
    });

    await editor.destroy();
  });

  test('keeps resize preview in place while geometry commit is pending', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;
    const handle = node.querySelector<HTMLElement>('.coco-ppt-resize-handle')!;
    const originalDispatch = editor.session().dispatch.bind(editor.session());
    let releaseCommit: (() => void) | null = null;
    let commitStarted = false;
    editor.session().dispatch = async (command) => {
      if (command.kind === 'set-shape-geometry') {
        commitStarted = true;
        await new Promise<void>((resolve) => {
          releaseCommit = resolve;
        });
      }
      await originalDispatch(command);
    };

    handle.dispatchEvent(mouse('mousedown', 20, 20));
    window.dispatchEvent(mouse('mousemove', 60, 50));
    const preview = { width: node.style.width, height: node.style.height };
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', 60, 50));

    expect(commitStarted).toBe(true);
    expect(node.style.width).toBe(preview.width);
    expect(node.style.height).toBe(preview.height);
    expect(
      root.querySelector<HTMLElement>('.coco-ppt-canvas')!.dataset.interactionMode,
    ).toBeUndefined();

    releaseCommit?.();
    await changed;
    await waitFor(() => {
      expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.width).toBeGreaterThan(2000000);
    });

    await editor.destroy();
  });

  test('keeps text unscaled while resizing text boxes', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-resize-text-wrap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: {
                text: '未来仍需全力以赴攀高峰',
                body: { autoFit: 'normal', wrap: 'none' },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textParagraphs = [
      {
        text: '未来仍需全力以赴攀高峰',
        runs: [{ text: '未来仍需全力以赴攀高峰', style: { color: '#1E3A8A' } }],
      },
    ];
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="2"]',
    )!;
    const richText = node.querySelector<HTMLElement>('.coco-ppt-rich-text')!;
    const handle = node.querySelector<HTMLElement>('.coco-ppt-resize-handle')!;
    richText.style.transform = 'scale(0.5)';
    richText.style.width = '200%';

    handle.dispatchEvent(mouse('mousedown', 10, 10, { cancelable: true }));
    window.dispatchEvent(mouse('mousemove', -40, -30));

    expect(richText.style.transform).toBe('');
    expect(richText.style.width).toBe('');
    expect(richText.style.whiteSpace).toBe('pre-wrap');

    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', -40, -30));
    await changed;

    await waitFor(() => {
      const resized = editor.session().scene().slides[0]!.nodes[0]!;
      expect(resized.textBody).toMatchObject({ autoFit: 'none', wrap: 'square' });
      expect(resized.bounds.width).toBeLessThan(2000000);
    });

    await editor.destroy();
  });

  test('does not reapply queued auto-fit scaling after manual text resize starts', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-resize-text-queued-autofit-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: {
                text: '未来仍需全力以赴攀高峰',
                body: { autoFit: 'normal', wrap: 'square' },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textParagraphs = [
      {
        text: '未来仍需全力以赴攀高峰',
        runs: [{ text: '未来仍需全力以赴攀高峰', style: { fontSize: 44 } }],
      },
    ];
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="2"]',
    )!;
    const richText = node.querySelector<HTMLElement>('.coco-ppt-rich-text')!;
    const handle = node.querySelector<HTMLElement>('.coco-ppt-resize-handle')!;
    Object.defineProperty(node, 'clientWidth', { configurable: true, value: 80 });
    Object.defineProperty(node, 'clientHeight', { configurable: true, value: 24 });
    Object.defineProperty(node, 'scrollWidth', { configurable: true, value: 240 });
    Object.defineProperty(node, 'scrollHeight', { configurable: true, value: 24 });

    handle.dispatchEvent(mouse('mousedown', 10, 10, { cancelable: true }));
    await new Promise((resolve) => requestAnimationFrame(resolve));

    expect(node.dataset.textResizePreview).toBe('wrap');
    expect(richText.dataset.textResizePreview).toBe('wrap');
    expect(richText.style.transform).toBe('');
    expect(richText.style.width).toBe('');

    window.dispatchEvent(mouse('mouseup', 10, 10));
    await editor.destroy();
  });

  test('prevents browser text selection while dragging text shapes', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    const allowedDefault = node.dispatchEvent(
      new MouseEvent('mousedown', {
        bubbles: true,
        cancelable: true,
        clientX: 10,
        clientY: 10,
        button: 0,
      }),
    );
    window.dispatchEvent(mouse('mousemove', 30, 40));
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', 30, 40));
    await changed;

    expect(allowedDefault).toBe(false);
    expect(root.querySelector('style')?.textContent).toContain('user-select: none');
    expect(root.querySelector('style')?.textContent).toContain('data-dragging="true"');
    expect(root.querySelector('style')?.textContent).toContain('data-snap-reference="true"');

    await waitFor(() =>
      expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).not.toBe(100000),
    );
    await editor.destroy();
  });

  test('drags rich text shapes from their rendered text spans without native text selection', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-rich-text-span-drag-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Rich Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: {
                text: '未来可期',
                body: { autoFit: 'shape', wrap: 'square' },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textParagraphs = [
      {
        text: '未来可期',
        runs: [
          { text: '未来', style: { bold: true } },
          { text: '可期', style: { color: '#4F73C9' } },
        ],
      },
    ];
    const editor = renderNativePptEditor(host, view);
    const span = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="2"] .coco-ppt-rich-text span',
    )!;

    const allowedDefault = span.dispatchEvent(mouse('mousedown', 10, 10, { cancelable: true }));
    window.dispatchEvent(mouse('mousemove', 40, 40));
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', 40, 40));
    await changed;

    expect(allowedDefault).toBe(false);
    await waitFor(() =>
      expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).not.toBe(100000),
    );

    await editor.destroy();
  });

  test('starts direct text editing from rendered rich text spans', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-rich-text-span-edit-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Rich Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: {
                text: '未来可期',
                body: { autoFit: 'shape', wrap: 'square' },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textParagraphs = [
      {
        text: '未来可期',
        runs: [
          { text: '未来', style: { bold: true } },
          { text: '可期', style: { color: '#4F73C9' } },
        ],
      },
    ];
    const editor = renderNativePptEditor(host, view);
    const span = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="2"] .coco-ppt-rich-text span',
    )!;

    span.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));

    await waitFor(() =>
      expect(
        host.shadowRoot!.querySelector<HTMLElement>(
          '.coco-ppt-canvas .coco-ppt-node[data-shape-id="2"]',
        )!.contentEditable,
      ).toBe('true'),
    );
    expect(editor.selection()).toMatchObject({ kind: 'shape', slideId: '256', shapeId: '2' });

    await editor.destroy();
  });

  test('blocks text selection on the canvas except while direct text editing is active', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    expect(node.dispatchEvent(new Event('selectstart', { bubbles: true, cancelable: true }))).toBe(
      false,
    );

    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    const editable = root.querySelector<HTMLElement>('[contenteditable="true"]')!;

    expect(
      editable.dispatchEvent(new Event('selectstart', { bubbles: true, cancelable: true })),
    ).toBe(true);

    await editor.destroy();
  });

  test('edits flattened group child text and keeps the parent group affordance visible', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-group-child-text-edit-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: 'child-title',
              name: 'Child Title',
              kind: 'textBox',
              geometry: { left: 300000, top: 200000, width: 1600000, height: 360000 },
              text: { text: '组内标题' },
            },
          ],
        },
      ],
    });
    Object.assign(view.scene.slides[0]!.nodes[0]!, {
      id: 'group-1:child-title',
      groupContext: {
        parentId: 'group-1',
        parentName: 'Parent Group',
        parentBounds: { left: 100000, top: 100000, width: 3000000, height: 1200000 },
        parentTransform: {
          childLeft: 0,
          childTop: 0,
          childWidth: 3000000,
          childHeight: 1200000,
        },
      },
    });
    const editor = renderNativePptEditor(host, view);
    const root = host.shadowRoot!;
    const child = root.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="group-1:child-title"]',
    )!;

    child.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));

    await waitFor(() => expect(child.contentEditable).toBe('true'));
    expect(editor.selection()).toMatchObject({
      kind: 'shape',
      slideId: '256',
      shapeId: 'group-1:child-title',
    });
    const parent = root.querySelector<HTMLElement>('.coco-ppt-parent-group-outline')!;
    expect(parent).not.toBeNull();
    expect(parent.dataset.shapeId).toBe('group-1');
    expect(parent.querySelector('.coco-ppt-parent-group-handle')).not.toBeNull();
    expect(parent.querySelectorAll('.coco-ppt-parent-group-edge')).toHaveLength(4);

    await editor.destroy();
  });

  test('drags a real parent group from a selected flattened child', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const widthDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    const heightDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 1200 : 0;
      },
    });
    Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 720 : 0;
      },
    });
    try {
      const view = await buildPptEditorViewFromPptx(await makeNativeGroupedTablePptx());
      const editor = renderNativePptEditor(host, view);
      const root = host.shadowRoot!;

      await editor.dispatch({
        kind: 'select-shape',
        slideIdOrIndex: 0,
        shapeIdOrName: '6:15',
      });

      const parentBefore = root.querySelector<HTMLElement>('.coco-ppt-parent-group-outline')!;
      expect(parentBefore).not.toBeNull();
      expect(parentBefore.dataset.shapeId).toBe('6');
      expect(parentBefore.dataset.selected).toBe('false');
      const childBefore = editor
        .session()
        .scene()
        .slides[0]!.nodes.find((node) => node.id === '6:15')!.bounds;
      const edge = parentBefore.querySelector<HTMLElement>(
        '.coco-ppt-parent-group-edge[data-edge="right"]',
      )!;

      edge.dispatchEvent(mouse('mousedown', 20, 20, { cancelable: true }));
      window.dispatchEvent(mouse('mousemove', 60, 45));
      const changed = waitForNextChange(editor);
      window.dispatchEvent(mouse('mouseup', 60, 45));
      await changed;

      await waitFor(() =>
        expect(editor.selection()).toMatchObject({
          kind: 'shape',
          slideId: '256',
          shapeId: '6',
        }),
      );
      const slide = editor.session().scene().slides[0]!;
      const groupAfter = slide.groups.find((node) => node.id === '6')!;
      const childAfter = slide.nodes.find((node) => node.id === '6:15')!;
      expect(groupAfter.bounds.left).toBeGreaterThan(1000);
      expect(groupAfter.bounds.top).toBeGreaterThan(1000);
      expect(childAfter.bounds.left).toBeGreaterThan(childBefore.left);
      expect(childAfter.bounds.top).toBeGreaterThan(childBefore.top);
      const parentAfter = root.querySelector<HTMLElement>('.coco-ppt-parent-group-outline')!;
      expect(parentAfter.dataset.shapeId).toBe('6');
      expect(parentAfter.dataset.selected).toBe('true');

      await editor.destroy();
    } finally {
      if (widthDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientWidth', widthDescriptor);
      }
      if (heightDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientHeight', heightDescriptor);
      }
    }
  });

  test('marks real flattened group children as geometry-editable on canvas', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromPptx(await makeNativeGroupedTablePptx());
    const editor = renderNativePptEditor(host, view);
    const child = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="6:15"]',
    )!;

    expect(child.dataset.draggable).toBe('true');
    expect(child.dataset.resizable).toBe('true');
    expect(child.querySelector('.coco-ppt-resize-handle')).not.toBeNull();
    expect(
      editor
        .session()
        .scene()
        .slides[0]!.nodes.find((node) => node.id === '6:15'),
    ).toMatchObject({
      capabilities: {
        editGeometry: true,
        editShapeStyle: false,
      },
    });

    await editor.destroy();
  });

  test('keeps small pointer movement as selection instead of drag', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;
    const canvas = root.querySelector<HTMLElement>('.coco-ppt-canvas')!;
    const beforeStyle = { left: node.style.left, top: node.style.top };

    node.dispatchEvent(mouse('mousedown', 10, 10, { cancelable: true }));
    window.dispatchEvent(mouse('mousemove', 30, 30));
    expect(canvas.dataset.interactionMode).toBe('dragging');
    expect(node.dataset.dragging).toBe('true');
    window.dispatchEvent(mouse('mousemove', 12, 12));
    expect(canvas.dataset.interactionMode).toBeUndefined();
    expect(node.dataset.dragging).toBeUndefined();
    window.dispatchEvent(mouse('mouseup', 12, 12));

    expect(node.style.left).toBe(beforeStyle.left);
    expect(node.style.top).toBe(beforeStyle.top);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(editor.session().scene().slides[0]!.nodes[0]!.bounds).toMatchObject({
      left: 100000,
      top: 100000,
    });

    await editor.destroy();
  });

  test('snaps dragged shapes to nearby object alignment guides', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const events: string[] = [];
    const widthDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    const heightDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 1200 : 0;
      },
    });
    Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 720 : 0;
      },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-drag-snap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Moving',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Move' },
            },
            {
              id: '3',
              name: 'Target',
              kind: 'textBox',
              geometry: { left: 1200000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Target' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view, {
      onObjectAction: (event) => events.push(event.action, event.reason),
    });
    try {
      const root = host.shadowRoot!;
      const moving = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="2"]')!;
      const target = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="3"]')!;
      const startLeft = Number.parseFloat(moving.style.left);
      const targetLeft = Number.parseFloat(target.style.left);
      const dx = Math.round(targetLeft - startLeft - 3);

      moving.dispatchEvent(mouse('mousedown', 20, 20));
      window.dispatchEvent(mouse('mousemove', 20 + dx, 20));

      expect(root.querySelectorAll('.coco-ppt-snap-guide[data-axis="x"]').length).toBeGreaterThan(
        0,
      );
      expect(
        root.querySelector<HTMLElement>('.coco-ppt-snap-guide[data-axis="x"]')?.dataset
          .referenceShapeId,
      ).toBe('3');
      expect(target.dataset.snapReference).toBe('true');
      expect(Number.parseFloat(moving.style.left)).toBeCloseTo(targetLeft, 1);

      const changed = waitForNextChange(editor);
      window.dispatchEvent(mouse('mouseup', 20 + dx, 20));
      await changed;

      await waitFor(() =>
        expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).toBe(1200000),
      );
      await waitFor(() => expect(events).toContain('snap-geometry'));
      await waitFor(() => expect(events).toContain('已吸附到 Target'));
      expect(root.querySelectorAll('.coco-ppt-snap-guide').length).toBe(0);
      expect(target.dataset.snapReference).toBeUndefined();
    } finally {
      if (widthDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientWidth', widthDescriptor);
      }
      if (heightDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientHeight', heightDescriptor);
      }
      await editor.destroy();
    }
  });

  test('temporarily disables drag snapping while Alt is pressed', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const events: string[] = [];
    const widthDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    const heightDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 1200 : 0;
      },
    });
    Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 720 : 0;
      },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-drag-alt-no-snap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Moving',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Move' },
            },
            {
              id: '3',
              name: 'Target',
              kind: 'textBox',
              geometry: { left: 1200000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Target' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view, {
      onObjectAction: (event) => events.push(event.action),
    });
    try {
      const root = host.shadowRoot!;
      const moving = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="2"]')!;
      const target = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="3"]')!;
      const startLeft = Number.parseFloat(moving.style.left);
      const targetLeft = Number.parseFloat(target.style.left);
      const dx = Math.round(targetLeft - startLeft - 3);

      moving.dispatchEvent(mouse('mousedown', 20, 20));
      window.dispatchEvent(mouse('mousemove', 20 + dx, 20, { altKey: true }));

      expect(root.querySelectorAll('.coco-ppt-snap-guide').length).toBe(0);
      expect(target.dataset.snapReference).toBeUndefined();
      expect(Number.parseFloat(moving.style.left)).not.toBeCloseTo(targetLeft, 1);

      const changed = waitForNextChange(editor);
      window.dispatchEvent(mouse('mouseup', 20 + dx, 20, { altKey: true }));
      await changed;

      await waitFor(() =>
        expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).not.toBe(1200000),
      );
      expect(events).not.toContain('snap-geometry');
    } finally {
      if (widthDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientWidth', widthDescriptor);
      }
      if (heightDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientHeight', heightDescriptor);
      }
      await editor.destroy();
    }
  });

  test('does not show move snap guides for mismatched edge and center references', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const widthDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    const heightDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 1200 : 0;
      },
    });
    Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 720 : 0;
      },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-drag-snap-role-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Moving',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Move' },
            },
            {
              id: '3',
              name: 'Target',
              kind: 'textBox',
              geometry: { left: 1200000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Target' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    try {
      const root = host.shadowRoot!;
      const moving = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="2"]')!;
      const target = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="3"]')!;
      const startLeft = Number.parseFloat(moving.style.left);
      const targetCenter =
        Number.parseFloat(target.style.left) + Number.parseFloat(target.style.width) / 2;
      const dx = Math.round(targetCenter - startLeft - 3);

      moving.dispatchEvent(mouse('mousedown', 20, 20));
      window.dispatchEvent(mouse('mousemove', 20 + dx, 20));

      expect(root.querySelectorAll('.coco-ppt-snap-guide[data-axis="x"]').length).toBe(0);

      const changed = waitForNextChange(editor);
      window.dispatchEvent(mouse('mouseup', 20 + dx, 20));
      await changed;
      await waitFor(() =>
        expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).not.toBe(100000),
      );
    } finally {
      if (widthDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientWidth', widthDescriptor);
      }
      if (heightDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientHeight', heightDescriptor);
      }
      await editor.destroy();
    }
  });

  test('distinguishes page snap guides from object snap guides', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const widthDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    const heightDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 1200 : 0;
      },
    });
    Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 720 : 0;
      },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-page-snap-guide-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Moving',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Move' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    try {
      const root = host.shadowRoot!;
      const moving = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="2"]')!;
      const startLeft = Number.parseFloat(moving.style.left);
      const width = Number.parseFloat(moving.style.width);
      const canvas = root.querySelector<HTMLElement>('.coco-ppt-canvas')!;
      const canvasWidth = Number.parseFloat(canvas.style.width);
      const dx = Math.round(canvasWidth / 2 - (startLeft + width / 2) - 3);

      moving.dispatchEvent(mouse('mousedown', 20, 20));
      window.dispatchEvent(mouse('mousemove', 20 + dx, 20));

      const guide = root.querySelector<HTMLElement>(
        '.coco-ppt-snap-guide[data-axis="x"][data-kind="slide"]',
      );
      expect(guide).not.toBeNull();
      expect(guide?.dataset.referenceShapeId).toBeUndefined();
      expect(root.querySelectorAll('.coco-ppt-node[data-snap-reference="true"]').length).toBe(0);

      const changed = waitForNextChange(editor);
      window.dispatchEvent(mouse('mouseup', 20 + dx, 20));
      await changed;
    } finally {
      if (widthDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientWidth', widthDescriptor);
      }
      if (heightDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientHeight', heightDescriptor);
      }
      await editor.destroy();
    }
  });

  test('snaps resized shapes to nearby object alignment guides', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const events: string[] = [];
    const widthDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    const heightDescriptor = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 1200 : 0;
      },
    });
    Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
      configurable: true,
      get() {
        return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage') ? 720 : 0;
      },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-resize-snap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Sizing',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Size' },
            },
            {
              id: '3',
              name: 'Target',
              kind: 'textBox',
              geometry: { left: 1200000, top: 100000, width: 800000, height: 300000 },
              text: { text: 'Target' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view, {
      onObjectAction: (event) => events.push(event.action),
    });
    try {
      const root = host.shadowRoot!;
      const sizing = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="2"]')!;
      const target = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="3"]')!;
      const handle = sizing.querySelector<HTMLElement>('.coco-ppt-resize-handle')!;
      const startWidth = Number.parseFloat(sizing.style.width);
      const expectedWidth =
        Number.parseFloat(target.style.left) - Number.parseFloat(sizing.style.left);
      const dx = Math.round(expectedWidth - startWidth - 3);

      handle.dispatchEvent(mouse('mousedown', 20, 20));
      window.dispatchEvent(mouse('mousemove', 20 + dx, 20));

      expect(root.querySelectorAll('.coco-ppt-snap-guide[data-axis="x"]').length).toBeGreaterThan(
        0,
      );
      expect(
        root.querySelector<HTMLElement>('.coco-ppt-snap-guide[data-axis="x"]')?.dataset
          .referenceShapeId,
      ).toBe('3');
      expect(target.dataset.snapReference).toBe('true');
      expect(Number.parseFloat(sizing.style.width)).toBeCloseTo(expectedWidth, 1);

      const changed = waitForNextChange(editor);
      window.dispatchEvent(mouse('mouseup', 20 + dx, 20));
      await changed;

      await waitFor(() =>
        expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.width).toBe(1100000),
      );
      expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).toBe(100000);
      await waitFor(() => expect(events).toContain('snap-geometry'));
      expect(root.querySelectorAll('.coco-ppt-snap-guide').length).toBe(0);
      expect(target.dataset.snapReference).toBeUndefined();
    } finally {
      if (widthDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientWidth', widthDescriptor);
      }
      if (heightDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'clientHeight', heightDescriptor);
      }
      await editor.destroy();
    }
  });

  test('nudges selected shapes with arrow keys', async () => {
    const { editor } = await buildEditorHost();
    const events: string[] = [];
    await editor.destroy();

    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-keyboard-nudge-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Draft' },
            },
          ],
        },
      ],
    });
    const nudgeEditor = renderNativePptEditor(host, view, {
      onObjectAction: (event) => events.push(event.reason),
    });

    await nudgeEditor.dispatch({ kind: 'select-shape', slideIdOrIndex: 0, shapeIdOrName: '2' });
    window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' }));
    await waitFor(() =>
      expect(nudgeEditor.session().scene().slides[0]!.nodes[0]!.bounds.left).toBe(191440),
    );
    window.dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowDown', shiftKey: true }),
    );
    await waitFor(() =>
      expect(nudgeEditor.session().scene().slides[0]!.nodes[0]!.bounds.top).toBe(1014400),
    );

    expect(events).toEqual(['已移动对象 右 1步', '已移动对象 下 10步']);
    expect(nudgeEditor.canUndo()).toBe(true);

    await nudgeEditor.destroy();
  });

  test('copies, pastes, duplicates, and deletes selected shapes with keyboard shortcuts', async () => {
    const { editor, root } = await buildEditorHost();
    const events: string[] = [];
    await editor.destroy();

    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-shape-clipboard-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Draft' },
            },
          ],
        },
      ],
    });
    const clipboardEditor = renderNativePptEditor(host, view, {
      onObjectAction: (event) => events.push(event.action),
    });

    await clipboardEditor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
    });
    window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'c', ctrlKey: true }));
    window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'v', ctrlKey: true }));
    await waitFor(() => expect(clipboardEditor.session().scene().slides[0]!.nodes).toHaveLength(2));
    expect(clipboardEditor.selection()).toMatchObject({ kind: 'shape', shapeId: '3' });

    window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'd', ctrlKey: true }));
    await waitFor(() => expect(clipboardEditor.session().scene().slides[0]!.nodes).toHaveLength(3));
    expect(clipboardEditor.selection()).toMatchObject({ kind: 'shape', shapeId: '4' });

    window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Delete' }));
    await waitFor(() => expect(clipboardEditor.session().scene().slides[0]!.nodes).toHaveLength(2));
    expect(root.querySelector('.coco-ppt-node')).toBeNull();
    expect(events).toEqual(['copy-shape', 'paste-shape', 'duplicate-shape', 'delete-shape']);
    const verified = await clipboardEditor.session().verifyExport();
    expect(verified.slides[0]!.shapes).toHaveLength(2);

    await clipboardEditor.destroy();
  });

  test('opens an object context menu and runs common object actions', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const events: string[] = [];
    const view = await buildPptEditorViewFromModel({
      id: 'native-context-menu-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Back',
              kind: 'geometricShape',
              raw: { presetGeometry: 'rect' },
              geometry: { left: 100000, top: 100000, width: 1000000, height: 500000 },
            },
            {
              id: '3',
              name: 'Front',
              kind: 'textBox',
              geometry: { left: 200000, top: 200000, width: 2000000, height: 500000 },
              text: { text: 'Front' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view, {
      onObjectAction: (event) => events.push(event.action),
    });
    const root = host.shadowRoot!;
    const front = root.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="3"]')!;

    front.dispatchEvent(contextMenu(80, 90));

    await waitFor(() => {
      const menu = root.querySelector<HTMLElement>('.coco-ppt-context-menu');
      expect(menu?.dataset.open).toBe('true');
      expect(menu?.dataset.shapeId).toBe('3');
      expect(menu?.textContent).toContain('复制');
    });
    expect(editor.selection()).toMatchObject({ kind: 'shape', shapeId: '3' });

    menuButton(root, '复制').click();
    front.dispatchEvent(contextMenu(80, 90));
    await waitFor(() => {
      expect(root.querySelector<HTMLElement>('.coco-ppt-context-menu')?.dataset.shapeId).toBe('3');
      expect(menuButton(root, '粘贴').disabled).toBe(false);
    });
    menuButton(root, '粘贴').click();
    await waitFor(() => expect(editor.session().scene().slides[0]!.nodes).toHaveLength(3));
    const pastedId = editor.selection()?.kind === 'shape' ? editor.selection()!.shapeId : '';
    const pasted = () =>
      root.querySelector<HTMLElement>(`.coco-ppt-node[data-shape-id="${pastedId}"]`)!;

    pasted().dispatchEvent(contextMenu(90, 100));
    await waitFor(() => {
      expect(root.querySelector<HTMLElement>('.coco-ppt-context-menu')?.dataset.shapeId).toBe(
        pastedId,
      );
      expect(menuButton(root, '置于底层').disabled).toBe(false);
    });
    menuButton(root, '置于底层').click();
    await waitFor(() => expect(editor.session().model().slides[0]!.shapes[0]!.id).toBe(pastedId));
    pasted().dispatchEvent(contextMenu(90, 100));
    await waitFor(() => {
      expect(root.querySelector<HTMLElement>('.coco-ppt-context-menu')?.dataset.shapeId).toBe(
        pastedId,
      );
      expect(menuButton(root, '置于顶层').disabled).toBe(false);
    });
    menuButton(root, '置于顶层').click();
    await waitFor(() =>
      expect(editor.session().model().slides[0]!.shapes.at(-1)!.id).toBe(pastedId),
    );
    pasted().dispatchEvent(contextMenu(90, 100));
    await waitFor(() => {
      expect(root.querySelector<HTMLElement>('.coco-ppt-context-menu')?.dataset.shapeId).toBe(
        pastedId,
      );
      expect(menuButton(root, '删除').disabled).toBe(false);
    });
    menuButton(root, '删除').click();
    await waitFor(() => expect(editor.session().scene().slides[0]!.nodes).toHaveLength(2));

    expect(events).toEqual([
      'copy-shape',
      'paste-shape',
      'send-shape-to-back',
      'bring-shape-to-front',
      'delete-shape',
    ]);
    const verified = await editor.session().verifyExport();
    expect(verified.slides[0]!.shapes).toHaveLength(2);

    await editor.destroy();
  });

  test('does not nudge while direct text editing is active', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    await editor.dispatch({ kind: 'select-shape', slideIdOrIndex: 0, shapeIdOrName: '2' });
    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' }));
    window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' }));

    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(editor.session().scene().slides[0]!.nodes[0]!.bounds.left).toBe(100000);

    await editor.destroy();
  });

  test('keeps dragged shapes partially visible on the slide', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;

    node.dispatchEvent(mouse('mousedown', 10, 10));
    window.dispatchEvent(mouse('mousemove', -10000, -10000));
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', -10000, -10000));
    await changed;

    await waitFor(() => {
      const bounds = editor.session().scene().slides[0]!.nodes[0]!.bounds;
      expect(bounds.left).toBeGreaterThanOrEqual(-bounds.width + 240000);
      expect(bounds.top).toBeGreaterThanOrEqual(-bounds.height + 240000);
      expect(bounds.left).toBeLessThan(100000);
      expect(bounds.top).toBeLessThan(100000);
    });

    await editor.destroy();
  });

  test('does not drag or resize shapes without geometry capability', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-locked-geometry-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: 'layout:2',
              name: 'Inherited',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Inherited' },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-node[data-shape-id="layout:2"]',
    )!;
    const handle = node.querySelector<HTMLElement>('.coco-ppt-resize-handle')!;

    expect(node.dataset.resizable).toBe('false');
    node.dispatchEvent(mouse('mousedown', 10, 10));
    window.dispatchEvent(mouse('mousemove', 30, 40));
    window.dispatchEvent(mouse('mouseup', 30, 40));
    handle.dispatchEvent(mouse('mousedown', 30, 40));
    window.dispatchEvent(mouse('mousemove', 60, 80));
    window.dispatchEvent(mouse('mouseup', 60, 80));

    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(editor.session().scene().slides[0]!.nodes[0]!.bounds).toMatchObject({
      left: 100000,
      top: 100000,
      width: 2000000,
      height: 500000,
    });
    expect(editor.canUndo()).toBe(false);

    await editor.destroy();
  });

  test('keeps resize interactions above the minimum shape size', async () => {
    const { editor, root } = await buildEditorHost();
    const node = root.querySelector<HTMLElement>('.coco-ppt-node')!;
    const handle = node.querySelector<HTMLElement>('.coco-ppt-resize-handle')!;

    handle.dispatchEvent(mouse('mousedown', 10, 10));
    window.dispatchEvent(mouse('mousemove', -10000, -10000));
    const changed = waitForNextChange(editor);
    window.dispatchEvent(mouse('mouseup', -10000, -10000));
    await changed;

    await waitFor(() => {
      const bounds = editor.session().scene().slides[0]!.nodes[0]!.bounds;
      expect(bounds.width).toBeGreaterThanOrEqual(91440);
      expect(bounds.height).toBeGreaterThanOrEqual(91440);
      expect(bounds.width).toBeLessThan(2000000);
      expect(bounds.height).toBeLessThan(500000);
    });

    await editor.destroy();
  });

  test('renders distributed fallback text inside an internal wrapper', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-distributed-text-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '5',
              name: 'Distributed title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1400000, height: 300000 },
              text: {
                text: '产品运营',
                body: { autoFit: 'shape', wrap: 'square' },
                style: { alignment: 'distributed' },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textStyle = { alignment: 'distributed' };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="5"]',
    )!;
    const wrapper = node.querySelector<HTMLElement>('.coco-ppt-distributed-text')!;

    expect(node.style.justifyContent).toBe('');
    expect(wrapper).not.toBeNull();
    expect(wrapper.children).toHaveLength(4);

    await editor.destroy();
  });

  test('clips preset shape backgrounds without clipping their text', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-preset-text-clip-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '7',
              name: 'Parallelogram title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 300000 },
              style: { fillColor: '#D8F1FF' },
              text: { text: '工作回顾' },
              raw: { presetGeometry: 'parallelogram' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'parallelogram';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="7"]',
    )!;
    const shape = node.querySelector<SVGPathElement>('.coco-ppt-preset-geometry > path')!;
    const textLayer = node.querySelector<HTMLElement>('.coco-ppt-text-layer')!;

    expect(node.style.clipPath).toBe('');
    expect(shape.getAttribute('d')).toBe('M 18 0 H 100 L 82 100 H 0 Z');
    expect(textLayer.textContent).toBe('工作回顾');

    await editor.destroy();
  });

  test('keeps text visible when only the shape fill is transparent', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-transparent-fill-text-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '8',
              name: 'Transparent fill text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 400000 },
              style: { fillColor: '#FFFFFF', opacity: 0 },
              text: { text: '具体执行：文本不能被透明填充隐藏。', style: { color: '#404040' } },
            },
          ],
        },
      ],
    });

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="8"]',
    )!;

    expect(node.style.opacity).toBe('');
    expect(node.textContent).toContain('具体执行');

    await editor.destroy();
  });

  test('normalizes CSS gradient stops for rectangle fills', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-css-gradient-stop-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '23',
              name: 'Rectangle Gradient',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1000000, height: 1000000 },
              raw: { presetGeometry: 'rect' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.fillGradient = {
      colors: ['#0958E9', 'rgba(150, 185, 248, 0.5)', '#0958E9'],
      angle: 225,
      positions: [65, 0, 89],
      opacities: [1, 0.5, 1],
    };
    view.scene.slides[0]!.nodes[0]!.opacity = 0.4;

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="23"]',
    )!;
    const style = node.getAttribute('style') ?? '';

    expect(style).toContain('linear-gradient(315deg');
    expect(style.indexOf('rgba(150, 185, 248, 0.2) 0%')).toBeLessThan(style.indexOf('65%'));
    expect(style).toContain('rgba(9, 88, 233, 0.4) 65%');
    expect(style.indexOf('65%')).toBeLessThan(style.indexOf('89%'));

    await editor.destroy();
  });

  test('renders frame preset geometry as a hollow gradient border', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-frame-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '6',
              name: 'Frame',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 400000, height: 400000 },
              style: {
                fillGradient: {
                  colors: ['#C80D14', '#9F080D'],
                  angle: 225,
                  positions: [0, 100],
                },
              },
              raw: { presetGeometry: 'frame' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'frame';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="6"]',
    )!;

    const style = node.getAttribute('style') ?? '';
    expect(style).not.toContain('background: linear-gradient');
    expect(style).toContain('border-style: solid');
    expect(style).toContain('border-width: 1px');

    await editor.destroy();
  });

  test('renders round2DiagRect with diagonal rounded corners', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-round2-diag-rect-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '13',
              name: 'Diagonal rounded card',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 900000 },
              style: { fillColor: '#C80D14' },
              raw: { presetGeometry: 'round2DiagRect' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'round2DiagRect';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="13"]',
    )!;

    expect(node.style.borderTopLeftRadius).not.toBe('');
    expect(node.style.borderTopRightRadius).toBe('0px');
    expect(node.style.borderBottomRightRadius).not.toBe('');
    expect(node.style.borderBottomLeftRadius).toBe('0px');

    await editor.destroy();
  });

  test('renders round2SameRect with same-side rounded corners', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-round2-same-rect-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '14',
              name: 'Top rounded panel',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 900000 },
              style: { fillColor: '#164A96' },
              raw: { presetGeometry: 'round2SameRect' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'round2SameRect';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="14"]',
    )!;

    expect(node.style.borderTopLeftRadius).not.toBe('');
    expect(node.style.borderTopRightRadius).not.toBe('');
    expect(node.style.borderBottomRightRadius).toBe('0px');
    expect(node.style.borderBottomLeftRadius).toBe('0px');

    await editor.destroy();
  });

  test('clips snip2DiagRect preset geometry', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-snip2-diag-rect-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '12',
              name: 'Gantt bar',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 180000 },
              style: { fillColor: '#00B3F8' },
              raw: { presetGeometry: 'snip2DiagRect' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'snip2DiagRect';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="12"]',
    )!;

    const path = node.querySelector<SVGPathElement>('.coco-ppt-preset-geometry > path')!;

    expect(node.style.background).toBe('transparent');
    expect(node.style.clipPath).toBe('');
    expect(path.getAttribute('d')).toBe('M 8 0 H 100 V 92 L 92 100 H 0 V 8 Z');

    await editor.destroy();
  });

  test('renders curvedUpArrow preset geometry as svg', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-curved-up-arrow-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '13',
              name: 'Curved Arrow',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 900000 },
              style: {
                fillGradient: {
                  colors: ['#3C6DB1', '#133E82'],
                  angle: 225,
                  positions: [0, 100],
                },
                lineGradient: {
                  colors: ['rgba(255, 255, 255, 0)', '#FFFFFF'],
                  angle: 90,
                  positions: [0, 100],
                },
                lineWidth: 12700,
              },
              raw: { presetGeometry: 'curvedUpArrow', threeD: { material: 'translucentPowder' } },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'curvedUpArrow';
    view.scene.slides[0]!.nodes[0]!.fillGradient = {
      colors: ['#3C6DB1', '#133E82'],
      angle: 225,
      positions: [0, 100],
    };
    view.scene.slides[0]!.nodes[0]!.lineGradient = {
      colors: ['rgba(255, 255, 255, 0)', '#FFFFFF'],
      angle: 90,
      positions: [0, 100],
    };
    view.scene.slides[0]!.nodes[0]!.lineWidth = 12700;
    view.scene.slides[0]!.nodes[0]!.threeD = { material: 'translucentPowder' };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="13"]',
    )!;

    expect(node.getAttribute('style') ?? '').toContain('background: transparent');
    const path = node.querySelector('.coco-ppt-preset-geometry path');
    expect(path?.getAttribute('d')).toContain('C 3 61 9 28 38 13');
    expect(path?.getAttribute('d')).toContain('L 100 38');
    expect(path?.getAttribute('fill-opacity')).toBe('0.52');
    expect(path?.getAttribute('stroke')).toContain('url(#coco-ppt-preset-line-gradient-13)');
    const gradient = node.querySelector('.coco-ppt-preset-geometry linearGradient')!;
    expect(gradient).not.toBeNull();
    expect(gradient.getAttribute('x1')).toBe('100%');
    expect(gradient.getAttribute('y1')).toBe('100%');
    expect(gradient.getAttribute('x2')).toBe('0%');
    expect(gradient.getAttribute('y2')).toBe('0%');

    await editor.destroy();
  });

  test('renders preset geometry extrusion as a light SVG back layer', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-preset-depth-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '15',
              name: 'Depth Panel',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1800000, height: 700000 },
              style: {
                fillColor: '#56B6EE',
                shadow: {
                  color: '#93AACA',
                  opacity: 0.71,
                  blur: 723900,
                  distance: 482600,
                  direction: 120,
                },
              },
              raw: {
                presetGeometry: 'trapezoid',
                threeD: { material: 'translucentPowder', extrusionHeight: 127000 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'trapezoid';
    view.scene.slides[0]!.nodes[0]!.fillColor = '#56B6EE';
    view.scene.slides[0]!.nodes[0]!.shadow = {
      color: '#93AACA',
      opacity: 0.71,
      blur: 723900,
      distance: 482600,
      direction: 120,
    };
    view.scene.slides[0]!.nodes[0]!.threeD = {
      material: 'translucentPowder',
      extrusionHeight: 127000,
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="15"]',
    )!;
    const svg = node.querySelector<SVGSVGElement>('.coco-ppt-preset-geometry')!;
    const extrusion = svg.querySelector<SVGPathElement>('.coco-ppt-preset-extrusion')!;
    const path = svg.querySelector<SVGPathElement>('path:not(.coco-ppt-preset-extrusion)')!;

    expect(svg.style.overflow).toBe('visible');
    expect(extrusion.getAttribute('d')).toBe(path.getAttribute('d'));
    expect(extrusion.getAttribute('transform')).toContain('translate(');
    expect(path.style.filter).toBe('');
    expect(path.getAttribute('fill-opacity')).toBe('0.52');

    await editor.destroy();
  });

  test('renders donut preset geometry as an even-odd SVG ring', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-donut-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '14',
              name: 'Donut',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 700000, height: 700000 },
              style: { fillColor: '#62D4F3' },
              raw: { presetGeometry: 'donut' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'donut';
    view.scene.slides[0]!.nodes[0]!.fillColor = '#62D4F3';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="14"]',
    )!;
    const path = node.querySelector<SVGPathElement>('.coco-ppt-preset-geometry > path')!;

    expect(node.getAttribute('style') ?? '').toContain('background: transparent');
    expect(path).not.toBeNull();
    expect(path.getAttribute('fill-rule')).toBe('evenodd');
    expect(path.getAttribute('d')).toContain('A 48 48');

    await editor.destroy();
  });

  test('renders homePlate preset geometry as an SVG pentagon', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-home-plate-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '19',
              name: 'Home Plate',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 900000, height: 700000 },
              style: { fillColor: '#254DE1' },
              raw: { presetGeometry: 'homePlate' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'homePlate';
    view.scene.slides[0]!.nodes[0]!.fillColor = '#254DE1';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="19"]',
    )!;
    const path = node.querySelector<SVGPathElement>('.coco-ppt-preset-geometry > path')!;

    expect(node.style.background).toBe('transparent');
    expect(node.style.clipPath).toBe('');
    expect(path.getAttribute('d')).toBe('M 50 0 L 100 25 V 100 H 0 V 25 Z');
    expect(path.getAttribute('fill')).toBe('#254DE1');

    await editor.destroy();
  });

  test('renders hexagon preset geometry as an SVG path with shape stroke', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-hexagon-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '11',
              name: 'Hexagon',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 900000, height: 900000 },
              style: { fillColor: '#164A96', lineColor: '#F59E0B', lineWidth: 25400 },
              raw: { presetGeometry: 'hexagon' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'hexagon';
    view.scene.slides[0]!.nodes[0]!.fillColor = '#164A96';
    view.scene.slides[0]!.nodes[0]!.lineColor = '#F59E0B';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="11"]',
    )!;
    const path = node.querySelector<SVGPathElement>('.coco-ppt-preset-geometry > path')!;

    expect(node.style.background).toBe('transparent');
    expect(node.style.clipPath).toBe('');
    expect(node.style.borderRadius).toBe('');
    expect(path.getAttribute('d')).toBe('M 25 0 H 75 L 100 50 L 75 100 H 25 L 0 50 Z');
    expect(path.getAttribute('fill')).toBe('#164A96');
    expect(path.getAttribute('stroke')).toBe('#F59E0B');

    await editor.destroy();
  });

  test('renders chevron preset geometry as an SVG path', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-chevron-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '12',
              name: 'Chevron',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 200000, height: 300000 },
              style: { fillColor: '#00B3F8' },
              raw: { presetGeometry: 'chevron' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'chevron';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="12"]',
    )!;
    const path = node.querySelector<SVGPathElement>('.coco-ppt-preset-geometry > path')!;

    expect(node.style.background).toBe('transparent');
    expect(node.style.clipPath).toBe('');
    expect(node.style.borderRadius).toBe('');
    expect(path.getAttribute('d')).toBe('M 0 0 H 72 L 100 50 L 72 100 H 0 L 28 50 Z');

    await editor.destroy();
  });

  test('renders plaque preset geometry with inset curved corners', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-plaque-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '13',
              name: 'Plaque',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 2200000, height: 400000 },
              style: { fillColor: '#21B9E7' },
              raw: { presetGeometry: 'plaque' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'plaque';
    view.scene.slides[0]!.nodes[0]!.fillColor = '#21B9E7';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="13"]',
    )!;
    const path = node.querySelector<SVGPathElement>('svg.coco-ppt-preset-geometry path')!;

    expect(node.style.background).toBe('transparent');
    expect(path.getAttribute('d')).toContain('Q 96 8 100 8');
    expect(path.getAttribute('fill')).toBe('#21B9E7');

    await editor.destroy();
  });

  test('renders arc preset geometry as an open SVG path with arrowhead', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-arc-preset-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '20',
              name: 'Arc',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 900000, height: 900000 },
              style: { lineColor: '#5B9BCB', lineHeadEnd: 'triangle' },
              raw: { presetGeometry: 'arc', presetAdjustments: { adj1: 14039509, adj2: 10149582 } },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'arc';
    view.scene.slides[0]!.nodes[0]!.presetAdjustments = { adj1: 14039509, adj2: 10149582 };
    view.scene.slides[0]!.nodes[0]!.lineColor = '#5B9BCB';
    view.scene.slides[0]!.nodes[0]!.lineWidth = undefined;
    view.scene.slides[0]!.nodes[0]!.lineHeadEnd = 'triangle';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="20"]',
    )!;
    const path = node.querySelector<SVGPathElement>('.coco-ppt-preset-geometry > path')!;

    expect(path.getAttribute('fill')).toBe('none');
    expect(path.getAttribute('d')).toContain('A 48 48');
    expect(Number(path.getAttribute('stroke-width'))).toBeGreaterThan(0);
    expect(path.getAttribute('marker-start') ?? '').toContain('url(#');
    expect(node.getAttribute('style') ?? '').toContain('background: transparent');
    expect(node.getAttribute('style') ?? '').not.toContain('border-style: solid');

    await editor.destroy();
  });

  test('renders blockArc preset geometry as a filled ring segment', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-block-arc-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '21',
              name: 'Progress arc',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 700000, height: 700000 },
              style: { fillColor: '#133E82' },
              raw: {
                presetGeometry: 'blockArc',
                presetAdjustments: { adj1: 10708716, adj2: 1500102, adj3: 17978 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'blockArc';
    view.scene.slides[0]!.nodes[0]!.presetAdjustments = {
      adj1: 10708716,
      adj2: 1500102,
      adj3: 17978,
    };
    view.scene.slides[0]!.nodes[0]!.fillColor = '#133E82';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="21"]',
    )!;
    const path = node.querySelector<SVGPathElement>('.coco-ppt-preset-geometry > path')!;

    expect(path.getAttribute('d')).toContain('A 48 48');
    expect(path.getAttribute('d')).toContain('Z');
    expect(path.getAttribute('fill')).toBe('#133E82');
    expect(path.getAttribute('fill-rule')).toBe('evenodd');
    expect(node.getAttribute('style') ?? '').toContain('background: transparent');
    expect(node.getAttribute('style') ?? '').not.toContain('border-radius: 9999px');

    await editor.destroy();
  });

  test('renders dashed preset shape outlines', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-dashed-outline-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '14',
              name: 'Dashed Box',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 900000, height: 500000 },
              style: { fillColor: '#FFFFFF', lineColor: '#123456', lineWidth: 9525 },
              raw: { presetGeometry: 'roundRect' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'roundRect';
    view.scene.slides[0]!.nodes[0]!.lineColor = '#123456';
    view.scene.slides[0]!.nodes[0]!.lineWidth = 9525;
    view.scene.slides[0]!.nodes[0]!.lineDash = 'dash';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="14"]',
    )!;

    expect(node.style.borderStyle).toBe('dashed');

    await editor.destroy();
  });

  test('does not round custom filled geometry through circle heuristics', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-custom-geometry-no-circle-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '15',
              name: 'Wave',
              kind: 'geometricShape',
              geometry: { left: 0, top: 0, width: 1000000, height: 1000000 },
              style: { fillColor: '#00C9D4' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.fillColor = '#00C9D4';
    view.scene.slides[0]!.nodes[0]!.customGeometry = {
      paths: [
        {
          width: 100000,
          height: 100000,
          commands: [
            { kind: 'moveTo', x: 0, y: 52000 },
            {
              kind: 'cubicBezierTo',
              x1: 22000,
              y1: 36000,
              x2: 44000,
              y2: 66000,
              x: 64000,
              y: 46000,
            },
            { kind: 'lineTo', x: 100000, y: 100000 },
            { kind: 'lineTo', x: 0, y: 100000 },
            { kind: 'close' },
          ],
        },
      ],
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="15"]',
    )!;

    expect(node.style.borderRadius).toBe('');
    expect(node.querySelector('.coco-ppt-custom-geometry path')).not.toBeNull();

    await editor.destroy();
  });

  test('renders ellipse-like cubic custom geometry as an SVG ellipse', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-custom-geometry-ellipse-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '16',
              name: 'Bezier circle',
              kind: 'geometricShape',
              geometry: { left: 0, top: 0, width: 1000000, height: 1000000 },
              style: { lineColor: '#133E82', lineWidth: 101600 },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.lineColor = '#133E82';
    view.scene.slides[0]!.nodes[0]!.lineWidth = 101600;
    view.scene.slides[0]!.nodes[0]!.customGeometry = {
      paths: [
        {
          width: 1000,
          height: 1000,
          commands: [
            { kind: 'moveTo', x: 1000, y: 500 },
            { kind: 'cubicBezierTo', x1: 1000, y1: 776, x2: 776, y2: 1000, x: 500, y: 1000 },
            { kind: 'cubicBezierTo', x1: 224, y1: 1000, x2: 0, y2: 776, x: 0, y: 500 },
            { kind: 'cubicBezierTo', x1: 0, y1: 224, x2: 224, y2: 0, x: 500, y: 0 },
            { kind: 'cubicBezierTo', x1: 776, y1: 0, x2: 1000, y2: 224, x: 1000, y: 500 },
            { kind: 'close' },
          ],
        },
      ],
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="16"]',
    )!;

    const ellipse = node.querySelector<SVGEllipseElement>('.coco-ppt-custom-geometry ellipse')!;
    expect(ellipse).not.toBeNull();
    expect(ellipse.getAttribute('rx')).toBe('500');
    expect(ellipse.getAttribute('ry')).toBe('500');
    expect(node.querySelector('.coco-ppt-custom-geometry path')).toBeNull();

    await editor.destroy();
  });

  test('renders custom geometry arc commands as svg arcs', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-custom-geometry-arc-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '18',
              name: 'ArcShape',
              kind: 'geometricShape',
              geometry: { left: 0, top: 0, width: 1000000, height: 1000000 },
              style: { lineColor: '#C43E1C', lineWidth: 63500 },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.customGeometry = {
      paths: [
        {
          width: 1000,
          height: 1000,
          commands: [
            { kind: 'moveTo', x: 1000, y: 500 },
            { kind: 'arcTo', wR: 500, hR: 500, stAng: 0, swAng: 5400000 },
          ],
        },
      ],
    };

    const editor = renderNativePptEditor(host, view);
    const path = host.shadowRoot!.querySelector<SVGPathElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="18"] .coco-ppt-custom-geometry path',
    )!;

    expect(path.getAttribute('d')).toBe('M 1000 500 A 500 500 0 0 1 500 1000');

    await editor.destroy();
  });

  test('does not round filled text boxes through circle heuristics', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-textbox-no-circle-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '16',
              name: 'Image Placeholder',
              kind: 'textBox',
              geometry: { left: 0, top: 0, width: 1000000, height: 920000 },
              style: { fillColor: '#AAAAAA' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.kind = 'textBox';
    view.scene.slides[0]!.nodes[0]!.fillColor = '#AAAAAA';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="16"]',
    )!;

    expect(node.style.borderRadius).toBe('');

    await editor.destroy();
  });

  test('renders dashed custom line geometry as stroke without fill', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-custom-dashed-line-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '17',
              name: 'Dashed Connector',
              kind: 'line',
              geometry: { left: 0, top: 0, width: 2600000, height: 330000 },
              style: { fillColor: '#00B3F8', lineColor: '#00B3F8', lineWidth: 15875 },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.kind = 'line';
    view.scene.slides[0]!.nodes[0]!.fillColor = '#00B3F8';
    view.scene.slides[0]!.nodes[0]!.lineColor = '#00B3F8';
    view.scene.slides[0]!.nodes[0]!.lineWidth = 15875;
    view.scene.slides[0]!.nodes[0]!.lineDash = 'cust:1500038:1500038';
    view.scene.slides[0]!.nodes[0]!.customGeometry = {
      paths: [
        {
          width: 21600,
          height: 21600,
          commands: [
            { kind: 'moveTo', x: 0, y: 1000 },
            { kind: 'lineTo', x: 21600, y: 1000 },
            { kind: 'lineTo', x: 21600, y: 21600 },
            { kind: 'lineTo', x: 0, y: 21600 },
            { kind: 'close' },
          ],
        },
      ],
    };

    const editor = renderNativePptEditor(host, view);
    const path = host.shadowRoot!.querySelector<SVGPathElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="17"] .coco-ppt-custom-geometry path',
    )!;

    expect(path.getAttribute('fill')).toBe('none');
    expect(Number(path.getAttribute('stroke-width'))).toBeLessThan(2000);
    expect(path.getAttribute('stroke-dasharray')).toMatch(/\d+(\.\d+)? \d+(\.\d+)?/);

    await editor.destroy();
  });

  test('renders near-axis connector shapes through line geometry', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-connector-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '7',
              name: 'Thin Connector',
              kind: 'line',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 10000 },
              style: {
                lineColor: '#FFFFFF',
                lineGradient: {
                  colors: ['#00B3F8', 'rgba(204, 240, 254, 0.5)'],
                  angle: 90,
                  positions: [0, 100],
                },
                lineWidth: 76200,
                lineCap: 'flat',
                lineTailEnd: 'arrow',
              },
              raw: { presetGeometry: 'straightConnector1' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'straightConnector1';
    view.scene.slides[0]!.nodes[0]!.lineGradient = {
      colors: ['#00B3F8', 'rgba(204, 240, 254, 0.5)'],
      angle: 90,
      positions: [0, 100],
    };
    view.scene.slides[0]!.nodes[0]!.lineTailEnd = 'arrow';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="7"]',
    )!;

    expect(node.querySelector('.coco-ppt-line-geometry')).not.toBeNull();
    expect(node.querySelector('.coco-ppt-line-geometry linearGradient')).not.toBeNull();
    const line = node.querySelector('.coco-ppt-line-geometry line')!;
    const arrow = node.querySelector('.coco-ppt-line-arrow')!;
    expect(line.getAttribute('stroke') ?? '').toContain('url(#');
    expect(line.getAttribute('stroke-linecap')).toBe('butt');
    expect(arrow.getAttribute('points') ?? '').toBe('1,6 -4,3.5 -4,8.5');
    expect(node.getAttribute('style') ?? '').toContain('background: transparent');

    await editor.destroy();
  });

  test('renders roundRect line gradients as an SVG outline', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-gradient-outline-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '8',
              name: 'Gradient Outline',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 800000, height: 500000 },
              style: {
                fillColor: 'rgba(255,255,255,0.24)',
                lineGradient: {
                  colors: ['#22D3EE', 'rgba(34,211,238,0)'],
                  angle: 90,
                  positions: [0, 100],
                },
              },
              raw: { presetGeometry: 'roundRect' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'roundRect';
    view.scene.slides[0]!.nodes[0]!.lineColor = undefined;
    view.scene.slides[0]!.nodes[0]!.lineGradient = {
      colors: ['#22D3EE', 'rgba(34,211,238,0)'],
      angle: 90,
      positions: [0, 100],
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="8"]',
    )!;

    expect(node.querySelector('.coco-ppt-gradient-outline')).not.toBeNull();
    expect(node.getAttribute('style') ?? '').not.toContain('border-style: solid');

    await editor.destroy();
  });

  test('renders ellipse line gradients as an SVG outline', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-ellipse-gradient-outline-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '80',
              name: 'Gradient Ellipse',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 100000, width: 1200000, height: 500000 },
              style: {
                lineGradient: {
                  colors: ['rgba(37,77,225,0)', '#254DE1'],
                  angle: 90,
                  positions: [18, 100],
                },
                lineWidth: 19050,
              },
              raw: { presetGeometry: 'ellipse' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'ellipse';
    view.scene.slides[0]!.nodes[0]!.lineColor = undefined;
    view.scene.slides[0]!.nodes[0]!.lineGradient = {
      colors: ['rgba(37,77,225,0)', '#254DE1'],
      angle: 90,
      positions: [18, 100],
    };
    view.scene.slides[0]!.nodes[0]!.lineWidth = 19050;

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="80"]',
    )!;
    const ellipse = node.querySelector<SVGElement>('.coco-ppt-gradient-outline ellipse')!;

    expect(ellipse).not.toBeNull();
    expect(ellipse.getAttribute('stroke') ?? '').toContain('url(#');
    expect(node.getAttribute('style') ?? '').not.toContain('border-style: solid');

    await editor.destroy();
  });

  test('allows compact centered CJK labels to wrap inside auto-fit text boxes', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-compact-label-wrap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '81',
              name: 'Circle Label Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1020234, height: 861695 },
              text: {
                text: '业务基础',
                body: { autoFit: 'shape', wrap: 'square' },
                style: {
                  fontSize: 12,
                  color: '#FFFFFF',
                  alignment: 'center',
                },
              },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="81"]',
    )!;

    expect(node.style.whiteSpace).not.toBe('nowrap');

    await editor.destroy();
  });

  test('does not shrink text for shape auto-fit text boxes', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-shape-autofit-no-scale-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '82',
              name: 'Shape AutoFit Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 380000, height: 260000 },
              text: {
                text: '未来',
                body: { autoFit: 'shape', wrap: 'square' },
                style: { fontSize: 44 },
              },
            },
          ],
        },
      ],
    });
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="82"]',
    )!;

    expect(node.style.transform).not.toContain('scale(');
    expect(node.style.whiteSpace).not.toBe('nowrap');

    await editor.destroy();
  });

  test('adds a narrow metric fallback for long uppercase Josefin Sans titles', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-latin-display-title-metric-fallback-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '83',
              name: 'Long Latin Title',
              kind: 'textBox',
              geometry: { left: 0, top: 0, width: 6096000, height: 1800000 },
              text: {
                text: 'HOW TO USE THIS PRESENTATION',
                body: { autoFit: 'shape', wrap: 'square' },
                style: {
                  fontSize: 65,
                  fontFamily: 'Josefin Sans',
                  alignment: 'center',
                  bold: true,
                },
              },
            },
            {
              id: '84',
              name: 'Explicit Tracking Title',
              kind: 'textBox',
              geometry: { left: 0, top: 1900000, width: 6096000, height: 1800000 },
              text: {
                text: 'HOW TO USE THIS PRESENTATION',
                body: { autoFit: 'shape', wrap: 'square' },
                style: {
                  fontSize: 65,
                  fontFamily: 'Josefin Sans',
                  alignment: 'center',
                  characterSpacing: 1.5,
                },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = { autoFit: 'shape', wrap: 'square' };
    view.scene.slides[0]!.nodes[0]!.textStyle = {
      fontSize: 65,
      fontFamily: 'Josefin Sans',
      alignment: 'center',
      bold: true,
    };
    view.scene.slides[0]!.nodes[1]!.textBody = { autoFit: 'shape', wrap: 'square' };
    view.scene.slides[0]!.nodes[1]!.textStyle = {
      fontSize: 65,
      fontFamily: 'Josefin Sans',
      alignment: 'center',
      characterSpacing: 1.5,
    };
    const editor = renderNativePptEditor(host, view);
    const title = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="83"]',
    )!;
    const explicit = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="84"]',
    )!;

    expect(title.style.letterSpacing).not.toBe('');
    expect(explicit.style.letterSpacing).not.toBe(title.style.letterSpacing);

    await editor.destroy();
  });

  test('lowers large Cuprum display titles without changing ordinary text boxes', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-cuprum-display-title-baseline-test',
      slideSize: { width: 12192000, height: 6858000 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '124',
              name: 'Large Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 6000000, height: 1400000 },
              text: {
                text: 'How to Use This Presentation',
                body: {
                  anchor: 'top',
                  autoFit: 'shape',
                  wrap: 'square',
                  insetLeft: 0,
                  insetTop: 0,
                  insetRight: 0,
                  insetBottom: 0,
                },
                style: {
                  fontSize: 90,
                  fontFamily: 'Cuprum',
                  color: '#FFFFFF',
                  alignment: 'left',
                },
              },
            },
            {
              id: '125',
              name: 'Body Copy',
              kind: 'textBox',
              geometry: { left: 100000, top: 1800000, width: 3000000, height: 1400000 },
              text: {
                text: 'Click on the button below this presentation preview.',
                body: {
                  anchor: 'top',
                  autoFit: 'shape',
                  wrap: 'square',
                  insetLeft: 0,
                  insetTop: 0,
                  insetRight: 0,
                  insetBottom: 0,
                },
                style: {
                  fontSize: 19,
                  fontFamily: 'Arial',
                  color: '#FFFFFF',
                  alignment: 'left',
                },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = {
      anchor: 'top',
      autoFit: 'shape',
      wrap: 'square',
      insetLeft: 0,
      insetTop: 0,
      insetRight: 0,
      insetBottom: 0,
    };
    view.scene.slides[0]!.nodes[0]!.textStyle = {
      fontSize: 90,
      fontFamily: 'Cuprum',
      color: '#FFFFFF',
      alignment: 'left',
    };
    view.scene.slides[0]!.nodes[1]!.textBody = {
      anchor: 'top',
      autoFit: 'shape',
      wrap: 'square',
      insetLeft: 0,
      insetTop: 0,
      insetRight: 0,
      insetBottom: 0,
    };
    view.scene.slides[0]!.nodes[1]!.textStyle = {
      fontSize: 19,
      fontFamily: 'Arial',
      color: '#FFFFFF',
      alignment: 'left',
    };
    const editor = renderNativePptEditor(host, view);
    const title = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="124"]',
    )!;
    const body = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="125"]',
    )!;

    expect(Number.parseFloat(title.style.paddingTop)).toBeGreaterThan(
      Number.parseFloat(body.style.paddingTop) + 0.01,
    );

    await editor.destroy();
  });

  test('does not leak numbered rich-text prefix styling into CJK title runs', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-numbered-title-run-style-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '83',
              name: 'Numbered Rich Title',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 400000 },
              text: {
                text: '04、请输入关键字',
                body: { autoFit: 'shape', wrap: 'square' },
                style: {
                  fontSize: 28,
                  color: '#56B6EE',
                  alignment: 'left',
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.textBody = { autoFit: 'shape', wrap: 'square' };
    nodeModel.textParagraphs = [
      {
        text: '04、请输入关键字',
        runs: [
          {
            text: '04',
            style: { fontSize: 28, color: '#56B6EE' },
          },
          {
            text: '、',
            style: { fontSize: 28, color: '#56B6EE' },
          },
          {
            text: '请',
            style: { color: '#93AACA' },
          },
          {
            text: '输入关键字',
            style: { fontFamily: '思源黑体 CN Bold' },
          },
        ],
      },
    ];
    const editor = renderNativePptEditor(host, view);
    const spans = host.shadowRoot!.querySelectorAll<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="83"] .coco-ppt-paragraph span',
    );
    const cjk = spans[3]!;

    expect(cjk.textContent).toBe('输入关键字');
    expect(cjk.style.color).toBe('#000000');
    expect(cjk.style.fontSize).not.toBe('');

    await editor.destroy();
  });

  test('keeps flipped text boxes readable by counter-flipping the text layer', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-flipped-text-readable-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '83',
              name: 'Flipped Text',
              kind: 'textBox',
              geometry: {
                left: 100000,
                top: 100000,
                width: 1200000,
                height: 300000,
                flipHorizontal: true,
              },
              text: {
                text: '单击输入标题',
                body: { autoFit: 'shape', wrap: 'square' },
                style: { fontSize: 28, alignment: 'right' },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.bounds.flipHorizontal = true;
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="83"]',
    )!;
    const textLayer = node.querySelector<HTMLElement>('.coco-ppt-text-layer')!;

    expect(node.style.transform).toContain('scaleX(-1)');
    expect(textLayer).not.toBeNull();
    expect(textLayer.textContent).toBe('单击输入标题');
    expect(textLayer.style.transform).toBe('scaleX(-1)');

    await editor.destroy();
  });

  test('shrinks overflowing text only for normal auto-fit text boxes', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const descriptors = {
      clientWidth: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth'),
      clientHeight: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight'),
      scrollWidth: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'scrollWidth'),
      scrollHeight: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'scrollHeight'),
    };
    Object.defineProperties(HTMLElement.prototype, {
      clientWidth: { configurable: true, value: 40 },
      clientHeight: { configurable: true, value: 40 },
      scrollWidth: { configurable: true, value: 120 },
      scrollHeight: { configurable: true, value: 60 },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-normal-autofit-scale-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '82',
              name: 'Normal AutoFit Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 380000, height: 260000 },
              text: {
                text: '未来',
                body: { autoFit: 'normal', wrap: 'square' },
                style: { fontSize: 44 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = { autoFit: 'normal', wrap: 'square' };
    const editor = renderNativePptEditor(host, view);
    try {
      const node = host.shadowRoot!.querySelector<HTMLElement>(
        '.coco-ppt-canvas .coco-ppt-node[data-shape-id="82"]',
      )!;

      await new Promise((resolve) => requestAnimationFrame(resolve));

      expect(node.style.transform).toBe('scale(0.3333333333333333)');
    } finally {
      await editor.destroy();
      for (const [key, descriptor] of Object.entries(descriptors)) {
        if (descriptor) Object.defineProperty(HTMLElement.prototype, key, descriptor);
        else delete (HTMLElement.prototype as unknown as Record<string, unknown>)[key];
      }
    }
  });

  test('does not wrap shape auto-fit text when the text body disables wrapping', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-shape-autofit-nowrap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '21',
              name: 'Percent',
              kind: 'textBox',
              geometry: {
                left: 556052,
                top: 5228769,
                width: 1577676,
                height: 923330,
                flipHorizontal: true,
              },
              text: {
                text: '80%',
                body: { autoFit: 'shape', wrap: 'none' },
                style: { fontSize: 54, color: '#FFFFFF', bold: true },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.textBody = { autoFit: 'shape', wrap: 'none' };
    nodeModel.bounds.flipHorizontal = true;

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="21"]',
    )!;
    const textLayer = node.querySelector<HTMLElement>('.coco-ppt-text-layer')!;

    expect(textLayer.textContent).toBe('80%');
    expect(textLayer.style.whiteSpace).toBe('nowrap');
    expect(node.style.transform).toContain('scaleX(-1)');
    expect(textLayer.style.transform).toBe('scaleX(-1)');

    await editor.destroy();
  });

  test('keeps text size and wraps after manually resizing normal auto-fit text boxes', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const descriptors = {
      clientWidth: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth'),
      clientHeight: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight'),
      scrollWidth: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'scrollWidth'),
      scrollHeight: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'scrollHeight'),
    };
    Object.defineProperties(HTMLElement.prototype, {
      clientWidth: {
        configurable: true,
        get() {
          return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage')
            ? 1200
            : 40;
        },
      },
      clientHeight: {
        configurable: true,
        get() {
          return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage')
            ? 720
            : 40;
        },
      },
      scrollWidth: { configurable: true, value: 120 },
      scrollHeight: { configurable: true, value: 60 },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-normal-autofit-manual-resize-wrap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '82',
              name: 'Normal AutoFit Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 380000, height: 260000 },
              text: {
                text: '未来可期',
                body: { autoFit: 'normal', wrap: 'square' },
                style: { fontSize: 44 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = { autoFit: 'normal', wrap: 'square' };
    const editor = renderNativePptEditor(host, view);
    try {
      let node = host.shadowRoot!.querySelector<HTMLElement>(
        '.coco-ppt-canvas .coco-ppt-node[data-shape-id="82"]',
      )!;
      await new Promise((resolve) => requestAnimationFrame(resolve));

      expect(node.style.transform).toBe('scale(0.3333333333333333)');

      const changed = waitForNextChange(editor);
      const handle = node.querySelector<HTMLElement>('.coco-ppt-resize-handle')!;
      handle.dispatchEvent(mouse('mousedown', 20, 20));
      window.dispatchEvent(mouse('mousemove', 70, 20));
      window.dispatchEvent(mouse('mouseup', 70, 20));
      await changed;

      await waitFor(() =>
        expect(editor.session().scene().slides[0]!.nodes[0]!.textBody?.autoFit).toBe('none'),
      );
      expect(editor.session().scene().slides[0]!.nodes[0]!.textBody?.wrap).toBe('square');
      node = host.shadowRoot!.querySelector<HTMLElement>(
        '.coco-ppt-canvas .coco-ppt-node[data-shape-id="82"]',
      )!;
      await new Promise((resolve) => requestAnimationFrame(resolve));
      expect(node.style.transform).not.toContain('scale(');
    } finally {
      await editor.destroy();
      for (const [key, descriptor] of Object.entries(descriptors)) {
        if (descriptor) Object.defineProperty(HTMLElement.prototype, key, descriptor);
        else delete (HTMLElement.prototype as unknown as Record<string, unknown>)[key];
      }
    }
  });

  test('keeps placeholder text size and wraps after manually resizing normal auto-fit titles', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const descriptors = {
      clientWidth: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth'),
      clientHeight: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight'),
      scrollWidth: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'scrollWidth'),
      scrollHeight: Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'scrollHeight'),
    };
    Object.defineProperties(HTMLElement.prototype, {
      clientWidth: {
        configurable: true,
        get() {
          return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage')
            ? 1200
            : 40;
        },
      },
      clientHeight: {
        configurable: true,
        get() {
          return this instanceof HTMLElement && this.classList.contains('coco-ppt-stage')
            ? 720
            : 40;
        },
      },
      scrollWidth: { configurable: true, value: 120 },
      scrollHeight: { configurable: true, value: 60 },
    });
    const view = await buildPptEditorViewFromModel({
      id: 'native-normal-autofit-placeholder-resize-wrap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '82',
              name: 'Normal AutoFit Placeholder',
              kind: 'placeholder',
              placeholder: { type: 'title' },
              geometry: { left: 100000, top: 100000, width: 380000, height: 260000 },
              text: {
                text: '未来可期',
                body: { autoFit: 'normal', wrap: 'square' },
                style: { fontSize: 44 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = { autoFit: 'normal', wrap: 'square' };
    const editor = renderNativePptEditor(host, view);
    try {
      let node = host.shadowRoot!.querySelector<HTMLElement>(
        '.coco-ppt-canvas .coco-ppt-node[data-shape-id="82"]',
      )!;
      await new Promise((resolve) => requestAnimationFrame(resolve));

      expect(node.style.transform).toBe('scale(0.3333333333333333)');

      const changed = waitForNextChange(editor);
      const handle = node.querySelector<HTMLElement>('.coco-ppt-resize-handle')!;
      handle.dispatchEvent(mouse('mousedown', 20, 20));
      window.dispatchEvent(mouse('mousemove', 70, 20));
      window.dispatchEvent(mouse('mouseup', 70, 20));
      await changed;

      await waitFor(() =>
        expect(editor.session().scene().slides[0]!.nodes[0]!.textBody?.autoFit).toBe('none'),
      );
      expect(editor.session().scene().slides[0]!.nodes[0]!.textBody?.wrap).toBe('square');
      node = host.shadowRoot!.querySelector<HTMLElement>(
        '.coco-ppt-canvas .coco-ppt-node[data-shape-id="82"]',
      )!;
      await new Promise((resolve) => requestAnimationFrame(resolve));
      expect(node.style.transform).not.toContain('scale(');
    } finally {
      await editor.destroy();
      for (const [key, descriptor] of Object.entries(descriptors)) {
        if (descriptor) Object.defineProperty(HTMLElement.prototype, key, descriptor);
        else delete (HTMLElement.prototype as unknown as Record<string, unknown>)[key];
      }
    }
  });

  test('does not clip glyphs for plain text boxes that are shorter than their font', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-text-overflow-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '83',
              name: 'Short Number',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 400000, height: 120000 },
              text: {
                text: '01',
                body: {
                  anchor: 'middle',
                  insetLeft: 0,
                  insetTop: 0,
                  insetRight: 0,
                  insetBottom: 0,
                },
                style: { fontSize: 30, alignment: 'center' },
              },
            },
            {
              id: '84',
              name: 'Filled Label',
              kind: 'textBox',
              geometry: { left: 600000, top: 100000, width: 400000, height: 120000 },
              text: {
                text: '02',
                body: {
                  anchor: 'middle',
                  insetLeft: 0,
                  insetTop: 0,
                  insetRight: 0,
                  insetBottom: 0,
                },
                style: { fontSize: 30, alignment: 'center' },
              },
              style: { fillColor: '#FFFFFF' },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[1]!.fillColor = '#FFFFFF';

    const editor = renderNativePptEditor(host, view);
    const textBox = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="83"]',
    )!;
    const filledShape = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="84"]',
    )!;

    expect(textBox.style.overflow).toBe('visible');
    expect(filledShape.style.overflow).toBe('');

    await editor.destroy();
  });

  test('allows long wrapped auto-fit body text to keep wrapping', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-long-autofit-wrap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '85',
              name: 'Body Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 800000 },
              text: {
                text: '您的内容打在这里或者通过复制您的文本后在此框中选择粘贴并选择只保留文字',
                body: { autoFit: 'shape', wrap: 'square' },
                style: { fontSize: 12 },
              },
            },
            {
              id: '86',
              name: 'Short Label',
              kind: 'textBox',
              geometry: { left: 1900000, top: 100000, width: 1600000, height: 300000 },
              text: {
                text: 'Project progress',
                body: { autoFit: 'shape', wrap: 'square' },
                style: { fontSize: 18 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = { autoFit: 'shape', wrap: 'square' };
    view.scene.slides[0]!.nodes[1]!.textBody = { autoFit: 'shape', wrap: 'square' };

    const editor = renderNativePptEditor(host, view);
    const bodyText = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="85"]',
    )!;
    const label = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="86"]',
    )!;

    expect(bodyText.style.whiteSpace).not.toBe('nowrap');
    expect(label.style.whiteSpace).not.toBe('nowrap');

    await editor.destroy();
  });

  test('renders east Asian vertical text bodies with vertical writing mode', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-vertical-text-body-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '87',
              name: 'Vertical Side Text',
              kind: 'textBox',
              geometry: { left: 8600000, top: 1000000, width: 240000, height: 4000000 },
              text: {
                text: 'WORK SUMMARY',
                body: { autoFit: 'shape', wrap: 'square', vertical: 'eastAsian' },
                style: { fontSize: 12, alignment: 'distributed', characterSpacing: 4 },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.textBody = { autoFit: 'shape', wrap: 'square', vertical: 'eastAsian' };
    nodeModel.textParagraphs = [
      {
        text: 'WORK SUMMARY',
        style: { alignment: 'distributed', spaceAfter: 76200 },
        runs: [
          {
            text: 'WORK SUMMARY',
            style: { fontSize: 12, alignment: 'distributed', characterSpacing: 4 },
          },
        ],
      },
      {
        text: 'NEXT COLUMN',
        style: { spaceBefore: 38100 },
        runs: [{ text: 'NEXT COLUMN', style: { fontSize: 12 } }],
      },
    ];

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="87"]',
    )!;
    const paragraphs = node.querySelectorAll<HTMLElement>('.coco-ppt-paragraph');
    const paragraph = paragraphs[0]!;

    expect(node.style.writingMode).toBe('vertical-rl');
    expect(node.style.textOrientation).toBe('upright');
    expect(paragraph.style.display).not.toBe('flex');
    expect(paragraph.style.marginLeft).not.toBe('');
    expect(paragraph.style.marginBottom).toBe('');
    expect(paragraphs[1]!.style.marginRight).not.toBe('');
    expect(paragraphs[1]!.style.marginTop).toBe('');

    await editor.destroy();
  });

  test('renders PPT text gradients with clipped text backgrounds', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-text-gradient-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '88',
              name: 'Gradient Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1600000, height: 800000 },
              text: {
                text: '可',
                body: { autoFit: 'shape', wrap: 'square' },
                style: {
                  fontSize: 80,
                  color: '#3878D0',
                  textGradient: {
                    colors: ['#3878D0', '#56B6EE'],
                    positions: [0, 64],
                    angle: 180,
                  },
                },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textStyle = {
      fontSize: 80,
      color: '#3878D0',
      textGradient: {
        colors: ['#3878D0', '#56B6EE'],
        positions: [0, 64],
        angle: 180,
      },
      textShadow: {
        color: '#FFFFFF',
        opacity: 0.5,
        blur: 60000,
        distance: 120000,
        direction: 90,
      },
    };
    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="88"]',
    )!;

    expect(node.style.backgroundImage).toContain('linear-gradient');
    expect(node.style.color).toBe('transparent');
    expect(node.style.getPropertyValue('-webkit-text-fill-color')).toBe('transparent');
    expect(node.style.textShadow).toContain('rgba(255, 255, 255, 0.5)');

    await editor.destroy();
  });

  test('renders PPT run text shadows on rich text spans', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-run-text-shadow-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '88',
              name: 'Shadow Run Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2200000, height: 800000 },
              text: {
                text: '标题',
                body: { autoFit: 'shape', wrap: 'square' },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textParagraphs = [
      {
        text: '标题',
        runs: [
          {
            text: '标题',
            style: {
              textShadow: {
                color: '#FFFFFF',
                opacity: 0.43,
                blur: 38100,
                distance: 38100,
                direction: 45,
              },
            },
          },
        ],
      },
    ];
    const editor = renderNativePptEditor(host, view);
    const span = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="88"] .coco-ppt-rich-text span',
    )!;

    expect(span.style.textShadow).toContain('rgba(255, 255, 255, 0.43)');

    await editor.destroy();
  });

  test('allows single-paragraph bullet text to wrap inside auto-fit text boxes', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-bullet-autofit-wrap-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '82',
              name: 'Bullet Text',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 900000 },
              text: {
                text: 'Click on the button below this presentation preview. Start editing your presentation.',
                body: { autoFit: 'shape', wrap: 'square' },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.textBody = { autoFit: 'shape', wrap: 'square' };
    nodeModel.textParagraphs = [
      {
        text: 'Click on the button below this presentation preview. Start editing your presentation.',
        style: { bullet: '•', lineHeight: 1.4, marginLeft: 474981, textIndent: -218440 },
        runs: [
          {
            text: 'Click on the button below this presentation preview. Start editing your presentation.',
          },
        ],
      },
    ];

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="82"]',
    )!;
    const paragraph = node.querySelector<HTMLElement>('.coco-ppt-paragraph')!;

    expect(node.style.whiteSpace).not.toBe('nowrap');
    expect(paragraph.style.whiteSpace).not.toBe('nowrap');
    expect(Number(paragraph.style.lineHeight)).toBeCloseTo(1.61);

    await editor.destroy();
  });

  test('renders picture borders as an SVG overlay', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-picture-outline-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '9',
              name: 'Picture',
              kind: 'picture',
              geometry: { left: 100000, top: 100000, width: 900000, height: 700000 },
              style: {
                lineColor: '#5271FF',
                lineWidth: 76200,
              },
              image: {
                targetPath: 'ppt/media/image1.png',
              },
              raw: { presetGeometry: 'roundRect' },
            },
          ],
        },
      ],
    });
    view.images['ppt/media/image1.png'] = {
      url: 'blob:coco-picture-test',
      contentType: 'image/png',
    };
    view.scene.slides[0]!.nodes[0]!.kind = 'picture';
    view.scene.slides[0]!.nodes[0]!.imagePath = 'ppt/media/image1.png';
    view.scene.slides[0]!.nodes[0]!.lineColor = '#5271FF';
    view.scene.slides[0]!.nodes[0]!.lineWidth = 76200;
    view.scene.slides[0]!.nodes[0]!.presetGeometry = 'roundRect';

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="9"]',
    )!;

    expect(node.querySelector('img')).not.toBeNull();
    const outline = node.querySelector<SVGSVGElement>('.coco-ppt-solid-outline');
    const rect = outline?.querySelector('rect');
    expect(outline).not.toBeNull();
    expect(outline?.style.overflow).toBe('visible');
    expect(rect?.getAttribute('x')).toBe('0');
    expect(node.getAttribute('style') ?? '').not.toContain('border-style: solid');

    await editor.destroy();
  });

  test('clips custom geometry pictures to their shape path', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-custom-picture-clip-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '9',
              name: 'Custom Picture',
              kind: 'picture',
              geometry: { left: 100000, top: 100000, width: 900000, height: 700000 },
              image: {
                targetPath: 'ppt/media/image1.png',
                crop: { top: 0.1 },
                opacity: 0.5,
              },
              raw: {
                customGeometry: {
                  paths: [
                    {
                      width: 1000,
                      height: 1000,
                      commands: [
                        { kind: 'moveTo', x: 0, y: 0 },
                        { kind: 'lineTo', x: 1000, y: 0 },
                        { kind: 'lineTo', x: 800, y: 1000 },
                        { kind: 'lineTo', x: 0, y: 1000 },
                        { kind: 'close' },
                      ],
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    view.images['ppt/media/image1.png'] = {
      url: 'blob:coco-picture-test',
      contentType: 'image/png',
    };
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'picture';
    nodeModel.imagePath = 'ppt/media/image1.png';
    nodeModel.imageCrop = { top: 0.1 };
    nodeModel.imageOpacity = 0.5;
    nodeModel.customGeometry = {
      paths: [
        {
          width: 1000,
          height: 1000,
          commands: [
            { kind: 'moveTo', x: 0, y: 0 },
            { kind: 'lineTo', x: 1000, y: 0 },
            { kind: 'lineTo', x: 800, y: 1000 },
            { kind: 'lineTo', x: 0, y: 1000 },
            { kind: 'close' },
          ],
        },
      ],
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="9"]',
    )!;
    const clipSvg = node.querySelector<SVGSVGElement>('.coco-ppt-picture-clip')!;
    const image = clipSvg.querySelector<SVGImageElement>('image')!;

    expect(node.querySelector('img')).toBeNull();
    expect(clipSvg.querySelector('clipPath path')).not.toBeNull();
    expect(image.getAttribute('clip-path')).toContain('url(#coco-ppt-picture-clip-9)');
    expect(image.getAttribute('opacity')).toBe('0.5');
    expect(Number(image.getAttribute('y'))).toBeLessThan(0);

    await editor.destroy();
  });

  test('keeps full-slide custom geometry pictures rectangular', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-full-slide-custom-picture-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '9',
              name: 'Full Slide Picture',
              kind: 'picture',
              geometry: { left: 0, top: 0, width: 9144000, height: 5143500 },
              image: { targetPath: 'ppt/media/image1.png' },
              raw: {
                customGeometry: {
                  paths: [
                    {
                      width: 18288000,
                      height: 10287000,
                      commands: [
                        { kind: 'moveTo', x: 0, y: 0 },
                        { kind: 'lineTo', x: 18288000, y: 0 },
                        { kind: 'lineTo', x: 18288000, y: 10287000 },
                        { kind: 'lineTo', x: 0, y: 10287000 },
                        { kind: 'close' },
                      ],
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    view.images['ppt/media/image1.png'] = {
      url: 'blob:coco-full-slide-picture-test',
      contentType: 'image/png',
    };
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'picture';
    nodeModel.imagePath = 'ppt/media/image1.png';
    nodeModel.customGeometry = {
      paths: [
        {
          width: 18288000,
          height: 10287000,
          commands: [
            { kind: 'moveTo', x: 0, y: 0 },
            { kind: 'lineTo', x: 18288000, y: 0 },
            { kind: 'lineTo', x: 18288000, y: 10287000 },
            { kind: 'lineTo', x: 0, y: 10287000 },
            { kind: 'close' },
          ],
        },
      ],
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="9"]',
    )!;

    expect(node.querySelector('img')).not.toBeNull();
    expect(node.querySelector('.coco-ppt-picture-clip')).toBeNull();

    await editor.destroy();
  });

  test('renders chart axis and bar border styles from the model', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-chart-style-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '10',
              name: 'Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1400000 },
              chart: {
                type: 'bar',
                data: {
                  series: [
                    {
                      values: [3, 6],
                      categories: ['Q1', 'Q2'],
                      fillColor: '#00AEEF',
                      borderColor: '#FFFFFF',
                      borderWidth: 24000,
                    },
                  ],
                },
                axes: {
                  category: {
                    lineColor: '#00FF00',
                    lineWidth: 12000,
                  },
                  value: {
                    majorGridlines: true,
                    gridlineColor: '#FFFFFF',
                    gridlineWidth: 12000,
                  },
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'bar',
      data: {
        series: [
          {
            values: [2654, 5845],
            categories: ['Q1', 'Q2'],
            fillColor: '#00AEEF',
            borderColor: '#FFFFFF',
            borderWidth: 24000,
          },
        ],
      },
      axes: {
        category: {
          lineColor: '#00FF00',
          lineWidth: 12000,
        },
        value: {
          majorGridlines: true,
          gridlineColor: '#FFFFFF',
          gridlineWidth: 12000,
        },
      },
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="10"]',
    )!;
    const rect = node.querySelector<SVGRectElement>('svg rect[stroke="#FFFFFF"]');
    const line = node.querySelector<SVGLineElement>('svg line[stroke="#FFFFFF"]');
    const valueAxisLabels = [...node.querySelectorAll<SVGTextElement>('svg text')].filter(
      (item) => item.getAttribute('text-anchor') === 'end',
    );
    const valueAxisLabel = valueAxisLabels[0];

    expect(rect).not.toBeNull();
    expect(Number(rect!.getAttribute('x'))).toBeGreaterThan(26);
    expect(Number(rect!.getAttribute('stroke-width'))).toBeGreaterThan(0);
    expect(line).not.toBeNull();
    expect(Number(line!.getAttribute('stroke-width'))).toBeLessThan(1);
    expect(valueAxisLabel).toBeDefined();
    expect(Number(valueAxisLabel?.getAttribute('x'))).toBeGreaterThanOrEqual(14);
    expect(valueAxisLabels.map((item) => item.textContent)).toContain('7000');

    await editor.destroy();
  });

  test('renders stacked bar grouping as overlaid category columns', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-stacked-bar-chart-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '17',
              name: 'Stacked Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1400000 },
              chart: {
                type: 'bar',
                grouping: 'stacked',
                data: {
                  series: [
                    { values: [4, 5.5], categories: ['Q1', 'Q2'], fillColor: '#56B6EE' },
                    { values: [2, 0.5], categories: ['Q1', 'Q2'], fillColor: '#3878D0' },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'bar',
      grouping: 'stacked',
      data: {
        series: [
          { values: [4, 5.5], categories: ['Q1', 'Q2'], fillColor: '#56B6EE' },
          { values: [2, 0.5], categories: ['Q1', 'Q2'], fillColor: '#3878D0' },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="17"]',
    )!;
    const rects = [...node.querySelectorAll<SVGRectElement>('svg rect')];
    const labels = [...node.querySelectorAll<SVGTextElement>('svg text')];
    const lowerFirstCategory = rects.find((rect) => rect.getAttribute('fill') === '#56B6EE');
    const upperFirstCategory = rects.find((rect) => rect.getAttribute('fill') === '#3878D0');

    expect(labels.map((item) => item.textContent)).toContain('6');
    expect(lowerFirstCategory).toBeDefined();
    expect(upperFirstCategory).toBeDefined();
    expect(lowerFirstCategory!.getAttribute('x')).toBe(upperFirstCategory!.getAttribute('x'));
    expect(Number(upperFirstCategory!.getAttribute('y'))).toBeLessThan(
      Number(lowerFirstCategory!.getAttribute('y')),
    );

    await editor.destroy();
  });

  test('renders separate bar chart groups on the same category center', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-overlay-bar-chart-groups-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '21',
              name: 'Overlay Bar Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1400000 },
              chart: {
                type: 'bar',
                grouping: 'clustered',
                data: {
                  series: [
                    {
                      chartGroup: 0,
                      values: [4, 5.5],
                      categories: ['Q1', 'Q2'],
                      fillColor: '#56B6EE',
                    },
                    {
                      chartGroup: 1,
                      values: [2.4, 3.2],
                      categories: ['Q1', 'Q2'],
                      fillColor: '#3878D0',
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'bar',
      grouping: 'clustered',
      title: '可EXCEL编辑图表标题',
      data: {
        series: [
          {
            name: '系列 2',
            chartGroup: 0,
            values: [4, 5.5],
            categories: ['Q1', 'Q2'],
            fillColor: '#56B6EE',
          },
          {
            name: '系列 1',
            chartGroup: 1,
            values: [2.4, 3.2],
            categories: ['Q1', 'Q2'],
            fillColor: '#3878D0',
          },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="21"]',
    )!;
    const rects = [...node.querySelectorAll<SVGRectElement>('svg rect')];
    const labels = [...node.querySelectorAll<SVGTextElement>('svg text')];
    const barRects = rects.filter((rect) => Number(rect.getAttribute('width')) > 5);
    const backFirstCategory = barRects.find((rect) => rect.getAttribute('fill') === '#56B6EE');
    const frontFirstCategory = barRects.find((rect) => rect.getAttribute('fill') === '#3878D0');

    expect(labels.map((item) => item.textContent)).toEqual(
      expect.arrayContaining(['可EXCEL编辑图表标题', '系列 2', '系列 1', '6', '5']),
    );
    expect(backFirstCategory).toBeDefined();
    expect(frontFirstCategory).toBeDefined();
    expect(backFirstCategory!.getAttribute('x')).toBe(frontFirstCategory!.getAttribute('x'));
    expect(Number(frontFirstCategory!.getAttribute('height'))).toBeLessThan(
      Number(backFirstCategory!.getAttribute('height')),
    );

    await editor.destroy();
  });

  test('keeps doughnut chart center labels readable', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-doughnut-center-label-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '18',
              name: 'Doughnut Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1600000 },
              chart: {
                type: 'doughnut',
                data: {
                  series: [
                    {
                      name: '市场渗透率指标',
                      values: [21.3, 33.5, 26.7],
                      categories: ['1月', '2月', '3月'],
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'doughnut',
      data: {
        series: [
          {
            name: '市场渗透率指标',
            values: [21.3, 33.5, 26.7],
            categories: ['1月', '2月', '3月'],
          },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const labels = [...host.shadowRoot!.querySelectorAll<SVGTextElement>('svg text')];
    const centerLabel = labels.find((item) => item.textContent === '市场渗透率…');
    const percentLabel = labels.find(
      (item) => item.textContent === '26.1%' && Number(item.getAttribute('font-size')) > 3,
    );

    expect(centerLabel).toBeDefined();
    expect(percentLabel).toBeDefined();
    expect(Number(centerLabel?.getAttribute('y'))).toBeLessThan(
      Number(percentLabel?.getAttribute('y')),
    );
    expect(Number(centerLabel?.getAttribute('font-size'))).toBeLessThan(
      Number(percentLabel?.getAttribute('font-size')),
    );

    await editor.destroy();
  });

  test('renders doughnut chart ratio values as direct percentages', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-doughnut-ratio-percent-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '19',
              name: 'Doughnut Ratio Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1600000 },
              chart: {
                type: 'doughnut',
                data: {
                  series: [
                    {
                      name: '市场渗透率指标',
                      values: [0.222, 0.335, 0.267, 0.178],
                      categories: ['1月', '2月', '3月', '4月'],
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'doughnut',
      data: {
        series: [
          {
            name: '市场渗透率指标',
            values: [0.222, 0.335, 0.267, 0.178],
            categories: ['1月', '2月', '3月', '4月'],
          },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const labels = [...host.shadowRoot!.querySelectorAll<SVGTextElement>('svg text')];

    expect(labels.some((item) => item.textContent === '22.2%')).toBe(true);

    await editor.destroy();
  });

  test('does not add default radar markers when the model has no marker metadata', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-radar-no-default-marker-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '22',
              name: 'Radar Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1600000 },
              chart: {
                type: 'radar',
                data: {
                  series: [
                    {
                      values: [32, 32, 28, 12, 15],
                      categories: ['化妆产品', '服装服饰', '母婴用品', '数码产品', '其他分类'],
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'radar',
      data: {
        series: [
          {
            values: [32, 32, 28, 12, 15],
            categories: ['化妆产品', '服装服饰', '母婴用品', '数码产品', '其他分类'],
          },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="22"]',
    )!;

    expect(node.querySelectorAll('svg circle')).toHaveLength(0);
    expect(node.querySelector('svg path[fill=\"rgba(37, 99, 235, 0.08)\"]')).not.toBeNull();

    await editor.destroy();
  });

  test('renders smooth chart series with cubic paths', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-chart-smooth-line-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '10',
              name: 'Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1400000 },
              chart: {
                type: 'line',
                data: {
                  series: [
                    {
                      values: [14, 18, 9, 16],
                      categories: ['1月', '2月', '3月', '4月'],
                      lineColor: '#22C7F2',
                      smooth: true,
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'line',
      data: {
        series: [
          {
            values: [14, 18, 9, 16],
            categories: ['1月', '2月', '3月', '4月'],
            lineColor: '#22C7F2',
            smooth: true,
          },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const path = host.shadowRoot!.querySelector<SVGPathElement>(
      '.coco-ppt-canvas .coco-ppt-node[data-shape-id="10"] svg path',
    )!;

    expect(path.getAttribute('d')).toContain('C ');

    await editor.destroy();
  });

  test('renders line chart value labels when requested by chart data', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-line-value-labels-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '10',
              name: 'Line Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1400000 },
              chart: {
                type: 'line',
                data: {
                  series: [
                    {
                      name: '产品A',
                      type: 'line',
                      values: [680, 450, 510],
                      categories: ['1月', '2月', '3月'],
                      lineColor: '#133E82',
                      marker: {
                        symbol: 'circle',
                        size: 6,
                        fillColor: '#FFFFFF',
                        lineColor: '#133E82',
                        lineWidth: 9525,
                      },
                      showValueLabels: true,
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'line',
      data: {
        series: [
          {
            name: '产品A',
            type: 'line',
            values: [680, 450, 510],
            categories: ['1月', '2月', '3月'],
            lineColor: '#133E82',
            marker: {
              symbol: 'circle',
              size: 6,
              fillColor: '#FFFFFF',
              lineColor: '#133E82',
              lineWidth: 9525,
            },
            showValueLabels: true,
          },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const labels = [...host.shadowRoot!.querySelectorAll<SVGTextElement>('svg text')].map(
      (item) => item.textContent,
    );
    const marker = host.shadowRoot!.querySelector<SVGCircleElement>('svg circle')!;
    const legendLine = [...host.shadowRoot!.querySelectorAll<SVGLineElement>('svg line')].find(
      (item) => item.getAttribute('stroke') === '#133E82',
    );

    expect(labels).toEqual(expect.arrayContaining(['680', '450', '510']));
    expect(marker.getAttribute('fill')).toBe('#FFFFFF');
    expect(marker.getAttribute('stroke')).toBe('#133E82');
    expect(legendLine).toBeTruthy();

    await editor.destroy();
  });

  test('renders value labels for every labelled clustered bar series', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-bar-series-value-labels-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '11',
              name: 'Bar Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1400000 },
              chart: {
                type: 'bar',
                data: {
                  series: [
                    {
                      name: 'A',
                      values: [2654, 965],
                      categories: ['1月', '2月'],
                      showValueLabels: true,
                    },
                    {
                      name: 'B',
                      values: [0, 2365],
                      categories: ['1月', '2月'],
                      showValueLabels: true,
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'bar',
      data: {
        series: [
          { name: 'A', values: [2654, 965], categories: ['1月', '2月'], showValueLabels: true },
          { name: 'B', values: [0, 2365], categories: ['1月', '2月'], showValueLabels: true },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const labels = [...host.shadowRoot!.querySelectorAll<SVGTextElement>('svg text')].map(
      (item) => item.textContent,
    );

    expect(labels).toEqual(expect.arrayContaining(['2654', '965', '0', '2365']));

    await editor.destroy();
  });

  test('renders chart series gradients with opacity stops', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-chart-gradient-opacity-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '13',
              name: 'Gradient Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1400000 },
              chart: {
                type: 'bar',
                data: {
                  series: [
                    {
                      type: 'bar',
                      values: [4, 5.5],
                      categories: ['类别 1', '类别 2'],
                      fillGradient: {
                        colors: ['#56B6EE', '#7FCEFB'],
                        positions: [0, 100],
                        opacities: [1, 0],
                        angle: 90,
                      },
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'bar',
      data: {
        series: [
          {
            type: 'bar',
            values: [4, 5.5],
            categories: ['类别 1', '类别 2'],
            fillGradient: {
              colors: ['#56B6EE', '#7FCEFB'],
              positions: [0, 100],
              opacities: [1, 0],
              angle: 90,
            },
          },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const stops = [...host.shadowRoot!.querySelectorAll<SVGStopElement>('svg linearGradient stop')];

    expect(stops.map((item) => item.getAttribute('stop-color'))).toEqual(
      expect.arrayContaining(['#56B6EE', '#7FCEFB']),
    );
    expect(stops.some((item) => item.getAttribute('stop-opacity') === '0')).toBe(true);

    await editor.destroy();
  });

  test('renders doughnut category and percent labels when requested by chart data', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-doughnut-labels-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '12',
              name: 'Doughnut Chart',
              kind: 'chart',
              geometry: { left: 100000, top: 100000, width: 2400000, height: 1400000 },
              chart: {
                type: 'doughnut',
                data: {
                  series: [
                    {
                      name: '销售占比',
                      values: [0.222, 0.335],
                      categories: ['1月', '2月'],
                      showCategoryNameLabels: true,
                      showPercentLabels: true,
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.kind = 'chart';
    nodeModel.chart = {
      type: 'doughnut',
      data: {
        series: [
          {
            name: '销售占比',
            values: [0.222, 0.335],
            categories: ['1月', '2月'],
            showCategoryNameLabels: true,
            showPercentLabels: true,
          },
        ],
      },
    };

    const editor = renderNativePptEditor(host, view);
    const labels = [...host.shadowRoot!.querySelectorAll<SVGTextElement>('svg text')].map(
      (item) => item.textContent,
    );

    expect(labels).toEqual(expect.arrayContaining(['1月', '22.2%', '2月', '33.5%']));

    await editor.destroy();
  });

  test('renders rich text bullets with explicit bullet color', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-bullet-style-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '11',
              name: 'Bullets',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1800000, height: 500000 },
              text: {
                text: 'Done',
                paragraphs: [
                  {
                    text: 'Done',
                    style: {
                      bullet: '✓',
                      bulletColor: '#22AA88',
                      bulletFontFamily: 'Wingdings',
                      bulletSize: 0.8,
                    },
                    runs: [{ text: 'Done' }],
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.text = 'Done';
    nodeModel.textParagraphs = [
      {
        text: 'Done',
        style: {
          bullet: '✓',
          bulletColor: '#22AA88',
          bulletFontFamily: 'Wingdings',
          bulletSize: 0.8,
        },
        runs: [{ text: 'Done' }],
      },
    ];

    const editor = renderNativePptEditor(host, view);
    const bullet = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-bullet')!;

    expect(bullet.textContent).toBe('✓');
    expect(bullet.style.color).toBe('#22AA88');
    expect(bullet.style.fontSize).toBe('0.8em');

    await editor.destroy();
  });

  test('renders compact centered Wingdings square bullets smaller than normal bullets', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-compact-bullet-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '11',
              name: 'Compact bullet label',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1800000, height: 500000 },
              text: {
                text: '目录',
                paragraphs: [
                  {
                    text: '目录',
                    style: {
                      alignment: 'center',
                      bullet: '□',
                      bulletFontFamily: 'Wingdings',
                      marginLeft: 285750,
                      textIndent: -285750,
                    },
                    runs: [{ text: '目录' }],
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.text = '目录';
    nodeModel.textParagraphs = [
      {
        text: '目录',
        style: {
          alignment: 'center',
          bullet: '□',
          bulletFontFamily: 'Wingdings',
          marginLeft: 285750,
          textIndent: -285750,
        },
        runs: [{ text: '目录' }],
      },
    ];

    const editor = renderNativePptEditor(host, view);
    const bullet = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-bullet')!;

    expect(bullet.textContent).toBe('□');
    expect(bullet.style.width).toBe('0.5em');
    expect(bullet.style.transform).toBe('scale(0.62)');

    await editor.destroy();
  });

  test('keeps Wingdings square bullets full-size in wider title bars', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-title-bar-bullet-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '11',
              name: 'Title bar bullet',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 4000000, height: 360000 },
              text: { text: '研发核心突破与成效' },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.text = '研发核心突破与成效';
    nodeModel.textParagraphs = [
      {
        text: '研发核心突破与成效',
        style: {
          alignment: 'center',
          bullet: '□',
          bulletFontFamily: 'Wingdings',
          marginLeft: 285750,
          textIndent: -285750,
        },
        runs: [{ text: '研发核心突破与成效' }],
      },
    ];

    const editor = renderNativePptEditor(host, view);
    const bullet = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-bullet')!;

    expect(bullet.textContent).toBe('□');
    expect(bullet.style.width).toBe('');
    expect(bullet.style.transform).toBe('');

    await editor.destroy();
  });

  test('allows transparent text boxes to overflow like native presentation text labels', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-transparent-text-overflow-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '11',
              name: 'Transparent text label',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1200000, height: 130000 },
              text: { text: 'Improve customer satisfaction' },
              style: {
                fillColor: '#000000',
                opacity: 0,
                lineColor: '#000000',
                lineOpacity: 0,
                lineWidth: 0,
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.fillColor = '#000000';
    nodeModel.opacity = 0;
    nodeModel.lineColor = '#000000';
    nodeModel.lineOpacity = 0;
    nodeModel.lineWidth = 0;

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="11"]')!;

    expect(node.style.overflow).toBe('visible');

    await editor.destroy();
  });

  test('keeps visible filled text boxes clipped to their shape bounds', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const view = await buildPptEditorViewFromModel({
      id: 'native-filled-text-clipping-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '11',
              name: 'Filled text label',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1200000, height: 130000 },
              text: { text: 'Visible fill keeps clipping' },
              style: {
                fillColor: '#FFFFFF',
                lineColor: '#000000',
                lineWidth: 9525,
              },
            },
          ],
        },
      ],
    });
    const nodeModel = view.scene.slides[0]!.nodes[0]!;
    nodeModel.fillColor = '#FFFFFF';
    nodeModel.lineColor = '#000000';
    nodeModel.lineWidth = 9525;

    const editor = renderNativePptEditor(host, view);
    const node = host.shadowRoot!.querySelector<HTMLElement>('.coco-ppt-node[data-shape-id="11"]')!;

    expect(node.style.overflow).toBe('');

    await editor.destroy();
  });
});

async function makeNativePlaceholderInheritancePptx(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await makeNativeGroupedTablePptx());
  const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
  zip.file(
    '[Content_Types].xml',
    contentTypes.replace(
      '</Types>',
      '  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>\n' +
        '  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>\n' +
        '</Types>',
    ),
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Office preview</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`,
  );
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
  );
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Layout Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Layout title</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sldLayout>`,
  );
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`,
  );
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Master Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:t>Master title</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle></p:txStyles>
</p:sldMaster>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeNativeGroupedTablePptx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`,
    { createFolders: false },
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`,
    { createFolders: false },
  );
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`,
    { createFolders: false },
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    { createFolders: false },
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="6" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="1200"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1200"/></a:xfrm></p:grpSpPr>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="15" name="GroupedTable"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="1450" y="1120"/><a:ext cx="450" cy="240"/></p:xfrm>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
              <a:tbl>
                <a:tr h="120">
                  <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>G1</a:t></a:r></a:p></a:txBody></a:tc>
                  <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>G2</a:t></a:r></a:p></a:txBody></a:tc>
                </a:tr>
              </a:tbl>
            </a:graphicData>
          </a:graphic>
        </p:graphicFrame>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
    { createFolders: false },
  );
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}
