import { describe, expect, test, vi } from 'vitest';
import JSZip from 'jszip';
import { Presentation } from '@coco-vexel/coco-office/ppt';
import {
  buildPptEditorViewFromModel,
  buildPptEditorViewFromPresentation,
  buildPptEditorViewFromPptx,
} from '../view.js';
import { PptEditorSession } from '../session.js';
import { buildPptPresentationScene } from '../scene.js';

describe('PptEditorSession', () => {
  test('allows selection but rejects edits and history while read-only', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'read-only-session-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [{
        id: '256',
        index: 0,
        shapes: [{
          id: '2',
          name: 'Title',
          kind: 'textBox',
          geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
          text: { text: 'Draft' },
        }],
      }],
    });
    const session = new PptEditorSession(view, { readOnly: true });

    await session.dispatch({ kind: 'select-shape', slideIdOrIndex: 0, shapeIdOrName: 'Title' });
    expect(session.selection()).toMatchObject({ kind: 'shape', shapeId: '2' });
    await expect(session.dispatch({ kind: 'set-shape-text', slideIdOrIndex: 0, shapeIdOrName: 'Title', text: 'blocked' }))
      .rejects.toMatchObject({ name: 'EditorReadOnlyError', code: 'EDITOR_READ_ONLY' });
    await expect(session.undo()).rejects.toMatchObject({ code: 'EDITOR_READ_ONLY' });
    expect(() => session.presentation()).toThrowError(/read-only/);
    expect(session.revision()).toBe(0);
    expect(session.scene().slides[0]!.nodes[0]!.text).toBe('Draft');

    session.setReadOnly(false);
    await session.dispatch({ kind: 'set-shape-text', slideIdOrIndex: 0, shapeIdOrName: 'Title', text: 'Final' });
    expect(session.scene().slides[0]!.nodes[0]!.text).toBe('Final');
    await session.destroy();
  });

  test('uses resolved image previews without replacing the exported image resource', async () => {
    const originalCreateObjectURL = URL.createObjectURL;
    Object.defineProperty(URL, 'createObjectURL', {
      configurable: true,
      value: vi.fn(() => 'blob:original-image'),
    });
    try {
      const bytes = new Uint8Array([1, 2, 3, 4]);
      const presentation = await Presentation.createBlank({ presentationId: 'image-preview-test' });
      await presentation.slides.getItem(0).shapes.addImage(
        { bytes, mimeType: 'image/x-emf', filename: 'diagram.emf' },
        {
          name: 'VectorPreview',
          geometry: { left: 100000, top: 100000, width: 2000000, height: 1000000 },
        },
      );

      const resolver = vi.fn(() => ({
        url: 'blob:preview-png',
        contentType: 'image/png',
      }));
      const view = await buildPptEditorViewFromPresentation(presentation, {
        resolveImagePreview: resolver,
      });

      expect(resolver).toHaveBeenCalledWith({
        path: 'ppt/media/image1.emf',
        contentType: 'image/x-emf',
        bytes,
      });
      expect(view.images['ppt/media/image1.emf']).toEqual({
        url: 'blob:preview-png',
        contentType: 'image/png',
      });
      expect(URL.createObjectURL).not.toHaveBeenCalled();
      await expect(presentation.exportImages()).resolves.toEqual([
        {
          path: 'ppt/media/image1.emf',
          contentType: 'image/x-emf',
          byteLength: bytes.byteLength,
          bytes,
        },
      ]);
    } finally {
      Object.defineProperty(URL, 'createObjectURL', {
        configurable: true,
        value: originalCreateObjectURL,
      });
    }
  });

  test('applies text commands through core and verifies exported pptx', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-test',
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
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-shape-text',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      text: 'Final',
    });

    expect(session.scene().slides[0]!.nodes[0]!.text).toBe('Final');
    const verified = await session.verifyExport();
    expect(verified.slides[0]!.shapes[0]!.text?.text).toBe('Final');

    await session.undo();
    expect(session.scene().slides[0]!.nodes[0]!.text).toBe('Draft');
    await session.redo();
    expect(session.scene().slides[0]!.nodes[0]!.text).toBe('Final');

    await session.destroy();
  });

  test('preserves shape selection across undo and redo when the shape still exists', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-selection-history-test',
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
    const session = new PptEditorSession(view);

    await session.dispatch({ kind: 'select-shape', slideIdOrIndex: 0, shapeIdOrName: 'Title' });
    await session.dispatch({
      kind: 'set-shape-text',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      text: 'Final',
    });
    await session.undo();
    expect(session.selection()).toMatchObject({ kind: 'shape', slideId: '256', shapeId: '2' });
    expect(session.scene().slides[0]!.nodes[0]!.text).toBe('Draft');
    await session.redo();
    expect(session.selection()).toMatchObject({ kind: 'shape', slideId: '256', shapeId: '2' });
    expect(session.scene().slides[0]!.nodes[0]!.text).toBe('Final');

    await session.destroy();
  });

  test('preserves text style in the scene after text edits', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-style-test',
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
    await view.presentation.slides.getItem(0).shapes.getItem('Title').text.setStyle({
      fontSize: 36,
      color: '#FFFFFF',
      bold: true,
      alignment: 'center',
    });
    const styledView = await buildPptEditorViewFromPresentation(view.presentation);
    const session = new PptEditorSession(styledView);

    await session.dispatch({
      kind: 'set-shape-text',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      text: 'Final',
    });

    const node = session.scene().slides[0]!.nodes[0]!;
    expect(node.text).toBe('Final');
    expect(node.textStyle).toMatchObject({
      fontSize: 36,
      color: '#FFFFFF',
      bold: true,
      alignment: 'center',
    });
    const verified = await session.verifyExport();
    expect(verified.slides[0]!.shapes[0]!.text?.style).toMatchObject({
      fontSize: 36,
      color: '#FFFFFF',
      bold: true,
      alignment: 'center',
    });
    expect(verified.slides[0]!.shapes[0]!.text?.paragraphs?.[0]?.runs).toHaveLength(1);
    expect(verified.slides[0]!.shapes[0]!.text?.paragraphs?.[0]?.runs[0]).toMatchObject({
      text: 'Final',
      style: {
        fontSize: 36,
        color: '#FFFFFF',
        bold: true,
        alignment: 'center',
      },
    });

    await session.destroy();
  });

  test('applies text style commands through core and verifies exported pptx', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-text-style-command-test',
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
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-shape-text-style',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      style: { fontSize: 28, color: '#336699', bold: true, alignment: 'center' },
    });

    expect(session.scene().slides[0]!.nodes[0]!.textStyle).toMatchObject({
      fontSize: 28,
      color: '#336699',
      bold: true,
      alignment: 'center',
    });
    const verified = await session.verifyExport();
    expect(verified.slides[0]!.shapes[0]!.text?.style).toMatchObject({
      fontSize: 28,
      color: '#336699',
      bold: true,
      alignment: 'center',
    });

    await session.destroy();
  });

  test('applies shape style and geometry commands through core', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-shape-style-geometry-command-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Box',
              kind: 'geometricShape',
              raw: { presetGeometry: 'rect' },
              geometry: { left: 100000, top: 200000, width: 1000000, height: 500000 },
              style: { fillColor: '#FFFFFF', lineColor: '#000000', lineWidth: 9525 },
            },
          ],
        },
      ],
    });
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-shape-style',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Box',
      style: { fillColor: '#FFCC00', lineColor: '#2255AA', lineWidth: 19050 },
    });
    await session.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Box',
      geometry: { left: 300000, top: 400000, width: 1200000, height: 650000 },
    });

    const shape = session.model().slides[0]!.shapes[0]!;
    expect(shape.style).toMatchObject({
      fillColor: '#FFCC00',
      lineColor: '#2255AA',
      lineWidth: 19050,
    });
    expect(shape.geometry).toMatchObject({
      left: 300000,
      top: 400000,
      width: 1200000,
      height: 650000,
    });
    const verified = await session.verifyExport();
    expect(verified.slides[0]!.shapes[0]!.style).toMatchObject({
      fillColor: '#FFCC00',
      lineColor: '#2255AA',
      lineWidth: 19050,
    });
    expect(verified.slides[0]!.shapes[0]!.geometry).toMatchObject({
      left: 300000,
      top: 400000,
      width: 1200000,
      height: 650000,
    });

    await session.destroy();
  });

  test('deletes and duplicates slide-owned shapes through core shape commands', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-shape-collection-command-test',
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
            {
              id: '3',
              name: 'Box',
              kind: 'geometricShape',
              raw: { presetGeometry: 'rect' },
              geometry: { left: 400000, top: 800000, width: 1000000, height: 500000 },
              style: { fillColor: '#FFFFFF' },
            },
          ],
        },
      ],
    });
    const session = new PptEditorSession(view);

    await session.dispatch({ kind: 'duplicate-shape', slideIdOrIndex: 0, shapeIdOrName: 'Title' });

    expect(session.model().slides[0]!.shapes).toHaveLength(3);
    const copy = session.model().slides[0]!.shapes[1]!;
    expect(copy.id).toBe('4');
    expect(copy.name).toBe('Title Copy 4');
    expect(copy.text?.text).toBe('Draft');
    expect(copy.geometry.left).toBe(280000);
    expect(session.selection()).toMatchObject({ kind: 'shape', slideId: '256', shapeId: '4' });

    await session.dispatch({ kind: 'delete-shape', slideIdOrIndex: 0, shapeIdOrName: 'Box' });

    expect(session.model().slides[0]!.shapes.map((shape) => shape.name)).toEqual([
      'Title',
      'Title Copy 4',
    ]);
    expect(session.selection()).toMatchObject({ kind: 'slide', slideId: '256' });
    const verified = await session.verifyExport();
    expect(verified.slides[0]!.shapes.map((shape) => shape.name)).toEqual([
      'Title',
      'Title Copy 4',
    ]);

    await session.undo();
    expect(session.model().slides[0]!.shapes.map((shape) => shape.name)).toEqual([
      'Title',
      'Title Copy 4',
      'Box',
    ]);
    await session.redo();
    expect(session.model().slides[0]!.shapes.map((shape) => shape.name)).toEqual([
      'Title',
      'Title Copy 4',
    ]);

    await session.destroy();
  });

  test('pastes a copied shape with a new id and preserves export content', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-paste-shape-command-test',
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
              text: { text: 'Copied text' },
            },
          ],
        },
      ],
    });
    const session = new PptEditorSession(view);

    await session.dispatch({ kind: 'paste-shape', slideIdOrIndex: '256', shapeIdOrName: 'Title' });

    const shapes = session.model().slides[0]!.shapes;
    expect(shapes).toHaveLength(2);
    expect(shapes[1]).toMatchObject({
      id: '3',
      name: 'Title Copy 3',
      text: { text: 'Copied text' },
    });
    expect(session.selection()).toMatchObject({ kind: 'shape', slideId: '256', shapeId: '3' });
    const verified = await session.verifyExport();
    expect(verified.slides[0]!.shapes[1]!.text?.text).toBe('Copied text');

    await session.destroy();
  });

  test('changes slide-owned shape z-order through core commands', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-shape-z-order-command-test',
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
              name: 'Middle',
              kind: 'geometricShape',
              raw: { presetGeometry: 'rect' },
              geometry: { left: 200000, top: 200000, width: 1000000, height: 500000 },
            },
            {
              id: '4',
              name: 'Front',
              kind: 'geometricShape',
              raw: { presetGeometry: 'rect' },
              geometry: { left: 300000, top: 300000, width: 1000000, height: 500000 },
            },
          ],
        },
      ],
    });
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'bring-shape-to-front',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Back',
    });
    expect(session.model().slides[0]!.shapes.map((shape) => shape.name)).toEqual([
      'Middle',
      'Front',
      'Back',
    ]);
    expect(session.selection()).toMatchObject({ kind: 'shape', shapeId: '2' });

    await session.dispatch({
      kind: 'send-shape-to-back',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Front',
    });
    expect(session.model().slides[0]!.shapes.map((shape) => shape.name)).toEqual([
      'Front',
      'Middle',
      'Back',
    ]);
    const verified = await session.verifyExport();
    expect(verified.slides[0]!.shapes.map((shape) => shape.name)).toEqual([
      'Front',
      'Middle',
      'Back',
    ]);

    await session.destroy();
  });

  test('exposes object edit capabilities by object boundary', () => {
    const scene = buildPptPresentationScene({
      id: 'editor-capability-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: 'layout:10',
              name: 'Layout Title',
              kind: 'textBox',
              geometry: { left: 0, top: 0, width: 1000000, height: 400000 },
              text: { text: 'Layout' },
            },
            {
              id: '2',
              name: 'Body',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Body' },
            },
            {
              id: '3',
              name: 'Image',
              kind: 'picture',
              geometry: { left: 0, top: 700000, width: 1000000, height: 1000000 },
              image: { targetPath: 'media/image1.png' },
            },
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 0, top: 1800000, width: 2000000, height: 800000 },
              table: { values: [['A1']], rowCount: 1, columnCount: 1 },
            },
            {
              id: '7',
              name: 'Grouped',
              kind: 'group',
              geometry: { left: 0, top: 0, width: 3000000, height: 2200000 },
              raw: {
                groupTransform: {
                  left: 0,
                  top: 0,
                  width: 3000000,
                  height: 2200000,
                  childLeft: 0,
                  childTop: 0,
                  childWidth: 3000000,
                  childHeight: 2200000,
                },
              },
              children: [
                {
                  id: '9',
                  name: 'Grouped Text',
                  kind: 'textBox',
                  geometry: { left: 200000, top: 100000, width: 1000000, height: 300000 },
                  text: { text: 'Grouped label' },
                },
                {
                  id: '8',
                  name: 'Grouped Table',
                  kind: 'table',
                  geometry: { left: 200000, top: 200000, width: 1000000, height: 500000 },
                  table: { values: [['G1']], rowCount: 1, columnCount: 1 },
                },
              ],
            },
            {
              id: '5',
              name: 'Chart',
              kind: 'chart',
              geometry: { left: 0, top: 2700000, width: 2000000, height: 1200000 },
              chart: { type: 'bar' },
            },
            {
              id: '6',
              name: 'SmartArt',
              kind: 'smartArt',
              geometry: { left: 0, top: 4000000, width: 2000000, height: 800000 },
            },
            {
              id: '10',
              name: 'BackgroundFillShape',
              kind: 'geometricShape',
              geometry: { left: 100000, top: 4200000, width: 1000000, height: 600000 },
              style: { fillColor: '#5B9BD5' },
              raw: { useBackgroundFill: true },
            },
          ],
        },
      ],
    });
    const nodes = scene.slides[0]!.nodes;
    const layoutTitle = nodes.find((node) => node.id === 'layout:10')!;
    const body = nodes.find((node) => node.id === '2')!;
    const image = nodes.find((node) => node.id === '3')!;
    const table = nodes.find((node) => node.id === '4')!;
    const groupedText = nodes.find((node) => node.id === '7:9')!;
    const groupedTable = nodes.find((node) => node.id === '7:8')!;
    const chart = nodes.find((node) => node.id === '5')!;
    const smartArt = nodes.find((node) => node.id === '6')!;
    const backgroundFillShape = nodes.find((node) => node.id === '10')!;

    expect(layoutTitle.locked).toBe(true);
    expect(layoutTitle.capabilities).toMatchObject({
      editText: false,
      editTextStyle: false,
      editGeometry: false,
      editShapeStyle: false,
      reasons: {
        editText: '继承自母版/版式，当前只支持查看',
        editGeometry: '继承自母版/版式，当前只支持查看',
      },
    });
    expect(body.capabilities).toMatchObject({
      editText: true,
      editTextStyle: true,
      editGeometry: true,
      editShapeStyle: true,
    });
    expect(body.capabilities.reasons?.editTableCells).toBe('当前对象不是表格');
    expect(image.capabilities).toMatchObject({
      editText: false,
      editTextStyle: false,
      editGeometry: true,
      editShapeStyle: true,
      reasons: {
        editText: '图片对象没有可编辑文字',
        editTableCells: '当前对象不是表格',
      },
    });
    expect(table.capabilities).toMatchObject({
      editText: false,
      editTableCells: true,
      editGeometry: true,
      editShapeStyle: true,
      reasons: {
        editText: '表格文字请使用单元格编辑',
      },
    });
    expect(groupedText.capabilities).toMatchObject({
      editText: true,
      editTextStyle: true,
      editGeometry: true,
      editShapeStyle: false,
      reason: 'Flattened group child',
      reasons: {
        editTableCells: '当前对象不是表格',
        editShapeStyle: '组内子对象暂不支持该项直接编辑',
      },
    });
    expect(groupedTable.capabilities).toMatchObject({
      editText: false,
      editTableCells: true,
      editGeometry: true,
      editShapeStyle: false,
      reason: 'Flattened group child',
      reasons: {
        editText: '当前对象没有可编辑文字内容',
        editShapeStyle: '组内子对象暂不支持该项直接编辑',
      },
    });
    expect(chart.capabilities).toMatchObject({
      editText: false,
      editTextStyle: false,
      editTableCells: false,
      editGeometry: true,
      editShapeStyle: false,
      reasons: {
        editText: '图表文字绑定在图表数据中，当前仅开放整体位置尺寸',
        editShapeStyle: '图表形状样式暂未开放',
      },
    });
    expect(smartArt.capabilities).toMatchObject({
      editText: false,
      editTextStyle: false,
      editTableCells: false,
      editGeometry: true,
      editShapeStyle: false,
      reasons: {
        editText: 'SmartArt 内部文本结构暂未开放',
        editShapeStyle: 'SmartArt 样式暂未开放',
      },
    });
    expect(backgroundFillShape.fillColor).toBeUndefined();
    expect(backgroundFillShape.fillGradient).toBeUndefined();
    expect(backgroundFillShape.usesBackgroundFill).toBe(true);
  });

  test('renders slide placeholders with geometry and text properties inherited through layout and master', () => {
    const scene = buildPptPresentationScene({
      id: 'editor-placeholder-inheritance-test',
      slideSize: { width: 9144000, height: 6858000 },
      slides: [
        {
          id: '256',
          index: 0,
          layout: {
            relationshipId: 'rId1',
            path: 'ppt/slideLayouts/slideLayout1.xml',
            shapes: [
              {
                id: 'layout-title',
                kind: 'placeholder',
                placeholder: { type: 'title' },
                geometry: {},
                text: { text: 'Layout title' },
              },
              {
                id: 'layout-body',
                kind: 'placeholder',
                placeholder: { index: '1' },
                geometry: {},
                text: { text: 'Layout body' },
              },
            ],
            master: {
              relationshipId: 'rId1',
              path: 'ppt/slideMasters/slideMaster1.xml',
              textStyles: {
                title: { fontSize: 44, color: '#111111' },
                body: { fontSize: 32, color: '#222222' },
              },
              shapes: [
                {
                  id: 'master-title',
                  kind: 'placeholder',
                  placeholder: { type: 'title' },
                  geometry: { left: 457200, top: 274638, width: 8229600, height: 1143000 },
                  text: { text: 'Master title', body: { anchor: 'middle', insetLeft: 91440 } },
                },
                {
                  id: 'master-body',
                  kind: 'placeholder',
                  placeholder: { type: 'body', index: '1' },
                  geometry: { left: 457200, top: 1600200, width: 8229600, height: 4525963 },
                  text: { text: 'Master body', body: { autoFit: 'normal', insetLeft: 91440 } },
                },
              ],
            },
          },
          shapes: [
            {
              id: '2',
              name: 'Title 1',
              kind: 'placeholder',
              placeholder: { type: 'title' },
              geometry: {},
              text: { text: 'Office preview' },
            },
            {
              id: '3',
              name: 'Content Placeholder 2',
              kind: 'placeholder',
              placeholder: { index: '1' },
              geometry: {},
              text: { text: 'Uploaded once, preview current version' },
            },
          ],
        },
      ],
    });

    expect(scene.slides[0]!.nodes).toEqual([
      expect.objectContaining({
        id: '2',
        bounds: { left: 457200, top: 274638, width: 8229600, height: 1143000 },
        text: 'Office preview',
        textStyle: { fontSize: 44, color: '#111111' },
        textBody: { anchor: 'middle', insetLeft: 91440 },
      }),
      expect.objectContaining({
        id: '3',
        bounds: { left: 457200, top: 1600200, width: 8229600, height: 4525963 },
        text: 'Uploaded once, preview current version',
        textStyle: { fontSize: 32, color: '#222222' },
        textBody: { autoFit: 'normal', insetLeft: 91440 },
      }),
    ]);
    expect(scene.slides[0]!.source.shapes.map((shape) => shape.geometry)).toEqual([{}, {}]);
  });

  test('blocks unsupported edit commands without adding history', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-command-boundary-test',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Body',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 2000000, height: 500000 },
              text: { text: 'Body' },
            },
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 0, top: 1800000, width: 2000000, height: 800000 },
              table: { values: [['A1']], rowCount: 1, columnCount: 1 },
            },
          ],
        },
      ],
    });
    const session = new PptEditorSession(view);

    await expect(
      session.dispatch({
        kind: 'set-shape-text',
        slideIdOrIndex: 0,
        shapeIdOrName: 'Data',
        text: 'Nope',
      }),
    ).rejects.toThrow('not text-editable');
    expect(session.canUndo()).toBe(false);

    await session.destroy();
  });

  test('loads pptx bytes as a full snapshot replacement', async () => {
    const presentation = await Presentation.fromModel({
      id: 'load-test',
      slides: [{ id: '256', index: 0, shapes: [] }],
    });
    const bytes = await presentation.toZip();
    const view = await buildPptEditorViewFromModel({
      id: 'initial',
      slides: [
        {
          id: '257',
          index: 0,
          shapes: [
            {
              id: '3',
              kind: 'textBox',
              geometry: {},
              text: { text: 'Initial' },
            },
          ],
        },
      ],
    });
    const session = new PptEditorSession(view);

    await session.loadPresentation(bytes);

    expect(session.model().slides).toHaveLength(1);
    expect(session.model().slides[0]!.shapes).toHaveLength(0);
    expect(session.canUndo()).toBe(false);

    await session.destroy();
    await presentation.dispose();
  });

  test('applies table cell text commands through core', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'table-test',
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 0, top: 0, width: 2000000, height: 800000 },
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
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-table-cell-text',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Data',
      row: 1,
      column: 1,
      text: 'Final',
    });

    expect(session.model().slides[0]!.shapes[0]!.table?.values[1]![1]).toBe('Final');
    const verified = await session.verifyExport();
    expect(verified.slides[0]!.shapes[0]!.table?.values[1]![1]).toBe('Final');

    await session.undo();
    expect(session.model().slides[0]!.shapes[0]!.table?.values[1]![1]).toBe('B2');
    await session.redo();
    expect(session.model().slides[0]!.shapes[0]!.table?.values[1]![1]).toBe('Final');

    await session.destroy();
  });

  test('applies grouped table cell text commands through core', async () => {
    const view = await buildPptEditorViewFromPptx(await makeGroupedTablePptx());
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-table-cell-text',
      slideIdOrIndex: 0,
      shapeIdOrName: '6:15',
      row: 0,
      column: 1,
      text: 'Grouped Final',
    });

    const groupedTable = session.scene().slides[0]!.nodes.find((node) => node.id === '6:15');
    expect(groupedTable?.capabilities.editTableCells).toBe(true);
    expect(groupedTable?.table?.values[0]?.[1]).toBe('Grouped Final');
    const verified = await session.verifyExport();
    const verifiedGroupedTable = verified.slides[0]!.shapes.find(
      (shape) => shape.name === 'Group 1',
    )?.children?.find((shape) => shape.name === 'GroupedTable');
    expect(verifiedGroupedTable?.table?.values[0]?.[1]).toBe('Grouped Final');

    await session.destroy();
  });

  test('persists no-autofit text body behavior when resizing text boxes', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-resize-text-autofit-roundtrip',
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
              geometry: { left: 100000, top: 100000, width: 1800000, height: 500000 },
              text: {
                text: 'Resize without shrinking',
                body: { autoFit: 'normal', wrap: 'square' },
                style: { fontSize: 44 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = { autoFit: 'normal', wrap: 'square' };
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-shape-text-style',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      style: { fontSize: 44 },
    });
    session.scene().slides[0]!.nodes[0]!.textBody = { autoFit: 'normal', wrap: 'square' };
    await session.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      geometry: { left: 100000, top: 100000, width: 800000, height: 500000 },
      preserveTextSizeOnResize: true,
    });

    expect(session.scene().slides[0]!.nodes[0]!.textBody).toMatchObject({
      autoFit: 'none',
      wrap: 'square',
    });
    const verified = await session.verifyExport();
    const title = verified.slides[0]!.shapes.find((shape) => shape.name === 'Title')!;
    expect(title.text?.body).toMatchObject({ autoFit: 'none', wrap: 'square' });
    expect(title.text?.style?.fontSize).toBe(44);

    await session.destroy();
  });

  test('keeps shape-auto-fit text size when manually resizing text boxes', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-resize-shape-autofit-roundtrip',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'ShapeFitTitle',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1800000, height: 500000 },
              text: {
                text: 'Resize without shrinking',
                body: { autoFit: 'shape', wrap: 'square' },
                style: { fontSize: 44 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = { autoFit: 'shape', wrap: 'square' };
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-shape-text-style',
      slideIdOrIndex: 0,
      shapeIdOrName: 'ShapeFitTitle',
      style: { fontSize: 44 },
    });
    session.scene().slides[0]!.nodes[0]!.textBody = { autoFit: 'shape', wrap: 'square' };
    await session.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: 0,
      shapeIdOrName: 'ShapeFitTitle',
      geometry: { width: 800000, height: 500000 },
      preserveTextSizeOnResize: true,
    });

    expect(session.scene().slides[0]!.nodes[0]!.textBody).toMatchObject({
      autoFit: 'none',
      wrap: 'square',
    });
    const verified = await session.verifyExport();
    const title = verified.slides[0]!.shapes.find((shape) => shape.name === 'ShapeFitTitle')!;
    expect(title.text?.body).toMatchObject({ autoFit: 'none', wrap: 'square' });
    expect(title.text?.style?.fontSize).toBe(44);

    await session.destroy();
  });

  test('keeps placeholder text size when manually resizing title placeholders', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-resize-placeholder-autofit-roundtrip',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title Placeholder',
              kind: 'placeholder',
              placeholder: { type: 'title' },
              geometry: { left: 100000, top: 100000, width: 2200000, height: 600000 },
              text: {
                text: 'Resize placeholder without shrinking',
                body: { autoFit: 'normal', wrap: 'square' },
                style: { fontSize: 44 },
              },
            },
          ],
        },
      ],
    });
    view.scene.slides[0]!.nodes[0]!.textBody = { autoFit: 'normal', wrap: 'square' };
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-shape-text-style',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title Placeholder',
      style: { fontSize: 44 },
    });
    session.scene().slides[0]!.nodes[0]!.textBody = { autoFit: 'normal', wrap: 'square' };
    await session.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title Placeholder',
      geometry: { width: 900000, height: 600000 },
      preserveTextSizeOnResize: true,
    });

    expect(session.scene().slides[0]!.nodes[0]!.textBody).toMatchObject({
      autoFit: 'none',
      wrap: 'square',
    });
    const verified = await session.verifyExport();
    const title = verified.slides[0]!.shapes.find((shape) => shape.name === 'Title Placeholder')!;
    expect(title.text?.body).toMatchObject({ autoFit: 'none', wrap: 'square' });
    expect(title.text?.style?.fontSize).toBe(44);

    await session.destroy();
  });

  test('adds wrapping body properties when resized text boxes did not expose text body data', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'editor-resize-missing-text-body-roundtrip',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'PlainTitle',
              kind: 'textBox',
              geometry: { left: 100000, top: 100000, width: 1800000, height: 500000 },
              text: {
                text: 'Resize without shrinking',
                style: { fontSize: 44 },
              },
            },
          ],
        },
      ],
    });
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-shape-text-style',
      slideIdOrIndex: 0,
      shapeIdOrName: 'PlainTitle',
      style: { fontSize: 44 },
    });
    await session.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: 0,
      shapeIdOrName: 'PlainTitle',
      geometry: { width: 800000, height: 500000 },
      preserveTextSizeOnResize: true,
    });

    expect(session.scene().slides[0]!.nodes[0]!.textBody).toMatchObject({
      autoFit: 'none',
      wrap: 'square',
    });
    const verified = await session.verifyExport();
    const title = verified.slides[0]!.shapes.find((shape) => shape.name === 'PlainTitle')!;
    expect(title.text?.body).toMatchObject({ autoFit: 'none', wrap: 'square' });
    expect(title.text?.style?.fontSize).toBe(44);

    await session.destroy();
  });

  test('persists combined text, style, geometry, and table edits in one export', async () => {
    const view = await buildPptEditorViewFromModel({
      id: 'combined-edit-roundtrip-test',
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
              geometry: { left: 100000, top: 200000, width: 2000000, height: 500000 },
              text: { text: 'Draft' },
              style: { fillColor: '#FFFFFF', lineColor: '#000000', lineWidth: 9525 },
            },
            {
              id: '4',
              name: 'Data',
              kind: 'table',
              geometry: { left: 100000, top: 900000, width: 2000000, height: 800000 },
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
    const session = new PptEditorSession(view);

    await session.dispatch({
      kind: 'set-shape-text',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      text: 'P2.2 TEXT ROUNDTRIP',
    });
    await session.dispatch({
      kind: 'set-shape-text-style',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      style: { fontSize: 30, color: '#336699', alignment: 'center' },
    });
    await session.dispatch({
      kind: 'set-shape-style',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      style: { fillColor: '#FFCC00', lineColor: '#2255AA', lineWidth: 19050 },
    });
    await session.dispatch({
      kind: 'set-shape-geometry',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Title',
      geometry: { left: 300000, top: 400000, width: 2400000, height: 700000 },
    });
    await session.dispatch({
      kind: 'set-table-cell-text',
      slideIdOrIndex: 0,
      shapeIdOrName: 'Data',
      row: 1,
      column: 0,
      text: 'P2.2 TABLE ROUNDTRIP',
    });

    const verified = await session.verifyExport();
    const title = verified.slides[0]!.shapes.find((shape) => shape.name === 'Title')!;
    const table = verified.slides[0]!.shapes.find((shape) => shape.name === 'Data')!;
    expect(title.text?.text).toBe('P2.2 TEXT ROUNDTRIP');
    expect(title.text?.style).toMatchObject({
      fontSize: 30,
      color: '#336699',
      alignment: 'center',
    });
    expect(title.style).toMatchObject({
      fillColor: '#FFCC00',
      lineColor: '#2255AA',
      lineWidth: 19050,
    });
    expect(title.geometry).toMatchObject({
      left: 300000,
      top: 400000,
      width: 2400000,
      height: 700000,
    });
    expect(table.table?.values[1]![0]).toBe('P2.2 TABLE ROUNDTRIP');
    expect(session.canUndo()).toBe(true);

    await session.destroy();
  });
});

async function makeGroupedTablePptx(): Promise<Uint8Array> {
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
