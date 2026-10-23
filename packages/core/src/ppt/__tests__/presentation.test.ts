import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
  bytesEqual,
  inspectSlideMotion,
  isAdvancedPptPart,
  packageAdvancedCounts,
  selectedZipEntryBytes,
  summarizeAdvancedShapes,
  summarizePptFixture,
  zipEntryBytes,
} from '../../../scripts/ppt-fixture-summary.mjs';
import { Presentation, type PptPresentationModel, type PptShapeModel } from '../index.js';

describe('Presentation P0 live-object SDK', () => {
  it('creates a pptx package from a model and reloads generated text and table shapes', async () => {
    const presentation = await Presentation.fromModel({
      id: 'model-deck',
      slides: [
        {
          id: '256',
          index: 0,
          title: 'Generated',
          shapes: [
            {
              id: '2',
              name: 'GeneratedTitle',
              kind: 'textBox',
              geometry: { left: 100, top: 200, width: 3000000, height: 600000 },
              text: { text: 'Generated Title' },
            },
            {
              id: '3',
              name: 'GeneratedTable',
              kind: 'table',
              geometry: { left: 100, top: 900000, width: 3000000, height: 900000 },
              table: {
                rowCount: 2,
                columnCount: 2,
                values: [
                  ['Q1', '10'],
                  ['Q2', '20'],
                ],
              },
            },
          ],
        },
      ],
    });

    const bytes = await presentation.toZip();
    const reloaded = await Presentation.fromZip(bytes);
    const slide = reloaded.slides.getItem(0);

    expect(await slide.shapes.getItem('GeneratedTitle').text.getText()).toBe('Generated Title');
    expect(await slide.shapes.getItem('GeneratedTable').table.values()).toEqual([
      ['Q1', '10'],
      ['Q2', '20'],
    ]);

    await slide.shapes.getItem('GeneratedTitle').text.setText('Edited Generated Title');
    const edited = await Presentation.fromZip(await reloaded.toZip());
    expect(await edited.slides.getItem(0).shapes.getItem('GeneratedTitle').text.getText()).toBe(
      'Edited Generated Title',
    );
  });

  it('createBlank exports a one-slide empty pptx package', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'blank' });
    const model = await presentation.toModel();
    expect(model.slides).toHaveLength(1);
    expect(await presentation.slides.getCount()).toBe(1);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getCount()).toBe(1);
  });

  it('writes notes and comments on slides without existing parts', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'notes-comments' });
    const slide = presentation.slides.getItem(0);

    await slide.notes.setText('Generated speaker notes');
    const added = await slide.comments.add({
      text: 'Generated comment',
      author: 'SDK Reviewer',
      initials: 'SR',
      date: '2026-08-25T10:00:00Z',
      x: 10,
      y: 20,
    });

    expect(added).toMatchObject({
      id: '1',
      authorId: '0',
      text: 'Generated comment',
      x: 10,
      y: 20,
    });

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedSlide = reloaded.slides.getItem(0);
    expect(await reloadedSlide.notes.getText()).toBe('Generated speaker notes');
    expect(await reloadedSlide.comments.items()).toEqual([
      {
        id: '1',
        authorId: '0',
        author: 'SDK Reviewer',
        initials: 'SR',
        date: '2026-08-25T10:00:00Z',
        text: 'Generated comment',
        x: 10,
        y: 20,
      },
    ]);
  });

  it('resolves inline mode and rejects planned worker execution explicitly', async () => {
    const resolutions: unknown[] = [];
    const presentation = await Presentation.createBlank({
      presentationId: 'runtime-mode',
      mode: 'auto',
      onModeResolved: (decision, presentationId) => resolutions.push({ decision, presentationId }),
    });

    expect(presentation.mode).toBe('inline');
    expect(presentation.modeSource).toBe('auto/inline');
    expect(resolutions).toEqual([
      {
        decision: { mode: 'inline', source: 'auto/inline' },
        presentationId: 'runtime-mode',
      },
    ]);

    await expect(Presentation.createBlank({ mode: 'worker' })).rejects.toThrow(
      'worker mode is planned',
    );
    await expect(presentation.executeCode()).rejects.toThrow('executeCode is planned');
  });

  it('rejects root operations after dispose', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'disposed' });

    await presentation.dispose();

    await expect(presentation.toModel()).rejects.toThrow('disposed');
    await expect(presentation.properties()).rejects.toThrow('disposed');
    await expect(presentation.toZip()).rejects.toThrow('disposed');
    await expect(presentation.findText('anything')).rejects.toThrow('disposed');
    await expect(presentation.replaceText('anything', 'else')).rejects.toThrow('disposed');
    await expect(presentation.exportImages()).rejects.toThrow('disposed');
    await expect(presentation.getMutations()).rejects.toThrow('disposed');
    await expect(presentation.clearMutations()).rejects.toThrow('disposed');
  });

  it('adds, copies, moves, and deletes slides while preserving package reloadability', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'slide-ops' });

    const added = await presentation.slides.add({ title: 'Added Slide' });
    expect(await added.title()).toBe('Added Slide');
    expect(await presentation.slides.getCount()).toBe(2);

    const copied = await presentation.slides.copy(await added.id());
    expect(await copied.title()).toBe('Added Slide');
    expect(await presentation.slides.getCount()).toBe(3);

    await copied.moveTo(0);
    expect(await presentation.slides.getItem(0).title()).toBe('Added Slide');

    await presentation.slides.getItem(1).delete();
    expect(await presentation.slides.getCount()).toBe(2);

    const duplicated = await presentation.slides.getItem(0).duplicate({ index: 'end' });
    expect(await duplicated.title()).toBe('Added Slide');
    expect(await presentation.slides.getCount()).toBe(3);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getCount()).toBe(3);
    expect(await reloaded.slides.getItem(0).title()).toBe('Added Slide');
  });

  it('rejects invalid slide operations without corrupting the package', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'slide-guards' });

    await expect(presentation.slides.getItemOrNullObject('missing-slide')).resolves.toBeNull();
    expect(() => presentation.slides.getItem(99)).toThrow('slide not found');
    await expect(presentation.slides.deleteItem(0)).rejects.toThrow(
      'cannot delete the final slide',
    );

    const added = await presentation.slides.add({ title: 'Safe Slide' });
    await presentation.clearMutations();

    await expect(added.moveTo(-10)).resolves.toBeUndefined();
    expect(await added.index()).toBe(0);
    await expect(presentation.slides.move('missing-slide', 0)).rejects.toThrow('slide not found');
    await expect(presentation.slides.deleteItem('missing-slide')).rejects.toThrow(
      'slide not found',
    );
    await expect(presentation.getMutations()).resolves.toEqual([]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getCount()).toBe(2);
  });

  it('finds and replaces text across slides, shapes, and table cells', async () => {
    const presentation = await Presentation.fromModel({
      id: 'search-deck',
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'IntroText',
              kind: 'textBox',
              geometry: {},
              text: { text: 'Hello Coco. Coco rocks.' },
            },
            {
              id: '3',
              name: 'SearchTable',
              kind: 'table',
              geometry: {},
              table: {
                rowCount: 1,
                columnCount: 2,
                values: [['Coco table', 'Other']],
              },
            },
          ],
        },
        {
          id: '257',
          index: 1,
          shapes: [
            {
              id: '2',
              name: 'SecondText',
              kind: 'textBox',
              geometry: {},
              text: { text: 'Second coco' },
            },
          ],
        },
      ],
    });

    const matches = await presentation.findText('coco');
    expect(
      matches.map((match) => [
        match.slideIndex,
        match.shapeName,
        match.source,
        match.row,
        match.column,
      ]),
    ).toEqual([
      [0, 'IntroText', 'shapeText', undefined, undefined],
      [0, 'IntroText', 'shapeText', undefined, undefined],
      [0, 'SearchTable', 'tableCell', 0, 0],
      [1, 'SecondText', 'shapeText', undefined, undefined],
    ]);

    const slideMatches = await presentation.slides
      .getItem(0)
      .findText('Coco', { caseSensitive: true });
    expect(slideMatches).toHaveLength(3);

    await expect(presentation.replaceText('coco', 'SDK')).resolves.toEqual({ count: 4 });
    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getItem(0).shapes.getItem('IntroText').text.getText()).toBe(
      'Hello SDK. SDK rocks.',
    );
    expect(await reloaded.slides.getItem(0).shapes.getItem('SearchTable').table.values()).toEqual([
      ['SDK table', 'Other'],
    ]);
    expect(await reloaded.slides.getItem(1).shapes.getItem('SecondText').text.getText()).toBe(
      'Second SDK',
    );
  });

  it('returns computed presentation and slide properties', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'props-deck' });
    const slide = presentation.slides.getItem(0);
    await slide.shapes.addTextBox('Hello', { name: 'Text' });
    await slide.shapes.addTable(1, 2, { name: 'Table', values: [['A', 'B']] });
    await slide.shapes.addImage({ bytes: pngBytes(), mimeType: 'image/png' }, { name: 'Picture' });

    await expect(presentation.properties()).resolves.toMatchObject({
      presentationId: 'props-deck',
      slideCount: 1,
      shapeCount: 3,
      textShapeCount: 1,
      tableCount: 1,
      pictureCount: 1,
      mediaCount: 1,
    });
    await expect(slide.properties()).resolves.toMatchObject({
      id: '256',
      index: 0,
      title: 'Hello',
      shapeCount: 3,
      textShapeCount: 1,
      tableCount: 1,
      pictureCount: 1,
    });
  });

  it('adds and deletes text boxes and tables while preserving package reloadability', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'shape-ops' });
    const slide = presentation.slides.getItem(0);

    const textBox = await slide.shapes.addTextBox('New Text', {
      name: 'NewTextBox',
      geometry: { left: 123, top: 456, width: 3000, height: 4000 },
    });
    const table = await slide.shapes.addTable(2, 2, {
      name: 'NewTable',
      values: [
        ['A', 'B'],
        ['C', 'D'],
      ],
    });

    expect(await slide.shapes.getCount()).toBe(2);
    expect(await textBox.text.getText()).toBe('New Text');
    expect(await table.table.values()).toEqual([
      ['A', 'B'],
      ['C', 'D'],
    ]);

    await table.table.getCell(1, 1).setText('Edited');
    await textBox.delete();
    expect(await slide.shapes.getCount()).toBe(1);

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'shape.add',
      'table.cell.text.set',
      'shape.delete',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedSlide = reloaded.slides.getItem(0);
    expect(await reloadedSlide.shapes.getCount()).toBe(1);
    expect(await reloadedSlide.shapes.getItem('NewTable').table.values()).toEqual([
      ['A', 'B'],
      ['C', 'Edited'],
    ]);
    expect(await reloadedSlide.shapes.getItemOrNullObject('NewTextBox')).toBeNull();

    await reloadedSlide.shapes.deleteItem('NewTable');
    expect(await reloadedSlide.shapes.getCount()).toBe(0);
  });

  it('rejects ambiguous and out-of-range shape lookups predictably', async () => {
    const presentation = await Presentation.fromModel({
      id: 'shape-lookup-guards',
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Duplicated',
              kind: 'textBox',
              geometry: {},
              text: { text: 'First' },
            },
            {
              id: '3',
              name: 'Duplicated',
              kind: 'textBox',
              geometry: {},
              text: { text: 'Second' },
            },
          ],
        },
      ],
    });
    const shapes = presentation.slides.getItem(0).shapes;

    expect(() => shapes.getItem('Duplicated')).toThrow('ambiguous');
    await expect(shapes.getItemOrNullObject('Duplicated')).resolves.toBeNull();
    await expect(shapes.getItemOrNullObject('missing-shape')).resolves.toBeNull();
    expect(() => shapes.getItemAt(2)).toThrow('shape index out of range');
    await expect(shapes.getItemAt(0).text.getText()).resolves.toBe('First');
    await expect(shapes.getItemAt(1).text.getText()).resolves.toBe('Second');
  });

  it('rejects unsupported shape-domain operations without recording mutations', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'shape-domain-guards' });
    const slide = presentation.slides.getItem(0);
    const text = await slide.shapes.addTextBox('Text only', { name: 'TextOnly' });
    const rect = await slide.shapes.addShape('rect', { name: 'RectOnly' });

    await presentation.clearMutations();

    await expect(text.image.style()).resolves.toEqual({});
    await expect(text.image.setOpacity(0.5)).rejects.toThrow('no image style surface');
    await expect(text.image.bytes()).rejects.toThrow('no image bytes');
    await expect(text.image.replace(pngBytes())).rejects.toThrow('no replaceable image');
    await expect(text.table.values()).rejects.toThrow('is not a table');
    await expect(rect.text.getText()).rejects.toThrow('no text body');
    await expect(rect.text.setText('Nope')).rejects.toThrow('no text body');

    expect(await presentation.getMutations()).toEqual([]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getItem(0).shapes.getCount()).toBe(2);
  });

  it('renames shapes and changes z-order in shape tree order', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'shape-z-order' });
    const slide = presentation.slides.getItem(0);
    const bottom = await slide.shapes.addTextBox('Bottom', { name: 'Bottom' });
    const middle = await slide.shapes.addTextBox('Middle', { name: 'Middle' });
    const top = await slide.shapes.addTextBox('Top', { name: 'Top' });

    await middle.setName('RenamedMiddle');
    await bottom.bringToFront();
    await top.sendToBack();

    const orderedNames = await Promise.all(
      (await slide.shapes.items()).map((shape) => shape.name()),
    );
    expect(orderedNames).toEqual(['Top', 'RenamedMiddle', 'Bottom']);

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'shape.add',
      'shape.add',
      'shape.name.set',
      'shape.zOrder.set',
      'shape.zOrder.set',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedNames = await Promise.all(
      (await reloaded.slides.getItem(0).shapes.items()).map((shape) => shape.name()),
    );
    expect(reloadedNames).toEqual(['Top', 'RenamedMiddle', 'Bottom']);
  });

  it('duplicates text, table, and image shapes with independent shape ids', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'shape-duplicate' });
    const slide = presentation.slides.getItem(0);
    const text = await slide.shapes.addTextBox('Copy me', { name: 'SourceText' });
    const table = await slide.shapes.addTable(1, 2, {
      name: 'SourceTable',
      values: [['A', 'B']],
    });
    const image = await slide.shapes.addImage(
      { bytes: pngBytes(), mimeType: 'image/png' },
      { name: 'SourceImage' },
    );

    const textCopy = await text.duplicate({
      name: 'TextCopy',
      geometry: { left: 111, top: 222 },
    });
    const tableCopy = await table.duplicate({ name: 'TableCopy' });
    const imageCopy = await image.duplicate({ name: 'ImageCopy' });

    expect(await textCopy.id()).not.toBe(await text.id());
    expect(await tableCopy.id()).not.toBe(await table.id());
    expect(await imageCopy.id()).not.toBe(await image.id());
    expect(await textCopy.text.getText()).toBe('Copy me');
    expect(await textCopy.geometry()).toMatchObject({ left: 111, top: 222 });
    expect(await tableCopy.table.values()).toEqual([['A', 'B']]);
    expect(await imageCopy.image.bytes()).toEqual(pngBytes());

    const replacement = new Uint8Array([137, 80, 78, 71, 9, 9, 9, 9]);
    await imageCopy.image.replace(replacement);
    expect(await image.image.bytes()).toEqual(pngBytes());
    expect(await imageCopy.image.bytes()).toEqual(replacement);

    expect((await presentation.getMutations()).map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'shape.add',
      'shape.add',
      'shape.add',
      'shape.add',
      'shape.add',
      'shape.image.replace',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedSlide = reloaded.slides.getItem(0);
    expect(await reloadedSlide.shapes.getItem('TextCopy').text.getText()).toBe('Copy me');
    expect(await reloadedSlide.shapes.getItem('TableCopy').table.values()).toEqual([['A', 'B']]);
    expect(await reloadedSlide.shapes.getItem('SourceImage').image.bytes()).toEqual(pngBytes());
    expect(await reloadedSlide.shapes.getItem('ImageCopy').image.bytes()).toEqual(replacement);
  });

  it('adds images with media parts and slide relationships', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'image-authoring' });
    const slide = presentation.slides.getItem(0);
    const image = await slide.shapes.addImage(
      {
        bytes: pngBytes(),
        mimeType: 'image/png',
        filename: 'tiny.png',
      },
      {
        name: 'AddedImage',
        geometry: { left: 10, top: 20, width: 30, height: 40 },
      },
    );

    expect(await image.kind()).toBe('picture');
    expect(await image.image.getInfo()).toMatchObject({
      targetPath: 'ppt/media/image1.png',
      contentType: 'image/png',
      filename: 'image1.png',
      byteLength: pngBytes().byteLength,
    });

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual(['shape.add']);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const added = reloaded.slides.getItem(0).shapes.getItem('AddedImage');
    expect(await added.kind()).toBe('picture');
    expect(await added.geometry()).toMatchObject({ left: 10, top: 20, width: 30, height: 40 });
    expect(await added.image.getInfo()).toMatchObject({
      targetPath: 'ppt/media/image1.png',
      contentType: 'image/png',
      byteLength: pngBytes().byteLength,
    });
  });

  it('reads Office SVG blip extension images as picture media', async () => {
    const presentation = await Presentation.fromZip(await makeSvgBlipFixturePptx(), {
      presentationId: 'svg-blip-fixture',
    });
    const slide = presentation.slides.getItem(0);
    const picture = slide.shapes.getItem('SvgLogo');

    await expect(picture.kind()).resolves.toBe('picture');
    await expect(picture.image.getInfo()).resolves.toMatchObject({
      relationshipId: 'rId2',
      targetPath: 'ppt/media/logo.svg',
      contentType: 'image/svg+xml',
      filename: 'logo.svg',
    });
    await expect(picture.image.bytes()).resolves.toEqual(svgBytes());
    await expect(presentation.exportImages()).resolves.toEqual([
      {
        path: 'ppt/media/logo.svg',
        contentType: 'image/svg+xml',
        byteLength: svgBytes().byteLength,
        bytes: svgBytes(),
      },
    ]);
  });

  it('classifies line-styled custom geometry as editable geometric shapes', async () => {
    const presentation = await Presentation.fromZip(await makeCustomGeometryFixturePptx(), {
      presentationId: 'custom-geometry-fixture',
    });
    const shape = presentation.slides.getItem(0).shapes.getItem('FreeformCircle');
    const model = await shape.toModel();
    const commands = model.raw?.customGeometry?.paths[0]?.commands ?? [];

    await expect(shape.kind()).resolves.toBe('geometricShape');
    expect(commands).toHaveLength(6);
    expect(commands[0]).toMatchObject({ kind: 'moveTo', x: 1000, y: 500 });
    expect(commands[1]).toMatchObject({
      kind: 'cubicBezierTo',
      x1: 1000,
      y1: 776,
      x2: 776,
      y2: 1000,
      x: 500,
      y: 1000,
    });
    expect(commands[5]).toMatchObject({ kind: 'close' });
    expect(model.style).toMatchObject({
      lineColor: '#133E82',
      lineWidth: 101600,
    });
  });

  it('preserves custom geometry arc commands', async () => {
    const presentation = await Presentation.fromZip(await makeCustomGeometryFixturePptx(), {
      presentationId: 'custom-geometry-arc-fixture',
    });
    const shape = presentation.slides.getItem(0).shapes.getItem('ArcShape');
    const model = await shape.toModel();
    const commands = model.raw?.customGeometry?.paths[0]?.commands ?? [];

    expect(commands).toEqual([
      { kind: 'moveTo', x: 1000, y: 500 },
      { kind: 'arcTo', wR: 500, hR: 500, stAng: 0, swAng: 5400000 },
    ]);
  });

  it('exports image bytes from presentation and picture shapes', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'image-export' });
    const slide = presentation.slides.getItem(0);
    const picture = await slide.shapes.addImage(
      { bytes: pngBytes(), mimeType: 'image/png' },
      { name: 'ExportedImage' },
    );

    await expect(picture.image.bytes()).resolves.toEqual(pngBytes());
    const exported = await presentation.exportImages();
    expect(exported).toHaveLength(1);
    expect(exported[0]).toMatchObject({
      path: 'ppt/media/image1.png',
      contentType: 'image/png',
      byteLength: pngBytes().byteLength,
    });
    expect(exported[0]!.bytes).toEqual(pngBytes());

    const replacement = new Uint8Array([137, 80, 78, 71, 1, 2, 3, 4]);
    await picture.image.replace(replacement);
    await expect(picture.image.bytes()).resolves.toEqual(replacement);
    expect((await presentation.exportImages())[0]!.bytes).toEqual(replacement);

    await picture.image.replace({
      bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xe0]),
      mimeType: 'image/jpeg',
      filename: 'replacement.jpeg',
    });
    expect(await picture.image.getInfo()).toMatchObject({
      targetPath: 'ppt/media/image1.png',
      contentType: 'image/png',
      filename: 'image1.png',
      byteLength: 4,
    });
  });

  it('sets picture crop and opacity style', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'image-style' });
    const slide = presentation.slides.getItem(0);
    const picture = await slide.shapes.addImage(
      { bytes: pngBytes(), mimeType: 'image/png' },
      { name: 'StyledPicture' },
    );

    await presentation.clearMutations();

    await picture.image.setStyle({
      crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 },
      fillRect: { left: 0.01, top: 0.02, right: 0.03, bottom: 0.04 },
      opacity: 0.65,
    });
    await picture.image.setCrop({ left: 0.15 });
    await picture.image.setFillRect({ right: 0.08 });
    await picture.image.setOpacity(0.5);

    expect(await picture.image.style()).toEqual({
      crop: { left: 0.15, top: 0.2, right: 0.3, bottom: 0.4 },
      fillRect: { left: 0.01, top: 0.02, right: 0.08, bottom: 0.04 },
      opacity: 0.5,
    });
    expect((await picture.toModel()).image).toMatchObject({
      crop: { left: 0.15, top: 0.2, right: 0.3, bottom: 0.4 },
      fillRect: { left: 0.01, top: 0.02, right: 0.08, bottom: 0.04 },
      opacity: 0.5,
    });
    expect((await presentation.getMutations()).map((mutation) => mutation.type)).toEqual([
      'shape.image.style.set',
      'shape.image.style.set',
      'shape.image.style.set',
      'shape.image.style.set',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getItem(0).shapes.getItem('StyledPicture').image.style()).toEqual({
      crop: { left: 0.15, top: 0.2, right: 0.3, bottom: 0.4 },
      fillRect: { left: 0.01, top: 0.02, right: 0.08, bottom: 0.04 },
      opacity: 0.5,
    });
  });

  it('adds preset geometric shapes and lines', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'shape-authoring' });
    const slide = presentation.slides.getItem(0);

    const rect = await slide.shapes.addShape('rect', {
      name: 'AddedRect',
      geometry: { left: 100, top: 200, width: 300, height: 400 },
    });
    const ellipse = await slide.shapes.addShape('ellipse', {
      name: 'AddedEllipse',
      geometry: { left: 500, top: 600, width: 700, height: 800 },
    });
    const line = await slide.shapes.addShape('line', {
      name: 'AddedLine',
      geometry: { left: 900, top: 1000, width: 1100, height: 0 },
    });

    expect(await rect.kind()).toBe('geometricShape');
    expect(await ellipse.kind()).toBe('geometricShape');
    expect(await line.kind()).toBe('line');

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedSlide = reloaded.slides.getItem(0);
    expect(await reloadedSlide.shapes.getItem('AddedRect').kind()).toBe('geometricShape');
    expect(await reloadedSlide.shapes.getItem('AddedRect').geometry()).toMatchObject({
      left: 100,
      top: 200,
      width: 300,
      height: 400,
    });
    expect(await reloadedSlide.shapes.getItem('AddedEllipse').kind()).toBe('geometricShape');
    expect(await reloadedSlide.shapes.getItem('AddedLine').kind()).toBe('line');

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'shape.add',
      'shape.add',
    ]);
  });

  it('sets and reads basic shape and text styles', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'style-editing' });
    const slide = presentation.slides.getItem(0);
    const rect = await slide.shapes.addShape('rect', { name: 'StyledRect' });
    const text = await slide.shapes.addTextBox('Styled Text', { name: 'StyledText' });

    await rect.setStyle({
      fillColor: '#ff0011',
      opacity: 0.4,
      lineColor: '00aaee',
      lineOpacity: 0.65,
      lineWidth: 25400,
      shadow: {
        color: '#223344',
        opacity: 0.35,
        blur: 12700,
        distance: 25400,
        direction: 45,
      },
    });
    await rect.setStyle({ opacity: 0.55, lineOpacity: 0.75 });
    await text.text.setStyle({
      fontSize: 18,
      fontFamily: 'Arial',
      color: '#123abc',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      baseline: 0.3,
      highlight: '#fff200',
      alignment: 'center',
      lineHeight: 1.25,
      characterSpacing: 2.4,
    });

    expect(await rect.style()).toEqual({
      fillColor: '#FF0011',
      opacity: 0.55,
      lineColor: '#00AAEE',
      lineOpacity: 0.75,
      lineWidth: 25400,
      shadow: {
        color: '#223344',
        opacity: 0.35,
        blur: 12700,
        distance: 25400,
        direction: 45,
      },
    });
    expect(await text.text.style()).toEqual({
      fontSize: 18,
      fontFamily: 'Arial',
      color: '#123ABC',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      baseline: 0.3,
      highlight: '#FFF200',
      alignment: 'center',
      lineHeight: 1.25,
      characterSpacing: 2.4,
    });

    await text.text.setBold(false);
    await text.text.setItalic(false);
    await text.text.setAlignment('right');

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'shape.add',
      'shape.style.set',
      'shape.style.set',
      'shape.text.style.set',
      'shape.text.style.set',
      'shape.text.style.set',
      'shape.text.style.set',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedRect = reloaded.slides.getItem(0).shapes.getItem('StyledRect');
    const reloadedText = reloaded.slides.getItem(0).shapes.getItem('StyledText');
    expect(await reloadedRect.style()).toEqual({
      fillColor: '#FF0011',
      opacity: 0.55,
      lineColor: '#00AAEE',
      lineOpacity: 0.75,
      lineWidth: 25400,
      shadow: {
        color: '#223344',
        opacity: 0.35,
        blur: 12700,
        distance: 25400,
        direction: 45,
      },
    });
    expect(await reloadedText.text.style()).toEqual({
      fontSize: 18,
      fontFamily: 'Arial',
      color: '#123ABC',
      bold: false,
      italic: false,
      underline: true,
      strike: true,
      baseline: 0.3,
      highlight: '#FFF200',
      alignment: 'right',
      lineHeight: 1.25,
      characterSpacing: 2.4,
    });
  });

  it('writes structured text paragraphs and run styles', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'structured-text' });
    const slide = presentation.slides.getItem(0);
    const shape = await slide.shapes.addTextBox('Initial', { name: 'StructuredText' });

    await shape.text.setParagraphs([
      {
        style: {
          alignment: 'center',
          bullet: '•',
          bulletColor: '#22AA88',
          bulletFontFamily: 'Arial',
          bulletSize: 0.72,
          lineHeight: 1.4,
          marginLeft: 342900,
          textIndent: -171450,
          spaceBefore: 1270,
          spaceAfter: 2540,
        },
        runs: [
          {
            text: 'Hello ',
            style: {
              fontSize: 20,
              fontFamily: 'Arial',
              color: '#112233',
              bold: true,
              underline: true,
              baseline: 0.2,
              highlight: '#ffee66',
              characterSpacing: 1.2,
            },
          },
          { text: 'Coco', style: { italic: true } },
        ],
      },
      {
        text: 'Second line',
        style: { alignment: 'right' },
      },
    ]);
    await shape.text.setRunStyle(1, {
      color: '#445566',
      bold: false,
      underline: false,
      strike: true,
      baseline: -0.15,
      highlight: '#ccffaa',
      characterSpacing: 3,
    });

    expect(await shape.text.paragraphs()).toEqual([
      {
        text: 'Hello Coco',
        style: {
          alignment: 'center',
          bullet: '•',
          bulletColor: '#22AA88',
          bulletFontFamily: 'Arial',
          bulletSize: 0.72,
          lineHeight: 1.4,
          marginLeft: 342900,
          textIndent: -171450,
          spaceBefore: 1270,
          spaceAfter: 2540,
        },
        runs: [
          {
            text: 'Hello ',
            style: {
              fontSize: 20,
              fontFamily: 'Arial',
              color: '#112233',
              bold: true,
              underline: true,
              baseline: 0.2,
              highlight: '#FFEE66',
              characterSpacing: 1.2,
              alignment: 'center',
            },
          },
          {
            text: 'Coco',
            style: {
              color: '#445566',
              bold: false,
              italic: true,
              underline: false,
              strike: true,
              baseline: -0.15,
              highlight: '#CCFFAA',
              characterSpacing: 3,
              alignment: 'center',
            },
          },
        ],
      },
      {
        text: 'Second line',
        style: { alignment: 'right' },
        runs: [{ text: 'Second line', style: { alignment: 'right' } }],
      },
    ]);

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'shape.text.paragraphs.set',
      'shape.text.run.style.set',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(
      await reloaded.slides.getItem(0).shapes.getItem('StructuredText').text.paragraphs(),
    ).toEqual(await shape.text.paragraphs());
  });

  it('round-trips explicit text line breaks as DrawingML breaks', async () => {
    const presentation = await Presentation.createBlank({
      presentationId: 'text-line-breaks',
    });
    const slide = presentation.slides.getItem(0);
    const shape = await slide.shapes.addTextBox('', { name: 'LineBreakText' });

    await shape.text.setParagraphs([
      {
        runs: [
          {
            text: 'HOW TO USE THIS\nPRESENTATION',
            style: { fontSize: 65, fontFamily: 'Josefin Sans', bold: true },
          },
        ],
      },
    ]);

    const output = await presentation.toZip();
    const zip = await JSZip.loadAsync(output);
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    expect(slideXml).toContain('<a:br');
    expect(slideXml).toContain('<a:t>HOW TO USE THIS</a:t>');
    expect(slideXml).toContain('<a:t>PRESENTATION</a:t>');

    const reloaded = await Presentation.fromZip(output);
    await expect(
      reloaded.slides.getItem(0).shapes.getItem('LineBreakText').text.paragraphs(),
    ).resolves.toEqual([
      {
        text: 'HOW TO USE THIS\nPRESENTATION',
        runs: [
          {
            text: 'HOW TO USE THIS',
            style: { fontSize: 65, fontFamily: 'Josefin Sans', bold: true },
          },
          { text: '\n' },
          {
            text: 'PRESENTATION',
            style: { fontSize: 65, fontFamily: 'Josefin Sans', bold: true },
          },
        ],
      },
    ]);
  });

  it('preserves first paragraph and run style when replacing rich text with plain text', async () => {
    const presentation = await Presentation.createBlank({
      presentationId: 'preserve-rich-text-style',
    });
    const slide = presentation.slides.getItem(0);
    const shape = await slide.shapes.addTextBox('Initial', { name: 'PreserveStyleText' });

    await shape.text.setParagraphs([
      {
        style: { alignment: 'center', lineHeight: 1.3 },
        runs: [
          {
            text: '未来',
            style: {
              fontSize: 44,
              fontFamily: 'Microsoft YaHei',
              color: '#112233',
              bold: true,
            },
          },
          {
            text: '可期',
            style: { color: '#4F73C9', italic: true },
          },
        ],
      },
      {
        text: '副标题',
        style: { alignment: 'right' },
      },
    ]);

    await shape.text.setText('继续攀登', { preserveFirstRunStyle: true });

    expect(await shape.text.paragraphs()).toEqual([
      {
        text: '继续攀登',
        style: { alignment: 'center', lineHeight: 1.3 },
        runs: [
          {
            text: '继续攀登',
            style: {
              fontSize: 44,
              fontFamily: 'Microsoft YaHei',
              color: '#112233',
              bold: true,
              alignment: 'center',
            },
          },
        ],
      },
    ]);
    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(
      await reloaded.slides.getItem(0).shapes.getItem('PreserveStyleText').text.paragraphs(),
    ).toEqual(await shape.text.paragraphs());
  });

  it('writes text body auto-fit properties for resized text boxes', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'text-body-autofit' });
    const slide = presentation.slides.getItem(0);
    const shape = await slide.shapes.addTextBox('Resize without shrinking', {
      name: 'BodyFitText',
    });

    await presentation.clearMutations();
    await shape.text.setBodyProperties({ autoFit: 'none', wrap: 'square' });

    expect((await shape.text.body()).autoFit).toBe('none');
    expect((await shape.text.body()).wrap).toBe('square');
    expect(await presentation.getMutations()).toEqual([
      {
        type: 'shape.text.body.set',
        slideId: '256',
        shapeId: '2',
        body: { autoFit: 'none', wrap: 'square' },
      },
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(
      await reloaded.slides.getItem(0).shapes.getItem('BodyFitText').text.body(),
    ).toMatchObject({
      autoFit: 'none',
      wrap: 'square',
    });
  });

  it('supports agent-friendly text style aliases', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'text-style-aliases' });
    const slide = presentation.slides.getItem(0);
    const shape = await slide.shapes.addTextBox('Alias Text', { name: 'AliasText' });

    await presentation.clearMutations();
    await shape.text.setTextStyle({
      fontFamily: 'Arial',
      color: '#0000ff',
      bold: true,
      underline: true,
      baseline: 0.25,
      highlight: '#ffcc00',
      lineHeight: 1.35,
      characterSpacing: 1.8,
    });
    await shape.text.setRunTextStyle(0, {
      color: '#00aaee',
      italic: true,
      strike: true,
      baseline: -0.1,
      highlight: '#ccffee',
    });

    expect(await shape.text.style()).toMatchObject({
      color: '#00AAEE',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      baseline: -0.1,
      highlight: '#CCFFEE',
      lineHeight: 1.35,
      characterSpacing: 1.8,
    });
    expect((await shape.text.runs())[0]?.style).toMatchObject({
      fontFamily: 'Arial',
      color: '#00AAEE',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      baseline: -0.1,
      highlight: '#CCFFEE',
      characterSpacing: 1.8,
    });
    expect((await shape.text.paragraphs())[0]?.style).toMatchObject({ lineHeight: 1.35 });
    expect((await presentation.getMutations()).map((mutation) => mutation.type)).toEqual([
      'shape.text.style.set',
      'shape.text.run.style.set',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedShape = reloaded.slides.getItem(0).shapes.getItem('AliasText');
    expect(await reloadedShape.text.style()).toMatchObject({
      color: '#00AAEE',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      baseline: -0.1,
      highlight: '#CCFFEE',
      lineHeight: 1.35,
      characterSpacing: 1.8,
    });
    expect((await reloadedShape.text.runs())[0]?.style).toMatchObject({
      fontFamily: 'Arial',
      color: '#00AAEE',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      baseline: -0.1,
      highlight: '#CCFFEE',
      characterSpacing: 1.8,
    });
  });

  it('rejects invalid style values before mutating the package', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'style-guards' });
    const slide = presentation.slides.getItem(0);
    const rect = await slide.shapes.addShape('rect', { name: 'GuardedRect' });
    const text = await slide.shapes.addTextBox('Guarded Text', { name: 'GuardedText' });
    const picture = await slide.shapes.addImage(
      { bytes: pngBytes(), mimeType: 'image/png' },
      { name: 'GuardedPicture' },
    );
    const table = await slide.shapes.addTable(1, 1, {
      name: 'GuardedTable',
      values: [['Guarded Cell']],
    });
    const cell = table.table.getCell(0, 0);

    await presentation.clearMutations();

    await expect(rect.setFillColor('not-a-color')).rejects.toThrow('invalid color');
    await expect(rect.setStyle({ opacity: 2 })).rejects.toThrow('opacity must be between 0 and 1');
    await expect(rect.setStyle({ lineOpacity: -0.1 })).rejects.toThrow(
      'opacity must be between 0 and 1',
    );
    await expect(rect.setStyle({ shadow: { opacity: 1.5 } })).rejects.toThrow(
      'opacity must be between 0 and 1',
    );
    await expect(rect.setStyle({ shadow: { blur: -1 } })).rejects.toThrow(
      'shadow blur must be a non-negative number',
    );
    await expect(text.text.setStyle({ color: '#12' })).rejects.toThrow('invalid color');
    await expect(text.text.setStyle({ highlight: '#12' })).rejects.toThrow('invalid color');
    await expect(text.text.setStyle({ baseline: Number.NaN })).rejects.toThrow(
      'text baseline must be a finite number',
    );
    await expect(text.text.setAlignment(undefined as never)).rejects.toThrow(
      'alignment is required',
    );
    await expect(picture.image.setOpacity(1.1)).rejects.toThrow('opacity must be between 0 and 1');
    await expect(picture.image.setCrop({ left: -0.1 })).rejects.toThrow(
      'image crop left must be between 0 and 1',
    );
    await expect(picture.image.setCrop({})).rejects.toThrow(
      'image crop must include at least one edge',
    );
    await expect(picture.image.setFillRect({ bottom: 1.1 })).rejects.toThrow(
      'image fillRect bottom must be between 0 and 1',
    );
    await expect(picture.image.setStyle({ fillRect: {} })).rejects.toThrow(
      'image fillRect must include at least one edge',
    );
    await expect(cell.setStyle({ body: { anchor: 'center' as never } })).rejects.toThrow(
      'invalid text body anchor',
    );
    await expect(cell.setStyle({ body: { insetLeft: -1 } })).rejects.toThrow(
      'cell body insetLeft must be a non-negative number',
    );
    await expect(cell.setStyle({ body: { autoFit: 'shape' } })).rejects.toThrow(
      'table cell body autoFit is not writable',
    );
    expect(await presentation.getMutations()).toEqual([]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getItem(0).shapes.getItem('GuardedRect').style()).toEqual({});
    expect(await reloaded.slides.getItem(0).shapes.getItem('GuardedText').text.style()).toEqual({});
    expect(await reloaded.slides.getItem(0).shapes.getItem('GuardedPicture').image.style()).toEqual(
      {},
    );
  });

  it('sets table cell text and basic cell styles', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'table-cell-styles' });
    const slide = presentation.slides.getItem(0);
    const tableShape = await slide.shapes.addTable(2, 2, {
      name: 'StyledTable',
      values: [
        ['A1', 'B1'],
        ['A2', 'B2'],
      ],
    });

    await tableShape.table.setCellText(0, 1, 'Changed B1');
    const cell = tableShape.table.getCell(0, 1);
    await cell.setStyle({
      fillColor: '#ddeeff',
      textStyle: {
        fontSize: 16,
        fontFamily: 'Arial',
        color: '#112233',
        bold: true,
        underline: true,
        strike: true,
        baseline: 0.2,
        highlight: '#ffdd00',
        alignment: 'center',
        lineHeight: 1.2,
        characterSpacing: 2,
      },
      body: {
        anchor: 'middle',
        wrap: 'square',
        insetLeft: 1000,
        insetTop: 2000,
        insetRight: 3000,
        insetBottom: 4000,
      },
    });

    expect(await tableShape.table.values()).toEqual([
      ['A1', 'Changed B1'],
      ['A2', 'B2'],
    ]);
    expect(await cell.style()).toEqual({
      fillColor: '#DDEEFF',
      body: {
        anchor: 'middle',
        wrap: 'square',
        insetLeft: 1000,
        insetTop: 2000,
        insetRight: 3000,
        insetBottom: 4000,
      },
      textStyle: {
        fontSize: 16,
        fontFamily: 'Arial',
        color: '#112233',
        bold: true,
        underline: true,
        strike: true,
        baseline: 0.2,
        highlight: '#FFDD00',
        alignment: 'center',
        lineHeight: 1.2,
        characterSpacing: 2,
      },
    });

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'table.cell.text.set',
      'table.cell.style.set',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedTable = reloaded.slides.getItem(0).shapes.getItem('StyledTable');
    expect(await reloadedTable.table.values()).toEqual([
      ['A1', 'Changed B1'],
      ['A2', 'B2'],
    ]);
    expect(await reloadedTable.table.getCell(0, 1).style()).toEqual({
      fillColor: '#DDEEFF',
      body: {
        anchor: 'middle',
        wrap: 'square',
        insetLeft: 1000,
        insetTop: 2000,
        insetRight: 3000,
        insetBottom: 4000,
      },
      textStyle: {
        fontSize: 16,
        fontFamily: 'Arial',
        color: '#112233',
        bold: true,
        underline: true,
        strike: true,
        baseline: 0.2,
        highlight: '#FFDD00',
        alignment: 'center',
        lineHeight: 1.2,
        characterSpacing: 2,
      },
    });
  });

  it('creates a table cell text run when editing an empty cell body', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'table-empty-cell-run' });
    const slide = presentation.slides.getItem(0);
    await slide.shapes.addTable(1, 1, {
      name: 'SparseTable',
      values: [['A1']],
    });

    const zip = await JSZip.loadAsync(await presentation.toZip());
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    zip.file('ppt/slides/slide1.xml', slideXml.replace('<a:t>A1</a:t>', ''), {
      createFolders: false,
    });

    const sparse = await Presentation.fromZip(await zip.generateAsync({ type: 'uint8array' }));
    const sparseTable = sparse.slides.getItem(0).shapes.getItem('SparseTable');
    expect(await sparseTable.table.values()).toEqual([['']]);

    await sparseTable.table.setCellText(0, 0, 'Filled');
    expect(await sparseTable.table.values()).toEqual([['Filled']]);

    const reloaded = await Presentation.fromZip(await sparse.toZip());
    expect(await reloaded.slides.getItem(0).shapes.getItem('SparseTable').table.values()).toEqual([
      ['Filled'],
    ]);
  });

  it('writes table cell paragraphs and run styles', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'table-cell-text-runs' });
    const slide = presentation.slides.getItem(0);
    const tableShape = await slide.shapes.addTable(1, 1, {
      name: 'TextRunTable',
      values: [['Initial']],
    });
    const cell = tableShape.table.getCell(0, 0);

    await cell.setParagraphs([
      {
        style: {
          alignment: 'center',
          bullet: '•',
          bulletColor: '#AA8822',
          bulletFontFamily: 'Arial',
          lineHeight: 1.45,
          marginLeft: 254000,
          textIndent: -127000,
          spaceBefore: 1270,
          spaceAfter: 1270,
        },
        runs: [
          {
            text: 'Cell ',
            style: {
              bold: true,
              underline: true,
              baseline: 0.18,
              highlight: '#ddeeff',
              characterSpacing: 1.5,
            },
          },
          { text: 'Run', style: { italic: true } },
        ],
      },
    ]);
    await cell.setRunStyle(1, {
      fontSize: 14,
      color: '#778899',
      underline: false,
      strike: true,
      baseline: -0.12,
      highlight: '#ffeecc',
      characterSpacing: 2.5,
    });

    expect(await tableShape.table.values()).toEqual([['Cell Run']]);
    expect(await cell.paragraphs()).toEqual([
      {
        text: 'Cell Run',
        style: {
          alignment: 'center',
          bullet: '•',
          bulletColor: '#AA8822',
          bulletFontFamily: 'Arial',
          lineHeight: 1.45,
          marginLeft: 254000,
          textIndent: -127000,
          spaceBefore: 1270,
          spaceAfter: 1270,
        },
        runs: [
          {
            text: 'Cell ',
            style: {
              bold: true,
              underline: true,
              baseline: 0.18,
              highlight: '#DDEEFF',
              characterSpacing: 1.5,
              alignment: 'center',
            },
          },
          {
            text: 'Run',
            style: {
              fontSize: 14,
              color: '#778899',
              italic: true,
              underline: false,
              strike: true,
              baseline: -0.12,
              highlight: '#FFEECC',
              characterSpacing: 2.5,
              alignment: 'center',
            },
          },
        ],
      },
    ]);

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'table.cell.text.paragraphs.set',
      'table.cell.text.run.style.set',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(
      await reloaded.slides
        .getItem(0)
        .shapes.getItem('TextRunTable')
        .table.getCell(0, 0)
        .paragraphs(),
    ).toEqual(await cell.paragraphs());
  });

  it('adds and deletes table rows and columns', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'table-structure' });
    const slide = presentation.slides.getItem(0);
    const tableShape = await slide.shapes.addTable(2, 2, {
      name: 'StructureTable',
      values: [
        ['A1', 'B1'],
        ['A2', 'B2'],
      ],
    });

    await tableShape.table.addRow(1, ['A1.5', 'B1.5']);
    await tableShape.table.addColumn(1, ['AB1', 'AB1.5', 'AB2']);
    expect(await tableShape.table.rowCount()).toBe(3);
    expect(await tableShape.table.columnCount()).toBe(3);
    expect(await tableShape.table.values()).toEqual([
      ['A1', 'AB1', 'B1'],
      ['A1.5', 'AB1.5', 'B1.5'],
      ['A2', 'AB2', 'B2'],
    ]);

    await tableShape.table.deleteRow(0);
    await tableShape.table.deleteColumn(2);
    expect(await tableShape.table.values()).toEqual([
      ['A1.5', 'AB1.5'],
      ['A2', 'AB2'],
    ]);

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.add',
      'table.row.add',
      'table.column.add',
      'table.row.delete',
      'table.column.delete',
    ]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedTable = reloaded.slides.getItem(0).shapes.getItem('StructureTable');
    expect(await reloadedTable.table.rowCount()).toBe(2);
    expect(await reloadedTable.table.columnCount()).toBe(2);
    expect(await reloadedTable.table.values()).toEqual([
      ['A1.5', 'AB1.5'],
      ['A2', 'AB2'],
    ]);
  });

  it('rejects invalid table dimensions and structural edits without corrupting the package', async () => {
    const presentation = await Presentation.createBlank({ presentationId: 'table-guards' });
    const slide = presentation.slides.getItem(0);

    await expect(slide.shapes.addTable(0, 2)).rejects.toThrow('rowCount must be positive');
    await expect(slide.shapes.addTable(2, 0)).rejects.toThrow('columnCount must be positive');
    await expect(slide.shapes.addTable(2, 2, { values: [['A', 'B']] })).rejects.toThrow(
      'must match 2x2',
    );

    const tableShape = await slide.shapes.addTable(1, 1, {
      name: 'GuardedTable',
      values: [['A1']],
    });
    await presentation.clearMutations();

    await expect(tableShape.table.setValues([['A1', 'B1']])).rejects.toThrow('must match 1x1');
    await expect(tableShape.table.getCell(1, 0).text()).rejects.toThrow('cell out of range');
    await expect(tableShape.table.setCellText(0, 1, 'B1')).rejects.toThrow('cell out of range');
    await expect(tableShape.table.setCellText('__proto__' as never, 0, 'polluted')).rejects.toThrow(
      'cell out of range',
    );
    await expect(
      tableShape.table.getCell(Number.NaN, 0).setStyle({ fillColor: '#ffffff' }),
    ).rejects.toThrow('cell out of range');
    expect(Object.prototype).not.toHaveProperty('polluted');
    await expect(tableShape.table.addRow('end', ['A2', 'B2'])).rejects.toThrow(
      'row values must have length 1',
    );
    await expect(tableShape.table.addColumn('end', ['B1', 'B2'])).rejects.toThrow(
      'column values must have length 1',
    );
    await expect(tableShape.table.deleteRow(0)).rejects.toThrow(
      'cannot delete the final table row',
    );
    await expect(tableShape.table.deleteColumn(0)).rejects.toThrow(
      'cannot delete the final table column',
    );
    expect(await presentation.getMutations()).toEqual([]);

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getItem(0).shapes.getItem('GuardedTable').table.values()).toEqual([
      ['A1'],
    ]);
  });

  it('reads slides, shapes, text, image metadata, and table values from a pptx package', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx(), {
      presentationId: 'unit-fixture',
    });

    const model = await presentation.toModel();
    expect(model.id).toBe('unit-fixture');
    expect(model.slides).toHaveLength(1);
    expect(model.themes).toEqual([
      {
        relationshipId: 'rId1',
        path: 'ppt/theme/theme1.xml',
        name: 'Coco Theme',
        colorScheme: {
          dk1: '#111111',
          lt1: '#FFFFFF',
          accent1: '#3366FF',
          accent2: '#22AA88',
        },
        fontScheme: {
          name: 'Coco Fonts',
          majorLatin: 'Aptos Display',
          minorLatin: 'Aptos',
        },
      },
    ]);
    expect(model.slides[0]!.layoutId).toBe('rId6');
    expect(model.slides[0]!.layout).toEqual({
      relationshipId: 'rId6',
      path: 'ppt/slideLayouts/slideLayout1.xml',
      name: 'Title and Content',
      master: {
        relationshipId: 'rId1',
        path: 'ppt/slideMasters/slideMaster1.xml',
        name: 'Office Theme',
        theme: model.themes![0],
      },
    });
    expect(model.slides[0]!.shapes.map((shape) => shape.kind)).toEqual([
      'placeholder',
      'textBox',
      'picture',
      'table',
      'chart',
      'group',
      'media',
      'oleObject',
      'smartArt',
      'table',
      'textBox',
    ]);
    expect(model.slides[0]!.shapes[4]).toMatchObject({
      id: '5',
      name: 'SalesChart',
      kind: 'chart',
      chart: {
        relationshipId: 'rId5',
        targetPath: 'ppt/charts/chart1.xml',
        type: 'bar',
        grouping: 'stacked',
        title: 'Quarterly Sales',
        seriesNames: ['North', 'South'],
      },
    });
    expect(model.slides[0]!.shapes[4]!.chart?.data?.series[0]).toMatchObject({
      chartGroup: 0,
      borderColor: '#FF0000',
      borderWidth: 12700,
      lineColor: '#FF0000',
      lineWidth: 12700,
    });
    expect(model.slides[0]!.shapes[4]!.chart?.axes).toMatchObject({
      category: {
        lineColor: '#00FF00',
        lineWidth: 9525,
      },
      value: {
        gridlineColor: '#FFFFFF',
        gridlineWidth: 9525,
      },
    });
    expect(model.slides[0]!.shapes[5]).toMatchObject({
      id: '6',
      name: 'Group 1',
      kind: 'group',
      children: [
        {
          id: '7',
          name: 'GroupedText',
          kind: 'textBox',
          text: { text: 'Grouped Coco', style: { color: '#FFFFFF' } },
        },
        {
          id: '8',
          name: 'GroupedRect',
          kind: 'geometricShape',
          style: {
            fillColor: '#05E1FF',
            opacity: 0.15,
          },
        },
        {
          id: '15',
          name: 'GroupedTable',
          kind: 'table',
          table: {
            values: [['G1', 'G2']],
            rowCount: 1,
            columnCount: 2,
          },
        },
      ],
    });
    expect(model.slides[0]!.shapes[6]).toMatchObject({
      id: '9',
      name: 'Video 1',
      kind: 'media',
      media: {
        relationshipId: 'rId7',
        targetPath: 'ppt/media/video1.mp4',
        mediaType: 'video',
        contentType: 'video/mp4',
        filename: 'video1.mp4',
      },
    });
    expect(model.slides[0]!.shapes[7]).toMatchObject({
      id: '10',
      name: 'Workbook OLE',
      kind: 'oleObject',
      oleObject: {
        relationshipId: 'rId8',
        targetPath: 'ppt/embeddings/oleObject1.bin',
        progId: 'Excel.Sheet.12',
        name: 'Workbook',
      },
    });
    expect(model.slides[0]!.shapes[8]).toMatchObject({
      id: '11',
      name: 'Process SmartArt',
      kind: 'smartArt',
      smartArt: {
        relationshipIds: {
          dm: 'rId9',
          lo: 'rId10',
          qs: 'rId11',
          cs: 'rId12',
        },
        targetPaths: {
          dm: 'ppt/diagrams/data1.xml',
          lo: 'ppt/diagrams/layout1.xml',
          qs: 'ppt/diagrams/quickStyle1.xml',
          cs: 'ppt/diagrams/colors1.xml',
        },
      },
    });
    expect(model.slides[0]!.shapes[9]).toMatchObject({
      id: '12',
      name: 'MergedTable',
      kind: 'table',
      table: {
        values: [
          ['Merged A', ''],
          ['C1', 'D1'],
        ],
        mergedCells: [
          { row: 0, column: 0, columnSpan: 2 },
          { row: 0, column: 1, hMerge: true },
        ],
      },
    });
    expect(model.slides[0]!.notes).toEqual({ text: 'Speaker note for Coco' });
    expect(model.slides[0]!.comments).toEqual([
      {
        id: '1',
        authorId: '0',
        author: 'Coco Reviewer',
        initials: 'CR',
        date: '2026-08-25T09:00:00Z',
        text: 'Review this grouped link',
        x: 100,
        y: 200,
      },
    ]);
    expect(model.slides[0]!.transition).toEqual({
      type: 'fade',
      speed: 'fast',
      duration: 750,
      advanceOnClick: true,
      advanceAfterTime: 3000,
      attributes: {
        spd: 'fast',
        'p14:dur': '750',
        advClick: '1',
        advTm: '3000',
      },
    });
    expect(model.slides[0]!.timing).toEqual({
      hasTiming: true,
      timeNodeCount: 2,
      mainSequenceCount: 1,
      interactiveSequenceCount: 0,
      rootDuration: 'indefinite',
      rootNodeType: 'tmRoot',
    });

    const slide = presentation.slides.getItem(0);
    expect(await presentation.themes.getCount()).toBe(1);
    expect(await presentation.themes.items()).toEqual(model.themes);
    expect(await slide.title()).toBe('Hello Coco');
    expect(await slide.layout()).toEqual(model.slides[0]!.layout);
    expect(model.slides[0]!.background).toMatchObject({
      imagePath: 'ppt/media/image1.png',
    });
    expect(await slide.transition()).toEqual(model.slides[0]!.transition);
    expect(await slide.timing()).toEqual(model.slides[0]!.timing);
    expect(await slide.notes.hasText()).toBe(true);
    expect(await slide.notes.getText()).toBe('Speaker note for Coco');
    expect(await slide.comments.getCount()).toBe(1);
    expect(await slide.comments.items()).toEqual(model.slides[0]!.comments);
    expect(await slide.shapes.getCount()).toBe(11);
    expect(await slide.shapes.getItem('TitleBox').toModel()).toMatchObject({
      placeholder: {
        type: 'title',
        index: '1',
        size: 'full',
        orientation: 'horz',
      },
      text: {
        body: {
          autoFit: 'shape',
          wrap: 'none',
        },
      },
    });
    expect((await slide.shapes.getItem('TitleBox').toModel()).text?.style?.bold).toBeUndefined();
    expect(await slide.shapes.getItem('SingleBoldBox').toModel()).toMatchObject({
      raw: {
        useBackgroundFill: true,
      },
      text: {
        style: {
          bold: true,
        },
        body: {
          autoFit: 'shape',
        },
        paragraphs: [
          {
            style: {
              bullet: '✓',
              bulletColor: '#22AA88',
              bulletFontFamily: 'Wingdings',
            },
          },
        ],
      },
    });
    expect(
      model.slides[0]!.shapes.find((shape) => shape.name === 'Group 1')?.children?.find(
        (shape) => shape.name === 'GroupedText',
      )?.text?.body,
    ).toEqual({});
    expect(await slide.shapes.getItem('VerticalText').toModel()).toMatchObject({
      text: {
        text: 'WORK SUMMARY',
        style: {
          color: '#3366FF',
          textGradient: {
            colors: ['#3366FF', '#22AA88'],
            positions: [0, 100],
          },
          textShadow: {
            color: '#FFFFFF',
            opacity: 0.5,
            blur: 60000,
            distance: 120000,
            direction: 90,
          },
        },
        body: {
          vertical: 'eastAsian',
          autoFit: 'shape',
          wrap: 'square',
        },
      },
    });
    expect(await slide.shapes.getItem('TitleBox').text.getText()).toBe('Hello Coco');
    expect(await slide.shapes.getItem('TitleBox').text.paragraphs()).toEqual([
      {
        text: 'Hello Coco',
        style: { alignment: 'center' },
        runs: [
          {
            text: 'Hello ',
            style: {
              fontSize: 20,
              color: '#112233',
              bold: true,
              alignment: 'center',
            },
          },
          {
            text: 'Coco',
            style: {
              italic: true,
              alignment: 'center',
            },
            hyperlink: {
              relationshipId: 'rId2',
              target: 'https://example.com/coco',
              targetMode: 'External',
              tooltip: 'Coco link',
            },
          },
        ],
      },
    ]);
    expect(await slide.shapes.getItem('TitleBox').text.runs()).toEqual([
      {
        text: 'Hello ',
        style: {
          fontSize: 20,
          color: '#112233',
          bold: true,
          alignment: 'center',
        },
      },
      {
        text: 'Coco',
        style: {
          italic: true,
          alignment: 'center',
        },
        hyperlink: {
          relationshipId: 'rId2',
          target: 'https://example.com/coco',
          targetMode: 'External',
          tooltip: 'Coco link',
        },
      },
    ]);
    expect(await slide.shapes.getItem('Picture 1').image.getInfo()).toMatchObject({
      targetPath: 'ppt/media/image1.png',
      filename: 'image1.png',
      byteLength: 8,
    });
    expect(await slide.shapes.getItem('DataTable').table.values()).toEqual([
      ['A1', 'B1\n\nB1 details'],
      ['A2', 'B2'],
    ]);
    const dataTableModel = (await presentation.toModel()).slides[0]!.shapes.find(
      (shape) => shape.name === 'DataTable',
    );
    expect(dataTableModel?.table?.cellStyles?.[0]?.[0]?.borders).toMatchObject({
      left: { hidden: true },
      top: { color: '#FF0000', opacity: 0.5, width: 25400, dash: 'dash' },
    });
    expect(await slide.shapes.getItem('DataTable').table.getCell(0, 0).runs()).toEqual([
      { text: 'A1' },
    ]);
    expect(await slide.shapes.getItem('MergedTable').table.hasMergedCells()).toBe(true);
    expect(await slide.shapes.getItem('MergedTable').table.mergedCells()).toEqual([
      { row: 0, column: 0, columnSpan: 2 },
      { row: 0, column: 1, hMerge: true },
    ]);
    expect(await slide.shapes.getItem('SalesChart').chart.isChart()).toBe(true);
    expect(await slide.shapes.getItem('SalesChart').chart.getInfo()).toMatchObject({
      relationshipId: 'rId5',
      targetPath: 'ppt/charts/chart1.xml',
      type: 'bar',
      title: 'Quarterly Sales',
      seriesNames: ['North', 'South'],
    });
    expect(await slide.shapes.getItem('SalesChart').chart.getData()).toEqual({
      title: 'Quarterly Sales',
      series: [
        { name: 'North', categories: ['Q1', 'Q2'], values: [10, 12] },
        { name: 'South', categories: ['Q1', 'Q2'], values: [8, 14] },
      ],
    });
    const group = slide.shapes.getItem('Group 1');
    expect(await group.group.isGroup()).toBe(true);
    expect(await group.text.hasText()).toBe(false);
    expect(await group.group.getCount()).toBe(3);
    const groupedText = group.group.getItem('GroupedText');
    expect(await groupedText.text.getText()).toBe('Grouped Coco');
    expect((await groupedText.toModel()).text?.style?.color).toBe('#FFFFFF');
    expect(await group.group.getItemAt(1).name()).toBe('GroupedRect');
    expect(await group.group.getItemAt(2).name()).toBe('GroupedTable');
    expect(await group.group.getItem('GroupedRect').hyperlink()).toEqual({
      relationshipId: 'rId2',
      target: 'https://example.com/coco',
      targetMode: 'External',
    });
    await expect(slide.shapes.getItem('TitleBox').group.getCount()).rejects.toThrow(
      'is not a group',
    );
    expect(
      (await presentation.findText('Grouped Coco', { caseSensitive: true })).map((match) => ({
        shapeId: match.shapeId,
        shapeName: match.shapeName,
        source: match.source,
      })),
    ).toEqual([{ shapeId: '7', shapeName: 'GroupedText', source: 'shapeText' }]);
    await expect(
      presentation.replaceText('Grouped Coco', 'Grouped SDK', { caseSensitive: true }),
    ).resolves.toEqual({ count: 1 });
    expect(await group.group.getItem('GroupedText').text.getText()).toBe('Grouped SDK');
  });

  it('round-trips grouped child geometry edits', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx());
    const groupedText = presentation.slides
      .getItem(0)
      .shapes.getItem('Group 1')
      .group.getItem('GroupedText');

    await groupedText.setGeometry({
      left: 1300,
      top: 1250,
      width: 900,
      height: 360,
      rotation: 180,
    });

    const immediateModel = await presentation.toModel();
    const immediateGroupedText = immediateModel.slides[0]!.shapes.find(
      (shape) => shape.name === 'Group 1',
    )?.children?.find((shape) => shape.name === 'GroupedText');
    expect(immediateGroupedText?.geometry).toMatchObject({
      left: 1300,
      top: 1250,
      width: 900,
      height: 360,
      rotation: 180,
    });

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedGroupedText = await reloaded.slides
      .getItem(0)
      .shapes.getItem('Group 1')
      .group.getItem('GroupedText')
      .toModel();
    expect(reloadedGroupedText.geometry).toMatchObject({
      left: 1300,
      top: 1250,
      width: 900,
      height: 360,
      rotation: 180,
    });
    expect(reloadedGroupedText.text?.text).toBe('Grouped Coco');
  });

  it('round-trips grouped table cell edits', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx());
    const groupedTable = presentation.slides
      .getItem(0)
      .shapes.getItem('Group 1')
      .group.getItem('GroupedTable');

    expect((await groupedTable.toModel()).table?.values).toEqual([['G1', 'G2']]);

    await groupedTable.table.setCellText(0, 1, 'Grouped Final');

    const immediateModel = await presentation.toModel();
    const immediateGroupedTable = immediateModel.slides[0]!.shapes.find(
      (shape) => shape.name === 'Group 1',
    )?.children?.find((shape) => shape.name === 'GroupedTable');
    expect(immediateGroupedTable?.table?.values[0]?.[1]).toBe('Grouped Final');

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedGroupedTable = await reloaded.slides
      .getItem(0)
      .shapes.getItem('Group 1')
      .group.getItem('GroupedTable')
      .toModel();
    expect(reloadedGroupedTable.table?.values[0]?.[1]).toBe('Grouped Final');
  });

  it('parses alpha values on gradient fill stops', async () => {
    const zip = await JSZip.loadAsync(await makeFixturePptx());
    const slidePath = 'ppt/slides/slide1.xml';
    const xml = await zip.file(slidePath)!.async('string');
    zip.file(
      slidePath,
      xml.replace(
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>',
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00AAFF"><a:alpha val="40000"/></a:srgbClr></a:gs><a:gs pos="39000"><a:srgbClr val="00AAFF"><a:alpha val="0"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill><a:ln><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00AAFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="00AAFF"><a:alpha val="0"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill><a:prstDash val="dash"/></a:ln></p:spPr>',
      ),
      { createFolders: false },
    );

    const presentation = await Presentation.fromZip(
      await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      }),
    );
    const model = await presentation.slides.getItem(0).shapes.getItem('TitleBox').toModel();

    expect(model.style?.fillGradient).toEqual({
      colors: ['#00AAFF', '#00AAFF'],
      angle: 90,
      scaled: true,
      positions: [0, 39],
      opacities: [0.4, 0],
    });
    expect(model.style?.lineGradient).toEqual({
      colors: ['#00AAFF', '#00AAFF'],
      angle: 90,
      scaled: true,
      positions: [0, 100],
      opacities: [1, 0],
    });
    expect(model.style?.lineDash).toBe('dash');
  });

  it('preserves shape 3D material metadata for rendering projections', async () => {
    const zip = await JSZip.loadAsync(await makeFixturePptx());
    const slidePath = 'ppt/slides/slide1.xml';
    const xml = await zip.file(slidePath)!.async('string');
    zip.file(
      slidePath,
      xml.replace(
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>',
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="trapezoid"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="56B6EE"/></a:solidFill><a:scene3d><a:camera prst="perspectiveRelaxedModerately"/></a:scene3d><a:sp3d extrusionH="127000" prstMaterial="translucentPowder"/></p:spPr>',
      ),
      { createFolders: false },
    );

    const presentation = await Presentation.fromZip(
      await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      }),
    );
    const model = await presentation.slides.getItem(0).shapes.getItem('TitleBox').toModel();

    expect(model.raw?.threeD).toEqual({
      material: 'translucentPowder',
      extrusionHeight: 127000,
    });
  });

  it('parses custom line dash patterns', async () => {
    const zip = await JSZip.loadAsync(await makeFixturePptx());
    const slidePath = 'ppt/slides/slide1.xml';
    const xml = await zip.file(slidePath)!.async('string');
    zip.file(
      slidePath,
      xml.replace(
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>',
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="12689"><a:solidFill><a:srgbClr val="BED0A6"><a:alpha val="42000"/></a:srgbClr></a:solidFill><a:custDash><a:ds d="1500038" sp="1500038"/></a:custDash></a:ln></p:spPr>',
      ),
      { createFolders: false },
    );

    const presentation = await Presentation.fromZip(
      await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      }),
    );
    const model = await presentation.slides.getItem(0).shapes.getItem('TitleBox').toModel();

    expect(model.style?.lineColor).toBe('#BED0A6');
    expect(model.style?.lineOpacity).toBe(0.42);
    expect(model.style?.lineDash).toBe('cust:1500038:1500038');
  });

  it('applies scheme color luminance modifiers in HSL space', async () => {
    const zip = await JSZip.loadAsync(await makeFixturePptx());
    const slidePath = 'ppt/slides/slide1.xml';
    const xml = await zip.file(slidePath)!.async('string');
    zip.file(
      slidePath,
      xml.replace(
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>',
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:ln w="101600"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill></a:ln></p:spPr>',
      ),
      { createFolders: false },
    );

    const presentation = await Presentation.fromZip(
      await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      }),
    );
    const model = await presentation.slides.getItem(0).shapes.getItem('TitleBox').toModel();

    expect(model.style?.lineColor).toBe('#85A3FF');
  });

  it('inherits chart-level value label visibility into series data', async () => {
    const zip = await JSZip.loadAsync(await makeFixturePptx());
    const chartPath = 'ppt/charts/chart1.xml';
    const xml = await zip.file(chartPath)!.async('string');
    const nextXml = xml.replace(
      '<c:grouping val="stacked"/>',
      '<c:grouping val="stacked"/><c:dLbls><c:showVal val="1"/><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls>',
    );
    expect(nextXml).toContain(
      '<c:dLbls><c:showVal val="1"/><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls>',
    );
    zip.file(chartPath, nextXml, { createFolders: false });

    const presentation = await Presentation.fromZip(
      await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      }),
    );
    const chart = await presentation.slides.getItem(0).shapes.getItem('SalesChart').chart.getData();

    expect(chart?.series.map((series) => series.showValueLabels)).toEqual([true, true]);
    expect(chart?.series.map((series) => series.showCategoryNameLabels)).toEqual([true, true]);
    expect(chart?.series.map((series) => series.showPercentLabels)).toEqual([true, true]);
  });

  it('parses gradient fill direction metadata', async () => {
    const zip = await JSZip.loadAsync(await makeFixturePptx());
    const slidePath = 'ppt/slides/slide1.xml';
    const xml = await zip.file(slidePath)!.async('string');
    zip.file(
      slidePath,
      xml.replace(
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>',
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:gradFill flip="x" rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="00AAFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/><a:tileRect l="10000" t="20000" r="30000" b="40000"/></a:gradFill></p:spPr>',
      ),
      { createFolders: false },
    );

    const presentation = await Presentation.fromZip(
      await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      }),
    );
    const model = await presentation.slides.getItem(0).shapes.getItem('TitleBox').toModel();

    expect(model.style?.fillGradient).toEqual({
      colors: ['#00AAFF', '#FFFFFF'],
      angle: 270,
      flip: 'x',
      rotWithShape: true,
      scaled: false,
      tileRect: { left: 10, top: 20, right: 30, bottom: 40 },
      positions: [0, 100],
    });
  });

  it('uses the first gradient stop as a fallback color for text runs', async () => {
    const zip = await JSZip.loadAsync(await makeFixturePptx());
    const slidePath = 'ppt/slides/slide1.xml';
    const xml = await zip.file(slidePath)!.async('string');
    zip.file(
      slidePath,
      xml.replace(
        '<a:r><a:rPr i="1"><a:hlinkClick r:id="rId2" tooltip="Coco link"/></a:rPr><a:t>Coco</a:t></a:r>',
        '<a:r><a:rPr i="1"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="3878D0"/></a:gs><a:gs pos="100000"><a:srgbClr val="56B6EE"/></a:gs></a:gsLst><a:lin ang="10800000" scaled="1"/></a:gradFill><a:hlinkClick r:id="rId2" tooltip="Coco link"/></a:rPr><a:t>Coco</a:t></a:r>',
      ),
      { createFolders: false },
    );

    const presentation = await Presentation.fromZip(
      await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      }),
    );

    const title = (await presentation.toModel()).slides[0]!.shapes.find(
      (shape) => shape.name === 'TitleBox',
    )!;

    expect(title.text?.paragraphs).toMatchObject([
      {
        runs: [
          {},
          {
            text: 'Coco',
            style: {
              color: '#3878D0',
              italic: true,
            },
          },
        ],
      },
    ]);
  });

  it('keeps gradient-only arrow connectors as editable line shapes', async () => {
    const zip = await JSZip.loadAsync(await makeFixturePptx());
    const slidePath = 'ppt/slides/slide1.xml';
    const xml = await zip.file(slidePath)!.async('string');
    zip.file(
      slidePath,
      xml.replace(
        '    </p:spTree>',
        `      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="42" name="Gradient Arrow"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1000" y="3000"/><a:ext cx="1800000" cy="0"/></a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="38100" cap="flat">
            <a:gradFill>
              <a:gsLst>
                <a:gs pos="0"><a:srgbClr val="00B3F8"/></a:gs>
                <a:gs pos="100000"><a:srgbClr val="CCF0FE"><a:alpha val="50000"/></a:srgbClr></a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="1"/>
            </a:gradFill>
            <a:tailEnd type="arrow" w="lg" len="lg"/>
          </a:ln>
        </p:spPr>
      </p:cxnSp>
    </p:spTree>`,
      ),
      { createFolders: false },
    );

    const presentation = await Presentation.fromZip(
      await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      }),
    );
    const model = await presentation.slides.getItem(0).shapes.getItem('Gradient Arrow').toModel();

    expect(model.kind).toBe('line');
    expect(model.raw?.presetGeometry).toBe('straightConnector1');
    expect(model.style?.lineCap).toBe('flat');
    expect(model.style?.lineTailEnd).toBe('arrow');
    expect(model.style?.lineGradient).toMatchObject({
      colors: ['#00B3F8', '#CCF0FE'],
      positions: [0, 100],
      opacities: [1, 0.5],
    });
  });

  it('updates text, geometry, table cells, and image bytes, then round-trips through toZip', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx());
    const slide = presentation.slides.getItem(0);
    const title = slide.shapes.getItem('TitleBox');

    await title.text.setText('Updated Title');
    await title.setGeometry({ left: 333.4, width: 444.6 });
    await slide.shapes.getItem('DataTable').table.getCell(1, 1).setText('Changed');
    await slide.shapes.getItem('Picture 1').image.replace(new Uint8Array([9, 8, 7, 6]));

    const mutations = await presentation.getMutations();
    expect(mutations.map((mutation) => mutation.type)).toEqual([
      'shape.text.set',
      'shape.geometry.set',
      'table.cell.text.set',
      'shape.image.replace',
    ]);

    const roundTrip = await Presentation.fromZip(await presentation.toZip());
    const nextSlide = roundTrip.slides.getItem(0);
    expect(await nextSlide.shapes.getItem('TitleBox').text.getText()).toBe('Updated Title');
    expect(await nextSlide.shapes.getItem('TitleBox').geometry()).toMatchObject({
      left: 333,
      width: 445,
    });
    expect(await nextSlide.shapes.getItem('DataTable').table.values()).toEqual([
      ['A1', 'B1\n\nB1 details'],
      ['A2', 'Changed'],
    ]);
    expect(await nextSlide.shapes.getItem('Picture 1').image.getInfo()).toMatchObject({
      byteLength: 4,
    });
  });

  it('replaces text and clears mutation history', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx());
    const shape = presentation.slides.getItem(0).shapes.getItem('TitleBox');

    await expect(shape.text.replace('coco', 'SDK')).resolves.toEqual({ count: 1 });
    expect(await shape.text.getText()).toBe('Hello SDK');
    expect(await presentation.getMutations()).toHaveLength(2);

    await presentation.clearMutations();
    expect(await presentation.getMutations()).toEqual([]);
  });

  it('writes and clears shape and text-run hyperlinks with slide relationships', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx());
    const slide = presentation.slides.getItem(0);
    const title = slide.shapes.getItem('TitleBox');
    const groupedRect = slide.shapes.getItem('Group 1').group.getItem('GroupedRect');

    await title.setHyperlink({
      target: 'https://example.com/title',
      tooltip: 'Title link',
    });
    await title.text.setRunHyperlink(0, {
      target: 'https://example.com/hello',
      tooltip: 'Hello link',
    });
    await groupedRect.clearHyperlink();

    const firstRoundtrip = await Presentation.fromZip(await presentation.toZip());
    const reloadedTitle = firstRoundtrip.slides.getItem(0).shapes.getItem('TitleBox');
    expect(await reloadedTitle.hyperlink()).toMatchObject({
      target: 'https://example.com/title',
      targetMode: 'External',
      tooltip: 'Title link',
    });
    expect((await reloadedTitle.text.runs())[0]!.hyperlink).toMatchObject({
      target: 'https://example.com/hello',
      targetMode: 'External',
      tooltip: 'Hello link',
    });
    expect((await reloadedTitle.text.runs())[1]!.hyperlink).toMatchObject({
      target: 'https://example.com/coco',
      targetMode: 'External',
      tooltip: 'Coco link',
    });
    expect(
      await firstRoundtrip.slides
        .getItem(0)
        .shapes.getItem('Group 1')
        .group.getItem('GroupedRect')
        .hyperlink(),
    ).toBeNull();

    await reloadedTitle.text.clearRunHyperlink(1);
    await reloadedTitle.clearHyperlink();
    const secondRoundtrip = await Presentation.fromZip(await firstRoundtrip.toZip());
    const clearedTitle = secondRoundtrip.slides.getItem(0).shapes.getItem('TitleBox');
    expect(await clearedTitle.hyperlink()).toBeNull();
    expect((await clearedTitle.text.runs())[0]!.hyperlink).toMatchObject({
      target: 'https://example.com/hello',
    });
    expect((await clearedTitle.text.runs())[1]!.hyperlink).toBeUndefined();

    expect(await firstRoundtrip.getMutations()).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ type: 'shape.text.run.hyperlink.clear', runIndex: 1 }),
        expect.objectContaining({ type: 'shape.hyperlink.clear' }),
      ]),
    );
  });

  it('updates and clears existing notes and comments', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx());
    const slide = presentation.slides.getItem(0);

    await slide.notes.setText('Updated note from SDK');
    const added = await slide.comments.add({
      text: 'Second review',
      author: 'Coco Reviewer',
      initials: 'CR',
      date: '2026-08-25T11:00:00Z',
      x: 300,
      y: 400,
    });
    expect(added.id).toBe('2');
    await slide.comments.delete('1');

    const roundTrip = await Presentation.fromZip(await presentation.toZip());
    const roundTripSlide = roundTrip.slides.getItem(0);
    expect(await roundTripSlide.notes.getText()).toBe('Updated note from SDK');
    expect(await roundTripSlide.comments.items()).toEqual([
      {
        id: '2',
        authorId: '0',
        author: 'Coco Reviewer',
        initials: 'CR',
        date: '2026-08-25T11:00:00Z',
        text: 'Second review',
        x: 300,
        y: 400,
      },
    ]);

    await roundTripSlide.notes.clear();
    await roundTripSlide.comments.clear();
    const cleared = await Presentation.fromZip(await roundTrip.toZip());
    expect(await cleared.slides.getItem(0).notes.getText()).toBe('');
    expect(await cleared.slides.getItem(0).comments.items()).toEqual([]);
  });

  it('reads merged table cells and protects unsafe edits', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx());
    const table = presentation.slides.getItem(0).shapes.getItem('MergedTable').table;

    expect(await table.hasMergedCells()).toBe(true);
    expect(await table.mergedCells()).toEqual([
      { row: 0, column: 0, columnSpan: 2 },
      { row: 0, column: 1, hMerge: true },
    ]);

    await expect(
      table.setValues([
        ['A', 'B'],
        ['C', 'D'],
      ]),
    ).rejects.toThrow('merged-cell table');
    await expect(table.addRow('end', ['E', 'F'])).rejects.toThrow('merged-cell table');
    await expect(table.deleteColumn(1)).rejects.toThrow('merged-cell table');
    await expect(table.setCellText(0, 1, 'Blocked')).rejects.toThrow('merged continuation cell');

    await table.setCellText(0, 0, 'Merged Updated');
    const reloaded = await Presentation.fromZip(await presentation.toZip());
    expect(await reloaded.slides.getItem(0).shapes.getItem('MergedTable').table.values()).toEqual([
      ['Merged Updated', ''],
      ['C1', 'D1'],
    ]);
  });

  it('replaces chart data conservatively and round-trips chart XML', async () => {
    const presentation = await Presentation.fromZip(await makeFixturePptx());
    const chart = presentation.slides.getItem(0).shapes.getItem('SalesChart').chart;

    await chart.setData({
      title: 'Updated Sales',
      series: [
        { name: 'East', categories: ['H1', 'H2'], values: [21, 22] },
        { name: 'West', categories: ['H1', 'H2'], values: [31, 32] },
      ],
    });
    await expect(
      chart.setData({
        series: [{ name: 'Only One', values: [1] }],
      }),
    ).rejects.toThrow('must contain 2 series');

    const reloaded = await Presentation.fromZip(await presentation.toZip());
    const reloadedChart = reloaded.slides.getItem(0).shapes.getItem('SalesChart').chart;
    expect(await reloadedChart.getInfo()).toMatchObject({
      title: 'Updated Sales',
      seriesNames: ['East', 'West'],
    });
    expect(await reloadedChart.getData()).toEqual({
      title: 'Updated Sales',
      series: [
        { name: 'East', categories: ['H1', 'H2'], values: [21, 22] },
        { name: 'West', categories: ['H1', 'H2'], values: [31, 32] },
      ],
    });
    expect((await presentation.getMutations()).map((mutation) => mutation.type)).toContain(
      'shape.chart.data.set',
    );
  });

  it('preserves advanced shapes and rejects unsupported advanced-domain edits', async () => {
    const bytes = await makeFixturePptx();
    const originalAdvancedParts = await selectedZipEntryBytes(bytes, isAdvancedPptPart);
    const presentation = await Presentation.fromZip(bytes);
    const slide = presentation.slides.getItem(0);
    const beforeAdvancedModels = await advancedShapeModels(slide);

    await presentation.clearMutations();
    await slide.shapes.getItem('TitleBox').text.setText('Ordinary edit beside advanced shapes');

    await expect(slide.shapes.getItem('SalesChart').text.setText('Not supported')).rejects.toThrow(
      'has no text body',
    );
    await expect(slide.shapes.getItem('Video 1').image.bytes()).rejects.toThrow(
      'has no image bytes',
    );
    await expect(slide.shapes.getItem('Workbook OLE').table.values()).rejects.toThrow(
      'is not a table',
    );
    await expect(slide.shapes.getItem('Process SmartArt').chart.getData()).resolves.toBeNull();
    await expect(
      slide.shapes.getItem('Process SmartArt').chart.setData({ series: [] }),
    ).rejects.toThrow('has no chart XML');

    expect((await presentation.getMutations()).map((mutation) => mutation.type)).toEqual([
      'shape.text.set',
    ]);

    const output = await presentation.toZip();
    const outputAdvancedParts = await selectedZipEntryBytes(output, isAdvancedPptPart);
    expect([...outputAdvancedParts.keys()].sort()).toEqual(
      [...originalAdvancedParts.keys()].sort(),
    );
    for (const [path, originalBytes] of originalAdvancedParts) {
      expect(bytesEqual(outputAdvancedParts.get(path), originalBytes), path).toBe(true);
    }

    const reloaded = await Presentation.fromZip(output);
    expect(await advancedShapeModels(reloaded.slides.getItem(0))).toEqual(beforeAdvancedModels);
  });

  it('does not mutate the package when unsupported advanced-domain writes fail', async () => {
    const bytes = await makeFixturePptx();
    const originalEntries = await zipEntryBytes(bytes);
    const presentation = await Presentation.fromZip(bytes);
    const slide = presentation.slides.getItem(0);
    const beforeAdvancedModels = await advancedShapeModels(slide);

    await presentation.clearMutations();

    await expect(slide.shapes.getItem('SalesChart').text.setStyle({ bold: true })).rejects.toThrow(
      'has no text body',
    );
    await expect(
      slide.shapes.getItem('SalesChart').text.setRunTextStyle(0, { bold: true }),
    ).rejects.toThrow('has no text body');
    await expect(slide.shapes.getItem('Video 1').image.replace(pngBytes())).rejects.toThrow(
      'has no replaceable image',
    );
    await expect(slide.shapes.getItem('Video 1').image.setOpacity(0.5)).rejects.toThrow(
      'has no image style surface',
    );
    await expect(slide.shapes.getItem('Workbook OLE').table.setValues([['x']])).rejects.toThrow(
      'is not a table',
    );
    await expect(slide.shapes.getItem('Workbook OLE').table.setCellText(0, 0, 'x')).rejects.toThrow(
      'is not a table',
    );
    await expect(
      slide.shapes.getItem('Process SmartArt').chart.setData({ title: 'x', series: [] }),
    ).rejects.toThrow('has no chart XML');
    await expect(slide.shapes.getItem('Group 1').image.setCrop({ left: 0.1 })).rejects.toThrow(
      'has no image style surface',
    );

    expect(await presentation.getMutations()).toEqual([]);

    const output = await presentation.toZip();
    const outputEntries = await zipEntryBytes(output);
    expect([...outputEntries.keys()].sort()).toEqual([...originalEntries.keys()].sort());
    for (const [path, originalBytes] of originalEntries) {
      expect(bytesEqual(outputEntries.get(path), originalBytes), path).toBe(true);
    }

    const reloaded = await Presentation.fromZip(output);
    expect(await advancedShapeModels(reloaded.slides.getItem(0))).toEqual(beforeAdvancedModels);
  });

  it('resolves effective text styles inherited from slide layouts and masters', async () => {
    const presentation = await Presentation.fromZip(await makeInheritedTextStyleFixturePptx());
    const title = presentation.slides.getItem(0).shapes.getItem('TitleBox');

    expect(await title.geometry()).toEqual({});
    expect(await title.effectiveGeometry()).toEqual({
      left: 457200,
      top: 274638,
      width: 8229600,
      height: 1143000,
    });
    expect(await title.placeholder()).toEqual({
      type: 'title',
      index: '1',
      size: 'full',
      orientation: 'horz',
    });
    expect(await title.effectivePlaceholder()).toEqual({
      type: 'title',
      index: '1',
      size: 'full',
      orientation: 'horz',
    });
    expect(await title.text.body()).toEqual({ wrap: 'none', autoFit: 'shape' });
    expect(await title.text.effectiveBody()).toMatchObject({
      anchor: 'middle',
      wrap: 'none',
      autoFit: 'shape',
      insetLeft: 91440,
    });
    expect((await title.text.style()).color).toBeUndefined();
    expect(await title.text.effectiveStyle()).toMatchObject({
      color: '#22AA88',
      fontSize: 20,
    });
    expect(await title.text.effectiveRuns()).toEqual([
      expect.objectContaining({
        text: 'Hello ',
        style: expect.objectContaining({
          color: '#22AA88',
          fontSize: 20,
          bold: true,
        }),
      }),
      expect.objectContaining({
        text: 'Coco',
        style: expect.objectContaining({
          color: '#22AA88',
          italic: true,
        }),
      }),
    ]);

    const model = await presentation.toModel();
    expect(model.slides[0]!.layout?.master?.textStyles).toEqual({
      title: { color: '#3366FF' },
    });
  });
});

describe.skipIf(!process.env.COCO_PPTX_FIXTURE_DIR)('Presentation real PPTX fixtures', () => {
  const fixtureDir = process.env.COCO_PPTX_FIXTURE_DIR || '';
  const files = existsSync(fixtureDir)
    ? readdirSync(fixtureDir).filter(
        (name) => name.toLowerCase().endsWith('.pptx') && !name.startsWith('~$'),
      )
    : [];

  it('has real pptx fixtures available', () => {
    expect(files.length).toBeGreaterThan(0);
  });

  it.each(files)('parses and round-trips %s', async (name) => {
    const bytes = readFileSync(join(fixtureDir, name));
    const originalEntries = await zipEntryBytes(bytes);
    const presentation = await Presentation.fromZip(bytes, { presentationId: name });
    const model = await presentation.toModel();

    expect(model.id).toBe(name);
    expect(model.slides.length).toBeGreaterThan(0);
    expect(model.slides.every((slide, index) => slide.index === index)).toBe(true);
    const expected = await inspectSlideMotion(bytes);
    expect(model.slides.filter((slide) => slide.transition).length).toBe(expected.transitions);
    expect(model.slides.filter((slide) => slide.timing).length).toBe(expected.timing);
    if (expected.transitions > 0) {
      const transition = model.slides.find((slide) => slide.transition)?.transition;
      expect(transition?.speed ?? transition?.type ?? transition?.duration).toBeDefined();
    }
    if (expected.timing > 0) {
      expect(model.slides.find((slide) => slide.timing)?.timing?.hasTiming).toBe(true);
    }

    const output = await presentation.toZip();
    expect(output.byteLength).toBeGreaterThan(0);
    const roundTripEntries = await zipEntryBytes(output);
    expect([...roundTripEntries.keys()].sort()).toEqual([...originalEntries.keys()].sort());
    for (const [path, originalBytes] of originalEntries) {
      expect(bytesEqual(roundTripEntries.get(path), originalBytes), path).toBe(true);
    }

    const reloaded = await Presentation.fromZip(output, { presentationId: `${name}:roundtrip` });
    const reloadedModel = await reloaded.toModel();
    expect(reloadedModel.slides.length).toBe(model.slides.length);
  });

  it.each(files)('supports SDK writes on real fixture %s', async (name) => {
    const bytes = readFileSync(join(fixtureDir, name));
    const preservedParts = await selectedZipEntryBytes(bytes, isAdvancedPptPart);
    const presentation = await Presentation.fromZip(bytes, { presentationId: `${name}:edit` });
    const before = await presentation.properties();
    const slide = presentation.slides.getItem(0);
    const marker = `COCO_REAL_EDIT_${name}`;

    await slide.shapes.addTextBox(marker, { name: 'CocoRealTextBox' });
    const table = await slide.shapes.addTable(2, 2, {
      name: 'CocoRealTable',
      values: [
        ['A', 'B'],
        ['C', 'D'],
      ],
    });
    await table.table.setCellText(1, 1, marker);
    await table.table.addRow('end', ['E', 'F']);
    await table.table.addColumn('end', ['G', 'H', 'I']);
    await slide.shapes.addImage(
      { bytes: pngBytes(), mimeType: 'image/png' },
      { name: 'CocoRealImage' },
    );
    await slide.shapes.addShape('rect', {
      name: 'CocoRealRect',
      geometry: { left: 10, top: 20, width: 300, height: 400 },
    });

    for (const shape of await slide.shapes.items()) {
      if (await shape.text.hasText()) {
        await shape.text.setText(`${marker}_TEXT`);
        break;
      }
    }

    const output = await presentation.toZip();
    const editedEntries = await zipEntryBytes(output);
    for (const [path, originalBytes] of preservedParts) {
      expect(bytesEqual(editedEntries.get(path), originalBytes), path).toBe(true);
    }

    const edited = await Presentation.fromZip(output, {
      presentationId: `${name}:edited-roundtrip`,
    });
    const after = await edited.properties();
    expect(after.slideCount).toBe(before.slideCount);
    expect(after.shapeCount).toBeGreaterThanOrEqual(before.shapeCount + 4);
    expect(await edited.findText(marker)).not.toEqual([]);
    expect(await edited.slides.getItem(0).shapes.getItem('CocoRealTable').table.values()).toEqual([
      ['A', 'B', 'G'],
      ['C', marker, 'H'],
      ['E', 'F', 'I'],
    ]);
    expect(await edited.slides.getItem(0).shapes.getItem('CocoRealImage').image.bytes()).toEqual(
      pngBytes(),
    );
    expect(await edited.slides.getItem(0).shapes.getItem('CocoRealRect').kind()).toBe(
      'geometricShape',
    );
  });

  it('preserves advanced shapes on the edited slide in real fixtures', async () => {
    let exercised = 0;

    for (const name of files) {
      const bytes = readFileSync(join(fixtureDir, name));
      const originalAdvancedParts = await selectedZipEntryBytes(bytes, isAdvancedPptPart);
      const presentation = await Presentation.fromZip(bytes, {
        presentationId: `${name}:advanced-slide-preservation`,
      });
      const model = await presentation.toModel();
      const advancedSlide = model.slides.find((slide) => advancedShapeSnapshots(slide).length > 0);
      if (!advancedSlide) continue;

      exercised += 1;
      const beforeAdvancedShapes = advancedShapeSnapshots(advancedSlide);
      const slide = presentation.slides.getItem(advancedSlide.id);
      const textCandidate = advancedSlide.shapes.find(
        (shape) => shape.text?.text.trim() && !shape.table && !isAdvancedShapeKind(shape.kind),
      );

      await presentation.clearMutations();
      if (textCandidate) {
        await slide.shapes
          .getItem(textCandidate.id)
          .text.setText(`COCO_ADVANCED_SAFE_TEXT_${name}`);
      } else {
        await slide.shapes.addTextBox(`COCO_ADVANCED_SAFE_TEXT_${name}`, {
          name: `CocoAdvancedSafeText_${advancedSlide.index}`,
        });
      }

      const output = await presentation.toZip();
      const editedEntries = await zipEntryBytes(output);
      for (const [path, originalBytes] of originalAdvancedParts) {
        expect(bytesEqual(editedEntries.get(path), originalBytes), `${name}:${path}`).toBe(true);
      }

      const reloaded = await Presentation.fromZip(output, {
        presentationId: `${name}:advanced-slide-preservation-roundtrip`,
      });
      const reloadedSlide = (await reloaded.toModel()).slides.find(
        (slideModel) => slideModel.id === advancedSlide.id,
      );
      expect(reloadedSlide, name).toBeDefined();
      expect(advancedShapeSnapshots(reloadedSlide!), name).toEqual(beforeAdvancedShapes);
      expect((await presentation.getMutations()).length).toBeGreaterThan(0);
    }

    expect(exercised).toBeGreaterThan(0);
  });

  it.each(files)('edits existing real fixture content when present %s', async (name) => {
    const bytes = readFileSync(join(fixtureDir, name));
    const presentation = await Presentation.fromZip(bytes, {
      presentationId: `${name}:existing-edit`,
    });
    const model = await presentation.toModel();
    const textCandidate = model.slides
      .flatMap((slide) => slide.shapes.map((shape) => ({ slide, shape })))
      .find(({ shape }) => shape.text?.text.trim() && !shape.table);
    const tableCandidate = model.slides
      .flatMap((slide) => slide.shapes.map((shape) => ({ slide, shape })))
      .find(({ shape }) => shape.table);
    const pictureCandidate = model.slides
      .flatMap((slide) => slide.shapes.map((shape) => ({ slide, shape })))
      .find(({ shape }) => shape.kind === 'picture' && shape.image?.targetPath);

    if (textCandidate) {
      const marker = `COCO_EXISTING_TEXT_${name}`;
      const sample = textCandidate.shape
        .text!.text.trim()
        .slice(0, Math.min(8, textCandidate.shape.text!.text.trim().length));
      const matches = await presentation.findText(sample, { caseSensitive: true });
      expect(matches.length).toBeGreaterThan(0);
      await presentation.slides
        .getItem(textCandidate.slide.id)
        .shapes.getItem(textCandidate.shape.id)
        .text.setText(marker);
      const reloaded = await Presentation.fromZip(await presentation.toZip());
      expect(await reloaded.findText(marker, { caseSensitive: true })).not.toEqual([]);
      return;
    }

    if (tableCandidate?.shape.table) {
      const marker = `COCO_EXISTING_TABLE_${name}`;
      await presentation.slides
        .getItem(tableCandidate.slide.id)
        .shapes.getItem(tableCandidate.shape.id)
        .table.setCellText(0, 0, marker);
      const reloaded = await Presentation.fromZip(await presentation.toZip());
      expect(await reloaded.findText(marker, { caseSensitive: true })).not.toEqual([]);
      return;
    }

    if (pictureCandidate) {
      const bytesBefore = await presentation.slides
        .getItem(pictureCandidate.slide.id)
        .shapes.getItem(pictureCandidate.shape.id)
        .image.bytes();
      expect(bytesBefore.byteLength).toBeGreaterThan(0);
      expect((await presentation.exportImages()).length).toBeGreaterThan(0);
      return;
    }

    expect(model.slides.length).toBeGreaterThan(0);
  });

  it('diagnoses advanced object coverage across real fixtures', async () => {
    const reports: PptRealFixtureDiagnostic[] = [];

    for (const name of files) {
      const bytes = readFileSync(join(fixtureDir, name));
      const presentation = await Presentation.fromZip(bytes, {
        presentationId: `${name}:diagnostic`,
      });
      reports.push({
        name,
        package: await packageAdvancedCounts(bytes),
        model: summarizeAdvancedShapes(await presentation.toModel()),
      });
    }

    expect(reports.length).toBeGreaterThan(0);
    const totals = reports.reduce(
      (next, report) => ({
        chartParts: next.chartParts + report.package.chartParts,
        mediaParts: next.mediaParts + report.package.mediaParts,
        videoParts: next.videoParts + report.package.videoParts,
        audioParts: next.audioParts + report.package.audioParts,
        embeddingParts: next.embeddingParts + report.package.embeddingParts,
        diagramParts: next.diagramParts + report.package.diagramParts,
        chartShapes: next.chartShapes + report.model.chartShapes,
        mediaShapes: next.mediaShapes + report.model.mediaShapes,
        oleObjectShapes: next.oleObjectShapes + report.model.oleObjectShapes,
        smartArtShapes: next.smartArtShapes + report.model.smartArtShapes,
        unknownShapes: next.unknownShapes + report.model.unknownShapes,
      }),
      {
        chartParts: 0,
        mediaParts: 0,
        videoParts: 0,
        audioParts: 0,
        embeddingParts: 0,
        diagramParts: 0,
        chartShapes: 0,
        mediaShapes: 0,
        oleObjectShapes: 0,
        smartArtShapes: 0,
        unknownShapes: 0,
      },
    );

    expect(totals.chartParts).toBeGreaterThan(0);
    expect(totals.mediaParts).toBeGreaterThan(0);
    expect(totals.chartShapes).toBeGreaterThan(0);
    expect(totals.videoParts).toBeGreaterThan(0);
    expect(totals.audioParts).toBeGreaterThan(0);
    expect(totals.mediaShapes).toBeGreaterThan(0);
    expect(totals.oleObjectShapes).toBeGreaterThan(0);
    expect(totals.smartArtShapes).toBeGreaterThan(0);
    expect(totals.unknownShapes).toBeGreaterThan(0);

    for (const report of reports) {
      if (report.package.chartParts > 0) {
        expect(report.model.chartShapes, report.name).toBeGreaterThan(0);
      }
      if (report.package.videoParts + report.package.audioParts > 0) {
        expect(report.model.mediaShapes, report.name).toBeGreaterThan(0);
      }
      if (report.package.embeddingParts > 0) {
        expect(report.model.oleObjectShapes, report.name).toBeGreaterThan(0);
      }
      if (report.package.diagramParts > 0) {
        expect(report.model.smartArtShapes, report.name).toBeGreaterThan(0);
      }
    }
  });

  it('recognizes generated real video and audio media fixture', async () => {
    const name = 'media_objects_video_audio.pptx';
    expect(files).toContain(name);

    const bytes = readFileSync(join(fixtureDir, name));
    const counts = await packageAdvancedCounts(bytes);
    expect(counts.videoParts).toBe(1);
    expect(counts.audioParts).toBe(1);

    const presentation = await Presentation.fromZip(bytes, {
      presentationId: `${name}:media-fixture`,
    });
    const model = await presentation.toModel();
    const shapes = model.slides.flatMap((slide) => slide.shapes);

    expect(shapes.find((shape) => shape.name === 'SDK Video Object')).toMatchObject({
      kind: 'media',
      media: {
        relationshipId: 'rIdVideo',
        targetPath: 'ppt/media/video1.mp4',
        mediaType: 'video',
        contentType: 'video/mp4',
        filename: 'video1.mp4',
      },
    });
    expect(shapes.find((shape) => shape.name === 'SDK Audio Object')).toMatchObject({
      kind: 'media',
      media: {
        relationshipId: 'rIdAudio',
        targetPath: 'ppt/media/audio1.wav',
        mediaType: 'audio',
        contentType: 'audio/wav',
        filename: 'audio1.wav',
      },
    });
  });

  it('recognizes generated real OLE and SmartArt fixture', async () => {
    const name = 'advanced_objects_ole_smartart.pptx';
    expect(files).toContain(name);

    const bytes = readFileSync(join(fixtureDir, name));
    const counts = await packageAdvancedCounts(bytes);
    expect(counts.embeddingParts).toBe(1);
    expect(counts.diagramParts).toBe(4);

    const presentation = await Presentation.fromZip(bytes, {
      presentationId: `${name}:advanced-fixture`,
    });
    const model = await presentation.toModel();
    const shapes = model.slides.flatMap((slide) => slide.shapes);

    expect(shapes.find((shape) => shape.name === 'SDK Workbook OLE')).toMatchObject({
      kind: 'oleObject',
      oleObject: {
        relationshipId: 'rIdOle',
        targetPath: 'ppt/embeddings/oleObject1.bin',
        progId: 'Excel.Sheet.12',
        name: 'SDK Embedded Workbook',
      },
    });
    expect(shapes.find((shape) => shape.name === 'SDK SmartArt Process')).toMatchObject({
      kind: 'smartArt',
      smartArt: {
        relationshipIds: {
          dm: 'rIdDgmData',
          lo: 'rIdDgmLayout',
          qs: 'rIdDgmStyle',
          cs: 'rIdDgmColors',
        },
        targetPaths: {
          dm: 'ppt/diagrams/data1.xml',
          lo: 'ppt/diagrams/layout1.xml',
          qs: 'ppt/diagrams/quickStyle1.xml',
          cs: 'ppt/diagrams/colors1.xml',
        },
      },
    });
  });

  it('recognizes generated real unknown graphic fixture', async () => {
    const name = 'unknown_graphic_object.pptx';
    expect(files).toContain(name);

    const bytes = readFileSync(join(fixtureDir, name));
    const presentation = await Presentation.fromZip(bytes, {
      presentationId: `${name}:unknown-fixture`,
    });
    const model = await presentation.toModel();
    const unknown = model.slides
      .flatMap((slide) => slide.shapes)
      .find((shape) => shape.name === 'SDK Unknown Graphic');

    expect(unknown).toMatchObject({
      kind: 'unknown',
      raw: {
        slidePath: 'ppt/slides/slide1.xml',
        hostType: 'graphicFrame',
      },
    });
    expect(summarizeAdvancedShapes(model).unknownShapes).toBe(1);
  });

  it('builds a structured capability summary for every real fixture', async () => {
    const summaries = [];

    for (const name of files) {
      summaries.push(
        await summarizePptFixture(Presentation, name, readFileSync(join(fixtureDir, name))),
      );
    }

    expect(summaries.length).toBe(files.length);
    expect(summaries.every((summary) => summary.slideCount > 0)).toBe(true);
    expect(summaries.every((summary) => summary.shapeCount >= 0)).toBe(true);
    expect(summaries.every((summary) => summary.noEditRoundTripPreserved)).toBe(true);
    expect(summaries.every((summary) => summary.advancedPartsPreservedAfterSdkWrites)).toBe(true);
    expect(
      summaries.every((summary) => summary.advancedShapesPreservedOnEditedSlide !== false),
    ).toBe(true);
    expect(summaries.some((summary) => summary.advancedShapesPreservedOnEditedSlide === true)).toBe(
      true,
    );
    expect(summaries.some((summary) => summary.packageParts.mediaParts > 0)).toBe(true);
    expect(summaries.some((summary) => summary.packageParts.imageParts > 0)).toBe(true);
    expect(summaries.some((summary) => summary.packageParts.chartParts > 0)).toBe(true);
    expect(summaries.some((summary) => summary.advancedShapes.chartShapes > 0)).toBe(true);
    expect(summaries.some((summary) => summary.editableCandidates.text)).toBe(true);
    expect(summaries.some((summary) => summary.editableCandidates.image)).toBe(true);

    const totals = summaries.reduce(
      (next, summary) => ({
        slides: next.slides + summary.slideCount,
        shapes: next.shapes + summary.shapeCount,
        chartParts: next.chartParts + summary.packageParts.chartParts,
        mediaParts: next.mediaParts + summary.packageParts.mediaParts,
        imageParts: next.imageParts + summary.packageParts.imageParts,
        videoParts: next.videoParts + summary.packageParts.videoParts,
        audioParts: next.audioParts + summary.packageParts.audioParts,
        otherMediaParts: next.otherMediaParts + summary.packageParts.otherMediaParts,
        embeddingParts: next.embeddingParts + summary.packageParts.embeddingParts,
        diagramParts: next.diagramParts + summary.packageParts.diagramParts,
        chartShapes: next.chartShapes + summary.advancedShapes.chartShapes,
        mediaShapes: next.mediaShapes + summary.advancedShapes.mediaShapes,
        oleObjectShapes: next.oleObjectShapes + summary.advancedShapes.oleObjectShapes,
        smartArtShapes: next.smartArtShapes + summary.advancedShapes.smartArtShapes,
        unknownShapes: next.unknownShapes + summary.advancedShapes.unknownShapes,
      }),
      {
        slides: 0,
        shapes: 0,
        chartParts: 0,
        mediaParts: 0,
        imageParts: 0,
        videoParts: 0,
        audioParts: 0,
        otherMediaParts: 0,
        embeddingParts: 0,
        diagramParts: 0,
        chartShapes: 0,
        mediaShapes: 0,
        oleObjectShapes: 0,
        smartArtShapes: 0,
        unknownShapes: 0,
      },
    );

    expect(totals.slides).toBeGreaterThanOrEqual(summaries.length);
    expect(totals.shapes).toBeGreaterThan(0);
    expect(totals.chartParts).toBeGreaterThan(0);
    expect(totals.mediaParts).toBeGreaterThan(0);
    expect(totals.imageParts + totals.videoParts + totals.audioParts + totals.otherMediaParts).toBe(
      totals.mediaParts,
    );
    expect(totals.videoParts).toBeGreaterThan(0);
    expect(totals.audioParts).toBeGreaterThan(0);
    expect(totals.embeddingParts).toBeGreaterThan(0);
    expect(totals.diagramParts).toBeGreaterThan(0);
    expect(totals.chartShapes).toBeGreaterThan(0);
    expect(totals.mediaShapes).toBeGreaterThan(0);
    expect(totals.oleObjectShapes).toBeGreaterThan(0);
    expect(totals.smartArtShapes).toBeGreaterThan(0);
    expect(totals.unknownShapes).toBeGreaterThan(0);
  });
});

async function makeInheritedTextStyleFixturePptx(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await makeFixturePptx());
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
  zip.file(
    'ppt/slides/slide1.xml',
    slide
      .replace(
        '<a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr>',
        '<a:rPr sz="2000" b="1"/>',
      )
      .replace(
        '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>',
        '<p:spPr/>',
      ),
  );
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld name="Title and Content">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="LayoutTitle"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:rPr><a:t>Layout title</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`,
  );
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld name="Office Theme">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="MasterTitle"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr anchor="ctr" lIns="91440"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:t>Master title</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle>
  </p:txStyles>
</p:sldMaster>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeFixturePptx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypesXml(), { createFolders: false });
  zip.file('_rels/.rels', rootRelsXml(), { createFolders: false });
  zip.file('ppt/presentation.xml', presentationXml(), { createFolders: false });
  zip.file('ppt/_rels/presentation.xml.rels', presentationRelsXml(), { createFolders: false });
  zip.file('ppt/slides/slide1.xml', slideXml(), { createFolders: false });
  zip.file('ppt/slides/_rels/slide1.xml.rels', slideRelsXml(), { createFolders: false });
  zip.file('ppt/notesSlides/notesSlide1.xml', notesSlideXml(), { createFolders: false });
  zip.file('ppt/comments/comment1.xml', commentsXml(), { createFolders: false });
  zip.file('ppt/commentAuthors.xml', commentAuthorsXml(), { createFolders: false });
  zip.file('ppt/charts/chart1.xml', chartXml(), { createFolders: false });
  zip.file('ppt/slideLayouts/slideLayout1.xml', slideLayoutXml(), { createFolders: false });
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', slideLayoutRelsXml(), {
    createFolders: false,
  });
  zip.file('ppt/slideMasters/slideMaster1.xml', slideMasterXml(), { createFolders: false });
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', slideMasterRelsXml(), {
    createFolders: false,
  });
  zip.file('ppt/theme/theme1.xml', themeXml(), { createFolders: false });
  zip.file('ppt/media/image1.png', new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), {
    createFolders: false,
  });
  zip.file('ppt/media/video1.mp4', new Uint8Array([0, 0, 0, 24, 102, 116, 121, 112]), {
    createFolders: false,
  });
  zip.file('ppt/embeddings/oleObject1.bin', new Uint8Array([208, 207, 17, 224]), {
    createFolders: false,
  });
  zip.file('ppt/diagrams/data1.xml', diagramPartXml('dataModel'), { createFolders: false });
  zip.file('ppt/diagrams/layout1.xml', diagramPartXml('layoutDef'), { createFolders: false });
  zip.file('ppt/diagrams/quickStyle1.xml', diagramPartXml('styleDef'), { createFolders: false });
  zip.file('ppt/diagrams/colors1.xml', diagramPartXml('colorsDef'), { createFolders: false });
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

async function makeSvgBlipFixturePptx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
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
  <p:sldSz cx="12192000" cy="6858000"/>
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
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="2" name="SvgLogo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill>
          <a:blip>
            <a:extLst>
              <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
                <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId2"/>
              </a:ext>
            </a:extLst>
          </a:blip>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>`,
    { createFolders: false },
  );
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.svg"/>
</Relationships>`,
    { createFolders: false },
  );
  zip.file('ppt/media/logo.svg', svgBytes(), { createFolders: false });
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

async function makeCustomGeometryFixturePptx(): Promise<Uint8Array> {
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
  <p:sldSz cx="12192000" cy="6858000"/>
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
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="FreeformCircle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="100000" y="200000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>
          <a:custGeom>
            <a:avLst/>
            <a:gdLst/>
            <a:pathLst>
              <a:path w="1000" h="1000">
                <a:moveTo><a:pt x="1000" y="500"/></a:moveTo>
                <a:cubicBezTo><a:pt x="1000" y="776"/><a:pt x="776" y="1000"/><a:pt x="500" y="1000"/></a:cubicBezTo>
                <a:cubicBezTo><a:pt x="224" y="1000"/><a:pt x="0" y="776"/><a:pt x="0" y="500"/></a:cubicBezTo>
                <a:cubicBezTo><a:pt x="0" y="224"/><a:pt x="224" y="0"/><a:pt x="500" y="0"/></a:cubicBezTo>
                <a:cubicBezTo><a:pt x="776" y="0"/><a:pt x="1000" y="224"/><a:pt x="1000" y="500"/></a:cubicBezTo>
                <a:close/>
              </a:path>
            </a:pathLst>
          </a:custGeom>
          <a:ln w="101600"><a:solidFill><a:srgbClr val="133E82"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="ArcShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1400000" y="200000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>
          <a:custGeom>
            <a:avLst/>
            <a:gdLst/>
            <a:pathLst>
              <a:path w="1000" h="1000">
                <a:moveTo><a:pt x="1000" y="500"/></a:moveTo>
                <a:arcTo wR="500" hR="500" stAng="0" swAng="5400000"/>
              </a:path>
            </a:pathLst>
          </a:custGeom>
          <a:noFill/>
          <a:ln w="63500"><a:solidFill><a:srgbClr val="C43E1C"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
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

function pngBytes(): Uint8Array {
  return new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
}

function svgBytes(): Uint8Array {
  return new TextEncoder().encode(
    '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="#36f"/></svg>',
  );
}

interface PptRealFixtureDiagnostic {
  name: string;
  package: PptPackageAdvancedCounts;
  model: PptAdvancedShapeSummary;
}

interface PptPackageAdvancedCounts {
  chartParts: number;
  mediaParts: number;
  imageParts: number;
  videoParts: number;
  audioParts: number;
  otherMediaParts: number;
  embeddingParts: number;
  diagramParts: number;
}

interface PptAdvancedShapeSummary {
  chartShapes: number;
  mediaShapes: number;
  oleObjectShapes: number;
  smartArtShapes: number;
  unknownShapes: number;
}

function advancedShapeSnapshots(slide: PptPresentationModel['slides'][number]) {
  return slide.shapes.flatMap((shape) => advancedShapeSnapshotsFromShape(shape));
}

function advancedShapeSnapshotsFromShape(shape: PptShapeModel): unknown[] {
  const own = isAdvancedShapeKind(shape.kind)
    ? [
        {
          id: shape.id,
          name: shape.name,
          kind: shape.kind,
          geometry: shape.geometry,
          chart: shape.chart,
          media: shape.media,
          oleObject: shape.oleObject,
          smartArt: shape.smartArt,
          raw: shape.raw,
        },
      ]
    : [];
  return [
    ...own,
    ...(shape.children ?? []).flatMap((child) => advancedShapeSnapshotsFromShape(child)),
  ];
}

function isAdvancedShapeKind(kind: PptShapeModel['kind']): boolean {
  return (
    kind === 'chart' ||
    kind === 'media' ||
    kind === 'oleObject' ||
    kind === 'smartArt' ||
    kind === 'unknown'
  );
}

async function advancedShapeModels(slide: Awaited<ReturnType<Presentation['slides']['getItem']>>) {
  return Promise.all(
    ['SalesChart', 'Video 1', 'Workbook OLE', 'Process SmartArt', 'Group 1'].map(async (name) => {
      const model = await slide.shapes.getItem(name).toModel();
      return {
        name,
        id: model.id,
        kind: model.kind,
        geometry: model.geometry,
        chart: model.chart,
        media: model.media,
        oleObject: model.oleObject,
        smartArt: model.smartArt,
        children: model.children?.map((child) => ({
          id: child.id,
          name: child.name,
          kind: child.kind,
          hyperlink: child.hyperlink,
        })),
        raw: model.raw,
      };
    }),
  );
}

function diagramPartXml(localName: string): string {
  return `<dgm:${localName} xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>`;
}

function contentTypesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="mp4" ContentType="video/mp4"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
  <Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>
  <Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>
  <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;
}

function rootRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
}

function presentationXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`;
}

function presentationRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>
</Relationships>`;
}

function slideRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/coco" TargetMode="External"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/video1.mp4"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
  <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
</Relationships>`;
}

function slideLayoutRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
}

function slideMasterRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
}

function slideXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
  <p:cSld>
    <p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TitleBox"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="1" sz="full" orient="horz"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>Hello </a:t></a:r><a:r><a:rPr i="1"><a:hlinkClick r:id="rId2" tooltip="Coco link"/></a:rPr><a:t>Coco</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp useBgFill="1">
        <p:nvSpPr><p:cNvPr id="13" name="SingleBoldBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="6000" y="2600"/><a:ext cx="1600" cy="400"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buFont typeface="Wingdings"/><a:buChar char="ü"/></a:pPr><a:r><a:rPr sz="1600" b="1"/><a:t>Single bold</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="3" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
        <p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm></p:spPr>
      </p:pic>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="4" name="DataTable"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="500" y="600"/><a:ext cx="700" cy="800"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tr h="100">
                <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL><a:noFill/></a:lnL><a:lnT w="25400"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/></a:lnT></a:tcPr></a:tc>
                <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>B1</a:t></a:r></a:p><a:p/><a:p><a:r><a:t>B1 details</a:t></a:r></a:p><a:p/><a:p/></a:txBody></a:tc>
              </a:tr>
              <a:tr h="100">
                <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>A2</a:t></a:r></a:p></a:txBody></a:tc>
                <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>B2</a:t></a:r></a:p></a:txBody></a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="5" name="SalesChart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="1000" y="1200"/><a:ext cx="2000" cy="1200"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId5"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="6" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="1200"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1200"/></a:xfrm><a:solidFill><a:srgbClr val="05E1FF"><a:alpha val="15000"/></a:srgbClr></a:solidFill></p:grpSpPr>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="7" name="GroupedText"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="1100" y="1100"/><a:ext cx="800" cy="300"/></a:xfrm></p:spPr>
          <p:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Grouped Coco</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="8" name="GroupedRect"><a:hlinkClick r:id="rId2"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="1200" y="1500"/><a:ext cx="500" cy="400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:grpFill/></p:spPr>
        </p:sp>
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
      <p:pic>
        <p:nvPicPr><p:cNvPr id="9" name="Video 1"/><p:cNvPicPr/><p:nvPr><a:videoFile r:link="rId7"/></p:nvPr></p:nvPicPr>
        <p:blipFill><a:blip/></p:blipFill>
        <p:spPr><a:xfrm><a:off x="1600" y="1800"/><a:ext cx="900" cy="600"/></a:xfrm></p:spPr>
      </p:pic>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="10" name="Workbook OLE"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="2600" y="1800"/><a:ext cx="900" cy="600"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
            <p:oleObj r:id="rId8" progId="Excel.Sheet.12" name="Workbook"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="11" name="Process SmartArt"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="3600" y="1800"/><a:ext cx="900" cy="600"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
            <dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId9" r:lo="rId10" r:qs="rId11" r:cs="rId12"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="12" name="MergedTable"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="4600" y="1800"/><a:ext cx="1200" cy="700"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblGrid><a:gridCol w="600"/><a:gridCol w="600"/></a:tblGrid>
              <a:tr h="100">
                <a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:t>Merged A</a:t></a:r></a:p></a:txBody></a:tc>
                <a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p><a:r><a:t></a:t></a:r></a:p></a:txBody></a:tc>
              </a:tr>
              <a:tr h="100">
                <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>C1</a:t></a:r></a:p></a:txBody></a:tc>
                <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>D1</a:t></a:r></a:p></a:txBody></a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="14" name="VerticalText"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="6000" y="200"/><a:ext cx="240" cy="2000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
        <p:txBody><a:bodyPr vert="eaVert" wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="1200"><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs></a:gsLst></a:gradFill><a:effectLst><a:outerShdw blurRad="60000" dist="120000" dir="5400000"><a:srgbClr val="FFFFFF"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst></a:rPr><a:t>WORK SUMMARY</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:transition spd="fast" p14:dur="750" advClick="1" advTm="3000"><p:fade/></p:transition>
  <p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq nodeType="mainSeq"><p:cTn id="2" dur="indefinite"/></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>
</p:sld>`;
}

function slideLayoutXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Title and Content">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;
}

function slideMasterXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Office Theme">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldMaster>`;
}

function themeXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Coco Theme">
  <a:themeElements>
    <a:clrScheme name="Coco Colors">
      <a:dk1><a:srgbClr val="111111"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="3366FF"/></a:accent1>
      <a:accent2><a:srgbClr val="22AA88"/></a:accent2>
    </a:clrScheme>
    <a:fontScheme name="Coco Fonts">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;
}

function notesSlideXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Speaker note for Coco</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`;
}

function chartXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarterly Sales</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:grouping val="stacked"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>12</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:order val="1"/>
          <c:tx><c:v>South</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>8</c:v></c:pt><c:pt idx="1"><c:v>14</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></c:spPr>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
}

function commentAuthorsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="0" name="Coco Reviewer" initials="CR" lastIdx="1" clrIdx="0"/>
</p:cmAuthorLst>`;
}

function commentsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2026-08-25T09:00:00Z" idx="1">
    <p:pos x="100" y="200"/>
    <p:text>Review this grouped link</p:text>
  </p:cm>
</p:cmLst>`;
}
