import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'vitest';
import {
  Presentation,
  diffPptPresentationModels,
  planPptWriteback,
  type PptPresentationModel,
  type PptMutation,
} from '../index.js';

const mutationSource = readFileSync(join(import.meta.dirname, '../mutations.ts'), 'utf8');

const BASE_MODEL: PptPresentationModel = {
  id: 'deck',
  slideSize: { width: 9144000, height: 5143500 },
  slides: [
    {
      id: '256',
      index: 0,
      title: 'Intro',
      shapes: [
        {
          id: '2',
          name: 'Title',
          kind: 'textBox',
          geometry: { left: 100, top: 100, width: 1000, height: 300 },
          text: { text: 'Intro' },
        },
      ],
    },
  ],
};

const ALL_MUTATION_FIXTURES: PptMutation[] = [
  { type: 'shape.add', slideId: '256', shapeId: '3', kind: 'textBox' },
  { type: 'shape.delete', slideId: '256', shapeId: '3' },
  { type: 'shape.name.set', slideId: '256', shapeId: '2', name: 'Renamed' },
  { type: 'shape.zOrder.set', slideId: '256', shapeId: '2', position: 'front' },
  {
    type: 'shape.geometry.set',
    slideId: '256',
    shapeId: '2',
    geometry: { left: 10, top: 20 },
  },
  {
    type: 'shape.style.set',
    slideId: '256',
    shapeId: '2',
    style: { fillColor: '#112233' },
  },
  {
    type: 'shape.hyperlink.set',
    slideId: '256',
    shapeId: '2',
    hyperlink: { target: 'https://example.com' },
  },
  { type: 'shape.hyperlink.clear', slideId: '256', shapeId: '2' },
  { type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'Updated' },
  { type: 'shape.text.body.set', slideId: '256', shapeId: '2', body: { wrap: 'square' } },
  {
    type: 'shape.text.style.set',
    slideId: '256',
    shapeId: '2',
    style: { bold: true },
  },
  {
    type: 'shape.text.paragraphs.set',
    slideId: '256',
    shapeId: '2',
    paragraphs: [{ runs: [{ text: 'Updated' }] }],
  },
  {
    type: 'shape.text.run.style.set',
    slideId: '256',
    shapeId: '2',
    runIndex: 0,
    style: { italic: true },
  },
  {
    type: 'shape.text.run.hyperlink.set',
    slideId: '256',
    shapeId: '2',
    runIndex: 0,
    hyperlink: { target: 'https://example.com/run' },
  },
  { type: 'shape.text.run.hyperlink.clear', slideId: '256', shapeId: '2', runIndex: 0 },
  {
    type: 'shape.text.replace',
    slideId: '256',
    shapeId: '2',
    search: 'Intro',
    replacement: 'Updated',
    count: 1,
  },
  { type: 'slide.notes.set', slideId: '256', text: 'Notes' },
  { type: 'slide.notes.clear', slideId: '256' },
  { type: 'slide.comment.add', slideId: '256', comment: { text: 'Comment' } },
  { type: 'slide.comment.delete', slideId: '256', commentId: '1' },
  { type: 'slide.comments.clear', slideId: '256' },
  {
    type: 'shape.chart.data.set',
    slideId: '256',
    shapeId: '2',
    data: { series: [{ values: [1, 2] }] },
  },
  { type: 'shape.image.replace', slideId: '256', shapeId: '2', imageRef: 'ppt/media/image1.png' },
  {
    type: 'shape.image.style.set',
    slideId: '256',
    shapeId: '2',
    style: { crop: { left: 0.1 }, opacity: 0.5 },
  },
  { type: 'table.cell.text.set', slideId: '256', shapeId: '2', row: 0, column: 0, text: 'Cell' },
  {
    type: 'table.cell.style.set',
    slideId: '256',
    shapeId: '2',
    row: 0,
    column: 0,
    style: { fillColor: '#445566' },
  },
  {
    type: 'table.cell.text.paragraphs.set',
    slideId: '256',
    shapeId: '2',
    row: 0,
    column: 0,
    paragraphs: [{ runs: [{ text: 'Cell' }] }],
  },
  {
    type: 'table.cell.text.run.style.set',
    slideId: '256',
    shapeId: '2',
    row: 0,
    column: 0,
    runIndex: 0,
    style: { color: '#778899' },
  },
  { type: 'table.values.set', slideId: '256', shapeId: '2', values: [['A']] },
  { type: 'table.row.add', slideId: '256', shapeId: '2', index: 1, values: ['A'] },
  { type: 'table.row.delete', slideId: '256', shapeId: '2', index: 1 },
  { type: 'table.column.add', slideId: '256', shapeId: '2', index: 1, values: ['A'] },
  { type: 'table.column.delete', slideId: '256', shapeId: '2', index: 1 },
];

function cloneModel(model = BASE_MODEL): PptPresentationModel {
  return structuredClone(model);
}

describe('PPT writeback planner', () => {
  test('returns no-op for unchanged presentation models', () => {
    const plan = planPptWriteback(cloneModel(), cloneModel(), {});

    expect(plan.strategy).toBe('no-op');
    expect(plan.operations).toEqual([]);
    expect(plan.diff.changed).toBe(false);
  });

  test('classifies slide and shape modifications', () => {
    const after = cloneModel();
    after.slides[0]!.title = 'Updated';
    after.slides[0]!.shapes[0]!.text = { text: 'Updated' };

    const diff = diffPptPresentationModels(cloneModel(), after);

    expect(diff.changed).toBe(true);
    expect(diff.slides).toHaveLength(1);
    expect(diff.slides[0]).toMatchObject({
      kind: 'modified',
      slideId: '256',
      beforeIndex: 0,
      afterIndex: 0,
      beforeTitle: 'Intro',
      afterTitle: 'Updated',
      propertiesChanged: true,
      moved: false,
      shapeOrderChanged: false,
    });
    expect(diff.slides[0]!.shapeChanges).toHaveLength(1);
    expect(diff.slides[0]!.shapeChanges[0]).toMatchObject({
      kind: 'modified',
      shapeId: '2',
      beforeName: 'Title',
      afterName: 'Title',
      beforeKind: 'textBox',
      afterKind: 'textBox',
    });
  });

  test('classifies slide add, delete, move, and shape order changes', () => {
    const before = cloneModel();
    before.slides.push({
      id: '257',
      index: 1,
      title: 'Remove me',
      shapes: [],
    });
    const after = cloneModel();
    after.slides[0]!.index = 1;
    after.slides[0]!.shapes.push({
      id: '3',
      name: 'Subtitle',
      kind: 'textBox',
      geometry: {},
      text: { text: 'Second' },
    });
    after.slides[0]!.shapes.reverse();
    after.slides.push({
      id: '258',
      index: 0,
      title: 'Added',
      shapes: [],
    });

    const diff = diffPptPresentationModels(before, after);

    expect(
      diff.slides.map((slide) => [slide.slideId, slide.kind, slide.moved, slide.shapeOrderChanged]),
    ).toEqual([
      ['258', 'added', false, false],
      ['256', 'modified', true, true],
      ['257', 'deleted', false, false],
    ]);
    expect(diff.slides[1]!.shapeChanges.map((shape) => shape.kind)).toEqual(['added']);
  });

  test('prefers presentation package replacement in auto mode when supported', () => {
    const after = cloneModel();
    after.title = 'Package update';

    const plan = planPptWriteback(cloneModel(), after, {
      canReplacePresentationPackage: true,
      canReplayMutations: true,
    });

    expect(plan.strategy).toBe('presentation-package-replace');
    expect(plan.operations).toEqual([{ type: 'replacePresentationPackage' }]);
  });

  test('plans slide package replacement when requested and supported', () => {
    const after = cloneModel();
    after.slides[0]!.shapes[0]!.text = { text: 'Slide package update' };

    const plan = planPptWriteback(
      cloneModel(),
      after,
      {
        canReplaceSlides: true,
      },
      { preference: 'slide-package' },
    );

    expect(plan.strategy).toBe('slide-package-replace');
    expect(plan.operations).toHaveLength(1);
    expect(plan.operations[0]).toMatchObject({
      type: 'replaceSlidePackage',
      slideId: '256',
      slideIndex: 0,
      targetSlide: {
        id: '256',
        shapes: [
          {
            id: '2',
            text: { text: 'Slide package update' },
          },
        ],
      },
    });
    expect(plan.warnings).toEqual([
      'slide package replacement is adapter-defined and must preserve relationships, layouts, and media parts',
    ]);
  });

  test('plans slide lifecycle operations when host declares matching capabilities', () => {
    const before = cloneModel();
    before.slides.push({ id: '257', index: 1, title: 'Delete', shapes: [] });
    const after = cloneModel();
    after.slides[0]!.index = 1;
    after.slides.push({ id: '258', index: 0, title: 'Insert', shapes: [] });

    const plan = planPptWriteback(
      before,
      after,
      {
        canInsertSlides: true,
        canDeleteSlides: true,
        canReorderSlides: true,
      },
      { preference: 'slide-package' },
    );

    expect(plan.strategy).toBe('slide-package-replace');
    expect(plan.operations.map((operation) => operation.type)).toEqual([
      'insertSlidePackage',
      'moveSlide',
      'deleteSlide',
    ]);
  });

  test('plans mutation replay when requested and supported', () => {
    const after = cloneModel();
    after.slides[0]!.shapes[0]!.text = { text: 'Mutation update' };

    const plan = planPptWriteback(
      cloneModel(),
      after,
      {
        canReplayMutations: true,
        supportedMutationTypes: ['shape.text.set'],
        maxMutationCount: 2,
      },
      {
        preference: 'mutation',
        mutations: [
          { type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'Mutation update' },
        ],
      },
    );

    expect(plan.strategy).toBe('mutation-replay');
    expect(plan.operations).toEqual([
      {
        type: 'replayMutations',
        mutations: [
          { type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'Mutation update' },
        ],
      },
    ]);
  });

  test('has planner fixture payloads for every SDK mutation type', () => {
    expect(ALL_MUTATION_FIXTURES.map((mutation) => mutation.type).sort()).toEqual(
      extractMutationTypeStrings(mutationSource).sort(),
    );
  });

  test('accepts every SDK mutation type when host allow-list declares support', () => {
    const after = cloneModel();
    after.slides[0]!.shapes[0]!.text = { text: 'All mutation types allowed' };

    const plan = planPptWriteback(
      cloneModel(),
      after,
      {
        canReplayMutations: true,
        supportedMutationTypes: ALL_MUTATION_FIXTURES.map((mutation) => mutation.type),
        maxMutationCount: ALL_MUTATION_FIXTURES.length,
      },
      { preference: 'mutation', mutations: ALL_MUTATION_FIXTURES },
    );

    expect(plan.strategy).toBe('mutation-replay');
    expect(plan.operations).toEqual([
      { type: 'replayMutations', mutations: ALL_MUTATION_FIXTURES },
    ]);
  });

  test('accepts replay without an allow-list and returns defensive mutation payloads', () => {
    const after = cloneModel();
    after.slides[0]!.shapes[0]!.text = { text: 'Loose replay' };
    const mutations: PptMutation[] = [
      { type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'Loose replay' },
    ];

    const plan = planPptWriteback(
      cloneModel(),
      after,
      {
        canReplayMutations: true,
      },
      { preference: 'mutation', mutations },
    );

    expect(plan.strategy).toBe('mutation-replay');
    const operation = plan.operations[0];
    if (operation?.type !== 'replayMutations') throw new Error('expected replayMutations');

    operation.mutations[0] = {
      type: 'shape.text.set',
      slideId: '256',
      shapeId: '2',
      text: 'Mutated by caller',
    };

    expect(mutations).toEqual([
      { type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'Loose replay' },
    ]);
  });

  test('warns when mutation replay is asked to cover presentation or media changes', () => {
    const after = cloneModel();
    after.title = 'Deck title changed';
    after.media = [{ path: 'ppt/media/image1.png', contentType: 'image/png', byteLength: 4 }];
    after.slides[0]!.shapes[0]!.text = { text: 'Mutation with warnings' };

    const plan = planPptWriteback(
      cloneModel(),
      after,
      {
        canReplayMutations: true,
        supportedMutationTypes: ['shape.text.set'],
      },
      {
        preference: 'mutation',
        mutations: [
          {
            type: 'shape.text.set',
            slideId: '256',
            shapeId: '2',
            text: 'Mutation with warnings',
          },
        ],
      },
    );

    expect(plan.strategy).toBe('mutation-replay');
    expect(plan.warnings).toEqual([
      'mutation replay must validate presentation-level changes before applying host operations',
      'mutation replay must validate media relationships and package image parts',
    ]);
  });

  test('reports unsupported reasons when capabilities cannot express the diff', () => {
    const after = cloneModel();
    after.slides[0]!.shapes.push({
      id: '3',
      name: 'Added Shape',
      kind: 'textBox',
      geometry: {},
      text: { text: 'Added' },
    });

    const plan = planPptWriteback(
      cloneModel(),
      after,
      { canReplayMutations: true },
      { preference: 'mutation' },
    );

    expect(plan.strategy).toBe('unsupported');
    expect(plan.unsupportedReasons).toEqual(['mutation replay requires SDK mutation records']);
  });

  test('rejects unsupported mutation types and replay thresholds', () => {
    const after = cloneModel();
    after.slides[0]!.shapes[0]!.text = { text: 'Unsupported mutation' };
    const mutations = [
      {
        type: 'shape.text.set' as const,
        slideId: '256',
        shapeId: '2',
        text: 'Unsupported mutation',
      },
    ];

    expect(
      planPptWriteback(
        cloneModel(),
        after,
        {
          canReplayMutations: true,
          supportedMutationTypes: ['shape.add'],
        },
        { preference: 'mutation', mutations },
      ).unsupportedReasons,
    ).toEqual(['host cannot replay mutation types: shape.text.set']);

    expect(
      planPptWriteback(
        cloneModel(),
        after,
        {
          canReplayMutations: true,
          maxMutationCount: 0,
        },
        { preference: 'mutation', mutations },
      ).unsupportedReasons,
    ).toEqual(['mutation count 1 exceeds maxMutationCount 0']);
  });

  test('slide package target slide payload is a defensive clone', () => {
    const after = cloneModel();
    after.slides[0]!.title = 'Planned';

    const plan = planPptWriteback(
      cloneModel(),
      after,
      {
        canReplaceSlides: true,
      },
      { preference: 'slide-package' },
    );
    const replace = plan.operations[0];
    if (replace?.type !== 'replaceSlidePackage') throw new Error('expected replaceSlidePackage');

    replace.targetSlide.title = 'Mutated';

    expect(after.slides[0]?.title).toBe('Planned');
  });

  test.skipIf(!process.env.COCO_PPTX_FIXTURE_DIR)('plans mutation replay from SDK edits on a real pptx fixture', async () => {
    const fixtureDir = process.env.COCO_PPTX_FIXTURE_DIR || '';
    const files = existsSync(fixtureDir)
      ? readdirSync(fixtureDir).filter((name) => name.toLowerCase().endsWith('.pptx'))
      : [];
    expect(files.length).toBeGreaterThan(0);

    const presentation = await Presentation.fromZip(readFileSync(join(fixtureDir, files[0]!)), {
      presentationId: 'real-writeback-plan',
    });
    const before = await presentation.toModel();
    await presentation.clearMutations();

    await presentation.slides
      .getItem(0)
      .shapes.addTextBox('COCO_WRITEBACK_PLAN', { name: 'CocoWritebackPlanText' });

    const after = await presentation.toModel();
    const mutations = await presentation.getMutations();
    const plan = planPptWriteback(
      before,
      after,
      {
        canReplayMutations: true,
        supportedMutationTypes: ['shape.add'],
      },
      { preference: 'mutation', mutations },
    );

    expect(plan.strategy).toBe('mutation-replay');
    expect(plan.operations).toHaveLength(1);
    expect(plan.operations[0]).toMatchObject({
      type: 'replayMutations',
      mutations: [{ type: 'shape.add', kind: 'textBox' }],
    });
  });
});

function extractMutationTypeStrings(source: string): string[] {
  return [...source.matchAll(/type: '([a-z]+(?:\.[a-zA-Z]+)+)'/g)].map((match) => match[1]!);
}
