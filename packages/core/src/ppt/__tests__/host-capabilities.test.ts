import { describe, expect, test } from 'vitest';
import {
  createPptHostWritebackProfileFromProbeReport,
  createPptHostWritebackProfile,
  getDefaultPptHostProbeChecks,
  getPptWritebackOperationRequirements,
  planPptWriteback,
  runPptHostProbe,
  summarizePptHostProbeReport,
  validatePptWritebackPlan,
  type PptHostProbeReport,
  type PptPresentationModel,
} from '../index.js';

const BASE_MODEL: PptPresentationModel = {
  id: 'deck',
  slides: [
    {
      id: '256',
      index: 0,
      shapes: [
        {
          id: '2',
          name: 'Title',
          kind: 'textBox',
          geometry: {},
          text: { text: 'Before' },
        },
      ],
    },
  ],
};

function changedModel(): PptPresentationModel {
  const model = structuredClone(BASE_MODEL);
  model.slides[0]!.shapes[0]!.text = { text: 'After' };
  return model;
}

describe('PPT host writeback capabilities', () => {
  test('uses conservative unproven profiles for live presentation hosts', () => {
    const microsoft = createPptHostWritebackProfile('microsoft-powerpoint');
    const wps = createPptHostWritebackProfile('wps-presentation');
    const onlyoffice = createPptHostWritebackProfile('onlyoffice-presentation');

    expect(microsoft).toMatchObject({
      host: 'microsoft-powerpoint',
      status: 'unproven',
      capabilities: {},
    });
    expect(wps).toMatchObject({ host: 'wps-presentation', status: 'unproven', capabilities: {} });
    expect(onlyoffice).toMatchObject({
      host: 'onlyoffice-presentation',
      status: 'unproven',
      capabilities: {},
    });
  });

  test('declares desktop package replacement as the only built-in supported host profile', () => {
    const profile = createPptHostWritebackProfile('desktop-package');

    expect(profile.status).toBe('supported');
    expect(profile.capabilities).toEqual({ canReplacePresentationPackage: true });

    const plan = planPptWriteback(BASE_MODEL, changedModel(), profile.capabilities);
    expect(plan.strategy).toBe('presentation-package-replace');
    expect(validatePptWritebackPlan(plan, profile)).toEqual({
      ok: true,
      unsupportedReasons: [],
      warnings: [],
    });
  });

  test('allows adapters to declare probe-backed capability overrides', () => {
    const profile = createPptHostWritebackProfile('wps-presentation', {
      status: 'probed',
      capabilities: {
        canReplayMutations: true,
        supportedMutationTypes: ['shape.text.set'],
        maxMutationCount: 3,
      },
      notes: ['Text replay probe passed in local WPS host.'],
    });

    const plan = planPptWriteback(BASE_MODEL, changedModel(), profile.capabilities, {
      preference: 'mutation',
      mutations: [{ type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'After' }],
    });

    expect(profile.status).toBe('probed');
    expect(plan.strategy).toBe('mutation-replay');
    expect(validatePptWritebackPlan(plan, profile)).toEqual({
      ok: true,
      unsupportedReasons: [],
      warnings: [],
    });
  });

  test('maps writeback operations to required host capabilities', () => {
    expect(getPptWritebackOperationRequirements({ type: 'replacePresentationPackage' })).toEqual([
      'canReplacePresentationPackage',
    ]);
    expect(
      getPptWritebackOperationRequirements({
        type: 'replaceSlidePackage',
        slideId: '256',
        slideIndex: 0,
        targetSlide: BASE_MODEL.slides[0]!,
      }),
    ).toEqual(['canReplaceSlides']);
    expect(
      getPptWritebackOperationRequirements({
        type: 'insertSlidePackage',
        slideId: '257',
        slideIndex: 1,
        targetSlide: { id: '257', index: 1, shapes: [] },
      }),
    ).toEqual(['canInsertSlides']);
    expect(
      getPptWritebackOperationRequirements({ type: 'deleteSlide', slideId: '256', slideIndex: 0 }),
    ).toEqual(['canDeleteSlides']);
    expect(
      getPptWritebackOperationRequirements({
        type: 'moveSlide',
        slideId: '256',
        fromIndex: 0,
        toIndex: 1,
      }),
    ).toEqual(['canReorderSlides']);
    expect(
      getPptWritebackOperationRequirements({
        type: 'replayMutations',
        mutations: [{ type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'After' }],
      }),
    ).toEqual(['canReplayMutations']);
  });

  test('validates plan operations against a host profile', () => {
    const plan = planPptWriteback(
      BASE_MODEL,
      changedModel(),
      {
        canReplayMutations: true,
      },
      {
        preference: 'mutation',
        mutations: [{ type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'After' }],
      },
    );

    const profile = createPptHostWritebackProfile('microsoft-powerpoint');
    const validation = validatePptWritebackPlan(plan, profile);

    expect(validation.ok).toBe(false);
    expect(validation.unsupportedReasons).toEqual([
      'operation replayMutations requires canReplayMutations',
    ]);
    expect(validation.warnings).toEqual([
      'host profile microsoft-powerpoint is unproven; adapter must run probes before writeback',
    ]);
  });

  test('validates replay mutation allow-lists and thresholds', () => {
    const plan = planPptWriteback(
      BASE_MODEL,
      changedModel(),
      {
        canReplayMutations: true,
      },
      {
        preference: 'mutation',
        mutations: [{ type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'After' }],
      },
    );

    expect(
      validatePptWritebackPlan(plan, {
        canReplayMutations: true,
        supportedMutationTypes: ['shape.add'],
        maxMutationCount: 0,
      }).unsupportedReasons,
    ).toEqual([
      'operation replayMutations has 1 mutations; maxMutationCount is 0',
      'operation replayMutations includes unsupported mutation types: shape.text.set',
    ]);
  });

  test('summarizes probe reports into conservative capabilities', () => {
    const report: PptHostProbeReport = {
      host: 'wps-presentation',
      adapterVersion: 'test-adapter',
      checks: [
        { name: 'save.freshness', passed: true },
        { name: 'package.replace', passed: false, message: 'SaveCopyAs returned stale bytes' },
        { name: 'slide.replace', passed: true },
        { name: 'slide.insert', passed: true },
        { name: 'shape.add', passed: true },
        { name: 'text.write', passed: true },
        { name: 'table.structure.write', passed: false, message: 'Merged cells lost styling' },
      ],
      maxMutationCount: 10,
      notes: ['Local smoke probe.'],
    };

    expect(summarizePptHostProbeReport(report)).toEqual({
      host: 'wps-presentation',
      passed: ['save.freshness', 'slide.replace', 'slide.insert', 'shape.add', 'text.write'],
      failed: ['package.replace', 'table.structure.write'],
      capabilities: {
        canReplaceSlides: true,
        canInsertSlides: true,
        canReplayMutations: true,
        supportedMutationTypes: ['shape.add', 'shape.text.set', 'shape.text.replace'],
        maxMutationCount: 10,
      },
      notes: [
        'Local smoke probe.',
        'package.replace: SaveCopyAs returned stale bytes',
        'table.structure.write: Merged cells lost styling',
      ],
    });
  });

  test('does not expose writeback capabilities without save freshness proof', () => {
    const report: PptHostProbeReport = {
      host: 'microsoft-powerpoint',
      checks: [
        { name: 'package.replace', passed: true },
        { name: 'slide.replace', passed: true },
        { name: 'text.write', passed: true },
      ],
    };

    expect(summarizePptHostProbeReport(report).capabilities).toEqual({});
  });

  test('creates probed profiles from report-backed mutation capabilities', () => {
    const profile = createPptHostWritebackProfileFromProbeReport({
      host: 'onlyoffice-presentation',
      checks: [
        { name: 'save.freshness', passed: true },
        { name: 'text.write', passed: true },
        { name: 'text.style.write', passed: true },
        { name: 'image.write', passed: true },
        { name: 'table.cell.write', passed: true },
        { name: 'table.cell.style.write', passed: true },
        { name: 'table.values.write', passed: true },
        { name: 'table.structure.write', passed: true },
      ],
      maxMutationCount: 20,
    });

    expect(profile).toMatchObject({
      host: 'onlyoffice-presentation',
      status: 'probed',
      capabilities: {
        canReplayMutations: true,
        supportedMutationTypes: [
          'shape.text.set',
          'shape.text.replace',
          'shape.text.style.set',
          'shape.image.replace',
          'table.cell.text.set',
          'table.cell.style.set',
          'table.values.set',
          'table.row.add',
          'table.row.delete',
          'table.column.add',
          'table.column.delete',
        ],
        maxMutationCount: 20,
      },
    });

    const plan = planPptWriteback(BASE_MODEL, changedModel(), profile.capabilities, {
      preference: 'mutation',
      mutations: [{ type: 'shape.text.set', slideId: '256', shapeId: '2', text: 'After' }],
    });

    expect(validatePptWritebackPlan(plan, profile).ok).toBe(true);
  });

  test('provides default probe checklists per host kind', () => {
    expect(getDefaultPptHostProbeChecks('desktop-package')).toEqual([
      'package.replace',
      'save.freshness',
    ]);
    expect(getDefaultPptHostProbeChecks('microsoft-powerpoint')).toEqual([
      'slide.replace',
      'slide.insert',
      'slide.delete',
      'slide.reorder',
      'shape.add',
      'shape.delete',
      'shape.name.write',
      'shape.zOrder.write',
      'shape.geometry.write',
      'shape.style.write',
      'text.write',
      'text.style.write',
      'image.write',
      'table.cell.write',
      'table.cell.style.write',
      'table.values.write',
      'table.structure.write',
      'save.freshness',
    ]);
  });

  test('runs probe adapters into plain JSON reports', async () => {
    const report = await runPptHostProbe(
      {
        host: 'microsoft-powerpoint',
        adapterVersion: 'probe-test',
        async runCheck(check) {
          if (check === 'text.write') {
            return { passed: true, details: { marker: 'COCO_PROBE' } };
          }
          if (check === 'image.write') return false;
          throw new Error(`not implemented: ${check}`);
        },
      },
      {
        generatedAt: '2026-08-25T00:00:00.000Z',
        checks: ['text.write', 'image.write', 'save.freshness'],
        maxMutationCount: 5,
        notes: ['Adapter smoke.'],
      },
    );

    expect(report).toEqual({
      host: 'microsoft-powerpoint',
      generatedAt: '2026-08-25T00:00:00.000Z',
      adapterVersion: 'probe-test',
      checks: [
        { name: 'text.write', passed: true, details: { marker: 'COCO_PROBE' } },
        { name: 'image.write', passed: false },
        { name: 'save.freshness', passed: false, message: 'not implemented: save.freshness' },
      ],
      maxMutationCount: 5,
      notes: ['Adapter smoke.'],
    });
  });
});
