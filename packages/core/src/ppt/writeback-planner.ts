import type { PptMutation } from './mutations.js';
import type { PptPresentationModel, PptShapeModel, PptSlideModel } from './model.js';

export type PptWritebackStrategy =
  | 'no-op'
  | 'presentation-package-replace'
  | 'slide-package-replace'
  | 'mutation-replay'
  | 'unsupported';

export type PptWritebackPreference = 'auto' | 'package' | 'slide-package' | 'mutation';

export interface PptWritebackCapabilities {
  /** Host can replace the active presentation with a full pptx package. */
  canReplacePresentationPackage?: boolean;
  /** Host can replace an existing slide from a generated slide/package payload. */
  canReplaceSlides?: boolean;
  /** Host can insert new slides from a generated slide/package payload. */
  canInsertSlides?: boolean;
  /** Host can delete slides. */
  canDeleteSlides?: boolean;
  /** Host can reorder slides. */
  canReorderSlides?: boolean;
  /** Host can replay SDK mutation records through live host APIs. */
  canReplayMutations?: boolean;
  /** Optional allow-list for mutation replay. Missing means every mutation type is accepted by the planner. */
  supportedMutationTypes?: PptMutation['type'][];
  /** Upper bound for replay operation count before planner should prefer package replacement. */
  maxMutationCount?: number;
}

export interface PptWritebackPlanOptions {
  preference?: PptWritebackPreference;
  mutations?: PptMutation[];
}

export type PptDeckDiffKind = 'added' | 'deleted' | 'modified';
export type PptShapeDiffKind = 'added' | 'deleted' | 'modified';

export interface PptShapeDiff {
  kind: PptShapeDiffKind;
  shapeId: string;
  beforeName?: string;
  afterName?: string;
  beforeKind?: PptShapeModel['kind'];
  afterKind?: PptShapeModel['kind'];
  before?: PptShapeModel;
  after?: PptShapeModel;
}

export interface PptSlideDiff {
  kind: PptDeckDiffKind;
  slideId: string;
  beforeIndex?: number;
  afterIndex?: number;
  beforeTitle?: string;
  afterTitle?: string;
  moved: boolean;
  propertiesChanged: boolean;
  shapeOrderChanged: boolean;
  shapeChanges: PptShapeDiff[];
  before?: PptSlideModel;
  after?: PptSlideModel;
}

export interface PptPresentationDiff {
  changed: boolean;
  presentationPropertiesChanged: boolean;
  slideSizeChanged: boolean;
  mediaChanged: boolean;
  slides: PptSlideDiff[];
}

export type PptWritebackOperation =
  | { type: 'replacePresentationPackage' }
  | { type: 'replaceSlidePackage'; slideId: string; slideIndex: number; targetSlide: PptSlideModel }
  | { type: 'insertSlidePackage'; slideId: string; slideIndex: number; targetSlide: PptSlideModel }
  | { type: 'deleteSlide'; slideId: string; slideIndex: number }
  | { type: 'moveSlide'; slideId: string; fromIndex: number; toIndex: number }
  | { type: 'replayMutations'; mutations: PptMutation[] };

export interface PptWritebackPlan {
  strategy: PptWritebackStrategy;
  diff: PptPresentationDiff;
  operations: PptWritebackOperation[];
  warnings: string[];
  unsupportedReasons: string[];
}

export function diffPptPresentationModels(
  before: PptPresentationModel,
  after: PptPresentationModel,
): PptPresentationDiff {
  const beforeById = new Map(before.slides.map((slide) => [slide.id, slide]));
  const afterById = new Map(after.slides.map((slide) => [slide.id, slide]));
  const slides: PptSlideDiff[] = [];

  for (const beforeSlide of before.slides) {
    const afterSlide = afterById.get(beforeSlide.id);
    if (!afterSlide) {
      slides.push({
        kind: 'deleted',
        slideId: beforeSlide.id,
        beforeIndex: beforeSlide.index,
        beforeTitle: beforeSlide.title,
        moved: false,
        propertiesChanged: true,
        shapeOrderChanged: false,
        shapeChanges: beforeSlide.shapes.map((shape) => ({
          kind: 'deleted',
          shapeId: shape.id,
          beforeName: shape.name,
          beforeKind: shape.kind,
          before: cloneModel(shape),
        })),
        before: cloneModel(beforeSlide),
      });
      continue;
    }

    const shapeChanges = diffShapes(beforeSlide.shapes, afterSlide.shapes);
    const moved = beforeSlide.index !== afterSlide.index;
    const propertiesChanged = slidePropertiesKey(beforeSlide) !== slidePropertiesKey(afterSlide);
    const shapeOrderChanged =
      shapeOrderKey(beforeSlide.shapes) !== shapeOrderKey(afterSlide.shapes);
    if (moved || propertiesChanged || shapeOrderChanged || shapeChanges.length > 0) {
      slides.push({
        kind: 'modified',
        slideId: beforeSlide.id,
        beforeIndex: beforeSlide.index,
        afterIndex: afterSlide.index,
        beforeTitle: beforeSlide.title,
        afterTitle: afterSlide.title,
        moved,
        propertiesChanged,
        shapeOrderChanged,
        shapeChanges,
        before: cloneModel(beforeSlide),
        after: cloneModel(afterSlide),
      });
    }
  }

  for (const afterSlide of after.slides) {
    if (!beforeById.has(afterSlide.id)) {
      slides.push({
        kind: 'added',
        slideId: afterSlide.id,
        afterIndex: afterSlide.index,
        afterTitle: afterSlide.title,
        moved: false,
        propertiesChanged: true,
        shapeOrderChanged: false,
        shapeChanges: afterSlide.shapes.map((shape) => ({
          kind: 'added',
          shapeId: shape.id,
          afterName: shape.name,
          afterKind: shape.kind,
          after: cloneModel(shape),
        })),
        after: cloneModel(afterSlide),
      });
    }
  }

  const presentationPropertiesChanged =
    presentationPropertiesKey(before) !== presentationPropertiesKey(after);
  const slideSizeChanged = stableJson(before.slideSize) !== stableJson(after.slideSize);
  const mediaChanged = stableJson(before.media) !== stableJson(after.media);

  return {
    changed: presentationPropertiesChanged || slideSizeChanged || mediaChanged || slides.length > 0,
    presentationPropertiesChanged,
    slideSizeChanged,
    mediaChanged,
    slides: slides.sort(
      (left, right) =>
        (left.afterIndex ?? left.beforeIndex ?? 0) - (right.afterIndex ?? right.beforeIndex ?? 0),
    ),
  };
}

export function planPptWriteback(
  before: PptPresentationModel,
  after: PptPresentationModel,
  capabilities: PptWritebackCapabilities,
  options: PptWritebackPlanOptions = {},
): PptWritebackPlan {
  const diff = diffPptPresentationModels(before, after);
  if (!diff.changed) {
    return { strategy: 'no-op', diff, operations: [], warnings: [], unsupportedReasons: [] };
  }

  const preference = options.preference ?? 'auto';
  const packagePlan = tryPlanPackage(diff, capabilities);
  if (preference === 'package') return packagePlan;

  const slidePlan = tryPlanSlidePackage(diff, after, capabilities);
  if (preference === 'slide-package') return slidePlan;

  const mutationPlan = tryPlanMutationReplay(diff, options.mutations, capabilities);
  if (preference === 'mutation') return mutationPlan;

  if (packagePlan.strategy !== 'unsupported') return packagePlan;
  if (slidePlan.strategy !== 'unsupported') return slidePlan;
  if (mutationPlan.strategy !== 'unsupported') return mutationPlan;

  return unsupported(diff, [
    ...packagePlan.unsupportedReasons,
    ...slidePlan.unsupportedReasons,
    ...mutationPlan.unsupportedReasons,
  ]);
}

function tryPlanPackage(
  diff: PptPresentationDiff,
  capabilities: PptWritebackCapabilities,
): PptWritebackPlan {
  if (!capabilities.canReplacePresentationPackage) {
    return unsupported(diff, ['host cannot replace the active presentation package in place']);
  }
  return {
    strategy: 'presentation-package-replace',
    diff,
    operations: [{ type: 'replacePresentationPackage' }],
    warnings: [],
    unsupportedReasons: [],
  };
}

function tryPlanSlidePackage(
  diff: PptPresentationDiff,
  after: PptPresentationModel,
  capabilities: PptWritebackCapabilities,
): PptWritebackPlan {
  const operations: PptWritebackOperation[] = [];
  const warnings: string[] = [];
  const afterById = new Map(after.slides.map((slide) => [slide.id, slide]));

  if (diff.presentationPropertiesChanged || diff.slideSizeChanged || diff.mediaChanged) {
    return unsupported(diff, [
      'slide package replacement cannot apply presentation-level or media-list changes alone',
    ]);
  }

  for (const slide of diff.slides) {
    if (slide.kind === 'added') {
      if (!capabilities.canInsertSlides)
        return unsupported(diff, ['slide package replacement cannot insert slides']);
      const targetSlide = afterById.get(slide.slideId);
      if (!targetSlide || slide.afterIndex === undefined) {
        return unsupported(diff, [`target slide model not found for slide id ${slide.slideId}`]);
      }
      operations.push({
        type: 'insertSlidePackage',
        slideId: slide.slideId,
        slideIndex: slide.afterIndex,
        targetSlide: cloneModel(targetSlide),
      });
      continue;
    }

    if (slide.kind === 'deleted') {
      if (!capabilities.canDeleteSlides)
        return unsupported(diff, ['slide package replacement cannot delete slides']);
      operations.push({
        type: 'deleteSlide',
        slideId: slide.slideId,
        slideIndex: slide.beforeIndex ?? 0,
      });
      continue;
    }

    if (slide.moved) {
      if (!capabilities.canReorderSlides)
        return unsupported(diff, ['slide package replacement cannot reorder slides']);
      operations.push({
        type: 'moveSlide',
        slideId: slide.slideId,
        fromIndex: slide.beforeIndex ?? 0,
        toIndex: slide.afterIndex ?? 0,
      });
    }

    if (slide.propertiesChanged || slide.shapeOrderChanged || slide.shapeChanges.length > 0) {
      if (!capabilities.canReplaceSlides)
        return unsupported(diff, ['host cannot replace individual slide packages']);
      const targetSlide = afterById.get(slide.slideId);
      if (!targetSlide || slide.afterIndex === undefined) {
        return unsupported(diff, [`target slide model not found for slide id ${slide.slideId}`]);
      }
      operations.push({
        type: 'replaceSlidePackage',
        slideId: slide.slideId,
        slideIndex: slide.afterIndex,
        targetSlide: cloneModel(targetSlide),
      });
    }
  }

  if (operations.some((operation) => operation.type === 'replaceSlidePackage')) {
    warnings.push(
      'slide package replacement is adapter-defined and must preserve relationships, layouts, and media parts',
    );
  }

  return { strategy: 'slide-package-replace', diff, operations, warnings, unsupportedReasons: [] };
}

function tryPlanMutationReplay(
  diff: PptPresentationDiff,
  mutations: PptMutation[] | undefined,
  capabilities: PptWritebackCapabilities,
): PptWritebackPlan {
  if (!capabilities.canReplayMutations) {
    return unsupported(diff, ['host cannot replay PPT SDK mutations']);
  }
  if (!mutations || mutations.length === 0) {
    return unsupported(diff, ['mutation replay requires SDK mutation records']);
  }
  const maxMutationCount = capabilities.maxMutationCount ?? Number.POSITIVE_INFINITY;
  if (mutations.length > maxMutationCount) {
    return unsupported(diff, [
      `mutation count ${mutations.length} exceeds maxMutationCount ${maxMutationCount}`,
    ]);
  }
  const supported = capabilities.supportedMutationTypes
    ? new Set(capabilities.supportedMutationTypes)
    : null;
  const unsupportedTypes = supported
    ? mutations.map((mutation) => mutation.type).filter((type) => !supported.has(type))
    : [];
  if (unsupportedTypes.length > 0) {
    return unsupported(diff, [
      `host cannot replay mutation types: ${[...new Set(unsupportedTypes)].join(', ')}`,
    ]);
  }

  const warnings: string[] = [];
  if (diff.presentationPropertiesChanged || diff.slideSizeChanged) {
    warnings.push(
      'mutation replay must validate presentation-level changes before applying host operations',
    );
  }
  if (diff.mediaChanged) {
    warnings.push('mutation replay must validate media relationships and package image parts');
  }

  return {
    strategy: 'mutation-replay',
    diff,
    operations: [{ type: 'replayMutations', mutations: mutations.map(cloneModel) }],
    warnings,
    unsupportedReasons: [],
  };
}

function unsupported(diff: PptPresentationDiff, unsupportedReasons: string[]): PptWritebackPlan {
  return {
    strategy: 'unsupported',
    diff,
    operations: [],
    warnings: [],
    unsupportedReasons: [...new Set(unsupportedReasons)],
  };
}

function diffShapes(before: PptShapeModel[], after: PptShapeModel[]): PptShapeDiff[] {
  const beforeById = new Map(before.map((shape) => [shape.id, shape]));
  const afterById = new Map(after.map((shape) => [shape.id, shape]));
  const changes: PptShapeDiff[] = [];

  for (const beforeShape of before) {
    const afterShape = afterById.get(beforeShape.id);
    if (!afterShape) {
      changes.push({
        kind: 'deleted',
        shapeId: beforeShape.id,
        beforeName: beforeShape.name,
        beforeKind: beforeShape.kind,
        before: cloneModel(beforeShape),
      });
    } else if (stableJson(beforeShape) !== stableJson(afterShape)) {
      changes.push({
        kind: 'modified',
        shapeId: beforeShape.id,
        beforeName: beforeShape.name,
        afterName: afterShape.name,
        beforeKind: beforeShape.kind,
        afterKind: afterShape.kind,
        before: cloneModel(beforeShape),
        after: cloneModel(afterShape),
      });
    }
  }

  for (const afterShape of after) {
    if (!beforeById.has(afterShape.id)) {
      changes.push({
        kind: 'added',
        shapeId: afterShape.id,
        afterName: afterShape.name,
        afterKind: afterShape.kind,
        after: cloneModel(afterShape),
      });
    }
  }

  return changes.sort((left, right) =>
    left.shapeId.localeCompare(right.shapeId, undefined, { numeric: true }),
  );
}

function presentationPropertiesKey(model: PptPresentationModel): string {
  return stableJson({
    id: model.id,
    title: model.title,
  });
}

function slidePropertiesKey(slide: PptSlideModel): string {
  return stableJson({
    id: slide.id,
    name: slide.name,
    title: slide.title,
    relationshipId: slide.relationshipId,
    path: slide.path,
    layoutId: slide.layoutId,
  });
}

function shapeOrderKey(shapes: PptShapeModel[]): string {
  return stableJson(shapes.map((shape) => shape.id));
}

function stableJson(value: unknown): string {
  return JSON.stringify(sortForJson(value));
}

function sortForJson(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortForJson);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, entryValue]) => entryValue !== undefined)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, entryValue]) => [key, sortForJson(entryValue)]),
  );
}

function cloneModel<T>(model: T): T {
  return JSON.parse(JSON.stringify(model)) as T;
}
