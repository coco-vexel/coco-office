import type { PackageIndices } from './indices.js';
import type {
  ApplyPlanParams,
  Op,
  PlanValidationIssue,
  PlanValidationItem,
  PlanValidationResult,
} from './plan.js';
import { getParagraphRuns, getRunText } from './runs.js';
import type { WordErrorCode } from './word-error.js';

interface PlanValidationContext {
  indices: PackageIndices;
  options: ApplyPlanParams['options'];
}

type OpValidator = (
  op: Op,
  ctx: PlanValidationContext,
) => PlanValidationIssue | undefined;

const validators: Partial<Record<Op['type'], OpValidator>> = {
  addComment: validateAddComment,
};

/**
 * Validate supported operations against the current immutable document state.
 * Unsupported operation types are reported as `unvalidated`, never assumed valid.
 */
export function validatePlanOps(
  params: ApplyPlanParams,
  indices: PackageIndices,
): PlanValidationResult {
  const ctx: PlanValidationContext = { indices, options: params.options };
  const items: PlanValidationItem[] = params.ops.map((op, opIndex) => {
    const validator = validators[op.type];
    if (!validator) return { opIndex, opType: op.type, status: 'unvalidated' };
    const issue = validator(op, ctx);
    return issue
      ? { opIndex, opType: op.type, status: 'invalid', issue }
      : { opIndex, opType: op.type, status: 'valid' };
  });
  return {
    valid: items.every((item) => item.status === 'valid'),
    items,
  };
}

function validateAddComment(op: Op, ctx: PlanValidationContext): PlanValidationIssue | undefined {
  if (op.type !== 'addComment') return undefined;
  if (ctx.options.trackChanges) {
    return issue(
      'UNSUPPORTED_OOXML_FEATURE',
      'addComment with trackChanges=true is not supported',
      { opType: op.type },
    );
  }
  if ('text' in op.range && op.range.scope === 'all') {
    return issue(
      'UNSUPPORTED_OOXML_FEATURE',
      "addComment: scope='all' is not supported for a text range",
      { opType: op.type, scope: 'all' },
    );
  }

  const paraId = op.range.paraId;
  const paragraph = ctx.indices.paraIdx.get(paraId.toUpperCase());
  if (!paragraph) {
    return issue('PARA_NOT_FOUND', `addComment: paragraph '${paraId}' not found`, { paraId });
  }
  if ('text' in op.range) {
    const search = op.range.text;
    const paragraphText = getParagraphRuns(paragraph).map((run) => getRunText(run)).join('');
    if (!search || !paragraphText.includes(search)) {
      return issue(
        'TEXT_NOT_FOUND_IN_PARA',
        `addComment: text '${search}' not found in paragraph ${paraId}`,
        { paraId, search },
      );
    }
  }
  return undefined;
}

function issue(
  code: WordErrorCode,
  message: string,
  details?: Record<string, unknown>,
): PlanValidationIssue {
  return { code, message, ...(details ? { details } : {}) };
}
