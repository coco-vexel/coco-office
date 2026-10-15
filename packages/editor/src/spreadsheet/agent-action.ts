import type {
  ExcelAgentTransactionGroupOptions,
  ExcelHistoryApplyResult,
} from '@coco-vexel/coco-office/excel';
import type { SpreadsheetRange } from './types.js';
import type {
  SpreadsheetAgentTransactionPreview,
  SpreadsheetEditorRuntime,
} from './runtime.js';

export interface SpreadsheetAgentActionContext {
  /** Current selection at the time the action is created. Hosts may pass richer app state later. */
  selection?: SpreadsheetRange;
  /** Optional opaque host payload, for example an agent plan id or UI prompt metadata. */
  metadata?: Record<string, unknown>;
}

export interface SpreadsheetAgentAction {
  id: string;
  label: string;
  description?: string;
  actor?: string;
  agentRunId?: string;
  agentStepId?: string;
  agentActionId?: string;
  metadata?: Record<string, unknown>;
  run(runtime: SpreadsheetEditorRuntime, context: SpreadsheetAgentActionContext): Promise<void>;
}

export interface SpreadsheetAgentActionPreview {
  action: SpreadsheetAgentAction;
  context: SpreadsheetAgentActionContext;
  transaction: SpreadsheetAgentTransactionPreview;
}

export interface SpreadsheetAgentActionApplyResult {
  action: SpreadsheetAgentAction;
  context: SpreadsheetAgentActionContext;
  transaction: ExcelHistoryApplyResult | null;
}

export async function previewSpreadsheetAgentAction(
  runtime: SpreadsheetEditorRuntime,
  action: SpreadsheetAgentAction,
  context: SpreadsheetAgentActionContext = {},
): Promise<SpreadsheetAgentActionPreview> {
  return {
    action,
    context,
    transaction: await runtime.previewAgentTransaction(agentActionTransactionOptions(action), (targetRuntime) =>
      action.run(targetRuntime, context),
    ),
  };
}

export async function applySpreadsheetAgentAction(
  runtime: SpreadsheetEditorRuntime,
  action: SpreadsheetAgentAction,
  context: SpreadsheetAgentActionContext = {},
): Promise<SpreadsheetAgentActionApplyResult> {
  return {
    action,
    context,
    transaction: await runtime.runAgentTransaction(agentActionTransactionOptions(action), (targetRuntime) =>
      action.run(targetRuntime, context),
    ),
  };
}

function agentActionTransactionOptions(action: SpreadsheetAgentAction): ExcelAgentTransactionGroupOptions {
  return {
    label: action.label,
    actor: action.actor,
    description: action.description,
    agentRunId: action.agentRunId,
    agentStepId: action.agentStepId,
    agentActionId: action.agentActionId ?? action.id,
    metadata: action.metadata,
  };
}
