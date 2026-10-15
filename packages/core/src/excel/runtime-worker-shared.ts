import { ExcelWorkbookCore } from './excel-workbook.js';
import type { ExcelWorkbookModel } from './model.js';
import {
  serializeExcelError,
  type ExcelRpcError,
  type ExcelRpcMessage,
  type ExcelRpcRequest,
  type ExcelRpcResponse,
} from './runtime-rpc.js';
import { executeExcelScript } from './runtime-script.js';

export type ExcelWorkerPost = (data: unknown) => void;

export function createExcelRuntimeDispatcher(post: ExcelWorkerPost): (raw: unknown) => Promise<void> {
  let workbook: ExcelWorkbookCore | null = null;

  return async function dispatch(raw: unknown): Promise<void> {
    if (raw === null || typeof raw !== 'object') return;
    const msg = raw as ExcelRpcMessage;
    if (msg.kind !== 'request') return;
    const req = msg as ExcelRpcRequest;

    try {
      let result: unknown;
      switch (req.method) {
        case 'load.zip': {
          const [bytes, opts] = req.args as [Uint8Array | ArrayBuffer, { workbookId?: string }];
          workbook = await ExcelWorkbookCore.fromZip(bytes, opts);
          result = {
            workbookId: workbook.workbookId,
          };
          break;
        }
        case 'load.model': {
          const [model, opts] = req.args as [ExcelWorkbookModel, { workbookId?: string }];
          workbook = ExcelWorkbookCore.fromModel(model, opts);
          result = {
            workbookId: workbook.workbookId,
          };
          break;
        }
        case 'to.model': {
          assertWorkbook(workbook);
          result = workbook.toModel();
          break;
        }
        case 'mutations.list': {
          assertWorkbook(workbook);
          result = workbook.getMutations();
          break;
        }
        case 'mutations.clear': {
          assertWorkbook(workbook);
          workbook.clearMutations();
          result = undefined;
          break;
        }
        case 'to.zip': {
          assertWorkbook(workbook);
          result = await workbook.toZip();
          break;
        }
        case 'script.execute': {
          assertWorkbook(workbook);
          const [code, args, timeoutMs] = req.args as [
            string,
            Record<string, unknown> | undefined,
            number,
          ];
          result = await executeExcelScript(workbook, code, args, timeoutMs);
          break;
        }
        case 'dispose': {
          workbook = null;
          result = undefined;
          break;
        }
        default:
          throw new Error(`unknown excel runtime method: ${req.method}`);
      }
      post({ kind: 'response', id: req.id, result } satisfies ExcelRpcResponse);
    } catch (error) {
      post({ kind: 'error', id: req.id, error: serializeExcelError(error) } satisfies ExcelRpcError);
    }
  };
}

function assertWorkbook(workbook: ExcelWorkbookCore | null): asserts workbook is ExcelWorkbookCore {
  if (!workbook) throw new Error('ExcelRuntimeWorkbook is not loaded');
}
