export interface ExcelRpcRequest {
  kind: 'request';
  id: string;
  method: string;
  args: unknown[];
}

export interface ExcelRpcResponse {
  kind: 'response';
  id: string;
  result: unknown;
}

export interface ExcelRpcError {
  kind: 'error';
  id: string;
  error: ExcelSerializableError;
}

export interface ExcelSerializableError {
  message: string;
  name?: string;
  stack?: string;
}

export type ExcelRpcMessage = ExcelRpcRequest | ExcelRpcResponse | ExcelRpcError;

export function serializeExcelError(error: unknown): ExcelSerializableError {
  const err = error as Error | null | undefined;
  return {
    message: err?.message ?? String(error),
    name: err?.name,
    stack: err?.stack,
  };
}

export function deserializeExcelError(error: ExcelSerializableError): Error {
  const err = new Error(error.message);
  err.name = error.name ?? 'Error';
  if (error.stack) err.stack = error.stack;
  return err;
}
