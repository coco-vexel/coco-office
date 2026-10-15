import type { ExcelWorkbookCore } from './excel-workbook.js';

export async function executeExcelScript(
  workbook: ExcelWorkbookCore,
  code: string,
  args: Record<string, unknown> | undefined,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<unknown> {
  if (signal?.aborted) {
    throw new Error('excel script aborted before dispatch');
  }

  const runner = new Function(
    'workbook',
    'args',
    `"use strict"; return (async () => {\n${code}\n})()`,
  ) as (workbook: ExcelWorkbookCore, args: Record<string, unknown> | undefined) => Promise<unknown>;

  const execution = runner(workbook, args);
  const races: Promise<unknown>[] = [execution];
  let timeoutHandle: ReturnType<typeof setTimeout> | undefined;
  let abortHandler: (() => void) | undefined;

  if (timeoutMs > 0) {
    races.push(
      new Promise((_, reject) => {
        timeoutHandle = setTimeout(
          () => reject(new Error(`excel script timeout after ${timeoutMs}ms`)),
          timeoutMs,
        );
      }),
    );
  }

  if (signal) {
    races.push(
      new Promise((_, reject) => {
        abortHandler = () => reject(new Error('excel script aborted'));
        signal.addEventListener('abort', abortHandler, { once: true });
      }),
    );
  }

  try {
    return await Promise.race(races);
  } finally {
    if (timeoutHandle) clearTimeout(timeoutHandle);
    if (signal && abortHandler) signal.removeEventListener('abort', abortHandler);
  }
}
