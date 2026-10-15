import { createExcelRuntimeDispatcher } from './runtime-worker-shared.js';

const ctx = self as unknown as DedicatedWorkerGlobalScope;
if (typeof ctx.postMessage !== 'function' || typeof ctx.addEventListener !== 'function') {
  throw new Error('excel runtime browser worker entry must run inside a DedicatedWorkerGlobalScope');
}

const dispatch = createExcelRuntimeDispatcher((data) => {
  ctx.postMessage(data);
});

ctx.addEventListener('message', (event: MessageEvent): void => {
  void dispatch(event.data);
});
