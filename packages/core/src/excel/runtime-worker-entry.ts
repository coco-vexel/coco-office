import { parentPort } from 'node:worker_threads';
import { createExcelRuntimeDispatcher } from './runtime-worker-shared.js';

if (!parentPort) {
  throw new Error('excel runtime worker entry must run inside a Node worker_threads Worker');
}

const port = parentPort;
const dispatch = createExcelRuntimeDispatcher((data) => {
  port.postMessage(data);
});

port.on('message', (raw: unknown): void => {
  void dispatch(raw);
});
