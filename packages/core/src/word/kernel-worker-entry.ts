/**
 * kernel-worker-entry —— Worker 的启动 bootstrap(Node `worker_threads` 场景)。
 *
 * tsup 把本文件单独打成 `dist/kernel-worker.{js,cjs}`,主线程用 `new Worker(bundleUrl)`
 * 加载。本 entry 仅负责 IO 层桥接:
 *   - `parentPort.on('message')` → 喂给 dispatcher
 *   - dispatcher 产出的 RpcMessage → `parentPort.postMessage`
 *
 * 实际 dispatch / Kernel 单例 / 错误序列化 都在 `kernel-worker-shared.ts`,与浏览器
 * entry(`kernel-worker-browser-entry.ts`)共享。
 */
import { parentPort } from 'node:worker_threads';
import { createKernelDispatcher } from './kernel-worker-shared.js';

if (!parentPort) {
  throw new Error(
    'kernel-worker-entry must run inside a Node worker_threads Worker (parentPort missing)',
  );
}

const port = parentPort;
const dispatch = createKernelDispatcher((data) => {
  port.postMessage(data);
});

port.on('message', (raw: unknown): void => {
  void dispatch(raw);
});
