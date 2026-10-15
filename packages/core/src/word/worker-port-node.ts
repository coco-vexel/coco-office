/**
 * Node worker_threads 的 WorkerPortLike 工厂。
 *
 * 把 `new Worker(bundleUrl)` 的 Node Worker 包成 `WorkerPortLike`,交给 WorkerTransport 使用。
 *
 * 仅在 Node 环境可用。浏览器 Worker 走 `worker-port-browser.ts`(待加)。
 */
import type { WorkerPortLike } from './worker-transport.js';

/**
 * 根据给定 bundle URL 启动 Node Worker,返回 WorkerPortLike。
 *
 * `bundleUrl` 通常是 `packages/core/dist/kernel-worker.mjs`(ESM)或 `.cjs`(CJS);
 * 具体由消费方决定并通过 `WordDocument.setWorkerBundleUrl` 注入。
 */
export async function createNodeWorkerPort(bundleUrl: string | URL): Promise<WorkerPortLike> {
  const { Worker: NodeWorker } = await import('node:worker_threads');
  const worker = new NodeWorker(bundleUrl);
  return {
    postMessage(data, _transferList) {
      worker.postMessage(data);
    },
    onMessage(listener) {
      worker.on('message', listener);
    },
    onError(listener) {
      worker.on('error', listener);
      worker.on('exit', (code) => {
        if (code !== 0) listener(new Error(`worker exited with code ${code}`));
      });
    },
    async terminate() {
      await worker.terminate();
    },
  };
}
