/**
 * 浏览器 `Worker` → `WorkerPortLike` 适配器。
 *
 * 用于配合 `WordDocument.setWorkerFactory()`:消费方拿到自己 bundler 解析出的 worker
 * 实例(Vite `?worker`、`new Worker(new URL(...))`、或预构造的 Worker)后,直接传进来:
 *
 * ```ts
 * import KernelWorker from '@coco-vexel/coco-office/kernel-worker-browser?worker';
 * import { WordDocument, createBrowserWorkerPort } from '@coco-vexel/coco-office/word';
 *
 * WordDocument.setWorkerFactory(() => createBrowserWorkerPort(new KernelWorker()));
 * ```
 *
 * 仅在浏览器 / WPS 等具备 `Worker` 全局的环境可用。Node 路径请用 `createNodeWorkerPort`。
 */
import type { WorkerPortLike } from './worker-transport.js';

/**
 * 把已构造好的浏览器 `Worker`(或 WPS Chromium 内的等价实例)包成 `WorkerPortLike`。
 *
 * `terminate()` 同步触发底层 `worker.terminate()`,并 resolve(浏览器 API 是 sync void,
 * 这里包成 Promise 与 Node 路径对齐)。
 */
export function createBrowserWorkerPort(worker: Worker): WorkerPortLike {
  return {
    postMessage(data, transferList) {
      // 浏览器 Worker.postMessage 第二参数要求 Transferable[];这里仅在传了 transferList 时透传
      if (transferList && transferList.length > 0) {
        worker.postMessage(data, transferList as Transferable[]);
      } else {
        worker.postMessage(data);
      }
    },
    onMessage(listener) {
      worker.addEventListener('message', (e: MessageEvent) => {
        listener(e.data);
      });
    },
    onError(listener) {
      worker.addEventListener('error', (e: ErrorEvent) => {
        // ErrorEvent 信息字段化后再上抛,跨边界保留更多上下文
        listener(e.error ?? new Error(e.message ?? 'worker error'));
      });
      worker.addEventListener('messageerror', () => {
        listener(new Error('worker messageerror (structured clone failed)'));
      });
    },
    async terminate() {
      worker.terminate();
    },
  };
}
