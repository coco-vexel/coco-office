/**
 * kernel-worker-browser-entry —— Worker 的启动 bootstrap(浏览器 `Worker` 场景)。
 *
 * tsup 单独打成 `dist/kernel-worker-browser.{js,cjs}`。消费方在主线程用 `new Worker(url,
 * { type: 'module' })` 加载,通过 `WordDocument.setWorkerFactory()` 注入到 SDK。
 *
 * 与 Node entry 唯一区别在 IO 层:`self.onmessage` / `self.postMessage`。dispatch 逻辑、
 * Kernel 单例、错误序列化全部走 `kernel-worker-shared.ts`。
 *
 * 运行时基线:WPS Chromium 104(项目最低支持版本)。不依赖 Chrome 105+ 新 API。
 */
import { createKernelDispatcher } from './kernel-worker-shared.js';

// `self` 在 DedicatedWorkerGlobalScope 里就是 worker 自身;不在 worker 上下文则报错
// (例如有人在主线程误 import 这个 entry)。
const ctx = self as unknown as DedicatedWorkerGlobalScope;
if (typeof ctx.postMessage !== 'function' || typeof ctx.addEventListener !== 'function') {
  throw new Error('kernel-worker-browser-entry must run inside a DedicatedWorkerGlobalScope');
}

const dispatch = createKernelDispatcher((data) => {
  ctx.postMessage(data);
});

ctx.addEventListener('message', (e: MessageEvent): void => {
  void dispatch(e.data);
});
