/**
 * Transport —— WordDocument(主线程代理)与 Kernel 之间的分发抽象。
 *
 * 两种实现:
 *   - `InlineTransport`:kernel 跑在主线程,call 直接委托到 kernelHandlers
 *   - `WorkerTransport`(worker-transport.ts):kernel 跑在 Worker,call 通过 postMessage
 *     编码成 RPC 消息,等待 response
 *
 * Transport 对外形态相同 → WordDocument 不需感知模式。
 */
import type { KernelChangeEvent } from './change-events.js';
import type { KernelMethod } from './kernel-handlers.js';
import { kernelHandlers } from './kernel-handlers.js';
import { Kernel, type KernelProgress } from './kernel.js';
import { executeScript } from './script-execute.js';

export interface Transport {
  /** 调用 kernel 方法;onProgress 用于流式进度(装载等) */
  call<T>(method: KernelMethod, args: unknown[], onProgress?: KernelProgress): Promise<T>;
  /**
   * 在 kernel 同上下文执行一段用户脚本(04.2 §六)。与 `call` 平级,**不**进
   * `KernelMethod` 联合 —— `KernelMethod` 表"数据 API",脚本执行是另一类能力。
   *
   * @param code 用户 JS 源码;函数体形态(允许顶层 await / return)
   * @param args 传给脚本的参数,经 structured clone 过 worker 边界
   * @param timeoutMs 超时毫秒数;0 = 不超时
   * @param signal 可选 AbortSignal(#151):
   *               - WorkerTransport:abort 触发 → terminate worker + reload 路径
   *                 (复用 #144 timeout reload 实施;pending RPC reject 为 \`WORKER_RELOADED\`)
   *               - InlineTransport:abort 触发 → race 拒绝(主线程立即返回,但脚本
   *                 仍在主线程后台跑直到自然结束 —— inline 无 worker 可 terminate)
   * @returns 脚本顶层 return 的值(structured-cloneable)
   * @throws  语法错误 / 脚本运行错 / 超时 / abort
   */
  executeScript(
    code: string,
    args: Record<string, unknown> | undefined,
    timeoutMs: number,
    signal?: AbortSignal,
  ): Promise<unknown>;
  /** 释放底层资源(worker 模式下 terminate worker) */
  dispose(): Promise<void>;
  /**
   * 可选:订阅 reload 完成事件。Worker 模式下每次 crash + reload 成功后触发;
   * InlineTransport 不触发(inline 无 reload 语义)。返回取消订阅函数。
   *
   * 主要用途(#23):SnapshotApi 订阅此事件,在 worker reload 完成后重建内存态
   * (nextId / changelog)——worker reload 会丢失主线程侧的自增栈,需要从 adapter
   * 重读已持久化快照重建 nextId。详见 docs/03.2 §9.1。
   *
   * Listener 可返回 Promise;transport 会串行 await 每个 listener 完成后再 flush
   * reload 期间排队的 RPC,避免 queued 的 take() 在 nextId 重建前执行。
   */
  onReload?(listener: () => void | Promise<void>): () => void;
  /**
   * 订阅 kernel 推出的 change 事件(applyPlan 成功后触发,涵盖 plan / script 两条入口)。
   * 返回取消订阅函数。
   *
   * - InlineTransport:直接转发 kernel.onChange
   * - WorkerTransport:监听 RpcEvent('word.change'),fan-out 给本地订阅者
   *
   * payload 是 `KernelChangeEvent`(精简版,不含全文档);WordDocument 上层根据订阅
   * 选项再决定是否拉取 toFlatOpc / toZip。监听器同步触发,抛错被吞掉。
   */
  onChange(listener: (ev: KernelChangeEvent) => void): () => void;
}

/** Inline:kernel 直接在主线程 */
export class InlineTransport implements Transport {
  readonly kernel: Kernel;

  constructor() {
    this.kernel = new Kernel();
  }

  async call<T>(method: KernelMethod, args: unknown[], onProgress?: KernelProgress): Promise<T> {
    const handler = kernelHandlers[method];
    const result = await handler({ kernel: this.kernel, onProgress }, args);
    return result as T;
  }

  async executeScript(
    code: string,
    args: Record<string, unknown> | undefined,
    timeoutMs: number,
    signal?: AbortSignal,
  ): Promise<unknown> {
    return executeScript(this.kernel, code, args, timeoutMs, signal);
  }

  async dispose(): Promise<void> {
    this.kernel.dispose();
  }

  onChange(listener: (ev: KernelChangeEvent) => void): () => void {
    return this.kernel.onChange(listener);
  }
}
