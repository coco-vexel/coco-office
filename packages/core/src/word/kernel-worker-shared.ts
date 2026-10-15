/**
 * kernel-worker-shared —— Worker bootstrap 的环境无关核心。
 *
 * Node `worker_threads` 与浏览器 `Worker` 的差异只在 IO 层(`parentPort.on('message')`
 * vs `self.onmessage`,`parentPort.postMessage` vs `self.postMessage`)。本模块把 RPC
 * dispatch 抽出来,Node entry / 浏览器 entry 都注入一个 `post` 函数即可复用。
 *
 * 不在 public API 暴露,仅供 `kernel-worker-entry.ts`(Node)/`kernel-worker-browser-entry.ts`
 * 引用。
 */
import { WORD_CHANGE_EVENT } from './change-events.js';
import { Kernel } from './kernel.js';
import { kernelHandlers, type KernelMethod } from './kernel-handlers.js';
import type {
  RpcError,
  RpcEvent,
  RpcMessage,
  RpcProgress,
  RpcRequest,
  RpcResponse,
  SerializableError,
} from './rpc-protocol.js';
import { executeScript } from './script-execute.js';
import { WordError } from './word-error.js';

/** 独立 RPC method —— 不进 KernelMethod 联合(04.2 §六)。 */
const SCRIPT_EXECUTE_METHOD = 'script.execute';

/** 由调用方注入,把消息回投给主线程。 */
export type PostFn = (data: unknown) => void;

/**
 * 创建一个 dispatcher 闭包,内部持单例 Kernel。每次收到 raw 消息调用即可。
 *
 * 调用方负责:
 *   - 用环境对应的事件机制把 raw 消息送进来
 *   - 提供 `post` 把 dispatcher 产出的 response/error/progress 回投给主线程
 */
export function createKernelDispatcher(post: PostFn): (raw: unknown) => Promise<void> {
  const kernel = new Kernel();

  // 把 worker 内的 change 事件 post 到主线程,主线程 WorkerTransport.handleIncoming
  // 在 RpcEvent 分支里 fan-out 给 WordDocument 的订阅者。每次 worker 实例只挂一次,
  // worker 生命周期内不取消(dispatcher 自身没有 dispose 入口;worker terminate 时
  // kernel 引用也释放,listener 不再被调用)。
  kernel.onChange((ev) => {
    post({ kind: 'event', event: WORD_CHANGE_EVENT, payload: ev } satisfies RpcEvent);
  });

  return async function dispatch(raw: unknown): Promise<void> {
    // 防御 host 投来非 RpcMessage 的 garbage(主线程库 bug / 浏览器 worker 兼容性消息等)
    if (raw === null || typeof raw !== 'object') return;
    const msg = raw as RpcMessage;
    if (msg.kind !== 'request') return;
    const req = msg as RpcRequest;

    // script.execute 走独立分支(04.2 §六),不进 kernelHandlers
    if (req.method === SCRIPT_EXECUTE_METHOD) {
      try {
        const [code, args, timeoutMs] = req.args as [
          string,
          Record<string, unknown> | undefined,
          number,
        ];
        const result = await executeScript(kernel, code, args, timeoutMs);
        post({ kind: 'response', id: req.id, result } satisfies RpcResponse);
      } catch (e) {
        post({ kind: 'error', id: req.id, error: serializeError(e) } satisfies RpcError);
      }
      return;
    }

    const handler = kernelHandlers[req.method as KernelMethod];
    if (!handler) {
      post({
        kind: 'error',
        id: req.id,
        error: { message: `unknown kernel method: ${req.method}` },
      } satisfies RpcError);
      return;
    }

    try {
      const result = await handler(
        {
          kernel,
          onProgress: (phase, pct) => {
            post({ kind: 'progress', id: req.id, phase, pct } satisfies RpcProgress);
          },
        },
        req.args,
      );
      post({ kind: 'response', id: req.id, result } satisfies RpcResponse);
    } catch (e) {
      post({ kind: 'error', id: req.id, error: serializeError(e) } satisfies RpcError);
    }
  };
}

export function serializeError(e: unknown): SerializableError {
  if (e instanceof WordError) {
    return {
      code: e.code,
      message: e.message,
      docId: e.docId,
      details: e.details,
      name: e.name,
    };
  }
  const err = e as Error | null | undefined;
  return {
    message: err?.message ?? String(e),
    name: err?.name,
    stack: err?.stack,
  };
}
