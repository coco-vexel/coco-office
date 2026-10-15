/**
 * WorkerTransport —— 主线程侧的 RPC 客户端。kernel 跑在 Worker,本类负责:
 *   - 生成 request id
 *   - postMessage 给 Worker
 *   - 监听 Worker 回投的 response/error/progress,关联到对应 pending request
 *   - 返回 Promise,主线程侧消费者像 InlineTransport 一样调用
 *   - **Worker crash → terminate + reload**(#19,docs/03 §4.7)
 *
 * **支持的 Worker 运行时**:
 *   - Node `worker_threads`(测试路径,#17 本 PR 主要覆盖)
 *   - 浏览器 `Worker`(代码路径已预留,但本 PR 未在 CI 里验证;#18 / 浏览器示例补齐)
 *
 * 错误:Worker 回投的 `SerializableError` 被重建为 `WordError`(如 `code` 存在)或
 * 原生 `Error`(非 WordError 情况)。
 *
 * **Reload 语义(#19)**:
 *   - 未配置 reload:port 崩溃 → 所有 pending 以底层 error reject(旧行为,防御性兜底)
 *   - 配置 reload:port 崩溃 → 旧 pending 全部以 `WordError{WORKER_RELOADED}` reject;
 *     起新 port,重放 load;期间到达的 `call()` 排队,reload 完成后按序 flush
 */
import { WORD_CHANGE_EVENT, type KernelChangeEvent } from './change-events.js';
import type { KernelLoadParams, KernelProgress } from './kernel.js';
import type { KernelMethod } from './kernel-handlers.js';
import {
  isRpcMessage,
  type RpcMessage,
  type RpcRequest,
  type SerializableError,
} from './rpc-protocol.js';
import type { Transport } from './transport.js';
import { WordError, type WordErrorCode } from './word-error.js';

/**
 * 抽象 WorkerPort —— Node worker_threads 的 `Worker` 与浏览器 `Worker` 接口不完全一致,
 * 本接口是公约数。
 */
export interface WorkerPortLike {
  postMessage(data: unknown, transferList?: unknown[]): void;
  /** 注册消息监听 */
  onMessage(listener: (data: unknown) => void): void;
  /** 注册错误监听(Worker 内部 crash;Node 侧 exit!=0 也折进这里) */
  onError(listener: (err: unknown) => void): void;
  /** terminate —— 返回 Promise 便于 await(某些 API 是 sync void,包成 Promise) */
  terminate(): Promise<void>;
}

/** Reload 所需参数(#19);缺省即关闭 reload,onError 退化为 reject all pending */
export interface WorkerReloadConfig {
  /** 起一个新 WorkerPortLike(与原 port 同 bundle、同配置) */
  factory: () => WorkerPortLike | Promise<WorkerPortLike>;
  /** reload 时重放的 load 参数(保留原始字节 / flat-opc 字符串) */
  loadParams: KernelLoadParams;
  /** docId —— WORKER_RELOADED 错误注入用 */
  docId: string;
}

export interface WorkerTransportOptions {
  /** 预先创建好的 WorkerPortLike;创建逻辑由调用方负责(环境差异太大,不在本类里) */
  port: WorkerPortLike;
  /** 可选 reload 配置(#19);未传则 worker 崩溃后不重启 */
  reload?: WorkerReloadConfig;
}

interface PendingCall {
  resolve: (value: unknown) => void;
  reject: (err: unknown) => void;
  onProgress?: KernelProgress;
  /** 原始请求,用于 reload 完成后重新入队已 flush 的"正在执行中"队列不需要;见 queued */
  req: RpcRequest;
}

/** reload 期间 call() 排队的请求;reload 完成后一次性 flush */
interface QueuedCall {
  method: KernelMethod;
  args: unknown[];
  onProgress?: KernelProgress;
  resolve: (value: unknown) => void;
  reject: (err: unknown) => void;
}

type TransportState = 'idle' | 'reloading' | 'disposed';

export class WorkerTransport implements Transport {
  private port: WorkerPortLike;
  private readonly reloadConfig?: WorkerReloadConfig;
  private readonly pending = new Map<string, PendingCall>();
  /** reload 期间新到达的 call() 排队于此,reload 完成后按 FIFO 执行 */
  private readonly queued: QueuedCall[] = [];
  private nextId = 0;
  private state: TransportState = 'idle';
  /** reload 重放 load RPC 的 promise,dispose 时 await 避免竞态 */
  private reloadingPromise?: Promise<void>;
  /**
   * 给 port handler 用的代际序号。每次 reload 把 generation++,旧 port 的 onMessage/onError
   * 回调里捕获了创建时的 gen,gen 与当前不符就静默丢弃 —— 避免旧 port 的延迟 exit 事件
   * 错误地触发第二次 reload。
   */
  private portGeneration = 0;
  /**
   * reload 完成后触发的监听器列表(#23)。SnapshotApi 订阅此以在 worker 重建后
   * 重新初始化内存栈(nextId / changelog)。每次 reload 成功后、flush queued 前
   * **串行 await** 所有监听器,避免 queued 的 take() 拿到陈旧 nextId 与 adapter
   * 里已持久化的 id 撞车。
   */
  private readonly reloadListeners = new Set<() => void | Promise<void>>();
  /**
   * change 事件订阅者(WordDocument.on('change', ...) 的底层挂载点)。worker 内
   * 每次 applyPlan 成功会 post 一条 RpcEvent('word.change'),本类在 handleIncoming
   * 里 fan-out 给这里的所有 listener。
   */
  private readonly changeListeners = new Set<(ev: KernelChangeEvent) => void>();

  constructor(opts: WorkerTransportOptions) {
    this.port = opts.port;
    this.reloadConfig = opts.reload;
    this.attachPortHandlers(this.port, this.portGeneration);
  }

  async call<T>(method: KernelMethod, args: unknown[], onProgress?: KernelProgress): Promise<T> {
    if (this.state === 'disposed') {
      throw new Error('WorkerTransport has been disposed');
    }
    // reload 期间入队;完成后按序 flush
    if (this.state === 'reloading') {
      return new Promise<T>((resolve, reject) => {
        this.queued.push({
          method,
          args,
          onProgress,
          resolve: (v) => resolve(v as T),
          reject,
        });
      });
    }
    return this.sendCall<T>(method, args, onProgress);
  }

  /**
   * 04.2 §六:走独立 RPC method `'script.execute'`,**不**进 KernelMethod 联合。
   * worker-entry 在收到该 method 时分支到 script-execute,绕过 kernelHandlers。
   *
   * **超时语义(04.2 §四,#144):**
   *
   * - 主线程持有**硬超时**:`timeoutMs > 0` 时 setTimeout;超时未 resolve → 调
   *   `handleWorkerError` 主动触发 reload 路径。所有 pending 以 `WordError(WORKER_RELOADED)`
   *   reject;`worker.terminate()` + 起新 worker + 重放 `load`。
   * - **不**把 `timeoutMs` 传给 worker —— worker 内 `script-execute.ts` 用 `0` 表示不 race,
   *   完全交主线程 terminate(避免 worker 内对死循环类同步阻塞无能为力)。
   * - `timeoutMs === 0` → 主线程不设硬超时(可能 hang,由调用方风险自负)。
   *
   * **Abort 语义(#151):**
   *
   * - 接 \`AbortSignal\`;abort 触发与 timeout 走**同一条路径**(handleWorkerError →
   *   reload)。pending RPC 以 \`WordError(WORKER_RELOADED)\` reject;\`worker.terminate()\`
   *   立即终止正在运行的脚本(包括同步死循环类阻塞)。
   * - 上层(plugin / agent)负责把 abort cause 转为 \`AgentAbortError\`(L2),不影响
   *   transport 层契约:transport 只管 reload,语义解读由 agent。
   */
  async executeScript(
    code: string,
    args: Record<string, unknown> | undefined,
    timeoutMs: number,
    signal?: AbortSignal,
  ): Promise<unknown> {
    if (this.state === 'disposed') {
      throw new Error('WorkerTransport has been disposed');
    }
    if (signal?.aborted) {
      throw new Error('script aborted before dispatch');
    }
    if (this.state === 'reloading') {
      return new Promise<unknown>((resolve, reject) => {
        this.queued.push({
          method: 'script.execute' as KernelMethod, // 类型边界:queued 槽复用
          // 注意:queued 在 reload 后 flush 时会重发,届时已经是新 worker。仍传 0
          // 让 worker 不做软 race,主线程超时由 flush 后的 executeScript 包装层接管。
          args: [code, args, 0],
          resolve,
          reject,
        });
      });
    }

    // 启动 RPC,主线程不让 worker 做软 race(传 0),由本方法的 setTimeout 兜底
    const callPromise = this.sendCall<unknown>('script.execute' as KernelMethod, [code, args, 0]);

    // 注册 timeout + abort listener;两者都通过 handleWorkerError 走 reload 路径
    const cleanups: (() => void)[] = [];

    if (timeoutMs > 0) {
      const timeoutHandle = setTimeout(() => {
        if (this.state === 'idle') {
          this.handleWorkerError(new Error(`script timeout after ${timeoutMs}ms`));
        }
      }, timeoutMs);
      cleanups.push(() => clearTimeout(timeoutHandle));
    }

    if (signal) {
      const abortHandler = (): void => {
        if (this.state === 'idle') {
          this.handleWorkerError(new Error('script aborted'));
        }
      };
      signal.addEventListener('abort', abortHandler, { once: true });
      cleanups.push(() => signal.removeEventListener('abort', abortHandler));
    }

    try {
      return await callPromise;
    } finally {
      for (const cleanup of cleanups) cleanup();
    }
  }

  async dispose(): Promise<void> {
    const wasState = this.state;
    this.state = 'disposed';
    // reject all pending
    for (const [id, call] of this.pending) {
      call.reject(new Error('WorkerTransport disposed with pending calls'));
      this.pending.delete(id);
    }
    // reject queued
    while (this.queued.length > 0) {
      const q = this.queued.shift()!;
      q.reject(new Error('WorkerTransport disposed with queued calls'));
    }
    // 若 reload 正在进行,等它结束再 terminate(避免 double-terminate 竞态)
    if (wasState === 'reloading' && this.reloadingPromise) {
      try {
        await this.reloadingPromise;
      } catch {
        /* ignore */
      }
    }
    try {
      await this.port.terminate();
    } catch {
      /* 已死的 port terminate 可能抛错,忽略 */
    }
  }

  // ─────────── 内部 ───────────

  private attachPortHandlers(port: WorkerPortLike, generation: number): void {
    port.onMessage((data) => {
      if (generation !== this.portGeneration) return;
      this.handleIncoming(data);
    });
    port.onError((err) => {
      if (generation !== this.portGeneration) return;
      this.handleWorkerError(err);
    });
  }

  private sendCall<T>(
    method: KernelMethod,
    args: unknown[],
    onProgress?: KernelProgress,
  ): Promise<T> {
    const id = `rpc-${this.nextId++}`;
    const req: RpcRequest = { kind: 'request', id, method, args };
    return new Promise<T>((resolve, reject) => {
      this.pending.set(id, {
        resolve: (v) => resolve(v as T),
        reject,
        onProgress,
        req,
      });
      try {
        this.port.postMessage(req);
      } catch (e) {
        this.pending.delete(id);
        reject(e);
      }
    });
  }

  private handleIncoming(data: unknown): void {
    if (!isRpcMessage(data)) return;
    const msg = data as RpcMessage;
    if (msg.kind === 'request') return; // 主线程只收 response/error/progress/event
    if (msg.kind === 'event') {
      // server-push 事件;无 id,与某个 request 无关
      if (msg.event === WORD_CHANGE_EVENT) {
        const ev = msg.payload as KernelChangeEvent;
        for (const listener of this.changeListeners) {
          try {
            listener(ev);
          } catch {
            // 监听器抛错不影响其它订阅者
          }
        }
      }
      return;
    }
    const call = this.pending.get(msg.id);
    if (!call) return;

    if (msg.kind === 'progress') {
      call.onProgress?.(msg.phase, msg.pct);
      return;
    }
    if (msg.kind === 'response') {
      call.resolve(msg.result);
      this.pending.delete(msg.id);
      return;
    }
    if (msg.kind === 'error') {
      call.reject(rebuildError(msg.error));
      this.pending.delete(msg.id);
    }
  }

  private handleWorkerError(err: unknown): void {
    if (this.state === 'disposed') return;
    // reloading 期间再次触发(新 port 又崩 / 旧 port 的延迟 exit):把 pending + queued
    // 全部以此 err reject,停在 reloading 态避免后续 call 再跑
    if (this.state === 'reloading') {
      for (const [id, call] of this.pending) {
        call.reject(err);
        this.pending.delete(id);
      }
      // 注意:不清 queued;旧 port 的 exit 延迟信号在新 port load 成功后才 fire 也可能
      //       出现,此处保守 reject pending 即可。如果 new port load 尚未发出,pending 为空。
      return;
    }

    if (!this.reloadConfig) {
      // 未配置 reload —— 退化到旧行为,所有 pending 以底层 error reject
      for (const [id, call] of this.pending) {
        call.reject(err);
        this.pending.delete(id);
      }
      return;
    }

    this.state = 'reloading';
    this.reloadingPromise = this.reloadAfterCrash(err).catch((reloadErr) => {
      // reload 本身失败 —— 所有 queued 也以此 reject,状态保留 'reloading' 避免后续 call 再跑
      while (this.queued.length > 0) {
        const q = this.queued.shift()!;
        q.reject(reloadErr);
      }
    });
  }

  private async reloadAfterCrash(originalErr: unknown): Promise<void> {
    const cfg = this.reloadConfig!;

    // 1. 把 pending 全部以 WORKER_RELOADED reject(保留 originalErr 到 details.cause)
    const causeMsg = originalErr instanceof Error ? originalErr.message : String(originalErr);
    for (const [id, call] of this.pending) {
      call.reject(
        new WordError(
          'WORKER_RELOADED',
          `Worker 崩溃,已 reload;本 RPC 被拒收(method=${call.req.method}): ${causeMsg}`,
          cfg.docId,
          { method: call.req.method, cause: causeMsg },
        ),
      );
      this.pending.delete(id);
    }

    // 2. terminate 旧 port(旧 port 已挂,terminate 可能抛错,忽略)
    try {
      await this.port.terminate();
    } catch {
      /* ignore */
    }

    // 3. 起新 port + 绑 handler(代际序号递增,旧 port 的延迟回调被过滤)
    const newPort = await cfg.factory();
    this.portGeneration++;
    this.port = newPort;
    this.attachPortHandlers(newPort, this.portGeneration);

    // 4. 在新 port 上重放 load RPC(state 仍 reloading,sendCall 直接发)
    //    load 期间如果新 port 再崩 → handleWorkerError 会忽略(state!==idle)。
    //    因此等 load 完成再切回 idle。
    try {
      await this.sendCall('load', [cfg.loadParams]);
    } catch (loadErr) {
      // reload 自身的 load 失败:所有 queued reject;状态保持 reloading(拒绝新 call)
      while (this.queued.length > 0) {
        const q = this.queued.shift()!;
        q.reject(loadErr);
      }
      throw loadErr;
    }

    // 5. 切回 idle;在 flush queued 之前**串行 await** reload 监听器(#23),
    //    让 SnapshotApi 重建 nextId —— 否则 queued 的 take() 可能拿到陈旧自增值
    //    和 adapter 里已持久化的 id 撞车。
    this.state = 'idle';
    this.reloadingPromise = undefined;
    for (const listener of this.reloadListeners) {
      try {
        await listener();
      } catch {
        // 监听器不应抛错;防御性吞掉,不影响后续 flush
      }
    }
    const toFlush = this.queued.splice(0);
    for (const q of toFlush) {
      this.sendCall(q.method, q.args, q.onProgress).then(q.resolve, q.reject);
    }
  }

  /**
   * 注册 reload 完成回调(#23)。返回取消订阅函数。listener 可返回 Promise,
   * reload 流程会串行 await 每个 listener 完成后再 flush queued RPC。
   *
   * 触发时机:worker crash → new port + load 重放成功 → state 切回 idle 之后、
   * flush queued 之前。SnapshotApi 在此重建 nextId(从 adapter.list 读已有快照)。
   */
  onReload(listener: () => void | Promise<void>): () => void {
    this.reloadListeners.add(listener);
    return () => {
      this.reloadListeners.delete(listener);
    };
  }

  onChange(listener: (ev: KernelChangeEvent) => void): () => void {
    this.changeListeners.add(listener);
    return () => {
      this.changeListeners.delete(listener);
    };
  }

  // ─────────── 测试辅助(test-only)───────────

  /**
   * @internal test-only:手动模拟一次 worker 崩溃。
   * 触发 reload 流程(若 reloadConfig 已配),供 `worker-reload.test.ts` 使用。
   */
  _triggerCrashForTest(err: Error = new Error('synthetic crash')): void {
    this.handleWorkerError(err);
  }

  /** @internal test-only:查状态机,用于断言 reload 完成 */
  _getStateForTest(): TransportState {
    return this.state;
  }
}

function rebuildError(raw: SerializableError): Error {
  if (raw.code) {
    return new WordError(
      raw.code as WordErrorCode,
      raw.message,
      raw.docId ?? '',
      raw.details ?? {},
    );
  }
  const e = new Error(raw.message);
  if (raw.name) e.name = raw.name;
  if (raw.stack) e.stack = raw.stack;
  return e;
}
