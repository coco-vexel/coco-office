/**
 * Worker RPC 协议 —— docs/03 §4.1 / §4.7。
 *
 * 主线程(WordDocument 代理)与 Worker 内 kernel 之间通过 postMessage 交换五类消息:
 *   - request:  { kind: 'request',  id, method, args }     —— 主 → Worker
 *   - response: { kind: 'response', id, result }           —— Worker → 主(成功)
 *   - error:    { kind: 'error',    id, error }            —— Worker → 主(失败)
 *   - progress: { kind: 'progress', id, phase, pct? }      —— Worker → 主(onProgress 回调)
 *   - event:    { kind: 'event',    event, payload }       —— Worker → 主(server-push,
 *               与某个 request 无关;v1 仅用于 'word.change')
 *
 * id 是一个 UUID-ish 字符串,主线程用它关联 pending request 与返回的 response/error。
 *
 * 错误使用 serializable 形态(code / message / docId / details),跨 Worker 边界后主
 * 线程用它重建 `WordError` 实例 → 抛给调用方。
 */
import type { WordErrorCode } from './word-error.js';

/** 主线程 → Worker 的方法调用 */
export interface RpcRequest {
  kind: 'request';
  id: string;
  method: string;
  args: unknown[];
}

/** Worker → 主线程:方法成功返回 */
export interface RpcResponse {
  kind: 'response';
  id: string;
  result: unknown;
}

/** Worker → 主线程:方法抛错。错误以字段形态序列化,主线程侧重建 WordError */
export interface RpcError {
  kind: 'error';
  id: string;
  error: SerializableError;
}

/** 序列化后的错误负载(跨 Worker 边界友好) */
export interface SerializableError {
  /** WordError.code;如果抛的是非 WordError,此字段为 undefined,主线程包装成 PACKAGE_CORRUPT */
  code?: WordErrorCode;
  message: string;
  /** WordError.docId;非 WordError 时可能为空串 */
  docId?: string;
  details?: Record<string, unknown>;
  /** 原始错误 name(Error/TypeError/...),调试用 */
  name?: string;
  /** 原始 stack(调试用;跨 worker 后行号不可靠,仅参考) */
  stack?: string;
}

/** Worker → 主线程:onProgress 回调 */
export interface RpcProgress {
  kind: 'progress';
  /** 关联的 request id */
  id: string;
  phase: string;
  pct?: number;
}

/**
 * Worker → 主线程:server-push 事件。与某个 request 无关,因此**没有** id 字段。
 * v1 仅用于 `'word.change'`(applyPlan 成功后从 worker 推到主线程,主线程 fan-out
 * 给 `WordDocument.on('change', ...)` 订阅者)。
 *
 * payload 必须 structured-cloneable;接收侧按 event 名做类型 narrowing。
 */
export interface RpcEvent {
  kind: 'event';
  event: string;
  payload: unknown;
}

export type RpcMessage = RpcRequest | RpcResponse | RpcError | RpcProgress | RpcEvent;

/** type guard */
export function isRpcMessage(v: unknown): v is RpcMessage {
  if (typeof v !== 'object' || v === null) return false;
  const m = v as { kind?: string };
  return (
    m.kind === 'request' ||
    m.kind === 'response' ||
    m.kind === 'error' ||
    m.kind === 'progress' ||
    m.kind === 'event'
  );
}
