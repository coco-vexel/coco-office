/**
 * kernel-worker-shared —— dispatcher 单元测试。
 *
 * 不起 Worker(Node 也好浏览器也好),直接调用 `createKernelDispatcher` + 注入 `post`
 * spy。验证消息 dispatch / 错误序列化 / 非 request 忽略 / unknown method 这些环境无关
 * 行为。
 *
 * 真实跨线程的 e2e 由 plan-worker / list-paragraphs-worker 等已有测试覆盖(走 Node
 * worker_threads bundle),浏览器 worker 真实跑由 #52 覆盖。
 */
import { describe, expect, test, vi } from 'vitest';
import { createKernelDispatcher, serializeError } from '../kernel-worker-shared.js';
import { WordError } from '../word-error.js';
import type { RpcRequest } from '../rpc-protocol.js';

describe('createKernelDispatcher', () => {
  test('非 request 消息直接忽略', async () => {
    const post = vi.fn();
    const dispatch = createKernelDispatcher(post);
    await dispatch({ kind: 'response', id: '1', result: null });
    await dispatch({ kind: 'progress', id: '1', phase: 'parse' });
    await dispatch(undefined);
    expect(post).not.toHaveBeenCalled();
  });

  test('dispose 请求 → 单条 response,id 透传', async () => {
    const post = vi.fn();
    const dispatch = createKernelDispatcher(post);
    const req: RpcRequest = { kind: 'request', id: 'r-1', method: 'dispose', args: [] };
    await dispatch(req);
    expect(post).toHaveBeenCalledTimes(1);
    expect(post).toHaveBeenCalledWith({ kind: 'response', id: 'r-1', result: undefined });
  });

  test('unknown method → error,message 含方法名', async () => {
    const post = vi.fn();
    const dispatch = createKernelDispatcher(post);
    const req = {
      kind: 'request',
      id: 'r-2',
      method: 'no.such.method',
      args: [],
    } as unknown as RpcRequest;
    await dispatch(req);
    expect(post).toHaveBeenCalledTimes(1);
    const [msg] = post.mock.calls[0] as [{ kind: string; id: string; error: { message: string } }];
    expect(msg.kind).toBe('error');
    expect(msg.id).toBe('r-2');
    expect(msg.error.message).toContain('no.such.method');
  });

  test('handler 抛 WordError → 序列化保留 code/docId/details', async () => {
    const post = vi.fn();
    const dispatch = createKernelDispatcher(post);
    // toZip 在 kernel 未 load 时会抛 WordError(PACKAGE_UNSUPPORTED_FORMAT)
    const req: RpcRequest = { kind: 'request', id: 'r-3', method: 'toZip', args: [] };
    await dispatch(req);
    expect(post).toHaveBeenCalledTimes(1);
    const [msg] = post.mock.calls[0] as [
      { kind: string; id: string; error: { code?: string; message: string; name?: string } },
    ];
    expect(msg.kind).toBe('error');
    expect(msg.id).toBe('r-3');
    expect(msg.error.code).toBeDefined();
    expect(msg.error.name).toBe('WordError');
  });

  test('script.execute 走独立分支 → 语法错抛回 error', async () => {
    const post = vi.fn();
    const dispatch = createKernelDispatcher(post);
    const req = {
      kind: 'request',
      id: 'r-4',
      method: 'script.execute',
      args: ['(((', undefined, 1000],
    } as unknown as RpcRequest;
    await dispatch(req);
    expect(post).toHaveBeenCalledTimes(1);
    const [msg] = post.mock.calls[0] as [{ kind: string; id: string; error: { message: string } }];
    expect(msg.kind).toBe('error');
    expect(msg.id).toBe('r-4');
  });

  test('每个 dispatcher 独立 Kernel —— 互不污染', async () => {
    const post1 = vi.fn();
    const post2 = vi.fn();
    const d1 = createKernelDispatcher(post1);
    const d2 = createKernelDispatcher(post2);
    const req: RpcRequest = { kind: 'request', id: 'x', method: 'dispose', args: [] };
    await d1(req);
    await d2(req);
    expect(post1).toHaveBeenCalledTimes(1);
    expect(post2).toHaveBeenCalledTimes(1);
  });
});

describe('serializeError', () => {
  test('WordError → 字段化(code/docId/details/name)', () => {
    const err = new WordError('PACKAGE_UNSUPPORTED_FORMAT', 'boom', 'doc-x', { extra: 1 });
    const out = serializeError(err);
    expect(out.code).toBe('PACKAGE_UNSUPPORTED_FORMAT');
    expect(out.message).toBe('boom');
    expect(out.docId).toBe('doc-x');
    expect(out.details).toEqual({ extra: 1 });
    expect(out.name).toBe('WordError');
  });

  test('普通 Error → message + name + stack', () => {
    const err = new TypeError('nope');
    const out = serializeError(err);
    expect(out.message).toBe('nope');
    expect(out.name).toBe('TypeError');
    expect(typeof out.stack).toBe('string');
  });

  test('非 Error 值 → message = String(value)', () => {
    expect(serializeError('plain').message).toBe('plain');
    expect(serializeError(42).message).toBe('42');
    expect(serializeError(null).message).toBe('null');
  });
});
