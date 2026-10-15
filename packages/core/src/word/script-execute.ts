/**
 * script-execute —— 在 worker / inline 内执行用户脚本(04.2 §六冻结的运行时环境)。
 *
 * 入口:
 *   `executeScript(kernel, code, args, timeoutMs)`
 *
 * 执行形态:
 *   `new Function('doc', 'args', 'fs', wrappedBody)(doc, args, fs)`
 *   wrappedBody 把用户 code 包进 async IIFE,允许顶层 await。
 *
 * 安全 / 隔离:
 *   - v1 不开启沙箱;脚本能访问 worker 全局(setTimeout / fetch 等 worker 原生能力)
 *   - 不暴露 require / 动态 import(由 `new Function` 隔离作用域天然不可见)
 *   - 错误穿过 worker 边界时序列化(由 worker-entry 处理)
 *
 * Timeout 简化实施(v1):
 *   - `timeoutMs > 0`:Promise.race 超时拒绝;**脚本本身继续运行**(setTimeout 不能终止
 *     运行中代码),Worker kill+reload 由 04.2 §四 留给后续 issue
 *   - `timeoutMs === 0`:不设超时
 *
 * doc 的形态(04.2 §6.1,#159 落地):
 *   - 注入与主线程 `WordDocument` **同形态**的 `doc`:`doc.paragraphs.list()` /
 *     `doc.bookmarks` / `doc.comments` / `doc.revisions` / `doc.plan().xxx().apply()` /
 *     `doc.docId`
 *   - 不暴露旧的平铺 / kernel-direct 形态(`doc.applyPlan` / `doc.listParagraphs` 等)
 *   - 不暴露 `doc.executeCode`(沙箱内递归无意义)
 *   - 实现细节见 sandbox-doc.ts —— 复用 namespace 工厂与 WordPlanImpl,加一层
 *     最小 sandbox transport,直接路由到本进程 kernel
 *
 * @see docs/04.2-脚本执行-Tool设计.md §二 §六
 */

import type { Kernel } from './kernel.js';
import { createSandboxDoc } from './sandbox-doc.js';

/**
 * 在当前线程内执行一段用户脚本,返回脚本顶层 return 的值。
 *
 * @param kernel 当前 worker / inline 的 Kernel 实例,作为脚本 `doc` 暴露
 * @param code   用户提供的 JS 源码;函数体形态(可以包含顶层 await / return)
 * @param args   传给脚本的参数,经 structured clone 过 worker 边界
 * @param timeoutMs 超时毫秒数;0 = 不超时
 * @param signal 可选 AbortSignal(#151):
 *   - InlineTransport 调本函数时透传:abort 触发 → race 拒绝(主线程立即返回,
 *     但脚本仍在后台跑直到自然结束 —— inline 无 worker 可 terminate)
 *   - WorkerTransport 在主线程侧用 signal 触发 \`worker.terminate()\`,worker 内
 *     的 executeScript 不直接收 signal(传 \`undefined\` 即可)
 *
 * @throws  脚本本身抛出的错误 / 语法错误 / 超时 / abort
 */
export async function executeScript(
  kernel: Kernel,
  code: string,
  args: Record<string, unknown> | undefined,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<unknown> {
  // 用户 code 可能包含顶层 return / await。包进 async IIFE,
  // 既保留 return 语义也允许 await。
  const wrappedBody = `return (async () => { ${code} })()`;

  let scriptFn: (doc: unknown, args: unknown, fs: unknown) => Promise<unknown>;
  try {
    scriptFn = new Function('doc', 'args', 'fs', wrappedBody) as (
      doc: unknown,
      args: unknown,
      fs: unknown,
    ) => Promise<unknown>;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`script syntax error: ${msg}`);
  }

  // doc = 与主线程 WordDocument 同形态的 facade(#159);见 sandbox-doc.ts。
  // fs 在 worker 内**不**注入(04.2 §6.1:"仅当宿主注入时存在",worker 边界外的
  // FileSystemPort 在 v1 内不跨边界传递 —— 主线程 plugin 的 fs 在加载脚本前已用过)。
  const doc = createSandboxDoc(kernel);
  const scriptPromise = (async () => scriptFn(doc, args, undefined))();

  // 提前触发的 abort:不启动脚本(实际上脚本已构造但 race 立即拒绝)
  if (signal?.aborted) {
    throw new Error('script aborted');
  }

  // 收集需要 race 的 promises:scriptPromise / timeout / abort
  const racers: Promise<unknown>[] = [scriptPromise];
  let timeoutHandle: ReturnType<typeof setTimeout> | undefined;
  let abortHandler: (() => void) | undefined;
  const cleanups: (() => void)[] = [];

  if (timeoutMs > 0) {
    racers.push(
      new Promise<never>((_, reject) => {
        timeoutHandle = setTimeout(() => {
          reject(new Error(`script timeout after ${timeoutMs}ms`));
        }, timeoutMs);
      }),
    );
    cleanups.push(() => {
      if (timeoutHandle) clearTimeout(timeoutHandle);
    });
  }

  if (signal) {
    racers.push(
      new Promise<never>((_, reject) => {
        abortHandler = () => reject(new Error('script aborted'));
        signal.addEventListener('abort', abortHandler, { once: true });
      }),
    );
    cleanups.push(() => {
      if (abortHandler) signal.removeEventListener('abort', abortHandler);
    });
  }

  try {
    return await Promise.race(racers);
  } finally {
    for (const cleanup of cleanups) cleanup();
  }
}
