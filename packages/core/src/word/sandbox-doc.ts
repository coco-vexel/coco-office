/**
 * sandbox-doc —— 在 sandbox 内构造与主线程 `WordDocument` 同形态的 `doc` 注入对象(#159)。
 *
 * 04.2 §6.1 声明沙箱内的 `doc` 与主线程 `WordDocument` 同名签名。本模块复用
 * `sub-apis.ts` 与 `WordPlanImpl` 的工厂(它们均 transport-agnostic),只额外
 * 提供一个最小的 sandbox transport —— 把 namespace / plan 的 RPC 调用转发到本
 * 进程内已存在的 Kernel,不经过 worker 边界。
 *
 * 注入的 `doc` 形态:
 *   - `doc.docId: string`
 *   - `doc.paragraphs.list(opts) / get(id)`
 *   - `doc.bookmarks.list(opts) / get(name)`
 *   - `doc.comments.list() / get(id)`
 *   - `doc.revisions.list(opts)`
 *   - `doc.plan(opts).<op>().apply()`
 *
 * 不暴露:
 *   - `doc.executeCode`(沙箱内递归执行无意义)
 *   - 旧的平铺 / kernel-direct 形态(`doc.applyPlan` / `doc.listParagraphs` 等):
 *     #159 决定一次性切换,不保留 backward-compat。
 */
import { kernelHandlers, type KernelMethod } from './kernel-handlers.js';
import type { Kernel, KernelProgress } from './kernel.js';
import { WordPlanImpl, type ApplyPlanParams, type PlanOptions, type WordPlan } from './plan.js';
import {
  createBookmarkApi,
  createCommentApi,
  createParagraphApi,
  createRevisionApi,
  createTablesApi,
  type BookmarkApi,
  type CommentApi,
  type ParagraphApi,
  type RevisionApi,
  type TablesApi,
} from './sub-apis.js';
import type { Transport } from './transport.js';
import { WordError } from './word-error.js';

export interface SandboxDoc {
  readonly docId: string;
  readonly paragraphs: ParagraphApi;
  readonly bookmarks: BookmarkApi;
  readonly comments: CommentApi;
  readonly revisions: RevisionApi;
  readonly tables: TablesApi;
  plan(opts?: PlanOptions): WordPlan;
}

/**
 * 构造 sandbox 用的最小 transport:只实现 `call` —— 把 RPC method 直接路由到
 * 已存在 kernel 的 handler。其它 Transport 能力(executeScript / dispose /
 * onChange)在 sandbox 内不需要也不应触发。
 *
 * `applyPlan` 特殊化:`kernelHandlers.applyPlan` 硬编码 `source: 'plan'`(主线程
 * `doc.plan().apply()` 路径),但脚本里 `doc.plan().apply()` 应当标 `source: 'script'`
 * —— 故 sandbox transport 直接调 `kernel.applyPlan(params, { source: 'script' })`,
 * 绕过 handler 默认值。
 */
function createSandboxTransport(kernel: Kernel): Transport {
  return {
    async call<T>(method: KernelMethod, args: unknown[], onProgress?: KernelProgress): Promise<T> {
      if (method === 'applyPlan') {
        const [params] = args as [ApplyPlanParams];
        const result = await kernel.applyPlan(params, { source: 'script' });
        return result as T;
      }
      const handler = kernelHandlers[method];
      const result = await handler({ kernel, onProgress }, args);
      return result as T;
    },
    executeScript() {
      return Promise.reject(
        new Error('nested executeScript is not supported inside script sandbox'),
      );
    },
    dispose() {
      return Promise.resolve();
    },
    onChange() {
      return () => {};
    },
  };
}

/**
 * 把 kernel 包成与主线程 `WordDocument` 同形态的 sandbox `doc`。
 *
 * 复用 sub-apis.ts 的 namespace 工厂 + WordPlanImpl —— 这两者只通过
 * `transport.call(method, args)` 访问 kernel,与主线程走同一条路径。
 *
 * 不做 `assertNotDisposed` 检查:sandbox 内 kernel 由宿主管理,脚本生命周期
 * 短于 kernel,disposed 状态不会出现。
 */
export function createSandboxDoc(kernel: Kernel): SandboxDoc {
  const docId = kernel.getDocId();
  if (!docId) {
    throw new WordError(
      'PACKAGE_UNSUPPORTED_FORMAT',
      'createSandboxDoc: kernel not loaded (docId is undefined)',
      '',
    );
  }
  const transport = createSandboxTransport(kernel);
  const noop = () => {};
  return {
    docId,
    paragraphs: createParagraphApi(transport, noop),
    bookmarks: createBookmarkApi(transport, noop),
    comments: createCommentApi(transport, noop),
    revisions: createRevisionApi(transport, noop),
    tables: createTablesApi(transport, noop),
    plan(opts?: PlanOptions): WordPlan {
      const resolved: PlanOptions = {
        trackChanges: opts?.trackChanges ?? false,
      };
      if (opts?.author !== undefined) resolved.author = opts.author;
      if (opts?.date !== undefined) resolved.date = opts.date;
      return new WordPlanImpl({ docId, transport, options: resolved });
    },
  };
}
