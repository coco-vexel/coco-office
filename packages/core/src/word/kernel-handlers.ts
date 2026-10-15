/**
 * kernel-handlers —— RPC 方法到 Kernel 方法的分发表。
 *
 * 同一张表被 InlineTransport(主线程直接调)与 kernel-worker-entry(Worker 内通过
 * postMessage 触发)共用,保证语义一致。
 */
import type { BookmarkName, ListBookmarksOptions } from './bookmark-queries.js';
import type { CommentId } from './comment-queries.js';
import type { ContextSearchOptions } from './context-search.js';
import type { Kernel, KernelLoadParams, KernelProgress, SnapshotPartEntry } from './kernel.js';
import type { ListParagraphsOptions } from './list-paragraphs.js';
import type { ListTablesOptions } from './list-tables.js';
import type { ListDrawingsOptions } from './list-drawings.js';
import type { ApplyPlanParams } from './plan.js';
import type { ListRevisionsOptions } from './revision-queries.js';
import type { ListStylesOptions } from './style-queries.js';

export type KernelMethod =
  | 'load'
  | 'toZip'
  | 'toFlatOpc'
  | 'dispose'
  | 'serializeParts'
  | 'rebuildFromParts'
  // 读 namespace 方法(#92,docs/03.1 §2);dots 在字符串里没有歧义,见 sub-apis.ts
  | 'paragraphs.list'
  | 'paragraphs.get'
  | 'bookmarks.list'
  | 'bookmarks.get'
  | 'comments.list'
  | 'comments.get'
  | 'context.search'
  | 'revisions.list'
  | 'revisions.get'
  | 'styles.list'
  | 'styles.get'
  | 'tables.list'
  | 'tables.get'
  | 'tables.getVerbatim'
  | 'drawings.list'
  | 'drawings.get'
  | 'drawings.getImageBytes'
  | 'validatePlan'
  | 'sections.list'
  | 'tocFields.list'
  | 'applyPlan';

export interface KernelHandlerContext {
  kernel: Kernel;
  onProgress?: KernelProgress;
}

/** 每个 method 的 handler 签名:接 context + args,返 Promise<result> */
export type KernelHandler = (ctx: KernelHandlerContext, args: unknown[]) => Promise<unknown>;

export const kernelHandlers: Record<KernelMethod, KernelHandler> = {
  load: async ({ kernel, onProgress }, args) => {
    const [params] = args as [KernelLoadParams];
    await kernel.load(params, onProgress);
    // 返回装载后的可导出快照,便于主线程代理保留 docId / 初始字节等(为 worker 模式设计)
    return {
      docId: kernel.getDocId(),
      hasInitialBytes: kernel.getInitialBytes() !== undefined,
      hasInitialFlatOpc: kernel.getInitialFlatOpc() !== undefined,
    };
  },
  toZip: async ({ kernel }) => kernel.toZip(),
  toFlatOpc: async ({ kernel }) => kernel.toFlatOpc(),
  dispose: async ({ kernel }) => {
    kernel.dispose();
  },
  serializeParts: async ({ kernel }) => kernel.serializeParts(),
  rebuildFromParts: async ({ kernel, onProgress }, args) => {
    const [parts] = args as [readonly SnapshotPartEntry[]];
    kernel.rebuildFromParts(parts, onProgress);
  },
  'paragraphs.list': async ({ kernel }, args) => {
    const [opts] = args as [ListParagraphsOptions | undefined];
    return kernel.listParagraphs(opts);
  },
  'paragraphs.get': async ({ kernel }, args) => {
    const [paraId] = args as [string];
    return kernel.getParagraph(paraId);
  },
  'bookmarks.list': async ({ kernel }, args) => {
    const [opts] = args as [ListBookmarksOptions | undefined];
    return kernel.listBookmarks(opts);
  },
  'bookmarks.get': async ({ kernel }, args) => {
    const [name] = args as [BookmarkName];
    return kernel.getBookmark(name);
  },
  'comments.list': async ({ kernel }) => kernel.listComments(),
  'comments.get': async ({ kernel }, args) => {
    const [id] = args as [CommentId];
    return kernel.getComment(id);
  },
  'context.search': async ({ kernel }, args) => {
    const [opts] = args as [ContextSearchOptions];
    return kernel.searchContext(opts);
  },
  'revisions.list': async ({ kernel }, args) => {
    const [opts] = args as [ListRevisionsOptions | undefined];
    return kernel.listRevisions(opts);
  },
  'revisions.get': async ({ kernel }, args) => {
    const [revId, opts] = args as [string, ListRevisionsOptions | undefined];
    return kernel.getRevision(revId, opts);
  },
  'styles.list': async ({ kernel }, args) => {
    const [opts] = args as [ListStylesOptions | undefined];
    return kernel.listStyles(opts);
  },
  'styles.get': async ({ kernel }, args) => {
    const [styleId] = args as [string];
    return kernel.getStyle(styleId);
  },
  'tables.list': async ({ kernel }, args) => {
    const [opts] = args as [ListTablesOptions | undefined];
    return kernel.listTables(opts);
  },
  'tables.get': async ({ kernel }, args) => {
    const [tableId] = args as [string];
    return kernel.getTable(tableId);
  },
  'tables.getVerbatim': async ({ kernel }, args) => {
    const [tableId] = args as [string];
    return kernel.getTableVerbatim(tableId);
  },
  'drawings.list': async ({ kernel }, args) => {
    const [opts] = args as [ListDrawingsOptions | undefined];
    return kernel.listDrawings(opts);
  },
  'drawings.get': async ({ kernel }, args) => {
    const [drawingId] = args as [string];
    return kernel.getDrawing(drawingId);
  },
  'drawings.getImageBytes': async ({ kernel }, args) => {
    const [drawingId] = args as [string];
    return kernel.getImageBytes(drawingId);
  },
  'sections.list': async ({ kernel }) => kernel.listSections(),
  'tocFields.list': async ({ kernel }) => kernel.listTocFields(),
  applyPlan: async ({ kernel }, args) => {
    const [params] = args as [ApplyPlanParams];
    // RPC 路径明确来源 'plan',对应 `doc.plan().apply()`(脚本路径走 script-execute,
    // 不进本 handler;那里 source 默认为 'script')。
    return kernel.applyPlan(params, { source: 'plan' });
  },
  validatePlan: async ({ kernel }, args) => {
    const [params] = args as [ApplyPlanParams];
    return kernel.validatePlan(params);
  },
};
