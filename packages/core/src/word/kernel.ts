interface CollectAffectedPartPathTiming {
  storiesMs: number;
  storyLookups: number;
  storyLookupMs: number;
  byType: Array<{
    type: Op['type'];
    count: number;
    ms: number;
    storyLookups: number;
    storyLookupMs: number;
  }>;
}

interface CollectAffectedPartPathsResult {
  paths: string[];
  timing: CollectAffectedPartPathTiming;
}

/**
 * Kernel —— 持有 Package + 三索引 + 执行装载/输出/变更的"真实运行体"。
 *
 * - Worker 模式:kernel 在 Worker 里跑,主线程通过 RPC 协议与之通信
 * - Inline 模式:kernel 直接在主线程跑,`InlineTransport` 调用它的同步 API
 *
 * 本文件不依赖任何 Worker / 主线程特定 API,可被两端复用。
 */
import type { Package, Part } from '../pack/index.js';
import {
  createBinaryPart,
  createPackage,
  createXmlPart,
  fromFlatOpc as pkgFromFlatOpc,
  fromZip as pkgFromZip,
  isXmlPart,
  listParts,
  setPart,
  toFlatOpc as pkgToFlatOpc,
  toZip as pkgToZip,
} from '../pack/index.js';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import {
  getBookmark as pureGetBookmark,
  listBookmarks as pureListBookmarks,
  type BookmarkInfo,
  type BookmarkName,
  type ListBookmarksOptions,
} from './bookmark-queries.js';
import type { KernelChangeEvent } from './change-events.js';
import {
  getComment as pureGetComment,
  listComments as pureListComments,
  type CommentId,
  type CommentInfo,
} from './comment-queries.js';
import {
  searchContext as pureSearchContext,
  type ContextSearchOptions,
  type ContextSearchResult,
} from './context-search.js';
import { BookmarkLoadError, buildIndices, type PackageIndices } from './indices.js';
import {
  getParagraphByParaId,
  listParagraphs as pureListParagraphs,
  type ListParagraphsOptions,
  type ListResult,
  type ParagraphInfo,
} from './list-paragraphs.js';
import {
  WordDomainCache,
  type DomainCacheName,
  type WordDomainCacheSnapshot,
} from './domain-cache.js';
import { ensureParaIds } from './para-id.js';
import { EMPTY_APPLY_RESULT, type ApplyPlanParams, type ApplyResult, type Op } from './plan.js';
import { validatePlanOps } from './plan-validation.js';
import { dispatchOp, type ExecutorContext } from './plan-executors/index.js';
import { sweepOrphanMediaParts } from './plan-executors/placeholder-image.js';
import { runTransaction } from './transactions.js';
import { listStories, type Story } from './stories.js';
import {
  listRevisions as pureListRevisions,
  type ListRevisionsOptions,
  type RevisionInfo,
} from './revision-queries.js';
import { listSections as pureListSections, type SectionInfo } from './section-queries.js';
import { listTocFields as pureListTocFields, type TocFieldRange } from './toc-queries.js';
import {
  listStyles as pureListStyles,
  getStyleById as pureGetStyleById,
  type ListStylesOptions,
  type StyleInfo,
} from './style-queries.js';
import {
  listTables as pureListTables,
  getTableById as pureGetTableById,
  getTableVerbatim as pureGetTableVerbatim,
  type ListTablesOptions,
  type TableInfo,
  type TableVerbatim,
} from './list-tables.js';
import {
  listDrawings as pureListDrawings,
  getDrawingById as pureGetDrawingById,
  type DrawingInfo,
  type ListDrawingsOptions,
} from './list-drawings.js';
import { WordError } from './word-error.js';
import { getImageBytesSync, type ImageBytesResult } from './plan-executors/insert-image-ops.js';

/** Kernel.load 的输入 */
export type KernelLoadSource =
  { type: 'zip'; bytes: Uint8Array } | { type: 'flat-opc'; xml: string };

export interface KernelLoadParams {
  source: KernelLoadSource;
  docId: string;
}

/**
 * 装载 / rebuild 进度回调。
 *
 * - load 使用 `'fetch' | 'parse' | 'normalize'`(LoadPhase 对齐)
 * - rebuildFromParts(#22)使用 `'parse' | 'build-indices' | 'done'`
 *
 * 上层 WordDocument.rollback 会把 phase 作为字符串透传,不强制 union 收敛。
 */
export type KernelProgress = (phase: string, pct?: number) => void;

/**
 * Kernel —— 持有单一 WordDocument 的内部状态。
 *
 * 一个 WordDocument 代理对应一个 Kernel。生命周期由代理控制:load → 多次操作 → dispose。
 */
export class Kernel {
  private pkg?: Package;
  private indices?: PackageIndices;
  private docId?: string;
  private initialBytes?: Uint8Array;
  private initialFlatOpc?: string;
  private disposed = false;
  private readonly domainCache = new WordDomainCache();
  /**
   * change 事件订阅者(applyPlan 成功后 fan-out)。InlineTransport 直接订阅;worker
   * 模式下 worker 内创建 Kernel 时由 worker entry 订阅一个把事件 post 到主线程的
   * 中转 listener(详见 kernel-worker-shared.ts)。
   *
   * 监听器异常被静默吞掉 —— 一个订阅者抛错不该影响其它订阅者,也不该影响 applyPlan
   * 的返回值(applyPlan 已经成功提交,事件只是事后通知)。
   */
  private readonly changeListeners = new Set<(ev: KernelChangeEvent) => void>();

  /**
   * applyPlan 互斥链 —— 串行化所有 applyPlan 调用,不论来源(plan / script)。
   *
   * 必要性:applyPlan 内部对 `pkg.parts[*].xmlDom` 做 deepClone(transaction
   * snapshot)+ 同步 mutate + 索引重建。多个 applyPlan 在 microtask 边界穿插
   * 时:
   *   - 多份 deepClone 同时占内存(50w 字 doc × 并发数 → WPS WebView OOM)
   *   - A 的 dispatch 与 B 的 dispatch 在共享 xmlDom 上交错,paraIdx 错位
   *   - A 的回滚 snapshot 包含 B 的 mutation,A 失败回滚会 wipe B 的写
   *   - finally 里 buildIndices 互相覆盖
   *
   * 用 promise chain 串行化:每次 applyPlan 等前一个完成才进入 inner。链上
   * resolve 永不 reject —— 把 inner 的成功/失败结果在外层 await 时直接抛给
   * 调用方,链本身只保证"前一个跑完了"。
   */
  private _applyPlanChain: Promise<void> = Promise.resolve();

  // ─────────── 装载 ───────────

  async load(params: KernelLoadParams, onProgress?: KernelProgress): Promise<void> {
    this.assertNotDisposed();
    const { source, docId } = params;

    onProgress?.('parse', 0);
    let pkg: Package;
    try {
      if (source.type === 'zip') {
        pkg = await pkgFromZip(source.bytes);
        this.initialBytes = source.bytes;
      } else {
        pkg = pkgFromFlatOpc(source.xml);
        this.initialFlatOpc = source.xml;
      }
    } catch (e) {
      const phase = source.type === 'zip' ? 'fromZip' : 'fromFlatOpc';
      throw wrapPackageError(e, docId, phase);
    }

    onProgress?.('normalize', 0);
    ensureParaIds(pkg);
    try {
      this.indices = buildIndices(pkg);
    } catch (e) {
      if (e instanceof BookmarkLoadError) {
        throw new WordError(e.code, e.message, docId, e.details);
      }
      throw wrapPackageError(e, docId, 'normalize');
    }
    onProgress?.('normalize', 100);

    this.pkg = pkg;
    this.docId = docId;
    this.domainCache.reset();
  }

  // ─────────── 输出 ───────────

  async toZip(): Promise<Uint8Array> {
    this.assertLoaded();
    try {
      return await pkgToZip(this.pkg!);
    } catch (e) {
      throw wrapPackageError(e, this.docId!, 'toZip');
    }
  }

  async toFlatOpc(): Promise<string> {
    this.assertLoaded();
    try {
      return pkgToFlatOpc(this.pkg!);
    } catch (e) {
      throw wrapPackageError(e, this.docId!, 'toFlatOpc');
    }
  }

  // ─────────── 只读查询(#24 起)───────────

  /**
   * 段落查询。纯同步,基于当前已装载的 pkg。
   *
   * 所有策略(scope / filter / textFormat / pagination)由入参决定,kernel **不**
   * 做"智能默认"——未传 pagination 时按 `DEFAULT_PAGINATION`(byCount(25,0))返回
   * 第一页;要全量请显式 `{ pagination: { mode: 'all' } }`(docs/03 §9.2)。
   *
   * 纯逻辑实现在 `list-paragraphs.ts`,本方法只做:装载态校验 + pkg 传入 + 错误
   * 包装(validation Error → WordError,附 docId)。
   */
  listParagraphs(opts?: ListParagraphsOptions): ListResult {
    this.assertLoaded();
    try {
      const cached = this.domainCache.paragraphs.getCacheableList(opts, () =>
        pureListParagraphs(this.pkg!, opts),
      );
      if (cached) return cached;
      return pureListParagraphs(this.pkg!, opts);
    } catch (e) {
      if (e instanceof WordError) throw e;
      const msg = e instanceof Error ? e.message : String(e);
      // listParagraphs 的抛错主要是 pagination 参数校验失败;归到 UNSUPPORTED_OOXML_FEATURE
      // 语义贴近"调用方传了不合法的选项",优于 PACKAGE_CORRUPT / PACKAGE_UNSUPPORTED_FORMAT
      throw new WordError('UNSUPPORTED_OOXML_FEATURE', msg, this.docId!);
    }
  }

  /**
   * 按 paraId 精确取段(#25)。走 paraIdx → O(1) 查找;不存在返回 null。
   * paraId 大小写不敏感(索引 key 是大写 8 hex)。
   */
  getParagraph(paraId: string): ParagraphInfo | null {
    this.assertLoaded();
    return this.domainCache.paragraphs.getParagraph(paraId, () =>
      getParagraphByParaId(this.pkg!, this.indices!, paraId),
    );
  }

  /**
   * 枚举书签(#25)。默认过滤 `_` 开头的隐藏书签(_Toc... / _GoBack / _Hlk...),
   * `opts.includeHidden=true` 可拿到全部。返回顺序遵循 bookmarkIdx 的迭代序 ——
   * JS Map 按插入序,装载期从上到下、逐 part 逐 `<w:bookmarkStart>` 收集。
   */
  listBookmarks(opts?: ListBookmarksOptions): BookmarkInfo[] {
    this.assertLoaded();
    return pureListBookmarks(this.pkg!, this.indices!, opts);
  }

  /**
   * 按 name 精确取书签(#25)。走 bookmarkIdx;不存在返回 null。
   * 不做 hidden 过滤 —— 调用方已知 name,通常就是来找隐藏书签的。
   */
  getBookmark(name: BookmarkName): BookmarkInfo | null {
    this.assertLoaded();
    return pureGetBookmark(this.pkg!, this.indices!, name);
  }

  /**
   * 枚举批注(#26)。走 commentIdx,按 Map 插入序(= 装载期 part 顺序 + w:id 升序)
   * 返回。主体缺失或无锚点的 entry 被跳过 —— 调用方看到的都是"可用"的批注。
   */
  listComments(): CommentInfo[] {
    this.assertLoaded();
    return pureListComments(this.indices!);
  }

  /**
   * 按 id 取批注(#26)。不存在或主体缺失返回 null 。
   */
  getComment(id: CommentId): CommentInfo | null {
    this.assertLoaded();
    return pureGetComment(this.indices!, id);
  }

  /**
   * 枚举修订(#26 + #33,Tier 1)。只覆盖 `<w:ins>` / `<w:del>` 标签级修订。
   *
   * **#33 加强**:扫到 `<w:rPrChange>` / `<w:pPrChange>` / `sectPrChange` /
   * `numberingChange` / `moveFrom` / `moveTo` 等 Tier 2 属性级修订 → 抛
   * `UNSUPPORTED_OOXML_FEATURE`(以前是默默忽略)。读写共用守护,与
   * `acceptRevisions` / `rejectRevisions` 行为一致。
   *
   * scope 与 `listParagraphs` 一致:默认 'body',可传 'all' 或 `StoryId[]`。
   */
  listRevisions(opts?: ListRevisionsOptions): RevisionInfo[] {
    this.assertLoaded();
    try {
      return pureListRevisions(this.pkg!, opts);
    } catch (e) {
      // 纯函数 assertNoTier2Revisions 抛错时 docId 占位 '';kernel 层注入真实 docId
      if (e instanceof WordError && e.docId === '') {
        throw new WordError(e.code, e.message, this.docId!, e.details);
      }
      throw e;
    }
  }

  getRevision(revId: string, opts?: ListRevisionsOptions): RevisionInfo | null {
    return this.listRevisions(opts).find((revision) => revision.revId === revId) ?? null;
  }

  searchContext(opts: ContextSearchOptions): ContextSearchResult {
    this.assertLoaded();
    return pureSearchContext(this.pkg!, this.indices!, opts);
  }

  // ─────────── 样式查询(#181)───────────

  /** 列出 styles.xml 中的样式，可按 type 过滤。styles.xml 不存在时返回空数组。 */
  listStyles(opts?: ListStylesOptions): StyleInfo[] {
    this.assertLoaded();
    return pureListStyles(this.indices!.stylesIdx, opts);
  }

  /** 按 styleId 精确查找，O(1)。不存在返回 null。 */
  getStyle(styleId: string): StyleInfo | null {
    this.assertLoaded();
    return pureGetStyleById(this.indices!.stylesIdx, styleId);
  }

  // ─────────── 表格查询（#185）───────────

  /** 枚举表格，默认只返回主文档 body 的顶层表格。*/
  listTables(opts?: ListTablesOptions): TableInfo[] {
    this.assertLoaded();
    const cached = this.domainCache.tables.getCacheableList(opts, () =>
      pureListTables(this.pkg!, this.indices!, opts),
    );
    if (cached) return cached;
    return pureListTables(this.pkg!, this.indices!, opts);
  }

  /** 按 tableId 精确取单个表格，O(1)。不存在返回 null。 */
  getTable(tableId: string): TableInfo | null {
    this.assertLoaded();
    return this.domainCache.tables.getTable(tableId, () =>
      pureGetTableById(this.pkg!, this.indices!, tableId),
    );
  }

  /**
   * 按 tableId verbatim 取一张表的 OOXML 子树 + 格内图字节
   * （task-20260616-table-content-driven，供跨文档搬表 insertTableFrom）。
   * 不存在返回 null。
   */
  getTableVerbatim(tableId: string): TableVerbatim | null {
    this.assertLoaded();
    return pureGetTableVerbatim(this.pkg!, this.indices!, tableId);
  }

  // ─────────── Drawing 枚举（#189）───────────

  /** 枚举 drawing 对象，默认只返回主文档 body 内的 drawing。*/
  listDrawings(opts?: ListDrawingsOptions): DrawingInfo[] {
    this.assertLoaded();
    return pureListDrawings(this.pkg!, this.indices!, opts);
  }

  /** 按 drawingId 精确取单个 drawing，O(1)。不存在返回 null。 */
  getDrawing(drawingId: string): DrawingInfo | null {
    this.assertLoaded();
    return pureGetDrawingById(this.pkg!, this.indices!, drawingId);
  }

  /**
   * 读取 drawingId 对应 picture 的原始图片字节（task-insert-image）。
   * drawingId 不存在 / 非 picture / media 找不到 → 返回 null（不抛）。
   */
  getImageBytes(drawingId: string): ImageBytesResult | null {
    this.assertLoaded();
    return getImageBytesSync(this.pkg!, this.indices!, drawingId);
  }

  /** 枚举主 body 的所有节（#211）。doc-order 排序。*/
  listSections(): SectionInfo[] {
    this.assertLoaded();
    return pureListSections(this.pkg!);
  }

  /** 枚举主 body 的所有 TOC 域（#197）。doc-order 排序。*/
  listTocFields(): TocFieldRange[] {
    this.assertLoaded();
    return pureListTocFields(this.pkg!);
  }

  // ─────────── 变更入口:Plan apply(#27)───────────

  /**
   * 执行 WordPlan 提交的 op 队列(docs/03.1 §3.2)。
   *
   * **#27 作用域**:本方法只承担 **Plan 容器 + ApplyResult 结构 + 错误状态机**。
   * 真正的 op executor 由后续 PR 分别落地(见 `plan.ts` 顶部注释)。因此本版本:
   *   - ops 为空 → 返回 `EMPTY_APPLY_RESULT`(WordPlanImpl 侧通常已拦,这里兜底)
   *   - ops 非空 → 抛 `UNSUPPORTED_OOXML_FEATURE`,message 指向待落地的 issue
   *
   * 注意:事务原子性(#16 runTransaction)+ 索引同步 + trackChanges 包装,都在首个
   * executor(#28)落地时接入。本 PR 不引入 executor registry,避免为空抽象付代价;
   * #28 直接在这里改成 switch(op.type) 即可。
   */
  /**
   * 执行 Plan op 队列(docs/03.1 §3.2)。
   *
   * 流程:
   *   1. 预扫 affected parts(每个 op 触达的 part path)
   *   2. `runTransaction` 包裹 dispatch 循环 —— 任一 op throw → 全回滚
   *   3. 成功后重建三索引(paraIdx / bookmarkIdx / commentIdx)—— executor 可能
   *      增删元素,最简单的保真方式是 buildIndices 一次性重跑
   *
   * 返回 Promise 因 runTransaction 是 async;#27 的 sync 签名承诺在 RPC 层
   * (handler `await` + `transport.call` 自带 await)对调用方完全透明。
   */
  async applyPlan(
    params: ApplyPlanParams,
    opts: { source?: 'plan' | 'script' } = {},
  ): Promise<ApplyResult> {
    this.assertLoaded();
    if (params.ops.length === 0) return cloneApplyResult(EMPTY_APPLY_RESULT);

    // 进入互斥队列:等前一个 applyPlan 跑完(链上 promise 永不 reject),拿到
    // 自己的 release;真正的 inner 在 await prev 之后才跑。
    const prev = this._applyPlanChain;
    let release!: () => void;
    this._applyPlanChain = new Promise<void>((resolve) => {
      release = resolve;
    });
    await prev;
    try {
      return await this._applyPlanInner(params, opts);
    } finally {
      release();
    }
  }

  async validatePlan(params: ApplyPlanParams) {
    this.assertLoaded();
    await this._applyPlanChain;
    return validatePlanOps(params, this.indices!);
  }

  private async _applyPlanInner(
    params: ApplyPlanParams,
    opts: { source?: 'plan' | 'script' } = {},
  ): Promise<ApplyResult> {
    const options: ExecutorContext['options'] = {
      trackChanges: params.options.trackChanges ?? false,
      author: params.options.author ?? '',
      date: params.options.date ? new Date(params.options.date) : new Date(),
    };
    const kernelStartedAt = Date.now();
    const collectAffectedPartsStartedAt = Date.now();
    const collectAffected = collectAffectedPartPaths(this, params.ops);
    const affectedPaths = collectAffected.paths;
    const collectAffectedPartsMs = Date.now() - collectAffectedPartsStartedAt;
    const result = cloneApplyResult(EMPTY_APPLY_RESULT);
    const ctx: ExecutorContext = {
      pkg: this.pkg!,
      indices: this.indices!,
      options,
      result,
      opIndex: 0,
      docId: this.docId!,
      orphanMediaCandidates: new Set<string>(),
    };
    const elapsed = (ms: number): string => `${Math.round(ms / 100) / 10}s`;
    const opStats = new Map<string, { count: number; ms: number }>();
    const applyStartedAt = Date.now();
    let txElapsedMs = 0;
    let indicesElapsedMs = 0;

    let txCommitted = false;
    try {
      const txStartedAt = Date.now();
      await runTransaction(this.pkg!, affectedPaths, () => {
        for (let i = 0; i < params.ops.length; i++) {
          ctx.opIndex = i;
          const op = params.ops[i]!;
          const opStartedAt = Date.now();
          dispatchOp(op, ctx);
          const stat = opStats.get(op.type) ?? { count: 0, ms: 0 };
          stat.count++;
          stat.ms += Date.now() - opStartedAt;
          opStats.set(op.type, stat);
        }
      });
      txElapsedMs = Date.now() - txStartedAt;
      txCommitted = true;
    } finally {
      // #190 孤儿 media 清理:仅在事务真正 commit 后才删 binary part。
      //
      // 设计动机:
      //   - 事务内删 binary part 可借助 transactions.ts 的 createBinaryPart 回滚路径
      //     恢复,但要求 collectAffectedPartPaths 提前 lookup 全部 drawing 的 a:blip
      //     reb 解析媒体路径,代价高且与 #190 的"延迟决策"语义不符。
      //   - 事务外删则简单:rels 已 commit,真正归零(全包 rels 都不再引用 media path)
      //     才删 binary;若事务回滚,rels 同步回到 pre-state,本步骤 short-circuit 不动。
      //   - 即便此清理本身抛错,主流程 result 已生效,降级为 warn。
      if (txCommitted && ctx.orphanMediaCandidates && ctx.orphanMediaCandidates.size > 0) {
        try {
          sweepOrphanMediaParts(this.pkg!, ctx.orphanMediaCandidates);
        } catch (e) {
          console.warn(
            `[kernel.applyPlan] orphan media sweep failed: ${e instanceof Error ? e.message : String(e)}`,
          );
        }
      }

      // 无论成功还是回滚,都重建三索引:
      //  - 成功:executor 可能增删元素,paraIdx / bookmarkIdx / commentIdx 条目
      //    的 Element 引用需同步
      //  - 回滚:runTransaction 用 deep-clone 快照恢复 xmlDom,原有索引条目指向
      //    **pre-clone 的 Element**(已游离),必须重建
      // 索引重建本身抛错(e.g. 回滚后重名书签 —— 不该发生但保险)时,原外层错误
      // 更重要:重建错误吞掉,走 warn(参照 snapshot-api reinitializeAfterReload)
      const indicesStartedAt = Date.now();
      try {
        this.indices = buildIndices(this.pkg!);
      } catch (e) {
        if (e instanceof BookmarkLoadError) {
          console.warn(
            `[kernel.applyPlan] indices rebuild after apply failed: ${e.code} ${e.message}`,
          );
        } else {
          console.warn(`[kernel.applyPlan] indices rebuild after apply failed: ${String(e)}`);
        }
      }
      indicesElapsedMs = Date.now() - indicesStartedAt;
      const opSummary = Array.from(opStats.entries())
        .map(([type, stat]) => `${type} x${stat.count} ${elapsed(stat.ms)}`)
        .join(', ');
      result.debugTiming = {
        ops: params.ops.length,
        affectedParts: affectedPaths.length,
        collectAffectedPartsMs,
        collectStoriesMs: collectAffected.timing.storiesMs,
        collectStoryLookups: collectAffected.timing.storyLookups,
        collectStoryLookupMs: collectAffected.timing.storyLookupMs,
        collectByType: collectAffected.timing.byType,
        transactionMs: txElapsedMs,
        indicesMs: indicesElapsedMs,
        totalMs: Date.now() - applyStartedAt,
        kernelMs: Date.now() - kernelStartedAt,
        byType: Array.from(opStats.entries()).map(([type, stat]) => ({
          type: type as Op['type'],
          count: stat.count,
          ms: stat.ms,
        })),
      };
      const collectSummary = collectAffected.timing.byType
        .map(
          (entry) =>
            `${entry.type} x${entry.count} ${elapsed(entry.ms)} story=${entry.storyLookups}/${elapsed(entry.storyLookupMs)}`,
        )
        .join(', ');
      console.info(
        `[kernel.applyPlan:timing] ops=${params.ops.length}, affectedParts=${affectedPaths.length}, collectAffectedParts=${elapsed(collectAffectedPartsMs)}, collectStories=${elapsed(collectAffected.timing.storiesMs)}, storyLookups=${collectAffected.timing.storyLookups}/${elapsed(collectAffected.timing.storyLookupMs)}, tx=${elapsed(txElapsedMs)}, indices=${elapsed(indicesElapsedMs)}, total=${elapsed(Date.now() - applyStartedAt)}, kernel=${elapsed(Date.now() - kernelStartedAt)}${opSummary ? `, ${opSummary}` : ''}${collectSummary ? `, collectByType: ${collectSummary}` : ''}`,
      );
    }

    // applyPlan 走到这里 = transaction 已 commit;事件在索引重建之后派发,确保
    // 订阅者拿到的 kernel 状态与 result 一致。kernel 自己不知道调用来源,源由
    // opts.source 显式标注:RPC handler 传 'plan',script-execute 经包装后传 'script'。
    this.domainCache.afterSuccessfulApply(params.ops);
    this.emitChange({
      ops: params.ops,
      result,
      source: opts.source ?? 'script',
    });

    return result;
  }

  // ─────────── change 事件订阅(`WordDocument.on('change', ...)` 走 transport 接此处)───────────

  /**
   * 订阅 applyPlan 成功事件。返回取消订阅函数。
   *
   * 监听器同步触发(在 applyPlan 内部 fan-out),抛错被吞掉(不影响其它订阅者
   * 与 applyPlan 返回)。要做 async 工作请在监听器内自行 catch 并 fire-and-forget。
   */
  onChange(listener: (ev: KernelChangeEvent) => void): () => void {
    this.changeListeners.add(listener);
    return () => {
      this.changeListeners.delete(listener);
    };
  }

  private emitChange(ev: KernelChangeEvent): void {
    for (const listener of this.changeListeners) {
      try {
        listener(ev);
      } catch {
        // 防御性:监听器异常不影响 applyPlan 的返回路径
      }
    }
  }

  // ─────────── 快照原语(供 SnapshotApi 使用,#21)───────────

  /**
   * 序列化当前 pkg 为 snapshot 用的 parts 数组。
   * 优化:
   *   - dirty=false 且有 _originalBytes → 直接复用原始字节(零 serialize,Uint8Array 引用共享)
   *   - dirty=true 的 XML part → slimdom serialize
   *   - 二进制 part → 用 bytes / _originalBytes
   */
  serializeParts(): SnapshotPartEntry[] {
    this.assertLoaded();
    const out: SnapshotPartEntry[] = [];
    for (const part of listParts(this.pkg!)) {
      out.push({
        path: part.path,
        contentType: part.contentType,
        bytes: partToBytes(part),
      });
    }
    return out;
  }

  /**
   * 从 snapshot 的 parts 数组重建 pkg + 三索引。
   *
   * 覆盖原 pkg:调用方应在 rollback 前确保 dispose/ 清理 transient 状态(如 changelog)。
   * 重建后 ensureParaIds + buildIndices 再跑一次,保证索引准确(docs/03.2 §10)。
   *
   * onProgress 分阶段上报(docs/03.2 §6):
   *   - `'parse'`        每个 part 被 parse 完成时,pct 0→100
   *   - `'build-indices'` buildIndices 开始前,pct 0
   *   - `'done'`         全部完成,pct 100
   */
  rebuildFromParts(parts: readonly SnapshotPartEntry[], onProgress?: KernelProgress): void {
    this.assertNotDisposed();
    const pkg = createPackage();
    const total = parts.length;
    let processed = 0;
    // 即便 parts 为空,也先触发一次 parse 0%,方便调用方感知开始
    onProgress?.('parse', 0);
    for (const entry of parts) {
      if (isXmlContentType(entry.contentType)) {
        const text = new TextDecoder().decode(entry.bytes);
        const dom = parseXmlDocument(text) as unknown as Document;
        setPart(
          pkg,
          createXmlPart({
            path: entry.path,
            contentType: entry.contentType,
            xmlDom: dom,
            originalBytes: entry.bytes,
          }),
        );
      } else {
        setPart(
          pkg,
          createBinaryPart({
            path: entry.path,
            contentType: entry.contentType,
            bytes: entry.bytes,
          }),
        );
      }
      processed++;
      const pct = total > 0 ? Math.round((processed / total) * 100) : 100;
      onProgress?.('parse', pct);
    }

    // 与 load 相同的 normalize 流程
    ensureParaIds(pkg);
    onProgress?.('build-indices', 0);
    try {
      this.indices = buildIndices(pkg);
    } catch (e) {
      if (e instanceof BookmarkLoadError) {
        throw new WordError(e.code, e.message, this.docId ?? '', e.details);
      }
      throw e;
    }
    this.pkg = pkg;
    this.domainCache.reset();
    onProgress?.('done', 100);
  }

  // ─────────── 释放 ───────────

  dispose(): void {
    this.disposed = true;
    // 丢引用,便于 GC(slimdom Document 比较占内存)
    this.pkg = undefined;
    this.indices = undefined;
    this.initialBytes = undefined;
    this.initialFlatOpc = undefined;
    this.domainCache.reset();
  }

  // ─────────── 内部访问 ───────────

  /** @internal */
  getDocId(): string | undefined {
    return this.docId;
  }

  /** @internal */
  getPackage(): Package | undefined {
    return this.pkg;
  }

  /** @internal */
  getIndices(): PackageIndices | undefined {
    return this.indices;
  }

  /** @internal */
  getInitialBytes(): Uint8Array | undefined {
    return this.initialBytes;
  }

  /** @internal */
  getInitialFlatOpc(): string | undefined {
    return this.initialFlatOpc;
  }

  /** @internal */
  isDisposed(): boolean {
    return this.disposed;
  }

  /** @internal */
  debugCacheSnapshot(): WordDomainCacheSnapshot {
    return this.domainCache.snapshot();
  }

  /** @internal */
  debugInvalidateCache(domain: DomainCacheName = 'all', reason = 'debug'): void {
    this.domainCache.invalidate(domain, reason);
  }

  private assertLoaded(): void {
    this.assertNotDisposed();
    if (!this.pkg || !this.docId) {
      throw new WordError('PACKAGE_UNSUPPORTED_FORMAT', 'Kernel not loaded', this.docId ?? '');
    }
  }

  private assertNotDisposed(): void {
    if (this.disposed) {
      throw new WordError(
        'PACKAGE_UNSUPPORTED_FORMAT',
        'Kernel has been disposed',
        this.docId ?? '',
      );
    }
  }
}

/**
 * 预扫 ops,收集事务要监护的所有 part paths(每个 op 读 + 写可能触达的 part)。
 *
 * #28 executor 只改 paraId 所在 part;future executors(#29 / #34 / #35)会扩展
 * 此函数的分支(e.g. addComment 触达 body part + comments.xml + Content_Types.xml)。
 */
function collectAffectedPartPaths(
  kernel: Kernel,
  ops: readonly Op[],
): CollectAffectedPartPathsResult {
  const paths = new Set<string>();
  const indices = kernel.getIndices();
  const pkg = kernel.getPackage();
  const emptyTiming: CollectAffectedPartPathTiming = {
    storiesMs: 0,
    storyLookups: 0,
    storyLookupMs: 0,
    byType: [],
  };
  if (!indices || !pkg) return { paths: [], timing: emptyTiming };

  const storiesStartedAt = Date.now();
  const stories = listStories(pkg);
  const timing: CollectAffectedPartPathTiming = {
    storiesMs: Date.now() - storiesStartedAt,
    storyLookups: 0,
    storyLookupMs: 0,
    byType: [],
  };
  const byType = new Map<
    Op['type'],
    { type: Op['type']; count: number; ms: number; storyLookups: number; storyLookupMs: number }
  >();
  let currentStat: {
    type: Op['type'];
    count: number;
    ms: number;
    storyLookups: number;
    storyLookupMs: number;
  } | null = null;
  const findStoryPartPathFor = (el: Element): string | null => {
    const startedAt = Date.now();
    let found: Story | null = null;
    for (const story of stories) {
      if (story.root.contains(el as unknown as Node)) {
        found = story;
        break;
      }
    }
    const ms = Date.now() - startedAt;
    timing.storyLookups++;
    timing.storyLookupMs += ms;
    if (currentStat) {
      currentStat.storyLookups++;
      currentStat.storyLookupMs += ms;
    }
    return found?.partPath ?? null;
  };
  // R1 multi-image fix: track maxN across all insertImageAfter ops in this plan
  // so each op's projected media path is distinct (mirrors allocateMediaPath's incremental behavior).
  // Initialized lazily inside the case 'insertImageAfter' branch on first encounter.
  let insertImageMaxN: number | null = null;
  for (const op of ops) {
    const stat = byType.get(op.type) ?? {
      type: op.type,
      count: 0,
      ms: 0,
      storyLookups: 0,
      storyLookupMs: 0,
    };
    stat.count++;
    byType.set(op.type, stat);
    currentStat = stat;
    const opStartedAt = Date.now();
    try {
      switch (op.type) {
        case 'replaceParagraph':
        case 'replaceTextInParagraph':
        case 'formatTextInParagraph':
        case 'formatParagraph':
        case 'insertParagraphBefore':
        case 'insertParagraphAfter': {
          const paraId =
            op.type === 'insertParagraphBefore' || op.type === 'insertParagraphAfter'
              ? op.anchorParaId
              : op.paraId;
          const p = indices.paraIdx.get(paraId.toUpperCase());
          if (p) {
            const partPath = findStoryPartPathFor(p);
            if (partPath) paths.add(partPath);
          }
          break;
        }
        case 'insertParagraphAtBodyStart': {
          // 无锚点 op：直接纳入 body part（<w:document>/<w:body> 所在 part）。
          const [bodyStory] = listStories(pkg, { kinds: ['body'] });
          if (bodyStory) paths.add(bodyStory.partPath);
          break;
        }
        case 'deleteParagraph': {
          const p = indices.paraIdx.get(op.paraId.toUpperCase());
          if (p) {
            const partPath = findStoryPartPathFor(p);
            if (partPath) paths.add(partPath);
          }
          // drop 模式可能删 comments.xml 里的 <w:comment> 主体，并在删光时
          // cleanup CT/rels —— 保守地无条件纳入事务范围，保证原子回滚。
          if (op.onAnchorConflict === 'drop') {
            paths.add('word/comments.xml');
            paths.add('[Content_Types].xml');
            paths.add('word/_rels/document.xml.rels');
          }
          break;
        }
        // #34 书签:createBookmark 在 paraId 所在 part 插元素;
        //         deleteBookmark 在 entry.partPath 所在 part 删元素
        case 'createBookmark': {
          const p = indices.paraIdx.get(op.range.paraId.toUpperCase());
          if (p) {
            const partPath = findStoryPartPathFor(p);
            if (partPath) paths.add(partPath);
          }
          break;
        }
        case 'deleteBookmark': {
          const entry = indices.bookmarkIdx.get(op.name);
          if (entry) paths.add(entry.partPath);
          break;
        }
        // #35 addComment:body part + comments.xml + Content_Types + rels(首次时可能改后两者)
        case 'addComment': {
          const p = indices.paraIdx.get(op.range.paraId.toUpperCase());
          if (p) {
            const partPath = findStoryPartPathFor(p);
            if (partPath) paths.add(partPath);
          }
          // comments.xml 无论是否存在,都纳入监护范围(首次创建时需要被回滚)
          paths.add('word/comments.xml');
          paths.add('[Content_Types].xml');
          paths.add('word/_rels/document.xml.rels');
          break;
        }
        // #35 deleteComment:锚点所在 body part + comments.xml + CT + rels
        // cleanup 可能在删光最后一条时修改 CT/rels，必须纳入事务范围保证原子回滚
        case 'deleteComment': {
          const entry = indices.commentIdx.get(op.commentId);
          if (entry) {
            const anchorEl = entry.anchors.start ?? entry.anchors.end ?? entry.anchors.ref;
            if (anchorEl) {
              const partPath = findStoryPartPathFor(anchorEl);
              if (partPath) paths.add(partPath);
            }
          }
          paths.add('word/comments.xml');
          paths.add('[Content_Types].xml');
          paths.add('word/_rels/document.xml.rels');
          break;
        }
        // #186 addTableComment:走 commentsPart + body part(+ 首次创建 comments.xml 时 CT/rels)
        case 'addTableComment': {
          const tbl = indices.tblIdx.get(op.tableId);
          if (tbl) {
            const partPath = findStoryPartPathFor(tbl);
            if (partPath) paths.add(partPath);
          }
          paths.add('word/comments.xml');
          paths.add('[Content_Types].xml');
          paths.add('word/_rels/document.xml.rels');
          break;
        }
        // #186 deleteTable:表格所在 part + drop 模式可能动 comments.xml/CT/rels
        case 'deleteTable': {
          const tbl = indices.tblIdx.get(op.tableId);
          if (tbl) {
            const partPath = findStoryPartPathFor(tbl);
            if (partPath) paths.add(partPath);
          }
          if (op.onAnchorConflict === 'drop') {
            paths.add('word/comments.xml');
            paths.add('[Content_Types].xml');
            paths.add('word/_rels/document.xml.rels');
          }
          break;
        }
        // #188 deleteTableRow:行所在 part + drop 模式可能动 comments.xml/CT/rels
        case 'deleteTableRow': {
          const tbl = indices.tblIdx.get(op.tableId);
          if (tbl) {
            const partPath = findStoryPartPathFor(tbl);
            if (partPath) paths.add(partPath);
          }
          if (op.onAnchorConflict === 'drop') {
            paths.add('word/comments.xml');
            paths.add('[Content_Types].xml');
            paths.add('word/_rels/document.xml.rels');
          }
          break;
        }
        // #190 drawing 替换 / 删除 + #210 shape 内文字替换
        //  - drawing 所在 part(从 drawingIdx + findStoryFor 解析)
        //  - 该 part 的 .rels 同伴(改 / 删 image relationship / hyperlink relationship)
        //  - word/media/placeholder.png(replace 时首次注入)
        //  - [Content_Types].xml(首次需要登记 image/png Default 扩展)
        //  - 旧 media binary part 的清理在 post-transaction 阶段(transactions.ts 不支持
        //    回滚 pre-existing part 的删除;orphan 清理放在 commit 后)
        case 'formatTableCell': {
          const tbl = indices.tblIdx.get(op.tableId);
          if (tbl) {
            const partPath = findStoryPartPathFor(tbl);
            if (partPath) paths.add(partPath);
          }
          break;
        }
        case 'insertTableCellParagraph': {
          const tbl = indices.tblIdx.get(op.tableId);
          if (tbl) {
            const partPath = findStoryPartPathFor(tbl);
            if (partPath) paths.add(partPath);
          }
          break;
        }
        case 'replaceDrawingWithPlaceholder':
        case 'deleteDrawing':
        case 'replaceShapeInnerText': {
          const dr = indices.drawingIdx.get(op.drawingId);
          if (dr) {
            const partPath = findStoryPartPathFor(dr);
            if (partPath) {
              paths.add(partPath);
              paths.add(relsPathFor(partPath));
            }
          }
          if (op.type === 'replaceDrawingWithPlaceholder') {
            paths.add('word/media/placeholder.png');
            paths.add('[Content_Types].xml');
          }
          break;
        }
        // task-insert-image：insertImageAfter
        //  - anchor 所在 part（通常 word/document.xml）
        //  - 该 part 的 .rels（新增 image rel）
        //  - [Content_Types].xml（登记 ext Default）
        //  - 动态新建的 word/media/imageN.ext（R1：必须列出才能回滚删除新 part）
        //
        //  注意：多个 insertImageAfter ops 在 apply 前一次性计算 affected paths；
        //  必须模拟 executor 的逐 op 累加（而非每次都扫初始 pkg），否则多张图同 ext
        //  时会算出相同的 imageN+1，漏列 image2+ 导致回滚不完整。
        //  用 insertImagePlannedPaths（在 collectAffectedPartPaths 调用方维护的闭包变量）
        //  按 ext 分别跟踪"本 plan 已规划的序号 offset"。
        case 'insertImageAfter': {
          const anchorEl = indices.paraIdx.get(op.anchorParaId.toUpperCase());
          if (anchorEl) {
            const partPath = findStoryPartPathFor(anchorEl);
            if (partPath) {
              paths.add(partPath);
              paths.add(relsPathFor(partPath));
            }
          }
          paths.add('[Content_Types].xml');
          {
            const ext = contentTypeToExt(op.contentType);
            if (ext) {
              // 初始 pkg 里最大 N（跨所有 ext 共享，与 allocateMediaPath 对齐）
              if (!insertImageMaxN) {
                insertImageMaxN = 0;
                for (const part of listParts(pkg)) {
                  const m = /^word\/media\/image(\d+)\./i.exec(part.path);
                  if (m) {
                    const n = parseInt(m[1]!, 10);
                    if (n > insertImageMaxN) insertImageMaxN = n;
                  }
                }
              }
              // 每规划一张图，序号 +1（模拟 executor allocateMediaPath 的 maxN++ 效果）
              insertImageMaxN++;
              // hasPart 防撞：若 image{maxN}.ext 已在初始 pkg 中存在，继续 +1
              while (pkg.parts.has(`word/media/image${insertImageMaxN}.${ext}`)) {
                insertImageMaxN++;
              }
              paths.add(`word/media/image${insertImageMaxN}.${ext}`);
            }
          }
          break;
        }
        // #32 acceptRevisions / rejectRevisions
        case 'acceptRevisions':
        case 'rejectRevisions': {
          if (op.scope.scope === 'paragraph') {
            const p = indices.paraIdx.get(op.scope.paraId.toUpperCase());
            if (p) {
              const partPath = findStoryPartPathFor(p);
              if (partPath) paths.add(partPath);
            }
          } else {
            // all / revision:保守把所有 XML Word 内容 part 都纳入(包括 comments.xml
            // —— 因为批注 part 里可能也有 ins/del)。未用到的 parts 入事务无害。
            for (const part of listParts(pkg)) {
              if (isXmlPart(part) && part.xmlDom) {
                const ct = part.contentType;
                if (
                  ct.includes('wordprocessingml') ||
                  ct.endsWith('.main+xml') ||
                  ct.endsWith('.header+xml') ||
                  ct.endsWith('.footer+xml') ||
                  ct.endsWith('.footnotes+xml') ||
                  ct.endsWith('.endnotes+xml') ||
                  ct.endsWith('.comments+xml')
                ) {
                  paths.add(part.path);
                }
              }
            }
          }
          break;
        }
        default:
          break;
      }
    } finally {
      stat.ms += Date.now() - opStartedAt;
      currentStat = null;
    }
  }
  timing.byType = Array.from(byType.values());
  return { paths: [...paths], timing };
}

/**
 * 计算某 part 的 .rels 同伴路径。
 *
 * 规则:`<dir>/<name>.<ext>` → `<dir>/_rels/<name>.<ext>.rels`
 *   - 'word/document.xml'    → 'word/_rels/document.xml.rels'
 *   - 'word/header1.xml'     → 'word/_rels/header1.xml.rels'
 *   - 'word/footer1.xml'     → 'word/_rels/footer1.xml.rels'
 *
 * 顶层 part('xxx.xml' 无目录前缀)走 '_rels/xxx.xml.rels'(与 _rels/.rels 同位)。
 */
function relsPathFor(partPath: string): string {
  const idx = partPath.lastIndexOf('/');
  if (idx < 0) return `_rels/${partPath}.rels`;
  const dir = partPath.slice(0, idx);
  const name = partPath.slice(idx + 1);
  return `${dir}/_rels/${name}.rels`;
}

/** contentType → ext（与 insert-image-ops.ts 同步；用于 collectAffectedPartPaths R1）*/
function contentTypeToExt(ct: string): string | null {
  const map: Record<string, string> = {
    'image/png': 'png',
    'image/jpeg': 'jpeg',
    'image/gif': 'gif',
    'image/bmp': 'bmp',
    'image/tiff': 'tiff',
  };
  return map[ct] ?? null;
}

/** 浅拷贝 ApplyResult(避免把冻结引用交出去被调用方误写)*/
function cloneApplyResult(r: ApplyResult): ApplyResult {
  return {
    insertedParagraphs: r.insertedParagraphs.slice(),
    insertedBodyStartParaIds: r.insertedBodyStartParaIds.slice(),
    createdComments: r.createdComments.slice(),
    createdBookmarks: r.createdBookmarks.slice(),
    acceptedRevisions: r.acceptedRevisions,
    rejectedRevisions: r.rejectedRevisions,
    replacedTableCells: r.replacedTableCells,
    ...(r.debugTiming !== undefined ? { debugTiming: r.debugTiming } : {}),
  };
}

function wrapPackageError(err: unknown, docId: string, phase: string): WordError {
  if (err instanceof WordError) return err;
  const msg = err instanceof Error ? err.message : String(err);
  return new WordError('PACKAGE_CORRUPT', `${phase} failed: ${msg}`, docId, { phase });
}

/** Kernel.serializeParts / rebuildFromParts 的 entry 形态(与 SnapshotPayload.parts 一致) */
export interface SnapshotPartEntry {
  path: string;
  contentType: string;
  bytes: Uint8Array;
}

function isXmlContentType(ct: string): boolean {
  return ct.endsWith('+xml') || ct === 'application/xml' || ct === 'text/xml';
}

function partToBytes(part: Part): Uint8Array {
  // dirty=false + 原始字节:零 serialize
  if (!part._dirty && part._originalBytes) {
    return part._originalBytes;
  }
  if (isXmlPart(part)) {
    if (!part.xmlDom) {
      throw new Error(`partToBytes: XML part '${part.path}' missing xmlDom`);
    }
    const text = serializeToWellFormedString(part.xmlDom as never);
    return new TextEncoder().encode(text);
  }
  // 二进制
  if (part.bytes) return part.bytes;
  if (part._originalBytes) return part._originalBytes;
  throw new Error(`partToBytes: binary part '${part.path}' has no bytes`);
}
