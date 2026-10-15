/**
 * WordDocument —— Word core 的顶层句柄(docs/03.1 §2)。
 *
 * 本 PR(#17)把装载 / 输出 / dispose 的真实执行改走 `Transport` 抽象:
 *   - **inline 模式**:`InlineTransport` 直接在主线程创建 Kernel,零 RPC 开销
 *   - **worker 模式**:`WorkerTransport` 把调用序列化成 RPC 消息经 postMessage 发给
 *     Worker 内的 Kernel(Node worker_threads 路径在本 PR 就绪,浏览器 Worker 路径
 *     留 #18 / 浏览器示例)
 *
 * 对消费方完全透明:`WordDocument.fromZip(buf, { mode: 'worker' })` 与 inline 行为一致,
 * 只是 Kernel 跑在另一线程。`_getPackage` / `_getIndices` 在 worker 模式下会抛错(
 * Package 在另一线程,无法直接返回引用)。查询类 API(#24 起)将改走 RPC 方法。
 */
import { InMemorySnapshotStorage, type SnapshotStorageAdapter } from '../shared/index.js';
import {
  type ChangeEvent,
  type ChangeListener,
  type ChangeSubscribeOptions,
  type KernelChangeEvent,
} from './change-events.js';
import type { KernelLoadParams } from './kernel.js';
import { WordPlanImpl, type ApplyResult, type PlanOptions, type WordPlan } from './plan.js';
import {
  createBookmarkApi,
  createCommentApi,
  createContextApi,
  createDrawingsApi,
  createParagraphApi,
  createRevisionApi,
  createSectionsApi,
  createStylesApi,
  createTablesApi,
  createTocFieldsApi,
  type BookmarkApi,
  type CommentApi,
  type ContextApi,
  type DrawingsApi,
  type ParagraphApi,
  type ParagraphInsertOptions,
  type RevisionApi,
  type SectionsApi,
  type StylesApi,
  type TablesApi,
  type TocFieldsApi,
} from './sub-apis.js';
import { InlineTransport, type Transport } from './transport.js';
import { WorkerTransport, type WorkerPortLike } from './worker-transport.js';
import { WordError } from './word-error.js';
import type { Package } from '../pack/index.js';
import type { PackageIndices } from './indices.js';
import type { DomainCacheName, WordDomainCacheSnapshot } from './domain-cache.js';
import { SnapshotApiImpl, type SnapshotApi } from './snapshot-api.js';
import { createBlankDocxBytes } from './blank-document.js';

export type RuntimeMode = 'auto' | 'worker' | 'inline';

/**
 * 模式决策的"为什么"。配合 `mode` 让外部平台一眼看清落 inline 还是 worker、以及是哪条规则。
 *
 * - `explicit-inline` / `explicit-worker`:消费方在 LoadOptions.mode 显式指定
 * - `auto/inside-worker-thread`:已在 Node worker_threads 子线程内,拒绝再 spawn
 * - `auto/node-with-worker-config`:Node 主线程 + 配了 workerFactory / workerBundleUrl
 * - `auto/node-no-worker-config`:Node 主线程但 worker 基础设施未配,静默 fallback
 * - `auto/browser-factory`:浏览器侧消费方提供了 workerFactory
 * - `auto/browser-global`:浏览器侧粗检测到全局 Worker 构造器
 * - `auto/fallback`:都不命中的兜底
 */
export type ModeSource =
  | 'explicit-inline'
  | 'explicit-worker'
  | 'auto/inside-worker-thread'
  | 'auto/node-with-worker-config'
  | 'auto/node-no-worker-config'
  | 'auto/browser-factory'
  | 'auto/browser-global'
  | 'auto/fallback';

export interface ModeDecision {
  mode: 'worker' | 'inline';
  source: ModeSource;
}

/** onProgress 回调的 phase 枚举 */
export type LoadPhase = 'fetch' | 'parse' | 'normalize';

export interface LoadOptions {
  /**
   * 文档身份标识(必填)。宿主提供,典型值 = 文件路径 hash。
   * - 作为 WordError.docId 自动注入
   * - 作为快照系统命名空间 key(详见 docs/03.2)
   */
  docId: string;
  /**
   * 运行时模式(docs/03 §4.2):
   * - `'inline'`:主线程运行 kernel
   * - `'worker'`:显式要求 Worker;不可用时抛 WORKER_CREATION_FAILED
   * - `'auto'`(默认):按 §4.3 顺序检测;不可用时静默 fallback 到 inline
   */
  mode?: RuntimeMode;
  /**
   * 装载时模式决策回调(#TBD,docs/03 §4.3)。
   *
   * - 设了:用 `decision.mode` / `decision.source` 自定义打点,**不再走默认 console.info**
   * - 未设:走默认 `console.info('[coco-office:word] mode=... source=... docId=...')`
   *
   * 平台侧可借此把决策写入自有日志体系,或在 UI 上展示当前 SDK 的执行模式。
   */
  onModeResolved?: (decision: ModeDecision, docId: string) => void;
  onProgress?: (phase: LoadPhase, pct?: number) => void;
  snapshotStorage?: SnapshotStorageAdapter;
  /**
   * 快照 FIFO 上限,未设则用 `DEFAULT_MAX_SNAPSHOTS`(10,docs/03.2 §3.2)。
   * 必须 >= 1 的整数;测试场景可调低方便验证淘汰行为。
   */
  maxSnapshots?: number;
}

/** 宿主可以自定义 Worker 创建方式(测试 / 非标准环境)*/
export type WorkerFactory = () => WorkerPortLike | Promise<WorkerPortLike>;

/** 全局默认 */
interface GlobalDefaults {
  mode: RuntimeMode;
  author: string;
  /** 自定义 Worker 工厂;未设则尝试 Node worker_threads(见 tryCreateNodeWorker) */
  workerFactory?: WorkerFactory;
  /** Node worker_threads 的 bundle URL;workerFactory 未设时使用 */
  workerBundleUrl?: string | URL;
  /** 全局默认快照存储(未在 LoadOptions 指定时使用;默认 InMemorySnapshotStorage) */
  snapshotStorage?: SnapshotStorageAdapter;
}

const defaults: GlobalDefaults = {
  mode: 'auto',
  author: '',
};

export class WordDocument {
  readonly docId: string;
  readonly mode: 'worker' | 'inline';
  /** 决策来源(配合 `mode` 让外部平台看清是哪条规则把它送到这个模式的) */
  readonly modeSource: ModeSource;

  // ─────────── 读 namespace(#92,docs/03.1 §2)───────────
  readonly paragraphs: ParagraphApi;
  readonly bookmarks: BookmarkApi;
  readonly comments: CommentApi;
  readonly context: ContextApi;
  readonly revisions: RevisionApi;
  readonly styles: StylesApi;
  /** 表格枚举与结构分析（#185）*/
  readonly tables: TablesApi;
  /** Drawing 对象枚举（#189）*/
  readonly drawings: DrawingsApi;
  /** 节（section）枚举（#211）*/
  readonly sections: SectionsApi;
  /** TOC 域枚举（#197）*/
  readonly tocFields: TocFieldsApi;
  /** 快照子 API(docs/03.2);每个 WordDocument 独立实例 */
  readonly snapshots: SnapshotApi;

  private readonly transport: Transport;
  private readonly initialBytes?: Uint8Array;
  private readonly initialFlatOpc?: string;
  readonly snapshotStorage: SnapshotStorageAdapter;
  private disposed = false;
  /**
   * change 订阅者表。每个条目记 listener + 它的订阅选项。第一次有人订阅时绑
   * `transport.onChange`(由 `subscribeTransportLazy()` 控制),最后一个取消订阅时
   * 不解绑(简化实现;listener 集合空时 dispatch 是 no-op)。
   *
   * 派发顺序:每个 listener 各自串行(`tail` 链),不同 listener 并行。这样保证
   * 同一订阅者按变更发生顺序看到事件,即便其 handler 是 async + 上一次 toFlatOpc
   * 还没回。跨订阅者不串行,避免一个慢 handler 卡住别人。
   */
  private readonly changeSubscribers = new Set<{
    listener: ChangeListener;
    fullDocument: 'none' | 'flatOpc' | 'zip';
    /** 串行链尾,保证 listener 调用顺序与变更顺序一致 */
    tail: Promise<void>;
  }>();
  private changeTransportUnsubscribe?: () => void;

  private constructor(init: {
    transport: Transport;
    docId: string;
    mode: 'worker' | 'inline';
    modeSource: ModeSource;
    initialBytes?: Uint8Array;
    initialFlatOpc?: string;
    snapshotStorage: SnapshotStorageAdapter;
    maxSnapshots?: number;
  }) {
    this.transport = init.transport;
    this.docId = init.docId;
    this.mode = init.mode;
    this.modeSource = init.modeSource;
    this.initialBytes = init.initialBytes;
    this.initialFlatOpc = init.initialFlatOpc;
    this.snapshotStorage = init.snapshotStorage;
    this.snapshots = new SnapshotApiImpl({
      docId: init.docId,
      transport: init.transport,
      adapter: init.snapshotStorage,
      ...(init.maxSnapshots !== undefined ? { maxSnapshots: init.maxSnapshots } : {}),
    });
    const assertNotDisposed = (): void => this.assertNotDisposed();
    this.paragraphs = createParagraphApi(init.transport, assertNotDisposed);
    this.bookmarks = createBookmarkApi(init.transport, assertNotDisposed);
    this.comments = createCommentApi(init.transport, assertNotDisposed);
    this.context = createContextApi(init.transport, assertNotDisposed);
    this.revisions = createRevisionApi(init.transport, assertNotDisposed);
    this.styles = createStylesApi(init.transport, assertNotDisposed);
    this.tables = createTablesApi(init.transport, assertNotDisposed);
    this.drawings = createDrawingsApi(init.transport, assertNotDisposed);
    this.sections = createSectionsApi(init.transport, assertNotDisposed);
    this.tocFields = createTocFieldsApi(init.transport, assertNotDisposed);
  }

  // ─────────── 静态:装载 ───────────

  static async createBlank(opts: LoadOptions): Promise<WordDocument> {
    validateLoadOptions(opts);
    const bytes = await createBlankDocxBytes();
    return WordDocument.fromZip(bytes, opts);
  }

  static async fromZip(buf: Uint8Array | ArrayBuffer, opts: LoadOptions): Promise<WordDocument> {
    validateLoadOptions(opts);
    const decision = await resolveMode(opts.mode ?? defaults.mode, opts.docId);
    emitModeDecision(decision, opts.docId, opts.onModeResolved);
    const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const params: KernelLoadParams = {
      source: { type: 'zip', bytes },
      docId: opts.docId,
    };
    // Worker 模式 transport 同时接收 reload 参数(#19);inline 忽略
    const transport = await createTransport(decision.mode, opts.docId, params);

    try {
      opts.onProgress?.('parse', 0);
      // LoadPhase 是 KernelProgress phase 的子集,这里做一次类型桥接
      const onProgress = opts.onProgress
        ? (phase: string, pct?: number) => opts.onProgress?.(phase as LoadPhase, pct)
        : undefined;
      await transport.call('load', [params], onProgress);
    } catch (err) {
      await transport.dispose();
      throw err;
    }

    const doc = new WordDocument({
      transport,
      docId: opts.docId,
      mode: decision.mode,
      modeSource: decision.source,
      initialBytes: bytes,
      snapshotStorage: resolveSnapshotStorage(opts.snapshotStorage),
      ...(opts.maxSnapshots !== undefined ? { maxSnapshots: opts.maxSnapshots } : {}),
    });
    try {
      await doc.snapshots.initializeFromStorage();
    } catch (err) {
      await doc.dispose();
      throw err;
    }
    return doc;
  }

  static async fromFlatOpc(xml: string, opts: LoadOptions): Promise<WordDocument> {
    validateLoadOptions(opts);
    const decision = await resolveMode(opts.mode ?? defaults.mode, opts.docId);
    emitModeDecision(decision, opts.docId, opts.onModeResolved);
    const params: KernelLoadParams = {
      source: { type: 'flat-opc', xml },
      docId: opts.docId,
    };
    const transport = await createTransport(decision.mode, opts.docId, params);

    try {
      opts.onProgress?.('parse', 0);
      // LoadPhase 是 KernelProgress phase 的子集,这里做一次类型桥接
      const onProgress = opts.onProgress
        ? (phase: string, pct?: number) => opts.onProgress?.(phase as LoadPhase, pct)
        : undefined;
      await transport.call('load', [params], onProgress);
    } catch (err) {
      await transport.dispose();
      throw err;
    }

    const doc = new WordDocument({
      transport,
      docId: opts.docId,
      mode: decision.mode,
      modeSource: decision.source,
      initialFlatOpc: xml,
      snapshotStorage: resolveSnapshotStorage(opts.snapshotStorage),
      ...(opts.maxSnapshots !== undefined ? { maxSnapshots: opts.maxSnapshots } : {}),
    });
    try {
      await doc.snapshots.initializeFromStorage();
    } catch (err) {
      await doc.dispose();
      throw err;
    }
    return doc;
  }

  // ─────────── 实例:输出 ───────────

  async toZip(): Promise<Uint8Array> {
    this.assertNotDisposed();
    return this.transport.call<Uint8Array>('toZip', []);
  }

  async toFlatOpc(): Promise<string> {
    this.assertNotDisposed();
    return this.transport.call<string>('toFlatOpc', []);
  }

  async dispose(): Promise<void> {
    if (this.disposed) return;
    this.disposed = true;
    await this.transport.dispose();
  }

  /**
   * 在与本文档同上下文(worker / inline)中执行一段用户脚本。
   *
   * 04.2 §六 冻结的运行时环境:
   * - 脚本以 `new Function('doc', 'args', 'fs', body)(doc, args, fs)` 形式跑
   * - `doc` = worker 内 Kernel 实例;脚本可同步访问 `doc.listParagraphs()` 等
   * - `args` = structured-cloneable 入参
   * - `fs` v1 在 worker 内不传(主线程 plugin 已用过)
   * - 顶层允许 `await`;`return expression` 即结果
   *
   * **agent 层 `execute_code` / `execute_file` tool 通过本方法桥接到执行通道**(04.2 §五)。
   *
   * @param code  JS 源码;函数体形态
   * @param args  传给脚本的参数
   * @param timeoutMs 超时毫秒数,默认 30000;0 = 不超时
   * @param signal 可选 AbortSignal(#151):abort 触发 → worker 立即 terminate + reload;
   *               inline 模式下 race 拒绝(主线程立即返回,但脚本仍后台跑)
   * @throws  语法错误 / 脚本运行错 / 超时 / abort
   *
   * @see docs/04.2-脚本执行-Tool设计.md §二 §六
   */
  async executeCode(
    code: string,
    args?: Record<string, unknown>,
    timeoutMs: number = 30000,
    signal?: AbortSignal,
  ): Promise<unknown> {
    this.assertNotDisposed();
    return this.transport.executeScript(code, args, timeoutMs, signal);
  }

  // ─────────── 变更事件订阅 ───────────

  /**
   * 订阅文档变更事件。在 kernel.applyPlan 成功后触发,涵盖两条入口:
   *
   *   - `doc.plan().apply()`        →  ev.source === 'plan'
   *   - `doc.executeCode(script)`   →  ev.source === 'script'(脚本调 `doc.applyPlan`)
   *
   * **不**触发的场景:空 plan;applyPlan 抛错回滚;脚本绕过 kernel.applyPlan 直接
   * 改 `doc.getPackage()` 的 DOM(v1 不覆盖,见 change-events.ts 注释)。
   *
   * **fullDocument 选项**:默认 `'none'`(只发 `{ ops, result, source }`),不付输出
   * 代价。设为 `'flatOpc'` / `'zip'` 时,WordDocument 在派发前调一次对应方法把全文
   * 档塞进 payload —— **每次变更一次开销**,worker 模式下还要再 RPC + structured-clone,
   * 大文档可达几百 ms。多订阅者同次事件共享拉取,因此打开多个 'flatOpc' 订阅不会
   * 重复跑 toFlatOpc。
   *
   * **派发顺序**:同一订阅者按变更发生顺序串行调用 handler(即便 handler 是 async),
   * 不同订阅者并行 —— 一个慢 handler 不会卡住别人。
   *
   * @returns 取消订阅函数。重复调用安全(幂等)。
   */
  on(event: 'change', handler: ChangeListener, opts?: ChangeSubscribeOptions): () => void {
    this.assertNotDisposed();
    if (event !== 'change') {
      // 留一个泛化签名,后续若要扩 'snapshot' / 'reload' 等只需此处扩 switch
      throw new Error(`unknown event: ${String(event)}`);
    }
    const subscriber: {
      listener: ChangeListener;
      fullDocument: 'none' | 'flatOpc' | 'zip';
      tail: Promise<void>;
    } = {
      listener: handler,
      fullDocument: opts?.fullDocument ?? 'none',
      tail: Promise.resolve(),
    };
    this.changeSubscribers.add(subscriber);
    this.subscribeTransportLazy();
    let unsubscribed = false;
    return () => {
      if (unsubscribed) return;
      unsubscribed = true;
      this.changeSubscribers.delete(subscriber);
    };
  }

  private subscribeTransportLazy(): void {
    if (this.changeTransportUnsubscribe) return;
    this.changeTransportUnsubscribe = this.transport.onChange((kernelEv) => {
      this.dispatchChange(kernelEv);
    });
  }

  /**
   * 把 kernel 推上来的 KernelChangeEvent 转成对外 ChangeEvent,fan-out 给订阅者。
   * 优化:多订阅者要同一种全文档输出时只拉一次。每个订阅者自己串行(tail chain),
   * 跨订阅者并行。
   */
  private dispatchChange(kernelEv: KernelChangeEvent): void {
    if (this.changeSubscribers.size === 0) return;
    if (this.disposed) return;

    // 按需求扫一遍订阅者,决定要不要拉 flatOpc / zip。已 disposed 时不拉。
    let needFlatOpc = false;
    let needZip = false;
    for (const sub of this.changeSubscribers) {
      if (sub.fullDocument === 'flatOpc') needFlatOpc = true;
      else if (sub.fullDocument === 'zip') needZip = true;
    }

    // 起一次 fetch promise(同次事件共享);失败则塞 undefined,handler 自负判空。
    // 失败时打 warn —— 派发链路出问题不该静悄悄。
    const flatOpcPromise: Promise<string | undefined> = needFlatOpc
      ? this.toFlatOpc().catch((err) => {
          console.warn(
            `[word-document.on('change')] toFlatOpc failed (docId=${this.docId}): ${String(err)}`,
          );
          return undefined;
        })
      : Promise.resolve(undefined);
    const zipPromise: Promise<Uint8Array | undefined> = needZip
      ? this.toZip().catch((err) => {
          console.warn(
            `[word-document.on('change')] toZip failed (docId=${this.docId}): ${String(err)}`,
          );
          return undefined;
        })
      : Promise.resolve(undefined);

    for (const sub of this.changeSubscribers) {
      const next = sub.tail.then(async () => {
        // 串行 await —— 上一个 handler 完成后再拉下一个 ChangeEvent。fetch promise
        // 是事件级共享的,await 同一个 promise 不会重复拉。
        const ev: ChangeEvent = {
          ops: kernelEv.ops,
          result: kernelEv.result,
          source: kernelEv.source,
        };
        if (sub.fullDocument === 'flatOpc') {
          const xml = await flatOpcPromise;
          if (xml !== undefined) ev.flatOpc = xml;
        } else if (sub.fullDocument === 'zip') {
          const bytes = await zipPromise;
          if (bytes !== undefined) ev.zip = bytes;
        }
        try {
          await sub.listener(ev);
        } catch (err) {
          console.warn(
            `[word-document.on('change')] listener threw (docId=${this.docId}): ${String(err)}`,
          );
        }
      });
      // 即便 next reject 也要把 tail 推进(.catch 兜底),否则下一次 dispatch 会等到
      // 永远 pending 的链。理论上 .then 内部已 try/catch,不会 reject,这里再保险一次。
      sub.tail = next.catch(() => undefined);
    }
  }

  // ─────────── 能力 1:查询段落(#24)───────────

  // ─────────── 读 API(#92 起)───────────
  // 已迁移到 namespace 形态:`doc.paragraphs` / `doc.bookmarks` / `doc.comments` /
  // `doc.revisions`(见构造函数上面的 `readonly` 字段)。详见 docs/03.1 §2。

  // ─────────── 变更入口:Plan(#27)───────────

  /**
   * 创建新 WordPlan(docs/03.1 §2 / §3)。
   *
   * - 同步返回,不走 RPC;链式方法只把 op push 进队列,真正 RPC 在 `apply()` 触发
   * - 多 plan 可并存,各自独立状态(每次调用返回新实例)
   * - `opts.author` 缺省时走 `WordDocument.setDefaultAuthor` 配置的值;都缺则空串
   * - `opts.trackChanges` 默认 false
   * - `opts.date` 默认 `new Date()`(在 addComment 等 op 里实际需要时再取)
   *
   * **能力实现进度**(#27 仅提供容器与 apply 状态机;op executor 分批落地):
   *   - 段落修改 → #28
   *   - 段落插入 / 删除 → #29
   *   - 书签 → #34
   *   - 批注 → #35
   *   - 修订 accept/reject → #31 / #32
   *
   * 本 PR 期间,任何非空 plan 的 `apply()` 会抛 `UNSUPPORTED_OOXML_FEATURE`。
   */
  plan(opts?: PlanOptions): WordPlan {
    this.assertNotDisposed();
    const resolved: PlanOptions = {
      trackChanges: opts?.trackChanges ?? false,
    };
    const author = opts?.author ?? WordDocument._getDefaultAuthor();
    if (author) resolved.author = author;
    if (opts?.date !== undefined) resolved.date = opts.date;
    return new WordPlanImpl({
      docId: this.docId,
      transport: this.transport,
      options: resolved,
    });
  }

  // ─────────── 内部 accessor(inline 模式才可用;worker 模式抛错)───────────

  /**
   * Appends a plain-text paragraph at the end of the main document body.
   */
  async appendParagraph(text: string, opts: ParagraphInsertOptions = {}): Promise<ApplyResult> {
    this.assertNotDisposed();
    const { paragraphs } = await this.paragraphs.list({
      scope: 'body',
      pagination: { mode: 'all' },
    });
    const last = paragraphs[paragraphs.length - 1];
    if (last) return last.insertAfter(text, opts);

    const result = await this.plan(opts).insertParagraphAtBodyStart().apply();
    const newParaId = result.insertedBodyStartParaIds[0];
    if (!newParaId) return result;
    const paragraph = await this.paragraphs.get(newParaId);
    return paragraph ? paragraph.replaceWith(text, opts) : result;
  }

  /** @internal 仅 inline 模式可用(Package 跨 Worker 边界不可直接引用)*/
  _getPackage(): Package {
    this.assertNotDisposed();
    if (this.mode !== 'inline') {
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        '_getPackage() is only available in inline mode; use query RPC methods in worker mode',
        this.docId,
      );
    }
    const inline = this.transport as InlineTransport;
    const pkg = inline.kernel.getPackage();
    if (!pkg) throw new WordError('PACKAGE_UNSUPPORTED_FORMAT', 'kernel has no pkg', this.docId);
    return pkg;
  }

  /** @internal 仅 inline 模式可用 */
  _getIndices(): PackageIndices {
    this.assertNotDisposed();
    if (this.mode !== 'inline') {
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        '_getIndices() is only available in inline mode',
        this.docId,
      );
    }
    const inline = this.transport as InlineTransport;
    const idx = inline.kernel.getIndices();
    if (!idx)
      throw new WordError('PACKAGE_UNSUPPORTED_FORMAT', 'kernel has no indices', this.docId);
    return idx;
  }

  /** @internal inline-only cache diagnostics for core tests. */
  _debugCacheSnapshot(): WordDomainCacheSnapshot {
    this.assertNotDisposed();
    if (this.mode !== 'inline') {
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        '_debugCacheSnapshot() is only available in inline mode',
        this.docId,
      );
    }
    const inline = this.transport as InlineTransport;
    return inline.kernel.debugCacheSnapshot();
  }

  /** @internal inline-only cache invalidation for consistency tests. */
  _debugInvalidateCache(domain: DomainCacheName = 'all', reason = 'debug'): void {
    this.assertNotDisposed();
    if (this.mode !== 'inline') {
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        '_debugInvalidateCache() is only available in inline mode',
        this.docId,
      );
    }
    const inline = this.transport as InlineTransport;
    inline.kernel.debugInvalidateCache(domain, reason);
  }

  /** @internal 装载时的初始字节(#19 worker reload 需要)*/
  _getInitialBytes(): Uint8Array | undefined {
    return this.initialBytes;
  }

  /** @internal 装载时的初始 Flat OPC 字符串 */
  _getInitialFlatOpc(): string | undefined {
    return this.initialFlatOpc;
  }

  /** @internal #19 测试用:拿到底层 transport 实例(模拟 worker crash) */
  _getTransport(): Transport {
    return this.transport;
  }

  private assertNotDisposed(): void {
    if (this.disposed) {
      throw new WordError(
        'PACKAGE_UNSUPPORTED_FORMAT',
        'WordDocument has been disposed',
        this.docId,
      );
    }
  }

  // ─────────── 静态:运行时默认 ───────────

  static setDefaultMode(mode: RuntimeMode): void {
    defaults.mode = mode;
  }

  /**
   * 不创建实例,查当前环境 `'auto'` 模式会选什么。真实反映决策结果
   * (docs/03 §4.3):
   *
   *   1. 已在 Node worker_threads 上下文中 → inline(避免二次 spawn)
   *   2. Node 环境:
   *        - 配了 workerFactory / workerBundleUrl → worker
   *        - 否则 → inline
   *   3. 浏览器(或类浏览器)环境:
   *        - 配了 workerFactory → worker
   *        - 全局 `Worker` 存在 → worker(粗检测,真实 ping 见 #52)
   *        - 否则 → inline
   *   4. 以上都不满足 → inline
   */
  static detectMode(): 'worker' | 'inline' {
    return detectAutoDecision().mode;
  }

  /**
   * 同 `detectMode()`,但返回决策来源,便于平台诊断为什么 auto 落到了某个模式。
   */
  static detectModeDecision(): ModeDecision {
    return detectAutoDecision();
  }

  static setDefaultAuthor(author: string): void {
    defaults.author = author;
  }

  /** 配置自定义 Worker 工厂(测试 / 非 Node-worker_threads 环境)*/
  static setWorkerFactory(factory: WorkerFactory | undefined): void {
    defaults.workerFactory = factory;
  }

  /** 配置 Node worker_threads 加载的 bundle URL(通常是 tsup 产物 dist/kernel-worker.mjs)*/
  static setWorkerBundleUrl(url: string | URL | undefined): void {
    defaults.workerBundleUrl = url;
  }

  /** 配置全局默认快照存储。未设 + LoadOptions 也未传 → `InMemorySnapshotStorage`(每次装载独立实例) */
  static setDefaultSnapshotStorage(adapter: SnapshotStorageAdapter | undefined): void {
    defaults.snapshotStorage = adapter;
  }

  /** @internal */
  static _getDefaultAuthor(): string {
    return defaults.author;
  }

  /** @internal */
  static _getDefaultMode(): RuntimeMode {
    return defaults.mode;
  }

  /** @internal */
  static _resetDefaults(): void {
    defaults.mode = 'auto';
    defaults.author = '';
    defaults.workerFactory = undefined;
    defaults.workerBundleUrl = undefined;
    defaults.snapshotStorage = undefined;
  }
}

// ─────────── 模块级辅助 ───────────

function validateLoadOptions(opts: LoadOptions): void {
  if (!opts || typeof opts !== 'object') {
    throw new WordError(
      'PACKAGE_UNSUPPORTED_FORMAT',
      'LoadOptions required (must contain docId)',
      '',
    );
  }
  if (typeof opts.docId !== 'string' || opts.docId.length === 0) {
    throw new WordError(
      'PACKAGE_UNSUPPORTED_FORMAT',
      'LoadOptions.docId is required and must be a non-empty string',
      '',
    );
  }
}

async function resolveMode(requested: RuntimeMode, docId: string): Promise<ModeDecision> {
  // 显式 inline:永远 inline,即便配了 bundle(docs/03 §4.2)
  if (requested === 'inline') return { mode: 'inline', source: 'explicit-inline' };

  // 显式 worker:必须可用,否则抛错(尊重显式意愿)
  if (requested === 'worker') {
    if (!canUseWorker()) {
      throw new WordError(
        'WORKER_CREATION_FAILED',
        "mode: 'worker' requested but no Worker runtime available; " +
          'in Node call WordDocument.setWorkerFactory() or setWorkerBundleUrl(), ' +
          'in browser provide a workerFactory (or ensure global Worker exists).',
        docId,
      );
    }
    return { mode: 'worker', source: 'explicit-worker' };
  }

  // 'auto':真实决策。不可用时静默 fallback 到 inline,绝不抛。
  return detectAutoDecision();
}

/**
 * 把模式决策派出去:有 onModeResolved 走它,没的话走默认 console.info。
 * 双重容错:onModeResolved 抛错也不影响装载主路径。
 */
function emitModeDecision(
  decision: ModeDecision,
  docId: string,
  onResolved?: (decision: ModeDecision, docId: string) => void,
): void {
  if (onResolved) {
    try {
      onResolved(decision, docId);
    } catch {
      // 宿主回调异常不能影响装载本身
    }
    return;
  }
  // 默认打点:平台 SDK 集成时,从控制台 / Node stdout 直接看到落地决策
  // eslint-disable-next-line no-console
  console.info(`[coco-office:word] mode=${decision.mode} source=${decision.source} docId=${docId}`);
}

/**
 * `'auto'` 模式的真实决策函数。按 docs/03 §4.3 顺序:
 *
 *   1. 已在 Node worker_threads 上下文中(非 main thread)→ inline
 *      避免在 worker 内又 spawn 子 worker 造成递归。
 *   2. 明显 Node(`process.versions.node` 存在且不是 worker_threads 内):
 *        - 配了 workerFactory / workerBundleUrl → worker
 *        - 否则 → inline(Node CLI 典型路径)
 *   3. 浏览器 / 类浏览器(无 process.versions.node):
 *        - 配了 workerFactory → worker(宿主自己管 ping)
 *        - 全局有 `Worker` 构造器 → worker(粗检测;真 ping 留 #52)
 *        - 否则 → inline
 *   4. 兜底 → inline
 *
 * 纯同步、无副作用,调用者可不 await。
 */
function detectAutoDecision(): ModeDecision {
  // 1) 已在 worker_threads 内:拒绝再开子 worker。
  if (isInsideNodeWorkerThread()) return { mode: 'inline', source: 'auto/inside-worker-thread' };

  const isNode = typeof process !== 'undefined' && process.versions?.node !== undefined;

  // 2) Node 主线程:仅在宿主提供 worker 基础设施时才走 worker。
  if (isNode) {
    return canUseWorker()
      ? { mode: 'worker', source: 'auto/node-with-worker-config' }
      : { mode: 'inline', source: 'auto/node-no-worker-config' };
  }

  // 3) 浏览器 / 类浏览器环境。
  if (defaults.workerFactory !== undefined) {
    return { mode: 'worker', source: 'auto/browser-factory' };
  }
  // 粗检测:全局有 Worker 构造器即视为可用。真 ping(加载 kernel + ready 握手)
  // 会引入时序复杂度,留给 #52 浏览器示例完善。
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  if (typeof (globalThis as any).Worker === 'function') {
    return { mode: 'worker', source: 'auto/browser-global' };
  }

  // 4) 兜底。
  return { mode: 'inline', source: 'auto/fallback' };
}

/** 是否跑在 Node worker_threads 子线程里。动态 require 避免浏览器打包报错。*/
function isInsideNodeWorkerThread(): boolean {
  if (typeof process === 'undefined' || !process.versions?.node) return false;
  try {
    // 用 Function 构造避免被静态分析工具/打包器识别为 require()
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const req = new Function('m', 'return require(m)') as (m: string) => any;
    const wt = req('node:worker_threads') as { isMainThread: boolean };
    return wt.isMainThread === false;
  } catch {
    return false;
  }
}

function canUseWorker(): boolean {
  // Node 场景:必须有 factory 或 bundleUrl。浏览器场景走 detectAutoChoice 里的 Worker 全局检测。
  if (defaults.workerFactory !== undefined || defaults.workerBundleUrl !== undefined) return true;
  // 浏览器场景无 factory 时,借全局 Worker 粗检测(和 detectAutoChoice 一致)。
  const isNode = typeof process !== 'undefined' && process.versions?.node !== undefined;
  if (!isNode) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (typeof (globalThis as any).Worker === 'function') return true;
  }
  return false;
}

function resolveSnapshotStorage(
  optsAdapter: SnapshotStorageAdapter | undefined,
): SnapshotStorageAdapter {
  // 优先级:LoadOptions.snapshotStorage > setDefaultSnapshotStorage > InMemorySnapshotStorage(本次装载独立)
  return optsAdapter ?? defaults.snapshotStorage ?? new InMemorySnapshotStorage();
}

async function createTransport(
  mode: 'worker' | 'inline',
  docId: string,
  loadParams?: KernelLoadParams,
): Promise<Transport> {
  if (mode === 'inline') {
    return new InlineTransport();
  }
  // mode === 'worker'
  try {
    // 闭包封装 "起一个 WorkerPortLike" 的逻辑,供首次使用 + reload 复用
    const makePort = async (): Promise<WorkerPortLike> => {
      if (defaults.workerFactory) {
        return await defaults.workerFactory();
      }
      if (defaults.workerBundleUrl) {
        const { createNodeWorkerPort } = await import('./worker-port-node.js');
        return createNodeWorkerPort(defaults.workerBundleUrl);
      }
      throw new Error('no worker factory or bundle URL configured');
    };
    const port = await makePort();
    return new WorkerTransport({
      port,
      // loadParams 存在时启用 reload(#19);fromZip/fromFlatOpc 总是传
      reload: loadParams ? { factory: makePort, loadParams, docId } : undefined,
    });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new WordError(
      'WORKER_CREATION_FAILED',
      `failed to create Worker transport: ${msg}`,
      docId,
    );
  }
}
