/**
 * SnapshotApi —— WordDocument.snapshots 子 API(docs/03.2 §2)。
 *
 * 职责:
 *   - `take(reason?, details?)`:通过 transport 让 kernel 序列化 parts,捆装 info + parts 为
 *     SnapshotPayload,交给 adapter.save,返回 id
 *   - `rollback(id, opts?)`:从 adapter.load 取回 payload,通过 transport 让 kernel 用 parts
 *     重建 pkg;消费语义 —— 成功后 adapter.delete;支持 `opts.onProgress` 分阶段上报
 *     (parse / build-indices / done,docs/03.2 §6)
 *   - `list / delete / clear`:直通 adapter
 *   - `recordChange / changes`:**主线程同步**,不过 transport(零 RPC 开销;docs/03.2 §7)
 *
 * 新增 #22:
 *   - FIFO 上限(默认 10,`SnapshotApiInit.maxSnapshots` 可配),take 成功后若超限则按
 *     `adapter.list` 的自增顺序淘汰最早的(`adapter.delete`)
 *   - `rollback` 的 `onProgress` 透传给 transport.call,InlineTransport 直达,
 *     WorkerTransport 通过 RpcProgress 穿过 Worker 边界
 */
import type {
  ChangeLog,
  SnapshotInfo,
  SnapshotPayload,
  SnapshotStorageAdapter,
} from '../shared/snapshot/index.js';
import type { KernelProgress, SnapshotPartEntry } from './kernel.js';
import type { Transport } from './transport.js';
import { WordError } from './word-error.js';
import { fingerprintSnapshotParts } from './snapshot-fingerprint.js';

/** 默认 FIFO 上限(docs/03.2 §3.2) */
export const DEFAULT_MAX_SNAPSHOTS = 10;

/** rollback 的进度回调:phase 分 'parse' | 'build-indices' | 'done'(docs/03.2 §6) */
export type RollbackProgress = (phase: string, pct?: number) => void;

export interface RollbackOptions {
  onProgress?: RollbackProgress;
}

/**
 * `take` 可选参数(#108,docs/03.2 §3.3)。
 *
 * - `pin: true` —— 标记该快照 **不参与 FIFO 淘汰**;即使达到 `maxSnapshots`,
 *   FIFO 也会跳过 pinned 的条目,转而淘汰最老的**非 pinned** 快照。
 * - `pin: undefined | false` —— 与不传 options 等价,普通快照。
 *
 * **使用约束**(agent 层自行约束,core 不强制):pinned 主要给 SDK 内部使用
 * (例 `WordAgent.create` 的 session-start 快照),plugin 层的 `AgentContext.snapshots`
 * 应强制 `pin: false`,避免 plugin 把存储塞满 pinned 导致 FIFO 失效。
 */
export interface TakeOptions {
  pin?: boolean;
  /**
   * 可选逻辑去重域。提供后，若该域最新快照的 contentHash 与当前文档相同，
   * `take` 直接返回已有 id，不创建新快照。不同 owner/用途必须使用不同 key。
   */
  dedupeKey?: string;
}

export interface SnapshotApi {
  /**
   * 保存当前文档状态为一个新快照,返回自增 id(snap-0 / snap-1 / ...)。
   * 已累加的 changelog 项会一次性吸入新快照的 `changesCaptured`,然后主线程 changelog 清零。
   *
   * 超过 `maxSnapshots`(默认 10,docs/03.2 §3.2)时,按 `adapter.list` 返回顺序淘汰
   * 最早的**非 pinned** 快照;pinned 快照永不淘汰(#108,docs/03.2 §3.3)。
   *
   * **能力边界**:若存储里的非 pinned 名额都满而仍要继续 take 普通快照,SDK 会尝试
   * 按正常 FIFO 淘汰最老非 pinned;若所有快照都是 pinned,则新 take 会直接成功
   * 保存(超出 maxSnapshots 但无可淘汰目标),SDK 发 `console.warn` 提示;
   * 上游自行控制 pinned 数量。
   */
  take(reason?: string, details?: unknown, options?: TakeOptions): Promise<string>;

  /**
   * 用给定 id 的快照覆盖当前文档状态,但**不删除**快照。
   *
   * 供需要把恢复结果提交到外部宿主的场景使用:宿主应先 restore,待外部写回成功后
   * 再显式 delete。这样宿主写回失败时仍可重试同一快照。
   */
  restore(id: string, opts?: RollbackOptions): Promise<void>;

  /**
   * 用给定 id 的快照覆盖当前文档状态,**消费式**:成功后该快照被 adapter.delete。
   * 其它快照不受影响。
   *
   * `opts.onProgress` 按阶段回调:'parse'(0→100)/ 'build-indices'(0)/ 'done'(100)。
   */
  rollback(id: string, opts?: RollbackOptions): Promise<void>;

  /** 按 id 自增序返回所有快照的 meta(不含 parts 字节) */
  list(): Promise<SnapshotInfo[]>;

  /** 按 id 删除某个快照。不存在时静默(幂等) */
  delete(id: string): Promise<void>;

  /** 清空该 docId 下的所有快照 */
  clear(): Promise<void>;

  /** 主线程同步累加一条 change,下次 take 时会一并吸入 */
  recordChange(change: { toolName: string; summary?: unknown }): void;

  /** 读当前未吸入的 change 数组(主线程同步) */
  changes(): ChangeLog[];

  /**
   * **内部用**:WordDocument 首次装载时从 adapter 恢复 nextId。失败会向上抛，
   * 防止在无法读取历史时从 snap-0 开始覆盖持久数据。
   */
  initializeFromStorage(): Promise<void>;

  /**
   * **内部/测试用**:worker reload 完成后由 transport 自动触发,主线程不应直接调。
   *
   * 语义(#23,docs/03.2 §9.1):
   *   - 从 adapter.list(docId) 读已持久化快照,把 `nextId` 重设为 `max(snap-<N>) + 1`
   *   - `changelog` 清零(reload 期间主线程的 changelog 视为丢失,无法恢复)
   *   - adapter 里已保存的字节不受影响,rollback 仍可用
   */
  reinitializeAfterReload(): Promise<void>;
}

export interface SnapshotApiInit {
  docId: string;
  transport: Transport;
  adapter: SnapshotStorageAdapter;
  /**
   * FIFO 上限,可选;默认 {@link DEFAULT_MAX_SNAPSHOTS}(10)。
   * 测试场景可调低便于验证淘汰行为。必须 >= 1。
   */
  maxSnapshots?: number;
}

export class SnapshotApiImpl implements SnapshotApi {
  private readonly docId: string;
  private readonly transport: Transport;
  private readonly adapter: SnapshotStorageAdapter;
  private readonly maxSnapshots: number;
  private nextId = 0;
  private changelog: ChangeLog[] = [];

  constructor(init: SnapshotApiInit) {
    this.docId = init.docId;
    this.transport = init.transport;
    this.adapter = init.adapter;
    const m = init.maxSnapshots ?? DEFAULT_MAX_SNAPSHOTS;
    if (!Number.isInteger(m) || m < 1) {
      throw new WordError(
        'PACKAGE_UNSUPPORTED_FORMAT',
        `SnapshotApiInit.maxSnapshots must be a positive integer, got: ${m}`,
        init.docId,
      );
    }
    this.maxSnapshots = m;
    // #23:订阅 transport reload 事件,worker crash + reload 后自动重建 nextId。
    // InlineTransport 无 onReload → 订阅 skip,符合预期(inline 无 reload 语义)。
    // transport 与 SnapshotApi 同生命周期(WordDocument 持有),dispose 链路由
    // WordDocument.dispose → transport.dispose 统一处理,这里不保存 unsubscribe。
    this.transport.onReload?.(() => this.reinitializeAfterReload());
  }

  async take(reason?: string, details?: unknown, options?: TakeOptions): Promise<string> {
    const pinned = options?.pin === true;
    const dedupeKey = options?.dedupeKey;
    if (dedupeKey !== undefined && dedupeKey.length === 0) {
      throw new WordError(
        'PACKAGE_UNSUPPORTED_FORMAT',
        'Snapshot TakeOptions.dedupeKey must not be empty',
        this.docId,
      );
    }

    // 1. 先序列化 parts(inline 直接取、worker 走 RPC)。
    //    **id 分配放在 await 之后**:worker reload 语义下,transport.call 可能被排队
    //    等待 reload + reinitializeAfterReload 完成;id 此时计算才能读到重建后的
    //    nextId,避免与 adapter 里已有 id 撞车(#23,docs/03.2 §9.1)。
    const parts = await this.transport.call<SnapshotPartEntry[]>('serializeParts', []);

    // 对规范排序后的整篇 OOXML parts 计算稳定指纹。只有显式提供 dedupeKey 的调用
    // 才启用逻辑去重，避免把 session-start / user / plugin 等不同所有权节点合并。
    const contentHash = fingerprintSnapshotParts(parts);
    if (dedupeKey !== undefined) {
      const infos = await this.adapter.list(this.docId);
      for (let i = infos.length - 1; i >= 0; i--) {
        const latestInScope = infos[i]!;
        if (latestInScope.dedupeKey !== dedupeKey) continue;
        if (latestInScope.contentHash === contentHash) {
          // 本次 take 已形成逻辑检查点，只是复用了已有节点；净零变化日志不应泄漏
          // 到下一张真正的新快照。
          this.changelog = [];
          return latestInScope.id;
        }
        break;
      }
    }

    const id = `snap-${this.nextId++}`;

    // 2. 吸入 changelog(主线程侧,take 结束前收集完毕)
    const changesCaptured = this.changelog;
    this.changelog = [];

    // 3. 组装 payload;pinned 一并写进 info,方便旧 adapter(不看 meta)也能带出
    const info: SnapshotInfo = {
      id,
      timestamp: Date.now(),
      reason,
      details,
      contentHash,
      ...(dedupeKey !== undefined ? { dedupeKey } : {}),
      changesCaptured,
      ...(pinned ? { pinned: true } : {}),
    };
    const payload: SnapshotPayload = {
      info,
      parts,
    };

    // 4. 持久化;pinned 通过 adapter 的 meta 通路传递(对不支持 pinned 的旧
    //    adapter,pinned 已写在 payload.info 上,list 时仍能还原)
    try {
      await this.adapter.save(this.docId, id, payload, pinned ? { pinned: true } : undefined);
    } catch (e) {
      // save 失败不算快照拍成功;回滚 id / changelog 以便下一次重试
      this.nextId--;
      this.changelog = changesCaptured.concat(this.changelog);
      throw e;
    }

    // 5. FIFO 淘汰(docs/03.2 §3.2 + §3.3):超出 maxSnapshots 时按 list 顺序
    //    删最老的**非 pinned** 条目。如果所有快照均为 pinned,则无可淘汰目标,
    //    console.warn 后放行(不抛错;保证 take 本身的成功语义)。
    try {
      await this.evictIfNeeded();
    } catch (err) {
      // save 已成功,此时 take 的事实已经成立。淘汰属于容量维护,失败不应把一个
      // 已存在的快照伪装成“拍摄失败”;保留超限状态并告警,后续 take 再重试淘汰。
      console.warn(
        `[snapshots] eviction failed after saving docId=${this.docId} id=${id}; snapshot kept`,
        err,
      );
    }

    return id;
  }

  async restore(id: string, opts?: RollbackOptions): Promise<void> {
    const payload = await this.adapter.load(this.docId, id);
    try {
      // onProgress 透传:inline 直达、worker 走 RpcProgress;KernelProgress 与 RollbackProgress
      // 形态一致,这里直接做一次轻量桥接(保留 opts 存在性判定,避免创建无用闭包)
      const onProgress: KernelProgress | undefined = opts?.onProgress
        ? (phase, pct) => opts.onProgress?.(phase, pct)
        : undefined;
      await this.transport.call('rebuildFromParts', [payload.parts], onProgress);
    } catch (e) {
      // 重建失败:保留 snapshot 以便排查。rollback 不算成功。
      throw toWordError(e, this.docId);
    }
  }

  async rollback(id: string, opts?: RollbackOptions): Promise<void> {
    await this.restore(id, opts);
    // 兼容既有消费语义:restore 成功后删除该快照。需要跨宿主原子提交的调用方
    // 应使用 restore → 宿主写回 → delete 的两阶段流程。
    await this.adapter.delete(this.docId, id);
  }

  /**
   * FIFO 淘汰:超限则按 adapter.list 顺序淘汰最早的**非 pinned** 条目。
   *
   * - 依赖 adapter.list 返回最早→最新;InMemory 按 snap-<N> 数字升序,契约中文档化
   * - pinned 快照永不被淘汰(#108,docs/03.2 §3.3)
   * - 可能一次触发多条(容错 maxSnapshots 在运行时被外部缩小)
   * - 若超限但全部都是 pinned,console.warn 提示后放行,不抛错
   */
  private async evictIfNeeded(): Promise<void> {
    const all = await this.adapter.list(this.docId);
    if (all.length <= this.maxSnapshots) return;

    // 目标:让**非 pinned 条目数**回到 (maxSnapshots - pinnedCount) 以内。
    // 更精准:淘汰到 all.length <= maxSnapshots,但只能动非 pinned。
    // 实现:按顺序扫,遇非 pinned 即删,直到 all.length <= maxSnapshots 或
    // 扫完仍无可淘汰。
    let remaining = all.length;
    let evicted = 0;
    for (const entry of all) {
      if (remaining <= this.maxSnapshots) break;
      if (entry.pinned === true) continue; // pinned 跳过
      await this.adapter.delete(this.docId, entry.id);
      remaining--;
      evicted++;
    }
    if (remaining > this.maxSnapshots && evicted === 0) {
      // 全是 pinned,无可淘汰;不抛错(保证 take 成功),只发警告
      console.warn(
        `[snapshots] docId=${this.docId} exceeds maxSnapshots=${this.maxSnapshots} but all ${all.length} snapshots are pinned; skipping eviction`,
      );
    }
  }

  async list(): Promise<SnapshotInfo[]> {
    return this.adapter.list(this.docId);
  }

  async delete(id: string): Promise<void> {
    return this.adapter.delete(this.docId, id);
  }

  async clear(): Promise<void> {
    return this.adapter.clear(this.docId);
  }

  recordChange(change: { toolName: string; summary?: unknown }): void {
    this.changelog.push({
      timestamp: Date.now(),
      toolName: change.toolName,
      ...(change.summary !== undefined ? { summary: change.summary } : {}),
    });
  }

  changes(): ChangeLog[] {
    // 返回浅拷贝,防止外部意外修改内部状态
    return this.changelog.slice();
  }

  /**
   * Worker reload 后内存态重建(#23,docs/03.2 §9.1)。
   *
   * 从 adapter 重新读出持久化快照列表,把 `nextId` 重设到 `max(snap-<N>) + 1`,
   * 保证后续 take 不与已有 id 撞车。`changelog` 清零——主线程在 reload 窗口累加
   * 的 ChangeLog 视为丢失,与 worker 侧 DOM 状态一起被重放覆盖,无法恢复。
   *
   * adapter.list 失败(极少见)时吞掉错误,保持现有 nextId 不变;下次 take
   * 仍可能冲突,但此处的抛错会阻塞 reload flush,不划算。
   */
  async reinitializeAfterReload(): Promise<void> {
    try {
      await this.initializeFromStorage();
    } catch (err) {
      // adapter.list 故障不阻塞 reload;nextId 保持旧值,下次 take 仍可能冲突,
      // 但抛错会阻塞 reload flush —— 此处选择不阻塞,让宿主自行保证 adapter 可靠。
      console.warn(
        `[snapshot-api] reinitializeAfterReload: adapter.list failed for docId=${this.docId}; nextId kept at ${this.nextId}`,
        err,
      );
    }
  }

  async initializeFromStorage(): Promise<void> {
    const infos = await this.adapter.list(this.docId);
    const maxNum = infos
      .map((i) => /^snap-(\d+)$/.exec(i.id)?.[1])
      .filter((v): v is string => v !== undefined)
      .map(Number)
      .reduce((a, b) => Math.max(a, b), -1);
    this.nextId = maxNum + 1;
    this.changelog = [];
  }
}

function toWordError(e: unknown, docId: string): Error {
  if (e instanceof WordError) return e;
  const msg = e instanceof Error ? e.message : String(e);
  return new WordError('PACKAGE_CORRUPT', `snapshot rollback failed: ${msg}`, docId);
}
