/**
 * 快照子系统 —— 与文档类型无关的存储层类型定义。
 *
 * 设计背景:docs/03.2-快照系统.md(尤其 §2 API 形态、§8 存储适配层)。
 *
 * 层次:
 *   - 本文件:SnapshotInfo / ChangeLog / SnapshotPayload / SnapshotStorageAdapter
 *   - `./in-memory.ts`:默认 InMemorySnapshotStorage 实现
 *   - 后续 #21 在 WordDocument 上挂 SnapshotApi 子 API,消费本适配器
 */

/** 消费方通过 `wordDoc.snapshots.recordChange()` 累加的变更日志项 */
export interface ChangeLog {
  /** 时间戳,Date.now() */
  timestamp: number;
  /** 工具名,例 'replace_paragraphs' */
  toolName: string;
  /** 工具自定义摘要(结构由消费方决定,可 JSON.stringify) */
  summary?: unknown;
}

/** 快照元信息(`list` 只返回这些,不含 parts 字节) */
export interface SnapshotInfo {
  /** 自增 id,'snap-0' / 'snap-1' / ... */
  id: string;
  /** 时间戳,Date.now() */
  timestamp: number;
  /** 用户提供的 take 原因 */
  reason?: string;
  /** 用户提供的 take 详细结构 */
  details?: unknown;
  /** 当前 OOXML Package 全部 parts 的稳定 SHA-256 指纹。旧快照可缺省。 */
  contentHash?: string;
  /** 逻辑去重域；仅同一 dedupeKey 的快照之间允许复用 id。 */
  dedupeKey?: string;
  /** take 时一次性吸入的主线程 changelog */
  changesCaptured: ChangeLog[];
  /**
   * Pinned 标记(#108,docs/03.2 §3.3):true 时该快照不参与 FIFO 淘汰。
   *
   * - `undefined` / `false` 等价,均视为普通快照
   * - 用于保护 `WordAgent.create` 的 session-start 快照不被 plugin 的
   *   常规 take 挤掉
   * - 向后兼容:旧 payload / 旧 adapter 未带该字段时视为未 pinned
   */
  pinned?: boolean;
}

/** 快照完整载荷(info + parts 字节) */
export interface SnapshotPayload {
  info: SnapshotInfo;
  parts: Array<{
    path: string;
    contentType: string;
    bytes: Uint8Array;
  }>;
}

/**
 * 存储适配器接口。宿主(WPS 面板 / Electron main / 浏览器 / Node server)提供实现。
 *
 * **语义承诺**:
 * - `save` / `load` 捆装 info + parts 为一次写入 / 一次读取
 * - `list` **仅返 meta**(SnapshotInfo 数组),不加载 parts(避免大档进主线程)
 * - `delete(docId, id)`:不存在时不抛错,幂等
 * - `clear(docId)`:清除该 docId 下所有快照,不影响其它 docId
 * - 所有方法 async(允许底层存储异步)
 *
 * **pinned 支持(#108)**:
 * - `save` 接受可选 `meta.pinned`,adapter 需把该 bit 与 payload 一起持久化,
 *   并在 `list` 返回的 `SnapshotInfo.pinned` 中还原
 * - 不支持 pinned 的旧 adapter 仍然合法:`list` 返回的 `pinned`
 *   默认为 `undefined`,被 SnapshotApi 视为未 pinned(向后兼容)
 * - 具体的 "FIFO 淘汰时跳过 pinned" 逻辑在 **SnapshotApi 主线程侧**,
 *   adapter 只负责存取该 bit
 *
 * SDK 不做:重试 / 超时 / 并发互斥。异常透传给消费方。
 */
export interface SnapshotStorageAdapter {
  /**
   * 保存快照。
   *
   * @param meta - 可选;目前仅 `pinned`(#108),表示本快照不应被 FIFO 淘汰。
   *   适配器应把 `pinned` 与 payload 一起持久化,并在 `list` 时体现在
   *   `SnapshotInfo.pinned` 上。未实现 pinned 的旧 adapter 可忽略该参数。
   */
  save(
    docId: string,
    id: string,
    payload: SnapshotPayload,
    meta?: { pinned?: boolean },
  ): Promise<void>;
  load(docId: string, id: string): Promise<SnapshotPayload>;
  delete(docId: string, id: string): Promise<void>;
  /**
   * 列出某 docId 下所有快照元信息。
   *
   * **必须按 id 自增数字升序返回(snap-0, snap-1, ..., snap-N)**。
   * SnapshotApi 的 FIFO 淘汰(按"最早快照"删)依赖该顺序;若 adapter 返回顺序与
   * id 自增序不一致,FIFO 会淘汰错对象。
   *
   * `SnapshotInfo.pinned` 反映 adapter 保存时收到的 `meta.pinned`;未实现 pinned
   * 的旧 adapter 不带该字段(视为未 pinned)。
   */
  list(docId: string): Promise<SnapshotInfo[]>;
  clear(docId: string): Promise<void>;
}
