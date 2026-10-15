/**
 * InMemorySnapshotStorage —— 默认实现,进程内 Map,刷新即失。
 *
 * 适用:测试 / CLI / 不要持久化的场景。宿主持久化请自行实现 SnapshotStorageAdapter
 * 并通过 `LoadOptions.snapshotStorage` 或 `WordDocument.setDefaultSnapshotStorage`
 * 注入(详见 docs/03.2 §8.5)。
 *
 * docId 命名空间隔离由嵌套 Map 实现:{ docId → { snapshotId → entry } }。
 *
 * **pinned 支持(#108,docs/03.2 §3.3)**:
 * - `save` 接受 `meta.pinned`,存进 entry 并在 `list` 的 SnapshotInfo 上还原
 * - `list` 里 `pinned === true` 的快照标志 SnapshotApi 的 FIFO 跳过它们
 */
import type { SnapshotInfo, SnapshotPayload, SnapshotStorageAdapter } from './types.js';

/** Map 里的条目:payload + pinned bit(bit 与 payload 解耦,便于未来扩展 meta) */
interface StoredEntry {
  payload: SnapshotPayload;
  pinned: boolean;
}

export class InMemorySnapshotStorage implements SnapshotStorageAdapter {
  /** 嵌套 Map:docId → (snapshotId → entry) */
  private readonly store = new Map<string, Map<string, StoredEntry>>();

  private getBucket(docId: string, createIfMissing: boolean): Map<string, StoredEntry> | undefined {
    let bucket = this.store.get(docId);
    if (!bucket && createIfMissing) {
      bucket = new Map();
      this.store.set(docId, bucket);
    }
    return bucket;
  }

  async save(
    docId: string,
    id: string,
    payload: SnapshotPayload,
    meta?: { pinned?: boolean },
  ): Promise<void> {
    const bucket = this.getBucket(docId, true)!;
    bucket.set(id, { payload, pinned: meta?.pinned === true });
  }

  async load(docId: string, id: string): Promise<SnapshotPayload> {
    const bucket = this.getBucket(docId, false);
    const entry = bucket?.get(id);
    if (!entry) {
      throw new Error(`snapshot not found: docId=${docId} id=${id}`);
    }
    return entry.payload;
  }

  async delete(docId: string, id: string): Promise<void> {
    // 幂等:不存在时静默
    this.getBucket(docId, false)?.delete(id);
  }

  async list(docId: string): Promise<SnapshotInfo[]> {
    const bucket = this.getBucket(docId, false);
    if (!bucket) return [];
    // [契约遵守] 按 id 自增数字升序返回(snap-0, snap-1, ...);
    // SnapshotStorageAdapter.list 契约要求,FIFO 淘汰依赖此顺序。
    // 截取 suffix 数字比较,非数字走字典序回退。
    // pinned 体现在 SnapshotInfo.pinned 字段(不参与排序,仅作为 FIFO 淘汰的筛选器)
    const infos = Array.from(bucket.values()).map((entry) => {
      // 重新构造 info 以携带 pinned(payload.info 来自 SnapshotApi,本 adapter 持有 pinned 真相)
      const base = entry.payload.info;
      return entry.pinned ? { ...base, pinned: true } : base;
    });
    return infos.sort((a, b) => {
      const na = extractIdNum(a.id);
      const nb = extractIdNum(b.id);
      if (na !== null && nb !== null) return na - nb;
      return a.id.localeCompare(b.id);
    });
  }

  async clear(docId: string): Promise<void> {
    this.store.delete(docId);
  }
}

/** 从 'snap-123' 提取 123;非 'snap-\d+' 模式返 null */
function extractIdNum(id: string): number | null {
  const m = /^snap-(\d+)$/.exec(id);
  return m ? Number(m[1]) : null;
}
