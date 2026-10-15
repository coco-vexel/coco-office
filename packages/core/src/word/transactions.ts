/**
 * 跨 parts 事务 —— docs/03 §13.5 的落地。
 *
 * Plan.apply 的底层原语:在一组 part 上执行可能失败的 work,失败时回滚到 work 前状态。
 *
 * 设计:
 *   - 进入事务前,对 `affectedPartPaths` 里每个已存在的 part 做深拷贝快照(xmlDom / bytes / _dirty)
 *   - 记录哪些 path 在 work 开始时**不存在**(新建场景)
 *   - work 内可自由变更 part(包括新增 / 删除 / 改 xmlDom)
 *   - 抛错时:
 *     * 原本存在的 part → 从快照恢复 xmlDom / bytes / _dirty
 *     * 原本不存在但 work 新增的 part → 从 pkg 删除
 *   - 成功时:保持 work 结束时的状态(work 自己负责 markDirty)
 *
 * 当前实现:deepClone 整个 xmlDom,在 10 万段落档上约 1-2s,优化(反向 op 记录)留后续 issue。
 */
import type { Package, Part } from '../pack/index.js';
import { createBinaryPart, createXmlPart, deletePart, getPart, setPart } from '../pack/index.js';

interface PartSnapshot {
  contentType: string;
  xmlDom: Document | undefined;
  bytes: Uint8Array | undefined;
  originalBytes: Uint8Array | undefined;
  dirty: boolean;
}

/**
 * 在 `pkg` 上执行可能失败的 `work`。事务作用域仅覆盖 `affectedPartPaths` 列出的 parts:
 *   - work 抛错 → 被列出的 parts 回滚到进入事务前状态(xmlDom / bytes / _dirty),新增的 parts 删除
 *   - work 成功 → 返回 work 的结果
 *
 * **调用方职责**:
 *   - `affectedPartPaths` 必须包含 work 里**所有可能触达**的 part 路径(读 + 写 + 新增)
 *   - 未在列表里的 part 即便被 work 改动,回滚时不会被恢复
 */
export async function runTransaction<T>(
  pkg: Package,
  affectedPartPaths: readonly string[],
  work: () => Promise<T> | T,
): Promise<T> {
  const snapshots = new Map<string, PartSnapshot>();
  const existedBefore = new Set<string>();

  for (const path of affectedPartPaths) {
    if (!pkg.parts.has(path)) continue;
    existedBefore.add(path);
    const part = getPart(pkg, path)!;
    snapshots.set(path, cloneSnapshot(part));
  }

  try {
    return await Promise.resolve(work());
  } catch (err) {
    rollback(pkg, affectedPartPaths, existedBefore, snapshots);
    throw err;
  }
}

function cloneSnapshot(part: Part): PartSnapshot {
  return {
    contentType: part.contentType,
    xmlDom: part.xmlDom ? (part.xmlDom.cloneNode(true) as Document) : undefined,
    bytes: part.bytes ? new Uint8Array(part.bytes) : undefined,
    originalBytes: part._originalBytes ? new Uint8Array(part._originalBytes) : undefined,
    dirty: part._dirty,
  };
}

function rollback(
  pkg: Package,
  affectedPartPaths: readonly string[],
  existedBefore: Set<string>,
  snapshots: Map<string, PartSnapshot>,
): void {
  for (const path of affectedPartPaths) {
    if (existedBefore.has(path)) {
      // 恢复到快照
      const part = getPart(pkg, path);
      const snap = snapshots.get(path)!;
      if (!part) {
        // 事务里把这个 part 删了 —— 用快照内 contentType + payload 重建。
        //
        // 这是 #190 引入的能力(占位图替换的旧 media 清理需要在事务内删 binary part)。
        // 用 createXmlPart / createBinaryPart 回填 readonly 字段,保持 Part 的"择一"约束。
        let restored: Part;
        if (snap.xmlDom) {
          restored = createXmlPart({
            path,
            contentType: snap.contentType,
            xmlDom: snap.xmlDom,
            ...(snap.originalBytes !== undefined ? { originalBytes: snap.originalBytes } : {}),
          });
        } else if (snap.bytes) {
          restored = createBinaryPart({
            path,
            contentType: snap.contentType,
            bytes: snap.bytes,
            ...(snap.originalBytes !== undefined ? { originalBytes: snap.originalBytes } : {}),
          });
        } else {
          throw new Error(
            `runTransaction rollback: snapshot for '${path}' has neither xmlDom nor bytes`,
          );
        }
        restored._dirty = snap.dirty;
        setPart(pkg, restored);
        continue;
      }
      part.xmlDom = snap.xmlDom;
      part.bytes = snap.bytes;
      part._dirty = snap.dirty;
    } else {
      // 事务里新建的 part(或从未存在过)→ 删除
      deletePart(pkg, path);
    }
  }
}
