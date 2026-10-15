/**
 * Plan executor 共享上下文类型。
 *
 * 独立文件的原因:打破 barrel 循环(#110)——若放在 `index.ts`,ops 文件反向
 * import type 会与 `index.ts` 正向 import executor 形成循环边,即使是 type-only
 * 也会被 dep-cruiser 拦下。
 */
import type { Package } from '../../pack/index.js';
import type { PackageIndices } from '../indices.js';
import type { ApplyResult } from '../plan.js';

/** Plan apply 期间传给 executor 的上下文 */
export interface ExecutorContext {
  pkg: Package;
  indices: PackageIndices;
  options: {
    trackChanges: boolean;
    author: string;
    /** 解码后的 Date(kernel.applyPlan 负责 new Date(iso))*/
    date: Date;
  };
  /** 当前 ApplyResult,executor mutate 积累字段 */
  result: ApplyResult;
  /** 当前 op 在 ops 数组里的 index(用于 insertedParagraphs 等的 opIndex 字段)*/
  opIndex: number;
  /** docId,用于 WordError.docId 注入 */
  docId: string;
  /**
   * 候选孤儿 media part 路径(#190):executor 在替换 / 删除 drawing 时把不再被
   * 任何 rels Relationship 引用的 media path 记录在此。kernel 在 runTransaction
   * **commit 之后**统一扫一遍这个集合,真正归零的才删 binary part。
   *
   * 为什么不在事务内删:transactions.ts 支持回滚 pre-existing part 的删除,但若
   * 后续 op 抛错回滚,rels 已经被回滚到含原 Relationship 的状态——media 应该跟着
   * "未删"。在事务外删可避免 collectAffectedPartPaths 提前 lookup 媒体路径的复杂度,
   * 且只在事务整体成功时才落地真正的二进制删除。
   *
   * Set 而非 Array:同一 media 可能被多个 op / drawing 引用,只需登记一次。
   *
   * 可选字段:仅 drawing-ops 使用,其它 executor 不感知。kernel.applyPlan 总会
   * 提供一个空 Set;旧 executor 单元测试不传也兼容(由 drawing-ops 防御性处理)。
   */
  orphanMediaCandidates?: Set<string>;
}
