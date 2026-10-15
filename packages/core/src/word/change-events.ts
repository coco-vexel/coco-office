/**
 * change-events —— `WordDocument.on('change', ...)` 的类型定义。
 *
 * 文档变更事件在 `kernel.applyPlan()` 成功后触发,涵盖两条入口:
 *   - `doc.plan().apply()`        →  source = 'plan'
 *   - `doc.executeCode(script)`   →  source = 'script'(脚本内调 `doc.applyPlan`)
 *
 * **不**触发的场景:
 *   - 空 plan(`ops.length === 0`)—— kernel 直接返回,不算一次变更
 *   - applyPlan 抛错回滚 —— 文档状态未变
 *   - 脚本绕过 kernel.applyPlan 直接改 `doc.getPackage()` 的 DOM —— v1 不覆盖,
 *     此类用法被视作"自负其责";后续可加 dirty-watcher,但代价大。
 *
 * Payload 默认只带 `ops` + `result` + `source`,**不**带全文档 —— 全文档输出
 * (toFlatOpc / toZip)在 worker 模式下要 RPC + serialize,10 万段落级别可达数百 ms,
 * 不是"白送"的开销。订阅时通过 `fullDocument: 'flatOpc' | 'zip'` 显式开启,本模块
 * 在派发前由 WordDocument 一次性拉取(多订阅者共享)。
 */
import type { ApplyResult, Op } from './plan.js';

/**
 * Kernel 内部的 change event payload —— 走 worker 边界时被 structured-clone,
 * 因此所有字段必须是 plain JSON。
 */
export interface KernelChangeEvent {
  /** 触发本次 apply 的 op 队列(完整传出,便于审计 / 重放) */
  ops: Op[];
  /** kernel.applyPlan 的返回值 */
  result: ApplyResult;
  /**
   * 'plan'   = 走 RPC `applyPlan`(`doc.plan().apply()` 路径)
   * 'script' = 走 `doc.applyPlan(...)`(`doc.executeCode(...)` 内部脚本路径)
   *
   * 由触发层(kernel-handlers / script-execute)显式标注;kernel.applyPlan 自己不知道
   * 来源,只接收 `opts.source`。
   */
  source: 'plan' | 'script';
}

/**
 * `WordDocument.on('change', handler)` 收到的 payload。比 KernelChangeEvent 多了
 * 可选的全文档字段 —— 是否填充由订阅时的 `fullDocument` 决定。
 */
export interface ChangeEvent extends KernelChangeEvent {
  /** 仅当 `subscribe.fullDocument === 'flatOpc'` 时存在 */
  flatOpc?: string;
  /** 仅当 `subscribe.fullDocument === 'zip'` 时存在 */
  zip?: Uint8Array;
}

/** 订阅选项;默认 `fullDocument: 'none'`(只发 metadata,不付全文档输出代价)。 */
export interface ChangeSubscribeOptions {
  /**
   * - `'none'`(默认):payload 只含 `{ ops, result, source }`
   * - `'flatOpc'`:派发前调一次 `toFlatOpc()`,payload.flatOpc 填充 XML 字符串
   * - `'zip'`:派发前调一次 `toZip()`,payload.zip 填充 docx 字节
   *
   * **代价警示**:`'flatOpc'` / `'zip'` 在 worker 模式下要 RPC + structured-clone
   * 整篇文档,大文档一次几百 ms;只在确实要"快照式消费"时启用。多订阅者共享同一次
   * 拉取,因此打开多个 'flatOpc' 订阅不会重复跑。
   */
  fullDocument?: 'none' | 'flatOpc' | 'zip';
}

/** 订阅回调;可同步可 async。WordDocument 串行 await 每个订阅者,保证派发顺序与变更顺序一致。 */
export type ChangeListener = (ev: ChangeEvent) => void | Promise<void>;

/** RPC event 通道里 word change 事件的 `event` 字符串常量。 */
export const WORD_CHANGE_EVENT = 'word.change';
