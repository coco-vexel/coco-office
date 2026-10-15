/**
 * WordError —— Word core 的统一错误模型(docs/03.1 §4)。
 *
 * 所有面向消费方的错误都用 WordError 抛出。`code` 枚举覆盖装载期 / 查询 / 变更前置 / Plan 状态
 * 四大类。`docId` 由 core 自动填充(WordDocument 内捕获底层错误后重抛时注入),方便日志
 * 与多文档场景定位。
 */

export type WordErrorCode =
  // ── 装载期 ──
  | 'PACKAGE_CORRUPT' // 字节 / XML 格式错
  | 'PACKAGE_UNSUPPORTED_FORMAT' // 不是 Word docx / Flat OPC
  | 'BOOKMARK_NAME_DUPLICATE_ON_LOAD' // 【已弃用 / 不再装载期抛出】重名书签现为容错(保留第一个),
  //                                  // code 保留仅作向后兼容,buildIndices 不再产生它(见 indices.ts)
  | 'BOOKMARK_NAME_EMPTY_ON_LOAD'
  | 'WORKER_CREATION_FAILED' // mode='worker' 显式但创建失败
  | 'WORKER_RELOADED' // worker 崩溃后 terminate + reload,执行中的 RPC 以此拒收(docs/03 §4.7)
  // ── 查询 ──
  | 'PARA_NOT_FOUND'
  | 'BOOKMARK_NOT_FOUND'
  | 'COMMENT_NOT_FOUND'
  | 'REVISION_NOT_FOUND'
  | 'TABLE_NOT_FOUND' // 表格 / 单元格 / 行按 (tableId, row?, col?) 查不到; details 带 row/col 详情
  | 'DRAWING_NOT_FOUND' // drawing 按 drawingId 查不到
  // ── 变更前置 ──
  | 'TEXT_NOT_FOUND_IN_PARA'
  | 'BOOKMARK_NAME_EXISTS'
  | 'BOOKMARK_NAME_INVALID'
  | 'ANCHOR_CONFLICT' // 变更区间覆盖活锚点
  | 'UNSUPPORTED_OOXML_FEATURE' // 本期不支持的 OOXML 结构(pPrChange/rPrChange 等)
  | 'INVALID_ARGUMENT' // op 参数越界 / 逻辑错误(如 outlineLevel 超出 0-8 范围)
  // ── Plan 状态 ──
  | 'PLAN_ALREADY_APPLIED'
  | 'PLAN_EMPTY' // apply() 空 plan;默认 warn 不抛,严格模式可抛
  // ── 排版（typeset Layer4）内容核对门禁（D5,接缝4「绝不丢内容」）──
  | 'CONTENT_LOST'; // 排版产物核对发现 content 元素丢失;details 带 lost[] 明细 + metrics

export class WordError extends Error {
  readonly code: WordErrorCode;
  readonly docId: string;
  readonly details: Record<string, unknown>;

  constructor(
    code: WordErrorCode,
    message: string,
    docId: string,
    details: Record<string, unknown> = {},
  ) {
    super(message);
    this.name = 'WordError';
    this.code = code;
    this.docId = docId;
    this.details = details;
  }
}
