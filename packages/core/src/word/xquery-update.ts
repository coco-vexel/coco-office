/**
 * XQuery Update Facility 封装 —— docs/03 §13.2 的 "PUL apply" 行。
 *
 * 对 fontoxpath 的 `evaluateUpdatingExpression` + `executePendingUpdateList` 做轻量包装:
 *   - 输入:XQuery Update 表达式 + 上下文节点
 *   - 输出:PUL 条目数量(mutations) + 应用后的 xdmValue
 *
 * 用法例:
 *   ```ts
 *   await applyXQueryUpdate(
 *     paraElement,
 *     'replace value of node ./w:r/w:t with "new text"',
 *     { namespaceResolver: (p) => p === 'w' ? W_NS : null }
 *   );
 *   ```
 *
 * 注意：fontoxpath 没有 package.json `exports` 字段，其 main 是 UMD bundle，
 * Node.js ESM 加载器无法静态解析命名导出，因此改用惰性动态 import()。
 * Node.js 在 ESM 中 import() 一个 CJS/UMD 模块时，会将其 module.exports 包装
 * 成 { default: <module.exports> }，故下方通过 `.default` 取出完整 API。
 *
 * 注意：XQuery 表达式自身性能需要调用方注意 —— 参见 docs/03 §13.2 的 460× 悬崖(
 * 无谓词 //w:p 在 10 万段落上 36s)；本函数不做防御，调用方应只传写针对特定 element
 * 的精确表达式。
 */

type EvaluateUpdatingFn = (
  expression: string,
  contextNode: unknown,
  domFacade: null,
  variables: unknown,
  options: unknown,
) => Promise<{ pendingUpdateList: unknown[]; xdmValue: unknown[] }>;

type ExecutePendingUpdateListFn = (pendingUpdateList: unknown[]) => void;

interface FontoxpathApi {
  evaluateUpdatingExpression: EvaluateUpdatingFn;
  executePendingUpdateList: ExecutePendingUpdateListFn;
}

/** 惰性加载 fontoxpath，避免 ESM 静态 import 命名导出失败。 */
let _fontoxpathPromise: Promise<FontoxpathApi> | undefined;
function getFontoxpath(): Promise<FontoxpathApi> {
  if (!_fontoxpathPromise) {
    _fontoxpathPromise = import('fontoxpath').then((mod) => {
      // ESM 导入 UMD/CJS 模块时，module.exports 挂在 .default 上
      const api = (mod as { default?: FontoxpathApi } & FontoxpathApi).default ?? mod;
      return api as FontoxpathApi;
    });
  }
  return _fontoxpathPromise;
}

export interface XQueryUpdateOptions {
  /** 命名空间 prefix → URI 的解析器。默认 null（表达式不含 prefix） */
  namespaceResolver?: (prefix: string) => string | null;
  /** 外部变量（$name → 值） */
  variables?: Record<string, unknown>;
}

export interface XQueryUpdateResult {
  /** PUL 条目数量（该次应用的变更数） */
  mutations: number;
  /** 表达式返回值（XQuery Update 通常无返回值，此处保持 opaque） */
  xdmValue: unknown[];
}

export async function applyXQueryUpdate(
  contextNode: Node,
  expression: string,
  options: XQueryUpdateOptions = {},
): Promise<XQueryUpdateResult> {
  const { evaluateUpdatingExpression, executePendingUpdateList } = await getFontoxpath();

  const fontOptions: { namespaceResolver?: (prefix: string) => string | null } = {};
  if (options.namespaceResolver) {
    fontOptions.namespaceResolver = options.namespaceResolver;
  }

  const { pendingUpdateList, xdmValue } = await evaluateUpdatingExpression(
    expression,
    contextNode as never,
    null,
    options.variables ?? {},
    // fontoxpath 的 UpdatingOptions 和 Options 结构部分重叠；仅传 namespaceResolver
    fontOptions as never,
  );

  const mutations = pendingUpdateList.length;
  executePendingUpdateList(pendingUpdateList);

  return { mutations, xdmValue };
}
