/**
 * Run 操作原语 —— 给后续 #28 replaceTextInParagraph 调用。
 *
 * 本模块只提供最小、可组合的基本操作:
 *   - getParagraphRuns / getRunText / getParagraphText:只读查询
 *   - splitRunAt:在 run 内 text offset 处拆一个 run 为两个(共享 <w:rPr>)
 *
 * **边界**(本期限制,后续按需扩展):
 *   - splitRunAt 仅支持 "run 下恰有一个 <w:t>" 的简单情况
 *   - 复杂结构(多个 <w:t>、含 <w:br> / <w:tab> / <w:sym> / <w:fldChar> 等)明确抛错
 *   - <w:t xml:space="preserve"> 的 xml:space 属性在拆分时保留
 *
 * **保真读 opt-in**（task-20260623-core-tab-faithful-read）：
 *   - `getRunText(run, { structuralWhitespace: true })` 按 appendSegmentsToRun 的精确逆运算还原
 *     `<w:tab/>` → `\t`、`<w:br/>` （无 type 或 textWrapping）→ `\n`。
 *   - 默认（不传 opts 或 structuralWhitespace: false）行为逐字节不变：只拼 `<w:t>`。
 *   - **偏移算术调用方**（bookmark-ops / comment-ops / paragraph-ops 共 7 处）
 *     **不得**传 `structuralWhitespace: true`——偏移依赖 tab 算 0 长保持稳定。
 */
import { W_NS } from './namespaces.js';

/**
 * `getRunText` / `getParagraphText` 可选参数（task-20260623-core-tab-faithful-read）。
 *
 * - `structuralWhitespace: true`：保真读模式——按 appendSegmentsToRun 的精确逆运算，
 *   `<w:tab/>` → `\t`、`<w:br/>` （无 w:type 或 w:type="textWrapping"）→ `\n`。
 * - 缺省 / `false`：现行为（只拼 `<w:t>`），**偏移算术安全**。
 */
export interface RunTextOptions {
  /**
   * 启用结构空白保真读：`<w:tab/>` → `\t`，`<w:br/>` → `\n`。
   * **仅用于内容读取**，不得用于偏移算术（bookmark/comment/replaceText 等）。
   */
  structuralWhitespace?: boolean;
}

function isWElement(node: Node, localName: string): boolean {
  return (
    node.nodeType === 1 &&
    (node as Element).namespaceURI === W_NS &&
    (node as Element).localName === localName
  );
}

/** 段落 `<w:p>` 下所有直接 `<w:r>` 子元素(不递归到嵌套段落 / smartTag 等) */
export function getParagraphRuns(paragraph: Element): Element[] {
  const runs: Element[] = [];
  const children = paragraph.childNodes;
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (child && isWElement(child, 'r')) {
      runs.push(child as Element);
    }
  }
  return runs;
}

/**
 * 取 run 的文本。
 *
 * - 默认（opts 缺省 / `structuralWhitespace: false`）：只拼 `<w:t>` 文本，忽略 `<w:br>` / `<w:tab>`。
 *   **偏移算术安全**：`<w:tab/>` 算 0 长，保持 bookmark/comment/replaceText 偏移稳定。
 * - `structuralWhitespace: true`：保真读——按文档顺序遍历子节点：
 *   `<w:t>` → textContent、`<w:tab/>` → `\t`、`<w:br/>` （无 w:type 或 textWrapping）→ `\n`。
 *   这是 `appendSegmentsToRun` 的精确逆运算。
 */
export function getRunText(run: Element, opts?: RunTextOptions): string {
  const faithful = opts?.structuralWhitespace === true;
  let text = '';
  const children = run.childNodes;
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (!child || child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI !== W_NS) continue;
    const local = el.localName;
    if (local === 't') {
      text += el.textContent ?? '';
    } else if (faithful) {
      if (local === 'tab') {
        text += '\t';
      } else if (local === 'br') {
        // 只还原无 type 或 type="textWrapping" 的 <w:br/>（写侧产出的两种）
        // page/column break 等不在写侧逆运算内，忽略
        const type = el.getAttributeNS(W_NS, 'type');
        if (type === null || type === '' || type === 'textWrapping') {
          text += '\n';
        }
      }
    }
  }
  return text;
}

/**
 * 段落的纯文本（所有直接 run 的文本拼接）。
 *
 * - 默认（opts 缺省）：只拼 `<w:t>`，偏移算术安全。
 * - `structuralWhitespace: true`：保真读，透传到 `getRunText`。
 */
export function getParagraphText(paragraph: Element, opts?: RunTextOptions): string {
  return getParagraphRuns(paragraph)
    .map((r) => getRunText(r, opts))
    .join('');
}

/** splitRunAt 返回值 */
export interface SplitResult {
  /** 原 run,文本截取到 [0, offset) */
  before: Element;
  /**
   * 分裂出的新 run,文本为 [offset, end)。
   * 与 `before` 共享 `<w:rPr>` 结构(深拷贝,不共享对象引用)。
   * offset=0 或 offset=runLen 时为 null(未发生分裂)
   */
  after: Element | null;
}

/**
 * 在 run 的文本 offset 处拆分。
 *
 * - `offset === 0` 或 `offset === getRunText(run).length` → 不拆,返回 { before: run, after: null }
 * - 其它情况 → 新 run 插入到原 run 的 nextSibling 位置,返回 { before: run, after: newRun }
 *
 * 本期限制:run 必须恰有 1 个 `<w:t>` 子;否则抛错(多 <w:t>、<w:br>、<w:tab> 等复杂结构交给
 * 后续 #28 的扩展实现)。
 */
export function splitRunAt(run: Element, offset: number): SplitResult {
  if (!isWElement(run, 'r')) {
    throw new Error('splitRunAt: argument is not a <w:r> element');
  }
  const tElements: Element[] = [];
  const children = run.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (c && isWElement(c, 't')) tElements.push(c as Element);
  }
  if (tElements.length === 0) {
    throw new Error(
      'splitRunAt: run has no <w:t> children; this primitive supports single-<w:t> runs only',
    );
  }
  if (tElements.length > 1) {
    throw new Error(
      `splitRunAt: run has ${tElements.length} <w:t> children; this primitive supports single-<w:t> runs only`,
    );
  }

  const originalT = tElements[0]!;
  const runText = originalT.textContent ?? '';
  if (offset < 0 || offset > runText.length) {
    throw new Error(`splitRunAt: offset ${offset} out of range [0, ${runText.length}]`);
  }
  if (offset === 0 || offset === runText.length) {
    return { before: run, after: null };
  }

  const leftText = runText.slice(0, offset);
  const rightText = runText.slice(offset);

  // 截取原 run 的文本到左半
  originalT.textContent = leftText;

  // 克隆整个 run 作为 after;克隆深拷贝会把 <w:rPr> 和 <w:t xml:space> 一起复制
  const afterRun = run.cloneNode(true) as Element;
  // 找到 after run 中的那个 <w:t>,写入右半文本
  const afterChildren = afterRun.childNodes;
  for (let i = 0; i < afterChildren.length; i++) {
    const c = afterChildren[i];
    if (c && isWElement(c, 't')) {
      (c as Element).textContent = rightText;
      break;
    }
  }

  // 插入 afterRun 到原 run 之后
  const parent = run.parentNode;
  if (!parent) {
    throw new Error('splitRunAt: run has no parent; cannot insert new run');
  }
  parent.insertBefore(afterRun, run.nextSibling);

  return { before: run, after: afterRun };
}
