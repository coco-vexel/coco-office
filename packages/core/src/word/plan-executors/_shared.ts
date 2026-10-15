/**
 * Executor 共享工具 —— 由 #29 提取,供 #34/#35 等后续 PR 复用。
 *
 * 从 paragraph-ops.ts 提出的三个内部工具:
 *   - createW / createT:创建带正确命名空间的 <w:*> / <w:t> 元素
 *   - appendTextAsRuns:把 text(含 \n / \t)追加为单 <w:r> 到段落,可选传入 rPr
 *   - firstDirectChildW:找 element 直属的首个 W_NS 子元素
 *
 * #30 追加:
 *   - ParagraphAnchors / collectParagraphAnchors:扫段内活锚点(bookmarks + comments)
 *   - collectRangeAnchors:扫一段 sibling 链内的锚点(供 replaceTextInParagraph 使用)
 *
 * task-core-replacepara-rpr 追加（提取自 table-ops.ts，统一事实源）:
 *   - pickFirstRunRPr：取段落首个 run 后代的 rPr，run 无 rPr 时兜底取 pPr/rPr
 *     供 replaceParagraph（paragraph-ops）和 replaceTableCellText（table-ops）共用
 */
import { W_NS, XML_NS } from '../namespaces.js';

// ─────────── #30 锚点检测 ───────────

/** 段落内活锚点汇总 */
export interface ParagraphAnchors {
  /** bookmarkStart 的 w:name 数组 */
  bookmarks: string[];
  /** commentRangeStart / commentReference 的 w:id 数组(去重) */
  comments: number[];
}

/**
 * 扫 `paragraph` 内所有 bookmarkStart(取 w:name)
 * 和 commentRangeStart / commentReference(取 w:id,去重),
 * 返回 `ParagraphAnchors`。
 *
 * 禁无谓词 XPath — 一律 `getElementsByTagNameNS`。
 */
export function collectParagraphAnchors(paragraph: Element): ParagraphAnchors {
  const bookmarks: string[] = [];
  const commentIds = new Set<number>();

  const bmStarts = paragraph.getElementsByTagNameNS(W_NS, 'bookmarkStart');
  for (let i = 0; i < bmStarts.length; i++) {
    const el = bmStarts[i];
    if (!el) continue;
    const name = el.getAttributeNS(W_NS, 'name') ?? el.getAttribute('w:name') ?? '';
    bookmarks.push(name);
  }

  const cStarts = paragraph.getElementsByTagNameNS(W_NS, 'commentRangeStart');
  for (let i = 0; i < cStarts.length; i++) {
    const el = cStarts[i];
    if (!el) continue;
    const raw = el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '';
    const id = parseInt(raw, 10);
    if (!isNaN(id)) commentIds.add(id);
  }

  const cRefs = paragraph.getElementsByTagNameNS(W_NS, 'commentReference');
  for (let i = 0; i < cRefs.length; i++) {
    const el = cRefs[i];
    if (!el) continue;
    const raw = el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '';
    const id = parseInt(raw, 10);
    if (!isNaN(id)) commentIds.add(id);
  }

  return { bookmarks, comments: Array.from(commentIds) };
}

/**
 * 扫 sibling 链 [startNode, endNode] 之间(含两端)的锚点 marker,
 * 用于 replaceTextInParagraph 的区间锚点检测。
 *
 * startNode / endNode 是 paragraph 直属 childNodes 内的节点。
 * 扫描范围:从 startNode 到 endNode(含)的所有 Element 子节点,
 * 以及这些节点自身若是 bookmarkStart / commentRangeStart / commentReference。
 *
 * 注意:锚点 marker 通常是段直属的 Element(非 run 子),
 * 遇到就算在区间内。
 */
export function collectRangeAnchors(startNode: Node, endNode: Node): ParagraphAnchors {
  const bookmarks: string[] = [];
  const commentIds = new Set<number>();

  let cur: Node | null = startNode;
  while (cur) {
    if (cur.nodeType === 1) {
      const el = cur as Element;
      const localName = el.localName;
      if (el.namespaceURI === W_NS) {
        if (localName === 'bookmarkStart') {
          const name = el.getAttributeNS(W_NS, 'name') ?? el.getAttribute('w:name') ?? '';
          bookmarks.push(name);
        } else if (localName === 'commentRangeStart' || localName === 'commentReference') {
          const raw = el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id') ?? '';
          const id = parseInt(raw, 10);
          if (!isNaN(id)) commentIds.add(id);
        }
      }
    }
    if (cur === endNode) break;
    cur = cur.nextSibling;
  }

  return { bookmarks, comments: Array.from(commentIds) };
}

/** 创建 `<w:localName>` 元素(W_NS)*/
export function createW(doc: Document, localName: string): Element {
  return doc.createElementNS(W_NS, `w:${localName}`);
}

/**
 * 创建 `<w:t xml:space="preserve">` 元素。
 * 所有 `<w:t>` 一律挂 preserve,保留首尾空格。
 */
export function createT(doc: Document, text: string): Element {
  const t = createW(doc, 't');
  t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  t.textContent = text;
  return t;
}

/**
 * 把 text 里的 `\n` / `\t` 映射为 `<w:br/>` / `<w:tab/>`,其它为 `<w:t>`。
 * 追加为单一 `<w:r>` 到 paragraph 末尾。
 * 空 text → 一个空 `<w:t/>` placeholder(段不能完全无 run)。
 *
 * `rPr`（可选）：若传入，插到 run 首位作为 `<w:rPr>`，透传字号/字体等格式。
 */
export function appendTextAsRuns(paragraph: Element, text: string, rPr?: Element): void {
  const doc = paragraph.ownerDocument;
  if (!doc) throw new Error('paragraph has no ownerDocument');

  const run = createW(doc, 'r');
  if (rPr) run.insertBefore(rPr, run.firstChild);
  appendSegmentsToRun(doc, run, text);
  paragraph.appendChild(run);
}

/**
 * 把 text 里的字符映射为 run 的子元素(`\n`→`<w:br/>` / `\t`→`<w:tab/>` / other→`<w:t>`)。
 * export 供 buildReplacementRun 等需要往已存在 run 里填内容的场景使用。
 */
export function appendSegmentsToRun(doc: Document, run: Element, text: string): void {
  if (text.length === 0) {
    run.appendChild(createT(doc, ''));
    return;
  }
  let buf = '';
  const flush = (): void => {
    if (buf.length > 0) {
      run.appendChild(createT(doc, buf));
      buf = '';
    }
  };
  for (const ch of text) {
    if (ch === '\n') {
      flush();
      run.appendChild(createW(doc, 'br'));
    } else if (ch === '\t') {
      flush();
      run.appendChild(createW(doc, 'tab'));
    } else {
      buf += ch;
    }
  }
  flush();
}

/** 找 `el` 直属子中首个 localName 匹配的 W_NS 元素(null 表示无)*/
export function firstDirectChildW(el: Element, localName: string): Element | null {
  const children = el.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === localName
    ) {
      return c as Element;
    }
  }
  return null;
}

/**
 * 取段落内**首个 `<w:r>` 后代**的 `<w:rPr>` 克隆；run 无 rPr 时兜底取 `pPr/rPr`。
 *
 * 通用化（task-core-cell-rpr R1）：
 *   用 `getElementsByTagNameNS(W_NS,'r')` 取后代 run（DOM document order），
 *   天然覆盖 SDT、hyperlink、smartTag、fldSimple、ins 等任意包裹容器内的 run，
 *   无需逐一判断包裹类型。
 *
 * pPr/rPr 兜底（task-core-cell-rpr R2）：
 *   首个 run 无显式 rPr 时，回退到段落 `<w:pPr><w:rPr>`——标准 OOXML"段落默认
 *   run 属性"来源，通用且符合规范语义。
 *
 * 优先级：
 *   1. 首个后代 run 有 rPr → 返回 run rPr 克隆
 *   2. 首个后代 run 无 rPr，但 pPr/rPr 存在 → 返回 pPr/rPr 克隆
 *   3. 两者都无 → 返回 undefined（裸 run，语义不变）
 *
 * 安全边界：OOXML 规范段落不嵌套表格（嵌套表是 `<w:tc>` 的兄弟，非段落子），
 *   故 `getElementsByTagNameNS` 只覆盖本段显示内容内的 run。
 *
 * 提取自 table-ops.ts（task-core-replacepara-rpr），paragraph-ops 与 table-ops 共用。
 *
 * @returns rPr 克隆，或 undefined（无可用格式来源）。
 */
export function pickFirstRunRPr(paragraph: Element): Element | undefined {
  // 1. 首个后代 run 的 rPr（覆盖任意包裹容器）
  const runs = paragraph.getElementsByTagNameNS(W_NS, 'r');
  if (runs.length > 0) {
    const rPr = firstDirectChildW(runs[0]!, 'rPr');
    if (rPr) return rPr.cloneNode(true) as Element;
  }
  // 2. 兜底：pPr/rPr（段落默认 run 属性）
  const pPr = firstDirectChildW(paragraph, 'pPr');
  if (pPr) {
    const pRPr = firstDirectChildW(pPr, 'rPr');
    if (pRPr) return pRPr.cloneNode(true) as Element;
  }
  return undefined;
}
