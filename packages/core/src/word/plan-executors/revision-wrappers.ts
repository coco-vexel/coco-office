/**
 * revision-wrappers.ts —— #31 trackChanges 包装 helper。
 *
 * 提供把 DOM run 包进 `<w:ins>` / `<w:del>` 的低层工具,以及给段落 mark 加
 * 插入/删除标记的 helper。不依赖 ExecutorContext,只依赖 Package 和选项原始值。
 *
 * **重要 OOXML 约束**:
 *   - `<w:ins>` / `<w:del>` 是 run 的包装层,出现在 `<w:p>` 子层(不跨 w:p)
 *   - `<w:del>` 内的 run 用 `<w:delText>` 而非 `<w:t>`(显示被删内容)
 *   - 段落 mark 的修订标记放在 `<w:pPr><w:rPr><w:ins/>` / `<w:pPr><w:rPr><w:del/>`
 *   - `w:id` 必须在 pkg 全局唯一(扫所有 part 的 ins/del,取 max+1)
 */
import type { Package } from '../../pack/index.js';
import { listParts, isXmlPart } from '../../pack/index.js';
import { W_NS } from '../namespaces.js';
import { createW } from './_shared.js';

/** 用于构建修订属性的最小选项 */
export interface RevisionAttrs {
  author: string;
  date: Date;
}

// ─────────── allocateRevisionId ───────────

/**
 * 扫 pkg 内所有 XML part 里的 `<w:ins>` / `<w:del>`,读取 `w:id` 属性,
 * 返回 max(存在的数字 id) + 1。
 * 无任何 valid id → 返回 0(首次分配)。
 */
export function allocateRevisionId(pkg: Package): number {
  let max = -1;
  const parts = listParts(pkg);
  for (const part of parts) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    const doc = part.xmlDom as Document;
    const insEls = doc.getElementsByTagNameNS(W_NS, 'ins');
    const delEls = doc.getElementsByTagNameNS(W_NS, 'del');
    max = readMaxId(insEls, max);
    max = readMaxId(delEls, max);
  }
  return max + 1;
}

function readMaxId(col: HTMLCollectionOf<Element>, current: number): number {
  let max = current;
  for (let i = 0; i < col.length; i++) {
    const el = col[i];
    if (!el) continue;
    const raw = el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id');
    if (raw === null) continue;
    const n = parseInt(raw, 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return max;
}

// ─────────── setRevisionAttrs ───────────

/** 给 `<w:ins>` / `<w:del>` 元素设置 w:id / w:author / w:date 属性 */
function setRevisionAttrs(el: Element, id: number, attrs: RevisionAttrs): void {
  el.setAttributeNS(W_NS, 'w:id', String(id));
  el.setAttributeNS(W_NS, 'w:author', attrs.author);
  el.setAttributeNS(W_NS, 'w:date', attrs.date.toISOString());
}

// ─────────── wrapRunsAsInsert ───────────

/**
 * 将 runs(深拷贝)包进新建的 `<w:ins w:id="N" w:author="..." w:date="...">`,
 * 并返回该 `<w:ins>` 元素。
 *
 * 调用方负责把返回元素插入 DOM;runs 本身(原始元素)由调用方管理。
 */
export function wrapRunsAsInsert(runs: Element[], id: number, attrs: RevisionAttrs): Element {
  if (runs.length === 0) throw new Error('wrapRunsAsInsert: runs must not be empty');
  const doc = runs[0]!.ownerDocument;
  if (!doc) throw new Error('wrapRunsAsInsert: run has no ownerDocument');

  const ins = createW(doc, 'ins');
  setRevisionAttrs(ins, id, attrs);
  for (const run of runs) {
    ins.appendChild(run.cloneNode(true));
  }
  return ins;
}

// ─────────── wrapRunsAsDelete ───────────

/**
 * 将 runs 深拷贝、把拷贝中所有 `<w:t>` 重建为 `<w:delText>` 后,包进新建的
 * `<w:del>`,返回该 `<w:del>` 元素。
 *
 * 约束:不试图 rename 节点(slimdom 不支持 renameNode),用 createElementNS 创建
 * 新 `<w:delText>`,把 textContent + xml:space 属性搬过来。
 */
export function wrapRunsAsDelete(runs: Element[], id: number, attrs: RevisionAttrs): Element {
  if (runs.length === 0) throw new Error('wrapRunsAsDelete: runs must not be empty');
  const doc = runs[0]!.ownerDocument;
  if (!doc) throw new Error('wrapRunsAsDelete: run has no ownerDocument');

  const del = createW(doc, 'del');
  setRevisionAttrs(del, id, attrs);

  for (const run of runs) {
    const cloned = run.cloneNode(true) as Element;
    convertTtoDelText(doc, cloned);
    del.appendChild(cloned);
  }
  return del;
}

/**
 * 递归遍历 `el` 及其后代,把所有 `<w:t>` 替换为 `<w:delText>`。
 * 属性和 textContent 全搬。
 *
 * 注意:不使用 renameNode(slimdom 不支持)。
 */
function convertTtoDelText(doc: Document, el: Element): void {
  // 先处理子节点(深度优先,避免在枚举时同时修改列表)
  const childrenSnapshot: Node[] = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (c) childrenSnapshot.push(c);
  }

  for (const child of childrenSnapshot) {
    if (
      child.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName === 't'
    ) {
      const t = child as Element;
      const delText = doc.createElementNS(W_NS, 'w:delText');
      // 搬属性(主要是 xml:space="preserve")
      for (let i = 0; i < t.attributes.length; i++) {
        const attr = t.attributes[i];
        if (!attr) continue;
        if (attr.namespaceURI) {
          delText.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
        } else {
          delText.setAttribute(attr.name, attr.value);
        }
      }
      delText.textContent = t.textContent;
      el.replaceChild(delText, t);
    } else if (child.nodeType === 1) {
      // 递归处理子元素
      convertTtoDelText(doc, child as Element);
    }
  }
}

// ─────────── ensureParaMarkInsert / ensureParaMarkDelete ───────────

/**
 * 在段落的 `<w:pPr><w:rPr>` 里插入 `<w:ins w:id="N" .../>`,
 * 表示该段落 mark 是新插入的(新段的修订标记)。
 *
 * - 如果 pPr 不存在 → 创建 pPr 并插到段落首
 * - 如果 pPr 存在但无 rPr → 创建 rPr 并插到 pPr 首
 * - 如果 rPr 已存在 → 直接 append ins
 */
export function ensureParaMarkInsert(paragraph: Element, id: number, attrs: RevisionAttrs): void {
  const doc = paragraph.ownerDocument;
  if (!doc) throw new Error('ensureParaMarkInsert: paragraph has no ownerDocument');

  const rPr = ensureRPrInPPr(doc, paragraph);
  const ins = createW(doc, 'ins');
  setRevisionAttrs(ins, id, attrs);
  rPr.appendChild(ins);
}

/**
 * 在段落的 `<w:pPr><w:rPr>` 里插入 `<w:del w:id="N" .../>`,
 * 表示该段落 mark 是被删除的(deleteParagraph trackChanges 下的段落 mark 标记)。
 */
export function ensureParaMarkDelete(paragraph: Element, id: number, attrs: RevisionAttrs): void {
  const doc = paragraph.ownerDocument;
  if (!doc) throw new Error('ensureParaMarkDelete: paragraph has no ownerDocument');

  const rPr = ensureRPrInPPr(doc, paragraph);
  const del = createW(doc, 'del');
  setRevisionAttrs(del, id, attrs);
  rPr.appendChild(del);
}

/** 确保段落有 `<w:pPr>`,并在 pPr 里确保有 `<w:rPr>`,返回 rPr 元素 */
function ensureRPrInPPr(doc: Document, paragraph: Element): Element {
  // 找或建 pPr
  let pPr = firstDirectChildW(paragraph, 'pPr');
  if (!pPr) {
    pPr = createW(doc, 'pPr');
    // pPr 必须是段落第一个子
    paragraph.insertBefore(pPr, paragraph.firstChild);
  }

  // 找或建 rPr
  let rPr = firstDirectChildW(pPr, 'rPr');
  if (!rPr) {
    rPr = createW(doc, 'rPr');
    pPr.insertBefore(rPr, pPr.firstChild);
  }

  return rPr;
}

/** 找 `el` 直属子中首个 localName 匹配的 W_NS 元素(null 表示无)*/
function firstDirectChildW(el: Element, localName: string): Element | null {
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
