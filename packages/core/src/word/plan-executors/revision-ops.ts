/**
 * acceptRevisions / rejectRevisions executor —— #32(Phase D 收尾)。
 *
 * 契约(docs/03.1 §3.1):
 *   - 三档 scope:all / paragraph(paraId) / revision(revId)
 *   - Accept `<w:ins>` wrapper:剥 wrapper,保留内容
 *   - Accept `<w:del>` wrapper:剥 wrapper + 删被删内容
 *   - Reject 语义相反
 *   - 段落 mark(`<w:pPr><w:rPr><w:ins/>` / `<w:del/>`)一并处理:
 *     * Accept `ins` mark:删 mark,段保留
 *     * Accept `del` mark:**删整段**(段 mark 的 del 简化为"段要被删")
 *     * Reject `ins` mark:**删整段**
 *     * Reject `del` mark:删 mark,段保留
 *   - `<w:rPrChange>` / `<w:pPrChange>` / `<w:sectPrChange>` 等 Tier 2 属性级修订 →
 *     抛 `UNSUPPORTED_OOXML_FEATURE`
 *
 * **段落 mark 简化说明**:OOXML 规范里段落 mark 的修订对应段落**合并 / 分裂**语义
 * (段 mark ins = 分裂点新增;段 mark del = 与下段合并)。本 PR 的 Tier 1 实现
 * 把"段 mark del accept / 段 mark ins reject"简化为"删整段",不做文本层面的
 * 段落合并。Tier 2 再精细化。
 *
 * 每处理一个 wrapper 或段落 mark 计入 ApplyResult.acceptedRevisions /
 * rejectedRevisions 一次。
 */
import { getPart, isXmlPart, listParts, markDirty } from '../../pack/index.js';
import { W_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { assertNoTier2Revisions } from '../revision-queries.js';
import { findStoryFor } from '../stories.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';

type Action = 'accept' | 'reject';
type RevisionKind = 'ins' | 'del';

interface RevisionTarget {
  el: Element;
  kind: RevisionKind;
  /** true 表示此 ins/del 是段落 mark(位于 pPr/rPr 下) */
  isParaMark: boolean;
  /** 段落 mark 对应的 `<w:p>` 元素(isParaMark=false 时也可能用到,作为 containing paragraph) */
  paragraph: Element | null;
}

// Tier 2 属性级修订拒绝:用共享的 `assertNoTier2Revisions`(revision-queries.ts),
// 读 listRevisions + 写 accept/reject 共用同一名单(#33)。

// ─────────── entry ───────────

export function executeAcceptRevisions(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'acceptRevisions' }>,
): void {
  processRevisions(ctx, op.scope, 'accept');
}

export function executeRejectRevisions(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'rejectRevisions' }>,
): void {
  processRevisions(ctx, op.scope, 'reject');
}

function processRevisions(
  ctx: ExecutorContext,
  scope: Extract<Op, { type: 'acceptRevisions' }>['scope'],
  action: Action,
): void {
  // 1. 收集 targets(根据 scope);同时扫 Tier 2 属性修订抛 UNSUPPORTED
  const targets = collectTargets(ctx, scope);

  // 2. 逐个处理;paragraph mark 要删整段的分支,**把对应 <w:p> 收集到批量删除队列**,
  //    避免同一段上多个 ins/del 导致多次尝试 removeChild
  const paragraphsToDelete = new Set<Element>();
  const markedParts = new Set<string>();
  const addDirty = (el: Element): void => {
    const story = findStoryFor(ctx.pkg, el);
    if (story) markedParts.add(story.partPath);
  };

  for (const t of targets) {
    // 先 markDirty(基于处理前 t.el 仍挂在 DOM 上的 parent chain),再应用变更 ——
    // applyWrapper 会 removeChild 导致 t.el.parentNode 变 null,太晚查不到 story
    if (t.paragraph) addDirty(t.paragraph);
    applyToTarget(t, action, paragraphsToDelete, addDirty);
    if (action === 'accept') ctx.result.acceptedRevisions++;
    else ctx.result.rejectedRevisions++;
  }

  // 3. 统一删除段落(去重后)
  for (const p of paragraphsToDelete) {
    addDirty(p);
    p.parentNode?.removeChild(p);
  }

  // 4. markDirty 所有触达的 parts
  for (const partPath of markedParts) {
    const part = getPart(ctx.pkg, partPath);
    if (part) markDirty(part);
  }
}

// ─────────── target 收集 ───────────

function collectTargets(
  ctx: ExecutorContext,
  scope: Extract<Op, { type: 'acceptRevisions' }>['scope'],
): RevisionTarget[] {
  if (scope.scope === 'all') {
    return collectFromPkg(ctx);
  }
  if (scope.scope === 'paragraph') {
    const p = ctx.indices.paraIdx.get(scope.paraId.toUpperCase());
    if (!p) {
      throw new WordError(
        'PARA_NOT_FOUND',
        `paragraph with paraId='${scope.paraId}' not found`,
        ctx.docId,
        { paraId: scope.paraId, opName: 'accept/rejectRevisions' },
      );
    }
    return collectFromSubtree(ctx, p);
  }
  // scope === 'revision'
  const all = collectFromPkg(ctx);
  const matched = all.filter((t) => {
    const id = t.el.getAttributeNS(W_NS, 'id') ?? t.el.getAttribute('w:id');
    return id === scope.revId;
  });
  if (matched.length === 0) {
    throw new WordError(
      'REVISION_NOT_FOUND',
      `revision with revId='${scope.revId}' not found`,
      ctx.docId,
      { revId: scope.revId, opName: 'accept/rejectRevisions' },
    );
  }
  return matched;
}

function collectFromPkg(ctx: ExecutorContext): RevisionTarget[] {
  const out: RevisionTarget[] = [];
  for (const part of listParts(ctx.pkg)) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    const doc = part.xmlDom as Document;
    assertNoTier2Revisions(doc, ctx.docId, 'acceptRevisions');
    collectInsDelFromRoot(doc, out);
  }
  return out;
}

function collectFromSubtree(ctx: ExecutorContext, root: Element): RevisionTarget[] {
  assertNoTier2Revisions(root, ctx.docId, 'acceptRevisions');
  const out: RevisionTarget[] = [];
  collectInsDelFromRoot(root, out);
  return out;
}

function collectInsDelFromRoot(root: Document | Element, out: RevisionTarget[]): void {
  const insEls = root.getElementsByTagNameNS(W_NS, 'ins');
  const delEls = root.getElementsByTagNameNS(W_NS, 'del');
  for (let i = 0; i < insEls.length; i++) {
    const el = insEls[i];
    if (!el) continue;
    out.push(buildTarget(el, 'ins'));
  }
  for (let i = 0; i < delEls.length; i++) {
    const el = delEls[i];
    if (!el) continue;
    out.push(buildTarget(el, 'del'));
  }
}

function buildTarget(el: Element, kind: RevisionKind): RevisionTarget {
  const parent = el.parentNode;
  // 段落 mark 检测:parent 是 rPr,rPr 的 parent 是 pPr,pPr 的 parent 是 w:p
  let isParaMark = false;
  let paragraph: Element | null = null;
  if (
    parent &&
    parent.nodeType === 1 &&
    (parent as Element).namespaceURI === W_NS &&
    (parent as Element).localName === 'rPr'
  ) {
    const pPr = parent.parentNode;
    if (
      pPr &&
      pPr.nodeType === 1 &&
      (pPr as Element).namespaceURI === W_NS &&
      (pPr as Element).localName === 'pPr'
    ) {
      const p = pPr.parentNode;
      if (
        p &&
        p.nodeType === 1 &&
        (p as Element).namespaceURI === W_NS &&
        (p as Element).localName === 'p'
      ) {
        isParaMark = true;
        paragraph = p as Element;
      }
    }
  }
  if (!isParaMark) {
    paragraph = findContainingParagraph(el);
  }
  return { el, kind, isParaMark, paragraph };
}

function findContainingParagraph(el: Element): Element | null {
  let cur: Node | null = el.parentNode;
  while (cur && cur.nodeType === 1) {
    const e = cur as Element;
    if (e.namespaceURI === W_NS && e.localName === 'p') return e;
    cur = cur.parentNode;
  }
  return null;
}

// assertNoPropertyRevisions 已在 #33 提到 revision-queries.ts,更名为 assertNoTier2Revisions
// 以读 / 写路径共用

// ─────────── 单 target 应用 ───────────

function applyToTarget(
  t: RevisionTarget,
  action: Action,
  paragraphsToDelete: Set<Element>,
  _addDirty: (el: Element) => void,
): void {
  if (t.isParaMark) {
    applyParaMark(t, action, paragraphsToDelete);
    return;
  }
  // 正文 wrapper(markDirty 已在上层循环基于 t.paragraph 做完)
  applyWrapper(t, action);
}

/** 段落 mark 处理 */
function applyParaMark(t: RevisionTarget, action: Action, paragraphsToDelete: Set<Element>): void {
  const p = t.paragraph!;
  const shouldDeleteParagraph =
    (t.kind === 'ins' && action === 'reject') || (t.kind === 'del' && action === 'accept');
  if (shouldDeleteParagraph) {
    paragraphsToDelete.add(p);
  } else {
    // 仅删 ins/del mark 标记,段保留
    t.el.parentNode?.removeChild(t.el);
  }
}

/** 正文 wrapper(`<w:ins>` / `<w:del>` 直接包裹 runs)处理 */
function applyWrapper(t: RevisionTarget, action: Action): void {
  const el = t.el;
  const parent = el.parentNode;
  if (!parent) return;

  const keepContent =
    (t.kind === 'ins' && action === 'accept') || (t.kind === 'del' && action === 'reject');

  if (keepContent) {
    // 剥 wrapper,把内部 children 移到 wrapper 原位
    // del wrapper 的 reject:内部 <w:delText> 要转回 <w:t>
    const children = Array.from(el.childNodes);
    if (t.kind === 'del' && action === 'reject') {
      for (const c of children) {
        if (c.nodeType === 1) convertDelTextToT(c as Element);
      }
    }
    for (const c of children) {
      parent.insertBefore(c, el);
    }
    parent.removeChild(el);
  } else {
    // 整个 wrapper 连带内容一起删
    parent.removeChild(el);
  }
}

/** 递归把 `<w:delText>` 转回 `<w:t>`(reject del 用);属性搬 xml:space */
function convertDelTextToT(el: Element): void {
  // 先递归处理子孙(深度优先,避免枚举中修改)
  const snapshot: Node[] = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (c) snapshot.push(c);
  }
  for (const child of snapshot) {
    if (
      child.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName === 'delText'
    ) {
      const delText = child as Element;
      const doc = delText.ownerDocument!;
      const t = doc.createElementNS(W_NS, 'w:t');
      // 搬属性(主要是 xml:space="preserve")
      for (let i = 0; i < delText.attributes.length; i++) {
        const attr = delText.attributes[i];
        if (!attr) continue;
        if (attr.namespaceURI) {
          t.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
        } else {
          t.setAttribute(attr.name, attr.value);
        }
      }
      t.textContent = delText.textContent;
      el.replaceChild(t, delText);
    } else if (child.nodeType === 1) {
      convertDelTextToT(child as Element);
    }
  }
}
