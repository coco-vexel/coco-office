/**
 * Revision(修订)只读查询(docs/03.1 §1.5 / §2,#26 + #33)。
 *
 * **Tier 1 覆盖面**:只看 `<w:ins>` / `<w:del>` —— 这是 run 级的插入/删除修订,
 * 承载 w:id / w:author / w:date 属性,以及子树内的 `<w:t>`(ins)/ `<w:delText>`(del)
 * 作为涉及文本。
 *
 * **Tier 2 属性级修订**(#33 起显式拒绝):
 *   - `<w:rPrChange>`(run 属性变更,在 `<w:rPr>` 里)
 *   - `<w:pPrChange>`(段落属性变更,在 `<w:pPr>` 里)
 *   - `<w:sectPrChange>` / `<w:tblPrChange>` / `<w:trPrChange>` / `<w:tcPrChange>` 等
 *   - `<w:numberingChange>` / `<w:moveFrom>` / `<w:moveTo>` 等
 *
 * 扫到任一 Tier 2 元素即抛 `UNSUPPORTED_OOXML_FEATURE`(`listRevisions` 与
 * `acceptRevisions` / `rejectRevisions` 共用此守护,保证读写一致)。
 *
 * scope 控制枚举的 story 集合(与 `listParagraphs` 共用 `resolveScope`)。
 */
import type { Package } from '../pack/index.js';
import { W14_NS, W_NS } from './namespaces.js';
import { resolveScope, type StoryId } from './stories.js';
import { WordError } from './word-error.js';

export type RevisionId = string;

export interface RevisionInfo {
  revId: RevisionId;
  kind: 'insert' | 'delete';
  author: string;
  date: string;
  story: StoryId;
  /** 修订归属的段落(若 ins/del 嵌在 `<w:p>` 里才有;极少数嵌在 `<w:body>` 直子层位的则缺) */
  paraId?: string;
  /** 涉及的文本(ins:全部 `<w:t>` 拼接;del:全部 `<w:delText>` 拼接)。为空串时省略 */
  text?: string;
}

export interface ListRevisionsOptions {
  /** 默认 'body' */
  scope?: 'body' | 'all' | StoryId[];
}

// ─────────── Tier 2 属性修订守护(#33,读/写共用)─────────────

/**
 * Tier 2 属性级修订 local name 集合。
 * 扫到任一即抛 `UNSUPPORTED_OOXML_FEATURE`(`listRevisions` 与 `accept/rejectRevisions`
 * 共用此守护,保证只读/只写对 Tier 2 的态度一致)。
 */
export const TIER2_REVISION_LOCAL_NAMES: readonly string[] = [
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'tcPrChange',
  'trPrChange',
  'tblPrExChange',
  'tblGridChange',
  'numberingChange',
  'moveFrom',
  'moveTo',
];

/**
 * 扫 `root`(Document 或 Element 子树)里是否存在 Tier 2 属性修订元素;
 * 发现一个立即抛 `WordError('UNSUPPORTED_OOXML_FEATURE')`。
 *
 * `docIdForError` 用于错误注入。调用方在读(listRevisions)/ 写(accept/reject)
 * 路径上各自提供自己的 docId。
 */
export function assertNoTier2Revisions(
  root: Document | Element,
  docIdForError: string,
  opContext: 'listRevisions' | 'acceptRevisions' | 'rejectRevisions',
): void {
  for (const local of TIER2_REVISION_LOCAL_NAMES) {
    const els = root.getElementsByTagNameNS(W_NS, local);
    if (els.length > 0) {
      throw new WordError(
        'UNSUPPORTED_OOXML_FEATURE',
        `Tier 2 revision '<w:${local}>' is not supported (op: ${opContext}); ` +
          `only <w:ins> / <w:del> are covered in Tier 1`,
        docIdForError,
        { revisionKind: local, opName: opContext },
      );
    }
  }
}

// ─────────── 公共入口 ───────────

export function listRevisions(pkg: Package, opts: ListRevisionsOptions = {}): RevisionInfo[] {
  const stories = resolveScope(pkg, opts.scope ?? 'body');
  const out: RevisionInfo[] = [];

  for (const story of stories) {
    // Tier 2 守护(#33):扫到 rPrChange / pPrChange / 等 → 抛
    // docId 本处没有(`listRevisions` 是 Package 级 helper,不知道 WordDocument.docId);
    // 用占位 '' 让 kernel 层的 toWordError 做二次注入。
    assertNoTier2Revisions(story.root, '', 'listRevisions');

    // 直接在 story root 上 getElementsByTagNameNS —— 只取该 story 子树内的 ins/del,
    // 不跨 sub-story(footnote 根单独一个 story,这里会在 footnote root 上再扫一轮)
    for (const { el, kind } of collectRevisionElementsInDocumentOrder(story.root)) {
      const info = buildRevisionInfo(el, kind, story.id);
      if (info) out.push(info);
    }
  }

  return out;
}

// ─────────── 内部 ───────────

function buildRevisionInfo(
  el: Element,
  kind: 'insert' | 'delete',
  story: StoryId,
): RevisionInfo | null {
  const revId = el.getAttributeNS(W_NS, 'id');
  if (revId === null) return null; // 缺 id 的 ins/del 跳过 —— 无法稳定引用

  const author = el.getAttributeNS(W_NS, 'author') ?? '';
  const date = el.getAttributeNS(W_NS, 'date') ?? '';
  const paraId = findContainingParaId(el);
  const text = extractRevisionText(el, kind);

  const info: RevisionInfo = { revId, kind, author, date, story };
  if (paraId !== undefined) info.paraId = paraId;
  if (text.length > 0) info.text = text;
  return info;
}

/** 从 ins/del 元素向上走到最近的 `<w:p>`,返回其 `w14:paraId`(大写);没有则返回 undefined */
function findContainingParaId(el: Element): string | undefined {
  let cur: Node | null = el.parentNode;
  while (cur && cur.nodeType === 1) {
    const e = cur as Element;
    if (e.namespaceURI === W_NS && e.localName === 'p') {
      const raw = e.getAttributeNS(W14_NS, 'paraId');
      return raw && raw.length > 0 ? raw.toUpperCase() : undefined;
    }
    cur = cur.parentNode;
  }
  return undefined;
}

/**
 * 涉及文本:
 *   - insert: 后代 `<w:t>`
 *   - delete: 后代 `<w:delText>`
 *
 * 说明:`<w:ins>` 可能嵌套内容 wrapper(如 `<w:ins><w:r><w:t>` 或
 * `<w:ins><w:hyperlink><w:r><w:t>`)—— 统一用 getElementsByTagNameNS 吃全部后代。
 */
function extractRevisionText(el: Element, kind: 'insert' | 'delete'): string {
  const tag = kind === 'insert' ? 't' : 'delText';
  const nodes = el.getElementsByTagNameNS(W_NS, tag);
  let out = '';
  for (let i = 0; i < nodes.length; i++) {
    out += nodes[i]!.textContent ?? '';
  }
  return out;
}

function collectRevisionElementsInDocumentOrder(
  root: Document | Element,
): Array<{ el: Element; kind: 'insert' | 'delete' }> {
  const out: Array<{ el: Element; kind: 'insert' | 'delete' }> = [];

  function visit(node: Node): void {
    if (node.nodeType === 1) {
      const el = node as Element;
      if (el.namespaceURI === W_NS) {
        if (el.localName === 'ins') {
          out.push({ el, kind: 'insert' });
          return;
        }
        if (el.localName === 'del') {
          out.push({ el, kind: 'delete' });
          return;
        }
      }
    }

    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child) visit(child);
    }
  }

  visit(root);
  return out;
}
