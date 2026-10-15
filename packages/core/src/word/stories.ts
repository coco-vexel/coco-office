/**
 * Word 跨 stories 枚举(docs/03 §10)。
 *
 * 一个 Word 文档可能包含 6 类 story:
 *   | 类型     | contentType 尾缀         | 根元素                            | 数量 |
 *   | body     | .document.main+xml       | <w:document>/<w:body>             | 1    |
 *   | header   | .header+xml              | <w:hdr>                           | 0..N |
 *   | footer   | .footer+xml              | <w:ftr>                           | 0..N |
 *   | footnote | .footnotes+xml 内每个    | 单个 <w:footnote>(该 part 多个)  | 0..M |
 *   | endnote  | .endnotes+xml 内每个     | 单个 <w:endnote>                  | 0..M |
 *   | comment  | .comments+xml 内每个     | 单个 <w:comment>                  | 0..M |
 *
 * **关键约束**:枚举实现**必须**走 `Document#getElementsByTagNameNS`。
 * `fontoxpath` 的 `//w:p` 无谓词 XPath 在 10 万段落上慢 460×(见 bench/xml-stack-worker)。
 * CI 的 `pnpm check:xpath` 会在源码里扫到此类调用即失败(scripts/check-no-unqualified-xpath.mjs)。
 */
import type { Package } from '../pack/index.js';
import { isXmlPart } from '../pack/index.js';
import { W_NS } from './namespaces.js';

export type StoryKind = 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'comment';

/**
 * Story 唯一标识:
 *   - 'body' — 唯一主 body
 *   - 'header:<part path>' / 'footer:<part path>' — 1 part 1 story
 *   - 'footnote:<id>' / 'endnote:<id>' / 'comment:<id>' — 多 sub-story 按 w:id 区分
 */
export type StoryId =
  | 'body'
  | `header:${string}`
  | `footer:${string}`
  | `footnote:${number}`
  | `endnote:${number}`
  | `comment:${number}`;

export interface Story {
  id: StoryId;
  kind: StoryKind;
  /** story 的根 Element(body = <w:body>,footnote = 单个 <w:footnote>,依此类推) */
  root: Element;
  /** 所在 part 的 path */
  partPath: string;
}

export interface ListStoriesOptions {
  /** 仅枚举指定 kinds(默认全枚举)*/
  kinds?: StoryKind[];
}

/** 把 HTMLCollection-ish 枚举结果快照成数组 */
function snapshot<T extends Element>(col: { length: number } & Record<number, T>): T[] {
  const out: T[] = [];
  for (let i = 0; i < col.length; i++) {
    const el = col[i];
    if (el) out.push(el);
  }
  return out;
}

/**
 * 枚举 Package 内所有 stories。
 *
 * 按 contentType 尾缀识别 part,内部用 `getElementsByTagNameNS` 提取根元素或
 * sub-story 元素。**不做**任何带谓词/路径的 XPath。
 */
export function listStories(pkg: Package, opts: ListStoriesOptions = {}): Story[] {
  const wanted = new Set<StoryKind>(
    opts.kinds ?? ['body', 'header', 'footer', 'footnote', 'endnote', 'comment'],
  );
  const out: Story[] = [];

  for (const part of pkg.parts.values()) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    const ct = part.contentType;

    if (
      wanted.has('body') &&
      (ct.endsWith('.document.main+xml') ||
        ct.endsWith('.document.macroEnabled.main+xml') ||
        ct.endsWith('.template.main+xml') ||
        ct.endsWith('.template.macroEnabledTemplate.main+xml'))
    ) {
      // body 是 <w:document>/<w:body> 的唯一子元素
      const bodies = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'body'));
      if (bodies.length > 0) {
        out.push({ id: 'body', kind: 'body', root: bodies[0]!, partPath: part.path });
      }
      continue;
    }

    if (wanted.has('header') && ct.endsWith('.header+xml')) {
      const hdrs = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'hdr'));
      if (hdrs.length > 0) {
        out.push({
          id: `header:${part.path}`,
          kind: 'header',
          root: hdrs[0]!,
          partPath: part.path,
        });
      }
      continue;
    }

    if (wanted.has('footer') && ct.endsWith('.footer+xml')) {
      const ftrs = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'ftr'));
      if (ftrs.length > 0) {
        out.push({
          id: `footer:${part.path}`,
          kind: 'footer',
          root: ftrs[0]!,
          partPath: part.path,
        });
      }
      continue;
    }

    if (wanted.has('footnote') && ct.endsWith('.footnotes+xml')) {
      const notes = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'footnote'));
      for (const note of notes) {
        const id = numAttr(note, W_NS, 'id');
        if (id === null) continue;
        out.push({
          id: `footnote:${id}`,
          kind: 'footnote',
          root: note,
          partPath: part.path,
        });
      }
      continue;
    }

    if (wanted.has('endnote') && ct.endsWith('.endnotes+xml')) {
      const notes = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'endnote'));
      for (const note of notes) {
        const id = numAttr(note, W_NS, 'id');
        if (id === null) continue;
        out.push({
          id: `endnote:${id}`,
          kind: 'endnote',
          root: note,
          partPath: part.path,
        });
      }
      continue;
    }

    if (wanted.has('comment') && ct.endsWith('.comments+xml')) {
      const comments = snapshot(part.xmlDom.getElementsByTagNameNS(W_NS, 'comment'));
      for (const c of comments) {
        const id = numAttr(c, W_NS, 'id');
        if (id === null) continue;
        out.push({
          id: `comment:${id}`,
          kind: 'comment',
          root: c,
          partPath: part.path,
        });
      }
    }
  }

  return out;
}

/**
 * 取 story 根下的所有 <w:p> 元素。
 *
 * - body / header / footer:story root 整个子树的 w:p
 * - footnote / endnote / comment:单个 sub-story root 下的 w:p(不跨 sub-story)
 */
export function getStoryParagraphs(story: Story): Element[] {
  // 使用 getElementsByTagNameNS 枚举 —— 根元素调用时只返回该子树下的匹配
  return snapshot(story.root.getElementsByTagNameNS(W_NS, 'p'));
}

/**
 * 反查元素所属 story(#25)。用 `Node.contains()` 逐个比对,命中首个返回。
 *
 * 一个 pkg 内 stories 数量通常 <20,contains 是 ~O(depth) 的祖先遍历,整体代价
 * 可忽略。找不到(元素不在任何已识别 story 的子树里)返回 null 。
 */
export function findStoryFor(pkg: Package, el: Element): Story | null {
  for (const s of listStories(pkg)) {
    if (s.root.contains(el as unknown as Node)) return s;
  }
  return null;
}

/**
 * scope 枚举解析为 Story 列表(#24 / #26 共用)。
 *
 * - `'body'` —— 只枚举 body
 * - `'all'` —— 全部 6 类 story
 * - `StoryId[]` —— 精确筛选;**未命中的 id 忽略**(调用方自己负责正确性)
 */
export function resolveScope(pkg: Package, scope: 'body' | 'all' | StoryId[]): Story[] {
  if (scope === 'body') return listStories(pkg, { kinds: ['body'] });
  if (scope === 'all') return listStories(pkg);
  const wanted = new Set<StoryId>(scope);
  return listStories(pkg).filter((s) => wanted.has(s.id));
}

function numAttr(el: Element, ns: string, local: string): number | null {
  const v = el.getAttributeNS(ns, local);
  if (v === null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
