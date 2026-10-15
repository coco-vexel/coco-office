/**
 * rangeKind 判定共享逻辑(#25 bookmark + #26 comment)。
 *
 * BookmarkInfo / CommentInfo 都有同款 `rangeKind`(paragraph / text / complex)语义,
 * 判定依赖一对 "start / end" marker 元素在 DOM 中的相对位置。本模块把这个判定 +
 * 区间文本收集抽出来,两端共享。
 *
 * 判定规则(docs/03.1 §1.5):
 *   1. end 为 null → 'complex'(缺失 end)
 *   2. start.parent ≠ end.parent → 'complex'(跨节点)
 *   3. parent 不是 `<w:p>` → 'complex'(e.g. `<w:tc>` / `<w:body>` 直挂)
 *   4. 同一 `<w:p>` 内:
 *      - start 之前 / end 之后**没有**内容节点 → 'paragraph'
 *      - 否则 → 'text'(text 字段 = start/end 间所有 `<w:t>` 拼接)
 *
 * "内容节点" = W_NS 下非 marker 元素(marker 集:pPr / rPr / bookmarkStart /
 * bookmarkEnd / commentRangeStart / commentRangeEnd / commentReference / proofErr /
 * permStart / permEnd)+ 任意非 W_NS 元素(保守视为内容,避免未知扩展误判成
 * paragraph 覆盖)。
 */
import { W14_NS, W_NS } from './namespaces.js';

export interface RangeClassification {
  kind: 'paragraph' | 'text' | 'complex';
  /** 'paragraph' / 'text' 命中时,containing `<w:p>` 的 paraId(大写 8 hex) */
  paraId?: string;
  /** 'text' 命中时,start/end 之间所有 `<w:t>` 的纯文本拼接 */
  text?: string;
}

export function classifyRange(start: Element | null, end: Element | null): RangeClassification {
  if (!start || !end) return { kind: 'complex' };

  const sp = start.parentNode;
  const ep = end.parentNode;
  if (!sp || sp !== ep) return { kind: 'complex' };

  const parent = sp as Element;
  if (parent.namespaceURI !== W_NS || parent.localName !== 'p') {
    return { kind: 'complex' };
  }

  const paraIdRaw = parent.getAttributeNS(W14_NS, 'paraId');
  const paraId = paraIdRaw && paraIdRaw.length > 0 ? paraIdRaw.toUpperCase() : undefined;

  const before = hasContentBefore(parent, start);
  const after = hasContentAfter(parent, end);

  if (!before && !after) {
    const info: RangeClassification = { kind: 'paragraph' };
    if (paraId !== undefined) info.paraId = paraId;
    return info;
  }

  const text = collectTextBetween(start, end);
  const info: RangeClassification = { kind: 'text', text };
  if (paraId !== undefined) info.paraId = paraId;
  return info;
}

// ─────────── 内部:内容 / 文本 ───────────

/**
 * `<w:p>` 直属 W_NS 子元素中,**不算内容**的 marker/meta 节点集。
 * 未知的 W_NS 元素 **不在此集中** → 被保守视为内容。
 */
const W_MARKER_LOCAL_NAMES = new Set<string>([
  'pPr',
  'rPr',
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'commentReference',
  'proofErr',
  'permStart',
  'permEnd',
]);

function isContentChild(n: Node | null): boolean {
  if (!n || n.nodeType !== 1) return false; // 文本 / 注释 节点不算内容
  const el = n as Element;
  if (el.namespaceURI === W_NS) {
    if (W_MARKER_LOCAL_NAMES.has(el.localName ?? '')) return false;
    // "marker run":一个 `<w:r>` 只包 rPr + marker 元素(典型:`<w:r><w:commentReference/>
    // </w:r>` 放在 commentRangeEnd 之后)—— 不算内容,否则所有含批注引用的段落都会
    // 被判成 'text' rangeKind。
    if (el.localName === 'r' && isMarkerOnlyRun(el)) return false;
    return true;
  }
  // 非 W_NS(mc:AlternateContent / drawing 等):保守当作内容
  return true;
}

/**
 * `<w:r>` 的子只含 rPr / commentReference / bookmarkStart / bookmarkEnd /
 * commentRangeStart / commentRangeEnd 时视为 marker-only(run 不贡献内容)。
 *
 * 任何其它 W_NS 子(w:t / w:br / w:tab / w:sym / w:drawing / w:pict / w:fldChar /
 * w:delText ...)或非 W_NS 子都视为 run 有内容。
 */
function isMarkerOnlyRun(run: Element): boolean {
  const children = run.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (!c || c.nodeType !== 1) continue; // 文本 / 注释节点忽略(空白不算内容)
    const el = c as Element;
    if (el.namespaceURI !== W_NS) return false;
    const local = el.localName ?? '';
    if (local === 'rPr') continue;
    if (W_MARKER_LOCAL_NAMES.has(local)) continue;
    return false;
  }
  return true;
}

function hasContentBefore(parent: Element, target: Element): boolean {
  const children = parent.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i]!;
    if (c === target) return false;
    if (isContentChild(c)) return true;
  }
  return false;
}

function hasContentAfter(parent: Element, target: Element): boolean {
  const children = parent.childNodes;
  let seen = false;
  for (let i = 0; i < children.length; i++) {
    const c = children[i]!;
    if (seen && isContentChild(c)) return true;
    if (c === target) seen = true;
  }
  return false;
}

/**
 * 收集 start..end(都 exclusive)之间所有 `<w:t>` 文本拼接。
 *
 * 区间内兄弟可能是 `<w:r>`(递归含 `<w:t>`)/ `<w:hyperlink>` / `<w:ins>` 等包装
 * 元素 —— 统一调 `getElementsByTagNameNS(W_NS, 't')` 吃全部后代 `<w:t>`。marker
 * 节点本身无 `<w:t>` 子,不会被计入。
 */
function collectTextBetween(start: Element, end: Element): string {
  const parent = start.parentNode;
  if (!parent) return '';
  const children = parent.childNodes;
  let inRange = false;
  let out = '';
  for (let i = 0; i < children.length; i++) {
    const c = children[i]!;
    if (c === start) {
      inRange = true;
      continue;
    }
    if (c === end) break;
    if (inRange && c.nodeType === 1) {
      const el = c as Element;
      const ts = el.getElementsByTagNameNS(W_NS, 't');
      for (let j = 0; j < ts.length; j++) {
        out += ts[j]!.textContent ?? '';
      }
    }
  }
  return out;
}
