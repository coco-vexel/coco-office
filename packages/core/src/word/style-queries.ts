/**
 * style-queries —— styles.xml 解析与样式索引构建（#181）。
 *
 * 职责：
 *   - `buildStylesIndex(pkg)`：扫 Package 找 styles.xml，构建 Map<styleId, StyleInfo>
 *   - `listStyles(stylesIdx, opts?)`：按可选 type 过滤后返回数组
 *   - `getStyleById(stylesIdx, styleId)`：O(1) 精确查找
 *
 * 健壮性契约：
 *   - styles.xml 不存在 → 返回空 Map，不抛错
 *   - `w:type` 未知值 → fallback 到 'paragraph'（OOXML 规范默认）
 *   - `w:name` 子元素缺失 → name fallback 到 styleId 本身
 *   - 同一 styleId 重复定义 → 后者覆盖前者（与 Word 行为一致）
 *   - basedOn 环路 → SDK 不检测，消费方遍历时须自行防环
 *
 * 不暴露 `builtIn` 字段：`<w:builtIn/>` 元素在实际文档中不一致（Word 某些版本省略），
 * 单向可信（存在则肯定是内置，不存在不代表是自定义）会误导消费方。
 * 消费方应通过 `name` + basedOn 链判断语义角色。
 */
import type { Package } from '../pack/index.js';
import { W_NS } from './namespaces.js';

// ─────────── 类型 ───────────

/** OOXML 定义的四种样式类型（`w:style@w:type`）。 */
export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

export interface StyleInfo {
  /** `w:style@w:styleId`。段落 `w:pStyle@w:val` 引用的 key，大小写敏感。 */
  styleId: string;
  /**
   * `w:name@w:val`。
   * - OOXML 规范：内置样式（如 Heading 1、Normal）的 name 永远是英文规范名，
   *   不随文档语言变化。消费方可据此判断语义角色。
   * - 自定义样式：name 为用户输入，可能是中文或任意字符串。
   * - 若 `<w:name>` 元素缺失（非规范文档），fallback 到 `styleId`。
   */
  name: string;
  /**
   * `w:style@w:type`。
   * 未知值 fallback 到 `'paragraph'`（OOXML 规范默认类型）。
   */
  type: StyleType;
  /**
   * 父样式的 `styleId`（`w:basedOn@w:val`）。
   * 根样式（如 Normal）无继承关系，此字段不存在。
   *
   * ⚠ 消费方遍历 basedOn 链时须自行防环，SDK 不检测循环继承。
   */
  basedOn?: string;
  /**
   * 样式定义中的大纲级别（`w:style/w:pPr/w:outlineLvl@w:val`），0–8。
   * 0 = 标题 1，1 = 标题 2，…，8 = 标题 9；9 表示正文（不视为标题）。
   *
   * WPS / Word 的自定义标题样式（如"公文-一级标题"）通常不 basedOn heading 1，
   * 而是在此字段直接编码级别。消费方可将其作为 basedOn 链之外的第二信号。
   *
   * 字段不存在 = 样式未设置大纲级别（继承自父样式或使用默认值）。
   */
  outlineLevel?: number;
}

export interface ListStylesOptions {
  /** 按样式类型过滤。不传则返回全部样式（含所有类型）。 */
  type?: StyleType;
}

// ─────────── styles.xml 的标准 ContentType ───────────

const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';

// ─────────── 内部工具 ───────────

/** 按 contentType 定位 styles.xml Part（不硬编码路径）。 */
function findStylesPart(pkg: Package): { xmlDom: Document } | undefined {
  for (const part of pkg.parts.values()) {
    if (part.contentType === STYLES_CONTENT_TYPE && part.xmlDom) {
      return part as { xmlDom: Document };
    }
  }
  return undefined;
}

/** 取第一个直接子元素（指定 namespace + localName）的某属性值。 */
function childAttr(
  el: Element,
  ns: string,
  localName: string,
  attrNs: string,
  attrLocal: string,
): string | null {
  const children = el.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === ns &&
      (c as Element).localName === localName
    ) {
      return (c as Element).getAttributeNS(attrNs, attrLocal);
    }
  }
  return null;
}

/**
 * 从 `<w:style>` 元素读取其 pPr 中的 outlineLvl。
 * 返回 0–8 的数字，或 undefined（未设置 / 值为 9 表示正文不视为标题）。
 */
function styleOutlineLevel(styleEl: Element): number | undefined {
  const children = styleEl.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (!c || c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI !== W_NS || el.localName !== 'pPr') continue;

    // 找到 pPr，再找 outlineLvl
    const pPrChildren = el.childNodes;
    for (let j = 0; j < pPrChildren.length; j++) {
      const pc = pPrChildren[j];
      if (!pc || pc.nodeType !== 1) continue;
      const pcel = pc as Element;
      if (pcel.namespaceURI !== W_NS || pcel.localName !== 'outlineLvl') continue;

      const raw = pcel.getAttributeNS(W_NS, 'val');
      if (raw === null) return undefined;
      const n = parseInt(raw, 10);
      // 0–8 才是标题级别；9 表示正文，不视为标题
      return Number.isFinite(n) && n >= 0 && n <= 8 ? n : undefined;
    }
  }
  return undefined;
}

/** `w:type` 原始字符串 → StyleType（未知值 fallback 到 'paragraph'）。 */
function parseStyleType(raw: string | null): StyleType {
  switch (raw) {
    case 'paragraph':
      return 'paragraph';
    case 'character':
      return 'character';
    case 'table':
      return 'table';
    case 'numbering':
      return 'numbering';
    default:
      return 'paragraph';
  }
}

// ─────────── 公开 API ───────────

/**
 * 扫 Package 构建样式索引。
 *
 * - styles.xml 不存在 → 返回空 Map，不抛错
 * - 同一 styleId 重复出现 → 后者覆盖（与 Word 行为一致）
 */
export function buildStylesIndex(pkg: Package): Map<string, StyleInfo> {
  const idx = new Map<string, StyleInfo>();

  const stylesPart = findStylesPart(pkg);
  if (!stylesPart) return idx;

  const styleEls = stylesPart.xmlDom.getElementsByTagNameNS(W_NS, 'style');
  for (let i = 0; i < styleEls.length; i++) {
    const el = styleEls[i];
    if (!el) continue;

    const styleId = el.getAttributeNS(W_NS, 'styleId');
    if (!styleId) continue; // styleId 缺失则跳过（非规范文档）

    const type = parseStyleType(el.getAttributeNS(W_NS, 'type'));
    const name = childAttr(el, W_NS, 'name', W_NS, 'val') ?? styleId;
    const basedOnRaw = childAttr(el, W_NS, 'basedOn', W_NS, 'val');

    const info: StyleInfo = { styleId, name, type };
    if (basedOnRaw) info.basedOn = basedOnRaw;

    // 大纲级别：从 w:style/w:pPr/w:outlineLvl@w:val 读取
    const outlineLvlVal = styleOutlineLevel(el);
    if (outlineLvlVal !== undefined) info.outlineLevel = outlineLvlVal;

    idx.set(styleId, info); // 重复时后者覆盖
  }

  return idx;
}

/**
 * 列出样式索引中的所有样式。
 *
 * @param stylesIdx  由 `buildStylesIndex` 构建的 Map
 * @param opts       可选过滤（目前支持按 type 过滤）
 */
export function listStyles(
  stylesIdx: Map<string, StyleInfo>,
  opts?: ListStylesOptions,
): StyleInfo[] {
  const all = Array.from(stylesIdx.values());
  if (!opts?.type) return all;
  return all.filter((s) => s.type === opts.type);
}

/**
 * 按 styleId 精确查找，O(1)。
 * styleId 大小写敏感（与 OOXML 规范一致）。
 */
export function getStyleById(stylesIdx: Map<string, StyleInfo>, styleId: string): StyleInfo | null {
  return stylesIdx.get(styleId) ?? null;
}
