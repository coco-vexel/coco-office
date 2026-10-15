/**
 * Section（节）查询（#211 / docs/03.1 §2）。
 *
 * 节边界由"承载 `<w:sectPr>` 的段"标定：
 *   - 主 body 的最后一个 `<w:sectPr>` 是 body 直接子元素（不属于任何段），是最后一节
 *   - 其余节的 sectPr 嵌在 `<w:p>/<w:pPr>/<w:sectPr>`，该段是节的最后一段
 *
 * 算法：遍历 body 直属节点（doc-order），收集 `<w:p>`，遇含 sectPr 的段闭合一节
 *      （lastParaId = 该段 paraId）；body 末尾的散 sectPr 闭合最后一节。
 *
 * 节信息字段（消费方主要诉求来自 skeleton #195 封面检测）：
 *   - sectionIndex / firstParaId / lastParaId：边界
 *   - type：'nextPage' / 'continuous' / 'evenPage' / 'oddPage' / 'nextColumn'（缺省 = nextPage，OOXML 默认）
 *   - headerRefs / footerRefs：resolve 到 partName，跨节比对页眉页脚是否相同
 *   - titlePage：'首页不同' 设置（封面强信号）
 */
import type { Package } from '../pack/index.js';
import { listParts } from '../pack/index.js';
import { W_NS, W14_NS, R_NS } from './namespaces.js';
import { resolveRelTarget } from './plan-executors/placeholder-image.js';
import { listStories } from './stories.js';

export type SectionType = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
export type HeaderFooterType = 'default' | 'first' | 'even';

export interface HeaderFooterRef {
  type: HeaderFooterType;
  /** rels 目标 part 路径（如 'word/header1.xml'） */
  partName: string;
}

export interface SectionInfo {
  sectionIndex: number;
  firstParaId: string;
  /** 节末段的 paraId（承载 sectPr 的段；最后一节为 body 末尾散 sectPr 时回退到 body 最后段）*/
  lastParaId: string;
  type?: SectionType;
  headerRefs?: HeaderFooterRef[];
  footerRefs?: HeaderFooterRef[];
  /** w:sectPr/w:titlePg 存在 = 首页不同（封面强信号）*/
  titlePage?: boolean;
}

export function listSections(pkg: Package): SectionInfo[] {
  const stories = listStories(pkg, { kinds: ['body'] });
  if (stories.length === 0) return [];
  const body = stories[0]!.root;
  const partPath = stories[0]!.partPath;
  const relsPath = relsPathFor(partPath);

  const out: SectionInfo[] = [];
  let sectionFirstParaId: string | undefined;
  let lastParaIdInBody: string | undefined;
  let sectionIndex = 0;

  // 遍历 body 直属子节点
  for (let i = 0; i < body.childNodes.length; i++) {
    const node = body.childNodes[i];
    if (!node || node.nodeType !== 1) continue;
    const el = node as Element;
    if (el.namespaceURI !== W_NS) continue;

    if (el.localName === 'p') {
      const paraId = el.getAttributeNS(W14_NS, 'paraId')?.toUpperCase() ?? '';
      if (paraId.length === 0) continue;
      if (sectionFirstParaId === undefined) sectionFirstParaId = paraId;
      lastParaIdInBody = paraId;

      // 该段是否承载 sectPr？
      const pPr = firstDirectChild(el, W_NS, 'pPr');
      const inlineSectPr = pPr ? firstDirectChild(pPr, W_NS, 'sectPr') : null;
      if (inlineSectPr) {
        out.push(
          buildSection(sectionIndex++, sectionFirstParaId, paraId, inlineSectPr, pkg, relsPath),
        );
        sectionFirstParaId = undefined;
      }
    } else if (el.localName === 'sectPr') {
      // body 末尾散 sectPr：最后一节，lastParaId 用 body 最后段
      if (sectionFirstParaId !== undefined && lastParaIdInBody !== undefined) {
        out.push(
          buildSection(sectionIndex++, sectionFirstParaId, lastParaIdInBody, el, pkg, relsPath),
        );
        sectionFirstParaId = undefined;
      }
    }
    // 其他节点（tbl 等）按 doc-order 跳过，不影响节边界
  }

  return out;
}

// ─────────── 内部 ───────────

function buildSection(
  sectionIndex: number,
  firstParaId: string,
  lastParaId: string,
  sectPr: Element,
  pkg: Package,
  relsPath: string | undefined,
): SectionInfo {
  const info: SectionInfo = { sectionIndex, firstParaId, lastParaId };

  // type
  const typeEl = firstDirectChild(sectPr, W_NS, 'type');
  if (typeEl) {
    const v = typeEl.getAttributeNS(W_NS, 'val');
    if (
      v === 'nextPage' ||
      v === 'continuous' ||
      v === 'evenPage' ||
      v === 'oddPage' ||
      v === 'nextColumn'
    ) {
      info.type = v;
    }
  }

  // header/footer references
  const headerRefs = collectRefs(sectPr, 'headerReference', pkg, relsPath);
  if (headerRefs.length > 0) info.headerRefs = headerRefs;
  const footerRefs = collectRefs(sectPr, 'footerReference', pkg, relsPath);
  if (footerRefs.length > 0) info.footerRefs = footerRefs;

  // titlePg
  if (firstDirectChild(sectPr, W_NS, 'titlePg')) info.titlePage = true;

  return info;
}

function collectRefs(
  sectPr: Element,
  localName: 'headerReference' | 'footerReference',
  pkg: Package,
  relsPath: string | undefined,
): HeaderFooterRef[] {
  const refs: HeaderFooterRef[] = [];
  if (!relsPath) return refs;
  const relMap = loadRels(pkg, relsPath);
  if (!relMap) return refs;
  const nodes = sectPr.getElementsByTagNameNS(W_NS, localName);
  for (let i = 0; i < nodes.length; i++) {
    const el = nodes[i];
    if (!el || el.parentNode !== sectPr) continue;
    const typeAttr = el.getAttributeNS(W_NS, 'type');
    const rId = el.getAttributeNS(R_NS, 'id');
    if (!rId) continue;
    const target = relMap.get(rId);
    if (!target) continue;
    const partName = resolveRelTarget(relsPath, target);
    const type: HeaderFooterType =
      typeAttr === 'first' || typeAttr === 'even' ? typeAttr : 'default';
    refs.push({ type, partName });
  }
  return refs;
}

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

function loadRels(pkg: Package, relsPath: string): Map<string, string> | null {
  for (const part of listParts(pkg)) {
    if (part.path !== relsPath) continue;
    if (!part.xmlDom) return null;
    const map = new Map<string, string>();
    const nodes = part.xmlDom.getElementsByTagNameNS(RELS_NS, 'Relationship');
    for (let i = 0; i < nodes.length; i++) {
      const el = nodes[i];
      if (!el) continue;
      const id = el.getAttribute('Id');
      const target = el.getAttribute('Target');
      if (id && target) map.set(id, target);
    }
    return map;
  }
  return null;
}

function relsPathFor(partPath: string): string | undefined {
  const lastSlash = partPath.lastIndexOf('/');
  const dir = lastSlash >= 0 ? partPath.slice(0, lastSlash) : '';
  const file = partPath.slice(lastSlash + 1);
  return dir.length > 0 ? `${dir}/_rels/${file}.rels` : `_rels/${file}.rels`;
}

function firstDirectChild(el: Element, ns: string, local: string): Element | null {
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (!c || c.nodeType !== 1) continue;
    const ce = c as Element;
    if (ce.namespaceURI === ns && ce.localName === local) return ce;
  }
  return null;
}
