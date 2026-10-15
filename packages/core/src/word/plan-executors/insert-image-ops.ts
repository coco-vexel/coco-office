/**
 * insertImageAfter executor —— task-insert-image。
 *
 * 在 anchorParaId 段之后新插一个 <w:p>，段内含 inline 图片 <w:drawing>。
 *
 * 写链（8 步）：
 *   1. anchor 定位（paraIdx → PARA_NOT_FOUND）
 *   2. 分配 media 文件名（image1..imageN.ext，无撞）
 *   3. 新建 binary part（createBinaryPart + setPart + markDirty）
 *   4. 加 image rel（rels 扫 max(rIdN)+1）
 *   5. 登记 content-type（ensureImageContentType）
 *   6. 构造 <w:drawing> DOM（inline picture 模板）
 *   7. 包 <w:r> + 新 <w:p>（含 w14:paraId 主动分配）
 *   8. markDirty
 *
 * 跨端：只用 Uint8Array，绝不碰 node fs/Buffer。
 */
import { parseXmlDocument } from 'slimdom';
import {
  createBinaryPart,
  createXmlPart,
  getPart,
  hasPart,
  isXmlPart,
  listParts,
  markDirty,
  setPart,
  type Package,
} from '../../pack/index.js';
import { generateParaId } from '../para-id.js';
import { A_NS, PIC_NS, R_NS, W_NS, W14_NS, WP_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { findStoryFor } from '../stories.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_PATH = '[Content_Types].xml';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const IMAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/**
 * contentType → 文件扩展名。
 *
 * 含位图 5 种 + 矢量图 2 种（EMF/WMF）。
 *
 * EMF/WMF 是 Word 原生支持的矢量图格式（粘贴的 Office 图表/示意图、老文档插图常见），
 * 在 OOXML 里和位图一样走 a:blip/@r:embed 引用机制 —— blip 对图片格式一视同仁，
 * 矢量图 part 的扩展名是 .emf/.wmf、content-type 是 image/x-emf/image/x-wmf。
 * typeset 透传源文档已有的 EMF/WMF 图时（getImageBytes 读出原始字节 + contentType，
 * insertImageAfter 重新引用），白名单必须含这两种，否则真实文档（含矢量图）排版崩溃。
 */
export const CONTENT_TYPE_TO_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/tiff': 'tiff',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
};

// ─────────── executor ───────────

export function executeInsertImageAfter(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'insertImageAfter' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'insertImageAfter with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'insertImageAfter' },
    );
  }

  // contentType 校验
  const ext = CONTENT_TYPE_TO_EXT[op.contentType];
  if (!ext) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `insertImageAfter: unsupported contentType '${op.contentType}'. Supported: image/png, image/jpeg, image/gif, image/bmp, image/tiff, image/x-emf, image/x-wmf`,
      ctx.docId,
      { opType: 'insertImageAfter', contentType: op.contentType },
    );
  }

  // 1. anchor 定位
  const anchor = ctx.indices.paraIdx.get(op.anchorParaId.toUpperCase());
  if (!anchor) {
    throw new WordError(
      'PARA_NOT_FOUND',
      `paragraph with paraId='${op.anchorParaId}' not found (op: insertImageAfter)`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: 'insertImageAfter' },
    );
  }
  const story = findStoryFor(ctx.pkg, anchor);
  if (!story) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `anchor paragraph '${op.anchorParaId}' not found in any story`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: 'insertImageAfter' },
    );
  }
  const hostPartPath = story.partPath;
  const relsPath = relsPathFor(hostPartPath);

  // 2. 分配 media 文件名（image1..imageN.ext）
  const mediaPath = allocateMediaPath(ctx.pkg, ext);

  // 3. 新建 binary part
  const mediaPart = createBinaryPart({
    path: mediaPath,
    contentType: op.contentType,
    bytes: op.bytes,
  });
  setPart(ctx.pkg, mediaPart);
  markDirty(mediaPart);

  // 4. 加 image rel（rels 扫 max(rIdN)+1）
  const rId = addImageRel(ctx.pkg, relsPath, mediaPath);

  // 5. 登记 content-type
  ensureImageContentType(ctx.pkg, ext, op.contentType);

  // 6. 构造 <w:drawing> DOM
  const hostPart = getPart(ctx.pkg, hostPartPath);
  if (!hostPart || !isXmlPart(hostPart) || !hostPart.xmlDom) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `host part '${hostPartPath}' missing or not XML`,
      ctx.docId,
      { opName: 'insertImageAfter' },
    );
  }
  const doc = hostPart.xmlDom;

  const docPrId = allocateDocPrId(doc);
  const drawing = buildInlineDrawing(doc, rId, docPrId, op.widthEmu, op.heightEmu);

  // 7. 新 <w:p>（含主动分配 w14:paraId）
  const seen = new Set(ctx.indices.paraIdx.keys());
  const newParaId = generateParaId(seen);

  const newPara = doc.createElementNS(W_NS, 'w:p');
  newPara.setAttributeNS(W14_NS, 'w14:paraId', newParaId);
  const run = doc.createElementNS(W_NS, 'w:r');
  run.appendChild(drawing);
  newPara.appendChild(run);

  // 插到 anchor 之后（insert-before anchor.nextSibling）
  const parent = anchor.parentNode;
  if (!parent) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `anchor paragraph '${op.anchorParaId}' has no parentNode`,
      ctx.docId,
      { paraId: op.anchorParaId, opName: 'insertImageAfter' },
    );
  }
  parent.insertBefore(newPara, anchor.nextSibling);

  // 同步 paraIdx（保证同一 plan 内后续 op 可查到新段）
  ctx.indices.paraIdx.set(newParaId, newPara);

  // 8. markDirty
  markDirty(hostPart);
  const relsPart = getPart(ctx.pkg, relsPath);
  if (relsPart) markDirty(relsPart);
  const ctPart = getPart(ctx.pkg, CONTENT_TYPES_PATH);
  if (ctPart) markDirty(ctPart);
}

// ─────────── helpers ───────────

/**
 * 分配唯一 media 文件名：扫 word/media/imageN.ext，取最大 N+1。
 * hasPart 二次校验防撞（理论不撞，保险用）。
 */
export function allocateMediaPath(pkg: Package, ext: string): string {
  let maxN = 0;
  for (const part of listParts(pkg)) {
    const m = /^word\/media\/image(\d+)\./i.exec(part.path);
    if (m) {
      const n = parseInt(m[1]!, 10);
      if (n > maxN) maxN = n;
    }
  }
  let candidate = `word/media/image${maxN + 1}.${ext}`;
  // 二次校验防撞（例如 image1.png 已存在但 image1.jpeg 也占了 N=1）
  while (hasPart(pkg, candidate)) {
    maxN++;
    candidate = `word/media/image${maxN + 1}.${ext}`;
  }
  return candidate;
}

/**
 * 在 relsPartPath 里新增一条 image relationship 指向 mediaPath，返回分配的 rId。
 * 若 rels part 不存在，先创建空 rels。
 */
export function addImageRelForHost(pkg: Package, relsPartPath: string, mediaPath: string): string {
  return addImageRel(pkg, relsPartPath, mediaPath);
}

function addImageRel(pkg: Package, relsPartPath: string, mediaPath: string): string {
  let relsPart = getPart(pkg, relsPartPath);
  if (!relsPart) {
    const empty = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}"/>`;
    const dom = parseXmlDocument(empty) as unknown as Document;
    relsPart = createXmlPart({
      path: relsPartPath,
      contentType: 'application/vnd.openxmlformats-package.relationships+xml',
      xmlDom: dom,
    });
    setPart(pkg, relsPart);
  }
  if (!isXmlPart(relsPart) || !relsPart.xmlDom) {
    throw new Error(`addImageRel: '${relsPartPath}' is not an XML part`);
  }

  const doc = relsPart.xmlDom;
  const rels = doc.getElementsByTagNameNS(RELS_NS, 'Relationship');
  let maxRId = 0;
  for (let i = 0; i < rels.length; i++) {
    const id = rels[i]!.getAttribute('Id') ?? '';
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = parseInt(m[1]!, 10);
      if (n > maxRId) maxRId = n;
    }
  }

  const rId = `rId${maxRId + 1}`;
  // Target 相对于宿主 part 目录：word/document.xml → word/media/imageN.ext → media/imageN.ext
  const target = computeRelTarget(relsPartPath, mediaPath);
  const rel = doc.createElementNS(RELS_NS, 'Relationship');
  rel.setAttribute('Id', rId);
  rel.setAttribute('Type', IMAGE_RELATIONSHIP_TYPE);
  rel.setAttribute('Target', target);
  doc.documentElement.appendChild(rel);
  markDirty(relsPart);
  return rId;
}

/**
 * 登记 content-type Default：补 `<Default Extension=ext ContentType=ct/>`。
 * 已有同 Extension 的 Default → 跳过。
 */
export function ensureImageContentType(pkg: Package, ext: string, contentType: string): void {
  const ctPart = getPart(pkg, CONTENT_TYPES_PATH);
  if (!ctPart || !isXmlPart(ctPart) || !ctPart.xmlDom) return;
  const doc = ctPart.xmlDom;

  const defaults = doc.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Default');
  for (let i = 0; i < defaults.length; i++) {
    if (defaults[i]!.getAttribute('Extension')?.toLowerCase() === ext.toLowerCase()) return;
  }

  const def = doc.createElementNS(CONTENT_TYPES_NS, 'Default');
  def.setAttribute('Extension', ext);
  def.setAttribute('ContentType', contentType);
  doc.documentElement.appendChild(def);
  markDirty(ctPart);
}

/**
 * 扫 document.xmlDom 里所有 wp:docPr/@id，取 max+1。
 */
function allocateDocPrId(doc: Document): number {
  const docPrs = doc.getElementsByTagNameNS(WP_NS, 'docPr');
  let maxId = 0;
  for (let i = 0; i < docPrs.length; i++) {
    const raw = docPrs[i]!.getAttribute('id');
    if (!raw) continue;
    const n = parseInt(raw, 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  return maxId + 1;
}

/**
 * 构造 inline picture <w:drawing> DOM。
 *
 * 结构：
 *   w:drawing
 *     wp:inline(distT/B/L/R=0)
 *       wp:extent(@cx @cy)
 *       wp:docPr(@id @name)
 *       wp:cNvGraphicFramePr > a:graphicFrameLocks(noChangeAspect=1)
 *       a:graphic > a:graphicData(uri=PIC_NS)
 *         pic:pic
 *           pic:nvPicPr > pic:cNvPr(@id @name) + pic:cNvPicPr > a:picLocks(noChangeAspect=1)
 *           pic:blipFill > a:blip(@r:embed) + a:stretch > a:fillRect
 *           pic:spPr > a:xfrm > (a:off x=0 y=0 + a:ext @cx @cy) + a:prstGeom(prst=rect) > a:avLst
 */
function buildInlineDrawing(
  doc: Document,
  rId: string,
  docPrId: number,
  widthEmu: number,
  heightEmu: number,
): Element {
  const imgName = `image${docPrId}`;

  function ns(nsUri: string, local: string): Element {
    // Use prefixed form for non-W_NS elements to match real OOXML output
    const prefix = nsPrefixOf(nsUri);
    return doc.createElementNS(nsUri, prefix ? `${prefix}:${local}` : local);
  }

  // w:drawing
  const drawing = doc.createElementNS(W_NS, 'w:drawing');

  // wp:inline
  const inline = ns(WP_NS, 'inline');
  inline.setAttribute('distT', '0');
  inline.setAttribute('distB', '0');
  inline.setAttribute('distL', '0');
  inline.setAttribute('distR', '0');
  drawing.appendChild(inline);

  // wp:extent
  const extent = ns(WP_NS, 'extent');
  extent.setAttribute('cx', String(widthEmu));
  extent.setAttribute('cy', String(heightEmu));
  inline.appendChild(extent);

  // wp:effectExtent
  const effectExtent = ns(WP_NS, 'effectExtent');
  effectExtent.setAttribute('l', '0');
  effectExtent.setAttribute('t', '0');
  effectExtent.setAttribute('r', '0');
  effectExtent.setAttribute('b', '0');
  inline.appendChild(effectExtent);

  // wp:docPr
  const docPr = ns(WP_NS, 'docPr');
  docPr.setAttribute('id', String(docPrId));
  docPr.setAttribute('name', imgName);
  inline.appendChild(docPr);

  // wp:cNvGraphicFramePr
  const cNvGFP = ns(WP_NS, 'cNvGraphicFramePr');
  const gfLocks = doc.createElementNS(A_NS, 'a:graphicFrameLocks');
  gfLocks.setAttribute('noChangeAspect', '1');
  cNvGFP.appendChild(gfLocks);
  inline.appendChild(cNvGFP);

  // a:graphic > a:graphicData
  const graphic = doc.createElementNS(A_NS, 'a:graphic');
  const graphicData = doc.createElementNS(A_NS, 'a:graphicData');
  graphicData.setAttribute('uri', PIC_NS);
  graphic.appendChild(graphicData);
  inline.appendChild(graphic);

  // pic:pic
  const pic = doc.createElementNS(PIC_NS, 'pic:pic');
  graphicData.appendChild(pic);

  // pic:nvPicPr
  const nvPicPr = doc.createElementNS(PIC_NS, 'pic:nvPicPr');
  pic.appendChild(nvPicPr);

  const cNvPr = doc.createElementNS(PIC_NS, 'pic:cNvPr');
  cNvPr.setAttribute('id', String(docPrId));
  cNvPr.setAttribute('name', imgName);
  nvPicPr.appendChild(cNvPr);

  const cNvPicPr = doc.createElementNS(PIC_NS, 'pic:cNvPicPr');
  const picLocks = doc.createElementNS(A_NS, 'a:picLocks');
  picLocks.setAttribute('noChangeAspect', '1');
  cNvPicPr.appendChild(picLocks);
  nvPicPr.appendChild(cNvPicPr);

  // pic:blipFill
  const blipFill = doc.createElementNS(PIC_NS, 'pic:blipFill');
  pic.appendChild(blipFill);

  const blip = doc.createElementNS(A_NS, 'a:blip');
  blip.setAttributeNS(R_NS, 'r:embed', rId);
  blipFill.appendChild(blip);

  const stretch = doc.createElementNS(A_NS, 'a:stretch');
  const fillRect = doc.createElementNS(A_NS, 'a:fillRect');
  stretch.appendChild(fillRect);
  blipFill.appendChild(stretch);

  // pic:spPr
  const spPr = doc.createElementNS(PIC_NS, 'pic:spPr');
  pic.appendChild(spPr);

  const xfrm = doc.createElementNS(A_NS, 'a:xfrm');
  const off = doc.createElementNS(A_NS, 'a:off');
  off.setAttribute('x', '0');
  off.setAttribute('y', '0');
  xfrm.appendChild(off);
  const extEl = doc.createElementNS(A_NS, 'a:ext');
  extEl.setAttribute('cx', String(widthEmu));
  extEl.setAttribute('cy', String(heightEmu));
  xfrm.appendChild(extEl);
  spPr.appendChild(xfrm);

  const prstGeom = doc.createElementNS(A_NS, 'a:prstGeom');
  prstGeom.setAttribute('prst', 'rect');
  const avLst = doc.createElementNS(A_NS, 'a:avLst');
  prstGeom.appendChild(avLst);
  spPr.appendChild(prstGeom);

  return drawing;
}

/** namespace URI → prefix map（用于创建有正确前缀的元素，与真实 OOXML 匹配）*/
function nsPrefixOf(nsUri: string): string {
  if (nsUri === WP_NS) return 'wp';
  if (nsUri === A_NS) return 'a';
  if (nsUri === PIC_NS) return 'pic';
  if (nsUri === R_NS) return 'r';
  return '';
}

/**
 * 把 rels part 路径转成宿主 part 所在目录，用于计算相对 Target。
 * 'word/_rels/document.xml.rels' → base='word'
 */
function computeRelTarget(relsPartPath: string, targetPartPath: string): string {
  const idx = relsPartPath.lastIndexOf('/_rels/');
  if (idx < 0) return targetPartPath;
  const hostDir = relsPartPath.slice(0, idx);
  if (targetPartPath.startsWith(`${hostDir}/`)) {
    return targetPartPath.slice(hostDir.length + 1);
  }
  return `/${targetPartPath}`;
}

/**
 * 把 part 路径转换为对应 .rels part 路径。
 *   'word/document.xml' → 'word/_rels/document.xml.rels'
 */
function relsPathFor(partPath: string): string {
  const idx = partPath.lastIndexOf('/');
  if (idx < 0) return `_rels/${partPath}.rels`;
  const dir = partPath.slice(0, idx);
  const name = partPath.slice(idx + 1);
  return `${dir}/_rels/${name}.rels`;
}

// ─────────── getImageBytes（内核侧同步实现，供 kernel.ts 调用）───────────

export interface ImageBytesResult {
  bytes: Uint8Array;
  contentType: string;
  extentEMU: { cx: number; cy: number };
}

/**
 * 读取 drawingId 对应 picture 的原始图片字节。
 * drawingId 不存在 / 非 picture / media 找不到 → 返回 null（不抛）。
 */
export function getImageBytesSync(
  pkg: Package,
  indices: { drawingIdx: Map<string, Element> },
  drawingId: string,
): ImageBytesResult | null {
  // 1. drawingId → <w:drawing>
  const dr = indices.drawingIdx.get(drawingId);
  if (!dr) return null;

  // 2. <w:drawing> → a:blip/@r:embed
  const blips = dr.getElementsByTagNameNS(A_NS, 'blip');
  if (blips.length === 0) return null;
  const blip = blips[0]!;
  const rId = blip.getAttributeNS(R_NS, 'embed');
  if (!rId) return null;

  // 3. <w:drawing> → 宿主 part path
  const story = findStoryFor(pkg, dr);
  if (!story) return null;
  const hostPartPath = story.partPath;

  // 4. partPath → rels path
  const relsPath = relsPathFor(hostPartPath);

  // 5. rId → media path via rels
  const relsPart = getPart(pkg, relsPath);
  if (!relsPart || !isXmlPart(relsPart) || !relsPart.xmlDom) return null;
  const rels = relsPart.xmlDom.getElementsByTagNameNS(RELS_NS, 'Relationship');
  let mediaPath: string | null = null;
  for (let i = 0; i < rels.length; i++) {
    const r = rels[i]!;
    if (r.getAttribute('Id') === rId) {
      const target = r.getAttribute('Target') ?? '';
      // resolve relative target → absolute part path
      mediaPath = resolveTarget(relsPath, target);
      break;
    }
  }
  if (!mediaPath) return null;

  // 6. media path → bytes
  const mediaPart = getPart(pkg, mediaPath);
  if (!mediaPart) return null;
  // isBinaryPart check: has bytes field
  const bytes = (mediaPart as unknown as { bytes?: Uint8Array }).bytes;
  if (!bytes) return null;
  const contentType = mediaPart.contentType;

  // extentEMU: from wp:extent inside wp:inline / wp:anchor
  const extentEMU = readDrawingExtent(dr);

  return { bytes, contentType, extentEMU };
}

/** wp:extent @cx/@cy from wp:inline or wp:anchor child */
function readDrawingExtent(drawing: Element): { cx: number; cy: number } {
  for (let i = 0; i < drawing.childNodes.length; i++) {
    const c = drawing.childNodes[i] as Element;
    if (!c || c.nodeType !== 1) continue;
    if (c.namespaceURI !== WP_NS) continue;
    if (c.localName !== 'inline' && c.localName !== 'anchor') continue;
    // find wp:extent child
    for (let j = 0; j < c.childNodes.length; j++) {
      const ext = c.childNodes[j] as Element;
      if (!ext || ext.nodeType !== 1) continue;
      if (ext.namespaceURI === WP_NS && ext.localName === 'extent') {
        const cx = parseInt(ext.getAttribute('cx') ?? '0', 10);
        const cy = parseInt(ext.getAttribute('cy') ?? '0', 10);
        return { cx: isNaN(cx) ? 0 : cx, cy: isNaN(cy) ? 0 : cy };
      }
    }
  }
  return { cx: 0, cy: 0 };
}

/** Resolve rels Target → absolute package path (mirrors resolveRelTarget in placeholder-image.ts) */
function resolveTarget(relsPartPath: string, target: string): string {
  let t = target.trim();
  if (t.startsWith('/')) return t.slice(1);
  while (t.startsWith('./')) t = t.slice(2);
  const idx = relsPartPath.lastIndexOf('/_rels/');
  if (idx < 0) return t;
  const hostDir = relsPartPath.slice(0, idx);
  // handle '../' traversal
  let parts = hostDir.split('/').filter(Boolean);
  while (t.startsWith('../')) {
    parts = parts.slice(0, -1);
    t = t.slice(3);
  }
  return parts.length === 0 ? t : `${parts.join('/')}/${t}`;
}

// Re-export for kernel.ts convenience
export { relsPathFor as relsPathForImage };
// Re-export for insert-table-ops.ts（跨文档搬表时算 host rels 路径）。
export { relsPathFor as relsPathForHost };
