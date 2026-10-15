/**
 * Drawing executor —— #190(replaceDrawingWithPlaceholder / deleteDrawing)+ #210
 * (replaceShapeInnerText)。
 *
 * 设计要点:
 *   - 入口都是 `drawingId`(装载期分配的 'dr-NNNN',来源:`indices.drawingIdx`),不接
 *     Element / paraId,避免调用方自己绕 DOM 反查。
 *   - replaceDrawingWithPlaceholder:仅接受 kind='picture' 的 drawing,把 a:blip/@r:embed
 *     改指共享 placeholder。布局完全保留(wp:inline / wp:anchor / wp:extent / 环绕 / position)。
 *     非 picture 抛 UNSUPPORTED_OOXML_FEATURE(由调用方走 deleteDrawing 或 replaceShapeInnerText)。
 *   - deleteDrawing:不限制 kind,物理删 w:drawing;承载 w:r 仅含该 drawing 时一并删除。
 *   - replaceShapeInnerText:仅接受含 w:txbxContent 的 drawing(textBox / shape-with-text),
 *     保留 w:drawing 整个外壳,只重写 txbxContent 内部为单段单 run 占位文字,保留首段 jc,
 *     并**继承原首 run rPr**(字体 / 字号 / 颜色)——红头艺术字字段化时占位 / 实际机关名
 *     不丢红色样式(解除 #210 v1「占位走默认字体」限制)。
 *     非 txbxContent drawing 抛 UNSUPPORTED_OOXML_FEATURE。这是 #192 内容型保留的工程依赖。
 *   - 旧 media / rels 引用计数清理:替换 / 删除后若旧 rId 在该 part 无其它引用,移除 rels
 *     条目;旧 media binary part 在全包无其它引用时一并删除。这是 #190 「原始图片二进制
 *     不出现在输出文档中」验收条件的保证。replaceShapeInnerText 同源清理 txbxContent
 *     内 hyperlink / 嵌套 picture 的 rId 引用。
 *   - trackChanges=true 暂不支持(与 deleteTable / replaceTableCellWithSdt 一致)。
 */
import { getPart, isXmlPart, markDirty } from '../../pack/index.js';
import { findStoryFor } from '../stories.js';
import { A_NS, PIC_NS, R_NS, W_NS, W14_NS, WP_NS } from '../namespaces.js';
import { generateParaId } from '../para-id.js';
import type { Op } from '../plan.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';
import {
  PLACEHOLDER_DESCR,
  ensurePlaceholderMediaPart,
  ensurePlaceholderRelInPart,
  ensurePngContentType,
  resolveRelTarget,
} from './placeholder-image.js';
import { createT, createW, firstDirectChildW, pickFirstRunRPr } from './_shared.js';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const PICTURE_GRAPHICDATA_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

// ─────────── replaceDrawingWithPlaceholder(#190)───────────

export function executeReplaceDrawingWithPlaceholder(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'replaceDrawingWithPlaceholder' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'replaceDrawingWithPlaceholder with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'replaceDrawingWithPlaceholder', drawingId: op.drawingId },
    );
  }

  const dr = requireDrawing(ctx, op.drawingId, 'replaceDrawingWithPlaceholder');

  // kind 校验:仅 picture 走本路径,其它(chart / smartArt / shape / textBox / group /
  // canvas / other)直接 reject —— 调用方应走 deleteDrawing 或 #192 的 shape 路径。
  const kindCheck = checkPictureKind(dr);
  if (!kindCheck.isPicture) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `replaceDrawingWithPlaceholder only supports kind='picture' drawings; got ${
        kindCheck.reason
      }. Use deleteDrawing for chart/smartArt/shape/textBox.`,
      ctx.docId,
      {
        opType: 'replaceDrawingWithPlaceholder',
        drawingId: op.drawingId,
        kindCheckReason: kindCheck.reason,
      },
    );
  }

  // 定位宿主 part 与对应 .rels
  const story = findStoryFor(ctx.pkg, dr);
  if (!story) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `drawing '${op.drawingId}' not found in any story`,
      ctx.docId,
      { drawingId: op.drawingId, opType: 'replaceDrawingWithPlaceholder' },
    );
  }
  const hostPart = getPart(ctx.pkg, story.partPath);
  if (!hostPart || !isXmlPart(hostPart)) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `story part '${story.partPath}' missing or not XML (op: replaceDrawingWithPlaceholder)`,
      ctx.docId,
      { drawingId: op.drawingId, partPath: story.partPath },
    );
  }
  const relsPath = relsPathFor(story.partPath);

  // 收集本 drawing 内所有 a:blip(picture 通常 1 个,但允许多张时统一处理)。
  // 同时记录每个 blip 替换前的 r:embed,供 S4 引用计数清理使用。
  const blips = snapshot(dr.getElementsByTagNameNS(A_NS, 'blip'));
  if (blips.length === 0) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `drawing '${op.drawingId}' has kind='picture' but contains no <a:blip>`,
      ctx.docId,
      { drawingId: op.drawingId, opType: 'replaceDrawingWithPlaceholder' },
    );
  }

  // 注入共享资源(幂等)
  const mediaCreated = ensurePlaceholderMediaPart(ctx.pkg);
  if (mediaCreated) ensurePngContentType(ctx.pkg);
  const placeholderRId = ensurePlaceholderRelInPart(ctx.pkg, relsPath);

  // 收集替换前的 r:embed / r:link,稍后由 S4 清理旧 media。
  // 当前 PR(S3):只做改写,清理路径占位,留 ctx 侧 channel 后续补。
  const oldRelIds: string[] = [];
  for (const blip of blips) {
    const oldEmbed = blip.getAttributeNS(R_NS, 'embed');
    const oldLink = blip.getAttributeNS(R_NS, 'link');
    if (oldEmbed) oldRelIds.push(oldEmbed);
    if (oldLink) oldRelIds.push(oldLink);
    // 移除 r:link(linked image 场景);只保留 r:embed 指向 placeholder
    if (oldLink !== null) blip.removeAttributeNS(R_NS, 'link');
    blip.setAttributeNS(R_NS, 'r:embed', placeholderRId);
  }

  // 写无障碍标注 wp:docPr/@descr(中文"[图片占位]"由此处体现)
  applyPlaceholderDescr(dr);

  markDirty(hostPart);

  // 旧 media 引用清理(S4)— 必须在 host part DOM 已经改完(blip 已重定向)之后再算引用数
  const relsPartObj = getPart(ctx.pkg, relsPath);
  if (relsPartObj && isXmlPart(relsPartObj)) {
    cleanupOrphanRels(ctx, hostPart, relsPartObj, relsPath, oldRelIds, placeholderRId);
  }
}

// ─────────── deleteDrawing(#190)───────────

export function executeDeleteDrawing(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'deleteDrawing' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'deleteDrawing with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'deleteDrawing', drawingId: op.drawingId },
    );
  }

  const dr = requireDrawing(ctx, op.drawingId, 'deleteDrawing');

  const story = findStoryFor(ctx.pkg, dr);
  if (!story) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `drawing '${op.drawingId}' not found in any story`,
      ctx.docId,
      { drawingId: op.drawingId, opType: 'deleteDrawing' },
    );
  }
  const hostPart = getPart(ctx.pkg, story.partPath);
  if (!hostPart || !isXmlPart(hostPart)) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `story part '${story.partPath}' missing or not XML (op: deleteDrawing)`,
      ctx.docId,
      { drawingId: op.drawingId, partPath: story.partPath },
    );
  }
  const relsPath = relsPathFor(story.partPath);

  // 1. 收集 drawing 子树里所有 rId(embed / link / id):删 drawing 后这些都可能成孤儿
  const referencedRelIds = collectRelIdsInSubtree(dr);

  // 2. 删 w:drawing 节点本身
  const parent = dr.parentNode;
  if (!parent) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `drawing '${op.drawingId}' has no parentNode`,
      ctx.docId,
      { drawingId: op.drawingId, opType: 'deleteDrawing' },
    );
  }
  parent.removeChild(dr);

  // 3. 父 w:r 若只剩 rPr / 已无实质内容,整段删(避免悬空空 run)
  if (
    parent.nodeType === 1 &&
    (parent as Element).namespaceURI === W_NS &&
    (parent as Element).localName === 'r' &&
    isRunEffectivelyEmpty(parent as Element)
  ) {
    const grand = parent.parentNode;
    if (grand) grand.removeChild(parent);
  }

  // 4. 同步 drawingIdx
  ctx.indices.drawingIdx.delete(op.drawingId);

  markDirty(hostPart);

  // 5. rels 引用清理(无 placeholder rId 需保留,传空串占位 keepRelId 不命中任何条目)
  const relsPartObj = getPart(ctx.pkg, relsPath);
  if (relsPartObj && isXmlPart(relsPartObj)) {
    cleanupOrphanRels(ctx, hostPart, relsPartObj, relsPath, referencedRelIds, '');
  }
}

// ─────────── replaceShapeInnerText(#210)───────────

export function executeReplaceShapeInnerText(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'replaceShapeInnerText' }>,
): void {
  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'replaceShapeInnerText with trackChanges=true not yet implemented',
      ctx.docId,
      { opType: 'replaceShapeInnerText', drawingId: op.drawingId },
    );
  }

  const dr = requireDrawing(ctx, op.drawingId, 'replaceShapeInnerText');

  // kind 校验:必须含 w:txbxContent(textBox / shape-with-text)。
  // - 嵌套 textBox 罕见但合法 —— 取 first descendant 在文档顺序上是最外层 txbxContent,符合直觉。
  // - mc:Fallback 内的 drawing 在装载期 buildDrawingIndex 已过滤,这里看到的是 Choice 路径。
  const txbxList = dr.getElementsByTagNameNS(W_NS, 'txbxContent');
  const txbx = txbxList.length > 0 ? (txbxList[0] as Element) : null;
  if (!txbx) {
    // 信息性错误:告诉调用方该走哪个 op 替代
    const kindCheck = checkPictureKind(dr);
    const hint = kindCheck.isPicture
      ? 'kind=picture; use replaceDrawingWithPlaceholder instead'
      : 'no w:txbxContent (shape without text / chart / smartArt / group / canvas / other); use deleteDrawing instead';
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      `replaceShapeInnerText only supports drawings containing w:txbxContent (textBox / shape-with-text); got ${hint}`,
      ctx.docId,
      { opType: 'replaceShapeInnerText', drawingId: op.drawingId, hint },
    );
  }

  // 定位宿主 part(同 #190 路径)
  const story = findStoryFor(ctx.pkg, dr);
  if (!story) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `drawing '${op.drawingId}' not found in any story`,
      ctx.docId,
      { drawingId: op.drawingId, opType: 'replaceShapeInnerText' },
    );
  }
  const hostPart = getPart(ctx.pkg, story.partPath);
  if (!hostPart || !isXmlPart(hostPart)) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      `story part '${story.partPath}' missing or not XML (op: replaceShapeInnerText)`,
      ctx.docId,
      { drawingId: op.drawingId, partPath: story.partPath },
    );
  }
  const relsPath = relsPathFor(story.partPath);

  // 替换前收集 txbxContent 子树内的 rId 引用(hyperlink / 嵌套 picture):删除后用于孤儿清理。
  // 注意:只扫 txbxContent 内部,外层 spPr/blipFill(若有作为 shape 填充图)不动。
  const referencedRelIds = collectRelIdsInSubtree(txbx);
  const oldParaIds = snapshot(txbx.getElementsByTagNameNS(W_NS, 'p'))
    .map((p) => p.getAttributeNS(W14_NS, 'paraId'))
    .filter((id): id is string => Boolean(id));

  // 保留首段 w:jc(对齐维度,#192 分组键依赖)
  const firstP = firstDirectChildW(txbx, 'p');
  const firstPPr = firstP ? firstDirectChildW(firstP, 'pPr') : null;

  // 保留被替换文字的原首 run rPr(字体 / 字号 / 颜色)——红头艺术字字段化的关键:
  // 占位文字(及后续排版填入的实际机关名)继承原 run 样式,而非走默认字体(解除 #210 v1 限制)。
  // pickFirstRunRPr 取传入元素内首个后代 run 的 rPr 克隆(覆盖 hyperlink / SDT 等包裹),
  // run 无显式 rPr 时兜底取该元素**直属** pPr/rPr ——故必须传**首段 w:p**(firstP)而非 txbx:
  //   pPr 是 w:p 的直属子,而 txbx 的直属子是 w:p,传 txbx 会让 pPr/rPr 兜底永远落空。
  // firstP 缺失(罕见,txbxContent 无段落)时退回 txbx,只能拿到后代 run 的 rPr(无 pPr/rPr 兜底)。
  // 两者都无 → undefined(裸 run,语义不变)。必须在清空 txbxContent 之前抓取。
  const firstRunRPr = pickFirstRunRPr(firstP ?? txbx);

  // 清空 txbxContent 全部子节点
  while (txbx.firstChild) txbx.removeChild(txbx.firstChild);
  for (const id of oldParaIds) ctx.indices.paraIdx.delete(id.toUpperCase());

  // 注入单段单 run 占位文字
  const doc = txbx.ownerDocument!;
  const p = createW(doc, 'p');
  const newParaId = generateParaId(new Set(ctx.indices.paraIdx.keys()));
  p.setAttributeNS(W14_NS, 'w14:paraId', newParaId);
  if (firstPPr) p.appendChild(firstPPr.cloneNode(true));
  const r = createW(doc, 'r');
  // 继承原 rPr(若有);rPr 必须是 run 首子(OOXML schema 要求 w:rPr 在 w:t 之前)
  if (firstRunRPr) r.appendChild(firstRunRPr);
  r.appendChild(createT(doc, op.placeholder));
  p.appendChild(r);
  txbx.appendChild(p);
  ctx.indices.paraIdx.set(newParaId, p);

  markDirty(hostPart);

  // rels 引用清理(同 #190):txbxContent 内若有 hyperlink/embed 的 rId,删后该 part 内归零则清 rels
  const relsPartObj = getPart(ctx.pkg, relsPath);
  if (relsPartObj && isXmlPart(relsPartObj)) {
    cleanupOrphanRels(ctx, hostPart, relsPartObj, relsPath, referencedRelIds, '');
  }
}

// ─────────── 内部 ───────────

function requireDrawing(ctx: ExecutorContext, drawingId: string, opName: string): Element {
  const dr = ctx.indices.drawingIdx.get(drawingId);
  if (!dr) {
    throw new WordError(
      'DRAWING_NOT_FOUND',
      `drawing with drawingId='${drawingId}' not found (op: ${opName})`,
      ctx.docId,
      { drawingId, opName },
    );
  }
  return dr;
}

function snapshot<T extends Element>(col: { length: number } & Record<number, T>): T[] {
  const out: T[] = [];
  for (let i = 0; i < col.length; i++) {
    const el = col[i];
    if (el) out.push(el);
  }
  return out;
}

/**
 * 判断 drawing 是否 kind='picture'。
 *
 * 与 list-drawings.ts 的判定逻辑对齐:看 a:graphicData @uri。这里独立实现一份
 * 避免反向依赖 list-drawings(那是只读 API,executor 走自己的轻量路径)。
 *
 * 进入路径:`w:drawing → wp:inline | wp:anchor → a:graphic → a:graphicData @uri`。
 * 若任一层缺失视为 corrupt(返回 reason=`malformed`),不当 picture 处理。
 */
function checkPictureKind(drawing: Element): { isPicture: boolean; reason: string } {
  // wp:inline / wp:anchor 任一
  let container: Element | null = null;
  for (let i = 0; i < drawing.childNodes.length; i++) {
    const c = drawing.childNodes[i];
    if (!c || c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI === WP_NS && (el.localName === 'inline' || el.localName === 'anchor')) {
      container = el;
      break;
    }
  }
  if (!container) return { isPicture: false, reason: 'malformed (no wp:inline / wp:anchor)' };

  // a:graphic → a:graphicData
  const graphic = firstChildNS(container, A_NS, 'graphic');
  if (!graphic) return { isPicture: false, reason: 'malformed (no a:graphic)' };
  const graphicData = firstChildNS(graphic, A_NS, 'graphicData');
  if (!graphicData) return { isPicture: false, reason: 'malformed (no a:graphicData)' };

  const uri = graphicData.getAttribute('uri');
  if (uri !== PICTURE_GRAPHICDATA_URI) {
    return { isPicture: false, reason: `kind=non-picture (a:graphicData/@uri='${uri ?? ''}')` };
  }

  // 进一步校验 pic:pic 存在(防御性:某些工具会用 picture URI 但内部嵌入非 pic:pic)
  const pic = firstChildNS(graphicData, PIC_NS, 'pic');
  if (!pic)
    return { isPicture: false, reason: 'malformed (graphicData/@uri=picture but no pic:pic)' };

  return { isPicture: true, reason: 'picture' };
}

function firstChildNS(parent: Element, ns: string, local: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const c = parent.childNodes[i];
    if (!c || c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI === ns && el.localName === local) return el;
  }
  return null;
}

/**
 * 设置 drawing 的中文无障碍标注。
 *
 * 路径:wp:inline | wp:anchor → wp:docPr。docPr 是 OOXML 规范保证的元素(Word
 * 装载期会自动补;若缺失则视作 corrupt,但本函数容错地不抛 —— 静默跳过)。
 *
 * 同时设 @descr(screen reader / OS accessibility 读取)与 @title(hover tooltip)。
 */
function applyPlaceholderDescr(drawing: Element): void {
  for (let i = 0; i < drawing.childNodes.length; i++) {
    const c = drawing.childNodes[i];
    if (!c || c.nodeType !== 1) continue;
    const container = c as Element;
    if (container.namespaceURI !== WP_NS) continue;
    if (container.localName !== 'inline' && container.localName !== 'anchor') continue;
    const docPr = firstChildNS(container, WP_NS, 'docPr');
    if (!docPr) return; // 罕见;不抛,保守跳过
    docPr.setAttribute('descr', PLACEHOLDER_DESCR);
    docPr.setAttribute('title', PLACEHOLDER_DESCR);
    return;
  }
}

/**
 * 把 part 路径转换为对应 .rels part 路径。
 *
 *   'word/document.xml' → 'word/_rels/document.xml.rels'
 *   'word/header1.xml'  → 'word/_rels/header1.xml.rels'
 *   '.rels' 顶层场景不出现在 #190 用例里(drawing 不会挂顶层 part)
 */
function relsPathFor(partPath: string): string {
  const idx = partPath.lastIndexOf('/');
  if (idx < 0) return `_rels/${partPath}.rels`;
  const dir = partPath.slice(0, idx);
  const name = partPath.slice(idx + 1);
  return `${dir}/_rels/${name}.rels`;
}

/**
 * 旧 rels / media 引用计数清理(#190 S4)。
 *
 * 流程:
 *   1. 对每个 oldRelId,在宿主 part XML 内统计剩余引用数(r:embed / r:link / r:id 三种属性)
 *   2. 引用归零 → 删 rels DOM 中对应 `<Relationship>`,记录 Target → 候选 media path
 *   3. media path 加入 ctx.orphanMediaCandidates(由 kernel 在事务 commit 后统一 sweep)
 *
 * 为什么不在这里直接删 binary part:同一 media 可能被多 drawing 共享(其它 part 或同 part
 * 后续未处理 drawing),只在全包 rels 都不再引用时才安全。判定与删除延迟到 sweep 阶段。
 *
 * @param hostPart  drawing 所在 XML part
 * @param relsPart  对应 .rels part(必须已 isXmlPart && xmlDom)
 * @param relsPath  rels part 路径(用于 Target → 绝对 path 解析)
 * @param oldRelIds 替换前的 r:embed / r:link 列表(允许重复;函数内去重)
 * @param keepRelId 新写入的 placeholder rId,不要把它当孤儿
 */
function cleanupOrphanRels(
  ctx: ExecutorContext,
  hostPart: { xmlDom?: Document },
  relsPart: { xmlDom?: Document },
  relsPath: string,
  oldRelIds: readonly string[],
  keepRelId: string,
): void {
  if (oldRelIds.length === 0) return;
  if (!hostPart.xmlDom || !relsPart.xmlDom) return;

  const uniqueOld = new Set(oldRelIds.filter((id) => id && id !== keepRelId));
  if (uniqueOld.size === 0) return;

  // 1. 宿主 part 内仍引用 oldRelId 的次数(replace 已经把所有 a:blip 改了;但同 part
  //    内可能有其它 drawing / VML / chart 也用同名 rId,本步骤是防误删)
  const orphanRelIds: string[] = [];
  for (const oldId of uniqueOld) {
    if (countRelRefsInDoc(hostPart.xmlDom, oldId) === 0) orphanRelIds.push(oldId);
  }
  if (orphanRelIds.length === 0) return;

  // 2. 从 rels DOM 删 orphan <Relationship>,记录 media path
  const rels = snapshot(relsPart.xmlDom.getElementsByTagNameNS(RELS_NS, 'Relationship'));
  let relsTouched = false;
  for (const rel of rels) {
    const id = rel.getAttribute('Id') ?? '';
    if (!orphanRelIds.includes(id)) continue;
    const target = rel.getAttribute('Target') ?? '';
    if (target && ctx.orphanMediaCandidates) {
      const mediaPath = resolveRelTarget(relsPath, target);
      ctx.orphanMediaCandidates.add(mediaPath);
    }
    rel.parentNode?.removeChild(rel);
    relsTouched = true;
  }
  if (relsTouched) {
    const rp = getPart(ctx.pkg, relsPath);
    if (rp) markDirty(rp);
  }
}

/**
 * 统计文档内引用某个 rId 的次数。
 *
 * 扫所有元素,匹配 r:embed / r:link / r:id 三类属性(OOXML 中 rId 引用的全部形态)。
 * 性能:全文遍历一遍,对中小文档可接受;若未来发现热点可改成按已知 tag 名(a:blip /
 * v:imagedata / c:chart / w:hyperlink 等)定向扫描。
 */
function countRelRefsInDoc(doc: Document, relId: string): number {
  let count = 0;
  const all = doc.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    const el = all[i] as Element | undefined;
    if (!el) continue;
    if (el.getAttributeNS(R_NS, 'embed') === relId) count++;
    if (el.getAttributeNS(R_NS, 'link') === relId) count++;
    if (el.getAttributeNS(R_NS, 'id') === relId) count++;
  }
  return count;
}

/**
 * 收集 drawing 子树内所有 rId 引用(r:embed / r:link / r:id),返回去重列表。
 *
 * 覆盖范围:不仅是 a:blip 的 embed/link,也包括 chart 的 r:id、SmartArt 的 r:dm/r:lo/r:qs/r:cs
 * (这些都用 R_NS 的 id 属性引用 chart.xml / data1.xml / styles.xml 等关联 part)。
 *
 * 注意:删除 drawing 时这些子树内所有 rId 都"可能"成为孤儿;sweepOrphanMediaParts
 * 只会删 isBinaryPart 的部分,XML 关联 part(如 chart.xml)会被跳过留存,行为合理。
 */
function collectRelIdsInSubtree(drawing: Element): string[] {
  const ids = new Set<string>();
  const all = drawing.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    const el = all[i] as Element | undefined;
    if (!el) continue;
    const embed = el.getAttributeNS(R_NS, 'embed');
    if (embed) ids.add(embed);
    const link = el.getAttributeNS(R_NS, 'link');
    if (link) ids.add(link);
    const rid = el.getAttributeNS(R_NS, 'id');
    if (rid) ids.add(rid);
  }
  return [...ids];
}

/**
 * 判断 `<w:r>` 是否仅含 `<w:rPr>` 或无内容(即可安全删除)。
 *
 * 复用自 table-ops.ts 的 isRunEffectivelyEmpty 语义,这里就近实现避免跨文件 export
 * shared util。判定:除 w:rPr 外没有其它元素子节点 = 空。
 */
function isRunEffectivelyEmpty(run: Element): boolean {
  for (let i = 0; i < run.childNodes.length; i++) {
    const c = run.childNodes[i];
    if (!c || c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI === W_NS && el.localName === 'rPr') continue;
    return false;
  }
  return true;
}
