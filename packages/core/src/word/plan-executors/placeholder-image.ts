/**
 * #190 占位图共享资源 —— PNG 字节 + 注入帮助。
 *
 * 设计:
 *   - 全包共享一份 `word/media/placeholder.png`(200x150 灰度,内含中部 X 标记)
 *   - 每个承载 drawing 的 part(`word/document.xml` / `word/header*.xml` 等)各自的
 *     `.rels` 里**最多注入一条** image relationship → placeholder.png,后续替换复用
 *   - [Content_Types].xml 首次时补 `<Default Extension="png" ContentType="image/png"/>`
 *     (已有相同 Default 或同名 Override 时跳过)
 *
 * 视觉表达:PNG 内是灰底 + 边框 + 中部 X 标记(国际通用"占位/图像缺失"图标语义)。
 * 中文"[图片占位]"标注由替换 executor 写入 `wp:docPr/@descr`(screen reader / hover
 * tooltip 可见),避免在 PNG 内手工渲染中文字形。
 *
 * 字节来源:`scripts/gen-placeholder-png.mjs` 一次性生成,baked 进 placeholder-image-bytes.ts。
 */
import { parseXmlDocument } from 'slimdom';
import type { Package } from '../../pack/index.js';
import {
  createBinaryPart,
  createXmlPart,
  deletePart,
  getPart,
  hasPart,
  isBinaryPart,
  isXmlPart,
  listParts,
  markDirty,
  setPart,
} from '../../pack/index.js';
import { PLACEHOLDER_PNG_BYTES } from './placeholder-image-bytes.js';

// ─────────── 常量 ───────────

/** 共享占位图在包内的路径 */
export const PLACEHOLDER_MEDIA_PATH = 'word/media/placeholder.png';
/** PNG MIME content type */
export const PNG_CONTENT_TYPE = 'image/png';
/** OOXML image relationship type(BLIP r:embed 的解析目标)*/
export const IMAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** 中文标注文字 —— 由 executor 写到 wp:docPr/@descr / @title */
export const PLACEHOLDER_DESCR = '[图片占位]';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

// ─────────── 公开 API ───────────

/**
 * 确保 `word/media/placeholder.png` 存在;不存在时创建。
 *
 * 幂等:多次调用安全,首次创建后续 short-circuit。
 *
 * @returns 是否本次刚创建(供调用方决定是否标记 [Content_Types] 等 dirty)
 */
export function ensurePlaceholderMediaPart(pkg: Package): boolean {
  if (hasPart(pkg, PLACEHOLDER_MEDIA_PATH)) return false;
  const part = createBinaryPart({
    path: PLACEHOLDER_MEDIA_PATH,
    contentType: PNG_CONTENT_TYPE,
    bytes: PLACEHOLDER_PNG_BYTES,
    // 不传 originalBytes 让其默认 = bytes;但显式 markDirty 让 toZip 触发实际写入
  });
  setPart(pkg, part);
  // createBinaryPart 默认 _dirty=false(装载即终态语义);新建场景必须 markDirty 才会被 toZip 序列化
  markDirty(part);
  return true;
}

/**
 * 确保 [Content_Types].xml 中 PNG 类型可解析。
 *
 * 优先策略:补 `<Default Extension="png" ContentType="image/png"/>`。
 *   - 若已有同 Extension 的 Default(无论 ContentType 是否一致)→ 不动
 *   - 若有 `<Override PartName="/word/media/placeholder.png" .../>` → 也算覆盖,不补 Default
 *
 * @returns 是否本次刚补
 */
export function ensurePngContentType(pkg: Package): boolean {
  const ctPart = getPart(pkg, CONTENT_TYPES_PATH);
  if (!ctPart || !isXmlPart(ctPart) || !ctPart.xmlDom) return false;

  const doc = ctPart.xmlDom;

  // 已有 Default Extension="png" → skip
  const defaults = doc.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Default');
  for (let i = 0; i < defaults.length; i++) {
    if (defaults[i]!.getAttribute('Extension')?.toLowerCase() === 'png') return false;
  }

  // 已有 Override 指向 placeholder.png → skip
  const overrides = doc.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
  for (let i = 0; i < overrides.length; i++) {
    if (overrides[i]!.getAttribute('PartName') === `/${PLACEHOLDER_MEDIA_PATH}`) return false;
  }

  // 补 Default
  const def = doc.createElementNS(CONTENT_TYPES_NS, 'Default');
  def.setAttribute('Extension', 'png');
  def.setAttribute('ContentType', PNG_CONTENT_TYPE);
  doc.documentElement.appendChild(def);
  markDirty(ctPart);
  return true;
}

/**
 * 确保指定 rels part 内有一条 image relationship 指向 placeholder.png;若已存在直接返回其 rId。
 *
 * @param pkg          包
 * @param relsPartPath 目标 rels part 路径,如 'word/_rels/document.xml.rels' / 'word/_rels/header1.xml.rels'
 *                     不存在时会创建空 rels(场景:某 header part 原本没有 .rels)
 * @returns 复用或新分配的 rId(形如 'rId7')
 */
export function ensurePlaceholderRelInPart(pkg: Package, relsPartPath: string): string {
  let relsPart = getPart(pkg, relsPartPath);
  if (!relsPart) {
    // 新建空 rels:某些 header / footer 原本无图,装载后第一次注入图片需要建 rels
    const empty = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}"/>`;
    const dom = parseXmlDocument(empty) as unknown as Document;
    relsPart = createXmlPart({
      path: relsPartPath,
      contentType: 'application/vnd.openxmlformats-package.relationships+xml',
      xmlDom: dom,
      // 不传 originalBytes → 默认 dirty=true(新建)
    });
    setPart(pkg, relsPart);
  }
  if (!isXmlPart(relsPart) || !relsPart.xmlDom) {
    throw new Error(`ensurePlaceholderRelInPart: '${relsPartPath}' is not an XML part`);
  }

  const doc = relsPart.xmlDom;
  const placeholderTarget = computeRelTarget(relsPartPath, PLACEHOLDER_MEDIA_PATH);

  // 已有指向 placeholder 的 image rel → 复用
  const rels = doc.getElementsByTagNameNS(RELS_NS, 'Relationship');
  let maxRId = 0;
  for (let i = 0; i < rels.length; i++) {
    const r = rels[i]!;
    const id = r.getAttribute('Id') ?? '';
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = parseInt(m[1]!, 10);
      if (n > maxRId) maxRId = n;
    }
    if (
      r.getAttribute('Type') === IMAGE_RELATIONSHIP_TYPE &&
      normalizeTarget(r.getAttribute('Target') ?? '') === placeholderTarget
    ) {
      return id;
    }
  }

  // 新增 rel
  const rId = `rId${maxRId + 1}`;
  const rel = doc.createElementNS(RELS_NS, 'Relationship');
  rel.setAttribute('Id', rId);
  rel.setAttribute('Type', IMAGE_RELATIONSHIP_TYPE);
  rel.setAttribute('Target', placeholderTarget);
  doc.documentElement.appendChild(rel);
  markDirty(relsPart);
  return rId;
}

/**
 * 把 .rels 文件内 `<Relationship>` 的 Target(相对路径)解析为包内绝对 part 路径。
 *
 * 规则:相对于宿主 part 的目录解析。
 *   - 'word/_rels/document.xml.rels' 内 Target='media/x.png' → 'word/media/x.png'
 *   - Target='/word/media/x.png'(以 '/' 开头的"包根绝对路径")→ 'word/media/x.png'
 *   - 顶层 '_rels/.rels' base 为包根
 */
export function resolveRelTarget(relsPartPath: string, target: string): string {
  let t = target.trim();
  if (t.startsWith('/')) return t.slice(1);
  while (t.startsWith('./')) t = t.slice(2);
  const base = hostDirOfRelsPart(relsPartPath);
  if (base === '') return t;
  // 处理 '../' 回溯(尾随段累计)
  let pathParts = base.split('/').filter(Boolean);
  while (t.startsWith('../')) {
    pathParts = pathParts.slice(0, -1);
    t = t.slice(3);
  }
  return pathParts.length === 0 ? t : `${pathParts.join('/')}/${t}`;
}

/**
 * 统计 media part 路径在全包所有 .rels 中的引用数。
 *
 * 用于孤儿 media 清理:执行器删除一条 rels 条目后,该 Relationship.Target 对应的
 * media path 可能仍被其它 .rels 引用(同图被多 part / 多 drawing 共享场景),
 * 仅在全包归零时才可真正删除 binary part。
 */
export function countMediaReferences(pkg: Package, mediaPath: string): number {
  let count = 0;
  for (const part of listParts(pkg)) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    if (!part.path.endsWith('.rels')) continue;
    const rels = part.xmlDom.getElementsByTagNameNS(RELS_NS, 'Relationship');
    for (let i = 0; i < rels.length; i++) {
      const target = rels[i]!.getAttribute('Target');
      if (!target) continue;
      if (resolveRelTarget(part.path, target) === mediaPath) count++;
    }
  }
  return count;
}

/**
 * 把候选孤儿 media 路径过一遍,只在全包真正归零时才 deletePart。
 *
 * @returns 实际删除的 path 列表(便于日志 / 测试断言)
 */
export function sweepOrphanMediaParts(pkg: Package, candidates: Iterable<string>): string[] {
  const removed: string[] = [];
  for (const path of candidates) {
    const part = getPart(pkg, path);
    if (!part) continue;
    if (!isBinaryPart(part)) continue;
    if (countMediaReferences(pkg, path) > 0) continue;
    deletePart(pkg, path);
    removed.push(path);
  }
  return removed;
}

// ─────────── 内部 ───────────

/**
 * 计算 rels 文件里 Relationship/@Target 应写的相对路径。
 *
 * OOXML rel Target 默认相对于宿主 part 的目录解析:
 *   - 'word/document.xml' 的 rels → base 'word/'
 *   - 'word/_rels/document.xml.rels' 内 Target='media/x.png' → 'word/media/x.png'
 *
 * 我们的 placeholder 在 'word/media/placeholder.png',对所有 'word/*.xml.rels' 来说
 * 都是 'media/placeholder.png'。对顶层 '_rels/.rels'(基 = 包根)需要 'word/media/placeholder.png'
 * —— 但 #190 的 drawing 不可能在顶层 rels 里出现,只保留兼容。
 */
function computeRelTarget(relsPartPath: string, targetPartPath: string): string {
  // 把 rels part 路径转成 host part 路径,再取目录做 base
  // 'word/_rels/document.xml.rels' → host = 'word/document.xml' → base = 'word/'
  // '_rels/.rels'                  → host = '.rels'             → base = ''
  const hostDir = hostDirOfRelsPart(relsPartPath);
  if (hostDir === '') return targetPartPath;
  if (targetPartPath.startsWith(`${hostDir}/`)) {
    return targetPartPath.slice(hostDir.length + 1);
  }
  // base 与目标不同前缀(罕见),退化为以 '../' 回溯。这里 #190 实际只走 word/ 前缀,
  // 简化处理:回 '/' 绝对路径形态(OOXML 允许)。
  return `/${targetPartPath}`;
}

function hostDirOfRelsPart(relsPath: string): string {
  // 找到 '/_rels/' 段:其前部 = host 所在目录
  const idx = relsPath.lastIndexOf('/_rels/');
  if (idx < 0) {
    // 如 '_rels/.rels' 形态:host 在包根
    return '';
  }
  return relsPath.slice(0, idx);
}

/** 标准化 Target 路径用于比较(去开头 './',统一斜杠,去重复斜杠)*/
function normalizeTarget(t: string): string {
  let s = t.trim();
  while (s.startsWith('./')) s = s.slice(2);
  return s.replace(/\\/g, '/');
}
