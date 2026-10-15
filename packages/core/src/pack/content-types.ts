/**
 * [Content_Types].xml 的合成与维护。
 *
 * 为什么单列一个模块:
 *   - zip 形态要求 [Content_Types].xml 存在(否则 Word/WPS 打不开)
 *   - Flat OPC 形态里 contentType 已内联在 <pkg:part> 上,不含 CT 文件
 *   - 因此 `fromFlatOpc` 出来的 Package 缺 CT;直接 `toZip` 会产出坏 docx
 *
 * 本模块提供 `ensureContentTypes(pkg)`:若 CT part 缺失,按当前 parts 的
 * (path → contentType)合成一份最小 CT XML 并加入 Package。
 *
 * `toZip` 会在打包前自动调一次,确保产出的 zip 始终可打开。
 */
import { parseXmlDocument } from 'slimdom';
import { createXmlPart } from './part.js';
import { getPart, setPart } from './package.js';
import type { Package } from './types.js';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const CONTENT_TYPES_CTYPE = 'application/vnd.openxmlformats-package.content-types+xml';

/** XML 属性值转义(Content Type 里可能含 `+`,但没有 `<` `&` 等需要转义的字符;仍做一遍以防) */
function xmlAttr(v: string): string {
  return v
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * 如果 Package 里已有 `[Content_Types].xml` part,返回 false,不做改动。
 * 否则,按当前 parts 的 contentType 合成一份最小 CT XML 并加入 Package,返回 true。
 *
 * 合成策略:每个非 CT part 产生一个 `<Override PartName="/path" ContentType="..."/>`。
 * 不尝试推导 Default 扩展名(正确性优先,简洁性次之)。
 */
export function ensureContentTypes(pkg: Package): boolean {
  if (getPart(pkg, CONTENT_TYPES_PATH)) return false;

  const overrides: string[] = [];
  for (const part of pkg.parts.values()) {
    if (part.path === CONTENT_TYPES_PATH) continue;
    overrides.push(
      `  <Override PartName="/${xmlAttr(part.path)}" ContentType="${xmlAttr(part.contentType)}"/>`,
    );
  }

  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Types xmlns="${CONTENT_TYPES_NS}">\n` +
    overrides.join('\n') +
    (overrides.length > 0 ? '\n' : '') +
    `</Types>`;

  setPart(
    pkg,
    createXmlPart({
      path: CONTENT_TYPES_PATH,
      contentType: CONTENT_TYPES_CTYPE,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
      originalBytes: new TextEncoder().encode(xml),
    }),
  );

  return true;
}
