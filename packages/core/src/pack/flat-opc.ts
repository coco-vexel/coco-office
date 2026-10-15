/**
 * Flat OPC ↔ Package 互转(docs/01 §3.3)。
 *
 * Flat OPC 是 WPS `Document.WordOpenXML` 返回的形态,整个包展开成一份大 XML:
 *
 *   <?xml version="1.0" standalone="yes"?>
 *   <?mso-application progid="Word.Document"?>
 *   <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
 *     <pkg:part pkg:name="/word/document.xml" pkg:contentType="...+xml">
 *       <pkg:xmlData>
 *         <w:document xmlns:w="...">...</w:document>
 *       </pkg:xmlData>
 *     </pkg:part>
 *     <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png" pkg:compression="store">
 *       <pkg:binaryData>iVBORw0KG...=</pkg:binaryData>
 *     </pkg:part>
 *     ...
 *   </pkg:package>
 *
 * **注意**:Flat OPC 里 contentType 直接写在每个 `<pkg:part>` 上,
 * 没有 `[Content_Types].xml` 这个 part 本身。互转时 `[Content_Types].xml` 只在 zip 形态下存在。
 */
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import type { Package, Part } from './types.js';
import { createBinaryPart, createXmlPart, isXmlPart } from './part.js';
import { createPackage, setPart } from './package.js';

const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const CONTENT_TYPES_PATH = '[Content_Types].xml';

function isXmlContentType(ct: string): boolean {
  return ct.endsWith('+xml') || ct === 'application/xml' || ct === 'text/xml';
}

/** Uint8Array → base64(兼容 Node 16+ / 浏览器 / WPS WebView) */
function base64Encode(bytes: Uint8Array): string {
  let binStr = '';
  for (let i = 0; i < bytes.length; i++) {
    binStr += String.fromCharCode(bytes[i]!);
  }
  return btoa(binStr);
}

/** base64 → Uint8Array */
function base64Decode(b64: string): Uint8Array {
  const cleaned = b64.replace(/\s+/g, '');
  const binStr = atob(cleaned);
  const bytes = new Uint8Array(binStr.length);
  for (let i = 0; i < binStr.length; i++) {
    bytes[i] = binStr.charCodeAt(i);
  }
  return bytes;
}

/** 取元素的第一个 Element 类型子节点(跳过文本 / 注释 / 空白) */
function firstElementChild(el: Element): Element | null {
  const kids = el.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const n = kids[i];
    if (n && n.nodeType === 1) return n as Element;
  }
  return null;
}

/**
 * 从 Flat OPC 字符串重建 Package。
 *
 * - 忽略顶部 `<?mso-application?>` PI(解析器自动处理)
 * - `<pkg:xmlData>` 里的根元素被提取重新序列化为该 part 的独立 XML 文档
 * - `<pkg:binaryData>` 做 base64 解码
 * - **不**产生 `[Content_Types].xml` part(Flat OPC 形态不含它)
 */
export function fromFlatOpc(xml: string): Package {
  const flatDoc = parseXmlDocument(xml) as unknown as Document;
  const pkg = createPackage();

  const partEls = flatDoc.getElementsByTagNameNS(PKG_NS, 'part');
  for (let i = 0; i < partEls.length; i++) {
    const partEl = partEls[i];
    if (!partEl) continue;

    const pkgName = partEl.getAttributeNS(PKG_NS, 'name');
    const ctype = partEl.getAttributeNS(PKG_NS, 'contentType');
    if (!pkgName || !ctype) {
      throw new Error(`fromFlatOpc: <pkg:part> missing name or contentType`);
    }
    const path = pkgName.replace(/^\//, '');
    if (!path) continue;

    // [Content_Types].xml 理论上不会出现在 Flat OPC 里;若遇到,跳过
    if (path === CONTENT_TYPES_PATH) continue;

    // 查找 xmlData 或 binaryData(只取第一个,跨命名空间断言)
    let xmlData: Element | null = null;
    let binaryData: Element | null = null;
    const kids = partEl.childNodes;
    for (let j = 0; j < kids.length; j++) {
      const k = kids[j];
      if (!k || k.nodeType !== 1) continue;
      const e = k as Element;
      if (e.namespaceURI !== PKG_NS) continue;
      if (e.localName === 'xmlData' && !xmlData) xmlData = e;
      else if (e.localName === 'binaryData' && !binaryData) binaryData = e;
    }

    if (xmlData) {
      const rootChild = firstElementChild(xmlData);
      if (!rootChild) {
        throw new Error(`fromFlatOpc: empty <pkg:xmlData> for part '${path}'`);
      }
      // 将子元素序列化出来,再独立解析为该 part 的 Document
      const xmlText = serializeToWellFormedString(rootChild as never);
      const originalBytes = new TextEncoder().encode(xmlText);
      const dom = parseXmlDocument(xmlText) as unknown as Document;
      setPart(pkg, createXmlPart({ path, contentType: ctype, xmlDom: dom, originalBytes }));
    } else if (binaryData) {
      const b64 = binaryData.textContent ?? '';
      const bytes = base64Decode(b64);
      setPart(pkg, createBinaryPart({ path, contentType: ctype, bytes }));
    } else {
      // 既无 xmlData 也无 binaryData — 按类型决定:XML contentType 用空元素,二进制用空字节
      if (isXmlContentType(ctype)) {
        throw new Error(`fromFlatOpc: XML part '${path}' has no <pkg:xmlData>`);
      }
      setPart(pkg, createBinaryPart({ path, contentType: ctype, bytes: new Uint8Array() }));
    }
  }

  return pkg;
}

/**
 * 将 Package 序列化为 Flat OPC 字符串。
 *
 * - 顶部插入 `<?mso-application progid="Word.Document"?>` PI
 * - 跳过 `[Content_Types].xml` part(Flat OPC 不需要)
 * - XML part:导入 `xmlDom` 的根元素到新 Flat OPC document,包进 `<pkg:xmlData>`
 * - 二进制 part:字节 base64 编码,包进 `<pkg:binaryData>`,加 `pkg:compression="store"`
 */
export function toFlatOpc(pkg: Package): string {
  // slimdom 的 DOMImplementation 构造签名与 lib.dom 不同(需要一个虚拟 document 作为 context);
  // 直接用 parseXmlDocument 生成一个带 pkg:package 根的空文档,避免跨类型适配。
  const seed = `<pkg:package xmlns:pkg="${PKG_NS}"/>`;
  const flatDoc = parseXmlDocument(seed) as unknown as Document;
  const root = flatDoc.documentElement;
  if (!root) {
    throw new Error('toFlatOpc: failed to create <pkg:package> root');
  }

  for (const part of pkg.parts.values()) {
    if (part.path === CONTENT_TYPES_PATH) continue;

    const partEl = flatDoc.createElementNS(PKG_NS, 'pkg:part');
    partEl.setAttributeNS(PKG_NS, 'pkg:name', '/' + part.path);
    partEl.setAttributeNS(PKG_NS, 'pkg:contentType', part.contentType);

    if (isXmlPart(part)) {
      appendXmlData(flatDoc, partEl, part);
    } else {
      partEl.setAttributeNS(PKG_NS, 'pkg:compression', 'store');
      appendBinaryData(flatDoc, partEl, part);
    }

    root.appendChild(partEl);
  }

  // 顶部加 <?mso-application?> PI
  const pi = flatDoc.createProcessingInstruction('mso-application', 'progid="Word.Document"');
  flatDoc.insertBefore(pi as never, flatDoc.documentElement as never);

  return serializeToWellFormedString(flatDoc as never);
}

function appendXmlData(flatDoc: Document, partEl: Element, part: Part): void {
  if (!part.xmlDom || !part.xmlDom.documentElement) {
    throw new Error(`toFlatOpc: XML part '${part.path}' missing xmlDom or root`);
  }
  const xmlData = flatDoc.createElementNS(PKG_NS, 'pkg:xmlData');
  // importNode 深拷贝原 part 的根元素到 Flat OPC document,命名空间保留
  const imported = flatDoc.importNode(
    part.xmlDom.documentElement as never,
    true,
  ) as unknown as Node;
  xmlData.appendChild(imported);
  partEl.appendChild(xmlData);
}

function appendBinaryData(flatDoc: Document, partEl: Element, part: Part): void {
  const bytes = part.bytes ?? part._originalBytes;
  if (!bytes) {
    throw new Error(`toFlatOpc: binary part '${part.path}' has no bytes`);
  }
  const binEl = flatDoc.createElementNS(PKG_NS, 'pkg:binaryData');
  binEl.textContent = base64Encode(bytes);
  partEl.appendChild(binEl);
}
