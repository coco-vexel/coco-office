/**
 * zip ↔ Package 互转(docs/01 §3.2)。
 *
 * - `fromZip(buf)` 解压 .docx,按 `[Content_Types].xml` 恢复每个 part 的 contentType
 * - `toZip(pkg)` 反向:dirty=false 的 part 复用 `_originalBytes`,dirty=true 的按 XML / 二进制重序列化
 *
 * 本实现目标:小文档正确性。极大档性能调优(lazy parse 等)归后续 issue。
 */
import JSZip from 'jszip';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import type { Package, Part } from './types.js';
import { createBinaryPart, createXmlPart, isXmlPart } from './part.js';
import { createPackage, setPart } from './package.js';
import { ensureContentTypes } from './content-types.js';

/** Content Types 子集 */
const CONTENT_TYPES_PATH = '[Content_Types].xml';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const CONTENT_TYPES_CTYPE = 'application/vnd.openxmlformats-package.content-types+xml';

/**
 * 判断 contentType 是否为 XML 类(决定走 parseXmlDocument 还是保留字节):
 *   - 以 `+xml` 结尾(OOXML 所有子 part 的惯例)
 *   - 或完全等于 `application/xml` / `text/xml`
 */
function isXmlContentType(ct: string): boolean {
  const normalized = ct.toLowerCase();
  if (
    normalized.startsWith('image/') ||
    normalized.startsWith('audio/') ||
    normalized.startsWith('video/') ||
    normalized === 'application/octet-stream' ||
    normalized === 'application/vnd.openxmlformats-officedocument.oleobject' ||
    normalized.startsWith('application/vnd.ms-office.')
  ) {
    return false;
  }
  return normalized.endsWith('+xml') || normalized === 'application/xml' || normalized === 'text/xml';
}

/** 解析 [Content_Types].xml 构建 partName → contentType 的查找器 */
interface ContentTypeResolver {
  /** partName 无前导 `/`,如 'word/document.xml' */
  resolve(partName: string): string | undefined;
}

function parseContentTypes(xml: string): ContentTypeResolver {
  const doc = parseXmlDocument(xml) as unknown as Document;
  const defaults = new Map<string, string>(); // extension(小写) → contentType
  const overrides = new Map<string, string>(); // partName(无 /)→ contentType

  const defaultEls = doc.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Default');
  for (let i = 0; i < defaultEls.length; i++) {
    const el = defaultEls[i];
    if (!el) continue;
    const ext = el.getAttribute('Extension');
    const ct = el.getAttribute('ContentType');
    if (ext && ct) defaults.set(ext.toLowerCase(), ct);
  }

  const overrideEls = doc.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
  for (let i = 0; i < overrideEls.length; i++) {
    const el = overrideEls[i];
    if (!el) continue;
    const partName = el.getAttribute('PartName');
    const ct = el.getAttribute('ContentType');
    if (partName && ct) {
      // PartName 以 `/` 开头,剥掉前导斜杠与 Package 内部 path 对齐
      overrides.set(partName.replace(/^\//, ''), ct);
    }
  }

  return {
    resolve(partName: string): string | undefined {
      const override = overrides.get(partName);
      if (override) return override;
      const dot = partName.lastIndexOf('.');
      if (dot === -1) return undefined;
      return defaults.get(partName.slice(dot + 1).toLowerCase());
    },
  };
}

/**
 * 从 zip 字节重建 Package。
 *
 * 行为:
 *   - 先读 `[Content_Types].xml`,构建 contentType 查找器
 *   - 遍历 zip entries,跳过目录
 *   - 每个 entry 按 contentType 决定 XML / 二进制,创建对应 Part
 *   - `[Content_Types].xml` 也作为 XML part 存入 Package
 */
export async function fromZip(buf: Uint8Array | ArrayBuffer): Promise<Package> {
  const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  const zip = await JSZip.loadAsync(bytes);
  const pkg = createPackage();

  const ctFile = zip.file(CONTENT_TYPES_PATH);
  if (!ctFile) {
    throw new Error(`fromZip: missing ${CONTENT_TYPES_PATH}; not a valid OOXML package`);
  }
  const ctXml = await ctFile.async('string');
  const resolver = parseContentTypes(ctXml);
  const ctBytes = await ctFile.async('uint8array');
  setPart(
    pkg,
    createXmlPart({
      path: CONTENT_TYPES_PATH,
      contentType: CONTENT_TYPES_CTYPE,
      xmlDom: parseXmlDocument(ctXml) as unknown as Document,
      originalBytes: ctBytes,
    }),
  );

  // 遍历非 [Content_Types].xml 的 entry
  const entryPaths: string[] = [];
  zip.forEach((relativePath, entry) => {
    if (entry.dir) return;
    if (relativePath === CONTENT_TYPES_PATH) return;
    entryPaths.push(relativePath);
  });

  for (const path of entryPaths) {
    const entry = zip.file(path);
    if (!entry) continue;
    const contentType = resolver.resolve(path);
    if (!contentType) {
      throw new Error(
        `fromZip: no ContentType for part '${path}' (not in Default by extension, not in Override by partName)`,
      );
    }
    const entryBytes = await entry.async('uint8array');

    if (isXmlContentType(contentType)) {
      const xmlText = await entry.async('string');
      const xmlDom = parseXmlDocument(xmlText) as unknown as Document;
      setPart(pkg, createXmlPart({ path, contentType, xmlDom, originalBytes: entryBytes }));
    } else {
      setPart(pkg, createBinaryPart({ path, contentType, bytes: entryBytes }));
    }
  }

  return pkg;
}

/**
 * 将 Package 序列化回 zip 字节。
 *
 * 行为:
 *   - `_dirty === false` 且持有 `_originalBytes` → 直接复用字节(零 serialize)
 *   - `_dirty === true` 的 XML part → slimdom serialize
 *   - `_dirty === true` 的二进制 part → 使用 `bytes`
 */
export async function toZip(pkg: Package): Promise<Uint8Array> {
  // Flat OPC → Package 的路径不会产生 [Content_Types].xml;此处补齐以保证产出的 zip 合法。
  // 对 fromZip 出来的 Package 无副作用(CT 已在)。
  ensureContentTypes(pkg);

  const zip = new JSZip();

  for (const part of pkg.parts.values()) {
    const bytes = partToBytes(part);
    // createFolders=false 显式禁止 JSZip 为含 '/' 的 path 自动创建中间目录 entry——
    // OOXML 规范不允许 zip 包内有空目录 entry,Word 见到会判"文档已损坏"拒打开
    // (WPS 宽松能开但产物不合规)。JSZip 3.x 默认 true,必须显式关。
    zip.file(part.path, bytes, { createFolders: false });
  }

  // OOXML 主体是高度可压缩的 XML;默认 STORE 会让产物比 Word 原生输出大 10-18×。
  return await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

function partToBytes(part: Part): Uint8Array {
  // dirty=false 且有原始字节:零 serialize
  if (!part._dirty && part._originalBytes) {
    return part._originalBytes;
  }

  if (isXmlPart(part)) {
    if (!part.xmlDom) {
      throw new Error(`toZip: XML part '${part.path}' missing xmlDom`);
    }
    // xmlDom 在我们的类型里是 lib.dom 的 Document(消费方友好),slimdom 的 Node 结构兼容但 TS
    // 类型不同;通过 `as never` 绕开 dts 检查,运行时正常。
    const text = serializeToWellFormedString(part.xmlDom as never);
    return new TextEncoder().encode(text);
  }

  // 二进制
  if (part.bytes) return part.bytes;
  if (part._originalBytes) return part._originalBytes;
  throw new Error(`toZip: binary part '${part.path}' has no bytes`);
}
