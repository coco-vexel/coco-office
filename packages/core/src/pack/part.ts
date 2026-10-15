/**
 * Part 构造 / 变更 / 查询 的辅助函数。
 *
 * 所有 Part 的"创建"都走这里(而非直接字面量),以便:
 *   - 守住 path 格式(无前导斜杠)
 *   - 守住 XML / 二进制 互斥
 *   - 统一 `_dirty` / `_originalBytes` 初始化语义
 */
import type { Part } from './types.js';

export interface CreateXmlPartSpec {
  /** Part 路径,无前导斜杠,如 'word/document.xml' */
  path: string;
  /** MIME content type */
  contentType: string;
  /** 已解析好的 XML DOM(slimdom Document) */
  xmlDom: Document;
  /**
   * 装载时的原始字节。
   * - 传入(来自装载路径):`_dirty` 初始为 `false`,可在 `toZip` 时零 serialize 复用
   * - 省略(新建 part 的场景,后续 #9/#10 的插入新 XML part):`_dirty` 初始为 `true`
   */
  originalBytes?: Uint8Array;
}

export interface CreateBinaryPartSpec {
  /** Part 路径,无前导斜杠,如 'word/media/image1.png' */
  path: string;
  /** MIME content type,如 'image/png' */
  contentType: string;
  /** 二进制载荷 */
  bytes: Uint8Array;
  /**
   * 装载时的原始字节;二进制通常装载即最终态,省略时默认 = `bytes`。
   * 二进制 part 初始 `_dirty` 恒为 `false`。
   */
  originalBytes?: Uint8Array;
}

function assertPath(path: string): void {
  if (typeof path !== 'string' || path.length === 0) {
    throw new Error(`invalid part path: must be non-empty string, got: ${JSON.stringify(path)}`);
  }
  if (path.startsWith('/')) {
    throw new Error(`invalid part path: must not start with '/', got: ${path}`);
  }
}

function assertContentType(contentType: string): void {
  if (typeof contentType !== 'string' || contentType.length === 0) {
    throw new Error(`invalid contentType: must be non-empty string`);
  }
}

/** 从已解析的 XML DOM 创建 XML part。`originalBytes` 传入 → dirty=false;否则 dirty=true(新建场景) */
export function createXmlPart(spec: CreateXmlPartSpec): Part {
  assertPath(spec.path);
  assertContentType(spec.contentType);
  if (!spec.xmlDom) {
    throw new Error(`createXmlPart: xmlDom required for ${spec.path}`);
  }
  return {
    path: spec.path,
    contentType: spec.contentType,
    xmlDom: spec.xmlDom,
    _originalBytes: spec.originalBytes,
    _dirty: spec.originalBytes === undefined,
  };
}

/** 从二进制字节创建 binary part。`originalBytes` 省略时默认 = `bytes`。dirty 恒为 false */
export function createBinaryPart(spec: CreateBinaryPartSpec): Part {
  assertPath(spec.path);
  assertContentType(spec.contentType);
  if (!spec.bytes) {
    throw new Error(`createBinaryPart: bytes required for ${spec.path}`);
  }
  return {
    path: spec.path,
    contentType: spec.contentType,
    bytes: spec.bytes,
    _originalBytes: spec.originalBytes ?? spec.bytes,
    _dirty: false,
  };
}

/** 标记 part 被改动过。toZip / toFlatOpc 会重 serialize 此 part */
export function markDirty(part: Part): void {
  part._dirty = true;
}

/** 查询 part 是否被改动(read-only;读操作不改 part) */
export function isDirty(part: Part): boolean {
  return part._dirty;
}

/** XML part 判定(拥有 xmlDom) */
export function isXmlPart(part: Part): boolean {
  return part.xmlDom !== undefined;
}

/** 二进制 part 判定(拥有 bytes 但无 xmlDom) */
export function isBinaryPart(part: Part): boolean {
  return part.bytes !== undefined && part.xmlDom === undefined;
}
