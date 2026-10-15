/**
 * w14:paraId 装载时 normalize —— docs/03 §11.2 / §13.6。
 *
 * 语义:
 *   - 扫 Package 里所有 XML part,对缺 `w14:paraId` 的 `<w:p>` 分配随机 8 位十六进制 id
 *   - **package 全局唯一**(body / headers / footers / footnotes / endnotes / comments 同集合)
 *   - 发生分配的 part 被 markDirty,便于后续 toZip / toFlatOpc 重 serialize
 *   - 10 万段落 ≈ 40ms(实测基线,见 bench/xml-stack-worker)
 */
import type { Package } from '../pack/index.js';
import { isXmlPart, markDirty } from '../pack/index.js';
import { W_NS, W14_NS } from './namespaces.js';

export interface EnsureParaIdsResult {
  /** 本次新分配的 paraId 数量 */
  added: number;
  /** 扫描到的 <w:p> 总数(含既有 paraId) */
  total: number;
  /** 发生分配的 part path 列表 */
  touchedParts: string[];
}

/** 扫全包收集已有的 paraId,保证新分配不冲突 */
function collectExistingParaIds(pkg: Package): Set<string> {
  const set = new Set<string>();
  for (const part of pkg.parts.values()) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    const paragraphs = part.xmlDom.getElementsByTagNameNS(W_NS, 'p');
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      if (!p) continue;
      const id = p.getAttributeNS(W14_NS, 'paraId');
      if (id) set.add(id.toUpperCase());
    }
  }
  return set;
}

/**
 * 生成 8 位十六进制 paraId,冲突重试。
 *
 * 使用 crypto.getRandomValues 取 4 字节(32-bit),转十六进制字符串,大写。
 * export 供 executor(#29 等)在 apply 时分配新段落 id 使用。
 */
export function generateParaId(seen: Set<string>, maxAttempts = 100): string {
  const bytes = new Uint8Array(4);
  for (let i = 0; i < maxAttempts; i++) {
    globalThis.crypto.getRandomValues(bytes);
    const hex =
      bytes[0]!.toString(16).padStart(2, '0') +
      bytes[1]!.toString(16).padStart(2, '0') +
      bytes[2]!.toString(16).padStart(2, '0') +
      bytes[3]!.toString(16).padStart(2, '0');
    const id = hex.toUpperCase();
    if (!seen.has(id)) {
      seen.add(id);
      return id;
    }
  }
  throw new Error(`ensureParaIds: unable to generate unique paraId after ${maxAttempts} attempts`);
}

/**
 * 为 Package 内所有缺 w14:paraId 的 <w:p> 分配随机 id(package 全局唯一)。
 *
 * 发生分配的 part 会被 markDirty。
 */
export function ensureParaIds(pkg: Package): EnsureParaIdsResult {
  const seen = collectExistingParaIds(pkg);
  let added = 0;
  let total = 0;
  const touched = new Set<string>();

  for (const part of pkg.parts.values()) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    const paragraphs = part.xmlDom.getElementsByTagNameNS(W_NS, 'p');
    let partTouched = false;
    // 快照成数组,避免下面 setAttributeNS 导致 live collection 行为变化
    const snapshot: Element[] = [];
    for (let i = 0; i < paragraphs.length; i++) {
      const el = paragraphs[i];
      if (el) snapshot.push(el);
    }
    total += snapshot.length;
    for (const p of snapshot) {
      const id = p.getAttributeNS(W14_NS, 'paraId');
      if (!id) {
        const newId = generateParaId(seen);
        p.setAttributeNS(W14_NS, 'w14:paraId', newId);
        added++;
        partTouched = true;
      }
    }
    if (partTouched) {
      markDirty(part);
      touched.add(part.path);
    }
  }

  return { added, total, touchedParts: Array.from(touched) };
}
