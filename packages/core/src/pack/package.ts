/**
 * Package 集合层:create / set / get / has / delete / list
 *
 * Package 只是 path→Part 的 Map 容器;零业务语义,存储格式由 adapter 层(#9 zip / #10 Flat OPC)决定。
 */
import type { Package, Part } from './types.js';

/** 空包。后续由 adapter 层填充 parts */
export function createPackage(): Package {
  return { parts: new Map() };
}

/** 以 part.path 为 key 放入 pkg。已存在同 path 时覆盖 */
export function setPart(pkg: Package, part: Part): void {
  pkg.parts.set(part.path, part);
}

/** 按 path 取回 part;不存在返 undefined */
export function getPart(pkg: Package, path: string): Part | undefined {
  return pkg.parts.get(path);
}

/** 按 path 判断存在 */
export function hasPart(pkg: Package, path: string): boolean {
  return pkg.parts.has(path);
}

/** 删除 part,返回是否删除成功(存在即删、不存在返 false) */
export function deletePart(pkg: Package, path: string): boolean {
  return pkg.parts.delete(path);
}

/** 所有 parts 快照数组(稳定遍历顺序 = Map 插入顺序) */
export function listParts(pkg: Package): Part[] {
  return Array.from(pkg.parts.values());
}
