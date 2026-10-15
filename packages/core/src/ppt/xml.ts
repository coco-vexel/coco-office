import type { Part } from '../pack/index.js';

export const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface PptRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export function childElementsNs(parent: Element, ns: string, localName: string): Element[] {
  const result: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI === ns && el.localName === localName) result.push(el);
  }
  return result;
}

export function firstChildNs(parent: Element, ns: string, localName: string): Element | null {
  return childElementsNs(parent, ns, localName)[0] ?? null;
}

export function firstChildLocal(parent: Element, localName: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child?.nodeType === 1 && (child as Element).localName === localName) {
      return child as Element;
    }
  }
  return null;
}

export function allDescendantsNs(
  parent: Element | Document,
  ns: string,
  localName: string,
): Element[] {
  return Array.from(parent.getElementsByTagNameNS(ns, localName)) as Element[];
}

export function allDescendantsLocal(parent: Element | Document, localName: string): Element[] {
  return Array.from(parent.getElementsByTagName('*')).filter((el) => el.localName === localName);
}

export function textOf(el: Element | null | undefined): string {
  return el?.textContent ?? '';
}

export function parseRelationships(part?: Part): Map<string, PptRelationship> {
  const result = new Map<string, PptRelationship>();
  if (!part?.xmlDom) return result;
  const rels = part.xmlDom.getElementsByTagNameNS(PKG_REL_NS, 'Relationship');
  for (let i = 0; i < rels.length; i++) {
    const rel = rels[i];
    if (!rel) continue;
    const id = rel.getAttribute('Id');
    const type = rel.getAttribute('Type');
    const target = rel.getAttribute('Target');
    if (!id || !type || !target) continue;
    result.set(id, {
      id,
      type,
      target,
      targetMode: rel.getAttribute('TargetMode') ?? undefined,
    });
  }
  return result;
}

export function dirname(path: string): string {
  const idx = path.lastIndexOf('/');
  return idx === -1 ? '' : path.slice(0, idx);
}

export function basename(path: string): string {
  const idx = path.lastIndexOf('/');
  return idx === -1 ? path : path.slice(idx + 1);
}

export function normalizePath(path: string): string {
  const parts: string[] = [];
  for (const part of path.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop();
    else parts.push(part);
  }
  return parts.join('/');
}

export function resolveRelationshipTarget(sourcePartPath: string, target: string): string {
  if (target.startsWith('/')) return normalizePath(target.slice(1));
  const base = dirname(sourcePartPath);
  return normalizePath(base ? `${base}/${target}` : target);
}

export function relsPathForPart(partPath: string): string {
  const dir = dirname(partPath);
  const name = basename(partPath);
  return dir ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

export function parseNumber(raw: string | null): number | undefined {
  if (raw === null || raw === '') return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

export function setAttrIfDefined(el: Element, name: string, value: number | undefined): void {
  if (value !== undefined && Number.isFinite(value)) el.setAttribute(name, String(Math.round(value)));
}
