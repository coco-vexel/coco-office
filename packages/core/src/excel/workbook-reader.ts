import type { Package, Part } from '../pack/index.js';
import { getPart } from '../pack/index.js';
import type { ExcelDefinedNameModel } from './model.js';
import { OFFICE_REL_NS, REL_NS, SS_NS } from './namespaces.js';

interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export interface WorkbookSheetRef {
  id: string;
  name: string;
  visibility?: string;
  relationshipId?: string;
  path?: string;
}

function requireXmlPart(pkg: Package, path: string): Part {
  const part = getPart(pkg, path);
  if (!part?.xmlDom) {
    throw new Error(`ExcelWorkbook: missing XML part ${path}`);
  }
  return part;
}

function optionalXmlPart(pkg: Package, path: string): Part | undefined {
  const part = getPart(pkg, path);
  return part?.xmlDom ? part : undefined;
}

function textOf(el: Element | null | undefined): string {
  if (!el) return '';
  return el.textContent ?? '';
}

function firstChildNs(parent: Element, ns: string, localName: string): Element | null {
  return parent.getElementsByTagNameNS(ns, localName)[0] ?? null;
}

function childElementsNs(parent: Element, ns: string, localName: string): Element[] {
  const elements = parent.getElementsByTagNameNS(ns, localName);
  return Array.from(elements).filter((el) => el.parentElement === parent);
}

function parseRelationships(part?: Part): Map<string, Relationship> {
  const rels = new Map<string, Relationship>();
  if (!part?.xmlDom) return rels;
  const relEls = part.xmlDom.getElementsByTagNameNS(REL_NS, 'Relationship');
  for (let i = 0; i < relEls.length; i++) {
    const rel = relEls[i];
    if (!rel) continue;
    const id = rel.getAttribute('Id');
    const type = rel.getAttribute('Type');
    const target = rel.getAttribute('Target');
    if (!id || !type || !target) continue;
    const model: Relationship = { id, type, target };
    const targetMode = rel.getAttribute('TargetMode');
    if (targetMode) model.targetMode = targetMode;
    rels.set(id, model);
  }
  return rels;
}

function dirname(path: string): string {
  const index = path.lastIndexOf('/');
  return index === -1 ? '' : path.slice(0, index);
}

function normalizePath(path: string): string {
  const parts: string[] = [];
  for (const part of path.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {
      parts.pop();
      continue;
    }
    parts.push(part);
  }
  return parts.join('/');
}

function resolveRelationshipTarget(sourcePartPath: string, target: string): string {
  if (target.startsWith('/')) return normalizePath(target.slice(1));
  return normalizePath(`${dirname(sourcePartPath)}/${target}`);
}

export function parseWorkbookSheets(pkg: Package): WorkbookSheetRef[] {
  const workbook = requireXmlPart(pkg, 'xl/workbook.xml');
  const rels = parseRelationships(optionalXmlPart(pkg, 'xl/_rels/workbook.xml.rels'));
  const sheetsEl = firstChildNs(workbook.xmlDom!.documentElement, SS_NS, 'sheets');
  if (!sheetsEl) return [];

  return childElementsNs(sheetsEl, SS_NS, 'sheet').map((sheetEl, index) => {
    const relationshipId = sheetEl.getAttributeNS(OFFICE_REL_NS, 'id') ?? undefined;
    const rel = relationshipId ? rels.get(relationshipId) : undefined;
    const path = rel ? resolveRelationshipTarget('xl/workbook.xml', rel.target) : undefined;
    const name = sheetEl.getAttribute('name') ?? `Sheet${index + 1}`;
    const id = sheetEl.getAttribute('sheetId') ?? String(index + 1);
    const visibility = sheetEl.getAttribute('state') ?? undefined;
    return { id, name, visibility, relationshipId, path };
  });
}

export function parseDefinedNames(pkg: Package): ExcelDefinedNameModel[] {
  const workbookPart = requireXmlPart(pkg, 'xl/workbook.xml');
  const definedNameEls = workbookPart.xmlDom!.getElementsByTagNameNS(SS_NS, 'definedName');
  const result: ExcelDefinedNameModel[] = [];
  for (let i = 0; i < definedNameEls.length; i++) {
    const el = definedNameEls[i];
    if (!el) continue;
    const name = el.getAttribute('name');
    if (!name) continue;
    const model: ExcelDefinedNameModel = {
      name,
      text: textOf(el),
    };
    const localSheetId = el.getAttribute('localSheetId');
    if (localSheetId !== null && localSheetId !== '') {
      const value = Number(localSheetId);
      if (Number.isFinite(value)) model.localSheetId = value;
    }
    result.push(model);
  }
  return result;
}

export function parseSharedStrings(pkg: Package): string[] {
  const part = optionalXmlPart(pkg, 'xl/sharedStrings.xml');
  if (!part?.xmlDom) return [];
  const strings: string[] = [];
  const siEls = part.xmlDom.getElementsByTagNameNS(SS_NS, 'si');
  for (let i = 0; i < siEls.length; i++) {
    const si = siEls[i];
    if (!si) continue;
    const tEls = si.getElementsByTagNameNS(SS_NS, 't');
    let text = '';
    for (let j = 0; j < tEls.length; j++) {
      text += textOf(tEls[j]);
    }
    strings.push(text);
  }
  return strings;
}
