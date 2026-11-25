import { createXmlPart, deletePart, getPart, markDirty, setPart, type Package, type Part } from '../pack/index.js';
import { R_NS, W_NS } from './namespaces.js';
import { resolveRelTarget } from './plan-executors/placeholder-image.js';
import { findSectionProperties, type HeaderFooterType } from './section-queries.js';
import { listStories } from './stories.js';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELS_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const WORD_PART = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const FOOTER_TYPES = ['default', 'even', 'first'] as const;
const SDK_FOOTER = /^coco-footer-\d+-(?:default|even|first)(?:-\d+)?\.xml$/u;
const AFTER_EVEN_ODD = new Set([
  'bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets',
  'drawingGridHorizontalSpacing', 'displayHorizontalDrawingGridEvery', 'characterSpacingControl',
  'updateFields', 'hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids',
  'mathPr', 'themeFontLang', 'clrSchemeMapping', 'shapeDefaults', 'decimalSymbol', 'listSeparator',
]);

export interface FooterPageNumber {
  font: string;
  fontSizePt: number;
  lineHeightPt?: number;
  alignment?: 'left' | 'center' | 'right';
  insetPt?: number;
  prefix?: string;
  suffix?: string;
}

export interface SectionFooter {
  /** Optional Word paragraphs or tables, imported into the generated footer part. */
  blocks?: readonly Element[];
  pageNumber?: FooterPageNumber;
}

export interface SectionFooters {
  /** Omitted variants retain their current reference. Pass null to remove a reference. */
  footers: Partial<Record<HeaderFooterType, SectionFooter | null>>;
  differentFirstPage?: boolean;
  /** Updates the document-wide evenAndOddHeaders setting for both headers and footers. */
  differentOddEven?: boolean;
}

function word(parent: Element, name: string, attrs: Record<string, string | number> = {}, before?: Element) {
  const element = parent.ownerDocument!.createElementNS(W_NS, `w:${name}`);
  for (const [key, value] of Object.entries(attrs)) element.setAttributeNS(W_NS, `w:${key}`, String(value));
  return parent.insertBefore(element, before ?? null);
}

function children(element: Element) {
  return Array.from(element.children);
}

function child(element: Element, name: string) {
  return children(element).find((node) => node.namespaceURI === W_NS && node.localName === name);
}

function createPart(doc: Document, path: string, namespace: string, root: string, contentType: string) {
  return createXmlPart({ path, contentType, xmlDom: doc.implementation.createDocument(namespace, root, null) });
}

function syncContentType(pkg: Package, path: string, contentType?: string) {
  const types = getPart(pkg, '[Content_Types].xml');
  if (!types?.xmlDom?.documentElement) return;
  const root = types.xmlDom.documentElement;
  const existing = children(root).find((node) => node.localName === 'Override' && node.getAttribute('PartName') === '/' + path);
  if (contentType === undefined) {
    if (!existing) return;
    existing.remove();
  } else {
    const node = existing ?? root.appendChild(types.xmlDom.createElementNS(CONTENT_TYPES_NS, 'Override'));
    node.setAttribute('PartName', '/' + path);
    node.setAttribute('ContentType', contentType);
  }
  markDirty(types);
}

function addPageNumber(root: Element, spec: FooterPageNumber) {
  const p = word(root, 'p');
  const pPr = word(p, 'pPr');
  word(pPr, 'snapToGrid', { val: 0 });
  word(pPr, 'spacing', { before: 0, after: 0, line: Math.round((spec.lineHeightPt ?? spec.fontSizePt + 2) * 20), lineRule: 'exact' });
  const inset = Math.round((spec.insetPt ?? 0) * 20);
  word(pPr, 'ind', { left: inset, right: inset });
  word(pPr, 'jc', { val: spec.alignment ?? 'right' });

  function text(parent: Element, value: string) {
    const run = word(parent, 'r');
    const rPr = word(run, 'rPr');
    word(rPr, 'rFonts', { ascii: spec.font, hAnsi: spec.font, eastAsia: spec.font, cs: spec.font });
    word(rPr, 'color', { val: '000000' });
    const size = Math.round(spec.fontSizePt * 2);
    word(rPr, 'sz', { val: size });
    word(rPr, 'szCs', { val: size });
    const t = word(run, 't');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = value;
  }
  if (spec.prefix) text(p, spec.prefix);
  text(word(p, 'fldSimple', { instr: ' PAGE ' }), '1');
  if (spec.suffix) text(p, spec.suffix);
}

/** Sets a section's footer variants and generates native Word PAGE fields when requested. */
export function setWordSectionFooters(pkg: Package, sectionIndex: number, layout: SectionFooters): void {
  const [body] = listStories(pkg, { kinds: ['body'] });
  if (!body || !Number.isSafeInteger(sectionIndex) || sectionIndex < 0) throw new Error('Invalid Word section');
  const existingSection = findSectionProperties(body.root, sectionIndex);
  if (existingSection === undefined) throw new Error(`Word section ${sectionIndex} was not found`);
  const doc = body.root.ownerDocument!;
  const split = body.partPath.lastIndexOf('/') + 1;
  const folder = body.partPath.slice(0, split);
  const relsPath = `${folder}_rels/${body.partPath.slice(split)}.rels`;
  const relPart = getPart(pkg, relsPath) ?? createPart(doc, relsPath, RELS_NS, 'Relationships', RELS_TYPE);
  const rels = relPart.xmlDom!.documentElement!;
  const ids = new Set(children(rels).map((rel) => rel.getAttribute('Id')));
  const makeRelationship = (type: string, target: string) => {
    let i = 1;
    while (ids.has(`rId${i}`)) i++;
    const id = `rId${i}`;
    ids.add(id);
    const rel = rels.ownerDocument!.createElementNS(RELS_NS, 'Relationship');
    rel.setAttribute('Id', id);
    rel.setAttribute('Type', `${R_NS}/${type}`);
    rel.setAttribute('Target', target);
    return rel;
  };
  const prepared = FOOTER_TYPES.flatMap<{ type: HeaderFooterType; part: Part | null; relationship: Element | null }>((type) => {
    const spec = layout.footers[type];
    if (spec === undefined) return [];
    if (spec === null) return [{ type, part: null, relationship: null }];
    const field = spec.pageNumber;
    if (field && (!field.font.trim() || !Number.isFinite(field.fontSizePt) || field.fontSizePt <= 0 ||
        (field.lineHeightPt !== undefined && (!Number.isFinite(field.lineHeightPt) || field.lineHeightPt <= 0)) ||
        (field.insetPt !== undefined && (!Number.isFinite(field.insetPt) || field.insetPt < 0)))) {
      throw new Error('Invalid Word PAGE field typography');
    }
    let n = 0;
    let target: string;
    do { target = `coco-footer-${sectionIndex}-${type}${n ? '-' + n : ''}.xml`; n++; } while (getPart(pkg, folder + target));
    const part = createPart(doc, folder + target, W_NS, 'w:ftr', `${WORD_PART}.footer+xml`);
    const root = part.xmlDom!.documentElement!;
    for (const block of spec.blocks ?? []) {
      if (block.namespaceURI !== W_NS || !['p', 'tbl'].includes(block.localName)) throw new Error('Invalid Word footer block');
      root.appendChild(root.ownerDocument!.importNode(block, true));
    }
    if (field) addPageNumber(root, field);
    if (!root.firstElementChild) word(root, 'p');
    return [{ type, part, relationship: makeRelationship('footer', target) }];
  });
  const section = existingSection ?? word(body.root, 'sectPr');
  const oldIds = new Set<string>();
  for (const { type, part, relationship } of prepared) {
    for (const node of children(section)) {
      if (node.namespaceURI !== W_NS || node.localName !== 'footerReference' || node.getAttributeNS(W_NS, 'type') !== type) continue;
      const id = node.getAttributeNS(R_NS, 'id');
      if (id) oldIds.add(id);
      node.remove();
    }
    if (!part || !relationship) continue;
    setPart(pkg, part);
    syncContentType(pkg, part.path, part.contentType);
    rels.appendChild(relationship);
    const after = children(section).find((node) => !['headerReference', 'footerReference'].includes(node.localName));
    const ref = word(section, 'footerReference', { type }, after);
    ref.setAttributeNS(R_NS, 'r:id', relationship.getAttribute('Id')!);
  }
  if (layout.differentFirstPage !== undefined) {
    child(section, 'titlePg')?.remove();
    if (layout.differentFirstPage) {
      const following = new Set(['textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange']);
      word(section, 'titlePg', {}, children(section).find((node) => following.has(node.localName)));
    }
  }

  if (layout.differentOddEven !== undefined) {
    const rel = children(rels).find((node) => node.getAttribute('Type') === `${R_NS}/settings`);
    const target = rel?.getAttribute('Target') ?? 'settings.xml';
    const path = resolveRelTarget(relsPath, target);
    const settings = getPart(pkg, path) ?? createPart(doc, path, W_NS, 'w:settings', `${WORD_PART}.settings+xml`);
    const root = settings.xmlDom!.documentElement!;
    child(root, 'evenAndOddHeaders')?.remove();
    if (layout.differentOddEven) word(root, 'evenAndOddHeaders', {}, children(root).find((node) => AFTER_EVEN_ODD.has(node.localName)));
    setPart(pkg, settings);
    syncContentType(pkg, settings.path, settings.contentType);
    markDirty(settings);
    if (!rel) rels.appendChild(makeRelationship('settings', target));
  }

  const referenced = new Set(Array.from(doc.getElementsByTagNameNS(W_NS, 'footerReference'), (ref) => ref.getAttributeNS(R_NS, 'id')));
  for (const rel of children(rels)) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && oldIds.has(id) && !referenced.has(id) && rel.getAttribute('Type') === `${R_NS}/footer` && target && SDK_FOOTER.test(target)) {
      rel.remove();
      if (!children(rels).some((relationship) => relationship.getAttribute('Target') === target)) {
        deletePart(pkg, folder + target);
        syncContentType(pkg, folder + target);
      }
    }
  }
  markDirty(getPart(pkg, body.partPath)!);
  setPart(pkg, relPart);
  markDirty(relPart);
}
