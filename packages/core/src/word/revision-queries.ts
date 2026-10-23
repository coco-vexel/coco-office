/** WordprocessingML revision discovery shared by the query and mutation paths. */
import type { Package } from '../pack/index.js';
import { W14_NS, W_NS } from './namespaces.js';
import { resolveScope, type StoryId } from './stories.js';
import { WordError } from './word-error.js';

export type RevisionId = string;

export type RevisionKind =
  | 'insert'
  | 'delete'
  | 'moveFrom'
  | 'moveTo'
  | 'runProperty'
  | 'paragraphProperty'
  | 'sectionProperty'
  | 'tableProperty'
  | 'tableRowProperty'
  | 'tableCellProperty'
  | 'tablePropertyException'
  | 'tableGrid'
  | 'numbering'
  | 'cellInsert'
  | 'cellDelete'
  | 'cellMerge'
  | 'customXmlInsert'
  | 'customXmlDelete'
  | 'customXmlMoveFrom'
  | 'customXmlMoveTo'
  | 'conflictInsert'
  | 'conflictDelete'
  | 'customXmlConflictInsert'
  | 'customXmlConflictDelete';

export interface RevisionInfo {
  revId: RevisionId;
  kind: RevisionKind;
  /** Exact WordprocessingML element name, for lossless classification. */
  ooxmlType: SupportedRevisionType;
  author: string;
  date: string;
  story: StoryId;
  paraId?: string;
  text?: string;
}

export interface ListRevisionsOptions {
  /** Defaults to `body`. */
  scope?: 'body' | 'all' | StoryId[];
}

export const SUPPORTED_REVISION_KINDS = {
  ins: 'insert',
  del: 'delete',
  moveFrom: 'moveFrom',
  moveTo: 'moveTo',
  rPrChange: 'runProperty',
  pPrChange: 'paragraphProperty',
  sectPrChange: 'sectionProperty',
  tblPrChange: 'tableProperty',
  trPrChange: 'tableRowProperty',
  tcPrChange: 'tableCellProperty',
  tblPrExChange: 'tablePropertyException',
  tblGridChange: 'tableGrid',
  numberingChange: 'numbering',
  cellIns: 'cellInsert',
  cellDel: 'cellDelete',
  cellMerge: 'cellMerge',
  customXmlInsRangeStart: 'customXmlInsert',
  customXmlDelRangeStart: 'customXmlDelete',
  customXmlMoveFromRangeStart: 'customXmlMoveFrom',
  customXmlMoveToRangeStart: 'customXmlMoveTo',
} as const satisfies Record<string, RevisionKind>;

export type SupportedRevisionLocalName = keyof typeof SUPPORTED_REVISION_KINDS;

export const SUPPORTED_W14_REVISION_KINDS = {
  conflictIns: 'conflictInsert',
  conflictDel: 'conflictDelete',
  customXmlConflictInsRangeStart: 'customXmlConflictInsert',
  customXmlConflictDelRangeStart: 'customXmlConflictDelete',
} as const satisfies Record<string, RevisionKind>;

export type SupportedW14RevisionLocalName = keyof typeof SUPPORTED_W14_REVISION_KINDS;
export type SupportedRevisionType = SupportedRevisionLocalName | SupportedW14RevisionLocalName;

/** Pair/position markers are part of a revision but are not separate SDK revisions. */
export const REVISION_RANGE_MARKER_NAMES = [
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'customXmlInsRangeEnd',
  'customXmlDelRangeEnd',
  'customXmlMoveFromRangeEnd',
  'customXmlMoveToRangeEnd',
] as const;

export const W14_REVISION_RANGE_MARKER_NAMES = [
  'customXmlConflictInsRangeEnd',
  'customXmlConflictDelRangeEnd',
] as const;

export function isSupportedRevisionElement(el: Element): boolean {
  return (
    (el.namespaceURI === W_NS &&
      Object.prototype.hasOwnProperty.call(SUPPORTED_REVISION_KINDS, el.localName)) ||
    (el.namespaceURI === W14_NS &&
      Object.prototype.hasOwnProperty.call(SUPPORTED_W14_REVISION_KINDS, el.localName))
  );
}

export function assertNoUnsupportedRevisions(
  root: Document | Element,
  docIdForError: string,
  opContext: 'listRevisions' | 'acceptRevisions' | 'rejectRevisions',
): void {
  const visit = (node: Node): void => {
    if (node.nodeType === 1) {
      const el = node as Element;
      if (looksLikeRevisionElement(el) && !isKnownRevisionOrMarker(el)) {
        throw new WordError(
          'UNSUPPORTED_OOXML_FEATURE',
          `Revision element '${el.nodeName}' is not supported (op: ${opContext})`,
          docIdForError,
          { revisionKind: el.localName, opName: opContext },
        );
      }
    }
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child) visit(child);
    }
  };
  visit(root);
}

function looksLikeRevisionElement(el: Element): boolean {
  if (el.namespaceURI === W14_NS) {
    return el.localName.startsWith('conflict') || el.localName.startsWith('customXmlConflict');
  }
  if (el.namespaceURI !== W_NS) return false;
  return (
    /(?:Pr|Grid)Change$/.test(el.localName) ||
    el.localName === 'numberingChange' ||
    /^(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge)$/.test(el.localName) ||
    /^(?:customXml(?:Ins|Del|MoveFrom|MoveTo)|move(?:From|To))Range(?:Start|End)$/.test(
      el.localName,
    )
  );
}

function isKnownRevisionOrMarker(el: Element): boolean {
  if (isSupportedRevisionElement(el)) return true;
  if (el.namespaceURI === W_NS) {
    return (REVISION_RANGE_MARKER_NAMES as readonly string[]).includes(el.localName);
  }
  if (el.namespaceURI === W14_NS) {
    return (W14_REVISION_RANGE_MARKER_NAMES as readonly string[]).includes(el.localName);
  }
  return false;
}

export function listRevisions(pkg: Package, opts: ListRevisionsOptions = {}): RevisionInfo[] {
  const stories = resolveScope(pkg, opts.scope ?? 'body');
  const out: RevisionInfo[] = [];
  for (const story of stories) {
    assertNoUnsupportedRevisions(story.root, '', 'listRevisions');
    for (const el of collectSupportedRevisionElements(story.root)) {
      const info = buildRevisionInfo(el, story.id);
      if (info) out.push(info);
    }
  }
  return out;
}

export function collectSupportedRevisionElements(root: Document | Element): Element[] {
  const out: Element[] = [];
  const visit = (node: Node): void => {
    if (node.nodeType === 1 && isSupportedRevisionElement(node as Element))
      out.push(node as Element);
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child) visit(child);
    }
  };
  visit(root);
  return out;
}

function buildRevisionInfo(el: Element, story: StoryId): RevisionInfo | null {
  const revId = el.getAttributeNS(W_NS, 'id');
  if (revId === null) return null;
  const ooxmlType = el.localName as SupportedRevisionType;
  const kind =
    el.namespaceURI === W14_NS
      ? SUPPORTED_W14_REVISION_KINDS[ooxmlType as SupportedW14RevisionLocalName]
      : SUPPORTED_REVISION_KINDS[ooxmlType as SupportedRevisionLocalName];
  const info: RevisionInfo = {
    revId,
    kind,
    ooxmlType,
    author: el.getAttributeNS(W_NS, 'author') ?? '',
    date: el.getAttributeNS(W_NS, 'date') ?? '',
    story,
  };
  const paraId = findContainingParaId(el);
  const text = extractRevisionText(el, ooxmlType);
  if (paraId !== undefined) info.paraId = paraId;
  if (text.length > 0) info.text = text;
  return info;
}

function findContainingParaId(el: Element): string | undefined {
  let cur: Node | null = el.parentNode;
  while (cur && cur.nodeType === 1) {
    const e = cur as Element;
    if (e.namespaceURI === W_NS && e.localName === 'p') {
      const raw = e.getAttributeNS(W14_NS, 'paraId');
      return raw && raw.length > 0 ? raw.toUpperCase() : undefined;
    }
    cur = cur.parentNode;
  }
  return undefined;
}

function extractRevisionText(el: Element, type: SupportedRevisionType): string {
  if (type.includes('RangeStart')) return extractRangeText(el, type);
  const textNames =
    type === 'del' || type === 'moveFrom' || type === 'conflictDel'
      ? ['delText', 't']
      : type === 'ins' || type === 'moveTo' || type === 'conflictIns'
        ? ['t']
        : [];
  let out = '';
  for (const name of textNames) {
    const nodes = el.getElementsByTagNameNS(W_NS, name);
    for (let i = 0; i < nodes.length; i++) out += nodes[i]?.textContent ?? '';
    if (out.length > 0) break;
  }
  return out;
}

function extractRangeText(el: Element, type: SupportedRevisionType): string {
  const endName =
    type === 'customXmlInsRangeStart'
      ? 'customXmlInsRangeEnd'
      : type === 'customXmlDelRangeStart'
        ? 'customXmlDelRangeEnd'
        : type === 'customXmlMoveFromRangeStart'
          ? 'customXmlMoveFromRangeEnd'
          : type === 'customXmlMoveToRangeStart'
            ? 'customXmlMoveToRangeEnd'
            : type === 'customXmlConflictInsRangeStart'
              ? 'customXmlConflictInsRangeEnd'
              : 'customXmlConflictDelRangeEnd';
  const id = el.getAttributeNS(W_NS, 'id');
  const all = flattenNodes(el.ownerDocument!);
  const startIndex = all.indexOf(el);
  const endIndex = all.findIndex(
    (node, index) =>
      index > startIndex &&
      node.nodeType === 1 &&
      (node as Element).namespaceURI === el.namespaceURI &&
      (node as Element).localName === endName &&
      (node as Element).getAttributeNS(W_NS, 'id') === id,
  );
  if (startIndex < 0 || endIndex < 0) return '';
  const preferDeleted =
    type === 'customXmlDelRangeStart' ||
    type === 'customXmlMoveFromRangeStart' ||
    type === 'customXmlConflictDelRangeStart';
  let deleted = '';
  let normal = '';
  for (let i = startIndex + 1; i < endIndex; i++) {
    const node = all[i];
    if (node?.nodeType !== 1 || (node as Element).namespaceURI !== W_NS) continue;
    if ((node as Element).localName === 'delText') deleted += node.textContent ?? '';
    if ((node as Element).localName === 't') normal += node.textContent ?? '';
  }
  return preferDeleted && deleted.length > 0 ? deleted : normal;
}

function flattenNodes(root: Node): Node[] {
  const out: Node[] = [];
  const visit = (node: Node): void => {
    out.push(node);
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child) visit(child);
    }
  };
  visit(root);
  return out;
}
