/** Accept/reject semantics for the WordprocessingML revision model. */
import { getPart, isXmlPart, listParts, markDirty } from '../../pack/index.js';
import { W14_NS, W_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import {
  REVISION_RANGE_MARKER_NAMES,
  W14_REVISION_RANGE_MARKER_NAMES,
  assertNoUnsupportedRevisions,
  collectSupportedRevisionElements,
  type SupportedRevisionType,
} from '../revision-queries.js';
import { findStoryFor } from '../stories.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';

type Action = 'accept' | 'reject';
type RevisionScope = Extract<Op, { type: 'acceptRevisions' }>['scope'];

interface RevisionTarget {
  el: Element;
  type: SupportedRevisionType;
}

const PROPERTY_CHANGE_NAMES = new Set<SupportedRevisionType>([
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'trPrChange',
  'tcPrChange',
  'tblPrExChange',
  'tblGridChange',
]);

const CUSTOM_RANGE_END: Partial<Record<SupportedRevisionType, string>> = {
  customXmlInsRangeStart: 'customXmlInsRangeEnd',
  customXmlDelRangeStart: 'customXmlDelRangeEnd',
  customXmlMoveFromRangeStart: 'customXmlMoveFromRangeEnd',
  customXmlMoveToRangeStart: 'customXmlMoveToRangeEnd',
  customXmlConflictInsRangeStart: 'customXmlConflictInsRangeEnd',
  customXmlConflictDelRangeStart: 'customXmlConflictDelRangeEnd',
};

export function executeAcceptRevisions(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'acceptRevisions' }>,
): void {
  processRevisions(ctx, op.scope, 'accept');
}

export function executeRejectRevisions(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'rejectRevisions' }>,
): void {
  processRevisions(ctx, op.scope, 'reject');
}

function processRevisions(ctx: ExecutorContext, scope: RevisionScope, action: Action): void {
  const targets = collectTargets(ctx, scope, action);
  const markedParts = new Set<string>();
  for (const target of targets) {
    const story = findStoryFor(ctx.pkg, target.el);
    if (story) markedParts.add(story.partPath);
  }

  // Descendants first: replacing a property container or deleting a row must not strand
  // independently tracked nested changes that were present in the original document.
  for (const target of [...targets].reverse()) {
    applyTarget(target, action, ctx.docId);
    if (action === 'accept') ctx.result.acceptedRevisions++;
    else ctx.result.rejectedRevisions++;
  }

  if (scope.scope === 'all') removeOrphanRangeMarkers(ctx);
  for (const partPath of markedParts) {
    const part = getPart(ctx.pkg, partPath);
    if (part) markDirty(part);
  }
}

function collectTargets(
  ctx: ExecutorContext,
  scope: RevisionScope,
  action: Action,
): RevisionTarget[] {
  if (scope.scope === 'paragraph') {
    const p = ctx.indices.paraIdx.get(scope.paraId.toUpperCase());
    if (!p) {
      throw new WordError(
        'PARA_NOT_FOUND',
        `paragraph with paraId='${scope.paraId}' not found`,
        ctx.docId,
        {
          paraId: scope.paraId,
          opName: `${action}Revisions`,
        },
      );
    }
    return collectFromRoot(p, ctx.docId, action);
  }

  const all: RevisionTarget[] = [];
  for (const part of listParts(ctx.pkg)) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    all.push(...collectFromRoot(part.xmlDom as Document, ctx.docId, action));
  }
  if (scope.scope === 'all') return all;

  const matched = all.filter((target) => revisionId(target.el) === scope.revId);
  if (matched.length === 0) {
    throw new WordError(
      'REVISION_NOT_FOUND',
      `revision with revId='${scope.revId}' not found`,
      ctx.docId,
      {
        revId: scope.revId,
        opName: `${action}Revisions`,
      },
    );
  }
  return matched;
}

function collectFromRoot(
  root: Document | Element,
  docId: string,
  action: Action,
): RevisionTarget[] {
  assertNoUnsupportedRevisions(
    root,
    docId,
    action === 'accept' ? 'acceptRevisions' : 'rejectRevisions',
  );
  return collectSupportedRevisionElements(root).map((el) => ({
    el,
    type: el.localName as SupportedRevisionType,
  }));
}

function applyTarget(target: RevisionTarget, action: Action, docId: string): void {
  const { el, type } = target;
  if (!el.parentNode) return;

  if (PROPERTY_CHANGE_NAMES.has(type)) {
    applyPropertyChange(el, action);
    return;
  }
  if (type === 'numberingChange') {
    remove(el);
    return;
  }
  if (type === 'cellIns' || type === 'cellDel') {
    applyCellInsertionDeletion(el, type, action);
    return;
  }
  if (type === 'cellMerge') {
    applyCellMerge(el, action);
    return;
  }
  if (CUSTOM_RANGE_END[type]) {
    applyCustomXmlRange(el, type, action, docId);
    return;
  }
  if (type === 'ins' || type === 'del') {
    applyInsertionDeletion(el, type, action);
    return;
  }
  if (type === 'moveFrom' || type === 'moveTo') {
    applyMove(el, type, action);
    return;
  }
  if (type === 'conflictIns' || type === 'conflictDel') {
    applyConflictRevision(el, type, action);
  }
}

function applyPropertyChange(change: Element, action: Action): void {
  const current = change.parentNode;
  if (!current || current.nodeType !== 1) return;
  if (action === 'accept') {
    current.removeChild(change);
    return;
  }

  const currentEl = current as Element;
  const expectedSnapshotName = currentEl.localName;
  const snapshot = directChild(change, expectedSnapshotName);
  const replacement = snapshot
    ? (snapshot.cloneNode(true) as Element)
    : currentEl.ownerDocument!.createElementNS(W_NS, `w:${expectedSnapshotName}`);

  // pPrChange snapshots do not carry the paragraph-mark rPr; it is orthogonal to pPr history.
  if (expectedSnapshotName === 'pPr') {
    const currentMarkRPr = directChild(currentEl, 'rPr');
    if (currentMarkRPr && !directChild(replacement, 'rPr')) {
      replacement.appendChild(currentMarkRPr.cloneNode(true));
    }
  }
  currentEl.parentNode?.replaceChild(replacement, currentEl);
}

function applyInsertionDeletion(el: Element, type: 'ins' | 'del', action: Action): void {
  const row = paragraphOrRowMarkOwner(el, 'tr');
  if (row) {
    const deleteRow =
      (type === 'ins' && action === 'reject') || (type === 'del' && action === 'accept');
    if (deleteRow) row.parentNode?.removeChild(row);
    else remove(el);
    return;
  }

  const paragraph = paragraphOrRowMarkOwner(el, 'p');
  if (paragraph) {
    const merge =
      (type === 'ins' && action === 'reject') || (type === 'del' && action === 'accept');
    if (merge) mergeParagraphWithNext(paragraph);
    else remove(el);
    return;
  }

  const keep = (type === 'ins' && action === 'accept') || (type === 'del' && action === 'reject');
  if (keep) unwrap(el, type === 'del');
  else remove(el);
}

function applyMove(el: Element, type: 'moveFrom' | 'moveTo', action: Action): void {
  const keep =
    (type === 'moveFrom' && action === 'reject') || (type === 'moveTo' && action === 'accept');
  const name = el.getAttributeNS(W_NS, 'name');
  const root = el.ownerDocument;
  if (keep) unwrap(el, type === 'moveFrom');
  else remove(el);
  removeMoveRangeMarkers(root, type, name);
}

function applyConflictRevision(
  el: Element,
  type: 'conflictIns' | 'conflictDel',
  action: Action,
): void {
  const keep =
    (type === 'conflictIns' && action === 'accept') ||
    (type === 'conflictDel' && action === 'reject');
  if (el.childNodes.length > 0) {
    if (keep) unwrap(el, type === 'conflictDel');
    else remove(el);
    return;
  }
  if (keep) {
    remove(el);
    return;
  }
  const owner =
    paragraphOrRowMarkOwner(el, 'p') ??
    paragraphOrRowMarkOwner(el, 'tr') ??
    nearestAncestor(el, 'r') ??
    nearestAncestor(el, 'tc') ??
    nearestAncestor(el, 'tr') ??
    nearestAncestor(el, 'p');
  if (owner) owner.parentNode?.removeChild(owner);
  else remove(el);
}

function applyCellInsertionDeletion(
  marker: Element,
  type: 'cellIns' | 'cellDel',
  action: Action,
): void {
  const cell = nearestAncestor(marker, 'tc');
  const deleteCell =
    (type === 'cellIns' && action === 'reject') || (type === 'cellDel' && action === 'accept');
  if (deleteCell && cell) cell.parentNode?.removeChild(cell);
  else remove(marker);
}

function applyCellMerge(marker: Element, action: Action): void {
  if (action === 'reject') {
    remove(marker);
    return;
  }
  const parent = marker.parentNode;
  if (!parent) return;
  const raw = marker.getAttributeNS(W_NS, 'vMerge');
  if (raw !== 'rest' && raw !== 'cont') {
    remove(marker);
    return;
  }
  const vMerge = marker.ownerDocument!.createElementNS(W_NS, 'w:vMerge');
  vMerge.setAttributeNS(W_NS, 'w:val', raw === 'rest' ? 'restart' : 'continue');
  parent.replaceChild(vMerge, marker);
}

function applyCustomXmlRange(
  start: Element,
  type: SupportedRevisionType,
  action: Action,
  docId: string,
): void {
  const endName = CUSTOM_RANGE_END[type]!;
  const id = revisionId(start);
  const end = findRangeEnd(start.ownerDocument!, start.namespaceURI!, endName, id);
  if (!end) {
    throw new WordError('UNSUPPORTED_OOXML_FEATURE', `Unpaired <w:${type}> revision range`, docId, {
      revisionKind: type,
      revId: id,
    });
  }
  const keep =
    (type === 'customXmlInsRangeStart' && action === 'accept') ||
    (type === 'customXmlDelRangeStart' && action === 'reject') ||
    (type === 'customXmlMoveFromRangeStart' && action === 'reject') ||
    (type === 'customXmlMoveToRangeStart' && action === 'accept') ||
    (type === 'customXmlConflictInsRangeStart' && action === 'accept') ||
    (type === 'customXmlConflictDelRangeStart' && action === 'reject');
  if (keep) {
    if (
      type === 'customXmlDelRangeStart' ||
      type === 'customXmlMoveFromRangeStart' ||
      type === 'customXmlConflictDelRangeStart'
    ) {
      for (const node of nodesBetween(start, end)) {
        if (node.nodeType === 1) convertDelTextToT(node as Element);
      }
    }
    remove(start);
    remove(end);
    return;
  }

  deleteRangeContents(start, end);
  remove(start);
  remove(end);
}

function paragraphOrRowMarkOwner(el: Element, ownerName: 'p' | 'tr'): Element | null {
  let parent = el.parentNode;
  if (!parent || parent.nodeType !== 1) return null;
  if (
    ownerName === 'p' &&
    (parent as Element).namespaceURI === W_NS &&
    (parent as Element).localName === 'rPr'
  ) {
    parent = parent.parentNode;
  }
  const propertyName = ownerName === 'p' ? 'pPr' : 'trPr';
  if (
    !parent ||
    parent.nodeType !== 1 ||
    (parent as Element).namespaceURI !== W_NS ||
    (parent as Element).localName !== propertyName
  ) {
    return null;
  }
  const owner = parent.parentNode;
  return owner &&
    owner.nodeType === 1 &&
    (owner as Element).namespaceURI === W_NS &&
    (owner as Element).localName === ownerName
    ? (owner as Element)
    : null;
}

function mergeParagraphWithNext(paragraph: Element): void {
  removeRevisionMarksFromParagraphProperties(paragraph);
  let next = paragraph.nextSibling;
  while (next && (next.nodeType !== 1 || (next as Element).localName !== 'p'))
    next = next.nextSibling;
  if (!next || (next as Element).namespaceURI !== W_NS) return;
  const nextParagraph = next as Element;
  for (const child of Array.from(nextParagraph.childNodes)) {
    if (
      child.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName === 'pPr'
    ) {
      continue;
    }
    paragraph.appendChild(child);
  }
  nextParagraph.parentNode?.removeChild(nextParagraph);
}

function removeRevisionMarksFromParagraphProperties(paragraph: Element): void {
  for (const name of ['ins', 'del']) {
    const nodes = Array.from(paragraph.getElementsByTagNameNS(W_NS, name));
    for (const node of nodes) {
      if (paragraphOrRowMarkOwner(node, 'p') === paragraph) remove(node);
    }
  }
}

function unwrap(el: Element, restoreDeletedText: boolean): void {
  const parent = el.parentNode;
  if (!parent) return;
  const children = Array.from(el.childNodes);
  if (restoreDeletedText) {
    for (const child of children) if (child.nodeType === 1) convertDelTextToT(child as Element);
  }
  for (const child of children) parent.insertBefore(child, el);
  parent.removeChild(el);
}

function convertDelTextToT(root: Element): void {
  const nodes = Array.from(root.getElementsByTagNameNS(W_NS, 'delText'));
  if (root.namespaceURI === W_NS && root.localName === 'delText') nodes.unshift(root);
  for (const delText of nodes) {
    if (!delText.parentNode) continue;
    const t = delText.ownerDocument!.createElementNS(W_NS, 'w:t');
    for (let i = 0; i < delText.attributes.length; i++) {
      const attr = delText.attributes[i];
      if (!attr) continue;
      if (attr.namespaceURI) t.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
      else t.setAttribute(attr.name, attr.value);
    }
    t.textContent = delText.textContent;
    delText.parentNode.replaceChild(t, delText);
  }
}

function removeMoveRangeMarkers(
  root: Document,
  type: 'moveFrom' | 'moveTo',
  name: string | null,
): void {
  if (name === null) return;
  const names =
    type === 'moveFrom'
      ? ['moveFromRangeStart', 'moveFromRangeEnd']
      : ['moveToRangeStart', 'moveToRangeEnd'];
  for (const local of names) {
    const markers = Array.from(root.getElementsByTagNameNS(W_NS, local));
    for (const marker of markers) {
      if (marker.getAttributeNS(W_NS, 'name') === name) remove(marker);
    }
  }
}

function removeOrphanRangeMarkers(ctx: ExecutorContext): void {
  for (const part of listParts(ctx.pkg)) {
    if (!isXmlPart(part) || !part.xmlDom) continue;
    for (const local of REVISION_RANGE_MARKER_NAMES) {
      for (const marker of Array.from(part.xmlDom.getElementsByTagNameNS(W_NS, local)))
        remove(marker);
    }
    for (const local of W14_REVISION_RANGE_MARKER_NAMES) {
      for (const marker of Array.from(part.xmlDom.getElementsByTagNameNS(W14_NS, local)))
        remove(marker);
    }
  }
}

function findRangeEnd(
  root: Document,
  namespace: string,
  localName: string,
  id: string | null,
): Element | null {
  const ends = root.getElementsByTagNameNS(namespace, localName);
  for (let i = 0; i < ends.length; i++) {
    const end = ends[i];
    if (end && revisionId(end) === id) return end;
  }
  return null;
}

function nodesBetween(start: Node, end: Node): Node[] {
  const all = flattenNodes(start.ownerDocument!);
  const from = all.indexOf(start);
  const to = all.indexOf(end);
  return from >= 0 && to > from ? all.slice(from + 1, to) : [];
}

function deleteRangeContents(start: Node, end: Node): void {
  const candidates = new Set(nodesBetween(start, end));
  for (const node of [...candidates]) {
    if (isAncestor(node, end)) candidates.delete(node);
  }
  for (const node of candidates) {
    if (node.parentNode && !candidates.has(node.parentNode)) node.parentNode.removeChild(node);
  }
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

function isAncestor(candidate: Node, node: Node): boolean {
  let current: Node | null = node.parentNode;
  while (current) {
    if (current === candidate) return true;
    current = current.parentNode;
  }
  return false;
}

function revisionId(el: Element): string | null {
  return el.getAttributeNS(W_NS, 'id') ?? el.getAttribute('w:id');
}

function directChild(parent: Element, localName: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName === localName
    ) {
      return child as Element;
    }
  }
  return null;
}

function nearestAncestor(el: Element, localName: string): Element | null {
  let current: Node | null = el.parentNode;
  while (current && current.nodeType === 1) {
    const ancestor = current as Element;
    if (ancestor.namespaceURI === W_NS && ancestor.localName === localName) return ancestor;
    current = ancestor.parentNode;
  }
  return null;
}

function remove(el: Element): void {
  el.parentNode?.removeChild(el);
}
