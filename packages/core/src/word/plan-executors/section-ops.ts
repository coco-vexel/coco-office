import { markDirty } from '../../pack/index.js';
import { W_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { findSectionProperties } from '../section-queries.js';
import { listStories } from '../stories.js';
import { WordError } from '../word-error.js';
import { createW, firstDirectChildW } from './_shared.js';
import type { ExecutorContext } from './executor-context.js';

// Match the page size and margins used by WordDocument.createBlank().
const DEFAULT_PAGE_WIDTH = 12240;
const DEFAULT_PAGE_HEIGHT = 15840;
const DEFAULT_MARGIN = 1440;
const DEFAULT_HEADER_FOOTER_DISTANCE = 720;
const TWIPS_PER_POINT = 20;
const MARGIN_SIDES = ['top', 'bottom', 'left', 'right'] as const;
const BEFORE_PAGE_SIZE = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type'];
const AFTER_VERTICAL_ALIGNMENT = [
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange',
];

export function executeSetSectionPageLayout(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'setSectionPageLayout' }>,
): void {
  const { sectionIndex, layout } = op;
  const details = { opType: op.type, sectionIndex };
  const invalid = (message: string) =>
    new WordError('INVALID_ARGUMENT', message, ctx.docId, details);

  if (ctx.options.trackChanges) {
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'setSectionPageLayout does not support tracked changes',
      ctx.docId,
      details,
    );
  }
  if (!Number.isSafeInteger(sectionIndex) || sectionIndex < 0) {
    throw invalid('sectionIndex must be a nonnegative integer');
  }
  if (!layout || typeof layout !== 'object' || Array.isArray(layout)) {
    throw invalid('layout must be an object');
  }
  const { orientation, margins, pageSize: requestedSize, verticalAlignment } = layout;
  if (orientation !== undefined && orientation !== 'portrait' && orientation !== 'landscape') {
    throw invalid('orientation must be portrait or landscape');
  }
  if (
    requestedSize !== undefined &&
    (!requestedSize ||
      ![requestedSize.width, requestedSize.height].every(
        (value) =>
          Number.isFinite(value) &&
          value >= 1 / TWIPS_PER_POINT &&
          value <= 4_294_967_295 / TWIPS_PER_POINT,
      ))
  ) {
    throw invalid('pageSize must contain positive width and height in points');
  }
  if (
    verticalAlignment !== undefined &&
    !['top', 'center', 'both', 'bottom'].includes(verticalAlignment)
  ) {
    throw invalid('verticalAlignment must be top, center, both or bottom');
  }
  if (
    margins !== undefined &&
    (!margins || typeof margins !== 'object' || Array.isArray(margins))
  ) {
    throw invalid('margins must be an object');
  }
  const marginChanges: Array<[side: (typeof MARGIN_SIDES)[number], twips: number]> = [];
  for (const side of MARGIN_SIDES) {
    const points = margins?.[side];
    if (points === undefined) continue;
    const signed = side === 'top' || side === 'bottom';
    const min = (signed ? -2_147_483_648 : 0) / TWIPS_PER_POINT;
    const max = (signed ? 2_147_483_647 : 4_294_967_295) / TWIPS_PER_POINT;
    if (!Number.isFinite(points) || points < min || points > max) {
      throw invalid(`${side} margin must be a finite number in points within the supported range`);
    }
    marginChanges.push([side, Math.round(points * TWIPS_PER_POINT)]);
  }

  const [body] = listStories(ctx.pkg, { kinds: ['body'] });
  if (!body) throw invalid('The document has no body section');
  const existing = findSectionProperties(body.root, sectionIndex);
  if (existing === undefined) throw invalid(`Section ${sectionIndex} was not found`);
  if (
    orientation === undefined &&
    !requestedSize &&
    marginChanges.length === 0 &&
    verticalAlignment === undefined
  )
    return;

  const section = existing ?? body.root.appendChild(createW(body.root.ownerDocument!, 'sectPr'));
  if (orientation !== undefined || requestedSize) {
    const pageSize =
      firstDirectChildW(section, 'pgSz') ?? insertProperty(section, 'pgSz', BEFORE_PAGE_SIZE);
    const currentWidth = Number(pageSize.getAttributeNS(W_NS, 'w'));
    const currentHeight = Number(pageSize.getAttributeNS(W_NS, 'h'));
    const hasSize =
      Number.isSafeInteger(currentWidth) &&
      Number.isSafeInteger(currentHeight) &&
      currentWidth > 0 &&
      currentHeight > 0;
    const width = requestedSize
      ? Math.round(requestedSize.width * TWIPS_PER_POINT)
      : hasSize
        ? currentWidth
        : DEFAULT_PAGE_WIDTH;
    const height = requestedSize
      ? Math.round(requestedSize.height * TWIPS_PER_POINT)
      : hasSize
        ? currentHeight
        : DEFAULT_PAGE_HEIGHT;
    const shortSide = Math.min(width, height);
    const longSide = Math.max(width, height);
    setTwips(
      pageSize,
      'w',
      orientation === 'landscape' ? longSide : orientation === 'portrait' ? shortSide : width,
    );
    setTwips(
      pageSize,
      'h',
      orientation === 'landscape' ? shortSide : orientation === 'portrait' ? longSide : height,
    );
    const direction = orientation ?? (width > height ? 'landscape' : 'portrait');
    pageSize.setAttributeNS(W_NS, 'w:orient', direction);
    if (requestedSize) pageSize.removeAttributeNS(W_NS, 'code');
  }
  if (marginChanges.length > 0) {
    let pageMargins = firstDirectChildW(section, 'pgMar');
    if (!pageMargins) {
      pageMargins = insertProperty(section, 'pgMar', [...BEFORE_PAGE_SIZE, 'pgSz']);
      for (const side of MARGIN_SIDES) setTwips(pageMargins, side, DEFAULT_MARGIN);
      setTwips(pageMargins, 'header', DEFAULT_HEADER_FOOTER_DISTANCE);
      setTwips(pageMargins, 'footer', DEFAULT_HEADER_FOOTER_DISTANCE);
      setTwips(pageMargins, 'gutter', 0);
    }
    for (const [side, twips] of marginChanges) setTwips(pageMargins, side, twips);
  }
  if (verticalAlignment !== undefined) {
    let alignment = firstDirectChildW(section, 'vAlign');
    if (!alignment) {
      const following = Array.from(section.children).find(
        (child) =>
          child.namespaceURI === W_NS && AFTER_VERTICAL_ALIGNMENT.includes(child.localName),
      );
      alignment = section.insertBefore(
        createW(section.ownerDocument!, 'vAlign'),
        following ?? null,
      );
    }
    alignment.setAttributeNS(W_NS, 'w:val', verticalAlignment);
  }
  markDirty(ctx.pkg.parts.get(body.partPath)!);
}

function insertProperty(section: Element, name: string, predecessors: readonly string[]): Element {
  const following = Array.from(section.children).find(
    (child) => child.namespaceURI === W_NS && !predecessors.includes(child.localName),
  );
  return section.insertBefore(createW(section.ownerDocument!, name), following ?? null);
}

function setTwips(element: Element, name: string, value: number): void {
  element.setAttributeNS(W_NS, `w:${name}`, String(value));
}
