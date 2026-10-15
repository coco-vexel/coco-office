import { fromFlatOpc, fromZip } from '@coco-vexel/coco-office/pack';
import type { Package } from '@coco-vexel/coco-office/pack';
import {
  buildIndices,
  buildParagraphInfo,
  getImageBytesSync,
  listDrawings,
  listTables,
  listSections,
  listStories,
  W14_NS,
  W_NS,
  type CommentInfo,
  type DrawingInfo,
  type ImageBytesResult,
  type PackageIndices,
  type ParagraphFormat,
  type RevisionInfo,
  type TableInfo,
} from '@coco-vexel/coco-office/word-view';
import type {
  EditorBlockView,
  EditorCommentSpan,
  EditorContentSegment,
  EditorDrawingSegment,
  EditorDocumentDefaults,
  EditorDocumentView,
  EditorBoxEdges,
  EditorFieldSegment,
  EditorImageSegment,
  EditorNumberingFormat,
  EditorPageLayout,
  EditorParagraphFormat,
  EditorParagraphSpan,
  EditorParagraphView,
  EditorPlaceholderSegment,
  EditorRevisionSegment,
  EditorRunFormat,
  EditorSectionView,
  EditorHeaderFooterStoryView,
  EditorHyperlink,
  EditorTableCellView,
  EditorTableBorders,
  EditorTableLook,
  EditorTableRowView,
  EditorTableView,
  EditorTextSegment,
} from './model.js';

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DEFAULT_TABLE_LOOK: Required<Pick<
  EditorTableLook,
  'firstRow' | 'lastRow' | 'firstColumn' | 'lastColumn' | 'noHorizontalBand' | 'noVerticalBand'
>> = {
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  noHorizontalBand: false,
  noVerticalBand: true,
};

export interface WordDocumentLike {
  _getPackage(): Package;
  _getIndices(): PackageIndices;
}

/** Build the editor view directly from browser-friendly DOCX package bytes. */
export async function buildWordEditorViewFromDocx(
  bytes: Uint8Array | ArrayBuffer,
): Promise<EditorDocumentView> {
  const pkg = await fromZip(bytes);
  return buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
}

/** Build the editor view directly from a browser/Electron Flat OPC payload. */
export function buildWordEditorViewFromFlatOpc(xml: string): EditorDocumentView {
  const pkg = fromFlatOpc(xml);
  return buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
}

export async function buildWordEditorView(doc: WordDocumentLike): Promise<EditorDocumentView> {
  return buildWordEditorViewFromPackage(doc._getPackage(), doc._getIndices());
}

export function buildWordEditorViewFromPackage(
  pkg: Package,
  indices: PackageIndices,
): EditorDocumentView {
  const blocks: EditorBlockView[] = [];
  const paragraphs: EditorParagraphView[] = [];
  const tables: EditorTableView[] = [];
  const attachedCommentIds = new Set<number>();
  const tableInfoById = new Map(
    listTables(pkg, indices, { scope: 'all', includeNested: true, includeCellTexts: true }).map((table) => [
      table.tableId,
      table,
    ]),
  );
  const tableIdByElement = new Map<Element, string>();
  for (const [tableId, tableEl] of indices.tblIdx) {
    tableIdByElement.set(tableEl, tableId);
  }
  const drawingIdByElement = new Map<Element, string>();
  const drawingInfoById = new Map(
    listDrawings(pkg, indices, { scope: 'all' }).map((drawing) => [drawing.drawingId, drawing]),
  );
  const drawingsByAnchorParaId = new Map<string, DrawingInfo[]>();
  for (const drawing of drawingInfoById.values()) {
    if (!drawing.anchorParaId) continue;
    const drawings = drawingsByAnchorParaId.get(drawing.anchorParaId) ?? [];
    drawings.push(drawing);
    drawingsByAnchorParaId.set(drawing.anchorParaId, drawings);
  }
  const imagesByDrawingId = new Map<string, EditorImageData>();
  for (const [drawingId, drawingEl] of indices.drawingIdx) {
    drawingIdByElement.set(drawingEl, drawingId);
    const image = buildEditorImage(pkg, indices, drawingId, drawingInfoById.get(drawingId));
    if (image) imagesByDrawingId.set(drawingId, image);
  }
  const theme = extractTheme(pkg);
  const context: ExtractContext = {
    indices,
    drawingIdByElement,
    drawingInfoById,
    drawingsByAnchorParaId,
    imagesByDrawingId,
    theme,
    documentDefaults: extractDocumentDefaults(pkg, theme),
    numbering: buildNumberingDefinitions(pkg),
    numberingCounters: new Map(),
    defaultParagraphStyleId: findDefaultParagraphStyleId(pkg),
    paragraphStyleFormats: buildParagraphStyleFormats(pkg, indices, theme),
    tableStyleFormats: buildTableStyleFormats(pkg, theme),
    hyperlinkByElement: buildHyperlinkMap(pkg),
  };

  for (const story of listStories(pkg, { kinds: ['body'] })) {
    const storyView = buildStoryBlocks(story.id, story.root, {
      context,
      attachedCommentIds,
      tableIdByElement,
      tableInfoById,
    });
    blocks.push(...storyView.blocks);
    paragraphs.push(...storyView.paragraphs);
    tables.push(...storyView.tables);
  }

  const headerFooterStories = extractHeaderFooterStories(pkg, {
    context,
    attachedCommentIds,
    tableIdByElement,
    tableInfoById,
  });

  return {
    kind: 'word',
    defaults: context.documentDefaults,
    pageLayout: extractPageLayout(pkg),
    sections: extractSectionsWithFallback(pkg),
    headers: headerFooterStories.headers,
    footers: headerFooterStories.footers,
    blocks,
    paragraphs,
    tables,
    detachedComments: buildDetachedComments(indices, attachedCommentIds),
    detachedRevisions: [],
  };
}

function extractPageLayout(pkg: Package): EditorPageLayout | undefined {
  const documentPart = pkg.parts.get('word/document.xml');
  const documentDom = documentPart?.xmlDom;
  if (!documentDom) return undefined;
  const body = documentDom.getElementsByTagNameNS(W_NS, 'body')[0];
  if (!body) return undefined;

  const sectPr = findLastBodySectPr(body);
  if (!sectPr) return undefined;
  return parsePageLayoutFromSectPr(sectPr);
}

function parsePageLayoutFromSectPr(sectPr: Element): EditorPageLayout | undefined {
  const pgSz = firstDirectChild(sectPr, W_NS, 'pgSz');
  const pgMar = firstDirectChild(sectPr, W_NS, 'pgMar');
  const docGrid = lastDirectChild(sectPr, W_NS, 'docGrid');
  const width = pgSz ? numAttr(pgSz, 'w') : undefined;
  const height = pgSz ? numAttr(pgSz, 'h') : undefined;
  if (width === undefined || height === undefined) return undefined;
  const orientation = pgSz ? pgSz.getAttributeNS(W_NS, 'orient') : null;

  const layout: EditorPageLayout = {
    width,
    height,
    margin: {
      top: pgMar ? numAttr(pgMar, 'top') ?? 1440 : 1440,
      right: pgMar ? numAttr(pgMar, 'right') ?? 1440 : 1440,
      bottom: pgMar ? numAttr(pgMar, 'bottom') ?? 1440 : 1440,
      left: pgMar ? numAttr(pgMar, 'left') ?? 1440 : 1440,
    },
  };
  const headerDistance = pgMar ? numAttr(pgMar, 'header') : undefined;
  const footerDistance = pgMar ? numAttr(pgMar, 'footer') : undefined;
  if (headerDistance !== undefined) layout.headerDistance = headerDistance;
  if (footerDistance !== undefined) layout.footerDistance = footerDistance;
  const parsedDocGrid = docGrid ? parseDocGrid(docGrid) : undefined;
  if (parsedDocGrid) layout.docGrid = parsedDocGrid;
  if (orientation === 'landscape') layout.orientation = 'landscape';
  else if (orientation === 'portrait') layout.orientation = 'portrait';
  return layout;
}

function parseDocGrid(docGrid: Element): NonNullable<EditorPageLayout['docGrid']> | undefined {
  const typeRaw = docGrid.getAttributeNS(W_NS, 'type');
  const type =
    typeRaw === 'default' || typeRaw === 'lines' || typeRaw === 'linesAndChars' || typeRaw === 'snapToChars'
      ? typeRaw
      : undefined;
  const linePitch = numAttr(docGrid, 'linePitch');
  const charSpace = numAttr(docGrid, 'charSpace');
  if (type === undefined && linePitch === undefined && charSpace === undefined) return undefined;
  return {
    ...(type !== undefined ? { type } : {}),
    ...(linePitch !== undefined ? { linePitch } : {}),
    ...(charSpace !== undefined ? { charSpace } : {}),
  };
}

function findLastBodySectPr(body: Element): Element | null {
  for (let i = body.childNodes.length - 1; i >= 0; i--) {
    const child = body.childNodes[i];
    if (!child || child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI !== W_NS) continue;
    if (el.localName === 'sectPr') return el;
    if (el.localName === 'p') {
      const pPr = firstDirectChild(el, W_NS, 'pPr');
      const sectPr = pPr ? firstDirectChild(pPr, W_NS, 'sectPr') : null;
      if (sectPr) return sectPr;
    }
  }
  return null;
}

function extractTheme(pkg: Package): EditorTheme {
  const theme: EditorTheme = { colors: new Map(), fonts: new Map() };
  const themeDom = pkg.parts.get('word/theme/theme1.xml')?.xmlDom;
  if (!themeDom) return theme;

  const clrScheme = themeDom.getElementsByTagNameNS(A_NS, 'clrScheme')[0];
  if (clrScheme) {
    for (const slot of [
      'dk1',
      'lt1',
      'dk2',
      'lt2',
      'accent1',
      'accent2',
      'accent3',
      'accent4',
      'accent5',
      'accent6',
      'hlink',
      'folHlink',
    ]) {
      const color = themeColorSlot(clrScheme, slot);
      if (color) theme.colors.set(slot, color);
    }
  }

  const fontScheme = themeDom.getElementsByTagNameNS(A_NS, 'fontScheme')[0];
  if (fontScheme) {
    applyThemeFonts(theme.fonts, firstDirectChild(fontScheme, A_NS, 'majorFont'), 'major');
    applyThemeFonts(theme.fonts, firstDirectChild(fontScheme, A_NS, 'minorFont'), 'minor');
  }
  return theme;
}

function themeColorSlot(clrScheme: Element, slot: string): string | undefined {
  const slotEl = firstDirectChild(clrScheme, A_NS, slot);
  if (!slotEl) return undefined;
  for (let i = 0; i < slotEl.childNodes.length; i++) {
    const child = slotEl.childNodes[i];
    if (!child || child.nodeType !== 1 || (child as Element).namespaceURI !== A_NS) continue;
    const colorEl = child as Element;
    const raw = colorEl.getAttribute('val') ?? colorEl.getAttribute('lastClr');
    const color = normalizeRgbColor(raw);
    if (color) return color;
  }
  return undefined;
}

function applyThemeFonts(fonts: Map<string, string>, fontEl: Element | null, prefix: 'major' | 'minor'): void {
  if (!fontEl) return;
  const latin = themeTypeface(firstDirectChild(fontEl, A_NS, 'latin'));
  const eastAsia = themeTypeface(firstDirectChild(fontEl, A_NS, 'ea')) || latin;
  const complexScript = themeTypeface(firstDirectChild(fontEl, A_NS, 'cs')) || latin;
  if (latin) {
    fonts.set(`${prefix}Ascii`, latin);
    fonts.set(`${prefix}HAnsi`, latin);
  }
  if (eastAsia) fonts.set(`${prefix}EastAsia`, eastAsia);
  if (complexScript) fonts.set(`${prefix}Bidi`, complexScript);
}

function themeTypeface(el: Element | null): string | undefined {
  const typeface = el?.getAttribute('typeface') ?? undefined;
  return typeface && typeface.length > 0 ? typeface : undefined;
}

function extractSectionsWithFallback(pkg: Package): EditorSectionView[] {
  const bodySections = extractSectionsFromBody(pkg);
  if (bodySections.length > 0) return bodySections;
  const sections = listSections(pkg);
  if (sections.length > 0) {
    return sections.map((section) => ({
      sectionIndex: section.sectionIndex,
      firstParaId: section.firstParaId,
      lastParaId: section.lastParaId,
      ...(section.type ? { type: section.type } : {}),
      headerRefs: section.headerRefs ?? [],
      footerRefs: section.footerRefs ?? [],
      titlePage: section.titlePage === true,
    }));
  }
  return [];
}

function extractSectionsFromBody(pkg: Package): EditorSectionView[] {
  const story = listStories(pkg, { kinds: ['body'] })[0];
  if (!story) return [];
  const relsPath = relsPathFor(story.partPath);
  const sections: EditorSectionView[] = [];
  let firstParaId: string | undefined;
  let lastParaId: string | undefined;
  let blockIndex = 0;

  for (let i = 0; i < story.root.childNodes.length; i++) {
    const child = story.root.childNodes[i];
    if (!child || child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI !== W_NS) continue;

    if (el.localName === 'p') {
      const paraId = paragraphIdWithFallback('body', el, blockIndex++);
      if (!firstParaId) firstParaId = paraId;
      lastParaId = paraId;
      const inlineSectPr = firstDirectDescendant(el, W_NS, ['pPr', 'sectPr']);
      if (inlineSectPr && firstParaId && lastParaId) {
        sections.push(buildFallbackSection(sections.length, firstParaId, lastParaId, inlineSectPr, pkg, relsPath));
        firstParaId = undefined;
      }
      continue;
    }

    if (el.localName === 'tbl') {
      blockIndex++;
      continue;
    }

    if (el.localName === 'sectPr' && firstParaId && lastParaId) {
      sections.push(buildFallbackSection(sections.length, firstParaId, lastParaId, el, pkg, relsPath));
      firstParaId = undefined;
    }
  }

  return sections;
}

function buildFallbackSection(
  sectionIndex: number,
  firstParaId: string,
  lastParaId: string,
  sectPr: Element,
  pkg: Package,
  relsPath: string | undefined,
): EditorSectionView {
  const section: EditorSectionView = {
    sectionIndex,
    firstParaId,
    lastParaId,
    headerRefs: collectHeaderFooterRefs(pkg, sectPr, relsPath, 'headerReference'),
    footerRefs: collectHeaderFooterRefs(pkg, sectPr, relsPath, 'footerReference'),
    titlePage: firstDirectChild(sectPr, W_NS, 'titlePg') !== null,
  };
  const pageLayout = parsePageLayoutFromSectPr(sectPr);
  if (pageLayout) section.pageLayout = pageLayout;
  const pgNumType = firstDirectChild(sectPr, W_NS, 'pgNumType');
  const pageNumberStart = pgNumType ? numAttr(pgNumType, 'start') : undefined;
  if (pageNumberStart !== undefined) section.pageNumberStart = pageNumberStart;
  const type = firstDirectChild(sectPr, W_NS, 'type')?.getAttributeNS(W_NS, 'val');
  if (
    type === 'nextPage' ||
    type === 'continuous' ||
    type === 'evenPage' ||
    type === 'oddPage' ||
    type === 'nextColumn'
  ) {
    section.type = type;
  }
  return section;
}

function collectHeaderFooterRefs(
  pkg: Package,
  sectPr: Element,
  relsPath: string | undefined,
  localName: 'headerReference' | 'footerReference',
): EditorSectionView['headerRefs'] {
  if (!relsPath) return [];
  const rels = loadRelationships(pkg, relsPath);
  if (!rels) return [];
  const refs: EditorSectionView['headerRefs'] = [];
  for (const el of directChildren(sectPr, W_NS, localName)) {
    const rId = el.getAttributeNS(R_NS, 'id');
    if (!rId) continue;
    const target = rels.get(rId);
    if (!target) continue;
    const typeAttr = el.getAttributeNS(W_NS, 'type');
    refs.push({
      type: typeAttr === 'first' || typeAttr === 'even' ? typeAttr : 'default',
      partName: resolveRelationshipTarget(relsPath, target),
    });
  }
  return refs;
}

function loadRelationships(pkg: Package, relsPath: string): Map<string, string> | null {
  const part = pkg.parts.get(relsPath);
  const relsDom = part?.xmlDom;
  if (!relsDom) return null;
  const rels = new Map<string, string>();
  const nodes = relsDom.getElementsByTagNameNS(RELS_NS, 'Relationship');
  for (let i = 0; i < nodes.length; i++) {
    const rel = nodes[i];
    const id = rel?.getAttribute('Id');
    const target = rel?.getAttribute('Target');
    if (id && target) rels.set(id, target);
  }
  return rels;
}

function resolveRelationshipTarget(relsPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const relsMarker = '/_rels/';
  const markerIndex = relsPath.lastIndexOf(relsMarker);
  if (markerIndex < 0) return target;
  return `${relsPath.slice(0, markerIndex)}/${target}`.replace(/\\/g, '/');
}

function relsPathFor(partPath: string): string | undefined {
  const lastSlash = partPath.lastIndexOf('/');
  const dir = lastSlash >= 0 ? partPath.slice(0, lastSlash) : '';
  const file = partPath.slice(lastSlash + 1);
  return dir.length > 0 ? `${dir}/_rels/${file}.rels` : `_rels/${file}.rels`;
}

function buildHyperlinkMap(pkg: Package): Map<Element, EditorHyperlink> {
  const result = new Map<Element, EditorHyperlink>();
  for (const story of listStories(pkg, { kinds: ['body', 'header', 'footer'] })) {
    const relsPath = relsPathFor(story.partPath);
    const relationships = relsPath ? loadRelationships(pkg, relsPath) : null;
    const hyperlinks = story.root.getElementsByTagNameNS(W_NS, 'hyperlink');
    for (let index = 0; index < hyperlinks.length; index += 1) {
      const element = hyperlinks[index];
      if (!element) continue;
      const relationshipId = element.getAttributeNS(R_NS, 'id');
      const target = relationshipId ? relationships?.get(relationshipId) : undefined;
      const anchor = element.getAttributeNS(W_NS, 'anchor') ?? undefined;
      const tooltip = element.getAttributeNS(W_NS, 'tooltip') ?? undefined;
      if (!target && !anchor) continue;
      result.set(element, {
        ...(target ? { target } : {}),
        ...(anchor ? { anchor } : {}),
        ...(tooltip ? { tooltip } : {}),
      });
    }
  }
  return result;
}

function paragraphIdWithFallback(storyId: EditorParagraphView['story'], paragraph: Element, index: number): string {
  return paragraph.getAttributeNS(W14_NS, 'paraId')?.toUpperCase() || `${storyId}:p:${index}`;
}

interface BuildStoryBlocksOptions {
  context: ExtractContext;
  attachedCommentIds: Set<number>;
  tableIdByElement: Map<Element, string>;
  tableInfoById: Map<string, TableInfo>;
  flattenSdt?: boolean;
  suppressPlaceholderSdtContent?: boolean;
}

function buildStoryBlocks(
  storyId: EditorParagraphView['story'],
  root: Element,
  options: BuildStoryBlocksOptions,
): Pick<EditorHeaderFooterStoryView, 'blocks' | 'paragraphs' | 'tables'> {
  const blocks: EditorBlockView[] = [];
  const paragraphs: EditorParagraphView[] = [];
  const tables: EditorTableView[] = [];
  const elements = options.flattenSdt
    ? collectStoryBlockElements(root, options.suppressPlaceholderSdtContent === true)
    : directBlockElements(root);
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]!;
    if (el.localName === 'p') {
      const paragraph = buildEditorParagraph(
        storyId,
        el,
        options.context,
        options.attachedCommentIds,
        paragraphIdWithFallback(storyId, el, i),
        options.suppressPlaceholderSdtContent === true,
      );
      blocks.push(paragraph);
      paragraphs.push(paragraph);
    } else if (el.localName === 'tbl') {
      const tableId = options.tableIdByElement.get(el);
      const info = tableId ? options.tableInfoById.get(tableId) : undefined;
      if (!tableId || !info) continue;
      const table = buildEditorTable(
        info,
        el,
        storyId,
        options.context,
        options.attachedCommentIds,
        options.tableIdByElement,
        options.tableInfoById,
        options.suppressPlaceholderSdtContent === true,
      );
      blocks.push(table);
      tables.push(table);
      paragraphs.push(...collectTableParagraphs(table));
      tables.push(...collectNestedTables(table));
    }
  }
  return { blocks, paragraphs, tables };
}

function directBlockElements(root: Element): Element[] {
  const blocks: Element[] = [];
  for (let i = 0; i < root.childNodes.length; i++) {
    const child = root.childNodes[i];
    if (!child || child.nodeType !== 1 || (child as Element).namespaceURI !== W_NS) continue;
    const el = child as Element;
    if (el.localName === 'p' || el.localName === 'tbl') blocks.push(el);
  }
  return blocks;
}

function collectStoryBlockElements(root: Element, suppressPlaceholderSdtContent = false): Element[] {
  const blocks: Element[] = [];
  const visit = (el: Element): void => {
    if (el.namespaceURI !== W_NS) return;
    if (suppressPlaceholderSdtContent && isShowingPlaceholderSdt(el)) return;
    if (el.localName === 'p' || el.localName === 'tbl') {
      blocks.push(el);
      return;
    }
    if (el.localName !== 'sdt' && el.localName !== 'sdtContent') return;
    for (let i = 0; i < el.childNodes.length; i++) {
      const child = el.childNodes[i];
      if (child?.nodeType === 1) visit(child as Element);
    }
  };
  for (let i = 0; i < root.childNodes.length; i++) {
    const child = root.childNodes[i];
    if (child?.nodeType === 1) visit(child as Element);
  }
  return blocks;
}

function extractHeaderFooterStories(
  pkg: Package,
  options: BuildStoryBlocksOptions,
): { headers: EditorHeaderFooterStoryView[]; footers: EditorHeaderFooterStoryView[] } {
  const headers: EditorHeaderFooterStoryView[] = [];
  const footers: EditorHeaderFooterStoryView[] = [];
  for (const story of listStories(pkg, { kinds: ['header', 'footer'] })) {
    if (story.kind !== 'header' && story.kind !== 'footer') continue;
    const blocks = buildStoryBlocks(story.id, story.root, {
      ...options,
      flattenSdt: true,
      suppressPlaceholderSdtContent: true,
    });
    const view: EditorHeaderFooterStoryView = {
      kind: story.kind,
      partName: story.partPath,
      blocks: blocks.blocks,
      paragraphs: blocks.paragraphs,
      tables: blocks.tables,
    };
    if (story.kind === 'header') headers.push(view);
    else footers.push(view);
  }
  return { headers, footers };
}

function buildEditorParagraph(
  storyId: EditorParagraphView['story'],
  paragraphEl: Element,
  context: ExtractContext,
  attachedCommentIds: Set<number>,
  fallbackParaId?: string,
  suppressPlaceholderSdtContent = false,
): EditorParagraphView {
  const paragraphInfo = buildParagraphInfo(storyId, paragraphEl, 'full', true, {
    structuralWhitespace: true,
  });
  const paraId = paragraphInfo.paraId || fallbackParaId || '';
  const extracted = extractParagraphContent(paragraphEl, paraId, context, suppressPlaceholderSdtContent);
  for (const span of extracted.spans) {
    if (span.kind === 'comment') attachedCommentIds.add(span.commentId);
  }

  const view: EditorParagraphView = {
    type: 'paragraph',
    paraId,
    story: paragraphInfo.story,
    text: extracted.text,
    flags: {
      hasComments: paragraphInfo.hasComments === true,
      hasRevisions: paragraphInfo.hasRevisions === true,
      hasBookmarks: paragraphInfo.hasBookmarks === true,
      hasMath: paragraphInfo.hasMath === true,
    },
    segments: extracted.segments,
    spans: extracted.spans,
  };
  if (paragraphInfo.styleId !== undefined) view.styleId = paragraphInfo.styleId;
  const styleFormat =
    paragraphInfo.styleId !== undefined
      ? context.paragraphStyleFormats.get(paragraphInfo.styleId)
      : context.defaultParagraphStyleId
        ? context.paragraphStyleFormats.get(context.defaultParagraphStyleId)
        : undefined;
  const firstRunFormat = firstParagraphRunFormat(paragraphEl, context.theme);
  const directFormat = parseParagraphProperties(firstDirectDescendant(paragraphEl, W_NS, ['pPr']), context.theme);
  const mergedFormat = mergeParagraphFormats(
    mergeParagraphFormats(
      mergeParagraphFormats(context.documentDefaults.paragraph, styleFormat),
      firstRunFormat,
    ),
    paragraphInfo.format,
  );
  const finalFormat = mergeParagraphFormats(
    mergedFormat,
    directFormat,
  );
  if (finalFormat?.numbering) {
    finalFormat.numbering = resolveNumberingLabel(finalFormat.numbering, context);
    applyNumberingLevelParagraphFormat(finalFormat, context);
  }
  if (finalFormat !== undefined) view.format = finalFormat;
  if (paragraphInfo.table !== undefined) view.table = paragraphInfo.table;
  return view;
}

function firstParagraphRunFormat(paragraph: Element, theme: EditorTheme): EditorRunFormat | undefined {
  for (const child of directChildren(paragraph, W_NS, 'r')) {
    const format = parseRunFormat(child, theme);
    if (format) return format;
  }
  for (const wrapper of directChildren(paragraph, W_NS, 'hyperlink')) {
    const run = directChildren(wrapper, W_NS, 'r')[0];
    const format = run ? parseRunFormat(run, theme) : undefined;
    if (format) return format;
  }
  for (const wrapperName of ['ins', 'del', 'fldSimple'] as const) {
    for (const wrapper of directChildren(paragraph, W_NS, wrapperName)) {
      const run = directChildren(wrapper, W_NS, 'r')[0];
      const format = run ? parseRunFormat(run, theme) : undefined;
      if (format) return format;
    }
  }
  return undefined;
}

function buildEditorTable(
  info: TableInfo,
  tableEl: Element,
  storyId: EditorParagraphView['story'],
  context: ExtractContext,
  attachedCommentIds: Set<number>,
  tableIdByElement: Map<Element, string>,
  tableInfoById: Map<string, TableInfo>,
  suppressPlaceholderSdtContent = false,
): EditorTableView {
  const rowSpans = calculateRowSpans(info.cellGrid ?? []);
  const tableRowEls = directTableRows(tableEl);
  const tableCellEls = directTableCells(tableEl);
  const tableStyle = info.tableStyleId ? context.tableStyleFormats.get(info.tableStyleId) : undefined;
  const tableMargins = parseTableCellMargins(tableEl) ?? tableStyle?.cellMargin;
  const borders = parseTableBorders(tableEl) ?? tableStyle?.borders;
  const layout = parseTableLayout(tableEl);
  const tableLook = mergeTableLooks(tableStyle?.tableLook, parseTableLook(tableEl));
  let physicalCellIndex = 0;
  return {
    type: 'table',
    tableId: info.tableId,
    ...(info.tableStyleId ? { tableStyleId: info.tableStyleId } : {}),
    columnWidths: info.columnWidths,
    ...(layout ? { layout } : {}),
    ...(tableLook ? { tableLook } : {}),
    ...(borders ? { borders } : {}),
    rows:
      info.cellGrid?.map((row, rowIndex) => {
        const rowProps = parseTableRowProperties(tableRowEls[rowIndex]);
        return {
        ...(info.rowHeights[rowIndex] ? { height: info.rowHeights[rowIndex]! } : {}),
        ...(rowProps.isHeader ? { isHeader: true } : {}),
        ...(rowProps.cantSplit ? { cantSplit: true } : {}),
        cells: row.map((cell, cellIndex) => {
          const conditionalStyle = tableStyle
            ? conditionalTableStyleForCell(tableStyle, rowIndex, cellIndex, info.cellGrid ?? [], tableLook)
            : undefined;
          const cellEl = tableCellEls[physicalCellIndex++];
          const cellBlocks = cellEl
            ? buildEditorCellBlocks(cellEl, {
                storyId,
                tableId: info.tableId,
                rowIndex,
                cellIndex,
                context,
                attachedCommentIds,
                tableIdByElement,
                tableInfoById,
                suppressPlaceholderSdtContent,
              })
            : undefined;
          const paragraphs =
            cellBlocks?.paragraphs.length
              ? cellBlocks.paragraphs
              : cell.paragraphIds
                  .map((id) => context.indices.paraIdx.get(id.toUpperCase()))
                  .filter((el): el is Element => el !== undefined)
                  .map((el) =>
                    buildEditorParagraph(
                      storyId,
                      el,
                      context,
                      attachedCommentIds,
                      undefined,
                      suppressPlaceholderSdtContent,
                    ),
                  );
          const fallbackParagraphs =
            paragraphs.length > 0
              ? paragraphs
              : buildFallbackCellParagraphs(info.tableId, rowIndex, cellIndex, cell.text);
          const view: EditorTableCellView = {
            text: cell.text,
            gridSpan: cell.gridSpan,
            vMerge: cell.vMerge,
            blocks: cellBlocks?.blocks.length ? cellBlocks.blocks : fallbackParagraphs,
            paragraphs: fallbackParagraphs,
          };
          if (conditionalStyle?.paragraphFormat) {
            for (const paragraph of cellBlocks?.directParagraphs ?? view.paragraphs) {
              paragraph.format = mergeParagraphFormats(conditionalStyle.paragraphFormat, paragraph.format);
            }
          }
          const rowSpan = rowSpans.get(`${rowIndex}:${cellIndex}`);
          if (rowSpan !== undefined && rowSpan > 1) view.rowSpan = rowSpan;
          if (cell.width !== undefined) view.width = cell.width;
          const directCellProperties = parseCellProperties(cellEl, context.theme);
          const margin =
            directCellProperties?.margin ??
            conditionalStyle?.cell?.margin ??
            tableStyle?.cell?.margin ??
            tableMargins;
          if (margin) view.margin = margin;
          const borders =
            directCellProperties?.borders ??
            conditionalStyle?.cell?.borders ??
            tableStyle?.cell?.borders;
          if (borders) view.borders = borders;
          if (tableStyle?.cell?.shading !== undefined) view.shading = tableStyle.cell.shading;
          if (conditionalStyle?.cell?.shading !== undefined) view.shading = conditionalStyle.cell.shading;
          if (cell.shading !== undefined) view.shading = cell.shading;
          if (directCellProperties?.shading !== undefined) view.shading = directCellProperties.shading;
          if (tableStyle?.cell?.verticalAlignment !== undefined) {
            view.verticalAlignment = tableStyle.cell.verticalAlignment;
          }
          if (conditionalStyle?.cell?.verticalAlignment !== undefined) {
            view.verticalAlignment = conditionalStyle.cell.verticalAlignment;
          }
          if (directCellProperties?.verticalAlignment !== undefined) {
            view.verticalAlignment = directCellProperties.verticalAlignment;
          }
          if (cell.verticalAlignment !== undefined) view.verticalAlignment = cell.verticalAlignment;
          if (directCellProperties?.noWrap !== undefined) view.noWrap = directCellProperties.noWrap;
          return view;
        }),
      };
      }) ?? [],
  };
}

interface BuildEditorCellBlocksOptions {
  storyId: EditorParagraphView['story'];
  tableId: string;
  rowIndex: number;
  cellIndex: number;
  context: ExtractContext;
  attachedCommentIds: Set<number>;
  tableIdByElement: Map<Element, string>;
  tableInfoById: Map<string, TableInfo>;
  suppressPlaceholderSdtContent: boolean;
}

interface EditorCellBlocks {
  blocks: EditorBlockView[];
  paragraphs: EditorParagraphView[];
  directParagraphs: EditorParagraphView[];
}

function buildEditorCellBlocks(cellEl: Element, options: BuildEditorCellBlocksOptions): EditorCellBlocks {
  const blocks: EditorBlockView[] = [];
  const paragraphs: EditorParagraphView[] = [];
  const directParagraphs: EditorParagraphView[] = [];
  const elements = directBlockElements(cellEl);
  for (let index = 0; index < elements.length; index += 1) {
    const el = elements[index]!;
    if (el.localName === 'p') {
      const paragraph = buildEditorParagraph(
        options.storyId,
        el,
        options.context,
        options.attachedCommentIds,
        paragraphIdWithFallback(options.storyId, el, index),
        options.suppressPlaceholderSdtContent,
      );
      blocks.push(paragraph);
      paragraphs.push(paragraph);
      directParagraphs.push(paragraph);
      continue;
    }
    const tableId = options.tableIdByElement.get(el);
    const info = tableId ? options.tableInfoById.get(tableId) : undefined;
    if (!tableId || !info) continue;
    const table = buildEditorTable(
      info,
      el,
      options.storyId,
      options.context,
      options.attachedCommentIds,
      options.tableIdByElement,
      options.tableInfoById,
      options.suppressPlaceholderSdtContent,
    );
    blocks.push(table);
    paragraphs.push(...collectTableParagraphs(table));
  }
  return { blocks, paragraphs, directParagraphs };
}

function collectTableParagraphs(table: EditorTableView): EditorParagraphView[] {
  return table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs));
}

function collectNestedTables(table: EditorTableView): EditorTableView[] {
  return table.rows.flatMap((row) =>
    row.cells.flatMap((cell) =>
      cell.blocks.flatMap((block): EditorTableView[] =>
        block.type === 'table' ? [block, ...collectNestedTables(block)] : [],
      ),
    ),
  );
}

function parseTableRowProperties(rowEl: Element | undefined): Pick<EditorTableRowView, 'isHeader' | 'cantSplit'> {
  const row: Pick<EditorTableRowView, 'isHeader' | 'cantSplit'> = {};
  if (!rowEl) return row;
  const trPr = firstDirectChild(rowEl, W_NS, 'trPr');
  if (!trPr) return row;
  const tblHeader = parseOnOff(firstDirectChild(trPr, W_NS, 'tblHeader'));
  if (tblHeader !== undefined ? tblHeader : firstDirectChild(trPr, W_NS, 'tblHeader') !== null) {
    row.isHeader = true;
  }
  const cantSplit = parseOnOff(firstDirectChild(trPr, W_NS, 'cantSplit'));
  if (cantSplit !== undefined ? cantSplit : firstDirectChild(trPr, W_NS, 'cantSplit') !== null) {
    row.cantSplit = true;
  }
  return row;
}

function parseTableLayout(tableEl: Element): EditorTableView['layout'] | undefined {
  const tblPr = firstDirectChild(tableEl, W_NS, 'tblPr');
  if (!tblPr) return undefined;
  const layout: NonNullable<EditorTableView['layout']> = {};

  const tblW = firstDirectChild(tblPr, W_NS, 'tblW');
  if (tblW) {
    const width = numAttr(tblW, 'w');
    const widthType = parseTableWidthType(tblW.getAttributeNS(W_NS, 'type'));
    if (width !== undefined) layout.width = width;
    if (widthType !== undefined) layout.widthType = widthType;
  }

  const tblInd = firstDirectChild(tblPr, W_NS, 'tblInd');
  const indent = tblInd ? numAttr(tblInd, 'w') : undefined;
  if (indent !== undefined) layout.indent = indent;

  const tblCellSpacing = firstDirectChild(tblPr, W_NS, 'tblCellSpacing');
  if (tblCellSpacing) {
    const cellSpacing = numAttr(tblCellSpacing, 'w');
    const cellSpacingType = parseTableWidthType(tblCellSpacing.getAttributeNS(W_NS, 'type'));
    if (cellSpacing !== undefined) layout.cellSpacing = cellSpacing;
    if (cellSpacingType !== undefined) layout.cellSpacingType = cellSpacingType;
  }

  const jc = firstDirectChild(tblPr, W_NS, 'jc');
  const alignment = jc ? parseTableAlignment(jc.getAttributeNS(W_NS, 'val')) : undefined;
  if (alignment !== undefined) layout.alignment = alignment;

  return Object.keys(layout).length > 0 ? layout : undefined;
}

function parseTableLook(tableEl: Element): EditorTableLook | undefined {
  const tblPr = firstDirectChild(tableEl, W_NS, 'tblPr');
  if (!tblPr) return undefined;
  return parseTableLookFromPr(tblPr);
}

function parseTableLookFromPr(tblPr: Element): EditorTableLook | undefined {
  const tblLook = firstDirectChild(tblPr, W_NS, 'tblLook');
  const look: EditorTableLook = {};
  const val = tblLook?.getAttributeNS(W_NS, 'val') ?? null;
  const bits = val ? Number.parseInt(val, 16) : Number.NaN;
  const fromBit = (mask: number): boolean | undefined =>
    Number.isFinite(bits) ? (bits & mask) !== 0 : undefined;
  const boolAttr = (name: string, mask: number): boolean | undefined => {
    const value = tblLook?.getAttributeNS(W_NS, name) ?? null;
    if (value !== null) return parseOnOffValue(value);
    return fromBit(mask);
  };

  const firstRow = boolAttr('firstRow', 0x0020);
  if (firstRow !== undefined) look.firstRow = firstRow;
  const lastRow = boolAttr('lastRow', 0x0040);
  if (lastRow !== undefined) look.lastRow = lastRow;
  const firstColumn = boolAttr('firstColumn', 0x0080);
  if (firstColumn !== undefined) look.firstColumn = firstColumn;
  const lastColumn = boolAttr('lastColumn', 0x0100);
  if (lastColumn !== undefined) look.lastColumn = lastColumn;
  const noHorizontalBand = boolAttr('noHBand', 0x0200);
  if (noHorizontalBand !== undefined) look.noHorizontalBand = noHorizontalBand;
  const noVerticalBand = boolAttr('noVBand', 0x0400);
  if (noVerticalBand !== undefined) look.noVerticalBand = noVerticalBand;
  const rowBandSize = parsePositiveTableBandSize(firstDirectChild(tblPr, W_NS, 'tblStyleRowBandSize'));
  if (rowBandSize !== undefined) look.rowBandSize = rowBandSize;
  const columnBandSize = parsePositiveTableBandSize(firstDirectChild(tblPr, W_NS, 'tblStyleColBandSize'));
  if (columnBandSize !== undefined) look.columnBandSize = columnBandSize;
  return Object.keys(look).length > 0 ? look : undefined;
}

function mergeTableLooks(
  base: EditorTableLook | undefined,
  override: EditorTableLook | undefined,
): EditorTableLook | undefined {
  if (!base && !override) return undefined;
  return { ...(base ?? {}), ...(override ?? {}) };
}

function parsePositiveTableBandSize(el: Element | null): number | undefined {
  const size = el ? numAttr(el, 'val') : undefined;
  return size !== undefined && size > 0 ? size : undefined;
}

function parseTableWidthType(value: string | null): NonNullable<EditorTableView['layout']>['widthType'] | undefined {
  if (value === 'auto' || value === 'dxa' || value === 'pct') return value;
  return undefined;
}

function parseTableAlignment(value: string | null): NonNullable<EditorTableView['layout']>['alignment'] | undefined {
  if (value === 'left' || value === 'center' || value === 'right') return value;
  return undefined;
}

function directTableCells(tableEl: Element): Element[] {
  const cells: Element[] = [];
  for (const row of directChildren(tableEl, W_NS, 'tr')) {
    cells.push(...directChildren(row, W_NS, 'tc'));
  }
  return cells;
}

function directTableRows(tableEl: Element): Element[] {
  return directChildren(tableEl, W_NS, 'tr');
}

function parseTableCellMargins(tableEl: Element): EditorBoxEdges | undefined {
  const tblPr = firstDirectChild(tableEl, W_NS, 'tblPr');
  const tblCellMar = tblPr ? firstDirectChild(tblPr, W_NS, 'tblCellMar') : null;
  return parseMargins(tblCellMar);
}

function parseCellProperties(
  cellEl: Element | undefined,
  theme?: EditorTheme,
): Pick<EditorTableCellView, 'margin' | 'borders' | 'shading' | 'verticalAlignment' | 'noWrap'> | undefined {
  if (!cellEl) return undefined;
  return parseTableCellProperties(firstDirectChild(cellEl, W_NS, 'tcPr'), theme);
}

function parseMargins(parent: Element | null): EditorBoxEdges | undefined {
  if (!parent) return undefined;
  const margins: EditorBoxEdges = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const el = firstDirectChild(parent, W_NS, side);
    const value = el ? numAttr(el, 'w') : undefined;
    if (value !== undefined) margins[side] = value;
  }
  return Object.keys(margins).length > 0 ? margins : undefined;
}

function parseTableBorders(tableEl: Element): EditorTableBorders | undefined {
  const tblPr = firstDirectChild(tableEl, W_NS, 'tblPr');
  return tblPr ? parseTableBordersFromPr(tblPr) : undefined;
}

function parseTableBordersFromPr(tblPr: Element): EditorTableBorders | undefined {
  const tblBorders = firstDirectChild(tblPr, W_NS, 'tblBorders');
  if (!tblBorders) return undefined;
  const borders: EditorTableBorders = {};
  for (const side of ['top', 'right', 'bottom', 'left', 'insideH', 'insideV'] as const) {
    const border = parseBorder(firstDirectChild(tblBorders, W_NS, side));
    if (border) borders[side] = border;
  }
  return Object.keys(borders).length > 0 ? borders : undefined;
}

function parseCellBordersFromPr(tcPr: Element): EditorTableBorders | undefined {
  const tcBorders = firstDirectChild(tcPr, W_NS, 'tcBorders');
  if (!tcBorders) return undefined;
  const borders: EditorTableBorders = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const border = parseBorder(firstDirectChild(tcBorders, W_NS, side));
    if (border) borders[side] = border;
  }
  return Object.keys(borders).length > 0 ? borders : undefined;
}

function parseShading(el: Element | null, theme?: EditorTheme): EditorTableCellView['shading'] | undefined {
  if (!el) return undefined;
  const fill =
    normalizeRgbColor(el.getAttributeNS(W_NS, 'fill')) ??
    themeColor(theme, el.getAttributeNS(W_NS, 'themeFill'));
  const color = el.getAttributeNS(W_NS, 'color') ?? undefined;
  const val = el.getAttributeNS(W_NS, 'val') ?? undefined;
  return {
    ...(fill
      ? {
          fill: applyThemeTintOrShade(
            fill,
            el.getAttributeNS(W_NS, 'themeFillTint'),
            el.getAttributeNS(W_NS, 'themeFillShade'),
          ),
        }
      : {}),
    ...(color ? { color } : {}),
    ...(val ? { val } : {}),
  };
}

function parseVerticalAlignment(el: Element | null): EditorTableCellView['verticalAlignment'] | undefined {
  const value = el?.getAttributeNS(W_NS, 'val');
  return value === 'top' || value === 'center' || value === 'bottom' || value === 'both'
    ? value
    : undefined;
}

function parseBorder(el: Element | null): NonNullable<EditorTableBorders['top']> | undefined {
  if (!el) return undefined;
  const val = el.getAttributeNS(W_NS, 'val') ?? undefined;
  const color = el.getAttributeNS(W_NS, 'color') ?? undefined;
  const size = numAttr(el, 'sz');
  return {
    ...(val ? { val } : {}),
    ...(color ? { color } : {}),
    ...(size !== undefined ? { size } : {}),
  };
}

function buildFallbackCellParagraphs(
  tableId: string,
  rowIndex: number,
  cellIndex: number,
  text: string,
): EditorParagraphView[] {
  if (text.length === 0) return [];
  return text.split('\n').map((part, paragraphIndex) => ({
    type: 'paragraph',
    paraId: `${tableId}:r${rowIndex}:c${cellIndex}:p${paragraphIndex}`,
    story: 'body',
    text: part,
    flags: {
      hasComments: false,
      hasRevisions: false,
      hasBookmarks: false,
      hasMath: false,
    },
    segments: part.length > 0 ? [{ kind: 'text', from: 0, to: part.length, text: part }] : [],
    spans: [],
  }));
}

function calculateRowSpans(cellGrid: NonNullable<TableInfo['cellGrid']>): Map<string, number> {
  const positions = calculateCellGridPositions(cellGrid);
  const spans = new Map<string, number>();

  for (let rowIndex = 0; rowIndex < positions.length; rowIndex++) {
    const row = positions[rowIndex] ?? [];
    for (let cellIndex = 0; cellIndex < row.length; cellIndex++) {
      const current = row[cellIndex];
      if (!current || current.cell.vMerge !== 'restart') continue;
      let rowSpan = 1;
      for (let nextRowIndex = rowIndex + 1; nextRowIndex < positions.length; nextRowIndex++) {
        const continuation = positions[nextRowIndex]?.find(
          (candidate) =>
            candidate.startColumn === current.startColumn &&
            candidate.cell.vMerge === 'continue',
        );
        if (!continuation) break;
        rowSpan += 1;
      }
      if (rowSpan > 1) spans.set(`${rowIndex}:${cellIndex}`, rowSpan);
    }
  }

  return spans;
}

function calculateCellGridPositions(cellGrid: NonNullable<TableInfo['cellGrid']>): TableCellGridPosition[][] {
  return cellGrid.map((row) => {
    let column = 0;
    return row.map((cell) => {
      const startColumn = column;
      const columnSpan = Math.max(1, cell.gridSpan);
      column += columnSpan;
      return { startColumn, endColumn: column - 1, columnSpan, cell };
    });
  });
}

interface TableCellGridPosition {
  startColumn: number;
  endColumn: number;
  columnSpan: number;
  cell: NonNullable<TableInfo['cellGrid']>[number][number];
}

function conditionalTableStyleForCell(
  tableStyle: EditorTableStyleFormat,
  rowIndex: number,
  cellIndex: number,
  cellGrid: NonNullable<TableInfo['cellGrid']>,
  tableLook: EditorTableLook | undefined,
): EditorTableConditionalStyle | undefined {
  const conditions = tableStyle.conditions;
  if (!conditions) return undefined;
  const positions = calculateCellGridPositions(cellGrid);
  const cellPosition = positions[rowIndex]?.[cellIndex];
  if (!cellPosition) return undefined;
  const rowCount = cellGrid.length;
  const maxGridColumn = Math.max(
    0,
    ...positions.flatMap((row) => row.map((position) => position.endColumn)),
  );
  const effectiveLook = effectiveTableLook(tableLook);
  const rowBandSize = effectiveLook.rowBandSize;
  const columnBandSize = effectiveLook.columnBandSize;
  let style: EditorTableConditionalStyle | undefined;
  if (!effectiveLook.noHorizontalBand && Math.floor(rowIndex / rowBandSize) % 2 === 0) {
    style = mergeTableConditionalStyles(style, conditions.band1Horz);
  } else if (!effectiveLook.noHorizontalBand) {
    style = mergeTableConditionalStyles(style, conditions.band2Horz);
  }
  if (!effectiveLook.noVerticalBand && Math.floor(cellPosition.startColumn / columnBandSize) % 2 === 0) {
    style = mergeTableConditionalStyles(style, conditions.band1Vert);
  } else if (!effectiveLook.noVerticalBand) {
    style = mergeTableConditionalStyles(style, conditions.band2Vert);
  }
  if (effectiveLook.firstRow && rowIndex === 0) style = mergeTableConditionalStyles(style, conditions.firstRow);
  if (effectiveLook.lastRow && rowIndex === rowCount - 1) {
    style = mergeTableConditionalStyles(style, conditions.lastRow);
  }
  if (effectiveLook.firstColumn && cellPosition.startColumn === 0) {
    style = mergeTableConditionalStyles(style, conditions.firstCol);
  }
  if (effectiveLook.lastColumn && cellPosition.endColumn === maxGridColumn) {
    style = mergeTableConditionalStyles(style, conditions.lastCol);
  }
  if (effectiveLook.firstRow && effectiveLook.firstColumn && rowIndex === 0 && cellPosition.startColumn === 0) {
    style = mergeTableConditionalStyles(style, conditions.nwCell);
  }
  if (effectiveLook.firstRow && effectiveLook.lastColumn && rowIndex === 0 && cellPosition.endColumn === maxGridColumn) {
    style = mergeTableConditionalStyles(style, conditions.neCell);
  }
  if (effectiveLook.lastRow && effectiveLook.firstColumn && rowIndex === rowCount - 1 && cellPosition.startColumn === 0) {
    style = mergeTableConditionalStyles(style, conditions.swCell);
  }
  if (effectiveLook.lastRow && effectiveLook.lastColumn && rowIndex === rowCount - 1 && cellPosition.endColumn === maxGridColumn) {
    style = mergeTableConditionalStyles(style, conditions.seCell);
  }
  return style;
}

function effectiveTableLook(tableLook: EditorTableLook | undefined): Required<EditorTableLook> {
  return {
    ...DEFAULT_TABLE_LOOK,
    ...(tableLook ?? {}),
    rowBandSize: tableLook?.rowBandSize ?? 1,
    columnBandSize: tableLook?.columnBandSize ?? 1,
  };
}

function extractParagraphContent(
  paragraph: Element,
  paraId: string,
  context: ExtractContext,
  suppressPlaceholderSdtContent = false,
): { text: string; segments: EditorContentSegment[]; spans: EditorParagraphSpan[] } {
  const segments: EditorContentSegment[] = [];
  const spans: EditorParagraphSpan[] = [];
  const activeComments = new Map<number, number>();
  const consumedDrawingIds = new Set<string>();
  let activeField: ActiveComplexField | null = null;
  let text = '';
  const collectParagraphVisibleText = (root: Element): string =>
    collectVisibleText(root, suppressPlaceholderSdtContent);

  const appendText = (value: string): { from: number; to: number } => {
    const from = text.length;
    text += value;
    return { from, to: text.length };
  };

  const children = paragraph.childNodes;
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (!child || child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI !== W_NS) {
      appendPlaceholder('unknown', el);
      continue;
    }

    if (el.localName === 'pPr') continue;

    if (el.localName === 'hyperlink') {
      appendHyperlink(el);
      continue;
    }

    const commentStartId = markerId(el, 'commentRangeStart');
    if (commentStartId !== null) {
      activeComments.set(commentStartId, text.length);
      continue;
    }

    const commentEndId = markerId(el, 'commentRangeEnd');
    if (commentEndId !== null) {
      const from = activeComments.get(commentEndId);
      if (from !== undefined) {
        const span = buildCommentSpan(context.indices, commentEndId, paraId, from, text.length, text);
        if (span) spans.push(span);
        activeComments.delete(commentEndId);
      }
      continue;
    }

    if (el.localName === 'ins' || el.localName === 'del') {
      const revisionKind = el.localName === 'ins' ? 'insert' : 'delete';
      const revisionText = collectText(el, revisionKind === 'delete' ? 'delText' : 't');
      const { from, to } = appendText(revisionText);
      const revisionId = el.getAttributeNS(W_NS, 'id') ?? '';
      const revisionRun = directChildren(el, W_NS, 'r')[0];
      const format = revisionRun ? parseRunFormat(revisionRun, context.theme) : undefined;
      const segment: EditorRevisionSegment = {
        kind: 'revision',
        from,
        to,
        text: revisionText,
        ...(format ? { format } : {}),
        revisionId,
        revisionKind,
        author: el.getAttributeNS(W_NS, 'author') ?? '',
        date: el.getAttributeNS(W_NS, 'date') ?? '',
      };
      segments.push(segment);
      if (revisionId.length > 0) {
        spans.push({
          kind: 'revision',
          id: `revision:${revisionId}`,
          revisionId,
          revisionKind,
          paraId,
          from,
          to,
          text: revisionText,
          author: segment.author,
          date: segment.date,
        });
      }
      continue;
    }

    if (el.localName === 'r' && containsFieldControl(el)) {
      appendRunWithFieldControls(el);
      continue;
    }

    if (activeField && el.localName === 'r') {
      appendFieldResultRun(el);
      continue;
    }

    if (el.localName === 'fldSimple') {
      if (appendSimpleField(el)) continue;
    }

    if (el.localName === 'r' && isMarkerOnlyRun(el)) continue;

    if (el.localName === 'r' && containsDirectPageBreak(el)) {
      appendRunContentWithPageBreaks(el);
      continue;
    }

    const directDrawings = el.localName === 'r' ? directDrawingChildren(el) : [];
    if (directDrawings.length > 0) {
      for (const drawing of directDrawings) appendDrawing(drawing);
      continue;
    }

    if (containsDirectDrawing(el)) {
      appendDrawing(el);
      continue;
    }

    const runText = collectParagraphVisibleText(el);
    if (runText.length > 0) {
      const { from, to } = appendText(runText);
      const segment: EditorContentSegment = { kind: 'text', from, to, text: runText };
      if (el.localName === 'r') {
        const format = parseRunFormat(el, context.theme);
        if (format) segment.format = format;
      }
      segments.push(segment);
    }
  }
  flushUnknownField();

  appendUnconsumedDescendantDrawings(paragraph);
  for (const drawing of context.drawingsByAnchorParaId.get(paraId) ?? []) {
    if (!consumedDrawingIds.has(drawing.drawingId)) appendDrawingInfo(drawing);
  }

  return {
    text,
    segments,
    spans: spans.sort((a, b) => a.from - b.from || a.to - b.to || a.id.localeCompare(b.id)),
  };

  function appendPlaceholder(
    placeholderKind: EditorPlaceholderSegment['placeholderKind'],
    source: Element,
  ): void {
    const { from, to } = appendText('\uFFFC');
    segments.push({
      kind: 'placeholder',
      from,
      to,
      placeholderKind,
      sourceTag: source.localName,
    });
  }

  function appendHyperlink(hyperlink: Element): void {
    const link = context.hyperlinkByElement.get(hyperlink);
    for (let index = 0; index < hyperlink.childNodes.length; index += 1) {
      const child = hyperlink.childNodes[index];
      if (!child || child.nodeType !== 1) continue;
      const run = child as Element;
      if (run.namespaceURI !== W_NS || run.localName !== 'r') continue;
      const value = collectParagraphVisibleText(run);
      if (value.length === 0) continue;
      const { from, to } = appendText(value);
      const segment: EditorTextSegment = { kind: 'text', from, to, text: value };
      const format = parseRunFormat(run, context.theme);
      if (format) segment.format = format;
      if (link) segment.hyperlink = { ...link };
      segments.push(segment);
    }
  }

  function appendRunWithFieldControls(run: Element): void {
    const format = parseRunFormat(run, context.theme);
    for (let i = 0; i < run.childNodes.length; i++) {
      const child = run.childNodes[i];
      if (!child || child.nodeType !== 1) continue;
      const childEl = child as Element;
      if (childEl.namespaceURI !== W_NS) continue;

      if (childEl.localName === 'fldChar') {
        const type = childEl.getAttributeNS(W_NS, 'fldCharType');
        if (type === 'begin') {
          flushUnknownField();
          activeField = { instruction: '', cachedText: '', inResult: false, format };
        } else if (type === 'separate' && activeField) {
          activeField.inResult = true;
        } else if (type === 'end' && activeField) {
          appendComplexField(activeField);
          activeField = null;
        }
        continue;
      }

      if (activeField && childEl.localName === 'instrText') {
        activeField.instruction += childEl.textContent ?? '';
        continue;
      }

      const value = visibleTextFromFieldRunChild(childEl);
      if (value.length === 0) continue;
      if (activeField) {
        if (activeField.inResult) activeField.cachedText += value;
        continue;
      }
      const { from, to } = appendText(value);
      const segment: EditorContentSegment = { kind: 'text', from, to, text: value };
      if (format) segment.format = format;
      segments.push(segment);
    }
  }

  function appendSimpleField(field: Element): boolean {
    const instruction = field.getAttributeNS(W_NS, 'instr') ?? '';
    const fieldKind = fieldKindFromInstruction(instruction);
    if (!fieldKind) return false;
    const cachedText = collectParagraphVisibleText(field);
    appendFieldSegment(fieldKind, instruction, cachedText, parseRunFormat(field, context.theme));
    return true;
  }

  function appendFieldResultRun(run: Element): void {
    const value = collectParagraphVisibleText(run);
    if (value.length > 0 && activeField?.inResult) activeField.cachedText += value;
  }

  function appendComplexField(field: ActiveComplexField): void {
    const fieldKind = fieldKindFromInstruction(field.instruction);
    if (fieldKind) {
      appendFieldSegment(fieldKind, field.instruction, field.cachedText, field.format);
      return;
    }
    appendCachedFieldText(field);
  }

  function appendFieldSegment(
    fieldKind: EditorFieldSegment['fieldKind'],
    instruction: string,
    cachedText: string,
    format: EditorRunFormat | undefined,
  ): void {
    const { from, to } = appendText('\uFFFC');
    const segment: EditorFieldSegment = {
      kind: 'field',
      from,
      to,
      fieldKind,
      instruction,
      cachedText,
    };
    if (format) segment.format = format;
    segments.push(segment);
  }

  function flushUnknownField(): void {
    if (!activeField) return;
    appendCachedFieldText(activeField);
    activeField = null;
  }

  function appendCachedFieldText(field: ActiveComplexField): void {
    if (field.cachedText.length === 0) return;
    const { from, to } = appendText(field.cachedText);
    const segment: EditorContentSegment = { kind: 'text', from, to, text: field.cachedText };
    if (field.format) segment.format = field.format;
    segments.push(segment);
  }

  function appendRunContentWithPageBreaks(run: Element): void {
    const format = parseRunFormat(run, context.theme);
    for (let i = 0; i < run.childNodes.length; i++) {
      const child = run.childNodes[i];
      if (!child || child.nodeType !== 1) continue;
      const childEl = child as Element;
      if (childEl.namespaceURI !== W_NS) continue;
      if (childEl.localName === 'br' && childEl.getAttributeNS(W_NS, 'type') === 'page') {
        appendPlaceholder('pageBreak', childEl);
        continue;
      }
      const runText = collectParagraphVisibleText(childEl);
      if (runText.length === 0) continue;
      const { from, to } = appendText(runText);
      const segment: EditorContentSegment = { kind: 'text', from, to, text: runText };
      if (format) segment.format = format;
      segments.push(segment);
    }
  }

  function appendDrawing(source: Element): void {
    const drawingId = context.drawingIdByElement.get(source);
    const image = drawingId ? context.imagesByDrawingId.get(drawingId) : undefined;
    const drawing = drawingId ? context.drawingInfoById.get(drawingId) : undefined;
    if (drawingId && drawing && drawing.kind !== 'picture') {
      consumedDrawingIds.add(drawingId);
      appendDrawingInfo(drawing);
      return;
    }
    if (!drawingId || !image) {
      appendPlaceholder('drawing', source);
      return;
    }
    const { from, to } = appendText('\uFFFC');
    consumedDrawingIds.add(drawingId);
    const segment: EditorImageSegment = {
      kind: 'image',
      from,
      to,
      drawingId,
      src: image.src,
      contentType: image.contentType,
    };
    if (image.width !== undefined) segment.width = image.width;
    if (image.height !== undefined) segment.height = image.height;
    if (image.anchoringMode !== undefined) segment.anchoringMode = image.anchoringMode;
    if (image.wrap !== undefined) segment.wrap = image.wrap;
    if (image.position !== undefined) segment.position = image.position;
    segments.push(segment);
  }

  function appendUnconsumedDescendantDrawings(source: Element): void {
    const drawings = source.getElementsByTagNameNS(W_NS, 'drawing');
    for (let i = 0; i < drawings.length; i++) {
      const drawing = drawings[i];
      if (!drawing) continue;
      const drawingId = context.drawingIdByElement.get(drawing);
      if (!drawingId || consumedDrawingIds.has(drawingId)) continue;
      appendDrawing(drawing);
    }
  }

  function appendDrawingInfo(drawing: DrawingInfo): void {
    if (drawing.kind === 'picture') return;
    consumedDrawingIds.add(drawing.drawingId);
    if (isHiddenTemplateFieldTextBox(drawing, context)) return;
    const visibleText = drawing.kind === 'textBox' ? (drawing.innerText ?? '') : '';
    const { from, to } = appendText(visibleText.length > 0 ? visibleText : '\uFFFC');
    const segment: EditorDrawingSegment = {
      kind: 'drawing',
      from,
      to,
      drawingId: drawing.drawingId,
      drawingKind: drawing.kind,
    };
    if (visibleText.length > 0) segment.text = visibleText;
    const size = emuExtentToCssPixels(drawing.extentEMU);
    if (size.width !== undefined) segment.width = size.width;
    if (size.height !== undefined) segment.height = size.height;
    if (drawing.anchoringMode !== undefined) segment.anchoringMode = drawing.anchoringMode;
    if (drawing.wrap !== undefined) segment.wrap = drawing.wrap;
    if (drawing.position !== undefined) segment.position = drawing.position;
    if (drawing.hasBorder !== undefined) segment.hasBorder = drawing.hasBorder;
    if (drawing.hasFill !== undefined) segment.hasFill = drawing.hasFill;
    if (drawing.isLineOrConnector !== undefined) segment.isLineOrConnector = drawing.isLineOrConnector;
    if (drawing.innerTextAlign !== undefined) segment.innerTextAlign = drawing.innerTextAlign;
    const innerParagraphs = buildDrawingInnerParagraphs(drawing, context);
    if (innerParagraphs.length > 0) segment.paragraphs = innerParagraphs;
    segments.push(segment);
  }
}

const HIDDEN_TEMPLATE_TEXT_BOX_MAX_EMU = 127_000; // 10 pt

/**
 * WPS/Word templates can keep SET fields in a tiny off-page text box. The cached
 * field results are data, not document copy; projecting them as normal text makes
 * the 5 pt box expand into several visible lines. Keep ordinary text boxes and
 * fields intact by requiring every one of the template's layout signals.
 */
function isHiddenTemplateFieldTextBox(drawing: DrawingInfo, context: ExtractContext): boolean {
  if (
    drawing.kind !== 'textBox' ||
    drawing.anchoringMode !== 'anchor' ||
    drawing.extentEMU.cx <= 0 ||
    drawing.extentEMU.cy <= 0 ||
    drawing.extentEMU.cx > HIDDEN_TEMPLATE_TEXT_BOX_MAX_EMU ||
    drawing.extentEMU.cy > HIDDEN_TEMPLATE_TEXT_BOX_MAX_EMU
  ) {
    return false;
  }

  const horizontalOffset = drawing.position?.horizontal?.posOffset;
  const verticalOffset = drawing.position?.vertical?.posOffset;
  if (!(
    (horizontalOffset !== undefined && horizontalOffset < 0) ||
    (verticalOffset !== undefined && verticalOffset < 0)
  )) {
    return false;
  }

  const source = context.indices.drawingIdx.get(drawing.drawingId);
  if (!source) return false;
  const textBoxContent = source.getElementsByTagNameNS(W_NS, 'txbxContent')[0];
  return textBoxContent ? containsOnlySetFieldResults(textBoxContent) : false;
}

function containsOnlySetFieldResults(root: Element): boolean {
  const fields: Array<{ instruction: string }> = [];
  let fieldCount = 0;
  const descendants = root.getElementsByTagName('*');

  for (let i = 0; i < descendants.length; i += 1) {
    const el = descendants[i];
    if (!el || el.namespaceURI !== W_NS) continue;

    if (el.localName === 'fldChar') {
      const type = el.getAttributeNS(W_NS, 'fldCharType');
      if (type === 'begin') {
        fields.push({ instruction: '' });
      } else if (type === 'end') {
        const field = fields.pop();
        if (!field || !/^SET\b/i.test(field.instruction.trim())) return false;
        fieldCount += 1;
      }
      continue;
    }

    if (el.localName === 'instrText') {
      const field = fields[fields.length - 1];
      if (!field) return false;
      field.instruction += el.textContent ?? '';
      continue;
    }

    if (el.localName === 't' && (el.textContent ?? '').trim().length > 0 && fields.length === 0) {
      return false;
    }
  }

  return fields.length === 0 && fieldCount > 0;
}

function buildDrawingInnerParagraphs(
  drawing: DrawingInfo,
  context: ExtractContext,
): EditorParagraphView[] {
  if (drawing.kind !== 'textBox' || !drawing.innerParagraphIds || drawing.innerParagraphIds.length === 0) {
    return [];
  }
  const attachedCommentIds = new Set<number>();
  return drawing.innerParagraphIds.flatMap((paraId) => {
    const paragraph = context.indices.paraIdx.get(paraId);
    if (!paragraph) return [];
    return [
      buildEditorParagraph(
        'body',
        paragraph,
        context,
        attachedCommentIds,
        paraId,
      ),
    ];
  });
}

interface ExtractContext {
  indices: PackageIndices;
  drawingIdByElement: Map<Element, string>;
  drawingInfoById: Map<string, DrawingInfo>;
  drawingsByAnchorParaId: Map<string, DrawingInfo[]>;
  imagesByDrawingId: Map<string, EditorImageData>;
  theme: EditorTheme;
  documentDefaults: EditorDocumentDefaults;
  numbering: NumberingDefinitions;
  numberingCounters: Map<string, number>;
  defaultParagraphStyleId?: string;
  paragraphStyleFormats: Map<string, EditorParagraphFormat>;
  tableStyleFormats: Map<string, EditorTableStyleFormat>;
  hyperlinkByElement: Map<Element, EditorHyperlink>;
}

interface EditorTheme {
  colors: Map<string, string>;
  fonts: Map<string, string>;
}

interface EditorImageData {
  src: string;
  contentType: string;
  width?: number;
  height?: number;
  anchoringMode?: EditorImageSegment['anchoringMode'];
  wrap?: EditorImageSegment['wrap'];
  position?: EditorImageSegment['position'];
}

interface ActiveComplexField {
  instruction: string;
  cachedText: string;
  inResult: boolean;
  format?: EditorRunFormat;
}

function buildEditorImage(
  pkg: Package,
  indices: PackageIndices,
  drawingId: string,
  drawingInfo?: DrawingInfo,
): EditorImageData | null {
  const result = getImageBytesSync(pkg, indices, drawingId);
  if (!result) return null;
  const image: EditorImageData = {
    src: `data:${result.contentType};base64,${bytesToBase64(result.bytes)}`,
    contentType: result.contentType,
  };
  if (drawingInfo) {
    image.anchoringMode = drawingInfo.anchoringMode;
    if (drawingInfo.wrap !== undefined) image.wrap = drawingInfo.wrap;
    if (drawingInfo.position !== undefined) image.position = drawingInfo.position;
  }
  const size = emuSizeToCssPixels(result);
  if (size.width !== undefined) image.width = size.width;
  if (size.height !== undefined) image.height = size.height;
  return image;
}

function emuSizeToCssPixels(result: ImageBytesResult): { width?: number; height?: number } {
  return emuExtentToCssPixels(result.extentEMU);
}

function emuExtentToCssPixels(extent: { cx: number; cy: number }): { width?: number; height?: number } {
  const { cx, cy } = extent;
  return {
    width: cx > 0 ? Math.round((cx / 914400) * 96) : undefined,
    height: cy > 0 ? Math.round((cy / 914400) * 96) : undefined,
  };
}

function bytesToBase64(bytes: Uint8Array): string {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  let out = '';
  let i = 0;
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i]! << 16) | (bytes[i + 1]! << 8) | bytes[i + 2]!;
    out += b64Char(chars, (n >> 18) & 63);
    out += b64Char(chars, (n >> 12) & 63);
    out += b64Char(chars, (n >> 6) & 63);
    out += b64Char(chars, n & 63);
  }
  if (i < bytes.length) {
    const a = bytes[i]!;
    const b = i + 1 < bytes.length ? bytes[i + 1]! : 0;
    const n = (a << 16) | (b << 8);
    out += b64Char(chars, (n >> 18) & 63);
    out += b64Char(chars, (n >> 12) & 63);
    out += i + 1 < bytes.length ? b64Char(chars, (n >> 6) & 63) : '=';
    out += '=';
  }
  return out;
}

function b64Char(chars: string, index: number): string {
  return chars.charAt(index);
}

function extractDocumentDefaults(pkg: Package, theme: EditorTheme): EditorDocumentDefaults {
  const defaults: EditorDocumentDefaults = {};
  const stylesDom = findStylesDom(pkg);
  if (!stylesDom) return defaults;
  const docDefaults = stylesDom.getElementsByTagNameNS(W_NS, 'docDefaults')[0];
  if (!docDefaults) return defaults;

  const rPrDefault = firstDirectDescendant(docDefaults, W_NS, ['rPrDefault', 'rPr']);
  const run = rPrDefault ? parseRunProperties(rPrDefault, theme) : undefined;
  if (run) defaults.run = run;

  const pPrDefault = firstDirectDescendant(docDefaults, W_NS, ['pPrDefault', 'pPr']);
  if (pPrDefault) {
    const paragraph = parseParagraphProperties(pPrDefault, theme);
    if (paragraph) defaults.paragraph = paragraph;
  }
  return defaults;
}

function findStylesDom(pkg: Package): Document | null {
  for (const part of pkg.parts.values()) {
    if (
      part.contentType ===
        'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' &&
      part.xmlDom
    ) {
      return part.xmlDom;
    }
  }
  return null;
}

interface EditorParagraphStyleData {
  basedOn?: string;
  format?: EditorParagraphFormat;
}

interface EditorTableStyleData {
  basedOn?: string;
  format?: EditorTableStyleFormat;
}

interface EditorTableStyleFormat {
  borders?: EditorTableBorders;
  cellMargin?: EditorBoxEdges;
  tableLook?: EditorTableLook;
  cell?: Pick<EditorTableCellView, 'margin' | 'borders' | 'shading' | 'verticalAlignment' | 'noWrap'>;
  paragraphFormat?: EditorParagraphFormat;
  conditions?: Partial<Record<EditorTableStyleCondition, EditorTableConditionalStyle>>;
}

type EditorTableStyleCondition =
  | 'firstRow'
  | 'lastRow'
  | 'firstCol'
  | 'lastCol'
  | 'band1Horz'
  | 'band2Horz'
  | 'band1Vert'
  | 'band2Vert'
  | 'nwCell'
  | 'neCell'
  | 'swCell'
  | 'seCell';

interface EditorTableConditionalStyle {
  cell?: Pick<EditorTableCellView, 'margin' | 'borders' | 'shading' | 'verticalAlignment' | 'noWrap'>;
  paragraphFormat?: EditorParagraphFormat;
}

interface NumberingLevelDefinition {
  ilvl: number;
  numFmt?: string;
  lvlText?: string;
  start: number;
  paragraphFormat?: EditorParagraphFormat;
}

interface NumberingDefinitions {
  nums: Map<string, Map<number, NumberingLevelDefinition>>;
}

function buildNumberingDefinitions(pkg: Package): NumberingDefinitions {
  const numberingDom = findNumberingDom(pkg);
  if (!numberingDom) return { nums: new Map() };

  const abstractNums = new Map<string, Map<number, NumberingLevelDefinition>>();
  const abstractNumEls = numberingDom.getElementsByTagNameNS(W_NS, 'abstractNum');
  for (let i = 0; i < abstractNumEls.length; i++) {
    const abstractNum = abstractNumEls[i];
    if (!abstractNum) continue;
    const abstractNumId = abstractNum.getAttributeNS(W_NS, 'abstractNumId');
    if (!abstractNumId) continue;
    const levels = new Map<number, NumberingLevelDefinition>();
    for (const lvl of directChildren(abstractNum, W_NS, 'lvl')) {
      const ilvl = numAttr(lvl, 'ilvl') ?? 0;
      const numFmt = firstDirectChild(lvl, W_NS, 'numFmt')?.getAttributeNS(W_NS, 'val') ?? undefined;
      const lvlText = firstDirectChild(lvl, W_NS, 'lvlText')?.getAttributeNS(W_NS, 'val') ?? undefined;
      const startEl = firstDirectChild(lvl, W_NS, 'start');
      const start = startEl ? numAttr(startEl, 'val') ?? 1 : 1;
      const paragraphFormat = parseParagraphProperties(firstDirectChild(lvl, W_NS, 'pPr'));
      levels.set(ilvl, {
        ilvl,
        start,
        ...(numFmt ? { numFmt } : {}),
        ...(lvlText ? { lvlText } : {}),
        ...(paragraphFormat ? { paragraphFormat } : {}),
      });
    }
    abstractNums.set(abstractNumId, levels);
  }

  const nums = new Map<string, Map<number, NumberingLevelDefinition>>();
  const numEls = numberingDom.getElementsByTagNameNS(W_NS, 'num');
  for (let i = 0; i < numEls.length; i++) {
    const num = numEls[i];
    if (!num) continue;
    const numId = num.getAttributeNS(W_NS, 'numId');
    const abstractNumId = firstDirectChild(num, W_NS, 'abstractNumId')?.getAttributeNS(W_NS, 'val');
    if (!numId || !abstractNumId) continue;
    const levels = abstractNums.get(abstractNumId);
    if (levels) nums.set(numId, levels);
  }
  return { nums };
}

function applyNumberingLevelParagraphFormat(
  format: EditorParagraphFormat,
  context: ExtractContext,
): void {
  if (!format.numbering) return;
  const level = context.numbering.nums.get(format.numbering.numId)?.get(format.numbering.ilvl);
  const levelFormat = level?.paragraphFormat;
  if (!levelFormat) return;

  if (levelFormat.indent !== undefined) {
    const existing = format.indent ?? {};
    format.indent = { ...existing };
    if (levelFormat.indent.left !== undefined && (existing.left === undefined || existing.left === 0)) {
      format.indent.left = levelFormat.indent.left;
    }
    if (levelFormat.indent.right !== undefined && existing.right === undefined) {
      format.indent.right = levelFormat.indent.right;
    }
    if (levelFormat.indent.firstLine !== undefined && existing.firstLine === undefined) {
      format.indent.firstLine = levelFormat.indent.firstLine;
    }
  }
  if (format.spacing === undefined && levelFormat.spacing !== undefined) {
    format.spacing = { ...levelFormat.spacing };
  }
  if (format.alignment === undefined && levelFormat.alignment !== undefined) {
    format.alignment = levelFormat.alignment;
  }
}

function findNumberingDom(pkg: Package): Document | null {
  for (const part of pkg.parts.values()) {
    if (
      part.contentType ===
        'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml' &&
      part.xmlDom
    ) {
      return part.xmlDom;
    }
  }
  return null;
}

function resolveNumberingLabel(
  numbering: EditorNumberingFormat,
  context: ExtractContext,
): EditorNumberingFormat {
  const level = context.numbering.nums.get(numbering.numId)?.get(numbering.ilvl);
  const numFmt = numbering.numFmt ?? level?.numFmt;
  const lvlText = numbering.lvlText ?? level?.lvlText;
  const resolved: EditorNumberingFormat = {
    ...numbering,
    ...(numFmt ? { numFmt } : {}),
    ...(lvlText ? { lvlText } : {}),
    ...(level?.start !== undefined ? { start: level.start } : {}),
  };
  const key = `${resolved.numId}:${resolved.ilvl}`;
  if (numFmt === 'bullet') {
    resolved.label = normalizeBulletLabel(lvlText);
    return resolved;
  }
  const current = (context.numberingCounters.get(key) ?? ((level?.start ?? 1) - 1)) + 1;
  context.numberingCounters.set(key, current);
  resolved.label = formatNumberingLabel(current, numFmt, lvlText);
  return resolved;
}

function normalizeBulletLabel(lvlText: string | undefined): string {
  const label = lvlText?.replace(/%\d+/g, '').trim();
  return label && label.length > 0 ? label : '•';
}

function formatNumberingLabel(value: number, numFmt: string | undefined, lvlText: string | undefined): string {
  const formatted =
    numFmt === 'lowerLetter'
      ? alphaLabel(value, false)
      : numFmt === 'upperLetter'
        ? alphaLabel(value, true)
        : String(value);
  if (!lvlText) return `${formatted}.`;
  return lvlText.replace(/%\d+/g, formatted);
}

function alphaLabel(value: number, upper: boolean): string {
  let n = Math.max(1, value);
  let out = '';
  while (n > 0) {
    n -= 1;
    out = String.fromCharCode((upper ? 65 : 97) + (n % 26)) + out;
    n = Math.floor(n / 26);
  }
  return out;
}

function buildParagraphStyleFormats(
  pkg: Package,
  indices: PackageIndices,
  theme: EditorTheme,
): Map<string, EditorParagraphFormat> {
  const styleData = new Map<string, EditorParagraphStyleData>();
  for (const part of pkg.parts.values()) {
    if (
      part.contentType !==
        'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' ||
      !part.xmlDom
    ) {
      continue;
    }
    const styleEls = part.xmlDom.getElementsByTagNameNS(W_NS, 'style');
    for (let i = 0; i < styleEls.length; i++) {
      const styleEl = styleEls[i];
      if (!styleEl) continue;
      const styleId = styleEl.getAttributeNS(W_NS, 'styleId');
      if (!styleId) continue;
      const type = styleEl.getAttributeNS(W_NS, 'type') ?? 'paragraph';
      if (type !== 'paragraph') continue;

      const data: EditorParagraphStyleData = {};
      const basedOn = firstDirectChild(styleEl, W_NS, 'basedOn')?.getAttributeNS(W_NS, 'val');
      if (basedOn) data.basedOn = basedOn;
      const format = parseStyleParagraphFormat(styleEl, theme);
      if (format) data.format = format;
      styleData.set(styleId, data);
    }
  }

  const resolved = new Map<string, EditorParagraphFormat>();
  for (const styleId of styleData.keys()) {
    const format = resolveParagraphStyleFormat(styleId, styleData, new Set());
    if (format) resolved.set(styleId, format);
  }
  for (const style of indices.stylesIdx.values()) {
    if (style.outlineLevel === undefined) continue;
    const existing = resolved.get(style.styleId);
    resolved.set(style.styleId, mergeParagraphFormats(existing, { outlineLevel: style.outlineLevel })!);
  }
  return resolved;
}

function findDefaultParagraphStyleId(pkg: Package): string | undefined {
  const stylesDom = findStylesDom(pkg);
  if (!stylesDom) return undefined;
  const styleEls = stylesDom.getElementsByTagNameNS(W_NS, 'style');
  for (let i = 0; i < styleEls.length; i++) {
    const styleEl = styleEls[i];
    if (!styleEl) continue;
    const type = styleEl.getAttributeNS(W_NS, 'type') ?? 'paragraph';
    if (type !== 'paragraph') continue;
    if (styleEl.getAttributeNS(W_NS, 'default') !== '1') continue;
    const styleId = styleEl.getAttributeNS(W_NS, 'styleId');
    if (styleId) return styleId;
  }
  return undefined;
}

function buildTableStyleFormats(pkg: Package, theme: EditorTheme): Map<string, EditorTableStyleFormat> {
  const styleData = new Map<string, EditorTableStyleData>();
  const stylesDom = findStylesDom(pkg);
  if (!stylesDom) return new Map();

  const styleEls = stylesDom.getElementsByTagNameNS(W_NS, 'style');
  for (let i = 0; i < styleEls.length; i++) {
    const styleEl = styleEls[i];
    if (!styleEl) continue;
    const styleId = styleEl.getAttributeNS(W_NS, 'styleId');
    if (!styleId) continue;
    const type = styleEl.getAttributeNS(W_NS, 'type') ?? 'paragraph';
    if (type !== 'table') continue;

    const data: EditorTableStyleData = {};
    const basedOn = firstDirectChild(styleEl, W_NS, 'basedOn')?.getAttributeNS(W_NS, 'val');
    if (basedOn) data.basedOn = basedOn;
    const format = parseTableStyleFormat(styleEl, theme);
    if (format) data.format = format;
    styleData.set(styleId, data);
  }

  const resolved = new Map<string, EditorTableStyleFormat>();
  for (const styleId of styleData.keys()) {
    const format = resolveTableStyleFormat(styleId, styleData, new Set());
    if (format) resolved.set(styleId, format);
  }
  return resolved;
}

function resolveTableStyleFormat(
  styleId: string,
  styleData: Map<string, EditorTableStyleData>,
  seen: Set<string>,
): EditorTableStyleFormat | undefined {
  if (seen.has(styleId)) return undefined;
  seen.add(styleId);
  const data = styleData.get(styleId);
  if (!data) return undefined;
  const base = data.basedOn ? resolveTableStyleFormat(data.basedOn, styleData, seen) : undefined;
  return mergeTableStyleFormats(base, data.format);
}

function parseTableStyleFormat(styleEl: Element, theme: EditorTheme): EditorTableStyleFormat | undefined {
  const wholeTable = mergeTableStyleFormats(
    parseTableStyleProperties(firstDirectChild(styleEl, W_NS, 'tblPr'), firstDirectChild(styleEl, W_NS, 'tcPr'), theme),
    parseWholeTableConditionalStyle(styleEl, theme),
  );
  const conditions = parseTableConditionalStyles(styleEl, theme);
  if (!conditions) return wholeTable;
  return mergeTableStyleFormats(wholeTable, { conditions });
}

function parseWholeTableConditionalStyle(styleEl: Element, theme: EditorTheme): EditorTableStyleFormat | undefined {
  for (const child of directChildren(styleEl, W_NS, 'tblStylePr')) {
    if (child.getAttributeNS(W_NS, 'type') !== 'wholeTable') continue;
    return parseTableStyleProperties(firstDirectChild(child, W_NS, 'tblPr'), firstDirectChild(child, W_NS, 'tcPr'), theme);
  }
  return undefined;
}

function parseTableConditionalStyles(
  styleEl: Element,
  theme: EditorTheme,
): EditorTableStyleFormat['conditions'] | undefined {
  const conditions: EditorTableStyleFormat['conditions'] = {};
  for (const child of directChildren(styleEl, W_NS, 'tblStylePr')) {
    const type = child.getAttributeNS(W_NS, 'type');
    if (
      type !== 'firstRow' &&
      type !== 'lastRow' &&
      type !== 'firstCol' &&
      type !== 'lastCol' &&
      type !== 'band1Horz' &&
      type !== 'band2Horz' &&
      type !== 'band1Vert' &&
      type !== 'band2Vert' &&
      type !== 'nwCell' &&
      type !== 'neCell' &&
      type !== 'swCell' &&
      type !== 'seCell'
    ) {
      continue;
    }
    const style = parseTableConditionalStyle(child, theme);
    if (style) conditions[type] = style;
  }
  return Object.keys(conditions).length > 0 ? conditions : undefined;
}

function parseTableConditionalStyle(styleEl: Element, theme: EditorTheme): EditorTableConditionalStyle | undefined {
  const style: EditorTableConditionalStyle = {};
  const cell = parseTableCellProperties(firstDirectChild(styleEl, W_NS, 'tcPr'), theme);
  if (cell) style.cell = cell;
  const paragraphFormat = mergeParagraphFormats(
    parseParagraphProperties(firstDirectChild(styleEl, W_NS, 'pPr'), theme),
    parseRunDefaultsAsParagraphFormat(firstDirectChild(styleEl, W_NS, 'rPr'), theme),
  );
  if (paragraphFormat) style.paragraphFormat = paragraphFormat;
  return Object.keys(style).length > 0 ? style : undefined;
}

function parseTableStyleProperties(
  tblPr: Element | null,
  tcPr: Element | null,
  theme: EditorTheme,
): EditorTableStyleFormat | undefined {
  const format: EditorTableStyleFormat = {};
  if (tblPr) {
    const borders = parseTableBordersFromPr(tblPr);
    if (borders) format.borders = borders;
    const cellMargin = parseMargins(firstDirectChild(tblPr, W_NS, 'tblCellMar'));
    if (cellMargin) format.cellMargin = cellMargin;
    const tableLook = parseTableLookFromPr(tblPr);
    if (tableLook) format.tableLook = tableLook;
  }
  if (tcPr) {
    const cell = parseTableCellProperties(tcPr, theme);
    if (cell) format.cell = cell;
  }
  return Object.keys(format).length > 0 ? format : undefined;
}

function parseTableCellProperties(
  tcPr: Element | null,
  theme?: EditorTheme,
): Pick<EditorTableCellView, 'margin' | 'borders' | 'shading' | 'verticalAlignment' | 'noWrap'> | undefined {
  if (!tcPr) return undefined;
  const cell: Pick<EditorTableCellView, 'margin' | 'borders' | 'shading' | 'verticalAlignment' | 'noWrap'> = {};
  const margin = parseMargins(firstDirectChild(tcPr, W_NS, 'tcMar'));
  if (margin) cell.margin = margin;
  const borders = parseCellBordersFromPr(tcPr);
  if (borders) cell.borders = borders;
  const shading = parseShading(firstDirectChild(tcPr, W_NS, 'shd'), theme);
  if (shading) cell.shading = shading;
  const verticalAlignment = parseVerticalAlignment(firstDirectChild(tcPr, W_NS, 'vAlign'));
  if (verticalAlignment) cell.verticalAlignment = verticalAlignment;
  const noWrap = parseOnOff(firstDirectChild(tcPr, W_NS, 'noWrap'));
  if (noWrap !== undefined ? noWrap : firstDirectChild(tcPr, W_NS, 'noWrap') !== null) cell.noWrap = true;
  return Object.keys(cell).length > 0 ? cell : undefined;
}

function mergeTableStyleFormats(
  base: EditorTableStyleFormat | undefined,
  override: EditorTableStyleFormat | undefined,
): EditorTableStyleFormat | undefined {
  if (!base && !override) return undefined;
  const merged: EditorTableStyleFormat = { ...(base ?? {}), ...(override ?? {}) };
  if (base?.borders || override?.borders) {
    merged.borders = { ...(base?.borders ?? {}), ...(override?.borders ?? {}) };
  }
  if (base?.cellMargin || override?.cellMargin) {
    merged.cellMargin = { ...(base?.cellMargin ?? {}), ...(override?.cellMargin ?? {}) };
  }
  if (base?.tableLook || override?.tableLook) {
    merged.tableLook = mergeTableLooks(base?.tableLook, override?.tableLook);
  }
  if (base?.cell || override?.cell) {
    merged.cell = { ...(base?.cell ?? {}), ...(override?.cell ?? {}) };
    if (base?.cell?.margin || override?.cell?.margin) {
      merged.cell.margin = { ...(base?.cell?.margin ?? {}), ...(override?.cell?.margin ?? {}) };
    }
    if (base?.cell?.borders || override?.cell?.borders) {
      merged.cell.borders = { ...(base?.cell?.borders ?? {}), ...(override?.cell?.borders ?? {}) };
    }
  }
  if (base?.paragraphFormat || override?.paragraphFormat) {
    merged.paragraphFormat = mergeParagraphFormats(base?.paragraphFormat, override?.paragraphFormat);
  }
  if (base?.conditions || override?.conditions) {
    merged.conditions = { ...(base?.conditions ?? {}) };
    for (const [condition, style] of Object.entries(override?.conditions ?? {}) as Array<
      [EditorTableStyleCondition, EditorTableConditionalStyle]
    >) {
      merged.conditions[condition] = mergeTableConditionalStyles(merged.conditions[condition], style);
    }
  }
  return merged;
}

function mergeTableConditionalStyles(
  base: EditorTableConditionalStyle | undefined,
  override: EditorTableConditionalStyle | undefined,
): EditorTableConditionalStyle | undefined {
  if (!base && !override) return undefined;
  const merged: EditorTableConditionalStyle = { ...(base ?? {}), ...(override ?? {}) };
  if (base?.cell || override?.cell) {
    merged.cell = { ...(base?.cell ?? {}), ...(override?.cell ?? {}) };
    if (base?.cell?.margin || override?.cell?.margin) {
      merged.cell.margin = { ...(base?.cell?.margin ?? {}), ...(override?.cell?.margin ?? {}) };
    }
    if (base?.cell?.borders || override?.cell?.borders) {
      merged.cell.borders = { ...(base?.cell?.borders ?? {}), ...(override?.cell?.borders ?? {}) };
    }
  }
  if (base?.paragraphFormat || override?.paragraphFormat) {
    merged.paragraphFormat = mergeParagraphFormats(base?.paragraphFormat, override?.paragraphFormat);
  }
  return merged;
}

function resolveParagraphStyleFormat(
  styleId: string,
  styleData: Map<string, EditorParagraphStyleData>,
  seen: Set<string>,
): EditorParagraphFormat | undefined {
  if (seen.has(styleId)) return undefined;
  seen.add(styleId);
  const data = styleData.get(styleId);
  if (!data) return undefined;
  const base = data.basedOn
    ? resolveParagraphStyleFormat(data.basedOn, styleData, seen)
    : undefined;
  return mergeParagraphFormats(base, data.format);
}

function parseStyleParagraphFormat(styleEl: Element, theme: EditorTheme): EditorParagraphFormat | undefined {
  const fromPPr = parseParagraphProperties(firstDirectChild(styleEl, W_NS, 'pPr'), theme);
  const fromRPr = parseRunDefaultsAsParagraphFormat(firstDirectChild(styleEl, W_NS, 'rPr'), theme);
  return mergeParagraphFormats(fromPPr, fromRPr);
}

function parseParagraphProperties(pPr: Element | null, theme?: EditorTheme): EditorParagraphFormat | undefined {
  const fmt: EditorParagraphFormat = {};
  let hasAny = false;
  if (pPr) {
    const jc = firstDirectChild(pPr, W_NS, 'jc');
    const alignment = jc?.getAttributeNS(W_NS, 'val');
    if (
      alignment === 'left' ||
      alignment === 'center' ||
      alignment === 'right' ||
      alignment === 'both' ||
      alignment === 'distribute'
    ) {
      fmt.alignment = alignment;
      hasAny = true;
    }

    const ind = firstDirectChild(pPr, W_NS, 'ind');
    if (ind) {
      const left = numAttr(ind, 'left');
      const right = numAttr(ind, 'right');
      const firstLineRaw = numAttr(ind, 'firstLine');
      const hangingRaw = numAttr(ind, 'hanging');
      const firstLine =
        firstLineRaw !== undefined
          ? firstLineRaw
          : hangingRaw !== undefined
            ? -hangingRaw
            : undefined;
      if (left !== undefined || right !== undefined || firstLine !== undefined) {
        fmt.indent = {};
        if (left !== undefined) fmt.indent.left = left;
        if (right !== undefined) fmt.indent.right = right;
        if (firstLine !== undefined) fmt.indent.firstLine = firstLine;
        hasAny = true;
      }
    }

    const spacing = firstDirectChild(pPr, W_NS, 'spacing');
    if (spacing) {
      const before = numAttr(spacing, 'before');
      const after = numAttr(spacing, 'after');
      const line = numAttr(spacing, 'line');
      const lineRuleRaw = spacing.getAttributeNS(W_NS, 'lineRule');
      const lineRule =
        lineRuleRaw === 'auto' || lineRuleRaw === 'exact' || lineRuleRaw === 'atLeast'
          ? lineRuleRaw
          : undefined;
      if (
        before !== undefined ||
        after !== undefined ||
        line !== undefined ||
        lineRule !== undefined
      ) {
        fmt.spacing = {};
        if (before !== undefined) fmt.spacing.before = before;
        if (after !== undefined) fmt.spacing.after = after;
        if (line !== undefined) fmt.spacing.line = line;
        if (lineRule !== undefined) fmt.spacing.lineRule = lineRule;
        hasAny = true;
      }
    }

    const tabStops = parseTabStops(firstDirectChild(pPr, W_NS, 'tabs'));
    if (tabStops !== undefined) {
      fmt.tabStops = tabStops;
      hasAny = true;
    }

    const outlineLvl = firstDirectChild(pPr, W_NS, 'outlineLvl');
    const outlineLevel = outlineLvl ? numAttr(outlineLvl, 'val') : undefined;
    if (outlineLevel !== undefined && outlineLevel >= 0 && outlineLevel <= 9) {
      fmt.outlineLevel = outlineLevel;
      hasAny = true;
    }

    const numPr = firstDirectChild(pPr, W_NS, 'numPr');
    if (numPr) {
      const numId = firstDirectChild(numPr, W_NS, 'numId')?.getAttributeNS(W_NS, 'val');
      if (numId && numId !== '0') {
        fmt.hasNumbering = true;
        fmt.numbering = {
          numId,
          ilvl: firstDirectChild(numPr, W_NS, 'ilvl')
            ? numAttr(firstDirectChild(numPr, W_NS, 'ilvl')!, 'val') ?? 0
            : 0,
        };
        hasAny = true;
      }
    }

    // Paragraph-mark run properties do not provide the default character
    // color for the paragraph text. WPS/Word keeps this color on the mark;
    // treating it as text color turns documents with blue paragraph marks blue.
    if (applyRunDefaultsToFormat(firstDirectChild(pPr, W_NS, 'rPr'), fmt, theme, false)) hasAny = true;
    if (applyParagraphBordersToFormat(pPr, fmt)) hasAny = true;
    const shading = parseShading(firstDirectChild(pPr, W_NS, 'shd'), theme);
    if (shading) {
      fmt.shading = shading;
      hasAny = true;
    }

    const keepNext = parseOnOff(firstDirectChild(pPr, W_NS, 'keepNext'));
    if (keepNext !== undefined) {
      fmt.keepNext = keepNext;
      hasAny = true;
    }
    const keepLines = parseOnOff(firstDirectChild(pPr, W_NS, 'keepLines'));
    if (keepLines !== undefined) {
      fmt.keepLines = keepLines;
      hasAny = true;
    }
    const widowControl = parseOnOff(firstDirectChild(pPr, W_NS, 'widowControl'));
    if (widowControl !== undefined) {
      fmt.widowControl = widowControl;
      hasAny = true;
    }

    const snapToGrid = parseOnOff(firstDirectChild(pPr, W_NS, 'snapToGrid'));
    if (snapToGrid !== undefined) {
      fmt.snapToGrid = snapToGrid;
      hasAny = true;
    }

    const pageBreakBefore = parseOnOff(firstDirectChild(pPr, W_NS, 'pageBreakBefore'));
    if (pageBreakBefore !== undefined) {
      fmt.pageBreakBefore = pageBreakBefore;
      hasAny = true;
    }

    if (firstDirectChild(pPr, W_NS, 'sectPr')) {
      fmt.hasSectionBreak = true;
      hasAny = true;
    }
  }
  return hasAny ? fmt : undefined;
}

function parseTabStops(tabs: Element | null): EditorParagraphFormat['tabStops'] | undefined {
  if (!tabs) return undefined;
  const stops: NonNullable<EditorParagraphFormat['tabStops']> = [];
  for (const tab of directChildren(tabs, W_NS, 'tab')) {
    const alignmentRaw = tab.getAttributeNS(W_NS, 'val');
    const alignment =
      alignmentRaw === 'left' ||
      alignmentRaw === 'center' ||
      alignmentRaw === 'right' ||
      alignmentRaw === 'decimal' ||
      alignmentRaw === 'bar'
        ? alignmentRaw
        : undefined;
    const position = numAttr(tab, 'pos');
    if (!alignment || position === undefined) continue;
    const leaderRaw = tab.getAttributeNS(W_NS, 'leader');
    const leader =
      leaderRaw === 'none' ||
      leaderRaw === 'dot' ||
      leaderRaw === 'hyphen' ||
      leaderRaw === 'underscore' ||
      leaderRaw === 'heavy' ||
      leaderRaw === 'middleDot'
        ? leaderRaw
        : undefined;
    stops.push({ alignment, position, ...(leader ? { leader } : {}) });
  }
  return stops.length > 0 ? stops : undefined;
}

function parseRunDefaultsAsParagraphFormat(rPr: Element | null, theme?: EditorTheme): EditorParagraphFormat | undefined {
  const fmt: EditorParagraphFormat = {};
  return applyRunDefaultsToFormat(rPr, fmt, theme) ? fmt : undefined;
}

function applyRunDefaultsToFormat(
  rPr: Element | null,
  fmt: EditorParagraphFormat,
  theme?: EditorTheme,
  includeColor = true,
): boolean {
  if (!rPr) return false;
  let hasAny = false;
  const fontFamily = parseFontFamily(rPr, theme);
  if (fontFamily) {
    fmt.fontFamily = fontFamily;
    hasAny = true;
  }
  const sz = firstDirectChild(rPr, W_NS, 'sz');
  const size = sz ? numAttr(sz, 'val') : undefined;
  if (size !== undefined) {
    fmt.fontSize = size / 2;
    hasAny = true;
  }
  const bold = parseOnOff(firstDirectChild(rPr, W_NS, 'b'));
  if (bold !== undefined) {
    fmt.bold = bold;
    hasAny = true;
  }
  if (includeColor) {
    const color = parseColor(firstDirectChild(rPr, W_NS, 'color'), theme);
    if (color) {
      fmt.color = color;
      hasAny = true;
    }
  }
  const characterSpacing = parseCharacterSpacing(rPr);
  if (characterSpacing !== undefined) {
    fmt.characterSpacing = characterSpacing;
    hasAny = true;
  }
  const highlight = parseHighlight(firstDirectChild(rPr, W_NS, 'highlight'));
  if (highlight) {
    fmt.highlight = highlight;
    hasAny = true;
  }
  const shading = parseShading(firstDirectChild(rPr, W_NS, 'shd'), theme);
  if (shading) {
    fmt.shading = shading;
    hasAny = true;
  }
  return hasAny;
}

function applyParagraphBordersToFormat(pPr: Element, fmt: EditorParagraphFormat): boolean {
  const pBdr = firstDirectChild(pPr, W_NS, 'pBdr');
  if (!pBdr) return false;
  const borders: NonNullable<EditorParagraphFormat['borders']> = {};
  for (const side of ['top', 'right', 'bottom', 'left', 'between', 'bar'] as const) {
    const border = parseParagraphBorder(firstDirectChild(pBdr, W_NS, side));
    if (border) borders[side] = border;
  }
  if (Object.keys(borders).length === 0) return false;
  fmt.borders = borders;
  if (borders.bottom) fmt.borderBottom = borders.bottom;
  return true;
}

function mergeParagraphFormats(
  base: EditorParagraphFormat | ParagraphFormat | undefined,
  override: EditorParagraphFormat | ParagraphFormat | undefined,
): EditorParagraphFormat | undefined {
  if (!base && !override) return undefined;
  const merged: EditorParagraphFormat = { ...(base ?? {}), ...(override ?? {}) };
  if (base?.indent || override?.indent) {
    merged.indent = { ...(base?.indent ?? {}), ...(override?.indent ?? {}) };
  }
  if (base?.spacing || override?.spacing) {
    merged.spacing = { ...(base?.spacing ?? {}), ...(override?.spacing ?? {}) };
  }
  if (
    (base && 'borders' in base && base.borders) ||
    (override && 'borders' in override && override.borders)
  ) {
    merged.borders = {
      ...((base as EditorParagraphFormat | undefined)?.borders ?? {}),
      ...((override as EditorParagraphFormat | undefined)?.borders ?? {}),
    };
    if (merged.borders.bottom) merged.borderBottom = merged.borders.bottom;
  }
  if (override && 'tabStops' in override && override.tabStops !== undefined) {
    merged.tabStops = override.tabStops.map((stop) => ({ ...stop }));
  } else if (base && 'tabStops' in base && base.tabStops !== undefined) {
    merged.tabStops = base.tabStops.map((stop) => ({ ...stop }));
  }
  return merged;
}

function numAttr(el: Element, localName: string): number | undefined {
  const raw = el.getAttributeNS(W_NS, localName);
  if (raw === null) return undefined;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function parseParagraphBorder(
  borderEl: Element | null,
): NonNullable<EditorParagraphFormat['borders']>['bottom'] | undefined {
  if (!borderEl) return undefined;
  const val = borderEl.getAttributeNS(W_NS, 'val') ?? undefined;
  if (val === 'none' || val === 'nil') return undefined;
  const border: NonNullable<EditorParagraphFormat['borders']>['bottom'] = { val };
  const color = borderEl.getAttributeNS(W_NS, 'color');
  const size = numAttr(borderEl, 'sz');
  const space = numAttr(borderEl, 'space');
  if (color) border.color = color;
  if (size !== undefined) border.size = size;
  if (space !== undefined) border.space = space;
  return border;
}

function parseRunFormat(run: Element, theme?: EditorTheme): EditorRunFormat | undefined {
  const rPr = firstDirectChild(run, W_NS, 'rPr');
  if (!rPr) return undefined;
  const parsed = parseRunProperties(rPr, theme);
  return parsed && Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseRunProperties(rPr: Element, theme?: EditorTheme): EditorRunFormat | undefined {
  const format: EditorRunFormat = {};

  const fontFamily = parseFontFamily(rPr, theme);
  if (fontFamily) format.fontFamily = fontFamily;

  const bold = parseOnOff(firstDirectChild(rPr, W_NS, 'b'));
  if (bold !== undefined) format.bold = bold;
  const italic = parseOnOff(firstDirectChild(rPr, W_NS, 'i'));
  if (italic !== undefined) format.italic = italic;
  const underline = firstDirectChild(rPr, W_NS, 'u');
  if (underline) {
    const val = underline.getAttributeNS(W_NS, 'val');
    format.underline = val !== 'none' && val !== '0' && val !== 'false';
  }
  const strike = parseOnOff(firstDirectChild(rPr, W_NS, 'strike'));
  if (strike !== undefined) format.strike = strike;
  const doubleStrike = parseOnOff(firstDirectChild(rPr, W_NS, 'dstrike'));
  if (doubleStrike !== undefined) format.doubleStrike = doubleStrike;
  const caps = parseOnOff(firstDirectChild(rPr, W_NS, 'caps'));
  if (caps !== undefined) format.caps = caps;
  const smallCaps = parseOnOff(firstDirectChild(rPr, W_NS, 'smallCaps'));
  if (smallCaps !== undefined) format.smallCaps = smallCaps;
  const sz = firstDirectChild(rPr, W_NS, 'sz');
  if (sz) {
    const raw = sz.getAttributeNS(W_NS, 'val');
    const parsed = raw !== null ? Number(raw) : NaN;
    if (Number.isFinite(parsed)) format.fontSize = parsed / 2;
  }
  const color = firstDirectChild(rPr, W_NS, 'color');
  if (color) {
    const val = parseColor(color, theme);
    if (val) format.color = val;
  }
  const characterSpacing = parseCharacterSpacing(rPr);
  if (characterSpacing !== undefined) format.characterSpacing = characterSpacing;
  const verticalAlign = parseVerticalAlign(firstDirectChild(rPr, W_NS, 'vertAlign'));
  if (verticalAlign) format.verticalAlign = verticalAlign;
  const highlight = parseHighlight(firstDirectChild(rPr, W_NS, 'highlight'));
  if (highlight) format.highlight = highlight;
  const shading = parseShading(firstDirectChild(rPr, W_NS, 'shd'), theme);
  if (shading) format.shading = shading;

  return Object.keys(format).length > 0 ? format : undefined;
}

function parseHighlight(highlight: Element | null): string | undefined {
  const val = highlight?.getAttributeNS(W_NS, 'val') ?? undefined;
  return val && val !== 'none' ? val : undefined;
}

function parseCharacterSpacing(rPr: Element): number | undefined {
  const spacing = firstDirectChild(rPr, W_NS, 'spacing');
  return spacing ? numAttr(spacing, 'val') : undefined;
}

function parseVerticalAlign(vertAlign: Element | null): EditorRunFormat['verticalAlign'] | undefined {
  const val = vertAlign?.getAttributeNS(W_NS, 'val');
  return val === 'baseline' || val === 'superscript' || val === 'subscript' ? val : undefined;
}

function parseFontFamily(rPr: Element, theme?: EditorTheme): string | undefined {
  const rFonts = firstDirectChild(rPr, W_NS, 'rFonts');
  if (!rFonts) return undefined;
  const eastAsia = rFonts.getAttributeNS(W_NS, 'eastAsia') || themeFont(theme, rFonts.getAttributeNS(W_NS, 'eastAsiaTheme'));
  const latin =
    rFonts.getAttributeNS(W_NS, 'ascii') ||
    rFonts.getAttributeNS(W_NS, 'hAnsi') ||
    themeFont(theme, rFonts.getAttributeNS(W_NS, 'asciiTheme')) ||
    themeFont(theme, rFonts.getAttributeNS(W_NS, 'hAnsiTheme'));
  if (latin && eastAsia && shouldPreserveLatinFont(latin)) {
    return latin === eastAsia ? latin : `${latin}, ${eastAsia}`;
  }
  return (
    eastAsia ||
    latin ||
    rFonts.getAttributeNS(W_NS, 'cs') ||
    themeFont(theme, rFonts.getAttributeNS(W_NS, 'csTheme')) ||
    undefined
  );
}

function shouldPreserveLatinFont(fontFamily: string): boolean {
  const normalized = fontFamily.trim().toLowerCase();
  return normalized === 'calibri' || normalized === 'aptos' || normalized === 'arial';
}

function parseColor(el: Element | null, theme?: EditorTheme): string | undefined {
  if (!el) return undefined;
  const color = normalizeRgbColor(el.getAttributeNS(W_NS, 'val'));
  const themed = color ?? themeColor(theme, el.getAttributeNS(W_NS, 'themeColor'));
  if (!themed) return undefined;
  return applyThemeTintOrShade(
    themed,
    el.getAttributeNS(W_NS, 'themeTint'),
    el.getAttributeNS(W_NS, 'themeShade'),
  );
}

function themeColor(theme: EditorTheme | undefined, value: string | null): string | undefined {
  if (!theme || !value) return undefined;
  return theme.colors.get(value);
}

function themeFont(theme: EditorTheme | undefined, value: string | null): string | undefined {
  if (!theme || !value) return undefined;
  return theme.fonts.get(normalizeThemeFontSlot(value));
}

function normalizeThemeFontSlot(value: string): string {
  const aliases: Record<string, string> = {
    '+mj-lt': 'majorHAnsi',
    '+mn-lt': 'minorHAnsi',
    '+mj-ea': 'majorEastAsia',
    '+mn-ea': 'minorEastAsia',
    '+mj-cs': 'majorBidi',
    '+mn-cs': 'minorBidi',
    majorAscii: 'majorHAnsi',
    minorAscii: 'minorHAnsi',
    majorLatin: 'majorHAnsi',
    minorLatin: 'minorHAnsi',
    majorCs: 'majorBidi',
    minorCs: 'minorBidi',
  };
  return aliases[value] ?? value;
}

function normalizeRgbColor(value: string | null | undefined): string | undefined {
  if (!value || value === 'auto') return undefined;
  if (/^[0-9a-f]{6}$/i.test(value)) return value.toUpperCase();
  if (/^[0-9a-f]{8}$/i.test(value)) return value.slice(2).toUpperCase();
  return undefined;
}

function applyThemeTintOrShade(color: string, tint: string | null, shade: string | null): string {
  const tintValue = parseHexByte(tint);
  if (tintValue !== undefined) return tintColor(color, tintValue / 255);
  const shadeValue = parseHexByte(shade);
  if (shadeValue !== undefined) return shadeColor(color, shadeValue / 255);
  return color;
}

function parseHexByte(value: string | null): number | undefined {
  if (!value || !/^[0-9a-f]{2}$/i.test(value)) return undefined;
  const parsed = Number.parseInt(value, 16);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function tintColor(color: string, amount: number): string {
  return transformRgb(color, (channel) => channel + (255 - channel) * amount);
}

function shadeColor(color: string, amount: number): string {
  return transformRgb(color, (channel) => channel * amount);
}

function transformRgb(color: string, transform: (channel: number) => number): string {
  const channels = [0, 2, 4].map((offset) => Number.parseInt(color.slice(offset, offset + 2), 16));
  return channels
    .map((channel) => Math.max(0, Math.min(255, Math.round(transform(channel)))))
    .map((channel) => channel.toString(16).toUpperCase().padStart(2, '0'))
    .join('');
}

function parseOnOff(el: Element | null): boolean | undefined {
  if (!el) return undefined;
  const val = el.getAttributeNS(W_NS, 'val');
  return parseOnOffValue(val);
}

function parseOnOffValue(value: string | null): boolean {
  return value !== '0' && value !== 'false' && value !== 'off';
}

function isShowingPlaceholderSdt(el: Element): boolean {
  if (el.namespaceURI !== W_NS || el.localName !== 'sdt') return false;
  const sdtPr = firstDirectChild(el, W_NS, 'sdtPr');
  return sdtPr ? parseOnOff(firstDirectChild(sdtPr, W_NS, 'showingPlcHdr')) === true : false;
}

function firstDirectChild(el: Element, ns: string, local: string): Element | null {
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === ns &&
      (child as Element).localName === local
    ) {
      return child as Element;
    }
  }
  return null;
}

function lastDirectChild(el: Element, ns: string, local: string): Element | null {
  for (let i = el.childNodes.length - 1; i >= 0; i--) {
    const child = el.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === ns &&
      (child as Element).localName === local
    ) {
      return child as Element;
    }
  }
  return null;
}

function directChildren(el: Element, ns: string, local: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === ns &&
      (child as Element).localName === local
    ) {
      out.push(child as Element);
    }
  }
  return out;
}

function firstDirectDescendant(el: Element, ns: string, path: readonly string[]): Element | null {
  let current: Element | null = el;
  for (const local of path) {
    current = current ? firstDirectChild(current, ns, local) : null;
    if (!current) return null;
  }
  return current;
}

function buildCommentSpan(
  indices: PackageIndices,
  commentId: number,
  paraId: string,
  from: number,
  to: number,
  paragraphText: string,
): EditorCommentSpan | null {
  const entry = indices.commentIdx.get(commentId);
  if (!entry?.comment) return null;
  const body = collectText(entry.comment, 't');
  return {
    kind: 'comment',
    id: `comment:${commentId}`,
    commentId,
    paraId,
    from,
    to,
    text: paragraphText.slice(from, to),
    rangeKind: from === 0 && to === paragraphText.length ? 'paragraph' : 'text',
    author: entry.comment.getAttributeNS(W_NS, 'author') ?? '',
    date: entry.comment.getAttributeNS(W_NS, 'date') ?? '',
    body,
  };
}

function buildDetachedComments(
  indices: PackageIndices,
  attachedCommentIds: Set<number>,
): CommentInfo[] {
  const detached: CommentInfo[] = [];
  for (const [id, entry] of indices.commentIdx) {
    if (attachedCommentIds.has(id) || !entry.comment) continue;
    detached.push({
      id,
      author: entry.comment.getAttributeNS(W_NS, 'author') ?? '',
      date: entry.comment.getAttributeNS(W_NS, 'date') ?? '',
      text: collectText(entry.comment, 't'),
      rangeKind: 'complex',
    });
  }
  return detached;
}

function markerId(el: Element, localName: string): number | null {
  if (el.namespaceURI !== W_NS || el.localName !== localName) return null;
  const raw = el.getAttributeNS(W_NS, 'id');
  if (raw === null) return null;
  const id = Number(raw);
  return Number.isFinite(id) ? id : null;
}

function collectVisibleText(root: Element, suppressPlaceholderSdtContent = false): string {
  let out = '';

  function visit(node: Node): void {
    if (node.nodeType !== 1) return;
    const el = node as Element;
    if (el.namespaceURI !== W_NS) return;
    if (suppressPlaceholderSdtContent && isShowingPlaceholderSdt(el)) return;
    if (el.localName === 't') {
      out += el.textContent ?? '';
      return;
    }
    if (el.localName === 'tab' || el.localName === 'ptab') {
      out += '\t';
      return;
    }
    if (el.localName === 'br') {
      const type = el.getAttributeNS(W_NS, 'type');
      if (type === null || type === '' || type === 'textWrapping') out += '\n';
      return;
    }
    if (el.localName === 'delText') return;
    if (el.localName === 'commentReference') return;

    for (let i = 0; i < el.childNodes.length; i++) {
      const child = el.childNodes[i];
      if (child) visit(child);
    }
  }

  visit(root);
  return out;
}

function containsFieldControl(run: Element): boolean {
  for (let i = 0; i < run.childNodes.length; i++) {
    const child = run.childNodes[i];
    if (!child || child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI !== W_NS) continue;
    if (el.localName === 'fldChar' || el.localName === 'instrText') return true;
  }
  return false;
}

function visibleTextFromFieldRunChild(el: Element): string {
  if (el.localName === 't') return el.textContent ?? '';
  if (el.localName === 'tab' || el.localName === 'ptab') return '\t';
  if (el.localName === 'br') {
    const type = el.getAttributeNS(W_NS, 'type');
    return type === null || type === '' || type === 'textWrapping' ? '\n' : '';
  }
  if (el.localName === 'delText' || el.localName === 'commentReference' || el.localName === 'instrText') {
    return '';
  }
  return collectVisibleText(el);
}

function fieldKindFromInstruction(instruction: string): EditorFieldSegment['fieldKind'] | null {
  const normalized = instruction
    .replace(/\\\*\s+MERGEFORMAT/gi, '')
    .replace(/\\\*\s+CHARFORMAT/gi, '')
    .trim()
    .toUpperCase();
  const fieldName = normalized.split(/\s+/)[0] ?? '';
  if (fieldName === 'PAGE') return 'pageNumber';
  if (fieldName === 'NUMPAGES') return 'totalPages';
  if (fieldName === 'SECTIONPAGES') return 'sectionPages';
  return null;
}

function collectText(root: Element, localName: 't' | 'delText'): string {
  const nodes = root.getElementsByTagNameNS(W_NS, localName);
  let out = '';
  for (let i = 0; i < nodes.length; i++) {
    out += nodes[i]?.textContent ?? '';
  }
  return out;
}

function containsDirectDrawing(el: Element): boolean {
  return el.namespaceURI === W_NS && (el.localName === 'drawing' || el.localName === 'pict');
}

function containsDirectPageBreak(run: Element): boolean {
  for (let i = 0; i < run.childNodes.length; i++) {
    const child = run.childNodes[i];
    if (!child || child.nodeType !== 1) continue;
    const el = child as Element;
    if (
      el.namespaceURI === W_NS &&
      el.localName === 'br' &&
      el.getAttributeNS(W_NS, 'type') === 'page'
    ) {
      return true;
    }
  }
  return false;
}

function directDrawingChildren(el: Element): Element[] {
  const drawings: Element[] = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes[i];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName === 'drawing'
    ) {
      drawings.push(child as Element);
    }
  }
  return drawings;
}

function isMarkerOnlyRun(run: Element): boolean {
  for (let i = 0; i < run.childNodes.length; i++) {
    const child = run.childNodes[i];
    if (!child || child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI !== W_NS) return false;
    if (
      el.localName === 'rPr' ||
      el.localName === 'commentReference' ||
      el.localName === 'bookmarkStart' ||
      el.localName === 'bookmarkEnd' ||
      el.localName === 'commentRangeStart' ||
      el.localName === 'commentRangeEnd'
    ) {
      continue;
    }
    return false;
  }
  return true;
}

export type { RevisionInfo };
