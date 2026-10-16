import { getPart, isBinaryPart, type Package, type Part } from '../pack/index.js';
import type {
  PptChartModel,
  PptChartData,
  PptChartSeriesData,
  PptCommentModel,
  PptCustomGeometryCommand,
  PptCustomGeometryModel,
  PptGroupTransformModel,
  PptHyperlinkModel,
  PptFontModel,
  PptGradientFillModel,
  PptImageModel,
  PptMediaModel,
  PptPresentationModel,
  PptShapeGeometry,
  PptShapeKind,
  PptShapeMediaModel,
  PptShapeModel,
  PptShapeStyle,
  PptShadowModel,
  PptSlideBackgroundModel,
  PptSlideLayoutModel,
  PptSlideMasterModel,
  PptSlideModel,
  PptSlideTimingModel,
  PptSlideTransitionModel,
  PptOleObjectModel,
  PptSmartArtModel,
  PptTableCellStyle,
  PptTableCellMerge,
  PptTableModel,
  PptTextBodyProperties,
  PptThemeModel,
  PptTextParagraph,
  PptTextRun,
  PptTextStyle,
} from './model.js';
import {
  A_NS,
  OFFICE_REL_NS,
  P_NS,
  R_NS,
  allDescendantsLocal,
  allDescendantsNs,
  childElementsNs,
  firstChildLocal,
  firstChildNs,
  parseNumber,
  parseRelationships,
  relsPathForPart,
  resolveRelationshipTarget,
  textOf,
  type PptRelationship,
} from './xml.js';

const PRESENTATION_REL_TYPE = `${OFFICE_REL_NS}/officeDocument`;
const SLIDE_REL_TYPE = `${OFFICE_REL_NS}/slide`;
const SLIDE_LAYOUT_REL_TYPE = `${OFFICE_REL_NS}/slideLayout`;
const SLIDE_MASTER_REL_TYPE = `${OFFICE_REL_NS}/slideMaster`;
const THEME_REL_TYPE = `${OFFICE_REL_NS}/theme`;
const IMAGE_REL_TYPE = `${OFFICE_REL_NS}/image`;
const CHART_REL_TYPE = `${OFFICE_REL_NS}/chart`;
const FONT_REL_TYPE = `${OFFICE_REL_NS}/font`;
const MEDIA_REL_TYPE = `${OFFICE_REL_NS}/media`;
const VIDEO_REL_TYPE = `${OFFICE_REL_NS}/video`;
const AUDIO_REL_TYPE = `${OFFICE_REL_NS}/audio`;
const OLE_OBJECT_REL_TYPE = `${OFFICE_REL_NS}/oleObject`;
const NOTES_SLIDE_REL_TYPE = `${OFFICE_REL_NS}/notesSlide`;
const COMMENTS_REL_TYPE = `${OFFICE_REL_NS}/comments`;
const COMMENT_AUTHORS_REL_TYPE = `${OFFICE_REL_NS}/commentAuthors`;
const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

export interface PptShapeInternal {
  slide: PptSlideInternal;
  model: PptShapeModel;
  element: Element;
  textElements: Element[];
  tableCells: Element[][];
  children: PptShapeInternal[];
}

export interface PptSlideInternal {
  model: PptSlideModel;
  part: Part;
  relationships: Map<string, PptRelationship>;
  shapes: PptShapeInternal[];
}

export interface PptReadResult {
  model: PptPresentationModel;
  presentationPart: Part;
  slides: PptSlideInternal[];
}

interface PptParseContext {
  colorScheme?: Record<string, string>;
  fontScheme?: PptThemeModel['fontScheme'];
  relationships?: Map<string, PptRelationship>;
  sourcePath?: string;
}

export function readPresentationPackage(pkg: Package, presentationId?: string): PptReadResult {
  const presentationPath = findPresentationPath(pkg);
  const presentationPart = requireXmlPart(pkg, presentationPath);
  const presentationDoc = presentationPart.xmlDom!;
  const presentationRels = parseRelationships(getPart(pkg, relsPathForPart(presentationPath)));
  const commentAuthors = parseCommentAuthors(pkg, presentationPath, presentationRels);
  const slideSize = parseSlideSize(presentationDoc.documentElement);
  const slideIds = allDescendantsNs(presentationDoc, P_NS, 'sldId');
  const slides: PptSlideInternal[] = [];

  for (let index = 0; index < slideIds.length; index++) {
    const slideIdEl = slideIds[index]!;
    const relId = slideIdEl.getAttributeNS(R_NS, 'id') ?? slideIdEl.getAttribute('r:id');
    if (!relId) continue;
    const rel = presentationRels.get(relId);
    if (!rel || rel.type !== SLIDE_REL_TYPE) continue;
    const slidePath = resolveRelationshipTarget(presentationPath, rel.target);
    const slidePart = requireXmlPart(pkg, slidePath);
    const slideRels = parseRelationships(getPart(pkg, relsPathForPart(slidePath)));
    const slideModel: PptSlideModel = {
      id: slideIdEl.getAttribute('id') ?? relId,
      index,
      relationshipId: relId,
      path: slidePath,
      shapes: [],
    };
    const slide: PptSlideInternal = {
      model: slideModel,
      part: slidePart,
      relationships: slideRels,
      shapes: [],
    };
    slide.model.layout = parseSlideLayout(pkg, slidePath, slideRels);
    const slideContext = contextFromTheme(slide.model.layout?.master?.theme);
    slide.shapes = parseSlideShapes(pkg, slide, slideContext);
    slide.model.shapes = slide.shapes.map((shape) => shape.model);
    slide.model.title = deriveSlideTitle(slide.model.shapes);
    slide.model.background = parseSlideBackground(
      slidePart.xmlDom!,
      contextWithRelationships(slideContext, slidePath, slideRels),
    );
    slide.model.layoutId = slide.model.layout?.relationshipId;
    slide.model.notes = parseSlideNotes(pkg, slidePath, slideRels);
    slide.model.comments = parseSlideComments(pkg, slidePath, slideRels, commentAuthors);
    slide.model.transition = parseSlideTransition(slidePart.xmlDom!);
    slide.model.timing = parseSlideTiming(slidePart.xmlDom!);
    slides.push(slide);
  }

  const model: PptPresentationModel = {
    id: presentationId,
    slideSize,
    slides: slides.map((slide) => slide.model),
    media: listMedia(pkg),
    fonts: listEmbeddedFonts(pkg, presentationPath, presentationDoc, presentationRels),
    themes: collectThemes(pkg, presentationPath, presentationRels, slides),
  };
  return { model, presentationPart, slides };
}

function findPresentationPath(pkg: Package): string {
  const rootRels = parseRelationships(getPart(pkg, '_rels/.rels'));
  for (const rel of rootRels.values()) {
    if (rel.type === PRESENTATION_REL_TYPE) {
      return rel.target.startsWith('/') ? rel.target.slice(1) : rel.target;
    }
  }
  if (getPart(pkg, 'ppt/presentation.xml')) return 'ppt/presentation.xml';
  throw new Error('Presentation: missing ppt/presentation.xml');
}

function requireXmlPart(pkg: Package, path: string): Part {
  const part = getPart(pkg, path);
  if (!part?.xmlDom) throw new Error(`Presentation: missing XML part ${path}`);
  return part;
}

function contextFromTheme(theme: PptThemeModel | undefined): PptParseContext {
  return {
    colorScheme: theme?.colorScheme,
    fontScheme: theme?.fontScheme,
  };
}

function contextWithRelationships(
  context: PptParseContext,
  sourcePath: string,
  relationships: Map<string, PptRelationship>,
): PptParseContext {
  return {
    ...context,
    sourcePath,
    relationships,
  };
}

function parseSlideSize(presentationEl: Element): { width: number; height: number } | undefined {
  const sldSz = firstChildNs(presentationEl, P_NS, 'sldSz');
  const width = parseNumber(sldSz?.getAttribute('cx') ?? null);
  const height = parseNumber(sldSz?.getAttribute('cy') ?? null);
  return width !== undefined && height !== undefined ? { width, height } : undefined;
}

function parseSlideLayout(
  pkg: Package,
  slidePath: string,
  slideRels: Map<string, PptRelationship>,
): PptSlideLayoutModel | undefined {
  const rel = Array.from(slideRels.values()).find((item) => item.type === SLIDE_LAYOUT_REL_TYPE);
  if (!rel) return undefined;
  const path = resolveRelationshipTarget(slidePath, rel.target);
  const part = getPart(pkg, path);
  const model: PptSlideLayoutModel = {
    relationshipId: rel.id,
    path,
  };
  if (part?.xmlDom) {
    const cSld = allDescendantsNs(part.xmlDom, P_NS, 'cSld')[0];
    const name = cSld?.getAttribute('name') ?? undefined;
    if (name) model.name = name;
    const layoutRels = parseRelationships(getPart(pkg, relsPathForPart(path)));
    model.master = parseSlideMaster(pkg, path, layoutRels);
    const context = contextFromTheme(model.master?.theme);
    model.background = parseSlideBackground(
      part.xmlDom,
      contextWithRelationships(context, path, layoutRels),
    );
    const layoutShapes = parsePartShapes(
      pkg,
      {
        model: { id: model.relationshipId ?? path, index: 0, path, shapes: [] },
        part,
        relationships: layoutRels,
        shapes: [],
      },
      context,
    );
    if (layoutShapes.length > 0) model.shapes = layoutShapes.map((shape) => shape.model);
  }
  return model;
}

function parseSlideMaster(
  pkg: Package,
  layoutPath: string,
  layoutRels: Map<string, PptRelationship>,
): PptSlideMasterModel | undefined {
  const rel = Array.from(layoutRels.values()).find((item) => item.type === SLIDE_MASTER_REL_TYPE);
  if (!rel) return undefined;
  const path = resolveRelationshipTarget(layoutPath, rel.target);
  const part = getPart(pkg, path);
  const model: PptSlideMasterModel = {
    relationshipId: rel.id,
    path,
  };
  if (part?.xmlDom) {
    const cSld = allDescendantsNs(part.xmlDom, P_NS, 'cSld')[0];
    const name = cSld?.getAttribute('name') ?? undefined;
    if (name) model.name = name;
    const masterRels = parseRelationships(getPart(pkg, relsPathForPart(path)));
    const themeRel = Array.from(masterRels.values()).find((item) => item.type === THEME_REL_TYPE);
    if (themeRel) model.theme = parseTheme(pkg, path, themeRel);
    const context = contextFromTheme(model.theme);
    const textStyles = parseSlideMasterTextStyles(part.xmlDom, context);
    if (Object.keys(textStyles).length > 0) model.textStyles = textStyles;
    model.background = parseSlideBackground(
      part.xmlDom,
      contextWithRelationships(context, path, masterRels),
    );
    const masterShapes = parsePartShapes(
      pkg,
      {
        model: { id: model.relationshipId ?? path, index: 0, path, shapes: [] },
        part,
        relationships: masterRels,
        shapes: [],
      },
      context,
    );
    if (masterShapes.length > 0) model.shapes = masterShapes.map((shape) => shape.model);
  }
  return model;
}

function parseSlideMasterTextStyles(
  doc: Document,
  context: PptParseContext,
): NonNullable<PptSlideMasterModel['textStyles']> {
  const textStyles = allDescendantsNs(doc, P_NS, 'txStyles')[0];
  if (!textStyles) return {};
  const result: NonNullable<PptSlideMasterModel['textStyles']> = {};
  const entries = [
    ['title', firstChildNs(textStyles, P_NS, 'titleStyle')],
    ['body', firstChildNs(textStyles, P_NS, 'bodyStyle')],
    ['other', firstChildNs(textStyles, P_NS, 'otherStyle')],
  ] as const;
  for (const [key, element] of entries) {
    if (!element) continue;
    const style = parseTextStyle(element, context);
    if (Object.keys(style).length > 0) result[key] = style;
  }
  return result;
}

function collectThemes(
  pkg: Package,
  presentationPath: string,
  presentationRels: Map<string, PptRelationship>,
  slides: PptSlideInternal[],
): PptThemeModel[] | undefined {
  const byPath = new Map<string, PptThemeModel>();
  for (const rel of presentationRels.values()) {
    if (rel.type !== THEME_REL_TYPE) continue;
    const theme = parseTheme(pkg, presentationPath, rel);
    if (theme.path) byPath.set(theme.path, theme);
  }
  for (const slide of slides) {
    const theme = slide.model.layout?.master?.theme;
    if (theme?.path) byPath.set(theme.path, theme);
  }
  return byPath.size > 0 ? Array.from(byPath.values()) : undefined;
}

function parseTheme(pkg: Package, sourcePath: string, rel: PptRelationship): PptThemeModel {
  const path = resolveRelationshipTarget(sourcePath, rel.target);
  const part = getPart(pkg, path);
  const model: PptThemeModel = {
    relationshipId: rel.id,
    path,
  };
  if (!part?.xmlDom) return model;
  const theme = allDescendantsNs(part.xmlDom, A_NS, 'theme')[0];
  const name = theme?.getAttribute('name') ?? undefined;
  if (name) model.name = name;
  const colorScheme = parseThemeColorScheme(part.xmlDom);
  if (Object.keys(colorScheme).length > 0) model.colorScheme = colorScheme;
  const fontScheme = parseThemeFontScheme(part.xmlDom);
  if (Object.keys(fontScheme).length > 0) model.fontScheme = fontScheme;
  return model;
}

function parseThemeColorScheme(doc: Document): Record<string, string> {
  const scheme = allDescendantsNs(doc, A_NS, 'clrScheme')[0];
  if (!scheme) return {};
  const result: Record<string, string> = {};
  for (let i = 0; i < scheme.childNodes.length; i++) {
    const child = scheme.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const item = child as Element;
    const color =
      allDescendantsNs(item, A_NS, 'srgbClr')[0]?.getAttribute('val') ??
      allDescendantsNs(item, A_NS, 'sysClr')[0]?.getAttribute('lastClr');
    if (color) result[item.localName] = `#${color.toUpperCase()}`;
  }
  return result;
}

function parseThemeFontScheme(doc: Document): NonNullable<PptThemeModel['fontScheme']> {
  const scheme = allDescendantsNs(doc, A_NS, 'fontScheme')[0];
  if (!scheme) return {};
  const majorFont = firstChildNs(scheme, A_NS, 'majorFont');
  const minorFont = firstChildNs(scheme, A_NS, 'minorFont');
  return {
    ...(scheme.getAttribute('name') ? { name: scheme.getAttribute('name')! } : {}),
    ...(majorFont
      ? {
          majorLatin: firstChildNs(majorFont, A_NS, 'latin')?.getAttribute('typeface') ?? undefined,
          majorEa: firstChildNs(majorFont, A_NS, 'ea')?.getAttribute('typeface') ?? undefined,
          majorCs: firstChildNs(majorFont, A_NS, 'cs')?.getAttribute('typeface') ?? undefined,
        }
      : {}),
    ...(minorFont
      ? {
          minorLatin: firstChildNs(minorFont, A_NS, 'latin')?.getAttribute('typeface') ?? undefined,
          minorEa: firstChildNs(minorFont, A_NS, 'ea')?.getAttribute('typeface') ?? undefined,
          minorCs: firstChildNs(minorFont, A_NS, 'cs')?.getAttribute('typeface') ?? undefined,
        }
      : {}),
  };
}

function parseSlideNotes(
  pkg: Package,
  slidePath: string,
  slideRels: Map<string, PptRelationship>,
): PptSlideModel['notes'] | undefined {
  const rel = Array.from(slideRels.values()).find((item) => item.type === NOTES_SLIDE_REL_TYPE);
  if (!rel) return undefined;
  const part = getPart(pkg, resolveRelationshipTarget(slidePath, rel.target));
  if (!part?.xmlDom) return undefined;
  const text = allDescendantsNs(part.xmlDom, A_NS, 't').map(textOf).join('');
  return { text };
}

interface PptCommentAuthorInfo {
  name?: string;
  initials?: string;
}

function parseCommentAuthors(
  pkg: Package,
  presentationPath: string,
  presentationRels: Map<string, PptRelationship>,
): Map<string, PptCommentAuthorInfo> {
  const result = new Map<string, PptCommentAuthorInfo>();
  const rel = Array.from(presentationRels.values()).find(
    (item) => item.type === COMMENT_AUTHORS_REL_TYPE,
  );
  const authorPath = rel
    ? resolveRelationshipTarget(presentationPath, rel.target)
    : 'ppt/commentAuthors.xml';
  const part = getPart(pkg, authorPath);
  if (!part?.xmlDom) return result;
  for (const author of allDescendantsNs(part.xmlDom, P_NS, 'cmAuthor')) {
    const id = author.getAttribute('id');
    if (!id) continue;
    result.set(id, {
      name: author.getAttribute('name') ?? undefined,
      initials: author.getAttribute('initials') ?? undefined,
    });
  }
  return result;
}

function parseSlideComments(
  pkg: Package,
  slidePath: string,
  slideRels: Map<string, PptRelationship>,
  authors: Map<string, PptCommentAuthorInfo>,
): PptCommentModel[] | undefined {
  const rel = Array.from(slideRels.values()).find((item) => item.type === COMMENTS_REL_TYPE);
  if (!rel) return undefined;
  const part = getPart(pkg, resolveRelationshipTarget(slidePath, rel.target));
  if (!part?.xmlDom) return undefined;
  const comments = allDescendantsNs(part.xmlDom, P_NS, 'cm').map((comment): PptCommentModel => {
    const authorId = comment.getAttribute('authorId') ?? undefined;
    const author = authorId ? authors.get(authorId) : undefined;
    const pos = firstChildNs(comment, P_NS, 'pos');
    const x = parseNumber(pos?.getAttribute('x') ?? null);
    const y = parseNumber(pos?.getAttribute('y') ?? null);
    return {
      ...(comment.getAttribute('idx') ? { id: comment.getAttribute('idx')! } : {}),
      ...(authorId ? { authorId } : {}),
      ...(author?.name ? { author: author.name } : {}),
      ...(author?.initials ? { initials: author.initials } : {}),
      ...(comment.getAttribute('dt') ? { date: comment.getAttribute('dt')! } : {}),
      text: textOf(firstChildNs(comment, P_NS, 'text')),
      ...(x !== undefined ? { x } : {}),
      ...(y !== undefined ? { y } : {}),
    };
  });
  return comments.length > 0 ? comments : undefined;
}

function parseSlideTransition(doc: Document): PptSlideTransitionModel | undefined {
  const transition =
    firstChildNs(doc.documentElement, P_NS, 'transition') ??
    firstChildLocal(doc.documentElement, 'transition') ??
    allDescendantsLocal(doc, 'transition')[0];
  if (!transition) return undefined;
  const model: PptSlideTransitionModel = {};
  const attributes = attributesOf(transition);
  if (Object.keys(attributes).length > 0) model.attributes = attributes;
  const speed = transition.getAttribute('spd') ?? undefined;
  const duration =
    parseNumber(transition.getAttribute('dur')) ??
    parseNumber(transition.getAttribute('p14:dur')) ??
    parseNumber(
      transition.getAttributeNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'dur'),
    );
  const advanceOnClick = parseBooleanAttr(transition.getAttribute('advClick'));
  const advanceAfterTime = parseNumber(transition.getAttribute('advTm'));
  const transitionType = firstTransitionChildName(transition);
  if (speed) model.speed = speed;
  if (duration !== undefined) model.duration = duration;
  if (advanceOnClick !== undefined) model.advanceOnClick = advanceOnClick;
  if (advanceAfterTime !== undefined) model.advanceAfterTime = advanceAfterTime;
  if (transitionType) model.type = transitionType;
  return model;
}

function parseSlideTiming(doc: Document): PptSlideTimingModel | undefined {
  const timing =
    firstChildNs(doc.documentElement, P_NS, 'timing') ??
    firstChildLocal(doc.documentElement, 'timing') ??
    allDescendantsLocal(doc, 'timing')[0];
  if (!timing) return undefined;
  const timeNodes = allDescendantsLocal(timing, 'cTn');
  const sequences = allDescendantsLocal(timing, 'seq');
  const root = timeNodes.find((node) => node.getAttribute('nodeType') === 'tmRoot') ?? timeNodes[0];
  const model: PptSlideTimingModel = {
    hasTiming: true,
    timeNodeCount: timeNodes.length,
    mainSequenceCount: sequences.filter((seq) => seq.getAttribute('nodeType') === 'mainSeq').length,
    interactiveSequenceCount: sequences.filter(
      (seq) => seq.getAttribute('nodeType') === 'interactiveSeq',
    ).length,
  };
  const rootDuration = root?.getAttribute('dur') ?? undefined;
  const rootNodeType = root?.getAttribute('nodeType') ?? undefined;
  if (rootDuration) model.rootDuration = rootDuration;
  if (rootNodeType) model.rootNodeType = rootNodeType;
  return model;
}

function firstTransitionChildName(transition: Element): string | undefined {
  for (let i = 0; i < transition.childNodes.length; i++) {
    const child = transition.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const element = child as Element;
    if (element.localName === 'sndAc' || element.localName === 'extLst') continue;
    return element.localName;
  }
  return undefined;
}

function attributesOf(element: Element): Record<string, string> {
  const result: Record<string, string> = {};
  for (let i = 0; i < element.attributes.length; i++) {
    const attr = element.attributes[i];
    if (!attr) continue;
    result[attr.name] = attr.value;
  }
  return result;
}

function parseSlideShapes(
  pkg: Package,
  slide: PptSlideInternal,
  context: PptParseContext = {},
): PptShapeInternal[] {
  return parsePartShapes(pkg, slide, context);
}

function parsePartShapes(
  pkg: Package,
  slide: PptSlideInternal,
  context: PptParseContext = {},
): PptShapeInternal[] {
  const spTree = allDescendantsNs(slide.part.xmlDom!, P_NS, 'spTree')[0];
  if (!spTree) return [];
  return parseShapeChildren(pkg, slide, spTree, context);
}

function parseShapeChildren(
  pkg: Package,
  slide: PptSlideInternal,
  parent: Element,
  context: PptParseContext = {},
): PptShapeInternal[] {
  const result: PptShapeInternal[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI !== P_NS) continue;
    if (element.localName === 'sp')
      result.push(parseShapeElement(pkg, slide, element, 'sp', context));
    else if (element.localName === 'pic')
      result.push(parseShapeElement(pkg, slide, element, 'pic', context));
    else if (element.localName === 'graphicFrame')
      result.push(parseShapeElement(pkg, slide, element, 'graphicFrame', context));
    else if (element.localName === 'grpSp')
      result.push(parseShapeElement(pkg, slide, element, 'grpSp', context));
    else if (element.localName === 'cxnSp' && isRenderableConnectorElement(element))
      result.push(parseShapeElement(pkg, slide, element, 'sp', context));
  }
  return result;
}

function isRenderableConnectorElement(element: Element): boolean {
  const spPr = firstChildNs(element, P_NS, 'spPr');
  const xfrm = spPr ? firstChildNs(spPr, A_NS, 'xfrm') : null;
  const ext = xfrm ? firstChildNs(xfrm, A_NS, 'ext') : null;
  const width = parseNumber(ext?.getAttribute('cx') ?? null) ?? 0;
  const height = parseNumber(ext?.getAttribute('cy') ?? null) ?? 0;
  const preset = parsePresetGeometry(element);
  const line = spPr ? firstChildNs(spPr, A_NS, 'ln') : null;
  const hasNoFill = !!(line && firstChildNs(line, A_NS, 'noFill'));
  const hasLineFill = !!(
    line &&
    (firstChildNs(line, A_NS, 'solidFill') || firstChildNs(line, A_NS, 'gradFill'))
  );
  const min = Math.min(width, height);
  const max = Math.max(width, height);
  const isAxisAligned = min === 0 || (max > 0 && min / max < 0.02);
  return (
    (preset === 'straightConnector1' || preset === 'line') &&
    isAxisAligned &&
    hasLineFill &&
    !hasNoFill
  );
}

function parseShapeElement(
  pkg: Package,
  slide: PptSlideInternal,
  element: Element,
  hostType: 'sp' | 'pic' | 'graphicFrame' | 'grpSp',
  context: PptParseContext = {},
): PptShapeInternal {
  const cNvPr = allDescendantsNs(element, P_NS, 'cNvPr')[0];
  const id = cNvPr?.getAttribute('id') ?? `${slide.model.id}:${slide.shapes.length}`;
  const name = cNvPr?.getAttribute('name') ?? undefined;
  const children = hostType === 'grpSp' ? parseShapeChildren(pkg, slide, element, context) : [];
  const textElements = hostType === 'grpSp' ? [] : allDescendantsNs(element, A_NS, 't');
  const tableCells = hostType === 'grpSp' ? [] : parseTableCells(element);
  const table =
    tableCells.length > 0 ? tableModelFromCells(element, tableCells, context) : undefined;
  const image = imageModelFromElement(slide, element);
  const chart = parseChartModel(pkg, slide, element, context);
  const media = parseMediaModel(pkg, slide, element);
  const oleObject = parseOleObjectModel(slide, element);
  const smartArt = parseSmartArtModel(slide, element);
  const kind = classifyShape(slide, element, hostType, textElements, {
    table,
    image,
    chart,
    media,
    oleObject,
    smartArt,
  });
  const textStyle = parseTextStyle(element, context);
  const textParagraphs =
    textElements.length > 0 ? parseTextParagraphs(slide, element, context) : undefined;
  const text =
    textElements.length > 0
      ? {
          text: textParagraphsToPlainText(textParagraphs ?? []),
          paragraphs: textParagraphs ?? [],
          ...(Object.keys(textStyle).length > 0 ? { style: textStyle } : {}),
          ...(parseTextBodyProperties(element) ? { body: parseTextBodyProperties(element) } : {}),
        }
      : undefined;
  const style = parseShapeStyle(element, context, {
    useStyleFillRef: !image && hostType !== 'pic',
  });
  const hyperlink = parseShapeHyperlink(slide, cNvPr);
  const placeholder = parsePlaceholder(element);
  const shapeProperties = firstChildNs(element, P_NS, 'spPr') ?? firstChildNs(element, P_NS, 'grpSpPr');
  const model: PptShapeModel = {
    id,
    kind,
    geometry: parseGeometry(element),
    raw: {
      slidePath: slide.model.path,
      relationshipId: image?.relationshipId,
      hostType,
      presetGeometry: parsePresetGeometry(element),
      presetAdjustments: parsePresetGeometryAdjustments(element),
      customGeometry: parseCustomGeometry(element),
      threeD: shapeProperties ? parseShape3d(shapeProperties) : undefined,
      groupTransform: hostType === 'grpSp' ? parseGroupTransform(element) : undefined,
      useBackgroundFill: parseBooleanAttr(element.getAttribute('useBgFill')) === true,
    },
  };
  if (name) model.name = name;
  if (Object.keys(style).length > 0) model.style = style;
  if (text) model.text = text;
  if (image) model.image = image;
  if (chart) model.chart = chart;
  if (media) model.media = media;
  if (oleObject) model.oleObject = oleObject;
  if (smartArt) model.smartArt = smartArt;
  if (table) model.table = table;
  if (placeholder) model.placeholder = placeholder;
  if (hyperlink) model.hyperlink = hyperlink;
  if (children.length > 0)
    model.children = children.map((child) => inheritGroupFill(child.model, child.element, style));
  return { slide, model, element, textElements, tableCells, children };
}

function parseShapeStyle(
  element: Element,
  context: PptParseContext = {},
  options: { useStyleFillRef?: boolean } = {},
): PptShapeStyle {
  const spPr = firstChildNs(element, P_NS, 'spPr') ?? firstChildNs(element, P_NS, 'grpSpPr');
  if (!spPr) return {};
  const solidFill = parseSolidFill(firstChildNs(spPr, A_NS, 'solidFill'), context);
  const fillGradient = parseGradientFill(firstChildNs(spPr, A_NS, 'gradFill'), context);
  const styleFillColor =
    !options.useStyleFillRef ||
    solidFill.color ||
    fillGradient ||
    firstChildNs(spPr, A_NS, 'noFill')
      ? undefined
      : colorFromStyleFillRef(element, context);
  const ln = firstChildNs(spPr, A_NS, 'ln');
  const lineFill = ln ? parseSolidFill(firstChildNs(ln, A_NS, 'solidFill'), context) : {};
  const lineGradient = ln
    ? parseGradientFill(firstChildNs(ln, A_NS, 'gradFill'), context)
    : undefined;
  const lineColor = lineFill.color;
  const lineWidth = parseNumber(ln?.getAttribute('w') ?? null);
  const lineCap = ln?.getAttribute('cap') ?? undefined;
  const lineDash = ln ? parseLineDash(ln) : undefined;
  const lineHeadEnd = ln
    ? (firstChildNs(ln, A_NS, 'headEnd')?.getAttribute('type') ?? undefined)
    : undefined;
  const lineTailEnd = ln
    ? (firstChildNs(ln, A_NS, 'tailEnd')?.getAttribute('type') ?? undefined)
    : undefined;
  const shadow = parseShapeShadow(spPr, context);
  return {
    ...((solidFill.color ?? styleFillColor)
      ? { fillColor: solidFill.color ?? styleFillColor }
      : {}),
    ...(fillGradient ? { fillGradient } : {}),
    ...(solidFill.opacity !== undefined ? { opacity: solidFill.opacity } : {}),
    ...(lineColor ? { lineColor } : {}),
    ...(lineGradient ? { lineGradient } : {}),
    ...(lineFill.opacity !== undefined ? { lineOpacity: lineFill.opacity } : {}),
    ...(lineWidth !== undefined ? { lineWidth } : {}),
    ...(lineCap ? { lineCap } : {}),
    ...(lineDash && lineDash !== 'solid' ? { lineDash } : {}),
    ...(lineHeadEnd ? { lineHeadEnd } : {}),
    ...(lineTailEnd ? { lineTailEnd } : {}),
    ...(shadow ? { shadow } : {}),
  };
}

function inheritGroupFill(
  model: PptShapeModel,
  element: Element,
  groupStyle: PptShapeStyle,
): PptShapeModel {
  if (!groupStyle.fillColor && !groupStyle.fillGradient && groupStyle.opacity === undefined)
    return model;
  const spPr = firstChildNs(element, P_NS, 'spPr');
  if (!spPr || !firstChildNs(spPr, A_NS, 'grpFill')) return model;
  const style = {
    ...model.style,
    ...(model.style?.fillColor || !groupStyle.fillColor ? {} : { fillColor: groupStyle.fillColor }),
    ...(model.style?.fillGradient || !groupStyle.fillGradient
      ? {}
      : { fillGradient: groupStyle.fillGradient }),
    ...(model.style?.opacity !== undefined || groupStyle.opacity === undefined
      ? {}
      : { opacity: groupStyle.opacity }),
  };
  return {
    ...model,
    style,
  };
}

function colorFromStyleFillRef(
  element: Element,
  context: PptParseContext = {},
): string | undefined {
  const style = firstChildNs(element, P_NS, 'style');
  const fillRef = style ? firstChildNs(style, A_NS, 'fillRef') : null;
  if (!fillRef) return undefined;
  return (
    colorFromColorElement(firstChildNs(fillRef, A_NS, 'srgbClr'), context) ??
    colorFromColorElement(firstChildNs(fillRef, A_NS, 'schemeClr'), context) ??
    colorFromColorElement(firstChildNs(fillRef, A_NS, 'sysClr'), context) ??
    colorFromColorElement(firstChildNs(fillRef, A_NS, 'prstClr'), context)
  );
}

function parseLineDash(line: Element): string | undefined {
  const preset = firstChildNs(line, A_NS, 'prstDash')?.getAttribute('val') ?? undefined;
  if (preset) return preset;
  const customDash = firstChildNs(line, A_NS, 'custDash');
  const firstDash = customDash ? firstChildNs(customDash, A_NS, 'ds') : undefined;
  const dash = parseNumber(firstDash?.getAttribute('d') ?? null);
  const space = parseNumber(firstDash?.getAttribute('sp') ?? null);
  if (dash === undefined || space === undefined || dash <= 0 || space < 0) return undefined;
  return `cust:${Math.round(dash)}:${Math.round(space)}`;
}

function colorFromStyleFontRef(
  element: Element,
  context: PptParseContext = {},
): string | undefined {
  const style = firstChildNs(element, P_NS, 'style');
  const fontRef = style ? firstChildNs(style, A_NS, 'fontRef') : null;
  if (!fontRef) return undefined;
  return (
    colorFromColorElement(firstChildNs(fontRef, A_NS, 'srgbClr'), context) ??
    colorFromColorElement(firstChildNs(fontRef, A_NS, 'schemeClr'), context) ??
    colorFromColorElement(firstChildNs(fontRef, A_NS, 'sysClr'), context) ??
    colorFromColorElement(firstChildNs(fontRef, A_NS, 'prstClr'), context)
  );
}

function parseShapeShadow(
  spPr: Element,
  context: PptParseContext = {},
): PptShapeStyle['shadow'] | undefined {
  return parseOuterShadow(firstChildNs(spPr, A_NS, 'effectLst'), context);
}

function parseShape3d(spPr: Element): NonNullable<PptShapeModel['raw']>['threeD'] | undefined {
  const sp3d = firstChildNs(spPr, A_NS, 'sp3d');
  if (!sp3d) return undefined;
  const material = sp3d.getAttribute('prstMaterial') ?? undefined;
  const extrusionHeight = parseNumber(sp3d.getAttribute('extrusionH'));
  return {
    ...(material ? { material } : {}),
    ...(extrusionHeight !== undefined ? { extrusionHeight } : {}),
  };
}

function firstTextShadow(
  elements: Element[],
  context: PptParseContext = {},
): PptShadowModel | undefined {
  for (const element of elements) {
    const shadow = parseOuterShadow(firstChildNs(element, A_NS, 'effectLst'), context);
    if (shadow) return shadow;
  }
  return undefined;
}

function parseOuterShadow(
  effectList: Element | null | undefined,
  context: PptParseContext = {},
): PptShadowModel | undefined {
  const shadow = effectList ? firstChildNs(effectList, A_NS, 'outerShdw') : undefined;
  if (!shadow) return undefined;
  const colorElement =
    firstChildNs(shadow, A_NS, 'srgbClr') ??
    firstChildNs(shadow, A_NS, 'schemeClr') ??
    firstChildNs(shadow, A_NS, 'sysClr') ??
    firstChildNs(shadow, A_NS, 'prstClr');
  const color = colorFromColorElement(colorElement, context);
  const opacity = colorElement
    ? parseNumber(firstChildNs(colorElement, A_NS, 'alpha')?.getAttribute('val') ?? null)
    : undefined;
  const blur = parseNumber(shadow.getAttribute('blurRad'));
  const distance = parseNumber(shadow.getAttribute('dist'));
  const direction = parseNumber(shadow.getAttribute('dir'));
  return {
    ...(color ? { color } : {}),
    ...(opacity !== undefined ? { opacity: opacity / 100000 } : {}),
    ...(blur !== undefined ? { blur } : {}),
    ...(distance !== undefined ? { distance } : {}),
    ...(direction !== undefined ? { direction: direction / 60000 } : {}),
  };
}

function parseSlideBackground(
  doc: Document,
  context: PptParseContext = {},
): PptSlideBackgroundModel | undefined {
  const cSld = allDescendantsNs(doc, P_NS, 'cSld')[0];
  const bg = cSld ? firstChildNs(cSld, P_NS, 'bg') : undefined;
  const bgPr = bg ? firstChildNs(bg, P_NS, 'bgPr') : undefined;
  if (!bgPr) return undefined;
  const solidFill = parseSolidFill(firstChildNs(bgPr, A_NS, 'solidFill'), context);
  const fillGradient = parseGradientFill(firstChildNs(bgPr, A_NS, 'gradFill'), context);
  const imagePath = parseBackgroundImagePath(bgPr, context);
  const background = {
    ...(solidFill.color ? { fillColor: solidFill.color } : {}),
    ...(fillGradient ? { fillGradient } : {}),
    ...(imagePath ? { imagePath } : {}),
  };
  return Object.keys(background).length > 0 ? background : undefined;
}

function parseBackgroundImagePath(
  bgPr: Element,
  context: PptParseContext = {},
): string | undefined {
  const blip = allDescendantsNs(bgPr, A_NS, 'blip')[0];
  const relId = blip?.getAttributeNS(R_NS, 'embed') ?? blip?.getAttribute('r:embed');
  if (!relId || !context.relationships || !context.sourcePath) return undefined;
  const rel = context.relationships.get(relId);
  if (!rel || rel.type !== IMAGE_REL_TYPE) return undefined;
  return resolveRelationshipTarget(context.sourcePath, rel.target);
}

function parsePresetGeometry(element: Element): string | undefined {
  const spPr = allDescendantsNs(element, P_NS, 'spPr')[0];
  if (!spPr) return undefined;
  return firstChildNs(spPr, A_NS, 'prstGeom')?.getAttribute('prst') ?? undefined;
}

function parsePresetGeometryAdjustments(element: Element): Record<string, number> | undefined {
  const spPr = allDescendantsNs(element, P_NS, 'spPr')[0];
  const presetGeometry = spPr ? firstChildNs(spPr, A_NS, 'prstGeom') : undefined;
  const adjustmentList = presetGeometry ? firstChildNs(presetGeometry, A_NS, 'avLst') : undefined;
  if (!adjustmentList) return undefined;
  const adjustments: Record<string, number> = {};
  for (const guide of childElementsNs(adjustmentList, A_NS, 'gd')) {
    const name = guide.getAttribute('name');
    const match = /^val\s+(-?\d+(?:\.\d+)?)$/i.exec(guide.getAttribute('fmla') ?? '');
    if (!name || !match) continue;
    adjustments[name] = Number(match[1]);
  }
  return Object.keys(adjustments).length > 0 ? adjustments : undefined;
}

function parseCustomGeometry(element: Element): PptCustomGeometryModel | undefined {
  const spPr = allDescendantsNs(element, P_NS, 'spPr')[0];
  const customGeometry = spPr ? firstChildNs(spPr, A_NS, 'custGeom') : undefined;
  const pathList = customGeometry ? firstChildNs(customGeometry, A_NS, 'pathLst') : undefined;
  if (!customGeometry || !pathList) return undefined;
  const paths = childElementsNs(pathList, A_NS, 'path')
    .map((path) => parseCustomGeometryPath(path, customGeometry))
    .filter((path) => path.commands.length > 0);
  return paths.length > 0 ? { paths } : undefined;
}

function parseCustomGeometryPath(
  path: Element,
  customGeometry: Element,
): PptCustomGeometryModel['paths'][number] {
  const width = parseNumber(path.getAttribute('w') ?? null);
  const height = parseNumber(path.getAttribute('h') ?? null);
  const guides = parseCustomGeometryGuides(customGeometry, width, height);
  const commands: PptCustomGeometryCommand[] = [];
  for (let i = 0; i < path.childNodes.length; i++) {
    const child = path.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const command = child as Element;
    if (command.namespaceURI !== A_NS) continue;
    if (command.localName === 'moveTo' || command.localName === 'lnTo') {
      const point = parseCustomGeometryPoint(firstChildNs(command, A_NS, 'pt'), guides);
      if (!point) continue;
      commands.push({ kind: command.localName === 'moveTo' ? 'moveTo' : 'lineTo', ...point });
    } else if (command.localName === 'cubicBezTo') {
      const points = childElementsNs(command, A_NS, 'pt')
        .map((point) => parseCustomGeometryPoint(point, guides))
        .filter((point): point is { x: number; y: number } => Boolean(point));
      if (points.length !== 3) continue;
      commands.push({
        kind: 'cubicBezierTo',
        x1: points[0]!.x,
        y1: points[0]!.y,
        x2: points[1]!.x,
        y2: points[1]!.y,
        x: points[2]!.x,
        y: points[2]!.y,
      });
    } else if (command.localName === 'arcTo') {
      const arc = parseCustomGeometryArc(command, guides);
      if (arc) commands.push(arc);
    } else if (command.localName === 'close') {
      commands.push({ kind: 'close' });
    }
  }
  return {
    ...(width !== undefined ? { width } : {}),
    ...(height !== undefined ? { height } : {}),
    commands,
  };
}

function parseCustomGeometryArc(
  command: Element,
  guides: Record<string, number>,
): Extract<PptCustomGeometryCommand, { kind: 'arcTo' }> | undefined {
  const wR = customGeometryNumber(command.getAttribute('wR'), undefined, undefined, guides);
  const hR = customGeometryNumber(command.getAttribute('hR'), undefined, undefined, guides);
  const stAng = customGeometryNumber(command.getAttribute('stAng'), undefined, undefined, guides);
  const swAng = customGeometryNumber(command.getAttribute('swAng'), undefined, undefined, guides);
  if (wR === undefined || hR === undefined || stAng === undefined || swAng === undefined)
    return undefined;
  return {
    kind: 'arcTo',
    wR,
    hR,
    stAng,
    swAng,
  };
}

function parseCustomGeometryGuides(
  customGeometry: Element,
  width: number | undefined,
  height: number | undefined,
): Record<string, number> {
  const guides: Record<string, number> = {};
  const guideList = firstChildNs(customGeometry, A_NS, 'gdLst');
  if (!guideList) return guides;
  for (const guide of childElementsNs(guideList, A_NS, 'gd')) {
    const name = guide.getAttribute('name');
    const value = evaluateCustomGeometryFormula(guide.getAttribute('fmla'), width, height, guides);
    if (name && value !== undefined) guides[name] = value;
  }
  return guides;
}

function evaluateCustomGeometryFormula(
  formula: string | null,
  width: number | undefined,
  height: number | undefined,
  guides: Record<string, number>,
): number | undefined {
  if (!formula) return undefined;
  const parts = formula.trim().split(/\s+/);
  if (parts[0] === 'val') return customGeometryNumber(parts[1], width, height, guides);
  if (parts[0] === '*/') {
    const a = customGeometryNumber(parts[1], width, height, guides);
    const b = customGeometryNumber(parts[2], width, height, guides);
    const c = customGeometryNumber(parts[3], width, height, guides);
    if (a === undefined || b === undefined || !c) return undefined;
    return (a * b) / c;
  }
  return undefined;
}

function parseCustomGeometryPoint(
  point: Element | null,
  guides: Record<string, number>,
): { x: number; y: number } | undefined {
  if (!point) return undefined;
  const x = customGeometryNumber(point.getAttribute('x'), undefined, undefined, guides);
  const y = customGeometryNumber(point.getAttribute('y'), undefined, undefined, guides);
  return x === undefined || y === undefined ? undefined : { x, y };
}

function customGeometryNumber(
  value: string | undefined | null,
  width: number | undefined,
  height: number | undefined,
  guides: Record<string, number>,
): number | undefined {
  if (!value) return undefined;
  if (value === 'w') return width;
  if (value === 'h') return height;
  if (guides[value] !== undefined) return guides[value];
  return parseNumber(value);
}

function parsePlaceholder(element: Element): PptShapeModel['placeholder'] | undefined {
  const ph = allDescendantsNs(element, P_NS, 'ph')[0];
  if (!ph) return undefined;
  const type = ph.getAttribute('type') ?? undefined;
  const index = ph.getAttribute('idx') ?? undefined;
  const size = ph.getAttribute('sz') ?? undefined;
  const orientation = ph.getAttribute('orient') ?? undefined;
  return {
    ...(type ? { type } : {}),
    ...(index ? { index } : {}),
    ...(size ? { size } : {}),
    ...(orientation ? { orientation } : {}),
  };
}

function parseTextStyle(element: Element, context: PptParseContext = {}): PptTextStyle {
  const textProperties = [
    ...allDescendantsNs(element, A_NS, 'rPr'),
    ...allDescendantsNs(element, A_NS, 'defRPr'),
  ];
  const size = firstParsedNumber(textProperties, 'sz');
  const color =
    firstSolidFillColor(textProperties, context) ??
    firstGradientFillColor(textProperties, context) ??
    colorFromStyleFontRef(element, context);
  const textGradient = firstGradientFill(textProperties, context);
  const textFill = textProperties.some((properties) => firstChildNs(properties, A_NS, 'noFill'))
    ? 'none'
    : undefined;
  const textOutlineColor = firstTextOutlineColor(textProperties, context);
  const textShadow = firstTextShadow(textProperties, context);
  const fontFamily = firstTypeface(textProperties, context);
  const bold =
    consistentParsedBoolean(textProperties, 'b') ?? singleTextRunParsedBoolean(element, 'b');
  const italic =
    consistentParsedBoolean(textProperties, 'i') ?? singleTextRunParsedBoolean(element, 'i');
  const underline =
    consistentParsedUnderline(textProperties) ?? singleTextRunParsedUnderline(element);
  const strike = consistentParsedStrike(textProperties) ?? singleTextRunParsedStrike(element);
  const baseline = firstParsedBaseline(textProperties);
  const highlight = firstHighlightColor(textProperties, context);
  const alignment = firstParagraphAlignment(element);
  const lineHeight = firstParagraphLineHeight(element);
  const characterSpacing = firstParsedCharacterSpacing(textProperties);
  return {
    ...(size !== undefined ? { fontSize: size / 100 } : {}),
    ...(fontFamily ? { fontFamily } : {}),
    ...(color ? { color } : {}),
    ...(textGradient ? { textGradient } : {}),
    ...(textFill ? { textFill } : {}),
    ...(textOutlineColor ? { textOutlineColor } : {}),
    ...(textShadow ? { textShadow } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underline !== undefined ? { underline } : {}),
    ...(strike !== undefined ? { strike } : {}),
    ...(baseline !== undefined ? { baseline } : {}),
    ...(highlight ? { highlight } : {}),
    ...(alignment ? { alignment } : {}),
    ...(lineHeight !== undefined ? { lineHeight } : {}),
    ...(characterSpacing !== undefined ? { characterSpacing } : {}),
  };
}

function parseTextBodyProperties(element: Element): PptTextBodyProperties | undefined {
  const bodyPr = allDescendantsNs(element, A_NS, 'bodyPr')[0];
  if (!bodyPr) return undefined;
  const anchor = parseTextBodyAnchor(bodyPr.getAttribute('anchor'));
  const autoFit = parseTextAutoFit(bodyPr);
  const wrap = parseTextBodyWrap(bodyPr.getAttribute('wrap'));
  const vertical = parseTextBodyVertical(bodyPr.getAttribute('vert'));
  const insetLeft = parseNumber(bodyPr.getAttribute('lIns'));
  const insetTop = parseNumber(bodyPr.getAttribute('tIns'));
  const insetRight = parseNumber(bodyPr.getAttribute('rIns'));
  const insetBottom = parseNumber(bodyPr.getAttribute('bIns'));
  const properties: PptTextBodyProperties = {
    ...(anchor ? { anchor } : {}),
    ...(autoFit ? { autoFit } : {}),
    ...(wrap ? { wrap } : {}),
    ...(vertical ? { vertical } : {}),
    ...(insetLeft !== undefined ? { insetLeft } : {}),
    ...(insetTop !== undefined ? { insetTop } : {}),
    ...(insetRight !== undefined ? { insetRight } : {}),
    ...(insetBottom !== undefined ? { insetBottom } : {}),
  };
  return properties;
}

function parseTextAutoFit(bodyPr: Element): PptTextBodyProperties['autoFit'] | undefined {
  if (firstChildNs(bodyPr, A_NS, 'spAutoFit')) return 'shape';
  if (firstChildNs(bodyPr, A_NS, 'normAutofit')) return 'normal';
  if (firstChildNs(bodyPr, A_NS, 'noAutofit')) return 'none';
  return undefined;
}

function parseTextBodyWrap(value: string | null): PptTextBodyProperties['wrap'] | undefined {
  if (value === 'none') return 'none';
  if (value === 'square') return 'square';
  return undefined;
}

function parseTextBodyVertical(
  value: string | null,
): PptTextBodyProperties['vertical'] | undefined {
  if (value === 'eaVert') return 'eastAsian';
  if (value === 'vert') return 'vertical';
  return undefined;
}

function parseTextBodyAnchor(value: string | null): PptTextBodyProperties['anchor'] | undefined {
  if (value === 'ctr') return 'middle';
  if (value === 'b') return 'bottom';
  if (value === 't') return 'top';
  return undefined;
}

function firstParsedNumber(elements: Element[], attribute: string): number | undefined {
  for (const element of elements) {
    const value = parseNumber(element.getAttribute(attribute));
    if (value !== undefined) return value;
  }
  return undefined;
}

function firstParsedBoolean(elements: Element[], attribute: string): boolean | undefined {
  for (const element of elements) {
    const value = parseBooleanAttr(element.getAttribute(attribute));
    if (value !== undefined) return value;
  }
  return undefined;
}

function consistentParsedBoolean(elements: Element[], attribute: string): boolean | undefined {
  let value: boolean | undefined;
  let hasMissingValue = false;
  for (const element of elements) {
    const parsed = parseBooleanAttr(element.getAttribute(attribute));
    if (parsed === undefined) {
      hasMissingValue = true;
      continue;
    }
    if (value !== undefined && value !== parsed) return undefined;
    value = parsed;
  }
  return hasMissingValue ? undefined : value;
}

function consistentParsedUnderline(elements: Element[]): boolean | undefined {
  let value: boolean | undefined;
  let hasMissingValue = false;
  for (const element of elements) {
    const parsed = parseUnderlineAttr(element.getAttribute('u'));
    if (parsed === undefined) {
      hasMissingValue = true;
      continue;
    }
    if (value !== undefined && value !== parsed) return undefined;
    value = parsed;
  }
  return hasMissingValue ? undefined : value;
}

function consistentParsedStrike(elements: Element[]): boolean | undefined {
  let value: boolean | undefined;
  let hasMissingValue = false;
  for (const element of elements) {
    const parsed = parseStrikeAttr(element.getAttribute('strike'));
    if (parsed === undefined) {
      hasMissingValue = true;
      continue;
    }
    if (value !== undefined && value !== parsed) return undefined;
    value = parsed;
  }
  return hasMissingValue ? undefined : value;
}

function singleTextRunParsedBoolean(element: Element, attribute: string): boolean | undefined {
  const textRuns = allDescendantsNs(element, A_NS, 'r').filter((run) =>
    childElementsNs(run, A_NS, 't').some((text) => (text.textContent ?? '').length > 0),
  );
  if (textRuns.length !== 1) return undefined;
  const runProperties = firstChildNs(textRuns[0]!, A_NS, 'rPr');
  return parseBooleanAttr(runProperties?.getAttribute(attribute) ?? null);
}

function singleTextRunParsedUnderline(element: Element): boolean | undefined {
  const textRuns = allDescendantsNs(element, A_NS, 'r').filter((run) =>
    childElementsNs(run, A_NS, 't').some((text) => (text.textContent ?? '').length > 0),
  );
  if (textRuns.length !== 1) return undefined;
  const runProperties = firstChildNs(textRuns[0]!, A_NS, 'rPr');
  return parseUnderlineAttr(runProperties?.getAttribute('u') ?? null);
}

function singleTextRunParsedStrike(element: Element): boolean | undefined {
  const textRuns = allDescendantsNs(element, A_NS, 'r').filter((run) =>
    childElementsNs(run, A_NS, 't').some((text) => (text.textContent ?? '').length > 0),
  );
  if (textRuns.length !== 1) return undefined;
  const runProperties = firstChildNs(textRuns[0]!, A_NS, 'rPr');
  return parseStrikeAttr(runProperties?.getAttribute('strike') ?? null);
}

function firstParsedCharacterSpacing(elements: Element[]): number | undefined {
  const value = firstParsedNumber(elements, 'spc');
  return value !== undefined ? value / 100 : undefined;
}

function firstParsedBaseline(elements: Element[]): number | undefined {
  const value = firstParsedNumber(elements, 'baseline');
  return value !== undefined ? value / 100000 : undefined;
}

function parseBaseline(value: string | null): number | undefined {
  const parsed = parseNumber(value);
  return parsed !== undefined ? parsed / 100000 : undefined;
}

function firstHighlightColor(
  elements: Element[],
  context: PptParseContext = {},
): string | undefined {
  for (const element of elements) {
    const highlight = firstChildNs(element, A_NS, 'highlight');
    const colorElement = highlight
      ? (firstChildNs(highlight, A_NS, 'srgbClr') ??
        firstChildNs(highlight, A_NS, 'schemeClr') ??
        firstChildNs(highlight, A_NS, 'sysClr') ??
        firstChildNs(highlight, A_NS, 'prstClr'))
      : null;
    const color = colorFromColorElement(colorElement, context);
    if (color) return color;
  }
  return undefined;
}

function firstTypeface(elements: Element[], context: PptParseContext = {}): string | undefined {
  for (const element of elements) {
    const typeface =
      firstChildNs(element, A_NS, 'ea')?.getAttribute('typeface') ??
      firstChildNs(element, A_NS, 'latin')?.getAttribute('typeface') ??
      firstChildNs(element, A_NS, 'cs')?.getAttribute('typeface') ??
      undefined;
    const family = normalizeTypeface(typeface, context);
    if (family) return family;
  }
  return undefined;
}

function normalizeTypeface(
  typeface: string | null | undefined,
  context: PptParseContext = {},
): string | undefined {
  if (!typeface) return undefined;
  if (typeface === '+mj-lt') return context.fontScheme?.majorLatin;
  if (typeface === '+mj-ea') return context.fontScheme?.majorEa ?? context.fontScheme?.majorLatin;
  if (typeface === '+mj-cs') return context.fontScheme?.majorCs ?? context.fontScheme?.majorLatin;
  if (typeface === '+mn-lt') return context.fontScheme?.minorLatin;
  if (typeface === '+mn-ea') return context.fontScheme?.minorEa ?? context.fontScheme?.minorLatin;
  if (typeface === '+mn-cs') return context.fontScheme?.minorCs ?? context.fontScheme?.minorLatin;
  if (typeface.startsWith('+')) return undefined;
  return typeface;
}

function firstSolidFillColor(
  elements: Element[],
  context: PptParseContext = {},
): string | undefined {
  for (const element of elements) {
    const color = colorFromSolidFill(firstChildNs(element, A_NS, 'solidFill'), context);
    if (color) return color;
  }
  return undefined;
}

function firstGradientFillColor(
  elements: Element[],
  context: PptParseContext = {},
): string | undefined {
  return firstGradientFill(elements, context)?.colors[0];
}

function firstGradientFill(
  elements: Element[],
  context: PptParseContext = {},
): PptShapeStyle['fillGradient'] | undefined {
  for (const element of elements) {
    const gradient = parseGradientFill(firstChildNs(element, A_NS, 'gradFill'), context);
    if (gradient) return gradient;
  }
  return undefined;
}

function firstTextOutlineColor(
  elements: Element[],
  context: PptParseContext = {},
): string | undefined {
  for (const element of elements) {
    const line = firstChildNs(element, A_NS, 'ln');
    const color = line
      ? colorFromSolidFill(firstChildNs(line, A_NS, 'solidFill'), context)
      : undefined;
    if (color) return color;
  }
  return undefined;
}

function firstParagraphAlignment(element: Element): PptTextStyle['alignment'] | undefined {
  for (const candidate of allDescendantsNs(element, A_NS, 'pPr')) {
    const alignment = parseAlignment(candidate.getAttribute('algn'));
    if (alignment) return alignment;
  }
  for (const listStyle of allDescendantsNs(element, A_NS, 'lstStyle')) {
    for (let i = 0; i < listStyle.childNodes.length; i++) {
      const child = listStyle.childNodes[i];
      if (child?.nodeType !== 1) continue;
      const candidate = child as Element;
      if (candidate.namespaceURI !== A_NS || !candidate.localName.endsWith('pPr')) continue;
      const alignment = parseAlignment(candidate.getAttribute('algn'));
      if (alignment) return alignment;
    }
  }
  return undefined;
}

function firstParagraphLineHeight(element: Element): number | undefined {
  for (const pPr of allDescendantsNs(element, A_NS, 'pPr')) {
    const lnSpc = firstChildNs(pPr, A_NS, 'lnSpc');
    const pct = lnSpc
      ? parseNumber(firstChildNs(lnSpc, A_NS, 'spcPct')?.getAttribute('val') ?? null)
      : undefined;
    if (pct !== undefined) return pct / 100000;
  }
  return undefined;
}

function parseTextParagraphs(
  slide: PptSlideInternal,
  element: Element,
  context: PptParseContext = {},
): PptTextParagraph[] {
  return allDescendantsNs(element, A_NS, 'p').map((paragraph) => {
    const pPr = firstChildNs(paragraph, A_NS, 'pPr');
    const alignment = parseAlignment(pPr?.getAttribute('algn') ?? null);
    const bullet = parseParagraphBullet(pPr, context);
    const paragraphStyle = parseParagraphStyle(pPr, alignment, bullet);
    const runs = parseTextParagraphRuns(slide, paragraph, alignment, context);
    const paragraphText = runs.map((run) => run.text).join('');
    return {
      text: paragraphText,
      ...(Object.keys(paragraphStyle).length > 0 ? { style: paragraphStyle } : {}),
      runs,
    };
  });
}

function parseTextParagraphRuns(
  slide: PptSlideInternal,
  paragraph: Element,
  alignment: PptTextStyle['alignment'],
  context: PptParseContext,
): PptTextRun[] {
  const runs: PptTextRun[] = [];
  for (let i = 0; i < paragraph.childNodes.length; i++) {
    const child = paragraph.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI !== A_NS) continue;
    if (element.localName === 'r') {
      const text = allDescendantsNs(element, A_NS, 't').map(textOf).join('');
      const style = parseRunTextStyle(element, alignment, context);
      const hyperlink = parseRunHyperlink(slide, element);
      runs.push({
        text,
        ...(Object.keys(style).length > 0 ? { style } : {}),
        ...(hyperlink ? { hyperlink } : {}),
      });
    } else if (element.localName === 'br') {
      runs.push({ text: '\n' });
    }
  }
  return runs;
}

function textParagraphsToPlainText(paragraphs: PptTextParagraph[]): string {
  return paragraphs.map((paragraph) => paragraph.text).join('\n');
}

function parseParagraphStyle(
  pPr: Element | null,
  alignment: PptTextStyle['alignment'],
  bullet: { char: string; color?: string; fontFamily?: string; size?: number } | undefined,
): NonNullable<PptTextParagraph['style']> {
  const lineHeight = parseParagraphLineHeight(pPr);
  const marginLeft = parseNumber(pPr?.getAttribute('marL') ?? null);
  const textIndent = parseNumber(pPr?.getAttribute('indent') ?? null);
  const spaceBefore = parseParagraphSpacing(pPr ? firstChildNs(pPr, A_NS, 'spcBef') : null);
  const spaceAfter = parseParagraphSpacing(pPr ? firstChildNs(pPr, A_NS, 'spcAft') : null);
  return {
    ...(alignment ? { alignment } : {}),
    ...(bullet ? { bullet: bullet.char } : {}),
    ...(bullet?.color ? { bulletColor: bullet.color } : {}),
    ...(bullet?.fontFamily ? { bulletFontFamily: bullet.fontFamily } : {}),
    ...(bullet?.size !== undefined ? { bulletSize: bullet.size } : {}),
    ...(lineHeight !== undefined ? { lineHeight } : {}),
    ...(marginLeft !== undefined ? { marginLeft } : {}),
    ...(textIndent !== undefined ? { textIndent } : {}),
    ...(spaceBefore !== undefined ? { spaceBefore } : {}),
    ...(spaceAfter !== undefined ? { spaceAfter } : {}),
  };
}

function parseParagraphLineHeight(pPr: Element | null): number | undefined {
  const lnSpc = pPr ? firstChildNs(pPr, A_NS, 'lnSpc') : null;
  const pct = lnSpc
    ? parseNumber(firstChildNs(lnSpc, A_NS, 'spcPct')?.getAttribute('val') ?? null)
    : undefined;
  if (pct !== undefined) return pct / 100000;
  return undefined;
}

function parseParagraphSpacing(element: Element | null): number | undefined {
  if (!element) return undefined;
  const points = parseNumber(firstChildNs(element, A_NS, 'spcPts')?.getAttribute('val') ?? null);
  if (points !== undefined) return points * 127;
  const pct = parseNumber(firstChildNs(element, A_NS, 'spcPct')?.getAttribute('val') ?? null);
  return pct !== undefined ? pct / 100000 : undefined;
}

function parseParagraphBullet(
  pPr: Element | null,
  context: PptParseContext = {},
): { char: string; color?: string; fontFamily?: string; size?: number } | undefined {
  if (!pPr || firstChildNs(pPr, A_NS, 'buNone')) return undefined;
  const typeface = firstChildNs(pPr, A_NS, 'buFont')?.getAttribute('typeface') ?? undefined;
  const color = parseBulletColor(pPr, context);
  const size = parseBulletSize(pPr);
  const buChar = firstChildNs(pPr, A_NS, 'buChar')?.getAttribute('char');
  if (buChar) {
    return {
      char: normalizeBulletCharacter(buChar, typeface),
      ...(color ? { color } : {}),
      ...(typeface ? { fontFamily: typeface } : {}),
      ...(size !== undefined ? { size } : {}),
    };
  }
  if (firstChildNs(pPr, A_NS, 'buAutoNum')) {
    return {
      char: '1.',
      ...(color ? { color } : {}),
      ...(typeface ? { fontFamily: typeface } : {}),
      ...(size !== undefined ? { size } : {}),
    };
  }
  return undefined;
}

function parseBulletSize(pPr: Element): number | undefined {
  const pct = parseNumber(firstChildNs(pPr, A_NS, 'buSzPct')?.getAttribute('val') ?? null);
  if (pct !== undefined) return pct / 100000;
  const pts = parseNumber(firstChildNs(pPr, A_NS, 'buSzPts')?.getAttribute('val') ?? null);
  return pts !== undefined ? pts / 100 : undefined;
}

function normalizeBulletCharacter(char: string, typeface: string | null | undefined): string {
  const family = typeface?.toLowerCase() ?? '';
  if (!family.includes('wingdings') && !family.includes('webdings')) return char;
  if (char === 'ü' || char === '\u00fc') return '✓';
  if (char === 'û' || char === '\u00fb') return '✗';
  if (char === 'p') return '□';
  if (char === 'n') return '■';
  if (char === 'l') return '●';
  if (char === 'u') return '◆';
  return '•';
}

function parseBulletColor(pPr: Element, context: PptParseContext = {}): string | undefined {
  const buClr = firstChildNs(pPr, A_NS, 'buClr');
  if (!buClr) return undefined;
  const colorElement =
    firstChildNs(buClr, A_NS, 'srgbClr') ??
    firstChildNs(buClr, A_NS, 'schemeClr') ??
    firstChildNs(buClr, A_NS, 'sysClr') ??
    firstChildNs(buClr, A_NS, 'prstClr');
  return colorFromColorElement(colorElement, context);
}

function parseRunHyperlink(slide: PptSlideInternal, run: Element): PptHyperlinkModel | undefined {
  const rPr = firstChildNs(run, A_NS, 'rPr');
  const hlink = rPr ? firstChildNs(rPr, A_NS, 'hlinkClick') : null;
  return parseHyperlink(slide, hlink);
}

function parseShapeHyperlink(
  slide: PptSlideInternal,
  cNvPr: Element | undefined,
): PptHyperlinkModel | undefined {
  if (!cNvPr) return undefined;
  return parseHyperlink(slide, firstChildNs(cNvPr, A_NS, 'hlinkClick'));
}

function parseHyperlink(
  slide: PptSlideInternal,
  hlink: Element | null,
): PptHyperlinkModel | undefined {
  if (!hlink) return undefined;
  const relationshipId =
    hlink.getAttributeNS(R_NS, 'id') ?? hlink.getAttribute('r:id') ?? undefined;
  const rel = relationshipId ? slide.relationships.get(relationshipId) : undefined;
  const action = hlink.getAttribute('action') ?? undefined;
  const tooltip = hlink.getAttribute('tooltip') ?? undefined;
  return {
    ...(relationshipId ? { relationshipId } : {}),
    ...(rel?.target ? { target: rel.target } : {}),
    ...(rel?.targetMode ? { targetMode: rel.targetMode } : {}),
    ...(action ? { action } : {}),
    ...(tooltip ? { tooltip } : {}),
  };
}

function parseRunTextStyle(
  run: Element,
  alignment: PptTextStyle['alignment'],
  context: PptParseContext = {},
): PptTextStyle {
  const rPr = firstChildNs(run, A_NS, 'rPr');
  const size = parseNumber(rPr?.getAttribute('sz') ?? null);
  const color = rPr
    ? (colorFromSolidFill(firstChildNs(rPr, A_NS, 'solidFill'), context) ??
      firstGradientFillColor([rPr], context))
    : undefined;
  const textGradient = rPr ? firstGradientFill([rPr], context) : undefined;
  const textFill = rPr && firstChildNs(rPr, A_NS, 'noFill') ? 'none' : undefined;
  const textOutlineColor = rPr ? firstTextOutlineColor([rPr], context) : undefined;
  const textShadow = rPr ? firstTextShadow([rPr], context) : undefined;
  const fontFamily = rPr ? firstTypeface([rPr], context) : undefined;
  const bold = parseBooleanAttr(rPr?.getAttribute('b') ?? null);
  const italic = parseBooleanAttr(rPr?.getAttribute('i') ?? null);
  const underline = parseUnderlineAttr(rPr?.getAttribute('u') ?? null);
  const strike = parseStrikeAttr(rPr?.getAttribute('strike') ?? null);
  const baseline = parseBaseline(rPr?.getAttribute('baseline') ?? null);
  const highlight = rPr ? firstHighlightColor([rPr], context) : undefined;
  const characterSpacing = rPr ? firstParsedCharacterSpacing([rPr]) : undefined;
  return {
    ...(size !== undefined ? { fontSize: size / 100 } : {}),
    ...(fontFamily ? { fontFamily } : {}),
    ...(color ? { color } : {}),
    ...(textGradient ? { textGradient } : {}),
    ...(textFill ? { textFill } : {}),
    ...(textOutlineColor ? { textOutlineColor } : {}),
    ...(textShadow ? { textShadow } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underline !== undefined ? { underline } : {}),
    ...(strike !== undefined ? { strike } : {}),
    ...(baseline !== undefined ? { baseline } : {}),
    ...(highlight ? { highlight } : {}),
    ...(alignment ? { alignment } : {}),
    ...(characterSpacing !== undefined ? { characterSpacing } : {}),
  };
}

function parseBooleanAttr(value: string | null): boolean | undefined {
  if (value === null) return undefined;
  return value === '1' || value === 'true';
}

function parseUnderlineAttr(value: string | null): boolean | undefined {
  if (value === null) return undefined;
  return value !== 'none';
}

function parseStrikeAttr(value: string | null): boolean | undefined {
  if (value === null) return undefined;
  return value !== 'noStrike';
}

function parseAlignment(value: string | null): PptTextStyle['alignment'] | undefined {
  if (value === 'ctr') return 'center';
  if (value === 'r') return 'right';
  if (value === 'just') return 'justify';
  if (value === 'dist' || value === 'thaiDist') return 'distributed';
  if (value === 'l') return 'left';
  return undefined;
}

function colorFromSolidFill(
  solidFill: Element | null,
  context: PptParseContext = {},
): string | undefined {
  return parseSolidFill(solidFill, context).color;
}

function parseSolidFill(
  solidFill: Element | null,
  context: PptParseContext = {},
): { color?: string; opacity?: number } {
  if (!solidFill) return {};
  const colorElement =
    firstChildNs(solidFill, A_NS, 'srgbClr') ??
    firstChildNs(solidFill, A_NS, 'schemeClr') ??
    firstChildNs(solidFill, A_NS, 'sysClr') ??
    firstChildNs(solidFill, A_NS, 'prstClr');
  const color = colorFromColorElement(colorElement, context);
  const alpha = colorElement
    ? parseNumber(firstChildNs(colorElement, A_NS, 'alpha')?.getAttribute('val') ?? null)
    : undefined;
  return {
    ...(color ? { color } : {}),
    ...(alpha !== undefined ? { opacity: alpha / 100000 } : {}),
  };
}

function parseGradientFill(
  gradFill: Element | null,
  context: PptParseContext = {},
): PptShapeStyle['fillGradient'] | undefined {
  if (!gradFill) return undefined;
  const stops = allDescendantsNs(gradFill, A_NS, 'gs')
    .map((stop) => {
      const colorElement =
        firstChildNs(stop, A_NS, 'srgbClr') ??
        firstChildNs(stop, A_NS, 'schemeClr') ??
        firstChildNs(stop, A_NS, 'sysClr') ??
        firstChildNs(stop, A_NS, 'prstClr');
      const color = colorElement && colorFromColorElement(colorElement, context);
      const position = parseNumber(stop.getAttribute('pos') ?? null);
      const alpha = colorElement
        ? parseNumber(firstChildNs(colorElement, A_NS, 'alpha')?.getAttribute('val') ?? null)
        : undefined;
      return { color, position, opacity: alpha === undefined ? undefined : alpha / 100000 };
    })
    .filter(
      (
        stop,
      ): stop is { color: string; position: number | undefined; opacity: number | undefined } =>
        !!stop.color,
    );
  const colors = stops.map((stop) => stop.color);
  if (colors.length === 0) return undefined;
  const positions = stops.map((stop) =>
    stop.position === undefined ? undefined : stop.position / 1000,
  );
  const opacities = stops.map((stop) => stop.opacity);
  const linear = firstChildNs(gradFill, A_NS, 'lin');
  const tileRect = firstChildNs(gradFill, A_NS, 'tileRect');
  const angle = parseNumber(linear?.getAttribute('ang') ?? null);
  const scaled = parseBooleanAttribute(linear?.getAttribute('scaled') ?? null);
  const rotWithShape = parseBooleanAttribute(gradFill.getAttribute('rotWithShape'));
  const flip = parseGradientFlip(gradFill.getAttribute('flip'));
  const tileRectModel = parseGradientTileRect(tileRect);
  return {
    colors,
    ...(angle !== undefined ? { angle: angle / 60000 } : {}),
    ...(flip ? { flip } : {}),
    ...(rotWithShape !== undefined ? { rotWithShape } : {}),
    ...(scaled !== undefined ? { scaled } : {}),
    ...(tileRectModel ? { tileRect: tileRectModel } : {}),
    ...(positions.some((position) => position !== undefined)
      ? { positions: positions.map((position) => position ?? 0) }
      : {}),
    ...(opacities.some((opacity) => opacity !== undefined)
      ? { opacities: opacities.map((opacity) => opacity ?? 1) }
      : {}),
  };
}

function parseBooleanAttribute(raw: string | null): boolean | undefined {
  if (raw === null || raw === '') return undefined;
  return raw === '1' || raw === 'true';
}

function parseGradientFlip(raw: string | null): PptGradientFillModel['flip'] | undefined {
  if (raw === 'none' || raw === 'x' || raw === 'y' || raw === 'xy') return raw;
  return undefined;
}

function parseGradientTileRect(
  tileRect: Element | null,
): NonNullable<PptGradientFillModel['tileRect']> | undefined {
  if (!tileRect) return undefined;
  const model: NonNullable<PptGradientFillModel['tileRect']> = {};
  const left = parsePercentageAttribute(tileRect.getAttribute('l'));
  const top = parsePercentageAttribute(tileRect.getAttribute('t'));
  const right = parsePercentageAttribute(tileRect.getAttribute('r'));
  const bottom = parsePercentageAttribute(tileRect.getAttribute('b'));
  if (left !== undefined) model.left = left;
  if (top !== undefined) model.top = top;
  if (right !== undefined) model.right = right;
  if (bottom !== undefined) model.bottom = bottom;
  return Object.keys(model).length > 0 ? model : {};
}

function parsePercentageAttribute(raw: string | null): number | undefined {
  const value = parseNumber(raw);
  return value === undefined ? undefined : value / 1000;
}

function colorFromColorElement(
  element: Element | null,
  context: PptParseContext = {},
): string | undefined {
  if (!element) return undefined;
  const schemeValue = element.getAttribute('val') ?? '';
  const raw =
    element.localName === 'srgbClr'
      ? element.getAttribute('val')
      : element.localName === 'sysClr'
        ? (element.getAttribute('lastClr') ?? element.getAttribute('val'))
        : element.localName === 'prstClr'
          ? presetColorToHex(element.getAttribute('val') ?? '')
          : (context.colorScheme?.[schemeValue] ??
            context.colorScheme?.[defaultColorSchemeAlias(schemeValue)]);
  if (!raw) return undefined;
  return applyColorTransforms(`#${raw.replace(/^#/, '').toUpperCase()}`, element);
}

function presetColorToHex(value: string): string | undefined {
  const colors: Record<string, string> = {
    black: '000000',
    white: 'FFFFFF',
    red: 'FF0000',
    green: '008000',
    blue: '0000FF',
    yellow: 'FFFF00',
  };
  return colors[value];
}

function cssColorFromColorElement(
  element: Element | null,
  context: PptParseContext = {},
): string | undefined {
  if (!element) return undefined;
  const color = colorFromColorElement(element, context);
  if (!color) return undefined;
  const alpha = parseNumber(firstChildNs(element, A_NS, 'alpha')?.getAttribute('val') ?? null);
  if (alpha === undefined || alpha >= 100000) return color;
  const [r, g, b] = [1, 3, 5].map((index) => Number.parseInt(color.slice(index, index + 2), 16));
  return `rgba(${r}, ${g}, ${b}, ${Math.max(0, Math.min(1, alpha / 100000))})`;
}

function defaultColorSchemeAlias(value: string): string {
  if (value === 'bg1') return 'lt1';
  if (value === 'tx1') return 'dk1';
  if (value === 'bg2') return 'lt2';
  if (value === 'tx2') return 'dk2';
  return value;
}

function applyColorTransforms(color: string, element: Element): string {
  const lumMod = parseNumber(firstChildNs(element, A_NS, 'lumMod')?.getAttribute('val') ?? null);
  const lumOff = parseNumber(firstChildNs(element, A_NS, 'lumOff')?.getAttribute('val') ?? null);
  const shade = parseNumber(firstChildNs(element, A_NS, 'shade')?.getAttribute('val') ?? null);
  const tint = parseNumber(firstChildNs(element, A_NS, 'tint')?.getAttribute('val') ?? null);
  if (lumMod === undefined && lumOff === undefined && shade === undefined && tint === undefined)
    return color;
  if ((lumMod !== undefined || lumOff !== undefined) && shade === undefined && tint === undefined) {
    return applyLuminanceTransforms(color, lumMod, lumOff);
  }
  const channels = [1, 3, 5].map((index) => Number.parseInt(color.slice(index, index + 2), 16));
  const transformed = channels.map((channel) => {
    const shaded = shade === undefined ? channel : channel * (shade / 100000);
    const tinted = tint === undefined ? shaded : shaded + (255 - shaded) * (tint / 100000);
    const modded = lumMod === undefined ? tinted : tinted * (lumMod / 100000);
    const offset = lumOff === undefined ? 0 : 255 * (lumOff / 100000);
    return Math.max(0, Math.min(255, Math.round(modded + offset)));
  });
  return `#${transformed
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()}`;
}

function applyLuminanceTransforms(
  color: string,
  lumMod: number | undefined,
  lumOff: number | undefined,
): string {
  const rgb = [1, 3, 5].map((index) => Number.parseInt(color.slice(index, index + 2), 16)) as [
    number,
    number,
    number,
  ];
  const [hue, saturation, luminance] = rgbToHsl(rgb);
  const nextLuminance = clamp01(
    luminance * ((lumMod ?? 100000) / 100000) + ((lumOff ?? 0) / 100000),
  );
  return rgbToHex(hslToRgb([hue, saturation, nextLuminance]));
}

function rgbToHsl([red, green, blue]: [number, number, number]): [number, number, number] {
  const r = red / 255;
  const g = green / 255;
  const b = blue / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const luminance = (max + min) / 2;
  if (max === min) return [0, 0, luminance];
  const delta = max - min;
  const saturation =
    luminance > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;
  if (max === r) hue = (g - b) / delta + (g < b ? 6 : 0);
  else if (max === g) hue = (b - r) / delta + 2;
  else hue = (r - g) / delta + 4;
  return [hue / 6, saturation, luminance];
}

function hslToRgb([hue, saturation, luminance]: [number, number, number]): [
  number,
  number,
  number,
] {
  if (saturation === 0) {
    const channel = luminance * 255;
    return [channel, channel, channel];
  }
  const q =
    luminance < 0.5
      ? luminance * (1 + saturation)
      : luminance + saturation - luminance * saturation;
  const p = 2 * luminance - q;
  return [
    hueToRgb(p, q, hue + 1 / 3) * 255,
    hueToRgb(p, q, hue) * 255,
    hueToRgb(p, q, hue - 1 / 3) * 255,
  ];
}

function hueToRgb(p: number, q: number, hue: number): number {
  let t = hue;
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
}

function rgbToHex(channels: [number, number, number]): string {
  return `#${channels
    .map((channel) =>
      Math.max(0, Math.min(255, Math.round(channel))).toString(16).padStart(2, '0'),
    )
    .join('')
    .toUpperCase()}`;
}

function clamp01(value: number): number {
  return Math.max(0, Math.min(1, value));
}

function classifyShape(
  slide: PptSlideInternal,
  element: Element,
  hostType: 'sp' | 'pic' | 'graphicFrame' | 'grpSp',
  textElements: Element[],
  domains: {
    table?: PptTableModel;
    image?: PptImageModel;
    chart?: PptChartModel;
    media?: PptShapeMediaModel;
    oleObject?: PptOleObjectModel;
    smartArt?: PptSmartArtModel;
  },
): PptShapeKind {
  if (hostType === 'grpSp') return 'group';
  if (domains.table) return 'table';
  if (domains.media) return 'media';
  if (domains.oleObject) return 'oleObject';
  if (domains.smartArt) return 'smartArt';
  if (domains.image || hostType === 'pic') return 'picture';
  if (domains.chart || hasChartRelationship(slide, element)) return 'chart';
  if (hostType === 'graphicFrame') return 'unknown';
  if (allDescendantsNs(element, P_NS, 'ph').length > 0) return 'placeholder';
  if (textElements.length > 0) return 'textBox';
  if (parseCustomGeometry(element)) return 'geometricShape';
  const presetGeometry = parsePresetGeometry(element);
  if (
    (presetGeometry === 'line' || presetGeometry === 'straightConnector1') &&
    allDescendantsNs(element, A_NS, 'ln').length > 0
  ) {
    return 'line';
  }
  if (
    allDescendantsNs(element, A_NS, 'ln').length > 0 &&
    allDescendantsNs(element, A_NS, 'prstGeom').length === 0
  ) {
    return 'line';
  }
  return hostType === 'sp' ? 'geometricShape' : 'unknown';
}

function parseChartModel(
  pkg: Package,
  slide: PptSlideInternal,
  element: Element,
  context: PptParseContext = {},
): PptChartModel | undefined {
  const chartRef = allDescendantsLocal(element, 'chart')[0];
  const relationshipId =
    chartRef?.getAttributeNS(R_NS, 'id') ?? chartRef?.getAttribute('r:id') ?? undefined;
  if (!relationshipId) return undefined;
  const rel = slide.relationships.get(relationshipId);
  if (!rel || rel.type !== CHART_REL_TYPE) return { relationshipId };
  const targetPath = resolveRelationshipTarget(slide.part.path, rel.target);
  const part = getPart(pkg, targetPath);
  const model: PptChartModel = { relationshipId, targetPath };
  if (!part?.xmlDom) return model;
  const type = parseChartType(part.xmlDom);
  const grouping = parseChartGrouping(part.xmlDom);
  const title = parseChartTitle(part.xmlDom);
  const data = parseChartData(part.xmlDom, context);
  const plotLayout = parseChartPlotLayout(part.xmlDom);
  const axes = parseChartAxes(part.xmlDom, context);
  const seriesNames = data.series
    .map((series) => series.name)
    .filter((value): value is string => Boolean(value));
  if (type) model.type = type;
  if (grouping) model.grouping = grouping;
  if (title) model.title = title;
  if (seriesNames.length > 0) model.seriesNames = seriesNames;
  if (data.series.length > 0) model.data = data;
  if (plotLayout) model.plotLayout = plotLayout;
  if (axes) model.axes = axes;
  return model;
}

function parseChartType(doc: Document): string | undefined {
  const plotArea = allDescendantsLocal(doc, 'plotArea')[0];
  if (!plotArea) return undefined;
  for (let i = 0; i < plotArea.childNodes.length; i++) {
    const child = plotArea.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const localName = (child as Element).localName;
    if (localName.endsWith('Chart')) return localName.slice(0, -'Chart'.length);
  }
  return undefined;
}

function parseChartGrouping(doc: Document): string | undefined {
  const plotArea = allDescendantsLocal(doc, 'plotArea')[0];
  if (!plotArea) return undefined;
  for (let i = 0; i < plotArea.childNodes.length; i++) {
    const child = plotArea.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const chart = child as Element;
    if (!chart.localName.endsWith('Chart')) continue;
    const grouping = firstChildLocal(chart, 'grouping')?.getAttribute('val') ?? undefined;
    if (grouping) return grouping;
  }
  return undefined;
}

function parseChartTitle(doc: Document): string | undefined {
  const title = allDescendantsLocal(doc, 'title')[0];
  if (!title) return undefined;
  const text = allDescendantsNs(title, A_NS, 't').map(textOf).join('');
  return text || undefined;
}

function parseChartData(doc: Document, context: PptParseContext = {}): PptChartData {
  const title = parseChartTitle(doc);
  const plotArea = allDescendantsLocal(doc, 'plotArea')[0];
  const series = plotArea ? parsePlotAreaSeries(plotArea, context) : [];
  return {
    ...(title ? { title } : {}),
    series,
  };
}

function parseChartPlotLayout(doc: Document): PptChartModel['plotLayout'] | undefined {
  const plotArea = allDescendantsLocal(doc, 'plotArea')[0];
  const manualLayout = plotArea ? allDescendantsLocal(plotArea, 'manualLayout')[0] : undefined;
  if (!manualLayout) return undefined;
  const numberFromChild = (name: string) =>
    parseNumber(firstChildLocal(manualLayout, name)?.getAttribute('val') ?? null);
  const layout = {
    ...(numberFromChild('x') !== undefined ? { x: numberFromChild('x')! } : {}),
    ...(numberFromChild('y') !== undefined ? { y: numberFromChild('y')! } : {}),
    ...(numberFromChild('w') !== undefined ? { width: numberFromChild('w')! } : {}),
    ...(numberFromChild('h') !== undefined ? { height: numberFromChild('h')! } : {}),
  };
  return Object.keys(layout).length > 0 ? layout : undefined;
}

function parseChartAxes(
  doc: Document,
  context: PptParseContext = {},
): PptChartModel['axes'] | undefined {
  const plotArea = allDescendantsLocal(doc, 'plotArea')[0];
  if (!plotArea) return undefined;
  const category = parseChartAxis(childElementsLocal(plotArea, 'catAx')[0], context);
  const value = parseChartAxis(childElementsLocal(plotArea, 'valAx')[0], context);
  const axes = {
    ...(category ? { category } : {}),
    ...(value ? { value } : {}),
  };
  return Object.keys(axes).length > 0 ? axes : undefined;
}

function parseChartAxis(
  axis: Element | undefined,
  context: PptParseContext = {},
): NonNullable<PptChartModel['axes']>['value'] | undefined {
  if (!axis) return undefined;
  const scaling = firstChildLocal(axis, 'scaling');
  const min = parseNumber(firstChildLocal(scaling ?? axis, 'min')?.getAttribute('val') ?? null);
  const max = parseNumber(firstChildLocal(scaling ?? axis, 'max')?.getAttribute('val') ?? null);
  const majorUnit = parseNumber(firstChildLocal(axis, 'majorUnit')?.getAttribute('val') ?? null);
  const deleted = parseBooleanAttr(firstChildLocal(axis, 'delete')?.getAttribute('val') ?? null);
  const tickLabelPosition = firstChildLocal(axis, 'tickLblPos')?.getAttribute('val') ?? undefined;
  const majorGridlinesElement = childElementsLocal(axis, 'majorGridlines')[0];
  const majorGridlines = Boolean(majorGridlinesElement);
  const axisSpPr = firstChildLocal(axis, 'spPr');
  const gridlineSpPr = majorGridlinesElement
    ? firstChildLocal(majorGridlinesElement, 'spPr')
    : null;
  const line = parseChartLineStyle(axisSpPr ? firstChildNs(axisSpPr, A_NS, 'ln') : null, context);
  const gridline = parseChartLineStyle(
    gridlineSpPr ? firstChildNs(gridlineSpPr, A_NS, 'ln') : null,
    context,
  );
  const result = {
    ...(deleted !== undefined ? { deleted } : {}),
    ...(min !== undefined ? { min } : {}),
    ...(max !== undefined ? { max } : {}),
    ...(majorUnit !== undefined ? { majorUnit } : {}),
    ...(majorGridlines ? { majorGridlines } : {}),
    ...(line.color ? { lineColor: line.color } : {}),
    ...(line.width !== undefined ? { lineWidth: line.width } : {}),
    ...(gridline.color ? { gridlineColor: gridline.color } : {}),
    ...(gridline.width !== undefined ? { gridlineWidth: gridline.width } : {}),
    ...(tickLabelPosition ? { tickLabelPosition } : {}),
  };
  return Object.keys(result).length > 0 ? result : undefined;
}

function parsePlotAreaSeries(
  plotArea: Element,
  context: PptParseContext = {},
): PptChartSeriesData[] {
  const series: PptChartSeriesData[] = [];
  let chartGroup = 0;
  for (let i = 0; i < plotArea.childNodes.length; i++) {
    const child = plotArea.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const chart = child as Element;
    if (!chart.localName.endsWith('Chart')) continue;
    const type = chart.localName.slice(0, -'Chart'.length);
    const chartLabels = parseChartDataLabelSettings(chart);
    for (const ser of childElementsLocal(chart, 'ser')) {
      const categories = parseChartSeriesCategories(ser);
      const name = parseChartSeriesName(ser);
      const spPr = firstChildLocal(ser, 'spPr');
      const fillColor = spPr
        ? colorFromSolidFill(firstChildNs(spPr, A_NS, 'solidFill'), context)
        : undefined;
      const fillGradient = spPr
        ? parseGradientFill(firstChildNs(spPr, A_NS, 'gradFill'), context)
        : undefined;
      const pointColors = parseChartPointColors(ser, context);
      const ln = spPr ? firstChildNs(spPr, A_NS, 'ln') : undefined;
      const lineStyle = parseChartLineStyle(ln, context);
      const lineGradient = ln
        ? parseGradientFill(firstChildNs(ln, A_NS, 'gradFill'), context)
        : undefined;
      const lineColor = lineStyle.color ?? lineGradient?.colors[0];
      const marker = parseChartSeriesMarker(ser, context);
      const seriesLabels = mergeChartDataLabelSettings(
        chartLabels,
        parseChartDataLabelSettings(ser),
      );
      const dataLabelLayouts = [
        ...parseChartDataLabelLayouts(chart),
        ...parseChartDataLabelLayouts(ser),
      ];
      const smooth = parseBooleanAttr(firstChildLocal(ser, 'smooth')?.getAttribute('val') ?? null);
      series.push({
        ...(name ? { name } : {}),
        type,
        chartGroup,
        ...(categories.length > 0 ? { categories } : {}),
        values: parseChartSeriesValues(ser),
        ...(fillColor ? { fillColor } : {}),
        ...(fillGradient ? { fillGradient } : {}),
        ...(pointColors.length > 0 ? { pointColors } : {}),
        ...(lineStyle.color ? { borderColor: lineStyle.color } : {}),
        ...(lineStyle.width !== undefined ? { borderWidth: lineStyle.width } : {}),
        ...(lineColor ? { lineColor } : {}),
        ...(lineStyle.width !== undefined ? { lineWidth: lineStyle.width } : {}),
        ...(marker ? { marker } : {}),
        ...(seriesLabels.showValueLabels !== undefined
          ? { showValueLabels: seriesLabels.showValueLabels }
          : {}),
        ...(seriesLabels.showCategoryNameLabels !== undefined
          ? { showCategoryNameLabels: seriesLabels.showCategoryNameLabels }
          : {}),
        ...(seriesLabels.showPercentLabels !== undefined
          ? { showPercentLabels: seriesLabels.showPercentLabels }
          : {}),
        ...(dataLabelLayouts.length > 0 ? { dataLabelLayouts } : {}),
        ...(smooth !== undefined ? { smooth } : {}),
      });
    }
    chartGroup++;
  }
  return series;
}

function parseChartLineStyle(
  ln: Element | null | undefined,
  context: PptParseContext = {},
): { color?: string; width?: number } {
  if (!ln || firstChildNs(ln, A_NS, 'noFill')) return {};
  const solidFill = firstChildNs(ln, A_NS, 'solidFill');
  const gradient = parseGradientFill(firstChildNs(ln, A_NS, 'gradFill'), context);
  const color = colorFromSolidFill(solidFill, context) ?? gradient?.colors[0];
  const width = parseNumber(ln.getAttribute('w') ?? null);
  return {
    ...(color ? { color } : {}),
    ...(width !== undefined ? { width } : {}),
  };
}

function parseChartPointColors(ser: Element, context: PptParseContext = {}): string[] {
  const colors: string[] = [];
  for (const point of childElementsLocal(ser, 'dPt')) {
    const index = parseNumber(firstChildLocal(point, 'idx')?.getAttribute('val') ?? null);
    const spPr = firstChildLocal(point, 'spPr');
    const color = spPr
      ? colorFromSolidFill(firstChildNs(spPr, A_NS, 'solidFill'), context)
      : undefined;
    if (index === undefined || !color) continue;
    colors[index] = color;
  }
  return colors;
}

function parseChartDataLabelSettings(element: Element): {
  showValueLabels?: boolean;
  showCategoryNameLabels?: boolean;
  showPercentLabels?: boolean;
} {
  const dLbls = firstChildLocal(element, 'dLbls');
  if (!dLbls) return {};
  const deleted = parseBooleanAttr(firstChildLocal(dLbls, 'delete')?.getAttribute('val') ?? null);
  if (deleted === true) {
    return {
      showValueLabels: false,
      showCategoryNameLabels: false,
      showPercentLabels: false,
    };
  }
  return {
    ...chartDataLabelBoolean(dLbls, 'showVal', 'showValueLabels'),
    ...chartDataLabelBoolean(dLbls, 'showCatName', 'showCategoryNameLabels'),
    ...chartDataLabelBoolean(dLbls, 'showPercent', 'showPercentLabels'),
  };
}

function parseChartDataLabelLayouts(element: Element): NonNullable<PptChartSeriesData['dataLabelLayouts']> {
  const dLbls = firstChildLocal(element, 'dLbls');
  if (!dLbls) return [];
  const result: NonNullable<PptChartSeriesData['dataLabelLayouts']> = [];
  for (const label of childElementsLocal(dLbls, 'dLbl')) {
    const index = parseNumber(firstChildLocal(label, 'idx')?.getAttribute('val') ?? null);
    const manualLayout = allDescendantsLocal(label, 'manualLayout')[0];
    if (index === undefined || !manualLayout) continue;
    const x = parseNumber(firstChildLocal(manualLayout, 'x')?.getAttribute('val') ?? null);
    const y = parseNumber(firstChildLocal(manualLayout, 'y')?.getAttribute('val') ?? null);
    if (x === undefined && y === undefined) continue;
    result.push({
      index,
      ...(x !== undefined ? { x } : {}),
      ...(y !== undefined ? { y } : {}),
    });
  }
  return result;
}

function chartDataLabelBoolean<K extends 'showValueLabels' | 'showCategoryNameLabels' | 'showPercentLabels'>(
  dLbls: Element,
  xmlName: string,
  key: K,
): Partial<Record<K, boolean>> {
  const value = parseBooleanAttr(firstChildLocal(dLbls, xmlName)?.getAttribute('val') ?? null);
  return value === undefined ? {} : { [key]: value } as Partial<Record<K, boolean>>;
}

function mergeChartDataLabelSettings(
  base: {
    showValueLabels?: boolean;
    showCategoryNameLabels?: boolean;
    showPercentLabels?: boolean;
  },
  override: {
    showValueLabels?: boolean;
    showCategoryNameLabels?: boolean;
    showPercentLabels?: boolean;
  },
): {
  showValueLabels?: boolean;
  showCategoryNameLabels?: boolean;
  showPercentLabels?: boolean;
} {
  return {
    ...base,
    ...override,
  };
}

function parseChartSeriesMarker(
  ser: Element,
  context: PptParseContext = {},
): PptChartSeriesData['marker'] | undefined {
  const marker = firstChildLocal(ser, 'marker');
  if (!marker) return undefined;
  const symbol = firstChildLocal(marker, 'symbol')?.getAttribute('val') ?? undefined;
  const size = parseNumber(firstChildLocal(marker, 'size')?.getAttribute('val') ?? null);
  const spPr = firstChildLocal(marker, 'spPr');
  const fillColor = spPr
    ? colorFromSolidFill(firstChildNs(spPr, A_NS, 'solidFill'), context)
    : undefined;
  const line = parseChartLineStyle(spPr ? firstChildNs(spPr, A_NS, 'ln') : null, context);
  return symbol || size !== undefined || fillColor || line.color || line.width !== undefined
    ? {
        ...(symbol ? { symbol } : {}),
        ...(size !== undefined ? { size } : {}),
        ...(fillColor ? { fillColor } : {}),
        ...(line.color ? { lineColor: line.color } : {}),
        ...(line.width !== undefined ? { lineWidth: line.width } : {}),
      }
    : undefined;
}

function childElementsLocal(parent: Element, localName: string): Element[] {
  const result: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const element = child as Element;
    if (element.localName === localName) result.push(element);
  }
  return result;
}

function parseChartSeriesName(ser: Element): string | undefined {
  const tx = firstChildLocal(ser, 'tx');
  if (!tx) return undefined;
  const inline = allDescendantsLocal(tx, 'v').map(textOf).join('');
  return valuesFromStringCache(tx)[0] ?? (inline || undefined);
}

function parseChartSeriesCategories(ser: Element): string[] {
  const cat = firstChildLocal(ser, 'cat') ?? firstChildLocal(ser, 'xVal');
  if (!cat) return [];
  return valuesFromStringCache(cat);
}

function parseChartSeriesValues(ser: Element): number[] {
  const val = firstChildLocal(ser, 'val') ?? firstChildLocal(ser, 'yVal');
  if (!val) return [];
  return valuesFromNumberCache(val);
}

function valuesFromStringCache(parent: Element): string[] {
  const cache = firstChildLocal(parent, 'strCache') ?? allDescendantsLocal(parent, 'strCache')[0];
  if (!cache) return [];
  return allDescendantsLocal(cache, 'pt')
    .map((pt) => allDescendantsLocal(pt, 'v')[0])
    .map(textOf)
    .filter(Boolean);
}

function valuesFromNumberCache(parent: Element): number[] {
  const cache = firstChildLocal(parent, 'numCache') ?? allDescendantsLocal(parent, 'numCache')[0];
  if (!cache) return [];
  return allDescendantsLocal(cache, 'pt')
    .map((pt) => Number(textOf(allDescendantsLocal(pt, 'v')[0])))
    .filter((value) => Number.isFinite(value));
}

function hasChartRelationship(slide: PptSlideInternal, element: Element): boolean {
  for (const el of allDescendantsLocal(element, 'chart')) {
    const relId = el.getAttributeNS(R_NS, 'id') ?? el.getAttribute('r:id');
    if (relId && slide.relationships.get(relId)?.type === CHART_REL_TYPE) return true;
  }
  return false;
}

function parseMediaModel(
  pkg: Package,
  slide: PptSlideInternal,
  element: Element,
): PptShapeMediaModel | undefined {
  const mediaRef =
    allDescendantsLocal(element, 'videoFile')[0] ??
    allDescendantsLocal(element, 'audioFile')[0] ??
    allDescendantsLocal(element, 'media')[0];
  const relId =
    mediaRef?.getAttributeNS(R_NS, 'link') ??
    mediaRef?.getAttribute('r:link') ??
    mediaRef?.getAttributeNS(R_NS, 'embed') ??
    mediaRef?.getAttribute('r:embed') ??
    undefined;
  if (!relId) return undefined;
  const rel = slide.relationships.get(relId);
  const mediaType = mediaTypeFromReference(mediaRef, rel?.type);
  const targetPath =
    rel && isMediaRelationship(rel.type)
      ? resolveRelationshipTarget(slide.part.path, rel.target)
      : undefined;
  const model: PptShapeMediaModel = { relationshipId: relId };
  if (mediaType) model.mediaType = mediaType;
  if (targetPath) {
    model.targetPath = targetPath;
    model.filename = targetPath.slice(targetPath.lastIndexOf('/') + 1);
    const part = getPart(pkg, targetPath);
    if (part?.contentType) model.contentType = part.contentType;
  }
  return model;
}

function parseOleObjectModel(
  slide: PptSlideInternal,
  element: Element,
): PptOleObjectModel | undefined {
  const ole = allDescendantsLocal(element, 'oleObj')[0];
  if (!ole) return undefined;
  const relId = ole.getAttributeNS(R_NS, 'id') ?? ole.getAttribute('r:id') ?? undefined;
  const model: PptOleObjectModel = {};
  const progId = ole.getAttribute('progId') ?? undefined;
  const name = ole.getAttribute('name') ?? undefined;
  if (relId) {
    model.relationshipId = relId;
    const rel = slide.relationships.get(relId);
    if (rel?.type === OLE_OBJECT_REL_TYPE) {
      model.targetPath = resolveRelationshipTarget(slide.part.path, rel.target);
    }
  }
  if (progId) model.progId = progId;
  if (name) model.name = name;
  return Object.keys(model).length > 0 ? model : undefined;
}

function parseSmartArtModel(
  slide: PptSlideInternal,
  element: Element,
): PptSmartArtModel | undefined {
  const graphicData = allDescendantsNs(element, A_NS, 'graphicData').find((el) =>
    (el.getAttribute('uri') ?? '') === DIAGRAM_URI,
  );
  if (!graphicData) return undefined;
  const relIds = allDescendantsLocal(graphicData, 'relIds')[0];
  if (!relIds) return {};
  const relationshipIds: Record<string, string> = {};
  const targetPaths: Record<string, string> = {};
  for (const attrName of ['dm', 'lo', 'qs', 'cs']) {
    const relId = relIds.getAttributeNS(R_NS, attrName) ?? relIds.getAttribute(`r:${attrName}`);
    if (!relId) continue;
    relationshipIds[attrName] = relId;
    const rel = slide.relationships.get(relId);
    if (rel) targetPaths[attrName] = resolveRelationshipTarget(slide.part.path, rel.target);
  }
  return {
    ...(Object.keys(relationshipIds).length > 0 ? { relationshipIds } : {}),
    ...(Object.keys(targetPaths).length > 0 ? { targetPaths } : {}),
  };
}

function isMediaRelationship(type: string): boolean {
  return type === MEDIA_REL_TYPE || type === VIDEO_REL_TYPE || type === AUDIO_REL_TYPE;
}

function mediaTypeFromReference(
  ref: Element | undefined,
  relationshipType: string | undefined,
): PptShapeMediaModel['mediaType'] | undefined {
  if (ref?.localName === 'videoFile' || relationshipType === VIDEO_REL_TYPE) return 'video';
  if (ref?.localName === 'audioFile' || relationshipType === AUDIO_REL_TYPE) return 'audio';
  if (relationshipType === MEDIA_REL_TYPE || ref?.localName === 'media') return 'media';
  return undefined;
}

function parseGeometry(element: Element): PptShapeGeometry {
  const xfrm = allDescendantsLocal(element, 'xfrm')[0];
  if (!xfrm) return {};
  const off = firstChildNs(xfrm, A_NS, 'off');
  const ext = firstChildNs(xfrm, A_NS, 'ext');
  const geometry: PptShapeGeometry = {};
  const left = parseNumber(off?.getAttribute('x') ?? null);
  const top = parseNumber(off?.getAttribute('y') ?? null);
  const width = parseNumber(ext?.getAttribute('cx') ?? null);
  const height = parseNumber(ext?.getAttribute('cy') ?? null);
  const rotation = parseNumber(xfrm.getAttribute('rot'));
  const flipHorizontal = parseBooleanAttr(xfrm.getAttribute('flipH'));
  const flipVertical = parseBooleanAttr(xfrm.getAttribute('flipV'));
  if (left !== undefined) geometry.left = left;
  if (top !== undefined) geometry.top = top;
  if (width !== undefined) geometry.width = width;
  if (height !== undefined) geometry.height = height;
  if (rotation !== undefined) geometry.rotation = rotation / 60000;
  if (flipHorizontal !== undefined) geometry.flipHorizontal = flipHorizontal;
  if (flipVertical !== undefined) geometry.flipVertical = flipVertical;
  return geometry;
}

function parseGroupTransform(element: Element): PptGroupTransformModel | undefined {
  const xfrm = allDescendantsLocal(element, 'xfrm')[0];
  if (!xfrm) return undefined;
  const chOff = firstChildNs(xfrm, A_NS, 'chOff');
  const chExt = firstChildNs(xfrm, A_NS, 'chExt');
  const childLeft = parseNumber(chOff?.getAttribute('x') ?? null);
  const childTop = parseNumber(chOff?.getAttribute('y') ?? null);
  const childWidth = parseNumber(chExt?.getAttribute('cx') ?? null);
  const childHeight = parseNumber(chExt?.getAttribute('cy') ?? null);
  if (
    childLeft === undefined &&
    childTop === undefined &&
    childWidth === undefined &&
    childHeight === undefined
  ) {
    return undefined;
  }
  return {
    ...(childLeft !== undefined ? { childLeft } : {}),
    ...(childTop !== undefined ? { childTop } : {}),
    ...(childWidth !== undefined ? { childWidth } : {}),
    ...(childHeight !== undefined ? { childHeight } : {}),
  };
}

function imageModelFromElement(
  slide: PptSlideInternal,
  element: Element,
): PptImageModel | undefined {
  const blip = allDescendantsNs(element, A_NS, 'blip')[0];
  const svgBlip = blip ? allDescendantsLocal(blip, 'svgBlip')[0] : undefined;
  const relId =
    blip?.getAttributeNS(R_NS, 'embed') ??
    blip?.getAttribute('r:embed') ??
    svgBlip?.getAttributeNS(R_NS, 'embed') ??
    svgBlip?.getAttribute('r:embed');
  if (!relId) return undefined;
  const rel = slide.relationships.get(relId);
  const targetPath =
    rel && rel.type === IMAGE_REL_TYPE
      ? resolveRelationshipTarget(slide.part.path, rel.target)
      : undefined;
  const model: PptImageModel = { relationshipId: relId };
  const crop = parseImageCrop(element);
  const fillRect = parseImageFillRect(element);
  const opacity = parseImageOpacity(blip);
  if (targetPath) {
    model.targetPath = targetPath;
    model.filename = targetPath.slice(targetPath.lastIndexOf('/') + 1);
  }
  if (crop) model.crop = crop;
  if (fillRect) model.fillRect = fillRect;
  if (opacity !== undefined) model.opacity = opacity;
  return model;
}

function parseImageCrop(element: Element): PptImageModel['crop'] | undefined {
  const srcRect = allDescendantsNs(element, A_NS, 'srcRect')[0];
  if (!srcRect) return undefined;
  const crop = {
    ...(parseNumber(srcRect.getAttribute('l')) !== undefined
      ? { left: parseNumber(srcRect.getAttribute('l'))! / 100000 }
      : {}),
    ...(parseNumber(srcRect.getAttribute('t')) !== undefined
      ? { top: parseNumber(srcRect.getAttribute('t'))! / 100000 }
      : {}),
    ...(parseNumber(srcRect.getAttribute('r')) !== undefined
      ? { right: parseNumber(srcRect.getAttribute('r'))! / 100000 }
      : {}),
    ...(parseNumber(srcRect.getAttribute('b')) !== undefined
      ? { bottom: parseNumber(srcRect.getAttribute('b'))! / 100000 }
      : {}),
  };
  return Object.keys(crop).length > 0 ? crop : undefined;
}

function parseImageFillRect(element: Element): PptImageModel['fillRect'] | undefined {
  const stretch = allDescendantsNs(element, A_NS, 'stretch')[0];
  const fillRect = stretch ? firstChildNs(stretch, A_NS, 'fillRect') : null;
  if (!fillRect) return undefined;
  const rect = {
    ...(parseNumber(fillRect.getAttribute('l')) !== undefined
      ? { left: parseNumber(fillRect.getAttribute('l'))! / 100000 }
      : {}),
    ...(parseNumber(fillRect.getAttribute('t')) !== undefined
      ? { top: parseNumber(fillRect.getAttribute('t'))! / 100000 }
      : {}),
    ...(parseNumber(fillRect.getAttribute('r')) !== undefined
      ? { right: parseNumber(fillRect.getAttribute('r'))! / 100000 }
      : {}),
    ...(parseNumber(fillRect.getAttribute('b')) !== undefined
      ? { bottom: parseNumber(fillRect.getAttribute('b'))! / 100000 }
      : {}),
  };
  return Object.keys(rect).length > 0 ? rect : undefined;
}

function parseImageOpacity(blip: Element | undefined): number | undefined {
  if (!blip) return undefined;
  const alphaModFix = firstChildNs(blip, A_NS, 'alphaModFix');
  const amount = parseNumber(alphaModFix?.getAttribute('amt') ?? null);
  return amount === undefined ? undefined : Math.max(0, Math.min(1, amount / 100000));
}

function parseTableCells(element: Element): Element[][] {
  const graphicData = allDescendantsNs(element, A_NS, 'graphicData').find(
    (el) => el.getAttribute('uri') === TABLE_URI,
  );
  if (!graphicData) return [];
  const table = firstChildNs(graphicData, A_NS, 'tbl');
  if (!table) return [];
  return childElementsNs(table, A_NS, 'tr').map((row) => childElementsNs(row, A_NS, 'tc'));
}

function tableModelFromCells(
  element: Element,
  cells: Element[][],
  context: PptParseContext = {},
): PptTableModel {
  const values = cells.map((row) => row.map((cell) => tableCellText(cell)));
  const cellStyles = cells.map((row) => row.map((cell) => parseTableCellStyle(cell, context)));
  const mergedCells = parseTableCellMerges(cells);
  const table = allDescendantsNs(element, A_NS, 'tbl')[0];
  const columnWidths = table ? parseTableColumnWidths(table) : [];
  const rowHeights = table ? parseTableRowHeights(table) : [];
  return {
    values,
    rowCount: values.length,
    columnCount: values.reduce((max, row) => Math.max(max, row.length), 0),
    ...(columnWidths.length > 0 ? { columnWidths } : {}),
    ...(rowHeights.length > 0 ? { rowHeights } : {}),
    ...(cellStyles.some((row) => row.some((style) => Object.keys(style).length > 0))
      ? { cellStyles }
      : {}),
    ...(mergedCells.length > 0 ? { mergedCells } : {}),
  };
}

function tableCellText(cell: Element): string {
  const paragraphs = allDescendantsNs(cell, A_NS, 'p').map((paragraph) =>
    allDescendantsNs(paragraph, A_NS, 't').map(textOf).join(''),
  );
  while (paragraphs.length > 0 && paragraphs[paragraphs.length - 1] === '') paragraphs.pop();
  if (paragraphs.length > 0) return paragraphs.join('\n');
  return allDescendantsNs(cell, A_NS, 't').map(textOf).join('');
}

function parseTableColumnWidths(table: Element): number[] {
  const grid = firstChildNs(table, A_NS, 'tblGrid');
  if (!grid) return [];
  return childElementsNs(grid, A_NS, 'gridCol')
    .map((column) => parseNumber(column.getAttribute('w')))
    .filter((value): value is number => value !== undefined);
}

function parseTableRowHeights(table: Element): number[] {
  return childElementsNs(table, A_NS, 'tr')
    .map((row) => parseNumber(row.getAttribute('h')))
    .filter((value): value is number => value !== undefined);
}

function parseTableCellMerges(cells: Element[][]): PptTableCellMerge[] {
  const result: PptTableCellMerge[] = [];
  for (let row = 0; row < cells.length; row++) {
    for (let column = 0; column < cells[row]!.length; column++) {
      const cell = cells[row]![column]!;
      const rowSpan = parseNumber(cell.getAttribute('rowSpan'));
      const columnSpan = parseNumber(cell.getAttribute('gridSpan'));
      const hMerge = parseBooleanAttr(cell.getAttribute('hMerge'));
      const vMerge = parseBooleanAttr(cell.getAttribute('vMerge'));
      if (
        rowSpan === undefined &&
        columnSpan === undefined &&
        hMerge === undefined &&
        vMerge === undefined
      ) {
        continue;
      }
      result.push({
        row,
        column,
        ...(rowSpan !== undefined ? { rowSpan } : {}),
        ...(columnSpan !== undefined ? { columnSpan } : {}),
        ...(hMerge !== undefined ? { hMerge } : {}),
        ...(vMerge !== undefined ? { vMerge } : {}),
      });
    }
  }
  return result;
}

function parseTableCellStyle(cell: Element, context: PptParseContext = {}): PptTableCellStyle {
  const tcPr = firstChildNs(cell, A_NS, 'tcPr');
  const fillColor = colorFromSolidFill(
    tcPr ? firstChildNs(tcPr, A_NS, 'solidFill') : null,
    context,
  );
  const fillGradient = parseGradientFill(
    tcPr ? firstChildNs(tcPr, A_NS, 'gradFill') : null,
    context,
  );
  const borders = tcPr ? parseTableCellBorders(tcPr, context) : undefined;
  const body = tcPr ? parseTableCellBodyProperties(tcPr) : undefined;
  const textStyle = parseTextStyle(cell, context);
  return {
    ...(fillColor ? { fillColor } : {}),
    ...(fillGradient ? { fillGradient } : {}),
    ...(borders ? { borders } : {}),
    ...(body ? { body } : {}),
    ...(Object.keys(textStyle).length > 0 ? { textStyle } : {}),
  };
}

function parseTableCellBorders(
  tcPr: Element,
  context: PptParseContext,
): PptTableCellStyle['borders'] | undefined {
  const borders = {
    left: parseTableCellBorder(firstChildNs(tcPr, A_NS, 'lnL'), context),
    right: parseTableCellBorder(firstChildNs(tcPr, A_NS, 'lnR'), context),
    top: parseTableCellBorder(firstChildNs(tcPr, A_NS, 'lnT'), context),
    bottom: parseTableCellBorder(firstChildNs(tcPr, A_NS, 'lnB'), context),
  };
  return Object.values(borders).some(Boolean) ? borders : undefined;
}

function parseTableCellBorder(
  line: Element | null,
  context: PptParseContext,
): NonNullable<PptTableCellStyle['borders']>['left'] | undefined {
  if (!line) return undefined;
  if (firstChildNs(line, A_NS, 'noFill')) {
    return { hidden: true };
  }
  const fill = parseSolidFill(firstChildNs(line, A_NS, 'solidFill'), context);
  const width = parseNumber(line.getAttribute('w') ?? null);
  const dash = parseLineDash(line);
  return {
    ...(fill.color ? { color: fill.color } : {}),
    ...(fill.opacity !== undefined ? { opacity: fill.opacity } : {}),
    ...(width !== undefined ? { width } : {}),
    ...(dash && dash !== 'solid' ? { dash } : {}),
  };
}

function parseTableCellBodyProperties(tcPr: Element): PptTextBodyProperties | undefined {
  const anchor = parseTextBodyAnchor(tcPr.getAttribute('anchor'));
  const wrap = parseTextBodyWrap(tcPr.getAttribute('wrap'));
  const insetLeft = parseNumber(tcPr.getAttribute('marL'));
  const insetTop = parseNumber(tcPr.getAttribute('marT'));
  const insetRight = parseNumber(tcPr.getAttribute('marR'));
  const insetBottom = parseNumber(tcPr.getAttribute('marB'));
  const properties: PptTextBodyProperties = {
    ...(anchor ? { anchor } : {}),
    ...(wrap ? { wrap } : {}),
    ...(insetLeft !== undefined ? { insetLeft } : {}),
    ...(insetTop !== undefined ? { insetTop } : {}),
    ...(insetRight !== undefined ? { insetRight } : {}),
    ...(insetBottom !== undefined ? { insetBottom } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

function deriveSlideTitle(shapes: PptShapeModel[]): string | undefined {
  const titleShape = shapes.find(
    (shape) => shape.kind === 'placeholder' && shape.text?.text.trim(),
  );
  if (titleShape?.text?.text.trim()) return titleShape.text.text.trim();
  const firstText = shapes.find((shape) => shape.text?.text.trim());
  return firstText?.text?.text.trim() || undefined;
}

function listMedia(pkg: Package): PptMediaModel[] {
  const media: PptMediaModel[] = [];
  for (const part of pkg.parts.values()) {
    if (!part.path.startsWith('ppt/media/')) continue;
    if (!isBinaryPart(part) && !part._originalBytes) continue;
    media.push({
      path: part.path,
      contentType: part.contentType,
      byteLength: part.bytes?.byteLength ?? part._originalBytes?.byteLength ?? 0,
    });
  }
  return media;
}

function listEmbeddedFonts(
  pkg: Package,
  presentationPath: string,
  presentationDoc: Document,
  presentationRels: Map<string, PptRelationship>,
): PptFontModel[] {
  const fonts: PptFontModel[] = [];
  for (const embeddedFont of allDescendantsNs(presentationDoc, P_NS, 'embeddedFont')) {
    const family = firstChildNs(embeddedFont, P_NS, 'font')?.getAttribute('typeface') ?? '';
    if (!family) continue;
    for (const variant of ['regular', 'bold', 'italic', 'boldItalic'] as const) {
      const ref = firstChildNs(embeddedFont, P_NS, variant);
      const relId = ref?.getAttributeNS(R_NS, 'id') ?? ref?.getAttribute('r:id') ?? undefined;
      if (!relId) continue;
      const rel = presentationRels.get(relId);
      if (!rel || rel.type !== FONT_REL_TYPE) continue;
      const path = resolveRelationshipTarget(presentationPath, rel.target);
      const part = getPart(pkg, path);
      const bytes = part?.bytes ?? part?._originalBytes;
      if (!part || !bytes) continue;
      fonts.push({
        family,
        path,
        contentType: part.contentType,
        byteLength: bytes.byteLength,
        ...(variant === 'bold' || variant === 'boldItalic' ? { weight: 700 } : { weight: 400 }),
        ...(variant === 'italic' || variant === 'boldItalic'
          ? { style: 'italic' as const }
          : { style: 'normal' as const }),
      });
    }
  }
  return fonts;
}

export function findOrCreateTransform(shape: Element): Element {
  let xfrm = allDescendantsLocal(shape, 'xfrm')[0];
  if (xfrm) return xfrm;
  const spPr = firstChildLocal(shape, 'spPr');
  if (!spPr) throw new Error('Presentation: shape has no transform container');
  xfrm = shape.ownerDocument.createElementNS(A_NS, 'a:xfrm');
  spPr.insertBefore(xfrm, spPr.firstChild);
  return xfrm;
}
