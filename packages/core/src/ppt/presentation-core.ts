import {
  createBinaryPart,
  createXmlPart,
  createPackage,
  deletePart,
  fromZip,
  getPart,
  markDirty,
  setPart,
  toZip,
  type Package,
  type Part,
} from '../pack/index.js';
import type {
  PptChartModel,
  PptChartData,
  PptChartSeriesData,
  PptCommentAddOptions,
  PptCommentModel,
  PptExportedImage,
  PptExportedFont,
  PptFindTextOptions,
  PptFontModel,
  PptHyperlinkModel,
  PptHyperlinkWriteOptions,
  PptImageInfo,
  PptImageSource,
  PptImageStyle,
  PptPresentationModel,
  PptPresentationProperties,
  PptPlaceholderModel,
  PptShapeGeometry,
  PptShapeStyle,
  PptSlideSize,
  PptShapeKind,
  PptShapeModel,
  PptSlideModel,
  PptSlideLayoutModel,
  PptSlideProperties,
  PptSlideTimingModel,
  PptSlideTransitionModel,
  PptTableCellStyle,
  PptTableCellMerge,
  PptTextBodyProperties,
  PptThemeModel,
  PptTextMatch,
  PptTextParagraph,
  PptTextParagraphWrite,
  PptTextStyle,
  PptTextRun,
  PptTextReplaceOptions,
  PptTextReplaceResult,
  PptTextWriteOptions,
} from './model.js';
import { PptMutationLog, type PptMutation } from './mutations.js';
import { resolvePptEffectivePlaceholderProperties } from './placeholder-inheritance.js';
import {
  findOrCreateTransform,
  readPresentationPackage,
  type PptReadResult,
  type PptShapeInternal,
} from './pptx-reader.js';
import { createPresentationPackageFromModel } from './pptx-writer.js';
import {
  A_NS,
  OFFICE_REL_NS,
  P_NS,
  PKG_REL_NS,
  R_NS,
  allDescendantsNs,
  allDescendantsLocal,
  childElementsNs,
  firstChildLocal,
  firstChildNs,
  parseNumber,
  parseRelationships,
  relsPathForPart,
  resolveRelationshipTarget,
  setAttrIfDefined,
  textOf,
} from './xml.js';
import { parseXmlDocument } from 'slimdom';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

export interface PresentationCoreLoadOptions {
  presentationId?: string;
}

export class PresentationCore {
  readonly presentationId: string;
  readonly slides: SlideCollectionCore;
  readonly themes: PresentationThemeCollectionCore;
  private readonly pkg: Package;
  private read: PptReadResult;
  private readonly mutations = new PptMutationLog();

  private constructor(init: { presentationId: string; pkg: Package; read: PptReadResult }) {
    this.presentationId = init.presentationId;
    this.pkg = init.pkg;
    this.read = init.read;
    this.slides = new SlideCollectionCore(this);
    this.themes = new PresentationThemeCollectionCore(this);
  }

  static async fromZip(
    bytes: Uint8Array | ArrayBuffer,
    opts: PresentationCoreLoadOptions = {},
  ): Promise<PresentationCore> {
    const pkg = await fromZip(bytes);
    return PresentationCore.fromPackage(pkg, opts);
  }

  static fromPackage(pkg: Package, opts: PresentationCoreLoadOptions = {}): PresentationCore {
    const presentationId = opts.presentationId ?? `ppt:${Date.now()}`;
    return new PresentationCore({
      presentationId,
      pkg,
      read: readPresentationPackage(pkg, presentationId),
    });
  }

  static fromModel(
    model: PptPresentationModel,
    opts: PresentationCoreLoadOptions = {},
  ): PresentationCore {
    const presentationId = opts.presentationId ?? model.id ?? `ppt:${Date.now()}`;
    const pkg = createPresentationPackageFromModel({ ...cloneModel(model), id: presentationId });
    return new PresentationCore({
      presentationId,
      pkg,
      read: readPresentationPackage(pkg, presentationId),
    });
  }

  toModel(): PptPresentationModel {
    return cloneModel(this.read.model);
  }

  properties(): PptPresentationProperties {
    const slides = this.read.model.slides;
    const shapes = slides.flatMap((slide) => slide.shapes);
    return {
      presentationId: this.presentationId,
      slideCount: slides.length,
      shapeCount: shapes.length,
      textShapeCount: shapes.filter((shape) => Boolean(shape.text) && !shape.table).length,
      tableCount: shapes.filter((shape) => Boolean(shape.table)).length,
      pictureCount: shapes.filter((shape) => shape.kind === 'picture').length,
      mediaCount: this.read.model.media?.length ?? 0,
      ...(this.read.model.slideSize ? { slideSize: { ...this.read.model.slideSize } } : {}),
    };
  }

  getMutations(): PptMutation[] {
    return this.mutations.list();
  }

  clearMutations(): void {
    this.mutations.clear();
  }

  findText(query: string | RegExp, options: PptFindTextOptions = {}): PptTextMatch[] {
    return this.read.slides.flatMap((slide) => findTextInSlide(slide, query, options));
  }

  replaceText(
    search: string | RegExp,
    replacement: string,
    options: PptTextReplaceOptions = {},
  ): PptTextReplaceResult {
    let count = 0;
    for (const slide of this.slides.items()) {
      count += slide.replaceText(search, replacement, options).count;
    }
    return { count };
  }

  exportImages(): PptExportedImage[] {
    return this.read.model.media?.map((media) => exportedImageFromPart(this, media.path)) ?? [];
  }

  exportFonts(): PptExportedFont[] {
    return this.read.model.fonts?.map((font) => exportedFontFromPart(this, font)) ?? [];
  }

  async toZip(): Promise<Uint8Array> {
    return toZip(this.pkg);
  }

  _slides(): PptSlideModel[] {
    return this.read.model.slides;
  }

  _getSlide(indexOrId: number | string): PptSlideModel | null {
    if (typeof indexOrId === 'number') return this.read.model.slides[indexOrId] ?? null;
    return (
      this.read.model.slides.find((slide) => slide.id === indexOrId || slide.name === indexOrId) ??
      null
    );
  }

  _getSlideInternal(slideId: string): PptReadResult['slides'][number] | null {
    return this.read.slides.find((slide) => slide.model.id === slideId) ?? null;
  }

  _record(mutation: PptMutation): void {
    this.mutations.record(mutation);
  }

  _package(): Package {
    return this.pkg;
  }

  _reloadFromPackage(): void {
    this.read = readPresentationPackage(this.pkg, this.presentationId);
  }
}

export class PresentationThemeCollectionCore {
  constructor(private readonly presentation: PresentationCore) {}

  items(): PptThemeModel[] {
    return cloneModel(this.presentation.toModel().themes ?? []);
  }

  getCount(): number {
    return this.presentation.toModel().themes?.length ?? 0;
  }
}

export class SlideCollectionCore {
  constructor(private readonly presentation: PresentationCore) {}

  items(): SlideCore[] {
    return this.presentation._slides().map((slide) => new SlideCore(this.presentation, slide.id));
  }

  getCount(): number {
    return this.presentation._slides().length;
  }

  getItem(indexOrId: number | string): SlideCore {
    const slide = this.presentation._getSlide(indexOrId);
    if (!slide) throw new Error(`Presentation: slide not found: ${String(indexOrId)}`);
    return new SlideCore(this.presentation, slide.id);
  }

  getItemOrNullObject(indexOrId: number | string): SlideCore | null {
    const slide = this.presentation._getSlide(indexOrId);
    return slide ? new SlideCore(this.presentation, slide.id) : null;
  }

  add(options: PptSlideAddOptions = {}): SlideCore {
    const slides = this.presentation._slides();
    const insertIndex = normalizeInsertIndex(options.index ?? 'end', slides.length);
    const nextNumber = nextSlidePartNumber(slides);
    const slidePath = `ppt/slides/slide${nextNumber}.xml`;
    const slideId = String(nextSlideId(slides));
    const relId = nextRelationshipId(this.presentation);
    const model: PptSlideModel = {
      id: slideId,
      index: insertIndex,
      relationshipId: relId,
      path: slidePath,
      shapes: [],
    };
    if (options.title) {
      model.title = options.title;
      model.shapes.push({
        id: '2',
        name: 'Title',
        kind: 'textBox',
        geometry: defaultTitleGeometry(this.presentation.toModel().slideSize),
        text: { text: options.title },
      });
    }
    setPart(
      this.presentation._package(),
      createXmlPart({
        path: slidePath,
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
        xmlDom: parseXmlDocument(slideXmlFromModel(model)) as unknown as Document,
      }),
    );
    addPresentationRelationship(this.presentation, relId, `slides/slide${nextNumber}.xml`);
    addSlideId(this.presentation, insertIndex, slideId, relId);
    deletePart(this.presentation._package(), '[Content_Types].xml');
    this.presentation._reloadFromPackage();
    return new SlideCore(this.presentation, slideId);
  }

  copy(indexOrId: number | string, options: PptSlideCopyOptions = {}): SlideCore {
    const source = this.presentation._getSlide(indexOrId);
    if (!source?.path) throw new Error(`Presentation: slide not found: ${String(indexOrId)}`);
    const sourcePart = getPart(this.presentation._package(), source.path);
    if (!sourcePart?.xmlDom)
      throw new Error(`Presentation: source slide has no XML part: ${source.path}`);
    const slides = this.presentation._slides();
    const insertIndex = normalizeInsertIndex(options.index ?? source.index + 1, slides.length);
    const nextNumber = nextSlidePartNumber(slides);
    const slidePath = `ppt/slides/slide${nextNumber}.xml`;
    const slideId = String(nextSlideId(slides));
    const relId = nextRelationshipId(this.presentation);
    setPart(
      this.presentation._package(),
      createXmlPart({
        path: slidePath,
        contentType: sourcePart.contentType,
        xmlDom: sourcePart.xmlDom.cloneNode(true) as Document,
      }),
    );
    const sourceRels = getPart(this.presentation._package(), relsPathForPart(source.path));
    if (sourceRels?.xmlDom) {
      setPart(
        this.presentation._package(),
        createXmlPart({
          path: relsPathForPart(slidePath),
          contentType: sourceRels.contentType,
          xmlDom: sourceRels.xmlDom.cloneNode(true) as Document,
        }),
      );
    }
    addPresentationRelationship(this.presentation, relId, `slides/slide${nextNumber}.xml`);
    addSlideId(this.presentation, insertIndex, slideId, relId);
    deletePart(this.presentation._package(), '[Content_Types].xml');
    this.presentation._reloadFromPackage();
    return new SlideCore(this.presentation, slideId);
  }

  move(indexOrId: number | string, toIndex: number): void {
    const slide = this.presentation._getSlide(indexOrId);
    if (!slide?.relationshipId)
      throw new Error(`Presentation: slide not found: ${String(indexOrId)}`);
    const sldIdLst = getSlideIdList(this.presentation);
    const nodes = allDescendantsNs(sldIdLst, P_NS, 'sldId');
    const target = nodes.find(
      (node) =>
        (node.getAttributeNS(R_NS, 'id') ?? node.getAttribute('r:id')) === slide.relationshipId,
    );
    if (!target)
      throw new Error(`Presentation: slide relationship not found: ${slide.relationshipId}`);
    const bounded = Math.max(0, Math.min(toIndex, nodes.length - 1));
    sldIdLst.removeChild(target);
    const refreshed = allDescendantsNs(sldIdLst, P_NS, 'sldId');
    sldIdLst.insertBefore(target, refreshed[bounded] ?? null);
    markDirty(requirePart(this.presentation, 'ppt/presentation.xml'));
    this.presentation._reloadFromPackage();
  }

  deleteItem(indexOrId: number | string): void {
    const slide = this.presentation._getSlide(indexOrId);
    if (!slide?.relationshipId || !slide.path)
      throw new Error(`Presentation: slide not found: ${String(indexOrId)}`);
    if (this.presentation._slides().length <= 1)
      throw new Error('Presentation: cannot delete the final slide');
    removeSlideId(this.presentation, slide.relationshipId);
    removePresentationRelationship(this.presentation, slide.relationshipId);
    deletePart(this.presentation._package(), slide.path);
    deletePart(this.presentation._package(), relsPathForPart(slide.path));
    deletePart(this.presentation._package(), '[Content_Types].xml');
    this.presentation._reloadFromPackage();
  }
}

export interface PptSlideAddOptions {
  index?: number | 'end';
  title?: string;
}

export interface PptSlideCopyOptions {
  index?: number | 'end';
}

export interface PptShapeAddOptions {
  name?: string;
  geometry?: Partial<PptShapeGeometry>;
}

export interface PptTableAddOptions extends PptShapeAddOptions {
  values?: string[][];
}

export interface PptImageAddOptions extends PptShapeAddOptions {
  mimeType?: string;
  filename?: string;
}

export interface PptShapeDuplicateOptions extends PptShapeAddOptions {}

export type PptPresetShape = 'rect' | 'roundRect' | 'ellipse' | 'line';

export class SlideCore {
  readonly shapes: ShapeCollectionCore;
  readonly notes: SlideNotesApiCore;
  readonly comments: SlideCommentsApiCore;

  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
  ) {
    this.shapes = new ShapeCollectionCore(presentation, slideId);
    this.notes = new SlideNotesApiCore(presentation, slideId);
    this.comments = new SlideCommentsApiCore(presentation, slideId);
  }

  id(): string {
    return this.model().id;
  }

  index(): number {
    return this.model().index;
  }

  name(): string | undefined {
    return this.model().name;
  }

  title(): string | undefined {
    return this.model().title;
  }

  layout(): PptSlideLayoutModel | null {
    const layout = this.model().layout;
    return layout ? cloneModel(layout) : null;
  }

  transition(): PptSlideTransitionModel | null {
    const transition = this.model().transition;
    return transition ? cloneModel(transition) : null;
  }

  timing(): PptSlideTimingModel | null {
    const timing = this.model().timing;
    return timing ? cloneModel(timing) : null;
  }

  properties(): PptSlideProperties {
    return slideProperties(this.model());
  }

  setTitle(text: string, opts?: PptTextWriteOptions): void {
    const firstTextShape = this.model().shapes.find((shape) => shape.text);
    if (!firstTextShape)
      throw new Error(`Presentation: slide ${this.slideId} has no text shape for title`);
    this.shapes.getItem(firstTextShape.id).text.setText(text, opts);
    this.model().title = text;
  }

  findText(query: string | RegExp, options: PptFindTextOptions = {}): PptTextMatch[] {
    const slide = this.presentation._getSlideInternal(this.slideId);
    if (!slide) throw new Error(`Presentation: slide not found: ${this.slideId}`);
    return findTextInSlide(slide, query, options);
  }

  replaceText(
    search: string | RegExp,
    replacement: string,
    options: PptTextReplaceOptions = {},
  ): PptTextReplaceResult {
    let count = 0;
    const slide = this.model();
    for (const shape of flattenShapeModels(slide.shapes)) {
      if (shape.text && !shape.table) {
        count += new ShapeCore(this.presentation, this.slideId, shape.id).text.replace(
          search,
          replacement,
          options,
        ).count;
      }
      if (shape.table) {
        for (let row = 0; row < shape.table.values.length; row++) {
          for (let column = 0; column < shape.table.values[row]!.length; column++) {
            const current = shape.table.values[row]![column]!;
            const result = replaceTextValue(current, search, replacement, options);
            if (result.count > 0) {
              new ShapeCore(this.presentation, this.slideId, shape.id).table.setCellText(
                row,
                column,
                result.text,
              );
              count += result.count;
            }
          }
        }
      }
    }
    return { count };
  }

  duplicate(options?: PptSlideCopyOptions): SlideCore {
    return this.presentation.slides.copy(this.slideId, options);
  }

  moveTo(index: number): void {
    this.presentation.slides.move(this.slideId, index);
  }

  delete(): void {
    this.presentation.slides.deleteItem(this.slideId);
  }

  toModel(): PptSlideModel {
    return cloneModel(this.model());
  }

  private model(): PptSlideModel {
    const slide = this.presentation._getSlide(this.slideId);
    if (!slide) throw new Error(`Presentation: slide not found: ${this.slideId}`);
    return slide;
  }
}

export class SlideNotesApiCore {
  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
  ) {}

  hasText(): boolean {
    return Boolean(this.model().notes?.text);
  }

  getText(): string {
    return this.model().notes?.text ?? '';
  }

  setText(text: string): void {
    const slide = this.internal();
    const part = ensureNotesSlidePart(this.presentation, slide);
    setNotesText(part.xmlDom!, text);
    markDirty(part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'slide.notes.set',
      slideId: this.slideId,
      text,
    });
  }

  clear(): void {
    const slide = this.internal();
    const part = getNotesSlidePart(this.presentation, slide);
    if (part?.xmlDom) {
      setNotesText(part.xmlDom, '');
      markDirty(part);
      this.presentation._reloadFromPackage();
    } else {
      slide.model.notes = { text: '' };
    }
    this.presentation._record({
      type: 'slide.notes.clear',
      slideId: this.slideId,
    });
  }

  private model(): PptSlideModel {
    const slide = this.presentation._getSlide(this.slideId);
    if (!slide) throw new Error(`Presentation: slide not found: ${this.slideId}`);
    return slide;
  }

  private internal(): PptReadResult['slides'][number] {
    const slide = this.presentation._getSlideInternal(this.slideId);
    if (!slide) throw new Error(`Presentation: slide not found: ${this.slideId}`);
    return slide;
  }
}

export class SlideCommentsApiCore {
  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
  ) {}

  items(): PptCommentModel[] {
    return cloneModel(this.model().comments ?? []);
  }

  getCount(): number {
    return this.model().comments?.length ?? 0;
  }

  add(options: PptCommentAddOptions): PptCommentModel {
    if (!options.text) throw new Error('Presentation: comment text is required');
    const slide = this.internal();
    const author = ensureCommentAuthor(this.presentation, {
      name: options.author ?? 'Coco Reviewer',
      initials: options.initials ?? 'CR',
    });
    const part = ensureSlideCommentsPart(this.presentation, slide);
    const comment = appendSlideComment(part.xmlDom!, {
      ...options,
      authorId: author.id,
    });
    markDirty(part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'slide.comment.add',
      slideId: this.slideId,
      comment: options,
    });
    return comment;
  }

  delete(commentId: string): void {
    const slide = this.internal();
    const part = getSlideCommentsPart(this.presentation, slide);
    if (!part?.xmlDom) return;
    const removed = removeSlideComment(part.xmlDom, commentId);
    if (!removed) throw new Error(`Presentation: comment not found: ${commentId}`);
    markDirty(part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'slide.comment.delete',
      slideId: this.slideId,
      commentId,
    });
  }

  clear(): void {
    const slide = this.internal();
    const part = getSlideCommentsPart(this.presentation, slide);
    if (part?.xmlDom) {
      for (const comment of allDescendantsNs(part.xmlDom, P_NS, 'cm')) {
        comment.parentNode?.removeChild(comment);
      }
      markDirty(part);
      this.presentation._reloadFromPackage();
    }
    this.presentation._record({
      type: 'slide.comments.clear',
      slideId: this.slideId,
    });
  }

  private model(): PptSlideModel {
    const slide = this.presentation._getSlide(this.slideId);
    if (!slide) throw new Error(`Presentation: slide not found: ${this.slideId}`);
    return slide;
  }

  private internal(): PptReadResult['slides'][number] {
    const slide = this.presentation._getSlideInternal(this.slideId);
    if (!slide) throw new Error(`Presentation: slide not found: ${this.slideId}`);
    return slide;
  }
}

export class ShapeCollectionCore {
  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
    private readonly parentShapeId?: string,
  ) {}

  items(): ShapeCore[] {
    return this.shapeModels().map(
      (shape) => new ShapeCore(this.presentation, this.slideId, shape.id),
    );
  }

  getCount(): number {
    return this.shapeModels().length;
  }

  getItem(idOrName: string): ShapeCore {
    const matches = this.shapeModels().filter(
      (shape) => shape.id === idOrName || shape.name === idOrName,
    );
    if (matches.length === 0) throw new Error(`Presentation: shape not found: ${idOrName}`);
    if (matches.length > 1) throw new Error(`Presentation: shape lookup is ambiguous: ${idOrName}`);
    return new ShapeCore(this.presentation, this.slideId, matches[0]!.id);
  }

  getItemAt(index: number): ShapeCore {
    const shape = this.shapeModels()[index];
    if (!shape) throw new Error(`Presentation: shape index out of range: ${index}`);
    return new ShapeCore(this.presentation, this.slideId, shape.id);
  }

  getItemOrNullObject(idOrName: string): ShapeCore | null {
    const matches = this.shapeModels().filter(
      (shape) => shape.id === idOrName || shape.name === idOrName,
    );
    return matches.length === 1
      ? new ShapeCore(this.presentation, this.slideId, matches[0]!.id)
      : null;
  }

  addTextBox(text: string, options: PptShapeAddOptions = {}): ShapeCore {
    const slide = this.slideModel();
    const shape = makeTextShapeModel(slide, text, options);
    appendShapeXml(this.presentation, this.slideId, simpleTextShapeXml(shape, slide.shapes.length));
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'shape.add',
      slideId: this.slideId,
      shapeId: shape.id,
      kind: 'textBox',
    });
    return new ShapeCore(this.presentation, this.slideId, shape.id);
  }

  addTable(rowCount: number, columnCount: number, options: PptTableAddOptions = {}): ShapeCore {
    if (!Number.isInteger(rowCount) || rowCount <= 0)
      throw new Error('Presentation: table rowCount must be positive');
    if (!Number.isInteger(columnCount) || columnCount <= 0)
      throw new Error('Presentation: table columnCount must be positive');
    const slide = this.slideModel();
    const shape = makeTableShapeModel(slide, rowCount, columnCount, options);
    appendShapeXml(this.presentation, this.slideId, tableShapeXml(shape, slide.shapes.length));
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'shape.add',
      slideId: this.slideId,
      shapeId: shape.id,
      kind: 'table',
    });
    return new ShapeCore(this.presentation, this.slideId, shape.id);
  }

  addImage(source: PptImageSource, options: PptImageAddOptions = {}): ShapeCore {
    const slide = this.slideModel();
    const image = imageSourceInfo(source, options);
    const mediaPath = nextMediaPath(this.presentation, image.extension);
    const relationshipId = addSlideImageRelationship(this.presentation, this.slideId, mediaPath);
    const shape = makeImageShapeModel(slide, relationshipId, mediaPath, image.contentType, options);
    setPart(
      this.presentation._package(),
      createBinaryPart({
        path: mediaPath,
        contentType: image.contentType,
        bytes: image.bytes,
        originalBytes: image.bytes,
      }),
    );
    appendShapeXml(this.presentation, this.slideId, pictureShapeXml(shape, slide.shapes.length));
    deletePart(this.presentation._package(), '[Content_Types].xml');
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'shape.add',
      slideId: this.slideId,
      shapeId: shape.id,
      kind: 'picture',
      imageRef: mediaPath,
    });
    return new ShapeCore(this.presentation, this.slideId, shape.id);
  }

  addShape(preset: PptPresetShape, options: PptShapeAddOptions = {}): ShapeCore {
    const slide = this.slideModel();
    const shape = makePresetShapeModel(slide, preset, options);
    appendShapeXml(
      this.presentation,
      this.slideId,
      presetShapeXml(shape, preset, slide.shapes.length),
    );
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'shape.add',
      slideId: this.slideId,
      shapeId: shape.id,
      kind: shape.kind === 'line' ? 'line' : 'geometricShape',
    });
    return new ShapeCore(this.presentation, this.slideId, shape.id);
  }

  deleteItem(idOrName: string): void {
    this.getItem(idOrName).delete();
  }

  private slideModel(): PptSlideModel {
    const slide = this.presentation._getSlide(this.slideId);
    if (!slide) throw new Error(`Presentation: slide not found: ${this.slideId}`);
    return slide;
  }

  private shapeModels(): PptShapeModel[] {
    if (!this.parentShapeId) return this.slideModel().shapes;
    const parent = getShapeInternal(this.presentation, this.slideId, this.parentShapeId);
    if (parent.model.kind !== 'group')
      throw new Error(`Presentation: shape ${this.parentShapeId} is not a group`);
    return parent.model.children ?? [];
  }
}

export class ShapeCore {
  readonly chart: ShapeChartApiCore;
  readonly text: ShapeTextApiCore;
  readonly image: ShapeImageApiCore;
  readonly table: TableApiCore;
  readonly group: ShapeGroupApiCore;

  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
    private readonly shapeId: string,
  ) {
    this.chart = new ShapeChartApiCore(presentation, slideId, shapeId);
    this.text = new ShapeTextApiCore(presentation, slideId, shapeId);
    this.image = new ShapeImageApiCore(presentation, slideId, shapeId);
    this.table = new TableApiCore(presentation, slideId, shapeId);
    this.group = new ShapeGroupApiCore(presentation, slideId, shapeId);
  }

  id(): string {
    return this.model().id;
  }

  name(): string | undefined {
    return this.model().name;
  }

  setName(name: string): void {
    const shape = this.internal();
    const cNvPr = allDescendantsNs(shape.element, P_NS, 'cNvPr')[0];
    if (!cNvPr) throw new Error(`Presentation: shape ${this.shapeId} has no cNvPr`);
    cNvPr.setAttribute('name', name);
    shape.model.name = name;
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.name.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      name,
    });
  }

  kind(): PptShapeKind {
    return this.model().kind;
  }

  geometry(): PptShapeGeometry {
    return { ...this.model().geometry };
  }

  effectiveGeometry(): PptShapeGeometry {
    const shape = this.internal();
    return resolvePptEffectivePlaceholderProperties(shape.model, shape.slide.model.layout).geometry;
  }

  placeholder(): PptPlaceholderModel | null {
    const placeholder = this.model().placeholder;
    return placeholder ? { ...placeholder } : null;
  }

  effectivePlaceholder(): PptPlaceholderModel | null {
    const shape = this.internal();
    const placeholder = resolvePptEffectivePlaceholderProperties(
      shape.model,
      shape.slide.model.layout,
    ).placeholder;
    return placeholder ? { ...placeholder } : null;
  }

  style(): PptShapeStyle {
    return { ...this.model().style };
  }

  effectiveStyle(): PptShapeStyle {
    const shape = this.internal();
    return {
      ...resolvePptEffectivePlaceholderProperties(shape.model, shape.slide.model.layout).style,
    };
  }

  hyperlink(): PptHyperlinkModel | null {
    const hyperlink = this.model().hyperlink;
    return hyperlink ? cloneModel(hyperlink) : null;
  }

  setHyperlink(hyperlink: PptHyperlinkWriteOptions): void {
    const shape = this.internal();
    const cNvPr = allDescendantsNs(shape.element, P_NS, 'cNvPr')[0];
    if (!cNvPr) throw new Error(`Presentation: shape ${this.shapeId} has no cNvPr`);
    const hlink = findOrCreateHyperlink(cNvPr);
    const model = setHyperlinkOnElement(this.presentation, shape.slide, hlink, hyperlink);
    shape.model.hyperlink = model;
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.hyperlink.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      hyperlink,
    });
  }

  clearHyperlink(): void {
    const shape = this.internal();
    const cNvPr = allDescendantsNs(shape.element, P_NS, 'cNvPr')[0];
    if (!cNvPr) return;
    const hlink = firstChildNs(cNvPr, A_NS, 'hlinkClick');
    clearHyperlinkElement(this.presentation, shape.slide, hlink);
    delete shape.model.hyperlink;
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.hyperlink.clear',
      slideId: this.slideId,
      shapeId: this.shapeId,
    });
  }

  setGeometry(geometry: Partial<PptShapeGeometry>): void {
    const shape = this.internal();
    const model = shape.model;
    model.geometry = { ...model.geometry, ...geometry };
    const xfrm = findOrCreateTransform(shape.element);
    let off = firstChildNs(xfrm, A_NS, 'off');
    if (!off) {
      off = xfrm.ownerDocument.createElementNS(A_NS, 'a:off');
      xfrm.appendChild(off);
    }
    let ext = firstChildNs(xfrm, A_NS, 'ext');
    if (!ext) {
      ext = xfrm.ownerDocument.createElementNS(A_NS, 'a:ext');
      xfrm.appendChild(ext);
    }
    setAttrIfDefined(off, 'x', model.geometry.left);
    setAttrIfDefined(off, 'y', model.geometry.top);
    setAttrIfDefined(ext, 'cx', model.geometry.width);
    setAttrIfDefined(ext, 'cy', model.geometry.height);
    setAttrIfDefined(xfrm, 'rot', degreesToOoxmlAngle(model.geometry.rotation));
    markDirty(shape.slide.part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'shape.geometry.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      geometry,
    });
  }

  setFillColor(color: string): void {
    this.setStyle({ fillColor: normalizeHexColor(color) });
  }

  setLineStyle(style: Pick<PptShapeStyle, 'lineColor' | 'lineWidth'>): void {
    this.setStyle({
      ...(style.lineColor ? { lineColor: normalizeHexColor(style.lineColor) } : {}),
      ...(style.lineWidth !== undefined ? { lineWidth: style.lineWidth } : {}),
    });
  }

  setStyle(style: Partial<PptShapeStyle>): void {
    const shape = this.internal();
    const normalized: Partial<PptShapeStyle> = {
      ...(style.fillColor ? { fillColor: normalizeHexColor(style.fillColor) } : {}),
      ...(style.opacity !== undefined ? { opacity: normalizeOpacity(style.opacity) } : {}),
      ...(style.lineColor ? { lineColor: normalizeHexColor(style.lineColor) } : {}),
      ...(style.lineOpacity !== undefined
        ? { lineOpacity: normalizeOpacity(style.lineOpacity) }
        : {}),
      ...(style.lineWidth !== undefined ? { lineWidth: style.lineWidth } : {}),
      ...(style.shadow ? { shadow: normalizeShadow(style.shadow) } : {}),
    };
    const spPr = findOrCreateShapeProperties(shape.element);
    if (normalized.fillColor) setSolidFill(spPr, normalized.fillColor, normalized.opacity);
    else if (normalized.opacity !== undefined)
      setExistingSolidFillOpacity(spPr, normalized.opacity, shape.model.style?.fillColor);
    if (
      normalized.lineColor ||
      normalized.lineOpacity !== undefined ||
      normalized.lineWidth !== undefined
    ) {
      const ln = findOrCreateLine(spPr);
      if (normalized.lineColor) setSolidFill(ln, normalized.lineColor, normalized.lineOpacity);
      else if (normalized.lineOpacity !== undefined)
        setExistingSolidFillOpacity(ln, normalized.lineOpacity, shape.model.style?.lineColor);
      if (normalized.lineWidth !== undefined) ln.setAttribute('w', String(normalized.lineWidth));
    }
    if (normalized.shadow) setOuterShadow(spPr, normalized.shadow);
    shape.model.style = { ...shape.model.style, ...normalized };
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.style.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      style: normalized,
    });
  }

  delete(): void {
    const shape = this.internal();
    const parent = shape.element.parentNode;
    if (!parent) throw new Error(`Presentation: shape ${this.shapeId} has no parent XML node`);
    parent.removeChild(shape.element);
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.delete',
      slideId: this.slideId,
      shapeId: this.shapeId,
    });
    this.presentation._reloadFromPackage();
  }

  duplicate(options: PptShapeDuplicateOptions = {}): ShapeCore {
    const shape = this.internal();
    const parent = shape.element.parentNode;
    if (!parent) throw new Error(`Presentation: shape ${this.shapeId} has no parent XML node`);
    const slide = shape.slide.model;
    const newId = nextShapeId(slide);
    const clone = shape.element.cloneNode(true) as Element;
    const name = options.name ?? `${shape.model.name ?? shape.model.kind} Copy ${newId}`;
    const cNvPr = allDescendantsNs(clone, P_NS, 'cNvPr')[0];
    if (!cNvPr) throw new Error(`Presentation: shape ${this.shapeId} has no cNvPr`);
    cNvPr.setAttribute('id', newId);
    cNvPr.setAttribute('name', name);
    if (options.geometry)
      applyGeometryToElement(clone, { ...shape.model.geometry, ...options.geometry });
    let imageRef: string | undefined;
    if (shape.model.kind === 'picture' && shape.model.image?.targetPath) {
      imageRef = duplicatePictureMedia(
        this.presentation,
        this.slideId,
        clone,
        shape.model.image.targetPath,
      );
    }
    parent.insertBefore(clone, shape.element.nextSibling);
    markDirty(shape.slide.part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'shape.add',
      slideId: this.slideId,
      shapeId: newId,
      kind: shape.model.kind,
      ...(imageRef ? { imageRef } : {}),
    });
    return new ShapeCore(this.presentation, this.slideId, newId);
  }

  bringToFront(): void {
    this.moveInZOrder('front');
  }

  sendToBack(): void {
    this.moveInZOrder('back');
  }

  toModel(): PptShapeModel {
    return cloneModel(this.model());
  }

  private model(): PptShapeModel {
    return this.internal().model;
  }

  private internal(): PptShapeInternal {
    return getShapeInternal(this.presentation, this.slideId, this.shapeId);
  }

  private moveInZOrder(position: 'front' | 'back'): void {
    const shape = this.internal();
    const parent = shape.element.parentNode;
    if (!parent) throw new Error(`Presentation: shape ${this.shapeId} has no parent XML node`);
    parent.removeChild(shape.element);
    if (position === 'front') {
      parent.appendChild(shape.element);
    } else {
      parent.insertBefore(shape.element, firstShapeTreeElement(parent) ?? null);
    }
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.zOrder.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      position,
    });
    this.presentation._reloadFromPackage();
  }
}

export class ShapeTextApiCore {
  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
    private readonly shapeId: string,
  ) {}

  hasText(): boolean {
    return this.internal().textElements.length > 0;
  }

  getText(): string {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    return shape.model.text?.text ?? '';
  }

  style(): PptTextStyle {
    return { ...this.internal().model.text?.style };
  }

  effectiveStyle(): PptTextStyle {
    return effectiveShapeTextStyle(this.internal());
  }

  body(): PptTextBodyProperties {
    return { ...this.internal().model.text?.body };
  }

  effectiveBody(): PptTextBodyProperties {
    const shape = this.internal();
    return {
      ...resolvePptEffectivePlaceholderProperties(shape.model, shape.slide.model.layout).textBody,
    };
  }

  paragraphs(): PptTextParagraph[] {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    return textParagraphsFromElement(shape.element, shape.slide);
  }

  runs(): PptTextRun[] {
    return this.paragraphs().flatMap((paragraph) => paragraph.runs);
  }

  effectiveRuns(): PptTextRun[] {
    const inherited = this.effectiveStyle();
    return this.paragraphs().flatMap((paragraph) =>
      paragraph.runs.map((run) => ({
        ...run,
        style: {
          ...inherited,
          ...(paragraph.style?.alignment ? { alignment: paragraph.style.alignment } : {}),
          ...(paragraph.style?.lineHeight !== undefined
            ? { lineHeight: paragraph.style.lineHeight }
            : {}),
          ...run.style,
        },
        ...(run.hyperlink ? { hyperlink: { ...run.hyperlink } } : {}),
      })),
    );
  }

  setParagraphs(paragraphs: PptTextParagraphWrite[]): void {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const textBody = findShapeTextBody(shape.element);
    if (!textBody) throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    replaceTextBodyParagraphs(textBody, paragraphs);
    markDirty(shape.slide.part);
    this.presentation._reloadFromPackage();
    const updated = this.internal();
    const nextParagraphs = textParagraphsFromElement(updated.element, updated.slide);
    const text = nextParagraphs.map((paragraph) => paragraph.text).join('\n');
    updated.model.text = { text, paragraphs: nextParagraphs };
    if (updated.slide.model.title === shape.model.text?.text) updated.slide.model.title = text;
    this.presentation._record({
      type: 'shape.text.paragraphs.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      paragraphs,
    });
  }

  setRunStyle(runIndex: number, style: Partial<PptTextStyle>): void {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const textBody = findShapeTextBody(shape.element);
    if (!textBody) throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const runs = flatTextRuns(textBody);
    const run = runs[requireIndex(runIndex, runs.length, 'run')];
    if (!run) throw new Error(`Presentation: text run out of range: ${runIndex}`);
    const normalized = normalizeTextStyle(style);
    applyTextStyleToRun(run, normalized);
    shape.model.text = {
      text: shape.model.text?.text ?? this.getText(),
      paragraphs: textParagraphsFromElement(shape.element, shape.slide),
      style: { ...shape.model.text?.style, ...normalized },
    };
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.text.run.style.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      runIndex,
      style: normalized,
    });
  }

  setRunTextStyle(runIndex: number, style: Partial<PptTextStyle>): void {
    this.setRunStyle(runIndex, style);
  }

  setRunHyperlink(runIndex: number, hyperlink: PptHyperlinkWriteOptions): void {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const textBody = findShapeTextBody(shape.element);
    if (!textBody) throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const runs = flatTextRuns(textBody);
    const run = runs[requireIndex(runIndex, runs.length, 'run')];
    if (!run) throw new Error(`Presentation: text run out of range: ${runIndex}`);
    const rPr = findOrCreateRunProperties(run);
    const hlink = findOrCreateHyperlink(rPr);
    setHyperlinkOnElement(this.presentation, shape.slide, hlink, hyperlink);
    shape.model.text = {
      text: shape.model.text?.text ?? this.getText(),
      paragraphs: textParagraphsFromElement(shape.element, shape.slide),
      style: shape.model.text?.style,
    };
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.text.run.hyperlink.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      runIndex,
      hyperlink,
    });
  }

  clearRunHyperlink(runIndex: number): void {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const textBody = findShapeTextBody(shape.element);
    if (!textBody) throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const runs = flatTextRuns(textBody);
    const run = runs[requireIndex(runIndex, runs.length, 'run')];
    if (!run) throw new Error(`Presentation: text run out of range: ${runIndex}`);
    const rPr = firstChildNs(run, A_NS, 'rPr');
    clearHyperlinkElement(
      this.presentation,
      shape.slide,
      rPr ? firstChildNs(rPr, A_NS, 'hlinkClick') : null,
    );
    shape.model.text = {
      text: shape.model.text?.text ?? this.getText(),
      paragraphs: textParagraphsFromElement(shape.element, shape.slide),
      style: shape.model.text?.style,
    };
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.text.run.hyperlink.clear',
      slideId: this.slideId,
      shapeId: this.shapeId,
      runIndex,
    });
  }

  setText(text: string, opts: PptTextWriteOptions = {}): void {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const oldText = shape.model.text?.text ?? '';
    if (opts.preserveFirstRunStyle) {
      const [firstParagraph] = textParagraphsFromElement(shape.element, shape.slide);
      const firstRun =
        firstParagraph?.runs.find((run) => run.text.length > 0) ?? firstParagraph?.runs[0];
      this.setParagraphs([
        {
          text,
          ...(firstParagraph?.style ? { style: firstParagraph.style } : {}),
          runs: [
            {
              text,
              ...(firstRun?.style ? { style: firstRun.style } : {}),
            },
          ],
        },
      ]);
      const updated = this.internal();
      updated.model.text = {
        ...(updated.model.text ?? { text }),
        ...(firstRun?.style ? { style: firstRun.style } : {}),
      };
      if (updated.slide.model.title === oldText) updated.slide.model.title = text;
      return;
    }
    const previousStyle = shape.model.text?.style;
    shape.textElements[0]!.textContent = text;
    for (const extra of shape.textElements.slice(1)) extra.textContent = '';
    shape.model.text = {
      text,
      paragraphs: textParagraphsFromElement(shape.element, shape.slide),
      ...(previousStyle ? { style: previousStyle } : {}),
    };
    if (shape.slide.model.title === oldText) shape.slide.model.title = text;
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.text.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      text,
    });
  }

  setStyle(style: Partial<PptTextStyle>): void {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const normalized = normalizeTextStyle(style);
    for (const textEl of shape.textElements) {
      const run = textEl.parentNode;
      if (!run || run.nodeType !== 1) continue;
      applyTextStyleToRun(run as Element, normalized);
    }
    shape.model.text = {
      text: shape.model.text?.text ?? this.getText(),
      style: { ...shape.model.text?.style, ...normalized },
    };
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.text.style.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      style: normalized,
    });
  }

  setTextStyle(style: Partial<PptTextStyle>): void {
    this.setStyle(style);
  }

  setBodyProperties(body: Partial<PptTextBodyProperties>): void {
    const shape = this.internal();
    if (shape.textElements.length === 0)
      throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const textBody = findShapeTextBody(shape.element);
    if (!textBody) throw new Error(`Presentation: shape ${this.shapeId} has no text body`);
    const bodyPr = findOrCreateTextBodyProperties(textBody);
    const normalized = normalizeShapeTextBodyProperties(body);
    applyShapeTextBodyProperties(bodyPr, normalized);
    shape.model.text = {
      ...(shape.model.text ?? { text: this.getText() }),
      body: { ...shape.model.text?.body, ...normalized },
      paragraphs: textParagraphsFromElement(shape.element, shape.slide),
    };
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.text.body.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      body: normalized,
    });
  }

  setBold(enabled: boolean): void {
    this.setStyle({ bold: enabled });
  }

  setItalic(enabled: boolean): void {
    this.setStyle({ italic: enabled });
  }

  setAlignment(alignment: PptTextStyle['alignment']): void {
    if (!alignment) throw new Error('Presentation: text alignment is required');
    this.setStyle({ alignment });
  }

  replace(
    search: string | RegExp,
    replacement: string,
    opts: PptTextReplaceOptions = {},
  ): PptTextReplaceResult {
    const current = this.getText();
    const pattern =
      typeof search === 'string'
        ? new RegExp(escapeRegExp(search), opts.caseSensitive ? 'g' : 'gi')
        : search;
    let count = 0;
    const next = current.replace(pattern, () => {
      count++;
      return replacement;
    });
    if (count > 0) {
      this.setText(next, opts);
      this.presentation._record({
        type: 'shape.text.replace',
        slideId: this.slideId,
        shapeId: this.shapeId,
        search: String(search),
        replacement,
        count,
      });
    }
    return { count };
  }

  private internal(): PptShapeInternal {
    return getShapeInternal(this.presentation, this.slideId, this.shapeId);
  }
}

function effectiveShapeTextStyle(shape: PptShapeInternal): PptTextStyle {
  return {
    ...resolvePptEffectivePlaceholderProperties(shape.model, shape.slide.model.layout).textStyle,
  };
}

export class ShapeChartApiCore {
  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
    private readonly shapeId: string,
  ) {}

  isChart(): boolean {
    const shape = this.internal().model;
    return shape.kind === 'chart' && Boolean(shape.chart);
  }

  getInfo(): PptChartModel | null {
    const chart = this.internal().model.chart;
    return chart ? cloneModel(chart) : null;
  }

  getData(): PptChartData | null {
    const part = this.chartPart();
    if (!part?.xmlDom) return null;
    return chartDataFromDocument(part.xmlDom);
  }

  setData(data: PptChartData): void {
    const shape = this.internal();
    const part = this.chartPart();
    if (!part?.xmlDom) throw new Error(`Presentation: chart ${this.shapeId} has no chart XML`);
    setChartData(part.xmlDom, data);
    markDirty(part);
    shape.model.chart = {
      ...shape.model.chart,
      ...(data.title ? { title: data.title } : {}),
      seriesNames: data.series
        .map((series) => series.name)
        .filter((value): value is string => Boolean(value)),
    };
    this.presentation._record({
      type: 'shape.chart.data.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      data,
    });
  }

  private internal(): PptShapeInternal {
    return getShapeInternal(this.presentation, this.slideId, this.shapeId);
  }

  private chartPart(): Part | undefined {
    const chart = this.internal().model.chart;
    return chart?.targetPath ? getPart(this.presentation._package(), chart.targetPath) : undefined;
  }
}

export class ShapeGroupApiCore {
  private readonly children: ShapeCollectionCore;

  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
    private readonly shapeId: string,
  ) {
    this.children = new ShapeCollectionCore(presentation, slideId, shapeId);
  }

  isGroup(): boolean {
    return this.internal().model.kind === 'group';
  }

  items(): ShapeCore[] {
    this.requireGroup();
    return this.children.items();
  }

  getCount(): number {
    this.requireGroup();
    return this.children.getCount();
  }

  getItem(idOrName: string): ShapeCore {
    this.requireGroup();
    return this.children.getItem(idOrName);
  }

  getItemAt(index: number): ShapeCore {
    this.requireGroup();
    return this.children.getItemAt(index);
  }

  getItemOrNullObject(idOrName: string): ShapeCore | null {
    this.requireGroup();
    return this.children.getItemOrNullObject(idOrName);
  }

  private requireGroup(): void {
    if (!this.isGroup()) throw new Error(`Presentation: shape ${this.shapeId} is not a group`);
  }

  private internal(): PptShapeInternal {
    return getShapeInternal(this.presentation, this.slideId, this.shapeId);
  }
}

function chartDataFromDocument(doc: Document): PptChartData {
  const title = chartTitleFromDocument(doc);
  const series = allDescendantsLocal(doc, 'ser').map((ser) => {
    const chartLabels = chartDataLabelSettings(closestChartElement(ser));
    const seriesLabels = mergeChartDataLabelSettings(chartLabels, chartDataLabelSettings(ser));
    const dataLabelLayouts = [
      ...chartDataLabelLayouts(closestChartElement(ser)),
      ...chartDataLabelLayouts(ser),
    ];
    return {
      name: chartSeriesName(ser),
      categories: chartSeriesCategories(ser),
      values: chartSeriesValues(ser),
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
    };
  });
  return {
    ...(title ? { title } : {}),
    series,
  };
}

function setChartData(doc: Document, data: PptChartData): void {
  const seriesNodes = allDescendantsLocal(doc, 'ser');
  if (seriesNodes.length !== data.series.length) {
    throw new Error(
      `Presentation: chart data must contain ${seriesNodes.length} series, got ${data.series.length}`,
    );
  }
  if (data.title !== undefined) setChartTitle(doc, data.title);
  for (let index = 0; index < seriesNodes.length; index++) {
    const ser = seriesNodes[index]!;
    const series = data.series[index]!;
    if (series.name !== undefined) setChartSeriesName(ser, series.name);
    if (series.categories) setChartSeriesCategories(ser, series.categories);
    setChartSeriesValues(ser, series.values);
  }
}

function chartTitleFromDocument(doc: Document): string | undefined {
  const title = allDescendantsLocal(doc, 'title')[0];
  if (!title) return undefined;
  const text =
    allDescendantsNs(title, A_NS, 't').map(textOf).join('') ||
    allDescendantsLocal(title, 'v').map(textOf).join('');
  return text || undefined;
}

function setChartTitle(doc: Document, title: string): void {
  const titleNode = allDescendantsLocal(doc, 'title')[0];
  if (!titleNode) throw new Error('Presentation: chart has no editable title');
  const textNodes = allDescendantsNs(titleNode, A_NS, 't');
  if (textNodes.length > 0) {
    textNodes[0]!.textContent = title;
    for (const extra of textNodes.slice(1)) extra.textContent = '';
    return;
  }
  const valueNodes = allDescendantsLocal(titleNode, 'v');
  if (valueNodes.length > 0) {
    valueNodes[0]!.textContent = title;
    for (const extra of valueNodes.slice(1)) extra.textContent = '';
    return;
  }
  throw new Error('Presentation: chart title has no editable text node');
}

function chartSeriesName(ser: Element): string | undefined {
  const tx = firstChildLocal(ser, 'tx');
  if (!tx) return undefined;
  const text =
    allDescendantsLocal(tx, 'v').map(textOf).join('') ||
    allDescendantsNs(tx, A_NS, 't').map(textOf).join('');
  return text || undefined;
}

function setChartSeriesName(ser: Element, name: string): void {
  const tx = findOrCreateChartChild(ser, 'tx');
  const strRef = firstChildLocal(tx, 'strRef');
  if (strRef) {
    const cache = findOrCreateChartChild(strRef, 'strCache');
    setStringCache(cache, [name]);
    return;
  }
  let value = firstChildLocal(tx, 'v');
  if (!value) {
    value = tx.ownerDocument.createElementNS(C_NS, 'c:v');
    tx.appendChild(value);
  }
  value.textContent = name;
}

function chartSeriesCategories(ser: Element): string[] | undefined {
  const cat = firstChildLocal(ser, 'cat');
  if (!cat) return undefined;
  const values = valuesFromStringCache(cat);
  return values.length > 0 ? values : undefined;
}

function setChartSeriesCategories(ser: Element, categories: string[]): void {
  const cat = findOrCreateChartChild(ser, 'cat');
  const strRef = firstChildLocal(cat, 'strRef') ?? appendChartChild(cat, 'strRef');
  const cache = findOrCreateChartChild(strRef, 'strCache');
  setStringCache(cache, categories);
}

function chartSeriesValues(ser: Element): number[] {
  const val = firstChildLocal(ser, 'val');
  if (!val) return [];
  return valuesFromNumberCache(val);
}

function chartDataLabelSettings(element: Element | null | undefined): {
  showValueLabels?: boolean;
  showCategoryNameLabels?: boolean;
  showPercentLabels?: boolean;
} {
  if (!element) return {};
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

function chartDataLabelLayouts(
  element: Element | null | undefined,
): NonNullable<PptChartSeriesData['dataLabelLayouts']> {
  if (!element) return [];
  const dLbls = firstChildLocal(element, 'dLbls');
  if (!dLbls) return [];
  const result: NonNullable<PptChartSeriesData['dataLabelLayouts']> = [];
  for (const label of allDescendantsLocal(dLbls, 'dLbl')) {
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

function chartDataLabelBoolean<
  K extends 'showValueLabels' | 'showCategoryNameLabels' | 'showPercentLabels',
>(dLbls: Element, xmlName: string, key: K): Partial<Record<K, boolean>> {
  const value = parseBooleanAttr(firstChildLocal(dLbls, xmlName)?.getAttribute('val') ?? null);
  return value === undefined ? {} : ({ [key]: value } as Partial<Record<K, boolean>>);
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

function closestChartElement(element: Element): Element | null {
  let current = element.parentNode;
  while (current?.nodeType === 1) {
    const candidate = current as Element;
    if (candidate.namespaceURI === C_NS && candidate.localName.endsWith('Chart')) return candidate;
    current = candidate.parentNode;
  }
  return null;
}

function setChartSeriesValues(ser: Element, values: number[]): void {
  const val = findOrCreateChartChild(ser, 'val');
  const numRef = firstChildLocal(val, 'numRef') ?? appendChartChild(val, 'numRef');
  const cache = findOrCreateChartChild(numRef, 'numCache');
  setNumberCache(cache, values);
}

function valuesFromStringCache(parent: Element): string[] {
  const cache = firstChildLocal(parent, 'strCache') ?? allDescendantsLocal(parent, 'strCache')[0];
  if (!cache) return [];
  return sortedCachePoints(cache).map(textOf);
}

function valuesFromNumberCache(parent: Element): number[] {
  const cache = firstChildLocal(parent, 'numCache') ?? allDescendantsLocal(parent, 'numCache')[0];
  if (!cache) return [];
  return sortedCachePoints(cache)
    .map((pt) => Number(textOf(firstChildLocal(pt, 'v')) || textOf(pt)))
    .filter((value) => Number.isFinite(value));
}

function sortedCachePoints(cache: Element): Element[] {
  return childElementsNs(cache, C_NS, 'pt').sort(
    (a, b) => Number(a.getAttribute('idx') ?? 0) - Number(b.getAttribute('idx') ?? 0),
  );
}

function setStringCache(cache: Element, values: string[]): void {
  replaceCachePoints(cache, values, (value) => value);
}

function setNumberCache(cache: Element, values: number[]): void {
  replaceCachePoints(cache, values, (value) => String(value));
}

function replaceCachePoints<T>(cache: Element, values: T[], serialize: (value: T) => string): void {
  for (const child of [
    ...childElementsNs(cache, C_NS, 'pt'),
    ...childElementsNs(cache, C_NS, 'ptCount'),
  ]) {
    child.parentNode?.removeChild(child);
  }
  const ptCount = cache.ownerDocument.createElementNS(C_NS, 'c:ptCount');
  ptCount.setAttribute('val', String(values.length));
  cache.appendChild(ptCount);
  values.forEach((value, index) => {
    const pt = cache.ownerDocument.createElementNS(C_NS, 'c:pt');
    pt.setAttribute('idx', String(index));
    const v = cache.ownerDocument.createElementNS(C_NS, 'c:v');
    v.textContent = serialize(value);
    pt.appendChild(v);
    cache.appendChild(pt);
  });
}

function findOrCreateChartChild(parent: Element, localName: string): Element {
  return firstChildLocal(parent, localName) ?? appendChartChild(parent, localName);
}

function appendChartChild(parent: Element, localName: string): Element {
  const child = parent.ownerDocument.createElementNS(C_NS, `c:${localName}`);
  parent.appendChild(child);
  return child;
}

export class ShapeImageApiCore {
  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
    private readonly shapeId: string,
  ) {}

  isImage(): boolean {
    const image = this.internal().model.image;
    return this.internal().model.kind === 'picture' && Boolean(image?.targetPath);
  }

  getInfo(): PptImageInfo | null {
    const image = this.internal().model.image;
    if (!image) return null;
    const part = image.targetPath
      ? getPart(this.presentation._package(), image.targetPath)
      : undefined;
    return {
      relationshipId: image.relationshipId,
      targetPath: image.targetPath,
      contentType: part?.contentType ?? image.contentType,
      filename: image.filename,
      byteLength: part?.bytes?.byteLength ?? part?._originalBytes?.byteLength,
    };
  }

  style(): PptImageStyle {
    const image = this.internal().model.image;
    if (!image) return {};
    return cloneModel({
      ...(image.crop ? { crop: image.crop } : {}),
      ...(image.fillRect ? { fillRect: image.fillRect } : {}),
      ...(image.opacity !== undefined ? { opacity: image.opacity } : {}),
    });
  }

  setStyle(style: Partial<PptImageStyle>): void {
    const shape = this.internal();
    const image = shape.model.image;
    if (!image) throw new Error(`Presentation: shape ${this.shapeId} has no image style surface`);
    const normalized = normalizeImageStyle(style);
    const blipFill = findImageBlipFill(shape.element);
    if (!blipFill) throw new Error(`Presentation: shape ${this.shapeId} has no image fill XML`);
    if (normalized.crop) setImageCrop(blipFill, normalized.crop);
    if (normalized.fillRect) setImageFillRect(blipFill, normalized.fillRect);
    if (normalized.opacity !== undefined) {
      const blip = firstChildNs(blipFill, A_NS, 'blip');
      if (!blip) throw new Error(`Presentation: shape ${this.shapeId} has no image blip XML`);
      setImageOpacity(blip, normalized.opacity);
    }
    shape.model.image = {
      ...image,
      ...normalized,
      ...(normalized.crop ? { crop: { ...image.crop, ...normalized.crop } } : {}),
      ...(normalized.fillRect ? { fillRect: { ...image.fillRect, ...normalized.fillRect } } : {}),
    };
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'shape.image.style.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      style: normalized,
    });
  }

  setCrop(crop: PptImageStyle['crop']): void {
    this.setStyle({ crop });
  }

  setFillRect(fillRect: PptImageStyle['fillRect']): void {
    this.setStyle({ fillRect });
  }

  setOpacity(opacity: number): void {
    this.setStyle({ opacity });
  }

  bytes(): Uint8Array {
    const image = this.internal().model.image;
    if (!image?.targetPath)
      throw new Error(`Presentation: shape ${this.shapeId} has no image bytes`);
    return exportedImageFromPart(this.presentation, image.targetPath).bytes;
  }

  replace(source: PptImageSource): void {
    const shape = this.internal();
    const image = shape.model.image;
    if (!image?.targetPath)
      throw new Error(`Presentation: shape ${this.shapeId} has no replaceable image`);
    const current = getPart(this.presentation._package(), image.targetPath);
    if (!current) throw new Error(`Presentation: image part not found: ${image.targetPath}`);
    const bytes = imageSourceBytes(source);
    setPart(
      this.presentation._package(),
      createBinaryPart({
        path: image.targetPath,
        contentType: current.contentType,
        bytes,
        originalBytes: bytes,
      }),
    );
    this.presentation._record({
      type: 'shape.image.replace',
      slideId: this.slideId,
      shapeId: this.shapeId,
      imageRef: image.targetPath,
    });
  }

  private internal(): PptShapeInternal {
    return getShapeInternal(this.presentation, this.slideId, this.shapeId);
  }
}

export class TableApiCore {
  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
    private readonly shapeId: string,
  ) {}

  isTable(): boolean {
    return Boolean(this.internal().model.table);
  }

  rowCount(): number {
    return this.table().rowCount;
  }

  columnCount(): number {
    return this.table().columnCount;
  }

  values(): string[][] {
    return this.table().values.map((row) => [...row]);
  }

  hasMergedCells(): boolean {
    return (this.table().mergedCells?.length ?? 0) > 0;
  }

  mergedCells(): PptTableCellMerge[] {
    return cloneModel(this.table().mergedCells ?? []);
  }

  setValues(values: string[][]): void {
    const table = this.table();
    this.assertNoMergedCells('set all table values');
    if (
      values.length !== table.rowCount ||
      values.some((row) => row.length !== table.columnCount)
    ) {
      throw new Error(
        `Presentation: table values must match ${table.rowCount}x${table.columnCount}`,
      );
    }
    for (let row = 0; row < values.length; row++) {
      for (let column = 0; column < values[row]!.length; column++) {
        this.writeCellText(row, column, values[row]![column]!);
      }
    }
    this.presentation._record({
      type: 'table.values.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      values,
    });
  }

  getCell(row: number, column: number): TableCellCore {
    return new TableCellCore(this.presentation, this.slideId, this.shapeId, row, column);
  }

  setCellText(row: number, column: number, text: string): void {
    this._setCellText(row, column, text);
  }

  addRow(index: number | 'end' = 'end', values?: string[]): void {
    const shape = this.internal();
    const table = this.table();
    this.assertNoMergedCells('add rows');
    const insertIndex = normalizeInsertIndex(index, table.rowCount);
    const rowValues = normalizeVector(values, table.columnCount, 'row');
    const tbl = getTableElement(shape);
    const rows = childElementsNs(tbl, A_NS, 'tr');
    const row = tableRowElement(tbl.ownerDocument, rowValues);
    tbl.insertBefore(row, rows[insertIndex] ?? null);
    markDirty(shape.slide.part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'table.row.add',
      slideId: this.slideId,
      shapeId: this.shapeId,
      index: insertIndex,
      values: rowValues,
    });
  }

  deleteRow(index: number): void {
    const shape = this.internal();
    const table = this.table();
    this.assertNoMergedCells('delete rows');
    if (table.rowCount <= 1) throw new Error('Presentation: cannot delete the final table row');
    const bounded = requireIndex(index, table.rowCount, 'row');
    const tbl = getTableElement(shape);
    const row = childElementsNs(tbl, A_NS, 'tr')[bounded];
    if (!row) throw new Error(`Presentation: table row out of range: ${index}`);
    tbl.removeChild(row);
    markDirty(shape.slide.part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'table.row.delete',
      slideId: this.slideId,
      shapeId: this.shapeId,
      index: bounded,
    });
  }

  addColumn(index: number | 'end' = 'end', values?: string[]): void {
    const shape = this.internal();
    const table = this.table();
    this.assertNoMergedCells('add columns');
    const insertIndex = normalizeInsertIndex(index, table.columnCount);
    const columnValues = normalizeVector(values, table.rowCount, 'column');
    const tbl = getTableElement(shape);
    const rows = childElementsNs(tbl, A_NS, 'tr');
    const tblGrid = firstChildNs(tbl, A_NS, 'tblGrid');
    if (!tblGrid) throw new Error(`Presentation: table ${this.shapeId} has no tblGrid`);
    const gridCols = childElementsNs(tblGrid, A_NS, 'gridCol');
    tblGrid.insertBefore(tableGridColElement(tbl.ownerDocument), gridCols[insertIndex] ?? null);
    for (let row = 0; row < rows.length; row++) {
      const cells = childElementsNs(rows[row]!, A_NS, 'tc');
      rows[row]!.insertBefore(
        tableCellElement(tbl.ownerDocument, columnValues[row] ?? ''),
        cells[insertIndex] ?? null,
      );
    }
    markDirty(shape.slide.part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'table.column.add',
      slideId: this.slideId,
      shapeId: this.shapeId,
      index: insertIndex,
      values: columnValues,
    });
  }

  deleteColumn(index: number): void {
    const shape = this.internal();
    const table = this.table();
    this.assertNoMergedCells('delete columns');
    if (table.columnCount <= 1)
      throw new Error('Presentation: cannot delete the final table column');
    const bounded = requireIndex(index, table.columnCount, 'column');
    const tbl = getTableElement(shape);
    const tblGrid = firstChildNs(tbl, A_NS, 'tblGrid');
    if (!tblGrid) throw new Error(`Presentation: table ${this.shapeId} has no tblGrid`);
    const gridCol = childElementsNs(tblGrid, A_NS, 'gridCol')[bounded];
    if (gridCol) tblGrid.removeChild(gridCol);
    for (const row of childElementsNs(tbl, A_NS, 'tr')) {
      const cell = childElementsNs(row, A_NS, 'tc')[bounded];
      if (cell) row.removeChild(cell);
    }
    markDirty(shape.slide.part);
    this.presentation._reloadFromPackage();
    this.presentation._record({
      type: 'table.column.delete',
      slideId: this.slideId,
      shapeId: this.shapeId,
      index: bounded,
    });
  }

  _cellText(row: number, column: number): string {
    const table = this.table();
    const [safeRow, safeColumn] = requireTableCellCoordinates(table, row, column);
    return table.values[safeRow]?.[safeColumn] ?? '';
  }

  _cellStyle(row: number, column: number): PptTableCellStyle {
    const table = this.table();
    const [safeRow, safeColumn] = requireTableCellCoordinates(table, row, column);
    return cloneModel(table.cellStyles?.[safeRow]?.[safeColumn] ?? {});
  }

  _cellParagraphs(row: number, column: number): PptTextParagraph[] {
    const shape = this.internal();
    const table = this.table();
    const [safeRow, safeColumn] = requireTableCellCoordinates(table, row, column);
    const cell = shape.tableCells[safeRow]?.[safeColumn];
    if (!cell) throw new Error(`Presentation: table cell out of range: ${row},${column}`);
    return textParagraphsFromElement(cell, shape.slide);
  }

  _setCellParagraphs(row: number, column: number, paragraphs: PptTextParagraphWrite[]): void {
    const shape = this.internal();
    const [safeRow, safeColumn] = this.assertEditableCell(row, column);
    const cell = this.requireCell(shape, safeRow, safeColumn);
    const textBody = firstChildNs(cell, A_NS, 'txBody');
    if (!textBody) throw new Error(`Presentation: table cell ${row},${column} has no text body`);
    replaceTextBodyParagraphs(textBody, paragraphs);
    shape.model.table!.values[safeRow]!.splice(
      safeColumn,
      1,
      textParagraphsFromElement(cell, shape.slide)
        .map((paragraph) => paragraph.text)
        .join('\n'),
    );
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'table.cell.text.paragraphs.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      row: safeRow,
      column: safeColumn,
      paragraphs,
    });
  }

  _setCellRunStyle(
    row: number,
    column: number,
    runIndex: number,
    style: Partial<PptTextStyle>,
  ): void {
    const shape = this.internal();
    const [safeRow, safeColumn] = this.assertEditableCell(row, column);
    const cell = this.requireCell(shape, safeRow, safeColumn);
    const runs = flatTextRuns(cell);
    const run = runs[requireIndex(runIndex, runs.length, 'run')];
    if (!run) throw new Error(`Presentation: text run out of range: ${runIndex}`);
    const normalized = normalizeTextStyle(style);
    applyTextStyleToRun(run, normalized);
    const table = shape.model.table!;
    table.cellStyles = ensureTableCellStyles(table.rowCount, table.columnCount, table.cellStyles);
    const rowStyles = table.cellStyles[safeRow]!;
    const previousStyle = rowStyles[safeColumn];
    rowStyles.splice(safeColumn, 1, {
      ...previousStyle,
      textStyle: { ...previousStyle?.textStyle, ...normalized },
    });
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'table.cell.text.run.style.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      row: safeRow,
      column: safeColumn,
      runIndex,
      style: normalized,
    });
  }

  _setCellText(row: number, column: number, text: string): void {
    this.writeCellText(row, column, text);
    this.presentation._record({
      type: 'table.cell.text.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      row,
      column,
      text,
    });
  }

  _setCellStyle(row: number, column: number, style: Partial<PptTableCellStyle>): void {
    const shape = this.internal();
    const [safeRow, safeColumn] = this.assertEditableCell(row, column);
    const cell = shape.tableCells[safeRow]?.[safeColumn];
    if (!cell) throw new Error(`Presentation: table cell out of range: ${row},${column}`);
    const normalized: Partial<PptTableCellStyle> = {
      ...(style.fillColor ? { fillColor: normalizeHexColor(style.fillColor) } : {}),
      ...(style.body ? { body: normalizeTableCellBodyProperties(style.body) } : {}),
      ...(style.textStyle ? { textStyle: normalizeTextStyle(style.textStyle) } : {}),
    };
    const tcPr =
      normalized.fillColor || normalized.body ? findOrCreateTableCellProperties(cell) : undefined;
    if (normalized.fillColor)
      setSolidFill(tcPr ?? findOrCreateTableCellProperties(cell), normalized.fillColor);
    if (normalized.body) applyTableCellBodyProperties(tcPr!, normalized.body);
    if (normalized.textStyle)
      applyTextStyleToTextElements(allDescendantsNs(cell, A_NS, 't'), normalized.textStyle);
    const table = shape.model.table!;
    table.cellStyles = ensureTableCellStyles(table.rowCount, table.columnCount, table.cellStyles);
    const rowStyles = table.cellStyles[safeRow]!;
    const previousStyle = rowStyles[safeColumn];
    rowStyles.splice(safeColumn, 1, {
      ...previousStyle,
      ...normalized,
      ...(normalized.body ? { body: { ...previousStyle?.body, ...normalized.body } } : {}),
      ...(normalized.textStyle
        ? { textStyle: { ...previousStyle?.textStyle, ...normalized.textStyle } }
        : {}),
    });
    markDirty(shape.slide.part);
    this.presentation._record({
      type: 'table.cell.style.set',
      slideId: this.slideId,
      shapeId: this.shapeId,
      row: safeRow,
      column: safeColumn,
      style: normalized,
    });
  }

  private writeCellText(row: number, column: number, text: string): void {
    const shape = this.internal();
    const [safeRow, safeColumn] = this.assertEditableCell(row, column);
    const cell = shape.tableCells[safeRow]?.[safeColumn];
    if (!cell) throw new Error(`Presentation: table cell out of range: ${row},${column}`);
    const textElements = tableCellTextElements(cell);
    if (textElements.length === 0) textElements.push(createTableCellTextRun(cell));
    textElements[0]!.textContent = text;
    for (const extra of textElements.slice(1)) extra.textContent = '';
    shape.model.table!.values[safeRow]!.splice(safeColumn, 1, text);
    markDirty(shape.slide.part);
  }

  private table() {
    const table = this.internal().model.table;
    if (!table) throw new Error(`Presentation: shape ${this.shapeId} is not a table`);
    return table;
  }

  private requireCell(shape: PptShapeInternal, row: number, column: number): Element {
    const table = this.table();
    const [safeRow, safeColumn] = requireTableCellCoordinates(table, row, column);
    const cell = shape.tableCells[safeRow]?.[safeColumn];
    if (!cell) throw new Error(`Presentation: table cell out of range: ${row},${column}`);
    return cell;
  }

  private assertNoMergedCells(operation: string): void {
    if (this.hasMergedCells()) {
      throw new Error(`Presentation: cannot ${operation} on a merged-cell table`);
    }
  }

  private assertEditableCell(row: number, column: number): [number, number] {
    const coordinates = requireTableCellCoordinates(this.table(), row, column);
    const [safeRow, safeColumn] = coordinates;
    const merge = this.table().mergedCells?.find(
      (item) => item.row === safeRow && item.column === safeColumn,
    );
    if (merge?.hMerge || merge?.vMerge) {
      throw new Error(
        `Presentation: cannot edit merged continuation cell ${safeRow},${safeColumn}`,
      );
    }
    return coordinates;
  }

  private internal(): PptShapeInternal {
    return getShapeInternal(this.presentation, this.slideId, this.shapeId);
  }
}

export class TableCellCore {
  constructor(
    private readonly presentation: PresentationCore,
    private readonly slideId: string,
    private readonly shapeId: string,
    private readonly row: number,
    private readonly column: number,
  ) {}

  text(): string {
    return new TableApiCore(this.presentation, this.slideId, this.shapeId)._cellText(
      this.row,
      this.column,
    );
  }

  paragraphs(): PptTextParagraph[] {
    return new TableApiCore(this.presentation, this.slideId, this.shapeId)._cellParagraphs(
      this.row,
      this.column,
    );
  }

  runs(): PptTextRun[] {
    return this.paragraphs().flatMap((paragraph) => paragraph.runs);
  }

  setParagraphs(paragraphs: PptTextParagraphWrite[]): void {
    new TableApiCore(this.presentation, this.slideId, this.shapeId)._setCellParagraphs(
      this.row,
      this.column,
      paragraphs,
    );
  }

  setRunStyle(runIndex: number, style: Partial<PptTextStyle>): void {
    new TableApiCore(this.presentation, this.slideId, this.shapeId)._setCellRunStyle(
      this.row,
      this.column,
      runIndex,
      style,
    );
  }

  setText(text: string): void {
    new TableApiCore(this.presentation, this.slideId, this.shapeId)._setCellText(
      this.row,
      this.column,
      text,
    );
  }

  style(): PptTableCellStyle {
    return new TableApiCore(this.presentation, this.slideId, this.shapeId)._cellStyle(
      this.row,
      this.column,
    );
  }

  setStyle(style: Partial<PptTableCellStyle>): void {
    new TableApiCore(this.presentation, this.slideId, this.shapeId)._setCellStyle(
      this.row,
      this.column,
      style,
    );
  }
}

function getShapeInternal(
  presentation: PresentationCore,
  slideId: string,
  shapeId: string,
): PptShapeInternal {
  const slide = presentation._getSlideInternal(slideId);
  const shape = slide ? findShapeInternal(slide.shapes, shapeId) : undefined;
  if (!shape) throw new Error(`Presentation: shape not found: ${shapeId}`);
  return shape;
}

function findShapeInternal(
  shapes: PptShapeInternal[],
  shapeId: string,
): PptShapeInternal | undefined {
  for (const shape of shapes) {
    if (shape.model.id === shapeId) return shape;
    const child = findShapeInternal(shape.children, shapeId);
    if (child) return child;
  }
  return undefined;
}

function exportedImageFromPart(presentation: PresentationCore, path: string): PptExportedImage {
  const part = getPart(presentation._package(), path);
  const bytes = part?.bytes ?? part?._originalBytes;
  if (!part || !bytes) throw new Error(`Presentation: image part not found: ${path}`);
  return {
    path,
    contentType: part.contentType,
    byteLength: bytes.byteLength,
    bytes: new Uint8Array(bytes),
  };
}

function exportedFontFromPart(presentation: PresentationCore, font: PptFontModel): PptExportedFont {
  const part = getPart(presentation._package(), font.path);
  const bytes = part?.bytes ?? part?._originalBytes;
  if (!part || !bytes) throw new Error(`Presentation: font part not found: ${font.path}`);
  return {
    ...font,
    contentType: part.contentType,
    byteLength: bytes.byteLength,
    bytes: new Uint8Array(bytes),
  };
}

function findTextInSlide(
  slide: PptReadResult['slides'][number],
  query: string | RegExp,
  options: PptFindTextOptions,
): PptTextMatch[] {
  const matches: PptTextMatch[] = [];
  for (const shape of flattenShapeModels(slide.model.shapes)) {
    if (shape.text && !shape.table) {
      matches.push(
        ...findMatchesInValue(shape.text.text, query, options).map((match) => ({
          slideId: slide.model.id,
          slideIndex: slide.model.index,
          shapeId: shape.id,
          shapeName: shape.name,
          source: 'shapeText' as const,
          text: shape.text!.text,
          start: match.start,
          length: match.length,
        })),
      );
    }
    if (shape.table) {
      for (let row = 0; row < shape.table.values.length; row++) {
        for (let column = 0; column < shape.table.values[row]!.length; column++) {
          const text = shape.table.values[row]![column]!;
          matches.push(
            ...findMatchesInValue(text, query, options).map((match) => ({
              slideId: slide.model.id,
              slideIndex: slide.model.index,
              shapeId: shape.id,
              shapeName: shape.name,
              source: 'tableCell' as const,
              text,
              start: match.start,
              length: match.length,
              row,
              column,
            })),
          );
        }
      }
    }
  }
  return matches;
}

function slideProperties(slide: PptSlideModel): PptSlideProperties {
  const shapes = flattenShapeModels(slide.shapes);
  return {
    id: slide.id,
    index: slide.index,
    name: slide.name,
    title: slide.title,
    shapeCount: shapes.length,
    textShapeCount: shapes.filter((shape) => Boolean(shape.text) && !shape.table).length,
    tableCount: shapes.filter((shape) => Boolean(shape.table)).length,
    pictureCount: shapes.filter((shape) => shape.kind === 'picture').length,
  };
}

function flattenShapeModels(shapes: PptShapeModel[]): PptShapeModel[] {
  return shapes.flatMap((shape) => [shape, ...flattenShapeModels(shape.children ?? [])]);
}

function findMatchesInValue(
  value: string,
  query: string | RegExp,
  options: PptFindTextOptions,
): Array<{ start: number; length: number }> {
  const regexp = textSearchRegExp(query, options);
  const matches: Array<{ start: number; length: number }> = [];
  let match: RegExpExecArray | null;
  while ((match = regexp.exec(value))) {
    matches.push({ start: match.index, length: match[0].length });
    if (match[0].length === 0) regexp.lastIndex++;
  }
  return matches;
}

function replaceTextValue(
  value: string,
  search: string | RegExp,
  replacement: string,
  options: PptTextReplaceOptions,
): { text: string; count: number } {
  const regexp = textSearchRegExp(search, options);
  let count = 0;
  const text = value.replace(regexp, () => {
    count++;
    return replacement;
  });
  return { text, count };
}

function textSearchRegExp(query: string | RegExp, options: PptFindTextOptions): RegExp {
  if (typeof query !== 'string') {
    const flags = query.flags.includes('g') ? query.flags : `${query.flags}g`;
    return new RegExp(query.source, flags);
  }
  return new RegExp(escapeRegExp(query), options.caseSensitive ? 'g' : 'gi');
}

function getTableElement(shape: PptShapeInternal): Element {
  const table = allDescendantsNs(shape.element, A_NS, 'tbl')[0];
  if (!table) throw new Error(`Presentation: shape ${shape.model.id} has no table XML`);
  return table;
}

function normalizeVector(
  values: string[] | undefined,
  length: number,
  axis: 'row' | 'column',
): string[] {
  if (values && values.length !== length) {
    throw new Error(`Presentation: ${axis} values must have length ${length}`);
  }
  return Array.from({ length }, (_, index) => values?.[index] ?? '');
}

function requireIndex(index: number, length: number, label: string): number {
  if (!Number.isInteger(index) || index < 0 || index >= length) {
    throw new Error(`Presentation: ${label} index out of range: ${index}`);
  }
  return index;
}

function requireTableCellCoordinates(
  table: { rowCount: number; columnCount: number },
  row: number,
  column: number,
): [number, number] {
  if (
    !Number.isInteger(row) ||
    !Number.isInteger(column) ||
    row < 0 ||
    row >= table.rowCount ||
    column < 0 ||
    column >= table.columnCount
  ) {
    throw new Error(`Presentation: table cell out of range: ${row},${column}`);
  }
  return [row, column];
}

function tableRowElement(doc: Document, values: string[]): Element {
  const row = doc.createElementNS(A_NS, 'a:tr');
  row.setAttribute('h', '370840');
  for (const value of values) row.appendChild(tableCellElement(doc, value));
  return row;
}

function tableCellElement(doc: Document, text: string): Element {
  const parsed = parseXmlDocument(
    `<root xmlns:a="${A_NS}">${tableCellXml(text)}</root>`,
  ) as unknown as Document;
  const cell = Array.from(parsed.documentElement.childNodes).find((node) => node.nodeType === 1);
  if (!cell) throw new Error('Presentation: generated table cell XML is empty');
  return doc.importNode(cell, true) as Element;
}

function tableGridColElement(doc: Document): Element {
  const gridCol = doc.createElementNS(A_NS, 'a:gridCol');
  gridCol.setAttribute('w', '1000000');
  return gridCol;
}

function findOrCreateShapeProperties(shapeElement: Element): Element {
  const existing = allDescendantsNs(shapeElement, P_NS, 'spPr')[0];
  if (existing) return existing;
  const spPr = shapeElement.ownerDocument.createElementNS(P_NS, 'p:spPr');
  const txBody = firstChildNs(shapeElement, P_NS, 'txBody');
  shapeElement.insertBefore(spPr, txBody ?? null);
  return spPr;
}

function findOrCreateLine(spPr: Element): Element {
  const existing = firstChildNs(spPr, A_NS, 'ln');
  if (existing) return existing;
  const ln = spPr.ownerDocument.createElementNS(A_NS, 'a:ln');
  spPr.appendChild(ln);
  return ln;
}

function findOrCreateRunProperties(run: Element): Element {
  const existing = firstChildNs(run, A_NS, 'rPr');
  if (existing) return existing;
  const rPr = run.ownerDocument.createElementNS(A_NS, 'a:rPr');
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}

function findOrCreateTableCellProperties(cell: Element): Element {
  const existing = firstChildNs(cell, A_NS, 'tcPr');
  if (existing) return existing;
  const tcPr = cell.ownerDocument.createElementNS(A_NS, 'a:tcPr');
  cell.appendChild(tcPr);
  return tcPr;
}

function tableCellTextElements(cell: Element): Element[] {
  return Array.from(cell.getElementsByTagNameNS(A_NS, 't')) as Element[];
}

function createTableCellTextRun(cell: Element): Element {
  const doc = cell.ownerDocument;
  let txBody = firstChildNs(cell, A_NS, 'txBody');
  if (!txBody) {
    txBody = doc.createElementNS(A_NS, 'a:txBody');
    txBody.appendChild(doc.createElementNS(A_NS, 'a:bodyPr'));
    const tcPr = firstChildNs(cell, A_NS, 'tcPr');
    cell.insertBefore(txBody, tcPr ?? null);
  }

  let paragraph = firstChildNs(txBody, A_NS, 'p');
  if (!paragraph) {
    paragraph = doc.createElementNS(A_NS, 'a:p');
    txBody.appendChild(paragraph);
  }

  const run = doc.createElementNS(A_NS, 'a:r');
  const textElement = doc.createElementNS(A_NS, 'a:t');
  run.appendChild(textElement);
  paragraph.appendChild(run);
  return textElement;
}

function findImageBlipFill(shapeElement: Element): Element | null {
  return (
    firstChildNs(shapeElement, P_NS, 'blipFill') ??
    allDescendantsNs(shapeElement, P_NS, 'blipFill')[0] ??
    null
  );
}

function findOrCreateImageSourceRect(blipFill: Element): Element {
  const existing = firstChildNs(blipFill, A_NS, 'srcRect');
  if (existing) return existing;
  const srcRect = blipFill.ownerDocument.createElementNS(A_NS, 'a:srcRect');
  const stretch = firstChildNs(blipFill, A_NS, 'stretch');
  blipFill.insertBefore(srcRect, stretch ?? null);
  return srcRect;
}

function findOrCreateImageStretch(blipFill: Element): Element {
  const existing = firstChildNs(blipFill, A_NS, 'stretch');
  if (existing) return existing;
  const stretch = blipFill.ownerDocument.createElementNS(A_NS, 'a:stretch');
  blipFill.appendChild(stretch);
  return stretch;
}

function findOrCreateImageFillRect(blipFill: Element): Element {
  const stretch = findOrCreateImageStretch(blipFill);
  const existing = firstChildNs(stretch, A_NS, 'fillRect');
  if (existing) return existing;
  const fillRect = blipFill.ownerDocument.createElementNS(A_NS, 'a:fillRect');
  stretch.appendChild(fillRect);
  return fillRect;
}

function findOrCreateParagraphProperties(paragraph: Element): Element {
  const existing = firstChildNs(paragraph, A_NS, 'pPr');
  if (existing) return existing;
  const pPr = paragraph.ownerDocument.createElementNS(A_NS, 'a:pPr');
  paragraph.insertBefore(pPr, paragraph.firstChild);
  return pPr;
}

function closestAncestor(element: Element, ns: string, localName: string): Element | null {
  let current = element.parentNode;
  while (current?.nodeType === 1) {
    const el = current as Element;
    if (el.namespaceURI === ns && el.localName === localName) return el;
    current = el.parentNode;
  }
  return null;
}

function alignmentToOoxml(alignment: NonNullable<PptTextStyle['alignment']>): string {
  if (alignment === 'center') return 'ctr';
  if (alignment === 'right') return 'r';
  if (alignment === 'justify') return 'just';
  if (alignment === 'distributed') return 'dist';
  return 'l';
}

function normalizeTextStyle(style: Partial<PptTextStyle>): Partial<PptTextStyle> {
  return {
    ...(style.fontSize !== undefined ? { fontSize: style.fontSize } : {}),
    ...(style.fontFamily ? { fontFamily: style.fontFamily } : {}),
    ...(style.color ? { color: normalizeHexColor(style.color) } : {}),
    ...(style.bold !== undefined ? { bold: style.bold } : {}),
    ...(style.italic !== undefined ? { italic: style.italic } : {}),
    ...(style.underline !== undefined ? { underline: style.underline } : {}),
    ...(style.strike !== undefined ? { strike: style.strike } : {}),
    ...(style.baseline !== undefined ? { baseline: normalizeTextBaseline(style.baseline) } : {}),
    ...(style.highlight ? { highlight: normalizeHexColor(style.highlight) } : {}),
    ...(style.alignment ? { alignment: style.alignment } : {}),
    ...(style.lineHeight !== undefined ? { lineHeight: style.lineHeight } : {}),
    ...(style.characterSpacing !== undefined ? { characterSpacing: style.characterSpacing } : {}),
  };
}

function normalizeTableCellBodyProperties(
  body: Partial<PptTextBodyProperties>,
): Partial<PptTextBodyProperties> {
  if (body.autoFit !== undefined)
    throw new Error('Presentation: table cell body autoFit is not writable');
  return {
    ...(body.anchor !== undefined ? { anchor: normalizeTextBodyAnchor(body.anchor) } : {}),
    ...(body.wrap !== undefined ? { wrap: normalizeTextBodyWrap(body.wrap) } : {}),
    ...(body.insetLeft !== undefined
      ? { insetLeft: normalizeNonNegativeNumber(body.insetLeft, 'cell body insetLeft') }
      : {}),
    ...(body.insetTop !== undefined
      ? { insetTop: normalizeNonNegativeNumber(body.insetTop, 'cell body insetTop') }
      : {}),
    ...(body.insetRight !== undefined
      ? { insetRight: normalizeNonNegativeNumber(body.insetRight, 'cell body insetRight') }
      : {}),
    ...(body.insetBottom !== undefined
      ? { insetBottom: normalizeNonNegativeNumber(body.insetBottom, 'cell body insetBottom') }
      : {}),
  };
}

function normalizeImageStyle(style: Partial<PptImageStyle>): Partial<PptImageStyle> {
  return {
    ...(style.crop ? { crop: normalizeImageRect(style.crop, 'crop') } : {}),
    ...(style.fillRect ? { fillRect: normalizeImageRect(style.fillRect, 'fillRect') } : {}),
    ...(style.opacity !== undefined ? { opacity: normalizeOpacity(style.opacity) } : {}),
  };
}

function normalizeImageRect(
  rect: PptImageStyle['crop'],
  name: 'crop' | 'fillRect',
): NonNullable<PptImageStyle['crop']> {
  if (!rect || Object.keys(rect).length === 0)
    throw new Error(`Presentation: image ${name} must include at least one edge`);
  return {
    ...(rect.left !== undefined ? { left: normalizeImageCropEdge(rect.left, `${name} left`) } : {}),
    ...(rect.top !== undefined ? { top: normalizeImageCropEdge(rect.top, `${name} top`) } : {}),
    ...(rect.right !== undefined
      ? { right: normalizeImageCropEdge(rect.right, `${name} right`) }
      : {}),
    ...(rect.bottom !== undefined
      ? { bottom: normalizeImageCropEdge(rect.bottom, `${name} bottom`) }
      : {}),
  };
}

function normalizeImageCropEdge(value: number, name: string): number {
  if (!Number.isFinite(value) || value < 0 || value > 1)
    throw new Error(`Presentation: image ${name} must be between 0 and 1: ${value}`);
  return value;
}

function normalizeTextBaseline(value: number): number {
  if (!Number.isFinite(value))
    throw new Error(`Presentation: text baseline must be a finite number: ${value}`);
  return value;
}

function normalizeTextBodyAnchor(
  anchor: NonNullable<PptTextBodyProperties['anchor']>,
): NonNullable<PptTextBodyProperties['anchor']> {
  if (anchor === 'top' || anchor === 'middle' || anchor === 'bottom') return anchor;
  throw new Error(`Presentation: invalid text body anchor ${anchor}`);
}

function normalizeTextBodyWrap(
  wrap: NonNullable<PptTextBodyProperties['wrap']>,
): NonNullable<PptTextBodyProperties['wrap']> {
  if (wrap === 'none' || wrap === 'square') return wrap;
  throw new Error(`Presentation: invalid text body wrap ${wrap}`);
}

function normalizeTextBodyVertical(
  vertical: NonNullable<PptTextBodyProperties['vertical']>,
): NonNullable<PptTextBodyProperties['vertical']> {
  if (vertical === 'eastAsian' || vertical === 'vertical') return vertical;
  throw new Error(`Presentation: invalid text body vertical ${vertical}`);
}

function normalizeTextBodyAutoFit(
  autoFit: NonNullable<PptTextBodyProperties['autoFit']>,
): NonNullable<PptTextBodyProperties['autoFit']> {
  if (autoFit === 'none' || autoFit === 'shape' || autoFit === 'normal') return autoFit;
  throw new Error(`Presentation: invalid text body autoFit ${autoFit}`);
}

function normalizeShapeTextBodyProperties(
  body: Partial<PptTextBodyProperties>,
): Partial<PptTextBodyProperties> {
  return {
    ...(body.anchor !== undefined ? { anchor: normalizeTextBodyAnchor(body.anchor) } : {}),
    ...(body.autoFit !== undefined ? { autoFit: normalizeTextBodyAutoFit(body.autoFit) } : {}),
    ...(body.wrap !== undefined ? { wrap: normalizeTextBodyWrap(body.wrap) } : {}),
    ...(body.vertical !== undefined ? { vertical: normalizeTextBodyVertical(body.vertical) } : {}),
    ...normalizeTextBodyInsets(body),
  };
}

function normalizeTextBodyInsets(
  body: Partial<PptTextBodyProperties>,
): Partial<PptTextBodyProperties> {
  return Object.fromEntries(
    (['insetLeft', 'insetTop', 'insetRight', 'insetBottom'] as const)
      .filter((key) => body[key] !== undefined)
      .map((key) => [key, normalizeNonNegativeNumber(body[key]!, `text body ${key}`)]),
  ) as Partial<PptTextBodyProperties>;
}

function findShapeTextBody(shapeElement: Element): Element | null {
  return (
    firstChildNs(shapeElement, P_NS, 'txBody') ??
    allDescendantsNs(shapeElement, P_NS, 'txBody')[0] ??
    allDescendantsNs(shapeElement, A_NS, 'txBody')[0] ??
    null
  );
}

function findOrCreateTextBodyProperties(textBody: Element): Element {
  const existing = firstChildNs(textBody, A_NS, 'bodyPr');
  if (existing) return existing;
  const bodyPr = textBody.ownerDocument.createElementNS(A_NS, 'a:bodyPr');
  textBody.insertBefore(bodyPr, firstChildNs(textBody, A_NS, 'lstStyle') ?? textBody.firstChild);
  return bodyPr;
}

function applyShapeTextBodyProperties(bodyPr: Element, body: Partial<PptTextBodyProperties>): void {
  if (body.anchor !== undefined) bodyPr.setAttribute('anchor', textBodyAnchorToOoxml(body.anchor));
  if (body.wrap !== undefined) bodyPr.setAttribute('wrap', textBodyWrapToOoxml(body.wrap));
  if (body.vertical !== undefined)
    bodyPr.setAttribute('vert', textBodyVerticalToOoxml(body.vertical));
  if (body.insetLeft !== undefined) bodyPr.setAttribute('lIns', String(body.insetLeft));
  if (body.insetTop !== undefined) bodyPr.setAttribute('tIns', String(body.insetTop));
  if (body.insetRight !== undefined) bodyPr.setAttribute('rIns', String(body.insetRight));
  if (body.insetBottom !== undefined) bodyPr.setAttribute('bIns', String(body.insetBottom));
  if (body.autoFit !== undefined) {
    removeTextAutoFitChildren(bodyPr);
    bodyPr.appendChild(textAutoFitElement(bodyPr.ownerDocument, body.autoFit));
  }
}

function removeTextAutoFitChildren(bodyPr: Element): void {
  const autoFitChildren = Array.from(bodyPr.childNodes).filter(
    (child): child is Element =>
      child.nodeType === 1 &&
      (child as Element).namespaceURI === A_NS &&
      ((child as Element).localName === 'spAutoFit' ||
        (child as Element).localName === 'normAutofit' ||
        (child as Element).localName === 'noAutofit'),
  );
  for (const child of autoFitChildren) {
    bodyPr.removeChild(child);
  }
}

function textAutoFitElement(
  doc: Document,
  autoFit: NonNullable<PptTextBodyProperties['autoFit']>,
): Element {
  const tag =
    autoFit === 'shape' ? 'spAutoFit' : autoFit === 'normal' ? 'normAutofit' : 'noAutofit';
  return doc.createElementNS(A_NS, `a:${tag}`);
}

function replaceTextBodyParagraphs(textBody: Element, paragraphs: PptTextParagraphWrite[]): void {
  const doc = textBody.ownerDocument;
  const normalized = paragraphs.length > 0 ? paragraphs : [{ text: '' }];
  for (const paragraph of childElementsNs(textBody, A_NS, 'p')) {
    textBody.removeChild(paragraph);
  }
  for (const paragraph of normalized) {
    textBody.appendChild(textParagraphElement(doc, paragraph));
  }
}

function textParagraphElement(doc: Document, paragraph: PptTextParagraphWrite): Element {
  const p = doc.createElementNS(A_NS, 'a:p');
  if (paragraph.style) applyParagraphStyle(p, paragraph.style);
  const runs =
    paragraph.runs && paragraph.runs.length > 0 ? paragraph.runs : [{ text: paragraph.text ?? '' }];
  for (const run of runs) {
    for (const element of textRunElements(doc, run)) p.appendChild(element);
  }
  return p;
}

function textRunElements(
  doc: Document,
  run: { text: string; style?: Partial<PptTextStyle> },
): Element[] {
  const parts = run.text.split('\n');
  const elements: Element[] = [];
  parts.forEach((part, index) => {
    if (index > 0) elements.push(doc.createElementNS(A_NS, 'a:br'));
    if (part.length > 0 || parts.length === 1) elements.push(textRunElement(doc, part, run.style));
  });
  return elements;
}

function textRunElement(
  doc: Document,
  textValue: string,
  style: Partial<PptTextStyle> | undefined,
): Element {
  const r = doc.createElementNS(A_NS, 'a:r');
  if (style && Object.keys(style).length > 0) applyTextStyleToRun(r, normalizeTextStyle(style));
  const text = doc.createElementNS(A_NS, 'a:t');
  text.textContent = textValue;
  if (/^\s|\s$| {2,}/.test(textValue)) text.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  r.appendChild(text);
  return r;
}

function flatTextRuns(element: Element): Element[] {
  return allDescendantsNs(element, A_NS, 'p').flatMap((paragraph) =>
    childElementsNs(paragraph, A_NS, 'r'),
  );
}

function textParagraphsFromElement(
  element: Element,
  slide?: PptReadResult['slides'][number],
): PptTextParagraph[] {
  return allDescendantsNs(element, A_NS, 'p').map((paragraph) => {
    const pPr = firstChildNs(paragraph, A_NS, 'pPr');
    const alignment = parseParagraphAlignment(pPr?.getAttribute('algn') ?? null);
    const lineHeight = parseParagraphLineHeight(pPr);
    const bullet = parseParagraphBullet(pPr);
    const marginLeft = parseNumber(pPr?.getAttribute('marL') ?? null);
    const textIndent = parseNumber(pPr?.getAttribute('indent') ?? null);
    const spaceBefore = parseParagraphSpacing(pPr ? firstChildNs(pPr, A_NS, 'spcBef') : null);
    const spaceAfter = parseParagraphSpacing(pPr ? firstChildNs(pPr, A_NS, 'spcAft') : null);
    const paragraphStyle = {
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
    const runs = textRunsFromParagraph(paragraph, alignment, slide);
    return {
      text: runs.map((run) => run.text).join(''),
      ...(Object.keys(paragraphStyle).length > 0 ? { style: paragraphStyle } : {}),
      runs,
    };
  });
}

function textRunsFromParagraph(
  paragraph: Element,
  alignment: PptTextStyle['alignment'],
  slide: PptReadResult['slides'][number] | undefined,
): PptTextRun[] {
  const runs: PptTextRun[] = [];
  for (let i = 0; i < paragraph.childNodes.length; i++) {
    const child = paragraph.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI !== A_NS) continue;
    if (element.localName === 'r') {
      const text = allDescendantsNs(element, A_NS, 't').map(textOf).join('');
      const style = textStyleFromRun(element, alignment);
      const hyperlink = slide ? hyperlinkFromRun(element, slide) : undefined;
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

function hyperlinkFromRun(
  run: Element,
  slide: PptReadResult['slides'][number],
): PptHyperlinkModel | undefined {
  const rPr = firstChildNs(run, A_NS, 'rPr');
  return hyperlinkFromElement(slide, rPr ? firstChildNs(rPr, A_NS, 'hlinkClick') : null);
}

function hyperlinkFromElement(
  slide: PptReadResult['slides'][number],
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

function findOrCreateHyperlink(parent: Element): Element {
  const existing = firstChildNs(parent, A_NS, 'hlinkClick');
  if (existing) return existing;
  const hlink = parent.ownerDocument.createElementNS(A_NS, 'a:hlinkClick');
  parent.appendChild(hlink);
  return hlink;
}

function setHyperlinkOnElement(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
  hlink: Element,
  hyperlink: PptHyperlinkWriteOptions,
): PptHyperlinkModel {
  if (!hyperlink.target) throw new Error('Presentation: hyperlink target is required');
  const existingRelId = hlink.getAttributeNS(R_NS, 'id') ?? hlink.getAttribute('r:id');
  if (existingRelId)
    removeSlideRelationshipIfUnreferenced(presentation, slide, existingRelId, hlink);
  const targetMode = hyperlink.targetMode ?? 'External';
  const relId = addSlideRelationship(presentation, slide, {
    type: `${OFFICE_REL_NS}/hyperlink`,
    target: hyperlink.target,
    targetMode,
  });
  hlink.setAttributeNS(R_NS, 'r:id', relId);
  setOptionalAttr(hlink, 'action', hyperlink.action);
  setOptionalAttr(hlink, 'tooltip', hyperlink.tooltip);
  slide.relationships.set(relId, {
    id: relId,
    type: `${OFFICE_REL_NS}/hyperlink`,
    target: hyperlink.target,
    targetMode,
  });
  return {
    relationshipId: relId,
    target: hyperlink.target,
    targetMode,
    ...(hyperlink.action ? { action: hyperlink.action } : {}),
    ...(hyperlink.tooltip ? { tooltip: hyperlink.tooltip } : {}),
  };
}

function clearHyperlinkElement(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
  hlink: Element | null,
): void {
  if (!hlink) return;
  const relId = hlink.getAttributeNS(R_NS, 'id') ?? hlink.getAttribute('r:id');
  hlink.parentNode?.removeChild(hlink);
  if (relId) removeSlideRelationshipIfUnreferenced(presentation, slide, relId);
}

function setOptionalAttr(element: Element, name: string, value: string | undefined): void {
  if (value === undefined) element.removeAttribute(name);
  else element.setAttribute(name, value);
}

function textStyleFromRun(run: Element, alignment: PptTextStyle['alignment']): PptTextStyle {
  const rPr = firstChildNs(run, A_NS, 'rPr');
  const size = parseNumber(rPr?.getAttribute('sz') ?? null);
  const fontFamily = rPr ? firstTypeface(rPr) : undefined;
  const color = rPr ? colorFromRunProperties(rPr) : undefined;
  const bold = parseBooleanAttr(rPr?.getAttribute('b') ?? null);
  const italic = parseBooleanAttr(rPr?.getAttribute('i') ?? null);
  const underline = parseUnderlineAttr(rPr?.getAttribute('u') ?? null);
  const strike = parseStrikeAttr(rPr?.getAttribute('strike') ?? null);
  const baseline = parseTextBaseline(rPr);
  const highlight = rPr ? highlightColorFromRunProperties(rPr) : undefined;
  const characterSpacing = parseCharacterSpacing(rPr);
  return {
    ...(size !== undefined ? { fontSize: size / 100 } : {}),
    ...(fontFamily ? { fontFamily } : {}),
    ...(color ? { color } : {}),
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

function firstTypeface(rPr: Element): string | undefined {
  return (
    firstChildNs(rPr, A_NS, 'ea')?.getAttribute('typeface') ??
    firstChildNs(rPr, A_NS, 'latin')?.getAttribute('typeface') ??
    firstChildNs(rPr, A_NS, 'cs')?.getAttribute('typeface') ??
    undefined
  );
}

function colorFromRunProperties(rPr: Element): string | undefined {
  const solidFill = firstChildNs(rPr, A_NS, 'solidFill');
  const srgbClr = solidFill ? firstChildNs(solidFill, A_NS, 'srgbClr') : null;
  const value = srgbClr?.getAttribute('val');
  if (value) return `#${value.toUpperCase()}`;
  const gradientFill = firstChildNs(rPr, A_NS, 'gradFill');
  const firstStop = gradientFill ? firstChildNs(gradientFill, A_NS, 'gs') : null;
  const gradientSrgbClr = firstStop ? firstChildNs(firstStop, A_NS, 'srgbClr') : null;
  const gradientValue = gradientSrgbClr?.getAttribute('val');
  return gradientValue ? `#${gradientValue.toUpperCase()}` : undefined;
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

function parseCharacterSpacing(rPr: Element | null): number | undefined {
  const value = parseNumber(rPr?.getAttribute('spc') ?? null);
  return value !== undefined ? value / 100 : undefined;
}

function parseTextBaseline(rPr: Element | null): number | undefined {
  const value = parseNumber(rPr?.getAttribute('baseline') ?? null);
  return value !== undefined ? value / 100000 : undefined;
}

function highlightColorFromRunProperties(rPr: Element): string | undefined {
  const highlight = firstChildNs(rPr, A_NS, 'highlight');
  const srgbClr = highlight ? firstChildNs(highlight, A_NS, 'srgbClr') : null;
  const value = srgbClr?.getAttribute('val');
  return value ? `#${value.toUpperCase()}` : undefined;
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
): { char: string; color?: string; fontFamily?: string; size?: number } | undefined {
  if (!pPr || firstChildNs(pPr, A_NS, 'buNone')) return undefined;
  const typeface = firstChildNs(pPr, A_NS, 'buFont')?.getAttribute('typeface') ?? undefined;
  const color = parseBulletColor(pPr);
  const size = parseBulletSize(pPr);
  const buChar = firstChildNs(pPr, A_NS, 'buChar')?.getAttribute('char');
  if (!buChar) return undefined;
  return {
    char: normalizeBulletCharacter(buChar, typeface),
    ...(color ? { color } : {}),
    ...(typeface ? { fontFamily: typeface } : {}),
    ...(size !== undefined ? { size } : {}),
  };
}

function parseBulletSize(pPr: Element): number | undefined {
  const pct = parseNumber(firstChildNs(pPr, A_NS, 'buSzPct')?.getAttribute('val') ?? null);
  if (pct !== undefined) return pct / 100000;
  const pts = parseNumber(firstChildNs(pPr, A_NS, 'buSzPts')?.getAttribute('val') ?? null);
  return pts !== undefined ? pts / 100 : undefined;
}

function parseBulletColor(pPr: Element): string | undefined {
  const buClr = firstChildNs(pPr, A_NS, 'buClr');
  if (!buClr) return undefined;
  const srgbClr = firstChildNs(buClr, A_NS, 'srgbClr');
  const value = srgbClr?.getAttribute('val');
  return value ? `#${value.toUpperCase()}` : undefined;
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

function parseParagraphAlignment(value: string | null): PptTextStyle['alignment'] | undefined {
  if (value === 'ctr') return 'center';
  if (value === 'r') return 'right';
  if (value === 'just') return 'justify';
  if (value === 'dist' || value === 'thaiDist') return 'distributed';
  if (value === 'l') return 'left';
  return undefined;
}

function applyTextStyleToRun(run: Element, style: Partial<PptTextStyle>): void {
  const rPr = findOrCreateRunProperties(run);
  if (style.fontSize !== undefined)
    rPr.setAttribute('sz', String(Math.round(style.fontSize * 100)));
  if (style.fontFamily) setRunTypeface(rPr, style.fontFamily);
  if (style.color) setSolidFill(rPr, style.color);
  if (style.bold !== undefined) rPr.setAttribute('b', style.bold ? '1' : '0');
  if (style.italic !== undefined) rPr.setAttribute('i', style.italic ? '1' : '0');
  if (style.underline !== undefined) rPr.setAttribute('u', style.underline ? 'sng' : 'none');
  if (style.strike !== undefined)
    rPr.setAttribute('strike', style.strike ? 'sngStrike' : 'noStrike');
  if (style.baseline !== undefined)
    rPr.setAttribute('baseline', String(Math.round(style.baseline * 100000)));
  if (style.highlight) setRunHighlight(rPr, style.highlight);
  if (style.characterSpacing !== undefined)
    rPr.setAttribute('spc', String(Math.round(style.characterSpacing * 100)));
  if (style.alignment) {
    const paragraph = closestAncestor(run, A_NS, 'p');
    if (paragraph)
      findOrCreateParagraphProperties(paragraph).setAttribute(
        'algn',
        alignmentToOoxml(style.alignment),
      );
  }
  if (style.lineHeight !== undefined) {
    const paragraph = closestAncestor(run, A_NS, 'p');
    if (paragraph)
      setParagraphLineHeight(findOrCreateParagraphProperties(paragraph), style.lineHeight);
  }
}

function setParagraphLineHeight(pPr: Element, lineHeight: number): void {
  let lnSpc = firstChildNs(pPr, A_NS, 'lnSpc');
  if (!lnSpc) {
    lnSpc = pPr.ownerDocument.createElementNS(A_NS, 'a:lnSpc');
    pPr.appendChild(lnSpc);
  }
  for (const child of childElementsNs(lnSpc, A_NS, 'spcPct')) {
    lnSpc.removeChild(child);
  }
  for (const child of childElementsNs(lnSpc, A_NS, 'spcPts')) {
    lnSpc.removeChild(child);
  }
  const spcPct = pPr.ownerDocument.createElementNS(A_NS, 'a:spcPct');
  spcPct.setAttribute('val', String(Math.round(lineHeight * 100000)));
  lnSpc.appendChild(spcPct);
}

function applyParagraphStyle(
  paragraph: Element,
  style: NonNullable<PptTextParagraphWrite['style']>,
): void {
  const pPr = findOrCreateParagraphProperties(paragraph);
  if (style.alignment) pPr.setAttribute('algn', alignmentToOoxml(style.alignment));
  if (style.lineHeight !== undefined) setParagraphLineHeight(pPr, style.lineHeight);
  if (style.marginLeft !== undefined)
    pPr.setAttribute('marL', String(Math.round(style.marginLeft)));
  if (style.textIndent !== undefined)
    pPr.setAttribute('indent', String(Math.round(style.textIndent)));
  if (style.spaceBefore !== undefined) setParagraphSpacing(pPr, 'spcBef', style.spaceBefore);
  if (style.spaceAfter !== undefined) setParagraphSpacing(pPr, 'spcAft', style.spaceAfter);
  if (style.bullet) setParagraphBullet(pPr, style);
}

function setParagraphSpacing(pPr: Element, localName: 'spcBef' | 'spcAft', value: number): void {
  let spacing = firstChildNs(pPr, A_NS, localName);
  if (!spacing) {
    spacing = pPr.ownerDocument.createElementNS(A_NS, `a:${localName}`);
    pPr.appendChild(spacing);
  }
  for (const child of Array.from(spacing.childNodes)) spacing.removeChild(child);
  const spcPts = pPr.ownerDocument.createElementNS(A_NS, 'a:spcPts');
  spcPts.setAttribute('val', String(Math.round(value / 127)));
  spacing.appendChild(spcPts);
}

function setParagraphBullet(
  pPr: Element,
  style: Pick<
    NonNullable<PptTextParagraphWrite['style']>,
    'bullet' | 'bulletColor' | 'bulletFontFamily' | 'bulletSize'
  >,
): void {
  if (!style.bullet) return;
  for (const localName of [
    'buNone',
    'buAutoNum',
    'buChar',
    'buClr',
    'buFont',
    'buSzPct',
    'buSzPts',
  ]) {
    for (const child of childElementsNs(pPr, A_NS, localName)) pPr.removeChild(child);
  }
  if (style.bulletSize !== undefined) {
    const size = pPr.ownerDocument.createElementNS(
      A_NS,
      style.bulletSize > 4 ? 'a:buSzPts' : 'a:buSzPct',
    );
    size.setAttribute(
      'val',
      String(Math.round(style.bulletSize > 4 ? style.bulletSize * 100 : style.bulletSize * 100000)),
    );
    pPr.appendChild(size);
  }
  if (style.bulletColor) {
    const buClr = pPr.ownerDocument.createElementNS(A_NS, 'a:buClr');
    const srgbClr = pPr.ownerDocument.createElementNS(A_NS, 'a:srgbClr');
    srgbClr.setAttribute('val', normalizeHexColor(style.bulletColor).slice(1));
    buClr.appendChild(srgbClr);
    pPr.appendChild(buClr);
  }
  if (style.bulletFontFamily) {
    const buFont = pPr.ownerDocument.createElementNS(A_NS, 'a:buFont');
    buFont.setAttribute('typeface', style.bulletFontFamily);
    pPr.appendChild(buFont);
  }
  const buChar = pPr.ownerDocument.createElementNS(A_NS, 'a:buChar');
  buChar.setAttribute('char', style.bullet);
  pPr.appendChild(buChar);
}

function setRunTypeface(rPr: Element, fontFamily: string): void {
  for (const localName of ['latin', 'ea', 'cs']) {
    let font = firstChildNs(rPr, A_NS, localName);
    if (!font) {
      font = rPr.ownerDocument.createElementNS(A_NS, `a:${localName}`);
      rPr.appendChild(font);
    }
    font.setAttribute('typeface', fontFamily);
  }
}

function setRunHighlight(rPr: Element, color: string): void {
  for (const child of childElementsNs(rPr, A_NS, 'highlight')) rPr.removeChild(child);
  const highlight = rPr.ownerDocument.createElementNS(A_NS, 'a:highlight');
  const srgbClr = rPr.ownerDocument.createElementNS(A_NS, 'a:srgbClr');
  srgbClr.setAttribute('val', color.slice(1));
  highlight.appendChild(srgbClr);
  rPr.appendChild(highlight);
}

function applyTextStyleToTextElements(textElements: Element[], style: Partial<PptTextStyle>): void {
  for (const textEl of textElements) {
    const run = textEl.parentNode;
    if (!run || run.nodeType !== 1) continue;
    applyTextStyleToRun(run as Element, style);
  }
}

function ensureTableCellStyles(
  rowCount: number,
  columnCount: number,
  styles: PptTableCellStyle[][] | undefined,
): PptTableCellStyle[][] {
  return Array.from({ length: rowCount }, (_, row) =>
    Array.from({ length: columnCount }, (_, column) => ({ ...styles?.[row]?.[column] })),
  );
}

function applyTableCellBodyProperties(tcPr: Element, body: Partial<PptTextBodyProperties>): void {
  if (body.anchor !== undefined) tcPr.setAttribute('anchor', textBodyAnchorToOoxml(body.anchor));
  if (body.wrap !== undefined) tcPr.setAttribute('wrap', body.wrap);
  if (body.insetLeft !== undefined) tcPr.setAttribute('marL', String(Math.round(body.insetLeft)));
  if (body.insetTop !== undefined) tcPr.setAttribute('marT', String(Math.round(body.insetTop)));
  if (body.insetRight !== undefined) tcPr.setAttribute('marR', String(Math.round(body.insetRight)));
  if (body.insetBottom !== undefined)
    tcPr.setAttribute('marB', String(Math.round(body.insetBottom)));
}

function textBodyAnchorToOoxml(anchor: NonNullable<PptTextBodyProperties['anchor']>): string {
  if (anchor === 'middle') return 'ctr';
  if (anchor === 'bottom') return 'b';
  return 't';
}

function textBodyWrapToOoxml(wrap: NonNullable<PptTextBodyProperties['wrap']>): string {
  return wrap;
}

function textBodyVerticalToOoxml(vertical: NonNullable<PptTextBodyProperties['vertical']>): string {
  return vertical === 'eastAsian' ? 'eaVert' : 'vert';
}

function setImageCrop(blipFill: Element, crop: NonNullable<PptImageStyle['crop']>): void {
  const srcRect = findOrCreateImageSourceRect(blipFill);
  setImageRectAttributes(srcRect, crop);
}

function setImageFillRect(
  blipFill: Element,
  fillRect: NonNullable<PptImageStyle['fillRect']>,
): void {
  const element = findOrCreateImageFillRect(blipFill);
  setImageRectAttributes(element, fillRect);
}

function setImageRectAttributes(element: Element, rect: NonNullable<PptImageStyle['crop']>): void {
  if (rect.left !== undefined) element.setAttribute('l', String(Math.round(rect.left * 100000)));
  if (rect.top !== undefined) element.setAttribute('t', String(Math.round(rect.top * 100000)));
  if (rect.right !== undefined) element.setAttribute('r', String(Math.round(rect.right * 100000)));
  if (rect.bottom !== undefined)
    element.setAttribute('b', String(Math.round(rect.bottom * 100000)));
}

function setImageOpacity(blip: Element, opacity: number): void {
  for (const child of childElementsNs(blip, A_NS, 'alphaModFix')) blip.removeChild(child);
  const alpha = blip.ownerDocument.createElementNS(A_NS, 'a:alphaModFix');
  alpha.setAttribute('amt', String(Math.round(opacity * 100000)));
  blip.appendChild(alpha);
}

function setSolidFill(parent: Element, color: string, opacity?: number): void {
  for (const localName of ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']) {
    for (const child of Array.from(parent.getElementsByTagNameNS(A_NS, localName))) {
      if (child.parentNode === parent) parent.removeChild(child);
    }
  }
  const solidFill = parent.ownerDocument.createElementNS(A_NS, 'a:solidFill');
  const srgbClr = parent.ownerDocument.createElementNS(A_NS, 'a:srgbClr');
  srgbClr.setAttribute('val', color.slice(1));
  if (opacity !== undefined) {
    const alpha = parent.ownerDocument.createElementNS(A_NS, 'a:alpha');
    alpha.setAttribute('val', String(Math.round(opacity * 100000)));
    srgbClr.appendChild(alpha);
  }
  solidFill.appendChild(srgbClr);
  parent.appendChild(solidFill);
}

function setExistingSolidFillOpacity(
  parent: Element,
  opacity: number,
  fallbackColor?: string,
): void {
  const solidFill = firstChildNs(parent, A_NS, 'solidFill');
  const srgbClr = solidFill ? firstChildNs(solidFill, A_NS, 'srgbClr') : null;
  if (!solidFill || !srgbClr) {
    if (fallbackColor) setSolidFill(parent, fallbackColor, opacity);
    return;
  }
  for (const child of childElementsNs(srgbClr, A_NS, 'alpha')) srgbClr.removeChild(child);
  const alpha = parent.ownerDocument.createElementNS(A_NS, 'a:alpha');
  alpha.setAttribute('val', String(Math.round(opacity * 100000)));
  srgbClr.appendChild(alpha);
}

function setOuterShadow(spPr: Element, shadow: NonNullable<PptShapeStyle['shadow']>): void {
  let effectList = firstChildNs(spPr, A_NS, 'effectLst');
  if (!effectList) {
    effectList = spPr.ownerDocument.createElementNS(A_NS, 'a:effectLst');
    spPr.appendChild(effectList);
  }
  for (const child of childElementsNs(effectList, A_NS, 'outerShdw')) {
    effectList.removeChild(child);
  }
  const outerShadow = spPr.ownerDocument.createElementNS(A_NS, 'a:outerShdw');
  if (shadow.blur !== undefined)
    outerShadow.setAttribute('blurRad', String(Math.round(shadow.blur)));
  if (shadow.distance !== undefined)
    outerShadow.setAttribute('dist', String(Math.round(shadow.distance)));
  if (shadow.direction !== undefined)
    outerShadow.setAttribute('dir', String(Math.round(shadow.direction * 60000)));
  const color = shadow.color ?? (shadow.opacity !== undefined ? '#000000' : undefined);
  if (color) {
    const srgbClr = spPr.ownerDocument.createElementNS(A_NS, 'a:srgbClr');
    srgbClr.setAttribute('val', color.slice(1));
    if (shadow.opacity !== undefined) {
      const alpha = spPr.ownerDocument.createElementNS(A_NS, 'a:alpha');
      alpha.setAttribute('val', String(Math.round(shadow.opacity * 100000)));
      srgbClr.appendChild(alpha);
    }
    outerShadow.appendChild(srgbClr);
  }
  effectList.appendChild(outerShadow);
}

function normalizeHexColor(color: string): string {
  const trimmed = color.trim();
  const raw = trimmed.startsWith('#') ? trimmed.slice(1) : trimmed;
  if (!/^[0-9a-fA-F]{6}$/.test(raw)) throw new Error(`Presentation: invalid color ${color}`);
  return `#${raw.toUpperCase()}`;
}

function normalizeOpacity(opacity: number): number {
  if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1)
    throw new Error(`Presentation: opacity must be between 0 and 1: ${opacity}`);
  return opacity;
}

function normalizeNonNegativeNumber(value: number, name: string): number {
  if (!Number.isFinite(value) || value < 0)
    throw new Error(`Presentation: ${name} must be a non-negative number: ${value}`);
  return value;
}

function normalizeFiniteNumber(value: number, name: string): number {
  if (!Number.isFinite(value))
    throw new Error(`Presentation: ${name} must be a finite number: ${value}`);
  return value;
}

function normalizeShadow(
  shadow: NonNullable<PptShapeStyle['shadow']>,
): NonNullable<PptShapeStyle['shadow']> {
  return {
    ...(shadow.color ? { color: normalizeHexColor(shadow.color) } : {}),
    ...(shadow.opacity !== undefined ? { opacity: normalizeOpacity(shadow.opacity) } : {}),
    ...(shadow.blur !== undefined
      ? { blur: normalizeNonNegativeNumber(shadow.blur, 'shadow blur') }
      : {}),
    ...(shadow.distance !== undefined
      ? { distance: normalizeNonNegativeNumber(shadow.distance, 'shadow distance') }
      : {}),
    ...(shadow.direction !== undefined
      ? { direction: normalizeFiniteNumber(shadow.direction, 'shadow direction') }
      : {}),
  };
}

function imageSourceBytes(source: PptImageSource): Uint8Array {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if ('bytes' in source)
    return source.bytes instanceof Uint8Array ? source.bytes : new Uint8Array(source.bytes);
  return Uint8Array.from(globalThis.atob(source.base64), (char) => char.charCodeAt(0));
}

function imageSourceInfo(
  source: PptImageSource,
  options: PptImageAddOptions = {},
): { bytes: Uint8Array; contentType: string; extension: string } {
  const bytes = imageSourceBytes(source);
  const contentType = options.mimeType ?? sourceMimeType(source) ?? sniffImageMimeType(bytes);
  return {
    bytes,
    contentType,
    extension: imageExtension(contentType, options.filename ?? sourceFilename(source)),
  };
}

function sourceMimeType(source: PptImageSource): string | undefined {
  if (source instanceof Uint8Array || source instanceof ArrayBuffer) return undefined;
  return source.mimeType;
}

function sourceFilename(source: PptImageSource): string | undefined {
  if (source instanceof Uint8Array || source instanceof ArrayBuffer) return undefined;
  return source.filename;
}

function sniffImageMimeType(bytes: Uint8Array): string {
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
    return 'image/png';
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return 'image/jpeg';
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'image/gif';
  return 'image/png';
}

function imageExtension(contentType: string, filename?: string): string {
  const filenameExt = filename?.match(/\.([A-Za-z0-9]+)$/)?.[1]?.toLowerCase();
  if (filenameExt) return filenameExt === 'jpg' ? 'jpeg' : filenameExt;
  if (contentType === 'image/jpeg') return 'jpeg';
  if (contentType === 'image/gif') return 'gif';
  if (contentType === 'image/svg+xml') return 'svg';
  return 'png';
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeInsertIndex(index: number | 'end', length: number): number {
  if (index === 'end') return length;
  return Math.max(0, Math.min(index, length));
}

function nextSlidePartNumber(slides: PptSlideModel[]): number {
  const used = new Set(
    slides
      .map((slide) => slide.path?.match(/ppt\/slides\/slide(\d+)\.xml$/)?.[1])
      .filter((value): value is string => Boolean(value))
      .map((value) => Number(value)),
  );
  let next = 1;
  while (used.has(next)) next++;
  return next;
}

function nextSlideId(slides: PptSlideModel[]): number {
  const max = slides.reduce((value, slide) => {
    const numeric = Number(slide.id);
    return Number.isFinite(numeric) ? Math.max(value, numeric) : value;
  }, 255);
  return max + 1;
}

function nextRelationshipId(presentation: PresentationCore): string {
  const rels = parseRelationships(requirePart(presentation, 'ppt/_rels/presentation.xml.rels'));
  const used = new Set(rels.keys());
  let index = 1;
  while (used.has(`rId${index}`)) index++;
  return `rId${index}`;
}

function addPresentationRelationship(
  presentation: PresentationCore,
  relId: string,
  target: string,
): void {
  addPresentationRelationshipOfType(presentation, {
    id: relId,
    type: `${OFFICE_REL_NS}/slide`,
    target,
  });
}

function addPresentationRelationshipOfType(
  presentation: PresentationCore,
  relationship: {
    id?: string;
    type: string;
    target: string;
    targetMode?: string;
  },
): string {
  const part = ensurePresentationRelationshipsPart(presentation);
  const relId = relationship.id ?? nextPresentationRelationshipIdFromPart(part);
  const rel = part.xmlDom!.createElementNS(PKG_REL_NS, 'Relationship');
  rel.setAttribute('Id', relId);
  rel.setAttribute('Type', relationship.type);
  rel.setAttribute('Target', relationship.target);
  if (relationship.targetMode) rel.setAttribute('TargetMode', relationship.targetMode);
  part.xmlDom!.documentElement.appendChild(rel);
  markDirty(part);
  return relId;
}

function ensurePresentationRelationshipsPart(presentation: PresentationCore): Part {
  const existing = getPart(presentation._package(), 'ppt/_rels/presentation.xml.rels');
  if (existing?.xmlDom) return existing;
  const part = createXmlPart({
    path: 'ppt/_rels/presentation.xml.rels',
    contentType: 'application/vnd.openxmlformats-package.relationships+xml',
    xmlDom: parseXmlDocument(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"/>`,
    ) as unknown as Document,
  });
  setPart(presentation._package(), part);
  return part;
}

function nextPresentationRelationshipIdFromPart(part: Part): string {
  const rels = parseRelationships(part);
  const used = new Set(rels.keys());
  let index = 1;
  while (used.has(`rId${index}`)) index++;
  return `rId${index}`;
}

function removePresentationRelationship(presentation: PresentationCore, relId: string): void {
  const part = requirePart(presentation, 'ppt/_rels/presentation.xml.rels');
  const rels = Array.from(
    part.xmlDom!.getElementsByTagNameNS(PKG_REL_NS, 'Relationship'),
  ) as Element[];
  const rel = rels.find((item) => item.getAttribute('Id') === relId);
  if (rel?.parentNode) {
    rel.parentNode.removeChild(rel);
    markDirty(part);
  }
}

function addSlideImageRelationship(
  presentation: PresentationCore,
  slideId: string,
  mediaPath: string,
): string {
  const slide = presentation._getSlideInternal(slideId);
  if (!slide?.model.path) throw new Error(`Presentation: slide not found: ${slideId}`);
  return addSlideRelationship(presentation, slide, {
    type: `${OFFICE_REL_NS}/image`,
    target: relativeMediaTarget(mediaPath),
  });
}

function addSlideRelationship(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
  relationship: {
    type: string;
    target: string;
    targetMode?: string;
  },
): string {
  if (!slide.model.path) throw new Error(`Presentation: slide not found: ${slide.model.id}`);
  const part = ensureSlideRelationshipsPart(presentation, slide.model.path);
  const rels = parseRelationships(part);
  const used = new Set(rels.keys());
  let index = 1;
  while (used.has(`rId${index}`)) index++;
  const relId = `rId${index}`;
  const rel = part.xmlDom!.createElementNS(PKG_REL_NS, 'Relationship');
  rel.setAttribute('Id', relId);
  rel.setAttribute('Type', relationship.type);
  rel.setAttribute('Target', relationship.target);
  if (relationship.targetMode) rel.setAttribute('TargetMode', relationship.targetMode);
  part.xmlDom!.documentElement.appendChild(rel);
  markDirty(part);
  slide.relationships.set(relId, {
    id: relId,
    type: relationship.type,
    target: relationship.target,
    targetMode: relationship.targetMode,
  });
  return relId;
}

function removeSlideRelationship(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
  relId: string,
): void {
  if (!slide.model.path) return;
  const part = getPart(presentation._package(), relsPathForPart(slide.model.path));
  if (!part?.xmlDom) return;
  const rels = Array.from(
    part.xmlDom.getElementsByTagNameNS(PKG_REL_NS, 'Relationship'),
  ) as Element[];
  const rel = rels.find((item) => item.getAttribute('Id') === relId);
  if (rel?.parentNode) {
    rel.parentNode.removeChild(rel);
    markDirty(part);
  }
  slide.relationships.delete(relId);
}

function removeSlideRelationshipIfUnreferenced(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
  relId: string,
  except?: Element,
): void {
  const refs = slide.part.xmlDom ? allDescendantsNs(slide.part.xmlDom, A_NS, 'hlinkClick') : [];
  const stillUsed = refs.some((item) => {
    if (item === except) return false;
    return (item.getAttributeNS(R_NS, 'id') ?? item.getAttribute('r:id')) === relId;
  });
  if (!stillUsed) removeSlideRelationship(presentation, slide, relId);
}

function getNotesSlidePart(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
): Part | undefined {
  const rel = Array.from(slide.relationships.values()).find(
    (item) => item.type === `${OFFICE_REL_NS}/notesSlide`,
  );
  return rel ? getPart(presentation._package(), resolveSlideTarget(slide, rel.target)) : undefined;
}

function ensureNotesSlidePart(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
): Part {
  const existing = getNotesSlidePart(presentation, slide);
  if (existing?.xmlDom) return existing;
  const path = nextNumberedPartPath(presentation, 'ppt/notesSlides/notesSlide', '.xml');
  const part = createXmlPart({
    path,
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    xmlDom: parseXmlDocument(notesSlideXml('')) as unknown as Document,
  });
  setPart(presentation._package(), part);
  addSlideRelationship(presentation, slide, {
    type: `${OFFICE_REL_NS}/notesSlide`,
    target: relativeSlideTarget(slide, path),
  });
  deletePart(presentation._package(), '[Content_Types].xml');
  return part;
}

function setNotesText(doc: Document, text: string): void {
  const textNodes = allDescendantsNs(doc, A_NS, 't');
  if (textNodes.length === 0) {
    const nextRoot = parseXmlDocument(notesSlideXml(text)).documentElement;
    if (!nextRoot) throw new Error('Presentation: failed to create notes XML');
    doc.replaceChild(nextRoot as unknown as Node, doc.documentElement as unknown as Node);
    return;
  }
  textNodes[0]!.textContent = text;
  for (const extra of textNodes.slice(1)) extra.textContent = '';
}

function getSlideCommentsPart(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
): Part | undefined {
  const rel = Array.from(slide.relationships.values()).find(
    (item) => item.type === `${OFFICE_REL_NS}/comments`,
  );
  return rel ? getPart(presentation._package(), resolveSlideTarget(slide, rel.target)) : undefined;
}

function ensureSlideCommentsPart(
  presentation: PresentationCore,
  slide: PptReadResult['slides'][number],
): Part {
  const existing = getSlideCommentsPart(presentation, slide);
  if (existing?.xmlDom) return existing;
  const path = nextNumberedPartPath(presentation, 'ppt/comments/comment', '.xml');
  const part = createXmlPart({
    path,
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml',
    xmlDom: parseXmlDocument(`<p:cmLst xmlns:p="${P_NS}"/>`) as unknown as Document,
  });
  setPart(presentation._package(), part);
  addSlideRelationship(presentation, slide, {
    type: `${OFFICE_REL_NS}/comments`,
    target: relativeSlideTarget(slide, path),
  });
  deletePart(presentation._package(), '[Content_Types].xml');
  return part;
}

function appendSlideComment(
  doc: Document,
  options: PptCommentAddOptions & { authorId: string },
): PptCommentModel {
  const root = doc.documentElement;
  const id = String(nextCommentIndex(doc));
  const cm = doc.createElementNS(P_NS, 'p:cm');
  cm.setAttribute('authorId', options.authorId);
  cm.setAttribute('idx', id);
  cm.setAttribute('dt', options.date ?? new Date().toISOString());
  const pos = doc.createElementNS(P_NS, 'p:pos');
  pos.setAttribute('x', String(options.x ?? 0));
  pos.setAttribute('y', String(options.y ?? 0));
  const text = doc.createElementNS(P_NS, 'p:text');
  text.textContent = options.text;
  cm.appendChild(pos);
  cm.appendChild(text);
  root.appendChild(cm);
  return {
    id,
    authorId: options.authorId,
    ...(options.author ? { author: options.author } : {}),
    ...(options.initials ? { initials: options.initials } : {}),
    ...(cm.getAttribute('dt') ? { date: cm.getAttribute('dt')! } : {}),
    text: options.text,
    x: options.x ?? 0,
    y: options.y ?? 0,
  };
}

function removeSlideComment(doc: Document, commentId: string): boolean {
  const comment = allDescendantsNs(doc, P_NS, 'cm').find(
    (item) => item.getAttribute('idx') === commentId,
  );
  if (!comment?.parentNode) return false;
  comment.parentNode.removeChild(comment);
  return true;
}

function ensureCommentAuthor(
  presentation: PresentationCore,
  author: { name: string; initials: string },
): { id: string } {
  const part = ensureCommentAuthorsPart(presentation);
  const existing = allDescendantsNs(part.xmlDom!, P_NS, 'cmAuthor').find(
    (item) =>
      item.getAttribute('name') === author.name &&
      item.getAttribute('initials') === author.initials,
  );
  if (existing?.getAttribute('id')) return { id: existing.getAttribute('id')! };
  const id = String(nextCommentAuthorId(part.xmlDom!));
  const node = part.xmlDom!.createElementNS(P_NS, 'p:cmAuthor');
  node.setAttribute('id', id);
  node.setAttribute('name', author.name);
  node.setAttribute('initials', author.initials);
  node.setAttribute('lastIdx', '0');
  node.setAttribute('clrIdx', id);
  part.xmlDom!.documentElement.appendChild(node);
  markDirty(part);
  return { id };
}

function ensureCommentAuthorsPart(presentation: PresentationCore): Part {
  const presentationPart = requirePart(presentation, 'ppt/presentation.xml');
  const relsPart = ensurePresentationRelationshipsPart(presentation);
  const rels = parseRelationships(relsPart);
  const existingRel = Array.from(rels.values()).find(
    (item) => item.type === `${OFFICE_REL_NS}/commentAuthors`,
  );
  const path = existingRel
    ? resolveRelationshipTarget(presentationPart.path, existingRel.target)
    : 'ppt/commentAuthors.xml';
  const existing = getPart(presentation._package(), path);
  if (existing?.xmlDom) return existing;
  const part = createXmlPart({
    path,
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml',
    xmlDom: parseXmlDocument(`<p:cmAuthorLst xmlns:p="${P_NS}"/>`) as unknown as Document,
  });
  setPart(presentation._package(), part);
  if (!existingRel) {
    addPresentationRelationshipOfType(presentation, {
      type: `${OFFICE_REL_NS}/commentAuthors`,
      target: 'commentAuthors.xml',
    });
  }
  deletePart(presentation._package(), '[Content_Types].xml');
  return part;
}

function nextCommentIndex(doc: Document): number {
  return (
    allDescendantsNs(doc, P_NS, 'cm').reduce((max, comment) => {
      const value = Number(comment.getAttribute('idx'));
      return Number.isFinite(value) ? Math.max(max, value) : max;
    }, 0) + 1
  );
}

function nextCommentAuthorId(doc: Document): number {
  return (
    allDescendantsNs(doc, P_NS, 'cmAuthor').reduce((max, author) => {
      const value = Number(author.getAttribute('id'));
      return Number.isFinite(value) ? Math.max(max, value) : max;
    }, -1) + 1
  );
}

function notesSlideXml(text: string): string {
  return `<p:notes xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXmlText(text)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
}

function nextNumberedPartPath(
  presentation: PresentationCore,
  prefix: string,
  suffix: string,
): string {
  let index = 1;
  let path = `${prefix}${index}${suffix}`;
  while (getPart(presentation._package(), path)) {
    index++;
    path = `${prefix}${index}${suffix}`;
  }
  return path;
}

function resolveSlideTarget(slide: PptReadResult['slides'][number], target: string): string {
  return resolveRelationshipTarget(slide.part.path, target);
}

function relativeSlideTarget(slide: PptReadResult['slides'][number], targetPath: string): string {
  if (targetPath.startsWith('ppt/notesSlides/')) {
    return `../notesSlides/${targetPath.slice('ppt/notesSlides/'.length)}`;
  }
  if (targetPath.startsWith('ppt/comments/')) {
    return `../comments/${targetPath.slice('ppt/comments/'.length)}`;
  }
  return `/${targetPath}`;
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function ensureSlideRelationshipsPart(presentation: PresentationCore, slidePath: string): Part {
  const relsPath = relsPathForPart(slidePath);
  const existing = presentation._package().parts.get(relsPath);
  if (existing?.xmlDom) return existing;
  const part = createXmlPart({
    path: relsPath,
    contentType: 'application/vnd.openxmlformats-package.relationships+xml',
    xmlDom: parseXmlDocument(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"/>`,
    ) as unknown as Document,
  });
  setPart(presentation._package(), part);
  return part;
}

function relativeMediaTarget(mediaPath: string): string {
  if (!mediaPath.startsWith('ppt/media/')) return `/${mediaPath}`;
  return `../media/${mediaPath.slice('ppt/media/'.length)}`;
}

function getSlideIdList(presentation: PresentationCore): Element {
  const part = requirePart(presentation, 'ppt/presentation.xml');
  const existing = allDescendantsNs(part.xmlDom!, P_NS, 'sldIdLst')[0];
  if (existing) return existing;
  const root = part.xmlDom!.documentElement;
  const sldIdLst = part.xmlDom!.createElementNS(P_NS, 'p:sldIdLst');
  root.appendChild(sldIdLst);
  markDirty(part);
  return sldIdLst;
}

function addSlideId(
  presentation: PresentationCore,
  index: number,
  slideId: string,
  relId: string,
): void {
  const part = requirePart(presentation, 'ppt/presentation.xml');
  const sldIdLst = getSlideIdList(presentation);
  const sldId = part.xmlDom!.createElementNS(P_NS, 'p:sldId');
  sldId.setAttribute('id', slideId);
  sldId.setAttributeNS(R_NS, 'r:id', relId);
  const nodes = allDescendantsNs(sldIdLst, P_NS, 'sldId');
  sldIdLst.insertBefore(sldId, nodes[index] ?? null);
  markDirty(part);
}

function removeSlideId(presentation: PresentationCore, relId: string): void {
  const part = requirePart(presentation, 'ppt/presentation.xml');
  const sldIdLst = getSlideIdList(presentation);
  const nodes = allDescendantsNs(sldIdLst, P_NS, 'sldId');
  const target = nodes.find(
    (node) => (node.getAttributeNS(R_NS, 'id') ?? node.getAttribute('r:id')) === relId,
  );
  if (target) {
    sldIdLst.removeChild(target);
    markDirty(part);
  }
}

function requirePart(presentation: PresentationCore, path: string): Part {
  const part = presentation._package().parts.get(path);
  if (!part?.xmlDom) throw new Error(`Presentation: missing XML part ${path}`);
  return part;
}

function defaultTitleGeometry(size: PptSlideSize | undefined): PptShapeGeometry {
  const width = size?.width ?? 9144000;
  return {
    left: Math.round(width * 0.08),
    top: 600000,
    width: Math.round(width * 0.84),
    height: 800000,
  };
}

function defaultShapeGeometry(): PptShapeGeometry {
  return {
    left: 600000,
    top: 900000,
    width: 3000000,
    height: 700000,
  };
}

function defaultTableGeometry(rowCount: number): PptShapeGeometry {
  return {
    left: 600000,
    top: 1200000,
    width: 4200000,
    height: Math.max(700000, rowCount * 380000),
  };
}

function defaultImageGeometry(): PptShapeGeometry {
  return {
    left: 600000,
    top: 900000,
    width: 2400000,
    height: 1600000,
  };
}

function defaultPresetShapeGeometry(preset: PptPresetShape): PptShapeGeometry {
  if (preset === 'line') {
    return {
      left: 600000,
      top: 900000,
      width: 2400000,
      height: 0,
    };
  }
  return {
    left: 600000,
    top: 900000,
    width: 1800000,
    height: 900000,
  };
}

function makeTextShapeModel(
  slide: PptSlideModel,
  text: string,
  options: PptShapeAddOptions,
): PptShapeModel {
  const id = nextShapeId(slide);
  return {
    id,
    name: options.name ?? `TextBox ${id}`,
    kind: 'textBox',
    geometry: { ...defaultShapeGeometry(), ...options.geometry },
    text: { text },
  };
}

function makePresetShapeModel(
  slide: PptSlideModel,
  preset: PptPresetShape,
  options: PptShapeAddOptions,
): PptShapeModel {
  const id = nextShapeId(slide);
  return {
    id,
    name: options.name ?? `${preset} ${id}`,
    kind: preset === 'line' ? 'line' : 'geometricShape',
    geometry: { ...defaultPresetShapeGeometry(preset), ...options.geometry },
    raw: {
      hostType: 'sp',
    },
  };
}

function makeImageShapeModel(
  slide: PptSlideModel,
  relationshipId: string,
  targetPath: string,
  contentType: string,
  options: PptImageAddOptions,
): PptShapeModel {
  const id = nextShapeId(slide);
  return {
    id,
    name: options.name ?? `Picture ${id}`,
    kind: 'picture',
    geometry: { ...defaultImageGeometry(), ...options.geometry },
    image: {
      relationshipId,
      targetPath,
      contentType,
      filename: targetPath.slice(targetPath.lastIndexOf('/') + 1),
    },
  };
}

function makeTableShapeModel(
  slide: PptSlideModel,
  rowCount: number,
  columnCount: number,
  options: PptTableAddOptions,
): PptShapeModel {
  const id = nextShapeId(slide);
  const values = normalizeTableValues(rowCount, columnCount, options.values);
  return {
    id,
    name: options.name ?? `Table ${id}`,
    kind: 'table',
    geometry: { ...defaultTableGeometry(rowCount), ...options.geometry },
    table: {
      rowCount,
      columnCount,
      values,
    },
  };
}

function normalizeTableValues(
  rowCount: number,
  columnCount: number,
  values?: string[][],
): string[][] {
  if (values && (values.length !== rowCount || values.some((row) => row.length !== columnCount))) {
    throw new Error(`Presentation: table values must match ${rowCount}x${columnCount}`);
  }
  return Array.from({ length: rowCount }, (_, row) =>
    Array.from({ length: columnCount }, (_, column) => values?.[row]?.[column] ?? ''),
  );
}

function nextShapeId(slide: PptSlideModel): string {
  const max = slide.shapes.reduce((value, shape) => {
    const numeric = Number(shape.id);
    return Number.isFinite(numeric) ? Math.max(value, numeric) : value;
  }, 1);
  return String(max + 1);
}

function nextMediaPath(presentation: PresentationCore, extension: string): string {
  let index = 1;
  let path = `ppt/media/image${index}.${extension}`;
  while (presentation._package().parts.has(path)) {
    index++;
    path = `ppt/media/image${index}.${extension}`;
  }
  return path;
}

function applyGeometryToElement(shape: Element, geometry: PptShapeGeometry): void {
  const xfrm = findOrCreateTransform(shape);
  let off = firstChildNs(xfrm, A_NS, 'off');
  if (!off) {
    off = xfrm.ownerDocument.createElementNS(A_NS, 'a:off');
    xfrm.appendChild(off);
  }
  let ext = firstChildNs(xfrm, A_NS, 'ext');
  if (!ext) {
    ext = xfrm.ownerDocument.createElementNS(A_NS, 'a:ext');
    xfrm.appendChild(ext);
  }
  setAttrIfDefined(off, 'x', geometry.left);
  setAttrIfDefined(off, 'y', geometry.top);
  setAttrIfDefined(ext, 'cx', geometry.width);
  setAttrIfDefined(ext, 'cy', geometry.height);
  setAttrIfDefined(xfrm, 'rot', degreesToOoxmlAngle(geometry.rotation));
}

function degreesToOoxmlAngle(rotation: number | undefined): number | undefined {
  return rotation === undefined ? undefined : Math.round(rotation * 60000);
}

function duplicatePictureMedia(
  presentation: PresentationCore,
  slideId: string,
  shape: Element,
  sourcePath: string,
): string {
  const sourcePart = getPart(presentation._package(), sourcePath);
  const bytes = sourcePart?.bytes ?? sourcePart?._originalBytes;
  if (!sourcePart || !bytes) throw new Error(`Presentation: image part not found: ${sourcePath}`);
  const contentType = sourcePart.contentType || 'image/png';
  const targetPath = nextMediaPath(
    presentation,
    imageExtension(contentType, sourcePath.slice(sourcePath.lastIndexOf('/') + 1)),
  );
  const relationshipId = addSlideImageRelationship(presentation, slideId, targetPath);
  setPart(
    presentation._package(),
    createBinaryPart({
      path: targetPath,
      contentType,
      bytes,
      originalBytes: bytes,
    }),
  );
  const blip = allDescendantsNs(shape, A_NS, 'blip')[0];
  if (!blip) throw new Error('Presentation: duplicated picture shape has no blip relationship');
  blip.setAttributeNS(R_NS, 'r:embed', relationshipId);
  deletePart(presentation._package(), '[Content_Types].xml');
  return targetPath;
}

function appendShapeXml(presentation: PresentationCore, slideId: string, xml: string): void {
  const internal = presentation._getSlideInternal(slideId);
  if (!internal) throw new Error(`Presentation: slide not found: ${slideId}`);
  const spTree = getSpTree(internal);
  const doc = parseXmlDocument(
    `<root xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">${xml}</root>`,
  ) as unknown as Document;
  const element = Array.from(doc.documentElement.childNodes).find((node) => node.nodeType === 1);
  if (!element) throw new Error('Presentation: generated shape XML is empty');
  spTree.appendChild(spTree.ownerDocument.importNode(element, true));
  markDirty(internal.part);
}

function getSpTree(slide: PptReadResult['slides'][number]): Element {
  const spTree = allDescendantsNs(slide.part.xmlDom!, P_NS, 'spTree')[0];
  if (!spTree) throw new Error(`Presentation: slide ${slide.model.id} has no shape tree`);
  return spTree;
}

function firstShapeTreeElement(parent: Node): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI !== P_NS) continue;
    if (
      element.localName === 'sp' ||
      element.localName === 'pic' ||
      element.localName === 'graphicFrame' ||
      element.localName === 'grpSp'
    ) {
      return element;
    }
  }
  return null;
}

function slideXmlFromModel(slide: PptSlideModel): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
    `<p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    slide.shapes.map((shape, index) => simpleTextShapeXml(shape, index)).join('') +
    `</p:spTree></p:cSld>` +
    `</p:sld>`
  );
}

function simpleTextShapeXml(shape: PptShapeModel, index: number): string {
  const geometry = shape.geometry ?? {};
  return (
    `<p:sp>` +
    `<p:nvSpPr><p:cNvPr id="${xmlAttr(shape.id || String(index + 2))}" name="${xmlAttr(shape.name || `Shape ${index + 1}`)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${geometry.left ?? 0}" y="${geometry.top ?? 0}"/><a:ext cx="${geometry.width ?? 3000000}" cy="${geometry.height ?? 600000}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlText(shape.text?.text ?? '')}</a:t></a:r></a:p></p:txBody>` +
    `</p:sp>`
  );
}

function tableShapeXml(shape: PptShapeModel, index: number): string {
  const geometry = shape.geometry ?? {};
  const rows = shape.table?.values ?? [];
  const columnCount =
    shape.table?.columnCount ?? rows.reduce((max, row) => Math.max(max, row.length), 0);
  const gridCols = Array.from({ length: columnCount }, () => `<a:gridCol w="1000000"/>`).join('');
  const tableRows = rows
    .map(
      (row) =>
        `<a:tr h="370840">` +
        Array.from({ length: columnCount }, (_, column) => tableCellXml(row[column] ?? '')).join(
          '',
        ) +
        `</a:tr>`,
    )
    .join('');
  return (
    `<p:graphicFrame>` +
    `<p:nvGraphicFramePr><p:cNvPr id="${xmlAttr(shape.id || String(index + 2))}" name="${xmlAttr(shape.name || `Table ${index + 1}`)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${geometry.left ?? 0}" y="${geometry.top ?? 0}"/><a:ext cx="${geometry.width ?? 3000000}" cy="${geometry.height ?? 900000}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
    `<a:tbl><a:tblPr/><a:tblGrid>${gridCols}</a:tblGrid>${tableRows}</a:tbl>` +
    `</a:graphicData></a:graphic>` +
    `</p:graphicFrame>`
  );
}

function pictureShapeXml(shape: PptShapeModel, index: number): string {
  const geometry = shape.geometry ?? {};
  const relId = shape.image?.relationshipId;
  if (!relId) throw new Error('Presentation: picture shape needs relationshipId');
  return (
    `<p:pic>` +
    `<p:nvPicPr><p:cNvPr id="${xmlAttr(shape.id || String(index + 2))}" name="${xmlAttr(shape.name || `Picture ${index + 1}`)}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${xmlAttr(relId)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${geometry.left ?? 0}" y="${geometry.top ?? 0}"/><a:ext cx="${geometry.width ?? 2400000}" cy="${geometry.height ?? 1600000}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `</p:pic>`
  );
}

function presetShapeXml(shape: PptShapeModel, preset: PptPresetShape, index: number): string {
  const geometry = shape.geometry ?? {};
  const isLine = preset === 'line';
  const shapeProperties = isLine
    ? `<p:spPr><a:xfrm><a:off x="${geometry.left ?? 0}" y="${geometry.top ?? 0}"/><a:ext cx="${geometry.width ?? 2400000}" cy="${geometry.height ?? 0}"/></a:xfrm><a:ln w="12700"/></p:spPr>`
    : `<p:spPr><a:xfrm><a:off x="${geometry.left ?? 0}" y="${geometry.top ?? 0}"/><a:ext cx="${geometry.width ?? 1800000}" cy="${geometry.height ?? 900000}"/></a:xfrm><a:prstGeom prst="${xmlAttr(preset)}"><a:avLst/></a:prstGeom></p:spPr>`;
  return (
    `<p:sp>` +
    `<p:nvSpPr><p:cNvPr id="${xmlAttr(shape.id || String(index + 2))}" name="${xmlAttr(shape.name || `${preset} ${index + 1}`)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    shapeProperties +
    `</p:sp>`
  );
}

function tableCellXml(text: string): string {
  return (
    `<a:tc>` +
    `<a:txBody><a:bodyPr/><a:p><a:r><a:t>${xmlText(text)}</a:t></a:r></a:p></a:txBody>` +
    `<a:tcPr/>` +
    `</a:tc>`
  );
}

function xmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function xmlAttr(value: string): string {
  return xmlText(value).replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function cloneModel<T>(model: T): T {
  return JSON.parse(JSON.stringify(model)) as T;
}
