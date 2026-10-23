import {
  PresentationCore,
  PresentationThemeCollectionCore,
  ShapeChartApiCore,
  ShapeCollectionCore,
  ShapeCore,
  ShapeGroupApiCore,
  ShapeImageApiCore,
  ShapeTextApiCore,
  SlideCommentsApiCore,
  SlideCollectionCore,
  SlideCore,
  SlideNotesApiCore,
  TableApiCore,
  TableCellCore,
  type PptImageAddOptions,
  type PptPresetShape,
  type PptShapeAddOptions,
  type PptShapeDuplicateOptions,
  type PptSlideAddOptions,
  type PptSlideCopyOptions,
  type PptTableAddOptions,
  type PresentationCoreLoadOptions,
} from './presentation-core.js';
import type {
  PresentationLoadOptions,
  PresentationRuntimeMode,
  PptChartData,
  PptChartModel,
  PptCommentAddOptions,
  PptCommentModel,
  PptExportedFont,
  PptExportedImage,
  PptFindTextOptions,
  PptHyperlinkModel,
  PptHyperlinkWriteOptions,
  PptImageCropModel,
  PptImageInfo,
  PptImageSource,
  PptImageStyle,
  PptParagraphAlignment,
  PptPlaceholderModel,
  PptPresentationModel,
  PptPresentationProperties,
  PptShapeGeometry,
  PptShapeKind,
  PptShapeModel,
  PptShapeStyle,
  PptSlideLayoutModel,
  PptSlideModel,
  PptSlideProperties,
  PptSlideTimingModel,
  PptSlideTransitionModel,
  PptTableCellMerge,
  PptTableCellStyle,
  PptThemeModel,
  PptTextBodyProperties,
  PptTextMatch,
  PptTextParagraph,
  PptTextParagraphWrite,
  PptTextReplaceOptions,
  PptTextReplaceResult,
  PptTextStyle,
  PptTextRun,
  PptTextWriteOptions,
} from './model.js';
import type { PptMutation } from './mutations.js';

export interface PresentationModeDecision {
  mode: 'inline' | 'worker';
  source: 'explicit-inline' | 'explicit-worker' | 'auto/inline';
}

export class Presentation {
  readonly presentationId: string;
  readonly mode: 'inline' | 'worker';
  readonly modeSource: PresentationModeDecision['source'];
  readonly slides: SlideCollection;
  readonly themes: PresentationThemeCollection;
  private disposed = false;

  private constructor(
    private readonly core: PresentationCore,
    decision: PresentationModeDecision,
  ) {
    this.presentationId = core.presentationId;
    this.mode = decision.mode;
    this.modeSource = decision.source;
    this.slides = new SlideCollection(core.slides);
    this.themes = new PresentationThemeCollection(core.themes);
  }

  static async fromZip(
    bytes: Uint8Array | ArrayBuffer,
    opts: PresentationLoadOptions = {},
  ): Promise<Presentation> {
    const decision = resolveMode(opts.mode);
    const core = await PresentationCore.fromZip(bytes, toCoreOptions(opts));
    opts.onModeResolved?.(decision, core.presentationId);
    return new Presentation(core, decision);
  }

  static async fromModel(
    model: PptPresentationModel,
    opts: PresentationLoadOptions = {},
  ): Promise<Presentation> {
    const decision = resolveMode(opts.mode);
    const core = PresentationCore.fromModel(model, toCoreOptions(opts));
    opts.onModeResolved?.(decision, core.presentationId);
    return new Presentation(core, decision);
  }

  static async createBlank(opts: PresentationLoadOptions = {}): Promise<Presentation> {
    return Presentation.fromModel(
      {
        id: opts.presentationId,
        slides: [
          {
            id: '256',
            index: 0,
            shapes: [],
          },
        ],
      },
      opts,
    );
  }

  async toModel(): Promise<PptPresentationModel> {
    this.assertNotDisposed();
    return this.core.toModel();
  }

  async properties(): Promise<PptPresentationProperties> {
    this.assertNotDisposed();
    return this.core.properties();
  }

  async toZip(): Promise<Uint8Array> {
    this.assertNotDisposed();
    return this.core.toZip();
  }

  async getMutations(): Promise<PptMutation[]> {
    this.assertNotDisposed();
    return this.core.getMutations();
  }

  async clearMutations(): Promise<void> {
    this.assertNotDisposed();
    this.core.clearMutations();
  }

  async findText(query: string | RegExp, options?: PptFindTextOptions): Promise<PptTextMatch[]> {
    this.assertNotDisposed();
    return this.core.findText(query, options);
  }

  async replaceText(
    search: string | RegExp,
    replacement: string,
    options?: PptTextReplaceOptions,
  ): Promise<PptTextReplaceResult> {
    this.assertNotDisposed();
    return this.core.replaceText(search, replacement, options);
  }

  async exportImages(): Promise<PptExportedImage[]> {
    this.assertNotDisposed();
    return this.core.exportImages();
  }

  async exportFonts(): Promise<PptExportedFont[]> {
    this.assertNotDisposed();
    return this.core.exportFonts();
  }

  async executeCode<T = unknown>(): Promise<T> {
    throw new Error('Presentation.executeCode is planned but not implemented in PPT P0 runtime');
  }

  async dispose(): Promise<void> {
    this.disposed = true;
  }

  private assertNotDisposed(): void {
    if (this.disposed) throw new Error('Presentation has been disposed');
  }
}

export class PresentationThemeCollection {
  constructor(private readonly core: PresentationThemeCollectionCore) {}

  async items(): Promise<PptThemeModel[]> {
    return this.core.items();
  }

  async getCount(): Promise<number> {
    return this.core.getCount();
  }
}

export class SlideCollection {
  constructor(private readonly core: SlideCollectionCore) {}

  async items(): Promise<Slide[]> {
    return this.core.items().map((slide) => new Slide(slide));
  }

  async getCount(): Promise<number> {
    return this.core.getCount();
  }

  getItem(indexOrId: number | string): Slide {
    return new Slide(this.core.getItem(indexOrId));
  }

  async getItemOrNullObject(indexOrId: number | string): Promise<Slide | null> {
    const slide = this.core.getItemOrNullObject(indexOrId);
    return slide ? new Slide(slide) : null;
  }

  async add(options?: PptSlideAddOptions): Promise<Slide> {
    return new Slide(this.core.add(options));
  }

  async copy(indexOrId: number | string, options?: PptSlideCopyOptions): Promise<Slide> {
    return new Slide(this.core.copy(indexOrId, options));
  }

  async move(indexOrId: number | string, toIndex: number): Promise<void> {
    this.core.move(indexOrId, toIndex);
  }

  async deleteItem(indexOrId: number | string): Promise<void> {
    this.core.deleteItem(indexOrId);
  }
}

export class Slide {
  readonly shapes: ShapeCollection;
  readonly notes: SlideNotesApi;
  readonly comments: SlideCommentsApi;

  constructor(private readonly core: SlideCore) {
    this.shapes = new ShapeCollection(core.shapes);
    this.notes = new SlideNotesApi(core.notes);
    this.comments = new SlideCommentsApi(core.comments);
  }

  async id(): Promise<string> {
    return this.core.id();
  }

  async index(): Promise<number> {
    return this.core.index();
  }

  async name(): Promise<string | undefined> {
    return this.core.name();
  }

  async title(): Promise<string | undefined> {
    return this.core.title();
  }

  async layout(): Promise<PptSlideLayoutModel | null> {
    return this.core.layout();
  }

  async transition(): Promise<PptSlideTransitionModel | null> {
    return this.core.transition();
  }

  async timing(): Promise<PptSlideTimingModel | null> {
    return this.core.timing();
  }

  async properties(): Promise<PptSlideProperties> {
    return this.core.properties();
  }

  async setTitle(text: string, opts?: PptTextWriteOptions): Promise<void> {
    this.core.setTitle(text, opts);
  }

  async findText(query: string | RegExp, options?: PptFindTextOptions): Promise<PptTextMatch[]> {
    return this.core.findText(query, options);
  }

  async replaceText(
    search: string | RegExp,
    replacement: string,
    options?: PptTextReplaceOptions,
  ): Promise<PptTextReplaceResult> {
    return this.core.replaceText(search, replacement, options);
  }

  async duplicate(options?: PptSlideCopyOptions): Promise<Slide> {
    return new Slide(this.core.duplicate(options));
  }

  async moveTo(index: number): Promise<void> {
    this.core.moveTo(index);
  }

  async delete(): Promise<void> {
    this.core.delete();
  }

  async toModel(): Promise<PptSlideModel> {
    return this.core.toModel();
  }
}

export class SlideNotesApi {
  constructor(private readonly core: SlideNotesApiCore) {}

  async hasText(): Promise<boolean> {
    return this.core.hasText();
  }

  async getText(): Promise<string> {
    return this.core.getText();
  }

  async setText(text: string): Promise<void> {
    this.core.setText(text);
  }

  async clear(): Promise<void> {
    this.core.clear();
  }
}

export class SlideCommentsApi {
  constructor(private readonly core: SlideCommentsApiCore) {}

  async items(): Promise<PptCommentModel[]> {
    return this.core.items();
  }

  async getCount(): Promise<number> {
    return this.core.getCount();
  }

  async add(options: PptCommentAddOptions): Promise<PptCommentModel> {
    return this.core.add(options);
  }

  async delete(commentId: string): Promise<void> {
    this.core.delete(commentId);
  }

  async clear(): Promise<void> {
    this.core.clear();
  }
}

export class ShapeCollection {
  constructor(private readonly core: ShapeCollectionCore) {}

  async items(): Promise<Shape[]> {
    return this.core.items().map((shape) => new Shape(shape));
  }

  async getCount(): Promise<number> {
    return this.core.getCount();
  }

  getItem(idOrName: string): Shape {
    return new Shape(this.core.getItem(idOrName));
  }

  getItemAt(index: number): Shape {
    return new Shape(this.core.getItemAt(index));
  }

  async getItemOrNullObject(idOrName: string): Promise<Shape | null> {
    const shape = this.core.getItemOrNullObject(idOrName);
    return shape ? new Shape(shape) : null;
  }

  async addTextBox(text: string, options?: PptShapeAddOptions): Promise<Shape> {
    return new Shape(this.core.addTextBox(text, options));
  }

  async addTable(
    rowCount: number,
    columnCount: number,
    options?: PptTableAddOptions,
  ): Promise<Shape> {
    return new Shape(this.core.addTable(rowCount, columnCount, options));
  }

  async addImage(source: PptImageSource, options?: PptImageAddOptions): Promise<Shape> {
    return new Shape(this.core.addImage(source, options));
  }

  async addShape(preset: PptPresetShape, options?: PptShapeAddOptions): Promise<Shape> {
    return new Shape(this.core.addShape(preset, options));
  }

  async deleteItem(idOrName: string): Promise<void> {
    this.core.deleteItem(idOrName);
  }
}

export class Shape {
  readonly chart: ShapeChartApi;
  readonly text: ShapeTextApi;
  readonly image: ShapeImageApi;
  readonly table: TableApi;
  readonly group: ShapeGroupApi;

  constructor(private readonly core: ShapeCore) {
    this.chart = new ShapeChartApi(core.chart);
    this.text = new ShapeTextApi(core.text);
    this.image = new ShapeImageApi(core.image);
    this.table = new TableApi(core.table);
    this.group = new ShapeGroupApi(core.group);
  }

  async id(): Promise<string> {
    return this.core.id();
  }

  async name(): Promise<string | undefined> {
    return this.core.name();
  }

  async setName(name: string): Promise<void> {
    this.core.setName(name);
  }

  async kind(): Promise<PptShapeKind> {
    return this.core.kind();
  }

  async geometry(): Promise<PptShapeGeometry> {
    return this.core.geometry();
  }

  async effectiveGeometry(): Promise<PptShapeGeometry> {
    return this.core.effectiveGeometry();
  }

  async placeholder(): Promise<PptPlaceholderModel | null> {
    return this.core.placeholder();
  }

  async effectivePlaceholder(): Promise<PptPlaceholderModel | null> {
    return this.core.effectivePlaceholder();
  }

  async style(): Promise<PptShapeStyle> {
    return this.core.style();
  }

  async effectiveStyle(): Promise<PptShapeStyle> {
    return this.core.effectiveStyle();
  }

  async hyperlink(): Promise<PptHyperlinkModel | null> {
    return this.core.hyperlink();
  }

  async setHyperlink(hyperlink: PptHyperlinkWriteOptions): Promise<void> {
    this.core.setHyperlink(hyperlink);
  }

  async clearHyperlink(): Promise<void> {
    this.core.clearHyperlink();
  }

  async setGeometry(geometry: Partial<PptShapeGeometry>): Promise<void> {
    this.core.setGeometry(geometry);
  }

  async setFillColor(color: string): Promise<void> {
    this.core.setFillColor(color);
  }

  async setLineStyle(style: Pick<PptShapeStyle, 'lineColor' | 'lineWidth'>): Promise<void> {
    this.core.setLineStyle(style);
  }

  async setStyle(style: Partial<PptShapeStyle>): Promise<void> {
    this.core.setStyle(style);
  }

  async delete(): Promise<void> {
    this.core.delete();
  }

  async duplicate(options?: PptShapeDuplicateOptions): Promise<Shape> {
    return new Shape(this.core.duplicate(options));
  }

  async bringToFront(): Promise<void> {
    this.core.bringToFront();
  }

  async sendToBack(): Promise<void> {
    this.core.sendToBack();
  }

  async toModel(): Promise<PptShapeModel> {
    return this.core.toModel();
  }
}

export class ShapeChartApi {
  constructor(private readonly core: ShapeChartApiCore) {}

  async isChart(): Promise<boolean> {
    return this.core.isChart();
  }

  async getInfo(): Promise<PptChartModel | null> {
    return this.core.getInfo();
  }

  async getData(): Promise<PptChartData | null> {
    return this.core.getData();
  }

  async setData(data: PptChartData): Promise<void> {
    this.core.setData(data);
  }
}

export class ShapeGroupApi {
  constructor(private readonly core: ShapeGroupApiCore) {}

  async isGroup(): Promise<boolean> {
    return this.core.isGroup();
  }

  async items(): Promise<Shape[]> {
    return this.core.items().map((shape) => new Shape(shape));
  }

  async getCount(): Promise<number> {
    return this.core.getCount();
  }

  getItem(idOrName: string): Shape {
    return new Shape(this.core.getItem(idOrName));
  }

  getItemAt(index: number): Shape {
    return new Shape(this.core.getItemAt(index));
  }

  async getItemOrNullObject(idOrName: string): Promise<Shape | null> {
    const shape = this.core.getItemOrNullObject(idOrName);
    return shape ? new Shape(shape) : null;
  }
}

export class ShapeTextApi {
  constructor(private readonly core: ShapeTextApiCore) {}

  async hasText(): Promise<boolean> {
    return this.core.hasText();
  }

  async getText(): Promise<string> {
    return this.core.getText();
  }

  async style(): Promise<PptTextStyle> {
    return this.core.style();
  }

  async effectiveStyle(): Promise<PptTextStyle> {
    return this.core.effectiveStyle();
  }

  async body(): Promise<PptTextBodyProperties> {
    return this.core.body();
  }

  async effectiveBody(): Promise<PptTextBodyProperties> {
    return this.core.effectiveBody();
  }

  async paragraphs(): Promise<PptTextParagraph[]> {
    return this.core.paragraphs();
  }

  async runs(): Promise<PptTextRun[]> {
    return this.core.runs();
  }

  async effectiveRuns(): Promise<PptTextRun[]> {
    return this.core.effectiveRuns();
  }

  async setParagraphs(paragraphs: PptTextParagraphWrite[]): Promise<void> {
    this.core.setParagraphs(paragraphs);
  }

  async setRunStyle(runIndex: number, style: Partial<PptTextStyle>): Promise<void> {
    this.core.setRunStyle(runIndex, style);
  }

  async setRunTextStyle(runIndex: number, style: Partial<PptTextStyle>): Promise<void> {
    this.core.setRunTextStyle(runIndex, style);
  }

  async setRunHyperlink(runIndex: number, hyperlink: PptHyperlinkWriteOptions): Promise<void> {
    this.core.setRunHyperlink(runIndex, hyperlink);
  }

  async clearRunHyperlink(runIndex: number): Promise<void> {
    this.core.clearRunHyperlink(runIndex);
  }

  async setText(text: string, opts?: PptTextWriteOptions): Promise<void> {
    this.core.setText(text, opts);
  }

  async setStyle(style: Partial<PptTextStyle>): Promise<void> {
    this.core.setStyle(style);
  }

  async setTextStyle(style: Partial<PptTextStyle>): Promise<void> {
    this.core.setTextStyle(style);
  }

  async setBodyProperties(body: Partial<PptTextBodyProperties>): Promise<void> {
    this.core.setBodyProperties(body);
  }

  async setBold(enabled: boolean): Promise<void> {
    this.core.setBold(enabled);
  }

  async setItalic(enabled: boolean): Promise<void> {
    this.core.setItalic(enabled);
  }

  async setAlignment(alignment: PptParagraphAlignment): Promise<void> {
    this.core.setAlignment(alignment);
  }

  async replace(
    search: string | RegExp,
    replacement: string,
    opts?: PptTextReplaceOptions,
  ): Promise<PptTextReplaceResult> {
    return this.core.replace(search, replacement, opts);
  }
}

export class ShapeImageApi {
  constructor(private readonly core: ShapeImageApiCore) {}

  async isImage(): Promise<boolean> {
    return this.core.isImage();
  }

  async getInfo(): Promise<PptImageInfo | null> {
    return this.core.getInfo();
  }

  async style(): Promise<PptImageStyle> {
    return this.core.style();
  }

  async setStyle(style: Partial<PptImageStyle>): Promise<void> {
    this.core.setStyle(style);
  }

  async setCrop(crop: PptImageCropModel): Promise<void> {
    this.core.setCrop(crop);
  }

  async setFillRect(fillRect: PptImageCropModel): Promise<void> {
    this.core.setFillRect(fillRect);
  }

  async setOpacity(opacity: number): Promise<void> {
    this.core.setOpacity(opacity);
  }

  async bytes(): Promise<Uint8Array> {
    return this.core.bytes();
  }

  async replace(source: PptImageSource): Promise<void> {
    this.core.replace(source);
  }
}

export class TableApi {
  constructor(private readonly core: TableApiCore) {}

  async isTable(): Promise<boolean> {
    return this.core.isTable();
  }

  async rowCount(): Promise<number> {
    return this.core.rowCount();
  }

  async columnCount(): Promise<number> {
    return this.core.columnCount();
  }

  async values(): Promise<string[][]> {
    return this.core.values();
  }

  async hasMergedCells(): Promise<boolean> {
    return this.core.hasMergedCells();
  }

  async mergedCells(): Promise<PptTableCellMerge[]> {
    return this.core.mergedCells();
  }

  async setValues(values: string[][]): Promise<void> {
    this.core.setValues(values);
  }

  async setCellText(row: number, column: number, text: string): Promise<void> {
    this.core.setCellText(row, column, text);
  }

  async addRow(index: number | 'end' = 'end', values?: string[]): Promise<void> {
    this.core.addRow(index, values);
  }

  async deleteRow(index: number): Promise<void> {
    this.core.deleteRow(index);
  }

  async addColumn(index: number | 'end' = 'end', values?: string[]): Promise<void> {
    this.core.addColumn(index, values);
  }

  async deleteColumn(index: number): Promise<void> {
    this.core.deleteColumn(index);
  }

  getCell(row: number, column: number): TableCell {
    return new TableCell(this.core.getCell(row, column));
  }
}

export class TableCell {
  constructor(private readonly core: TableCellCore) {}

  async text(): Promise<string> {
    return this.core.text();
  }

  async paragraphs(): Promise<PptTextParagraph[]> {
    return this.core.paragraphs();
  }

  async runs(): Promise<PptTextRun[]> {
    return this.core.runs();
  }

  async setParagraphs(paragraphs: PptTextParagraphWrite[]): Promise<void> {
    this.core.setParagraphs(paragraphs);
  }

  async setRunStyle(runIndex: number, style: Partial<PptTextStyle>): Promise<void> {
    this.core.setRunStyle(runIndex, style);
  }

  async setText(text: string): Promise<void> {
    this.core.setText(text);
  }

  async style(): Promise<PptTableCellStyle> {
    return this.core.style();
  }

  async setStyle(style: Partial<PptTableCellStyle>): Promise<void> {
    this.core.setStyle(style);
  }
}

function resolveMode(mode: PresentationRuntimeMode = 'auto'): PresentationModeDecision {
  if (mode === 'worker') {
    throw new Error('Presentation worker mode is planned but not implemented in PPT P0');
  }
  if (mode === 'inline') return { mode: 'inline', source: 'explicit-inline' };
  return { mode: 'inline', source: 'auto/inline' };
}

function toCoreOptions(opts: PresentationLoadOptions): PresentationCoreLoadOptions {
  return { presentationId: opts.presentationId };
}
