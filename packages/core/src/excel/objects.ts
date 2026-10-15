import type {
  ExcelChartAddOptions,
  ExcelChartModel,
  ExcelCommentModel,
  ExcelConditionalFormatModel,
  ExcelConditionalFormatRuleModel,
  ExcelDataValidationModel,
  ExcelDrawingAnchorModel,
  ExcelDrawingMarkerModel,
  ExcelDrawingModel,
  ExcelHyperlinkModel,
  ExcelImageModel,
  ExcelShapeModel,
  ExcelTableAutoFilterModel,
  ExcelTableColumnModel,
  ExcelTableModel,
  ExcelTableSortStateModel,
  ExcelTableStyleModel,
  ExcelThreadedCommentModel,
} from './model.js';
import { parseRangeAddress } from './address.js';

export class ExcelTable {
  readonly columns: ExcelTableColumnCollection;
  readonly style: ExcelTableStyle;

  constructor(
    private readonly model: ExcelTableModel,
    private readonly remove?: () => void,
  ) {
    this.columns = new ExcelTableColumnCollection(() => this.model.columns);
    this.style = new ExcelTableStyle(() => (this.model.style ??= {}));
  }

  id(): number | undefined {
    return this.model.id;
  }

  name(): string | undefined {
    return this.model.name;
  }

  displayName(): string | undefined {
    return this.model.displayName;
  }

  setName(name: string): void {
    this.model.name = name;
  }

  setDisplayName(displayName: string): void {
    this.model.displayName = displayName;
  }

  rangeAddress(): string | undefined {
    return this.model.ref;
  }

  setRange(address: string): void {
    parseRangeAddress(address);
    this.model.ref = address;
    if (this.model.autoFilterRef) this.setAutoFilterRange(address);
  }

  resize(address: string): void {
    this.setRange(address);
    const range = parseRangeAddress(address);
    const current = this.model.columns;
    if (range.columnCount < current.length) {
      current.splice(range.columnCount);
    }
    while (current.length < range.columnCount) {
      const id = nextTableColumnId(current);
      current.push({ id, name: `Column${id}` });
    }
  }

  autoFilterRangeAddress(): string | undefined {
    return this.model.autoFilterRef;
  }

  setAutoFilterRange(address: string | undefined): void {
    if (address !== undefined) parseRangeAddress(address);
    this.model.autoFilterRef = address;
    if (address === undefined) {
      delete this.model.autoFilter;
    } else {
      this.model.autoFilter ??= { ref: address, filterColumns: [] };
      this.model.autoFilter.ref = address;
    }
  }

  headerRowCount(): number | undefined {
    return this.model.headerRowCount;
  }

  setHeaderRowCount(count: number | undefined): void {
    assertOptionalNonNegativeInteger(count, 'ExcelTable.setHeaderRowCount');
    this.model.headerRowCount = count;
  }

  totalsRowCount(): number | undefined {
    return this.model.totalsRowCount;
  }

  setTotalsRowCount(count: number | undefined): void {
    assertOptionalNonNegativeInteger(count, 'ExcelTable.setTotalsRowCount');
    this.model.totalsRowCount = count;
  }

  totalsRowShown(): boolean | undefined {
    return this.model.totalsRowShown;
  }

  setTotalsRowShown(value: boolean | undefined): void {
    this.model.totalsRowShown = value;
  }

  autoFilter(): ExcelTableAutoFilterModel | undefined {
    return this.model.autoFilter ? cloneAutoFilter(this.model.autoFilter) : undefined;
  }

  setAutoFilter(model: ExcelTableAutoFilterModel | undefined): void {
    if (!model) {
      delete this.model.autoFilter;
      delete this.model.autoFilterRef;
      return;
    }
    const copy = cloneAutoFilter(model);
    if (copy.ref) parseRangeAddress(copy.ref);
    this.model.autoFilter = copy;
    this.model.autoFilterRef = copy.ref;
  }

  sortState(): ExcelTableSortStateModel | undefined {
    return this.model.autoFilter?.sortState ? cloneSortState(this.model.autoFilter.sortState) : undefined;
  }

  setSortState(sortState: ExcelTableSortStateModel | undefined): void {
    if (!sortState) {
      if (this.model.autoFilter) delete this.model.autoFilter.sortState;
      return;
    }
    this.model.autoFilter ??= { ref: this.model.autoFilterRef ?? this.model.ref, filterColumns: [] };
    this.model.autoFilter.sortState = cloneSortState(sortState);
  }

  toModel(): ExcelTableModel {
    return cloneTable(this.model);
  }

  delete(): void {
    this.remove?.();
  }
}

export class ExcelTableColumn {
  constructor(
    private readonly model: ExcelTableColumnModel,
    private readonly remove?: () => void,
  ) {}

  id(): number {
    return this.model.id;
  }

  name(): string {
    return this.model.name;
  }

  setName(name: string): void {
    this.model.name = name;
  }

  totalsRowLabel(): string | undefined {
    return this.model.totalsRowLabel;
  }

  setTotalsRowLabel(label: string | undefined): void {
    this.model.totalsRowLabel = label;
  }

  totalsRowFunction(): string | undefined {
    return this.model.totalsRowFunction;
  }

  setTotalsRowFunction(fn: string | undefined): void {
    this.model.totalsRowFunction = fn;
  }

  toModel(): ExcelTableColumnModel {
    return { ...this.model };
  }

  delete(): void {
    this.remove?.();
  }
}

export class ExcelTableColumnCollection {
  constructor(private readonly getModels: () => ExcelTableColumnModel[]) {}

  items(): ExcelTableColumn[] {
    return this.getModels().map((model) => this.wrap(model));
  }

  getItem(nameOrId: string | number): ExcelTableColumn {
    const item = this.getItemOrNullObject(nameOrId);
    if (!item) throw new Error(`ExcelTableColumnCollection: column not found: ${nameOrId}`);
    return item;
  }

  getItemOrNullObject(nameOrId: string | number): ExcelTableColumn | null {
    const model = this.getModels().find((column) => column.name === nameOrId || column.id === nameOrId);
    return model ? this.wrap(model) : null;
  }

  add(model: ExcelTableColumnModel): ExcelTableColumn {
    const models = this.getModels();
    if (models.some((column) => column.name === model.name || column.id === model.id)) {
      throw new Error(`ExcelTableColumnCollection: duplicate column: ${model.name || model.id}`);
    }
    const copy = { ...model, id: model.id ?? nextTableColumnId(models) };
    this.getModels().push(copy);
    return this.wrap(copy);
  }

  delete(nameOrId: string | number): void {
    const models = this.getModels();
    const index = models.findIndex((column) => column.name === nameOrId || column.id === nameOrId);
    if (index >= 0) models.splice(index, 1);
  }

  deleteItem(nameOrId: string | number): void {
    const models = this.getModels();
    const index = models.findIndex((column) => column.name === nameOrId || column.id === nameOrId);
    if (index === -1) throw new Error(`ExcelTableColumnCollection: column not found: ${nameOrId}`);
    models.splice(index, 1);
  }

  toModel(): ExcelTableColumnModel[] {
    return this.getModels().map((column) => ({ ...column }));
  }

  private wrap(model: ExcelTableColumnModel): ExcelTableColumn {
    return new ExcelTableColumn(model, () => {
      const models = this.getModels();
      const index = models.indexOf(model);
      if (index >= 0) models.splice(index, 1);
    });
  }
}

export class ExcelTableStyle {
  constructor(private readonly getModel: () => ExcelTableStyleModel) {}

  name(): string | undefined {
    return this.getModel().name;
  }

  setName(name: string | undefined): void {
    this.getModel().name = name;
  }

  showFirstColumn(): boolean | undefined {
    return this.getModel().showFirstColumn;
  }

  setShowFirstColumn(value: boolean | undefined): void {
    this.getModel().showFirstColumn = value;
  }

  showLastColumn(): boolean | undefined {
    return this.getModel().showLastColumn;
  }

  setShowLastColumn(value: boolean | undefined): void {
    this.getModel().showLastColumn = value;
  }

  showRowStripes(): boolean | undefined {
    return this.getModel().showRowStripes;
  }

  setShowRowStripes(value: boolean | undefined): void {
    this.getModel().showRowStripes = value;
  }

  showColumnStripes(): boolean | undefined {
    return this.getModel().showColumnStripes;
  }

  setShowColumnStripes(value: boolean | undefined): void {
    this.getModel().showColumnStripes = value;
  }

  toModel(): ExcelTableStyleModel {
    return { ...this.getModel() };
  }
}

export class ExcelTableCollection {
  constructor(private readonly getModels: () => ExcelTableModel[]) {}

  items(): ExcelTable[] {
    return this.getModels().map((model) => this.wrap(model));
  }

  getItem(nameOrDisplayName: string): ExcelTable {
    const item = this.getItemOrNullObject(nameOrDisplayName);
    if (!item) throw new Error(`ExcelTableCollection: table not found: ${nameOrDisplayName}`);
    return item;
  }

  getItemOrNullObject(nameOrDisplayName: string): ExcelTable | null {
    const model =
      this.getModels().find((table) => table.name === nameOrDisplayName || table.displayName === nameOrDisplayName) ??
      null;
    return model ? this.wrap(model) : null;
  }

  add(model: ExcelTableModel): ExcelTable {
    const copy = cloneTable(model);
    this.getModels().push(copy);
    return this.wrap(copy);
  }

  delete(nameOrDisplayName: string): void {
    const models = this.getModels();
    const index = models.findIndex((table) => table.name === nameOrDisplayName || table.displayName === nameOrDisplayName);
    if (index >= 0) models.splice(index, 1);
  }

  toModel(): ExcelTableModel[] {
    return this.getModels().map((model) => cloneTable(model));
  }

  private wrap(model: ExcelTableModel): ExcelTable {
    return new ExcelTable(model, () => {
      const models = this.getModels();
      const index = models.indexOf(model);
      if (index >= 0) models.splice(index, 1);
    });
  }
}

export class ExcelHyperlink {
  constructor(
    private readonly model: ExcelHyperlinkModel,
    private readonly remove?: () => void,
  ) {}

  ref(): string {
    return this.model.ref;
  }

  setRef(ref: string): void {
    this.model.ref = ref;
  }

  target(): string | undefined {
    return this.model.target;
  }

  targetMode(): string | undefined {
    return this.model.targetMode;
  }

  display(): string | undefined {
    return this.model.display;
  }

  setDisplay(display: string | undefined): void {
    this.model.display = display;
  }

  setLocation(location: string | undefined): void {
    this.model.location = location;
  }

  toModel(): ExcelHyperlinkModel {
    return { ...this.model };
  }

  delete(): void {
    this.remove?.();
  }
}

export class ExcelHyperlinkCollection {
  constructor(private readonly getModels: () => ExcelHyperlinkModel[]) {}

  items(): ExcelHyperlink[] {
    return this.getModels().map((model) => this.wrap(model));
  }

  add(model: ExcelHyperlinkModel): ExcelHyperlink {
    const copy = { ...model };
    this.getModels().push(copy);
    return this.wrap(copy);
  }

  getItem(ref: string): ExcelHyperlink {
    const item = this.getItemOrNullObject(ref);
    if (!item) throw new Error(`hyperlink not found: ${ref}`);
    return item;
  }

  getItemOrNullObject(ref: string): ExcelHyperlink | null {
    const model = this.getModels().find((item) => item.ref === ref);
    return model ? this.wrap(model) : null;
  }

  delete(ref: string): void {
    const models = this.getModels();
    for (let i = models.length - 1; i >= 0; i--) {
      if (models[i]?.ref === ref) models.splice(i, 1);
    }
  }

  toModel(): ExcelHyperlinkModel[] {
    return this.getModels().map((model) => ({ ...model }));
  }

  private wrap(model: ExcelHyperlinkModel): ExcelHyperlink {
    return new ExcelHyperlink(model, () => {
      const models = this.getModels();
      const index = models.indexOf(model);
      if (index >= 0) models.splice(index, 1);
    });
  }
}

export class ExcelDataValidation {
  constructor(
    private readonly model: ExcelDataValidationModel,
    private readonly remove?: () => void,
  ) {}

  rangeAddress(): string | undefined {
    return this.model.sqref;
  }

  setRange(address: string): void {
    this.model.sqref = address;
  }

  type(): string | undefined {
    return this.model.type;
  }

  setType(type: string | undefined): void {
    this.model.type = type;
  }

  operator(): string | undefined {
    return this.model.operator;
  }

  setOperator(operator: string | undefined): void {
    this.model.operator = operator;
  }

  allowBlank(): boolean | undefined {
    return this.model.allowBlank;
  }

  setAllowBlank(value: boolean | undefined): void {
    this.model.allowBlank = value;
  }

  showDropDown(): boolean | undefined {
    return this.model.showDropDown;
  }

  setShowDropDown(value: boolean | undefined): void {
    this.model.showDropDown = value;
  }

  showInputMessage(): boolean | undefined {
    return this.model.showInputMessage;
  }

  setShowInputMessage(value: boolean | undefined): void {
    this.model.showInputMessage = value;
  }

  showErrorMessage(): boolean | undefined {
    return this.model.showErrorMessage;
  }

  setShowErrorMessage(value: boolean | undefined): void {
    this.model.showErrorMessage = value;
  }

  errorStyle(): string | undefined {
    return this.model.errorStyle;
  }

  setErrorStyle(style: string | undefined): void {
    this.model.errorStyle = style;
  }

  errorTitle(): string | undefined {
    return this.model.errorTitle;
  }

  error(): string | undefined {
    return this.model.error;
  }

  setError(title: string | undefined, message: string | undefined): void {
    this.model.errorTitle = title;
    this.model.error = message;
  }

  promptTitle(): string | undefined {
    return this.model.promptTitle;
  }

  prompt(): string | undefined {
    return this.model.prompt;
  }

  setPrompt(title: string | undefined, message: string | undefined): void {
    this.model.promptTitle = title;
    this.model.prompt = message;
  }

  formula1(): string | undefined {
    return this.model.formula1;
  }

  setFormula1(formula: string | undefined): void {
    this.model.formula1 = formula;
  }

  formula2(): string | undefined {
    return this.model.formula2;
  }

  setFormula2(formula: string | undefined): void {
    this.model.formula2 = formula;
  }

  toModel(): ExcelDataValidationModel {
    return { ...this.model };
  }

  delete(): void {
    this.remove?.();
  }
}

export class ExcelDataValidationCollection {
  constructor(private readonly getModels: () => ExcelDataValidationModel[]) {}

  items(): ExcelDataValidation[] {
    return this.getModels().map((model) => this.wrap(model));
  }

  add(model: ExcelDataValidationModel): ExcelDataValidation {
    const copy = { ...model };
    this.getModels().push(copy);
    return this.wrap(copy);
  }

  getItem(rangeAddress: string): ExcelDataValidation {
    const item = this.getItemOrNullObject(rangeAddress);
    if (!item) throw new Error(`data validation not found: ${rangeAddress}`);
    return item;
  }

  getItemOrNullObject(rangeAddress: string): ExcelDataValidation | null {
    const model = this.getModels().find((validation) => validation.sqref === rangeAddress);
    return model ? this.wrap(model) : null;
  }

  delete(rangeAddress: string): void {
    const models = this.getModels();
    for (let i = models.length - 1; i >= 0; i--) {
      if (models[i]?.sqref === rangeAddress) models.splice(i, 1);
    }
  }

  toModel(): ExcelDataValidationModel[] {
    return this.getModels().map((model) => ({ ...model }));
  }

  private wrap(model: ExcelDataValidationModel): ExcelDataValidation {
    return new ExcelDataValidation(model, () => {
      const models = this.getModels();
      const index = models.indexOf(model);
      if (index >= 0) models.splice(index, 1);
    });
  }
}

export class ExcelConditionalFormat {
  readonly rules: ExcelConditionalFormatRuleCollection;

  constructor(
    private readonly model: ExcelConditionalFormatModel,
    private readonly remove?: () => void,
  ) {
    this.rules = new ExcelConditionalFormatRuleCollection(() => this.model.rules);
  }

  rangeAddress(): string | undefined {
    return this.model.sqref;
  }

  setRange(address: string): void {
    this.model.sqref = address;
  }

  toModel(): ExcelConditionalFormatModel {
    return cloneConditionalFormat(this.model);
  }

  delete(): void {
    this.remove?.();
  }
}

export class ExcelConditionalFormatRule {
  constructor(
    private readonly model: ExcelConditionalFormatRuleModel,
    private readonly remove?: () => void,
  ) {}

  type(): string | undefined {
    return this.model.type;
  }

  setType(type: string | undefined): void {
    this.model.type = type;
  }

  priority(): number | undefined {
    return this.model.priority;
  }

  setPriority(priority: number | undefined): void {
    this.model.priority = priority;
  }

  operator(): string | undefined {
    return this.model.operator;
  }

  setOperator(operator: string | undefined): void {
    this.model.operator = operator;
  }

  stopIfTrue(): boolean | undefined {
    return this.model.stopIfTrue;
  }

  setStopIfTrue(value: boolean | undefined): void {
    this.model.stopIfTrue = value;
  }

  formulas(): string[] {
    return [...this.model.formulas];
  }

  setFormula(index: number, formula: string): void {
    this.model.formulas[index] = formula;
  }

  addFormula(formula: string): void {
    this.model.formulas.push(formula);
  }

  deleteFormula(index: number): void {
    this.model.formulas.splice(index, 1);
  }

  toModel(): ExcelConditionalFormatRuleModel {
    return { ...this.model, formulas: [...this.model.formulas] };
  }

  delete(): void {
    this.remove?.();
  }
}

export class ExcelConditionalFormatRuleCollection {
  constructor(private readonly getModels: () => ExcelConditionalFormatRuleModel[]) {}

  items(): ExcelConditionalFormatRule[] {
    return this.getModels().map((model) => this.wrap(model));
  }

  getItem(index: number): ExcelConditionalFormatRule {
    const item = this.getItemOrNullObject(index);
    if (!item) throw new Error(`conditional format rule not found: ${index}`);
    return item;
  }

  getItemOrNullObject(index: number): ExcelConditionalFormatRule | null {
    const model = this.getModels()[index];
    return model ? this.wrap(model) : null;
  }

  add(model: ExcelConditionalFormatRuleModel): ExcelConditionalFormatRule {
    const copy = { ...model, formulas: [...model.formulas] };
    this.getModels().push(copy);
    return this.wrap(copy);
  }

  delete(index: number): void {
    this.getModels().splice(index, 1);
  }

  toModel(): ExcelConditionalFormatRuleModel[] {
    return this.getModels().map((model) => ({ ...model, formulas: [...model.formulas] }));
  }

  private wrap(model: ExcelConditionalFormatRuleModel): ExcelConditionalFormatRule {
    return new ExcelConditionalFormatRule(model, () => {
      const models = this.getModels();
      const index = models.indexOf(model);
      if (index >= 0) models.splice(index, 1);
    });
  }
}

export class ExcelConditionalFormatCollection {
  constructor(private readonly getModels: () => ExcelConditionalFormatModel[]) {}

  items(): ExcelConditionalFormat[] {
    return this.getModels().map((model) => this.wrap(model));
  }

  add(model: ExcelConditionalFormatModel): ExcelConditionalFormat {
    const copy = cloneConditionalFormat(model);
    this.getModels().push(copy);
    return this.wrap(copy);
  }

  getItem(rangeAddress: string): ExcelConditionalFormat {
    const item = this.getItemOrNullObject(rangeAddress);
    if (!item) throw new Error(`conditional format not found: ${rangeAddress}`);
    return item;
  }

  getItemOrNullObject(rangeAddress: string): ExcelConditionalFormat | null {
    const model = this.getModels().find((format) => format.sqref === rangeAddress);
    return model ? this.wrap(model) : null;
  }

  delete(rangeAddress: string): void {
    const models = this.getModels();
    for (let i = models.length - 1; i >= 0; i--) {
      if (models[i]?.sqref === rangeAddress) models.splice(i, 1);
    }
  }

  toModel(): ExcelConditionalFormatModel[] {
    return this.getModels().map((model) => cloneConditionalFormat(model));
  }

  private wrap(model: ExcelConditionalFormatModel): ExcelConditionalFormat {
    return new ExcelConditionalFormat(model, () => {
      const models = this.getModels();
      const index = models.indexOf(model);
      if (index >= 0) models.splice(index, 1);
    });
  }
}

export class ExcelComment {
  constructor(
    private readonly model: ExcelCommentModel,
    private readonly remove?: () => void,
  ) {}

  ref(): string {
    return this.model.ref;
  }

  setRef(ref: string): void {
    this.model.ref = ref;
  }

  text(): string {
    return this.model.text;
  }

  setText(text: string): void {
    this.model.text = text;
  }

  toModel(): ExcelCommentModel {
    return { ...this.model };
  }

  delete(): void {
    this.remove?.();
  }
}

export class ExcelCommentCollection {
  constructor(private readonly getModels: () => ExcelCommentModel[]) {}

  items(): ExcelComment[] {
    return this.getModels().map((model) => this.wrap(model));
  }

  add(model: ExcelCommentModel): ExcelComment {
    const copy = { ...model };
    this.getModels().push(copy);
    return this.wrap(copy);
  }

  getItem(ref: string): ExcelComment {
    const item = this.getItemOrNullObject(ref);
    if (!item) throw new Error(`comment not found: ${ref}`);
    return item;
  }

  getItemOrNullObject(ref: string): ExcelComment | null {
    const model = this.getModels().find((comment) => comment.ref === ref);
    return model ? this.wrap(model) : null;
  }

  delete(ref: string): void {
    const models = this.getModels();
    for (let i = models.length - 1; i >= 0; i--) {
      if (models[i]?.ref === ref) models.splice(i, 1);
    }
  }

  toModel(): ExcelCommentModel[] {
    return this.getModels().map((model) => ({ ...model }));
  }

  private wrap(model: ExcelCommentModel): ExcelComment {
    return new ExcelComment(model, () => {
      const models = this.getModels();
      const index = models.indexOf(model);
      if (index >= 0) models.splice(index, 1);
    });
  }
}

export class ExcelThreadedComment {
  constructor(
    private readonly model: ExcelThreadedCommentModel,
    private readonly deleteSelf?: () => void,
  ) {}

  ref(): string {
    return this.model.ref;
  }

  setRef(ref: string): void {
    this.model.ref = ref;
  }

  text(): string {
    return this.model.text;
  }

  setText(text: string): void {
    this.model.text = text;
  }

  delete(): void {
    this.deleteSelf?.();
  }

  toModel(): ExcelThreadedCommentModel {
    return { ...this.model };
  }
}

export class ExcelThreadedCommentCollection {
  constructor(private readonly getModels: () => ExcelThreadedCommentModel[]) {}

  items(): ExcelThreadedComment[] {
    return this.getModels().map((model) => this.wrap(model));
  }

  getItem(ref: string): ExcelThreadedComment {
    const item = this.getItemOrNullObject(ref);
    if (!item) throw new Error(`threaded comment not found: ${ref}`);
    return item;
  }

  getItemOrNullObject(ref: string): ExcelThreadedComment | null {
    const model = this.getModels().find((comment) => comment.ref === ref);
    return model ? this.wrap(model) : null;
  }

  delete(ref: string): void {
    const models = this.getModels();
    const index = models.findIndex((comment) => comment.ref === ref);
    if (index >= 0) models.splice(index, 1);
  }

  toModel(): ExcelThreadedCommentModel[] {
    return this.getModels().map((model) => ({ ...model }));
  }

  private wrap(model: ExcelThreadedCommentModel): ExcelThreadedComment {
    return new ExcelThreadedComment(model, () => {
      const models = this.getModels();
      const index = models.indexOf(model);
      if (index >= 0) models.splice(index, 1);
    });
  }
}

export class ExcelDrawingCollection {
  constructor(private readonly getModels: () => ExcelDrawingModel[]) {}

  items(): ExcelDrawing[] {
    return this.getModels().map((model) => new ExcelDrawing(model));
  }

  toModel(): ExcelDrawingModel[] {
    return this.getModels().map((drawing) => ({
      ...drawing,
      anchors: drawing.anchors.map((anchor) => ({
        ...anchor,
        from: anchor.from ? { ...anchor.from } : undefined,
        to: anchor.to ? { ...anchor.to } : undefined,
        object: anchor.object ? { ...anchor.object } : undefined,
      })),
    }));
  }
}

export class ExcelDrawing {
  readonly anchors: ExcelDrawingAnchorCollection;

  constructor(private readonly model: ExcelDrawingModel) {
    this.anchors = new ExcelDrawingAnchorCollection(() => this.model.anchors);
  }

  relationshipId(): string | undefined {
    return this.model.relationshipId;
  }

  path(): string | undefined {
    return this.model.path;
  }

  toModel(): ExcelDrawingModel {
    return cloneDrawing(this.model);
  }
}

export class ExcelDrawingAnchorCollection {
  constructor(private readonly getModels: () => ExcelDrawingAnchorModel[]) {}

  items(): ExcelDrawingAnchor[] {
    return this.getModels().map((model) => new ExcelDrawingAnchor(model));
  }

  getItem(index: number): ExcelDrawingAnchor {
    const item = this.getItemOrNullObject(index);
    if (!item) throw new Error(`drawing anchor not found: ${index}`);
    return item;
  }

  getItemOrNullObject(index: number): ExcelDrawingAnchor | null {
    const model = this.getModels()[index];
    return model ? new ExcelDrawingAnchor(model) : null;
  }

  toModel(): ExcelDrawingAnchorModel[] {
    return this.getModels().map((anchor) => cloneDrawingAnchor(anchor));
  }
}

export class ExcelDrawingAnchor {
  readonly from: ExcelDrawingMarker;
  readonly to: ExcelDrawingMarker;

  constructor(private readonly model: ExcelDrawingAnchorModel) {
    this.from = new ExcelDrawingMarker(() => (this.model.from ??= {}));
    this.to = new ExcelDrawingMarker(() => (this.model.to ??= {}));
  }

  kind(): ExcelDrawingAnchorModel['kind'] {
    return this.model.kind;
  }

  object(): ExcelDrawingAnchorModel['object'] | undefined {
    return this.model.object ? { ...this.model.object } : undefined;
  }

  toModel(): ExcelDrawingAnchorModel {
    return cloneDrawingAnchor(this.model);
  }
}

export class ExcelDrawingMarker {
  constructor(private readonly getModel: () => ExcelDrawingMarkerModel) {}

  row(): number | undefined {
    return this.getModel().row;
  }

  setRow(row: number | undefined): void {
    this.getModel().row = row;
  }

  column(): number | undefined {
    return this.getModel().column;
  }

  setColumn(column: number | undefined): void {
    this.getModel().column = column;
  }

  rowOffset(): number | undefined {
    return this.getModel().rowOffset;
  }

  setRowOffset(offset: number | undefined): void {
    this.getModel().rowOffset = offset;
  }

  columnOffset(): number | undefined {
    return this.getModel().columnOffset;
  }

  setColumnOffset(offset: number | undefined): void {
    this.getModel().columnOffset = offset;
  }

  toModel(): ExcelDrawingMarkerModel {
    return { ...this.getModel() };
  }
}

export class ExcelChartCollection {
  constructor(
    private readonly getModels: () => ExcelChartModel[],
    private readonly getDrawings?: () => ExcelDrawingModel[],
  ) {}

  items(): ExcelChart[] {
    return this.getModels().map((model) => new ExcelChart(model));
  }

  add(options: ExcelChartAddOptions): ExcelChart {
    validateChartAddOptions(options);
    const models = this.getModels();
    const chart: ExcelChartModel = {
      type: options.type,
      title: options.title,
      categoryRange: options.categoryRange,
      series: options.series.map((series) => ({ ...series })),
      position: {
        from: { ...options.position.from },
        to: { ...options.position.to },
      },
      formulaRefs: chartFormulaRefs(options),
    };
    models.push(chart);
    const drawings = this.getDrawings?.();
    if (drawings) {
      const drawing = drawings[0] ?? { anchors: [] };
      if (!drawings[0]) drawings.push(drawing);
      drawing.anchors.push({
        kind: 'twoCell',
        from: { ...options.position.from },
        to: { ...options.position.to },
        object: {
          kind: 'chart',
          name: options.title,
        },
      });
    }
    return new ExcelChart(chart);
  }

  toModel(): ExcelChartModel[] {
    return this.getModels().map((chart) => ({
      ...chart,
      series: chart.series?.map((series) => ({ ...series })),
      position: chart.position
        ? { from: { ...chart.position.from }, to: { ...chart.position.to } }
        : undefined,
      formulaRefs: [...chart.formulaRefs],
    }));
  }
}

function chartFormulaRefs(options: ExcelChartAddOptions): string[] {
  return [
    options.categoryRange,
    ...options.series.flatMap((series) =>
      series.name?.includes('!') || series.name?.includes('$')
        ? [series.name, series.values]
        : [series.values],
    ),
  ];
}

function validateChartAddOptions(options: ExcelChartAddOptions): void {
  if (!options.type) throw new Error('ExcelChartCollection.add: chart type is required');
  if (!options.categoryRange) throw new Error('ExcelChartCollection.add: categoryRange is required');
  if (!options.series.length) throw new Error('ExcelChartCollection.add: at least one series is required');
  if ((options.type === 'pie' || options.type === 'doughnut') && options.series.length !== 1) {
    throw new Error(`ExcelChartCollection.add: ${options.type} charts require exactly one series`);
  }
  for (const series of options.series) {
    if (!series.values) throw new Error('ExcelChartCollection.add: series values range is required');
  }
}

export class ExcelChart {
  readonly formulaRefs: ExcelFormulaRefCollection;

  constructor(private readonly model: ExcelChartModel) {
    this.formulaRefs = new ExcelFormulaRefCollection(() => this.model.formulaRefs);
  }

  relationshipId(): string | undefined {
    return this.model.relationshipId;
  }

  drawingPath(): string | undefined {
    return this.model.drawingPath;
  }

  path(): string | undefined {
    return this.model.path;
  }

  title(): string | undefined {
    return this.model.title;
  }

  setTitle(title: string | undefined): void {
    this.model.title = title;
  }

  toModel(): ExcelChartModel {
    return { ...this.model, formulaRefs: [...this.model.formulaRefs] };
  }
}

export class ExcelFormulaRefCollection {
  constructor(private readonly getModels: () => string[]) {}

  items(): string[] {
    return [...this.getModels()];
  }

  getItem(index: number): string | undefined {
    return this.getModels()[index];
  }

  setItem(index: number, formulaRef: string): void {
    this.getModels()[index] = formulaRef;
  }

  add(formulaRef: string): void {
    this.getModels().push(formulaRef);
  }

  delete(index: number): void {
    this.getModels().splice(index, 1);
  }

  toModel(): string[] {
    return this.items();
  }
}

export class ExcelImageCollection {
  constructor(private readonly getModels: () => ExcelImageModel[]) {}

  items(): ExcelImage[] {
    return this.getModels().map((model) => new ExcelImage(model));
  }

  toModel(): ExcelImageModel[] {
    return this.getModels().map((image) => ({ ...image }));
  }
}

export class ExcelImage {
  constructor(private readonly model: ExcelImageModel) {}

  relationshipId(): string | undefined {
    return this.model.relationshipId;
  }

  drawingPath(): string | undefined {
    return this.model.drawingPath;
  }

  path(): string | undefined {
    return this.model.path;
  }

  targetMode(): string | undefined {
    return this.model.targetMode;
  }

  name(): string | undefined {
    return this.model.name;
  }

  toModel(): ExcelImageModel {
    return { ...this.model };
  }
}

export class ExcelShapeCollection {
  constructor(private readonly getDrawings: () => ExcelDrawingModel[]) {}

  items(): ExcelShape[] {
    return this.getShapeEntries().map((entry) => new ExcelShape(entry));
  }

  getItem(name: string): ExcelShape {
    const item = this.getItemOrNullObject(name);
    if (!item) throw new Error(`shape not found: ${name}`);
    return item;
  }

  getItemOrNullObject(name: string): ExcelShape | null {
    const entry = this.getShapeEntries().find((item) => item.anchor.object?.name === name);
    return entry ? new ExcelShape(entry) : null;
  }

  toModel(): ExcelShapeModel[] {
    return this.getShapeEntries().map((entry) => ({
      drawingPath: entry.drawing.path,
      name: entry.anchor.object?.name,
      anchor: cloneDrawingAnchor(entry.anchor),
    }));
  }

  private getShapeEntries(): ExcelShapeEntry[] {
    return this.getDrawings().flatMap((drawing) =>
      drawing.anchors
        .filter((anchor) => anchor.object?.kind === 'shape')
        .map((anchor) => ({ drawing, anchor })),
    );
  }
}

export class ExcelShape {
  readonly anchor: ExcelDrawingAnchor;

  constructor(private readonly entry: ExcelShapeEntry) {
    this.anchor = new ExcelDrawingAnchor(entry.anchor);
  }

  drawingPath(): string | undefined {
    return this.entry.drawing.path;
  }

  name(): string | undefined {
    return this.entry.anchor.object?.name;
  }

  setName(name: string | undefined): void {
    this.entry.anchor.object ??= { kind: 'shape' };
    this.entry.anchor.object.name = name;
  }

  toModel(): ExcelShapeModel {
    return {
      drawingPath: this.entry.drawing.path,
      name: this.entry.anchor.object?.name,
      anchor: cloneDrawingAnchor(this.entry.anchor),
    };
  }
}

interface ExcelShapeEntry {
  drawing: ExcelDrawingModel;
  anchor: ExcelDrawingAnchorModel;
}

function cloneTable(table: ExcelTableModel): ExcelTableModel {
  return {
    ...table,
    columns: table.columns.map((column) => ({ ...column })),
    autoFilter: table.autoFilter ? cloneAutoFilter(table.autoFilter) : undefined,
    style: table.style ? { ...table.style } : undefined,
  };
}

function cloneAutoFilter(autoFilter: ExcelTableAutoFilterModel): ExcelTableAutoFilterModel {
  const clone = {
    ...autoFilter,
    filterColumns: autoFilter.filterColumns.map((column) => ({
      ...column,
      filters: column.filters ? [...column.filters] : undefined,
    })),
    sortState: autoFilter.sortState ? cloneSortState(autoFilter.sortState) : undefined,
  };
  if (clone.sortState === undefined) delete clone.sortState;
  for (const column of clone.filterColumns) {
    if (column.filters === undefined) delete column.filters;
  }
  return clone;
}

function cloneSortState(sortState: ExcelTableSortStateModel): ExcelTableSortStateModel {
  return {
    ...sortState,
    conditions: sortState.conditions.map((condition) => ({ ...condition })),
  };
}

function nextTableColumnId(columns: ExcelTableColumnModel[]): number {
  return columns.reduce((max, column) => Math.max(max, column.id), 0) + 1;
}

function assertOptionalNonNegativeInteger(value: number | undefined, operation: string): void {
  if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
    throw new Error(`${operation}: value must be a non-negative integer`);
  }
}

function cloneDrawing(drawing: ExcelDrawingModel): ExcelDrawingModel {
  return {
    ...drawing,
    anchors: drawing.anchors.map((anchor) => cloneDrawingAnchor(anchor)),
  };
}

function cloneDrawingAnchor(anchor: ExcelDrawingAnchorModel): ExcelDrawingAnchorModel {
  return {
    ...anchor,
    from: anchor.from ? { ...anchor.from } : undefined,
    to: anchor.to ? { ...anchor.to } : undefined,
    object: anchor.object ? { ...anchor.object } : undefined,
  };
}

function cloneConditionalFormat(format: ExcelConditionalFormatModel): ExcelConditionalFormatModel {
  return {
    ...format,
    rules: format.rules.map((rule) => ({
      ...rule,
      formulas: [...rule.formulas],
    })),
  };
}
