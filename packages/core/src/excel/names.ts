import type { ExcelDefinedNameModel } from './model.js';

export class ExcelNamedItem {
  private readonly model: ExcelDefinedNameModel;
  private readonly names: ExcelDefinedNameModel[];

  constructor(model: ExcelDefinedNameModel, names: ExcelDefinedNameModel[]) {
    this.model = model;
    this.names = names;
  }

  name(): string {
    return this.model.name;
  }

  text(): string {
    return this.model.text;
  }

  setText(text: string): void {
    this.model.text = text;
  }

  rename(name: string): void {
    assertValidName(name);
    if (this.names.some((item) => item !== this.model && item.name === name)) {
      throw new Error(`ExcelWorkbook: defined name already exists: ${name}`);
    }
    this.model.name = name;
  }

  localSheetId(): number | undefined {
    return this.model.localSheetId;
  }

  setLocalSheetId(localSheetId: number | undefined): void {
    if (localSheetId !== undefined && (!Number.isInteger(localSheetId) || localSheetId < 0)) {
      throw new Error(`ExcelNamedItem.setLocalSheetId: localSheetId must be a non-negative integer`);
    }
    this.model.localSheetId = localSheetId;
  }

  toModel(): ExcelDefinedNameModel {
    return { ...this.model };
  }
}

export class ExcelNamedItemCollection {
  private readonly itemsByModel: ExcelNamedItem[];
  private readonly models: ExcelDefinedNameModel[];

  constructor(names: ExcelDefinedNameModel[]) {
    this.models = names;
    this.itemsByModel = names.map((name) => new ExcelNamedItem(name, names));
  }

  items(): ExcelNamedItem[] {
    return [...this.itemsByModel];
  }

  getItem(name: string): ExcelNamedItem {
    const item = this.getItemOrNullObject(name);
    if (!item) throw new Error(`ExcelWorkbook: defined name not found: ${name}`);
    return item;
  }

  getItemOrNullObject(name: string): ExcelNamedItem | null {
    return this.itemsByModel.find((item) => item.name() === name) ?? null;
  }

  add(model: ExcelDefinedNameModel): ExcelNamedItem {
    assertValidName(model.name);
    if (this.getItemOrNullObject(model.name)) {
      throw new Error(`ExcelWorkbook: defined name already exists: ${model.name}`);
    }
    if (model.localSheetId !== undefined && (!Number.isInteger(model.localSheetId) || model.localSheetId < 0)) {
      throw new Error(`ExcelNamedItemCollection.add: localSheetId must be a non-negative integer`);
    }
    const next = { ...model };
    this.models.push(next);
    const item = new ExcelNamedItem(next, this.models);
    this.itemsByModel.push(item);
    return item;
  }

  deleteItem(name: string): void {
    const index = this.models.findIndex((item) => item.name === name);
    if (index === -1) throw new Error(`ExcelWorkbook: defined name not found: ${name}`);
    this.models.splice(index, 1);
    this.itemsByModel.splice(index, 1);
  }
}

function assertValidName(name: string): void {
  if (!name || !/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) {
    throw new Error(`ExcelNamedItem: invalid defined name: ${name}`);
  }
}
