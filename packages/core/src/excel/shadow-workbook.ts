import { toZip, type Package } from '../pack/index.js';
import type {
  ExcelDefinedNameModel,
  ExcelSheetModel,
  ExcelStylesModel,
  ExcelWorkbookModel,
} from './model.js';
import { createWorkbookPackageFromModel } from './xlsx-writer.js';

export interface ExcelShadowWorkbookOptions {
  /** Name written into the one-sheet workbook. Usually the planner's temporary shadow sheet name. */
  sheetName?: string;
  /** Sheet id written into the one-sheet workbook. Defaults to the source sheet id. */
  sheetId?: string;
  /** Workbook styles required by cells, rows, columns, tables, and formats referenced by style ids. */
  styles?: ExcelStylesModel;
  /** Optional workbook defined names to include in the generated package. */
  definedNames?: ExcelDefinedNameModel[];
}

export interface ExcelShadowWorkbookPayload {
  sheetName: string;
  model: ExcelWorkbookModel;
  package: Package;
  bytes: Uint8Array;
  base64: string;
}

export async function createShadowWorkbookPayload(
  sheet: ExcelSheetModel,
  options: ExcelShadowWorkbookOptions = {},
): Promise<ExcelShadowWorkbookPayload> {
  const model = createShadowWorkbookModel(sheet, options);
  const pkg = createWorkbookPackageFromModel(model);
  const bytes = await toZip(pkg);
  return {
    sheetName: model.sheets[0]!.name,
    model,
    package: pkg,
    bytes,
    base64: bytesToBase64(bytes),
  };
}

export function createShadowWorkbookModel(
  sheet: ExcelSheetModel,
  options: ExcelShadowWorkbookOptions = {},
): ExcelWorkbookModel {
  const shadowSheet = cloneJson(sheet);
  shadowSheet.id = options.sheetId ?? shadowSheet.id;
  shadowSheet.name = options.sheetName ?? shadowSheet.name;
  shadowSheet.relationshipId = undefined;
  shadowSheet.path = undefined;

  return {
    sharedStrings: [],
    definedNames: (options.definedNames ?? []).map((name) => ({ ...name })),
    styles: cloneStyles(options.styles ?? emptyStyles()),
    sheets: [shadowSheet],
  };
}

function emptyStyles(): ExcelStylesModel {
  return {
    numberFormats: [],
    fonts: [],
    fills: [],
    borders: [],
    cellFormats: [],
  };
}

function cloneStyles(styles: ExcelStylesModel): ExcelStylesModel {
  return cloneJson(styles);
}

function cloneJson<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}
