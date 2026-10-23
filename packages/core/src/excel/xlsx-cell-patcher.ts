import { deletePart, getPart, markDirty, type Package } from '../pack/index.js';
import { formatRangeAddress, parseCellAddress, parseRangeAddress } from './address.js';
import type { ExcelCellModel, ExcelSheetModel, ExcelWorkbookModel } from './model.js';
import { REL_NS, SS_NS } from './namespaces.js';

const CALC_CHAIN_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

export interface ExcelWorksheetCellPatch {
  sheet: ExcelSheetModel;
  addresses: string[];
  dimension: string | undefined;
}

/**
 * Detect the narrow package-backed edit case that can be written without rebuilding
 * workbook.xml, styles.xml, or an entire worksheet from the normalized SDK model.
 */
export function collectContentOnlyCellPatches(
  before: ExcelWorkbookModel,
  after: ExcelWorkbookModel,
): ExcelWorksheetCellPatch[] | null {
  if (
    !sameJson(before.sharedStrings, after.sharedStrings) ||
    !sameJson(before.definedNames, after.definedNames) ||
    !sameJson(before.styles, after.styles) ||
    before.sheets.length !== after.sheets.length
  ) {
    return null;
  }

  const patches: ExcelWorksheetCellPatch[] = [];
  for (let index = 0; index < before.sheets.length; index += 1) {
    const previousSheet = before.sheets[index]!;
    const currentSheet = after.sheets[index]!;
    if (!sameJson(sheetWithoutCellContent(previousSheet), sheetWithoutCellContent(currentSheet))) {
      return null;
    }

    const previousCells = new Map(previousSheet.cells.map((cell) => [cell.address, cell]));
    const currentCells = new Map(currentSheet.cells.map((cell) => [cell.address, cell]));
    const addresses = new Set([...previousCells.keys(), ...currentCells.keys()]);
    const changed: string[] = [];
    for (const address of addresses) {
      const previous = previousCells.get(address);
      const current = currentCells.get(address);
      if (sameJson(previous, current)) continue;
      if (!isContentOnlyCellChange(previous, current)) return null;
      changed.push(address);
    }
    if (changed.length) {
      changed.sort(compareCellAddresses);
      patches.push({
        sheet: currentSheet,
        addresses: changed,
        dimension: preservedDimension(
          previousSheet.dimension,
          currentSheet.dimension,
          changed,
          currentCells,
        ),
      });
    }
  }
  return patches;
}

export function patchWorkbookCellsInPackage(
  pkg: Package,
  patches: readonly ExcelWorksheetCellPatch[],
): void {
  if (patches.length === 0) return;
  assertPatchableFormulas(pkg, patches);
  invalidateCalculationChain(pkg);
  for (const patch of patches) {
    const path = patch.sheet.path;
    if (!path) throw new Error(`ExcelWorkbook: worksheet ${patch.sheet.name} has no package path`);
    const part = getPart(pkg, path);
    if (!part?.xmlDom) throw new Error(`ExcelWorkbook: missing XML part ${path}`);
    const worksheet = part.xmlDom.documentElement;
    const sheetData = directChild(worksheet, 'sheetData');
    if (!sheetData) throw new Error(`ExcelWorkbook: worksheet ${path} has no sheetData`);

    const currentCells = new Map(patch.sheet.cells.map((cell) => [cell.address, cell]));
    for (const address of patch.addresses) {
      patchCell(part.xmlDom, sheetData, address, currentCells.get(address));
    }
    patchDimension(part.xmlDom, worksheet, patch.dimension);
    markDirty(part);
  }
}

function assertPatchableFormulas(pkg: Package, patches: readonly ExcelWorksheetCellPatch[]): void {
  for (const patch of patches) {
    if (!patch.sheet.path) continue;
    const worksheet = getPart(pkg, patch.sheet.path)?.xmlDom?.documentElement;
    const sheetData = worksheet ? directChild(worksheet, 'sheetData') : undefined;
    if (!sheetData) continue;
    const addresses = new Set(patch.addresses);
    for (const row of directChildren(sheetData, 'row')) {
      for (const cell of directChildren(row, 'c')) {
        const address = cell.getAttribute('r');
        if (!address || !addresses.has(parseCellAddress(address).address)) continue;
        const formula = directChild(cell, 'f');
        const formulaType = formula?.getAttribute('t');
        if (formulaType === 'shared' || formulaType === 'array' || formulaType === 'dataTable') {
          throw new Error(
            `ExcelWorkbook: fidelity-preserving cell patch does not support ${formulaType} formula cell ${patch.sheet.name}!${address}`,
          );
        }
      }
    }
  }
}

function invalidateCalculationChain(pkg: Package): void {
  const relsPart = getPart(pkg, 'xl/_rels/workbook.xml.rels');
  const relationships = relsPart?.xmlDom?.getElementsByTagNameNS(REL_NS, 'Relationship') ?? [];
  const removedPaths: string[] = [];
  for (const relationship of Array.from(relationships)) {
    if (relationship.getAttribute('Type') !== CALC_CHAIN_REL_TYPE) continue;
    const target = relationship.getAttribute('Target');
    if (target) removedPaths.push(resolveRelationshipTarget('xl/workbook.xml', target));
    relationship.parentNode?.removeChild(relationship);
  }
  if (!removedPaths.length) return;
  if (relsPart) markDirty(relsPart);

  for (const path of removedPaths) deletePart(pkg, path);
  const contentTypes = getPart(pkg, '[Content_Types].xml');
  if (contentTypes?.xmlDom) {
    const overrides = contentTypes.xmlDom.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
    for (const override of Array.from(overrides)) {
      const path = override.getAttribute('PartName')?.replace(/^\//, '');
      if (path && removedPaths.includes(path)) override.parentNode?.removeChild(override);
    }
    markDirty(contentTypes);
  }

  const workbookPart = getPart(pkg, 'xl/workbook.xml');
  const workbook = workbookPart?.xmlDom?.documentElement;
  if (!workbookPart?.xmlDom || !workbook) return;
  let calcPr = directChild(workbook, 'calcPr');
  if (!calcPr) {
    calcPr = workbookPart.xmlDom.createElementNS(SS_NS, 'calcPr');
    workbook.insertBefore(calcPr, firstDirectChild(workbook, 'extLst'));
  }
  calcPr.setAttribute('fullCalcOnLoad', '1');
  calcPr.setAttribute('forceFullCalc', '1');
  markDirty(workbookPart);
}

function resolveRelationshipTarget(sourcePath: string, target: string): string {
  if (target.startsWith('/')) return normalizePath(target.slice(1));
  const slash = sourcePath.lastIndexOf('/');
  const directory = slash === -1 ? '' : sourcePath.slice(0, slash);
  return normalizePath(directory ? `${directory}/${target}` : target);
}

function normalizePath(path: string): string {
  const parts: string[] = [];
  for (const part of path.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop();
    else parts.push(part);
  }
  return parts.join('/');
}

function sheetWithoutCellContent(
  sheet: ExcelSheetModel,
): Omit<ExcelSheetModel, 'cells' | 'dimension'> {
  const { cells: _cells, dimension: _dimension, ...rest } = sheet;
  return rest;
}

function isContentOnlyCellChange(
  previous: ExcelCellModel | undefined,
  current: ExcelCellModel | undefined,
): boolean {
  if (!previous) return current?.styleId === undefined;
  if (!current) return true;
  return (
    previous.address === current.address &&
    previous.row === current.row &&
    previous.column === current.column &&
    previous.styleId === current.styleId
  );
}

function patchCell(
  document: Document,
  sheetData: Element,
  address: string,
  model: ExcelCellModel | undefined,
): void {
  const parsed = parseCellAddress(address);
  let row = directChildren(sheetData, 'row').find(
    (candidate) => Number(candidate.getAttribute('r')) === parsed.row,
  );
  let cell = row
    ? directChildren(row, 'c').find((candidate) => candidate.getAttribute('r') === parsed.address)
    : undefined;

  if (!model) {
    if (cell?.parentNode) cell.parentNode.removeChild(cell);
    return;
  }

  if (!row) {
    row = document.createElementNS(SS_NS, 'row');
    row.setAttribute('r', String(parsed.row));
    const nextRow = directChildren(sheetData, 'row').find(
      (candidate) => Number(candidate.getAttribute('r')) > parsed.row,
    );
    sheetData.insertBefore(row, nextRow ?? null);
  }
  if (!cell) {
    cell = document.createElementNS(SS_NS, 'c');
    cell.setAttribute('r', parsed.address);
    const nextCell = directChildren(row, 'c').find((candidate) => {
      const ref = candidate.getAttribute('r');
      return ref ? parseCellAddress(ref).column > parsed.column : false;
    });
    row.insertBefore(cell, nextCell ?? firstDirectChild(row, 'extLst') ?? null);
  }
  patchCellContent(document, cell, model);
}

function patchCellContent(document: Document, cell: Element, model: ExcelCellModel): void {
  for (const child of directChildren(cell)) {
    if (child.localName === 'f' || child.localName === 'v' || child.localName === 'is') {
      cell.removeChild(child);
    }
  }
  cell.removeAttribute('t');
  const insertionPoint = firstDirectChild(cell, 'extLst');

  if (model.kind === 'string') {
    cell.setAttribute('t', 'inlineStr');
    const inline = document.createElementNS(SS_NS, 'is');
    const text = document.createElementNS(SS_NS, 't');
    const value = String(model.value ?? '');
    if (/^\s|\s$/.test(value)) {
      text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    }
    text.textContent = value;
    inline.appendChild(text);
    cell.insertBefore(inline, insertionPoint ?? null);
    return;
  }

  if (model.kind === 'boolean') cell.setAttribute('t', 'b');
  if (model.kind === 'error') cell.setAttribute('t', 'e');
  if (model.formula) appendTextElement(document, cell, 'f', model.formula, insertionPoint);
  if (model.kind === 'blank' && !model.formula) return;
  if (model.value !== null && model.value !== undefined) {
    const value = model.kind === 'boolean' ? (model.value ? '1' : '0') : String(model.value);
    appendTextElement(document, cell, 'v', value, insertionPoint);
  }
}

function appendTextElement(
  document: Document,
  parent: Element,
  name: string,
  value: string,
  before: Element | null,
): void {
  const element = document.createElementNS(SS_NS, name);
  element.textContent = value;
  parent.insertBefore(element, before);
}

function patchDimension(
  document: Document,
  worksheet: Element,
  dimension: string | undefined,
): void {
  const existing = directChild(worksheet, 'dimension');
  if (!dimension) {
    if (existing?.parentNode) existing.parentNode.removeChild(existing);
    return;
  }
  if (existing) {
    existing.setAttribute('ref', dimension);
    return;
  }
  const next = ['sheetViews', 'sheetFormatPr', 'cols', 'sheetData']
    .map((name) => directChild(worksheet, name))
    .find((element): element is Element => Boolean(element));
  const created = document.createElementNS(SS_NS, 'dimension');
  created.setAttribute('ref', dimension);
  worksheet.insertBefore(created, next ?? null);
}

function directChild(parent: Element, localName: string): Element | undefined {
  return directChildren(parent, localName)[0];
}

function firstDirectChild(parent: Element, localName: string): Element | null {
  return directChild(parent, localName) ?? null;
}

function directChildren(parent: Element, localName?: string): Element[] {
  const result: Element[] = [];
  for (let index = 0; index < parent.childNodes.length; index += 1) {
    const child = parent.childNodes[index];
    if (child?.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI !== SS_NS) continue;
    if (localName && element.localName !== localName) continue;
    result.push(element);
  }
  return result;
}

function compareCellAddresses(left: string, right: string): number {
  const a = parseCellAddress(left);
  const b = parseCellAddress(right);
  return a.row - b.row || a.column - b.column;
}

function preservedDimension(
  previous: string | undefined,
  current: string | undefined,
  changedAddresses: readonly string[],
  currentCells: ReadonlyMap<string, ExcelCellModel>,
): string | undefined {
  if (!previous) return current;
  const range = parseRangeAddress(previous);
  let minRow = range.startRow;
  let minColumn = range.startColumn;
  let maxRow = range.startRow + range.rowCount - 1;
  let maxColumn = range.startColumn + range.columnCount - 1;
  for (const address of changedAddresses) {
    if (!currentCells.has(address)) continue;
    const cell = parseCellAddress(address);
    minRow = Math.min(minRow, cell.row);
    minColumn = Math.min(minColumn, cell.column);
    maxRow = Math.max(maxRow, cell.row);
    maxColumn = Math.max(maxColumn, cell.column);
  }
  return formatRangeAddress(minRow, minColumn, maxRow - minRow + 1, maxColumn - minColumn + 1);
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}
