import type { Package, Part } from '../pack/index.js';
import { getPart } from '../pack/index.js';
import { OFFICE_REL_NS, REL_NS, SS_NS } from './namespaces.js';
import type {
  ExcelCellKind,
  ExcelCellModel,
  ExcelCommentModel,
  ExcelConditionalFormatModel,
  ExcelConditionalFormatRuleModel,
  ExcelDataValidationModel,
  ExcelChartModel,
  ExcelDrawingAnchorModel,
  ExcelDrawingMarkerModel,
  ExcelDrawingModel,
  ExcelDrawingObjectModel,
  ExcelImageModel,
  ExcelColumnModel,
  ExcelHyperlinkModel,
  ExcelMergedRangeModel,
  ExcelRowModel,
  ExcelSheetModel,
  ExcelSheetViewModel,
  ExcelPaneModel,
  ExcelSelectionModel,
  ExcelTableColumnModel,
  ExcelTableAutoFilterModel,
  ExcelTableFilterColumnModel,
  ExcelTableSortConditionModel,
  ExcelTableSortStateModel,
  ExcelTableModel,
  ExcelTableStyleModel,
  ExcelThreadedCommentModel,
  ExcelWorkbookModel,
} from './model.js';
import { parseCellAddress } from './address.js';
import { parseStyles } from './styles-reader.js';
import type { WorkbookSheetRef } from './workbook-reader.js';
import { parseDefinedNames, parseSharedStrings, parseWorkbookDate1904, parseWorkbookSheets } from './workbook-reader.js';

interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

function requireXmlPart(pkg: Package, path: string): Part {
  const part = getPart(pkg, path);
  if (!part?.xmlDom) {
    throw new Error(`ExcelWorkbook: missing XML part ${path}`);
  }
  return part;
}

function optionalXmlPart(pkg: Package, path: string): Part | undefined {
  const part = getPart(pkg, path);
  return part?.xmlDom ? part : undefined;
}

function textOf(el: Element | null | undefined): string {
  if (!el) return '';
  return el.textContent ?? '';
}

function firstChildNs(parent: Element, ns: string, localName: string): Element | null {
  return parent.getElementsByTagNameNS(ns, localName)[0] ?? null;
}

function childElementsNs(parent: Element, ns: string, localName: string): Element[] {
  const result: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child?.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI === ns && el.localName === localName) {
      result.push(el);
    }
  }
  return result;
}

function parseRelationships(part?: Part): Map<string, Relationship> {
  const result = new Map<string, Relationship>();
  if (!part?.xmlDom) return result;

  const rels = part.xmlDom.getElementsByTagNameNS(REL_NS, 'Relationship');
  for (let i = 0; i < rels.length; i++) {
    const rel = rels[i];
    if (!rel) continue;
    const id = rel.getAttribute('Id');
    const type = rel.getAttribute('Type');
    const target = rel.getAttribute('Target');
    if (!id || !type || !target) continue;
    result.set(id, {
      id,
      type,
      target,
      targetMode: rel.getAttribute('TargetMode') ?? undefined,
    });
  }
  return result;
}

function dirname(path: string): string {
  const idx = path.lastIndexOf('/');
  return idx === -1 ? '' : path.slice(0, idx);
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

function resolveRelationshipTarget(sourcePartPath: string, target: string): string {
  if (target.startsWith('/')) return normalizePath(target.slice(1));
  const base = dirname(sourcePartPath);
  return normalizePath(base ? `${base}/${target}` : target);
}

function parseNumber(raw: string | undefined): number | null {
  if (raw === undefined || raw === '') return null;
  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

function parseOptionalNumber(raw: string | null): number | undefined {
  if (raw === null || raw === '') return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function parseBooleanAttr(raw: string | null): boolean | undefined {
  if (raw === null) return undefined;
  return raw === '1' || raw === 'true';
}

const EXCEL_1900_EPOCH_OFFSET_DAYS = 25_569;
const EXCEL_1904_EPOCH_OFFSET_DAYS = 24_107;
const MILLISECONDS_PER_DAY = 86_400_000;

function parseIsoDateSerial(raw: string | undefined, date1904: boolean): number | null {
  if (!raw) return null;
  const match =
    /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?(?:Z|[+-]\d{2}:\d{2})?)?$/.exec(
      raw,
    );
  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4] ?? 0);
  const minute = Number(match[5] ?? 0);
  const second = Number(match[6] ?? 0);
  const millisecond = Number((match[7] ?? '').padEnd(3, '0').slice(0, 3));
  if (month < 1 || month > 12 || hour > 23 || minute > 59 || second > 59) return null;

  // Spreadsheet dates are timezone-naive. Preserve the ISO wall-clock fields even
  // when the serialized value includes an offset instead of shifting the date.
  const timestamp = Date.UTC(year, month - 1, day, hour, minute, second, millisecond);
  const date = new Date(timestamp);
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day ||
    date.getUTCHours() !== hour ||
    date.getUTCMinutes() !== minute ||
    date.getUTCSeconds() !== second ||
    date.getUTCMilliseconds() !== millisecond
  ) return null;

  if (date1904) return timestamp / MILLISECONDS_PER_DAY + EXCEL_1904_EPOCH_OFFSET_DAYS;

  let serial = timestamp / MILLISECONDS_PER_DAY + EXCEL_1900_EPOCH_OFFSET_DAYS;
  // Excel reserves serial 60 for the fictitious 1900-02-29.
  if (serial < 61) serial -= 1;
  return serial;
}

function setIfDefined<T extends object, K extends keyof T>(target: T, key: K, value: T[K] | undefined): void {
  if (value !== undefined) target[key] = value;
}

function parseCell(
  cell: Element,
  sharedStrings: string[],
  date1904: boolean,
): ExcelCellModel | null {
  const address = cell.getAttribute('r');
  if (!address) return null;
  const { row, column } = parseCellAddress(address);
  const dataType = cell.getAttribute('t') ?? undefined;
  const styleIdRaw = cell.getAttribute('s');
  const formula = textOf(firstChildNs(cell, SS_NS, 'f')) || undefined;
  const rawValue = textOf(firstChildNs(cell, SS_NS, 'v')) || undefined;
  let kind: ExcelCellKind = formula ? 'formula' : 'blank';
  let value: string | number | boolean | null = null;
  let sharedStringIndex: number | undefined;

  if (dataType === 's') {
    sharedStringIndex = rawValue === undefined ? undefined : Number(rawValue);
    value =
      sharedStringIndex === undefined || !Number.isInteger(sharedStringIndex)
        ? null
        : (sharedStrings[sharedStringIndex] ?? null);
    kind = formula ? 'formula' : 'string';
  } else if (dataType === 'inlineStr') {
    const inline = firstChildNs(cell, SS_NS, 'is');
    const tEls = inline?.getElementsByTagNameNS(SS_NS, 't') ?? [];
    let text = '';
    for (let i = 0; i < tEls.length; i++) text += textOf(tEls[i]);
    value = text;
    kind = formula ? 'formula' : 'string';
  } else if (dataType === 'b') {
    value = rawValue === '1';
    kind = formula ? 'formula' : 'boolean';
  } else if (dataType === 'str') {
    value = rawValue ?? null;
    kind = formula ? 'formula' : 'string';
  } else if (dataType === 'e') {
    value = rawValue ?? null;
    kind = 'error';
  } else if (dataType === 'd') {
    const serial = parseIsoDateSerial(rawValue, date1904);
    value = serial ?? rawValue ?? null;
    kind = formula ? 'formula' : serial === null ? 'string' : 'number';
  } else if (rawValue !== undefined) {
    value = parseNumber(rawValue);
    kind = formula ? 'formula' : 'number';
  }

  return {
    address,
    row,
    column,
    kind,
    value,
    rawValue,
    formula,
    dataType,
    styleId: styleIdRaw === null ? undefined : Number(styleIdRaw),
    sharedStringIndex,
  };
}

function parseColumns(sheetDom: Document): ExcelColumnModel[] {
  const colEls = sheetDom.getElementsByTagNameNS(SS_NS, 'col');
  const columns: ExcelColumnModel[] = [];
  for (let i = 0; i < colEls.length; i++) {
    const col = colEls[i];
    if (!col) continue;
    const min = parseOptionalNumber(col.getAttribute('min'));
    const max = parseOptionalNumber(col.getAttribute('max'));
    if (min === undefined || max === undefined) continue;
    const column: ExcelColumnModel = { min, max };
    setIfDefined(column, 'width', parseOptionalNumber(col.getAttribute('width')));
    setIfDefined(column, 'hidden', parseBooleanAttr(col.getAttribute('hidden')));
    setIfDefined(column, 'styleId', parseOptionalNumber(col.getAttribute('style')));
    setIfDefined(column, 'customWidth', parseBooleanAttr(col.getAttribute('customWidth')));
    columns.push(column);
  }
  return columns;
}

function parseRows(sheetDom: Document): ExcelRowModel[] {
  const rowEls = sheetDom.getElementsByTagNameNS(SS_NS, 'row');
  const rows: ExcelRowModel[] = [];
  for (let i = 0; i < rowEls.length; i++) {
    const row = rowEls[i];
    if (!row) continue;
    const index = parseOptionalNumber(row.getAttribute('r'));
    if (index === undefined) continue;
    const rowModel: ExcelRowModel = { index };
    setIfDefined(rowModel, 'height', parseOptionalNumber(row.getAttribute('ht')));
    setIfDefined(rowModel, 'hidden', parseBooleanAttr(row.getAttribute('hidden')));
    setIfDefined(rowModel, 'styleId', parseOptionalNumber(row.getAttribute('s')));
    setIfDefined(rowModel, 'customHeight', parseBooleanAttr(row.getAttribute('customHeight')));
    rows.push(rowModel);
  }
  return rows;
}

function parseMergedRanges(sheetDom: Document): ExcelMergedRangeModel[] {
  const mergeCellEls = sheetDom.getElementsByTagNameNS(SS_NS, 'mergeCell');
  const mergedRanges: ExcelMergedRangeModel[] = [];
  for (let i = 0; i < mergeCellEls.length; i++) {
    const ref = mergeCellEls[i]?.getAttribute('ref');
    if (ref) mergedRanges.push({ address: ref });
  }
  return mergedRanges;
}

function sheetRelationshipPartPath(sheetPath: string): string {
  return `${dirname(sheetPath)}/_rels/${sheetPath.slice(sheetPath.lastIndexOf('/') + 1)}.rels`;
}

function parseHyperlinks(pkg: Package, sheetPath: string, sheetDom: Document): ExcelHyperlinkModel[] {
  const rels = parseRelationships(optionalXmlPart(pkg, sheetRelationshipPartPath(sheetPath)));
  const hyperlinkEls = sheetDom.getElementsByTagNameNS(SS_NS, 'hyperlink');
  const hyperlinks: ExcelHyperlinkModel[] = [];
  for (let i = 0; i < hyperlinkEls.length; i++) {
    const hyperlink = hyperlinkEls[i];
    if (!hyperlink) continue;
    const ref = hyperlink.getAttribute('ref');
    if (!ref) continue;
    const relationshipId = hyperlink.getAttributeNS(OFFICE_REL_NS, 'id') ?? undefined;
    const rel = relationshipId ? rels.get(relationshipId) : undefined;
    const model: ExcelHyperlinkModel = { ref };
    setIfDefined(model, 'relationshipId', relationshipId);
    setIfDefined(
      model,
      'target',
      rel ? (rel.targetMode === 'External' ? rel.target : resolveRelationshipTarget(sheetPath, rel.target)) : undefined,
    );
    setIfDefined(model, 'targetMode', rel?.targetMode);
    setIfDefined(model, 'location', hyperlink.getAttribute('location') ?? undefined);
    setIfDefined(model, 'display', hyperlink.getAttribute('display') ?? undefined);
    setIfDefined(model, 'tooltip', hyperlink.getAttribute('tooltip') ?? undefined);
    hyperlinks.push(model);
  }
  return hyperlinks;
}

function parseTableColumns(tableDom: Document): ExcelTableColumnModel[] {
  const tableColumnEls = tableDom.getElementsByTagNameNS(SS_NS, 'tableColumn');
  const columns: ExcelTableColumnModel[] = [];
  for (let i = 0; i < tableColumnEls.length; i++) {
    const el = tableColumnEls[i];
    if (!el) continue;
    const id = parseOptionalNumber(el.getAttribute('id'));
    const name = el.getAttribute('name');
    if (id === undefined || name === null) continue;
    const column: ExcelTableColumnModel = { id, name };
    setIfDefined(column, 'totalsRowLabel', el.getAttribute('totalsRowLabel') ?? undefined);
    setIfDefined(column, 'totalsRowFunction', el.getAttribute('totalsRowFunction') ?? undefined);
    columns.push(column);
  }
  return columns;
}

function parseTableStyle(tableDom: Document): ExcelTableStyleModel | undefined {
  const styleEl = tableDom.getElementsByTagNameNS(SS_NS, 'tableStyleInfo')[0];
  if (!styleEl) return undefined;
  const style: ExcelTableStyleModel = {};
  setIfDefined(style, 'name', styleEl.getAttribute('name') ?? undefined);
  setIfDefined(style, 'showFirstColumn', parseBooleanAttr(styleEl.getAttribute('showFirstColumn')));
  setIfDefined(style, 'showLastColumn', parseBooleanAttr(styleEl.getAttribute('showLastColumn')));
  setIfDefined(style, 'showRowStripes', parseBooleanAttr(styleEl.getAttribute('showRowStripes')));
  setIfDefined(style, 'showColumnStripes', parseBooleanAttr(styleEl.getAttribute('showColumnStripes')));
  return Object.keys(style).length ? style : undefined;
}

function parseTableAutoFilter(tableDom: Document): ExcelTableAutoFilterModel | undefined {
  const tableEl = tableDom.getElementsByTagNameNS(SS_NS, 'table')[0];
  const autoFilterEl = tableEl ? firstChildNs(tableEl, SS_NS, 'autoFilter') : null;
  if (!autoFilterEl) return undefined;
  const model: ExcelTableAutoFilterModel = {
    ref: autoFilterEl.getAttribute('ref') ?? undefined,
    filterColumns: parseFilterColumns(autoFilterEl),
  };
  setIfDefined(model, 'sortState', parseSortState(autoFilterEl));
  if (!model.filterColumns.length && !model.sortState) return undefined;
  return model;
}

function parseFilterColumns(autoFilterEl: Element): ExcelTableFilterColumnModel[] {
  return childElementsNs(autoFilterEl, SS_NS, 'filterColumn').flatMap((filterColumnEl) => {
    const colId = parseOptionalNumber(filterColumnEl.getAttribute('colId'));
    if (colId === undefined) return [];
    const filterColumn: ExcelTableFilterColumnModel = { colId };
    setIfDefined(filterColumn, 'hiddenButton', parseBooleanAttr(filterColumnEl.getAttribute('hiddenButton')));
    setIfDefined(filterColumn, 'showButton', parseBooleanAttr(filterColumnEl.getAttribute('showButton')));
    const filtersEl = firstChildNs(filterColumnEl, SS_NS, 'filters');
    const filters = filtersEl
      ? childElementsNs(filtersEl, SS_NS, 'filter')
          .map((filterEl) => filterEl.getAttribute('val') ?? undefined)
          .filter((value): value is string => value !== undefined)
      : [];
    if (filters.length) filterColumn.filters = filters;
    return [filterColumn];
  });
}

function parseSortState(autoFilterEl: Element): ExcelTableSortStateModel | undefined {
  const sortStateEl = firstChildNs(autoFilterEl, SS_NS, 'sortState');
  if (!sortStateEl) return undefined;
  const state: ExcelTableSortStateModel = { conditions: [] };
  setIfDefined(state, 'ref', sortStateEl.getAttribute('ref') ?? undefined);
  setIfDefined(state, 'caseSensitive', parseBooleanAttr(sortStateEl.getAttribute('caseSensitive')));
  setIfDefined(state, 'sortMethod', sortStateEl.getAttribute('sortMethod') ?? undefined);
  state.conditions = childElementsNs(sortStateEl, SS_NS, 'sortCondition').map((conditionEl) => {
    const condition: ExcelTableSortConditionModel = {};
    setIfDefined(condition, 'ref', conditionEl.getAttribute('ref') ?? undefined);
    setIfDefined(condition, 'descending', parseBooleanAttr(conditionEl.getAttribute('descending')));
    setIfDefined(condition, 'sortBy', conditionEl.getAttribute('sortBy') ?? undefined);
    return condition;
  });
  return state;
}

function parseTablePart(pkg: Package, relationshipId: string, path: string): ExcelTableModel {
  const part = optionalXmlPart(pkg, path);
  if (!part?.xmlDom) {
    return { relationshipId, path, columns: [] };
  }

  const tableEl = part.xmlDom.getElementsByTagNameNS(SS_NS, 'table')[0];
  if (!tableEl) {
    return { relationshipId, path, columns: [] };
  }

  const table: ExcelTableModel = {
    relationshipId,
    path,
    columns: parseTableColumns(part.xmlDom),
  };
  setIfDefined(table, 'id', parseOptionalNumber(tableEl.getAttribute('id')));
  setIfDefined(table, 'name', tableEl.getAttribute('name') ?? undefined);
  setIfDefined(table, 'displayName', tableEl.getAttribute('displayName') ?? undefined);
  setIfDefined(table, 'ref', tableEl.getAttribute('ref') ?? undefined);
  setIfDefined(table, 'headerRowCount', parseOptionalNumber(tableEl.getAttribute('headerRowCount')));
  setIfDefined(table, 'totalsRowCount', parseOptionalNumber(tableEl.getAttribute('totalsRowCount')));
  setIfDefined(table, 'totalsRowShown', parseBooleanAttr(tableEl.getAttribute('totalsRowShown')));
  const autoFilterEl = firstChildNs(tableEl, SS_NS, 'autoFilter');
  const autoFilter = parseTableAutoFilter(part.xmlDom);
  setIfDefined(table, 'autoFilterRef', autoFilterEl?.getAttribute('ref') ?? undefined);
  setIfDefined(table, 'autoFilter', autoFilter);
  setIfDefined(table, 'style', parseTableStyle(part.xmlDom));
  return table;
}

function parseTables(pkg: Package, sheetPath: string, sheetDom: Document): ExcelTableModel[] {
  const rels = parseRelationships(optionalXmlPart(pkg, sheetRelationshipPartPath(sheetPath)));
  const tablePartEls = sheetDom.getElementsByTagNameNS(SS_NS, 'tablePart');
  const tables: ExcelTableModel[] = [];
  for (let i = 0; i < tablePartEls.length; i++) {
    const relationshipId = tablePartEls[i]?.getAttributeNS(OFFICE_REL_NS, 'id') ?? undefined;
    if (!relationshipId) continue;
    const rel = rels.get(relationshipId);
    if (!rel) {
      tables.push({ relationshipId, columns: [] });
      continue;
    }
    tables.push(parseTablePart(pkg, relationshipId, resolveRelationshipTarget(sheetPath, rel.target)));
  }
  return tables;
}

function parseDataValidations(sheetDom: Document): ExcelDataValidationModel[] {
  const validationEls = sheetDom.getElementsByTagNameNS(SS_NS, 'dataValidation');
  const validations: ExcelDataValidationModel[] = [];
  for (let i = 0; i < validationEls.length; i++) {
    const el = validationEls[i];
    if (!el) continue;
    const validation: ExcelDataValidationModel = {};
    setIfDefined(validation, 'sqref', el.getAttribute('sqref') ?? undefined);
    setIfDefined(validation, 'type', el.getAttribute('type') ?? undefined);
    setIfDefined(validation, 'operator', el.getAttribute('operator') ?? undefined);
    setIfDefined(validation, 'allowBlank', parseBooleanAttr(el.getAttribute('allowBlank')));
    setIfDefined(validation, 'showDropDown', parseBooleanAttr(el.getAttribute('showDropDown')));
    setIfDefined(validation, 'showInputMessage', parseBooleanAttr(el.getAttribute('showInputMessage')));
    setIfDefined(validation, 'showErrorMessage', parseBooleanAttr(el.getAttribute('showErrorMessage')));
    setIfDefined(validation, 'errorStyle', el.getAttribute('errorStyle') ?? undefined);
    setIfDefined(validation, 'errorTitle', el.getAttribute('errorTitle') ?? undefined);
    setIfDefined(validation, 'error', el.getAttribute('error') ?? undefined);
    setIfDefined(validation, 'promptTitle', el.getAttribute('promptTitle') ?? undefined);
    setIfDefined(validation, 'prompt', el.getAttribute('prompt') ?? undefined);
    setIfDefined(validation, 'formula1', textOf(firstChildNs(el, SS_NS, 'formula1')) || undefined);
    setIfDefined(validation, 'formula2', textOf(firstChildNs(el, SS_NS, 'formula2')) || undefined);
    validations.push(validation);
  }
  return validations;
}

function parseConditionalFormatRule(ruleEl: Element): ExcelConditionalFormatRuleModel {
  const rule: ExcelConditionalFormatRuleModel = { formulas: [] };
  setIfDefined(rule, 'type', ruleEl.getAttribute('type') ?? undefined);
  setIfDefined(rule, 'priority', parseOptionalNumber(ruleEl.getAttribute('priority')));
  setIfDefined(rule, 'operator', ruleEl.getAttribute('operator') ?? undefined);
  setIfDefined(rule, 'dxfId', parseOptionalNumber(ruleEl.getAttribute('dxfId')));
  setIfDefined(rule, 'stopIfTrue', parseBooleanAttr(ruleEl.getAttribute('stopIfTrue')));
  setIfDefined(rule, 'aboveAverage', parseBooleanAttr(ruleEl.getAttribute('aboveAverage')));
  setIfDefined(rule, 'percent', parseBooleanAttr(ruleEl.getAttribute('percent')));
  setIfDefined(rule, 'bottom', parseBooleanAttr(ruleEl.getAttribute('bottom')));
  setIfDefined(rule, 'rank', parseOptionalNumber(ruleEl.getAttribute('rank')));
  setIfDefined(rule, 'text', ruleEl.getAttribute('text') ?? undefined);
  setIfDefined(rule, 'timePeriod', ruleEl.getAttribute('timePeriod') ?? undefined);
  const formulaEls = childElementsNs(ruleEl, SS_NS, 'formula');
  rule.formulas = formulaEls.map((formula) => textOf(formula));
  return rule;
}

function parseConditionalFormats(sheetDom: Document): ExcelConditionalFormatModel[] {
  const conditionalFormattingEls = sheetDom.getElementsByTagNameNS(SS_NS, 'conditionalFormatting');
  const conditionalFormats: ExcelConditionalFormatModel[] = [];
  for (let i = 0; i < conditionalFormattingEls.length; i++) {
    const el = conditionalFormattingEls[i];
    if (!el) continue;
    const format: ExcelConditionalFormatModel = {
      rules: childElementsNs(el, SS_NS, 'cfRule').map((rule) => parseConditionalFormatRule(rule)),
    };
    setIfDefined(format, 'sqref', el.getAttribute('sqref') ?? undefined);
    conditionalFormats.push(format);
  }
  return conditionalFormats;
}

function relationshipTargetsByType(pkg: Package, sheetPath: string, typeSuffix: string): Array<{ id: string; path: string }> {
  const rels = parseRelationships(optionalXmlPart(pkg, sheetRelationshipPartPath(sheetPath)));
  const targets: Array<{ id: string; path: string }> = [];
  for (const rel of rels.values()) {
    if (rel.type.endsWith(typeSuffix)) {
      targets.push({ id: rel.id, path: resolveRelationshipTarget(sheetPath, rel.target) });
    }
  }
  return targets;
}

function parseComments(pkg: Package, sheetPath: string): ExcelCommentModel[] {
  const targets = relationshipTargetsByType(pkg, sheetPath, '/comments');
  const comments: ExcelCommentModel[] = [];
  for (const target of targets) {
    const part = optionalXmlPart(pkg, target.path);
    if (!part?.xmlDom) continue;
    const authors = Array.from(part.xmlDom.getElementsByTagNameNS(SS_NS, 'author')).map((author) => textOf(author));
    const commentEls = part.xmlDom.getElementsByTagNameNS(SS_NS, 'comment');
    for (let i = 0; i < commentEls.length; i++) {
      const el = commentEls[i];
      if (!el) continue;
      const ref = el.getAttribute('ref');
      if (!ref) continue;
      const authorId = parseOptionalNumber(el.getAttribute('authorId'));
      const comment: ExcelCommentModel = {
        ref,
        text: textOf(firstChildNs(el, SS_NS, 'text')),
      };
      setIfDefined(comment, 'authorId', authorId);
      setIfDefined(comment, 'author', authorId === undefined ? undefined : authors[authorId]);
      setIfDefined(comment, 'shapeId', parseOptionalNumber(el.getAttribute('shapeId')));
      comments.push(comment);
    }
  }
  return comments;
}

function parseThreadedComments(pkg: Package, sheetPath: string): ExcelThreadedCommentModel[] {
  const targets = relationshipTargetsByType(pkg, sheetPath, '/threadedComment');
  const comments: ExcelThreadedCommentModel[] = [];
  for (const target of targets) {
    const part = optionalXmlPart(pkg, target.path);
    if (!part?.xmlDom) continue;
    const commentEls = part.xmlDom.getElementsByTagNameNS('*', 'threadedComment');
    for (let i = 0; i < commentEls.length; i++) {
      const el = commentEls[i];
      if (!el) continue;
      const ref = el.getAttribute('ref');
      if (!ref) continue;
      const comment: ExcelThreadedCommentModel = {
        ref,
        text: textOf(firstChildNs(el, el.namespaceURI ?? '', 'text')) || textOf(el.getElementsByTagNameNS('*', 'text')[0]),
      };
      setIfDefined(comment, 'id', el.getAttribute('id') ?? undefined);
      setIfDefined(comment, 'personId', el.getAttribute('personId') ?? undefined);
      setIfDefined(comment, 'dateTime', el.getAttribute('dT') ?? undefined);
      comments.push(comment);
    }
  }
  return comments;
}

function firstDescendantLocal(parent: Element, localName: string): Element | null {
  return parent.getElementsByTagNameNS('*', localName)[0] ?? null;
}

function descendantsLocal(parent: Element | Document, localName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS('*', localName));
}

function parseDrawingMarker(parent: Element, localName: string): ExcelDrawingMarkerModel | undefined {
  const markerEl = firstDescendantLocal(parent, localName);
  if (!markerEl) return undefined;
  const marker: ExcelDrawingMarkerModel = {};
  setIfDefined(marker, 'column', parseOptionalNumber(firstDescendantLocal(markerEl, 'col')?.textContent ?? null));
  setIfDefined(marker, 'columnOffset', parseOptionalNumber(firstDescendantLocal(markerEl, 'colOff')?.textContent ?? null));
  setIfDefined(marker, 'row', parseOptionalNumber(firstDescendantLocal(markerEl, 'row')?.textContent ?? null));
  setIfDefined(marker, 'rowOffset', parseOptionalNumber(firstDescendantLocal(markerEl, 'rowOff')?.textContent ?? null));
  return Object.keys(marker).length ? marker : undefined;
}

function drawingObjectName(parent: Element): string | undefined {
  return firstDescendantLocal(parent, 'cNvPr')?.getAttribute('name') ?? undefined;
}

function parseDrawingObject(anchorEl: Element, rels: Map<string, Relationship>, drawingPath: string): ExcelDrawingObjectModel | undefined {
  const chartEl = firstDescendantLocal(anchorEl, 'chart');
  if (chartEl) {
    const relationshipId = chartEl.getAttributeNS(OFFICE_REL_NS, 'id') ?? undefined;
    const rel = relationshipId ? rels.get(relationshipId) : undefined;
    return {
      kind: 'chart',
      name: drawingObjectName(anchorEl),
      relationshipId,
      targetPath: rel ? resolveRelationshipTarget(drawingPath, rel.target) : undefined,
      targetMode: rel?.targetMode,
    };
  }

  const blipEl = firstDescendantLocal(anchorEl, 'blip');
  if (blipEl) {
    const relationshipId =
      blipEl.getAttributeNS(OFFICE_REL_NS, 'embed') ?? blipEl.getAttributeNS(OFFICE_REL_NS, 'link') ?? undefined;
    const rel = relationshipId ? rels.get(relationshipId) : undefined;
    return {
      kind: 'image',
      name: drawingObjectName(anchorEl),
      relationshipId,
      targetPath: rel ? resolveRelationshipTarget(drawingPath, rel.target) : undefined,
      targetMode: rel?.targetMode,
    };
  }

  if (firstDescendantLocal(anchorEl, 'sp')) {
    return { kind: 'shape', name: drawingObjectName(anchorEl) };
  }
  if (firstDescendantLocal(anchorEl, 'graphicFrame')) {
    return { kind: 'graphicFrame', name: drawingObjectName(anchorEl) };
  }
  return undefined;
}

function parseDrawingAnchors(drawingDom: Document, rels: Map<string, Relationship>, drawingPath: string): ExcelDrawingAnchorModel[] {
  const anchors: ExcelDrawingAnchorModel[] = [];
  const specs = [
    ['twoCellAnchor', 'twoCell'],
    ['oneCellAnchor', 'oneCell'],
    ['absoluteAnchor', 'absolute'],
  ] as const;

  for (const [localName, kind] of specs) {
    for (const anchorEl of descendantsLocal(drawingDom, localName)) {
      const anchor: ExcelDrawingAnchorModel = { kind };
      setIfDefined(anchor, 'from', parseDrawingMarker(anchorEl, 'from'));
      setIfDefined(anchor, 'to', parseDrawingMarker(anchorEl, 'to'));
      setIfDefined(anchor, 'object', parseDrawingObject(anchorEl, rels, drawingPath));
      anchors.push(anchor);
    }
  }
  return anchors;
}

function drawingRelationshipPartPath(drawingPath: string): string {
  return `${dirname(drawingPath)}/_rels/${drawingPath.slice(drawingPath.lastIndexOf('/') + 1)}.rels`;
}

function chartTitle(chartDom: Document): string | undefined {
  const titleEl = descendantsLocal(chartDom, 'title')[0];
  if (!titleEl) return undefined;
  const text = descendantsLocal(titleEl, 't').map((el) => textOf(el)).join('');
  return text || undefined;
}

function parseChartPart(pkg: Package, relationshipId: string | undefined, drawingPath: string, path: string): ExcelChartModel {
  const part = optionalXmlPart(pkg, path);
  const chart: ExcelChartModel = { relationshipId, drawingPath, path, formulaRefs: [] };
  if (!part?.xmlDom) return chart;
  setIfDefined(chart, 'title', chartTitle(part.xmlDom));
  chart.formulaRefs = descendantsLocal(part.xmlDom, 'f').map((el) => textOf(el)).filter(Boolean);
  return chart;
}

function parseDrawings(pkg: Package, sheetPath: string, sheetDom: Document): {
  drawings: ExcelDrawingModel[];
  charts: ExcelChartModel[];
  images: ExcelImageModel[];
} {
  const sheetRels = parseRelationships(optionalXmlPart(pkg, sheetRelationshipPartPath(sheetPath)));
  const drawings: ExcelDrawingModel[] = [];
  const charts: ExcelChartModel[] = [];
  const images: ExcelImageModel[] = [];

  for (const drawingEl of descendantsLocal(sheetDom, 'drawing')) {
    const relationshipId = drawingEl.getAttributeNS(OFFICE_REL_NS, 'id') ?? undefined;
    if (!relationshipId) continue;
    const sheetRel = sheetRels.get(relationshipId);
    const drawingPath = sheetRel ? resolveRelationshipTarget(sheetPath, sheetRel.target) : undefined;
    if (!drawingPath) {
      drawings.push({ relationshipId, anchors: [] });
      continue;
    }

    const drawingPart = optionalXmlPart(pkg, drawingPath);
    const drawingRels = parseRelationships(optionalXmlPart(pkg, drawingRelationshipPartPath(drawingPath)));
    const anchors = drawingPart?.xmlDom ? parseDrawingAnchors(drawingPart.xmlDom, drawingRels, drawingPath) : [];
    drawings.push({ relationshipId, path: drawingPath, anchors });

    for (const anchor of anchors) {
      if (anchor.object?.kind === 'chart' && anchor.object.targetPath) {
        charts.push(parseChartPart(pkg, anchor.object.relationshipId, drawingPath, anchor.object.targetPath));
      } else if (anchor.object?.kind === 'image' && anchor.object.targetPath) {
        images.push({
          relationshipId: anchor.object.relationshipId,
          drawingPath,
          path: anchor.object.targetPath,
          targetMode: anchor.object.targetMode,
          name: anchor.object.name,
        });
      }
    }
  }

  return { drawings, charts, images };
}

function parsePane(sheetViewEl: Element): ExcelPaneModel | undefined {
  const paneEl = firstChildNs(sheetViewEl, SS_NS, 'pane');
  if (!paneEl) return undefined;
  const pane: ExcelPaneModel = {};
  setIfDefined(pane, 'xSplit', parseOptionalNumber(paneEl.getAttribute('xSplit')));
  setIfDefined(pane, 'ySplit', parseOptionalNumber(paneEl.getAttribute('ySplit')));
  setIfDefined(pane, 'topLeftCell', paneEl.getAttribute('topLeftCell') ?? undefined);
  setIfDefined(pane, 'activePane', paneEl.getAttribute('activePane') ?? undefined);
  setIfDefined(pane, 'state', paneEl.getAttribute('state') ?? undefined);
  return Object.keys(pane).length ? pane : undefined;
}

function parseSelections(sheetViewEl: Element): ExcelSelectionModel[] {
  return childElementsNs(sheetViewEl, SS_NS, 'selection').map((selectionEl) => {
    const selection: ExcelSelectionModel = {};
    setIfDefined(selection, 'pane', selectionEl.getAttribute('pane') ?? undefined);
    setIfDefined(selection, 'activeCell', selectionEl.getAttribute('activeCell') ?? undefined);
    setIfDefined(selection, 'activeCellId', parseOptionalNumber(selectionEl.getAttribute('activeCellId')));
    setIfDefined(selection, 'sqref', selectionEl.getAttribute('sqref') ?? undefined);
    return selection;
  });
}

function parseSheetViews(sheetDom: Document): ExcelSheetViewModel[] {
  const sheetViewEls = sheetDom.getElementsByTagNameNS(SS_NS, 'sheetView');
  const views: ExcelSheetViewModel[] = [];
  for (let i = 0; i < sheetViewEls.length; i++) {
    const el = sheetViewEls[i];
    if (!el) continue;
    const view: ExcelSheetViewModel = { selections: parseSelections(el) };
    setIfDefined(view, 'workbookViewId', parseOptionalNumber(el.getAttribute('workbookViewId')));
    setIfDefined(view, 'view', el.getAttribute('view') ?? undefined);
    setIfDefined(view, 'tabSelected', parseBooleanAttr(el.getAttribute('tabSelected')));
    setIfDefined(view, 'showGridLines', parseBooleanAttr(el.getAttribute('showGridLines')));
    setIfDefined(view, 'showRowColHeaders', parseBooleanAttr(el.getAttribute('showRowColHeaders')));
    setIfDefined(view, 'zoomScale', parseOptionalNumber(el.getAttribute('zoomScale')));
    setIfDefined(view, 'zoomScaleNormal', parseOptionalNumber(el.getAttribute('zoomScaleNormal')));
    setIfDefined(view, 'topLeftCell', el.getAttribute('topLeftCell') ?? undefined);
    setIfDefined(view, 'pane', parsePane(el));
    views.push(view);
  }
  return views;
}

function parseSheet(
  pkg: Package,
  ref: WorkbookSheetRef,
  sharedStrings: string[],
  date1904: boolean,
): ExcelSheetModel {
  if (!ref.path) {
    return {
      ...ref,
      sheetViews: [],
      columns: [],
      rows: [],
      mergedRanges: [],
      hyperlinks: [],
      tables: [],
      dataValidations: [],
      conditionalFormats: [],
      comments: [],
      threadedComments: [],
      drawings: [],
      charts: [],
      images: [],
      cells: [],
    };
  }
  const part = optionalXmlPart(pkg, ref.path);
  if (!part?.xmlDom) {
    return {
      ...ref,
      sheetViews: [],
      columns: [],
      rows: [],
      mergedRanges: [],
      hyperlinks: [],
      tables: [],
      dataValidations: [],
      conditionalFormats: [],
      comments: [],
      threadedComments: [],
      drawings: [],
      charts: [],
      images: [],
      cells: [],
    };
  }
  const dimension = part.xmlDom.getElementsByTagNameNS(SS_NS, 'dimension')[0]?.getAttribute('ref') ?? undefined;
  const cellEls = part.xmlDom.getElementsByTagNameNS(SS_NS, 'c');
  const cells: ExcelCellModel[] = [];
  for (let i = 0; i < cellEls.length; i++) {
    const parsed = cellEls[i] ? parseCell(cellEls[i]!, sharedStrings, date1904) : null;
    if (parsed) cells.push(parsed);
  }
  const drawingInfo = parseDrawings(pkg, ref.path, part.xmlDom);
  return {
    ...ref,
    dimension,
    sheetViews: parseSheetViews(part.xmlDom),
    columns: parseColumns(part.xmlDom),
    rows: parseRows(part.xmlDom),
    mergedRanges: parseMergedRanges(part.xmlDom),
    hyperlinks: parseHyperlinks(pkg, ref.path, part.xmlDom),
    tables: parseTables(pkg, ref.path, part.xmlDom),
    dataValidations: parseDataValidations(part.xmlDom),
    conditionalFormats: parseConditionalFormats(part.xmlDom),
    comments: parseComments(pkg, ref.path),
    threadedComments: parseThreadedComments(pkg, ref.path),
    drawings: drawingInfo.drawings,
    charts: drawingInfo.charts,
    images: drawingInfo.images,
    cells,
  };
}

export function readWorkbookModel(pkg: Package): ExcelWorkbookModel {
  const sharedStrings = parseSharedStrings(pkg);
  const sheetRefs = parseWorkbookSheets(pkg);
  const date1904 = parseWorkbookDate1904(pkg);
  return {
    date1904,
    sharedStrings,
    definedNames: parseDefinedNames(pkg),
    styles: parseStyles(pkg),
    sheets: sheetRefs.map((sheet) => parseSheet(pkg, sheet, sharedStrings, date1904)),
  };
}
