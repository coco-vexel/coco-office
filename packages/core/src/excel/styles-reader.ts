import type { Package } from '../pack/index.js';
import type {
  ExcelAlignmentModel,
  ExcelBorderModel,
  ExcelBorderSideModel,
  ExcelCellFormatModel,
  ExcelColorModel,
  ExcelFillModel,
  ExcelFontModel,
  ExcelNumberFormatModel,
  ExcelStylesModel,
} from './model.js';
import { SS_NS } from './namespaces.js';
import { getPart } from '../pack/index.js';
import { applyExcelTint, ooxmlArgbToExcelColor } from './color.js';

const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const THEME_COLOR_ORDER = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;
const INDEXED_COLORS = [
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
] as const;

type ThemeColors = Map<number, ExcelColorModel>;

function firstChildNs(parent: Element, ns: string, localName: string): Element | null {
  return parent.getElementsByTagNameNS(ns, localName)[0] ?? null;
}

function childElementsNs(parent: Element, ns: string, localName: string): Element[] {
  const elements = parent.getElementsByTagNameNS(ns, localName);
  return Array.from(elements).filter((el) => el.parentElement === parent);
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

function setIfDefined<T extends object, K extends keyof T>(
  target: T,
  key: K,
  value: T[K] | undefined,
): void {
  if (value !== undefined) target[key] = value;
}

function parseColor(
  el: Element | null | undefined,
  themeColors: ThemeColors,
): ExcelColorModel | undefined {
  if (!el) return undefined;
  let color = ooxmlArgbToExcelColor(el.getAttribute('rgb') ?? undefined);
  const theme = parseOptionalNumber(el.getAttribute('theme'));
  if (!color && theme !== undefined) color = themeColors.get(theme);
  const indexed = parseOptionalNumber(el.getAttribute('indexed'));
  if (!color && indexed !== undefined) {
    const rgb =
      INDEXED_COLORS[indexed] ??
      (indexed === 64 ? '000000' : indexed === 65 ? 'FFFFFF' : undefined);
    if (rgb) color = { rgb: `#${rgb}` };
  }
  if (!color && parseBooleanAttr(el.getAttribute('auto'))) color = { rgb: '#000000' };
  return color ? applyExcelTint(color, parseOptionalNumber(el.getAttribute('tint'))) : undefined;
}

function parseThemeColors(pkg: Package): ThemeColors {
  const themeColors: ThemeColors = new Map();
  const themeDom = getPart(pkg, 'xl/theme/theme1.xml')?.xmlDom;
  if (!themeDom) return themeColors;
  const scheme = themeDom.getElementsByTagNameNS(DRAWING_NS, 'clrScheme')[0];
  if (!scheme) return themeColors;
  THEME_COLOR_ORDER.forEach((name, index) => {
    const slot = Array.from(scheme.getElementsByTagNameNS(DRAWING_NS, name)).find(
      (element) => element.parentElement === scheme,
    );
    const source = slot?.firstElementChild;
    const raw = source?.getAttribute('val') ?? source?.getAttribute('lastClr') ?? undefined;
    const color = ooxmlArgbToExcelColor(raw);
    if (color) themeColors.set(index, color);
  });
  return themeColors;
}

function parseValNumber(parent: Element, name: string): number | undefined {
  return parseOptionalNumber(firstChildNs(parent, SS_NS, name)?.getAttribute('val') ?? null);
}

function parseValString(parent: Element, name: string): string | undefined {
  return firstChildNs(parent, SS_NS, name)?.getAttribute('val') ?? undefined;
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

function parseNumberFormats(stylesDom: Document): {
  list: ExcelNumberFormatModel[];
  byId: Map<number, string>;
} {
  const list: ExcelNumberFormatModel[] = [];
  const byId = new Map<number, string>();
  const numFmtEls = stylesDom.getElementsByTagNameNS(SS_NS, 'numFmt');
  for (let i = 0; i < numFmtEls.length; i++) {
    const el = numFmtEls[i];
    if (!el) continue;
    const id = parseOptionalNumber(el.getAttribute('numFmtId'));
    const code = el.getAttribute('formatCode');
    if (id === undefined || code === null) continue;
    list.push({ id, code });
    byId.set(id, code);
  }
  return { list, byId };
}

function parseFonts(stylesDom: Document, themeColors: ThemeColors): ExcelFontModel[] {
  const fontsEl = stylesDom.getElementsByTagNameNS(SS_NS, 'fonts')[0];
  if (!fontsEl) return [];
  return childElementsNs(fontsEl, SS_NS, 'font').map((fontEl) => {
    const font: ExcelFontModel = {};
    setIfDefined(font, 'name', parseValString(fontEl, 'name'));
    setIfDefined(font, 'size', parseValNumber(fontEl, 'sz'));
    setIfDefined(font, 'family', parseValNumber(fontEl, 'family'));
    setIfDefined(font, 'charset', parseValNumber(fontEl, 'charset'));
    setIfDefined(font, 'scheme', parseValString(fontEl, 'scheme'));
    setIfDefined(font, 'bold', firstChildNs(fontEl, SS_NS, 'b') ? true : undefined);
    setIfDefined(font, 'italic', firstChildNs(fontEl, SS_NS, 'i') ? true : undefined);
    setIfDefined(font, 'underline', firstChildNs(fontEl, SS_NS, 'u') ? true : undefined);
    setIfDefined(font, 'color', parseColor(firstChildNs(fontEl, SS_NS, 'color'), themeColors));
    return font;
  });
}

function parseFills(stylesDom: Document, themeColors: ThemeColors): ExcelFillModel[] {
  const fillsEl = stylesDom.getElementsByTagNameNS(SS_NS, 'fills')[0];
  if (!fillsEl) return [];
  return childElementsNs(fillsEl, SS_NS, 'fill').map((fillEl) => {
    const pattern = firstChildNs(fillEl, SS_NS, 'patternFill');
    const fill: ExcelFillModel = {};
    if (!pattern) return fill;
    setIfDefined(fill, 'patternType', pattern.getAttribute('patternType') ?? undefined);
    setIfDefined(
      fill,
      'foregroundColor',
      parseColor(firstChildNs(pattern, SS_NS, 'fgColor'), themeColors),
    );
    setIfDefined(
      fill,
      'backgroundColor',
      parseColor(firstChildNs(pattern, SS_NS, 'bgColor'), themeColors),
    );
    return fill;
  });
}

function parseBorderSide(
  borderEl: Element,
  name: string,
  themeColors: ThemeColors,
): ExcelBorderSideModel | undefined {
  const sideEl = firstChildNs(borderEl, SS_NS, name);
  if (!sideEl) return undefined;
  const side: ExcelBorderSideModel = {};
  setIfDefined(side, 'style', sideEl.getAttribute('style') ?? undefined);
  setIfDefined(side, 'color', parseColor(firstChildNs(sideEl, SS_NS, 'color'), themeColors));
  return Object.keys(side).length ? side : undefined;
}

function parseBorders(stylesDom: Document, themeColors: ThemeColors): ExcelBorderModel[] {
  const bordersEl = stylesDom.getElementsByTagNameNS(SS_NS, 'borders')[0];
  if (!bordersEl) return [];
  return childElementsNs(bordersEl, SS_NS, 'border').map((borderEl) => {
    const border: ExcelBorderModel = {};
    setIfDefined(border, 'left', parseBorderSide(borderEl, 'left', themeColors));
    setIfDefined(border, 'right', parseBorderSide(borderEl, 'right', themeColors));
    setIfDefined(border, 'top', parseBorderSide(borderEl, 'top', themeColors));
    setIfDefined(border, 'bottom', parseBorderSide(borderEl, 'bottom', themeColors));
    setIfDefined(border, 'diagonal', parseBorderSide(borderEl, 'diagonal', themeColors));
    return border;
  });
}

function parseAlignment(xfEl: Element): ExcelAlignmentModel | undefined {
  const alignmentEl = firstChildNs(xfEl, SS_NS, 'alignment');
  if (!alignmentEl) return undefined;
  const alignment: ExcelAlignmentModel = {};
  setIfDefined(alignment, 'horizontal', alignmentEl.getAttribute('horizontal') ?? undefined);
  setIfDefined(alignment, 'vertical', alignmentEl.getAttribute('vertical') ?? undefined);
  setIfDefined(
    alignment,
    'textRotation',
    parseOptionalNumber(alignmentEl.getAttribute('textRotation')),
  );
  setIfDefined(alignment, 'wrapText', parseBooleanAttr(alignmentEl.getAttribute('wrapText')));
  setIfDefined(alignment, 'shrinkToFit', parseBooleanAttr(alignmentEl.getAttribute('shrinkToFit')));
  setIfDefined(alignment, 'indent', parseOptionalNumber(alignmentEl.getAttribute('indent')));
  return Object.keys(alignment).length ? alignment : undefined;
}

function parseCellFormats(
  stylesDom: Document,
  numberFormatsById: Map<number, string>,
): ExcelCellFormatModel[] {
  const cellXfsEl = stylesDom.getElementsByTagNameNS(SS_NS, 'cellXfs')[0];
  if (!cellXfsEl) return [];
  return childElementsNs(cellXfsEl, SS_NS, 'xf').map((xfEl) => {
    const numberFormatId = parseOptionalNumber(xfEl.getAttribute('numFmtId'));
    const format: ExcelCellFormatModel = {};
    setIfDefined(format, 'numberFormatId', numberFormatId);
    setIfDefined(
      format,
      'numberFormatCode',
      numberFormatId === undefined
        ? undefined
        : (numberFormatsById.get(numberFormatId) ?? builtInNumberFormatCode(numberFormatId)),
    );
    setIfDefined(format, 'fontId', parseOptionalNumber(xfEl.getAttribute('fontId')));
    setIfDefined(format, 'fillId', parseOptionalNumber(xfEl.getAttribute('fillId')));
    setIfDefined(format, 'borderId', parseOptionalNumber(xfEl.getAttribute('borderId')));
    setIfDefined(format, 'xfId', parseOptionalNumber(xfEl.getAttribute('xfId')));
    setIfDefined(
      format,
      'applyNumberFormat',
      parseBooleanAttr(xfEl.getAttribute('applyNumberFormat')),
    );
    setIfDefined(format, 'applyFont', parseBooleanAttr(xfEl.getAttribute('applyFont')));
    setIfDefined(format, 'applyFill', parseBooleanAttr(xfEl.getAttribute('applyFill')));
    setIfDefined(format, 'applyBorder', parseBooleanAttr(xfEl.getAttribute('applyBorder')));
    setIfDefined(format, 'applyAlignment', parseBooleanAttr(xfEl.getAttribute('applyAlignment')));
    setIfDefined(format, 'alignment', parseAlignment(xfEl));
    return format;
  });
}

function builtInNumberFormatCode(id: number): string | undefined {
  const builtIns: Record<number, string> = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
  };
  return builtIns[id];
}

export function parseStyles(pkg: Package): ExcelStylesModel {
  const part = getPart(pkg, 'xl/styles.xml');
  if (!part?.xmlDom) return emptyStyles();
  const numberFormats = parseNumberFormats(part.xmlDom);
  const themeColors = parseThemeColors(pkg);
  return {
    numberFormats: numberFormats.list,
    fonts: parseFonts(part.xmlDom, themeColors),
    fills: parseFills(part.xmlDom, themeColors),
    borders: parseBorders(part.xmlDom, themeColors),
    cellFormats: parseCellFormats(part.xmlDom, numberFormats.byId),
  };
}
