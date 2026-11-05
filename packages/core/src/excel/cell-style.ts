import type { ExcelAlignmentModel, ExcelBorderModel, ExcelCellFormatModel, ExcelFillModel, ExcelFontModel, ExcelStylesModel } from './model.js';

/** A cell's appearance without workbook-local font, fill, border or number-format indices. */
export interface ExcelCellStyleSnapshot {
  font: ExcelFontModel;
  fill: ExcelFillModel;
  border: ExcelBorderModel;
  alignment?: ExcelAlignmentModel;
  numberFormatId?: number;
  numberFormatCode?: string;
}

export function captureExcelCellStyle(styles: ExcelStylesModel, styleId?: number): ExcelCellStyleSnapshot {
  const format = styles.cellFormats[styleId ?? 0] ?? {};
  const code = format.numberFormatCode ?? styles.numberFormats.find((item) => item.id === format.numberFormatId)?.code;
  return clone({
    font: styles.fonts[format.fontId ?? 0] ?? {},
    fill: styles.fills[format.fillId ?? 0] ?? {},
    border: styles.borders[format.borderId ?? 0] ?? {},
    ...(format.alignment ? { alignment: format.alignment } : {}),
    ...(code !== undefined ? { numberFormatCode: code } : { numberFormatId: format.numberFormatId ?? 0 }),
  });
}

export function importExcelCellStyles(styles: ExcelStylesModel, snapshots: readonly ExcelCellStyleSnapshot[]): number[] {
  const known = new Map(styles.cellFormats.map((_, id) => [key(captureExcelCellStyle(styles, id)), id]));
  const component = <T>(items: T[], value: T): number => {
    const fingerprint = key(value);
    const existing = items.findIndex((item) => key(item) === fingerprint);
    if (existing !== -1) return existing;
    items.push(clone(value));
    return items.length - 1;
  };
  return snapshots.map((snapshot) => {
    const fingerprint = key(snapshot);
    const found = known.get(fingerprint);
    if (found !== undefined) return found;
    let numberFormatId = snapshot.numberFormatId ?? 0;
    if (snapshot.numberFormatCode !== undefined) {
      const existing = styles.numberFormats.find((item) => item.code === snapshot.numberFormatCode);
      numberFormatId = existing?.id ?? styles.numberFormats.reduce((max, item) => Math.max(max, item.id), 163) + 1;
      if (!existing) styles.numberFormats.push({ id: numberFormatId, code: snapshot.numberFormatCode });
    }
    const format: ExcelCellFormatModel = {
      fontId: component(styles.fonts, snapshot.font),
      fillId: component(styles.fills, snapshot.fill),
      borderId: component(styles.borders, snapshot.border),
      numberFormatId,
      ...(snapshot.numberFormatCode !== undefined ? { numberFormatCode: snapshot.numberFormatCode } : {}),
      ...(snapshot.alignment ? { alignment: clone(snapshot.alignment), applyAlignment: true } : {}),
      applyFont: true, applyFill: true, applyBorder: true, applyNumberFormat: true,
    };
    styles.cellFormats.push(format);
    const id = styles.cellFormats.length - 1;
    known.set(fingerprint, id);
    return id;
  });
}

export function equalExcelCellStyles(left: ExcelCellStyleSnapshot, right: ExcelCellStyleSnapshot): boolean {
  return key(left) === key(right);
}

function clone<T>(value: T): T { return JSON.parse(JSON.stringify(value)) as T; }
function key(value: unknown): string {
  if (value === undefined) return '';
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(key).join(',')}]`;
  return `{${Object.keys(value).sort().filter((name) => (value as Record<string, unknown>)[name] !== undefined).map((name) => `${JSON.stringify(name)}:${key((value as Record<string, unknown>)[name])}`).join(',')}}`;
}
