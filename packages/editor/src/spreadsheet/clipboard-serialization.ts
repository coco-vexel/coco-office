import type { ExcelCellStyleSnapshot, ExcelCellValue } from '@coco-vexel/coco-office/excel';
import type { SpreadsheetClipboardPayload } from './clipboard.js';

export const SPREADSHEET_CLIPBOARD_MIME = 'application/x-coco-spreadsheet';
const MAX_PAYLOAD_LENGTH = 32 * 1024 * 1024;
const MAX_CELLS = 1_048_576;

export function serializeSpreadsheetClipboardPayload(payload: SpreadsheetClipboardPayload): string {
  return JSON.stringify({ ...payload, schema: 1 });
}

function serializedPayload(data: Pick<DataTransfer, 'getData'>): string | null {
  const custom = data.getData(SPREADSHEET_CLIPBOARD_MIME);
  if (custom.trimStart().startsWith('{')) return custom;
  const encoded = /\bdata-coco-spreadsheet-payload\s*=\s*(["'])(.*?)\1/is.exec(data.getData('text/html'))?.[2];
  if (!encoded) return null;
  if (encoded.length > MAX_PAYLOAD_LENGTH * 3) throw new Error('复制的单元格数据过大。');
  try { return decodeURIComponent(encoded); } catch { throw new Error('复制的单元格数据已损坏，请重新复制。'); }
}

export function clipboardToken(data: Pick<DataTransfer, 'getData'>): string | null {
  const text = serializedPayload(data);
  if (text) {
    if (text.length > MAX_PAYLOAD_LENGTH) throw new Error('复制的单元格数据过大。');
    try { const value = JSON.parse(text) as { id?: unknown }; return typeof value.id === 'string' ? value.id : null; }
    catch { throw new Error('复制的单元格数据已损坏，请重新复制。'); }
  }
  return data.getData(SPREADSHEET_CLIPBOARD_MIME) || /\bdata-coco-spreadsheet="([^"]+)"/i.exec(data.getData('text/html'))?.[1] || null;
}

export function readSpreadsheetClipboardPayload(data: Pick<DataTransfer, 'getData'>): SpreadsheetClipboardPayload | null {
  const text = serializedPayload(data);
  if (!text) return null;
  if (text.length > MAX_PAYLOAD_LENGTH) throw new Error('复制的单元格数据过大。');
  try { return validate(JSON.parse(text)); }
  catch { throw new Error('复制的单元格数据格式无效，请重新复制。'); }
}

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Expected object');
  if (Object.keys(value).some((key) => key === '__proto__' || key === 'constructor' || key === 'prototype')) throw new Error('Invalid property');
  return value as Record<string, unknown>;
}
function string(value: unknown, max = 32768): string {
  if (typeof value !== 'string' || value.length > max) throw new Error('Invalid string');
  return value;
}
function integer(value: unknown, max: number, min = 1): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < min || value > max) throw new Error('Invalid integer');
  return value;
}
function scalar(value: unknown): ExcelCellValue {
  if (value === null || typeof value === 'boolean' || typeof value === 'string' || (typeof value === 'number' && Number.isFinite(value))) return value;
  throw new Error('Invalid cell value');
}
function matrix<T>(value: unknown, rows: number, columns: number, read: (item: unknown) => T): T[][] {
  if (!Array.isArray(value) || value.length !== rows) throw new Error('Invalid rows');
  return value.map((row: unknown) => {
    if (!Array.isArray(row) || row.length !== columns) throw new Error('Invalid columns');
    return row.map(read);
  });
}

function color(value: unknown): { rgb: string } {
  const rgb = string(record(value).rgb, 7);
  if (!/^#[0-9a-f]{6}$/i.test(rgb)) throw new Error('Invalid color');
  return { rgb };
}
function fields(value: unknown, strings: string[], numbers: string[], booleans: string[]): Record<string, unknown> {
  const input = record(value); const result: Record<string, unknown> = {};
  for (const name of strings) if (input[name] !== undefined) result[name] = string(input[name], name === 'name' ? 1024 : 256);
  for (const name of numbers) if (input[name] !== undefined) {
    const number = input[name]; if (typeof number !== 'number' || !Number.isFinite(number) || Math.abs(number) > 65535) throw new Error('Invalid style number');
    result[name] = number;
  }
  for (const name of booleans) if (input[name] !== undefined) {
    if (typeof input[name] !== 'boolean') throw new Error('Invalid style boolean'); result[name] = input[name];
  }
  return result;
}
function style(value: unknown): ExcelCellStyleSnapshot {
  const input = record(value);
  const font = fields(input.font, ['name', 'scheme'], ['size', 'family', 'charset'], ['bold', 'italic', 'underline']);
  const fontInput = record(input.font); if (fontInput.color !== undefined) font.color = color(fontInput.color);
  const fill = fields(input.fill, ['patternType'], [], []);
  const fillInput = record(input.fill);
  for (const name of ['foregroundColor', 'backgroundColor']) if (fillInput[name] !== undefined) fill[name] = color(fillInput[name]);
  const borderInput = record(input.border); const border: Record<string, unknown> = {};
  for (const name of ['top', 'bottom', 'left', 'right', 'diagonal']) if (borderInput[name] !== undefined) {
    const side = record(borderInput[name]);
    border[name] = { ...(side.style !== undefined ? { style: string(side.style, 64) } : {}), ...(side.color !== undefined ? { color: color(side.color) } : {}) };
  }
  return {
    font, fill, border,
    ...(input.alignment !== undefined ? { alignment: fields(input.alignment, ['horizontal', 'vertical'], ['textRotation', 'indent'], ['wrapText', 'shrinkToFit']) } : {}),
    ...(input.numberFormatId !== undefined ? { numberFormatId: integer(input.numberFormatId, 65535, 0) } : {}),
    ...(input.numberFormatCode !== undefined ? { numberFormatCode: string(input.numberFormatCode, 4096) } : {}),
  } as ExcelCellStyleSnapshot;
}

function validate(value: unknown): SpreadsheetClipboardPayload {
  const input = record(value);
  if (input.schema !== 1) throw new Error('Unsupported clipboard schema');
  const range = record(input.sourceRange);
  const sourceRange = {
    startRow: integer(range.startRow, 1_048_576), endRow: integer(range.endRow, 1_048_576),
    startColumn: integer(range.startColumn, 16384), endColumn: integer(range.endColumn, 16384),
  };
  const rows = sourceRange.endRow - sourceRange.startRow + 1; const columns = sourceRange.endColumn - sourceRange.startColumn + 1;
  if (rows < 1 || columns < 1 || rows * columns > MAX_CELLS) throw new Error('Invalid range');
  const values = matrix(input.values, rows, columns, scalar);
  const formulas = matrix(input.formulas, rows, columns, (item) => item === null ? null : string(item));
  const styleIds = matrix(input.styleIds, rows, columns, (item) => item === null || item === undefined ? undefined : integer(item, 1_048_576, 0));
  if (!Array.isArray(input.merges) || input.merges.length > rows * columns) throw new Error('Invalid merges');
  const covered = new Set<number>();
  const merges = input.merges.map((item: unknown) => {
    const merge = record(item);
    const result = { startRowOffset: integer(merge.startRowOffset, rows - 1, 0), startColumnOffset: integer(merge.startColumnOffset, columns - 1, 0), rowCount: integer(merge.rowCount, rows), columnCount: integer(merge.columnCount, columns) };
    if (result.startRowOffset + result.rowCount > rows || result.startColumnOffset + result.columnCount > columns) throw new Error('Merge out of range');
    for (let r = result.startRowOffset; r < result.startRowOffset + result.rowCount; r++) for (let c = result.startColumnOffset; c < result.startColumnOffset + result.columnCount; c++) {
      const key = r * columns + c; if (covered.has(key)) throw new Error('Overlapping merges'); covered.add(key);
    }
    return result;
  });
  const styles: Record<string, ExcelCellStyleSnapshot> = {};
  for (const [key, snapshot] of Object.entries(record(input.styles))) {
    if (key !== 'default' && !/^(0|[1-9]\d{0,6})$/.test(key)) throw new Error('Invalid style key');
    styles[key] = style(snapshot);
  }
  for (const row of styleIds) for (const id of row) if (!Object.prototype.hasOwnProperty.call(styles, id === undefined ? 'default' : String(id))) throw new Error('Missing style definition');
  return {
    schema: 1, id: string(input.id, 256), sourceSheet: string(input.sourceSheet, 128),
    ...(input.sourceWorkbookId !== undefined ? { sourceWorkbookId: string(input.sourceWorkbookId, 512) } : {}),
    sourceRange, text: string(input.text, MAX_PAYLOAD_LENGTH), values, formulas, styleIds, styles, merges,
  };
}
