import { describe, expect, test } from 'vitest';
import { captureExcelCellStyle, importExcelCellStyles } from '../cell-style.js';
import type { ExcelStylesModel } from '../model.js';

function styles(name: string, code: string): ExcelStylesModel {
  return {
    fonts: [{}, { name, size: 26, color: { rgb: '#AA3311' } }],
    fills: [{}, { patternType: 'solid', foregroundColor: { rgb: '#FFFFAA' } }],
    borders: [{}, { left: { style: 'dashed', color: { rgb: '#228866' } } }],
    numberFormats: [{ id: 164, code }],
    cellFormats: [{}, { fontId: 1, fillId: 1, borderId: 1, numberFormatId: 164, alignment: { horizontal: 'center', wrapText: true } }],
  };
}

describe('portable Excel cell styles', () => {
  test('resolves workbook-local style and custom number-format collisions by definition', () => {
    const source = styles('KaiTi', '0.00" 元"');
    const target = styles('Arial', 'yyyy-mm-dd');
    const snapshot = captureExcelCellStyle(source, 1);
    const originalSource = structuredClone(source);
    const [id] = importExcelCellStyles(target, [snapshot]);
    expect(id).toBe(2);
    expect(captureExcelCellStyle(target, id)).toEqual(snapshot);
    expect(target.cellFormats[id!]!.numberFormatId).toBe(165);
    expect(target.numberFormats).toEqual([{ id: 164, code: 'yyyy-mm-dd' }, { id: 165, code: '0.00" 元"' }]);
    expect(source).toEqual(originalSource);
    const after = structuredClone(target);
    expect(importExcelCellStyles(target, [snapshot, snapshot])).toEqual([id, id]);
    expect(target).toEqual(after);
  });

  test('reuses equivalent format definitions even when their component ids differ', () => {
    const source = styles('KaiTi', '0.00%');
    const target = styles('Arial', 'yyyy-mm-dd');
    target.fonts.unshift({ name: 'Calibri' });
    target.cellFormats.forEach((format) => { if (format.fontId !== undefined) format.fontId++; });
    const snapshot = captureExcelCellStyle(source, 1);
    const [first] = importExcelCellStyles(target, [snapshot]);
    const counts = JSON.stringify(target);
    expect(importExcelCellStyles(target, [captureExcelCellStyle(target, first)])).toEqual([first]);
    expect(JSON.stringify(target)).toBe(counts);
    expect(captureExcelCellStyle(target, 0).font).toEqual({ name: 'Calibri' });
  });
});
