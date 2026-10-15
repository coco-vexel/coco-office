import { describe, expect, test } from 'vitest';
import {
  columnIndexToName,
  columnNameToIndex,
  formatCellAddress,
  formatRangeAddress,
  parseCellAddress,
  parseRangeAddress,
  splitSheetQualifiedAddress,
} from '../index.js';

describe('Excel address utilities', () => {
  test('converts column names and indexes', () => {
    expect(columnNameToIndex('A')).toBe(1);
    expect(columnNameToIndex('Z')).toBe(26);
    expect(columnNameToIndex('AA')).toBe(27);
    expect(columnNameToIndex('XFD')).toBe(16384);

    expect(columnIndexToName(1)).toBe('A');
    expect(columnIndexToName(26)).toBe('Z');
    expect(columnIndexToName(27)).toBe('AA');
    expect(columnIndexToName(16384)).toBe('XFD');
  });

  test('parses and formats cells with normalized casing and absolute markers', () => {
    expect(parseCellAddress('$aa$10')).toEqual({ address: 'AA10', row: 10, column: 27 });
    expect(parseCellAddress(' Sheet1!b2 ')).toEqual({ address: 'B2', row: 2, column: 2 });
    expect(formatCellAddress(3, 28)).toBe('AB3');
  });

  test('parses range addresses and normalizes reversed coordinates', () => {
    expect(parseRangeAddress('A1:D10')).toEqual({
      address: 'A1:D10',
      sheetName: undefined,
      startRow: 1,
      startColumn: 1,
      rowCount: 10,
      columnCount: 4,
    });
    expect(parseRangeAddress('$D$10:$A$1')).toEqual({
      address: 'A1:D10',
      sheetName: undefined,
      startRow: 1,
      startColumn: 1,
      rowCount: 10,
      columnCount: 4,
    });
    expect(parseRangeAddress('B2')).toMatchObject({
      address: 'B2',
      startRow: 2,
      startColumn: 2,
      rowCount: 1,
      columnCount: 1,
    });
    expect(formatRangeAddress(2, 2, 3, 4)).toBe('B2:E4');
  });

  test('parses whole-column and whole-row Excel A1 references', () => {
    expect(parseRangeAddress('A:A')).toEqual({
      address: 'A1:A1048576',
      sheetName: undefined,
      startRow: 1,
      startColumn: 1,
      rowCount: 1048576,
      columnCount: 1,
    });
    expect(parseRangeAddress('$C:$A')).toEqual({
      address: 'A1:C1048576',
      sheetName: undefined,
      startRow: 1,
      startColumn: 1,
      rowCount: 1048576,
      columnCount: 3,
    });
    expect(parseRangeAddress('1:7')).toEqual({
      address: 'A1:XFD7',
      sheetName: undefined,
      startRow: 1,
      startColumn: 1,
      rowCount: 7,
      columnCount: 16384,
    });
    expect(parseRangeAddress("'My Sheet'!$7:$1")).toEqual({
      address: 'A1:XFD7',
      sheetName: 'My Sheet',
      startRow: 1,
      startColumn: 1,
      rowCount: 7,
      columnCount: 16384,
    });
  });

  test('parses sheet-qualified addresses including quoted sheet names', () => {
    expect(splitSheetQualifiedAddress('Sheet1!A1:D2')).toEqual({
      sheetName: 'Sheet1',
      address: 'A1:D2',
    });
    expect(splitSheetQualifiedAddress("'My Sheet'!$A$1")).toEqual({
      sheetName: 'My Sheet',
      address: '$A$1',
    });
    expect(splitSheetQualifiedAddress("'Bob''s Sheet'!C3")).toEqual({
      sheetName: "Bob's Sheet",
      address: 'C3',
    });
    expect(parseRangeAddress("'My Sheet'!$D$4:$B$2")).toEqual({
      address: 'B2:D4',
      sheetName: 'My Sheet',
      startRow: 2,
      startColumn: 2,
      rowCount: 3,
      columnCount: 3,
    });
  });

  test('rejects invalid addresses', () => {
    expect(() => parseCellAddress('A0')).toThrow(/invalid cell address/);
    expect(() => parseCellAddress('1A')).toThrow(/invalid cell address/);
    expect(() => parseRangeAddress('A')).toThrow(/invalid cell address/);
    expect(() => parseRangeAddress('1')).toThrow(/invalid cell address/);
    expect(() => parseRangeAddress('A:1')).toThrow(/invalid range address/);
    expect(() => parseRangeAddress('A1:B2:C3')).toThrow(/invalid range address/);
    expect(() => columnIndexToName(0)).toThrow(/invalid column index/);
    expect(() => formatRangeAddress(1, 1, 0, 1)).toThrow(/invalid row count/);
  });
});
