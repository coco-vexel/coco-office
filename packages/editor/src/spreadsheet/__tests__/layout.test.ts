import { describe, expect, test } from 'vitest';
import {
  cssPixelsToExcelColumnWidth,
  excelColumnWidthToCssPixels,
} from '../layout.js';

describe('spreadsheet column width conversion', () => {
  test('matches the ISO 29500 Calibri 11 example', () => {
    expect(excelColumnWidthToCssPixels(8.7109375, 7)).toBe(61);
  });

  test('uses the workbook Normal font metric and round-trips editor resizing', () => {
    expect(excelColumnWidthToCssPixels(6.85546875, 8.2)).toBe(56);
    const width = cssPixelsToExcelColumnWidth(53, 8.2);
    expect(width).toBe(6.45703125);
    expect(excelColumnWidthToCssPixels(width, 8.2)).toBe(53);
  });

  test('matches the WPS copybook grid at 100% zoom', () => {
    expect(excelColumnWidthToCssPixels(5.889999866485596, 8.2)).toBe(48);
    expect(39.75 * 1.2).toBeCloseTo(47.7);
  });
});
