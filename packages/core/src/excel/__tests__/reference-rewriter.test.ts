import { describe, expect, test } from 'vitest';
import {
  renameSheetInFormula,
  renameSheetInReference,
  rewriteFormulaForColumnDelete,
  rewriteFormulaForColumnInsert,
  rewriteFormulaForRowDelete,
  rewriteFormulaForRowInsert,
  rewriteReferenceForColumnDelete,
  rewriteReferenceForColumnInsert,
  rewriteReferenceForRowDelete,
  rewriteReferenceForRowInsert,
  shiftFormulaReferences,
  shiftReferenceRelative,
  rewriteFormulaForRowSort,
} from '../index.js';

describe('Excel reference rewriter', () => {
  test('rewrites relative references for row sorting', () => {
    const rowMap = new Map([[2, 3], [3, 2]]);
    expect(rewriteFormulaForRowSort('=A2+$B$2+C$3+"A2"', rowMap)).toBe('=A3+$B$2+C$3+"A2"');
  });

  test('renames sheet-qualified references', () => {
    expect(renameSheetInReference('Sheet1!A1:B2', 'Sheet1', 'Renamed')).toBe('Renamed!A1:B2');
    expect(renameSheetInReference("'My Sheet'!A1", 'My Sheet', "Bob's Sheet")).toBe(
      "'Bob''s Sheet'!A1",
    );
    expect(renameSheetInReference('Other!A1', 'Sheet1', 'Renamed')).toBe('Other!A1');
  });

  test('rewrites references for row insertion', () => {
    expect(rewriteReferenceForRowInsert('A1', 2)).toEqual({ ok: true, address: 'A1' });
    expect(rewriteReferenceForRowInsert('A2', 2)).toEqual({ ok: true, address: 'A3' });
    expect(rewriteReferenceForRowInsert('A1:B2', 2, 3)).toEqual({
      ok: true,
      address: 'A1:B5',
    });
    expect(rewriteReferenceForRowInsert('Other!A2', 2, 1, { sheetName: 'Sheet1' })).toEqual({
      ok: true,
      address: 'Other!A2',
    });
  });

  test('rewrites references for row deletion', () => {
    expect(rewriteReferenceForRowDelete('A5', 2, 2)).toEqual({ ok: true, address: 'A3' });
    expect(rewriteReferenceForRowDelete('A2', 2)).toEqual({ ok: false, address: '#REF!' });
    expect(rewriteReferenceForRowDelete('A1:A5', 2, 2)).toEqual({
      ok: true,
      address: 'A1:A3',
    });
    expect(rewriteReferenceForRowDelete('A2:A3', 2, 2)).toEqual({
      ok: false,
      address: '#REF!',
    });
  });

  test('rewrites references for column insertion', () => {
    expect(rewriteReferenceForColumnInsert('A1', 2)).toEqual({ ok: true, address: 'A1' });
    expect(rewriteReferenceForColumnInsert('B1', 2)).toEqual({ ok: true, address: 'C1' });
    expect(rewriteReferenceForColumnInsert('A1:B2', 2, 2)).toEqual({
      ok: true,
      address: 'A1:D2',
    });
  });

  test('rewrites references for column deletion', () => {
    expect(rewriteReferenceForColumnDelete('E1', 2, 2)).toEqual({ ok: true, address: 'C1' });
    expect(rewriteReferenceForColumnDelete('B1', 2)).toEqual({ ok: false, address: '#REF!' });
    expect(rewriteReferenceForColumnDelete('A1:E2', 2, 2)).toEqual({
      ok: true,
      address: 'A1:C2',
    });
    expect(rewriteReferenceForColumnDelete('B1:C2', 2, 2)).toEqual({
      ok: false,
      address: '#REF!',
    });
  });

  test('normalizes absolute and reversed references', () => {
    expect(rewriteReferenceForRowInsert('$D$4:$B$2', 3)).toEqual({
      ok: true,
      address: 'B2:D5',
    });
  });

  test('preserves absolute markers for non-reversed references', () => {
    expect(rewriteReferenceForRowInsert('Sheet1!$A$2:$A$10', 2)).toEqual({
      ok: true,
      address: 'Sheet1!$A$3:$A$11',
    });
    expect(rewriteFormulaForColumnInsert('$A1+B$2', 2)).toBe('$A1+C$2');
  });

  test('rejects invalid rewrite arguments', () => {
    expect(() => rewriteReferenceForRowInsert('A1', 0)).toThrow(/rowIndex/);
    expect(() => rewriteReferenceForColumnDelete('A1', 1, 0)).toThrow(/count/);
  });

  test('renames sheet-qualified references inside formulas', () => {
    expect(renameSheetInFormula('SUM(Sheet1!A1:B2)+Other!A1', 'Sheet1', 'Renamed')).toBe(
      'SUM(Renamed!A1:B2)+Other!A1',
    );
    expect(renameSheetInFormula("'My Sheet'!A1+1", 'My Sheet', "Bob's Sheet")).toBe(
      "'Bob''s Sheet'!A1+1",
    );
  });

  test('rewrites formula references for row and column insertion', () => {
    expect(rewriteFormulaForRowInsert('SUM(A1:B2)+C3', 2, 2)).toBe('SUM(A1:B4)+C5');
    expect(rewriteFormulaForColumnInsert('A1+B1+Sheet2!B2', 2, 1)).toBe(
      'A1+C1+Sheet2!C2',
    );
    expect(rewriteFormulaForColumnInsert('A1+B1+Sheet2!B2', 2, 1, { sheetName: 'Sheet1' })).toBe(
      'A1+C1+Sheet2!B2',
    );
  });

  test('rewrites formula references for row and column deletion', () => {
    expect(rewriteFormulaForRowDelete('SUM(A1:A5)+B2', 2, 2)).toBe('SUM(A1:A3)+#REF!');
    expect(rewriteFormulaForColumnDelete('A1+E1+B1:C2', 2, 2)).toBe('A1+C1+#REF!');
  });

  test('shifts only relative references when a formula cell is filled', () => {
    expect(shiftReferenceRelative('Sheet1!$A1:B$2', 3, 2)).toEqual({
      ok: true,
      address: 'Sheet1!$A4:D$2',
    });
    expect(shiftFormulaReferences('SUM(A1:B2)+$C3+D$4+"E5"', 2, 1)).toBe(
      'SUM(B3:C4)+$C5+E$4+"E5"',
    );
    expect(shiftFormulaReferences('A1', -1, 0)).toBe('#REF!');
  });

  test('skips formula string literals and function-like identifiers', () => {
    expect(rewriteFormulaForRowInsert('IF(A1="B2",LOG10(100),A2)', 2)).toBe(
      'IF(A1="B2",LOG10(100),A3)',
    );
    expect(rewriteFormulaForRowInsert('CONCAT("A1:B2"," "" C3")&A1', 1)).toBe(
      'CONCAT("A1:B2"," "" C3")&A2',
    );
  });
});
