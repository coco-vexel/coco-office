import { describe, expect, test } from 'vitest';
import { rewriteMovedCellFormula, type SpreadsheetCellMove } from '../clipboard-move.js';

const move: SpreadsheetCellMove = {
  sourceSheet: 'Sheet1', targetSheet: 'Sheet1', formulaSheet: 'Sheet1', resultSheet: 'Sheet1',
  source: { startRow: 2, endRow: 3, startColumn: 2, endColumn: 3 }, rowOffset: 4, columnOffset: 4,
};

describe('cut formula references', () => {
  test.each([
    ['B2+$B$2+B$2+$B2', 'F6+$F$6+F$6+$F6'],
    ['SUM($B2:C$3)', 'SUM($F6:G$7)'],
    ['SUM(C3:$B$2)', 'SUM(G7:$F$6)'],
    ['SUM(A1:D4)', 'SUM(A1:D4)'],
    ['SUM(B2:B10)', 'SUM(B4:B10)'],
    ['SUM(A2:C3)', 'SUM(A2:A3)'],
    ['SUM(B1:C3)', 'SUM(B1:C1)'],
    ['SUM(B2:D3)', 'SUM(D2:D3)'],
    ['SUM(B1:B10)', 'SUM(B1:B10)'],
    ["'sheet1'!B2+'Unrelated'!b2", "'Sheet1'!F6+'Unrelated'!b2"],
    ['[Other.xlsx]Sheet1!B2+B2', '[Other.xlsx]Sheet1!B2+F6'],
    ["'[Other.xlsx]Sheet1'!B2+B2", "'[Other.xlsx]Sheet1'!B2+F6"],
    ['SUM(Table1[B2])+B2', 'SUM(Table1[B2])+F6'],
    ['SUM(Sheet1:Sheet3!B2)+B2', 'SUM(Sheet1:Sheet3!B2)+F6'],
    ['IF(B2="B2",B2,0)', 'IF(F6="B2",F6,0)'],
    ['SUM(F6:G7)', 'SUM(#REF!)'],
    ['F6+$G$7', '#REF!+#REF!'],
    ['SUM(F5:G7)', 'SUM(F5:G5)'],
    ['SUM(F6:H7)', 'SUM(H6:H7)'],
    ['SUM(E6:G7)', 'SUM(E6:E7)'],
    ['SUM(F6:G8)', 'SUM(F8:G8)'],
    ['SUM(E5:H8)', 'SUM(E5:H8)'],
  ])('%s → %s', (formula, expected) => {
    expect(rewriteMovedCellFormula(formula, move)).toBe(expected);
  });

  test('overlapping moves follow the source once and only invalidate overwritten destination cells', () => {
    expect(rewriteMovedCellFormula('A1+B1+C1+D1', {
      ...move, source: { startRow: 1, endRow: 1, startColumn: 1, endColumn: 3 }, rowOffset: 0, columnOffset: 1,
    })).toBe('B1+C1+D1+#REF!');
  });

  test('cutting the first member into the middle of a sum range shrinks its source edge', () => {
    expect(rewriteMovedCellFormula('SUM($A$2:$D$2)', {
      ...move, source: { startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 }, rowOffset: 0, columnOffset: 1,
    })).toBe('SUM($B$2:$D$2)');
  });

  test('cutting the last member into the middle of a sum range shrinks its source edge', () => {
    expect(rewriteMovedCellFormula('SUM($A$2:$D$2)', {
      ...move, source: { startRow: 2, endRow: 2, startColumn: 4, endColumn: 4 }, rowOffset: 0, columnOffset: -2,
    })).toBe('SUM($A$2:$C$2)');
  });

  test('cross-sheet moves retain outside references, escape sheet names, and update dependents on other sheets', () => {
    const cross = { ...move, targetSheet: "New ' Sheet", resultSheet: "New ' Sheet" };
    expect(rewriteMovedCellFormula('B2+A1', cross)).toBe("F6+'Sheet1'!A1");
    expect(rewriteMovedCellFormula("'Sheet1'!$B$2+'Other'!B2", { ...cross, formulaSheet: 'Other', resultSheet: 'Other' }))
      .toBe("'New '' Sheet'!$F$6+'Other'!B2");
    expect(rewriteMovedCellFormula("'New '' Sheet'!F6+B2", { ...cross, resultSheet: 'Sheet1' }))
      .toBe("#REF!+'New '' Sheet'!F6");
  });

  test('resolves an unquoted Chinese sheet prefix rather than treating it as the current sheet', () => {
    expect(rewriteMovedCellFormula('来源!B2+其他!B2', { ...move, sourceSheet: '来源', formulaSheet: '其他', resultSheet: '其他' }))
      .toBe("'Sheet1'!F6+其他!B2");
  });
});
