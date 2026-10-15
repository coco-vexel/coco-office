import { describe, expect, test } from 'vitest';
import type { ExcelSheetModel } from '../model.js';
import {
  SpreadsheetFormulaEngine,
  calculateSpreadsheetSheetFormulas,
  parseSpreadsheetFormula,
} from '../formula/engine.js';

describe('spreadsheet formula parser', () => {
  test('parses calls, ranges, references, and arithmetic precedence', () => {
    expect(parseSpreadsheetFormula('SUM(A1:A3)+B1*2')).toMatchObject({
      kind: 'binary',
      operator: '+',
      left: { kind: 'call', name: 'SUM' },
      right: { kind: 'binary', operator: '*' },
    });
  });

  test('parses string concatenation and percentage postfix', () => {
    expect(parseSpreadsheetFormula('"Rate: "&A1*10%')).toMatchObject({
      kind: 'binary',
      operator: '&',
      right: {
        kind: 'binary',
        operator: '*',
        right: { kind: 'unary', operator: '%' },
      },
    });
  });

  test('parses quoted cross-sheet ranges', () => {
    expect(parseSpreadsheetFormula("SUMPRODUCT(('收款单'!B$2:B$7=A2)*('收款单'!D$2:D$7))")).toMatchObject({
      kind: 'call',
      name: 'SUMPRODUCT',
    });
  });

  test('parses unquoted cross-sheet cells', () => {
    expect(parseSpreadsheetFormula('SUM(Sheet2!A1:A2)')).toMatchObject({
      kind: 'call',
      name: 'SUM',
      args: [{ kind: 'range', sheetName: 'Sheet2', startAddress: 'A1', endAddress: 'A2' }],
    });
  });

  test('parses LET names separately from function calls', () => {
    expect(parseSpreadsheetFormula('LET(total,A1+B1,total*2)')).toMatchObject({
      kind: 'call',
      name: 'LET',
      args: [{ kind: 'name', name: 'total' }, { kind: 'binary', operator: '+' }, { kind: 'binary', operator: '*' }],
    });
  });
});

describe('spreadsheet formula engine', () => {
  test('calculates same-sheet cell references and arithmetic', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'C7', row: 7, column: 3, kind: 'number', value: 12171 },
        { address: 'D7', row: 7, column: 4, kind: 'number', value: 12171 },
        { address: 'E7', row: 7, column: 5, kind: 'formula', value: null, formula: 'C7+D7' },
      ]),
    );

    expect(results.get('E7')).toMatchObject({
      value: 24342,
      displayText: '24342',
      error: null,
    });
  });

  test('supports aggregate functions over ranges', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 1 },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 2 },
        { address: 'A3', row: 3, column: 1, kind: 'number', value: 3 },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'SUM(A1:A3)' },
        {
          address: 'B2',
          row: 2,
          column: 2,
          kind: 'formula',
          value: null,
          formula: 'AVERAGE(A1:A3)',
        },
      ]),
    );

    expect(results.get('B1')?.value).toBe(6);
    expect(results.get('B2')?.value).toBe(2);
  });

  test('supports IF and comparisons', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 10 },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'IF(A1>5,"ok","no")' },
      ]),
    );

    expect(results.get('B1')?.value).toBe('ok');
  });

  test('supports logical functions and IFERROR', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 10 },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 0 },
        {
          address: 'B1',
          row: 1,
          column: 2,
          kind: 'formula',
          value: null,
          formula: 'AND(A1>5,NOT(A2>0),OR(FALSE(),TRUE()))',
        },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'IFERROR(A1/A2,"fallback")' },
      ]),
    );

    expect(results.get('B1')?.value).toBe(true);
    expect(results.get('B2')?.value).toBe('fallback');
  });

  test('supports text functions, concatenation, and percentages', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: '  Coco Office  ' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'Agent' },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'UPPER(TRIM(A1))&" "&LEFT(A2,2)' },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'TEXTJOIN("-",TRUE(),RIGHT(A2,3),MID(A2,2,2))' },
        { address: 'B3', row: 3, column: 2, kind: 'formula', value: null, formula: '200*10%' },
      ]),
    );

    expect(results.get('B1')?.value).toBe('COCO OFFICE Ag');
    expect(results.get('B2')?.value).toBe('ent-ge');
    expect(results.get('B3')?.value).toBe(20);
  });

  test('supports TEXTBEFORE and TEXTAFTER for delimiter-based extraction', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: '张三 13812345678 北京' },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'TEXTBEFORE(A1," ")' },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'TEXTBEFORE(TEXTAFTER(A1," ")," ")' },
        { address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: 'TEXTAFTER(A1," ")' },
        { address: 'E1', row: 1, column: 5, kind: 'formula', value: null, formula: 'TEXTBEFORE(A1," ",-1)' },
        { address: 'F1', row: 1, column: 6, kind: 'formula', value: null, formula: 'TEXTAFTER(A1," ",-1)' },
      ]),
    );

    expect(results.get('B1')?.value).toBe('张三');
    expect(results.get('C1')?.value).toBe('13812345678');
    expect(results.get('D1')?.value).toBe('13812345678 北京');
    expect(results.get('E1')?.value).toBe('张三 13812345678');
    expect(results.get('F1')?.value).toBe('北京');
  });

  test('supports TODAY and DATEDIF for completed age in years', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'G2', row: 2, column: 7, kind: 'string', value: '1990-06-15' },
        { address: 'H2', row: 2, column: 8, kind: 'formula', value: null, formula: 'DATEDIF(G2,TODAY(),"Y")' },
      ]),
      { today: new Date('2026-06-14T12:00:00Z') },
    );

    expect(results.get('H2')).toMatchObject({ value: 35, error: null });
  });

  test('supports common date, text, validation, rounding, and lookup helpers', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'Coco-Office-Coco' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'B-2' },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 12.345 },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 18 },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'FIND("Office",A1)' },
        { address: 'C2', row: 2, column: 3, kind: 'formula', value: null, formula: 'SUBSTITUTE(A1,"Coco","Core",2)' },
        { address: 'C3', row: 3, column: 3, kind: 'formula', value: null, formula: 'TEXT(B1,"0.00")' },
        { address: 'C4', row: 4, column: 3, kind: 'formula', value: null, formula: 'ROUNDUP(B1,2)' },
        { address: 'C5', row: 5, column: 3, kind: 'formula', value: null, formula: 'XLOOKUP(A2,A2:A2,B2:B2,"missing")' },
        { address: 'C6', row: 6, column: 3, kind: 'formula', value: null, formula: 'XMATCH(A2,A2:A2,0)' },
        { address: 'C7', row: 7, column: 3, kind: 'formula', value: null, formula: 'DATE(2024,2,29)' },
        { address: 'C8', row: 8, column: 3, kind: 'formula', value: null, formula: 'YEAR(C7)' },
        { address: 'C9', row: 9, column: 3, kind: 'formula', value: null, formula: 'COUNTBLANK(A3:A4)' },
        { address: 'C10', row: 10, column: 3, kind: 'formula', value: null, formula: 'ISNUMBER(B1)' },
        { address: 'C11', row: 11, column: 3, kind: 'formula', value: null, formula: 'IFNA(XMATCH("missing",A2:A2,0),99)' },
        { address: 'C12', row: 12, column: 3, kind: 'formula', value: null, formula: 'ISBLANK(A3)' },
        { address: 'C13', row: 13, column: 3, kind: 'formula', value: null, formula: 'ISTEXT(A1)' },
      ]),
    );

    expect(results.get('C1')?.value).toBe(6);
    expect(results.get('C2')?.value).toBe('Coco-Office-Core');
    expect(results.get('C3')?.value).toBe('12.35');
    expect(results.get('C4')?.value).toBe(12.35);
    expect(results.get('C5')?.value).toBe(18);
    expect(results.get('C6')?.value).toBe(1);
    expect(results.get('C7')?.value).toBe(45351);
    expect(results.get('C8')?.value).toBe(2024);
    expect(results.get('C9')?.value).toBe(2);
    expect(results.get('C10')?.value).toBe(true);
    expect(results.get('C11')?.value).toBe(99);
    expect(results.get('C12')?.value).toBe(true);
    expect(results.get('C13')?.value).toBe(true);
  });

  test('supports error predicates without surfacing the inspected error', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: '1/0' },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'ISERROR(A1)' },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'IFNA(MATCH("x",B1:B1,0),"not found")' },
      ]),
    );

    expect(results.get('B1')?.value).toBe(true);
    expect(results.get('C1')?.value).toBe('not found');
  });

  test('supports LET, IFS, SWITCH, and common error helpers', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 6 },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 4 },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'LET(total,A1+B1,total*2)' },
        { address: 'C2', row: 2, column: 3, kind: 'formula', value: null, formula: 'IFS(A1>10,"large",A1>5,"medium",TRUE(),"small")' },
        { address: 'C3', row: 3, column: 3, kind: 'formula', value: null, formula: 'SWITCH(A1,1,"one",6,"six","other")' },
        { address: 'C4', row: 4, column: 3, kind: 'formula', value: null, formula: 'ISNA(NA())' },
        { address: 'C5', row: 5, column: 3, kind: 'formula', value: null, formula: 'N(TRUE())+N("text")' },
      ]),
    );

    expect(results.get('C1')?.value).toBe(20);
    expect(results.get('C2')?.value).toBe('medium');
    expect(results.get('C3')?.value).toBe('six');
    expect(results.get('C4')?.value).toBe(true);
    expect(results.get('C5')?.value).toBe(1);
  });

  test('supports workday and networkdays with holidays', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: '2024-01-05' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: '2024-01-08' },
        { address: 'A3', row: 3, column: 1, kind: 'string', value: '2024-01-01' },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'WORKDAY(A1,1,A3)' },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'NETWORKDAYS(A1,A2,A3)' },
      ]),
    );

    expect(results.get('B1')?.value).toBe(45299);
    expect(results.get('B2')?.value).toBe(2);
  });

  test('supports rank and order statistics', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 10 },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 30 },
        { address: 'A3', row: 3, column: 1, kind: 'number', value: 20 },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'RANK(A3,A1:A3)' },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'RANK(A3,A1:A3,1)' },
        { address: 'B3', row: 3, column: 2, kind: 'formula', value: null, formula: 'LARGE(A1:A3,2)' },
        { address: 'B4', row: 4, column: 2, kind: 'formula', value: null, formula: 'SMALL(A1:A3,2)' },
        { address: 'B5', row: 5, column: 2, kind: 'formula', value: null, formula: 'MEDIAN(A1:A3)' },
      ]),
    );

    expect(results.get('B1')?.value).toBe(2);
    expect(results.get('B2')?.value).toBe(2);
    expect(results.get('B3')?.value).toBe(20);
    expect(results.get('B4')?.value).toBe(20);
    expect(results.get('B5')?.value).toBe(20);
  });

  test('supports date and time helpers', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: '2024-02-29' },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'DATEVALUE(A1)' },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'TIME(13,45,30)' },
        { address: 'B3', row: 3, column: 2, kind: 'formula', value: null, formula: 'HOUR(B2)' },
        { address: 'B4', row: 4, column: 2, kind: 'formula', value: null, formula: 'MINUTE(B2)' },
        { address: 'B5', row: 5, column: 2, kind: 'formula', value: null, formula: 'SECOND(B2)' },
      ]),
    );

    expect(results.get('B1')?.value).toBe(45351);
    expect(results.get('B2')?.value).toBeCloseTo(49530 / 86400, 10);
    expect(results.get('B3')?.value).toBe(13);
    expect(results.get('B4')?.value).toBe(45);
    expect(results.get('B5')?.value).toBe(30);
  });

  test('supports wildcard criteria in conditional aggregation', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'East-1' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'East-2' },
        { address: 'A3', row: 3, column: 1, kind: 'string', value: 'West-1' },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 10 },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 20 },
        { address: 'B3', row: 3, column: 2, kind: 'number', value: 30 },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'SUMIF(A1:A3,"East-*",B1:B3)' },
        { address: 'C2', row: 2, column: 3, kind: 'formula', value: null, formula: 'COUNTIF(A1:A3,"East-?")' },
      ]),
    );

    expect(results.get('C1')?.value).toBe(30);
    expect(results.get('C2')?.value).toBe(2);
  });

  test('supports SUBTOTAL and AGGREGATE hidden-row options', () => {
    const sheet = makeSheet([
      { address: 'A1', row: 1, column: 1, kind: 'number', value: 10 },
      { address: 'A2', row: 2, column: 1, kind: 'number', value: 20 },
      { address: 'A3', row: 3, column: 1, kind: 'number', value: 30 },
      { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'SUBTOTAL(9,A1:A3)' },
      { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'SUBTOTAL(109,A1:A3)' },
      { address: 'B3', row: 3, column: 2, kind: 'formula', value: null, formula: 'AGGREGATE(9,5,A1:A3)' },
      { address: 'B4', row: 4, column: 2, kind: 'formula', value: null, formula: 'AGGREGATE(14,5,A1:A3,2)' },
    ]);
    sheet.rows = [{ index: 2, hidden: true }];

    const results = calculateSpreadsheetSheetFormulas(sheet);

    expect(results.get('B1')?.value).toBe(60);
    expect(results.get('B2')?.value).toBe(40);
    expect(results.get('B3')?.value).toBe(40);
    expect(results.get('B4')?.value).toBe(10);
  });

  test('supports dynamic array functions and spill results', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 3 },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 1 },
        { address: 'A3', row: 3, column: 1, kind: 'number', value: 3 },
        { address: 'A4', row: 4, column: 1, kind: 'number', value: 2 },
        { address: 'B1', row: 1, column: 2, kind: 'boolean', value: true },
        { address: 'B2', row: 2, column: 2, kind: 'boolean', value: false },
        { address: 'B3', row: 3, column: 2, kind: 'boolean', value: true },
        { address: 'B4', row: 4, column: 2, kind: 'boolean', value: false },
        { address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: 'FILTER(A1:A4,B1:B4)' },
        { address: 'F1', row: 1, column: 6, kind: 'formula', value: null, formula: 'SORT(A1:A4)' },
        { address: 'H1', row: 1, column: 8, kind: 'formula', value: null, formula: 'UNIQUE(A1:A4)' },
        { address: 'J1', row: 1, column: 10, kind: 'formula', value: null, formula: 'SEQUENCE(2,2,10,5)' },
      ]),
    );

    expect(results.get('D1')?.value).toEqual([[3], [3]]);
    expect(results.get('D2')?.value).toBe(3);
    expect(results.get('F1')?.value).toEqual([[1], [2], [3], [3]]);
    expect(results.get('F4')?.value).toBe(3);
    expect(results.get('H1')?.value).toEqual([[3], [1], [2]]);
    expect(results.get('J1')?.value).toEqual([[10, 15], [20, 25]]);
    expect(results.get('K2')?.value).toBe(25);
  });

  test('returns SPILL when a dynamic array target is occupied', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: 'SEQUENCE(2)' },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 99 },
      ]),
    );

    expect(results.get('A1')?.error).toBe('#SPILL!');
    expect(results.get('A2')).toBeUndefined();
  });

  test('supports SORTBY, TAKE, and DROP over dynamic arrays', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'A' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'B' },
        { address: 'A3', row: 3, column: 1, kind: 'string', value: 'C' },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 30 },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 10 },
        { address: 'B3', row: 3, column: 2, kind: 'number', value: 20 },
        { address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: 'SORTBY(A1:B3,B1:B3,1)' },
        { address: 'G1', row: 1, column: 7, kind: 'formula', value: null, formula: 'TAKE(A1:B3,2)' },
        { address: 'J1', row: 1, column: 10, kind: 'formula', value: null, formula: 'DROP(A1:B3,1,1)' },
      ]),
    );

    expect(results.get('D1')?.value).toEqual([['B', 10], ['C', 20], ['A', 30]]);
    expect(results.get('D2')?.value).toEqual('C');
    expect(results.get('E3')?.value).toBe(30);
    expect(results.get('G1')?.value).toEqual([['A', 30], ['B', 10]]);
    expect(results.get('H2')?.value).toBe(10);
    expect(results.get('J1')?.value).toEqual([[10], [20]]);
    expect(results.get('J2')?.value).toBe(20);
  });

  test('recalculates dynamic spill results after a source change', () => {
    const sheet = makeSheet([
      { address: 'A1', row: 1, column: 1, kind: 'number', value: 10 },
      { address: 'A2', row: 2, column: 1, kind: 'number', value: 20 },
      { address: 'B1', row: 1, column: 2, kind: 'boolean', value: true },
      { address: 'B2', row: 2, column: 2, kind: 'boolean', value: false },
      { address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: 'FILTER(A1:A2,B1:B2)' },
    ]);
    const engine = new SpreadsheetFormulaEngine(sheet);

    engine.calculate();
    const updated = engine.updateCellValue('B2', true);

    expect(updated.get('D1')?.value).toEqual([[10], [20]]);
    expect(updated.get('D2')?.value).toBe(20);
  });

  test('supports dynamic array reshaping helpers', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'a' },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 1 },
        { address: 'C1', row: 1, column: 3, kind: 'string', value: '' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'b' },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 2 },
        { address: 'C2', row: 2, column: 3, kind: 'string', value: 'x' },
        { address: 'E1', row: 1, column: 5, kind: 'formula', value: null, formula: 'CHOOSECOLS(A1:C2,3,1)' },
        { address: 'H1', row: 1, column: 8, kind: 'formula', value: null, formula: 'CHOOSEROWS(A1:C2,-1,1)' },
        { address: 'L1', row: 1, column: 12, kind: 'formula', value: null, formula: 'TOCOL(A1:C2,1)' },
        { address: 'N1', row: 1, column: 14, kind: 'formula', value: null, formula: 'TOROW(A1:C2,1)' },
      ]),
    );

    expect(results.get('E1')?.value).toEqual([['', 'a'], ['x', 'b']]);
    expect(results.get('H1')?.value).toEqual([['b', 2, 'x'], ['a', 1, '']]);
    expect(results.get('L1')?.value).toEqual([['a'], [1], ['b'], [2], ['x']]);
    expect(results.get('N1')?.value).toEqual([['a', 1, 'b', 2, 'x']]);
  });

  test('supports DATEDIF month and day units and rejects reversed dates', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: '2020-01-31' },
        { address: 'B1', row: 1, column: 2, kind: 'string', value: '2021-03-01' },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'DATEDIF(A1,B1,"M")' },
        { address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: 'DATEDIF(A1,B1,"D")' },
        { address: 'E1', row: 1, column: 5, kind: 'formula', value: null, formula: 'DATEDIF(B1,A1,"Y")' },
      ]),
    );

    expect(results.get('C1')?.value).toBe(13);
    expect(results.get('D1')?.value).toBe(395);
    expect(results.get('E1')?.error).toBe('#NUM!');
  });

  test('supports conditional aggregation', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'East' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'West' },
        { address: 'A3', row: 3, column: 1, kind: 'string', value: 'East' },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 10 },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 20 },
        { address: 'B3', row: 3, column: 2, kind: 'number', value: 30 },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'SUMIF(A1:A3,"East",B1:B3)' },
        { address: 'C2', row: 2, column: 3, kind: 'formula', value: null, formula: 'COUNTIF(B1:B3,">=20")' },
        { address: 'C3', row: 3, column: 3, kind: 'formula', value: null, formula: 'AVERAGEIFS(B1:B3,A1:A3,"East",B1:B3,">10")' },
      ]),
    );

    expect(results.get('C1')?.value).toBe(40);
    expect(results.get('C2')?.value).toBe(2);
    expect(results.get('C3')?.value).toBe(30);
  });

  test('supports conditional extrema and common math helpers', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'East' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'West' },
        { address: 'A3', row: 3, column: 1, kind: 'string', value: 'East' },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 10 },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 20 },
        { address: 'B3', row: 3, column: 2, kind: 'number', value: 30 },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'MAXIFS(B1:B3,A1:A3,"East")' },
        { address: 'C2', row: 2, column: 3, kind: 'formula', value: null, formula: 'MINIFS(B1:B3,A1:A3,"West")' },
        { address: 'C3', row: 3, column: 3, kind: 'formula', value: null, formula: 'PRODUCT(B1:B3)' },
        { address: 'C4', row: 4, column: 3, kind: 'formula', value: null, formula: 'SUMSQ(B1:B3)' },
        { address: 'C5', row: 5, column: 3, kind: 'formula', value: null, formula: 'POWER(2,3)+SQRT(81)' },
        { address: 'C6', row: 6, column: 3, kind: 'formula', value: null, formula: 'SIGN(-4)+EXP(0)+LN(1)+LOG(100)' },
      ]),
    );

    expect(results.get('C1')?.value).toBe(30);
    expect(results.get('C2')?.value).toBe(20);
    expect(results.get('C3')?.value).toBe(6000);
    expect(results.get('C4')?.value).toBe(1400);
    expect(results.get('C5')?.value).toBe(17);
    expect(results.get('C6')?.value).toBe(2);
  });

  test('supports common financial functions', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: 'PMT(0.1/12,12,1000)' },
        { address: 'A2', row: 2, column: 1, kind: 'formula', value: null, formula: 'FV(0.1/12,12,-100)' },
        { address: 'A3', row: 3, column: 1, kind: 'formula', value: null, formula: 'PV(0.1/12,12,-100)' },
        { address: 'A4', row: 4, column: 1, kind: 'formula', value: null, formula: 'NPV(0.1,100,200,300)' },
        { address: 'A5', row: 5, column: 1, kind: 'formula', value: null, formula: 'PMT(0,10,100)' },
        { address: 'A6', row: 6, column: 1, kind: 'formula', value: null, formula: 'FV(0,10,-100)' },
        { address: 'A7', row: 7, column: 1, kind: 'formula', value: null, formula: 'PV(0,10,-100)' },
      ]),
    );

    expect(results.get('A1')?.value).toBeCloseTo(-87.915887, 5);
    expect(results.get('A2')?.value).toBeCloseTo(1256.5568, 3);
    expect(results.get('A3')?.value).toBeCloseTo(1137.4508, 3);
    expect(results.get('A4')?.value).toBeCloseTo(481.5928, 3);
    expect(results.get('A5')?.value).toBe(-10);
    expect(results.get('A6')?.value).toBe(1000);
    expect(results.get('A7')?.value).toBe(1000);
  });

  test('supports IRR, RATE, variance, and standard deviation functions', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'B1', row: 1, column: 2, kind: 'number', value: -100 },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 60 },
        { address: 'B3', row: 3, column: 2, kind: 'number', value: 60 },
        { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: 'IRR(B1:B3)' },
        { address: 'A2', row: 2, column: 1, kind: 'formula', value: null, formula: 'RATE(10,-100,1000)' },
        { address: 'A3', row: 3, column: 1, kind: 'formula', value: null, formula: 'VAR.P(1,2,3,4)' },
        { address: 'A4', row: 4, column: 1, kind: 'formula', value: null, formula: 'VAR.S(1,2,3,4)' },
        { address: 'A5', row: 5, column: 1, kind: 'formula', value: null, formula: 'STDEV.P(1,2,3,4)' },
        { address: 'A6', row: 6, column: 1, kind: 'formula', value: null, formula: 'STDEV.S(1,2,3,4)' },
      ]),
    );

    expect(results.get('A1')?.value).toBeCloseTo(0.1306623863, 8);
    expect(results.get('A2')?.value).toBeCloseTo(0, 8);
    expect(results.get('A3')?.value).toBe(1.25);
    expect(results.get('A4')?.value).toBeCloseTo(1.666666667, 8);
    expect(results.get('A5')?.value).toBeCloseTo(1.118033989, 8);
    expect(results.get('A6')?.value).toBeCloseTo(1.290994449, 8);
  });

  test('supports percentile, quartile, and statistical aliases', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 1 },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 2 },
        { address: 'B3', row: 3, column: 2, kind: 'number', value: 3 },
        { address: 'B4', row: 4, column: 2, kind: 'number', value: 4 },
        { address: 'B5', row: 5, column: 2, kind: 'number', value: 3 },
        { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: 'PERCENTILE.INC(B1:B4,0.25)' },
        { address: 'A2', row: 2, column: 1, kind: 'formula', value: null, formula: 'PERCENTILE.EXC(B1:B4,0.25)' },
        { address: 'A3', row: 3, column: 1, kind: 'formula', value: null, formula: 'QUARTILE.INC(B1:B4,3)' },
        { address: 'A4', row: 4, column: 1, kind: 'formula', value: null, formula: 'QUARTILE.EXC(B1:B4,2)' },
        { address: 'A5', row: 5, column: 1, kind: 'formula', value: null, formula: 'RANK.EQ(3,B1:B5)' },
        { address: 'A6', row: 6, column: 1, kind: 'formula', value: null, formula: 'RANK.AVG(3,B1:B5)' },
        { address: 'A7', row: 7, column: 1, kind: 'formula', value: null, formula: 'VAR(B1:B4)' },
        { address: 'A8', row: 8, column: 1, kind: 'formula', value: null, formula: 'STDEV(B1:B4)' },
      ]),
    );

    expect(results.get('A1')?.value).toBe(1.75);
    expect(results.get('A2')?.value).toBe(1.25);
    expect(results.get('A3')?.value).toBe(3.25);
    expect(results.get('A4')?.value).toBe(2.5);
    expect(results.get('A5')?.value).toBe(2);
    expect(results.get('A6')?.value).toBe(2.5);
    expect(results.get('A7')?.value).toBeCloseTo(1.666666667, 8);
    expect(results.get('A8')?.value).toBeCloseTo(1.290994449, 8);
  });

  test('supports SUMPRODUCT with quoted cross-sheet ranges', () => {
    const orders = makeSheet([
      { address: 'A2', row: 2, column: 1, kind: 'string', value: 'SO001' },
      { address: 'B2', row: 2, column: 2, kind: 'string', value: 'C001' },
      { address: 'I2', row: 2, column: 9, kind: 'formula', value: null, formula: "SUMPRODUCT(('收款单'!B$2:B$4=A2)*('收款单'!D$2:D$4))" },
      { address: 'A3', row: 3, column: 1, kind: 'string', value: 'SO002' },
      { address: 'I3', row: 3, column: 9, kind: 'formula', value: null, formula: "SUMPRODUCT(('收款单'!B$2:B$4=A3)*('收款单'!D$2:D$4))" },
    ]);
    const receipts = makeSheet([
      { address: 'B2', row: 2, column: 2, kind: 'string', value: 'SO001' },
      { address: 'D2', row: 2, column: 4, kind: 'number', value: 50000 },
      { address: 'B3', row: 3, column: 2, kind: 'string', value: 'SO002' },
      { address: 'D3', row: 3, column: 4, kind: 'number', value: 90000 },
      { address: 'B4', row: 4, column: 2, kind: 'string', value: 'SO001' },
      { address: 'D4', row: 4, column: 4, kind: 'number', value: 100000 },
    ]);
    orders.name = '销售订单';
    receipts.name = '收款单';
    const results = calculateSpreadsheetSheetFormulas(orders, { sheets: [orders, receipts] });
    expect(results.get('I2')).toMatchObject({ value: 150000, error: null });
    expect(results.get('I3')).toMatchObject({ value: 90000, error: null });
  });

  test('regresses the full sales-order and employee formula scenario', () => {
    const sales = makeSheet([
      ...Array.from({ length: 11 }, (_, index) => {
        const row = index + 2;
        const names = ['张三', '李四', '王五', '老钱', '赵高', '张三', '王五', '李四', '张三', '赵高', '孙海洋'];
        const amounts = [150000, 89000, 215000, 75000, 132000, 92000, 125000, 180000, 95000, 160000, 110000];
        const ids = ['SO001', 'SO002', 'SO003', 'SO004', 'SO005', 'SO006', 'SO007', 'SO008', 'SO009', 'SO010', 'SO011'];
        return [
          { address: `A${row}`, row, column: 1, kind: 'string' as const, value: ids[index]! },
          { address: `D${row}`, row, column: 4, kind: 'number' as const, value: amounts[index]! },
          { address: `G${row}`, row, column: 7, kind: 'string' as const, value: names[index]! },
          { address: `I${row}`, row, column: 9, kind: 'formula' as const, value: null, formula: `SUMPRODUCT(('收款单'!B$2:B$7=A${row})*('收款单'!D$2:D$7))` },
          { address: `J${row}`, row, column: 10, kind: 'formula' as const, value: null, formula: `IF(I${row}>=D${row},"已结清",IF(I${row}>0,"部分收款","未收款"))` },
        ];
      }).flat(),
    ]);
    sales.name = '销售订单';

    const receipts = makeSheet([
      { address: 'B2', row: 2, column: 2, kind: 'string', value: 'SO001' },
      { address: 'D2', row: 2, column: 4, kind: 'number', value: 50000 },
      { address: 'B3', row: 3, column: 2, kind: 'string', value: 'SO002' },
      { address: 'D3', row: 3, column: 4, kind: 'number', value: 89000 },
      { address: 'B4', row: 4, column: 2, kind: 'string', value: 'SO001' },
      { address: 'D4', row: 4, column: 4, kind: 'number', value: 100000 },
      { address: 'B5', row: 5, column: 2, kind: 'string', value: 'SO003' },
      { address: 'D5', row: 5, column: 4, kind: 'number', value: 100000 },
      { address: 'B6', row: 6, column: 2, kind: 'string', value: 'SO005' },
      { address: 'D6', row: 6, column: 4, kind: 'number', value: 80000 },
      { address: 'B7', row: 7, column: 2, kind: 'string', value: 'SO003' },
      { address: 'D7', row: 7, column: 4, kind: 'number', value: 115000 },
    ]);
    receipts.name = '收款单';

    const employees = makeSheet([
      ...['张三', '李四', '王五', '老钱', '赵高'].flatMap((name, index) => {
        const row = index + 2;
        return [
          { address: `C${row}`, row, column: 3, kind: 'string' as const, value: name },
          { address: `K${row}`, row, column: 11, kind: 'formula' as const, value: null, formula: `SUMPRODUCT(('销售订单'!G$2:G$12=C${row})*('销售订单'!D$2:D$12))` },
          { address: `L${row}`, row, column: 12, kind: 'formula' as const, value: null, formula: `IF(K${row}<=100000,K${row}*0.05,IF(K${row}<=300000,K${row}*0.08,K${row}*0.12))` },
        ];
      }),
    ]);
    employees.name = '员工';

    const sheets = [sales, receipts, employees];
    const salesResults = calculateSpreadsheetSheetFormulas(sales, { sheets });
    const employeeResults = calculateSpreadsheetSheetFormulas(employees, { sheets });
    const salesFormulaResults = [...salesResults.entries()].filter(([address]) => /^I|^J/.test(address));
    const employeeFormulaResults = [...employeeResults.entries()].filter(([address]) => /^K|^L/.test(address));

    expect(salesFormulaResults).toHaveLength(22);
    expect(employeeFormulaResults).toHaveLength(10);
    expect(salesFormulaResults.every(([, result]) => result.error === null)).toBe(true);
    expect(employeeFormulaResults.every(([, result]) => result.error === null)).toBe(true);
    expect(['I2', 'I3', 'I4', 'I5', 'I6'].map((address) => salesResults.get(address)?.value)).toEqual([150000, 89000, 215000, 0, 80000]);
    expect(['J2', 'J5', 'J6'].map((address) => salesResults.get(address)?.value)).toEqual(['已结清', '未收款', '部分收款']);
    expect(['K2', 'K3', 'K4', 'K5', 'K6'].map((address) => employeeResults.get(address)?.value)).toEqual([337000, 269000, 340000, 75000, 292000]);
    expect(['L2', 'L3', 'L4', 'L5', 'L6'].map((address) => employeeResults.get(address)?.value)).toEqual([40440, 21520, 40800, 3750, 23360]);
  });

  test('evaluates formula cells referenced from another sheet', () => {
    const orders = makeSheet([
      { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: 'Receipts!A1*2' },
    ]);
    const receipts = makeSheet([
      { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: '100+50' },
    ]);
    orders.name = '销售订单';
    receipts.name = 'Receipts';

    const results = calculateSpreadsheetSheetFormulas(orders, { sheets: [orders, receipts] });

    expect(results.get('A1')).toMatchObject({ value: 300, error: null });
  });

  test('supports exact lookups and index match', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'SKU' },
        { address: 'B1', row: 1, column: 2, kind: 'string', value: 'Price' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'A-1' },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 12 },
        { address: 'A3', row: 3, column: 1, kind: 'string', value: 'B-2' },
        { address: 'B3', row: 3, column: 2, kind: 'number', value: 18 },
        { address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: 'VLOOKUP("B-2",A2:B3,2,0)' },
        { address: 'D2', row: 2, column: 4, kind: 'formula', value: null, formula: 'INDEX(B2:B3,MATCH("A-1",A2:A3,0),1)' },
      ]),
    );

    expect(results.get('D1')?.value).toBe(18);
    expect(results.get('D2')?.value).toBe(12);
  });

  test('returns formula errors', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 1 },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 0 },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'A1/B1' },
      ]),
    );

    expect(results.get('C1')).toMatchObject({
      displayText: '#DIV/0!',
      error: '#DIV/0!',
    });
  });

  test('detects direct cycles', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: 'A1+1' },
      ]),
    );

    expect(results.get('A1')?.error).toBe('#CYCLE!');
  });

  test('short-circuits IF branches with errors', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 1 },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 0 },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'IF(TRUE(),"safe",A1/A2)' },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'IF(FALSE(),A1/A2,"safe")' },
      ]),
    );

    expect(results.get('B1')).toMatchObject({ value: 'safe', error: null });
    expect(results.get('B2')).toMatchObject({ value: 'safe', error: null });
  });

  test('propagates nested formula errors through dependent cells', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 10 },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 0 },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'A1/A2' },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'B1+1' },
      ]),
    );

    expect(results.get('B1')?.error).toBe('#DIV/0!');
    expect(results.get('C1')?.error).toBe('#DIV/0!');
  });

  test('detects indirect cycles', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'formula', value: null, formula: 'B1+1' },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'C1+1' },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'A1+1' },
      ]),
    );

    expect(results.get('A1')?.error).toBe('#CYCLE!');
    expect(results.get('B1')?.error).toBe('#CYCLE!');
    expect(results.get('C1')?.error).toBe('#CYCLE!');
  });

  test('returns errors for unsupported lookup modes and missing values', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'A' },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 1 },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'B' },
        { address: 'B2', row: 2, column: 2, kind: 'number', value: 2 },
        { address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: 'VLOOKUP("C",A1:B2,2,0)' },
        { address: 'D2', row: 2, column: 4, kind: 'formula', value: null, formula: 'MATCH("C",A1:A2,0)' },
        { address: 'D3', row: 3, column: 4, kind: 'formula', value: null, formula: 'VLOOKUP("A",A1:B2,2)' },
      ]),
    );

    expect(results.get('D1')?.error).toBe('#N/A');
    expect(results.get('D2')?.error).toBe('#N/A');
    expect(results.get('D3')?.error).toBe('#N/A');
  });

  test('returns value errors for invalid argument shapes', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'string', value: 'East' },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: 'West' },
        { address: 'B1', row: 1, column: 2, kind: 'number', value: 10 },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'SUMIF(A1:A2,"East",B1:B1)' },
        { address: 'C2', row: 2, column: 3, kind: 'formula', value: null, formula: 'INDEX(B1:B1,2,1)' },
        { address: 'C3', row: 3, column: 3, kind: 'formula', value: null, formula: 'LEFT("abc",-1)' },
      ]),
    );

    expect(results.get('C1')?.error).toBe('#VALUE!');
    expect(results.get('C2')?.error).toBe('#REF!');
    expect(results.get('C3')?.error).toBe('#VALUE!');
  });

  test('handles mixed scalar and range values in text and numeric functions', () => {
    const results = calculateSpreadsheetSheetFormulas(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 1 },
        { address: 'A2', row: 2, column: 1, kind: 'string', value: '2' },
        { address: 'A3', row: 3, column: 1, kind: 'string', value: 'x' },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'SUM(A1:A3,3)' },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'COUNT(A1:A3)' },
        { address: 'B3', row: 3, column: 2, kind: 'formula', value: null, formula: 'CONCAT("v",A1:A3)' },
      ]),
    );

    expect(results.get('B1')?.value).toBe(6);
    expect(results.get('B2')?.value).toBe(2);
    expect(results.get('B3')?.value).toBe('v12x');
  });

  test('tracks direct, range, and transitive formula dependents', () => {
    const engine = new SpreadsheetFormulaEngine(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 1 },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 2 },
        { address: 'A3', row: 3, column: 1, kind: 'number', value: 3 },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'A1+1' },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'SUM(A1:A3)' },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'B1+B2' },
      ]),
    );

    engine.calculate();

    expect(engine.getDependencies('C1')).toEqual([
      { kind: 'cell', address: 'B1' },
      { kind: 'cell', address: 'B2' },
    ]);
    expect(engine.getDependents('A1')).toEqual(['B1', 'B2']);
    expect(engine.getAffectedFormulaAddresses(['A1'])).toEqual(['B1', 'B2', 'C1']);
  });

  test('recalculates affected formula subset after changed addresses', () => {
    const engine = new SpreadsheetFormulaEngine(
      makeSheet([
        { address: 'A1', row: 1, column: 1, kind: 'number', value: 4 },
        { address: 'A2', row: 2, column: 1, kind: 'number', value: 8 },
        { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'A1*2' },
        { address: 'B2', row: 2, column: 2, kind: 'formula', value: null, formula: 'A2*2' },
        { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'B1+1' },
      ]),
    );

    engine.calculate();
    const recalculated = engine.recalculate(['A1']);

    expect([...recalculated.keys()]).toEqual(['B1', 'C1']);
    expect(recalculated.get('B1')?.value).toBe(8);
    expect(recalculated.get('C1')?.value).toBe(9);
  });

  test('updates a cell value and recalculates dependent formulas', () => {
    const sheet = makeSheet([
      { address: 'A1', row: 1, column: 1, kind: 'number', value: 4 },
      { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'A1*2' },
      { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'B1+1' },
      { address: 'D1', row: 1, column: 4, kind: 'formula', value: null, formula: '10+1' },
    ]);
    const engine = new SpreadsheetFormulaEngine(sheet);

    const updated = engine.updateCellValue('A1', 7);

    const updatedCell = sheet.cells.find((cell) => cell.address === 'A1');
    expect(updatedCell).toMatchObject({
      kind: 'number',
      value: 7,
    });
    expect(updatedCell).not.toHaveProperty('formula');
    expect([...updated.keys()]).toEqual(['B1', 'C1']);
    expect(updated.get('B1')?.value).toBe(14);
    expect(updated.get('C1')?.value).toBe(15);
  });

  test('updates a formula and recalculates downstream formulas', () => {
    const sheet = makeSheet([
      { address: 'A1', row: 1, column: 1, kind: 'number', value: 3 },
      { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'A1*2' },
      { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'B1+1' },
    ]);
    const engine = new SpreadsheetFormulaEngine(sheet);

    const updated = engine.updateCellFormula('B1', '=A1*3');

    expect(sheet.cells.find((cell) => cell.address === 'B1')).toMatchObject({
      kind: 'formula',
      formula: 'A1*3',
      value: null,
    });
    expect([...updated.keys()]).toEqual(['B1', 'C1']);
    expect(updated.get('B1')?.value).toBe(9);
    expect(updated.get('C1')?.value).toBe(10);
  });

  test('creates new cells when updating missing addresses', () => {
    const sheet = makeSheet([
      { address: 'A1', row: 1, column: 1, kind: 'number', value: 3 },
      { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'B1+1' },
    ]);
    const engine = new SpreadsheetFormulaEngine(sheet);

    const updated = engine.updateCellFormula('B1', 'A1*4');

    expect(sheet.cells.find((cell) => cell.address === 'B1')).toMatchObject({
      row: 1,
      column: 2,
      kind: 'formula',
      formula: 'A1*4',
    });
    expect([...updated.keys()]).toEqual(['B1', 'C1']);
    expect(updated.get('B1')?.value).toBe(12);
    expect(updated.get('C1')?.value).toBe(13);
  });

  test('clears a formula cell and recalculates downstream formulas against blank value', () => {
    const sheet = makeSheet([
      { address: 'A1', row: 1, column: 1, kind: 'number', value: 3 },
      { address: 'B1', row: 1, column: 2, kind: 'formula', value: null, formula: 'A1*4' },
      { address: 'C1', row: 1, column: 3, kind: 'formula', value: null, formula: 'B1+1' },
    ]);
    const engine = new SpreadsheetFormulaEngine(sheet);

    const updated = engine.updateCellFormula('B1', null);

    expect(sheet.cells.some((cell) => cell.address === 'B1')).toBe(false);
    expect([...updated.keys()]).toEqual(['C1']);
    expect(updated.get('C1')?.value).toBe(1);
  });
});

function makeSheet(cells: ExcelSheetModel['cells']): ExcelSheetModel {
  return {
    id: '1',
    name: 'Sheet1',
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
    cells,
  };
}
