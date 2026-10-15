import type { ExcelCellValue } from '../model.js';

export type SpreadsheetFormulaError =
  | '#DIV/0!'
  | '#NUM!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#N/A'
  | '#CALC!'
  | '#SPILL!'
  | '#CYCLE!';

export type SpreadsheetFormulaScalar = ExcelCellValue | boolean;

export type SpreadsheetFormulaValue = SpreadsheetFormulaScalar | SpreadsheetFormulaScalar[][];

export interface SpreadsheetFormulaCellReference {
  kind: 'cell';
  address: string;
  sheetName?: string;
}

export interface SpreadsheetFormulaRangeReference {
  kind: 'range';
  startAddress: string;
  endAddress: string;
  sheetName?: string;
}

export type SpreadsheetFormulaReference =
  | SpreadsheetFormulaCellReference
  | SpreadsheetFormulaRangeReference;

export type SpreadsheetFormulaAst =
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }
  | { kind: 'name'; name: string }
  | { kind: 'cell'; address: string; sheetName?: string }
  | { kind: 'range'; startAddress: string; endAddress: string; sheetName?: string }
  | { kind: 'unary'; operator: '+' | '-' | '%'; argument: SpreadsheetFormulaAst }
  | {
      kind: 'binary';
      operator: '+' | '-' | '*' | '/' | '&' | '=' | '<>' | '>' | '>=' | '<' | '<=';
      left: SpreadsheetFormulaAst;
      right: SpreadsheetFormulaAst;
    }
  | { kind: 'call'; name: string; args: SpreadsheetFormulaAst[] };

export interface SpreadsheetFormulaResult {
  value: SpreadsheetFormulaValue;
  displayText: string;
  error: SpreadsheetFormulaError | null;
  dependencies: SpreadsheetFormulaReference[];
}

export function formulaError(
  error: SpreadsheetFormulaError,
  dependencies: SpreadsheetFormulaReference[] = [],
): SpreadsheetFormulaResult {
  return {
    value: null,
    displayText: error,
    error,
    dependencies,
  };
}

export function formulaValue(
  value: SpreadsheetFormulaValue,
  dependencies: SpreadsheetFormulaReference[] = [],
): SpreadsheetFormulaResult {
  const displayValue = Array.isArray(value) ? value[0]?.[0] ?? null : value;
  return {
    value,
    displayText: displayValue === null ? '' : String(displayValue),
    error: null,
    dependencies,
  };
}
