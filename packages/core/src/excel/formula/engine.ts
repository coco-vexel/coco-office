import type { ExcelCellModel, ExcelCellValue, ExcelSheetModel } from '../model.js';
import { evaluateSpreadsheetFormulaAst } from './evaluator.js';
import { parseSpreadsheetFormula } from './parser.js';
import {
  formulaError,
  formulaValue,
  type SpreadsheetFormulaReference,
  type SpreadsheetFormulaError,
  type SpreadsheetFormulaResult,
} from './types.js';

export interface SpreadsheetFormulaEngineOptions {
  /** Evaluation date used by TODAY(). Defaults to the host's current local date. */
  today?: Date;
  /** Workbook sheets used to resolve quoted cross-sheet references. */
  sheets?: readonly ExcelSheetModel[];
}

class SpreadsheetFormulaEvaluationError extends Error {
  constructor(
    readonly formulaError: SpreadsheetFormulaError,
    readonly dependencies: SpreadsheetFormulaReference[] = [],
  ) {
    super(formulaError);
  }
}

export class SpreadsheetFormulaEngine {
  private readonly cellsByAddress: Map<string, ExcelCellModel>;
  private readonly resultsByAddress = new Map<string, SpreadsheetFormulaResult>();
  private readonly resultsBySheetAddress = new Map<string, SpreadsheetFormulaResult>();
  private readonly dependenciesByAddress = new Map<string, SpreadsheetFormulaReference[]>();
  private readonly dependentsByAddress = new Map<string, Set<string>>();
  private readonly evaluating = new Set<string>();
  private readonly cellsBySheet = new Map<string, Map<string, ExcelCellModel>>();
  private readonly sheetsByName = new Map<string, ExcelSheetModel>();

  constructor(
    private readonly sheet: ExcelSheetModel,
    private readonly options: SpreadsheetFormulaEngineOptions = {},
  ) {
    this.cellsByAddress = new Map(sheet.cells.map((cell) => [cell.address, cell]));
    const sheets = options.sheets ?? [sheet];
    for (const item of sheets) {
      this.sheetsByName.set(item.name, item);
      this.cellsBySheet.set(item.name, new Map(item.cells.map((cell) => [cell.address, cell])));
    }
  }

  calculate(): Map<string, SpreadsheetFormulaResult> {
    for (const cell of this.sheet.cells) {
      if (cell.formula) this.evaluateFormulaCell(cell.address);
    }
    return this.withSpillResults(new Map(this.resultsByAddress));
  }

  recalculate(changedAddresses: Iterable<string>): Map<string, SpreadsheetFormulaResult> {
    this.calculate();
    const affected = this.getAffectedFormulaAddresses(changedAddresses);
    for (const address of affected) {
      this.resultsByAddress.delete(address);
      this.resultsBySheetAddress.delete(sheetAddressKey(this.sheet.name, address));
      this.dependenciesByAddress.delete(address);
    }
    this.rebuildDependents();
    for (const address of affected) {
      this.evaluateFormulaCell(address);
    }
    return this.withSpillResults(new Map(
      [...affected]
        .map((address) => [address, this.resultsByAddress.get(address)] as const)
        .filter((entry): entry is readonly [string, SpreadsheetFormulaResult] => entry[1] !== undefined),
    ));
  }

  updateCellValue(address: string, value: ExcelCellValue): Map<string, SpreadsheetFormulaResult> {
    const normalized = addressFromKey(normalizeAddressKey(address));
    const affected = new Set([
      ...this.getAffectedFormulaAddresses([normalized]),
      normalized,
    ]);
    this.setCellValue(normalized, value);
    return this.recalculateKnownAffected(affected);
  }

  updateCellFormula(address: string, formula: string | null): Map<string, SpreadsheetFormulaResult> {
    const normalized = addressFromKey(normalizeAddressKey(address));
    const affected = new Set([
      ...this.getAffectedFormulaAddresses([normalized]),
      normalized,
    ]);
    this.setCellFormula(normalized, formula);
    return this.recalculateKnownAffected(affected);
  }

  evaluateFormulaCell(address: string): SpreadsheetFormulaResult {
    return this.evaluateFormulaCellAt(this.sheet, address);
  }

  private evaluateFormulaCellAt(targetSheet: ExcelSheetModel, address: string): SpreadsheetFormulaResult {
    const cacheKey = sheetAddressKey(targetSheet.name, address);
    const cached = this.resultsBySheetAddress.get(cacheKey);
    if (cached) return cached;
    const cell = (targetSheet === this.sheet ? this.cellsByAddress : this.cellsBySheet.get(targetSheet.name))?.get(address);
    if (!cell?.formula) return formulaValue(cell?.value ?? null);
    if (this.evaluating.has(cacheKey)) {
      const result = formulaError('#CYCLE!');
      this.resultsBySheetAddress.set(cacheKey, result);
      return result;
    }
    this.evaluating.add(cacheKey);
    let result: SpreadsheetFormulaResult;
    try {
      const ast = parseSpreadsheetFormula(cell.formula);
      const evaluated = evaluateSpreadsheetFormulaAst(ast, {
        getToday: () => excelSerialFromDate(this.options.today ?? new Date()),
        isRowHidden: (row, sheetName) => {
          const referencedSheet = sheetName ? this.sheetsByName.get(sheetName) : targetSheet;
          return referencedSheet?.rows.find((item) => item.index === row)?.hidden === true;
        },
        getCellEvaluation: (cellAddress, sheetName) => {
          const referencedSheet = sheetName ? this.sheetsByName.get(sheetName) : targetSheet;
          const referenced = (sheetName ? this.cellsBySheet.get(sheetName) : targetSheet === this.sheet ? this.cellsByAddress : this.cellsBySheet.get(targetSheet.name))?.get(cellAddress);
          if (!referenced) return formulaValue(null);
          if (!referenced.formula) return formulaValue(referenced.value);
          return this.evaluateFormulaCellAt(referencedSheet ?? targetSheet, cellAddress);
        },
        getCellValue: (cellAddress, sheetName) => {
          const referencedSheet = sheetName ? this.sheetsByName.get(sheetName) : targetSheet;
          const referenced = (sheetName ? this.cellsBySheet.get(sheetName) : targetSheet === this.sheet ? this.cellsByAddress : this.cellsBySheet.get(targetSheet.name))?.get(cellAddress);
          if (!referenced) return null;
          if (!referenced.formula) return referenced.value;
          const nested = this.evaluateFormulaCellAt(referencedSheet ?? targetSheet, cellAddress);
          if (nested.error) {
            throw new SpreadsheetFormulaEvaluationError(nested.error, [
              { kind: 'cell', address: cellAddress },
              ...nested.dependencies,
            ]);
          }
          return Array.isArray(nested.value) ? nested.value[0]?.[0] ?? null : nested.value;
        },
      });
      result = evaluated.error
        ? formulaError(evaluated.error, evaluated.dependencies)
        : formulaValue(evaluated.value, evaluated.dependencies);
    } catch (error) {
      result =
        error instanceof SpreadsheetFormulaEvaluationError
          ? formulaError(error.formulaError, error.dependencies)
          : formulaError('#VALUE!');
    } finally {
      this.evaluating.delete(cacheKey);
    }
    this.resultsBySheetAddress.set(cacheKey, result);
    if (targetSheet === this.sheet) {
      this.resultsByAddress.set(address, result);
      this.registerDependencies(address, result.dependencies);
    }
    return result;
  }

  getDependencies(address: string): SpreadsheetFormulaReference[] {
    if (!this.dependenciesByAddress.has(address)) this.evaluateFormulaCell(address);
    return [...(this.dependenciesByAddress.get(address) ?? [])];
  }

  getDependents(address: string): string[] {
    this.calculate();
    return [...(this.dependentsByAddress.get(normalizeAddressKey(address)) ?? [])].sort(compareAddresses);
  }

  getAffectedFormulaAddresses(changedAddresses: Iterable<string>): string[] {
    this.calculate();
    const affected = new Set<string>();
    const queue = [...changedAddresses].map(normalizeAddressKey);
    for (let index = 0; index < queue.length; index++) {
      const address = queue[index]!;
      const dependents = this.dependentsByAddress.get(address);
      if (!dependents) continue;
      for (const dependent of dependents) {
        if (affected.has(dependent)) continue;
        affected.add(dependent);
        queue.push(normalizeAddressKey(dependent));
      }
    }
    return this.sortAffectedFormulaAddresses(affected);
  }

  private registerDependencies(
    formulaAddress: string,
    dependencies: SpreadsheetFormulaReference[],
  ): void {
    this.dependenciesByAddress.set(formulaAddress, dependencies);
    this.addDependentEdges(formulaAddress, dependencies);
  }

  private addDependentEdges(
    formulaAddress: string,
    dependencies: SpreadsheetFormulaReference[],
  ): void {
    for (const dependency of dependencies) {
      for (const address of expandDependencyAddresses(dependency)) {
        let dependents = this.dependentsByAddress.get(address);
        if (!dependents) {
          dependents = new Set<string>();
          this.dependentsByAddress.set(address, dependents);
        }
        dependents.add(formulaAddress);
      }
    }
  }

  private rebuildDependents(): void {
    this.dependentsByAddress.clear();
    for (const [address, dependencies] of this.dependenciesByAddress) {
      this.addDependentEdges(address, dependencies);
    }
  }

  private recalculateKnownAffected(
    affected: Set<string>,
  ): Map<string, SpreadsheetFormulaResult> {
    this.clearCalculationState();
    this.calculate();
    const sorted = this.sortAffectedFormulaAddresses(
      new Set([...affected].filter((address) => this.cellsByAddress.get(address)?.formula)),
    );
    return this.withSpillResults(new Map(
      sorted
        .map((address) => [address, this.resultsByAddress.get(address)] as const)
        .filter((entry): entry is readonly [string, SpreadsheetFormulaResult] => entry[1] !== undefined),
    ));
  }

  private clearCalculationState(): void {
    this.resultsByAddress.clear();
    this.resultsBySheetAddress.clear();
    this.dependenciesByAddress.clear();
    this.dependentsByAddress.clear();
    this.evaluating.clear();
  }

  private withSpillResults(results: Map<string, SpreadsheetFormulaResult>): Map<string, SpreadsheetFormulaResult> {
    const occupiedSpillCells = new Set<string>();
    for (const cell of this.sheet.cells) {
      if (!cell.formula) continue;
      const result = results.get(cell.address);
      if (!result || result.error || !Array.isArray(result.value)) continue;
      const array = result.value;
      const occupied = new Set(this.sheet.cells.filter((item) => item.address !== cell.address && (item.formula || item.value !== null)).map((item) => `${item.row}:${item.column}`));
      const hasConflict = array.some((row, rowOffset) => row.some((_value, columnOffset) => {
        const key = `${cell.row + rowOffset}:${cell.column + columnOffset}`;
        return occupied.has(key) || occupiedSpillCells.has(key);
      }));
      if (hasConflict) {
        results.set(cell.address, formulaError('#SPILL!', result.dependencies));
        continue;
      }
      for (let rowOffset = 0; rowOffset < array.length; rowOffset++) {
        for (let columnOffset = 0; columnOffset < (array[rowOffset]?.length ?? 0); columnOffset++) {
          if (rowOffset === 0 && columnOffset === 0) continue;
          const address = addressFromRowColumn(cell.row + rowOffset, cell.column + columnOffset);
          occupiedSpillCells.add(`${cell.row + rowOffset}:${cell.column + columnOffset}`);
          results.set(address, formulaValue(array[rowOffset]?.[columnOffset] ?? null, [{ kind: 'cell', address: cell.address }]));
        }
      }
    }
    return results;
  }

  private setCellValue(address: string, value: ExcelCellValue): void {
    if (value === null) {
      this.deleteCell(address);
      return;
    }
    const cell = this.ensureCell(address);
    cell.value = value;
    cell.kind =
      typeof value === 'boolean' ? 'boolean' : typeof value === 'number' ? 'number' : 'string';
    delete cell.formula;
    delete cell.rawValue;
    delete cell.dataType;
    delete cell.sharedStringIndex;
  }

  private setCellFormula(address: string, formula: string | null): void {
    if (formula === null) {
      this.deleteCell(address);
      return;
    }
    const cell = this.ensureCell(address);
    cell.kind = 'formula';
    cell.formula = formula.startsWith('=') ? formula.slice(1) : formula;
    cell.value = null;
    delete cell.rawValue;
    delete cell.dataType;
    delete cell.sharedStringIndex;
  }

  private ensureCell(address: string): ExcelCellModel {
    const existing = this.cellsByAddress.get(address);
    if (existing) return existing;
    const parsed = parseAddress(address);
    const cell: ExcelCellModel = {
      address,
      row: parsed.row,
      column: parsed.column,
      kind: 'blank',
      value: null,
    };
    this.sheet.cells.push(cell);
    this.sheet.cells.sort((left, right) => left.row - right.row || left.column - right.column);
    this.cellsByAddress.set(address, cell);
    return cell;
  }

  private deleteCell(address: string): void {
    const cell = this.cellsByAddress.get(address);
    if (!cell) return;
    this.cellsByAddress.delete(address);
    const index = this.sheet.cells.indexOf(cell);
    if (index !== -1) this.sheet.cells.splice(index, 1);
  }

  private sortAffectedFormulaAddresses(affected: Set<string>): string[] {
    const remaining = new Set(affected);
    const sorted: string[] = [];
    while (remaining.size > 0) {
      const ready = [...remaining]
        .filter((address) => {
          const dependencies = this.dependenciesByAddress.get(address) ?? [];
          return !dependencies.some((dependency) =>
            expandDependencyAddresses(dependency).some((dependencyAddress) =>
              remaining.has(addressFromKey(dependencyAddress)),
            ),
          );
        })
        .sort(compareAddresses);
      if (ready.length === 0) {
        sorted.push(...[...remaining].sort(compareAddresses));
        break;
      }
      for (const address of ready) {
        remaining.delete(address);
        sorted.push(address);
      }
    }
    return sorted;
  }
}

function sheetAddressKey(sheetName: string, address: string): string {
  return `${sheetName}\u0000${address}`;
}

function addressFromRowColumn(row: number, column: number): string {
  return `${columnIndexToName(column)}${row}`;
}

export function calculateSpreadsheetSheetFormulas(
  sheet: ExcelSheetModel,
  options: SpreadsheetFormulaEngineOptions = {},
): Map<string, SpreadsheetFormulaResult> {
  return new SpreadsheetFormulaEngine(sheet, options).calculate();
}

function excelSerialFromDate(date: Date): number {
  const utcMidnight = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.floor((utcMidnight - Date.UTC(1899, 11, 30)) / 86_400_000);
}

function expandDependencyAddresses(dependency: SpreadsheetFormulaReference): string[] {
  if (dependency.kind === 'cell') return [normalizeAddressKey(dependency.address)];
  const start = parseAddress(dependency.startAddress);
  const end = parseAddress(dependency.endAddress);
  const addresses: string[] = [];
  for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
    for (
      let column = Math.min(start.column, end.column);
      column <= Math.max(start.column, end.column);
      column++
    ) {
      addresses.push(`${row}:${column}`);
    }
  }
  return addresses;
}

function normalizeAddressKey(address: string): string {
  const parsed = parseAddress(address);
  return `${parsed.row}:${parsed.column}`;
}

function addressFromKey(key: string): string {
  const [row, column] = key.split(':').map(Number);
  if (!row || !column) return key;
  return `${columnIndexToName(column)}${row}`;
}

function parseAddress(address: string): { row: number; column: number } {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(address.trim());
  if (!match) return { row: 0, column: 0 };
  return {
    row: Number(match[2]),
    column: columnNameToIndex(match[1]!),
  };
}

function columnNameToIndex(columnName: string): number {
  let value = 0;
  for (const ch of columnName.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64);
  return value;
}

function columnIndexToName(columnIndex: number): string {
  let value = columnIndex;
  let name = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function compareAddresses(left: string, right: string): number {
  const leftRef = parseAddress(left);
  const rightRef = parseAddress(right);
  return leftRef.row - rightRef.row || leftRef.column - rightRef.column || left.localeCompare(right);
}

export * from './types.js';
export * from './parser.js';
export * from './evaluator.js';
