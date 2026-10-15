import { formatCellAddress, parseCellAddress } from '../address.js';
import type {
  SpreadsheetFormulaAst,
  SpreadsheetFormulaError,
  SpreadsheetFormulaReference,
  SpreadsheetFormulaScalar,
  SpreadsheetFormulaValue,
} from './types.js';

export interface SpreadsheetFormulaEvaluationContext {
  getCellValue(address: string, sheetName?: string): SpreadsheetFormulaScalar;
  getCellEvaluation?(address: string, sheetName?: string): SpreadsheetFormulaEvaluation;
  getNameValue?(name: string): SpreadsheetFormulaEvaluation;
  isRowHidden?(row: number, sheetName?: string): boolean;
  /** Returns today's date as an Excel serial number. Injectable for deterministic hosts/tests. */
  getToday?: () => number;
}

export interface SpreadsheetFormulaEvaluation {
  value: SpreadsheetFormulaValue;
  error: SpreadsheetFormulaError | null;
  dependencies: SpreadsheetFormulaReference[];
}

export function evaluateSpreadsheetFormulaAst(
  ast: SpreadsheetFormulaAst,
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  return evaluateNode(ast, context);
}

function evaluateNode(
  ast: SpreadsheetFormulaAst,
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  switch (ast.kind) {
    case 'number':
    case 'string':
      return ok(ast.value);
    case 'name':
      return context.getNameValue?.(ast.name) ?? error('#NAME?');
    case 'cell': {
      const evaluated = context.getCellEvaluation?.(ast.address, ast.sheetName);
      if (evaluated) {
        const reference: SpreadsheetFormulaReference = ast.sheetName
          ? { kind: 'cell', address: ast.address, sheetName: ast.sheetName }
          : { kind: 'cell', address: ast.address };
        return {
          value: evaluated.value,
          error: evaluated.error,
          dependencies: [reference],
        };
      }
      const value = context.getCellValue(ast.address, ast.sheetName);
      return ok(value, [{ kind: 'cell', address: ast.address, sheetName: ast.sheetName }]);
    }
    case 'range':
      return evaluateRange(ast.startAddress, ast.endAddress, ast.sheetName, context);
    case 'unary':
      return evaluateUnary(ast, context);
    case 'binary':
      return evaluateBinary(ast, context);
    case 'call':
      return evaluateCall(ast.name, ast.args, context);
  }
}

function evaluateRange(
  startAddress: string,
  endAddress: string,
  sheetName: string | undefined,
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  const start = parseCellAddress(startAddress);
  const end = parseCellAddress(endAddress);
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startColumn = Math.min(start.column, end.column);
  const endColumn = Math.max(start.column, end.column);
  const values: SpreadsheetFormulaScalar[][] = [];
  for (let row = startRow; row <= endRow; row++) {
    const rowValues: SpreadsheetFormulaScalar[] = [];
    for (let column = startColumn; column <= endColumn; column++) {
      rowValues.push(context.getCellValue(formatCellAddress(row, column), sheetName));
    }
    values.push(rowValues);
  }
  return ok(values, [{ kind: 'range', startAddress, endAddress, sheetName }]);
}

function evaluateUnary(
  ast: Extract<SpreadsheetFormulaAst, { kind: 'unary' }>,
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  const argument = evaluateNode(ast.argument, context);
  if (argument.error) return argument;
  const number = toNumber(argument.value);
  if (number === null) return error('#VALUE!', argument.dependencies);
  if (ast.operator === '%') return ok(number / 100, argument.dependencies);
  return ok(ast.operator === '-' ? -number : number, argument.dependencies);
}

function evaluateBinary(
  ast: Extract<SpreadsheetFormulaAst, { kind: 'binary' }>,
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  const left = evaluateNode(ast.left, context);
  if (left.error) return left;
  const right = evaluateNode(ast.right, context);
  const dependencies = [...left.dependencies, ...right.dependencies];
  if (right.error) return { ...right, dependencies };
  if (Array.isArray(left.value) || Array.isArray(right.value)) {
    return evaluateMatrixBinary(ast.operator, left.value, right.value, dependencies);
  }
  if (['=', '<>', '>', '>=', '<', '<='].includes(ast.operator)) {
    return ok(compareValues(left.value, right.value, ast.operator), dependencies);
  }
  if (ast.operator === '&') {
    return ok(`${toText(left.value)}${toText(right.value)}`, dependencies);
  }
  const leftNumber = toNumber(left.value);
  const rightNumber = toNumber(right.value);
  if (leftNumber === null || rightNumber === null) return error('#VALUE!', dependencies);
  if (ast.operator === '/' && rightNumber === 0) return error('#DIV/0!', dependencies);
  const value =
    ast.operator === '+'
      ? leftNumber + rightNumber
      : ast.operator === '-'
        ? leftNumber - rightNumber
        : ast.operator === '*'
          ? leftNumber * rightNumber
          : leftNumber / rightNumber;
  return ok(value, dependencies);
}

function evaluateCall(
  name: string,
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (name === 'IF') return evaluateIf(args, context);
  if (name === 'IFERROR') return evaluateIfError(args, context);
  if (name === 'IFNA') return evaluateIfNa(args, context);
  if (name === 'ISERROR' || name === 'ISERR' || name === 'ISNA') return evaluateErrorPredicate(name, args, context);
  if (name === 'LET') return evaluateLet(args, context);
  if (name === 'IFS') return evaluateIfs(args, context);
  if (name === 'SWITCH') return evaluateSwitch(args, context);
  if (name === 'SUBTOTAL') return evaluateSubtotal(args, context);
  if (name === 'AGGREGATE') return evaluateAggregate(args, context);
  const values: SpreadsheetFormulaValue[] = [];
  let dependencies: SpreadsheetFormulaReference[] = [];
  for (const arg of args) {
    const evaluated = evaluateNode(arg, context);
    dependencies = [...dependencies, ...evaluated.dependencies];
    if (evaluated.error) return { ...evaluated, dependencies };
    values.push(evaluated.value);
  }
  if (name === 'SUMPRODUCT') return evaluateSumProduct(values, dependencies);
  switch (name) {
    case 'TRUE':
      return args.length === 0 ? ok(true, dependencies) : error('#VALUE!', dependencies);
    case 'FALSE':
      return args.length === 0 ? ok(false, dependencies) : error('#VALUE!', dependencies);
    case 'NA':
      return args.length === 0 ? error('#N/A', dependencies) : error('#VALUE!', dependencies);
    case 'N': {
      const value = values[0] ?? null;
      if (typeof value === 'number') return ok(value, dependencies);
      if (typeof value === 'boolean') return ok(value ? 1 : 0, dependencies);
      return ok(0, dependencies);
    }
    case 'AND':
      return ok(values.every(truthy), dependencies);
    case 'OR':
      return ok(values.some(truthy), dependencies);
    case 'NOT':
      return values.length === 1 ? ok(!truthy(values[0]!), dependencies) : error('#VALUE!', dependencies);
    case 'SUM':
      return ok(numericValues(values).reduce((sum, value) => sum + value, 0), dependencies);
    case 'PRODUCT': {
      const numbers = numericValues(values);
      return ok(numbers.reduce((product, value) => product * value, 1), dependencies);
    }
    case 'SUMSQ':
      return ok(numericValues(values).reduce((sum, value) => sum + value * value, 0), dependencies);
    case 'AVERAGE': {
      const numbers = numericValues(values);
      return numbers.length === 0
        ? error('#DIV/0!', dependencies)
        : ok(numbers.reduce((sum, value) => sum + value, 0) / numbers.length, dependencies);
    }
    case 'COUNT':
      return ok(numericValues(values).length, dependencies);
    case 'COUNTA':
      return ok(flatten(values).filter((value) => value !== null && value !== '').length, dependencies);
    case 'MIN': {
      const numbers = numericValues(values);
      return numbers.length === 0 ? ok(0, dependencies) : ok(Math.min(...numbers), dependencies);
    }
    case 'MAX': {
      const numbers = numericValues(values);
      return numbers.length === 0 ? ok(0, dependencies) : ok(Math.max(...numbers), dependencies);
    }
    case 'MEDIAN': {
      const numbers = numericValues(values).sort((left, right) => left - right);
      if (!numbers.length) return error('#NUM!', dependencies);
      const middle = Math.floor(numbers.length / 2);
      return ok(numbers.length % 2 ? numbers[middle]! : (numbers[middle - 1]! + numbers[middle]!) / 2, dependencies);
    }
    case 'LARGE':
    case 'SMALL': {
      const numbers = numericValues([values[0] ?? null]).sort((left, right) => name === 'LARGE' ? right - left : left - right);
      const index = toNumber(values[1] ?? null);
      if (index === null || !Number.isInteger(index) || index < 1 || index > numbers.length) return error('#NUM!', dependencies);
      return ok(numbers[index - 1]!, dependencies);
    }
    case 'RANK':
    case 'RANK.EQ':
    case 'RANK.AVG': {
      const number = toNumber(values[0] ?? null);
      const reference = numericValues([values[1] ?? null]);
      const order = toNumber(values[2] ?? 0);
      if (number === null || order === null || !reference.length) return error('#N/A', dependencies);
      const rank = reference.filter((value) => order === 0 ? value > number : value < number).length + 1;
      if (name === 'RANK.AVG') {
        const tieCount = reference.filter((value) => value === number).length;
        return ok(tieCount > 1 ? rank + (tieCount - 1) / 2 : rank, dependencies);
      }
      return ok(rank, dependencies);
    }
    case 'ABS': {
      const number = toNumber(values[0] ?? null);
      return number === null ? error('#VALUE!', dependencies) : ok(Math.abs(number), dependencies);
    }
    case 'SIGN': {
      const number = toNumber(values[0] ?? null);
      return number === null ? error('#VALUE!', dependencies) : ok(number < 0 ? -1 : number > 0 ? 1 : 0, dependencies);
    }
    case 'POWER': {
      const base = toNumber(values[0] ?? null);
      const exponent = toNumber(values[1] ?? null);
      if (base === null || exponent === null) return error('#VALUE!', dependencies);
      const result = base ** exponent;
      return Number.isFinite(result) ? ok(result, dependencies) : error('#NUM!', dependencies);
    }
    case 'SQRT': {
      const number = toNumber(values[0] ?? null);
      return number === null ? error('#VALUE!', dependencies) : number < 0 ? error('#NUM!', dependencies) : ok(Math.sqrt(number), dependencies);
    }
    case 'EXP': {
      const number = toNumber(values[0] ?? null);
      if (number === null) return error('#VALUE!', dependencies);
      const result = Math.exp(number);
      return Number.isFinite(result) ? ok(result, dependencies) : error('#NUM!', dependencies);
    }
    case 'LN': {
      const number = toNumber(values[0] ?? null);
      return number === null ? error('#VALUE!', dependencies) : number <= 0 ? error('#NUM!', dependencies) : ok(Math.log(number), dependencies);
    }
    case 'LOG': {
      const number = toNumber(values[0] ?? null);
      const base = toNumber(values[1] ?? 10);
      return number === null || base === null || number <= 0 || base <= 0 || base === 1
        ? error(number === null || base === null ? '#VALUE!' : '#NUM!', dependencies)
        : ok(Math.log(number) / Math.log(base), dependencies);
    }
    case 'PMT':
    case 'FV':
    case 'PV':
      return evaluateFinancial(name, values, dependencies);
    case 'NPV': {
      const rate = toNumber(values[0] ?? null);
      const cashFlows = numericValues(values.slice(1));
      if (rate === null || rate <= -1) return error(rate === null ? '#VALUE!' : '#NUM!', dependencies);
      return ok(cashFlows.reduce((total, cashFlow, index) => total + cashFlow / (1 + rate) ** (index + 1), 0), dependencies);
    }
    case 'IRR':
      return evaluateIrr(values, dependencies);
    case 'RATE':
      return evaluateRate(values, dependencies);
    case 'VAR':
    case 'VAR.P':
    case 'VAR.S':
    case 'STDEV':
    case 'STDEV.P':
    case 'STDEV.S':
      return evaluateVariance(name, values, dependencies);
    case 'PERCENTILE':
    case 'PERCENTILE.INC':
    case 'PERCENTILE.EXC':
      return evaluatePercentile(name, values, dependencies);
    case 'QUARTILE':
    case 'QUARTILE.INC':
    case 'QUARTILE.EXC':
      return evaluateQuartile(name, values, dependencies);
    case 'ROUND': {
      const number = toNumber(values[0] ?? null);
      const digits = toNumber(values[1] ?? 0);
      if (number === null || digits === null) return error('#VALUE!', dependencies);
      const factor = 10 ** Math.trunc(digits);
      return ok(Math.round(number * factor) / factor, dependencies);
    }
    case 'ROUNDUP':
    case 'ROUNDDOWN': {
      const number = toNumber(values[0] ?? null);
      const digits = toNumber(values[1] ?? 0);
      if (number === null || digits === null) return error('#VALUE!', dependencies);
      return ok(roundDirectional(number, Math.trunc(digits), name === 'ROUNDUP'), dependencies);
    }
    case 'CEILING':
    case 'FLOOR': {
      const number = toNumber(values[0] ?? null);
      const significance = toNumber(values[1] ?? 1);
      if (number === null || significance === null || significance === 0) return error('#VALUE!', dependencies);
      const step = Math.abs(significance);
      const rounded = name === 'CEILING'
        ? Math.ceil(number / step) * step
        : Math.floor(number / step) * step;
      return ok(rounded, dependencies);
    }
    case 'INT': {
      const number = toNumber(values[0] ?? null);
      return number === null ? error('#VALUE!', dependencies) : ok(Math.floor(number), dependencies);
    }
    case 'MOD': {
      const number = toNumber(values[0] ?? null);
      const divisor = toNumber(values[1] ?? null);
      if (number === null || divisor === null) return error('#VALUE!', dependencies);
      return divisor === 0 ? error('#DIV/0!', dependencies) : ok(number % divisor, dependencies);
    }
    case 'LEN':
      return values.length === 1 ? ok(toText(values[0]!).length, dependencies) : error('#VALUE!', dependencies);
    case 'TRIM':
      return values.length === 1 ? ok(toText(values[0]!).trim().replace(/\s+/g, ' '), dependencies) : error('#VALUE!', dependencies);
    case 'LOWER':
      return values.length === 1 ? ok(toText(values[0]!).toLowerCase(), dependencies) : error('#VALUE!', dependencies);
    case 'UPPER':
      return values.length === 1 ? ok(toText(values[0]!).toUpperCase(), dependencies) : error('#VALUE!', dependencies);
    case 'LEFT': {
      const text = toText(values[0] ?? '');
      const count = toNumber(values[1] ?? 1);
      return count === null || count < 0 ? error('#VALUE!', dependencies) : ok(text.slice(0, Math.trunc(count)), dependencies);
    }
    case 'RIGHT': {
      const text = toText(values[0] ?? '');
      const count = toNumber(values[1] ?? 1);
      return count === null || count < 0 ? error('#VALUE!', dependencies) : ok(text.slice(Math.max(0, text.length - Math.trunc(count))), dependencies);
    }
    case 'MID': {
      const text = toText(values[0] ?? '');
      const start = toNumber(values[1] ?? null);
      const count = toNumber(values[2] ?? null);
      if (start === null || count === null || start < 1 || count < 0) return error('#VALUE!', dependencies);
      return ok(text.slice(Math.trunc(start) - 1, Math.trunc(start) - 1 + Math.trunc(count)), dependencies);
    }
    case 'FIND':
    case 'SEARCH': {
      const needle = toText(values[0] ?? '');
      const haystack = toText(values[1] ?? '');
      const start = toNumber(values[2] ?? 1);
      if (start === null || start < 1) return error('#VALUE!', dependencies);
      const source = name === 'SEARCH' ? haystack.toLowerCase() : haystack;
      const target = name === 'SEARCH' ? needle.toLowerCase() : needle;
      const index = source.indexOf(target, Math.trunc(start) - 1);
      return index < 0 ? error('#VALUE!', dependencies) : ok(index + 1, dependencies);
    }
    case 'SUBSTITUTE': {
      const text = toText(values[0] ?? '');
      const oldText = toText(values[1] ?? '');
      const newText = toText(values[2] ?? '');
      if (oldText === '') return ok(text, dependencies);
      const instance = values[3] === undefined ? null : toNumber(values[3]);
      if (instance !== null && (instance === null || instance < 1)) return error('#VALUE!', dependencies);
      if (instance === null) return ok(text.split(oldText).join(newText), dependencies);
      let occurrence = 0;
      return ok(text.split(oldText).map((part, index, parts) => {
        if (index === parts.length - 1) return part;
        occurrence++;
        return `${part}${occurrence === Math.trunc(instance) ? newText : oldText}`;
      }).join(''), dependencies);
    }
    case 'REPLACE': {
      const text = toText(values[0] ?? '');
      const start = toNumber(values[1] ?? null);
      const count = toNumber(values[2] ?? null);
      if (start === null || count === null || start < 1 || count < 0) return error('#VALUE!', dependencies);
      return ok(text.slice(0, Math.trunc(start) - 1) + toText(values[3] ?? '') + text.slice(Math.trunc(start) - 1 + Math.trunc(count)), dependencies);
    }
    case 'TEXT':
      return values.length === 2 ? ok(formatTextValue(values[0]!, toText(values[1]!)), dependencies) : error('#VALUE!', dependencies);
    case 'CONCAT':
    case 'CONCATENATE':
      return ok(flatten(values).map((value) => toText(value)).join(''), dependencies);
    case 'TEXTJOIN': {
      const delimiter = toText(values[0] ?? '');
      const ignoreEmpty = truthy(values[1] ?? false);
      const textValues = flatten(values.slice(2)).map((value) => toText(value));
      return ok(textValues.filter((value) => !ignoreEmpty || value !== '').join(delimiter), dependencies);
    }
    case 'TEXTBEFORE':
      return evaluateTextBoundary('before', values, dependencies);
    case 'TEXTAFTER':
      return evaluateTextBoundary('after', values, dependencies);
    case 'DATEDIF':
      return evaluateDateDifference(values, dependencies);
    case 'TODAY':
      return args.length === 0 ? ok(context.getToday?.() ?? excelSerialFromDate(new Date()), dependencies) : error('#VALUE!', dependencies);
    case 'VALUE': {
      const number = toNumber(values[0] ?? null);
      return number === null ? error('#VALUE!', dependencies) : ok(number, dependencies);
    }
    case 'COUNTBLANK': {
      const matrix = toMatrix(values[0]);
      return !matrix ? error('#VALUE!', dependencies) : ok(flatten([matrix]).filter((value) => value === null || value === '').length, dependencies);
    }
    case 'ISBLANK':
      return values.length === 1 ? ok(values[0] === null || values[0] === '', dependencies) : error('#VALUE!', dependencies);
    case 'ISNUMBER':
      return values.length === 1 ? ok(typeof values[0] === 'number' && Number.isFinite(values[0]), dependencies) : error('#VALUE!', dependencies);
    case 'ISTEXT':
      return values.length === 1 ? ok(typeof values[0] === 'string', dependencies) : error('#VALUE!', dependencies);
    case 'DATE':
      return evaluateDate(values, dependencies);
    case 'DATEVALUE': {
      const date = toDateParts(values[0]);
      return values.length === 1 && date ? ok(excelSerialFromTime(date.time), dependencies) : error('#VALUE!', dependencies);
    }
    case 'TIME': {
      const hour = toNumber(values[0] ?? null);
      const minute = toNumber(values[1] ?? null);
      const second = toNumber(values[2] ?? null);
      if (values.length !== 3 || hour === null || minute === null || second === null) return error('#VALUE!', dependencies);
      return ok((Math.trunc(hour) * 3600 + Math.trunc(minute) * 60 + Math.trunc(second)) / 86400, dependencies);
    }
    case 'YEAR':
    case 'MONTH':
    case 'DAY':
      return evaluateDatePart(name, values, dependencies);
    case 'HOUR':
    case 'MINUTE':
    case 'SECOND':
      return evaluateTimePart(name, values, dependencies);
    case 'DAYS':
      return evaluateDays(values, dependencies);
    case 'WEEKDAY':
      return evaluateWeekday(values, dependencies);
    case 'EDATE':
    case 'EOMONTH':
      return evaluateMonthShift(name, values, dependencies);
    case 'WORKDAY':
      return evaluateWorkday(values, dependencies);
    case 'NETWORKDAYS':
      return evaluateNetworkdays(values, dependencies);
    case 'SUMIF':
      return evaluateConditionalAggregate('SUMIF', values, dependencies);
    case 'COUNTIF':
      return evaluateConditionalAggregate('COUNTIF', values, dependencies);
    case 'AVERAGEIF':
      return evaluateConditionalAggregate('AVERAGEIF', values, dependencies);
    case 'SUMIFS':
      return evaluateMultiConditionalAggregate('SUMIFS', values, dependencies);
    case 'COUNTIFS':
      return evaluateMultiConditionalAggregate('COUNTIFS', values, dependencies);
    case 'AVERAGEIFS':
      return evaluateMultiConditionalAggregate('AVERAGEIFS', values, dependencies);
    case 'MAXIFS':
      return evaluateMultiConditionalAggregate('MAXIFS', values, dependencies);
    case 'MINIFS':
      return evaluateMultiConditionalAggregate('MINIFS', values, dependencies);
    case 'VLOOKUP':
      return evaluateLookup('VLOOKUP', values, dependencies);
    case 'HLOOKUP':
      return evaluateLookup('HLOOKUP', values, dependencies);
    case 'INDEX':
      return evaluateIndex(values, dependencies);
    case 'MATCH':
      return evaluateMatch(values, dependencies);
    case 'XLOOKUP':
      return evaluateXLookup(values, dependencies);
    case 'XMATCH':
      return evaluateXMatch(values, dependencies);
    case 'FILTER':
      return evaluateFilter(values, dependencies);
    case 'SORT':
      return evaluateSort(values, dependencies);
    case 'UNIQUE':
      return evaluateUnique(values, dependencies);
    case 'SEQUENCE':
      return evaluateSequence(values, dependencies);
    case 'SORTBY':
      return evaluateSortBy(values, dependencies);
    case 'TAKE':
      return evaluateTakeDrop('TAKE', values, dependencies);
    case 'DROP':
      return evaluateTakeDrop('DROP', values, dependencies);
    case 'CHOOSECOLS':
      return evaluateChooseAxis('columns', values, dependencies);
    case 'CHOOSEROWS':
      return evaluateChooseAxis('rows', values, dependencies);
    case 'TOCOL':
      return evaluateToLine('column', values, dependencies);
    case 'TOROW':
      return evaluateToLine('row', values, dependencies);
    default:
      return error('#NAME?', dependencies);
  }
}

function evaluateMatrixBinary(
  operator: Extract<SpreadsheetFormulaAst, { kind: 'binary' }>['operator'],
  left: SpreadsheetFormulaValue,
  right: SpreadsheetFormulaValue,
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const leftMatrix = toMatrix(left);
  const rightMatrix = toMatrix(right);
  if (!leftMatrix || !rightMatrix) return error('#VALUE!', dependencies);
  const rows = Math.max(leftMatrix.length, rightMatrix.length);
  const columns = Math.max(leftMatrix[0]?.length ?? 0, rightMatrix[0]?.length ?? 0);
  if (!isScalar(left) && !isScalar(right) && !sameShape(leftMatrix, rightMatrix)) return error('#VALUE!', dependencies);
  const output: SpreadsheetFormulaScalar[][] = [];
  for (let row = 0; row < rows; row++) {
    const values: SpreadsheetFormulaScalar[] = [];
    for (let column = 0; column < columns; column++) {
      const leftValue = isScalar(left) ? left : leftMatrix[row]?.[column] ?? null;
      const rightValue = isScalar(right) ? right : rightMatrix[row]?.[column] ?? null;
      const result = evaluateScalarBinary(operator, leftValue, rightValue, dependencies);
      if (result.error) return result;
      if (Array.isArray(result.value)) return error('#VALUE!', dependencies);
      values.push(result.value);
    }
    output.push(values);
  }
  return ok(output, dependencies);
}

function evaluateScalarBinary(
  operator: Extract<SpreadsheetFormulaAst, { kind: 'binary' }>['operator'],
  left: SpreadsheetFormulaScalar,
  right: SpreadsheetFormulaScalar,
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (['=', '<>', '>', '>=', '<', '<='].includes(operator)) {
    return ok(compareValues(left, right, operator), dependencies);
  }
  if (operator === '&') return ok(`${toText(left)}${toText(right)}`, dependencies);
  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  if (leftNumber === null || rightNumber === null) return error('#VALUE!', dependencies);
  if (operator === '/' && rightNumber === 0) return error('#DIV/0!', dependencies);
  const value = operator === '+'
    ? leftNumber + rightNumber
    : operator === '-'
      ? leftNumber - rightNumber
      : operator === '*'
        ? leftNumber * rightNumber
        : leftNumber / rightNumber;
  return ok(value, dependencies);
}

function evaluateSumProduct(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const matrices = values.map(toMatrix);
  if (matrices.some((matrix) => !matrix) || matrices.length === 0) return error('#VALUE!', dependencies);
  const first = matrices[0]!;
  if (matrices.some((matrix) => !sameShape(first, matrix!))) return error('#VALUE!', dependencies);
  let total = 0;
  forEachMatrix(first, (_value, row, column) => {
    let product = 1;
    for (const matrix of matrices) {
      const number = toNumber(matrix![row]?.[column] ?? null);
      product *= number ?? 0;
    }
    total += product;
  });
  return ok(total, dependencies);
}

function isScalar(value: SpreadsheetFormulaValue): value is SpreadsheetFormulaScalar {
  return !Array.isArray(value);
}

function evaluateDateDifference(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length !== 3) return error('#VALUE!', dependencies);
  const start = toDateParts(values[0]);
  const end = toDateParts(values[1]);
  const unit = toText(values[2] ?? '').toUpperCase();
  if (!start || !end || !['Y', 'M', 'D'].includes(unit)) return error('#VALUE!', dependencies);
  if (start.time > end.time) return error('#NUM!', dependencies);
  if (unit === 'D') return ok(Math.floor((end.time - start.time) / MS_PER_DAY), dependencies);
  const totalMonths = (end.year - start.year) * 12 + end.month - start.month;
  const adjustedMonths = end.day < start.day ? totalMonths - 1 : totalMonths;
  return ok(unit === 'M' ? adjustedMonths : Math.floor(adjustedMonths / 12), dependencies);
}

function evaluateTextBoundary(
  mode: 'before' | 'after',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 2 || values.length > 4) return error('#VALUE!', dependencies);
  const text = toText(values[0] ?? '');
  const delimiter = toText(values[1] ?? '');
  if (delimiter === '') return error('#VALUE!', dependencies);
  const instanceValue = values[2] === undefined ? 1 : toNumber(values[2]);
  if (instanceValue === null || instanceValue === 0 || !Number.isFinite(instanceValue)) {
    return error('#VALUE!', dependencies);
  }
  const instance = Math.trunc(instanceValue);
  const index = findDelimiterIndex(text, delimiter, instance);
  if (index < 0) {
    return values.length === 4 ? ok(toText(values[3] ?? ''), dependencies) : error('#N/A', dependencies);
  }
  return ok(mode === 'before' ? text.slice(0, index) : text.slice(index + delimiter.length), dependencies);
}

function findDelimiterIndex(text: string, delimiter: string, instance: number): number {
  if (instance > 0) {
    let start = 0;
    let index = -1;
    for (let occurrence = 0; occurrence < instance; occurrence++) {
      index = text.indexOf(delimiter, start);
      if (index < 0) return -1;
      start = index + delimiter.length;
    }
    return index;
  }

  let end = text.length;
  let index = -1;
  for (let occurrence = 0; occurrence < Math.abs(instance); occurrence++) {
    index = text.lastIndexOf(delimiter, end);
    if (index < 0) return -1;
    end = index - 1;
  }
  return index;
}

function evaluateIf(
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length < 2 || args.length > 3) return error('#VALUE!');
  const condition = evaluateNode(args[0]!, context);
  if (condition.error) return condition;
  const branch = truthy(condition.value) ? args[1]! : args[2];
  if (!branch) return ok(false, condition.dependencies);
  const evaluated = evaluateNode(branch, context);
  return {
    ...evaluated,
    dependencies: [...condition.dependencies, ...evaluated.dependencies],
  };
}

function evaluateIfError(
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length !== 2) return error('#VALUE!');
  const value = evaluateNode(args[0]!, context);
  if (!value.error) return value;
  const fallback = evaluateNode(args[1]!, context);
  return {
    ...fallback,
    dependencies: [...value.dependencies, ...fallback.dependencies],
  };
}

function evaluateIfNa(
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length !== 2) return error('#VALUE!');
  const value = evaluateNode(args[0]!, context);
  if (!value.error || value.error !== '#N/A') return value;
  const fallback = evaluateNode(args[1]!, context);
  return {
    ...fallback,
    dependencies: [...value.dependencies, ...fallback.dependencies],
  };
}

function evaluateErrorPredicate(
  name: 'ISERROR' | 'ISERR' | 'ISNA',
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length !== 1) return error('#VALUE!');
  const evaluated = evaluateNode(args[0]!, context);
  const isError = evaluated.error !== null;
  const matches = name === 'ISERROR'
    ? isError
    : name === 'ISERR'
      ? isError && evaluated.error !== '#N/A'
      : evaluated.error === '#N/A';
  return ok(matches, evaluated.dependencies);
}

function evaluateLet(
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length < 3 || args.length % 2 === 0) return error('#VALUE!');
  const variables = new Map<string, SpreadsheetFormulaEvaluation>();
  let dependencies: SpreadsheetFormulaReference[] = [];
  const scopedContext = (): SpreadsheetFormulaEvaluationContext => ({
    ...context,
    getNameValue: (name) => variables.get(name.toUpperCase()) ?? context.getNameValue?.(name) ?? error('#NAME?'),
  });
  for (let index = 0; index < args.length - 1; index += 2) {
    const name = args[index];
    if (name?.kind !== 'name') return error('#NAME?', dependencies);
    const value = evaluateNode(args[index + 1]!, scopedContext());
    dependencies = [...dependencies, ...value.dependencies];
    if (value.error) return { ...value, dependencies };
    variables.set(name.name.toUpperCase(), value);
  }
  const result = evaluateNode(args[args.length - 1]!, scopedContext());
  return { ...result, dependencies: [...dependencies, ...result.dependencies] };
}

function evaluateIfs(
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length < 2 || args.length % 2 !== 0) return error('#VALUE!');
  let dependencies: SpreadsheetFormulaReference[] = [];
  for (let index = 0; index < args.length; index += 2) {
    const condition = evaluateNode(args[index]!, context);
    dependencies = [...dependencies, ...condition.dependencies];
    if (condition.error) return { ...condition, dependencies };
    if (truthy(condition.value)) {
      const result = evaluateNode(args[index + 1]!, context);
      return { ...result, dependencies: [...dependencies, ...result.dependencies] };
    }
  }
  return error('#N/A', dependencies);
}

function evaluateSwitch(
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length < 3) return error('#VALUE!');
  const expression = evaluateNode(args[0]!, context);
  if (expression.error) return expression;
  const hasDefault = (args.length - 1) % 2 === 1;
  const pairEnd = hasDefault ? args.length - 1 : args.length;
  let dependencies = [...expression.dependencies];
  for (let index = 1; index < pairEnd; index += 2) {
    const match = evaluateNode(args[index]!, context);
    dependencies = [...dependencies, ...match.dependencies];
    if (match.error) return { ...match, dependencies };
    if (compareValues(expression.value, match.value, '=')) {
      const result = evaluateNode(args[index + 1]!, context);
      return { ...result, dependencies: [...dependencies, ...result.dependencies] };
    }
  }
  if (hasDefault) {
    const result = evaluateNode(args[args.length - 1]!, context);
    return { ...result, dependencies: [...dependencies, ...result.dependencies] };
  }
  return error('#N/A', dependencies);
}

function evaluateSubtotal(
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length < 2) return error('#VALUE!');
  const functionNumber = evaluateNode(args[0]!, context);
  const number = toNumber(functionNumber.value);
  if (functionNumber.error || number === null || !Number.isInteger(number) || number < 1 || number > 111) {
    return error('#VALUE!', functionNumber.dependencies);
  }
  const baseFunction = number > 100 ? number - 100 : number;
  if (![1, 2, 3, 4, 5, 6, 9, 101, 102, 103, 104, 105, 106, 109].includes(number)) {
    return error('#VALUE!', functionNumber.dependencies);
  }
  let dependencies = [...functionNumber.dependencies];
  const values: SpreadsheetFormulaValue[] = [];
  for (const argument of args.slice(1)) {
    const evaluated = evaluateNode(argument, context);
    dependencies = [...dependencies, ...evaluated.dependencies];
    if (evaluated.error) return { ...evaluated, dependencies };
    values.push(number >= 100 ? filterHiddenRows(evaluated.value, argument, context) : evaluated.value);
  }
  return aggregateValues(baseFunction, values, dependencies);
}

function evaluateAggregate(
  args: SpreadsheetFormulaAst[],
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaEvaluation {
  if (args.length < 3) return error('#VALUE!');
  const functionNumber = evaluateNode(args[0]!, context);
  const options = evaluateNode(args[1]!, context);
  const number = toNumber(functionNumber.value);
  const option = toNumber(options.value);
  if (functionNumber.error || options.error || number === null || option === null || !Number.isInteger(number) || !Number.isInteger(option)) {
    return error('#VALUE!', [...functionNumber.dependencies, ...options.dependencies]);
  }
  const hiddenRows = [1, 3, 5, 7].includes(option);
  let dependencies = [...functionNumber.dependencies, ...options.dependencies];
  const values: SpreadsheetFormulaValue[] = [];
  for (const argument of args.slice(2, number === 14 || number === 15 ? 3 : undefined)) {
    const evaluated = evaluateNode(argument, context);
    dependencies = [...dependencies, ...evaluated.dependencies];
    if (evaluated.error) return { ...evaluated, dependencies };
    values.push(hiddenRows ? filterHiddenRows(evaluated.value, argument, context) : evaluated.value);
  }
  if (number === 14 || number === 15) {
    const k = toNumber(args[3] ? (evaluateNode(args[3], context).value) : null);
    if (k === null) return error('#VALUE!', dependencies);
    values.push(k);
  }
  return aggregateValues(number, values, dependencies);
}

function filterHiddenRows(
  value: SpreadsheetFormulaValue,
  ast: SpreadsheetFormulaAst,
  context: SpreadsheetFormulaEvaluationContext,
): SpreadsheetFormulaValue {
  const isRowHidden = context.isRowHidden;
  if (!isRowHidden || !Array.isArray(value)) return value;
  if (ast.kind === 'cell') {
    return isRowHidden(parseCellAddress(ast.address).row, ast.sheetName) ? [] : value;
  }
  if (ast.kind !== 'range') return value;
  const start = parseCellAddress(ast.startAddress);
  const end = parseCellAddress(ast.endAddress);
  return value.filter((_row, offset) => !isRowHidden(start.row + offset, ast.sheetName));
}

function aggregateValues(
  functionNumber: number,
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (functionNumber === 1) {
    const numbers = numericValues(values);
    return numbers.length ? ok(numbers.reduce((sum, value) => sum + value, 0) / numbers.length, dependencies) : error('#DIV/0!', dependencies);
  }
  if (functionNumber === 2) return ok(numericValues(values).length, dependencies);
  if (functionNumber === 3) return ok(flatten(values).filter((value) => value !== null && value !== '').length, dependencies);
  if (functionNumber === 4) {
    const numbers = numericValues(values);
    return numbers.length ? ok(Math.max(...numbers), dependencies) : ok(0, dependencies);
  }
  if (functionNumber === 5) {
    const numbers = numericValues(values);
    return numbers.length ? ok(Math.min(...numbers), dependencies) : ok(0, dependencies);
  }
  if (functionNumber === 6) {
    return ok(numericValues(values).reduce((product, value) => product * value, 1), dependencies);
  }
  if (functionNumber === 9) return ok(numericValues(values).reduce((sum, value) => sum + value, 0), dependencies);
  if (functionNumber === 12) {
    const numbers = numericValues(values).sort((left, right) => left - right);
    if (!numbers.length) return error('#NUM!', dependencies);
    const middle = Math.floor(numbers.length / 2);
    return ok(numbers.length % 2 ? numbers[middle]! : (numbers[middle - 1]! + numbers[middle]!) / 2, dependencies);
  }
  if (functionNumber === 14 || functionNumber === 15) {
    const k = toNumber(values[values.length - 1] ?? null);
    const numbers = numericValues(values.slice(0, -1)).sort((left, right) => functionNumber === 14 ? right - left : left - right);
    if (k === null || !Number.isInteger(k) || k < 1 || k > numbers.length) return error('#NUM!', dependencies);
    return ok(numbers[k - 1]!, dependencies);
  }
  return error('#VALUE!', dependencies);
}

function roundDirectional(number: number, digits: number, roundUp: boolean): number {
  const factor = 10 ** digits;
  const scaled = number * factor;
  const rounded = roundUp
    ? Math.sign(scaled) * Math.ceil(Math.abs(scaled))
    : Math.trunc(scaled);
  return rounded / factor;
}

function evaluateFinancial(
  name: 'PMT' | 'FV' | 'PV',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 3 || values.length > 5) return error('#VALUE!', dependencies);
  const rate = toNumber(values[0] ?? null);
  const periods = toNumber(values[1] ?? null);
  const third = toNumber(values[2] ?? null);
  const fourth = toNumber(values[3] ?? 0);
  const type = toNumber(values[4] ?? 0);
  if (rate === null || periods === null || third === null || fourth === null || type === null) {
    return error('#VALUE!', dependencies);
  }
  if (periods === 0 || ![0, 1].includes(Math.trunc(type))) return error('#NUM!', dependencies);
  const normalizedType = Math.trunc(type);
  const growth = (1 + rate) ** periods;
  if (!Number.isFinite(growth)) return error('#NUM!', dependencies);
  if (rate === 0) {
    if (name === 'PMT') return ok(-(third + fourth) / periods, dependencies);
    if (name === 'FV') return ok(-(fourth + third * periods), dependencies);
    return ok(-(fourth + third * periods), dependencies);
  }
  const annuity = ((growth - 1) / rate) * (1 + rate * normalizedType);
  if (name === 'PMT') return ok(-(third * growth + fourth) * rate / (growth - 1), dependencies);
  if (name === 'FV') return ok(-(fourth * growth + third * annuity), dependencies);
  return ok(-(fourth + third * annuity) / growth, dependencies);
}

function evaluateIrr(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const cashFlows = numericValues([values[0] ?? null]);
  const guess = toNumber(values[1] ?? 0.1);
  if (!cashFlows.length || guess === null) return error('#VALUE!', dependencies);
  const rate = solveRate(cashFlows, guess);
  return rate === null ? error('#NUM!', dependencies) : ok(rate, dependencies);
}

function evaluateRate(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 3 || values.length > 6) return error('#VALUE!', dependencies);
  const periods = toNumber(values[0] ?? null);
  const payment = toNumber(values[1] ?? null);
  const presentValue = toNumber(values[2] ?? null);
  const futureValue = toNumber(values[3] ?? 0);
  const type = toNumber(values[4] ?? 0);
  const guess = toNumber(values[5] ?? 0.1);
  if ([periods, payment, presentValue, futureValue, type, guess].some((value) => value === null)) {
    return error('#VALUE!', dependencies);
  }
  if (periods === 0 || ![0, 1].includes(Math.trunc(type!))) return error('#NUM!', dependencies);
  const normalizedPeriods = Math.trunc(periods!);
  const normalizedType = Math.trunc(type!);
  const cashFlows = [presentValue! + (normalizedType === 1 ? payment! : 0)];
  for (let index = 1; index <= normalizedPeriods; index++) {
    cashFlows.push((normalizedType === 1 && index === normalizedPeriods ? 0 : payment!) + (index === normalizedPeriods ? futureValue! : 0));
  }
  const rate = solveRate(cashFlows, guess!);
  return rate === null ? error('#NUM!', dependencies) : ok(rate, dependencies);
}

function solveRate(cashFlows: number[], initialGuess: number): number | null {
  const evaluate = (rate: number): number => cashFlows.reduce((sum, cashFlow, index) => sum + cashFlow / (1 + rate) ** index, 0);
  const derivative = (rate: number): number => cashFlows.reduce((sum, cashFlow, index) => (
    index === 0 ? sum : sum - index * cashFlow / (1 + rate) ** (index + 1)
  ), 0);
  let rate = Math.max(-0.9999, initialGuess);
  for (let iteration = 0; iteration < 100; iteration++) {
    const value = evaluate(rate);
    if (Math.abs(value) < 1e-10) return rate;
    const slope = derivative(rate);
    if (!Number.isFinite(value) || !Number.isFinite(slope) || Math.abs(slope) < 1e-14) break;
    const next = rate - value / slope;
    if (!Number.isFinite(next) || next <= -1 || Math.abs(next - rate) > 10) break;
    rate = next;
  }
  let left = -0.9999;
  let leftValue = evaluate(left);
  for (let index = 1; index <= 2000; index++) {
    const right = -0.9999 + index * 0.01;
    const rightValue = evaluate(right);
    if (Number.isFinite(leftValue) && Number.isFinite(rightValue) && leftValue * rightValue <= 0) {
      let low = left;
      let high = right;
      for (let iteration = 0; iteration < 80; iteration++) {
        const middle = (low + high) / 2;
        const middleValue = evaluate(middle);
        if (Math.abs(middleValue) < 1e-10) return middle;
        if (leftValue * middleValue <= 0) {
          high = middle;
        } else {
          low = middle;
          leftValue = middleValue;
        }
      }
      return (low + high) / 2;
    }
    left = right;
    leftValue = rightValue;
  }
  return null;
}

function evaluateVariance(
  name: 'VAR' | 'VAR.P' | 'VAR.S' | 'STDEV' | 'STDEV.P' | 'STDEV.S',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const numbers = numericValues(values);
  const sample = name === 'VAR' || name === 'STDEV' || name.endsWith('.S');
  if (numbers.length < (sample ? 2 : 1)) return error('#DIV/0!', dependencies);
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const variance = numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (numbers.length - (sample ? 1 : 0));
  return ok(name.startsWith('STDEV') ? Math.sqrt(variance) : variance, dependencies);
}

function evaluatePercentile(
  name: 'PERCENTILE' | 'PERCENTILE.INC' | 'PERCENTILE.EXC',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length !== 2) return error('#VALUE!', dependencies);
  return percentileValue(values[0], values[1]!, name === 'PERCENTILE.EXC', dependencies);
}

function evaluateQuartile(
  name: 'QUARTILE' | 'QUARTILE.INC' | 'QUARTILE.EXC',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length !== 2) return error('#VALUE!', dependencies);
  const quartile = toNumber(values[1] ?? null);
  if (quartile === null || !Number.isInteger(quartile) || quartile < 0 || quartile > 4) {
    return error('#NUM!', dependencies);
  }
  return percentileValue(values[0], quartile / 4, name === 'QUARTILE.EXC', dependencies);
}

function percentileValue(
  source: SpreadsheetFormulaValue | undefined,
  percentile: SpreadsheetFormulaValue,
  exclusive: boolean,
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const numbers = numericValues([source ?? null]).sort((left, right) => left - right);
  const k = toNumber(percentile);
  if (!numbers.length || k === null) return error('#NUM!', dependencies);
  const position = exclusive ? k * (numbers.length + 1) - 1 : k * (numbers.length - 1);
  if (position < 0 || position > numbers.length - 1) return error('#NUM!', dependencies);
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return ok(numbers[lower]!, dependencies);
  const fraction = position - lower;
  return ok(numbers[lower]! + (numbers[upper]! - numbers[lower]!) * fraction, dependencies);
}

function evaluateDate(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length !== 3) return error('#VALUE!', dependencies);
  const year = toNumber(values[0] ?? null);
  const month = toNumber(values[1] ?? null);
  const day = toNumber(values[2] ?? null);
  if (year === null || month === null || day === null) return error('#VALUE!', dependencies);
  const normalizedYear = Math.trunc(year) < 1900 ? Math.trunc(year) + 1900 : Math.trunc(year);
  const time = Date.UTC(normalizedYear, Math.trunc(month) - 1, Math.trunc(day));
  return Number.isFinite(time) ? ok(Math.floor((time - Date.UTC(1899, 11, 30)) / MS_PER_DAY), dependencies) : error('#NUM!', dependencies);
}

function evaluateDatePart(
  name: 'YEAR' | 'MONTH' | 'DAY',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const date = toDateParts(values[0]);
  if (!date || values.length !== 1) return error('#VALUE!', dependencies);
  return ok(name === 'YEAR' ? date.year : name === 'MONTH' ? date.month : date.day, dependencies);
}

function evaluateTimePart(
  name: 'HOUR' | 'MINUTE' | 'SECOND',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const value = values[0];
  const date = toDateParts(value);
  if (!date || values.length !== 1) return error('#VALUE!', dependencies);
  const time = ((date.time % MS_PER_DAY) + MS_PER_DAY) % MS_PER_DAY;
  const totalSeconds = Math.floor(time / 1000);
  if (name === 'HOUR') return ok(Math.floor(totalSeconds / 3600), dependencies);
  if (name === 'MINUTE') return ok(Math.floor((totalSeconds % 3600) / 60), dependencies);
  return ok(totalSeconds % 60, dependencies);
}

function evaluateDays(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const end = toDateParts(values[0]);
  const start = toDateParts(values[1]);
  if (!end || !start || values.length !== 2) return error('#VALUE!', dependencies);
  return ok(Math.floor((end.time - start.time) / MS_PER_DAY), dependencies);
}

function evaluateWeekday(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const date = toDateParts(values[0]);
  const returnType = toNumber(values[1] ?? 1);
  if (!date || values.length < 1 || values.length > 2 || returnType === null || ![1, 2, 3].includes(Math.trunc(returnType))) return error('#NUM!', dependencies);
  const day = new Date(date.time).getUTCDay();
  if (Math.trunc(returnType) === 2) return ok(day === 0 ? 7 : day, dependencies);
  if (Math.trunc(returnType) === 3) return ok(day === 0 ? 6 : day - 1, dependencies);
  return ok(day + 1, dependencies);
}

function evaluateMonthShift(
  name: 'EDATE' | 'EOMONTH',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const date = toDateParts(values[0]);
  const months = toNumber(values[1] ?? null);
  if (!date || months === null || values.length !== 2) return error('#VALUE!', dependencies);
  const shifted = new Date(Date.UTC(date.year, date.month - 1 + Math.trunc(months), name === 'EOMONTH' ? 0 : date.day));
  return ok(Math.floor((shifted.getTime() - Date.UTC(1899, 11, 30)) / MS_PER_DAY), dependencies);
}

function evaluateWorkday(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 2 || values.length > 3) return error('#VALUE!', dependencies);
  const start = toDateParts(values[0]);
  const days = toNumber(values[1] ?? null);
  const holidays = holidayTimes(values[2]);
  if (!start || days === null || holidays === null) return error('#VALUE!', dependencies);
  const direction = days < 0 ? -1 : 1;
  let remaining = Math.abs(Math.trunc(days));
  let time = start.time;
  while (remaining > 0 || isWeekend(time) || holidays.has(time)) {
    time += direction * MS_PER_DAY;
    if (!isWeekend(time) && !holidays.has(time)) remaining--;
  }
  return ok(excelSerialFromTime(time), dependencies);
}

function evaluateNetworkdays(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 2 || values.length > 3) return error('#VALUE!', dependencies);
  const start = toDateParts(values[0]);
  const end = toDateParts(values[1]);
  const holidays = holidayTimes(values[2]);
  if (!start || !end || holidays === null) return error('#VALUE!', dependencies);
  const direction = start.time <= end.time ? 1 : -1;
  let time = direction === 1 ? start.time : end.time;
  const limit = direction === 1 ? end.time : start.time;
  let count = 0;
  while (time <= limit) {
    if (!isWeekend(time) && !holidays.has(time)) count++;
    time += MS_PER_DAY;
  }
  return ok(direction * count, dependencies);
}

function holidayTimes(value: SpreadsheetFormulaValue | undefined): Set<number> | null {
  if (value === undefined) return new Set();
  const dates = flatten([value]).map(toDateParts);
  if (dates.some((date) => !date)) return null;
  return new Set(dates.map((date) => date!.time));
}

function isWeekend(time: number): boolean {
  const day = new Date(time).getUTCDay();
  return day === 0 || day === 6;
}

function excelSerialFromTime(time: number): number {
  return Math.floor((time - Date.UTC(1899, 11, 30)) / MS_PER_DAY);
}

function formatTextValue(value: SpreadsheetFormulaValue, format: string): string {
  const number = toNumber(value);
  if (number === null) return toText(value);
  if (format.includes('%')) {
    const decimals = decimalPlaces(format.replace('%', ''));
    return `${(number * 100).toFixed(decimals)}%`;
  }
  const decimals = decimalPlaces(format);
  if (format.includes(',')) return number.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  return decimals > 0 ? number.toFixed(decimals) : String(Math.round(number));
}

function decimalPlaces(format: string): number {
  const match = /\.([0#]+)/.exec(format);
  return match?.[1]?.length ?? 0;
}

function evaluateXLookup(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 3 || values.length > 6) return error('#VALUE!', dependencies);
  const lookupValue = values[0];
  const lookupArray = flatten([values[1]!]);
  const returnArray = flatten([values[2]!]);
  const matchMode = toNumber(values[4] ?? 0);
  if (matchMode === null || Math.trunc(matchMode) !== 0 || lookupArray.length !== returnArray.length) return error('#VALUE!', dependencies);
  const index = lookupArray.findIndex((value) => compareValues(value, lookupValue ?? null, '='));
  if (index < 0) return values[3] === undefined ? error('#N/A', dependencies) : ok(values[3]!, dependencies);
  return ok(returnArray[index] ?? null, dependencies);
}

function evaluateXMatch(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 2 || values.length > 4) return error('#VALUE!', dependencies);
  const lookupArray = flatten([values[1]!]);
  const matchMode = toNumber(values[2] ?? 0);
  if (matchMode === null || Math.trunc(matchMode) !== 0) return error('#N/A', dependencies);
  const index = lookupArray.findIndex((value) => compareValues(value, values[0] ?? null, '='));
  return index < 0 ? error('#N/A', dependencies) : ok(index + 1, dependencies);
}

function evaluateFilter(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 2 || values.length > 3) return error('#VALUE!', dependencies);
  const source = toMatrix(values[0]);
  const include = toMatrix(values[1]);
  if (!source || !include) return error('#VALUE!', dependencies);
  const rowMode = include.length === source.length && include[0]?.length === 1;
  const columnMode = include.length === 1 && include[0]?.length === source[0]?.length;
  if (!rowMode && !columnMode) return error('#VALUE!', dependencies);
  if (rowMode) {
    const rows = source.filter((_row, index) => truthy(include[index]?.[0] ?? false));
    return rows.length > 0 ? ok(rows, dependencies) : values[2] !== undefined ? ok(toMatrix(values[2])!, dependencies) : error('#CALC!', dependencies);
  }
  const columns = source[0]?.map((_value, column) => column).filter((column) => truthy(include[0]?.[column] ?? false)) ?? [];
  if (columns.length === 0) return values[2] !== undefined ? ok(toMatrix(values[2])!, dependencies) : error('#CALC!', dependencies);
  return ok(source.map((row) => columns.map((column) => row[column] ?? null)), dependencies);
}

function evaluateSort(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 1 || values.length > 4) return error('#VALUE!', dependencies);
  const source = toMatrix(values[0]);
  const sortIndex = toNumber(values[1] ?? 1);
  const sortOrder = toNumber(values[2] ?? 1);
  const byColumn = truthy(values[3] ?? false);
  if (!source || sortIndex === null || sortOrder === null || ![1, -1].includes(Math.trunc(sortOrder))) return error('#VALUE!', dependencies);
  const index = Math.trunc(sortIndex) - 1;
  if (byColumn) {
    if (index < 0 || index >= source.length) return error('#VALUE!', dependencies);
    const columns = Array.from({ length: source[0]?.length ?? 0 }, (_, column) => column);
    columns.sort((left, right) => compareForSort(source[index]?.[left] ?? null, source[index]?.[right] ?? null) * Math.trunc(sortOrder));
    return ok(source.map((row) => columns.map((column) => row[column] ?? null)), dependencies);
  }
  if (index < 0 || index >= (source[0]?.length ?? 0)) return error('#VALUE!', dependencies);
  const rows = source.map((_row, row) => row);
  rows.sort((left, right) => compareForSort(source[left]?.[index] ?? null, source[right]?.[index] ?? null) * Math.trunc(sortOrder));
  return ok(rows.map((row) => [...source[row]!]), dependencies);
}

function evaluateUnique(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 1 || values.length > 3) return error('#VALUE!', dependencies);
  const source = toMatrix(values[0]);
  if (!source) return error('#VALUE!', dependencies);
  const byColumn = truthy(values[1] ?? false);
  const exactlyOnce = truthy(values[2] ?? false);
  const items = byColumn
    ? (source[0] ?? []).map((_value, column) => source.map((row) => row[column] ?? null))
    : source.map((row) => [...row]);
  const keys = items.map((item) => JSON.stringify(item));
  const counts = new Map<string, number>();
  keys.forEach((key) => counts.set(key, (counts.get(key) ?? 0) + 1));
  const uniqueItems = items.filter((item, index) => {
    const first = keys.indexOf(keys[index]!);
    return first === index && (!exactlyOnce || counts.get(keys[index]!) === 1);
  });
  if (byColumn) {
    return ok(uniqueItems[0]?.map((_value, row) => uniqueItems.map((item) => item[row] ?? null)) ?? [], dependencies);
  }
  return ok(uniqueItems, dependencies);
}

function evaluateSequence(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 1 || values.length > 4) return error('#VALUE!', dependencies);
  const rows = toNumber(values[0] ?? null);
  const columns = toNumber(values[1] ?? 1);
  const start = toNumber(values[2] ?? 1);
  const step = toNumber(values[3] ?? 1);
  if (rows === null || columns === null || start === null || step === null || rows < 1 || columns < 1 || !Number.isInteger(rows) || !Number.isInteger(columns)) return error('#VALUE!', dependencies);
  return ok(Array.from({ length: Math.trunc(rows) }, (_row, row) =>
    Array.from({ length: Math.trunc(columns) }, (_column, column) => start + (row * Math.trunc(columns) + column) * step)), dependencies);
}

function compareForSort(left: SpreadsheetFormulaScalar, right: SpreadsheetFormulaScalar): number {
  if (compareValues(left, right, '=')) return 0;
  return compareValues(left, right, '<') ? -1 : 1;
}

function evaluateSortBy(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 3 || values.length % 2 === 0) return error('#VALUE!', dependencies);
  const source = toMatrix(values[0]);
  if (!source) return error('#VALUE!', dependencies);
  const rowCount = source.length;
  const columnCount = source[0]?.length ?? 0;
  const keys = [] as { values: SpreadsheetFormulaScalar[]; order: number }[];
  for (let index = 1; index < values.length; index += 2) {
    const keyMatrix = toMatrix(values[index]);
    const order = toNumber(values[index + 1] ?? 1);
    if (!keyMatrix || order === null || ![1, -1].includes(Math.trunc(order))) return error('#VALUE!', dependencies);
    const rowKey = keyMatrix.length === rowCount && keyMatrix.every((row) => row.length === 1);
    const columnKey = keyMatrix.length === 1 && keyMatrix[0]!.length === columnCount;
    if (!rowKey && !columnKey) return error('#VALUE!', dependencies);
    keys.push({
      values: rowKey ? keyMatrix.map((row) => row[0] ?? null) : keyMatrix[0]!.slice(),
      order: Math.trunc(order),
    });
  }
  const rowMode = keys[0]!.values.length === rowCount;
  const positions = Array.from({ length: rowMode ? rowCount : columnCount }, (_, index) => index);
  positions.sort((left, right) => {
    for (const key of keys) {
      const comparison = compareForSort(key.values[left] ?? null, key.values[right] ?? null) * key.order;
      if (comparison !== 0) return comparison;
    }
    return left - right;
  });
  return ok(rowMode
    ? positions.map((row) => [...source[row]!])
    : source.map((row) => positions.map((column) => row[column] ?? null)), dependencies);
}

function evaluateTakeDrop(
  name: 'TAKE' | 'DROP',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 2 || values.length > 3) return error('#VALUE!', dependencies);
  const source = toMatrix(values[0]);
  const rowCount = toNumber(values[1] ?? null);
  const columnCount = values[2] === undefined ? null : toNumber(values[2]);
  if (!source || rowCount === null || columnCount === null && values[2] !== undefined) return error('#VALUE!', dependencies);
  const rows = sliceMatrix(name, source, Math.trunc(rowCount), 0);
  if (!rows.length) return error('#CALC!', dependencies);
  if (columnCount === null) return ok(rows, dependencies);
  const columns = sliceMatrix(name, rows, Math.trunc(columnCount), 1);
  return columns.length && columns[0]!.length ? ok(columns, dependencies) : error('#CALC!', dependencies);
}

function sliceMatrix(
  name: 'TAKE' | 'DROP',
  source: SpreadsheetFormulaScalar[][],
  count: number,
  axis: 0 | 1,
): SpreadsheetFormulaScalar[][] {
  const size = axis === 0 ? source.length : source[0]?.length ?? 0;
  if (count === 0 || Math.abs(count) > size) return [];
  const start = name === 'TAKE'
    ? count > 0 ? 0 : size + count
    : count > 0 ? count : 0;
  const end = name === 'TAKE'
    ? count > 0 ? count : size
    : count > 0 ? size : size + count;
  if (axis === 0) return source.slice(start, end).map((row) => [...row]);
  return source.map((row) => row.slice(start, end));
}

function evaluateChooseAxis(
  axis: 'rows' | 'columns',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 2) return error('#VALUE!', dependencies);
  const source = toMatrix(values[0]);
  if (!source) return error('#VALUE!', dependencies);
  const size = axis === 'rows' ? source.length : source[0]?.length ?? 0;
  const indexes: number[] = [];
  for (const value of values.slice(1)) {
    const index = toNumber(value);
    if (index === null || !Number.isInteger(index) || index === 0) return error('#VALUE!', dependencies);
    const normalized = index > 0 ? index - 1 : size + index;
    if (normalized < 0 || normalized >= size) return error('#VALUE!', dependencies);
    indexes.push(normalized);
  }
  if (axis === 'rows') return ok(indexes.map((row) => [...source[row]!]), dependencies);
  return ok(source.map((row) => indexes.map((column) => row[column] ?? null)), dependencies);
}

function evaluateToLine(
  axis: 'row' | 'column',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  if (values.length < 1 || values.length > 3) return error('#VALUE!', dependencies);
  const source = toMatrix(values[0]);
  const ignore = toNumber(values[1] ?? 0);
  const scanByColumn = truthy(values[2] ?? false);
  if (!source || ignore === null || ![0, 1, 2, 3].includes(Math.trunc(ignore))) return error('#VALUE!', dependencies);
  const flattened: SpreadsheetFormulaScalar[] = [];
  if (scanByColumn) {
    for (let column = 0; column < (source[0]?.length ?? 0); column++) {
      for (const row of source) flattened.push(row[column] ?? null);
    }
  } else {
    for (const row of source) flattened.push(...row);
  }
  const filtered = flattened.filter((value) => {
    const blank = value === null || value === '';
    return ignore === 0 || (ignore === 1 && !blank) || (ignore === 2) || (ignore === 3 && !blank);
  });
  return ok(axis === 'column' ? filtered.map((value) => [value]) : [filtered], dependencies);
}

function evaluateConditionalAggregate(
  name: 'SUMIF' | 'COUNTIF' | 'AVERAGEIF',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const criteriaRange = toMatrix(values[0]);
  const criteria = values[1];
  if (!criteriaRange || criteria === undefined) return error('#VALUE!', dependencies);
  const sumRange = toMatrix(name === 'COUNTIF' ? values[0] : (values[2] ?? values[0]));
  if (!sumRange || !sameShape(criteriaRange, sumRange)) return error('#VALUE!', dependencies);
  const matched: SpreadsheetFormulaScalar[] = [];
  forEachMatrix(criteriaRange, (criteriaValue, row, column) => {
    if (matchesCriteria(criteriaValue, criteria)) matched.push(sumRange[row]?.[column] ?? null);
  });
  if (name === 'COUNTIF') return ok(matched.length, dependencies);
  const numbers = numericValues(matched);
  if (name === 'SUMIF') return ok(numbers.reduce((sum, value) => sum + value, 0), dependencies);
  return numbers.length === 0
    ? error('#DIV/0!', dependencies)
    : ok(numbers.reduce((sum, value) => sum + value, 0) / numbers.length, dependencies);
}

function evaluateMultiConditionalAggregate(
  name: 'SUMIFS' | 'COUNTIFS' | 'AVERAGEIFS' | 'MAXIFS' | 'MINIFS',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const valueRange = toMatrix(values[0]);
  if (!valueRange || values.length < 3) return error('#VALUE!', dependencies);
  const criteriaPairs = name === 'COUNTIFS' ? values : values.slice(1);
  if (criteriaPairs.length % 2 !== 0) return error('#VALUE!', dependencies);
  const pairs: { range: SpreadsheetFormulaScalar[][]; criteria: SpreadsheetFormulaValue }[] = [];
  for (let index = 0; index < criteriaPairs.length; index += 2) {
    const range = toMatrix(criteriaPairs[index]);
    const criteria = criteriaPairs[index + 1];
    if (!range || criteria === undefined || !sameShape(valueRange, range)) return error('#VALUE!', dependencies);
    pairs.push({ range, criteria });
  }
  const matched: SpreadsheetFormulaScalar[] = [];
  forEachMatrix(valueRange, (value, row, column) => {
    if (pairs.every((pair) => matchesCriteria(pair.range[row]?.[column] ?? null, pair.criteria))) {
      matched.push(value);
    }
  });
  if (name === 'COUNTIFS') return ok(matched.length, dependencies);
  const numbers = numericValues(matched);
  if (name === 'MAXIFS') return ok(numbers.length ? Math.max(...numbers) : 0, dependencies);
  if (name === 'MINIFS') return ok(numbers.length ? Math.min(...numbers) : 0, dependencies);
  if (name === 'SUMIFS') return ok(numbers.reduce((sum, value) => sum + value, 0), dependencies);
  return numbers.length === 0
    ? error('#DIV/0!', dependencies)
    : ok(numbers.reduce((sum, value) => sum + value, 0) / numbers.length, dependencies);
}

function evaluateLookup(
  name: 'VLOOKUP' | 'HLOOKUP',
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const lookupValue = values[0];
  const table = toMatrix(values[1]);
  const index = toNumber(values[2] ?? null);
  const approximate = values.length >= 4 ? truthy(values[3]!) : true;
  if (!table || index === null || index < 1 || approximate) return error(approximate ? '#N/A' : '#VALUE!', dependencies);
  const targetIndex = Math.trunc(index) - 1;
  if (name === 'VLOOKUP') {
    if (targetIndex >= (table[0]?.length ?? 0)) return error('#REF!', dependencies);
    const row = table.find((item) => compareValues(item[0] ?? null, lookupValue ?? null, '='));
    return row ? ok(row[targetIndex] ?? null, dependencies) : error('#N/A', dependencies);
  }
  if (targetIndex >= table.length) return error('#REF!', dependencies);
  const column = table[0]?.findIndex((value) => compareValues(value, lookupValue ?? null, '=')) ?? -1;
  return column >= 0 ? ok(table[targetIndex]?.[column] ?? null, dependencies) : error('#N/A', dependencies);
}

function evaluateIndex(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const matrix = toMatrix(values[0]);
  const row = toNumber(values[1] ?? 1);
  const column = toNumber(values[2] ?? 1);
  if (!matrix || row === null || column === null || row < 1 || column < 1) return error('#VALUE!', dependencies);
  const value = matrix[Math.trunc(row) - 1]?.[Math.trunc(column) - 1];
  return value === undefined ? error('#REF!', dependencies) : ok(value, dependencies);
}

function evaluateMatch(
  values: SpreadsheetFormulaValue[],
  dependencies: SpreadsheetFormulaReference[],
): SpreadsheetFormulaEvaluation {
  const lookupValue = values[0];
  const matrix = toMatrix(values[1]);
  const matchType = toNumber(values[2] ?? 1);
  if (!matrix || matchType === null || matchType !== 0) return error('#N/A', dependencies);
  const vector = matrix.length === 1 ? matrix[0]! : matrix.map((row) => row[0] ?? null);
  const index = vector.findIndex((value) => compareValues(value, lookupValue ?? null, '='));
  return index >= 0 ? ok(index + 1, dependencies) : error('#N/A', dependencies);
}

function ok(
  value: SpreadsheetFormulaValue,
  dependencies: SpreadsheetFormulaReference[] = [],
): SpreadsheetFormulaEvaluation {
  return { value, error: null, dependencies };
}

function error(
  value: SpreadsheetFormulaError,
  dependencies: SpreadsheetFormulaReference[] = [],
): SpreadsheetFormulaEvaluation {
  return { value: null, error: value, dependencies };
}

function numericValues(values: SpreadsheetFormulaValue[]): number[] {
  return flatten(values)
    .map((value) => toNumber(value))
    .filter((value): value is number => value !== null);
}

function toMatrix(value: SpreadsheetFormulaValue | undefined): SpreadsheetFormulaScalar[][] | null {
  if (value === undefined) return null;
  return Array.isArray(value) ? value : [[value]];
}

function sameShape(
  left: SpreadsheetFormulaScalar[][],
  right: SpreadsheetFormulaScalar[][],
): boolean {
  return left.length === right.length && left.every((row, index) => row.length === (right[index]?.length ?? -1));
}

function forEachMatrix(
  matrix: SpreadsheetFormulaScalar[][],
  visit: (value: SpreadsheetFormulaScalar, row: number, column: number) => void,
): void {
  for (let row = 0; row < matrix.length; row++) {
    for (let column = 0; column < matrix[row]!.length; column++) {
      visit(matrix[row]![column]!, row, column);
    }
  }
}

function flatten(values: SpreadsheetFormulaValue[]): SpreadsheetFormulaScalar[] {
  return values.flatMap((value) => (Array.isArray(value) ? value.flat() : [value]));
}

function toText(value: SpreadsheetFormulaValue): string {
  if (Array.isArray(value)) return value.flat().map((item) => toText(item)).join('');
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

function toNumber(value: SpreadsheetFormulaValue): number | null {
  if (Array.isArray(value)) return null;
  if (value === null || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'string') {
    const number = Number(value);
    return Number.isFinite(number) ? number : null;
  }
  return null;
}

function matchesCriteria(
  value: SpreadsheetFormulaScalar,
  criteria: SpreadsheetFormulaValue,
): boolean {
  if (Array.isArray(criteria)) return false;
  if (typeof criteria !== 'string') return compareValues(value, criteria, '=');
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(criteria);
  if (!match) return wildcardMatches(value, criteria) ?? compareValues(value, criteria, '=');
  const operator = match[1]!;
  const rawRight = match[2]!;
  const numericRight = Number(rawRight);
  const right: SpreadsheetFormulaScalar = rawRight !== '' && Number.isFinite(numericRight) ? numericRight : rawRight;
  if (operator === '=' || operator === '<>') {
    const wildcard = wildcardMatches(value, rawRight);
    if (wildcard !== null) return operator === '=' ? wildcard : !wildcard;
  }
  return compareValues(value, right, operator);
}

function wildcardMatches(value: SpreadsheetFormulaScalar, criteria: string): boolean | null {
  if (!criteria.includes('*') && !criteria.includes('?')) return null;
  let pattern = '';
  for (let index = 0; index < criteria.length; index++) {
    const char = criteria[index]!;
    if (char === '~' && index + 1 < criteria.length) {
      pattern += escapeRegex(criteria[++index]!);
    } else if (char === '*') {
      pattern += '.*';
    } else if (char === '?') {
      pattern += '.';
    } else {
      pattern += escapeRegex(char);
    }
  }
  return new RegExp(`^${pattern}$`, 'i').test(toText(value));
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function truthy(value: SpreadsheetFormulaValue): boolean {
  if (Array.isArray(value)) return value.flat().some(truthy);
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') return value.length > 0;
  return false;
}

function compareValues(
  left: SpreadsheetFormulaValue,
  right: SpreadsheetFormulaValue,
  operator: string,
): boolean {
  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  const comparableLeft = leftNumber ?? String(left ?? '');
  const comparableRight = rightNumber ?? String(right ?? '');
  switch (operator) {
    case '=':
      return comparableLeft === comparableRight;
    case '<>':
      return comparableLeft !== comparableRight;
    case '>':
      return comparableLeft > comparableRight;
    case '>=':
      return comparableLeft >= comparableRight;
    case '<':
      return comparableLeft < comparableRight;
    case '<=':
      return comparableLeft <= comparableRight;
    default:
      return false;
  }
}

const MS_PER_DAY = 86_400_000;

interface FormulaDateParts {
  year: number;
  month: number;
  day: number;
  time: number;
}

function toDateParts(value: SpreadsheetFormulaValue | undefined): FormulaDateParts | null {
  if (Array.isArray(value) || value === null || value === '') return null;
  if (typeof value === 'number' && Number.isFinite(value)) {
    const time = Date.UTC(1899, 11, 30) + value * MS_PER_DAY;
    return datePartsFromTime(time);
  }
  if (typeof value !== 'string') return null;
  const match = /^\s*(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})\s*$/.exec(value);
  const time = match
    ? Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]))
    : Date.parse(value);
  return Number.isFinite(time) ? datePartsFromTime(time) : null;
}

function datePartsFromTime(time: number): FormulaDateParts | null {
  const date = new Date(time);
  if (!Number.isFinite(date.getTime())) return null;
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate(), time };
}

function excelSerialFromDate(date: Date): number {
  const utcMidnight = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.floor((utcMidnight - Date.UTC(1899, 11, 30)) / MS_PER_DAY);
}
