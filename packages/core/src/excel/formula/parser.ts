import { formatCellAddress, parseCellAddress } from '../address.js';
import type { SpreadsheetFormulaAst, SpreadsheetFormulaReference } from './types.js';

type Token =
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }
  | { kind: 'identifier'; value: string }
  | { kind: 'cell'; value: string; sheetName?: string }
  | { kind: 'operator'; value: string }
  | { kind: 'paren'; value: '(' | ')' }
  | { kind: 'comma'; value: ',' }
  | { kind: 'colon'; value: ':' }
  | { kind: 'eof' };

export function parseSpreadsheetFormula(formula: string): SpreadsheetFormulaAst {
  const parser = new FormulaParser(tokenizeFormula(stripFormulaPrefix(formula)));
  return parser.parse();
}

export function collectSpreadsheetFormulaReferences(
  formula: string,
): SpreadsheetFormulaReference[] {
  const references: SpreadsheetFormulaReference[] = [];
  try {
    collectReferences(parseSpreadsheetFormula(formula), references);
  } catch {
    return [];
  }
  return dedupeReferences(references);
}

function collectReferences(
  ast: SpreadsheetFormulaAst,
  output: SpreadsheetFormulaReference[],
): void {
  switch (ast.kind) {
    case 'cell':
      output.push({ kind: 'cell', address: ast.address, sheetName: ast.sheetName });
      return;
    case 'range':
      output.push({
        kind: 'range',
        startAddress: ast.startAddress,
        endAddress: ast.endAddress,
        sheetName: ast.sheetName,
      });
      return;
    case 'unary':
      collectReferences(ast.argument, output);
      return;
    case 'binary':
      collectReferences(ast.left, output);
      collectReferences(ast.right, output);
      return;
    case 'call':
      ast.args.forEach((arg) => collectReferences(arg, output));
      return;
    case 'number':
    case 'string':
    case 'name':
      return;
  }
}

function dedupeReferences(
  references: SpreadsheetFormulaReference[],
): SpreadsheetFormulaReference[] {
  const seen = new Set<string>();
  return references.filter((reference) => {
    const key =
      reference.kind === 'cell'
        ? `cell:${reference.sheetName ?? ''}:${reference.address}`
        : `range:${reference.sheetName ?? ''}:${reference.startAddress}:${reference.endAddress}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function stripFormulaPrefix(formula: string): string {
  return formula.trim().startsWith('=') ? formula.trim().slice(1) : formula.trim();
}

function tokenizeFormula(input: string): Token[] {
  const tokens: Token[] = [];
  let index = 0;
  while (index < input.length) {
    const ch = input[index]!;
    if (/\s/.test(ch)) {
      index++;
      continue;
    }
    if (/\d/.test(ch) || (ch === '.' && /\d/.test(input[index + 1] ?? ''))) {
      const start = index;
      index++;
      while (index < input.length && /[\d.]/.test(input[index]!)) index++;
      const value = Number(input.slice(start, index));
      if (!Number.isFinite(value)) throw new Error('Invalid number literal');
      tokens.push({ kind: 'number', value });
      continue;
    }
    if (ch === '"') {
      let value = '';
      index++;
      while (index < input.length) {
        const next = input[index]!;
        if (next === '"') {
          if (input[index + 1] === '"') {
            value += '"';
            index += 2;
            continue;
          }
          index++;
          break;
        }
        value += next;
        index++;
      }
      tokens.push({ kind: 'string', value });
      continue;
    }
    if (ch === "'") {
      const sheetStart = index + 1;
      let sheetName = '';
      index = sheetStart;
      while (index < input.length) {
        if (input[index] === "'") {
          if (input[index + 1] === "'") {
            sheetName += "'";
            index += 2;
            continue;
          }
          break;
        }
        sheetName += input[index]!;
        index++;
      }
      if (input[index] !== "'" || input[index + 1] !== '!') throw new Error('Invalid sheet reference');
      index += 2;
      const start = index;
      while (index < input.length && /[A-Za-z0-9_$]/.test(input[index]!)) index++;
      const cell = normalizeCellReference(input.slice(start, index));
      if (!cell) throw new Error('Invalid sheet cell reference');
      tokens.push({ kind: 'cell', value: cell, sheetName });
      continue;
    }
    if (ch === '(' || ch === ')') {
      tokens.push({ kind: 'paren', value: ch });
      index++;
      continue;
    }
    if (ch === ',') {
      tokens.push({ kind: 'comma', value: ch });
      index++;
      continue;
    }
    if (ch === ':') {
      tokens.push({ kind: 'colon', value: ch });
      index++;
      continue;
    }
    const twoChar = input.slice(index, index + 2);
    if (twoChar === '>=' || twoChar === '<=' || twoChar === '<>') {
      tokens.push({ kind: 'operator', value: twoChar });
      index += 2;
      continue;
    }
    if ('+-*/&%=><'.includes(ch)) {
      tokens.push({ kind: 'operator', value: ch });
      index++;
      continue;
    }
    if (/[A-Za-z_$]/.test(ch)) {
      const start = index;
      index++;
      while (index < input.length && /[A-Za-z0-9_$.]/.test(input[index]!)) index++;
      const raw = input.slice(start, index);
      if (input[index] === '!') {
        index++;
        const cellStart = index;
        while (index < input.length && /[A-Za-z0-9_$]/.test(input[index]!)) index++;
        const cell = normalizeCellReference(input.slice(cellStart, index));
        if (!cell) throw new Error('Invalid sheet cell reference');
        tokens.push({ kind: 'cell', value: cell, sheetName: raw });
        continue;
      }
      const cell = normalizeCellReference(raw);
      tokens.push(cell ? { kind: 'cell', value: cell } : { kind: 'identifier', value: raw });
      continue;
    }
    throw new Error(`Unexpected token: ${ch}`);
  }
  tokens.push({ kind: 'eof' });
  return tokens;
}

class FormulaParser {
  private index = 0;

  constructor(private readonly tokens: Token[]) {}

  parse(): SpreadsheetFormulaAst {
    const expression = this.parseComparison();
    this.expect('eof');
    return expression;
  }

  private parseComparison(): SpreadsheetFormulaAst {
    let left = this.parseConcatenation();
    while (this.isOperator('=', '<>', '>', '>=', '<', '<=')) {
      const operator = this.advanceOperator() as '=' | '<>' | '>' | '>=' | '<' | '<=';
      left = { kind: 'binary', operator, left, right: this.parseConcatenation() };
    }
    return left;
  }

  private parseConcatenation(): SpreadsheetFormulaAst {
    let left = this.parseAdditive();
    while (this.isOperator('&')) {
      this.advanceOperator();
      left = { kind: 'binary', operator: '&', left, right: this.parseAdditive() };
    }
    return left;
  }

  private parseAdditive(): SpreadsheetFormulaAst {
    let left = this.parseMultiplicative();
    while (this.isOperator('+', '-')) {
      const operator = this.advanceOperator() as '+' | '-';
      left = { kind: 'binary', operator, left, right: this.parseMultiplicative() };
    }
    return left;
  }

  private parseMultiplicative(): SpreadsheetFormulaAst {
    let left = this.parseUnary();
    while (this.isOperator('*', '/')) {
      const operator = this.advanceOperator() as '*' | '/';
      left = { kind: 'binary', operator, left, right: this.parseUnary() };
    }
    return left;
  }

  private parseUnary(): SpreadsheetFormulaAst {
    if (this.isOperator('+', '-')) {
      const operator = this.advanceOperator() as '+' | '-';
      return { kind: 'unary', operator, argument: this.parseUnary() };
    }
    return this.parsePostfix();
  }

  private parsePostfix(): SpreadsheetFormulaAst {
    let expression = this.parsePrimary();
    while (this.isOperator('%')) {
      this.advanceOperator();
      expression = { kind: 'unary', operator: '%', argument: expression };
    }
    return expression;
  }

  private parsePrimary(): SpreadsheetFormulaAst {
    const token = this.advance();
    if (token.kind === 'number') return { kind: 'number', value: token.value };
    if (token.kind === 'string') return { kind: 'string', value: token.value };
    if (token.kind === 'cell') {
      if (this.peek().kind === 'colon') {
        this.advance();
        const end = this.advance();
        if (end.kind !== 'cell') throw new Error('Range reference must end with a cell');
        if (end.sheetName && end.sheetName !== token.sheetName) throw new Error('Range sheets must match');
        return {
          kind: 'range',
          startAddress: token.value,
          endAddress: end.value,
          sheetName: token.sheetName,
        };
      }
      return { kind: 'cell', address: token.value, sheetName: token.sheetName };
    }
    if (token.kind === 'identifier') {
      return this.isParen('(') ? this.parseCall(token.value) : { kind: 'name', name: token.value };
    }
    if (token.kind === 'paren' && token.value === '(') {
      const expression = this.parseComparison();
      this.expect('paren', ')');
      return expression;
    }
    throw new Error('Expected expression');
  }

  private parseCall(name: string): SpreadsheetFormulaAst {
    this.expect('paren', '(');
    const args: SpreadsheetFormulaAst[] = [];
    if (!this.isParen(')')) {
      do {
        args.push(this.parseComparison());
      } while (this.consume('comma'));
    }
    this.expect('paren', ')');
    return { kind: 'call', name: name.toUpperCase(), args };
  }

  private isOperator(...operators: string[]): boolean {
    const token = this.peek();
    return token.kind === 'operator' && operators.includes(token.value);
  }

  private isParen(value: '(' | ')'): boolean {
    const token = this.peek();
    return token.kind === 'paren' && token.value === value;
  }

  private advanceOperator(): string {
    const token = this.advance();
    if (token.kind !== 'operator') throw new Error('Expected operator');
    return token.value;
  }

  private consume(kind: Token['kind']): boolean {
    if (this.peek().kind !== kind) return false;
    this.advance();
    return true;
  }

  private expect(kind: Token['kind'], value?: string): Token {
    const token = this.advance();
    if (token.kind !== kind || (value !== undefined && 'value' in token && token.value !== value)) {
      throw new Error(`Expected ${kind}`);
    }
    return token;
  }

  private peek(): Token {
    return this.tokens[this.index] ?? { kind: 'eof' };
  }

  private advance(): Token {
    return this.tokens[this.index++] ?? { kind: 'eof' };
  }
}

function normalizeCellReference(raw: string): string | null {
  try {
    const parsed = parseCellAddress(raw.replace(/\$/g, ''));
    return formatCellAddress(parsed.row, parsed.column);
  } catch {
    return null;
  }
}
