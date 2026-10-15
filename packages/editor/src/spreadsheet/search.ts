import type { ExcelCellValue } from '@coco-vexel/coco-office/excel';
import type { SpreadsheetRange, SpreadsheetRenderCell, SpreadsheetSheetRenderModel } from './types.js';

export interface SpreadsheetFindOptions {
  query: string;
  matchCase?: boolean;
}

export interface SpreadsheetFindMatch {
  row: number;
  column: number;
  address: string;
  text: string;
}

export interface SpreadsheetReplaceTarget {
  address: string;
  value?: ExcelCellValue;
  formula?: string;
}

export function buildSpreadsheetReplacePlan(
  model: SpreadsheetSheetRenderModel,
  options: SpreadsheetFindOptions & { replacement: string },
): SpreadsheetReplaceTarget[] {
  const query = options.query.trim();
  if (!query) return [];
  const matches: SpreadsheetReplaceTarget[] = [];
  const cells = [...model.cellsByAddress.values()].sort((a, b) => a.row - b.row || a.column - b.column);
  for (const cell of cells) {
    if (!searchableCellText(cell, options.matchCase === true).includes(normalizeSearchText(query, options.matchCase === true))) continue;
    if (cell.formula) {
      const formulaQuery = query.startsWith('=') ? query.slice(1) : query;
      matches.push({ address: cell.address, formula: replaceText(cell.formula, formulaQuery, options.replacement, options.matchCase === true) });
    } else {
      const source = cell.value === null ? '' : String(cell.value);
      matches.push({ address: cell.address, value: replaceText(source, query, options.replacement, options.matchCase === true) });
    }
  }
  return matches;
}

export function findSpreadsheetMatches(
  model: SpreadsheetSheetRenderModel,
  options: SpreadsheetFindOptions,
): SpreadsheetFindMatch[] {
  const query = normalizeSearchText(options.query, options.matchCase === true);
  if (!query) return [];
  const matches: SpreadsheetFindMatch[] = [];
  const cells = [...model.cellsByAddress.values()].sort((a, b) => a.row - b.row || a.column - b.column);
  for (const cell of cells) {
    const haystack = searchableCellText(cell, options.matchCase === true);
    if (!haystack.includes(query)) continue;
    matches.push({ row: cell.row, column: cell.column, address: cell.address, text: cell.displayText });
  }
  return matches;
}

export function nextSpreadsheetMatch(
  matches: readonly SpreadsheetFindMatch[],
  active: { row: number; column: number },
  direction: 'forward' | 'backward' = 'forward',
): SpreadsheetFindMatch | null {
  if (matches.length === 0) return null;
  if (direction === 'backward') {
    return [...matches].reverse().find((match) => compareCellPosition(match, active) < 0) ?? matches[matches.length - 1] ?? null;
  }
  return matches.find((match) => compareCellPosition(match, active) > 0) ?? matches[0] ?? null;
}

export function rangeForSpreadsheetMatch(match: SpreadsheetFindMatch): SpreadsheetRange {
  return {
    startRow: match.row,
    startColumn: match.column,
    endRow: match.row,
    endColumn: match.column,
  };
}

function searchableCellText(cell: SpreadsheetRenderCell, matchCase: boolean): string {
  const parts = [
    cell.displayText,
    rawCellValueText(cell.value),
    cell.formula ? `=${cell.formula}` : '',
  ].filter(Boolean);
  return normalizeSearchText(parts.join('\n'), matchCase);
}

function rawCellValueText(value: ExcelCellValue): string {
  return value === null ? '' : String(value);
}

function normalizeSearchText(value: string, matchCase: boolean): string {
  const text = value.trim();
  return matchCase ? text : text.toLocaleLowerCase();
}

function compareCellPosition(
  a: { row: number; column: number },
  b: { row: number; column: number },
): number {
  return a.row === b.row ? a.column - b.column : a.row - b.row;
}

function replaceText(source: string, query: string, replacement: string, matchCase: boolean): string {
  if (!query) return source;
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return source.replace(new RegExp(escaped, matchCase ? 'g' : 'gi'), replacement);
}
