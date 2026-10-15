#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2).filter((arg) => arg !== '--');
const json = args.includes('--json');
const check = args.includes('--check');
const requireFormulas = args.includes('--require-formulas');
const inputDirs = args.filter((arg) => !arg.startsWith('--'));
const fixtureDirs = inputDirs.length
  ? inputDirs
  : (process.env.COCO_EXCEL_FIXTURE_DIR ?? '').split(';').filter(Boolean);

async function main() {
  const distPath = resolve(here, '../dist/excel.js');
  if (!existsSync(distPath)) {
    throw new Error(
      `Missing ${distPath}. Run "corepack pnpm --filter @coco-vexel/coco-office build" first.`,
    );
  }
  if (!fixtureDirs.length) {
    throw new Error('Pass one or more Excel fixture directories, or set COCO_EXCEL_FIXTURE_DIR.');
  }

  const { ExcelWorkbook, calculateSpreadsheetSheetFormulas } = await import(pathToFileURL(distPath).href);
  const files = [...new Set(fixtureDirs.flatMap(findExcelFiles))].sort((left, right) => left.localeCompare(right));
  const reports = [];
  for (const filePath of files) {
    reports.push(await inspectWorkbook(filePath, ExcelWorkbook, calculateSpreadsheetSheetFormulas));
  }

const report = { generatedAt: new Date().toISOString(), fixtureDirs, files: reports };
  if (json) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    printMarkdownReport(report);
  }
  if (check) validateReport(report, { requireFormulas });
}

function findExcelFiles(directory) {
  const resolved = resolve(directory);
  if (!existsSync(resolved)) throw new Error(`Excel fixture directory not found: ${resolved}`);
  const stat = statSync(resolved);
  if (stat.isFile()) return resolved.toLowerCase().endsWith('.xlsx') ? [resolved] : [];
  const files = [];
  for (const entry of readdirSync(resolved, { withFileTypes: true })) {
    if (entry.name.startsWith('~$')) continue;
    const path = join(resolved, entry.name);
    if (entry.isDirectory()) files.push(...findExcelFiles(path));
    else if (entry.isFile() && entry.name.toLowerCase().endsWith('.xlsx')) files.push(path);
  }
  return files;
}

async function inspectWorkbook(filePath, ExcelWorkbook, calculateSpreadsheetSheetFormulas) {
  const workbook = await ExcelWorkbook.fromZip(readFileSync(filePath), { mode: 'inline' });
  const model = await workbook.toModel();
  const functions = Object.create(null);
  const errors = Object.create(null);
  const unresolved = [];
  let formulaCells = 0;
  let cachedFormulaValues = 0;
  let comparedFormulaValues = 0;
  let mismatchedFormulaValues = 0;

  for (const sheet of model.sheets) {
    const results = calculateSpreadsheetSheetFormulas(sheet, { sheets: model.sheets });
    for (const cell of sheet.cells) {
      if (!cell.formula) continue;
      formulaCells++;
      for (const name of extractFunctionNames(cell.formula)) functions[name] = (functions[name] ?? 0) + 1;
      const result = results.get(cell.address);
      if (!result) continue;
      if (result.error) {
        errors[result.error] = (errors[result.error] ?? 0) + 1;
        if (result.error === '#NAME?') unresolved.push({ sheet: sheet.name, address: cell.address, formula: cell.formula });
      }
      if (cell.value !== null && cell.value !== undefined) {
        cachedFormulaValues++;
        if (!result.error && !Array.isArray(result.value)) {
          comparedFormulaValues++;
          if (!sameFormulaValue(cell.value, result.value)) mismatchedFormulaValues++;
        }
      }
    }
  }

  return {
    file: filePath,
    sheets: model.sheets.length,
    formulaCells,
    functions: sortCounts(functions),
    errors: sortCounts(errors),
    unresolved,
    cachedFormulaValues,
    comparedFormulaValues,
    mismatchedFormulaValues,
  };
}

function extractFunctionNames(formula) {
  return [...formula.matchAll(/\b([A-Z][A-Z0-9_.]*)\s*\(/gi)].map((match) => match[1].toUpperCase());
}

function sortCounts(counts) {
  return Object.fromEntries(Object.entries(counts).sort(([left], [right]) => left.localeCompare(right)));
}

function sameFormulaValue(left, right) {
  if (typeof left === 'number' && typeof right === 'number') return Math.abs(left - right) < 1e-7;
  return left === right;
}

function printMarkdownReport(report) {
  const totals = summarizeReport(report);

  console.log('# Excel Formula Coverage Report');
  console.log();
  console.log(`Scanned files: ${totals.files}`);
  console.log(`Formula cells: ${totals.formulaCells}`);
  console.log(`Formula cache values: ${totals.cachedFormulaValues}`);
  console.log(`Compared cache values: ${totals.comparedFormulaValues}`);
  console.log(`Mismatches: ${totals.mismatchedFormulaValues}`);
  console.log();
  console.log('## Function Frequency');
  console.log();
  console.log('| Function | Count |');
  console.log('| --- | ---: |');
  for (const [name, count] of Object.entries(sortCounts(totals.functions))) console.log(`| ${name} | ${count} |`);
  console.log();
  console.log('## Errors');
  console.log();
  if (!Object.keys(totals.errors).length) console.log('No formula errors.');
  else for (const [name, count] of Object.entries(sortCounts(totals.errors))) console.log(`- ${name}: ${count}`);
  if (totals.unresolved.length) {
    console.log();
    console.log('## Unresolved Formulas');
    console.log();
    for (const item of totals.unresolved) console.log(`- ${basename(item.file)} / ${item.sheet}!${item.address}: \`${item.formula}\``);
  }
  console.log();
  console.log('## Files');
  console.log();
  console.log('| File | Formula cells | Errors | Cached values | Mismatches |');
  console.log('| --- | ---: | ---: | ---: | ---: |');
  for (const file of report.files) {
    const errorCount = Object.values(file.errors).reduce((sum, count) => sum + count, 0);
    console.log(`| ${basename(file.file)} | ${file.formulaCells} | ${errorCount} | ${file.cachedFormulaValues} | ${file.mismatchedFormulaValues} |`);
  }
}

function summarizeReport(report) {
  return report.files.reduce((result, file) => {
    result.files++;
    result.formulaCells += file.formulaCells;
    result.cachedFormulaValues += file.cachedFormulaValues;
    result.comparedFormulaValues += file.comparedFormulaValues;
    result.mismatchedFormulaValues += file.mismatchedFormulaValues;
    for (const [name, count] of Object.entries(file.functions)) result.functions[name] = (result.functions[name] ?? 0) + count;
    for (const [name, count] of Object.entries(file.errors)) result.errors[name] = (result.errors[name] ?? 0) + count;
    result.unresolved.push(...file.unresolved.map((item) => ({ ...item, file: file.file })));
    return result;
  }, { files: 0, formulaCells: 0, cachedFormulaValues: 0, comparedFormulaValues: 0, mismatchedFormulaValues: 0, functions: {}, errors: {}, unresolved: [] });
}

function validateReport(report, options) {
  const totals = summarizeReport(report);
  const failures = [];
  if (!totals.files) failures.push('no Excel files were scanned');
  if (options.requireFormulas && totals.formulaCells === 0) failures.push('no formula cells were found');
  if (Object.keys(totals.errors).length) failures.push(`formula errors: ${Object.entries(totals.errors).map(([name, count]) => `${name} (${count})`).join(', ')}`);
  if (totals.unresolved.length) failures.push(`unresolved formulas: ${totals.unresolved.length}`);
  if (totals.mismatchedFormulaValues) failures.push(`cached-value mismatches: ${totals.mismatchedFormulaValues}`);
  if (failures.length) throw new Error(`Formula coverage check failed: ${failures.join('; ')}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
