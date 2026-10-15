#!/usr/bin/env node
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const distPath = resolve(here, '../dist/excel.js');
const outputPath = resolve(here, '../test-fixtures/excel/formula-business-regression.xlsx');

const { ExcelWorkbook } = await import(pathToFileURL(distPath).href);

const textCell = (address, row, column, value) => ({ address, row, column, kind: 'string', value });
const numberCell = (address, row, column, value) => ({ address, row, column, kind: 'number', value });
const formulaCell = (address, row, column, formula) => ({ address, row, column, kind: 'formula', value: null, formula });
const sheet = (id, name, cells) => ({
  id,
  name,
  dimension: 'A1:L8',
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
});

const workbook = await ExcelWorkbook.fromModel({
  sharedStrings: [],
  definedNames: [],
  styles: { numberFormats: [], fonts: [], fills: [], borders: [], cellFormats: [] },
  sheets: [
    sheet('1', '销售订单', [
      textCell('A1', 1, 1, '合同ID'), textCell('C1', 1, 3, '客户ID'), textCell('D1', 1, 4, '合同金额'), textCell('G1', 1, 7, '姓名'),
      textCell('A2', 2, 1, 'SO001'), textCell('C2', 2, 3, 'C001'), numberCell('D2', 2, 4, 150000), textCell('G2', 2, 7, '张三'),
      textCell('A3', 3, 1, 'SO002'), textCell('C3', 3, 3, 'C002'), numberCell('D3', 3, 4, 89000), textCell('G3', 3, 7, '李四'),
      textCell('A4', 4, 1, 'SO003'), textCell('C4', 4, 3, 'C003'), numberCell('D4', 4, 4, 215000), textCell('G4', 4, 7, '王五'),
    ]),
    sheet('2', '收款单', [
      textCell('A1', 1, 1, '收款ID'), textCell('B1', 1, 2, '合同ID'), textCell('D1', 1, 4, '收款金额'),
      textCell('A2', 2, 1, 'R001'), textCell('B2', 2, 2, 'SO001'), numberCell('D2', 2, 4, 150000),
      textCell('A3', 3, 1, 'R002'), textCell('B3', 3, 2, 'SO002'), numberCell('D3', 3, 4, 30000),
      textCell('A4', 4, 1, 'R003'), textCell('B4', 4, 2, 'SO003'), numberCell('D4', 4, 4, 0),
    ]),
    sheet('3', '员工', [
      textCell('A1', 1, 1, '员工ID'), textCell('C1', 1, 3, '姓名'), textCell('G1', 1, 7, '生日'), textCell('H1', 1, 8, '年龄'), textCell('K1', 1, 11, '销售额'), textCell('L1', 1, 12, '提成'),
      textCell('A2', 2, 1, 'E001'), textCell('C2', 2, 3, '张三'), numberCell('G2', 2, 7, 32874), formulaCell('H2', 2, 8, 'DATEDIF(G2,TODAY(),"Y")'), formulaCell('K2', 2, 11, "SUMPRODUCT(('销售订单'!G$2:G$4=C2)*('销售订单'!D$2:D$4))"), formulaCell('L2', 2, 12, 'IF(K2<=100000,K2*0.05,IF(K2<=300000,K2*0.08,K2*0.12))'),
      textCell('A3', 3, 1, 'E002'), textCell('C3', 3, 3, '李四'), numberCell('G3', 3, 7, 35600), formulaCell('H3', 3, 8, 'DATEDIF(G3,TODAY(),"Y")'), formulaCell('K3', 3, 11, "SUMPRODUCT(('销售订单'!G$2:G$4=C3)*('销售订单'!D$2:D$4))"), formulaCell('L3', 3, 12, 'IF(K3<=100000,K3*0.05,IF(K3<=300000,K3*0.08,K3*0.12))'),
      textCell('A4', 4, 1, 'E003'), textCell('C4', 4, 3, '王五'), numberCell('G4', 4, 7, 36526), formulaCell('H4', 4, 8, 'DATEDIF(G4,TODAY(),"Y")'), formulaCell('K4', 4, 11, "SUMPRODUCT(('销售订单'!G$2:G$4=C4)*('销售订单'!D$2:D$4))"), formulaCell('L4', 4, 12, 'IF(K4<=100000,K4*0.05,IF(K4<=300000,K4*0.08,K4*0.12))'),
    ]),
  ],
}, { mode: 'inline', workbookId: 'formula-business-regression' });

mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, await workbook.toZip());
console.log(`Wrote ${outputPath}`);
