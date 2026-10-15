import {
  ExcelWorkbook,
  ExcelWorkbookCore,
  createShadowWorkbookModel,
  planExcelWriteback,
  type ExcelCellValue,
  type ExcelComment,
  type ExcelConditionalFormatRule,
  type ExcelHyperlink,
  type ExcelNamedItem,
  type ExcelRange,
  type ExcelSheetModel,
  type ExcelTable,
  type ExcelWorkbookMutation,
  type ExcelWorkbookModel,
  type ExcelWorksheet,
} from '@coco-vexel/coco-office/excel';

const sheet: ExcelSheetModel = {
  id: '1',
  name: 'Sheet1',
  dimension: 'A1:B2',
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
  cells: [],
};

const model: ExcelWorkbookModel = {
  sheets: [sheet],
  sharedStrings: [],
  definedNames: [],
  styles: {
    numberFormats: [],
    fonts: [],
    fills: [],
    borders: [],
    cellFormats: [],
  },
};

async function smoke(): Promise<void> {
  ExcelWorkbook.setDefaultMode('inline');
  const workbook = await ExcelWorkbook.fromModel(model, { workbookId: 'public-api-smoke' });
  const worksheets = await workbook.worksheets.items();
  const worksheet: ExcelWorksheet = workbook.worksheets.getItem('Sheet1');
  const maybeWorksheet: ExcelWorksheet | null =
    await workbook.worksheets.getItemOrNullObject('Sheet1');
  const addedWorksheet: ExcelWorksheet = await workbook.worksheets.add('Added');
  await addedWorksheet.setVisibility('hidden');
  const copiedWorksheet: ExcelWorksheet = await workbook.worksheets.copy('Sheet1', 'Sheet Copy');
  await workbook.worksheets.move('Added', 0);
  await workbook.worksheets.deleteItem('Sheet Copy');
  const range: ExcelRange = worksheet.getRange('A1:B2');

  await range.setValues([
    ['Name', 'Amount'],
    ['alpha', 42],
  ]);
  const values: ExcelCellValue[][] = await range.values();
  await range.setStyleId(1);
  const styleIds: (number | undefined)[][] = await range.styleIds();
  await range.setRowHeight(22);
  await range.setColumnWidth(16);
  const rowHeights: Array<number | undefined> = await range.rowHeights();
  const columnWidths: Array<number | undefined> = await range.columnWidths();
  await range.merge();
  await range.unmerge();
  await range.format.font.setBold(true);
  await range.format.font.setItalic(true);
  await range.format.font.setUnderline(true);
  await range.format.font.setName('Calibri');
  await range.format.font.setSize(12);
  await range.format.font.setColor({ rgb: '#FF0000' });
  await range.format.fill.setPatternType('solid');
  await range.format.fill.setForegroundColor({ rgb: '#FFFF00' });
  await range.format.fill.setBackgroundColor({ rgb: '#000000' });
  await range.format.border.setBottom({ style: 'thin', color: { rgb: '#000000' } });
  await range.format.alignment.setHorizontal('center');
  await range.format.alignment.setVertical('center');
  await range.format.alignment.setWrapText(true);
  await range.format.setNumberFormatCode('0.00%');
  const fontBold: boolean | undefined = await range.format.font.bold();
  const numberFormatCode: string | undefined = await range.format.numberFormatCode();
  const cellValue: ExcelCellValue = await worksheet.getCell('B2').value();
  await worksheet.getCell('C2').setFormula('B2*2');
  await worksheet.getCell('C2').setStyleId(1);
  const cellStyleId: number | undefined = await worksheet.getCell('C2').styleId();

  const namedItem: ExcelNamedItem = await workbook.names.add({
    name: 'AddedName',
    text: 'Sheet1!$A$1',
  });
  await namedItem.setText('Sheet1!$A$1:$B$2');
  await namedItem.rename('RenamedName');
  await namedItem.setLocalSheetId(0);
  await workbook.names.deleteItem('RenamedName');

  const table: ExcelTable = await worksheet.tables.add({
    name: 'Table1',
    displayName: 'Table1',
    ref: 'A1:B2',
    columns: [
      { id: 1, name: 'Name' },
      { id: 2, name: 'Amount' },
    ],
  });
  await table.setDisplayName('TableRenamed');
  await table.resize('A1:C3');
  await table.setHeaderRowCount(1);
  await table.setTotalsRowCount(1);
  await table.setTotalsRowShown(true);
  await table.setAutoFilter({
    ref: 'A1:C3',
    filterColumns: [{ colId: 1, filters: ['42'] }],
    sortState: { ref: 'A2:C3', conditions: [{ ref: 'B2:B3', descending: true }] },
  });
  await table.setSortState({ ref: 'A2:C3', conditions: [{ ref: 'B2:B3' }] });
  const autoFilter = await table.autoFilter();
  const sortState = await table.sortState();
  await table.style.setShowRowStripes(true);
  await table.columns.getItem('Amount').setTotalsRowFunction('sum');
  await table.columns.getItem('Column3').setTotalsRowLabel('total');
  await table.columns.add({ id: 4, name: 'Temp' });
  await table.columns.deleteItem('Temp');
  const writebackPlan = planExcelWriteback(model, await workbook.toModel(), {
    canPatchCells: true,
    canInsertShadowSheets: true,
    canCopySheetContent: true,
    canDeleteSheet: true,
    canRenameSheet: true,
  });
  const firstCellChangeAddress: string | undefined =
    writebackPlan.operations[0]?.type === 'patchCells'
      ? writebackPlan.operations[0].changes[0]?.address
      : undefined;
  const firstShadowTargetName: string | undefined =
    writebackPlan.operations[0]?.type === 'insertShadowSheet'
      ? writebackPlan.operations[0].targetSheet.name
      : undefined;
  const shadowModel = createShadowWorkbookModel(sheet, {
    sheetName: '__coco_shadow_1',
    styles: model.styles,
  });

  const hyperlink: ExcelHyperlink = await worksheet.hyperlinks.add({
    ref: 'A2',
    target: 'https://example.com',
    display: 'alpha',
  });
  await hyperlink.setDisplay('Alpha');

  const comment: ExcelComment = await worksheet.comments.add({ ref: 'B2', text: 'checked' });
  await comment.setText('updated');

  const validation = await worksheet.dataValidations.add({
    sqref: 'B2',
    type: 'whole',
    formula1: '1',
    formula2: '100',
  });
  await validation.setAllowBlank(true);

  const format = await worksheet.conditionalFormats.add({
    sqref: 'B2',
    rules: [{ type: 'cellIs', priority: 1, formulas: ['10'] }],
  });
  const rule: ExcelConditionalFormatRule = format.rules.getItem(0);
  await rule.setPriority(2);

  const workbookModel: ExcelWorkbookModel = await workbook.toModel();
  const mutations: ExcelWorkbookMutation[] = await workbook.getMutations();
  await workbook.clearMutations();
  const bytes: Uint8Array = await workbook.toZip();
  const core: ExcelWorkbookCore = await ExcelWorkbookCore.fromZip(bytes);

  await workbook.dispose();

  void worksheets;
  void maybeWorksheet;
  void copiedWorksheet;
  void values;
  void styleIds;
  void rowHeights;
  void columnWidths;
  void fontBold;
  void numberFormatCode;
  void cellValue;
  void cellStyleId;
  void autoFilter;
  void sortState;
  void writebackPlan;
  void firstCellChangeAddress;
  void firstShadowTargetName;
  void shadowModel;
  void workbookModel;
  void mutations;
  void core;
}

void smoke();

// Public API guardrails: getters that cross the runtime are async.
// @ts-expect-error values() returns Promise<ExcelCellValue[][]>, not a matrix directly.
const invalidValues: ExcelCellValue[][] = new Promise<ExcelWorkbook>((resolve) => {
  void resolve;
}) as unknown as ReturnType<ExcelRange['values']>;

// getItem returns a synchronous handle; existence checks use getItemOrNullObject.
const handleIsSync: ExcelWorksheet = {} as ReturnType<ExcelWorkbook['worksheets']['getItem']>;
void handleIsSync;
