import type {
  ExcelAlignmentModel,
  ExcelBorderModel,
  ExcelCellValue,
  ExcelColorModel,
  ExcelFontModel,
  ExcelSheetViewModel,
} from './model.js';
import type {
  ExcelWorksheetLineOptions,
  ExcelWorksheetPrintSetupOptions,
} from './worksheet-layout.js';

export type ExcelWorkbookMutation =
  | ExcelRangeSetValuesMutation
  | ExcelRangeSetFormulasMutation
  | ExcelRangeClearMutation
  | ExcelRangeMergeMutation
  | ExcelRangeUnmergeMutation
  | ExcelRangeSetRowHeightMutation
  | ExcelRangeSetColumnWidthMutation
  | ExcelRangeSetFontColorMutation
  | ExcelRangeSetFontMutation
  | ExcelRangeSetAlignmentMutation
  | ExcelRangeSetFillForegroundColorMutation
  | ExcelRangeSetBorderMutation
  | ExcelWorksheetAddMutation
  | ExcelWorksheetDeleteMutation
  | ExcelWorksheetRenameMutation
  | ExcelWorksheetAddLinesMutation
  | ExcelWorksheetSetPrintSetupMutation
  | ExcelWorksheetSetSheetViewsMutation
  | ExcelUnsafeMutation;

export interface ExcelMutationSheetRef {
  sheetId?: string;
  sheetName: string;
}

export interface ExcelRangeMutationBase extends ExcelMutationSheetRef {
  address: string;
}

export interface ExcelRangeSetValuesMutation extends ExcelRangeMutationBase {
  kind: 'range.setValues';
  values: ExcelCellValue[][];
}

export interface ExcelRangeSetFormulasMutation extends ExcelRangeMutationBase {
  kind: 'range.setFormulas';
  formulas: (string | null)[][];
}

export interface ExcelRangeClearMutation extends ExcelRangeMutationBase {
  kind: 'range.clear';
}

export interface ExcelRangeMergeMutation extends ExcelRangeMutationBase {
  kind: 'range.merge';
}

export interface ExcelRangeUnmergeMutation extends ExcelRangeMutationBase {
  kind: 'range.unmerge';
}

export interface ExcelRangeSetRowHeightMutation extends ExcelRangeMutationBase {
  kind: 'range.setRowHeight';
  height: number | undefined;
}

export interface ExcelRangeSetColumnWidthMutation extends ExcelRangeMutationBase {
  kind: 'range.setColumnWidth';
  width: number | undefined;
}

export interface ExcelRangeSetFontColorMutation extends ExcelRangeMutationBase {
  kind: 'range.format.font.setColor';
  color: ExcelColorModel | undefined;
}

export interface ExcelRangeSetFontMutation extends ExcelRangeMutationBase {
  kind: 'range.format.font.setValue';
  font: ExcelFontModel;
}

export interface ExcelRangeSetAlignmentMutation extends ExcelRangeMutationBase {
  kind: 'range.format.alignment.setValue';
  alignment: ExcelAlignmentModel | undefined;
}

export interface ExcelRangeSetFillForegroundColorMutation extends ExcelRangeMutationBase {
  kind: 'range.format.fill.setForegroundColor';
  color: ExcelColorModel | undefined;
}

export interface ExcelRangeSetBorderMutation extends ExcelRangeMutationBase {
  kind: 'range.format.border.setValue';
  border: ExcelBorderModel;
}

export interface ExcelWorksheetAddMutation extends ExcelMutationSheetRef {
  kind: 'worksheet.add';
}

export interface ExcelWorksheetDeleteMutation extends ExcelMutationSheetRef {
  kind: 'worksheet.delete';
}

export interface ExcelWorksheetRenameMutation extends ExcelMutationSheetRef {
  kind: 'worksheet.rename';
  oldName: string;
  newName: string;
}

export interface ExcelWorksheetAddLinesMutation extends ExcelMutationSheetRef {
  kind: 'worksheet.addLines';
  lines: ExcelWorksheetLineOptions[];
}

export interface ExcelWorksheetSetPrintSetupMutation extends ExcelMutationSheetRef {
  kind: 'worksheet.setPrintSetup';
  options: ExcelWorksheetPrintSetupOptions;
}

export interface ExcelWorksheetSetSheetViewsMutation extends ExcelMutationSheetRef {
  kind: 'worksheet.setSheetViews';
  sheetViews: ExcelSheetViewModel[];
}

export interface ExcelUnsafeMutation {
  kind: 'unsafe';
  scope: 'workbook' | 'worksheet' | 'range' | 'object';
  sheetId?: string;
  sheetName?: string;
  reason: string;
}

export type ExcelMutationRecorder = (mutation: ExcelWorkbookMutation) => void;

export class ExcelMutationLog {
  private readonly items: ExcelWorkbookMutation[] = [];

  record(mutation: ExcelWorkbookMutation): void {
    this.items.push(cloneMutation(mutation));
  }

  list(): ExcelWorkbookMutation[] {
    return this.items.map(cloneMutation);
  }

  clear(): void {
    this.items.length = 0;
  }

  hasUnsafe(): boolean {
    return this.items.some((item) => item.kind === 'unsafe');
  }
}

export function cloneMutation(mutation: ExcelWorkbookMutation): ExcelWorkbookMutation {
  return JSON.parse(JSON.stringify(mutation)) as ExcelWorkbookMutation;
}

export function changedSheetNamesFromMutations(mutations: readonly ExcelWorkbookMutation[]): string[] {
  const names = new Set<string>();
  for (const mutation of mutations) {
    if (mutation.kind === 'worksheet.rename') {
      names.add(mutation.oldName);
      names.add(mutation.newName);
      continue;
    }
    if ('sheetName' in mutation && mutation.sheetName) {
      names.add(mutation.sheetName);
    }
  }
  return [...names];
}
