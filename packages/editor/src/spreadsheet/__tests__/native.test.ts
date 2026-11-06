// @vitest-environment happy-dom
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { captureExcelCellStyle, ExcelWorkbook } from '@coco-vexel/coco-office/excel';
import { SPREADSHEET_CLIPBOARD_MIME } from '../clipboard.js';
import {
  copybookGuideAppearance,
  renderSpreadsheetEditor,
  type SpreadsheetEditor,
  type SpreadsheetEditorOptions,
} from '../native.js';
import { SpreadsheetCanvasRenderer } from '../canvas-renderer.js';
import { SpreadsheetLayout } from '../layout.js';

const editors: SpreadsheetEditor[] = [];

test('recognizes Tianzi and Pinyin guide shapes used by copybook workbooks', () => {
  const shape = (name: string) => ({
    kind: 'twoCell' as const,
    object: { kind: 'shape' as const, name },
  });
  expect(copybookGuideAppearance(shape('coco-tianzige-horizontal-3-2')))
    .toEqual({ orientation: 'horizontal', dash: 'dashed' });
  expect(copybookGuideAppearance(shape('coco-tianzige-vertical-3-2')))
    .toEqual({ orientation: 'vertical', dash: 'dashed' });
  expect(copybookGuideAppearance(shape('coco-pinyin-guide-2-1')))
    .toEqual({ orientation: 'horizontal', dash: 'dashed' });
  expect(copybookGuideAppearance(shape('coco-pinyin-outline-2-top')))
    .toEqual({ orientation: 'horizontal', dash: 'solid' });
  expect(copybookGuideAppearance(shape('coco-pinyin-outline-2-vertical-4')))
    .toEqual({ orientation: 'vertical', dash: 'solid' });
  expect(copybookGuideAppearance(shape('unrelated-shape'))).toBeNull();
});

beforeEach(() => {
  vi.stubGlobal('confirm', vi.fn(() => { throw new Error('Native confirm must not be used by the editor'); }));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    return new Proxy({
      canvas: this,
      globalAlpha: 1,
      measureText: (text: string) => ({ width: text.length * 7, actualBoundingBoxAscent: 9, actualBoundingBoxDescent: 3 }),
      getLineDash: () => [],
    }, { get: (target, key) => Reflect.get(target, key) ?? (() => {}) }) as unknown as CanvasRenderingContext2D;
  });
});

afterEach(async () => {
  for (const editor of editors.splice(0)) editor.destroy();
  await new Promise((resolve) => setTimeout(resolve, 0));
  document.body.replaceChildren();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

async function setup(options: SpreadsheetEditorOptions = {}, prepare?: (workbook: ExcelWorkbook) => Promise<void>) {
  const workbook = await ExcelWorkbook.fromModel({
    sharedStrings: [], definedNames: [],
    styles: { numberFormats: [], fonts: [{}], fills: [{}], borders: [{}], cellFormats: [{}] },
    sheets: [{
      id: '1', name: 'Sheet1', sheetViews: [], columns: [], rows: [], mergedRanges: [],
      hyperlinks: [], tables: [], dataValidations: [], conditionalFormats: [], comments: [],
      threadedComments: [], drawings: [], charts: [], images: [],
      cells: [{ address: 'A1', row: 1, column: 1, kind: 'string', value: '原内容' }],
    }],
  }, { mode: 'inline' });
  await prepare?.(workbook);
  const bytes = await workbook.toZip();
  await workbook.dispose();
  const host = document.createElement('div');
  document.body.appendChild(host);
  vi.spyOn(host, 'getBoundingClientRect').mockReturnValue({
    x: 0, y: 0, left: 0, top: 0, right: 640, bottom: 480, width: 640, height: 480, toJSON: () => ({}),
  });
  host.setPointerCapture = vi.fn();
  host.releasePointerCapture = vi.fn();
  host.hasPointerCapture = vi.fn(() => false);
  const editor = await renderSpreadsheetEditor(host, bytes, options);
  editors.push(editor);
  return { editor, host, root: host.shadowRoot! };
}

async function edit(host: HTMLElement, root: ShadowRoot, value: string) {
  host.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true, composed: true }));
  const textarea = root.querySelector<HTMLTextAreaElement>('textarea')!;
  await vi.waitFor(() => expect(textarea.style.display).toBe('block'));
  textarea.value = value;
  textarea.dispatchEvent(new Event('input', { bubbles: true }));
  return textarea;
}

function button(root: ShadowRoot, label: string): HTMLButtonElement {
  const found = Array.from(root.querySelectorAll<HTMLButtonElement>('button')).find((item) =>
    item.getAttribute('aria-label') === label || item.textContent === label,
  );
  if (!found) throw new Error(`Missing button: ${label}`);
  return found;
}

async function borderedGrid(workbook: ExcelWorkbook, value: string | null = null) {
  const range = workbook.worksheets.getItem('Sheet1').getRange('A1:E5');
  await range.setValues(Array.from({ length: 5 }, () => Array(5).fill(value)));
  const side = { style: 'thin', color: { rgb: '#2C7A69' } };
  await range.format.border.setValue({ top: side, bottom: side, left: side, right: side });
  await range.format.font.setValue({ name: 'KaiTi', size: 26 });
}

function selectCell(host: HTMLElement, editor: SpreadsheetEditor, row: number, column: number) {
  const start = editor.selection().active;
  const keys = [
    ...Array(Math.abs(row - start.row)).fill(row < start.row ? 'ArrowUp' : 'ArrowDown'),
    ...Array(Math.abs(column - start.column)).fill(column < start.column ? 'ArrowLeft' : 'ArrowRight'),
  ];
  for (const key of keys) host.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, composed: true }));
  expect(editor.selection().active).toEqual({ row, column });
}

const gridEdges = [
  { side: 'left', address: 'A3', row: 3, column: 1 },
  { side: 'right', address: 'E3', row: 3, column: 5 },
  { side: 'top', address: 'C1', row: 1, column: 3 },
  { side: 'bottom', address: 'C5', row: 5, column: 3 },
];

function pressKey(host: HTMLElement, key: string, options: KeyboardEventInit = {}) {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, composed: true, cancelable: true, ...options });
  host.dispatchEvent(event);
  return event;
}

function cellPoint(editor: SpreadsheetEditor, row: number, column: number) {
  const layout = new SpreadsheetLayout(editor.runtime().activeSheet());
  const rect = layout.getCellLogicalRect(row, column);
  return { clientX: layout.rowHeaderWidth + rect.x + rect.width / 2, clientY: layout.columnHeaderHeight + rect.y + rect.height / 2 };
}

function copyCells(host: HTMLElement, kind: 'copy' | 'cut' = 'copy') {
  const data = new DataTransfer();
  host.dispatchEvent(new ClipboardEvent(kind, { clipboardData: data, cancelable: true }));
  return data;
}

describe('spreadsheet common shortcuts', () => {
  test('separate edit commits to the same cell remain separate undo actions', async () => {
    const { editor, host, root } = await setup();
    await edit(host, root, 'first'); await editor.flushEdits();
    await edit(host, root, 'second'); await editor.flushEdits();
    pressKey(host, 'z', { ctrlKey: true }); await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('first');
    pressKey(host, 'z', { ctrlKey: true }); await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('原内容');
  });

  test('typing replaces the selected cell while F2 edits its existing content', async () => {
    const { editor, host, root } = await setup();
    pressKey(host, '7');
    const textarea = root.querySelector<HTMLTextAreaElement>('textarea')!;
    await vi.waitFor(() => expect(textarea.style.display).toBe('block'));
    expect(textarea.value).toBe('7');
    expect(textarea.selectionStart).toBe(1);
    pressKey(textarea, 'Enter');
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe(7);
    selectCell(host, editor, 1, 1);
    pressKey(host, 'F2');
    await vi.waitFor(() => expect(textarea.style.display).toBe('block'));
    expect(textarea.value).toBe('7');
  });

  test('Ctrl+D and Ctrl+R fill from the leading row/column without generating a numeric sequence', async () => {
    const { editor, host } = await setup({}, async (workbook) => {
      await workbook.worksheets.getItem('Sheet1').getRange('A1:B1').setValues([[5, 10]]);
    });
    pressKey(host, 'ArrowDown', { shiftKey: true });
    pressKey(host, 'ArrowDown', { shiftKey: true });
    pressKey(host, 'd', { ctrlKey: true });
    await editor.flushEdits();
    expect(editor.cellState('A2').value).toBe(5);
    expect(editor.cellState('A3').value).toBe(5);
    pressKey(host, 'ArrowRight');
    pressKey(host, 'ArrowRight', { shiftKey: true });
    pressKey(host, 'r', { ctrlKey: true });
    await editor.flushEdits();
    expect(editor.cellState('C1').value).toBe(10);
    pressKey(host, 'z', { ctrlKey: true });
    await editor.flushEdits();
    expect(editor.cellState('C1').value).toBeNull();
  });

  test.each([
    { modifier: 'ctrlKey', redo: 'y', shift: false },
    { modifier: 'ctrlKey', redo: 'z', shift: true },
    { modifier: 'metaKey', redo: 'z', shift: true },
    { modifier: 'metaKey', redo: 'y', shift: false },
  ])('$modifier undo and $redo redo are queued in order', async ({ modifier, redo, shift }) => {
    const { editor, host } = await setup();
    await editor.setCellValue('A1', 'changed');
    await editor.setCellValue('B1', 42);
    expect(pressKey(host, 'z', { [modifier]: true }).defaultPrevented).toBe(true);
    pressKey(host, 'z', { [modifier]: true });
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('原内容');
    expect(editor.cellState('B1').value).toBeNull();
    pressKey(host, redo, { [modifier]: true, shiftKey: shift });
    pressKey(host, redo, { [modifier]: true, shiftKey: shift });
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('changed');
    expect(editor.cellState('B1').value).toBe(42);
  });

  test('Enter and Tab navigate, Shift reverses, and edit commits move once', async () => {
    const { editor, host, root } = await setup();
    pressKey(host, 'Enter');
    expect(editor.selection().active).toEqual({ row: 2, column: 1 });
    pressKey(host, 'Tab');
    expect(editor.selection().active).toEqual({ row: 2, column: 2 });
    pressKey(host, 'Tab', { shiftKey: true });
    pressKey(host, 'Enter', { shiftKey: true });
    expect(editor.selection().active).toEqual({ row: 1, column: 1 });
    const textarea = await edit(host, root, 'Tab 保存');
    pressKey(textarea, 'Tab');
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('Tab 保存');
    expect(editor.selection().active).toEqual({ row: 1, column: 2 });
    await edit(host, root, '换行保留');
    expect(pressKey(textarea, 'Enter', { altKey: true }).defaultPrevented).toBe(false);
    expect(textarea.style.display).toBe('block');
    pressKey(textarea, 'Enter');
    await editor.flushEdits();
    expect(editor.selection().active).toEqual({ row: 2, column: 2 });
  });

  test('Shift arrows extend and shrink the selection around a stable anchor', async () => {
    const { editor, host } = await setup();
    selectCell(host, editor, 3, 3);
    pressKey(host, 'ArrowRight', { shiftKey: true });
    pressKey(host, 'ArrowRight', { shiftKey: true });
    expect(editor.selection().range).toEqual({ startRow: 3, endRow: 3, startColumn: 3, endColumn: 5 });
    for (let i = 0; i < 3; i++) pressKey(host, 'ArrowLeft', { shiftKey: true });
    expect(editor.selection().active).toEqual({ row: 3, column: 3 });
    expect(editor.selection().range).toEqual({ startRow: 3, endRow: 3, startColumn: 2, endColumn: 3 });
    pressKey(host, 'ArrowDown', { shiftKey: true });
    expect(editor.selection().range).toEqual({ startRow: 3, endRow: 4, startColumn: 2, endColumn: 3 });
  });

  test('selects the current region, then the sheet, and supports row/column selection', async () => {
    const { editor, host } = await setup({}, async (workbook) => {
      await workbook.worksheets.getItem('Sheet1').getRange('A1:B2').setValues([[1, 2], [3, 4]]);
    });
    pressKey(host, 'a', { ctrlKey: true });
    expect(editor.selection().range).toEqual({ startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
    pressKey(host, 'a', { ctrlKey: true });
    expect(editor.selection().range.endRow).toBe(100);
    expect(editor.selection().range.endColumn).toBe(26);
    selectCell(host, editor, 3, 3);
    pressKey(host, ' ', { shiftKey: true });
    expect(editor.selection().kind).toBe('row');
    expect(editor.selection().range.startRow).toBe(3);
    pressKey(host, ' ', { ctrlKey: true });
    expect(editor.selection().kind).toBe('column');
  });

  test('Ctrl arrows jump data boundaries; Home/End and sheet switching use the workbook', async () => {
    const { editor, host } = await setup({}, async (workbook) => {
      await workbook.worksheets.getItem('Sheet1').getRange('A1:A4').setValues([[1], [2], [null], [4]]);
      await workbook.worksheets.getItem('Sheet1').getRange('D10').setValues([['last']]);
    });
    pressKey(host, 'ArrowDown', { ctrlKey: true });
    expect(editor.selection().active.row).toBe(2);
    pressKey(host, 'ArrowDown', { ctrlKey: true });
    expect(editor.selection().active.row).toBe(4);
    pressKey(host, 'End', { ctrlKey: true });
    expect(editor.selection().active).toEqual({ row: 10, column: 4 });
    pressKey(host, 'Home', { ctrlKey: true });
    expect(editor.selection().active).toEqual({ row: 1, column: 1 });
    await editor.addSheet('Other');
    pressKey(host, 'PageUp', { ctrlKey: true });
    await editor.flushEdits();
    expect(editor.activeSheetName()).toBe('Sheet1');
    pressKey(host, 'PageDown', { ctrlKey: true });
    await editor.flushEdits();
    expect(editor.activeSheetName()).toBe('Other');
  });

  test('keyboard navigation scrolls the active cell into view', async () => {
    const { editor, host, root } = await setup();
    for (let i = 0; i < 40; i++) pressKey(host, 'ArrowDown');
    expect(editor.selection().active.row).toBe(41);
    const textarea = await edit(host, root, 'visible');
    expect(Number.parseFloat(textarea.style.top)).toBeGreaterThanOrEqual(24);
    expect(Number.parseFloat(textarea.style.top)).toBeLessThan(420);
  });

  test.each([['b', 'bold'], ['i', 'italic'], ['u', 'underline']] as const)('Ctrl/Command+%s toggles %s instead of always enabling it', async (key, property) => {
    const { editor, host } = await setup();
    pressKey(host, key, { ctrlKey: true });
    await editor.flushEdits();
    expect(editor.cellState('A1').format?.font).toMatchObject({ [property]: true });
    pressKey(host, key, { metaKey: true });
    await editor.flushEdits();
    expect(editor.cellState('A1').format?.font).toMatchObject({ [property]: false });
    pressKey(host, 'z', { ctrlKey: true });
    await editor.flushEdits();
    expect(editor.cellState('A1').format?.font).toMatchObject({ [property]: true });
  });

  test('copy/paste preserves formulas; values-only paste preserves the destination format', async () => {
    const { editor, host } = await setup({}, async (workbook) => {
      const sheet = workbook.worksheets.getItem('Sheet1');
      await sheet.getRange('A1').setFormulas([['1+1']]);
      await sheet.getRange('A1').format.font.setBold(true);
      await sheet.getRange('B1').format.font.setItalic(true);
    });
    expect(pressKey(host, 'c', { ctrlKey: true }).defaultPrevented).toBe(false);
    const data = copyCells(host);
    selectCell(host, editor, 1, 2);
    pressKey(host, 'v', { ctrlKey: true, shiftKey: true });
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('B1').value).toBe(2);
    expect(editor.cellState('B1').formula).toBeUndefined();
    expect(editor.cellState('B1').format?.font).toMatchObject({ italic: true });
    selectCell(host, editor, 1, 3);
    pressKey(host, 'v', { metaKey: true });
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('C1').formula).toBe('1+1');
    expect(editor.cellState('C1').format?.font).toMatchObject({ bold: true });
  });

  test('external text identical to a previous copy does not reuse its styles', async () => {
    const { editor, host } = await setup({}, async (workbook) => {
      await workbook.worksheets.getItem('Sheet1').getRange('A1').format.font.setBold(true);
      await workbook.worksheets.getItem('Sheet1').getRange('B1').format.font.setItalic(true);
    });
    copyCells(host);
    selectCell(host, editor, 1, 2);
    const data = new DataTransfer(); data.setData('text/plain', '原内容');
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('B1').format?.font).toMatchObject({ italic: true });
    expect(editor.cellState('B1').value).toBe('原内容');
  });

  test.each(['custom', 'html'])('copies formulas and actual formats between independent workbooks via %s, even after the source closes', async (kind) => {
    const prepare = async (workbook: ExcelWorkbook, source: boolean) => {
      const sheet = workbook.worksheets.getItem('Sheet1');
      await sheet.getRange('A1:B2').setValues([[12.5, null], [0, false]]);
      await sheet.getRange('B1').setFormulas([['A1+$A$1+A$1+$A1']]);
      const range = sheet.getRange(source ? 'A1:B2' : 'C3:D4');
      await range.format.font.setValue({ name: source ? 'KaiTi' : 'Arial', size: source ? 26 : 9, color: { rgb: source ? '#AA3311' : '#1133AA' } });
      await range.format.fill.setValue({ patternType: 'solid', foregroundColor: { rgb: source ? '#FFFFAA' : '#AAFFFF' } });
      const side = { style: source ? 'thin' : 'dashed', color: { rgb: '#228866' } };
      await range.format.border.setValue({ top: side, bottom: side, left: side, right: side });
      await range.format.alignment.setValue({ horizontal: source ? 'center' : 'right', vertical: 'center', wrapText: true });
      await range.format.setNumberFormatCode(source ? '0.00" 元"' : 'yyyy-mm-dd');
      if (!source) {
        await sheet.getRange('A1').setValues([[2]]);
        await sheet.getRange('C1').setValues([[3]]);
        await sheet.getRange('A3').setValues([[4]]);
      }
    };
    const source = await setup({ workbookId: 'source' }, (workbook) => prepare(workbook, true));
    const target = await setup({ workbookId: 'target' }, (workbook) => prepare(workbook, false));
    const format = source.editor.cellState('A1').format;
    const snapshot = captureExcelCellStyle(source.editor.runtime().model().styles, source.editor.cellState('A1').styleId);
    const previous = target.editor.cellState('C3');
    expect(source.editor.cellState('A1').styleId).toBe(previous.styleId);
    pressKey(source.host, 'ArrowRight', { shiftKey: true });
    pressKey(source.host, 'ArrowDown', { shiftKey: true });
    const copied = copyCells(source.host);
    const clipboard = new DataTransfer();
    clipboard.setData('text/plain', copied.getData('text/plain'));
    const type = kind === 'custom' ? SPREADSHEET_CLIPBOARD_MIME : 'text/html';
    clipboard.setData(type, copied.getData(type));
    source.editor.destroy(); editors.splice(editors.indexOf(source.editor), 1);
    selectCell(target.host, target.editor, 3, 3);
    target.host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: clipboard, cancelable: true }));
    await target.editor.flushEdits();
    expect(target.editor.cellState('C3')).toMatchObject({ value: 12.5, format });
    expect(target.editor.cellState('C3').styleId).not.toBe(previous.styleId);
    expect(target.editor.cellState('D3')).toMatchObject({ formula: 'C3+$A$1+C$1+$A3', value: 21.5 });
    expect(target.editor.cellState('C4').value).toBe(0);
    expect(target.editor.cellState('D4').value).toBe(false);
    pressKey(target.host, 'z', { metaKey: true }); await target.editor.flushEdits();
    expect(target.editor.cellState('C3')).toEqual(previous);
    pressKey(target.host, 'z', { metaKey: true, shiftKey: true }); await target.editor.flushEdits();
    expect(target.editor.cellState('C3').format).toEqual(format);
    const reopened = await ExcelWorkbook.fromZip(await target.editor.toZip(), { mode: 'inline' });
    try {
      const model = await reopened.toModel();
      const cell = model.sheets[0]!.cells.find((cell) => cell.address === 'C3')!;
      expect(captureExcelCellStyle(model.styles, cell.styleId)).toEqual(snapshot);
      expect(await reopened.worksheets.getItem('Sheet1').getRange('D3').formulas()).toEqual([['C3+$A$1+C$1+$A3']]);
    } finally { await reopened.dispose(); }
    const styles = JSON.stringify(target.editor.runtime().model().styles);
    selectCell(target.host, target.editor, 6, 3);
    target.host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: clipboard, cancelable: true }));
    await target.editor.flushEdits();
    expect(JSON.stringify(target.editor.runtime().model().styles)).toBe(styles);
  });

  test('cross-document paste preserves merged geometry and styled empty follower cells', async () => {
    const source = await setup({}, async (workbook) => {
      await borderedGrid(workbook, '田');
      await workbook.worksheets.getItem('Sheet1').getRange('A1:B2').merge();
    });
    const target = await setup();
    const copied = copyCells(source.host);
    const clipboard = new DataTransfer(); clipboard.setData('text/html', copied.getData('text/html'));
    selectCell(target.host, target.editor, 3, 3);
    target.host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: clipboard, cancelable: true }));
    await target.editor.flushEdits();
    expect((await target.editor.runtime().workbook().worksheets.getItem('Sheet1').mergedRanges()).map((merge) => merge.address)).toEqual(['C3:D4']);
    expect(target.editor.cellState('C3').value).toBe('田');
    expect(target.editor.cellState('D4').format?.border).toMatchObject({ right: { style: 'thin' }, bottom: { style: 'thin' } });
  });

  test('cross-document values-only paste leaves target formats and style pools untouched', async () => {
    const source = await setup({}, async (workbook) => {
      const range = workbook.worksheets.getItem('Sheet1').getRange('A1');
      await range.setFormulas([['1+1']]); await range.format.font.setValue({ name: 'KaiTi', size: 26 });
    });
    const target = await setup({}, async (workbook) => {
      await workbook.worksheets.getItem('Sheet1').getRange('A1').format.font.setItalic(true);
    });
    const styles = structuredClone(target.editor.runtime().model().styles);
    const format = target.editor.cellState('A1').format;
    pressKey(target.host, 'v', { ctrlKey: true, shiftKey: true });
    target.host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: copyCells(source.host), cancelable: true }));
    await target.editor.flushEdits();
    expect(target.editor.cellState('A1')).toMatchObject({ value: 2, format });
    expect(target.editor.cellState('A1').formula).toBeUndefined();
    expect(target.editor.runtime().model().styles).toEqual(styles);
  });

  test('damaged cross-document metadata reports an error without falling back to text or changing the document', async () => {
    const source = await setup();
    const onError = vi.fn(); const target = await setup({ onError });
    const clipboard = copyCells(source.host);
    const payload = JSON.parse(clipboard.getData(SPREADSHEET_CLIPBOARD_MIME)); payload.styles = {};
    clipboard.setData(SPREADSHEET_CLIPBOARD_MIME, JSON.stringify(payload));
    const before = structuredClone(target.editor.runtime().model());
    target.host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: clipboard, cancelable: true }));
    await vi.waitFor(() => expect(onError).toHaveBeenCalledOnce());
    await target.editor.flushEdits();
    expect(target.editor.runtime().model()).toEqual(before);
    expect(target.editor.runtime().canUndo()).toBe(false);
  });

  test('cut is deferred until paste, can be cancelled, and moves in one undo group', async () => {
    const { editor, host } = await setup();
    let data = copyCells(host, 'cut');
    expect(editor.cellState('A1').value).toBe('原内容');
    pressKey(host, 'Escape');
    selectCell(host, editor, 1, 2);
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('原内容');
    data = copyCells(host, 'cut');
    selectCell(host, editor, 1, 3);
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('B1').value).toBeNull();
    expect(editor.cellState('C1').value).toBe('原内容');
    pressKey(host, 'z', { ctrlKey: true });
    await editor.flushEdits();
    expect(editor.cellState('B1').value).toBe('原内容');
    expect(editor.cellState('C1').value).toBeNull();
  });

  test('Enter pastes an active copy and exits copy mode, including after changing sheets', async () => {
    const { editor, host } = await setup();
    copyCells(host);
    await editor.addSheet('Other');
    pressKey(host, 'Enter');
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('原内容');
    pressKey(host, 'Enter');
    expect(editor.selection().active.row).toBe(2);
  });

  test('paste captures its destination and is saved before immediate export', async () => {
    const { editor, host } = await setup();
    const data = copyCells(host);
    selectCell(host, editor, 1, 2);
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, cancelable: true }));
    pressKey(host, 'ArrowRight');
    const workbook = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try { expect(await workbook.worksheets.getItem('Sheet1').getRange('A1:C1').values()).toEqual([['原内容', '原内容', null]]); }
    finally { await workbook.dispose(); }
    expect(editor.selection().active).toEqual({ row: 1, column: 3 });
  });

  test('external blank and ragged TSV pastes preserve formatting and all columns', async () => {
    const { editor, host } = await setup({}, (workbook) => borderedGrid(workbook, '田'));
    const before = editor.cellState('A1').format;
    const blank = new DataTransfer(); blank.setData('text/plain', '');
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: blank, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBeNull();
    expect(editor.cellState('A1').format).toEqual(before);
    const ragged = new DataTransfer(); ragged.setData('text/plain', '1\n2\t3');
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: ragged, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe(1);
    expect(editor.cellState('B1').value).toBeNull();
    expect(editor.cellState('B2').value).toBe(3);
  });

  test('pastes into a merged cell without unmerging or filling hidden follower cells', async () => {
    const { editor, host } = await setup({}, async (workbook) => {
      await borderedGrid(workbook, '田');
      await workbook.worksheets.getItem('Sheet1').getRange('A1:B2').merge();
    });
    const data = new DataTransfer(); data.setData('text/plain', '=1+2');
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('A1')).toMatchObject({ formula: '1+2', value: 3 });
    expect((await editor.runtime().workbook().worksheets.getItem('Sheet1').mergedRanges()).map((merge) => merge.address)).toEqual(['A1:B2']);
    await editor.runtime().workbook().worksheets.getItem('Sheet1').getRange('A1:B2').unmerge();
    expect(await editor.runtime().workbook().worksheets.getItem('Sheet1').getRange('A1:B2').values()).toEqual([[null, null], [null, null]]);
    expect(await editor.runtime().workbook().worksheets.getItem('Sheet1').getRange('A1').formulas()).toEqual([['1+2']]);
  });

  test('input controls and read-only mode retain their own editing and clipboard behavior', async () => {
    const onChange = vi.fn();
    const { editor, host, root } = await setup({ onChange });
    const input = document.createElement('input'); root.appendChild(input);
    for (const key of ['a', 'z', 'y', 'b', 'c', 'v']) expect(pressKey(input, key, { ctrlKey: true }).defaultPrevented).toBe(false);
    for (const type of ['copy', 'cut', 'paste']) {
      const event = new ClipboardEvent(type, { clipboardData: new DataTransfer(), bubbles: true, composed: true, cancelable: true });
      input.dispatchEvent(event); expect(event.defaultPrevented).toBe(false);
    }
    const data = copyCells(host);
    editor.setReadOnly(true);
    for (const key of ['z', 'y', 'b']) pressKey(host, key, { ctrlKey: true });
    host.dispatchEvent(new ClipboardEvent('cut', { clipboardData: data, cancelable: true }));
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('原内容');
    expect(onChange).not.toHaveBeenCalled();
    expect(editor.canUndo()).toBe(false);
  });
});

describe('spreadsheet keyboard clearing', () => {
  test.each(['Delete', 'Backspace'])('%s clears only the selected cell and keeps formatting, undo and writeback', async (key) => {
    const onChange = vi.fn();
    const { editor, host } = await setup({ onChange }, async (workbook) => {
      await borderedGrid(workbook, '田');
      await workbook.worksheets.getItem('Sheet1').getRange('A1').setFormulas([['1+1']]);
    });
    const before = editor.cellState('A1');
    expect(pressKey(host, key).defaultPrevented).toBe(true);
    await editor.flushEdits();
    expect(editor.cellState('A1')).toMatchObject({ value: null, format: before.format, styleId: before.styleId });
    expect(editor.cellState('A1').formula).toBeUndefined();
    expect(editor.cellState('B1').value).toBe('田');
    expect(onChange).toHaveBeenCalledOnce();
    await editor.undo();
    expect(editor.cellState('A1').formula).toBe('1+1');
    expect(editor.cellState('A1').value).toBe(2);
    await editor.redo();
    const workbook = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try {
      const cell = workbook.worksheets.getItem('Sheet1').getRange('A1');
      expect(await cell.values()).toEqual([[null]]);
      expect(await cell.formulas()).toEqual([[null]]);
      expect(await cell.styleIds()).toEqual([[before.styleId]]);
    } finally { await workbook.dispose(); }
  });

  test('clears a dragged range as one undo action even if selection moves immediately afterwards', async () => {
    const { editor, host } = await setup({}, (workbook) => borderedGrid(workbook, '田'));
    const externalInput = document.createElement('input');
    document.body.appendChild(externalInput);
    externalInput.focus();
    for (const [type, row, column] of [['pointerdown', 1, 1], ['pointermove', 2, 2], ['pointerup', 2, 2]] as const) {
      host.dispatchEvent(new PointerEvent(type, { ...cellPoint(editor, row, column), pointerId: 1, button: 0, bubbles: true, composed: true }));
    }
    expect(document.activeElement).toBe(host);
    expect(editor.selection().range).toEqual({ startRow: 1, startColumn: 1, endRow: 2, endColumn: 2 });
    pressKey(host, 'Delete');
    for (let i = 0; i < 3; i++) pressKey(host, 'ArrowRight');
    await editor.flushEdits();
    for (const address of ['A1', 'A2', 'B1', 'B2']) expect(editor.cellState(address).value).toBeNull();
    expect(editor.cellState('D1').value).toBe('田');
    expect(editor.runtime().workbook().history().undoGroups()).toHaveLength(1);
    await editor.undo();
    for (const address of ['A1', 'A2', 'B1', 'B2']) expect(editor.cellState(address).value).toBe('田');
  });

  test('export waits for a keyboard clear and its persistence callback', async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const onChange = vi.fn(() => gate);
    const { editor, host } = await setup({ onChange });
    pressKey(host, 'Delete');
    const exporting = editor.toZip();
    let finished = false;
    void exporting.then(() => { finished = true; });
    try {
      await vi.waitFor(() => expect(onChange).toHaveBeenCalledOnce());
      expect(finished).toBe(false);
    } finally { release(); }
    const workbook = await ExcelWorkbook.fromZip(await exporting, { mode: 'inline' });
    try { expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([[null]]); }
    finally { await workbook.dispose(); }
  });

  test('finishes keyboard clearing in the original sheet before switching sheets', async () => {
    const { editor, host } = await setup();
    await editor.addSheet('Other');
    await editor.setCellValue('A1', '另一张表');
    await editor.setActiveSheet('Sheet1');
    pressKey(host, 'Delete');
    await editor.setActiveSheet('Other');
    expect(editor.cellState('A1').value).toBe('另一张表');
    await editor.setActiveSheet('Sheet1');
    expect(editor.cellState('A1').value).toBeNull();
  });

  test('persists keyboard clearing if the host immediately unmounts the editor', async () => {
    const onChange = vi.fn();
    const { editor, host } = await setup({ onChange });
    pressKey(host, 'Delete');
    editor.destroy();
    await vi.waitFor(() => expect(onChange).toHaveBeenCalledOnce());
    const workbook = await ExcelWorkbook.fromZip(onChange.mock.calls[0]![0].bytes, { mode: 'inline' });
    try { expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([[null]]); }
    finally { await workbook.dispose(); }
  });

  test.each(['Delete', 'Backspace'])('leaves native text editing and read-only cells alone for %s', async (key) => {
    const onChange = vi.fn();
    const { editor, host, root } = await setup({ onChange });
    const textarea = await edit(host, root, '正在输入');
    expect(pressKey(textarea, key).defaultPrevented).toBe(false);
    expect(textarea.style.display).toBe('block');
    expect(editor.cellState('A1').value).toBe('原内容');
    pressKey(textarea, 'Escape');
    editor.setReadOnly(true);
    pressKey(host, key);
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('原内容');
    expect(onChange).not.toHaveBeenCalled();
    expect(editor.canUndo()).toBe(false);
  });

  test('does not clear cells for modified keys, IME, or a focused input in the editor', async () => {
    const onChange = vi.fn();
    const { editor, host, root } = await setup({ onChange });
    for (const options of [{ ctrlKey: true }, { metaKey: true }, { altKey: true }, { shiftKey: true }, { isComposing: true }]) {
      expect(pressKey(host, 'Delete', options).defaultPrevented).toBe(false);
    }
    const input = document.createElement('input');
    root.appendChild(input);
    expect(pressKey(input, 'Delete').defaultPrevented).toBe(false);
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('原内容');
    expect(onChange).not.toHaveBeenCalled();
  });
});

describe('spreadsheet copy outline lifecycle', () => {
  test.each(['double-click', 'F2', 'switch-sheet', 'new-sheet', 'Escape'])('dismisses the outline on %s but keeps copied data available', async (action) => {
    const onChange = vi.fn();
    const { editor, host, root } = await setup({ onChange });
    await editor.addSheet('Other');
    await editor.setActiveSheet('Sheet1');
    const outline = vi.spyOn(SpreadsheetCanvasRenderer.prototype, 'renderCopiedRange');
    const clipboard = new DataTransfer();
    host.dispatchEvent(new ClipboardEvent('copy', { clipboardData: clipboard, cancelable: true }));
    pressKey(host, 'ArrowRight');
    expect(outline).toHaveBeenCalled();
    outline.mockClear();
    if (action === 'switch-sheet') await editor.setActiveSheet('Other');
    else if (action === 'new-sheet') await editor.addSheet('Third');
    else if (action === 'Escape') pressKey(host, 'Escape');
    else {
      if (action === 'double-click') host.dispatchEvent(new MouseEvent('dblclick', { ...cellPoint(editor, 2, 2), bubbles: true, composed: true, cancelable: true }));
      else pressKey(host, action);
      const textarea = root.querySelector<HTMLTextAreaElement>('textarea')!;
      await vi.waitFor(() => expect(textarea.style.display).toBe('block'));
      pressKey(textarea, 'Escape');
    }
    expect(outline).not.toHaveBeenCalled();
    selectCell(host, editor, 5, 2);
    const changes = onChange.mock.calls.length;
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: clipboard, cancelable: true }));
    await vi.waitFor(() => expect(onChange).toHaveBeenCalledTimes(changes + 1));
    expect(editor.cellState('B5').value).toBe('原内容');
    expect(outline).not.toHaveBeenCalled();
  });
});

describe('spreadsheet browser editing', () => {
  test('pastes a copied blank cell over content and preserves its border through undo, redo, and export', async () => {
    const onChange = vi.fn();
    const { editor, host } = await setup({ onChange }, async (workbook) => {
      await borderedGrid(workbook);
      await workbook.worksheets.getItem('Sheet1').getRange('A3').setValues([['田']]);
    });
    const clipboard = new DataTransfer();
    host.dispatchEvent(new ClipboardEvent('copy', { clipboardData: clipboard, cancelable: true }));
    expect(clipboard.getData('text/plain')).toBe('');
    selectCell(host, editor, 3, 1);
    const before = editor.cellState('A3').format;
    host.dispatchEvent(new ClipboardEvent('paste', { clipboardData: clipboard, cancelable: true }));
    await vi.waitFor(() => expect(onChange).toHaveBeenCalledOnce());
    expect(editor.cellState('A3').value).toBeNull();
    expect(editor.cellState('A3').format).toEqual(before);
    await editor.undo();
    expect(editor.cellState('A3').value).toBe('田');
    await editor.redo();
    const reopened = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try {
      const target = reopened.worksheets.getItem('Sheet1').getRange('A3');
      expect(await target.values()).toEqual([[null]]);
      expect((await target.styleIds())[0]![0]).toBeDefined();
    } finally { await reopened.dispose(); }
  });

  test.each(['0', 'FALSE', ' '] as const)('commits %j without treating it as blank', async (text) => {
    const { editor, host, root } = await setup();
    await edit(host, root, text);
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe(text === '0' ? 0 : text === 'FALSE' ? false : text);
  });

  test.each(gridEdges)('keeps the $side grid border when leaving an unchanged blank editor', async ({ address, row, column }) => {
    const onChange = vi.fn();
    const { editor, host, root } = await setup({ onChange }, borderedGrid);
    selectCell(host, editor, row, column);
    const before = editor.cellState(address);
    for (const exit of ['blur', 'Enter', 'Escape']) {
      const textarea = await edit(host, root, '');
      if (exit === 'blur') textarea.dispatchEvent(new FocusEvent('blur'));
      else textarea.dispatchEvent(new KeyboardEvent('keydown', { key: exit, bubbles: true, composed: true }));
      await editor.flushEdits();
      expect(editor.cellState(address).format).toEqual(before.format);
      expect(editor.canUndo()).toBe(false);
      expect(onChange).not.toHaveBeenCalled();
    }
    const reopened = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try {
      expect(await reopened.worksheets.getItem('Sheet1').getRange(address).styleIds()).toEqual([[before.styleId]]);
    } finally { await reopened.dispose(); }
  });

  test('preserves all four grid edges when deleting text, undoing, redoing and exporting', async () => {
    const { editor, host, root } = await setup({}, (workbook) => borderedGrid(workbook, '田'));
    const styles = new Map<string, number | undefined>();
    for (const { address, row, column } of gridEdges) {
      selectCell(host, editor, row, column);
      const before = editor.cellState(address);
      styles.set(address, before.styleId);
      const textarea = await edit(host, root, '');
      textarea.dispatchEvent(new FocusEvent('blur'));
      await editor.flushEdits();
      expect(editor.cellState(address).value).toBeNull();
      expect(editor.cellState(address).format).toEqual(before.format);
      await editor.undo();
      expect(editor.cellState(address).value).toBe('田');
      expect(editor.cellState(address).format).toEqual(before.format);
      await editor.redo();
      expect(editor.cellState(address).value).toBeNull();
      expect(editor.cellState(address).format).toEqual(before.format);
    }
    const reopened = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try {
      for (const { address } of gridEdges) {
        expect(await reopened.worksheets.getItem('Sheet1').getRange(address).styleIds()).toEqual([[styles.get(address)]]);
      }
    } finally { await reopened.dispose(); }
  });

  test('commits on blur once, persists through export, and supports undo/redo', async () => {
    const onChange = vi.fn();
    const { editor, host, root } = await setup({ onChange });
    const textarea = await edit(host, root, '失焦后保留');
    textarea.dispatchEvent(new FocusEvent('blur'));
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('失焦后保留');
    expect(onChange).toHaveBeenCalledTimes(1);
    const exported = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try {
      expect(await exported.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([['失焦后保留']]);
    } finally { await exported.dispose(); }
    await editor.undo();
    expect(editor.cellState('A1').value).toBe('原内容');
    await editor.redo();
    expect(editor.cellState('A1').value).toBe('失焦后保留');
  });

  test('commits when selecting another cell instead of cancelling the input', async () => {
    const { editor, host, root } = await setup();
    await edit(host, root, '点击其他格仍保留');
    host.dispatchEvent(new PointerEvent('pointerdown', { clientX: 200, clientY: 80, pointerId: 1, button: 0, bubbles: true, composed: true }));
    await editor.flushEdits();
    expect(editor.runtime().activeSheet().cellsByAddress.get('A1')?.value).toBe('点击其他格仍保留');
  });

  test('Enter commits without reopening the editor; Escape cancels without a later blur commit', async () => {
    const { editor, host, root } = await setup();
    const textarea = await edit(host, root, 'Enter 保存');
    textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, composed: true, cancelable: true }));
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('Enter 保存');
    expect(textarea.style.display).toBe('none');
    await edit(host, root, '不应保存');
    textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }));
    textarea.dispatchEvent(new FocusEvent('blur'));
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('Enter 保存');
  });

  test('does not commit the Enter key used to finish IME composition', async () => {
    const { editor, host, root } = await setup();
    const textarea = await edit(host, root, '拼音输入');
    textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', isComposing: true, bubbles: true, composed: true }));
    expect(textarea.style.display).toBe('block');
    expect(editor.cellState('A1').value).toBe('原内容');
    await editor.flushEdits();
    expect(editor.cellState('A1').value).toBe('拼音输入');
  });

  test('export waits for an in-flight blur commit and its persistence callback', async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const onChange = vi.fn(async () => gate);
    const { editor, host, root } = await setup({ onChange });
    const textarea = await edit(host, root, '立即导出');
    textarea.dispatchEvent(new FocusEvent('blur'));
    const exporting = editor.toZip();
    let finished = false;
    void exporting.then(() => { finished = true; });
    await vi.waitFor(() => expect(onChange).toHaveBeenCalledOnce());
    expect(finished).toBe(false);
    release();
    const workbook = await ExcelWorkbook.fromZip(await exporting, { mode: 'inline' });
    try {
      expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([['立即导出']]);
    } finally { await workbook.dispose(); }
  });

  test('flushes before switching sheets so an edit cannot land on the next sheet', async () => {
    const { editor, host, root } = await setup();
    await editor.addSheet('Other');
    await editor.setActiveSheet('Sheet1');
    await edit(host, root, '属于原表');
    await editor.setActiveSheet('Other');
    expect(editor.cellState('A1').value).toBeNull();
    await editor.setActiveSheet('Sheet1');
    expect(editor.cellState('A1').value).toBe('属于原表');
  });

  test('persists an active edit when the host unmounts the editor', async () => {
    const onChange = vi.fn();
    const { editor, host, root } = await setup({ onChange });
    await edit(host, root, '切换文档也保留');
    editor.destroy();
    await vi.waitFor(() => expect(onChange).toHaveBeenCalledOnce());
    const workbook = await ExcelWorkbook.fromZip(onChange.mock.calls[0]![0].bytes, { mode: 'inline' });
    try {
      expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([['切换文档也保留']]);
    } finally { await workbook.dispose(); }
  });
});

describe('spreadsheet worksheet controls', () => {
  test('creates a uniquely named sheet and exports it together with the committed original edit', async () => {
    const { editor, host, root } = await setup();
    await edit(host, root, '创建前保存');
    button(root, '新建工作表').click();
    await vi.waitFor(() => expect(editor.activeSheetName()).toBe('Sheet2'));
    await vi.waitFor(() => expect(button(root, '新建工作表').disabled).toBe(false));
    expect(button(root, 'Sheet2').getAttribute('aria-pressed')).toBe('true');
    expect(button(root, '删除工作表“Sheet2”').disabled).toBe(false);
    expect(root.querySelector('[aria-label="删除工作表“Sheet1”"]')).toBeNull();
    expect(root.querySelector('[aria-haspopup="menu"]')).toBeNull();
    const workbook = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try {
      expect((await workbook.toModel()).sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Sheet2']);
      expect(await workbook.worksheets.getItem('Sheet1').getRange('A1').values()).toEqual([['创建前保存']]);
    } finally { await workbook.dispose(); }
    button(root, 'Sheet1').click();
    await vi.waitFor(() => expect(editor.activeSheetName()).toBe('Sheet1'));
    await vi.waitFor(() => expect(button(root, '删除工作表“Sheet1”').disabled).toBe(false));
    expect(root.querySelector('[aria-label="删除工作表“Sheet2”"]')).toBeNull();
  });

  test('confirms deletion, switches to a neighbour, and prevents deleting the last sheet', async () => {
    const { editor, root } = await setup();
    await editor.addSheet('Other');
    await editor.setCellValue('A1', 'edited before deletion');
    expect(editor.canUndo()).toBe(true);
    button(root, '删除工作表“Other”').click();
    const dialog = root.querySelector<HTMLElement>('[role="dialog"]')!;
    expect(dialog.style.display).toBe('block');
    expect(dialog.getAttribute('aria-label')).toBe('删除工作表？');
    expect(dialog.textContent).toContain('Other');
    expect(dialog.textContent).toContain('此操作无法撤销。');
    expect(root.activeElement).toBe(button(root, '取消'));
    expect(window.confirm).not.toHaveBeenCalled();
    button(root, '取消').click();
    expect(dialog.style.display).toBe('none');
    expect(root.activeElement).toBe(button(root, '删除工作表“Other”'));
    expect(editor.activeSheetName()).toBe('Other');
    button(root, '删除工作表“Other”').click();
    const confirmDelete = button(root, '删除');
    confirmDelete.click();
    confirmDelete.click();
    await vi.waitFor(() => expect(editor.activeSheetName()).toBe('Sheet1'));
    expect(editor.canUndo()).toBe(false);
    await vi.waitFor(() => expect(button(root, '新建工作表').disabled).toBe(false));
    expect(Array.from(root.querySelectorAll('button')).some((item) => item.textContent === 'Other')).toBe(false);
    expect(button(root, '删除工作表“Sheet1”').disabled).toBe(true);
    const workbook = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try { expect((await workbook.toModel()).sheets.map((sheet) => sheet.name)).toEqual(['Sheet1']); }
    finally { await workbook.dispose(); }
  });

  test('dismisses deletion without changing the workbook on Escape, outside click, sheet switch or read-only mode', async () => {
    const { editor, root } = await setup();
    await editor.addSheet('Other');
    const dialog = root.querySelector<HTMLElement>('[role="dialog"]')!;
    button(root, '删除工作表“Other”').click();
    button(root, '取消').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }));
    expect(dialog.style.display).toBe('none');
    expect(root.activeElement).toBe(button(root, '删除工作表“Other”'));
    button(root, '删除工作表“Other”').click();
    document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, composed: true }));
    expect(dialog.style.display).toBe('none');
    button(root, '删除工作表“Other”').click();
    button(root, '新建工作表').focus();
    expect(dialog.style.display).toBe('none');
    button(root, '删除工作表“Other”').click();
    button(root, 'Sheet1').click();
    await vi.waitFor(() => expect(editor.activeSheetName()).toBe('Sheet1'));
    await vi.waitFor(() => expect(button(root, '新建工作表').disabled).toBe(false));
    expect(dialog.style.display).toBe('none');
    button(root, '删除工作表“Sheet1”').click();
    expect(dialog.style.display).toBe('block');
    editor.setReadOnly(true);
    expect(dialog.style.display).toBe('none');
    expect(root.querySelector('button[aria-label^="删除工作表"]')).toBeNull();
    const workbook = await ExcelWorkbook.fromZip(await editor.toZip(), { mode: 'inline' });
    try { expect((await workbook.toModel()).sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Other']); }
    finally { await workbook.dispose(); }
    expect(window.confirm).not.toHaveBeenCalled();
  });

  test('hides mutation controls in read-only mode and restores them when editing is enabled', async () => {
    const { editor, root } = await setup({ readOnly: true });
    expect(root.querySelector('[aria-label="新建工作表"]')).toBeNull();
    expect(root.querySelector('button[aria-label^="删除工作表"]')).toBeNull();
    editor.setReadOnly(false);
    expect(button(root, '新建工作表')).toBeDefined();
    expect(button(root, '删除工作表“Sheet1”')).toBeDefined();
    editor.setReadOnly(true);
    expect(root.querySelector('[aria-label="新建工作表"]')).toBeNull();
    expect(root.querySelector('button[aria-label^="删除工作表"]')).toBeNull();
  });
});
