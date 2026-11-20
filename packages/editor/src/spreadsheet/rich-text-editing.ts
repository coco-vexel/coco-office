import type { ExcelRichTextRun } from '@coco-vexel/coco-office/excel';
import type { SpreadsheetTextStyle } from './style-resolver.js';

export interface SpreadsheetEditableRun extends ExcelRichTextRun {
  style: SpreadsheetTextStyle;
}

export function writeEditableRuns(input: HTMLElement, runs: readonly SpreadsheetEditableRun[]): void {
  input.replaceChildren(...runs.map(({ text, style }, index) => {
    const span = input.ownerDocument.createElement('span');
    span.dataset.runIndex = String(index);
    span.textContent = text;
    Object.assign(span.style, {
      font: style.font, color: style.fontColor,
      textDecoration: style.underline ? 'underline' : 'none',
    });
    return span;
  }));
}

export function readEditableRuns(input: HTMLElement, source: readonly ExcelRichTextRun[]): ExcelRichTextRun[] {
  const result: ExcelRichTextRun[] = [];
  let lastIndex: number | undefined;
  let length = 0;
  const append = (text: string, index: number | undefined) => {
    if (!text) return;
    length += text.length;
    const last = result[result.length - 1];
    if (last && lastIndex === index) {
      last.text += text;
    } else {
      const font = index === undefined ? undefined : source[index]?.font;
      result.push(font ? { text, font } : { text });
      lastIndex = index;
    }
  };
  const walk = (parent: HTMLElement, index?: number) => {
    let previousBlock = false;
    let previousEmptyBlock = false;
    for (const child of Array.from(parent.childNodes)) {
      const element = child.nodeType === 1 ? child as HTMLElement : undefined;
      const block = element?.tagName === 'DIV' || element?.tagName === 'P';
      const last = result[result.length - 1];
      if ((block || previousBlock) && (last || previousBlock) &&
          (!last?.text.endsWith('\n') || previousEmptyBlock)) append('\n', lastIndex);
      const before = length;
      if (child.nodeType === 3) append(child.textContent ?? '', index);
      else if (element?.tagName === 'BR') {
        if (parent.childNodes.length > 1) append('\n', index);
      } else if (element) {
        const runIndex = element.dataset.runIndex === undefined ? index : Number(element.dataset.runIndex);
        walk(element, runIndex);
      }
      previousBlock = block;
      previousEmptyBlock = block && length === before;
    }
  };
  walk(input);
  return result;
}

function editorSelection(input: HTMLElement): Selection | null {
  const root = input.getRootNode() as Document | (ShadowRoot & { getSelection?(): Selection | null });
  return root.getSelection?.() ?? input.ownerDocument.getSelection();
}

export function focusEditableStart(input: HTMLElement): void {
  input.focus({ preventScroll: true });
  const selection = editorSelection(input);
  if (!selection) return;
  const range = input.ownerDocument.createRange();
  const firstText = input.firstElementChild?.firstChild;
  if (firstText) range.setStart(firstText, 0);
  else range.selectNodeContents(input);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
}

export function insertEditableText(input: HTMLElement, text: string): void {
  const doc = input.ownerDocument;
  const value = text.replace(/\r\n?/g, '\n');
  // insertText retains the in-progress editor's native undo queue.
  if (doc.execCommand?.('insertText', false, value)) return;
  const selection = editorSelection(input);
  if (!selection?.rangeCount) return;
  const range = selection.getRangeAt(0);
  if (!input.contains(range.commonAncestorContainer)) return;
  const node = doc.createTextNode(value);
  range.deleteContents();
  range.insertNode(node);
  range.setStartAfter(node);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
}
