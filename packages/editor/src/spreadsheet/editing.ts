import type { ExcelRichTextRun } from '@coco-vexel/coco-office/excel';
import type { SpreadsheetRect } from './types.js';
import { focusEditableStart, insertEditableText, readEditableRuns, writeEditableRuns, type SpreadsheetEditableRun } from './rich-text-editing.js';

export interface SpreadsheetCellEditorOptions {
  textarea: HTMLTextAreaElement;
  richFrame: HTMLElement;
  richInput: HTMLElement;
  onCommit: (value: string, runs?: readonly ExcelRichTextRun[]) => void | Promise<void>;
  onInput?: (value: string) => void;
  onCancel?: () => void;
  onNavigate?: (key: 'Enter' | 'Tab', reverse: boolean) => void | Promise<void>;
  onError?: (error: unknown) => void;
}

export class SpreadsheetCellEditor {
  private active = false;
  private richRuns: readonly SpreadsheetEditableRun[] | undefined;
  private pendingCommit: Promise<void> = Promise.resolve();

  constructor(private readonly options: SpreadsheetCellEditorOptions) {
    for (const input of [options.textarea, options.richInput]) {
      input.addEventListener('keydown', this.handleKeyDown);
      input.addEventListener('blur', this.handleBlur);
      input.addEventListener('input', this.handleInput);
    }
    options.richInput.addEventListener('paste', this.handleRichPaste);
  }

  isRich(): boolean {
    return this.richRuns !== undefined;
  }

  contains(event: Event): boolean {
    const path = event.composedPath();
    return path.includes(this.options.textarea) || path.includes(this.options.richFrame);
  }

  open(rect: SpreadsheetRect, value: string, richRuns?: readonly SpreadsheetEditableRun[]): void {
    this.active = true;
    this.richRuns = richRuns;
    if (richRuns) {
      const { richFrame, richInput } = this.options;
      writeEditableRuns(richInput, richRuns);
      this.place(rect);
      focusEditableStart(richInput);
      richFrame.scrollTop = 0;
      richFrame.scrollLeft = 0;
      return;
    }
    const textarea = this.options.textarea;
    textarea.value = value;
    this.place(rect);
    textarea.focus({ preventScroll: true });
    textarea.setSelectionRange(value.length, value.length);
  }

  place(rect: SpreadsheetRect | null): void {
    if (!this.active) return;
    const surface = this.richRuns ? this.options.richFrame : this.options.textarea;
    if (!rect) {
      surface.style.display = 'none';
      return;
    }
    Object.assign(surface.style, {
      display: this.richRuns ? 'flex' : 'block',
      left: `${rect.x}px`, top: `${rect.y}px`,
      width: `${Math.max(1, rect.width)}px`, height: `${Math.max(1, rect.height)}px`,
    });
  }

  commit(navigation?: { key: 'Enter' | 'Tab'; reverse: boolean }): Promise<void> {
    if (!this.active) return this.pendingCommit;
    this.active = false;
    const runs = this.richRuns && readEditableRuns(this.options.richInput, this.richRuns);
    const value = runs ? runs.map((run) => run.text).join('') : this.options.textarea.value;
    try {
      this.pendingCommit = Promise.resolve(this.options.onCommit(value, runs));
      if (navigation) this.pendingCommit = this.pendingCommit.then(() => this.options.onNavigate?.(navigation.key, navigation.reverse));
    } catch (error) {
      this.pendingCommit = Promise.reject(error);
    }
    // Mark inactive and start the commit before hiding: hiding may synchronously blur.
    this.close();
    return this.pendingCommit;
  }

  cancel(): void {
    this.active = false;
    this.close();
    this.options.onCancel?.();
  }

  close(): void {
    this.options.textarea.style.display = 'none';
    this.options.richFrame.style.display = 'none';
    this.richRuns = undefined;
  }

  destroy(): void {
    this.active = false;
    for (const input of [this.options.textarea, this.options.richInput]) {
      input.removeEventListener('keydown', this.handleKeyDown);
      input.removeEventListener('blur', this.handleBlur);
      input.removeEventListener('input', this.handleInput);
    }
    this.options.richInput.removeEventListener('paste', this.handleRichPaste);
    this.close();
  }

  private readonly handleBlur = (): void => {
    void this.commit().catch((error) => this.options.onError?.(error));
  };

  private readonly handleInput = (): void => {
    if (!this.active || !this.options.onInput) return;
    const value = this.richRuns ? this.options.richInput.textContent ?? '' : this.options.textarea.value;
    this.options.onInput(value);
  };

  private readonly handleRichPaste = (event: ClipboardEvent): void => {
    event.preventDefault();
    if (event.clipboardData?.types.includes('text/plain')) {
      insertEditableText(this.options.richInput, event.clipboardData.getData('text/plain'));
      this.handleInput();
    }
  };

  private readonly handleKeyDown = (event: KeyboardEvent): void => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      this.cancel();
    } else if (this.richRuns && event.key === 'Enter' && event.altKey && !event.ctrlKey && !event.metaKey) {
      event.preventDefault();
      event.stopPropagation();
      insertEditableText(this.options.richInput, '\n');
      this.handleInput();
    } else if ((event.key === 'Enter' || event.key === 'Tab') && !event.altKey && !event.ctrlKey && !event.metaKey) {
      event.preventDefault();
      event.stopPropagation();
      void this.commit({ key: event.key, reverse: event.shiftKey }).catch((error) => this.options.onError?.(error));
    }
  };
}

