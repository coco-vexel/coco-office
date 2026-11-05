import type { SpreadsheetRect } from './types.js';

export interface SpreadsheetCellEditorOptions {
  textarea: HTMLTextAreaElement;
  onCommit: (value: string) => void | Promise<void>;
  onCancel?: () => void;
  onNavigate?: (key: 'Enter' | 'Tab', reverse: boolean) => void | Promise<void>;
  onError?: (error: unknown) => void;
}

export class SpreadsheetCellEditor {
  private active = false;
  private pendingCommit: Promise<void> = Promise.resolve();

  constructor(private readonly options: SpreadsheetCellEditorOptions) {
    options.textarea.addEventListener('keydown', this.handleKeyDown);
    options.textarea.addEventListener('blur', this.handleBlur);
  }

  open(rect: SpreadsheetRect, value: string): void {
    this.active = true;
    const textarea = this.options.textarea;
    textarea.value = value;
    textarea.style.display = 'block';
    textarea.style.position = 'absolute';
    textarea.style.left = `${rect.x}px`;
    textarea.style.top = `${rect.y}px`;
    textarea.style.width = `${Math.max(1, rect.width)}px`;
    textarea.style.height = `${Math.max(1, rect.height)}px`;
    textarea.focus();
    textarea.select();
  }

  commit(navigation?: { key: 'Enter' | 'Tab'; reverse: boolean }): Promise<void> {
    if (!this.active) return this.pendingCommit;
    this.active = false;
    const value = this.options.textarea.value;
    try {
      this.pendingCommit = Promise.resolve(this.options.onCommit(value));
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
  }

  destroy(): void {
    this.active = false;
    this.options.textarea.removeEventListener('keydown', this.handleKeyDown);
    this.options.textarea.removeEventListener('blur', this.handleBlur);
    this.close();
  }

  private readonly handleBlur = (): void => {
    void this.commit().catch((error) => this.options.onError?.(error));
  };

  private readonly handleKeyDown = (event: KeyboardEvent): void => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      this.cancel();
    } else if ((event.key === 'Enter' || event.key === 'Tab') && !event.altKey && !event.ctrlKey && !event.metaKey) {
      event.preventDefault();
      event.stopPropagation();
      void this.commit({ key: event.key, reverse: event.shiftKey }).catch((error) => this.options.onError?.(error));
    }
  };
}

