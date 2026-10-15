import type { SpreadsheetRect } from './types.js';

export interface SpreadsheetCellEditorOptions {
  textarea: HTMLTextAreaElement;
  onCommit: (value: string) => void | Promise<void>;
  onCancel?: () => void;
}

export class SpreadsheetCellEditor {
  constructor(private readonly options: SpreadsheetCellEditorOptions) {
    this.options.textarea.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        this.cancel();
      } else if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        void this.commit();
      }
    });
  }

  open(rect: SpreadsheetRect, value: string): void {
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

  async commit(): Promise<void> {
    const value = this.options.textarea.value;
    this.close();
    await this.options.onCommit(value);
  }

  cancel(): void {
    this.close();
    this.options.onCancel?.();
  }

  close(): void {
    this.options.textarea.style.display = 'none';
  }
}

