/** Raised when a document mutation is attempted while an editor is read-only. */
export class EditorReadOnlyError extends Error {
  readonly code = 'EDITOR_READ_ONLY';

  constructor(editor: 'word' | 'spreadsheet' | 'ppt') {
    super(`${editor} editor is read-only`);
    this.name = 'EditorReadOnlyError';
  }
}
