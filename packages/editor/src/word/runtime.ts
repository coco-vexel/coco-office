import type { EditorDocumentView } from './model.js';
import type { NativeWordChange } from './native-commands.js';
import type { NativeWordDocument } from './native-model.js';

export type WordEditorEngineKind = 'native-word';

export type WordEditorSelection =
  | {
      kind: 'caret';
      paraId: string;
      offset: number;
    }
  | {
      kind: 'text-range';
      anchor: WordEditorSelectionPoint;
      focus: WordEditorSelectionPoint;
    }
  | {
      kind: 'object';
      paraId: string;
      objectKind: 'image';
      objectId: string;
      offset: number;
    }
  | {
      kind: 'table-cell-range';
      tableId: string;
      anchor: WordEditorTableCellPoint;
      focus: WordEditorTableCellPoint;
    };

export interface WordEditorSelectionPoint {
  paraId: string;
  offset: number;
}

export interface WordEditorTableCellPoint {
  rowIndex: number;
  columnIndex: number;
}

export interface WordEditorSearchOptions {
  query: string;
  matchCase?: boolean;
  wholeWord?: boolean;
}

export interface WordEditorSearchMatch {
  id: string;
  paraId: string;
  from: number;
  to: number;
  text: string;
}

export interface WordEditorSearchState {
  options: WordEditorSearchOptions;
  matches: readonly WordEditorSearchMatch[];
  activeIndex: number;
}

export interface WordEditorHyperlink {
  target?: string;
  anchor?: string;
  tooltip?: string;
}

export type WordEditorFormatValue<T> = T | 'mixed' | null;

export interface WordEditorLineSpacingState {
  line: number;
  lineRule: 'auto' | 'exact' | 'atLeast';
}

export interface WordEditorCommandState {
  selectionKind: WordEditorSelection['kind'] | null;
  trackRevisions: boolean;
  bold: boolean | 'mixed';
  italic: boolean | 'mixed';
  underline: boolean | 'mixed';
  strike: boolean | 'mixed';
  fontFamily: WordEditorFormatValue<string>;
  fontSize: WordEditorFormatValue<number>;
  textColor: WordEditorFormatValue<string>;
  highlightColor: WordEditorFormatValue<string>;
  hyperlink: WordEditorFormatValue<WordEditorHyperlink>;
  paragraphAlignment: WordEditorFormatValue<'left' | 'center' | 'right' | 'both'>;
  leftIndent: number | 'mixed';
  rightIndent: number | 'mixed';
  firstLineIndent: number | 'mixed';
  spacingBefore: number | 'mixed';
  spacingAfter: number | 'mixed';
  lineSpacing: WordEditorFormatValue<WordEditorLineSpacingState>;
  listType: WordEditorFormatValue<'bullet' | 'numbered'>;
  listLevel: WordEditorFormatValue<number>;
  image: WordEditorImageState | null;
  table: WordEditorTableState | null;
}

export interface WordEditorImageState {
  objectId: string;
  width: number | null;
  height: number | null;
  alignment: 'left' | 'center' | 'right' | 'both' | null;
}

export interface WordEditorImageInsertOptions {
  width?: number;
  height?: number;
}

export interface WordEditorTableState {
  tableId: string;
  rowIndex: number;
  rowCount: number;
  columnIndex: number;
  columnCount: number;
  range: {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
    cellCount: number;
  } | null;
  structuralEditingSupported: boolean;
  backgroundColor: WordEditorFormatValue<string>;
  verticalAlignment: WordEditorFormatValue<'top' | 'center' | 'bottom' | 'both'>;
  borderColor: WordEditorFormatValue<string>;
  bordersVisible: boolean | 'mixed';
}

export type WordEditorRuntimeEvent = 'change' | 'selectionChange';

export interface WordEditorChangeEvent {
  readonly revision: number;
  readonly changes: readonly NativeWordChange[];
  readonly revertedChanges: readonly NativeWordChange[];
  readonly totalChangeCount: number;
}

export interface WordEditorSelectionChangeEvent {
  readonly selection: WordEditorSelection | null;
  readonly commandState: WordEditorCommandState;
}

export interface WordEditorRuntimeEventMap {
  change: WordEditorChangeEvent;
  selectionChange: WordEditorSelectionChangeEvent;
}

export type WordEditorCommand =
  | {
      kind: 'insert-text';
      text: string;
    }
  | {
      kind: 'delete-backward';
    }
  | {
      kind: 'delete-forward';
    }
  | {
      kind: 'delete-range';
      anchor: WordEditorSelectionPoint;
      focus: WordEditorSelectionPoint;
    }
  | {
      kind: 'replace-range';
      anchor: WordEditorSelectionPoint;
      focus: WordEditorSelectionPoint;
      text: string;
    }
  | {
      kind: 'split-paragraph';
    }
  | {
      kind: 'insert-page-break';
    }
  | {
      kind: 'delete-object';
      objectKind: 'image';
      objectId: string;
    }
  | {
      kind: 'insert-image';
      src: string;
      width?: number;
      height?: number;
    }
  | {
      kind: 'set-image-size';
      objectId: string;
      width: number;
      height: number;
    }
  | {
      kind: 'set-image-alignment';
      objectId: string;
      alignment: 'left' | 'center' | 'right';
    }
  | {
      kind: 'insert-table-row';
      position: 'before' | 'after';
    }
  | {
      kind: 'insert-table';
      rows: number;
      columns: number;
    }
  | {
      kind: 'delete-table';
    }
  | {
      kind: 'delete-table-row';
    }
  | {
      kind: 'insert-table-column';
      position: 'before' | 'after';
    }
  | {
      kind: 'delete-table-column';
    }
  | {
      kind: 'merge-table-cells';
    }
  | {
      kind: 'split-table-cell';
    }
  | {
      kind: 'set-table-cell-background';
      color: string | null;
    }
  | {
      kind: 'set-table-cell-vertical-alignment';
      alignment: 'top' | 'center' | 'bottom' | null;
    }
  | {
      kind: 'set-table-cell-borders';
      color: string | null;
    }
  | {
      kind: 'clear-table-cells';
    }
  | {
      kind: 'set-hyperlink';
      hyperlink: WordEditorHyperlink;
    }
  | {
      kind: 'remove-hyperlink';
    }
  | {
      kind: 'add-comment';
      body: string;
      author?: string;
      date?: string;
    }
  | {
      kind: 'edit-comment';
      commentId: number;
      body: string;
    }
  | {
      kind: 'delete-comment';
      commentId: number;
    }
  | {
      kind: 'reply-comment';
      commentId: number;
      body: string;
    }
  | {
      kind: 'add-revision';
      revisionKind: 'insert' | 'delete';
      author?: string;
      date?: string;
    }
  | {
      kind: 'toggle-track-revisions';
    }
  | {
      kind: 'accept-revision';
      revisionId?: string;
      revisionKind?: 'insert' | 'delete' | 'format';
    }
  | {
      kind: 'reject-revision';
      revisionId?: string;
      revisionKind?: 'insert' | 'delete' | 'format';
    }
  | {
      kind: 'accept-all-revisions' | 'reject-all-revisions';
    }
  | {
      kind: 'toggle-bold';
    }
  | {
      kind: 'toggle-italic';
    }
  | {
      kind: 'toggle-underline';
    }
  | {
      kind: 'toggle-strike';
    }
  | {
      kind: 'set-font-family';
      fontFamily: string;
    }
  | {
      kind: 'set-font-size';
      fontSize: number;
    }
  | {
      kind: 'set-text-color';
      color: string;
    }
  | {
      kind: 'set-highlight-color';
      color: string | null;
    }
  | {
      kind: 'clear-text-format';
    }
  | {
      kind: 'set-paragraph-alignment';
      alignment: 'left' | 'center' | 'right' | 'both';
    }
  | {
      kind: 'set-paragraph-indent';
      indent: { left?: number; right?: number; firstLine?: number };
    }
  | {
      kind: 'adjust-paragraph-left-indent';
      delta: number;
    }
  | {
      kind: 'set-paragraph-spacing';
      spacing: { before?: number; after?: number };
    }
  | {
      kind: 'set-paragraph-line-spacing';
      line: number;
      lineRule: 'auto' | 'exact' | 'atLeast';
    }
  | {
      kind: 'toggle-list';
      listType: 'bullet' | 'numbered';
    }
  | {
      kind: 'remove-list';
    }
  | {
      kind: 'adjust-list-level';
      delta: number;
    };

export interface WordEditorRuntime {
  readonly kind: WordEditorEngineKind;
  readonly destroyed: boolean;
  destroy(): void;
  loadDocument(view: EditorDocumentView): void;
  getSourceView(): EditorDocumentView;
  getDocument(): NativeWordDocument;
  getSelection(): WordEditorSelection | null;
  setSelection(selection: WordEditorSelection | null): boolean;
  focus(): boolean;
  insertImage(source: string | Blob, options?: WordEditorImageInsertOptions): Promise<boolean>;
  search(options: WordEditorSearchOptions): WordEditorSearchState;
  getSearchState(): WordEditorSearchState | null;
  findNext(direction?: 'forward' | 'backward'): WordEditorSearchState | null;
  replaceCurrent(replacement: string): boolean;
  replaceAll(replacement: string): number;
  clearSearch(): void;
  getCommandState(): WordEditorCommandState;
  getChanges(): readonly NativeWordChange[];
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): boolean;
  redo(): boolean;
  canDispatch(command: WordEditorCommand): boolean;
  dispatch(command: WordEditorCommand): boolean;
  subscribe<Event extends WordEditorRuntimeEvent>(
    event: Event,
    listener: (payload: WordEditorRuntimeEventMap[Event]) => void,
  ): () => void;
}

export interface WordEditorRuntimeOptions {
  onChange?: (event: WordEditorChangeEvent) => void;
  onSelectionChange?: (event: WordEditorSelectionChangeEvent) => void;
  contextMenu?: boolean;
  trackRevisions?: boolean;
  /** Keep the editor DOM and styles inside a component boundary by default. */
  styleIsolation?: 'shadow' | 'none';
  /** CSS text to install inside the ShadowRoot when styleIsolation is 'shadow'. */
  styleText?: string;
}
