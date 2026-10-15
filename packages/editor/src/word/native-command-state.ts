import type { EditorParagraphFormat, EditorRunFormat } from './model.js';
import {
  findNativeParagraph,
  findNativeParagraphLocation,
  findNativeTableCellLocation,
  findNativeTableLocation,
  nativeParagraphText,
  type NativeWordDocument,
  type NativeWordImageNode,
  type NativeWordParagraphNode,
} from './native-model.js';
import { nativeTableCellsInRange, nativeTableSupportsStructuralEditing } from './native-commands.js';
import type {
  WordEditorCommand,
  WordEditorCommandState,
  WordEditorFormatValue,
  WordEditorHyperlink,
  WordEditorSelection,
  WordEditorSelectionPoint,
} from './runtime.js';
import { nativeListType } from './native-numbering.js';

type TextRange = { paraId: string; from: number; to: number };
type Alignment = Exclude<WordEditorCommandState['paragraphAlignment'], 'mixed' | null>;

export function getNativeWordCommandState(
  doc: NativeWordDocument,
  selection: WordEditorSelection | null,
): WordEditorCommandState {
  const runFormats = selectedRunFormats(doc, selection);
  const paragraphs = selectedParagraphs(doc, selection);
  const image = selectedImage(doc, selection);
  const table = selectedTableState(doc, selection);
  return {
    selectionKind: selection?.kind ?? null,
    trackRevisions: false,
    bold: mixedValue(runFormats.map((format) => format.bold ?? false), false) ?? false,
    italic: mixedValue(runFormats.map((format) => format.italic ?? false), false) ?? false,
    underline: mixedValue(runFormats.map((format) => format.underline ?? false), false) ?? false,
    strike: mixedValue(runFormats.map((format) => format.strike ?? false), false) ?? false,
    fontFamily: mixedValue(runFormats.map((format) => format.fontFamily ?? null), null),
    fontSize: mixedValue(runFormats.map((format) => format.fontSize ?? null), null),
    textColor: mixedValue(runFormats.map((format) => normalizeColor(format.color)), null),
    highlightColor: mixedValue(runFormats.map((format) => normalizeHighlight(format.highlight)), null),
    hyperlink: selectedHyperlinkState(doc, selection),
    paragraphAlignment: mixedValue(
      paragraphs.map((paragraph) => normalizeAlignment(paragraph.format?.alignment)),
      null,
    ),
    leftIndent: mixedValue(paragraphs.map((paragraph) => paragraph.format?.indent?.left ?? 0), 0),
    rightIndent: mixedValue(paragraphs.map((paragraph) => paragraph.format?.indent?.right ?? 0), 0),
    firstLineIndent: mixedValue(paragraphs.map((paragraph) => paragraph.format?.indent?.firstLine ?? 0), 0),
    spacingBefore: mixedValue(paragraphs.map((paragraph) => paragraph.format?.spacing?.before ?? 0), 0),
    spacingAfter: mixedValue(paragraphs.map((paragraph) => paragraph.format?.spacing?.after ?? 0), 0),
    lineSpacing: mixedStructuredValue(
      paragraphs.map((paragraph) => {
        const line = paragraph.format?.spacing?.line;
        if (line === undefined) return null;
        return { line, lineRule: paragraph.format?.spacing?.lineRule ?? 'auto' };
      }),
    ),
    listType: mixedValue(
      paragraphs.map((paragraph) =>
        paragraph.format?.hasNumbering ? nativeListType(paragraph.format.numbering) : null,
      ),
      null,
    ),
    listLevel: mixedValue(
      paragraphs.map((paragraph) =>
        paragraph.format?.hasNumbering ? (paragraph.format.numbering?.ilvl ?? 0) : null,
      ),
      null,
    ),
    image: image && selection?.kind === 'object'
      ? {
          objectId: image.drawingId,
          width: image.width ?? null,
          height: image.height ?? null,
          alignment: normalizeAlignment(paragraphs[0]?.format?.alignment),
        }
      : null,
    table,
  };
}

export function canDispatchNativeWordCommand(
  doc: NativeWordDocument,
  selection: WordEditorSelection | null,
  command: WordEditorCommand,
): boolean {
  if (
    command.kind === 'toggle-track-revisions' ||
    command.kind === 'accept-revision' ||
    command.kind === 'reject-revision' ||
    command.kind === 'accept-all-revisions' ||
    command.kind === 'reject-all-revisions'
  ) return true;
  if (!selection) return false;
  if (command.kind === 'insert-image') {
    return (selection.kind === 'caret' || selection.kind === 'text-range') && command.src.trim().length > 0;
  }
  if (command.kind === 'insert-table') {
    if (
      selection.kind !== 'caret' ||
      !Number.isInteger(command.rows) ||
      !Number.isInteger(command.columns) ||
      command.rows < 1 ||
      command.rows > 20 ||
      command.columns < 1 ||
      command.columns > 20
    ) return false;
    return findNativeParagraphLocation(doc, selection.paraId)?.container === doc.blocks;
  }
  if (command.kind === 'delete-table') {
    const table = selectedTableState(doc, selection);
    return !!table && findNativeTableLocation(doc, table.tableId)?.container === doc.blocks;
  }
  if (command.kind === 'set-image-size' || command.kind === 'set-image-alignment') {
    return (
      selection.kind === 'object' &&
      selection.objectKind === 'image' &&
      selection.objectId === command.objectId &&
      !!selectedImage(doc, selection)
    );
  }
  if (
    command.kind === 'insert-table-row' ||
    command.kind === 'delete-table-row' ||
    command.kind === 'insert-table-column' ||
    command.kind === 'delete-table-column'
  ) {
    const table = selectedTableState(doc, selection);
    if (!table?.structuralEditingSupported) return false;
    if (command.kind === 'delete-table-row') return table.rowCount > 1;
    if (command.kind === 'delete-table-column') return table.columnCount > 1;
    return true;
  }
  if (
    command.kind === 'set-table-cell-background' ||
    command.kind === 'set-table-cell-vertical-alignment' ||
    command.kind === 'set-table-cell-borders'
  ) {
    return selectedTableState(doc, selection) !== null;
  }
  if (command.kind === 'merge-table-cells' || command.kind === 'split-table-cell') {
    if (selection.kind !== 'table-cell-range') return false;
    const location = findNativeTableLocation(doc, selection.tableId);
    if (!location) return false;
    const cells = nativeTableCellsInRange(location.table, selection);
    if (command.kind === 'merge-table-cells') {
      if (!nativeTableSupportsStructuralEditing(location.table)) return false;
      const sameRow = selection.anchor.rowIndex === selection.focus.rowIndex;
      return sameRow && cells.length >= 2 && cells.every((item) => item.cell.gridSpan === 1 && item.cell.vMerge == null);
    }
    const cell = cells.length === 1 ? cells[0]!.cell : null;
    return !!cell && (
      (cell.gridSpan > 1 && cell.vMerge == null && (cell.rowSpan ?? 1) === 1) ||
      (cell.vMerge === 'restart' && (cell.rowSpan ?? 1) > 1)
    );
  }
  if (command.kind === 'clear-table-cells') {
    return selection.kind === 'table-cell-range' && selectedTableState(doc, selection) !== null;
  }
  if (command.kind === 'set-hyperlink') {
    return selection.kind === 'text-range' && !!(command.hyperlink.target?.trim() || command.hyperlink.anchor?.trim());
  }
  if (command.kind === 'remove-hyperlink') {
    return selection.kind === 'text-range' && selectedHyperlinkState(doc, selection) !== null;
  }
  if (command.kind === 'insert-text' || command.kind === 'split-paragraph' || command.kind === 'insert-page-break') {
    return selection.kind === 'caret' && !!findNativeParagraph(doc, selection.paraId);
  }
  if (command.kind === 'delete-object') return selection.kind === 'object' && selection.objectKind === command.objectKind;
  if (command.kind === 'replace-range' || command.kind === 'delete-range') {
    return (
      command.anchor.paraId === command.focus.paraId &&
      !!findNativeParagraph(doc, command.anchor.paraId)
    );
  }
  if (
    command.kind === 'toggle-bold' ||
    command.kind === 'toggle-italic' ||
    command.kind === 'toggle-underline' ||
    command.kind === 'toggle-strike' ||
    command.kind === 'set-font-family' ||
    command.kind === 'set-font-size' ||
    command.kind === 'set-text-color' ||
    command.kind === 'set-highlight-color' ||
    command.kind === 'clear-text-format'
  ) {
    return selection.kind === 'text-range' && selectedTextRanges(doc, selection).some((range) => range.to > range.from);
  }
  if (
    command.kind === 'set-paragraph-alignment' ||
    command.kind === 'set-paragraph-indent' ||
    command.kind === 'adjust-paragraph-left-indent' ||
    command.kind === 'set-paragraph-spacing' ||
    command.kind === 'set-paragraph-line-spacing'
    || command.kind === 'toggle-list'
    || command.kind === 'remove-list'
    || command.kind === 'adjust-list-level'
  ) {
    return selectedParagraphs(doc, selection).length > 0;
  }
  if (command.kind === 'delete-backward') {
    return selection.kind === 'caret' || selection.kind === 'object' || selection.kind === 'table-cell-range';
  }
  return command.kind === 'delete-forward' && (selection.kind === 'caret' || selection.kind === 'table-cell-range');
}

function selectedTableState(
  doc: NativeWordDocument,
  selection: WordEditorSelection | null,
): WordEditorCommandState['table'] {
  if (!selection) return null;
  if (selection.kind === 'table-cell-range') {
    const location = findNativeTableLocation(doc, selection.tableId);
    if (!location) return null;
    const cells = nativeTableCellsInRange(location.table, selection);
    if (cells.length === 0) return null;
    const startRow = Math.min(selection.anchor.rowIndex, selection.focus.rowIndex);
    const endRow = Math.max(selection.anchor.rowIndex, selection.focus.rowIndex);
    const startColumn = Math.min(selection.anchor.columnIndex, selection.focus.columnIndex);
    const endColumn = Math.max(selection.anchor.columnIndex, selection.focus.columnIndex);
    return {
      tableId: location.table.tableId,
      rowIndex: selection.anchor.rowIndex,
      rowCount: location.table.rows.length,
      columnIndex: selection.anchor.columnIndex,
      columnCount: Math.max(0, ...location.table.rows.map((row) =>
        row.cells.reduce((count, cell) => count + Math.max(1, cell.gridSpan), 0))),
      range: { startRow, endRow, startColumn, endColumn, cellCount: cells.length },
      structuralEditingSupported: nativeTableSupportsStructuralEditing(location.table),
      backgroundColor: mixedValue(cells.map(({ cell }) => normalizeColor(cell.shading?.fill)), null),
      verticalAlignment: mixedValue(cells.map(({ cell }) => cell.verticalAlignment ?? null), null),
      borderColor: mixedValue(cells.map(({ cell }) => selectedCellBorderColor(cell.borders)), null),
      bordersVisible: mixedValue(cells.map(({ cell }) => selectedCellBordersVisible(cell.borders)), false),
    };
  }
  const paraId = selection.kind === 'text-range' ? selection.anchor.paraId : selection.paraId;
  const location = findNativeTableCellLocation(doc, paraId);
  if (!location) return null;
  return {
    tableId: location.table.tableId,
    rowIndex: location.rowIndex,
    rowCount: location.table.rows.length,
    columnIndex: location.columnIndex,
    columnCount: location.table.rows[0]?.cells.length ?? 0,
    range: null,
    structuralEditingSupported: nativeTableSupportsStructuralEditing(location.table),
    backgroundColor: normalizeColor(location.cell.shading?.fill),
    verticalAlignment: location.cell.verticalAlignment ?? null,
    borderColor: selectedCellBorderColor(location.cell.borders),
    bordersVisible: selectedCellBordersVisible(location.cell.borders),
  };
}

function selectedCellBorderColor(
  borders: import('./model.js').EditorTableBorders | undefined,
): string | null {
  const border = [borders?.top, borders?.right, borders?.bottom, borders?.left]
    .find((candidate) => candidate && candidate.val !== 'nil' && candidate.val !== 'none');
  return normalizeColor(border?.color);
}

function selectedCellBordersVisible(
  borders: import('./model.js').EditorTableBorders | undefined,
): boolean {
  return [borders?.top, borders?.right, borders?.bottom, borders?.left]
    .some((border) => !!border && border.val !== 'nil' && border.val !== 'none');
}

function selectedImage(
  doc: NativeWordDocument,
  selection: WordEditorSelection | null,
): NativeWordImageNode | null {
  if (selection?.kind !== 'object' || selection.objectKind !== 'image') return null;
  const paragraph = findNativeParagraph(doc, selection.paraId);
  return paragraph?.content.find(
    (child): child is NativeWordImageNode => child.kind === 'image' && child.drawingId === selection.objectId,
  ) ?? null;
}

function selectedRunFormats(doc: NativeWordDocument, selection: WordEditorSelection | null): EditorRunFormat[] {
  if (!selection || selection.kind === 'object' || selection.kind === 'table-cell-range') return [];
  if (selection.kind === 'caret') {
    const paragraph = findNativeParagraph(doc, selection.paraId);
    if (!paragraph) return [];
    const format = runFormatAt(doc, paragraph, selection.offset);
    return format ? [format] : [];
  }
  return selectedTextRanges(doc, selection).flatMap((range) => runFormatsInRange(doc, range));
}

function selectedHyperlinkState(
  doc: NativeWordDocument,
  selection: WordEditorSelection | null,
): WordEditorCommandState['hyperlink'] {
  if (!selection || selection.kind === 'object' || selection.kind === 'table-cell-range') return null;
  if (selection.kind === 'caret') {
    const paragraph = findNativeParagraph(doc, selection.paraId);
    if (!paragraph) return null;
    let position = 0;
    let previous: NativeWordParagraphNode['content'][number] | null = null;
    for (const child of paragraph.content) {
      const length = inlineLength(child);
      if (child.kind === 'text') {
        if (selection.offset > position && selection.offset <= position + length) return child.hyperlink ?? null;
        if (selection.offset === position) {
          return previous?.kind === 'text'
            ? previous.hyperlink ?? child.hyperlink ?? null
            : child.hyperlink ?? null;
        }
        previous = child;
      }
      position += length;
    }
    return previous?.kind === 'text' ? previous.hyperlink ?? null : null;
  }
  const links: Array<WordEditorHyperlink | null> = selectedTextRanges(doc, selection).flatMap((range) => {
    const paragraph = findNativeParagraph(doc, range.paraId);
    if (!paragraph) return [];
    const values: Array<WordEditorHyperlink | null> = [];
    let position = 0;
    for (const child of paragraph.content) {
      const length = inlineLength(child);
      if (child.kind === 'text' && position + length > range.from && position < range.to) {
        values.push(child.hyperlink ?? null);
      }
      position += length;
    }
    return values;
  });
  return mixedStructuredValue(links);
}

function runFormatsInRange(doc: NativeWordDocument, range: TextRange): EditorRunFormat[] {
  const paragraph = findNativeParagraph(doc, range.paraId);
  if (!paragraph) return [];
  const formats: EditorRunFormat[] = [];
  let position = 0;
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind === 'text' && position + length > range.from && position < range.to) {
      formats.push(effectiveRunFormat(doc, paragraph, child.format));
    }
    position += length;
  }
  return formats;
}

function runFormatAt(
  doc: NativeWordDocument,
  paragraph: NativeWordParagraphNode,
  offset: number,
): EditorRunFormat | null {
  let position = 0;
  let previous: EditorRunFormat | null = null;
  for (const child of paragraph.content) {
    const length = inlineLength(child);
    if (child.kind === 'text') {
      const format = effectiveRunFormat(doc, paragraph, child.format);
      if (offset > position && offset <= position + length) return format;
      if (offset === position) return previous ?? format;
      previous = format;
    }
    position += length;
  }
  return previous ?? effectiveRunFormat(doc, paragraph, null);
}

function effectiveRunFormat(
  doc: NativeWordDocument,
  paragraph: NativeWordParagraphNode,
  run: EditorRunFormat | null,
): EditorRunFormat {
  const paragraphRun = paragraph.format as EditorRunFormat | null;
  return { ...(doc.defaults?.run ?? {}), ...(paragraphRun ?? {}), ...(run ?? {}) };
}

function selectedParagraphs(doc: NativeWordDocument, selection: WordEditorSelection | null): NativeWordParagraphNode[] {
  if (!selection) return [];
  if (selection.kind === 'table-cell-range') return [];
  if (selection.kind === 'caret' || selection.kind === 'object') {
    const paragraph = findNativeParagraph(doc, selection.paraId);
    return paragraph ? [paragraph] : [];
  }
  return selectedTextRanges(doc, selection)
    .map((range) => findNativeParagraph(doc, range.paraId))
    .filter((paragraph): paragraph is NativeWordParagraphNode => paragraph !== null);
}

function selectedTextRanges(
  doc: NativeWordDocument,
  selection: Extract<WordEditorSelection, { kind: 'text-range' }>,
): TextRange[] {
  const ordered = orderedSelection(doc, selection);
  if (!ordered) return [];
  if (ordered.anchor.paraId === ordered.focus.paraId) {
    return [{ paraId: ordered.anchor.paraId, from: ordered.anchor.offset, to: ordered.focus.offset }];
  }
  const start = findNativeParagraphLocation(doc, ordered.anchor.paraId);
  const end = findNativeParagraphLocation(doc, ordered.focus.paraId);
  if (!start || !end || start.container !== end.container) return [];
  const ranges: TextRange[] = [];
  for (let index = start.index; index <= end.index; index += 1) {
    const block = start.container[index];
    if (block?.kind !== 'paragraph') return [];
    ranges.push({
      paraId: block.paraId,
      from: index === start.index ? ordered.anchor.offset : 0,
      to: index === end.index ? ordered.focus.offset : nativeParagraphText(block).length,
    });
  }
  return ranges;
}

function orderedSelection(
  doc: NativeWordDocument,
  selection: Extract<WordEditorSelection, { kind: 'text-range' }>,
): { anchor: WordEditorSelectionPoint; focus: WordEditorSelectionPoint } | null {
  if (selection.anchor.paraId === selection.focus.paraId) {
    return selection.anchor.offset <= selection.focus.offset
      ? { anchor: selection.anchor, focus: selection.focus }
      : { anchor: selection.focus, focus: selection.anchor };
  }
  const anchor = findNativeParagraphLocation(doc, selection.anchor.paraId);
  const focus = findNativeParagraphLocation(doc, selection.focus.paraId);
  if (!anchor || !focus || anchor.container !== focus.container) return null;
  return anchor.index < focus.index
    ? { anchor: selection.anchor, focus: selection.focus }
    : { anchor: selection.focus, focus: selection.anchor };
}

function inlineLength(child: NativeWordParagraphNode['content'][number]): number {
  if (child.kind === 'text') return child.text.length;
  if (child.kind === 'field') return (child.cachedText ?? '\uFFFC').length;
  if (child.kind === 'drawing' && child.text) return child.text.length;
  return 1;
}

function mixedValue<T>(values: T[], empty: null): WordEditorFormatValue<T>;
function mixedValue<T>(values: T[], empty: T): T | 'mixed';
function mixedValue<T>(values: T[], empty: T | null): WordEditorFormatValue<T> {
  if (values.length === 0) return empty;
  return values.every((value) => value === values[0]) ? values[0]! : 'mixed';
}

function mixedStructuredValue<T>(values: Array<T | null>): WordEditorFormatValue<T> {
  if (values.length === 0 || values.every((value) => value === null)) return null;
  const first = JSON.stringify(values[0]);
  return values.every((value) => JSON.stringify(value) === first) ? (values[0] ?? null) : 'mixed';
}

function normalizeColor(color: string | undefined): string | null {
  if (!color || color.toLowerCase() === 'auto') return null;
  const value = color.replace(/^#/, '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(value) ? `#${value}` : color;
}

function normalizeHighlight(highlight: string | undefined): string | null {
  if (!highlight || highlight.toLowerCase() === 'none') return null;
  return normalizeColor(highlight) ?? highlight;
}

function normalizeAlignment(alignment: EditorParagraphFormat['alignment']): Alignment {
  if (alignment === 'center' || alignment === 'right' || alignment === 'both') return alignment;
  if (alignment === 'distribute') return 'both';
  return 'left';
}
