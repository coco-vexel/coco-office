import type { EditorDocumentView, EditorParagraphView } from './model.js';
import type { NativeWordDocument, NativeWordParagraphNode } from './native-model.js';

export interface WordEditorCoverage {
  topLevelBlocks: number;
  topLevelParagraphs: number;
  tableBlocks: number;
  tableParagraphs: number;
  allParagraphs: number;
  topLevelImages: number;
  tableImages: number;
  allImages: number;
  topLevelDrawings: number;
  tableDrawings: number;
  allDrawings: number;
  pageBreaks: number;
}

export interface NativeWordCoverage {
  editor: WordEditorCoverage;
  native: {
    blocks: number;
    paragraphs: number;
    images: number;
    drawings: number;
    pageBreaks: number;
  };
  missing: {
    paragraphs: number;
    images: number;
    drawings: number;
    pageBreaks: number;
  };
}

export function summarizeEditorCoverage(wordView: EditorDocumentView): WordEditorCoverage {
  const topLevelParagraphs = wordView.blocks.filter((block) => block.type === 'paragraph');
  const tableParagraphs = collectEditorTableParagraphs(wordView);
  const allParagraphs = collectEditorParagraphs(wordView);
  return {
    topLevelBlocks: wordView.blocks.length,
    topLevelParagraphs: topLevelParagraphs.length,
    tableBlocks: wordView.tables.length,
    tableParagraphs: tableParagraphs.length,
    allParagraphs: allParagraphs.length,
    topLevelImages: countImages(topLevelParagraphs),
    tableImages: countImages(tableParagraphs),
    allImages: countImages(allParagraphs),
    topLevelDrawings: countDrawings(topLevelParagraphs),
    tableDrawings: countDrawings(tableParagraphs),
    allDrawings: countDrawings(allParagraphs),
    pageBreaks: countPageBreaks(allParagraphs),
  };
}

export function summarizeNativeCoverage(
  wordView: EditorDocumentView,
  nativeDocument: NativeWordDocument | undefined,
): NativeWordCoverage {
  const editor = summarizeEditorCoverage(wordView);
  const nativeParagraphs = nativeDocument ? collectNativeBodyParagraphs(nativeDocument) : [];
  const nativeImages = countNativeImages(nativeParagraphs);
  return {
    editor,
    native: {
      blocks: nativeDocument?.blocks.length ?? 0,
      paragraphs: nativeParagraphs.length,
      images: nativeImages,
      drawings: countNativeDrawings(nativeParagraphs),
      pageBreaks: countNativePageBreaks(nativeParagraphs),
    },
    missing: {
      paragraphs: editor.allParagraphs - nativeParagraphs.length,
      images: editor.allImages - nativeImages,
      drawings: editor.allDrawings - countNativeDrawings(nativeParagraphs),
      pageBreaks: editor.pageBreaks - countNativePageBreaks(nativeParagraphs),
    },
  };
}

export function collectNativeParagraphs(nativeDocument: NativeWordDocument): NativeWordParagraphNode[] {
  return [
    ...collectNativeParagraphsFromBlocks(nativeDocument.blocks),
    ...nativeDocument.headers.flatMap((story) => collectNativeParagraphsFromBlocks(story.blocks)),
    ...nativeDocument.footers.flatMap((story) => collectNativeParagraphsFromBlocks(story.blocks)),
  ];
}

function collectNativeBodyParagraphs(nativeDocument: NativeWordDocument): NativeWordParagraphNode[] {
  return collectNativeParagraphsFromBlocks(nativeDocument.blocks);
}

function collectNativeParagraphsFromBlocks(blocks: NativeWordDocument['blocks']): NativeWordParagraphNode[] {
  return blocks.flatMap((block) => {
    if (block.kind === 'paragraph') return [block];
    return block.rows.flatMap((row) =>
      row.cells.flatMap((cell) =>
        cell.vMerge === 'continue' ? [] : collectNativeParagraphsFromBlocks(cell.blocks),
      ),
    );
  });
}

function collectEditorParagraphs(wordView: EditorDocumentView): EditorParagraphView[] {
  const topLevelParagraphs = wordView.blocks.filter(
    (block): block is EditorParagraphView => block.type === 'paragraph',
  );
  return [
    ...topLevelParagraphs,
    ...collectEditorTableParagraphs(wordView),
  ];
}

function collectEditorTableParagraphs(wordView: EditorDocumentView): EditorParagraphView[] {
  return collectEditorTableParagraphsFromTables(wordView.tables);
}

function collectEditorTableParagraphsFromTables(tables: EditorDocumentView['tables']): EditorParagraphView[] {
  return tables.flatMap((table) =>
    table.rows.flatMap((row) =>
      row.cells.flatMap((cell) => (cell.vMerge === 'continue' ? [] : cell.paragraphs)),
    ),
  );
}

function countImages(paragraphs: EditorParagraphView[]): number {
  return paragraphs.reduce(
    (count, paragraph) => count + paragraph.segments.filter((segment) => segment.kind === 'image').length,
    0,
  );
}

function countDrawings(paragraphs: EditorParagraphView[]): number {
  return paragraphs.reduce(
    (count, paragraph) => count + paragraph.segments.filter((segment) => segment.kind === 'drawing').length,
    0,
  );
}

function countNativeImages(paragraphs: NativeWordParagraphNode[]): number {
  return paragraphs.reduce(
    (count, paragraph) => count + paragraph.content.filter((child) => child.kind === 'image').length,
    0,
  );
}

function countNativeDrawings(paragraphs: NativeWordParagraphNode[]): number {
  return paragraphs.reduce(
    (count, paragraph) => count + paragraph.content.filter((child) => child.kind === 'drawing').length,
    0,
  );
}

function countPageBreaks(paragraphs: EditorParagraphView[]): number {
  return paragraphs.reduce(
    (count, paragraph) =>
      count +
      paragraph.segments.filter(
        (segment) => segment.kind === 'placeholder' && segment.placeholderKind === 'pageBreak',
      ).length,
    0,
  );
}

function countNativePageBreaks(paragraphs: NativeWordParagraphNode[]): number {
  return paragraphs.reduce(
    (count, paragraph) => count + paragraph.content.filter((child) => child.kind === 'pageBreak').length,
    0,
  );
}
