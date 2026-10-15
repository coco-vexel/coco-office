import {
  nativeParagraphText,
  type NativeWordBlockNode,
  type NativeWordDocument,
  type NativeWordParagraphNode,
} from './native-model.js';
import type { WordEditorSearchMatch, WordEditorSearchOptions } from './runtime.js';

export function searchNativeWordDocument(
  doc: NativeWordDocument,
  options: WordEditorSearchOptions,
): WordEditorSearchMatch[] {
  if (options.query.length === 0) return [];
  const flags = options.matchCase ? 'gu' : 'giu';
  const pattern = new RegExp(escapeRegExp(options.query), flags);
  const matches: WordEditorSearchMatch[] = [];
  for (const paragraph of bodyParagraphs(doc.blocks)) {
    const text = nativeParagraphText(paragraph);
    for (const match of text.matchAll(pattern)) {
      const from = match.index;
      const to = from + match[0].length;
      if (options.wholeWord && !isWholeWord(text, from, to)) continue;
      matches.push({
        id: `${paragraph.paraId}:${from}:${to}`,
        paraId: paragraph.paraId,
        from,
        to,
        text: match[0],
      });
    }
  }
  return matches;
}

export function bodyParagraphs(blocks: readonly NativeWordBlockNode[]): NativeWordParagraphNode[] {
  return blocks.flatMap((block) => {
    if (block.kind === 'paragraph') return [block];
    return block.rows.flatMap((row) => row.cells.flatMap((cell) => bodyParagraphs(cell.blocks)));
  });
}

function isWholeWord(text: string, from: number, to: number): boolean {
  const before = from > 0 ? text[from - 1]! : '';
  const after = to < text.length ? text[to]! : '';
  return !isWordCharacter(before) && !isWordCharacter(after);
}

function isWordCharacter(value: string): boolean {
  return value.length > 0 && /[\p{L}\p{N}_]/u.test(value);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
