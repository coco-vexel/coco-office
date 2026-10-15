import { describe, expect, test } from 'vitest';
import { searchNativeWordDocument } from '../native-search.js';
import type {
  NativeWordBlockNode,
  NativeWordDocument,
  NativeWordParagraphNode,
} from '../native-model.js';

function paragraph(paraId: string, text: string): NativeWordParagraphNode {
  return {
    kind: 'paragraph',
    paraId,
    story: 'body',
    styleId: null,
    format: null,
    content: [{ kind: 'text', text, format: null }],
  };
}

function cell(paraId: string, text: string) {
  const block = paragraph(paraId, text);
  return { gridSpan: 1, blocks: [block], paragraphs: [block] };
}

function fixture(): NativeWordDocument {
  const nestedTable: NativeWordBlockNode = {
    kind: 'table',
    tableId: 'nested',
    columnWidths: [1200],
    rows: [{ cells: [cell('table-p1', 'Alpha 中文 Alpha')] }],
  };
  return {
    kind: 'word',
    sections: [],
    headers: [{ kind: 'header', partName: 'header1.xml', blocks: [paragraph('header-p1', 'Alpha')] }],
    footers: [],
    blocks: [paragraph('body-p1', 'alpha alphabet ALPHA'), nestedTable],
  };
}

describe('native Word search', () => {
  test('searches body and nested table paragraphs while excluding headers', () => {
    const matches = searchNativeWordDocument(fixture(), { query: 'alpha' });

    expect(matches.map((match) => [match.paraId, match.from, match.text])).toEqual([
      ['body-p1', 0, 'alpha'],
      ['body-p1', 6, 'alpha'],
      ['body-p1', 15, 'ALPHA'],
      ['table-p1', 0, 'Alpha'],
      ['table-p1', 9, 'Alpha'],
    ]);
  });

  test('supports case-sensitive and Unicode-aware whole-word matching', () => {
    expect(searchNativeWordDocument(fixture(), { query: 'Alpha', matchCase: true })).toHaveLength(2);
    expect(searchNativeWordDocument(fixture(), { query: 'alpha', wholeWord: true })).toHaveLength(4);
    expect(searchNativeWordDocument(fixture(), { query: '中', wholeWord: true })).toHaveLength(0);
    expect(searchNativeWordDocument(fixture(), { query: '中文', wholeWord: true })).toHaveLength(1);
    expect(searchNativeWordDocument(fixture(), { query: '' })).toEqual([]);
  });
});
