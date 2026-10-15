import type { Package } from '../pack/index.js';
import type { BookmarkInfo, ListBookmarksOptions } from './bookmark-queries.js';
import { getBookmark, listBookmarks } from './bookmark-queries.js';
import type { PackageIndices } from './indices.js';
import type { ParagraphInfo } from './list-paragraphs.js';
import { listParagraphs } from './list-paragraphs.js';

export type ContextSearchScope = 'body' | 'all';

export interface ContextSearchInclude {
  paragraphs?: boolean;
  tables?: boolean;
  headings?: boolean;
  bookmarks?: 'none' | 'metadata' | 'coveredText';
  comments?: 'none' | 'anchors';
  revisions?: 'none' | 'flag' | 'visibleText';
  headersFooters?: boolean;
  footnotesEndnotes?: boolean;
}

export interface ContextSearchOptions {
  query: string;
  keywords?: string[];
  scope?: ContextSearchScope;
  include?: ContextSearchInclude;
  context?: {
    before?: number;
    after?: number;
  };
  limit?: number;
  bookmarkFilter?: {
    name?: string;
    prefix?: string;
    includeHidden?: boolean;
  };
}

export interface ContextSearchResult {
  matches: ContextSearchMatch[];
  truncated?: boolean;
}

export interface ContextSearchMatch {
  matchId: string;
  score: number;
  reason: 'keyword' | 'bookmark' | 'heading' | 'table' | 'metadata';
  target: ContextSearchTarget;
  text: string;
  highlights?: ContextSearchHighlight[];
  before?: ContextBlock[];
  after?: ContextBlock[];
  headingPath?: string[];
  flags?: {
    hasComments?: boolean;
    hasRevisions?: boolean;
  };
  bookmarks?: ContextBookmark[];
}

export interface ContextSearchHighlight {
  text: string;
  start: number;
  end: number;
}

export type ContextSearchTarget =
  | { kind: 'paragraph'; paraId: string; story: string }
  | { kind: 'tableCell'; tableId: string; row: number; col: number; paraId?: string; story: string }
  | {
      kind: 'bookmark';
      bookmarkName: string;
      rangeKind: 'paragraph' | 'text' | 'complex';
      paraId?: string;
      story: string;
    }
  | { kind: 'heading'; paraId: string; story: string };

export interface ContextBlock {
  paraId: string;
  story: string;
  text: string;
  table?: { tableId: string; row: number; col: number };
}

export interface ContextBookmark {
  name: string;
  rangeKind: 'paragraph' | 'text' | 'complex';
  writable: boolean;
}

interface ParagraphEntry {
  paragraph: ParagraphInfo;
  index: number;
  headingPath: string[];
}

const DEFAULT_LIMIT = 10;
const DEFAULT_CONTEXT_SIZE = 2;

export function searchContext(
  pkg: Package,
  indices: PackageIndices,
  options: ContextSearchOptions,
): ContextSearchResult {
  const include = normalizeInclude(options.include);
  const terms = normalizeTerms(options.query, options.keywords);
  const limit = normalizeLimit(options.limit);
  const beforeCount = normalizeContextCount(options.context?.before);
  const afterCount = normalizeContextCount(options.context?.after);
  const scope = options.scope ?? 'body';

  const paragraphEntries = buildParagraphEntries(pkg, scope);
  const bookmarkInfos = collectBookmarks(pkg, indices, options.bookmarkFilter, include.bookmarks);
  const bookmarksByParaId = buildBookmarksByParaId(bookmarkInfos);

  const matches: ContextSearchMatch[] = [];

  if (terms.length > 0) {
    for (const entry of paragraphEntries) {
      const paragraph = entry.paragraph;
      const isTable = paragraph.table !== undefined;
      const isHeadingMatch = isHeading(paragraph);

      if (isTable && include.tables !== true) continue;
      if (!isTable && include.paragraphs !== true && !isHeadingMatch) continue;
      if (isHeadingMatch && include.headings !== true && include.paragraphs !== true) continue;

      const highlights = findHighlights(paragraph.text, terms);
      if (highlights.length === 0) continue;

      const reason = isHeadingMatch ? 'heading' : isTable ? 'table' : 'keyword';
      matches.push(
        buildParagraphMatch(
          matches.length + 1,
          entry,
          paragraphEntries,
          beforeCount,
          afterCount,
          reason,
          highlights,
          bookmarksByParaId.get(paragraph.paraId),
          include,
        ),
      );
      if (matches.length >= limit) return { matches, truncated: true };
    }
  }

  if (include.bookmarks === 'coveredText') {
    for (const bookmark of bookmarkInfos) {
      const text = resolveBookmarkText(bookmark, paragraphEntries);
      if (text.length === 0) continue;
      const highlights = terms.length > 0 ? findHighlights(text, terms) : [];
      if (terms.length > 0 && highlights.length === 0) continue;
      const entry = bookmark.paraId
        ? paragraphEntries.find((candidate) => candidate.paragraph.paraId === bookmark.paraId)
        : undefined;
      matches.push(
        buildBookmarkMatch(
          matches.length + 1,
          bookmark,
          text,
          highlights,
          entry,
          paragraphEntries,
          beforeCount,
          afterCount,
        ),
      );
      if (matches.length >= limit) return { matches, truncated: true };
    }
  }

  return { matches };
}

function normalizeInclude(
  include: ContextSearchInclude | undefined,
): Required<ContextSearchInclude> {
  return {
    paragraphs: include?.paragraphs ?? true,
    tables: include?.tables ?? true,
    headings: include?.headings ?? true,
    bookmarks: include?.bookmarks ?? 'metadata',
    comments: include?.comments ?? 'anchors',
    revisions: include?.revisions ?? 'flag',
    headersFooters: include?.headersFooters ?? false,
    footnotesEndnotes: include?.footnotesEndnotes ?? false,
  };
}

function normalizeTerms(query: string, keywords: string[] | undefined): string[] {
  const out: string[] = [];
  for (const raw of [query, ...(keywords ?? [])]) {
    const term = raw.trim();
    if (term.length > 0 && !out.includes(term)) out.push(term);
  }
  return out;
}

function normalizeLimit(limit: number | undefined): number {
  if (limit === undefined) return DEFAULT_LIMIT;
  if (!Number.isFinite(limit) || limit < 1) return DEFAULT_LIMIT;
  return Math.floor(limit);
}

function normalizeContextCount(value: number | undefined): number {
  if (value === undefined) return DEFAULT_CONTEXT_SIZE;
  if (!Number.isFinite(value) || value < 0) return DEFAULT_CONTEXT_SIZE;
  return Math.floor(value);
}

function buildParagraphEntries(pkg: Package, scope: ContextSearchScope): ParagraphEntry[] {
  const result = listParagraphs(pkg, {
    scope,
    pagination: { mode: 'all' },
    includeFormat: true,
    includeTableInfo: true,
  });
  const headingStack: string[] = [];
  return result.paragraphs.map((paragraph, index) => {
    if (paragraph.table === undefined && isHeading(paragraph)) {
      const level = headingLevel(paragraph);
      if (level !== undefined) {
        headingStack.length = level;
        headingStack[level] = paragraph.text;
      } else {
        headingStack.length = 0;
        headingStack.push(paragraph.text);
      }
    }
    return {
      paragraph,
      index,
      headingPath: headingStack.filter((value) => value.length > 0),
    };
  });
}

function collectBookmarks(
  pkg: Package,
  indices: PackageIndices,
  filter: ContextSearchOptions['bookmarkFilter'],
  includeMode: Required<ContextSearchInclude>['bookmarks'],
): BookmarkInfo[] {
  if (includeMode === 'none') return [];
  if (filter?.name) {
    const bookmark = getBookmark(pkg, indices, filter.name);
    return bookmark ? [bookmark] : [];
  }
  const opts: ListBookmarksOptions = {
    ...(filter?.prefix !== undefined ? { prefix: filter.prefix } : {}),
    includeHidden:
      filter?.includeHidden === true ||
      (filter?.prefix !== undefined && filter.prefix.startsWith('_')),
  };
  return listBookmarks(pkg, indices, opts);
}

function buildBookmarksByParaId(bookmarks: BookmarkInfo[]): Map<string, ContextBookmark[]> {
  const out = new Map<string, ContextBookmark[]>();
  for (const bookmark of bookmarks) {
    if (!bookmark.paraId) continue;
    const existing = out.get(bookmark.paraId) ?? [];
    existing.push({
      name: bookmark.name,
      rangeKind: bookmark.rangeKind,
      writable: bookmark.rangeKind === 'paragraph' || bookmark.rangeKind === 'text',
    });
    out.set(bookmark.paraId, existing);
  }
  return out;
}

function isHeading(paragraph: ParagraphInfo): boolean {
  return headingLevel(paragraph) !== undefined;
}

function headingLevel(paragraph: ParagraphInfo): number | undefined {
  const outline = paragraph.format?.outlineLevel;
  if (outline !== undefined && outline >= 0 && outline <= 8) return outline;
  const style = paragraph.styleId?.toLowerCase();
  if (!style) return undefined;
  const headingMatch = style.match(/^heading([1-9])$/);
  if (headingMatch) return Number(headingMatch[1]) - 1;
  const chineseMatch = style.match(/^标题([1-9])$/);
  if (chineseMatch) return Number(chineseMatch[1]) - 1;
  return undefined;
}

function findHighlights(text: string, terms: string[]): ContextSearchHighlight[] {
  const highlights: ContextSearchHighlight[] = [];
  for (const term of terms) {
    const start = text.indexOf(term);
    if (start >= 0) highlights.push({ text: term, start, end: start + term.length });
  }
  return highlights;
}

function buildParagraphMatch(
  ordinal: number,
  entry: ParagraphEntry,
  entries: ParagraphEntry[],
  beforeCount: number,
  afterCount: number,
  reason: 'keyword' | 'heading' | 'table',
  highlights: ContextSearchHighlight[],
  bookmarks: ContextBookmark[] | undefined,
  include: Required<ContextSearchInclude>,
): ContextSearchMatch {
  const paragraph = entry.paragraph;
  const target: ContextSearchTarget = paragraph.table
    ? {
        kind: 'tableCell',
        tableId: paragraph.table.tableId,
        row: paragraph.table.row,
        col: paragraph.table.col,
        paraId: paragraph.paraId,
        story: paragraph.story,
      }
    : reason === 'heading'
      ? { kind: 'heading', paraId: paragraph.paraId, story: paragraph.story }
      : { kind: 'paragraph', paraId: paragraph.paraId, story: paragraph.story };

  return cleanMatch({
    matchId: `ctx-${ordinal}`,
    score: scoreFor(reason, highlights.length),
    reason,
    target,
    text: paragraph.text,
    highlights,
    before: contextBlocks(entries, entry.index, -beforeCount, -1),
    after: contextBlocks(entries, entry.index, 1, afterCount),
    headingPath: entry.headingPath,
    flags: {
      ...(include.comments !== 'none' && paragraph.hasComments ? { hasComments: true } : {}),
      ...(include.revisions !== 'none' && paragraph.hasRevisions ? { hasRevisions: true } : {}),
    },
    bookmarks,
  });
}

function buildBookmarkMatch(
  ordinal: number,
  bookmark: BookmarkInfo,
  text: string,
  highlights: ContextSearchHighlight[],
  entry: ParagraphEntry | undefined,
  entries: ParagraphEntry[],
  beforeCount: number,
  afterCount: number,
): ContextSearchMatch {
  return cleanMatch({
    matchId: `ctx-${ordinal}`,
    score: highlights.length > 0 ? 0.9 : 0.72,
    reason: 'bookmark',
    target: {
      kind: 'bookmark',
      bookmarkName: bookmark.name,
      rangeKind: bookmark.rangeKind,
      ...(bookmark.paraId !== undefined ? { paraId: bookmark.paraId } : {}),
      story: bookmark.story,
    },
    text,
    highlights,
    before: entry ? contextBlocks(entries, entry.index, -beforeCount, -1) : undefined,
    after: entry ? contextBlocks(entries, entry.index, 1, afterCount) : undefined,
    headingPath: entry?.headingPath,
    bookmarks: [
      {
        name: bookmark.name,
        rangeKind: bookmark.rangeKind,
        writable: bookmark.rangeKind === 'paragraph' || bookmark.rangeKind === 'text',
      },
    ],
  });
}

function scoreFor(reason: 'keyword' | 'heading' | 'table', hitCount: number): number {
  const base = reason === 'heading' ? 0.86 : reason === 'table' ? 0.84 : 0.82;
  return Math.min(0.99, base + Math.max(0, hitCount - 1) * 0.03);
}

function contextBlocks(
  entries: ParagraphEntry[],
  index: number,
  fromOffset: number,
  toOffset: number,
): ContextBlock[] | undefined {
  const start = Math.max(0, index + fromOffset);
  const end = Math.min(entries.length - 1, index + toOffset);
  if (start > end) return undefined;
  const blocks: ContextBlock[] = [];
  for (let i = start; i <= end; i++) {
    if (i === index) continue;
    const paragraph = entries[i]!.paragraph;
    if (paragraph.text.length === 0) continue;
    blocks.push({
      paraId: paragraph.paraId,
      story: paragraph.story,
      text: paragraph.text,
      ...(paragraph.table ? { table: paragraph.table } : {}),
    });
  }
  return blocks.length > 0 ? blocks : undefined;
}

function resolveBookmarkText(bookmark: BookmarkInfo, entries: ParagraphEntry[]): string {
  if (bookmark.rangeKind === 'text') return bookmark.text ?? '';
  if (!bookmark.paraId) return '';
  return entries.find((entry) => entry.paragraph.paraId === bookmark.paraId)?.paragraph.text ?? '';
}

function cleanMatch(match: ContextSearchMatch): ContextSearchMatch {
  if (match.highlights?.length === 0) delete match.highlights;
  if (match.before?.length === 0) delete match.before;
  if (match.after?.length === 0) delete match.after;
  if (match.headingPath?.length === 0) delete match.headingPath;
  if (match.bookmarks?.length === 0) delete match.bookmarks;
  if (match.flags && Object.keys(match.flags).length === 0) delete match.flags;
  return match;
}
