# Context Search APIs

This domain defines the Core Word API for local document context search. It is for locating actionable context inside the current Word document before an agent reads, comments, edits, or formats content.

Context search is not external retrieval, a knowledge base, or host-specific search. It runs on the loaded `WordDocument` model and returns Core-owned references such as paragraph ids, table coordinates, bookmark metadata, and text snippets.

## Skill Summary

Use this document when the task says: find context, search keyword, locate clause, find by person/name/date/amount, inspect a section, find bookmarked content, find appended comparison text, or operate on the part about a topic.

The public entry point is:

```ts
const result = await doc.context.search({
  query: '违约金',
  keywords: ['违约金', '违约责任'],
  context: { before: 2, after: 2 },
  limit: 10,
});
```

The API returns matches with enough local context for an agent to decide the next operation without reading the whole document.

## Entry Point

```ts
doc.context.search(options): Promise<ContextSearchResult>
```

`ContextSearchOptions`:

```ts
interface ContextSearchOptions {
  query: string;
  keywords?: string[];
  scope?: 'body' | 'all';
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

interface ContextSearchInclude {
  paragraphs?: boolean;
  tables?: boolean;
  headings?: boolean;
  bookmarks?: 'none' | 'metadata' | 'coveredText';
  comments?: 'none' | 'anchors';
  revisions?: 'none' | 'flag' | 'visibleText';
  headersFooters?: boolean;
  footnotesEndnotes?: boolean;
}
```

Recommended defaults:

```ts
{
  scope: 'body',
  include: {
    paragraphs: true,
    tables: true,
    headings: true,
    bookmarks: 'metadata',
    comments: 'anchors',
    revisions: 'flag',
  },
  context: { before: 2, after: 2 },
  limit: 10,
}
```

## Searchable Content

### v1 Required

| Content         | Default             | Notes                                                                  |
| --------------- | ------------------- | ---------------------------------------------------------------------- |
| Body paragraphs | on                  | Main path. Returns `paraId`.                                           |
| Table cell text | on                  | Returns `tableId`, `row`, `col`, and containing `paraId`.              |
| Headings        | on                  | Heading text can match; paragraph matches may include `headingPath`.   |
| Comment anchors | marker only         | Return `hasComments`; v1 does not need to search comment body text.    |
| Revisions       | marker only         | Return `hasRevisions`; v1 does not need to expose every revision text. |
| Bookmarks       | metadata by default | Reliable content domain. See bookmark rules below.                     |

### Bookmarks

Bookmarks are reliable enough for context search because Core already exposes live objects:

- `doc.bookmarks.list(opts?)`
- `doc.bookmarks.get(name)`
- `bookmark.getText()`
- `bookmark.getParagraph()`

Bookmark search has two modes:

- `bookmarks:'metadata'`: attach bookmark metadata to paragraph/table matches.
- `bookmarks:'coveredText'`: also search bookmark-covered text as its own match source.

`bookmarkFilter.prefix` supports agent-owned trace bookmarks such as `_co_append`. Hidden bookmarks are excluded unless `includeHidden:true` or a hidden prefix is explicitly requested.

Range reliability:

| `rangeKind` | Read behavior                                      | Write follow-up                                             |
| ----------- | -------------------------------------------------- | ----------------------------------------------------------- |
| `paragraph` | Search and return whole paragraph text.            | Writable through `paraId`.                                  |
| `text`      | Search and return covered text.                    | Writable through containing `paraId` plus exact text range. |
| `complex`   | Return as read-only context when text can be read. | Do not write directly; re-locate a concrete range first.    |

### Later Extensions

These are useful but should not block v1:

- Headers and footers.
- Footnotes and endnotes.
- Comment body text.
- Revision visible text versus deleted text policies.
- Text boxes and shape text.
- Field results such as TOC entries, page numbers, and cross references.
- OCR text from images.

## Result Shape

Planned result:

```ts
interface ContextSearchResult {
  matches: ContextSearchMatch[];
  truncated?: boolean;
}

interface ContextSearchMatch {
  matchId: string;
  score: number;
  reason: 'keyword' | 'bookmark' | 'heading' | 'table' | 'metadata';
  target: ContextSearchTarget;
  text: string;
  highlights?: Array<{ text: string; start: number; end: number }>;
  before?: ContextBlock[];
  after?: ContextBlock[];
  headingPath?: string[];
  flags?: {
    hasComments?: boolean;
    hasRevisions?: boolean;
  };
  bookmarks?: ContextBookmark[];
}

type ContextSearchTarget =
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

interface ContextBlock {
  paraId: string;
  story: string;
  text: string;
  table?: { tableId: string; row: number; col: number };
}

interface ContextBookmark {
  name: string;
  rangeKind: 'paragraph' | 'text' | 'complex';
  writable: boolean;
}
```

## CodeAgent Rules

1. Prefer `doc.context.search()` before whole-document reads when the user gives keywords, entities, clauses, headings, or bookmark names.
2. Treat search results as current-document context only; do not merge external knowledge into this API.
3. Follow-up writes must use returned `paraId`, table coordinates, exact text, or a writable bookmark-derived range.
4. Do not write directly to `rangeKind:'complex'` bookmark matches.
5. If several matches are plausible and the write is risky, ask the user which target to operate on.
6. Use `doc.paragraphs.list()` for full-document processing, explicit pagination, or when context search cannot locate the target.
