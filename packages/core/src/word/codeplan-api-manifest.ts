import type { SnapshotApi } from './snapshot-api.js';
import type {
  Bookmark,
  BookmarkApi,
  Comment,
  CommentApi,
  ContextApi,
  Drawing,
  DrawingsApi,
  Paragraph,
  ParagraphApi,
  PictureDrawing,
  Revision,
  RevisionApi,
  SectionsApi,
  ShapeDrawing,
  StylesApi,
  Table,
  TableCell,
  TableColumn,
  TableRow,
  TablesApi,
  TextBoxDrawing,
  TextRange,
  TocFieldsApi,
} from './sub-apis.js';
import type { WordDocument } from './word-document.js';

export type WordCodePlanApiDomain =
  | 'document'
  | 'paragraph'
  | 'bookmark'
  | 'table'
  | 'drawing'
  | 'range'
  | 'comment'
  | 'revision'
  | 'style'
  | 'section'
  | 'context'
  | 'toc'
  | 'snapshot';

export type WordCodePlanApiKind = 'locator' | 'read' | 'write' | 'format' | 'lifecycle';

type WordCodePlanReceiverMap = {
  doc: WordDocument;
  'doc.paragraphs': ParagraphApi;
  'doc.bookmarks': BookmarkApi;
  'doc.comments': CommentApi;
  'doc.context': ContextApi;
  'doc.revisions': RevisionApi;
  'doc.styles': StylesApi;
  'doc.tables': TablesApi;
  'doc.drawings': DrawingsApi;
  'doc.sections': SectionsApi;
  'doc.tocFields': TocFieldsApi;
  'doc.snapshots': SnapshotApi;
  paragraph: Paragraph;
  range: TextRange;
  bookmark: Bookmark;
  comment: Comment;
  revision: Revision;
  table: Table;
  cell: TableCell;
  row: TableRow;
  column: TableColumn;
  drawing: Drawing;
  picture: PictureDrawing;
  textBox: TextBoxDrawing;
  shape: ShapeDrawing;
};

type ReceiverName = keyof WordCodePlanReceiverMap;
type MethodName<T> = {
  [K in keyof T]-?: T[K] extends (...args: never[]) => unknown ? K : never;
}[keyof T] &
  string;

export type WordCodePlanSdkApiPath = {
  [R in ReceiverName]: `${R}.${MethodName<WordCodePlanReceiverMap[R]>}`;
}[ReceiverName];

export interface WordCodePlanApiManifestEntry {
  path: WordCodePlanSdkApiPath;
  domain: WordCodePlanApiDomain;
  kind: WordCodePlanApiKind;
  codeplan: boolean;
  requiresAwait: boolean;
  reason?: string;
}

interface WordCodePlanApiGroup<R extends ReceiverName = ReceiverName> {
  receiver: R;
  domain: WordCodePlanApiDomain;
  kind: WordCodePlanApiKind;
  methods: readonly MethodName<WordCodePlanReceiverMap[R]>[];
  codeplan: boolean;
  requiresAwait: boolean;
  reason?: string;
}

function group<R extends ReceiverName>(
  receiver: R,
  domain: WordCodePlanApiDomain,
  kind: WordCodePlanApiKind,
  methods: readonly MethodName<WordCodePlanReceiverMap[R]>[],
  options: { codeplan?: boolean; requiresAwait?: boolean; reason?: string } = {},
): WordCodePlanApiGroup<R> {
  return {
    receiver,
    domain,
    kind,
    methods,
    codeplan: options.codeplan ?? true,
    requiresAwait: options.requiresAwait ?? true,
    ...(options.reason ? { reason: options.reason } : {}),
  };
}

const groups = [
  group('doc.paragraphs', 'paragraph', 'locator', ['list', 'get']),
  group('paragraph', 'paragraph', 'read', ['getInfo', 'getText']),
  group('paragraph', 'paragraph', 'write', [
    'replaceWith',
    'setAlignment',
    'setOutlineLevel',
    'setPageBreakBefore',
    'setIndent',
    'setSpacing',
    'insertBefore',
    'insertAfter',
    'insertImageAfter',
    'insertTableAfter',
    'delete',
    'addComment',
    'createBookmark',
  ]),
  group('paragraph', 'range', 'locator', ['findText'], { requiresAwait: false }),
  group('range', 'range', 'write', ['replaceWith', 'delete', 'addComment', 'createBookmark']),
  group('range', 'range', 'format', [
    'setBold',
    'setColor',
    'setItalic',
    'setUnderline',
    'setHighlight',
    'setFontSize',
    'setFontFamily',
  ]),
  group('doc.comments', 'comment', 'locator', ['list', 'get']),
  group('comment', 'comment', 'read', ['getInfo', 'getText', 'getAnchorText', 'getParagraph']),
  group('comment', 'comment', 'write', ['delete']),
  group('doc.revisions', 'revision', 'locator', ['list', 'get']),
  group('doc.revisions', 'revision', 'write', ['accept', 'reject']),
  group('revision', 'revision', 'read', ['getInfo', 'getText', 'getParagraph']),
  group('revision', 'revision', 'write', ['accept', 'reject']),
  group('doc.bookmarks', 'bookmark', 'locator', ['list', 'get']),
  group('bookmark', 'bookmark', 'read', ['getInfo', 'getText', 'getParagraph']),
  group('bookmark', 'bookmark', 'write', ['delete']),
  group('doc.tables', 'table', 'locator', ['list', 'get', 'getVerbatim']),
  group('table', 'table', 'read', ['getInfo', 'getText', 'getCells', 'getVerbatim']),
  group('table', 'table', 'locator', ['cell', 'row', 'column'], { requiresAwait: false }),
  group('table', 'table', 'write', [
    'addComment',
    'delete',
    'replaceCellText',
    'deleteRow',
    'insertRow',
    'setColumnWidths',
    'setRowHeights',
  ]),
  group('cell', 'table', 'read', ['getInfo', 'getBlocks', 'getParagraphs', 'getText']),
  group('cell', 'range', 'locator', ['findText'], { requiresAwait: false }),
  group('cell', 'table', 'write', [
    'replaceWith',
    'insertParagraphAtStart',
    'insertParagraphAtEnd',
    'setShading',
    'setVerticalAlignment',
    'setWidth',
  ]),
  group('row', 'table', 'read', ['getCells', 'getText']),
  group('row', 'table', 'write', ['delete', 'insertBefore', 'insertAfter', 'setHeight']),
  group('column', 'table', 'read', ['getCells', 'getText']),
  group('column', 'table', 'write', ['setWidth']),
  group('doc.drawings', 'drawing', 'locator', [
    'list',
    'listPictures',
    'listTextBoxes',
    'listShapes',
    'get',
  ]),
  group('drawing', 'drawing', 'read', ['getInfo', 'getImageBytes']),
  group('drawing', 'drawing', 'write', [
    'delete',
    'replaceWithPlaceholder',
    'replaceShapeInnerText',
  ]),
  group('picture', 'drawing', 'read', ['getImageBytes']),
  group('picture', 'drawing', 'write', ['delete', 'replaceWithPlaceholder']),
  group('textBox', 'drawing', 'read', ['getText', 'getParagraphs', 'getBlocks']),
  group('textBox', 'drawing', 'write', ['replaceText']),
  group('shape', 'drawing', 'read', ['getInfo']),
  group('shape', 'drawing', 'write', ['delete']),
  group('doc.styles', 'style', 'read', ['list', 'get']),
  group('doc.context', 'context', 'read', ['search']),
  group('doc.sections', 'section', 'read', ['list']),
  group('doc.tocFields', 'toc', 'read', ['list']),
  group('doc', 'document', 'write', ['appendParagraph']),
  group(
    'doc.snapshots',
    'snapshot',
    'lifecycle',
    ['take', 'restore', 'rollback', 'list', 'delete', 'clear', 'recordChange', 'changes'],
    {
      codeplan: false,
      reason:
        'Snapshot lifecycle is owned by the host runtime, not generated document-operation code.',
    },
  ),
  group(
    'doc',
    'document',
    'lifecycle',
    ['toZip', 'toFlatOpc', 'dispose', 'executeCode', 'plan', 'on'],
    {
      codeplan: false,
      reason:
        'The host owns export, disposal, subscriptions, nested execution, and low-level plan orchestration.',
    },
  ),
] as const;

export const WORD_CODEPLAN_API_MANIFEST_VERSION = 1;

export const WORD_CODEPLAN_API_MANIFEST: readonly WordCodePlanApiManifestEntry[] = groups.flatMap(
  ({ receiver, domain, kind, methods, codeplan, requiresAwait, reason }) =>
    methods.map((method) => ({
      path: `${receiver}.${method}` as WordCodePlanSdkApiPath,
      domain,
      kind,
      codeplan,
      requiresAwait,
      ...(reason ? { reason } : {}),
    })),
);

export const WORD_CODEPLAN_ALLOWED_API_PATHS = WORD_CODEPLAN_API_MANIFEST.filter(
  (entry) => entry.codeplan,
).map((entry) => entry.path);
