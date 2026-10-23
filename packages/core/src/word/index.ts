export {
  W_NS,
  R_NS,
  A_NS,
  WP_NS,
  PIC_NS,
  CHART_NS,
  DGM_NS,
  WPS_NS,
  WPG_NS,
  WPC_NS,
  MC_NS,
  VML_NS,
  W14_NS,
  PKG_NS,
  XML_NS,
} from './namespaces.js';
export { ensureParaIds, type EnsureParaIdsResult } from './para-id.js';
export {
  getParagraphRuns,
  getRunText,
  getParagraphText,
  splitRunAt,
  type SplitResult,
  type RunTextOptions,
} from './runs.js';
export {
  buildIndices,
  BookmarkLoadError,
  type BookmarkLoadErrorCode,
  type BookmarkEntry,
  type CommentEntry,
  type PackageIndices,
} from './indices.js';
export {
  listStories,
  getStoryParagraphs,
  findStoryFor,
  resolveScope,
  type Story,
  type StoryId,
  type StoryKind,
  type ListStoriesOptions,
} from './stories.js';
export {
  listParagraphs,
  getParagraphByParaId,
  buildParagraphInfo,
  DEFAULT_PAGINATION,
  PREVIEW_CHAR_LIMIT,
  type ListParagraphsOptions,
  type ListParagraphsScope,
  type ListParagraphsTextFormat,
  type ListParagraphsFilter,
  type ListParagraphsPagination,
  type ParagraphInfo,
  type ParagraphFormat,
  type ListResult,
} from './list-paragraphs.js';
export {
  listBookmarks,
  getBookmark,
  type BookmarkInfo,
  type BookmarkName,
  type ListBookmarksOptions,
} from './bookmark-queries.js';
export { listComments, getComment, type CommentId, type CommentInfo } from './comment-queries.js';
export {
  searchContext,
  type ContextSearchOptions,
  type ContextSearchInclude,
  type ContextSearchResult,
  type ContextSearchMatch,
  type ContextSearchHighlight,
  type ContextSearchTarget,
  type ContextBlock,
  type ContextBookmark,
} from './context-search.js';
export {
  listRevisions,
  type RevisionInfo,
  type RevisionKind,
  type SupportedRevisionLocalName,
  type SupportedW14RevisionLocalName,
  type SupportedRevisionType,
  type ListRevisionsOptions,
} from './revision-queries.js';
export {
  listSections,
  type SectionInfo,
  type SectionType,
  type HeaderFooterRef,
  type HeaderFooterType,
} from './section-queries.js';
export { listTocFields, type TocFieldRange } from './toc-queries.js';
export {
  WordPlanImpl,
  EMPTY_APPLY_RESULT,
  type WordPlan,
  type WordPlanImplInit,
  type PlanOptions,
  type ApplyResult,
  type ApplyPlanParams,
  type PlanValidationIssue,
  type PlanValidationItem,
  type PlanValidationResult,
  type Op,
  type Range,
  type ParagraphRange,
  type TextRange as TextRangeDescriptor,
  type RevisionScope,
  type InsertParagraphSpec,
  type ParaId,
  type RevisionId,
} from './plan.js';
export { classifyRange, type RangeClassification } from './range-kind.js';
export { WordError, type WordErrorCode } from './word-error.js';
export {
  WordDocument,
  type LoadOptions,
  type LoadPhase,
  type RuntimeMode,
  type ModeDecision,
  type ModeSource,
  type WorkerFactory,
} from './word-document.js';
export {
  WORD_CHANGE_EVENT,
  type ChangeEvent,
  type ChangeListener,
  type ChangeSubscribeOptions,
  type KernelChangeEvent,
} from './change-events.js';
export {
  type ParagraphApi,
  type ParagraphList,
  type Paragraph,
  type LiveObjectWriteOptions,
  type ParagraphDeleteOptions,
  type ParagraphInsertOptions,
  type ParagraphInsertImageSpec,
  type ParagraphInsertTableSpec,
  type ParagraphReplaceOptions,
  type TextRange,
  type BookmarkApi,
  type Bookmark,
  type CommentApi,
  type Comment,
  type ContextApi,
  type RevisionApi,
  type Revision,
  type StylesApi,
  type TablesApi,
  type Table,
  type TableCell,
  type TableCellBlock,
  type CellTextRange,
  type TableRow,
  type TableColumn,
  type Drawing,
  type PictureDrawing,
  type TextBoxDrawing,
  type ShapeDrawing,
  type DrawingImageBytes,
  type DrawingsApi,
  type SectionsApi,
  type TocFieldsApi,
} from './sub-apis.js';
export {
  listTables,
  getTableById,
  getTableVerbatim,
  type TableInfo,
  type TableCellInfo,
  type TableCellBlockInfo,
  type RowHeight,
  type ListTablesOptions,
  type TableVerbatim,
  type TableVerbatimImage,
} from './list-tables.js';
export {
  listDrawings,
  getDrawingById,
  buildDrawingIndex,
  hasDrawingAncestor,
  type DrawingInfo,
  type ListDrawingsOptions,
  type DrawingAnchoringMode,
  type DrawingKind,
  type DrawingWrapType,
  type DrawingPosition,
  type DrawingPositionAxis,
} from './list-drawings.js';
export {
  buildInlineSdt,
  hashSdtId,
  type SdtSpec,
  type PlainTextSdtSpec,
  type DateSdtSpec,
  type RichTextSdtSpec,
} from './sdt-builders.js';
export {
  buildStylesIndex,
  listStyles,
  getStyleById,
  type StyleInfo,
  type StyleType,
  type ListStylesOptions,
} from './style-queries.js';
export { runTransaction } from './transactions.js';
export {
  applyXQueryUpdate,
  type XQueryUpdateOptions,
  type XQueryUpdateResult,
} from './xquery-update.js';
export {
  Kernel,
  type KernelLoadParams,
  type KernelLoadSource,
  type SnapshotPartEntry,
} from './kernel.js';
export {
  DEFAULT_MAX_SNAPSHOTS,
  type SnapshotApi,
  type RollbackOptions,
  type RollbackProgress,
} from './snapshot-api.js';
export type {
  RpcMessage,
  RpcRequest,
  RpcResponse,
  RpcError,
  RpcProgress,
  RpcEvent,
  SerializableError,
} from './rpc-protocol.js';
export { InlineTransport, type Transport } from './transport.js';
export { WorkerTransport, type WorkerPortLike } from './worker-transport.js';
export { createBrowserWorkerPort } from './worker-port-browser.js';
export {
  WORD_CODEPLAN_API_MANIFEST,
  WORD_CODEPLAN_API_MANIFEST_VERSION,
  WORD_CODEPLAN_ALLOWED_API_PATHS,
  type WordCodePlanApiDomain,
  type WordCodePlanApiKind,
  type WordCodePlanApiManifestEntry,
  type WordCodePlanSdkApiPath,
} from './codeplan-api-manifest.js';
