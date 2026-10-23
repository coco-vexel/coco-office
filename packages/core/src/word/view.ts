export { W_NS, W14_NS } from './namespaces.js';
export { buildIndices, type PackageIndices } from './indices.js';
export { buildParagraphInfo, type ParagraphFormat, type ParagraphInfo } from './list-paragraphs.js';
export { listTables, type TableInfo, type TableCellInfo } from './list-tables.js';
export { listDrawings, type DrawingInfo } from './list-drawings.js';
export {
  listSections,
  type HeaderFooterRef,
  type HeaderFooterType,
  type SectionInfo,
  type SectionType,
} from './section-queries.js';
export { getImageBytesSync, type ImageBytesResult } from './plan-executors/insert-image-ops.js';
export { getStoryParagraphs, listStories, type StoryId } from './stories.js';
export type { CommentInfo } from './comment-queries.js';
export type {
  RevisionInfo,
  RevisionKind,
  SupportedRevisionLocalName,
  SupportedW14RevisionLocalName,
  SupportedRevisionType,
} from './revision-queries.js';
