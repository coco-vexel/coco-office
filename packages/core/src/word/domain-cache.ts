import type { Op } from './plan.js';
import type { ListParagraphsOptions, ListResult, ParagraphInfo } from './list-paragraphs.js';
import type { ListTablesOptions, RowHeight, TableCellInfo, TableInfo } from './list-tables.js';

export type DomainCacheState = 'empty' | 'ready' | 'dirty';
export type DomainCacheName = 'all' | 'paragraphs' | 'tables';

export interface ParagraphDomainCacheSnapshot {
  state: DomainCacheState;
  listAllBodyReady: boolean;
  cachedParagraphIds: string[];
  dirtyReason?: string;
  buildCount: number;
}

export interface WordDomainCacheSnapshot {
  revision: number;
  paragraphs: ParagraphDomainCacheSnapshot;
  tables: TableDomainCacheSnapshot;
}

export interface TableDomainCacheSnapshot {
  state: DomainCacheState;
  cachedProfiles: string[];
  cachedTableIds: string[];
  dirtyReason?: string;
  buildCount: number;
}

export class WordDomainCache {
  private revision = 0;
  readonly paragraphs = new ParagraphDomainCache();
  readonly tables = new TableDomainCache();

  reset(): void {
    this.revision = 0;
    this.paragraphs.clear();
    this.tables.clear();
  }

  currentRevision(): number {
    return this.revision;
  }

  afterSuccessfulApply(ops: readonly Op[]): void {
    this.revision++;
    if (ops.some(affectsParagraphTextOrStructure)) {
      this.paragraphs.markDirty('content mutation');
    }
    if (ops.some(affectsTableStructureOrContent)) {
      this.tables.markDirty('content mutation');
    }
  }

  invalidate(domain: DomainCacheName = 'all', reason = 'manual'): void {
    if (domain === 'all' || domain === 'paragraphs') {
      this.paragraphs.markDirty(reason);
    }
    if (domain === 'all' || domain === 'tables') {
      this.tables.markDirty(reason);
    }
  }

  snapshot(): WordDomainCacheSnapshot {
    return {
      revision: this.revision,
      paragraphs: this.paragraphs.snapshot(),
      tables: this.tables.snapshot(),
    };
  }
}

export class ParagraphDomainCache {
  private state: DomainCacheState = 'empty';
  private dirtyReason: string | undefined;
  private listAllBody: ListResult | undefined;
  private byId = new Map<string, ParagraphInfo | null>();
  private buildCount = 0;

  getCacheableList(
    opts: ListParagraphsOptions | undefined,
    build: () => ListResult,
  ): ListResult | undefined {
    if (!isCacheableParagraphList(opts)) return undefined;
    this.flushIfDirty();
    if (!this.listAllBody) {
      const result = build();
      this.listAllBody = cloneListResult(result);
      this.byId.clear();
      for (const paragraph of this.listAllBody.paragraphs) {
        this.byId.set(paragraph.paraId.toUpperCase(), paragraph);
      }
      this.state = 'ready';
      this.dirtyReason = undefined;
      this.buildCount++;
    }
    return cloneListResult(this.listAllBody);
  }

  getParagraph(paraId: string, build: () => ParagraphInfo | null): ParagraphInfo | null {
    this.flushIfDirty();
    const key = paraId.toUpperCase();
    if (!this.byId.has(key)) {
      const value = build();
      this.byId.set(key, value ? cloneParagraphInfo(value) : null);
      this.state = 'ready';
      this.dirtyReason = undefined;
      this.buildCount++;
    }
    const value = this.byId.get(key) ?? null;
    return value ? cloneParagraphInfo(value) : null;
  }

  markDirty(reason: string): void {
    if (this.state === 'empty') return;
    this.state = 'dirty';
    this.dirtyReason = reason;
  }

  clear(): void {
    this.state = 'empty';
    this.dirtyReason = undefined;
    this.listAllBody = undefined;
    this.byId.clear();
    this.buildCount = 0;
  }

  snapshot(): ParagraphDomainCacheSnapshot {
    return {
      state: this.state,
      listAllBodyReady: !!this.listAllBody && this.state !== 'dirty',
      cachedParagraphIds: Array.from(this.byId.keys()),
      dirtyReason: this.dirtyReason,
      buildCount: this.buildCount,
    };
  }

  private flushIfDirty(): void {
    if (this.state !== 'dirty') return;
    this.state = 'empty';
    this.dirtyReason = undefined;
    this.listAllBody = undefined;
    this.byId.clear();
  }
}

export function isCacheableParagraphList(opts: ListParagraphsOptions | undefined): boolean {
  if (!opts) return false;
  if (opts.scope !== undefined && opts.scope !== 'body') return false;
  if (opts.filter !== undefined) return false;
  if (opts.textFormat !== undefined && opts.textFormat !== 'full') return false;
  if (opts.includeFormat === true) return false;
  if (opts.structuralWhitespace === true) return false;
  if (opts.includeTableInfo === true) return false;
  return opts.pagination?.mode === 'all';
}

type TableListProfile = 'basic' | 'withCellTexts';

export class TableDomainCache {
  private state: DomainCacheState = 'empty';
  private dirtyReason: string | undefined;
  private lists = new Map<TableListProfile, TableInfo[]>();
  private byIdBasic = new Map<string, TableInfo | null>();
  private buildCount = 0;

  getCacheableList(
    opts: ListTablesOptions | undefined,
    build: () => TableInfo[],
  ): TableInfo[] | undefined {
    const profile = tableListProfile(opts);
    if (!profile) return undefined;
    this.flushIfDirty();
    if (!this.lists.has(profile)) {
      const tables = build().map(cloneTableInfo);
      this.lists.set(profile, tables);
      if (profile === 'basic') {
        this.byIdBasic.clear();
        for (const table of tables) {
          this.byIdBasic.set(table.tableId, table);
        }
      }
      this.state = 'ready';
      this.dirtyReason = undefined;
      this.buildCount++;
    }
    return this.lists.get(profile)!.map(cloneTableInfo);
  }

  getTable(tableId: string, build: () => TableInfo | null): TableInfo | null {
    this.flushIfDirty();
    if (!this.byIdBasic.has(tableId)) {
      const table = build();
      this.byIdBasic.set(tableId, table ? cloneTableInfo(table) : null);
      this.state = 'ready';
      this.dirtyReason = undefined;
      this.buildCount++;
    }
    const table = this.byIdBasic.get(tableId) ?? null;
    return table ? cloneTableInfo(table) : null;
  }

  markDirty(reason: string): void {
    if (this.state === 'empty') return;
    this.state = 'dirty';
    this.dirtyReason = reason;
  }

  clear(): void {
    this.state = 'empty';
    this.dirtyReason = undefined;
    this.lists.clear();
    this.byIdBasic.clear();
    this.buildCount = 0;
  }

  snapshot(): TableDomainCacheSnapshot {
    return {
      state: this.state,
      cachedProfiles: this.state === 'dirty' ? [] : Array.from(this.lists.keys()),
      cachedTableIds: this.state === 'dirty' ? [] : Array.from(this.byIdBasic.keys()),
      dirtyReason: this.dirtyReason,
      buildCount: this.buildCount,
    };
  }

  private flushIfDirty(): void {
    if (this.state !== 'dirty') return;
    this.state = 'empty';
    this.dirtyReason = undefined;
    this.lists.clear();
    this.byIdBasic.clear();
  }
}

export function isCacheableTableList(opts: ListTablesOptions | undefined): boolean {
  return tableListProfile(opts) !== null;
}

function tableListProfile(opts: ListTablesOptions | undefined): TableListProfile | null {
  if (opts?.scope !== undefined && opts.scope !== 'body') return null;
  if (opts?.includeNested === true) return null;
  return opts?.includeCellTexts === true ? 'withCellTexts' : 'basic';
}

function affectsParagraphTextOrStructure(op: Op): boolean {
  switch (op.type) {
    case 'replaceParagraph':
    case 'replaceTextInParagraph':
    case 'formatParagraph':
    case 'insertParagraphBefore':
    case 'insertParagraphAfter':
    case 'insertParagraphAtBodyStart':
    case 'deleteParagraph':
    case 'deleteTable':
    case 'replaceTableCellWithSdt':
    case 'deleteTableRow':
    case 'replaceTableCellText':
    case 'insertTableCellParagraph':
    case 'replaceShapeInnerText':
    case 'insertTableRow':
    case 'cloneTableAfter':
    case 'insertImageAfter':
    case 'insertTableFrom':
      return true;
    default:
      return false;
  }
}

function affectsTableStructureOrContent(op: Op): boolean {
  switch (op.type) {
    case 'deleteTable':
    case 'replaceTableCellWithSdt':
    case 'deleteTableRow':
    case 'replaceTableCellText':
    case 'formatTableCell':
    case 'insertTableCellParagraph':
    case 'insertTableRow':
    case 'cloneTableAfter':
    case 'insertTableFrom':
    case 'setTableColumnWidths':
    case 'setTableRowHeights':
      return true;
    case 'deleteParagraph':
    case 'insertParagraphBefore':
    case 'insertParagraphAfter':
    case 'insertParagraphAtBodyStart':
    case 'insertImageAfter':
    case 'deleteDrawing':
    case 'replaceDrawingWithPlaceholder':
      return true;
    default:
      return false;
  }
}

function cloneListResult(result: ListResult): ListResult {
  return {
    paragraphs: result.paragraphs.map(cloneParagraphInfo),
    pagination: { ...result.pagination },
  };
}

function cloneParagraphInfo(info: ParagraphInfo): ParagraphInfo {
  return {
    ...info,
    format: info.format
      ? {
          ...info.format,
          indent: info.format.indent ? { ...info.format.indent } : undefined,
          spacing: info.format.spacing ? { ...info.format.spacing } : undefined,
        }
      : undefined,
    table: info.table ? { ...info.table } : undefined,
  };
}

function cloneTableInfo(info: TableInfo): TableInfo {
  return {
    ...info,
    columnWidths: [...info.columnWidths],
    rowHeights: info.rowHeights.map(cloneRowHeightOrNull),
    cells: info.cells ? info.cells.map((row) => [...row]) : undefined,
    cellGrid: info.cellGrid ? info.cellGrid.map((row) => row.map(cloneTableCellInfo)) : undefined,
  };
}

function cloneRowHeightOrNull(rowHeight: RowHeight | null): RowHeight | null {
  return rowHeight ? { ...rowHeight } : null;
}

function cloneTableCellInfo(cell: TableCellInfo): TableCellInfo {
  return {
    ...cell,
    shading: cell.shading ? { ...cell.shading } : undefined,
    paragraphIds: [...cell.paragraphIds],
    blocks: cell.blocks.map((block) => ({ ...block })),
    imageDrawingIds: [...cell.imageDrawingIds],
  };
}
