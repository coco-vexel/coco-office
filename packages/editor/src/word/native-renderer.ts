import type {
  NativeWordBlockNode,
  NativeWordDocument,
  NativeWordDrawingNode,
  NativeWordFieldNode,
  NativeWordHeaderFooterStory,
  NativeWordImageNode,
  NativeWordParagraphNode,
  NativeWordTableCellNode,
  NativeWordTableNode,
} from './native-model.js';
import type { EditorTableBorders } from './model.js';
import { resolveParagraphStyle, resolveRunStyle, resolveTableCellStyle } from './style-resolver.js';
import {
  safeEmbeddedRasterImageUrl,
  safeRenderedHyperlinkHref,
} from './native-url-safety.js';

export interface NativeAnchorLayoutMetadata {
  classNames: string[];
  dataset: Record<string, string>;
  styleProperties: Record<string, string>;
}

interface NativeRenderOptions {
  implicitGridForParagraphs?: boolean;
  implicitGridLinePitch?: number;
}

const NATIVE_ANNOTATION_RAIL_WIDTH_PX = 260;

export function defineNativeWordElements(): void {
  defineElement('coco-word-native-document');
  defineElement('coco-word-native-page');
  defineElement('coco-word-native-page-body');
  defineElement('coco-word-native-story');
  defineElement('coco-word-native-p');
  defineElement('coco-word-native-drawing');
  defineElement('coco-word-native-table');
}

export function renderNativeDocument(doc: NativeWordDocument): HTMLElement {
  const root = document.createElement('coco-word-native-document');
  root.contentEditable = 'true';
  root.spellcheck = false;
  if (shouldUseNaturalParagraphFlow(doc)) root.classList.add('uses-natural-paragraph-flow');
  if (shouldUseCompactLatinTables(doc)) root.classList.add('uses-compact-latin-tables');
  if (shouldSplitRowspanTableGroups(doc)) root.classList.add('uses-rowspan-table-group-splitting');
  const defaultRunStyle = resolveRunStyle(doc.defaults?.run).cssText;
  if (defaultRunStyle.length > 0) root.setAttribute('style', defaultRunStyle);
  for (const story of doc.headers) root.appendChild(renderNativeHeaderFooterStory(story));
  const implicitGridLinePitch = documentGridLinePitch(doc);
  let implicitGridForParagraphs = implicitGridLinePitch !== undefined && nativeDocumentHasCjkText(doc);
  for (const block of doc.blocks) {
    root.appendChild(renderNativeBlock(block, { implicitGridForParagraphs, implicitGridLinePitch }));
    if (nativeBlockHasSnapToGridDisabled(block)) implicitGridForParagraphs = false;
  }
  for (const story of doc.footers) root.appendChild(renderNativeHeaderFooterStory(story));
  const hasAnnotations = root.querySelector(
    '.coco-word-comment, .coco-word-revision-insert, .coco-word-revision-delete, .coco-word-revision-format, [data-revision-id]',
  ) !== null;
  if (hasAnnotations) {
    root.classList.add('has-annotation-rail');
    if (doc.pageLayout?.width !== undefined) {
      root.style.minWidth = `${twipsToPxNumber(doc.pageLayout.width) + NATIVE_ANNOTATION_RAIL_WIDTH_PX}px`;
    }
  }
  return root;
}

function shouldUseNaturalParagraphFlow(doc: NativeWordDocument): boolean {
  if (doc.blocks.length > 300) return false;
  const layouts = [doc.pageLayout, ...doc.sections.map((section) => section.pageLayout)].filter(
    (layout): layout is NonNullable<NativeWordDocument['pageLayout']> => layout !== undefined,
  );
  if (layouts.length === 0) return false;
  if (layouts.every((layout) => layout.docGrid === undefined)) return true;
  return !nativeDocumentHasCjkText(doc);
}

function documentGridLinePitch(doc: NativeWordDocument): number | undefined {
  const layouts = [doc.pageLayout, ...doc.sections.map((section) => section.pageLayout)].filter(
    (layout): layout is NonNullable<NativeWordDocument['pageLayout']> => layout !== undefined,
  );
  return layouts.find((layout) => layout.docGrid?.linePitch !== undefined)?.docGrid?.linePitch;
}

function shouldUseCompactLatinTables(doc: NativeWordDocument): boolean {
  return doc.blocks.some((block) => block.kind === 'table') && !nativeDocumentHasCjkText(doc);
}

function shouldSplitRowspanTableGroups(doc: NativeWordDocument): boolean {
  const tables = doc.blocks.filter((block) => block.kind === 'table');
  if (tables.length < 10) return false;
  return nativeDocumentImageCount(doc) <= 5;
}

function nativeDocumentImageCount(doc: NativeWordDocument): number {
  return doc.blocks.reduce((count, block) => count + nativeBlockImageCount(block), 0);
}

function nativeBlockImageCount(block: NativeWordBlockNode): number {
  if (block.kind === 'paragraph') {
    return block.content.filter((child) => child.kind === 'image' || child.kind === 'drawing').length;
  }
  return block.rows.reduce(
    (count, row) =>
      count + row.cells.reduce((cellCount, cell) => cellCount + cell.blocks.reduce((sum, child) => sum + nativeBlockImageCount(child), 0), 0),
    0,
  );
}

function nativeDocumentHasCjkText(doc: NativeWordDocument): boolean {
  let cjk = 0;
  let total = 0;
  for (const block of doc.blocks.slice(0, 160)) {
    for (const text of nativeBlockTextSamples(block)) {
      for (const char of text) {
        if (/\s/.test(char)) continue;
        total++;
        if (/[\u3400-\u9FFF\uF900-\uFAFF]/.test(char)) cjk++;
      }
    }
    if (total >= 1200) break;
  }
  return cjk >= 24 && cjk / Math.max(total, 1) >= 0.08;
}

function nativeBlockTextSamples(block: NativeWordBlockNode): string[] {
  if (block.kind === 'paragraph') {
    return block.content.filter((child) => child.kind === 'text').map((child) => child.text);
  }
  const text: string[] = [];
  for (const row of block.rows) {
    for (const cell of row.cells) {
      for (const paragraph of cell.paragraphs) {
        text.push(...paragraph.content.filter((child) => child.kind === 'text').map((child) => child.text));
      }
    }
  }
  return text;
}

export function paginateNativeDocument(root: HTMLElement, doc: NativeWordDocument): void {
  const defaultLayout = doc.pageLayout;
  if (!defaultLayout || root.dataset.paginated === 'true') return;
  root.dataset.paginated = 'true';
  root.classList.add('is-paginated');
  const originalChildren = Array.from(root.children);
  const headerTemplates = originalChildren.filter(
    (child) =>
      child instanceof HTMLElement &&
      child.tagName.toLowerCase() === 'coco-word-native-story' &&
      child.dataset.storyKind === 'header',
  ) as HTMLElement[];
  const footerTemplates = originalChildren.filter(
    (child) =>
      child instanceof HTMLElement &&
      child.tagName.toLowerCase() === 'coco-word-native-story' &&
      child.dataset.storyKind === 'footer',
  ) as HTMLElement[];
  const bodyBlocks = originalChildren.filter(
    (child) =>
      child instanceof HTMLElement &&
      !(child.tagName.toLowerCase() === 'coco-word-native-story'),
  ) as HTMLElement[];
  root.replaceChildren();

  const sectionPageCounts = new Map<number, number>();
  let currentSectionIndex = 0;
  let currentPageLayout = pageLayoutForSection(doc, currentSectionIndex, defaultLayout);
  const createPage = (sectionIndex = currentSectionIndex) => {
    const sectionPageNumber = (sectionPageCounts.get(sectionIndex) ?? 0) + 1;
    sectionPageCounts.set(sectionIndex, sectionPageNumber);
    currentSectionIndex = sectionIndex;
    currentPageLayout = pageLayoutForSection(doc, sectionIndex, defaultLayout);
    return appendNativePage(
      root,
      doc,
      currentPageLayout,
      headerTemplates,
      footerTemplates,
      root.children.length + 1,
      sectionIndex,
      sectionPageNumber,
    );
  };
  let page = createPage();
  let body = page.querySelector<HTMLElement>('coco-word-native-page-body')!;
  const bodyLimit = () =>
    Math.max(120, Math.round(body.getBoundingClientRect().height) || contentHeightForLayout(currentPageLayout));
  const nextPage = () => {
    page = createPage();
    body = page.querySelector<HTMLElement>('coco-word-native-page-body')!;
    return body;
  };
  const moveEmptyPageToSection = (sectionIndex: number) => {
    if (sectionIndex === currentSectionIndex || body.children.length > 0) return;
    const oldCount = sectionPageCounts.get(currentSectionIndex) ?? 0;
    if (oldCount > 1) sectionPageCounts.set(currentSectionIndex, oldCount - 1);
    else sectionPageCounts.delete(currentSectionIndex);
    const sectionPageNumber = (sectionPageCounts.get(sectionIndex) ?? 0) + 1;
    sectionPageCounts.set(sectionIndex, sectionPageNumber);
    currentSectionIndex = sectionIndex;
    currentPageLayout = pageLayoutForSection(doc, sectionIndex, defaultLayout);
    refreshNativePageChrome(
      page,
      doc,
      currentPageLayout,
      headerTemplates,
      footerTemplates,
      Number(page.dataset.pageNumber ?? '1'),
      sectionIndex,
      sectionPageNumber,
    );
  };
  const appendBlock = (block: HTMLElement, isFinalBodyBlock: boolean) => {
    const blockSectionIndex = sectionIndexForBlock(block, doc) ?? currentSectionIndex;
    if (body.children.length === 0) {
      moveEmptyPageToSection(blockSectionIndex);
    }
    if (shouldStartBlockOnNewPage(block) && body.children.length > 0) {
      nextPage();
      moveEmptyPageToSection(blockSectionIndex);
    }
    body.appendChild(block);
    const promotesToEmptyFlow = nativeBlockPromotesToEmptyFlow(block);
    if (promotesToEmptyFlow) promotePageRelativeAnchors(page);
    const hasForcedPageBreak = block.querySelector('.coco-word-page-break') !== null;
    if (body.scrollHeight > bodyLimit()) {
      const keepWithPrevious = previousKeepNextBlock(body, block);
      if (keepWithPrevious) {
        if (block.matches('coco-word-native-table')) {
          body.removeChild(block);
          const split = splitNativeTableAcrossPages(block, body, bodyLimit, nextPage);
          if (split) {
            if (!promotesToEmptyFlow) promotePageRelativeAnchors(page);
            const sectionBreakType = sectionBreakTypeAfterBlock(block, doc);
            const nextSectionIndex = sectionIndexAfterBlock(block, doc);
            if (nextSectionIndex !== null) currentSectionIndex = nextSectionIndex;
            if (hasForcedPageBreak || sectionBreakType) advanceAfterForcedBreak(nextPage, sectionBreakType);
            return;
          }
          body.appendChild(block);
        }
        body.removeChild(block);
        body.removeChild(keepWithPrevious);
        nextPage();
        moveEmptyPageToSection(blockSectionIndex);
        body.appendChild(keepWithPrevious);
        body.appendChild(block);
        if (block.matches('coco-word-native-table') && body.scrollHeight > bodyLimit()) {
          body.removeChild(block);
          if (!splitNativeTableAcrossPages(block, body, bodyLimit, nextPage)) {
            body.appendChild(block);
          }
        }
      } else {
        body.removeChild(block);
        const splitParagraphTail = splitNativeParagraphToFitCurrentPage(block, body, bodyLimit);
        if (splitParagraphTail) {
          nextPage();
          moveEmptyPageToSection(blockSectionIndex);
          body.appendChild(splitParagraphTail);
        } else {
          const split = splitNativeTableAcrossPages(block, body, bodyLimit, nextPage);
          if (!split) {
            nextPage();
            moveEmptyPageToSection(blockSectionIndex);
            body.appendChild(block);
          }
        }
      }
    } else if (block.matches('coco-word-native-table')) {
      splitNativeTableTailForBottomReserve(block, body, bodyLimit, nextPage, isFinalBodyBlock);
    }
    if (!promotesToEmptyFlow) promotePageRelativeAnchors(page);
    const sectionBreakType = sectionBreakTypeAfterBlock(block, doc);
    const nextSectionIndex = sectionIndexAfterBlock(block, doc);
    if (nextSectionIndex !== null) currentSectionIndex = nextSectionIndex;
    if (hasForcedPageBreak || sectionBreakType) advanceAfterForcedBreak(nextPage, sectionBreakType);
  };
  for (let index = 0; index < bodyBlocks.length; index += 1) {
    appendBlock(bodyBlocks[index]!, index === bodyBlocks.length - 1);
  }
  compactTrailingEmptyFlowPages(root);
  compactEmptyNativePages(root);
  renumberNativePages(root);
  refreshNativePageFields(root, doc);
  refreshNativePageCommentBubbles(root);
}

export function reflowNativeDocumentFromParagraph(
  root: HTMLElement,
  doc: NativeWordDocument,
  paraId: string,
): boolean {
  const paragraph = root.querySelector<HTMLElement>(
    `coco-word-native-page-body coco-word-native-p[data-para-id="${paraId}"]`,
  );
  const startPage = paragraph?.closest<HTMLElement>('coco-word-native-page');
  if (!paragraph || !startPage || !doc.pageLayout) return false;

  let pages = directNativePages(root);
  let startIndex = pages.indexOf(startPage);
  if (startIndex < 0) return false;
  const templates = collectIncrementalStoryTemplates(root);
  let changed = false;
  let pageCountChanged = false;

  for (let index = startIndex; index < pages.length; index += 1) {
    const page = pages[index]!;
    const body = nativePageBody(page);
    if (!body) continue;
    while (nativePageBodyOverflows(body)) {
      if (moveTrailingIncrementalTableRows(root, doc, page, pages[index + 1], templates)) {
        pages = directNativePages(root);
        changed = true;
        continue;
      }
      const moving = trailingIncrementalBlockGroup(body);
      if (moving.length === 0 || moving.length >= body.children.length) break;
      let nextPage = pages[index + 1];
      if (!nextPage || cannotReceivePreviousPageContent(page, nextPage, doc)) {
        nextPage = insertIncrementalNativePage(root, doc, page, nextPage, templates);
        pages = directNativePages(root);
        pageCountChanged = true;
      }
      const nextBody = nativePageBody(nextPage);
      if (!nextBody) break;
      for (const block of [...moving].reverse()) nextBody.insertBefore(block, nextBody.firstChild);
      changed = true;
    }
  }

  pages = directNativePages(root);
  startIndex = Math.max(0, Math.min(startIndex, Math.max(0, pages.length - 1)) - 1);
  for (let index = startIndex; index < pages.length - 1; index += 1) {
    const page = pages[index]!;
    const nextPage = pages[index + 1]!;
    const body = nativePageBody(page);
    const nextBody = nativePageBody(nextPage);
    if (!body || !nextBody || cannotReceivePreviousPageContent(page, nextPage, doc)) continue;
    while (backfillIncrementalTableRows(body, nextBody)) changed = true;
    while (true) {
      const candidates = leadingIncrementalBlockGroup(nextBody);
      if (candidates.length === 0 || !canBackfillIncrementalBlock(body, candidates[0]!, doc)) break;
      for (const candidate of candidates) body.appendChild(candidate);
      if (nativePageBodyOverflows(body, 8)) {
        for (const candidate of [...candidates].reverse()) nextBody.insertBefore(candidate, nextBody.firstChild);
        if (candidates.length > 1) {
          const first = candidates[0]!;
          body.appendChild(first);
          if (!nativePageBodyOverflows(body, 8)) {
            changed = true;
            continue;
          }
          nextBody.insertBefore(first, nextBody.firstChild);
        }
        break;
      }
      changed = true;
    }
    if (isEmptyNativePage(nextPage)) {
      nextPage.remove();
      pages = directNativePages(root);
      index -= 1;
      pageCountChanged = true;
    }
  }

  if (!changed) return false;
  compactTrailingEmptyFlowPages(root);
  compactEmptyNativePages(root);
  renumberNativePages(root);
  if (pageCountChanged) refreshIncrementalPageChrome(root, doc, templates);
  refreshNativePageFields(root, doc);
  refreshNativePageCommentBubbles(root);
  return true;
}

export interface NativeDocumentReflowMeasurement {
  overflowPx: number;
  canBackfillPrevious: boolean;
  canBackfillNext: boolean;
}

export function measureNativeDocumentReflowFromParagraph(
  root: HTMLElement,
  doc: NativeWordDocument,
  paraId: string,
): NativeDocumentReflowMeasurement {
  const empty = { overflowPx: 0, canBackfillPrevious: false, canBackfillNext: false };
  const paragraph = root.querySelector<HTMLElement>(
    `coco-word-native-page-body coco-word-native-p[data-para-id="${paraId}"]`,
  );
  const page = paragraph?.closest<HTMLElement>('coco-word-native-page');
  const pages = directNativePages(root);
  const pageIndex = page ? pages.indexOf(page) : -1;
  if (!paragraph || !page || pageIndex < 0) return empty;
  const body = nativePageBody(page);
  if (!body) return empty;
  const limit = Math.max(120, Math.round(body.getBoundingClientRect().height));
  const overflowPx = Math.max(0, body.scrollHeight - limit);

  const previousPage = pages[pageIndex - 1];
  const previousBody = previousPage ? nativePageBody(previousPage) : null;
  const canBackfillPrevious = !!(
    previousPage &&
    previousBody &&
    leadingIncrementalBlockGroup(body).some((block) => blockContainsParagraph(block, paraId)) &&
    nativePagePairCanBackfill(previousPage, page, previousBody, body, doc)
  );

  const nextPage = pages[pageIndex + 1];
  const nextBody = nextPage ? nativePageBody(nextPage) : null;
  const canBackfillNext = !!(
    nextPage &&
    nextBody &&
    nativePagePairCanBackfill(page, nextPage, body, nextBody, doc)
  );
  return { overflowPx, canBackfillPrevious, canBackfillNext };
}

function nativePagePairCanBackfill(
  page: HTMLElement,
  nextPage: HTMLElement,
  body: HTMLElement,
  nextBody: HTMLElement,
  doc: NativeWordDocument,
): boolean {
  if (cannotReceivePreviousPageContent(page, nextPage, doc)) return false;
  const available = Math.max(0, Math.round(body.getBoundingClientRect().height) - nativeBodyContentBottom(body));
  if (available <= 0) return false;

  const sourceTable = body.lastElementChild;
  const targetTable = nextBody.firstElementChild;
  if (
    sourceTable instanceof HTMLElement &&
    targetTable instanceof HTMLElement &&
    sourceTable.matches('coco-word-native-table') &&
    targetTable.matches('coco-word-native-table') &&
    sourceTable.dataset.tableId === targetTable.dataset.tableId
  ) {
    const targetRows = Array.from(
      targetTable.querySelectorAll<HTMLTableRowElement>(':scope > table > tbody > tr'),
    ).filter((row) => row.dataset.repeatedHeader !== 'true');
    const rowGroup = groupTableRowsForPagination(targetRows)[0];
    if (rowGroup && !tableRowGroupStartsWithHeader(rowGroup) && tableRowGroupHeight(rowGroup) <= available + 8) {
      return true;
    }
  }

  const candidates = leadingIncrementalBlockGroup(nextBody);
  if (candidates.length === 0 || !canBackfillIncrementalBlock(body, candidates[0]!, doc)) return false;
  return incrementalBlockGroupHeight(candidates) <= available + 8;
}

function incrementalBlockGroupHeight(blocks: HTMLElement[]): number {
  return blocks.reduce((height, block) => {
    const style = window.getComputedStyle(block);
    const marginTop = Number.parseFloat(style.marginTop) || 0;
    const marginBottom = Number.parseFloat(style.marginBottom) || 0;
    return height + block.getBoundingClientRect().height + marginTop + marginBottom;
  }, 0);
}

export function rebuildNativeTableFragmentsFromParagraph(
  root: HTMLElement,
  doc: NativeWordDocument,
  paraId: string,
): boolean {
  const paragraph = root.querySelector<HTMLElement>(
    `coco-word-native-page-body coco-word-native-p[data-para-id="${paraId}"]`,
  );
  const startPage = paragraph?.closest<HTMLElement>('coco-word-native-page');
  const pages = directNativePages(root);
  const startPageIndex = startPage ? pages.indexOf(startPage) : -1;
  if (startPageIndex < 0) return false;

  let rebuilt = false;
  for (const block of doc.blocks) {
    if (block.kind !== 'table') continue;
    const fragments = Array.from(
      root.querySelectorAll<HTMLElement>(`coco-word-native-table[data-table-id="${block.tableId}"]`),
    );
    const first = fragments[0];
    const firstPage = first?.closest<HTMLElement>('coco-word-native-page');
    if (!first || !firstPage || pages.indexOf(firstPage) < startPageIndex) continue;
    first.replaceWith(renderNativeTable(block));
    for (const fragment of fragments.slice(1)) fragment.remove();
    rebuilt = true;
  }
  return rebuilt;
}

function directNativePages(root: HTMLElement): HTMLElement[] {
  return Array.from(root.children).filter(
    (child): child is HTMLElement =>
      child instanceof HTMLElement && child.tagName.toLowerCase() === 'coco-word-native-page',
  );
}

function nativePageBody(page: HTMLElement): HTMLElement | null {
  return page.querySelector<HTMLElement>(':scope > coco-word-native-page-body');
}

function nativePageBodyOverflows(body: HTMLElement, tolerance = 1): boolean {
  const limit = Math.max(120, Math.round(body.getBoundingClientRect().height));
  return body.scrollHeight > limit + tolerance;
}

function trailingIncrementalBlockGroup(body: HTMLElement): HTMLElement[] {
  const last = body.lastElementChild;
  if (!(last instanceof HTMLElement)) return [];
  const moving = [last];
  const previous = last.previousElementSibling;
  if (previous instanceof HTMLElement && previous.matches('coco-word-native-p[data-keep-next="true"]')) {
    moving.unshift(previous);
  }
  return moving;
}

function leadingIncrementalBlockGroup(body: HTMLElement): HTMLElement[] {
  const first = body.firstElementChild;
  if (!(first instanceof HTMLElement)) return [];
  const moving = [first];
  const next = first.nextElementSibling;
  if (first.matches('coco-word-native-p[data-keep-next="true"]') && next instanceof HTMLElement) moving.push(next);
  return moving;
}

function moveTrailingIncrementalTableRows(
  root: HTMLElement,
  doc: NativeWordDocument,
  page: HTMLElement,
  existingNextPage: HTMLElement | undefined,
  templates: { headers: HTMLElement[]; footers: HTMLElement[] },
): boolean {
  const body = nativePageBody(page);
  const sourceBlock = body?.lastElementChild;
  if (!(sourceBlock instanceof HTMLElement) || !sourceBlock.matches('coco-word-native-table')) return false;
  const sourceTable = sourceBlock.querySelector<HTMLTableElement>(':scope > table');
  const sourceBody = sourceTable?.querySelector<HTMLTableSectionElement>(':scope > tbody');
  if (!sourceTable || !sourceBody) return false;
  const rows = Array.from(sourceBody.children).filter(isTableRow);
  const dataRows = rows.filter((row) => row.dataset.repeatedHeader !== 'true');
  const groups = groupTableRowsForPagination(dataRows);
  if (groups.length < 2) return false;
  const moving = groups[groups.length - 1]!;
  if (tableRowGroupStartsWithHeader(moving)) return false;

  let nextPage = existingNextPage;
  if (!nextPage || cannotReceivePreviousPageContent(page, nextPage, doc)) {
    nextPage = insertIncrementalNativePage(root, doc, page, nextPage, templates);
  }
  const nextBody = nativePageBody(nextPage);
  if (!nextBody) return false;
  const tableId = sourceBlock.dataset.tableId;
  let targetBlock = Array.from(nextBody.children).find(
    (child): child is HTMLElement =>
      child instanceof HTMLElement &&
      child.matches('coco-word-native-table') &&
      child.dataset.tableId === tableId,
  );
  if (!targetBlock) {
    targetBlock = createNativeTableChunk(sourceBlock, sourceTable);
    nextBody.insertBefore(targetBlock, nextBody.firstChild);
    const targetBody = targetBlock.querySelector<HTMLTableSectionElement>(':scope > table > tbody');
    if (targetBody) appendClonedTableHeaderRows(targetBody, tableHeaderRowsForIncrementalTable(root, tableId));
  }
  const targetBody = targetBlock.querySelector<HTMLTableSectionElement>(':scope > table > tbody');
  if (!targetBody) return false;
  const insertionPoint = Array.from(targetBody.children).find(
    (row) => row instanceof HTMLTableRowElement && row.dataset.repeatedHeader !== 'true',
  );
  removeTableRowGroup(sourceBody, moving);
  for (const row of moving) targetBody.insertBefore(row, insertionPoint ?? null);
  return true;
}

function backfillIncrementalTableRows(body: HTMLElement, nextBody: HTMLElement): boolean {
  const sourceBlock = body.lastElementChild;
  const targetBlock = nextBody.firstElementChild;
  if (
    !(sourceBlock instanceof HTMLElement) ||
    !(targetBlock instanceof HTMLElement) ||
    !sourceBlock.matches('coco-word-native-table') ||
    !targetBlock.matches('coco-word-native-table') ||
    sourceBlock.dataset.tableId !== targetBlock.dataset.tableId
  ) {
    return false;
  }
  const sourceBody = sourceBlock.querySelector<HTMLTableSectionElement>(':scope > table > tbody');
  const targetBody = targetBlock.querySelector<HTMLTableSectionElement>(':scope > table > tbody');
  if (!sourceBody || !targetBody) return false;
  const dataRows = Array.from(targetBody.children)
    .filter(isTableRow)
    .filter((row) => row.dataset.repeatedHeader !== 'true');
  const moving = groupTableRowsForPagination(dataRows)[0];
  if (!moving || tableRowGroupStartsWithHeader(moving)) return false;
  const restoreBefore = moving[moving.length - 1]?.nextElementSibling ?? null;
  appendTableRowGroup(sourceBody, moving);
  if (nativePageBodyOverflows(body, 8)) {
    for (const row of moving) targetBody.insertBefore(row, restoreBefore);
    return false;
  }
  if (targetBody.querySelector('tr:not([data-repeated-header="true"])') === null) targetBlock.remove();
  return true;
}

function tableHeaderRowsForIncrementalTable(root: HTMLElement, tableId: string | undefined): HTMLTableRowElement[] {
  if (!tableId) return [];
  const table = Array.from(root.querySelectorAll<HTMLElement>('coco-word-native-table')).find(
    (candidate) => candidate.dataset.tableId === tableId && candidate.querySelector('tr[data-repeat-header="true"]'),
  );
  return table ? Array.from(table.querySelectorAll<HTMLTableRowElement>('tr[data-repeat-header="true"]')) : [];
}

function cannotReceivePreviousPageContent(
  page: HTMLElement,
  nextPage: HTMLElement,
  doc: NativeWordDocument,
): boolean {
  if (page.dataset.sectionIndex !== nextPage.dataset.sectionIndex) return true;
  const body = nativePageBody(page);
  const nextBody = nativePageBody(nextPage);
  const last = body?.lastElementChild;
  const first = nextBody?.firstElementChild;
  if (first instanceof HTMLElement && shouldStartBlockOnNewPage(first)) return true;
  if (!(last instanceof HTMLElement)) return false;
  if (last.querySelector('.coco-word-page-break')) return true;
  const breakType = sectionBreakTypeAfterBlock(last, doc);
  return breakType !== null && breakType !== 'continuous' && breakType !== 'nextColumn';
}

function canBackfillIncrementalBlock(body: HTMLElement, candidate: HTMLElement, doc: NativeWordDocument): boolean {
  if (shouldStartBlockOnNewPage(candidate)) return false;
  const last = body.lastElementChild;
  if (!(last instanceof HTMLElement)) return true;
  if (last.querySelector('.coco-word-page-break')) return false;
  const breakType = sectionBreakTypeAfterBlock(last, doc);
  return breakType === null || breakType === 'continuous' || breakType === 'nextColumn';
}

function collectIncrementalStoryTemplates(root: HTMLElement): { headers: HTMLElement[]; footers: HTMLElement[] } {
  const seen = new Set<string>();
  const headers: HTMLElement[] = [];
  const footers: HTMLElement[] = [];
  for (const story of Array.from(root.querySelectorAll<HTMLElement>(':scope > coco-word-native-page > coco-word-native-story'))) {
    const kind = story.dataset.storyKind;
    const key = `${kind ?? ''}:${story.dataset.partName ?? ''}`;
    if (seen.has(key)) continue;
    seen.add(key);
    if (kind === 'header') headers.push(story);
    if (kind === 'footer') footers.push(story);
  }
  return { headers, footers };
}

function insertIncrementalNativePage(
  root: HTMLElement,
  doc: NativeWordDocument,
  sourcePage: HTMLElement,
  beforePage: HTMLElement | undefined,
  templates: { headers: HTMLElement[]; footers: HTMLElement[] },
): HTMLElement {
  const pages = directNativePages(root);
  const sourceIndex = pages.indexOf(sourcePage);
  const sectionIndex = parseNonNegativeInt(sourcePage.dataset.sectionIndex, 0);
  const sectionPageNumber =
    pages.slice(0, sourceIndex + 1).filter((page) => parseNonNegativeInt(page.dataset.sectionIndex, 0) === sectionIndex)
      .length + 1;
  const page = appendNativePage(
    root,
    doc,
    pageLayoutForSection(doc, sectionIndex, doc.pageLayout!),
    templates.headers,
    templates.footers,
    sourceIndex + 2,
    sectionIndex,
    sectionPageNumber,
  );
  if (beforePage) root.insertBefore(page, beforePage);
  return page;
}

function refreshIncrementalPageChrome(
  root: HTMLElement,
  doc: NativeWordDocument,
  templates: { headers: HTMLElement[]; footers: HTMLElement[] },
): void {
  for (const page of directNativePages(root)) {
    const pageNumber = parseNonNegativeInt(page.dataset.pageNumber, 1);
    const sectionIndex = parseNonNegativeInt(page.dataset.sectionIndex, 0);
    const sectionPageNumber = parseNonNegativeInt(page.dataset.sectionPageNumber, 1);
    refreshNativePageChrome(
      page,
      doc,
      pageLayoutForSection(doc, sectionIndex, doc.pageLayout!),
      templates.headers,
      templates.footers,
      pageNumber,
      sectionIndex,
      sectionPageNumber,
    );
  }
}

function shouldStartBlockOnNewPage(block: HTMLElement): boolean {
  return block.matches('coco-word-native-p[data-page-break-before="true"]');
}

function previousKeepNextBlock(body: HTMLElement, block: HTMLElement): HTMLElement | null {
  const previous = block.previousElementSibling;
  if (!(previous instanceof HTMLElement)) return null;
  if (!previous.matches('coco-word-native-p[data-keep-next="true"]')) return null;
  return previous.previousElementSibling instanceof HTMLElement ? previous : null;
}

function sectionBreakTypeAfterBlock(
  block: HTMLElement,
  doc: NativeWordDocument,
): NonNullable<NativeWordDocument['sections'][number]['type']> | 'nextPage' | null {
  if (!block.matches('coco-word-native-p[data-section-break="true"]')) return null;
  const paraId = block.dataset.paraId;
  if (!paraId) return 'nextPage';
  const section = doc.sections.find((item) => item.lastParaId === paraId);
  return section?.type ?? 'nextPage';
}

function sectionIndexAfterBlock(block: HTMLElement, doc: NativeWordDocument): number | null {
  const section = doc.sections.find((item) => item.lastParaId && blockContainsParagraph(block, item.lastParaId));
  if (!section) return null;
  return Math.min(section.sectionIndex + 1, Math.max(0, doc.sections.length - 1));
}

function sectionIndexForBlock(block: HTMLElement, doc: NativeWordDocument): number | null {
  if (doc.sections.length === 0) return null;
  const paraIds = paragraphIdsInBlock(block);
  if (paraIds.length === 0) return null;
  const directFirstMatch = doc.sections.find((section) => paraIds.includes(section.firstParaId));
  if (directFirstMatch) return directFirstMatch.sectionIndex;
  const directLastMatch = doc.sections.find((section) => paraIds.includes(section.lastParaId));
  if (directLastMatch) return directLastMatch.sectionIndex;
  return null;
}

function blockContainsParagraph(block: HTMLElement, paraId: string): boolean {
  return paragraphIdsInBlock(block).includes(paraId);
}

function paragraphIdsInBlock(block: HTMLElement): string[] {
  if (block.matches('coco-word-native-p')) {
    const paraId = block.dataset.paraId;
    return paraId ? [paraId] : [];
  }
  return Array.from(block.querySelectorAll<HTMLElement>('coco-word-native-p'))
    .map((paragraph) => paragraph.dataset.paraId)
    .filter((paraId): paraId is string => paraId !== undefined && paraId.length > 0);
}

function advanceAfterForcedBreak(
  nextPage: () => HTMLElement,
  sectionBreakType: NonNullable<NativeWordDocument['sections'][number]['type']> | 'nextPage' | null,
): void {
  if (sectionBreakType === 'continuous' || sectionBreakType === 'nextColumn') return;
  const nextBody = nextPage();
  const page = nextBody.closest('coco-word-native-page');
  const pageNumber = Number(page?.getAttribute('data-page-number') ?? '0');
  if (sectionBreakType === 'oddPage' && pageNumber % 2 === 0) nextPage();
  if (sectionBreakType === 'evenPage' && pageNumber % 2 === 1) nextPage();
}

/**
 * Keep ordinary paragraphs flowing across pages. Word's widow control only
 * limits orphan lines; it does not make a paragraph an indivisible block.
 */
function splitNativeParagraphToFitCurrentPage(
  block: HTMLElement,
  body: HTMLElement,
  bodyLimit: () => number,
): HTMLElement | null {
  if (!block.matches('coco-word-native-p')) return null;
  if (block.dataset.keepLines === 'true' || block.dataset.pageBreakBefore === 'true') return null;
  if (
    block.querySelector(
      'img, coco-word-native-drawing, coco-word-native-table, .coco-word-page-break, .coco-word-tab-row',
    )
  ) {
    return null;
  }

  const total = textLengthOfNode(block);
  if (total < 2) return null;

  let best: { head: HTMLElement; tail: HTMLElement } | null = null;
  let low = 1;
  let high = total - 1;
  while (low <= high) {
    const offset = Math.floor((low + high) / 2);
    const split = splitElementByTextOffset(block, offset);
    if (!split || textLengthOfNode(split.head) === 0 || textLengthOfNode(split.tail) === 0) {
      high = offset - 1;
      continue;
    }
    body.appendChild(split.head);
    if (body.scrollHeight <= bodyLimit()) {
      if (best?.head.parentElement === body) body.removeChild(best.head);
      best = split;
      low = offset + 1;
    } else {
      body.removeChild(split.head);
      high = offset - 1;
    }
  }

  if (!best || best.head.parentElement !== body) return null;
  return best.tail;
}

function splitNativeTableAcrossPages(
  block: HTMLElement,
  currentBody: HTMLElement,
  bodyLimit: () => number,
  nextPage: () => HTMLElement,
): boolean {
  if (block.tagName.toLowerCase() !== 'coco-word-native-table') return false;
  const table = block.querySelector(':scope > table');
  const sourceBody = table?.querySelector(':scope > tbody');
  const rows = sourceBody ? Array.from(sourceBody.children).filter(isTableRow) : [];
  const rowGroups = groupTableRowsForPagination(rows);
  if (!table || rowGroups.length < 2) return false;
  let targetBody = currentBody;
  let chunk = createNativeTableChunk(block, table);
  let chunkBody = chunk.querySelector('tbody') as HTMLTableSectionElement;
  const headerRows = rows.filter((row) => row.dataset.repeatHeader === 'true');
  const maxDataGroupsPerChunk = maxTableDataGroupsPerChunk(rowGroups, headerRows);
  let dataGroupsInChunk = 0;
  targetBody.appendChild(chunk);
  for (const group of rowGroups) {
    const startsWithHeader = tableRowGroupStartsWithHeader(group);
    if (!startsWithHeader && dataGroupsInChunk >= maxDataGroupsPerChunk) {
      targetBody = nextPage();
      chunk = createNativeTableChunk(block, table);
      chunkBody = chunk.querySelector('tbody') as HTMLTableSectionElement;
      targetBody.appendChild(chunk);
      appendClonedTableHeaderRows(chunkBody, headerRows);
      dataGroupsInChunk = 0;
    }
    appendTableRowGroup(chunkBody, group);
    if (!startsWithHeader) dataGroupsInChunk += 1;
    if (targetBody.children.length <= 1 && chunkBody.children.length <= group.length) continue;
    if (targetBody.scrollHeight <= bodyLimit()) continue;
    removeTableRowGroup(chunkBody, group);
    if (!startsWithHeader) dataGroupsInChunk = Math.max(0, dataGroupsInChunk - 1);
    const rowSplitTail = splitTableRowGroupByRowsToFitCurrentPage(chunkBody, group, targetBody, bodyLimit());
    const splitTail = rowSplitTail ?? splitTableRowGroupByBlocksToFitCurrentPage(chunkBody, group, targetBody, bodyLimit());
    const textSplitTail = splitTail ?? splitTableRowGroupByTextToFitCurrentPage(chunkBody, group, targetBody, bodyLimit());
    if (textSplitTail) {
      if (!startsWithHeader) dataGroupsInChunk += 1;
      targetBody = nextPage();
      chunk = createNativeTableChunk(block, table);
      chunkBody = chunk.querySelector('tbody') as HTMLTableSectionElement;
      targetBody.appendChild(chunk);
      dataGroupsInChunk = 0;
      appendClonedTableHeaderRows(chunkBody, headerRows);
      appendTableRowGroup(chunkBody, textSplitTail);
      if (!startsWithHeader) dataGroupsInChunk = 1;
      continue;
    }
    if (chunkBody.children.length === 0) targetBody.removeChild(chunk);
    targetBody = nextPage();
    chunk = createNativeTableChunk(block, table);
    chunkBody = chunk.querySelector('tbody') as HTMLTableSectionElement;
    targetBody.appendChild(chunk);
    dataGroupsInChunk = 0;
    if (!startsWithHeader) appendClonedTableHeaderRows(chunkBody, headerRows);
    appendTableRowGroup(chunkBody, group);
    if (!startsWithHeader) dataGroupsInChunk = 1;
  }
  return true;
}

function maxTableDataGroupsPerChunk(
  rowGroups: HTMLTableRowElement[][],
  headerRows: HTMLTableRowElement[],
): number {
  const dataGroupCount = rowGroups.filter((group) => !tableRowGroupStartsWithHeader(group)).length;
  if (headerRows.length === 0 || dataGroupCount < 30) return Number.POSITIVE_INFINITY;
  return 8;
}

function splitNativeTableTailForBottomReserve(
  block: HTMLElement,
  currentBody: HTMLElement,
  bodyLimit: () => number,
  nextPage: () => HTMLElement,
  isFinalBodyBlock: boolean,
): boolean {
  const limit = bodyLimit();
  const remaining = limit - nativeBodyContentBottom(currentBody);
  const reserve = tableBottomReservePx(block, limit);
  if (remaining >= reserve) return false;
  if (isFinalBodyBlock && isShortFlexFilledTablePage(currentBody, limit)) return false;
  const table = block.querySelector(':scope > table');
  const sourceBody = table?.querySelector<HTMLTableSectionElement>(':scope > tbody');
  const rows = sourceBody ? Array.from(sourceBody.children).filter(isTableRow) : [];
  const rowGroups = groupTableRowsForPagination(rows);
  if (!table || !sourceBody || rowGroups.length < 4) return false;

  const headerRows = rows.filter((row) => row.dataset.repeatHeader === 'true');
  const initialRemaining = remaining;
  let recoveredHeight = 0;
  const movedGroups: HTMLTableRowElement[][] = [];
  while (rowGroups.length > 1 && initialRemaining + recoveredHeight < reserve) {
    const group = rowGroups.pop();
    if (!group || tableRowGroupStartsWithHeader(group)) break;
    recoveredHeight += tableRowGroupHeight(group);
    removeTableRowGroup(sourceBody, group);
    movedGroups.unshift(group);
  }
  if (movedGroups.length === 0) return false;

  const targetBody = nextPage();
  const chunk = createNativeTableChunk(block, table);
  const chunkBody = chunk.querySelector('tbody') as HTMLTableSectionElement;
  targetBody.appendChild(chunk);
  if (!tableRowGroupStartsWithHeader(movedGroups[0] ?? [])) appendClonedTableHeaderRows(chunkBody, headerRows);
  for (const group of movedGroups) appendTableRowGroup(chunkBody, group);
  return true;
}

function tableRowGroupHeight(group: HTMLTableRowElement[]): number {
  return group.reduce((height, row) => height + row.getBoundingClientRect().height, 0);
}

function splitTableRowGroupByRowsToFitCurrentPage(
  target: HTMLTableSectionElement,
  group: HTMLTableRowElement[],
  targetBody: HTMLElement,
  limit: number,
): HTMLTableRowElement[] | null {
  if (!target.closest('coco-word-native-document.uses-rowspan-table-group-splitting')) return null;
  if (group.length < 2) return null;
  if (group.some((row) => row.dataset.repeatHeader === 'true' || row.dataset.cantSplit === 'true')) return null;

  let best: { head: HTMLTableRowElement[]; tail: HTMLTableRowElement[] } | null = null;
  for (let splitIndex = 1; splitIndex < group.length; splitIndex += 1) {
    const split = splitTableRowGroupAtIndex(group, splitIndex);
    if (!split) continue;
    if (best) removeTableRowGroup(target, best.head);
    for (const row of split.head) target.appendChild(row);
    if (targetBody.scrollHeight <= limit) {
      best = split;
    } else {
      removeTableRowGroup(target, split.head);
      if (best) appendTableRowGroup(target, best.head);
      break;
    }
  }
  if (!best) return null;
  return best.tail;
}

function splitTableRowGroupAtIndex(
  group: HTMLTableRowElement[],
  splitIndex: number,
): { head: HTMLTableRowElement[]; tail: HTMLTableRowElement[] } | null {
  if (splitIndex <= 0 || splitIndex >= group.length) return null;
  const slots = tableRowGroupCellSlots(group);
  const head = group.slice(0, splitIndex).map((row) => row.cloneNode(true) as HTMLTableRowElement);
  const tail = group.slice(splitIndex).map((row) => row.cloneNode(true) as HTMLTableRowElement);
  const continuations = slots.filter((slot) => slot.row < splitIndex && slot.row + slot.rowSpan > splitIndex);
  if (continuations.length === 0) return null;

  for (const slot of continuations) {
    const headRow = head[slot.row];
    const headCell = headRow?.cells[slot.cellIndex];
    if (headCell) headCell.rowSpan = splitIndex - slot.row;

    const continuation = slot.cell.cloneNode(false) as HTMLTableCellElement;
    continuation.rowSpan = slot.row + slot.rowSpan - splitIndex;
    continuation.colSpan = slot.colSpan;
    continuation.dataset.rowFragmentContinuation = 'true';
    continuation.classList.add('is-rowspan-continuation');
    insertTableCellAtColumn(tail[0]!, continuation, slot.column, slots, splitIndex);
  }
  return { head, tail };
}

interface TableCellSlot {
  row: number;
  cellIndex: number;
  column: number;
  colSpan: number;
  rowSpan: number;
  cell: HTMLTableCellElement;
}

function tableRowGroupCellSlots(group: HTMLTableRowElement[]): TableCellSlot[] {
  const slots: TableCellSlot[] = [];
  const activeSpans: number[] = [];
  for (let rowIndex = 0; rowIndex < group.length; rowIndex += 1) {
    const row = group[rowIndex]!;
    let column = 0;
    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      while ((activeSpans[column] ?? 0) > rowIndex) column += 1;
      const cell = row.cells[cellIndex]!;
      const colSpan = Math.max(1, cell.colSpan);
      const rowSpan = Math.max(1, cell.rowSpan);
      slots.push({ row: rowIndex, cellIndex, column, colSpan, rowSpan, cell });
      for (let offset = 0; offset < colSpan; offset += 1) {
        activeSpans[column + offset] = Math.max(activeSpans[column + offset] ?? 0, rowIndex + rowSpan);
      }
      column += colSpan;
    }
  }
  return slots;
}

function insertTableCellAtColumn(
  row: HTMLTableRowElement,
  cell: HTMLTableCellElement,
  column: number,
  slots: TableCellSlot[],
  sourceRowIndex: number,
): void {
  const rowSlots = slots
    .filter((slot) => slot.row === sourceRowIndex)
    .sort((a, b) => a.column - b.column);
  const beforeSlot = rowSlots.find((slot) => slot.column > column);
  const beforeCell = beforeSlot ? row.cells[beforeSlot.cellIndex] : null;
  row.insertBefore(cell, beforeCell ?? null);
}

function splitTableRowGroupByBlocksToFitCurrentPage(
  target: HTMLTableSectionElement,
  group: HTMLTableRowElement[],
  targetBody: HTMLElement,
  limit: number,
): HTMLTableRowElement[] | null {
  if (group.length !== 1) return null;
  const row = group[0]!;
  if (!canSplitTableRowByBlocks(row)) return null;

  let best: { head: HTMLTableRowElement; tail: HTMLTableRowElement } | null = null;
  let low = 0;
  let high = 1;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const ratio = (low + high) / 2;
    const split = splitTableRowByBlockRatio(row, ratio);
    if (!split) {
      high = ratio;
      continue;
    }
    target.appendChild(split.head);
    if (targetBody.scrollHeight <= limit) {
      if (best?.head.parentElement === target) target.removeChild(best.head);
      best = split;
      low = ratio;
    } else {
      target.removeChild(split.head);
      high = ratio;
    }
  }
  if (!best || best.head.parentElement !== target) return null;
  return [best.tail];
}

function splitTableRowGroupByTextToFitCurrentPage(
  target: HTMLTableSectionElement,
  group: HTMLTableRowElement[],
  targetBody: HTMLElement,
  limit: number,
): HTMLTableRowElement[] | null {
  if (group.length !== 1) return null;
  const row = group[0]!;
  if (!canSplitTableRowByText(row)) return null;

  let best: { head: HTMLTableRowElement; tail: HTMLTableRowElement } | null = null;
  let low = 0;
  let high = 1;
  for (let attempt = 0; attempt < 9; attempt += 1) {
    const ratio = (low + high) / 2;
    const split = splitTableRowByTextRatio(row, ratio);
    if (!split) {
      high = ratio;
      continue;
    }
    target.appendChild(split.head);
    if (targetBody.scrollHeight <= limit) {
      if (best?.head.parentElement === target) target.removeChild(best.head);
      best = split;
      low = ratio;
    } else {
      target.removeChild(split.head);
      high = ratio;
    }
  }
  if (!best || best.head.parentElement !== target) return null;
  return [best.tail];
}

function canSplitTableRowByText(row: HTMLTableRowElement): boolean {
  if (row.dataset.repeatHeader === 'true' || row.dataset.cantSplit === 'true') return false;
  if (Array.from(row.cells).some((cell) => cell.rowSpan > 1 || cell.colSpan > 1)) return false;
  const cells = Array.from(row.cells);
  if (cells.some((cell) => tableCellHasHardSplitContent(cell))) return false;
  if (tableRowAllowsStaticCellSplits(row)) return cells.some((cell) => canSplitTableCellByText(cell));
  return cells.every((cell) => canSplitTableCellByText(cell));
}

function splitTableRowByTextRatio(
  row: HTMLTableRowElement,
  ratio: number,
): { head: HTMLTableRowElement; tail: HTMLTableRowElement } | null {
  const head = cloneTableRowForTextSplit(row);
  const tail = cloneTableRowForTextSplit(row);
  for (const cell of Array.from(row.cells)) {
    const split = splitTableCellByTextRatio(cell, ratio);
    if (!split) return null;
    head.appendChild(split.head);
    tail.appendChild(split.tail);
  }
  return { head, tail };
}

function splitTableCellByTextRatio(
  cell: HTMLTableCellElement,
  ratio: number,
): { head: HTMLTableCellElement; tail: HTMLTableCellElement } | null {
  if (!canSplitTableCellByText(cell)) return splitStaticTableCell(cell);
  const paragraph = Array.from(cell.children).find(
    (child): child is HTMLElement => child instanceof HTMLElement && child.tagName.toLowerCase() === 'coco-word-native-p',
  );
  if (!paragraph) return null;
  const textLength = textLengthOfNode(paragraph);
  const offset = Math.max(1, Math.min(textLength - 1, Math.round(textLength * ratio)));
  const split = splitElementByTextOffset(paragraph, offset);
  if (!split || textLengthOfNode(split.head) === 0 || textLengthOfNode(split.tail) === 0) return null;
  const head = cell.cloneNode(false) as HTMLTableCellElement;
  const tail = cell.cloneNode(false) as HTMLTableCellElement;
  head.appendChild(split.head);
  tail.appendChild(split.tail);
  return { head, tail };
}

function splitElementByTextOffset(element: HTMLElement, offset: number): { head: HTMLElement; tail: HTMLElement } | null {
  const total = textLengthOfNode(element);
  if (!Number.isInteger(offset) || offset <= 0 || offset >= total) return null;
  const head = element.cloneNode(false) as HTMLElement;
  const tail = element.cloneNode(false) as HTMLElement;
  splitChildNodesByTextOffset(Array.from(element.childNodes), offset, head, tail);
  head.dataset.splitFragment = 'head';
  tail.dataset.splitFragment = 'tail';
  return { head, tail };
}

function splitChildNodesByTextOffset(
  nodes: Node[],
  offset: number,
  headParent: Node,
  tailParent: Node,
): void {
  let remaining = offset;
  for (const node of nodes) {
    const length = textLengthOfNode(node);
    if (length === 0) {
      if (remaining > 0) headParent.appendChild(node.cloneNode(true));
      else tailParent.appendChild(node.cloneNode(true));
      continue;
    }
    if (length <= remaining) {
      headParent.appendChild(node.cloneNode(true));
      remaining -= length;
      continue;
    }
    if (remaining <= 0) {
      tailParent.appendChild(node.cloneNode(true));
      continue;
    }
    splitNodeAcrossParents(node, remaining, headParent, tailParent);
    remaining = 0;
  }
}

function splitNodeAcrossParents(node: Node, offset: number, headParent: Node, tailParent: Node): void {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? '';
    headParent.appendChild(document.createTextNode(text.slice(0, offset)));
    tailParent.appendChild(document.createTextNode(text.slice(offset)));
    return;
  }
  if (!(node instanceof HTMLElement)) return;
  const head = node.cloneNode(false) as HTMLElement;
  const tail = node.cloneNode(false) as HTMLElement;
  splitChildNodesByTextOffset(Array.from(node.childNodes), offset, head, tail);
  if (textLengthOfNode(head) > 0) headParent.appendChild(head);
  if (textLengthOfNode(tail) > 0) tailParent.appendChild(tail);
}

function textLengthOfNode(node: Node): number {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent?.length ?? 0;
  return Array.from(node.childNodes).reduce((length, child) => length + textLengthOfNode(child), 0);
}

function canSplitTableRowByBlocks(row: HTMLTableRowElement): boolean {
  if (row.dataset.repeatHeader === 'true' || row.dataset.cantSplit === 'true') return false;
  if (Array.from(row.cells).some((cell) => cell.rowSpan > 1 || cell.colSpan > 1)) return false;
  const cells = Array.from(row.cells);
  if (cells.some((cell) => tableCellHasHardSplitContent(cell))) return false;
  if (tableRowAllowsStaticCellSplits(row)) return cells.some((cell) => canSplitTableCellByBlocks(cell));
  return cells.every((cell) => canSplitTableCellByBlocks(cell));
}

function splitTableRowByBlockRatio(
  row: HTMLTableRowElement,
  ratio: number,
): { head: HTMLTableRowElement; tail: HTMLTableRowElement } | null {
  const head = cloneTableRowForBlockSplit(row);
  const tail = cloneTableRowForBlockSplit(row);
  for (const cell of Array.from(row.cells)) {
    const split = splitTableCellByBlockRatio(cell, ratio);
    if (!split) return null;
    head.appendChild(split.head);
    tail.appendChild(split.tail);
  }
  return { head, tail };
}

function splitTableCellByBlockRatio(
  cell: HTMLTableCellElement,
  ratio: number,
): { head: HTMLTableCellElement; tail: HTMLTableCellElement } | null {
  if (!canSplitTableCellByBlocks(cell)) return splitStaticTableCell(cell);
  const blocks = Array.from(cell.childNodes).filter((node): node is HTMLElement => node instanceof HTMLElement);
  if (blocks.length < 2) return null;
  const pivot = Math.max(1, Math.min(blocks.length - 1, Math.round(blocks.length * ratio)));
  const head = cell.cloneNode(false) as HTMLTableCellElement;
  const tail = cell.cloneNode(false) as HTMLTableCellElement;
  for (let index = 0; index < blocks.length; index += 1) {
    const clone = blocks[index]!.cloneNode(true);
    if (index < pivot) head.appendChild(clone);
    else tail.appendChild(clone);
  }
  return { head, tail };
}

function tableCellHasHardSplitContent(cell: HTMLTableCellElement): boolean {
  return (
    cell.querySelector(
      'coco-word-native-drawing, coco-word-native-table, .coco-word-native-field, .coco-word-page-break, .coco-word-comment, .coco-word-revision-delete, .coco-word-revision-insert',
    ) !== null
  );
}

function tableRowAllowsStaticCellSplits(row: HTMLTableRowElement): boolean {
  return row.closest('coco-word-native-document')?.classList.contains('uses-natural-paragraph-flow') === true;
}

function canSplitTableCellByBlocks(cell: HTMLTableCellElement): boolean {
  return Array.from(cell.childNodes).filter((node) => node instanceof HTMLElement).length >= 2;
}

function canSplitTableCellByText(cell: HTMLTableCellElement): boolean {
  if (cell.querySelector('img')) return false;
  const blocks = Array.from(cell.children).filter((child): child is HTMLElement => child instanceof HTMLElement);
  return blocks.length === 1 && blocks[0]!.tagName.toLowerCase() === 'coco-word-native-p' && textLengthOfNode(blocks[0]!) > 1;
}

function splitStaticTableCell(cell: HTMLTableCellElement): { head: HTMLTableCellElement; tail: HTMLTableCellElement } {
  const head = cell.cloneNode(true) as HTMLTableCellElement;
  const tail = cell.cloneNode(false) as HTMLTableCellElement;
  tail.dataset.rowFragmentStatic = 'true';
  return { head, tail };
}

function cloneTableRowForBlockSplit(row: HTMLTableRowElement): HTMLTableRowElement {
  const clone = row.cloneNode(false) as HTMLTableRowElement;
  clone.dataset.rowFragment = 'block';
  clone.style.removeProperty('height');
  clone.style.removeProperty('--coco-word-row-height');
  delete clone.dataset.heightRule;
  return clone;
}

function cloneTableRowForTextSplit(row: HTMLTableRowElement): HTMLTableRowElement {
  const clone = row.cloneNode(false) as HTMLTableRowElement;
  clone.dataset.rowFragment = 'text';
  clone.style.removeProperty('height');
  clone.style.removeProperty('--coco-word-row-height');
  delete clone.dataset.heightRule;
  return clone;
}

function isShortFlexFilledTablePage(body: HTMLElement, limit: number): boolean {
  if (body.scrollHeight > limit + 1) return false;
  return nativeBodyContentHeight(body) <= limit * 0.75;
}

function nativeBodyContentHeight(body: HTMLElement): number {
  let height = 0;
  for (const child of Array.from(body.children)) {
    if (!(child instanceof HTMLElement)) continue;
    const rect = child.getBoundingClientRect();
    const style = window.getComputedStyle(child);
    const marginTop = Number.parseFloat(style.marginTop) || 0;
    const marginBottom = Number.parseFloat(style.marginBottom) || 0;
    height += rect.height + marginTop + marginBottom;
  }
  return height;
}

function nativeBodyContentBottom(body: HTMLElement): number {
  const bodyRect = body.getBoundingClientRect();
  let bottom = 0;
  for (const child of Array.from(body.children)) {
    if (!(child instanceof HTMLElement)) continue;
    const rect = child.getBoundingClientRect();
    const style = window.getComputedStyle(child);
    const marginBottom = Number.parseFloat(style.marginBottom) || 0;
    bottom = Math.max(bottom, rect.bottom - bodyRect.top + marginBottom);
  }
  return bottom;
}

function tableBottomReservePx(block: HTMLElement, bodyLimit: number): number {
  if (bodyLimit < 700) return 0;
  const rowCount = block.querySelectorAll('tr').length;
  const minRowsForReserve = block.closest('coco-word-native-document.uses-compact-latin-tables') ? 8 : 4;
  if (rowCount < minRowsForReserve) return 0;
  return 96;
}

function groupTableRowsForPagination(rows: HTMLTableRowElement[]): HTMLTableRowElement[][] {
  const groups: HTMLTableRowElement[][] = [];
  let start = 0;
  let end = 0;
  for (let index = 0; index < rows.length; index++) {
    end = Math.max(end, index);
    for (const cell of Array.from(rows[index]!.cells)) {
      end = Math.max(end, index + Math.max(1, cell.rowSpan) - 1);
    }
    if (index !== end) continue;
    groups.push(rows.slice(start, index + 1));
    start = index + 1;
  }
  if (start < rows.length) groups.push(rows.slice(start));
  return groups;
}

function appendTableRowGroup(target: HTMLTableSectionElement, group: HTMLTableRowElement[]): void {
  for (const row of group) target.appendChild(row);
  const table = target.closest<HTMLElement>('coco-word-native-table');
  if (table) markNativeTableEditable(table);
}

function tableRowGroupStartsWithHeader(group: HTMLTableRowElement[]): boolean {
  return group[0]?.dataset.repeatHeader === 'true';
}

function appendClonedTableHeaderRows(target: HTMLTableSectionElement, headerRows: HTMLTableRowElement[]): void {
  for (const row of headerRows) {
    const clone = row.cloneNode(true) as HTMLTableRowElement;
    clone.dataset.repeatedHeader = 'true';
    target.appendChild(clone);
  }
  const table = target.closest<HTMLElement>('coco-word-native-table');
  if (table) markNativeTableEditable(table);
}

function removeTableRowGroup(target: HTMLTableSectionElement, group: HTMLTableRowElement[]): void {
  for (const row of group) {
    if (row.parentElement === target) target.removeChild(row);
  }
}

function createNativeTableChunk(sourceBlock: HTMLElement, sourceTable: Element): HTMLElement {
  const block = sourceBlock.cloneNode(false) as HTMLElement;
  block.classList.add('is-table-fragment');
  const table = sourceTable.cloneNode(false) as HTMLElement;
  const colgroup = sourceTable.querySelector(':scope > colgroup');
  if (colgroup) table.appendChild(colgroup.cloneNode(true));
  table.appendChild(document.createElement('tbody'));
  block.appendChild(table);
  return block;
}

function markNativeTableEditable(table: HTMLElement): void {
  table.contentEditable = 'true';
  for (const cell of Array.from(table.querySelectorAll<HTMLTableCellElement>('td, th'))) {
    cell.contentEditable = 'true';
  }
  for (const paragraph of Array.from(table.querySelectorAll<HTMLElement>('coco-word-native-p'))) {
    paragraph.contentEditable = 'true';
    if (paragraph.childNodes.length === 0) paragraph.appendChild(document.createTextNode(''));
  }
}

function compactTrailingEmptyFlowPages(root: HTMLElement): void {
  while (true) {
    const pages = Array.from(root.querySelectorAll<HTMLElement>(':scope > coco-word-native-page'));
    if (pages.length < 2) return;
    const page = pages[pages.length - 1]!;
    if (!isEmptyFlowAnchorPage(page)) return;
    const previousPage = pages[pages.length - 2]!;
    const previousBody = previousPage.querySelector<HTMLElement>(':scope > coco-word-native-page-body');
    const body = page.querySelector<HTMLElement>(':scope > coco-word-native-page-body');
    const overlay = page.querySelector<HTMLElement>(':scope > .coco-word-native-page-overlay');
    const behindLayer = page.querySelector<HTMLElement>(':scope > .coco-word-native-page-behind-layer');
    if (!previousBody || !body || (!overlay && !behindLayer)) return;
    if (behindLayer) {
      const previousBehindLayer = ensureNativePageLayer(previousPage, 'behind');
      for (const anchor of Array.from(behindLayer.children)) previousBehindLayer.appendChild(anchor);
    }
    if (overlay) {
      const previousOverlay = ensureNativePageLayer(previousPage, 'front');
      for (const anchor of Array.from(overlay.children)) previousOverlay.appendChild(anchor);
    }
    for (const block of Array.from(body.children)) previousBody.appendChild(block);
    page.remove();
  }
}

function compactEmptyNativePages(root: HTMLElement): void {
  for (const page of Array.from(root.querySelectorAll<HTMLElement>('coco-word-native-page'))) {
    if (isEmptyNativePage(page)) page.remove();
  }
}

function isEmptyNativePage(page: HTMLElement): boolean {
  const body = page.querySelector<HTMLElement>(':scope > coco-word-native-page-body');
  if (!body || body.children.length > 0) return false;
  const pageLayerAnchorCount =
    (page.querySelector<HTMLElement>(':scope > .coco-word-native-page-overlay')?.children.length ?? 0) +
    (page.querySelector<HTMLElement>(':scope > .coco-word-native-page-behind-layer')?.children.length ?? 0);
  return pageLayerAnchorCount === 0;
}

function renumberNativePages(root: HTMLElement): void {
  const sectionPageCounts = new Map<number, number>();
  const pages = Array.from(root.querySelectorAll<HTMLElement>(':scope > coco-word-native-page'));
  pages.forEach((page, index) => {
    const sectionIndex = parseNonNegativeInt(page.dataset.sectionIndex, 0);
    const sectionPageNumber = (sectionPageCounts.get(sectionIndex) ?? 0) + 1;
    sectionPageCounts.set(sectionIndex, sectionPageNumber);
    page.dataset.pageNumber = String(index + 1);
    page.dataset.sectionPageNumber = String(sectionPageNumber);
  });
}

function isEmptyFlowAnchorPage(page: HTMLElement): boolean {
  const body = page.querySelector<HTMLElement>(':scope > coco-word-native-page-body');
  const overlay = page.querySelector<HTMLElement>(':scope > .coco-word-native-page-overlay');
  const behindLayer = page.querySelector<HTMLElement>(':scope > .coco-word-native-page-behind-layer');
  const pageLayerAnchorCount = (overlay?.children.length ?? 0) + (behindLayer?.children.length ?? 0);
  if (!body || pageLayerAnchorCount === 0) return false;
  const blocks = Array.from(body.children).filter((child): child is HTMLElement => child instanceof HTMLElement);
  if (blocks.length === 0) return false;
  return blocks.every(
    (block) =>
      block.matches('coco-word-native-p.is-flow-empty') &&
      (block.textContent ?? '').trim().length === 0 &&
      block.getBoundingClientRect().height <= 1,
  );
}

function isTableRow(node: Element): node is HTMLTableRowElement {
  return node instanceof HTMLTableRowElement;
}

function promotePageRelativeAnchors(page: HTMLElement): void {
  const sourceParagraphs = new Set<HTMLElement>();
  const anchors = Array.from(
    page.querySelectorAll<HTMLElement>(
      ':scope .is-absolute-anchor:not([data-page-layer-anchor="true"])',
    ),
  );
  for (const anchor of anchors) {
    if (!shouldPromoteAnchorToPageLayer(anchor)) continue;
    const sourceParagraph = anchor.closest<HTMLElement>('coco-word-native-p');
    if (sourceParagraph) sourceParagraphs.add(sourceParagraph);
    const sourceStory = anchor.closest<HTMLElement>('coco-word-native-story');
    if (sourceStory?.dataset.storyKind) anchor.dataset.sourceStoryKind = sourceStory.dataset.storyKind;
    if (sourceStory?.dataset.partName) anchor.dataset.sourcePartName = sourceStory.dataset.partName;
    anchor.dataset.pageLayerAnchor = 'true';
    const layerKind = anchor.dataset.wrap === 'behind' ? 'behind' : 'front';
    anchor.dataset.pageLayer = layerKind;
    anchor.classList.add('is-page-layer-anchor');
    ensureNativePageLayer(page, layerKind).appendChild(anchor);
  }
  for (const paragraph of sourceParagraphs) {
    paragraph.classList.toggle('is-flow-empty', !nativeParagraphHasFlowContent(paragraph));
  }
  cleanupEmptyNativePageLayers(page);
}

function ensureNativePageLayer(page: HTMLElement, layer: 'behind' | 'front'): HTMLElement {
  const className = layer === 'behind' ? 'coco-word-native-page-behind-layer' : 'coco-word-native-page-overlay';
  const existing = page.querySelector<HTMLElement>(`:scope > .${className}`);
  if (existing) return existing;
  const el = document.createElement('div');
  el.className = className;
  el.contentEditable = 'false';
  if (layer === 'behind') {
    const firstFlowChild = Array.from(page.children).find((child) => child.tagName.toLowerCase() !== 'coco-word-native-story');
    page.insertBefore(el, firstFlowChild ?? page.firstChild);
  } else {
    page.appendChild(el);
  }
  return el;
}

function cleanupEmptyNativePageLayers(page: HTMLElement): void {
  page.querySelector<HTMLElement>(':scope > .coco-word-native-page-behind-layer:empty')?.remove();
  page.querySelector<HTMLElement>(':scope > .coco-word-native-page-overlay:empty')?.remove();
}

function shouldPromoteAnchorToPageLayer(anchor: HTMLElement): boolean {
  const vertical = anchor.dataset.positionVerticalRelativeFrom;
  const horizontal = anchor.dataset.positionHorizontalRelativeFrom;
  if (vertical) return isPageLayerRelativeFrom(vertical);
  return isPageLayerRelativeFrom(horizontal);
}

function isPageLayerRelativeFrom(value: string | undefined): boolean {
  return value === 'page' || value === 'margin';
}

function nativeBlockPromotesToEmptyFlow(block: HTMLElement): boolean {
  if (!block.matches('coco-word-native-p')) return false;
  if ((block.textContent ?? '').trim().length > 0) return false;
  const children = Array.from(block.children);
  if (children.length === 0) return false;
  return children.every(
    (child) =>
      child instanceof HTMLElement &&
      child.classList.contains('is-absolute-anchor') &&
      shouldPromoteAnchorToPageLayer(child),
  );
}

function nativeParagraphHasFlowContent(paragraph: HTMLElement): boolean {
  if ((paragraph.textContent ?? '').trim().length > 0) return true;
  for (const child of Array.from(paragraph.children)) {
    if (!(child instanceof HTMLElement)) continue;
    if (child.dataset.pageLayerAnchor === 'true' || child.classList.contains('is-page-layer-anchor')) {
      continue;
    }
    if (child.classList.contains('coco-word-tab-marker')) continue;
    return true;
  }
  return false;
}

function appendNativePage(
  root: HTMLElement,
  doc: NativeWordDocument,
  layout: NonNullable<NativeWordDocument['pageLayout']>,
  headerTemplates: HTMLElement[],
  footerTemplates: HTMLElement[],
  pageNumber: number,
  sectionIndex: number,
  sectionPageNumber: number,
): HTMLElement {
  const page = document.createElement('coco-word-native-page');
  page.dataset.pageNumber = String(pageNumber);
  page.dataset.sectionIndex = String(sectionIndex);
  page.dataset.sectionPageNumber = String(sectionPageNumber);
  const displayPageNumber = displayPageNumberForSection(doc, sectionIndex, pageNumber, sectionPageNumber);
  page.dataset.displayPageNumber = String(displayPageNumber);
  applyNativePageLayout(page, layout);
  for (const header of selectHeaderFooterTemplates(doc, 'header', pageNumber, sectionIndex, sectionPageNumber, headerTemplates)) {
    page.appendChild(cloneNativeHeaderFooterForPage(header, displayPageNumber, sectionPageNumber));
  }
  page.appendChild(document.createElement('coco-word-native-page-body'));
  for (const footer of selectHeaderFooterTemplates(doc, 'footer', pageNumber, sectionIndex, sectionPageNumber, footerTemplates)) {
    page.appendChild(cloneNativeHeaderFooterForPage(footer, displayPageNumber, sectionPageNumber));
  }
  promotePageRelativeAnchors(page);
  root.appendChild(page);
  return page;
}

function refreshNativePageChrome(
  page: HTMLElement,
  doc: NativeWordDocument,
  layout: NonNullable<NativeWordDocument['pageLayout']>,
  headerTemplates: HTMLElement[],
  footerTemplates: HTMLElement[],
  pageNumber: number,
  sectionIndex: number,
  sectionPageNumber: number,
): void {
  const body = page.querySelector<HTMLElement>('coco-word-native-page-body');
  if (!body) return;
  page.dataset.sectionIndex = String(sectionIndex);
  page.dataset.sectionPageNumber = String(sectionPageNumber);
  const displayPageNumber = displayPageNumberForSection(doc, sectionIndex, pageNumber, sectionPageNumber);
  page.dataset.displayPageNumber = String(displayPageNumber);
  applyNativePageLayout(page, layout);
  page.querySelector<HTMLElement>(':scope > .coco-word-native-page-overlay')?.remove();
  page.querySelector<HTMLElement>(':scope > .coco-word-native-page-behind-layer')?.remove();
  for (const story of Array.from(page.querySelectorAll(':scope > coco-word-native-story'))) {
    page.removeChild(story);
  }
  for (const header of selectHeaderFooterTemplates(doc, 'header', pageNumber, sectionIndex, sectionPageNumber, headerTemplates)) {
    page.insertBefore(cloneNativeHeaderFooterForPage(header, displayPageNumber, sectionPageNumber), body);
  }
  for (const footer of selectHeaderFooterTemplates(doc, 'footer', pageNumber, sectionIndex, sectionPageNumber, footerTemplates)) {
    page.appendChild(cloneNativeHeaderFooterForPage(footer, displayPageNumber, sectionPageNumber));
  }
  promotePageRelativeAnchors(page);
}

export function refreshNativePageCommentBubbles(root: HTMLElement): void {
  for (const page of Array.from(root.querySelectorAll<HTMLElement>(':scope > coco-word-native-page'))) {
    page.querySelector<HTMLElement>(':scope > .coco-word-native-comment-layer')?.remove();
    const comments = uniquePageComments(page);
    if (comments.length > 0) {
      const layer = document.createElement('div');
      layer.className = 'coco-word-native-comment-layer';
      layer.contentEditable = 'false';
      page.appendChild(layer);
      comments.sort((a, b) => a.top - b.top || a.commentId - b.commentId).forEach((comment, index) => {
        const bubble = document.createElement('aside');
        bubble.className = 'coco-word-native-comment-bubble';
        bubble.dataset.annotationAnchorTop = String(comment.top);
        bubble.dataset.commentId = String(comment.commentId);
        bubble.dataset.commentBody = comment.body;
        if (comment.parentCommentId !== undefined && Number.isFinite(comment.parentCommentId)) {
          bubble.dataset.parentCommentId = String(comment.parentCommentId);
          bubble.classList.add('coco-word-native-comment-reply');
        }
        if (comment.paraId) bubble.dataset.paraId = comment.paraId;
        if (comment.from !== undefined) bubble.dataset.from = String(comment.from);
        if (comment.to !== undefined) bubble.dataset.to = String(comment.to);
        bubble.style.setProperty('--coco-word-comment-index', String(index));

        const meta = document.createElement('div');
        meta.className = 'coco-word-native-comment-meta';
        meta.textContent = [comment.author || `Comment ${comment.commentId}`, comment.date ? formatRevisionDate(comment.date) : '']
          .filter(Boolean)
          .join(' · ');
        if (comment.parentCommentId !== undefined && Number.isFinite(comment.parentCommentId)) {
          meta.textContent = `回复 #${comment.parentCommentId} · ${meta.textContent}`;
        }

        const body = document.createElement('div');
        body.className = 'coco-word-native-comment-body';
        body.textContent = comment.body || '(empty comment)';

        const actions = document.createElement('div');
        actions.className = 'coco-word-native-comment-actions';
        const menuButton = document.createElement('button');
        menuButton.type = 'button';
        menuButton.className = 'coco-word-native-comment-menu-button';
        menuButton.dataset.commentMenu = 'true';
        menuButton.dataset.commentId = String(comment.commentId);
        menuButton.setAttribute('aria-label', '批注操作');
        menuButton.title = '批注操作';
        menuButton.textContent = '⋯';
        actions.appendChild(menuButton);

        const menu = document.createElement('div');
        menu.className = 'coco-word-native-comment-menu';
        menu.hidden = true;
        menu.dataset.commentMenuFor = String(comment.commentId);
        for (const [action, text] of [['reply', '回复'], ['edit', '编辑'], ['delete', '删除']] as const) {
          const button = document.createElement('button');
          button.type = 'button';
          button.className = 'coco-word-native-comment-menu-item';
          button.dataset.commentAction = action;
          button.dataset.commentId = String(comment.commentId);
          button.textContent = text;
          menu.appendChild(button);
        }
        actions.appendChild(menu);

        bubble.appendChild(meta);
        bubble.appendChild(body);
        bubble.appendChild(actions);
        layer.appendChild(bubble);
      });
    }
    refreshNativePageRevisionBubbles(page);
    layoutNativePageAnnotationBubbles(page);
  }
}

function refreshNativePageRevisionBubbles(page: HTMLElement): void {
  page.querySelector<HTMLElement>(':scope > .coco-word-native-revision-layer')?.remove();
  const revisions = new Map<string, { id: string; kind: 'insert' | 'delete' | 'format'; author: string; date: string; text: string; top: number; paraId?: string; from?: number; to?: number }>();
  const pageRect = page.getBoundingClientRect();
  for (const mark of Array.from(page.querySelectorAll<HTMLElement>('.coco-word-revision-insert[data-revision-id], .coco-word-revision-delete[data-revision-id], .coco-word-revision-format[data-revision-id], img[data-revision-id], coco-word-native-table[data-revision-id]'))) {
    const id = mark.dataset.revisionId ?? '';
    if (!id) continue;
    const kind = mark.dataset.revisionKind === 'format'
      ? 'format'
      : mark.dataset.revisionKind === 'delete' || mark.classList.contains('coco-word-revision-delete') ? 'delete' : 'insert';
    const rect = mark.getBoundingClientRect();
    const key = `${kind}:${id}`;
    const previous = revisions.get(key);
    const sourceRange = nativeMarkRange(mark);
    revisions.set(key, {
      id,
      kind,
      author: mark.dataset.revisionAuthor ?? '',
      date: mark.dataset.revisionDate ?? '',
      text: `${previous?.text ?? ''}${mark.dataset.revisionText ?? mark.textContent ?? ''}`,
      top: previous?.top ?? Math.max(0, rect.top - pageRect.top),
      paraId: previous?.paraId ?? sourceRange?.paraId,
      from: previous?.from ?? sourceRange?.from,
      to: sourceRange?.to ?? previous?.to,
    });
  }
  if (revisions.size === 0) return;
  const layer = document.createElement('div');
  layer.className = 'coco-word-native-revision-layer';
  layer.contentEditable = 'false';
  page.appendChild(layer);
  const sortedRevisions = Array.from(revisions.values()).sort((a, b) => a.top - b.top || a.id.localeCompare(b.id));
  const groups = new Map<string, typeof sortedRevisions>();
  for (const revision of sortedRevisions) {
    const key = revision.paraId ? `paragraph:${revision.paraId}` : `position:${Math.round(revision.top / 48)}`;
    const group = groups.get(key);
    if (group) group.push(revision);
    else groups.set(key, [revision]);
  }

  const createBubble = (revision: (typeof sortedRevisions)[number]) => {
    const bubble = document.createElement('aside');
    bubble.className = `coco-word-native-revision-bubble is-${revision.kind}`;
    bubble.dataset.annotationAnchorTop = String(revision.top);
    bubble.dataset.revisionId = revision.id;
    bubble.dataset.revisionKind = revision.kind;
    if (revision.paraId) bubble.dataset.paraId = revision.paraId;
    if (revision.from !== undefined) bubble.dataset.from = String(revision.from);
    if (revision.to !== undefined) bubble.dataset.to = String(revision.to);
    const header = document.createElement('div');
    header.className = 'coco-word-native-revision-header';
    const meta = document.createElement('span');
    meta.className = 'coco-word-native-revision-meta';
    meta.textContent = [revision.author || 'User', formatRevisionDate(revision.date)].filter(Boolean).join(' · ');
    header.appendChild(meta);
    bubble.appendChild(header);
    const content = document.createElement('div');
    content.className = 'coco-word-native-revision-content';
    content.textContent = revision.text || '(空内容)';
    bubble.appendChild(content);
    const actions = document.createElement('div');
    actions.className = 'coco-word-native-revision-actions';
    for (const [action, text, label] of [['accept', '✓', '接受修订'], ['reject', '×', '拒绝修订']] as const) {
      const button = document.createElement('button');
      button.type = 'button';
      button.dataset.revisionAction = action;
      button.dataset.revisionId = revision.id;
      button.dataset.revisionKind = revision.kind;
      button.textContent = text;
      button.setAttribute('aria-label', label);
      button.title = label;
      button.className = `is-${action}`;
      actions.appendChild(button);
    }
    bubble.appendChild(actions);
    return bubble;
  };

  for (const revisionsInGroup of groups.values()) {
    if (revisionsInGroup.length === 1) {
      layer.appendChild(createBubble(revisionsInGroup[0]!));
      continue;
    }
    const group = document.createElement('section');
    group.className = 'coco-word-native-revision-group';
    group.dataset.annotationAnchorTop = String(revisionsInGroup[0]!.top);
    const paraId = revisionsInGroup[0]!.paraId;
    if (paraId) group.dataset.paraId = paraId;

    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'coco-word-native-revision-group-toggle';
    toggle.dataset.revisionGroupToggle = 'true';
    toggle.setAttribute('aria-expanded', 'false');
    const title = document.createElement('strong');
    title.textContent = `本段 ${revisionsInGroup.length} 条修订`;
    const counts = revisionsInGroup.reduce((result, revision) => {
      result[revision.kind] += 1;
      return result;
    }, { insert: 0, delete: 0, format: 0 });
    const summary = document.createElement('span');
    summary.textContent = [
      counts.delete ? `${counts.delete} 删` : '',
      counts.insert ? `${counts.insert} 增` : '',
      counts.format ? `${counts.format} 格式` : '',
    ].filter(Boolean).join(' · ');
    const affordance = document.createElement('span');
    affordance.className = 'coco-word-native-revision-group-affordance';
    affordance.textContent = '展开';
    toggle.append(title, summary, affordance);

    const list = document.createElement('div');
    list.className = 'coco-word-native-revision-group-list';
    list.hidden = true;
    revisionsInGroup.forEach((revision) => list.appendChild(createBubble(revision)));
    group.append(toggle, list);
    layer.appendChild(group);
  }
}

interface NativeAnnotationBubbleLayoutEntry {
  element: HTMLElement;
  anchorTop: number;
  height: number;
  top: number;
}

/** Keep comments and revisions in one ordered, non-overlapping page rail. */
export function layoutNativePageAnnotationBubbles(page: HTMLElement): void {
  const gap = 8;
  const entries: NativeAnnotationBubbleLayoutEntry[] = Array.from(
    page.querySelectorAll<HTMLElement>(
      ':scope > .coco-word-native-comment-layer > .coco-word-native-comment-bubble, :scope > .coco-word-native-revision-layer > .coco-word-native-revision-bubble, :scope > .coco-word-native-revision-layer > .coco-word-native-revision-group',
    ),
  ).map((element) => ({
    element,
    anchorTop: Math.max(0, Number(element.dataset.annotationAnchorTop) || 0),
    height: Math.max(44, Math.ceil(element.getBoundingClientRect().height)),
    top: 0,
  }));

  if (entries.length === 0) {
    return;
  }

  entries.sort((a, b) => a.anchorTop - b.anchorTop);
  let previousBottom = 0;
  entries.forEach((entry, index) => {
    entry.top = Math.max(entry.anchorTop, index === 0 ? 0 : previousBottom + gap);
    previousBottom = entry.top + entry.height;
  });

  const pageHeight = Math.max(0, Math.ceil(page.getBoundingClientRect().height));
  if (pageHeight > 0 && previousBottom > pageHeight) {
    let nextTop = pageHeight;
    for (let index = entries.length - 1; index >= 0; index -= 1) {
      const entry = entries[index]!;
      entry.top = Math.min(entry.top, nextTop - entry.height);
      nextTop = entry.top - gap;
    }
    if (entries[0]!.top < 0) {
      const shift = -entries[0]!.top;
      entries.forEach((entry) => {
        entry.top += shift;
      });
    }
  }

  for (const entry of entries) {
    const property = entry.element.classList.contains('coco-word-native-comment-bubble')
      ? '--coco-word-comment-top'
      : '--coco-word-revision-top';
    entry.element.style.setProperty(property, `${Math.round(entry.top)}px`);
  }
}

function formatRevisionDate(value: string): string {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString('zh-CN', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
}

function uniquePageComments(page: HTMLElement): Array<{ commentId: number; author: string; date: string; body: string; top: number; paraId?: string; from?: number; to?: number; parentCommentId?: number }> {
  const pageRect = page.getBoundingClientRect();
  const comments = new Map<number, { commentId: number; author: string; date: string; body: string; top: number; paraId?: string; from?: number; to?: number; parentCommentId?: number }>();
  for (const mark of Array.from(page.querySelectorAll<HTMLElement>('.coco-word-comment[data-comment-id]'))) {
    const commentId = Number(mark.dataset.commentId);
    if (!Number.isFinite(commentId)) continue;
    const markRect = mark.getBoundingClientRect();
    const sourceRange = nativeMarkRange(mark);
    const previous = comments.get(commentId);
    const sameParagraph = !previous?.paraId || !sourceRange?.paraId || previous.paraId === sourceRange.paraId;
    comments.set(commentId, {
      commentId,
      author: previous?.author ?? mark.dataset.commentAuthor ?? '',
      date: previous?.date ?? mark.dataset.commentDate ?? '',
      body: previous?.body ?? mark.dataset.commentBody ?? '',
      top: previous?.top ?? Math.max(0, markRect.top - pageRect.top),
      parentCommentId: previous?.parentCommentId ?? (mark.dataset.parentCommentId ? Number(mark.dataset.parentCommentId) : undefined),
      paraId: previous?.paraId ?? sourceRange?.paraId,
      from: sameParagraph
        ? Math.min(previous?.from ?? sourceRange?.from ?? Number.POSITIVE_INFINITY, sourceRange?.from ?? Number.POSITIVE_INFINITY)
        : previous?.from ?? sourceRange?.from,
      to: sameParagraph
        ? Math.max(previous?.to ?? sourceRange?.to ?? Number.NEGATIVE_INFINITY, sourceRange?.to ?? Number.NEGATIVE_INFINITY)
        : previous?.to ?? sourceRange?.to,
    });
  }
  return Array.from(comments.values()).map((comment) => ({
    ...comment,
    ...(Number.isFinite(comment.from) ? { from: comment.from } : { from: undefined }),
    ...(Number.isFinite(comment.to) ? { to: comment.to } : { to: undefined }),
  })).sort((a, b) => a.top - b.top || a.commentId - b.commentId);
}

function nativeMarkRange(mark: HTMLElement): { paraId: string; from: number; to: number } | null {
  const paragraph = mark.closest<HTMLElement>('coco-word-native-p[data-para-id]');
  if (!paragraph) return null;
  let offset = 0;
  let found = false;
  const walk = (node: Node): void => {
    if (found) return;
    if (node === mark) {
      found = true;
      return;
    }
    if (node.nodeType === Node.TEXT_NODE) {
      offset += node.textContent?.length ?? 0;
      return;
    }
    if (node instanceof HTMLElement && node.classList.contains('coco-word-comment-widget')) return;
    for (const child of Array.from(node.childNodes)) walk(child);
  };
  walk(paragraph);
  if (!found) return null;
  const length = Array.from(mark.childNodes).reduce((total, child) => total + nativeInlineDomLength(child), 0);
  return { paraId: paragraph.dataset.paraId ?? '', from: offset, to: offset + Math.max(1, length) };
}

function nativeInlineDomLength(node: Node): number {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent?.length ?? 0;
  if (node instanceof HTMLElement && node.classList.contains('coco-word-comment-widget')) return 0;
  if (node instanceof HTMLImageElement || node instanceof HTMLElement && node.matches('coco-word-native-drawing, coco-word-native-table, .coco-word-page-break')) return 1;
  return Array.from(node.childNodes).reduce((total, child) => total + nativeInlineDomLength(child), 0);
}

function pageLayoutForSection(
  doc: NativeWordDocument,
  sectionIndex: number,
  fallback: NonNullable<NativeWordDocument['pageLayout']>,
): NonNullable<NativeWordDocument['pageLayout']> {
  return doc.sections.find((section) => section.sectionIndex === sectionIndex)?.pageLayout ?? fallback;
}

function displayPageNumberForSection(
  doc: NativeWordDocument,
  sectionIndex: number,
  pageNumber: number,
  sectionPageNumber: number,
): number {
  const section = doc.sections.find((item) => item.sectionIndex === sectionIndex);
  if (section?.pageNumberStart !== undefined) return section.pageNumberStart + sectionPageNumber - 1;
  return pageNumber;
}

function applyNativePageLayout(
  page: HTMLElement,
  layout: NonNullable<NativeWordDocument['pageLayout']>,
): void {
  page.style.width = twipsToPx(layout.width);
  page.style.height = twipsToPx(layout.height);
  page.style.paddingTop = twipsToPx(layout.margin.top);
  page.style.paddingRight = twipsToPx(layout.margin.right);
  page.style.paddingBottom = twipsToPx(layout.margin.bottom);
  page.style.paddingLeft = twipsToPx(layout.margin.left);
  page.style.setProperty('--coco-word-page-margin-top', twipsToPx(layout.margin.top));
  page.style.setProperty('--coco-word-page-margin-left', twipsToPx(layout.margin.left));
  page.style.setProperty('--coco-word-page-margin-right', twipsToPx(layout.margin.right));
  page.style.setProperty('--coco-word-page-margin-bottom', twipsToPx(layout.margin.bottom));
  page.style.setProperty('--coco-word-header-distance', twipsToPx(layout.headerDistance ?? Math.round(layout.margin.top / 2)));
  page.style.setProperty('--coco-word-footer-distance', twipsToPx(layout.footerDistance ?? Math.round(layout.margin.bottom / 2)));
  if (layout.docGrid?.linePitch !== undefined && layout.docGrid.linePitch > 0) {
    page.style.setProperty('--coco-word-doc-grid-line-height', twipsToPx(layout.docGrid.linePitch));
  } else {
    page.style.removeProperty('--coco-word-doc-grid-line-height');
  }
}

function contentHeightForLayout(layout: NonNullable<NativeWordDocument['pageLayout']>): number {
  return Math.max(120, twipsToPxNumber(layout.height - layout.margin.top - layout.margin.bottom));
}

function selectHeaderFooterTemplates(
  doc: NativeWordDocument,
  kind: 'header' | 'footer',
  pageNumber: number,
  sectionIndex: number,
  sectionPageNumber: number,
  templates: HTMLElement[],
): HTMLElement[] {
  if (templates.length <= 1) return templates;
  const section = doc.sections.find((item) => item.sectionIndex === sectionIndex) ?? doc.sections[0];
  const refs = effectiveHeaderFooterRefs(doc, kind, section?.sectionIndex ?? sectionIndex);
  const selectedPartName = selectHeaderFooterPartName(
    refs,
    section?.titlePage === true,
    pageNumber,
    sectionPageNumber,
  );
  if (!selectedPartName && !hasHeaderFooterRefs(doc, kind)) return templates;
  if (selectedPartName) {
    const template = templates.find((story) => story.dataset.partName === selectedPartName);
    return template ? [template] : [];
  }
  const visibleTemplate = templates.find((story) => !story.classList.contains('is-empty'));
  return visibleTemplate ? [visibleTemplate] : [templates[0]!];
}

function effectiveHeaderFooterRefs(
  doc: NativeWordDocument,
  kind: 'header' | 'footer',
  sectionIndex: number,
): Array<{ type: 'default' | 'first' | 'even'; partName: string }> {
  for (let index = sectionIndex; index >= 0; index--) {
    const section = doc.sections.find((item) => item.sectionIndex === index);
    const refs = kind === 'header' ? section?.headerRefs : section?.footerRefs;
    if (refs && refs.length > 0) return refs;
  }
  return [];
}

function hasHeaderFooterRefs(doc: NativeWordDocument, kind: 'header' | 'footer'): boolean {
  return doc.sections.some((section) =>
    kind === 'header' ? section.headerRefs.length > 0 : section.footerRefs.length > 0,
  );
}

function selectHeaderFooterPartName(
  refs: Array<{ type: 'default' | 'first' | 'even'; partName: string }>,
  titlePage: boolean,
  pageNumber: number,
  sectionPageNumber: number,
): string | null {
  if (refs.length === 0) return null;
  if (sectionPageNumber === 1 && titlePage) {
    const first = refs.find((ref) => ref.type === 'first');
    if (first) return first.partName;
  }
  if (pageNumber % 2 === 0) {
    const even = refs.find((ref) => ref.type === 'even');
    if (even) return even.partName;
  }
  return refs.find((ref) => ref.type === 'default')?.partName ?? refs[0]!.partName;
}

function renderNativeHeaderFooterStory(story: NativeWordHeaderFooterStory): HTMLElement {
  const el = document.createElement('coco-word-native-story');
  el.setAttribute('data-story-kind', story.kind);
  el.setAttribute('data-part-name', story.partName);
  if (!hasVisibleStoryContent(story)) el.classList.add('is-empty');
  for (const block of story.blocks) el.appendChild(renderNativeBlock(block));
  return el;
}

function cloneNativeHeaderFooterForPage(
  template: HTMLElement,
  pageNumber: number,
  sectionPageNumber: number,
): HTMLElement {
  const clone = template.cloneNode(true) as HTMLElement;
  for (const field of Array.from(clone.querySelectorAll<HTMLElement>('[data-field-kind]'))) {
    if (field.dataset.fieldKind === 'pageNumber') field.textContent = String(pageNumber);
    field.dataset.pageNumber = String(pageNumber);
    field.dataset.sectionPageNumber = String(sectionPageNumber);
  }
  return clone;
}

function refreshNativePageFields(root: HTMLElement, doc: NativeWordDocument): void {
  const pages = Array.from(root.querySelectorAll<HTMLElement>(':scope > coco-word-native-page'));
  const sectionPageCounts = new Map<number, number>();
  for (const page of pages) {
    const sectionIndex = Number(page.dataset.sectionIndex ?? '0');
    sectionPageCounts.set(sectionIndex, (sectionPageCounts.get(sectionIndex) ?? 0) + 1);
  }
  const displayPageStarts = displayPageStartsForPages(doc, pages, sectionPageCounts);
  for (const page of pages) {
    const pageNumber = parsePositiveInt(page.dataset.pageNumber, 1);
    const sectionIndex = parseNonNegativeInt(page.dataset.sectionIndex, 0);
    const sectionPageNumber = parsePositiveInt(page.dataset.sectionPageNumber, pageNumber);
    const displayPageNumber = (displayPageStarts.get(sectionIndex) ?? pageNumber) + sectionPageNumber - 1;
    const sectionPages = sectionPageCounts.get(sectionIndex) ?? sectionPageNumber;
    page.dataset.displayPageNumber = String(displayPageNumber);
    for (const field of Array.from(page.querySelectorAll<HTMLElement>('[data-field-kind]'))) {
      if (field.dataset.fieldKind === 'pageNumber') field.textContent = String(displayPageNumber);
      else if (field.dataset.fieldKind === 'totalPages') field.textContent = String(pages.length);
      else if (field.dataset.fieldKind === 'sectionPages') field.textContent = String(sectionPages);
      field.dataset.pageNumber = String(displayPageNumber);
      field.dataset.sectionPageNumber = String(sectionPageNumber);
      field.dataset.totalPages = String(pages.length);
      field.dataset.sectionPages = String(sectionPages);
    }
  }
}

function displayPageStartsForPages(
  doc: NativeWordDocument,
  pages: HTMLElement[],
  sectionPageCounts: Map<number, number>,
): Map<number, number> {
  const sectionOrder: number[] = [];
  const seen = new Set<number>();
  for (const page of pages) {
    const sectionIndex = parseNonNegativeInt(page.dataset.sectionIndex, 0);
    if (seen.has(sectionIndex)) continue;
    seen.add(sectionIndex);
    sectionOrder.push(sectionIndex);
  }
  return resolveNativeSectionPageNumberStarts(doc.sections, sectionOrder, sectionPageCounts);
}

export function resolveNativeSectionPageNumberStarts(
  sections: Array<{ sectionIndex: number; pageNumberStart?: number }>,
  sectionOrder: number[],
  sectionPageCounts: Map<number, number>,
): Map<number, number> {
  const starts = new Map<number, number>();
  let nextPageNumber = 1;
  for (const sectionIndex of sectionOrder) {
    const section = sections.find((item) => item.sectionIndex === sectionIndex);
    const start = section?.pageNumberStart ?? nextPageNumber;
    starts.set(sectionIndex, start);
    nextPageNumber = start + (sectionPageCounts.get(sectionIndex) ?? 0);
  }
  return starts;
}

function parsePositiveInt(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function parseNonNegativeInt(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : fallback;
}

function hasVisibleStoryContent(story: NativeWordHeaderFooterStory): boolean {
  return story.blocks.some((block) => {
    if (block.kind === 'table') return block.rows.length > 0;
    return block.content.some((child) => {
      if (child.kind === 'text') return child.text.trim().length > 0;
      if (child.kind === 'field') return true;
      return true;
    });
  });
}

export function renderNativeBlock(block: NativeWordBlockNode, options: NativeRenderOptions = {}): HTMLElement {
  return block.kind === 'paragraph' ? renderNativeParagraph(block, options) : renderNativeTable(block);
}

export function renderNativeParagraph(paragraph: NativeWordParagraphNode, options: NativeRenderOptions = {}): HTMLElement {
  const el = document.createElement('coco-word-native-p');
  el.setAttribute('data-para-id', paragraph.paraId);
  el.setAttribute('data-story', paragraph.story);
  if (paragraph.styleId) el.setAttribute('data-style-id', paragraph.styleId);
  if (paragraph.format?.outlineLevel !== undefined) {
    el.setAttribute('data-outline-level', String(paragraph.format.outlineLevel));
  }
  if (paragraph.format?.hasNumbering) {
    el.setAttribute('data-numbering', 'true');
    el.setAttribute('data-numbering-label', paragraph.format.numbering?.label ?? '•');
  }
  if (paragraph.format?.pageBreakBefore) el.setAttribute('data-page-break-before', 'true');
  if (paragraph.format?.hasSectionBreak) el.setAttribute('data-section-break', 'true');
  if (paragraph.format?.keepNext) el.setAttribute('data-keep-next', 'true');
  if (paragraph.format?.keepLines) el.setAttribute('data-keep-lines', 'true');
  if (paragraph.format?.widowControl !== undefined) {
    el.setAttribute('data-widow-control', String(paragraph.format.widowControl));
  }
  if (paragraph.format?.snapToGrid !== undefined) {
    el.setAttribute('data-snap-to-grid', String(paragraph.format.snapToGrid));
  }
  const paragraphFormat = paragraphUsesImplicitGrid(paragraph, options)
    ? {
        ...paragraph.format,
        snapToGrid: true,
        spacing: {
          ...(paragraph.format?.spacing ?? {}),
          line: options.implicitGridLinePitch,
          lineRule: 'exact' as const,
        },
      }
    : paragraph.format;
  const paragraphStyle = resolveParagraphStyle(paragraphFormat).cssText;
  if (paragraphStyle.length > 0) el.setAttribute('style', paragraphStyle);
  el.contentEditable = 'inherit';
  el.tabIndex = 0;
  el.setAttribute('role', 'textbox');
  el.spellcheck = false;
  const tabRow = renderTabStopRow(paragraph);
  if (tabRow) {
    el.classList.add('has-tab-stops');
    el.appendChild(tabRow);
    return el;
  }
  for (const child of paragraph.content) {
    if (child.kind === 'text') {
      el.appendChild(renderNativeText(child));
    } else if (child.kind === 'field') {
      el.appendChild(renderNativeField(child));
    } else if (child.kind === 'image') {
      el.appendChild(renderNativeImage(child));
    } else if (child.kind === 'drawing') {
      el.appendChild(renderNativeDrawing(child));
    } else {
      const pageBreak = document.createElement('span');
      pageBreak.className = 'coco-word-page-break';
      pageBreak.contentEditable = 'false';
      pageBreak.setAttribute('data-kind', 'page-break');
      pageBreak.setAttribute('aria-label', 'Page break');
      el.appendChild(pageBreak);
    }
  }
  // Keep an empty paragraph keyboard-focusable, including newly inserted table cells.
  if (el.childNodes.length === 0) el.appendChild(document.createTextNode(''));
  return el;
}

export function renderNativeImage(child: NativeWordImageNode): HTMLImageElement {
  const img = document.createElement('img');
  const safeSrc = safeEmbeddedRasterImageUrl(child.src);
  if (safeSrc) img.src = safeSrc;
  img.alt = child.drawingId;
  img.dataset.drawingId = child.drawingId;
  if (child.revision) {
    img.classList.add(child.revision.revisionKind === 'delete' ? 'coco-word-revision-delete' : 'coco-word-revision-insert');
    img.dataset.revisionId = child.revision.revisionId;
    img.dataset.revisionKind = child.revision.revisionKind;
    img.dataset.revisionAuthor = child.revision.author ?? '';
    img.dataset.revisionDate = child.revision.date ?? '';
    img.dataset.revisionText = '图片';
  }
  if (child.formatRevision) {
    img.classList.add('coco-word-revision-format');
    img.dataset.revisionId = child.formatRevision.revisionId;
    img.dataset.revisionKind = 'format';
    img.dataset.revisionAuthor = child.formatRevision.author ?? '';
    img.dataset.revisionDate = child.formatRevision.date ?? '';
    img.dataset.revisionText = '图片尺寸';
  }
  if (child.anchoringMode) img.dataset.anchoringMode = child.anchoringMode;
  if (child.wrap) img.dataset.wrap = child.wrap;
  if (child.width) img.style.width = `${child.width}px`;
  if (child.height) img.style.height = `${child.height}px`;
  applyImageLayoutStyle(img, child);
  return img;
}

function paragraphUsesImplicitGrid(paragraph: NativeWordParagraphNode, options: NativeRenderOptions): boolean {
  if (!options.implicitGridForParagraphs) return false;
  if (paragraph.format?.snapToGrid !== undefined) return false;
  // Word's document grid is the default line rhythm when a paragraph does
  // not explicitly opt out or provide its own line spacing.
  return paragraph.format?.spacing?.line === undefined && options.implicitGridLinePitch !== undefined;
}

function nativeBlockHasSnapToGridDisabled(block: NativeWordBlockNode): boolean {
  if (block.kind === 'paragraph') return block.format?.snapToGrid === false;
  return block.rows.some((row) =>
    row.cells.some((cell) => cell.paragraphs.some((paragraph) => paragraph.format?.snapToGrid === false)),
  );
}

function renderTabStopRow(paragraph: NativeWordParagraphNode): HTMLElement | null {
  const tabStops = paragraph.format?.tabStops;
  if (!tabStops || tabStops.length === 0) return null;
  const tabCount = paragraph.content.reduce(
    (count, child) => count + (child.kind === 'text' ? countTabs(child.text) : 0),
    0,
  );
  if (tabCount !== 1 && tabCount !== 2) return null;
  if (!paragraph.content.every((child) => child.kind === 'text')) return null;
  const rightTab = tabStops.find((stop) => stop.alignment === 'right');
  if (!rightTab) return null;
  const hasLeader = rightTab.leader !== undefined && rightTab.leader !== 'none';
  const textParts = paragraphTabTextParts(paragraph);
  if (!hasLeader && !isCompactTabStopLayout(textParts)) return null;

  const row = document.createElement('span');
  row.className = tabCount === 2 ? 'coco-word-tab-row is-three-part-tab' : 'coco-word-tab-row is-right-tab';
  row.style.setProperty('--coco-word-tab-stop', twipsToPt(rightTab.position));
  if (hasLeader) {
    row.classList.add(`has-leader-${rightTab.leader}`);
    row.dataset.tabLeader = rightTab.leader;
  }
  const segments = Array.from({ length: tabCount + 1 }, (_, index) => {
    const segment = document.createElement('span');
    segment.className = `coco-word-tab-segment is-segment-${index}`;
    return segment;
  });
  let targetIndex = 0;

  for (const child of paragraph.content) {
    if (child.kind !== 'text') continue;
    const parts = child.text.split('\t');
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i] ?? '';
      if (part.length > 0) segments[targetIndex]?.appendChild(renderNativeText({ ...child, text: part }));
      if (i < parts.length - 1) targetIndex = Math.min(targetIndex + 1, segments.length - 1);
    }
  }

  for (let index = 0; index < segments.length; index++) {
    if (index > 0) {
      const marker = document.createElement('span');
      marker.className = 'coco-word-tab-marker';
      marker.textContent = '\t';
      row.appendChild(marker);
    }
    row.appendChild(segments[index]!);
  }
  return row;
}

function paragraphTabTextParts(paragraph: NativeWordParagraphNode): string[] {
  const parts = [''];
  for (const child of paragraph.content) {
    if (child.kind !== 'text') continue;
    const chunks = child.text.split('\t');
    parts[parts.length - 1] += chunks[0] ?? '';
    for (let index = 1; index < chunks.length; index++) {
      parts.push(chunks[index] ?? '');
    }
  }
  return parts;
}

function isCompactTabStopLayout(parts: string[]): boolean {
  if (parts.length < 2 || parts.length > 3) return false;
  return parts.every((part) => part.trim().length <= 32);
}

function countTabs(value: string): number {
  return value.split('\t').length - 1;
}

function renderNativeDrawing(child: NativeWordDrawingNode): HTMLElement {
  const drawing = document.createElement('coco-word-native-drawing');
  drawing.contentEditable = 'false';
  drawing.dataset.drawingId = child.drawingId;
  drawing.dataset.drawingKind = child.drawingKind;
  if (child.anchoringMode) drawing.dataset.anchoringMode = child.anchoringMode;
  if (child.wrap) drawing.dataset.wrap = child.wrap;
  drawing.classList.add(`is-${child.drawingKind.toLowerCase()}`);
  if (child.hasBorder) drawing.classList.add('has-border');
  if (child.hasFill) drawing.classList.add('has-fill');
  if (child.isLineOrConnector) {
    drawing.classList.add('is-line-or-connector');
    drawing.classList.add(isVerticalLine(child) ? 'is-vertical-line' : 'is-horizontal-line');
  }
  const hasVisibleText =
    (child.text?.trim().length ?? 0) > 0 ||
    (child.paragraphs?.some((paragraph) =>
      paragraph.content.some((node) => node.kind === 'text' && node.text.trim().length > 0),
    ) ??
      false);
  const isEmptyTextBoxPlaceholder =
    child.drawingKind === 'textBox' && !hasVisibleText && !child.hasBorder;
  if (
    isEmptyTextBoxPlaceholder ||
    (!hasVisibleText && !child.hasBorder && !child.hasFill && !child.isLineOrConnector)
  ) {
    drawing.classList.add('is-empty');
  }
  const width = nativeDrawingWidthCss(child);
  if (width) drawing.style.width = width;
  if (child.isLineOrConnector) {
    if (child.height) drawing.style.height = `${child.height}px`;
  } else if (child.height && shouldApplyDrawingMinHeight(child)) {
    drawing.style.minHeight = `${child.height}px`;
  }
  applyDrawingLayoutStyle(drawing, child);

  if (child.paragraphs && child.paragraphs.length > 0) {
    for (const paragraph of child.paragraphs) drawing.appendChild(renderNativeDrawingParagraph(paragraph));
  } else if (child.text && child.text.length > 0) {
    const text = document.createElement('span');
    text.className = 'coco-word-native-drawing-text';
    if (child.innerTextAlign) text.style.textAlign = mapDrawingTextAlign(child.innerTextAlign);
    text.textContent = child.text;
    drawing.appendChild(text);
  } else {
    drawing.setAttribute('aria-label', child.drawingKind);
  }
  return drawing;
}

function isVerticalLine(child: NativeWordDrawingNode): boolean {
  return (child.height ?? 0) > Math.max(1, (child.width ?? 0) * 2);
}

function shouldApplyDrawingMinHeight(child: NativeWordDrawingNode): boolean {
  if (child.drawingKind === 'textBox' && !child.hasBorder && !child.hasFill) return false;
  return true;
}

export function nativeDrawingWidthCss(
  child: Pick<NativeWordDrawingNode, 'width' | 'anchoringMode' | 'wrap' | 'isLineOrConnector'>,
): string | undefined {
  if (!child.width) return undefined;
  if (child.anchoringMode === 'anchor' && isAbsoluteAnchorWrap(child.wrap)) return `${child.width}px`;
  const availableWidth = child.isLineOrConnector ? '100%' : 'max(1px, calc(100% - 4pt))';
  return `min(${child.width}px, ${availableWidth})`;
}

function isAbsoluteAnchorWrap(wrap: NativeWordDrawingNode['wrap']): boolean {
  return wrap === 'none' || wrap === 'behind' || wrap === 'inFront';
}

function renderNativeDrawingParagraph(paragraph: NativeWordParagraphNode): HTMLElement {
  const line = document.createElement('div');
  line.className = 'coco-word-native-drawing-paragraph';
  const paragraphStyle = resolveParagraphStyle(paragraph.format).cssText;
  if (paragraphStyle.length > 0) line.setAttribute('style', paragraphStyle);
  for (const child of paragraph.content) {
    if (child.kind === 'text') {
      line.appendChild(renderNativeText(child));
    } else if (child.kind === 'field') {
      line.appendChild(renderNativeField(child));
    } else if (child.kind === 'image' || child.kind === 'drawing') {
      line.appendChild(document.createTextNode(child.kind === 'drawing' && child.text ? child.text : '\uFFFC'));
    }
  }
  return line;
}

function renderNativeText(child: NativeWordParagraphNode['content'][number]): Node {
  if (child.kind !== 'text') return document.createTextNode('');
  const runStyle = resolveRunStyle(child.format).cssText;
  const commentMark = child.marks?.find((mark) => mark.kind === 'comment');
  const revisionMark = child.marks?.find((mark) => mark.kind === 'revision');
  const formatRevision = child.formatRevision;
  if (!commentMark && !revisionMark && !formatRevision && !child.hyperlink && runStyle.length === 0) return document.createTextNode(child.text);

  const span = child.hyperlink ? document.createElement('a') : document.createElement('span');
  if (runStyle.length > 0) span.setAttribute('style', runStyle);
  if (child.hyperlink) {
    const href = safeRenderedHyperlinkHref(child.hyperlink.target, child.hyperlink.anchor);
    span.classList.add('coco-word-hyperlink');
    span.setAttribute('href', href);
    span.dataset.hyperlinkTarget = child.hyperlink.target ?? '';
    span.dataset.hyperlinkAnchor = child.hyperlink.anchor ?? '';
    if (child.hyperlink.tooltip) span.title = child.hyperlink.tooltip;
    if (child.hyperlink.target) {
      span.setAttribute('target', '_blank');
      span.setAttribute('rel', 'noopener noreferrer');
    }
  }
  if (commentMark) {
    span.classList.add('coco-word-comment');
    span.dataset.commentId = String(commentMark.commentId);
    span.dataset.commentAuthor = commentMark.author;
    span.dataset.commentDate = commentMark.date;
    span.dataset.commentBody = commentMark.body;
    if (commentMark.parentCommentId !== undefined) {
      span.dataset.parentCommentId = String(commentMark.parentCommentId);
      span.classList.add('coco-word-comment-reply');
    }
    span.title = commentTitle(commentMark);
  }
  if (revisionMark) {
    span.classList.add(
      revisionMark.revisionKind === 'delete' ? 'coco-word-revision-delete' : 'coco-word-revision-insert',
    );
    span.dataset.revisionId = revisionMark.revisionId ?? '';
    span.dataset.revisionKind = revisionMark.revisionKind;
    span.dataset.revisionAuthor = revisionMark.author ?? '';
    span.dataset.revisionDate = revisionMark.date ?? '';
    span.title = revisionTitle(revisionMark);
  }
  if (formatRevision) {
    span.classList.add('coco-word-revision-format');
    span.dataset.revisionId = formatRevision.revisionId;
    span.dataset.revisionKind = 'format';
    span.dataset.revisionAuthor = formatRevision.author ?? '';
    span.dataset.revisionDate = formatRevision.date ?? '';
    span.title = '格式修订';
  }
  span.appendChild(document.createTextNode(child.text));
  return span;
}

function renderNativeField(child: NativeWordFieldNode): HTMLElement {
  const span = document.createElement('span');
  span.className = 'coco-word-native-field';
  span.dataset.fieldKind = child.fieldKind;
  span.dataset.fieldInstruction = child.instruction;
  const runStyle = resolveRunStyle(child.format).cssText;
  if (runStyle.length > 0) span.setAttribute('style', runStyle);
  span.contentEditable = 'false';
  span.textContent = child.cachedText ?? '';
  return span;
}

function commentTitle(mark: { commentId: number; author?: string; body?: string }): string {
  const author = mark.author ? `${mark.author}: ` : '';
  const body = mark.body?.trim() ?? '';
  return body.length > 0 ? `Comment ${mark.commentId} ${author}${body}` : `Comment ${mark.commentId}`;
}

function revisionTitle(mark: {
  revisionId?: string;
  revisionKind?: 'insert' | 'delete';
  author?: string;
  date?: string;
}): string {
  const label = mark.revisionKind === 'delete' ? 'Deleted' : 'Inserted';
  const author = mark.author ? ` by ${mark.author}` : '';
  const date = mark.date ? ` on ${mark.date}` : '';
  const id = mark.revisionId ? ` (${mark.revisionId})` : '';
  return `${label}${author}${date}${id}`;
}

function applyDrawingLayoutStyle(el: HTMLElement, child: NativeWordDrawingNode): void {
  if (child.anchoringMode !== 'anchor') return;
  el.classList.add('coco-word-anchor-drawing');
  applyAnchorLayoutStyle(el, child);
}

function applyImageLayoutStyle(img: HTMLImageElement, child: NativeWordImageNode): void {
  if (child.anchoringMode !== 'anchor') return;
  img.classList.add('coco-word-anchor-image');
  applyAnchorLayoutStyle(img, child);
}

function applyAnchorLayoutStyle(
  el: HTMLElement,
  child: Pick<NativeWordImageNode | NativeWordDrawingNode, 'position' | 'wrap'>,
): void {
  const metadata = nativeAnchorLayoutMetadata(child);
  for (const className of metadata.classNames) el.classList.add(className);
  for (const [name, value] of Object.entries(metadata.dataset)) el.dataset[name] = value;
  for (const [name, value] of Object.entries(metadata.styleProperties)) el.style.setProperty(name, value);
}

export function nativeAnchorLayoutMetadata(
  child: Pick<NativeWordImageNode | NativeWordDrawingNode, 'position' | 'wrap'>,
): NativeAnchorLayoutMetadata {
  const classNames: string[] = [];
  const dataset: Record<string, string> = {};
  const styleProperties: Record<string, string> = {};
  collectAnchorPositionMetadata(child.position, dataset, styleProperties);
  const side = anchorHorizontalSide(child.position);
  const horizontalRelative = child.position?.horizontal?.relativeFrom;
  const verticalRelative = child.position?.vertical?.relativeFrom;

  if (child.wrap === 'square' || child.wrap === 'tight' || child.wrap === 'through') {
    if (side === 'left') classNames.push('is-float-left');
    else if (side === 'right') classNames.push('is-float-right');
    else if (side === 'center') classNames.push('is-centered');
    else classNames.push('is-float-right');
  } else if (child.wrap === 'behind' || child.wrap === 'inFront' || child.wrap === 'none') {
    classNames.push('is-overlay-like');
    classNames.push('is-absolute-anchor');
    if (isPageLayerRelativeFrom(verticalRelative)) classNames.push('is-page-relative-anchor');
    else classNames.push('is-paragraph-relative-anchor');
    if (horizontalRelative === 'column') classNames.push('is-column-relative-anchor');
    if (side === 'left') classNames.push('is-anchor-left');
    if (side === 'right') classNames.push('is-anchor-right');
    if (side === 'center') classNames.push('is-anchor-center');
    if (child.wrap === 'behind') classNames.push('is-behind-text');
    if (child.wrap === 'inFront') classNames.push('is-in-front-of-text');
  } else if (child.wrap === 'topAndBottom') {
    classNames.push('is-top-bottom');
    if (side === 'center') classNames.push('is-centered');
  }
  return { classNames, dataset, styleProperties };
}

function collectAnchorPositionMetadata(
  position: NativeWordImageNode['position'] | NativeWordDrawingNode['position'] | undefined,
  dataset: Record<string, string>,
  styleProperties: Record<string, string>,
): void {
  const horizontal = position?.horizontal;
  const vertical = position?.vertical;
  if (horizontal?.relativeFrom) dataset.positionHorizontalRelativeFrom = horizontal.relativeFrom;
  if (vertical?.relativeFrom) dataset.positionVerticalRelativeFrom = vertical.relativeFrom;
  if (horizontal?.align) dataset.positionHorizontalAlign = horizontal.align;
  if (vertical?.align) dataset.positionVerticalAlign = vertical.align;
  if (horizontal?.posOffset !== undefined) {
    dataset.positionHorizontalOffset = String(horizontal.posOffset);
    styleProperties['--coco-word-anchor-left'] = `${emuToPxNumber(horizontal.posOffset)}px`;
  }
  if (vertical?.posOffset !== undefined) {
    dataset.positionVerticalOffset = String(vertical.posOffset);
    styleProperties['--coco-word-anchor-top'] = `${emuToPxNumber(vertical.posOffset)}px`;
  }
}

function anchorHorizontalSide(
  position: NativeWordImageNode['position'] | NativeWordDrawingNode['position'] | undefined,
): 'left' | 'right' | 'center' | null {
  const horizontal = position?.horizontal;
  if (!horizontal) return null;
  if (horizontal.align === 'left' || horizontal.align === 'inside') return 'left';
  if (horizontal.align === 'right' || horizontal.align === 'outside') return 'right';
  if (horizontal.align === 'center') return 'center';
  if (horizontal.posOffset === undefined) return null;
  const px = emuToPxNumber(horizontal.posOffset);
  if (px <= 8) return 'left';
  if (px >= 240) return 'right';
  return null;
}

function mapDrawingTextAlign(value: string): string {
  if (value === 'both') return 'justify';
  return value;
}

export function renderNativeTable(table: NativeWordTableNode): HTMLElement {
  const el = document.createElement('coco-word-native-table');
  el.setAttribute('data-table-id', table.tableId);
  if (table.revision) {
    el.classList.add(table.revision.revisionKind === 'delete' ? 'coco-word-revision-delete' : 'coco-word-revision-insert');
    el.dataset.revisionId = table.revision.revisionId;
    el.dataset.revisionKind = table.revision.revisionKind;
    el.dataset.revisionAuthor = table.revision.author ?? '';
    el.dataset.revisionDate = table.revision.date ?? '';
    el.dataset.revisionText = '表格结构';
  }
  if (table.tableStyleId) el.setAttribute('data-style-id', table.tableStyleId);
  const tableEl = document.createElement('table');
  const border = tableBorderCss(table);
  if (border) tableEl.style.setProperty('--coco-word-table-border', border);
  applyNativeTableLayout(tableEl, table);
  if (table.columnWidths.some((width) => width > 0)) {
    const colgroup = document.createElement('colgroup');
    let tableWidth = 0;
    for (const width of table.columnWidths) tableWidth += Math.max(0, width);
    for (const width of table.columnWidths) {
      const col = document.createElement('col');
      if (width > 0) {
        col.style.width =
          table.layout?.widthType === 'pct' && tableWidth > 0
            ? `${(width / tableWidth) * 100}%`
            : twipsToPt(width);
      }
      colgroup.appendChild(col);
    }
    tableEl.appendChild(colgroup);
    if (tableWidth > 0 && tableEl.style.width.length === 0) tableEl.style.width = twipsToPt(tableWidth);
  }
  const tbody = document.createElement('tbody');
  const columnCount = nativeTableColumnCount(table);
  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex += 1) {
    const row = table.rows[rowIndex]!;
    const tr = document.createElement('tr');
    tr.dataset.rowIndex = String(rowIndex);
    if (row.isHeader) tr.dataset.repeatHeader = 'true';
    if (row.cantSplit) tr.dataset.cantSplit = 'true';
    if (row.height && row.height.val > 0) {
      const height = twipsToPt(row.height.val);
      tr.dataset.heightRule = row.height.hRule ?? 'auto';
      tr.style.setProperty('--coco-word-row-height', height);
      if (row.height.hRule === 'exact') {
        tr.style.height = height;
      } else {
        tr.style.height = height;
      }
    }
    let columnIndex = 0;
    for (const cell of row.cells) {
      const gridSpan = Math.max(1, cell.gridSpan);
      const cellColumnIndex = columnIndex;
      columnIndex += gridSpan;
      if (cell.vMerge === 'continue') continue;
      const td = document.createElement('td');
      td.dataset.rowIndex = String(rowIndex);
      td.dataset.cellIndex = String(row.cells.indexOf(cell));
      td.dataset.columnIndex = String(cellColumnIndex);
      if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
      if (cell.rowSpan && cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      const cellStyle = resolveTableCellStyle({
        ...cell,
        borders: resolveNativeTableCellBorders(table, cell, {
          rowIndex,
          rowCount: table.rows.length,
          columnIndex: cellColumnIndex,
          columnSpan: gridSpan,
          columnCount,
        }),
      }).cssText;
      if (cellStyle.length > 0) td.setAttribute('style', cellStyle);
      td.contentEditable = 'true';
      for (const block of cell.blocks) {
        const renderedBlock = renderNativeBlock(block);
        for (const paragraph of [
          ...(renderedBlock.matches('coco-word-native-p') ? [renderedBlock] : []),
          ...Array.from(renderedBlock.querySelectorAll<HTMLElement>('coco-word-native-p')),
        ]) {
          paragraph.contentEditable = 'true';
        }
        td.appendChild(renderedBlock);
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  tableEl.appendChild(tbody);
  el.appendChild(tableEl);
  markNativeTableEditable(el);
  return el;
}

export interface NativeTableCellPosition {
  rowIndex: number;
  rowCount: number;
  columnIndex: number;
  columnSpan: number;
  columnCount: number;
}

function nativeTableColumnCount(table: NativeWordTableNode): number {
  const gridColumnCount = table.columnWidths.filter((width) => width > 0).length;
  if (gridColumnCount > 0) return gridColumnCount;
  return Math.max(
    1,
    ...table.rows.map((row) =>
      row.cells.reduce((sum, cell) => sum + Math.max(1, cell.gridSpan), 0),
    ),
  );
}

export function resolveNativeTableCellBorders(
  table: NativeWordTableNode,
  cell: NativeWordTableCellNode,
  position: NativeTableCellPosition,
): EditorTableBorders | undefined {
  const tableBorders = table.borders;
  if (!tableBorders && !cell.borders) return undefined;
  return {
    ...effectiveTableCellBorders(tableBorders, position),
    ...(cell.borders ?? {}),
  };
}

function effectiveTableCellBorders(
  borders: EditorTableBorders | undefined,
  position: NativeTableCellPosition,
): EditorTableBorders {
  if (!borders) return {};
  const nil = { val: 'nil' };
  const isFirstRow = position.rowIndex === 0;
  const isLastRow = position.rowIndex >= position.rowCount - 1;
  const isFirstColumn = position.columnIndex === 0;
  const isLastColumn = position.columnIndex + position.columnSpan >= position.columnCount;
  return {
    top: (isFirstRow ? borders.top : borders.insideH) ?? nil,
    right: (isLastColumn ? borders.right : borders.insideV) ?? nil,
    bottom: (isLastRow ? borders.bottom : borders.insideH) ?? nil,
    left: (isFirstColumn ? borders.left : borders.insideV) ?? nil,
  };
}

function applyNativeTableLayout(tableEl: HTMLElement, table: NativeWordTableNode): void {
  const layout = table.layout;
  if (!layout) return;

  if (layout.width !== undefined && layout.width > 0) {
    if (layout.widthType === 'pct') tableEl.style.width = `${layout.width / 50}%`;
    else if (layout.widthType !== 'auto') tableEl.style.width = twipsToPt(layout.width);
  }

  if (layout.alignment === 'center') {
    tableEl.style.marginLeft = 'auto';
    tableEl.style.marginRight = 'auto';
  } else if (layout.alignment === 'right') {
    tableEl.style.marginLeft = 'auto';
    tableEl.style.marginRight = '0';
  } else if (layout.indent !== undefined && layout.indent > 0) {
    tableEl.style.marginLeft = twipsToPt(layout.indent);
  }

  if (layout.cellSpacing !== undefined && layout.cellSpacing > 0) {
    tableEl.style.borderCollapse = 'separate';
    tableEl.style.borderSpacing =
      layout.cellSpacingType === 'pct' ? `${layout.cellSpacing / 50}%` : twipsToPt(layout.cellSpacing);
  }
}

function twipsToPt(twips: number): string {
  return `${twips / 20}pt`;
}

function twipsToPx(twips: number): string {
  return `${twipsToPxNumber(twips)}px`;
}

function twipsToPxNumber(twips: number): number {
  return twips / 15;
}

function emuToPxNumber(emu: number): number {
  return (emu / 914400) * 96;
}

function tableBorderCss(table: NativeWordTableNode): string | null {
  const border =
    table.borders?.insideH ??
    table.borders?.insideV ??
    table.borders?.top ??
    table.borders?.left ??
    table.borders?.bottom ??
    table.borders?.right;
  if (!border || border.val === 'none' || border.val === 'nil') return null;
  const color = border.color && /^[0-9a-f]{6}$/i.test(border.color) ? `#${border.color}` : '#d4d4d8';
  const width = border.size !== undefined ? Math.max(0.5, border.size / 8) : 0.75;
  const style = border.val === 'double' ? 'double' : 'solid';
  return `${width}pt ${style} ${color}`;
}

function defineElement(name: string): void {
  if (customElements.get(name)) return;
  customElements.define(name, class extends HTMLElement {});
}
