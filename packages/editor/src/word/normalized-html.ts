import type {
  EditorBlockView,
  EditorContentSegment,
  EditorDocumentView,
  EditorParagraphView,
  EditorTableView,
} from './model.js';
import { resolveParagraphStyle, resolveRunStyle, resolveTableCellStyle } from './style-resolver.js';

export interface NormalizedDocumentHtmlPageOptions {
  /** Zero-based top-level block offset. */
  startBlock?: number;
  /** Approximate serialized HTML budget. A block is never split. */
  maxChars?: number;
}

export interface NormalizedDocumentHtmlPage {
  html: string;
  /** Zero-based block offset for the next page. Missing means end of document. */
  nextBlock?: number;
}

/**
 * Serialize an editor view into compact, semantic HTML for model inspection.
 *
 * The projection intentionally omits Word style ids. Effective formatting has
 * already been resolved by the editor extractor and is emitted as inline CSS.
 * Stable paragraph/table ids remain available for later document operations.
 */
export function renderNormalizedDocumentHtmlPage(
  view: EditorDocumentView,
  options: NormalizedDocumentHtmlPageOptions = {},
): NormalizedDocumentHtmlPage {
  const startBlock = clampInteger(options.startBlock ?? 0, 0, view.blocks.length);
  const maxChars = Math.max(1, Math.floor(options.maxChars ?? 12_000));
  const rendered: string[] = [];
  let htmlChars = '<article></article>'.length;
  let blockIndex = startBlock;

  while (blockIndex < view.blocks.length) {
    const blockHtml = renderBlock(view.blocks[blockIndex]!);
    const separatorChars = rendered.length > 0 ? 1 : 0;
    if (rendered.length > 0 && htmlChars + separatorChars + blockHtml.length > maxChars) break;
    rendered.push(blockHtml);
    htmlChars += separatorChars + blockHtml.length;
    blockIndex += 1;
  }

  const html =
    rendered.length > 0 ? `<article>\n${rendered.join('\n')}\n</article>` : '<article></article>';
  return blockIndex < view.blocks.length ? { html, nextBlock: blockIndex } : { html };
}

function renderBlock(block: EditorBlockView): string {
  return block.type === 'table' ? renderTable(block) : renderParagraph(block);
}

function renderParagraph(paragraph: EditorParagraphView): string {
  const tag = headingTag(paragraph);
  const style = resolveParagraphStyle(paragraph.format).cssText;
  const attributes = [
    `data-para-id="${escapeAttribute(paragraph.paraId)}"`,
    style ? `style="${escapeAttribute(style)}"` : '',
  ]
    .filter(Boolean)
    .join(' ');
  const content = renderParagraphContent(paragraph);
  return `<${tag} ${attributes}>${content || '&nbsp;'}</${tag}>`;
}

function headingTag(paragraph: EditorParagraphView): `h${number}` | 'p' {
  const level = paragraph.format?.outlineLevel;
  if (level === undefined || level < 0 || level > 5) return 'p';
  return `h${level + 1}`;
}

function renderParagraphContent(paragraph: EditorParagraphView): string {
  const label = paragraph.format?.numbering?.label;
  const prefix = label ? `${escapeHtml(label)} ` : '';
  if (paragraph.segments.length === 0) return prefix + renderText(paragraph.text);
  return prefix + paragraph.segments.map(renderSegment).join('');
}

function renderSegment(segment: EditorContentSegment): string {
  if (segment.kind === 'image') {
    return `<img data-drawing-id="${escapeAttribute(segment.drawingId)}" alt="[图片]">`;
  }
  if (segment.kind === 'drawing') {
    const text = segment.text ? renderText(segment.text) : '[绘图]';
    return `<span data-drawing-id="${escapeAttribute(segment.drawingId)}">${text}</span>`;
  }
  if (segment.kind === 'placeholder') {
    if (segment.placeholderKind === 'pageBreak') return '<br data-page-break="true">';
    return `<span data-object="${escapeAttribute(segment.placeholderKind)}">[对象]</span>`;
  }
  if (segment.kind === 'field') {
    const content = renderText(segment.cachedText ?? '');
    return `<span data-field="${escapeAttribute(segment.fieldKind)}">${content}</span>`;
  }

  const style = resolveRunStyle(segment.format).cssText;
  const content = renderText(segment.text);
  const styleAttribute = style ? ` style="${escapeAttribute(style)}"` : '';
  if (segment.kind === 'revision') {
    const tag = segment.revisionKind === 'delete' ? 'del' : 'ins';
    return `<${tag}${styleAttribute}>${content}</${tag}>`;
  }
  if (segment.hyperlink) {
    const href = segment.hyperlink.target ?? segment.hyperlink.anchor;
    const hrefAttribute = href ? ` href="${escapeAttribute(href)}"` : '';
    return `<a${hrefAttribute}${styleAttribute}>${content}</a>`;
  }
  return style ? `<span${styleAttribute}>${content}</span>` : content;
}

function renderTable(table: EditorTableView): string {
  const rows = table.rows
    .map((row) => {
      const cells = row.cells
        .map((cell) => {
          const style = resolveTableCellStyle(cell).cssText;
          const attributes = [
            cell.gridSpan > 1 ? `colspan="${cell.gridSpan}"` : '',
            cell.rowSpan !== undefined && cell.rowSpan > 1 ? `rowspan="${cell.rowSpan}"` : '',
            style ? `style="${escapeAttribute(style)}"` : '',
          ]
            .filter(Boolean)
            .join(' ');
          const content = cell.blocks.map(renderBlock).join('\n');
          return `<td${attributes ? ` ${attributes}` : ''}>${content || '&nbsp;'}</td>`;
        })
        .join('');
      return `<tr>${cells}</tr>`;
    })
    .join('\n');
  return `<table data-table-id="${escapeAttribute(table.tableId)}">\n${rows}\n</table>`;
}

function renderText(value: string): string {
  return escapeHtml(value)
    .replace(/\t/g, '&#9;')
    .replace(/\r\n|\r|\n/g, '<br>');
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttribute(value: string): string {
  return escapeHtml(value).replace(/"/g, '&quot;');
}

function clampInteger(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, Math.floor(value)));
}
