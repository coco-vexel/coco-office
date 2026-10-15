import type { EditorHyperlink, EditorRunFormat } from './model.js';
import { nativeParagraphText, type NativeWordBlockNode } from './native-model.js';
import type { NativeWordListType } from './native-numbering.js';
import { safeRenderedHyperlinkHref } from './native-url-safety.js';

export const NATIVE_BLOCK_CLIPBOARD_TYPE = 'application/x-coco-word-blocks+json';

export interface NativeClipboardRun {
  from: number;
  to: number;
  format: Partial<EditorRunFormat>;
  hyperlink?: EditorHyperlink;
}

export interface NativeClipboardParagraph {
  text: string;
  runs: NativeClipboardRun[];
  listType?: NativeWordListType;
}

export function parseNativeClipboardHtml(html: string): NativeClipboardParagraph[] {
  const document = new DOMParser().parseFromString(html, 'text/html');
  const candidates = Array.from(document.body.querySelectorAll('p, li, h1, h2, h3, h4, h5, h6')).filter(
    (element) => element.tagName !== 'LI' || !element.querySelector('p'),
  );
  const blocks = candidates.length > 0 ? candidates : [document.body];
  return blocks.map((block) => {
    const paragraph: NativeClipboardParagraph = { text: '', runs: [] };
    const list = block.closest('li');
    if (list) paragraph.listType = list.parentElement?.tagName === 'OL' ? 'numbered' : 'bullet';
    appendClipboardNode(block, {}, paragraph);
    paragraph.text = paragraph.text.replace(/\u00a0/g, ' ');
    return paragraph;
  });
}

export function serializeNativeClipboardHtml(paragraphs: NativeClipboardParagraph[]): string {
  let html = '';
  let activeList: NativeWordListType | null = null;
  for (const paragraph of paragraphs) {
    if (paragraph.listType !== activeList) {
      if (activeList) html += activeList === 'numbered' ? '</ol>' : '</ul>';
      activeList = paragraph.listType ?? null;
      if (activeList) html += activeList === 'numbered' ? '<ol>' : '<ul>';
    }
    const content = serializeClipboardParagraphContent(paragraph);
    html += activeList ? `<li>${content}</li>` : `<p>${content}</p>`;
  }
  if (activeList) html += activeList === 'numbered' ? '</ol>' : '</ul>';
  return html;
}

export function serializeNativeBlockClipboard(blocks: readonly NativeWordBlockNode[]): string {
  return JSON.stringify({ version: 1, blocks });
}

export function parseNativeBlockClipboard(value: string): NativeWordBlockNode[] | null {
  if (!value || value.length > 20_000_000) return null;
  try {
    const payload = JSON.parse(value) as unknown;
    if (!isRecord(payload) || payload.version !== 1 || !Array.isArray(payload.blocks)) return null;
    if (payload.blocks.length === 0 || payload.blocks.length > 1_000) return null;
    return payload.blocks.every(isNativeClipboardBlock) ? payload.blocks as NativeWordBlockNode[] : null;
  } catch {
    return null;
  }
}

export function nativeBlockClipboardText(blocks: readonly NativeWordBlockNode[]): string {
  return blocks.map(nativeClipboardBlockText).join('\n');
}

function nativeClipboardBlockText(block: NativeWordBlockNode): string {
  if (block.kind === 'paragraph') return nativeParagraphText(block);
  return block.rows
    .map((row) => row.cells.map((cell) => cell.blocks.map(nativeClipboardBlockText).join('\n')).join('\t'))
    .join('\n');
}

function isNativeClipboardBlock(value: unknown): boolean {
  if (!isRecord(value)) return false;
  if (value.kind === 'paragraph') {
    return typeof value.paraId === 'string' && typeof value.story === 'string' && Array.isArray(value.content) &&
      value.content.length <= 100_000 && value.content.every(isNativeClipboardInline);
  }
  if (value.kind !== 'table' || typeof value.tableId !== 'string' || !Array.isArray(value.rows) || value.rows.length > 10_000) {
    return false;
  }
  return value.rows.every((row) =>
    isRecord(row) && Array.isArray(row.cells) && row.cells.length <= 1_000 && row.cells.every((cell) =>
      isRecord(cell) && Array.isArray(cell.blocks) && cell.blocks.length <= 10_000 && cell.blocks.every(isNativeClipboardBlock),
    ),
  );
}

function isNativeClipboardInline(value: unknown): boolean {
  if (!isRecord(value) || typeof value.kind !== 'string') return false;
  if (value.kind === 'text') return typeof value.text === 'string';
  if (value.kind === 'image') return typeof value.drawingId === 'string' && typeof value.src === 'string';
  if (value.kind === 'drawing') return typeof value.drawingId === 'string';
  if (value.kind === 'pageBreak') return true;
  if (value.kind === 'field') return typeof value.fieldKind === 'string' && typeof value.instruction === 'string';
  return false;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function serializeClipboardParagraphContent(paragraph: NativeClipboardParagraph): string {
  let html = '';
  let offset = 0;
  const runs = [...paragraph.runs]
    .map((run) => ({ ...run, from: Math.max(0, run.from), to: Math.min(paragraph.text.length, run.to) }))
    .filter((run) => run.to > run.from)
    .sort((left, right) => left.from - right.from || left.to - right.to);
  for (const run of runs) {
    if (run.from > offset) html += escapeClipboardText(paragraph.text.slice(offset, run.from));
    if (run.from < offset) continue;
    const text = escapeClipboardText(paragraph.text.slice(run.from, run.to));
    const style = clipboardRunStyle(run.format);
    let content = style ? `<span style="${escapeHtmlAttribute(style)}">${text}</span>` : text;
    if (run.hyperlink) {
      const href = safeRenderedHyperlinkHref(run.hyperlink.target, run.hyperlink.anchor);
      const title = run.hyperlink.tooltip ? ` title="${escapeHtmlAttribute(run.hyperlink.tooltip)}"` : '';
      content = `<a href="${escapeHtmlAttribute(href)}"${title}>${content}</a>`;
    }
    html += content;
    offset = run.to;
  }
  if (offset < paragraph.text.length) html += escapeClipboardText(paragraph.text.slice(offset));
  return html || '<br>';
}

function clipboardRunStyle(format: Partial<EditorRunFormat>): string {
  const declarations: string[] = [];
  if (format.bold !== undefined) declarations.push(`font-weight:${format.bold ? '700' : '400'}`);
  if (format.italic !== undefined) declarations.push(`font-style:${format.italic ? 'italic' : 'normal'}`);
  const decorations: string[] = [];
  if (format.underline) decorations.push('underline');
  if (format.strike || format.doubleStrike) decorations.push('line-through');
  if (decorations.length > 0) declarations.push(`text-decoration-line:${decorations.join(' ')}`);
  else if (format.underline === false || format.strike === false || format.doubleStrike === false) {
    declarations.push('text-decoration-line:none');
  }
  if (format.fontFamily) declarations.push(`font-family:${format.fontFamily}`);
  if (format.fontSize !== undefined) declarations.push(`font-size:${format.fontSize}pt`);
  if (format.color) declarations.push(`color:#${format.color.replace(/^#/, '')}`);
  if (format.highlight && format.highlight !== 'none') {
    declarations.push(`background-color:#${format.highlight.replace(/^#/, '')}`);
  }
  if (format.verticalAlign) declarations.push(`vertical-align:${format.verticalAlign}`);
  if (format.caps) declarations.push('text-transform:uppercase');
  if (format.smallCaps) declarations.push('font-variant-caps:small-caps');
  if (format.characterSpacing !== undefined) declarations.push(`letter-spacing:${format.characterSpacing / 20}pt`);
  return declarations.join(';');
}

function escapeClipboardText(value: string): string {
  return escapeHtml(value).replace(/\r\n?|\n/g, '<br>');
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeHtmlAttribute(value: string): string {
  return escapeHtml(value).replace(/"/g, '&quot;');
}

function appendClipboardNode(
  node: Node,
  inherited: Partial<EditorRunFormat>,
  paragraph: NativeClipboardParagraph,
  inheritedHyperlink?: EditorHyperlink,
): void {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? '';
    if (!text) return;
    const from = paragraph.text.length;
    paragraph.text += text;
    if (Object.keys(inherited).length > 0 || inheritedHyperlink) {
      paragraph.runs.push({
        from,
        to: paragraph.text.length,
        format: inherited,
        ...(inheritedHyperlink ? { hyperlink: inheritedHyperlink } : {}),
      });
    }
    return;
  }
  if (!(node instanceof HTMLElement)) return;
  if (node.tagName === 'BR') {
    paragraph.text += '\n';
    return;
  }
  const format = { ...inherited, ...clipboardElementFormat(node) };
  const hyperlink = node.tagName === 'A' ? clipboardElementHyperlink(node) : inheritedHyperlink;
  for (const child of Array.from(node.childNodes)) appendClipboardNode(child, format, paragraph, hyperlink);
}

function clipboardElementHyperlink(element: HTMLElement): EditorHyperlink | undefined {
  const href = element.getAttribute('href')?.trim();
  if (!href) return undefined;
  const tooltip = element.getAttribute('title')?.trim();
  if (href.startsWith('#')) {
    let anchor = href.slice(1);
    try {
      anchor = decodeURIComponent(anchor);
    } catch {
      return undefined;
    }
    return anchor ? { anchor, ...(tooltip ? { tooltip } : {}) } : undefined;
  }
  const safeHref = safeRenderedHyperlinkHref(href, undefined);
  return safeHref === '#' ? undefined : { target: safeHref, ...(tooltip ? { tooltip } : {}) };
}

function clipboardElementFormat(element: HTMLElement): Partial<EditorRunFormat> {
  const style = element.style;
  const format: Partial<EditorRunFormat> = {};
  if (element.matches('b, strong') || Number.parseInt(style.fontWeight, 10) >= 600) format.bold = true;
  if (element.matches('i, em') || style.fontStyle === 'italic') format.italic = true;
  const decoration = style.textDecorationLine || style.textDecoration;
  if (element.matches('u') || decoration.includes('underline')) format.underline = true;
  if (element.matches('s, strike, del') || decoration.includes('line-through')) format.strike = true;
  if (style.fontFamily) format.fontFamily = style.fontFamily.replace(/["']/g, '');
  const fontSize = cssFontSizePoints(style.fontSize);
  if (fontSize !== null) format.fontSize = fontSize;
  const color = normalizeCssColor(style.color);
  const background = normalizeCssColor(style.backgroundColor);
  if (color) format.color = color;
  if (background) format.highlight = background;
  return format;
}

function normalizeCssColor(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed || trimmed === 'transparent') return null;
  const hex = /^#([\da-f]{3}|[\da-f]{6})$/i.exec(trimmed);
  if (hex) {
    const digits = hex[1]!;
    return (digits.length === 3 ? digits.split('').map((digit) => digit + digit).join('') : digits).toUpperCase();
  }
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(trimmed);
  if (!rgb) return null;
  return rgb.slice(1, 4).map((part) => Math.max(0, Math.min(255, Number(part))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function cssFontSizePoints(value: string): number | null {
  const match = /^([\d.]+)(pt|px)$/i.exec(value.trim());
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return null;
  return match[2]?.toLowerCase() === 'px' ? amount * 0.75 : amount;
}
