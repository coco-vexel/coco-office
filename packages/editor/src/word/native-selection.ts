import type { WordEditorSelection } from './runtime.js';

export interface NativeCursorPoint {
  paraId: string;
  offset: number;
}

export interface NativeSelectedImage {
  paraId: string;
  drawingId: string;
}

export function nativeSelection(host: HTMLElement): Selection | null {
  const root = host.getRootNode();
  if (root instanceof ShadowRoot) {
    return (root as ShadowRoot & { getSelection?: () => Selection | null }).getSelection?.() ?? null;
  }
  return host.ownerDocument.getSelection();
}

export function currentNativeCursor(host: HTMLElement): NativeCursorPoint | null {
  const selection = nativeSelection(host);
  if (!selection || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  const paragraph = parentNativeParagraph(range.startContainer);
  if (!paragraph || !host.contains(paragraph)) return null;
  return {
    paraId: paragraph.dataset.paraId ?? '',
    offset: textOffsetWithin(paragraph, range.startContainer, range.startOffset),
  };
}

export function nativeSelectionFromCursor(host: HTMLElement): WordEditorSelection | null {
  const point = currentNativeCursor(host);
  return point ? { kind: 'caret', paraId: point.paraId, offset: point.offset } : null;
}

export function currentNativeSelection(host: HTMLElement): WordEditorSelection | null {
  const selection = nativeSelection(host);
  if (!selection || selection.rangeCount === 0) return null;
  if (!selection.anchorNode || !selection.focusNode) return null;
  const anchorParagraph = parentNativeParagraph(selection.anchorNode);
  const focusParagraph = parentNativeParagraph(selection.focusNode);
  if (!anchorParagraph || !focusParagraph || !host.contains(anchorParagraph) || !host.contains(focusParagraph)) {
    return null;
  }

  const anchor = {
    paraId: anchorParagraph.dataset.paraId ?? '',
    offset: textOffsetWithin(anchorParagraph, selection.anchorNode, selection.anchorOffset),
  };
  const focus = {
    paraId: focusParagraph.dataset.paraId ?? '',
    offset: textOffsetWithin(focusParagraph, selection.focusNode, selection.focusOffset),
  };

  if (selection.isCollapsed) return { kind: 'caret', paraId: anchor.paraId, offset: anchor.offset };
  return { kind: 'text-range', anchor, focus };
}

export function nativeCursorFromPoint(host: HTMLElement, x: number, y: number): NativeCursorPoint | null {
  const caret = caretNodeFromPoint(x, y);
  if (!caret) return null;
  const paragraph = parentNativeParagraph(caret.node);
  if (!paragraph || !host.contains(paragraph)) return null;
  return {
    paraId: paragraph.dataset.paraId ?? '',
    offset: textOffsetWithin(paragraph, caret.node, caret.offset),
  };
}

export function restoreNativeCursor(host: HTMLElement, cursor: NativeCursorPoint): void {
  const paragraph = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${cursor.paraId}"]`);
  if (!paragraph) return;
  paragraph.focus({ preventScroll: true });
  const range = document.createRange();
  const target = textNodeAt(paragraph, cursor.offset);
  range.setStart(target.node, target.offset);
  range.setEnd(target.node, target.offset);
  const selection = nativeSelection(host);
  if (!selection) return;
  selection.removeAllRanges();
  selection.addRange(range);
}

export function restoreNativeTextSelection(
  host: HTMLElement,
  anchor: NativeCursorPoint,
  focus: NativeCursorPoint,
): void {
  const anchorParagraph = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${anchor.paraId}"]`);
  const focusParagraph = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${focus.paraId}"]`);
  if (!anchorParagraph || !focusParagraph) return;
  const anchorTarget = textNodeAt(anchorParagraph, anchor.offset);
  const focusTarget = textNodeAt(focusParagraph, focus.offset);
  const selection = nativeSelection(host);
  if (!selection) return;
  anchorParagraph.focus({ preventScroll: true });
  selection.removeAllRanges();
  selection.setBaseAndExtent(anchorTarget.node, anchorTarget.offset, focusTarget.node, focusTarget.offset);
}

export function restoreNativeElementSelection(host: HTMLElement, element: HTMLElement): boolean {
  const paragraph = parentNativeParagraph(element);
  if (!paragraph || !host.contains(paragraph)) return false;
  const range = document.createRange();
  range.selectNodeContents(element);
  const selection = nativeSelection(host);
  if (!selection) return false;
  paragraph.focus({ preventScroll: true });
  selection.removeAllRanges();
  selection.addRange(range);
  return true;
}

export function createNativeTextRange(
  host: HTMLElement,
  paraId: string,
  from: number,
  to: number,
): Range | null {
  const paragraph = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${paraId}"]`);
  if (!paragraph) return null;
  const start = textNodeAt(paragraph, from);
  const end = textNodeAt(paragraph, to);
  const range = document.createRange();
  range.setStart(start.node, start.offset);
  range.setEnd(end.node, end.offset);
  return range;
}

export function createNativeSelectionRange(
  host: HTMLElement,
  start: NativeCursorPoint,
  end: NativeCursorPoint,
): Range | null {
  const startParagraph = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${start.paraId}"]`);
  const endParagraph = host.querySelector<HTMLElement>(`coco-word-native-p[data-para-id="${end.paraId}"]`);
  if (!startParagraph || !endParagraph) return null;
  const startTarget = textNodeAt(startParagraph, start.offset);
  const endTarget = textNodeAt(endParagraph, end.offset);
  const range = document.createRange();
  range.setStart(startTarget.node, startTarget.offset);
  range.setEnd(endTarget.node, endTarget.offset);
  return range;
}

function caretNodeFromPoint(x: number, y: number): { node: Node; offset: number } | null {
  const doc = document as Document & {
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    caretRangeFromPoint?: (x: number, y: number) => Range | null;
  };
  const position = doc.caretPositionFromPoint?.(x, y);
  if (position) return { node: position.offsetNode, offset: position.offset };
  const range = doc.caretRangeFromPoint?.(x, y);
  if (range) return { node: range.startContainer, offset: range.startOffset };
  return null;
}

export function selectNativeImage(host: HTMLElement, selectedImage: NativeSelectedImage): void {
  clearNativeImageSelection(host);
  const image = host.querySelector<HTMLImageElement>(
    `coco-word-native-p[data-para-id="${selectedImage.paraId}"] img[data-drawing-id="${selectedImage.drawingId}"]`,
  );
  image?.classList.add('is-selected');
}

export function clearNativeImageSelection(host: HTMLElement): void {
  host.querySelectorAll('coco-word-native-p img.is-selected').forEach((image) => {
    image.classList.remove('is-selected');
  });
}

export function imageOffset(paragraph: HTMLElement, image: HTMLImageElement): number {
  let offset = 0;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  let current = walker.nextNode();
  while (current) {
    if (current === image) return offset;
    if (current.nodeType === Node.TEXT_NODE) {
      offset += current.textContent?.length ?? 0;
    } else if (current instanceof HTMLImageElement) {
      offset += 1;
    }
    current = walker.nextNode();
  }
  return offset;
}

export function parentNativeParagraph(node: Node): HTMLElement | null {
  let current: Node | null = node.nodeType === Node.ELEMENT_NODE ? node : node.parentNode;
  while (current) {
    if (current instanceof HTMLElement && current.tagName === 'COCO-WORD-NATIVE-P') return current;
    current = current.parentNode;
  }
  return null;
}

function textOffsetWithin(root: HTMLElement, node: Node, offset: number): number {
  let total = 0;
  let found: number | null = null;

  const visit = (current: Node): void => {
    if (found !== null) return;
    if (current === node) {
      if (current.nodeType === Node.TEXT_NODE) {
        found = total + Math.max(0, Math.min(offset, current.textContent?.length ?? 0));
        return;
      }
      const childCount = Math.max(0, Math.min(offset, current.childNodes.length));
      for (let index = 0; index < childCount; index += 1) total += domInlineLength(current.childNodes[index]!);
      found = total;
      return;
    }
    if (current.nodeType === Node.TEXT_NODE) {
      total += current.textContent?.length ?? 0;
      return;
    }
    if (isAtomicInlineElement(current)) {
      total += 1;
      return;
    }
    for (let index = 0; index < current.childNodes.length; index += 1) visit(current.childNodes[index]!);
  };

  visit(root);
  return found ?? total;
}

function domInlineLength(node: Node): number {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent?.length ?? 0;
  if (isAtomicInlineElement(node)) return 1;
  let length = 0;
  for (let index = 0; index < node.childNodes.length; index += 1) length += domInlineLength(node.childNodes[index]!);
  return length;
}

function textNodeAt(root: HTMLElement, offset: number): { node: Node; offset: number } {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  let current = walker.nextNode();
  let remaining = offset;
  while (current) {
    if (current.nodeType === Node.TEXT_NODE) {
      const length = current.textContent?.length ?? 0;
      if (remaining <= length) return { node: current, offset: remaining };
      remaining -= length;
    } else if (isAtomicInlineElement(current)) {
      if (remaining === 0) return elementBoundary(current, 'before');
      if (remaining === 1) return elementBoundary(current, 'after');
      remaining -= 1;
    }
    current = walker.nextNode();
  }
  const fallback = document.createTextNode('');
  root.appendChild(fallback);
  return { node: fallback, offset: 0 };
}

function isAtomicInlineElement(node: Node): node is HTMLElement {
  return node instanceof HTMLElement &&
    (node instanceof HTMLImageElement || node.classList.contains('coco-word-page-break'));
}

function elementBoundary(element: HTMLElement, side: 'before' | 'after'): { node: Node; offset: number } {
  const parent = element.parentNode;
  if (!parent) return { node: element, offset: 0 };
  const index = Array.prototype.indexOf.call(parent.childNodes, element) as number;
  return { node: parent, offset: Math.max(0, index + (side === 'after' ? 1 : 0)) };
}
