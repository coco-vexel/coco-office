import { W_NS } from './namespaces.js';

export type WordTableBorderStyle =
  'nil' | 'none' | 'single' | 'double' | 'dotted' | 'dashed' | 'dotDash' | 'dotDotDash';

export interface WordTableBorder {
  val: WordTableBorderStyle;
  /** 6-digit RGB color without '#', or 'auto'. */
  color?: string;
  /** Border width in eighths of a point. */
  size?: number;
  /** Border spacing in points. */
  space?: number;
}

export interface WordTableBorders {
  top?: WordTableBorder;
  right?: WordTableBorder;
  bottom?: WordTableBorder;
  left?: WordTableBorder;
  insideH?: WordTableBorder;
  insideV?: WordTableBorder;
}

const BORDER_SIDES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;
const TABLE_BORDER_FOLLOWING_PROPERTIES = new Set([
  'shd',
  'tblLayout',
  'tblCellMar',
  'tblLook',
  'tblCaption',
  'tblDescription',
  'tblPrChange',
]);
const CELL_BORDER_FOLLOWING_PROPERTIES = new Set([
  'shd',
  'noWrap',
  'tcMar',
  'textDirection',
  'tcFitText',
  'vAlign',
  'hideMark',
  'headers',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
]);

/**
 * Writes table or cell borders using the single Core OOXML implementation.
 * Callers own target lookup and dirty-part tracking.
 */
export function writeWordTableBorders(
  properties: Element,
  kind: 'table' | 'cell',
  borders: WordTableBorders,
): void {
  const containerName = kind === 'table' ? 'tblBorders' : 'tcBorders';
  let container = firstWordChild(properties, containerName);
  if (!container) {
    container = properties.ownerDocument!.createElementNS(W_NS, `w:${containerName}`);
    const following =
      kind === 'table' ? TABLE_BORDER_FOLLOWING_PROPERTIES : CELL_BORDER_FOLLOWING_PROPERTIES;
    properties.insertBefore(container, firstFollowingWordChild(properties, following));
  }

  for (const side of BORDER_SIDES) {
    const spec = borders[side];
    if (!spec) continue;
    let border = firstWordChild(container, side);
    if (!border) {
      border = properties.ownerDocument!.createElementNS(W_NS, `w:${side}`);
      container.appendChild(border);
    }
    border.setAttributeNS(W_NS, 'w:val', spec.val);
    setOptionalAttribute(border, 'color', spec.color?.toUpperCase());
    setOptionalAttribute(border, 'sz', spec.size);
    setOptionalAttribute(border, 'space', spec.space);
  }
}

function firstWordChild(parent: Element, localName: string): Element | null {
  for (let index = 0; index < parent.childNodes.length; index += 1) {
    const child = parent.childNodes[index];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      (child as Element).localName === localName
    ) {
      return child as Element;
    }
  }
  return null;
}

function firstFollowingWordChild(parent: Element, localNames: Set<string>): Element | null {
  for (let index = 0; index < parent.childNodes.length; index += 1) {
    const child = parent.childNodes[index];
    if (
      child?.nodeType === 1 &&
      (child as Element).namespaceURI === W_NS &&
      localNames.has((child as Element).localName)
    ) {
      return child as Element;
    }
  }
  return null;
}

function setOptionalAttribute(
  element: Element,
  localName: string,
  value: string | number | undefined,
): void {
  if (value === undefined) element.removeAttributeNS(W_NS, localName);
  else element.setAttributeNS(W_NS, `w:${localName}`, String(value));
}
