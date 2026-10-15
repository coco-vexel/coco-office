import type { EditorParagraphFormat, EditorRunFormat, EditorTableBorder, EditorTableCellView } from './model.js';

export interface ResolvedStyle {
  declarations: Record<string, string>;
  cssText: string;
}

export function resolveParagraphStyle(format: EditorParagraphFormat | undefined | null): ResolvedStyle {
  const declarations: Record<string, string> = {};
  if (!format) return resolved(declarations);

  applyOutlineDefaults(format, declarations);
  if (format.alignment) declarations['text-align'] = alignmentToCss(format.alignment);
  if (format.indent?.left !== undefined && shouldEmitLeftIndent(format)) {
    declarations['margin-left'] = twipsToPt(format.indent.left);
  }
  if (format.indent?.right !== undefined) declarations['margin-right'] = twipsToPt(format.indent.right);
  if (format.indent?.firstLine !== undefined) declarations['text-indent'] = twipsToPt(format.indent.firstLine);
  if (format.spacing?.before !== undefined) declarations['margin-top'] = twipsToPt(format.spacing.before);
  if (format.spacing?.after !== undefined) declarations['margin-bottom'] = twipsToPt(format.spacing.after);
  const lineHeight = lineHeightToCss(format.spacing?.line, format.spacing?.lineRule, format.snapToGrid);
  if (lineHeight) declarations['line-height'] = lineHeight;
  if (format.fontSize !== undefined) declarations['font-size'] = `${format.fontSize}pt`;
  if (format.fontFamily !== undefined) declarations['font-family'] = fontFamilyToCss(format.fontFamily);
  if (format.characterSpacing !== undefined) declarations['letter-spacing'] = twipsToPt(format.characterSpacing);
  if (format.bold !== undefined) declarations['font-weight'] = format.bold ? '700' : '400';
  // `format.color` comes from paragraph-mark properties (`w:pPr/w:rPr`).
  // It is not the character color for the paragraph text; run-level color is
  // applied by the inline renderer instead.
  const backgroundColor = shadingToCss(format.shading) ?? highlightToCss(format.highlight);
  if (backgroundColor) declarations['background-color'] = backgroundColor;
  applyParagraphBorders(format, declarations);
  return resolved(declarations);
}

function shouldEmitLeftIndent(format: EditorParagraphFormat): boolean {
  return !(format.numbering && format.indent?.left === 0 && format.indent.firstLine === undefined);
}

function applyOutlineDefaults(
  format: EditorParagraphFormat,
  declarations: Record<string, string>,
): void {
  if (format.outlineLevel === undefined) return;
  if (format.outlineLevel === 0) {
    declarations['font-size'] = '16pt';
    declarations['font-weight'] = '700';
    declarations['line-height'] = '1.28';
  } else if (format.outlineLevel === 1) {
    declarations['font-size'] = '14pt';
    declarations['font-weight'] = '700';
    declarations['line-height'] = '1.3';
  } else if (format.outlineLevel === 2) {
    declarations['font-size'] = '12pt';
    declarations['font-weight'] = '700';
    declarations['line-height'] = '1.35';
  }
}

export function resolveRunStyle(format: EditorRunFormat | undefined | null): ResolvedStyle {
  const declarations: Record<string, string> = {};
  if (!format) return resolved(declarations);

  if (format.bold !== undefined) declarations['font-weight'] = format.bold ? '700' : '400';
  if (format.italic !== undefined) declarations['font-style'] = format.italic ? 'italic' : 'normal';
  applyTextDecoration(format, declarations);
  if (format.caps) declarations['text-transform'] = 'uppercase';
  if (format.smallCaps) declarations['font-variant-caps'] = 'small-caps';
  if (format.fontSize !== undefined) declarations['font-size'] = `${format.fontSize}pt`;
  if (format.fontFamily !== undefined) declarations['font-family'] = fontFamilyToCss(format.fontFamily);
  if (format.characterSpacing !== undefined) declarations['letter-spacing'] = twipsToPt(format.characterSpacing);
  if (format.verticalAlign) {
    declarations['vertical-align'] = runVerticalAlignToCss(format.verticalAlign);
    if (format.verticalAlign !== 'baseline' && format.fontSize === undefined) declarations['font-size'] = 'smaller';
  }
  const color = colorToCss(format.color);
  if (color) declarations.color = color;
  const backgroundColor = highlightToCss(format.highlight) ?? shadingToCss(format.shading);
  if (backgroundColor) declarations['background-color'] = backgroundColor;
  return resolved(declarations);
}

export function resolveTableCellStyle(cell: Pick<EditorTableCellView, 'margin' | 'borders' | 'shading' | 'verticalAlignment' | 'noWrap' | 'width'>): ResolvedStyle {
  const declarations: Record<string, string> = {};
  if (cell.width !== undefined && cell.width > 0) declarations.width = twipsToPt(cell.width);
  if (cell.margin) {
    if (cell.margin.top !== undefined) declarations['padding-top'] = twipsToPt(cell.margin.top);
    if (cell.margin.right !== undefined) declarations['padding-right'] = twipsToPt(cell.margin.right);
    if (cell.margin.bottom !== undefined) declarations['padding-bottom'] = twipsToPt(cell.margin.bottom);
    if (cell.margin.left !== undefined) declarations['padding-left'] = twipsToPt(cell.margin.left);
  }
  const fill = colorToCss(cell.shading?.fill);
  if (fill) declarations['background-color'] = fill;
  if (cell.verticalAlignment) {
    declarations['vertical-align'] = tableCellVerticalAlignmentToCss(cell.verticalAlignment);
  }
  if (cell.noWrap) declarations['white-space'] = 'nowrap';
  if (cell.borders) {
    applyTableCellBorder(declarations, 'top', cell.borders.top);
    applyTableCellBorder(declarations, 'right', cell.borders.right);
    applyTableCellBorder(declarations, 'bottom', cell.borders.bottom);
    applyTableCellBorder(declarations, 'left', cell.borders.left);
  }
  return resolved(declarations);
}

export function colorToCss(color: string | undefined): string | null {
  if (!color || color === 'auto') return null;
  return /^[0-9a-f]{6}$/i.test(color) ? `#${color}` : null;
}

function shadingToCss(shading: { fill?: string; val?: string } | undefined): string | null {
  if (!shading || shading.val === 'nil' || shading.val === 'clear' && !shading.fill) return null;
  return colorToCss(shading.fill);
}

function highlightToCss(highlight: string | undefined): string | null {
  if (!highlight || highlight === 'none') return null;
  const named: Record<string, string> = {
    black: '#000000',
    blue: '#0000ff',
    cyan: '#00ffff',
    green: '#00ff00',
    magenta: '#ff00ff',
    red: '#ff0000',
    yellow: '#ffff00',
    white: '#ffffff',
    darkBlue: '#000080',
    darkCyan: '#008080',
    darkGreen: '#008000',
    darkMagenta: '#800080',
    darkRed: '#800000',
    darkYellow: '#808000',
    darkGray: '#808080',
    lightGray: '#c0c0c0',
  };
  return named[highlight] ?? colorToCss(highlight);
}

function resolved(declarations: Record<string, string>): ResolvedStyle {
  return {
    declarations,
    cssText: Object.entries(declarations)
      .map(([property, value]) => `${property}: ${value}`)
      .join('; '),
  };
}

function alignmentToCss(alignment: NonNullable<EditorParagraphFormat['alignment']>): string {
  if (alignment === 'both' || alignment === 'distribute') return 'justify';
  return alignment;
}

function tableCellVerticalAlignmentToCss(
  alignment: NonNullable<EditorTableCellView['verticalAlignment']>,
): string {
  if (alignment === 'center' || alignment === 'both') return 'middle';
  return alignment;
}

function applyTextDecoration(format: EditorRunFormat, declarations: Record<string, string>): void {
  const lines: string[] = [];
  if (format.underline) lines.push('underline');
  if (format.strike || format.doubleStrike) lines.push('line-through');
  if (lines.length > 0) {
    declarations['text-decoration-line'] = lines.join(' ');
    if (format.doubleStrike) declarations['text-decoration-style'] = 'double';
    return;
  }
  if (format.underline === false || format.strike === false || format.doubleStrike === false) {
    declarations['text-decoration-line'] = 'none';
  }
}

function runVerticalAlignToCss(alignment: NonNullable<EditorRunFormat['verticalAlign']>): string {
  if (alignment === 'superscript') return 'super';
  if (alignment === 'subscript') return 'sub';
  return 'baseline';
}

function applyParagraphBorders(format: EditorParagraphFormat, declarations: Record<string, string>): void {
  const borders = format.borders ?? (format.borderBottom ? { bottom: format.borderBottom } : undefined);
  if (!borders) return;
  applyParagraphBorder(declarations, 'top', borders.top);
  applyParagraphBorder(declarations, 'right', borders.right);
  applyParagraphBorder(declarations, 'bottom', borders.bottom ?? borders.between);
  applyParagraphBorder(declarations, 'left', borders.left ?? borders.bar);
}

function applyParagraphBorder(
  declarations: Record<string, string>,
  side: 'top' | 'right' | 'bottom' | 'left',
  border: EditorParagraphFormat['borderBottom'] | undefined,
): void {
  const css = borderToCss(border);
  if (!css) return;
  declarations[`border-${side}`] = css;
  declarations[`padding-${side}`] = border?.space !== undefined ? `${border.space}pt` : '4pt';
}

function borderToCss(border: EditorParagraphFormat['borderBottom']): string | null {
  if (!border || border.val === 'none' || border.val === 'nil') return null;
  const color = colorToCss(border.color) ?? '#000000';
  const widthPt = border.size !== undefined ? Math.max(0.75, border.size / 8) : 0.75;
  const style = border.val === 'double' ? 'double' : 'solid';
  return `${widthPt}pt ${style} ${color}`;
}

function applyTableCellBorder(
  declarations: Record<string, string>,
  side: 'top' | 'right' | 'bottom' | 'left',
  border: EditorTableBorder | undefined,
): void {
  if (!border) return;
  const css = tableBorderToCss(border);
  declarations[`border-${side}`] = css ?? '0';
}

function tableBorderToCss(border: EditorTableBorder): string | null {
  if (border.val === 'none' || border.val === 'nil') return null;
  const color = colorToCss(border.color) ?? '#000000';
  const widthPt = border.size !== undefined ? Math.max(0.5, border.size / 8) : 0.75;
  const style = border.val === 'double' ? 'double' : 'solid';
  return `${widthPt}pt ${style} ${color}`;
}

function lineHeightToCss(
  line: number | undefined,
  lineRule: 'auto' | 'exact' | 'atLeast' | undefined,
  snapToGrid: boolean | undefined,
): string | null {
  if (line === undefined) return null;
  if (lineRule === 'exact' || lineRule === 'atLeast') return twipsToPt(line);
  if (snapToGrid !== true) return String(line / 240);
  return `calc(var(--coco-word-doc-grid-line-height, 1em) * ${line / 240})`;
}

function twipsToPt(twips: number): string {
  return `${twips / 20}pt`;
}

function fontFamilyToCss(fontFamily: string): string {
  const families = fontFamily
    .split(',')
    .map((family) => family.trim())
    .filter((family) => family.length > 0);
  const quotedFamilies = families.map((family) =>
    `"${family.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\r\n\f]/g, ' ')}"`,
  );
  return [...quotedFamilies, 'SimSun', '"Songti SC"', '"Noto Serif CJK SC"', 'serif'].join(', ');
}
