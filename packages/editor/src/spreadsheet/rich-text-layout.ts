import type { SpreadsheetTextStyle } from './style-resolver.js';

interface StyledRun {
  text: string;
  style: SpreadsheetTextStyle;
}

interface TextFragment extends StyledRun {
  width: number;
}

interface TextLine {
  fragments: TextFragment[];
  width: number;
  ascent: number;
  descent: number;
  minimumHeight: number;
}

const graphemes = typeof Intl.Segmenter === 'function' ? new Intl.Segmenter(undefined, { granularity: 'grapheme' }) : undefined;

export function layoutRichText(
  ctx: CanvasRenderingContext2D,
  runs: readonly StyledRun[],
  maxWidth: number,
  wrap: boolean,
): Array<TextLine & { height: number; baseline: number }> {
  const lines: TextLine[] = [];
  const metricsByFont = new Map<string, TextMetrics>();
  let line = emptyLine();
  let trailingBreak = false;
  let lastStyle: SpreadsheetTextStyle | undefined;

  const includeMetrics = (style: SpreadsheetTextStyle, textMetrics?: TextMetrics) => {
    let fontMetrics = metricsByFont.get(style.font);
    if (!fontMetrics) {
      fontMetrics = ctx.measureText('国Ag');
      metricsByFont.set(style.font, fontMetrics);
    }
    line.ascent = Math.max(line.ascent, fontMetrics.actualBoundingBoxAscent || style.fontSize, textMetrics?.actualBoundingBoxAscent || 0);
    line.descent = Math.max(line.descent, fontMetrics.actualBoundingBoxDescent || style.fontSize * 0.22, textMetrics?.actualBoundingBoxDescent || 0);
    const fontHeight = (fontMetrics.fontBoundingBoxAscent || 0) + (fontMetrics.fontBoundingBoxDescent || 0);
    line.minimumHeight = Math.max(line.minimumHeight, fontHeight || style.fontSize * 1.2);
  };

  const append = (text: string, style: SpreadsheetTextStyle) => {
    if (!text) return;
    const metrics = ctx.measureText(text);
    includeMetrics(style, metrics);
    line.fragments.push({ text, style, width: metrics.width });
    line.width += metrics.width;
  };

  const nextLine = () => {
    lines.push(line);
    line = emptyLine();
  };

  for (const { text, style } of runs) {
    if (!text) continue;
    lastStyle = style;
    ctx.font = style.font;
    const paragraphs = text.split(/\r\n?|\n/);
    for (const [index, paragraph] of paragraphs.entries()) {
      if (index > 0) {
        if (line.fragments.length === 0) includeMetrics(style);
        nextLine();
        trailingBreak = true;
      }
      if (!paragraph) continue;
      trailingBreak = false;
      if (!wrap) {
        append(paragraph, style);
        continue;
      }
      let fragment = '';
      const characters = graphemes
        ? Array.from(graphemes.segment(paragraph), ({ segment }) => segment)
        : Array.from(paragraph);
      for (const character of characters) {
        const candidate = fragment + character;
        if ((fragment || line.fragments.length > 0) && line.width + ctx.measureText(candidate).width > maxWidth) {
          append(fragment, style);
          nextLine();
          fragment = character;
        } else {
          fragment = candidate;
        }
      }
      append(fragment, style);
    }
  }

  if (trailingBreak && lastStyle) includeMetrics(lastStyle);
  if (line.fragments.length > 0 || trailingBreak) nextLine();
  return lines.map((item) => {
    const height = Math.max(item.minimumHeight, item.ascent + item.descent);
    return { ...item, height, baseline: item.ascent + (height - item.ascent - item.descent) / 2 };
  });
}

function emptyLine(): TextLine {
  return { fragments: [], width: 0, ascent: 0, descent: 0, minimumHeight: 0 };
}
