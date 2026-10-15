import { Schema, type Node as ProseMirrorNode } from 'prosemirror-model';
import { baseKeymap } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import { EditorState, Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type {
  EditorBlockView,
  EditorDocumentView,
  EditorParagraphFormat,
  EditorParagraphView,
  EditorRunFormat,
  EditorTableCellView,
  EditorTableView,
} from './model.js';

export interface WordProjection {
  doc: ProseMirrorNode;
  decorations: DecorationSet;
  paragraphContentPos: Map<string, number>;
}

export interface CreateWordEditorStateOptions {
  plugins?: Plugin[];
}

export const wordEditorSchema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: {
        paraId: { default: '' },
        story: { default: 'body' },
        styleId: { default: null },
        format: { default: null },
      },
      parseDOM: [{ tag: 'p' }],
      toDOM: (node) => ['p', paragraphDomAttrs(node.attrs), 0],
    },
    table: {
      group: 'block',
      content: 'table_row+',
      attrs: {
        tableId: { default: '' },
        columnWidths: { default: null },
      },
      tableRole: 'table',
      toDOM: (node) => [
        'div',
        { class: 'coco-word-table-scroll' },
        ['table', tableDomAttrs(node.attrs), ['tbody', 0]],
      ],
    },
    table_row: {
      content: 'table_cell+',
      tableRole: 'row',
      toDOM: () => ['tr', 0],
    },
    table_cell: {
      content: 'block+',
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        width: { default: null },
        shading: { default: null },
        verticalAlignment: { default: null },
      },
      tableRole: 'cell',
      toDOM: (node) => ['td', tableCellDomAttrs(node.attrs), 0],
    },
    image: {
      inline: true,
      group: 'inline',
      atom: true,
      attrs: {
        src: {},
        alt: { default: '' },
        width: { default: null },
        height: { default: null },
        drawingId: { default: null },
      },
      toDOM: (node) => ['img', imageDomAttrs(node.attrs)],
    },
    text: { group: 'inline' },
  },
  marks: {
    strong: {
      parseDOM: [{ tag: 'strong' }, { tag: 'b' }],
      toDOM: () => ['strong', 0],
    },
    em: {
      parseDOM: [{ tag: 'em' }, { tag: 'i' }],
      toDOM: () => ['em', 0],
    },
    underline: {
      parseDOM: [{ tag: 'u' }],
      toDOM: () => ['u', 0],
    },
    textStyle: {
      attrs: {
        color: { default: null },
        fontSize: { default: null },
      },
      parseDOM: [{ tag: 'span' }],
      toDOM: (mark) => ['span', textStyleDomAttrs(mark.attrs), 0],
    },
  },
});

export const wordDecorationsKey = new PluginKey<DecorationSet>('coco-word-decorations');
export const wordStructureNormalizationKey = new PluginKey('coco-word-structure-normalization');

export function createWordEditorState(
  view: EditorDocumentView,
  options: CreateWordEditorStateOptions = {},
): EditorState {
  const projection = projectWordViewToProseMirror(view);
  return EditorState.create({
    schema: wordEditorSchema,
    doc: projection.doc,
    plugins: [
      createWordDecorationsPlugin(projection.decorations),
      keymap(baseKeymap),
      createWordStructureNormalizationPlugin(),
      ...(options.plugins ?? []),
    ],
  });
}

export function projectWordViewToProseMirror(view: EditorDocumentView): WordProjection {
  const paragraphContentPos = new Map<string, number>();
  const blocks = view.blocks.map((block) => blockToNode(block));
  const doc = wordEditorSchema.nodes.doc.create(null, blocks);

  doc.descendants((node, pos) => {
    if (node.type === wordEditorSchema.nodes.paragraph && typeof node.attrs.paraId === 'string') {
      paragraphContentPos.set(node.attrs.paraId, pos + 1);
    }
  });

  const decorations: Decoration[] = [];
  for (const paragraph of view.paragraphs) {
    const base = paragraphContentPos.get(paragraph.paraId)!;
    for (const span of paragraph.spans) {
      if (span.to <= span.from) {
        if (span.kind === 'comment') {
          decorations.push(createCommentWidget(base + span.from, span));
        }
        continue;
      }
      const className =
        span.kind === 'comment'
          ? 'coco-word-comment'
          : span.revisionKind === 'delete'
            ? 'coco-word-revision-delete'
            : 'coco-word-revision-insert';
      decorations.push(
        Decoration.inline(
          base + span.from,
          base + span.to,
          {
            class: className,
            ...(span.kind === 'comment' ? commentDomAttrs(span) : {}),
          },
          {
            id: span.id,
            kind: span.kind,
            paraId: paragraph.paraId,
          },
        ),
      );
    }
  }

  return {
    doc,
    decorations: DecorationSet.create(doc, decorations),
    paragraphContentPos,
  };
}

function createCommentWidget(pos: number, span: { id: string; commentId: number; body: string }) {
  return Decoration.widget(
    pos,
    () => {
      const marker = document.createElement('span');
      marker.className = 'coco-word-comment-widget';
      marker.dataset.commentId = String(span.commentId);
      marker.title = commentTitle(span);
      marker.textContent = '!';
      return marker;
    },
    {
      id: span.id,
      kind: 'comment',
      side: 1,
    },
  );
}

function commentDomAttrs(span: { commentId: number; body: string }): Record<string, string> {
  return {
    'data-comment-id': String(span.commentId),
    title: commentTitle(span),
  };
}

function commentTitle(span: { commentId: number; body: string }): string {
  const body = span.body.trim();
  return body.length > 0 ? `Comment ${span.commentId}: ${body}` : `Comment ${span.commentId}`;
}

function blockToNode(block: EditorBlockView): ProseMirrorNode {
  return block.type === 'table' ? tableToNode(block) : paragraphToNode(block);
}

function paragraphToNode(paragraph: EditorParagraphView): ProseMirrorNode {
  return wordEditorSchema.nodes.paragraph.create(
    {
      paraId: paragraph.paraId,
      story: paragraph.story,
      styleId: paragraph.styleId ?? null,
      format: paragraph.format ?? null,
    },
    paragraph.segments.length > 0 ? paragraph.segments.flatMap((segment) => {
      if (segment.kind === 'placeholder') {
        return [wordEditorSchema.text('\uFFFC')];
      }
      if (segment.kind === 'drawing') {
        return [wordEditorSchema.text(segment.text && segment.text.length > 0 ? segment.text : '\uFFFC')];
      }
      if (segment.kind === 'field') {
        return [wordEditorSchema.text(segment.cachedText ?? '\uFFFC', marksForRunFormat(segment.format))];
      }
      if (segment.kind === 'image') {
        return [
          wordEditorSchema.nodes.image.create({
            src: segment.src,
            alt: segment.drawingId,
            width: segment.width ?? null,
            height: segment.height ?? null,
            drawingId: segment.drawingId,
          }),
        ];
      }
      if (segment.text.length === 0) return [];
      const format = segment.kind === 'text' ? segment.format : undefined;
      return [wordEditorSchema.text(segment.text, marksForRunFormat(format))];
    }) : paragraph.text.length > 0 ? wordEditorSchema.text(paragraph.text) : undefined,
  );
}

function tableToNode(table: EditorTableView): ProseMirrorNode {
  return wordEditorSchema.nodes.table.create(
    { tableId: table.tableId, columnWidths: table.columnWidths },
    table.rows.map((row) =>
      wordEditorSchema.nodes.table_row.create(
        null,
        row.cells
          .filter((cell) => cell.vMerge !== 'continue')
          .map((cell) => tableCellToNode(cell)),
      ),
    ),
  );
}

function tableCellToNode(cell: EditorTableCellView): ProseMirrorNode {
  const paragraphs =
    cell.paragraphs.length > 0
      ? cell.paragraphs.map((paragraph) => paragraphToNode(paragraph))
      : [wordEditorSchema.nodes.paragraph.create(null)];
  return wordEditorSchema.nodes.table_cell.create(
    {
      colspan: Math.max(1, cell.gridSpan),
      rowspan: Math.max(1, cell.rowSpan ?? 1),
      width: cell.width ?? null,
      shading: cell.shading ?? null,
      verticalAlignment: cell.verticalAlignment ?? null,
    },
    paragraphs,
  );
}

function marksForRunFormat(format: EditorRunFormat | undefined) {
  if (!format) return undefined;
  const marks = [];
  if (format.bold) marks.push(wordEditorSchema.marks.strong.create());
  if (format.italic) marks.push(wordEditorSchema.marks.em.create());
  if (format.underline) marks.push(wordEditorSchema.marks.underline.create());
  if (format.color || format.fontSize) {
    marks.push(
      wordEditorSchema.marks.textStyle.create({
        color: format.color ?? null,
        fontSize: format.fontSize ?? null,
      }),
    );
  }
  return marks.length > 0 ? marks : undefined;
}

function paragraphDomAttrs(attrs: {
  paraId?: unknown;
  story?: unknown;
  styleId?: unknown;
  format?: unknown;
}): Record<string, string> {
  const domAttrs: Record<string, string> = {};
  if (typeof attrs.paraId === 'string' && attrs.paraId.length > 0) {
    domAttrs['data-para-id'] = attrs.paraId;
  }
  if (typeof attrs.story === 'string' && attrs.story.length > 0) {
    domAttrs['data-story'] = attrs.story;
  }
  if (typeof attrs.styleId === 'string' && attrs.styleId.length > 0) {
    domAttrs['data-style-id'] = attrs.styleId;
  }

  const format = isParagraphFormat(attrs.format) ? attrs.format : null;
  if (format?.outlineLevel !== undefined) {
    domAttrs['data-outline-level'] = String(format.outlineLevel);
  }

  const style = paragraphStyle(format);
  if (style.length > 0) domAttrs.style = style;
  return domAttrs;
}

function tableDomAttrs(attrs: { tableId?: unknown; columnWidths?: unknown }): Record<string, string> {
  const domAttrs: Record<string, string> = {};
  if (typeof attrs.tableId === 'string' && attrs.tableId.length > 0) {
    domAttrs['data-table-id'] = attrs.tableId;
  }
  return domAttrs;
}

function tableCellDomAttrs(attrs: {
  colspan?: unknown;
  rowspan?: unknown;
  width?: unknown;
  shading?: unknown;
  verticalAlignment?: unknown;
}): Record<string, string> {
  const domAttrs: Record<string, string> = {};
  if (typeof attrs.colspan === 'number' && attrs.colspan > 1) {
    domAttrs.colspan = String(attrs.colspan);
  }
  if (typeof attrs.rowspan === 'number' && attrs.rowspan > 1) {
    domAttrs.rowspan = String(attrs.rowspan);
  }
  const style: string[] = [];
  if (typeof attrs.width === 'number' && attrs.width > 0) {
    style.push(`width: ${twipsToPt(attrs.width)}`);
  }
  if (isShading(attrs.shading)) {
    const fill = colorToCss(attrs.shading.fill);
    if (fill) style.push(`background-color: ${fill}`);
  }
  if (typeof attrs.verticalAlignment === 'string') {
    style.push(`vertical-align: ${attrs.verticalAlignment === 'center' ? 'middle' : attrs.verticalAlignment}`);
  }
  if (style.length > 0) domAttrs.style = style.join('; ');
  return domAttrs;
}

function imageDomAttrs(attrs: {
  src?: unknown;
  alt?: unknown;
  width?: unknown;
  height?: unknown;
  drawingId?: unknown;
}): Record<string, string> {
  const domAttrs: Record<string, string> = {
    src: typeof attrs.src === 'string' ? attrs.src : '',
    alt: typeof attrs.alt === 'string' ? attrs.alt : '',
  };
  if (typeof attrs.drawingId === 'string') domAttrs['data-drawing-id'] = attrs.drawingId;
  const style: string[] = [];
  if (typeof attrs.width === 'number' && attrs.width > 0) {
    style.push(`width: ${attrs.width}px`);
  }
  if (typeof attrs.height === 'number' && attrs.height > 0) {
    style.push(`height: ${attrs.height}px`);
  }
  if (style.length > 0) domAttrs.style = style.join('; ');
  return domAttrs;
}

function isShading(value: unknown): value is { fill?: string } {
  return value !== null && typeof value === 'object';
}

function textStyleDomAttrs(attrs: { color?: unknown; fontSize?: unknown }): Record<string, string> {
  const style: string[] = [];
  const color = typeof attrs.color === 'string' ? colorToCss(attrs.color) : null;
  if (color) style.push(`color: ${color}`);
  if (typeof attrs.fontSize === 'number') style.push(`font-size: ${attrs.fontSize}pt`);
  return style.length > 0 ? { style: style.join('; ') } : {};
}

function isParagraphFormat(value: unknown): value is EditorParagraphFormat {
  return value !== null && typeof value === 'object';
}

function paragraphStyle(format: EditorParagraphFormat | null): string {
  if (!format) return '';
  const declarations: string[] = [];

  if (format.alignment) declarations.push(`text-align: ${alignmentToCss(format.alignment)}`);
  if (format.indent?.left !== undefined) declarations.push(`margin-left: ${twipsToPt(format.indent.left)}`);
  if (format.indent?.right !== undefined) declarations.push(`margin-right: ${twipsToPt(format.indent.right)}`);
  if (format.indent?.firstLine !== undefined) {
    declarations.push(`text-indent: ${twipsToPt(format.indent.firstLine)}`);
  }
  if (format.spacing?.before !== undefined) {
    declarations.push(`margin-top: ${twipsToPt(format.spacing.before)}`);
  }
  if (format.spacing?.after !== undefined) {
    declarations.push(`margin-bottom: ${twipsToPt(format.spacing.after)}`);
  }
  const lineHeight = lineHeightToCss(format.spacing?.line, format.spacing?.lineRule);
  if (lineHeight) declarations.push(`line-height: ${lineHeight}`);
  if (format.fontSize !== undefined) declarations.push(`font-size: ${format.fontSize}pt`);
  if (format.bold !== undefined) declarations.push(`font-weight: ${format.bold ? '700' : '400'}`);
  const borderBottom = borderToCss(format.borderBottom);
  if (borderBottom) {
    declarations.push(`border-bottom: ${borderBottom}`);
    declarations.push(`padding-bottom: ${format.borderBottom?.space ?? 4}pt`);
  }

  return declarations.join('; ');
}

function alignmentToCss(alignment: NonNullable<EditorParagraphFormat['alignment']>): string {
  if (alignment === 'both' || alignment === 'distribute') return 'justify';
  return alignment;
}

function borderToCss(border: EditorParagraphFormat['borderBottom']): string | null {
  if (!border || border.val === 'none' || border.val === 'nil') return null;
  const color = colorToCss(border.color) ?? '#000000';
  const widthPt = border.size !== undefined ? Math.max(0.75, border.size / 8) : 0.75;
  const style = border.val === 'double' ? 'double' : 'solid';
  return `${widthPt}pt ${style} ${color}`;
}

function lineHeightToCss(
  line: number | undefined,
  lineRule: 'auto' | 'exact' | 'atLeast' | undefined,
): string | null {
  if (line === undefined) return null;
  if (lineRule === 'exact' || lineRule === 'atLeast') return twipsToPt(line);
  return String(line / 240);
}

function twipsToPt(twips: number): string {
  return `${twips / 20}pt`;
}

function colorToCss(color: string | undefined): string | null {
  if (!color || color === 'auto') return null;
  return /^[0-9a-f]{6}$/i.test(color) ? `#${color}` : null;
}

function createWordDecorationsPlugin(decorations: DecorationSet): Plugin {
  return new Plugin({
    key: wordDecorationsKey,
    state: {
      init: () => decorations,
      apply: (tr, set) => set.map(tr.mapping, tr.doc),
    },
    props: {
      decorations(state) {
        return wordDecorationsKey.getState(state);
      },
    },
  });
}

function createWordStructureNormalizationPlugin(): Plugin {
  return new Plugin({
    key: wordStructureNormalizationKey,
    appendTransaction(transactions, _oldState, newState) {
      if (!transactions.some((tr) => tr.docChanged)) return null;
      if (transactions.some((tr) => tr.getMeta(wordStructureNormalizationKey))) return null;

      const seen = new Set<string>();
      let tr = newState.tr;
      let changed = false;
      newState.doc.descendants((node, pos) => {
        if (node.type !== wordEditorSchema.nodes.paragraph) return true;
        const paraId = typeof node.attrs.paraId === 'string' ? node.attrs.paraId : '';
        if (paraId && !seen.has(paraId)) {
          seen.add(paraId);
          return true;
        }

        const nextParaId = nextUniqueParaId(seen);
        seen.add(nextParaId);
        tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, paraId: nextParaId });
        changed = true;
        return true;
      });

      return changed ? tr.setMeta(wordStructureNormalizationKey, { normalized: true }) : null;
    },
  });
}

function nextUniqueParaId(seen: Set<string>): string {
  let paraId = '';
  do {
    paraId = Math.floor(Math.random() * 0xffffffff)
      .toString(16)
      .toUpperCase()
      .padStart(8, '0');
  } while (!paraId || seen.has(paraId));
  return paraId;
}
