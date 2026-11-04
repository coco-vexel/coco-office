import { describe, expect, it } from 'vitest';
import type { EditorDocumentView, EditorParagraphView } from '../model.js';
import { renderNormalizedDocumentHtmlPage } from '../normalized-html.js';

function paragraph(
  paraId: string,
  text: string,
  format: EditorParagraphView['format'] = {},
): EditorParagraphView {
  return {
    type: 'paragraph',
    paraId,
    story: 'body',
    text,
    format,
    flags: {
      hasComments: false,
      hasRevisions: false,
      hasBookmarks: false,
      hasMath: false,
    },
    segments: [
      {
        kind: 'text',
        from: 0,
        to: text.length,
        text,
        format: { fontFamily: '宋体', fontSize: 12, color: '112233' },
      },
    ],
    spans: [],
  };
}

function documentView(paragraphs: EditorParagraphView[]): EditorDocumentView {
  return {
    kind: 'word',
    sections: [],
    headers: [],
    footers: [],
    blocks: paragraphs,
    paragraphs,
    tables: [],
    detachedComments: [],
    detachedRevisions: [],
  };
}

describe('renderNormalizedDocumentHtmlPage', () => {
  it('renders semantic HTML with stable ids and resolved inline styles', () => {
    const title = paragraph('A1', '客户服务项目复盘报告', {
      outlineLevel: 0,
      alignment: 'center',
      fontFamily: '黑体',
      fontSize: 22,
      bold: true,
    });
    title.styleId = 'Heading1';

    const result = renderNormalizedDocumentHtmlPage(documentView([title]));

    expect(result.html).toContain('<h1 data-para-id="A1"');
    expect(result.html).toContain('text-align: center');
    expect(result.html).toContain('font-size: 22pt');
    expect(result.html).toContain('font-weight: 700');
    expect(result.html).toContain('<span style="font-size: 12pt; font-family:');
    expect(result.html).toContain('color: #112233');
    expect(result.html).not.toContain('Heading1');
    expect(result.html).not.toContain('style-id');
  });

  it('paginates only between top-level blocks', () => {
    const first = paragraph('A1', '第一段');
    const second = paragraph('A2', '第二段');
    const view = documentView([first, second]);

    const page1 = renderNormalizedDocumentHtmlPage(view, { maxChars: 1 });
    const page2 = renderNormalizedDocumentHtmlPage(view, {
      startBlock: page1.nextBlock,
      maxChars: 1,
    });

    expect(page1.html).toContain('data-para-id="A1"');
    expect(page1.html).not.toContain('data-para-id="A2"');
    expect(page1.nextBlock).toBe(1);
    expect(page2.html).toContain('data-para-id="A2"');
    expect(page2.nextBlock).toBeUndefined();
  });
});
