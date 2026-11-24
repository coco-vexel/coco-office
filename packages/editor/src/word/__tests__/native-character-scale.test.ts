// @vitest-environment happy-dom

import { afterEach, expect, test, vi } from 'vitest';
import { fromFlatOpc } from '@coco-vexel/coco-office/pack';
import { buildIndices, W14_NS, W_NS } from '@coco-vexel/coco-office/word-view';
import { buildWordEditorViewFromPackage } from '../extract.js';
import { createNativeWordDocument, type NativeWordDocument } from '../native-model.js';
import {
  defineNativeWordElements,
  paginateNativeDocument,
  renderNativeDocument,
} from '../native-renderer.js';

afterEach(() => {
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

test('a scaled DOCX text run reserves its compressed width at the original font size', () => {
  const packageXml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
      <pkg:xmlData><w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>
        <w:p w14:paraId="A0000001"><w:r><w:rPr><w:sz w:val="1034"/><w:w w:val="36"/></w:rPr><w:t>禁止吸烟</w:t></w:r></w:p>
        <w:p w14:paraId="A0000002"><w:r><w:rPr><w:w w:val="100"/></w:rPr><w:t>正常宽度</w:t></w:r></w:p>
        <w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:left="567" w:right="567"/><w:vAlign w:val="center"/></w:sectPr>
      </w:body></w:document></pkg:xmlData>
    </pkg:part>
  </pkg:package>`;
  const pkg = fromFlatOpc(packageXml);
  const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
  const doc = createNativeWordDocument(view);
  defineNativeWordElements();
  vi.spyOn(HTMLElement.prototype, 'offsetWidth', 'get').mockReturnValue(2757);
  const root = renderNativeDocument(doc);
  document.body.appendChild(root);
  paginateNativeDocument(root, doc);

  expect(root.querySelector<HTMLElement>('coco-word-native-page')?.dataset.verticalAlignment).toBe(
    'center',
  );
  const [scaledParagraph, normalParagraph] = root.querySelectorAll('coco-word-native-p');
  const run = scaledParagraph!.querySelector<HTMLElement>('coco-word-native-scaled-run')!;
  expect(run.textContent).toBe('禁止吸烟');
  expect(run.parentElement?.style.fontSize).toBe('517pt');
  expect(run.firstElementChild?.getAttribute('style')).toContain('scaleX(0.36)');
  const usableWidthPx = (16838 - 2 * 567) / 15;
  expect(Number.parseFloat(run.style.width)).toBeCloseTo(2757 * 0.36);
  expect(Number.parseFloat(run.style.width)).toBeLessThan(usableWidthPx);
  expect(normalParagraph!.querySelector('coco-word-native-scaled-run')).toBeNull();
});

test('keeps three exact-height large-print lines on the same centered page', () => {
  const lines = ['国家社会', '公共服务', '平台'];
  const pageLayout = {
    width: 16838,
    height: 11906,
    margin: { top: 567, right: 567, bottom: 567, left: 567 },
    verticalAlignment: 'center' as const,
  };
  const doc: NativeWordDocument = {
    kind: 'word',
    pageLayout,
    sections: [],
    headers: [],
    footers: [],
    blocks: lines.map((text, index) => ({
      kind: 'paragraph',
      paraId: `A000000${index + 1}`,
      story: 'body',
      styleId: null,
      format: { alignment: 'center', spacing: { line: 3583, lineRule: 'exact' } },
      content: [{ kind: 'text', text, format: { fontSize: 172, characterScale: 109 } }],
    })),
  };
  defineNativeWordElements();
  const contentHeight = (pageLayout.height - pageLayout.margin.top - pageLayout.margin.bottom) / 15;
  vi.spyOn(HTMLElement.prototype, 'scrollHeight', 'get').mockImplementation(function (
    this: HTMLElement,
  ) {
    if (this.localName !== 'coco-word-native-page-body') return 0;
    const textHeight = Array.from(this.children).reduce((total, paragraph) => {
      const line = paragraph.firstElementChild as HTMLElement | null;
      const heightPt = Number.parseFloat(line?.style.height ?? '');
      // A browser's exact-height block reserves its specified height. An inline
      // run without a height can expand the paragraph's line box.
      return total + (Number.isFinite(heightPt) ? heightPt * (4 / 3) : 250);
    }, 0);
    return Math.round(Math.max(contentHeight, textHeight));
  });
  const root = renderNativeDocument(doc);
  document.body.appendChild(root);
  paginateNativeDocument(root, doc);

  const pages = root.querySelectorAll<HTMLElement>('coco-word-native-page');
  expect(pages).toHaveLength(1);
  expect(pages[0]!.dataset.verticalAlignment).toBe('center');
  const paragraphs = Array.from(pages[0]!.querySelectorAll('coco-word-native-p'));
  expect(paragraphs.map((paragraph) => paragraph.textContent)).toEqual(lines);
  for (const paragraph of paragraphs) {
    const line = paragraph.firstElementChild as HTMLElement;
    expect(line.style.height).toBe('179.15pt');
    expect(line.style.justifyContent).toBe('center');
  }
});
