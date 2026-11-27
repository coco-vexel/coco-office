// @vitest-environment happy-dom

import { afterEach, expect, test, vi } from 'vitest';
import { fromFlatOpc } from '@coco-vexel/coco-office/pack';
import { buildIndices, W_NS } from '@coco-vexel/coco-office/word-view';
import { buildWordEditorViewFromPackage } from '../extract.js';
import { createNativeWordDocument } from '../native-model.js';
import { paginateNativeDocument, renderNativeDocument } from '../native-renderer.js';

afterEach(() => {
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

test.each([
  { precedingHeight: 270, expectedPage: '1' },
  { precedingHeight: 330, expectedPage: '2' },
])('keeps a bottom-anchored table intact on page $expectedPage', ({ precedingHeight, expectedPage }) => {
  const row = (text: string) => `<w:tr><w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`;
  const pkg = fromFlatOpc(`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
      <pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body>
        <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
        <w:tbl>${row('Attachment')}</w:tbl>
        <w:tbl><w:tblPr><w:tblpPr w:horzAnchor="margin" w:vertAnchor="margin"
          w:tblpXSpec="left" w:tblpYSpec="bottom"/></w:tblPr>${row('Copy')}${row('Publication')}</w:tbl>
        <w:p/>
        <w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="1500" w:bottom="1500"/></w:sectPr>
      </w:body></w:document></pkg:xmlData>
    </pkg:part>
  </pkg:package>`);
  const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
  expect(view.tables.map((table) => table.layout?.anchorToBottomMargin)).toEqual([undefined, true]);

  vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    if (this.matches('coco-word-native-page-body')) return new DOMRect(0, 0, 400, 400);
    if (this.matches('coco-word-native-table')) return new DOMRect(0, 0, 400, this.querySelectorAll('tr').length * 40);
    return new DOMRect(0, 0, 400, this.textContent === 'Body text' ? precedingHeight : 2);
  });
  vi.spyOn(HTMLElement.prototype, 'scrollHeight', 'get').mockImplementation(function (this: HTMLElement) {
    return Array.from(this.children).reduce((height, block) => height + block.getBoundingClientRect().height, 0);
  });
  const doc = createNativeWordDocument(view);
  const root = renderNativeDocument(doc);
  document.body.appendChild(root);
  paginateNativeDocument(root, doc);

  const tables = root.querySelectorAll('coco-word-native-table');
  expect(tables).toHaveLength(2);
  expect(tables[0]!.closest('coco-word-native-page')?.getAttribute('data-page-number')).toBe('1');
  expect(tables[1]!.hasAttribute('data-anchor-bottom-margin')).toBe(true);
  expect(tables[1]!.closest('coco-word-native-page')?.getAttribute('data-page-number')).toBe(expectedPage);
  expect(tables[1]!.querySelectorAll('tr')).toHaveLength(2);
});
