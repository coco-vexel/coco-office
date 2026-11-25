import { expect, test } from 'vitest';
import { fromFlatOpc, getPart, toFlatOpc } from '../../pack/index.js';
import { R_NS, setWordSectionFooters, W_NS, WordDocument } from '../index.js';

test('writes separate footer variants with native fields and removes only replaced SDK parts', async () => {
  const doc = await WordDocument.createBlank({ docId: 'section-footers', mode: 'inline' });
  try {
    const pkg = fromFlatOpc(await doc.toFlatOpc());
    const pageNumber = { font: '宋体', fontSizePt: 14, prefix: '—', suffix: '—' };
    setWordSectionFooters(pkg, 0, {
      differentOddEven: true,
      differentFirstPage: true,
      footers: {
        default: { pageNumber },
        even: { pageNumber: { ...pageNumber, alignment: 'left' } },
        first: {},
      },
    });
    setWordSectionFooters(pkg, 0, { differentFirstPage: false, footers: { default: { pageNumber }, first: null } });
    const roundTrip = fromFlatOpc(toFlatOpc(pkg));
    const main = getPart(roundTrip, 'word/document.xml')!.xmlDom!;
    const refs = Array.from(main.getElementsByTagNameNS(W_NS, 'footerReference'));
    expect(refs.map((ref) => ref.getAttributeNS(W_NS, 'type')).sort()).toEqual(['default', 'even']);
    expect(refs.every((ref) => ref.hasAttributeNS(R_NS, 'id'))).toBe(true);
    expect(main.getElementsByTagNameNS(W_NS, 'titlePg')).toHaveLength(0);
    expect(getPart(roundTrip, 'word/settings.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'evenAndOddHeaders')).toHaveLength(1);
    expect(getPart(roundTrip, 'word/coco-footer-0-default.xml')).toBeUndefined();
    expect(getPart(roundTrip, 'word/coco-footer-0-first.xml')).toBeUndefined();
    const even = getPart(roundTrip, 'word/coco-footer-0-even.xml')!.xmlDom!;
    expect(even.documentElement?.textContent).toBe('—1—');
    expect(even.getElementsByTagNameNS(W_NS, 'fldSimple')[0]?.getAttributeNS(W_NS, 'instr')).toBe(' PAGE ');
  } finally {
    await doc.dispose();
  }
});
