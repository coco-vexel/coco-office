import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
import { W14_NS, W_NS } from '../namespaces.js';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

async function makeFixture(bodyInner: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeDoc(bodyInner: string, docId: string): Promise<WordDocument> {
  return WordDocument.fromZip(await makeFixture(bodyInner), { docId, mode: 'inline' });
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('Word domain cache', () => {
  test('paragraph list cache builds lazily only for the cacheable full-body profile', async () => {
    const doc = await makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>one</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>two</w:t></w:r></w:p>`,
      'cache-lazy',
    );

    expect(doc._debugCacheSnapshot().paragraphs.state).toBe('empty');

    await doc.paragraphs.list();
    expect(doc._debugCacheSnapshot().paragraphs.state).toBe('empty');

    const first = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    const afterFirst = doc._debugCacheSnapshot();
    expect(await paragraphTexts(first)).toEqual(['one', 'two']);
    expect(afterFirst.paragraphs.state).toBe('ready');
    expect(afterFirst.paragraphs.listAllBodyReady).toBe(true);
    expect(afterFirst.paragraphs.buildCount).toBe(1);

    await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(doc._debugCacheSnapshot().paragraphs.buildCount).toBe(1);

    await doc.dispose();
  });

  test('paragraph cache is invalidated after text mutation and rebuilds to current content', async () => {
    const doc = await makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>old</w:t></w:r></w:p>
       <w:p w14:paraId="AAAAAAA2"><w:r><w:t>stable</w:t></w:r></w:p>`,
      'cache-apply',
    );

    await doc.paragraphs.list({ pagination: { mode: 'all' } });
    await doc.plan().replaceParagraph('AAAAAAA1', 'new').apply();

    const dirtySnapshot = doc._debugCacheSnapshot();
    expect(dirtySnapshot.revision).toBe(1);
    expect(dirtySnapshot.paragraphs.state).toBe('dirty');

    const cachedRead = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await paragraphTexts(cachedRead)).toEqual(['new', 'stable']);

    doc._debugInvalidateCache('paragraphs', 'test rescan');
    const rescanned = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(cachedRead).toEqual(rescanned);

    await doc.dispose();
  });

  test('failed transaction leaves cached paragraph data unchanged', async () => {
    const doc = await makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>keep</w:t></w:r></w:p>`,
      'cache-rollback',
    );

    await doc.paragraphs.list({ pagination: { mode: 'all' } });
    const before = doc._debugCacheSnapshot();

    await expect(
      doc.plan().replaceParagraph('AAAAAAA1', 'mutated').replaceParagraph('NOPE', 'x').apply(),
    ).rejects.toBeInstanceOf(WordError);

    expect(doc._debugCacheSnapshot()).toEqual(before);
    expect(await (await doc.paragraphs.get('AAAAAAA1'))?.getText()).toBe('keep');

    await doc.dispose();
  });

  test('live paragraphs read and write the current document state', async () => {
    const doc = await makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>original</w:t></w:r></w:p>`,
      'paragraph-object',
    );

    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    const paragraph = list.paragraphs[0]!;
    expect(paragraph.paraId).toBe('AAAAAAA1');
    expect(paragraph.id).toBe('AAAAAAA1');
    expect(await paragraph.getText()).toBe('original');

    await paragraph.replaceWith('changed by object');

    expect(await (await doc.paragraphs.get('AAAAAAA1'))?.getText()).toBe('changed by object');

    await doc.dispose();
  });

  test('format-only mutation keeps text cache usable', async () => {
    const doc = await makeDoc(
      `<w:p w14:paraId="AAAAAAA1"><w:r><w:t>hello typo</w:t></w:r></w:p>`,
      'cache-format',
    );

    await doc.paragraphs.list({ pagination: { mode: 'all' } });
    await doc
      .plan()
      .formatTextInParagraph('AAAAAAA1', 'typo', { bold: true, color: 'FF0000' })
      .apply();

    const afterApply = doc._debugCacheSnapshot();
    expect(afterApply.revision).toBe(1);
    expect(afterApply.paragraphs.state).toBe('ready');
    expect(afterApply.paragraphs.buildCount).toBe(1);

    const list = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    expect(await list.paragraphs[0]?.getText()).toBe('hello typo');
    expect(doc._debugCacheSnapshot().paragraphs.buildCount).toBe(1);

    await doc.dispose();
  });

  test('table list cache builds lazily for default body top-level tables', async () => {
    const doc = await makeDoc(
      `${table1x1('AAAAAAA1', 'cell-a')}
       ${table1x1('AAAAAAA2', 'cell-b')}
       <w:p w14:paraId="AAAAAAAF"><w:r><w:t>end</w:t></w:r></w:p>`,
      'cache-table-lazy',
    );

    expect(doc._debugCacheSnapshot().tables.state).toBe('empty');

    const first = await doc.tables.list();
    const afterFirst = doc._debugCacheSnapshot();
    expect(first.map((table) => table.tableId)).toEqual(['tbl-0001', 'tbl-0002']);
    expect(afterFirst.tables.state).toBe('ready');
    expect(afterFirst.tables.cachedProfiles).toEqual(['basic']);
    expect(afterFirst.tables.buildCount).toBe(1);

    await doc.tables.list();
    expect(doc._debugCacheSnapshot().tables.buildCount).toBe(1);

    await doc.dispose();
  });

  test('table cache returns defensive copies for nested DTO fields', async () => {
    const doc = await makeDoc(
      `${table1x1('AAAAAAA1', 'cell-a')}
       <w:p w14:paraId="AAAAAAAF"><w:r><w:t>end</w:t></w:r></w:p>`,
      'cache-table-copy',
    );

    const list = await doc.tables.list({ includeCellTexts: true });
    list[0]!.cells![0]![0] = 'external mutation';
    list[0]!.cellGrid![0]![0]!.imageDrawingIds.push('external-drawing');

    const again = await doc.tables.list({ includeCellTexts: true });
    expect(again[0]?.cells?.[0]?.[0]).toBe('cell-a');
    expect(again[0]?.cellGrid?.[0]?.[0]?.imageDrawingIds).toEqual([]);

    await doc.dispose();
  });

  test('table cache is invalidated after table content mutation and rebuilds current cells', async () => {
    const doc = await makeDoc(
      `${table1x1('AAAAAAA1', 'old-cell')}
       <w:p w14:paraId="AAAAAAAF"><w:r><w:t>end</w:t></w:r></w:p>`,
      'cache-table-apply',
    );

    await doc.tables.list({ includeCellTexts: true });
    await doc
      .plan()
      .replaceTableCellText({
        tableId: 'tbl-0001',
        row: 0,
        col: 0,
        text: 'new-cell',
        onAnchorConflict: 'error',
      })
      .apply();

    const dirtySnapshot = doc._debugCacheSnapshot();
    expect(dirtySnapshot.tables.state).toBe('dirty');

    const cachedRead = await doc.tables.list({ includeCellTexts: true });
    expect(cachedRead[0]?.cells?.[0]?.[0]).toBe('new-cell');

    doc._debugInvalidateCache('tables', 'test rescan');
    const rescanned = await doc.tables.list({ includeCellTexts: true });
    expect(cachedRead).toEqual(rescanned);

    await doc.dispose();
  });
});

function table1x1(paraId: string, text: string): string {
  return `<w:tbl>
  <w:tblPr/>
  <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>`;
}

async function paragraphTexts(list: Awaited<ReturnType<WordDocument['paragraphs']['list']>>) {
  return Promise.all(list.paragraphs.map((paragraph) => paragraph.getText()));
}
