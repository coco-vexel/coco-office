/* eslint-disable no-console -- CLI verification output. */
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { WordDocument } from '@coco-vexel/coco-office/word';
import { ExcelWorkbook } from '@coco-vexel/coco-office/excel';
import { Presentation } from '@coco-vexel/coco-office/ppt';
import { createSample } from './samples.mjs';

const output = new URL('./output/', import.meta.url);
await mkdir(output, { recursive: true });
for (const kind of ['docx', 'xlsx', 'pptx']) {
  const original = await createSample(kind);
  await writeFile(new URL(`sample.${kind}`, output), original);
  let edited;
  if (kind === 'docx') {
    const doc = await WordDocument.fromZip(original, { docId: 'example-word', mode: 'inline' });
    try {
      const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });
      await paragraphs[0].replaceWith('Edited with Coco Office');
      edited = await doc.toZip();
    } finally {
      await doc.dispose();
    }
    const reloaded = await WordDocument.fromZip(edited, { docId: 'verify-word', mode: 'inline' });
    try {
      const { paragraphs } = await reloaded.paragraphs.list({ textFormat: 'preview' });
      assert.equal(await paragraphs[0].getText(), 'Edited with Coco Office');
    } finally {
      await reloaded.dispose();
    }
  } else if (kind === 'xlsx') {
    const book = await ExcelWorkbook.fromZip(original, {
      workbookId: 'example-excel',
      mode: 'inline',
    });
    try {
      await book.worksheets
        .getItem('Sample')
        .getRange('B2')
        .setValues([[42]]);
      edited = await book.toZip();
    } finally {
      await book.dispose();
    }
    const reloaded = await ExcelWorkbook.fromZip(edited, {
      workbookId: 'verify-excel',
      mode: 'inline',
    });
    try {
      assert.equal(await reloaded.worksheets.getItem('Sample').getCell('B2').value(), 42);
    } finally {
      await reloaded.dispose();
    }
  } else {
    const deck = await Presentation.fromZip(original);
    try {
      await deck.slides.getItem(0).shapes.getItem('Title').text.setText('Edited with Coco Office');
      edited = await deck.toZip();
    } finally {
      await deck.dispose();
    }
    const reloaded = await Presentation.fromZip(edited);
    try {
      assert.equal(
        await reloaded.slides.getItem(0).shapes.getItem('Title').text.getText(),
        'Edited with Coco Office',
      );
    } finally {
      await reloaded.dispose();
    }
  }
  await writeFile(new URL(`edited.${kind}`, output), edited);
  console.log(`${kind}: created, edited, exported and reloaded successfully`);
}
