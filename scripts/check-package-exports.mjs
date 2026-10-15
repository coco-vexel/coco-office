// Copy into a fresh consumer project with the three tarballs installed, then run.
/* eslint-disable no-console -- CLI verification output. */
import assert from 'node:assert/strict';
import { readFile, access } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { WordDocument } from '@coco-vexel/coco-office/word';
import { ExcelWorkbook } from '@coco-vexel/coco-office/excel';

const require = createRequire(import.meta.url);
for (const name of ['@coco-vexel/coco-office', '@coco-vexel/coco-office-editor', '@coco-vexel/coco-office-oojson']) {
  const metadataPath = require.resolve(`${name}/package.json`);
  const root = dirname(metadataPath);
  const pkg = JSON.parse(await readFile(metadataPath, 'utf8'));
  assert.equal(pkg.license, 'Apache-2.0');
  assert.notEqual(pkg.private, true);
  assert.equal(pkg.publishConfig.access, 'public');
  if (name === '@coco-vexel/coco-office') {
    const notices = await readFile(resolve(root, 'THIRD_PARTY_NOTICES.md'), 'utf8');
    assert(
      notices.includes('Jean-loup Gailly and Mark Adler'),
      'embedded zlib attribution must ship',
    );
    assert(notices.includes('License selected for Coco Office distribution: MIT'));
  }
  assert(!JSON.stringify(pkg.dependencies ?? {}).includes('workspace:'));
  for (const file of ['LICENSE', 'README.md', 'THIRD_PARTY_NOTICES.md'])
    await access(resolve(root, file));
  for (const target of Object.values(pkg.exports)) {
    for (const path of typeof target === 'string' ? [target] : Object.values(target))
      await access(resolve(root, path));
  }
  console.log(`${name}: package metadata, notices and export targets verified`);
}
for (const entry of ['pack', 'word', 'word-view', 'excel', 'ppt', 'shared']) {
  assert(Object.keys(await import(`@coco-vexel/coco-office/${entry}`)).length > 0);
  assert(Object.keys(require(`@coco-vexel/coco-office/${entry}`)).length > 0);
}
const doc = await WordDocument.createBlank({ docId: 'worker-source', mode: 'inline' });
const bytes = await doc.toZip();
await doc.dispose();
WordDocument.setWorkerBundleUrl(require.resolve('@coco-vexel/coco-office/kernel-worker'));
const workerDoc = await WordDocument.fromZip(bytes, { docId: 'worker-consumer', mode: 'worker' });
try {
  const { paragraphs } = await workerDoc.paragraphs.list({ pagination: { mode: 'all' } });
  await paragraphs[0].replaceWith('Node worker roundtrip');
  const reloaded = await WordDocument.fromZip(await workerDoc.toZip(), {
    docId: 'worker-reload',
    mode: 'inline',
  });
  try {
    assert.equal(
      await (await reloaded.paragraphs.list({})).paragraphs[0].getText(),
      'Node worker roundtrip',
    );
  } finally {
    await reloaded.dispose();
  }
} finally {
  await workerDoc.dispose();
}
const { createSample } = await import('./examples/samples.mjs');
ExcelWorkbook.setWorkerBundleUrl(require.resolve('@coco-vexel/coco-office/excel-worker'));
const workerBook = await ExcelWorkbook.fromZip(await createSample('xlsx'), {
  workbookId: 'worker-excel',
  mode: 'worker',
});
try {
  await workerBook.worksheets.getItem('Sample').getCell('B2').setValue(99);
  const reloaded = await ExcelWorkbook.fromZip(await workerBook.toZip(), {
    workbookId: 'verify-worker-excel',
    mode: 'inline',
  });
  try {
    assert.equal(await reloaded.worksheets.getItem('Sample').getCell('B2').value(), 99);
  } finally {
    await reloaded.dispose();
  }
} finally {
  await workerBook.dispose();
}
console.log('Core ESM/CJS imports and Node Word/Excel worker export/reload verified');

await import('./examples/oojson-roundtrip.mjs');
