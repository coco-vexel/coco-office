# @coco-vexel/coco-office

Read, query, modify and export Word (`.docx`), Excel (`.xlsx`) and PowerPoint (`.pptx`) documents in JavaScript. Core supplies document/package APIs; it does not render a full Office application or call Office host APIs.

## Installation and requirements

```bash
npm install @coco-vexel/coco-office
```

These instructions target the initial npm release; availability is subject to publication. Use Node.js 22 or 24 for the examples below, saved as `.mjs` files. ESM, CommonJS and TypeScript declarations are included. Browser applications should use a bundler and supply file selection/download handling. Neither Microsoft Office nor a backend is required for local document operations.

## Quick start: create a Word document

Save as `word.mjs`, then run `node word.mjs`. No input document is required. It creates `hello.docx` in the current directory; `wx` prevents overwriting an existing file.

```js
import { writeFile } from 'node:fs/promises';
import { WordDocument } from '@coco-vexel/coco-office/word';

const doc = await WordDocument.createBlank({ docId: 'hello', mode: 'inline' });
try {
  const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });
  await paragraphs[0].replaceWith('Hello from Coco Office');
  await writeFile('hello.docx', await doc.toZip(), { flag: 'wx' });
} finally {
  await doc.dispose();
}
```

For an existing DOCX, use `WordDocument.fromZip(await readFile(inputPath), { docId: 'edit', mode: 'inline' })`. Write the returned bytes to a separate output path. `paragraphs.list()` supports pagination; request `pagination: { mode: 'all' }` when the operation needs all paragraphs.

## Edit an Excel workbook

Save as `excel.mjs`; run `node excel.mjs input.xlsx Sheet1`. The named sheet must already exist. This writes `edited.xlsx` and leaves the input file unchanged. The example writes a literal value, not a formula.

```js
import { readFile, writeFile } from 'node:fs/promises';
import { ExcelWorkbook } from '@coco-vexel/coco-office/excel';

const [, , inputPath, sheetName = 'Sheet1'] = process.argv;
if (!inputPath) throw new Error('Usage: node excel.mjs input.xlsx [sheetName]');
const book = await ExcelWorkbook.fromZip(await readFile(inputPath), {
  workbookId: 'example',
  mode: 'inline',
});
try {
  await book.worksheets
    .getItem(sheetName)
    .getRange('B2')
    .setValues([[42]]);
  await writeFile('edited.xlsx', await book.toZip(), { flag: 'wx' });
} finally {
  await book.dispose();
}
```

## Edit a PowerPoint text shape

Save as `ppt.mjs`; run `node ppt.mjs input.pptx Title`. The first slide must contain a text shape with the supplied name. This produces `edited.pptx`; it does not render slide images.

```js
import { readFile, writeFile } from 'node:fs/promises';
import { Presentation } from '@coco-vexel/coco-office/ppt';

const [, , inputPath, shapeName = 'Title'] = process.argv;
if (!inputPath) throw new Error('Usage: node ppt.mjs input.pptx [shapeName]');
const deck = await Presentation.fromZip(await readFile(inputPath));
try {
  await deck.slides.getItem(0).shapes.getItem(shapeName).text.setText('Updated title');
  await writeFile('edited.pptx', await deck.toZip(), { flag: 'wx' });
} finally {
  await deck.dispose();
}
```

## Public entrypoints

| Import                                                            | Purpose                                                      |
| ----------------------------------------------------------------- | ------------------------------------------------------------ |
| `@coco-vexel/coco-office/pack`                                    | OOXML ZIP/Flat OPC package I/O                               |
| `@coco-vexel/coco-office/word`                                    | Word document, paragraphs/ranges and supported edits         |
| `@coco-vexel/coco-office/word-view`                               | Word projection/index helpers for views                      |
| `@coco-vexel/coco-office/excel`                                   | Worksheets, cells/ranges, formula subset and workbook export |
| `@coco-vexel/coco-office/ppt`                                     | Slides, text/shapes/tables and presentation export           |
| `@coco-vexel/coco-office/shared`                                  | Shared utilities                                             |
| `@coco-vexel/coco-office/kernel-worker`, `/kernel-worker-browser` | Word Node/browser worker bootstraps                          |
| `@coco-vexel/coco-office/excel-worker`, `/excel-worker-browser`   | Excel Node/browser worker bootstraps                         |

ESM example: `import { WordDocument } from '@coco-vexel/coco-office/word'`.
CommonJS example: `const { WordDocument } = require('@coco-vexel/coco-office/word')`.
Use domain entrypoints for the APIs you need; do not import `src/` or undocumented `dist/` paths.

## Lifecycle, browser use and limitations

Call `dispose()` in `finally`, including when an edit/export fails. The examples use inline mode, which requires no Worker setup. Node Worker mode requires `WordDocument.setWorkerBundleUrl(...)` or `ExcelWorkbook.setWorkerBundleUrl(...)`; browser Workers require the corresponding browser factory. Do not load Node Worker bootstraps as browser entrypoints. Inline mode can block the UI for large documents; profile representative files before choosing a worker configuration.

In a browser, read `new Uint8Array(await file.arrayBuffer())` and pass it to `fromZip`. Download output with a Blob/object URL; loading and saving files are application responsibilities. To display/edit documents visually, use `@coco-vexel/coco-office-editor`.

Word layout, Excel formulas/rich objects, and PPT advanced objects/animations have feature-specific limits. Unknown-part preservation and export/reparse tests do not guarantee all edits are supported or pixel-identical rendering in Office. A missing sheet/shape, malformed input or unsupported operation should be handled by the caller; never discard the original file after an unchecked export.

## Documentation

| Area       | Start here                                                                                                | Detailed reference                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| ---------- | --------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| General    | [Getting started](https://github.com/coco-vexel/coco-office/blob/main/docs/core/getting-started.md)       | [API index](https://github.com/coco-vexel/coco-office/blob/main/docs/core/api-reference.md) and [cookbook](https://github.com/coco-vexel/coco-office/blob/main/docs/core/cookbook.md)                                                                                                                                                                                                                                                                                                       |
| Word       | [Word documentation](https://github.com/coco-vexel/coco-office/blob/main/docs/core/README.md)             | [Paragraphs](https://github.com/coco-vexel/coco-office/blob/main/docs/core/domains/paragraphs.md), [ranges](https://github.com/coco-vexel/coco-office/blob/main/docs/core/domains/ranges.md), [tables](https://github.com/coco-vexel/coco-office/blob/main/docs/core/domains/tables.md), [comments](https://github.com/coco-vexel/coco-office/blob/main/docs/core/domains/comments.md), and [revisions](https://github.com/coco-vexel/coco-office/blob/main/docs/core/domains/revisions.md) |
| Excel      | [Excel documentation](https://github.com/coco-vexel/coco-office/blob/main/docs/core/excel/README.md)      | [API reference](https://github.com/coco-vexel/coco-office/blob/main/docs/core/excel/api-reference.md) and [formula support](https://github.com/coco-vexel/coco-office/blob/main/docs/core/excel/formula-support.md)                                                                                                                                                                                                                                                                         |
| PowerPoint | [PowerPoint documentation](https://github.com/coco-vexel/coco-office/blob/main/docs/core/ppt/README.md)   | [API reference](https://github.com/coco-vexel/coco-office/blob/main/docs/core/ppt/api-reference.md) and [style capabilities](https://github.com/coco-vexel/coco-office/blob/main/docs/core/ppt/style-capabilities.md)                                                                                                                                                                                                                                                                       |
| Runtime    | [Runtime, tests and demos](https://github.com/coco-vexel/coco-office/blob/main/docs/runtime-and-demos.md) | [Runtime modes](https://github.com/coco-vexel/coco-office/blob/main/docs/core/advanced/runtime.md) and [host integration](https://github.com/coco-vexel/coco-office/blob/main/docs/core/advanced/host-integration.md)                                                                                                                                                                                                                                                                       |

GitHub Releases also provide a self-contained SDK bundle with this package,
offline documentation and runnable Word/Excel/PowerPoint examples.

## Runnable repository demo

For synthetic three-format fixtures and round-trip assertions, clone the repository, run `corepack pnpm install --frozen-lockfile`, `corepack pnpm build`, then `corepack pnpm example`. These are repository commands, not scripts shipped with the installed npm package.

Report reproducible SDK issues with package version and a minimal synthetic document at [GitHub Issues](https://github.com/coco-vexel/coco-office/issues). Follow [SECURITY.md](https://github.com/coco-vexel/coco-office/blob/main/SECURITY.md) for private vulnerability reports.

Apache-2.0 for original code; see LICENSE and THIRD_PARTY_NOTICES.md in the package.
