# @coco-vexel/coco-office-oojson

ONLYOFFICE-shaped JSON ↔ Word OOXML / Flat OPC conversion, plus PPT content/deck adapters. This is a conversion library, not an ONLYOFFICE server, editor or Excel JSON converter.

## Install and runtime

After the initial npm release: `npm install @coco-vexel/coco-office-oojson`.

ESM JavaScript for Node.js and browser bundlers. It depends on `@coco-vexel/coco-office`. The initial release has no TypeScript declaration files or CommonJS entry. `@coco-vexel/coco-office-oojson/browser-global` ships an IIFE exposing `oojson` for a script host; it is not an ESM API entry. No Office application is required for pure conversion.

```js
import { jsonToOoxml, ooxmlToJson, jsonToFlatOpc } from '@coco-vexel/coco-office-oojson';
const document = {
  type: 'document',
  content: [
    {
      type: 'paragraph',
      pPr: {},
      content: [{ type: 'run', rPr: {}, content: ['Hello'] }],
    },
  ],
};
const xml = jsonToOoxml(document);
const { doc, gaps } = ooxmlToJson(xml);
const flatOpc = jsonToFlatOpc(document);
```

## Create a DOCX and inspect writeback diagnostics

If you directly import Core as in this example, declare it in your application too:

```bash
npm install @coco-vexel/coco-office-oojson @coco-vexel/coco-office
```

Save the following as `convert.mjs`, then run `node convert.mjs` with Node.js 22 or 24. It needs no input file and writes `converted.docx`. The exclusive write flag prevents overwriting an existing file. Browser bundlers can use the same conversion calls, replacing Node file I/O with Blob downloads.

```js
import { writeFile } from 'node:fs/promises';
import {
  jsonToFlatOpc,
  jsonToOoxml,
  ooxmlToJson,
  flatOpcToOnlyOfficeWriteback,
} from '@coco-vexel/coco-office-oojson';
import { fromFlatOpc, toZip } from '@coco-vexel/coco-office/pack';

const input = {
  type: 'document',
  content: [
    {
      type: 'paragraph',
      pPr: {},
      content: [{ type: 'run', rPr: {}, content: ['Hello OOJSON'] }],
    },
  ],
};
const { doc, gaps } = ooxmlToJson(jsonToOoxml(input));
if (gaps.size) throw new Error(`Unmapped XML: ${JSON.stringify(gaps.top())}`);
const flatOpc = jsonToFlatOpc(input);
const result = flatOpcToOnlyOfficeWriteback(flatOpc);
if (result.diagnostics.unsupportedDomains.length) {
  throw new Error(`Unsupported writeback: ${JSON.stringify(result.diagnostics.unsupportedItems)}`);
}
console.log('Returned document:', doc);
console.log('Writeback warnings:', result.diagnostics.warnings);
await writeFile('converted.docx', await toZip(fromFlatOpc(flatOpc)), { flag: 'wx' });
```

`jsonToOoxml()` returns a Word `document.xml` string; it does not return ZIP bytes. `jsonToFlatOpc()` returns an XML package string. `toZip(fromFlatOpc(...))` creates DOCX bytes. `ooxmlToJson()` returns `{ doc, gaps }`, not the document directly. JSON conversion and writeback planning are synchronous; ZIP I/O is asynchronous.

`flatOpcToOnlyOfficeWriteback()` prepares a body plus gap information and diagnostics; it does not mutate a live editor. A separate host adapter must apply these using the correct ONLYOFFICE host APIs and report failures. `gaps.size === 0` is not proof that every unsupported package feature can be written back; inspect both conversion gaps and writeback diagnostics.

For a script-tag host, copy the installed `dist/oojson.global.js` to your static assets and load it with `<script src="/assets/oojson.global.js"></script>`; functions are available on `globalThis.oojson`. Do not expect named ESM exports from the `/browser-global` entry. Host scripts and DOM access remain your application's responsibility.

## API groups and limits

- Word body: `jsonToOoxml`, `ooxmlToJson`, `Gaps`; paragraphs, runs, tables and supported formatting. Private extension data preserves otherwise unmapped JSON fields where implemented; this does not guarantee arbitrary DOCX fidelity.
- Word package: `jsonToFlatOpc`, `jsonToFlatOpcWithStyles`, `buildStylesXml`, `buildNumberingXml`, default theme/font/settings parts and `withPackageLongTailParts`.
- Writeback preparation: `flatOpcToOnlyOfficeWriteback` returns `bodyJson`, `gaps` and `diagnostics` (`warnings`, `unsupportedDomains`, `unsupportedItems`). A host adapter must apply the body and check diagnostics; the library does not call ONLYOFFICE APIs itself.
- PPT: ONLYOFFICE text content ↔ Core text bodies; deck ↔ Core models; supported text/style/model updates. Advanced geometry, charts, animations and layout are not covered by a blanket fidelity claim.
- Long-tail limits include complex drawings/wrapping, threaded comments, picture bullets and theme effects. See [implementation notes](https://github.com/coco-vexel/coco-office/blob/main/packages/oojson/TODO.md).

## Runnable verification

From the repository root after install/build:

```bash
node examples/oojson-roundtrip.mjs
corepack pnpm --filter @coco-vexel/coco-office-oojson test:portable
corepack pnpm demo
```

Open `/oojson.html` from the demo URL for a synthetic text conversion and round-trip check. The Node example also crosses Core Flat OPC/ZIP and writeback boundaries. Neither requires DocumentBuilder or a private corpus.

Optional host fidelity checks use `DOCBUILDER_EXE` and `test:ooxml`; a live ONLYOFFICE editor and its external adapter are needed to validate actual ToJSON/FromJSON execution. These are not prerequisites for installing this npm package.

Apache-2.0; see LICENSE and THIRD_PARTY_NOTICES.md. Core is a separate runtime dependency; its attribution is also included for the bundled browser-global artifact.

Report reproducible conversion issues with package version, synthetic input JSON and diagnostics at [GitHub Issues](https://github.com/coco-vexel/coco-office/issues). Do not attach customer files or account metadata.
