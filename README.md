# Coco Office

Host-independent Office document SDKs and experimental browser editors. Read,
query, modify and export DOCX, XLSX and PPTX locally using JavaScript.

[中文说明](README.zh-CN.md) · [Contributing](CONTRIBUTING.md) ·
[Release review](docs/open-source-readiness.md)

The project is open source and the three packages are published on npm at
version `1.0.2`. See [LICENSING.md](LICENSING.md) for license scope.

## Packages

| Package                          | Purpose                                                       | Status                |
| -------------------------------- | ------------------------------------------------------------- | --------------------- |
| `@coco-vexel/coco-office`        | OOXML / Flat OPC package operations and Word, Excel, PPT APIs | Published             |
| `@coco-vexel/coco-office-editor` | Word, spreadsheet and fixed-canvas PPT browser editors        | Experimental          |
| `@coco-vexel/coco-office-oojson` | ONLYOFFICE OOJSON ↔ OOXML / Flat OPC conversion               | Apache-2.0; published |

## Architecture boundary

Core has no dependency on Editor, OOJSON, agents or a particular Office host.
Editor uses Core as the document truth layer. Office.js, WPS and ONLYOFFICE host
calls and product orchestration belong in external adapters.

```text
ONLYOFFICE host API
       ↕
@coco-vexel/coco-office-oojson
       ↕
OOXML / Flat OPC
       ↕
@coco-vexel/coco-office
```

## Quick start

Use Node.js 22 or 24 and Corepack. The pnpm version is pinned in `package.json`.

```bash
git clone https://github.com/coco-vexel/coco-office.git
cd coco-office
corepack pnpm install --frozen-lockfile
corepack pnpm build
corepack pnpm example
```

The [runnable example](examples/roundtrip.mjs) generates synthetic Word, Excel and
PowerPoint documents, edits them, exports separate copies, reloads the results
and checks the changed content. Files are written to `examples/output/`.
No Office installation, private repository or external sample is needed.

For an offline evaluation handoff, `corepack pnpm bundle:sdk` builds a ZIP under
`.release-check/sdk/` containing the Core tarball, public SDK documentation and
runnable three-format examples. The same ZIP can be attached to a GitHub Release.

For example, using the Core PPT API:

```js
import { Presentation } from '@coco-vexel/coco-office/ppt';

const deck = await Presentation.fromZip(inputBytes);
try {
  await deck.slides.getItem(0).shapes.getItem('Title').text.setText('Updated title');
  const outputBytes = await deck.toZip();
  // Save outputBytes as a separate file using your application's file API.
} finally {
  await deck.dispose();
}
```

The example assumes the first slide contains a text shape named `Title`. See
[Word getting started](docs/core/getting-started.md), [Excel](docs/core/excel/README.md)
and [PPT](docs/core/ppt/README.md) for their public APIs and constraints.

## Browser demo

After building the packages:

```bash
corepack pnpm demo
```

Open the local URL printed by Vite. Choose a generated sample or open your own
`.docx`, `.xlsx` or `.pptx`, edit it, then select **Export copy**. Document bytes
stay in the browser; this demo has no document-upload backend. Opening another
document replaces the current editing session, so export changes first.

`pnpm build:demo` produces static files in `packages/editor/demo-dist/` for review
or later hosting. Building does not publish them. The demo consumes package
exports rather than importing private products or Core source aliases.

## Capability boundaries

| Area         | Available path                                                         | Limits                                                                                |
| ------------ | ---------------------------------------------------------------------- | ------------------------------------------------------------------------------------- |
| Word         | Paragraph/range edits, selected formatting and structure, OOXML export | Native layout and advanced editing are experimental; export reports unsupported edits |
| Excel        | Cell/range APIs, formulas, workbook edits and spreadsheet UI           | Formula and rich-object support is a subset of Excel; review domain docs              |
| PPT          | Slides, text/shapes/tables, fixed-canvas editing and PPTX export       | Advanced objects and animations have feature-specific preservation/editing limits     |
| Environments | Core browser/Node entries; Editor browser DOM                          | No promise of complete Office application behavior or pixel-identical rendering       |

Package round-trip tests establish content/structure behavior, not visual parity
with Microsoft Office, WPS or ONLYOFFICE. Embedded fonts and external resources
are not supplied by this project.

## Development and checks

```bash
corepack pnpm test:ci
corepack pnpm --filter @coco-vexel/coco-office test:types:excel
corepack pnpm build:demo
corepack pnpm licenses:packages
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for portable suite scope and optional host
checks. The internal regression shell is separate: `pnpm example` then
`pnpm dev:editor`. Corpus paths are configurable; see
[maintainer tools](packages/editor/dev/README.md).

## Community and support

Use the repository issue forms for reproducible bugs, feature proposals,
documentation problems, and focused usage questions. Read
[CONTRIBUTING.md](CONTRIBUTING.md) before opening a pull request. Report
vulnerabilities privately according to [SECURITY.md](SECURITY.md); never attach
customer documents, credentials, or undisclosed exploits to public issues.
Participation is governed by [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).

## npm distribution

All three libraries are published on npm. Install only what you need:

```bash
npm install @coco-vexel/coco-office
npm install @coco-vexel/coco-office-editor
npm install @coco-vexel/coco-office-oojson
```

Editor and OOJSON bring Core as a dependency. OOJSON is ESM JavaScript without
TypeScript declarations in this initial version. See the [OOJSON package guide](packages/oojson/README.md),
[runtime and demo checklist](docs/runtime-and-demos.md), and [npm release procedure](docs/npm-release.md).
The browser demo includes an OOJSON conversion page at `/oojson.html`.

## Source history

`.source-projection.json` records historical import provenance. It is not a
synchronization instruction or a dependency needed to build this checkout.
Changes here must not automatically overwrite an older integration repository.

## License

Original Core/Editor/OOJSON code, synthetic examples and the three self-generated
Excel fixtures use Apache-2.0. See [LICENSING.md](LICENSING.md) for scope.
Dependencies retain their licenses; package distributions carry their notices.
