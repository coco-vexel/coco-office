# Coco Native Word Dev Tools

For the standalone public demo, run `pnpm build` then `pnpm demo` from the repository root.
It generates samples in the browser and accepts local DOCX/XLSX/PPTX files.

These maintainer tools use `pnpm example` output for Word/PPT fixtures by default.
Set `COCO_DOCX_FIXTURE_DIR`, `COCO_PPTX_FIXTURE_DIR`, and optionally
`COCO_EXTRA_DOCX_FIXTURE` to use a local corpus. Fixed index suites below refer to
the maintainer corpus and are not portable CI acceptance tests.

These tools validate native Word rendering for content coverage and layout risk signals. They do not enforce pixel-perfect parity with Microsoft Word.

## Native Regression

Run native browser rendering against fixture documents:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor dev:regression -- --start-server --port=5189 --fixtures=0 --screenshot-pages=1
```

Run the fixed P0 baseline suite:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor dev:regression -- --start-server --port=5189 --suite=word-native-p0 --no-screenshots
```

Run the extra rendering closure suite:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor dev:regression -- --start-server --port=5190 --suite=word-native-p1-extra --no-screenshots
```

Outputs:

- `native-regression-output/latest-report.json`
- `native-regression-output/latest-report.html`
- native browser page screenshots

The report separates acceptance from visual review:

- `PASS`: content coverage is complete and there are no layout review notes.
- `REVIEW`: content coverage is complete, but there is page count drift, a layout risk signal, or a missing Word screenshot. This is acceptable for long documents when no key content is lost.
- `FAIL`: parsed content is not rendered, console errors occur, or a page body overflows.

Content coverage includes body paragraphs, tables, body images, anchored images, body drawings, text boxes, numbering, page breaks, comments, and revisions. Headers and footers are checked through their own visibility and anchor metrics because Word files can contain unused header/footer templates. Word reference screenshots are for human review; page count drift is not treated as pixel-diff failure.

The `word-native-p0` suite uses fixture indices instead of name matching so it stays stable when similarly named files are added. It covers complex product docs, image-heavy docs, long CJK docs, 200+ page docs, rowspan/table pagination, generated stress coverage, docGrid/line-height behavior, tracked revisions, text boxes/red-head layouts, and comments.

The `word-native-p1-extra` suite is a rendering-closure smoke suite for generated formal documents, vMerge tables, page-specific headers/footers, Unicode headers/footers, complex numbering, different first-page headers/footers, and comment replies.

## Microsoft Word Reference

Generate Microsoft Word reference images for visual comparison:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor dev:word-reference -- --input="../../examples/output/sample.docx" --pages=3
```

This opens the document through Microsoft Word automation in read-only mode, exports a PDF, and renders PNG pages under `native-regression-output/word-reference/`.

Word reference screenshots are for human review. They should guide fidelity work, but they are not pixel-diff assertions.

## PPT Edit Roundtrip

Run the browser PPT editor through real toolbar controls, export PPTX bytes, reparse the export through Core, and verify the exported file opens through Microsoft PowerPoint COM:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor dev:ppt-edit-roundtrip -- --base-url=http://127.0.0.1:5178 --fixture=0
```

Outputs are written under `native-regression-output/ppt-edit-roundtrip/`:

- edited PPTX
- browser screenshot
- JSON summary with selected object ids, edited values, DOM style samples, and `verifyExport()` slide count
- PowerPoint COM open verification for the exported PPTX unless `--no-office-verify` is provided

The Office verification may start a PowerPoint process, but it does not drive the desktop with mouse or keyboard input.

`dev:ppt-editor-smoke:fast` is the inner-loop check for PPT editor work. It keeps one high-frequency edit sample, one interaction fixture, one fixture content audit, and no Office or screenshot gallery work by default:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor dev:ppt-editor-smoke:fast -- --start-server --port=5178
```

`dev:ppt-editor-smoke:quick` also runs a small Office export smoke by default. It reuses this round-trip runner for representative cases (`0:default,13:group-child`) so quick editor regressions catch exported PPTX files that Core can reparse but PowerPoint cannot open. Use `--no-office-export-smoke` to skip that check or `--office-export-smoke-cases=0:default,13:group-child` to choose cases.

`dev:ppt-editor-smoke:full` broadens the sampled fixtures and includes the parent-group drag export case (`13:group-parent-drag`) in the Office smoke matrix.

## PPT Reference Cross Check

Run browser PPT rendering against a presentation-app PNG export:

```bash
corepack pnpm --filter @coco-vexel/coco-office-editor dev:ppt-reference-cross-check -- --start-server --port=5178 --fixtures=12 --reference=auto
```

The default reference is Microsoft PowerPoint automation because it is usually more stable in the local regression loop. Use `dev:ppt-office-cross-check` or `--reference=office` as a PowerPoint-friendly alias when WPS is unstable locally. `--reference=wps` keeps the strict WPS-only path. `--reference=auto` tries WPS first and falls back to PowerPoint when WPS COM is busy, rejected, or opens a repair-blocked/empty presentation. Reference PNGs are written under each fixture's `reference/` directory, and legacy `wps` fields are kept in JSON rows for older analysis helpers. The summary records the reference app used per fixture. When both reference apps cannot export a fixture, the run still captures browser-side render diagnostics and marks the fixture as `referenceUnavailable` instead of treating it as a browser rendering failure.

Runs with review slides also write `review-contact-sheet.jpg`, pairing reference and browser screenshots for the highest-difference slides in each fixture. Use `--contact-sheet-per-fixture=N` to tune how many review slides are included per fixture, or `--no-contact-sheet` for machine-only runs.

Each run keeps its timestamped output directory and also updates `ppt-reference-cross-check/latest-summary.json` plus `ppt-reference-cross-check/latest-review-contact-sheet.jpg` when a contact sheet is generated.

The pixel mean / weighted difference is a review sorter, not a pass/fail gate for the PPT editor. The MVP target is structural fidelity: no missing slides, no dropped unsupported images, no empty browser canvases, no scene/DOM node-count drift, and no major rendered-text coverage loss. The cross-check summary therefore reports `structural.status` and `structural.issues` separately from `review`; treat `structural.status: "REVIEW"` as an implementation risk, while `review` slides are candidates for human visual inspection of font metrics, gradients, preset shapes, and other non-structural differences. `layout.status` is a softer review signal for browser text nodes whose rendered content overflows a clipped box, which helps find wrapping/resize regressions without equating every native PPT text-clipping case with a structural failure.

## Line Grid Fixture

Generate a focused line-height/docGrid fixture:

```bash
python packages/editor/dev/create-line-grid-fixture.py
corepack pnpm --filter @coco-vexel/coco-office-editor dev:word-reference -- --input="../../examples/output/sample.docx" --pages=8
corepack pnpm --filter @coco-vexel/coco-office-editor dev:regression -- --start-server --port=5289 --fixtures=line_grid --no-screenshots
```

This fixture isolates default paragraph style inheritance, `docGrid`, explicit `snapToGrid`, and table pagination behavior.
