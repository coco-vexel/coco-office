# Changelog

All notable changes to Coco Office are documented in this file.

## [1.0.8] - 2026-11-09

### Fixed

- Excluded text-box story paragraphs from the main Word document-body paragraph
  listing while preserving explicit text-box story access.

## [1.0.7] - 2026-11-06

### Added

- Added portable worksheet layout, print setup and drawing-line mutations for
  generating host-compatible spreadsheet templates such as copybooks.
- Added standard spreadsheet editing, keyboard navigation and rich clipboard
  transfers across the browser editor runtime.
- Added Word table border APIs and editor gridline support.

### Fixed

- Preserved blank cell styles, worksheet print settings and spreadsheet
  geometry across OOXML round trips and Office host rendering.
- Applied paragraph formatting across selected Word table cells.

## [1.0.6] - 2026-11-04

### Added

- Added normalized semantic HTML rendering for Word documents so agents can
  inspect effective paragraph and run formatting without relying on style IDs.
- Exposed stable Word paragraph identities and document-level sandbox helpers
  for structured formatting inspection and execution.

### Fixed

- Preserved valid empty-body paragraph insertion while tightening structured
  paragraph operation handling in direct and worker execution modes.

## [1.0.5] - 2026-11-02

### Fixed

- Rendered legacy Word VML horizontal lines with their original width, stroke
  color and weight in the browser editor.
- Added category-aware Chinese font fallbacks so Heiti, Kaiti, Fangsong and
  Songti document fonts retain a closer visual match across platforms.

## [1.0.4] - 2026-10-31

### Fixed

- Converted ISO 8601 SpreadsheetML date cells to Excel serial values so they
  render, sort and calculate as dates instead of appearing blank.
- Added 1904 date-system preservation and common built-in date/time number
  formats, including correct handling of fractional-day values.

## [1.0.3] - 2026-10-29

### Added

- Added structural `moveRows` and `moveColumns` Excel APIs that preserve cell
  values, formulas, styles and row or column metadata while updating supported
  dependent references.
- Documented structural move semantics and CodePlan guidance so whole-row and
  whole-column reordering does not rebuild the worksheet data matrix.

## [1.0.2] - 2026-10-23

### Added

- Added complete Word revision modeling and accept/reject operations across the
  SDK and editor runtime.
- Added a consistent read-only mode for the Word, Spreadsheet and PPT editors,
  including runtime mutation guards and narrow PPT preview support.

### Fixed

- Preserved untouched XLSX package content when applying cell edits through
  targeted OOXML patching.
- Improved PPT inherited placeholder properties, effective text styles and
  thumbnail fidelity.
- Preserved Word page width when the annotation rail is visible.

### Security

- Hardened document URL handling and other untrusted document input paths.

## [1.0.1] - 2026-10-15

### Security

- Raised the minimum `prosemirror-view` version to `1.42.6` and `jszip` to
  `3.10.2`.
- Refreshed package third-party notices after dependency updates.

### Maintenance

- Updated compatible build and test dependencies.
- Updated GitHub Actions to current major versions.

## [1.0.0] - 2026-10-15

- Initial open-source release of Core, Editor and OOJSON packages.

[1.0.8]: https://github.com/coco-vexel/coco-office/compare/v1.0.7...v1.0.8
[1.0.7]: https://github.com/coco-vexel/coco-office/compare/v1.0.6...v1.0.7
[1.0.6]: https://github.com/coco-vexel/coco-office/compare/v1.0.5...v1.0.6
[1.0.5]: https://github.com/coco-vexel/coco-office/compare/v1.0.4...v1.0.5
[1.0.4]: https://github.com/coco-vexel/coco-office/compare/v1.0.3...v1.0.4
[1.0.3]: https://github.com/coco-vexel/coco-office/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/coco-vexel/coco-office/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/coco-vexel/coco-office/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/coco-vexel/coco-office/releases/tag/v1.0.0
