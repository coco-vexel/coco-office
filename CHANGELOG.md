# Changelog

All notable changes to Coco Office are documented in this file.

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

[1.0.2]: https://github.com/coco-vexel/coco-office/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/coco-vexel/coco-office/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/coco-vexel/coco-office/releases/tag/v1.0.0
