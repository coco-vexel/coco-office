# Native Word Editor Capability Matrix

This matrix describes the active custom-element Native runtime. It separates live editing, rendering, and DOCX writeback so host integrations do not infer export support from a visual-only feature.

| Area | Live browser/Electron editor | DOCX writeback | Verification |
| --- | --- | --- | --- |
| Text, paragraphs, selection, IME | Supported | Supported for ordinary text, split/merge and field-safe insertion | Unit, runtime, and P0 regression |
| Run formatting | Supported | Supported for common run formatting and format revisions | Round-trip tests |
| Lists, numbering, spacing, alignment | Supported | Basic/common cases supported | Native tests and P1 samples |
| Images | Inline insert, copy/paste, delete, resize, alignment, undo/redo | Ordinary inline insert, delete, resize, and tracked insert resolution supported | Round-trip tests and image fixture |
| Floating drawings and advanced wrapping | Rendered conservatively where parsed | Not a general editable/writeback contract | Image-heavy fixture; no destructive fallback |
| Tables | Regular row/column edits, cell editing, formatting, merge/split, selection, undo/redo | Regular structures, cell content/format, selected merge/split and table revisions supported | Round-trip tests and table-heavy P0/P1 samples |
| Comments | Add, display, locate, edit, reply, delete, undo/redo | Body/header anchors and basic reply metadata supported | Round-trip tests and comment fixture |
| Revisions | Track insert/delete/format/table actions; individual and batch accept/reject | Text, format, image, table, cross-paragraph and story-part cases covered | Runtime, round-trip, and revision fixture |
| Headers/footers | Render and direct paragraph editing | Text, formatting, fields, images, comments and basic revisions supported | Header/footer round-trip tests and P1 samples |
| Fields and page breaks | Render and preserve basic fields/page breaks | Field-safe text insertion and page-break handling covered | Round-trip and runtime tests |
| Hyperlinks | External/document-anchor rendering and commands | External relationship writeback supported; bookmark projection remains limited | Round-trip tests |
| Find/replace | Body and nested-table text | Changes flow through normal text change records | Search tests |

## Conservative boundaries

- Complex floating layout, footnotes, advanced section behavior, bookmark-aware anchor editing, and unsupported drawing types are preserved or skipped rather than rewritten destructively.
- Complex merged-grid structural edits remain disabled when the original structure cannot be represented safely.
- Collaborative editing and remote-operation reconciliation are outside the current package contract.

## Host contract

Browser and Electron hosts use the same `renderNativeWordEditor` runtime. Hosts own file acquisition, persistence, download/IPC, and toolbar UI; the editor exposes the live model, command state, change records, and lifecycle events.
