# Core Excel API Docs

This directory documents the public `@coco-vexel/coco-office/excel` workbook-operation APIs.

The docs follow the same structure as the Word core docs: an entry README, a compact API index, domain-level API specifications, model documentation, and host integration notes. The goal is that humans and small code agents can retrieve only the required domain documents instead of loading the entire Excel surface.

## Recommended Reading Order

1. [api-reference.md](./api-reference.md): public API index and retrieval map.
2. [domains/workbook.md](./domains/workbook.md): `ExcelWorkbook` loading, exporting, styles, names, and lifecycle.
3. [runtime.md](./runtime.md): `ExcelWorkbook` inline/worker execution boundary.
4. [domains/worksheets.md](./domains/worksheets.md): worksheet discovery, structural edits, and sheet-level domains.
5. [domains/ranges-cells.md](./domains/ranges-cells.md): cell and range reads/writes.
6. [domains/tables.md](./domains/tables.md): Excel table objects and table columns.
7. [domains/annotations.md](./domains/annotations.md): hyperlinks, comments, and threaded comments.
8. [domains/formatting-rules.md](./domains/formatting-rules.md): data validations and conditional formats.
9. [domains/drawings.md](./domains/drawings.md): drawings, charts, images, and shapes.
10. [model.md](./model.md): normalized model shape used by package and host bridges.
11. [writeback-planner.md](./writeback-planner.md): host-neutral planning layer for package, sheet-shadow, and diff writeback.
12. [mutation-writeback.md](./mutation-writeback.md): SDK-recorded operation replay for fast host writeback.
13. [host-integration.md](./host-integration.md): Microsoft Excel, WPS, ONLYOFFICE, and desktop writeback guidance.
14. [formula-support.md](./formula-support.md): supported formula families, known boundaries, and coverage audit commands.

## Functional Domains

| Domain | Document | Main Use |
| --- | --- | --- |
| Workbook | [domains/workbook.md](./domains/workbook.md) | Load `.xlsx`, load normalized model, export `.xlsx`, read styles, read defined names. |
| Runtime | [runtime.md](./runtime.md) | Run the same `ExcelWorkbook` API inline or inside a worker through async boundaries. |
| Worksheets | [domains/worksheets.md](./domains/worksheets.md) | List/get sheets, rename, insert/delete rows and columns, access sheet domains. |
| Ranges And Cells | [domains/ranges-cells.md](./domains/ranges-cells.md) | Read and write values, formulas, text, value types, and range navigation. |
| Tables | [domains/tables.md](./domains/tables.md) | Read/edit table metadata, table range, columns, and table style flags. |
| Annotations | [domains/annotations.md](./domains/annotations.md) | Read/edit hyperlinks, comments, and threaded comments. |
| Formatting Rules | [domains/formatting-rules.md](./domains/formatting-rules.md) | Read/edit data validations and conditional format rules. |
| Drawings | [domains/drawings.md](./domains/drawings.md) | Read drawing/chart/image/shape snapshots, create basic charts, and plan host object repair. |
| Model | [model.md](./model.md) | Use or produce `ExcelWorkbookModel` for bridges and tests. |
| Writeback Planner | [writeback-planner.md](./writeback-planner.md) | Convert model diffs and host capabilities into portable writeback intents. |
| Mutation Writeback | [mutation-writeback.md](./mutation-writeback.md) | Record SDK write operations and replay safe mutations through live host APIs before falling back. |
| Host Integration | [host-integration.md](./host-integration.md) | Map package/model operations to host-specific read and writeback mechanisms. |
| Formula Support | [formula-support.md](./formula-support.md) | Track formula coverage without claiming full Excel compatibility. |

## Core Boundary

`@coco-vexel/coco-office/excel` is an OOXML-native workbook operation kernel. Callers provide either `.xlsx` package bytes or a normalized `ExcelWorkbookModel`. Core parses, queries, mutates, and serializes the workbook.

Core does not directly handle:

- Host APIs such as Office.js, WPS JS, ONLYOFFICE plugin APIs, or desktop COM automation.
- File picking, save dialogs, download flows, or local path policy.
- UI confirmation, task panes, ribbons, or plugin lifecycle.
- Business routing, product policy, login state, or model/provider configuration.

## CodeAgent Guidance

1. Generated workbook-operation code should normally assume a ready `workbook: ExcelWorkbook`.
2. Prefer async live-object APIs such as `workbook.worksheets.getItem('Sheet1').getRange('A1:B2')`.
3. Use `ExcelWorkbook` for both inline and worker execution. Configure `mode`, `setWorkerFactory`, or `setWorkerBundleUrl` when the host wants worker isolation.
4. Do not invent host-specific APIs inside core operation code.
5. Treat `toModel()` as a snapshot. Mutating the returned object does not mutate the workbook.
6. Use `toZip()` only when the runtime needs package bytes for host writeback or file output.
7. Retrieve only the domain document needed for the task.

## Maintenance

- `packages/core/src/excel/index.ts` is the export source of truth.
- `packages/core/src/excel/runtime.ts` is the public async live-object API source of truth.
- `packages/core/src/excel/excel-workbook.ts`, `worksheet.ts`, `range.ts`, and `objects.ts` are the synchronous core-kernel source of truth exposed as `*Core` APIs for advanced use and tests.
- `packages/core/src/excel/writeback-planner.ts` is the host-neutral writeback planning source of truth.
- `packages/core/src/excel/model.ts` is the normalized model source of truth.
- Update the matching domain document whenever a public API changes.
