# Excel Runtime APIs

This domain covers `ExcelWorkbook` runtime mode, worker isolation, script execution, export, and disposal.

Use this document when the task says: worker mode, inline mode, runtime, execute generated Excel code, isolate large workbook work, avoid UI freeze, script timeout, or dispose workbook.

## Skill Summary

`ExcelWorkbook` is the public async root object. It follows the same runtime shape as `WordDocument`: callers use one live-object API surface, and `mode` decides whether the core kernel runs inline in the current thread or inside a worker.

The synchronous parser/mutator kernel is exposed as `ExcelWorkbookCore` for advanced integrations and unit tests. Generated workbook-operation code should prefer the public async `ExcelWorkbook` unless it is intentionally testing the low-level kernel.

## Runtime Contract

The runtime contract is intentionally Word-like:

| Contract | Meaning |
| --- | --- |
| One public API | Inline and worker mode expose the same `ExcelWorkbook` live-object API. Calling code must not fork by runtime mode. |
| Async operation boundary | Any operation that resolves workbook state, mutates workbook state, serializes output, or checks existence is async. |
| Sync handle boundary | Methods that only create handles are sync and cheap: `getItem`, `getRange`, `getCell`, range navigation, and object-domain `getItem`. |
| Worker isolation | In worker mode, the `ExcelWorkbookCore` instance lives in the worker. Public live-object calls are RPC operations. |
| Structured data only | Arguments and return values crossing the worker boundary must be structured-cloneable. |
| No leaked core objects | Public live-object handles do not expose `ExcelWorkbookCore`, `ExcelWorksheetCore`, `ExcelRangeCore`, or object-domain `*Core` instances. |
| Explicit disposal | The caller that creates an `ExcelWorkbook` should call `dispose()` when finished. |

This means generated SDK code should look the same in both modes:

```ts
const workbook = await ExcelWorkbook.fromZip(bytes, { mode: 'auto' });
const sheet = workbook.worksheets.getItem('Sheet1');
await sheet.getRange('A1:B2').setValues([[1, 2], [3, 4]]);
const output = await workbook.toZip();
await workbook.dispose();
```

Do not write generated code that reaches into core internals:

```ts
// Wrong for public SDK operation code.
const core = await ExcelWorkbookCore.fromZip(bytes);
core.worksheets.getItem('Sheet1').getCell('A1').setValue('x');
```

## Load APIs

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `ExcelWorkbook.fromZip(bytes, opts?)` | Load `.xlsx` bytes into an inline or worker workbook. | `Uint8Array \| ArrayBuffer`, `ExcelWorkbookLoadOptions?` | `Promise<ExcelWorkbook>` |
| `ExcelWorkbook.fromModel(model, opts?)` | Load normalized model into an inline or worker workbook. | `ExcelWorkbookModel`, `ExcelWorkbookLoadOptions?` | `Promise<ExcelWorkbook>` |
| `ExcelWorkbook.fromPackage(pkg, opts?)` | Load an already parsed package into an inline workbook. | `Package`, `ExcelWorkbookLoadOptions?` | `ExcelWorkbook` |

`fromPackage` is inline-only because a parsed package object is not a stable structured-clone boundary. Use `fromZip` when worker mode is required.

`ExcelWorkbookLoadOptions` extends the core load options:

| Field | Type | Required | Description |
| --- | --- | --- | --- |
| `workbookId` | `string` | no | Stable workbook id. |
| `mode` | `'auto' \| 'worker' \| 'inline'` | no | Runtime mode. Defaults to global default, initially `auto`. |
| `onModeResolved` | `(decision, workbookId) => void` | no | Receives runtime mode decision. |

## Runtime Properties

| Property | Type | Description |
| --- | --- | --- |
| `workbook.workbookId` | `string` | Loaded workbook id. |
| `workbook.mode` | `'worker' \| 'inline'` | Actual runtime mode. |
| `workbook.modeSource` | `ExcelModeSource` | Why the mode was chosen. |
| `workbook.worksheets` | `ExcelWorksheetCollection` | Worksheet live-object domain. |
| `workbook.names` | `ExcelNamedItemCollection` | Defined-name live-object domain. |

## Runtime Methods

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `workbook.executeCode(code, args?, timeoutMs?, signal?)` | Execute JS against a ready core `workbook` inside the selected runtime. | `string`, `Record<string, unknown>?`, `number?`, `AbortSignal?` | `Promise<unknown>` |
| `workbook.toModel()` | Export a normalized workbook model snapshot. | none | `Promise<ExcelWorkbookModel>` |
| `workbook.getMutations()` | Advanced host-runtime API: list SDK-recorded mutations in the selected runtime. | none | `Promise<ExcelWorkbookMutation[]>` |
| `workbook.clearMutations()` | Advanced host-runtime API: clear the mutation log in the selected runtime. | none | `Promise<void>` |
| `workbook.toZip()` | Export `.xlsx` bytes. | none | `Promise<Uint8Array>` |
| `workbook.dispose()` | Release runtime resources and terminate worker when present. | none | `Promise<void>` |

`executeCode` receives two variables:

| Variable | Type | Description |
| --- | --- | --- |
| `workbook` | `ExcelWorkbookCore` | The synchronous core workbook object living inside the runtime. |
| `args` | `Record<string, unknown> \| undefined` | Structured-cloneable caller arguments. |

Inside `executeCode`, use synchronous core calls and return plain structured-cloneable data:

```ts
const result = await workbook.executeCode(
  `
  const sheet = workbook.worksheets.getItem(args.sheetName);
  sheet.getRange('A1:B2').setValues(args.values);
  return sheet.getCell('A1').value();
  `,
  { sheetName: 'Sheet1', values: [[1, 2], [3, 4]] },
);
```

For normal SDK usage, prefer live-object methods instead of `executeCode`:

```ts
const workbook = await ExcelWorkbook.fromZip(bytes, { workbookId: 'book.xlsx', mode: 'auto' });
const sheet = workbook.worksheets.getItem('Sheet1');
await sheet.getRange('A1:B2').setValues([[1, 2], [3, 4]]);
const value = await sheet.getCell('A1').value();
const output = await workbook.toZip();
```

`timeoutMs` defaults to `30000`. In worker mode, a script timeout or abort terminates the worker runtime. The workbook becomes disposed after terminal script timeout/abort.

`getMutations()` and `clearMutations()` cross the same inline/worker boundary as `toModel()`. They are intended for host runtimes that need to verify and replay safe SDK mutations during writeback. Generated workbook-operation code should not call them.

## When To Use `executeCode`

`executeCode` is the escape hatch for advanced adapters and generated bulk plans. It runs JavaScript inside the selected runtime with access to a ready synchronous `ExcelWorkbookCore`.

Use it when:

| Situation | Why |
| --- | --- |
| A bulk edit would otherwise require many RPC round trips | One script can run close to the core kernel. |
| A host adapter wants to apply a precompiled plan | The plan can execute where the workbook state lives. |
| A low-level test needs to assert kernel behavior through the public runtime | It exercises the same inline/worker dispatch path. |

Prefer live-object APIs when:

| Situation | Preferred API |
| --- | --- |
| Read or write rectangular cells | `sheet.getRange(...).values()`, `setValues(...)`, `formulas()`, `setFormulas(...)` |
| Edit one cell | `sheet.getCell(...).value()`, `setValue(...)`, `formula()`, `setFormula(...)` |
| Edit object metadata | `sheet.tables`, `sheet.comments`, `sheet.hyperlinks`, `sheet.dataValidations`, `sheet.conditionalFormats` |
| Serialize output | `workbook.toModel()`, `workbook.toZip()` |

`executeCode` must return plain data, not live handles or core instances:

```ts
const snapshot = await workbook.executeCode(
  `
  const sheet = workbook.worksheets.getItem(args.sheetName);
  const range = sheet.getRange(args.address);
  return { address: range.address(), values: range.values() };
  `,
  { sheetName: 'Sheet1', address: 'A1:B2' },
);
```

## Static Runtime Configuration

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `ExcelWorkbook.setDefaultMode(mode)` | Set global default runtime mode. | `'auto' \| 'worker' \| 'inline'` | `void` |
| `ExcelWorkbook.setWorkerFactory(factory)` | Configure custom worker creation. | `ExcelWorkerFactory \| undefined` | `void` |
| `ExcelWorkbook.setWorkerBundleUrl(url)` | Configure Node worker bundle URL. | `string \| URL \| undefined` | `void` |
| `ExcelWorkbook.detectMode()` | Detect current auto mode. | none | `'worker' \| 'inline'` |
| `ExcelWorkbook.detectModeDecision(mode?)` | Detect mode and decision source. | `ExcelRuntimeMode?` | `ExcelModeDecision` |

Worker bundles are exported as:

| Entry | Use |
| --- | --- |
| `@coco-vexel/coco-office/excel-worker` | Node `worker_threads` entry. |
| `@coco-vexel/coco-office/excel-worker-browser` | Browser dedicated-worker entry. |

## Runtime Mode Rules

| Requested Mode | Behavior |
| --- | --- |
| `inline` | Always runs in current thread. |
| `worker` | Requires `workerFactory` or `workerBundleUrl`; throws if unavailable. |
| `auto` | Uses worker when a factory or bundle URL is configured; otherwise falls back to inline. |

## Non-Breaking Principles

These principles explain what runtime mode is allowed to change:

| Area | Principle |
| --- | --- |
| Public API shape | Runtime mode must not change the public method names, return types, or live-object hierarchy. |
| Workbook semantics | Inline and worker mode should produce the same workbook model and package output for the same operations. |
| Error surface | Worker transport errors may differ, but SDK-level validation and workbook operation errors should stay semantically equivalent. |
| Host integration | Runtime mode does not read from or write to a host application by itself. Host bridges remain responsible for source acquisition and writeback. |
| Performance | Worker mode may reduce UI blocking, but it does not change workbook semantics. |

If an operation cannot be made equivalent in worker mode, do not expose it only on worker objects. Add it to the shared public facade after core support and tests exist.

## CodeAgent Rules

1. Use `ExcelWorkbook` as the public root object for both inline and worker mode.
2. Prefer async live-object APIs for generated workbook-operation code.
3. Use `executeCode` only for advanced bulk operations that are easier to express as one runtime-side script.
4. Keep `args` and return values structured-cloneable in worker mode.
5. Do not return live object references from `executeCode`; return plain data.
6. Call `dispose()` when the caller owns the workbook lifecycle.
7. Do not use `ExcelWorkbookCore` in ordinary generated SDK code.
8. Do not branch generated operation logic on `workbook.mode`; configure runtime before loading the workbook.
9. Treat `getMutations()` and `clearMutations()` as host-runtime writeback APIs, not as task-level workbook operations.
