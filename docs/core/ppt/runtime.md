# PPT Runtime

Use this document when a task mentions async semantics, inline/worker execution,
`executeCode`, or runtime boundaries.

## Public Runtime Shape

`Presentation` is async by design:

```ts
const presentation = await Presentation.fromZip(bytes, { mode: 'auto' });
const slide = presentation.slides.getItem(0); // sync handle
const count = await slide.shapes.getCount(); // async read
```

Handles are lightweight and synchronous. Reads, writes, serialization, search,
existence checks, and properties snapshots return `Promise`.

## Skill Summary

Use this runtime document to decide what SDK operation code may assume about the
execution environment:

```ts
export async function run(presentation, args) {
  const slide = presentation.slides.getItem(args.slideIndex ?? 0);
  await slide.shapes.addTextBox(args.text, { name: 'GeneratedText' });
  return await presentation.properties();
}
```

SDK operation code should receive a ready `presentation` from the
runtime. It should not create host sessions, call `Presentation.fromZip(...)`,
choose worker mode, or call `executeCode(...)`.

## Runtime Modes

```ts
type PresentationRuntimeMode = 'auto' | 'inline' | 'worker';
```

Current implementation:

- `auto` resolves to inline.
- `inline` is supported.
- `worker` is planned and currently rejected by runtime mode resolution.

Mode decision fields:

| Field                      | Meaning                                    |
| -------------------------- | ------------------------------------------ |
| `presentation.mode`        | Resolved runtime mode: currently `inline`. |
| `presentation.modeSource`  | Why the mode was selected.                 |
| `opts.onModeResolved(...)` | Optional callback for runtime diagnostics. |

`PresentationModeDecision['source']` is one of:

- `explicit-inline`
- `explicit-worker`
- `auto/inline`

The public API should remain worker-compatible: SDK operation code must not depend
on object identity, host APIs, DOM state, or non-structured-cloneable return
values.

## executeCode Status

`presentation.executeCode(...)` exists as a planned script boundary but currently
throws. SDK callers should use direct live-object APIs until a runtime document
provider exists.

Do not use `executeCode(...)` as a shortcut for PPT operations today. Call
live-object APIs directly from SDK operation code.

For future runtime execution, mirror the Word/Excel pattern:

- runtime initializes the `Presentation`;
- operation code receives a ready `presentation`;
- runtime exports/runs `async function run(presentation, args)`;
- code uses public live-object APIs;
- runtime commits `toZip()` or mutation writeback after successful execution.

## Runtime Boundary

Runtime-executed SDK operation code should look like:

```ts
export async function run(presentation, args) {
  const slide = presentation.slides.getItem(args.slideIndex ?? 0);
  await slide.shapes.addTextBox(args.text, { name: 'GeneratedText' });
  return await presentation.properties();
}
```

It should not call:

- `Presentation.fromZip(...)`
- `coco.init(...)`
- Office.js/WPSJS/ONLYOFFICE APIs
- `PresentationCore` APIs
- raw filesystem APIs
- `presentation.executeCode(...)`
- `presentation.getMutations()` / `presentation.clearMutations()`

unless the host/runtime explicitly asks for an adapter or test.

## Return Values

SDK operation code should return structured-cloneable values:

| Allowed Return                                | Notes                               |
| --------------------------------------------- | ----------------------------------- |
| Plain JSON objects, arrays, strings, numbers  | Preferred for structured results.   |
| `Uint8Array` package bytes                    | Only when the task asks for output. |
| `PptPresentationModel` / properties snapshots | OK for inspection/reporting.        |

Do not return live handles such as `Slide`, `Shape`, `TableCell`, or host-native
objects.

## SDK Operation Rules

1. Assume a ready `presentation` object.
2. Always `await` reads and writes.
3. Do not use worker mode until PPT worker runtime is implemented.
4. Do not return live handles from SDK operation code.
5. Return plain JSON, strings, numbers, arrays, or package bytes when requested.
6. Do not call `Presentation.fromZip(...)` in ordinary SDK operation code.
7. Do not call `executeCode(...)` until the PPT runtime provider is implemented.
8. Do not use mutation APIs in ordinary SDK operation code.
