# PPT Host Integration

Use this document when a task mentions Microsoft PowerPoint, WPS Presentation,
ONLYOFFICE, desktop bridges, package bytes, writeback, or PPTAgent runtime.

## Core Boundary

Core accepts `.pptx` bytes or `PptPresentationModel`, exposes live-object
operations, and exports `.pptx` bytes. Host bridges are responsible for
obtaining source bytes and writing results back to the host.

## Skill Summary

Use this document only for host adapter or runtime orchestration work. Ordinary
generated PPT editing code should use the domain docs and a ready
`presentation: Presentation`.

```ts
// Host/runtime boundary, not user-operation CodePlan:
const presentation = await Presentation.fromZip(bytes);
await presentation.clearMutations();
await runGeneratedCode(presentation, args);
const output = await presentation.toZip();
```

Core SDK docs intentionally do not expose Office.js, WPSJS, ONLYOFFICE, COM, or
selection APIs as public CodePlan dependencies.

## Host Responsibilities

Hosts provide:

- document/package bytes or a model bridge;
- save/writeback policy;
- user confirmation UI;
- file path and permissions policy;
- selection/task pane/ribbon lifecycle;
- PPTAgent CodePlan runtime document provider;
- fallback behavior when live-host replay is unsafe.

Core provides:

- OOXML package parsing and writing;
- normalized snapshots;
- live-object editing APIs;
- mutation logs;
- model diff and writeback plans;
- host capability profiles and plan validation;
- package bytes via `toZip()`;
- image bytes via `exportImages()` / `shape.image.bytes()`.

## Capability Profiles

Core exposes host-neutral capability helpers:

```ts
const profile = createPptHostWritebackProfile('microsoft-powerpoint');
const validation = validatePptWritebackPlan(plan, profile);
```

Default live-host profiles are intentionally conservative:

| Host                    | Default Status | Default Capability Claim |
| ----------------------- | -------------- | ------------------------ |
| Microsoft PowerPoint    | `unproven`     | none                     |
| WPS Presentation        | `unproven`     | none                     |
| ONLYOFFICE Presentation | `unproven`     | none                     |
| Desktop package bridge  | `supported`    | full package replace     |
| Custom                  | `unproven`     | none                     |

Adapters must opt in to capabilities after runnable probes. A profile is a
contract declaration for `planPptWriteback(...)` and
`validatePptWritebackPlan(...)`; it is not a host implementation.

Profile states:

| Status      | Meaning                                                         |
| ----------- | --------------------------------------------------------------- |
| `unproven`  | Capability exists only as a possibility; do not execute it yet. |
| `probed`    | A runnable probe produced evidence for declared capabilities.   |
| `supported` | Built-in package-only path or adapter-provided support.         |

Default Microsoft PowerPoint, WPS, and ONLYOFFICE profiles are deliberately
`unproven`; they are not evidence that mutation replay or slide replay works.

## Probe Report Contract

Host probes should return plain JSON:

```ts
interface PptHostProbeReport {
  host:
    | 'microsoft-powerpoint'
    | 'wps-presentation'
    | 'onlyoffice-presentation'
    | 'desktop-package'
    | 'custom';
  generatedAt?: string;
  adapterVersion?: string;
  checks: PptHostProbeCheck[];
  maxMutationCount?: number;
  notes?: string[];
}

interface PptHostProbeCheck {
  name: PptHostProbeCheckName;
  passed: boolean;
  message?: string;
  details?: Record<string, unknown>;
}
```

Core derives profiles through:

```ts
const report = await runPptHostProbe(adapter);
const summary = summarizePptHostProbeReport(report);
const profile = createPptHostWritebackProfileFromProbeReport(report);
```

Supported probe check names:

- `package.replace`
- `slide.replace`
- `slide.insert`
- `slide.delete`
- `slide.reorder`
- `shape.add`
- `shape.delete`
- `shape.name.write`
- `shape.zOrder.write`
- `shape.geometry.write`
- `shape.style.write`
- `text.write`
- `text.style.write`
- `image.write`
- `table.cell.write`
- `table.cell.style.write`
- `table.values.write`
- `table.structure.write`
- `save.freshness`

Every writeback capability is gated by `save.freshness`. A host may prove that
it can edit text or slides, but core will not derive replay or package
capabilities unless the probe also proves that saved/exported output is fresh.

Probe-derived mutation support is intentionally conservative. It covers only the
mutation types mapped from passed probe checks. Host adapters may declare
additional `supportedMutationTypes`, but those declarations should come from
host-specific replay tests, not from the generic probe checklist alone.

Probe adapters implement a small runner interface:

```ts
interface PptHostProbeAdapter {
  host: PptHostKind;
  adapterVersion?: string;
  runCheck(check: PptHostProbeCheckName): Promise<Omit<PptHostProbeCheck, 'name'> | boolean>;
}
```

`runPptHostProbe(adapter, options?)` executes the requested checks and catches
thrown errors as failed check messages. `getDefaultPptHostProbeChecks(host)`
returns the default checklist. Live hosts use the broad slide/shape/text/image
and table checklist. `desktop-package` only checks `package.replace` and
`save.freshness`.

## Microsoft PowerPoint / Office.js

Office.js may use `PowerPoint.run` internally in a host adapter, but generated
core SDK code should not use Office.js. The adapter should map host operations
to core package/model operations or use mutation replay only when verified safe.

Default capability profile: `unproven`, no writeback capabilities. Enable
`canReplayMutations`, slide operations, or package replacement only after
Office.js probes prove the target runtime supports the exact operation.

Minimum probe report focus:

- slide enumeration and zero/one-based index translation;
- shape id/name enumeration;
- text write and style write;
- image replacement or delete/reinsert fallback;
- table cell and structure writes;
- slide insert/delete/reorder if enabled;
- package or host save freshness.

## WPS Presentation / WPSJS

WPSJS is adapter-level. Its COM/VBA-like object model should not leak into
`@coco-vexel/coco-office/ppt` generated code. Numeric indexes in public core APIs are
0-based; WPS adapters translate to host-specific indexing.

Default capability profile: `unproven`, no writeback capabilities. WPS adapters
should probe slide count/indexing, shape IDs/names, text write, image write,
table write, slide lifecycle, and save freshness before enabling replay.

Minimum probe report focus:

- `ActivePresentation` access;
- `Slides.Item(1)` to core index `0` mapping;
- `Shape.Id`, `Shape.Name`, and z-order behavior;
- `TextFrame.TextRange.Text` write and readback;
- picture replacement or delete/reinsert fallback;
- table cell writes;
- `SaveCopyAs` or equivalent freshness.

## ONLYOFFICE

ONLYOFFICE plugin APIs belong in the host bridge. If a JSON bridge is used, it
must preserve data needed by core or fall back to package replacement. Do not
return native host objects from generated CodePlan code.

Default capability profile: `unproven`, no writeback capabilities. ONLYOFFICE
adapters should prove JSON bridge fidelity and plugin write APIs before enabling
mutation or slide-level replay.

Minimum probe report focus:

- plugin command JSON boundary;
- slide and shape enumeration;
- text and table write/readback;
- image insert or replacement;
- presentation save/export freshness;
- unsupported native objects serialized as plain diagnostics, not host objects.

## Desktop

Desktop integrations should treat `toZip()` as the safe full document snapshot
boundary. Mutation replay can be added per host capability, but failures should
fall back to package writeback.

Default capability profile: `supported` for `canReplacePresentationPackage`.
This profile covers package replacement only; it does not imply live selection,
slide rendering, or host-native mutation replay.

## PPTAgent CodePlan

Recommended future boundary:

1. host initializes a `Presentation`;
2. host passes `presentation` and `args` to the PPT CodePlan runtime;
3. generated code uses public live-object APIs only;
4. runtime commits `toZip()` or validated mutation replay after success;
5. failed runs do not commit writeback.

Until that runtime exists, PPT Agent work should consume SDK docs only as a
planning reference. It should not assume live host selection, worker execution,
or `executeCode(...)`.

## Current Limitations

- Worker mode is planned but not implemented.
- `executeCode` currently throws.
- Slide thumbnails/rendering are not in core.
- Transition and animation editing, layout/master editing, linked workbook chart
  editing, SmartArt editing, OLE editing/rendering, merged-cell structural
  editing, and rich media playback are P2/P3 domains.
- Speaker notes and comments have basic read/write APIs, but richer formatting
  and threading are not implemented.
- Chart cache data replacement exists for conservative existing-series updates;
  linked workbook edits and series add/remove are not implemented.

## CodePlan Rules

1. Do not call host APIs from generated core operation code.
2. Do not initialize documents inside generated code.
3. Prefer full package writeback unless adapter capability proves safe mutation replay.
4. Keep host-specific indexing, relationship, and selection rules outside core code.
5. Treat default live-host profiles as unproven until probe reports enable capabilities.
6. Do not return native host objects from probe reports or generated CodePlan.
7. Keep `PowerPoint.run`, WPSJS, ONLYOFFICE commands, COM, and filesystem writeback inside adapters.
