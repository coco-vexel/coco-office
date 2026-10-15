# PPT Writeback Planner

Use this document when a task mentions PPT model diff, writeback strategy,
package replacement, slide replacement, mutation replay, or host capability
planning.

The planner is host-neutral. It does not call Office.js, WPSJS, ONLYOFFICE, COM,
file-system replacement, or UI flows. It compares normalized presentation models,
inspects declared host capabilities, and returns portable writeback intents for
host adapters.

## Skill Summary

Use the planner only after SDK operations have already run:

```ts
const before = await presentation.toModel();
await presentation.clearMutations();

// Run user operation code with live-object APIs.

const after = await presentation.toModel();
const mutations = await presentation.getMutations();
const plan = planPptWriteback(before, after, capabilities, { mutations });
```

Ordinary SDK operation code should not call this planner. Runtime and adapter
code can call it to choose between package replacement, slide package
operations, or mutation replay.

## APIs

```ts
import {
  createPptHostWritebackProfile,
  diffPptPresentationModels,
  planPptWriteback,
  runPptHostProbe,
  validatePptWritebackPlan,
} from '@coco-vexel/coco-office/ppt';

const before = await presentation.toModel();
await presentation.clearMutations();

// Execute SDK operations.

const after = await presentation.toModel();
const mutations = await presentation.getMutations();

const plan = planPptWriteback(
  before,
  after,
  {
    canReplacePresentationPackage: true,
    canReplayMutations: true,
  },
  { mutations },
);

const profile = createPptHostWritebackProfile('desktop-package');
const validation = validatePptWritebackPlan(plan, profile);
```

API responsibilities:

| API                                                       | Responsibility                                     |
| --------------------------------------------------------- | -------------------------------------------------- |
| `diffPptPresentationModels(before, after)`                | Classify normalized model changes.                 |
| `planPptWriteback(before, after, capabilities, options?)` | Choose portable writeback strategy and operations. |
| `createPptHostWritebackProfile(host, options?)`           | Build conservative host capability profile.        |
| `createPptHostWritebackProfileFromProbeReport(report)`    | Convert a plain JSON probe report into a profile.  |
| `runPptHostProbe(adapter, options?)`                      | Execute adapter-provided probe checks.             |
| `summarizePptHostProbeReport(report)`                     | Summarize probe results and derived capabilities.  |
| `validatePptWritebackPlan(plan, profileOrCapabilities)`   | Validate a plan against declared capabilities.     |
| `getPptWritebackOperationRequirements(operation)`         | List capability flags required by one operation.   |

## Strategies

| Strategy                       | Meaning                                      | Typical Adapter Action                                    |
| ------------------------------ | -------------------------------------------- | --------------------------------------------------------- |
| `no-op`                        | No model diff.                               | Do nothing.                                               |
| `presentation-package-replace` | Replace the active deck with a full package. | Use the safest full `.pptx` writeback path.               |
| `slide-package-replace`        | Apply slide-level package operations.        | Insert, replace, delete, or reorder slides.               |
| `mutation-replay`              | Replay recorded SDK mutations.               | Map each mutation to host-native operations.              |
| `unsupported`                  | Capabilities cannot express the diff safely. | Stop, ask for another path, or fall back outside planner. |

Auto mode prefers the safest available strategy in this order:

1. `presentation-package-replace`
2. `slide-package-replace`
3. `mutation-replay`

Use `preference: 'mutation'` or `preference: 'slide-package'` when an adapter
wants to test a specific host path.

Strategy selection should remain conservative. If a host can safely replace the
whole package and the workflow allows it, package replacement is the preferred
baseline path. Mutation replay is for proven fast paths after baseline checks.
Mutation replay may return warnings when the model diff includes
presentation-level, slide-size, or media-list changes; adapters must validate
those warnings before executing host operations.

## Capabilities

```ts
interface PptWritebackCapabilities {
  canReplacePresentationPackage?: boolean;
  canReplaceSlides?: boolean;
  canInsertSlides?: boolean;
  canDeleteSlides?: boolean;
  canReorderSlides?: boolean;
  canReplayMutations?: boolean;
  supportedMutationTypes?: PptMutation['type'][];
  maxMutationCount?: number;
}
```

Capability flags are declarations from a host adapter. The planner trusts these
flags only to select a strategy; adapters must still validate live document state
before executing operations.

When `supportedMutationTypes` is omitted, mutation replay planning treats every
provided mutation type as allowed. Use this only for trusted adapters with their
own replay validation; otherwise provide an explicit allow-list.

## Host Profiles

```ts
const profile = createPptHostWritebackProfile('wps-presentation', {
  status: 'probed',
  capabilities: {
    canReplayMutations: true,
    supportedMutationTypes: ['shape.text.set'],
    maxMutationCount: 20,
  },
});
```

Built-in host kinds:

- `microsoft-powerpoint`
- `wps-presentation`
- `onlyoffice-presentation`
- `desktop-package`
- `custom`

Default profile rules:

- Microsoft PowerPoint, WPS, ONLYOFFICE, and custom hosts default to
  `status: 'unproven'` and no capabilities.
- Adapters may enable capabilities only after runnable probes prove the behavior.
- `desktop-package` defaults to `status: 'supported'` with
  `canReplacePresentationPackage: true`.

Use `validatePptWritebackPlan(plan, profile)` before executing a plan. It checks
operation-level capability requirements, mutation replay allow-lists, mutation
count limits, and unproven-profile warnings.

Probe-backed profiles:

```ts
const report = await runPptHostProbe(adapter, { maxMutationCount: 20 });
const profile = createPptHostWritebackProfileFromProbeReport(report);
```

`summarizePptHostProbeReport(report)` returns passed checks, failed checks,
derived capabilities, and notes. Failed check messages become profile notes.
Every derived writeback capability requires `save.freshness`.

## Diff Model

`diffPptPresentationModels(before, after)` returns:

- `changed`
- `presentationPropertiesChanged`
- `slideSizeChanged`
- `mediaChanged`
- `slides: PptSlideDiff[]`

Each slide diff includes:

- `kind: 'added' | 'deleted' | 'modified'`
- `slideId`
- `beforeIndex` / `afterIndex`
- `moved`
- `propertiesChanged`
- `shapeOrderChanged`
- `shapeChanges`
- defensive `before` / `after` slide snapshots when available

Each shape diff includes:

- `kind: 'added' | 'deleted' | 'modified'`
- `shapeId`
- `beforeName` / `afterName`
- `beforeKind` / `afterKind`
- defensive `before` / `after` shape snapshots when available

Numeric slide indexes remain zero-based.

## Operations

Planner operations are plain JSON intents:

- `replacePresentationPackage`
- `replaceSlidePackage`
- `insertSlidePackage`
- `deleteSlide`
- `moveSlide`
- `replayMutations`

`getPptWritebackOperationRequirements(operation)` maps each operation to the
required capability:

| Operation                    | Required Capability             |
| ---------------------------- | ------------------------------- |
| `replacePresentationPackage` | `canReplacePresentationPackage` |
| `replaceSlidePackage`        | `canReplaceSlides`              |
| `insertSlidePackage`         | `canInsertSlides`               |
| `deleteSlide`                | `canDeleteSlides`               |
| `moveSlide`                  | `canReorderSlides`              |
| `replayMutations`            | `canReplayMutations`            |

The planner does not include package bytes. Host runtimes should obtain bytes
from `presentation.toZip()` or construct slide payloads through adapter-specific
packaging code after accepting the plan.

## Adapter Rules

1. Treat planner output as an execution plan, not proof that the live host still
   matches the SDK baseline.
2. Validate the live host document against the baseline before replaying
   mutations.
3. Validate slide ids, indexes, relationship ids, and media paths before
   slide-level replacement.
4. Prefer full package replacement when available and the user workflow allows
   it.
5. Use mutation replay only when every mutation type is explicitly supported by
   the host adapter.
6. Keep unsupported plans explicit; do not silently downgrade to partial edits.

## SDK Operation Rules

Ordinary SDK operation code should not call writeback planning APIs. These APIs
belong in runtimes and adapter orchestration after SDK operations finish.

1. Do not call `planPptWriteback(...)` from ordinary slide-editing SDK code.
2. Do not use host capability profiles to decide what SDK operation to run; use
   the SDK docs for supported operations.
3. Runtime code may pass SDK mutations into the planner after clearing mutations
   at baseline.
4. Adapter code must validate live host state before executing a returned plan.
5. Unsupported plans should stop or choose an explicit fallback; do not silently
   perform partial edits.
