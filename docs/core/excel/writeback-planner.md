# Excel Writeback Planner

`writeback-planner` is the host-neutral planning layer for applying a changed workbook back to a live host.

It does not call Office.js, WPS JS, ONLYOFFICE APIs, COM, file-system replacement, or any UI/download flow. It compares workbook models, inspects declared host capabilities, and returns an execution plan made of portable writeback intents. A host adapter is responsible for executing those intents.

## Boundary

```text
before ExcelWorkbookModel
after ExcelWorkbookModel
host capabilities
        |
        v
planExcelWriteback(...)
        |
        v
ExcelWritebackPlan
        |
        v
Office / WPS / ONLYOFFICE / Desktop adapter executes operations
```

The planner belongs to the Excel SDK core because it encodes portable workbook writeback decisions. Host adapters remain outside core.

## APIs

```ts
function diffExcelWorkbookModels(
  before: ExcelWorkbookModel,
  after: ExcelWorkbookModel,
): ExcelWorkbookDiff;

function planExcelWriteback(
  before: ExcelWorkbookModel,
  after: ExcelWorkbookModel,
  capabilities: ExcelWritebackCapabilities,
  options?: ExcelWritebackPlanOptions,
): ExcelWritebackPlan;
```

## Capabilities

```ts
interface ExcelWritebackCapabilities {
  canReplaceWorkbookPackage?: boolean;
  canInsertShadowSheets?: boolean;
  canCopySheetContent?: boolean;
  canDeleteSheet?: boolean;
  canRenameSheet?: boolean;
  canPatchCells?: boolean;
  canPatchSheetObjects?: boolean;
  maxCellPatchCount?: number;
}
```

Capabilities describe what the host adapter has proven it can do. They must not be guessed.

## Strategies

| Strategy | Meaning | Typical Host Fit |
| --- | --- | --- |
| `no-op` | Models are equivalent for writeback planning. | All hosts. |
| `workbook-package-replace` | Replace active workbook with the full target `.xlsx` package. | Desktop/file contexts where in-place replacement is proven. |
| `sheet-shadow-replace` | Insert generated/shadow sheets, copy content into target sheets, repair objects, delete shadows. | Office/WPS-style plugin bridges when package read exists but package writeback does not. |
| `diff-patch` | Patch cells and optionally sheet objects directly through live APIs. | Fallback or small-edit path when package/sheet-shadow writeback is unavailable or explicitly requested. |
| `unsupported` | Declared capabilities cannot express the diff safely. | Adapter should stop or ask for a different strategy. |

## Operations

Planner operations are intents, not host API calls:

- `replaceWorkbookPackage`
- `insertShadowSheet`
- `copySheetContent`
- `repairSheetObjects`
- `deleteSheet`
- `renameSheet`
- `patchCells`
- `patchSheetObjects`

Adapters map these operations to concrete host calls. For example, Office may execute `insertShadowSheet` with `insertWorksheetsFromBase64`, while WPS may open a generated workbook and copy ranges across workbooks.

`insertShadowSheet` contains a defensive `targetSheet` snapshot:

```ts
{
  type: 'insertShadowSheet',
  sheetId: '1',
  sheetName: '__coco_shadow_1',
  targetSheet: targetSheetModel
}
```

`sheetName` is the temporary shadow sheet name requested by the planner. `targetSheet` is the desired final sheet model. The adapter may generate a one-sheet workbook/package from this model, insert it as the shadow sheet, then execute `copySheetContent`.

Use the adapter-facing SDK helper to generate the workbook package:

```ts
const payload = await createShadowWorkbookPayload(operation.targetSheet, {
  sheetName: operation.sheetName,
  styles: operation.targetStyles
});

// payload.bytes: xlsx bytes
// payload.base64: base64 xlsx, suitable for hosts such as Office insert APIs
```

`patchCells` contains executable cell payloads:

```ts
{
  type: 'patchCells',
  sheetName: 'Sheet1',
  cellCount: 2,
  changes: [
    { kind: 'modified', address: 'A1', before: oldCell, after: newCell },
    { kind: 'added', address: 'B2', after: newCell }
  ]
}
```

Deleted cells are emitted as `{ kind: 'deleted', address, before }`; the adapter should clear that cell.

## Notes

- `diff-patch` is intentionally conservative. In `auto` mode, package replacement and sheet-shadow replacement are preferred when available. Direct patching is selected when stronger writeback paths are unavailable, or when `preference: 'diff'` is explicitly requested.
- `diff-patch` only plans direct patches when the declared cell-change count is within `maxCellPatchCount` and structural/object changes are supported.
- `sheet-shadow-replace` may include `repairSheetObjects` when object changes are detected. The adapter must decide how to repair charts, shapes, images, comments, validations, conditional formats, and host-specific object fidelity.
- `workbook-package-replace` is never selected unless `canReplaceWorkbookPackage` is true or explicitly preferred and supported.
- The planner does not execute filter/sort semantics. Table filter/sort changes are model/object changes and are handled by the chosen writeback strategy.
