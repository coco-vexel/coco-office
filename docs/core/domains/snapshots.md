# Snapshot APIs

This domain covers document snapshots, rollback, and per-document change logs.

Use this document when the task says: snapshot, rollback, undo checkpoint, restore previous state, save document state, or change log.

## Skill Summary

Snapshots are document-level safety checkpoints. They are not document editing APIs. A snapshot stores the current OOXML package state through the configured snapshot storage adapter.

## Entry Point

Snapshots live on `doc.snapshots`.

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.snapshots.take(reason?, details?, options?)` | Save current document state. | `reason?: string`, `details?: unknown`, `TakeOptions?` | `Promise<string>` snapshot id |
| `doc.snapshots.restore(id, opts?)` | Restore a snapshot without deleting it. Use before an external host commit. | `id: string`, `RollbackOptions?` | `Promise<void>` |
| `doc.snapshots.rollback(id, opts?)` | Restore a snapshot and delete it after success. | `id: string`, `RollbackOptions?` | `Promise<void>` |
| `doc.snapshots.list()` | List snapshot metadata. | none | `Promise<SnapshotInfo[]>` |
| `doc.snapshots.delete(id)` | Delete one snapshot. | `id: string` | `Promise<void>` |
| `doc.snapshots.clear()` | Delete all snapshots for this document. | none | `Promise<void>` |
| `doc.snapshots.recordChange(change)` | Record a pending change log item. | `{toolName:string; summary?:unknown}` | `void` |
| `doc.snapshots.changes()` | Read pending change log items. | none | `ChangeLog[]` |

`TakeOptions`:

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `pin` | `boolean` | `false` | Pinned snapshots are not evicted by FIFO cleanup. |
| `dedupeKey` | `string` | none | Reuse the latest snapshot id in this logical scope when the whole-package SHA-256 is unchanged. |

`RollbackOptions`:

| Field | Type | Description |
| --- | --- | --- |
| `onProgress` | `(phase:string, pct?:number) => void` | Progress callback. Phases include `parse`, `build-indices`, and `done`. |

## Defaults

| Constant | Value | Description |
| --- | --- | --- |
| `DEFAULT_MAX_SNAPSHOTS` | `10` | Default FIFO limit for non-pinned snapshots. |

## Examples

Create a checkpoint before edits:

```ts
const snapshotId = await doc.snapshots.take('before-formatting');
```

Avoid duplicate user checkpoints while keeping other owners isolated:

```ts
const snapshotId = await doc.snapshots.take('turn-start', details, {
  dedupeKey: `user:${sessionId}`,
});
```

Rollback if a validation fails:

```ts
const snapshotId = await doc.snapshots.take('before-table-edit');
try {
  const table = await doc.tables.get('tbl-0001');
  await table?.cell(0, 0).replaceWith('新值');
} catch (err) {
  await doc.snapshots.rollback(snapshotId);
  throw err;
}
```

Commit a restore to an external host atomically:

```ts
await doc.snapshots.restore(snapshotId);
const xml = await doc.toFlatOpc();
await host.write(xml);
await doc.snapshots.delete(snapshotId);
```

## CodeAgent Rules

1. Use snapshots as safety checkpoints around risky multi-step edits.
2. Do not use snapshots to read document content.
3. Rollback consumes the target snapshot after success.
4. When an external host write is part of the operation, use `restore` and delete only after the host write succeeds.
5. Do not pin ordinary plugin/user snapshots unless the host explicitly owns snapshot retention.
