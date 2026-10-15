import { createHash } from 'node:crypto';
import { describe, expect, test } from 'vitest';
import type { SnapshotPartEntry } from '../kernel.js';
import { fingerprintSnapshotParts } from '../snapshot-fingerprint.js';

describe('fingerprintSnapshotParts', () => {
  test('与 Node SHA-256 参考实现一致，且不受 parts 输入顺序影响', () => {
    const parts: SnapshotPartEntry[] = [
      {
        path: 'word/document.xml',
        contentType: 'application/xml',
        bytes: new Uint8Array([1, 2, 3]),
      },
      {
        path: '_rels/.rels',
        contentType: 'application/relationships+xml',
        bytes: new Uint8Array([9, 8]),
      },
    ];
    const expected = referenceFingerprint(parts);
    expect(fingerprintSnapshotParts(parts)).toBe(expected);
    expect(fingerprintSnapshotParts(parts.slice().reverse())).toBe(expected);
  });
});

function referenceFingerprint(parts: SnapshotPartEntry[]): string {
  const hash = createHash('sha256');
  hash.update('coco-snapshot-v1');
  const sorted = parts.slice().sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
  for (const part of sorted) {
    updateFramed(hash, new TextEncoder().encode(part.path));
    updateFramed(hash, new TextEncoder().encode(part.contentType));
    updateFramed(hash, part.bytes);
  }
  return hash.digest('hex');
}

function updateFramed(hash: ReturnType<typeof createHash>, bytes: Uint8Array): void {
  const length = new Uint8Array(4);
  new DataView(length.buffer).setUint32(0, bytes.length, false);
  hash.update(length);
  hash.update(bytes);
}
