import { describe, expect, test } from 'vitest';
import type { NativeWordChange } from '../native-commands.js';
import { createWordEditorChangeEvent } from '../runtime-events.js';

const insert: NativeWordChange = { kind: 'text', op: 'insert', paraId: 'P1', offset: 0, text: 'A' };
const format: NativeWordChange = {
  kind: 'format',
  op: 'run',
  paraId: 'P1',
  from: 0,
  to: 1,
  format: { bold: true },
};

describe('native Word runtime change events', () => {
  test('reports only changes added since the previous notification', () => {
    expect(createWordEditorChangeEvent([insert], [insert, format], 2)).toEqual({
      revision: 2,
      changes: [format],
      revertedChanges: [],
      totalChangeCount: 2,
    });
  });

  test('reports changes removed by undo', () => {
    expect(createWordEditorChangeEvent([insert, format], [insert], 3)).toEqual({
      revision: 3,
      changes: [],
      revertedChanges: [format],
      totalChangeCount: 1,
    });
  });

  test('reports a replaced suffix as reverted and added changes', () => {
    const replacement: NativeWordChange = {
      kind: 'text',
      op: 'replace',
      paraId: 'P1',
      offset: 0,
      text: 'B',
      deletedText: 'A',
    };
    expect(createWordEditorChangeEvent([insert, format], [insert, replacement], 4)).toEqual({
      revision: 4,
      changes: [replacement],
      revertedChanges: [format],
      totalChangeCount: 2,
    });
  });
});
