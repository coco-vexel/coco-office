import { describe, expect, it } from 'vitest';
import * as coreEntry from '../index.js';

describe('@coco-vexel/coco-office entry', () => {
  it('loads without side effects', () => {
    expect(coreEntry).toBeDefined();
    expect(typeof coreEntry).toBe('object');
  });
});
