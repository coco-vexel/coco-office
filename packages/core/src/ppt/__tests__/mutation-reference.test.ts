import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const here = dirname(fileURLToPath(import.meta.url));
const mutationsPath = resolve(here, '../mutations.ts');
const mutationDocsPath = resolve(here, '../../../../../docs/core/ppt/mutation-writeback.md');

const mutationsSource = readFileSync(mutationsPath, 'utf8');
const mutationDocs = readFileSync(mutationDocsPath, 'utf8');

describe('PPT SDK mutation reference', () => {
  it('documents every PptMutation type string', () => {
    const documentedTypes = new Set(extractMutationTypeStrings(mutationDocs));
    const missing = extractMutationTypeStrings(mutationsSource).filter(
      (mutationType) => !documentedTypes.has(mutationType),
    );

    expect(missing).toEqual([]);
  });
});

function extractMutationTypeStrings(source: string): string[] {
  return [...source.matchAll(/['`]([a-z]+(?:\.[a-zA-Z]+)+)['`]/g)]
    .map((match) => match[1]!)
    .filter((value) => /^(?:shape|slide|table)\./.test(value));
}
