import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const here = dirname(fileURLToPath(import.meta.url));
const pptRoot = resolve(here, '..');

const indexSource = readFileSync(resolve(pptRoot, 'index.ts'), 'utf8');

describe('PPT SDK public exports', () => {
  it('re-exports the SDK-facing public surface from the package entrypoint', () => {
    const actual = extractIndexExports(indexSource);
    const expected = [
      ...extractModuleExports('runtime.ts'),
      ...extractModuleExports('model.ts'),
      ...extractModuleExports('mutations.ts'),
      ...extractModuleExports('writeback-planner.ts'),
      ...extractModuleExports('host-capabilities.ts'),
      ...extractModuleExports('pptx-writer.ts'),
      ...extractModuleExports('presentation-core.ts').filter(
        (name) => name.endsWith('Core') || /^Ppt.*Options$/.test(name) || name === 'PptPresetShape',
      ),
    ].filter((name) => !name.startsWith('PptRead') && !name.startsWith('PptShapeInternal'));

    expect([...new Set(expected)].filter((name) => !actual.has(name)).sort()).toEqual([]);
  });
});

function extractModuleExports(filename: string): string[] {
  const source = readFileSync(resolve(pptRoot, filename), 'utf8');
  return [...source.matchAll(/^export (?:class|interface|type|function) (\w+)/gm)].map(
    (match) => match[1]!,
  );
}

function extractIndexExports(source: string): Set<string> {
  const names = new Set<string>();
  for (const block of source.matchAll(/export(?: type)? \{([\s\S]*?)\} from/g)) {
    for (const raw of block[1]!.split(',')) {
      const name = raw.trim().replace(/^type\s+/, '');
      if (name) names.add(name);
    }
  }
  return names;
}
