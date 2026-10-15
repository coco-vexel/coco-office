import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const here = dirname(fileURLToPath(import.meta.url));
const modelPath = resolve(here, '../model.ts');
const docsRoot = resolve(here, '../../../../../docs/core/ppt');

const modelSource = readFileSync(modelPath, 'utf8');
const docsSource = readMarkdownTree(docsRoot);

describe('PPT SDK model references', () => {
  it('keeps every exported model type discoverable from PPT SDK docs', () => {
    const missing = extractExportedModelTypes(modelSource).filter((typeName) => {
      const pattern = new RegExp(`\\b${escapeRegExp(typeName)}\\b`);
      return !pattern.test(docsSource);
    });

    expect(missing).toEqual([]);
  });
});

function extractExportedModelTypes(source: string): string[] {
  return [...source.matchAll(/^export (?:interface|type) (\w+)/gm)].map((match) => match[1]!);
}

function readMarkdownTree(root: string): string {
  let combined = '';
  for (const entry of readdirSync(root)) {
    const path = join(root, entry);
    const stats = statSync(path);
    if (stats.isDirectory()) {
      combined += `\n${readMarkdownTree(path)}`;
    } else if (entry.endsWith('.md')) {
      combined += `\n${readFileSync(path, 'utf8')}`;
    }
  }
  return combined;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
