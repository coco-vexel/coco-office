/* eslint-disable no-console -- CLI verification output. */
import { readdir, readFile } from 'node:fs/promises';
import { extname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const packageRoots = ['core', 'editor', 'oojson'].map((name) =>
  join(root, 'packages', name, 'dist'),
);
const textExtensions = new Set(['.cjs', '.cts', '.js', '.json', '.map', '.mjs', '.ts']);
const forbidden = [
  { label: 'private repository name', pattern: /coco-(?:engine|suite|desktop)/i },
  { label: 'internal source marker', pattern: /internal:coco/i },
  { label: 'macOS user path', pattern: /\/Users\// },
  { label: 'Windows user path', pattern: /[A-Z]:\\Users\\/i },
];

async function filesUnder(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...(await filesUnder(path)));
    else if (entry.isFile() && textExtensions.has(extname(entry.name))) files.push(path);
  }
  return files;
}

const findings = [];
for (const directory of packageRoots) {
  for (const path of await filesUnder(directory)) {
    const contents = await readFile(path, 'utf8');
    for (const rule of forbidden) {
      if (rule.pattern.test(contents))
        findings.push(`${relative(root, path)}: ${rule.label} (${rule.pattern})`);
    }
  }
}

if (findings.length > 0) {
  console.error('Package output contains private build provenance:');
  for (const finding of findings) console.error(`- ${finding}`);
  process.exitCode = 1;
} else {
  console.log('Package output contains no private repository names or local user paths.');
}
