/* eslint-disable no-console -- CLI release artifact output. */
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { access, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { basename, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = fileURLToPath(new URL('../', import.meta.url));
const coreRoot = join(root, 'packages', 'core');
const output = join(root, '.release-check', 'sdk');
const packageOutput = join(output, 'packages');
const packageMetadata = JSON.parse(await readFile(join(coreRoot, 'package.json'), 'utf8'));
const tarballName = `${packageMetadata.name.slice(1).replace('/', '-')}-${packageMetadata.version}.tgz`;
const zipName = `coco-office-sdk-${packageMetadata.version}.zip`;
const fixedDate = new Date('1980-01-01T00:00:00.000Z');

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, stdio: 'inherit' });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${command} failed (${result.status}) in ${cwd}`);
}

async function addFile(zip, source, target) {
  zip.file(target.replaceAll('\\', '/'), await readFile(source), {
    createFolders: false,
    date: fixedDate,
  });
}

async function addTree(zip, sourceRoot, targetRoot, include = () => true, currentRelative = '') {
  const entries = await readdir(sourceRoot, { withFileTypes: true });
  for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
    const source = join(sourceRoot, entry.name);
    const relativePath = join(currentRelative, relative(sourceRoot, source));
    if (!include(relativePath, entry)) continue;
    const target = `${targetRoot}/${entry.name}`;
    if (entry.isDirectory()) await addTree(zip, source, target, include, relativePath);
    else if (entry.isFile()) await addFile(zip, source, target);
  }
}

await access(join(coreRoot, 'dist', 'index.js')).catch(() => {
  throw new Error('Core dist is missing. Run this command through "corepack pnpm bundle:sdk".');
});
await rm(output, { recursive: true, force: true });
await mkdir(packageOutput, { recursive: true });

const pnpmCli = process.env.npm_execpath;
if (!pnpmCli) throw new Error('Run this command through pnpm.');
run(process.execPath, [pnpmCli, 'pack', '--pack-destination', packageOutput], coreRoot);

const tarballPath = join(packageOutput, tarballName);
await access(tarballPath);
const zip = new JSZip();
const readme = `# Coco Office SDK ${packageMetadata.version}

This bundle contains the installable \`${packageMetadata.name}\` package, offline
documentation and runnable synthetic Word, Excel and PowerPoint examples.

## Requirements

- Node.js 22 or 24
- npm with access to the public npm registry for third-party dependencies

## Run the examples

\`\`\`bash
cd examples
npm install
npm run example
\`\`\`

The example creates, edits, exports and reloads synthetic DOCX, XLSX and PPTX
files under \`examples/output/\`. It does not require Microsoft Office, WPS,
ONLYOFFICE, a backend or private Coco repositories.

## Documentation

Start with \`docs/core/getting-started.md\`. The Word documentation index is
\`docs/core/README.md\`; Excel and PowerPoint start at
\`docs/core/excel/README.md\` and \`docs/core/ppt/README.md\`.

## License

The Coco Office package is Apache-2.0. See \`LICENSE\` and
\`THIRD_PARTY_NOTICES.md\`.
`;
zip.file('README.md', readme, { createFolders: false, date: fixedDate });
await addFile(zip, join(coreRoot, 'LICENSE'), 'LICENSE');
await addFile(zip, join(coreRoot, 'THIRD_PARTY_NOTICES.md'), 'THIRD_PARTY_NOTICES.md');
await addFile(zip, join(coreRoot, 'README.md'), 'docs/package-readme.md');
await addFile(zip, tarballPath, `packages/${tarballName}`);

const excludedDocs = new Set([
  'advanced/agent-notes.md',
  'advanced/codeplan.md',
  'ppt/sdk-plan.md',
  'ppt/sdk-review.md',
]);
await addTree(zip, join(root, 'docs', 'core'), 'docs/core', (path) => {
  const normalized = path.replaceAll('\\', '/');
  return !excludedDocs.has(normalized);
});
await addFile(zip, join(root, 'examples', 'samples.mjs'), 'examples/samples.mjs');
await addFile(zip, join(root, 'examples', 'roundtrip.mjs'), 'examples/roundtrip.mjs');
zip.file(
  'examples/package.json',
  `${JSON.stringify(
    {
      name: 'coco-office-sdk-examples',
      private: true,
      type: 'module',
      scripts: { example: 'node roundtrip.mjs' },
      dependencies: {
        [packageMetadata.name]: `file:../packages/${tarballName}`,
      },
    },
    null,
    2,
  )}\n`,
  { createFolders: false, date: fixedDate },
);

const bytes = await zip.generateAsync({
  type: 'nodebuffer',
  platform: 'UNIX',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});
const zipPath = join(output, zipName);
await writeFile(zipPath, bytes);
const sha256 = createHash('sha256').update(bytes).digest('hex');
await writeFile(join(output, `${zipName}.sha256`), `${sha256}  ${zipName}\n`);
await writeFile(
  join(output, 'bundle.json'),
  `${JSON.stringify(
    {
      package: packageMetadata.name,
      version: packageMetadata.version,
      archive: basename(zipPath),
      sha256,
    },
    null,
    2,
  )}\n`,
);
console.log(`SDK bundle: ${zipPath}`);
console.log(`SHA-256: ${sha256}`);
