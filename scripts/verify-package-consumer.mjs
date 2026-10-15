/* eslint-disable no-console -- CLI verification output. */
import { spawnSync } from 'node:child_process';
import { copyFile, cp, mkdir, mkdtemp, readFile, realpath, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const pnpmCli = process.env.npm_execpath;
if (!pnpmCli) throw new Error('Run this check with pnpm verify:packages.');
const output = join(root, '.release-check', 'packages');
await mkdir(output, { recursive: true });
function tarballName(pkg) {
  return `${pkg.name.slice(1).replace('/', '-')}-${pkg.version}.tgz`;
}
function run(command, args, cwd, shell = false) {
  const result = spawnSync(command, args, { cwd, stdio: 'inherit', shell });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${command} failed (${result.status}) in ${cwd}`);
}
const tarballs = {};
run(process.execPath, [join(root, 'scripts', 'check-package-content.mjs')], root);
for (const name of ['core', 'editor', 'oojson']) {
  const dir = join(root, 'packages', name);
  const pkg = JSON.parse(await readFile(join(dir, 'package.json'), 'utf8'));
  run(process.execPath, [pnpmCli, 'pack', '--pack-destination', output], dir);
  tarballs[pkg.name] = `file:${join(output, tarballName(pkg)).replaceAll('\\', '/')}`;
}
// Windows can expose TEMP through an 8.3 short path while Vite expands HTML
// imports to the long path. Normalize once so its HTML proxy keys stay stable.
const consumer = await realpath(await mkdtemp(join(tmpdir(), 'coco-office-consumer-')));
await writeFile(
  join(consumer, 'package.json'),
  JSON.stringify(
    {
      name: 'coco-office-package-consumer',
      private: true,
      type: 'module',
      dependencies: tarballs,
      devDependencies: { vite: '5.4.21', typescript: '5.9.3' },
    },
    null,
    2,
  ),
);
await copyFile(join(root, '.npmrc'), join(consumer, '.npmrc'));
// Only fixed arguments enter the Windows shell; paths are passed via cwd/package.json.
run(
  'npm',
  ['install', '--ignore-scripts', '--no-audit', '--no-fund'],
  consumer,
  process.platform === 'win32',
);
await mkdir(join(consumer, 'examples'), { recursive: true });
for (const name of ['samples.mjs', 'roundtrip.mjs', 'oojson-sample.mjs', 'oojson-roundtrip.mjs'])
  await copyFile(join(root, 'examples', name), join(consumer, 'examples', name));
await copyFile(
  join(root, 'scripts', 'check-package-exports.mjs'),
  join(consumer, 'check-package-exports.mjs'),
);
run(process.execPath, ['check-package-exports.mjs'], consumer);
run(process.execPath, ['examples/roundtrip.mjs'], consumer);
await cp(join(root, 'packages/editor/demo'), join(consumer, 'packages/editor/demo'), {
  recursive: true,
});
await copyFile(
  join(root, 'packages/editor/vite.demo.config.ts'),
  join(consumer, 'packages/editor/vite.demo.config.ts'),
);
run(
  process.execPath,
  [
    join(consumer, 'node_modules/vite/bin/vite.js'),
    'build',
    '--config',
    'packages/editor/vite.demo.config.ts',
  ],
  consumer,
);
await writeFile(
  join(consumer, 'api-check.ts'),
  `
import { WordDocument } from '@coco-vexel/coco-office/word';
import { ExcelWorkbook } from '@coco-vexel/coco-office/excel';
import { Presentation, type PptChartDataLabelLayoutModel } from '@coco-vexel/coco-office/ppt';
import { renderNativePptEditor } from '@coco-vexel/coco-office-editor/ppt';
import { renderSpreadsheetEditor } from '@coco-vexel/coco-office-editor/spreadsheet';
const layout: PptChartDataLabelLayoutModel = { index: 0, x: 0.5 };
void [WordDocument, ExcelWorkbook, Presentation, renderNativePptEditor, renderSpreadsheetEditor, layout];
`,
);
run(
  process.execPath,
  [
    join(consumer, 'node_modules/typescript/bin/tsc'),
    '--noEmit',
    '--skipLibCheck',
    '--target',
    'ES2022',
    '--module',
    'NodeNext',
    '--moduleResolution',
    'NodeNext',
    'api-check.ts',
  ],
  consumer,
);
await writeFile(
  join(root, '.release-check', 'consumer.json'),
  JSON.stringify({ consumer, tarballs, passed: true }, null, 2),
);
console.log(`Standalone tarball consumer passed: ${consumer}`);
