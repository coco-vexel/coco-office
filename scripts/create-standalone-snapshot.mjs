/* eslint-disable no-console -- CLI verification output. */
import { execFileSync } from 'node:child_process';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const dirtyTracked = execFileSync('git', ['status', '--porcelain', '--untracked-files=no'], {
  cwd: root,
  encoding: 'utf8',
}).trim();
if (dirtyTracked)
  throw new Error(
    'Tracked files are modified. Commit or stash them before creating a public snapshot.',
  );

const publicRootFiles = new Set([
  '.gitignore',
  '.npmrc',
  '.prettierrc.json',
  'CONTRIBUTING.md',
  'LICENSE',
  'LICENSING.md',
  'README.md',
  'README.zh-CN.md',
  'SECURITY.md',
  'eslint.config.js',
  'package.json',
  'pnpm-lock.yaml',
  'pnpm-workspace.yaml',
  'tsconfig.base.json',
]);
const publicDocFiles = new Set([
  'docs/npm-release.md',
  'docs/open-source-readiness.md',
  'docs/runtime-and-demos.md',
  'docs/word-editor-capability-matrix.md',
]);
const excludedCoreDocs = new Set([
  'docs/core/advanced/agent-notes.md',
  'docs/core/advanced/codeplan.md',
  'docs/core/ppt/sdk-plan.md',
  'docs/core/ppt/sdk-review.md',
]);

function isPublic(path) {
  if (publicRootFiles.has(path) || publicDocFiles.has(path)) return true;
  if (path.startsWith('.github/')) return true;
  if (path.startsWith('examples/')) return true;
  if (path.startsWith('packages/')) return true;
  if (path.startsWith('scripts/')) return true;
  if (path.startsWith('docs/core/') && !excludedCoreDocs.has(path)) return true;
  return false;
}

const tracked = execFileSync('git', ['ls-files', '-z'], {
  cwd: root,
  encoding: 'utf8',
})
  .split('\0')
  .filter(Boolean);
const files = tracked.filter(isPublic);
const sourceCommit = execFileSync('git', ['rev-parse', 'HEAD'], {
  cwd: root,
  encoding: 'utf8',
}).trim();
const snapshot = await mkdtemp(join(tmpdir(), 'coco-office-standalone-'));

for (const path of files) {
  const target = resolve(snapshot, path);
  if (!target.startsWith(snapshot + '\\') && !target.startsWith(snapshot + '/'))
    throw new Error(`Invalid path: ${path}`);
  await mkdir(dirname(target), { recursive: true });
  const contents = execFileSync('git', ['show', `${sourceCommit}:${path}`], {
    cwd: root,
    maxBuffer: 64 * 1024 * 1024,
  });
  await writeFile(target, contents);
}
await mkdir(join(root, '.release-check'), { recursive: true });
await writeFile(
  join(root, '.release-check', 'snapshot.json'),
  `${JSON.stringify(
    {
      snapshot,
      sourceCommit,
      files: files.length,
      excludedTrackedFiles: tracked.length - files.length,
    },
    null,
    2,
  )}\n`,
);
console.log(JSON.stringify({ snapshot, sourceCommit, files: files.length }));
