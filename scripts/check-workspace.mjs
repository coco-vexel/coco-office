/* eslint-disable no-console -- Workspace check wrapper. */
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const pnpmCli = process.env.npm_execpath;
if (!pnpmCli) throw new Error('Run this script through pnpm.');

function run(args) {
  const result = spawnSync(process.execPath, [pnpmCli, ...args], {
    cwd: root,
    stdio: 'inherit',
  });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

run(['build']);
run(['test:ci']);
console.log('Workspace build and CI tests passed.');
