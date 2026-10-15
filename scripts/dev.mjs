import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const pnpmCli = process.env.npm_execpath;
if (!pnpmCli) throw new Error('Run this script through pnpm.');

const target = process.argv[2] ?? 'editor';
const script = target === 'demo' ? 'demo' : 'dev:editor';
const result = spawnSync(process.execPath, [pnpmCli, script], {
  cwd: root,
  stdio: 'inherit',
});
if (result.error) throw result.error;
process.exit(result.status ?? 1);
