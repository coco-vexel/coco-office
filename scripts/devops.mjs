/* eslint-disable no-console -- CLI help output. */
const commands = [
  ['Build packages', 'corepack pnpm build'],
  ['Run CI tests', 'corepack pnpm test:ci'],
  ['Run package consumer verification', 'corepack pnpm verify:packages'],
  ['Build offline SDK bundle', 'corepack pnpm bundle:sdk'],
  ['Prepare package notices', 'corepack pnpm licenses:packages'],
  ['Build browser demo', 'corepack pnpm build:demo'],
];

for (const [label, command] of commands) {
  console.log(`${label}: ${command}`);
}
