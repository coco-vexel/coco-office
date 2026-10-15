/* eslint-disable no-console -- CLI verification output. */
import { readFile, writeFile, readdir, realpath } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const license = await readFile(join(root, 'LICENSE'), 'utf8');
const normalizeText = (value) => value.replace(/\r\n?/g, '\n').trimEnd();
const workspacePackages = new Set([
  '@coco-vexel/coco-office',
  '@coco-vexel/coco-office-editor',
  '@coco-vexel/coco-office-oojson',
]);
for (const target of ['core', 'editor', 'oojson']) {
  const packageDir = join(root, 'packages', target);
  const collected = new Map();
  async function visit(dir) {
    const pkg = JSON.parse(await readFile(join(dir, 'package.json'), 'utf8'));
    const require = createRequire(join(dir, 'package.json'));
    for (const name of Object.keys(pkg.dependencies ?? {}).sort()) {
      if (workspacePackages.has(name)) {
        if (target === 'oojson' && name === '@coco-vexel/coco-office')
          await visit(join(root, 'packages', 'core'));
        continue;
      }
      let dependencyDir;
      try {
        dependencyDir = dirname(require.resolve(`${name}/package.json`));
      } catch {
        dependencyDir = dirname(require.resolve(name));
        for (;;) {
          try {
            const meta = JSON.parse(await readFile(join(dependencyDir, 'package.json'), 'utf8'));
            if (meta.name === name) break;
          } catch {
            // Entry files can be nested below the package metadata directory.
          }
          const parent = dirname(dependencyDir);
          if (parent === dependencyDir) throw new Error(`Cannot locate metadata for ${name}`);
          dependencyDir = parent;
        }
      }
      dependencyDir = await realpath(dependencyDir);
      const meta = JSON.parse(await readFile(join(dependencyDir, 'package.json'), 'utf8'));
      const key = `${meta.name}@${meta.version}`;
      if (collected.has(key)) continue;
      const files = (await readdir(dependencyDir))
        .filter((name) => /^(license|licence|copying|notice)([.-]|$)/i.test(name))
        .sort();
      const texts = await Promise.all(
        files.map(
          async (name) =>
            `### ${name}\n\n${normalizeText(await readFile(join(dependencyDir, name), 'utf8'))}`,
        ),
      );
      if (!texts.length) {
        const readmeName = (await readdir(dependencyDir)).find((name) =>
          /^readme(\.md)?$/i.test(name),
        );
        const readme = readmeName ? await readFile(join(dependencyDir, readmeName), 'utf8') : '';
        const section = readme.match(
          /^#{1,3}\s+Licen[cs]e\s*\r?\n([\s\S]*?)(?=^#{1,3}\s|$(?![\s\S]))/im,
        );
        if (!section || !section[1].includes('Permission is hereby granted'))
          throw new Error(`No license text found for ${key}; review manually.`);
        texts.push(`### ${readmeName} (license section)\n\n${section[1].trim()}`);
      }
      // Pako's top-level LICENSE covers its MIT wrapper; the embedded zlib
      // implementation has additional notices in source-file headers.
      if (meta.name === 'pako') {
        const zlibDir = join(dependencyDir, 'lib', 'zlib');
        const headers = new Map();
        for (const file of (await readdir(zlibDir)).filter((name) => name.endsWith('.js')).sort()) {
          const source = await readFile(join(zlibDir, file), 'utf8');
          const ending = 'This notice may not be removed or altered from any source distribution.';
          const start = source.indexOf('// (C)');
          const end = source.indexOf(ending);
          if (start < 0 || end < start)
            throw new Error(`Review embedded zlib license: ${key}/${file}`);
          const header = normalizeText(source.slice(start, end + ending.length));
          headers.set(header, [...(headers.get(header) ?? []), `lib/zlib/${file}`]);
        }
        for (const [header, paths] of headers) {
          texts.push(
            `### Embedded zlib notice: ${paths.join(', ')}\n\n\`\`\`text\n${header}\n\`\`\``,
          );
        }
      }
      collected.set(
        key,
        `## ${key}\n\nDeclared license: ${typeof meta.license === 'string' ? meta.license : JSON.stringify(meta.license)}\n\n${meta.name === 'jszip' ? 'License selected for Coco Office distribution: MIT (upstream offers MIT OR GPL).\n\n' : ''}${texts.join('\n\n')}`,
      );
      await visit(dependencyDir);
    }
  }
  await visit(packageDir);
  await writeFile(join(packageDir, 'LICENSE'), license);
  await writeFile(
    join(packageDir, 'THIRD_PARTY_NOTICES.md'),
    '# Third-party notices\n\nGenerated from installed runtime dependencies by `pnpm licenses:packages`.\n' +
      'Dependencies retain their own licenses. The Core package supplies its own notices.\n\n' +
      [...collected.entries()]
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([, text]) => text)
        .join('\n\n') +
      '\n',
  );
  console.log(`${target}: license and ${collected.size} dependency notices prepared`);
}
