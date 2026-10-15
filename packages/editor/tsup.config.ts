import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { defineConfig } from 'tsup';

const release = process.env.COCO_RELEASE === '1';
const editorCssPath = resolve('src/styles/editor.css');
const generatedStylesPath = resolve('src/generated/editor-styles.ts');
mkdirSync(resolve('src/generated'), { recursive: true });
writeFileSync(generatedStylesPath, `export default ${JSON.stringify(readFileSync(editorCssPath, 'utf8'))};\n`);

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    word: 'src/word/index.ts',
    spreadsheet: 'src/spreadsheet/index.ts',
    ppt: 'src/ppt/index.ts',
    'ppt-native': 'src/ppt/native.ts',
    'word-editing': 'src/word/editing.ts',
    'word-native': 'src/word/native.ts',
    'word-prosemirror': 'src/word/prosemirror.ts',
  },
  format: ['esm', 'cjs'],
  dts: true,
  clean: true,
  sourcemap: !release,
  minify: release,
  keepNames: !release,
  target: 'es2017',
  splitting: false,
  treeshake: true,
  esbuildOptions(options) {
    options.legalComments = 'eof';
    options.sourcesContent = !release;
  },
});
