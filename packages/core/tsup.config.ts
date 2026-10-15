import { defineConfig } from 'tsup';

const release = process.env.COCO_RELEASE === '1';

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    pack: 'src/pack/index.ts',
    shared: 'src/shared/index.ts',
    word: 'src/word/index.ts',
    'word-view': 'src/word/view.ts',
    excel: 'src/excel/index.ts',
    ppt: 'src/ppt/index.ts',
    // Worker entry:Node worker_threads 的 kernel bootstrap(#17)
    'kernel-worker': 'src/word/kernel-worker-entry.ts',
    // Worker entry:浏览器 Worker(WPS Chromium / 普通浏览器)的 kernel bootstrap
    'kernel-worker-browser': 'src/word/kernel-worker-browser-entry.ts',
    'excel-worker': 'src/excel/runtime-worker-entry.ts',
    'excel-worker-browser': 'src/excel/runtime-worker-browser-entry.ts',
  },
  format: ['esm', 'cjs'],
  dts: true,
  clean: true,
  sourcemap: !release,
  minify: release,
  keepNames: !release,
  target: 'es2017',
  splitting: false,
  // Keep lazy Node worker imports lazy. The optional Rollup tree-shaking pass
  // hoists them into static imports, breaking browser consumers of word/excel.
  treeshake: false,
  esbuildOptions(options) {
    options.treeShaking = true;
    options.legalComments = 'eof';
    options.sourcesContent = !release;
  },
});
