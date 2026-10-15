import { defineConfig } from 'vitest/config';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageDir = fileURLToPath(new URL('.', import.meta.url));

export default defineConfig({
  root: '.',
  resolve: {
    alias: {
      '@coco-vexel/coco-office/pack': resolve(packageDir, '../core/src/pack/index.ts'),
      '@coco-vexel/coco-office/excel': resolve(packageDir, '../core/src/excel/index.ts'),
      '@coco-vexel/coco-office/ppt': resolve(packageDir, '../core/src/ppt/index.ts'),
      '@coco-vexel/coco-office/word': resolve(packageDir, '../core/src/word/index.ts'),
      '@coco-vexel/coco-office/word-view': resolve(packageDir, '../core/src/word/view.ts'),
    },
  },
  test: {
    include: ['src/**/*.test.ts'],
  },
});
