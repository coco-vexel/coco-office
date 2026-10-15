import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';

export default defineConfig({
  root: fileURLToPath(new URL('./demo', import.meta.url)),
  base: './',
  server: { host: '127.0.0.1' },
  build: {
    outDir: '../demo-dist',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        editor: fileURLToPath(new URL('./demo/index.html', import.meta.url)),
        oojson: fileURLToPath(new URL('./demo/oojson.html', import.meta.url)),
      },
    },
  },
});
