import { defineConfig } from 'vite';
import { resolve } from 'node:path';

// Library build: the engine ships as a dependency-free ES module.
export default defineConfig({
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      formats: ['es'],
      fileName: () => 'index.js',
    },
    rollupOptions: { external: [] },
    sourcemap: true,
    target: 'es2022',
  },
});
