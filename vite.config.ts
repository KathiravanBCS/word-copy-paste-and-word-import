import { defineConfig } from 'vite';
import { resolve } from 'node:path';

// Dev / demo application config. The demo lives in src/demo and imports the
// engine directly from source so that debugging a real clipboard payload never
// requires a build step.
export default defineConfig({
  root: resolve(__dirname, 'src/demo'),
  base: './',
  server: { port: 5180, open: '/clipboard-lab/' },
  build: {
    outDir: resolve(__dirname, 'dist-demo'),
    emptyOutDir: true,
    rollupOptions: {
      input: {
        index: resolve(__dirname, 'src/demo/index.html'),
        lab: resolve(__dirname, 'src/demo/clipboard-lab/index.html'),
        rooster: resolve(__dirname, 'src/demo/rooster-editor/index.html'),
        tiptap: resolve(__dirname, 'src/demo/tiptap-editor/index.html'),
      },
    },
  },
});
