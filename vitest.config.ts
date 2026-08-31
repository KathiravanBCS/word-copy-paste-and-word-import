import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
    environment: 'happy-dom',
    // The security fixtures deliberately contain an <iframe> and remote
    // references. happy-dom would otherwise try to fetch them, which a real
    // browser's DOMParser never does — its documents are inert.
    environmentOptions: {
      happyDOM: {
        settings: {
          disableIframePageLoading: true,
          disableJavaScriptEvaluation: true,
          disableJavaScriptFileLoading: true,
          disableCSSFileLoading: true,
        },
      },
    },
    include: ['src/tests/**/*.test.ts'],
    reporters: ['default'],
  },
});
