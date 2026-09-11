import { defineConfig } from 'vite';

/**
 * POC build configuration.
 *
 * - `npm run dev` serves the demo page directly (browse-first welcome screen).
 * - `npm run build` runs `tsc` first (type-check + declarations → dist/),
 *   then Vite bundles demo/index.html into dist-web/.
 */
export default defineConfig({
  server: {
    open: '/demo/index.html',
  },
  build: {
    // Separate from tsc's "dist" output so nothing gets wiped.
    outDir: 'dist-web',
    rollupOptions: {
      input: 'demo/index.html',
    },
  },
});