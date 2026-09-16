import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';

// Bundles the whole app (including Recharts, PapaParse, and SheetJS) into a
// single self-contained dist/index.html — no separate JS/CSS files, so it
// can be dropped anywhere (GitHub Pages, a static file share, a single
// email attachment) and just work.
export default defineConfig({
  plugins: [react(), viteSingleFile()],
  build: {
    cssCodeSplit: false,
    assetsInlineLimit: 100000000,
    chunkSizeWarningLimit: 100000,
    outDir: 'dist',
  },
});
