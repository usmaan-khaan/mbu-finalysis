import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// The post-build script inlines generated assets into dist/index.html.
export default defineConfig({
  plugins: [react()],
  build: {
    cssCodeSplit: false,
    assetsInlineLimit: 100000000,
    chunkSizeWarningLimit: 100000,
    outDir: 'dist',
  },
});
