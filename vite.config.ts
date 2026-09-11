import { resolve } from 'node:path'
import { defineConfig } from 'vite'

const rootDir = import.meta.dirname

export default defineConfig({
  base: './',
  root: resolve(rootDir),
  server: {
    host: '127.0.0.1',
    port: 5173,
    strictPort: true,
    open: true
  },
  optimizeDeps: {
    // Pre-bundle every format entrypoint at server start. The viewer imports each
    // format on demand (dynamic import)
    include: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx']
  },
  build: {
    lib: {
      entry: resolve(rootDir, 'src/index.ts'),
      name: 'OfficeViewer',
      fileName: (format) => `office-viewer.${format}.js`,
      formats: ['es']
    },
    rollupOptions: {
      output: {
        entryFileNames: 'office-viewer.es.js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash][extname]',
        globals: {}
      }
    },
    outDir: resolve(rootDir, 'dist'),
    copyPublicDir: false,
    emptyOutDir: true
  }
})
