import { resolve } from 'node:path'
import { defineConfig } from 'vite'

const rootDir = import.meta.dirname

export default defineConfig({
  base: './',
  root: resolve(rootDir),
  publicDir: resolve(rootDir, 'public'),
  server: {
    host: '127.0.0.1',
    port: 5173,
    strictPort: true,
    open: true
  },
  optimizeDeps: {
    // Pre-bundle every format entrypoint at server start. The viewer imports each
    // format on demand (dynamic import), so without this Vite discovers them lazily
    // and the first request can fail with "504 Outdated Optimize Dep".
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
      external: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx'],
      output: {
        entryFileNames: 'office-viewer.es.js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash][extname]',
        globals: {}
      }
    },
    outDir: resolve(rootDir, 'dist'),
    emptyOutDir: true
  }
})
