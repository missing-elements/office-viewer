import { resolve } from 'node:path'
import { defineConfig } from 'vite'

// Dev server and static build for the demo. The demo imports the package by name;
// this alias points that at the source so the demo always exercises the current code
// and the built site under demo/dist is self-contained.
const rootDir = import.meta.dirname

export default defineConfig({
  root: resolve(rootDir, 'demo'),
  publicDir: resolve(rootDir, 'public'),
  resolve: {
    alias: {
      '@missing-elements/office-viewer': resolve(rootDir, 'src/index.ts')
    }
  },
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
    outDir: 'dist',
    emptyOutDir: true
  }
})
