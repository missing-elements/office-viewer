import { resolve } from 'node:path'
import { defineConfig } from 'vite'

const rootDir = import.meta.dirname

export default defineConfig({
  base: './',
  root: resolve(rootDir, 'demo'),
  publicDir: resolve(rootDir, 'public'),
  server: {
    host: '127.0.0.1',
    port: 5173,
    strictPort: true,
    open: true
  },
  build: {
    lib: {
      entry: resolve(rootDir, 'src/index.ts'),
      name: 'OfficeViewer',
      fileName: (format) => `office-viewer.${format}.js`,
      formats: ['es']
    },
    rollupOptions: {
      external: ['@silurus/ooxml'],
      output: {
        entryFileNames: 'office-viewer.es.js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash][extname]',
        globals: {}
      }
    }
  }
})
