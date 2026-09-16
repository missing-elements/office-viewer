import { resolve } from 'node:path'
import { defineConfig } from 'vite'

// Library build only. The demo has its own config in vite.demo.config.ts.
const rootDir = import.meta.dirname

export default defineConfig({
  base: './',
  root: resolve(rootDir),
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
