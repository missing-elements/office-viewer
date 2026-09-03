import { resolve } from 'node:path'
import { defineConfig } from 'vite'

const rootDir = import.meta.dirname

export default defineConfig({
  base: './',
  build: {
    rollupOptions: {
      input: {
        demo: resolve(rootDir, 'demo/index.html'),
        'measure-combined': resolve(rootDir, 'demo/measure-combined.html'),
        'measure-docx': resolve(rootDir, 'demo/measure-docx.html'),
        'measure-xlsx': resolve(rootDir, 'demo/measure-xlsx.html'),
        'measure-pptx': resolve(rootDir, 'demo/measure-pptx.html')
      }
    }
  }
})
