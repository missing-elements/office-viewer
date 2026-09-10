import { webdriverio } from '@vitest/browser-webdriverio'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  // Pre-bundle every format entrypoint: the viewer imports each format on demand,
  // so without this the browser dev server can 504 on the first lazy request.
  optimizeDeps: {
    include: ['@silurus/ooxml/docx', '@silurus/ooxml/xlsx', '@silurus/ooxml/pptx']
  },
  test: {
    include: ['tests/browser/**/*.test.ts'],
    browser: {
      enabled: true,
      provider: webdriverio(),
      instances: [{ browser: 'firefox' }]
    }
  }
})
