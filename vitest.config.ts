import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
    exclude: ['tests/browser/**/*.test.ts', 'tests/build-output.test.ts', 'tests/build-cdn.test.ts']
  }
})
