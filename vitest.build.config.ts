import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/build-output.test.ts', 'tests/build-cdn.test.ts']
  }
})
