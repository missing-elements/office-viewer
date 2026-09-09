import { readFile, readdir } from 'node:fs/promises'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

describe('production build output', () => {
  it('emits parser assets, render workers, and demo fixtures', async () => {
    const assetsDir = join(process.cwd(), 'dist', 'assets')
    const fixturesDir = join(process.cwd(), 'dist', 'fixtures')
    const assetNames = await readdir(assetsDir)
    const fixtureNames = await readdir(fixturesDir)

    const wasmAssets = assetNames.filter((name) => name.endsWith('.wasm'))
    const renderWorkers = assetNames.filter((name) => name.includes('render-worker') && name.endsWith('.js'))

    // WASM may be bundled as base64 data URLs inside format chunks; verify worker hosts and format chunks instead.
    expect(renderWorkers.length).toBeGreaterThanOrEqual(3)
    const formatChunks = assetNames.filter((name) => /^docx-.*\.js$/.test(name) || /^xlsx-.*\.js$/.test(name) || /^pptx-.*\.js$/.test(name))
    expect(formatChunks.length).toBeGreaterThanOrEqual(3)
    expect(fixtureNames).toEqual(expect.arrayContaining(['sample.docx', 'sample.xlsx', 'sample.pptx']))
  })

  it('keeps the library entry small and free of upstream internals', async () => {
    const bundlePath = join(process.cwd(), 'dist', 'office-viewer.es.js')
    const content = await readFile(bundlePath, 'utf-8')

    expect(content.length).toBeLessThan(100_000)
    expect(content).not.toContain('@silurus/ooxml')
    expect(content).not.toContain('data:application/wasm')
  })

  it('keeps render-worker URLs relative and bundler-managed', async () => {
    const assetsDir = join(process.cwd(), 'dist', 'assets')
    const assetNames = await readdir(assetsDir)
    const hostAssets = assetNames.filter((name) => name.includes('render-worker-host') && name.endsWith('.js'))

    expect(hostAssets.length).toBeGreaterThanOrEqual(3)

    for (const name of hostAssets) {
      const content = await readFile(join(assetsDir, name), 'utf-8')
      // Worker host modules use import.meta.url to resolve workers; they must not hardcode dev URLs.
      expect(content).toContain('import.meta.url')
      expect(content).not.toContain('localhost')
      expect(content).not.toContain('127.0.0.1')
    }
  })
})
