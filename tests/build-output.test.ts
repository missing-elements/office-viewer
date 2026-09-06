import { readdir } from 'node:fs/promises'
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

    expect(wasmAssets.length).toBeGreaterThanOrEqual(3)
    expect(renderWorkers.length).toBeGreaterThanOrEqual(3)
    expect(fixtureNames).toEqual(expect.arrayContaining(['sample.docx', 'sample.xlsx', 'sample.pptx']))
  })
})
