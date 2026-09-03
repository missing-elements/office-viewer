import { readdir } from 'node:fs/promises'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

describe('production build output', () => {
  it('emits wasm assets and render workers', async () => {
    const assetsDir = join(process.cwd(), 'dist', 'assets')
    const assetNames = await readdir(assetsDir)

    const wasmAssets = assetNames.filter((name) => name.endsWith('.wasm'))
    const renderWorkers = assetNames.filter((name) => name.includes('render-worker') && name.endsWith('.js'))

    expect(wasmAssets.length).toBeGreaterThanOrEqual(3)
    expect(renderWorkers.length).toBeGreaterThanOrEqual(3)
  })
})
