import { readFile, stat } from 'node:fs/promises'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

describe('CDN / plain-module consumption', () => {
  const distDir = join(process.cwd(), 'dist')
  const bundlePath = join(distDir, 'office-viewer.es.js')
  const typesDir = join(distDir, 'types')
  const indexTypesPath = join(typesDir, 'index.d.ts')

  it('ships a single ESM bundle', async () => {
    const bundle = await readFile(bundlePath, 'utf-8')
    expect(bundle.length).toBeGreaterThan(0)
    expect(bundle).toContain('export')
    expect(bundle).not.toContain('import.meta.url')
  })

  it('ships generated TypeScript declarations', async () => {
    const indexTypes = await readFile(indexTypesPath, 'utf-8')
    expect(indexTypes).toContain('OfficeViewerElement')
    expect(indexTypes).toContain('defineOfficeViewerElement')

    const elementTypesPath = join(typesDir, 'office-viewer-element.d.ts')
    const elementTypes = await readFile(elementTypesPath, 'utf-8')
    expect(elementTypes).toContain('getViewer')
    expect(elementTypes).toContain('getDocument')
    expect(elementTypes).toContain('getEngine')

    const publicTypesPath = join(typesDir, 'types.d.ts')
    const publicTypes = await readFile(publicTypesPath, 'utf-8')
    expect(publicTypes).toContain('OfficeFormat')
    expect(publicTypes).toContain('OfficeSource')
  })

  it('does not bundle the upstream engine into the library entry', async () => {
    const bundle = await readFile(bundlePath, 'utf-8')
    // The orchestration layer should be small; the upstream parser assets are loaded on demand.
    const stats = await stat(bundlePath)
    expect(stats.size).toBeLessThan(100_000)
    expect(bundle).not.toContain('@silurus/ooxml')
  })

  it('declares the public API surface in package.json', async () => {
    const pkg = JSON.parse(await readFile(join(process.cwd(), 'package.json'), 'utf-8'))
    expect(pkg.main).toBe('./dist/office-viewer.es.js')
    expect(pkg.module).toBe('./dist/office-viewer.es.js')
    expect(pkg.types).toBe('./dist/types/index.d.ts')
    expect(pkg.exports['.'].types).toBe('./dist/types/index.d.ts')
    expect(pkg.exports['.'].import).toBe('./dist/office-viewer.es.js')
  })
})
