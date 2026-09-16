import { execSync } from 'node:child_process'
import { mkdtempSync, readdirSync, readFileSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { beforeAll, describe, expect, it } from 'vitest'

// Guards the published shape of the library: one small entry that leaves
// @silurus/ooxml external. Bundling it inlines WASM and workers as base64 and
// produces a 14 MB package that CDN transpilers such as esm.sh time out on.
const root = resolve(import.meta.dirname, '..')
const outDir = mkdtempSync(join(tmpdir(), 'office-viewer-dist-'))
const entry = join(outDir, 'office-viewer.es.js')

describe('library build output', () => {
  beforeAll(() => {
    execSync(`pnpm exec vite build --outDir ${JSON.stringify(outDir)}`, { cwd: root, stdio: 'pipe' })
  }, 120_000)

  it('emits a single small ES module entry', () => {
    expect(readdirSync(outDir)).toEqual(['office-viewer.es.js'])
    expect(statSync(entry).size).toBeLessThan(50_000)
  })

  it('leaves @silurus/ooxml external so consumers and CDNs resolve it themselves', () => {
    const source = readFileSync(entry, 'utf8')
    for (const format of ['docx', 'xlsx', 'pptx']) {
      expect(source).toContain(`import("@silurus/ooxml/${format}")`)
    }
    expect(source).not.toMatch(/_parser_bg\.wasm|render-worker|@silurus\/ooxml\/dist/)
  })
})
