import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { OoxmlIntegrationSpike, SAMPLE_FIXTURES } from '../../src'

describe('browser integration spike', () => {
  let container: HTMLDivElement
  let harness: OoxmlIntegrationSpike

  beforeEach(() => {
    container = document.createElement('div')
    container.style.width = '1024px'
    container.style.height = '768px'
    document.body.append(container)
    harness = new OoxmlIntegrationSpike(container)
  })

  afterEach(() => {
    harness.destroy()
    container.remove()
  })

  it('loads the representative fixtures in worker mode by default', async () => {
    const docx = await harness.load(SAMPLE_FIXTURES.docx, { format: 'docx' })
    expect(docx.requestedMode).toBe('worker')
    expect(['worker', 'main']).toContain(docx.effectiveMode)

    const xlsx = await harness.load(SAMPLE_FIXTURES.xlsx, { format: 'xlsx' })
    expect(xlsx.totalCount).toBeGreaterThanOrEqual(1)
    expect(xlsx.sheetNames?.length).toBeGreaterThanOrEqual(1)

    const pptx = await harness.load(SAMPLE_FIXTURES.pptx, { format: 'pptx' })
    expect(pptx.totalCount).toBeGreaterThanOrEqual(1)
    expect(['worker', 'main']).toContain(pptx.effectiveMode)
  })

  it('supports main mode and File inputs', async () => {
    const response = await fetch(SAMPLE_FIXTURES.docx)
    const file = new File([await response.arrayBuffer()], 'sample.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
    const result = await harness.load(file, { mode: 'main' })

    expect(result.sourceKind).toBe('file')
    expect(result.requestedMode).toBe('main')
    expect(result.effectiveMode).toBe('main')
  })

  it('reloads cleanly and can be reused after destroy', async () => {
    const first = await harness.load(SAMPLE_FIXTURES.docx, { format: 'docx' })
    const reloaded = await harness.reload()

    expect(reloaded.generation).toBeGreaterThan(first.generation)
    expect(container.childElementCount).toBeGreaterThan(0)

    harness.destroy()
    expect(container.childElementCount).toBe(0)

    const second = await harness.load(SAMPLE_FIXTURES.xlsx, { format: 'xlsx' })
    expect(second.format).toBe('xlsx')
    expect(container.childElementCount).toBeGreaterThan(0)
  })

  it('surfaces a useful malformed-input diagnostic', async () => {
    await expect(
      harness.load('/fixtures/malformed.docx', { format: 'docx', mode: 'main' })
    ).rejects.toThrowError()

    const summary = harness.getSummary()
    expect(summary?.lastError?.name).toBeTruthy()
    expect(summary?.lastError?.message.length).toBeGreaterThan(0)
    expect(summary?.diagnostics.at(-1)?.length).toBeGreaterThan(0)
  })
})
