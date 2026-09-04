import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { defineOfficeViewerElement, OoxmlIntegrationSpike, SAMPLE_FIXTURES, type OfficeViewerElement } from '../../src'

const supportsCustomElementWrapper = detectCustomElementWrapperSupport()

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

  it.skipIf(!supportsCustomElementWrapper)('supports the custom-element wrapper API for the same fixture path', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    element.style.width = '100%'
    element.style.height = '100%'
    container.append(element)

    const result = await element.load(SAMPLE_FIXTURES.docx, { format: 'docx' })
    expect(result.format).toBe('docx')
    expect(element.getSummary()?.format).toBe('docx')

    const viewport = (element.shadowRoot ?? element).querySelector<HTMLElement>('[part="viewport"]')
    expect(viewport).toBeTruthy()
    expect(viewport?.childElementCount ?? 0).toBeGreaterThan(0)

    element.destroy()
    expect(viewport?.childElementCount ?? 0).toBe(0)
  })

  it.skipIf(!supportsCustomElementWrapper)('reconnects and reloads src-driven content after disconnect', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    element.setAttribute('src', SAMPLE_FIXTURES.docx)
    element.setAttribute('file-type', 'docx')
    element.style.width = '100%'
    element.style.height = '100%'

    container.append(element)
    await waitForEvent(element, 'ready')

    expect(element.getSummary()?.format).toBe('docx')

    element.remove()
    container.append(element)

    await waitForEvent(element, 'ready')
    expect(element.getSummary()?.format).toBe('docx')
  })

  it.skipIf(!supportsCustomElementWrapper)('supports canceling an externally aborted load request', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    element.style.width = '100%'
    element.style.height = '100%'
    container.append(element)

    const controller = new AbortController()
    controller.abort('manual cancellation')

    await expect(
      element.load(SAMPLE_FIXTURES.docx, {
        format: 'docx',
        signal: controller.signal
      })
    ).rejects.toMatchObject({ name: 'AbortError' })
  })
})

function detectCustomElementWrapperSupport(): boolean {
  try {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as Partial<OfficeViewerElement>
    return typeof element.load === 'function' && typeof element.getSummary === 'function' && typeof element.destroy === 'function'
  } catch {
    return false
  }
}

function waitForEvent(target: EventTarget, eventName: string, timeoutMs = 10_000): Promise<Event> {
  return new Promise((resolve, reject) => {
    const timeoutId = window.setTimeout(() => {
      target.removeEventListener(eventName, onEvent)
      reject(new Error(`Timed out waiting for ${eventName}`))
    }, timeoutMs)

    const onEvent = (event: Event) => {
      window.clearTimeout(timeoutId)
      target.removeEventListener(eventName, onEvent)
      resolve(event)
    }

    target.addEventListener(eventName, onEvent, { once: true })
  })
}
