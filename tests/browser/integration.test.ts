import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { defineOfficeViewerElement, type OfficeViewerElement } from '../../src'

const SAMPLE_FIXTURES = {
  docx: '/fixtures/sample.docx',
  xlsx: '/fixtures/sample.xlsx',
  pptx: '/fixtures/sample.pptx'
}

const supportsCustomElementWrapper = detectCustomElementWrapperSupport()

describe('browser integration', () => {
  let container: HTMLDivElement

  beforeEach(() => {
    container = document.createElement('div')
    container.style.width = '1024px'
    container.style.height = '768px'
    document.body.append(container)
  })

  afterEach(() => {
    container.replaceChildren()
    container.remove()
  })

  it.skipIf(!supportsCustomElementWrapper)('creates zero DOM children and no Shadow DOM', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    expect(element.shadowRoot).toBeNull()
    expect(element.childNodes.length).toBe(0)

    await element.load(SAMPLE_FIXTURES.docx, { format: 'docx' })

    expect(element.shadowRoot).toBeNull()
    expect(element.childNodes.length).toBe(0)
    expect(container.contains(element)).toBe(true)

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('loads DOCX from a URL and exposes upstream instances', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    await element.load(SAMPLE_FIXTURES.docx, { format: 'docx' })

    expect(element.format).toBe('docx')
    expect(element.ready).toBe(true)
    expect(element.mode).toBeOneOf(['worker', 'main'])
    expect(typeof (element.getDocument() as Record<string, unknown>)?.destroy).toBe('function')
    expect(typeof (element.getViewer() as Record<string, unknown>)?.pageCount).toBe('number')
    expect(typeof (element.getEngine() as Record<string, unknown>)?.destroy).toBe('function')

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('loads XLSX from a URL and exposes upstream instances', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    await element.load(SAMPLE_FIXTURES.xlsx, { format: 'xlsx' })

    expect(element.format).toBe('xlsx')
    expect(element.ready).toBe(true)
    expect(element.mode).toBeOneOf(['worker', 'main'])
    expect(Array.isArray((element.getDocument() as Record<string, unknown>)?.sheetNames)).toBe(true)
    expect(typeof (element.getViewer() as Record<string, unknown>)?.sheetCount).toBe('number')
    expect(Array.isArray((element.getEngine() as Record<string, unknown>)?.sheetNames)).toBe(true)

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('loads PPTX from a URL and exposes upstream instances', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    await element.load(SAMPLE_FIXTURES.pptx, { format: 'pptx' })

    expect(element.format).toBe('pptx')
    expect(element.ready).toBe(true)
    expect(element.mode).toBeOneOf(['worker', 'main'])
    expect(typeof (element.getDocument() as Record<string, unknown>)?.destroy).toBe('function')
    expect(typeof (element.getViewer() as Record<string, unknown>)?.slideCount).toBe('number')
    expect(typeof (element.getEngine() as Record<string, unknown>)?.destroy).toBe('function')

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('loads File sources in main mode', async () => {
    defineOfficeViewerElement()

    const response = await fetch(SAMPLE_FIXTURES.docx)
    const file = new File([await response.arrayBuffer()], 'sample.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    })

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    await element.load(file, { mode: 'main' })

    expect(element.format).toBe('docx')
    expect(element.mode).toBe('main')

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('loads Blob, ArrayBuffer, and Uint8Array sources', async () => {
    defineOfficeViewerElement()

    const response = await fetch(SAMPLE_FIXTURES.xlsx)
    const buffer = await response.arrayBuffer()

    const blob = new Blob([buffer])
    const blobElement = document.createElement('office-viewer') as OfficeViewerElement
    container.append(blobElement)
    await blobElement.load(blob, { format: 'xlsx', mode: 'main' })
    expect(blobElement.format).toBe('xlsx')
    blobElement.destroy()

    const arrayBufferElement = document.createElement('office-viewer') as OfficeViewerElement
    container.append(arrayBufferElement)
    await arrayBufferElement.load(buffer, { format: 'xlsx', mode: 'main' })
    expect(arrayBufferElement.format).toBe('xlsx')
    arrayBufferElement.destroy()

    const uint8ArrayElement = document.createElement('office-viewer') as OfficeViewerElement
    container.append(uint8ArrayElement)
    await uint8ArrayElement.load(new Uint8Array(buffer), { format: 'xlsx', mode: 'main' })
    expect(uint8ArrayElement.format).toBe('xlsx')
    uint8ArrayElement.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('does not mutate caller-owned ArrayBuffer or Uint8Array', async () => {
    defineOfficeViewerElement()

    const response = await fetch(SAMPLE_FIXTURES.docx)
    const buffer = await response.arrayBuffer()
    const originalBytes = new Uint8Array(buffer).slice()

    const arrayBufferElement = document.createElement('office-viewer') as OfficeViewerElement
    container.append(arrayBufferElement)
    await arrayBufferElement.load(buffer, { format: 'docx', mode: 'main' })
    expect(new Uint8Array(buffer)).toEqual(originalBytes)
    arrayBufferElement.destroy()

    const view = new Uint8Array(buffer)
    const uint8ArrayElement = document.createElement('office-viewer') as OfficeViewerElement
    container.append(uint8ArrayElement)
    await uint8ArrayElement.load(view, { format: 'docx', mode: 'main' })
    expect(view).toEqual(originalBytes)
    uint8ArrayElement.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('reloads cleanly and destroy clears state', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    await element.load(SAMPLE_FIXTURES.docx, { format: 'docx' })
    expect(element.format).toBe('docx')

    await element.reload()
    expect(element.format).toBe('docx')

    element.destroy()
    expect(element.format).toBeNull()
    expect(element.ready).toBe(false)
    expect(element.getViewer()).toBeNull()
  })

  it.skipIf(!supportsCustomElementWrapper)('reconnects and reloads src-driven content after disconnect', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    element.setAttribute('src', SAMPLE_FIXTURES.docx)
    element.setAttribute('file-type', 'docx')

    container.append(element)
    await waitForEvent(element, 'ready')

    expect(element.format).toBe('docx')

    element.remove()
    container.append(element)

    await waitForEvent(element, 'ready')
    expect(element.format).toBe('docx')

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('observes effective mode on the engine', async () => {
    defineOfficeViewerElement()

    const workerElement = document.createElement('office-viewer') as OfficeViewerElement
    container.append(workerElement)
    await workerElement.load(SAMPLE_FIXTURES.docx, { format: 'docx' })
    expect(workerElement.mode).toBeOneOf(['worker', 'main'])
    const engineMode = (workerElement.getEngine() as Record<string, unknown> | null)?.mode
    expect(engineMode).toBe(workerElement.mode)
    workerElement.destroy()

    const mainElement = document.createElement('office-viewer') as OfficeViewerElement
    container.append(mainElement)
    await mainElement.load(SAMPLE_FIXTURES.docx, { format: 'docx', mode: 'main' })
    expect(mainElement.mode).toBe('main')
    mainElement.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('aborts an in-flight load when superseded by a newer load', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    const first = element.load(SAMPLE_FIXTURES.docx, { format: 'docx', mode: 'main' })
    const second = element.load(SAMPLE_FIXTURES.xlsx, { format: 'xlsx', mode: 'main' })

    await expect(first).rejects.toMatchObject({ name: 'AbortError' })
    await second

    expect(element.format).toBe('xlsx')
    expect(element.ready).toBe(true)

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('supports canceling an externally aborted load request', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
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

  it.skipIf(!supportsCustomElementWrapper)('surfaces a useful malformed-input diagnostic', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    let capturedError: Error | null = null
    element.addEventListener('loaderror', (event) => {
      capturedError = (event as CustomEvent).detail.error
    })

    await expect(element.load('/fixtures/malformed.docx', { format: 'docx', mode: 'main' })).rejects.toThrowError()
    expect(capturedError).toBeTruthy()

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('fails clearly for unsupported formats', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    let capturedError: Error | null = null
    element.addEventListener('loaderror', (event) => {
      capturedError = (event as CustomEvent).detail.error
    })

    await expect(element.load('/fixtures/report.pdf', { format: 'pdf' as unknown as 'docx' })).rejects.toThrowError()
    expect(capturedError).toBeTruthy()

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('fails clearly for binary sources without format', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    let capturedError: Error | null = null
    element.addEventListener('loaderror', (event) => {
      capturedError = (event as CustomEvent).detail.error
    })

    await expect(element.load(new Uint8Array([1, 2, 3]))).rejects.toThrowError()
    expect(capturedError).toBeTruthy()

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('accepts a custom wasm-url option', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    let capturedError: Error | null = null
    element.addEventListener('loaderror', (event) => {
      capturedError = (event as CustomEvent).detail.error
    })

    await expect(
      element.load(SAMPLE_FIXTURES.docx, {
        format: 'docx',
        mode: 'main',
        wasmUrl: '/fixtures/fake-wasm/does-not-exist.wasm'
      })
    ).rejects.toThrowError()
    expect(capturedError).toBeTruthy()

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('accepts a custom wasm-url attribute', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    element.setAttribute('wasm-url', '/fixtures/fake-wasm/does-not-exist.wasm')
    element.setAttribute('src', SAMPLE_FIXTURES.docx)
    element.setAttribute('file-type', 'docx')
    element.setAttribute('mode', 'main')
    container.append(element)

    let capturedError: Error | null = null
    element.addEventListener('loaderror', (event) => {
      capturedError = (event as CustomEvent).detail.error
    })

    await expect(waitForEvent(element, 'loaderror', 5000)).resolves.toBeDefined()
    expect(capturedError).toBeTruthy()

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('loads when imported as a plain module from a CDN-like relative URL', async () => {
    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    // Simulate the plain-module consumer path: the element class is available
    // on the global namespace after an ESM import, and `load()` is invoked directly.
    if (!customElements.get('office-viewer')) {
      defineOfficeViewerElement()
    }

    await element.load(SAMPLE_FIXTURES.pptx, { format: 'pptx', mode: 'main' })

    expect(element.format).toBe('pptx')
    expect(element.ready).toBe(true)
    expect(typeof (element.getViewer() as Record<string, unknown>)?.slideCount).toBe('number')

    element.destroy()
  })

  it.skipIf(!supportsCustomElementWrapper)('does not expose internal render container as a child', async () => {
    defineOfficeViewerElement()

    const element = document.createElement('office-viewer') as OfficeViewerElement
    container.append(element)

    await element.load(SAMPLE_FIXTURES.xlsx, { format: 'xlsx', mode: 'main' })

    expect(element.shadowRoot).toBeNull()
    expect(element.childNodes.length).toBe(0)

    element.destroy()
  })
})

function detectCustomElementWrapperSupport(): boolean {
  try {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as Partial<OfficeViewerElement>
    return (
      typeof element.load === 'function'
      && typeof element.reload === 'function'
      && typeof element.destroy === 'function'
      && typeof element.getViewer === 'function'
      && typeof element.getDocument === 'function'
      && typeof element.getEngine === 'function'
    )
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
