import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { defineOfficeViewerElement, OFFICE_FORMATS, OfficeViewerElement } from '../src'
import type { OfficeFormat, OfficeViewerLoadErrorDetail, OfficeViewerLoadOptions } from '../src'
import { FakeDocxViewer, FakePptxViewer, FakeViewer, FakeXlsxViewer } from './helpers/fake-viewer'

vi.mock('@silurus/ooxml/docx', async () => ({
  DocxScrollViewer: (await import('./helpers/fake-viewer')).FakeDocxViewer
}))
vi.mock('@silurus/ooxml/xlsx', async () => ({
  XlsxViewer: (await import('./helpers/fake-viewer')).FakeXlsxViewer
}))
vi.mock('@silurus/ooxml/pptx', async () => ({
  PptxScrollViewer: (await import('./helpers/fake-viewer')).FakePptxViewer
}))

const LIFECYCLE_EVENTS = ['loadstart', 'ready', 'loaderror', 'destroy']

function createElement(): OfficeViewerElement {
  const element = document.createElement('office-viewer') as OfficeViewerElement
  document.body.append(element)
  return element
}

function recordEvents(element: EventTarget, names = LIFECYCLE_EVENTS): string[] {
  const events: string[] = []
  for (const name of names) {
    element.addEventListener(name, () => events.push(name))
  }
  return events
}

function once(element: EventTarget, type: string): Promise<Event> {
  return new Promise((resolve) => element.addEventListener(type, resolve, { once: true }))
}

function settled(element: OfficeViewerElement): Promise<Event> {
  return Promise.race([once(element, 'ready'), once(element, 'loaderror')])
}

// A macrotask boundary so every queued microtask (attribute sync, deferred teardown) has run.
function flush(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0))
}

function viewerCount(count: number): Promise<void> {
  return vi.waitFor(() => {
    expect(FakeViewer.instances).toHaveLength(count)
  })
}

function bytesOf(source: string | ArrayBuffer | undefined): number[] {
  expect(source).toBeInstanceOf(ArrayBuffer)
  return Array.from(new Uint8Array(source as ArrayBuffer))
}

function streamOf(...chunks: number[][]): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) {
        controller.enqueue(Uint8Array.from(chunk))
      }
      controller.close()
    }
  })
}

async function loadViaAttributes(
  element: OfficeViewerElement,
  src: string,
  format: string,
  extra: Record<string, string> = {}
): Promise<Event> {
  element.setAttribute('src', src)
  element.setAttribute('file-type', format)
  for (const [name, value] of Object.entries(extra)) {
    element.setAttribute(name, value)
  }
  return settled(element)
}

// Moves the element to a new parent synchronously: disconnect and reconnect in one task.
async function moveSynchronously(element: OfficeViewerElement): Promise<void> {
  const wrapper = document.createElement('div')
  document.body.append(wrapper)
  wrapper.append(element)
  await flush()
}

describe('office-viewer element', () => {
  beforeEach(() => {
    defineOfficeViewerElement()
    FakeViewer.reset()
  })

  afterEach(async () => {
    for (const element of document.querySelectorAll('office-viewer')) {
      ;(element as OfficeViewerElement).destroy()
      element.remove()
    }
    await flush()
  })

  describe('definition and idle state', () => {
    it('exports the headless custom-element surface', () => {
      expect(OfficeViewerElement).toBeTypeOf('function')
      expect(defineOfficeViewerElement).toBeTypeOf('function')
      expect(defineOfficeViewerElement()).toBe(OfficeViewerElement)
      expect(OFFICE_FORMATS).toEqual(['docx', 'xlsx', 'pptx'])
    })

    it('creates an open Shadow DOM with the #viewer container on connect', () => {
      const element = createElement()
      expect(element.shadowRoot?.mode).toBe('open')
      expect(element.shadowRoot?.getElementById('viewer')).not.toBeNull()
    })

    it('reports idle state before any load', () => {
      const element = createElement()
      expect(element.ready).toBe(false)
      expect(element.error).toBeNull()
      expect(element.format).toBeNull()
      expect(element.mode).toBeNull()
      expect(element.getViewer()).toBeNull()
    })
  })

  describe('load()', () => {
    it('emits loadstart then ready and exposes the upstream viewer', async () => {
      const element = createElement()
      const events = recordEvents(element)

      await element.load('/sample.docx', { format: 'docx' })

      expect(events).toEqual(['loadstart', 'ready'])
      expect(element.getViewer()).toBeInstanceOf(FakeDocxViewer)
      expect(element.getViewer()).toBe(FakeViewer.last)
      expect(FakeViewer.last.container).toBe(element.shadowRoot?.getElementById('viewer'))
      expect(FakeViewer.last.sources).toEqual(['/sample.docx'])
      expect(element.ready).toBe(true)
      expect(element.error).toBeNull()
      expect(element.format).toBe('docx')
    })

    it('reports the mode the viewer was created with, defaulting to worker', async () => {
      const element = createElement()

      await element.load('/sample.docx', { format: 'docx' })
      expect(element.mode).toBe('worker')
      expect(FakeViewer.last.options.mode).toBe('worker')

      await element.load('/sample.docx', { format: 'docx', mode: 'main' })
      expect(element.mode).toBe('main')
      expect(FakeViewer.last.options.mode).toBe('main')
    })

    it('selects the upstream viewer for each format and forwards wasmUrl', async () => {
      const element = createElement()

      await element.load('/sample.xlsx', { format: 'xlsx', wasmUrl: 'https://cdn.example/xlsx.wasm' })
      expect(element.getViewer()).toBeInstanceOf(FakeXlsxViewer)
      expect(FakeViewer.last.options.wasmUrl).toBe('https://cdn.example/xlsx.wasm')

      await element.load('/sample.pptx', { format: 'pptx' })
      expect(element.getViewer()).toBeInstanceOf(FakePptxViewer)
      expect(element.format).toBe('pptx')
    })

    it('rejects an unsupported format with loaderror before creating a viewer', async () => {
      const element = createElement()
      const events = recordEvents(element)

      await expect(element.load('/sample.pdf', { format: 'pdf' as OfficeFormat })).rejects.toThrow(
        'Unsupported format "pdf"'
      )

      expect(events).toEqual(['loadstart', 'loaderror'])
      expect(FakeViewer.instances).toHaveLength(0)
      expect(element.error?.message).toContain('Unsupported format "pdf"')
      expect(element.ready).toBe(false)
    })

    it('rejects an unsupported source type with loaderror before creating a viewer', async () => {
      const element = createElement()
      const errors: Error[] = []
      element.addEventListener('loaderror', (event) => {
        errors.push((event as CustomEvent<OfficeViewerLoadErrorDetail>).detail.error)
      })

      await expect(element.load(42 as never, { format: 'docx' })).rejects.toThrow('Unsupported source type')

      expect(errors).toHaveLength(1)
      expect(errors[0]).toBe(element.error)
      expect(FakeViewer.instances).toHaveLength(0)
    })

    it('resolves Blob sources to an ArrayBuffer', async () => {
      const element = createElement()

      await element.load(new Blob([Uint8Array.from([1, 2, 3, 4])]), { format: 'docx' })

      expect(bytesOf(FakeViewer.last.sources[0])).toEqual([1, 2, 3, 4])
    })

    it('resolves ReadableStream sources to an ArrayBuffer', async () => {
      const element = createElement()

      await element.load(streamOf([1, 2], [3]), { format: 'docx' })

      expect(bytesOf(FakeViewer.last.sources[0])).toEqual([1, 2, 3])
    })

    it('stores the error, destroys the failed viewer, and keeps the previous document', async () => {
      const element = createElement()
      await element.load('/first.docx', { format: 'docx' })
      const first = FakeViewer.last
      const events = recordEvents(element)

      FakeViewer.loadBehavior = 'hold'
      const failing = element.load('/second.xlsx', { format: 'xlsx' })
      await viewerCount(2)
      const second = FakeViewer.last
      second.fail(new Error('boom'))
      await expect(failing).rejects.toThrow('boom')

      expect(events).toEqual(['loadstart', 'loaderror'])
      expect(element.error?.message).toBe('boom')
      expect(second.destroyed).toBe(true)
      expect(first.destroyed).toBe(false)
      expect(element.getViewer()).toBe(first)
      expect(element.ready).toBe(true)
      expect(element.format).toBe('docx')
    })

    it('uses the options as passed even if the caller mutates them later', async () => {
      const element = createElement()
      FakeViewer.loadBehavior = 'hold'
      const options: OfficeViewerLoadOptions = { format: 'docx', mode: 'main' }
      const loading = element.load('/sample.docx', options)
      await viewerCount(1)

      options.format = 'xlsx'
      options.mode = 'worker'
      FakeViewer.last.finish()
      await loading

      expect(element.getViewer()).toBeInstanceOf(FakeDocxViewer)
      expect(element.format).toBe('docx')
      expect(element.mode).toBe('main')

      FakeViewer.loadBehavior = 'resolve'
      await element.reload()
      expect(FakeViewer.last.options.mode).toBe('main')
    })
  })

  describe('reload()', () => {
    it('rejects with an explicit error before any load()', async () => {
      const element = createElement()
      await expect(element.reload()).rejects.toThrow('Nothing to reload')
    })

    it('repeats the most recent request, including one that failed', async () => {
      const element = createElement()
      FakeViewer.loadBehavior = 'hold'
      const failing = element.load('/sample.docx', { format: 'docx', mode: 'main' })
      await viewerCount(1)
      FakeViewer.last.fail(new Error('boom'))
      await expect(failing).rejects.toThrow('boom')

      FakeViewer.loadBehavior = 'resolve'
      await element.reload()

      expect(FakeViewer.instances).toHaveLength(2)
      expect(FakeViewer.last.sources).toEqual(['/sample.docx'])
      expect(FakeViewer.last.options.mode).toBe('main')
      expect(element.ready).toBe(true)
      expect(element.error).toBeNull()
    })

    it('replays an ArrayBuffer source even though upstream detached the original', async () => {
      const element = createElement()
      const buffer = Uint8Array.from([1, 2, 3, 4]).buffer

      await element.load(buffer, { format: 'docx' })
      expect(buffer.byteLength).toBe(0)

      await element.reload()

      expect(FakeViewer.instances).toHaveLength(2)
      expect(bytesOf(FakeViewer.last.sources[0])).toEqual([1, 2, 3, 4])
    })

    it('re-reads a Blob source instead of the buffer upstream consumed', async () => {
      const element = createElement()

      await element.load(new Blob([Uint8Array.from([5, 6])]), { format: 'docx' })
      await element.reload()

      expect(bytesOf(FakeViewer.last.sources[0])).toEqual([5, 6])
    })

    it('replays a consumed ReadableStream source from the bytes read the first time', async () => {
      const element = createElement()

      await element.load(streamOf([1, 2], [3]), { format: 'docx' })
      await element.reload()

      expect(FakeViewer.instances).toHaveLength(2)
      expect(bytesOf(FakeViewer.last.sources[0])).toEqual([1, 2, 3])
      expect(element.ready).toBe(true)
    })
  })

  describe('superseding loads', () => {
    it('destroys the superseded viewer immediately and rejects its promise with AbortError', async () => {
      const element = createElement()
      const events = recordEvents(element)
      FakeViewer.loadBehavior = 'hold'

      const first = element.load('/first.docx', { format: 'docx' })
      await viewerCount(1)
      const firstViewer = FakeViewer.last

      const second = element.load('/second.docx', { format: 'docx' })
      expect(firstViewer.destroyed).toBe(true)
      await expect(first).rejects.toMatchObject({ name: 'AbortError' })

      await viewerCount(2)
      const secondViewer = FakeViewer.last
      secondViewer.finish()
      await second

      expect(element.getViewer()).toBe(secondViewer)
      expect(events).toEqual(['loadstart', 'loadstart', 'ready'])
      expect(element.error).toBeNull()
      expect(element.shadowRoot?.querySelectorAll('.fake-viewer')).toHaveLength(1)
    })

    it('settles a superseded promise even though upstream never settles its load()', async () => {
      const element = createElement()
      FakeViewer.loadBehavior = 'hold'

      const first = element.load('/first.docx', { format: 'docx' })
      await viewerCount(1)
      element.destroy()

      // The fake, like upstream, leaves the in-flight load() pending after destroy().
      await expect(first).rejects.toMatchObject({ name: 'AbortError', message: 'Destroyed while loading.' })
    })

    it('reports AbortError rather than the upstream error when a superseded load fails', async () => {
      const element = createElement()
      const errors = recordEvents(element, ['loaderror'])
      FakeViewer.loadBehavior = 'hold'

      const first = element.load('/first.docx', { format: 'docx' })
      await viewerCount(1)
      FakeViewer.last.fail(new Error('boom'))
      const second = element.load('/second.docx', { format: 'docx' })

      await expect(first).rejects.toMatchObject({ name: 'AbortError' })
      expect(errors).toEqual([])
      expect(element.error).toBeNull()

      await viewerCount(2)
      FakeViewer.last.finish()
      await second
      expect(element.ready).toBe(true)
    })

    it('cancels a pending ReadableStream source and never mounts its viewer', async () => {
      const element = createElement()
      let cancelled = false
      const stalled = new ReadableStream<Uint8Array>({
        pull() {
          return new Promise(() => {})
        },
        cancel() {
          cancelled = true
        }
      })

      const first = element.load(stalled, { format: 'docx' })
      await flush()
      const second = element.load('/second.docx', { format: 'docx' })

      await expect(first).rejects.toMatchObject({ name: 'AbortError' })
      expect(cancelled).toBe(true)
      await second
      expect(FakeViewer.instances).toHaveLength(1)
      expect(FakeViewer.last.sources).toEqual(['/second.docx'])
    })

    it('stops before reading the source when a loadstart listener destroys the element', async () => {
      const element = createElement()
      let pulled = false
      const stream = new ReadableStream<Uint8Array>(
        {
          pull() {
            pulled = true
            return new Promise(() => {})
          }
        },
        { highWaterMark: 0 }
      )
      element.addEventListener('loadstart', () => element.destroy(), { once: true })

      await expect(element.load(stream, { format: 'docx' })).rejects.toMatchObject({ name: 'AbortError' })

      expect(pulled).toBe(false)
      expect(stream.locked).toBe(false)
      expect(FakeViewer.instances).toHaveLength(0)
      expect(element.error).toBeNull()
    })
  })

  describe('destroy()', () => {
    it('is a safe no-op before any load and still emits destroy', () => {
      const element = createElement()
      const events = recordEvents(element)

      expect(() => {
        element.destroy()
        element.destroy()
      }).not.toThrow()

      expect(events).toEqual(['destroy', 'destroy'])
      expect(element.ready).toBe(false)
    })

    it('cancels an in-flight load, destroys its viewer, and resets state', async () => {
      const element = createElement()
      FakeViewer.loadBehavior = 'hold'
      const loading = element.load('/sample.docx', { format: 'docx' })
      await viewerCount(1)
      const viewer = FakeViewer.last
      const events = recordEvents(element)

      element.destroy()

      expect(viewer.destroyed).toBe(true)
      await expect(loading).rejects.toMatchObject({ name: 'AbortError' })
      expect(events).toEqual(['destroy'])
      expect(element.ready).toBe(false)
      expect(element.getViewer()).toBeNull()
      expect(element.format).toBeNull()
      expect(element.mode).toBeNull()
      expect(element.error).toBeNull()
      await expect(element.reload()).rejects.toThrow('Nothing to reload')
    })

    it('tears down a loaded viewer and allows loading again', async () => {
      const element = createElement()
      await element.load('/sample.docx', { format: 'docx' })
      const viewer = FakeViewer.last

      element.destroy()
      expect(viewer.destroyed).toBe(true)
      expect(element.shadowRoot?.querySelectorAll('.fake-viewer')).toHaveLength(0)

      await element.load('/sample.xlsx', { format: 'xlsx' })
      expect(element.getViewer()).toBeInstanceOf(FakeXlsxViewer)
    })
  })

  describe('attributes', () => {
    it('coalesces src, file-type, and mode set together into one load', async () => {
      const element = createElement()
      const events = recordEvents(element)

      await loadViaAttributes(element, '/sample.xlsx', 'xlsx', { mode: 'main' })

      expect(events).toEqual(['loadstart', 'ready'])
      expect(FakeViewer.instances).toHaveLength(1)
      expect(element.getViewer()).toBeInstanceOf(FakeXlsxViewer)
      expect(FakeViewer.last.sources).toEqual(['/sample.xlsx'])
      expect(element.mode).toBe('main')
    })

    it('loads from attributes set before the element is connected', async () => {
      const element = document.createElement('office-viewer') as OfficeViewerElement
      element.setAttribute('src', '/sample.pptx')
      element.setAttribute('file-type', 'PPTX')
      await flush()
      expect(FakeViewer.instances).toHaveLength(0)

      document.body.append(element)
      await settled(element)

      expect(element.getViewer()).toBeInstanceOf(FakePptxViewer)
      expect(element.format).toBe('pptx')
    })

    it('forwards wasm-url to the upstream viewer', async () => {
      const element = createElement()

      await loadViaAttributes(element, '/sample.docx', 'docx', { 'wasm-url': 'https://cdn.example/docx.wasm' })

      expect(FakeViewer.last.options.wasmUrl).toBe('https://cdn.example/docx.wasm')
    })

    it('emits loaderror when src is set without a valid file-type', async () => {
      const element = createElement()

      element.setAttribute('src', '/sample.docx')
      const event = (await once(element, 'loaderror')) as CustomEvent<OfficeViewerLoadErrorDetail>

      expect(event.detail.error.message).toContain('Missing format')
      expect(element.error).toBe(event.detail.error)
      expect(FakeViewer.instances).toHaveLength(0)
    })

    it('lets an explicit load() outrank attribute changes made earlier in the same task', async () => {
      const element = createElement()
      element.setAttribute('src', '/attr.docx')
      element.setAttribute('file-type', 'docx')
      const programmatic = element.load('/explicit.xlsx', { format: 'xlsx' })

      await programmatic
      await flush()

      expect(FakeViewer.instances).toHaveLength(1)
      expect(element.getViewer()).toBeInstanceOf(FakeXlsxViewer)
      expect(element.format).toBe('xlsx')
    })

    it('destroys an attribute-driven document when src is cleared', async () => {
      const element = createElement()
      await loadViaAttributes(element, '/sample.docx', 'docx')
      const viewer = FakeViewer.last
      const events = recordEvents(element)

      element.removeAttribute('src')
      await flush()

      expect(viewer.destroyed).toBe(true)
      expect(events).toEqual(['destroy'])
      expect(element.ready).toBe(false)
    })

    it('leaves a programmatic document alone when src is cleared', async () => {
      const element = createElement()
      await loadViaAttributes(element, '/sample.docx', 'docx')
      await element.load('/explicit.xlsx', { format: 'xlsx' })
      const viewer = FakeViewer.last
      const events = recordEvents(element)

      element.removeAttribute('src')
      await flush()

      expect(viewer.destroyed).toBe(false)
      expect(element.getViewer()).toBe(viewer)
      expect(events).toEqual([])
    })

    it('keeps the current document when the element is moved without attribute changes', async () => {
      const element = createElement()
      await loadViaAttributes(element, '/sample.docx', 'docx')
      const viewer = FakeViewer.last
      const events = recordEvents(element)

      await moveSynchronously(element)

      expect(viewer.destroyed).toBe(false)
      expect(element.getViewer()).toBe(viewer)
      expect(FakeViewer.instances).toHaveLength(1)
      expect(events).toEqual([])
    })

    it('keeps a programmatic load across a move even when src is set', async () => {
      const element = createElement()
      await loadViaAttributes(element, '/sample.docx', 'docx')
      await element.load('/other.xlsx', { format: 'xlsx' })
      const viewer = FakeViewer.last

      await moveSynchronously(element)

      expect(element.getViewer()).toBe(viewer)
      expect(element.format).toBe('xlsx')
    })

    it('does not retry a failed attribute load when the element is moved', async () => {
      const element = createElement()
      element.setAttribute('src', '/sample.docx')
      await once(element, 'loaderror')
      const events = recordEvents(element)

      await moveSynchronously(element)

      expect(events).toEqual([])
      expect(element.error?.message).toContain('Missing format')
    })

    it('reloads when attributes change while the element is detached', async () => {
      const element = createElement()
      await loadViaAttributes(element, '/sample.docx', 'docx')
      const previous = FakeViewer.last

      element.remove()
      element.setAttribute('src', '/sample.pptx')
      element.setAttribute('file-type', 'pptx')
      document.body.append(element)
      await settled(element)

      expect(previous.destroyed).toBe(true)
      expect(element.getViewer()).toBeInstanceOf(FakePptxViewer)
      expect(FakeViewer.instances).toHaveLength(2)
    })
  })

  describe('removal from the document', () => {
    it('destroys the viewer once the element has been removed', async () => {
      const element = createElement()
      await element.load('/sample.docx', { format: 'docx' })
      const viewer = FakeViewer.last
      const events = recordEvents(element)

      element.remove()
      expect(viewer.destroyed).toBe(false)
      await flush()

      expect(viewer.destroyed).toBe(true)
      expect(events).toEqual(['destroy'])
      expect(element.ready).toBe(false)
      expect(element.getViewer()).toBeNull()
    })

    it('cancels an in-flight load once the element has been removed', async () => {
      const element = createElement()
      FakeViewer.loadBehavior = 'hold'
      const loading = element.load('/sample.docx', { format: 'docx' })
      await viewerCount(1)

      element.remove()

      await expect(loading).rejects.toMatchObject({ name: 'AbortError' })
      expect(FakeViewer.last.destroyed).toBe(true)
    })

    it('restores a programmatic document when the element is added again', async () => {
      const element = createElement()
      await element.load(new Blob([Uint8Array.from([7, 8])]), { format: 'xlsx', mode: 'main' })
      const events = recordEvents(element)

      element.remove()
      await flush()
      document.body.append(element)
      await settled(element)

      expect(events).toEqual(['destroy', 'loadstart', 'ready'])
      expect(FakeViewer.instances).toHaveLength(2)
      expect(element.getViewer()).toBeInstanceOf(FakeXlsxViewer)
      expect(bytesOf(FakeViewer.last.sources[0])).toEqual([7, 8])
      expect(element.mode).toBe('main')
    })

    it('restores an attribute-driven document when the element is added again', async () => {
      const element = createElement()
      await loadViaAttributes(element, '/sample.docx', 'docx')

      element.remove()
      await flush()
      expect(element.ready).toBe(false)

      document.body.append(element)
      await settled(element)

      expect(element.ready).toBe(true)
      expect(FakeViewer.instances).toHaveLength(2)
      expect(FakeViewer.last.sources).toEqual(['/sample.docx'])
    })

    it('prefers attributes changed while detached over restoring the previous document', async () => {
      const element = createElement()
      await element.load('/explicit.xlsx', { format: 'xlsx' })

      element.remove()
      await flush()
      element.setAttribute('src', '/sample.pptx')
      element.setAttribute('file-type', 'pptx')
      document.body.append(element)
      await settled(element)

      expect(element.getViewer()).toBeInstanceOf(FakePptxViewer)
      expect(FakeViewer.instances).toHaveLength(2)
    })

    it('does not restore after an explicit destroy()', async () => {
      const element = createElement()
      await element.load('/sample.docx', { format: 'docx' })
      const events = recordEvents(element)

      element.destroy()
      element.remove()
      await flush()
      document.body.append(element)
      await flush()

      expect(events).toEqual(['destroy'])
      expect(element.ready).toBe(false)
      expect(FakeViewer.instances).toHaveLength(1)
    })

    it('does nothing when an idle element is removed', async () => {
      const element = createElement()
      const events = recordEvents(element)

      element.remove()
      await flush()

      expect(events).toEqual([])
    })

    it('leaves a failed load untouched when the element is removed and re-added', async () => {
      const element = createElement()
      await expect(element.load(42 as never, { format: 'docx' })).rejects.toThrow()
      const events = recordEvents(element)

      element.remove()
      await flush()
      document.body.append(element)
      await flush()

      expect(events).toEqual([])
      expect(element.error?.message).toContain('Unsupported source type')
      expect(FakeViewer.instances).toHaveLength(0)
    })
  })
})
