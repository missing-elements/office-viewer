import { fileTypeFromBuffer } from 'file-type'
import { describe, expect, it } from 'vitest'

import { defineOfficeViewerElement, OfficeViewerElement } from '../../src'
import type { OfficeFormat, OfficeViewerMode } from '../../src/types'

// Fixture documents are served by the Vitest browser dev server from /public.
const FIXTURES: Record<OfficeFormat, string> = {
  docx: '/fixtures/sample.docx',
  xlsx: '/fixtures/sample.xlsx',
  pptx: '/fixtures/sample.pptx'
}

// Exercise both render modes so worker-mode WASM loading is covered, not just main.
const MODES: OfficeViewerMode[] = ['main', 'worker']

async function loadFixture(
  format: OfficeFormat,
  mode: OfficeViewerMode
): Promise<OfficeViewerElement> {
  defineOfficeViewerElement()
  const element = document.createElement('office-viewer') as OfficeViewerElement
  document.body.append(element)
  try {
    await element.load(FIXTURES[format], { format, mode })
    return element
  } catch (error) {
    element.remove()
    throw error
  }
}

describe('office-viewer file formats', () => {
  it('detects each fixture format from its contents', async () => {
    for (const [format, fixture] of Object.entries(FIXTURES)) {
      const buffer = await (await fetch(fixture)).arrayBuffer()
      const detected = await fileTypeFromBuffer(buffer)

      expect(detected?.ext).toBe(format)
    }
  })

  for (const format of Object.keys(FIXTURES) as OfficeFormat[]) {
    for (const mode of MODES) {
      it(`loads a ${format} document to ready in ${mode} mode`, async () => {
        const element = await loadFixture(format, mode)
        try {
          expect(element.error).toBeNull()
          expect(element.ready).toBe(true)
          expect(element.format).toBe(format)
          expect(element.mode).toBe(mode)
          expect(element.getViewer()).not.toBeNull()
          expect(element.shadowRoot?.getElementById('viewer')).not.toBeNull()
        } finally {
          element.destroy()
          element.remove()
        }
      }, 30_000)
    }
  }

  it('emits loadstart and ready events on success', async () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)

    const events: string[] = []
    element.addEventListener('loadstart', () => events.push('loadstart'))
    element.addEventListener('ready', () => events.push('ready'))

    try {
      await element.load(FIXTURES.docx, { format: 'docx' })
      expect(events).toEqual(['loadstart', 'ready'])
      // mode was omitted, so the element must report the effective default.
      expect(element.mode).toBe('worker')
    } finally {
      element.destroy()
      element.remove()
    }
  }, 30_000)

  it('surfaces an error for a malformed document', async () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)

    try {
      await expect(element.load('/fixtures/malformed.docx', { format: 'docx' })).rejects.toThrow()
      expect(element.ready).toBe(false)
      expect(element.error).not.toBeNull()
      expect(element.format).toBeNull()
    } finally {
      element.destroy()
      element.remove()
    }
  }, 30_000)

  // Regression: the demo "Load file" button passes a File/Blob, which previously
  // failed because only string URLs were accepted.
  it('loads a document from a File source', async () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)

    try {
      const response = await fetch(FIXTURES.docx)
      const blob = await response.blob()
      const file = new File([blob], 'sample.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      })

      await element.load(file, { format: 'docx' })
      expect(element.error).toBeNull()
      expect(element.ready).toBe(true)
      expect(element.getViewer()).not.toBeNull()
    } finally {
      element.destroy()
      element.remove()
    }
  }, 30_000)

  it('rejects a superseded load with AbortError and commits only the newest one', async () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)

    const errors: string[] = []
    element.addEventListener('loaderror', () => errors.push('loaderror'))

    try {
      const first = element.load(FIXTURES.docx, { format: 'docx' })
      const second = element.load(FIXTURES.xlsx, { format: 'xlsx' })

      await expect(first).rejects.toMatchObject({ name: 'AbortError' })
      await second

      expect(errors).toEqual([])
      expect(element.error).toBeNull()
      expect(element.format).toBe('xlsx')
      expect(element.ready).toBe(true)
    } finally {
      element.destroy()
      element.remove()
    }
  }, 30_000)

  // Regression: upstream may take ownership of the bytes it is given (docx and pptx detach
  // the buffer after load), so reload() must replay from a re-readable copy.
  it('reloads a document that was loaded from a File', async () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)

    try {
      const blob = await (await fetch(FIXTURES.docx)).blob()
      const file = new File([blob], 'sample.docx')
      await element.load(file, { format: 'docx' })
      const first = element.getViewer()

      await element.reload()

      expect(element.error).toBeNull()
      expect(element.ready).toBe(true)
      expect(element.getViewer()).not.toBe(first)
    } finally {
      element.destroy()
      element.remove()
    }
  }, 30_000)

  it('reloads a document that was loaded from an ArrayBuffer', async () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)

    try {
      const buffer = await (await fetch(FIXTURES.docx)).arrayBuffer()
      await element.load(buffer, { format: 'docx' })

      await element.reload()

      expect(element.error).toBeNull()
      expect(element.ready).toBe(true)
    } finally {
      element.destroy()
      element.remove()
    }
  }, 30_000)

  it('settles a load cancelled by destroy() without waiting for the parser', async () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)

    try {
      const loading = element.load(FIXTURES.pptx, { format: 'pptx' })
      element.destroy()
      const started = performance.now()

      await expect(loading).rejects.toMatchObject({ name: 'AbortError' })

      expect(performance.now() - started).toBeLessThan(200)
      expect(element.ready).toBe(false)
    } finally {
      element.destroy()
      element.remove()
    }
  }, 30_000)

  it('loads a document from an ArrayBuffer source', async () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)

    try {
      const buffer = await (await fetch(FIXTURES.xlsx)).arrayBuffer()
      await element.load(buffer, { format: 'xlsx' })
      expect(element.error).toBeNull()
      expect(element.ready).toBe(true)
    } finally {
      element.destroy()
      element.remove()
    }
  }, 30_000)
})
