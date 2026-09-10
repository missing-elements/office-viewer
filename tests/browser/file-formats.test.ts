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
