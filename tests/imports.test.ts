import { describe, expect, it } from 'vitest'

import { defineOfficeViewerElement, detectOfficeFormat, normalizeOfficeSource, OfficeViewerElement, OoxmlIntegrationSpike } from '../src'

describe('OOXML spike imports', () => {
  it('imports the three upstream format entry points', async () => {
    const [docx, xlsx, pptx] = await Promise.all([
      import('@silurus/ooxml/docx'),
      import('@silurus/ooxml/xlsx'),
      import('@silurus/ooxml/pptx')
    ])

    expect(docx.DocxScrollViewer).toBeTypeOf('function')
    expect(xlsx.XlsxViewer).toBeTypeOf('function')
    expect(pptx.PptxScrollViewer).toBeTypeOf('function')
  })

  it('detects formats conservatively', () => {
    expect(detectOfficeFormat({ fileName: 'report.docx' })).toBe('docx')
    expect(detectOfficeFormat({ fileName: 'book.xlsx' })).toBe('xlsx')
    expect(detectOfficeFormat({ fileName: 'deck.pptx' })).toBe('pptx')
  })

  it('normalizes Uint8Array and preserves the spike API surface', async () => {
    const normalized = await normalizeOfficeSource(new Uint8Array([1, 2, 3]), {
      format: 'docx',
      fileName: 'sample.docx'
    })

    expect(normalized.source).toBeInstanceOf(ArrayBuffer)
    expect(normalized.sourceKind).toBe('uint8array')
    expect(normalized.format).toBe('docx')
    expect(OoxmlIntegrationSpike).toBeTypeOf('function')
  })

  it('retains independent copies of caller-owned binary inputs', async () => {
    const arrayBuffer = new Uint8Array([1, 2, 3]).buffer
    const arrayBufferLoad = await normalizeOfficeSource(arrayBuffer, { format: 'docx' })
    new Uint8Array(arrayBuffer)[0] = 9

    const sourceView = new Uint8Array([4, 5, 6])
    const uint8ArrayLoad = await normalizeOfficeSource(sourceView, { format: 'xlsx' })
    sourceView[0] = 9

    expect(arrayBufferLoad.source).toBeInstanceOf(ArrayBuffer)
    expect(uint8ArrayLoad.source).toBeInstanceOf(ArrayBuffer)
    expect(new Uint8Array(arrayBufferLoad.source as ArrayBuffer)).toEqual(new Uint8Array([1, 2, 3]))
    expect(new Uint8Array(uint8ArrayLoad.source as ArrayBuffer)).toEqual(new Uint8Array([4, 5, 6]))
  })

  it('exports the initial custom-element wrapper surface', () => {
    expect(OfficeViewerElement).toBeTypeOf('function')
    expect(defineOfficeViewerElement).toBeTypeOf('function')
  })
})
