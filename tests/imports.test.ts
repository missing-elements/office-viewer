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

  it('exports the initial custom-element wrapper surface', () => {
    expect(OfficeViewerElement).toBeTypeOf('function')
    expect(defineOfficeViewerElement).toBeTypeOf('function')
  })
})
