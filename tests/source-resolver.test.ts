import { describe, expect, it } from 'vitest'

import { detectOfficeFormat, resolveOfficeSource } from '../src'

describe('source-resolver', () => {
  describe('detectOfficeFormat', () => {
    it('prefers the explicit format override', () => {
      expect(detectOfficeFormat('report.docx', 'xlsx')).toBe('xlsx')
      expect(detectOfficeFormat(undefined, 'pptx')).toBe('pptx')
    })

    it('detects formats from filename extensions', () => {
      expect(detectOfficeFormat('report.docx')).toBe('docx')
      expect(detectOfficeFormat('book.xlsx')).toBe('xlsx')
      expect(detectOfficeFormat('deck.pptx')).toBe('pptx')
    })

    it('is case-insensitive for extensions', () => {
      expect(detectOfficeFormat('Report.DOCX')).toBe('docx')
      expect(detectOfficeFormat('Book.XLSX')).toBe('xlsx')
      expect(detectOfficeFormat('Deck.PPTX')).toBe('pptx')
    })

    it('throws for unsupported or missing extensions', () => {
      expect(() => detectOfficeFormat('archive.zip')).toThrow('Unable to determine Office format')
      expect(() => detectOfficeFormat(undefined)).toThrow('Unable to determine Office format')
    })
  })

  describe('resolveOfficeSource', () => {
    it('treats strings as URLs', async () => {
      const resolved = await resolveOfficeSource('http://localhost/fixtures/report.docx', undefined, undefined)
      expect(resolved.sourceKind).toBe('url')
      expect(resolved.format).toBe('docx')
      expect(resolved.fileName).toBe('report.docx')
    })

    it('converts URL objects to strings', async () => {
      const resolved = await resolveOfficeSource(new URL('/fixtures/book.xlsx', 'http://localhost'), undefined, undefined)
      expect(resolved.sourceKind).toBe('url')
      expect(resolved.format).toBe('xlsx')
      expect(resolved.source).toBe('http://localhost/fixtures/book.xlsx')
    })

    it('uses explicit format over filename inference', async () => {
      const resolved = await resolveOfficeSource('/fixtures/unknown', undefined, 'pptx')
      expect(resolved.format).toBe('pptx')
    })

    it('uses file-name when URL has no extension', async () => {
      const resolved = await resolveOfficeSource('http://localhost/fixtures/unknown', 'report.docx', undefined)
      expect(resolved.format).toBe('docx')
      expect(resolved.fileName).toBe('report.docx')
    })

    it('copies File contents to an ArrayBuffer', async () => {
      const file = new File([new Uint8Array([1, 2, 3])], 'sample.docx')
      const resolved = await resolveOfficeSource(file, undefined, undefined)

      expect(resolved.sourceKind).toBe('file')
      expect(resolved.format).toBe('docx')
      expect(resolved.source).toBeInstanceOf(ArrayBuffer)
      expect(new Uint8Array(resolved.source as ArrayBuffer)).toEqual(new Uint8Array([1, 2, 3]))
    })

    it('copies Blob contents to an ArrayBuffer', async () => {
      const blob = new Blob([new Uint8Array([4, 5, 6])])
      const resolved = await resolveOfficeSource(blob, 'book.xlsx', undefined)

      expect(resolved.sourceKind).toBe('blob')
      expect(resolved.format).toBe('xlsx')
      expect(resolved.source).toBeInstanceOf(ArrayBuffer)
      expect(new Uint8Array(resolved.source as ArrayBuffer)).toEqual(new Uint8Array([4, 5, 6]))
    })

    it('copies ArrayBuffer contents', async () => {
      const buffer = new Uint8Array([7, 8, 9]).buffer
      const resolved = await resolveOfficeSource(buffer, 'deck.pptx', undefined)

      expect(resolved.sourceKind).toBe('array-buffer')
      expect(resolved.format).toBe('pptx')
      expect(resolved.source).toBeInstanceOf(ArrayBuffer)
      expect(new Uint8Array(resolved.source as ArrayBuffer)).toEqual(new Uint8Array([7, 8, 9]))

      // Original buffer must remain untouched.
      new Uint8Array(buffer)[0] = 99
      expect(new Uint8Array(resolved.source as ArrayBuffer)[0]).toBe(7)
    })

    it('copies Uint8Array contents', async () => {
      const view = new Uint8Array([10, 11, 12])
      const resolved = await resolveOfficeSource(view, 'report.docx', undefined)

      expect(resolved.sourceKind).toBe('uint8array')
      expect(resolved.format).toBe('docx')
      expect(resolved.source).toBeInstanceOf(ArrayBuffer)
      expect(new Uint8Array(resolved.source as ArrayBuffer)).toEqual(new Uint8Array([10, 11, 12]))

      view[0] = 99
      expect(new Uint8Array(resolved.source as ArrayBuffer)[0]).toBe(10)
    })

    it('throws for binary sources without format or filename', async () => {
      await expect(resolveOfficeSource(new Uint8Array([1]), undefined, undefined)).rejects.toThrow(
        'Unable to determine Office format'
      )
    })
  })
})
