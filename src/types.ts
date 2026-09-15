import type { DocxScrollViewer } from '@silurus/ooxml/docx'
import type { PptxScrollViewer } from '@silurus/ooxml/pptx'
import type { XlsxViewer } from '@silurus/ooxml/xlsx'

export const OFFICE_FORMATS = ['docx', 'xlsx', 'pptx'] as const

export type OfficeFormat = (typeof OFFICE_FORMATS)[number]
export type OfficeSource = string | ArrayBuffer | Blob | ReadableStream<Uint8Array>
export type OfficeViewerMode = 'worker' | 'main'

/** Upstream viewer instance returned by `getViewer()`. */
export type OfficeViewer = DocxScrollViewer | XlsxViewer | PptxScrollViewer

export interface OfficeViewerLoadOptions {
  format: OfficeFormat
  mode?: OfficeViewerMode
  wasmUrl?: string | URL
}

/** `detail` of the `loaderror` event. */
export interface OfficeViewerLoadErrorDetail {
  error: Error
}
