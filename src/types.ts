export type OfficeSource = string | ArrayBuffer | Blob | ReadableStream<Uint8Array>
export type OfficeFormat = 'docx' | 'xlsx' | 'pptx'
export type OfficeViewerMode = 'worker' | 'main'

export type DocxScrollViewer = typeof import('@silurus/ooxml/docx').DocxScrollViewer
export type PptxScrollViewer = typeof import('@silurus/ooxml/pptx').PptxScrollViewer
export type XlsxViewer = typeof import('@silurus/ooxml/xlsx').XlsxViewer

export type OfficeViewer = DocxScrollViewer | PptxScrollViewer | XlsxViewer

export interface OfficeViewerLoadOptions {
  format: OfficeFormat
  mode?: OfficeViewerMode
}
