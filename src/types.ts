export type OfficeSource = string | ArrayBuffer | Blob | ReadableStream<Uint8Array>
export type OfficeFormat = 'docx' | 'xlsx' | 'pptx'
export type OfficeViewerMode = 'worker' | 'main'
export type OfficeViewerStatus = 'idle' | 'loading' | 'ready' | 'error'

export interface OfficeViewerLoadOptions {
  format: OfficeFormat
  mode?: OfficeViewerMode
  wasmUrl?: string | URL
}
