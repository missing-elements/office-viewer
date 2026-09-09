/**
 * Public source types accepted by `<office-viewer>`.
 *
 * String values are treated as URLs. Caller-owned binary buffers (ArrayBuffer,
 * Uint8Array) are copied before transfer to workers.
 */
export type OfficeSource =
  | string
  | URL
  | File
  | Blob
  | ArrayBuffer
  | Uint8Array

export type OfficeSourceKind = 'url' | 'file' | 'blob' | 'array-buffer' | 'uint8array'

export type OfficeFormat = 'docx' | 'xlsx' | 'pptx'
export type OfficeViewerMode = 'worker' | 'main'

export interface OfficeViewerLoadOptions {
  /** Explicit format override. */
  format?: OfficeFormat
  /** Metadata filename for sources without an intrinsic name. */
  fileName?: string
  /** Preferred rendering mode. Defaults to `'worker'`. */
  mode?: OfficeViewerMode
  /** Custom WASM asset URL, passed through to upstream engine loaders. */
  wasmUrl?: string | URL
  /** Abort signal that cancels the load. */
  signal?: AbortSignal
}

export interface OfficeViewerReadyEventDetail {
  format: OfficeFormat
  requestedMode: OfficeViewerMode
  effectiveMode: OfficeViewerMode
}

export interface OfficeViewerLoadErrorEventDetail {
  error: Error
}
