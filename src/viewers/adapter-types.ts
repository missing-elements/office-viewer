import type { OfficeFormat, OfficeViewerMode } from '../types'

export interface AdapterSource {
  readonly source: string | ArrayBuffer
  readonly format: OfficeFormat
  readonly fileName?: string
}

export interface AdapterLoadOptions {
  mode?: OfficeViewerMode
  wasmUrl?: string | URL
  container?: HTMLElement
}

export interface ViewerAdapter {
  readonly format: OfficeFormat
  readonly viewer: unknown | null
  readonly document: unknown | null
  readonly engine: unknown | null

  load(source: AdapterSource, options: AdapterLoadOptions): Promise<void>
  destroy(): void
}
