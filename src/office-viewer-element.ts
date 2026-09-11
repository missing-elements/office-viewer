import type {
  OfficeFormat,
  OfficeSource,
  OfficeViewerLoadOptions,
  OfficeViewerMode,
  OfficeViewerStatus
} from './types'

export const OFFICE_VIEWER_TAG_NAME = 'office-viewer'

const RELOAD_ATTRIBUTE_NAMES = new Set(['src', 'file-type', 'mode', 'wasm-url'])

// Allow module evaluation in non-DOM runtimes (SSR/tests); defineOfficeViewerElement()
// still throws a clear error when customElements is unavailable.
if (typeof globalThis.HTMLElement === 'undefined') {
  ;(globalThis as { HTMLElement: typeof HTMLElement }).HTMLElement = class {} as typeof HTMLElement
}

import type { DocxScrollViewer } from '@silurus/ooxml/docx'
import type { XlsxViewer } from '@silurus/ooxml/xlsx'
import type { PptxScrollViewer } from '@silurus/ooxml/pptx'

export type OfficeViewer = DocxScrollViewer | XlsxViewer | PptxScrollViewer

export class OfficeViewerElement extends HTMLElement {
  static get observedAttributes(): string[] {
    return ['src', 'file-type', 'mode', 'wasm-url']
  }

  static readonly tagName = OFFICE_VIEWER_TAG_NAME
  static readonly shadowRootMode: ShadowRootMode = 'open'

  private viewer: OfficeViewer | null = null
  private _status: OfficeViewerStatus = 'idle'
  private lastError: Error | null = null
  private retainedSource: OfficeSource | null = null
  private retainedOptions: OfficeViewerLoadOptions | null = null
  private abortController: AbortController | null = null

  get src(): string | null {
    return this.getAttribute('src')
  }

  set src(value: string | null) {
    if (!value) {
      this.removeAttribute('src')
      return
    }
    this.setAttribute('src', value)
  }

  get status(): OfficeViewerStatus {
    return this._status
  }

  get ready(): boolean {
    return this._status === 'ready'
  }

  get error(): Error | null {
    return this.lastError
  }

  get format(): OfficeFormat | null {
    return this.retainedOptions?.format ?? null
  }

  get mode(): OfficeViewerMode | null {
    const mode = (this.viewer as { mode?: OfficeViewerMode } | null)?.mode
    if (mode === 'worker' || mode === 'main') {
      return mode
    }
    return this.retainedOptions?.mode ?? null
  }

  getViewer(): OfficeViewer | null {
    return this.viewer
  }

  connectedCallback(): void {
    this.ensureContainer()
    if (this.src?.trim()) {
      void this.loadFromAttributes()
    }
  }

  attributeChangedCallback(name: string, oldValue: string | null, newValue: string | null): void {
    if (oldValue === newValue || !this.isConnected || !RELOAD_ATTRIBUTE_NAMES.has(name)) {
      return
    }

    if (name === 'src' && !newValue?.trim()) {
      this.destroy()
      return
    }

    if (this.src?.trim()) {
      void this.loadFromAttributes()
    }
  }

  async load(source: OfficeSource, options: OfficeViewerLoadOptions): Promise<void> {
    this.cancelActiveLoad()
    const controller = new AbortController()
    this.abortController = controller

    this.retainedSource = source
    this.retainedOptions = options
    this.lastError = null
    this.setStatus('loading')
    this.emit('loadstart')

    const previous = this.viewer
    let viewer: OfficeViewer | null = null

    try {
      // Await the promise returned by createViewer (which contains a dynamic import)
      viewer = await createViewer(options, this.ensureContainer())
      // Upstream only accepts a URL string or ArrayBuffer; normalize Blob/File and
      // ReadableStream sources to an ArrayBuffer so they load the same way.
      await viewer.load(await resolveSource(source))
    } catch (reason) {
      viewer?.destroy()
      const error = toError(reason)
      if (this.abortController !== controller) {
        // A newer load or destroy() owns the element state now; leave it untouched.
        throw error
      }
      this.lastError = error
      this.setStatus('error')
      this.emit('loaderror', { error })
      throw error
    }

    if (this.abortController !== controller) {
      // Superseded or destroyed while loading; leave the current state untouched.
      viewer.destroy()
      throw createAbortError('Load aborted.')
    }

    this.viewer = viewer
    previous?.destroy()
    this.setStatus('ready')
    this.emit('ready')
  }

  async reload(): Promise<void> {
    if (!this.retainedSource || !this.retainedOptions) {
      throw new Error('Nothing has been loaded yet.')
    }
    await this.load(this.retainedSource, this.retainedOptions)
  }

  destroy(): void {
    this.cancelActiveLoad()
    this.viewer?.destroy()
    this.viewer = null
    this.retainedSource = null
    this.retainedOptions = null
    this.lastError = null
    this.setStatus('idle')
    this.emit('destroy')
  }

  private setStatus(value: OfficeViewerStatus): void {
    this._status = value
  }

  private cancelActiveLoad(): void {
    if (this.abortController) {
      this.abortController.abort()
      this.abortController = null
    }
  }

  private ensureContainer(): HTMLElement {
    if (!this.shadowRoot && typeof this.attachShadow === 'function') {
      const root = this.attachShadow({ mode: OfficeViewerElement.shadowRootMode })
      const style = document.createElement('style')
      style.textContent = `:host{display:block}#viewer{width:100%;height:100%;overflow:auto}`
      const container = document.createElement('div')
      container.id = 'viewer'
      root.append(style, container)
      return container
    }

    const existing = this.shadowRoot?.getElementById('viewer')
    if (existing) {
      return existing
    }

    throw new Error('Unable to create Shadow DOM render container for office-viewer.')
  }

  private async loadFromAttributes(): Promise<void> {
    const source = this.src?.trim()
    const format = parseFormat(this.getAttribute('file-type'))
    if (!source || !format) {
      return
    }

    try {
      await this.load(source, {
        format,
        mode: parseMode(this.getAttribute('mode')),
        wasmUrl: normalizeWasmUrl(this.getAttribute('wasm-url') ?? undefined)
      })
    } catch {
      // load() emits loaderror.
    }
  }

  private emit<T>(name: string, detail?: T): void {
    if (typeof CustomEvent === 'function') {
      this.dispatchEvent(new CustomEvent(name, { detail }))
      return
    }
    this.dispatchEvent(new Event(name))
  }
}

export function defineOfficeViewerElement(tagName = OFFICE_VIEWER_TAG_NAME): typeof OfficeViewerElement {
  const existing = typeof customElements === 'undefined' ? undefined : customElements.get(tagName)
  if (existing) {
    return existing as typeof OfficeViewerElement
  }

  if (typeof customElements === 'undefined') {
    throw new Error('Custom elements are not available in this runtime.')
  }

  customElements.define(tagName, OfficeViewerElement)
  return OfficeViewerElement
}

async function createViewer(
  options: OfficeViewerLoadOptions,
  container: HTMLElement
): Promise<OfficeViewer> {
  const loadOptions = {
    mode: options.mode ?? 'worker',
    wasmUrl: normalizeWasmUrl(options.wasmUrl)
  }

  // Ensure format is normalized for case-insensitive matching
  const format = (options.format ?? '').trim().toLowerCase()

  const viewerPromise = (async () => {
    switch (format) {
      case 'docx': {
        const { DocxScrollViewer } = await import('@silurus/ooxml/docx')
        return new DocxScrollViewer(container, loadOptions)
      }
      case 'xlsx': {
        const { XlsxViewer } = await import('@silurus/ooxml/xlsx')
        return new XlsxViewer(container, loadOptions)
      }
      case 'pptx': {
        const { PptxScrollViewer } = await import('@silurus/ooxml/pptx')
        return new PptxScrollViewer(container, loadOptions)
      }
      default:
        throw new Error(`Unsupported format: ${options.format}`)
    }
  })()

  return viewerPromise
}

function parseFormat(value: string | null): OfficeFormat | undefined {
  const normalized = value?.trim().toLowerCase()
  if (normalized === 'docx' || normalized === 'xlsx' || normalized === 'pptx') {
    return normalized
  }
  return undefined
}

function parseMode(value: string | null): OfficeViewerMode | undefined {
  const normalized = value?.trim().toLowerCase()
  if (normalized === 'worker' || normalized === 'main') {
    return normalized
  }
  return undefined
}

function normalizeWasmUrl(wasmUrl: string | URL | undefined): string | undefined {
  if (wasmUrl instanceof URL) {
    return wasmUrl.toString()
  }
  return wasmUrl
}

// Converts any supported OfficeSource into the string | ArrayBuffer form upstream
// accepts. Strings pass through unchanged; Blob/File and ReadableStream are read
// fully into an ArrayBuffer.
async function resolveSource(source: OfficeSource): Promise<string | ArrayBuffer> {
  if (typeof source === 'string' || source instanceof ArrayBuffer) {
    return source
  }
  if (typeof Blob !== 'undefined' && source instanceof Blob) {
    return source.arrayBuffer()
  }
  if (typeof ReadableStream !== 'undefined' && source instanceof ReadableStream) {
    return readStreamToArrayBuffer(source)
  }
  throw new Error('Unsupported source type. Expected a URL string, ArrayBuffer, Blob, File, or ReadableStream<Uint8Array>.')
}

async function readStreamToArrayBuffer(stream: ReadableStream<Uint8Array>): Promise<ArrayBuffer> {
  const reader = stream.getReader()
  const chunks: Uint8Array[] = []
  let total = 0
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) {
        break
      }
      if (value) {
        chunks.push(value)
        total += value.byteLength
      }
    }
  } finally {
    reader.releaseLock()
  }

  const buffer = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    buffer.set(chunk, offset)
    offset += chunk.byteLength
  }
  return buffer.buffer
}

function createAbortError(reason: string): DOMException {
  return new DOMException(reason, 'AbortError')
}

function toError(reason: unknown): Error {
  if (reason instanceof Error) {
    return reason
  }
  return new Error(String(reason))
}
