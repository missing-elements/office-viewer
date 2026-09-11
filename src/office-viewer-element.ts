import type {
  OfficeFormat,
  OfficeSource,
  OfficeViewerLoadOptions,
  OfficeViewerMode,
} from './types'

export const OFFICE_VIEWER_TAG_NAME = 'office-viewer'

const RELOAD_ATTRIBUTE_NAMES = new Set(['src', 'file-type', 'mode'])

import { type DocxScrollViewer as DocxScrollViewerType } from '@silurus/ooxml/docx'
import { type PptxScrollViewer as PptxScrollViewerType } from '@silurus/ooxml/pptx'
import { type XlsxViewer as XlsxViewerType } from '@silurus/ooxml/xlsx'

export type OfficeViewer = DocxScrollViewerType | XlsxViewerType | PptxScrollViewerType

export class OfficeViewerElement extends HTMLElement {
  static get observedAttributes(): string[] {
    return ['src', 'file-type', 'mode']
  }

  static readonly tagName = OFFICE_VIEWER_TAG_NAME
  static readonly shadowRootMode: ShadowRootMode = 'open'

  private viewer: OfficeViewer | null = null
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
    this.dispatchEvent(new CustomEvent('loadstart'))

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
      this.dispatchEvent(new CustomEvent('loaderror', { detail: { error } }))
      throw error
    }

    if (this.abortController !== controller) {
      // Superseded or destroyed while loading; leave the current state untouched.
      viewer.destroy()
      throw createAbortError('Load aborted.')
    }

    this.viewer = viewer
    previous?.destroy()
    this.dispatchEvent(new CustomEvent('ready'))
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
    this.dispatchEvent(new CustomEvent('destroy'))
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
        mode: parseMode(this.getAttribute('mode'))
      })
    } catch {
      // load() emits loaderror.
    }
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
    mode: options.mode ?? 'worker'
  }

  if (options.format === 'docx') {
    const { DocxScrollViewer } = await import('@silurus/ooxml/docx')
    return new DocxScrollViewer(container, loadOptions)
  }
  if (options.format === 'pptx') {
    const { PptxScrollViewer } = await import('@silurus/ooxml/pptx')
    return new PptxScrollViewer(container, loadOptions)
  }
  if (options.format === 'xlsx') {
    const { XlsxViewer } = await import('@silurus/ooxml/xlsx')
    return new XlsxViewer(container, loadOptions)
  }

  throw new Error(`Unsupported format: ${options.format}`)
}

function parseFormat(value: string | null): OfficeFormat | undefined {
  const normalized = value?.trim().toLowerCase()
  if (['docx', 'xlsx', 'pptx'].includes(normalized ?? '')) {
    return normalized as OfficeFormat
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
