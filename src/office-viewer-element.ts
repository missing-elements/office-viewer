import { DocxScrollViewer } from '@silurus/ooxml/docx'
import { PptxScrollViewer } from '@silurus/ooxml/pptx'
import { XlsxViewer } from '@silurus/ooxml/xlsx'
import type {
  OfficeFormat,
  OfficeSource,
  OfficeViewerLoadOptions,
  OfficeViewerMode,
  OfficeViewerStatus
} from './types'

export const OFFICE_VIEWER_TAG_NAME = 'office-viewer'

export type OfficeViewer = DocxScrollViewer | XlsxViewer | PptxScrollViewer

const RELOAD_ATTRIBUTE_NAMES = new Set(['src', 'file-type', 'mode', 'wasm-url'])

if (typeof globalThis.HTMLElement === 'undefined') {
  ;(globalThis as { HTMLElement: typeof HTMLElement }).HTMLElement = class {} as typeof HTMLElement
}

export class OfficeViewerElement extends HTMLElement {
  static get observedAttributes(): string[] {
    return ['src', 'file-type', 'mode', 'wasm-url']
  }

  static readonly tagName = OFFICE_VIEWER_TAG_NAME
  static readonly shadowRootMode: ShadowRootMode = 'open'

  private viewer: OfficeViewer | null = null
  private status: OfficeViewerStatus = 'idle'
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

  get ready(): boolean {
    return this.status === 'ready'
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
      viewer = createViewer(options, this.ensureContainer())
      await viewer.load(source)
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
    this.status = value
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
      style.textContent = `
        :host { display: block; }
        #viewer { width: 100%; height: 100%; overflow: auto; }
      `
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

function createViewer(
  options: OfficeViewerLoadOptions,
  container: HTMLElement
): OfficeViewer {
  const loadOptions = {
    mode: options.mode ?? 'worker',
    wasmUrl: normalizeWasmUrl(options.wasmUrl)
  }

  switch (options.format) {
    case 'docx':
      return new DocxScrollViewer(container, loadOptions)
    case 'xlsx':
      return new XlsxViewer(container, loadOptions)
    case 'pptx':
      return new PptxScrollViewer(container, loadOptions)
    default:
      throw new Error(`Unsupported format: ${options.format}`)
  }
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

function createAbortError(reason: string): DOMException {
  return new DOMException(reason, 'AbortError')
}

function toError(reason: unknown): Error {
  if (reason instanceof Error) {
    return reason
  }
  return new Error(String(reason))
}
