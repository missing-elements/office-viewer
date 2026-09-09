import type {
  OfficeFormat,
  OfficeSource,
  OfficeViewerLoadErrorEventDetail,
  OfficeViewerLoadOptions,
  OfficeViewerMode,
  OfficeViewerReadyEventDetail
} from './types'
import { resolveOfficeSource } from './source-resolver'
import { LoadController } from './load-controller'
import type { ViewerAdapter } from './viewers/adapter-types'
import { createDocxAdapter, createPptxAdapter, createXlsxAdapter } from './viewers'

export const OFFICE_VIEWER_TAG_NAME = 'office-viewer'

const RELOAD_ATTRIBUTE_NAMES = new Set(['src', 'file-name', 'file-type', 'mode', 'wasm-url'])

type LifecycleState = 'idle' | 'loading' | 'ready' | 'error' | 'detached' | 'destroyed'

if (typeof globalThis.HTMLElement === 'undefined') {
  ;(globalThis as { HTMLElement: typeof HTMLElement }).HTMLElement = class {} as typeof HTMLElement
}

export class OfficeViewerElement extends HTMLElement {
  static get observedAttributes(): string[] {
    return ['src', 'file-name', 'file-type', 'mode', 'wasm-url']
  }

  private loadController = new LoadController()
  private lifecycleState: LifecycleState = 'idle'
  private adapter: ViewerAdapter | null = null
  private retainedSource: OfficeSource | null = null
  private retainedOptions: OfficeViewerLoadOptions | null = null
  private lastError: Error | null = null
  private renderContainer: HTMLElement | null = null

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
    return this.lifecycleState === 'ready'
  }

  get error(): Error | null {
    return this.lastError
  }

  get format(): OfficeFormat | null {
    return this.adapter?.format ?? null
  }

  get mode(): OfficeViewerMode | null {
    const mode = (this.adapter?.engine as { mode?: OfficeViewerMode } | null)?.mode
    if (mode === 'worker' || mode === 'main') {
      return mode
    }
    return null
  }

  connectedCallback(): void {
    if (this.lifecycleState === 'destroyed') {
      return
    }

    this.lifecycleState = 'idle'

    if (this.src?.trim()) {
      void this.loadFromAttributes()
    }
  }

  disconnectedCallback(): void {
    if (this.lifecycleState === 'destroyed') {
      return
    }

    this.loadController.cancel(createAbortError('Viewer disconnected.'))
    this.adapter?.destroy()
    this.adapter = null
    this.renderContainer = null
    this.lifecycleState = 'detached'
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

  async load(source: OfficeSource, options: OfficeViewerLoadOptions = {}): Promise<void> {
    if (this.lifecycleState === 'destroyed') {
      throw new Error('OfficeViewerElement has been destroyed. Call load() on a new element.')
    }

    const request = this.loadController.beginRequest(options.signal)

    if (request.signal.aborted) {
      request.detach()
      throw toAbortError(request.signal.reason)
    }

    this.retainedSource = source
    this.retainedOptions = options
    this.lastError = null
    this.lifecycleState = 'loading'
    this.emit('loadstart')

    try {
      const resolved = await resolveOfficeSource(source, this.readFileName(), this.readFormat(options.format))
      const adapter = await this.createAdapter(resolved.format)

      if (!this.loadController.isCurrent(request.id) || request.signal.aborted) {
        adapter.destroy()
        throw toAbortError(request.signal.reason)
      }

      this.adapter?.destroy()
      this.renderContainer = this.createRenderContainer()
      this.adapter = adapter

      await adapter.load(
        { source: resolved.source, format: resolved.format, fileName: resolved.fileName },
        {
          mode: this.readMode(options.mode),
          wasmUrl: this.readWasmUrl(options.wasmUrl),
          container: this.renderContainer
        }
      )

      if (!this.loadController.isCurrent(request.id) || request.signal.aborted) {
        throw toAbortError(request.signal.reason)
      }

      this.lifecycleState = 'ready'
      const detail: OfficeViewerReadyEventDetail = {
        format: resolved.format,
        requestedMode: this.readMode(options.mode),
        effectiveMode: this.mode ?? this.readMode(options.mode)
      }
      this.emit<OfficeViewerReadyEventDetail>('ready', detail)
    } catch (error) {
      if (!this.loadController.isCurrent(request.id)) {
        throw error
      }

      this.lifecycleState = 'error'
      this.lastError = normalizeError(error)
      this.emit<OfficeViewerLoadErrorEventDetail>('loaderror', { error: this.lastError })

      if (isAbortError(error) || request.signal.aborted) {
        throw toAbortError(request.signal.reason ?? error)
      }

      throw error
    } finally {
      request.detach()
    }
  }

  async reload(): Promise<void> {
    if (this.lifecycleState === 'destroyed') {
      throw new Error('OfficeViewerElement has been destroyed. Call load() on a new element.')
    }

    if (!this.retainedSource) {
      throw new Error('Nothing has been loaded yet.')
    }

    await this.load(this.retainedSource, this.retainedOptions ?? undefined)
  }

  destroy(): void {
    if (this.lifecycleState === 'destroyed') {
      return
    }

    this.loadController.dispose()
    this.adapter?.destroy()
    this.adapter = null
    this.renderContainer = null
    this.retainedSource = null
    this.retainedOptions = null
    this.lastError = null
    this.lifecycleState = 'destroyed'
    this.emit('destroy')
  }

  getViewer(): unknown | null {
    return this.adapter?.viewer ?? null
  }

  getDocument(): unknown | null {
    return this.adapter?.document ?? null
  }

  getEngine(): unknown | null {
    return this.adapter?.engine ?? null
  }

  private async loadFromAttributes(): Promise<void> {
    const source = this.src?.trim()
    if (!source) {
      return
    }

    try {
      await this.load(source)
    } catch {
      // load() already emits loaderror when needed.
    }
  }

  private async createAdapter(format: OfficeFormat): Promise<ViewerAdapter> {
    switch (format) {
      case 'docx':
        return createDocxAdapter()
      case 'xlsx':
        return createXlsxAdapter()
      case 'pptx':
        return createPptxAdapter()
      default:
        throw new Error(`Unsupported format: ${format}`)
    }
  }

  private createRenderContainer(): HTMLElement {
    if (typeof document === 'undefined') {
      throw new Error('OfficeViewerElement requires a browser document.')
    }

    const container = document.createElement('div')
    container.style.width = '100%'
    container.style.height = '100%'
    container.style.overflow = 'auto'
    return container
  }

  private readFormat(explicit?: OfficeFormat): OfficeFormat | undefined {
    const fromAttribute = parseFormat(this.getAttribute('file-type'))
    return explicit ?? fromAttribute
  }

  private readFileName(): string | undefined {
    return sanitizeAttribute(this.getAttribute('file-name'))
  }

  private readMode(explicit?: OfficeViewerMode): OfficeViewerMode {
    if (explicit) {
      return explicit
    }

    const fromAttribute = parseMode(this.getAttribute('mode'))
    return fromAttribute ?? 'worker'
  }

  private readWasmUrl(explicit?: string | URL): string | URL | undefined {
    if (explicit) {
      return explicit
    }

    const fromAttribute = sanitizeAttribute(this.getAttribute('wasm-url'))
    return fromAttribute
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

function parseMode(value: string | null): OfficeViewerMode | undefined {
  const normalized = value?.trim().toLowerCase()
  if (normalized === 'worker' || normalized === 'main') {
    return normalized
  }

  return undefined
}

function parseFormat(value: string | null): OfficeFormat | undefined {
  const normalized = value?.trim().toLowerCase()
  if (normalized === 'docx' || normalized === 'xlsx' || normalized === 'pptx') {
    return normalized
  }

  return undefined
}

function sanitizeAttribute(value: string | null): string | undefined {
  const trimmed = value?.trim()
  return trimmed ? trimmed : undefined
}

function normalizeError(error: unknown): Error {
  if (error instanceof Error) {
    return error
  }

  return new Error(String(error))
}

function isAbortError(error: unknown): error is DOMException {
  return error instanceof DOMException && error.name === 'AbortError'
}

function createAbortError(reason: string): DOMException {
  return new DOMException(reason, 'AbortError')
}

function toAbortError(reason: unknown): DOMException {
  if (reason instanceof DOMException && reason.name === 'AbortError') {
    return reason
  }

  if (reason instanceof Error && reason.name === 'AbortError') {
    return new DOMException(reason.message, 'AbortError')
  }

  return new DOMException(typeof reason === 'string' ? reason : 'Aborted.', 'AbortError')
}
