import { OFFICE_FORMATS } from './types'
import type {
  OfficeFormat,
  OfficeSource,
  OfficeViewer,
  OfficeViewerLoadErrorDetail,
  OfficeViewerLoadOptions,
  OfficeViewerMode
} from './types'

export const OFFICE_VIEWER_TAG_NAME = 'office-viewer'

const OBSERVED_ATTRIBUTES = ['src', 'file-type', 'mode', 'wasm-url'] as const
const DEFAULT_MODE: OfficeViewerMode = 'worker'

type ResolvedLoadOptions = OfficeViewerLoadOptions & { mode: OfficeViewerMode }

type ViewerConstructor = new (
  container: HTMLElement,
  options: { mode: OfficeViewerMode; wasmUrl?: string | URL }
) => OfficeViewer

// Dynamic imports keep each format's module and WASM out of the bundle until requested.
const VIEWER_MODULES: Record<OfficeFormat, () => Promise<ViewerConstructor>> = {
  docx: async () => (await import('@silurus/ooxml/docx')).DocxScrollViewer,
  xlsx: async () => (await import('@silurus/ooxml/xlsx')).XlsxViewer,
  pptx: async () => (await import('@silurus/ooxml/pptx')).PptxScrollViewer
}

interface RetainedRequest {
  source: OfficeSource
  options: ResolvedLoadOptions
}

interface PendingLoad {
  controller: AbortController
  viewer: OfficeViewer | null
}

interface ResolvedSource {
  /** Bytes handed to upstream, which may take ownership of them. */
  bytes: string | ArrayBuffer
  /** Re-readable form kept for reload(). */
  retain: string | Blob
}

export class OfficeViewerElement extends HTMLElement {
  static get observedAttributes(): string[] {
    return [...OBSERVED_ATTRIBUTES]
  }

  static readonly tagName = OFFICE_VIEWER_TAG_NAME
  static readonly shadowRootMode: ShadowRootMode = 'open'

  private viewer: OfficeViewer | null = null
  private viewerOptions: ResolvedLoadOptions | null = null
  private pending: PendingLoad | null = null
  private retained: RetainedRequest | null = null
  private lastError: Error | null = null
  // Attributes changed since the last sync. A later explicit load() outranks them.
  private attributesDirty = false
  private syncScheduled = false
  // The current request came from attributes, so clearing src unloads it.
  private attributeDriven = false
  // Torn down by removal from the document; reconnecting restores the retained request.
  private restoreOnConnect = false

  get ready(): boolean {
    return this.viewer !== null
  }

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
    return this.viewerOptions?.format ?? null
  }

  get mode(): OfficeViewerMode | null {
    return this.viewerOptions?.mode ?? null
  }

  getViewer(): OfficeViewer | null {
    return this.viewer
  }

  connectedCallback(): void {
    this.ensureContainer()
    this.scheduleSync()
  }

  disconnectedCallback(): void {
    // Deferred so a synchronous move (remove, then append) keeps the viewer alive.
    queueMicrotask(() => {
      if (this.isConnected || (!this.viewer && !this.pending)) {
        return
      }
      // Release the worker and DOM now; the request stays retained so reconnecting restores it.
      this.teardown('Removed from the document while loading.')
      this.restoreOnConnect = true
      this.dispatchEvent(new CustomEvent('destroy'))
    })
  }

  attributeChangedCallback(_name: string, oldValue: string | null, newValue: string | null): void {
    if (oldValue !== newValue) {
      this.attributesDirty = true
      this.scheduleSync()
    }
  }

  async load(source: OfficeSource, options: OfficeViewerLoadOptions): Promise<void> {
    await this.start(source, options, false)
  }

  async reload(): Promise<void> {
    if (!this.retained) {
      throw new Error('Nothing to reload: no load() has been requested yet.')
    }
    await this.start(this.retained.source, this.retained.options, this.attributeDriven)
  }

  destroy(): void {
    this.teardown('Destroyed while loading.')
    this.retained = null
    this.attributeDriven = false
    this.restoreOnConnect = false
    this.dispatchEvent(new CustomEvent('destroy'))
  }

  private async start(source: OfficeSource, options: OfficeViewerLoadOptions, attributeDriven: boolean): Promise<void> {
    this.cancelActiveLoad('Superseded by a newer load.')
    // An explicit load() outranks attribute changes made earlier in the same task.
    this.attributesDirty = false
    this.attributeDriven = attributeDriven
    this.restoreOnConnect = false

    const pending: PendingLoad = { controller: new AbortController(), viewer: null }
    const { signal } = pending.controller
    // Copied and resolved once, so caller-side mutation cannot change what was requested.
    const request: RetainedRequest = { source, options: { ...options, mode: options.mode ?? DEFAULT_MODE } }
    this.pending = pending
    this.retained = request
    this.lastError = null
    this.dispatchEvent(new CustomEvent('loadstart'))

    let viewer: OfficeViewer | null = null
    try {
      // A loadstart listener may have called load() or destroy() synchronously.
      signal.throwIfAborted()
      assertSupportedFormat(request.options.format)
      const [{ bytes, retain }, Viewer] = await abortable(
        Promise.all([resolveSource(source, signal), VIEWER_MODULES[request.options.format]()]),
        signal
      )
      signal.throwIfAborted()
      // Upstream may detach the bytes it is given, so reload() replays a re-readable copy.
      request.source = retain
      viewer = new Viewer(this.ensureContainer(), { mode: request.options.mode, wasmUrl: request.options.wasmUrl })
      pending.viewer = viewer
      // Upstream settles load() only after its fetch and parse finish, even once destroyed,
      // so a cancelled load is settled here instead of waiting for it.
      await abortable(viewer.load(bytes), signal)
      signal.throwIfAborted()
    } catch (reason) {
      if (this.pending !== pending) {
        // A newer load or destroy() owns the element now; this load's viewer is already destroyed.
        viewer?.destroy()
        throw signal.reason
      }
      this.pending = null
      // Stops a source read that may still be running alongside the step that failed.
      pending.controller.abort(reason)
      viewer?.destroy()
      const error = toError(reason)
      this.lastError = error
      this.dispatchEvent(new CustomEvent<OfficeViewerLoadErrorDetail>('loaderror', { detail: { error } }))
      throw error
    }

    this.pending = null
    const previous = this.viewer
    this.viewer = viewer
    this.viewerOptions = request.options
    previous?.destroy()
    this.dispatchEvent(new CustomEvent('ready'))
  }

  private teardown(reason: string): void {
    this.cancelActiveLoad(reason)
    this.viewer?.destroy()
    this.viewer = null
    this.viewerOptions = null
    this.lastError = null
  }

  private cancelActiveLoad(reason: string): void {
    const pending = this.pending
    if (!pending) {
      return
    }
    this.pending = null
    pending.controller.abort(createAbortError(reason))
    pending.viewer?.destroy()
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

  private scheduleSync(): void {
    if (this.syncScheduled) {
      return
    }
    this.syncScheduled = true
    // Coalesced so attributes set together start one load, not one per attribute.
    queueMicrotask(() => {
      this.syncScheduled = false
      if (this.isConnected) {
        this.sync()
      }
    })
  }

  private sync(): void {
    if (this.attributesDirty) {
      this.attributesDirty = false
      if (this.loadFromAttributes()) {
        return
      }
      if (this.attributeDriven) {
        // src was cleared, so an attribute-driven document has nothing left to show.
        this.destroy()
        return
      }
    }
    if (this.restoreOnConnect && this.retained) {
      this.reload().catch(() => {
        // start() has already emitted loaderror.
      })
    }
  }

  private loadFromAttributes(): boolean {
    const src = this.getAttribute('src')?.trim()
    if (!src) {
      return false
    }
    const options: OfficeViewerLoadOptions = {
      // Validated by start() so a missing or unknown file-type surfaces as loaderror.
      format: (this.getAttribute('file-type')?.trim().toLowerCase() ?? '') as OfficeFormat,
      mode: parseMode(this.getAttribute('mode')),
      wasmUrl: this.getAttribute('wasm-url')?.trim() || undefined
    }
    this.start(src, options, true).catch(() => {
      // start() has already emitted loaderror.
    })
    return true
  }
}

export function defineOfficeViewerElement(tagName = OFFICE_VIEWER_TAG_NAME): typeof OfficeViewerElement {
  if (typeof customElements === 'undefined') {
    throw new Error('Custom elements are not available in this runtime.')
  }
  const existing = customElements.get(tagName)
  if (existing) {
    return existing as typeof OfficeViewerElement
  }
  customElements.define(tagName, OfficeViewerElement)
  return OfficeViewerElement
}

function assertSupportedFormat(format: unknown): asserts format is OfficeFormat {
  if (typeof format === 'string' && (OFFICE_FORMATS as readonly string[]).includes(format)) {
    return
  }
  const expected = `Expected one of ${OFFICE_FORMATS.map((name) => `"${name}"`).join(', ')}.`
  throw new Error(
    format === undefined || format === null || format === ''
      ? `Missing format. ${expected}`
      : `Unsupported format "${String(format)}". ${expected}`
  )
}

function parseMode(value: string | null): OfficeViewerMode | undefined {
  const normalized = value?.trim().toLowerCase()
  if (normalized === 'worker' || normalized === 'main') {
    return normalized
  }
  return undefined
}

async function resolveSource(source: OfficeSource, signal: AbortSignal): Promise<ResolvedSource> {
  if (typeof source === 'string') {
    return { bytes: source, retain: source }
  }
  if (source instanceof ArrayBuffer) {
    return { bytes: source, retain: new Blob([source]) }
  }
  if (typeof Blob !== 'undefined' && source instanceof Blob) {
    return { bytes: await source.arrayBuffer(), retain: source }
  }
  if (typeof ReadableStream !== 'undefined' && source instanceof ReadableStream) {
    const blob = await readStreamToBlob(source, signal)
    return { bytes: await blob.arrayBuffer(), retain: blob }
  }
  throw new Error(
    'Unsupported source type. Expected a URL string, ArrayBuffer, Blob, File, or ReadableStream<Uint8Array>.'
  )
}

async function readStreamToBlob(stream: ReadableStream<Uint8Array>, signal: AbortSignal): Promise<Blob> {
  const reader = stream.getReader()
  // Cancelling makes a pending read() resolve as done, so a cancelled load stops consuming the stream.
  const cancel = (): void => {
    reader.cancel(signal.reason).catch(() => {})
  }
  signal.addEventListener('abort', cancel, { once: true })

  const chunks: BlobPart[] = []
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) {
        break
      }
      // Chunks typed over ArrayBufferLike are still valid Blob parts at runtime.
      chunks.push(value as BlobPart)
    }
  } finally {
    signal.removeEventListener('abort', cancel)
    reader.releaseLock()
  }
  return new Blob(chunks)
}

// Settles with the signal's reason as soon as it aborts, even if the wrapped promise never settles.
function abortable<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const onAbort = (): void => reject(signal.reason)
    if (signal.aborted) {
      onAbort()
      return
    }
    signal.addEventListener('abort', onAbort, { once: true })
    promise.then(resolve, reject).finally(() => signal.removeEventListener('abort', onAbort))
  })
}

function createAbortError(message: string): DOMException {
  return new DOMException(message, 'AbortError')
}

function toError(reason: unknown): Error {
  return reason instanceof Error ? reason : new Error(String(reason))
}
