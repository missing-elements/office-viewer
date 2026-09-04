import {
  OoxmlIntegrationSpike,
  type OfficeFormat,
  type OfficeSpikeError,
  type OfficeSpikeLoadOptions,
  type OfficeSpikeRunSummary,
  type OfficeSource,
  type OfficeViewerMode
} from './ooxml-spike'
import { CHANGE_EVENT_BY_FORMAT } from './viewer-adapter'

export interface OfficeViewerLoadOptions extends OfficeSpikeLoadOptions {
  signal?: AbortSignal
}

export interface OfficeViewerReadyEventDetail {
  summary: OfficeSpikeRunSummary
}

export interface OfficeViewerLoadErrorEventDetail {
  error: OfficeSpikeError
  summary: OfficeSpikeRunSummary | null
}

export interface OfficeViewerProgressEventDetail {
  summary: OfficeSpikeRunSummary
}

export interface OfficeViewerIndexChangeEventDetail {
  format: OfficeFormat
  index: number
  totalCount?: number
  layoutComplete?: boolean
  sheetNames?: string[]
  summary: OfficeSpikeRunSummary
}

export const OFFICE_VIEWER_TAG_NAME = 'office-viewer'

const RELOAD_ATTRIBUTE_NAMES = new Set(['src', 'file-name', 'file-type', 'mode', 'wasm-url'])

interface ActiveRequest {
  id: number
  signal: AbortSignal
  detach(): void
}

if (typeof globalThis.HTMLElement === 'undefined') {
  ;(globalThis as { HTMLElement: typeof HTMLElement }).HTMLElement = class {} as typeof HTMLElement
}

export class OfficeViewerElement extends HTMLElement {
  static get observedAttributes(): string[] {
    return ['src', 'file-name', 'file-type', 'mode', 'wasm-url']
  }

  private status: HTMLParagraphElement | null = null
  private viewport: HTMLDivElement | null = null
  private harness: OoxmlIntegrationSpike | null = null
  private unsubscribeSummary: (() => void) | null = null
  private lastProgressSignature: string | null = null
  private requestGeneration = 0
  private activeRequestController: AbortController | null = null

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

  connectedCallback(): void {
    this.ensureInitialized()
    this.setBusy(false)

    if (this.src?.trim()) {
      void this.loadFromAttributes()
    }
  }

  disconnectedCallback(): void {
    this.cancelActiveRequest(createAbortError('Viewer disconnected.'))
    this.harness?.destroy()
    this.setBusy(false)
  }

  attributeChangedCallback(name: string, oldValue: string | null, newValue: string | null): void {
    if (oldValue === newValue || !this.isConnected || !RELOAD_ATTRIBUTE_NAMES.has(name)) {
      return
    }

    if (name === 'src' && !newValue?.trim()) {
      this.destroy()
      this.updateStatus('Idle', true)
      return
    }

    if (this.src?.trim()) {
      void this.loadFromAttributes()
    }
  }

  async load(source: OfficeSource, options: OfficeViewerLoadOptions = {}): Promise<OfficeSpikeRunSummary> {
    const { signal, ...loadOptions } = options
    const harness = this.getHarness()
    const request = this.beginRequest(signal)

    if (request.signal.aborted) {
      request.detach()
      throw toAbortError(request.signal.reason)
    }

    this.setBusy(true)
    this.updateStatus('Loading document...')
    this.emit('loadstart')

    try {
      const summary = await harness.load(source, this.mergeWithAttributes(loadOptions))

      if (!this.isCurrentRequest(request.id) || request.signal.aborted) {
        throw toAbortError(request.signal.reason)
      }

      this.setBusy(false)
      this.updateStatus(`${summary.format.toUpperCase()} ready (${summary.effectiveMode})`, true)
      this.emit<OfficeViewerReadyEventDetail>('ready', { summary })
      return summary
    } catch (error) {
      if (!this.isCurrentRequest(request.id)) {
        throw error
      }

      if (isAbortError(error) || request.signal.aborted) {
        this.setBusy(false)
        this.updateStatus('Load canceled.')
        throw toAbortError(request.signal.reason ?? error)
      }

      this.setBusy(false)
      const normalized = normalizeError(error)
      this.updateStatus(`Load failed: ${normalized.message}`)
      this.emit<OfficeViewerLoadErrorEventDetail>('loaderror', {
        error: normalized,
        summary: harness.getSummary()
      })
      throw error
    } finally {
      request.detach()
    }
  }

  async reload(): Promise<OfficeSpikeRunSummary> {
    const harness = this.getHarness()
    const request = this.beginRequest()

    if (request.signal.aborted) {
      request.detach()
      throw toAbortError(request.signal.reason)
    }

    this.setBusy(true)
    this.updateStatus('Reloading document...')
    this.emit('loadstart')

    try {
      const summary = await harness.reload()

      if (!this.isCurrentRequest(request.id) || request.signal.aborted) {
        throw toAbortError(request.signal.reason)
      }

      this.setBusy(false)
      this.updateStatus(`${summary.format.toUpperCase()} reloaded (${summary.effectiveMode})`, true)
      this.emit<OfficeViewerReadyEventDetail>('ready', { summary })
      return summary
    } catch (error) {
      if (!this.isCurrentRequest(request.id)) {
        throw error
      }

      if (isAbortError(error) || request.signal.aborted) {
        this.setBusy(false)
        this.updateStatus('Reload canceled.')
        throw toAbortError(request.signal.reason ?? error)
      }

      this.setBusy(false)
      const normalized = normalizeError(error)
      this.updateStatus(`Reload failed: ${normalized.message}`)
      this.emit<OfficeViewerLoadErrorEventDetail>('loaderror', {
        error: normalized,
        summary: harness.getSummary()
      })
      throw error
    } finally {
      request.detach()
    }
  }

  getSummary(): OfficeSpikeRunSummary | null {
    return this.harness?.getSummary() ?? null
  }

  goToPage(pageIndex: number): boolean {
    return this.getHarness().goToPage(pageIndex)
  }

  goToSlide(slideIndex: number): boolean {
    return this.getHarness().goToSlide(slideIndex)
  }

  goToSheet(sheetIndex: number): boolean {
    return this.getHarness().goToSheet(sheetIndex)
  }

  destroy(): void {
    this.cancelActiveRequest(createAbortError('Viewer destroyed.'))
    this.requestGeneration += 1
    this.harness?.destroy()
    this.setBusy(false)
    this.emit('destroy')
  }

  private mergeWithAttributes(options: OfficeViewerLoadOptions): OfficeSpikeLoadOptions {
    const attributeOptions = this.readAttributeOptions()
    return {
      format: options.format ?? attributeOptions.format,
      fileName: options.fileName ?? attributeOptions.fileName,
      mode: options.mode ?? attributeOptions.mode,
      wasmUrl: options.wasmUrl ?? attributeOptions.wasmUrl
    }
  }

  private readAttributeOptions(): OfficeSpikeLoadOptions {
    return {
      format: parseFormat(this.getAttribute('file-type')),
      fileName: sanitizeAttribute(this.getAttribute('file-name')),
      mode: parseMode(this.getAttribute('mode')),
      wasmUrl: sanitizeAttribute(this.getAttribute('wasm-url'))
    }
  }

  private async loadFromAttributes(): Promise<void> {
    const source = this.src?.trim()
    if (!source) {
      return
    }

    try {
      await this.load(source)
    } catch {
      // load() already updates status and emits loaderror when needed.
    }
  }

  private setBusy(isBusy: boolean): void {
    this.setAttribute('aria-busy', isBusy ? 'true' : 'false')
  }

  private updateStatus(message: string, hide = false): void {
    const status = this.status
    if (!status) {
      return
    }

    status.hidden = hide
    status.textContent = message
  }

  private emit<T>(name: string, detail?: T): void {
    if (typeof CustomEvent === 'function') {
      this.dispatchEvent(new CustomEvent(name, { detail }))
      return
    }

    this.dispatchEvent(new Event(name))
  }

  private getRenderRoot(): ShadowRoot | this {
    if (typeof this.attachShadow !== 'function') {
      return this
    }

    try {
      return this.attachShadow({ mode: 'open' })
    } catch {
      // Some browser automation environments disable Shadow DOM APIs.
      return this
    }
  }

  private ensureInitialized(): void {
    if (this.harness && this.status && this.viewport) {
      return
    }

    if (typeof document === 'undefined') {
      throw new Error('OfficeViewerElement requires a browser document.')
    }

    const renderRoot = this.getRenderRoot()
    const style = document.createElement('style')
    style.textContent = [
      ':host { display: block; min-height: 16rem; }',
      '.status { margin: 0 0 0.5rem; color: #9ca3af; font: 500 0.875rem/1.4 ui-sans-serif, system-ui, sans-serif; }',
      '.status[hidden] { display: none; }',
      '.viewport { height: 100%; min-height: 0; border-radius: 0.5rem; overflow: hidden; }'
    ].join('')

    const status = document.createElement('p')
    status.className = 'status'
    status.setAttribute('part', 'status')
    status.setAttribute('role', 'status')
    status.setAttribute('aria-live', 'polite')

    const viewport = document.createElement('div')
    viewport.className = 'viewport'
    viewport.setAttribute('part', 'viewport')

    renderRoot.append(style, status, viewport)

    this.status = status
    this.viewport = viewport
    this.unsubscribeSummary?.()
    this.harness = new OoxmlIntegrationSpike(viewport)
    this.unsubscribeSummary = this.harness.subscribeSummary((summary) => {
      this.handleSummary(summary)
    })
    this.updateStatus('Idle', true)
  }

  private beginRequest(externalSignal?: AbortSignal): ActiveRequest {
    const id = ++this.requestGeneration
    this.lastProgressSignature = null
    this.cancelActiveRequest(createAbortError('Superseded by a newer request.'))

    const controller = new AbortController()
    this.activeRequestController = controller

    let removeExternalListener = () => {}
    if (externalSignal) {
      if (externalSignal.aborted) {
        controller.abort(externalSignal.reason)
      } else {
        const onExternalAbort = () => {
          controller.abort(externalSignal.reason)
        }

        externalSignal.addEventListener('abort', onExternalAbort, { once: true })
        removeExternalListener = () => {
          externalSignal.removeEventListener('abort', onExternalAbort)
        }
      }
    }

    const onAbort = () => {
      if (!this.isCurrentRequest(id)) {
        return
      }

      this.harness?.destroy()
      this.setBusy(false)
      this.updateStatus('Load canceled.')
    }

    controller.signal.addEventListener('abort', onAbort, { once: true })

    return {
      id,
      signal: controller.signal,
      detach: () => {
        removeExternalListener()
        controller.signal.removeEventListener('abort', onAbort)
        if (this.activeRequestController === controller) {
          this.activeRequestController = null
        }
      }
    }
  }

  private cancelActiveRequest(reason: unknown): void {
    if (!this.activeRequestController) {
      return
    }

    this.activeRequestController.abort(reason)
    this.activeRequestController = null
  }

  private isCurrentRequest(id: number): boolean {
    return this.requestGeneration === id
  }

  private getHarness(): OoxmlIntegrationSpike {
    this.ensureInitialized()
    if (!this.harness) {
      throw new Error('OfficeViewerElement failed to initialize harness.')
    }

    return this.harness
  }

  private handleSummary(summary: OfficeSpikeRunSummary): void {
    const signature = summarySignature(summary)
    if (signature === this.lastProgressSignature) {
      return
    }

    this.lastProgressSignature = signature
    this.emit<OfficeViewerProgressEventDetail>('progress', {
      summary
    })

    if (typeof summary.currentIndex !== 'number') {
      return
    }

    this.emit<OfficeViewerIndexChangeEventDetail>(CHANGE_EVENT_BY_FORMAT[summary.format], {
      format: summary.format,
      index: summary.currentIndex,
      totalCount: summary.totalCount,
      layoutComplete: summary.layoutComplete,
      sheetNames: summary.sheetNames ? [...summary.sheetNames] : undefined,
      summary
    })
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

function normalizeError(error: unknown): OfficeSpikeError {
  const value = error as { name?: string, message?: string, code?: string } | undefined
  return {
    name: value?.name ?? 'Error',
    message: value?.message ?? String(error),
    code: value?.code
  }
}

function isAbortError(error: unknown): error is Error {
  return error instanceof Error && error.name === 'AbortError'
}

function createAbortError(message: string): Error {
  if (typeof DOMException === 'function') {
    return new DOMException(message, 'AbortError')
  }

  const error = new Error(message)
  error.name = 'AbortError'
  return error
}

function toAbortError(reason: unknown): Error {
  if (isAbortError(reason)) {
    return reason
  }

  if (reason instanceof Error) {
    if (reason.name === 'AbortError') {
      return reason
    }

    return createAbortError(reason.message)
  }

  if (typeof reason === 'string' && reason.trim()) {
    return createAbortError(reason)
  }

  return createAbortError('The operation was aborted.')
}

function summarySignature(summary: OfficeSpikeRunSummary): string {
  const sheetNames = summary.sheetNames?.join(',') ?? ''
  return [
    summary.generation,
    summary.format,
    summary.currentIndex ?? -1,
    summary.totalCount ?? -1,
    summary.layoutComplete ? 1 : 0,
    sheetNames,
    summary.lastError?.message ?? ''
  ].join('|')
}