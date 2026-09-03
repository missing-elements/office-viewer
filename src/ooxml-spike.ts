export type OfficeFormat = 'docx' | 'xlsx' | 'pptx'
export type OfficeViewerMode = 'worker' | 'main'
export type OfficeSource = string | URL | File | Blob | ArrayBuffer | Uint8Array
export type OfficeSourceKind = 'url' | 'file' | 'blob' | 'array-buffer' | 'uint8array'

export interface OfficeSpikeLoadOptions {
  format?: OfficeFormat
  fileName?: string
  mode?: OfficeViewerMode
  wasmUrl?: string | URL
}

export interface OfficeSpikeError {
  name: string
  message: string
  code?: string
}

export interface OfficeSpikeRunSummary {
  format: OfficeFormat
  viewerKind: 'DocxScrollViewer' | 'XlsxViewer' | 'PptxScrollViewer'
  sourceKind: OfficeSourceKind
  fileName?: string
  requestedMode: OfficeViewerMode
  effectiveMode: OfficeViewerMode
  firstVisibleMs?: number
  loadCompletedMs?: number
  currentIndex?: number
  totalCount?: number
  layoutComplete?: boolean
  sheetNames?: string[]
  lastError?: OfficeSpikeError
  generation: number
  diagnostics: string[]
}

interface RetainedLoad {
  readonly source: string | ArrayBuffer
  readonly sourceKind: OfficeSourceKind
  readonly format: OfficeFormat
  readonly fileName?: string
  readonly options: OfficeSpikeLoadOptions
}

type ViewerInstance = {
  destroy(): void
}

type EngineInstance = {
  readonly mode: OfficeViewerMode
  destroy(): void
}

type PendingState = {
  readonly startedAt: number
  summary: OfficeSpikeRunSummary
}

const EXTENSION_TO_FORMAT: Record<string, OfficeFormat> = {
  docx: 'docx',
  xlsx: 'xlsx',
  pptx: 'pptx'
}

export const SAMPLE_FIXTURES: Record<OfficeFormat, string> = {
  docx: '/fixtures/sample.docx',
  xlsx: '/fixtures/sample.xlsx',
  pptx: '/fixtures/sample.pptx'
}

export function detectOfficeFormat(source: Pick<RetainedLoad, 'fileName'> | { fileName?: string }, explicitFormat?: OfficeFormat): OfficeFormat {
  if (explicitFormat) {
    return explicitFormat
  }

  const fileName = source.fileName?.trim()
  const extension = fileName?.split('.').pop()?.toLowerCase()
  if (extension && extension in EXTENSION_TO_FORMAT) {
    return EXTENSION_TO_FORMAT[extension]
  }

  throw new Error('Unable to determine Office format. Provide options.format, fileName, or a URL/file name with .docx, .xlsx, or .pptx.')
}

export async function normalizeOfficeSource(source: OfficeSource, options: OfficeSpikeLoadOptions = {}): Promise<RetainedLoad> {
  const explicitFileName = options.fileName?.trim() || undefined

  if (typeof source === 'string') {
    return {
      source,
      sourceKind: 'url',
      format: detectOfficeFormat({ fileName: explicitFileName ?? inferFileNameFromUrl(source) }, options.format),
      fileName: explicitFileName ?? inferFileNameFromUrl(source),
      options
    }
  }

  if (source instanceof URL) {
    const href = source.toString()
    return {
      source: href,
      sourceKind: 'url',
      format: detectOfficeFormat({ fileName: explicitFileName ?? inferFileNameFromUrl(href) }, options.format),
      fileName: explicitFileName ?? inferFileNameFromUrl(href),
      options
    }
  }

  if (source instanceof File) {
    return {
      source: copyArrayBuffer(await source.arrayBuffer()),
      sourceKind: 'file',
      format: detectOfficeFormat({ fileName: explicitFileName ?? source.name }, options.format),
      fileName: explicitFileName ?? source.name,
      options
    }
  }

  if (source instanceof Blob) {
    return {
      source: copyArrayBuffer(await source.arrayBuffer()),
      sourceKind: 'blob',
      format: detectOfficeFormat({ fileName: explicitFileName }, options.format),
      fileName: explicitFileName,
      options
    }
  }

  if (source instanceof Uint8Array) {
    return {
      source: source.slice().buffer,
      sourceKind: 'uint8array',
      format: detectOfficeFormat({ fileName: explicitFileName }, options.format),
      fileName: explicitFileName,
      options
    }
  }

  return {
    source: copyArrayBuffer(source),
    sourceKind: 'array-buffer',
    format: detectOfficeFormat({ fileName: explicitFileName }, options.format),
    fileName: explicitFileName,
    options
  }
}

export class OoxmlIntegrationSpike {
  private readonly container: HTMLElement
  private retainedLoad: RetainedLoad | null = null
  private activeViewer: ViewerInstance | null = null
  private activeEngine: EngineInstance | null = null
  private generation = 0
  private pendingState: PendingState | null = null

  constructor(container: HTMLElement) {
    this.container = container
  }

  getSummary(): OfficeSpikeRunSummary | null {
    return this.pendingState?.summary ?? null
  }

  async load(source: OfficeSource, options: OfficeSpikeLoadOptions = {}): Promise<OfficeSpikeRunSummary> {
    const retainedLoad = await normalizeOfficeSource(source, {
      mode: 'worker',
      ...options
    })

    const generation = ++this.generation
    this.destroyActive()
    this.retainedLoad = retainedLoad
    this.pendingState = {
      startedAt: performance.now(),
      summary: {
        format: retainedLoad.format,
        viewerKind: retainedLoad.format === 'docx' ? 'DocxScrollViewer' : retainedLoad.format === 'xlsx' ? 'XlsxViewer' : 'PptxScrollViewer',
        sourceKind: retainedLoad.sourceKind,
        fileName: retainedLoad.fileName,
        requestedMode: retainedLoad.options.mode ?? 'worker',
        effectiveMode: retainedLoad.options.mode ?? 'worker',
        generation,
        diagnostics: []
      }
    }

    if (retainedLoad.sourceKind === 'file' || retainedLoad.sourceKind === 'blob' || retainedLoad.sourceKind === 'uint8array') {
      this.pendingState.summary.diagnostics.push('Normalized File, Blob, and Uint8Array inputs to ArrayBuffer because upstream declarations only type load(source) as string | ArrayBuffer.')
    }

    try {
      const loaded = retainedLoad.format === 'docx'
        ? await this.loadDocx(retainedLoad, generation)
        : retainedLoad.format === 'xlsx'
          ? await this.loadXlsx(retainedLoad, generation)
          : await this.loadPptx(retainedLoad, generation)

      if (!this.isCurrent(generation)) {
        loaded.viewer.destroy()
        loaded.engine.destroy()
        throw new Error('Stale load ignored.')
      }

      this.activeViewer = loaded.viewer
      this.activeEngine = loaded.engine
      this.pendingState.summary.effectiveMode = loaded.engine.mode
      this.finishVisibleTick()
      return this.pendingState.summary
    } catch (error) {
      if (this.isCurrent(generation)) {
        this.pendingState.summary.lastError = toOfficeSpikeError(error)
        this.pendingState.summary.diagnostics.push(formatErrorDiagnostic(error))
      }

      throw error
    }
  }

  async reload(): Promise<OfficeSpikeRunSummary> {
    if (!this.retainedLoad) {
      throw new Error('Nothing has been loaded yet.')
    }

    const source = typeof this.retainedLoad.source === 'string'
      ? this.retainedLoad.source
      : copyArrayBuffer(this.retainedLoad.source)

    return this.load(source, this.retainedLoad.options)
  }

  destroy(): void {
    this.generation += 1
    this.retainedLoad = null
    this.pendingState = null
    this.destroyActive()
  }

  private destroyActive(): void {
    this.activeViewer?.destroy()
    this.activeEngine?.destroy()
    this.activeViewer = null
    this.activeEngine = null
    this.container.replaceChildren()
  }

  private isCurrent(generation: number): boolean {
    return this.pendingState?.summary.generation === generation
  }

  private markVisible(generation: number, currentIndex: number, totalCount?: number, layoutComplete?: boolean): void {
    if (!this.isCurrent(generation) || !this.pendingState) {
      return
    }

    const { summary } = this.pendingState
    summary.currentIndex = currentIndex
    summary.totalCount = totalCount
    summary.layoutComplete = layoutComplete
    this.finishVisibleTick()
  }

  private finishVisibleTick(): void {
    if (!this.pendingState) {
      return
    }

    const elapsed = performance.now() - this.pendingState.startedAt
    this.pendingState.summary.firstVisibleMs ??= elapsed
    this.pendingState.summary.loadCompletedMs = elapsed
  }

  private async loadDocx(retainedLoad: RetainedLoad, generation: number): Promise<{ engine: EngineInstance, viewer: ViewerInstance }> {
    const { DocxDocument, DocxScrollViewer } = await import('@silurus/ooxml/docx')
    const engine = await DocxDocument.load(cloneRetainedSource(retainedLoad), {
      mode: retainedLoad.options.mode ?? 'worker',
      wasmUrl: retainedLoad.options.wasmUrl,
      progressiveLayout: true
    })

    const viewer = DocxScrollViewer.fromDocument(this.container, engine, {
      enableHyperlinks: true,
      enableTextSelection: true,
      onVisiblePageChange: (topIndex, total, layoutComplete) => {
        this.markVisible(generation, topIndex, total, layoutComplete)
      },
      onError: (error) => {
        if (this.isCurrent(generation) && this.pendingState) {
          this.pendingState.summary.lastError = toOfficeSpikeError(error)
          this.pendingState.summary.diagnostics.push(formatErrorDiagnostic(error))
        }
      }
    })

    this.markVisible(generation, viewer.topVisiblePage, viewer.pageCount, viewer.layoutComplete)
    return { engine, viewer }
  }

  private async loadXlsx(retainedLoad: RetainedLoad, generation: number): Promise<{ engine: EngineInstance, viewer: ViewerInstance }> {
    const { XlsxViewer, XlsxWorkbook } = await import('@silurus/ooxml/xlsx')
    const engine = await XlsxWorkbook.load(cloneRetainedSource(retainedLoad), {
      mode: retainedLoad.options.mode ?? 'worker',
      wasmUrl: retainedLoad.options.wasmUrl
    })

    const viewer = XlsxViewer.fromWorkbook(this.container, engine, {
      enableHyperlinks: true,
      onReady: (sheetNames) => {
        if (this.isCurrent(generation) && this.pendingState) {
          this.pendingState.summary.sheetNames = [...sheetNames]
          this.markVisible(generation, 0, sheetNames.length, true)
        }
      },
      onSheetChange: (index, total) => {
        this.markVisible(generation, index, total, true)
      },
      onError: (error) => {
        if (this.isCurrent(generation) && this.pendingState) {
          this.pendingState.summary.lastError = toOfficeSpikeError(error)
          this.pendingState.summary.diagnostics.push(formatErrorDiagnostic(error))
        }
      }
    })
    this.markVisible(generation, viewer.sheetIndex, viewer.sheetCount, true)
    this.pendingState!.summary.sheetNames = [...viewer.sheetNames]
    return { engine, viewer }
  }

  private async loadPptx(retainedLoad: RetainedLoad, generation: number): Promise<{ engine: EngineInstance, viewer: ViewerInstance }> {
    const { PptxPresentation, PptxScrollViewer } = await import('@silurus/ooxml/pptx')
    const engine = await PptxPresentation.load(cloneRetainedSource(retainedLoad), {
      mode: retainedLoad.options.mode ?? 'worker',
      wasmUrl: retainedLoad.options.wasmUrl,
      progressiveLayout: true
    })

    const viewer = PptxScrollViewer.fromPresentation(this.container, engine, {
      enableHyperlinks: true,
      enableTextSelection: true,
      onVisibleSlideChange: (topIndex, total, layoutComplete) => {
        this.markVisible(generation, topIndex, total, layoutComplete)
      },
      onError: (error) => {
        if (this.isCurrent(generation) && this.pendingState) {
          this.pendingState.summary.lastError = toOfficeSpikeError(error)
          this.pendingState.summary.diagnostics.push(formatErrorDiagnostic(error))
        }
      }
    })

    this.markVisible(generation, viewer.topVisibleSlide, viewer.slideCount, viewer.layoutComplete)
    return { engine, viewer }
  }
}

function inferFileNameFromUrl(urlString: string): string | undefined {
  try {
    const url = new URL(urlString, window.location.href)
    const fileName = url.pathname.split('/').filter(Boolean).at(-1)
    return fileName || undefined
  } catch {
    return undefined
  }
}

function cloneRetainedSource(retainedLoad: RetainedLoad): string | ArrayBuffer {
  return typeof retainedLoad.source === 'string' ? retainedLoad.source : copyArrayBuffer(retainedLoad.source)
}

function copyArrayBuffer(buffer: ArrayBuffer): ArrayBuffer {
  return buffer.slice(0)
}

function toOfficeSpikeError(error: unknown): OfficeSpikeError {
  const value = error as { name?: string, message?: string, code?: string } | undefined
  return {
    name: value?.name ?? 'Error',
    message: value?.message ?? String(error),
    code: value?.code
  }
}

function formatErrorDiagnostic(error: unknown): string {
  const value = toOfficeSpikeError(error)
  return value.code ? `${value.name} (${value.code}): ${value.message}` : `${value.name}: ${value.message}`
}
