import type { AdapterLoadOptions, AdapterSource, ViewerAdapter } from './adapter-types'

export async function createXlsxAdapter(): Promise<ViewerAdapter> {
  const xlsx = await import('@silurus/ooxml/xlsx')

  let engine: Awaited<ReturnType<typeof xlsx.XlsxWorkbook.load>> | null = null
  let viewer: ReturnType<typeof xlsx.XlsxViewer.fromWorkbook> | null = null

  return {
    format: 'xlsx',
    get viewer() {
      return viewer
    },
    get document() {
      return engine
    },
    get engine() {
      return engine
    },
    async load(source: AdapterSource, options: AdapterLoadOptions): Promise<void> {
      const wasmUrl = normalizeWasmUrl(options.wasmUrl)
      engine = await xlsx.XlsxWorkbook.load(source.source, {
        mode: options.mode ?? 'worker',
        wasmUrl
      })
      const container = options.container ?? createDetachedContainer()
      viewer = xlsx.XlsxViewer.fromWorkbook(container, engine)
    },
    destroy(): void {
      viewer?.destroy()
      engine?.destroy()
      viewer = null
      engine = null
    }
  }
}

function normalizeWasmUrl(wasmUrl: string | URL | undefined): string | undefined {
  if (wasmUrl instanceof URL) {
    return wasmUrl.toString()
  }
  return wasmUrl
}

function createDetachedContainer(): HTMLElement {
  if (typeof document === 'undefined') {
    throw new Error('A document is required to create a viewer container.')
  }

  const container = document.createElement('div')
  container.style.width = '100%'
  container.style.height = '100%'
  container.style.overflow = 'auto'
  return container
}
