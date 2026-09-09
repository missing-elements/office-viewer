import {
  defineOfficeViewerElement,
  type OfficeFormat,
  type OfficeSource,
  type OfficeViewerElement,
  type OfficeViewerReadyEventDetail
} from '../src'

const SAMPLE_FIXTURES: Record<OfficeFormat, string> = {
  docx: '/fixtures/sample.docx',
  xlsx: '/fixtures/sample.xlsx',
  pptx: '/fixtures/sample.pptx'
}

const viewer = document.querySelector<HTMLDivElement>('#viewer')
const formatSelect = document.querySelector<HTMLSelectElement>('#format')
const modeSelect = document.querySelector<HTMLSelectElement>('#mode')
const urlInput = document.querySelector<HTMLInputElement>('#url')
const fileInput = document.querySelector<HTMLInputElement>('#file')
const wasmUrlInput = document.querySelector<HTMLInputElement>('#wasmUrl')
const status = document.querySelector<HTMLElement>('#status')
const summary = document.querySelector<HTMLElement>('#summary')

if (!viewer || !formatSelect || !modeSelect || !urlInput || !fileInput || !wasmUrlInput || !status || !summary) {
  throw new Error('Demo UI is missing required elements.')
}

const ui = { viewer, formatSelect, modeSelect, urlInput, fileInput, wasmUrlInput, status, summary }
const viewerElement = createViewerElement(ui)

const currentFormat = (): OfficeFormat => ui.formatSelect.value as OfficeFormat

function buildOptions(): import('../src').OfficeViewerLoadOptions {
  return {
    format: currentFormat(),
    mode: ui.modeSelect.value === 'main' ? 'main' : 'worker',
    wasmUrl: ui.wasmUrlInput.value.trim() ? ui.wasmUrlInput.value.trim() : undefined
  }
}

function refreshSummary(): void {
  if (!viewerElement) {
    ui.summary.textContent = 'null'
    return
  }

  ui.summary.textContent = JSON.stringify(
    {
      ready: viewerElement.ready,
      error: viewerElement.error ? { name: viewerElement.error.name, message: viewerElement.error.message } : null,
      format: viewerElement.format,
      mode: viewerElement.mode,
      viewerType: viewerElement.getViewer()?.constructor?.name ?? null,
      documentType: viewerElement.getDocument()?.constructor?.name ?? null,
      engineType: viewerElement.getEngine()?.constructor?.name ?? null
    },
    null,
    2
  )
}

async function runLoad(source: OfficeSource, options = buildOptions()): Promise<void> {
  if (!viewerElement) {
    ui.status.textContent = 'Viewer failed to initialize. Check console for details.'
    return
  }

  ui.status.textContent = 'Loading…'
  refreshSummary()

  try {
    await viewerElement.load(source, options)
    const detail: OfficeViewerReadyEventDetail = {
      format: viewerElement.format ?? options.format ?? currentFormat(),
      requestedMode: options.mode ?? 'worker',
      effectiveMode: viewerElement.mode ?? options.mode ?? 'worker'
    }
    ui.status.textContent = `${detail.format.toUpperCase()} ready (${detail.effectiveMode})`
    refreshSummary()
  } catch (error) {
    ui.status.textContent = `Load failed: ${error instanceof Error ? error.message : String(error)}`
    refreshSummary()
  }
}

function syncSampleUrl(): void {
  ui.urlInput.value = SAMPLE_FIXTURES[currentFormat()]
}

if (viewerElement) {
  for (const eventName of ['ready', 'loaderror', 'destroy']) {
    viewerElement.addEventListener(eventName, () => {
      refreshSummary()
    })
  }
}

ui.formatSelect.addEventListener('change', syncSampleUrl)
document.querySelector('#load-url')?.addEventListener('click', () => {
  void runLoad(ui.urlInput.value, buildOptions())
})
document.querySelector('#load-file')?.addEventListener('click', () => {
  const file = ui.fileInput.files?.[0]
  if (!file) {
    ui.status.textContent = 'Choose a local file first.'
    return
  }

  void runLoad(file, buildOptions())
})
document.querySelector('#sample-url')?.addEventListener('click', () => {
  const format = currentFormat()
  void runLoad(SAMPLE_FIXTURES[format], {
    ...buildOptions(),
    format
  })
})
document.querySelector('#sample-file')?.addEventListener('click', async () => {
  const format = currentFormat()
  const fixtureUrl = SAMPLE_FIXTURES[format]
  const response = await fetch(fixtureUrl)
  const file = new File([await response.arrayBuffer()], fixtureUrl.split('/').pop() ?? `sample.${format}`, { type: 'application/octet-stream' })
  void runLoad(file, {
    ...buildOptions(),
    format
  })
})
document.querySelector('#reload')?.addEventListener('click', () => {
  if (!viewerElement) {
    ui.status.textContent = 'Viewer failed to initialize. Check console for details.'
    return
  }

  void viewerElement
    .reload()
    .then(() => {
      ui.status.textContent = `Reloaded ${viewerElement.format?.toUpperCase() ?? ''} (${viewerElement.mode ?? ''})`
      refreshSummary()
    })
    .catch((error) => {
      ui.status.textContent = `Reload failed: ${error instanceof Error ? error.message : String(error)}`
      refreshSummary()
    })
})
document.querySelector('#destroy')?.addEventListener('click', () => {
  if (!viewerElement) {
    ui.status.textContent = 'Viewer failed to initialize. Check console for details.'
    return
  }

  viewerElement.destroy()
  ui.status.textContent = 'Destroyed active viewer.'
  refreshSummary()
})

syncSampleUrl()
refreshSummary()

function createViewerElement(model: typeof ui): OfficeViewerElement | null {
  try {
    defineOfficeViewerElement()
    const candidate = document.createElement('office-viewer') as Partial<OfficeViewerElement>

    if (
      typeof candidate.load !== 'function'
      || typeof candidate.reload !== 'function'
      || typeof candidate.destroy !== 'function'
      || typeof candidate.getViewer !== 'function'
      || typeof candidate.getDocument !== 'function'
      || typeof candidate.getEngine !== 'function'
    ) {
      throw new Error('Custom element upgraded to an unexpected shape.')
    }

    const element = candidate as OfficeViewerElement
    element.style.width = '100%'
    element.style.height = '100%'
    model.viewer.replaceChildren(element)
    return element
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    model.status.textContent = `Viewer bootstrap failed: ${message}`
    model.summary.textContent = JSON.stringify(
      {
        error: {
          name: error instanceof Error ? error.name : 'Error',
          message
        }
      },
      null,
      2
    )

    console.error('office-viewer bootstrap failed', error)
    return null
  }
}
