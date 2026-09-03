import { OoxmlIntegrationSpike, SAMPLE_FIXTURES, type OfficeFormat, type OfficeSpikeLoadOptions, type OfficeSource } from '../src'

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
const harness = new OoxmlIntegrationSpike(ui.viewer)

const currentFormat = (): OfficeFormat => ui.formatSelect.value as OfficeFormat

function buildOptions(): OfficeSpikeLoadOptions {
  return {
    format: currentFormat(),
    mode: ui.modeSelect.value === 'main' ? 'main' : 'worker',
    wasmUrl: ui.wasmUrlInput.value.trim() ? ui.wasmUrlInput.value.trim() : undefined
  }
}

function refreshSummary(): void {
  ui.summary.textContent = JSON.stringify(harness.getSummary(), null, 2)
}

async function runLoad(source: OfficeSource, options = buildOptions()): Promise<void> {
  ui.status.textContent = 'Loading…'
  refreshSummary()

  try {
    const result = await harness.load(source, options)
    ui.status.textContent = `${result.format.toUpperCase()} ready in ${Math.round(result.loadCompletedMs ?? 0)}ms (${result.effectiveMode})`
    refreshSummary()
  } catch (error) {
    ui.status.textContent = `Load failed: ${error instanceof Error ? error.message : String(error)}`
    refreshSummary()
  }
}

function syncSampleUrl(): void {
  ui.urlInput.value = SAMPLE_FIXTURES[currentFormat()]
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
  void harness.reload()
    .then((result) => {
      ui.status.textContent = `Reloaded ${result.format.toUpperCase()} in ${Math.round(result.loadCompletedMs ?? 0)}ms (${result.effectiveMode})`
      refreshSummary()
    })
    .catch((error) => {
      ui.status.textContent = `Reload failed: ${error instanceof Error ? error.message : String(error)}`
      refreshSummary()
    })
})
document.querySelector('#destroy')?.addEventListener('click', () => {
  harness.destroy()
  ui.status.textContent = 'Destroyed active viewer.'
  refreshSummary()
})

syncSampleUrl()
refreshSummary()
