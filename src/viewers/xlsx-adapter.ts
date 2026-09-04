import type { RetainedLoad } from '../ooxml-spike'
import type { AdapterLoadHooks, LoadedAdapter } from './adapter-types'
import { invokeViewerMethod } from '../viewer-adapter'

export async function loadXlsxAdapter(retainedLoad: RetainedLoad, container: HTMLElement, hooks: AdapterLoadHooks): Promise<LoadedAdapter> {
  const { XlsxViewer, XlsxWorkbook } = await import('@silurus/ooxml/xlsx')
  const engine = await XlsxWorkbook.load(cloneRetainedSource(retainedLoad), {
    mode: retainedLoad.options.mode ?? 'worker',
    wasmUrl: retainedLoad.options.wasmUrl
  })

  const viewer = XlsxViewer.fromWorkbook(container, engine, {
    enableHyperlinks: true,
    onReady: (sheetNames) => {
      hooks.onSummaryUpdate({
        currentIndex: 0,
        totalCount: sheetNames.length,
        layoutComplete: true,
        sheetNames: [...sheetNames]
      })
    },
    onSheetChange: (index, total) => {
      hooks.onSummaryUpdate({
        currentIndex: index,
        totalCount: total,
        layoutComplete: true
      })
    },
    onError: (error) => {
      hooks.onError(error)
    }
  })

  hooks.onSummaryUpdate({
    currentIndex: viewer.sheetIndex,
    totalCount: viewer.sheetCount,
    layoutComplete: true,
    sheetNames: [...viewer.sheetNames]
  })

  return {
    engine,
    viewer
  }
}

export function navigateXlsxViewer(viewer: { destroy(): void }, targetIndex: number): boolean {
  return invokeViewerMethod(viewer, ['setSheetIndex', 'goToSheet', 'showSheet', 'setActiveSheetIndex'], targetIndex)
}

function cloneRetainedSource(retainedLoad: RetainedLoad): string | ArrayBuffer {
  return typeof retainedLoad.source === 'string' ? retainedLoad.source : retainedLoad.source.slice(0)
}
