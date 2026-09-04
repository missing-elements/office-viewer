import type { RetainedLoad } from '../ooxml-spike'
import type { AdapterLoadHooks, LoadedAdapter } from './adapter-types'
import { invokeViewerMethod } from '../viewer-adapter'

export async function loadDocxAdapter(retainedLoad: RetainedLoad, container: HTMLElement, hooks: AdapterLoadHooks): Promise<LoadedAdapter> {
  const { DocxDocument, DocxScrollViewer } = await import('@silurus/ooxml/docx')
  const engine = await DocxDocument.load(cloneRetainedSource(retainedLoad), {
    mode: retainedLoad.options.mode ?? 'worker',
    wasmUrl: retainedLoad.options.wasmUrl,
    progressiveLayout: true
  })

  const viewer = DocxScrollViewer.fromDocument(container, engine, {
    enableHyperlinks: true,
    enableTextSelection: true,
    onVisiblePageChange: (topIndex, total, layoutComplete) => {
      hooks.onSummaryUpdate({
        currentIndex: topIndex,
        totalCount: total,
        layoutComplete
      })
    },
    onError: (error) => {
      hooks.onError(error)
    }
  })

  hooks.onSummaryUpdate({
    currentIndex: viewer.topVisiblePage,
    totalCount: viewer.pageCount,
    layoutComplete: viewer.layoutComplete
  })

  return {
    engine,
    viewer
  }
}

export function navigateDocxViewer(viewer: { destroy(): void }, targetIndex: number): boolean {
  return invokeViewerMethod(viewer, ['scrollToPage', 'goToPage', 'setPageIndex', 'setVisiblePageIndex'], targetIndex)
}

function cloneRetainedSource(retainedLoad: RetainedLoad): string | ArrayBuffer {
  return typeof retainedLoad.source === 'string' ? retainedLoad.source : retainedLoad.source.slice(0)
}
