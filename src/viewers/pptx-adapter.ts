import type { RetainedLoad } from '../ooxml-spike'
import type { AdapterLoadHooks, LoadedAdapter } from './adapter-types'
import { invokeViewerMethod } from '../viewer-adapter'

export async function loadPptxAdapter(retainedLoad: RetainedLoad, container: HTMLElement, hooks: AdapterLoadHooks): Promise<LoadedAdapter> {
  const { PptxPresentation, PptxScrollViewer } = await import('@silurus/ooxml/pptx')
  const engine = await PptxPresentation.load(cloneRetainedSource(retainedLoad), {
    mode: retainedLoad.options.mode ?? 'worker',
    wasmUrl: retainedLoad.options.wasmUrl,
    progressiveLayout: true
  })

  const viewer = PptxScrollViewer.fromPresentation(container, engine, {
    enableHyperlinks: true,
    enableTextSelection: true,
    onVisibleSlideChange: (topIndex, total, layoutComplete) => {
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
    currentIndex: viewer.topVisibleSlide,
    totalCount: viewer.slideCount,
    layoutComplete: viewer.layoutComplete
  })

  return {
    engine,
    viewer
  }
}

export function navigatePptxViewer(viewer: { destroy(): void }, targetIndex: number): boolean {
  return invokeViewerMethod(viewer, ['scrollToSlide', 'goToSlide', 'setSlideIndex', 'setVisibleSlideIndex'], targetIndex)
}

function cloneRetainedSource(retainedLoad: RetainedLoad): string | ArrayBuffer {
  return typeof retainedLoad.source === 'string' ? retainedLoad.source : retainedLoad.source.slice(0)
}
