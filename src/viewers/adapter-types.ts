import type { EngineInstance, OfficeSpikeSummaryUpdate, RetainedLoad, ViewerInstance } from '../ooxml-spike'

export interface AdapterLoadHooks {
  onSummaryUpdate(update: OfficeSpikeSummaryUpdate): void
  onError(error: unknown): void
}

export interface LoadedAdapter {
  engine: EngineInstance
  viewer: ViewerInstance
}

export interface OfficeViewerAdapter {
  load(retainedLoad: RetainedLoad, container: HTMLElement, hooks: AdapterLoadHooks): Promise<LoadedAdapter>
  navigate(viewer: ViewerInstance, targetIndex: number): boolean
}
