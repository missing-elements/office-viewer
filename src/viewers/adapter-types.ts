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
  
  // --- General Navigation/Interaction ---
  navigate(viewer: ViewerInstance, targetIndex: number): boolean
  
  // --- DOCX Specific ---
  goToPage?(pageIndex: number): boolean
  setScale?(scale: number): void
  findText?(query: string): void
  clearFind?(): void

  // --- XLSX Specific ---
  goToSheet?(sheetIndex: number): boolean
  selectRange?(rangeIdentifier: string): void
  // Add other XLSX specific methods as defined in the architecture
  
  // --- PPTX Specific ---
  goToSlide?(slideIndex: number): boolean
}
