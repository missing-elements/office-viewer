import type { OfficeFormat, ViewerInstance } from './ooxml-spike'
import {
  loadDocxAdapter,
  loadPptxAdapter,
  loadXlsxAdapter,
  navigateDocxViewer,
  navigatePptxViewer,
  navigateXlsxViewer,
  type OfficeViewerAdapter
} from './viewers'

export const VIEWER_KIND_BY_FORMAT: Record<OfficeFormat, 'DocxScrollViewer' | 'XlsxViewer' | 'PptxScrollViewer'> = {
  docx: 'DocxScrollViewer',
  xlsx: 'XlsxViewer',
  pptx: 'PptxScrollViewer'
}

export const CHANGE_EVENT_BY_FORMAT: Record<OfficeFormat, 'pagechange' | 'sheetchange' | 'slidechange'> = {
  docx: 'pagechange',
  xlsx: 'sheetchange',
  pptx: 'slidechange'
}

const ADAPTERS: Record<OfficeFormat, OfficeViewerAdapter> = {
  docx: {
    load: loadDocxAdapter,
    navigate: navigateDocxViewer
  },
  xlsx: {
    load: loadXlsxAdapter,
    navigate: navigateXlsxViewer
  },
  pptx: {
    load: loadPptxAdapter,
    navigate: navigatePptxViewer
  }
}

export function getViewerAdapter(format: OfficeFormat): OfficeViewerAdapter {
  return ADAPTERS[format]
}

export function invokeViewerMethod(viewer: ViewerInstance, methodNames: string[], targetIndex: number): boolean {
  const normalizedIndex = Math.max(0, Math.trunc(targetIndex))
  const shape = viewer as unknown as Record<string, unknown>

  for (const methodName of methodNames) {
    const candidate = shape[methodName]
    if (typeof candidate !== 'function') {
      continue
    }

    try {
      ;(candidate as (index: number) => void).call(viewer, normalizedIndex)
      return true
    } catch {
      // Try the next candidate to stay resilient across upstream API changes.
    }
  }

  return false
}
