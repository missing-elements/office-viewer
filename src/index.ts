import { defineOfficeViewerElement, OfficeViewerElement } from './office-viewer-element'
import { createDocxAdapter } from './viewers/docx-adapter'
import { createPptxAdapter } from './viewers/pptx-adapter'
import { createXlsxAdapter } from './viewers/xlsx-adapter'

export * from './types'
export * from './source-resolver'
export * from './load-controller'
export * from './office-viewer-element'
export * from './viewers/adapter-types'
export * from './viewers/index'

export const officeViewer = {
  OfficeViewerElement,
  defineOfficeViewerElement,

  docx: {
    adapter: createDocxAdapter
  },
  xlsx: {
    adapter: createXlsxAdapter
  },
  pptx: {
    adapter: createPptxAdapter
  }
}

export default officeViewer
