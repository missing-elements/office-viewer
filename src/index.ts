import { defineOfficeViewerElement, OfficeViewerElement } from './office-viewer-element';
import { loadDocxAdapter } from './viewers/docx-adapter';
import { loadPptxAdapter } from './viewers/pptx-adapter';
import { loadXlsxAdapter } from './viewers/xlsx-adapter';

/**
 * @module office-viewer
 * @description Aggregated entry point for the Office Viewer library.
 * Provides access to core classes, types, and helper functions.
 * 
 * @exports {OoxmlIntegrationSpike}
 * @exports {OfficeViewerElement}
 * @exports {defineOfficeViewerElement}
 * @exports * from './viewers/adapter-types'
 * @exports * from './viewers/index'
 */
export * from './ooxml-spike';
export * from './office-viewer-element';
export * from './viewers/adapter-types';
export * from './viewers/index';

/**
 * The primary export object grouping all public APIs.
 * @module office-viewer
 */
export const officeViewer = {
  /** Core Spike Engine */
  // OoxmlIntegrationSpike is available via export * from './ooxml-spike'
  
  /** Custom Element Definition */
  OfficeViewerElement,
  defineOfficeViewerElement,
  
  
  /** Format Adapters */
  docx: {
    adapter: loadDocxAdapter,
  },
  xlsx: {
    adapter: loadXlsxAdapter,
  },
  pptx: {
    adapter: loadPptxAdapter,
  },
};

export default officeViewer;
