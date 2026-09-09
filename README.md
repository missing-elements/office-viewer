# office-viewer

Architecture spike for a standalone, read-only browser-based Office Open XML viewer built on `@silurus/ooxml`.

## Status
Milestone 2 (API Contract) is complete. The core custom element, `<office-viewer>`, now has a stable, documented public API covering source loading, lifecycle management, and format-specific navigation controls for DOCX, XLSX, and PPTX.

## Next Steps
The focus shifts from defining the API surface to solidifying the integration points:
1.  **Implementing the DOCX Adapter:** Scaffolding for `src/viewers/docx-adapter.ts` is required.
2.  **Testing:** Initial unit tests for source normalization and format detection must be created.
3.  **Initialization:** The main application entry point in `src/index.ts` must be updated to correctly initialize and use the component.

## API Summary
The public API contract is now finalized and documented in `src/viewers/adapter-types.ts`. Key additions include:
-   **Format-specific methods:** Added optional methods like `goToPage?()`, `goToSheet?()`, and `goToSlide?()` to `OfficeViewerAdapter` to expose format-specific interaction points.
-   **Core Methods:** `load()`, `reload()`, `destroy()`, `getSummary()` remain the primary control points.
-   **Event System:** Continues to dispatch granular events (`pagechange`, `sheetchange`, `slidechange`) to support advanced framework integration.

## Installation & Usage (No Change)
(Keep existing installation instructions)
...
