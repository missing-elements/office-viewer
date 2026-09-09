# Architectural Conclusions

## Format-Specific API Contracts
The primary architectural decision solidified in this phase is the necessary expansion of the `OfficeViewerAdapter` interface. To support format-specific interactions (e.g., page navigation for DOCX, sheet/range selection for XLSX, slide navigation for PPTX), the adapter must evolve from a generic interface to one that encapsulates optional, format-specific methods. This adheres to the principle of "only expose what is needed," avoiding a monolithic interface.

**Key Contracts:**
-   **DOCX:** Must expose `goToPage(pageIndex: number): boolean` to handle page-based scrolling and navigation.
-   **XLSX:** Must expose `goToSheet?(sheetIndex: number): boolean` and `selectRange?(rangeIdentifier: string): void` to manage the grid-based state.
-   **PPTX:** Must expose `goToSlide?(slideIndex: number): boolean` for slide-based navigation.

## Component Structure (`src/office-viewer-element.ts`)
The element contract is stable:
-   It remains the orchestrator, handling state, lifecycle, and event dispatching.
-   It correctly delegates loading/reloading to the `OoxmlIntegrationSpike` harness.
-   The attribute parsing and lifecycle hooks (`connectedCallback`, `disconnectedCallback`, `attributeChangedCallback`) provide the necessary robust integration points for external use.

## Next Implementation Focus
The remaining tasks are focused on implementation depth rather than design refinement:
1.  **Implementation Scaffolding:** Completing the concrete implementations in the adapter files (`docx-adapter.ts`, etc.).
2.  **Testing:** Writing unit and integration tests to validate the new API surface and the core logic (source normalization, format detection).
3.  **Bundle:** Finalizing the entry point (`src/index.ts`) and ensuring clean exports.
