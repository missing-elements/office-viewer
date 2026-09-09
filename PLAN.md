# Office Viewer Plan

Headless Web Component orchestration layer for `@silurus/ooxml`.

## Scope

- Formats: DOCX, XLSX, PPTX.
- No UI chrome, no wrapper methods, no document/engine abstraction.
- Shadow DOM is used only as an internal style boundary for the viewer surface.

## Public API

```ts
load(source: string | ArrayBuffer, options: { format, mode?, wasmUrl? })
reload()
destroy()
getViewer() -> DocxScrollViewer | XlsxViewer | PptxScrollViewer | null

ready: boolean
error: Error | null
format: OfficeFormat | null
mode: OfficeViewerMode | null
```

## Architecture

```text
<office-viewer> (open Shadow DOM)
  |
  +-- Shadow DOM container #viewer
  |
  +-- new DocxScrollViewer(container, opts)
  +-- new XlsxViewer(container, opts)
  +-- new PptxScrollViewer(container, opts)
```

The element owns:
- attribute parsing
- Shadow DOM container
- viewer lifecycle (load, cancel, reload, destroy)
- simple status/error properties

The element does not own:
- document models (they are internal to upstream viewers)
- wrapper methods
- UI chrome

## Testing

- Unit: jsdom-based custom element lifecycle tests.
- Browser: load each format from URL and ArrayBuffer, verify `getViewer()` type.

## Milestones

- [x] Simplified headless component with Shadow DOM.
- [ ] Browser integration tests against real fixtures.
- [ ] Documentation and release.
