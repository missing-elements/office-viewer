# Office Viewer Architecture

## Goals

- Provide a headless custom element that loads DOCX, XLSX, and PPTX via `@silurus/ooxml`.
- Expose upstream viewer instances directly.
- Shield the viewer surface from global styles with an open Shadow DOM.
- Keep the orchestration layer small.

## Non-goals

- PDF support, editing, saving, printing, UI chrome.
- Wrapping upstream APIs with new method names.

## Component

`<office-viewer>` creates an open Shadow DOM on connection. Inside, it mounts a `<div id="viewer">` with `width: 100%; height: 100%; overflow: auto;`.

On `load(source, options)` it constructs the appropriate upstream viewer:

```ts
new DocxScrollViewer(container, { mode, wasmUrl })
new XlsxViewer(container, { mode, wasmUrl })
new PptxScrollViewer(container, { mode, wasmUrl })
```

Then calls `viewer.load(source)`.

## Lifecycle

- `connectedCallback`: create Shadow DOM container; auto-load if `src` attribute present.
- `load`: cancel previous load, create viewer, call `viewer.load(source)`.
- `reload`: re-run `load` with retained source and options.
- `destroy`: cancel load, destroy viewer, reset state.

## Public surface

Only `load`, `reload`, `destroy`, `getViewer`, and read-only state properties. Consumers interact with the upstream viewer returned by `getViewer()`.
