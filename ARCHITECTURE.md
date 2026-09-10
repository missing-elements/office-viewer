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

The format-specific modules are loaded dynamically, so only the requested
viewer and its WASM assets are fetched. The element accepts URL strings and
`ArrayBuffer` sources directly; it resolves `Blob`/`File` and
`ReadableStream<Uint8Array>` sources to an `ArrayBuffer` before calling
`viewer.load(source)`.

## Lifecycle

- `connectedCallback`: create Shadow DOM container; auto-load if `src` attribute present.
- `load`: cancel a previous load, create a viewer, normalize the source, then
	call `viewer.load(source)`. Superseded loads destroy their own viewer and do
	not update the element's state.
- `reload`: re-run `load` with retained source and options.
- `destroy`: cancel load, destroy viewer, reset state.

`loadstart` is emitted when loading begins. A successful load emits `ready`; a
failed load stores the error and emits `loaderror` with `{ error }`. `destroy`
is emitted after teardown.

## Public surface

Only `load`, `reload`, `destroy`, `getViewer`, and read-only `ready`, `error`,
`format`, and `mode` state properties are exposed. Consumers interact with the
upstream viewer returned by `getViewer()`.
