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

- `connectedCallback`: create the Shadow DOM container and schedule a sync.
- `disconnectedCallback`: after one microtask, if the element is still
  disconnected and holds a viewer or an in-flight load, release them, emit
  `destroy`, keep the retained request, and mark the element for restore. A
  synchronous move (remove, then append) therefore keeps the viewer, and
  `destroy()` followed by `remove()` emits a single `destroy`.
- `attributeChangedCallback`: mark the attributes dirty and schedule a sync.
- Sync runs at most once per microtask and only while connected. If attributes
  are dirty and `src` is set, it starts an attribute-driven load; if attributes
  are dirty, `src` is empty, and the current document came from attributes, it
  calls `destroy()`; otherwise, if the element was marked for restore, it calls
  `reload()`. Moves without attribute changes and failed loads never start work.
- `load`: cancel the in-flight load, clear the dirty flag (an explicit call
  outranks earlier attribute changes in the same task), retain the request, clear
  `error`, emit `loadstart`, validate the format, resolve the source and import
  the format module concurrently, construct the upstream viewer, call
  `viewer.load()`, then commit: swap in the new viewer, record its options,
  destroy the previous viewer, and emit `ready`.
- `reload`: re-run `load` with the retained request, keeping whether it was
  attribute-driven.
- `destroy`: cancel the in-flight load, destroy the viewer, forget the retained
  request, reset state, emit `destroy`.

### Load ownership

Each `load()` owns an `AbortController`. Cancelling it, whether by a newer
`load()`, `destroy()`, or removal from the document, aborts the signal and
destroys whatever the load has constructed so far:

- a pending `ReadableStream` read is cancelled;
- no viewer is constructed once the source and module have arrived;
- the upstream viewer is destroyed immediately.

Upstream settles its own `load()` only after the fetch and parse finish, even
when the viewer has been destroyed, so every await in `load()` is raced against
the signal. The cancelled promise therefore rejects at once with the abort
reason, a `DOMException` named `AbortError`, and never touches element state or
emits events. Only the owning load stores `error`, emits `loaderror`, or commits
a viewer.

### Retained request

`reload()` replays the most recent request, including one that failed. Options
are copied and resolved (default mode applied) once per request. Upstream may
take ownership of the `ArrayBuffer` it is given and detach it, so the element
never retains those bytes: URL strings and `Blob`/`File` sources are retained as-is,
and `ArrayBuffer` and stream sources are retained as a `Blob` copy. `destroy()`
clears the retained request; removal from the document does not.

### State properties

`ready`, `getViewer()`, `format`, and `mode` all describe the committed viewer
and change together when `ready` fires. `mode` is the mode the viewer was
created with, `worker` when the caller gave none; upstream may still fall back
to main-thread rendering internally. `error` describes the most recent owning
load and is cleared at `loadstart` and on `destroy()`.

## Public surface

Only `load`, `reload`, `destroy`, `getViewer`, and read-only `ready`, `error`,
`format`, and `mode` state properties are exposed. Consumers interact with the
upstream viewer returned by `getViewer()`.
