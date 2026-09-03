# Office Viewer Architecture

## Purpose

This document defines the architecture for `@missing-elements/office-viewer`, a read-only Web Component for rendering Office Open XML documents in the browser.

Initial formats:

- DOCX
- XLSX
- PPTX

The component uses `@silurus/ooxml` for parsing, layout, and Canvas rendering. It does not use ONLYOFFICE, `x2t`, PDF.js, or `pdfjs-viewer-element`.

Architecture decisions recorded here are authoritative as of September 3, 2026. The upstream API, declarations, package exports, and asset behavior must be rechecked when upgrading `@silurus/ooxml`.

## Goals

- Render DOCX, XLSX, and PPTX locally in the browser.
- Avoid uploading local document bytes to a service.
- Provide a stable framework-neutral custom-element API.
- Follow the TypeScript/Vite/pnpm/Vitest/WebdriverIO conventions of `missing-elements/pdfjs-viewer`.
- Use worker rendering by default.
- Keep optional renderer modules out of the default application graph.
- Make WASM and worker assets reliable in Vite, plain modules, static hosting, and verified bundler environments.
- Make load, reload, disconnect, reconnect, and destroy deterministic.
- Preserve format-specific behavior rather than forcing DOCX, XLSX, and PPTX into one model.

## Non-goals

- PDF support. PDF remains a separate component.
- Editing, mutation, saving, or lossless round-tripping.
- Collaboration or document-server protocols.
- Macro execution.
- Running embedded OLE applications.
- Legacy binary formats.
- Common printing API in the initial release.
- Reimplementing the OOXML parsers or renderers.

## Upstream engine model

```text
Office Open XML ZIP package
          |
          v
Rust format parser compiled to WebAssembly
          |
          v
validated format-specific document model
          |
          +-----------------------+
          |                       |
          v                       v
main-mode layout/paint       worker-mode layout/paint
on the main thread           in an upstream render Worker
          |                       |
          v                       v
HTML Canvas                  ImageBitmap
```

Upstream format entry points:

```text
@silurus/ooxml/docx
  DocxDocument / DocxViewer / DocxScrollViewer

@silurus/ooxml/xlsx
  XlsxWorkbook / XlsxViewer / XlsxSheetViewer

@silurus/ooxml/pptx
  PptxPresentation / PptxViewer / PptxScrollViewer
```

`office-viewer` is an orchestration layer. The upstream engine owns parsing, layout, rendering, and format-specific interaction. The custom element owns source handling, lifecycle, status UI, and adapter selection.

## Integration boundary

```text
<office-viewer>
  |
  +-- source normalization
  +-- conservative format selection
  +-- cancellation and generation control
  +-- Shadow DOM shell and status UI
  +-- common events and verified commands
  +-- resize handling
  |
  +-- DocxScrollViewer adapter
  +-- XlsxViewer adapter
  +-- PptxScrollViewer adapter
```

Do not create a second render worker. Configure the upstream viewer and use its worker lifecycle.

## Format-specific viewer choice

### DOCX

Use `DocxScrollViewer` where appropriate. DOCX is a sequence of pages and benefits from virtualization, progressive layout, text selection, find, hyperlinks, zoom, and page navigation.

### XLSX

Use `XlsxViewer`. A workbook has worksheets, cells, ranges, selection state, tabs, and grid scrolling. XLSX must not be modeled as a paginated document.

### PPTX

Use `PptxScrollViewer` where appropriate. Slides can be virtualized and displayed in a continuous reading surface while preserving slide navigation, zoom, text selection, find, hyperlinks, and verified media behavior.

All constructor options and method names must be confirmed against current upstream declarations before adapter implementation.

## Rendering mode

The public mode is:

```text
mode="worker|main"
```

- Worker mode is the default.
- `mode="worker"` prefers worker rendering.
- `mode="main"` explicitly selects main-thread rendering.
- There is no `auto` mode.
- There is no `worker-required` option initially.
- A document-specific upstream compatibility fallback may use main mode even when worker mode was requested.

The implementation should track:

```ts
requestedMode: "worker" | "main"
effectiveMode: "worker" | "main"
```

Expose the effective mode publicly only if the upstream viewer makes it reliable and stable.

### Worker requirements

Verify:

- `Worker` support;
- `OffscreenCanvas` support;
- `ImageBitmap` presentation and ownership;
- worker asset paths;
- WASM asset paths;
- CSP and static-hosting behavior;
- DOCX browser-only shaping fallback;
- custom renderer limitations in worker mode.

Every returned `ImageBitmap` must be transferred or closed. Stale bitmaps must not survive reload or destroy.

## WASM asset model

The upstream package ships parser WASM as separate assets referenced from JavaScript using `new URL(..., import.meta.url)` and loaded at runtime.

```text
format JavaScript entry
        |
        +-- parser_bg.wasm asset
        |
        v
streaming or ArrayBuffer WASM initialization
```

Do not inline parser WASM into the main JavaScript bundle unless the upstream build requires it. Separate assets provide independent caching, smaller JavaScript, clearer measurements, and CDN/static-hosting flexibility.

### `wasmUrl`

Consumers may need to host WASM at a custom location for CDNs, static asset directories, non-root deployments, offline packaging, or esbuild-based builds.

The component may expose a `wasmUrl` load option only if the current upstream declarations support it. The component must not invent a URL shape that is not supported upstream.

A provisional common type may be considered during the spike:

```ts
interface OfficeViewerLoadOptions {
  wasmUrl?: string
}
```

If the upstream engine requires a format-specific option or different type, adapters must translate internally. Do not promise one URL works for all formats without verification.

### Worker URL

Do not expose `workerUrl` in the initial public API. The upstream worker is an implementation detail unless the current package explicitly provides a supported public configuration.

## Asset verification matrix

| Environment | Main mode | Worker mode | WASM override | Required result |
|---|---:|---:|---:|---|
| Vite development | yes | yes | verify | emitted assets load |
| Vite production | yes | yes | verify | production paths work |
| Plain module | yes | yes | verify | no bundler-specific failure |
| Webpack/Next.js | verify | verify | verify | documented result |
| esbuild-based build | verify | verify | likely | copied WASM path documented if needed |
| Static subpath | yes | yes | verify | relative paths remain correct |
| CDN assets | optional | optional | yes | CORS and MIME documented |

Required checks:

- correct WASM MIME configuration for streaming;
- reachable worker assets;
- correct non-root base-path resolution;
- no development-only worker URL in production;
- no accidental inline WASM;
- optional modules absent unless explicitly imported.

## Source model and ownership

Public source types:

```ts
export type OfficeSource =
  | string
  | URL
  | File
  | Blob
  | ArrayBuffer
  | Uint8Array
```

Rules:

- A string is a URL unless explicitly documented otherwise.
- `File` and `Blob` are read locally.
- ArrayBuffer and Uint8Array values are caller-owned.
- `load()` is canonical; `src` is a convenience URL attribute.
- `reload()` uses the retained source and options.
- `file-name` supplies missing metadata.
- `file-type` is an explicit override.
- Caller buffers must not be detached or mutated unexpectedly.
- Copy caller-owned buffers before transferring them to a worker.
- Revoke only object URLs created by the component.
- Retain the original source sufficiently for `downloadOriginal()`.

A URL source is not equivalent to a local source for privacy: the browser contacts the URL origin when fetching it.

## Format detection

Use this order:

1. `file-type` attribute.
2. `format` load option.
3. `File.name`, `file-name`, or URL pathname extension.
4. Clear failure for extensionless binary input without an explicit format.

Do not treat every ZIP file as DOCX. ZIP package-part inspection is deferred beyond the MVP:

```text
word/document.xml       → DOCX
xl/workbook.xml         → XLSX
ppt/presentation.xml    → PPTX
```

Add package inspection later only if the use case justifies another archive-detection dependency.

## Lifecycle and cancellation

```text
idle → loading → ready
                 ↘ error

ready/error → loading on replacement
any temporary state → detached on disconnect
any state → destroyed on permanent destroy()
```

Every load receives:

- an `AbortController`;
- a generation token;
- a retained source/configuration record.

Async work may update the component only if its generation is current.

### Disconnect

`disconnectedCallback()` is temporary lifecycle cleanup:

- abort active work;
- destroy/release the active upstream viewer;
- remove listeners;
- release bitmaps and other render resources;
- retain source and configuration for possible reconnection.

### Reconnect

When reconnected, recreate the upstream viewer and reload the retained source when appropriate. Reconnection must not run after permanent destruction.

### Permanent destroy

`destroy()`:

- aborts active work;
- releases the upstream viewer and its worker resources;
- releases image bitmaps;
- revokes component-owned object URLs;
- clears active source/runtime state;
- requires a new explicit `load()` for future use.

## Adapter contract

The exact types must be based on current upstream declarations. The conceptual boundary is:

```ts
interface ViewerAdapter {
  readonly format: OfficeFormat
  readonly effectiveMode: "worker" | "main"

  load(source: AdapterSource, options: AdapterLoadOptions): Promise<void>
  destroy(): void

  setScale?(scale: number): void
  fitWidth?(): void
  fitPage?(): void
  findText?(query: string): void
  clearFind?(): void
}
```

Format-specific capabilities remain separate:

```text
DOCX: page navigation and visible-page callbacks
XLSX: sheet navigation, range/cell selection, workbook state
PPTX: slide navigation and visible-slide callbacks
```

Do not expose parser models as part of the stable custom-element API.

## Public API decisions

Initial attributes:

```text
src
file-name
file-type
mode="worker|main"
theme
zoom
page
slide
sheet
enable-text-selection
enable-hyperlinks
enable-download
show-toolbar
```

`locale` is excluded. `print()` is excluded from the common API.

Initial methods:

```ts
load(source: OfficeSource, options?: OfficeViewerLoadOptions): Promise<void>
reload(): Promise<void>
setScale(scale: number): void
fitWidth(): void
fitPage(): void
findText(query: string): void
clearFind(): void
downloadOriginal(): Promise<void>
destroy(): void
```

`downloadOriginal()` downloads original source bytes only. It does not create a PDF, image, or modified Office file.

## Styling

Use an open Shadow DOM, but style only the component shell where useful:

- background/desk;
- page and slide gaps;
- shadows;
- status UI;
- optional toolbar;
- XLSX shell elements where the upstream surface allows it.

Do not claim that authored document content can be themed. A Paper & Ink visual direction is optional inspiration, not an architectural dependency.

The host must provide a bounded height:

```css
office-viewer {
  display: block;
  height: 100dvh;
}
```

Required accessibility behavior:

- `aria-busy` while loading;
- accessible loading/error status;
- labelled viewer region;
- keyboard navigation where supported;
- focus-visible styling;
- clear fallback for unavailable Canvas/worker features.

## Optional renderer modules

The default graph must omit:

- MathJax;
- ChartEx;
- 3-D charts;
- Region Maps;
- TIFF decoding.

Expose optional configuration only after verifying upstream injection types and worker compatibility. Optional assets should load only when explicitly requested and needed.

## Dependency and packaging strategy

Use `@silurus/ooxml` as a normal runtime dependency initially. Do not use a peer dependency in the first release.

Start with one public entry point:

```text
@missing-elements/office-viewer
```

Potential future format-specific entries remain deferred:

```text
@missing-elements/office-viewer/docx
@missing-elements/office-viewer/xlsx
@missing-elements/office-viewer/pptx
```

Choose combined imports, dynamic imports, and subpath exports using measured application graphs and reliable asset behavior—not npm unpacked size alone.

## Privacy and security

- Local `File`, `Blob`, ArrayBuffer, and Uint8Array sources are processed locally and are not uploaded by the component.
- URL sources are fetched from the URL origin supplied by the consumer.
- Do not execute macros or embedded OLE applications.
- Preserve upstream resource limits.
- Avoid logging document contents or extracted text.
- Sanitize or constrain hyperlink behavior according to upstream APIs.
- Release workers, bitmaps, listeners, and object URLs on teardown.

## Performance measurements

Measure for each format:

- JavaScript transferred;
- WASM transferred;
- worker assets transferred;
- parser initialization time;
- time to first visible page/sheet/slide;
- complete layout time;
- peak memory;
- warm-cache load time;
- main-thread blocking time.

Preferred behavior:

- lazy-load the format adapter where reliable;
- use worker mode by default;
- use progressive layout for large DOCX/PPTX documents where supported;
- use upstream virtualization;
- omit optional modules by default;
- avoid parsing a document more than once when future multi-view APIs are introduced.

## Testing strategy

Use Vitest and WebdriverIO as the primary test stack, matching `missing-elements/pdfjs-viewer`.

### Unit tests

- source normalization;
- extension inference;
- format detection;
- attribute parsing;
- worker capability selection;
- stale-load and abort handling;
- buffer ownership;
- error mapping;
- event payloads;
- cleanup.

### Browser tests

- URL, File, Blob, ArrayBuffer, and Uint8Array loading;
- DOCX, XLSX, and PPTX;
- worker and main modes;
- verified fallback;
- reload and replacement;
- disconnect/reconnect;
- permanent destroy;
- resize/refit;
- zoom and format-specific navigation;
- find and hyperlinks;
- original download;
- malformed, encrypted, legacy, and unsupported files.

### Asset tests

- emitted WASM and worker files;
- production paths;
- static subpath deployment;
- `wasmUrl` override where supported;
- no inline WASM;
- no development-only paths;
- optional modules not fetched by default.

### Visual fixtures

```text
simple.docx
tables.docx
images.docx
equations.docx
simple.xlsx
formatted.xlsx
charts.xlsx
simple.pptx
shapes.pptx
images.pptx
```

## Architecture-spike exit criteria

The spike is complete when:

- all three upstream viewers load real fixtures;
- worker mode has been tested in a production build;
- main-mode behavior and any fallback are known;
- WASM and worker assets resolve in supported environments;
- `wasmUrl` behavior is documented if exposed;
- format-specific bundle/runtime sizes are recorded;
- optional modules are confirmed absent by default;
- source ownership and destroy/reload behavior are tested;
- the adapter contract is based on actual declarations;
- known upstream pitfalls are recorded in `docs/ooxml-integration-notes.md`.

Only then should the stable public API of `<office-viewer>` be finalized.
