# Office Viewer Architecture

## Purpose

This document defines the architecture for `@missing-elements/office-viewer`, a read-only Web Component for rendering Office Open XML documents in the browser.

The initial supported formats are:

- DOCX
- XLSX
- PPTX

The component uses `@silurus/ooxml` for parsing, layout, and Canvas rendering. It does not use ONLYOFFICE, `x2t`, PDF.js, or `pdfjs-viewer-element`.

> Architecture study date: September 3, 2026.
>
> The public `@silurus/ooxml` API, generated declarations, package exports, and asset behavior must be rechecked when upgrading the dependency. This document records integration decisions, not a substitute for the upstream API reference.

## Goals

- Render DOCX, XLSX, and PPTX privately in the browser.
- Avoid uploading documents to a server.
- Provide a stable, framework-neutral custom-element API.
- Keep the implementation consistent with `missing-elements/pdfjs-viewer`.
- Prefer off-main-thread parsing and rendering where supported.
- Keep optional renderer modules out of the default application graph.
- Make WASM and worker asset loading work with Vite, plain browser modules, static hosting, and common framework bundlers.
- Make load, reload, disconnect, reconnect, and destroy behavior deterministic.
- Preserve the distinction between document formats instead of forcing them into one artificial model.

## Non-goals

- Editing or mutation APIs.
- Lossless document round-tripping.
- Saving modified DOCX, XLSX, or PPTX files.
- Collaboration or server-side document sessions.
- Macro execution.
- Running embedded OLE applications.
- Legacy binary formats such as DOC, XLS, and PPT.
- PDF support. PDF remains a separate project: `@missing-elements/pdfjs-viewer-element`.
- Reimplementing the OOXML parsers or renderers.

## Upstream engine model

`@silurus/ooxml` is a read-only rendering library. Its architecture is approximately:

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
on the main thread           in a render Worker
          |                       |
          v                       v
HTML Canvas                  ImageBitmap
```

The upstream package has separate engines for each format:

```text
@silurus/ooxml/docx
  DocxDocument / DocxViewer / DocxScrollViewer

@silurus/ooxml/xlsx
  XlsxWorkbook / XlsxViewer / XlsxSheetViewer

@silurus/ooxml/pptx
  PptxPresentation / PptxViewer / PptxScrollViewer
```

The package also has shared layout and paint primitives in `@silurus/ooxml-core` and optional renderer entries such as MathJax, ChartEx, 3-D charts, Region Maps, and TIFF decoding.

## Integration boundary

`office-viewer` is an orchestration and presentation layer. It should not duplicate the upstream document model.

```text
<office-viewer>
  |
  +-- source normalization
  +-- format selection
  +-- lifecycle and cancellation
  +-- Shadow DOM and status UI
  +-- common events
  +-- common navigation commands
  |
  +-- DocxScrollViewer
  +-- XlsxViewer
  +-- PptxScrollViewer
```

The element owns the viewer instance and its host container. The upstream viewer owns parsing, layout, rendering, and format-specific interaction.

## Repository structure

The implementation should follow the `pdfjs-viewer` project style:

```text
.
├── demo/
├── docs/
├── scripts/
├── src/
│   ├── index.ts
│   ├── office-viewer-element.ts
│   ├── types.ts
│   ├── source-resolver.ts
│   ├── format-detector.ts
│   ├── load-controller.ts
│   ├── viewer-adapter.ts
│   ├── viewers/
│   │   ├── docx-adapter.ts
│   │   ├── xlsx-adapter.ts
│   │   └── pptx-adapter.ts
│   ├── ui/
│   │   ├── loading-view.ts
│   │   ├── error-view.ts
│   │   └── toolbar.ts
│   └── themes/
│       └── paper-and-ink.css
├── tests/
├── types/
├── package.json
├── tsconfig.json
├── vite.config.ts
└── PLAN.md
```

## Viewer selection

### DOCX

Use `DocxScrollViewer` for the default component surface. DOCX is naturally represented as a sequence of pages, and the scroll viewer provides virtualization, zoom, page visibility callbacks, and progressive layout support.

Recommended options to investigate and use where available:

```ts
{
  mode: "worker",
  progressiveLayout: true,
  enableTextSelection: true,
  enableHyperlinks: true,
}
```

Do not assume every option has identical constructor and load signatures. Confirm the declarations before implementation.

### XLSX

Use `XlsxViewer`, not a document-page scroll abstraction. A workbook has worksheets, cells, ranges, selection state, and sheet tabs.

The adapter should preserve spreadsheet-specific operations rather than hiding them behind page terminology:

```text
sheet selection
cell/range selection
sheet navigation
viewport scrolling
zoom
find
hyperlinks
```

### PPTX

Use `PptxScrollViewer` for the default continuous presentation surface. It provides a natural slide-by-slide reading experience while retaining slide-specific navigation and media behavior.

Recommended options to investigate and use where available:

```ts
{
  mode: "worker",
  progressiveLayout: true,
  enableTextSelection: true,
  enableHyperlinks: true,
}
```

## Main mode and worker mode

The upstream library exposes two important execution models.

### Main mode

```text
parser Worker
  |
  v
validated model
  |
  v
main-thread layout and Canvas paint
```

This mode is the compatibility fallback. It is useful when `Worker` or `OffscreenCanvas` is unavailable, or when a particular document requires browser-only shaping behavior.

### Worker mode

```text
render Worker
  |
  +-- parse
  +-- layout
  +-- paint using OffscreenCanvas
  |
  v
ImageBitmap
  |
  v
main-thread presentation
```

Worker mode should be preferred for large documents because it keeps expensive parsing, layout, and painting away from the main thread.

Important integration rules:

1. Detect support before choosing worker mode.
2. Treat the requested mode and effective mode as separate values.
3. Do not use canvas-target APIs when the underlying viewer is in worker mode if the upstream API marks them unavailable.
4. Ensure every returned `ImageBitmap` is either transferred to the presentation canvas or closed by the owner.
5. Do not retain stale bitmaps after a reload or destroy operation.
6. Expect a DOCX shaping fallback to main mode in cases documented by the upstream package.
7. Do not pass main-realm custom renderer objects into a worker unless the upstream API explicitly supports them.

Suggested mode type:

```ts
export type RequestedRenderMode = "main" | "worker" | "auto"

export type EffectiveRenderMode = "main" | "worker"
```

Recommended default:

```text
requested mode: auto
preferred mode: worker
fallback: main
```

The `mode="worker"` attribute may mean “prefer worker mode” in the MVP. If strict behavior is needed later, add a separate `worker-required` option rather than silently changing the meaning of `mode`.

## WASM asset model

Each OOXML format has a parser WASM asset. The package documentation describes these as real `.wasm` files referenced beside JavaScript using `new URL(..., import.meta.url)` and fetched at runtime.

The architecture must preserve that separation:

```text
JavaScript module
     |
     +-- new URL("...parser_bg.wasm", import.meta.url)
     |
     v
WASM file served as a separate asset
```

Do not inline the parser WASM into the main JavaScript bundle unless the upstream package or a future build explicitly requires it. External assets provide:

- independent browser caching;
- smaller initial JavaScript;
- clearer network behavior;
- CDN/static-hosting flexibility;
- easier bundle-size measurement.

### Asset resolution requirements

The integration must test:

- Vite development server;
- Vite production build;
- plain `<script type="module">` usage;
- Next.js or other webpack-based application;
- static hosting under a non-root base path;
- CDN-hosted WASM using an explicit `wasmUrl` option;
- worker mode where the worker and WASM are separate assets.

The upstream documentation identifies two important compatibility areas:

- Vite 7 development may require excluding `@silurus/ooxml` from dependency optimization.
- esbuild-based builders may not process `new URL` WASM references and may require copying the WASM asset and supplying `wasmUrl`.

These should become documented integration cases, not assumptions hidden in the component.

### `wasmUrl` policy

The component should not rewrite an upstream `wasmUrl` unless necessary. Consumers may need to control the parser asset location for:

- a CDN;
- a static asset directory;
- an Angular/esbuild build;
- a custom deployment base path;
- an offline bundle.

The initial API can expose a load option:

```ts
interface OfficeViewerLoadOptions {
  wasmUrl?: string | Partial<Record<OfficeFormat, string>>
}
```

However, the actual shape must match the upstream viewer constructors/load options. If `@silurus/ooxml` accepts one format-specific URL rather than a map, the adapter should translate the common option internally.

Do not promise that one URL works for all three formats unless the upstream library confirms that behavior.

## Worker asset model

There are two different workers that must not be confused:

```text
Parser worker / render worker
  supplied or created by @silurus/ooxml

Office Viewer custom-element code
  orchestration on the main thread
```

`office-viewer` should not create a second rendering worker unless the upstream API requires it. The adapter should configure the upstream viewer and let it manage its worker lifecycle.

If the bundler requires a worker URL or a copied worker asset, the package must expose a controlled configuration rather than relying on an undocumented path.

Potential configuration shape:

```ts
interface WorkerAssetOptions {
  workerUrl?: string
  wasmUrl?: string
}
```

This is provisional and must be validated against the actual package declarations.

### Worker lifecycle

For every load:

```text
create or acquire upstream viewer
  |
  v
start parse/render
  |
  v
present output
  |
  v
destroy viewer on replacement or element teardown
```

On failure or cancellation:

- reject the current load promise;
- terminate or destroy the upstream viewer if its API requires it;
- release image bitmaps;
- remove event listeners;
- prevent late callbacks from updating the new viewer;
- leave the component in a predictable error or empty state.

## Source model

The public component should accept:

```ts
export type OfficeSource =
  | string
  | URL
  | File
  | Blob
  | ArrayBuffer
  | Uint8Array
```

Interpretation:

- `string` is a URL unless explicitly documented otherwise.
- `URL` is fetched as a URL.
- `File` and `Blob` are read locally.
- `ArrayBuffer` and `Uint8Array` are caller-provided bytes.

The resolver returns an internal object:

```ts
interface ResolvedSource {
  bytes?: ArrayBuffer | Uint8Array
  url?: string
  fileName?: string
  format?: OfficeFormat
  original?: OfficeSource
}
```

The exact source shape passed to `@silurus/ooxml` must be selected according to its documented API. Do not convert every source into a Blob URL by default; direct byte loading may avoid unnecessary copying.

### Source ownership

The component must document ownership clearly:

- Caller-owned `ArrayBuffer` and `Uint8Array` must not be unexpectedly detached or mutated.
- If transferables are used internally, copy the data first or explicitly document transfer ownership.
- Object URLs created by the component must be revoked.
- A caller-provided URL must not be revoked by the component.
- The original source should remain available for the optional “download original” action.

## Format detection

```ts
export type OfficeFormat = "docx" | "xlsx" | "pptx"
```

Resolution order:

1. `file-type` attribute.
2. `format` supplied to `load()`.
3. Filename extension from `File`, URL, or `file-name`.
4. Container signature and package-part inspection.
5. Clear unsupported-format error.

The component should not treat every ZIP file as DOCX. OOXML packages are ZIP containers, but the package parts distinguish Word, Excel, and PowerPoint documents.

Initial detection should be conservative:

```text
word/document.xml       → docx
xl/workbook.xml         → xlsx
ppt/presentation.xml    → pptx
```

If package inspection is not part of the first milestone, require an explicit filename or `file-type` for extensionless binary sources and document that behavior.

## Lifecycle state machine

The custom element should model its state explicitly:

```text
idle
  |
  v
loading
  |
  +--> ready
  |
  +--> error
  |
  +--> idle (aborted/replaced)

ready --new load--> loading
ready --destroy--> destroyed
error --new load--> loading
idle --destroy--> destroyed
```

Suggested internal state:

```ts
type ViewerStatus =
  | "idle"
  | "loading"
  | "ready"
  | "error"
  | "destroyed"
```

Every load receives a generation number:

```ts
const generation = ++this.loadGeneration
```

Async work may update the element only if its generation is still current.

## Attribute behavior

Attributes should be observed only when their behavior is well-defined.

Initial attributes:

```text
src
file-name
file-type
mode
theme
locale
zoom
page
slide
sheet
enable-text-selection
enable-hyperlinks
enable-download
enable-print
show-toolbar
```

Rules:

- Changing `src` starts a new load.
- Changing `file-type` may trigger a reload only if a source is already loaded.
- Changing `mode` should affect the next load; live mode switching should not be promised initially.
- Changing theme may be applied live if the upstream viewer supports it; otherwise apply on reload.
- Changing page/slide/sheet should call the format-specific navigation API after readiness.
- Invalid numeric values should produce a controlled error or be ignored with a diagnostic; never pass `NaN` into the upstream viewer.

## Adapter design

Use a common adapter lifecycle but preserve format-specific capabilities.

```ts
interface ViewerAdapter {
  readonly format: OfficeFormat
  readonly effectiveMode: EffectiveRenderMode

  load(source: ResolvedSource, options: AdapterLoadOptions): Promise<void>
  destroy(): void

  setScale?(scale: number): void
  fitWidth?(): void
  fitPage?(): void
  findText?(query: string): void
  clearFind?(): void
  print?(): Promise<void>
}
```

Format-specific extensions should be typed separately:

```ts
interface DocxViewerAdapter extends ViewerAdapter {
  goToPage(pageIndex: number): Promise<void>
}

interface XlsxViewerAdapter extends ViewerAdapter {
  goToSheet(sheetIndex: number): Promise<void>
  setSelection?(selection: unknown): void
}

interface PptxViewerAdapter extends ViewerAdapter {
  goToSlide(slideIndex: number): Promise<void>
}
```

Do not expose `unknown` in the final public API where upstream declarations provide a stable type. The placeholder only indicates that XLSX selection requires a format-specific contract.

## Optional rendering modules

The upstream package treats several renderers as opt-in:

- MathJax equations.
- ChartEx.
- 3-D charts.
- Region Maps.
- TIFF decoding.

The MVP should omit these modules by default. This keeps them outside the ordinary application graph and avoids fetching optional assets for documents that do not need them.

Future configuration should be explicit:

```ts
interface OptionalRenderers {
  math?: unknown
  chartEx?: unknown
  threeD?: unknown
  regionMap?: unknown
  tiff?: unknown
}
```

The final types must use the actual upstream renderer types. Optional modules must be passed at the documented construction/load boundary and not injected into every render call.

Worker-mode compatibility must be tested for each enabled module. A module that works in main mode but not worker mode must produce a documented fallback or error.

## Text selection, hyperlinks, and find

Canvas does not provide native text semantics by itself. The upstream project uses overlay techniques for selection and interaction.

The component should enable these features only through documented upstream options:

```text
enableTextSelection
findText / findNext / findPrev / clearFind
enableHyperlinks
onHyperlinkClick
```

Do not build a second text layer in the first version. Duplicating upstream text geometry would create synchronization and accessibility problems.

The adapter should translate upstream callbacks into component events without leaking the entire internal document model.

## Resize and layout

The host element must have a bounded height:

```css
office-viewer {
  display: block;
  height: 100dvh;
}
```

The component should use `ResizeObserver` where the upstream viewer supports explicit relayout/refit operations.

Resize policy:

1. Observe the viewer container.
2. Coalesce rapid resize notifications.
3. Call the format-specific relayout/refit API.
4. Avoid forcing a full document reload.
5. Ignore resize notifications after destroy.

The component should not assume that a zero-width container is permanently invalid. If the upstream engine defers layout for zero-width containers, allow the first non-zero resize to trigger layout.

## Error handling

The upstream package documents stable error categories including:

- encrypted document;
- invalid password;
- unsupported encryption;
- legacy binary format;
- not OOXML;
- resource-limit failures;
- decoded-image limits;
- parser crashes;
- ordinary fetch, worker, render, and configuration errors.

The adapter should preserve stable error codes where available and wrap them in a component-level error shape:

```ts
interface OfficeViewerErrorDetail {
  error: Error
  code?: string
  format?: OfficeFormat
  phase: "resolve" | "detect" | "load" | "render" | "destroy"
  recoverable: boolean
}
```

Do not classify errors by parsing human-readable error messages when a stable upstream code or class exists.

## Shadow DOM and UI

Use an open Shadow DOM, matching the style of `pdfjs-viewer-element`.

```text
<office-viewer>
  #shadow-root
    <style>
    <div part="root">
      <div part="status" aria-live="polite"></div>
      <div part="viewer"></div>
    </div>
</office-viewer>
```

Recommended styling hooks:

```css
--office-viewer-background
--office-viewer-page-gap
--office-viewer-page-shadow
--office-viewer-accent-color
--office-viewer-status-color
```

The upstream viewer owns its Canvas and internal rendering surface. The component owns only the outer shell and status UI unless the upstream API requires a specific target element.

## Proposed package exports

Start with one public entry point:

```text
@missing-elements/office-viewer
```

Possible future entries:

```text
@missing-elements/office-viewer/docx
@missing-elements/office-viewer/xlsx
@missing-elements/office-viewer/pptx
```

Do not publish format-specific entries until bundle measurements and worker/WASM asset behavior justify them.

The combined custom element may use dynamic imports internally:

```text
<office-viewer src="report.docx">
  |
  +-- detect DOCX
  +-- dynamically import DOCX adapter
  +-- load DOCX WASM
```

Dynamic imports are desirable for first-load size but must be tested carefully with Vite, worker assets, and plain module usage. If dynamic imports make asset resolution unreliable, prefer stable format-specific builds over a fragile automatic loader.

## Build and asset verification matrix

Before finalizing the package build, verify:

| Environment | Main mode | Worker mode | WASM URL override | Notes |
|---|---:|---:|---:|---|
| Vite dev | required | required | required | Include Vite 7 and current Vite behavior if supported |
| Vite production | required | required | required | Check emitted asset paths |
| Plain browser module | required | required | required | No bundler |
| Webpack 5 / Next.js | required | required | required | Check `new URL` handling |
| esbuild-based build | required | required | required | Copy WASM and use `wasmUrl` if necessary |
| Static host subpath | required | required | required | Test non-root deployment |
| CDN assets | optional | optional | required | Verify CORS and MIME configuration |

Required checks:

- WASM is served with a correct MIME type where streaming compilation is used.
- Worker files are reachable from the deployed URL.
- Relative asset paths resolve under a non-root base path.
- No parser WASM is accidentally inlined into the main bundle.
- No worker URL points to a development-only source file in production.
- Optional modules are absent from the default graph unless explicitly imported.

## Performance strategy

Measure real browser cost rather than npm unpacked size.

Record for each format:

- JavaScript transferred.
- WASM transferred.
- Worker assets transferred.
- Time to parser initialization.
- Time to first visible page/sheet/slide.
- Time to complete layout.
- Peak memory.
- Repeat-load cost with a warm cache.
- Main-thread blocking time.

Recommended defaults:

- Lazy-load the format adapter.
- Prefer worker mode.
- Use progressive layout for long DOCX/PPTX documents where supported.
- Virtualize continuous scroll surfaces.
- Avoid optional modules by default.
- Do not parse a document more than once when multiple views are added later.

A future advanced API may allow a caller to load a headless engine once and create multiple views from it. This should not be part of the first custom-element contract unless the use case is demonstrated.

## Security and privacy

The component processes document bytes locally, but browser security still applies.

Required practices:

- Do not upload document bytes from the component.
- Use normal browser fetch behavior for URL sources.
- Document CORS requirements for cross-origin URLs.
- Do not execute macros or embedded OLE applications.
- Sanitize or constrain hyperlink navigation according to upstream behavior.
- Keep resource limits enabled and configurable only through documented upstream options.
- Avoid logging document contents or extracted text.
- Revoke object URLs and release worker resources on destroy.

Local processing is a product property, not a guarantee that every URL source is private; a URL fetch still contacts the URL origin.

## Testing strategy

### Unit tests

- Source normalization.
- Filename inference.
- Extension and package-part detection.
- Attribute parsing.
- Mode selection.
- Worker capability detection.
- Stale-load invalidation.
- Abort handling.
- Error normalization.
- Object URL ownership.
- Event payload creation.

### Browser tests

- Load each supported format from a URL.
- Load each supported format from a `File`.
- Load a Blob and binary buffer.
- Use worker mode.
- Use main mode.
- Trigger worker fallback.
- Reload a second document.
- Disconnect while loading.
- Reconnect after disconnect.
- Destroy after ready.
- Resize the host.
- Navigate pages, sheets, and slides.
- Zoom and fit.
- Find text.
- Click or disable hyperlinks.
- Download the original source.
- Handle malformed, encrypted, legacy, and unsupported files.

### Asset tests

- Verify emitted WASM files exist.
- Verify worker files exist.
- Verify `wasmUrl` override.
- Verify static subpath deployment.
- Verify no inline WASM.
- Verify production worker paths.
- Verify optional modules are not fetched by default.

### Visual tests

Use representative fixtures:

```text
simple.docx
tables.docx
images.docx
large.docx
equations.docx
simple.xlsx
formatted.xlsx
charts.xlsx
simple.pptx
shapes.pptx
images.pptx
```

The visual baseline should belong to the upstream renderer. Component-specific screenshots should additionally verify:

- loading shell;
- error shell;
- resize behavior;
- theme;
- toolbar;
- scroll container;
- page/slide spacing;
- XLSX sheet surface.

## Recommended implementation sequence

### Step 1 — Confirm upstream APIs

Before writing the public element:

- Read the current `@silurus/ooxml` README and API declarations.
- Confirm constructors and load methods for all three viewers.
- Confirm worker-mode options.
- Confirm source types.
- Confirm destroy behavior.
- Confirm navigation and zoom methods.
- Confirm find and hyperlink APIs.
- Confirm error classes and codes.
- Confirm optional renderer injection.

Record any differences in `docs/ooxml-integration-notes.md`.

### Step 2 — Build a raw integration harness

Create a small internal demo that directly imports the three upstream viewers and loads one fixture for each format.

The harness must test:

- main mode;
- worker mode;
- local File input;
- URL input;
- WASM URL override;
- production build asset paths;
- destroy/reload.

Do not add the custom element until this harness is reliable.

### Step 3 — Implement one adapter

Start with DOCX because it is closest to the PDF-style continuous page viewer.

Implement:

- source loading;
- worker preference;
- loading completion;
- destroy;
- zoom;
- page navigation;
- errors;
- resize/refit.

### Step 4 — Add the custom element shell

Add:

- Shadow DOM;
- status UI;
- attributes;
- lifecycle state machine;
- load generation protection;
- common events.

### Step 5 — Add XLSX and PPTX adapters

Preserve their format-specific interaction models. Avoid premature common abstractions for cells, pages, and slides.

### Step 6 — Harden assets and lifecycle

Test all bundler and static-hosting cases. Test disconnect/reconnect, abort, and stale asynchronous work.

### Step 7 — Add common UX

Add only verified common controls:

- zoom;
- fit;
- find;
- download original;
- print where supported;
- theme;
- optional toolbar.

### Step 8 — Measure and document

Publish bundle and runtime measurements. Document supported features, known limitations, worker requirements, and asset deployment instructions.

## Decisions to defer

The following decisions should remain open until the architecture spike is complete:

- Whether the combined package uses dynamic imports.
- Whether to publish separate format entry points.
- Whether `mode` means strict mode or preferred mode.
- Whether `print()` can be common across all formats.
- Whether optional renderers belong in the public custom-element API.
- Whether the component should expose headless-engine sharing.
- Whether a built-in toolbar belongs in the core package or a separate package.

## Definition of done for the architecture spike

The spike is complete when:

- DOCX, XLSX, and PPTX load through their real upstream viewers.
- Worker mode has been tested in a production build.
- Main-mode fallback behavior is known.
- WASM and worker assets resolve in the supported build environments.
- `wasmUrl` behavior is documented and verified where needed.
- Format-specific bundle sizes are recorded.
- Optional modules are confirmed to be excluded by default.
- Destroy/reload behavior is tested.
- The adapter API is based on actual declarations rather than assumptions.
- All known upstream pitfalls are recorded.

Only after these conditions are met should the stable public API of `<office-viewer>` be finalized.
