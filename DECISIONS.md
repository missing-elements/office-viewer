# Accepted Architecture Decisions

**Date:** September 3, 2026

This document records the decisions accepted after reviewing `PLAN.md` and `ARCHITECTURE.md`. Where an earlier document describes a provisional alternative, this document is authoritative for implementation.

## 1. Rendering mode

The component uses worker rendering by default.

Public API:

```text
mode="worker|main"
```

Semantics:

- `worker` is the default and preferred mode.
- `main` explicitly selects main-thread rendering.
- An `auto` mode is not needed.
- If worker mode cannot be used because of browser or document limitations, the implementation may fall back to main mode only when this is required by `@silurus/ooxml` compatibility behavior.
- The effective mode should be observable internally and may be exposed through a read-only property if the upstream API makes this reliable.

Do not add a strict `worker-required` option in the initial release.

## 2. Locale

Remove `locale` from the initial public API.

The first release does not define component UI localization, document language behavior, or spreadsheet number-format localization. These may be revisited later after a concrete upstream integration requirement is identified.

## 3. Printing

Remove `print()` from the initial common API.

Printing is not yet a reliable cross-format operation because DOCX, XLSX, and PPTX have different pagination and virtualization behavior. It can be added later as a format-specific capability after implementation and browser testing prove it works correctly.

## 4. Download

The public method is:

```ts
downloadOriginal(): Promise<void>
```

It downloads the original source bytes. It does not export a rendered image, PDF, or modified Office document.

Behavior:

- URL sources download the fetched original bytes when available.
- `File`, `Blob`, and binary sources download the retained source bytes.
- A generated filename is used when the source has no filename.
- The component must not claim to save or round-trip an edited document.

## 5. `src` and `load()` relationship

`load(source)` is the canonical operation.

The `src` attribute is a convenience API:

- Changing `src` resolves and loads the new URL.
- Calling `load(source)` does not need to reflect the source into the `src` attribute.
- `reload()` reloads the last retained source and options.
- `file-name` supplies metadata for sources that do not have a filename.
- `file-type` supplies an explicit format override.

A direct binary load must remain reloadable. The implementation may retain an internal copy of caller-provided bytes when necessary.

## 6. Source ownership and transferables

The default behavior must preserve caller-owned data.

- Do not detach caller-owned `ArrayBuffer` values unexpectedly.
- Do not mutate caller-owned `Uint8Array` values.
- If a buffer must be transferred to a Worker, make an internal copy first.
- Do not add a public transfer-ownership option in the initial release.
- Object URLs created by the component must be revoked by the component.
- Caller-provided URLs must never be revoked by the component.

Performance optimizations must not silently change ownership semantics.

## 7. Worker and WASM asset configuration

Use the worker implementation managed by `@silurus/ooxml`. Do not create a second rendering worker in `office-viewer`.

For the initial public API:

- Expose `wasmUrl` only if the current upstream declarations support the required configuration.
- Do not expose `workerUrl` until the upstream package explicitly provides a supported worker URL contract.
- Treat worker URL resolution as an upstream/bundler concern unless verified otherwise.

The integration must still test emitted worker paths, WASM paths, static hosting, non-root base paths, and explicit WASM URL overrides.

## 8. Format detection

Use a conservative staged approach.

### Initial release

Determine the format from:

1. Explicit `file-type`.
2. `load()` format option.
3. `File.name`, `file-name`, or URL pathname extension.

If an extensionless binary source has no explicit format, fail with a clear error.

### Later enhancement

Add ZIP package-part inspection if real use cases justify it:

```text
word/document.xml       → DOCX
xl/workbook.xml         → XLSX
ppt/presentation.xml    → PPTX
```

Do not add a ZIP inspection dependency to the MVP solely to support ambiguous extensionless files.

## 9. Worker URL

Do not promise a public `workerUrl` option in the initial release. Only expose configuration that is confirmed by the actual `@silurus/ooxml` declarations and build behavior.

The architecture spike must document how the upstream package creates workers and how Vite and plain browser builds resolve them.

## 10. Styling

Do not require a Paper & Ink theme unless it provides real styling value for the outer component shell.

The component may style its own:

- background or desk;
- page/slide gaps;
- page shadows;
- loading and error UI;
- toolbar, if added;
- XLSX shell elements, if the upstream viewer permits styling.

It must not imply that authored document content can be themed. If the upstream viewers do not provide meaningful styling hooks, skip the custom theme and provide only minimal layout variables.

## 11. DOCX and PPTX scrolling

Use the upstream continuous scroll viewers where they provide the best fit:

- `DocxScrollViewer` for DOCX.
- `PptxScrollViewer` for PPTX.

When progressive layout is enabled:

- page/slide availability may be provisional;
- final counts must not be reported until the upstream layout-completion signal;
- printing, final page-count UI, and final snapshots must wait for layout completion;
- resize/refit must not force a complete reload.

The adapter must translate visible-page/visible-slide callbacks without inventing unsupported guarantees.

## 12. XLSX surface

Use `XlsxViewer` and preserve spreadsheet-specific behavior.

XLSX is not a paginated document. The common component API must not make `page` or `slide` concepts apply to XLSX.

Sheet navigation, cell/range selection, and workbook-specific operations remain format-specific. Do not expose them as universal methods unless the final typed API can represent their semantics accurately.

## 13. Disconnect and reconnect

Distinguish temporary detachment from permanent destruction.

### `disconnectedCallback()`

- Destroy or release the active upstream viewer.
- Remove listeners and release render resources.
- Retain the last source and configuration needed for reattachment.

### Reconnection

- Recreate the upstream viewer and reload the retained source when appropriate.
- Prevent reconnect logic from running after permanent destruction.

### `destroy()`

- Permanently release the upstream viewer.
- Abort active work.
- Release image bitmaps and workers through the upstream lifecycle.
- Revoke component-owned object URLs.
- Clear active resources.
- Require a new explicit `load()` call for future use.

This behavior is designed for framework lifecycles that temporarily remove and reinsert elements.

## 14. Dependency strategy

Use `@silurus/ooxml` as a normal runtime dependency for the initial package.

Do not make it a peer dependency in the first release. A normal dependency gives the component a tested upstream version and simpler installation. Reconsider this only if consumers need to control the engine version independently.

Dynamic imports and format-specific exports remain implementation decisions for the architecture spike. They must be selected using actual bundle and asset measurements.

## 15. Test stack

Use the same primary test stack as `missing-elements/pdfjs-viewer`:

- Vitest.
- Vitest browser tests.
- WebdriverIO browser provider.

Do not add Playwright to the initial project unless Vitest and WebdriverIO cannot reliably test production worker behavior or visual output.

## 16. Privacy wording

Use precise privacy language:

- For `File`, `Blob`, `ArrayBuffer`, and `Uint8Array` sources, the component processes bytes locally and does not upload them.
- For URL sources, the browser fetches the URL supplied by the consumer; that network request contacts the URL origin.
- The component does not send document bytes to a service of its own.

Avoid claiming that every source is private merely because rendering occurs locally.

## Updated MVP API direction

```ts
export type OfficeFormat = "docx" | "xlsx" | "pptx"
export type OfficeViewerMode = "worker" | "main"

export type OfficeSource =
  | string
  | URL
  | File
  | Blob
  | ArrayBuffer
  | Uint8Array

export interface OfficeViewerLoadOptions {
  format?: OfficeFormat
  fileName?: string
  mode?: OfficeViewerMode
  wasmUrl?: string
  signal?: AbortSignal
}

interface OfficeViewerElement extends HTMLElement {
  load(source: OfficeSource, options?: OfficeViewerLoadOptions): Promise<void>
  reload(): Promise<void>
  setScale(scale: number): void
  fitWidth(): void
  fitPage(): void
  findText(query: string): void
  clearFind(): void
  downloadOriginal(): Promise<void>
  destroy(): void
}
```

The exact `wasmUrl` type and adapter options must be aligned with the current upstream declarations before implementation. Do not publish this interface unchanged without verification.

## Final implementation priority

1. Verify `@silurus/ooxml` declarations and examples.
2. Build a raw DOCX/XLSX/PPTX integration harness.
3. Verify worker mode and WASM assets in production builds.
4. Measure per-format bundle and runtime costs.
5. Implement the DOCX adapter and custom-element shell.
6. Add XLSX and PPTX adapters.
7. Harden reconnect, cancellation, and source ownership.
8. Add only verified common interactions.
9. Add minimal shell styling if it provides clear value.
10. Finalize and document the public API.
