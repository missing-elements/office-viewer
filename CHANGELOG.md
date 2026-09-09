# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2025-01-10

### Added

- Headless `<office-viewer>` custom element with no shadow DOM and no rendered children.
- Unified source loading for `string` URLs, `Blob`, `File`, `ArrayBuffer`, and `ReadableStream<Uint8Array>`.
- Automatic format detection from file extension, MIME type, and ZIP content inspection for DOCX, XLSX, and PPTX.
- Direct upstream engine access through `getViewer()`, `getDocument()`, and `getEngine()`.
- Load lifecycle events: `loadstart`, `ready`, and `loaderror`.
- Cancellation and stale-load protection via `AbortController` and generation tokens.
- Thin adapter factories for `@silurus/ooxml` viewers (`DocxScrollViewer`, `XlsxViewer`, `PptxScrollViewer`).
- TypeScript declaration output published alongside the ESM build.
- Unit tests for source resolution and load control.
- Browser integration tests covering all three formats and malformed documents.
- Build-output tests verifying production asset shape.
- Bundle-size measurement pages for DOCX, XLSX, PPTX, and combined loading.

### Changed

- Pivoted from a full-featured viewer wrapper to a headless orchestration shell per `DECISIONS.md`.

### Removed

- Wrapper methods and internal DOM rendering; the element now exposes upstream APIs directly.
