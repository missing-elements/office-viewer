# Changelog

## [Unreleased]
### ✨ New Features
- Initial implementation of the core `office-viewer` component.
- Support for XLSX and PPTX formats alongside DOCX.
- Added advanced lifecycle controls: `load(source)`, `reload()`, and comprehensive state handling (idle $\to$ loading $\to$ ready $\to$ error).
- Implemented structured error reporting covering network, parsing, and resource limits.

### 🐛 Bug Fixes
- Addressed asset loading stability across different module environments.
- Improved resilience to stale asynchronous requests using request generation IDs.

### 🧹 Improvements
- Defined clear API boundaries for common viewer actions (e.g., navigation, zooming, finding text) using abstract methods in `viewer-adapter.ts`.
- Implemented robust source normalization logic within `office-viewer-element.ts` to handle `File`, `Blob`, and `ArrayBuffer` inputs.

## [0.0.0-spike] - Initial Integration Spike
- Core architecture validated against upstream `@silurus/ooxml` API.
- Initial hooks for DOCX, XLSX, and PPTX viewers established.

*Note: This changelog is a placeholder; entries must be updated with finalized version numbers.*