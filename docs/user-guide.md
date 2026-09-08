# Troubleshooting and Usage Guide

This guide addresses common issues encountered when using `<office-viewer>` and provides best practices for integration.

## Common Errors and Solutions

### 1. `Load failed: The operation was aborted.`
*   **Cause:** The component was disconnected or a new document was loaded before the previous load completed.
*   **Solution:** Ensure that state changes or component removals are handled gracefully using `element.addEventListener('disconnect', ...)` or `element.addEventListener('destroy', ...)` to clean up the viewer.
*   **API Guidance:** Use `element.reload()` to initiate a new load after a temporary interruption, rather than relying solely on attribute changes if the source is complex.

### 2. `Load failed: Cannot find parser assets.`
*   **Cause:** The required WebAssembly (`*.wasm`) or worker JavaScript files were not correctly bundled or served by the web server/CDN.
*   **Solution:**
    *   **Development:** Ensure `pnpm run build` runs successfully and all assets in the `dist/assets/` directory are available.
    *   **Production/CDN:** When deploying, the entire contents of the `dist/assets/` directory must be accessible at the root path used by the component, especially the `*parser_bg*.wasm` and `*render-worker-*.js` files. Check your build configuration for asset versioning/hashing to ensure these files are never lost or inaccessible.

### 3. `Load failed: Invalid source type.`
*   **Cause:** The provided `src` attribute or `load()` source is neither a valid URL string, a local `File`/`Blob`, nor a correctly formatted data URI.
*   **Solution:**
    *   **Local Files:** Always use `File` objects or `Blob`s generated from `File` inputs, and pass them via `load(fileObject)` rather than setting the `src` attribute on a local file path.
    *   **Buffers:** If working with raw bytes, ensure you pass a `Uint8Array` or `ArrayBuffer` directly to `load()`. Do not rely on string representations of binary data unless absolutely necessary.

## Best Practices

*   **State Management:** Always check the component's state first. Check for the `ready` event before attempting UI interaction methods like `findText()` or `downloadOriginal()`.
*   **Asynchronous Operations:** All major operations (`load`, `reload`, `downloadOriginal()`) return a `Promise` and must be awaited.
*   **Cancellation:** For performance-critical sequences, always utilize the `AbortController` pattern when calling `load()` to prevent resource leaks from stale, background requests.