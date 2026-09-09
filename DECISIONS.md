# Architectural Conclusions

## Headless-Only Scope (Revised September 9, 2026)

The project has pivoted from a full-featured viewer component to a **headless orchestration shell**. All UI concerns — Shadow DOM, status indicators, toolbars, themes, navigation buttons, and wrapper methods — are explicitly out of scope.

**Rationale:** Consumers need direct access to `@silurus/ooxml` APIs without abstraction leakage. The component's value is lifecycle management, source normalization, and format detection — not visual presentation.

## Component Structure

The element contract is now minimal and stable:

- **No Shadow DOM.** The element creates zero DOM children.
- **No visual chrome.** No loading spinners, error messages, or toolbars rendered by the component.
- **No wrapper methods.** No `setScale()`, `fitWidth()`, `findText()`, `downloadOriginal()`, or navigation methods on the element.
- **Direct API exposure.** Consumers call `getViewer()`, `getDocument()`, or `getEngine()` and interact with upstream instances directly.

The element remains responsible for:
- Attribute parsing (`src`, `file-name`, `file-type`, `mode`, `wasm-url`)
- Source normalization and conservative format detection
- Load lifecycle (generation tokens, abort, stale-load guards)
- Adapter selection and upstream viewer instantiation
- Disconnect/reconnect and destroy semantics
- Lifecycle events (`loadstart`, `ready`, `loaderror`, `destroy`)

## Adapter Contract

Adapters are thin factories, not wrappers. They do not re-expose upstream methods with new names or semantics.

```ts
interface ViewerAdapter {
  readonly format: OfficeFormat
  readonly viewer: DocxScrollViewer | XlsxViewer | PptxScrollViewer | null
  readonly document: DocxDocument | XlsxWorkbook | PptxPresentation | null
  readonly engine: unknown | null

  load(source: AdapterSource, options: AdapterLoadOptions): Promise<void>
  destroy(): void
}
```

Format-specific navigation (pages, sheets, slides) is **not** part of the adapter interface. Consumers access these via the upstream viewer returned by `getViewer()`.

## Events

Only four lifecycle events:

- `loadstart`
- `ready`
- `loaderror`
- `destroy`

No `pagechange`, `slidechange`, `sheetchange`, `zoomchange`, `progress`, or `find` events. Upstream callbacks (e.g., `onVisiblePageChange`) are attached by consumers directly to the viewer instance.

## Next Implementation Focus

1. **Strip existing `OfficeViewerElement`**: Remove Shadow DOM, status UI, `viewport` container, and all wrapper methods.
2. **Simplify adapters**: Remove format-specific navigation methods from adapter interfaces. Expose upstream viewer/document/engine directly.
3. **Update tests**: Assert zero DOM children, no Shadow DOM, and correct upstream instance types from accessors.
4. **Remove `OoxmlIntegrationSpike` harness**: The element should manage adapters directly without an intermediate harness abstraction.
