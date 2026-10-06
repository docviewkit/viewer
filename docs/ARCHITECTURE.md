# Architecture

## Runtime API seam

The browser SDK keeps package and rendering internals behind one small document API:

```text
OfficeEngine.open(bytes)
        │
        ▼
OfficeDocument
  ├─ render(request)  → ImageBitmap
  ├─ hitTest(point)   → document objects + ancestors
  ├─ searchText(query) → immutable object text ranges
  ├─ listObjects(filter) → bounded object/text-layer view with resolved font runs
  ├─ getObject(id)    → document object
  └─ close()
```

`createOfficeEngine({ formatPack })` optionally wraps this same interface with six closed native format families. Omitted/`false` means no optional module is imported. A dynamic factory is evaluated only for a content-recognized native candidate; only the selected native Wasm is then loaded.

Callers do not see ZIP entries, XML events, format models, layout passes, display-list nodes, Worker commands, or Wasm memory. The separately exported accuracy module is described below.

## Runtime flow

```text
host bytes
  → per-document module Worker
  → Rust/Wasm safety budgets
  → preliminary content dispatch
      ├─ bounded CSV / RTF adapter
      └─ ZIP + Deflate + CRC
          → OPC content types/relationships
          → native OOXML presentation / spreadsheet / text parser
             (documents, templates, and slideshow aliases)
      └─ optional fallback after native rejection
          → bounded ODF ZIP, iWork ZIP, OLE CFB, PDF-header, XPS, or OFD package probe
          → lazy family Wasm load
          → original-byte ODF, native iWork, DOC/XLS/PPT, WPS/ET/DPS, PDF, XPS, or OFD subset parser
  → font demand (family/style/weight/stretch + Unicode scalars)
  → embedded/static exact-face resolution
  → host fontProvider handshake for unresolved faces
  → Worker FontFace registration + Canvas advance measurement
  → final format-specific layout compiler
      └─ DOCX repeats line breaking/pagination with bounded OVFM metrics
  → render primitives + retained object table + diagnostics
  → versioned binary snapshot
  → Canvas2D replay
      ├─ browser-native image decode in the document Worker
      └─ bounded reusable image-codec Worker pool with packaged SVG Wasm/font assets
  → optional render-tail text watermark
  → ImageBitmap
```

Package aliases are normalized to their parser family: for example PPSX/POTX use the base PPTX parser and OTS uses the optional ODS parser. Macro-enabled variants follow the same read-only path while security inspection blocks active content. TXT, standalone HTML/HTM, and XLSB never enter a text fallback. ODF, OLE compound DOC/XLS/PPT, PDF, XPS, and OFD enter their parsers only when the matching optional pack is configured.

The common display list describes how to draw: affine layers, groups, opacity, paint, paths, effects, images, and text layout. It does not claim that a PPTX shape, XLSX cell, DOCX paragraph, and CSV cell are the same semantic object. Native source references remain format-discriminated through the layout compiler.

The optional text watermark is painted once at the shared render tail, after document objects and before `ImageBitmap` transfer. It therefore covers every format and the Viewer page, thumbnail, and print paths without becoming a document object or changing search, hit testing, source mapping, or `renderedObjectCount`. Both Viewer and Engine expose this option without a runtime license.

## Startup and reuse contract

The host may start and memoize the configured `formatPack` dynamic import before `open()`. This warms only the JavaScript dispatcher; candidate detection still selects one optional Wasm family, and no unused family is fetched or instantiated. The Viewer creates its Engine on the first `open()`, reuses it after `close()`, and releases it on `destroy()`.

IronCalc lives only in `office-viewer-calc.wasm`. Base Core reports a calculation requirement only when an XLSX formula has no cached value; the document Worker then evaluates the original workbook in the optional module, validates its format-neutral cell-result payload, applies it to the same retained Core document, and continues lazy sheet materialization. Cached XLSX files never fetch or compile the calculator. Calculator failure preserves the usable workbook and its localized diagnostic.

First visible content stays ahead of background work: unit materialization, continuous-page rendering, thumbnails, image decoding, and font-provider requests are demand-driven and bounded. Hosts should not turn document-wide search, all-unit rendering, every optional Wasm binary, or a complete remote font catalog into startup prerequisites. Versioned HTTP caching is the supported warm-load mechanism for packaged Worker, Wasm, codec, and font assets.

## Wasm and JavaScript responsibilities

Rust/Wasm owns deterministic work:

- content identification;
- ZIP, Deflate, CRC and safe part lookup;
- bounded XML tokenization;
- bounded flat-text parsing and sanitization;
- native OOXML plus optional ODF format semantics and source mapping;
- bounded native OLE, PDF, and OFD syntax/content parsing without a third-party document engine;
- deterministic geometry and document limits;
- the retained object table and hit-test order;
- structured diagnostics.

The JavaScript runtime owns browser capabilities:

- module Worker lifecycle, cancellation and wall-clock timeout;
- Canvas2D replay and ImageBitmap transfer;
- image signature/dimension validation, browser-native decode, and a bounded pool of reusable codec Workers for Office-specific formats;
- shared visible-image prefetch with concurrency and pixel budgets; slide, page, and sheet-viewport selection remains the unit-level bound;
- exact-face font policy, document-scoped registration, provider handshakes, and Canvas measurement;
- rich-text measurement/replay, including CJK wrapping, tab/list prefixes, paragraph direction, vertical text, auto-fit, and vertical anchoring;
- the small handwritten Wasm ABI adapter.

Large scenes use a versioned binary snapshot rather than JSON. Snapshot v31 carries the current affine layers, paints, paths, effects, rich text, speaker-note formatting, media descriptors, and native optional-format source locators; older supported snapshot versions are decoded explicitly. No `wasm-bindgen`, serialization framework, PDF/Office document library, or complete third-party document renderer is included. The local runtime includes narrowly scoped image codecs and pinned preset-geometry data; see [third-party notices](../THIRD_PARTY_NOTICES.md).

The public font seam remains two options on `createOfficeEngine`: `fontProvider` and `fontPolicy`. Provider caching, request batching, Worker messages, aliases, and the compact `OVFM` table are internal. Static font binaries are prepared once per Engine, while each document Worker must still register its own `FontFace` objects because Worker font realms are isolated. The metric table contains only validated face descriptors, Unicode scalars, and finite `advanceEm` values; Wasm never receives a font URL.

## Coordinates and identity

- Surface and object bounds use 96-DPI logical pixels.
- Render scale and pixel ratio affect output pixels, not hit-test coordinates.
- Unit and object IDs are opaque and stable for the lifetime of one open document.
- Source mappings expose native identifiers where they exist: PPTX Shape ID, ODF element ID/path, XLSX A1 address, ODS row/column, DOCX/ODT paragraph/table/range data, PDF object/stream offsets, OFD part/object paths, and flat-format row/column/text ranges.
- Mapping quality is explicit: `exact`, `derived`, or `approximate`.

## Product boundaries

There is no arbitrary format plugin registry, dependency-injection container, event bus, public XML/PDF tree, generic Office semantic DOM, network conversion service, editor, or save pipeline. The optional-format seam is deliberately closed to native-family Wasm binaries reusing the existing document model. iWork and PDF share one binary because iWork embeds PDF objects. WPS is a separate package/candidate boundary while intentionally sharing the audited legacy CFB and Word/BIFF/PowerPoint binary; an explicit WPS extension never bypasses structural validation.

The browser-local Viewer under `examples/` is deliberately a reference integration rather than a published framework component. It composes the SDK's unit metadata and bounded render requests into thumbnails/navigation, search and text overlays, selected-text and document-font resolution, lazy continuous pages, virtual sheets with frozen-pane overlays, zoom, fullscreen, print, diagnostics, and source inspection. It does not enumerate system fonts or add hidden format semantics to the engine.

## Accuracy test seam

Accuracy evaluation is a separate deep module exported from `@docviewkit/sdk/accuracy`. Its small interface accepts independent expected/actual observations and returns an `AccuracyReport`; corpus adapters hide reference applications, secure enterprise stores, browser pixel capture, and renderer instrumentation. This keeps test-only ZIP mutation and oracle infrastructure out of the browser SDK while applying one scoring and failure vocabulary to every format.

The accuracy module never treats successful parsing or one whole-page screenshot as proof of fidelity. Evidence-free snapshots are invalid and force a zero score. Structure/content, normalized geometry, text line layout, four visual metrics, hit/source mapping, compatibility diagnostics, repeated-render determinism, and four package metamorphisms are independent report layers.

For the optional external pixel seam, reviewed Microsoft Office, Apple iWork, and WPS Office references stay outside the SDK. Microsoft PowerPoint is the oracle for PPTX/PPT/ODP, Word for DOCX/DOC/ODT, Excel for XLSX/XLS/ODS, Apple iWork for its native formats, and WPS Office for WPS/ET/DPS. LibreOffice and unrelated applications cannot supply native or release acceptance references. The browser harness independently captures every OfficeViewer page, slide, or fixed-range sheet viewport. The comparison runner enforces the format-to-oracle mapping, reference and pixel hashes, complete unit topology, declared oracle application/rasterizer fields, the measured candidate environment fingerprint, whole-surface similarity, and deterministic local regions. WPS references are manual reviewed imports because the exporter has no trusted WPS automation path. For manual imports, producer and reviewer provenance remain human-governed attestations rather than cryptographically verified local facts. These goldens are regression evidence for pinned fixtures and environments, not runtime dependencies or a claim of universal pixel parity.
