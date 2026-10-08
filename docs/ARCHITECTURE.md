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

## Integration reference

### Engine quick start

```ts
import { createOfficeEngine } from "@docviewkit/viewer/engine";

const engine = await createOfficeEngine();
const document = await engine.open(await file.arrayBuffer());

const frame = await document.render({
  unitIndex: 0,
  scale: 1,
  watermark: "Confidential", // Optional repeated text in the rendered bitmap.
});
canvas.width = frame.pixelWidth;
canvas.height = frame.pixelHeight;
canvas.getContext("2d")!.drawImage(frame.bitmap, 0, 0);
frame.bitmap.close();

const hits = await document.hitTest({ unitIndex: 0, x: 120, y: 80 });
console.log(hits[0]?.object.type, hits[0]?.object.source);

const matches = await document.searchText({ query: "budget", wholeWord: true });
const textLayerObjects = await document.listObjects({ unitIndex: 0, textOnly: true });
console.log(textLayerObjects[0]?.fontRuns); // authored family, rendered family, source, and text range

document.close();
engine.close();
```

### Default Viewer Web Component

Use the framework-independent Viewer when the host needs a complete read-only UI instead of a custom Engine integration. The host acquires the document and owns permissions, document lists, downloads, business workflow, and AI; the component owns parsing, rendering, navigation, search, source location, diagnostics, accessibility, and cleanup.

Install `@docviewkit/viewer` and mount the component in your page before querying it:

```html
<docviewkit-viewer></docviewkit-viewer>
```

```ts
import "@docviewkit/viewer";

const viewer = document.querySelector("docviewkit-viewer");
viewer.config = {
  locale: "zh-CN", // Any BCP 47 locale; unsupported catalogs fall back to English.
  theme: "auto",   // Follows prefers-color-scheme; "light" and "dark" are explicit.
  navigation: "auto",
  initialPageMode: "continuous", // Default; set "single" for one-page-at-a-time navigation.
  minimal: false, // Hide Viewer chrome and leave only the rendering workspace.
  watermark: "Confidential", // Included in pages, thumbnails, and printing.
  features: {
    search: true, // Set false to hide both the search button and search box.
    diagnostics: true, // Hidden by default; opt in to the compatibility report.
    print: false,
    fullscreen: true,
    interactionModeSwitcher: true, // Hidden by default; shows the three-segment mode control.
    interactionMode: "text", // "object" (default), "display", or selectable/copyable "text".
  },
};

await viewer.open(file);
await viewer.reveal({ kind: "object", objectId: "paragraph-42" });
```

`watermark` accepts 1–256 characters and paints repeated diagonal text into rendered bitmaps without changing the source document, object list, search results, hit testing, or source mapping. The option is available to all integrations. Engine integrations can set the same option per `document.render()` request.

The Viewer never acquires documents on its own. Production hosts should call
`open()` only after applying their own authorization and document-selection
workflow. The reference page enables direct file drop for testing; append
`?fileDrop=false` to disable it in an embedded reference-page integration. Its
local-file button is hidden by default; append `?filePicker=true` to show it.

Visible commands are icon-first. Localized `aria-label`, `title`, search, state, error, and diagnostic text remain available to assistive technology and pointer users. English and Simplified Chinese are built in; applications can supply another catalog without replacing the component:

```ts
viewer.config = {
  locale: "fr",
  messages: {
    fr: {
      search: "Rechercher dans le document",
      searchPlaceholder: "Rechercher",
    },
  },
};
```

Customization is intentionally bounded: `features` controls commands and interactions. Set `features.search: false` to hide both the search button and search box. Diagnostics and the segmented interaction-mode switcher are hidden by default and enabled with `features.diagnostics: true` and `features.interactionModeSwitcher: true`; print and fullscreen buttons are independently controlled by `features.print` and `features.fullscreen`. `interactionMode` has exactly three mutually exclusive values: `"object"` preserves the default point-and-click object location behavior and emits `docviewkit-objectselect`; `"display"` provides a non-interactive rendered surface without object hit-testing or selectable text; `"text"` enables native text selection/copy and disables object hit-testing. Safe hyperlinks are active only in `"object"` mode and remain independently configurable with `features.hyperlinks`. `navigation` controls the rail. Set `minimal: true` to hide the toolbar, navigation, status bar, and spreadsheet chrome, leaving only the rendering workspace. Paginated documents and presentations open in lazy `continuous` mode by default; set `initialPageMode: "single"` for one-page-at-a-time navigation. Spreadsheets retain their virtualized single-sheet viewport. Semantic CSS variables such as `--dv-accent`, `--dv-background`, and `--dv-navigation-width` control appearance, stable `::part` regions expose the shell, toolbar, navigation, workspace, document, continuous-page surface, status bar, and diagnostics, and `toolbar-start` / `toolbar-end` slots accept host-owned actions. Applications that replace the information architecture should use the Engine API.

For compatibility, the former `textSelection` and `objectSelection` feature flags are still accepted but deprecated. `textSelection: true` maps to `"text"`; otherwise `objectSelection: false` maps to `"display"`; all other legacy combinations map to `"object"`. An explicit `interactionMode` always wins.

Legacy `license` options and `evaluateLicense()` remain as deprecated compatibility APIs and no longer restrict capabilities.

The toolbar theme button cycles through system, light, and dark modes. It updates the `theme` attribute, which takes precedence over `config.theme`; remove the attribute to return control to the host configuration. The dark theme changes only Viewer chrome and workspace colors. Rendered document surfaces remain white and retain authored colors rather than being inverted.

Viewport-driven hosts can pass `{ priority: "interactive", supersedeKey: "surface", signal }`
as the second `render()` argument. Worker execution runs one prioritized render queue, drops
older work with the same key, and lets interactive work interrupt prefetch work. To avoid the
default defensive input copy, pass an `ArrayBuffer` to `open(buffer, { transferInput: true })`;
Worker mode transfers and detaches that caller-owned buffer.

Hosts may provide local font binaries without granting the engine network access:

```ts
const engine = await createOfficeEngine({
  fonts: [{
    family: "Acme Sans",
    bytes: await (await fetch("/assets/acme-sans.woff2")).arrayBuffer(),
    weight: 400,
  }],
});
```

Large font catalogs can be loaded lazily through one host-owned provider:

```ts
const engine = await createOfficeEngine({
  fontPolicy: "deterministic",
  fontProvider: async (requests, { signal }) => {
    // Match the requested family/style/weight/stretch in an allowlisted manifest.
    // Return binary FontAsset values; the document never supplies a URL.
    return fetchFontAssets(requests, { signal });
  },
});
```

The SDK batches exact face and Unicode-scalar demand, bounds provider calls, verifies an optional SHA-256 expectation, deduplicates concurrent loads, and keeps an Engine-scoped LRU byte cache. The provider remains responsible for its allowlisted service, credentials, licensing, and CORS/CSP configuration. Static and provider assets are byte-limited and copied before ownership crosses a document boundary. Worker documents register independent `FontFace` values and release their font realm when the Worker terminates.

`local-first` is the default: it resolves each exact face as embedded → static host → exact requested `local(...)` face → known platform-compatible `local(...)` candidates → provider → explicit approximation, and does not call the provider when a requested or compatible local face loads successfully. Compatible candidates are a bounded, requested-family-scoped allowlist for common Office fonts; they do not enumerate the system font inventory. `deterministic` is opt-in and skips unmanaged local fonts, resolving as embedded → static host → provider → explicit approximation. Neither policy enumerates the user's installed fonts or requests local-font access permission. Missing faces use an explicit generic approximation (`serif`, `sans-serif`, or `monospace`) with `FONT_SUBSTITUTED`; fallback is never silent. Embedded fonts are registered under document-scoped aliases and released with the document.

DOCX uses a bounded two-pass path: the first parse discovers the actual font runs, the Worker registers the selected faces and measures their Canvas advances, then Wasm repeats line breaking and pagination with that immutable metric table. This materially improves wrapping and page counts, but the result remains approximate because Microsoft Word's complete pagination engine and compatibility rules are not public.

Text-bearing `DocumentObject` values expose immutable `fontRuns` with UTF-16 ranges, the document-authored family, the exact renderer family or alias, and its resolution source. This is the same resolution used for Canvas paint in inline and Worker execution.

`x` and `y` are 96-DPI logical document pixels with the surface origin at its top-left. They do not include render scale or device pixel ratio.

The default runtime creates one module Worker per open document. This is deliberate: cancellation and timeout can terminate an untrusted document without affecting other sessions. `execution: "inline"` exists for controlled tests only and has weaker isolation.

### Supported formats

- Presentation: `.pptx`/`.pptm`, `.ppsx`/`.ppsm`, `.potx`/`.potm`
- Spreadsheet: `.xlsx`/`.xlsm`, `.xltx`/`.xltm`
- Text: `.docx`/`.docm`, `.dotx`/`.dotm`
- Flat formats: `.csv` and `.rtf`

Optional native format packs add:

- OpenDocument: `.odp`/`.otp`/`.fodp`, `.odg`/`.otg`/`.fodg`, `.ods`/`.ots`/`.fods`, and `.odt`/`.ott`/`.fodt` with the same source-mapped rendering model
- iWork: `.pages` single-file packages (saved pagination, styled body text, inline raster images, bounded table reconstruction, cached chart geometry, and embedded vector-PDF drawing content); `.numbers` single-file packages (native sheets, tables, tile storage, strings, rich text, numbers, dates, booleans, and source-mapped cells); `.key` single-file packages (a bounded browser-local subset of static text and original JPEG/PNG objects)
- Legacy Office: formatted binary `.doc` paragraphs, `.xls` cells, and `.ppt` text atoms with bounded approximate layout
- WPS Office: `.wps`, `.et`, and `.dps` CFB documents through a dedicated lazy pack, reusing the validated Word/BIFF/PowerPoint record parsers without launching WPS Office or substituting an embedded preview; see the [WPS FormatPack design and acceptance contract](WPS-FORMAT-PACK.md)
- PDF (`.pdf`): fixed pages with a retained object index, first-page-priority/on-demand parsing, searchable text, vector paths, images, Form XObjects, and password-protected Standard Security documents using RC4, AES-128, or AES-256 revision 5
- XPS/OpenXPS (`.xps`, `.oxps`): fixed-document sequences and pages with searchable exact-outline glyphs, vector paths, PNG/JPEG/TIFF/JPEG XR images, ICC/ContextColor and ColorConvertedBitmap color management, XPS gradients and opacity masks, transforms, clips, ImageBrush/VisualBrush viewboxes/viewports/tiling, resources, and deobfuscated embedded fonts
- OFD (`.ofd`, versions 1.0 and 1.1): lazily materialized fixed pages with searchable/source-mapped positioned text, embedded fonts and raster images, paths and composite graphics, solid/axial/radial and Gouraud mesh paint, inherited drawing parameters, transforms/opacity, geometry clips, templates, static annotation appearances, vector OFD seals, outlines, and local SM3/SM2 signature-integrity checks

Final rendering must come from parsed document content and structure. Embedded thumbnails, previews and cover images are limited to navigation or temporary loading placeholders. Some legacy iWork paths still return embedded previews; these are a known implementation gap and are not accepted final rendering or native-fidelity evidence. Isolatable unsupported or damaged content should produce diagnostics while the remaining document stays usable. See [the format support boundary](SUPPORT.md).

The optional families are disabled and unloaded by default. Enable them through a dynamic import so only the matching native parser is loaded. The WPS and legacy Office candidates intentionally share the audited CFB/record binary:

```ts
const engine = await createOfficeEngine({
  formatPack: () => import("@docviewkit/viewer/extended-formats")
    .then(({ extendedFormatPack }) => extendedFormatPack),
});
```

Applications that need only WPS formats can load the smaller dedicated subpath:

```ts
const engine = await createOfficeEngine({
  formatPack: () => import("@docviewkit/viewer/wps-formats")
    .then(({ wpsFormatPack }) => wpsFormatPack),
});
const document = await engine.open(bytes, { fileName: "report.wps" });
```

Applications that need only XPS/OpenXPS can load its dedicated subpath:

```ts
const engine = await createOfficeEngine({
  formatPack: () => import("@docviewkit/viewer/xps-formats")
    .then(({ xpsFormatPack }) => xpsFormatPack),
});
const document = await engine.open(bytes);
```

Applications that need only OFD can load its dedicated Rust/Wasm pack:

```ts
const engine = await createOfficeEngine({
  formatPack: () => import("@docviewkit/viewer/ofd-formats")
    .then(({ ofdFormatPack }) => ofdFormatPack),
});
const document = await engine.open(bytes);
```

Set `formatPack: false` (or omit it) to keep ODF, iWork, legacy Office, WPS Office, PDF, XPS/OpenXPS, and OFD inputs unsupported. Native-family detection is content-based. Because WPS reuses the same CFB/record families as binary DOC/XLS/PPT, `.wps`, `.et`, or `.dps` in the original `fileName` selects the isolated WPS pack only after the CFB signature is present; the native parser still validates the reachable main stream and record signature. The selected parser receives a byte-for-byte snapshot of the original input, and no conversion service or external application is involved. iWork and PDF share one Wasm because iWork embeds PDF objects; OFD has its own Wasm; the other native Wasm modules remain independently loaded.

Spreadsheet calculation is also isolated. XLSX files use their authored cached formula results without loading `office-viewer-calc.wasm`; only a formula cell with no cached result requests that module. Set `calculationWasm: false` to keep best-effort blank results and diagnostics, or provide a custom `WasmSource` when runtime assets are hosted separately.

### First-open performance

Configure the Viewer before its first `open()` and start the optional format-pack import as soon as the application knows that those formats are allowed:

```ts
import "@docviewkit/viewer";

const formatPack = import("@docviewkit/viewer/extended-formats")
  .then(({ extendedFormatPack }) => extendedFormatPack);

const viewer = document.querySelector("docviewkit-viewer");
viewer.config = {
  engine: {
    formatPack: () => formatPack,
  },
};

await viewer.open(file);
```

Creating the promise starts fetching the small JavaScript format dispatcher before the document is opened. It does **not** fetch every optional Wasm binary: after bounded content detection, the runtime loads only the matching ODF, PDF/iWork, legacy Office/WPS, XPS, or OFD module. Applications limited to WPS, XPS, or OFD should use the dedicated `wps-formats`, `xps-formats`, or `ofd-formats` subpath instead of `extended-formats`.

For repeat visits, serve versioned Worker, Wasm, font, and codec assets with normal HTTP caching from the same deployment as the Viewer. Keep one Viewer instance across documents and call `close()` between them; reserve `destroy()` for removing the component because it also releases the reusable Engine. Do not preload every format binary, render every page, start document-wide search, or fetch a complete font catalog before first paint. Those actions move bounded on-demand work back into the startup path.

Measure cold-cache and warm-cache results separately. Record the package version, browser, hardware, input byte size and unit count, `open()` completion, first visible render, and diagnostics; compare changes with the same artifact and completion boundary.

Password-protected PDFs report `PDF_PASSWORD_REQUIRED` when no password is supplied and `PDF_PASSWORD_INCORRECT` after a failed attempt. Engine callers can retry with the same bytes and a password; the password is copied only into the document Worker/Wasm open call and cleared from those temporary buffers afterward:

```ts
const document = await engine.open(pdfBytes, { password });
// The Viewer component exposes the same option:
await viewer.open(file, { password });
```

If `Viewer.open()` receives no password, or receives an incorrect password, the Viewer requests it locally through a masked dialog and allows up to three attempts.

Package formats are identified from OPC content types/relationships or the ODF `mimetype`; filenames and extensions are never trusted for binary dispatch. Template and slideshow variants use the corresponding presentation, spreadsheet, or text parser. Macro-enabled packages are readable, but macros and other active content never execute. Flat formats use bounded content sniffing.

TXT, standalone HTML/HTM, and XLSB are not rendered and return `UNSUPPORTED_FORMAT`. PDF remains unsupported while its pack is disabled. ZIP64 packages are supported within the same input, entry-count, expanded-size, compression-ratio, and operation limits as ordinary ZIP packages. The current feature depth and intentional fallbacks are documented in [docs/SUPPORT.md](SUPPORT.md).

Recognized unsupported, approximate, substituted, external, or active content produces structured diagnostics instead of disappearing silently.

#### Embedded picture formats

Embedded picture parts in supported presentation, spreadsheet, and text-document positions are identified from their bytes as well as their declared media type. The modern Office/Open XML set is covered locally:

- PNG, JPEG/JPG/JPE/JFIF, GIF, WebP, BMP/DIB/RLE, self-contained graphic SVG, TIFF/TIF, ICO, PCX, and JPEG 2000 (JP2/J2K/JPC) use local decode paths.
- EMF, EMF+ records carried by EMF, WMF, and their gzip-compressed EMZ/WMZ forms use a best-effort metafile renderer and report approximate fidelity.
- SVG text reuses successfully registered document-embedded and host-provided font binaries, with the same embedded-before-host priority as normal document text. Browser-only fonts cannot expose their bytes to the isolated codec Worker; when no reusable binary font exists, SVG decoding fails explicitly and an Office-provided raster fallback is used when available.
- EPS and legacy PICT/PCT/PIC/PCZ are deliberately blocked, matching the current Office for Windows security posture; they are never handed to an image decoder.

“Non-approximate” describes the engine's codec path, not pixel identity with a particular Office release. Animated GIF/WebP content is currently rendered as a static frame, and external resources referenced by SVG are never fetched. See the precise surface and per-document limits in [docs/SUPPORT.md](SUPPORT.md).

#### Embedded audio and video

PPTX-family media attached through DrawingML media relationships and ODP `draw:plugin` media are exposed by `render()` in `RenderResult.media`. The Inspector places local, user-controlled playback over the authored object bounds while retaining the document's poster image in the static Canvas result. Supported browser media containers are MP4/M4A, QuickTime, WebM, Ogg, MP3, WAV, and AAC; actual codec availability follows the host browser. Media stays in memory, external relationships are blocked, declared types are checked against payload signatures, and playback never autostarts. Hosts with a Content Security Policy must permit `media-src blob:` for the in-memory player URLs.

### Security model

All input is untrusted. The runtime includes its own bounded ZIP/Deflate, XML, OLE, PDF syntax/content, and PDF Standard Security implementations, rejects DTD/entities and unsupported encryption handlers, validates CRC and package paths, blocks active/external content, enforces document and render budgets, and performs parsing off the main thread. Non-native picture codecs run in a bounded pool of isolated module Workers with byte, pixel, queue, and wall-clock limits. See [docs/SECURITY.md](SECURITY.md).

The shipped runtime contains pinned, local-only image codecs rather than a network conversion service or a general Office renderer. It never sends document bytes to those projects or to any remote endpoint. Their versions, purposes, and licenses are recorded in [THIRD_PARTY_NOTICES.md](../THIRD_PARTY_NOTICES.md).

### Frontend runtime compatibility

The Viewer runs in frontend environments that provide DOM/Custom Elements, module Workers, WebAssembly, Canvas/OffscreenCanvas, ImageBitmap, and FontFace. These include browsers and compatible WebView hosts in Electron, Tauri, Ionic and Capacitor. React Native and Flutter applications can embed it through a WebView that provides the same APIs; their native rendering layers do not run the Web Component directly. Validate the actual host engine, asset loading, CSP, fonts, and required interactions on each target platform.

Browser engine baseline:

- Chrome and Edge 103+
- Firefox 114+
- Safari and iOS 16.4+
- Cross-platform application WebViews must provide the same required Web APIs

Host CSP must allow the configured `worker-src` and same-origin `connect-src` access to the packaged codec assets. No document bytes are uploaded or fetched by the engine.
