# DocViewKit Viewer and Engine

DocViewKit is a lightweight document viewer for everyday OA attachment previews, approval workflows, admin portals, CRM/ERP systems, cloud drives, and customer-facing apps, as well as enterprise SaaS, AI knowledge bases, legal review, and financial audit. Quick integration starts with the ready-made Viewer: import the component, mount `<docviewkit-viewer>`, and call `open(file)`. A compact core and on-demand format packs keep unused parsers out of the initial load, with no document-conversion server to deploy. Use the Engine API when you need custom rendering or source-object access.

Apache-2.0 software with optional paid support, customization and enterprise delivery.
[Source repository](https://github.com/docviewkit/viewer) · [Online demo](https://docviewkit.com/en/demo/) · [Documentation](https://docviewkit.com/docs/quickstart/) · [Contact services](mailto:novalag778@gmail.com)

A frontend-local, object-aware rendering SDK with a compact OOXML/flat-format core and optional OpenDocument, iWork, legacy Office, WPS Office, PDF, XPS/OpenXPS, and OFD extensions. It accepts document bytes, identifies binary formats from content, parses and renders locally, and maps hit-tested objects back to their source coordinates.

The public `@docviewkit/viewer` package includes the ready-made Viewer, the Engine API at `/engine`, and optional format packs. A frontend-local reference Viewer is included under `examples/` to exercise navigation, search, selectable text, selected-text and document-font resolution, virtualized sheets, continuous pages, printing, fullscreen, diagnostics, and source-object inspection.

## Engine quick start

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

## Default Viewer Web Component

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

## License and paid services

DocViewKit Viewer, Engine, format adapters and project source code are available
under [Apache License 2.0](LICENSE). No runtime license, domain registration or
mandatory Viewer branding is required. Engine APIs, UI slots, CSS parts, minimal
mode and text watermarks are available to everyone.

Third-party components, fonts and test assets retain their own licenses and
copyrights; see [NOTICE](NOTICE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
The project license does not grant rights to the DocViewKit trademark.

Paid services cover integration support, customization, compatibility work,
performance tuning and enterprise delivery under an agreed scope and SLA.
Contact [DocViewKit](mailto:novalag778@gmail.com). Payment is for service delivery and
support, not permission to use, modify or redistribute the open source software.
Legacy `license` options and `evaluateLicense()` remain as deprecated compatibility
APIs and no longer restrict capabilities.

These terms apply to releases built from this source; older published versions retain their included licenses. The website requires no account, sign-in or application registration. Public bug reports belong in [GitHub Issues](https://github.com/docviewkit/viewer/issues); share confidential support files only through an agreed private channel.

## Supported formats

- Presentation: `.pptx`/`.pptm`, `.ppsx`/`.ppsm`, `.potx`/`.potm`
- Spreadsheet: `.xlsx`/`.xlsm`, `.xltx`/`.xltm`
- Text: `.docx`/`.docm`, `.dotx`/`.dotm`
- Flat formats: `.csv` and `.rtf`

Optional native format packs add:

- OpenDocument: `.odp`/`.otp`/`.fodp`, `.odg`/`.otg`/`.fodg`, `.ods`/`.ots`/`.fods`, and `.odt`/`.ott`/`.fodt` with the same source-mapped rendering model
- iWork: `.pages` single-file packages (saved pagination, styled body text, inline raster images, bounded table reconstruction, cached chart geometry, and embedded vector-PDF drawing content); `.numbers` single-file packages (native sheets, tables, tile storage, strings, rich text, numbers, dates, booleans, and source-mapped cells); `.key` single-file packages (a bounded browser-local subset of static text and original JPEG/PNG objects)
- Legacy Office: formatted binary `.doc` paragraphs, `.xls` cells, and `.ppt` text atoms with bounded approximate layout
- WPS Office: `.wps`, `.et`, and `.dps` CFB documents through a dedicated lazy pack, reusing the validated Word/BIFF/PowerPoint record parsers without launching WPS Office or substituting an embedded preview; see the [WPS FormatPack design and acceptance contract](docs/WPS-FORMAT-PACK.md)
- PDF (`.pdf`): fixed pages with a retained object index, first-page-priority/on-demand parsing, searchable text, vector paths, images, Form XObjects, and password-protected Standard Security documents using RC4, AES-128, or AES-256 revision 5
- XPS/OpenXPS (`.xps`, `.oxps`): fixed-document sequences and pages with searchable exact-outline glyphs, vector paths, PNG/JPEG/TIFF/JPEG XR images, ICC/ContextColor and ColorConvertedBitmap color management, XPS gradients and opacity masks, transforms, clips, ImageBrush/VisualBrush viewboxes/viewports/tiling, resources, and deobfuscated embedded fonts
- OFD (`.ofd`, versions 1.0 and 1.1): lazily materialized fixed pages with searchable/source-mapped positioned text, embedded fonts and raster images, paths and composite graphics, solid/axial/radial and Gouraud mesh paint, inherited drawing parameters, transforms/opacity, geometry clips, templates, static annotation appearances, vector OFD seals, outlines, and local SM3/SM2 signature-integrity checks

Final rendering must come from parsed document content and structure. Embedded thumbnails, previews and cover images are limited to navigation or temporary loading placeholders. Some legacy iWork paths still return embedded previews; these are a known implementation gap and are not accepted final rendering or native-fidelity evidence. Isolatable unsupported or damaged content should produce diagnostics while the remaining document stays usable. See [the format support boundary](docs/SUPPORT.md).

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

## First-open performance

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

TXT, standalone HTML/HTM, and XLSB are not rendered and return `UNSUPPORTED_FORMAT`. PDF remains unsupported while its pack is disabled. ZIP64 packages are supported within the same input, entry-count, expanded-size, compression-ratio, and operation limits as ordinary ZIP packages. The current feature depth and intentional fallbacks are documented in [docs/SUPPORT.md](docs/SUPPORT.md).

Recognized unsupported, approximate, substituted, external, or active content produces structured diagnostics instead of disappearing silently.

### Embedded picture formats

Embedded picture parts in supported presentation, spreadsheet, and text-document positions are identified from their bytes as well as their declared media type. The modern Office/Open XML set is covered locally:

- PNG, JPEG/JPG/JPE/JFIF, GIF, WebP, BMP/DIB/RLE, self-contained graphic SVG, TIFF/TIF, ICO, PCX, and JPEG 2000 (JP2/J2K/JPC) use local decode paths.
- EMF, EMF+ records carried by EMF, WMF, and their gzip-compressed EMZ/WMZ forms use a best-effort metafile renderer and report approximate fidelity.
- SVG text reuses successfully registered document-embedded and host-provided font binaries, with the same embedded-before-host priority as normal document text. Browser-only fonts cannot expose their bytes to the isolated codec Worker; when no reusable binary font exists, SVG decoding fails explicitly and an Office-provided raster fallback is used when available.
- EPS and legacy PICT/PCT/PIC/PCZ are deliberately blocked, matching the current Office for Windows security posture; they are never handed to an image decoder.

“Non-approximate” describes the engine's codec path, not pixel identity with a particular Office release. Animated GIF/WebP content is currently rendered as a static frame, and external resources referenced by SVG are never fetched. See the precise surface and per-document limits in [docs/SUPPORT.md](docs/SUPPORT.md).

### Embedded audio and video

PPTX-family media attached through DrawingML media relationships and ODP `draw:plugin` media are exposed by `render()` in `RenderResult.media`. The Inspector places local, user-controlled playback over the authored object bounds while retaining the document's poster image in the static Canvas result. Supported browser media containers are MP4/M4A, QuickTime, WebM, Ogg, MP3, WAV, and AAC; actual codec availability follows the host browser. Media stays in memory, external relationships are blocked, declared types are checked against payload signatures, and playback never autostarts. Hosts with a Content Security Policy must permit `media-src blob:` for the in-memory player URLs.

## Security model

All input is untrusted. The runtime includes its own bounded ZIP/Deflate, XML, OLE, PDF syntax/content, and PDF Standard Security implementations, rejects DTD/entities and unsupported encryption handlers, validates CRC and package paths, blocks active/external content, enforces document and render budgets, and performs parsing off the main thread. Non-native picture codecs run in a bounded pool of isolated module Workers with byte, pixel, queue, and wall-clock limits. See [docs/SECURITY.md](docs/SECURITY.md).

The shipped runtime contains pinned, local-only image codecs rather than a network conversion service or a general Office renderer. It never sends document bytes to those projects or to any remote endpoint. Their versions, purposes, and licenses are recorded in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

## Frontend runtime compatibility

The Viewer runs in frontend environments that provide DOM/Custom Elements, module Workers, WebAssembly, Canvas/OffscreenCanvas, ImageBitmap, and FontFace. These include browsers and compatible WebView hosts in Electron, Tauri, Ionic and Capacitor. React Native and Flutter applications can embed it through a WebView that provides the same APIs; their native rendering layers do not run the Web Component directly. Validate the actual host engine, asset loading, CSP, fonts, and required interactions on each target platform.

Browser engine baseline:

- Chrome and Edge 103+
- Firefox 114+
- Safari and iOS 16.4+
- Cross-platform application WebViews must provide the same required Web APIs

Host CSP must allow the configured `worker-src` and same-origin `connect-src` access to the packaged codec assets. No document bytes are uploaded or fetched by the engine.

## Development

Requirements: Node.js 24 (the release CI runtime), npm, and a rustup toolchain with `wasm32-unknown-unknown`.

```sh
git clone https://github.com/docviewkit/viewer.git
cd viewer
npm ci
rustup target add wasm32-unknown-unknown
npm test
```

`npm run build` produces ESM declarations, source maps, the document and image-codec Worker modules, the base Wasm, the optional native format-pack Wasm binaries (with iWork and PDF sharing one binary), and the required local font assets.

The root `@docviewkit/sdk` package is a private build workspace to prevent accidental npm publication. Public `@docviewkit/viewer` artifacts are generated by `scripts/prepare-release.mjs`; the workspace flag does not restrict Apache-2.0 source usage. See [release instructions](docs/RELEASING.md) and [website source and deployment instructions](https://github.com/docviewkit/website).

Contributions must include a real case that fails before the change and passes afterward, preserve existing passing cases, and keep equivalent format behavior in the shared implementation. Rendering changes also need actual browser evidence for the affected behavior; corpus documents require permission for public redistribution. Follow [AGENTS.md](AGENTS.md) and the [accuracy policy](docs/ACCURACY.md).

Run repeatable browser performance samples, or compare them with a checked baseline:

```sh
npm run performance -- --write-baseline /tmp/office-viewer-perf.json sample.pptx sample.xlsx
npm run performance -- --baseline /tmp/office-viewer-perf.json sample.pptx sample.xlsx
npm run performance -- --unit 400 large-document.pdf
```

### Upload inspector

Run the browser-local reference Viewer:

```sh
npm run test:server -- start
```

The managed test server always rebuilds the latest code before it starts.
Use `restart`, `status`, `logs`, or `stop` in place of `start` to manage the
Inspector at [http://127.0.0.1:4173/](http://127.0.0.1:4173/). Logs are stored
under `.cache/test-server/`.

The official website runs independently from [docviewkit/website](https://github.com/docviewkit/website).

The lower-level `npm run inspect` command starts an already-built runtime without process management. Open [http://127.0.0.1:4173/](http://127.0.0.1:4173/). The page accepts every format listed above. It provides slide/page navigation and thumbnails, sheet tabs with virtual scrolling and frozen panes, lazy continuous text pages, search, selectable/copyable text, URL-text safe links, zoom, fullscreen, printing, and document/render diagnostics with hit-tested native source mappings. It never uploads document bytes and does not edit or save documents.

### Official-source corpus validation

The fast corpus check downloads nine pinned OOXML/ODF fixtures, verifies their SHA-256 values, and exercises identification, parsing, diagnostics, source mapping, hit testing, and a bounded-rejection case:

```sh
npm run test:corpus
```

The full check uses sparse checkouts pinned to Microsoft Open XML SDK, OASIS ODF TC, and TDF ODF Toolkit commits:

```sh
npm run test:corpus:full
```

Assets stay in `.cache/official-corpus/`; the machine-readable report is written to `output/official-corpus-report.json`. See [the validation baseline](docs/research/CORPUS-VALIDATION.md), [OOXML source research](docs/research/OOXML-TEST-ASSETS.md), and [ODF source research](docs/research/ODF-TEST-ASSETS.md). These upstream assets do not provide a six-format visual oracle and do not constitute ISO, ECMA, or OASIS certification.

PDF compatibility has a separate pinned, hash-verified
[`py-pdf/sample-files` corpus](docs/research/PY-PDF-SAMPLE-FILES.md):

```sh
npm run test:pdf-corpus       # representative smoke set
npm run test:pdf-corpus:full  # all files and pages
npm run visual:pdf-corpus     # selected macOS PDFKit comparisons
npm run visual:pdf-corpus:full # all 293 pages; every page must be >=99%
```

Downloaded PDFs stay in `.cache/pdf-corpus/`; results are written to
`output/pdf-corpus-report.json`, `output/pdf-corpus-visual/`, and
`output/pdf-corpus-visual-full/`. The corpus is compatibility evidence, not
certification or a malicious-input security suite.

Architecture details live in [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## Accuracy testing

The separate `@docviewkit/viewer/accuracy` test module evaluates independent source-oracle and renderer snapshots across structure/content, normalized geometry, text layout, exact and tolerant pixels, windowed SSIM, object regions, hit/source mapping, compatibility diagnostics, determinism, and package metamorphism. It emits one `AccuracyReport` per file; a high whole-page visual score cannot conceal missing objects or incorrect mappings.

```sh
npm run test:accuracy
npm run accuracy -- suites/release-accuracy.mjs --output output/accuracy
```

Release corpus manifests must cover minimal single-feature, combination, authorized real enterprise, large, malformed, and malicious documents. Every compatibility Bug must register a minimal reproduction, independent oracle, and automated test. See [docs/ACCURACY.md](docs/ACCURACY.md).

External visual acceptance uses Microsoft Office as the only golden for PPTX/PPT/ODP, DOCX/DOC/ODT, and XLSX/XLS/ODS; Apple iWork for Keynote/Pages/Numbers; and WPS Office for WPS/ET/DPS. LibreOffice and unrelated applications are prohibited as native or release acceptance oracles. `visual:native:reference` records reviewed native goldens, and `visual:native:capture` captures every OfficeViewer unit in Chromium under the locked `native-visual-v1` policy. Its 99.5% tolerant-pixel and 99% SSIM/object-region gates are stricter than a 2% visual-difference target. A release suite must cover every format and trusted corpus class, verify the fixture/reference/golden/candidate hash chain and pinned machine-checkable fields, and reject missing pages, slides, sheets, or candidate-environment drift. Manual imports additionally require a hash-bound human review attestation; local tooling does not cryptographically authenticate the reviewer or producer. See [External visual acceptance](docs/NATIVE-VISUAL-ACCEPTANCE.md).
