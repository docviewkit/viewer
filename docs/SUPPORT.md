# Format support

OfficeViewer implements bounded, source-mapped subsets of each format. In this document:

- **Supported** means the feature produces real render objects, can participate in hit testing, and retains a native source reference.
- **Approximate** means the content remains visible and source-mapped through a deterministic fallback, accompanied by a structured diagnostic.
- **Unsupported** means the engine does not pretend to render the feature; it emits a diagnostic when the format parser can identify it, or rejects an unsupported top-level format.

None of these terms promises pixel identity with a particular Office release. Fidelity also depends on fonts, browser text shaping, color management, and the exact producer-specific XML.

## Accepted inputs

| Family | Accepted file extensions | Dispatch and execution policy |
| --- | --- | --- |
| Presentation | `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, `.potm` | OPC content types; slideshow/template packages open as ordinary read-only presentations |
| Spreadsheet | `.xlsx`, `.xlsm`, `.xltx`, `.xltm` | OPC content types; template packages open as ordinary read-only workbooks |
| Text document | `.docx`, `.docm`, `.dotx`, `.dotm` | OPC content types; template packages open as ordinary read-only documents |
| Flat formats | `.csv`, `.rtf` | Bounded content sniffing; no filename trust |
| Optional ODF pack | `.odp`, `.otp`, `.fodp`, `.odg`, `.otg`, `.fodg`, `.ods`, `.ots`, `.fods`, `.odt`, `.ott`, `.fodt` | ODF `mimetype`; packaged formats also use the manifest, while FODP, FODG, FODS, and FODT are parsed directly as flat XML; disabled and unloaded by default |
| Optional iWork pack | `.pages`, `.numbers`, `.key` single-file packages | ZIP/IWA content identity; bounded preview/native-static subsets; disabled and unloaded by default |
| Optional legacy Office pack | `.doc`, `.xls`, `.ppt` in OLE CFB | Validated CFB directory/main stream plus native FIB/CLX, BIFF, or PowerPoint records; disabled and unloaded by default |
| Optional WPS Office pack | `.wps`, `.et`, `.dps` in OLE CFB | CFB content identity plus explicit original extension to choose the isolated WPS pack; native main-stream and Word/BIFF/PowerPoint record validation; disabled and unloaded by default |
| Optional PDF pack | `.pdf` (PDF 1.x fixed-layout documents) | `%PDF-` content identity within the specification header window; native syntax, object-stream, page-tree, and content-operator parsing; disabled and unloaded by default |
| Optional XPS pack | `.xps`, `.oxps` | OPC fixed-representation relationship plus fixed-document sequence, document, and page parts; disabled and unloaded by default |
| Optional OFD pack | `.ofd` (OFD 1.0 and 1.1 fixed-layout documents) | ZIP structure containing `OFD.xml`, a document root, and page content parts; disabled and unloaded by default |

Macro-enabled packages are parsed as their XML family, but VBA, scripts, ActiveX, and other active content never execute. TXT, standalone HTML/HTM, and XLSB are unsupported and return `UNSUPPORTED_FORMAT`. Optional families, including PDF, XPS, and OFD, remain unsupported when their pack is omitted or set to `false`. ZIP64 packages share the ordinary package parser's hard input, entry-count, expanded-size, compression-ratio, and operation limits.

## Optional format packs

The optional pack seam has a closed seven-family vocabulary (`odf`, `iwork`, `legacy-office`, `wps`, `pdf`, `xps`, and `ofd`). The official subpaths supply native-family Wasm binaries; OFD has a standalone Rust/Wasm binary, while WPS and legacy candidates intentionally share the same audited CFB/record binary. The runtime passes immutable input bytes into the parsers and never invokes an external application, network service, or intermediate document format. Native candidates reach the pack only after bounded content probing and an allowed base-parser rejection.

| Format | Implemented native subset | Intentional limits |
| --- | --- | --- |
| ODP / FODP / ODS / ODT / FODT | Existing source-mapped OpenDocument presentation, spreadsheet, and text adapters, including presentation/text flat XML and template aliases | Loaded only through the optional ODF Wasm; format-specific limits remain documented in the detail tables below |
| Pages | Validates the single-file ZIP, Apple IWA/Snappy root identity, metadata plist, saved page hints, body/style tables, inline raster images, bounded table cell text, and embedded vector-PDF drawing content; supported saved-layout packages render searchable/source-mapped pages without raster page substitution | Static subset: complex tables, floating drawings, headers/footers, footnotes, equations, advanced chart labels/effects, advanced wrapping, and drawing effects are omitted or approximate; older packages outside the saved-layout subset retain the bounded root-preview fallback, and directory packages are rejected with an actionable diagnostic |
| Numbers | Validates the single-file ZIP, Apple IWA/Snappy root identity, native sheet/table references, table dimensions, data lists, tile storage, row offsets, and old/new cell storage; renders every decoded table as searchable, source-mapped cells without using the root preview | Native static subset: advanced table styling, formula expressions, merges, custom row/column metrics, shapes, charts, comments, and floating canvas objects remain omitted or approximate; malformed or preview-only Numbers packages are rejected |
| Keynote | Reads the bounded native slide order and static slide components; renders solid slide backgrounds, paragraph-styled text, nested groups, original JPEG/PNG image bytes, and package-local audio/video with an authored poster as searchable/source-mapped objects; falls back per slide to the embedded preview when no supported native drawable is available | Static subset: some shape styling, masks/crops, PDF/vector objects, tables, charts, masters, builds, animations, and transitions are omitted or approximate; media playback is user-controlled and codec support is browser-dependent; fallback thumbnails may be low resolution or absent |
| Binary DOC | Validates CFB/FIB and Word 97+ CLX pieces; resolves font tables, STSH paragraph/character style inheritance, CHPX/PAPX FKP runs, and PCD property modifiers; preserves common fonts, sizes, colors, emphasis, alignment, indents, and spacing with native character/record ranges | Bounded approximate pagination; fields, drawings, headers/footers, tables, uncommon legacy properties, and exact Word line/page layout are not rendered |
| Binary XLS | Validates CFB and BIFF workbook/sheet streams, reads common string/numeric/boolean/error/cached-formula cells and sheet names, and retains row/column/record offsets | Default cell metrics; no formula evaluation, legacy formatting, drawings, charts, comments, controls, or external workbook traversal |
| Binary PPT | Validates CFB, `Current User`, the newest-to-oldest edit chain, merged live persist directories, and live document/slide containers; extracts bounded slide text atoms, common legacy drawing geometry, native record offsets, and shape-bound embedded WAV audio | Legacy layout remains approximate; linked audio/video, embedded video/OLE, animations, and exact PowerPoint layout are not activated |
| WPS / ET / DPS | Dedicated WPS Wasm and package entrypoint; validates the CFB sector graph and reachable `WordDocument`, `Workbook`/`Book`, or `PowerPoint Document` plus `Current User`, then renders the decoded Word, BIFF, or PowerPoint records as source-mapped objects. Unallocated producer trailers are isolated and diagnosed; embedded thumbnails/previews are never a render fallback | The current record coverage is the corresponding Binary DOC/XLS/PPT subset above. ET formatting/drawings/charts and exact WPS Writer/DPS pagination/layout are not yet complete, so a 2% claim is accepted only for a reviewed corpus that passes the WPS Office native visual release suite |
| PDF | Directly parses indirect and compressed object streams, inherited page trees/boxes/rotation, common text operators and ToUnicode `bfchar`/`bfrange` maps, RGB/CMYK/gray vector paths, encoded JPEG/JPEG 2000 image XObjects, and nested Form XObjects; password-protected Standard Security documents using RC4, AES-128, or AES-256 revision 5 are decrypted locally; the object index is retained for the document lifetime, the first page is materialized during open, and later pages append bounded delta snapshots to the existing renderer cache, while document-wide queries intentionally materialize the complete document | Fixed-layout rendering subset: font widths/kerning and rotated text placement are approximate; clipping, shadings/patterns, inline images, raw sampled-image filters, annotations/forms, optional-content groups, signatures, multimedia, and advanced transparency are omitted or diagnosed; public-key handlers and AES-256 revision 6 are rejected explicitly; JavaScript, actions, launch targets, and attachments never execute |
| XPS / OpenXPS | Parses fixed-document sequences, documents, and pages; UTF-8/UTF-16 XML; nested canvases; searchable/source-mapped outline glyphs with authored indices, clusters, bidi and sideways layout; abbreviated and expanded path geometry including arcs; transforms, clips, opacity and opacity masks; complete XPS linear/radial gradient mapping, spread, interpolation, and transform semantics; ImageBrush and VisualBrush viewboxes/viewports, alignment, stretch, all tile/flip modes, transforms, and visual resources; local/remote resource dictionaries; PNG, JPEG, TIFF, and JPEG XR images; sRGB, scRGB, ContextColor/ICC colors, ColorConvertedBitmap color management; and XPS-obfuscated embedded fonts | No-degradation contract: unsupported visual elements fail the open with `FORMAT_INVALID`; they are never omitted, replaced, or reported as a successful approximate render. Print tickets and signatures are non-visual package metadata and never execute |
| OFD 1.0 / 1.1 | Validates the bounded ZIP/XML package, retains the page index, materializes pages on demand, and renders searchable/source-mapped positioned text, browser fonts, raster images, paths and bounded composite graphics, solid/axial/radial paint, adaptively tessellated Gouraud and lattice Gouraud mesh shadings, inherited DrawParam colors/line styles, affine transforms, opacity, geometry clips, background/foreground templates, static annotation appearances, vector OFD seals, and document outlines. Signed references are checked with SM3 and supported SES V1/V4/V5 SM2-with-SM3 signature values are verified locally | Best-effort fixed-layout subset: other patterns, complex or post-1.1 clip transforms, non-OFD seal appearances, attachments, and interactive actions are not complete. Signature validity does not establish certificate trust, revocation, timestamp status, signer identity, or regulatory compliance; those unavailable checks are reported explicitly |

Every limited path emits structured approximate/omitted diagnostics. Encrypted legacy Office payloads and unsupported PDF security handlers are rejected; macros, links, embedded OLE objects, PDF actions, and attachments never execute or activate.

## Implemented P0/P1/P2 scope

The priority labels describe the implemented engineering scope, not a compatibility certification.

- **P0 shared renderer:** affine transforms, rotation and flips, nested groups, z-order and opacity; rectangle, rounded rectangle, ellipse, line, and path geometry; solid, linear-gradient, and radial-gradient paint; shadows and geometry clips; source-mapped rich text.
- **P0 text layout:** per-run typography and color, CJK wrapping, literal bullets/numbering prefixes, tab stops and hanging indents, horizontal LTR/RTL paragraph direction, vertical text, shrink-to-fit, and top/center/bottom anchoring. Full Unicode BiDi, every vertical-writing convention, and application-specific line breaking remain approximate.
- **P0 font resolution:** document-embedded DOCX/PPTX and packaged OpenDocument fonts, static host assets, Engine-cached lazy provider assets, optional exact requested-face `local(...)` probing under `local-first`, deterministic/local-first policy, and explicit serif/sans-serif/monospace approximation with structured diagnostics.
- **P0 presentation depth:** masters/layouts/placeholders, theme colors/fonts, backgrounds, groups, common/custom paths, gradients, shadows, rich text, merged tables, cached charts, common equations, and SmartArt/diagram fallbacks.
- **P1 document depth:** common page flow, sections/columns, headers/footers, lists, tables, images/floats, notes/comments, cached field results, and accepted-revision view.
- **P1 spreadsheet depth:** authored dimensions/styles/number formats, merges, conditional formatting, images, cached charts, frozen panes, data-validation annotations, and sparklines where listed below.
- **P1 reference Viewer:** navigation, search, selectable/copyable text, lazy continuous pages, sheet virtualization and frozen panes, zoom, fullscreen, print, diagnostics, and source inspection.
- **P2 breadth:** OOXML and optional ODF template/slideshow aliases plus bounded CSV and RTF adapters; unsupported formats are rejected rather than misidentified as text.

## Presentation detail

| Capability | PPTX family | ODP family | Important limits |
| --- | --- | --- | --- |
| Page model | Slides, slide size, backgrounds, master/layout inheritance, placeholder replacement | Pages, page size, background styles, default/named/parent style inheritance | Vendor-specific inheritance outside the tested subset may be approximate |
| Shapes and layers | All 187 standard DrawingML preset geometries with guide formulas, authored adjustments, arcs, multi-path coordinate normalization, and per-path fill/darken/lighten/stroke semantics; numeric `custGeom` paths; connectors; groups; affine transforms; rotation/flips; fill/stroke; opacity | Rectangle, ellipse, line, polygon/polyline/SVG path, groups, transforms, fill/stroke, opacity | Vendor-only presets and malformed/unsupported custom geometry still use an explicit visible fallback and diagnostic |
| Theme and effects | Theme color schemes, tint/shade/luminance/alpha transforms, major/minor Latin/EA/CS fonts, solid and linear-gradient fills, nested group/object effect composition, shadows, soft edges, and DrawingML 3D camera/light/backdrop/material/Z/extrusion/contour/bevel/flat-text properties | Default and parent styles, solid/linear-gradient fills, shadows | DrawingML 3D is rendered as a deterministic Canvas 2D projection; it is not pixel-identical to Office's GPU rasterizer, and embedded 3D model parts are not decoded |
| Text | Rich runs, theme fonts/colors, paragraph alignment, paragraph-local line height and before/after spacing, common vertical anchoring | Rich spans, inherited fonts/colors, paragraph alignment, common vertical anchoring | Browser shaping and producer-specific glyph metrics can differ |
| Tables | Table/cell object tree, merges, authored fill/borders/text | Table/cell object tree, row/column spans, authored fill/borders/text | Complex table style bands and every border conflict rule are incomplete |
| Charts | Source-mapped bar, line, area, scatter, radar, pie, and doughnut charts from cached series, including mapped legend labels | Inline or safely embedded bar, line, and circle/pie charts from cached/local table data | No formula recalculation; unsupported or data-less presentation charts use an explicit mapped placeholder |
| Diagrams | Uses the authored diagram drawing part when present; otherwise emits a deterministic source-mapped hierarchy with parent-child connectors from diagram data | No general ODF diagram engine | Data-only hierarchy layout is intentionally approximate |
| Equations | Common OMML fractions, scripts, roots, n-ary structures, delimiters, matrices, functions, accents, bars, limits, group characters, and boxes become readable mapped math text | Common embedded MathML structures become readable mapped math text | This is linearized text, not a full equation-layout engine; complex constructs are diagnosed as approximate |
| Pictures | Embedded and inherited pictures, crop, rotation/flips, SVG preference with local fallback | Embedded/inline pictures, transforms, package-safe lookup | See the codec table below |
| Audio and video | Embedded `media`, `audioFile`, and `videoFile` relationships with authored poster and bounds | Embedded `draw:plugin` media with an optional poster | MP4/M4A, QuickTime, WebM, Ogg, MP3, WAV, and AAC containers are exposed for local controls; codec support is browser-dependent; external media and autoplay are blocked |

Native Keynote exposes package-local audio/video data references with their authored poster and bounds through the same local controls. Binary PPT exposes shape-bound WAV data from its `SoundCollection`; the legacy format's path-based movie and linked-audio records remain blocked as external resources.

Speaker notes expose both a plain-text compatibility view and authored paragraph/run formatting per slide (including their source part); the Inspector renders the formatted note body. Printable notes-page layout is not implemented. Native action links, animations, transitions, timed slideshow media sequencing, embedded OLE/ActiveX, embedded 3D models, full chart styling, and full SmartArt layout semantics are not implemented. Slideshow variants and embedded media do not autoplay.

## Text-document detail

| Capability | DOCX family | ODT family | Important limits |
| --- | --- | --- | --- |
| Text and styles | Paragraphs, rich runs, alignment/indent/spacing, tabs, numbering | Paragraphs, styled spans, alignment/indent/spacing, tabs, list styles | DOCX repeats line breaking with browser-measured face advances; shaping and pagination still approximate a word processor's private engine |
| Page flow | Page size/margins, sections, equal-width columns, explicit/page-before breaks, keep/widow controls | Page size/margins and deterministic page flow | Non-equal DOCX columns and many compatibility-mode pagination rules are incomplete; ODT section columns are not implemented |
| Page backgrounds | Enabled document colors, VML linear and radial gradients, multicolor stops, tiled textures and stretched pictures | Master-page colors and tiled/stretched background images | VML nonlinear interpolation and ODT positioned non-repeating images use diagnosed approximations |
| Headers/footers | Default, first, and even-page stories, including usable images | Master-page header/footer text | Complex fields and arbitrary drawing content in stories remain limited |
| Tables | Grid widths, row/cell objects, horizontal and vertical merges, row-boundary pagination | Repeated rows/cells, row/column spans, authored column widths, cell images, page-local fragments | Nested tables are rejected/omitted with diagnostics |
| Images and floating content | Inline and anchored DrawingML images, common position and wrap modes, header/footer images | Inline/floating frames, images, text frames, common transforms and wrap modes | Complex contour wrapping, legacy VML/WordArt, and generic DOCX text-box drawing layout are incomplete |
| Shapes | Image/drawing bounds and source mapping | Rectangle, ellipse, line, and common custom shape frames | ODT freeform path/polygon/connector/caption objects are not rendered |
| Notes and review | Footnotes, endnotes, and comments in a dedicated source-mapped note flow | Not implemented as a review-flow system | Footnotes are not placed at the page bottom; comments are not rendered as margin balloons |
| Fields/revisions | Visible cached field/TOC results; inserted/move-to content visible and deleted/move-from hidden | Empty ODF fields with `office:date-value`/`time-value`/`string-value` and dynamic `text:page-number`/`text:page-count` in body and master stories; other fields use cached display when present | Fields are not recalculated from formulas; revision markup is not an interactive review UI |

## Spreadsheet detail

| Capability | XLSX family | ODS family | Important limits |
| --- | --- | --- | --- |
| Sheets and cells | Visible-sheet state, used range, hidden rows/columns, strings, numbers, booleans, errors, cached formula values, and bounded read-only calculation when formula caches are absent | Tables, repeated rows/columns/cells, common value types | XLSX calculation covers same-workbook references, names, shared/array formulas, implicit intersection, and common functions; unsupported formulas remain locally diagnosed |
| Worksheet tab colors | RGB, indexed and theme colors with tint, including chartsheets | Table-style tab colors and parent-style inheritance | Viewer preserves authored color as a tab indicator; binary XLS uses its workbook palette fallback |
| Geometry | Authored row heights, column widths, hidden state, merged cells | Authored row/column sizes, hidden state, repeated geometry, merged cells | Very large sheets rely on bounded viewport rendering in the reference Viewer |
| Styling | Fonts, fills, borders, alignment, built-in/common custom number formats | Common cell styles, fonts, fills, borders, alignment, number display | The complete Excel/Calc formatting language is not implemented |
| Conditional formatting | `cellIs`, 2/3-point color scales, data bars, and 3–5 icon sets | Numeric comparisons, 2/3-point color scales, positive/negative data bars | Unsupported rules are diagnosed and omitted |
| Drawings | Authored DrawingML images, common shapes/connectors, groups, transforms, text, chart anchors, and tiled worksheet background pictures beneath the grid and cells | Embedded image frames | Unsupported drawing payloads are omitted with diagnostics |
| Charts | Bar, line, and pie from cached series | Inline or safely embedded bar, line, and pie/circle from local table data | Chart series are not rebuilt from formulas; pivot charts, secondary-axis fidelity, and full chart styling remain limited |
| Frozen panes | Frozen row/column counts and authored pixel extents | Frozen rows/columns from view settings and style geometry | The reference Viewer composites frozen overlays; the SDK exposes metadata and bounded render viewports |
| Print setup | Print area, paper-size code, portrait/landscape orientation, scale/fit settings, margins, first/even/odd header-footer text, and page-layout/page-break-preview view state are retained; browser printing paginates each sheet on cell boundaries | Not implemented | Printing remains read-only and host-controlled; automatic pagination uses common paper-size codes and does not mutate sheet geometry |
| Data validation | List, whole, decimal, date, time, text-length, and custom-rule annotations, including prompt/error metadata | `table:content-validation` read-only annotations with condition and help/error metadata; validity macros/event listeners are blocked | Read-only markers only; rules are not executed or enforced |
| Sparklines | Line, column, and win/loss from bounded same-workbook source ranges | Not implemented | Unsupported groups/references are diagnosed |

Cached pivot result cells remain visible, but pivot rebuilding, pivot charts, slicers, timelines, Power Query/data-model refresh, external-workbook refresh, and interactive filtering/editing are unsupported. Formula calculation never executes macros, active content, external refreshes, or collaborative/editing operations.

## Flat-format detail

| Format | Supported subset | Important limits |
| --- | --- | --- |
| CSV | RFC 4180 quoting, escaped quotes, commas, CR/LF, and embedded newlines; rendered as a source-mapped sheet | Content sniffing requires at least two records with at least two fields each; ambiguous input returns `UNSUPPORTED_FORMAT` |
| RTF | Bounded groups/destinations, Windows-1252 escapes, font/color tables, Unicode, tabs/paragraphs/pages, and common bold/italic/underline/strike/size/color runs | Pictures and embedded objects are omitted/blocked; advanced layout, fields, lists, and drawing objects are unsupported |

Malformed CSV/RTF and resource-limit violations are rejected instead of being rendered as damaged text.

## Reference Viewer surface

The browser-local Viewer under `examples/` is a validation and integration reference, not a published framework component. It supports:

- slide/page navigation and thumbnails, sheet tabs, keyboard unit navigation, and fit/custom zoom;
- text search with result navigation, object-bounds text selection/copy, and safe links for visible text that is itself an `http`, `https`, `mailto`, or `tel` URL;
- lazy continuous mode for text pages, bounded virtual scrolling for large sheets, and authored frozen-row/column overlays;
- fullscreen, current-view printing, diagnostics, object hierarchy, source mapping, and hit inspection.

It does not provide editing, saving, comments/review controls, native document-hyperlink actions, or timed slideshow playback. Embedded presentation audio/video uses direct user-controlled local playback.

## Embedded audio/video

`RenderResult.media` contains only supported, document-local media intersecting the requested viewport, including its object ID, authored bounds/transform, canonical media type, and copied bytes. The Inspector creates and revokes local Blob URLs per displayed frame. The static bitmap continues to contain the authored poster so printing and hosts that omit controls retain a deterministic fallback.

Container signatures are checked for MP4/M4A/QuickTime, WebM, Ogg, MP3, WAV, and AAC. Decoding is delegated to the browser's native media stack, so a recognized container can still be unplayable when its codec is unavailable. External relationships, malformed signatures, unsupported containers, active content, autoplay, and network fallback are never activated.

## Embedded picture codecs

This table applies wherever a format parser above supports picture placement. “Non-approximate” describes the selected codec path, not pixel identity with a particular application, browser color pipeline, or font environment.

| Family | Accepted forms | Fidelity |
| --- | --- | --- |
| Browser-native raster | PNG, JPEG/JPG/JPE/JFIF, GIF, WebP, BMP | Non-approximate for a supported static frame; animated images currently render one frame |
| Office bitmap | DIB and RLE bitmap variants, TIFF/TIF, ICO, PCX | Non-approximate for supported layouts; four-plane PCX color conversion emits an approximate diagnostic |
| Vector | Self-contained SVG | Local rasterization; graphic-only SVG is non-approximate, while SVG text reuses registered embedded/host font binaries and emits an approximate diagnostic. A missing reusable binary font is explicit and can select an Office-provided raster fallback |
| JPEG 2000 | JP2 and raw J2K/JPC codestreams | Non-approximate for unambiguous 8-bit grayscale/sRGB; alpha, raw/ICC, subsampled, signed, high-bit-depth, or four-component color is approximate |
| Windows metafile | EMF, supported EMF+ records inside EMF, WMF | Best effort with an approximate diagnostic; unsupported GDI records can reduce fidelity |
| Compressed metafile | EMZ and WMZ | Bounded gzip expansion followed by the same best-effort EMF/WMF path |
| Disabled legacy | EPS; PICT/PCT/PIC/PCZ | Blocked with a security/fidelity diagnostic; no decoder is invoked |

Picture type is determined from both the relationship/media declaration and payload signature. Misleading extensions are not trusted. External SVG resources are never fetched. Invalid signatures, malformed dimensions, unsupported codec features, and resource-limit violations produce structured diagnostics.

## Compatibility and accuracy policy

- Font selection is face-aware. The default `local-first` policy uses document-embedded, static host, an exact requested `local(...)` face, known platform-compatible local candidates, provider, then explicit approximation. The bounded compatibility allowlist covers common Office CJK, Arabic, Hebrew, Indic, Thai, and Latin families. Missing Cambria faces use the packaged Caladea 1.002 Latin fallback after provider lookup; the four regular/bold/italic faces are loaded only on demand and reported as substitutions. The opt-in `deterministic` policy skips unmanaged local fonts. DOCX obfuscated font parts are deobfuscated only in the document realm; PPTX and packaged OpenDocument font parts use the same document-scoped lifecycle.
- Hosts may provide byte-only local fonts with style, weight, and stretch metadata. Optional local probing is requested-family-scoped and demand-driven through `FontFace` with `local(...)`; the SDK never enumerates installed fonts, calls `queryLocalFonts()`, or treats `FontFaceSet.check()` success as proof that a named font exists.
- A final approximation emits `FONT_SUBSTITUTED` with the authored and fallback families. Missing provider faces, provider failure/timeout, integrity mismatch, unavailable metrics, and font-byte budget omissions remain visible through structured diagnostics.
- External resources are never loaded. Supported parsers report them as `EXTERNAL_RESOURCE_BLOCKED`.
- OOXML and OpenDocument package XML is currently accepted as UTF-8 only; XPS fixed-page XML accepts UTF-8 and UTF-16.
- Complex-script shaping and some text-document pagination depend on browser text behavior and remain approximate.
- An accuracy snapshot with no observable evidence is invalid, scores zero, and cannot create a false pass.
- Embedded previews and thumbnails are fallback diagnostics, not native-format support or fidelity evidence. A format is accepted as natively supported only when its authored structure and content are decoded into independently renderable, searchable, source-mapped objects; preview-only output cannot satisfy compatibility or release gates.
- The external visual-acceptance path uses a closed oracle mapping: Microsoft PowerPoint for PPTX/PPT/ODP, Microsoft Word for DOCX/DOC/ODT, Microsoft Excel for XLSX/XLS/ODS, Apple iWork for Keynote/Pages/Numbers, and WPS Office for WPS/ET/DPS. LibreOffice and non-owning applications are prohibited as native or release acceptance oracles. The path verifies complete units, the fixture/reference/golden/candidate hash chain, declared application/rasterizer fields, and the measured candidate browser/font environment before gating whole-page and local-region similarity. A manual import's reviewer, actual producer, and complete production environment remain human-governed provenance rather than cryptographically proven local facts. Threshold success is evidence only for the reviewed corpus and pinned environments, not universal pixel identity or format certification.

The disabled EPS/PICT policy follows Microsoft's current Office for Windows behavior for [EPS](https://support.microsoft.com/en-us/office/support-for-eps-images-has-been-turned-off-in-office-a069d664-4bcf-415e-a1b5-cbb0c334a840) and [PICT](https://support.microsoft.com/en-us/office/lifecycle/support-for-pict-images-is-being-turned-off-in-office-for-windows).
