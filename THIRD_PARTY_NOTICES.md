# Third-party notices

DocViewKit SDK bundles the following pinned image-codec components, fallback fonts, DrawingML geometry data, and legacy symbol-font mappings. They run locally on bytes already supplied to the engine; none is a hosted service and none receives document data over the network.

| Component | Version | Purpose | License | Upstream source |
| --- | ---: | --- | --- | --- |
| Caladea | 1.002 (crosextra 20130214) | Cambria-compatible Latin fallback | Apache-2.0 | [LibreOffice more fonts](https://cgit.freedesktop.org/libreoffice/core/tree/external/more_fonts/ExternalPackage_crosextra_caladea.mk) |
| `@resvg/resvg-wasm` | 2.6.2 | Deterministic SVG rasterization | MPL-2.0 | [yisibl/resvg-js](https://github.com/yisibl/resvg-js) |
| `emf-converter` | 1.5.0 + OfficeViewer compatibility fixes | Best-effort EMF/EMF+/WMF rendering | Apache-2.0 | [ChristopherVR/emf-converter](https://github.com/ChristopherVR/emf-converter) |
| `jpeg2000` | 1.1.1 | JP2/J2K decoding | Apache-2.0 | [runk/jpeg2000](https://github.com/runk/jpeg2000) |
| `utif2` | 4.1.0 | TIFF decoding | MIT | [photopea/UTIF.js](https://github.com/photopea/UTIF.js) |
| `pako` | 1.0.11 | TIFF support and bounded EMZ/WMZ gzip expansion | MIT and Zlib | [nodeca/pako](https://github.com/nodeca/pako) |
| `mtx-decompressor` | 1.6.0 | Browser-compatible decoding of PPTX Embedded OpenType fonts | MPL-2.0 | [twardoch/mtx-decompressor](https://github.com/twardoch/mtx-decompressor) |
| `encoding_rs` / `cfg-if` | 0.8.35 / 1.0.4 | GBK decoding for predefined Chinese PDF CMaps | MIT and BSD-3-Clause / MIT | [hsivonen/encoding_rs](https://github.com/hsivonen/encoding_rs) |
| `jpeg-decoder` | 0.3.2 | Bounded CMYK/YCCK JPEG decoding for PDF images | MIT or Apache-2.0 | [image-rs/jpeg-decoder](https://github.com/image-rs/jpeg-decoder) |
| Fontations (`skrifa`, `read-fonts`, `font-types`) | 0.44.0 / 0.41.0 / 0.12.2 | Rust-side OpenType/CFF glyph outline parsing for browser-independent PDF math rendering | MIT or Apache-2.0 | [googlefonts/fontations](https://github.com/googlefonts/fontations) |
| Apache POI preset shape definitions | revision `913c7889` | Normative DrawingML preset geometry data compiled into the PPTX parser | Apache-2.0 | [apache/poi](https://github.com/apache/poi/blob/913c78891bd0cd20945b050c63abfb8c66c88009/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml) |
| Unicode Symbol and Wingdings mappings | Adobe table 1.0 and WG2 N4363 | Readable Unicode semantics for legacy symbol-font source bytes | [Unicode Terms of Use](https://www.unicode.org/copyright.html) | [Adobe Symbol mapping](https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt), [WG2 N4363](https://www.unicode.org/wg2/docs/n4363.pdf) |

Package dependency versions are locked in `package-lock.json`; generated geometry and symbol mapping sources carry pinned provenance and SHA-256 metadata. Copyright remains with each upstream author and contributor. The build copies the applicable MPL-2.0, Apache-2.0, MIT, and Zlib license texts into `dist/third-party-licenses/`. Redistribution must retain those texts and notices. This notice does not modify their terms. Original DocViewKit code is licensed under Apache-2.0; see LICENSE and NOTICE. The build-generated inventory in dist/THIRD_PARTY_NOTICES.md also records the Rust Wasm dependency trees, Carlito (OFL-1.1), PDFium standard-font data, and the other bundled codecs. Source-only vendored ironcalc_base retains MIT OR Apache-2.0 with the license text in third_party/licenses/ironcalc-MIT.txt. Test documents and their embedded content require their individual source permissions.
