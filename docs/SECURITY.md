# Security model

Every document and embedded resource is untrusted.

## Isolation

- The default runtime uses one module Worker per document.
- `open()` timeout or cancellation terminates that Worker.
- No document parser or layout loop runs on the host's main thread.
- Non-native Office picture formats, ICO, plus the BMP browser fallback, are decoded in a bounded pool of reusable module Workers. Healthy Workers are recycled after an idle deadline; failed, invalid, and timed-out Workers are terminated.
- Browser-native PNG/JPEG/GIF/WebP/BMP decode remains inside the per-document Worker after preflight validation.
- Inline execution is opt-in and intended only for controlled automated tests.
- Optional ODF, iWork, legacy Office, WPS Office, PDF, XPS, and OFD parsers use the same per-document Worker lifecycle; lazy loading, compilation, and parsing share one absolute open deadline.

Workers reduce the impact of decoder failure and make timeouts enforceable; they are not a separate-origin security sandbox. Deployments must still serve the package from a trusted origin and apply an appropriate `worker-src` Content Security Policy.

Configured text watermarks are presentation metadata painted into returned bitmaps, not access control or DRM. They do not modify the source document, and a determined user who controls the browser runtime can alter client-side code or pixels. Use server-side authorization and contractual controls for document access.

## Package and XML controls

- ZIP central and local headers are cross-checked; the central CRC remains authoritative and is verified against the decoded payload even when a producer left a stale local-header CRC.
- CRC-32, entry ranges, duplicate names, path traversal, encryption, unsupported flags, multi-disk archives, and ZIP64 metadata are validated or rejected. Backslashes are normalized before duplicate-name and traversal checks.
- Stored and raw-Deflate entries are decoded with output limits and one document-wide Deflate operation budget; repeated references do not reset it.
- DTD, general entities, non-UTF-8 encodings, excessive depth, excessive nodes, attributes, and XML size are rejected.
- OPC relationships are resolved within the package root and are edge-limited.
- External relationships are never fetched.
- ZIP64 EOCD records, locators, central fields, local fields, and descriptors are cross-checked and reduced to addressable values before use; existing input, entry-count, expanded-size, compression-ratio, and operation limits still apply. A redundant local ZIP64 size pair is ignored only when it is all-zero or agrees with already validated ordinary sizes and no data descriptor is present.

## Legacy compound-file controls

- OLE CFB header geometry, FAT/DIFAT/miniFAT chains, sector ranges, cycles, declared stream sizes, and directory reachability are validated before a DOC/XLS/PPT main stream is read.
- WPS/ET/DPS use the same CFB graph checks. A referenced partial final sector is accepted only when it contains every directory-declared stream byte and omits padding alone; otherwise it remains a truncation failure. Unreachable trailing bytes can be ignored. Both compatibility paths emit an approximate container diagnostic.
- Only a reachable `WordDocument`, `Workbook`/`Book`, or `PowerPoint Document` stream selects a format. Orphan directory slots cannot trigger dispatch.
- Binary PPT also requires a reachable `Current User` stream. Its current edit chain and persist directories select the live `Document` and `Slide` containers; stale incremental-edit records are never scanned as visible content.
- FIB/CLX, BIFF, and live PowerPoint records are consumed within declared bounds and the normal input, materialized-byte, relationship/entry, object, and render budgets.
- Encrypted/obfuscated payloads are rejected. VBA storage, embedded OLE objects, controls, and external workbook/presentation links are diagnosed but never instantiated, executed, or followed.

## PDF controls

- PDF identification uses a bounded `%PDF-` header scan, never a filename or caller MIME type.
- The Wasm pack parses PDF values, indirect objects, object streams, page trees, resource dictionaries, and page-content operators directly; it does not invoke a browser PDF viewer, Acrobat, Preview, a conversion service, or a third-party PDF engine.
- Object/value depth, dictionary/array size, stream bytes, decoded Flate/ASCII output, page pixels, path commands, Form XObject recursion, and materialized document objects are bounded before rendering.
- FlateDecode validates its zlib wrapper and performs bounded DEFLATE decoding before bytes are consumed. A corrupt Adler-32 checksum is tolerated only when the bounded stream decoded successfully, matching established PDF viewers; unsupported stream/image filters fail explicitly or omit only the identified image feature with a diagnostic.
- Encrypted PDFs are rejected. JavaScript/actions, launch targets, attachments, forms, signatures, optional-content behavior, and multimedia never execute; recognized active/attached content is blocked and diagnosed.

## OFD controls

- OFD is accepted only after bounded ZIP structure identifies `OFD.xml`, the document root, and page content; package paths, CRCs, expanded bytes, XML depth, nodes, attributes, resources, and document objects use the shared hard limits.
- Pages are materialized on demand. A malformed page, template, annotation, image, or embedded font is isolated with a diagnostic when the remaining fixed document can still render.
- Signature containers, signed values, attachments, and actions are never executed. SM3 protected-file digests and supported SES V1/V4/V5 SM2-with-SM3 signature values are checked locally; vector OFD seal appearances are parsed through the same bounded OFD pipeline, and nested seals are not expanded.
- A valid SM2 signature establishes only mathematical integrity for the embedded signer key. The SDK does not establish certificate chains, revocation, timestamp status, signer identity, hardware-key provenance, or regulatory compliance, and reports `OFD_SIGNATURE_CERTIFICATE_TRUST_UNAVAILABLE`. The RustCrypto SM2 implementation used here has not undergone an independent security audit, so this check is not a certified electronic-signature validation service.

## Active content

Macros, scripts, ActiveX, OLE and embedded executable content are never instantiated or executed. They produce `ACTIVE_CONTENT_BLOCKED` diagnostics. ODF `Scripts/` content is treated the same way.

## Embedded picture and media controls

- Declared media types are allowlisted and checked against payload signatures. File extensions alone are never trusted.
- Dimensions and pixel counts are validated before decode whenever the format exposes them. Formats whose final dimensions are only known after decode are checked before their bitmap is accepted and charged to the document-wide image budget.
- SVG is rasterized locally with system-font loading disabled. SVG text receives only successfully registered document-embedded and host-provided font binaries, using the normal embedded-before-host priority. Document-selected external images, fonts, stylesheets, entities, or other dependencies are rejected and never fetched.
- EMZ/WMZ gzip output is bounded before the decompressed EMF/WMF signature is revalidated. Compressed data cannot bypass the expanded-entry or image-pixel limits.
- EMF/EMF+/WMF rendering is best effort and reports approximate fidelity. EMF+ bitmap objects and the classic EMR_BITBLT/EMR_STRETCHDIBITS DIB paths are preflighted against record-local offsets plus cumulative byte and pixel budgets before the converter runs. Nested EMF+ metafile image objects are skipped because the pinned converter does not preserve limits across its recursive path; top-level EMF, WMF, EMZ, and WMZ remain renderable. The converter's WMF path does not consume embedded bitmap or metafile records. No metafile record is executed as native platform code.
- EPS and PICT/PCT/PIC/PCZ are rejected before decode, following the current Office for Windows security posture.
- Embedded audio/video is accepted only from validated internal PPTX/ODP/Keynote package parts or a binary PPT `SoundDataBlob` reached through its shape, external-object, and sound identifiers. Its container signature must match an allowlisted declared or inferred media type; external relationships, legacy path-based movies, and network fallback are blocked. Hosts receive copied bytes and the Inspector uses revocable local Blob URLs with user-initiated controls, never autoplay.
- A host Content Security Policy must allow `media-src blob:` (optionally alongside `'self'`) for these in-memory players; `default-src 'self'` alone blocks Blob media. The bundled Inspector server sets this directive explicitly.
- Media decoding runs in the browser's native media stack. The expanded-entry and total materialized-byte budgets bound the encoded payload before it reaches a media element; media is not passed to the image-codec Worker and cannot select a URL, header, credential, or device.
- Third-party codecs are pinned and bundled locally. The codec path receives only image bytes, bounded reusable font binaries, and numeric limits; it is not supplied a document URL or credentials. The codec built-in fetch is the fixed, same-origin packaged `resvg.wasm` asset; image content cannot select a URL. If the host configures `fontProvider`, its allowlisted endpoint must also be permitted by the deployment's `connect-src` policy.

## Resource budgets

Defaults are lower than immutable hard ceilings. A host may lower them but cannot disable them or configure infinity.

| Budget | Default |
| --- | ---: |
| Input bytes | 128 MiB |
| ZIP entries | 20,000 |
| One expanded entry | 128 MiB |
| Total declared expansion | 512 MiB |
| Compression ratio | 200:1 above 1 MiB |
| XML part | 64 MiB |
| XML depth | 128 |
| XML nodes | 2,000,000 |
| Relationship edges | 50,000 |
| Document objects | 1,000,000 |
| One render | 32 megapixels |
| One decoded image | 32 megapixels |
| All decoded images | 128 megapixels |
| Static, provider, and document-embedded font data | 64 MiB |

The isolated codec pool uses at most eight module Workers. Each request's fixed 9-second deadline covers queueing and decoding; a timed-out or invalid Worker is terminated, while healthy idle Workers are recycled after 30 seconds. The renderer has a 10-second outer image-decode deadline. The 128 MiB expanded-entry limit also bounds one encoded or internally decompressed image. SVG, TIFF, PCX, DIB/RLE, JPEG 2000, and Windows metafiles are not passed directly to the browser's general image decoder.

Static and provider document-font assets are binary values supplied through the host API. A document can request a face but cannot choose a service URL, headers, credentials, or cache key. The Engine batches exact-face requests, applies provider timeout/cancellation, optional SHA-256 verification, in-flight/content deduplication, and an LRU byte budget. The host remains responsible for endpoint allowlisting, authentication, licensing, CORS, and CSP.

The font resolver can fetch four fixed, same-origin packaged Caladea 1.002 files for missing Cambria faces (at most 64 KiB per face); document content cannot provide a font URL.

Embedded fonts are read only from validated internal DOCX/PPTX/OpenDocument package parts; external relationships and unsafe OpenDocument URIs are ignored and never fetched. DOCX font obfuscation is reversed in memory for document viewing, and no decoded font is written to disk. Aggregate encoded bytes are bounded before registration. Browser-measured DOCX metrics use a separate bounded, strictly decoded table of face descriptors, Unicode scalars, and finite advances. Each default-Worker document receives its own static/provider-font copy and owns its extracted embedded fonts, so browser font parsing stays in the disposable document Worker. Successfully registered binary fonts may be cloned into a pooled SVG codec Worker under the same font-byte budget; browser-loaded/local fonts are not exportable and therefore are not available to that isolated rasterizer. Inline execution uses the same budget and `FontFace` loading path, but font parsing occurs in the host realm and remains appropriate only for controlled inputs/tests.

## Reporting a security issue

Report vulnerabilities through GitHub's private vulnerability reporting for [docviewkit/viewer](https://github.com/docviewkit/viewer/security/advisories/new), following the repository's [Security policy](https://github.com/docviewkit/viewer/security/policy). Do not post confidential source documents, credentials or personal data in public Issues.
