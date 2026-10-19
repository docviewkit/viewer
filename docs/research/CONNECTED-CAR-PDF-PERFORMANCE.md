# Connected Car PDF opening regression

Measured locally on 2026-10-19 using the supplied `2. Connected Car Service.pdf`.
The original is not checked into the repository.

- SHA-256: `7aaae37a40a621226db3bca18a730fb301418a728bd403c339da8e59f4edaaab`
- 66,470,565 bytes, one A1 page; Illustrator 30.8 export.
- Largest CMYK JPEG: 5,824 × 3,264 pixels, 3,110,981 compressed bytes.
- CPU sampling attributed about 1.58 seconds to the DeviceCMYK interpolation function.

## Fix and semantic boundary

`rgba_bmp` caches exact eight-bit CMYK tuples for one image conversion, using
at most 16,384 entries / 128 KiB. It compares the complete key on every lookup;
collisions replace entries without changing results. Alpha remains per pixel.
Small images use smaller tables; Gray/RGB images allocate no cache. The color
table, floating-point operations, source resolution and image decoding are unchanged.

Both current changes and HEAD were searched for equivalent conversion paths.
PDF raster conversions already share `rgba_bmp` and `cmyk_to_rgb`. PDF vector
colors accept floating-point components and cannot use an eight-bit tuple cache
without quantization. XPS ContextColor carries its own ICC profile to LittleCMS;
it is not equivalent to PDF's fixed DeviceCMYK fallback. Neither is merged into
this raster-specific cache.

## Retained color regression

`tests/fixtures/connected-car-cmyk-colors.bin` contains 4,096 seven-byte records:
four CMYK bytes followed by three expected RGB bytes. They were sampled at
`floor(i * (5824 * 3264 - 1) / 4095)`, for `i = 0..4095`, from the largest
JPEG decoded by jpeg-decoder 0.3.2 and inverted with `255 - sample`. RGB values
were captured with the original conversion before the cache was implemented.
Fixture SHA-256: `009e6118a559e2528b926c8aabf4564ec8ec2916ce9617135f0c36cb199e28f6`.

The fixture has 2,255 distinct tuples and 144 cache-slot collisions. The regression
replays the colors forward, backward, then forward with varying alpha, plus
zero-CMYK initialization and successive single-pixel images. All 389 PDF-feature
tests pass, including the existing JPEG Decode-array inversion and palette tests.

## Reproduction

After building current Core, PDF Wasm and JS, provide the original file as `$PDF`:

```sh
node scripts/run-performance-suite.mjs --samples 3 --max-open-ms 1500 "$PDF"
cargo test --release --no-default-features --features pdf-formats
```

The same 1,500 ms opening budget failed before the change (2,429.6 ms).
Timing is machine-dependent; this is a local regression threshold, not a general
latency guarantee. The user's reported 8–9 seconds was not reproduced at that
exact duration on this machine.

| Measurement | Before | After |
| --- | ---: | ---: |
| Chromium Worker open, median of 3 | 2,416.4 ms | 1,281.8 ms |
| Chromium Worker first frame, median of 3 | 2,657.9 ms | 1,561.9 ms |
| Chromium Viewer cold open, one run | 3,418.3 ms | 2,417.4 ms |
| WebKit Viewer cold open, one run | 4,116 ms | 3,167 ms |

The 1,500 ms budget passes after the fix. Full-page renders at 2,248 × 3,179
retain 1,994 objects. Before/after SHA-256 of decoded RGBA pixels is identical
within each browser (browser engines are compared to their own baselines):

- Chromium: `1fd4afea22cd3ed837dc5a84db2039e7eef252ec54ba5a16f441ecbf0f924c6f`
- WebKit: `e3359f36a124ddd1dc8416cbee8a43a30d770b8e0d2cc98a70cbcac019b82181`

Both real Viewer runs have no page errors or console errors. The rebuilt combined
PDF/iWork Wasm grows by 804 bytes, from 3,147,011 to 3,147,815 bytes. Local evidence
(JSON, CPU profile, full-page PNGs and browser probe) is in `output/connected-car/`.
The 101 package-size and format-pack unit/integration tests also pass:
`node --test tests/core-size.test.mjs tests/format-pack.test.mjs tests/format-pack-e2e.test.mjs`.
Existing unrelated DOCX, protocol and Viewer working-tree edits were preserved;
these measurements do not constitute validation of all those changes.

## Second optimization pass

The next CPU profile still showed about 233 ms in full-file byte searches and
216 ms in the CMYK JPEG conversion path (individual samples vary). PDF recovery,
worksheet XML searches and iWork plist searches now use one exact byte-search
helper backed by the already locked memchr 2.8.3 implementation. PDF keeps its
empty-needle rejection; the XML/plist wrappers retain their former empty-window
panic behavior. Search offsets, binary bytes and first-match semantics are covered
by a shared boundary regression. No raw PDF security scan was removed.

CMYK JPEG Decode-array mappings are precomputed for every channel and all 256
sample values, using the unchanged floating-point expression and rounding. The
1 KiB table replaces repeated arithmetic for every pixel; reversed and custom
Decode arrays continue through the same expression. Source resolution and the
calibrated RGB conversion remain unchanged.

Native regressions: PDF 390 passed; iWork 481 passed; base formats 872 passed,
1 existing ignored. Before this pass, the same supplied PDF failed a 1,000 ms
opening budget (1,530.7 ms during compilation load); quiet-run timing is recorded
separately below so compilation load is not treated as a performance gain.

Quiet consecutive comparison after the complete build:

| Measurement | Previous pass | Second pass |
| --- | ---: | ---: |
| Chromium Worker open, median of 3 | 1,184.0 ms | 865.9 ms |
| Chromium Worker first frame, median of 3 | 1,442.9 ms | 1,163.5 ms |
| Chromium Viewer cold open | 2,156.8 ms | 1,555.5 ms |
| WebKit Viewer cold open, awaited teardown probe | 2,843 ms | 2,596 ms |

The 1,000 ms opening gate passes. Full build and 101 package/format-pack
integration checks pass. Core is 3,837,920 bytes and PDF/iWork is 3,154,935 bytes.
Chromium's complete RGBA hash remains unchanged. Normalized complete Core scenes,
including SHA-256 of every binary resource, compare exactly before/after
(`scene-before.json` and `scene-after.json` in the evidence directory).

WebKit pixel checks have a reproducibility limitation: unchanged-version reruns
also produce different pixel hashes in several image regions. Before and after
runs both produced `ec0950ee01c90f9fc01682f61f94bc1cdbb4ae140188b4599f98bef01ed7e85e`;
the original `e3359f36a124ddd1dc8416cbee8a43a30d770b8e0d2cc98a70cbcac019b82181`
was also observed after the change. These are not grounds to claim universally
stable WebKit pixel equality. The inconsistent probe originally failed to await
Viewer.close(); awaiting it did not eliminate the variability. Console/page errors
are absent in the final runs. WebKit's roughly 2.6-second cold first frame still
needs further work; this pass does not establish sub-two-second opening on WebKit.

## Firefox 157: native-session profiling and BMP fallback fix

The user explicitly authorized Firefox automation for this investigation. The
actual running `/Applications/Firefox.app` (Firefox 157, macOS 27.0.1) was driven
through native UI; no browser was installed. A local Gecko Profiler recording
captured the original PDF. One codec Worker had 2,152 leaf samples in
`maskChannel`, plus per-pixel `decodeDib`, `setPixel` and DataView overhead.
This is the fallback used after native BMP decoding fails on this input.

The shared BMP/DIB decoder now handles standard 32-bit BGRA bitfields by direct
byte reordering. It still validates headers, pixel ranges and masks; top-down and
bottom-up row order are preserved. Nonstandard masks use the original decoder.
ICO still applies its AND mask and all-zero-alpha fallback after the shared loop.
The same decoder serves BMP, DIB, ICO and embedded metafile DIBs; no PDF-specific
parallel decoder was added.

Real-file decoded BMP: 76,038,214 bytes, 5,824 × 3,264 pixels. Three local Node
codec runs before: 1,047.7 / 1,060.8 / 1,039.4 ms; after: 45.7 / 43.5 / 39.0 ms.
All 76,038,144 decoded RGBA bytes retain SHA-256
`c021024581a813bd51cf5a056a61a3a845a97b1ac7c50279f43f957e1950d741`.

A temporary local test page first opened `visual-baseline.pptx`, then opened the
original PDF twice without refreshing, timing `await viewer.open(File)` through
first frame. The page and file were identical for both versions, in the user's
running Firefox; no profiler ran during these measurements:

| Actual Firefox Viewer | Before | After |
| --- | ---: | ---: |
| First PDF open after the default PPTX | 5,322 ms | 2,022 ms |
| Reopening PDF without refreshing | 4,985 ms | 1,657 ms |

The first PDF-open test includes optional pack initialization; subsequent opens
reuse the engine. Timing excludes test-fixture HTTP download (the local file was
already read before clicking Open), consistently before and after. Browser assets
were rebuilt normally using `npm run build:js`. Local profile, original-image
benchmark and regression logs are in `output/playwright/firefox-pdf/`.

Twenty image-codec tests pass, including BMP/DIB orientation and unusual-mask
boundaries, alpha endpoints, ICO AND masks and ICO all-zero-alpha handling.

Final verification: 195 codec/display-list tests and the image identification /
codec-worker bundle checks pass. Temporary 63 MiB PDF and local measurement-page
copies were removed from tests/fixtures; the user's original Downloads PDF was
not modified. The profiler preset was restored and Firefox returned to the
user's original Viewer tab. Refresh that tab before manually retrying so its
already-loaded Worker scripts are replaced with the rebuilt version.

## Bounded renderer image bucket fingerprints

The native Firefox recording also attributed 514 leaf samples to the bundled
`byteFingerprint` function (`runtime-ACXH3LBG.js`, line 1). It traversed every byte
of the 76 MB bitmap just to choose its image-cache bucket. Renderer bucket keys
now sample at most 4,096 evenly spaced bytes. Existing full-byte hashing remains
the default for protocol callers; renderer cache candidates still require object
identity or exact full-byte equality. This changes bucket placement, not image
content or the criterion for sharing a decoded bitmap.

A forced-collision regression supplies two same-length images whose sampled
bytes agree but whose unsampled byte differs. The equal copies share one decode;
the distinct image decodes separately and draws separately. Combined protocol,
image, codec and display-list regressions: 247 passed. No parallel format-specific
fingerprint implementation was added.

Using the same temporary page, original File and running Firefox 157 as above:

| Viewer first-frame measurement | Previous version | Bounded buckets |
| --- | ---: | ---: |
| First PDF open after default PPTX | 2,022 ms | 1,526 ms |
| Reopening without refreshing | 1,657 ms | 1,158 ms |

These are Viewer measurements, not a controlled comparison with Firefox's built-in
PDF viewer. The built-in viewer was inspected on the same original local PDF,
but no quantitatively comparable first-frame timing was obtained, so parity or
superiority is not claimed. This change does not alter source resolution,
CMYK conversion, alpha, bitmap decoding, drawing or text layout.

## Native-viewer presentation timing and rejected interpolation experiment

A new local recording of Firefox's built-in reader on the original file was
saved as `output/playwright/firefox-pdf/native-profile.json.gz`. Relative to
navigation start, its first contentful paint is 39.2 ms; the LCP marker at 70 ms
identifies `span#numPages`, not the PDF canvas. Actual `paintImageXObject` samples
occur from 1,070 to 1,407 ms, with image creation in the PDF worker through
1,402 ms. Document load completes at 1,410.6 ms. Sampling does not give a precise
full-page completion event, so these numbers are not a matched first-frame
benchmark. Its default Automatic Zoom also differs from our fit-to-view setting.
The early toolbar paint must not be presented as PDF page rendering latency.

The current Viewer recording (`current-profile.json.gz`) still identifies CMYK
interpolation and JPEG decoding as major worker costs: nearest-script sample
counts include 372 for `wasm-function[494]` and 281 for `[592]`. A subsequent
unprofiled repeat-open measurement was 1,171 ms.

An experiment reused identical C/M interpolation prefixes across Y/K corners,
keeping floating-point multiplication and corner accumulation order unchanged.
All 390 PDF tests passed, as did 4,096 source-image color comparisons and 100,000
float-bit comparisons against the original function. Full original-document scene
hash and resource hash stayed identical. Node Core opening improved only slightly
(about 797 -> 774 ms warm in one pair), while actual Firefox measured 1,609 ms on
first open and 1,364 ms on repeat, without a profiler. This failed to demonstrate
a cross-browser end-to-end benefit, so the source change and Wasm were restored.
After restoration, the same running Firefox measured 1,521 ms first open and
1,163 ms repeat, confirming the stable baseline. The experiment is not retained
as a product optimization. No fidelity or
resolution changes were made. Remaining perceived-speed work must measure first
real page content separately from complete first frame, including cold-start
pack costs; current evidence does not establish parity with the built-in reader.
