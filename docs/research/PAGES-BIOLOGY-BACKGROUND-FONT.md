# Pages biology: section fills and missing display font

The supplied three-page `Untitled.pages` is retained byte-for-byte as
`tests/fixtures/pages-biology.pages`, SHA-256
`ef8b161ca75d17b8f14a5c566bc8456c8053fc6ef1b852f5d379de3dddef1089`.
It is a different document from the earlier two-page brochure with the same
Downloads filename.

Native Pages shows dark blue on the first two pages, a white third page, and a
narrow uppercase cover title on four lines: THE / STUDY / OF CELL / BIOLOGY.
The first two section archives (208428, 208254, message type 10011) store a TSD
FillArchive in field 30, with RGB values rounding to #383e4e. Section 208399 has
no fill. The parser previously ignored that field. Pages now reuses the existing
Keynote/Numbers fill decoder and emits a full-page shape below section artwork.
No fill is inherited from the preceding section. Malformed optional fills retain
a diagnostic; safety/resource failures still propagate. Page backgrounds have a
Pages-specific source locator and use the existing shared shape/paint protocol.
This does not introduce another fill parser or use packaged preview images.

The source title requests BebasNeue, 128 pt, bold. That font is not available to
the tested browsers. Default local-first resolution previously selected the broad
Calibri fallback, breaking STUDY and BIOLOGY across lines. The existing shared
local compatibility allowlist now tries the spaced/unspaced original face name,
then DINCondensed-Bold and ArialNarrow-Bold. Exact fonts retain priority, and
opt-in deterministic resolution still skips unmanaged local fonts. No font bytes
were added. The tested machine resolves the title to DIN Condensed; this restores
the narrow appearance and four complete lines, but is a substitute rather than
an exact Bebas Neue glyph match. A host/provider can supply the original face.
AvenirNext body faces already resolved to browser fonts before this change and
remain unchanged.

Validation:

- The unchanged-file Rust test fails before the fix (zero backgrounds instead
  of two), passes after, and checks bounds, color, page ownership and z order.
- The font regression fails before the allowlist change; after it, exact-face
  priority, narrow substitution and deterministic-policy boundaries pass.
- `node scripts/test-pages-biology.mjs` opens the real Viewer and renders the
  original document through the worker engine in Chromium and WebKit. It checks
  all three page corner colors, title line breaks, Avenir face resolutions and
  empty error consoles. The old JS and Wasm fail in both browsers, with white
  backgrounds and split title words. Captures/results/logs are under
  `output/playwright/pages-biology/`, including `baseline/`.
- `node scripts/test-pages-brochure.mjs` passes in both browsers.
- All 123 iWork module tests and 93 font/protocol tests pass. The existing Pages,
  Keynote indent and Numbers scene hashes remain unchanged.
- The optimized PDF/iWork pack is 3,220,624 bytes (+660 from 3,219,964), within the
  existing 3,223,000-byte budget.

## Firefox follow-up: face names, family names and synthetic weight

The same fixture still failed in the user's installed Firefox 158 on macOS after
`d7936af`. This was reproduced in that browser with the unchanged fixture (the
Downloads file and fixture hashes match). A local harness used the actual worker
engine and then opened the same bytes in the real Viewer.

Two shared font-resolution defects were confirmed:

1. `local("DINCondensed-Bold")` loads successfully. A registered alias measures
   the title probe at 634.33 px (72 px font, weight 400), but CSS selection by the
   PostScript name `DINCondensed-Bold` measures 1072.20 px in the worker, the
   missing-name fallback. CSS family `DIN Condensed` measures 634.33 px. The old
   compatibility check required the face name to also be a measurable CSS family,
   so it removed the loaded usable alias and selected Calibri. Loaded aliases
   now survive when the source name is not CSS-selectable. An alias must still
   pass the existing non-fallback measurement probe; when the source CSS name
   is measurable, different alias/source metrics are still rejected.
2. Compatible local faces were registered with weight 400 unconditionally.
   The named `-Bold` face already has bold outlines. Firefox synthesizes bold
   again when requesting weight 700 from that registration, changing the probe
   width from 634.33 to 671.83 px. Fixing only the first defect leaves BIOLOGY
   split into BIOLOG / Y. Named bold compatibility faces are now registered as
   weight 700, and validation uses those same descriptors. Other compatibility
   faces retain weight 400. Exact and document-declared faces retain their
   authored descriptors. No font assets or browser settings were changed.

Before: THE / STUD / Y OF / CELL / BIOL / OGY; font source=fallback, family=Calibri.
After: THE / STUDY / OF CELL / BIOLOGY; source=browser, using the loaded DIN alias.
The title remains a DIN Condensed substitute for the unavailable Bebas Neue.

Evidence is under `output/playwright/pages-biology/firefox-native/`: before/after
JSON, native rendered page PNGs, `after-viewer.png`, the intermediate one-defect
result, and before/intermediate/after unit logs. After checks cover the four
complete lines, three page background colors, Viewer ready state, and no captured
runtime errors. The regression extracted from the real file fails on the old
resolver, also fails on the intermediate synthetic-weight defect, and passes on
the completed fix. Existing loaded-fallback rejection, measured alias/source
mismatch rejection, exact-font priority, deterministic-policy and teardown tests
remain passing (94 font/protocol tests total). Biology and brochure scripts pass
in Chromium and WebKit. This JS-only change does not alter Core scenes or Wasm.

The focused browser script retains Chromium/WebKit as its default. Firefox is
explicitly opt-in for this user-requested investigation:

```sh
PAGES_BIOLOGY_FIREFOX=1 node scripts/test-pages-biology.mjs
```

The cached Playwright Firefox builds could not launch in this desktop execution
environment (macOS process/sandbox launch errors, before any document test).
Consequently the successful Firefox proof above comes from the user's running
Firefox 158, driven through the native UI and the local harness, not a claimed
successful Playwright Firefox run. It used the same application assets and
worker engine. The opt-in script is available for hosts where Firefox launches.
