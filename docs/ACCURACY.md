# Accuracy test system

Opening a package is only a prerequisite. Accuracy is evaluated from an independent source oracle and a renderer observation at the `@docviewkit/sdk/accuracy` seam. The evaluator is browser-independent and has no runtime dependency; corpus adapters own source extraction, browser capture, and any licensed reference-renderer integration.

## Per-file pipeline

```text
source document
  ├─ independent oracle adapter ─→ expected AccuracySnapshot
  ├─ OfficeViewer adapter ───────→ actual AccuracySnapshot
  ├─ repeated render adapter ────→ determinism runs
  └─ package metamorpher ────────→ four invariant variants
                                      │
                                      ▼
                              evaluateAccuracy(...)
                                      │
                                      ▼
                         one AccuracyReport JSON per file
```

Expected values must come from a source of truth independent of OfficeViewer: authored feature metadata, a worked specification example, a trusted reference application's structured export, or a manually reviewed golden. Do not generate expected geometry, line breaks, object IDs, or pixels with the code under test.

## Observation contract

An `AccuracySnapshot` contains:

- units: page, slide, or sheet index, type, and dimensions;
- objects matched by stable native `sourceKey`, with renderer-local `id`, type, unit, bounds, rotation, normalized crop, z-order, group parent, and source content;
- text line boxes with text range, baseline, font size, line height, overflow, and auto-scale;
- RGBA visual surfaces and their document-space viewport;
- normalized hit probes and front-to-back source-object hits;
- stable structured diagnostics;
- a serializable Display List observation.

`content` stores source-oracle facts, not raster data. Typical values are text, an image SHA-256, table row/column shape and cell source keys, or a cell address/formula/value. Large binary assets stay outside the report.

A snapshot must contain observable evidence: at least one unit, object, visual surface, hit probe, diagnostic, or non-empty Display List. Matching `{ units: [], objects: [], diagnostics: [] }` values are invalid on both sides, produce `EXPECTED_SNAPSHOT_EMPTY` / `ACTUAL_SNAPSHOT_EMPTY`, and force `overallScore` to `0`. A rejected malformed or malicious document may still be a valid diagnostics-only observation.

## Accuracy layers

| Layer | Checks |
| --- | --- |
| Structure and content | Unit count/type, object count/type, text, image digest, table shape, cell address/value, missing and unexpected objects |
| Geometry | Unit size, normalized x/y/width/height, rotation, normalized crop, exact z-order, group parent |
| Text layout | Line count, exact text ranges and wrap positions, normalized baseline, relative font size/line spacing, overflow, auto-scale |
| Visual | Exact RGBA pixel ratio, per-channel tolerance ratio, windowed 8×8 luminance SSIM, and object-region tolerance comparison |
| Object mapping | Center, edge, rotated, Group, overlap, and z-order probes; front-to-back order; every hit resolves to an actual native source key |
| Compatibility | Multiset comparison of code/fidelity/phase/part/source for unsupported, approximate, font substitution, external resource, and active content diagnostics |
| Determinism | Display List, renderer-local object IDs/native mapping, and text layout are identical across repeated runs |
| Metamorphic | Rendering state is unchanged after XML attribute reordering, namespace-prefix renaming, ZIP entry reordering, and irrelevant ZIP-comment metadata changes |

Geometry is compared after normalization:

```text
nx = object.x / unit.width
ny = object.y / unit.height
nw = object.width / unit.width
nh = object.height / unit.height
```

Negative x/y remains valid for off-canvas objects. Invalid dimensions must be rejected by the observation adapter, not silently clamped.

## Defaults and scoring

Default tolerances are deliberately explicit:

| Policy | Default |
| --- | ---: |
| Normalized geometry and baseline | 0.002 |
| Rotation | 0.25° |
| Relative font size, line height, auto-scale | 1% |
| Tolerant pixel delta | 8 per RGBA channel |
| Exact pixel ratio | 1.000 |
| Tolerant pixel ratio | 0.995 |
| SSIM | 0.990 |
| Object-region visual similarity | 0.990 |
| Overall score | 95 |

The exact-pixel metric is always calculated. Its gate can be lowered only for a reviewed platform matrix; tolerant diff and SSIM do not erase the exact-diff evidence.

Applicable layer weights are content 18%, geometry 18%, text layout 15%, visual 18%, object mapping 13%, compatibility 8%, determinism 5%, and metamorphic 5%. A non-applicable layer reports `score: 100, applicable: false` and is excluded from the weighted denominator. A file passes only when no layer emitted a failure and the overall threshold is met.

Pixel goldens must pin browser build, OS image, font files, font loading completion, locale, timezone, color space, device pixel ratio, scale, and background. Store their environment fingerprint with the oracle. Cross-platform runs use separate reviewed baselines; they must not silently widen one global tolerance.

## Corpus governance

Every release manifest must contain all six classes:

1. `minimal`: one source feature per file, including every compatibility Bug reproduction;
2. `combination`: interactions such as rotated text in a Group over an image or merged cells with styles;
3. `enterprise`: authorized, de-identified real business documents from the secure corpus provider;
4. `large`: high page/slide/sheet/object counts and large embedded assets within configured budgets;
5. `malformed`: truncated, inconsistent, duplicate, invalid-relationship, and invalid-XML packages;
6. `malicious`: ZIP bombs, path traversal, DTD/entity, active content, external resources, and resource-limit attacks.

The full required feature list is exported as `REQUIRED_ACCURACY_COVERAGE`. `validateAccuracyCorpus()` fails on a missing class or feature. A manifest declaration is not proof by itself: the suite runner passes each case's declarations to `evaluateAccuracy()`, which verifies them against independent expected observations. For example, `content-image` requires an image object with a digest, `hit-z-order` requires a z-order probe, visual claims require a valid RGBA surface, and determinism claims require at least two meaningful runs. Missing evidence produces `DECLARED_COVERAGE_NOT_OBSERVED`, appears in `coverageValidation.missing`, and forces the overall score to `0`. Licensed enterprise files should remain in an access-controlled CI corpus, while their manifest IDs, hashes, oracle revisions, and report retention policy remain auditable.

For every compatibility Bug:

1. reduce the failing file to a minimal legal or intentionally malformed reproduction;
2. store the file and an independently reviewed oracle;
3. add its automated suite adapter;
4. register the Bug ID in `compatibilityBugs` and on the `minimal` case;
5. run `validateAccuracyCorpus()` in CI.

Missing reproduction, oracle, automation, or a non-minimal regression case is a manifest failure.

## Running a suite

The suite module exports `manifest` and executable `cases`. Each case returns expected and actual snapshots, with optional repeated runs, metamorphic variants, and per-file policy:

```js
export const manifest = {
  compatibilityBugs: ["OV-314"],
  cases: [{
    id: "ov-314-rotated-group.pptx",
    corpusClass: "minimal",
    format: "pptx",
    fixture: "corpus/minimal/ov-314-rotated-group.pptx",
    oracle: "corpus/minimal/ov-314-rotated-group.oracle.json",
    automatedTest: "suites/pptx.mjs#ov-314",
    coverage: [/* declared AccuracyCoverage values */],
    regressionBugId: "OV-314",
  }],
};

export const cases = [{
  id: "ov-314-rotated-group.pptx",
  async observe() {
    return { expected, actual, repeatedRuns, metamorphicVariants };
  },
}];
```

Run:

```sh
npm run accuracy -- suites/release-accuracy.mjs --output output/accuracy
```

The command validates the complete corpus manifest, rejects empty or evidence-free coverage claims, writes `<file-id>.json` for every file plus `summary.json`, and exits nonzero if any file fails. The Node-only metamorphic generator is `scripts/accuracy-metamorphic.mjs`; it reads ordinary non-ZIP64 OOXML/ODF packages, validates CRCs, and emits the four independent variants.

Use `npm run test:accuracy` for the evaluator, runner, mutation generator, and public-SDK metamorphic regression tests.

## Local Excel-calculated XLSX golden

Authorized enterprise workbooks stay outside Git, but can still drive exact cell-display regression. The CBAM 2025 workbook used for the missing-formula-cache regression has source SHA-256 `eb43553472c93f4fb393e2d30e21bc135660ee35bb3c8dae8b1b30fd6eb93b94`. Its reviewed Microsoft Excel full-recalculation copy has SHA-256 `bcc5705cbe92fbcf95995c2bbb42a0ec627410b7edd6a62663b499727baaee04`; both contain 52,673 formula cells, and the Excel copy contains a cached `<v>` element for every formula.

Create the local golden from a copy so the source is never modified:

```sh
cp "$SOURCE_XLSX" output/goldens/cbam-2025-excel-recalculated.xlsx
osascript scripts/recalculate-xlsx-golden.applescript \
  output/goldens/cbam-2025-excel-recalculated.xlsx
```

Then run the exact regression. It compares every non-empty rendered cell string, the 14 visible-sheet state, representative uncached formula output, and the workbook's A4/portrait/margin print metadata:

```sh
OFFICEVIEWER_XLSX_SOURCE="$SOURCE_XLSX" \
OFFICEVIEWER_XLSX_EXCEL_GOLDEN=output/goldens/cbam-2025-excel-recalculated.xlsx \
cargo test --release --features native-formats \
  renders_uncached_cbam_formulas_from_the_supplied_workbook --lib -- --ignored
```

Excel automation forces macros off, suppresses link updates, performs a full calculation rebuild, and restores the user's Excel settings. This local-only oracle supplements rather than replaces the hash-bound browser visual suite.

## External PNG visual oracle

Use `scripts/run-external-visual-suite.mjs` schema v2 for release-grade external visual comparison. The oracle mapping is closed: Microsoft PowerPoint for PPTX/PPT/ODP, Microsoft Word for DOCX/DOC/ODT, Microsoft Excel for XLSX/XLS/ODS, Apple iWork for Keynote/Pages/Numbers, and WPS Office for WPS/ET/DPS. LibreOffice and non-owning applications are prohibited as native or release acceptance oracles; retained research output cannot enter a reviewed golden or suite manifest. Manual PDF/viewport imports require a reviewer declaration bound to source and artifact hashes, but local tooling does not cryptographically authenticate the reviewer, actual producer, or complete production environment. Both captures contain only the read-only document viewport; application chrome, editing placeholders, selection handles, inspectors, and transient loading UI are outside the scoring surface. The complete workflow, locked threshold policy, sheet-range contract, and manifest are specified in [NATIVE-VISUAL-ACCEPTANCE.md](NATIVE-VISUAL-ACCEPTANCE.md).

The older generic schema v1 remains supported for tool-level and research tests, but it is not a native release gate. Its manifest shape is:

```json
{
  "schemaVersion": 1,
  "oracleMode": "read-only",
  "environmentFingerprint": {
    "os": "macOS",
    "osVersion": "15.5",
    "architecture": "arm64",
    "locale": "zh-CN",
    "timezone": "Asia/Shanghai",
    "colorSpace": "srgb",
    "devicePixelRatio": 2,
    "scale": 1,
    "background": "#ffffff",
    "fontSetDigest": "sha256:0000000000000000000000000000000000000000000000000000000000000000",
    "referenceRenderer": "Reviewed Chromium golden",
    "referenceRendererVersion": "Chromium 138.0.7204.50",
    "candidateRenderer": "OfficeViewer Chromium capture",
    "candidateRendererVersion": "git-revision-or-build-id"
  },
  "cases": [{
    "id": "strict-presentation-size-slide1",
    "corpusClass": "minimal",
    "fixture": "strict-presentation-size.pptx",
    "unitIndex": 0,
    "unitType": "slide",
    "goldenPng": "strict-presentation-size-slide1-golden.png",
    "actualPng": "strict-presentation-size-slide1-officeviewer.png"
  }]
}
```

PNG paths are resolved relative to the manifest. The current capture job must write a separate fingerprint JSON with the same fields and values; the runner refuses to read or score images when any field differs:

```sh
node scripts/run-external-visual-suite.mjs \
  tests/goldens/browser/strict-presentation-size-suite.json \
  --fingerprint output/accuracy/current-environment.json \
  --output output/accuracy/external
```

The runner accepts non-interlaced 8-bit grayscale, RGB, indexed, grayscale-alpha, and RGBA PNGs, validates chunk CRCs, and rejects images above 100 million pixels. Each case JSON embeds the standard `AccuracyReport` plus encoded-PNG and decoded-RGBA SHA-256 digests, image dimensions, exact/tolerant differing-pixel counts, exact/tolerant ratios, maximum and mean channel delta, and the exact-difference bounding box. `summary.json` records the fingerprint comparison and one report pointer per case. A dimension mismatch or threshold failure exits nonzero.

## AccuracyReport

Every report includes the required outcomes and the evidence used to explain them:

```json
{
  "fileId": "ov-314-rotated-group.pptx",
  "coverageValidation": {
    "declared": ["structure-units", "geometry-bounds"],
    "observed": ["structure-units", "geometry-bounds"],
    "missing": []
  },
  "contentCompleteness": { "score": 100, "applicable": true, "metrics": {} },
  "geometryAccuracy": { "score": 98.5, "applicable": true, "metrics": {} },
  "textLayoutAccuracy": { "score": 100, "applicable": true, "metrics": {} },
  "visualSimilarity": {
    "score": 96,
    "applicable": true,
    "metrics": {
      "exactPixelSimilarity": 0.97,
      "tolerantPixelSimilarity": 0.998,
      "ssim": 0.995,
      "objectRegionSimilarity": 0.992
    }
  },
  "objectMappingAccuracy": { "score": 100, "applicable": true, "metrics": {} },
  "overallScore": 98.1,
  "passed": false,
  "failureReasons": [{
    "code": "NORMALIZED_BOUNDS_MISMATCH",
    "layer": "geometry",
    "sourceKey": "ppt/slides/slide1.xml#shape:7",
    "message": "Normalized object position or size exceeds tolerance"
  }]
}
```

The report also carries compatibility, determinism, and metamorphic sections so a high visual score cannot conceal a mapping, diagnostic, or nondeterminism failure.
