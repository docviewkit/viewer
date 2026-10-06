# py-pdf sample-files corpus

OfficeViewer uses the upstream [`py-pdf/sample-files`](https://github.com/py-pdf/sample-files)
repository as a PDF compatibility corpus. The suite is pinned to commit
`818dc013ad1f537198e9fcfae8a6b0dffe25ffa3`; every selected PDF also has a
byte length and SHA-256 recorded in
[`tests/corpus/py-pdf-sample-files.json`](../../tests/corpus/py-pdf-sample-files.json).
It never follows the upstream default branch implicitly.

The upstream files are licensed under
[CC BY-SA 4.0](https://github.com/py-pdf/sample-files/blob/818dc013ad1f537198e9fcfae8a6b0dffe25ffa3/LICENSE).
They are downloaded into `.cache/pdf-corpus/` for testing and are not bundled
in the npm package or release artifacts.

## Suites

```sh
# 11 representative files: basic parsing, password protection, images,
# forms, Arabic text, malformed metadata, CMYK, annotations, and transforms
npm run test:pdf-corpus

# All 34 PDFs and all 293 pages
npm run test:pdf-corpus:full

# Selected pages compared with macOS PDFKit
npm run visual:pdf-corpus

# All 293 pages compared with macOS PDFKit; no rejected or sub-99% page passes
npm run visual:pdf-corpus:full
```

The parser suites verify the pinned checkout, file hashes and lengths, PDF
identification, password-required behavior, page counts, lazy page
materialization, object geometry/source mapping, structured diagnostics, and
explicit baselines for known rejections.
Their machine-readable result is `output/pdf-corpus-report.json`.

The visual suite reuses the PDF Preview harness and writes comparisons under
`output/pdf-corpus-visual/`. It checks compact image, Arabic, CMYK, transform,
image-reference, and GeoTopo samples. The 117-page GeoTopo document samples
only its first, middle, and last pages. Each case must either pass the standard
visual thresholds or stay above its recorded compatibility floor, so known
fidelity gaps remain visible without making the suite permanently red.
Its parser baseline also asserts representative Type 1 math glyphs on page 59
and rejects leaked control characters.

The full visual suite accounts for every manifest page. Pages that the engine
cannot open remain explicit failures, and every rendered page must reach at
least `0.99` tolerant pixel similarity. Cross-raster comparison keeps the
channel tolerance at 8 and uses a symmetric 3-pixel neighborhood so differences
in PDFKit and Chromium edge antialiasing do not conceal real missing content,
geometry, or color errors. Its machine-readable result is
`output/pdf-corpus-visual-full/full-corpus-summary.json`.

This corpus is broad interoperability evidence, not a PDF specification
certification, security/malicious-input suite, or proof that OfficeViewer
exposes every form, attachment, outline, or annotation feature. Those public
capabilities still require focused semantic assertions. PDFKit is an
independent visual reference for this PDF research suite, not an oracle for
Office or iWork formats.

To update the corpus, review the upstream license and changes, choose a new
commit deliberately, regenerate every size and SHA-256 in the manifest, and
run both the full parser suite and selected visual suite.
