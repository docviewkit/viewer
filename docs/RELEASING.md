# Release pipeline

Pushing an annotated `vX.Y.Z` tag starts `.github/workflows/release.yml`.
The tag version must match `package.json`, `package-lock.json`, `Cargo.toml`,
and `Cargo.lock`.

Ordinary branch pushes and pull requests do not start automated builds. The
full `.github/workflows/ci.yml` gate remains available through manual dispatch;
automatic compilation is reserved for `v*` release tags.

The workflow:

1. runs the JavaScript, package, Rust, and native release gates;
2. builds the Apache-2.0 SDK, public Viewer package and Pages demo from the
   same commit, then uploads them to a draft release;
3. dispatches `publish.yml` to verify checksums, publish `@docviewkit/viewer`,
   deploy Pages and publish the GitHub Release.

## Source repository configuration

The canonical source and release repository is `docviewkit/viewer`. Its
`GITHUB_TOKEN` creates one draft containing all artifacts and triggers the
existing `publish.yml` with `repository_dispatch` (`release-ready`),
which retains the npm Trusted Publishing identity. No cross-repository token
or GitHub App is needed. Release jobs run only in `docviewkit/viewer`;
keep version tags attached to the actual source commit.

Commit `.github/workflows/ci.yml`, `release.yml` and `publish.yml` together
with the source on the default `main` branch before pushing a new version tag.
`repository_dispatch` uses the default branch's publishing workflow. Import
source through a normal commit or merge that preserves the public repository's
history. Existing binary-only release tags must not be pushed again or moved.

Repository Settings must keep Actions enabled and Pages configured to deploy
from GitHub Actions. Keep the `npm-production` and `github-pages` environments.
Workflow files grant the required write permissions explicitly; the repository's
default token permission can remain read-only. Old `DOCVIEWKIT_APP_ID`,
`DOCVIEWKIT_APP_PRIVATE_KEY` and `DOCVIEWKIT_RELEASE_TOKEN` secrets are no longer
used by this pipeline.

## Independent official website

The official website, online documentation and Demo source live in
[docviewkit/website](https://github.com/docviewkit/website). That repository
installs an exact, locked `@docviewkit/viewer` npm version and deploys tested
website commits independently. It does not compile Core or Wasm. Publishing a
Viewer version does not depend on the website build or hosting availability.

Website production SSH secrets belong only to the website repository's
`website-production` environment. This core repository does not package or
deploy the website. Keep the existing npm Trusted Publishing identity here.

The website reports its source commit at `/site-version.json`. Its Demo and
online documentation report the installed Viewer version, which may differ
from the website's own version. Dependabot proposes Viewer patch updates in
the website repository; required website and browser checks gate their merge
and deployment. Major and minor updates are reviewed separately.

## npm and Pages configuration

The public workflow is
`docviewkit/viewer/.github/workflows/publish.yml`.

The existing `@docviewkit/viewer` package uses npm Trusted Publishing. Preserve
and verify its configured identity:

- Organization or user: `docviewkit`
- Repository: `viewer`
- Workflow: `publish.yml`
- Environment: `npm-production`
- Allowed action: direct `npm publish`

Keep `publish.yml` and `npm-production` named exactly as configured. OIDC needs
`id-token: write`, a GitHub-hosted runner, npm >=11.5.1 and Node >=22.14.0;
the workflow uses Node 24. It continues to publish the generated
`@docviewkit/viewer` tarball rather than the private root workspace package.
Source migration requires no new npm token. Only initial publication of a new
package needs the optional `NPM_TOKEN` fallback before setting up OIDC.
See [npm Trusted Publishing](https://docs.npmjs.com/trusted-publishers/) and
[GitHub workflow triggers](https://docs.github.com/en/actions/how-tos/write-workflows/choose-when-workflows-run/trigger-a-workflow).

## Open source license

Viewer, Engine and project source use Apache-2.0. Every package and Pages runtime
includes LICENSE, NOTICE and the third-party licenses. The independent website
preserves these files from its installed npm dependency.
The public package preserves its Viewer entry and adds `/engine`, `/viewer`
and `/accuracy` exports. Runtime licenses and mandatory branding are removed.
Support, customization and enterprise delivery are arranged by email;
the website has no customer accounts or database and signing keys are never bundled.
Root `private: true` prevents accidental npm publication of the build workspace;
it does not restrict rights under Apache-2.0.

Before importing source into the existing public repository, review Git history
for credentials and confidential content, and review external test documents
for redistribution rights. Preserve existing public Releases and Issues; do not
force-push over public history or move old version tags.

## Retry

If a downstream service fails after artifacts have been built, run the source
workflow manually with the existing source tag. Draft assets can be replaced
until publication completes, and an already published npm version is not
republished. Once the GitHub Release is public, a retry leaves its notes and
attachments intact. Website deployments are retried in the website repository
using the tested website commit.

## Local development

Requirements: Node.js 24 (the release CI runtime), npm, and a rustup toolchain with `wasm32-unknown-unknown`.

```sh
git clone https://github.com/docviewkit/viewer.git
cd viewer
npm ci
rustup target add wasm32-unknown-unknown
npm test
```

`npm run build` produces ESM declarations, source maps, the document and image-codec Worker modules, the base Wasm, the optional native format-pack Wasm binaries (with iWork and PDF sharing one binary), and the required local font assets.

The root `@docviewkit/sdk` package is a private build workspace to prevent accidental npm publication. Public `@docviewkit/viewer` artifacts are generated by `scripts/prepare-release.mjs`; the workspace flag does not restrict Apache-2.0 source usage. See the release pipeline above and [website source and deployment instructions](https://github.com/docviewkit/website).

Contributions must include a real case that fails before the change and passes afterward, preserve existing passing cases, and keep equivalent format behavior in the shared implementation. Rendering changes also need actual browser evidence for the affected behavior; corpus documents require permission for public redistribution. Follow [AGENTS.md](../AGENTS.md) and the [accuracy policy](ACCURACY.md).

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

The lower-level `npm run inspect` command starts an already-built runtime without process management. Open [http://127.0.0.1:4173/](http://127.0.0.1:4173/). The page accepts the [supported formats](SUPPORT.md). It provides slide/page navigation and thumbnails, sheet tabs with virtual scrolling and frozen panes, lazy continuous text pages, search, selectable/copyable text, URL-text safe links, zoom, fullscreen, printing, and document/render diagnostics with hit-tested native source mappings. It never uploads document bytes and does not edit or save documents.

### Official-source corpus validation

The fast corpus check downloads nine pinned OOXML/ODF fixtures, verifies their SHA-256 values, and exercises identification, parsing, diagnostics, source mapping, hit testing, and a bounded-rejection case:

```sh
npm run test:corpus
```

The full check uses sparse checkouts pinned to Microsoft Open XML SDK, OASIS ODF TC, and TDF ODF Toolkit commits:

```sh
npm run test:corpus:full
```

Assets stay in `.cache/official-corpus/`; the machine-readable report is written to `output/official-corpus-report.json`. See [the validation baseline](research/CORPUS-VALIDATION.md), [OOXML source research](research/OOXML-TEST-ASSETS.md), and [ODF source research](research/ODF-TEST-ASSETS.md). These upstream assets do not provide a six-format visual oracle and do not constitute ISO, ECMA, or OASIS certification.

PDF compatibility has a separate pinned, hash-verified
[`py-pdf/sample-files` corpus](research/PY-PDF-SAMPLE-FILES.md):

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

Architecture details live in [docs/ARCHITECTURE.md](ARCHITECTURE.md).

### Accuracy testing

The separate `@docviewkit/viewer/accuracy` test module evaluates independent source-oracle and renderer snapshots across structure/content, normalized geometry, text layout, exact and tolerant pixels, windowed SSIM, object regions, hit/source mapping, compatibility diagnostics, determinism, and package metamorphism. It emits one `AccuracyReport` per file; a high whole-page visual score cannot conceal missing objects or incorrect mappings.

```sh
npm run test:accuracy
npm run accuracy -- suites/release-accuracy.mjs --output output/accuracy
```

Release corpus manifests must cover minimal single-feature, combination, authorized real enterprise, large, malformed, and malicious documents. Every compatibility Bug must register a minimal reproduction, independent oracle, and automated test. See [docs/ACCURACY.md](ACCURACY.md).

External visual acceptance uses Microsoft Office as the only golden for PPTX/PPT/ODP, DOCX/DOC/ODT, and XLSX/XLS/ODS; Apple iWork for Keynote/Pages/Numbers; and WPS Office for WPS/ET/DPS. LibreOffice and unrelated applications are prohibited as native or release acceptance oracles. `visual:native:reference` records reviewed native goldens, and `visual:native:capture` captures every OfficeViewer unit in Chromium under the locked `native-visual-v1` policy. Its 99.5% tolerant-pixel and 99% SSIM/object-region gates are stricter than a 2% visual-difference target. A release suite must cover every format and trusted corpus class, verify the fixture/reference/golden/candidate hash chain and pinned machine-checkable fields, and reject missing pages, slides, sheets, or candidate-environment drift. Manual imports additionally require a hash-bound human review attestation; local tooling does not cryptographically authenticate the reviewer or producer. See [External visual acceptance](NATIVE-VISUAL-ACCEPTANCE.md).
