import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, isAbsolute, resolve, sep } from "node:path";
import { platform } from "node:os";
import { fileURLToPath } from "node:url";

import { createOfficeEngine } from "../dist/engine.js";
import { extendedFormatPack } from "../dist/extended-formats.js";
import { OfficeEngineError } from "../dist/types.js";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const manifestPath = resolve(root, "tests/corpus/py-pdf-sample-files.json");
const checkout = resolve(root, ".cache/pdf-corpus/py-pdf-sample-files");
const reportPath = resolve(root, "output/pdf-corpus-report.json");
const visualRoot = resolve(root, "output/pdf-corpus-visual");
const fullVisualRoot = resolve(root, "output/pdf-corpus-visual-full");
const arguments_ = process.argv.slice(2);
const full = arguments_.includes("--full");
const visual = arguments_.includes("--visual");
const visualFull = arguments_.includes("--visual-full");
const quiet = arguments_.includes("--quiet");
const knownArguments = ["--full", "--visual", "--visual-full", "--quiet"];
const unknown = arguments_.filter((argument) => !knownArguments.includes(argument));

if (unknown.length !== 0) throw new Error(`Unknown argument(s): ${unknown.join(", ")}`);
if ([full, visual, visualFull].filter(Boolean).length > 1) {
  throw new Error("--full, --visual, and --visual-full are separate suites");
}

const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
validateManifest(manifest);
const selected = full ? manifest.samples : manifest.samples.filter(({ smoke }) => smoke);

function digest(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: root,
    encoding: "utf8",
    stdio: options.inherit ? "inherit" : ["ignore", "pipe", "pipe"],
  });
  if (result.status !== 0) {
    if (options.allowFailure) return undefined;
    const detail = options.inherit ? "" : `: ${(result.stderr || result.stdout).trim()}`;
    throw new Error(`${command} ${args.join(" ")} failed${detail}`);
  }
  return options.inherit ? "" : result.stdout.trim();
}

function validateManifest(value) {
  if (value?.schemaVersion !== 1 || typeof value.suite !== "object" || !Array.isArray(value.samples)) {
    throw new Error("Invalid PDF corpus manifest");
  }
  if (!/^[0-9a-f]{40}$/u.test(value.suite.commit)
      || value.suite.license !== "CC-BY-SA-4.0"
      || value.samples.length === 0) {
    throw new Error("PDF corpus suite metadata is incomplete");
  }
  const paths = new Set();
  for (const sample of value.samples) {
    const normalized = sample.path.split("/").join(sep);
    if (typeof sample.path !== "string" || isAbsolute(normalized)
        || normalized.split(sep).includes("..") || !sample.path.endsWith(".pdf")
        || paths.has(sample.path)) {
      throw new Error(`Unsafe or duplicate PDF path: ${sample.path}`);
    }
    paths.add(sample.path);
    if (!/^[0-9a-f]{64}$/u.test(sample.sha256)
        || !Number.isSafeInteger(sample.bytes) || sample.bytes <= 0 || sample.bytes > 16 * 1024 * 1024
        || !Number.isSafeInteger(sample.pages) || sample.pages <= 0
        || typeof sample.encrypted !== "boolean" || typeof sample.smoke !== "boolean") {
      throw new Error(`Invalid PDF metadata: ${sample.path}`);
    }
    const expectation = sample.expectation ?? { status: "opened" };
    if (!["opened", "rejected"].includes(expectation.status)
        || (expectation.status === "rejected" && typeof expectation.code !== "string")) {
      throw new Error(`Invalid PDF expectation: ${sample.path}`);
    }
    for (const probe of sample.textProbes ?? []) {
      if (!Number.isSafeInteger(probe.page) || probe.page < 1 || probe.page > sample.pages
          || !Array.isArray(probe.includes) || probe.includes.some((text) => typeof text !== "string")
          || typeof probe.excludesControlCharacters !== "boolean") {
        throw new Error(`Invalid PDF text probe: ${sample.path}`);
      }
    }
  }
  if (!value.samples.some(({ smoke }) => smoke)) throw new Error("PDF corpus has no smoke samples");
}

function prepareCheckout() {
  if (existsSync(resolve(checkout, ".git"))) {
    const head = run("git", ["-C", checkout, "rev-parse", "HEAD"], { allowFailure: true });
    const dirty = run("git", [
      "-C", checkout, "status", "--porcelain", "--untracked-files=no",
    ]);
    if (head === manifest.suite.commit && dirty === "") return;
  }

  rmSync(checkout, { recursive: true, force: true });
  mkdirSync(checkout, { recursive: true });
  run("git", ["-C", checkout, "init", "--quiet"]);
  run("git", ["-C", checkout, "remote", "add", "origin", manifest.suite.repository]);
  run("git", ["-C", checkout, "fetch", "--quiet", "--depth", "1", "origin", manifest.suite.commit]);
  run("git", ["-C", checkout, "checkout", "--quiet", "--detach", "FETCH_HEAD"]);
  if (run("git", ["-C", checkout, "rev-parse", "HEAD"]) !== manifest.suite.commit) {
    throw new Error("PDF corpus checkout does not match its pinned commit");
  }
}

function verifiedPath(sample) {
  const path = resolve(checkout, sample.path);
  if (!path.startsWith(`${checkout}${sep}`)) throw new Error(`PDF escaped checkout: ${sample.path}`);
  const bytes = readFileSync(path);
  if (bytes.length !== sample.bytes) {
    throw new Error(`${sample.path} has ${bytes.length} bytes; expected ${sample.bytes}`);
  }
  const sha256 = digest(bytes);
  if (sha256 !== sample.sha256) {
    throw new Error(`${sample.path} has SHA-256 ${sha256}; expected ${sample.sha256}`);
  }
  return { path, bytes };
}

function validateObject(object) {
  const bounds = Object.values(object.bounds);
  if (bounds.some((value) => !Number.isFinite(value))
      || object.bounds.width < 0 || object.bounds.height < 0
      || object.source.format !== "pdf") {
    throw new Error(`${object.id} has invalid PDF geometry or source mapping`);
  }
}

async function openSample(engine, sample) {
  const started = performance.now();
  let document;
  try {
    const { bytes } = verifiedPath(sample);
    if (sample.encrypted) {
      try {
        const unexpected = await engine.open(bytes);
        unexpected.close();
        throw new Error("encrypted PDF opened without a password");
      } catch (cause) {
        if (cause?.code !== "PDF_PASSWORD_REQUIRED") throw cause;
      }
      document = await engine.open(bytes, { password: "openpassword" });
    } else {
      document = await engine.open(bytes);
    }
    if (document.info.format !== "pdf" || document.info.kind !== "text") {
      throw new Error(`identified as ${document.info.format}/${document.info.kind}`);
    }
    if (document.info.units.length !== sample.pages) {
      throw new Error(`reported ${document.info.units.length} pages; expected ${sample.pages}`);
    }

    let objects = 0;
    for (let unitIndex = 0; unitIndex < sample.pages; unitIndex += 1) {
      const pageObjects = await document.listObjects(
        { unitIndex },
        { signal: AbortSignal.timeout(15_000) },
      );
      pageObjects.forEach(validateObject);
      const probe = sample.textProbes?.find(({ page }) => page === unitIndex + 1);
      if (probe !== undefined) {
        const text = pageObjects.map((object) => object.text ?? "").join("");
        for (const expected of probe.includes) {
          if (!text.includes(expected)) {
            throw new Error(`page ${probe.page} is missing expected text ${JSON.stringify(expected)}`);
          }
        }
        if (probe.excludesControlCharacters && /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(text)) {
          throw new Error(`page ${probe.page} contains decoded control characters`);
        }
      }
      objects += pageObjects.length;
    }
    return {
      path: sample.path,
      status: "opened",
      pages: sample.pages,
      objects,
      diagnostics: document.diagnostics().map(({ code, severity, fidelity, phase }) => ({
        code, severity, fidelity, phase,
      })),
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } catch (cause) {
    return {
      path: sample.path,
      status: cause instanceof OfficeEngineError ? "rejected" : "failed",
      code: cause?.code ?? "UNEXPECTED_ERROR",
      message: cause instanceof Error ? cause.message : String(cause),
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } finally {
    document?.close();
  }
}

function runVisualSuite() {
  if (platform() !== "darwin") throw new Error("The PDFKit visual corpus requires macOS");
  const cases = [
    ["001-trivial/minimal-document.pdf", "1", 99],
    ["007-imagemagick-images/imagemagick-ASCII85Decode.pdf", "1", 85],
    ["015-arabic/habibi-oneline-cmap.pdf", "1", 99],
    ["023-cmyk-image/cmyk-image.pdf", "1", 92, 0.65],
    ["027-cropped-rotated-scaled/cropped-rotated-scaled.pdf", "all", 85],
    ["028-image-references-deduplication/wrong-references.pdf", "all", 85],
    ["009-pdflatex-geotopo/GeoTopo.pdf", "1,59,117", 85],
  ];
  for (const [samplePath, pages, minimumScore, minimumTolerantSimilarity = 0] of cases) {
    const sample = manifest.samples.find(({ path }) => path === samplePath);
    if (sample === undefined) throw new Error(`Missing visual sample: ${samplePath}`);
    const { path } = verifiedPath(sample);
    const id = samplePath.slice(0, samplePath.indexOf("/"));
    const output = resolve(visualRoot, id);
    rmSync(output, { recursive: true, force: true });
    run(process.execPath, [
      "scripts/run-pdf-preview-suite.mjs",
      "--pages", pages,
      "--output", output,
      path,
    ], { allowFailure: true });
    const summary = JSON.parse(readFileSync(resolve(output, "reports/summary.json"), "utf8"));
    if (summary.environmentFingerprint?.matches !== true
        || summary.reports.length === 0
        || summary.reports.some(
          ({ overallScore, tolerantPixelSimilarity }) =>
            overallScore < minimumScore
              || tolerantPixelSimilarity < minimumTolerantSimilarity,
        )) {
      throw new Error(
        `${samplePath} fell below its PDFKit visual baseline `
          + `(score ${minimumScore}, tolerant pixels ${minimumTolerantSimilarity})`,
      );
    }
    console.log(
      `${summary.status === "passed" ? "PASS" : "BASE"}  ${samplePath} — `
        + `${summary.passed}/${summary.total} standard passes, score=${summary.averageScore}`,
    );
  }
}

function runFullVisualSuite() {
  if (platform() !== "darwin") throw new Error("The PDFKit visual corpus requires macOS");
  const renderable = manifest.samples.filter(
    ({ expectation }) => (expectation?.status ?? "opened") === "opened",
  );
  const unrenderable = manifest.samples.filter(({ expectation }) => expectation?.status === "rejected");
  rmSync(fullVisualRoot, { recursive: true, force: true });
  run(process.execPath, [
    "scripts/run-pdf-preview-suite.mjs",
    "--pages", "all",
    "--output", fullVisualRoot,
    ...renderable.map((sample) => verifiedPath(sample).path),
  ], { allowFailure: true, inherit: true });
  const summaryPath = resolve(fullVisualRoot, "reports/summary.json");
  if (!existsSync(summaryPath)) {
    throw new Error("The full PDF visual corpus stopped before producing a comparison report");
  }
  const summary = JSON.parse(readFileSync(summaryPath, "utf8"));
  const expectedRenderedPages = renderable.reduce((total, sample) => total + sample.pages, 0);
  const unrenderedPages = unrenderable.reduce((total, sample) => total + sample.pages, 0);
  const belowTarget = summary.reports
    .filter(({ tolerantPixelSimilarity }) => tolerantPixelSimilarity < 0.99)
    .sort((left, right) => left.tolerantPixelSimilarity - right.tolerantPixelSimilarity);
  const result = {
    schemaVersion: 1,
    targetSimilarity: 0.99,
    totalFiles: manifest.samples.length,
    totalPages: manifest.samples.reduce((total, sample) => total + sample.pages, 0),
    renderedPages: summary.reports.length,
    unrenderedPages,
    pagesAtOrAboveTarget: summary.reports.length - belowTarget.length,
    belowTarget,
    unrenderable: unrenderable.map(({ path, pages, expectation }) => ({
      path,
      pages,
      code: expectation.code,
    })),
  };
  writeFileSync(
    resolve(fullVisualRoot, "full-corpus-summary.json"),
    `${JSON.stringify(result, null, 2)}\n`,
  );
  if (summary.environmentFingerprint?.matches !== true
      || summary.reports.length !== expectedRenderedPages
      || belowTarget.length !== 0
      || unrenderedPages !== 0) {
    const worst = belowTarget[0];
    throw new Error(
      `Full PDF corpus missed the 99% target: ${result.pagesAtOrAboveTarget}/${result.totalPages} `
        + `pages passed, ${unrenderedPages} unrendered`
        + (worst === undefined
          ? ""
          : `, worst ${worst.id}=${worst.tolerantPixelSimilarity}`),
    );
  }
  console.log(`PASS  all ${result.totalPages} PDF pages are at least 99% similar to PDFKit`);
}

prepareCheckout();
if (visualFull) {
  runFullVisualSuite();
  process.exit(0);
}
if (visual) {
  runVisualSuite();
  console.log(`PASS  PDFKit visual corpus (${visualRoot})`);
  process.exit(0);
}

const wasm = readFileSync(resolve(root, "dist/office-viewer-core.wasm"));
const engine = await createOfficeEngine({
  execution: "inline",
  wasm,
  formatPack: async () => ({
    async load(candidate) {
      return readFileSync(await extendedFormatPack.load(candidate));
    },
  }),
});
const results = [];
try {
  for (const sample of selected) {
    const result = await openSample(engine, sample);
    const expectation = sample.expectation ?? { status: "opened" };
    const accepted = result.status === expectation.status
      && (expectation.code === undefined || result.code === expectation.code);
    results.push({ ...result, outcomeAccepted: accepted });
    if (!quiet || !accepted) {
      const detail = result.status === "opened"
        ? `${result.pages} pages, ${result.objects} objects, ${result.diagnostics.length} diagnostics`
        : `${result.code}: ${result.message}`;
      const label = !accepted ? "FAIL" : result.status === "rejected" ? "XFAIL" : "PASS";
      console.log(`${label.padEnd(5)} ${sample.path} — ${detail}`);
    }
  }
} finally {
  engine.close();
}

const failed = results.filter(({ outcomeAccepted }) => !outcomeAccepted);
const summary = {
  total: results.length,
  accepted: results.length - failed.length,
  unexpected: failed.length,
  opened: results.filter(({ status }) => status === "opened").length,
  rejected: results.filter(({ status }) => status === "rejected").length,
  manifestPages: selected.reduce((sum, sample) => sum + sample.pages, 0),
  materializedPages: results.reduce((sum, result) => sum + (result.pages ?? 0), 0),
};
mkdirSync(dirname(reportPath), { recursive: true });
writeFileSync(reportPath, `${JSON.stringify({
  generatedAt: new Date().toISOString(),
  mode: full ? "full" : "smoke",
  suite: manifest.suite,
  summary,
  cases: results,
}, null, 2)}\n`);

if (failed.length !== 0 || summary.total !== selected.length) {
  process.exitCode = 1;
} else {
  console.log(
    `PASS  ${full ? "full" : "smoke"} PDF corpus: ${summary.total} files, `
      + `${summary.manifestPages} manifest pages, ${summary.rejected} expected rejections`,
  );
}
