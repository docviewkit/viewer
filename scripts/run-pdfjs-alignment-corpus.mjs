import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { arch, platform, release } from "node:os";
import { basename, dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { chromium } from "playwright-core";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const corpusRoot = resolve(root, "tests/external/pdfjs/test/pdfs");
const outputRoot = resolve(root, "output/accuracy/pdfjs-reference-corpus");
const referenceLedgerPath = resolve(outputRoot, "ledger.json");
const pdfjsManifestPath = resolve(root, "tests/external/pdfjs/test/test_manifest.json");
const acrobatOraclePath = resolve(root, "tests/pdfjs-acrobat-oracle.json");
const candidateLedgerPath = resolve(outputRoot, "officeviewer-ledger.json");
const comparisonSuitePath = resolve(outputRoot, "alignment-suite.json");
const fingerprintPath = resolve(outputRoot, "alignment-environment.json");
const reportRoot = resolve(outputRoot, "alignment-reports");
const channelTolerance = 8;
const pixelRadius = 3;
const maxComparisonPixels = 25_000_000;
const require = createRequire(resolve(root, "tests/external/pdfjs/package.json"));

function usage() {
  return [
    "Usage: node scripts/run-pdfjs-alignment-corpus.mjs [options]",
    "",
    "Options:",
    "  --concurrency <count>        Browser workers (default: 2)",
    "  --timeout-ms <milliseconds>  Per-PDF timeout (default: 90000)",
    "  --limit <count>              Deterministic prefix for smoke tests",
    "  --match <pattern>            Only basenames containing this text",
    "  --include-acrobat-required   Capture PDFs that need an Acrobat oracle",
    "  --resume                     Reuse completed captures with the same SHA-256",
    "  --capture-only               Skip the final visual comparison",
    "  --help                       Show this help",
  ].join("\n");
}

function positiveInteger(value, option) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 1) throw new Error(`${option} must be a positive integer`);
  return parsed;
}

function parseArguments(arguments_) {
  const options = {
    concurrency: 2,
    timeoutMs: 90_000,
    limit: undefined,
    match: undefined,
    resume: false,
    captureOnly: false,
    includeAcrobatRequired: false,
  };
  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index];
    if (argument === "--help") return { ...options, help: true };
    if (argument === "--resume") { options.resume = true; continue; }
    if (argument === "--capture-only") { options.captureOnly = true; continue; }
    if (argument === "--include-acrobat-required") { options.includeAcrobatRequired = true; continue; }
    const value = arguments_[++index];
    if (value === undefined) throw new Error(`${usage()}\nMissing ${argument} value`);
    if (argument === "--concurrency") options.concurrency = positiveInteger(value, argument);
    else if (argument === "--timeout-ms") options.timeoutMs = positiveInteger(value, argument);
    else if (argument === "--limit") options.limit = positiveInteger(value, argument);
    else if (argument === "--match") options.match = value;
    else throw new Error(`${usage()}\nUnknown option ${argument}`);
  }
  return { ...options, help: false };
}

function writeJsonAtomic(path, value) {
  mkdirSync(dirname(path), { recursive: true });
  const temporary = `${path}.tmp-${process.pid}`;
  writeFileSync(temporary, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 });
  renameSync(temporary, path);
}

function candidateDigest() {
  const digest = createHash("sha256");
  for (const path of [
    "dist/office-viewer-core.wasm",
    "dist/office-viewer-pdf.wasm",
    "dist/worker-engine.js",
    "dist/render.js",
    "dist/image-codec-worker.js",
    "examples/visual-harness.js",
  ]) {
    digest.update(path).update("\0").update(readFileSync(resolve(root, path)));
  }
  return digest.digest("hex");
}

function startProcess(command, arguments_) {
  return spawn(command, arguments_, { cwd: root, stdio: ["ignore", "pipe", "pipe"] });
}

function startServer() {
  const child = startProcess(process.execPath, [
    "scripts/serve.mjs", "--port", "0", "--fixture-root", corpusRoot,
  ]);
  return new Promise((resolvePromise, reject) => {
    let output = "";
    const timeout = setTimeout(() => {
      child.kill("SIGTERM");
      reject(new Error(`Timed out starting Viewer server: ${output.trim()}`));
    }, 20_000);
    const inspect = (chunk) => {
      output += chunk;
      const match = /DocViewKit Viewer: (http:\/\/[^\s]+\/)/u.exec(output);
      if (match === null) return;
      clearTimeout(timeout);
      resolvePromise({ child, baseUrl: match[1] });
    };
    child.stdout.on("data", inspect);
    child.stderr.on("data", inspect);
    child.once("error", reject);
    child.once("exit", (code) => {
      clearTimeout(timeout);
      reject(new Error(`Viewer server exited with ${code}: ${output.trim()}`));
    });
  });
}

async function stopProcess(child) {
  if (child.exitCode !== null || child.signalCode !== null) return;
  await new Promise((resolvePromise) => {
    const timeout = setTimeout(() => { child.kill("SIGKILL"); resolvePromise(); }, 5_000);
    child.once("exit", () => { clearTimeout(timeout); resolvePromise(); });
    child.kill("SIGTERM");
  });
}

async function closeBrowser(browser) {
  let timeout;
  await Promise.race([
    browser.close().catch(() => undefined),
    new Promise((resolvePromise) => {
      timeout = setTimeout(resolvePromise, 5_000);
    }),
  ]);
  clearTimeout(timeout);
}

function stemFor(entry) {
  if (!Array.isArray(entry.pages) || entry.pages.length === 0) {
    return `${basename(entry.path, ".pdf")}-${entry.sha256.slice(0, 12)}`;
  }
  const prefix = "reference/";
  const directory = dirname(entry.pages[0].png).split("\\").join("/");
  if (!directory.startsWith(prefix)) throw new Error(`${entry.path} has an invalid reference path`);
  return directory.slice(prefix.length);
}

function fixturePasswords() {
  const manifest = JSON.parse(readFileSync(pdfjsManifestPath, "utf8"));
  const passwords = new Map();
  for (const test of manifest) {
    if (typeof test?.file !== "string" || !test.file.startsWith("pdfs/")) continue;
    const path = test.file.slice("pdfs/".length);
    if (typeof test.password === "string") passwords.set(path, test.password);
  }
  for (const path of ["auth-event-ef-open.pdf", "empty_protected.pdf", "issue7665.pdf", "secHandler.pdf"]) {
    passwords.set(path, "");
  }
  passwords.set("print_protection.pdf", "password");
  return passwords;
}

async function capture(browser, baseUrl, entry, password, timeoutMs) {
  const started = performance.now();
  const page = await browser.newPage({ locale: "en-US", timezoneId: "Asia/Shanghai", deviceScaleFactor: 1 });
  const browserErrors = [];
  page.on("pageerror", (cause) => browserErrors.push(cause.message));
  page.on("console", (message) => { if (message.type() === "error") browserErrors.push(message.text()); });
  try {
    const url = new URL("examples/visual-harness.html", baseUrl);
    url.searchParams.set("fixture", entry.path);
    url.searchParams.set("unit", "0");
    url.searchParams.set("scale", "0.75");
    if (password !== undefined) url.searchParams.set("password", password);
    await page.goto(url.href, { waitUntil: "domcontentloaded", timeout: timeoutMs });
    await page.waitForFunction(
      () => ["pass", "error"].includes(document.body.dataset.state),
      undefined,
      { timeout: timeoutMs },
    );
    const harness = await page.evaluate(() => globalThis.visualHarness);
    if (harness?.state !== "pass") throw new Error(harness?.message ?? "OfficeViewer harness failed");
    if (harness.format !== "pdf") throw new Error(`Opened as ${harness.format}, expected pdf`);
    const stem = stemFor(entry);
    const actualDirectory = resolve(outputRoot, "actual", stem);
    mkdirSync(actualDirectory, { recursive: true });
    const pages = [];
    const hasReferencePages = Array.isArray(entry.pages) && Number.isSafeInteger(entry.pageCount);
    const commonPages = hasReferencePages ? Math.min(harness.unitCount, entry.pageCount) : harness.unitCount;
    for (let index = 0; index < commonPages; index += 1) {
      const pageStarted = performance.now();
      const expected = hasReferencePages ? entry.pages[index] : undefined;
      const capture_ = await page.evaluate(async ({ unitIndex, expectedWidth, expectedHeight }) => {
        const rendered = await globalThis.visualHarness.renderUnit(unitIndex);
        const hasExpectedSize = Number.isSafeInteger(expectedWidth) && Number.isSafeInteger(expectedHeight);
        const widthDelta = hasExpectedSize ? Math.abs(rendered.width - expectedWidth) : 0;
        const heightDelta = hasExpectedSize ? Math.abs(rendered.height - expectedHeight) : 0;
        if (hasExpectedSize && (widthDelta !== 0 || heightDelta !== 0) && widthDelta <= 1 && heightDelta <= 1) {
          const image = new Image();
          image.src = rendered.dataUrl;
          await image.decode();
          const canvas = document.createElement("canvas");
          canvas.width = expectedWidth;
          canvas.height = expectedHeight;
          const context = canvas.getContext("2d");
          context.fillStyle = "#ffffff";
          context.fillRect(0, 0, expectedWidth, expectedHeight);
          context.drawImage(image, 0, 0);
          return {
            ...rendered,
            rawWidth: rendered.width,
            rawHeight: rendered.height,
            width: expectedWidth,
            height: expectedHeight,
            dataUrl: canvas.toDataURL("image/png"),
            comparisonNormalized: true,
          };
        }
        return { ...rendered, rawWidth: rendered.width, rawHeight: rendered.height };
      }, { unitIndex: index, expectedWidth: expected?.width, expectedHeight: expected?.height });
      if (typeof capture_?.dataUrl !== "string" || !capture_.dataUrl.startsWith("data:image/png;base64,")) {
        throw new Error(`Page ${index + 1} did not export a PNG`);
      }
      const png = Buffer.from(capture_.dataUrl.slice("data:image/png;base64,".length), "base64");
      const filename = `page-${String(index + 1).padStart(4, "0")}.png`;
      writeFileSync(resolve(actualDirectory, filename), png, { mode: 0o600 });
      pages.push({
        page: index + 1,
        width: capture_.width,
        height: capture_.height,
        rawWidth: capture_.rawWidth,
        rawHeight: capture_.rawHeight,
        comparisonNormalized: capture_.comparisonNormalized === true,
        png: `actual/${stem}/${filename}`,
        diagnostics: capture_.diagnostics.map(({ code, severity, fidelity, phase, message }) => ({
          code, severity, fidelity, phase, message,
        })),
        durationMs: Math.round((performance.now() - pageStarted) * 10) / 10,
      });
    }
    return {
      path: entry.path,
      sha256: entry.sha256,
      status: !hasReferencePages || harness.unitCount === entry.pageCount ? "captured" : "unit-count-mismatch",
      oracle: hasReferencePages ? "pdf.js" : "acrobat-required",
      referencePages: hasReferencePages ? entry.pageCount : undefined,
      candidatePages: harness.unitCount,
      pages,
      browserErrors,
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } catch (cause) {
    const oracleAligned = entry.acrobatOracle?.outcome === "damaged"
      || entry.acrobatOracle?.outcome === "password-required";
    return {
      path: entry.path,
      sha256: entry.sha256,
      status: oracleAligned ? "oracle-aligned-failure" : "officeviewer-failed",
      oracle: oracleAligned ? "Adobe Acrobat Reader" : entry.oracle,
      oracleOutcome: entry.acrobatOracle?.outcome,
      code: cause?.name ?? cause?.code ?? "OFFICEVIEWER_ERROR",
      message: cause instanceof Error ? cause.message : String(cause),
      browserErrors,
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } finally {
    await page.close().catch(() => undefined);
  }
}

async function normalizedComparisonPng(relativePath, stem, filename, width, height) {
  if (width * height <= maxComparisonPixels) return relativePath;
  const scale = Math.sqrt(maxComparisonPixels / (width * height));
  const targetWidth = Math.max(1, Math.floor(width * scale));
  const targetHeight = Math.max(1, Math.floor(height * scale));
  const source = resolve(outputRoot, relativePath);
  const targetRelative = `comparison/${stem}/${filename}`;
  const target = resolve(outputRoot, targetRelative);
  const { createCanvas, loadImage } = require("@napi-rs/canvas");
  const image = await loadImage(source);
  const canvas = createCanvas(targetWidth, targetHeight);
  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, targetWidth, targetHeight);
  context.drawImage(image, 0, 0, targetWidth, targetHeight);
  mkdirSync(dirname(target), { recursive: true });
  writeFileSync(target, canvas.toBuffer("image/png"), { mode: 0o600 });
  return targetRelative;
}

async function comparisonCases(referenceEntries, candidates) {
  const referenceByPath = new Map(referenceEntries.map((entry) => [entry.path, entry]));
  const cases = [];
  for (const candidate of candidates) {
    const reference = referenceByPath.get(candidate.path);
    if (reference === undefined || !Array.isArray(reference.pages) || !Array.isArray(candidate.pages)) continue;
    for (const actual of candidate.pages) {
      const expected = reference.pages[actual.page - 1];
      if (expected === undefined) continue;
      const filename = `page-${String(actual.page).padStart(4, "0")}.png`;
      const goldenPng = await normalizedComparisonPng(
        expected.png,
        `reference/${stemFor(reference)}`,
        filename,
        expected.width,
        expected.height,
      );
      const actualPng = await normalizedComparisonPng(
        actual.png,
        `actual/${stemFor(reference)}`,
        filename,
        actual.width,
        actual.height,
      );
      cases.push({
        id: `${stemFor(reference)}-page-${String(actual.page).padStart(4, "0")}`,
        corpusClass: reference.pageCount > 100 ? "large" : "enterprise",
        unitType: "page",
        unitIndex: actual.page - 1,
        fixture: resolve(corpusRoot, reference.path),
        goldenPng,
        actualPng,
        declaredCoverage: ["structure-units", "visual-exact", "visual-tolerant", "visual-ssim"],
        policy: {
          pixelChannel: channelTolerance,
          pixelRadius,
          minExactPixelRatio: 0,
          minTolerantPixelRatio: 0.98,
          minSsim: 0.90,
          minOverallScore: 95,
        },
      });
    }
  }
  return cases;
}

async function runComparison(fingerprint, cases) {
  writeJsonAtomic(comparisonSuitePath, {
    schemaVersion: 1,
    oracleMode: "read-only",
    environmentFingerprint: fingerprint,
    cases,
  });
  writeJsonAtomic(fingerprintPath, fingerprint);
  await new Promise((resolvePromise, reject) => {
    const child = startProcess(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      comparisonSuitePath,
      "--fingerprint", fingerprintPath,
      "--actual-root", outputRoot,
      "--output", reportRoot,
    ]);
    let output = "";
    child.stdout.on("data", (chunk) => { output += chunk; process.stdout.write(chunk); });
    child.stderr.on("data", (chunk) => { output += chunk; process.stderr.write(chunk); });
    child.once("error", reject);
    child.once("exit", (code) => code === 0
      ? resolvePromise()
      : reject(new Error(`Visual comparison exited with ${code}: ${output.trim().slice(-4_000)}`)));
  });
}

async function main(options) {
  const referenceLedger = JSON.parse(readFileSync(referenceLedgerPath, "utf8"));
  const acrobatOracle = JSON.parse(readFileSync(acrobatOraclePath, "utf8")).files;
  let references = referenceLedger.files.filter(({ status }) => status === "rendered"
    || (options.includeAcrobatRequired && status === "pdfjs-failed"))
    .map((entry) => entry.status === "pdfjs-failed"
      ? { ...entry, acrobatOracle: acrobatOracle[entry.path] }
      : entry);
  const missingAcrobatOracle = references.find((entry) => entry.status === "pdfjs-failed"
    && entry.acrobatOracle === undefined);
  if (missingAcrobatOracle !== undefined) {
    throw new Error(`Missing Acrobat oracle result for ${missingAcrobatOracle.path}`);
  }
  if (options.match !== undefined) {
    references = references.filter(({ path }) => basename(path).toLowerCase().includes(options.match.toLowerCase()));
  }
  if (options.limit !== undefined) references = references.slice(0, options.limit);
  if (references.length === 0) throw new Error("No rendered PDF.js references matched this run");
  const previous = options.resume && existsSync(candidateLedgerPath)
    ? JSON.parse(readFileSync(candidateLedgerPath, "utf8"))
    : undefined;
  const buildDigest = candidateDigest();
  const completed = new Map((previous?.candidateDigest === buildDigest ? previous.files : [])
    .filter(({ status }) => ["captured", "unit-count-mismatch", "oracle-aligned-failure"].includes(status))
    .map((entry) => [`${entry.path}\0${entry.sha256}`, entry]));
  const results = [];
  const pending = [];
  const passwords = fixturePasswords();
  for (const entry of references) {
    const prior = completed.get(`${entry.path}\0${entry.sha256}`);
    if (prior === undefined) pending.push(entry); else results.push(prior);
  }
  const server = await startServer();
  const browsers = [];
  let next = 0;
  const startedAt = new Date().toISOString();
  const writeLedger = () => writeJsonAtomic(candidateLedgerPath, {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),
    startedAt,
    referenceLedger: referenceLedgerPath,
    candidateDigest: buildDigest,
    summary: {
      selectedPdfs: references.length,
      completedPdfs: results.length,
      pendingPdfs: references.length - results.length,
      capturedPdfs: results.filter(({ status }) => status === "captured").length,
      oracleAlignedFailures: results.filter(({ status }) => status === "oracle-aligned-failure").length,
      unitCountMismatches: results.filter(({ status }) => status === "unit-count-mismatch").length,
      failedPdfs: results.filter(({ status }) => status === "officeviewer-failed").length,
      capturedPages: results.reduce((sum, entry) => sum + (entry.pages?.length ?? 0), 0),
    },
    files: [...results].sort((left, right) => left.path.localeCompare(right.path, "en")),
  });
  writeLedger();
  try {
    for (let index = 0; index < Math.min(options.concurrency, pending.length); index += 1) {
      const browser = await chromium.launch({ channel: "chrome", headless: true });
      browsers.push(browser);
      void index;
    }
    await Promise.all(browsers.map(async (browser) => {
      while (next < pending.length) {
        const entry = pending[next++];
        let result = await capture(
          browser,
          server.baseUrl,
          entry,
          passwords.get(entry.path),
          options.timeoutMs,
        );
        if (result.status === "officeviewer-failed") {
          const retryBrowser = await chromium.launch({ channel: "chrome", headless: true });
          try {
            result = await capture(
              retryBrowser,
              server.baseUrl,
              entry,
              passwords.get(entry.path),
              options.timeoutMs,
            );
          } finally {
            await closeBrowser(retryBrowser);
          }
        }
        results.push(result);
        writeLedger();
        console.log(
          `${String(results.length).padStart(String(references.length).length)}/${references.length} `
          + `${result.status.padEnd(20)} ${entry.path}, ${result.pages?.length ?? 0} pages`,
        );
      }
    }));
  } finally {
    await Promise.all(browsers.map(closeBrowser));
    await stopProcess(server.child);
  }
  writeLedger();
  const browserVersion = browsers[0]?.version() ?? "unknown";
  const fingerprint = {
    os: platform(),
    osVersion: release(),
    architecture: arch(),
    locale: "en-US",
    timezone: "Asia/Shanghai",
    colorSpace: "srgb",
    devicePixelRatio: 1,
    scale: 0.75,
    background: "#ffffff",
    fontSetDigest: `sha256:${createHash("sha256").update(readFileSync(
      resolve(root, "node_modules/@embedpdf/fonts-sc/fonts/NotoSansHans-Regular.otf"),
    )).digest("hex")}`,
    referenceRenderer: `pdf.js ${referenceLedger.environment.pdfjsCommit}`,
    referenceRendererVersion: referenceLedger.environment.pdfjsCommit,
    candidateRenderer: "OfficeViewer PDF format pack",
    candidateRendererVersion: JSON.parse(readFileSync(resolve(root, "package.json"), "utf8")).version,
    browser: `Chrome ${browserVersion}`,
  };
  const cases = await comparisonCases(references, results);
  if (!options.captureOnly) await runComparison(fingerprint, cases);
  console.log(`OfficeViewer corpus capture complete: ${results.length}/${references.length}, ${cases.length} comparable pages`);
}

const options = parseArguments(process.argv.slice(2));
if (options.help) console.log(usage());
else await main(options);
