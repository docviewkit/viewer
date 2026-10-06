import { createHash } from "node:crypto";
import { spawn } from "node:child_process";
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { arch, platform, release, tmpdir } from "node:os";
import { basename, resolve } from "node:path";

import { chromium } from "playwright-core";

const PDF_PIXEL_RADIUS = 3;

function usage() {
  return [
    "Usage: node scripts/run-pdf-preview-suite.mjs [options] <file.pdf> [...]",
    "",
    "Options:",
    "  --pages <selection>          1-based pages: 1,3-5 or all (default: 1)",
    "  --output <directory>         Captures and reports (default: output/accuracy/pdf-preview)",
    "  --browser-executable <path> Chromium-compatible browser executable",
    "  --help                       Show this help",
  ].join("\n");
}

function parseArguments(arguments_) {
  if (arguments_.includes("--help")) return { help: true };
  let pages = "1";
  let output = "output/accuracy/pdf-preview";
  let browserExecutable;
  const pdfs = [];
  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index];
    if (["--pages", "--output", "--browser-executable"].includes(argument)) {
      const value = arguments_[++index];
      if (value === undefined) throw new Error(`${usage()}\nMissing ${argument} value`);
      if (argument === "--pages") pages = value;
      if (argument === "--output") output = value;
      if (argument === "--browser-executable") browserExecutable = value;
    } else if (argument.startsWith("-")) {
      throw new Error(`${usage()}\nUnknown option ${argument}`);
    } else {
      pdfs.push(resolve(argument));
    }
  }
  if (pdfs.length === 0) throw new Error(usage());
  if (!pdfs.every((path) => path.toLowerCase().endsWith(".pdf"))) {
    throw new Error("Every input must have a .pdf extension");
  }
  return {
    help: false,
    pages,
    output: resolve(output),
    browserExecutable: browserExecutable === undefined ? undefined : resolve(browserExecutable),
    pdfs,
  };
}

function selectedPageIndices(specification, count) {
  if (specification === "all") return Array.from({ length: count }, (_, index) => index);
  const selected = new Set();
  for (const part of specification.split(",")) {
    const match = /^(\d+)(?:-(\d+))?$/u.exec(part);
    if (match === null) throw new Error(`Invalid page selection: ${part}`);
    const first = Number(match[1]);
    const last = Number(match[2] ?? match[1]);
    if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last)
      || first < 1 || last < first || last > count) {
      throw new Error(`Page selection is outside the ${count}-page PDF: ${part}`);
    }
    for (let page = first; page <= last; page += 1) selected.add(page - 1);
  }
  if (selected.size === 0) throw new Error("Page selection cannot be empty");
  return [...selected].sort((left, right) => left - right);
}

function safeStem(path, index) {
  const stem = basename(path, ".pdf").normalize("NFKC")
    .replace(/[^A-Za-z0-9._-]+/gu, "-").replace(/^-+|-+$/gu, "");
  return `${String(index + 1).padStart(2, "0")}-${stem || "document"}`;
}

function startProcess(command, arguments_, options = {}) {
  return spawn(command, arguments_, { cwd: resolve("."), stdio: ["ignore", "pipe", "pipe"], ...options });
}

function run(command, arguments_, label) {
  return new Promise((resolvePromise, reject) => {
    const child = startProcess(command, arguments_);
    let output = "";
    child.stdout.on("data", (chunk) => {
      output += chunk;
      process.stdout.write(chunk);
    });
    child.stderr.on("data", (chunk) => {
      output += chunk;
      process.stderr.write(chunk);
    });
    child.once("error", reject);
    child.once("exit", (code, signal) => {
      if (code === 0) resolvePromise(output);
      else reject(new Error(`${label} exited with ${code ?? signal}: ${output.trim()}`));
    });
  });
}

function startServer(fixtureRoot) {
  const child = startProcess(process.execPath, [
    "scripts/serve.mjs", "--port", "0", "--fixture-root", fixtureRoot,
  ]);
  return new Promise((resolvePromise, reject) => {
    let output = "";
    const timeout = setTimeout(() => {
      child.kill("SIGTERM");
      reject(new Error(`Timed out starting the visual harness server: ${output.trim()}`));
    }, 20_000);
    const inspect = (chunk) => {
      output += chunk;
      const match = /(?:DocViewKit Viewer|Office Viewer Inspector): (http:\/\/[^\s]+\/)/u.exec(output);
      if (match !== null) {
        clearTimeout(timeout);
        resolvePromise({ child, baseUrl: match[1] });
      }
    };
    child.stdout.on("data", inspect);
    child.stderr.on("data", inspect);
    child.once("error", reject);
    child.once("exit", (code) => {
      clearTimeout(timeout);
      reject(new Error(`Visual harness server exited with ${code}: ${output.trim()}`));
    });
  });
}

async function stopProcess(child) {
  if (child.exitCode !== null || child.signalCode !== null) return;
  await new Promise((resolvePromise) => {
    const timeout = setTimeout(() => {
      child.kill("SIGKILL");
      resolvePromise();
    }, 5_000);
    child.once("exit", () => {
      clearTimeout(timeout);
      resolvePromise();
    });
    child.kill("SIGTERM");
  });
}

async function macOsVersion() {
  if (platform() !== "darwin") return release();
  return (await run("sw_vers", ["-productVersion"], "sw_vers")).trim();
}

async function fontDigest() {
  const bytes = await readFile(resolve("node_modules/@embedpdf/fonts-sc/fonts/NotoSansHans-Regular.otf"));
  return `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
}

async function writeJson(path, value) {
  await writeFile(path, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 });
}

const options = parseArguments(process.argv.slice(2));
if (options.help) {
  console.log(usage());
  process.exit(0);
}
if (platform() !== "darwin") throw new Error("The Preview oracle requires macOS PDFKit");

await mkdir(options.output, { recursive: true });
const fixtureRoot = await mkdtemp(resolve(tmpdir(), "officeviewer-pdf-preview-"));
const inputs = [];
for (const [index, input] of options.pdfs.entries()) {
  const stem = safeStem(input, index);
  const fixture = `${stem}.pdf`;
  await copyFile(input, resolve(fixtureRoot, fixture));
  inputs.push({ input, fixture, stem });
}

const server = await startServer(fixtureRoot);
let browser;
let comparisonFailed = false;
try {
  browser = await chromium.launch(options.browserExecutable === undefined
    ? { channel: "chrome", headless: true }
    : { executablePath: options.browserExecutable, headless: true });
  const context = await browser.newContext({ locale: "en-US", timezoneId: "Asia/Shanghai", deviceScaleFactor: 1 });
  const cases = [];
  for (const input of inputs) {
    const page = await context.newPage();
    const browserErrors = [];
    page.on("console", (message) => {
      if (message.type() === "error") browserErrors.push(`console: ${message.text()}`);
    });
    page.on("pageerror", (cause) => browserErrors.push(`pageerror: ${cause.message}`));
    const url = new URL("examples/visual-harness.html", server.baseUrl);
    url.searchParams.set("fixture", input.fixture);
    url.searchParams.set("unit", "0");
    // The PDF model uses CSS points (96 dpi); Preview/PDFKit references use PDF points (72 dpi).
    url.searchParams.set("scale", "0.75");
    await page.goto(url.href, { waitUntil: "domcontentloaded", timeout: 90_000 });
    await page.waitForFunction(() => ["pass", "error"].includes(document.body.dataset.state), undefined, { timeout: 90_000 });
    const harness = await page.evaluate(() => globalThis.visualHarness);
    if (harness?.state !== "pass") throw new Error(`${input.input}: ${harness?.message ?? "visual harness failed"}`);
    if (harness.format !== "pdf") throw new Error(`${input.input}: expected PDF but opened ${harness.format}`);
    const pageIndices = selectedPageIndices(options.pages, harness.unitCount);
    const referenceDirectory = resolve(options.output, "reference", input.stem);
    const actualDirectory = resolve(options.output, "actual", input.stem);
    const referenceObservationDirectory = resolve(options.output, "observations", "reference", input.stem);
    const actualObservationDirectory = resolve(options.output, "observations", "actual", input.stem);
    await mkdir(referenceDirectory, { recursive: true });
    await mkdir(actualDirectory, { recursive: true });
    await mkdir(referenceObservationDirectory, { recursive: true });
    await mkdir(actualObservationDirectory, { recursive: true });
    const dimensions = {};
    for (const pageIndex of pageIndices) {
      const capture = await page.evaluate(async (index) => globalThis.visualHarness.renderUnit(index), pageIndex);
      if (typeof capture?.dataUrl !== "string" || !capture.dataUrl.startsWith("data:image/png;base64,")) {
        throw new Error(`${input.input} page ${pageIndex + 1} did not export a PNG`);
      }
      if (!Array.isArray(capture.objects) || !Array.isArray(capture.diagnostics)) {
        throw new Error(`${input.input} page ${pageIndex + 1} did not export font-aware visual evidence`);
      }
      const attention = capture.attention;
      if (attention === null
        || typeof attention !== "object"
        || !["complete", "bounded"].includes(attention.status)
        || !Number.isSafeInteger(attention.padding)
        || attention.padding < PDF_PIXEL_RADIUS
        || ![
          attention.sourceTextObjectCount,
          attention.visibleTextObjectCount,
          attention.candidateCount,
          attention.mergedRegionCount,
          attention.regionCount,
          attention.pixelCount,
          attention.omittedCandidateCount,
          attention.omittedRegionCount,
        ].every((value) => Number.isSafeInteger(value) && value >= 0)
        || attention.regionCount !== capture.objects.length) {
        throw new Error(`${input.input} page ${pageIndex + 1} exported invalid text-attention evidence`);
      }
      if (attention.status === "bounded") {
        throw new Error(
          `${input.input} page ${pageIndex + 1} text-attention evidence exceeded its bounded limits `
          + `(omitted candidates=${attention.omittedCandidateCount}, regions=${attention.omittedRegionCount})`,
        );
      }
      if (attention.visibleTextObjectCount !== 0 && capture.objects.length === 0) {
        throw new Error(`${input.input} page ${pageIndex + 1} exported no visible text-attention regions`);
      }
      const fontFailures = capture.diagnostics.filter(({ code }) => (
        code === "FONT_LOAD_FAILED"
        || code === "FONT_LOAD_TIMEOUT"
        || code === "FONT_ENVIRONMENT_UNAVAILABLE"
      ));
      if (fontFailures.length !== 0) {
        throw new Error(
          `${input.input} page ${pageIndex + 1} rejected embedded fonts: ${fontFailures.map(({ message }) => message).join("; ")}`,
        );
      }
      const filename = `page-${String(pageIndex + 1).padStart(4, "0")}.png`;
      const observationFilename = `page-${String(pageIndex + 1).padStart(4, "0")}.json`;
      await writeFile(
        resolve(actualDirectory, filename),
        Buffer.from(capture.dataUrl.slice("data:image/png;base64,".length), "base64"),
        { mode: 0o600 },
      );
      // Candidate text bounds only select independently rendered PDFKit pixels
      // for closer inspection; they are not treated as reference geometry.
      const attentionObservation = { objects: capture.objects, attention };
      await writeJson(resolve(referenceObservationDirectory, observationFilename), attentionObservation);
      await writeJson(resolve(actualObservationDirectory, observationFilename), attentionObservation);
      dimensions[String(pageIndex + 1)] = { width: capture.width, height: capture.height };
      cases.push({
        id: `${input.stem}-page-${String(pageIndex + 1).padStart(4, "0")}`,
        corpusClass: harness.unitCount > 100 ? "large" : "enterprise",
        unitType: "page",
        unitIndex: pageIndex,
        fixture: input.input,
        goldenPng: `reference/${input.stem}/${filename}`,
        actualPng: `actual/${input.stem}/${filename}`,
        expectedObservationJson: `observations/reference/${input.stem}/${observationFilename}`,
        actualObservationJson: `observations/actual/${input.stem}/${observationFilename}`,
        declaredCoverage: [
          "structure-units",
          "visual-exact",
          "visual-tolerant",
          "visual-ssim",
          ...(capture.objects.length === 0 ? [] : ["visual-object-region"]),
        ],
        policy: {
          pixelChannel: 8,
          pixelRadius: PDF_PIXEL_RADIUS,
          minExactPixelRatio: 0,
          minTolerantPixelRatio: 0.98,
          minSsim: 0.90,
          // PDFKit and Canvas rasterize the same valid face differently. This
          // threshold remains above the measured missing-glyph regression while
          // allowing the independently rendered, browser-valid glyph outlines.
          minObjectRegionSimilarity: 0.65,
          minOverallScore: 95,
        },
      });
      console.log(`Captured ${basename(input.input)} page ${pageIndex + 1}/${harness.unitCount}`);
    }
    const dimensionsPath = resolve(options.output, `${input.stem}-dimensions.json`);
    await writeJson(dimensionsPath, dimensions);
    await run("swift", [
      "scripts/render-pdf-preview-reference.swift", input.input, referenceDirectory, "1", options.pages, dimensionsPath,
    ], `PDFKit reference rendering for ${input.input}`);
    if (browserErrors.length !== 0) throw new Error(`${input.input}: ${browserErrors.join("; ")}`);
    await page.close();
  }

  const fingerprint = {
    os: "macOS",
    osVersion: await macOsVersion(),
    architecture: arch(),
    locale: "en-US",
    timezone: "Asia/Shanghai",
    colorSpace: "srgb",
    devicePixelRatio: 1,
    scale: 0.75,
    background: "#ffffff",
    fontSetDigest: await fontDigest(),
    referenceRenderer: "macOS PDFKit (Preview rendering stack)",
    referenceRendererVersion: await macOsVersion(),
    candidateRenderer: "OfficeViewer PDF format pack",
    candidateRendererVersion: JSON.parse(await readFile(resolve("package.json"), "utf8")).version,
  };
  const suitePath = resolve(options.output, "suite.json");
  const fingerprintPath = resolve(options.output, "environment.json");
  await writeJson(suitePath, { schemaVersion: 1, oracleMode: "read-only", environmentFingerprint: fingerprint, cases });
  await writeJson(fingerprintPath, fingerprint);
  try {
    await run(process.execPath, [
      "scripts/run-external-visual-suite.mjs", suitePath,
      "--fingerprint", fingerprintPath,
      "--actual-root", options.output,
      "--output", resolve(options.output, "reports"),
    ], "PDF Preview comparison");
  } catch (cause) {
    comparisonFailed = true;
    console.error(cause instanceof Error ? cause.message : cause);
  }
} finally {
  await browser?.close();
  await stopProcess(server.child);
  await rm(fixtureRoot, { recursive: true, force: true });
}

if (comparisonFailed) process.exitCode = 1;
