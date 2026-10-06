import { createHash } from "node:crypto";
import { spawn, spawnSync } from "node:child_process";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { arch, platform, release } from "node:os";
import { basename, dirname, resolve } from "node:path";

import { chromium } from "playwright-core";

import {
  flattenNativeVisualCases,
  resolveSafeRelativePath,
  sha256,
  validateNativeReference,
  validateNativeVisualSuite,
} from "./native-visual-policy.mjs";

function usage() {
  return [
    "Usage: node scripts/capture-native-visual-suite.mjs <suite.json> [options]",
    "",
    "Options:",
    "  --output <directory>            Candidate images, provenance, and reports",
    "  --font-root <directory>         Font files and optional manifest",
    "  --font-manifest <file.json>     Deterministic font manifest served to the harness",
    "  --browser-executable <path>     Chromium-compatible browser executable",
    "  --validate-only                 Validate manifest, fixture/golden hashes, and font digest",
    "  --help                          Show this help",
  ].join("\n");
}

function parseArguments(arguments_) {
  if (arguments_.includes("--help")) return { help: true };
  let suitePath;
  let output;
  let fontRoot;
  let fontManifest;
  let browserExecutable;
  let validateOnly = false;
  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index];
    if (argument === "--output" || argument === "--font-root" || argument === "--font-manifest" || argument === "--browser-executable") {
      const value = arguments_[++index];
      if (value === undefined) throw new Error(`${usage()}\nMissing ${argument} value`);
      if (argument === "--output") output = value;
      if (argument === "--font-root") fontRoot = value;
      if (argument === "--font-manifest") fontManifest = value;
      if (argument === "--browser-executable") browserExecutable = value;
    } else if (argument === "--validate-only") {
      validateOnly = true;
    } else if (argument.startsWith("-")) {
      throw new Error(`${usage()}\nUnknown option ${argument}`);
    } else if (suitePath === undefined) {
      suitePath = argument;
    } else {
      throw new Error(`${usage()}\nUnexpected argument ${argument}`);
    }
  }
  if (suitePath === undefined) throw new Error(usage());
  const absoluteSuitePath = resolve(suitePath);
  return {
    help: false,
    suitePath: absoluteSuitePath,
    output: resolve(output ?? resolve("output/accuracy/native", basename(suitePath, ".json"))),
    fontRoot: resolve(fontRoot ?? "node_modules/@embedpdf/fonts-sc/fonts"),
    fontManifest,
    browserExecutable: browserExecutable === undefined ? undefined : resolve(browserExecutable),
    validateOnly,
  };
}

function canonicalFontEntry(entry) {
  return JSON.stringify({
    family: entry.family,
    file: entry.file,
    style: entry.style ?? "normal",
    weight: entry.weight ?? 400,
    stretch: entry.stretch ?? "normal",
  });
}

async function fontSetDigest(fontRoot, fontManifest) {
  const entries = fontManifest === undefined
    ? [{ family: "Noto Sans S Chinese", file: "NotoSansHans-Regular.otf" }]
    : JSON.parse(await readFile(resolveSafeRelativePath(fontRoot, fontManifest, "fontManifest"), "utf8")).fonts;
  if (!Array.isArray(entries) || entries.length === 0 || entries.length > 256) {
    throw new Error("Font manifest must contain from 1 through 256 fonts");
  }
  const hash = createHash("sha256");
  hash.update("officeviewer-native-font-set-v1\0");
  for (const entry of [...entries].sort((left, right) => canonicalFontEntry(left).localeCompare(canonicalFontEntry(right)))) {
    if (entry === null || typeof entry !== "object" || typeof entry.file !== "string" || typeof entry.family !== "string") {
      throw new Error("Every font manifest entry must contain family and file strings");
    }
    const bytes = await readFile(resolveSafeRelativePath(fontRoot, entry.file, `font ${entry.file}`));
    hash.update(canonicalFontEntry(entry));
    hash.update("\0");
    hash.update(sha256(bytes));
    hash.update("\0");
  }
  return `sha256:${hash.digest("hex")}`;
}

async function validateAssets(suitePath, suite, cases) {
  const suiteDirectory = dirname(suitePath);
  const fixtureDigests = new Map();
  const goldenAssets = new Map();
  const referenceCache = new Map();
  for (const testCase of cases) {
    if (!fixtureDigests.has(testCase.fixture)) {
      const bytes = await readFile(resolveSafeRelativePath(suiteDirectory, testCase.fixture, `${testCase.id}.fixture`));
      fixtureDigests.set(testCase.fixture, sha256(bytes));
    }
    if (fixtureDigests.get(testCase.fixture) !== testCase.fixtureSha256) {
      throw new Error(`${testCase.id} fixture SHA-256 does not match the manifest`);
    }
    if (!goldenAssets.has(testCase.goldenPng)) {
      const bytes = await readFile(resolveSafeRelativePath(suiteDirectory, testCase.goldenPng, `${testCase.id}.goldenPng`));
      goldenAssets.set(testCase.goldenPng, { bytes, digest: sha256(bytes) });
    }
    const golden = goldenAssets.get(testCase.goldenPng);
    if (golden.digest !== testCase.goldenPngSha256) {
      throw new Error(`${testCase.id} golden PNG SHA-256 does not match the manifest`);
    }
    await validateNativeReference(suiteDirectory, suite, testCase, golden.bytes, referenceCache);
  }
}

function startServer(suiteDirectory, fontRoot) {
  const child = spawn(process.execPath, [
    "scripts/serve.mjs",
    "--port", "0",
    "--fixture-root", suiteDirectory,
    "--font-root", fontRoot,
  ], { cwd: resolve("."), stdio: ["ignore", "pipe", "pipe"] });
  return new Promise((resolvePromise, reject) => {
    let settled = false;
    let output = "";
    const timeout = setTimeout(() => {
      if (settled) return;
      settled = true;
      child.kill("SIGTERM");
      reject(new Error(`Timed out starting the visual harness server: ${output.trim()}`));
    }, 20_000);
    const inspect = (chunk) => {
      output += chunk.toString();
      const match = /(?:DocViewKit Viewer|Office Viewer Inspector): (http:\/\/[^\s]+\/)/u.exec(output);
      if (match !== null && !settled) {
        settled = true;
        clearTimeout(timeout);
        resolvePromise({ child, baseUrl: match[1] });
      }
    };
    child.stdout.on("data", inspect);
    child.stderr.on("data", inspect);
    child.once("exit", (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      reject(new Error(`Visual harness server exited with ${code}: ${output.trim()}`));
    });
    child.once("error", (cause) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      reject(cause);
    });
  });
}

function macOsVersion() {
  if (platform() !== "darwin") return release();
  const result = spawnSync("sw_vers", ["-productVersion"], { encoding: "utf8" });
  return result.status === 0 && result.stdout.trim() !== "" ? result.stdout.trim() : release();
}

function osName() {
  if (platform() === "darwin") return "macOS";
  if (platform() === "win32") return "Windows";
  return "Linux";
}

function buildRevision() {
  const revision = spawnSync("git", ["rev-parse", "HEAD"], { cwd: resolve("."), encoding: "utf8" });
  if (revision.status !== 0 || revision.stdout.trim() === "") return "unrecorded";
  const status = spawnSync("git", ["status", "--porcelain"], { cwd: resolve("."), encoding: "utf8" });
  return `${revision.stdout.trim()}${status.status === 0 && status.stdout.trim() !== "" ? "+dirty" : ""}`;
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

async function writeJson(path, value) {
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 });
}

const options = parseArguments(process.argv.slice(2));
if (options.help) {
  console.log(usage());
  process.exit(0);
}

const suite = JSON.parse(await readFile(options.suitePath, "utf8"));
validateNativeVisualSuite(suite);
const cases = flattenNativeVisualCases(suite);
await validateAssets(options.suitePath, suite, cases);
const fontsDigest = await fontSetDigest(options.fontRoot, options.fontManifest);
if (fontsDigest !== suite.candidateFingerprint.fontSetDigest) {
  throw new Error("Candidate font set does not match candidateFingerprint.fontSetDigest");
}
console.log(`Validated ${cases.length} native visual units; fontSetDigest=${fontsDigest}`);
if (options.validateOnly) process.exit(0);

await mkdir(options.output, { recursive: true });
const server = await startServer(dirname(options.suitePath), options.fontRoot);
let browser;
try {
  browser = await chromium.launch(options.browserExecutable === undefined
    ? { channel: "chrome", headless: true }
    : { executablePath: options.browserExecutable, headless: true });
  const context = await browser.newContext({
    locale: suite.candidateFingerprint.locale,
    timezoneId: suite.candidateFingerprint.timezone,
    deviceScaleFactor: 1,
  });
  const probe = await context.newPage();
  const browserEnvironment = await probe.evaluate(() => ({
    locale: navigator.language,
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
    devicePixelRatio,
  }));
  await probe.close();
  const actualFingerprint = {
    os: osName(),
    osVersion: macOsVersion(),
    architecture: arch(),
    locale: browserEnvironment.locale,
    timezone: browserEnvironment.timezone,
    colorSpace: "srgb",
    devicePixelRatio: browserEnvironment.devicePixelRatio,
    scale: suite.candidateFingerprint.scale,
    background: "#ffffff",
    fontSetDigest: fontsDigest,
    browser: "Chromium",
    browserVersion: browser.version(),
    buildRevision: process.env.GIT_COMMIT ?? buildRevision(),
  };
  const fingerprintPath = resolve(options.output, "current-candidate-fingerprint.json");
  await writeJson(fingerprintPath, actualFingerprint);

  for (const testCase of cases) {
    const page = await context.newPage();
    const browserErrors = [];
    page.on("console", (message) => {
      if (message.type() === "error") browserErrors.push(`console: ${message.text()}`);
    });
    page.on("pageerror", (cause) => browserErrors.push(`pageerror: ${cause.message}`));
    const url = new URL("examples/visual-harness.html", server.baseUrl);
    url.searchParams.set("fixture", testCase.fixture);
    url.searchParams.set("unit", String(testCase.unitIndex));
    url.searchParams.set("scale", String(suite.candidateFingerprint.scale));
    if (testCase.sheetRange !== undefined) url.searchParams.set("sheetRange", testCase.sheetRange);
    if (options.fontManifest !== undefined) url.searchParams.set("fontManifest", options.fontManifest);
    await page.goto(url.href, { waitUntil: "domcontentloaded", timeout: 90_000 });
    await page.waitForFunction(() => ["pass", "error"].includes(document.body.dataset.state), undefined, { timeout: 90_000 });
    const captured = await page.evaluate(() => ({
      harness: globalThis.visualHarness,
      dataUrl: document.querySelector("#candidate-export")?.href,
    }));
    await page.close();
    if (captured.harness?.state !== "pass") {
      throw new Error(`${testCase.id} harness failed: ${captured.harness?.message ?? "unknown error"}`);
    }
    if (browserErrors.length !== 0) throw new Error(`${testCase.id} browser errors: ${browserErrors.join("; ")}`);
    if (typeof captured.dataUrl !== "string" || !captured.dataUrl.startsWith("data:image/png;base64,")) {
      throw new Error(`${testCase.id} did not export a PNG`);
    }
    const pngBytes = Buffer.from(captured.dataUrl.slice("data:image/png;base64,".length), "base64");
    const pngPath = resolveSafeRelativePath(options.output, testCase.actualPng, `${testCase.id}.actualPng`);
    await mkdir(dirname(pngPath), { recursive: true });
    await writeFile(pngPath, pngBytes, { mode: 0o600 });
    const capture = {
      schemaVersion: 1,
      fixtureSha256: captured.harness.fixtureSha256,
      format: captured.harness.format,
      unitIndex: captured.harness.unitIndex,
      unitType: captured.harness.unitType,
      unitName: captured.harness.unitName,
      unitCount: captured.harness.unitCount,
      width: captured.harness.width,
      height: captured.harness.height,
      pngSha256: sha256(pngBytes),
      diagnostics: captured.harness.diagnostics,
      ...(testCase.sheetRange === undefined ? {} : {
        sheetRange: captured.harness.sheetRange,
        viewport: captured.harness.viewport,
      }),
    };
    await writeJson(
      resolveSafeRelativePath(options.output, testCase.actualObservationJson, `${testCase.id}.actualObservationJson`),
      capture,
    );
    console.log(`Captured ${testCase.id} ${capture.width}x${capture.height} ${capture.pngSha256}`);
  }
  await context.close();
  await browser.close();
  browser = undefined;

  const reports = resolve(options.output, "reports");
  const comparison = spawnSync(process.execPath, [
    "scripts/run-external-visual-suite.mjs",
    options.suitePath,
    "--fingerprint", fingerprintPath,
    "--actual-root", options.output,
    "--output", reports,
  ], { cwd: resolve("."), encoding: "utf8", stdio: "inherit" });
  if (comparison.error !== undefined) throw comparison.error;
  if (comparison.status !== 0) process.exitCode = comparison.status ?? 1;
} finally {
  await browser?.close();
  await stopProcess(server.child);
}
