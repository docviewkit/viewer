import { spawn } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { dirname, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

import playwright from "playwright-core";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const pdfRoot = resolve(root, "tests/external/pdfjs/test/pdfs");
const pdfjsRoot = resolve(root, "tests/external/pdfjs");
const pdfjsWebRoot = resolve(pdfjsRoot, "build/generic/web");
const maximumPixels = 1_200_000;
const defaultCases = [
  { file: "S2.pdf", page: 1, feature: "vector and text page" },
  { file: "bug1721218_reduced.pdf", page: 1, feature: "repeated vector artwork" },
  { file: "tracemonkey.pdf", page: 6, feature: "complex text and image page" },
  { file: "bomb_giant.pdf", page: 1, feature: "giant page bounds" },
  { file: "blendmode.pdf", page: 1, feature: "blend modes" },
  { file: "transparency_group.pdf", page: 1, feature: "transparency groups" },
  { file: "type4psfunc.pdf", page: 1, feature: "Type 4 shading function" },
  { file: "bug_jpx.pdf", page: 1, feature: "JPEG 2000 image" },
  { file: "XiaoBiaoSong.pdf", page: 1, feature: "embedded CJK font" },
  { file: "vertical.pdf", page: 1, feature: "vertical Japanese text" },
  { file: "bitmap-halftone-template1.pdf", page: 1, feature: "halftone image mask" },
  { file: "Brotli-Prototype-FileA.pdf", page: 1, feature: "Brotli-compressed PDF data" },
];

function argumentsOf(values) {
  const options = { samples: 7, output: resolve(root, "output/performance/pdfjs-comparison.json") };
  for (let index = 0; index < values.length; index += 1) {
    const argument = values[index];
    if (argument === "--samples") options.samples = Number(values[++index]);
    else if (argument === "--output") options.output = resolve(values[++index]);
    else if (argument === "--browser") options.browser = values[++index];
    else if (argument === "--case") options.case = values[++index];
    else if (argument === "--help") options.help = true;
    else throw new Error(`Unknown option: ${argument}`);
  }
  if (!Number.isSafeInteger(options.samples) || options.samples < 3 || options.samples > 20) {
    throw new Error("--samples must be an integer from 3 through 20");
  }
  return options;
}

function usage() {
  return [
    "Usage: node scripts/run-pdfjs-performance-comparison.mjs [options]",
    "  --samples <count>  Timed samples per renderer and PDF (default: 7)",
    "  --output <json>    Report path (default: output/performance/pdfjs-comparison.json)",
    "  --browser <path>   Chrome/Chromium executable",
    "  --case <filename>  Run one default case for profiling",
  ].join("\n");
}

function startServer() {
  const child = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], {
    cwd: root,
    stdio: ["ignore", "pipe", "pipe"],
  });
  return new Promise((resolveServer, reject) => {
    let output = "";
    const timeout = setTimeout(() => reject(new Error("Inspector server did not start within 10 seconds")), 10_000);
    const finish = (cause, url) => {
      clearTimeout(timeout);
      if (cause === undefined) resolveServer({ child, url });
      else reject(cause);
    };
    child.stdout.on("data", (chunk) => {
      output += chunk;
      const url = output.match(/(?:DocViewKit Viewer|Office Viewer Inspector): (http:\/\/[^\s]+)/u)?.[1];
      if (url !== undefined) finish(undefined, url);
    });
    child.stderr.on("data", (chunk) => { output += chunk; });
    child.once("exit", (code) => finish(new Error(`Inspector server exited (${code}): ${output.trim()}`)));
    child.once("error", finish);
  });
}

function routedPdfjsAsset(pathname) {
  const roots = [
    ["/build/", resolve(pdfjsRoot, "build/generic/build")],
    ["/cmaps/", resolve(pdfjsWebRoot, "cmaps")],
    ["/standard_fonts/", resolve(pdfjsWebRoot, "standard_fonts")],
    ["/wasm/", resolve(pdfjsWebRoot, "wasm")],
  ];
  for (const [prefix, directory] of roots) {
    if (!pathname.startsWith(prefix)) continue;
    const path = resolve(directory, pathname.slice(prefix.length));
    if (path.startsWith(`${directory}${sep}`) && existsSync(path) && statSync(path).isFile()) return path;
  }
  return undefined;
}

function contentType(path) {
  if (path.endsWith(".mjs") || path.endsWith(".js")) return "text/javascript; charset=utf-8";
  if (path.endsWith(".wasm")) return "application/wasm";
  return "application/octet-stream";
}

function encodedSource(path) {
  return readFileSync(path).toString("base64");
}

function median(values) {
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.floor(sorted.length / 2)];
}

function rounded(value) {
  return Math.round(value * 10) / 10;
}

function summarize(samples) {
  return Object.fromEntries(["openMs", "renderMs", "firstFrameMs"].map((metric) => [metric, rounded(median(
    samples.map((sample) => sample[metric]),
  ))]));
}

function advantage(officeviewer, pdfjs) {
  return rounded((pdfjs - officeviewer) / pdfjs * 100);
}

async function prepareOfficeViewer(page, serverUrl) {
  await page.goto(`${serverUrl}examples/inspector.html`, { waitUntil: "domcontentloaded" });
  await page.evaluate(async () => {
    const { createOfficeEngine } = await import("/dist/engine.js");
    globalThis.benchmarkOfficeEngine = await createOfficeEngine({
      formatPack: () => import("/dist/extended-formats.js")
        .then(({ extendedFormatPack }) => extendedFormatPack),
    });
  });
}

async function preparePdfjs(context) {
  await context.addInitScript(() => {
    Math.sumPrecise ??= (values) => {
      let sum = 0;
      let compensation = 0;
      for (const value of values) {
        const next = sum + value;
        compensation += Math.abs(sum) >= Math.abs(value)
          ? sum - next + value
          : value - next + sum;
        sum = next;
      }
      return sum + compensation;
    };
  });
  await context.route("http://pdfjs-performance.local/**", async (route) => {
    const pathname = new URL(route.request().url()).pathname;
    if (pathname === "/") {
      await route.fulfill({ contentType: "text/html", body: "<!doctype html><html><body></body></html>" });
      return;
    }
    const asset = routedPdfjsAsset(pathname);
    if (asset === undefined) {
      await route.fulfill({ status: 404, body: "Not Found" });
      return;
    }
    await route.fulfill({ contentType: contentType(asset), body: readFileSync(asset) });
  });
  const page = await context.newPage();
  await page.goto("http://pdfjs-performance.local/", { waitUntil: "domcontentloaded" });
  await page.evaluate(async () => {
    globalThis.benchmarkPdfjs = await import("/build/pdf.mjs");
    globalThis.benchmarkPdfjs.GlobalWorkerOptions.workerSrc = "/build/pdf.worker.mjs";
  });
  return page;
}

async function installSource(page, base64) {
  await page.evaluate((encodedSource) => {
    const encoded = atob(encodedSource);
    globalThis.benchmarkSource = new Uint8Array(encoded.length);
    for (let index = 0; index < encoded.length; index += 1) {
      globalThis.benchmarkSource[index] = encoded.charCodeAt(index);
    }
  }, base64);
}

async function measureOfficeViewer(page, testCase) {
  await page.bringToFront();
  return page.evaluate(async ({ testCase, maximumPixels }) => {
    const source = globalThis.benchmarkSource.slice();
    const started = performance.now();
    const document = await globalThis.benchmarkOfficeEngine.open(source.buffer, {
      fileName: testCase.file,
      transferInput: true,
    });
    const opened = performance.now();
    const unit = document.info.units[testCase.page - 1];
    if (unit === undefined) throw new Error(`${testCase.file} has no page ${testCase.page}`);
    // OfficeViewer scene units are CSS pixels at 96 DPI, while PDF.js viewports
    // are PDF points at 72 DPI. Convert to the same physical raster dimensions.
    const pointScale = 0.75;
    const targetArea = unit.width * pointScale * unit.height * pointScale;
    const scale = pointScale * Math.min(1, Math.sqrt(maximumPixels / Math.max(1, targetArea)));
    const frame = await document.render({
      unitIndex: unit.index,
      scale,
      pixelRatio: 1,
      background: "#fff",
    }, { priority: "interactive", supersedeKey: "pdfjs-performance-comparison" });
    const rendered = performance.now();
    const dimensions = { width: frame.bitmap.width, height: frame.bitmap.height };
    const renderedObjectCount = frame.renderedObjectCount;
    frame.bitmap.close();
    document.close();
    return {
      openMs: opened - started,
      renderMs: rendered - opened,
      firstFrameMs: rendered - started,
      renderedObjectCount,
      ...dimensions,
    };
  }, { testCase, maximumPixels });
}

async function inspectOfficeViewer(page, testCase) {
  await page.bringToFront();
  return page.evaluate(async (currentCase) => {
    const document = await globalThis.benchmarkOfficeEngine.open(globalThis.benchmarkSource.slice().buffer, {
      fileName: currentCase.file,
      transferInput: true,
    });
    const objects = await document.listObjects({ unitIndex: currentCase.page - 1 });
    document.close();
    const types = Object.fromEntries([...objects.reduce((counts, object) => {
      counts.set(object.type, (counts.get(object.type) ?? 0) + 1);
      return counts;
    }, new Map()).entries()].sort(([left], [right]) => left.localeCompare(right)));
    const textObjects = objects.filter(({ text }) => typeof text === "string" && text.length > 0);
    return {
      objectCount: objects.length,
      types,
      textObjectCount: textObjects.length,
      textCodeUnits: textObjects.reduce((sum, object) => sum + object.text.length, 0),
    };
  }, testCase);
}

async function measurePdfjs(page, testCase) {
  await page.bringToFront();
  return page.evaluate(async ({ testCase, maximumPixels }) => {
    const source = globalThis.benchmarkSource.slice();
    const started = performance.now();
    const loadingTask = globalThis.benchmarkPdfjs.getDocument({
      data: source,
      cMapUrl: "/cmaps/",
      cMapPacked: true,
      standardFontDataUrl: "/standard_fonts/",
      wasmUrl: "/wasm/",
      useWasm: true,
      useSystemFonts: true,
      enableXfa: true,
      isEvalSupported: false,
    });
    const pdfDocument = await loadingTask.promise;
    const pdfPage = await pdfDocument.getPage(testCase.page);
    const baseViewport = pdfPage.getViewport({ scale: 1 });
    const scale = Math.min(1, Math.sqrt(maximumPixels / Math.max(1, baseViewport.width * baseViewport.height)));
    const viewport = pdfPage.getViewport({ scale });
    const opened = performance.now();
    const canvas = globalThis.document.createElement("canvas");
    canvas.width = Math.max(1, Math.ceil(viewport.width));
    canvas.height = Math.max(1, Math.ceil(viewport.height));
    await pdfPage.render({
      canvasContext: canvas.getContext("2d"),
      viewport,
      background: "#ffffff",
      intent: "display",
    }).promise;
    const rendered = performance.now();
    const dimensions = { width: canvas.width, height: canvas.height };
    pdfPage.cleanup();
    await loadingTask.destroy();
    return {
      openMs: opened - started,
      renderMs: rendered - opened,
      firstFrameMs: rendered - started,
      ...dimensions,
    };
  }, { testCase, maximumPixels });
}

const options = argumentsOf(process.argv.slice(2));
if (options.help) {
  console.log(usage());
  process.exit(0);
}

const selectedCases = options.case === undefined
  ? defaultCases
  : defaultCases.filter(({ file }) => file === options.case);
if (selectedCases.length === 0) throw new Error(`Unknown default case: ${options.case}`);
const required = [
  resolve(root, "dist/engine.js"),
  resolve(root, "dist/extended-formats.js"),
  resolve(pdfjsRoot, "build/generic/build/pdf.mjs"),
  resolve(pdfjsRoot, "build/generic/build/pdf.worker.mjs"),
  ...selectedCases.map(({ file }) => resolve(pdfRoot, file)),
];
const missing = required.filter((path) => !existsSync(path));
if (missing.length > 0) throw new Error(`Missing benchmark inputs:\n${missing.join("\n")}\nRun npm run build and build pdf.js generic assets first.`);

const server = await startServer();
let browser;
try {
  browser = await playwright.chromium.launch(options.browser === undefined
    ? { channel: "chrome", headless: true }
    : { executablePath: options.browser, headless: true });
  const context = await browser.newContext({
    locale: "en-US",
    timezoneId: "Asia/Shanghai",
    deviceScaleFactor: 1,
  });
  const officePage = await context.newPage();
  await prepareOfficeViewer(officePage, server.url);
  const pdfjsPage = await preparePdfjs(context);
  const results = [];
  for (const testCase of selectedCases) {
    const path = resolve(pdfRoot, testCase.file);
    const base64 = encodedSource(path);
    await installSource(officePage, base64);
    await installSource(pdfjsPage, base64);
    await measureOfficeViewer(officePage, testCase);
    await measurePdfjs(pdfjsPage, testCase);
    const scene = await inspectOfficeViewer(officePage, testCase);
    const officeviewerSamples = [];
    const pdfjsSamples = [];
    for (let sample = 0; sample < options.samples; sample += 1) {
      const order = sample % 2 === 0 ? ["officeviewer", "pdfjs"] : ["pdfjs", "officeviewer"];
      for (const renderer of order) {
        const result = renderer === "officeviewer"
          ? await measureOfficeViewer(officePage, testCase)
          : await measurePdfjs(pdfjsPage, testCase);
        (renderer === "officeviewer" ? officeviewerSamples : pdfjsSamples).push(result);
      }
    }
    const officeviewer = summarize(officeviewerSamples);
    const pdfjs = summarize(pdfjsSamples);
    results.push({
      ...testCase,
      bytes: statSync(path).size,
      scene,
      officeviewer,
      pdfjs,
      samples: { officeviewer: officeviewerSamples, pdfjs: pdfjsSamples },
      firstFrameAdvantagePct: advantage(officeviewer.firstFrameMs, pdfjs.firstFrameMs),
      renderAdvantagePct: advantage(officeviewer.renderMs, pdfjs.renderMs),
      officeviewerDimensions: {
        width: officeviewerSamples[0].width,
        height: officeviewerSamples[0].height,
      },
      pdfjsDimensions: { width: pdfjsSamples[0].width, height: pdfjsSamples[0].height },
    });
    console.log(`${testCase.file}: OfficeViewer ${officeviewer.firstFrameMs}ms, pdf.js ${pdfjs.firstFrameMs}ms, advantage ${results.at(-1).firstFrameAdvantagePct}%`);
  }
  await officePage.evaluate(() => globalThis.benchmarkOfficeEngine.close());
  const geometricRatio = Math.exp(results.reduce(
    (sum, result) => sum + Math.log(result.officeviewer.firstFrameMs / result.pdfjs.firstFrameMs),
    0,
  ) / results.length);
  const geometricMeanAdvantagePct = rounded((1 - geometricRatio) * 100);
  const passingCaseCount = results.filter((result) => result.firstFrameAdvantagePct >= 10).length;
  const report = {
    version: 1,
    generatedAt: new Date().toISOString(),
    samples: options.samples,
    environment: { browser: "Chromium", deviceScaleFactor: 1, maximumPixels },
    criterion: "OfficeViewer geometric-mean open-to-first-frame must be at least 10% faster than pdf.js",
    caseCriterion: "A case passes when its median open-to-first-frame is at least 10% faster than pdf.js",
    geometricMeanAdvantagePct,
    passing: geometricMeanAdvantagePct >= 10,
    passingCaseCount,
    caseCount: results.length,
    results,
  };
  mkdirSync(dirname(options.output), { recursive: true });
  writeFileSync(options.output, `${JSON.stringify(report, null, 2)}\n`, { mode: 0o600 });
  console.log(JSON.stringify(report, null, 2));
} finally {
  await browser?.close();
  server.child.kill("SIGTERM");
}
