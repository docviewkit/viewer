import { spawn } from "node:child_process";
import { readFile, writeFile } from "node:fs/promises";
import { basename, resolve } from "node:path";

import playwright from "playwright-core";

function argumentsOf(values) {
  const options = { inputs: [], samples: 3, tolerance: 1.25, unit: 1 };
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (value === "--browser") options.browser = values[++index];
    else if (value === "--baseline") options.baseline = values[++index];
    else if (value === "--write-baseline") options.writeBaseline = values[++index];
    else if (value === "--samples") options.samples = Number(values[++index]);
    else if (value === "--unit") options.unit = Number(values[++index]);
    else if (value === "--all-query") options.allQuery = true;
    else if (value === "--inline") options.inline = true;
    else if (value === "--cpu-throttle") options.cpuThrottle = Number(values[++index]);
    else if (value === "--max-open-ms") options.maxOpenMs = Number(values[++index]);
    else if (value === "--tolerance") options.tolerance = Number(values[++index]);
    else if (value === "--help") options.help = true;
    else if (value.startsWith("--")) throw new Error(`Unknown option: ${value}`);
    else options.inputs.push(resolve(value));
  }
  if (!Number.isInteger(options.samples) || options.samples < 1 || options.samples > 20) {
    throw new Error("--samples must be an integer from 1 through 20");
  }
  if (!Number.isFinite(options.tolerance) || options.tolerance < 1) {
    throw new Error("--tolerance must be at least 1");
  }
  if (!Number.isInteger(options.unit) || options.unit < 1) {
    throw new Error("--unit must be a positive 1-based unit number");
  }
  if (options.maxOpenMs !== undefined && (!Number.isFinite(options.maxOpenMs) || options.maxOpenMs <= 0)) {
    throw new Error("--max-open-ms must be a positive number");
  }
  if (options.cpuThrottle !== undefined
    && (!Number.isFinite(options.cpuThrottle) || options.cpuThrottle < 1 || options.cpuThrottle > 20)) {
    throw new Error("--cpu-throttle must be between 1 and 20");
  }
  return options;
}

function usage() {
  return [
    "Usage: npm run performance -- [options] <document-file>...",
    "  --browser <path>          Chrome/Chromium executable",
    "  --samples <count>         Median sample count (default: 3)",
    "  --unit <number>           1-based page/sheet/slide to render (default: 1)",
    "  --all-query               Measure a document-wide object query after rendering",
    "  --inline                  Run Wasm inline for CPU-throttling diagnostics",
    "  --cpu-throttle <rate>     Apply Chromium CPU slowdown (1 through 20)",
    "  --max-open-ms <number>    Fail when median document open time exceeds this budget",
    "  --baseline <json>         Fail on open/render/first-render regression",
    "  --write-baseline <json>   Write the current result as baseline",
    "  --tolerance <ratio>       Allowed regression ratio (default: 1.25)",
  ].join("\n");
}

function startServer() {
  const child = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], {
    cwd: process.cwd(),
    stdio: ["ignore", "pipe", "pipe"],
  });
  return new Promise((resolveServer, reject) => {
    let output = "";
    const timer = setTimeout(() => reject(new Error("Inspector server did not start within 10 seconds")), 10_000);
    const finish = (cause, url) => {
      clearTimeout(timer);
      if (cause !== undefined) reject(cause);
      else resolveServer({ child, url });
    };
    child.stdout.on("data", (chunk) => {
      output += chunk;
      const url = output.match(/(?:DocViewKit Viewer|Office Viewer Inspector): (http:\/\/[^\s]+)/u)?.[1];
      if (url !== undefined) finish(undefined, url);
    });
    child.stderr.on("data", (chunk) => { output += chunk; });
    child.once("exit", (code) => finish(new Error(`Inspector server exited (${code}): ${output.trim()}`)));
    child.once("error", (cause) => finish(cause));
  });
}

async function measure(page, path, sampleCount, unitNumber, allQuery) {
  const file = await readFile(path);
  return page.evaluate(async ({ base64, fileName, path, sampleCount, unitNumber, allQuery }) => {
    const encoded = atob(base64);
    const source = new Uint8Array(encoded.length);
    for (let index = 0; index < encoded.length; index += 1) source[index] = encoded.charCodeAt(index);
    const samples = [];
    for (let iteration = 0; iteration < sampleCount; iteration += 1) {
      const started = performance.now();
      const document = await globalThis.performanceEngine.open(source.buffer.slice(0), {
        fileName,
        transferInput: true,
      });
      const opened = performance.now();
      const unit = document.info.units[unitNumber - 1];
      if (unit === undefined) throw new Error(`${path} has no unit ${unitNumber}`);
      const viewport = unit.type === "sheet"
        ? { x: 0, y: 0, width: Math.min(1_200, unit.width), height: Math.min(900, unit.height) }
        : undefined;
      const area = (viewport?.width ?? unit.width) * (viewport?.height ?? unit.height);
      const scale = Math.min(1, Math.sqrt(1_200_000 / Math.max(1, area)));
      const frame = await document.render({
        unitIndex: unit.index,
        ...(viewport === undefined ? {} : { viewport }),
        scale,
        pixelRatio: 1,
        background: "#fff",
      }, { priority: "interactive", supersedeKey: "performance-suite" });
      const rendered = performance.now();
      frame.bitmap.close();
      const repeatedFrame = await document.render({
        unitIndex: unit.index,
        ...(viewport === undefined ? {} : { viewport }),
        scale,
        pixelRatio: 1,
        background: "#fff",
      }, { priority: "interactive", supersedeKey: "performance-suite" });
      const repeated = performance.now();
      repeatedFrame.bitmap.close();
      const nextUnit = document.info.units[unitNumber];
      let nextPageRenderMs;
      if (nextUnit?.type === "page") {
        const nextStarted = performance.now();
        const nextFrame = await document.render({
          unitIndex: nextUnit.index,
          scale,
          pixelRatio: 1,
          background: "#fff",
        }, { priority: "interactive", supersedeKey: "performance-suite" });
        nextPageRenderMs = performance.now() - nextStarted;
        nextFrame.bitmap.close();
      }
      const queryStarted = performance.now();
      const objects = await document.listObjects({
        unitIndex: unit.index,
        ...(viewport === undefined ? {} : { viewport }),
      });
      const queried = performance.now();
      let allQueryMs;
      let allObjectCount;
      if (allQuery) {
        const allQueryStarted = performance.now();
        allObjectCount = (await document.listObjects()).length;
        allQueryMs = performance.now() - allQueryStarted;
      }
      const format = document.info.format;
      document.close();
      samples.push({
        format,
        openMs: opened - started,
        firstRenderMs: rendered - started,
        renderMs: rendered - opened,
        repeatRenderMs: repeated - rendered,
        queryMs: queried - queryStarted,
        ...(allQueryMs === undefined ? {} : { allQueryMs, allObjectCount }),
        ...(nextPageRenderMs === undefined ? {} : { nextPageRenderMs }),
        objectCount: objects.length,
      });
    }
    const median = (metric) => {
      const values = samples.map((sample) => sample[metric]).sort((left, right) => left - right);
      return values[Math.floor(values.length / 2)];
    };
    return {
      path,
      format: samples[0]?.format,
      bytes: source.byteLength,
      unit: unitNumber,
      openMs: median("openMs"),
      firstRenderMs: median("firstRenderMs"),
      renderMs: median("renderMs"),
      repeatRenderMs: median("repeatRenderMs"),
      ...(samples.every((sample) => sample.nextPageRenderMs !== undefined)
        ? { nextPageRenderMs: median("nextPageRenderMs") }
        : {}),
      queryMs: median("queryMs"),
      ...(samples.every((sample) => sample.allQueryMs !== undefined)
        ? { allQueryMs: median("allQueryMs"), allObjectCount: median("allObjectCount") }
        : {}),
      objectCount: median("objectCount"),
    };
  }, { base64: file.toString("base64"), fileName: basename(path), path, sampleCount, unitNumber, allQuery });
}

function verifyBaseline(results, baseline, tolerance) {
  if (baseline?.version !== 2 || !Array.isArray(baseline.results)) {
    throw new Error("Performance baseline is not schema v2; regenerate it with --write-baseline");
  }
  const failures = [];
  const expected = new Map(baseline.results.map((result) => [result.path, result]));
  for (const result of results) {
    const prior = expected.get(result.path);
    if (prior === undefined) continue;
    for (const metric of ["openMs", "renderMs", "firstRenderMs"]) {
      if (!Number.isFinite(prior[metric])) {
        failures.push(`${result.path} baseline ${metric} is missing or invalid; regenerate it`);
        continue;
      }
      const limit = Math.max(prior[metric] * tolerance, prior[metric] + 5);
      if (result[metric] > limit) failures.push(`${result.path} ${metric} ${result[metric].toFixed(1)}ms > ${limit.toFixed(1)}ms`);
    }
  }
  if (failures.length > 0) throw new Error(`Performance regression:\n${failures.join("\n")}`);
}

function verifyOpenBudget(results, maximum) {
  if (maximum === undefined) return;
  const failures = results
    .filter((result) => result.openMs > maximum)
    .map((result) => `${result.path} openMs ${result.openMs.toFixed(1)}ms > ${maximum.toFixed(1)}ms`);
  if (failures.length > 0) throw new Error(`Open-time budget exceeded:\n${failures.join("\n")}`);
}

const options = argumentsOf(process.argv.slice(2));
if (options.help) {
  console.log(usage());
  process.exit(0);
}
if (options.inputs.length === 0) throw new Error(`At least one document file is required.\n${usage()}`);

const server = await startServer();
let browser;
try {
  browser = await playwright.chromium.launch(options.browser === undefined
    ? { channel: "chrome", headless: true }
    : { executablePath: options.browser, headless: true });
  const page = await browser.newPage();
  if (options.cpuThrottle !== undefined) {
    const session = await page.context().newCDPSession(page);
    await session.send("Emulation.setCPUThrottlingRate", { rate: options.cpuThrottle });
  }
  await page.goto(`${server.url}examples/inspector.html`);
  await page.evaluate(async ({ inline }) => {
    const { createOfficeEngine } = await import("/dist/engine.js");
    globalThis.performanceEngine = await createOfficeEngine({
      ...(inline ? { execution: "inline" } : {}),
      formatPack: () => import("/dist/extended-formats.js")
        .then(({ extendedFormatPack }) => extendedFormatPack),
    });
  }, { inline: options.inline === true });
  const results = [];
  for (const input of options.inputs) {
    results.push(await measure(page, input, options.samples, options.unit, options.allQuery === true));
  }
  verifyOpenBudget(results, options.maxOpenMs);
  await page.evaluate(() => globalThis.performanceEngine.close());
  const report = {
    version: 2,
    samples: options.samples,
    execution: options.inline === true ? "inline" : "worker",
    ...(options.cpuThrottle === undefined ? {} : { cpuThrottle: options.cpuThrottle }),
    results,
  };
  if (options.baseline !== undefined) {
    verifyBaseline(results, JSON.parse(await readFile(resolve(options.baseline), "utf8")), options.tolerance);
  }
  if (options.writeBaseline !== undefined) {
    await writeFile(resolve(options.writeBaseline), `${JSON.stringify(report, null, 2)}\n`);
  }
  console.log(JSON.stringify(report, null, 2));
} finally {
  await browser?.close();
  server.child.kill("SIGTERM");
}
