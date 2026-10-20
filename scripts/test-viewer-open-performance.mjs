// Real local documents stay outside the repository. Measures the public Viewer,
// including file read, parsing, font loading, layout and the first visible paint.
// node scripts/test-viewer-open-performance.mjs --output report.json <file>...
//   [--dist directory] [--baseline report.json] [--samples 3] [--budget-ms 2000]
//   [--all-pages] [--inline] [--reference-wasm file | --reference-dist directory] [--max-reference-ratio 0.9]
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { createServer } from "node:http";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { resolve, dirname, basename, extname, sep } from "node:path";
import { chromium, webkit } from "playwright-core";

const options = { dist: "dist", samples: 3, budgetMs: 2000, webkitTolerance: 0, maxReferenceRatio: 0.9, inputs: [] };
const keys = { "--dist": "dist", "--reference-wasm": "referenceWasm", "--reference-dist": "referenceDist", "--max-reference-ratio": "maxReferenceRatio", "--output": "output", "--baseline": "baseline", "--samples": "samples", "--budget-ms": "budgetMs", "--webkit-channel-tolerance": "webkitTolerance" };
for (let i = 2; i < process.argv.length; i++) {
  const value = process.argv[i];
  if (value === "--all-pages") options.allPages = true;
  else if (value === "--inline") options.inline = true;
  else if (keys[value]) options[keys[value]] = process.argv[++i];
  else { assert.ok(!value.startsWith("--"), `Unknown option ${value}`); options.inputs.push(resolve(value)); }
}
assert.ok(options.output && options.inputs.length, "Supply --output report.json and local document paths");
options.samples = Number(options.samples); options.budgetMs = Number(options.budgetMs);
options.maxReferenceRatio = Number(options.maxReferenceRatio);
assert.ok(Number.isFinite(options.maxReferenceRatio) && options.maxReferenceRatio > 0 && options.maxReferenceRatio < 1);
assert.ok(Number.isInteger(options.samples) && options.samples > 0 && options.samples <= 20);
assert.ok(Number.isFinite(options.budgetMs) && options.budgetMs > 0);
const root = resolve("."), dist = resolve(options.dist), output = resolve(options.output);
await mkdir(dirname(output), { recursive: true });
assert.ok(!(options.referenceWasm && options.referenceDist), "Choose one reference runtime");
const referenceCore = options.referenceWasm ?? (options.referenceDist ? resolve(options.referenceDist, "office-viewer-core.wasm") : undefined);
const referenceWasmSha256 = referenceCore
  ? createHash("sha256").update(await readFile(referenceCore)).digest("hex") : undefined;
const baseline = options.baseline ? JSON.parse(await readFile(options.baseline, "utf8")) : undefined;
const server = createServer(async (request, response) => {
  try {
    const pathname = decodeURIComponent(new URL(request.url, "http://localhost").pathname);
    const selectedDist = options.referenceDist && request.headers["x-office-viewer-reference"] === "true"
      ? resolve(options.referenceDist) : dist;
    const base = pathname.startsWith("/dist/") ? selectedDist : root;
    const path = resolve(base, "." + (base === selectedDist ? pathname.slice(5) : pathname));
    assert.ok(path.startsWith(base + sep));
    response.setHeader("Content-Type", ({ ".js": "text/javascript", ".html": "text/html", ".wasm": "application/wasm" })[extname(path)] ?? "application/octet-stream");
    const data = await readFile(path);
    response.end(options.inline && path === resolve(root, "examples/viewer.js")
      ? data.toString().replace("engine: {", 'engine: { execution: "inline",') : data);
  } catch { response.writeHead(404).end(); }
});
await new Promise(resolveServer => server.listen(0, "127.0.0.1", resolveServer));
const results = [];
assert.ok([0, 1].includes(Number(options.webkitTolerance)), "WebKit channel tolerance must be zero or one");
try {
  for (const [browserName, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      for (const path of options.inputs) {
        const result = { browser: browserName, file: basename(path), sha256: createHash("sha256").update(await readFile(path)).digest("hex"), samples: [], pages: [], pixelDifferences: [], errors: [], ...(referenceCore ? { referenceSamples: [] } : {}) };
        const variants = referenceCore ? 2 : 1;
        for (let attempt = 0; attempt < options.samples * variants; attempt++) {
          const sample = Math.floor(attempt / variants);
          // Alternate order to compare the same original file under the same load.
          const reference = variants === 2 && sample % 2 === attempt % 2;
          const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, reducedMotion: "reduce" });
          try {
            if (reference && options.referenceDist) await page.setExtraHTTPHeaders({ "x-office-viewer-reference": "true" });
            else if (reference) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: resolve(referenceCore), contentType: "application/wasm" }));
            page.on("pageerror", error => result.errors.push(error.message));
            page.on("console", message => { if (message.type() === "error") result.errors.push(message.text()); });
            await page.goto(`http://127.0.0.1:${server.address().port}/examples/viewer.html?filePicker=true`);
            await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
            await page.evaluate(inline => {
              const viewer = document.querySelector("#viewer");
              delete document.documentElement.dataset.ready;
              viewer.addEventListener("docviewkit-ready", () => {
                const ready = performance.now();
                requestAnimationFrame(() => { globalThis.openTiming = { readyMs: ready - globalThis.openStarted, paintMs: performance.now() - globalThis.openStarted }; });
              }, { once: true });
              globalThis.openStarted = performance.now();
            }, options.inline === true);
            await page.locator("#viewer-file-input").setInputFiles(path);
            await page.waitForFunction(() => globalThis.openTiming || document.documentElement.dataset.error, undefined, { timeout: 180000 });
            const timing = await page.evaluate(() => ({ ...globalThis.openTiming, error: document.documentElement.dataset.error, state: document.querySelector("#viewer").state.status }));
            assert.equal(timing.error, undefined); assert.equal(timing.state, "ready");
            (reference ? result.referenceSamples : result.samples).push(timing);
            if (!reference && sample === 0 && options.allPages) {
              const count = await page.evaluate(() => document.querySelector("#viewer").state.info.units.length);
              for (let index = 0; index < count; index++) {
                await page.evaluate(index => document.querySelector("#viewer").reveal({ kind: "unit", unitIndex: index }), index);
                await page.waitForFunction(index => document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${index}"]`)?.dataset.rendered === "true", index, { timeout: 180000 });
                result.pages.push(await page.evaluate(async index => {
                  const canvas = document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${index}"] canvas`);
                  const data = canvas.getContext("2d").getImageData(0, 0, canvas.width, canvas.height).data;
                  const hash = [...new Uint8Array(await crypto.subtle.digest("SHA-256", data))].map(value => value.toString(16).padStart(2, "0")).join("");
                  return { index, width: canvas.width, height: canvas.height, hash };
                }, index));
                const png = await page.evaluate(index => document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${index}"] canvas`).toDataURL("image/png").split(",")[1], index);
                await writeFile(resolve(dirname(output), `${basename(output, ".json")}-${browserName}-${result.sha256.slice(0, 8)}-page${index + 1}-canvas.png`), Buffer.from(png, "base64"));
                const reference = baseline?.results.find(value => value.browser === browserName && value.sha256 === result.sha256)?.pages[index];
                if (reference && reference.hash !== result.pages[index].hash) {
                  const beforePath = resolve(dirname(options.baseline), `${basename(options.baseline, ".json")}-${browserName}-${result.sha256.slice(0, 8)}-page${index + 1}-canvas.png`);
                  const beforePng = (await readFile(beforePath)).toString("base64");
                  const difference = await page.evaluate(async ({ beforePng, index }) => {
                    const source = document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${index}"] canvas`);
                    const image = new Image(); image.src = `data:image/png;base64,${beforePng}`; await image.decode();
                    if (image.width !== source.width || image.height !== source.height) throw new Error("Page dimensions changed");
                    const canvas = document.createElement("canvas"); canvas.width = image.width; canvas.height = image.height;
                    const ctx = canvas.getContext("2d"); ctx.drawImage(image, 0, 0);
                    const before = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
                    const after = source.getContext("2d").getImageData(0, 0, canvas.width, canvas.height).data;
                    let maxChannelDifference = 0, changedPixels = 0;
                    for (let offset = 0; offset < before.length; offset += 4) {
                      let changed = false;
                      for (let channel = 0; channel < 4; channel++) {
                        const delta = Math.abs(before[offset + channel] - after[offset + channel]);
                        maxChannelDifference = Math.max(maxChannelDifference, delta); changed ||= delta > 0;
                      }
                      if (changed) changedPixels++;
                    }
                    return { index, maxChannelDifference, changedPixels };
                  }, { beforePng, index });
                  result.pixelDifferences.push(difference);
                }

                if (index === 0 || index === 7) await page.screenshot({ path: resolve(dirname(output), `${basename(output, ".json")}-${browserName}-${result.sha256.slice(0, 8)}-page${index + 1}.png`) });
              }
            }
          } finally { await page.close(); }
        }
        if (referenceCore) {
          result.referenceMedianMs = result.referenceSamples.map(value => value.paintMs).sort((a, b) => a - b)[Math.floor(options.samples / 2)];
        }
        result.medianMs = result.samples.map(value => value.paintMs).sort((a, b) => a - b)[Math.floor(options.samples / 2)];
        results.push(result);
        console.log(JSON.stringify({ ...result, pages: result.pages.length }));
      }
    } finally { await browser.close(); }
  }
} finally {
  server.close();
  await writeFile(output, JSON.stringify({ measurement: "local file selection to first visible Viewer paint", viewport: [1440, 1000], budgetMs: options.budgetMs, webkitChannelTolerance: Number(options.webkitTolerance), execution: options.inline ? "inline" : "auto", ...(referenceWasmSha256 ? { referenceWasmSha256, ...(options.referenceDist ? { referenceDist: resolve(options.referenceDist) } : {}), maxReferenceRatio: options.maxReferenceRatio } : {}), results }, null, 2) + "\n");
}
for (const result of results) {
  assert.deepEqual(result.errors, [], `${result.browser}/${result.file} browser errors`);
  if (referenceCore) assert.ok(result.medianMs <= result.referenceMedianMs * options.maxReferenceRatio,
    `${result.browser}/${result.file}: ${result.medianMs.toFixed(1)}ms must be <= ${options.maxReferenceRatio} × reference ${result.referenceMedianMs.toFixed(1)}ms`);
}
if (options.baseline) {
  for (const result of results) {
    const before = baseline.results.find(value => value.browser === result.browser && value.sha256 === result.sha256);
    assert.ok(before, "Baseline must use the unchanged original document");
    if (options.allPages) {
      assert.deepEqual(result.pages.map(({ hash, ...shape }) => shape), before.pages.map(({ hash, ...shape }) => shape));
      const tolerance = result.browser === "webkit" ? Number(options.webkitTolerance) : 0;
      assert.ok(result.pixelDifferences.every(diff => diff.maxChannelDifference <= tolerance),
        `${result.browser}/${result.file} pixel differences exceed ${tolerance}: ${JSON.stringify(result.pixelDifferences)}`);
    }
  }
}
assert.deepEqual(results.filter(result => result.medianMs >= options.budgetMs).map(({ browser, file, medianMs }) => ({ browser, file, medianMs })), [], "Viewer first-paint budget exceeded");
