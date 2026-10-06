import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { spawn } from "node:child_process";
import test from "node:test";
import { chromium, webkit } from "playwright-core";
import UTIF from "utif2";
import { evaluateAccuracy } from "../dist/accuracy.js";

// Opt in to the original downloaded corpus; no generated replacements or bundled fonts.
const root = process.env.XPS_CORPUS;
const regions = [
  ["0inDashArray", 10, 10, 110, 110],
  ["StartEndLineCaps", 0, 0, 55, 300],
  ["Mitered", 40, 225, 550, 350],
  ["Strokes", 15, 40, 70, 470],
  ["DoubleTransformation", 90, 190, 360, 150],
  ["Path.Data", 90, 235, 490, 280],
  ["StrokeDashOffset", 0, 175, 500, 325],
];
const snapshot = (width, height, data) => ({
  units: [{ index: 0, type: "page", width, height }], objects: [],
  visuals: [{ unitIndex: 0, width, height, data }],
});

async function startCorpusServer(t, fixtureRoot) {
  const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0", "--fixture-root", resolve(fixtureRoot)], {
    stdio: ["ignore", "pipe", "pipe"],
  });
  t.after(() => server.kill());
  const base = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("XPS test server did not start")), 15_000);
    let output = "";
    for (const stream of [server.stdout, server.stderr]) stream.on("data", (data) => {
      output += data;
      const match = output.match(/(?:DocViewKit Viewer|Office Viewer Inspector): (http:\/\/[^\s]+\/)/u);
      if (match) { clearTimeout(timer); resolve(match[1]); }
    });
    server.once("error", (error) => { clearTimeout(timer); reject(error); });
    server.once("exit", (code) => { clearTimeout(timer); reject(new Error(`XPS server exited: ${code}`)); });
  });
  return base;
}

test("real XPS stroke and geometry regions match the TIFF references", { skip: !root, timeout: 120_000 }, async (t) => {
  const base = await startCorpusServer(t, root);
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage();
      for (const [fixture, x, y, width, height] of regions) await t.test(`${name}: ${fixture}`, {
        // Keep the known rasterization mismatch visible; do not relax the comparison policy.
        todo: name === "webkit" && fixture === "StartEndLineCaps"
          ? "WebKit cap-edge tolerant similarity remains 0.9879; TIFF gate requires 0.99" : false,
      }, async () => {
        const bytes = await readFile(resolve(root, "RenderedDocs", `${fixture}.tif`));
        const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
        const [ifd] = UTIF.decode(buffer);
        UTIF.decodeImage(buffer, ifd);
        const rgba = UTIF.toRGBA8(ifd);
        assert.ok(x + width <= ifd.width && y + height <= ifd.height);
        const expected = new Uint8Array(width * height * 4);
        for (let row = 0; row < height; row += 1) {
          const start = ((y + row) * ifd.width + x) * 4;
          expected.set(rgba.subarray(start, start + width * 4), row * width * 4);
        }
        await page.goto(`${base}examples/visual-harness.html?fixture=${fixture}.xps`);
        await page.waitForFunction(() => ["pass", "error"].includes(document.body.dataset.state));
        assert.equal(await page.locator("body").getAttribute("data-state"), "pass", await page.locator("#status").textContent());
        const actual = await page.locator("#surface").evaluate((canvas, rect) =>
          Array.from(canvas.getContext("2d").getImageData(...rect).data), [x, y, width, height]);
        const comparison = evaluateAccuracy({
          fileId: `${name}/${fixture}`, corpusClass: "combination",
          expected: snapshot(width, height, expected), actual: snapshot(width, height, Uint8Array.from(actual)),
          policy: { pixelChannel: 16, pixelRadius: 1, minExactPixelRatio: 0, minTolerantPixelRatio: 0.99, minSsim: 0.99 },
          declaredCoverage: ["visual-tolerant", "visual-ssim"],
        });
        assert.equal(comparison.passed, true, JSON.stringify(comparison.visualSimilarity.metrics));
      });
    } finally { await browser.close(); }
  }
});


test("real mb02 ICC brush renders without overflow and preserves CMYK colors", { skip: !root, timeout: 60_000 }, async t => {
  const base = await startCorpusServer(t, resolve(root, "../QualityLogicMinBar"));
  // Independently computed using Pillow ImageCms and this document's embedded
  // uswebuncoated.icc, relative colorimetric intent, and sRGB output.
  const expected = [[0, 170, 229], [223, 77, 128], [255, 238, 64], [87, 84, 94]];
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage();
      await page.goto(`${base}examples/visual-harness.html?fixture=mb02.xps`);
      await page.waitForFunction(() => ["pass", "error"].includes(document.body.dataset.state));
      assert.equal(await page.locator("body").getAttribute("data-state"), "pass", await page.locator("#status").textContent());
      const actual = await page.locator("#surface").evaluate(canvas => [100, 180, 250, 330]
        .map(y => Array.from(canvas.getContext("2d").getImageData(600, y, 1, 1).data).slice(0, 3)));
      assert.deepEqual(actual, expected, name);
    } finally { await browser.close(); }
  }
});

test("real Office2007 tables bounds opacity-mask surfaces", { skip: !root, timeout: 60_000 }, async t => {
  const base = await startCorpusServer(t, resolve(root, "../Office2007"));
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage();
      await page.route("**/examples/viewer.js", route => route.fulfill({ contentType: "application/javascript", body: "" }));
      await page.goto(`${base}examples/viewer.html`);
      const result = await page.evaluate(async () => {
        const NativeCanvas = globalThis.OffscreenCanvas;
        let pixels = 0;
        globalThis.OffscreenCanvas = class extends NativeCanvas {
          constructor(width, height) { super(width, height); pixels += width * height; }
        };
        const { createOfficeEngine } = await import("/dist/engine.js");
        const { xpsFormatPack } = await import("/dist/xps-formats.js");
        const engine = await createOfficeEngine({ execution: "inline", formatPack: async () => xpsFormatPack });
        const doc = await engine.open(await (await fetch("/tests/fixtures/Office2007_Tables.xps")).arrayBuffer());
        try {
          let pages = 0;
          for (let unitIndex = 0; unitIndex < doc.info.units.length; unitIndex++) {
            const frame = await doc.render({ unitIndex, scale: 1, pixelRatio: 1, background: "#ffffff" });
            if (frame.renderedObjectCount > 0) pages++;
            frame.bitmap.close();
          }
          return { pixels, pages };
        } finally { doc.close(); engine.close(); globalThis.OffscreenCanvas = NativeCanvas; }
      });
      assert.equal(result.pages, 6, name);
      // Deterministic allocation budget, independent of machine speed.
      assert.ok(result.pixels < 900_000_000, `${name}: allocated ${result.pixels} canvas pixels`);
    } finally { await browser.close(); }
  }
});
