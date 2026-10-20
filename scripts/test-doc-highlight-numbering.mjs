import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { mkdir, writeFile } from "node:fs/promises";
import { chromium, webkit } from "playwright-core";

// The confidential source stays outside the repository; supply its local path.
const source = process.env.DOC_REGRESSION_FILE;
assert.ok(source, "Set DOC_REGRESSION_FILE to the original DOC");
const output = "output/playwright/doc-highlight-numbering";
await mkdir(output, { recursive: true });
const results = [];
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10000);
    server.once("error", reject);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 1000 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.route("**/tests/fixtures/private-regression.doc", route => route.fulfill({ path: source, contentType: "application/msword" }));
      const baseline = process.env.DOC_BASELINE_WASM;
      if (baseline) await page.route("**/dist/office-viewer-legacy-office.wasm", route => route.fulfill({ path: baseline, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=private-regression.doc`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      const checks = await page.evaluate(async () => {
        const { createOfficeEngine } = await import("/dist/engine.js");
        const { extendedFormatPack } = await import("/dist/extended-formats.js");
        const engine = await createOfficeEngine({ execution: "inline", formatPack: async () => extendedFormatPack });
        const doc = await engine.open(new Uint8Array(await (await fetch("/tests/fixtures/private-regression.doc")).arrayBuffer()));
        try {
          let yellowPixels = 0;
          const numbered = [];
          const manual = [];
          for (let unitIndex = 0; unitIndex < doc.info.units.length; unitIndex++) {
            const objects = await doc.listObjects({ unitIndex });
            const frame = await doc.render({ unitIndex, scale: 1, includeTextFragments: true });
            try {
              const canvas = document.createElement("canvas");
              canvas.width = frame.bitmap.width; canvas.height = frame.bitmap.height;
              const ctx = canvas.getContext("2d"); ctx.drawImage(frame.bitmap, 0, 0);
              const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
              for (let i = 0; i < pixels.length; i += 4) if (pixels[i] > 240 && pixels[i + 1] > 240 && pixels[i + 2] < 30 && pixels[i + 3] > 240) yellowPixels++;
              for (const object of objects.filter(o => /^(?:\d[\d.]*|[a-f]\.|[-•])(?:\t|\s|[\u2e80-\uffff])/u.test(o.text ?? ""))) {
                if (/^\d[\d.]*\t/u.test(object.text)) continue;
                const fragments = frame.textFragments.filter(f => f.objectId === object.id && f.text.trim());
                manual.push({ unitIndex, id: object.id, label: object.text.match(/^[\d.a-f•-]+/u)[0], markerX: fragments[0]?.transform.e, bodyX: fragments[1]?.transform.e, tabbed: /^(?:[a-f]\.|[-•])\t/u.test(object.text) });
              }
              for (const object of objects.filter(o => /^\d[\d.]*\t/u.test(o.text ?? ""))) {
                const fragments = frame.textFragments.filter(f => f.objectId === object.id);
                const marker = fragments.find(f => /^\d/u.test(f.text));
                const body = fragments.find(f => f !== marker && f.text.trim());
                numbered.push({ unitIndex, id: object.id, label: object.text.match(/^[\d.]+/u)[0], markerX: marker?.transform.e, bodyX: body?.transform.e });
              }
            } finally { frame.bitmap.close(); }
          }
          return { yellowPixels, numbered, manual, units: doc.info.units.length };
        } finally { doc.close(); engine.close(); }
      });
      await page.waitForFunction(() => document.querySelector("#viewer").shadowRoot.querySelector('canvas'));
      for (let unitIndex = 0; unitIndex < checks.units; unitIndex++) {
        await page.evaluate(unitIndex => document.querySelector("#viewer").reveal({ kind: "unit", unitIndex }), unitIndex);
        await page.waitForFunction(unitIndex => document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${unitIndex}"]`)?.dataset.rendered === "true", unitIndex);
        const png = await page.evaluate(unitIndex => document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${unitIndex}"] canvas`).toDataURL().split(",")[1], unitIndex);
        await writeFile(`${output}/${name}-round2-${baseline ? "before" : "after"}-page${unitIndex + 1}.png`, Buffer.from(png, "base64"));
      }
      await writeFile(`${output}/${name}-${baseline ? "before" : "after"}.json`, JSON.stringify({ checks, errors }, null, 2));
      console.log(name, JSON.stringify({ ...checks, errors }));
      results.push({ name, checks, errors });
    } finally { await browser.close(); }
  }
} finally { server.kill(); }

const failures = [];
for (const { name, checks, errors } of results) {
  if (checks.yellowPixels <= 100) failures.push(`${name}: missing yellow highlighting`);
  if (checks.numbered.length !== 7) failures.push(`${name}: missing source numbered paragraphs`);
  for (const numbered of checks.numbered) {
    const expectedGap = numbered.label.split(".").length > 2 ? 28 : 24;
    if (Math.abs(numbered.markerX - 96.13333) >= 0.1 || Math.abs(numbered.bodyX - numbered.markerX - expectedGap) >= 0.1) {
      failures.push(`${name}: incorrect numbering indentation ${JSON.stringify(numbered)}`);
    }
  }
  if (checks.manual.length !== 45) failures.push(`${name}: missing manual, alphabetic or dash list coverage`);
  for (const manual of checks.manual) {
    if (!manual.tabbed && manual.label === "1.2." && manual.bodyX - manual.markerX > 32) failures.push(`${name}: manual Chinese numbering gap is stretched by justification ${JSON.stringify(manual)}`);
    if (manual.tabbed && Math.abs(manual.bodyX - manual.markerX - 24) > 0.1) failures.push(`${name}: alphabetic/dash tab disagrees with native Word ${JSON.stringify(manual)}`);
    if (Math.abs(manual.markerX - 96.16) > 0.1) failures.push(`${name}: manual/alphabetic/dash left edge disagrees with native Word ${JSON.stringify(manual)}`);
  }
  assert.deepEqual(errors, []);
}
assert.deepEqual(failures, []);
