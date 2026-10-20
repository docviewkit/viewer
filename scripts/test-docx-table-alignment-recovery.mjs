import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { mkdir, writeFile } from "node:fs/promises";
import { chromium, webkit } from "playwright-core";

const source = process.env.DOCX_RECOVERY_FILE;
assert.ok(source, "Set DOCX_RECOVERY_FILE to the supplied AN13013 DOCX");
const output = "output/playwright/nxp-open";
await mkdir(output, { recursive: true });
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
  const results = [];
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 1000 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.route("**/tests/fixtures/nxp-recovery.docx", route => route.fulfill({ path: source }));
      const baseline = process.env.DOCX_BASELINE_WASM;
      if (baseline) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: baseline, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=nxp-recovery.docx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true" || document.documentElement.dataset.error);
      const checks = await page.evaluate(async () => {
        if (document.documentElement.dataset.error) return { error: document.documentElement.dataset.error };
        const { createOfficeEngine } = await import("/dist/engine.js");
        const engine = await createOfficeEngine({ execution: "inline" });
        const doc = await engine.open(new Uint8Array(await (await fetch("/tests/fixtures/nxp-recovery.docx")).arrayBuffer()));
        try {
          const text = [];
          let imageObjects = 0;
          for (let unitIndex = 0; unitIndex < doc.info.units.length; unitIndex++) {
            const objects = await doc.listObjects({ unitIndex });
            text.push(...objects.map(object => object.text ?? ""));
            imageObjects += objects.filter(object => object.type === "image").length;
          }
          const combined = text.join(" ");
          const rendered = [];
          for (const unitIndex of [...new Set([0, Math.floor(doc.info.units.length / 2), doc.info.units.length - 1])]) {
            const frame = await doc.render({ unitIndex, scale: 1, includeTextFragments: true });
            try {
              rendered.push({ unitIndex, objects: frame.renderedObjectCount, fragments: frame.textFragments.length });
            } finally { frame.bitmap.close(); }
          }
          return { units: doc.info.units.length, imageObjects, rendered,
            content: ["AN13013", "EdgeLock", "Date of release: 14 September 2022"].map(value => combined.includes(value)),
            diagnostics: doc.diagnostics().filter(diagnostic => diagnostic.part === "word/styles.xml") };
        } finally { doc.close(); engine.close(); }
      });
      if (!checks.error) {
        for (const unitIndex of [0, checks.units - 1]) {
          await page.evaluate(unitIndex => document.querySelector("#viewer").reveal({ kind: "unit", unitIndex }), unitIndex);
          await page.waitForFunction(unitIndex => document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${unitIndex}"]`)?.dataset.rendered === "true", unitIndex);
          const png = await page.evaluate(unitIndex => document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${unitIndex}"] canvas`).toDataURL().split(",")[1], unitIndex);
          await writeFile(`${output}/${name}-page${unitIndex + 1}.png`, Buffer.from(png, "base64"));
        }
      }
      await writeFile(`${output}/${name}-${baseline ? "before" : "after"}.json`, JSON.stringify({ checks, errors }, null, 2));
      console.log(name, JSON.stringify({ checks, errors }));
      results.push({ name, checks, errors });
    } finally { await browser.close(); }
  }
  for (const { name, checks, errors } of results) {
    assert.equal(checks.error, undefined, `${name}: failed to open the real DOCX`);
    assert.ok(checks.units >= 40);
    assert.ok(checks.imageObjects > 100);
    assert.ok(checks.content.every(Boolean));
    assert.ok(checks.rendered.every(frame => frame.objects > 0 && frame.fragments > 0));
    assert.ok(checks.diagnostics.some(diagnostic => diagnostic.code === "FORMAT_INVALID" && diagnostic.severity === "warning"));
    assert.deepEqual(errors, []);
  }
} finally { server.kill(); }
