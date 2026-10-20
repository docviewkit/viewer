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
      if (process.env.NXP_OVERLAP_FONT) await page.route("**/nxp-stress-font.ttf", route => route.fulfill({ path: process.env.NXP_OVERLAP_FONT, contentType: "font/ttf" }));
      else await page.route("**/nxp-stress-font.ttf", route => route.fulfill({ status: 204 }));
      await page.goto(`${url}examples/viewer.html?fixture=nxp-recovery.docx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true" || document.documentElement.dataset.error);
      const checks = await page.evaluate(async () => {
        if (document.documentElement.dataset.error) return { error: document.documentElement.dataset.error };
        const { createOfficeEngine } = await import("/dist/engine.js");
        const fontResponse = await fetch("/nxp-stress-font.ttf");
        const engine = await createOfficeEngine({ execution: "inline",
          // Optional locally supplied wide face exercises fallback wrapping without bundling a font.
          ...(fontResponse.status === 200 ? { fonts: [{ family: "CIJBKI+Arial Bold", weight: 700, bytes: await fontResponse.arrayBuffer() }], fontPolicy: "deterministic" } : {}) });
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
          const firstObjects = await doc.listObjects({ unitIndex: 0 });
          const title = firstObjects.find(object => object.text?.startsWith("Get started with"));
          const revision = firstObjects.find(object => object.text?.startsWith("Rev."));
          const titleLayout = { title: title.bounds, revision: revision.bounds };
          const rendered = [];
          const titleRendering = [];
          let firstPagePng;
          for (const unitIndex of [...new Set([0, Math.floor(doc.info.units.length / 2), doc.info.units.length - 1])]) {
            const frame = await doc.render({ unitIndex, scale: 1, includeTextFragments: true });
            try {
              rendered.push({ unitIndex, declaredObjects: (await doc.listObjects({ unitIndex })).length, objects: frame.renderedObjectCount, fragments: frame.textFragments.length });
              if (unitIndex === 0) {
                const titleFragments = frame.textFragments.filter(fragment => fragment.objectId === title.id && fragment.text.trim());
                const revisionFragments = frame.textFragments.filter(fragment => fragment.objectId === revision.id && fragment.text.trim());
                titleRendering.push({ scale: 1, titleLines: new Set(titleFragments.map(fragment => fragment.line)).size,
                  titleBottom: Math.max(...titleFragments.map(fragment => fragment.transform.f + fragment.height)),
                  revisionTop: Math.min(...revisionFragments.map(fragment => fragment.transform.f)) });
                const canvas = document.createElement("canvas");
                canvas.width = frame.bitmap.width; canvas.height = frame.bitmap.height;
                canvas.getContext("2d").drawImage(frame.bitmap, 0, 0);
                firstPagePng = canvas.toDataURL().split(",")[1];
              }
            } finally { frame.bitmap.close(); }
          }
          const scaled = await doc.render({ unitIndex: 0, scale: 2, includeTextFragments: true });
          try {
            const titleFragments = scaled.textFragments.filter(fragment => fragment.objectId === title.id && fragment.text.trim());
            const revisionFragments = scaled.textFragments.filter(fragment => fragment.objectId === revision.id && fragment.text.trim());
            titleRendering.push({ scale: 2, titleLines: new Set(titleFragments.map(fragment => fragment.line)).size,
              titleBottom: Math.max(...titleFragments.map(fragment => fragment.transform.f + fragment.height)),
              revisionTop: Math.min(...revisionFragments.map(fragment => fragment.transform.f)) });
          } finally { scaled.bitmap.close(); }
          return { units: doc.info.units.length, imageObjects, rendered, titleLayout, titleRendering, firstPagePng,
            content: ["AN13013", "EdgeLock", "Date of release: 14 September 2022"].map(value => combined.includes(value)),
            diagnostics: doc.diagnostics().filter(diagnostic => diagnostic.part === "word/styles.xml") };
        } finally { doc.close(); engine.close(); }
      });
      if (checks.firstPagePng) {
        const mode = process.env.NXP_OVERLAP_FONT ? "wide-font" : "default";
        await writeFile(`${output}/${name}-${mode}-${baseline ? "before" : "after"}-page1.png`, Buffer.from(checks.firstPagePng, "base64"));
        delete checks.firstPagePng;
      }
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
    assert.ok(checks.titleLayout.revision.y >= checks.titleLayout.title.y + checks.titleLayout.title.height,
      `${name}: wrapped title overlaps the revision line: ${JSON.stringify(checks.titleLayout)}`);
    assert.ok(checks.titleRendering.every(sample => sample.titleBottom <= sample.revisionTop + 0.01),
      `${name}: actual painted title overlaps revision: ${JSON.stringify(checks.titleRendering)}`);
    if (process.env.NXP_OVERLAP_FONT) assert.ok(checks.titleRendering.every(sample => sample.titleLines === 2), "exercise the original wrapped-title failure path");
    assert.ok(checks.units >= 40);
    assert.ok(checks.imageObjects > 100);
    assert.ok(checks.content.every(Boolean));
    assert.ok(checks.rendered.every(frame => frame.objects > 0 && frame.fragments > 0));
    assert.ok(checks.rendered.every(frame => frame.objects === frame.declaredObjects),
      `${name}: frame avoidance must retain page content: ${JSON.stringify(checks.rendered)}`);
    assert.ok(checks.diagnostics.some(diagnostic => diagnostic.code === "FORMAT_INVALID" && diagnostic.severity === "warning"));
    assert.deepEqual(errors, []);
  }
} finally { server.kill(); }
