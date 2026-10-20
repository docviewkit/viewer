import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { mkdir, writeFile } from "node:fs/promises";
import { chromium, webkit } from "playwright-core";

const fixture = "word-legacy-numbering-tabs.docx";
const output = "output/playwright/docx-legacy-numbering-tabs";
await mkdir(output, { recursive: true });
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10_000);
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
      const baseline = process.env.DOCX_NUMBERING_BASELINE_WASM;
      if (baseline) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: baseline, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=${fixture}`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      const checks = await page.evaluate(async fixture => {
        const { createOfficeEngine } = await import("/dist/engine.js");
        const engine = await createOfficeEngine({ execution: "inline", wasm: await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer() });
        const doc = await engine.open(new Uint8Array(await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()));
        try {
          const checks = [];
          for (let unitIndex = 0; unitIndex < doc.info.units.length; unitIndex++) {
            const objects = await doc.listObjects({ unitIndex });
            const bullets = objects.filter(o => o.text === "•");
            if (!bullets.length) continue;
            const frame = await doc.render({ unitIndex, scale: 1, includeTextFragments: true });
            try {
              for (const bullet of bullets) {
                const marker = frame.textFragments.find(f => f.objectId === bullet.id && f.text === "•");
                if (!marker) throw new Error(`Missing marker ${bullet.id}`);
                const body = objects.find(o => o.id.startsWith("docx:paragraph:") && Math.abs(o.bounds.y - bullet.bounds.y) < 0.01);
                if (!body) throw new Error(`Missing bullet body ${bullet.id}: ${JSON.stringify(objects.slice(0, 2))}`);
                const first = frame.textFragments.find(f => f.objectId === body.id && f.text.trim());
                if (!first) throw new Error(`Missing rendered body ${body.id}`);
                checks.push({ unitIndex, markerX: marker.transform.e, bodyX: first.transform.e, gap: first.transform.e - marker.transform.e });
              }
            } finally { frame.bitmap.close(); }
          }
          return checks;
        } finally { doc.close(); engine.close(); }
      }, fixture);
      const unitIndex = checks[0]?.unitIndex ?? 0;
      await page.evaluate(unitIndex => document.querySelector("#viewer").reveal({ kind: "unit", unitIndex }), unitIndex);
      await page.waitForFunction(unitIndex => document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${unitIndex}"]`)?.dataset.rendered === "true", unitIndex);
      const png = await page.evaluate(unitIndex => document.querySelector("#viewer").shadowRoot.querySelector(`.continuous-page[data-unit-index="${unitIndex}"] canvas`).toDataURL().split(",")[1], unitIndex);
      await writeFile(`${output}/${name}${baseline ? "-before" : "-after"}.png`, Buffer.from(png, "base64"));
      console.log(name, JSON.stringify({ bullets: checks.length, gaps: [...new Set(checks.map(c => +c.gap.toFixed(3)))], errors }));
      assert.ok(checks.length > 20);
      for (const check of checks) assert.ok(Math.abs(check.gap - 24) < 0.1, `${name}: ${JSON.stringify(check)}`);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
