import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { mkdir, writeFile } from "node:fs/promises";
import { chromium, webkit } from "playwright-core";

const fixture = "ccc-planning-survey.docx";
const output = "output/playwright/docx-zero-style";
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
      if (process.env.DOCX_STYLE_BASELINE_WASM) {
        await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.DOCX_STYLE_BASELINE_WASM, contentType: "application/wasm" }));
      }
      await page.goto(`${url}examples/viewer.html?fixture=${fixture}`);
      await page.waitForFunction(() => document.documentElement.dataset.ready || document.documentElement.dataset.error);
      assert.equal(await page.evaluate(() => document.documentElement.dataset.error), undefined);
      const result = await page.evaluate(async fixture => {
        const { createOfficeEngine } = await import("/dist/engine.js");
        const engine = await createOfficeEngine({ execution: "inline", wasm: await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer() });
        const doc = await engine.open(new Uint8Array(await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()));
        try {
          const pages = [];
          for (let unitIndex = 0; unitIndex < doc.info.units.length; unitIndex++) {
            const objects = await doc.listObjects({ unitIndex });
            const frame = await doc.render({ unitIndex, scale: 1, includeTextFragments: true });
            try {
              pages.push({ objects: objects.length, text: objects.map(o => o.text ?? "").join(" "), fragments: frame.textFragments.length });
            } finally { frame.bitmap.close(); }
          }
          return { pages, diagnostics: doc.info.diagnostics };
        } finally { doc.close(); engine.close(); }
      }, fixture);
      assert.ok(result.pages.length > 1);
      assert.ok(result.pages.every(p => p.objects > 0 && p.fragments > 0));
      assert.match(result.pages.map(p => p.text).join(" "), /SURVEY: CCC DIGITAL KEY/);
      assert.match(result.pages.at(-1).text, /Question 13:/);
      assert.deepEqual(errors, []);
      await page.screenshot({ path: `${output}/${name}-after.png`, fullPage: true });
      await writeFile(`${output}/${name}-result.json`, JSON.stringify(result, null, 2));
      console.log(name, JSON.stringify({ pages: result.pages.length, fragments: result.pages.map(p => p.fragments), errors }));
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
