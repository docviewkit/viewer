import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";

const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10_000);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit").split(",")) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      if (process.env.PPT_RTL_BASELINE_WASM) await page.route("**/dist/office-viewer-legacy-office.wasm", route => route.fulfill({ path: process.env.PPT_RTL_BASELINE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=tdf77747.ppt`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-legacy-office.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/tdf77747.ppt")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const { data, width, height } = context.getImageData(0, 0, canvas.width, canvas.height);
          let minX = width, maxX = 0, ink = 0;
          for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
            const i = (y * width + x) * 4;
            if (data[i + 3] > 128 && data[i] < 100 && data[i + 1] < 100 && data[i + 2] < 100) {
              minX = Math.min(minX, x); maxX = Math.max(maxX, x); ink++;
            }
          }
          return { minX, maxX, ink, width, diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); }
      });
      assert.ok(result.ink > 100, "authored text must be visible");
      assert.ok(result.minX > result.width / 2, `text must be on the right: ${JSON.stringify(result)}`);
      assert.ok(result.maxX > 850 && result.maxX < 930, `text must end inside its right inset: ${JSON.stringify(result)}`);
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
      await page.screenshot({ path: `/tmp/tdf77747-${name}.png` });
      console.log(`${name}: Hebrew paragraphs rendered on the right; no render diagnostics or page errors`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
