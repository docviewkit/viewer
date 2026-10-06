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
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: process.env.DOCVIEWKIT_HEADED !== "1", timeout: 30_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      if (process.env.LIST_BASELINE_WASM) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.LIST_BASELINE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=fill-color-list.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/fill-color-list.pptx")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
          const colors = [0xc0504d, 0x9bbb59, 0x8064a2];
          const regions = colors.map(color => {
            let count = 0, minX = Infinity, minY = Infinity, maxX = 0, maxY = 0;
            for (let y = 0; y < canvas.height; y++) for (let x = 0; x < canvas.width; x++) {
              const i = (y * canvas.width + x) * 4;
              if ((pixels[i] << 16 | pixels[i+1] << 8 | pixels[i+2]) !== color) continue;
              count++; minX = Math.min(minX, x); minY = Math.min(minY, y);
              maxX = Math.max(maxX, x); maxY = Math.max(maxY, y);
            }
            return { count, minX, minY, maxX, maxY };
          });
          const bodies = [240, 480, 720].map(x => {
            const i = (350 * canvas.width + x) * 4;
            return Array.from(pixels.slice(i, i + 3));
          });
          return { regions, bodies, diagnostics: frame.diagnostics };

        } finally { renderer.close(); doc.close(); core.close(); }
      });
      await page.screenshot({ path: `output/playwright/fill-color-list-${name}${process.env.LIST_BASELINE_WASM ? "-baseline" : ""}.png` });
      console.log(`${name}: ${JSON.stringify(result)}`);
      for (const [i, region] of result.regions.entries()) {
        assert.ok(region.count > 15000, `column ${i} retains its authored fill`);
        assert.ok(Math.abs(region.minX - [133, 374, 615][i]) < 4);
        assert.ok(Math.abs(region.minY - 103) < 4);
        assert.ok(Math.abs((region.maxY - region.minY) / (region.maxX - region.minX) - 0.4) < 0.03);
      }
      for (const [i, pixel] of result.bodies.entries()) {
        assert.ok(Math.min(...pixel) > 195 && Math.max(...pixel) < 245, `column ${i} has a pale body`);
      }
      assert.ok(result.bodies[0][0] > result.bodies[0][1]);
      assert.ok(result.bodies[1][1] > result.bodies[1][0]);
      assert.ok(result.bodies[2][2] > result.bodies[2][1]);
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
