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
      if (process.env.FILL_BASELINE_WASM) await page.route("**/dist/office-viewer-legacy-office.wasm", route => route.fulfill({ path: process.env.FILL_BASELINE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=FillPatterns.ppt`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-legacy-office.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/FillPatterns.ppt")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 3 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const shapes = doc.scene.objects.filter(object => object.bounds.width < 50 && object.bounds.height < 50);
          const tiles = shapes.map(object => {
            const { x, y } = object.bounds;
            const pixels = context.getImageData(Math.round(x * 3) + 6, Math.round(y * 3) + 6, 48, 48).data;
            let red = 0, white = 0;
            for (let i = 0; i < pixels.length; i += 4) {
              if (pixels[i] > 200 && pixels[i + 1] < 180 && pixels[i + 2] < 180) red++;
              if (pixels[i] > 240 && pixels[i + 1] > 180 && pixels[i + 2] > 180) white++;
            }
            return { red, white };
          });
          return { tiles, diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); }
      });
      assert.equal(result.tiles.length, 48);
      result.tiles.forEach((tile, index) => assert.ok(tile.red > 0 && tile.white > 0, `tile ${index}: ${JSON.stringify(tile)}`));
      assert.ok(result.tiles[0].red < result.tiles[0].white / 4, "first pattern must be sparse red dots on white");
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
      await page.screenshot({ path: `/tmp/FillPatterns-${name}.png` });
      console.log(`${name}: 48 authored pattern fills rendered in red and white; no render diagnostics or page errors`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
