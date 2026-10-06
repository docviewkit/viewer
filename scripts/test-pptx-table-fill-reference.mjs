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
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      if (process.env.TABLE_BASELINE_WASM || process.env.TABLE_WASM) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.TABLE_BASELINE_WASM || process.env.TABLE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=bnc910045.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/bnc910045.pptx")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const pixels = context.getImageData(75, 85, 270, 28).data;
          let blue = 0, white = 0;
          for (let i = 0; i < pixels.length; i += 4) {
            if (pixels[i] === 79 && pixels[i + 1] === 129 && pixels[i + 2] === 189) blue++;
            if (pixels[i] > 240 && pixels[i + 1] > 240 && pixels[i + 2] > 240) white++;
          }
          const cell = doc.scene.objects.find(o => o.text?.includes("Table with"));
          const lower = context.getImageData(75, 148, 270, 20).data;
          let lowerBlue = 0, lowerWhite = 0;
          for (let i = 0; i < lower.length; i += 4) {
            if (lower[i] === 79 && lower[i + 1] === 129 && lower[i + 2] === 189) lowerBlue++;
            if (lower[i] > 240 && lower[i + 1] > 240 && lower[i + 2] > 240) lowerWhite++;
          }
          return { blue, white, lowerBlue, lowerWhite, bounds: cell.bounds, text: doc.scene.objects.map(o => o.text ?? "").join(" "), diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); core.close(); }
      });
      await page.screenshot({ path: `output/playwright/bnc910045-${name}${process.env.TABLE_BASELINE_WASM ? "-baseline" : ""}.png` });
      console.log(`${name}: ${JSON.stringify(result)}`);
      assert.ok(result.blue > 2000, "first row must have its authored blue background");
      assert.ok(result.white > 50 && result.white < 4000, "white text must be visible on the blue background");
      assert.ok(result.bounds.height >= 96, "wrapped text must fit within its cell background");
      assert.ok(result.lowerBlue > 2000 && result.lowerWhite > 50 && result.lowerWhite < 2000,
        "last line must also remain visible against blue");
      assert.match(result.text, /Table with blue background in the first row\./);
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
