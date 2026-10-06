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
      if (process.env.BACKGROUND_BASELINE_WASM) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.BACKGROUND_BASELINE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=bnc904423.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/bnc904423.pptx")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const colors = ["Green", "Blue", "Red"].map(label => {
            const object = doc.scene.objects.find(o => o.text?.startsWith(label));
            return [...context.getImageData(Math.ceil(object.bounds.x + 10), Math.ceil(object.bounds.y + 10), 1, 1).data];
          });
          return { colors, diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); core.close(); }
      });
      await page.screenshot({ path: `/tmp/bnc904423-${name}${process.env.BACKGROUND_BASELINE_WASM ? "-baseline" : ""}.png` });
      assert.deepEqual(result.colors, [[0, 204, 153, 255], [51, 51, 204, 255], [255, 0, 0, 255]]);
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
      console.log(`${name}: green inherited, blue theme and red direct fills verified`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
