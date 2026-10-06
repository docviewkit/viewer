import assert from "node:assert/strict";
import { mkdirSync } from "node:fs";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";

mkdirSync("output/playwright", { recursive: true });
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
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30_000, ...(name === "firefox" && process.env.FIREFOX_EXECUTABLE ? { executablePath: process.env.FIREFOX_EXECUTABLE } : {}) });
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (["error", "warning"].includes(message.type())) errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=Ink2.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/Ink2.pptx")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 2 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const data = context.getImageData(0, 0, canvas.width, canvas.height).data;
          const pixels = { purple: 0, red: 0, yellow: 0 };
          for (let i = 0; i < data.length; i += 4) {
            const [r, g, b] = data.subarray(i, i + 3);
            if (r < 180 && g < 100 && b > r + 25) pixels.purple++;
            if (r > 230 && g < 100 && b < 100) pixels.red++;
            if (r > 245 && g > 245 && b > 145 && b < 190) pixels.yellow++;
          }
          const ink = doc.scene.objects.filter(o => o.source.part === "ppt/drawings/vmlDrawing1.vml");
          return { pixels, count: ink.length, diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); core.close(); }
      });
      await page.screenshot({ path: `output/playwright/Ink2-${name}.png` });
      console.log(name, JSON.stringify(result));
      assert.equal(result.count, 19);
      for (const [color, count] of Object.entries(result.pixels)) assert.ok(count > 1000, `${color} handwriting must be visibly painted`);
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
