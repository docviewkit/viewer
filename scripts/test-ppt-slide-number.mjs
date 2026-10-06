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
      if (process.env.PPT_NUMBER_BASELINE_WASM) await page.route("**/dist/office-viewer-legacy-office.wasm", route => route.fulfill({ path: process.env.PPT_NUMBER_BASELINE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=tdf169705.ppt`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-legacy-office.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/tdf169705.ppt")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 4 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const footer = doc.scene.objects.find(object => object.unitIndex === 0 && (object.text === "1" || object.text === "*"));
          if (!footer) throw new Error("missing page number");
          let visual = footer.visual;
          while (visual.visual) visual = visual.visual;
          const bounds = footer.bounds;
          const pixels = context.getImageData(Math.floor(bounds.x * 4), Math.floor(bounds.y * 4), Math.ceil(bounds.width * 4), Math.ceil(bounds.height * 4));
          let minX = pixels.width, maxX = -1, minY = pixels.height, maxY = -1, ink = 0;
          for (let y = 0; y < pixels.height; y++) for (let x = 0; x < pixels.width; x++) {
            const i = (y * pixels.width + x) * 4;
            if (pixels.data[i] < 80 && pixels.data[i + 1] < 80 && pixels.data[i + 2] < 80) {
              minX = Math.min(minX, x); maxX = Math.max(maxX, x);
              minY = Math.min(minY, y); maxY = Math.max(maxY, y); ink++;
            }
          }
          return { text: footer.text, rendered: visual.runs.map(run => run.text).join(""), ink,
            width: maxX - minX + 1, height: maxY - minY + 1, diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); }
      });
      assert.equal(result.text, "1");
      assert.equal(result.rendered, "1");
      assert.ok(result.ink > 20 && result.height > result.width, `visible digit 1: ${JSON.stringify(result)}`);
      assert.deepEqual(result.diagnostics.map(item => item.code), ["IMAGE_FIDELITY_APPROXIMATE"]);
      assert.deepEqual(errors, []);
      console.log(`${name}: page number 1 is visible`, result);
      await page.screenshot({ path: `/tmp/ppt-number-${name}${process.env.PPT_NUMBER_BASELINE_WASM ? "-baseline" : ""}.png` });
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
