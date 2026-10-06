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
      if (process.env.ARC_BASELINE_WASM) await page.route("**/dist/office-viewer-legacy-office.wasm", route => route.fulfill({ path: process.env.ARC_BASELINE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=tdf122899_Arc_90_to_91_clockwise.ppt`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-legacy-office.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/tdf122899_Arc_90_to_91_clockwise.ppt")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 4 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const tipPixels = context.getImageData(460 * 4, 237 * 4, 17 * 4, 9 * 4).data;
          let tipInk = 0;
          for (let i = 0; i < tipPixels.length; i += 4) {
            if (tipPixels[i] > 80 && tipPixels[i] < 200 && tipPixels[i + 1] < 130 && tipPixels[i + 2] < 80) tipInk++;
          }
          return {
            tipInk,
            dot: [...context.getImageData(480.5 * 4, 246 * 4, 1, 1).data],
            gap: [...context.getImageData(479.5 * 4, 270 * 4, 1, 1).data],
            center: [...context.getImageData(500 * 4, 370 * 4, 1, 1).data],
            corner: [...context.getImageData(315 * 4, 253 * 4, 1, 1).data],
            objects: doc.scene.objects.length,
            diagnostics: frame.diagnostics,
          };
        } finally { renderer.close(); doc.close(); }
      });
      assert.equal(result.objects, 3);
      assert.ok(result.dot[0] > 80 && result.dot[0] < 200 && result.dot[1] < 100 && result.dot[2] < 80, `round endpoint: ${result.dot}`);
      assert.ok(result.tipInk > 30, `endpoint marks must extend above the arc: ${result.tipInk}`);
      assert.ok(result.gap.slice(0, 3).every(value => value > 240), `radial gap must remain white: ${result.gap}`);
      assert.ok(result.center[0] > 240 && result.center[1] > 150 && result.center[2] < 30, "authored yellow arc interior");
      assert.ok(result.corner.slice(0, 3).every(value => value > 240), "arc bounding-box corner remains outside the ellipse");
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
      console.log(`${name}: original PPT renders arc, endpoint marks and white radial gap`);
      await page.screenshot({ path: `/tmp/ppt-arc-${name}${process.env.ARC_BASELINE_WASM ? "-baseline" : ""}.png` });
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
