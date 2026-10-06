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
      if (process.env.INDENT_BASELINE_WASM) await page.route("**/dist/office-viewer-legacy-office.wasm", route => route.fulfill({ path: process.env.INDENT_BASELINE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=ppt-indentation-bullets.ppt`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-legacy-office.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/ppt-indentation-bullets.ppt")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
          const lines = [];
          let current;
          for (let y = 0; y < canvas.height; y++) {
            let left = canvas.width;
            for (let x = 0; x < canvas.width; x++) {
              const i = (y * canvas.width + x) * 4;
              if (pixels[i] < 100 && pixels[i + 1] < 100 && pixels[i + 2] < 100) left = Math.min(left, x);
            }
            if (left < canvas.width) {
              if (!current) { current = { left, top: y }; lines.push(current); }
              current.left = Math.min(current.left, left);
            } else { current = undefined; }
          }
          return { lines, diagnostics: frame.diagnostics };

        } finally { renderer.close(); doc.close(); }
      });
      assert.equal(result.lines.length, 12);
      const origin = 88 + 9.6;
      const cm = 96 / 2.54;
      const expected = [1, 2.5, 2.5, 2.5, 1.5, 3, 3, 3, 2, 3, 3, 3];
      result.lines.forEach((line, index) => assert.ok(Math.abs(line.left - origin - expected[index] * cm) < 6,
        `line ${index}: ${line.left}, expected ${origin + expected[index] * cm}`));
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
      console.log(`${name}: all 3 bullet anchors and 9 continuation lines match authored indentation`);
      await page.screenshot({ path: `/tmp/ppt-indentation-${name}${process.env.INDENT_BASELINE_WASM ? "-baseline" : ""}.png` });
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
