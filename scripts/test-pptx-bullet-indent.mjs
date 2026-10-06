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
      if (process.env.BULLET_INDENT_BASELINE_WASM) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.BULLET_INDENT_BASELINE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=bullet-indent.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/bullet-indent.pptx")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const rows = [];
          for (let y = 160; y < 240; y++) {
            const xs = [];
            for (let x = 40; x < 160; x++) {
              const pixel = context.getImageData(x, y, 1, 1).data;
              if (pixel[0] < 100 && pixel[1] < 100 && pixel[2] < 100) xs.push(x);
            }
            if (xs.length) rows.push({ y, xs });
          }
          const groups = [];
          for (const row of rows) {
            if (!groups.length || row.y > groups.at(-1).at(-1).y + 1) groups.push([]);
            groups.at(-1).push(row);
          }
          const anchors = groups.map(group => {
            const xs = [...new Set(group.flatMap(row => row.xs))].sort((a, b) => a - b);
            const body = xs.find((x, index) => index > 0 && x - xs[index - 1] > 10);
            return { bullet: xs[0], body };
          });
          return { anchors, diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); core.close(); }
      });
      await page.screenshot({ path: `/tmp/bullet-indent-${name}${process.env.BULLET_INDENT_BASELINE_WASM ? "-baseline" : ""}.png` });
      assert.equal(result.anchors.length, 2);
      for (const anchor of result.anchors) {
        assert.ok(Math.abs(anchor.bullet - 48) < 4, JSON.stringify(result.anchors));
        assert.ok(Math.abs(anchor.body - (48 + 324000 / 9525)) < 4, JSON.stringify(result.anchors));
      }
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
      console.log(`${name}: both bullet and body pixel anchors match PowerPoint`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
