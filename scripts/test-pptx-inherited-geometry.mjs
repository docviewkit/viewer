import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";

const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.once("exit", code => reject(new Error(`Server exited: ${code}`)));
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) resolve(match[0]);
    });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit").split(",")) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 60_000 });
    let timer;
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 950 } });
      timer = setTimeout(() => { void page.close(); }, 60_000);
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=customshape-bitmapfill-srcrect.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.screenshot({ path: `/tmp/pptx-inherited-geometry-${name}.png` });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/customshape-bitmapfill-srcrect.pptx")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const ctx = canvas.getContext("2d"); ctx.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const picture = doc.scene.objects.find(object => object.name === "Content Placeholder 5");
          const { x, y, width, height } = picture.bounds;
          const pixel = (px, py) => [...ctx.getImageData(Math.floor(px), Math.floor(py), 1, 1).data];
          return { corners: [[x + 3, y + 3], [x + width - 4, y + 3], [x + 3, y + height - 4], [x + width - 4, y + height - 4]].map(([px, py]) => pixel(px, py)), interior: pixel(x + width / 2, y + 10), diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); core.close(); }
      });
      console.log(name, result);
      assert.ok(result.corners.every(pixel => pixel.slice(0, 3).every(value => value > 248)), "all four rounded corners must expose the white slide");
      assert.ok(result.interior.slice(0, 3).some(value => value < 230), "image content must remain visible");
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
    } finally { clearTimeout(timer); await browser.close(); }
  }
} finally { server.kill(); }
