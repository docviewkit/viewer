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
  for (const fixture of (process.env.DOCVIEWKIT_FIXTURE ?? "activex_commandbutton.pptx,activex_togglebutton.pptx").split(",")) {
    for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit").split(",")) {
      const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true });
      try {
        const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
        const errors = [];
        page.on("pageerror", error => errors.push(error.message));
        await page.goto(`${url}examples/viewer.html?fixture=${fixture}`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
        const result = await page.evaluate(async (fixture) => {
          const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
          const { SceneRenderer } = await import("/dist/render.js");
          const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
          const doc = core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()));
          const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
          try {
            const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 4 });
            const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
            const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
            const buttons = doc.scene.objects.filter(o => o.visual.kind === "image");
            const centers = buttons.map(o => {
              const b = o.bounds;
              const x = Math.round((b.x + b.width * .2) * 4), y = Math.round((b.y + b.height * .2) * 4);
              const w = Math.round(b.width * .6 * 4), h = Math.round(b.height * .6 * 4);
              const data = context.getImageData(x, y, w, h).data;
              let min = h, max = -1;
              for (let row = 0; row < h; row++) for (let col = 0; col < w; col++) {
                const i = (row * w + col) * 4;
                if (Math.abs(data[i] - data[i+2]) < 20 && data[i] < 220) { min = Math.min(min, row); max = Math.max(max, row); }
              }
              const colors = {};
              for (let row = 0; row < Math.min(16, h); row++) for (let col = 0; col < w; col++) {
                const i = (row * w + col) * 4;
                const key = Array.from(data.slice(i, i + 3)).join(",");
                colors[key] = (colors[key] ?? 0) + 1;
              }
              return { colors, offset: (y + (min + max) / 2) / 4 - (b.y + b.height / 2), min, max };
            });
            return { centers, diagnostics: frame.diagnostics };
          } finally { renderer.close(); doc.close(); }
        }, fixture);
        await page.screenshot({ path: `output/playwright/${fixture}-${name}.png` });
        console.log(fixture, name, JSON.stringify(result));
        assert.equal(result.centers.length, 3);
        if (fixture === "activex_togglebutton.pptx") {
          assert.ok(result.centers[0].colors["240,240,240"] > 100, "normal button keeps its gray face");
          const entries = Object.entries(result.centers[1].colors);
          const count = entries.reduce((n, [, count]) => n + count, 0);
          const green = entries.reduce((n, [rgb, count]) => n + Number(rgb.split(",")[1]) * count, 0) / count;
          assert.ok(entries.every(([rgb]) => { const [r, , b] = rgb.split(",").map(Number); return r === 255 && b === 255; }), "pressed face stays pink after scaling");
          assert.ok(Math.abs(green - 191.5) < 10, `equal pink/white dither averages to green 191.5, got ${green}`);
          assert.ok(result.centers[2].colors["255,255,255"] > 100, "transparent button keeps the slide background");
        }
        for (const center of result.centers) assert.ok(center.max >= center.min && Math.abs(center.offset) < 4, `caption must be vertically centered: ${JSON.stringify(center)}`);
        assert.ok(result.diagnostics.every(d => d.code === "IMAGE_FIDELITY_APPROXIMATE"));
        assert.deepEqual(errors, []);
      } finally { await browser.close(); }
    }
  }
} finally { server.kill(); }
