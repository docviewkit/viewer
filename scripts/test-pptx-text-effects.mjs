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
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 20_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      if (process.env.TEXT_EFFECTS_BASELINE) await page.route("**/dist/render.js", route =>
        route.fulfill({ path: process.env.TEXT_EFFECTS_BASELINE, contentType: "text/javascript" }));
      const started = Date.now();
      await page.goto(`${url}examples/viewer.html?fixture=Text_withEffects_100chars.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true", null, { timeout: 15_000 });
      const openMs = Date.now() - started;
      await page.screenshot({ path: `/tmp/text-effects-${name}.png` });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/Text_withEffects_100chars.pptx")).arrayBuffer()));
        doc.loadUnit(0);
        const { registerFonts, collectFontRequests } = await import("/dist/font.js");
        const fonts = await registerFonts([], undefined, undefined, undefined, [], [],
          { requestedFaces: collectFontRequests(doc.scene.objects) });
        const NativeCanvas = OffscreenCanvas;
        let pixels = 0;
        globalThis.OffscreenCanvas = class extends NativeCanvas {
          constructor(width, height) { super(width, height); pixels += width * height; }
        };
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS, fonts);
        try {
          const start = performance.now();
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0 });
          const renderMs = performance.now() - start;
          const canvas = new NativeCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d");
          context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const countBlue = (x, y, width, height) => {
            const data = context.getImageData(x, y, width, height).data;
            let count = 0;
            for (let i = 0; i < data.length; i += 4) {
              if (data[i + 2] > data[i] + 25 && data[i] < 220 && data[i + 3] > 100) count++;
            }
            return count;
          };
          const allocatedPixels = pixels;
          globalThis.OffscreenCanvas = NativeCanvas;
          const controls = structuredClone(doc.scene.objects);
          const strip = visual => {
            for (const effect of visual.effects ?? []) { delete effect.reflection; delete effect.glow; }
            if (visual.visual) strip(visual.visual);
          };
          controls.forEach(object => strip(object.visual));
          const control = new SceneRenderer(controls, DEFAULT_LIMITS, fonts);
          let plain;
          try { plain = await control.render(doc.scene.info.units[0], { unitIndex: 0 }); }
          finally { control.close(); }
          const reference = new NativeCanvas(canvas.width, canvas.height).getContext("2d");
          reference.drawImage(plain.bitmap, 0, 0); plain.bitmap.close();
          const original = context.getImageData(0, 0, canvas.width, canvas.height).data;
          const withoutEffects = reference.getImageData(0, 0, canvas.width, canvas.height).data;
          let title = 0, halo = 0, lastInkRow = 0, darkBlue = 0;
          for (let y = 0; y < 500; y++) for (let x = 90; x < 900; x++) {
            const i = (y * canvas.width + x) * 4;
            if (y < 150 && original[i] > original[i + 2] + 40) title++;
            if (y < 150 && withoutEffects[i + 1] > 250 && original[i + 1] < 245) halo++;
            if (y > 170 && withoutEffects[i + 2] > withoutEffects[i] + 25 && withoutEffects[i] < 200) lastInkRow = y;
            if (y > 170 && original[i + 2] > original[i] + 25 && original[i + 1] < 110) darkBlue++;
          }
          return { pixels: allocatedPixels, renderMs, title, halo, darkBlue,
            body: countBlue(90, 175, 800, 200), bullet: countBlue(48, 175, 40, 40),
            reflection: countBlue(90, lastInkRow + 1, 700, 18), diagnostics: frame.diagnostics };
        } finally {
          globalThis.OffscreenCanvas = NativeCanvas;
          renderer.close(); fonts.close(); doc.close(); core.close();
        }
      });
      console.log(name, { openMs, ...result });
      assert.ok(result.title > 1_000, "white title retains its orange glyph outline");
      assert.ok(result.halo > 100, "title glow extends outside the outline");
      assert.ok(result.bullet > 20, "bullet inherits the authored blue gradient");
      assert.ok(result.darkBlue > 1_000, "gradient reaches its dark stop within the glyph height");
      assert.ok(result.body > 10_000, "authored gradient text must be visible");
      assert.ok(result.reflection > 100, "reflection must start immediately below the glyphs, not below the line box");
      assert.ok(result.pixels < 10_000_000, "reflection fragments must not each allocate an entire slide");
      assert.deepEqual(result.diagnostics.filter(diagnostic => diagnostic.code !== "FONT_SUBSTITUTED"), []);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
