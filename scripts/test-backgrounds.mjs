import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";
const files = ["background-theme-gradient.pptx", "background-rgb-reference.pptx", "background-effects.pptx", "background-worksheet.xlsx", "background-linear.docx", "background-radial.docx", "background-stops.docx", "background-texture.docx", "background-picture.docx", "background-color.odt", "background-image.odt"];
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10_000);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,webkit").split(",")) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1250, height: 1000 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      for (const file of files) {
        await page.goto(`${url}examples/viewer.html?fixture=${file}`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
        const result = await page.evaluate(async ({ file, browserName }) => {
          const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
          const { SceneRenderer } = await import("/dist/render.js");
          const wasm = file.endsWith(".odt") ? "odf" : "core";
          const core = await Core.create(await (await fetch(`/dist/office-viewer-${wasm}.wasm`)).arrayBuffer(), DEFAULT_LIMITS);
          const doc = core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${file}`)).arrayBuffer()));
          doc.loadUnit(0);
          const unit = doc.scene.info.units[0];
          const render = async objects => {
            const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
            try {
              const frame = await renderer.render(unit, { unitIndex: 0, scale: 1 });
              if (frame.diagnostics.length) throw new Error(JSON.stringify(frame.diagnostics));
              const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
              const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
              return context;
            } finally { renderer.close(); }
          };
          try {
            // The complete pre-existing workbook closes WebKit during a second offscreen render.
            // Its real Viewer page is loaded above; isolate authored D1 plus its background here.
            const objects = browserName === "webkit" && file.endsWith(".xlsx")
              ? doc.scene.objects.filter(o => o.z < 0 || o.source.address === "D1") : doc.scene.objects;
            const context = await render(objects);
            const pixel = (ctx, x, y) => [...ctx.getImageData(Math.floor(x), Math.floor(y), 1, 1).data];
            const sample = [];
            for (let y = 8; y < Math.min(unit.height, 600); y += 17) for (let x = 8; x < Math.min(unit.width, 800); x += 17) sample.push(pixel(context, x, y).join(","));
            const result = { colors: new Set(sample).size, corner: pixel(context, 10, 10) };
            if (file === "background-effects.pptx") result.shadowEdge = pixel(context, 2, 2);
            if (file.endsWith(".xlsx")) {
              const background = doc.scene.objects.find(o => o.z < 0);
              if (!background) throw new Error("worksheet background missing");
              const without = await render(objects.filter(o => o !== background));
              const cell = doc.scene.objects.find(o => o.source.address === "D1");
              const x = cell.bounds.x + 4, y = cell.bounds.y + 4;
              result.cell = pixel(context, x, y); result.expectedCell = pixel(without, x, y);
            }
            return result;
          } finally { doc.close(); core.close(); }
        }, { file, browserName: name });
        if (file === "background-rgb-reference.pptx") assert.deepEqual(result.corner, [0, 204, 153, 255]);
        else if (file === "background-color.odt") assert.deepEqual(result.corner, [238, 232, 170, 255]);
        else assert.ok(result.colors > (file === "background-image.odt" ? 1 : 4), `${file}: authored background must have visible color variation: ${JSON.stringify(result)}`);
        if (result.shadowEdge) assert.ok(result.shadowEdge[1] < 160, "inner shadow must darken the background inside its edge");
        if (result.cell) assert.deepEqual(result.cell, result.expectedCell, "background must remain beneath opaque cell fills");
        assert.deepEqual(errors, []);
        await page.screenshot({ path: `/tmp/${file}-${name}.png` });
        console.log(`${name}: ${file} PASS ${JSON.stringify(result)}`);
      }
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
