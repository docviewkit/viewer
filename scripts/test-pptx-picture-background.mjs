import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import { createInterface } from "node:readline";
import { chromium, webkit } from "playwright-core";

// Extracted from the supplied VW presentation's slide 19; source XML/media are unchanged.
const fixture = "snowball-vw-picture-background.pptx";
const unitIndex = process.env.PICTURE_ORIGINAL ? 18 : 0;
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
mkdirSync("output/playwright", { recursive: true });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10_000);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const [name, engine] of Object.entries({ chromium, webkit })) {
    const browser = await engine.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 950 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      if (process.env.PICTURE_BASELINE_WASM) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.PICTURE_BASELINE_WASM, contentType: "application/wasm" }));
      if (process.env.PICTURE_ORIGINAL) await page.route(`**/tests/fixtures/${fixture}`, route => route.fulfill({ path: process.env.PICTURE_ORIGINAL }));
      await page.goto(`${url}examples/viewer.html?fixture=${fixture}`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.evaluate(unitIndex => document.querySelector("#viewer").reveal({ kind: "unit", unitIndex }), unitIndex);
      const result = await page.evaluate(async ({ fixture, unitIndex }) => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()));
        doc.loadUnit(unitIndex);
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[unitIndex], { unitIndex, scale: 0.5 });
          const canvas = document.createElement("canvas");
          canvas.width = frame.bitmap.width; canvas.height = frame.bitmap.height;
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const picture = doc.scene.objects.find(o => o.unitIndex === unitIndex && o.source.shapeId === 9);
          const pixel = (x, y) => [...context.getImageData(x, y, 1, 1).data];
          // Transparent top-left image margin must show its authored white fill.
          const background = pixel(Math.ceil(picture.bounds.x / 2) + 8, 8);
          const slideBackground = pixel(5, 5);
          return { background, slideBackground, diagnostics: frame.diagnostics, png: canvas.toDataURL() };
        } finally { renderer.close(); doc.close(); core.close(); }
      }, { fixture, unitIndex });
      const suffix = `${process.env.PICTURE_ORIGINAL ? "-original" : ""}${process.env.PICTURE_BASELINE_WASM ? "-before" : "-after"}`;
      writeFileSync(`output/playwright/picture-background-${name}${suffix}.png`, Buffer.from(result.png.split(",")[1], "base64"));
      delete result.png;
      await page.screenshot({ path: `output/playwright/picture-background-viewer-${name}${suffix}.png` });
      console.log(name, JSON.stringify(result));
      assert.deepEqual(result.background, [255, 255, 255, 255], "transparent image margin must retain the authored white background");
      assert.deepEqual(result.slideBackground, [0, 0, 0, 255], "slide background must stay black");
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
