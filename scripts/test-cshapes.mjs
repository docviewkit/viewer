import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";

const name = process.env.DOCVIEWKIT_BROWSER ?? "chromium";
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], {
  stdio: ["ignore", "pipe", "inherit"],
});
let browser;
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10_000);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1250, height: 1000 } });
  const errors = [];
  page.on("pageerror", error => errors.push(error.message));
  page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
  await page.goto(`${url}examples/viewer.html?fixture=cshapes.pptx`);
  await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
  const checks = await page.evaluate(async () => {
    const { SceneRenderer } = await import("/dist/render.js");
    const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
    const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
    const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/cshapes.pptx")).arrayBuffer()));
    try {
      doc.loadUnit(0);
      const checks = [];
      const names = ["chartStar", "chartX", "chartPlus", "accentCallout1", "accentCallout2",
        "accentCallout3", "accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3"];
      for (const name of names) {
        const source = doc.scene.objects.find(o => o.name === name);
        let visual = source.visual;
        while (visual.visual) visual = visual.visual;
        if (visual.geometry.kind !== "layered-path") throw new Error(`${name} lost its authored paths`);
        // Exercise both scene paint adapters with the same real geometry and solid colors.
        for (const legacy of [false, true]) {
          const object = legacy ? { ...source, visual: { ...visual, kind: "shape",
            fill: visual.fill.color, stroke: visual.stroke.color } } : source;
          const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
          try {
            const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 2 });
            const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
            const context = canvas.getContext("2d");
            context.drawImage(frame.bitmap, 0, 0);
            frame.bitmap.close();
            if (frame.diagnostics.length) throw new Error(JSON.stringify(frame.diagnostics));
            const { x, y, width, height } = object.bounds;
            const pixel = [...context.getImageData(Math.round((x + width / 2) * 2), Math.round((y + height / 2) * 2), 1, 1).data];
            checks.push({ name, legacy, pixel });
          } finally { renderer.close(); }
        }
      }
      return checks;
    } finally { doc.close(); core.close(); }
  });
  await page.screenshot({ path: `/tmp/cshapes-${name}.png` });
  assert.deepEqual(errors, []);
  for (const check of checks) {
    assert.deepEqual(check.pixel, check.name.startsWith("chart") ? [160, 160, 96, 255] : [255, 255, 127, 255],
      `${check.name} (${check.legacy ? "legacy" : "painted"}) must preserve authored fill and visible strokes`);
  }
  console.log(`${name}: ${checks.length} cshapes pixel checks passed; clean console`);
} finally { await browser?.close(); server.kill(); }
