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
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 950 } });
      const timer = setTimeout(() => { void page.close(); }, 60_000);
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      if (process.env.SOFT_EDGE_BASELINE) await page.route("**/dist/render.js", route => route.fulfill({ path: process.env.SOFT_EDGE_BASELINE, contentType: "text/javascript" }));
      await page.goto(`${url}examples/viewer.html?fixture=croppedTo0.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.screenshot({ path: `/tmp/croppedTo0-${name}${process.env.SOFT_EDGE_BASELINE ? "-before" : "-after"}.png` });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/croppedTo0.pptx")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const ctx = canvas.getContext("2d"); ctx.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const data = ctx.getImageData(770, 40, 160, 180).data;
          const edge = [...ctx.getImageData(742, 130, 1, 1).data];
          let blue = 0;
          for (let i = 0; i < data.length; i += 4) if (data[i + 2] > data[i] + 50 && data[i + 2] > 100) blue++;
          const picture = doc.scene.objects.find(object => object.name === "Picture 2");
          const control = new SceneRenderer([{ ...picture, visual: picture.visual.visual }], DEFAULT_LIMITS);
          const plain = await control.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          ctx.clearRect(0, 0, canvas.width, canvas.height); ctx.drawImage(plain.bitmap, 0, 0);
          plain.bitmap.close(); control.close();
          const reference = ctx.getImageData(770, 40, 160, 180).data;
          let controlBlue = 0;
          for (let i = 0; i < reference.length; i += 4) if (reference[i + 2] > reference[i] + 50 && reference[i + 2] > 100) controlBlue++;
          // The same negative-crop semantics apply to standalone images and shape fills.
          const paint = { ...picture.visual.visual.fill, cropLeft: -.5, cropTop: -.5, cropRight: -.5, cropBottom: -.5 };
          const padding = [];
          for (const visual of [paint, { ...picture.visual.visual, geometry: "rectangle", fill: paint }]) {
            const check = new SceneRenderer([{ ...picture, bounds: { x: 20, y: 20, width: 160, height: 160 }, visual }], DEFAULT_LIMITS);
            try {
              const frame = await check.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
              ctx.clearRect(0, 0, canvas.width, canvas.height); ctx.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
              padding.push({ outside: [...ctx.getImageData(30, 100, 1, 1).data], inside: [...ctx.getImageData(80, 80, 1, 1).data] });
            } finally { check.close(); }
          }
          return { blue, controlBlue, edge, padding, diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); core.close(); }
      });
      console.log(name, result);
      clearTimeout(timer);
      assert.ok(result.blue > 300, "soft edges must retain the image's blue interior lines");
      assert.ok(result.blue > result.controlBlue * .9, "interior lines must stay as clear as the unfeathered source");
      assert.ok(result.edge[1] > 20 && result.edge[1] < 230, "the authored soft edge must still fade into the white slide");
      for (const { outside, inside } of result.padding) {
        assert.ok(outside[3] === 0 || outside.slice(0, 3).every(value => value === 255), "negative crop retains empty padding");
        assert.ok(inside[0] > 200 && inside[1] < 30 && inside[3] > 200, "source content retains its mapped position and size");
      }
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
