import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";

const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], {
  stdio: ["ignore", "pipe", "inherit"],
});
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10_000);
    createInterface({ input: server.stdout }).on("line", (line) => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit").split(",")) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1300, height: 1100 } });
      if (process.env.FILL_BASELINE_WASM) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.FILL_BASELINE_WASM, contentType: "application/wasm" }));
      const errors = [];
      page.on("pageerror", (error) => errors.push(error.message));
      page.on("console", (message) => { if (message.type() === "error") errors.push(message.text()); if (message.text().startsWith("fill-check:")) console.log(`${name}: ${message.text()}`); });
      await page.goto(`${url}examples/viewer.html?fixture=placeholder-priority.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.screenshot({ path: `/tmp/placeholder-priority-${name}${process.env.FILL_BASELINE_WASM ? "-baseline" : ""}.png` });
      await page.goto(`${url}examples/viewer.html?fixture=n820786.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        globalThis.fillCheckCore = core;
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/n820786.pptx")).arrayBuffer()));
        try {
          const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const ctx = canvas.getContext("2d");
          ctx.drawImage(frame.bitmap, 0, 0);
          const pixels = ctx.getImageData(800, 453, 12, 18).data;
          let white = 0, dark = 0;
          for (let i = 0; i < pixels.length; i += 4) {
            if (pixels[i] > 200) white++;
            if (pixels[i] < 100) dark++;
          }
          frame.bitmap.close();
          return { white, dark };
        } finally { doc.close(); }
      });
      const fills = await page.evaluate(async (browserName) => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = globalThis.fillCheckCore;
        const open = async name => core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${name}`)).arrayBuffer()));
        const check = (condition, message) => { if (!condition) throw new Error(message); };
        const render = async (doc, index = 0, objects = doc.scene.objects) => {
          const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
          const frame = await renderer.render(doc.scene.info.units[index], { unitIndex: index, scale: 1 });
          check(frame.diagnostics.length === 0, JSON.stringify(frame.diagnostics));
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0);
          frame.bitmap.close(); renderer.close();
          return context;
        };
        const pixel = (context, x, y) => [...context.getImageData(Math.floor(x), Math.floor(y), 1, 1).data];
        const distance = (a, b) => a.reduce((sum, value, index) => sum + Math.abs(value - b[index]), 0);
        const result = {};
        let authoredImagePaint;
        try {
          console.info("fill-check: placeholder font colors");
          const placeholders = await open("placeholder-priority.pptx");
          try {
            const context = await render(placeholders);
            const colors = [[255, 106, 82], [0, 176, 80]];
            const pixels = context.getImageData(0, 0, context.canvas.width, context.canvas.height).data;
            result.placeholderColors = colors.map(color => {
              let count = 0;
              for (let i = 0; i < pixels.length; i += 4) {
                if (color.every((value, channel) => Math.abs(pixels[i + channel] - value) < 8)) count++;
              }
              return count;
            });
            check(result.placeholderColors.every(count => count > 20), "aaa must render orange-red and bbb green");
          } finally { placeholders.close(); }
          console.info("fill-check: spreadsheet");
          const sheet = await open("basicspreadsheet-fills.xlsx");
          try {
            console.info("fill-check: spreadsheet opened");
            sheet.loadUnit(0);
            console.info("fill-check: spreadsheet loaded");
            // The complete workbook closes WebKit even with the pre-change Wasm.
            // Keep its authored fill cells as a focused cross-engine regression.
            const context = await render(sheet, 0, browserName === "webkit"
              ? sheet.scene.objects.filter(o => ["D1", "D2", "A8"].includes(o.source.address)) : sheet.scene.objects);
            console.info("fill-check: spreadsheet rendered");
            result.cellGradient = distance(pixel(context, 390, 2), pixel(context, 390, 16));
            check(result.cellGradient > 70, "D1 gradient must change from top to bottom");
            const cells = context.getImageData(360, 22, 64, 12).data;
            result.cellPattern = new Set(Array.from({ length: cells.length / 4 }, (_, i) => cells[i * 4])).size;
            check(result.cellPattern >= 2, "D2 must retain its grid foreground and background");
          } finally { sheet.close(); }
          console.info("fill-check: image tiles");
          const image = await open("tdf152070.pptx");
          try {
            const context = await render(image);
            const paint = image.scene.objects[0].visual.fill;
            authoredImagePaint = paint;
            const bitmap = await createImageBitmap(new Blob([paint.bytes], { type: paint.mediaType }));
            const tileWidth = bitmap.width * 96 / 300 * paint.mapping.scaleX;
            const tileHeight = bitmap.height * 96 / 300 * paint.mapping.scaleY;
            let error = 0;
            for (let i = 0; i < 20; i++) error += distance(pixel(context, 100 + i * 7, 100), pixel(context, 100 + i * 7 + tileWidth, 100));
            result.imageRepeatError = error / 20;
            result.imageTile = [tileWidth, tileHeight];
            check(result.imageRepeatError < 40, "Image must repeat using authored scale and intrinsic 300 DPI");
            const mirrored = structuredClone(image.scene.objects[0]);
            mirrored.visual.fill.mapping.flip = "flip-x";
            const mirror = await render(image, 0, [mirrored]);
            let mirrorError = 0;
            for (let i = 10; i < 30; i++) mirrorError += distance(
              pixel(mirror, paint.mapping.offsetX + i, 75),
              pixel(mirror, paint.mapping.offsetX + tileWidth * 2 - i - 1, 75));
            result.imageMirrorError = mirrorError / 20;
            check(result.imageMirrorError < 50, "Mirrored tiles must reverse alternate image columns");
            bitmap.close();
          } finally { image.close(); }
          console.info("fill-check: gradients");
          const gradients = await open("tdf114848.pptx");
          try {
            const text = gradients.scene.objects.find(o => o.text?.includes("Word Art Perspective"));
            const context = await render(gradients, 0, [text]);
            const data = context.getImageData(0, 0, context.canvas.width, context.canvas.height).data;
            const colors = new Set();
            for (let i = 0; i < data.length; i += 4) if (data[i] > data[i + 2] + 30 && data[i] > 80) colors.add(`${data[i]},${data[i + 1]},${data[i + 2]}`);
            result.textColors = colors.size;
            check(colors.size > 20, "Authored gradient text must contain multiple colored tones");
            gradients.loadUnit(4);
            const shape = await render(gradients, 4);
            result.shapeGradient = distance(pixel(shape, 480, 360), pixel(shape, 10, 10));
            check(result.shapeGradient > 80, "Shape gradient must preserve focus and boundary colors");
          } finally { gradients.close(); }
          console.info("fill-check: gradient mapping and outlines");
          const mapping = await open("tdf114848-fill-mapping.pptx");
          try {
            await render(mapping);
            const object = structuredClone(mapping.scene.objects.find(o => o.text?.includes("Word Art Perspective")));
            object.bounds = { x: 0, y: 0, width: 100, height: 60 };
            object.visual = { kind: "painted-shape", geometry: "rectangle", strokeWidth: 0, stroke: { kind: "none" },
              fill: { kind: "mapped-gradient", tile: { left: 0, top: 0, right: .5, bottom: 0 }, flip: "flip-x", rotateWithShape: true,
                paint: { kind: "linear-gradient", start: { x: 0, y: 0 }, end: { x: 100, y: 0 }, stops: [{offset:0,color:0xff0000ff},{offset:1,color:0x0000ffff}] } } };
            const gradient = await render(mapping, 0, [object]);
            result.gradientMirrorError = distance(pixel(gradient, 10, 30), pixel(gradient, 89, 30));
            check(result.gradientMirrorError < 15 && distance(pixel(gradient, 10, 30), pixel(gradient, 49, 30)) > 200,
              "Mapped gradient must retain its ramp and mirrored tiling");
            object.visual.stroke = { kind: "pattern", preset: "smCheck", foreground: 0x000000ff, background: 0xffffffff };
            object.visual.strokeWidth = 12; object.visual.fill = { kind: "none" };
            object.bounds = { x: 10, y: 10, width: 100, height: 60 };
            const outline = await render(mapping, 0, [object]);
            const strip = outline.getImageData(20, 10, 64, 4).data;
            let dark = 0, light = 0;
            for (let i = 0; i < strip.length; i += 4) { if (strip[i] < 30) dark++; if (strip[i] > 220) light++; }
            result.patternStroke = { dark, light };
            check(dark > 40 && light > 40, "Pattern outline must preserve both hatch colors");
            const originalText = mapping.scene.objects.find(o => o.text?.includes("Word Art Perspective"));
            let rich = originalText.visual; while (rich.visual) rich = rich.visual;
            object.bounds = { x: 10, y: 10, width: 600, height: 90 };
            object.visual = { kind: "rich-text", geometry: "rectangle", align: "start", lineHeight: 0,
              fill: { kind: "none" }, stroke: { kind: "none" }, strokeWidth: 0,
              runs: [{ ...rich.runs[0], text: "MMMMMMMM", fontSize: 60, color: 0x000000ff, paint: { kind: "solid", color: 0x000000ff } }] };
            const blackText = await render(mapping, 0, [object]);
            object.visual.runs[0].paint = { kind: "pattern", preset: "smCheck", foreground: 0x000000ff, background: 0xffffffff };
            const patternText = await render(mapping, 0, [object]);
            const blackCount = context => { const data = context.getImageData(10,10,600,90).data; let count=0; for(let i=0;i<data.length;i+=4) if(data[i]<40 && data[i+1]<40 && data[i+2]<40) count++; return count; };
            result.patternTextRatio = blackCount(patternText) / blackCount(blackText);
            check(result.patternTextRatio > .3 && result.patternTextRatio < .7, "Text glyphs must be filled by the checker brush");
            object.visual.runs[0].paint = authoredImagePaint;
            const imageText = await render(mapping, 0, [object]);
            const glyphPixels = imageText.getImageData(10,10,600,90).data;
            const textureColors = new Set();
            for(let i=0;i<glyphPixels.length;i+=4) if(Math.max(glyphPixels[i],glyphPixels[i+1],glyphPixels[i+2])-Math.min(glyphPixels[i],glyphPixels[i+1],glyphPixels[i+2])>25) textureColors.add(`${glyphPixels[i]},${glyphPixels[i+1]},${glyphPixels[i+2]}`);
            result.imageTextColors = textureColors.size;
            check(textureColors.size > 30, "Text glyphs must retain the image texture");
          } finally { mapping.close(); }
          return result;
        } finally { core.close(); delete globalThis.fillCheckCore; }
      }, name);
      console.log(`${name}: original Office fill pixels PASS ${JSON.stringify(fills)}`);
      await page.screenshot({ path: `/tmp/n820786-${name}.png` });
      assert.ok(result.white > 40 && result.dark > 20, `${name}: pct30 must show both background and dots: ${JSON.stringify(result)}`);
      assert.deepEqual(errors, []);
      console.log(`${name}: grouped pct30 pixels PASS ${JSON.stringify(result)}`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
