import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, webkit, firefox } from "playwright-core";
const fixtures = process.env.CHART_FIXTURES?.split(",") ?? ["smartart-chevron.pptx", "smartart-cycle.pptx", "smartart-dir.pptx", "funnel-pp1.pptx", "color_funnel.xlsx", "testStockChart.docx", "tdf128207.docx", "SimpleHistogram.xlsx", "paretoLine.xlsx", "sunburst.xlsx", "tdf163727_histogram_underflow_overflow.xlsx", "pieOfPieChart.xlsx", "chart-text-can-overlap.xlsx"];
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("server timeout")), 10000);
    createInterface({ input: server.stdout }).on("line", line => { const m = line.match(/https?:\/\/\S+/u); if (m) { clearTimeout(timer); resolve(m[0]); } });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,webkit").split(",")) {
    const browser = await ({chromium, firefox, webkit})[name].launch({headless: true});
    try {
      const page = await browser.newPage({ viewport: {width: 1200, height: 900} });
      const errors = []; page.on("pageerror", e => errors.push(e.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      const wasmPath = process.env.CHART_WASM ?? process.env.CHART_BASELINE_WASM;
      if (wasmPath) await page.route("**/dist/office-viewer-core.wasm", r => r.fulfill({path: wasmPath, contentType: "application/wasm"}));
      for (const fixture of fixtures) {
        await page.goto(`${url}examples/viewer.html?fixture=${fixture}`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
        const result = await page.evaluate(async fixture => {
          const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
          const { SceneRenderer } = await import("/dist/render.js");
          const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
          const doc = core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()));
          doc.loadUnit(0);
          const objects = doc.scene.objects;
          const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
          const axisGlyphs = [];
          const contextPrototype = OffscreenCanvasRenderingContext2D.prototype;
          const fillText = contextPrototype.fillText;
          if (fixture === "chart-text-can-overlap.xlsx") contextPrototype.fillText = function(text, x, y, ...args) {
            const transform = this.getTransform();
            if (text.trim() && Math.abs(transform.b) > 0.01) {
              const metrics = this.measureText(text);
              const points = [-metrics.actualBoundingBoxLeft, metrics.actualBoundingBoxRight].flatMap(dx =>
                [-metrics.actualBoundingBoxAscent, metrics.actualBoundingBoxDescent].map(dy =>
                  new DOMPoint(x + dx, y + dy).matrixTransform(transform)));
              axisGlyphs.push({ text, top: Math.min(...points.map(p => p.y)), bottom: Math.max(...points.map(p => p.y)),
                left: Math.min(...points.map(p => p.x)), right: Math.max(...points.map(p => p.x)) });
            }
            return fillText.call(this, text, x, y, ...args);
          };
          try {
            const frame = await renderer.render(doc.scene.info.units[0], {unitIndex: 0, scale: 1});
            const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
            const context = canvas.getContext("2d"); context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
            const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
            let greenBorderPixels = 0;
            for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i+1] >= 150 && pixels[i+1] <= 190 && pixels[i+2] >= 60 && pixels[i+2] <= 100) greenBorderPixels++;
            return { axisGlyphs, greenBorderPixels, count: objects.length, objects: objects.filter(o => o.type !== "cell").map(o => ({id:o.id, parentId:o.parentId, kind:o.type, part:o.source?.part, text:o.text, bounds:o.bounds, visual: o.visual?.kind, geometry: typeof o.visual?.geometry === "string" ? o.visual.geometry : o.visual?.geometry?.kind})), diagnostics: frame.diagnostics };
          } finally { contextPrototype.fillText = fillText; renderer.close(); doc.close(); core.close(); }
        }, fixture);
        await page.screenshot({path: `output/playwright/completion-${fixture}-${name}${process.env.CHART_BASELINE_WASM ? "-baseline" : process.env.CHART_WASM ? "-candidate" : ""}.png`});
        assert.ok(result.count > 3, `${fixture} has materialized content`);
        const chartObjects = result.objects.filter(o => o.part?.includes("/charts/"));
        const labels = (fixture.startsWith("smartart-") ? result.objects : chartObjects).filter(o => o.text).map(o => o.text);
        if (fixture === "chart-text-can-overlap.xlsx") {
          const label = chartObjects.find(o => o.text === "Here is a long 1");
          const frame = result.objects.find(o => o.id === label.parentId).bounds;
          const axisY = Math.max(...chartObjects.filter(o => !o.text && o.bounds.width > frame.width / 2
            && o.bounds.height <= 1).map(o => o.bounds.y));
          assert.equal(result.axisGlyphs.map(g => g.text).join("").replace(/\s/gu, ""),
            Array.from({length: 7}, (_, i) => `Hereisalong${i + 1}`).join(""), "all seven authored labels are painted at 45 degrees");
          for (const glyph of result.axisGlyphs) {
            assert.ok(glyph.top >= axisY + 1, `${name}: ${glyph.text} crosses X axis (${glyph.top} <= ${axisY})`);
            assert.ok(glyph.bottom <= frame.y + frame.height, `${name}: ${glyph.text} exceeds chart bottom`);
            assert.ok(glyph.left >= frame.x && glyph.right <= frame.x + frame.width, `${name}: ${glyph.text} exceeds chart sides`);
          }
          console.log(`${name} rotated axis glyphs ${JSON.stringify({axisY,
            minimumGap: Math.min(...result.axisGlyphs.map(g => g.top - axisY)),
            bottomMargin: frame.y + frame.height - Math.max(...result.axisGlyphs.map(g => g.bottom)),
            rightMargin: frame.x + frame.width - Math.max(...result.axisGlyphs.map(g => g.right))})}`);
        }
        console.log(`${name} ${fixture} ${JSON.stringify({objects: result.count, charts: chartObjects.length, labels, greenBorderPixels: result.greenBorderPixels, diagnostics: result.diagnostics})}`);
        if (fixture === "smartart-chevron.pptx") {
          const panels = ["a", "b", "c"].map(t => result.objects.find(o => o.text === t));
          assert.ok(panels.every(p => p && Math.abs(p.bounds.y - panels[0].bounds.y) < 1), "chevron nodes share a row");
        }
        if (fixture === "smartart-dir.pptx") {
          assert.ok(result.objects.find(o => o.text === "second").bounds.x < result.objects.find(o => o.text === "first").bounds.x);
          assert.ok(labels.some(t => t.includes("• first")));
        }
        if (fixture === "funnel-pp1.pptx") assert.ok(labels.includes("Thing 4"), "funnel comes from actual values");
        if (fixture === "tdf163727_histogram_underflow_overflow.xlsx") assert.ok(labels.includes("≤ 11") && labels.includes("> 14"));
        if (fixture === "color_funnel.xlsx") assert.ok(result.greenBorderPixels > 100, "authored green funnel outlines reach rendered pixels");
        if (fixture === "testStockChart.docx") assert.ok(labels.includes("70") && labels.includes("160") && labels.includes("Volume"), "volume and price axes survive");
        if (fixture === "tdf128207.docx") assert.ok(labels.includes("0–2") && labels.includes("2–4") && labels.includes("4–6"), "surface uses value bands rather than series legend");
        if (fixture === "tdf128207.docx") assert.ok(chartObjects.filter(o => o.part.endsWith("chart2.xml") && o.geometry === "path").length >= 8, "surface bands are rendered");
        if (fixture === "SimpleHistogram.xlsx") assert.equal(labels[0], "Chart Title");
        if (fixture === "paretoLine.xlsx") assert.equal(labels[0], "ParetoLine");
        if (fixture === "SimpleHistogram.xlsx") assert.ok(labels.some(t => t.startsWith("[")), "numeric observations are binned");
        if (fixture === "pieOfPieChart.xlsx") assert.ok(chartObjects.filter(o => o.geometry === "path").length >= 9, "two pie plots and connectors");
        if (fixture === "smartart-cycle.pptx") {
          const a = result.objects.find(o => o.text === "a"), b = result.objects.find(o => o.text === "b"), e = result.objects.find(o => o.text === "e");
          assert.ok(b.bounds.x > a.bounds.x && e.bounds.x < a.bounds.x, "nodes follow the circular topology");
        }
        if (fixture === "paretoLine.xlsx") assert.ok(labels.includes("100%"), "cumulative percentage axis");
        if (fixture === "sunburst.xlsx") assert.ok(labels.includes("Best") && labels.includes("Worst"), "hierarchical labels");
        assert.deepEqual(result.diagnostics, []);
        assert.deepEqual(errors, []);
      }
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
