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
      if (process.env.TABLE_BASELINE_WASM || process.env.TABLE_WASM) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.TABLE_BASELINE_WASM || process.env.TABLE_WASM, contentType: "application/wasm" }));
      await page.goto(`${url}examples/viewer.html?fixture=bnc910045.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true").catch(async error => { console.log(await page.locator("body").innerText(), errors); throw error; });
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const { createOfficeEngine } = await import("/dist/engine.js");
        const engine = await createOfficeEngine({ execution: "inline", wasm: await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer() });
        const original = await engine.open(new Uint8Array(await (await fetch("/tests/fixtures/bnc910045.pptx")).arrayBuffer()));
        const originalFrame = await original.render({ unitIndex: 0, scale: 1, includeTextFragments: true });
        const originalFragments = originalFrame.textFragments.filter(f=>f.text.trim());
        console.log("original table fragments", JSON.stringify(originalFragments.map(f=>({text:f.text,line:f.line,font:f.font}))));
        const originalLines = new Set(originalFragments.map(f=>f.line)).size;
        originalFrame.bitmap.close(); original.close(); engine.close();
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const failures = [], checks = [];
        const check = (condition, label) => { checks.push(label); if (!condition) failures.push(label); };
        check(originalLines === 2, "supplied file keeps its two authored text lines with metric compatible Calibri fallback");
        const unwrap = object => { let visual = object.visual; while (visual.visual) visual = visual.visual; return visual; };
        const open = async name => core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${name}`)).arrayBuffer()));
        const cascade = await open("table-style-cascade.pptx");
        try {
          const cells = cascade.scene.objects.filter(o => o.kind === "cell" || o.text?.startsWith("R"));
          check(cells.length === 16, "16 source cells survive");
          for (const [label, color] of [["R0C0",0xff0000ff],["R0C3",0x00ff00ff],["R3C0",0x0000ffff],["R3C3",0xffff00ff]]) {
            const cell = cells.find(o => o.text?.startsWith(label));
            check(unwrap(cell).fill.color === color, `${label} corner color`);
          }
          const corner = unwrap(cells.find(o => o.text?.startsWith("R0C0")));
          check(corner.runs[0].fontFamily === "Arial" && !corner.runs[0].bold && !corner.runs[0].italic, "corner font and explicit style off");
          const direct = corner.runs.find(run => run.text.includes("EXPLICIT"));
          check(direct.fontFamily === "Times New Roman" && direct.bold && direct.italic && direct.color === 0xff00ffff, "direct run formatting wins");
          const body = unwrap(cells.find(o => o.text === "R2C2"));
          check(body.runs[0].fontFamily === "Courier New" && body.runs[0].italic && !body.runs[0].bold, "whole table text style inheritance");
          check(unwrap(cells.find(o=>o.text === "R0C1")).fill.color === 0xeeeeeeff, "banding excludes header");
          check(unwrap(cells.find(o=>o.text === "R3C1")).fill.color === 0xeeeeeeff, "banding excludes footer");
          check(cascade.scene.objects.some(o=>!o.text && unwrap(o).fill?.color === 0xeeeeeeff), "direct table background fill");
          const lines = cascade.scene.objects.map(unwrap);
          check(lines.some(v => v.stroke?.color === 0x990099ff && v.strokeWidth === 4), "first row border override");
          check(lines.some(v => v.stroke?.color === 0x009999ff && v.strokeWidth === 3), "first column border override");
          check(cascade.scene.objects.some(o => o.visual.kind === "advanced-effect" && o.visual.outerShadow), "theme table shadow reference");
          const renderer = new SceneRenderer(cascade.scene.objects, DEFAULT_LIMITS);
          const frame = await renderer.render(cascade.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height), ctx = canvas.getContext("2d");
          ctx.drawImage(frame.bitmap,0,0);
          for (const [x,y,r,g,b] of [[80,100,255,0,0],[580,100,0,255,0],[80,400,0,0,255],[580,400,255,255,0]]) {
            const p = ctx.getImageData(x,y,1,1).data;
            check(p[0]===r && p[1]===g && p[2]===b, `corner pixels ${x},${y}`);
          }
          check(frame.diagnostics.length === 0, "cascade render diagnostics");
          frame.bitmap.close(); renderer.close();
        } finally { cascade.close(); }
        for (const fixture of ["table-style-image.pptx", "table-style-theme-image.pptx"]) {
        const texture = await open(fixture);
        try {
          const cell = texture.scene.objects.find(o => o.text?.startsWith("R0C0"));
          if (fixture === "table-style-theme-image.pptx") {
            check(texture.scene.objects.some(o=>o.visual.kind === "advanced-effect" && o.visual.glow), "direct table background glow");
            let effect = cell.visual;
            while (effect.visual && effect.kind !== "advanced-effect") effect = effect.visual;
            check(effect.threeD?.bevelTop?.width === 2 && effect.threeD.lightRig === "threePt", "cell3D bevel and lighting map to shared 3D rendering");
          }
          const fill = unwrap(cell).fill;
          check(fill.kind === "image" && fill.bytes.length > 100 && fill.tile && fill.mapping?.flip === "flip-x", "table style image relationship and tiling");
          const renderer = new SceneRenderer(texture.scene.objects, DEFAULT_LIMITS);
          const frame = await renderer.render(texture.scene.info.units[0], {unitIndex:0,scale:1});
          const canvas = new OffscreenCanvas(frame.bitmap.width,frame.bitmap.height), ctx=canvas.getContext("2d");ctx.drawImage(frame.bitmap,0,0);
          const pixels=ctx.getImageData(80,100,80,50).data, colors=new Set();
          for(let i=0;i<pixels.length;i+=4) colors.add(`${pixels[i]},${pixels[i+1]},${pixels[i+2]}`);
          check(colors.size>30,"real texture pixels");
          check(frame.diagnostics.length===0,"texture render diagnostics");frame.bitmap.close();renderer.close();
        } finally { texture.close(); }
        }
        const invalid = await open("table-style-image-invalid.pptx");
        try {
          check(invalid.scene.objects.filter(o=>o.text?.startsWith("R")).length === 16, "invalid table images preserve all text cells");
          check(invalid.scene.diagnostics.filter(d=>d.message.includes("table image fill")).length === 2 && invalid.scene.diagnostics.some(d=>d.code === "EXTERNAL_RESOURCE_BLOCKED"), "invalid crop, external and missing table images are diagnosed");
          check(invalid.scene.objects.some(o=>unwrap(o).fill?.color === 0xeeeeeeff), "invalid images retain the authored table background");
        } finally { invalid.close(); }
        let compoundBorders = 0;
        const all = await open("table-styles-built-in.pptx"), explicit = await open("table-styles-explicit.pptx");
        try {
          check(all.scene.info.units.length === 74, "74 built-in styles");
          for(let i=0;i<74;i++) {
            all.loadUnit(i); explicit.loadUnit(i);
            compoundBorders += all.scene.objects.filter(o=>o.unitIndex===i && o.visual.kind === "stroke-style" && o.visual.style.compound !== "single").length;
            const cells=all.scene.objects.filter(o=>o.unitIndex===i && o.text?.startsWith("R"));
            check(cells.length===16,`built-in ${i+1} retains cells`);
            const signature = doc => JSON.stringify(doc.scene.objects.filter(o=>o.unitIndex===i).map(o=>({text:o.text,bounds:o.bounds,visual:o.visual})));
            check(signature(all)===signature(explicit),`built-in ${i+1} matches its normative definition`);
          }
          check(compoundBorders > 0, "built-in compound borders retain shared stroke semantics");
        } finally { all.close(); explicit.close(); core.close(); }
        return { failures, checks: checks.length, originalLines };
      });
      console.log(`${name}: ${JSON.stringify(result)}`);
      await page.goto(`${url}examples/viewer.html?fixture=table-style-cascade.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.screenshot({ path: `output/playwright/table-styles-${name}${process.env.TABLE_BASELINE_WASM ? "-baseline" : ""}.png` });
      assert.deepEqual(result.failures, []);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
