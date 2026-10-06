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
      const page = await browser.newPage({ viewport: { width: 1400, height: 950 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=font-scale.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.screenshot({ path: `/tmp/font-scale-${name}.png` });
      const result = await page.evaluate(async () => {
        const font = new FontFace("Calibri", 'url(/dist/Carlito-Regular.ttf)');
        document.fonts.add(await font.load());
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch("/tests/fixtures/font-scale.pptx")).arrayBuffer()));
        const renderer = new SceneRenderer(doc.scene.objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, scale: 1 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const ctx = canvas.getContext("2d"); ctx.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
          const rows = [];
          for (let y = 75; y < 650; y++) {
            let body = 0, bullet = 0;
            for (let x = 65; x < 950; x++) {
              const p = ctx.getImageData(x, y, 1, 1).data;
              if (p[0] < 120 && p[1] < 120 && p[2] < 120) { if (x < 100) bullet++; else body++; }
            }
            if (body || bullet) rows.push({y, body, bullet});
          }
          const groups = [];
          for (const row of rows) {
            if (!groups.length || row.y > groups.at(-1).at(-1).y + 1) groups.push([]);
            groups.at(-1).push(row);
          }
          return { rows: groups.map(g => ({top:g[0].y, bottom:g.at(-1).y, bullet:g.some(r=>r.bullet>0), body:g.some(r=>r.body>0)})), diagnostics: frame.diagnostics };
        } finally { renderer.close(); doc.close(); core.close(); }
      });
      console.log(name, JSON.stringify(result));
      assert.deepEqual(result.rows.flatMap((row, index) => row.bullet ? [index] : []), [0, 3, 7]);
      assert.ok(result.rows.filter(row => row.bullet).every(row => row.body), "each bullet shares the first text line");
      assert.equal(result.rows.filter(row => row.body).length, 10, "PowerPoint: A/C use three lines, B uses four");
      assert.deepEqual(result.diagnostics, []);
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
