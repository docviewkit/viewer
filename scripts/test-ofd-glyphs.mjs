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
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,webkit").split(",")) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1300, height: 1100 } });
      const errors = [];
      page.on("pageerror", (error) => errors.push(error.message));
      page.on("console", (message) => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=ofdrw-z.ofd`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      const result = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-ofd.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const document = core.open(new Uint8Array(await (await fetch("/tests/fixtures/ofdrw-z.ofd")).arrayBuffer()));
        try {
          const objects = document.scene.objects.filter((object) => ["GlyphCount", "Glyphs", "ST_Array"].includes(object.text));
          const output = {};
          for (const scenario of ["fallback", "embedded", "browser", "direct"]) {
            // Identical real Canvas font isolates substitution policy from installed fonts.
            const fonts = { resolve: () => ({ family: "serif", source: scenario === "direct" ? "fallback" : scenario,
              ...(scenario === "browser" ? { spaceAdvanceEm: .5 } : {}) }) };
            const renderedObjects = scenario !== "direct" ? objects : objects.flatMap((object) =>
              object.visual.visual.children.map((child, index) => ({ ...object,
                numericId: object.numericId * 100 + index,
                bounds: child.bounds,
                text: child.visual.runs[0].text,
                visual: { ...object.visual, visual: child.visual },
              })));
            const renderer = new SceneRenderer(renderedObjects, DEFAULT_LIMITS, fonts);
            const frame = await renderer.render(document.scene.info.units[0], { unitIndex: 0, includeTextFragments: true });
            output[scenario] = frame.textFragments.map((fragment) => ({
              text: fragment.text,
              width: fragment.width * fragment.transform.a,
              x: fragment.transform.e,
            }));
            frame.bitmap.close();
          }
          output.expected = objects.flatMap((object) => object.visual.visual.children.map((child) => ({
            x: child.bounds.x,
            width: child.visual.runs[0].fontSize / 2,
          })));
          return output;
        } finally { document.close(); core.close(); }
      });
      assert.equal(result.fallback.map((glyph) => glyph.text).join(""), "GlyphCountGlyphsST_Array");
      for (const scenario of ["fallback", "browser", "direct"]) {
        assert.equal(result[scenario].length, result.expected.length);
        for (let index = 0; index < result.expected.length; index += 1) {
          const glyph = result[scenario][index];
          const expected = result.expected[index];
          assert.ok(Math.abs(glyph.width - expected.width) < .02, `${name} ${scenario} ${glyph.text}: ${glyph.width} != ${expected.width}`);
          assert.ok(Math.abs(glyph.x - expected.x) < .02, "authored DeltaX positions must survive");
        }
      }
      assert.ok(result.embedded[0].width > result.expected[0].width * 1.3, "embedded font metrics must remain untouched");
      assert.deepEqual(errors, []);
      console.log(`${name}: ${result.fallback.length} authored glyph positions and half-em widths passed; embedded metrics preserved`);
      if (process.env.DOCVIEWKIT_SCREENSHOT) await page.screenshot({ path: process.env.DOCVIEWKIT_SCREENSHOT });
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
