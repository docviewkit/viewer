import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { mkdir } from "node:fs/promises";
import { chromium, webkit } from "playwright-core";

// Native anchors and real-file provenance: tests/fixtures/list-indent-reference.json.
const fixtures = ["odp-paragraph-list-indent.odp", "odg-paragraph-list-indent.odg", "lo-bullet-label.rtf", "lo-legacy-bullet-label.rtf"];
const output = "output/playwright/list-indent";
await mkdir(output, { recursive: true });
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
const failures = [];
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10_000);
    server.once("error", reject);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      const baseline = process.env.LIST_INDENT_BASELINE === "1";
      if (baseline) {
        for (const pack of ["core", "odf"]) await page.route(`**/dist/office-viewer-${pack}.wasm`, route => route.fulfill({ path: `/tmp/ov-list-before-${pack}.wasm`, contentType: "application/wasm" }));
      }
      for (const fixture of fixtures) {
        await page.goto(`${url}examples/viewer.html?fixture=${fixture.endsWith(".odg") ? "empty.fodp" : fixture}`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
        if (fixture.endsWith(".odg")) {
          await page.evaluate(() => { delete document.documentElement.dataset.ready; });
          await page.locator("#viewer-file-input").setInputFiles(`tests/fixtures/${fixture}`);
          await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
        }
        const result = await page.evaluate(async fixture => {
          const { createOfficeEngine } = await import("/dist/engine.js");
          const engine = await createOfficeEngine({ execution: "inline", formatPack: () => import("/dist/extended-formats.js").then(m => m.extendedFormatPack) });
          const doc = await engine.open(new Uint8Array(await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()));
          try {
            const objects = await doc.listObjects({ unitIndex: 0 });
            const frame = await doc.render({ unitIndex: 0, scale: 1, includeTextFragments: true });
            try {
              if (fixture.endsWith(".odp") || fixture.endsWith(".odg")) {
                const object = objects.find(o => o.text?.includes("副标题"));
                const fragments = frame.textFragments.filter(f => f.objectId === object?.id);
                const first = fragments.find(f => f.text.includes("副标题"));
                const second = fragments.find(f => f.text.includes("第二级别标题"));
                return { firstX: first?.transform.e, secondX: second?.transform.e, offset: second?.transform.e - first?.transform.e };
              }
              const body = fixture.includes("legacy") ? "Bullet" : "AAA";
              const marker = fixture.includes("legacy") ? "•" : "·";
              const object = objects.find(o => o.text?.includes(body));
              const fragments = frame.textFragments.filter(f => f.objectId === object?.id);
              const label = fragments.find(f => f.text.trim() === marker);
              const text = fragments.find(f => f.text.includes(body));
              return { text: object?.text, markerX: label?.transform.e, bodyX: text?.transform.e, gap: text?.transform.e - label?.transform.e };
            } finally { frame.bitmap.close(); }
          } finally { doc.close(); engine.close(); }
        }, fixture);
        await page.waitForFunction(() => document.querySelector("#viewer").shadowRoot.querySelector("canvas")?.width > 0);
        await page.screenshot({ path: `${output}/${name}-${fixture}-${baseline ? "before" : "after"}.png` });
        console.log(name, fixture, JSON.stringify(result));
        try {
          if (fixture.endsWith(".rtf")) {
            assert.ok(Math.abs(result.markerX - 120) < 0.2, "native marker anchor");
            assert.ok(Math.abs(result.bodyX - 144) < 0.2, "native body anchor");
            assert.ok(Math.abs(result.gap - 24) < 0.2, "native marker/body gap");
          } else assert.ok(Math.abs(result.offset - 96) < 0.2, "native nested list offset");
        } catch (error) { failures.push(`${name} ${fixture}: ${error.message}`); }
      }
      assert.deepEqual(errors, [], `${name}: clean console`);
    } finally { await browser.close(); }
  }
  assert.deepEqual(failures, []);
} finally { server.kill(); }
