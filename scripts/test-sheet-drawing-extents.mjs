import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { mkdir, readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { chromium, webkit } from "playwright-core";

const file = process.env.SHEET_DRAWING_FILE
  ?? fileURLToPath(new URL("../tests/fixtures/numbers-multiple-sheets-charts.numbers", import.meta.url));
assert.deepEqual(await readFile(file), await readFile(new URL(
  "../tests/fixtures/numbers-multiple-sheets-charts.numbers", import.meta.url,
)), "Verification must use the supplied original Numbers file");
const output = process.env.SHEET_DRAWING_OUTPUT ?? "output/playwright/sheet-drawing-extents";
await mkdir(output, { recursive: true });
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
let failed = false;
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10_000);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
    server.once("exit", code => reject(new Error(`Server exited: ${code}`)));
  });
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1440, height: 1200 } });
    try {
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      if (process.env.SHEET_DRAWING_BASELINE_WASM) {
        const bytes = await readFile(process.env.SHEET_DRAWING_BASELINE_WASM);
        await page.route("**/dist/office-viewer-pdf.wasm", route => route.fulfill({
          contentType: "application/wasm", body: bytes,
        }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector("#viewer").state.status === "ready");
      await page.locator("#viewer-file-input").setInputFiles(file);
      await page.waitForFunction(() => document.querySelector("#viewer").state.status === "ready"
        && document.querySelector("#viewer").shadowRoot.querySelector(".sheet-tab")?.textContent === "Budget");
      const viewer = page.locator("#viewer");
      await viewer.evaluate(v => v.reveal({ kind: "unit", unitIndex: 0 }));
      await page.screenshot({ path: `${output}/${name}.png` });
      const full = await viewer.evaluate(async v => {
        const { createOfficeEngine } = await import("/dist/engine.js");
        const { extendedFormatPack } = await import("/dist/extended-formats.js");
        const engine = await createOfficeEngine({ formatPack: async () => extendedFormatPack });
        const doc = await engine.open(await (await fetch("/tests/fixtures/numbers-multiple-sheets-charts.numbers")).arrayBuffer());
        try {
          const unit = doc.info.units[0];
          const objects = await doc.listObjects();
          const table = objects.find(object => object.unitIndex === 0 && object.type === "table");
          const title = objects.find(object => object.unitIndex === 0 && object.text === "Monthly Budget");
          const instructions = objects.find(object => object.unitIndex === 0 && object.text?.startsWith("HOW TO USE"));
          const nativeLayout = {
            // Read from Numbers' own accessibility geometry for the supplied file.
            tableY: table?.bounds.y, expectedTableY: 440 * 4 / 3,
            instructionsY: instructions?.bounds.y, expectedInstructionsY: 44 * 4 / 3,
            titleWidth: title?.bounds.width,
            charts: objects.filter(object => object.unitIndex === 0 && object.source.kind === "chart").length,
          };
          const canvas = v.shadowRoot.querySelector(".surface");
          const actual = { width: canvas.width / v.state.zoom, height: canvas.height / v.state.zoom };
          const direct = await doc.render({ unitIndex: 0, scale: v.state.zoom, pixelRatio: 1,
            viewport: { x: 0, y: 0, width: actual.width, height: actual.height },
            sheetSizes: { rows: [], columns: [] }, includeTextFragments: true });
          try {
            const reference = new OffscreenCanvas(direct.pixelWidth, direct.pixelHeight);
            reference.getContext("2d").drawImage(direct.bitmap, 0, 0);
            const expected = reference.getContext("2d").getImageData(0, 0, canvas.width, canvas.height).data;
            const pixels = canvas.getContext("2d").getImageData(0, 0, canvas.width, canvas.height).data;
            const gridHeight = unit.rows * unit.rowAxis.defaultSize + unit.rowAxis.spans.reduce(
              (delta, span) => delta + (span.end - span.start + 1) * (span.size - unit.rowAxis.defaultSize), 0);
            let maxDelta = 0, overflowInk = 0;
            for (let y = 1; y < canvas.height - 1; y++) for (let x = 1; x < canvas.width - 1; x++) {
              const offset = (y * canvas.width + x) * 4;
              for (let channel = 0; channel < 4; channel++) {
                maxDelta = Math.max(maxDelta, Math.abs(pixels[offset + channel] - expected[offset + channel]));
              }
              if (y / v.state.zoom > gridHeight && pixels[offset + 3] > 0
                && Math.min(pixels[offset], pixels[offset + 1], pixels[offset + 2]) < 220) overflowInk++;
            }
            return { actual, expected: { width: unit.width, height: unit.height }, maxDelta, overflowInk, nativeLayout,
              percentageFragments: direct.textFragments.filter(fragment => fragment.text.includes("%") || /^\d+$/.test(fragment.text)).map(fragment => fragment.text) };
          } finally { direct.bitmap.close(); }
        } finally { doc.close(); engine.close(); }
      });
      assert.ok(full.actual.width >= full.expected.width - 1 && full.actual.height >= full.expected.height - 1,
        `${name}: chart clipped to table: ${JSON.stringify(full)}`);
      assert.equal(full.nativeLayout.charts, 2, `${name}: missing native donut or column chart`);
      assert.ok(Math.abs(full.nativeLayout.tableY - full.nativeLayout.expectedTableY) < 1,
        `${name}: table must remain below both native charts`);
      assert.ok(Math.abs(full.nativeLayout.instructionsY - full.nativeLayout.expectedInstructionsY) < 1
        && full.nativeLayout.titleWidth > 670, `${name}: native title/instructions geometry lost`);
      for (const percentage of ["29%", "21%", "17%"])
        assert.ok(full.percentageFragments.includes(percentage), `${name}: wrapped doughnut label ${percentage}: ${JSON.stringify(full.percentageFragments)}`);
      assert.ok(full.overflowInk > 100, `${name}: missing chart pixels below the last table row`);
      assert.ok(full.maxDelta <= 1, `${name}: tiled chart differs from direct rendering: ${JSON.stringify(full)}`);
      // Exercise zoom, scroll, and cached tiles beyond the last cell in both axes.
      await page.setViewportSize({ width: 600, height: 400 });
      await viewer.locator(".zoom-input").fill("200");
      await viewer.locator(".zoom-input").press("Enter");
      await page.waitForFunction(expected => {
        const v = document.querySelector("#viewer"), wrap = v.shadowRoot.querySelector(".stage-wrap");
        return v.state.zoom === 2 && v.state.status === "ready"
          && parseFloat(wrap.style.width) >= expected.width * 2 + 47
          && parseFloat(wrap.style.height) >= expected.height * 2 + 23;
      }, full.expected);
      await viewer.evaluate(v => { const w = v.shadowRoot.querySelector(".workspace"); w.scrollTo(w.scrollWidth, w.scrollHeight); });
      await page.waitForFunction(expected => {
        const v = document.querySelector("#viewer"), root = v.shadowRoot;
        const stage = root.querySelector(".stage"), canvas = root.querySelector(".surface");
        return (parseFloat(stage.style.left) - 48 + parseFloat(canvas.style.width)) / v.state.zoom >= expected.width - 1
          && (parseFloat(stage.style.top) - 24 + parseFloat(canvas.style.height)) / v.state.zoom >= expected.height - 1;
      }, full.expected);
      await page.screenshot({ path: `${output}/${name}-scrolled.png` });
      await viewer.locator(".sheet-tab").nth(1).click();
      await page.waitForFunction(() => document.querySelector("#viewer").state.currentUnitIndex === 1);
      await viewer.locator(".sheet-tab").nth(0).click();
      await page.waitForFunction(() => document.querySelector("#viewer").state.currentUnitIndex === 0);
      assert.deepEqual(errors, []);
      console.log(`${name}: full chart, zoom/scroll and sheet switching passed ${JSON.stringify(full)}`);
    } catch (error) {
      failed = true;
      console.error(error);
      console.error(await page.locator("#viewer").evaluate(v => {
        const root = v.shadowRoot, w = root.querySelector(".workspace"), c = root.querySelector(".surface");
        return { state: v.state, stage: root.querySelector(".stage").style.cssText,
          canvas: c?.style.cssText, scroll: [w.scrollLeft, w.scrollTop, w.scrollWidth, w.scrollHeight, w.clientWidth, w.clientHeight] };
      }));
      await page.screenshot({ path: `${output}/${name}-failure.png` });
    }
    finally { await browser.close(); }
  }
} finally { server.kill(); }
if (failed) process.exitCode = 1;
