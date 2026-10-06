// Run after building JS and the ODF Wasm. DOCVIEWKIT_BROWSERS selects engines.
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";

const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timeout")), 10000);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/[^\s]+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit").split(",")) {
    const browser = await { chromium, firefox, webkit }[name].launch({ timeout: 20000, headless: process.env.DOCVIEWKIT_HEADED !== "1" });
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 1100 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=empty.fodp`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.evaluate(() => { delete document.documentElement.dataset.ready; });
      await page.locator("#viewer-file-input").setInputFiles("tests/fixtures/two_columns.odg");
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.waitForTimeout(500);
      const ink = await page.locator("#viewer").evaluate(viewer => {
        const canvas = viewer.shadowRoot.querySelector(".stage canvas");
        const { data } = canvas.getContext("2d").getImageData(0, 0, canvas.width, canvas.height);
        // Authored A4 page, frame at (3.5, 2.75) cm, 13.75 cm wide; gap centered at 10.375 cm.
        return [[4, 9.5], [10.2, 10.55], [11, 16.8]].map(([left, right]) => {
          let count = 0;
          for (let y = Math.ceil(3 / 29.7 * canvas.height); y < 16 / 29.7 * canvas.height; y++) {
            for (let x = Math.ceil(left / 21 * canvas.width); x < right / 21 * canvas.width; x++) {
              const i = (y * canvas.width + x) * 4;
              if (data[i + 3] > 128 && data[i] < 100 && data[i + 1] < 100 && data[i + 2] < 100) count++;
            }
          }
          return count;
        });
      });
      assert.ok(ink[0] > 500 && ink[2] > 500, `${name}: both columns contain text: ${ink}`);
      assert.equal(ink[1], 0, `${name}: authored column gap is empty`);
      assert.deepEqual(errors, []);
      await page.screenshot({ path: `/tmp/odg-two-columns-${name}.png` });
      console.log(name, "two text columns, empty gap and clean console PASS", ink);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
