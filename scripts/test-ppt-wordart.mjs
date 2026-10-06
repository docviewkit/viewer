// Run after building JS and the legacy Office Wasm. DOCVIEWKIT_BROWSERS selects engines.
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
    const browser = await { chromium, firefox, webkit }[name].launch({ timeout: 20000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 1100 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=tdf143315-WordartWithoutBullet.ppt`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      const ink = await page.locator("#viewer").evaluate(viewer => {
        const canvas = viewer.shadowRoot.querySelector(".stage canvas");
        const { data } = canvas.getContext("2d").getImageData(0, 0, canvas.width, canvas.height);
        let minX = canvas.width, maxX = -1, minY = canvas.height, maxY = -1, count = 0;
        for (let y = 0; y < canvas.height; y++) for (let x = 0; x < canvas.width; x++) {
          const i = (y * canvas.width + x) * 4;
          if (data[i + 3] > 128 && data[i] > 80 && data[i + 1] < 70 && data[i + 2] < 100) {
            minX = Math.min(minX, x); maxX = Math.max(maxX, x);
            minY = Math.min(minY, y); maxY = Math.max(maxY, y); count++;
          }
        }
        return { width: (maxX - minX + 1) / canvas.width * 960,
          height: (maxY - minY + 1) / canvas.height * 720, count };
      });
      await page.screenshot({ path: `/tmp/ppt-wordart-${name}.png` });
      assert.ok(ink.count > 100, `${name}: authored red glyphs are visible`);
      assert.ok(ink.width > 280 && ink.width < 360, `${name}: horizontal WordArt fills its 340px frame: ${JSON.stringify(ink)}`);
      assert.ok(ink.height > 25 && ink.height < 80, `${name}: one horizontal line: ${JSON.stringify(ink)}`);
      assert.deepEqual(errors, []);
      console.log(name, "horizontal WordArt and clean console PASS", ink);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
