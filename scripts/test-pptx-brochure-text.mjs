// Original-document regression: the three page-7 cards explicitly inherit
// #404040 from lvl1pPr, despite a white fontRef and white higher list levels.
// node scripts/test-pptx-brochure-text.mjs <original English brochure.pptx> [output-prefix]
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { chromium, webkit } from "playwright-core";

assert.ok(process.argv[2], "Supply the unchanged original English brochure");
const output = resolve(process.argv[3] ?? "output/playwright/brochure-text/after");
await mkdir(dirname(output), { recursive: true });
const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
const results = [];
try {
  const url = await new Promise((resolveUrl, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timed out")), 10000);
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolveUrl(match[0]); }
    });
  });
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, reducedMotion: "reduce" });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.evaluate(() => { delete document.documentElement.dataset.ready; });
      await page.locator("#viewer-file-input").setInputFiles(resolve(process.argv[2]));
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true", undefined, { timeout: 120000 });
      await page.evaluate(() => document.querySelector("#viewer").reveal({ kind: "unit", unitIndex: 6 }));
      await page.waitForFunction(() => document.querySelector("#viewer").shadowRoot.querySelector('.continuous-page[data-unit-index="6"]')?.dataset.rendered === "true");
      const { png, ...pixels } = await page.evaluate(() => {
        const canvas = document.querySelector("#viewer").shadowRoot.querySelector('.continuous-page[data-unit-index="6"] canvas');
        const context = canvas.getContext("2d");
        // Interior text regions, excluding card edges, header, and shadows.
        const regions = [[110, 1140, 285, 220], [475, 1140, 230, 220], [785, 1140, 200, 220]];
        const grayPixels = regions.map(([x, y, w, h]) => {
          const data = context.getImageData(Math.round(x * canvas.width / 1080), Math.round(y * canvas.height / 1560), Math.round(w * canvas.width / 1080), Math.round(h * canvas.height / 1560)).data;
          let count = 0;
          for (let i = 0; i < data.length; i += 4) if (data[i] >= 60 && data[i] <= 100 && Math.abs(data[i] - data[i + 1]) <= 1 && Math.abs(data[i] - data[i + 2]) <= 1 && data[i + 3] === 255) count++;
          return count;
        });
        return { grayPixels, width: canvas.width, height: canvas.height, png: canvas.toDataURL("image/png").split(",")[1] };
      });
      await writeFile(`${output}-${name}.png`, Buffer.from(png, "base64"));
      results.push({ browser: name, ...pixels, errors });
      console.log(JSON.stringify(results.at(-1)));
    } finally { await browser.close(); }
  }
} finally {
  server.kill();
  await writeFile(`${output}.json`, JSON.stringify(results, null, 2) + "\n");
}
for (const result of results) {
  assert.deepEqual(result.errors, []);
  result.grayPixels.forEach((count, index) => assert.ok(count > 1000, `${result.browser} card ${index + 1}: expected visible authored dark gray text, got ${count} pixels`));
}
