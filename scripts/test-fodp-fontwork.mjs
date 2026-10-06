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
    const browser = await { chromium, firefox, webkit }[name].launch({ timeout: 20000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 1100 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=fit-to-size.fodp`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      const ink = await page.locator("#viewer").evaluate(viewer => {
        const canvas = viewer.shadowRoot.querySelector(".stage canvas");
        const { data } = canvas.getContext("2d").getImageData(0, 0, canvas.width, canvas.height);
        // The source page is 28 x 21 cm; all six source objects have explicit cm bounds.
        return [[1.4, 2, 12.4, 4.686], [14, 2.4, 12.599, 5.085],
          [1.201, 10.115, 12.599, 4.485], [14.2, 10.2, 12.399, 3.885],
          [2.6, 18, 10, 1.4], [15, 18, 10, 1.4]].map(([x, y, w, h]) => {
          const left = Math.floor(x / 28 * canvas.width), top = Math.floor(y / 21 * canvas.height);
          const width = Math.floor(w / 28 * canvas.width), height = Math.floor(h / 21 * canvas.height);
          let minX = width, maxX = -1, count = 0;
          for (let row = 0; row < height; row++) for (let column = 0; column < width; column++) {
            const offset = ((top + row) * canvas.width + left + column) * 4;
            if (data[offset + 3] > 128 && data[offset] < 100 && data[offset + 1] < 100 && data[offset + 2] < 100) {
              minX = Math.min(minX, column); maxX = Math.max(maxX, column); count++;
            }
          }
          return { width: (maxX - minX + 1) / width, center: (maxX + minX) / 2 / width, density: count / width / height };
        });
      });
      for (const index of [0, 3]) assert.ok(ink[index].width < .25, `${name}: unfitted text ${index + 1}`);
      for (const index of [1, 2]) {
        assert.ok(ink[index].width > .88, `${name}: fitted width ${index + 1}`);
        assert.ok(Math.abs(ink[index].center - .5) < .05, `${name}: fitted center ${index + 1}`);
      }
      for (const index of [4, 5]) {
        assert.ok(ink[index].width > .9, `${name}: Fontwork spans its source envelope`);
        assert.ok(ink[index].density > .025 && ink[index].density < .5, `${name}: glyphs, not a filled rectangle`);
      }
      assert.deepEqual(errors, []);
      await page.screenshot({ path: `/tmp/fodp-fontwork-${name}.png` });
      console.log(name, "six source text effects and clean console PASS", ink);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
