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
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit").split(",")) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1000, height: 900 } });
      const errors = [];
      page.on("pageerror", (error) => errors.push(error.message));
      page.on("console", (message) => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=ofdrw-zsbk.ofd`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      for (const width of [1000, 600]) {
        await page.setViewportSize({ width, height: 900 });
        await page.waitForFunction(() => {
          const root = document.querySelector("#viewer").shadowRoot;
          return root.querySelectorAll('.continuous-page[data-rendered="true"]').length === 2;
        });
        // Allow the resize observer and queued layout frame to settle.
        await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
        const sizes = await page.locator("#viewer").evaluate((viewer) =>
          [...viewer.shadowRoot.querySelectorAll(".continuous-page")].map((element) => {
            const { width, height } = element.getBoundingClientRect();
            return { width, height };
          }));
        assert.equal(sizes.length, 2);
        assert.ok(Math.abs(sizes[0].width / sizes[1].width - 297 / 210) < .002, JSON.stringify(sizes));
        assert.ok(Math.abs(sizes[0].height / sizes[1].height - 210 / 297) < .002, JSON.stringify(sizes));
        console.log(`${name} ${width}px: ${JSON.stringify(sizes)}`);
      }
      assert.deepEqual(errors, []);
    } finally {
      await browser.close();
    }
  }
} finally {
  server.kill();
}
