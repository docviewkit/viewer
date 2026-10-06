// Run with: node --test commercial/tests/home-layout.mjs
import test from "node:test";
import assert from "node:assert/strict";
import { once } from "node:events";
import { chromium } from "playwright-core";
import { createCommercialApp } from "../src/app.mjs";

test("localized home format cards fit desktop and mobile without browser errors", async () => {
  const app = createCommercialApp({ publicDir: new URL("../public", import.meta.url).pathname });
  let browser;
  try {
    app.server.listen(0, "127.0.0.1");
    await once(app.server, "listening");
    browser = await chromium.launch({ headless: true });
    for (const locale of ["en", "zh-cn"]) {
      for (const width of [1440, 375]) {
        const page = await browser.newPage({ viewport: { width, height: 1000 } });
        const errors = [];
        page.on("pageerror", (error) => errors.push(error.message));
        page.on("console", (message) => { if (message.type() === "error") errors.push(message.text()); });
        await page.goto(`http://127.0.0.1:${app.server.address().port}/${locale}/`, { waitUntil: "networkidle" });
        await page.locator("#formats").scrollIntoViewIfNeeded();
        assert.equal(await page.locator("#formats article").count(), 7);
        assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), `${locale} at ${width}px`);
        assert.deepEqual(errors, []);
        await page.close();
      }
    }
  } finally {
    await browser?.close();
    await app.close();
  }
});
