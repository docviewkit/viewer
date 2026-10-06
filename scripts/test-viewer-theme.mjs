// Run after npm run build:js. DOCVIEWKIT_BROWSERS optionally selects engines.
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";

const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timeout")), 10_000);
    createInterface({ input: server.stdout }).on("line", (line) => {
      const match = line.match(/https?:\/\/[^\s]+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const [name, engine] of Object.entries({ chromium, firefox, webkit })) {
    if (process.env.DOCVIEWKIT_BROWSERS && !process.env.DOCVIEWKIT_BROWSERS.split(",").includes(name)) continue;
    console.log(`${name}: starting`);
    const browser = await engine.launch({ timeout: 20_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, colorScheme: "light", reducedMotion: "reduce" });
      const errors = [];
      page.on("pageerror", (error) => errors.push(error.message));
      page.on("console", (message) => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=visual-baseline.pptx&locale=zh-CN`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      assert.equal(await page.title(), "DocViewKit Viewer");
      const viewer = page.locator("docviewkit-viewer");
      const button = viewer.locator('[data-action="theme"]');
      assert.equal(await button.count(), 1, `${name}: missing theme control`);
      const theme = () => viewer.evaluate((v) => v.state.theme);
      const pixels = () => viewer.locator(".stage canvas").first().evaluate((c) => c.toDataURL());
      const before = await pixels();
      await page.emulateMedia({ colorScheme: "dark" });
      await page.waitForFunction(() => document.querySelector("#viewer").state.theme === "dark");
      await button.click(); // auto -> light
      assert.equal(await theme(), "light");
      await button.focus();
      await page.keyboard.press("Enter"); // light -> dark
      assert.equal(await theme(), "dark");
      assert.match(await button.getAttribute("aria-label"), /深色/);
      assert.equal(await viewer.evaluate((v) => getComputedStyle(v).getPropertyValue("--dv-accent").trim()), "#6aa7ff");
      await page.waitForFunction(() => getComputedStyle(document.querySelector("#viewer").shadowRoot.querySelector(".toolbar")).backgroundColor === "rgb(34, 37, 43)");
      assert.equal(await pixels(), before, "theme must preserve document pixels");
      await page.emulateMedia({ colorScheme: "light" });
      assert.equal(await theme(), "dark", "explicit theme overrides system");
      await page.screenshot({ path: `/tmp/viewer-theme-${name}-dark.png` });
      for (const width of [1440, 930, 720, 640, 601, 600, 541, 540, 480, 375, 320]) {
        await page.setViewportSize({ width, height: 900 });
        await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
        assert.equal(await button.isVisible(), true);
        const overlap = await viewer.evaluate((v) => {
          const toolbar = v.shadowRoot.querySelector(".toolbar").getBoundingClientRect();
          const controls = [...v.shadowRoot.querySelectorAll(".toolbar button, .toolbar .zoom-value")]
            .map((e) => e.getBoundingClientRect()).filter((r) => r.width > 0);
          return controls.some((r, i) => r.left < toolbar.left || r.right > toolbar.right
            || controls.slice(i + 1).some((s) => r.left < s.right && s.left < r.right));
        });
        assert.equal(overlap, false, `${name}: toolbar overlap at ${width}px`);
      }
      await page.setViewportSize({ width: 375, height: 812 });
      await page.waitForTimeout(500); // Let the resized document frame finish rendering before visual capture.
      await page.screenshot({ path: `/tmp/viewer-theme-${name}-mobile.png` });
      await button.click(); // dark -> auto
      assert.equal(await theme(), "light");
      await page.emulateMedia({ colorScheme: "dark" });
      await page.waitForFunction(() => document.querySelector("#viewer").state.theme === "dark");
      await viewer.evaluate((v) => { v.removeAttribute("theme"); v.config = { ...v.config, theme: "light", locale: "en" }; });
      assert.equal(await theme(), "light");
      assert.match(await button.getAttribute("aria-label"), /Light/);
      await button.click();
      assert.equal(await theme(), "dark");
      await viewer.evaluate((v) => v.close());
      assert.equal(await button.isEnabled(), true);
      await button.click();
      assert.equal(await theme(), "dark"); // auto, OS dark, even without a document
      assert.deepEqual(errors, []);
      console.log(`${name}: theme cycle, system/config, keyboard, unchanged document pixels, 11 widths, console PASS`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
