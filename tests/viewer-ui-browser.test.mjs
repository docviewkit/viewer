import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { before, after, test } from "node:test";
import { chromium, webkit } from "playwright-core";

let server, url;
before(async () => {
  server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
  url = await new Promise((resolve, reject) => {
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) resolve(match[0]);
    });
    server.once("exit", code => reject(new Error(`Viewer server exited: ${code}`)));
  });
});
after(() => server?.kill());

for (const engine of [chromium, webkit]) {
  test(`${engine.name()}: Viewer input respects IME and keeps visible keyboard focus`, { timeout: 90_000 }, async t => {
    const browser = await engine.launch();
    t.after(() => browser.close());
    const page = await browser.newPage();
    const errors = [];
    page.on("pageerror", error => errors.push(error.message));
    page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
    await page.goto(`${url}examples/viewer.html?fixture=visual-baseline.pptx`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
    const viewer = page.locator("docviewkit-viewer");
    await t.test("IME cancellation does not close document search", async () => {
      for (const composition of [{ isComposing: true }, { keyCode: 229 }]) {
        await viewer.locator('[data-action="search"]').click();
        const input = viewer.locator(".search-input");
        await input.fill("中文");
        await input.dispatchEvent("keydown", { key: "Escape", bubbles: true, ...composition });
        assert.equal(await input.isVisible(), true, `IME Escape must preserve search: ${JSON.stringify(composition)}`);
        assert.equal(await input.inputValue(), "中文");
        await input.press("Escape");
        assert.equal(await input.isVisible(), false);
        assert.equal(await viewer.locator('[data-action="search"]').evaluate(el => el.matches(":focus")), true);
      }
    });
    await t.test("IME confirmation does not commit a partial zoom", async () => {
      const zoom = viewer.locator(".zoom-input");
      for (const composition of [{ isComposing: true }, { keyCode: 229 }]) {
        const before = await viewer.evaluate(v => v.state.zoom);
        await zoom.fill("123");
        await zoom.dispatchEvent("keydown", { key: "Enter", bubbles: true, ...composition });
        assert.equal(await viewer.evaluate(v => v.state.zoom), before);
        await zoom.press("Escape");
      }
      await zoom.fill("125");
      await zoom.press("Enter");
      assert.equal(await viewer.evaluate(v => v.state.zoom), 1.25);
    });
    await t.test("search focus survives forced colors without relying on shadows", async () => {
      await page.emulateMedia({ forcedColors: "active" });
      const input = viewer.locator(".search-input");
      if (!await input.isVisible()) await viewer.locator('[data-action="search"]').click();
      await input.focus();
      const outline = await viewer.locator(".search").evaluate(el => {
        const style = getComputedStyle(el);
        return { width: parseFloat(style.outlineWidth), style: style.outlineStyle };
      });
      assert.ok(outline.width >= 2 && outline.style !== "none", JSON.stringify(outline));
    });
    assert.deepEqual(errors, []);
  });
}
