// Run after npm run build. Optional argument: the original WNCDKP2 proposal PDF.
import assert from "node:assert/strict";
import { readFile, mkdir } from "node:fs/promises";
import { createServer } from "node:http";
import { resolve, extname, sep } from "node:path";
import { chromium, webkit } from "playwright-core";

const root = resolve(".");
const original = process.argv[2] ? await readFile(process.argv[2]) : undefined;
const server = createServer(async (request, response) => {
  const pathname = new URL(request.url, "http://localhost").pathname;
  const path = resolve(root, `.${pathname}`);
  try {
    if (!path.startsWith(root + sep)) throw new Error("Invalid path");
    const bytes = pathname === "/original.pdf" && original ? original : await readFile(path);
    response.setHeader("Content-Type", ({ ".js": "text/javascript", ".wasm": "application/wasm", ".html": "text/html" })[extname(path)] ?? "application/octet-stream");
    response.end(bytes);
  } catch { response.writeHead(404).end(); }
});
const cases = [
  ["01-single-wrapper", ["Child A", "Child B"], [0, 0]],
  ["02-multiple-roots", ["Root A", "Child A", "Root B", "Child B"], [0, 1, 0, 1]],
  ["03-chain", ["Middle", "Inner", "Child A", "Child B"], [0, 1, 2, 2]],
  ["04-different-destination", ["Child A", "Child B"], [0, 0]],
  ["05-single-leaf", ["Only bookmark"], [0]],
  ["06-single-child", ["Only child"], [0]],
  ["07-nested-branches", ["Section A", "Detail A", "Section B", "Detail B"], [0, 1, 0, 1]],
  ["08-default-theme-sibling", ["Default Theme", "Child A", "Other root"], [0, 1, 0]],
];
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
try {
  const url = `http://127.0.0.1:${server.address().port}`;
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,webkit").split(",")) {
    assert.ok(name === "chromium" || name === "webkit", "Only Chromium and WebKit are supported by this regression");
    const browser = await ({ chromium, webkit })[name].launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      await page.route("**/outline-regression.html", route => route.fulfill({ contentType: "text/html",
        body: '<link rel="icon" href="data:,"><style>html,body{margin:0;height:100%}docviewkit-viewer{height:100%}</style>' }));
      await page.goto(`${url}/outline-regression.html`);
      await page.evaluate(async () => {
        await import("/dist/viewer.js");
        const viewer = document.createElement("docviewkit-viewer");
        viewer.config = { navigation: "visible", engine: {
          formatPack: () => import("/dist/extended-formats.js").then(module => module.extendedFormatPack),
        } };
        document.body.append(viewer);
        window.viewer = viewer;
      });
      const open = path => page.evaluate(async path => {
        const info = await window.viewer.open(new File([await (await fetch(path)).arrayBuffer()], path.split("/").pop()));
        return info.outline;
      }, path);
      const rows = () => page.locator(".outline-button").evaluateAll(buttons => buttons.map(button => ({
        title: button.title, level: Number(button.style.getPropertyValue("--dv-outline-level")),
        ariaLevel: Number(button.getAttribute("aria-level")), unitIndex: Number(button.dataset.unitIndex),
        index: Number(button.dataset.outlineIndex), parent: Number(button.dataset.outlineParent),
        pos: Number(button.getAttribute("aria-posinset")), size: Number(button.getAttribute("aria-setsize")),
      })));
      const unchanged = async outline => assert.deepEqual(await page.evaluate(() => window.viewer.state.info.outline), outline,
        "Presentation must not mutate the authored SDK outline");
      if (original) {
        const outline = await open("/original.pdf");
        assert.equal(outline[0].title, "Default Theme");
        assert.equal(outline.length, 34);
        await mkdir("output/playwright", { recursive: true });
        await page.locator(".navigation").screenshot({ path: `output/playwright/pdf-outline-${name}.png` });
        const actual = await rows();
        assert.deepEqual(actual.map(row => row.title), outline.slice(1).map(item => item.title),
          `${name}: original PDF must expose the 33 slides directly`);
        assert.ok(actual.every(row => row.level === 0 && row.ariaLevel === 1 && row.parent === -1));
        await page.locator(".outline-button").nth(1).click();
        await page.waitForFunction(() => window.viewer.state.currentUnitIndex === 1);
        await unchanged(outline);
        console.log(`${name}: original PDF passed (34 authored bookmarks, 33 visible slide bookmarks)`);
      }
      for (const [fixture, titles, levels] of cases) {
        const path = `/tests/fixtures/pdf-outline/${fixture}.pdf`;
        const outline = await open(path);
        const actual = await rows();
        assert.deepEqual(actual.map(row => row.title), titles, `${name}/${fixture}: titles`);
        assert.deepEqual(actual.map(row => row.level), levels, `${name}/${fixture}: indentation`);
        assert.deepEqual(actual.map(row => row.ariaLevel), levels.map(level => level + 1));
        assert.deepEqual(actual.map(row => row.index), actual.map((_, index) => index));
        for (const row of actual) {
          const siblings = actual.filter(item => item.parent === row.parent);
          assert.equal(row.pos, siblings.findIndex(item => item.index === row.index) + 1);
          assert.equal(row.size, siblings.length);
        }
        for (const row of actual) {
          await page.locator(".outline-button").nth(row.index).click({ position: { x: 100, y: 16 } });
          await page.waitForFunction(unitIndex => window.viewer.state.currentUnitIndex === unitIndex, row.unitIndex);
          assert.equal(await page.locator(".outline-button").nth(row.index).getAttribute("aria-current"), "page");
        }
        const branch = actual.findIndex((row, index) => actual[index + 1]?.level > row.level);
        if (branch >= 0) {
          const parent = page.locator(".outline-button").nth(branch);
          const child = page.locator(".outline-button").nth(branch + 1);
          await parent.locator(".outline-toggle").click();
          await child.waitFor({ state: "hidden" });
          assert.equal(await child.isVisible(), false);
          await page.locator('[data-action="navigation-thumbnails"]').click();
          await page.locator('[data-action="navigation-outline"]').click();
          assert.equal(await parent.getAttribute("aria-expanded"), "false");
          await parent.focus();
          await page.keyboard.press("ArrowRight");
          await child.waitFor({ state: "visible" });
          assert.equal(await child.isVisible(), true);
          await page.keyboard.press("ArrowRight");
          assert.equal(await child.evaluate(button => button === button.getRootNode().activeElement), true);
          if (await child.getAttribute("aria-expanded") === "true") await page.keyboard.press("ArrowLeft");
          await page.keyboard.press("ArrowLeft");
          assert.equal(await parent.evaluate(button => button === button.getRootNode().activeElement), true);
        }
        await unchanged(outline);
        await open(path);
        assert.equal(await page.locator('.outline-button[hidden]').count(), 0, "Reopening resets collapsed state");
        console.log(`${name}: ${fixture} passed`);
      }
      // Real DOCX headings share the same viewer policy; parsing stays format-specific.
      for (const fixture of ["complex0.docx", "word-numbered-body.docx"]) {
        const outline = await open(`/tests/fixtures/${fixture}`);
        const rootLevel = Math.min(...outline.map(item => item.level));
        const roots = outline.filter(item => item.level === rootLevel);
        const expected = roots.length === 1 && outline.length > 1 ? outline.slice(1) : outline;
        assert.deepEqual((await rows()).map(row => row.title), expected.map(item => item.title));
        await unchanged(outline);
        console.log(`${name}: ${fixture} shared outline policy passed`);
      }
      assert.deepEqual(await open("/tests/fixtures/visual-baseline.pptx"), []);
      assert.equal(await page.locator(".outline-button").count(), 0);
      assert.equal(await page.locator(".unit-list").getAttribute("data-navigation-mode"), "thumbnails");
      assert.deepEqual(errors, [], `${name}: browser errors`);
    } finally { await browser.close(); }
  }
} finally { server.close(); server.closeAllConnections(); }
