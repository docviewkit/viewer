// Run after npm run build. Real documents and Workers; delay only the competing render requests.
import assert from "node:assert/strict";
import { mkdir, readFile } from "node:fs/promises";
import { createServer } from "node:http";
import { resolve, extname, sep } from "node:path";
import { chromium, firefox, webkit } from "playwright-core";

const root = resolve(".");
const server = createServer(async (request, response) => {
  const path = resolve(root, `.${new URL(request.url, "http://localhost").pathname}`);
  try {
    if (!path.startsWith(root + sep)) throw new Error("Invalid path");
    const bytes = await readFile(path);
    response.setHeader("Content-Type", ({ ".js": "text/javascript", ".wasm": "application/wasm", ".html": "text/html" })[extname(path)] ?? "application/octet-stream");
    response.end(bytes);
  } catch { response.writeHead(404).end(); }
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
try {
  const url = `http://127.0.0.1:${server.address().port}/`;
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit").split(",")) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 20_000 });
    const timeout = setTimeout(() => void browser.close(), 60_000);
    try {
      for (const fixture of ["corpus-stacked-mix.pptx", "word-footnote-page-break.docx"]) {
        for (const mode of ["single", "continuous"]) {
          const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
          const errors = [];
          page.on("pageerror", (error) => errors.push(error.message));
          page.on("console", (message) => { if (message.type() === "error") errors.push(message.text()); });
          await page.route("**/examples/thumbnail-regression.html", (route) => route.fulfill({
            contentType: "text/html", body: '<style>html,body{margin:0;height:100%}docviewkit-viewer{height:100%}</style>',
          }));
          await page.goto(`${url}examples/thumbnail-regression.html`);
          const forward = await page.evaluate(async ({ mode, fixture }) => {
            await import("/dist/viewer.js");
            window.held = new Map();
            window.requests = [];
            window.hold = "thumbnails";
            const viewer = document.createElement("docviewkit-viewer");
            window.viewer = viewer;
            viewer.config = { initialPageMode: mode, navigation: "visible", engine: {
              workerFactory: (url) => {
                const worker = new Worker(url, { type: "module" });
                const post = worker.postMessage.bind(worker);
                worker.postMessage = (message, ...rest) => {
                  if (message.type === "cancel-render") window.held.delete(message.targetId);
                  if (message.type === "render") {
                    window.requests.push(message);
                    const thumbnail = message.supersedeKey?.startsWith("viewer-thumbnail:");
                    if ((window.hold === "thumbnails" && thumbnail)
                      || (window.hold === "pages" && !thumbnail)) {
                      window.held.set(message.id, () => post(message, ...rest));
                      return;
                    }
                  }
                  post(message, ...rest);
                };
                return worker;
              },
            } };
            document.body.append(viewer);
            const response = await fetch(`/tests/fixtures/${fixture}`);
            window.fixture = new File([await response.arrayBuffer()], fixture);
            const info = await viewer.open(window.fixture);
            const root = viewer.shadowRoot;
            const readyThumbnail = root.querySelector('.unit-button[data-unit-index="0"] canvas')?.dataset.rendered;
            root.querySelector('[data-action="navigation-thumbnails"]').click();
            await new Promise(requestAnimationFrame);
            const thumbnail = root.querySelector('.unit-button[data-unit-index="0"] canvas');
            const surface = mode === "single" ? root.querySelector(".surface")
              : root.querySelector('.continuous-page[data-unit-index="0"] canvas');
            const reference = document.createElement("canvas");
            reference.width = thumbnail.width;
            reference.height = thumbnail.height;
            const context = reference.getContext("2d");
            context.imageSmoothingQuality = "high";
            context.drawImage(surface, 0, 0, reference.width, reference.height);
            const expected = context.getImageData(0, 0, reference.width, reference.height).data;
            const actual = thumbnail.getContext("2d").getImageData(0, 0, reference.width, reference.height).data;
            let totalDifference = 0;
            let maxDifference = 0;
            for (let index = 0; index < expected.length; index++) {
              const difference = Math.abs(expected[index] - actual[index]);
              totalDifference += difference;
              maxDifference = Math.max(maxDifference, difference);
            }
            return { pages: info.units.length, mode: viewer.state.pageMode,
              readyThumbnail, rendered: thumbnail.dataset.rendered, equal: reference.toDataURL() === thumbnail.toDataURL(),
              meanDifference: totalDifference / expected.length, maxDifference,
              thumbnailWidth: thumbnail.width, surfaceWidth: surface.width,
              pendingCurrent: window.requests.filter((request) => request.supersedeKey === "viewer-thumbnail:0" && window.held.has(request.id)).length };
          }, { mode, fixture });
          assert.equal(forward.mode, mode);
          if (fixture.endsWith(".pptx")) assert.equal(forward.readyThumbnail, "true", "Thumbnail must be present when the main page becomes ready");
          assert.equal(forward.rendered, "true", `${name}/${mode}: main page ready but thumbnail missing`);
          assert.equal(forward.pendingCurrent, 0);
          assert.equal(forward.equal, true, `${name}/${mode}: thumbnail must contain the actual page pixels: ${JSON.stringify(forward)}`);
          assert.ok(forward.surfaceWidth > forward.thumbnailWidth);
          // Release remaining thumbnails, then retain them while delaying full-size page rendering.
          await page.evaluate(() => { window.hold = "none"; for (const release of window.held.values()) release(); window.held.clear(); });
          await page.waitForFunction(() => [...window.viewer.shadowRoot.querySelectorAll(".unit-button canvas")]
            .filter((canvas) => canvas.closest(".thumbnail").dataset.visible === "true")
            .every((canvas) => canvas.dataset.rendered === "true"));
          const reverse = await page.evaluate((mode) => {
            window.hold = "pages";
            const root = window.viewer.shadowRoot;
            // A mode switch keeps navigation thumbnails and creates an unrendered primary surface.
            root.querySelector('[data-action="page-mode"]').click();
            const nextMode = mode === "single" ? "continuous" : "single";
            const surface = nextMode === "single" ? root.querySelector(".surface")
              : root.querySelector('.continuous-page[data-unit-index="0"] canvas');
            const thumbnail = root.querySelector('.unit-button[data-unit-index="0"] canvas');
            return { nextMode, preview: surface.dataset.preview,
              equal: surface.toDataURL() === thumbnail.toDataURL(),
              cssWidth: surface.getBoundingClientRect().width, pixelWidth: surface.width,
              zoom: window.viewer.state.zoom, unitWidth: window.viewer.state.info.units[0].width,
              textCount: surface.parentElement.querySelectorAll(".text-layer-item").length };
          }, mode);
          assert.equal(reverse.preview, "true", `${name}/${mode}: missing thumbnail preview`);
          assert.equal(reverse.equal, true);
          assert.ok(reverse.cssWidth > reverse.pixelWidth * 2, "Preview should fill the page without allocating a full-size bitmap");
          assert.equal(reverse.textCount, 0);
          assert.ok(Math.abs(reverse.cssWidth / reverse.unitWidth - reverse.zoom) < .001);
          if (name === "chromium" && fixture.endsWith(".docx")) {
            await mkdir("output/playwright", { recursive: true });
            await page.screenshot({ path: `output/playwright/thumbnail-preview-${reverse.nextMode}.png` });
          }
          await page.evaluate(() => { window.hold = "none"; for (const release of window.held.values()) release(); window.held.clear(); });
          await page.waitForFunction((mode) => {
            const root = window.viewer.shadowRoot;
            const surface = mode === "single" ? root.querySelector(".surface")
              : root.querySelector('.continuous-page[data-unit-index="0"] canvas');
            return surface.dataset.preview === undefined && surface.width > 200;
          }, reverse.nextMode);
          if (name === "chromium" && fixture.endsWith(".docx")) {
            await page.screenshot({ path: `output/playwright/thumbnail-full-${reverse.nextMode}.png` });
          }
          if (reverse.nextMode === "single") {
            const rapid = await page.evaluate(async () => {
              window.hold = "pages";
              const first = window.viewer.reveal({ kind: "unit", unitIndex: 1 });
              const root = window.viewer.shadowRoot;
              const surface = root.querySelector(".surface");
              const preview = surface.dataset.preview;
              const equal = surface.toDataURL() === root.querySelector('.unit-button[data-unit-index="1"] canvas').toDataURL();
              const second = window.viewer.reveal({ kind: "unit", unitIndex: 0 });
              const currentPreview = surface.toDataURL() === root.querySelector('.unit-button[data-unit-index="0"] canvas').toDataURL();
              window.hold = "none";
              for (const release of window.held.values()) release();
              window.held.clear();
              await Promise.all([first, second]);
              return { preview, equal, currentPreview, finalPreview: surface.dataset.preview, unitIndex: window.viewer.state.currentUnitIndex };
            });
            assert.equal(rapid.preview, "true");
            assert.equal(rapid.equal, true);
            assert.equal(rapid.currentPreview, true);
            assert.equal(rapid.finalPreview, undefined);
            assert.equal(rapid.unitIndex, 0);
          }
          await page.evaluate(async () => { await window.viewer.close(); });
          assert.deepEqual(errors, []);
          console.log(JSON.stringify({ browser: name, fixture, mode, forward, reverse }));
          await page.close();
        }
      }
    } finally { clearTimeout(timeout); await browser.close(); }
  }
} finally { server.close(); server.closeAllConnections(); }
