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
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,webkit").split(",")) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 20_000 });
    const timeout = setTimeout(() => void browser.close(), 120_000);
    try {
      for (const fixture of ["complex0.docx", "word-numbered-body.docx"]) {
        const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
        const errors = [];
        page.on("pageerror", error => errors.push(error.message));
        page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
        await page.goto(`${url}examples/viewer.html?fixture=${fixture}`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
        const authoredOutline = await page.locator("docviewkit-viewer").evaluate(viewer => viewer.state.info.outline);
        const rootLevel = Math.min(...authoredOutline.map(item => item.level));
        const rootCount = authoredOutline.filter(item => item.level === rootLevel).length;
        const outline = rootCount === 1 && authoredOutline.length > 1 ? authoredOutline.slice(1) : authoredOutline;
        assert.deepEqual(await page.locator(".outline-button").evaluateAll(items => items.map(item => item.title)),
          outline.map(item => item.title), "Navigation indexes address the displayed outline after root simplification");
        const branch = outline.findIndex((item, index) => outline[index + 1]?.level > item.level);
        assert.ok(branch >= 0, `${fixture} must exercise a real nested outline`);
        const item = index => page.locator(`.outline-button[data-outline-index="${index}"]`);
        const settleOutline = () => page.locator(".unit-list").evaluate(async list => {
          await Promise.allSettled(list.getAnimations({ subtree: true }).map(animation => animation.finished));
        });
        const visible = async () => {
          await settleOutline();
          return page.locator(".outline-button:visible").evaluateAll(items => items.map(item => Number(item.dataset.outlineIndex)));
        };
        const currentPage = () => page.locator("docviewkit-viewer").evaluate(viewer => viewer.state.currentUnitIndex);
        const all = outline.map((_, index) => index);
        const descendants = index => {
          const end = outline.findIndex((entry, i) => i > index && entry.level <= outline[index].level);
          return all.slice(index + 1, end < 0 ? undefined : end);
        };
        const motion = await page.locator("docviewkit-viewer").evaluate(async (viewer, index) => {
          const root = viewer.shadowRoot;
          const list = root.querySelector(".unit-list");
          const parent = root.querySelector(`.outline-button[data-outline-index="${index}"]`);
          const child = root.querySelector(`.outline-button[data-outline-index="${index + 1}"]`);
          const sample = () => ({ height: child.getBoundingClientRect().height,
            opacity: Number(getComputedStyle(child).opacity),
            arrow: getComputedStyle(parent.querySelector(".outline-toggle"), "::before").transform });
          const states = [];
          for (const direction of ["close", "open"]) {
            const before = sample();
            parent.querySelector(".outline-toggle").click();
            const animations = list.getAnimations({ subtree: true });
            const running = animations.some(animation => animation.playState === "running");
            for (const animation of animations) {
              animation.pause();
              animation.currentTime = animation.effect.getTiming().duration / 2;
            }
            await new Promise(requestAnimationFrame);
            const middle = sample();
            for (const animation of animations) animation.finish();
            await new Promise(requestAnimationFrame);
            states.push({ direction, running, before, middle, after: sample() });
          }
          return states;
        }, branch);
        for (const { direction, running, before, middle, after } of motion) {
          assert.ok(running, `${name}/${fixture}: outline ${direction} must animate`);
          assert.ok(middle.height > Math.min(before.height, after.height)
            && middle.height < Math.max(before.height, after.height), `${direction} must interpolate the actual row height: ${JSON.stringify({ before, middle, after })}`);
          assert.ok(middle.opacity > 0 && middle.opacity < 1, `${direction} must fade the content`);
          assert.notEqual(middle.arrow, before.arrow);
          assert.notEqual(middle.arrow, after.arrow);
        }
        assert.equal(await item(branch).getAttribute("aria-expanded"), "true", `${name}/${fixture}: parent has no disclosure control`);
        assert.deepEqual(await visible(), all);
        const pageBefore = await currentPage();
        await item(branch).locator(".outline-toggle").click();
        assert.equal(await item(branch + 1).evaluate(button => button.inert), true, "Collapsing descendants must stop accepting input immediately");
        assert.deepEqual(await visible(), all.filter(index => !descendants(branch).includes(index)));
        assert.equal(await currentPage(), pageBefore, "Collapsing must not navigate");
        await page.locator('[data-action="navigation-thumbnails"]').click();
        await page.locator('[data-action="navigation-outline"]').click();
        assert.equal(await item(branch).getAttribute("aria-expanded"), "false", "Mode changes must retain collapsed branches");
        await item(branch).focus();
        await page.keyboard.press("ArrowRight");
        assert.deepEqual(await visible(), all);
        await page.keyboard.press("ArrowRight");
        assert.equal(await item(branch + 1).evaluate(button => button === button.getRootNode().activeElement), true);
        const nested = descendants(branch).find(index => outline[index + 1]?.level > outline[index].level);
        if (nested !== undefined) {
          await item(nested).locator(".outline-toggle").click();
          await item(branch).locator(".outline-toggle").click();
          await item(branch).locator(".outline-toggle").click();
          assert.deepEqual(await visible(), all.filter(index => !descendants(nested).includes(index)), "Reopening a parent must preserve a child's collapse");
          await item(nested).focus();
          await page.keyboard.press("ArrowDown");
          const next = all.find(index => index > nested && !descendants(nested).includes(index));
          assert.equal(await item(next).evaluate(button => button === button.getRootNode().activeElement), true, "Keyboard navigation must skip hidden descendants");
          await item(nested).focus();
          await page.keyboard.press("ArrowRight");
        }
        const target = descendants(branch).find(index => outline[index].unitIndex !== pageBefore) ?? branch + 1;
        await item(target).click({ position: { x: 60, y: 16 } });
        assert.equal(await currentPage(), outline[target].unitIndex, "Title clicks must still navigate");
        assert.equal(await item(target).getAttribute("aria-current"), "page");
        await page.locator('[data-action="navigation-thumbnails"]').click();
        await page.locator('[data-action="navigation-outline"]').click();
        assert.equal(await item(target).getAttribute("aria-current"), "page", "Mode changes must retain the selected heading on the same page");
        await item(target).focus();
        await item(branch).locator(".outline-toggle").click();
        assert.equal(await item(branch).evaluate(button => button === button.getRootNode().activeElement), true, "Collapsing a focused descendant must restore visible focus");
        assert.equal(await item(branch).getAttribute("tabindex"), "0", "A hidden selection must keep its visible ancestor keyboard accessible");
        await item(branch).press("ArrowLeft");
        assert.equal(await currentPage(), outline[target].unitIndex, "Outline arrows must not change pages");
        await item(branch).press("ArrowRight");
        assert.equal(await item(target).getAttribute("aria-current"), "page", "Reopening must retain the selected heading");
        if (name === "chromium" && fixture === "complex0.docx") {
          await mkdir("output/playwright", { recursive: true });
          await settleOutline();
          await page.locator(".navigation").screenshot({ path: "output/playwright/outline-expanded.png" });
          await item(branch).locator(".outline-toggle").click();
          await settleOutline();
          await page.locator(".navigation").screenshot({ path: "output/playwright/outline-collapsed.png" });
        }
        await page.locator("docviewkit-viewer").evaluate(async (viewer, fixture) => {
          await viewer.open(new File([await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()], fixture));
        }, fixture);
        assert.deepEqual(await visible(), all, "Opening a document must reset collapsed state");
        await item(branch).evaluate(async button => {
          for (let index = 0; index < 6; index++) {
            button.querySelector(".outline-toggle").click();
            await new Promise(requestAnimationFrame);
          }
        });
        assert.deepEqual(await visible(), all, "Rapid reversals must finish fully expanded");
        await page.emulateMedia({ reducedMotion: "reduce" });
        for (const expanded of [false, true]) {
          const reduced = await item(branch).evaluate(async button => {
            button.querySelector(".outline-toggle").click();
            await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            const child = button.nextElementSibling;
            const style = getComputedStyle(child);
            return { height: child.getBoundingClientRect().height, inert: child.inert,
              opacity: Number(style.opacity), transition: style.transitionProperty,
              arrowTransition: getComputedStyle(button.querySelector(".outline-toggle"), "::before").transitionProperty };
          });
          assert.equal(reduced.transition, "none", "Reduced motion must disable row transitions");
          assert.equal(reduced.arrowTransition, "none", "Reduced motion must disable arrow transitions");
          assert.equal(reduced.height, expanded ? motion[0].before.height : 0, "Reduced motion must reach the final geometry immediately");
          assert.equal(reduced.opacity, expanded ? 1 : 0);
          assert.equal(reduced.inert, !expanded);
        }
        assert.deepEqual(errors, []);
        console.log(JSON.stringify({ browser: name, fixture, outlineCollapse: "passed", headings: outline.length }));
        await page.close();
      }
      for (const deviceScaleFactor of [1, 2]) {
        for (const fixture of ["digital-key-roadmap-btcs.pptx", "corpus-stacked-mix.pptx", "word-footnote-page-break.docx"]) {
          for (const mode of ["single", "continuous"]) {
            const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor });
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
              const thumbnail = root.querySelector('.unit-button[data-unit-index="0"] canvas');
              // Outline documents create thumbnails on mode switch. Their
              // IntersectionObserver callback may run after the next frame.
              for (let frame = 0; thumbnail.dataset.rendered !== "true" && frame < 120; frame++) {
                await new Promise(requestAnimationFrame);
              }
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
                thumbnailWidth: thumbnail.width, thumbnailHeight: thumbnail.height,
                cssWidth: thumbnail.getBoundingClientRect().width, cssHeight: thumbnail.getBoundingClientRect().height,
                surfaceWidth: surface.width,
                pendingCurrent: window.requests.filter((request) => request.supersedeKey === "viewer-thumbnail:0" && window.held.has(request.id)).length };
            }, { mode, fixture });
            assert.equal(forward.mode, mode);
            if (fixture.endsWith(".pptx")) assert.equal(forward.readyThumbnail, "true", "Thumbnail must be present when the main page becomes ready");
            assert.equal(forward.rendered, "true", `${name}/${mode}: main page ready but thumbnail missing`);
            assert.equal(forward.pendingCurrent, 0);
            assert.equal(forward.equal, true, `${name}/${mode}: thumbnail must contain the actual page pixels: ${JSON.stringify(forward)}`);
            assert.ok(forward.surfaceWidth > forward.thumbnailWidth);
            if (name === "chromium" && fixture === "digital-key-roadmap-btcs.pptx" && mode === "single") {
              await mkdir("output/playwright", { recursive: true });
              await page.locator(".navigation").screenshot({ path: `output/playwright/thumbnail-quality-dpr${deviceScaleFactor}.png` });
            }
            assert.ok(forward.thumbnailWidth >= forward.cssWidth * deviceScaleFactor * .98
              && forward.thumbnailHeight >= forward.cssHeight * deviceScaleFactor * .98,
            `${name}/${fixture}/${mode}/DPR${deviceScaleFactor}: thumbnail is enlarged beyond its raster resolution: ${JSON.stringify(forward)}`);
            // Release remaining thumbnails, then retain them while delaying full-size page rendering.
            await page.evaluate(() => { window.hold = "none"; for (const release of window.held.values()) release(); window.held.clear(); });
            await page.waitForFunction(() => [...window.viewer.shadowRoot.querySelectorAll(".unit-button canvas")]
              .every((canvas) => canvas.dataset.rendered === "true"));
            const quality = await page.evaluate(() => ({
              thumbnails: [...window.viewer.shadowRoot.querySelectorAll(".unit-button canvas")]
                .filter((canvas) => canvas.closest(".thumbnail").dataset.visible === "true")
                .map((canvas) => ({ width: canvas.width, height: canvas.height,
                  cssWidth: canvas.getBoundingClientRect().width, cssHeight: canvas.getBoundingClientRect().height })),
              requests: window.requests.filter((request) => request.supersedeKey?.startsWith("viewer-thumbnail:"))
                .map(({ request }) => {
                  const unit = window.viewer.state.info.units[request.unitIndex];
                  const canvas = window.viewer.shadowRoot.querySelector(`.unit-button[data-unit-index="${request.unitIndex}"] canvas`);
                  return { pixelRatio: request.pixelRatio,
                    width: Math.ceil(unit.width * request.scale * request.pixelRatio),
                    height: Math.ceil(unit.height * request.scale * request.pixelRatio),
                    thumbnailWidth: canvas.width, thumbnailHeight: canvas.height };
                }),
            }));
            for (const thumbnail of quality.thumbnails) {
              assert.ok(thumbnail.width >= thumbnail.cssWidth * deviceScaleFactor * .98
                && thumbnail.height >= thumbnail.cssHeight * deviceScaleFactor * .98,
              `${name}/${fixture}/${mode}/DPR${deviceScaleFactor}: prefetched thumbnail is undersized: ${JSON.stringify(thumbnail)}`);
            }
            for (const request of quality.requests) {
              assert.equal(request.pixelRatio, deviceScaleFactor);
              assert.ok(request.width >= request.thumbnailWidth && request.height >= request.thumbnailHeight,
                `Prefetched source must not be upscaled: ${JSON.stringify(request)}`);
            }
            if (name === "chromium" && fixture === "digital-key-roadmap-btcs.pptx" && mode === "single") {
              await page.locator(".navigation").screenshot({ path: `output/playwright/thumbnail-quality-dpr${deviceScaleFactor}.png` });
            }
            if (fixture === "digital-key-roadmap-btcs.pptx" && mode === "single") {
              const resizer = page.locator(".navigation-resizer");
              for (const [key, steps, width] of [["Shift+ArrowRight", 10, 400], ["Shift+ArrowLeft", 12, 120]]) {
                for (let step = 0; step < steps; step++) {
                  await resizer.press(key);
                  await page.waitForFunction(() => {
                    const root = window.viewer.shadowRoot;
                    return Math.abs(root.querySelector(".navigation").getBoundingClientRect().width
                      - parseFloat(root.querySelector(".shell").style.getPropertyValue("--dv-navigation-width"))) < .01;
                  });
                }
                assert.equal(await page.locator(".navigation").evaluate((navigation) => Math.round(navigation.getBoundingClientRect().width)), width);
                const thumbnail = await page.locator('.unit-button[data-unit-index="0"] canvas').evaluate((canvas) => ({
                  width: canvas.width, height: canvas.height,
                  cssWidth: canvas.getBoundingClientRect().width, cssHeight: canvas.getBoundingClientRect().height,
                }));
                assert.ok(thumbnail.width >= thumbnail.cssWidth * deviceScaleFactor * .98
                  && thumbnail.height >= thumbnail.cssHeight * deviceScaleFactor * .98,
                `Resizing navigation must retain thumbnail resolution: ${JSON.stringify(thumbnail)}`);
              }
              for (let step = 0; step < 3; step++) {
                await resizer.press(step === 2 ? "ArrowRight" : "Shift+ArrowRight");
                await page.waitForFunction(() => {
                  const root = window.viewer.shadowRoot;
                  return Math.abs(root.querySelector(".navigation").getBoundingClientRect().width
                    - parseFloat(root.querySelector(".shell").style.getPropertyValue("--dv-navigation-width"))) < .01;
                });
              }
              await page.waitForFunction(() => Math.abs(window.viewer.shadowRoot.querySelector(".navigation").getBoundingClientRect().width - 176) < 1);
            }
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
            console.log(JSON.stringify({ browser: name, fixture, mode, deviceScaleFactor, forward, quality, reverse }));
            await page.close();
          }
        }
      }
    } finally { clearTimeout(timeout); await browser.close(); }
  }
} finally { server.close(); server.closeAllConnections(); }
