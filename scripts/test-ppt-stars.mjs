// Build JS and the legacy Office Wasm first. CURLZ_FONT_PATH optionally verifies exact typography.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium, firefox, webkit } from "playwright-core";

const url = process.env.VIEWER_URL ?? "http://127.0.0.1:8080/";
const font = process.env.CURLZ_FONT_PATH ? [...await readFile(process.env.CURLZ_FONT_PATH)] : null;
for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit").split(",")) {
  const browser = await { chromium, firefox, webkit }[name].launch({ timeout: 20000 });
  try {
    const page = await browser.newPage({ viewport: { width: 1100, height: 1300 } });
    const errors = [];
    page.on("pageerror", error => errors.push(error.message));
    page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
    await page.goto(`${url}examples/viewer.html?fixture=tdf168786.ppt`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
    if (font) await page.evaluate(async bytes => {
      const previous = document.querySelector("#viewer");
      const viewer = document.createElement(previous.localName);
      viewer.id = "viewer";
      viewer.config = { ...previous.config, engine: { ...previous.config.engine,
        fonts: [{ family: "Curlz MT", weight: 700, bytes: new Uint8Array(bytes) }] } };
      previous.replaceWith(viewer);
      previous.destroy();
      await viewer.open(new File([await (await fetch("/tests/fixtures/tdf168786.ppt")).arrayBuffer()], "tdf168786.ppt"));
    }, font);
    const result = await page.evaluate(() => {
      const canvas = document.querySelector("#viewer").shadowRoot.querySelector(".stage canvas");
      const context = canvas.getContext("2d");
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
      const region = (left, top, right, bottom, predicate) => {
        let count = 0, minY = 1, maxY = 0;
        for (let y = Math.ceil(top * canvas.height); y < bottom * canvas.height; y++) {
          for (let x = Math.ceil(left * canvas.width); x < right * canvas.width; x++) {
            const i = (y * canvas.width + x) * 4;
            if (predicate(pixels[i], pixels[i + 1], pixels[i + 2])) {
              count++; minY = Math.min(minY, y / canvas.height); maxY = Math.max(maxY, y / canvas.height);
            }
          }
        }
        return { count, minY, maxY };
      };
      return {
        whiteStarCorners: region(0.015, 0.035, 0.07, 0.08, (r, g, b) => r > 245 && g > 245 && b > 245).count,
        darkStar: region(0.02, 0.04, 0.15, 0.22, (r, g, b) => r < 70 && g < 90 && b < 130).count,
        title: region(0.1, 0.15, 0.34, 0.28, (r, g, b) => r > 100 && g < 65 && b < 140),
      };
    });
    await page.screenshot({ path: `/tmp/tdf168786-${name}${font ? "-curlz" : ""}.png` });
    assert.equal(result.whiteStarCorners, 0, "star bounding boxes must remain transparent");
    assert.ok(result.darkStar > 100, "dark offset star remains visible");
    if (font) {
      assert.ok(result.title.count > 100, "authored Curlz title is visible");
      assert.ok(result.title.maxY - result.title.minY < 0.06, "title stays on one line");
      assert.ok(result.title.minY > 0.21, "title clears the logo");
    }
    assert.deepEqual(errors, []);
    console.log(name, result, font ? "Curlz typography PASS" : "geometry PASS; exact font not supplied");
  } finally { await browser.close(); }
}
