// Capture 3928 pie charts after pie-offset / data-label / bottom-legend fixes.
import { mkdir } from "node:fs/promises";
import { chromium } from "playwright-core";

const outDir = "output/playwright";
await mkdir(outDir, { recursive: true });
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1400, height: 1100 } });
await page.goto(
  "http://127.0.0.1:4173/examples/viewer.html?fixture=oasis-3928-pie-explode.ods",
  { waitUntil: "networkidle" },
);
await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
await page.waitForTimeout(800);
await page.screenshot({ path: `${outDir}/3928-full-sheet.png` });

// Sample canvas ink for the exploded yellow slice (approx #ffd320) near the second chart.
const stats = await page.locator("#viewer").evaluate((viewer) => {
  const canvas = viewer.shadowRoot.querySelector(".stage canvas");
  const ctx = canvas.getContext("2d");
  const { data, width, height } = ctx.getImageData(0, 0, canvas.width, canvas.height);
  let yellow = 0;
  let blue = 0;
  let yellowXs = [];
  let yellowYs = [];
  for (let y = 0; y < height; y += 2) {
    for (let x = 0; x < width; x += 2) {
      const i = (y * width + x) * 4;
      const r = data[i];
      const g = data[i + 1];
      const b = data[i + 2];
      const a = data[i + 3];
      if (a < 200) continue;
      if (r > 220 && g > 180 && g < 230 && b < 80) {
        yellow += 1;
        yellowXs.push(x);
        yellowYs.push(y);
      }
      if (r < 40 && g < 90 && b > 100 && b < 170) {
        blue += 1;
      }
    }
  }
  const cluster = (values) => {
    if (values.length === 0) return null;
    values.sort((a, b) => a - b);
    return {
      min: values[0],
      max: values[values.length - 1],
      mid: values[Math.floor(values.length / 2)],
    };
  };
  return {
    width,
    height,
    yellow,
    blue,
    yellowX: cluster(yellowXs),
    yellowY: cluster(yellowYs),
  };
});
console.log("canvas stats", JSON.stringify(stats, null, 2));
await page.screenshot({ path: `${outDir}/3928-capture.png`, fullPage: true });
await browser.close();
console.log("wrote", `${outDir}/3928-full-sheet.png`, `${outDir}/3928-capture.png`);
