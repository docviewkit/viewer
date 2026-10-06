import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const runner = await readFile(
  new URL("../scripts/run-pdfjs-performance-comparison.mjs", import.meta.url),
  "utf8",
);

test("PDF.js performance comparison uses twelve representative physical-size cases", () => {
  const cases = runner.match(/\{ file: "[^"]+\.pdf", page: \d+, feature: "[^"]+" \}/gu) ?? [];
  assert.equal(cases.length, 12);
  assert.match(runner, /file: "S2\.pdf"/u);
  assert.match(runner, /const pointScale = 0\.75/u);
  assert.match(runner, /maximumPixels = 1_200_000/u);
});

test("PDF.js performance comparison alternates foreground renderers and reports the 10% gate", () => {
  assert.match(runner, /sample % 2 === 0/u);
  assert.match(runner, /await page\.bringToFront\(\)/u);
  assert.match(runner, /geometricMeanAdvantagePct >= 10/u);
  assert.match(runner, /firstFrameAdvantagePct >= 10/u);
});
