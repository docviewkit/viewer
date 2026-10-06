import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

const manifest = JSON.parse(readFileSync(
  new URL("./corpus/py-pdf-sample-files.json", import.meta.url),
  "utf8",
));
const runner = readFileSync(
  new URL("../scripts/run-pdf-corpus.mjs", import.meta.url),
  "utf8",
);

test("py-pdf corpus is immutable, bounded, and split into smoke and full suites", () => {
  assert.equal(manifest.schemaVersion, 1);
  assert.match(manifest.suite.commit, /^[0-9a-f]{40}$/u);
  assert.equal(manifest.suite.license, "CC-BY-SA-4.0");
  assert.equal(manifest.samples.length, 34);
  assert.equal(manifest.samples.reduce((sum, sample) => sum + sample.pages, 0), 293);
  assert.equal(manifest.samples.filter(({ smoke }) => smoke).length, 11);
  assert.equal(new Set(manifest.samples.map(({ path }) => path)).size, 34);
  for (const sample of manifest.samples) {
    assert.match(sample.path, /^[^/].*\.pdf$/u);
    assert.equal(sample.path.includes(".."), false);
    assert.match(sample.sha256, /^[0-9a-f]{64}$/u);
    assert.equal(Number.isSafeInteger(sample.bytes) && sample.bytes > 0, true);
    assert.equal(Number.isSafeInteger(sample.pages) && sample.pages > 0, true);
  }
});

test("PDF corpus runner verifies bytes, materializes pages, and reuses the visual oracle", () => {
  assert.match(runner, /fetch", "--quiet", "--depth", "1", "origin", manifest\.suite\.commit/u);
  assert.match(runner, /bytes\.length !== sample\.bytes/u);
  assert.match(runner, /sha256 !== sample\.sha256/u);
  assert.match(runner, /PDF_PASSWORD_REQUIRED/u);
  assert.match(runner, /password: "openpassword"/u);
  assert.match(runner, /document\.listObjects/u);
  assert.match(runner, /run-pdf-preview-suite\.mjs/u);
  assert.match(runner, /environmentFingerprint\?\.matches !== true/u);
  assert.match(runner, /overallScore.*minimumScore/u);
  assert.match(runner, /tolerantPixelSimilarity.*minimumTolerantSimilarity/u);
  assert.match(runner, /023-cmyk-image\/cmyk-image\.pdf", "1", 92, 0\.65/u);
  assert.match(runner, /"--pages", "all"/u);
  assert.match(runner, /tolerantPixelSimilarity < 0\.99/u);
  assert.match(runner, /unrenderedPages !== 0/u);
  assert.doesNotMatch(runner, /refs\/heads\/main|origin\/main/u);
});
