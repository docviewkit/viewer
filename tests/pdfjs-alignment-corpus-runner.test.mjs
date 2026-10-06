import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const runner = await readFile(
  new URL("../scripts/run-pdfjs-alignment-corpus.mjs", import.meta.url),
  "utf8",
);
const acrobatOracle = JSON.parse(await readFile(
  new URL("./pdfjs-acrobat-oracle.json", import.meta.url),
  "utf8",
));

test("PDF.js alignment runner consumes rendered pages and captures Acrobat-required files", () => {
  assert.match(runner, /options\.includeAcrobatRequired && status === "pdfjs-failed"/u);
  assert.match(runner, /hasReferencePages \? Math\.min\(harness\.unitCount, entry\.pageCount\) : harness\.unitCount/u);
  assert.match(runner, /!hasReferencePages \|\| harness\.unitCount === entry\.pageCount \? "captured" : "unit-count-mismatch"/u);
  assert.match(runner, /oracle: hasReferencePages \? "pdf\.js" : "acrobat-required"/u);
  assert.match(runner, /!Array\.isArray\(reference\.pages\)/u);
  assert.match(runner, /pendingPdfs:\s*references\.length\s*-\s*results\.length/u);
});

test("Acrobat fallback outcomes cover every PDF.js reference failure", () => {
  assert.equal(acrobatOracle.schemaVersion, 1);
  assert.deepEqual(Object.keys(acrobatOracle.files).sort(), [
    "Pages-tree-refs.pdf",
    "REDHAT-1531897-0.pdf",
    "bug1020226.pdf",
    "encrypted-attachment.pdf",
    "poppler-395-0-fuzzed.pdf",
    "poppler-742-0-fuzzed.pdf",
    "poppler-937-0-fuzzed.pdf",
    "pr6531_1.pdf",
    "pr6531_2.pdf",
  ]);
});

test("PDF.js alignment runner compares like-sized browser captures without application chrome", () => {
  assert.match(runner, /examples\/visual-harness\.html/u);
  assert.match(runner, /test_manifest\.json/u);
  assert.match(runner, /url\.searchParams\.set\("password", password\)/u);
  assert.match(runner, /url\.searchParams\.set\("scale", "0\.75"\)/u);
  assert.match(runner, /globalThis\.visualHarness\.renderUnit\(unitIndex\)/u);
  assert.match(runner, /pixelChannel:\s*channelTolerance/u);
  assert.match(runner, /pixelRadius/u);
  assert.match(runner, /minTolerantPixelRatio:\s*0\.98/u);
  assert.match(runner, /minSsim:\s*0\.90/u);
  assert.match(runner, /widthDelta <= 1 && heightDelta <= 1/u);
  assert.match(runner, /context\.drawImage\(image, 0, 0\)/u);
  assert.match(runner, /comparisonNormalized:\s*true/u);
});

test("OfficeViewer failures remain explicit and resumable", () => {
  assert.match(runner, /status:\s*oracleAligned \? "oracle-aligned-failure" : "officeviewer-failed"/u);
  assert.match(runner, /Missing Acrobat oracle result/u);
  assert.match(runner, /oracleAlignedFailures/u);
  assert.match(runner, /writeJsonAtomic\(candidateLedgerPath/u);
  assert.match(runner, /--resume/u);
  assert.match(runner, /failedPdfs/u);
  assert.match(runner, /previous\?\.candidateDigest === buildDigest/u);
  assert.match(runner, /dist\/office-viewer-pdf\.wasm/u);
  assert.match(runner, /result\.status === "officeviewer-failed"/u);
  assert.match(runner, /const retryBrowser = await chromium\.launch/u);
  assert.match(runner, /await closeBrowser\(retryBrowser\)/u);
});
