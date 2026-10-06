import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { deflateSync } from "node:zlib";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");

function crc32(bytes) {
  let crc = 0xffff_ffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb8_8320 & -(crc & 1));
  }
  return (crc ^ 0xffff_ffff) >>> 0;
}

function pngChunk(type, data) {
  const typeBytes = Buffer.from(type, "ascii");
  const chunk = Buffer.alloc(12 + data.length);
  chunk.writeUInt32BE(data.length, 0);
  typeBytes.copy(chunk, 4);
  data.copy(chunk, 8);
  chunk.writeUInt32BE(crc32(Buffer.concat([typeBytes, data])), 8 + data.length);
  return chunk;
}

function rgbaPng(width, height, rgba) {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8;
  header[9] = 6;
  const scanlines = Buffer.alloc(height * (width * 4 + 1));
  for (let y = 0; y < height; y += 1) {
    const row = y * (width * 4 + 1);
    scanlines[row] = 0;
    Buffer.from(rgba).copy(scanlines, row + 1, y * width * 4, (y + 1) * width * 4);
  }
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    pngChunk("IHDR", header),
    pngChunk("IDAT", deflateSync(scanlines)),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

function environmentFingerprint(overrides = {}) {
  return {
    os: "macOS",
    osVersion: "15.5",
    architecture: "arm64",
    locale: "zh-CN",
    timezone: "Asia/Shanghai",
    colorSpace: "srgb",
    devicePixelRatio: 2,
    scale: 1,
    background: "#ffffff",
    fontSetDigest: `sha256:${"0".repeat(64)}`,
    referenceRenderer: "Reviewed Chromium golden",
    referenceRendererVersion: "Chromium 138.0.7204.50",
    candidateRenderer: "OfficeViewer Chromium capture",
    candidateRendererVersion: "test-revision",
    ...overrides,
  };
}

test("external visual runner compares PNG goldens and writes machine-readable diffs", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-external-accuracy-"));
  try {
    const white = [255, 255, 255, 255, 255, 255, 255, 255];
    const changed = [255, 255, 255, 255, 0, 0, 0, 255];
    await writeFile(resolve(temporary, "golden.png"), rgbaPng(2, 1, white));
    await writeFile(resolve(temporary, "same.png"), rgbaPng(2, 1, white));
    await writeFile(resolve(temporary, "changed.png"), rgbaPng(2, 1, changed));
    const fingerprint = environmentFingerprint();
    await writeFile(resolve(temporary, "current-environment.json"), JSON.stringify(fingerprint));
    await writeFile(resolve(temporary, "suite.json"), JSON.stringify({
      schemaVersion: 1,
      oracleMode: "read-only",
      environmentFingerprint: fingerprint,
      cases: [
        { id: "same", corpusClass: "minimal", fixture: "sample.pptx", unitIndex: 0, unitType: "slide", goldenPng: "golden.png", actualPng: "same.png" },
        { id: "different", corpusClass: "minimal", fixture: "sample.pptx", unitIndex: 0, unitType: "slide", goldenPng: "golden.png", actualPng: "changed.png" },
      ],
    }));
    const outputPath = resolve(temporary, "reports");
    const result = spawnSync(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      resolve(temporary, "suite.json"),
      "--fingerprint",
      resolve(temporary, "current-environment.json"),
      "--output",
      outputPath,
    ], { cwd: root, encoding: "utf8" });

    assert.equal(result.status, 1, result.stderr || result.stdout);
    const same = JSON.parse(await readFile(resolve(outputPath, "same.json"), "utf8"));
    assert.equal(same.passed, true);
    assert.equal(same.visualDiff.differingPixelCount, 0);
    assert.equal(same.visualDiff.differenceBounds, null);
    const different = JSON.parse(await readFile(resolve(outputPath, "different.json"), "utf8"));
    assert.equal(different.passed, false);
    assert.equal(different.visualDiff.differingPixelCount, 1);
    assert.equal(different.visualDiff.tolerantDifferingPixelCount, 1);
    assert.equal(different.visualDiff.maxChannelDelta, 255);
    assert.equal(different.visualDiff.meanAbsoluteChannelDelta, 95.625);
    assert.deepEqual(different.visualDiff.differenceBounds, { x: 1, y: 0, width: 1, height: 1 });
    assert.equal(different.failureReasons.some(({ code }) => code === "EXACT_PIXEL_DIFF"), true);
    const summary = JSON.parse(await readFile(resolve(outputPath, "summary.json"), "utf8"));
    assert.equal(summary.environmentFingerprint.matches, true);
    assert.equal(summary.total, 2);
    assert.equal(summary.passed, 1);
    assert.equal(summary.failed, 1);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("external visual runner refuses an unpinned capture environment", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-external-environment-"));
  try {
    const expected = environmentFingerprint();
    const actual = environmentFingerprint({ candidateRendererVersion: "Chromium 139" });
    await writeFile(resolve(temporary, "expected.png"), rgbaPng(1, 1, [255, 255, 255, 255]));
    await writeFile(resolve(temporary, "actual.png"), rgbaPng(1, 1, [255, 255, 255, 255]));
    await writeFile(resolve(temporary, "current-environment.json"), JSON.stringify(actual));
    await writeFile(resolve(temporary, "suite.json"), JSON.stringify({
      schemaVersion: 1,
      oracleMode: "read-only",
      environmentFingerprint: expected,
      cases: [{
        id: "environment-mismatch",
        corpusClass: "minimal",
        fixture: "sample.pptx",
        unitIndex: 0,
        unitType: "slide",
        goldenPng: "expected.png",
        actualPng: "actual.png",
      }],
    }));
    const outputPath = resolve(temporary, "reports");
    const result = spawnSync(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      resolve(temporary, "suite.json"),
      "--fingerprint",
      resolve(temporary, "current-environment.json"),
      "--output",
      outputPath,
    ], { cwd: root, encoding: "utf8" });

    assert.equal(result.status, 1);
    const summary = JSON.parse(await readFile(resolve(outputPath, "summary.json"), "utf8"));
    assert.equal(summary.environmentFingerprint.matches, false);
    assert.deepEqual(summary.environmentFingerprint.changedFields, ["candidateRendererVersion"]);
    assert.equal(summary.failureReasons[0].code, "ENVIRONMENT_FINGERPRINT_MISMATCH");
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("external visual runner rejects editor-mode screenshots as oracles", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-external-mode-"));
  try {
    const fingerprint = environmentFingerprint();
    const png = rgbaPng(1, 1, [255, 255, 255, 255]);
    await writeFile(resolve(temporary, "golden.png"), png);
    await writeFile(resolve(temporary, "actual.png"), png);
    await writeFile(resolve(temporary, "current-environment.json"), JSON.stringify(fingerprint));
    await writeFile(resolve(temporary, "suite.json"), JSON.stringify({
      schemaVersion: 1,
      oracleMode: "editor",
      environmentFingerprint: fingerprint,
      cases: [{
        id: "editor-is-not-an-oracle",
        corpusClass: "minimal",
        fixture: "sample.pptx",
        unitIndex: 0,
        unitType: "slide",
        goldenPng: "golden.png",
        actualPng: "actual.png",
      }],
    }));
    const outputPath = resolve(temporary, "reports");
    const result = spawnSync(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      resolve(temporary, "suite.json"),
      "--fingerprint",
      resolve(temporary, "current-environment.json"),
      "--output",
      outputPath,
    ], { cwd: root, encoding: "utf8" });

    assert.equal(result.status, 1);
    const summary = JSON.parse(await readFile(resolve(outputPath, "summary.json"), "utf8"));
    assert.equal(summary.status, "invalid-suite");
    assert.equal(summary.failureReasons[0].code, "EXTERNAL_VISUAL_SUITE_INVALID");
    assert.match(summary.failureReasons[0].message, /oracleMode must be read-only/u);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("external visual runner evaluates paired object and text observations", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-external-objects-"));
  try {
    const fingerprint = environmentFingerprint();
    const png = rgbaPng(2, 1, [255, 255, 255, 255, 255, 255, 255, 255]);
    await writeFile(resolve(temporary, "golden.png"), png);
    await writeFile(resolve(temporary, "actual.png"), png);
    await writeFile(resolve(temporary, "current-environment.json"), JSON.stringify(fingerprint));
    const snapshot = {
      units: [{ index: 0, type: "slide", width: 2, height: 1 }],
      objects: [{
        id: "object:1",
        sourceKey: "ppt/slides/slide1.xml#shape=7",
        type: "text-box",
        unitIndex: 0,
        bounds: { x: 0, y: 0, width: 2, height: 1 },
        content: { text: "Title" },
        textLayout: {
          lines: [{ text: "Title", start: 0, end: 5, baseline: 0.8, fontSize: 12, lineHeight: 14 }],
          overflow: false,
          autoScale: 1,
        },
      }],
    };
    await writeFile(resolve(temporary, "expected.json"), JSON.stringify(snapshot));
    await writeFile(resolve(temporary, "actual.json"), JSON.stringify(snapshot));
    await writeFile(resolve(temporary, "suite.json"), JSON.stringify({
      schemaVersion: 1,
      oracleMode: "read-only",
      environmentFingerprint: fingerprint,
      cases: [{
        id: "slide-object-observation",
        corpusClass: "minimal",
        fixture: "sample.pptx",
        unitIndex: 0,
        unitType: "slide",
        goldenPng: "golden.png",
        actualPng: "actual.png",
        expectedObservationJson: "expected.json",
        actualObservationJson: "actual.json",
        declaredCoverage: [
          "structure-units",
          "structure-objects",
          "content-text",
          "geometry-bounds",
          "text-lines",
          "text-baseline",
          "text-font-size",
          "text-line-spacing",
          "visual-exact",
          "visual-tolerant",
          "visual-ssim",
        ],
      }],
    }));
    const outputPath = resolve(temporary, "reports");
    const result = spawnSync(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      resolve(temporary, "suite.json"),
      "--fingerprint",
      resolve(temporary, "current-environment.json"),
      "--output",
      outputPath,
    ], { cwd: root, encoding: "utf8" });

    assert.equal(result.status, 0, result.stderr || result.stdout);
    const report = JSON.parse(await readFile(resolve(outputPath, "slide-object-observation.json"), "utf8"));
    assert.equal(report.contentCompleteness.applicable, true);
    assert.equal(report.geometryAccuracy.applicable, true);
    assert.equal(report.textLayoutAccuracy.applicable, true);
    assert.equal(report.textLayoutAccuracy.score, 100);
    assert.deepEqual(report.coverageValidation.missing, []);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});
