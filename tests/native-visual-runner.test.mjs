import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { deflateSync } from "node:zlib";
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

import { sha256 } from "../scripts/native-visual-policy.mjs";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const digest = `sha256:${"0".repeat(64)}`;

function crc32(bytes) {
  let crc = 0xffff_ffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb8_8320 & -(crc & 1));
  }
  return (crc ^ 0xffff_ffff) >>> 0;
}

function chunk(type, data) {
  const typeBytes = Buffer.from(type, "ascii");
  const result = Buffer.alloc(12 + data.length);
  result.writeUInt32BE(data.length, 0);
  typeBytes.copy(result, 4);
  data.copy(result, 8);
  result.writeUInt32BE(crc32(Buffer.concat([typeBytes, data])), 8 + data.length);
  return result;
}

function rgbaPng(width, height, rgba) {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8;
  header[9] = 6;
  const scanlines = Buffer.alloc(height * (1 + width * 4));
  for (let row = 0; row < height; row += 1) {
    const sourceOffset = row * width * 4;
    const destinationOffset = row * (1 + width * 4);
    scanlines[destinationOffset] = 0;
    rgba.copy(scanlines, destinationOffset + 1, sourceOffset, sourceOffset + width * 4);
  }
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk("IHDR", header),
    chunk("IDAT", deflateSync(scanlines)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

function whitePng(width = 1, height = 1) {
  return rgbaPng(width, height, Buffer.alloc(width * height * 4, 255));
}

function candidateFingerprint() {
  return {
    os: "macOS",
    osVersion: "26.0",
    architecture: "arm64",
    locale: "zh-CN",
    timezone: "Asia/Shanghai",
    colorSpace: "srgb",
    devicePixelRatio: 1,
    scale: 1,
    background: "#ffffff",
    fontSetDigest: digest,
    browser: "Chromium",
    browserVersion: "150.0.0.0",
  };
}

async function fixture(temporary, images = {}) {
  const fixtureBytes = Buffer.from("reviewed-pptx-fixture");
  const goldenPng = images.golden ?? whitePng();
  const actualPng = images.actual ?? goldenPng;
  const actualWidth = actualPng.readUInt32BE(16);
  const actualHeight = actualPng.readUInt32BE(20);
  await mkdir(resolve(temporary, "fixtures"));
  await mkdir(resolve(temporary, "goldens"));
  await mkdir(resolve(temporary, "actual", "candidate"), { recursive: true });
  await writeFile(resolve(temporary, "fixtures", "sample.pptx"), fixtureBytes);
  await writeFile(resolve(temporary, "goldens", "sample.png"), goldenPng);
  const referencePdf = Buffer.from("%PDF-1.7\nsynthetic reviewed reference\n", "ascii");
  await writeFile(resolve(temporary, "goldens", "reference.pdf"), referencePdf);
  await writeFile(resolve(temporary, "actual", "candidate", "sample.png"), actualPng);
  await writeFile(resolve(temporary, "actual", "candidate", "sample.json"), JSON.stringify({
    schemaVersion: 1,
    fixtureSha256: sha256(fixtureBytes),
    format: "pptx",
    unitIndex: 0,
    unitType: "slide",
    unitCount: 1,
    width: actualWidth,
    height: actualHeight,
    pngSha256: sha256(actualPng),
    diagnostics: [],
  }));
  const reference = Buffer.from(`${JSON.stringify({
    schemaVersion: 1,
    oracleSuite: "microsoft-office",
    source: { file: "sample.pptx", format: "pptx", sha256: sha256(fixtureBytes).slice("sha256:".length) },
    referenceApplication: {
      name: "Microsoft PowerPoint",
      bundleId: "com.microsoft.Powerpoint",
      version: "16.111",
      build: "16.111.1",
    },
    captureSemantics: {
      unitType: "slide",
      capture: "pdf-export",
      rasterDpi: 96,
      nativeApplicationAutomation: true,
    },
    rasterizer: { name: "pdftoppm", version: "26.05.0", dpi: 96 },
    pdfs: [{
      file: "reference.pdf",
      bytes: referencePdf.length,
      sha256: sha256(referencePdf).slice("sha256:".length),
    }],
    pages: [{ file: "sample.png", pageIndex: 0, sha256: sha256(goldenPng).slice("sha256:".length) }],
  })}\n`);
  await writeFile(resolve(temporary, "goldens", "reference.json"), reference);
  const fingerprint = candidateFingerprint();
  const suite = {
    schemaVersion: 2,
    suiteKind: "regression",
    oracleMode: "read-only",
    oracleSuite: "microsoft-office",
    thresholdPolicy: "native-visual-v1",
    oracleFingerprint: {
      os: "macOS",
      osVersion: "26.0",
      architecture: "arm64",
      locale: "zh-CN",
      timezone: "Asia/Shanghai",
      colorSpace: "srgb",
      scale: 1,
      background: "#ffffff",
      fontSetDigest: digest,
      applications: { powerpoint: { version: "16.111", build: "16.111.1", capture: "pdf-export" } },
      rasterizer: { name: "pdftoppm", version: "26.05.0", dpi: 96 },
    },
    candidateFingerprint: fingerprint,
    documents: [{
      id: "sample-pptx",
      corpusClass: "minimal",
      format: "pptx",
      fixture: "fixtures/sample.pptx",
      fixtureSha256: sha256(fixtureBytes),
      unitCount: 1,
      units: [{
        index: 0,
        referenceJson: "goldens/reference.json",
        referenceJsonSha256: sha256(reference),
        goldenPng: "goldens/sample.png",
        goldenPngSha256: sha256(goldenPng),
        actualPng: "candidate/sample.png",
        actualObservationJson: "candidate/sample.json",
      }],
    }],
  };
  await writeFile(resolve(temporary, "suite.json"), JSON.stringify(suite));
  await writeFile(resolve(temporary, "fingerprint.json"), JSON.stringify({ ...fingerprint, buildRevision: "test" }));
  return suite;
}

test("schema v2 visual runner verifies native provenance and uses the locked policy", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-native-runner-"));
  try {
    await fixture(temporary);
    const output = resolve(temporary, "reports");
    const result = spawnSync(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      resolve(temporary, "suite.json"),
      "--fingerprint", resolve(temporary, "fingerprint.json"),
      "--actual-root", resolve(temporary, "actual"),
      "--output", output,
    ], { cwd: root, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr || result.stdout);
    const report = JSON.parse(await readFile(resolve(output, "sample-pptx-unit-0.json"), "utf8"));
    assert.equal(report.passed, true);
    assert.equal(report.oracle.suite, "microsoft-office");
    assert.equal(report.oracle.renderer, "powerpoint");
    assert.equal(report.oracle.reference.file, "goldens/reference.json");
    assert.equal(report.oracle.reference.capture.capture, "pdf-export");
    assert.equal(report.oracle.reference.capture.nativeApplicationAutomation, true);
    assert.equal(report.thresholdPolicy, "native-visual-v1");
    assert.equal(report.visualDiff.pixelRadius, 1);
    const summary = JSON.parse(await readFile(resolve(output, "summary.json"), "utf8"));
    assert.equal(summary.status, "passed");
    assert.deepEqual(summary.formats.pptx, { total: 1, passed: 1, failed: 0 });
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("schema v2 visual runner rejects a concentrated local defect even when whole-page metrics pass", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-native-local-region-"));
  try {
    const width = 200;
    const height = 200;
    const actualRgba = Buffer.alloc(width * height * 4, 255);
    for (let y = 96; y < 99; y += 1) {
      for (let x = 96; x < 99; x += 1) {
        const offset = (y * width + x) * 4;
        actualRgba[offset] = 0;
        actualRgba[offset + 1] = 0;
        actualRgba[offset + 2] = 0;
      }
    }
    await fixture(temporary, { golden: whitePng(width, height), actual: rgbaPng(width, height, actualRgba) });
    const output = resolve(temporary, "reports");
    const result = spawnSync(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      resolve(temporary, "suite.json"),
      "--fingerprint", resolve(temporary, "fingerprint.json"),
      "--actual-root", resolve(temporary, "actual"),
      "--output", output,
    ], { cwd: root, encoding: "utf8" });
    assert.equal(result.status, 1, result.stderr || result.stdout);
    const report = JSON.parse(await readFile(resolve(output, "sample-pptx-unit-0.json"), "utf8"));
    assert.equal(report.passed, false);
    assert.equal(report.visualSimilarity.metrics.tolerantPixelSimilarity > 0.995, true);
    assert.equal(report.visualSimilarity.metrics.ssim > 0.99, true);
    assert.equal(report.failureReasons.some(({ code }) => code === "OBJECT_REGION_DIFF"), true);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("schema v2 visual runner rejects a changed reviewed golden", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-native-golden-"));
  try {
    await fixture(temporary);
    await writeFile(resolve(temporary, "goldens", "sample.png"), Buffer.from("not-the-reviewed-png"));
    const output = resolve(temporary, "reports");
    const result = spawnSync(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      resolve(temporary, "suite.json"),
      "--fingerprint", resolve(temporary, "fingerprint.json"),
      "--actual-root", resolve(temporary, "actual"),
      "--output", output,
    ], { cwd: root, encoding: "utf8" });
    assert.equal(result.status, 1);
    const report = JSON.parse(await readFile(resolve(output, "sample-pptx-unit-0.json"), "utf8"));
    assert.equal(report.passed, false);
    assert.match(report.failureReasons[0].message, /golden PNG SHA-256/u);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("schema v2 visual runner rejects reference metadata that no longer authenticates the golden", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-native-reference-chain-"));
  try {
    await fixture(temporary);
    await writeFile(resolve(temporary, "goldens", "reference.json"), "{}\n");
    const output = resolve(temporary, "reports");
    const result = spawnSync(process.execPath, [
      "scripts/run-external-visual-suite.mjs",
      resolve(temporary, "suite.json"),
      "--fingerprint", resolve(temporary, "fingerprint.json"),
      "--actual-root", resolve(temporary, "actual"),
      "--output", output,
    ], { cwd: root, encoding: "utf8" });
    assert.equal(result.status, 1);
    const report = JSON.parse(await readFile(resolve(output, "sample-pptx-unit-0.json"), "utf8"));
    assert.match(report.failureReasons[0].message, /reference\.json SHA-256/u);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});
