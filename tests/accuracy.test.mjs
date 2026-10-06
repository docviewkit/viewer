import assert from "node:assert/strict";
import test from "node:test";

import {
  evaluateAccuracy,
  REQUIRED_ACCURACY_COVERAGE,
  validateAccuracyCorpus,
} from "../dist/accuracy.js";

function exactSnapshot() {
  return {
    units: [{ index: 0, type: "slide", width: 1000, height: 500 }],
    objects: [{
      id: "object:7",
      sourceKey: "ppt/slides/slide1.xml#shape:7",
      type: "text-box",
      unitIndex: 0,
      bounds: { x: 100, y: 50, width: 300, height: 100 },
      rotation: 15,
      crop: { x: 0, y: 0, width: 1, height: 1 },
      zOrder: 2,
      content: { text: "hello", digest: "sha256:text" },
      textLayout: {
        lines: [{ text: "hello", start: 0, end: 5, baseline: 72, fontSize: 20, lineHeight: 24 }],
        overflow: false,
        autoScale: 1,
      },
    }],
    visuals: [{
      unitIndex: 0,
      width: 2,
      height: 2,
      data: new Uint8Array([
        255, 255, 255, 255, 0, 0, 0, 255,
        0, 0, 0, 255, 255, 255, 255, 255,
      ]),
    }],
    hitProbes: [{
      id: "center-7",
      kind: "center",
      unitIndex: 0,
      x: 0.25,
      y: 0.2,
      hits: ["ppt/slides/slide1.xml#shape:7"],
    }],
    diagnostics: [{
      code: "FONT_SUBSTITUTED",
      fidelity: "approximate",
      phase: "render",
      part: "ppt/slides/slide1.xml",
      sourceKey: "ppt/slides/slide1.xml#shape:7",
    }],
    displayList: [{ op: "text", objectId: "object:7", x: 100, y: 50 }],
  };
}

function solidSurface(width, height, level) {
  const pixels = new Uint8Array(width * height * 4);
  for (let offset = 0; offset < pixels.length; offset += 4) {
    pixels[offset] = level;
    pixels[offset + 1] = level;
    pixels[offset + 2] = level;
    pixels[offset + 3] = 255;
  }
  return pixels;
}

function paintVerticalStroke(pixels, width, x, top, bottom, level) {
  for (let y = top; y < bottom; y += 1) {
    const offset = (y * width + x) * 4;
    pixels[offset] = level;
    pixels[offset + 1] = level;
    pixels[offset + 2] = level;
  }
}

function evaluateTextRegion({
  fileId,
  width,
  height,
  expectedPixels,
  actualPixels,
  bounds,
  pixelRadius = 3,
}) {
  const base = {
    units: [{ index: 0, type: "page", width, height }],
    objects: [{
      id: "text:1",
      sourceKey: "text:1",
      type: "text-region",
      unitIndex: 0,
      bounds,
    }],
  };
  return evaluateAccuracy({
    fileId,
    corpusClass: "minimal",
    expected: {
      ...base,
      visuals: [{ unitIndex: 0, width, height, data: expectedPixels }],
    },
    actual: {
      ...base,
      visuals: [{ unitIndex: 0, width, height, data: actualPixels }],
    },
    policy: {
      minExactPixelRatio: 0,
      minTolerantPixelRatio: 0.98,
      minSsim: 0,
      minObjectRegionSimilarity: 0.65,
      pixelRadius,
    },
  });
}

test("empty snapshots cannot produce a perfect accuracy result", () => {
  const empty = { units: [], objects: [], diagnostics: [] };
  const report = evaluateAccuracy({
    fileId: "empty.pptx",
    corpusClass: "minimal",
    expected: empty,
    actual: empty,
  });

  assert.equal(report.passed, false);
  assert.equal(report.overallScore, 0);
  assert.deepEqual(
    report.failureReasons.map(({ code }) => code),
    ["EXPECTED_SNAPSHOT_EMPTY", "ACTUAL_SNAPSHOT_EMPTY"],
  );
});

test("declared coverage must be backed by independent oracle observations", () => {
  const snapshot = exactSnapshot();
  const declaredCoverage = [
    "structure-units",
    "content-text",
    "content-image",
    "hit-z-order",
    "determinism-display-list",
  ];
  const report = evaluateAccuracy({
    fileId: "coverage.pptx",
    corpusClass: "minimal",
    expected: snapshot,
    actual: snapshot,
    declaredCoverage,
  });

  assert.deepEqual(report.coverageValidation, {
    declared: declaredCoverage,
    observed: ["structure-units", "content-text"],
    missing: ["content-image", "hit-z-order", "determinism-display-list"],
  });
  assert.equal(report.passed, false);
  assert.equal(report.overallScore, 0);
  assert.deepEqual(
    report.failureReasons
      .filter(({ code }) => code === "DECLARED_COVERAGE_NOT_OBSERVED")
      .map(({ coverage }) => coverage),
    ["content-image", "hit-z-order", "determinism-display-list"],
  );
});

test("matching observations produce a complete perfect AccuracyReport", () => {
  const snapshot = exactSnapshot();
  const report = evaluateAccuracy({
    fileId: "minimal-text.pptx",
    corpusClass: "minimal",
    expected: snapshot,
    actual: snapshot,
    repeatedRuns: [snapshot, snapshot],
    metamorphicVariants: [
      { kind: "xml-attribute-order", snapshot },
      { kind: "namespace-prefix", snapshot },
      { kind: "zip-entry-order", snapshot },
      { kind: "irrelevant-metadata", snapshot },
    ],
  });

  assert.equal(report.fileId, "minimal-text.pptx");
  assert.equal(report.contentCompleteness.score, 100);
  assert.equal(report.geometryAccuracy.score, 100);
  assert.equal(report.textLayoutAccuracy.score, 100);
  assert.equal(report.visualSimilarity.score, 100);
  assert.equal(report.objectMappingAccuracy.score, 100);
  assert.equal(report.compatibilityAccuracy.score, 100);
  assert.equal(report.determinismAccuracy.score, 100);
  assert.equal(report.metamorphicAccuracy.score, 100);
  assert.equal(report.overallScore, 100);
  assert.equal(report.passed, true);
  assert.deepEqual(report.failureReasons, []);
});

test("geometry and text baselines use normalized coordinates with configurable tolerances", () => {
  const expected = exactSnapshot();
  const original = expected.objects[0];
  const actual = {
    ...expected,
    objects: [{
      ...original,
      bounds: { x: 101.5, y: 50.5, width: 301.5, height: 100.5 },
      rotation: 15.2,
      crop: { x: 0.001, y: 0, width: 0.999, height: 1 },
      textLayout: {
        ...original.textLayout,
        lines: [{ ...original.textLayout.lines[0], baseline: 72.5 }],
      },
    }],
  };

  const tolerant = evaluateAccuracy({ fileId: "geometry.pptx", corpusClass: "minimal", expected, actual });
  assert.equal(tolerant.geometryAccuracy.score, 100);
  assert.equal(tolerant.textLayoutAccuracy.score, 100);

  const strict = evaluateAccuracy({
    fileId: "geometry.pptx",
    corpusClass: "minimal",
    expected,
    actual,
    policy: { geometry: 0.0001, rotationDegrees: 0.1 },
  });
  const codes = new Set(strict.failureReasons.map(({ code }) => code));
  assert.equal(codes.has("NORMALIZED_BOUNDS_MISMATCH"), true);
  assert.equal(codes.has("ROTATION_MISMATCH"), true);
  assert.equal(codes.has("CROP_MISMATCH"), true);
  assert.equal(codes.has("BASELINE_MISMATCH"), true);
});

test("content completeness distinguishes text, image, table, and cell failures", () => {
  const unit = { index: 0, type: "sheet", width: 100, height: 100 };
  const object = (sourceKey, type, content) => ({
    id: sourceKey,
    sourceKey,
    type,
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 10, height: 10 },
    content,
  });
  const expected = {
    units: [unit],
    objects: [
      object("text", "text-box", { text: "Quarterly report" }),
      object("image", "image", { digest: "sha256:expected" }),
      object("table", "table", { rows: 1, columns: 1, cellSourceKeys: ["cell:A1"] }),
      object("cell:A1", "cell", { address: "A1", value: "42" }),
    ],
  };
  const actual = {
    units: [unit],
    objects: [
      object("text", "text-box", { text: "Quarterly report" }),
      object("image", "image", { digest: "sha256:wrong" }),
      object("table", "table", { rows: 1, columns: 1, cellSourceKeys: ["cell:A1"] }),
    ],
  };

  const report = evaluateAccuracy({ fileId: "content.xlsx", corpusClass: "combination", expected, actual });
  assert.equal(report.contentCompleteness.score < 100, true);
  assert.equal(report.contentCompleteness.metrics["text-boxCompleteness"], 100);
  assert.equal(report.contentCompleteness.metrics.imageCompleteness, 0);
  assert.equal(report.contentCompleteness.metrics.tableCompleteness, 100);
  assert.equal(report.contentCompleteness.metrics.cellCompleteness, 0);
  const codes = new Set(report.failureReasons.map(({ code }) => code));
  assert.equal(codes.has("OBJECT_CONTENT_MISMATCH"), true);
  assert.equal(codes.has("OBJECT_MISSING"), true);
});

test("visual accuracy reports exact, tolerant, SSIM, and object-region metrics", () => {
  const base = {
    units: [{ index: 0, type: "slide", width: 4, height: 1 }],
    objects: [{
      id: "shape:1",
      sourceKey: "shape:1",
      type: "shape",
      unitIndex: 0,
      bounds: { x: 1, y: 0, width: 1, height: 1 },
    }],
  };
  const expectedPixels = new Uint8Array([
    10, 10, 10, 255, 20, 20, 20, 255, 30, 30, 30, 255, 40, 40, 40, 255,
  ]);
  const smallDiff = expectedPixels.slice();
  smallDiff[4] += 5;
  const expected = { ...base, visuals: [{ unitIndex: 0, width: 4, height: 1, data: expectedPixels }] };
  const actual = { ...base, visuals: [{ unitIndex: 0, width: 4, height: 1, data: smallDiff }] };
  const report = evaluateAccuracy({
    fileId: "visual.pptx",
    corpusClass: "minimal",
    expected,
    actual,
    policy: { minExactPixelRatio: 0.75 },
  });
  assert.equal(report.visualSimilarity.metrics.exactPixelSimilarity, 0.75);
  assert.equal(report.visualSimilarity.metrics.tolerantPixelSimilarity, 1);
  assert.equal(report.visualSimilarity.metrics.ssim > 0.99, true);
  assert.equal(report.visualSimilarity.metrics.objectRegionSimilarity, 1);

  const largeDiff = expectedPixels.slice();
  largeDiff[4] = 255;
  const failed = evaluateAccuracy({
    fileId: "visual.pptx",
    corpusClass: "minimal",
    expected,
    actual: { ...base, visuals: [{ unitIndex: 0, width: 4, height: 1, data: largeDiff }] },
    policy: { minExactPixelRatio: 0.75 },
  });
  const codes = new Set(failed.failureReasons.map(({ code }) => code));
  assert.equal(codes.has("TOLERANT_PIXEL_DIFF"), true);
  assert.equal(codes.has("OBJECT_REGION_DIFF"), true);
});

test("object-region evidence catches sparse missing text that whole-page pixels tolerate", () => {
  const width = 100;
  const height = 100;
  const expectedPixels = solidSurface(width, height, 255);
  const actualPixels = expectedPixels.slice();
  // A one-pixel stroke occupies only 10% of its text region. Background-heavy
  // region averages must not let the complete loss of that stroke pass.
  paintVerticalStroke(expectedPixels, width, 14, 10, 20, 0);
  const report = evaluateTextRegion({
    fileId: "sparse-text.pdf",
    width,
    height,
    expectedPixels,
    actualPixels,
    bounds: { x: 10, y: 10, width: 10, height: 10 },
  });
  const codes = new Set(report.failureReasons.map(({ code }) => code));

  assert.equal(report.visualSimilarity.metrics.tolerantPixelSimilarity > 0.98, true);
  assert.equal(report.visualSimilarity.metrics.objectRegionSimilarity < 0.65, true);
  assert.equal(codes.has("TOLERANT_PIXEL_DIFF"), false);
  assert.equal(codes.has("OBJECT_REGION_DIFF"), true);
});

test("object-region evidence catches sparse missing text on a dark background", () => {
  const width = 40;
  const height = 40;
  const expectedPixels = solidSurface(width, height, 24);
  const actualPixels = expectedPixels.slice();
  paintVerticalStroke(expectedPixels, width, 14, 10, 20, 245);
  const report = evaluateTextRegion({
    fileId: "dark-sparse-text.pdf",
    width,
    height,
    expectedPixels,
    actualPixels,
    bounds: { x: 10, y: 10, width: 10, height: 10 },
  });

  assert.equal(report.visualSimilarity.metrics.objectRegionSimilarity < 0.65, true);
  assert.equal(
    report.failureReasons.some(({ code }) => code === "OBJECT_REGION_DIFF"),
    true,
  );
});

test("object-region edge energy catches an adjacent missing stroke", () => {
  const width = 30;
  const height = 30;
  const expectedPixels = solidSurface(width, height, 255);
  const actualPixels = expectedPixels.slice();
  paintVerticalStroke(expectedPixels, width, 10, 10, 20, 0);
  paintVerticalStroke(expectedPixels, width, 12, 10, 20, 0);
  paintVerticalStroke(actualPixels, width, 10, 10, 20, 0);
  const report = evaluateTextRegion({
    fileId: "adjacent-stroke.pdf",
    width,
    height,
    expectedPixels,
    actualPixels,
    bounds: { x: 8, y: 8, width: 7, height: 14 },
  });

  assert.equal(report.visualSimilarity.metrics.objectRegionSimilarity < 0.65, true);
  assert.equal(
    report.failureReasons.some(({ code }) => code === "OBJECT_REGION_DIFF"),
    true,
  );
});

test("object-region tolerance accepts bounded rasterizer displacement at a region edge", () => {
  const width = 20;
  const height = 20;
  const expectedPixels = solidSurface(width, height, 255);
  const actualPixels = expectedPixels.slice();
  paintVerticalStroke(expectedPixels, width, 8, 8, 12, 0);
  paintVerticalStroke(actualPixels, width, 5, 8, 12, 0);
  const report = evaluateTextRegion({
    fileId: "shifted-text.pdf",
    width,
    height,
    expectedPixels,
    actualPixels,
    bounds: { x: 5, y: 8, width: 7, height: 4 },
  });

  assert.equal(report.visualSimilarity.metrics.tolerantPixelSimilarity, 1);
  assert.equal(report.visualSimilarity.metrics.objectRegionSimilarity, 1);
  assert.equal(report.failureReasons.some(({ code }) => code === "OBJECT_REGION_DIFF"), false);
});

test("object-region edge energy treats matching uniform regions as neutral", () => {
  for (const level of [255, 24]) {
    const width = 20;
    const height = 20;
    const expectedPixels = solidSurface(width, height, level);
    const report = evaluateTextRegion({
      fileId: `uniform-${level}.pdf`,
      width,
      height,
      expectedPixels,
      actualPixels: expectedPixels.slice(),
      bounds: { x: 5, y: 5, width: 10, height: 10 },
    });

    assert.equal(report.visualSimilarity.metrics.objectRegionSimilarity, 1);
    assert.equal(
      report.failureReasons.some(({ code }) => code === "OBJECT_REGION_DIFF"),
      false,
    );
  }
});

test("object-region coverage rejects empty or off-surface evidence", () => {
  const pixels = new Uint8Array(4 * 4 * 4).fill(255);
  const report = evaluateAccuracy({
    fileId: "empty-region.pdf",
    corpusClass: "minimal",
    expected: {
      units: [{ index: 0, type: "page", width: 4, height: 4 }],
      objects: [{
        id: "text:outside",
        sourceKey: "text:outside",
        type: "text-box",
        unitIndex: 0,
        bounds: { x: 10, y: 10, width: 2, height: 2 },
      }],
      visuals: [{ unitIndex: 0, width: 4, height: 4, data: pixels }],
    },
    actual: {
      units: [{ index: 0, type: "page", width: 4, height: 4 }],
      objects: [{
        id: "text:outside",
        sourceKey: "text:outside",
        type: "text-box",
        unitIndex: 0,
        bounds: { x: 10, y: 10, width: 2, height: 2 },
      }],
      visuals: [{ unitIndex: 0, width: 4, height: 4, data: pixels.slice() }],
    },
    declaredCoverage: ["visual-object-region"],
  });

  assert.deepEqual(report.coverageValidation.missing, ["visual-object-region"]);
  assert.equal(
    report.failureReasons.some(({ code }) => code === "DECLARED_COVERAGE_NOT_OBSERVED"),
    true,
  );
});

test("object-region comparison accepts an identical region at the shared pixel budget", () => {
  const width = 500;
  const height = 500;
  const expectedPixels = solidSurface(width, height, 255);
  const report = evaluateTextRegion({
    fileId: "bounded-identical-region.pdf",
    width,
    height,
    expectedPixels,
    actualPixels: expectedPixels.slice(),
    bounds: { x: 0, y: 0, width, height },
  });

  assert.equal(report.visualSimilarity.metrics.objectRegionSimilarity, 1);
  assert.equal(
    report.failureReasons.some(({ code }) => code === "OBJECT_REGION_LIMIT"),
    false,
  );
});

test("object-region comparison fails visibly above the shared pixel budget", () => {
  const width = 501;
  const height = 500;
  const expectedPixels = solidSurface(width, height, 255);
  const report = evaluateTextRegion({
    fileId: "oversized-region.pdf",
    width,
    height,
    expectedPixels,
    actualPixels: expectedPixels.slice(),
    bounds: { x: 0, y: 0, width, height },
  });

  assert.equal(
    report.failureReasons.some(({ code }) => code === "OBJECT_REGION_LIMIT"),
    true,
  );
  assert.equal(report.passed, false);
});

test("object-region work budget accounts for both comparison directions", () => {
  const width = 300;
  const height = 200;
  const expectedPixels = solidSurface(width, height, 255);
  const report = evaluateTextRegion({
    fileId: "neighbor-work-region.pdf",
    width,
    height,
    expectedPixels,
    actualPixels: expectedPixels.slice(),
    bounds: { x: 0, y: 0, width, height },
    pixelRadius: 8,
  });
  const failure = report.failureReasons.find(({ code }) => code === "OBJECT_REGION_LIMIT");

  assert.equal(failure?.actual?.neighborWork, 138_720_000);
  assert.equal(report.passed, false);
});

test("object-region comparison fails visibly instead of exceeding its work budget", () => {
  const width = 23;
  const height = 23;
  const pixels = new Uint8Array(width * height * 4).fill(255);
  const objects = Array.from({ length: 513 }, (_, index) => ({
    id: `region:${index}`,
    sourceKey: `region:${index}`,
    type: "text-region",
    unitIndex: 0,
    bounds: { x: index % width, y: Math.floor(index / width), width: 1, height: 1 },
  }));
  const snapshot = {
    units: [{ index: 0, type: "page", width, height }],
    objects,
    visuals: [{ unitIndex: 0, width, height, data: pixels }],
  };
  const report = evaluateAccuracy({
    fileId: "bounded-regions.pdf",
    corpusClass: "minimal",
    expected: snapshot,
    actual: {
      ...snapshot,
      visuals: [{ unitIndex: 0, width, height, data: pixels.slice() }],
    },
    policy: { minExactPixelRatio: 0, minSsim: 0 },
  });

  assert.equal(report.failureReasons.some(({ code }) => code === "OBJECT_REGION_LIMIT"), true);
  assert.equal(report.passed, false);
});

test("tolerant visual accuracy can absorb bounded rasterizer displacement", () => {
  const expectedPixels = new Uint8Array([
    0, 0, 0, 255, 255, 255, 255, 255, 0, 0, 0, 255,
  ]);
  const shiftedPixels = new Uint8Array([
    255, 255, 255, 255, 0, 0, 0, 255, 255, 255, 255, 255,
  ]);
  const snapshot = (data) => ({
    units: [{ index: 0, type: "slide", width: 3, height: 1 }],
    objects: [],
    visuals: [{ unitIndex: 0, width: 3, height: 1, data }],
  });
  const report = evaluateAccuracy({
    fileId: "rasterizer-shift.pptx",
    corpusClass: "minimal",
    expected: snapshot(expectedPixels),
    actual: snapshot(shiftedPixels),
    policy: {
      pixelChannel: 0,
      pixelRadius: 1,
      minExactPixelRatio: 0,
      minTolerantPixelRatio: 1,
      minSsim: 0,
      minOverallScore: 0,
    },
  });

  assert.equal(report.visualSimilarity.metrics.exactPixelSimilarity, 0);
  assert.equal(report.visualSimilarity.metrics.tolerantPixelSimilarity, 1);
  assert.equal(report.visualSimilarity.metrics.pixelRadiusTolerance, 1);
  assert.equal(report.failureReasons.some(({ code }) => code === "TOLERANT_PIXEL_DIFF"), false);

  const white = new Uint8Array([
    255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  ]);
  const extraDarkPixel = evaluateAccuracy({
    fileId: "rasterizer-extra-pixel.pptx",
    corpusClass: "minimal",
    expected: snapshot(white),
    actual: snapshot(shiftedPixels),
    policy: {
      pixelChannel: 0,
      pixelRadius: 1,
      minExactPixelRatio: 0,
      minTolerantPixelRatio: 1,
      minSsim: 0,
      minOverallScore: 0,
    },
  });
  assert.equal(extraDarkPixel.visualSimilarity.metrics.tolerantPixelSimilarity, 0.6667);
});

test("hit probes verify geometry cases, front-to-back order, groups, and source mapping", () => {
  const objects = [
    { id: "group", sourceKey: "group", type: "group", unitIndex: 0, bounds: { x: 0, y: 0, width: 1, height: 1 } },
    { id: "back", sourceKey: "back", type: "shape", unitIndex: 0, bounds: { x: 0, y: 0, width: 1, height: 1 }, parentSourceKey: "group", zOrder: 1 },
    { id: "front", sourceKey: "front", type: "shape", unitIndex: 0, bounds: { x: 0, y: 0, width: 1, height: 1 }, zOrder: 2 },
  ];
  const probes = [
    { id: "center", kind: "center", unitIndex: 0, x: 0.5, y: 0.5, hits: ["front", "back", "group"] },
    { id: "edge", kind: "edge", unitIndex: 0, x: 0, y: 0.5, hits: ["front", "back", "group"] },
    { id: "rotated", kind: "rotated", unitIndex: 0, x: 0.2, y: 0.2, hits: ["front"] },
    { id: "group", kind: "group", unitIndex: 0, x: 0.4, y: 0.4, hits: ["back", "group"] },
    { id: "overlap", kind: "overlap", unitIndex: 0, x: 0.5, y: 0.5, hits: ["front", "back"] },
    { id: "z-order", kind: "z-order", unitIndex: 0, x: 0.5, y: 0.5, hits: ["front", "back"] },
  ];
  const expected = { units: [{ index: 0, type: "slide", width: 1, height: 1 }], objects, hitProbes: probes };
  const actualProbes = probes.map((probe) => probe.id === "overlap"
    ? { ...probe, hits: ["back", "front"] }
    : probe.id === "rotated"
      ? { ...probe, hits: ["unmapped"] }
      : probe);
  const report = evaluateAccuracy({
    fileId: "hit-testing.pptx",
    corpusClass: "combination",
    expected,
    actual: { ...expected, hitProbes: actualProbes },
  });

  assert.equal(report.objectMappingAccuracy.score < 100, true);
  const codes = new Set(report.failureReasons.map(({ code }) => code));
  assert.equal(codes.has("HIT_MAPPING_MISMATCH"), true);
  assert.equal(codes.has("HIT_SOURCE_UNMAPPED"), true);
});

test("compatibility diagnostics classify unsupported, approximate, font, external, and active content", () => {
  const diagnostics = [
    { code: "UNSUPPORTED_FEATURE", fidelity: "unsupported", phase: "parse", part: "slide1.xml" },
    { code: "APPROXIMATE_LAYOUT", fidelity: "approximate", phase: "layout", part: "document.xml" },
    { code: "FONT_SUBSTITUTED", fidelity: "approximate", phase: "render", part: "slide1.xml" },
    { code: "EXTERNAL_RESOURCE_BLOCKED", fidelity: "not-rendered", phase: "security", part: "document.xml" },
    { code: "ACTIVE_CONTENT_BLOCKED", fidelity: "not-rendered", phase: "security", part: "vbaProject.bin" },
  ];
  const snapshot = { units: [], objects: [], diagnostics };
  const exact = evaluateAccuracy({ fileId: "diagnostics.docx", corpusClass: "malicious", expected: snapshot, actual: snapshot });
  assert.equal(exact.compatibilityAccuracy.metrics.unsupportedCount, 1);
  assert.equal(exact.compatibilityAccuracy.metrics.approximateCount, 2);
  assert.equal(exact.compatibilityAccuracy.metrics.fontSubstitutionCount, 1);
  assert.equal(exact.compatibilityAccuracy.metrics.externalResourceCount, 1);
  assert.equal(exact.compatibilityAccuracy.metrics.activeContentCount, 1);

  const missing = evaluateAccuracy({
    fileId: "diagnostics.docx",
    corpusClass: "malicious",
    expected: snapshot,
    actual: { ...snapshot, diagnostics: diagnostics.slice(0, -1) },
  });
  assert.equal(missing.compatibilityAccuracy.score < 100, true);
  assert.equal(missing.failureReasons.some(({ code }) => code === "DIAGNOSTIC_MISMATCH"), true);
});

test("determinism and metamorphic checks identify the changed invariant", () => {
  const snapshot = exactSnapshot();
  const changedObject = {
    ...snapshot.objects[0],
    id: "object:changed",
    textLayout: {
      ...snapshot.objects[0].textLayout,
      lines: [{ ...snapshot.objects[0].textLayout.lines[0], baseline: 80 }],
    },
  };
  const changed = {
    ...snapshot,
    objects: [changedObject],
    displayList: [{ op: "text", objectId: "object:changed", x: 101, y: 50 }],
  };
  const variants = [
    { kind: "xml-attribute-order", snapshot },
    { kind: "namespace-prefix", snapshot },
    { kind: "zip-entry-order", snapshot: changed },
    { kind: "irrelevant-metadata", snapshot },
  ];
  const report = evaluateAccuracy({
    fileId: "invariants.pptx",
    corpusClass: "minimal",
    expected: snapshot,
    actual: snapshot,
    repeatedRuns: [snapshot, changed],
    metamorphicVariants: variants,
  });
  const codes = new Set(report.failureReasons.map(({ code }) => code));
  assert.equal(codes.has("DISPLAY_LIST_NONDETERMINISTIC"), true);
  assert.equal(codes.has("OBJECT_ID_NONDETERMINISTIC"), true);
  assert.equal(codes.has("LAYOUT_NONDETERMINISTIC"), true);
  assert.equal(codes.has("METAMORPHIC_RENDER_CHANGED"), true);
  assert.equal(report.determinismAccuracy.score, 0);
  assert.equal(report.metamorphicAccuracy.score < 100, true);
});

test("corpus governance requires all classes, all layers, and a minimal automated regression per bug", () => {
  const cases = [
    {
      id: "bug-123-minimal",
      corpusClass: "minimal",
      format: "pptx",
      fixture: "corpus/minimal/bug-123.pptx",
      oracle: "corpus/minimal/bug-123.oracle.json",
      automatedTest: "tests/accuracy-regressions.test.mjs#BUG-123",
      coverage: REQUIRED_ACCURACY_COVERAGE,
      regressionBugId: "BUG-123",
    },
    ...["combination", "enterprise", "large", "malformed", "malicious"].map((corpusClass) => ({
      id: `${corpusClass}-suite`,
      corpusClass,
      format: "mixed",
      fixture: `provider:${corpusClass}`,
      oracle: `oracle:${corpusClass}`,
      automatedTest: `runner:${corpusClass}`,
      coverage: [],
    })),
  ];
  const manifest = { compatibilityBugs: ["BUG-123"], cases };
  assert.deepEqual(validateAccuracyCorpus(manifest), { valid: true, issues: [] });
  for (const format of ["ppt", "doc", "xls", "keynote", "pages", "numbers", "pdf"]) {
    const validation = validateAccuracyCorpus({
      ...manifest,
      cases: cases.map((entry, index) => index === 0 ? { ...entry, format } : entry),
    });
    assert.equal(validation.issues.some(({ code }) => code === "CORPUS_FORMAT_INVALID"), false);
  }

  const broken = validateAccuracyCorpus({
    compatibilityBugs: ["BUG-123", "BUG-404"],
    cases: cases
      .filter(({ corpusClass }) => corpusClass !== "enterprise")
      .map((entry) => entry.regressionBugId === "BUG-123" ? { ...entry, automatedTest: "", format: "txt" } : entry),
  });
  assert.equal(broken.valid, false);
  const codes = new Set(broken.issues.map(({ code }) => code));
  assert.equal(codes.has("CORPUS_CLASS_MISSING"), true);
  assert.equal(codes.has("REGRESSION_AUTOMATION_MISSING"), true);
  assert.equal(codes.has("REGRESSION_CASE_MISSING"), true);
  assert.equal(codes.has("CORPUS_FORMAT_INVALID"), true);
});
