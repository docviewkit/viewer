import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import test from "node:test";

import { REQUIRED_ACCURACY_COVERAGE } from "../dist/accuracy.js";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const corpusClasses = ["minimal", "combination", "enterprise", "large", "malformed", "malicious"];

function manifestCases() {
  return corpusClasses.map((corpusClass, index) => ({
    id: `${corpusClass}-case`,
    corpusClass,
    format: "mixed",
    fixture: `fixture:${corpusClass}`,
    oracle: `oracle:${corpusClass}`,
    automatedTest: `suite:${corpusClass}`,
    coverage: index === 0 ? REQUIRED_ACCURACY_COVERAGE : [],
  }));
}

function comprehensiveSuitePrelude(cases) {
  return `
    export const manifest = ${JSON.stringify({ compatibilityBugs: [], cases })};
    const objects = [
      {
        id: "text", sourceKey: "text", type: "text-box", unitIndex: 0,
        bounds: { x: 1, y: 1, width: 2, height: 2 }, rotation: 15, zOrder: 3,
        content: { text: "hello world" },
        textLayout: {
          lines: [
            { text: "hello", start: 0, end: 5, baseline: 1, fontSize: 12, lineHeight: 14 },
            { text: "world", start: 6, end: 11, baseline: 2, fontSize: 12, lineHeight: 14 },
          ],
          overflow: false,
          autoScale: 1,
        },
      },
      {
        id: "image", sourceKey: "image", type: "image", unitIndex: 0,
        bounds: { x: 0, y: 0, width: 1, height: 1 }, crop: { x: 0, y: 0, width: 1, height: 1 },
        content: { digest: "sha256:image" },
      },
      {
        id: "table", sourceKey: "table", type: "table", unitIndex: 0,
        bounds: { x: 2, y: 2, width: 1, height: 1 }, content: { rows: 1, columns: 1 },
      },
      {
        id: "cell", sourceKey: "cell", type: "cell", unitIndex: 0,
        bounds: { x: 2, y: 2, width: 1, height: 1 }, content: { address: "A1", value: "42" },
      },
    ];
    const snapshot = {
      units: [{ index: 0, type: "slide", width: 4, height: 4 }],
      objects,
      visuals: [{ unitIndex: 0, width: 4, height: 4, data: new Uint8Array(64).fill(255) }],
      hitProbes: ["center", "edge", "rotated", "group", "overlap", "z-order"].map((kind) => ({
        id: kind, kind, unitIndex: 0, x: 0.5, y: 0.5, hits: ["text"],
      })),
      diagnostics: [
        { code: "UNSUPPORTED_FEATURE", fidelity: "unsupported", phase: "parse" },
        { code: "APPROXIMATE_LAYOUT", fidelity: "approximate", phase: "layout" },
        { code: "FONT_SUBSTITUTED", fidelity: "approximate", phase: "render" },
        { code: "EXTERNAL_RESOURCE_BLOCKED", fidelity: "not-rendered", phase: "security" },
        { code: "ACTIVE_CONTENT_BLOCKED", fidelity: "not-rendered", phase: "security" },
      ],
      displayList: [{ op: "text", objectId: "text" }],
    };
    const observation = {
      expected: snapshot,
      actual: snapshot,
      repeatedRuns: [snapshot, snapshot],
      metamorphicVariants: ["xml-attribute-order", "namespace-prefix", "zip-entry-order", "irrelevant-metadata"]
        .map((kind) => ({ kind, snapshot })),
    };
  `;
}

test("accuracy suite runner rejects coverage claims without matching oracle evidence", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-accuracy-coverage-"));
  try {
    const suitePath = resolve(temporary, "suite.mjs");
    const outputPath = resolve(temporary, "reports");
    const source = `
      export const manifest = ${JSON.stringify({ compatibilityBugs: [], cases: manifestCases() })};
      const snapshot = { units: [{ index: 0, type: "slide", width: 100, height: 100 }], objects: [], diagnostics: [] };
      export const cases = manifest.cases.map(({ id }) => ({ id, observe: async () => ({ expected: snapshot, actual: snapshot }) }));
    `;
    await writeFile(suitePath, source);
    const result = spawnSync(process.execPath, ["scripts/run-accuracy-suite.mjs", suitePath, "--output", outputPath], {
      cwd: root,
      encoding: "utf8",
    });

    assert.equal(result.status, 1, result.stderr || result.stdout);
    const report = JSON.parse(await readFile(resolve(outputPath, "minimal-case.json"), "utf8"));
    assert.equal(report.passed, false);
    assert.equal(report.overallScore, 0);
    assert.equal(report.coverageValidation.observed.includes("structure-units"), true);
    assert.equal(report.coverageValidation.missing.includes("visual-exact"), true);
    assert.equal(report.failureReasons.some(({ code }) => code === "DECLARED_COVERAGE_NOT_OBSERVED"), true);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("accuracy suite runner writes one AccuracyReport per file and a summary", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-accuracy-"));
  try {
    const suitePath = resolve(temporary, "suite.mjs");
    const outputPath = resolve(temporary, "reports");
    const cases = manifestCases();
    const source = `
      ${comprehensiveSuitePrelude(cases)}
      export const cases = manifest.cases.map(({ id }) => ({ id, observe: async () => observation }));
    `;
    await writeFile(suitePath, source);
    const result = spawnSync(process.execPath, ["scripts/run-accuracy-suite.mjs", suitePath, "--output", outputPath], {
      cwd: root,
      encoding: "utf8",
    });
    assert.equal(result.status, 0, result.stderr || result.stdout);
    const files = (await readdir(outputPath)).sort();
    assert.equal(files.length, corpusClasses.length + 1);
    assert.equal(files.includes("summary.json"), true);
    const report = JSON.parse(await readFile(resolve(outputPath, "minimal-case.json"), "utf8"));
    assert.equal(report.fileId, "minimal-case");
    assert.equal(report.contentCompleteness.score, 100);
    assert.equal(report.geometryAccuracy.score, 100);
    assert.equal(report.textLayoutAccuracy.score, 100);
    assert.equal(report.visualSimilarity.score, 100);
    assert.equal(report.objectMappingAccuracy.score, 100);
    assert.equal(report.overallScore, 100);
    assert.deepEqual(report.failureReasons, []);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("accuracy suite runner still writes a failed report when observation throws", async () => {
  const temporary = await mkdtemp(resolve(tmpdir(), "officeviewer-accuracy-failure-"));
  try {
    const suitePath = resolve(temporary, "suite.mjs");
    const outputPath = resolve(temporary, "reports");
    const source = `
      ${comprehensiveSuitePrelude(manifestCases())}
      export const cases = manifest.cases.map(({ id }) => ({
        id,
        observe: async () => {
          if (id === "large-case") throw new Error("reference renderer unavailable");
          return observation;
        },
      }));
    `;
    await writeFile(suitePath, source);
    const result = spawnSync(process.execPath, ["scripts/run-accuracy-suite.mjs", suitePath, "--output", outputPath], {
      cwd: root,
      encoding: "utf8",
    });
    assert.equal(result.status, 1, result.stderr || result.stdout);
    const report = JSON.parse(await readFile(resolve(outputPath, "large-case.json"), "utf8"));
    assert.equal(report.passed, false);
    assert.equal(report.overallScore, 0);
    assert.equal(report.failureReasons[0].code, "OBSERVATION_FAILED");
    const summary = JSON.parse(await readFile(resolve(outputPath, "summary.json"), "utf8"));
    assert.equal(summary.failed, 1);
    assert.equal(summary.total, corpusClasses.length);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});
