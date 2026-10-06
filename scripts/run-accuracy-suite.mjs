import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { pathToFileURL } from "node:url";

import { evaluateAccuracy, validateAccuracyCorpus } from "../dist/accuracy.js";

function usage() {
  return "Usage: node scripts/run-accuracy-suite.mjs <suite-module.mjs> [--output <directory>]";
}

function parseArguments(arguments_) {
  let modulePath;
  let output = "output/accuracy";
  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index];
    if (argument === "--output") {
      const value = arguments_[++index];
      if (value === undefined) throw new Error(`${usage()}\nMissing --output value`);
      output = value;
    } else if (argument.startsWith("-")) {
      throw new Error(`${usage()}\nUnknown option ${argument}`);
    } else if (modulePath === undefined) {
      modulePath = argument;
    } else {
      throw new Error(`${usage()}\nUnexpected argument ${argument}`);
    }
  }
  if (modulePath === undefined) throw new Error(usage());
  return { modulePath: resolve(modulePath), output: resolve(output) };
}

function safeName(id) {
  const value = id.normalize("NFKC").replace(/[^A-Za-z0-9._-]+/gu, "-").replace(/^-+|-+$/gu, "");
  if (value === "" || value === "." || value === "..") throw new Error(`File id ${JSON.stringify(id)} cannot form a report filename`);
  return value;
}

function assertSuite(module) {
  if (module.manifest === undefined || !Array.isArray(module.cases)) {
    throw new Error("Accuracy suite must export manifest and cases");
  }
  const validation = validateAccuracyCorpus(module.manifest);
  if (!validation.valid) {
    const detail = validation.issues.map((issue) => `${issue.code}: ${issue.message}`).join("\n");
    throw new Error(`Accuracy corpus manifest is invalid:\n${detail}`);
  }
  const manifestIds = new Set(module.manifest.cases.map(({ id }) => id));
  const caseIds = new Set();
  const reportNames = new Map();
  for (const testCase of module.cases) {
    if (typeof testCase?.id !== "string" || typeof testCase.observe !== "function") {
      throw new Error("Every accuracy suite case must contain a string id and an observe() function");
    }
    if (!manifestIds.has(testCase.id)) throw new Error(`Suite case ${testCase.id} is absent from the manifest`);
    if (caseIds.has(testCase.id)) throw new Error(`Suite case ${testCase.id} is duplicated`);
    caseIds.add(testCase.id);
    const reportName = safeName(testCase.id);
    const collision = reportNames.get(reportName);
    if (collision !== undefined) throw new Error(`Suite cases ${collision} and ${testCase.id} map to the same report filename`);
    reportNames.set(reportName, testCase.id);
  }
  for (const id of manifestIds) if (!caseIds.has(id)) throw new Error(`Manifest case ${id} has no observe() adapter`);
}

function observationFailureReport(fileId, corpusClass, declaredCoverage, cause) {
  const failedSection = { score: 0, applicable: true, metrics: {} };
  return {
    fileId,
    corpusClass,
    passed: false,
    coverageValidation: { declared: declaredCoverage, observed: [], missing: declaredCoverage },
    contentCompleteness: failedSection,
    geometryAccuracy: failedSection,
    textLayoutAccuracy: failedSection,
    visualSimilarity: failedSection,
    objectMappingAccuracy: failedSection,
    compatibilityAccuracy: failedSection,
    determinismAccuracy: failedSection,
    metamorphicAccuracy: failedSection,
    overallScore: 0,
    failureReasons: [{
      code: "OBSERVATION_FAILED",
      layer: "overall",
      message: cause instanceof Error ? cause.message : String(cause),
    }],
  };
}

const { modulePath, output } = parseArguments(process.argv.slice(2));
const suite = await import(pathToFileURL(modulePath).href);
assertSuite(suite);
await mkdir(output, { recursive: true });

const manifestById = new Map(suite.manifest.cases.map((testCase) => [testCase.id, testCase]));
const reports = [];
for (const testCase of suite.cases) {
  const manifestCase = manifestById.get(testCase.id);
  let report;
  try {
    const observation = await testCase.observe();
    report = evaluateAccuracy({
      fileId: testCase.id,
      corpusClass: manifestCase.corpusClass,
      expected: observation.expected,
      actual: observation.actual,
      ...(observation.repeatedRuns === undefined ? {} : { repeatedRuns: observation.repeatedRuns }),
      ...(observation.metamorphicVariants === undefined ? {} : { metamorphicVariants: observation.metamorphicVariants }),
      declaredCoverage: manifestCase.coverage,
      ...(observation.policy === undefined ? {} : { policy: observation.policy }),
    });
  } catch (cause) {
    report = observationFailureReport(testCase.id, manifestCase.corpusClass, manifestCase.coverage, cause);
  }
  reports.push(report);
  const reportPath = resolve(output, `${safeName(testCase.id)}.json`);
  await mkdir(dirname(reportPath), { recursive: true });
  await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`, { mode: 0o600 });
  console.log(`${report.passed ? "PASS" : "FAIL"} ${testCase.id} score=${report.overallScore} report=${reportPath}`);
}

const summary = {
  generatedAt: new Date().toISOString(),
  suite: modulePath,
  total: reports.length,
  passed: reports.filter(({ passed }) => passed).length,
  failed: reports.filter(({ passed }) => !passed).length,
  averageScore: reports.length === 0
    ? 100
    : Math.round((reports.reduce((sum, report) => sum + report.overallScore, 0) / reports.length) * 10_000) / 10_000,
  reports: reports.map(({ fileId, corpusClass, passed, overallScore }) => ({ fileId, corpusClass, passed, overallScore })),
};
const summaryPath = resolve(output, "summary.json");
await writeFile(summaryPath, `${JSON.stringify(summary, null, 2)}\n`, { mode: 0o600 });
console.log(`Summary: ${summary.passed}/${summary.total} passed, average=${summary.averageScore}, report=${summaryPath}`);
if (summary.failed !== 0) process.exitCode = 1;
