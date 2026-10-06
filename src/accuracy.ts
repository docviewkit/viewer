/** Pure, dependency-free accuracy evaluation for renderer test harnesses. */

import { bytesEqual } from "./bytes.js";

export type AccuracyCorpusClass =
  | "minimal"
  | "combination"
  | "enterprise"
  | "large"
  | "malformed"
  | "malicious";

export type HitProbeKind = "center" | "edge" | "rotated" | "group" | "overlap" | "z-order";
export type MetamorphicVariantKind =
  | "xml-attribute-order"
  | "namespace-prefix"
  | "zip-entry-order"
  | "irrelevant-metadata";

export interface AccuracyRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface AccuracyUnitObservation {
  readonly index: number;
  readonly type: "slide" | "sheet" | "page";
  readonly width: number;
  readonly height: number;
}

export interface TextLineObservation {
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly baseline: number;
  readonly fontSize: number;
  readonly lineHeight: number;
}

export interface TextLayoutObservation {
  readonly lines: readonly TextLineObservation[];
  readonly overflow: boolean;
  readonly autoScale: number;
}

export interface AccuracyObjectObservation {
  /** Renderer-local identity. Determinism checks require this to stay stable. */
  readonly id: string;
  /** Stable native-source identity used to match expected and actual objects. */
  readonly sourceKey: string;
  readonly type: string;
  readonly unitIndex: number;
  readonly bounds: AccuracyRect;
  readonly rotation?: number;
  /** Image crop in normalized [0, 1] coordinates. */
  readonly crop?: AccuracyRect;
  readonly zOrder?: number;
  readonly parentSourceKey?: string;
  /** Text, image digest, table shape, cell address/value, or another source oracle. */
  readonly content?: Readonly<Record<string, unknown>>;
  readonly textLayout?: TextLayoutObservation;
}

export interface PixelSurfaceObservation {
  readonly unitIndex: number;
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
  /** Document-space area represented by this surface; defaults to the full unit. */
  readonly viewport?: AccuracyRect;
}

export interface HitProbeObservation {
  readonly id: string;
  readonly kind: HitProbeKind;
  readonly unitIndex: number;
  /** Normalized unit-space coordinates. */
  readonly x: number;
  readonly y: number;
  /** Native source keys in front-to-back hit order. */
  readonly hits: readonly string[];
}

export interface AccuracyDiagnosticObservation {
  readonly code: string;
  readonly fidelity: "exact" | "approximate" | "unsupported" | "not-rendered";
  readonly phase: string;
  readonly part?: string;
  readonly sourceKey?: string;
}

export interface AccuracySnapshot {
  readonly units: readonly AccuracyUnitObservation[];
  readonly objects: readonly AccuracyObjectObservation[];
  readonly visuals?: readonly PixelSurfaceObservation[];
  readonly hitProbes?: readonly HitProbeObservation[];
  readonly diagnostics?: readonly AccuracyDiagnosticObservation[];
  readonly displayList?: unknown;
}

export interface MetamorphicVariantObservation {
  readonly kind: MetamorphicVariantKind;
  readonly snapshot: AccuracySnapshot;
}

export interface AccuracyTolerancePolicy {
  /** Maximum normalized difference for x/y/width/height and crop values. */
  readonly geometry?: number;
  readonly rotationDegrees?: number;
  /** Maximum relative difference for font size, line height, and auto scale. */
  readonly textRelative?: number;
  /** Maximum per-channel RGBA delta for the tolerant pixel metric. */
  readonly pixelChannel?: number;
  /** Maximum pixel displacement searched by the tolerant metric. */
  readonly pixelRadius?: number;
  readonly minExactPixelRatio?: number;
  readonly minTolerantPixelRatio?: number;
  readonly minSsim?: number;
  readonly minObjectRegionSimilarity?: number;
  readonly minOverallScore?: number;
}

export interface EvaluateAccuracyRequest {
  readonly fileId: string;
  readonly corpusClass: AccuracyCorpusClass;
  readonly expected: AccuracySnapshot;
  readonly actual: AccuracySnapshot;
  readonly repeatedRuns?: readonly AccuracySnapshot[];
  readonly metamorphicVariants?: readonly MetamorphicVariantObservation[];
  /** Coverage claims to verify against independent oracle observations. */
  readonly declaredCoverage?: readonly AccuracyCoverage[];
  readonly policy?: AccuracyTolerancePolicy;
}

export type AccuracyLayer =
  | "content"
  | "geometry"
  | "text-layout"
  | "visual"
  | "object-mapping"
  | "compatibility"
  | "determinism"
  | "metamorphic"
  | "overall";

export interface AccuracyFailure {
  readonly code: string;
  readonly layer: AccuracyLayer;
  readonly message: string;
  readonly sourceKey?: string;
  readonly coverage?: AccuracyCoverage;
  readonly expected?: unknown;
  readonly actual?: unknown;
}

export interface AccuracySectionReport {
  readonly score: number;
  readonly applicable: boolean;
  readonly metrics: Readonly<Record<string, number>>;
}

export interface AccuracyCoverageValidation {
  readonly declared: readonly AccuracyCoverage[];
  readonly observed: readonly AccuracyCoverage[];
  readonly missing: readonly AccuracyCoverage[];
}

export interface AccuracyReport {
  readonly fileId: string;
  readonly corpusClass: AccuracyCorpusClass;
  readonly passed: boolean;
  readonly coverageValidation: AccuracyCoverageValidation;
  readonly contentCompleteness: AccuracySectionReport;
  readonly geometryAccuracy: AccuracySectionReport;
  readonly textLayoutAccuracy: AccuracySectionReport;
  readonly visualSimilarity: AccuracySectionReport;
  readonly objectMappingAccuracy: AccuracySectionReport;
  readonly compatibilityAccuracy: AccuracySectionReport;
  readonly determinismAccuracy: AccuracySectionReport;
  readonly metamorphicAccuracy: AccuracySectionReport;
  readonly overallScore: number;
  readonly failureReasons: readonly AccuracyFailure[];
}

export type AccuracyCoverage =
  | "structure-units"
  | "structure-objects"
  | "content-text"
  | "content-image"
  | "content-table"
  | "content-cell"
  | "geometry-bounds"
  | "geometry-rotation"
  | "geometry-crop"
  | "geometry-z-order"
  | "text-lines"
  | "text-breaks"
  | "text-baseline"
  | "text-font-size"
  | "text-line-spacing"
  | "text-overflow"
  | "text-auto-scale"
  | "visual-exact"
  | "visual-tolerant"
  | "visual-ssim"
  | "visual-object-region"
  | "hit-center"
  | "hit-edge"
  | "hit-rotated"
  | "hit-group"
  | "hit-overlap"
  | "hit-z-order"
  | "diagnostic-unsupported"
  | "diagnostic-approximate"
  | "diagnostic-font-substitution"
  | "diagnostic-external-resource"
  | "diagnostic-active-content"
  | "determinism-display-list"
  | "determinism-object-id"
  | "determinism-layout"
  | "metamorphic-xml-attribute-order"
  | "metamorphic-namespace-prefix"
  | "metamorphic-zip-entry-order"
  | "metamorphic-irrelevant-metadata";

export const REQUIRED_ACCURACY_COVERAGE: readonly AccuracyCoverage[] = Object.freeze([
  "structure-units",
  "structure-objects",
  "content-text",
  "content-image",
  "content-table",
  "content-cell",
  "geometry-bounds",
  "geometry-rotation",
  "geometry-crop",
  "geometry-z-order",
  "text-lines",
  "text-breaks",
  "text-baseline",
  "text-font-size",
  "text-line-spacing",
  "text-overflow",
  "text-auto-scale",
  "visual-exact",
  "visual-tolerant",
  "visual-ssim",
  "visual-object-region",
  "hit-center",
  "hit-edge",
  "hit-rotated",
  "hit-group",
  "hit-overlap",
  "hit-z-order",
  "diagnostic-unsupported",
  "diagnostic-approximate",
  "diagnostic-font-substitution",
  "diagnostic-external-resource",
  "diagnostic-active-content",
  "determinism-display-list",
  "determinism-object-id",
  "determinism-layout",
  "metamorphic-xml-attribute-order",
  "metamorphic-namespace-prefix",
  "metamorphic-zip-entry-order",
  "metamorphic-irrelevant-metadata",
]);

export interface AccuracyCorpusCase {
  readonly id: string;
  readonly corpusClass: AccuracyCorpusClass;
  readonly format: "pptx" | "ppt" | "odp" | "xlsx" | "xls" | "ods" | "docx" | "doc" | "odt"
    | "keynote" | "pages" | "numbers" | "pdf" | "mixed";
  readonly fixture: string;
  readonly oracle: string;
  readonly automatedTest: string;
  readonly coverage: readonly AccuracyCoverage[];
  readonly regressionBugId?: string;
}

export interface AccuracyCorpusManifest {
  readonly compatibilityBugs: readonly string[];
  readonly cases: readonly AccuracyCorpusCase[];
}

export interface AccuracyCorpusIssue {
  readonly code: string;
  readonly message: string;
  readonly caseId?: string;
  readonly bugId?: string;
  readonly coverage?: AccuracyCoverage;
  readonly corpusClass?: AccuracyCorpusClass;
}

export interface AccuracyCorpusValidation {
  readonly valid: boolean;
  readonly issues: readonly AccuracyCorpusIssue[];
}

/** Validate corpus breadth and the compatibility-bug regression invariant. */
export function validateAccuracyCorpus(manifest: AccuracyCorpusManifest): AccuracyCorpusValidation {
  const issues: AccuracyCorpusIssue[] = [];
  const corpusClasses: readonly AccuracyCorpusClass[] = [
    "minimal",
    "combination",
    "enterprise",
    "large",
    "malformed",
    "malicious",
  ];
  const corpusClassSet = new Set<string>(corpusClasses);
  const formatSet = new Set<string>([
    "pptx", "ppt", "odp", "xlsx", "xls", "ods", "docx", "doc", "odt", "keynote", "pages", "numbers", "pdf", "mixed",
  ]);
  const requiredCoverageSet = new Set<string>(REQUIRED_ACCURACY_COVERAGE);
  const seenIds = new Set<string>();
  for (const testCase of manifest.cases) {
    if (seenIds.has(testCase.id)) {
      issues.push({ code: "CORPUS_CASE_DUPLICATE", message: `Corpus case ${testCase.id} is duplicated`, caseId: testCase.id });
    }
    seenIds.add(testCase.id);
    if (!corpusClassSet.has(testCase.corpusClass)) {
      issues.push({ code: "CORPUS_CLASS_INVALID", message: `Corpus case ${testCase.id} has an invalid class`, caseId: testCase.id });
    }
    if (!formatSet.has(testCase.format)) {
      issues.push({ code: "CORPUS_FORMAT_INVALID", message: `Corpus case ${testCase.id} has an invalid format`, caseId: testCase.id });
    }
    if (testCase.fixture.trim() === "" || testCase.oracle.trim() === "" || testCase.automatedTest.trim() === "") {
      issues.push({ code: "CORPUS_CASE_INCOMPLETE", message: `Corpus case ${testCase.id} must reference a fixture, oracle, and automated test`, caseId: testCase.id });
    }
    for (const coverage of testCase.coverage) {
      if (!requiredCoverageSet.has(coverage)) {
        issues.push({ code: "ACCURACY_COVERAGE_INVALID", message: `Corpus case ${testCase.id} declares unknown coverage ${coverage}`, caseId: testCase.id });
      }
    }
  }
  for (const corpusClass of corpusClasses) {
    if (!manifest.cases.some((testCase) => testCase.corpusClass === corpusClass)) {
      issues.push({ code: "CORPUS_CLASS_MISSING", message: `Corpus class ${corpusClass} has no case`, corpusClass });
    }
  }
  const covered = new Set(manifest.cases.flatMap(({ coverage }) => coverage));
  for (const coverage of REQUIRED_ACCURACY_COVERAGE) {
    if (!covered.has(coverage)) {
      issues.push({ code: "ACCURACY_COVERAGE_MISSING", message: `Accuracy coverage ${coverage} is missing`, coverage });
    }
  }
  const seenBugs = new Set<string>();
  for (const bugId of manifest.compatibilityBugs) {
    if (seenBugs.has(bugId)) {
      issues.push({ code: "COMPATIBILITY_BUG_DUPLICATE", message: `Compatibility bug ${bugId} is duplicated`, bugId });
      continue;
    }
    seenBugs.add(bugId);
    const regressions = manifest.cases.filter((testCase) => testCase.regressionBugId === bugId);
    if (regressions.length === 0) {
      issues.push({ code: "REGRESSION_CASE_MISSING", message: `Compatibility bug ${bugId} has no regression case`, bugId });
      continue;
    }
    for (const regression of regressions) {
      if (regression.corpusClass !== "minimal") {
        issues.push({ code: "REGRESSION_NOT_MINIMAL", message: `Compatibility bug ${bugId} must use a minimal reproduction`, bugId, caseId: regression.id });
      }
      if (regression.fixture.trim() === "") {
        issues.push({ code: "REGRESSION_FIXTURE_MISSING", message: `Compatibility bug ${bugId} has no reproduction file`, bugId, caseId: regression.id });
      }
      if (regression.oracle.trim() === "") {
        issues.push({ code: "REGRESSION_ORACLE_MISSING", message: `Compatibility bug ${bugId} has no independent oracle`, bugId, caseId: regression.id });
      }
      if (regression.automatedTest.trim() === "") {
        issues.push({ code: "REGRESSION_AUTOMATION_MISSING", message: `Compatibility bug ${bugId} is not referenced by an automated test`, bugId, caseId: regression.id });
      }
    }
  }
  for (const regression of manifest.cases.filter(({ regressionBugId }) => regressionBugId !== undefined)) {
    const bugId = regression.regressionBugId;
    if (bugId !== undefined && !seenBugs.has(bugId)) {
      issues.push({
        code: "REGRESSION_BUG_UNREGISTERED",
        message: `Regression case ${regression.id} references an unregistered compatibility bug`,
        caseId: regression.id,
        bugId,
      });
    }
  }
  return Object.freeze({ valid: issues.length === 0, issues: Object.freeze(issues) });
}

const DEFAULT_POLICY = Object.freeze({
  geometry: 0.002,
  rotationDegrees: 0.25,
  textRelative: 0.01,
  pixelChannel: 8,
  pixelRadius: 0,
  minExactPixelRatio: 1,
  minTolerantPixelRatio: 0.995,
  minSsim: 0.99,
  minObjectRegionSimilarity: 0.99,
  minOverallScore: 95,
});

const SECTION_WEIGHTS = Object.freeze({
  contentCompleteness: 18,
  geometryAccuracy: 18,
  textLayoutAccuracy: 15,
  visualSimilarity: 18,
  objectMappingAccuracy: 13,
  compatibilityAccuracy: 8,
  determinismAccuracy: 5,
  metamorphicAccuracy: 5,
});

interface Counter {
  passed: number;
  total: number;
}

function check(counter: Counter, condition: boolean): boolean {
  counter.total += 1;
  if (condition) counter.passed += 1;
  return condition;
}

function rounded(value: number): number {
  return Math.round(value * 10_000) / 10_000;
}

function score(counter: Counter): number {
  return counter.total === 0 ? 100 : rounded((counter.passed / counter.total) * 100);
}

function section(counter: Counter, metrics: Record<string, number>, applicable = counter.total !== 0): AccuracySectionReport {
  return Object.freeze({ score: score(counter), applicable, metrics: Object.freeze(metrics) });
}

function canonical(value: unknown, seen = new Set<object>()): string {
  if (value === null || typeof value !== "object") {
    if (typeof value === "number" && !Number.isFinite(value)) return JSON.stringify(String(value));
    return JSON.stringify(value);
  }
  if (seen.has(value)) throw new TypeError("Accuracy observations must not contain cycles");
  seen.add(value);
  let result: string;
  if (value instanceof Uint8Array) {
    result = `[${Array.from(value).join(",")}]`;
  } else if (Array.isArray(value)) {
    result = `[${value.map((entry) => canonical(entry, seen)).join(",")}]`;
  } else {
    const record = value as Record<string, unknown>;
    result = `{${Object.keys(record).sort().map((key) => `${JSON.stringify(key)}:${canonical(record[key], seen)}`).join(",")}}`;
  }
  seen.delete(value);
  return result;
}

function same(left: unknown, right: unknown): boolean {
  return canonical(left) === canonical(right);
}

function bySource(objects: readonly AccuracyObjectObservation[]): Map<string, AccuracyObjectObservation> {
  const result = new Map<string, AccuracyObjectObservation>();
  for (const object of objects) {
    if (!result.has(object.sourceKey)) result.set(object.sourceKey, object);
  }
  return result;
}

function byUnit(units: readonly AccuracyUnitObservation[]): Map<number, AccuracyUnitObservation> {
  return new Map(units.map((unit) => [unit.index, unit]));
}

function addFailure(
  failures: AccuracyFailure[],
  layer: AccuracyLayer,
  code: string,
  message: string,
  values: { sourceKey?: string; coverage?: AccuracyCoverage; expected?: unknown; actual?: unknown } = {},
): void {
  failures.push(Object.freeze({ layer, code, message, ...values }));
}

function compareContent(
  expected: AccuracySnapshot,
  actual: AccuracySnapshot,
  failures: AccuracyFailure[],
): AccuracySectionReport {
  const counter: Counter = { passed: 0, total: 0 };
  check(counter, expected.units.length === actual.units.length);
  if (expected.units.length !== actual.units.length) {
    addFailure(failures, "content", "UNIT_COUNT_MISMATCH", "Page, slide, or sheet count differs", {
      expected: expected.units.length,
      actual: actual.units.length,
    });
  }
  const actualUnits = byUnit(actual.units);
  for (const unit of expected.units) {
    const observed = actualUnits.get(unit.index);
    const matches = observed?.type === unit.type;
    check(counter, matches);
    if (!matches) {
      addFailure(failures, "content", "UNIT_TYPE_MISMATCH", `Unit ${unit.index} type differs`, {
        expected: unit.type,
        actual: observed?.type,
      });
    }
  }

  check(counter, expected.objects.length === actual.objects.length);
  if (expected.objects.length !== actual.objects.length) {
    addFailure(failures, "content", "OBJECT_COUNT_MISMATCH", "Object count differs", {
      expected: expected.objects.length,
      actual: actual.objects.length,
    });
  }
  const actualObjects = bySource(actual.objects);
  const expectedObjects = bySource(expected.objects);
  if (expectedObjects.size !== expected.objects.length) throw new TypeError("Expected accuracy objects must have unique sourceKey values");
  const actualSourcesUnique = actualObjects.size === actual.objects.length;
  check(counter, actualSourcesUnique);
  if (!actualSourcesUnique) {
    addFailure(failures, "content", "DUPLICATE_SOURCE_MAPPING", "Actual output maps more than one object to the same native source key");
  }
  const contentByType: Record<string, Counter> = {};
  for (const object of expected.objects) {
    const typeCounter = contentByType[object.type] ?? { passed: 0, total: 0 };
    contentByType[object.type] = typeCounter;
    const observed = actualObjects.get(object.sourceKey);
    const exists = observed !== undefined;
    check(counter, exists);
    if (!exists) {
      check(counter, false);
      check(counter, false);
      check(typeCounter, false);
      addFailure(failures, "content", "OBJECT_MISSING", "Expected source object is missing", { sourceKey: object.sourceKey });
      continue;
    }
    const typeMatches = object.type === observed.type;
    check(counter, typeMatches);
    if (!typeMatches) {
      addFailure(failures, "content", "OBJECT_TYPE_MISMATCH", "Object type differs", {
        sourceKey: object.sourceKey,
        expected: object.type,
        actual: observed.type,
      });
    }
    const contentMatches = same(object.content, observed.content);
    check(counter, contentMatches);
    check(typeCounter, contentMatches);
    if (!contentMatches) {
      addFailure(failures, "content", "OBJECT_CONTENT_MISMATCH", "Text, image, table, or cell content differs", {
        sourceKey: object.sourceKey,
        expected: object.content,
        actual: observed.content,
      });
    }
  }
  for (const object of actual.objects) {
    if (!expectedObjects.has(object.sourceKey)) {
      addFailure(failures, "content", "UNEXPECTED_OBJECT", "Actual output contains an unexpected source object", {
        sourceKey: object.sourceKey,
      });
    }
  }
  const metrics: Record<string, number> = {
    unitCount: actual.units.length,
    objectCount: actual.objects.length,
  };
  for (const [type, typeCounter] of Object.entries(contentByType)) {
    metrics[`${type}Completeness`] = score(typeCounter);
  }
  return section(counter, metrics);
}

function finitePositive(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}

function normalizedBounds(object: AccuracyObjectObservation, unit: AccuracyUnitObservation): AccuracyRect | undefined {
  if (!finitePositive(unit.width) || !finitePositive(unit.height)) return undefined;
  const { x, y, width, height } = object.bounds;
  if (![x, y, width, height].every(Number.isFinite)) return undefined;
  return { x: x / unit.width, y: y / unit.height, width: width / unit.width, height: height / unit.height };
}

function within(left: number, right: number, tolerance: number): boolean {
  return Number.isFinite(left) && Number.isFinite(right) && Math.abs(left - right) <= tolerance;
}

function relativeWithin(left: number, right: number, tolerance: number): boolean {
  const denominator = Math.max(Math.abs(left), Math.abs(right), Number.EPSILON);
  return Number.isFinite(left) && Number.isFinite(right) && Math.abs(left - right) / denominator <= tolerance;
}

function rotationDistance(left: number, right: number): number {
  const delta = Math.abs(((left - right) % 360 + 360) % 360);
  return Math.min(delta, 360 - delta);
}

function compareGeometry(
  expected: AccuracySnapshot,
  actual: AccuracySnapshot,
  failures: AccuracyFailure[],
  geometryTolerance: number,
  rotationTolerance: number,
): AccuracySectionReport {
  const counter: Counter = { passed: 0, total: 0 };
  const expectedUnits = byUnit(expected.units);
  const actualUnits = byUnit(actual.units);
  for (const unit of expected.units) {
    const observed = actualUnits.get(unit.index);
    if (observed === undefined) continue;
    const dimensionsMatch = relativeWithin(unit.width, observed.width, geometryTolerance)
      && relativeWithin(unit.height, observed.height, geometryTolerance);
    check(counter, dimensionsMatch);
    if (!dimensionsMatch) {
      addFailure(failures, "geometry", "UNIT_SIZE_MISMATCH", `Unit ${unit.index} dimensions differ`, {
        expected: { width: unit.width, height: unit.height },
        actual: { width: observed.width, height: observed.height },
      });
    }
  }
  const actualObjects = bySource(actual.objects);
  for (const object of expected.objects) {
    const observed = actualObjects.get(object.sourceKey);
    const expectedUnit = expectedUnits.get(object.unitIndex);
    const actualUnit = observed === undefined ? undefined : actualUnits.get(observed.unitIndex);
    if (observed === undefined || expectedUnit === undefined || actualUnit === undefined) continue;
    const expectedBounds = normalizedBounds(object, expectedUnit);
    const actualBounds = normalizedBounds(observed, actualUnit);
    const boundsMatch = expectedBounds !== undefined && actualBounds !== undefined
      && within(expectedBounds.x, actualBounds.x, geometryTolerance)
      && within(expectedBounds.y, actualBounds.y, geometryTolerance)
      && within(expectedBounds.width, actualBounds.width, geometryTolerance)
      && within(expectedBounds.height, actualBounds.height, geometryTolerance);
    check(counter, boundsMatch);
    if (!boundsMatch) {
      addFailure(failures, "geometry", "NORMALIZED_BOUNDS_MISMATCH", "Normalized object position or size exceeds tolerance", {
        sourceKey: object.sourceKey,
        expected: expectedBounds,
        actual: actualBounds,
      });
    }
    const rotationMatches = rotationDistance(object.rotation ?? 0, observed.rotation ?? 0) <= rotationTolerance;
    check(counter, rotationMatches);
    if (!rotationMatches) {
      addFailure(failures, "geometry", "ROTATION_MISMATCH", "Object rotation exceeds tolerance", {
        sourceKey: object.sourceKey,
        expected: object.rotation ?? 0,
        actual: observed.rotation ?? 0,
      });
    }
    const expectedCrop = object.crop ?? { x: 0, y: 0, width: 1, height: 1 };
    const actualCrop = observed.crop ?? { x: 0, y: 0, width: 1, height: 1 };
    const cropMatches = within(expectedCrop.x, actualCrop.x, geometryTolerance)
      && within(expectedCrop.y, actualCrop.y, geometryTolerance)
      && within(expectedCrop.width, actualCrop.width, geometryTolerance)
      && within(expectedCrop.height, actualCrop.height, geometryTolerance);
    check(counter, cropMatches);
    if (!cropMatches) {
      addFailure(failures, "geometry", "CROP_MISMATCH", "Normalized image crop exceeds tolerance", {
        sourceKey: object.sourceKey,
        expected: object.crop,
        actual: observed.crop,
      });
    }
    const zMatches = (object.zOrder ?? 0) === (observed.zOrder ?? 0);
    check(counter, zMatches);
    if (!zMatches) {
      addFailure(failures, "geometry", "Z_ORDER_MISMATCH", "Object z-order differs", {
        sourceKey: object.sourceKey,
        expected: object.zOrder ?? 0,
        actual: observed.zOrder ?? 0,
      });
    }
    const parentMatches = object.parentSourceKey === observed.parentSourceKey;
    check(counter, parentMatches);
    if (!parentMatches) {
      addFailure(failures, "geometry", "GROUP_PARENT_MISMATCH", "Object group parent differs", {
        sourceKey: object.sourceKey,
        expected: object.parentSourceKey,
        actual: observed.parentSourceKey,
      });
    }
  }
  return section(counter, {
    normalizedCoordinateTolerance: geometryTolerance,
    rotationToleranceDegrees: rotationTolerance,
  });
}

function compareTextLayout(
  expected: AccuracySnapshot,
  actual: AccuracySnapshot,
  failures: AccuracyFailure[],
  geometryTolerance: number,
  relativeTolerance: number,
): AccuracySectionReport {
  const counter: Counter = { passed: 0, total: 0 };
  const actualObjects = bySource(actual.objects);
  const expectedUnits = byUnit(expected.units);
  const actualUnits = byUnit(actual.units);
  for (const object of expected.objects) {
    if (object.textLayout === undefined) continue;
    const observed = actualObjects.get(object.sourceKey);
    const layout = observed?.textLayout;
    const exists = observed !== undefined && layout !== undefined;
    check(counter, exists);
    if (!exists || observed === undefined || layout === undefined) {
      addFailure(failures, "text-layout", "TEXT_LAYOUT_MISSING", "Text layout observation is missing", { sourceKey: object.sourceKey });
      continue;
    }
    const expectedLayout = object.textLayout;
    const lineCountMatches = expectedLayout.lines.length === layout.lines.length;
    check(counter, lineCountMatches);
    if (!lineCountMatches) {
      addFailure(failures, "text-layout", "LINE_COUNT_MISMATCH", "Rendered text line count differs", {
        sourceKey: object.sourceKey,
        expected: expectedLayout.lines.length,
        actual: layout.lines.length,
      });
    }
    const expectedUnit = expectedUnits.get(object.unitIndex);
    const actualUnit = actualUnits.get(observed.unitIndex);
    const comparableLines = Math.min(expectedLayout.lines.length, layout.lines.length);
    for (let index = 0; index < comparableLines; index += 1) {
      const expectedLine = expectedLayout.lines[index]!;
      const actualLine = layout.lines[index]!;
      const breakMatches = expectedLine.text === actualLine.text
        && expectedLine.start === actualLine.start
        && expectedLine.end === actualLine.end;
      check(counter, breakMatches);
      if (!breakMatches) {
        addFailure(failures, "text-layout", "LINE_BREAK_MISMATCH", `Text wrapping differs at line ${index}`, {
          sourceKey: object.sourceKey,
          expected: { text: expectedLine.text, start: expectedLine.start, end: expectedLine.end },
          actual: { text: actualLine.text, start: actualLine.start, end: actualLine.end },
        });
      }
      const baselineMatches = expectedUnit !== undefined && actualUnit !== undefined
        && within(expectedLine.baseline / expectedUnit.height, actualLine.baseline / actualUnit.height, geometryTolerance);
      check(counter, baselineMatches);
      if (!baselineMatches) {
        addFailure(failures, "text-layout", "BASELINE_MISMATCH", `Text baseline differs at line ${index}`, {
          sourceKey: object.sourceKey,
          expected: expectedLine.baseline,
          actual: actualLine.baseline,
        });
      }
      const fontMatches = relativeWithin(expectedLine.fontSize, actualLine.fontSize, relativeTolerance);
      check(counter, fontMatches);
      if (!fontMatches) {
        addFailure(failures, "text-layout", "FONT_SIZE_MISMATCH", `Font size differs at line ${index}`, {
          sourceKey: object.sourceKey,
          expected: expectedLine.fontSize,
          actual: actualLine.fontSize,
        });
      }
      const spacingMatches = relativeWithin(expectedLine.lineHeight, actualLine.lineHeight, relativeTolerance);
      check(counter, spacingMatches);
      if (!spacingMatches) {
        addFailure(failures, "text-layout", "LINE_SPACING_MISMATCH", `Line spacing differs at line ${index}`, {
          sourceKey: object.sourceKey,
          expected: expectedLine.lineHeight,
          actual: actualLine.lineHeight,
        });
      }
    }
    const overflowMatches = expectedLayout.overflow === layout.overflow;
    check(counter, overflowMatches);
    if (!overflowMatches) {
      addFailure(failures, "text-layout", "TEXT_OVERFLOW_MISMATCH", "Text-box overflow state differs", {
        sourceKey: object.sourceKey,
        expected: expectedLayout.overflow,
        actual: layout.overflow,
      });
    }
    const scaleMatches = relativeWithin(expectedLayout.autoScale, layout.autoScale, relativeTolerance);
    check(counter, scaleMatches);
    if (!scaleMatches) {
      addFailure(failures, "text-layout", "TEXT_AUTOSCALE_MISMATCH", "Text auto-scale differs", {
        sourceKey: object.sourceKey,
        expected: expectedLayout.autoScale,
        actual: layout.autoScale,
      });
    }
  }
  return section(counter, { relativeTolerance, normalizedBaselineTolerance: geometryTolerance });
}

function validSurface(surface: PixelSurfaceObservation): boolean {
  return Number.isInteger(surface.width) && surface.width > 0
    && Number.isInteger(surface.height) && surface.height > 0
    && surface.data.length === surface.width * surface.height * 4;
}

function pixelRatios(
  expected: Uint8Array,
  actual: Uint8Array,
  width: number,
  height: number,
  channelTolerance: number,
  pixelRadius: number,
): { exact: number; tolerant: number } {
  const pixels = expected.length / 4;
  let exact = 0;
  const matchesNeighborhood = (
    reference: Uint8Array,
    candidate: Uint8Array,
    pixel: number,
  ): boolean => {
    const offset = pixel * 4;
    const x = pixel % width;
    const y = Math.floor(pixel / width);
    for (let candidateY = Math.max(0, y - pixelRadius);
      candidateY <= Math.min(height - 1, y + pixelRadius);
      candidateY += 1) {
      for (let candidateX = Math.max(0, x - pixelRadius);
        candidateX <= Math.min(width - 1, x + pixelRadius);
        candidateX += 1) {
        const candidateOffset = (candidateY * width + candidateX) * 4;
        let matches = true;
        for (let channel = 0; channel < 4; channel += 1) {
          if (Math.abs(reference[offset + channel]! - candidate[candidateOffset + channel]!) > channelTolerance) {
            matches = false;
            break;
          }
        }
        if (matches) return true;
      }
    }
    return false;
  };
  let forwardTolerant = 0;
  for (let pixel = 0; pixel < pixels; pixel += 1) {
    const offset = pixel * 4;
    let exactPixel = true;
    for (let channel = 0; channel < 4; channel += 1) {
      const delta = Math.abs(expected[offset + channel]! - actual[offset + channel]!);
      if (delta !== 0) exactPixel = false;
    }
    if (exactPixel) exact += 1;
    if (matchesNeighborhood(expected, actual, pixel)) forwardTolerant += 1;
  }
  let reverseTolerant = forwardTolerant;
  if (pixelRadius > 0) {
    reverseTolerant = 0;
    for (let pixel = 0; pixel < pixels; pixel += 1) {
      if (matchesNeighborhood(actual, expected, pixel)) reverseTolerant += 1;
    }
  }
  return { exact: exact / pixels, tolerant: Math.min(forwardTolerant, reverseTolerant) / pixels };
}

function luminance(data: Uint8Array, pixel: number): number {
  const offset = pixel * 4;
  return data[offset]! * 0.2126 + data[offset + 1]! * 0.7152 + data[offset + 2]! * 0.0722;
}

/** Windowed luminance SSIM with 8×8 windows. */
function ssim(expected: PixelSurfaceObservation, actual: PixelSurfaceObservation): number {
  const windowSize = 8;
  const c1 = (0.01 * 255) ** 2;
  const c2 = (0.03 * 255) ** 2;
  let total = 0;
  let windows = 0;
  for (let top = 0; top < expected.height; top += windowSize) {
    for (let left = 0; left < expected.width; left += windowSize) {
      const width = Math.min(windowSize, expected.width - left);
      const height = Math.min(windowSize, expected.height - top);
      const count = width * height;
      let meanExpected = 0;
      let meanActual = 0;
      for (let y = 0; y < height; y += 1) {
        for (let x = 0; x < width; x += 1) {
          const pixel = (top + y) * expected.width + left + x;
          meanExpected += luminance(expected.data, pixel);
          meanActual += luminance(actual.data, pixel);
        }
      }
      meanExpected /= count;
      meanActual /= count;
      let varianceExpected = 0;
      let varianceActual = 0;
      let covariance = 0;
      for (let y = 0; y < height; y += 1) {
        for (let x = 0; x < width; x += 1) {
          const pixel = (top + y) * expected.width + left + x;
          const deltaExpected = luminance(expected.data, pixel) - meanExpected;
          const deltaActual = luminance(actual.data, pixel) - meanActual;
          varianceExpected += deltaExpected * deltaExpected;
          varianceActual += deltaActual * deltaActual;
          covariance += deltaExpected * deltaActual;
        }
      }
      const denominator = Math.max(1, count - 1);
      varianceExpected /= denominator;
      varianceActual /= denominator;
      covariance /= denominator;
      total += ((2 * meanExpected * meanActual + c1) * (2 * covariance + c2))
        / ((meanExpected ** 2 + meanActual ** 2 + c1) * (varianceExpected + varianceActual + c2));
      windows += 1;
    }
  }
  return Math.max(-1, Math.min(1, total / windows));
}

interface PixelRegion {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

const MAX_OBJECT_REGION_COUNT = 256;
const MAX_OBJECT_REGION_PIXELS = 250_000;
const MAX_OBJECT_REGION_NEIGHBOR_WORK = 128_000_000;

function regionPixelBounds(
  object: AccuracyObjectObservation,
  unit: AccuracyUnitObservation,
  surface: PixelSurfaceObservation,
): PixelRegion | undefined {
  const viewport = surface.viewport ?? { x: 0, y: 0, width: unit.width, height: unit.height };
  if (![viewport.x, viewport.y, viewport.width, viewport.height].every(Number.isFinite)
    || viewport.width <= 0
    || viewport.height <= 0
    || ![
      object.bounds.x,
      object.bounds.y,
      object.bounds.width,
      object.bounds.height,
    ].every(Number.isFinite)
    || object.bounds.width <= 0
    || object.bounds.height <= 0) return undefined;
  const left = Math.max(0, Math.floor(((object.bounds.x - viewport.x) / viewport.width) * surface.width));
  const top = Math.max(0, Math.floor(((object.bounds.y - viewport.y) / viewport.height) * surface.height));
  const right = Math.min(surface.width, Math.ceil(((object.bounds.x + object.bounds.width - viewport.x) / viewport.width) * surface.width));
  const bottom = Math.min(surface.height, Math.ceil(((object.bounds.y + object.bounds.height - viewport.y) / viewport.height) * surface.height));
  return right > left && bottom > top ? { left, top, right, bottom } : undefined;
}

function pixelDiffers(
  left: Uint8Array,
  leftOffset: number,
  right: Uint8Array,
  rightOffset: number,
  tolerance: number,
): boolean {
  for (let channel = 0; channel < 4; channel += 1) {
    if (Math.abs(left[leftOffset + channel]! - right[rightOffset + channel]!) > tolerance) {
      return true;
    }
  }
  return false;
}

function regionAttentionMask(
  source: Uint8Array,
  other: Uint8Array,
  width: number,
  height: number,
  region: PixelRegion,
  tolerance: number,
): Uint8Array {
  const regionWidth = region.right - region.left;
  const regionHeight = region.bottom - region.top;
  const mask = new Uint8Array(regionWidth * regionHeight);
  const edgeTolerance = Math.max(8, tolerance);
  for (let y = region.top; y < region.bottom; y += 1) {
    for (let x = region.left; x < region.right; x += 1) {
      const pixel = y * width + x;
      const offset = pixel * 4;
      let attention = pixelDiffers(source, offset, other, offset, tolerance);
      if (!attention) {
        attention = x > 0 && pixelDiffers(
          source,
          offset,
          source,
          (pixel - 1) * 4,
          edgeTolerance,
        ) || x + 1 < width && pixelDiffers(
          source,
          offset,
          source,
          (pixel + 1) * 4,
          edgeTolerance,
        ) || y > 0 && pixelDiffers(
          source,
          offset,
          source,
          (pixel - width) * 4,
          edgeTolerance,
        ) || y + 1 < height && pixelDiffers(
          source,
          offset,
          source,
          (pixel + width) * 4,
          edgeTolerance,
        );
      }
      if (attention) {
        mask[(y - region.top) * regionWidth + x - region.left] = 1;
      }
    }
  }
  return mask;
}

function maskContains(mask: Uint8Array, region: PixelRegion, x: number, y: number): boolean {
  if (x < region.left || x >= region.right || y < region.top || y >= region.bottom) return false;
  return mask[(y - region.top) * (region.right - region.left) + x - region.left] === 1;
}

function regionEdgeEnergy(
  data: Uint8Array,
  width: number,
  region: PixelRegion,
  tolerance: number,
): number {
  const noiseFloor = Math.max(8, tolerance);
  const edgeEnergy = (leftPixel: number, rightPixel: number): number => {
    const leftOffset = leftPixel * 4;
    const rightOffset = rightPixel * 4;
    let gradient = 0;
    for (let channel = 0; channel < 3; channel += 1) {
      gradient = Math.max(
        gradient,
        Math.abs(data[leftOffset + channel]! - data[rightOffset + channel]!),
      );
    }
    return Math.max(0, gradient - noiseFloor);
  };
  let energy = 0;
  for (let y = region.top; y < region.bottom; y += 1) {
    for (let x = region.left; x < region.right; x += 1) {
      const pixel = y * width + x;
      if (x + 1 < region.right) energy += edgeEnergy(pixel, pixel + 1);
      if (y + 1 < region.bottom) energy += edgeEnergy(pixel, pixel + width);
    }
  }
  return energy;
}

interface DirectionalRegionSimilarity {
  readonly all: number;
  readonly attention: number;
  readonly attentionPixels: number;
}

function directionalRegionSimilarity(
  reference: Uint8Array,
  candidate: Uint8Array,
  referenceAttention: Uint8Array,
  candidateAttention: Uint8Array,
  region: PixelRegion,
  searchRegion: PixelRegion,
  width: number,
  tolerance: number,
  radius: number,
): DirectionalRegionSimilarity {
  const pixels = (region.right - region.left) * (region.bottom - region.top);
  let allMatches = 0;
  let attentionMatches = 0;
  let attentionPixels = 0;
  for (let y = region.top; y < region.bottom; y += 1) {
    for (let x = region.left; x < region.right; x += 1) {
      const referenceOffset = (y * width + x) * 4;
      const isAttention = maskContains(referenceAttention, searchRegion, x, y);
      if (isAttention) attentionPixels += 1;
      let matches = false;
      let attentionMatchesCandidate = false;
      for (let candidateY = Math.max(searchRegion.top, y - radius);
        candidateY <= Math.min(searchRegion.bottom - 1, y + radius);
        candidateY += 1) {
        for (let candidateX = Math.max(searchRegion.left, x - radius);
          candidateX <= Math.min(searchRegion.right - 1, x + radius);
          candidateX += 1) {
          const candidateOffset = (candidateY * width + candidateX) * 4;
          if (pixelDiffers(reference, referenceOffset, candidate, candidateOffset, tolerance)) {
            continue;
          }
          matches = true;
          if (!isAttention || maskContains(
            candidateAttention,
            searchRegion,
            candidateX,
            candidateY,
          )) {
            attentionMatchesCandidate = true;
          }
          if (attentionMatchesCandidate) break;
        }
        if (attentionMatchesCandidate) break;
      }
      if (matches) allMatches += 1;
      if (isAttention && attentionMatchesCandidate) attentionMatches += 1;
    }
  }
  return {
    all: allMatches / pixels,
    attention: attentionPixels === 0 ? 1 : attentionMatches / attentionPixels,
    attentionPixels,
  };
}

function regionSimilarity(
  expected: Uint8Array,
  actual: Uint8Array,
  region: PixelRegion,
  width: number,
  height: number,
  tolerance: number,
  radius: number,
): number {
  const searchRegion = {
    left: Math.max(0, region.left - radius),
    top: Math.max(0, region.top - radius),
    right: Math.min(width, region.right + radius),
    bottom: Math.min(height, region.bottom + radius),
  };
  const expectedAttention = regionAttentionMask(
    expected,
    actual,
    width,
    height,
    searchRegion,
    tolerance,
  );
  const actualAttention = regionAttentionMask(
    actual,
    expected,
    width,
    height,
    searchRegion,
    tolerance,
  );
  const forward = directionalRegionSimilarity(
    expected,
    actual,
    expectedAttention,
    actualAttention,
    region,
    searchRegion,
    width,
    tolerance,
    radius,
  );
  const reverse = directionalRegionSimilarity(
    actual,
    expected,
    actualAttention,
    expectedAttention,
    region,
    searchRegion,
    width,
    tolerance,
    radius,
  );
  const attentionSimilarity = forward.attentionPixels === 0 && reverse.attentionPixels === 0
    ? 1
    : forward.attentionPixels === 0 || reverse.attentionPixels === 0
      ? 0
      : Math.min(forward.attention, reverse.attention);
  const expectedEdgeEnergy = regionEdgeEnergy(expected, width, searchRegion, tolerance);
  const actualEdgeEnergy = regionEdgeEnergy(actual, width, searchRegion, tolerance);
  const edgeEnergySimilarity = expectedEdgeEnergy === 0 && actualEdgeEnergy === 0
    ? 1
    : expectedEdgeEnergy === 0 || actualEdgeEnergy === 0
      ? 0
      : Math.min(expectedEdgeEnergy, actualEdgeEnergy)
        / Math.max(expectedEdgeEnergy, actualEdgeEnergy);
  return Math.min(forward.all, reverse.all, attentionSimilarity, edgeEnergySimilarity);
}

function compareVisuals(
  expected: AccuracySnapshot,
  actual: AccuracySnapshot,
  failures: AccuracyFailure[],
  policy: Required<AccuracyTolerancePolicy>,
): AccuracySectionReport {
  const expectedVisuals = expected.visuals ?? [];
  const actualVisuals = new Map((actual.visuals ?? []).map((surface) => [surface.unitIndex, surface]));
  const counter: Counter = { passed: 0, total: 0 };
  const units = byUnit(expected.units);
  const exactRatios: number[] = [];
  const tolerantRatios: number[] = [];
  const ssimValues: number[] = [];
  const regionValues: number[] = [];
  for (const surface of expectedVisuals) {
    const observed = actualVisuals.get(surface.unitIndex);
    const compatible = observed !== undefined && validSurface(surface) && validSurface(observed)
      && surface.width === observed.width && surface.height === observed.height;
    check(counter, compatible);
    if (!compatible) {
      addFailure(failures, "visual", "VISUAL_SURFACE_MISMATCH", `Visual surface ${surface.unitIndex} is missing or has incompatible dimensions`, {
        expected: { width: surface.width, height: surface.height, rgbaBytes: surface.data.length },
        actual: observed === undefined ? undefined : { width: observed.width, height: observed.height, rgbaBytes: observed.data.length },
      });
      continue;
    }
    const ratios = pixelRatios(
      surface.data,
      observed.data,
      surface.width,
      surface.height,
      policy.pixelChannel,
      policy.pixelRadius,
    );
    const similarity = ssim(surface, observed);
    exactRatios.push(ratios.exact);
    tolerantRatios.push(ratios.tolerant);
    ssimValues.push(similarity);
    const exactPasses = ratios.exact >= policy.minExactPixelRatio;
    const tolerantPasses = ratios.tolerant >= policy.minTolerantPixelRatio;
    const ssimPasses = similarity >= policy.minSsim;
    check(counter, exactPasses);
    check(counter, tolerantPasses);
    check(counter, ssimPasses);
    if (!exactPasses) addFailure(failures, "visual", "EXACT_PIXEL_DIFF", "Exact pixel similarity is below threshold", { expected: policy.minExactPixelRatio, actual: ratios.exact });
    if (!tolerantPasses) addFailure(failures, "visual", "TOLERANT_PIXEL_DIFF", "Tolerance pixel similarity is below threshold", { expected: policy.minTolerantPixelRatio, actual: ratios.tolerant });
    if (!ssimPasses) addFailure(failures, "visual", "SSIM_BELOW_THRESHOLD", "Windowed SSIM is below threshold", { expected: policy.minSsim, actual: similarity });

    const unit = units.get(surface.unitIndex);
    if (unit !== undefined) {
      let regionCount = 0;
      let regionPixels = 0;
      let regionNeighborWork = 0;
      for (const object of expected.objects.filter((candidate) => candidate.unitIndex === surface.unitIndex)) {
        const region = regionPixelBounds(object, unit, surface);
        if (region === undefined) continue;
        const pixels = (region.right - region.left) * (region.bottom - region.top);
        const diameter = policy.pixelRadius * 2 + 1;
        const neighborWork = pixels * diameter * diameter * 4 * 2;
        if (regionCount === MAX_OBJECT_REGION_COUNT
          || regionPixels + pixels > MAX_OBJECT_REGION_PIXELS
          || regionNeighborWork + neighborWork > MAX_OBJECT_REGION_NEIGHBOR_WORK) {
          addFailure(
            failures,
            "visual",
            "OBJECT_REGION_LIMIT",
            "Object-region visual comparison exceeds its bounded work budget",
            {
              sourceKey: object.sourceKey,
              expected: {
                maxRegions: MAX_OBJECT_REGION_COUNT,
                maxPixels: MAX_OBJECT_REGION_PIXELS,
                maxNeighborWork: MAX_OBJECT_REGION_NEIGHBOR_WORK,
              },
              actual: {
                regions: regionCount + 1,
                pixels: regionPixels + pixels,
                neighborWork: regionNeighborWork + neighborWork,
              },
            },
          );
          break;
        }
        regionCount += 1;
        regionPixels += pixels;
        regionNeighborWork += neighborWork;
        const objectSimilarity = regionSimilarity(
          surface.data,
          observed.data,
          region,
          surface.width,
          surface.height,
          policy.pixelChannel,
          policy.pixelRadius,
        );
        regionValues.push(objectSimilarity);
        const regionPasses = objectSimilarity >= policy.minObjectRegionSimilarity;
        check(counter, regionPasses);
        if (!regionPasses) {
          addFailure(failures, "visual", "OBJECT_REGION_DIFF", "Object-region visual similarity is below threshold", {
            sourceKey: object.sourceKey,
            expected: policy.minObjectRegionSimilarity,
            actual: objectSimilarity,
          });
        }
      }
    }
  }
  const average = (values: readonly number[]) => values.length === 0 ? 1 : values.reduce((sum, value) => sum + value, 0) / values.length;
  return section(counter, {
    exactPixelSimilarity: rounded(average(exactRatios)),
    tolerantPixelSimilarity: rounded(average(tolerantRatios)),
    pixelChannelTolerance: policy.pixelChannel,
    pixelRadiusTolerance: policy.pixelRadius,
    ssim: rounded(average(ssimValues)),
    objectRegionSimilarity: rounded(average(regionValues)),
  }, expectedVisuals.length !== 0);
}

function compareMapping(expected: AccuracySnapshot, actual: AccuracySnapshot, failures: AccuracyFailure[]): AccuracySectionReport {
  const expectedProbes = expected.hitProbes ?? [];
  const actualProbes = new Map((actual.hitProbes ?? []).map((probe) => [probe.id, probe]));
  const actualSources = bySource(actual.objects);
  const counter: Counter = { passed: 0, total: 0 };
  let mappedSourceCount = 0;
  for (const probe of expectedProbes) {
    const observed = actualProbes.get(probe.id);
    const matches = observed !== undefined
      && observed.kind === probe.kind
      && observed.unitIndex === probe.unitIndex
      && same(observed.hits, probe.hits);
    check(counter, matches);
    if (!matches) {
      addFailure(failures, "object-mapping", "HIT_MAPPING_MISMATCH", `Hit probe ${probe.id} differs`, {
        expected: probe,
        actual: observed,
      });
    }
    if (observed !== undefined) {
      for (const sourceKey of observed.hits) {
        const mapped = actualSources.has(sourceKey);
        check(counter, mapped);
        if (mapped) mappedSourceCount += 1;
        else {
          addFailure(failures, "object-mapping", "HIT_SOURCE_UNMAPPED", `Hit probe ${probe.id} returned an object without a source mapping`, {
            sourceKey,
          });
        }
      }
    }
  }
  return section(counter, {
    probeCount: expectedProbes.length,
    matchedProbeCount: expectedProbes.filter((probe) => same(actualProbes.get(probe.id)?.hits, probe.hits)).length,
    mappedSourceCount,
  }, expectedProbes.length !== 0);
}

function diagnosticKey(diagnostic: AccuracyDiagnosticObservation): string {
  return canonical({
    code: diagnostic.code,
    fidelity: diagnostic.fidelity,
    phase: diagnostic.phase,
    part: diagnostic.part ?? null,
    sourceKey: diagnostic.sourceKey ?? null,
  });
}

function compareDiagnostics(expected: AccuracySnapshot, actual: AccuracySnapshot, failures: AccuracyFailure[]): AccuracySectionReport {
  const expectedDiagnostics = expected.diagnostics ?? [];
  const actualDiagnostics = actual.diagnostics ?? [];
  const expectedCounts = new Map<string, number>();
  const actualCounts = new Map<string, number>();
  for (const diagnostic of expectedDiagnostics) expectedCounts.set(diagnosticKey(diagnostic), (expectedCounts.get(diagnosticKey(diagnostic)) ?? 0) + 1);
  for (const diagnostic of actualDiagnostics) actualCounts.set(diagnosticKey(diagnostic), (actualCounts.get(diagnosticKey(diagnostic)) ?? 0) + 1);
  const keys = new Set([...expectedCounts.keys(), ...actualCounts.keys()]);
  const counter: Counter = { passed: 0, total: 0 };
  for (const key of keys) {
    const expectedCount = expectedCounts.get(key) ?? 0;
    const actualCount = actualCounts.get(key) ?? 0;
    const matches = expectedCount === actualCount;
    check(counter, matches);
    if (!matches) {
      addFailure(failures, "compatibility", "DIAGNOSTIC_MISMATCH", "Unsupported, approximate, font, external-resource, or active-content diagnostic differs", {
        expected: { diagnostic: JSON.parse(key) as unknown, count: expectedCount },
        actual: { diagnostic: JSON.parse(key) as unknown, count: actualCount },
      });
    }
  }
  return section(counter, {
    expectedDiagnosticCount: expectedDiagnostics.length,
    actualDiagnosticCount: actualDiagnostics.length,
    unsupportedCount: actualDiagnostics.filter(({ fidelity }) => fidelity === "unsupported").length,
    approximateCount: actualDiagnostics.filter(({ fidelity }) => fidelity === "approximate").length,
    fontSubstitutionCount: actualDiagnostics.filter(({ code }) => code === "FONT_SUBSTITUTED").length,
    externalResourceCount: actualDiagnostics.filter(({ code }) => code === "EXTERNAL_RESOURCE_BLOCKED").length,
    activeContentCount: actualDiagnostics.filter(({ code }) => code === "ACTIVE_CONTENT_BLOCKED").length,
  }, true);
}

function identityState(snapshot: AccuracySnapshot): unknown {
  return snapshot.objects.map(({ id, sourceKey }) => ({ id, sourceKey }));
}

function layoutState(snapshot: AccuracySnapshot): unknown {
  return snapshot.objects.map(({ sourceKey, textLayout }) => ({ sourceKey, textLayout: textLayout ?? null }));
}

function compareDeterminism(runs: readonly AccuracySnapshot[], failures: AccuracyFailure[]): AccuracySectionReport {
  const counter: Counter = { passed: 0, total: 0 };
  if (runs.length < 2) return section(counter, { runCount: runs.length }, false);
  const baseline = runs[0]!;
  for (let index = 1; index < runs.length; index += 1) {
    const run = runs[index]!;
    const displayListMatches = same(baseline.displayList, run.displayList);
    const identityMatches = same(identityState(baseline), identityState(run));
    const layoutMatches = same(layoutState(baseline), layoutState(run));
    check(counter, displayListMatches);
    check(counter, identityMatches);
    check(counter, layoutMatches);
    if (!displayListMatches) addFailure(failures, "determinism", "DISPLAY_LIST_NONDETERMINISTIC", `Display List differs in repeated run ${index + 1}`);
    if (!identityMatches) addFailure(failures, "determinism", "OBJECT_ID_NONDETERMINISTIC", `Object IDs differ in repeated run ${index + 1}`);
    if (!layoutMatches) addFailure(failures, "determinism", "LAYOUT_NONDETERMINISTIC", `Text layout differs in repeated run ${index + 1}`);
  }
  return section(counter, { runCount: runs.length }, true);
}

function sameRenderState(left: AccuracySnapshot, right: AccuracySnapshot): boolean {
  if (!same(left.units, right.units) || !same(left.objects, right.objects) || !same(left.displayList, right.displayList)) return false;
  const leftVisuals = left.visuals ?? [];
  const rightVisuals = right.visuals ?? [];
  if (leftVisuals.length !== rightVisuals.length) return false;
  for (let index = 0; index < leftVisuals.length; index += 1) {
    const a = leftVisuals[index]!;
    const b = rightVisuals[index]!;
    if (a.unitIndex !== b.unitIndex || a.width !== b.width || a.height !== b.height || !same(a.viewport, b.viewport) || !bytesEqual(a.data, b.data)) return false;
  }
  return true;
}

function compareMetamorphic(
  baseline: AccuracySnapshot,
  variants: readonly MetamorphicVariantObservation[],
  failures: AccuracyFailure[],
): AccuracySectionReport {
  const counter: Counter = { passed: 0, total: 0 };
  if (variants.length === 0) return section(counter, { variantCount: 0 }, false);
  const required: readonly MetamorphicVariantKind[] = [
    "xml-attribute-order",
    "namespace-prefix",
    "zip-entry-order",
    "irrelevant-metadata",
  ];
  for (const kind of required) {
    const matches = variants.filter((variant) => variant.kind === kind);
    const existsOnce = matches.length === 1;
    check(counter, existsOnce);
    if (!existsOnce) {
      addFailure(failures, "metamorphic", "METAMORPHIC_VARIANT_MISSING", `Expected exactly one ${kind} variant`, {
        expected: 1,
        actual: matches.length,
      });
      continue;
    }
    const invariant = sameRenderState(baseline, matches[0]!.snapshot);
    check(counter, invariant);
    if (!invariant) {
      addFailure(failures, "metamorphic", "METAMORPHIC_RENDER_CHANGED", `Rendering changed after ${kind} transformation`);
    }
  }
  return section(counter, { variantCount: variants.length }, true);
}

function resolvePolicy(policy: AccuracyTolerancePolicy | undefined): Required<AccuracyTolerancePolicy> {
  const resolved = { ...DEFAULT_POLICY, ...policy };
  for (const [name, value] of Object.entries(resolved)) {
    if (!Number.isFinite(value) || value < 0) throw new TypeError(`Accuracy policy ${name} must be a finite non-negative number`);
  }
  for (const name of ["minExactPixelRatio", "minTolerantPixelRatio", "minSsim"] as const) {
    if (resolved[name] > 1) throw new TypeError(`Accuracy policy ${name} must not exceed 1`);
  }
  if (resolved.minObjectRegionSimilarity > 1) throw new TypeError("Accuracy policy minObjectRegionSimilarity must not exceed 1");
  if (resolved.minOverallScore > 100) throw new TypeError("Accuracy policy minOverallScore must not exceed 100");
  if (resolved.pixelChannel > 255) throw new TypeError("Accuracy policy pixelChannel must not exceed 255");
  if (!Number.isInteger(resolved.pixelRadius) || resolved.pixelRadius > 8) {
    throw new TypeError("Accuracy policy pixelRadius must be an integer not exceeding 8");
  }
  return resolved;
}

function weightedScore(sections: Record<keyof typeof SECTION_WEIGHTS, AccuracySectionReport>): number {
  let weighted = 0;
  let weights = 0;
  for (const [name, weight] of Object.entries(SECTION_WEIGHTS) as [keyof typeof SECTION_WEIGHTS, number][]) {
    const value = sections[name];
    if (!value.applicable) continue;
    weighted += value.score * weight;
    weights += weight;
  }
  return weights === 0 ? 100 : rounded(weighted / weights);
}

function hasObservationValue(value: unknown): boolean {
  if (value === undefined || value === null) return false;
  if (Array.isArray(value) || value instanceof Uint8Array) return value.length !== 0;
  if (typeof value === "object") return Object.keys(value).length !== 0;
  return true;
}

function hasSnapshotEvidence(snapshot: AccuracySnapshot): boolean {
  return snapshot.units.length !== 0
    || snapshot.objects.length !== 0
    || (snapshot.visuals?.length ?? 0) !== 0
    || (snapshot.hitProbes?.length ?? 0) !== 0
    || (snapshot.diagnostics?.length ?? 0) !== 0
    || hasObservationValue(snapshot.displayList);
}

function hasContentFact(object: AccuracyObjectObservation, key: string, type: "string" | "number"): boolean {
  return object.content !== undefined && typeof object.content[key] === type;
}

function hasTextLine(snapshot: AccuracySnapshot): boolean {
  return snapshot.objects.some(({ textLayout }) => (textLayout?.lines.length ?? 0) !== 0);
}

function hasRepeatedEvidence(
  request: EvaluateAccuracyRequest,
  predicate: (snapshot: AccuracySnapshot) => boolean,
): boolean {
  const runs = request.repeatedRuns ?? [];
  return runs.length >= 2 && runs.every(predicate);
}

function coverageIsObserved(coverage: AccuracyCoverage, request: EvaluateAccuracyRequest): boolean {
  const snapshot = request.expected;
  const visuals = snapshot.visuals ?? [];
  const probes = snapshot.hitProbes ?? [];
  const diagnostics = snapshot.diagnostics ?? [];
  switch (coverage) {
    case "structure-units": return snapshot.units.length !== 0;
    case "structure-objects": return snapshot.objects.length !== 0;
    case "content-text": return snapshot.objects.some((object) => hasContentFact(object, "text", "string"));
    case "content-image": return snapshot.objects.some((object) => object.type === "image" && hasContentFact(object, "digest", "string"));
    case "content-table": return snapshot.objects.some((object) => object.type === "table"
      && hasContentFact(object, "rows", "number") && hasContentFact(object, "columns", "number"));
    case "content-cell": return snapshot.objects.some((object) => object.type === "cell" && hasContentFact(object, "address", "string"));
    case "geometry-bounds": return snapshot.objects.some(({ bounds }) => [bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite));
    case "geometry-rotation": return snapshot.objects.some(({ rotation }) => rotation !== undefined && Number.isFinite(rotation));
    case "geometry-crop": return snapshot.objects.some(({ crop }) => crop !== undefined && [crop.x, crop.y, crop.width, crop.height].every(Number.isFinite));
    case "geometry-z-order": return snapshot.objects.some(({ zOrder }) => zOrder !== undefined && Number.isFinite(zOrder));
    case "text-lines": return hasTextLine(snapshot);
    case "text-breaks": return snapshot.objects.some(({ textLayout }) => (textLayout?.lines.length ?? 0) >= 2);
    case "text-baseline": return snapshot.objects.some(({ textLayout }) => textLayout?.lines.some(({ baseline }) => Number.isFinite(baseline)) === true);
    case "text-font-size": return snapshot.objects.some(({ textLayout }) => textLayout?.lines.some(({ fontSize }) => Number.isFinite(fontSize)) === true);
    case "text-line-spacing": return snapshot.objects.some(({ textLayout }) => textLayout?.lines.some(({ lineHeight }) => Number.isFinite(lineHeight)) === true);
    case "text-overflow": return snapshot.objects.some(({ textLayout }) => typeof textLayout?.overflow === "boolean");
    case "text-auto-scale": return snapshot.objects.some(({ textLayout }) => Number.isFinite(textLayout?.autoScale));
    case "visual-exact":
    case "visual-tolerant":
    case "visual-ssim": return visuals.some(validSurface);
    case "visual-object-region": {
      const units = byUnit(snapshot.units);
      return visuals.some((surface) => {
        if (!validSurface(surface)) return false;
        const unit = units.get(surface.unitIndex);
        return unit !== undefined && snapshot.objects.some((object) => (
          object.unitIndex === surface.unitIndex
          && regionPixelBounds(object, unit, surface) !== undefined
        ));
      });
    }
    case "hit-center": return probes.some(({ kind }) => kind === "center");
    case "hit-edge": return probes.some(({ kind }) => kind === "edge");
    case "hit-rotated": return probes.some(({ kind }) => kind === "rotated");
    case "hit-group": return probes.some(({ kind }) => kind === "group");
    case "hit-overlap": return probes.some(({ kind }) => kind === "overlap");
    case "hit-z-order": return probes.some(({ kind }) => kind === "z-order");
    case "diagnostic-unsupported": return diagnostics.some(({ fidelity }) => fidelity === "unsupported");
    case "diagnostic-approximate": return diagnostics.some(({ fidelity }) => fidelity === "approximate");
    case "diagnostic-font-substitution": return diagnostics.some(({ code }) => code === "FONT_SUBSTITUTED");
    case "diagnostic-external-resource": return diagnostics.some(({ code }) => code === "EXTERNAL_RESOURCE_BLOCKED");
    case "diagnostic-active-content": return diagnostics.some(({ code }) => code === "ACTIVE_CONTENT_BLOCKED");
    case "determinism-display-list": return hasRepeatedEvidence(request, ({ displayList }) => hasObservationValue(displayList));
    case "determinism-object-id": return hasRepeatedEvidence(request, ({ objects }) => objects.length !== 0 && objects.every(({ id }) => id !== ""));
    case "determinism-layout": return hasRepeatedEvidence(request, hasTextLine);
    case "metamorphic-xml-attribute-order": return hasMetamorphicEvidence(request, "xml-attribute-order");
    case "metamorphic-namespace-prefix": return hasMetamorphicEvidence(request, "namespace-prefix");
    case "metamorphic-zip-entry-order": return hasMetamorphicEvidence(request, "zip-entry-order");
    case "metamorphic-irrelevant-metadata": return hasMetamorphicEvidence(request, "irrelevant-metadata");
  }
}

function hasMetamorphicEvidence(request: EvaluateAccuracyRequest, kind: MetamorphicVariantKind): boolean {
  const variants = (request.metamorphicVariants ?? []).filter((variant) => variant.kind === kind);
  return variants.length === 1 && hasSnapshotEvidence(variants[0]!.snapshot);
}

function validateDeclaredCoverage(request: EvaluateAccuracyRequest): AccuracyCoverageValidation {
  const declared = [...new Set(request.declaredCoverage ?? [])];
  const observed = declared.filter((coverage) => coverageIsObserved(coverage, request));
  const observedSet = new Set(observed);
  const missing = declared.filter((coverage) => !observedSet.has(coverage));
  return Object.freeze({
    declared: Object.freeze(declared),
    observed: Object.freeze(observed),
    missing: Object.freeze(missing),
  });
}

/** Evaluate independent source-oracle and renderer observations for one file. */
export function evaluateAccuracy(request: EvaluateAccuracyRequest): AccuracyReport {
  const policy = resolvePolicy(request.policy);
  const failures: AccuracyFailure[] = [];
  const expectedHasEvidence = hasSnapshotEvidence(request.expected);
  const actualHasEvidence = hasSnapshotEvidence(request.actual);
  if (!expectedHasEvidence) {
    addFailure(failures, "overall", "EXPECTED_SNAPSHOT_EMPTY", "Expected accuracy snapshot contains no oracle evidence");
  }
  if (!actualHasEvidence) {
    addFailure(failures, "overall", "ACTUAL_SNAPSHOT_EMPTY", "Actual accuracy snapshot contains no renderer evidence");
  }
  const coverageValidation = validateDeclaredCoverage(request);
  for (const coverage of coverageValidation.missing) {
    addFailure(failures, "overall", "DECLARED_COVERAGE_NOT_OBSERVED", `Declared accuracy coverage ${coverage} has no matching oracle evidence`, { coverage });
  }
  const contentCompleteness = compareContent(request.expected, request.actual, failures);
  const geometryAccuracy = compareGeometry(
    request.expected,
    request.actual,
    failures,
    policy.geometry,
    policy.rotationDegrees,
  );
  const textLayoutAccuracy = compareTextLayout(
    request.expected,
    request.actual,
    failures,
    policy.geometry,
    policy.textRelative,
  );
  const visualSimilarity = compareVisuals(request.expected, request.actual, failures, policy);
  const objectMappingAccuracy = compareMapping(request.expected, request.actual, failures);
  const compatibilityAccuracy = compareDiagnostics(request.expected, request.actual, failures);
  const determinismAccuracy = compareDeterminism(request.repeatedRuns ?? [], failures);
  const metamorphicAccuracy = compareMetamorphic(request.actual, request.metamorphicVariants ?? [], failures);
  const sections = {
    contentCompleteness,
    geometryAccuracy,
    textLayoutAccuracy,
    visualSimilarity,
    objectMappingAccuracy,
    compatibilityAccuracy,
    determinismAccuracy,
    metamorphicAccuracy,
  };
  const observationsValid = expectedHasEvidence && actualHasEvidence && coverageValidation.missing.length === 0;
  const overallScore = observationsValid ? weightedScore(sections) : 0;
  if (observationsValid && overallScore < policy.minOverallScore) {
    addFailure(failures, "overall", "OVERALL_SCORE_BELOW_THRESHOLD", "Overall accuracy score is below threshold", {
      expected: policy.minOverallScore,
      actual: overallScore,
    });
  }
  return Object.freeze({
    fileId: request.fileId,
    corpusClass: request.corpusClass,
    passed: failures.length === 0,
    coverageValidation,
    ...sections,
    overallScore,
    failureReasons: Object.freeze(failures),
  });
}
