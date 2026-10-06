import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { dirname, isAbsolute, normalize, resolve, sep } from "node:path";

export const NATIVE_VISUAL_POLICY_ID = "native-visual-v1";
export const NATIVE_REFERENCE_REVIEW_SCOPE = "all-rendered-units-visual-fidelity-v1";

export const NATIVE_VISUAL_POLICY = Object.freeze({
  geometry: 0.002,
  rotationDegrees: 0.25,
  textRelative: 0.01,
  pixelChannel: 8,
  pixelRadius: 1,
  minExactPixelRatio: 0,
  minTolerantPixelRatio: 0.995,
  minSsim: 0.99,
  minObjectRegionSimilarity: 0.99,
  minOverallScore: 95,
});

export const NATIVE_FORMAT_RULES = Object.freeze({
  pptx: Object.freeze({ oracleSuite: "microsoft-office", application: "powerpoint", applicationName: "Microsoft PowerPoint", bundleId: "com.microsoft.Powerpoint", extension: ".pptx", unitType: "slide", capture: "pdf-export" }),
  ppt: Object.freeze({ oracleSuite: "microsoft-office", application: "powerpoint", applicationName: "Microsoft PowerPoint", bundleId: "com.microsoft.Powerpoint", extension: ".ppt", unitType: "slide", capture: "pdf-export" }),
  docx: Object.freeze({ oracleSuite: "microsoft-office", application: "word", applicationName: "Microsoft Word", bundleId: "com.microsoft.Word", extension: ".docx", unitType: "page", capture: "pdf-export" }),
  doc: Object.freeze({ oracleSuite: "microsoft-office", application: "word", applicationName: "Microsoft Word", bundleId: "com.microsoft.Word", extension: ".doc", unitType: "page", capture: "pdf-export" }),
  xlsx: Object.freeze({ oracleSuite: "microsoft-office", application: "excel", applicationName: "Microsoft Excel", bundleId: "com.microsoft.Excel", extension: ".xlsx", unitType: "sheet", capture: "sheet-viewport" }),
  xls: Object.freeze({ oracleSuite: "microsoft-office", application: "excel", applicationName: "Microsoft Excel", bundleId: "com.microsoft.Excel", extension: ".xls", unitType: "sheet", capture: "sheet-viewport" }),
  keynote: Object.freeze({ oracleSuite: "apple-iwork", application: "keynote", applicationName: "Keynote", bundleId: "com.apple.iWork.Keynote", extension: ".key", unitType: "slide", capture: "pdf-export" }),
  pages: Object.freeze({ oracleSuite: "apple-iwork", application: "pages", applicationName: "Pages", bundleId: "com.apple.iWork.Pages", extension: ".pages", unitType: "page", capture: "pdf-export" }),
  numbers: Object.freeze({ oracleSuite: "apple-iwork", application: "numbers", applicationName: "Numbers", bundleId: "com.apple.iWork.Numbers", extension: ".numbers", unitType: "sheet", capture: "sheet-viewport" }),
  odp: Object.freeze({ oracleSuite: "microsoft-office", application: "powerpoint", applicationName: "Microsoft PowerPoint", bundleId: "com.microsoft.Powerpoint", extension: ".odp", unitType: "slide", capture: "pdf-export" }),
  odt: Object.freeze({ oracleSuite: "microsoft-office", application: "word", applicationName: "Microsoft Word", bundleId: "com.microsoft.Word", extension: ".odt", unitType: "page", capture: "pdf-export" }),
  ods: Object.freeze({ oracleSuite: "microsoft-office", application: "excel", applicationName: "Microsoft Excel", bundleId: "com.microsoft.Excel", extension: ".ods", unitType: "sheet", capture: "sheet-viewport" }),
  wps: Object.freeze({ oracleSuite: "wps-office", application: "wps-writer", applicationName: "WPS Office", bundleId: "com.kingsoft.wpsoffice.mac", extension: ".wps", unitType: "page", capture: "pdf-export" }),
  et: Object.freeze({ oracleSuite: "wps-office", application: "wps-spreadsheets", applicationName: "WPS Office", bundleId: "com.kingsoft.wpsoffice.mac", extension: ".et", unitType: "sheet", capture: "sheet-viewport" }),
  dps: Object.freeze({ oracleSuite: "wps-office", application: "wps-presentation", applicationName: "WPS Office", bundleId: "com.kingsoft.wpsoffice.mac", extension: ".dps", unitType: "slide", capture: "pdf-export" }),
});

export const CANDIDATE_FINGERPRINT_FIELDS = Object.freeze([
  "os",
  "osVersion",
  "architecture",
  "locale",
  "timezone",
  "colorSpace",
  "devicePixelRatio",
  "scale",
  "background",
  "fontSetDigest",
  "browser",
  "browserVersion",
]);

const CORPUS_CLASSES = new Set(["minimal", "combination", "enterprise", "large"]);
const ORACLE_SUITES = new Set(["microsoft-office", "apple-iwork", "wps-office"]);
const SUITE_KINDS = new Set(["regression", "release"]);
const SHA256 = /^sha256:[0-9a-f]{64}$/u;
const SHEET_RANGE = /^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/u;

function sheetColumnNumber(label) {
  return [...label].reduce((number, character) => number * 26 + character.charCodeAt(0) - 64, 0);
}

function assertSheetRange(value, label) {
  const match = typeof value === "string" ? SHEET_RANGE.exec(value) : null;
  if (match === null) throw new Error(`${label} must be an uppercase bounded A1 range such as A1:H40`);
  const startColumn = sheetColumnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = sheetColumnNumber(match[3]);
  const endRow = Number(match[4]);
  if (startColumn > 16_384 || endColumn > 16_384 || startRow > 1_048_576 || endRow > 1_048_576
    || startColumn > endColumn || startRow > endRow) {
    throw new Error(`${label} must stay within A1:XFD1048576 and run from top-left to bottom-right`);
  }
}

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function assertRecord(value, label) {
  if (!isRecord(value)) throw new Error(`${label} must be a JSON object`);
}

function assertKnownKeys(value, allowed, label) {
  for (const key of Object.keys(value)) {
    if (!allowed.has(key)) throw new Error(`${label}.${key} is not supported`);
  }
}

function assertNonEmptyString(value, label) {
  if (typeof value !== "string" || value.trim() !== value || value === "") {
    throw new Error(`${label} must be a trimmed non-empty string`);
  }
}

function assertSha256(value, label) {
  if (typeof value !== "string" || !SHA256.test(value)) {
    throw new Error(`${label} must be a lowercase SHA-256 digest`);
  }
}

export function assertSafeRelativePath(value, label) {
  assertNonEmptyString(value, label);
  if (value.length > 320 || isAbsolute(value) || value.includes("\0") || value.includes("\\")
    || value.split("/").some((segment) => segment === "" || segment === "." || segment === ".."
      || !/^[A-Za-z0-9][A-Za-z0-9._ -]*$/u.test(segment))) {
    throw new Error(`${label} must be a safe relative path`);
  }
  const normalized = normalize(value);
  if (normalized === "." || normalized === ".." || normalized.startsWith(`..${sep}`)) {
    throw new Error(`${label} must stay inside its configured root`);
  }
  return normalized;
}

export function resolveSafeRelativePath(root, value, label) {
  const normalized = assertSafeRelativePath(value, label);
  const target = resolve(root, normalized);
  const resolvedRoot = resolve(root);
  if (target !== resolvedRoot && !target.startsWith(`${resolvedRoot}${sep}`)) {
    throw new Error(`${label} must stay inside its configured root`);
  }
  return target;
}

function validateApplicationEntry(entry, key, expectedCapture) {
  const label = `oracleFingerprint.applications.${key}`;
  assertRecord(entry, label);
  assertKnownKeys(entry, new Set(["version", "build", "capture"]), label);
  assertNonEmptyString(entry.version, `${label}.version`);
  assertNonEmptyString(entry.build, `${label}.build`);
  if (entry.capture !== expectedCapture) {
    throw new Error(`${label}.capture must be ${expectedCapture}`);
  }
}

function validateOracleFingerprint(fingerprint, formats) {
  assertRecord(fingerprint, "oracleFingerprint");
  assertKnownKeys(fingerprint, new Set([
    "os", "osVersion", "architecture", "locale", "timezone", "colorSpace", "scale", "background",
    "fontSetDigest", "applications", "rasterizer",
  ]), "oracleFingerprint");
  for (const field of ["os", "osVersion", "architecture", "locale", "timezone", "colorSpace", "background"]) {
    assertNonEmptyString(fingerprint[field], `oracleFingerprint.${field}`);
  }
  if (!Number.isFinite(fingerprint.scale) || fingerprint.scale <= 0) {
    throw new Error("oracleFingerprint.scale must be a finite positive number");
  }
  if (fingerprint.scale !== 1) throw new Error("oracleFingerprint.scale must be 1");
  if (fingerprint.background !== "#ffffff") throw new Error("oracleFingerprint.background must be #ffffff");
  if (fingerprint.colorSpace !== "srgb") throw new Error("oracleFingerprint.colorSpace must be srgb");
  assertSha256(fingerprint.fontSetDigest, "oracleFingerprint.fontSetDigest");
  assertRecord(fingerprint.applications, "oracleFingerprint.applications");
  const applications = new Map();
  for (const format of formats) {
    const rule = NATIVE_FORMAT_RULES[format];
    const previousCapture = applications.get(rule.application);
    if (previousCapture !== undefined && previousCapture !== rule.capture) {
      throw new Error(`Formats for ${rule.application} require inconsistent capture methods`);
    }
    applications.set(rule.application, rule.capture);
  }
  assertKnownKeys(fingerprint.applications, new Set(applications.keys()), "oracleFingerprint.applications");
  for (const [application, capture] of applications) {
    validateApplicationEntry(fingerprint.applications[application], application, capture);
  }
  const needsRasterizer = formats.some((format) => NATIVE_FORMAT_RULES[format].capture === "pdf-export");
  if (needsRasterizer) {
    assertRecord(fingerprint.rasterizer, "oracleFingerprint.rasterizer");
    assertKnownKeys(fingerprint.rasterizer, new Set(["name", "version", "dpi"]), "oracleFingerprint.rasterizer");
    if (fingerprint.rasterizer.name !== "pdftoppm") throw new Error("oracleFingerprint.rasterizer.name must be pdftoppm");
    assertNonEmptyString(fingerprint.rasterizer.version, "oracleFingerprint.rasterizer.version");
    if (fingerprint.rasterizer.dpi !== 96) throw new Error("oracleFingerprint.rasterizer.dpi must be 96");
  } else if (fingerprint.rasterizer !== undefined) {
    throw new Error("oracleFingerprint.rasterizer is not used by a sheet-only suite");
  }
}

export function validateCandidateFingerprint(fingerprint, label = "candidateFingerprint", allowRuntimeMetadata = false) {
  assertRecord(fingerprint, label);
  assertKnownKeys(
    fingerprint,
    new Set([...CANDIDATE_FINGERPRINT_FIELDS, ...(allowRuntimeMetadata ? ["buildRevision"] : [])]),
    label,
  );
  for (const field of CANDIDATE_FINGERPRINT_FIELDS) {
    const value = fingerprint[field];
    if (field === "devicePixelRatio" || field === "scale") {
      if (!Number.isFinite(value) || value <= 0) throw new Error(`${label}.${field} must be a finite positive number`);
    } else {
      assertNonEmptyString(value, `${label}.${field}`);
    }
  }
  assertSha256(fingerprint.fontSetDigest, `${label}.fontSetDigest`);
  if (fingerprint.devicePixelRatio !== 1) throw new Error(`${label}.devicePixelRatio must be 1`);
  if (fingerprint.scale !== 1) throw new Error(`${label}.scale must be 1`);
  if (fingerprint.background !== "#ffffff") throw new Error(`${label}.background must be #ffffff`);
  if (fingerprint.colorSpace !== "srgb") throw new Error(`${label}.colorSpace must be srgb`);
}

function formatForFixture(fixture) {
  const lower = fixture.toLowerCase();
  return Object.entries(NATIVE_FORMAT_RULES).find(([, rule]) => lower.endsWith(rule.extension))?.[0];
}

function validateDocument(document, suite, ids, reportNames, assetPaths) {
  const label = `documents[${suite.documents.indexOf(document)}]`;
  assertRecord(document, label);
  assertKnownKeys(document, new Set([
    "id", "corpusClass", "format", "fixture", "fixtureSha256", "unitCount", "regressionBugId", "units",
  ]), label);
  assertNonEmptyString(document.id, `${label}.id`);
  if (!/^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/u.test(document.id)) {
    throw new Error(`${label}.id must be filename-safe ASCII`);
  }
  if (ids.has(document.id)) throw new Error(`Native visual document ${document.id} is duplicated`);
  ids.add(document.id);
  if (!CORPUS_CLASSES.has(document.corpusClass)) {
    throw new Error(`${label}.corpusClass must be minimal, combination, enterprise, or large`);
  }
  const rule = NATIVE_FORMAT_RULES[document.format];
  if (rule === undefined) throw new Error(`${label}.format is not one of the fifteen native-golden formats`);
  if (rule.oracleSuite !== suite.oracleSuite) {
    throw new Error(`${label}.format ${document.format} belongs to ${rule.oracleSuite}, not ${suite.oracleSuite}`);
  }
  assertSafeRelativePath(document.fixture, `${label}.fixture`);
  if (formatForFixture(document.fixture) !== document.format) {
    throw new Error(`${label}.fixture extension does not match format ${document.format}`);
  }
  assertSha256(document.fixtureSha256, `${label}.fixtureSha256`);
  if (!Number.isInteger(document.unitCount) || document.unitCount < 1 || document.unitCount > 100_000) {
    throw new Error(`${label}.unitCount must be an integer from 1 through 100000`);
  }
  if (document.regressionBugId !== undefined) assertNonEmptyString(document.regressionBugId, `${label}.regressionBugId`);
  if (!Array.isArray(document.units) || document.units.length !== document.unitCount) {
    throw new Error(`${label}.units must contain exactly unitCount entries`);
  }
  const indices = new Set();
  for (let unitOffset = 0; unitOffset < document.units.length; unitOffset += 1) {
    const unit = document.units[unitOffset];
    const unitLabel = `${label}.units[${unitOffset}]`;
    assertRecord(unit, unitLabel);
    assertKnownKeys(unit, new Set([
      "index", "sheetRange", "referenceJson", "referenceJsonSha256", "goldenPng", "goldenPngSha256",
      "actualPng", "actualObservationJson",
    ]), unitLabel);
    if (!Number.isInteger(unit.index) || unit.index < 0 || unit.index >= document.unitCount || indices.has(unit.index)) {
      throw new Error(`${unitLabel}.index must uniquely cover 0 through ${document.unitCount - 1}`);
    }
    indices.add(unit.index);
    if (rule.unitType === "sheet") {
      assertSheetRange(unit.sheetRange, `${unitLabel}.sheetRange`);
    } else if (unit.sheetRange !== undefined) {
      throw new Error(`${unitLabel}.sheetRange is valid only for sheet formats`);
    }
    for (const field of ["referenceJson", "goldenPng", "actualPng", "actualObservationJson"]) {
      assertSafeRelativePath(unit[field], `${unitLabel}.${field}`);
      if (assetPaths[field] !== undefined) {
        if (assetPaths[field].has(unit[field])) throw new Error(`${unitLabel}.${field} is duplicated`);
        assetPaths[field].add(unit[field]);
      }
    }
    if (!unit.referenceJson.toLowerCase().endsWith(".json")) throw new Error(`${unitLabel}.referenceJson must end with .json`);
    assertSha256(unit.referenceJsonSha256, `${unitLabel}.referenceJsonSha256`);
    if (unit.goldenPng === unit.actualPng) throw new Error(`${unitLabel} goldenPng and actualPng must differ`);
    if (!unit.goldenPng.toLowerCase().endsWith(".png") || !unit.actualPng.toLowerCase().endsWith(".png")) {
      throw new Error(`${unitLabel} PNG paths must end with .png`);
    }
    if (!unit.actualObservationJson.toLowerCase().endsWith(".json")) {
      throw new Error(`${unitLabel}.actualObservationJson must end with .json`);
    }
    assertSha256(unit.goldenPngSha256, `${unitLabel}.goldenPngSha256`);
    const reportName = `${document.id}-unit-${unit.index}`;
    if (reportNames.has(reportName)) throw new Error(`Native visual report name ${reportName} is duplicated`);
    reportNames.add(reportName);
  }
  for (let index = 0; index < document.unitCount; index += 1) {
    if (!indices.has(index)) throw new Error(`${label}.units does not cover unit ${index}`);
  }
  return rule;
}

export function validateNativeVisualSuite(suite) {
  assertRecord(suite, "Native visual suite");
  assertKnownKeys(suite, new Set([
    "schemaVersion", "suiteKind", "oracleMode", "oracleSuite", "thresholdPolicy", "oracleFingerprint",
    "candidateFingerprint", "documents",
  ]), "Native visual suite");
  if (suite.schemaVersion !== 2) throw new Error("Native visual suite schemaVersion must be 2");
  if (!SUITE_KINDS.has(suite.suiteKind)) throw new Error("Native visual suiteKind must be regression or release");
  if (suite.oracleMode !== "read-only") throw new Error("Native visual suite oracleMode must be read-only");
  if (!ORACLE_SUITES.has(suite.oracleSuite)) {
    throw new Error("Native visual oracleSuite must be microsoft-office, apple-iwork, or wps-office");
  }
  if (suite.thresholdPolicy !== NATIVE_VISUAL_POLICY_ID) {
    throw new Error(`Native visual thresholdPolicy must be ${NATIVE_VISUAL_POLICY_ID}`);
  }
  if (Object.hasOwn(suite, "policy")) throw new Error("Native visual suites cannot override the locked threshold policy");
  validateCandidateFingerprint(suite.candidateFingerprint);
  if (!Array.isArray(suite.documents) || suite.documents.length === 0) {
    throw new Error("Native visual suite must contain at least one document");
  }
  const ids = new Set();
  const reportNames = new Set();
  const assetPaths = {
    goldenPng: new Set(),
    actualPng: new Set(),
    actualObservationJson: new Set(),
  };
  const formats = [];
  for (const document of suite.documents) {
    const rule = validateDocument(document, suite, ids, reportNames, assetPaths);
    formats.push(document.format);
    if (rule === undefined) throw new Error("unreachable format validation failure");
  }
  if (suite.suiteKind === "release") {
    const required = Object.entries(NATIVE_FORMAT_RULES)
      .filter(([, rule]) => rule.oracleSuite === suite.oracleSuite)
      .map(([format]) => format);
    const present = new Set(formats);
    const missing = required.filter((format) => !present.has(format));
    if (missing.length !== 0) throw new Error(`Native release suite is missing formats: ${missing.join(", ")}`);
    const presentClasses = new Set(suite.documents.map(({ corpusClass }) => corpusClass));
    const missingClasses = [...CORPUS_CLASSES].filter((corpusClass) => !presentClasses.has(corpusClass));
    if (missingClasses.length !== 0) {
      throw new Error(`Native release suite is missing trusted corpus classes: ${missingClasses.join(", ")}`);
    }
  }
  validateOracleFingerprint(suite.oracleFingerprint, formats);
  return suite;
}

export function flattenNativeVisualCases(suite) {
  validateNativeVisualSuite(suite);
  return suite.documents.flatMap((document) => {
    const rule = NATIVE_FORMAT_RULES[document.format];
    return document.units.map((unit) => ({
      id: `${document.id}-unit-${unit.index}`,
      documentId: document.id,
      corpusClass: document.corpusClass,
      regressionBugId: document.regressionBugId,
      format: document.format,
      fixture: document.fixture,
      fixtureSha256: document.fixtureSha256,
      unitIndex: unit.index,
      unitCount: document.unitCount,
      unitType: rule.unitType,
      sheetRange: unit.sheetRange,
      referenceJson: unit.referenceJson,
      referenceJsonSha256: unit.referenceJsonSha256,
      goldenPng: unit.goldenPng,
      goldenPngSha256: unit.goldenPngSha256,
      actualPng: unit.actualPng,
      actualObservationJson: unit.actualObservationJson,
      oracleApplication: rule.application,
      capture: rule.capture,
    }));
  });
}

function rawSha256(value, label) {
  if (typeof value !== "string" || !/^[0-9a-f]{64}$/u.test(value)) {
    throw new Error(`${label} must be a raw lowercase SHA-256 digest`);
  }
  return `sha256:${value}`;
}

function validateReviewAttestation(reference, rule, artifactSha256, label) {
  const attestation = reference.reviewAttestation;
  const attestationLabel = `${label}.reviewAttestation`;
  assertRecord(attestation, attestationLabel);
  assertKnownKeys(attestation, new Set([
    "status", "reviewer", "reviewedAt", "scope", "oracleSuite", "application", "sourceSha256", "artifactSha256",
  ]), attestationLabel);
  if (attestation.status !== "reviewed") throw new Error(`${attestationLabel}.status must be reviewed`);
  assertNonEmptyString(attestation.reviewer, `${attestationLabel}.reviewer`);
  if (!/^[A-Za-z0-9](?:[A-Za-z0-9._@:-]{0,126}[A-Za-z0-9])?$/u.test(attestation.reviewer)) {
    throw new Error(`${attestationLabel}.reviewer must be a stable non-whitespace ID`);
  }
  assertNonEmptyString(attestation.reviewedAt, `${attestationLabel}.reviewedAt`);
  const reviewedAt = Date.parse(attestation.reviewedAt);
  if (!Number.isFinite(reviewedAt) || new Date(reviewedAt).toISOString() !== attestation.reviewedAt) {
    throw new Error(`${attestationLabel}.reviewedAt must be a canonical ISO-8601 instant`);
  }
  if (attestation.reviewedAt !== reference.generatedAt) {
    throw new Error(`${attestationLabel}.reviewedAt must equal generatedAt`);
  }
  if (attestation.scope !== NATIVE_REFERENCE_REVIEW_SCOPE) {
    throw new Error(`${attestationLabel}.scope must be ${NATIVE_REFERENCE_REVIEW_SCOPE}`);
  }
  if (attestation.oracleSuite !== rule.oracleSuite || attestation.application !== rule.applicationName) {
    throw new Error(`${attestationLabel} oracle suite or application does not match the format rule`);
  }
  if (attestation.sourceSha256 !== reference.source.sha256) {
    throw new Error(`${attestationLabel}.sourceSha256 does not match source.sha256`);
  }
  rawSha256(attestation.sourceSha256, `${attestationLabel}.sourceSha256`);
  if (attestation.artifactSha256 !== artifactSha256) {
    throw new Error(`${attestationLabel}.artifactSha256 does not match the reviewed artifact`);
  }
  rawSha256(attestation.artifactSha256, `${attestationLabel}.artifactSha256`);
}

async function validateReferencePdf(reference, referencePath, loaded, label) {
  if (!Array.isArray(reference.pdfs) || reference.pdfs.length !== 1) {
    throw new Error(`${label}.pdfs must contain exactly one PDF record`);
  }
  const record = reference.pdfs[0];
  const recordLabel = `${label}.pdfs[0]`;
  assertRecord(record, recordLabel);
  assertSafeRelativePath(record.file, `${recordLabel}.file`);
  if (!record.file.toLowerCase().endsWith(".pdf")) throw new Error(`${recordLabel}.file must end with .pdf`);
  if (!Number.isSafeInteger(record.bytes) || record.bytes <= 0) {
    throw new Error(`${recordLabel}.bytes must be a positive integer`);
  }
  const expectedSha256 = rawSha256(record.sha256, `${recordLabel}.sha256`);
  if (loaded.pdf !== undefined) return loaded.pdf;

  const pdfPath = resolveSafeRelativePath(dirname(referencePath), record.file, `${recordLabel}.file`);
  let bytes;
  try {
    bytes = await readFile(pdfPath);
  } catch (cause) {
    throw new Error(`${label} PDF could not be read`, { cause });
  }
  if (bytes.length === 0) throw new Error(`${label} PDF must not be empty`);
  if (bytes.length !== record.bytes) throw new Error(`${label} PDF byte length does not match reference.json`);
  if (!bytes.subarray(0, 5).equals(Buffer.from("%PDF-", "ascii"))) {
    throw new Error(`${label} PDF does not have a PDF signature`);
  }
  const actualSha256 = sha256(bytes);
  if (actualSha256 !== expectedSha256) throw new Error(`${label} PDF SHA-256 does not match reference.json`);
  loaded.pdf = Object.freeze({ file: record.file, bytes: record.bytes, sha256: actualSha256 });
  return loaded.pdf;
}

function referenceArtifact(reference, testCase, label) {
  const rule = NATIVE_FORMAT_RULES[testCase.format];
  if (rule.unitType === "sheet") {
    if (reference.captureSemantics?.capture !== "sheet-viewport"
      || reference.captureSemantics?.unitType !== "sheet"
      || reference.captureSemantics?.unitIndex !== testCase.unitIndex
      || reference.captureSemantics?.range !== testCase.sheetRange) {
      throw new Error(`${label} sheet capture semantics do not match the suite unit`);
    }
    assertRecord(reference.viewport, `${label}.viewport`);
    if (reference.viewport.unitIndex !== testCase.unitIndex || reference.viewport.range !== testCase.sheetRange) {
      throw new Error(`${label}.viewport unit index or range does not match the suite unit`);
    }
    if (reference.captureSemantics.nativeApplicationAutomation !== false) {
      throw new Error(`${label} sheet viewport must be a manually reviewed import`);
    }
    validateReviewAttestation(reference, rule, reference.viewport.sha256, label);
    return reference.viewport;
  }
  const capture = reference.captureSemantics?.capture;
  if (reference.captureSemantics?.unitType !== rule.unitType
    || reference.captureSemantics?.rasterDpi !== 96
    || (capture !== rule.capture && capture !== "reviewed-native-pdf-import")) {
    throw new Error(`${label} PDF capture semantics do not match the suite unit`);
  }
  if (capture === rule.capture) {
    if (rule.oracleSuite === "wps-office") {
      throw new Error(`${label} WPS PDF references must be explicitly reviewed native PDF imports`);
    }
    if (rule.capture !== "pdf-export" || reference.captureSemantics.nativeApplicationAutomation !== true) {
      throw new Error(`${label} automated PDF export must declare nativeApplicationAutomation true`);
    }
  } else {
    if (reference.captureSemantics.nativeApplicationAutomation !== false) {
      throw new Error(`${label} reviewed native PDF import must declare nativeApplicationAutomation false`);
    }
    if (!Array.isArray(reference.pdfs) || reference.pdfs.length !== 1 || !isRecord(reference.pdfs[0])) {
      throw new Error(`${label}.pdfs must contain exactly one manually reviewed PDF`);
    }
    validateReviewAttestation(reference, rule, reference.pdfs[0].sha256, label);
  }
  if (reference.rasterizer?.name !== "pdftoppm" || reference.rasterizer?.dpi !== 96) {
    throw new Error(`${label}.rasterizer must be pdftoppm at 96 DPI`);
  }
  if (!Array.isArray(reference.pages) || reference.pages.length !== testCase.unitCount) {
    throw new Error(`${label}.pages must contain exactly unitCount entries`);
  }
  const pages = new Map(reference.pages.map((page) => [page?.pageIndex, page]));
  if (pages.size !== reference.pages.length || !pages.has(testCase.unitIndex)) {
    throw new Error(`${label}.pages must uniquely cover every unit index`);
  }
  return pages.get(testCase.unitIndex);
}

export async function validateNativeReference(suiteDirectory, suite, testCase, goldenBytes, cache = new Map()) {
  const referencePath = resolveSafeRelativePath(suiteDirectory, testCase.referenceJson, `${testCase.id}.referenceJson`);
  let loaded = cache.get(referencePath);
  if (loaded === undefined) {
    const bytes = await readFile(referencePath);
    let value;
    try {
      value = JSON.parse(bytes.toString("utf8"));
    } catch (cause) {
      throw new Error(`${testCase.id} reference.json is not valid JSON`, { cause });
    }
    loaded = { digest: sha256(bytes), value };
    cache.set(referencePath, loaded);
  }
  if (loaded.digest !== testCase.referenceJsonSha256) {
    throw new Error(`${testCase.id} reference.json SHA-256 does not match the reviewed manifest`);
  }
  const reference = loaded.value;
  const label = `${testCase.id} reference.json`;
  assertRecord(reference, label);
  if (reference.schemaVersion !== 1) throw new Error(`${label}.schemaVersion must be 1`);
  const rule = NATIVE_FORMAT_RULES[testCase.format];
  if (reference.oracleSuite !== rule.oracleSuite) {
    throw new Error(`${label}.oracleSuite must be ${rule.oracleSuite}`);
  }
  if (reference.source?.format !== rule.extension.slice(1)
    || rawSha256(reference.source?.sha256, `${label}.source.sha256`) !== testCase.fixtureSha256) {
    throw new Error(`${label} source format or SHA-256 does not match the fixture`);
  }
  const expectedApplication = suite.oracleFingerprint.applications[rule.application];
  if (reference.referenceApplication?.name !== rule.applicationName
    || reference.referenceApplication?.bundleId !== rule.bundleId
    || reference.referenceApplication?.version !== expectedApplication.version
    || reference.referenceApplication?.build !== expectedApplication.build) {
    throw new Error(`${label} application identity does not match oracleFingerprint`);
  }
  if (rule.capture === "pdf-export"
    && reference.rasterizer?.version !== suite.oracleFingerprint.rasterizer.version) {
    throw new Error(`${label} rasterizer version does not match oracleFingerprint`);
  }
  const pdf = rule.capture === "pdf-export"
    ? await validateReferencePdf(reference, referencePath, loaded, label)
    : undefined;
  const artifact = referenceArtifact(reference, testCase, label);
  assertRecord(artifact, `${label} unit artifact`);
  assertNonEmptyString(artifact.file, `${label} unit artifact.file`);
  if (rawSha256(artifact.sha256, `${label} unit artifact.sha256`) !== testCase.goldenPngSha256
    || sha256(goldenBytes) !== testCase.goldenPngSha256) {
    throw new Error(`${label} unit artifact SHA-256 does not match the golden PNG`);
  }
  const artifactPath = resolveSafeRelativePath(dirname(referencePath), artifact.file, `${label} unit artifact.file`);
  const goldenPath = resolveSafeRelativePath(suiteDirectory, testCase.goldenPng, `${testCase.id}.goldenPng`);
  if (artifactPath !== goldenPath) throw new Error(`${label} unit artifact path does not match goldenPng`);
  return Object.freeze({
    file: testCase.referenceJson,
    sha256: testCase.referenceJsonSha256,
    application: reference.referenceApplication,
    capture: reference.captureSemantics,
    reviewAttestation: reference.reviewAttestation,
    rasterizer: reference.rasterizer,
    pdf,
    artifact: { file: artifact.file, sha256: testCase.goldenPngSha256 },
  });
}

export function changedCandidateFingerprintFields(expected, actual) {
  validateCandidateFingerprint(expected, "candidateFingerprint");
  validateCandidateFingerprint(actual, "currentCandidateFingerprint", true);
  return CANDIDATE_FINGERPRINT_FIELDS.filter((field) => JSON.stringify(expected[field]) !== JSON.stringify(actual[field]));
}

export function sha256(bytes) {
  return `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
}
