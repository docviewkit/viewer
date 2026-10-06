import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, resolve } from "node:path";
import test from "node:test";

import {
  NATIVE_FORMAT_RULES,
  NATIVE_REFERENCE_REVIEW_SCOPE,
  NATIVE_VISUAL_POLICY,
  NATIVE_VISUAL_POLICY_ID,
  flattenNativeVisualCases,
  sha256,
  validateNativeReference,
  validateNativeVisualSuite,
} from "../scripts/native-visual-policy.mjs";

const digest = (value = "0") => `sha256:${value.repeat(64)}`;
const REFERENCE_PDF_BYTES = Buffer.from("%PDF-1.7\nreviewed-reference\n", "ascii");
const REFERENCE_PDF_SHA256 = sha256(REFERENCE_PDF_BYTES).slice(7);

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
    fontSetDigest: digest(),
    browser: "Chromium",
    browserVersion: "150.0.0.0",
  };
}

function documentFor(format, id = format) {
  const rule = NATIVE_FORMAT_RULES[format];
  return {
    id,
    corpusClass: "minimal",
    format,
    fixture: `fixtures/${id}${rule.extension}`,
    fixtureSha256: digest("1"),
    unitCount: 1,
    units: [{
      index: 0,
      ...(rule.unitType === "sheet" ? { sheetRange: "A1:H40" } : {}),
      referenceJson: `goldens/${id}/reference.json`,
      referenceJsonSha256: digest("3"),
      goldenPng: `goldens/${id}-01.png`,
      goldenPngSha256: digest("2"),
      actualPng: `candidate/${id}-01.png`,
      actualObservationJson: `candidate/${id}-01.json`,
    }],
  };
}

function oracleFingerprint(formats) {
  const applications = {};
  for (const format of formats) {
    const rule = NATIVE_FORMAT_RULES[format];
    applications[rule.application] = { version: "1.0", build: "100", capture: rule.capture };
  }
  return {
    os: "macOS",
    osVersion: "26.0",
    architecture: "arm64",
    locale: "zh-CN",
    timezone: "Asia/Shanghai",
    colorSpace: "srgb",
    scale: 1,
    background: "#ffffff",
    fontSetDigest: digest(),
    applications,
    ...(formats.some((format) => NATIVE_FORMAT_RULES[format].capture === "pdf-export")
      ? { rasterizer: { name: "pdftoppm", version: "26.05.0", dpi: 96 } }
      : {}),
  };
}

function suiteFor(oracleSuite, formats, suiteKind = "regression") {
  const suite = {
    schemaVersion: 2,
    suiteKind,
    oracleMode: "read-only",
    oracleSuite,
    thresholdPolicy: NATIVE_VISUAL_POLICY_ID,
    oracleFingerprint: oracleFingerprint(formats),
    candidateFingerprint: candidateFingerprint(),
    documents: formats.map((format) => documentFor(format)),
  };
  if (suiteKind === "release") {
    const classes = ["minimal", "combination", "enterprise", "large"];
    while (suite.documents.length < classes.length) {
      suite.documents.push(documentFor(formats[0], `${formats[0]}-${suite.documents.length + 1}`));
    }
    classes.forEach((corpusClass, index) => { suite.documents[index].corpusClass = corpusClass; });
  }
  return suite;
}

async function writeReference(directory, suite, reference, goldenBytes, pdfBytes = REFERENCE_PDF_BYTES) {
  const unit = suite.documents[0].units[0];
  unit.referenceJson = `goldens/${suite.documents[0].id}/reference.json`;
  unit.goldenPng = `goldens/${suite.documents[0].id}/golden.png`;
  unit.goldenPngSha256 = sha256(goldenBytes);
  const referencePath = resolve(directory, unit.referenceJson);
  await mkdir(dirname(referencePath), { recursive: true });
  const bytes = Buffer.from(`${JSON.stringify(reference, null, 2)}\n`);
  await writeFile(referencePath, bytes);
  await writeFile(resolve(directory, unit.goldenPng), goldenBytes);
  if (reference.pdfs?.[0]?.file === "reference.pdf" && pdfBytes !== undefined) {
    await writeFile(resolve(dirname(referencePath), "reference.pdf"), pdfBytes);
  }
  unit.referenceJsonSha256 = sha256(bytes);
  return flattenNativeVisualCases(suite)[0];
}

function officeReference(suite, goldenBytes) {
  const format = suite.documents[0].format;
  const rule = NATIVE_FORMAT_RULES[format];
  const reference = {
    schemaVersion: 1,
    oracleSuite: rule.oracleSuite,
    generatedAt: "2026-07-19T00:00:00.000Z",
    source: { format, sha256: "1".repeat(64) },
    referenceApplication: {
      name: rule.applicationName,
      bundleId: rule.bundleId,
      version: "1.0",
      build: "100",
    },
    captureSemantics: {
      unitType: rule.unitType,
      capture: rule.capture,
    },
  };
  if (rule.unitType === "sheet") {
    Object.assign(reference.captureSemantics, {
      unitIndex: 0,
      range: "A1:H40",
      nativeApplicationAutomation: false,
    });
    reference.viewport = {
      unitIndex: 0,
      range: "A1:H40",
      file: "golden.png",
      sha256: sha256(goldenBytes).slice(7),
    };
    reference.reviewAttestation = reviewAttestation(reference, rule, reference.viewport.sha256);
  } else {
    reference.captureSemantics.nativeApplicationAutomation = true;
    reference.captureSemantics.rasterDpi = 96;
    reference.rasterizer = { name: "pdftoppm", version: "26.05.0", dpi: 96 };
    reference.pdfs = [{
      file: "reference.pdf",
      bytes: REFERENCE_PDF_BYTES.length,
      sha256: REFERENCE_PDF_SHA256,
    }];
    reference.pages = [{ pageIndex: 0, file: "golden.png", sha256: sha256(goldenBytes).slice(7) }];
  }
  return reference;
}

function reviewAttestation(reference, rule, artifactSha256) {
  return {
    status: "reviewed",
    reviewer: "reviewer-id",
    reviewedAt: reference.generatedAt,
    scope: NATIVE_REFERENCE_REVIEW_SCOPE,
    oracleSuite: rule.oracleSuite,
    application: rule.applicationName,
    sourceSha256: reference.source.sha256,
    artifactSha256,
  };
}

function manualPdfReference(suite, goldenBytes) {
  const reference = officeReference(suite, goldenBytes);
  const rule = NATIVE_FORMAT_RULES[suite.documents[0].format];
  reference.captureSemantics.capture = "reviewed-native-pdf-import";
  reference.captureSemantics.nativeApplicationAutomation = false;
  reference.reviewAttestation = reviewAttestation(reference, rule, reference.pdfs[0].sha256);
  return reference;
}

test("native visual suites enforce the Office, iWork, and WPS oracle mapping", () => {
  const officeFormats = ["pptx", "ppt", "docx", "doc", "xlsx", "xls", "odp", "odt", "ods"];
  const office = suiteFor("microsoft-office", officeFormats, "release");
  const iwork = suiteFor("apple-iwork", ["keynote", "pages", "numbers"], "release");
  const wps = suiteFor("wps-office", ["wps", "et", "dps"], "release");
  assert.equal(validateNativeVisualSuite(office), office);
  assert.equal(validateNativeVisualSuite(iwork), iwork);
  assert.equal(validateNativeVisualSuite(wps), wps);
  assert.deepEqual(flattenNativeVisualCases(iwork).slice(0, 3).map(({ format, oracleApplication, unitType }) => (
    [format, oracleApplication, unitType]
  )), [
    ["keynote", "keynote", "slide"],
    ["pages", "pages", "page"],
    ["numbers", "numbers", "sheet"],
  ]);
  assert.deepEqual(flattenNativeVisualCases(office).slice(-3).map(({
    format, oracleApplication, unitType, capture,
  }) => [format, oracleApplication, unitType, capture]), [
    ["odp", "powerpoint", "slide", "pdf-export"],
    ["odt", "word", "page", "pdf-export"],
    ["ods", "excel", "sheet", "sheet-viewport"],
  ]);
  assert.deepEqual(flattenNativeVisualCases(wps).slice(0, 3).map(({
    format, oracleApplication, unitType, capture,
  }) => [format, oracleApplication, unitType, capture]), [
    ["wps", "wps-writer", "page", "pdf-export"],
    ["et", "wps-spreadsheets", "sheet", "sheet-viewport"],
    ["dps", "wps-presentation", "slide", "pdf-export"],
  ]);
  assert.deepEqual([
    NATIVE_FORMAT_RULES.pptx.bundleId,
    NATIVE_FORMAT_RULES.docx.bundleId,
    NATIVE_FORMAT_RULES.xlsx.bundleId,
    NATIVE_FORMAT_RULES.keynote.bundleId,
    NATIVE_FORMAT_RULES.pages.bundleId,
    NATIVE_FORMAT_RULES.numbers.bundleId,
    NATIVE_FORMAT_RULES.wps.bundleId,
    NATIVE_FORMAT_RULES.et.bundleId,
    NATIVE_FORMAT_RULES.dps.bundleId,
  ], [
    "com.microsoft.Powerpoint",
    "com.microsoft.Word",
    "com.microsoft.Excel",
    "com.apple.iWork.Keynote",
    "com.apple.iWork.Pages",
    "com.apple.iWork.Numbers",
    "com.kingsoft.wpsoffice.mac",
    "com.kingsoft.wpsoffice.mac",
    "com.kingsoft.wpsoffice.mac",
  ]);

  const families = [
    ["microsoft-office", office],
    ["apple-iwork", iwork],
    ["wps-office", wps],
  ];
  for (const [source, suite] of families) {
    for (const [target] of families) {
      if (target === source) continue;
      const crossed = structuredClone(suite);
      crossed.oracleSuite = target;
      assert.throws(() => validateNativeVisualSuite(crossed), new RegExp(`belongs to ${source}`, "u"));
    }
  }

  const forbidden = structuredClone(office);
  forbidden.oracleSuite = "libreoffice";
  assert.throws(
    () => validateNativeVisualSuite(forbidden),
    /oracleSuite must be microsoft-office, apple-iwork, or wps-office/u,
  );
});

test("native visual release suites require every format in their oracle family", () => {
  const incomplete = suiteFor("microsoft-office", ["pptx"], "release");
  assert.throws(
    () => validateNativeVisualSuite(incomplete),
    /missing formats: ppt, docx, doc, xlsx, xls, odp, odt, ods/u,
  );

  const shallow = suiteFor("apple-iwork", ["keynote", "pages", "numbers"]);
  shallow.suiteKind = "release";
  assert.throws(() => validateNativeVisualSuite(shallow), /missing trusted corpus classes/u);

  const incompleteWps = suiteFor("wps-office", ["wps"], "release");
  assert.throws(
    () => validateNativeVisualSuite(incompleteWps),
    /missing formats: et, dps/u,
  );
});

test("native visual manifests require complete units and deterministic sheet ranges", () => {
  const missingUnit = suiteFor("microsoft-office", ["pptx"]);
  missingUnit.documents[0].unitCount = 2;
  assert.throws(() => validateNativeVisualSuite(missingUnit), /exactly unitCount entries/u);

  const unboundedSheet = suiteFor("microsoft-office", ["xlsx"]);
  delete unboundedSheet.documents[0].units[0].sheetRange;
  assert.throws(() => validateNativeVisualSuite(unboundedSheet), /sheetRange/u);

  const unboundedOds = suiteFor("microsoft-office", ["ods"]);
  delete unboundedOds.documents[0].units[0].sheetRange;
  assert.throws(() => validateNativeVisualSuite(unboundedOds), /sheetRange/u);

  const pageWithRange = suiteFor("apple-iwork", ["pages"]);
  pageWithRange.documents[0].units[0].sheetRange = "A1:H40";
  assert.throws(() => validateNativeVisualSuite(pageWithRange), /valid only for sheet formats/u);

  for (const invalidRange of ["H1:A40", "A40:H1", "XFE1:XFE2", "A1048577:A1048577"]) {
    const invalidSheet = suiteFor("microsoft-office", ["xlsx"]);
    invalidSheet.documents[0].units[0].sheetRange = invalidRange;
    assert.throws(() => validateNativeVisualSuite(invalidSheet), /A1:XFD1048576|top-left/u);
  }
});

test("native visual policy is versioned, immutable, and cannot be overridden per suite or unit", () => {
  assert.equal(Object.isFrozen(NATIVE_VISUAL_POLICY), true);
  assert.deepEqual(NATIVE_VISUAL_POLICY, {
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
  const suitePolicy = suiteFor("microsoft-office", ["pptx"]);
  suitePolicy.policy = { minSsim: 0 };
  assert.throws(() => validateNativeVisualSuite(suitePolicy), /policy/u);

  const unitPolicy = suiteFor("microsoft-office", ["pptx"]);
  unitPolicy.documents[0].units[0].policy = { minOverallScore: 0 };
  assert.throws(() => validateNativeVisualSuite(unitPolicy), /policy/u);

  const scaledCandidate = suiteFor("microsoft-office", ["pptx"]);
  scaledCandidate.candidateFingerprint.scale = 2;
  assert.throws(() => validateNativeVisualSuite(scaledCandidate), /candidateFingerprint\.scale must be 1/u);

  const nonSrgbOracle = suiteFor("microsoft-office", ["pptx"]);
  nonSrgbOracle.oracleFingerprint.colorSpace = "display-p3";
  assert.throws(() => validateNativeVisualSuite(nonSrgbOracle), /oracleFingerprint\.colorSpace must be srgb/u);

  const wrongOfficeCapture = suiteFor("microsoft-office", ["ods"]);
  wrongOfficeCapture.oracleFingerprint.applications.excel.capture = "pdf-export";
  assert.throws(
    () => validateNativeVisualSuite(wrongOfficeCapture),
    /oracleFingerprint\.applications\.excel\.capture must be sheet-viewport/u,
  );
});

test("native visual manifests reject path traversal and extension aliases", () => {
  const traversal = suiteFor("microsoft-office", ["docx"]);
  traversal.documents[0].fixture = "../sample.docx";
  assert.throws(() => validateNativeVisualSuite(traversal), /safe relative path|inside its configured root/u);

  const wrongExtension = suiteFor("microsoft-office", ["docx"]);
  wrongExtension.documents[0].fixture = "fixtures/sample.doc";
  assert.throws(() => validateNativeVisualSuite(wrongExtension), /extension does not match/u);
});

test("PDF references require explicit automated export or a reviewed manual import", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "officeviewer-pdf-policy-"));
  const goldenBytes = Buffer.from("reviewed-golden");
  const suite = suiteFor("microsoft-office", ["odp"]);

  try {
    const automated = officeReference(suite, goldenBytes);
    let testCase = await writeReference(directory, suite, automated, goldenBytes);
    await validateNativeReference(directory, suite, testCase, goldenBytes);

    delete automated.captureSemantics.capture;
    testCase = await writeReference(directory, suite, automated, goldenBytes);
    await assert.rejects(validateNativeReference(directory, suite, testCase, goldenBytes), /PDF capture semantics/u);

    automated.captureSemantics.capture = "pdf-export";
    automated.captureSemantics.nativeApplicationAutomation = false;
    testCase = await writeReference(directory, suite, automated, goldenBytes);
    await assert.rejects(
      validateNativeReference(directory, suite, testCase, goldenBytes),
      /nativeApplicationAutomation true/u,
    );

    const manual = manualPdfReference(suite, goldenBytes);
    delete manual.reviewAttestation;
    testCase = await writeReference(directory, suite, manual, goldenBytes);
    await assert.rejects(validateNativeReference(directory, suite, testCase, goldenBytes), /reviewAttestation/u);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("WPS page and slide references require reviewed imports from WPS Office", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "officeviewer-wps-policy-"));
  const goldenBytes = Buffer.from("reviewed-wps-golden");
  const suite = suiteFor("wps-office", ["wps"]);

  try {
    const automated = officeReference(suite, goldenBytes);
    let testCase = await writeReference(directory, suite, automated, goldenBytes);
    await assert.rejects(
      validateNativeReference(directory, suite, testCase, goldenBytes),
      /WPS PDF references must be explicitly reviewed/u,
    );

    const manual = manualPdfReference(suite, goldenBytes);
    testCase = await writeReference(directory, suite, manual, goldenBytes);
    const report = await validateNativeReference(directory, suite, testCase, goldenBytes);
    assert.equal(report.application.name, "WPS Office");
    assert.equal(report.reviewAttestation.oracleSuite, "wps-office");
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("PDF references authenticate the physical PDF once per shared reference", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "officeviewer-pdf-bytes-"));
  const goldenBytes = Buffer.from("reviewed-golden");
  const suite = suiteFor("microsoft-office", ["odp"]);
  const reference = officeReference(suite, goldenBytes);

  try {
    let testCase = await writeReference(directory, suite, reference, goldenBytes);
    const pdfPath = resolve(directory, "goldens/odp/reference.pdf");
    const cache = new Map();
    const report = await validateNativeReference(directory, suite, testCase, goldenBytes, cache);
    assert.deepEqual(report.pdf, {
      file: "reference.pdf",
      bytes: REFERENCE_PDF_BYTES.length,
      sha256: sha256(REFERENCE_PDF_BYTES),
    });

    await rm(pdfPath);
    await validateNativeReference(directory, suite, testCase, goldenBytes, cache);
    await assert.rejects(
      validateNativeReference(directory, suite, testCase, goldenBytes, new Map()),
      /PDF could not be read/u,
    );

    await writeFile(pdfPath, Buffer.alloc(0));
    await assert.rejects(validateNativeReference(directory, suite, testCase, goldenBytes), /must not be empty/u);

    const invalidSignature = Buffer.from(REFERENCE_PDF_BYTES);
    invalidSignature.set(Buffer.from("NOTPD", "ascii"), 0);
    await writeFile(pdfPath, invalidSignature);
    await assert.rejects(validateNativeReference(directory, suite, testCase, goldenBytes), /PDF signature/u);

    await writeFile(pdfPath, Buffer.concat([REFERENCE_PDF_BYTES, Buffer.from("x")]));
    await assert.rejects(validateNativeReference(directory, suite, testCase, goldenBytes), /byte length/u);

    const replacement = Buffer.from(REFERENCE_PDF_BYTES);
    replacement[replacement.length - 2] ^= 1;
    await writeFile(pdfPath, replacement);
    await assert.rejects(validateNativeReference(directory, suite, testCase, goldenBytes), /PDF SHA-256/u);

    for (const [mutate, expected] of [
      [(value) => { value.pdfs = []; }, /exactly one PDF record/u],
      [(value) => { value.pdfs[0].file = "../reference.pdf"; }, /safe relative path/u],
      [(value) => { value.pdfs[0].file = "reference.bin"; }, /must end with \.pdf/u],
      [(value) => { value.pdfs[0].bytes = 0; }, /positive integer/u],
      [(value) => { value.pdfs[0].sha256 = "invalid"; }, /raw lowercase SHA-256/u],
      [(value) => { value.pdfs[0].sha256 = "0".repeat(64); }, /PDF SHA-256/u],
    ]) {
      const invalid = structuredClone(reference);
      mutate(invalid);
      testCase = await writeReference(directory, suite, invalid, goldenBytes);
      await assert.rejects(validateNativeReference(directory, suite, testCase, goldenBytes), expected);
    }
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("manual PDF and sheet references require review attestations bound to their source and artifact", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "officeviewer-review-policy-"));
  const goldenBytes = Buffer.from("reviewed-golden");

  try {
    const pdfSuite = suiteFor("microsoft-office", ["odp"]);
    const pdfReference = manualPdfReference(pdfSuite, goldenBytes);
    let testCase = await writeReference(directory, pdfSuite, pdfReference, goldenBytes);
    await validateNativeReference(directory, pdfSuite, testCase, goldenBytes);

    const mutations = [
      [(reference) => { delete reference.reviewAttestation; }, /reviewAttestation/u],
      [(reference) => { reference.reviewAttestation.status = "pending"; }, /status/u],
      [(reference) => { reference.reviewAttestation.reviewer = "reviewer id"; }, /reviewer/u],
      [(reference) => { reference.reviewAttestation.reviewer = "审核员"; }, /reviewer/u],
      [(reference) => { reference.reviewAttestation.reviewer = "-reviewer"; }, /reviewer/u],
      [(reference) => { reference.reviewAttestation.reviewedAt = "2026-07-19T00:00:00Z"; }, /canonical/u],
      [(reference) => { reference.reviewAttestation.reviewedAt = "2026-07-19T00:00:01.000Z"; }, /equal generatedAt/u],
      [(reference) => { reference.reviewAttestation.scope = "partial"; }, /scope/u],
      [(reference) => { reference.reviewAttestation.oracleSuite = "libreoffice"; }, /oracle suite or application/u],
      [(reference) => { reference.reviewAttestation.application = "LibreOffice"; }, /oracle suite or application/u],
      [(reference) => { reference.reviewAttestation.sourceSha256 = "8".repeat(64); }, /sourceSha256/u],
      [(reference) => { reference.reviewAttestation.artifactSha256 = "9".repeat(64); }, /artifactSha256/u],
    ];
    for (const [mutate, expected] of mutations) {
      const invalid = structuredClone(pdfReference);
      mutate(invalid);
      testCase = await writeReference(directory, pdfSuite, invalid, goldenBytes);
      await assert.rejects(validateNativeReference(directory, pdfSuite, testCase, goldenBytes), expected);
    }

    const sheetSuite = suiteFor("microsoft-office", ["ods"]);
    const sheetReference = officeReference(sheetSuite, goldenBytes);
    testCase = await writeReference(directory, sheetSuite, sheetReference, goldenBytes);
    await validateNativeReference(directory, sheetSuite, testCase, goldenBytes);

    delete sheetReference.reviewAttestation;
    testCase = await writeReference(directory, sheetSuite, sheetReference, goldenBytes);
    await assert.rejects(validateNativeReference(directory, sheetSuite, testCase, goldenBytes), /reviewAttestation/u);

    sheetReference.reviewAttestation = reviewAttestation(
      sheetReference,
      NATIVE_FORMAT_RULES.ods,
      "9".repeat(64),
    );
    testCase = await writeReference(directory, sheetSuite, sheetReference, goldenBytes);
    await assert.rejects(validateNativeReference(directory, sheetSuite, testCase, goldenBytes), /artifactSha256/u);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("ODF references require Microsoft Office and reject LibreOffice as an oracle", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "officeviewer-odf-policy-"));
  const goldenBytes = Buffer.from("reviewed-golden");

  try {
    for (const format of ["odp", "odt", "ods"]) {
      const suite = suiteFor("microsoft-office", [format]);
      const reference = officeReference(suite, goldenBytes);
      let testCase = await writeReference(directory, suite, reference, goldenBytes);
      const report = await validateNativeReference(directory, suite, testCase, goldenBytes);
      assert.equal(report.application.name, NATIVE_FORMAT_RULES[format].applicationName);

      delete reference.oracleSuite;
      testCase = await writeReference(directory, suite, reference, goldenBytes);
      await assert.rejects(
        validateNativeReference(directory, suite, testCase, goldenBytes),
        /oracleSuite must be microsoft-office/u,
      );

      reference.oracleSuite = "libreoffice";
      testCase = await writeReference(directory, suite, reference, goldenBytes);
      await assert.rejects(
        validateNativeReference(directory, suite, testCase, goldenBytes),
        /oracleSuite must be microsoft-office/u,
      );

      reference.oracleSuite = "microsoft-office";
      reference.referenceApplication.bundleId = "org.libreoffice.LibreOffice";
      testCase = await writeReference(directory, suite, reference, goldenBytes);
      await assert.rejects(
        validateNativeReference(directory, suite, testCase, goldenBytes),
        /application identity/u,
      );

      reference.referenceApplication.bundleId = NATIVE_FORMAT_RULES[format].bundleId;
      reference.referenceApplication.name = "LibreOffice";
      testCase = await writeReference(directory, suite, reference, goldenBytes);
      await assert.rejects(
        validateNativeReference(directory, suite, testCase, goldenBytes),
        /application identity/u,
      );
    }
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
