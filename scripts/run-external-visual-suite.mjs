import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { inflateSync } from "node:zlib";

import { evaluateAccuracy } from "../dist/accuracy.js";
import {
  NATIVE_VISUAL_POLICY,
  changedCandidateFingerprintFields,
  flattenNativeVisualCases,
  resolveSafeRelativePath,
  sha256,
  validateCandidateFingerprint,
  validateNativeReference,
  validateNativeVisualSuite,
} from "./native-visual-policy.mjs";

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const MAX_PIXELS = 100_000_000;
const VISUAL_COVERAGE = ["structure-units", "visual-exact", "visual-tolerant", "visual-ssim"];
const CORPUS_CLASSES = new Set(["minimal", "combination", "enterprise", "large", "malformed", "malicious"]);
const UNIT_TYPES = new Set(["slide", "sheet", "page"]);
const FINGERPRINT_FIELDS = [
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
  "referenceRenderer",
  "referenceRendererVersion",
  "candidateRenderer",
  "candidateRendererVersion",
];

function usage() {
  return "Usage: node scripts/run-external-visual-suite.mjs <suite.json> --fingerprint <current-environment.json> [--actual-root <directory>] [--output <directory>]";
}

function parseArguments(arguments_) {
  let suitePath;
  let fingerprintPath;
  let actualRoot;
  let output = "output/accuracy/external";
  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index];
    if (argument === "--fingerprint") {
      fingerprintPath = arguments_[++index];
      if (fingerprintPath === undefined) throw new Error(`${usage()}\nMissing --fingerprint value`);
    } else if (argument === "--output") {
      output = arguments_[++index];
      if (output === undefined) throw new Error(`${usage()}\nMissing --output value`);
    } else if (argument === "--actual-root") {
      actualRoot = arguments_[++index];
      if (actualRoot === undefined) throw new Error(`${usage()}\nMissing --actual-root value`);
    } else if (argument.startsWith("-")) {
      throw new Error(`${usage()}\nUnknown option ${argument}`);
    } else if (suitePath === undefined) {
      suitePath = argument;
    } else {
      throw new Error(`${usage()}\nUnexpected argument ${argument}`);
    }
  }
  if (suitePath === undefined || fingerprintPath === undefined) throw new Error(usage());
  return {
    suitePath: resolve(suitePath),
    fingerprintPath: resolve(fingerprintPath),
    actualRoot: actualRoot === undefined ? undefined : resolve(actualRoot),
    output: resolve(output),
  };
}

function safeName(id) {
  const value = id.normalize("NFKC").replace(/[^A-Za-z0-9._-]+/gu, "-").replace(/^-+|-+$/gu, "");
  if (value === "" || value === "." || value === "..") throw new Error(`Case id ${JSON.stringify(id)} cannot form a report filename`);
  return value;
}

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function canonical(value) {
  if (Array.isArray(value)) return `[${value.map(canonical).join(",")}]`;
  if (isRecord(value)) {
    return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical(value[key])}`).join(",")}}`;
  }
  return JSON.stringify(value);
}

function validateFingerprint(fingerprint, label) {
  if (!isRecord(fingerprint)) throw new Error(`${label} must be a JSON object`);
  for (const field of FINGERPRINT_FIELDS) {
    const value = fingerprint[field];
    if (field === "devicePixelRatio" || field === "scale") {
      if (!Number.isFinite(value) || value <= 0) throw new Error(`${label}.${field} must be a finite positive number`);
    } else if (typeof value !== "string" || value.trim() === "") {
      throw new Error(`${label}.${field} must be a non-empty string`);
    }
  }
  if (!/^sha256:[0-9a-f]{64}$/u.test(fingerprint.fontSetDigest)) {
    throw new Error(`${label}.fontSetDigest must be a lowercase SHA-256 digest`);
  }
}

function validateLegacySuite(suite) {
  if (!isRecord(suite) || suite.schemaVersion !== 1) throw new Error("External visual suite schemaVersion must be 1");
  if (suite.oracleMode !== "read-only") {
    throw new Error("External visual suite oracleMode must be read-only; application chrome and editing placeholders are not visual oracles");
  }
  validateFingerprint(suite.environmentFingerprint, "environmentFingerprint");
  if (!Array.isArray(suite.cases) || suite.cases.length === 0) throw new Error("External visual suite must contain at least one case");
  const ids = new Set();
  const reportNames = new Set();
  for (const testCase of suite.cases) {
    if (!isRecord(testCase) || typeof testCase.id !== "string" || testCase.id.trim() === "") throw new Error("Every external visual case must have a non-empty id");
    if (ids.has(testCase.id)) throw new Error(`External visual case ${testCase.id} is duplicated`);
    ids.add(testCase.id);
    const reportName = safeName(testCase.id);
    if (reportNames.has(reportName)) throw new Error(`External visual case ${testCase.id} collides with another report filename`);
    reportNames.add(reportName);
    if (!CORPUS_CLASSES.has(testCase.corpusClass)) throw new Error(`External visual case ${testCase.id} has an invalid corpusClass`);
    if (!UNIT_TYPES.has(testCase.unitType)) throw new Error(`External visual case ${testCase.id} must declare unitType slide, sheet, or page`);
    if (!Number.isInteger(testCase.unitIndex) || testCase.unitIndex < 0) throw new Error(`External visual case ${testCase.id} has an invalid unitIndex`);
    for (const field of ["fixture", "goldenPng", "actualPng"]) {
      if (typeof testCase[field] !== "string" || testCase[field].trim() === "") throw new Error(`External visual case ${testCase.id}.${field} must be a non-empty string`);
    }
    const hasExpectedObservation = testCase.expectedObservationJson !== undefined;
    const hasActualObservation = testCase.actualObservationJson !== undefined;
    if (hasExpectedObservation !== hasActualObservation) {
      throw new Error(`External visual case ${testCase.id} must pair expectedObservationJson with actualObservationJson`);
    }
    for (const field of ["expectedObservationJson", "actualObservationJson"]) {
      if (testCase[field] !== undefined && (typeof testCase[field] !== "string" || testCase[field].trim() === "")) {
        throw new Error(`External visual case ${testCase.id}.${field} must be a non-empty string`);
      }
    }
    if (testCase.declaredCoverage !== undefined
      && (!Array.isArray(testCase.declaredCoverage)
        || testCase.declaredCoverage.some((coverage) => typeof coverage !== "string" || coverage === ""))) {
      throw new Error(`External visual case ${testCase.id}.declaredCoverage must be an array of coverage names`);
    }
    if (testCase.policy !== undefined && !isRecord(testCase.policy)) throw new Error(`External visual case ${testCase.id}.policy must be an object`);
  }
}

function changedFingerprintFields(expected, actual) {
  return [...new Set([...Object.keys(expected), ...Object.keys(actual)])]
    .filter((field) => canonical(expected[field]) !== canonical(actual[field]))
    .sort();
}

function crc32(bytes) {
  let crc = 0xffff_ffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb8_8320 & -(crc & 1));
  }
  return (crc ^ 0xffff_ffff) >>> 0;
}

function paeth(left, up, upperLeft) {
  const prediction = left + up - upperLeft;
  const leftDistance = Math.abs(prediction - left);
  const upDistance = Math.abs(prediction - up);
  const upperLeftDistance = Math.abs(prediction - upperLeft);
  if (leftDistance <= upDistance && leftDistance <= upperLeftDistance) return left;
  return upDistance <= upperLeftDistance ? up : upperLeft;
}

function decodePng(bytes, label) {
  if (bytes.length < PNG_SIGNATURE.length || !bytes.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)) {
    throw new Error(`${label} is not a PNG file`);
  }
  let offset = PNG_SIGNATURE.length;
  let header;
  let palette;
  let transparency;
  let ended = false;
  const compressed = [];
  while (offset < bytes.length) {
    if (offset + 12 > bytes.length) throw new Error(`${label} has a truncated PNG chunk`);
    const length = bytes.readUInt32BE(offset);
    const end = offset + 12 + length;
    if (end > bytes.length) throw new Error(`${label} has an out-of-bounds PNG chunk`);
    const typeBytes = bytes.subarray(offset + 4, offset + 8);
    const type = typeBytes.toString("ascii");
    const data = bytes.subarray(offset + 8, offset + 8 + length);
    const expectedCrc = bytes.readUInt32BE(offset + 8 + length);
    const actualCrc = crc32(Buffer.concat([typeBytes, data]));
    if (expectedCrc !== actualCrc) throw new Error(`${label} PNG chunk ${type} has an invalid CRC`);
    if (type === "IHDR") {
      if (header !== undefined || length !== 13) throw new Error(`${label} has an invalid PNG header`);
      header = {
        width: data.readUInt32BE(0),
        height: data.readUInt32BE(4),
        bitDepth: data[8],
        colorType: data[9],
        compression: data[10],
        filter: data[11],
        interlace: data[12],
      };
    } else if (type === "PLTE") {
      palette = data;
    } else if (type === "tRNS") {
      transparency = data;
    } else if (type === "IDAT") {
      compressed.push(data);
    } else if (type === "IEND") {
      ended = true;
      offset = end;
      break;
    }
    offset = end;
  }
  if (!ended || header === undefined || compressed.length === 0) throw new Error(`${label} is missing required PNG chunks`);
  if (offset !== bytes.length) throw new Error(`${label} contains data after the PNG end chunk`);
  const { width, height, bitDepth, colorType, compression, filter, interlace } = header;
  if (!Number.isInteger(width) || !Number.isInteger(height) || width <= 0 || height <= 0 || width * height > MAX_PIXELS) {
    throw new Error(`${label} PNG dimensions are invalid or exceed ${MAX_PIXELS} pixels`);
  }
  if (bitDepth !== 8 || compression !== 0 || filter !== 0 || interlace !== 0) {
    throw new Error(`${label} must be a non-interlaced 8-bit PNG`);
  }
  const channels = new Map([[0, 1], [2, 3], [3, 1], [4, 2], [6, 4]]).get(colorType);
  if (channels === undefined) throw new Error(`${label} uses unsupported PNG color type ${colorType}`);
  if (colorType === 3 && (palette === undefined || palette.length === 0 || palette.length % 3 !== 0)) {
    throw new Error(`${label} has an invalid indexed-color palette`);
  }
  const stride = width * channels;
  const expectedInflatedLength = height * (stride + 1);
  const inflated = inflateSync(Buffer.concat(compressed), { maxOutputLength: expectedInflatedLength });
  if (inflated.length !== expectedInflatedLength) throw new Error(`${label} PNG scanline length is invalid`);
  const unfiltered = Buffer.alloc(height * stride);
  for (let y = 0; y < height; y += 1) {
    const inputRow = y * (stride + 1);
    const outputRow = y * stride;
    const filterType = inflated[inputRow];
    if (filterType > 4) throw new Error(`${label} uses invalid PNG filter ${filterType}`);
    for (let x = 0; x < stride; x += 1) {
      const encoded = inflated[inputRow + 1 + x];
      const left = x >= channels ? unfiltered[outputRow + x - channels] : 0;
      const up = y > 0 ? unfiltered[outputRow - stride + x] : 0;
      const upperLeft = y > 0 && x >= channels ? unfiltered[outputRow - stride + x - channels] : 0;
      const predictor = filterType === 0 ? 0
        : filterType === 1 ? left
          : filterType === 2 ? up
            : filterType === 3 ? Math.floor((left + up) / 2)
              : paeth(left, up, upperLeft);
      unfiltered[outputRow + x] = (encoded + predictor) & 0xff;
    }
  }
  const rgba = new Uint8Array(width * height * 4);
  for (let pixel = 0; pixel < width * height; pixel += 1) {
    const input = pixel * channels;
    const output = pixel * 4;
    if (colorType === 6) {
      rgba.set(unfiltered.subarray(input, input + 4), output);
    } else if (colorType === 2) {
      rgba[output] = unfiltered[input];
      rgba[output + 1] = unfiltered[input + 1];
      rgba[output + 2] = unfiltered[input + 2];
      const transparent = transparency?.length === 6
        && transparency.readUInt16BE(0) === unfiltered[input]
        && transparency.readUInt16BE(2) === unfiltered[input + 1]
        && transparency.readUInt16BE(4) === unfiltered[input + 2];
      rgba[output + 3] = transparent ? 0 : 255;
    } else if (colorType === 3) {
      const index = unfiltered[input];
      if (index * 3 + 2 >= palette.length) throw new Error(`${label} references an invalid PNG palette index`);
      rgba[output] = palette[index * 3];
      rgba[output + 1] = palette[index * 3 + 1];
      rgba[output + 2] = palette[index * 3 + 2];
      rgba[output + 3] = transparency?.[index] ?? 255;
    } else if (colorType === 4) {
      rgba[output] = unfiltered[input];
      rgba[output + 1] = unfiltered[input];
      rgba[output + 2] = unfiltered[input];
      rgba[output + 3] = unfiltered[input + 1];
    } else {
      rgba[output] = unfiltered[input];
      rgba[output + 1] = unfiltered[input];
      rgba[output + 2] = unfiltered[input];
      const transparent = transparency?.length === 2 && transparency.readUInt16BE(0) === unfiltered[input];
      rgba[output + 3] = transparent ? 0 : 255;
    }
  }
  return { width, height, data: rgba };
}

function digest(bytes) {
  return `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
}

function rounded(value) {
  return Math.round(value * 1_000_000) / 1_000_000;
}

function visualDiff(expected, actual, channelTolerance, pixelRadius = 0) {
  const comparable = expected.width === actual.width && expected.height === actual.height;
  if (!comparable) {
    return {
      comparable: false,
      totalPixels: null,
      differingPixelCount: null,
      tolerantDifferingPixelCount: null,
      exactPixelSimilarity: null,
      tolerantPixelSimilarity: null,
      maxChannelDelta: null,
      meanAbsoluteChannelDelta: null,
      differenceBounds: null,
    };
  }
  const totalPixels = expected.width * expected.height;
  let differingPixelCount = 0;
  let tolerantDifferingPixelCount = 0;
  let maxChannelDelta = 0;
  let absoluteDelta = 0;
  let minX = expected.width;
  let minY = expected.height;
  let maxX = -1;
  let maxY = -1;
  const matchesNeighborhood = (reference, candidate, pixel) => {
    const referenceOffset = pixel * 4;
    const x = pixel % expected.width;
    const y = Math.floor(pixel / expected.width);
    for (let candidateY = Math.max(0, y - pixelRadius);
      candidateY <= Math.min(expected.height - 1, y + pixelRadius);
      candidateY += 1) {
      for (let candidateX = Math.max(0, x - pixelRadius);
        candidateX <= Math.min(expected.width - 1, x + pixelRadius);
        candidateX += 1) {
        const candidateOffset = (candidateY * expected.width + candidateX) * 4;
        let matches = true;
        for (let channel = 0; channel < 4; channel += 1) {
          if (Math.abs(reference.data[referenceOffset + channel] - candidate.data[candidateOffset + channel]) > channelTolerance) {
            matches = false;
            break;
          }
        }
        if (matches) return true;
      }
    }
    return false;
  };
  for (let pixel = 0; pixel < totalPixels; pixel += 1) {
    let differs = false;
    let differsBeyondTolerance = false;
    for (let channel = 0; channel < 4; channel += 1) {
      const offset = pixel * 4 + channel;
      const delta = Math.abs(expected.data[offset] - actual.data[offset]);
      absoluteDelta += delta;
      maxChannelDelta = Math.max(maxChannelDelta, delta);
      if (delta !== 0) differs = true;
      if (delta > channelTolerance) differsBeyondTolerance = true;
    }
    if (differsBeyondTolerance && pixelRadius > 0 && matchesNeighborhood(expected, actual, pixel)) {
      differsBeyondTolerance = false;
    }
    if (differs) {
      differingPixelCount += 1;
      const x = pixel % expected.width;
      const y = Math.floor(pixel / expected.width);
      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }
    if (differsBeyondTolerance) tolerantDifferingPixelCount += 1;
  }
  if (pixelRadius > 0) {
    let reverseTolerantDifferingPixelCount = 0;
    for (let pixel = 0; pixel < totalPixels; pixel += 1) {
      if (!matchesNeighborhood(actual, expected, pixel)) reverseTolerantDifferingPixelCount += 1;
    }
    tolerantDifferingPixelCount = Math.max(
      tolerantDifferingPixelCount,
      reverseTolerantDifferingPixelCount,
    );
  }
  return {
    comparable: true,
    totalPixels,
    differingPixelCount,
    tolerantDifferingPixelCount,
    exactPixelSimilarity: rounded((totalPixels - differingPixelCount) / totalPixels),
    tolerantPixelSimilarity: rounded((totalPixels - tolerantDifferingPixelCount) / totalPixels),
    channelTolerance,
    pixelRadius,
    maxChannelDelta,
    meanAbsoluteChannelDelta: rounded(absoluteDelta / (totalPixels * 4)),
    differenceBounds: differingPixelCount === 0 ? null : {
      x: minX,
      y: minY,
      width: maxX - minX + 1,
      height: maxY - minY + 1,
    },
  };
}

function failedSection() {
  return { score: 0, applicable: true, metrics: {} };
}

function observationFailure(testCase, suite, environmentFingerprint, cause) {
  const native = suite.schemaVersion === 2;
  return {
    schemaVersion: native ? 2 : 1,
    oracle: {
      suite: native ? suite.oracleSuite : undefined,
      renderer: native ? testCase.oracleApplication : environmentFingerprint.referenceRenderer,
      mode: "read-only",
      fixture: testCase.fixture,
      goldenPng: testCase.goldenPng,
    },
    candidate: {
      renderer: native ? environmentFingerprint.browser : environmentFingerprint.candidateRenderer,
      actualPng: testCase.actualPng,
    },
    environmentFingerprint,
    fileId: testCase.id,
    corpusClass: testCase.corpusClass,
    passed: false,
    coverageValidation: { declared: VISUAL_COVERAGE, observed: [], missing: VISUAL_COVERAGE },
    contentCompleteness: failedSection(),
    geometryAccuracy: failedSection(),
    textLayoutAccuracy: failedSection(),
    visualSimilarity: failedSection(),
    objectMappingAccuracy: failedSection(),
    compatibilityAccuracy: failedSection(),
    determinismAccuracy: failedSection(),
    metamorphicAccuracy: failedSection(),
    overallScore: 0,
    failureReasons: [{
      code: "EXTERNAL_VISUAL_OBSERVATION_FAILED",
      layer: "visual",
      message: cause instanceof Error ? cause.message : String(cause),
    }],
    visualDiff: null,
  };
}

function validateNativeCapture(capture, testCase, actualBytes, actual) {
  if (!isRecord(capture)) throw new Error(`${testCase.id} candidate capture must be a JSON object`);
  const expected = {
    fixtureSha256: testCase.fixtureSha256,
    format: testCase.format,
    unitIndex: testCase.unitIndex,
    unitType: testCase.unitType,
    unitCount: testCase.unitCount,
    pngSha256: sha256(actualBytes),
  };
  for (const [field, value] of Object.entries(expected)) {
    if (capture[field] !== value) {
      throw new Error(`${testCase.id} candidate capture ${field} does not match the manifest or PNG`);
    }
  }
  if (testCase.unitType === "sheet") {
    if (capture.sheetRange !== testCase.sheetRange) {
      throw new Error(`${testCase.id} candidate capture sheetRange does not match the manifest`);
    }
    const viewport = capture.viewport;
    if (!isRecord(viewport)
      || ![viewport.x, viewport.y, viewport.width, viewport.height].every(Number.isFinite)
      || viewport.width <= 0 || viewport.height <= 0) {
      throw new Error(`${testCase.id} candidate capture must record the resolved positive sheet viewport`);
    }
  } else if (capture.sheetRange !== undefined || capture.viewport !== undefined) {
    throw new Error(`${testCase.id} non-sheet capture cannot declare sheetRange or viewport`);
  }
  if (!Array.isArray(capture.diagnostics)) throw new Error(`${testCase.id} candidate capture must record diagnostics`);
  if (!Number.isInteger(capture.width) || capture.width < 1 || !Number.isInteger(capture.height) || capture.height < 1) {
    throw new Error(`${testCase.id} candidate capture dimensions are invalid`);
  }
  if (capture.width !== actual.width || capture.height !== actual.height) {
    throw new Error(`${testCase.id} candidate capture dimensions do not match the PNG`);
  }
}

function nativeVisualRegions(unitIndex, width, height) {
  const columns = Math.min(16, width);
  const rows = Math.min(16, height);
  const regions = [];
  for (let row = 0; row < rows; row += 1) {
    const top = Math.floor((row * height) / rows);
    const bottom = Math.floor(((row + 1) * height) / rows);
    for (let column = 0; column < columns; column += 1) {
      const left = Math.floor((column * width) / columns);
      const right = Math.floor(((column + 1) * width) / columns);
      const sourceKey = `visual-region-${row}-${column}`;
      regions.push({
        id: sourceKey,
        sourceKey,
        type: "visual-region",
        unitIndex,
        bounds: { x: left, y: top, width: right - left, height: bottom - top },
      });
    }
  }
  return regions;
}

async function writeJson(path, value) {
  await writeFile(path, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 });
}

const { suitePath, fingerprintPath, actualRoot, output } = parseArguments(process.argv.slice(2));
await mkdir(output, { recursive: true });

let suite;
let actualFingerprint;
let cases = [];
let inputsValid = false;
try {
  suite = JSON.parse(await readFile(suitePath, "utf8"));
  actualFingerprint = JSON.parse(await readFile(fingerprintPath, "utf8"));
  if (suite?.schemaVersion === 2) {
    validateNativeVisualSuite(suite);
    validateCandidateFingerprint(actualFingerprint, "currentCandidateFingerprint", true);
    cases = flattenNativeVisualCases(suite);
  } else {
    validateLegacySuite(suite);
    validateFingerprint(actualFingerprint, "currentEnvironmentFingerprint");
    cases = suite.cases;
  }
  inputsValid = true;
} catch (cause) {
  const summary = {
    schemaVersion: suite?.schemaVersion === 2 ? 2 : 1,
    generatedAt: new Date().toISOString(),
    suite: suitePath,
    status: "invalid-suite",
    total: 0,
    passed: 0,
    failed: 0,
    reports: [],
    failureReasons: [{ code: "EXTERNAL_VISUAL_SUITE_INVALID", message: cause instanceof Error ? cause.message : String(cause) }],
  };
  await writeJson(resolve(output, "summary.json"), summary);
  console.error(summary.failureReasons[0].message);
  process.exitCode = 1;
}

if (inputsValid) {
  const native = suite.schemaVersion === 2;
  const expectedFingerprint = native ? suite.candidateFingerprint : suite.environmentFingerprint;
  const changedFields = native
    ? changedCandidateFingerprintFields(expectedFingerprint, actualFingerprint)
    : changedFingerprintFields(expectedFingerprint, actualFingerprint);
  const environmentFingerprint = {
    expected: expectedFingerprint,
    actual: actualFingerprint,
    matches: changedFields.length === 0,
    changedFields,
  };
  if (changedFields.length !== 0) {
    const summary = {
      schemaVersion: native ? 2 : 1,
      generatedAt: new Date().toISOString(),
      suite: suitePath,
      oracleMode: "read-only",
      status: "environment-mismatch",
      environmentFingerprint,
      total: cases.length,
      passed: 0,
      failed: cases.length,
      reports: [],
      failureReasons: [{
        code: "ENVIRONMENT_FINGERPRINT_MISMATCH",
        message: `Current capture environment differs in: ${changedFields.join(", ")}`,
      }],
    };
    await writeJson(resolve(output, "summary.json"), summary);
    console.error(summary.failureReasons[0].message);
    process.exitCode = 1;
  } else {
    const suiteDirectory = dirname(suitePath);
    const candidateDirectory = actualRoot ?? suiteDirectory;
    const reports = [];
    const nativeReferenceCache = new Map();
    for (const testCase of cases) {
      let report;
      try {
        let nativeReference;
        const goldenPath = native
          ? resolveSafeRelativePath(suiteDirectory, testCase.goldenPng, `${testCase.id}.goldenPng`)
          : resolve(suiteDirectory, testCase.goldenPng);
        const actualPath = native
          ? resolveSafeRelativePath(candidateDirectory, testCase.actualPng, `${testCase.id}.actualPng`)
          : resolve(candidateDirectory, testCase.actualPng);
        const goldenBytes = await readFile(goldenPath);
        const actualBytes = await readFile(actualPath);
        if (native) {
          const fixturePath = resolveSafeRelativePath(suiteDirectory, testCase.fixture, `${testCase.id}.fixture`);
          const fixtureBytes = await readFile(fixturePath);
          if (sha256(fixtureBytes) !== testCase.fixtureSha256) {
            throw new Error(`${testCase.id} fixture SHA-256 does not match the reviewed manifest`);
          }
          if (sha256(goldenBytes) !== testCase.goldenPngSha256) {
            throw new Error(`${testCase.id} golden PNG SHA-256 does not match the reviewed manifest`);
          }
          nativeReference = await validateNativeReference(
            suiteDirectory,
            suite,
            testCase,
            goldenBytes,
            nativeReferenceCache,
          );
        }
        const golden = decodePng(goldenBytes, `${testCase.id} reference golden`);
        const actual = decodePng(actualBytes, `${testCase.id} OfficeViewer capture`);
        const expectedObservation = native || testCase.expectedObservationJson === undefined
          ? {}
          : JSON.parse(await readFile(resolve(suiteDirectory, testCase.expectedObservationJson), "utf8"));
        const actualObservation = native || testCase.actualObservationJson === undefined
          ? {}
          : JSON.parse(await readFile(resolve(suiteDirectory, testCase.actualObservationJson), "utf8"));
        if (!isRecord(expectedObservation) || !isRecord(actualObservation)) {
          throw new Error(`${testCase.id} observation JSON must contain snapshot objects`);
        }
        if (native) {
          const capturePath = resolveSafeRelativePath(
            candidateDirectory,
            testCase.actualObservationJson,
            `${testCase.id}.actualObservationJson`,
          );
          const capture = JSON.parse(await readFile(capturePath, "utf8"));
          validateNativeCapture(capture, testCase, actualBytes, actual);
        }
        const expectedObjects = native
          ? nativeVisualRegions(testCase.unitIndex, golden.width, golden.height)
          : Array.isArray(expectedObservation.objects) ? expectedObservation.objects : [];
        const actualObjects = native
          ? nativeVisualRegions(testCase.unitIndex, actual.width, actual.height)
          : Array.isArray(actualObservation.objects) ? actualObservation.objects : [];
        const expectedSnapshot = {
          ...expectedObservation,
          units: [{ index: testCase.unitIndex, type: testCase.unitType, width: golden.width, height: golden.height }],
          objects: expectedObjects,
          visuals: [{ unitIndex: testCase.unitIndex, width: golden.width, height: golden.height, data: golden.data }],
        };
        const actualSnapshot = {
          ...actualObservation,
          units: [{ index: testCase.unitIndex, type: testCase.unitType, width: actual.width, height: actual.height }],
          objects: actualObjects,
          visuals: [{ unitIndex: testCase.unitIndex, width: actual.width, height: actual.height, data: actual.data }],
        };
        const accuracyReport = evaluateAccuracy({
          fileId: testCase.id,
          corpusClass: testCase.corpusClass,
          expected: expectedSnapshot,
          actual: actualSnapshot,
          declaredCoverage: native ? VISUAL_COVERAGE : testCase.declaredCoverage ?? VISUAL_COVERAGE,
          policy: native ? NATIVE_VISUAL_POLICY : testCase.policy,
        });
        report = {
          schemaVersion: native ? 2 : 1,
          oracle: {
            suite: native ? suite.oracleSuite : undefined,
            renderer: native ? testCase.oracleApplication : suite.environmentFingerprint.referenceRenderer,
            mode: "read-only",
            fixture: testCase.fixture,
            fixtureSha256: native ? testCase.fixtureSha256 : undefined,
            format: native ? testCase.format : undefined,
            capture: native ? testCase.capture : undefined,
            reference: native ? nativeReference : undefined,
            goldenPng: testCase.goldenPng,
            pngSha256: digest(goldenBytes),
            rgbaSha256: digest(golden.data),
            width: golden.width,
            height: golden.height,
          },
          candidate: {
            renderer: native ? suite.candidateFingerprint.browser : suite.environmentFingerprint.candidateRenderer,
            actualPng: testCase.actualPng,
            pngSha256: digest(actualBytes),
            rgbaSha256: digest(actual.data),
            width: actual.width,
            height: actual.height,
          },
          oracleFingerprint: native ? suite.oracleFingerprint : undefined,
          environmentFingerprint: expectedFingerprint,
          thresholdPolicy: native ? suite.thresholdPolicy : undefined,
          unitCount: native ? testCase.unitCount : undefined,
          sheetRange: native ? testCase.sheetRange : undefined,
          ...accuracyReport,
          visualDiff: visualDiff(
            golden,
            actual,
            native ? NATIVE_VISUAL_POLICY.pixelChannel : testCase.policy?.pixelChannel ?? 8,
            native ? NATIVE_VISUAL_POLICY.pixelRadius : testCase.policy?.pixelRadius ?? 0,
          ),
        };
      } catch (cause) {
        report = observationFailure(testCase, suite, expectedFingerprint, cause);
      }
      reports.push(report);
      const reportName = `${safeName(testCase.id)}.json`;
      await writeJson(resolve(output, reportName), report);
      console.log(`${report.passed ? "PASS" : "FAIL"} ${testCase.id} score=${report.overallScore} report=${resolve(output, reportName)}`);
    }
    const passed = reports.filter((report) => report.passed).length;
    const formatSummary = native
      ? Object.fromEntries([...new Set(cases.map(({ format }) => format))].sort().map((format) => {
        const formatReports = reports.filter((report) => cases.find(({ id }) => id === report.fileId)?.format === format);
        const formatPassed = formatReports.filter(({ passed: reportPassed }) => reportPassed).length;
        return [format, { total: formatReports.length, passed: formatPassed, failed: formatReports.length - formatPassed }];
      }))
      : undefined;
    const summary = {
      schemaVersion: native ? 2 : 1,
      generatedAt: new Date().toISOString(),
      suite: suitePath,
      oracleMode: "read-only",
      status: passed === reports.length ? "passed" : "failed",
      oracleSuite: native ? suite.oracleSuite : undefined,
      thresholdPolicy: native ? suite.thresholdPolicy : undefined,
      oracleFingerprint: native ? suite.oracleFingerprint : undefined,
      environmentFingerprint,
      total: reports.length,
      passed,
      failed: reports.length - passed,
      averageScore: reports.length === 0 ? 0 : rounded(reports.reduce((sum, report) => sum + report.overallScore, 0) / reports.length),
      formats: formatSummary,
      reports: reports.map((report) => ({
        id: report.fileId,
        passed: report.passed,
        overallScore: report.overallScore,
        exactPixelSimilarity: report.visualDiff?.exactPixelSimilarity ?? null,
        tolerantPixelSimilarity: report.visualDiff?.tolerantPixelSimilarity ?? null,
        report: `${safeName(report.fileId)}.json`,
      })),
      failureReasons: [],
    };
    await writeJson(resolve(output, "summary.json"), summary);
    console.log(`Summary: ${summary.passed}/${summary.total} passed, average=${summary.averageScore}, report=${resolve(output, "summary.json")}`);
    if (summary.failed !== 0) process.exitCode = 1;
  }
}
