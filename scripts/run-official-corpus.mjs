import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { basename, dirname, extname, relative, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { createOfficeEngine } from "../dist/engine.js";
import { OfficeEngineError } from "../dist/types.js";
import { extendedFormatPack } from "../dist/extended-formats.js";
import { readZipEntries } from "./accuracy-metamorphic.mjs";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const cacheRoot = resolve(root, ".cache/official-corpus");
const reportPath = resolve(root, "output/official-corpus-report.json");
const mode = process.argv.includes("--full") ? "full" : "smoke";
const quiet = process.argv.includes("--quiet");
const unknownArguments = process.argv.slice(2).filter((argument) => argument !== "--full" && argument !== "--quiet");

if (unknownArguments.length !== 0) {
  throw new Error(`Unknown argument(s): ${unknownArguments.join(", ")}`);
}

const formats = Object.freeze({
  ".docx": { format: "docx", kind: "text" },
  ".xlsx": { format: "xlsx", kind: "spreadsheet" },
  ".pptx": { format: "pptx", kind: "presentation" },
  ".odt": { format: "odt", kind: "text" },
  ".ods": { format: "ods", kind: "spreadsheet" },
  ".odp": { format: "odp", kind: "presentation" },
});

const libreOfficeCommit = "db5d24b4e0d0350125babbaf377d2ed7f043e161";
const libreOfficeVisualReferences = Object.freeze([
  { fixture: "sd/qa/unit/data/TextFittingComparisonWithMSO_1.pptx", embeddedPngs: 14 },
  { fixture: "sd/qa/unit/data/TextFittingComparisonWithMSO_2.pptx", embeddedPngs: 6 },
  { fixture: "sd/qa/unit/data/TextFittingComparisonWithMSO_3.pptx", embeddedPngs: 6 },
  { fixture: "sd/qa/unit/data/TextFittingComparisonWithMSO_TopBottomMiddleAlignment.pptx", embeddedPngs: 6 },
]);

const suites = Object.freeze([
  {
    id: "microsoft-open-xml-sdk-v3.5.1",
    publisher: "Microsoft / .NET Foundation",
    classification: "official implementation regression assets",
    repository: "https://github.com/dotnet/Open-XML-SDK.git",
    commit: "3139fdfd27414548a41555f7848d5728f6e71a42",
    roots: ["test/DocumentFormat.OpenXml.Tests.Assets/assets"],
  },
  {
    id: "oasis-odf-tc-1.3-feature-tests",
    publisher: "OASIS OpenDocument Format Technical Committee",
    classification: "standards committee feature-test documents",
    repository: "https://github.com/oasis-tcs/odf-tc.git",
    commit: "16a59d945875bd74834e82e166bebded316478da",
    roots: ["src/test/resources/odf1.3/testfiles"],
  },
  {
    id: "tdf-odf-toolkit-v0.13.0",
    publisher: "The Document Foundation ODF Toolkit project",
    classification: "official implementation regression assets",
    repository: "https://github.com/tdf/odftoolkit.git",
    commit: "b926a6134a2fee782076500dfc02c47c2d651cff",
    roots: [
      "odfdom/src/test/resources/test-input",
      "validator/src/test/resources",
    ],
  },
  {
    id: "libreoffice-core-qa-db5d24b4",
    publisher: "The Document Foundation / LibreOffice",
    classification: "official implementation regression assets with selected MSO embedded render references",
    repository: "https://github.com/LibreOffice/core.git",
    commit: libreOfficeCommit,
    roots: [
      "sw/qa/extras/ooxmlimport/data",
      "sw/qa/extras/ooxmlexport/data",
      "sw/qa/extras/odfimport/data",
      "sw/qa/extras/odfexport/data",
      "sc/qa/unit/data/xlsx",
      "sc/qa/unit/data/ods",
      "sd/qa/unit/data",
      "chart2/qa/extras/chart2dump/data",
      "chart2/qa/extras/data",
      "chart2/qa/extras/xshape/data",
      "chart2/qa/unit/data",
    ],
    excludeSecurityFixtures: true,
    contentFormatOverrides: {
      "sw/qa/extras/ooxmlexport/data/tdf171025_pageAfter.docx": "odt",
      "sw/qa/extras/ooxmlexport/data/tdf171038_pageAfter.docx": "odt",
      "sw/qa/extras/odfimport/data/tdf76322_columnBreakInHeader.docx": "odt",
    },
    visualReferences: libreOfficeVisualReferences,
  },
]);

const smokeCases = Object.freeze([
  smokeCase(
    "microsoft-open-xml-sdk-v3.5.1",
    "strict-word-comment-range.docx",
    "https://raw.githubusercontent.com/dotnet/Open-XML-SDK/3139fdfd27414548a41555f7848d5728f6e71a42/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Word/CommentRangeStart.docx",
    "a6aa3203a42e2095294e107687bb2c392b3a8cb6bc6cc1b20946960a8e8026d3",
  ),
  smokeCase(
    "microsoft-open-xml-sdk-v3.5.1",
    "strict-spreadsheet-filter.xlsx",
    "https://raw.githubusercontent.com/dotnet/Open-XML-SDK/3139fdfd27414548a41555f7848d5728f6e71a42/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Excel/filter_type.xlsx",
    "9e32e6f06bcec398a6b1d89d5564b8b2ff7c3ef93c61ec0ab3da6497e46bf4b1",
  ),
  smokeCase(
    "microsoft-open-xml-sdk-v3.5.1",
    "strict-presentation-size.pptx",
    "https://raw.githubusercontent.com/dotnet/Open-XML-SDK/3139fdfd27414548a41555f7848d5728f6e71a42/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/PowerPoint/sldsz-Slide%20Size-type-screen16x9.pptx",
    "e5d50eb72deb0d1282227b2300aeec3072d1957779b93f9b931583262e7c2614",
  ),
  smokeCase(
    "oasis-odf-tc-1.3-feature-tests",
    "odf13-metadata.odt",
    "https://raw.githubusercontent.com/oasis-tcs/odf-tc/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.3/testfiles/3776%20meta_creator-initials.odt",
    "85c4ca375c022152d06d05a43fdfdfc8547fb5239628ff5249d04538ea4d39e1",
  ),
  smokeCase(
    "oasis-odf-tc-1.3-feature-tests",
    "odf13-sheet-tab-color.ods",
    "https://raw.githubusercontent.com/oasis-tcs/odf-tc/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.3/testfiles/2173%20table_tab-color.ods",
    "1eb597df7d3bd497aa432b311df26cec6aac10dbd093c92ead181b19b23de3aa",
  ),
  smokeCase(
    "oasis-odf-tc-1.3-feature-tests",
    "odf13-presentation-linecap.odp",
    "https://raw.githubusercontent.com/oasis-tcs/odf-tc/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.3/testfiles/3742%20linecap%20combine%20draw%20with%20svg.odp",
    "4fb38a082820a196046ad41116c4f0d1916cebd2902eb7fbdbb925686d09ec31",
  ),
  smokeCase(
    "tdf-odf-toolkit-v0.13.0",
    "toolkit-hello-world.odt",
    "https://raw.githubusercontent.com/tdf/odftoolkit/b926a6134a2fee782076500dfc02c47c2d651cff/odfdom/src/test/resources/test-input/HelloWorld.odt",
    "8bcab0ccafe0ce80fb85b4b90c4aabef92b101576457e3752a89ecc0f2e73c1b",
  ),
  smokeCase(
    "tdf-odf-toolkit-v0.13.0",
    "toolkit-basic-sheet.ods",
    "https://raw.githubusercontent.com/tdf/odftoolkit/b926a6134a2fee782076500dfc02c47c2d651cff/odfdom/src/test/resources/test-input/Basic.ods",
    "0a85fa4cff2979c997970c5abfb7ea98397d1df6b1efe62ce63d5b6fdc898f2a",
    { status: "rejected", code: "OBJECT_LIMIT" },
  ),
  smokeCase(
    "tdf-odf-toolkit-v0.13.0",
    "toolkit-slide.odp",
    "https://raw.githubusercontent.com/tdf/odftoolkit/b926a6134a2fee782076500dfc02c47c2d651cff/odfdom/src/test/resources/test-input/SlideTest1.odp",
    "75c8f889bdb55b1d5b6b3501ccec7d376d5509c033e81285612ec4039afd5d01",
  ),
  smokeCase(
    "libreoffice-core-qa-db5d24b4",
    "libreoffice-page-content-bottom.docx",
    `https://raw.githubusercontent.com/LibreOffice/core/${libreOfficeCommit}/sw/qa/extras/ooxmlexport/data/page-content-bottom.docx`,
    "9c5563f533a3b32483bf0c86a409ae4e5da91b35437c058d12a449008cd774d7",
  ),
  smokeCase(
    "libreoffice-core-qa-db5d24b4",
    "libreoffice-space.odt",
    `https://raw.githubusercontent.com/LibreOffice/core/${libreOfficeCommit}/sw/qa/extras/odfimport/data/space.odt`,
    "389e4e6b8d768bdf1cac893d35727652262141a4458e2b72225da2d3ca858c56",
  ),
  smokeCase(
    "libreoffice-core-qa-db5d24b4",
    "libreoffice-empty-noconf.xlsx",
    `https://raw.githubusercontent.com/LibreOffice/core/${libreOfficeCommit}/sc/qa/unit/data/xlsx/empty-noconf.xlsx`,
    "556489821a2d6c42aaff824c8165c652a402ffa73ca45c65758302a893d5d785",
  ),
  smokeCase(
    "libreoffice-core-qa-db5d24b4",
    "libreoffice-tdf76310.ods",
    `https://raw.githubusercontent.com/LibreOffice/core/${libreOfficeCommit}/sc/qa/unit/data/ods/tdf76310.ods`,
    "c83255e2dcff4ce7af02d3cbdd18148aeae7f16b10ddf3911c74791075e2c7a4",
  ),
  smokeCase(
    "libreoffice-core-qa-db5d24b4",
    "libreoffice-mso-text-fitting.pptx",
    `https://raw.githubusercontent.com/LibreOffice/core/${libreOfficeCommit}/sd/qa/unit/data/TextFittingComparisonWithMSO_2.pptx`,
    "6f7edbcdfe083546e1d145a5929c419c74d4fabfcd9622bb193c55b784155a6a",
  ),
  smokeCase(
    "libreoffice-core-qa-db5d24b4",
    "libreoffice-16-9.odp",
    `https://raw.githubusercontent.com/LibreOffice/core/${libreOfficeCommit}/sd/qa/unit/data/odp/16-9.odp`,
    "a6851686ee63dee52b32fef9ff8819e48bca65034334a8bc0caa2e2d4352690f",
  ),
]);

const fullBaseline = Object.freeze({
  total: 4_736,
  minimumOpened: 990,
  forbiddenRejectedCodePrefixes: Object.freeze(["CORE_"]),
  cases: Object.freeze({
    "microsoft-open-xml-sdk-v3.5.1/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/asSources/wordprocessing/complex1_NOR.docx": Object.freeze({ units: 11 }),
  }),
  suites: Object.freeze({
    "microsoft-open-xml-sdk-v3.5.1": Object.freeze({ total: 707, minimumOpened: 651 }),
    "oasis-odf-tc-1.3-feature-tests": Object.freeze({ total: 34, minimumOpened: 34 }),
    "tdf-odf-toolkit-v0.13.0": Object.freeze({ total: 350, minimumOpened: 305 }),
    "libreoffice-core-qa-db5d24b4": Object.freeze({ total: 3_645, minimumOpened: 3_336 }),
  }),
});

function smokeCase(suite, name, url, sha256, expectation = { status: "opened" }) {
  const expected = formats[extname(name).toLowerCase()];
  if (expected === undefined) throw new Error(`Unsupported smoke fixture extension: ${name}`);
  return { suite, name, url, sha256, expectation, ...expected };
}

function digest(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function cachedDownload(testCase) {
  const target = resolve(cacheRoot, "smoke", testCase.name);
  if (existsSync(target) && digest(readFileSync(target)) === testCase.sha256) return target;

  const response = await fetch(testCase.url, { signal: AbortSignal.timeout(30_000) });
  if (!response.ok) throw new Error(`Download failed (${response.status}): ${testCase.url}`);
  const contentLength = Number(response.headers.get("content-length") ?? 0);
  if (contentLength > 16 * 1024 * 1024) throw new Error(`Fixture exceeds 16 MiB: ${testCase.name}`);
  const bytes = Buffer.from(await response.arrayBuffer());
  if (bytes.byteLength > 16 * 1024 * 1024) throw new Error(`Fixture exceeds 16 MiB: ${testCase.name}`);
  const actual = digest(bytes);
  if (actual !== testCase.sha256) {
    throw new Error(`SHA-256 mismatch for ${testCase.name}: expected ${testCase.sha256}, received ${actual}`);
  }
  mkdirSync(dirname(target), { recursive: true });
  const temporary = `${target}.${process.pid}.tmp`;
  writeFileSync(temporary, bytes, { mode: 0o600 });
  renameSync(temporary, target);
  return target;
}

function runGit(arguments_, cwd) {
  const result = spawnSync("git", arguments_, { cwd, encoding: "utf8" });
  if (result.status !== 0) {
    throw new Error(`git ${arguments_.join(" ")} failed: ${(result.stderr || result.stdout).trim()}`);
  }
  return result.stdout.trim();
}

function checkoutSuite(suite) {
  const checkout = resolve(cacheRoot, "full", suite.id);
  if (existsSync(resolve(checkout, ".git"))) {
    const head = spawnSync("git", ["rev-parse", "HEAD"], { cwd: checkout, encoding: "utf8" });
    if (head.status === 0) {
      const clean = runGit(["status", "--porcelain", "--untracked-files=all"], checkout) === "";
      if (head.stdout.trim() === suite.commit && clean) {
        runGit(["sparse-checkout", "set", ...suite.roots], checkout);
        if (runGit(["status", "--porcelain", "--untracked-files=all"], checkout) === "") return checkout;
      }
    }
  }

  rmSync(checkout, { recursive: true, force: true });
  mkdirSync(checkout, { recursive: true });
  runGit(["init", "--quiet"], checkout);
  runGit(["remote", "add", "origin", suite.repository], checkout);
  runGit(["sparse-checkout", "init", "--cone"], checkout);
  runGit(["sparse-checkout", "set", ...suite.roots], checkout);
  runGit(["fetch", "--quiet", "--depth", "1", "origin", suite.commit], checkout);
  runGit(["checkout", "--quiet", "--detach", "FETCH_HEAD"], checkout);
  const current = runGit(["rev-parse", "HEAD"], checkout);
  if (current !== suite.commit) throw new Error(`Pinned commit mismatch for ${suite.id}`);
  return checkout;
}

function filesBelow(directory, output = []) {
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const path = resolve(directory, entry.name);
    if (entry.isDirectory()) filesBelow(path, output);
    else if (entry.isFile() && formats[extname(entry.name).toLowerCase()] !== undefined) output.push(path);
  }
  return output;
}

function isSecurityFixture(path) {
  return /(?:^|\/)fail(?:\/|$)/u.test(path)
    || /^(?:BID|CVE|EDB|RC4|ofz)/iu.test(basename(path));
}

function xmlElement(xml, name) {
  return xml.match(new RegExp(`<${name}>([^<]*)</${name}>`, "u"))?.[1];
}

function inspectVisualReferences(suite, checkout) {
  if (suite.visualReferences === undefined) return [];
  const decoder = new TextDecoder();
  return suite.visualReferences.map((reference) => {
    const fixture = resolve(checkout, reference.fixture);
    const fixtureBytes = readFileSync(fixture);
    const entries = readZipEntries(fixtureBytes);
    const byName = new Map(entries.map((entry) => [entry.name, entry]));
    const applicationProperties = byName.get("docProps/app.xml");
    if (applicationProperties === undefined) throw new Error(`${reference.fixture} has no docProps/app.xml`);
    const appXml = decoder.decode(applicationProperties.data);
    const application = xmlElement(appXml, "Application");
    const appVersion = xmlElement(appXml, "AppVersion");
    if (application !== "Microsoft Office PowerPoint" || appVersion !== "16.0000") {
      throw new Error(`${reference.fixture} has unexpected producer ${application ?? "unknown"} ${appVersion ?? "unknown"}`);
    }

    const images = entries
      .filter(({ name }) => /^ppt\/media\/image\d+\.png$/u.test(name))
      .sort(({ name: left }, { name: right }) => left.localeCompare(right, "en", { numeric: true }));
    if (images.length !== reference.embeddedPngs) {
      throw new Error(`${reference.fixture} has ${images.length} MSO PNG references; expected ${reference.embeddedPngs}`);
    }

    const relationships = entries.filter(({ name }) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/u.test(name));
    const mappings = images.map((entry) => {
      const file = basename(entry.name);
      const slides = relationships
        .filter((relationship) => decoder.decode(relationship.data).includes(`Target="../media/${file}"`))
        .map(({ name }) => Number(name.match(/slide(\d+)\.xml\.rels$/u)[1]))
        .sort((left, right) => left - right);
      if (slides.length === 0) throw new Error(`${reference.fixture} does not map ${file} to a slide`);
      return { file, slides, bytes: entry.data.length, sha256: digest(entry.data) };
    });

    return {
      suite: suite.id,
      fixture: reference.fixture,
      fixtureSha256: digest(fixtureBytes),
      provenance: "mso-embedded-render-reference",
      application,
      appVersion,
      mappings,
    };
  });
}

const visualReferenceResults = [];

async function prepareCases() {
  if (mode === "smoke") {
    return Promise.all(smokeCases.map(async (testCase) => ({
      ...testCase,
      path: await cachedDownload(testCase),
    })));
  }

  const cases = [];
  for (const suite of suites) {
    const checkout = checkoutSuite(suite);
    visualReferenceResults.push(...inspectVisualReferences(suite, checkout));
    for (const corpusRoot of suite.roots) {
      const absoluteRoot = resolve(checkout, corpusRoot);
      for (const path of filesBelow(absoluteRoot).sort()) {
        const name = relative(checkout, path);
        if (suite.excludeSecurityFixtures === true && isSecurityFixture(name)) continue;
        const override = suite.contentFormatOverrides?.[name];
        const expected = formats[override === undefined ? extname(path).toLowerCase() : `.${override}`];
        cases.push({
          suite: suite.id,
          name,
          path,
          ...expected,
        });
      }
    }
  }
  return cases;
}

function validateObject(object, document) {
  const values = Object.values(object.bounds);
  if (values.some((value) => !Number.isFinite(value)) || object.bounds.width < 0 || object.bounds.height < 0) {
    throw new Error(`${object.id} has invalid bounds`);
  }
  if (object.source.format !== document.info.format) {
    throw new Error(`${object.id} has a mismatched source format`);
  }
}

async function runCase(engine, testCase) {
  const started = performance.now();
  let document;
  let sha256;
  try {
    const bytes = readFileSync(testCase.path);
    sha256 = digest(bytes);
    document = await engine.open(bytes);
    if (document.info.format !== testCase.format || document.info.kind !== testCase.kind) {
      throw new Error(
        `content identified as ${document.info.format}/${document.info.kind}; expected ${testCase.format}/${testCase.kind}`,
      );
    }
    if (document.info.units.length === 0) throw new Error("document has no renderable units");

    const object = await document.getObject("object:0");
    let objectProbe = "empty";
    if (object !== undefined) {
      validateObject(object, document);
      objectProbe = object.source.mapping;
      if (object.bounds.width > 0 && object.bounds.height > 0) {
        const hits = await document.hitTest({
          unitIndex: object.unitIndex,
          x: object.bounds.x + object.bounds.width / 2,
          y: object.bounds.y + object.bounds.height / 2,
          limit: 256,
        });
        if (!hits.some((hit) => hit.object.id === object.id)) {
          throw new Error(`${object.id} was not returned by a center-point hit test`);
        }
      }
    }

    const diagnostics = document.diagnostics();
    return {
      suite: testCase.suite,
      name: testCase.name,
      expectedFormat: testCase.format,
      sha256,
      status: "opened",
      format: document.info.format,
      kind: document.info.kind,
      units: document.info.units.length,
      diagnostics: diagnostics.map(({ code, severity, fidelity, phase }) => ({ code, severity, fidelity, phase })),
      objectProbe,
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } catch (cause) {
    const structured = cause instanceof OfficeEngineError;
    return {
      suite: testCase.suite,
      name: testCase.name,
      expectedFormat: testCase.format,
      ...(sha256 === undefined ? {} : { sha256 }),
      status: structured ? "rejected" : "failed",
      code: structured ? cause.code : "UNEXPECTED_ERROR",
      message: cause instanceof Error ? cause.message : String(cause),
      diagnostics: structured
        ? cause.diagnostics.map(({ code, severity, fidelity, phase }) => ({ code, severity, fidelity, phase }))
        : [],
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } finally {
    document?.close();
  }
}

const cases = await prepareCases();
const wasm = readFileSync(resolve(root, "dist/office-viewer-core.wasm"));
const engine = await createOfficeEngine({
  execution: "inline",
  wasm,
  formatPack: async () => ({
    async load(candidate) {
      return readFileSync(await extendedFormatPack.load(candidate));
    },
  }),
});
const results = [];
try {
  for (const testCase of cases) {
    const result = await runCase(engine, testCase);
    const outcomeAccepted = mode === "full"
      ? result.status !== "failed"
      : result.status === testCase.expectation.status
        && (testCase.expectation.code === undefined || result.code === testCase.expectation.code);
    results.push({ ...result, outcomeAccepted });
    const detail = result.status === "opened"
      ? `${result.format} units=${result.units} diagnostics=${result.diagnostics.length} object=${result.objectProbe}`
      : `${result.code} ${result.message}`;
    const label = outcomeAccepted ? (mode === "full" ? "DONE" : "PASS") : "FAIL";
    if (!quiet || !outcomeAccepted || result.status === "failed") {
      console.log(`${label.padEnd(5)} ${testCase.suite} / ${basename(testCase.name)} — ${result.status}: ${detail}`);
    }
  }
} finally {
  engine.close();
}

const bySuite = Object.fromEntries(suites.map((suite) => {
  const suiteResults = results.filter((result) => result.suite === suite.id);
  return [suite.id, {
    total: suiteResults.length,
    opened: suiteResults.filter(({ status }) => status === "opened").length,
    rejected: suiteResults.filter(({ status }) => status === "rejected").length,
    failed: suiteResults.filter(({ status }) => status === "failed").length,
  }];
}));
const summary = {
  total: results.length,
  opened: results.filter(({ status }) => status === "opened").length,
  rejected: results.filter(({ status }) => status === "rejected").length,
  failed: results.filter(({ status }) => status === "failed").length,
  acceptedOutcomes: results.filter(({ outcomeAccepted }) => outcomeAccepted).length,
  unexpected: results.filter(({ outcomeAccepted }) => !outcomeAccepted).length,
  documentDiagnostics: results.reduce((count, result) => count + result.diagnostics.length, 0),
  uniqueInputs: new Set(results.map(({ sha256 }) => sha256).filter(Boolean)).size,
  bySuite,
};
const baselineFailures = [];
if (mode === "full") {
  if (summary.total !== fullBaseline.total) {
    baselineFailures.push(`corpus size ${summary.total} != ${fullBaseline.total}`);
  }
  if (summary.opened < fullBaseline.minimumOpened) {
    baselineFailures.push(`opened ${summary.opened} < ${fullBaseline.minimumOpened}`);
  }
  for (const result of results.filter(({ status }) => status === "rejected")) {
    if (fullBaseline.forbiddenRejectedCodePrefixes.some((prefix) => result.code.startsWith(prefix))) {
      baselineFailures.push(`${result.name} was rejected with forbidden internal code ${result.code}`);
    }
  }
  for (const [suiteId, expectation] of Object.entries(fullBaseline.suites)) {
    const actual = bySuite[suiteId];
    if (actual.total !== expectation.total) {
      baselineFailures.push(`${suiteId} size ${actual.total} != ${expectation.total}`);
    }
    if (actual.opened < expectation.minimumOpened) {
      baselineFailures.push(`${suiteId} opened ${actual.opened} < ${expectation.minimumOpened}`);
    }
  }
  for (const [id, expectation] of Object.entries(fullBaseline.cases)) {
    const result = results.find(({ suite, name }) => `${suite}/${name}` === id);
    if (result?.status !== "opened" || result.units !== expectation.units) {
      baselineFailures.push(`${id} units ${result?.units ?? "unavailable"} != ${expectation.units}`);
    }
  }
}
const report = {
  generatedAt: new Date().toISOString(),
  mode,
  scope: "Content identification, bounded package parsing, unit metadata, diagnostics, source mapping and hit-test smoke; visual rendering is verified in the browser inspector.",
  certification: "These assets do not provide a complete six-format visual oracle and this report is not an ISO, ECMA or OASIS certification.",
  suites,
  visualReferences: visualReferenceResults,
  ...(mode === "full" ? {
    regressionBaseline: {
      ...fullBaseline,
      met: baselineFailures.length === 0,
      failures: baselineFailures,
    },
  } : {}),
  summary,
  cases: results,
};
mkdirSync(dirname(reportPath), { recursive: true });
writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`);

const completion = mode === "full"
  ? `${summary.acceptedOutcomes}/${summary.total} completed with a structured outcome`
  : `${summary.acceptedOutcomes}/${summary.total} expectations met`;
console.log(`\nSummary: ${completion}; ${summary.opened} opened, ${summary.rejected} structured rejections, ${summary.failed} failures`);
if (baselineFailures.length !== 0) console.error(`Regression baseline failed: ${baselineFailures.join("; ")}`);
console.log(`Report: ${reportPath}`);
if (summary.unexpected !== 0 || baselineFailures.length !== 0) process.exitCode = 1;
