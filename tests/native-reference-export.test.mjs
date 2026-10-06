import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { readFile, mkdtemp, stat, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { acquireNativeWorkspace, cacheNativeReference, restoreNativeReference } from "../scripts/export-native-reference.mjs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const source = await readFile(resolve(root, "scripts/export-native-reference.mjs"), "utf8");

test("native reference exporter exposes only the closed Office, iWork, and WPS golden applications", () => {
  const result = spawnSync(process.execPath, ["scripts/export-native-reference.mjs", "--help"], {
    cwd: root,
    encoding: "utf8",
  });
  assert.equal(result.status, 0, result.stderr);
  for (const application of ["Microsoft PowerPoint", "Microsoft Word", "Microsoft Excel", "Keynote", "Pages", "Numbers", "WPS Office"]) {
    assert.match(source, new RegExp(application, "u"));
  }
  assert.match(result.stdout, /PPTX\/PPT\/ODP\s+Microsoft PowerPoint/u);
  assert.match(result.stdout, /DOCX\/DOC\/ODT\/RTF\s+Microsoft Word/u);
  assert.match(result.stdout, /XLSX\/XLS\/ODS\s+Microsoft Excel/u);
  assert.match(result.stdout, /WPS\/DPS\s+WPS Office/u);
  assert.match(result.stdout, /ET\s+WPS Office/u);
  assert.match(result.stdout, /--sheet-viewport/u);
  assert.match(result.stdout, /--native-pdf/u);
  assert.match(result.stdout, /--range/u);
  assert.match(result.stdout, /--unit-index/u);
  assert.match(result.stdout, /--allow-print-pages/u);
  assert.match(result.stdout, /--reviewed-by <stable-id>/u);
  assert.doesNotMatch(source, /LibreOffice/iu);
  for (const [application, suite, bundleId] of [
    ["Microsoft PowerPoint", "microsoft-office", "com.microsoft.Powerpoint"],
    ["Microsoft Word", "microsoft-office", "com.microsoft.Word"],
    ["Microsoft Excel", "microsoft-office", "com.microsoft.Excel"],
    ["Keynote", "apple-iwork", "com.apple.iWork.Keynote"],
    ["Pages", "apple-iwork", "com.apple.iWork.Pages"],
    ["Numbers", "apple-iwork", "com.apple.iWork.Numbers"],
    ["WPS Office", "wps-office", "com.kingsoft.wpsoffice.mac"],
  ]) {
    assert.match(source, new RegExp(`name: "${application}",\\n\\s+oracleSuite: "${suite}",\\n\\s+expectedBundleId: "${bundleId}"`, "u"));
  }
  assert.match(source, /if \(bundleId !== application\.expectedBundleId\)/u);
  for (const [extension, application] of [
    ["odp", "powerpoint"],
    ["pptm", "powerpoint"],
    ["potx", "powerpoint"],
    ["potm", "powerpoint"],
    ["ppsx", "powerpoint"],
    ["ppsm", "powerpoint"],
    ["docm", "word"],
    ["dotx", "word"],
    ["dotm", "word"],
    ["ott", "word"],
    ["xlsm", "excel"],
    ["xltx", "excel"],
    ["xltm", "excel"],
    ["odt", "word"],
    ["rtf", "word"],
    ["ods", "excel"],
    ["wps", "wps-writer"],
    ["et", "wps-spreadsheets"],
    ["dps", "wps-presentation"],
  ]) {
    assert.match(source, new RegExp(`\\["\\.${extension}", "${application}"\\]`, "u"));
  }
});

test("every native reference manifest records its closed oracle suite identity", () => {
  assert.equal(source.match(/schemaVersion: 1,\n\s+oracleSuite: application\.oracleSuite/gu)?.length, 3);
  assert.equal(source.match(/oracleSuite: "microsoft-office"/gu)?.length, 3);
  assert.equal(source.match(/oracleSuite: "apple-iwork"/gu)?.length, 3);
  assert.equal(source.match(/oracleSuite: "wps-office"/gu)?.length, 3);
  assert.doesNotMatch(source, /oracleSuite:\s*"libreoffice"/iu);
});

test("manual reference imports require a stable human reviewer and emit a bounded attestation", () => {
  const invalidInvocations = [
    {
      arguments: ["example.pptx", "output", "--native-pdf", "reviewed.pdf"],
      message: /--reviewed-by is required with --native-pdf or --sheet-viewport/u,
    },
    {
      arguments: [
        "example.ods", "output", "--sheet-viewport", "reviewed.png", "--range", "A1:H40", "--unit-index", "0",
      ],
      message: /--reviewed-by is required with --native-pdf or --sheet-viewport/u,
    },
    {
      arguments: ["example.pptx", "output", "--reviewed-by", "qa.bot"],
      message: /--reviewed-by is valid only with --native-pdf or --sheet-viewport/u,
    },
    {
      arguments: ["example.pptx", "output", "--native-pdf", "reviewed.pdf", "--reviewed-by", "qa bot"],
      message: /--reviewed-by must be a stable 1-128 character identifier/u,
    },
    {
      arguments: ["example.pptx", "output", "--native-pdf", "reviewed.pdf", "--reviewed-by"],
      message: /Missing --reviewed-by value/u,
    },
  ];
  for (const invocation of invalidInvocations) {
    const result = spawnSync(process.execPath, ["scripts/export-native-reference.mjs", ...invocation.arguments], {
      cwd: root,
      encoding: "utf8",
    });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, invocation.message);
  }

  assert.equal(source.match(/reviewAttestation: \{/gu)?.length, 2);
  assert.equal(source.match(/reviewedAt: generatedAt/gu)?.length, 2);
  assert.equal(source.match(/scope: "all-rendered-units-visual-fidelity-v1"/gu)?.length, 2);
  assert.equal(source.match(/provenanceControl: "human-controlled"/gu)?.length, 2);
  assert.equal(source.match(/provenanceVerification: "not performed by this script"/gu)?.length, 2);
  assert.match(source, /artifactSha256: copiedPdfHash/u);
  assert.match(source, /artifactSha256: viewportHash/u);
  assert.match(source, /capture: "pdf-export"/u);
  assert.match(source, /nativeApplicationAutomation: true/u);
  assert.match(source, /operational observation is not cryptographic producer verification/u);
});

test("native reference exporter isolates source files and distinguishes sheet viewports from print pages", () => {
  assert.match(source, /mkdtemp\(resolve\(outputParent, "\.native-reference-"\)\)/u);
  assert.match(source, /const copiedFixture = resolve\(nativeDirectory, `source\$\{extension\}`\)/u);
  assert.match(source, /copyFile\(fixture, copiedFixture\)/u);
  assert.match(source, /await rm\(copiedFixture\)/u);
  assert.match(source, /Original fixture changed during native reference export/u);
  assert.match(source, /capture:\s*"sheet-viewport"/u);
  assert.match(source, /capture:\s*"reviewed-native-pdf-import"/u);
  assert.match(source, /provenanceVerification:\s*"not performed by this script"/u);
  assert.match(source, /assertPdfSignature/u);
  assert.match(source, /Imported native PDF does not match the reviewed PDF/u);
  assert.match(source, /unitType:\s*"sheet"/u);
  assert.match(source, /unitIndex,/u);
  assert.match(source, /--sheet-viewport, --range, and --unit-index must be provided together/u);
  assert.match(source, /provenanceVerification:\s*"not performed by this script"/u);
  assert.match(source, /print-page/u);
  assert.match(source, /RASTER_DPI = 96/u);
  assert.match(source, /function pngDimensions/u);
  assert.match(source, /Buffer\.from\(\[137, 80, 78, 71, 13, 10, 26, 10\]\)/u);
  assert.match(source, /Output path already exists/u);
});

test("Word native export tolerates version-specific print setting setters", () => {
  for (const setting of ["update links at open", "update fields at print", "update links at print"]) {
    assert.match(source, new RegExp(`try\\n\\s+set ${setting} of my wordSettings to false\\n\\s+end try`, "u"));
  }
});

test("native PDF import is restricted to page and slide formats and cannot be mixed with other capture modes", () => {
  const invalidInvocations = [
    {
      arguments: ["example.pptx", "output", "--native-pdf", "reviewed.pdf", "--allow-print-pages"],
      message: /--native-pdf cannot be combined/u,
    },
    {
      arguments: [
        "example.pptx", "output", "--native-pdf", "reviewed.pdf", "--sheet-viewport", "reviewed.png", "--range", "A1:H40",
        "--unit-index", "0",
      ],
      message: /--native-pdf cannot be combined/u,
    },
    {
      arguments: ["example.ods", "output", "--native-pdf", "reviewed.pdf", "--reviewed-by", "qa.bot"],
      message: /--native-pdf is valid only for PPTX, PPT, ODP, DOCX, DOC, ODT, RTF, KEY, PAGES, WPS, or DPS/u,
    },
    {
      arguments: [
        "example.odt", "output", "--sheet-viewport", "reviewed.png", "--range", "A1:H40", "--unit-index", "0",
        "--reviewed-by", "qa.bot",
      ],
      message: /--sheet-viewport is valid only for XLSX, XLS, ODS, NUMBERS, or ET/u,
    },
  ];
  for (const invocation of invalidInvocations) {
    const result = spawnSync(process.execPath, ["scripts/export-native-reference.mjs", ...invocation.arguments], {
      cwd: root,
      encoding: "utf8",
    });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, invocation.message);
  }
});

test("WPS references require explicit reviewed imports and are never GUI-automated", () => {
  for (const [fixture, option] of [
    ["example.wps", "--native-pdf"],
    ["example.et", "--sheet-viewport"],
    ["example.dps", "--native-pdf"],
  ]) {
    const result = spawnSync(process.execPath, [
      "scripts/export-native-reference.mjs", fixture, "output",
    ], { cwd: root, encoding: "utf8" });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, new RegExp(`require an explicitly reviewed ${option}`, "u"));
  }
  assert.doesNotMatch(source, /APPLE_SCRIPTS\["wps/u);
});


test("native workspace preserves the granted directory and excludes concurrent exports", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "native-workspace-"));
  try {
    const before = await stat(directory);
    const release = await acquireNativeWorkspace(directory);
    await assert.rejects(acquireNativeWorkspace(directory), /workspace is busy/u);
    await writeFile(resolve(directory, "source.pptx"), "source copy");
    await writeFile(resolve(directory, "reference.pdf"), "generated PDF");
    await writeFile(resolve(directory, "keep.txt"), "unrelated");
    await release();
    await assert.rejects(stat(resolve(directory, "reference.pdf")), { code: "ENOENT" });
    await assert.rejects(stat(resolve(directory, "source.pptx")), { code: "ENOENT" });
    assert.equal(await readFile(resolve(directory, "keep.txt"), "utf8"), "unrelated");
    const releaseAgain = await acquireNativeWorkspace(directory);
    assert.equal((await stat(directory)).ino, before.ino);
    await releaseAgain();
  } finally { await rm(directory, { recursive: true, force: true }); }
});


test("supplied encrypted presentation cannot export an unrelated active document", { skip: process.platform !== "darwin" }, async () => {
  const result = spawnSync(process.execPath, [
    "scripts/export-native-reference.mjs", "tests/fixtures/native-encrypted-presentation.pptx", "output/encrypted-native-must-not-exist",
  ], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Encrypted Office package requires credentials/u);
  await assert.rejects(stat(resolve(root, "output/encrypted-native-must-not-exist")), { code: "ENOENT" });
  assert.ok(source.indexOf("previousPresentationCount") < source.indexOf("save my openedPresentation"));
  assert.match(source, /Requested source identity mismatch/u);
  assert.match(source, /save as my openedDocument/u);
});


test("persistent native references restore identical bytes and reject tampering", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "native-library-"));
  try {
    const hash = bytes => createHash("sha256").update(bytes).digest("hex");
    const manifest = { source: { sha256: hash("source"), format: "xlsx", integrityVerifiedAfterExport: true },
      referenceApplication: { version: "1" }, environment: {}, rasterizer: { version: "1" },
      captureSemantics: { capture: "pdf-export", sourceIdentityCheck: "opened-workbook-source-path", allowPrintPages: true },
      pdfs: [{ file: "sheet-0001.pdf", sha256: hash("pdf") }],
      pages: [{ file: "sheet-0001-page-0001.png", sha256: hash("pixels") }] };
    await writeFile(resolve(directory, "reference.json"), JSON.stringify(manifest));
    await writeFile(resolve(directory, "sheet-0001.pdf"), "pdf");
    await writeFile(resolve(directory, "sheet-0001-page-0001.png"), "pixels");
    const library = resolve(directory, "library");
    // Keep the source outside the library to avoid recursive copies.
    const source = await mkdtemp(resolve(tmpdir(), "native-entry-"));
    try {
      for (const file of ["reference.json", "sheet-0001.pdf", "sheet-0001-page-0001.png"]) await writeFile(resolve(source, file), await readFile(resolve(directory, file)));
      assert.equal(await restoreNativeReference("0".repeat(64), resolve(directory, "miss"), library), false);
      const key = await cacheNativeReference(source, library);
      assert.equal(await cacheNativeReference(source, library), key);
      assert.equal(await restoreNativeReference(key, resolve(directory, "restored"), library), true);
      assert.equal(await readFile(resolve(directory, "restored/sheet-0001-page-0001.png"), "utf8"), "pixels");
      await writeFile(resolve(library, key, "sheet-0001-page-0001.png"), "corrupted");
      await assert.rejects(restoreNativeReference(key, resolve(directory, "bad"), library), /integrity failure/u);
      await assert.rejects(stat(resolve(directory, "bad")), { code: "ENOENT" });
    } finally { await rm(source, { recursive: true, force: true }); }
  } finally { await rm(directory, { recursive: true, force: true }); }
});
