import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");

test("native release runner exposes one command for all fifteen formats", () => {
  const result = spawnSync(process.execPath, ["scripts/run-native-visual-release.mjs", "--help"], {
    cwd: root,
    encoding: "utf8",
  });
  assert.equal(result.status, 0, result.stderr);
  for (const format of [
    "pptx", "ppt", "docx", "doc", "xlsx", "xls", "odp", "odt", "ods",
    "keynote", "pages", "numbers", "wps", "et", "dps",
  ]) {
    assert.match(result.stdout, new RegExp(`\\b${format}\\b`, "u"));
  }
});

test("native release runner reports every missing manifest without raw ENOENT failures", async () => {
  const empty = await mkdtemp(resolve(tmpdir(), "officeviewer-native-release-"));
  try {
    const result = spawnSync(process.execPath, [
      "scripts/run-native-visual-release.mjs",
      "--validate-only",
      "--suite-root", empty,
    ], { cwd: root, encoding: "utf8" });
    assert.equal(result.status, 1);
    assert.match(result.stdout, /office: missing/u);
    assert.match(result.stdout, /iwork: missing/u);
    assert.match(result.stdout, /wps: missing/u);
    assert.match(result.stdout, /office-release\.json/u);
    assert.match(result.stdout, /iwork-release\.json/u);
    assert.match(result.stdout, /wps-release\.json/u);
    assert.match(result.stdout, /hard gate/u);
    assert.doesNotMatch(`${result.stdout}\n${result.stderr}`, /ENOENT/u);
  } finally {
    await rm(empty, { recursive: true, force: true });
  }
});

test("native release advisory mode permits only missing reviewed manifests", async () => {
  const empty = await mkdtemp(resolve(tmpdir(), "officeviewer-native-release-advisory-"));
  try {
    const missing = spawnSync(process.execPath, [
      "scripts/run-native-visual-release.mjs",
      "--validate-only",
      "--allow-missing",
      "--suite-root", empty,
    ], { cwd: root, encoding: "utf8" });
    assert.equal(missing.status, 0, `${missing.stdout}\n${missing.stderr}`);
    assert.match(missing.stderr, /not certified/u);
    assert.match(missing.stderr, /must be marked not-certified/u);

    await writeFile(resolve(empty, "office-release.json"), "{}\n");
    const invalid = spawnSync(process.execPath, [
      "scripts/run-native-visual-release.mjs",
      "--validate-only",
      "--allow-missing",
      "--suite-root", empty,
    ], { cwd: root, encoding: "utf8" });
    assert.equal(invalid.status, 1);
    assert.match(invalid.stdout, /office: failed/u);
  } finally {
    await rm(empty, { recursive: true, force: true });
  }
});
