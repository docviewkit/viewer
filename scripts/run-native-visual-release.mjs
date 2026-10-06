import { access } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const defaultSuiteRoot = resolve(root, "suites/native");
const defaultOutputRoot = resolve(root, "output/accuracy/native-release");
const releaseSuites = Object.freeze([
  Object.freeze({
    id: "office",
    file: "office-release.json",
    formats: Object.freeze(["pptx", "ppt", "docx", "doc", "xlsx", "xls", "odp", "odt", "ods"]),
  }),
  Object.freeze({
    id: "iwork",
    file: "iwork-release.json",
    formats: Object.freeze(["keynote", "pages", "numbers"]),
  }),
  Object.freeze({
    id: "wps",
    file: "wps-release.json",
    formats: Object.freeze(["wps", "et", "dps"]),
  }),
]);

function usage() {
  return [
    "Usage: node scripts/run-native-visual-release.mjs [options]",
    "",
    "Runs all native release suites and covers all fifteen golden formats:",
    `  Office: ${releaseSuites[0].formats.join(", ")}`,
    `  iWork:  ${releaseSuites[1].formats.join(", ")}`,
    `  WPS:    ${releaseSuites[2].formats.join(", ")}`,
    "",
    "Options:",
    "  --validate-only                 Validate manifests, hashes, references, and fonts only",
    "  --allow-missing                 Warn instead of failing when reviewed release manifests are absent",
    "  --suite-root <directory>        Directory containing office-release.json, iwork-release.json, and wps-release.json",
    "  --output <directory>            Root for per-suite candidates and reports",
    "  --font-root <directory>         Forwarded to the native capture runner",
    "  --font-manifest <file.json>     Forwarded to the native capture runner",
    "  --browser-executable <path>     Forwarded to the native capture runner",
    "  --help                          Show this help",
  ].join("\n");
}

function parseArguments(values) {
  let suiteRoot = defaultSuiteRoot;
  let outputRoot = defaultOutputRoot;
  let validateOnly = false;
  let allowMissing = false;
  const forwarded = [];
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (value === "--help") return { help: true };
    if (value === "--validate-only") {
      validateOnly = true;
      continue;
    }
    if (value === "--allow-missing") {
      allowMissing = true;
      continue;
    }
    if (["--suite-root", "--output", "--font-root", "--font-manifest", "--browser-executable"].includes(value)) {
      const optionValue = values[++index];
      if (optionValue === undefined || optionValue.startsWith("-")) throw new Error(`Missing ${value} value\n${usage()}`);
      if (value === "--suite-root") suiteRoot = resolve(optionValue);
      else if (value === "--output") outputRoot = resolve(optionValue);
      else forwarded.push(value, optionValue);
      continue;
    }
    throw new Error(`Unknown option ${value}\n${usage()}`);
  }
  return {
    help: false,
    suiteRoot,
    outputRoot,
    validateOnly,
    allowMissing,
    forwarded,
  };
}

async function exists(path) {
  try {
    await access(path);
    return true;
  } catch (cause) {
    if (cause?.code === "ENOENT") return false;
    throw cause;
  }
}

async function runRelease(options) {
  const results = [];
  for (const suite of releaseSuites) {
    const suitePath = resolve(options.suiteRoot, suite.file);
    if (!await exists(suitePath)) {
      results.push({ ...suite, status: "missing", suitePath });
      continue;
    }
    const arguments_ = [
      "scripts/capture-native-visual-suite.mjs",
      suitePath,
      ...options.forwarded,
      ...(options.validateOnly ? ["--validate-only"] : ["--output", resolve(options.outputRoot, suite.id)]),
    ];
    const result = spawnSync(process.execPath, arguments_, { cwd: root, stdio: "inherit" });
    if (result.error !== undefined) throw result.error;
    results.push({ ...suite, status: result.status === 0 ? "passed" : "failed", suitePath });
  }
  return results;
}

function printSummary(results, allowMissing) {
  console.log("\nNative visual release summary");
  for (const result of results) {
    console.log(`- ${result.id}: ${result.status}; formats=${result.formats.join(",")}; suite=${result.suitePath}`);
  }
  const missing = results.filter(({ status }) => status === "missing");
  if (missing.length !== 0) {
    const names = missing.map(({ file }) => file).join(", ");
    if (allowMissing) {
      console.warn(
        `::warning title=Native visual acceptance not certified::Missing reviewed release manifests: ${names}. `
        + "The release may continue but must be marked not-certified; fixtures, references, and goldens must not be synthesized.",
      );
    } else {
      console.log("Missing release manifests are a hard gate; reviewed fixtures, references, and goldens must not be synthesized by this runner.");
    }
  }
}

const options = parseArguments(process.argv.slice(2));
if (options.help) {
  console.log(usage());
} else {
  const results = await runRelease(options);
  printSummary(results, options.allowMissing);
  if (
    results.some(({ status }) => status === "failed")
    || (!options.allowMissing && results.some(({ status }) => status === "missing"))
  ) {
    process.exitCode = 1;
  }
}
