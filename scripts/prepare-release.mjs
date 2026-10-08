import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import {
  cp,
  mkdir,
  readFile,
  readdir,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import {
  basename,
  isAbsolute,
  relative,
  resolve,
  sep,
} from "node:path";
import { tmpdir } from "node:os";

const root = resolve(import.meta.dirname, "..");
const tagPattern = /^v(\d+\.\d+\.\d+(?:-[0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?)$/u;
const nativeVisualStatuses = new Set(["passed", "not-certified", "not-evaluated"]);
const runtimeFiles = new Set([
  "index.js",
  "viewer.js",
  "accuracy.js",
  "worker.js",
  "image-codec-worker.js",
  "extended-formats.js",
  "wps-formats.js",
  "xps-formats.js",
  "ofd-formats.js",
  "types.js",
  "version.json",
]);

// Two metric-compatible Office fallbacks only; adding fonts requires a benefit/license review.
const bundledFontFamilies = [
  ["Carlito", "1.103", "OFL-1.1", 832 * 1024, "bbcf8ce9c8acc91355ab5d263ec9e37d498de2222336e7d42e589f791afe56c4"],
  ["Caladea", "1.002", "Apache-2.0", 64 * 1024, "6f1041c12f758ed86d804acbcb54ad822d053fa15520184c28c3b8eabb8f66f6"],
];
const bundledFontVariants = ["Regular", "Bold", "Italic", "BoldItalic"];

export async function validateBundledFonts(directory = resolve(root, "dist"), files) {
  const inventory = (files ?? await readdir(directory, { recursive: true })).map((name) => name.replaceAll("\\", "/"));
  const fonts = inventory.filter((name) => /\.(?:ttf|otf|ttc|woff2?)$/iu.test(name)).sort();
  const approved = bundledFontFamilies.flatMap(([family]) => bundledFontVariants.map((variant) => `${family}-${variant}.ttf`)).sort();
  if (JSON.stringify(fonts) !== JSON.stringify(approved)) {
    throw new Error(`Bundled font inventory must contain only the approved 2 families / 8 faces: ${fonts.join(", ")}`);
  }
  const sizes = await Promise.all(fonts.map(async (name) => (await stat(resolve(directory, name))).size));
  if (sizes.reduce((total, size) => total + size, 0) > 3 * 1024 * 1024) {
    throw new Error("Bundled fonts exceed the 3 MiB total budget");
  }
  const notices = await readFile(resolve(directory, "THIRD_PARTY_NOTICES.md"), "utf8");
  for (const [family, version, license, maxBytes, licenseHash] of bundledFontFamilies) {
    const source = resolve(root, "third_party", family.toLowerCase());
    const provenance = await readFile(resolve(source, "README.md"), "utf8");
    for (const variant of bundledFontVariants) {
      const name = `${family}-${variant}.ttf`;
      if (sizes[fonts.indexOf(name)] > maxBytes) throw new Error(`Bundled font exceeds its size budget: ${name}`);
      const expected = provenance.split("\n").find((line) => line.startsWith(`- \`${name}\`:`))?.match(/[0-9a-f]{64}/u)?.[0];
      const actual = createHash("sha256").update(await readFile(resolve(directory, name))).digest("hex");
      if (expected !== actual) throw new Error(`Bundled font differs from the reviewed upstream file: ${name}`);
    }
    const licensePath = `third-party-licenses/${family}-${version}-LICENSE.txt`;
    if (!inventory.includes(licensePath)) throw new Error(`Bundled font license is missing: ${licensePath}`);
    for (const path of [resolve(source, "LICENSE.txt"), resolve(directory, licensePath)]) {
      if (createHash("sha256").update(await readFile(path)).digest("hex") !== licenseHash) {
        throw new Error(`Bundled font license differs from the reviewed text: ${path}`);
      }
    }
    if (!notices.includes(`| ${family} | ${version} | ${license} | ${licensePath} |`)) {
      throw new Error(`Bundled font attribution is missing: ${family}`);
    }
  }
}

function argument(name) {
  const index = process.argv.indexOf(name);
  return index === -1 ? undefined : process.argv[index + 1];
}

function packageVersion(lock) {
  return lock.packages?.[""]?.version ?? lock.version;
}

function tomlVersion(source, packageName) {
  const packageBlock = source.match(
    new RegExp(`\\[\\[package\\]\\][\\s\\S]*?name = "${packageName}"[\\s\\S]*?version = "([^"]+)"`, "u"),
  );
  return packageBlock?.[1];
}

function collectPackageTargets(value, targets) {
  if (typeof value === "string") {
    targets.add(value);
    return;
  }
  if (Array.isArray(value)) {
    for (const entry of value) collectPackageTargets(entry, targets);
    return;
  }
  if (value !== null && typeof value === "object") {
    for (const entry of Object.values(value)) collectPackageTargets(entry, targets);
  }
}

function localPackageEntryFiles(manifest) {
  const targets = new Set(["package.json"]);
  for (const key of ["main", "module", "types"]) {
    if (typeof manifest[key] === "string") targets.add(manifest[key]);
  }
  collectPackageTargets(manifest.exports, targets);
  return [...targets]
    .map((entry) => entry.startsWith("./") ? entry.slice(2) : entry)
    .filter((entry) => entry.length !== 0 && !entry.includes("*") && !entry.endsWith("/"));
}

function isTracked(projectRoot, file) {
  try {
    execFileSync(
      "git",
      ["ls-files", "--error-unmatch", "--", relative(projectRoot, file)],
      { cwd: projectRoot, stdio: "ignore" },
    );
    return true;
  } catch {
    return false;
  }
}

export async function validateLocalDependencies(projectRoot = root) {
  const manifest = JSON.parse(await readFile(resolve(projectRoot, "package.json"), "utf8"));
  const dependencies = {
    ...manifest.dependencies,
    ...manifest.optionalDependencies,
    ...manifest.devDependencies,
  };
  for (const [name, specifier] of Object.entries(dependencies)) {
    if (typeof specifier !== "string" || !specifier.startsWith("file:")) continue;
    const packageRoot = resolve(projectRoot, specifier.slice("file:".length));
    const packageRelative = relative(projectRoot, packageRoot);
    if (
      packageRelative === ".."
      || packageRelative.startsWith(`..${sep}`)
      || isAbsolute(packageRelative)
    ) {
      throw new Error(`Local dependency ${name} escapes the repository: ${specifier}`);
    }
    const packageManifestPath = resolve(packageRoot, "package.json");
    const packageManifest = JSON.parse(await readFile(packageManifestPath, "utf8"));
    for (const entry of localPackageEntryFiles(packageManifest)) {
      const file = resolve(packageRoot, entry);
      const fileRelative = relative(packageRoot, file);
      if (
        fileRelative === ".."
        || fileRelative.startsWith(`..${sep}`)
        || isAbsolute(fileRelative)
      ) {
        throw new Error(`Local dependency ${name} exports a path outside its package: ${entry}`);
      }
      const metadata = await stat(file).catch(() => undefined);
      if (metadata?.isFile() !== true) {
        throw new Error(`Local dependency ${name} is missing exported file: ${relative(projectRoot, file)}`);
      }
      if (!isTracked(projectRoot, file)) {
        throw new Error(`Local dependency ${name} exports an untracked file: ${relative(projectRoot, file)}`);
      }
    }
  }
}

async function releaseVersion(tag) {
  const match = tagPattern.exec(tag);
  if (match === null) {
    throw new Error(`Release tag must match vX.Y.Z or vX.Y.Z-prerelease: ${tag}`);
  }
  const expected = match[1];
  const manifest = JSON.parse(await readFile(resolve(root, "package.json"), "utf8"));
  const lock = JSON.parse(await readFile(resolve(root, "package-lock.json"), "utf8"));
  const cargoManifest = await readFile(resolve(root, "Cargo.toml"), "utf8");
  const cargoLock = await readFile(resolve(root, "Cargo.lock"), "utf8");
  const versions = new Map([
    ["package.json", manifest.version],
    ["package-lock.json", packageVersion(lock)],
    ["Cargo.toml", cargoManifest.match(/^\s*version\s*=\s*"([^"]+)"/mu)?.[1]],
    ["Cargo.lock", tomlVersion(cargoLock, "office-viewer-core")],
  ]);
  const mismatches = [...versions].filter(([, version]) => version !== expected);
  if (mismatches.length !== 0) {
    throw new Error(
      `Tag ${tag} does not match ${mismatches.map(([file, version]) => `${file} (${version ?? "missing"})`).join(", ")}`,
    );
  }
  await validateLocalDependencies();
  return { version: expected, prerelease: expected.includes("-") };
}

async function copyFreeRuntime(target) {
  const dist = resolve(root, "dist");
  const entries = await readdir(dist, { withFileTypes: true });
  for (const entry of entries) {
    if (
      (entry.isFile() && (
        runtimeFiles.has(entry.name)
        || /^runtime-[A-Z0-9]+\.js$/u.test(entry.name)
        || entry.name.endsWith(".d.ts")
        || entry.name.endsWith(".wasm")
        || /^(?:Carlito|Caladea)-(?:Regular|Bold|Italic|BoldItalic)\.ttf$/u.test(entry.name)
      ))
      || (entry.isDirectory() && entry.name === "third-party-licenses")
    ) {
      await cp(resolve(dist, entry.name), resolve(target, entry.name), { recursive: entry.isDirectory() });
    }
  }
  await cp(resolve(dist, "THIRD_PARTY_NOTICES.md"), resolve(target, "THIRD_PARTY_NOTICES.md"));
  for (const entry of ["LICENSE", "NOTICE"]) await cp(resolve(root, entry), resolve(target, entry));
  await validateBundledFonts(target);
}

async function replaceReleaseTokens(directory, replacements) {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = resolve(directory, entry.name);
    if (entry.isDirectory()) {
      await replaceReleaseTokens(path, replacements);
      continue;
    }
    if (!/\.(?:html|js|css|json|md)$/u.test(entry.name)) continue;
    let source = await readFile(path, "utf8");
    for (const [token, value] of replacements) source = source.replaceAll(token, value);
    await writeFile(path, source);
  }
}

function npmPack(directory, destination) {
  const cache = mkdtempSync(resolve(tmpdir(), "docviewkit-npm-cache-"));
  try {
    const output = execFileSync(
      process.platform === "win32" ? "npm.cmd" : "npm",
      ["pack", "--pack-destination", destination, "--ignore-scripts", "--cache", cache],
      { cwd: directory, encoding: "utf8" },
    );
    return basename(output.trim().split(/\r?\n/u).at(-1));
  } finally {
    rmSync(cache, { recursive: true, force: true });
  }
}

async function prepareFreePackage(output, version) {
  const stage = resolve(output, "free-package");
  await mkdir(stage, { recursive: true });
  await copyFreeRuntime(stage);
  await cp(resolve(root, "README.md"), resolve(stage, "README.md"));
  await cp(resolve(root, "release/FREE_VIEWER_LICENSE.md"), resolve(stage, "FREE_VIEWER_LICENSE.md"));
  const manifest = {
    name: "@docviewkit/viewer",
    version,
    description: "Frontend-local, read-only Office, PDF and OFD Viewer Web Component",
    keywords: [
      "document-viewer",
      "document-preview",
      "file-viewer",
      "file-preview",
      "office-viewer",
      "office-preview",
      "word-viewer",
      "excel-viewer",
      "powerpoint-viewer",
      "docx-viewer",
      "xlsx-viewer",
      "pptx-viewer",
      "pdf-viewer",
      "pdf-preview",
      "ofd-viewer",
      "ofd-preview",
      "spreadsheet-viewer",
      "presentation-viewer",
      "web-component",
      "web-components",
      "custom-element",
      "browser",
      "frontend",
      "client-side",
      "frontend-only",
      "local-first",
      "offline",
      "no-server",
      "read-only",
      "javascript",
      "typescript",
      "webassembly",
      "wasm",
      "canvas",
      "web-worker",
      "lazy-loading",
      "webview",
      "electron",
      "tauri",
      "ionic",
      "capacitor",
      "react",
      "vue",
      "angular",
      "svelte",
      "office",
      "microsoft-office",
      "word",
      "excel",
      "powerpoint",
      "ooxml",
      "doc",
      "docx",
      "docm",
      "dotx",
      "dotm",
      "xls",
      "xlsx",
      "xlsm",
      "xltx",
      "xltm",
      "ppt",
      "pptx",
      "pptm",
      "ppsx",
      "ppsm",
      "potx",
      "potm",
      "opendocument",
      "odf",
      "libreoffice",
      "openoffice",
      "odt",
      "ott",
      "fodt",
      "ods",
      "ots",
      "fods",
      "odp",
      "otp",
      "fodp",
      "odg",
      "otg",
      "fodg",
      "iwork",
      "pages",
      "numbers",
      "keynote",
      "wps-office",
      "wps",
      "et",
      "dps",
      "pdf",
      "xps",
      "oxps",
      "openxps",
      "ofd",
      "rtf",
      "csv",
      "文档预览",
      "文件预览",
      "在线预览",
      "办公文档预览",
      "纯前端预览",
      "word预览",
      "excel预览",
      "ppt预览",
      "pdf预览",
      "ofd预览",
    ],
    type: "module",
    sideEffects: ["./viewer.js"],
    main: "./viewer.js",
    types: "./viewer.d.ts",
    exports: {
      "./engine": { types: "./index.d.ts", import: "./index.js" },
      "./viewer": { types: "./viewer.d.ts", import: "./viewer.js" },
      "./accuracy": { types: "./accuracy.d.ts", import: "./accuracy.js" },
      ".": {
        types: "./viewer.d.ts",
        import: "./viewer.js",
      },
      "./extended-formats": {
        types: "./extended-formats.d.ts",
        import: "./extended-formats.js",
      },
      "./wps-formats": {
        types: "./wps-formats.d.ts",
        import: "./wps-formats.js",
      },
      "./xps-formats": {
        types: "./xps-formats.d.ts",
        import: "./xps-formats.js",
      },
      "./ofd-formats": {
        types: "./ofd-formats.d.ts",
        import: "./ofd-formats.js",
      },
      "./package.json": "./package.json",
    },
    repository: {
      type: "git",
      url: "git+https://github.com/docviewkit/viewer.git",
    },
    homepage: "https://docviewkit.com/en/demo/",
    bugs: {
      url: "https://github.com/docviewkit/viewer/issues",
    },
    license: "Apache-2.0",
    engines: {
      node: ">=20",
    },
  };
  await writeFile(resolve(stage, "package.json"), `${JSON.stringify(manifest, null, 2)}\n`);
  const archive = npmPack(stage, output);
  return { stage, archive };
}

async function preparePages(output, version, tag, freeStage) {
  const pages = resolve(output, "pages");
  await cp(resolve(root, "release/pages"), pages, { recursive: true });
  const sdk = resolve(pages, "sdk", `v${version}`);
  await mkdir(sdk, { recursive: true });
  await copyFreeRuntime(sdk);
  await writeFile(
    resolve(pages, "version.json"),
    `${JSON.stringify({ name: "@docviewkit/viewer", version, tag }, null, 2)}\n`,
  );
  for (const entry of ["FREE_VIEWER_LICENSE.md", "LICENSE", "NOTICE"]) await cp(resolve(freeStage, entry), resolve(pages, entry));
  await replaceReleaseTokens(pages, new Map([
    ["__DOCVIEWKIT_VERSION__", version],
    ["__DOCVIEWKIT_TAG__", tag],
  ]));
  return pages;
}

export async function prepareRelease({
  tag,
  output,
  verifyOnly = false,
  nativeVisualStatus = "not-evaluated",
}) {
  if (!nativeVisualStatuses.has(nativeVisualStatus)) {
    throw new Error(`Unsupported native visual acceptance status: ${nativeVisualStatus}`);
  }
  const release = await releaseVersion(tag);
  if (verifyOnly) return release;
  await validateBundledFonts();
  await rm(output, { recursive: true, force: true });
  await mkdir(output, { recursive: true });
  const free = await prepareFreePackage(output, release.version);
  const sdkArchive = npmPack(root, output);
  await preparePages(output, release.version, tag, free.stage);
  const result = {
    tag,
    version: release.version,
    prerelease: release.prerelease,
    npmTag: release.prerelease ? "next" : "latest",
    freeArchive: free.archive,
    sdkArchive,
    nativeVisualAcceptance: {
      policy: "native-visual-v1",
      status: nativeVisualStatus,
    },
  };
  await writeFile(resolve(output, "release.json"), `${JSON.stringify(result, null, 2)}\n`);
  return result;
}

if (process.argv[1] === import.meta.filename) {
  const tag = argument("--tag") ?? process.env.GITHUB_REF_NAME;
  if (tag === undefined) throw new Error("Pass --tag vX.Y.Z or set GITHUB_REF_NAME");
  const output = resolve(root, argument("--output") ?? "release-output");
  const result = await prepareRelease({
    tag,
    output,
    verifyOnly: process.argv.includes("--verify-only"),
    nativeVisualStatus: argument("--native-visual-status") ?? "not-evaluated",
  });
  console.log(JSON.stringify(result));
}
