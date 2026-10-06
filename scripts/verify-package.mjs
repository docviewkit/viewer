import { mkdtemp, readFile, rm } from "node:fs/promises";
import { execFileSync } from "node:child_process";
import { resolve } from "node:path";
import { tmpdir } from "node:os";
import { fileURLToPath } from "node:url";
import { validateBundledFonts } from "./prepare-release.mjs";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const packageSizeBudget = Object.freeze({
  // Linux release output is slightly larger than the macOS build.
  packed: 8_200_000,
  // Shared XLS chart rendering and XPS compatibility add about 370 KiB of Wasm.
  unpacked: 22_300_000,
});
const manifest = JSON.parse(await readFile(resolve(root, "package.json"), "utf8"));
const issues = JSON.parse(await readFile(resolve(root, "dist/third-party-license-issues.json"), "utf8"));
if (manifest.license !== "Apache-2.0") {
  throw new Error("DocViewKit must declare Apache-2.0");
}
if (issues.length !== 0) {
  throw new Error(`Third-party license inventory has unresolved entries: ${issues.map((entry) => entry.package).join(", ")}`);
}

const cache = await mkdtemp(resolve(tmpdir(), "docviewkit-pack-"));
try {
  const output = execFileSync("npm", ["pack", "--dry-run", "--json", "--cache", cache], {
    cwd: root,
    encoding: "utf8",
  });
  const [packed] = JSON.parse(output);
  const files = packed.files.map(({ path }) => path);
  await validateBundledFonts(resolve(root, "dist"), files.filter((path) => path.startsWith("dist/")).map((path) => path.slice(5)));
  const excludedRuntimeFiles = new Set([
    "dist/image-codecs.js", "dist/image-codecs.d.ts",
    "dist/metafile-safety.js", "dist/metafile-safety.d.ts",
    "dist/worker-protocol.js", "dist/worker-protocol.d.ts",
    "dist/noto-sans-hans.otf",
    "dist/office-viewer-iwork.wasm",
  ]);
  const forbidden = files.filter((path) => path.endsWith(".map")
    || (/\.(?:ttf|otf|ttc|woff2?)$/iu.test(path) && !path.startsWith("dist/"))
    || excludedRuntimeFiles.has(path)
    || /(^|\/)(?:examples|research|scripts|src|tests)(\/|$)/u.test(path)
    || (path.endsWith(".ts") && !path.endsWith(".d.ts")));
  if (forbidden.length !== 0) throw new Error(`Forbidden package files: ${forbidden.join(", ")}`);
  for (const required of [
    "LICENSE",
    "NOTICE",
    "dist/viewer.js",
    "dist/worker.js",
    "dist/office-viewer-core.wasm",
    "dist/office-viewer-calc.wasm",
    "dist/office-viewer-xps.wasm",
    "dist/xps-formats.js",
    "dist/office-viewer-ofd.wasm",
    "dist/ofd-formats.js",
    "dist/THIRD_PARTY_NOTICES.md",
  ]) {
    if (!files.includes(required)) throw new Error(`Package is missing ${required}`);
  }
  if (
    packed.size > packageSizeBudget.packed
    || packed.unpackedSize > packageSizeBudget.unpacked
  ) {
    throw new Error(
      `Package size budget exceeded: packed=${packed.size}/${packageSizeBudget.packed}, `
      + `unpacked=${packed.unpackedSize}/${packageSizeBudget.unpacked}`,
    );
  }
  const viewer = await readFile(resolve(root, "dist/viewer.js"), "utf8");
  if (/from\s*["']\.\/engine\.js["']/u.test(viewer)) {
    throw new Error("Viewer bundle still exposes a source-module dependency");
  }
  const worker = await readFile(resolve(root, "dist/worker.js"), "utf8");
  for (const chunk of [...viewer.matchAll(/from["']\.\/(runtime-[^"']+\.js)["']/gu),
    ...worker.matchAll(/from["']\.\/(runtime-[^"']+\.js)["']/gu)].map((match) => match[1])) {
    if (!files.includes(`dist/${chunk}`)) throw new Error(`Package is missing shared runtime chunk ${chunk}`);
  }
  console.log(`Package gate passed: ${files.length} files, ${packed.size} bytes packed, ${packed.unpackedSize} bytes unpacked`);
} finally {
  await rm(cache, { recursive: true, force: true });
}
