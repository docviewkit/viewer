import { execFile } from "node:child_process";
import { mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { basename, dirname, resolve } from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const dist = resolve(root, "dist");
const licenseRoot = resolve(dist, "third-party-licenses");
const metadataFiles = [
  resolve(dist, ".image-codec-metafile.json"),
];
const execute = promisify(execFile);

function packageName(input) {
  const normalized = input.replaceAll("\\", "/");
  const marker = "/node_modules/";
  const offset = normalized.lastIndexOf(marker);
  if (offset === -1) return null;
  const parts = normalized.slice(offset + marker.length).split("/");
  return parts[0].startsWith("@") ? `${parts[0]}/${parts[1]}` : parts[0];
}

async function findLicense(packageRoot) {
  const names = ["license", "license.md", "license.txt", "license-mit.txt", "licence", "licence.md", "copying"];
  const entries = (await readdir(packageRoot, { withFileTypes: true }))
    .filter((entry) => entry.isFile())
    .sort((left, right) => left.name.localeCompare(right.name));
  const actualNames = new Map();
  for (const entry of entries) {
    const normalized = entry.name.toLowerCase();
    if (!actualNames.has(normalized)) actualNames.set(normalized, entry.name);
  }
  for (const candidate of names) {
    const name = actualNames.get(candidate);
    if (name === undefined) continue;
    const content = await readFile(resolve(packageRoot, name), "utf8");
    if (content.trim() !== "") return { name, content };
  }
  return null;
}

const packages = new Set(["@kittl/little-cms", "@resvg/resvg-wasm", "mtx-decompressor"]);
for (const metadataFile of metadataFiles) {
  const metadata = JSON.parse(await readFile(metadataFile, "utf8"));
  for (const input of Object.keys(metadata.inputs)) {
    const name = packageName(resolve(root, input));
    if (name !== null) packages.add(name);
  }
}

await mkdir(licenseRoot, { recursive: true });
const rows = [];
const issues = [];
for (const name of [...packages].sort()) {
  const packageRoot = resolve(root, "node_modules", name);
  const manifest = JSON.parse(await readFile(resolve(packageRoot, "package.json"), "utf8"));
  const licenseId = manifest.license ?? (name === "khroma" ? "MIT" : null);
  const license = name === "@resvg/resvg-wasm"
    ? { name: "LICENSE", content: await readFile(resolve(root, "third_party/licenses/MPL-2.0.txt"), "utf8") }
    : await findLicense(packageRoot);
  if (typeof licenseId !== "string" || licenseId.trim() === "" || license === null) {
    issues.push({
      package: name,
      version: manifest.version,
      problem: typeof licenseId !== "string" ? "missing-license-metadata" : "missing-license-text",
    });
    rows.push(`| ${name} | ${manifest.version} | UNKNOWN | unavailable |`);
    continue;
  }
  const target = `${name.replaceAll("/", "_")}@${manifest.version}-${basename(license.name)}`;
  await writeFile(resolve(licenseRoot, target), license.content);
  rows.push(`| ${name} | ${manifest.version} | ${licenseId} | third-party-licenses/${target} |`);
}

await writeFile(resolve(licenseRoot, "Caladea-1.002-LICENSE.txt"),
  await readFile(resolve(root, "third_party/caladea/LICENSE.txt"), "utf8"));
rows.push("| Caladea | 1.002 | Apache-2.0 | third-party-licenses/Caladea-1.002-LICENSE.txt |");
await writeFile(resolve(licenseRoot, "Carlito-1.103-LICENSE.txt"),
  await readFile(resolve(root, "third_party/carlito/LICENSE.txt"), "utf8"));
rows.push("| Carlito | 1.103 | OFL-1.1 | third-party-licenses/Carlito-1.103-LICENSE.txt |");
await writeFile(resolve(licenseRoot, "PDFium-FoxitDingbats-LICENSE.txt"),
  await readFile(resolve(root, "third_party/pdf-standard-fonts/LICENSE_FOXIT"), "utf8"));
rows.push("| PDFium FoxitDingbats | SHA256 845c752392b6c914fb989c75a08b7792b88f542d2499042ef2889f8c814a16ed | BSD-3-Clause | third-party-licenses/PDFium-FoxitDingbats-LICENSE.txt |");

const cargo = process.env.CARGO ?? "cargo";
const rustBuilds = [
  [],
  ["--no-default-features", "--features", "odf-formats"],
  ["--no-default-features", "--features", "legacy-office-formats"],
  ["--no-default-features", "--features", "iwork-formats,pdf-formats"],
  ["--no-default-features", "--features", "xps-formats"],
  ["--no-default-features", "--features", "ofd-formats"],
];
const rustPackages = new Set();
for (const features of rustBuilds) {
  const { stdout } = await execute(cargo, [
    "tree",
    "--locked",
    "--target",
    "wasm32-unknown-unknown",
    "--edges",
    "normal",
    "--prefix",
    "none",
    "--format",
    "{p}",
    "-p",
    "office-viewer-core",
    ...features,
  ], { cwd: root, maxBuffer: 16 * 1024 * 1024 });
  for (const line of stdout.split("\n")) {
    const normalized = line.replace(/ \(\*\)$/, "").replace(/ \(proc-macro\)$/, "").trim();
    const match = /^([^ ]+) v([^ ]+)/.exec(normalized);
    if (match !== null && match[1] !== "office-viewer-core") {
      rustPackages.add(`${match[1]}@${match[2]}`);
    }
  }
}

const { stdout: cargoMetadataText } = await execute(cargo, [
  "metadata",
  "--format-version",
  "1",
  "--locked",
  "--all-features",
], { cwd: root, maxBuffer: 64 * 1024 * 1024 });
const cargoMetadata = JSON.parse(cargoMetadataText);
const rustMetadata = new Map(
  cargoMetadata.packages.map((manifest) => [`${manifest.name}@${manifest.version}`, manifest]),
);
for (const key of [...rustPackages].sort()) {
  const manifest = rustMetadata.get(key);
  if (manifest === undefined) {
    issues.push({ package: key, version: "UNKNOWN", problem: "missing-cargo-metadata" });
    rows.push(`| ${key} | UNKNOWN | UNKNOWN | unavailable |`);
    continue;
  }
  const packageRoot = dirname(manifest.manifest_path);
  const entries = (await readdir(packageRoot, { withFileTypes: true }))
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .filter((name) => /^(license|licence|copying|unlicense)/i.test(name))
    .sort((left, right) => left.localeCompare(right));
  if (typeof manifest.license_file === "string") {
    const explicit = basename(manifest.license_file);
    if (!entries.includes(explicit)) entries.unshift(explicit);
  }
  let licenseSources = entries.map((name) => ({ name, path: resolve(packageRoot, name) }));
  if (licenseSources.length === 0 && typeof manifest.repository === "string") {
    const repository = manifest.repository.replace(/\/$/, "");
    for (const peer of cargoMetadata.packages) {
      if (peer.name === manifest.name
        || typeof peer.repository !== "string"
        || peer.repository.replace(/\/$/, "") !== repository) continue;
      const peerRoot = dirname(peer.manifest_path);
      const peerLicenses = (await readdir(peerRoot, { withFileTypes: true }))
        .filter((entry) => entry.isFile() && /^(license|licence|copying|unlicense)/i.test(entry.name))
        .map((entry) => ({ name: entry.name, path: resolve(peerRoot, entry.name) }));
      if (peerLicenses.length !== 0) {
        licenseSources = peerLicenses;
        break;
      }
    }
  }
  if (licenseSources.length === 0 && manifest.name === "ironcalc_base") {
    licenseSources.push({
      name: "LICENSE-MIT",
      path: resolve(root, "third_party/licenses/ironcalc-MIT.txt"),
    });
  }
  const licenses = [];
  for (const { name, path } of licenseSources) {
    const content = await readFile(path, "utf8");
    if (content.trim() !== "") licenses.push({ name, content });
  }
  if (typeof manifest.license !== "string" || manifest.license.trim() === "" || licenses.length === 0) {
    issues.push({
      package: manifest.name,
      version: manifest.version,
      problem: typeof manifest.license !== "string" ? "missing-license-metadata" : "missing-license-text",
    });
    rows.push(`| ${manifest.name} | ${manifest.version} | ${manifest.license ?? "UNKNOWN"} | unavailable |`);
    continue;
  }
  const target = `rust-${manifest.name}@${manifest.version}-LICENSES.txt`;
  await writeFile(
    resolve(licenseRoot, target),
    licenses.map(({ name, content }) => `===== ${name} =====\n${content.trim()}\n`).join("\n"),
  );
  rows.push(`| ${manifest.name} | ${manifest.version} | ${manifest.license} | third-party-licenses/${target} |`);
}

const notice = [
  "# Bundled third-party software",
  "",
  "This inventory is generated from the JavaScript bundle dependency graphs and every Rust package in the shipped WebAssembly feature trees. The corresponding license texts are shipped with the package.",
  "",
  "| Package | Version | License | License text |",
  "| --- | --- | --- | --- |",
  ...rows,
  "",
].join("\n");
await writeFile(resolve(dist, "THIRD_PARTY_NOTICES.md"), notice);
await writeFile(resolve(dist, "third-party-license-issues.json"), `${JSON.stringify(issues, null, 2)}\n`);
if (issues.length !== 0) {
  console.warn(`Bundled dependency license gate has ${issues.length} unresolved item(s); commercial release verification will fail.`);
}
await Promise.all(metadataFiles.map((path) => rm(path, { force: true })));
