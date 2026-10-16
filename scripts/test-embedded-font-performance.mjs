// Real document regression; embedded fonts stay local and are never committed.
// Run after build: node scripts/test-embedded-font-performance.mjs <pptx>
//   [--output report.json] [--baseline before.json]
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { readZipEntries } from "./accuracy-metamorphic.mjs";

const [path, ...args] = process.argv.slice(2);
assert.ok(path, "Supply the original PPTX path");
const options = {};
for (let i = 0; i < args.length; i += 2) {
  assert.ok(["--output", "--baseline", "--decoder"].includes(args[i]) && args[i + 1], "Invalid option");
  options[args[i].slice(2)] = args[i + 1];
}
const { eotToTtf, parseEotMetadata } = await import(options.decoder ? pathToFileURL(resolve(options.decoder)).href : "../dist/mtx-decompressor.js");
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const input = await readFile(path);
const entries = readZipEntries(input).filter(({ name }) => /^ppt\/fonts\/.*\.fntdata$/u.test(name));
assert.ok(entries.length > 0, "The real document must contain embedded fonts");
const results = [];
for (const { name, data } of entries) {
  const metadata = parseEotMetadata(data);
  let allocatedBytes = 0;
  const NativeUint8Array = globalThis.Uint8Array;
  const started = performance.now();
  let decoded;
  try {
    // Count actual allocations at the runtime boundary, independently of the
    // decoder implementation and machine speed. Views do not allocate storage.
    globalThis.Uint8Array = class extends NativeUint8Array {
      constructor(...values) {
        super(...values);
        if (typeof values[0] === "number") allocatedBytes += this.byteLength;
      }
    };
    decoded = eotToTtf(data);
  } finally {
    globalThis.Uint8Array = NativeUint8Array;
  }
  results.push({ name, compressed: metadata.compressed, inputBytes: data.length,
    outputBytes: decoded.length, sha256: hash(decoded), allocatedBytes,
    decodeMs: performance.now() - started });
}
const report = { inputSha256: hash(input), results };
if (options.output) await writeFile(options.output, JSON.stringify(report, null, 2) + "\n");
console.log(JSON.stringify(report, null, 2));
if (options.baseline) {
  const baseline = JSON.parse(await readFile(options.baseline, "utf8"));
  assert.equal(report.inputSha256, baseline.inputSha256, "Must compare the unchanged real document");
  assert.deepEqual(results.map(({ name, outputBytes, sha256 }) => ({ name, outputBytes, sha256 })),
    baseline.results.map(({ name, outputBytes, sha256 }) => ({ name, outputBytes, sha256 })),
    "Every decoded font must be byte-identical to the pre-change decoder");
}
const failures = results.filter(result => result.allocatedBytes > Math.max(16 * 1024 ** 2, result.outputBytes * 16));
assert.deepEqual(failures.map(({ name, allocatedBytes }) => ({ name, allocatedBytes })), [],
  "Glyph decoding must use linear allocation volume, not repeatedly copy the growing font");
