import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { createOfficeEngine } from "../dist/engine.js";
import { DEFAULT_LIMITS, assertInputSize, resolveLimits } from "../dist/core.js";

// Keep this private, 166 MiB customer presentation outside the repository.
// Run with DOCVIEWKIT_LARGE_PPTX pointing to the original, unmodified file.
test("opens all 35 slides of the supplied video-heavy Honda presentation", {
  skip: !process.env.DOCVIEWKIT_LARGE_PPTX,
}, async () => {
  const bytes = await readFile(process.env.DOCVIEWKIT_LARGE_PPTX);
  assert.equal(bytes.byteLength, 173545373, "use the original customer file");
  const engine = await createOfficeEngine({ execution: "inline", wasm: await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)) });
  let document;
  try {
    document = await engine.open(bytes);
    assert.equal(document.info.units.length, 35);
    for (let unitIndex = 0; unitIndex < 35; unitIndex++) {
      const objects = await document.listObjects({ unitIndex });
      assert.ok(objects.length > 0, `slide ${unitIndex + 1} must retain its parsed content`);
    }
    assert.ok(document.diagnostics().every(d => d.severity !== "fatal"));
  } finally { document?.close(); engine.close(); }
});

test("default input budget admits video-heavy packages and preserves explicit and hard limits", () => {
  assertInputSize(173545373, DEFAULT_LIMITS.inputBytes);
  assertInputSize(DEFAULT_LIMITS.inputBytes, DEFAULT_LIMITS.inputBytes);
  assert.throws(() => assertInputSize(DEFAULT_LIMITS.inputBytes + 1, DEFAULT_LIMITS.inputBytes),
    error => error.code === "INPUT_SIZE_LIMIT");
  const restricted = resolveLimits({ inputBytes: 128 * 1024 * 1024 });
  assert.throws(() => assertInputSize(173545373, restricted.inputBytes),
    error => error.code === "INPUT_SIZE_LIMIT");
  assert.throws(() => resolveLimits({ inputBytes: 512 * 1024 * 1024 + 1 }),
    error => error.code === "INVALID_LIMITS");
});
