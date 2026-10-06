import assert from "node:assert/strict";
import test from "node:test";

import { compileCore } from "../dist/core.js";

const EMPTY_WASM = "AGFzbQEAAAA=";

test("URL Wasm compilation is shared and falls back when streaming MIME is unavailable", async () => {
  const streamingUrl = `data:application/wasm;base64,${EMPTY_WASM}`;
  const first = await compileCore(streamingUrl);
  const second = await compileCore(streamingUrl);
  assert.equal(second, first);

  const buffered = await compileCore(`data:application/octet-stream;base64,${EMPTY_WASM}`);
  assert.ok(buffered instanceof WebAssembly.Module);
});
