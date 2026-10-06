import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { createOfficeEngine } from "../dist/index.js";
import { evaluateLicense } from "../dist/license.js";

test("Apache-2.0 opens a real presentation through the public Engine without runtime licensing", async () => {
  const bytes = await readFile(new URL("./fixtures/visual-baseline.pptx", import.meta.url));
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  for (const license of [undefined, { token: "expired-or-invalid", publicKey: "", origin: "https://customer.example.com" }]) {
    const engine = await createOfficeEngine({ wasm, execution: "inline", license });
    try {
      const document = await engine.open(bytes);
      assert.equal(document.info.format, "pptx");
      assert.ok(document.info.units.length > 0);
      assert.ok((await document.listObjects({ unitIndex: 0 })).length > 0);
      document.close();
      const state = await evaluateLicense(license);
      assert.equal(state.branding, "hidden");
      assert.deepEqual(state.entitlements, {
        viewer: true, engine: true, businessSlots: true, advancedCustomization: true,
      });
    } finally {
      engine.close();
    }
  }
});
