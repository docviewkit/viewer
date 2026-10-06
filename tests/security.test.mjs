import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";
import { createZip } from "./zip-fixture.mjs";

async function inlineEngine(wasmFile = "office-viewer-core.wasm") {
  const wasm = await readFile(new URL(`../dist/${wasmFile}`, import.meta.url));
  return createOfficeEngine({ wasm, execution: "inline" });
}

test("rejects encrypted ODF before document XML is parsed", async () => {
  const encrypted = createZip({
    mimetype: "application/vnd.oasis.opendocument.text",
    "META-INF/manifest.xml": `<?xml version="1.0"?>
      <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
        <manifest:file-entry manifest:full-path="content.xml">
          <manifest:encryption-data manifest:checksum-type="SHA256/1K"/>
        </manifest:file-entry>
      </manifest:manifest>`,
    "content.xml": new Uint8Array([0xde, 0xad, 0xbe, 0xef]),
  }, { compress: false });
  const engine = await inlineEngine("office-viewer-odf.wasm");
  try {
    await assert.rejects(
      engine.open(encrypted),
      (error) => error?.code === "PACKAGE_ENCRYPTED"
        && error.diagnostics?.[0]?.phase === "security",
    );
  } finally {
    engine.close();
  }
});

test("opens macro-enabled OOXML as inert content and reports the blocked payload", async () => {
  const docm = createZip({
    "[Content_Types].xml": `<?xml version="1.0"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/>
      </Types>`,
    "_rels/.rels": `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>`,
    "word/document.xml": `<?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>Inert macro document</w:t></w:r></w:p></w:body>
      </w:document>`,
    "word/vbaProject.bin": new Uint8Array([1, 2, 3, 4]),
  });
  const engine = await inlineEngine();
  const document = await engine.open(docm);
  try {
    assert.equal(document.info.format, "docx");
    const findings = document.diagnostics().filter((diagnostic) => diagnostic.code === "ACTIVE_CONTENT_BLOCKED");
    assert.equal(findings.length, 1);
    assert.equal(findings[0].fidelity, "not-rendered");
    assert.equal(findings[0].phase, "security");
    assert.match(findings[0].part, /vbaProject\.bin/u);
  } finally {
    document.close();
    engine.close();
  }
});

test("never follows an external officeDocument relationship", async () => {
  const external = createZip({
    "[Content_Types].xml": "<Types/>",
    "_rels/.rels": `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" TargetMode="External" Target="https://example.invalid/document.xml"/>
      </Relationships>`,
  });
  const engine = await inlineEngine();
  try {
    await assert.rejects(
      engine.open(external),
      (error) => error?.code === "EXTERNAL_RESOURCE_BLOCKED",
    );
  } finally {
    engine.close();
  }
});
