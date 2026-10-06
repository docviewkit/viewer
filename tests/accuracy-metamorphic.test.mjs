import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";
import {
  createMetamorphicPackages,
  readZipComment,
  readZipEntries,
} from "../scripts/accuracy-metamorphic.mjs";
import { createZip } from "./zip-fixture.mjs";

const decoder = new TextDecoder();

test("metamorphic package generator applies four independent semantics-preserving changes", () => {
  const input = createZip({
    "[Content_Types].xml": "<Types xmlns=\"urn:types\"><Override ContentType=\"application/xml\" PartName=\"/word/document.xml\"/></Types>",
    "_rels/.rels": "<Relationships xmlns=\"urn:relationships\"><Relationship Target=\"word/document.xml\" Type=\"officeDocument\" Id=\"rId1\"/></Relationships>",
    "word/document.xml": "<w:document xmlns:w=\"urn:word\" w:version=\"1\"><w:p w:id=\"7\">Hello</w:p></w:document>",
  });
  const originalNames = readZipEntries(input).map(({ name }) => name);
  const variants = createMetamorphicPackages(input);

  assert.deepEqual(variants.map(({ kind }) => kind), [
    "xml-attribute-order",
    "namespace-prefix",
    "zip-entry-order",
    "irrelevant-metadata",
  ]);
  for (const variant of variants) assert.notDeepEqual(variant.bytes, input);

  const attributeXml = decoder.decode(readZipEntries(variants[0].bytes)[0].data);
  assert.match(attributeXml, /PartName="\/word\/document\.xml" ContentType="application\/xml"/u);

  const namespaceXml = decoder.decode(readZipEntries(variants[1].bytes).find(({ name }) => name === "word/document.xml").data);
  assert.match(namespaceXml, /<ovm:document xmlns:ovm="urn:word"/u);
  assert.doesNotMatch(namespaceXml, /<w:document/u);

  assert.deepEqual(readZipEntries(variants[2].bytes).map(({ name }) => name), [...originalNames].reverse());
  assert.deepEqual(readZipEntries(variants[3].bytes).map(({ name }) => name), originalNames);
  assert.equal(readZipComment(variants[3].bytes), "OfficeViewer accuracy harness metadata mutation");
});

test("public SDK structure, layout, IDs, and hit mapping survive all four package transformations", async () => {
  const docx = createZip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>`,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="0A1B2C3D"><w:pPr><w:spacing w:line="360"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Hello DOCX</w:t></w:r></w:p>
          <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
        </w:body>
      </w:document>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const observe = async (bytes) => {
    const document = await engine.open(bytes);
    try {
      return {
        info: document.info,
        diagnostics: document.diagnostics(),
        hits: await document.hitTest({ unitIndex: 0, x: 100, y: 100 }),
      };
    } finally {
      document.close();
    }
  };
  try {
    const baseline = await observe(docx);
    for (const variant of createMetamorphicPackages(docx)) {
      assert.deepEqual(await observe(variant.bytes), baseline, variant.kind);
    }
  } finally {
    engine.close();
  }
});
