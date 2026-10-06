import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";
import { createZip } from "./zip-fixture.mjs";

const workbook = createZip({
  "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Search" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Budget</t></is></c><c r="B1" t="inlineStr"><is><t>预算</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>budget budgeted budget</t></is></c></row></sheetData></worksheet>`,
});

const batchedPdfText = new TextEncoder().encode(`%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >> endobj
4 0 obj << >> stream
BT /F1 10 Tf 1 0 0 1 10 80 Tm [(A) 0 (A) 0 (A)] TJ ET
endstream endobj
5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Test /FirstChar 65 /LastChar 65 /Widths [700] /FontDescriptor 6 0 R >> endobj
6 0 obj << /Type /FontDescriptor /FontName /Test /Ascent 750 /Descent -250 >> endobj
trailer << /Root 1 0 R >>
%%EOF`);

test("searchText finds immutable text ranges with case, whole-word, unit, and result limits", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(workbook);
  try {
    const matches = await document.searchText({
      query: "budget",
      unitIndex: 0,
      wholeWord: true,
      limit: 2,
    });
    assert.deepEqual(matches.map(({ object, ranges }) => ({
      text: object.text,
      ranges,
    })), [
      { text: "Budget", ranges: [[0, 6]] },
      { text: "budget budgeted budget", ranges: [[0, 6], [16, 22]] },
    ]);
    assert.equal(Object.isFrozen(matches), true);
    assert.equal(Object.isFrozen(matches[0].ranges), true);

    const cjk = await document.searchText({ query: "预算", matchCase: true });
    assert.deepEqual(cjk.map(({ object }) => object.text), ["预算"]);
  } finally {
    document.close();
    engine.close();
  }
});

test("listObjects exposes an immutable, bounded unit view for viewer text layers", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(workbook);
  try {
    const textObjects = await document.listObjects({
      unitIndex: 0,
      types: ["cell"],
      textOnly: true,
      limit: 2,
    });
    assert.deepEqual(textObjects.map((object) => object.text), ["Budget", "预算"]);
    assert.equal(Object.isFrozen(textObjects), true);
    assert.equal(Object.isFrozen(textObjects[0]), true);
    await assert.rejects(
      document.listObjects({ unitIndex: 1 }),
      (error) => error?.code === "INVALID_UNIT",
    );
  } finally {
    document.close();
    engine.close();
  }
});

test("PDF text batches expose one copyable object while retaining hit testing", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(batchedPdfText);
  try {
    const objects = await document.listObjects({ unitIndex: 0, types: ["text-box"], textOnly: true });
    assert.equal(objects.length, 1);
    assert.equal(objects[0].text, "AAA");
    assert.equal(objects[0].source.mapping, "derived");
    assert.deepEqual(
      objects[0].fontRuns?.map(({ start, end }) => ({ start, end })),
      [{ start: 0, end: 1 }, { start: 1, end: 2 }, { start: 2, end: 3 }],
    );
    const { x, y, width, height } = objects[0].bounds;
    const hits = await document.hitTest({ unitIndex: 0, x: x + width * 0.8, y: y + height / 2 });
    assert.equal(hits[0]?.object.id, objects[0].id);
    assert.equal(hits[0]?.object.text, "AAA");
  } finally {
    document.close();
    engine.close();
  }
});
