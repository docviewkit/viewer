import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";
import { extendedFormatPack } from "../dist/extended-formats.js";
import { sheetPrintPages } from "../dist/render.js";
import { createZip } from "./zip-fixture.mjs";
import { readZipEntries } from "../scripts/accuracy-metamorphic.mjs";

function workbook(worksheet, styles, extraParts = {}) {
  return createZip({
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>${styles === undefined ? "" : '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'}</Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>${styles === undefined ? "" : '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'}</Relationships>`,
    "xl/worksheets/sheet1.xml": worksheet,
    ...(styles === undefined ? {} : { "xl/styles.xml": styles }),
    ...extraParts,
  });
}

function spreadsheet(content, extraParts = {}) {
  return createZip({
    mimetype: "application/vnd.oasis.opendocument.spreadsheet",
    "META-INF/manifest.xml": `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`,
    "content.xml": content,
    ...extraParts,
  }, { compress: false });
}

async function open(bytes, engineOptions = {}) {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({
    wasm,
    execution: "inline",
    fontPolicy: "local-first",
    formatPack: async () => ({
      async load(candidate) {
        return readFile(await extendedFormatPack.load(candidate));
      },
    }),
    ...engineOptions,
  });
  return { engine, document: await engine.open(bytes) };
}

test("XLSX dialog sheets do not prevent remaining sheets from opening", async () => {
  const { engine, document } = await open(await readFile(new URL("./fixtures/Dialogsheet.xlsx", import.meta.url)));
  try {
    assert.deepEqual(document.info.units.map(unit => unit.name), ["Sheet1", "Sheet2", "Sheet3"]);
    assert.ok(document.diagnostics().some(diagnostic => diagnostic.code === "UNSUPPORTED_FEATURE"
      && diagnostic.message.includes("Dialog1")));
  } finally {
    document.close();
    engine.close();
  }
});

test("ODS supplied minimum decimal places use numeric formats rather than cached separators", async () => {
  const { engine, document } = await open(await readFile(new URL(
    "./fixtures/oasis-3695-min-decimal-placed.ods", import.meta.url,
  )));
  try {
    assert.deepEqual((await document.listObjects()).map(cell => cell.text),
      ["7.89", "7.9", "7.890", "7.89 ", "7.89  "]);
  } finally {
    document.close();
    engine.close();
  }
});

test("ODS decimal formats preserve explicit locales and bound unsupported formatting", async () => {
  const original = await readFile(new URL("./fixtures/oasis-3695-min-decimal-placed.ods", import.meta.url));
  for (const [attributes, expected] of [
    ['number:language="de" number:country="DE"', "7,890"],
    ['number:rfc-language-tag="de-DE"', "7,890"],
    ["", "7.890"],
  ]) {
    const parts = Object.fromEntries(readZipEntries(original).map(({ name, data }) => [name, data]));
    parts["styles.xml"] = new TextDecoder().decode(parts["styles.xml"])
      .replace('style:name="N108"', `style:name="N108" ${attributes}`);
    const { engine, document } = await open(createZip(parts));
    try { assert.equal((await document.listObjects())[2].text, expected); }
    finally { document.close(); engine.close(); }
  }
  for (const [format, expected, value = "7.89"] of [
    ['number:decimal-places="3" number:min-decimal-places="1"', "7.89"],
    ['number:decimal-places="3" number:min-decimal-places="3"', "7.890"],
    ['number:decimal-places="3" number:min-decimal-places="0" number:decimal-replacement=" "', "7.89 "],
    ['number:decimal-places="2" number:grouping="true" number:min-integer-digits="6"', "001,234.50", "1234.5"],
    ['number:decimal-places="2" number:min-integer-digits="0"', ".25", "0.25"],
    ['number:decimal-places="999999999"', "cached"],
    ['number:decimal-places="2" number:min-decimal-places="5"', "cached"],
    ['number:decimal-places="2" number:display-factor="100"', "cached"],
  ]) {
    const { engine, document } = await open(spreadsheet(`<office:document-content>
      <office:automatic-styles><number:number-style style:name="n"><number:number ${format}/></number:number-style>
      <style:style style:family="table-cell" style:name="c" style:data-style-name="n"/></office:automatic-styles>
      <office:body><office:spreadsheet><table:table table:name="Test"><table:table-row>
      <table:table-cell table:style-name="c" office:value-type="float" office:value="${value}"><text:p>cached</text:p></table:table-cell>
      </table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`));
    try {
      assert.equal((await document.listObjects())[0].text, expected);
      if (expected === "cached") assert.ok(document.diagnostics().some(d => /numeric format preserves cached/.test(d.message)));
    }
    finally { document.close(); engine.close(); }
  }
});

function installRecordingCanvas() {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const imageBitmapDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const pathDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Path2D");
  const calls = { fills: [], fillRects: [], paintOrder: [], strokes: [], paintedStrokes: [], fonts: [], text: [], rects: [], lines: [], beziers: [], transforms: [], rotations: [], drawImages: [], filters: [], shadowBlurs: [], clips: 0, clippedText: [] };
  const savedClipDepths = [];
  let clipDepth = 0;
  let strokeStyle;
  let fillStyle;
  let lineWidth;
  const context = {
    save() { savedClipDepths.push(clipDepth); },
    restore() { clipDepth = savedClipDepths.pop() ?? 0; },
    getTransform() { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; }, setTransform() {},
    scale() {}, translate() {}, rotate(...args) { calls.rotations.push(args); }, transform(...args) { calls.transforms.push(args); },
    drawImage(...args) { calls.drawImages.push(args); },
    getImageData(x, y, width, height) { return { width, height, data: new Uint8ClampedArray(width * height * 4) }; },
    putImageData() {},
    createLinearGradient() { return { addColorStop() {} }; },
    createPattern() { return { setTransform() {} }; },
    fillRect(...args) { calls.fillRects.push([fillStyle, ...args]); }, beginPath() {}, closePath() {}, rect(...args) { calls.rects.push(args); }, ellipse() {}, moveTo() {}, lineTo(...args) { calls.lines.push(args); }, bezierCurveTo(...args) { calls.beziers.push(args); },
    fill() { calls.paintOrder.push(fillStyle); }, stroke() { calls.paintedStrokes.push([strokeStyle, lineWidth]); }, clip() { calls.clips += 1; clipDepth += 1; }, measureText(value) { return { width: value.length * 8 }; },
    fillText(...args) { calls.text.push(args); if (clipDepth > 0) calls.clippedText.push(args[0]); },
    set fillStyle(value) { fillStyle = value; calls.fills.push(value); },
    set strokeStyle(value) { strokeStyle = value; calls.strokes.push(value); },
    set font(value) { calls.fonts.push(value); },
    set filter(value) { calls.filters.push(value); },
    set shadowBlur(value) { calls.shadowBlurs.push(value); },
    set lineWidth(value) { lineWidth = value; }, set textBaseline(_value) {}, set textAlign(_value) {},
    globalAlpha: 1, globalCompositeOperation: "source-over", imageSmoothingEnabled: true, imageSmoothingQuality: "high",
  };
  class RecordingCanvas {
    constructor(width = 0, height = 0) { this.width = width; this.height = height; }
    getContext() { context.canvas = this; return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  // Arbitrary picture clips now use the same DrawingML paths as presentation pictures.
  Object.defineProperty(globalThis, "Path2D", { configurable: true, value: class {
    moveTo() {} lineTo() {} closePath() {} quadraticCurveTo() {}
    bezierCurveTo(...args) { calls.beziers.push(args); }
  } });
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: RecordingCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async (source) => ({ width: source.width ?? 1, height: source.height ?? 1, close() {} }),
  });
  return {
    calls,
    restore() {
      if (descriptor) Object.defineProperty(globalThis, "OffscreenCanvas", descriptor);
      else delete globalThis.OffscreenCanvas;
      if (pathDescriptor) Object.defineProperty(globalThis, "Path2D", pathDescriptor);
      else delete globalThis.Path2D;
      if (imageBitmapDescriptor) Object.defineProperty(globalThis, "createImageBitmap", imageBitmapDescriptor);
      else delete globalThis.createImageBitmap;
    },
  };
}

test("renders the original Numbers chart through the public protocol", async () => {
  const { engine, document } = await open(await readFile(new URL("./fixtures/chart-original.numbers", import.meta.url)));
  const canvas = installRecordingCanvas();
  let frame;
  try {
    const objects = await document.listObjects();
    const charts = objects.filter((object) => object.source.kind === "chart");
    assert.equal(charts.length, 1);
    assert.equal(charts[0].source.format, "numbers");
    assert.equal(charts[0].source.part, "Index/CalculationEngine-905377.iwa");
    assert.ok(charts[0].bounds.x < 170 && charts[0].bounds.width > 450,
      "chart bounds must include labels outside the authored plot frame");
    assert.equal(objects.filter((object) => object.type === "cell").length, 4);
    const unit = document.info.units[0];
    frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: unit.width, height: unit.height }, scale: 1 });
    assert.ok(canvas.calls.paintOrder.some((color) => color !== "#000000" && color !== "#ffffff"));
    assert.equal(canvas.calls.paintOrder.filter((color) => color === "rgba(91, 155, 213, 1)").length, 2,
      "both columns use the inherited native series color");
    for (const value of ["10", "58"]) {
      assert.equal(canvas.calls.text.filter(([text]) => text === value).length, 1, `${value} belongs only to its cell; native iWork hides chart value labels`);
    }
    for (const tick of ["15", "30", "45", "60"]) {
      assert.equal(canvas.calls.text.filter(([text]) => text === tick).length, 1, `axis tick ${tick} must stay on one line`);
    }
    assert.equal(canvas.calls.drawImages.length, 0, "chart must be native shapes, not a preview image");
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("derives XLSX dimensions and renders inline strings, booleans, and cached formulas", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t>Inline </t></r><r><t>text</t></r></is></c><c r="B1" t="b"><v>1</v></c><c r="C1"><f>1+1</f><v>2</v></c><c r="D1"><f>NOW()</f></c></row></sheetData></worksheet>`);
  const { engine, document } = await open(bytes);
  try {
    assert.deepEqual(document.info.units, [
      { type: "sheet", index: 0, id: "unit:0", name: "Data", width: 244, height: 24, rows: 1, columns: 4, frozenRows: 0, frozenColumns: 0, frozenWidth: 0, frozenHeight: 0, rowAxis: { defaultSize: 24, spans: [] }, columnAxis: { defaultSize: 61, spans: [] }, showGridLines: true },
    ]);
    const expected = ["Inline text", "true", "2", undefined];
    for (let column = 0; column < expected.length; column += 1) {
      const hit = await document.hitTest({ unitIndex: 0, x: column * 61 + 2, y: 2 });
      assert.equal(hit[0]?.object.text, expected[column]);
      if (column >= 2) assert.equal(hit[0]?.object.source.formula, column === 2 ? "=1+1" : "=NOW()");
    }
    assert.equal(
      document.diagnostics().some((diagnostic) => diagnostic.code === "UNSUPPORTED_FEATURE" && /cached/u.test(diagnostic.message)),
      true,
    );
  } finally {
    document.close();
    engine.close();
  }
});

test("loads the optional calculator only for XLSX formulas without cached results", async () => {
  const cached = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>`);
  const cachedOpen = await open(cached, { calculationWasm: false });
  try {
    assert.equal((await cachedOpen.document.hitTest({ unitIndex: 0, x: 2, y: 2 }))[0]?.object.text, "2");
    assert.equal(cachedOpen.document.diagnostics().some(({ message }) => /optional calculation module/u.test(message)), false);
  } finally {
    cachedOpen.document.close();
    cachedOpen.engine.close();
  }

  const uncached = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>1+1</f></c></row></sheetData></worksheet>`);
  const calculationWasm = await readFile(new URL("../dist/office-viewer-calc.wasm", import.meta.url));
  const calculatedOpen = await open(uncached, { calculationWasm });
  try {
    assert.equal((await calculatedOpen.document.hitTest({ unitIndex: 0, x: 2, y: 2 }))[0]?.object.text, "2");
    assert.equal(calculatedOpen.document.diagnostics().some(({ message }) => /optional calculation module/u.test(message)), false);
  } finally {
    calculatedOpen.document.close();
    calculatedOpen.engine.close();
  }
});

test("calculator runtime ABI preserves dependent formulas, text and error results", async () => {
  const formulas = [
    ["SUM(2,3,4)", "9"],
    ["A1*2", "18"],
    ['IF(A2&gt;10,"ok","bad")', "ok"],
    ["ROUND(1.2345,2)", "1.23"],
    ["MAX(4,2)", "4"],
    ["IFERROR(1/0,7)", "7"],
    ["1/0", "#DIV/0!"],
    ['TEXT(12.5,"0.00")', "12.50"],
  ];
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${formulas.map(([formula], index) =>
    `<row r="${index + 1}"><c r="A${index + 1}"><f>${formula}</f></c></row>`).join("")}</sheetData></worksheet>`);
  const calculationWasm = await readFile(new URL("../dist/office-viewer-calc.wasm", import.meta.url));
  const { engine, document } = await open(bytes, { calculationWasm });
  try {
    for (const [index, [, expected]] of formulas.entries()) {
      const hits = await document.hitTest({ unitIndex: 0, x: 2, y: index * 24 + 2 });
      assert.equal(hits[0]?.object.text, expected);
    }
  } finally {
    document.close();
    engine.close();
  }
});

test("recalculates cached TODAY formulas when the workbook opens", async () => {
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><f>TODAY()</f><v>42189</v></c></row></sheetData></worksheet>`,
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>`,
  );
  const calculationWasm = await readFile(new URL("../dist/office-viewer-calc.wasm", import.meta.url));
  const before = new Date();
  const { engine, document } = await open(bytes, { calculationWasm });
  try {
    const after = new Date();
    const dates = [before, after].map((date) => `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`);
    assert.ok(dates.includes((await document.hitTest({ unitIndex: 0, x: 2, y: 2 }))[0]?.object.text));
  } finally {
    document.close();
    engine.close();
  }
});

test("supplied NoExtDataA1 refreshes every shared NOW cell and preserves other caches", async () => {
  const bytes = await readFile(new URL("./fixtures/NoExtDataA1.xlsx", import.meta.url));
  const calculationWasm = await readFile(new URL("../dist/office-viewer-calc.wasm", import.meta.url));
  const before = new Date();
  const { engine, document } = await open(bytes, { calculationWasm });
  try {
    const after = new Date();
    const times = [before, after].map(date => `${date.getMonth() + 1}/${date.getDate()}/${String(date.getFullYear()).slice(-2)} ${date.getHours()}:${String(date.getMinutes()).padStart(2, "0")}`);
    const objects = await document.listObjects({ unitIndex: 0, limit: 40_000 });
    const cells = new Map(objects.filter(cell => cell.text !== undefined).map(cell => [cell.source.address, cell.text]));
    assert.ok(times.includes(cells.get("B2")), cells.get("B2"));
    for (let row = 2; row <= 1250; row++) {
      for (const column of ["B", "C", "H", "I"]) {
        assert.equal(cells.get(`${column}${row}`), cells.get(`${column}2`), `${column}${row}`);
      }
    }
    assert.equal(cells.get("D3"), "0.0524322");
    assert.equal(cells.get("J3"), "35902.274");
    assert.equal(cells.get("F3"), "3.71");
    assert.deepEqual(document.diagnostics(), []);
  } finally {
    document.close();
    engine.close();
  }
});

test("opens an empty XLSX worksheet without a dimension element", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`);
  const { engine, document } = await open(bytes);
  try {
    assert.deepEqual(document.info.units, [
      { type: "sheet", index: 0, id: "unit:0", name: "Data", width: 61, height: 24, rows: 1, columns: 1, frozenRows: 0, frozenColumns: 0, frozenWidth: 0, frozenHeight: 0, rowAxis: { defaultSize: 24, spans: [] }, columnAxis: { defaultSize: 61, spans: [] }, showGridLines: true },
    ]);
    assert.deepEqual(await document.hitTest({ unitIndex: 0, x: 2, y: 2 }), []);
  } finally {
    document.close();
    engine.close();
  }
});

test("keeps the public Strict XLSX sheet extent beyond text overflowing its last used column", async () => {
  const bytes = workbook(`<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><dimension ref="I13"/><sheetData><row r="13"><c r="I13" s="0" t="inlineStr"><is><t>image with  a style applied</t></is></c></row></sheetData></worksheet>`);
  const { engine, document } = await open(bytes);
  try {
    const unit = document.info.units[0];
    const [cell] = await document.listObjects({ unitIndex: 0, textOnly: true });
    assert.ok(unit.width > cell.bounds.x + cell.bounds.width, JSON.stringify({ unit, cell }));
  } finally {
    document.close();
    engine.close();
  }
});

test("renders the supplied empty XLSX chart title as Office Chart Title", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/chart-empty-title-original.xlsx", import.meta.url)));
    const objects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "xl/charts/chart1.xml");
    assert.equal(objects.filter(({ text }) => text === "Chart Title").length, 1);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders the supplied scatter chart standard error bars in every OOXML host", async (t) => {
  const original = await readFile(new URL("./fixtures/testErrorBarProp.xlsx", import.meta.url));
  const parts = new Map(readZipEntries(original).map(({ name, data }) => [name, data]));
  const chart = parts.get("xl/charts/chart1.xml");
  const rels = (body) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  const rel = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
  const types = (part, type) => `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.${type}.main+xml"/></Types>`;
  const graphic = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="chart"/></a:graphicData></a:graphic>`;
  const ns = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
  const inputs = [
    ["xlsx", original],
    ["docx", createZip({
      "[Content_Types].xml": types("word/document.xml", "wordprocessingml.document"),
      "_rels/.rels": rels(rel("root", "officeDocument", "word/document.xml")),
      "word/document.xml": `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ${ns}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Error bars"/>${graphic}</wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`,
      "word/_rels/document.xml.rels": rels(rel("chart", "chart", "charts/chart1.xml") + rel("theme", "theme", "theme/theme1.xml")),
      "word/charts/chart1.xml": chart,
      "word/theme/theme1.xml": parts.get("xl/theme/theme1.xml"),
    })],
    ["pptx", createZip({
      "[Content_Types].xml": types("ppt/presentation.xml", "presentationml.presentation"),
      "_rels/.rels": rels(rel("root", "officeDocument", "ppt/presentation.xml")),
      "ppt/presentation.xml": `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${ns}><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
      "ppt/_rels/presentation.xml.rels": rels(rel("slide", "slide", "slides/slide1.xml") + rel("theme", "theme", "theme/theme1.xml")),
      "ppt/slides/slide1.xml": `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${ns}><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="1" name="Error bars"/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="5486400" cy="3200400"/></p:xfrm>${graphic}</p:graphicFrame></p:spTree></p:cSld></p:sld>`,
      "ppt/slides/_rels/slide1.xml.rels": rels(rel("chart", "chart", "../charts/chart1.xml")),
      "ppt/charts/chart1.xml": chart,
      "ppt/theme/theme1.xml": parts.get("xl/theme/theme1.xml"),
    })],
  ];
  for (const [host, bytes] of inputs) await t.test(host, async () => {
    const { engine, document } = await open(bytes);
    const canvas = installRecordingCanvas();
    let frame;
    try {
      const unit = document.info.units[0];
      frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: unit.width, height: unit.height }, scale: 1 });
      assert.equal(canvas.calls.paintedStrokes.filter(([color]) => color === "rgba(255, 0, 0, 1)").length, 10,
        "ten vertical standard error bars must reach Canvas");
      assert.equal(canvas.calls.paintedStrokes.filter(([color, width]) => color === "rgba(89, 89, 89, 1)" && width === 1).length, 10,
        "ten horizontal standard error bars must retain their gray one-pixel stroke");
      assert.ok(canvas.calls.strokes.includes("rgba(68, 114, 196, 1)"), "error-bar styles must not overwrite the blue series");
      const labels = canvas.calls.text.map(([text]) => text);
      assert.ok(labels.includes("1918") && labels.includes("1932"), "numeric X axis surrounds the horizontal errors");
      assert.ok(labels.includes("88") && labels.includes("91.5"), "Y axis and points share their error-aware scale");
    } finally {
      frame?.bitmap.close(); document.close(); engine.close(); canvas.restore();
    }
  });
});

test("renders an XLSX chart legend at its Office-defined position with markers", async () => {
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:J20"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
      "xl/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:ser><c:tx><c:v>Left Legend</c:v></c:tx><c:marker><c:symbol val="square"/><c:size val="5"/></c:marker><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea><c:legend><c:legendPos val="l"/></c:legend></c:chart></c:chartSpace>`,
    },
  );
  const { engine, document } = await open(bytes);
  try {
    const objects = await document.listObjects({ unitIndex: 0 });
    const chart = objects.find((object) => object.type === "group");
    const legend = objects.find((object) => object.text === "Left Legend");
    assert.ok(chart && legend, JSON.stringify(objects));
    assert.ok(legend.bounds.x < chart.bounds.x + chart.bounds.width / 2, JSON.stringify({ chart: chart.bounds, legend: legend.bounds }));
    assert.ok(objects.some(({ text, bounds }) => text === undefined
      && bounds.x < legend.bounds.x
      && bounds.width > 2 && bounds.width < 20
      && bounds.height > 2 && bounds.height < 20
      && Math.abs(bounds.y + bounds.height / 2 - (legend.bounds.y + legend.bounds.height / 2)) < 2), JSON.stringify(objects));
  } finally {
    document.close();
    engine.close();
  }
});

test("renders an XLSX chartsheet through its drawing relationship", async () => {
  const bytes = createZip({
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Chart" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>`,
    "xl/chartsheets/sheet1.xml": `<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><tabColor rgb="FF123456"/></sheetPr><drawing r:id="rIdDrawing"/></chartsheet>`,
    "xl/chartsheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="6457950" cy="4391025"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:absoluteAnchor></xdr:wsDr>`,
    "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
    "xl/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
  });
  const { engine, document } = await open(bytes);
  try {
    assert.equal(document.info.units[0].tabColor, 0x123456ff);
    const objects = await document.listObjects({ unitIndex: 0 });
    assert.ok(objects.some(({ type }) => type === "group"), JSON.stringify(objects));
  } finally {
    document.close();
    engine.close();
  }
});

test("renders built-in XLSX table fills including Spreadsheet.xlsx TableStyleMedium2", async () => {
  for (const tableStyle of ["TableStyleMedium2", "TableStyleMedium9"]) {
    const bytes = workbook(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`,
      undefined,
      {
        "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`,
        "xl/tables/table1.xml": `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ref="A1:A3"><tableStyleInfo name="${tableStyle}" showRowStripes="1"/></table>`,
      },
    );
    const canvas = installRecordingCanvas();
    const { engine, document } = await open(bytes);
    try {
      const frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: 64, height: 60 }, scale: 1, pixelRatio: 1 });
      frame.bitmap.close();
      assert.ok(canvas.calls.fills.includes("rgba(91, 155, 213, 1)"), `missing ${tableStyle} header fill: ${JSON.stringify(canvas.calls.fills)}`);
      assert.ok(canvas.calls.fills.includes("rgba(222, 235, 247, 1)"), `missing ${tableStyle} row stripe: ${JSON.stringify(canvas.calls.fills)}`);
    } finally {
      canvas.restore();
      document.close();
      engine.close();
    }
  }
});

test("renders custom whole-table fills from basicspreadsheet.xlsx", async () => {
  const bytes = await readFile(new URL("./fixtures/basicspreadsheet-custom-styles.xlsx", import.meta.url));
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  try {
    const frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: 900, height: 900 }, scale: 1, pixelRatio: 1 });
    frame.bitmap.close();
    assert.equal(canvas.calls.paintOrder.filter((fill) => fill === "rgba(247, 150, 70, 1)").length, 36);
  } finally {
    canvas.restore();
    document.close();
    engine.close();
  }
});

test("renders cached XLSX pivot table fills across omitted blank cells", async () => {
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D4"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Sales</t></is></c><c r="B1" t="inlineStr"><is><t>Column Labels</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Row Labels</t></is></c><c r="B2" t="inlineStr"><is><t>Canada</t></is></c><c r="C2" t="inlineStr"><is><t>France</t></is></c><c r="D2" t="inlineStr"><is><t>Germany</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>CY 2003</t></is></c><c r="B3"><v>42</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>Grand Total</t></is></c><c r="B4"><v>42</v></c></row></sheetData></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>`,
      "xl/pivotTables/pivotTable1.xml": `<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><location ref="A1:D4" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><rowFields count="1"><field x="0"/></rowFields><rowItems count="2"><i><x/></i><i t="grand"><x/></i></rowItems><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showLastColumn="1"/></pivotTableDefinition>`,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  try {
    const frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: 244, height: 96 }, scale: 1, pixelRatio: 1 });
    frame.bitmap.close();
    assert.equal(canvas.calls.paintOrder.filter((fill) => fill === "rgba(222, 235, 247, 1)").length, 12);
    assert.equal((await document.hitTest({ unitIndex: 0, x: 3 * 61 + 2, y: 2 }))[0]?.object.source.address, "D1");
  } finally {
    canvas.restore();
    document.close();
    engine.close();
  }
});

test("clips centered XLSX text before an occupied preceding cell", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><name val="Arial"/></font></fonts><fills><fill><patternFill patternType="none"/></fill></fills><borders><border/></borders><cellXfs><xf fontId="0" fillId="0" borderId="0"/><xf fontId="0" fillId="0" borderId="0"><alignment horizontal="center"/></xf></cellXfs></styleSheet>`;
  const text = "qqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqq";
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>${text}</t></is></c><c r="C1"/></row></sheetData></worksheet>`, styles);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  try {
    const frame = await document.render({
      unitIndex: 0,
      viewport: { x: 0, y: 0, width: 288, height: 24 },
      scale: 1,
      pixelRatio: 1,
    });
    frame.bitmap.close();
    assert.ok(
      canvas.calls.clippedText.some((value) => value.includes("qqqq")),
      `centered text must be clipped before the occupied preceding cell: ${JSON.stringify({ text: canvas.calls.text, clippedText: canvas.calls.clippedText })}`,
    );
  } finally {
    canvas.restore();
    document.close();
    engine.close();
  }
});

test("expands a lazily loaded XLSX viewport after another sheet without duplicate object IDs", async () => {
  const bytes = createZip({
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>first</t></is></c></row></sheetData></worksheet>`,
    "xl/worksheets/sheet2.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A40"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>top</t></is></c></row><row r="40"><c r="A40" t="inlineStr"><is><t>bottom</t></is></c></row></sheetData></worksheet>`,
  });
  const { engine, document } = await open(bytes);
  try {
    await document.listObjects({ unitIndex: 0, viewport: { x: 0, y: 0, width: 96, height: 24 } });
    await document.listObjects({ unitIndex: 1, viewport: { x: 0, y: 0, width: 96, height: 24 } });
    const expanded = await document.listObjects({
      unitIndex: 1,
      viewport: { x: 0, y: 936, width: 96, height: 48 },
    });
    assert.deepEqual(expanded.map((object) => object.text), ["bottom"]);
  } finally {
    document.close();
    engine.close();
  }
});

test("exposes authored XLSX print settings through the public protocol", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView view="pageLayout"/></sheetViews><sheetData/>
    <pageMargins left="0.7" right="0.7" top="0.8" bottom="0.8" header="0.3" footer="0.3"/>
    <pageSetup paperSize="9" orientation="portrait" scale="85" fitToWidth="1" fitToHeight="2"/>
    <headerFooter differentFirst="1"><oddHeader>&amp;CCBAM report</oddHeader><firstFooter>&amp;P</firstFooter></headerFooter>
  </worksheet>`);
  const { engine, document } = await open(bytes);
  try {
    const settings = document.info.units[0].printSettings;
    const { margins, ...withoutMargins } = settings;
    assert.deepEqual(withoutMargins, {
      viewMode: "pageLayout",
      paperSize: 9,
      orientation: "portrait",
      scale: 85,
      fitToWidth: 1,
      fitToHeight: 2,
      fitToPage: true,
      differentOddEven: false,
      differentFirst: true,
      oddHeader: "&CCBAM report",
      firstFooter: "&P",
    });
    for (const [name, expected] of Object.entries({ left: 0.7, right: 0.7, top: 0.8, bottom: 0.8, header: 0.3, footer: 0.3 })) {
      assert.ok(Math.abs(margins[name] - expected) < 1e-6, `${name} margin`);
    }
    assert.equal(Object.isFrozen(settings), true);
    assert.equal(Object.isFrozen(margins), true);
  } finally {
    document.close();
    engine.close();
  }
});

test("retains saved XLSX row heights when the rows are not marked custom", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs>
  </styleSheet>`;
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:H43"/><sheetFormatPr defaultRowHeight="15"/>
    <cols><col min="1" max="1" width="12" customWidth="1"/><col min="3" max="3" width="15.25" customWidth="1"/><col min="4" max="4" width="18.375" customWidth="1"/><col min="5" max="5" width="14.5" customWidth="1"/><col min="6" max="6" width="17.25" customWidth="1"/><col min="7" max="8" width="12" customWidth="1"/></cols>
    <sheetData><row r="37" ht="60"/><row r="41" ht="30.75" customHeight="1"/><row r="43"><c r="H43"/></row></sheetData>
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="portrait"/>
  </worksheet>`, styles);
  const { engine, document } = await open(bytes);
  try {
    const unit = document.info.units[0];
    assert.equal(unit.rowAxis.defaultSize, 20);
    assert.deepEqual(unit.rowAxis.spans, [
      { start: 36, end: 36, size: 80 },
      { start: 40, end: 40, size: 41 },
    ]);
    await document.listObjects({ unitIndex: 0 });
    assert.deepEqual(document.info.units[0].rowAxis, unit.rowAxis);
    // Saved heights total 941 px, requiring two vertical pages as well as two horizontal pages.
    assert.equal(unit.height, 941);
    assert.equal(sheetPrintPages(unit).length, 4);
  } finally {
    document.close();
    engine.close();
  }
});

test("WNC Architecture pictures retain their saved row-anchor proportions", async () => {
  // Architecture sheet, styles, drawing and image bytes extracted unchanged from
  // WNC_Digital_Key_Requirements_and_Scenario.xlsx (SHA-256
  // 46ceadd2e6686d24e4d1f974fbce4854545a30bcdde9e9c5750e8fd97c7a6dd7).
  const { engine, document } = await open(await readFile(new URL(
    "./fixtures/wnc-picture-geometry.xlsx", import.meta.url,
  )));
  try {
    assert.equal(document.info.units[0].rowAxis.defaultSize, 20);
    const images = (await document.listObjects({ unitIndex: 0 }))
      .filter(object => object.type === "image");
    assert.equal(images.length, 2);
    for (const [index, height] of [
      21 * 20 + (26248 - 106680) / 9525,
      30 * 20 + (117841 - 144780) / 9525,
    ].entries()) {
      assert.ok(Math.abs(images[index].bounds.height - height) < 0.001);
      assert.ok(Math.abs(images[index].bounds.width - (18 * 69 + 250814 / 9525)) < 0.001);
    }
  } finally {
    document.close();
    engine.close();
  }
});

test("renders view-only row heights and column widths without changing document metadata", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders>
    <cellXfs count="2"><xf/><xf borderId="1"/></cellXfs>
  </styleSheet>`;
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>A</t></is></c><c r="B1" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>`, styles);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({
      unitIndex: 0,
      sheetSizes: { rows: [{ index: 0, size: 30 }], columns: [{ index: 0, size: 100 }] },
    });
    assert.deepEqual(frame.viewport, { x: 0, y: 0, width: 169, height: 30 });
    assert.equal(canvas.calls.rects.some((rect) => rect[0] === 0 && rect[1] === 0 && rect[2] === 100 && rect[3] === 30), true);
    assert.equal(canvas.calls.rects.some((rect) => rect[0] === 100 && rect[1] === 0 && rect[2] === 1 && rect[3] === 30), true, JSON.stringify(canvas.calls.rects));
    assert.equal(canvas.calls.rects.some((rect) => rect[0] === 0 && rect[1] === 29 && rect[2] === 100 && rect[3] === 1), true, JSON.stringify(canvas.calls.rects));
    assert.equal(document.info.units[0].width, 138);
    assert.equal(document.info.units[0].height, 24);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("paints XLSX cell borders after adjacent cell fills", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill></fills>
    <borders count="7"><border/><border><left style="medium"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="medium"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border><border><left style="thin"><color rgb="FF000000"/></left><right style="medium"><color rgb="FF000000"/></right><top style="medium"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border><border><left style="medium"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border><border><left style="thin"><color rgb="FF000000"/></left><right style="medium"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border><border><left style="medium"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="medium"><color rgb="FF000000"/></bottom></border><border><left style="thin"><color rgb="FF000000"/></left><right style="medium"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="medium"><color rgb="FF000000"/></bottom></border></borders>
    <cellXfs count="7"><xf/><xf fillId="1" borderId="1"/><xf fillId="1" borderId="2"/><xf fillId="1" borderId="3"/><xf fillId="1" borderId="4"/><xf fillId="1" borderId="5"/><xf fillId="1" borderId="6"/></cellXfs>
  </styleSheet>`;
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B5"/><sheetData><row r="1" ht="18"><c r="A1" s="1"/><c r="B1" s="2"/></row><row r="2" ht="71"><c r="A2" s="3"/><c r="B2" s="4"/></row><row r="3" ht="71"><c r="A3" s="3"/><c r="B3" s="4"/></row><row r="4" ht="53"><c r="A4" s="3"/><c r="B4" s="4"/></row><row r="5" ht="53.75"><c r="A5" s="5"/><c r="B5" s="6"/></row></sheetData></worksheet>`, styles);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({
      unitIndex: 0,
      scale: 1,
      pixelRatio: 1,
      sheetSizes: { rows: [], columns: [] },
    });
    const lastCellFill = canvas.calls.paintOrder.lastIndexOf("rgba(255, 255, 255, 1)");
    const firstBorder = canvas.calls.paintOrder.indexOf("rgba(0, 0, 0, 1)");
    assert.ok(firstBorder > lastCellFill, JSON.stringify(canvas.calls.paintOrder));
    const verticalBorders = canvas.calls.rects.filter(([, , width]) => width === 1);
    assert.equal(verticalBorders.filter(([x]) => x === 69).length, 10, JSON.stringify(verticalBorders));
    assert.equal(verticalBorders.some(([x]) => x === 68 || x === 70), false, JSON.stringify(verticalBorders));
    const horizontalBorders = canvas.calls.rects.filter(([, , width, height]) => width === 69 && height === 1);
    const horizontalCoordinates = new Set(horizontalBorders.map(([, y]) => y.toFixed(3)));
    assert.equal(horizontalCoordinates.size, 4, JSON.stringify(horizontalBorders));
    assert.equal(canvas.calls.rects.some(([x, y, width]) => x === 0 && y === 0 && width === 2), true, JSON.stringify(canvas.calls.rects));
    assert.equal(canvas.calls.rects.some(([x, , width]) => x === 136 && width === 2), true, JSON.stringify(canvas.calls.rects));
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("keeps XLSX diagonal borders inside their cell", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="2"><border/><border diagonalUp="1" diagonalDown="1"><left style="thick"/><right style="thick"/><top style="thick"/><bottom style="thick"/><diagonal style="thick"/></border></borders>
    <cellXfs count="2"><xf borderId="0"/><xf borderId="1"/></cellXfs>
  </styleSheet>`;
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="B1:C1"/><sheetData><row r="1"><c r="B1" s="1"/><c r="C1"/></row></sheetData></worksheet>`, styles);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({
      unitIndex: 0,
      scale: 1,
      pixelRatio: 1,
      sheetSizes: { rows: [], columns: [] },
    });
    assert.deepEqual(canvas.calls.lines.slice(-2), [[138, 24], [138, 0]]);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("uses XLSX row and column geometry for merged-cell bounds and hit testing", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:D3"/>
    <cols>
      <col min="1" max="1" width="10" customWidth="1"/>
      <col min="2" max="2" hidden="1"/>
      <col min="3" max="4" width="12" customWidth="1"/>
    </cols>
    <sheetData>
      <row r="1" ht="30" customHeight="1">
        <c r="A1" t="inlineStr"><is><t>Visible</t></is></c>
        <c r="C1" t="inlineStr"><is><t>Merged</t></is></c>
        <c r="D1" t="inlineStr"><is><t>covered</t></is></c>
      </row>
      <row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>hidden row</t></is></c></row>
      <row r="3"><c r="A3" t="inlineStr"><is><t>Last</t></is></c><c r="C3" t="inlineStr"><is><t>covered too</t></is></c></row>
    </sheetData>
    <mergeCells count="1"><mergeCell ref="C1:D3"/></mergeCells>
  </worksheet>`);
  const { engine, document } = await open(bytes);
  try {
    assert.deepEqual(document.info.units, [
      { type: "sheet", index: 0, id: "unit:0", name: "Data", width: 238, height: 64, rows: 3, columns: 4, frozenRows: 0, frozenColumns: 0, frozenWidth: 0, frozenHeight: 0, rowAxis: { defaultSize: 24, spans: [{ start: 0, end: 0, size: 40 }, { start: 1, end: 1, size: 0 }] }, columnAxis: { defaultSize: 61, spans: [{ start: 0, end: 0, size: 70 }, { start: 1, end: 1, size: 0 }, { start: 2, end: 3, size: 84 }] }, showGridLines: true },
    ]);

    const visible = await document.hitTest({ unitIndex: 0, x: 10, y: 10 });
    assert.equal(visible[0]?.object.text, "Visible");
    assert.deepEqual(visible[0]?.object.bounds, { x: 0, y: 0, width: 70, height: 40 });

    const merged = await document.hitTest({ unitIndex: 0, x: 230, y: 50 });
    assert.equal(merged.length, 1);
    assert.equal(merged[0]?.object.text, "Merged");
    assert.deepEqual(merged[0]?.object.bounds, { x: 70, y: 0, width: 168, height: 64 });

    const last = await document.hitTest({ unitIndex: 0, x: 10, y: 50 });
    assert.deepEqual(last.map(({ object }) => object.text), ["Last"]);
    const hiddenRowBoundary = await document.hitTest({ unitIndex: 0, x: 10, y: 40 });
    assert.equal(hiddenRowBoundary.some(({ object }) => object.text === "hidden row"), false);
    assert.equal(document.diagnostics().some((diagnostic) => /merged cells use independent/u.test(diagnostic.message)), false);
  } finally {
    document.close();
    engine.close();
  }
});

test("uses the Normal style maximum digit width for Office-authored XLSX columns", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:C1"/><sheetFormatPr defaultColWidth="9.1640625"/>
    <cols><col min="2" max="2" width="5.5" customWidth="1"/><col min="3" max="3" width="28.5" customWidth="1"/></cols>
    <sheetData><row r="1"><c r="A1"/><c r="B1"/><c r="C1"/></row></sheetData>
  </worksheet>`, `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts><font><sz val="11"/><name val="宋体"/></font></fonts><fills><fill/></fills><borders><border/></borders>
    <cellXfs><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
  </styleSheet>`);
  const { engine, document } = await open(bytes);
  try {
    assert.deepEqual(document.info.units[0].columnAxis, {
      defaultSize: 73,
      spans: [{ start: 1, end: 1, size: 44 }, { start: 2, end: 2, size: 228 }],
    });
    assert.equal(document.info.units[0].width, 345);
  } finally {
    document.close();
    engine.close();
  }
});

test("keeps an entirely hidden XLSX sheet renderable without exposing hidden cells", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1"/><cols><col min="1" max="1" hidden="1"/></cols>
    <sheetData><row r="1" hidden="1"><c r="A1" t="inlineStr"><is><t>hidden</t></is></c></row></sheetData>
  </worksheet>`);
  const { engine, document } = await open(bytes);
  try {
    assert.deepEqual(document.info.units, [
      { type: "sheet", index: 0, id: "unit:0", name: "Data", width: 1, height: 1, rows: 1, columns: 1, frozenRows: 0, frozenColumns: 0, frozenWidth: 0, frozenHeight: 0, rowAxis: { defaultSize: 24, spans: [{ start: 0, end: 0, size: 0 }] }, columnAxis: { defaultSize: 61, spans: [{ start: 0, end: 0, size: 0 }] }, showGridLines: true },
    ]);
    assert.deepEqual(await document.hitTest({ unitIndex: 0, x: 0, y: 0 }), []);
  } finally {
    document.close();
    engine.close();
  }
});

test("exposes XLSX frozen pane counts and real pixel extents", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:C3"/>
    <sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>
    <cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="12" customWidth="1"/></cols>
    <sheetData><row r="1" ht="30" customHeight="1"/></sheetData>
  </worksheet>`);
  const { engine, document } = await open(bytes);
  try {
    assert.deepEqual(document.info.units, [{
      type: "sheet", index: 0, id: "unit:0", name: "Data", width: 215, height: 88,
      rows: 3, columns: 3, frozenRows: 1, frozenColumns: 2, frozenWidth: 154, frozenHeight: 40,
      rowAxis: { defaultSize: 24, spans: [{ start: 0, end: 0, size: 40 }] },
      columnAxis: { defaultSize: 61, spans: [{ start: 0, end: 0, size: 70 }, { start: 1, end: 1, size: 84 }] },
      showGridLines: true,
    }]);
  } finally {
    document.close();
    engine.close();
  }
});

test("applies XLSX cell font, fill, border, alignment, and number format when rendering", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts>
    <fonts count="2"><font/><font><b/><i/><sz val="14"/><color rgb="FF336699"/><name val="Aptos"/></font></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></fills>
    <borders count="2"><border/><border><left style="thin"><color rgb="FFFF0000"/></left></border></borders>
    <cellXfs count="2"><xf/><xf numFmtId="164" fontId="1" fillId="1" borderId="1"><alignment horizontal="center" wrapText="1"/></xf></cellXfs>
  </styleSheet>`;
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" s="1"><v>0.125</v></c></row></sheetData></worksheet>`,
    styles,
  );
  const canvas = installRecordingCanvas();
  const { calls } = canvas;

  const { engine, document } = await open(bytes);
  let frame;
  try {
    const hit = await document.hitTest({ unitIndex: 0, x: 10, y: 10 });
    assert.equal(hit[0]?.object.text, "12.50%");
    assert.equal(hit[0]?.object.wrapText, true);
    assert.ok(Math.abs(hit[0].object.fontRuns[0].fontSize - 14 * 96 / 72) < 0.001);
    assert.equal(hit[0].object.fontRuns[0].bold, true);
    assert.equal(hit[0].object.fontRuns[0].italic, true);
    frame = await document.render({ unitIndex: 0 });
    assert.equal(calls.fills.includes("rgba(255, 204, 0, 1)"), true);
    assert.equal(calls.fills.includes("rgba(255, 0, 0, 1)"), true);
    assert.equal(calls.fonts.some((font) => /Aptos/u.test(font) && /italic/u.test(font) && /700/u.test(font)), true);
    assert.equal(calls.text.some(([text]) => text === "12.50%"), true);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders General XLSX decimals at Excel's visible precision", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">
    <c r="A1"><v>83.116531372070313</v></c><c r="B1"><v>1.4164407253265381</v></c><c r="C1"><v>0.60000002384185791</v></c>
  </row></sheetData></worksheet>`);
  const { engine, document } = await open(bytes);
  try {
    assert.deepEqual((await document.listObjects({ unitIndex: 0, textOnly: true })).map(({ text }) => text), ["83.11653", "1.416441", "0.6"]);
  } finally {
    document.close();
    engine.close();
  }
});

test("renders the complete XLSX indexed color palette", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="2"><font/><font><color indexed="16"/></font></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor indexed="22"/></patternFill></fill></fills>
    <borders count="1"><border/></borders><cellXfs count="2"><xf/><xf fontId="1" fillId="1"/></cellXfs>
  </styleSheet>`;
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Indexed</t></is></c></row></sheetData></worksheet>`, styles);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.includes("rgba(192, 192, 192, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(128, 0, 0, 1)"), true);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("does not paint worksheet gridlines through filled XLSX cells", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill></fills>
    <borders count="1"><border/></borders><cellXfs count="2"><xf/><xf fillId="1"/></cellXfs>
  </styleSheet>`;
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/></row></sheetData></worksheet>`, styles);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.paintedStrokes.some(([color]) => color === "rgba(208, 208, 208, 1)"), false);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders XLSX worksheet gridlines for empty cells in the visible viewport", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetViews><sheetView showGridLines="1"/></sheetViews><sheetData/></worksheet>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    assert.equal(document.info.units[0].showGridLines, true);
    frame = await document.render({
      unitIndex: 0,
      viewport: { x: 61, y: 24, width: 61, height: 24 },
      sheetSizes: { rows: [], columns: [] },
    });
    const grid = canvas.calls.fillRects.filter(([fill]) => fill === "rgba(208, 208, 208, 1)");
    assert.deepEqual(grid, [
      ["rgba(208, 208, 208, 1)", 61, 24, 1, 24],
      ["rgba(208, 208, 208, 1)", 121, 24, 1, 24],
      ["rgba(208, 208, 208, 1)", 61, 24, 61, 1],
      ["rgba(208, 208, 208, 1)", 61, 47, 61, 1],
    ]);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("honors disabled XLSX worksheet gridlines", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetViews><sheetView showGridLines="0"/></sheetViews><sheetData/></worksheet>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    assert.equal(document.info.units[0].showGridLines, false);
    frame = await document.render({ unitIndex: 0, sheetSizes: { rows: [], columns: [] } });
    assert.equal(canvas.calls.fillRects.some(([fill]) => fill === "rgba(208, 208, 208, 1)"), false);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders locale-qualified XLSX currency symbols", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="1"><numFmt numFmtId="164" formatCode="[$$-409]#,##0.00"/></numFmts>
    <fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="1"><border/></borders><cellXfs count="2"><xf/><xf numFmtId="164"/></cellXfs>
  </styleSheet>`;
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>3033784.21</v></c></row></sheetData></worksheet>`, styles);
  const { engine, document } = await open(bytes);
  try {
    const objects = await document.listObjects({ unitIndex: 0 });
    assert.equal(objects[0]?.text, "$3,033,784.21");
  } finally {
    document.close();
    engine.close();
  }
});

test("resolves XLSX theme fills and SpreadsheetML tints", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts>
    <fills count="5">
      <fill><patternFill patternType="none"/></fill>
      <fill><patternFill patternType="solid"><fgColor theme="0"/></patternFill></fill>
      <fill><patternFill patternType="solid"><fgColor theme="1"/></patternFill></fill>
      <fill><patternFill patternType="solid"><fgColor theme="4"/></patternFill></fill>
      <fill><patternFill patternType="solid"><fgColor theme="4" tint="0.5"/></patternFill></fill>
    </fills>
    <borders count="1"><border/></borders>
    <cellXfs count="5"><xf/><xf fillId="1"/><xf fillId="2"/><xf fillId="3"/><xf fillId="4"/></cellXfs>
  </styleSheet>`;
  const theme = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom">
    <a:themeElements><a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="112233"/></a:dk1>
      <a:lt1><a:srgbClr val="445566"/></a:lt1>
      <a:dk2><a:srgbClr val="222222"/></a:dk2>
      <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
      <a:accent2><a:srgbClr val="00FF00"/></a:accent2>
      <a:accent3><a:srgbClr val="0000FF"/></a:accent3>
      <a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
      <a:accent5><a:srgbClr val="FF00FF"/></a:accent5>
      <a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme></a:themeElements>
  </a:theme>`;
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D1"/><sheetData><row r="1"><c r="A1" s="1"/><c r="B1" s="2"/><c r="C1" s="3"/><c r="D1" s="4"/></row></sheetData></worksheet>`,
    styles,
    {
      "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`,
      "xl/theme/theme1.xml": theme,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.includes("rgba(68, 85, 102, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(17, 34, 51, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(255, 0, 0, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(255, 128, 128, 1)"), true);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("shared ChartML unit and multilevel labels reach every OOXML host", async () => {
  for (const [source, expected] of [
    ["ooxml-display-units.docx", ["Billions", "2E-9"]],
    ["ooxml-multilevel-axis.docx", ["Categoria 1", "Categoria 2", "Categoria 3", "Categoria 4", "2011", "2012"]],
  ]) {
    const original = await readFile(new URL(`./fixtures/${source}`, import.meta.url));
    const chart = readZipEntries(original).find(entry => entry.name === "word/charts/chart1.xml").data;
    for (const carrier of [undefined, "corpus-chart-wall.pptx", "chart-empty-title-original.xlsx"]) {
      let bytes = original;
      if (carrier !== undefined) {
        // Adapter boundary: transplant the real authored ChartML cache into an existing host.
        const parts = Object.fromEntries(readZipEntries(await readFile(new URL(`./fixtures/${carrier}`, import.meta.url)))
          .map(({ name, data }) => [name, data]));
        const part = Object.keys(parts).find(name => /^(ppt|xl)\/charts\/chart1\.xml$/u.test(name));
        assert.ok(part);
        parts[part] = chart;
        bytes = createZip(parts);
      }
      const { engine, document } = await open(bytes, { calculationWasm: false });
      try {
        const text = (await document.listObjects()).map(object => object.text ?? "").join("|");
        for (const label of expected) assert.ok(text.includes(label), `${carrier ?? source}: missing ${label}: ${text}`);
      } finally { document.close(); engine.close(); }
    }
  }
});

test("real OOXML text conditions paint through the optional calculator", async () => {
  const bytes = await readFile(new URL("./fixtures/ooxml-conditional-priority.xlsx", import.meta.url));
  const calculationWasm = await readFile(new URL("../dist/office-viewer-calc.wasm", import.meta.url));
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes, { calculationWasm });
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.ok(canvas.calls.fills.includes("rgba(255, 0, 0, 1)"), JSON.stringify(canvas.calls.fills));
    // The x14 rules in this original file remain explicitly unsupported.
    assert.ok(document.diagnostics().some(d => /conditional-format rules are not rendered/u.test(d.message)));
  } finally {
    frame?.bitmap.close(); document.close(); engine.close(); canvas.restore();
  }
});

test("conditional styles honor global priority and stopIfTrue across sqref blocks", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData>
    <conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>
    <conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan" stopIfTrue="1"><formula>5</formula></cfRule></conditionalFormatting>
    <conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="2" priority="3" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>
    </worksheet>`, `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs><dxfs count="3">
    <dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf>
    <dxf><font><b/></font><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></dxf>
    <dxf><font><i/></font><fill><patternFill patternType="solid"><fgColor rgb="FF0000FF"/></patternFill></fill></dxf>
    </dxfs></styleSheet>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes, { calculationWasm: false });
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.ok(canvas.calls.fills.includes("rgba(255, 0, 0, 1)"));
    assert.ok(!canvas.calls.fills.includes("rgba(0, 255, 0, 1)"));
    assert.ok(!canvas.calls.fills.includes("rgba(0, 0, 255, 1)"));
    assert.ok(canvas.calls.fonts.some(font => /bold|700/u.test(font)), JSON.stringify(canvas.calls.fonts));
    assert.ok(canvas.calls.fonts.every(font => !font.includes("italic")));
    assert.ok(!document.diagnostics().some(d => /optional calculation/u.test(d.message)));
  } finally {
    frame?.bitmap.close(); document.close(); engine.close(); canvas.restore();
  }
});

test("unsupported conditional date systems degrade only the rule", async () => {
  const calculationWasm = await readFile(new URL("../dist/office-viewer-calc.wasm", import.meta.url));
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData>
    <conditionalFormatting sqref="A1"><cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="today"><formula>TRUE()</formula></cfRule></conditionalFormatting>
    </worksheet>`, `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs>
    <dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>`);
  for (const date1904 of [false, true]) {
    const parts = Object.fromEntries(readZipEntries(bytes).map(({ name, data }) => [name, data]));
    if (date1904) parts["xl/workbook.xml"] = new TextDecoder().decode(parts["xl/workbook.xml"]).replace('<sheets>', '<workbookPr date1904="1"/><sheets>');
    const { engine, document } = await open(createZip(parts), { calculationWasm });
    try {
      assert.equal((await document.listObjects()).find(object => object.text === "10")?.text, "10");
      assert.equal(document.diagnostics().some(d => /conditional-format rules are not rendered/u.test(d.message)), date1904);
    } finally { document.close(); engine.close(); }
  }
});

test("renders XLSX cellIs conditional formatting from differential styles", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs>
    <dxfs count="1"><dxf><font><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>
  </styleSheet>`;
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/>
      <sheetData><row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>3</v></c></row></sheetData>
      <conditionalFormatting sqref="A1:B1"><cfRule type="cellIs" dxfId="0" operator="greaterThan" priority="1"><formula>5</formula></cfRule></conditionalFormatting>
    </worksheet>`,
    styles,
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.filter((fill) => fill === "rgba(255, 0, 0, 1)").length, 1);
    assert.equal(canvas.calls.fills.includes("rgba(255, 255, 255, 1)"), true);
    assert.equal(
      document.diagnostics().some((diagnostic) => /conditional formatting is not rendered/u.test(diagnostic.message)),
      false,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("FilterByColor renders symbols instead of circle text", async () => {
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(await readFile(new URL("./fixtures/FilterByColor.xlsx", import.meta.url)));
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.text.some(([text]) => text === "●"), false);
    for (const color of ["rgba(99, 163, 106, 1)", "rgba(240, 174, 76, 1)", "rgba(211, 106, 103, 1)"]) {
      assert.ok(canvas.calls.fills.includes(color), `${color} symbol must be a filled path`);
    }
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("XLSX icon spacing respects alignment, hidden values and stopped rules", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>${[1, 2, 3, 4].map(row => `<row r="${row}"><c r="A${row}"><v>${row}</v></c></row>`).join('')}</sheetData>
    <conditionalFormatting sqref="A3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan" stopIfTrue="1"><formula>0</formula></cfRule></conditionalFormatting>
    <conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="2"><iconSet><cfvo type="num" val="0"/><cfvo type="num" val="2"/><cfvo type="num" val="3"/></iconSet></cfRule></conditionalFormatting>
    <conditionalFormatting sqref="A4"><cfRule type="iconSet" priority="3"><iconSet showValue="0"><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting>
    </worksheet>`, `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts><font/></fonts><fills><fill><patternFill patternType="none"/></fill></fills>
    <borders><border/></borders><cellXfs><xf><alignment horizontal="left"/></xf></cellXfs><dxfs><dxf/></dxfs></styleSheet>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    const objects = await document.listObjects();
    const icons = objects.filter(object => object.type === "shape");
    assert.deepEqual(icons.map(object => object.source.address), ["A1", "A2", "A4"]);
    for (const value of ["1", "2"]) {
      const icon = icons.find(object => object.source.address === `A${value}`);
      const text = canvas.calls.text.find(([text]) => text === value);
      assert.ok(text[1] >= icon.bounds.x + icon.bounds.width + 1, `${value} must leave room for the icon`);
    }
    assert.equal(canvas.calls.text.find(([text]) => text === "3")[1], 2, "a stopped icon rule must not indent text");
    assert.equal(canvas.calls.text.some(([text]) => text === "4" || text === "●"), false);
  } finally {
    frame?.bitmap.close(); document.close(); engine.close(); canvas.restore();
  }
});

test("renders XLSX color scales, data bars, and icon sets as visible objects", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/>
    <sheetData>
      <row r="1"><c r="A1"><v>0</v></c><c r="B1"><v>50</v></c><c r="C1"><v>100</v></c></row>
      <row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>50</v></c><c r="C2"><v>100</v></c></row>
      <row r="3"><c r="A3"><v>1</v></c><c r="B3"><v>2</v></c><c r="C3"><v>3</v></c></row>
    </sheetData>
    <conditionalFormatting sqref="A1:C1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/><color rgb="FFFF0000"/></colorScale></cfRule></conditionalFormatting>
    <conditionalFormatting sqref="A2:C2"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF00AA44"/></dataBar></cfRule></conditionalFormatting>
    <conditionalFormatting sqref="A3:C3"><cfRule type="iconSet" priority="3"><iconSet iconSet="3Arrows"><cfvo type="min"/><cfvo type="num" val="2"/><cfvo type="num" val="3"/></iconSet></cfRule></conditionalFormatting>
  </worksheet>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    for (const color of [
      "rgba(0, 0, 255, 1)",
      "rgba(255, 0, 0, 1)",
      "rgba(0, 170, 68, 0.6)",
    ]) {
      assert.equal(canvas.calls.fills.includes(color), true, `${color} must be painted`);
    }
    assert.equal(canvas.calls.text.some(([text]) => ["↓", "→", "↑"].includes(text)), false);
    for (const color of ["rgba(255, 80, 80, 1)", "rgba(255, 192, 0, 1)", "rgba(112, 173, 71, 1)"]) {
      assert.ok(canvas.calls.fills.includes(color), `${color} arrow must be painted as a path`);
    }
    assert.equal(
      document.diagnostics().some((diagnostic) => /conditional-format rules are not rendered/u.test(diagnostic.message)),
      false,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders case-insensitive text conditional fills and omits read-only validation markers", async () => {
  const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs>
    <dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFC00000"/></patternFill></fill></dxf></dxfs>
  </styleSheet>`;
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/>
    <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Open</t></is></c></row></sheetData>
    <conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>"OPEN"</formula></cfRule></conditionalFormatting>
    <dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"Open,Closed"</formula1></dataValidation></dataValidations>
  </worksheet>`, styles);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.includes("rgba(192, 0, 0, 1)"), true);
    const objects = await document.listObjects({ unitIndex: 0 });
    assert.deepEqual(objects.map(({ type, text }) => ({ type, text })), [{ type: "cell", text: "Open" }]);
    assert.equal(document.diagnostics().some((diagnostic) => /conditional-format rules are not rendered/u.test(diagnostic.message)), false);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("resolves XLSX drawing relationships and two-cell image anchors", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="7" name="Logo"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm rot="1800000"><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
      "xl/media/image1.png": png,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    assert.deepEqual(
      { rows: document.info.units[0].rows, columns: document.info.units[0].columns, width: document.info.units[0].width, height: document.info.units[0].height },
      { rows: 2, columns: 2, width: 122, height: 48 },
    );
    const hits = await document.hitTest({ unitIndex: 0, x: 61, y: 24 });
    assert.equal(hits[0]?.object.type, "image");
    assert.deepEqual(hits[0]?.object.bounds, { x: 0, y: 0, width: 122, height: 48 });
    assert.equal(
      document.diagnostics().some((diagnostic) => /drawing layer/u.test(diagnostic.message)),
      false,
    );
    frame = await document.render({ unitIndex: 0 });
    assert.equal(
      canvas.calls.transforms.some(([a, b, c, d]) => Math.abs(a - Math.cos(Math.PI / 6)) < 0.001
        && Math.abs(b - Math.sin(Math.PI / 6)) < 0.001
        && Math.abs(c + Math.sin(Math.PI / 6)) < 0.001
        && Math.abs(d - Math.cos(Math.PI / 6)) < 0.001),
      true,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders authored XLSX picture reflection and glow effects", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:row>2</xdr:row></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="7"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/></xdr:blipFill><xdr:spPr><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8594"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"><a:shade val="85000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst><a:glow rad="19050"><a:srgbClr val="FF0000"/></a:glow><a:reflection stA="50000" endA="0" endPos="50000" dist="9525"/></a:effectLst></xdr:spPr></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
      "xl/media/image1.png": png,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    assert.ok(document.info.units[0].height >= 99, `reflection clipped by sheet extent: ${document.info.units[0].height}`);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(canvas.calls.drawImages.length >= 2, `missing reflected image composition: ${canvas.calls.drawImages.length}`);
    assert.ok(canvas.calls.shadowBlurs.filter((blur) => blur === 1).length >= 2, `missing source/reflection glow: ${JSON.stringify(canvas.calls.shadowBlurs)}`);
    assert.ok(canvas.calls.paintOrder.includes("rgba(237, 237, 237, 1)"), `missing authored picture fill: ${JSON.stringify(canvas.calls.paintOrder)}`);
    assert.ok(canvas.calls.beziers.some((curve) => Math.abs(curve[4] - 4.125) < 0.01 && Math.abs(curve[5]) < 0.01), `wrong authored roundRect radius: ${JSON.stringify(canvas.calls.beziers[0])}`);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders embedded images as XLSX worksheet shape fills", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:row>2</xdr:row></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="ellipse"/><a:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></xdr:spPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
      "xl/media/image1.png": png,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.drawImages.length, 1);
    assert.equal((await document.listObjects({ unitIndex: 0 }))[0]?.type, "shape");
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("clips XLSX picture fills to shared DrawingML preset geometry", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="irregularSeal2"/><a:blipFill><a:blip r:embed="rIdImage"/></a:blipFill></xdr:spPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
      "xl/media/image1.png": png,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.ok(canvas.calls.lines.length > 20, `picture fill used fallback rectangle geometry: ${JSON.stringify(canvas.calls.lines)}`);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("maps an XLSX group picture fill through each grpFill child", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:row>4</xdr:row></xdr:to><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="1"/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm><a:blipFill><a:blip r:embed="rIdImage"/></a:blipFill></xdr:grpSpPr><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="snip1Rect"/><a:grpFill/></xdr:spPr></xdr:sp><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="plus"/><a:grpFill/></xdr:spPr></xdr:sp></xdr:grpSp></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
      "xl/media/image1.png": png,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.drawImages.length, 2);
    assert.ok(canvas.calls.drawImages[1][1] > canvas.calls.drawImages[0][1], JSON.stringify(canvas.calls.drawImages));
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders an XLSX ActiveX control's VML appearance without activating it", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rIdDrawing"/><legacyDrawing r:id="rIdVml"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1025" name="Object 1"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/vmlDrawing1.vml": `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="CommandButton1" o:spid="_x0000_s1025"><v:imagedata o:relid="rIdPreview"/></v:shape></xml>`,
      "xl/drawings/_rels/vmlDrawing1.vml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/preview.png"/></Relationships>`,
      "xl/media/preview.png": png,
    },
  );
  const { engine, document } = await open(bytes);
  try {
    const objects = await document.listObjects({ unitIndex: 0 });
    assert.equal(objects.length, 1);
    assert.equal(objects[0]?.type, "image");
    assert.deepEqual(objects[0]?.bounds, { x: 0, y: 0, width: 122, height: 48 });
  } finally {
    document.close();
    engine.close();
  }
});

test("applies authored rotation to XLSX worksheet shapes", async () => {
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="7"/></xdr:nvSpPr><xdr:spPr><a:xfrm rot="1800000"><a:off x="0" y="0"/><a:ext cx="1219200" cy="381000"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Rotated</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(
      canvas.calls.transforms.some(([a, b, c, d]) => Math.abs(a - Math.cos(Math.PI / 6)) < 0.001
        && Math.abs(b - Math.sin(Math.PI / 6)) < 0.001
        && Math.abs(c + Math.sin(Math.PI / 6)) < 0.001
        && Math.abs(d - Math.cos(Math.PI / 6)) < 0.001),
      true,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders XLSX SmartArt data nodes with an approximation diagnostic", async () => {
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="7"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdDiagram"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDiagram" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>`,
      "xl/diagrams/data1.xml": `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>Node</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>`,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal((await document.listObjects({ unitIndex: 0 })).some(({ text }) => text === "Node"), true);
    assert.match(JSON.stringify(document.diagnostics()), /XLSX SmartArt used a deterministic data-model layout/u);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders XLSX SmartArt picture nodes and background fills", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:row>2</xdr:row></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="7"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><dgm:relIds r:dm="rIdDiagram"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDiagram" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>`,
      "xl/diagrams/data1.xml": `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dgm:bg><a:blipFill><a:blip r:embed="rIdBackground"/></a:blipFill></dgm:bg><dgm:ptLst><dgm:pt modelId="node"><dgm:prSet phldr="1"/></dgm:pt><dgm:pt modelId="presentation" type="pres"><dgm:prSet presAssocID="node" presName="imagNode"/><dgm:spPr><a:blipFill><a:blip r:embed="rIdNode"/></a:blipFill></dgm:spPr></dgm:pt></dgm:ptLst></dgm:dataModel>`,
      "xl/diagrams/_rels/data1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBackground" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/background.png"/><Relationship Id="rIdNode" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/node.png"/></Relationships>`,
      "xl/media/background.png": png,
      "xl/media/node.png": png,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.drawImages.length, 2);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders XLSX bar, line, and pie charts from cached series data", async () => {
  const anchor = (id, relationship, fromColumn, fromRow, toColumn, toRow) => `<xdr:twoCellAnchor><xdr:from><xdr:col>${fromColumn}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${toColumn}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="${relationship}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
  const chart = (kind, values, style = "") => `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${style}<c:chart><c:plotArea><c:${kind}Chart><c:ser>${kind === "bar" ? '<c:dLbls><c:dLbl><c:idx val="2"/><c:tx><c:rich><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Custom</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:ln><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:ln></c:spPr></c:dLbl></c:dLbls>' : ""}${kind === "line" ? '<c:marker><c:symbol val="square"/><c:size val="8"/></c:marker><c:cat><c:numLit><c:formatCode>mmm\\ yyyy</c:formatCode><c:pt idx="0"><c:v>41521</c:v></c:pt><c:pt idx="1"><c:v>41551</c:v></c:pt><c:pt idx="2"><c:v>41582</c:v></c:pt></c:numLit></c:cat>' : ""}<c:val><c:numRef><c:numCache>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:val></c:ser></c:${kind}Chart>${kind === "line" ? '<c:dateAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Months</a:t></a:r></a:p></c:rich></c:tx></c:title><c:txPr><a:bodyPr rot="-5400000"/></c:txPr></c:dateAx><c:valAx><c:axPos val="l"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Usage</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>' : ""}</c:plotArea></c:chart></c:chartSpace>`;
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:F12"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchor(11, "rIdBar", 0, 0, 2, 4)}${anchor(12, "rIdLine", 2, 0, 4, 4)}${anchor(13, "rIdPie", 4, 0, 6, 4)}</xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBar" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rIdLine" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/><Relationship Id="rIdPie" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/></Relationships>`,
      "xl/charts/chart1.xml": chart("bar", [10, 30, 20], '<c:style val="47"/>'),
      "xl/charts/chart2.xml": chart("line", [5, 25, 15]),
      "xl/charts/chart3.xml": chart("pie", [1, 2, 3]),
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    const lineObjects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "xl/charts/chart2.xml");
    assert.equal(canvas.calls.fills.includes("rgba(63, 63, 63, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(255, 0, 0, 1)"), true);
    assert.equal(canvas.calls.strokes.includes("rgba(255, 255, 0, 1)"), true);
    assert.equal(canvas.calls.text.map(([text]) => text).join("").includes("Custom"), true);
    assert.equal(canvas.calls.fills.includes("rgba(68, 114, 196, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(237, 125, 49, 1)"), true);
    assert.equal(canvas.calls.strokes.includes("rgba(68, 114, 196, 1)"), true);
    assert.ok(lineObjects.some(({ text }) => text === "Months"));
    assert.ok(lineObjects.some(({ text }) => text === "Usage"));
    assert.equal(lineObjects.filter(({ bounds }) => bounds.width > 10 && bounds.width < 11 && bounds.height > 10 && bounds.height < 11).length, 3);
    assert.ok(canvas.calls.rotations.some(([rotation]) => Math.abs(rotation + Math.PI / 2) < 0.001));
    assert.equal(
      document.diagnostics().some((diagnostic) => /chart drawing is not rendered/u.test(diagnostic.message)),
      false,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders XLSX pie3D charts with shared curved side walls", async () => {
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:row>8</xdr:row></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
      "xl/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:view3D><c:rotX val="30"/><c:perspective val="30"/></c:view3D><c:plotArea><c:pie3DChart><c:varyColors val="1"/><c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:dLbls><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls></c:pie3DChart></c:plotArea></c:chart></c:chartSpace>`,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.ok(canvas.calls.paintOrder.length >= 6, `pie3D side walls were omitted: ${JSON.stringify(canvas.calls.paintOrder)}`);
    assert.ok(canvas.calls.beziers.length >= 28, `pie3D arcs were not curved: ${canvas.calls.beziers.length}`);
    const pieShapes = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ type, source }) => type === "shape" && source.part === "xl/charts/chart1.xml");
    const pieWidth = Math.max(...pieShapes.map(({ bounds }) => bounds.width));
    const pieHeight = Math.max(...pieShapes.map(({ bounds }) => bounds.height));
    assert.ok(pieWidth > 140 && pieHeight < pieWidth * 0.75, JSON.stringify({ pieWidth, pieHeight }));
    assert.ok(
      canvas.calls.text.some(([text]) => text.includes("A")) && canvas.calls.text.some(([text]) => text.includes("33%")),
      `pie labels were omitted: ${JSON.stringify(canvas.calls.text)}`,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders every cached series in an XLSX mixed chart with centered area span", async () => {
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:F12"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:row>12</xdr:row></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
      "xl/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:areaChart><c:ser><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:areaChart><c:barChart><c:barDir val="col"/><c:ser><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart><c:lineChart><c:ser><c:spPr><a:ln><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart><c:valAx><c:crossBetween val="between"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.includes("rgba(255, 0, 0, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(0, 255, 0, 1)"), true);
    assert.equal(canvas.calls.strokes.includes("rgba(0, 0, 255, 1)"), true);
    assert.ok(
      canvas.calls.paintOrder.indexOf("rgba(0, 255, 0, 1)") < canvas.calls.paintOrder.indexOf("rgba(255, 0, 0, 1)"),
      `mixed chart layers ignored OOXML order: ${JSON.stringify(canvas.calls.paintOrder)}`,
    );
    const objects = await document.listObjects({ unitIndex: 0 });
    const chart = objects.find(({ type }) => type === "group");
    const centeredArea = objects.find(({ text, bounds }) => text === undefined
      && bounds.width > chart.bounds.width * 0.35 && bounds.width < chart.bounds.width * 0.65
      && bounds.height > chart.bounds.height * 0.6);
    assert.ok(centeredArea, "cross-between area must span only the category centers");
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders XLSX chart series with workbook theme colors", async () => {
  const bytes = workbook(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C6"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`,
    undefined,
    {
      "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`,
      "xl/theme/theme1.xml": `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4F81BD"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`,
      "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:row>6</xdr:row></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>`,
      "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
      "xl/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
    },
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.includes("rgba(79, 129, 189, 1)"), true);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("keeps XLSX data validations non-visual in the read-only viewer", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:G1"/><sheetData><row r="1">
      <c r="A1" t="inlineStr"><is><t>One</t></is></c><c r="B1"><v>3</v></c><c r="C1"><v>1.5</v></c>
      <c r="D1"><v>46000</v></c><c r="E1"><v>0.5</v></c><c r="F1" t="inlineStr"><is><t>short</t></is></c><c r="G1"><v>7</v></c>
    </row></sheetData>
    <dataValidations count="7">
      <dataValidation type="list" allowBlank="1" showInputMessage="1" promptTitle="Choice" prompt="Choose an allowed value" showErrorMessage="1" errorTitle="Invalid" error="Use the list" sqref="A1"><formula1>&quot;One,Two,Three&quot;</formula1></dataValidation>
      <dataValidation type="whole" operator="between" showErrorMessage="1" errorTitle="Range" error="Enter 1 through 5" sqref="B1"><formula1>1</formula1><formula2>5</formula2></dataValidation>
      <dataValidation type="decimal" sqref="C1"><formula1>0</formula1><formula2>2</formula2></dataValidation>
      <dataValidation type="date" sqref="D1"><formula1>DATE(2025,1,1)</formula1></dataValidation>
      <dataValidation type="time" sqref="E1"><formula1>TIME(8,0,0)</formula1></dataValidation>
      <dataValidation type="textLength" sqref="F1"><formula1>10</formula1></dataValidation>
      <dataValidation type="custom" sqref="G1"><formula1>ISNUMBER(G1)</formula1></dataValidation>
    </dataValidations>
  </worksheet>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    const listHits = await document.hitTest({ unitIndex: 0, x: 58, y: 10 });
    const customHits = await document.hitTest({ unitIndex: 0, x: 6 * 61 + 30, y: 10 });
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.text.some(([text]) => ["▼", "ⓘ", "!"].includes(text)), false);
    assert.equal(listHits.some(({ object }) => object.source.address === "A1" && object.source.mapping === "exact"), true);
    assert.equal(customHits.some(({ object }) => object.source.address === "G1" && object.source.mapping === "exact"), true);
    assert.equal((await document.listObjects({ unitIndex: 0 })).every(({ type }) => type === "cell"), true);
    assert.equal(document.diagnostics().some((diagnostic) => /data validation .*not rendered/u.test(diagnostic.message)), false);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders XLSX line, column, and win-loss sparklines from referenced cached values", async () => {
  const bytes = workbook(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
    <dimension ref="A1:D3"/><sheetData>
      <row r="1"><c r="A1"><v>5</v></c><c r="B1"><v>20</v></c><c r="C1"><v>10</v></c><c r="D1"/></row>
      <row r="2"><c r="A2"><v>-5</v></c><c r="B2"><v>10</v></c><c r="C2"><v>20</v></c><c r="D2"/></row>
      <row r="3"><c r="A3"><v>-1</v></c><c r="B3"><v>0</v></c><c r="C3"><v>1</v></c><c r="D3"/></row>
    </sheetData><extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups>
      <x14:sparklineGroup type="line" displayXAxis="1" minAxisType="custom" manualMin="0" maxAxisType="custom" manualMax="30">
        <x14:colorSeries rgb="FF4472C4"/><x14:colorAxis rgb="FF111111"/>
        <x14:sparklines><x14:sparkline><xm:f>Data!A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline></x14:sparklines>
      </x14:sparklineGroup>
      <x14:sparklineGroup type="column" displayXAxis="1">
        <x14:colorSeries rgb="FFED7D31"/><x14:colorNegative rgb="FFFF0000"/><x14:colorAxis rgb="FF222222"/>
        <x14:sparklines><x14:sparkline><xm:f>Data!A2:C2</xm:f><xm:sqref>D2</xm:sqref></x14:sparkline></x14:sparklines>
      </x14:sparklineGroup>
      <x14:sparklineGroup type="stacked" displayXAxis="1">
        <x14:colorSeries rgb="FF70AD47"/><x14:colorNegative rgb="FF7030A0"/><x14:colorAxis rgb="FF333333"/>
        <x14:sparklines><x14:sparkline><xm:f>Data!A3:C3</xm:f><xm:sqref>D3</xm:sqref></x14:sparkline></x14:sparklines>
      </x14:sparklineGroup>
    </x14:sparklineGroups></ext></extLst>
  </worksheet>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    const lineHits = await document.hitTest({ unitIndex: 0, x: 3 * 61 + 30, y: 10 });
    const columnHits = await document.hitTest({ unitIndex: 0, x: 3 * 61 + 30, y: 30 });
    const winLossHits = await document.hitTest({ unitIndex: 0, x: 3 * 61 + 30, y: 55 });
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.strokes.includes("rgba(68, 114, 196, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(237, 125, 49, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(255, 0, 0, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(112, 173, 71, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(112, 48, 160, 1)"), true);
    for (const [hits, address] of [[lineHits, "D1"], [columnHits, "D2"], [winLossHits, "D3"]]) {
      assert.equal(hits.some(({ object }) => object.type === "shape" && object.source.address === address && object.source.mapping === "exact"), true);
    }
    assert.equal(document.diagnostics().some((diagnostic) => /sparkline .*not rendered/u.test(diagnostic.message)), false);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("uses ODS row, column, merged-cell, and basic cell styles when rendering", async () => {
  const bytes = spreadsheet(`<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
    <office:automatic-styles>
      <style:style style:name="wide" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style>
      <style:style style:name="tall" style:family="table-row"><style:table-row-properties style:row-height="0.5in"/></style:style>
      <style:style style:name="accent" style:family="table-cell">
        <style:table-cell-properties fo:background-color="#ffe0b2" fo:border="0.75pt solid #ff0000"/>
        <style:text-properties fo:font-family="Aptos" fo:font-size="14pt" fo:font-weight="bold" fo:font-style="italic" fo:color="#336699"/>
        <style:paragraph-properties fo:text-align="center"/>
      </style:style>
    </office:automatic-styles>
    <office:body><office:spreadsheet><table:table table:name="Data">
      <table:table-column table:style-name="wide"/>
      <table:table-column table:visibility="collapse"/>
      <table:table-column table:number-columns-repeated="2"/>
      <table:table-row table:style-name="tall">
        <table:table-cell table:style-name="accent" office:value-type="string"><text:p>Styled</text:p></table:table-cell>
        <table:table-cell office:value-type="string"><text:p>hidden column</text:p></table:table-cell>
        <table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="3" office:value-type="string"><text:p>Merged ODS</text:p></table:table-cell>
        <table:covered-table-cell/>
      </table:table-row>
      <table:table-row table:visibility="collapse"><table:covered-table-cell table:number-columns-repeated="4"/></table:table-row>
      <table:table-row>
        <table:table-cell office:value-type="string"><text:p>Last</text:p></table:table-cell>
        <table:table-cell/>
        <table:covered-table-cell table:number-columns-repeated="2"/>
      </table:table-row>
    </table:table></office:spreadsheet></office:body>
  </office:document-content>`);
  const canvas = installRecordingCanvas();
  const { calls } = canvas;
  const { engine, document } = await open(bytes);
  let frame;
  try {
    assert.deepEqual(document.info.units, [
      { type: "sheet", index: 0, id: "unit:0", name: "Data", width: 288, height: 72, rows: 3, columns: 4, frozenRows: 0, frozenColumns: 0, frozenWidth: 0, frozenHeight: 0, rowAxis: { defaultSize: 24, spans: [{ start: 0, end: 0, size: 48 }, { start: 1, end: 1, size: 0 }] }, columnAxis: { defaultSize: 96, spans: [{ start: 1, end: 1, size: 0 }] }, showGridLines: true },
    ]);
    const styled = await document.hitTest({ unitIndex: 0, x: 20, y: 20 });
    assert.deepEqual(styled[0]?.object.bounds, { x: 0, y: 0, width: 96, height: 48 });
    const merged = await document.hitTest({ unitIndex: 0, x: 220, y: 60 });
    assert.equal(merged.length, 1);
    assert.equal(merged[0]?.object.text, "Merged ODS");
    assert.deepEqual(merged[0]?.object.bounds, { x: 96, y: 0, width: 192, height: 72 });
    const hiddenColumnBoundary = await document.hitTest({ unitIndex: 0, x: 96, y: 20 });
    assert.equal(hiddenColumnBoundary.some(({ object }) => object.text === "hidden column"), false);
    frame = await document.render({ unitIndex: 0 });
    assert.equal(calls.fills.includes("rgba(255, 224, 178, 1)"), true);
    assert.equal(calls.strokes.includes("rgba(255, 0, 0, 1)"), true);
    assert.equal(calls.fonts.some((font) => /Aptos/u.test(font) && /italic/u.test(font) && /700/u.test(font)), true);
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("exposes ODS frozen panes from view settings with real style geometry", async () => {
  const bytes = spreadsheet(`<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:automatic-styles><style:style style:name="wide" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style><style:style style:name="tall" style:family="table-row"><style:table-row-properties style:row-height="0.5in"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Data"><table:table-column table:style-name="wide" table:number-columns-repeated="2"/><table:table-row table:style-name="tall"><table:table-cell/></table:table-row><table:table-row/></table:table></office:spreadsheet></office:body></office:document-content>`, {
    "settings.xml": `<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="Data"><config:config-item config:name="HorizontalSplitMode" config:type="short">2</config:config-item><config:config-item config:name="HorizontalSplitPosition" config:type="int">1</config:config-item><config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item><config:config-item config:name="VerticalSplitPosition" config:type="int">2</config:config-item></config:config-item-map-entry></config:config-item-map-named></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`,
  });
  const { engine, document } = await open(bytes);
  try {
    assert.deepEqual(document.info.units, [{
      type: "sheet", index: 0, id: "unit:0", name: "Data", width: 192, height: 48,
      rows: 1, columns: 2, frozenRows: 1, frozenColumns: 2, frozenWidth: 192, frozenHeight: 48,
      rowAxis: { defaultSize: 24, spans: [{ start: 0, end: 0, size: 48 }] },
      columnAxis: { defaultSize: 96, spans: [] },
      showGridLines: true,
    }]);
  } finally {
    document.close();
    engine.close();
  }
});

test("renders embedded ODS images at draw frame geometry", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = spreadsheet(`<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row><table:table-cell/><draw:frame draw:id="logo" svg:x="0in" svg:y="0in" svg:width="0.5in" svg:height="0.2in"><draw:image xlink:href="Pictures/logo.png"/></draw:frame></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`, {
    "Pictures/logo.png": png,
  });
  const { engine, document } = await open(bytes);
  try {
    const hits = await document.hitTest({ unitIndex: 0, x: 10, y: 10 });
    assert.equal(hits[0]?.object.type, "image");
    assert.deepEqual(hits[0]?.object.bounds, { x: 0, y: 0, width: 48, height: 19.200000762939453 });
    assert.equal(
      document.diagnostics().some((diagnostic) => /drawing layer/u.test(diagnostic.message)),
      false,
    );
  } finally {
    document.close();
    engine.close();
  }
});

test("supplied ODS pie renders all seven legend labels in its custom three-column box", async () => {
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(await readFile(new URL("./fixtures/oasis-3883-chart-legend.ods", import.meta.url)));
  let frame;
  try {
    // Verify the emitted canvas text as well as the parser's objects: legend text must survive sheet clipping.
    const unit = document.info.units[0];
    frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: unit.width, height: unit.height }, scale: 1 });
    const names = ["Anne", "Tim", "Bob", "John", "Eve", "Georg", "James"];
    for (const name of names) {
      assert.equal(canvas.calls.text.filter(([text, x]) => text === name && x > 550).length, 1, name);
    }
    assert.equal(canvas.calls.drawImages.length, 0);
  } finally { frame?.bitmap.close(); document.close(); engine.close(); canvas.restore(); }
});

test("ODS floating images extend the viewport through unused authored columns", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const bytes = spreadsheet(`<office:document-content><office:automatic-styles>
    <style:style style:family="table-column" style:name="narrow"><style:table-column-properties style:column-width="40px"/></style:style>
    </office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Data">
    <table:table-column table:style-name="narrow" table:number-columns-repeated="10"/>
    <table:table-row><table:table-cell><text:p>cell</text:p></table:table-cell></table:table-row>
    <table:shapes><draw:frame svg:x="100px" svg:y="100px" svg:width="48px" svg:height="20px"><draw:image xlink:href="Pictures/logo.png"/></draw:frame></table:shapes>
    </table:table></office:spreadsheet></office:body></office:document-content>`, { "Pictures/logo.png": png });
  const { engine, document } = await open(bytes);
  try {
    const unit = document.info.units[0];
    assert.equal(unit.columns, 4);
    assert.equal(unit.rows, 5);
    assert.equal(unit.width, 160);
    assert.equal(unit.height, 120);
    assert.deepEqual(unit.columnAxis.spans, [{ start: 0, end: 3, size: 40 }]);
    assert.equal((await document.hitTest({ unitIndex: 0, x: 110, y: 110 }))[0]?.object.type, "image");
  } finally { document.close(); engine.close(); }
});

test("renders ODS value-comparison conditional styles", async () => {
  const bytes = spreadsheet(`<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
    <office:automatic-styles>
      <style:style style:name="alert" style:family="table-cell">
        <style:table-cell-properties fo:background-color="#ff0000"/>
        <style:text-properties fo:color="#ffffff"/>
      </style:style>
    </office:automatic-styles>
    <office:body><office:spreadsheet><table:table table:name="Data">
      <table:table-row>
        <table:table-cell office:value-type="float" office:value="10"/>
        <table:table-cell office:value-type="float" office:value="3"/>
      </table:table-row>
      <calcext:conditional-formats>
        <calcext:conditional-format calcext:target-range-address="Data.A1:Data.B1">
          <calcext:condition calcext:apply-style-name="alert" calcext:value="cell-content()&gt;5" calcext:base-cell-address="Data.A1"/>
        </calcext:conditional-format>
      </calcext:conditional-formats>
    </table:table></office:spreadsheet></office:body>
  </office:document-content>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.filter((fill) => fill === "rgba(255, 0, 0, 1)").length, 1);
    assert.equal(canvas.calls.fills.includes("rgba(255, 255, 255, 1)"), true);
    assert.equal(
      document.diagnostics().some((diagnostic) => /conditional formatting is not rendered/u.test(diagnostic.message)),
      false,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders ODS color scales and data bars", async () => {
  const bytes = spreadsheet(`<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0">
    <office:body><office:spreadsheet><table:table table:name="Data">
      <table:table-row>
        <table:table-cell office:value-type="float" office:value="0"/>
        <table:table-cell office:value-type="float" office:value="5"/>
        <table:table-cell office:value-type="float" office:value="10"/>
      </table:table-row>
      <table:table-row>
        <table:table-cell office:value-type="float" office:value="0"/>
        <table:table-cell office:value-type="float" office:value="5"/>
        <table:table-cell office:value-type="float" office:value="10"/>
      </table:table-row>
      <calcext:conditional-formats>
        <calcext:conditional-format calcext:target-range-address="Data.A1:Data.C1">
          <calcext:color-scale>
            <calcext:color-scale-entry calcext:type="minimum" calcext:value="0" calcext:color="#0000ff"/>
            <calcext:color-scale-entry calcext:type="percent" calcext:value="50" calcext:color="#ffff00"/>
            <calcext:color-scale-entry calcext:type="maximum" calcext:value="0" calcext:color="#ff0000"/>
          </calcext:color-scale>
        </calcext:conditional-format>
        <calcext:conditional-format calcext:target-range-address="Data.A2:Data.C2">
          <calcext:data-bar calcext:max-length="100" calcext:negative-color="#cc0000" calcext:positive-color="#00aa44" calcext:axis-color="#000000">
            <calcext:formatting-entry calcext:type="minimum" calcext:value="0"/>
            <calcext:formatting-entry calcext:type="maximum" calcext:value="0"/>
          </calcext:data-bar>
        </calcext:conditional-format>
      </calcext:conditional-formats>
    </table:table></office:spreadsheet></office:body>
  </office:document-content>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.includes("rgba(0, 0, 255, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(255, 255, 0, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(255, 0, 0, 1)"), true);
    assert.equal(canvas.calls.fills.some((fill) => /rgba\(0, 170, 68,/u.test(fill)), true);
    assert.equal(
      document.diagnostics().some((diagnostic) => /conditional-format rules could not be rendered/u.test(diagnostic.message)),
      false,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders inline ODS bar, line, and pie charts from inline table data", async () => {
  const chart = (kind, values) => `<chart:chart chart:class="chart:${kind}">
    <chart:plot-area><chart:series chart:values-cell-range-address="local.A1:local.A${values.length}"/></chart:plot-area>
    <table:table table:name="local">${values.map((value) => `<table:table-row><table:table-cell office:value-type="float" office:value="${value}"/></table:table-row>`).join("")}</table:table>
  </chart:chart>`;
  const bytes = spreadsheet(`<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0">
    <office:body><office:spreadsheet><table:table table:name="Data">
      <table:shapes>
        <draw:frame draw:id="bar" svg:x="0in" svg:y="0in" svg:width="2in" svg:height="1.5in">${chart("bar", [10, 30, 20])}</draw:frame>
        <draw:frame draw:id="line" svg:x="2.2in" svg:y="0in" svg:width="2in" svg:height="1.5in">${chart("line", [5, 25, 15])}</draw:frame>
        <draw:frame draw:id="pie" svg:x="4.4in" svg:y="0in" svg:width="2in" svg:height="1.5in">${chart("circle", [1, 2, 3])}</draw:frame>
      </table:shapes>
      <table:table-row><table:table-cell/></table:table-row>
    </table:table></office:spreadsheet></office:body>
  </office:document-content>`);
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    const barHits = await document.hitTest({ unitIndex: 0, x: 40, y: 110 });
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.includes("rgba(68, 114, 196, 1)"), true);
    assert.equal(canvas.calls.fills.includes("rgba(237, 125, 49, 1)"), true);
    assert.equal(canvas.calls.strokes.includes("rgba(68, 114, 196, 1)"), true);
    assert.equal(frame.renderedObjectCount >= 12, true);
    assert.equal(barHits.some(({ object }) => object.source.elementId === "bar" && object.source.part === "content.xml"), true);
    assert.equal(
      document.diagnostics().some((diagnostic) => /ODS chart .*no usable/u.test(diagnostic.message)),
      false,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("renders safely referenced ODS chart objects with embedded-part source mapping", async () => {
  const embeddedChart = `<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0">
    <office:body><office:chart><chart:chart chart:class="chart:bar">
      <chart:plot-area><chart:series chart:values-cell-range-address="local.A1:local.A3"/></chart:plot-area>
      <table:table table:name="local">
        <table:table-row><table:table-cell office:value-type="float" office:value="10"/></table:table-row>
        <table:table-row><table:table-cell office:value-type="float" office:value="30"/></table:table-row>
        <table:table-row><table:table-cell office:value-type="float" office:value="20"/></table:table-row>
      </table:table>
    </chart:chart></office:chart></office:body>
  </office:document-content>`;
  const bytes = spreadsheet(`<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink">
    <office:body><office:spreadsheet><table:table table:name="Data">
      <table:shapes><draw:frame draw:id="embedded" svg:x="1in" svg:y="0in" svg:width="2in" svg:height="1.5in"><draw:object xlink:href="./Object 1"/><draw:image xlink:href="./ObjectReplacements/Object 1"/></draw:frame></table:shapes>
      <table:table-row><table:table-cell/></table:table-row>
    </table:table></office:spreadsheet></office:body>
  </office:document-content>`, { "Object 1/content.xml": embeddedChart });
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  let frame;
  try {
    const hits = await document.hitTest({ unitIndex: 0, x: 135, y: 110 });
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.fills.includes("rgba(68, 114, 196, 1)"), true);
    assert.equal(hits.some(({ object }) => object.source.part === "Object 1/content.xml" && object.source.elementId === "embedded"), true);
    assert.equal(
      document.diagnostics().some((diagnostic) => /embedded chart has no usable/u.test(diagnostic.message)),
      false,
    );
    assert.equal(
      document.diagnostics().some((diagnostic) => /image alternative is missing/u.test(diagnostic.message)),
      true,
    );
  } finally {
    frame?.bitmap.close();
    document.close();
    engine.close();
    canvas.restore();
  }
});

test("diagnoses blocked and data-less ODS chart objects without placeholders", async () => {
  const emptyChart = `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:plot-area/></chart:chart></office:chart></office:body></office:document-content>`;
  const bytes = spreadsheet(`<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink">
    <office:body><office:spreadsheet><table:table table:name="Data"><table:shapes>
      <draw:frame svg:x="0in" svg:y="0in" svg:width="1in" svg:height="1in"><draw:object xlink:href="https://example.test/chart"/></draw:frame>
      <draw:frame svg:x="1in" svg:y="0in" svg:width="1in" svg:height="1in"><draw:object xlink:href="./Object 2"/></draw:frame>
    </table:shapes><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body>
  </office:document-content>`, { "Object 2/content.xml": emptyChart });
  const { engine, document } = await open(bytes);
  try {
    const diagnostics = document.diagnostics();
    assert.equal(diagnostics.some((diagnostic) => diagnostic.code === "EXTERNAL_RESOURCE_BLOCKED" && /chart object reference/u.test(diagnostic.message)), true);
    assert.equal(diagnostics.some((diagnostic) => /embedded chart has no usable/u.test(diagnostic.message)), true);
    assert.deepEqual(await document.hitTest({ unitIndex: 0, x: 48, y: 48 }), []);
    assert.deepEqual(await document.hitTest({ unitIndex: 0, x: 144, y: 48 }), []);
  } finally {
    document.close();
    engine.close();
  }
});


test("renders Japanese era dates from tdf161301.xlsx", async () => {
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(await readFile(new URL("./fixtures/tdf161301.xlsx", import.meta.url)));
  try {
    const frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: 900, height: 300 }, scale: 1, pixelRatio: 1 });
    frame.bitmap.close();
    assert.equal(canvas.calls.text.map(([text]) => text).join("").split("令和6年5月28日").length - 1, 2, JSON.stringify(canvas.calls.text));
    assert.equal(canvas.calls.text.some(([text]) => text.includes("ggge")), false);
  } finally {
    canvas.restore();
    document.close();
    engine.close();
  }
});

test("preserves Japanese era transitions and first-year notation", async () => {
  const cases = [[43586, "令和元年5月1日"], [43585, "平成31年4月30日"], [32516, "平成元年1月8日"], [32515, "昭和64年1月7日"], [9856, "昭和元年12月25日"], [9855, "大正15年12月24日"], [4595, "大正元年7月30日"], [4594, "明治45年7月29日"]];
  const bytes = workbook(
    `<worksheet><sheetData>${cases.map(([serial], i) => `<row r="${i + 1}"><c r="A${i + 1}" s="0"><v>${serial}</v></c></row>`).join("")}</sheetData></worksheet>`,
    `<styleSheet><numFmts><numFmt numFmtId="164" formatCode="[$-ja-JP-x-gannen]ggge&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/></numFmts><cellXfs><xf numFmtId="164"/></cellXfs></styleSheet>`,
  );
  const canvas = installRecordingCanvas();
  const { engine, document } = await open(bytes);
  try {
    const frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: 900, height: 900 }, scale: 1, pixelRatio: 1 });
    frame.bitmap.close();
    assert.equal(canvas.calls.text.map(([text]) => text).join(""), cases.map(([, expected]) => expected).join(""));
  } finally { canvas.restore(); document.close(); engine.close(); }
});

// XLSX/XLS fixtures are LibreOffice exports of the supplied OASIS 2173 workbook.
for (const extension of ['ods', 'xlsx', 'xlsm', 'xls']) {
  test(`preserves authored worksheet tab colors through ${extension.toUpperCase()} metadata and rendering`, async () => {
    let bytes = await readFile(new URL(`./fixtures/oasis-2173-tab-color.${extension === 'xlsm' ? 'xlsx' : extension}`, import.meta.url));
    if (extension === 'xlsm') {
      const parts = Object.fromEntries(readZipEntries(bytes).map(({ name, data }) => [name, data]));
      parts['[Content_Types].xml'] = new TextDecoder().decode(parts['[Content_Types].xml'])
        .replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml');
      bytes = createZip(parts);
    }
    const { engine, document } = await open(bytes);
    try {
      assert.deepEqual(document.info.units.map(({ name, tabColor }) => [name, tabColor]), [
        ['Bar', 0xff0000ff], ['Bubble', 0x0000ffff],
      ]);
      for (const unit of document.info.units) {
        await document.hitTest({ unitIndex: unit.index, x: 4, y: 4 });
      }
      assert.deepEqual(document.info.units.map(unit => unit.tabColor), [0xff0000ff, 0x0000ffff]);
    } finally { document.close(); engine.close(); }
  });
}

test('worksheet tab colors preserve white, inheritance, theme/indexed colors, and optional invalid-color recovery', async () => {
  for (const [extension, cases] of [
    ['ods', [
      ['table:tab-color="#ffffff"', 0xffffffff],
      ['', undefined],
      ['table:tab-color="invalid"', undefined],
    ]],
    ['xlsx', [
      ['rgb="FFFFFFFF"', 0xffffffff],
      ['theme="4"', 0x18a303ff],
      ['theme="4" tint="1"', 0xffffffff],
      ['indexed="10"', 0xff0000ff],
      ['auto="1"', undefined],
      ['rgb="invalid"', undefined],
    ]],
  ]) {
    const original = await readFile(new URL(`./fixtures/oasis-2173-tab-color.${extension}`, import.meta.url));
    for (const [replacement, expected] of cases) {
      const parts = Object.fromEntries(readZipEntries(original).map(({ name, data }) => [name, data]));
      const part = extension === 'ods' ? 'content.xml' : 'xl/worksheets/sheet1.xml';
      const xml = new TextDecoder().decode(parts[part]);
      parts[part] = extension === 'ods'
        ? xml.replace('table:tab-color="#ff0000"', replacement)
        : xml.replace('tabColor rgb="FFFF0000"', `tabColor ${replacement}`);
      const { engine, document } = await open(createZip(parts));
      try {
        assert.equal(document.info.units[0].tabColor, expected, `${extension}: ${replacement}`);
        await document.hitTest({ unitIndex: 0, x: 4, y: 4 });
        assert.equal(document.info.units[0].tabColor, expected);
        assert.equal(document.info.units[1].tabColor, 0x0000ffff);
        if (replacement.includes('invalid')) {
          assert.ok(document.diagnostics().some(d => d.message.includes('tab color')));
        }
      } finally { document.close(); engine.close(); }
    }
  }
  const original = await readFile(new URL('./fixtures/oasis-2173-tab-color.ods', import.meta.url));
  const parts = Object.fromEntries(readZipEntries(original).map(({ name, data }) => [name, data]));
  parts['content.xml'] = new TextDecoder().decode(parts['content.xml'])
    .replace('style:name="ta1" style:family="table"', 'style:name="ta1" style:family="table" style:parent-style-name="tab-parent"')
    .replace('table:tab-color="#ff0000"', '')
    .replace('</office:automatic-styles>', '<style:style style:name="tab-parent" style:family="table"><style:table-properties table:tab-color="#ff0000"/></style:style></office:automatic-styles>');
  const { engine, document } = await open(createZip(parts));
  try { assert.equal(document.info.units[0].tabColor, 0xff0000ff); }
  finally { document.close(); engine.close(); }
});

for (const [file, check] of [
    ["ooxml-print-titles.xlsx", unit => assert.equal(unit.printSettings?.printTitles, "DIET!$5:$5")],
    ["ooxml-manual-breaks.xlsx", unit => {
      assert.deepEqual(unit.printSettings?.rowBreaks, [[8, 0, 16383]]);
      assert.ok(Object.isFrozen(unit.printSettings.rowBreaks));
      assert.ok(Object.isFrozen(unit.printSettings.rowBreaks[0]));
    }],
]) {
  test(`real XLSX print metadata ${file}`, async () => {
    const { engine, document } = await open(await readFile(new URL(`fixtures/${file}`, import.meta.url)));
    try { check(document.info.units[0]); }
    finally { document.close(); engine.close(); }
  });
}

test("real XLSX manual breaks and repeated titles compose printable cell regions", async () => {
  const opened = await open(await readFile(new URL("fixtures/ooxml-manual-breaks.xlsx", import.meta.url)));
  try {
    const original = opened.document.info.units[0];
    const unit = { ...original, printSettings: { ...original.printSettings, printArea: "Start!$B$1:$B$41" } };
    const pages = sheetPrintPages({ ...unit, printSettings: { ...unit.printSettings, fitToPage: false } });
    const row8 = unit.rowAxis.defaultSize * 8 + unit.rowAxis.spans.reduce((sum, span) =>
      sum + Math.max(0, Math.min(8, span.end + 1) - span.start) * (span.size - unit.rowAxis.defaultSize), 0);
    assert.equal(pages[0].viewport.height, row8);
    assert.equal(pages[1].viewport.y, row8);
    const fitted = sheetPrintPages({ ...unit, printSettings: { ...unit.printSettings, fitToPage: true, fitToWidth: 1, fitToHeight: 1 } });
    assert.equal(fitted.length, 1, "fit-to-page ignores manual breaks");
  } finally { opened.document.close(); opened.engine.close(); }
  const { engine, document } = await open(await readFile(new URL("fixtures/ooxml-print-titles.xlsx", import.meta.url)));
  try {
    const unit = document.info.units[0];
    const pages = sheetPrintPages({ ...unit, printSettings: {
      ...unit.printSettings, fitToPage: false, scale: 100, rowBreaks: [[10, 0, 16383]],
    } });
    const repeated = pages.find(page => page.fragments.length > 1);
    assert.ok(repeated, "later print page must include the authored row 5 title");
    const [title, body] = repeated.fragments;
    assert.ok(title.viewport.y < body.viewport.y);
    assert.equal(title.y, 0);
    assert.equal(body.y, title.viewport.height);
    assert.equal(repeated.height, title.viewport.height + body.viewport.height);
    const both = sheetPrintPages({ ...unit, printSettings: { ...unit.printSettings, fitToPage: false, scale: 100,
      printTitles: "'DIET, chart'!$5:$5,'DIET, chart'!$B:$B", rowBreaks: [[10,0,16383]], columnBreaks: [[5,0,1048575]],
    } });
    assert.ok(both.some(page => page.fragments.length === 4), "row, column, corner and body regions");
    const overlapping = sheetPrintPages({ ...unit, printSettings: { ...unit.printSettings, fitToPage: false,
      printArea: "DIET!$B$5:$I$20", printTitles: "DIET!$4:$6",
    } });
    assert.equal(overlapping[0].fragments.length, 2, "title portion before print area must remain visible");
    assert.equal(overlapping[0].fragments[0].viewport.y + overlapping[0].fragments[0].viewport.height,
      overlapping[0].fragments[1].viewport.y, "overlapping title rows must not be duplicated");
  } finally { document.close(); engine.close(); }
});

for (const carrier of ["ooxml-minor-ticks.xlsx", "ooxml-display-units.docx", "corpus-chart-wall.pptx"]) {
  test(`real ChartML minor ticks reach ${carrier}`, async () => {
    const source = await readFile(new URL("fixtures/ooxml-minor-ticks.xlsx", import.meta.url));
    const chart = readZipEntries(source).find(entry => entry.name === "xl/charts/chart1.xml").data;
    const parts = Object.fromEntries(readZipEntries(await readFile(new URL(`fixtures/${carrier}`, import.meta.url))).map(({name, data}) => [name, data]));
    parts[Object.keys(parts).find(name => /^(word|ppt|xl)\/charts\/chart1\.xml$/u.test(name))] = chart;
    const {engine, document} = await open(createZip(parts));
    try {
      const objects = await document.listObjects();
      assert.ok(objects.filter(object => Math.abs(object.bounds.width - 3) < .01 && object.bounds.height <= .02).length >= 12);
    } finally { document.close(); engine.close(); }
  });
}

// Exercise authored minor-grid styling on the real line chart in each host.
for (const carrier of ["ooxml-minor-ticks.xlsx", "ooxml-display-units.docx", "corpus-chart-wall.pptx"]) {
  test(`derived ChartML minor gridlines reach ${carrier}`, async () => {
    const source = readZipEntries(await readFile(new URL("fixtures/ooxml-minor-ticks.xlsx", import.meta.url)));
    const chart = new TextDecoder().decode(source.find(entry => entry.name === "xl/charts/chart1.xml").data)
      .replace("</c:valAx>", '<c:minorGridlines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr></c:minorGridlines><c:minorUnit val="0.5"/></c:valAx>');
    const parts = Object.fromEntries(readZipEntries(await readFile(new URL(`fixtures/${carrier}`, import.meta.url))).map(({name, data}) => [name, data]));
    parts[Object.keys(parts).find(name => /^(word|ppt|xl)\/charts\/chart1\.xml$/u.test(name))] = chart;
    const {engine, document} = await open(createZip(parts));
    try {
      const canvas = installRecordingCanvas();
      let frame;
      try {
        const unit = document.info.units[0];
        frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: unit.width, height: unit.height }, scale: 1 });
        assert.equal(canvas.calls.paintedStrokes.filter(([color, width]) => color === "rgba(18, 52, 86, 1)" && width === 2).length, 7);
      } finally { frame?.bitmap.close(); canvas.restore(); }
    } finally { document.close(); engine.close(); }
  });
}

test("real date axis paints calendar month labels and weekly gridlines", async () => {
  const {engine, document} = await open(await readFile(new URL("fixtures/ooxml-calendar-axis.xlsx", import.meta.url)));
  const canvas = installRecordingCanvas();
  let frame;
  try {
    const all = await document.listObjects();
    const chart = all.filter(object => object.source.part === "xl/charts/chart1.xml");
    const months = ["1/1/2019", "2/1/2019", "3/1/2019", "4/1/2019", "5/1/2019"];
    for (const month of months) assert.ok(chart.some(object => object.text === month), `missing calendar month ${month}`);
    const unit = document.info.units[0];
    frame = await document.render({ unitIndex: 0, viewport: { x: 0, y: 0, width: unit.width, height: unit.height }, scale: 1 });
    assert.ok(canvas.calls.paintedStrokes.filter(([color]) => color === "rgba(221, 221, 221, 1)").length >= 12);
  } finally { frame?.bitmap.close(); canvas.restore(); document.close(); engine.close(); }
});

test("UTF-16 OOXML package metadata and worksheet preserve real workbook display", async () => {
  const source = await readFile(new URL("fixtures/ooxml-minor-ticks.xlsx", import.meta.url));
  const parts = Object.fromEntries(readZipEntries(source).map(({name, data}) => [name, data]));
  for (const name of ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels", "xl/worksheets/sheet1.xml"]) {
    const xml = new TextDecoder().decode(parts[name]).replace('encoding="UTF-8"', 'encoding="UTF-16"');
    parts[name] = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(xml, "utf16le")]);
  }
  const original = await open(source);
  let converted;
  try {
    converted = await open(createZip(parts));
    const expected = (await original.document.listObjects()).map(({text, bounds}) => ({text, bounds}));
    const actual = (await converted.document.listObjects()).map(({text, bounds}) => ({text, bounds}));
    assert.ok(expected.some(({text}) => text === "col1"));
    assert.deepEqual(actual, expected);
  } finally {
    original.document.close(); original.engine.close();
    converted?.document.close(); converted?.engine.close();
  }
});

for (const [carrier, part, token] of [
  ["ooxml-nested-alternate.docx", "word/document.xml", "ABC"],
  ["corpus-chart-wall.pptx", "ppt/slides/slide1.xml", null],
]) {
  test(`UTF-16 main XML preserves ${carrier} display`, async () => {
    const source = await readFile(new URL(`fixtures/${carrier}`, import.meta.url));
    const parts = Object.fromEntries(readZipEntries(source).map(({name, data}) => [name, data]));
    const xml = new TextDecoder().decode(parts[part]).replace('encoding="UTF-8"', 'encoding="UTF-16"');
    parts[part] = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(xml, "utf16le")]);
    const original = await open(source);
    let converted;
    try {
      converted = await open(createZip(parts));
      const visible = async document => (await document.listObjects()).map(({text, bounds}) => ({text, bounds}));
      const expected = await visible(original.document);
      assert.ok(expected.length > 0);
      if (token) assert.ok(expected.some(({text}) => text?.includes(token)));
      assert.deepEqual(await visible(converted.document), expected);
    } finally {
      original.document.close(); original.engine.close();
      converted?.document.close(); converted?.engine.close();
    }
  });
}

for (const carrier of ["ooxml-display-units.docx", "corpus-chart-wall.pptx"]) {
  test(`real calendar ChartML reaches ${carrier}`, async () => {
    const source = readZipEntries(await readFile(new URL("fixtures/ooxml-calendar-axis.xlsx", import.meta.url)));
    const chart = source.find(entry => entry.name === "xl/charts/chart1.xml").data;
    const parts = Object.fromEntries(readZipEntries(await readFile(new URL(`fixtures/${carrier}`, import.meta.url))).map(({name, data}) => [name, data]));
    parts[Object.keys(parts).find(name => /^(word|ppt)\/charts\/chart1\.xml$/u.test(name))] = chart;
    const {engine, document} = await open(createZip(parts));
    try {
      const objects = await document.listObjects();
      for (const month of ["1/1/2019", "2/1/2019", "3/1/2019", "4/1/2019", "5/1/2019"])
        assert.ok(objects.some(object => object.text === month), `${carrier}: missing ${month}`);
    } finally { document.close(); engine.close(); }
  });
}
