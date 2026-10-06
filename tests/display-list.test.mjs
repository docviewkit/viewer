import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { Core, DEFAULT_LIMITS } from "../dist/core.js";
import { decodeSnapshot } from "../dist/protocol.js";
import { collectFontRequests } from "../dist/font.js";
import {
  layoutTextRuns,
  SceneRenderer,
  sheetAutoFitColumnWidth,
  sheetAutoFitRowHeight,
} from "../dist/render.js";

test("xl8galry preserves chart styles, geometry and legends across OOXML hosts", async () => {
  const { readZipEntries } = await import("../scripts/accuracy-metamorphic.mjs");
  const { createZip } = await import("./zip-fixture.mjs");
  const fixture = await readFile(new URL("./fixtures/xl8galry.xlsx", import.meta.url));
  const charts = new Map(readZipEntries(fixture).map(({name, data}) => [name, data]));
  const core = await Core.create(await readFile(process.env.XL8GALRY_WASM ?? new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  try {
    for (const [host, carrier, part] of [
      ["xlsx", null, "xl/charts/chart1.xml"],
      ["docx", "chart-original.docx", "word/charts/chart1.xml"],
      ["pptx", "chart_pt_color_bg1.pptx", "ppt/charts/chart1.xml"],
    ]) {
      const parts = carrier ? Object.fromEntries(readZipEntries(await readFile(new URL(`./fixtures/${carrier}`, import.meta.url))).map(({name, data}) => [name, data])) : null;
      if (parts) {
        const directory = part.slice(0, part.lastIndexOf('/') + 1);
        parts[`${directory}xl8galry-theme.xml`] = charts.get('xl/theme/theme1.xml');
        parts[`${directory}_rels/chart1.xml.rels`] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="xl8galryTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" Target="xl8galry-theme.xml"/></Relationships>`;
      }
      for (const number of [1, 7, 10, 11, 17]) {
        if (parts) parts[part] = charts.get(`xl/charts/chart${number}.xml`);
        const document = core.open(parts ? createZip(parts) : fixture);
        try {
          const index = parts ? 0 : number - 1;
          document.loadUnit(index);
          const objects = document.scene.objects.filter(o => o.unitIndex === index && o.source?.part.includes("/charts/"));
          const visuals = objects.map(o => { let v = o.visual; while (v.visual) v = v.visual; return v; });
          if (number === 1) {
            assert.ok(objects.some(o => o.text === "1189.679009"), `${host}: General preserves source precision`);
            const bars = objects.filter(o => { let v = o.visual; while (v.visual) v = v.visual; return v.fill?.kind === "solid" && v.fill.color === 0x9999ffff && o.bounds.width > 20; });
            assert.ok(bars.length > 0, `${host}: North bar`);
            for (const bar of bars) {
              let v = bar.visual;
              assert.ok(v.shadow || v.outerShadow, `${host}: theme shadow on North bar`);
              while (v.visual) v = v.visual;
              assert.equal(v.stroke?.color, 0x000000ff, `${host}: authored black border`);
              assert.ok(v.strokeWidth > 1, `${host}: authored border width`);
            }
          }
          if (number === 7) {
            for (const color of [0x000080ff, 0xff00ffff]) {
              const line = objects.find(o => { let v = o.visual; while (v.visual) v = v.visual; return v.geometry === 'line' && v.stroke?.color === color; });
              assert.ok(line, `${host}: legend line exists`);
              const marker = objects.find(o => { let v = o.visual; while (v.visual) v = v.visual; return v.geometry !== 'line' && v.fill?.color === color && Math.abs(o.bounds.x + o.bounds.width / 2 - line.bounds.x - line.bounds.width / 2) < 1 && Math.abs(o.bounds.y + o.bounds.height / 2 - line.bounds.y) < 1; });
              assert.ok(marker, `${host}: centered legend marker`);
              assert.ok(line.bounds.width >= marker.bounds.width * 2, `${host}: line must remain visible on both sides of marker`);
            }
          }
          if (number === 10) {
            const labels = objects.filter(o => /^Series\s*\d+$/u.test(o.text ?? '')).map(o => o.text.replace(/\s/gu, ''));
            assert.deepEqual(labels, Array.from({ length: 10 }, (_, i) => `Series${i + 1}`), `${host}: Tubes legend follows left-to-right series order`);
          }
          if (number === 11) {
            assert.ok(visuals.filter(v => v.geometry?.kind === "path").length >= 24, `${host}: pie3DChart retains side faces despite point bubble3D=0`);
            const labels = objects.filter(o => o.text && o.text !== 'Pie Explosion');
            assert.equal(labels.length, 12, `${host}: all pie categories survive`);
            for (const label of labels) {
              let v = label.visual;
              while (v.visual) v = v.visual;
              assert.ok(label.bounds.height <= v.runs[0].fontSize * 1.5, `${host}: ${label.text} stays on one compact line`);
            }
            const leaders = visuals.filter(v => v.geometry?.commands?.length === 2 && v.stroke?.color === 0xffffffff);
            assert.ok(leaders.length <= 4, `${host}: adjacent labels do not need leader lines`);
            for (const [near, far] of [[0xea7575ff, 0x005dbbff], [0xbbeaeaff, 0xeaeabbff]]) {
              const front = visuals.findIndex(v => v.fill?.stops?.[0]?.color === near && v.geometry?.commands?.length > 5);
              const back = visuals.findIndex(v => v.fill?.stops?.[0]?.color === far && v.geometry?.commands?.length === 5);
              assert.ok(front >= 0 && back >= 0 && front > back, `${host}: rear radial wall must not cut through the nearer curved face`);
            }
          }
          if (number === 17) assert.ok(visuals.some(v => v.fill?.kind === "pattern" && v.fill.preset === "pct75"), `${host}: point pattern reaches the scene`);
        } finally { document.close(); }
      }
    }
  } finally { core.close(); }
});

test("supplied tdf105517 preserves chart formatting across PPTX DOCX and XLSX", async () => {
  const { readZipEntries } = await import("../scripts/accuracy-metamorphic.mjs");
  const { createZip } = await import("./zip-fixture.mjs");
  const original = await readFile(new URL("./fixtures/corpus-tdf105517.pptx", import.meta.url));
  const parts = new Map(readZipEntries(original).map(({ name, data }) => [name, data]));
  const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const chartParts = {
    "charts/chart1.xml": parts.get("ppt/charts/chart1.xml"),
    "charts/_rels/chart1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/themeOverride" Target="../theme/themeOverride1.xml"/></Relationships>`,
    "theme/themeOverride1.xml": parts.get("ppt/theme/themeOverride1.xml"),
  };
  const common = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`,
  };
  const docx = createZip({ ...common, ...chartParts,
    "[Content_Types].xml": common["[Content_Types].xml"].replace("</Types>", '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/document.xml": `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="9036496" cy="4718478"/><wp:docPr id="1" name="chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="chart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="18000" w:h="14000"/></w:sectPr></w:body></w:document>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart" Type="${rel}/chart" Target="../charts/chart1.xml"/></Relationships>`,
  });
  const xlsx = createZip({ ...common, ...chartParts,
    "[Content_Types].xml": common["[Content_Types].xml"].replace("</Types>", '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>'),
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Chart" sheetId="1" r:id="s1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="d1"/></worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="d1" Type="${rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9036496" cy="4718478"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="chart"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="chart"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:absoluteAnchor></xdr:wsDr>`,
    "xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart" Type="${rel}/chart" Target="../../charts/chart1.xml"/></Relationships>`,
  });
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  try {
    for (const [host, bytes] of [["PPTX", original], ["DOCX", docx], ["XLSX", xlsx]]) {
      const document = core.open(bytes);
      try {
        assert.ok(document.scene.info?.units.length > 0, `${host} must expose its chart page: ${JSON.stringify(document.scene.diagnostics)}`);
        document.loadUnit(0);
        const objects = document.scene.objects;
        for (const text of ["1,200,000", "1,000,000", "220,000"]) {
          const labels = objects.filter(o => o.text === text);
          assert.equal(labels.length, 1, `${host} ${text}`);
          let visual = labels[0].visual;
          while (visual.visual) visual = visual.visual;
          assert.equal(visual.runs[0].fontSize, 32, `${host} authored 24pt`);
          assert.equal(visual.runs[0].bold, true);
          assert.ok(labels[0].bounds.width > 120, `${host} label room`);
        }
        const shapes = objects.map(o => { let v = o.visual; while (v.visual) v = v.visual; return v; });
        assert.equal(shapes.filter(v => v.fill?.color === 0x226ca9ff && v.geometry?.commands?.length === 4).length, 5, `${host} triangles`);
        for (const color of [0xf0ab00ff, 0x226ca9ff]) {
          assert.equal(shapes.filter(v => v.stroke?.color === color && v.geometry?.commands?.length === 5).length, 1, `${host} one series path preserves all five points and its color`);
        }
      } finally { document.close(); }
    }
  } finally { core.close(); }
});

test("supplied bnc904423 inherits the green layout fill and preserves local fills", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/bnc904423.pptx", import.meta.url)));
  try {
    for (const [label, color] of [["Green", 0x00cc99ff], ["Blue", 0x3333ccff], ["Red", 0xff0000ff]]) {
      const object = document.scene.objects.find(object => object.text?.startsWith(label));
      assert.ok(object, label);
      let visual = object.visual;
      while (visual.visual) visual = visual.visual;
      assert.deepEqual(visual.fill, { kind: "solid", color }, label);
    }
  } finally { document.close(); core.close(); }
});

for (const name of ["gears", "curved arrows"]) {
  test(`supplied cshapes preserves ${name}`, async () => {
    const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
    const document = core.open(await readFile(new URL("./fixtures/cshapes.pptx", import.meta.url)));
    try {
      document.loadUnit(0);
      for (const shape of name === "gears" ? ["gear6", "gear9"] : ["curvedLeftArrow", "curvedRightArrow", "curvedUpArrow", "curvedDownArrow"]) {
        const object = document.scene.objects.find(o => o.name === shape);
        let visual = object.visual;
        while (visual.visual) visual = visual.visual;
        assert.equal(typeof visual.geometry, "object", `${shape} must not fall back to a rectangle`);
        const commands = visual.geometry.commands ?? visual.geometry.layers.flatMap(layer => layer.commands);
        assert.ok(commands.length > 20, `${shape} must retain its authored outline`);
        for (const command of commands.filter(c => c.kind === "bezierCurveTo")) {
          assert.ok(command.x >= -.01 && command.x <= object.bounds.width + .01
            && command.y >= -.01 && command.y <= object.bounds.height + .01,
          `${shape} arc endpoint must stay on the authored shape: ${JSON.stringify(command)}`);
        }
      }
    } finally { document.close(); core.close(); }
  });
}

const encoder = new TextEncoder();

test("supplied ODS number fills survive the snapshot and render at all three positions", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-odf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/oasis-3765-number-fill-character.ods", import.meta.url)));
  const context = recordingContext();
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: class {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  } });
  try {
    const cells = document.scene.objects.filter(object => object.type === "cell");
    assert.equal(cells.length, 12);
    for (const [index, cell] of cells.entries()) {
      context.calls.length = 0;
      const frame = await new SceneRenderer([cell], DEFAULT_LIMITS).render(document.scene.info.units[0], { unitIndex: 0 });
      frame.bitmap.close();
      const painted = context.calls.filter(call => call[0] === "fillText").map(call => call[1]).join("");
      assert.equal(painted.replaceAll("―", ""), cell.text);
      if (index < 3) assert.equal(painted, cell.text);
      else {
        assert.match(painted, /―+/u);
        assert.equal(painted.indexOf("―"), index < 6 ? 0 : index < 9 ? cell.text.length - 1 : cell.text.length);
      }
    }
  } finally {
    if (original === undefined) delete globalThis.OffscreenCanvas;
    else Object.defineProperty(globalThis, "OffscreenCanvas", original);
    document.close(); core.close();
  }
});

test("fill characters use measured spare width and preserve source offsets", () => {
  const context = recordingContext();
  context.measureText = function(text) {
    return { width: [...text].length * Number(/([\d.]+)px/u.exec(this.font)?.[1] ?? 7) / 1.4 };
  };
  const run = { text: "7,89€", fontFamily: "Arial", fontSize: 14,
    color: 0xff, bold: false, italic: false, letterSpacing: 0 };
  for (const offset of [0, 4, 5]) {
    for (const maxWidth of [40, 55, 100, 130]) {
      const layout = layoutTextRuns(context, [run], { maxWidth, wrap: false,
        fillCharacter: { offset, character: "―" } });
      const count = Math.max(0, Math.floor((maxWidth - 50) / 10));
      assert.equal(layout.runs.map(item => item.text).join(""),
        run.text.slice(0, offset) + "―".repeat(count) + run.text.slice(offset));
      for (const item of layout.runs.filter(item => item.text.includes("―"))) {
        assert.equal(item.start, offset);
        assert.equal(item.end, offset);
      }
    }
  }
});

test("supplied fit-to-size FODP keeps Fontwork envelopes and centered stretched text", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-odf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/fit-to-size.fodp", import.meta.url)));
  try {
    const fontwork = document.scene.objects.filter(o => o.text?.includes("Fontwork"));
    assert.equal(fontwork.length, 2);
    for (const object of fontwork) {
      let visual = object.visual;
      while (visual.kind === "layer") visual = visual.visual;
      assert.equal(visual.layout.warp, "text-envelope");
      assert.equal(visual.visual.geometry.kind, "layered-path");
      assert.equal(visual.visual.geometry.layers.length, 2);
      assert.equal(visual.visual.fill.kind, "none");
      assert.ok(visual.visual.runs.filter(run => run.text.trim()).every(run => run.fontFamily === "Arial Black"));
    }
    const run = { text: "2\nfoo", fontFamily: "Arial", fontSize: 32,
      color: 0x000000ff, bold: false, italic: false, letterSpacing: 0 };
    const context = recordingContext();
    context.measureText = function(text) { return { width: text.length * Number(/([\d.]+)px/u.exec(this.font)?.[1] ?? 32) / 2 }; };
    const layout = layoutTextRuns(context, [run], {
      maxWidth: 476, maxHeight: 192, lineHeight: 38.4,
      align: "center", autoFit: "fit-frame", wrap: false,
    });
    for (const placement of layout.runs) {
      assert.ok(Math.abs(placement.x + placement.width / 2 - layout.width / 2) < .01,
        "stretch must preserve the shared paragraph center without scaling its outer alignment gap");
    }
  } finally { document.close(); core.close(); }
});

test("supplied bullet-indent PPTX aligns positive and hanging bullet paragraphs", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/bullet-indent.pptx", import.meta.url)));
  try {
    const object = document.scene.objects.find(o => o.text?.includes("5678"));
    const layout = layoutTextRuns(recordingContext(), object.visual.visual.runs, {
      ...object.visual.layout, maxWidth: object.bounds.width,
      align: object.visual.visual.align, lineHeight: object.visual.visual.lineHeight,
    });
    assert.deepEqual(layout.runs.filter(run => run.text === "●").map(run => run.x), [0, 0]);
    for (const text of ["5678", "1234"]) {
      assert.ok(Math.abs(layout.runs.find(run => run.text === text).x - 324000 / 9525) < .01);
    }
  } finally { document.close(); core.close(); }
});

test("supplied n759180 PPTX keeps list anchors after manual line breaks", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/n759180.pptx", import.meta.url)));
  try {
    const object = document.scene.objects.find(o => o.text?.includes("Textrun3"));
    const layout = layoutTextRuns(recordingContext(), object.visual.visual.runs, {
      ...object.visual.layout, maxWidth: object.bounds.width,
      align: object.visual.visual.align, lineHeight: object.visual.visual.lineHeight,
    });
    const items = layout.runs.filter(run => /textrun/i.test(run.text));
    assert.equal(items.length, 3);
    assert.deepEqual(items.map(run => run.line), [0, 2, 4]);
    assert.equal(object.visual.layout.paragraphs.length, 3);
    assert.deepEqual(items.map(run => run.x), [items[0].x, items[0].x, items[0].x]);
    const markers = layout.runs.filter(run => run.text === "•");
    assert.equal(markers.length, 3);
    assert.deepEqual(markers.map(run => run.x), [markers[0].x, markers[0].x, markers[0].x]);
  } finally { document.close(); core.close(); }
});

test("supplied CBAM investment paragraph keeps terminal punctuation on line two", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/cbam-investor-plan.docx", import.meta.url)));
  try {
    const object = document.scene.objects.find(o => o.text?.startsWith("强制合规带来进口商入口"));
    assert.ok(object);
    const context = recordingContext();
    context.measureText = text => ({ width: [...text].length * 16 });
    const layout = layoutTextRuns(context, object.visual.visual.runs, {
      ...object.visual.layout, maxWidth: object.bounds.width,
      align: object.visual.visual.align, lineHeight: object.visual.visual.lineHeight,
    });
    const lines = Array.from({ length: layout.lineCount }, (_, line) =>
      layout.runs.filter(run => run.line === line).map(run => run.text).join(""));
    assert.equal(lines.length, 2, JSON.stringify(lines));
    assert.ok(lines[1].endsWith("基础设施。"));
    assert.equal(lines.join(""), object.text);
    assert.ok(Math.abs(object.bounds.height - 2 * object.visual.visual.lineHeight) < .01);
  } finally { document.close(); core.close(); }
});

test("shared text layout never soft-wraps Chinese closing punctuation to line start", () => {
  for (const punctuation of [..."、。，．！？：；）］】》〉」』〕”’,.;:!?)]}…"]) {
    for (const eastAsianLineBreaks of [true, false]) {
      const run = { text: `中文${punctuation}后`, fontFamily: "sans-serif", fontSize: 12,
        bold: false, italic: false, letterSpacing: 0, eastAsianLineBreaks };
      const layout = layoutTextRuns(recordingContext(), [run], { maxWidth: 10, lineHeight: 18 });
      const lines = Array.from({ length: layout.lineCount }, (_, line) =>
        layout.runs.filter(run => run.line === line).map(run => run.text).join(""));
      assert.deepEqual(lines, [`中文${punctuation}`, "后"], `${punctuation}, ${eastAsianLineBreaks}`);
    }
  }
  const run = { fontFamily: "sans-serif", fontSize: 12, bold: false, italic: false,
    letterSpacing: 0, eastAsianLineBreaks: false };
  for (const [text, expected] of [
    ["中文 。后", ["中文 。", "后"]],
    ["中文\n。后", ["中文", "。后"]],
    ["中文)ABC", ["中文)", "AB", "C"]],
  ]) {
    const layout = layoutTextRuns(recordingContext(), [{ ...run, text }], { maxWidth: 10, lineHeight: 18 });
    const lines = Array.from({ length: layout.lineCount }, (_, line) =>
      layout.runs.filter(run => run.line === line).map(run => run.text).join(""));
    assert.deepEqual(lines, expected, text);
  }
});

test("supplied binary DOC resolves all seven graphics from their own picture records", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-legacy-office.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/complex1-nor.doc", import.meta.url)));
  try {
    const pictures = document.scene.objects.filter(o => o.type === "image"
      && o.source.textRange?.[0] >= 11705 && o.source.textRange[0] < 11737);
    assert.equal(pictures.length, 7, "the authored graphics page has seven pictures");
    // Source PICF geometry checked against Word, with decoded authored payload hashes.
    const expected = [
      [11705, "image/jpeg", 432, 168, 256.3233, 192.1333, "4e3246913ccca32f3bb19adea2c40141f793e591fc51369f718d3fe0168c0815"],
      [11717, "image/png", 288, 240, 56.9893, 76.992, "5ee15462d6fbe05b638dddf5e6114393a024cfa867968f3add316dba52507ea1"],
      [11719, "image/png", 291, 459, 196.248, 70.6453, "b7558d065f7fe67f8000d875c3241e58311bd4dd5b9d09bc28654c8707d8e40b"],
      [11723, "image/png", 73, 241, 187.768, 283.568, "c3c4a50f77bf5f9c965199148704a00f7134228dc78bc3f9b37abe1d5c7da1a6"],
      [11725, "image/png", 418.6667, 686.6666, 277.3333, 273.3333, "60e0a96b819b21a0e957dfa583c76a882180380a1d64eeeca122a9bd1d04a803"],
      [11733, "image/x-wmf", 75, 627, 189, 143.4, "7fb0a8295a682e29a5ef5a72b49f5fc67e2b0643208c08bd7c40d73a99cfa83e"],
      [11735, "image/x-wmf", 528, 504, 144.6667, 106.4747, "28d9a323fd81344ce87178c10225799e109e64202e185485f0853470eeedf750"],
    ];
    for (const [cp, mediaType, x, y, width, height, hash] of expected) {
      const picture = pictures.find(p => p.source.textRange[0] === cp);
      assert.ok(picture, `picture at source position ${cp}`);
      assert.equal(picture.unitIndex, 6);
      assert.equal(picture.visual.mediaType, mediaType);
      assert.equal(createHash("sha256").update(picture.visual.bytes).digest("hex"), hash);
      for (const [key, value] of Object.entries({ x, y, width, height })) {
        assert.ok(Math.abs(picture.bounds[key] - value) < .01, `${cp}: ${key}`);
      }
    }
    const watermark = document.scene.objects.find(o => o.type === "image" && o.source.textRange?.[0] === 14356);
    assert.ok(watermark && watermark.unitIndex === 0);
    assert.ok(Math.abs(watermark.bounds.width - 569.5326) < .01);
    assert.equal(watermark.bounds.y, 672);
  } finally { document.close(); core.close(); }
});

test("supplied binary DOC header picture and framed description share a row above the body", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-legacy-office.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/complex1-nor.doc", import.meta.url)));
  try {
    const graphics = document.scene.objects.find(o => o.text === "Graphics");
    const page = document.scene.objects.filter(o => o.unitIndex === graphics.unitIndex);
    const card = page.find(o => o.type === "image" && o.source.textRange?.[0] === 14657);
    const description = page.find(o => o.text?.startsWith("This is a monochrome OS2 DIB"));
    assert.equal(description.bounds.x, 192);
    assert.equal(description.bounds.width, 288);
    assert.ok(Math.abs(card.bounds.y - description.bounds.y) < 2);
    assert.ok(graphics.bounds.y > card.bounds.y + card.bounds.height);
  } finally { document.close(); core.close(); }
});

test("supplied binary DOC preserves its explicit page breaks around the graphics page", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-legacy-office.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/complex1-nor.doc", import.meta.url)));
  try {
    const graphics = document.scene.objects.find(o => o.text === "Graphics");
    const after = document.scene.objects.find(o => o.text?.startsWith("The previous page contains 7 pictures."));
    assert.equal(graphics.unitIndex, 6, "the graphics occupy Word's seventh page");
    assert.equal(after.unitIndex, 7, "the explanation starts after the authored page break");
  } finally { document.close(); core.close(); }
});

test("supplied binary DOC uses the measured Times New Roman natural line height", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-legacy-office.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/complex1-nor.doc", import.meta.url));
  const original = core.open(bytes);
  let measured;
  try {
    const body = original.scene.objects.find(o => o.text?.startsWith('\t1" top and bottom margins'));
    const metrics = measureSceneFontMetrics([body], {
      resolve: () => ({ family: "Times New Roman", source: "host" }),
    }, {
      createContext: () => ({ font: "", measureText(text) {
        const size = Number(this.font.match(/([\d.]+)px/u)[1]);
        return { width: text.length * size * .45,
          fontBoundingBoxAscent: size * 1825 / 2048, fontBoundingBoxDescent: size * 443 / 2048 };
      } }),
    });
    measured = core.open(bytes, metrics.table);
    const paragraph = measured.scene.objects.find(o => o.text?.startsWith('\t1" top and bottom margins'));
    assert.ok(Math.abs(paragraph.visual.visual.lineHeight - 40 / 3 * 2355 / 2048) < .001,
      `natural spacing must use the shared exact-face metrics: ${paragraph.visual.visual.lineHeight}`);
  } finally { measured?.close(); original.close(); core.close(); }
});

test("supplied binary DOC keeps leader tabs before a separately formatted symbol", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-legacy-office.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/complex1-nor.doc", import.meta.url)));
  try {
    const paragraph = document.scene.objects.find(o => o.text?.startsWith("Bullet (Alt+0183)"));
    assert.ok(paragraph);
    const layout = layoutTextRuns(recordingContext(), paragraph.visual.visual.runs, {
      ...paragraph.visual.layout, maxWidth: paragraph.bounds.width,
    });
    assert.ok(layout.runs.some(run => run.text === "\t" && run.leader === "dot"));
    assert.equal(layout.runs.find(run => run.text === "•")?.x, 192);
  } finally { document.close(); core.close(); }
});

test("supplied Pages missing photo keeps the following column heading below its frame", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/newletter.pages", import.meta.url)));
  const context = recordingContext();
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const column = document.scene.objects.find(o => o.unitIndex === 1 && o.text?.startsWith("Column Title"));
    assert.ok(column);
    const frame = await new SceneRenderer([column], DEFAULT_LIMITS)
      .render(document.scene.info.units[1], { unitIndex: 1 });
    frame.bitmap.close();
    const heading = context.calls.find(c => c[0] === "fillText" && c[1].startsWith("Column"));
    assert.ok(heading && heading[3] > 250,
      `the native column heading follows the missing photo, not the frame top: baseline ${heading?.[3]}`);
    const photo = document.scene.objects.find(o => o.unitIndex === 1 && o.id.endsWith("-31953"));
    assert.equal(photo?.type, "image", "missing bytes must not remove the image geometry");
    assert.equal(photo.visual.kind, "none", "a missing photo leaves empty space, never a preview substitute");
    assert.ok(Math.abs(photo.bounds.height - 107.77959442138672 * 4 / 3) < .01);
    const measured = layoutTextRuns(context, column.visual.visual.runs, {
      ...column.visual.layout, maxWidth: column.bounds.width - column.visual.layout.insetLeft - column.visual.layout.insetRight,
      lineHeight: column.visual.visual.lineHeight,
    });
    assert.ok(Math.abs(column.bounds.y + column.visual.layout.insetTop + measured.runs[0].lineTop
      - (photo.bounds.y + photo.bounds.height + 16)) < .01,
    "the column clears the photo plus the native 12pt wrap margin");
    assert.ok(document.scene.diagnostics.some(d => d.message.includes("IWORK_IMAGE_UNAVAILABLE: image 31953")));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    document.close(); core.close();
  }
});

test("supplied Pages mixed-size title keeps subtitle leading and both paragraph rules", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/newletter.pages", import.meta.url)));
  const context = recordingContext();
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const title = document.scene.objects.find(o => o.unitIndex === 0 && o.text?.startsWith("TODAY’S NEWS"));
    assert.ok(title);
    const [heading, subtitle] = title.visual.layout.paragraphs;
    assert.ok(subtitle.lineHeight < 40 && subtitle.lineHeight < heading.lineHeight,
      `13pt subtitle must not inherit the 68pt heading's ${subtitle.lineHeight}px line height`);
    assert.ok(subtitle.ruleAbove, "the subtitle's top paragraph border must survive parsing");
    assert.ok(subtitle.ruleBelow, "the subtitle's bottom paragraph border must survive parsing");
    assert.equal(subtitle.ruleAbove.strokeWidth, 1, "native Pages uses a 0.75pt border");
    assert.equal(subtitle.ruleAbove.offsetY, -8, "native Pages uses 6pt border spacing");
    assert.equal(subtitle.ruleBelow.offsetY, 8);
    const frame = await new SceneRenderer([title], DEFAULT_LIMITS)
      .render(document.scene.info.units[0], { unitIndex: 0 });
    frame.bitmap.close();
    const rules = context.calls.filter((c, i) => c[0] === "moveTo"
      && context.calls[i + 1]?.[0] === "lineTo" && context.calls[i + 2]?.[0] === "stroke");
    assert.equal(rules.length, 2, "both horizontal borders must actually paint");
    const topic = context.calls.find(c => c[0] === "fillText" && c[1].startsWith("Key"));
    assert.ok(topic && topic[3] > rules[0][2] && topic[3] < rules[1][2],
      "the subtitle must paint between its two borders");
    const sidebar = document.scene.objects.find(o => o.unitIndex === 0 && o.text?.startsWith("Heading 1"));
    assert.ok(rules[1][2] < sidebar.bounds.y, "the subtitle must not overlap the next content row");
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    document.close(); core.close();
  }
});

test("supplied Pages sidebar soft breaks do not duplicate paragraph borders", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/newletter2.pages", import.meta.url)));
  try {
    const sidebar = document.scene.objects.find(o => o.unitIndex === 0 && o.text?.startsWith("Heading 1"));
    const paragraphs = sidebar.visual.layout.paragraphs;
    assert.equal(paragraphs.filter(p => p.ruleBelow).length, 3,
      "native Pages has one separator per sidebar item, not one per forced line break");
    assert.equal(paragraphs[1].spaceAfter, 0, "a soft break must not add paragraph-after spacing");
    assert.equal(paragraphs[1].ruleBelow, undefined);
  } finally { document.close(); core.close(); }
});

test("supplied Pages image captions paint independently of missing image assets", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const context = recordingContext();
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    for (const [file, page, text, count] of [
      ["newletter3", 1, "Tap or click here and start typing to replace this image caption.", 1],
      ["report", 0, "Drag your own photo onto the image placeholder above", 1],
      ["untitled", 1, "Tap or click to replace the caption.", 6],
    ]) {
      const document = core.open(await readFile(new URL(`./fixtures/${file}.pages`, import.meta.url)));
      try {
        const captions = document.scene.objects.filter(o => o.unitIndex === page && o.text?.startsWith(text));
        assert.equal(captions.length, count, `${file} keeps every stored image caption`);
        context.calls.length = 0;
        const frame = await new SceneRenderer(captions, DEFAULT_LIMITS)
          .render(document.scene.info.units[page], { unitIndex: page });
        frame.bitmap.close();
        const painted = context.calls.filter(c => c[0] === "fillText").map(c => c[1]).join("");
        assert.equal(painted.split(text).length - 1, count, `${file} captions must actually paint`);
        if (file === "newletter3") {
          assert.ok(Math.abs(captions[0].bounds.x - 388 * 4 / 3) < .01);
          assert.ok(Math.abs(captions[0].bounds.width - 138 * 4 / 3) < .01,
            "the caption uses the masked image width, not the original image width");
          assert.equal(captions[0].visual.layout.paragraphs[0].align, "center");
        }
      } finally { document.close(); }
    }
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    core.close();
  }
});

test("supplied Pages newsletter headers survive clipping with paragraph leading", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/newletter3.pages", import.meta.url)));
  const headers = document.scene.objects.filter(o => o.unitIndex === 0 && /SEPTEMBER|ISSUE/u.test(o.text ?? ""));
  assert.equal(headers.length, 2);
  const context = recordingContext();
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const frame = await new SceneRenderer(headers, DEFAULT_LIMITS)
      .render(document.scene.info.units[0], { unitIndex: 0 });
    frame.bitmap.close();
    const text = context.calls.filter(c => c[0] === "fillText").map(c => c[1]).join("");
    assert.ok(text.includes("SEPTEMBER 2, 2026"), "line leading must not make a fitting header disappear");
    assert.ok(text.includes("ISSUE 1"));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    document.close(); core.close();
  }
});

test("supplied Pages saved body ranges retain the URL line at the page boundary", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/pages-text-alignment.pages", import.meta.url));
  const initial = core.open(bytes);
  const context = recordingContext();
  context.measureText = function(text) {
    const size = Number(this.font.match(/([\d.]+)px/u)[1]);
    return { width: [...text].reduce((sum, c) => sum + size * (/[\u2e80-\uffff]/u.test(c) ? 1 : .6), 0),
      fontBoundingBoxAscent: size * 1.06, fontBoundingBoxDescent: size * .34 };
  };
  const metrics = measureSceneFontMetrics(initial.scene.objects, {
    resolve: family => ({ family, source: "browser" }),
  }, { createContext: () => context });
  const document = core.open(bytes, metrics.table);
  initial.close();
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: class {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  } });
  try {
    assert.equal(document.scene.info.units.length, 18);
    const url = "https://sxzhdjgl.sxdygbjy.gov.cn:8082/";
    for (const unitIndex of [0, 1]) {
      const objects = document.scene.objects.filter(o => o.unitIndex === unitIndex);
      context.calls.length = 0;
      const frame = await new SceneRenderer(objects, DEFAULT_LIMITS)
        .render(document.scene.info.units[unitIndex], { unitIndex });
      frame.bitmap.close();
      const painted = context.calls.filter(c => c[0] === "fillText").map(c => c[1]).join("");
      if (unitIndex === 0) {
        assert.ok(objects[0].text.endsWith(url), "the saved first-page range contains the URL");
        assert.ok(painted.endsWith(url), "approximate font boxes must not discard the final URL line");
        assert.ok(!painted.includes("login"), "the next page's content stays on the next page");
      } else {
        assert.ok(painted.startsWith("login）和手机APP端。"));
        assert.ok(!painted.includes(url), "the URL must not be repeated on the next page");
      }
    }
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    document.close(); core.close();
  }
});

test("supplied Pages Report retains its three-line drop cap and logical text", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/report.pages", import.meta.url)));
  try {
    const body = document.scene.objects.find(o => o.text?.includes("To get started"));
    assert.ok(body.visual.layout.paragraphs[5].dropCap, "Pages drop-cap metadata reaches the shared layout");
    const context = recordingContext();
    context.measureText = function(text) {
      const size = Number(this.font.match(/([\d.]+)px/u)[1]);
      return { width: [...text].length * size * .5, actualBoundingBoxAscent: size * .7,
        actualBoundingBoxDescent: 0, fontBoundingBoxAscent: size * .9, fontBoundingBoxDescent: size * .2 };
    };
    const result = layoutTextRuns(context, body.visual.visual.runs, {
      ...body.visual.layout, maxWidth: body.bounds.width,
    });
    const cap = result.runs.find(r => r.text === "T" && r.baselineOffset !== undefined);
    assert.ok(cap?.style.bold && cap.style.fontSize > 60, "T is enlarged in its authored bold face");
    const following = result.runs.filter(r => r.line >= cap.line && r.line < cap.line + 3 && r !== cap && r.text.trim());
    assert.ok(following.length > 0);
    assert.ok(following.every(r => r.x >= cap.x + cap.width), "the first three lines clear the initial");
    assert.equal(result.runs.map(r => r.text).join(""), body.text.replaceAll("\n", ""), "visible text retains T exactly once");
    assert.equal(body.text.slice(cap.start, cap.end), "T", "the enlarged initial retains its logical selection range");
    const after = result.runs.find(r => r.line === cap.line + 3 && r.text.trim());
    assert.equal(after.x, 0, "text returns to its authored margin after three lines");
    const short = layoutTextRuns(context, [{ ...body.visual.visual.runs.at(-1), text: "To\nNext" }], {
      maxWidth: 300, paragraphs: [body.visual.layout.paragraphs[5], { ...body.visual.layout.paragraphs[5], dropCap: undefined }],
    });
    assert.equal(short.runs.map(r => r.text).join(""), "ToNext", "a drop cap may share its input run with body text");
    assert.ok(short.runs.find(r => r.text === "Next").lineTop >= 3 * body.visual.layout.paragraphs[5].lineHeight,
      "a short paragraph reserves the full initial before starting the next paragraph");
  } finally { document.close(); core.close(); }
});

test("supplied Pages Report preserves measured spacing, tight heading ascent, and anchored body text", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/report.pages", import.meta.url));
  const initial = core.open(bytes);
  const context = recordingContext();
  context.measureText = function (text) {
    const size = Number(this.font.match(/([\d.]+)px/u)?.[1] ?? 10);
    const heading = this.font.includes("Medium");
    return { width: [...text].length * size * .3,
      fontBoundingBoxAscent: size * (this.textBaseline === "top" ? .08 : heading ? .98 : .95),
      fontBoundingBoxDescent: size * (heading ? .22 : .21) };
  };
  const metrics = measureSceneFontMetrics(initial.scene.objects, {
    resolve: family => ({ family, source: "browser" }),
  }, { createContext: () => context });
  const document = core.open(bytes, metrics.table);
  initial.close();
  const body = document.scene.objects.find(o => o.unitIndex === 1 && o.text?.includes("You can use Pages for both"));
  assert.ok(Math.abs(body.visual.layout.paragraphs[2].lineHeight - 16 * (0.95 + 0.21 + 0.028) * 1.1) < .01,
    "the shared exact-face metrics and Helvetica Neue leading must reach Pages paragraph layout");
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const frame = await new SceneRenderer([body], DEFAULT_LIMITS)
      .render(document.scene.info.units[1], { unitIndex: 1 });
    frame.bitmap.close();
    const heading = context.calls.find(c => c[0] === "fillText" && c[1].includes("Heading"));
    const first = body.visual.layout.paragraphs[0];
    // At this measured width the continuation fits on one line.
    const expected = body.bounds.y + first.lineHeight + first.spaceAfter + 32 * .98;
    assert.ok(Math.abs(heading[3] - expected) < .01, `heading baseline ${heading[3]} != ${expected}`);
    const firstPage = document.scene.objects.find(o => o.unitIndex === 0 && o.text?.includes("It’s easy to edit text"));
    context.calls.length = 0;
    const firstFrame = await new SceneRenderer([firstPage], DEFAULT_LIMITS)
      .render(document.scene.info.units[0], { unitIndex: 0 });
    firstFrame.bitmap.close();
    const painted = context.calls.filter(c => c[0] === "fillText").map(c => c[1]).join("");
    assert.ok(painted.includes("It’s easy to edit text"));
    assert.ok(painted.includes("To add photos"), "text following the anchored image remains visible");
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    document.close(); core.close();
  }
});

test("supplied Pages floating stories use native point-rounded font line boxes", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  try {
    for (const [file, bodyText, index, expected] of [
      ["newletter3", "To get started", 1, 19.2],
      ["newletter2", "This newsletter template uses linked", 2, 22],
      ["newletter4", "can add additional", 0, 20.8],
    ]) {
      const bytes = await readFile(new URL(`./fixtures/${file}.pages`, import.meta.url));
      const initial = core.open(bytes);
      const metrics = measureSceneFontMetrics(initial.scene.objects, { resolve: family => ({ family, source: "browser" }) }, {
        createContext: () => ({ font: "", measureText(text) {
          const size = Number(this.font.match(/([\d.]+)px/u)[1]);
          const [ascent, descent] = this.font.includes("Charter") ? [2007 / 2048, 492 / 2048]
            : this.font.includes("Superclarendon") ? [.98, .225]
            : this.font.includes("HelveticaNeue") ? [.952, .213] : [1878 / 2048, 449 / 2048];
          return { width: text.length * size * .3, fontBoundingBoxAscent: size * ascent, fontBoundingBoxDescent: size * descent };
        } }),
      });
      initial.close();
      const document = core.open(bytes, metrics.table);
      try {
        const object = document.scene.objects.find(o => o.text?.includes(bodyText));
        assert.ok(Math.abs(object.visual.layout.paragraphs[index].lineHeight - expected) < .001,
          `${file} floating line height must match native Pages`);
        if (file === "newletter2") {
          const quote = document.scene.objects.find(o => o.text?.includes("“This is an example"));
          const paragraph = quote.text.split("\n").findIndex(t => t.includes("“This is an example"));
          assert.ok(Math.abs(quote.visual.layout.paragraphs[paragraph].lineHeight - 92 / 3) < .001);
        }
      } finally { document.close(); }
    }
  } finally { core.close(); }
});

// Firefox 154/macOS, Times New Roman at 1000px, ASCII 32..126.
const firefoxTimesAdvances = [250,333,408.20001,500,500,833,777.83331,180.16667,333,333,500,563.96667,250,333,250,277.83334,500,500,500,500,500,500,500,500,500,500,277.83334,277.83334,563.96667,563.96667,563.96667,443.83334,920.90002,722.16669,667,667,722.16669,610.83331,556.16669,722.16669,722.16669,333,389.16666,722.16669,610.83331,889.16669,722.16669,722.16669,556.16669,722.16669,667,556.16669,610.83331,722.16669,722.16669,943.83331,722.16669,722.16669,610.83331,333,277.83334,333,469.23334,500,333,443.83334,500,443.83334,500,443.83334,333,500,500,277.83334,277.83334,500,277.83334,777.83331,500,500,500,500,333,389.16666,277.83334,500,500,722.16669,500,500,443.83334,479.96667,200.2,479.96667,541];

test("original Word character table keeps the B row on page one", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/word-symbol-original.docx", import.meta.url));
  const initial = core.open(bytes);
  const metrics = measureSceneFontMetrics(initial.scene.objects, {
    resolve: family => ({ family, source: "browser" }),
  }, { createContext: () => ({ font: "", measureText(text) {
    const size = Number(this.font.match(/([\d.]+)px/u)[1]);
    return { width: [...text].reduce((sum, char) => sum + (firefoxTimesAdvances[char.codePointAt(0) - 32] ?? 500), 0) * size / 1000,
      fontBoundingBoxAscent: 1825 / 2048 * size, fontBoundingBoxDescent: 443 / 2048 * size };
  } }) });
  initial.close();
  const document = core.open(bytes, metrics.table);
  try {
    for (const text of ["B-", "176", "191"]) {
      const cell = document.scene.objects.find(o => o.text === text);
      assert.ok(cell, `character table contains ${text}`);
      assert.equal(cell.unitIndex, 0, `${text} belongs on the first page, as in Word`);
    }
    assert.equal(document.scene.objects.find(o => o.text === "C-").unitIndex, 1,
      "the next character row still begins on page two");
  } finally { document.close(); core.close(); }
});

test("original Word double table border retains two full strokes and a gap", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-symbol-original.docx", import.meta.url)));
  try {
    const edges = document.scene.objects.filter(o => o.id.startsWith("docx:table:1:") && o.id.includes(":border:"));
    const doubles = edges.filter(o => o.visual.style?.compound === "double");
    const single = edges.find(o => o.id.endsWith("row:0:column:0:fragment:0:border:right"));
    assert.equal(doubles.reduce((sum, edge) => sum + edge.visual.visual.geometry.commands.filter(p => p.kind === "lineTo").length, 0),
      12, "joining corners must retain all twelve outer border segments");
    for (const edge of doubles) {
      assert.ok(Math.abs(edge.visual.visual.strokeWidth - 3 * single.visual.strokeWidth) < .01,
        `double border envelope ${edge.visual.visual.strokeWidth}px must contain two ${single.visual.strokeWidth}px strokes and a gap`);
    }
  } finally { document.close(); core.close(); }
});

test("original Word double border corners use connected miter joins", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-symbol-original.docx", import.meta.url)));
  try {
    const table = document.scene.objects.find(o => o.id === "docx:table:1:fragment:0");
    const edges = document.scene.objects.filter(o => o.id.startsWith("docx:table:1:") && o.visual.style?.compound === "double");
    for (const edge of edges) assert.equal(edge.visual.style.join, "miter");
    for (const x of [table.bounds.x, table.bounds.x + table.bounds.width]) {
      for (const y of [table.bounds.y, table.bounds.y + table.bounds.height]) {
        assert.ok(edges.some(edge => {
          const points = edge.visual.visual.geometry.commands;
          return points.some((p, i) => i > 0 && i < points.length - 1
            && Math.abs(edge.bounds.x + p.x - x) < .01
            && Math.abs(edge.bounds.y + p.y - y) < .01
            && points[i - 1].kind !== "closePath" && points[i + 1].kind !== "moveTo");
        }), `outer corner (${x}, ${y}) must join in one continuous path, not two butt caps`);
      }
    }
  } finally { document.close(); core.close(); }
});

test("original Word matrix parentheses enclose every row", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-symbol-original.docx", import.meta.url)));
  try {
    const equation = document.scene.objects.find(o => o.type === "group" && o.text === "(1 2 3\n4 5 6)(5\n4\n3)");
    assert.ok(equation, "render both matrices from the original EQ field");
    const parts = document.scene.objects.filter(o => o.parentNumericId === equation.numericId);
    const delimiters = parts.filter(o => o.text === "(" || o.text === ")" || o.visual.geometry?.kind === "path")
      .sort((a, b) => a.bounds.x - b.bounds.x);
    assert.equal(delimiters.length, 4);
    for (let i = 0; i < 4; i += 2) {
      const [left, right] = delimiters.slice(i, i + 2);
      const cells = parts.filter(o => /^\d$/u.test(o.text ?? "") && o.bounds.x > left.bounds.x && o.bounds.x < right.bounds.x);
      assert.equal(cells.length, i === 0 ? 6 : 3);
      const top = Math.min(...cells.map(o => o.bounds.y));
      const bottom = Math.max(...cells.map(o => o.bounds.y + o.bounds.height));
      for (const delimiter of [left, right]) {
        assert.ok(delimiter.bounds.y <= top + .01 && delimiter.bounds.y + delimiter.bounds.height >= bottom - .01,
          `${i === 0 ? "two" : "three"}-row matrix: bracket height ${delimiter.bounds.height} must enclose ${bottom - top}px`);
      }
    }
  } finally { document.close(); core.close(); }
});

test("original Word radical clears its radicand and joins its overbar", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-symbol-original.docx", import.meta.url)));
  try {
    const equation = document.scene.objects.find(o => o.type === "group" && o.text?.includes("4ac"));
    const parts = document.scene.objects.filter(o => o.parentNumericId === equation.numericId);
    const radicand = parts.find(o => o.text === "b");
    const legacyGlyph = parts.find(o => o.text === "√");
    if (legacyGlyph) assert.ok(legacyGlyph.bounds.x + legacyGlyph.bounds.width <= radicand.bounds.x,
      `root hook overlaps b by ${legacyGlyph.bounds.x + legacyGlyph.bounds.width - radicand.bounds.x}px`);
    const radical = parts.find(o => o.visual.geometry?.kind === "path");
    assert.ok(radical, "a continuous scalable hook and overbar must enclose the radicand");
    const points = radical.visual.geometry.commands;
    const join = points.at(-2);
    const end = points.at(-1);
    assert.equal(join.y, end.y, "overbar connects directly to the rising stroke");
    assert.ok(radical.bounds.x + join.x + radical.visual.strokeWidth / 2 < radicand.bounds.x,
      "leave space between the rising stroke and b");
    const raised = parts.find(o => o.text === "2 ");
    assert.ok(radical.bounds.y + end.y + radical.visual.strokeWidth / 2 < raised.bounds.y,
      "overbar clears the raised exponent");
    const tail = parts.find(o => o.text === "- 4ac");
    assert.ok(radical.bounds.x + end.x >= tail.bounds.x + tail.bounds.width - .01);
  } finally { document.close(); core.close(); }
});

test("original Word matrix caption paints on its middle-row baseline", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-symbol-original.docx", import.meta.url)));
  try {
    const paragraph = document.scene.objects.find(o => o.text?.includes("3 by 3 matrix"));
    const middle = document.scene.objects.find(o => o.id === `${paragraph.id}:equation:0:part:3`);
    const context = recordingContext();
    const drawn = [];
    context.measureText = function(text) {
      const size = Number(this.font.match(/([\d.]+)px/u)[1]);
      return { width: [...text].length * size * .45,
        fontBoundingBoxAscent: this.textBaseline === "alphabetic" ? size * .82 : 0,
        fontBoundingBoxDescent: size * .18 };
    };
    context.fillText = function(text, x, y) { drawn.push({ text, x, y }); };
    class Canvas {
      getContext() { return context; }
      transferToImageBitmap() { return { close() {} }; }
    }
    Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: Canvas });
    const renderer = new SceneRenderer([paragraph, middle], DEFAULT_LIMITS);
    const frame = await renderer.render(document.scene.info.units[paragraph.unitIndex], { unitIndex: paragraph.unitIndex });
    frame.bitmap.close();
    const caption = drawn.find(r => r.text === "3" && Math.abs(r.x - paragraph.bounds.x - 48) < .01);
    const row = drawn.find(r => r.text === "2" && Math.abs(r.x - middle.bounds.x) < .01);
    assert.ok(caption && row, "paint both the original caption and the matrix middle row");
    assert.ok(Math.abs(caption.y - row.y) < 2,
      `caption baseline ${caption.y} must match matrix middle row ${row.y}`);
    assert.ok(Math.abs(caption.x - paragraph.bounds.x - 48) < .01, "preserve the authored tab");
  } finally {
    document.close(); core.close();
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("original DOCX columns fit their allocated lines with Firefox size-quantized advances", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  // Firefox 154/macOS, Times New Roman, ASCII 32..126, captured from Canvas.
  const large = firefoxTimesAdvances;
  const small = [3.33333,4.46667,5.46667,6.66667,6.66667,11.13333,10.4,2.4,4.46667,4.46667,6.66667,7.53333,3.33333,4.46667,3.33333,3.7,6.66667,6.66667,6.66667,6.66667,6.66667,6.66667,6.66667,6.66667,6.66667,6.66667,3.7,3.7,7.53333,7.53333,7.53333,5.93333,12.3,9.66667,8.9,8.9,9.66667,8.16667,7.43333,9.66667,9.66667,4.46667,5.2,9.66667,8.16667,11.9,9.66667,9.66667,7.43333,9.66667,8.9,7.43333,8.16667,9.66667,9.66667,12.6,9.66667,9.66667,8.16667,4.46667,3.7,4.46667,6.26667,6.66667,4.46667,5.93333,6.66667,5.93333,6.66667,5.93333,4.46667,6.66667,6.66667,3.7,3.7,6.66667,3.7,10.4,6.66667,6.66667,6.66667,6.66667,4.46667,5.2,3.7,6.66667,6.66667,9.66667,6.66667,6.66667,5.93333,6.4,2.66667,6.4,7.23333];
  const actualSize = 13.333333969116211;
  const context = { font: "", letterSpacing: "0px", measureText(text) {
    const size = Number(this.font.match(/([\d.]+)px/u)[1]);
    return { width: [...text].reduce((sum, char) => {
      const index = char.codePointAt(0) - 32;
      return sum + (index < 0 || index >= large.length ? size * .5
        : size === actualSize ? small[index] : large[index] * size / 1000);
    }, 0) };
  } };
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/word-symbol-original.docx", import.meta.url));
  const initial = core.open(bytes);
  let measured;
  try {
    const metrics = measureSceneFontMetrics(initial.scene.objects, { resolve: family => ({ family, source: "browser" }) }, { createContext: () => context });
    measured = core.open(bytes, metrics.table);
    const columns = measured.scene.objects.filter(o => o.source.paragraphId === "9c893ab" && o.text);
    const heading = measured.scene.objects.find(o => o.source.paragraphId === "2b4a7e1" && o.text);
    assert.equal(columns.length, 2);
    for (const column of columns) {
      const { visual: rich, layout: options } = column.visual;
      const painted = layoutTextRuns(context, rich.runs, { ...options, maxWidth: column.bounds.width, lineHeight: rich.lineHeight, align: rich.align });
      assert.ok(painted.lineCount * painted.lineHeight <= column.bounds.height + .01,
        column.id + ": allocated " + column.bounds.height / rich.lineHeight + " lines, painted " + painted.lineCount);
      assert.ok(column.unitIndex < heading.unitIndex || column.bounds.y + painted.lineCount * painted.lineHeight <= heading.bounds.y + .01);
    }
  } finally { measured?.close(); initial.close(); core.close(); }
});
const TEXT_ALIGN_CODES = {
  start: 0,
  center: 1,
  end: 2,
  justify: 3,
  distribute: 4,
  "medium-kashida": 5,
  "high-kashida": 6,
  "low-kashida": 7,
  "thai-distribute": 8,
};

test("supplied Word subsection headings use the authored numbering tab and marker font", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-numbered-body.docx", import.meta.url)));
  try {
    const context = {
      font: "", letterSpacing: "0px",
      measureText(text) {
        const size = Number(this.font.match(/([\d.]+)px/u)[1]);
        return { width: [...text].reduce((sum, c) => sum + size * (/[\u2e80-\uffff]/u.test(c) ? 1 : .5), 0) };
      },
    };
    for (const id of ["3DF73E91", "1B21E377"]) {
      const marker = document.scene.objects.find(o => o.id.startsWith("docx:numbering:") && o.source.paragraphId === id);
      const body = document.scene.objects.find(o => o.id.startsWith("docx:paragraph:") && o.source.paragraphId === id);
      const layout = layoutTextRuns(context, body.visual.visual.runs, {
        ...body.visual.layout, lineHeight: body.visual.visual.lineHeight,
        align: body.visual.visual.align, maxWidth: body.bounds.width,
      });
      const title = layout.runs.find(run => run.text.trim());
      assert.ok(Math.abs(body.bounds.x + title.x - marker.bounds.x - 42) < .01,
        `${body.text}: native Word places both titles at the 630-twip tab, including the leading-tab case`);
      assert.ok(Math.abs(body.visual.layout.hangingIndent - 1247 / 15) < .01,
        "first-line positioning must preserve the continuation indent");
      const style = marker.visual.visual.runs[0];
      assert.equal(style.fontFamily, "Times New Roman");
      assert.equal(style.bold, true, "numbering-level formatting is independent of the first body run");
      const textStyle = body.visual.visual.runs.at(-1);
      assert.equal(textStyle.fontFamily, "SimHei");
      assert.equal(textStyle.bold, false, "retain the authored unbolded SimHei title");
    }
  } finally {
    document.close(); core.close();
  }
});

test("supplied Word URL footnote stays inside its allocated lines without covering the next note", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/word-footnote-url.docx", import.meta.url));
  const initial = core.open(bytes);
  try {
    const url = initial.scene.objects.find(o => o.id.startsWith("docx:footnote:8:paragraph")).visual.visual.runs.at(-1).text;
    // Native Word's complete URL is 401.026008 pt wide. Also exercise a wider fallback.
    for (const [source, advance] of [["browser", 401.026008 * 4 / 3 / url.length / 12], ["fallback", .55]]) {
      const context = {
        font: "", letterSpacing: "0px",
        measureText(text) {
          const size = Number(this.font.match(/([\d.]+)px/u)[1]);
          return { width: [...text].reduce((sum, c) => sum + size * (/[\u2e80-\uffff]/u.test(c) ? 1 : advance), 0) };
        },
      };
      const metrics = measureSceneFontMetrics(initial.scene.objects, {
        resolve: family => ({ family, source }),
      }, { createContext: () => context });
      const document = core.open(bytes, metrics.table);
      try {
        const note = document.scene.objects.find(o => o.id.startsWith("docx:footnote:8:paragraph"));
        const next = document.scene.objects.find(o => o.id.startsWith("docx:footnote:9:paragraph"));
        const layout = layoutTextRuns(context, note.visual.visual.runs, {
          ...note.visual.layout, maxWidth: note.bounds.width, lineHeight: note.visual.visual.lineHeight,
        });
        assert.equal(note.unitIndex, next.unitIndex);
        assert.ok(layout.lineCount * layout.lineHeight <= note.bounds.height + .01, `${source}: painted note exceeds its allocated height`);
        assert.ok(note.bounds.y + layout.lineCount * layout.lineHeight <= next.bounds.y, `${source}: next footnote is covered`);
        assert.equal(layout.runs.map(r => r.text).join(""), note.text);
        if (source === "browser") {
          const lines = Array.from({ length: layout.lineCount }, (_, i) => layout.runs.filter(r => r.line === i).map(r => r.text).join(""));
          assert.ok(lines[0].endsWith("）"), "closing punctuation stays with the Chinese sentence");
          assert.equal(lines[1], url, "the fitting URL occupies one complete line, as in Word");
        }
      } finally { document.close(); }
    }
  } finally { initial.close(); core.close(); }
});

test("supplied Word numbering separates the first line without indenting continuation text", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-numbered-body.docx", import.meta.url)));
  try {
    const context = {
      font: "", letterSpacing: "0px",
      measureText(text) {
        const size = Number(this.font.match(/([\d.]+)px/u)[1]);
        return { width: [...text].reduce((sum, c) => sum + size * (/[\u2e80-\uffff]/u.test(c) ? 1 : 0.5), 0) };
      },
    };
    const render = object => layoutTextRuns(context, object.visual.visual.runs, {
      ...object.visual.layout, lineHeight: object.visual.visual.lineHeight,
      align: object.visual.visual.align, maxWidth: object.bounds.width,
    });
    for (const paragraphId of ["6F330EF3", "28A42473"]) {
      const marker = document.scene.objects.find(o => o.id.startsWith("docx:numbering:") && o.source.paragraphId === paragraphId);
      const body = document.scene.objects.find(o => o.id.startsWith("docx:paragraph:") && o.source.paragraphId === paragraphId);
      assert.ok(body.visual.visual.runs[0].letterSpacing > 3, "original 第4 opening retains its automatic CJK/number gap");
      const run = marker.visual.visual.runs[0];
      assert.equal(run.letterSpacing, 0, "automatic body spacing must not spread list-marker digits");
      assert.ok(Math.abs(render(marker).width - marker.text.length * run.fontSize * .5) < .001,
        "the actual marker paints with ordinary digit advances");
    }
    for (const [id, prefix] of [[2, "1 目前正在"], [3, "2 环境署2022a"]]) {
      const note = document.scene.objects.find(o => o.id.startsWith(`docx:footnote:${id}:`) && o.text);
      assert.ok(note.text.startsWith(prefix), "paragraph tab definitions must not become note text");
      const tabs = note.visual.layout.tabStops;
      assert.ok(tabs.some(tab => Math.abs(tab.position - 1210 / 15) < 0.01), "authored note tab is preserved");
      assert.ok(!tabs.some(tab => Math.abs(tab.position - 1247 / 15) < 0.01), "cleared inherited tab is absent");
      const lines = render(note).runs;
      assert.ok(Math.abs(lines[0].x - 1247 / 15) < 0.01, "native first-line footnote indent");
      if (id === 3) {
        const continuation = lines.find(run => run.line === 1);
        assert.ok(continuation, "the original long note wraps");
        assert.ok(Math.abs(continuation.x - lines[0].x) < 0.01, "note continuation aligns with its first line");
      }
    }
    const abbreviations = document.scene.objects.filter(o => /^[A-Z][A-Za-z0-9]+\t/u.test(o.text ?? ""));
    assert.equal(abbreviations.length, 19, "all original abbreviation rows");
    for (const object of abbreviations) {
      const title = render(object).runs.find(run => /[\u2e80-\uffff]/u.test(run.text));
      assert.ok(Math.abs(title.x - 75.6) < 0.01, `native hanging tab: ${object.text}`);
      const split = object.visual.visual.runs.flatMap(run => [...run.text].map(text => ({ ...run, text })));
      const splitTitle = layoutTextRuns(context, split, { ...object.visual.layout, maxWidth: object.bounds.width }).runs.find(run => /[\u2e80-\uffff]/u.test(run.text));
      assert.equal(splitTitle.x, title.x, "run segmentation must not turn ordinary text into a list marker");
    }
    const heading = document.scene.objects.find(o => o.source.paragraphId === "77EF5E93");
    const headingInk = render(heading).runs.filter(run => run.text.trim().length > 0);
    assert.ok(Math.abs(headingInk[0].x - 6) < 0.01, "manual (a) follows the authored 90-twip tab, not a generated list gap");
    const title = headingInk.find(run => run.text === "渗");
    assert.ok(Math.abs(title.x - 42) < 0.01, "native heading title follows the 630-twip tab");
    const markerEnd = headingInk.filter(run => !/[\u2e80-\uffff]/u.test(run.text)).at(-1);
    assert.ok(markerEnd.x + markerEnd.width < title.x, "manual marker and heading must not overlap");
    for (let index = 0; index < 4; index += 1) {
      const body = document.scene.objects.find(o => o.id === `docx:paragraph:${index}:fragment:0`);
      const marker = document.scene.objects.find(o => o.id === `docx:numbering:${index}`);
      const text = render(body);
      const label = render(marker).runs.at(-1);
      const textX = body.bounds.x + text.runs[0].x;
      assert.ok(Math.abs(marker.bounds.x - 144) < 0.01, "native marker position");
      assert.ok(Math.abs(textX - (102 + 1238 / 15)) < 0.01, `native first-line tab position for ${marker.text}`);
      assert.ok(marker.bounds.x + label.x + label.width < textX, "the painted marker must not overlap text");
      const continuation = text.runs.find(run => run.line === 1);
      if (index < 3) assert.ok(continuation, "the real paragraph wraps");
      if (continuation) assert.ok(Math.abs(body.bounds.x + continuation.x - 144) < 0.01, "continuation retains the authored left indent");
    }
    const markers = document.scene.objects.filter(o => o.id.startsWith("docx:numbering:") && o.visual.visual.align === "start");
    assert.ok(markers.length > 10, "include the original document's small positive and narrow hanging indents");
    for (const marker of markers) {
      const body = document.scene.objects.find(o => o.id === `docx:paragraph:${marker.id.split(":")[2]}:fragment:0`);
      const ink = render(body).runs.find(run => run.text.trim().length > 0);
      const label = render(marker).runs.at(-1);
      assert.ok(marker.bounds.x + label.x + label.width <= body.bounds.x + ink.x + 0.01, `painted overlap: ${marker.text} ${body.text.slice(0,20)}`);
    }
  } finally {
    document.close();
    core.close();
  }
});

test("supplied Word TOC inherits title tabs without moving page numbers onto the next line", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-toc.docx", import.meta.url)));
  try {
    const entries = document.scene.objects.filter((object) => object.text?.includes("\t") && /\d$/u.test(object.text));
    assert.equal(entries.length, 60);
    const context = {
      font: "", letterSpacing: "0px",
      measureText(text) {
        const size = Number(this.font.match(/([\d.]+)px/u)[1]);
        // Native Times New Roman advances keep the Roman labels within their hanging indent.
        const advances = { I: 0.333, V: 0.722, ".": 0.25, " ": 0.25, "…": 1 };
        return { width: [...text].reduce((sum, character) => sum + size * (advances[character] ?? (/[\u2e80-\uffff]/u.test(character) ? 1 : 0.5)), 0) };
      },
    };
    for (const object of entries) {
      const { layout, visual } = object.visual;
      const rendered = layoutTextRuns(context, visual.runs, {
        ...layout, lineHeight: visual.lineHeight, align: visual.align, maxWidth: object.bounds.width,
      });
      if (!object.text.startsWith("Annex")) assert.equal(rendered.lineCount, 1, object.text);
      const page = rendered.runs.at(-1);
      // One source entry uses literal ellipses instead of a page-number tab.
      if (/\t\d+$/u.test(object.text)) {
        assert.ok(Math.abs(page.x + page.width - 632.4) < 0.01, `page number alignment: ${object.text}`);
      }
      if (object.text.startsWith("I. ")) {
        const title = rendered.runs.find((run) => run.text === "导");
        assert.ok(Math.abs(title.x - 1814 / 15) < 0.01, "title follows the inherited left tab");
        assert.equal(rendered.runs.filter((run) => run.leader === "dot").length, 1, "only the page-number tab has a dot leader");
      }
    }
  } finally {
    document.close();
    core.close();
  }
});

test("supplied DOC preserves first-page capacity and justifies only continued fragments", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-legacy-office.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-table-auto-height.doc", import.meta.url)));
  try {
    assert.equal(document.scene.info.units.length, 24, "matches the native Word page count");
    const leading = document.scene.objects.filter((object) => object.unitIndex === 0).at(-1);
    const continuation = document.scene.objects.find((object) => object.unitIndex === 1);
    assert.ok(Math.abs(leading.bounds.height - 124.8) < 0.01, "page one ends with three lines");
    const context = {
      font: "", letterSpacing: "0px",
      measureText(text) {
        const fontSize = Number(this.font.match(/([\d.]+)px/u)[1]);
        return { width: [...text].reduce((sum, character) => sum + fontSize * (/\s/u.test(character) ? 0.3 : /[\x00-\x7f]/u.test(character) ? 0.53 : 1), 0) };
      },
    };
    const render = (object) => layoutTextRuns(context, object.visual.visual.runs, {
      ...object.visual.layout,
      align: object.visual.visual.align,
      lineHeight: object.visual.visual.lineHeight,
      maxWidth: object.bounds.width,
    });
    const leadingLayout = render(leading);
    assert.equal(leadingLayout.lineCount, 3);
    const tail = leadingLayout.runs.filter((run) => run.text.trim() !== "").at(-1);
    assert.ok(Math.abs(tail.x + tail.width - leading.bounds.width) < 0.01, "continued paragraph reaches the right edge");
    assert.equal(leading.visual.layout.continuesAfter, true);
    assert.equal(leading.visual.layout.compressPunctuation, true);
    assert.ok(leading.text.endsWith("第三方提"), "native first-page break");
    assert.equal(leading.visual.visual.runs.map((run) => run.text).join(""), leading.text, "pagination must not invent newline characters");
    const heading = document.scene.objects.find((object) => object.text === "提示条款：");
    const headingTail = render(heading).runs.at(-1);
    assert.ok(headingTail.x + headingTail.width < heading.bounds.width / 2, "a real paragraph end remains unstretched");
    const deletion = document.scene.objects.find((object) => object.text?.startsWith("个人信息删除："));
    const deletionLine = render(deletion).runs.filter((run) => run.line === 0).map((run) => run.text).join("");
    assert.ok(deletionLine.includes("指在"), "NBSP must not split the label from its definition");
    assert.equal(render(continuation).runs[0].x, 0, "continuation starts at the left margin");
  } finally {
    document.close();
    core.close();
  }
});

test("text layout avoids rectangular bands and restores line width below them", () => {
  const run = {fontFamily: "Arial", fontSize: 10, color: 0xff, bold: false, italic: false,
    underline: false, strike: false, baselineShift: 0, letterSpacing: 0, text: "one two three four ".repeat(18)};
  const wrapRegions = [{x: 60, y: 0, width: 40, height: 40},
    {x: 0, y: 20, width: 45, height: 60}, {x: 0, y: 80, width: 100, height: 20}];
  for (const align of ["start", "end", "justify"]) {
    const result = layoutTextRuns({font: "", letterSpacing: "", measureText: text => ({width: text.length * 5})},
      [run], {maxWidth: 100, lineHeight: 10, align, wrapRegions});
    for (const item of result.runs.filter(r => r.text.trim())) {
      for (const b of wrapRegions) assert.ok(item.x + item.width <= b.x + .01 || item.x >= b.x + b.width - .01
        || item.y + item.height <= b.y + .01 || item.y >= b.y + b.height - .01, `${align}: overlapping ${item.text}`);
    }
    assert.ok(result.runs.some(r => r.y >= 100 && r.x + r.width > 60));
    assert.equal(result.runs.map(r => r.text).join(""), run.text);
  }
});

test("supplied Pages newsletter wraps body around the pull quote and staircase", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/newletter4.pages", import.meta.url)));
  try {
    const body = document.scene.objects.find(o => o.unitIndex === 1 && o.text?.startsWith("can add additional"));
    const obstacles = document.scene.objects.filter(o => o.unitIndex === 1
      && (o.text?.includes("pull quote") || o.id.endsWith("-33453")));
    assert.equal(obstacles.length, 2);
    const {layout, visual} = body.visual;
    const context = {font: "", letterSpacing: "0px", measureText(text) {
      return {width: text.length * Number(this.font.match(/([\d.]+)px/u)[1]) * 0.5};
    }};
    const rendered = layoutTextRuns(context, visual.runs, {
      ...layout, lineHeight: visual.lineHeight,
      maxWidth: body.bounds.width - layout.insetLeft - layout.insetRight,
    });
    for (const run of rendered.runs.filter(r => r.text.trim())) {
      const x = body.bounds.x + layout.insetLeft + run.x;
      const y = body.bounds.y + layout.insetTop + run.y;
      for (const obstacle of obstacles) {
        const b = obstacle.bounds;
        assert.ok(x + run.width <= b.x + 0.01 || x >= b.x + b.width - 0.01
          || y + run.height <= b.y + 0.01 || y >= b.y + b.height - 0.01,
        `body ${JSON.stringify(run.text)} at ${x},${y} overlaps ${obstacle.id}`);
      }
    }
    assert.equal(rendered.runs.map(r => r.text).join("").replace(/\s/gu, ""), body.text.replace(/\s/gu, ""));
  } finally { document.close(); core.close(); }
});

test("text layout preserves logical UTF-16 ranges across wrapped placements", () => {
  const runs = [
    {
      text: "甲😀 ", fontFamily: "Arial", fontSize: 12, color: 0xff,
      bold: false, italic: false, underline: false, strikethrough: false,
      highlight: 0, baselineShift: 0, letterSpacing: 0,
    },
    {
      text: "AB乙", fontFamily: "Arial", fontSize: 12, color: 0xff,
      bold: true, italic: false, underline: false, strikethrough: false,
      highlight: 0, baselineShift: 0, letterSpacing: 0,
    },
  ];
  const source = runs.map(({ text }) => text).join("");
  const layout = layoutTextRuns(recordingContext(), runs, {
    maxWidth: 28,
    lineHeight: 18,
    wrap: true,
  });
  const selectable = layout.runs.filter(({ start, end }) => start !== undefined && end !== undefined);
  assert.ok(selectable.length >= 4);
  for (const placement of selectable) {
    assert.equal(source.slice(placement.start, placement.end), placement.text);
  }
  assert.equal(selectable[0].start, 0);
  assert.equal(selectable.at(-1).end, source.length);
  assert.ok(new Set(selectable.map(({ line }) => line)).size > 1);
});

test("continued text frames justify CJK and Latin without adding a line or changing left alignment", () => {
  for (const text of ["甲乙丙丁", "one two"]) {
    const run = {
      text, fontFamily: "Arial", fontSize: 12, color: 0xff,
      bold: false, italic: false, underline: false, strikethrough: false,
      highlight: 0, baselineShift: 0, letterSpacing: 0,
    };
    for (const align of ["justify", "start"]) {
      const options = { maxWidth: 100, lineHeight: 18, align };
      const terminal = layoutTextRuns(recordingContext(), [run], options);
      const continued = layoutTextRuns(recordingContext(), [run], { ...options, continuesAfter: true });
      assert.equal(continued.lineCount, terminal.lineCount);
      assert.equal(continued.height, terminal.height);
      assert.equal(continued.runs.map((placement) => placement.text).join(""), text);
      if (align === "justify") {
        const last = continued.runs.at(-1);
        assert.ok(Math.abs(last.x + last.width - 100) < 0.01);
        assert.ok(terminal.width < 100);
      } else {
        assert.deepEqual(continued, terminal);
      }
    }
  }
});

test("punctuation compression fits the supplied continuation without shrinking Chinese glyphs", () => {
  const text = "中去除个人信息的行为，使其保持不可被检索、访问的状态。";
  const run = { text, fontFamily: "仿宋", fontSize: 64 / 3, color: 255, bold: false,
    italic: false, underline: false, strikethrough: false, highlight: 0, baselineShift: 0, letterSpacing: 0 };
  const context = {font: "", letterSpacing: "0px", measureText: (text) => ({width: [...text].length * run.fontSize})};
  const options = {maxWidth: 553.7333, lineHeight: 41.6, align: "justify"};
  assert.equal(layoutTextRuns(context, [run], options).lineCount, 2);
  const compressed = layoutTextRuns(context, [run], {...options, compressPunctuation: true});
  assert.equal(compressed.lineCount, 1);
  assert.equal(compressed.runs.map((r) => r.text).join(""), text);
  assert.ok(compressed.width <= options.maxWidth + 0.001);
  for (const r of compressed.runs.filter((r) => /\p{Script=Han}/u.test(r.text))) {
    assert.equal(r.width, run.fontSize);
  }
});

test("spreadsheet row auto-fit uses the renderer's authored wrapping and font size", () => {
  const context = {
    font: "",
    letterSpacing: "0px",
    measureText(text) { return { width: text.length * 8 }; },
  };
  const cell = {
    text: "AA AA AA AA",
    width: 28,
    fontRuns: [{
      start: 0,
      end: 11,
      authoredFamily: "Arial",
      renderedFamily: "Arial",
      source: "browser",
      fontSize: 16,
    }],
  };
  assert.ok(sheetAutoFitRowHeight(context, [{ ...cell, wrap: true }], 24) > 24);
  assert.equal(sheetAutoFitRowHeight(context, [{ ...cell, wrap: false }], 24), 24);
});

test("spreadsheet column auto-fit uses the widest displayed line and rendered font metrics", () => {
  const context = {
    font: "",
    letterSpacing: "0px",
    measureText(text) { return { width: text.length * (this.font.includes("700") ? 11 : 8) }; },
  };
  const cells = [{ text: "tiny\nlonger", width: 24, wrap: false }, {
    text: "Wider",
    width: 24,
    wrap: false,
    fontRuns: [{
      start: 0,
      end: 5,
      authoredFamily: "Arial",
      renderedFamily: "Arial",
      source: "browser",
      fontSize: 20,
      bold: true,
    }],
  }];
  assert.equal(sheetAutoFitColumnWidth(context, cells, 12), 60);
});

class Writer {
  bytes = [];

  u8(value) { this.bytes.push(value & 0xff); }
  u16(value) { this.u8(value); this.u8(value >>> 8); }
  u32(value) { this.u16(value); this.u16(value >>> 16); }
  i32(value) { this.u32(value >>> 0); }
  f32(value) {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setFloat32(0, value, true);
    this.bytes.push(...bytes);
  }
  string(value) {
    const bytes = encoder.encode(value);
    this.u32(bytes.length);
    this.bytes.push(...bytes);
  }
  absentString() { this.u32(0xffff_ffff); }
  padding(count) { for (let index = 0; index < count; index += 1) this.u8(0); }
}

function source(writer, shapeId) {
  writer.string("ppt/slides/slide1.xml");
  writer.string("shape");
  writer.u8(1);
  writer.string("shapeId");
  writer.u8(1);
  writer.u32(shapeId);
}

function geometry(writer, value) {
  if (typeof value === "string") {
    writer.u8({ rectangle: 0, ellipse: 1, line: 2 }[value]);
    return;
  }
  if (value.kind === "rounded-rectangle") {
    writer.u8(3);
    writer.f32(value.radiusX);
    writer.f32(value.radiusY);
    return;
  }
  writer.u8(4);
  writer.u8(value.fillRule === "evenodd" ? 1 : 0);
  writer.padding(3);
  writer.u32(value.commands.length);
  for (const command of value.commands) {
    const code = { moveTo: 0, lineTo: 1, quadraticCurveTo: 2, bezierCurveTo: 3, closePath: 4 }[command.kind];
    writer.u8(code);
    writer.padding(3);
    for (const coordinate of command.values ?? []) writer.f32(coordinate);
  }
}

function paint(writer, value) {
  const code = value.kind === "none" ? 0 : value.kind === "solid" ? 1 : value.kind === "linear-gradient" ? 2 : 3;
  writer.u8(code);
  writer.padding(3);
  if (value.kind === "solid") writer.u32(value.color);
  if (value.kind === "linear-gradient") {
    writer.f32(value.x0);
    writer.f32(value.y0);
    writer.f32(value.x1);
    writer.f32(value.y1);
    writer.u32(value.stops.length);
    for (const stop of value.stops) {
      writer.f32(stop.offset);
      writer.u32(stop.color);
    }
  }
  if (value.kind === "radial-gradient") {
    writer.f32(value.x0);
    writer.f32(value.y0);
    writer.f32(value.r0);
    writer.f32(value.x1);
    writer.f32(value.y1);
    writer.f32(value.r1);
    writer.u32(value.stops.length);
    for (const stop of value.stops) {
      writer.f32(stop.offset);
      writer.u32(stop.color);
    }
  }
}

function visual(writer, value, version) {
  if (value.kind === "none") return;
  if (value.kind === "media") {
    writer.u8(value.mediaKind === "video" ? 1 : 0);
    writer.padding(3);
    writer.string(value.mediaType);
    writer.u32(value.bytes.length);
    writer.bytes.push(...value.bytes);
    writer.u8({ none: 0 }[value.visual.kind]);
    writer.padding(3);
    visual(writer, value.visual, version);
    return;
  }
  if (value.kind === "image-color-change") {
    writer.u32(value.from);
    writer.u32(value.to);
    if (version >= 23) {
      writer.u8(value.useAlpha ? 1 : 0);
      writer.u8({ none: 0 }[value.visual.kind]);
      writer.padding(2);
    } else {
      writer.u8({ none: 0 }[value.visual.kind]);
      writer.padding(3);
    }
    visual(writer, value.visual, version);
    return;
  }
  if (value.kind === "layer") {
    for (const component of value.transform) writer.f32(component);
    writer.f32(value.opacity);
    writer.u8({ none: 0, layer: 5, "painted-shape": 6, "rich-text": 7, effect: 8, "text-layout": 9 }[value.visual.kind]);
    writer.padding(3);
    visual(writer, value.visual, version);
    return;
  }
  if (value.kind === "effect") {
    writer.u8(value.rawFlags ?? ((value.shadow === undefined ? 0 : 1) | (value.clip === undefined ? 0 : 2)));
    writer.padding(3);
    if (value.shadow !== undefined) {
      writer.u32(value.shadow.color);
      writer.f32(value.shadow.blur);
      writer.f32(value.shadow.offsetX);
      writer.f32(value.shadow.offsetY);
    }
    if (value.clip !== undefined) geometry(writer, value.clip);
    writer.u8({ none: 0, layer: 5, "painted-shape": 6, "rich-text": 7, effect: 8, "text-layout": 9 }[value.visual.kind]);
    writer.padding(3);
    visual(writer, value.visual, version);
    return;
  }
  if (value.kind === "text-layout") {
    writer.u8({ auto: 0, ltr: 1, rtl: 2 }[value.layout.direction]);
    writer.u8({ horizontal: 0, "vertical-rl": 1, "vertical-lr": 2 }[value.layout.orientation]);
    writer.u8({ none: 0, shrink: 1, "fit-frame": 2 }[value.layout.autoFit]);
    writer.u8(value.rawFlags ?? (
      (value.layout.prefix === undefined ? 0 : 1)
      | (({ top: 0, center: 1, bottom: 2 }[value.layout.verticalAlign ?? "top"]) << 1)
      | (version >= 61 && value.layout.continuesAfter === true ? 8 : 0)
      | (version >= 62 && value.layout.compressPunctuation === true ? 16 : 0)
      | (version >= 64 && value.layout.fixedLineHeight === true ? 32 : 0)
    ));
    writer.f32(value.layout.defaultTabStop);
    writer.f32(value.layout.hangingIndent);
    writer.f32(value.layout.minScale);
    if (version >= 7) writer.f32(value.layout.paragraphSpacing ?? 0);
    if (version >= 8) {
      writer.f32(value.layout.insetLeft ?? 4);
      writer.f32(value.layout.insetRight ?? 4);
      writer.f32(value.layout.insetTop ?? 4);
      writer.f32(value.layout.insetBottom ?? 4);
      writer.f32(value.layout.marginLeft ?? 0);
      writer.f32(value.layout.marginRight ?? 0);
      writer.f32(value.layout.firstLineIndent ?? 0);
    }
    if (version >= 16) {
      writer.u32(value.layout.columnCount ?? 1);
      writer.f32(value.layout.columnSpacing ?? 0);
      writer.f32(value.layout.rotationDegrees ?? 0);
      writer.f32(value.layout.fontScale ?? 1);
      writer.f32(value.layout.lineSpacingReduction ?? 0);
      writer.u8(value.layout.horizontalOverflow === "clip" ? 1 : 0);
      writer.u8({ overflow: 0, clip: 1, ellipsis: 2 }[value.layout.verticalOverflow ?? "overflow"]);
      writer.u8(value.layout.wrap === false ? 0 : 1);
      writer.u8(value.layout.warp === undefined ? 0 : 1);
    }
    if (version >= 32) {
      writer.u8((value.layout.textFill === false ? 0 : 1)
        | (value.layout.textScaleToFit === true ? 2 : 0)
        | (value.layout.lowResolutionSupersample === true ? 4 : 0)
        | (version >= 41 && value.layout.textMatrixScaleToFit === true ? 8 : 0));
      writer.padding(3);
      writer.u32(value.layout.textStrokeColor ?? 0);
      writer.f32(value.layout.textStrokeWidth ?? 0);
      writer.f32(value.layout.textBaseline ?? 0);
      if (version >= 40) paint(writer, value.layout.textPaint ?? { kind: "none" });
      if (version >= 41) paint(writer, value.layout.textStrokePaint ?? { kind: "none" });
    }
    writer.u32(value.layout.tabStops.length);
    for (const item of value.layout.tabStops) {
      const stop = typeof item === "number"
        ? { position: item, align: "start", leader: "none" }
        : item;
      writer.f32(stop.position);
      if (version >= 28) {
        writer.u8({ start: 0, center: 1, end: 2 }[stop.align]);
        writer.u8({ none: 0, dot: 1, hyphen: 2, underscore: 3, "middle-dot": 4 }[stop.leader]);
        writer.padding(2);
      }
    }
    if (value.layout.prefix !== undefined) writer.string(value.layout.prefix);
    if (version >= 16 && value.layout.warp !== undefined) writer.string(value.layout.warp);
    if (version >= 9) {
      writer.u32(value.layout.paragraphs?.length ?? 0);
      for (const paragraph of value.layout.paragraphs ?? []) {
        writer.u8(TEXT_ALIGN_CODES[paragraph.align]);
        writer.padding(3);
        writer.f32(paragraph.marginLeft);
        writer.f32(paragraph.marginRight);
        writer.f32(paragraph.firstLineIndent);
        writer.f32(paragraph.defaultTabStop);
        if (version >= 19) {
          writer.f32(paragraph.lineHeight ?? 0);
          writer.f32(paragraph.spaceBefore ?? 0);
          writer.f32(paragraph.spaceAfter ?? 0);
        }
        if (version >= 24) {
          writer.u8((paragraph.latinLineBreak === false ? 0 : 1)
            | (paragraph.hangingPunctuation === true ? 2 : 0)
            | (version >= 68 && paragraph.dropCap !== undefined ? 16 : 0));
          writer.padding(3);
          if (version >= 65) writer.padding(24);
          if (version >= 66) writer.padding(24);
          if (version >= 68) {
            const cap = paragraph.dropCap;
            writer.u32(cap?.characters ?? 0);
            writer.u32(cap?.lines ?? 0);
            writer.u32(cap?.raisedLines ?? 0);
            writer.f32(cap?.padding ?? 0);
            writer.f32(cap?.outdent ?? 0);
          }
        }
      }
    }
    if (version >= 67) {
      writer.u32(value.layout.wrapRegions?.length ?? 0);
      for (const rect of value.layout.wrapRegions ?? []) {
        for (const value of [rect.x, rect.y, rect.width, rect.height]) writer.f32(value);
      }
    }
    writer.u8({ none: 0, layer: 5, "painted-shape": 6, "rich-text": 7, effect: 8, "text-layout": 9 }[value.visual.kind]);
    writer.padding(3);
    visual(writer, value.visual, version);
    return;
  }
  if (value.kind === "painted-shape") {
    geometry(writer, value.geometry);
    paint(writer, value.fill);
    paint(writer, value.stroke);
    writer.f32(value.strokeWidth);
    return;
  }
  geometry(writer, value.geometry);
  paint(writer, value.fill);
  paint(writer, value.stroke);
  writer.f32(value.strokeWidth);
  writer.u8(TEXT_ALIGN_CODES[value.align]);
  writer.padding(3);
  writer.f32(value.lineHeight);
  writer.u32(value.runs.length);
  for (const run of value.runs) {
    writer.string(run.text);
    writer.string(run.fontFamily);
    writer.f32(run.fontSize);
    writer.u32(run.color);
    writer.u8((run.bold ? 1 : 0)
      | (run.italic ? 2 : 0)
      | (run.underline ? 4 : 0)
      | (run.strikethrough ? 8 : 0));
    writer.padding(3);
    writer.f32(run.letterSpacing);
    writer.u32(run.highlight ?? 0);
    writer.f32(run.baselineShift ?? 0);
    if (version >= 56) writer.f32(run.horizontalScale ?? 1);
  }
}

function snapshot(objects, version = 4) {
  const writer = new Writer();
  writer.u32(0x3144564f);
  writer.u16(version);
  writer.u8(0);
  writer.u8(1);
  writer.u8(1);
  writer.padding(3);
  writer.u32(0);
  writer.u32(1);
  writer.u8(1);
  writer.padding(3);
  writer.u32(0);
  writer.string("unit:0");
  writer.string("Slide 1");
  writer.f32(960);
  writer.f32(720);
  writer.u32(0);
  writer.u32(0);
  if (version >= 5) {
    writer.u32(0);
    writer.u32(0);
    writer.f32(0);
    writer.f32(0);
  }
  if (version >= 27) {
    writer.f32(0);
    writer.u32(0);
    writer.f32(0);
    writer.u32(0);
  }
  if (version >= 55) writer.padding(4);
  if (version >= 11) {
    writer.absentString();
    writer.absentString();
    if (version >= 17) {
      writer.absentString();
      writer.absentString();
    }
    if (version >= 30) writer.u32(0);
    writer.u32(1);
    writer.u8(0);
    writer.padding(3);
  }
  if (version >= 38) {
    writer.u8(0);
    writer.u8(0);
    writer.u8(0);
    writer.padding(1);
    writer.u32(0);
    writer.u32(0);
    writer.u32(0);
    writer.u32(0);
    for (let index = 0; index < 6; index += 1) writer.f32(0);
    for (let index = 0; index < 7; index += 1) writer.absentString();
  }
  if (version >= 37) writer.u32(0);
  if (version >= 6) writer.u32(0);
  writer.u32(objects.length);
  for (const object of objects) {
    writer.u32(object.numericId);
    writer.i32(object.parentNumericId ?? -1);
    writer.u32(0);
    writer.u8(object.type === "group" ? 1 : object.type === "text-box" ? 2 : 5);
    writer.u8({ none: 0, layer: 5, "painted-shape": 6, "rich-text": 7, effect: 8, "text-layout": 9, "image-color-change": 12, media: 13 }[object.visual.kind]);
    writer.u8(0);
    writer.padding(1);
    writer.i32(object.numericId);
    writer.f32(object.x ?? 0);
    writer.f32(object.y ?? 0);
    writer.f32(object.width ?? 100);
    writer.f32(object.height ?? 100);
    writer.string(object.id);
    if (object.parentId === undefined) writer.absentString();
    else writer.string(object.parentId);
    writer.absentString();
    source(writer, object.numericId);
    visual(writer, object.visual, version);
  }
  return Uint8Array.from(writer.bytes);
}

const gradient = {
  kind: "linear-gradient",
  x0: 0,
  y0: 0,
  x1: 80,
  y1: 0,
  stops: [
    { offset: 0, color: 0xff0000ff },
    { offset: 1, color: 0x0000ffff },
  ],
};

test("protocol v41 preserves text stroke paint and matrix scaling intent", () => {
  const scene = decodeSnapshot(snapshot([{
    numericId: 1,
    id: "text:stroke-pattern",
    type: "text-box",
    visual: {
      kind: "text-layout",
      layout: {
        direction: "ltr",
        orientation: "horizontal",
        autoFit: "none",
        verticalAlign: "top",
        defaultTabStop: 36,
        hangingIndent: 0,
        minScale: 0.1,
        tabStops: [],
        textFill: false,
        textScaleToFit: true,
        textMatrixScaleToFit: true,
        textStrokePaint: gradient,
        textStrokeWidth: 2,
        textBaseline: 12,
      },
      visual: { kind: "none" },
    },
  }], 41));

  assert.equal(scene.objects[0].visual.kind, "text-layout");
  assert.equal(scene.objects[0].visual.layout.textMatrixScaleToFit, true);
  assert.deepEqual(scene.objects[0].visual.layout.textStrokePaint, {
    kind: "linear-gradient",
    start: { x: 0, y: 0 },
    end: { x: 80, y: 0 },
    stops: gradient.stops,
  });
});

test("protocol v29 preserves embedded media and its poster", () => {
  const bytes = Uint8Array.of(0x49, 0x44, 0x33);
  const scene = decodeSnapshot(snapshot([{
    numericId: 1,
    id: "media:1",
    type: "shape",
    visual: {
      kind: "media",
      mediaKind: "audio",
      mediaType: "audio/mpeg",
      bytes,
      visual: { kind: "none" },
    },
  }], 29));

  assert.equal(scene.objects[0].visual.kind, "media");
  assert.equal(scene.objects[0].visual.mediaKind, "audio");
  assert.equal(scene.objects[0].visual.mediaType, "audio/mpeg");
  assert.deepEqual(scene.objects[0].visual.bytes, bytes);
  assert.equal(scene.objects[0].visual.visual.kind, "none");
});

test("protocol v4 preserves affine layers, rounded geometry, gradient paint, paths, and rich runs", () => {
  const scene = decodeSnapshot(snapshot([
    {
      numericId: 1,
      id: "group:1",
      type: "group",
      visual: {
        kind: "layer",
        transform: [0, 1, -1, 0, 100, 20],
        opacity: 0.5,
        visual: { kind: "none" },
      },
    },
    {
      numericId: 2,
      id: "shape:2",
      parentNumericId: 1,
      parentId: "group:1",
      type: "shape",
      visual: {
        kind: "painted-shape",
        geometry: { kind: "rounded-rectangle", radiusX: 8, radiusY: 6 },
        fill: gradient,
        stroke: { kind: "solid", color: 0x112233ff },
        strokeWidth: 2,
      },
    },
    {
      numericId: 3,
      id: "text:3",
      type: "text-box",
      visual: {
        kind: "rich-text",
        geometry: {
          kind: "path",
          fillRule: "evenodd",
          commands: [
            { kind: "moveTo", values: [0, 0] },
            { kind: "lineTo", values: [100, 0] },
            { kind: "bezierCurveTo", values: [100, 20, 80, 40, 60, 40] },
            { kind: "closePath" },
          ],
        },
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align: "center",
        lineHeight: 18,
        runs: [
          { text: "Hello ", fontFamily: "Aptos", fontSize: 12, color: 0x000000ff, bold: true, italic: false, underline: false, strikethrough: false, highlight: 0, baselineShift: 0, letterSpacing: 0 },
          { text: "Office", fontFamily: "DengXian", fontSize: 14, color: 0x336699ff, bold: false, italic: true, underline: true, strikethrough: true, highlight: 0xffee88ff, baselineShift: 3, letterSpacing: 0.5 },
        ],
      },
    },
  ]));

  assert.deepEqual(scene.objects[0].visual, {
    kind: "layer",
    transform: { a: 0, b: 1, c: -1, d: 0, e: 100, f: 20 },
    opacity: 0.5,
    visual: { kind: "none" },
  });
  assert.deepEqual(scene.objects[1].visual, {
    kind: "painted-shape",
    geometry: { kind: "rounded-rectangle", radiusX: 8, radiusY: 6 },
    fill: {
      kind: "linear-gradient",
      start: { x: 0, y: 0 },
      end: { x: 80, y: 0 },
      stops: [
        { offset: 0, color: 0xff0000ff },
        { offset: 1, color: 0x0000ffff },
      ],
    },
    stroke: { kind: "solid", color: 0x112233ff },
    strokeWidth: 2,
  });
  assert.equal(scene.objects[2].visual.kind, "rich-text");
  assert.deepEqual(scene.objects[2].visual.runs.map((run) => run.text), ["Hello ", "Office"]);
  assert.deepEqual(scene.objects[2].visual.runs[1], {
    text: "Office",
    fontFamily: "DengXian",
    fontSize: 14,
    color: 0x336699ff,
    bold: false,
    italic: true,
    underline: true,
    strikethrough: true,
    highlight: 0xffee88ff,
    baselineShift: 3,
    letterSpacing: 0.5,
    horizontalScale: 1,
  });
  assert.deepEqual(scene.objects[2].visual.geometry.commands[2], {
    kind: "bezierCurveTo",
    cp1x: 100,
    cp1y: 20,
    cp2x: 80,
    cp2y: 40,
    x: 60,
    y: 40,
  });
});

test("protocol v23 preserves DrawingML color-change useA and defaults v22 payloads to true", () => {
  const object = (useAlpha) => ({
    numericId: 1,
    id: "image:1",
    type: "shape",
    visual: {
      kind: "image-color-change",
      from: 0xffff_ff80,
      to: 0x3366_9980,
      useAlpha,
      visual: { kind: "none" },
    },
  });

  assert.deepEqual(decodeSnapshot(snapshot([object(false)], 23)).objects[0].visual, {
    kind: "image-color-change",
    from: 0xffff_ff80,
    to: 0x3366_9980,
    useAlpha: false,
    visual: { kind: "none" },
  });
  assert.deepEqual(decodeSnapshot(snapshot([object(false)], 22)).objects[0].visual, {
    kind: "image-color-change",
    from: 0xffff_ff80,
    to: 0x3366_9980,
    useAlpha: true,
    visual: { kind: "none" },
  });
});

test("protocol v61 preserves continuation and rejects its flag in older snapshots", () => {
  const visual = {
    kind: "text-layout",
    layout: {
      direction: "ltr", orientation: "horizontal", autoFit: "none",
      tabStops: [], defaultTabStop: 36, hangingIndent: 0, minScale: 0.1,
      continuesAfter: true,
    },
    visual: { kind: "none" },
  };
  const object = { numericId: 1, id: "text:1", type: "text-box", visual };
  assert.equal(decodeSnapshot(snapshot([object], 61)).objects[0].visual.layout.continuesAfter, true);
  assert.equal(decodeSnapshot(snapshot([object], 60)).objects[0].visual.layout.continuesAfter, undefined);
  assert.throws(() => decodeSnapshot(snapshot([{ ...object, visual: { ...visual, rawFlags: 8 } }], 60)));
  const compressed = {...object, visual: {...visual, layout: {...visual.layout, compressPunctuation: true}}};
  assert.equal(decodeSnapshot(snapshot([compressed], 62)).objects[0].visual.layout.compressPunctuation, true);
  assert.equal(decodeSnapshot(snapshot([compressed], 61)).objects[0].visual.layout.compressPunctuation, undefined);
  assert.throws(() => decodeSnapshot(snapshot([{...object, visual: {...visual, rawFlags: 16}}], 61)));
  const fixed = {...object, visual: {...visual, layout: {...visual.layout, fixedLineHeight: true}}};
  assert.equal(decodeSnapshot(snapshot([fixed], 64)).objects[0].visual.layout.fixedLineHeight, true);
  assert.equal(decodeSnapshot(snapshot([fixed], 63)).objects[0].visual.layout.fixedLineHeight, undefined);
  assert.throws(() => decodeSnapshot(snapshot([{...object, visual: {...visual, rawFlags: 32}}], 63)));
  const region = {x: -4, y: 12, width: 30, height: 40};
  const wrapped = regions => ({...object, visual: {...visual, layout: {...visual.layout, wrapRegions: regions}}});
  assert.deepEqual(decodeSnapshot(snapshot([wrapped([region])], 67)).objects[0].visual.layout.wrapRegions, [region]);
  assert.equal(decodeSnapshot(snapshot([wrapped([region])], 66)).objects[0].visual.layout.wrapRegions, undefined);
  for (const regions of [[{...region, x: NaN}], [{...region, width: 0}], Array(1025).fill(region)]) {
    assert.throws(() => decodeSnapshot(snapshot([wrapped(regions)], 67)), /text wrap|text layout/iu);
  }
  const dropCap = { characters: 1, lines: 3, raisedLines: 0, padding: 2, outdent: 0 };
  const capped = cap => ({ ...object, visual: { ...visual, layout: { ...visual.layout, paragraphs: [{
    align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, dropCap: cap,
  }] } } });
  assert.deepEqual(decodeSnapshot(snapshot([capped(dropCap)], 68)).objects[0].visual.layout.paragraphs[0].dropCap, dropCap);
  assert.equal(decodeSnapshot(snapshot([capped(dropCap)], 67)).objects[0].visual.layout.paragraphs[0].dropCap, undefined);
  for (const cap of [{ ...dropCap, lines: 0 }, { ...dropCap, characters: 33 }, { ...dropCap, padding: NaN }]) {
    assert.throws(() => decodeSnapshot(snapshot([capped(cap)], 68)), /text layout/iu);
  }

});

test("protocol v24 preserves DrawingML paragraph line-break properties", () => {
  const paragraph = {
    align: "start",
    marginLeft: 0,
    marginRight: 0,
    firstLineIndent: 0,
    defaultTabStop: 36,
    lineHeight: 18,
    spaceBefore: 2,
    spaceAfter: 3,
    latinLineBreak: false,
    hangingPunctuation: true,
  };
  const visual = {
    kind: "text-layout",
    layout: {
      direction: "ltr",
      orientation: "horizontal",
      autoFit: "none",
      verticalAlign: "top",
      tabStops: [],
      defaultTabStop: 36,
      hangingIndent: 0,
      minScale: 0.1,
      paragraphs: [paragraph],
    },
    visual: { kind: "none" },
  };

  const decoded = decodeSnapshot(snapshot([{
    numericId: 1,
    id: "text:1",
    type: "text-box",
    visual,
  }], 24));

  assert.deepEqual(decoded.objects[0].visual.layout.paragraphs, [paragraph]);
});

test("protocol v5 preserves radial paint, effects, and advanced text layout", () => {
  const visual = {
    kind: "text-layout",
    layout: {
      direction: "rtl",
      orientation: "vertical-rl",
      autoFit: "fit-frame",
      verticalAlign: "center",
      prefix: "1. ",
      tabStops: [
        { position: 24, align: "start", leader: "none" },
        { position: 48, align: "start", leader: "none" },
      ],
      defaultTabStop: 36,
      hangingIndent: 12,
      minScale: 0.5,
    },
    visual: {
      kind: "effect",
      shadow: { color: 0x11223380, blur: 6, offsetX: 3, offsetY: 4 },
      clip: { kind: "rounded-rectangle", radiusX: 8, radiusY: 6 },
      visual: {
        kind: "painted-shape",
        geometry: "ellipse",
        fill: {
          kind: "radial-gradient",
          x0: 20,
          y0: 20,
          r0: 0,
          x1: 20,
          y1: 20,
          r1: 40,
          stops: [{ offset: 0, color: 0xffffffff }, { offset: 1, color: 0x000000ff }],
        },
        stroke: { kind: "none" },
        strokeWidth: 0,
      },
    },
  };

  const scene = decodeSnapshot(snapshot([{
    numericId: 1,
    id: "shape:1",
    type: "shape",
    visual,
  }], 5));

  assert.deepEqual(scene.objects[0].visual, {
    kind: "text-layout",
    layout: {
      direction: "rtl",
      orientation: "vertical-rl",
      autoFit: "fit-frame",
      verticalAlign: "center",
      prefix: "1. ",
      tabStops: [
        { position: 24, align: "start", leader: "none" },
        { position: 48, align: "start", leader: "none" },
      ],
      defaultTabStop: 36,
      hangingIndent: 12,
      paragraphSpacing: 0,
      insetLeft: 4,
      insetRight: 4,
      insetTop: 4,
      insetBottom: 4,
      marginLeft: 0,
      marginRight: 0,
      firstLineIndent: 0,
      columnCount: 1,
      columnSpacing: 0,
      rotationDegrees: 0,
      fontScale: 1,
      lineSpacingReduction: 0,
      horizontalOverflow: "overflow",
      verticalOverflow: "overflow",
      wrap: true,
      minScale: 0.5,
    },
    visual: {
      kind: "effect",
      shadow: { color: 0x11223380, blur: 6, offsetX: 3, offsetY: 4 },
      clip: { kind: "rounded-rectangle", radiusX: 8, radiusY: 6 },
      visual: {
        kind: "painted-shape",
        geometry: "ellipse",
        fill: {
          kind: "radial-gradient",
          start: { x: 20, y: 20, radius: 0 },
          end: { x: 20, y: 20, radius: 40 },
          stops: [{ offset: 0, color: 0xffffffff }, { offset: 1, color: 0x000000ff }],
        },
        stroke: { kind: "none" },
        strokeWidth: 0,
      },
    },
  });
});

test("protocol v28 preserves tab alignment and leaders", () => {
  const tabStops = [{ position: 90, align: "end", leader: "dot" }];
  const decoded = decodeSnapshot(snapshot([{
    numericId: 1,
    id: "text:1",
    type: "text-box",
    visual: {
      kind: "text-layout",
      layout: {
        direction: "ltr",
        orientation: "horizontal",
        autoFit: "none",
        verticalAlign: "top",
        tabStops,
        defaultTabStop: 36,
        hangingIndent: 0,
        minScale: 0.1,
      },
      visual: { kind: "none" },
    },
  }], 28));

  assert.deepEqual(decoded.objects[0].visual.layout.tabStops, tabStops);
});

test("protocol v7 preserves paragraph spacing in advanced text layout", () => {
  const scene = decodeSnapshot(snapshot([{
    numericId: 1,
    id: "shape:paragraph-spacing",
    type: "shape",
    visual: {
      kind: "text-layout",
      layout: {
        direction: "ltr",
        orientation: "horizontal",
        autoFit: "none",
        verticalAlign: "top",
        tabStops: [],
        defaultTabStop: 36,
        hangingIndent: 0,
        minScale: 1,
        paragraphSpacing: 13.5,
      },
      visual: {
        kind: "painted-shape",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
      },
    },
  }], 7));

  assert.equal(scene.objects[0].visual.layout.paragraphSpacing, 13.5);
});

test("protocol v9 preserves paragraph-local text indentation", () => {
  const paragraphs = [
    { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 96 },
    { align: "justify", marginLeft: 36, marginRight: 4, firstLineIndent: -36, defaultTabStop: 96 },
  ];
  const scene = decodeSnapshot(snapshot([{
    numericId: 1,
    id: "shape:paragraph-indentation",
    type: "shape",
    visual: {
      kind: "text-layout",
      layout: {
        direction: "ltr",
        orientation: "horizontal",
        autoFit: "none",
        verticalAlign: "top",
        tabStops: [],
        defaultTabStop: 36,
        hangingIndent: 0,
        paragraphSpacing: 0,
        insetLeft: 4,
        insetRight: 4,
        insetTop: 4,
        insetBottom: 4,
        marginLeft: 0,
        marginRight: 0,
        firstLineIndent: 0,
        paragraphs,
        minScale: 1,
      },
      visual: {
        kind: "painted-shape",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
      },
    },
  }], 9));

  assert.deepEqual(
    scene.objects[0].visual.layout.paragraphs,
    paragraphs.map((paragraph) => ({
      ...paragraph,
      latinLineBreak: false,
      hangingPunctuation: true,
    })),
  );
});

test("protocol v18 preserves extended Word paragraph alignment modes", () => {
  const modes = [
    "justify",
    "distribute",
    "medium-kashida",
    "high-kashida",
    "low-kashida",
    "thai-distribute",
  ];
  for (const align of modes) {
    const scene = decodeSnapshot(snapshot([{
      numericId: 1,
      id: `text:${align}`,
      type: "text-box",
      visual: {
        kind: "rich-text",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align,
        lineHeight: 18,
        runs: [],
      },
    }], 18));
    assert.equal(scene.objects[0].visual.align, align);
  }
});

test("protocol v19 preserves paragraph-local line height and spacing", () => {
  const paragraph = {
    align: "center",
    marginLeft: 12,
    marginRight: 8,
    firstLineIndent: -4,
    defaultTabStop: 36,
    lineHeight: 24,
    spaceBefore: 6,
    spaceAfter: 10,
  };
  const scene = decodeSnapshot(snapshot([{
    numericId: 1,
    id: "text:paragraph-spacing",
    type: "text-box",
    visual: {
      kind: "text-layout",
      layout: {
        direction: "ltr",
        orientation: "horizontal",
        autoFit: "none",
        verticalAlign: "top",
        tabStops: [],
        defaultTabStop: 36,
        hangingIndent: 0,
        paragraphSpacing: 0,
        insetLeft: 4,
        insetRight: 4,
        insetTop: 4,
        insetBottom: 4,
        marginLeft: 0,
        marginRight: 0,
        firstLineIndent: 0,
        paragraphs: [paragraph],
        minScale: 1,
      },
      visual: {
        kind: "painted-shape",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
      },
    },
  }], 19));

  assert.deepEqual(scene.objects[0].visual.layout.paragraphs, [{
    ...paragraph,
    latinLineBreak: false,
    hangingPunctuation: true,
  }]);
});

test("protocol v17 rejects alignment codes introduced by v18", () => {
  assert.throws(
    () => decodeSnapshot(snapshot([{
      numericId: 1,
      id: "text:legacy-distribute",
      type: "text-box",
      visual: {
        kind: "rich-text",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align: "distribute",
        lineHeight: 18,
        runs: [],
      },
    }], 17)),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /text alignment/u.test(error.message),
  );
});

test("protocol v5 rejects reserved effect flags and excessive tab stops", () => {
  const shape = {
    kind: "painted-shape",
    geometry: "rectangle",
    fill: { kind: "solid", color: 0xffffffff },
    stroke: { kind: "none" },
    strokeWidth: 0,
  };
  assert.throws(
    () => decodeSnapshot(snapshot([{
      numericId: 1,
      id: "shape:vertical-align-flags",
      type: "shape",
      visual: {
        kind: "text-layout",
        rawFlags: 0b110,
        layout: {
          direction: "ltr",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          tabStops: [],
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 0.1,
        },
        visual: shape,
      },
    }], 5)),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /text-layout flags/u.test(error.message),
  );

  assert.throws(
    () => decodeSnapshot(snapshot([{
      numericId: 1,
      id: "shape:flags",
      type: "shape",
      visual: {
        kind: "effect",
        rawFlags: 5,
        shadow: { color: 0x00000080, blur: 4, offsetX: 1, offsetY: 1 },
        visual: shape,
      },
    }], 5)),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /effect flags/u.test(error.message),
  );

  assert.throws(
    () => decodeSnapshot(snapshot([{
      numericId: 1,
      id: "shape:tabs",
      type: "shape",
      visual: {
        kind: "text-layout",
        layout: {
          direction: "ltr",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          tabStops: Array.from({ length: 257 }, (_, index) => index + 1),
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 0.1,
        },
        visual: shape,
      },
    }], 5)),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /tab-stop count/u.test(error.message),
  );
});

function recordingContext() {
  const calls = [];
  const gradients = [];
  const filters = [];
  const strokeShadows = [];
  const state = [];
  let activeFilter = "none";
  const context = {
    calls,
    filters,
    gradients,
    strokeShadows,
    font: "",
    fillStyle: "",
    strokeStyle: "",
    lineWidth: 1,
    textBaseline: "alphabetic",
    textAlign: "start",
    letterSpacing: "0px",
    direction: "inherit",
    get filter() { return activeFilter; },
    set filter(value) { activeFilter = value; filters.push(value); },
    globalAlpha: 1,
    globalCompositeOperation: "source-over",
    shadowColor: "rgba(0, 0, 0, 0)",
    shadowBlur: 0,
    shadowOffsetX: 0,
    shadowOffsetY: 0,
    save() {
      state.push({
        globalAlpha: this.globalAlpha,
        globalCompositeOperation: this.globalCompositeOperation,
        shadowColor: this.shadowColor,
        shadowBlur: this.shadowBlur,
        shadowOffsetX: this.shadowOffsetX,
        shadowOffsetY: this.shadowOffsetY,
        direction: this.direction,
        filter: this.filter,
      });
      calls.push(["save"]);
    },
    restore() {
      const saved = state.pop();
      this.globalAlpha = saved?.globalAlpha ?? 1;
      this.globalCompositeOperation = saved?.globalCompositeOperation ?? "source-over";
      this.shadowColor = saved?.shadowColor ?? "rgba(0, 0, 0, 0)";
      this.shadowBlur = saved?.shadowBlur ?? 0;
      this.shadowOffsetX = saved?.shadowOffsetX ?? 0;
      this.shadowOffsetY = saved?.shadowOffsetY ?? 0;
      this.direction = saved?.direction ?? "inherit";
      this.filter = saved?.filter ?? "none";
      calls.push(["restore"]);
    },
    scale(...args) { calls.push(["scale", ...args]); },
    translate(...args) { calls.push(["translate", ...args]); },
    rotate(...args) { calls.push(["rotate", ...args]); },
    transform(...args) { calls.push(["transform", ...args]); },
    setTransform(...args) { calls.push(["setTransform", ...args]); },
    fillRect(...args) { calls.push(["fillRect", ...args, this.globalCompositeOperation]); },
    clearRect(...args) { calls.push(["clearRect", ...args]); },
    beginPath() { calls.push(["beginPath"]); },
    rect(...args) { calls.push(["rect", ...args]); },
    ellipse(...args) { calls.push(["ellipse", ...args]); },
    moveTo(...args) { calls.push(["moveTo", ...args]); },
    lineTo(...args) { calls.push(["lineTo", ...args]); },
    quadraticCurveTo(...args) { calls.push(["quadraticCurveTo", ...args]); },
    bezierCurveTo(...args) { calls.push(["bezierCurveTo", ...args]); },
    closePath() { calls.push(["closePath"]); },
    clip(...args) { calls.push(["clip", ...args]); },
    fill(...args) {
      calls.push(["fill", ...args, {
        shadowColor: this.shadowColor,
        shadowBlur: this.shadowBlur,
        shadowOffsetX: this.shadowOffsetX,
        shadowOffsetY: this.shadowOffsetY,
      }, this.globalAlpha]);
    },
    stroke() {
      strokeShadows.push(this.shadowColor);
      calls.push(["stroke", this.lineWidth]);
    },
    measureText(text) { return { width: text.length * 5 }; },
    fillText(...args) {
      calls.push(["fillText", ...args, this.direction, this.textAlign, this.letterSpacing]);
    },
    strokeText(...args) {
      calls.push(["strokeText", ...args, this.strokeStyle]);
    },
    drawImage(...args) { calls.push(["drawImage", ...args]); },
    createPattern(...args) {
      calls.push(["createPattern", ...args]);
      return { setTransform(...transform) { calls.push(["patternTransform", ...transform]); } };
    },
    createLinearGradient(...args) {
      calls.push(["createLinearGradient", ...args]);
      const stops = [];
      const result = { addColorStop(offset, color) { stops.push([offset, color]); } };
      gradients.push({ kind: "linear", args, stops, result });
      return result;
    },
    createRadialGradient(...args) {
      const stops = [];
      const result = { addColorStop(offset, color) { stops.push([offset, color]); } };
      gradients.push({ kind: "radial", args, stops, result });
      return result;
    },
  };
  return context;
}

test("text layout preserves authored horizontal character scaling", () => {
  const layout = layoutTextRuns(recordingContext(), [{
    text: "x",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
    horizontalScale: 0.33,
  }], { maxWidth: 100 });

  assert.ok(Math.abs((layout.runs[0]?.width ?? 0) - 1.65) < 0.000_001);
});

function sceneObject(overrides) {
  return {
    numericId: overrides.numericId,
    id: overrides.id,
    type: overrides.type ?? "shape",
    unitIndex: 0,
    bounds: overrides.bounds ?? { x: 10, y: 20, width: 80, height: 40 },
    source: {
      format: "pptx",
      part: "ppt/slides/slide1.xml",
      kind: "shape",
      shapeId: overrides.numericId,
      mapping: "exact",
    },
    z: overrides.numericId,
    visual: overrides.visual,
    ...(overrides.parentNumericId === undefined ? {} : { parentNumericId: overrides.parentNumericId }),
    ...(overrides.parentId === undefined ? {} : { parentId: overrides.parentId }),
  };
}

test("supplied Word date paints its final note on the same line with fallback fonts", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-first-page-text.docx", import.meta.url)));
  const object = document.scene.objects.find(o => o.text?.startsWith("2021"));
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  context.measureText = function (text) {
    const size = Number(this.font.match(/([\d.]+)px/u)[1]);
    // Native TNR digits / Apple Symbols fallback star and full-width CJK.
    const advance = [...text].reduce((sum, c) => sum + (/^[0-9]$/u.test(c) ? 0.5 : c === "∗" ? 0.5234375 : 1), 0);
    return { width: advance * size + [...text].length * (parseFloat(this.letterSpacing) || 0) };
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    let nativeInk;
    for (const mode of ["browser", "cjk-fallback", "fallback"]) {
      const resolve = family => ({ family, source: mode === "cjk-fallback" ? (family === "等线" ? "fallback" : "browser") : mode });
      const fonts = { resolve, resolveFace: face => ({ ...face, ...resolve(face.family) }) };
      context.calls.length = 0;
      const frame = await new SceneRenderer([object], DEFAULT_LIMITS, fonts).render(document.scene.info.units[0], { unitIndex: 0 });
      frame.bitmap.close();
      const ink = context.calls.filter(call => call[0] === "fillText" && call[1].trim());
      assert.equal(ink.map(call => call[1]).join(""), object.text, "paint all original text, including the note");
      assert.equal(new Set(ink.map(call => call[3])).size, 1, `${mode}: the star must not wrap`);
      const star = ink.at(-1);
      assert.ok(star[2] > object.bounds.x + 350, `${mode}: note follows the date`);
      const positions = ink.map(call => call.slice(1, 4));
      if (mode === "browser") nativeInk = positions;
      else assert.deepEqual(positions, nativeInk, "equal measured advances must produce equal positions regardless of font source");
    }
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    document.close();
    core.close();
  }
});

test("renderer returns embedded media with authored bounds and transforms", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "media:1",
      visual: {
        kind: "layer",
        transform: { a: 1, b: 0, c: 0, d: 1, e: 12, f: 8 },
        opacity: 1,
        visual: {
          kind: "media",
          mediaKind: "audio",
          mediaType: "audio/mpeg",
          bytes: Uint8Array.of(0x49, 0x44, 0x33),
          visual: { kind: "none" },
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 300, height: 200 },
      { unitIndex: 0 },
    );

    assert.equal(result.renderedObjectCount, 0);
    assert.equal(result.media.length, 1);
    assert.deepEqual(result.media[0], {
      objectId: "media:1",
      kind: "audio",
      mediaType: "audio/mpeg",
      bytes: Uint8Array.of(0x49, 0x44, 0x33),
      bounds: { x: 10, y: 20, width: 80, height: 40 },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 12, f: 8 },
    });
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer covers fractional viewport edge pixels with the requested background", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const renderer = new SceneRenderer([sceneObject({
      visual: {
        kind: "painted-shape",
        geometry: "rectangle",
        fill: { kind: "solid", color: 0xff0000ff },
        stroke: { kind: "none" },
        strokeWidth: 0,
      },
    })], DEFAULT_LIMITS);
    await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 595.2, height: 841.9 },
      { unitIndex: 0, background: "#ffffff" },
    );
    const backgroundIndex = context.calls.findIndex((call) => (
      call[0] === "fillRect" && call[1] === 0 && call[2] === 0
        && call[3] === 596 && call[4] === 842 && call[5] === "destination-over"
    ));
    assert.ok(backgroundIndex > context.calls.findIndex((call) => call[0] === "fill"));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer paints a bounded text watermark without changing document objects", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; }
    getContext() { return context; }
    transferToImageBitmap() { return { width: this.width, height: this.height, close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const renderer = new SceneRenderer([], DEFAULT_LIMITS);
    const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 400, height: 200 };
    const result = await renderer.render(unit, { unitIndex: 0, watermark: "Confidential" });
    assert.equal(result.renderedObjectCount, 0);
    assert.ok(context.calls.some((call) => call[0] === "fillText" && call[1] === "Confidential"));
    await assert.rejects(
      renderer.render(unit, { unitIndex: 0, watermark: " ".repeat(257) }),
      (error) => error?.code === "INVALID_RENDER_REQUEST",
    );
    result.bitmap.close();
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer supersamples low-scale PDF pages before returning requested pixels", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const canvases = [];
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
      this.context = recordingContext();
      canvases.push(this);
    }
    getContext() { return this.context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = {
      ...sceneObject({
        numericId: 1,
        id: "pdf:shape",
        visual: {
          kind: "text-layout",
          layout: { lowResolutionSupersample: true },
          visual: { kind: "none" },
        },
      }),
      source: { format: "pdf", part: "document.pdf", kind: "path", mapping: "exact" },
    };
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 50 },
      { unitIndex: 0, scale: 0.5 },
    );

    assert.deepEqual(canvases.map(({ width, height }) => [width, height]), [[75, 38], [50, 25]]);
    assert.equal(result.pixelWidth, 50);
    assert.equal(result.pixelHeight, 25);
    assert.equal(canvases[1].context.calls.some(([name]) => name === "drawImage"), true);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer applies PPTX text-run shadows and reflections without affecting sibling runs", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  const textEffects = [];
  context.fillText = function fillText(...args) {
    textEffects.push([this.filter, this.globalCompositeOperation, this.fillStyle, ...args]);
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const run = (text) => ({
      text, fontFamily: "Arial", fontSize: 18, color: 0x000000ff,
      bold: false, italic: false, underline: false, strikethrough: false,
      highlight: 0, baselineShift: 0, letterSpacing: 0,
    });
    const renderer = new SceneRenderer([sceneObject({
      numericId: 1,
      id: "pptx:run-shadow",
      type: "text-box",
      visual: {
        kind: "text-effects",
        effects: [
          {
            shadow: { color: 0x0000006e, blur: 5, offsetX: 3, offsetY: -3 },
            innerShadow: { color: 0x00000080, blur: 2, offsetX: 1, offsetY: 1 },
            reflection: {
              startOpacity: 0.5, endOpacity: 0.5, startPosition: 0, endPosition: 1,
              directionDegrees: 90, blur: 0, distance: 2, scaleX: 1, scaleY: 1,
            },
            shadowScaleX: 1,
            shadowScaleY: 0.23,
            shadowSkewX: 20,
            shadowSkewY: 0,
            shadowAlignment: 7,
          },
          {
            shadowScaleX: 1,
            shadowScaleY: 1,
            shadowSkewX: 0,
            shadowSkewY: 0,
            shadowAlignment: 7,
          },
        ],
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "start",
          lineHeight: 22,
          runs: [run("Shadow"), run("Plain")],
        },
      },
    })], DEFAULT_LIMITS);

    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 200, height: 100 },
      { unitIndex: 0 },
    );

    assert.equal(result.renderedObjectCount, 1);
    assert.ok(context.calls.some((call) => call[0] === "transform"
      && call[1] === 1 && call[3] === Math.tan(20 * Math.PI / 180) && call[4] === 0.23));
    assert.ok(textEffects.some(([filter, composite, fillStyle]) => filter === "blur(2.5px)"
      && composite === "source-over" && fillStyle === "rgba(0, 0, 0, 0.43137254901960786)"));
    assert.ok(textEffects.filter(([, , , text]) => text === "Shadow").length > 2);
    assert.equal(textEffects.filter(([, , , text]) => text === "Plain").length, 1);
    assert.ok(!textEffects.some(([filter]) => filter === "blur(1.25px)"));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("scaled-to-fit text and its shadow use the fitted box origin", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  context.measureText = () => ({
    width: 200,
    actualBoundingBoxLeft: 0,
    actualBoundingBoxRight: 160,
    actualBoundingBoxAscent: 30,
    actualBoundingBoxDescent: 10,
  });
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const renderer = new SceneRenderer([{
      ...sceneObject({
        numericId: 1,
        id: "docx:vml-fitpath",
        type: "text-box",
        bounds: { x: 10, y: 20, width: 80, height: 40 },
        visual: {
          kind: "text-effects",
          effects: [{
            shadow: { color: 0x80808080, blur: 0, offsetX: 0, offsetY: 0 },
            shadowScaleX: 1,
            shadowScaleY: 1,
            shadowSkewX: 0,
            shadowSkewY: 0,
            shadowAlignment: 8,
          }],
          visual: {
            kind: "text-layout",
            layout: {
              insetLeft: 0, insetRight: 0, insetTop: 0, insetBottom: 0,
              verticalAlign: "center", wrap: false, textScaleToFit: true,
              warp: "text-path-fit",
            },
            visual: {
              kind: "text", geometry: "rectangle", fill: 0, stroke: 0, strokeWidth: 0,
              fontFamily: "Arial", fontSize: 40, color: 0x000000ff,
              bold: false, italic: false, align: "center", lineHeight: 48,
            },
          },
        },
      }),
      text: "Your Text Here",
      source: { format: "docx", part: "word/document.xml", kind: "shape", mapping: "exact" },
    }], DEFAULT_LIMITS);

    await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 80 },
      { unitIndex: 0 },
    );

    assert.ok(context.calls.some(([name, x, y]) => name === "translate" && x === 90 && y === 60));
    assert.ok(context.calls.some(([name, x]) => name === "translate" && x === 10));
    assert.ok(context.calls.some(([name, x, y]) => name === "scale" && x === 0.5 && y === 1));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer paints zero-width and authored PDF hairlines at one device pixel", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  context.getTransform = () => ({ a: 0, b: 2, c: -2, d: 0, e: 0, f: 0 });
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = {
      ...sceneObject({
        numericId: 1,
        id: "pdf:hairline",
        bounds: { x: 10, y: 20, width: 100, height: 0.01 },
        visual: {
          kind: "painted-shape",
          geometry: {
            kind: "path",
            fillRule: "nonzero",
            commands: [
              { kind: "moveTo", x: 0, y: 0 },
              { kind: "lineTo", x: 100, y: 0 },
            ],
          },
          fill: { kind: "none" },
          stroke: { kind: "solid", color: 0x000000ff },
          strokeWidth: 0,
        },
      }),
      source: { format: "pdf", part: "document.pdf", kind: "path", mapping: "exact" },
    };
    const authoredHairline = {
      ...object,
      numericId: 2,
      id: "pdf:authored-hairline",
      visual: { ...object.visual, strokeWidth: 0.1 },
    };
    const renderer = new SceneRenderer([object, authoredHairline], DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 120, height: 80 },
      { unitIndex: 0, scale: 2 },
    );

    assert.equal(result.renderedObjectCount, 2);
    assert.equal(
      context.calls.filter(([name, width]) => name === "stroke" && width === 0.5).length,
      2,
    );
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("sheet viewport rerenders reuse prepared objects outside the new viewport", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  let offscreenBoundsReads = 0;
  const objects = Array.from({ length: 2_000 }, (_, index) => {
    const object = sceneObject({
      numericId: index + 1,
      id: `cell:${index + 1}`,
      type: "cell",
      bounds: { x: 0, y: index * 20, width: 80, height: 20 },
      visual: { kind: "none" },
    });
    if (index === 0) {
      const bounds = object.bounds;
      Object.defineProperty(object, "bounds", {
        configurable: true,
        get() {
          offscreenBoundsReads += 1;
          return bounds;
        },
      });
    }
    return object;
  });
  const unit = {
    type: "sheet",
    index: 0,
    id: "unit:0",
    name: "Sheet 1",
    width: 80,
    height: 40_000,
    rows: 2_000,
    columns: 1,
    frozenRows: 0,
    frozenColumns: 0,
    frozenWidth: 0,
    frozenHeight: 0,
    rowAxis: { defaultSize: 20, spans: [] },
    columnAxis: { defaultSize: 80, spans: [] },
  };
  try {
    const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
    await renderer.render(unit, {
      unitIndex: 0,
      viewport: { x: 0, y: 0, width: 80, height: 800 },
      sheetSizes: { rows: [], columns: [] },
    });
    offscreenBoundsReads = 0;
    await renderer.render(unit, {
      unitIndex: 0,
      viewport: { x: 0, y: 20_000, width: 80, height: 800 },
      sheetSizes: { rows: [], columns: [] },
    });
    assert.equal(offscreenBoundsReads, 0);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer inherits group transforms and opacity while painting rounded gradients and paths", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const group = sceneObject({
    numericId: 1,
    id: "group:1",
    type: "group",
    bounds: { x: 0, y: 0, width: 200, height: 200 },
    visual: {
      kind: "layer",
      transform: { a: 1, b: 0, c: 0, d: 1, e: 12, f: 8 },
      opacity: 0.5,
      visual: { kind: "none" },
    },
  });
  const rounded = sceneObject({
    numericId: 2,
    id: "shape:2",
    parentNumericId: 1,
    parentId: "group:1",
    visual: {
      kind: "painted-shape",
      geometry: { kind: "rounded-rectangle", radiusX: 8, radiusY: 6 },
      fill: {
        kind: "linear-gradient",
        start: { x: 0, y: 0 },
        end: { x: 80, y: 0 },
        stops: [{ offset: 0, color: 0xff0000ff }, { offset: 1, color: 0x0000ffff }],
      },
      stroke: { kind: "solid", color: 0x112233ff },
      strokeWidth: 2,
    },
  });
  const path = sceneObject({
    numericId: 3,
    id: "shape:3",
    visual: {
      kind: "layer",
      transform: { a: 0, b: 1, c: -1, d: 0, e: 100, f: 0 },
      opacity: 1,
      visual: {
        kind: "painted-shape",
        geometry: {
          kind: "path",
          fillRule: "evenodd",
          commands: [
            { kind: "moveTo", x: 0, y: 0 },
            { kind: "lineTo", x: 20, y: 0 },
            { kind: "quadraticCurveTo", cpx: 30, cpy: 10, x: 20, y: 20 },
            { kind: "bezierCurveTo", cp1x: 10, cp1y: 20, cp2x: 0, cp2y: 10, x: 0, y: 0 },
            { kind: "closePath" },
          ],
        },
        fill: { kind: "solid", color: 0x00ff00ff },
        stroke: { kind: "none" },
        strokeWidth: 0,
      },
    },
  });
  try {
    const renderer = new SceneRenderer([group, rounded, path], DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 300, height: 200 },
      { unitIndex: 0 },
    );

    assert.equal(result.renderedObjectCount, 2);
    assert.equal(context.calls.some((call) => call[0] === "transform" && call.slice(1).join() === "1,0,0,1,12,8"), true);
    assert.equal(context.calls.some((call) => call[0] === "transform" && call.slice(1).join() === "0,1,-1,0,100,0"), true);
    assert.deepEqual(context.gradients[0].args, [10, 20, 90, 20]);
    assert.deepEqual(context.gradients[0].stops, [
      [0, "rgba(255, 0, 0, 1)"],
      [1, "rgba(0, 0, 255, 1)"],
    ]);
    assert.equal(context.calls.some((call) => call[0] === "fill" && call.at(-1) === 0.5), true);
    assert.equal(context.calls.some((call) => call[0] === "bezierCurveTo"), true);
    assert.equal(context.calls.some((call) => call[0] === "quadraticCurveTo"), true);
    assert.equal(context.calls.some((call) => call[0] === "fill" && call[1] === "evenodd"), true);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

// Derived from tdf114848: map its first linear fill through defRPr and append
// a plain paragraph to check that the default paint does not leak between paragraphs.
test("DrawingML gradient mapping survives the PPTX adapter and protocol", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("./fixtures/tdf114848-fill-mapping.pptx", import.meta.url)));
  try {
    const mapped = [];
    const collect = value => {
      if (!value || typeof value !== "object" || value instanceof Uint8Array) return;
      if (value.kind === "mapped-gradient") mapped.push(value);
      for (const child of Object.values(value)) collect(child);
    };
    doc.scene.objects.forEach(collect);
    assert.ok(mapped.some(p => p.tile.right === .5 && p.flip === "flip-x" && p.rotateWithShape === false));
    let text = doc.scene.objects.find(o => o.text?.includes("Plain after gradient")).visual;
    while (text.visual) text = text.visual;
    assert.equal(text.runs.find(r => r.text === "Plain after gradient").paint, undefined, "paragraph fill must not leak into the next paragraph");
  } finally { doc.close(); core.close(); }
});

test("supplied tdf152070 preserves image tile scale and offset", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("./fixtures/tdf152070.pptx", import.meta.url)));
  try {
    const paints = [];
    const collect = visual => { if (visual.fill?.kind === "image") paints.push(visual.fill); if (visual.visual) collect(visual.visual); };
    doc.scene.objects.forEach(o => collect(o.visual));
    assert.equal(paints.length, 1);
    assert.ok(Math.abs(paints[0].mapping?.scaleX - .6) < .001);
    assert.ok(Math.abs(paints[0].mapping?.offsetY - 54) < .001);
  } finally { doc.close(); core.close(); }
});

test("supplied tdf114848 retains shape-following gradient semantics", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("./fixtures/tdf114848.pptx", import.meta.url)));
  try {
    doc.loadUnit(4);
    assert.ok(doc.scene.objects.some(o => o.unitIndex === 4 && JSON.stringify(o.visual).includes('"kind":"shape-gradient"')));
  } finally { doc.close(); core.close(); }
});

test("supplied tdf114848 preserves gradient text paints", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("./fixtures/tdf114848.pptx", import.meta.url)));
  try {
    const object = doc.scene.objects.find(o => o.text?.includes("Word Art Perspective"));
    let visual = object.visual;
    while (visual.visual) visual = visual.visual;
    assert.equal(visual.runs[0].paint?.kind, "linear-gradient");
    assert.ok(visual.runs[0].paint.stops.length >= 2);
  } finally { doc.close(); core.close(); }
});

test("supplied basicspreadsheet preserves authored cell pattern fills", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("./fixtures/basicspreadsheet-fills.xlsx", import.meta.url)));
  try {
    doc.loadUnit(0);
    for (const [address, kind] of [["D1", "linear-gradient"], ["D2", "pattern"], ["A8", "linear-gradient"]]) {
      const cell = doc.scene.objects.find(o => o.source.address === address);
      assert.equal(cell?.visual.visual.fill.kind, kind, `${address} must retain its authored fill`);
    }
  } finally { doc.close(); core.close(); }
});

test("supplied n820786 checker patterns do not fall back to diagonal hatching", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/n820786.pptx", import.meta.url)));
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  let tiles = 0;
  class Canvas {
    constructor(width, height) { if (width === 8 && height === 8) tiles++; }
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: Canvas });
  try {
    const objects = document.scene.objects.filter(o => ["smCheck", "lgCheck"].includes(o.visual.fill?.preset));
    assert.equal(objects.length, 2);
    const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
    (await renderer.render(document.scene.info.units[0], { unitIndex: 0 })).bitmap.close();
    assert.equal(tiles, 2, "both authored checker fills must use the repeating checker tile");
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    document.close(); core.close();
  }
});

test("renderer retraces a curved outline after a pattern fill mutates the canvas path", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const geometry = {
    kind: "path",
    fillRule: "nonzero",
    commands: [
      { kind: "moveTo", x: 20, y: 0 },
      { kind: "lineTo", x: 80, y: 0 },
      { kind: "lineTo", x: 100, y: 20 },
      { kind: "lineTo", x: 100, y: 60 },
      { kind: "lineTo", x: 0, y: 60 },
      { kind: "lineTo", x: 0, y: 20 },
      { kind: "quadraticCurveTo", cpx: 0, cpy: 0, x: 20, y: 0 },
      { kind: "closePath" },
    ],
  };
  const object = sceneObject({
    numericId: 20,
    id: "shape:pattern-with-curved-outline",
    bounds: { x: 10, y: 20, width: 100, height: 60 },
    visual: {
      kind: "stroke-style",
      style: {
        cap: "flat",
        join: "miter",
        compound: "single",
        alignment: "center",
        miterLimit: 4,
        dash: [32, 12],
      },
      visual: {
        kind: "painted-shape",
        geometry,
        fill: {
          kind: "pattern",
          preset: "pct5",
          foreground: 0x4472c4ff,
          background: 0xffff_ffff,
        },
        stroke: { kind: "solid", color: 0x172c51ff },
        strokeWidth: 6,
      },
    },
  });

  try {
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 140, height: 100 },
      { unitIndex: 0 },
    );

    const strokeIndex = context.calls.findLastIndex((call) => call[0] === "stroke");
    const beginPathIndex = context.calls.findLastIndex(
      (call, index) => index < strokeIndex && call[0] === "beginPath",
    );
    const strokedPath = context.calls.slice(beginPathIndex, strokeIndex);
    assert.equal(strokedPath.some((call) => call[0] === "quadraticCurveTo"), true,
      "the outline stroke must use the curved shape path rather than the final pattern dot");
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("DrawingML percentage fills use Office-sized binary stipple cells", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  const tiles = [];
  context.createPattern = () => ({ setTransform() {} });
  class FakeCanvas {
    constructor(width, height) {
      this.context = width === 8 && height === 8 ? recordingContext() : context;
      if (this.context !== context) tiles.push(this.context);
    }
    getContext() { return this.context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const object = sceneObject({
    numericId: 21,
    id: "shape:pct90",
    bounds: { x: 0, y: 0, width: 80, height: 80 },
    visual: {
      kind: "painted-shape",
      geometry: "rectangle",
      fill: {
        kind: "pattern",
        preset: "pct90",
        foreground: 0x000000ff,
        background: 0xffffffff,
      },
      stroke: { kind: "none" },
      strokeWidth: 0,
    },
  });

  try {
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 80, height: 80 },
      { unitIndex: 0 },
    );

    const cells = tiles.flatMap((tile) => tile.calls.filter((call) => call[0] === "fillRect" && call[3] === 1 && call[4] === 1));
    assert.equal(cells.length, 62, "pct90 must retain the native named hatch bitmap, rather than a computed percentage");
    assert.equal(cells.every((call) => call[3] === 1 && call[4] === 1), true);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer preserves authored subpixel PDF pattern stroke coverage", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  context.stroke = function stroke() {
    this.calls.push(["stroke", this.lineWidth, this.globalAlpha]);
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const object = sceneObject({
    numericId: 21,
    id: "shape:pdf-subpixel-pattern",
    visual: {
      kind: "painted-shape",
      geometry: "rectangle",
      fill: {
        kind: "pattern",
        preset: "pdf:down:6:0.4",
        foreground: 0x0000_00ff,
        background: 0,
      },
      stroke: { kind: "none" },
      strokeWidth: 0,
    },
  });

  try {
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 140, height: 100 },
      { unitIndex: 0 },
    );

    assert.equal(
      context.calls.some((call) => call[0] === "stroke" && call[1] === 0.4 && call[2] === 1),
      true,
    );
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

for (const compound of ["double", "triple"]) test(`renderer applies ${compound} stroke styles to legacy shapes`, async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  context.setLineDash = (dash) => context.calls.push(["setLineDash", dash]);
  context.stroke = function stroke() {
    this.calls.push(["stroke", this.lineWidth, this.strokeStyle]);
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });

  try {
    const object = sceneObject({
      numericId: 21,
      id: "legacy-shape:dashed",
      bounds: { x: 10, y: 20, width: 100, height: 60 },
      visual: {
        kind: "stroke-style",
        style: {
          cap: "flat",
          join: "miter",
          compound,
          alignment: "center",
          miterLimit: 4,
          dash: [12, 4],
        },
        visual: {
          kind: "shape",
          geometry: "rectangle",
          fill: 0,
          stroke: 0x172c51ff,
          strokeWidth: 4,
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 140, height: 100 },
      { unitIndex: 0 },
    );

    assert.deepEqual(
      context.calls.find((call) => call[0] === "setLineDash"),
      ["setLineDash", [12, 4]],
    );
    assert.deepEqual(
      context.calls.filter((call) => call[0] === "stroke"),
      compound === "double" ? [
        ["stroke", 4, "rgba(23, 44, 81, 1)"],
        ["stroke", 4 / 3, "rgba(255, 255, 255, 1)"],
      ] : [
        ["stroke", 7.2, "rgba(23, 44, 81, 1)"],
        ["stroke", 4.8, "rgba(255, 255, 255, 1)"],
        ["stroke", 3.2, "rgba(23, 44, 81, 1)"],
        ["stroke", 1.6, "rgba(255, 255, 255, 1)"],
        ["stroke", 0.8, "rgba(23, 44, 81, 1)"],
      ],
    );
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer snaps only subpixel axis-aligned hairlines in device space", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  const multiply = (left, right) => ({
    a: left.a * right.a + left.c * right.b,
    b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d,
    d: left.b * right.c + left.d * right.d,
    e: left.a * right.e + left.c * right.f + left.e,
    f: left.b * right.e + left.d * right.f + left.f,
  });
  let transform = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const transforms = [];
  const save = context.save.bind(context);
  const restore = context.restore.bind(context);
  const scale = context.scale.bind(context);
  const translate = context.translate.bind(context);
  const applyTransform = context.transform.bind(context);
  context.save = function saveWithTransform() {
    transforms.push({ ...transform });
    save();
  };
  context.restore = function restoreWithTransform() {
    restore();
    transform = transforms.pop() ?? { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  };
  context.scale = function scaleWithTransform(x, y) {
    scale(x, y);
    transform = multiply(transform, { a: x, b: 0, c: 0, d: y, e: 0, f: 0 });
  };
  context.translate = function translateWithTransform(x, y) {
    translate(x, y);
    transform = multiply(transform, { a: 1, b: 0, c: 0, d: 1, e: x, f: y });
  };
  context.transform = function transformWithTransform(a, b, c, d, e, f) {
    applyTransform(a, b, c, d, e, f);
    transform = multiply(transform, { a, b, c, d, e, f });
  };
  context.getTransform = () => ({ ...transform });
  context.setLineDash = (...dash) => context.calls.push(["setLineDash", ...dash]);
  let lineWidth = context.lineWidth;
  Object.defineProperty(context, "lineWidth", {
    configurable: true,
    get: () => lineWidth,
    set(value) {
      lineWidth = value;
      context.calls.push(["lineWidth", value]);
    },
  });

  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const solidStroke = { kind: "solid", color: 0x000080ff };
  const noFill = { kind: "none" };
  try {
    const objects = [
      sceneObject({
        numericId: 1,
        id: "line:native-device-width",
        bounds: { x: 10, y: 20.1, width: 40, height: 0 },
        visual: {
          kind: "layer",
          transform: { a: 1, b: 0, c: 0, d: 1, e: 0.2, f: 0.2 },
          opacity: 1,
          visual: {
            kind: "painted-shape",
            geometry: "line",
            fill: noFill,
            stroke: solidStroke,
            strokeWidth: 0.5,
          },
        },
      }),
      sceneObject({
        numericId: 2,
        id: "path:dashed-arrow",
        bounds: { x: 40, y: 30.1, width: 20, height: 4 },
        visual: {
          kind: "stroke-style",
          style: {
            cap: "flat",
            join: "miter",
            compound: "single",
            alignment: "center",
            miterLimit: 4,
            dash: [4, 2],
          },
          visual: {
            kind: "painted-shape",
            geometry: {
              kind: "path",
              fillRule: "nonzero",
              commands: [
                { kind: "moveTo", x: 0, y: 0 },
                { kind: "lineTo", x: 20, y: 0 },
                { kind: "moveTo", x: 20, y: 0 },
                { kind: "lineTo", x: 17, y: -2 },
                { kind: "lineTo", x: 17, y: 2 },
                { kind: "closePath" },
              ],
            },
            fill: noFill,
            stroke: solidStroke,
            strokeWidth: 0.5,
          },
        },
      }),
      sceneObject({
        numericId: 3,
        id: "rectangle:multi-device-width",
        bounds: { x: 70.2, y: 40.2, width: 10, height: 10 },
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: noFill,
          stroke: solidStroke,
          strokeWidth: 1,
        },
      }),
      sceneObject({
        numericId: 4,
        id: "line:diagonal",
        bounds: { x: 90.1, y: 55.1, width: 10, height: 5 },
        visual: {
          kind: "painted-shape",
          geometry: "line",
          fill: noFill,
          stroke: solidStroke,
          strokeWidth: 0.5,
        },
      }),
      sceneObject({
        numericId: 5,
        id: "path:curve",
        bounds: { x: 105, y: 60, width: 10, height: 5 },
        visual: {
          kind: "painted-shape",
          geometry: {
            kind: "path",
            fillRule: "nonzero",
            commands: [
              { kind: "moveTo", x: 0, y: 0 },
              { kind: "quadraticCurveTo", cpx: 5, cpy: 5, x: 10, y: 0 },
            ],
          },
          fill: noFill,
          stroke: solidStroke,
          strokeWidth: 0.5,
        },
      }),
      sceneObject({
        numericId: 6,
        id: "line:axis-swapped-transform",
        bounds: { x: 5, y: 5.1, width: 10, height: 0 },
        visual: {
          kind: "layer",
          transform: { a: 0, b: 1, c: -1, d: 0, e: 130, f: 0 },
          opacity: 1,
          visual: {
            kind: "painted-shape",
            geometry: "line",
            fill: noFill,
            stroke: solidStroke,
            strokeWidth: 0.5,
          },
        },
      }),
      sceneObject({
        numericId: 7,
        id: "line:office-hairline",
        bounds: { x: 10, y: 70.1, width: 20, height: 0 },
        visual: {
          kind: "layer",
          transform: { a: 0.5, b: 0, c: 0, d: 0.5, e: 0, f: 0 },
          opacity: 1,
          visual: {
            kind: "painted-shape",
            geometry: "line",
            fill: noFill,
            stroke: solidStroke,
            strokeWidth: 2 / 3,
          },
        },
      }),
      sceneObject({
        numericId: 8,
        id: "line:scaled-beyond-hairline",
        bounds: { x: 10, y: 75.1, width: 20, height: 0 },
        visual: {
          kind: "painted-shape",
          geometry: "line",
          fill: noFill,
          stroke: solidStroke,
          strokeWidth: 2 / 3,
        },
      }),
    ];
    const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 140, height: 80 },
      { unitIndex: 0, pixelRatio: 2 },
    );

    assert.equal(context.calls.filter((call) => (
      call[0] === "moveTo" && call[1] === 10 && call[2] === 20.1
    )).length, 1, "a native one-device-pixel stroke should retain its authored center");
    assert.equal(context.calls.some((call) => (
      call[0] === "moveTo" && call[1] === 40 && call[2] === 30.1
    )), true, "the native one-pixel dashed path should retain its tangent origin");
    assert.equal(context.calls.some((call) => (
      call[0] === "lineTo" && call[1] === 57 && call[2] === 30.1
    )), true, "the stroked path should still terminate at the arrow attachment");
    assert.equal(context.calls.some((call) => (
      call[0] === "moveTo" && call[1] === 60 && call[2] === 30.1
    )), true, "the separately filled arrowhead should retain its authored coordinates");
    assert.deepEqual(context.calls.find((call) => call[0] === "setLineDash"), ["setLineDash", [4, 2]]);
    assert.equal(context.calls.some((call) => (
      call[0] === "rect" && call[1] === 70.2 && call[2] === 40.2
    )), true, "multi-pixel rectangle strokes should retain their authored geometry");
    assert.equal(context.calls.filter((call) => (
      call[0] === "moveTo" && call[1] === 90.1 && call[2] === 55.1
    )).length, 1, "diagonal strokes should retain their authored path");
    assert.equal(context.calls.filter((call) => call[0] === "quadraticCurveTo").length, 1,
      "curved strokes should not be rebuilt by pixel snapping");
    assert.equal(context.calls.filter((call) => (
      call[0] === "moveTo" && call[1] === 5 && call[2] === 5.1
    )).length, 1, "axis-swapped native one-pixel strokes should retain their authored center");
    assert.equal(context.calls.some((call) => (
      call[0] === "moveTo" && call[1] === 10 && Math.abs(call[2] - 70.5) < 1e-9
    )), true, "a subpixel Office hairline should snap to a one-device-pixel center");
    const lineWidths = context.calls
      .filter((call) => call[0] === "lineWidth")
      .map((call) => call[1]);
    assert.deepEqual(lineWidths.slice(-3), [2 / 3, 1, 2 / 3],
      "only the <=1 device-pixel stroke should be quantized to one pixel");
    assert.equal(context.calls.filter((call) => (
      call[0] === "moveTo" && call[1] === 10 && call[2] === 75.1
    )).length, 1, "the same hairline scaled above one device pixel should retain its authored width and path");
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer applies object shadows and geometry clips to radial paint", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "shape:1",
      visual: {
        kind: "effect",
        shadow: { color: 0x11223380, blur: 6, offsetX: 3, offsetY: 4 },
        clip: { kind: "rounded-rectangle", radiusX: 8, radiusY: 6 },
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: {
            kind: "radial-gradient",
            start: { x: 20, y: 20, radius: 0 },
            end: { x: 20, y: 20, radius: 40 },
            stops: [{ offset: 0, color: 0xffffffff }, { offset: 1, color: 0x000000ff }],
          },
          stroke: { kind: "solid", color: 0x000000ff },
          strokeWidth: 1,
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 80 },
      { unitIndex: 0 },
    );

    assert.equal(result.renderedObjectCount, 1);
    assert.deepEqual(context.gradients[0].args, [30, 40, 0, 30, 40, 40]);
    assert.deepEqual(context.gradients[0].stops, [
      [0, "rgba(255, 255, 255, 1)"],
      [1, "rgba(0, 0, 0, 1)"],
    ]);
    assert.equal(context.calls.some((call) => call[0] === "clip"), true);
    const painted = context.calls.find((call) => call[0] === "fill" && call[1] === "nonzero");
    assert.deepEqual(painted?.at(-2), {
      shadowColor: "rgba(17, 34, 51, 0.5019607843137255)",
      shadowBlur: 6,
      shadowOffsetX: 3,
      shadowOffsetY: 4,
    });
    assert.equal(context.strokeShadows.at(-1), "transparent");
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer retains a shared path clip across consecutive objects", async () => {
  const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const originalPath = Object.getOwnPropertyDescriptor(globalThis, "Path2D");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  class FakePath2D {
    moveTo() {}
    lineTo() {}
    closePath() {}
  }
  Object.defineProperties(globalThis, {
    OffscreenCanvas: { configurable: true, value: FakeCanvas },
    Path2D: { configurable: true, value: FakePath2D },
  });
  try {
    const clip = {
      kind: "path",
      fillRule: "nonzero",
      commands: [
        { kind: "moveTo", x: 0, y: 0 },
        { kind: "lineTo", x: 80, y: 0 },
        { kind: "lineTo", x: 80, y: 40 },
        { kind: "closePath" },
      ],
    };
    const objects = [1, 2].map((numericId) => sceneObject({
      numericId,
      id: `shape:${numericId}`,
      visual: {
        kind: "effect",
        clip,
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0xffffffff },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      },
    }));
    const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 80 },
      { unitIndex: 0 },
    );

    assert.equal(result.renderedObjectCount, 2);
    assert.equal(context.calls.filter(([name]) => name === "clip").length, 1);
  } finally {
    if (originalCanvas) Object.defineProperty(globalThis, "OffscreenCanvas", originalCanvas);
    else delete globalThis.OffscreenCanvas;
    if (originalPath) Object.defineProperty(globalThis, "Path2D", originalPath);
    else delete globalThis.Path2D;
  }
});

test("renderer batches consecutive opaque PDF paths with the same paint", async () => {
  const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const originalPath = Object.getOwnPropertyDescriptor(globalThis, "Path2D");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  class FakePath2D {
    moveTo() {}
    lineTo() {}
    quadraticCurveTo() {}
    bezierCurveTo() {}
    closePath() {}
  }
  Object.defineProperties(globalThis, {
    OffscreenCanvas: { configurable: true, value: FakeCanvas },
    Path2D: { configurable: true, value: FakePath2D },
  });
  try {
    const path = {
      kind: "path",
      fillRule: "nonzero",
      commands: [
        { kind: "moveTo", x: 0, y: 0 },
        { kind: "lineTo", x: 10, y: 0 },
        { kind: "lineTo", x: 10, y: 10 },
        { kind: "closePath" },
      ],
    };
    const object = (numericId, color) => ({
      ...sceneObject({
        numericId,
        id: `pdf:${numericId}`,
        bounds: { x: numericId * 12, y: 0, width: 10, height: 10 },
        visual: {
          kind: "painted-shape",
          geometry: path,
          fill: { kind: "solid", color },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      }),
      source: { format: "pdf", part: "document.pdf", kind: "path", mapping: "exact" },
    });
    const renderer = new SceneRenderer([
      object(1, 0x112233ff),
      object(2, 0x112233ff),
      object(3, 0x11223380),
    ], DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 120, height: 80 },
      { unitIndex: 0 },
    );

    assert.equal(result.renderedObjectCount, 3);
    assert.equal(context.calls.filter(([name]) => name === "fill").length, 2);

    const overlapping = { ...object(2, 0x112233ff), bounds: object(1, 0).bounds };
    const fills = context.calls.filter(([name]) => name === "fill").length;
    const overlapResult = await new SceneRenderer([
      object(1, 0x112233ff),
      overlapping,
    ], DEFAULT_LIMITS).render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 120, height: 80 },
      { unitIndex: 0 },
    );
    assert.equal(overlapResult.renderedObjectCount, 2);
    assert.equal(context.calls.filter(([name]) => name === "fill").length - fills, 2);
  } finally {
    if (originalCanvas) Object.defineProperty(globalThis, "OffscreenCanvas", originalCanvas);
    else delete globalThis.OffscreenCanvas;
    if (originalPath) Object.defineProperty(globalThis, "Path2D", originalPath);
    else delete globalThis.Path2D;
  }
});

test("renderer reuses a bounded final raster for identical requests", async () => {
  const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const originalBitmap = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const context = recordingContext();
  let canvases = 0;
  let clones = 0;
  let cachedCloses = 0;
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
      canvases += 1;
    }
    getContext() { return context; }
    transferToImageBitmap() {
      return { width: this.width, height: this.height, close() { cachedCloses += 1; } };
    }
  }
  Object.defineProperties(globalThis, {
    OffscreenCanvas: { configurable: true, value: FakeCanvas },
    createImageBitmap: {
      configurable: true,
      value: async (source) => {
        clones += 1;
        return { width: source.width, height: source.height, close() {} };
      },
    },
  });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "shape:1",
      visual: {
        kind: "painted-shape",
        geometry: "rectangle",
        fill: { kind: "solid", color: 0xffffffff },
        stroke: { kind: "none" },
        strokeWidth: 0,
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const unit = { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 80 };
    const first = await renderer.render(unit, { unitIndex: 0 });
    const second = await renderer.render(unit, { unitIndex: 0 });
    assert.equal(canvases, 1);
    assert.equal(context.calls.filter(([name]) => name === "fill").length, 1);
    assert.equal(clones, 2);
    first.bitmap.close();
    second.bitmap.close();

    const different = await renderer.render(unit, { unitIndex: 0, background: "#000000" });
    assert.equal(canvases, 2);
    different.bitmap.close();
    renderer.appendObjects([{ ...object, numericId: 2, id: "shape:2", z: 2 }]);
    assert.equal(cachedCloses, 2);
    const invalidated = await renderer.render(unit, { unitIndex: 0 });
    assert.equal(canvases, 3);
    invalidated.bitmap.close();
    renderer.close();

    const limited = new SceneRenderer(
      [object],
      { ...DEFAULT_LIMITS, imagePixels: 100, totalImagePixels: 100 },
    );
    (await limited.render(unit, { unitIndex: 0 })).bitmap.close();
    (await limited.render(unit, { unitIndex: 0 })).bitmap.close();
    assert.equal(canvases, 5);
    assert.equal(clones, 4);
    limited.close();
  } finally {
    if (originalCanvas) Object.defineProperty(globalThis, "OffscreenCanvas", originalCanvas);
    else delete globalThis.OffscreenCanvas;
    if (originalBitmap) Object.defineProperty(globalThis, "createImageBitmap", originalBitmap);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer skips fully transparent shapes before Canvas drawing", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const objects = [
      sceneObject({
        numericId: 1,
        id: "shape:transparent",
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0xffffff00 },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      }),
      sceneObject({
        numericId: 2,
        id: "shape:opaque",
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0xffffffff },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      }),
    ];
    const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 80 },
      { unitIndex: 0 },
    );

    assert.equal(result.renderedObjectCount, 1);
    assert.equal(context.calls.filter(([name]) => name === "fill").length, 1);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer skips shapes strictly covered by a later opaque rectangle", async () => {
  const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const rectangle = {
      kind: "painted-shape",
      geometry: "rectangle",
      fill: { kind: "solid", color: 0xffffffff },
      stroke: { kind: "none" },
      strokeWidth: 0,
    };
    const objects = [
      sceneObject({
        numericId: 1,
        id: "shape:covered",
        bounds: { x: 10, y: 10, width: 10, height: 10 },
        visual: rectangle,
      }),
      sceneObject({
        numericId: 2,
        id: "shape:cover",
        bounds: { x: 0, y: 0, width: 40, height: 40 },
        visual: rectangle,
      }),
    ];
    const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 80 },
      { unitIndex: 0 },
    );

    assert.equal(result.renderedObjectCount, 1);
  } finally {
    if (originalCanvas) Object.defineProperty(globalThis, "OffscreenCanvas", originalCanvas);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer keeps an off-viewport object when its shadow intersects the viewport", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "shape:shadow-edge",
      bounds: { x: 105, y: 20, width: 10, height: 10 },
      visual: {
        kind: "effect",
        shadow: { color: 0x00000080, blur: 4, offsetX: -10, offsetY: 0 },
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0xffffffff },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 80 },
      { unitIndex: 0, viewport: { x: 0, y: 0, width: 100, height: 80 } },
    );

    assert.equal(result.renderedObjectCount, 1);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer keeps an off-viewport object when its soft edge intersects the viewport", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
      this.context = recordingContext();
      this.context.canvas = this;
      this.context.getTransform = () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
    }
    getContext() { return this.context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "shape:soft-edge",
      bounds: { x: 102, y: 20, width: 10, height: 10 },
      visual: {
        kind: "advanced-effect",
        softEdge: 4,
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0xffffffff },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const result = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 80 },
      { unitIndex: 0, viewport: { x: 0, y: 0, width: 100, height: 80 } },
    );

    assert.equal(result.renderedObjectCount, 1);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer paints a scaled DrawingML outer shadow as a separate alpha mask", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const contexts = [];
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
      this.context = recordingContext();
      this.context.canvas = this;
      this.context.getTransform = () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
      this.context.setTransform = (...args) => this.context.calls.push(["setTransform", ...args]);
      this.context.resetTransform = () => this.context.calls.push(["resetTransform"]);
      contexts.push(this.context);
    }
    getContext() { return this.context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "shape:scaled-shadow",
      bounds: { x: 10, y: 10, width: 80, height: 40 },
      visual: {
        kind: "advanced-effect",
        outerShadow: {
          color: 0x00000026,
          blur: 16,
          offsetX: 0,
          offsetY: 100 / 3,
          scaleX: 0.9,
          scaleY: -0.19,
          skewX: 0,
          skewY: 0,
          alignment: 7,
        },
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0x4c91cfff },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 100 },
      { unitIndex: 0 },
    );

    assert.ok(contexts.some((context) => context.calls.some((call) => (
      call[0] === "transform" && call[1] === 0.9 && call[4] === -0.19
    ))));
    assert.ok(contexts.some((context) => context.calls.some((call) => (
      call[0] === "fillRect" && call.at(-1) === "source-in"
    ))));
    assert.ok(contexts.some((context) => context.filters.includes("blur(8px)")));
    assert.ok(contexts.some((context) => context.calls.some((call) => call[0] === "drawImage")));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

function recordingGlowContext() {
  const context = recordingContext();
  context.imageShadows = [];
  context.getTransform = () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
  context.getImageData = (x, y, width, height) => ({ width, height, data: new Uint8ClampedArray(width * height * 4) });
  context.putImageData = () => {};
  const drawImage = context.drawImage.bind(context);
  context.drawImage = (...args) => {
    context.imageShadows.push({ color: context.shadowColor, blur: context.shadowBlur });
    drawImage(...args);
  };
  return context;
}

test("renderer composes inherited and object-local advanced effects", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingGlowContext();
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; }
    getContext() { context.canvas = this; return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const group = sceneObject({
      numericId: 1,
      id: "group:effects",
      type: "group",
      bounds: { x: 0, y: 0, width: 50, height: 50 },
      visual: {
        kind: "advanced-effect",
        glow: { color: 0xff000080, radius: 2 },
        visual: { kind: "none" },
      },
    });
    const child = sceneObject({
      numericId: 2,
      parentNumericId: 1,
      parentId: "group:effects",
      id: "shape:effects",
      bounds: { x: 10, y: 10, width: 20, height: 20 },
      visual: {
        kind: "advanced-effect",
        glow: { color: 0x00ff0080, radius: 2 },
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0xffffffff },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      },
    });
    const renderer = new SceneRenderer([group, child], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 60 },
      { unitIndex: 0 },
    );

    assert.equal(context.calls.filter((call) => call[0] === "fill").length, 1);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer includes a DrawingML glow in the reflected image", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingGlowContext();
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; }
    getContext() { context.canvas = this; return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const renderer = new SceneRenderer([sceneObject({
      numericId: 1,
      id: "shape:glow-reflection",
      bounds: { x: 10, y: 10, width: 40, height: 30 },
      visual: {
        kind: "advanced-effect",
        glow: { color: 0x00b0f080, radius: 4 },
        reflection: {
          startOpacity: 0.5, endOpacity: 0.5, startPosition: 0, endPosition: 1,
          directionDegrees: 90, blur: 0, distance: 2, scaleX: 1, scaleY: 1,
        },
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0xffffffff },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      },
    })], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 80 },
      { unitIndex: 0 },
    );

    assert.equal(context.imageShadows.filter(({ color }) => color === "rgba(0, 176, 240, 0.5019607843137255)").length, 2);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer applies a glow inside the same DrawingML 3D camera pass", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingGlowContext();
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; }
    getContext() { context.canvas = this; return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "shape:glow-camera",
      bounds: { x: 0, y: 0, width: 100, height: 80 },
      visual: {
        kind: "advanced-effect",
        glow: { color: 0xff000080, radius: 6 },
        threeD: {
          cameraPreset: "isometricLeftDown",
          cameraFov: 0,
          cameraZoom: 1,
          cameraLatitude: 35,
          cameraLongitude: 45,
          cameraRevolution: 0,
          lightRig: "threePt",
          lightDirection: "t",
          lightLatitude: 0,
          lightLongitude: 0,
          lightRevolution: 0,
          z: 0,
          extrusionHeight: 0,
          contourWidth: 0,
          material: "matte",
        },
        visual: {
          kind: "effect",
          clip: "rectangle",
          visual: {
            kind: "painted-shape",
            geometry: "rectangle",
            fill: { kind: "solid", color: 0xffffffff },
            stroke: { kind: "none" },
            strokeWidth: 0,
          },
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 120, height: 100 },
      { unitIndex: 0 },
    );

    const fills = context.calls.filter((call) => call[0] === "fill");
    assert.equal(fills.length, 1);
    assert.ok(
      context.calls.findIndex((call) => call[0] === "clip")
        > context.calls.findIndex((call) => call[0] === "transform"),
      `picture crop ran before its 3D camera: ${JSON.stringify(context.calls)}`,
    );
    assert.ok(context.imageShadows.some(({ color, blur }) => color === "rgba(255, 0, 0, 0.5019607843137255)" && blur === 3));
    const camera = context.calls.find((call) => call[0] === "transform" && Math.abs(call[2]) > 0.1);
    assert.ok(Math.abs(camera[1] - Math.SQRT1_2) < 0.001);
    assert.ok(Math.abs(camera[2] - Math.sin(35 * Math.PI / 180) / Math.sqrt(2)) < 0.001);
    assert.ok(Math.abs(camera[4] - Math.cos(35 * Math.PI / 180)) < 0.001);
    assert.ok(context.strokeShadows.every(value => !value.startsWith("rgba(255, 0, 0,")));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("supplied radar renders distinct lit bevel facets from the real chart paths", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/chart-Radar.docx", import.meta.url)));
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { context.canvas = this; return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const polygons = document.scene.objects.filter(o => o.visual.threeD && o.visual.visual?.geometry?.kind === "path");
    assert.equal(polygons.length, 2);
    assert.notEqual(polygons[0].visual.visual.fill.color, polygons[1].visual.visual.fill.color);
    for (const polygon of polygons) {
      // Isolate the bevel from the shadow's independent raster/composite work.
      const object = { ...polygon, visual: { ...polygon.visual, outerShadow: undefined } };
      context.gradients.length = 0;
      await new SceneRenderer([object], DEFAULT_LIMITS).render(
        { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 794, height: 1123 }, { unitIndex: 0 });
      assert.equal(context.gradients.length, 5, "each edge needs its own lighting normal");
      assert.ok(new Set(context.gradients.map(g => JSON.stringify(g.stops))).size > 2,
        "facets facing different directions must have different lighting");
    }
  } finally {
    document.close(); core.close();
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer paints the DrawingML top bevel above the opaque front face", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { context.canvas = this; return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "shape:top-bevel",
      visual: {
        kind: "advanced-effect",
        threeD: {
          cameraPreset: "orthographicFront",
          cameraFov: 0,
          cameraZoom: 1,
          cameraLatitude: 0,
          cameraLongitude: 0,
          cameraRevolution: 0,
          lightRig: "threePt",
          lightDirection: "t",
          lightLatitude: 0,
          lightLongitude: 0,
          lightRevolution: 0,
          z: 0,
          extrusionHeight: 0,
          contourWidth: 0,
          material: "warmMatte",
          bevelTop: { width: 8, height: 8, preset: "circle" },
        },
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: { kind: "solid", color: 0x558ac8ff },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 60 },
      { unitIndex: 0 },
    );

    assert.ok(context.calls.some(call => call[0] === "clearRect"), "bevel overlay must not duplicate the opaque front face");
    assert.equal(context.gradients.length, 4);
    const stops = context.gradients.flatMap(gradient => gradient.stops);
    assert.ok(stops.some(([, value]) => value.startsWith("rgba(255,255,255,")));
    assert.ok(stops.some(([, value]) => value.startsWith("rgba(0,0,0,")));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("rich text drawing consumes the same run layout returned by the measurement API", async () => {
  const context = recordingContext();
  const runs = [
    { text: "Hi", fontFamily: "Aptos", fontSize: 12, color: 0x000000ff, bold: false, italic: false, underline: false, strikethrough: false, highlight: 0, baselineShift: 0, letterSpacing: 0 },
    { text: "!", fontFamily: "Aptos", fontSize: 12, color: 0xff0000ff, bold: true, italic: false, underline: true, strikethrough: true, highlight: 0xffee88ff, baselineShift: 2, letterSpacing: 0 },
  ];
  const layout = layoutTextRuns(context, runs, { maxWidth: 92, lineHeight: 18, align: "start" });
  assert.deepEqual(layout.runs.map(({ text, x, y, width }) => ({ text, x, y, width })), [
    { text: "Hi", x: 0, y: 0, width: 10 },
    { text: "!", x: 10, y: 0, width: 5 },
  ]);
  assert.deepEqual({ width: layout.width, height: layout.height, lineCount: layout.lineCount }, {
    width: 15,
    height: 18,
    lineCount: 1,
  });

  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "text:1",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 100, height: 30 },
      visual: {
        kind: "rich-text",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align: "start",
        lineHeight: 18,
        runs,
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 30 },
      { unitIndex: 0 },
    );
    assert.deepEqual(
      context.calls.filter((call) => call[0] === "fillText").map(([, text, x, y]) => ({ text, x, y })),
      [{ text: "Hi", x: 4, y: 4 }, { text: "!", x: 14, y: 2 }],
    );
    assert.equal(context.calls.some((call) => call[0] === "fillRect" && call[1] === 14 && call[2] === 2), true);
    assert.equal(context.calls.filter((call) => call[0] === "stroke").length >= 2, true);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer distinguishes word-processing and DrawingML text baselines", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const renderLine = async (lineHeight, type = "text-box") => {
    const context = recordingContext();
    const drawn = [];
    context.measureText = function measureText(text) {
      return this.textBaseline === "alphabetic"
        ? {
            width: text.length * 5,
            fontBoundingBoxAscent: 20,
            fontBoundingBoxDescent: 4,
          }
        : {
            width: text.length * 5,
            fontBoundingBoxAscent: 6,
            fontBoundingBoxDescent: 18,
          };
    };
    context.fillText = function fillText(text, x, y) {
      drawn.push({ text, x, y, baseline: this.textBaseline });
    };
    class FakeCanvas {
      getContext() { return context; }
      transferToImageBitmap() { return { close() {} }; }
    }
    Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
    const object = sceneObject({
      numericId: 1,
      id: `text:baseline:${lineHeight}`,
      type,
      bounds: { x: 0, y: 0, width: 100, height: 80 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "ltr",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 1,
          insetLeft: 4,
          insetRight: 4,
          insetTop: 4,
          insetBottom: 4,
          paragraphs: [{
            align: "start",
            marginLeft: 0,
            marginRight: 0,
            firstLineIndent: 0,
            defaultTabStop: 36,
            lineHeight,
            spaceBefore: 0,
            spaceAfter: 0,
          }],
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "start",
          lineHeight: 24,
          runs: [{
            text: "Office",
            fontFamily: "Aptos",
            fontSize: 20,
            color: 0x000000ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baselineShift: 0,
            letterSpacing: 0,
          }],
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const frame = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 80 },
      { unitIndex: 0 },
    );
    frame.bitmap.close();
    return drawn;
  };

  try {
    assert.deepEqual(await renderLine(24), [
      { text: "Office", x: 4, y: 24, baseline: "alphabetic" },
    ]);
    assert.deepEqual(await renderLine(21.6), [
      { text: "Office", x: 4, y: 22, baseline: "alphabetic" },
    ]);
    assert.deepEqual(await renderLine(21.6, "paragraph"), [
      { text: "Office", x: 4, y: 24, baseline: "alphabetic" },
    ]);
    assert.deepEqual(await renderLine(48), [
      { text: "Office", x: 4, y: 42, baseline: "alphabetic" },
    ]);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer excludes DOCX paragraph leading from character highlights", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  context.measureText = function measureText(text) {
    return this.textBaseline === "alphabetic"
      ? {
          width: text.length * 5,
          actualBoundingBoxAscent: 12,
          fontBoundingBoxAscent: 14,
          fontBoundingBoxDescent: 4,
        }
      : {
          width: text.length * 5,
          fontBoundingBoxAscent: 3,
          fontBoundingBoxDescent: 15,
        };
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "docx:highlight",
      type: "paragraph",
      bounds: { x: 0, y: 0, width: 100, height: 27 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "ltr",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 1,
          insetLeft: 0,
          insetRight: 0,
          insetTop: 0,
          insetBottom: 0,
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "start",
          lineHeight: 27,
          runs: [{
            text: "S32K358",
            fontFamily: "DengXian",
            fontSize: 13.333,
            color: 0x000000ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0xffff00ff,
            baselineShift: 0,
            letterSpacing: 0,
          }],
        },
      },
    });
    object.source.format = "docx";
    object.source.part = "word/document.xml";
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 40 },
      { unitIndex: 0 },
    );

    const highlight = context.calls.find(
      (call) => call[0] === "fillRect" && call[3] === 35,
    );
    assert.equal(highlight?.[4], 18);
    assert.equal(highlight?.[2], 9, "highlight follows the font box below natural leading");
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer paints adjacent CJK tokens in one shaping run", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "docx:cjk-shaping",
      type: "paragraph",
      bounds: { x: 0, y: 0, width: 200, height: 24 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "ltr",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 1,
          insetLeft: 0,
          insetRight: 0,
          insetTop: 0,
          insetBottom: 0,
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "start",
          lineHeight: 24,
          runs: [
            {
              text: "甲方指定地点（（上海市浦东新区白莲泾",
              fontFamily: "DengXian",
              fontSize: 16,
              color: 0x000000ff,
              bold: true,
              italic: false,
              underline: false,
              strikethrough: false,
              highlight: 0,
              baselineShift: 0,
              letterSpacing: 0,
            },
            {
              text: "路",
              fontFamily: "DengXian",
              fontSize: 16,
              color: 0x000000ff,
              bold: true,
              italic: false,
              underline: false,
              strikethrough: false,
              highlight: 0,
              baselineShift: 0,
              letterSpacing: 4,
            },
            {
              text: "127号21楼））",
              fontFamily: "DengXian",
              fontSize: 16,
              color: 0x000000ff,
              bold: true,
              italic: false,
              underline: false,
              strikethrough: false,
              highlight: 0,
              baselineShift: 0,
              letterSpacing: 0,
            },
          ],
        },
      },
    });
    object.source.format = "docx";
    object.source.part = "word/document.xml";
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 200, height: 30 },
      { unitIndex: 0 },
    );

    const painted = context.calls
      .filter((call) => call[0] === "fillText")
      .map(([, text]) => text);
    assert.equal(painted.some((text) => text.includes("上海")), true);
    assert.equal(painted.some((text) => text.includes("白莲泾路")), true);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer shares one alphabetic baseline across mixed-script fonts on a line", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  const drawn = [];
  context.measureText = function measureText(text) {
    const eastAsian = this.font.includes("East Asian Test");
    if (this.textBaseline === "alphabetic") {
      return {
        width: text.length * 5,
        actualBoundingBoxAscent: eastAsian ? 20 : 17,
        fontBoundingBoxAscent: eastAsian ? 21 : 26,
        fontBoundingBoxDescent: eastAsian ? 3 : 6,
      };
    }
    return {
      width: text.length * 5,
      fontBoundingBoxAscent: eastAsian ? -0.125 : 7,
      fontBoundingBoxDescent: eastAsian ? 24 : 25,
    };
  };
  context.fillText = function fillText(text, x, y) {
    drawn.push({ text, x, y, baseline: this.textBaseline });
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const object = sceneObject({
    numericId: 1,
    id: "text:mixed-script-baseline",
    type: "text-box",
    bounds: { x: 0, y: 0, width: 100, height: 40 },
    visual: {
      kind: "text-layout",
      layout: {
        direction: "ltr",
        orientation: "horizontal",
        autoFit: "none",
        verticalAlign: "top",
        defaultTabStop: 36,
        hangingIndent: 0,
        minScale: 1,
        insetLeft: 4,
        insetRight: 4,
        insetTop: 4,
        insetBottom: 4,
      },
      visual: {
        kind: "rich-text",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align: "start",
        lineHeight: 24,
        runs: [
          {
            text: "中",
            fontFamily: "East Asian Test",
            fontSize: 20,
            color: 0x000000ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baselineShift: 0,
            letterSpacing: 0,
          },
          {
            text: "npm",
            fontFamily: "Latin Test",
            fontSize: 20,
            color: 0x000000ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baselineShift: 0,
            letterSpacing: 0,
          },
        ],
      },
    },
  });
  const renderer = new SceneRenderer([object], DEFAULT_LIMITS);

  try {
    const frame = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 40 },
      { unitIndex: 0 },
    );
    frame.bitmap.close();

    assert.deepEqual(drawn, [
      { text: "中", x: 4, y: 30, baseline: "alphabetic" },
      { text: "npm", x: 9, y: 30, baseline: "alphabetic" },
    ]);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer resolves and reuses the exact managed face for each text run", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  const measured = [];
  const painted = [];
  const requestedFaces = [];
  context.measureText = function measureText(text) {
    measured.push({ text, font: this.font });
    return {
      width: text.length * 5,
      fontBoundingBoxAscent: 10,
      fontBoundingBoxDescent: 2,
    };
  };
  context.fillText = function fillText(text) {
    painted.push({ text, font: this.font });
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const fonts = {
    resolve(family) { return { family, source: "browser" }; },
    resolveFace(face) {
      requestedFaces.push(face);
      const variant = face.style === "italic" ? "italic" : face.weight === 700 ? "bold" : "regular";
      return { ...face, family: `Managed ${variant}`, source: "provider" };
    },
  };
  const run = (text, overrides = {}) => ({
    text,
    fontFamily: "Authored Sans",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
    ...overrides,
  });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "text:managed-faces",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 100, height: 30 },
      visual: {
        kind: "rich-text",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align: "start",
        lineHeight: 18,
        runs: [run("N"), run("B", { bold: true }), run("I", { italic: true })],
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS, fonts);
    const frame = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 30 },
      { unitIndex: 0 },
    );
    frame.bitmap.close();

    assert.deepEqual(requestedFaces.map(({ family, style, weight, codePoints }) => ({
      family,
      style,
      weight,
      codePoints,
    })), [
      { family: "Authored Sans", style: "normal", weight: 400, codePoints: [0x4e] },
      { family: "Authored Sans", style: "normal", weight: 700, codePoints: [0x42] },
      { family: "Authored Sans", style: "italic", weight: 400, codePoints: [0x49] },
    ]);
    assert.equal(measured.some(({ text, font }) => text === "N" && font === '400 12px "Managed regular"'), true);
    assert.equal(measured.some(({ text, font }) => text === "B" && font === '700 12px "Managed bold"'), true);
    assert.equal(measured.some(({ text, font }) => text === "I" && font === 'italic 400 12px "Managed italic"'), true);
    assert.deepEqual(painted, [
      { text: "N", font: '400 12px "Managed regular"' },
      { text: "B", font: '700 12px "Managed bold"' },
      { text: "I", font: 'italic 400 12px "Managed italic"' },
    ]);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer reuses an exact embedded PDF face across different text", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const requestedFaces = [];
  const fonts = {
    resolve(family) { return { family, source: "embedded" }; },
    resolveFace(face) {
      requestedFaces.push(face);
      return { ...face, source: "embedded" };
    },
  };
  const textObject = (numericId, text, y) => ({
    ...sceneObject({
      numericId,
      id: `text:pdf-face:${numericId}`,
      type: "text-box",
      bounds: { x: 0, y, width: 100, height: 20 },
      visual: {
        kind: "text",
        geometry: "rectangle",
        fill: 0x00000000,
        stroke: 0x00000000,
        strokeWidth: 0,
        fontFamily: "ABCDEF+Example PDF 5 0",
        fontSize: 12,
        color: 0x000000ff,
        bold: false,
        italic: false,
        align: "start",
        lineHeight: 14,
      },
    }),
    text,
  });
  try {
    const renderer = new SceneRenderer([
      textObject(1, "First", 0),
      textObject(2, "Second", 20),
    ], DEFAULT_LIMITS, fonts);
    const frame = await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 40 },
      { unitIndex: 0 },
    );
    frame.bitmap.close();
    assert.equal(requestedFaces.length, 1);
    assert.deepEqual(requestedFaces[0].codePoints, [..."First"].map((character) => character.codePointAt(0))
      .filter((codePoint, index, codePoints) => codePoints.indexOf(codePoint) === index)
      .sort((left, right) => left - right));
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("PDF text batches replay the original positioned text draws", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const run = (text) => ({
    text,
    fontFamily: "Helvetica",
    fontSize: 10,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  });
  const child = (text) => ({
    kind: "text-layout",
    layout: {
      direction: "ltr",
      orientation: "horizontal",
      autoFit: "none",
      verticalAlign: "top",
      defaultTabStop: 36,
      hangingIndent: 0,
      minScale: 0.1,
      insetLeft: 0,
      insetRight: 0,
      insetTop: 0,
      insetBottom: 0,
      marginLeft: 0,
      marginRight: 0,
      columnCount: 1,
      columnSpacing: 0,
      rotationDegrees: 0,
      horizontalOverflow: "overflow",
      verticalOverflow: "overflow",
      wrap: false,
      textFill: true,
      textBaseline: 10,
    },
    visual: {
      kind: "rich-text",
      geometry: "rectangle",
      fill: { kind: "none" },
      stroke: { kind: "none" },
      strokeWidth: 0,
      align: "start",
      lineHeight: 10,
      runs: [run(text)],
    },
  });
  const pdfObject = (numericId, text, bounds, visual) => ({
    ...sceneObject({ numericId, id: `pdf:text:${numericId}`, type: "text-box", bounds, visual }),
    text,
    source: {
      format: "pdf",
      part: "document.pdf",
      kind: "text",
      objectNumber: 4,
      byteOffset: 10,
      mapping: "derived",
    },
  });
  const firstBounds = { x: 10, y: 20, width: 7, height: 10 };
  const secondBounds = { x: 17, y: 20, width: 7, height: 10 };
  const renderCalls = async (objects) => {
    const context = recordingContext();
    context.getTransform = () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
    class FakeCanvas {
      getContext() { return context; }
      transferToImageBitmap() { return { close() {} }; }
    }
    Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
    const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
    const frame = await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 50 },
      { unitIndex: 0, includeTextFragments: true },
    );
    frame.bitmap.close();
    return { calls: context.calls, textFragments: frame.textFragments };
  };
  try {
    const separate = await renderCalls([
      pdfObject(1, "A", firstBounds, child("A")),
      pdfObject(2, "B", secondBounds, child("B")),
    ]);
    const batch = await renderCalls([
      pdfObject(1, "AB", { x: 10, y: 20, width: 14, height: 10 }, {
        kind: "group",
        children: [
          { bounds: firstBounds, visual: child("A") },
          { bounds: secondBounds, visual: child("B") },
        ],
      }),
    ]);
    assert.deepEqual(batch.calls, separate.calls);
    const genericGroup = await renderCalls([{
      ...pdfObject(1, "AB", { x: 10, y: 20, width: 14, height: 10 }, {
        kind: "group",
        children: [
          { bounds: firstBounds, visual: child("A") },
          { bounds: secondBounds, visual: child("B") },
        ],
      }),
      source: { format: "xps", part: "page.fpage", kind: "text", mapping: "exact" },
    }]);
    assert.deepEqual(genericGroup.textFragments.map(({ text, start, end, transform }) =>
      ({ text, start, end, x: transform.e, y: transform.f })), [
      { text: "A", start: 0, end: 1, x: 10, y: 20 },
      { text: "B", start: 1, end: 2, x: 17, y: 20 },
    ]);
    const semanticBatch = await renderCalls([
      pdfObject(1, "AB", { x: 10, y: 20, width: 14, height: 10 }, {
        kind: "group",
        children: [
          { bounds: firstBounds, visual: child("\ue041") },
          { bounds: secondBounds, visual: child("\ue042") },
        ],
      }),
    ]);
    assert.deepEqual(semanticBatch.textFragments.map(({ text }) => text), ["A", "B"]);
    const semanticLigatureBatch = await renderCalls([
      pdfObject(1, "fix", { x: 10, y: 20, width: 14, height: 10 }, {
        kind: "group",
        children: [
          { bounds: firstBounds, visual: child("\ue01f") },
          { bounds: secondBounds, visual: child("\ue169") },
        ],
      }),
    ]);
    assert.deepEqual(semanticLigatureBatch.textFragments.map(({ text }) => text), ["fi", "x"]);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer distinguishes embedded PDF advances from authored matrix scaling and stroke paint", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  let measurements = 0;
  let paintedKerning;
  context.measureText = () => {
    measurements += 1;
    return { width: 40 };
  };
  const fillText = context.fillText;
  context.fillText = function (...args) {
    paintedKerning = this.fontKerning;
    return fillText.apply(this, args);
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const object = {
    ...sceneObject({
      numericId: 1,
      id: "text:pdf-advance",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 40, height: 20 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "ltr",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 0.1,
          insetLeft: 0,
          insetRight: 0,
          insetTop: 0,
          insetBottom: 0,
          marginLeft: 0,
          marginRight: 0,
          columnCount: 1,
          columnSpacing: 0,
          rotationDegrees: 0,
          horizontalOverflow: "overflow",
          verticalOverflow: "overflow",
          wrap: false,
          textScaleToFit: true,
          textBaseline: 12,
          textPaint: {
            kind: "linear-gradient",
            start: { x: 0, y: 0 },
            end: { x: 40, y: 0 },
            stops: gradient.stops,
          },
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "start",
          lineHeight: 14,
          runs: [{
            text: "Fast",
            fontFamily: "ABCDEF+Example PDF 5 0",
            fontSize: 12,
            color: 0x000000ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baselineShift: 0,
            letterSpacing: 0,
          }],
        },
      },
    }),
    source: {
      format: "pdf",
      part: "page:1",
      kind: "text",
      mapping: "exact",
    },
  };
  const fonts = {
    resolve: (family) => ({ family, source: "embedded" }),
    resolveFace: (face) => ({ ...face, source: "embedded" }),
  };
  try {
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS, fonts);
    const frame = await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 40, height: 20 },
      { unitIndex: 0 },
    );
    frame.bitmap.close();
    assert.equal(paintedKerning, "none");
    assert.equal(measurements, 0);
    assert.deepEqual(context.gradients[0].args, [0, -12, 40, -12]);
    assert.ok(
      context.calls.findIndex((call) => call[0] === "translate")
        < context.calls.findIndex((call) => call[0] === "createLinearGradient"),
    );

    const patternedObject = {
      ...object,
      bounds: { ...object.bounds, width: 80 },
      visual: {
        ...object.visual,
        layout: {
          ...object.visual.layout,
          textPaint: undefined,
          textFill: false,
          textMatrixScaleToFit: true,
          textStrokePaint: {
            kind: "linear-gradient",
            start: { x: 0, y: 0 },
            end: { x: 80, y: 0 },
            stops: gradient.stops,
          },
          textStrokeWidth: 2,
        },
      },
    };
    const patternedFrame = await new SceneRenderer([patternedObject], DEFAULT_LIMITS, fonts).render(
      { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 80, height: 20 },
      { unitIndex: 0 },
    );
    patternedFrame.bitmap.close();
    assert.equal(measurements, 1);
    assert.equal(context.calls.some((call) => call[0] === "scale" && call[1] === 2 && call[2] === 1), true);
    assert.equal(context.gradients.length, 2);
    assert.equal(context.calls.some((call) => (
      call[0] === "strokeText" && call.at(-1) === context.gradients[1].result
    )), true);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer scopes browser font availability to the run's Unicode coverage", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const checks = [];
  const fonts = {
    add() {},
    delete() { return true; },
    check(font, text) {
      checks.push({ font, text });
      return false;
    },
  };
  const run = (text) => ({
    text,
    fontFamily: "Coverage Sans",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "text:coverage-probe",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 100, height: 30 },
      visual: {
        kind: "rich-text",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align: "start",
        lineHeight: 18,
        runs: [run("AB"), run("BA"), run("Ω")],
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS, fonts);
    const frame = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 30 },
      { unitIndex: 0 },
    );
    frame.bitmap.close();

    assert.deepEqual(checks, [
      { font: '400 12px "Coverage Sans"', text: "AB" },
      { font: '400 12px "Coverage Sans"', text: "Ω" },
    ]);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer positions text decorations relative to an alphabetic baseline", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  context.measureText = function measureText(text) {
    return this.textBaseline === "alphabetic"
      ? {
          width: text.length * 5,
          actualBoundingBoxAscent: 14,
          fontBoundingBoxAscent: 20,
          fontBoundingBoxDescent: 4,
        }
      : {
          width: text.length * 5,
          fontBoundingBoxAscent: 6,
          fontBoundingBoxDescent: 18,
        };
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "text:decorations",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 100, height: 40 },
      visual: {
        kind: "rich-text",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align: "start",
        lineHeight: 24,
        runs: [{
          text: "Office",
          fontFamily: "Aptos",
          fontSize: 20,
          color: 0x000000ff,
          bold: false,
          italic: false,
          underline: true,
          strikethrough: true,
          highlight: 0,
          baselineShift: 0,
          letterSpacing: 0,
        }],
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const frame = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 40 },
      { unitIndex: 0 },
    );
    frame.bitmap.close();

    assert.deepEqual(
      context.calls.filter((call) => call[0] === "moveTo"),
      [["moveTo", 4, 26], ["moveTo", 4, 17]],
    );
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer aligns same-size underlines across mixed-script fonts", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  context.measureText = function measureText(text) {
    const eastAsian = this.font.includes("DengXian");
    return this.textBaseline === "alphabetic"
      ? {
          width: text.length * 10,
          actualBoundingBoxAscent: 14,
          fontBoundingBoxAscent: 20,
          fontBoundingBoxDescent: eastAsian ? 0 : 4,
        }
      : {
          width: text.length * 10,
          fontBoundingBoxAscent: 5,
          fontBoundingBoxDescent: eastAsian ? 0 : 4,
        };
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const shared = {
      fontSize: 20,
      color: 0x000000ff,
      bold: true,
      italic: false,
      underline: true,
      strikethrough: false,
      highlight: 0,
      baselineShift: 0,
      letterSpacing: 0,
    };
    const object = sceneObject({
      numericId: 1,
      id: "text:mixed-script-underline",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 100, height: 40 },
      visual: {
        kind: "rich-text",
        geometry: "rectangle",
        fill: { kind: "none" },
        stroke: { kind: "none" },
        strokeWidth: 0,
        align: "start",
        lineHeight: 24,
        runs: [
          { ...shared, text: "2026", fontFamily: "Arial" },
          { ...shared, text: "年", fontFamily: "DengXian" },
        ],
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const frame = await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 40 },
      { unitIndex: 0 },
    );
    frame.bitmap.close();

    const underlineRows = context.calls
      .filter((call) => call[0] === "moveTo")
      .map(([, , y]) => Math.round(y * 1_000) / 1_000);
    assert.equal(underlineRows.length, 2);
    assert.equal(new Set(underlineRows).size, 1);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer preserves DrawingML's default vertical text overflow", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "text:vertical-overflow",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 100, height: 10 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "ltr",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 0.1,
          insetLeft: 0,
          insetRight: 0,
          insetTop: 0,
          insetBottom: 0,
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "start",
          lineHeight: 18,
          runs: [{
            text: "A\nB",
            fontFamily: "Arial",
            fontSize: 12,
            color: 0x000000ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baselineShift: 0,
            letterSpacing: 0,
          }],
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 60 },
      { unitIndex: 0 },
    );

    assert.deepEqual(
      context.calls.filter((call) => call[0] === "fillText").map(([, text, , y]) => ({ text, y })),
      [{ text: "A", y: 0 }, { text: "B", y: 18 }],
    );
    assert.equal(context.calls.some((call) => call[0] === "clip"), false);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer consumes text-layout prefix, tabs, and paragraph direction", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const run = {
      text: "甲\t乙",
      fontFamily: "DengXian",
      fontSize: 12,
      color: 0x000000ff,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      highlight: 0,
      baselineShift: 0,
      letterSpacing: 0,
    };
    const object = sceneObject({
      numericId: 1,
      id: "text:layout",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 40, height: 30 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "rtl",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          prefix: "• ",
          tabStops: [20],
          defaultTabStop: 10,
          hangingIndent: 10,
          minScale: 0.5,
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "start",
          lineHeight: 18,
          runs: [run],
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 40, height: 30 },
      { unitIndex: 0 },
    );

    assert.deepEqual(
      context.calls.filter((call) => call[0] === "fillText").map(([, text, x, , direction, align]) => ({
        text,
        x,
        direction,
        align,
      })),
      [
        { text: "• ", x: 26, direction: "rtl", align: "left" },
        { text: "甲", x: 21, direction: "rtl", align: "left" },
        { text: "乙", x: 11, direction: "rtl", align: "left" },
      ],
    );
    const directCalls = context.calls.filter((call) => call[0] === "fillText");
    context.calls.length = 0;
    const grouped = {
      ...object,
      visual: {
        ...object.visual,
        visual: {
          kind: "group",
          children: [{ bounds: object.bounds, visual: object.visual.visual }],
        },
      },
    };
    await new SceneRenderer([grouped], DEFAULT_LIMITS).render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 40, height: 30 },
      { unitIndex: 0 },
    );
    assert.deepEqual(context.calls.filter((call) => call[0] === "fillText"), directCalls,
      "group children must inherit the same text layout as a direct visual");

  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("text layout wraps CJK text without requiring spaces", () => {
  const context = recordingContext();
  const runs = [{
    text: "甲乙丙丁",
    fontFamily: "DengXian",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, { maxWidth: 10, lineHeight: 18 });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "甲", x: 0, line: 0 },
      { text: "乙", x: 5, line: 0 },
      { text: "丙", x: 0, line: 1 },
      { text: "丁", x: 5, line: 1 },
    ],
  );
  assert.equal(layout.lineCount, 2);
});

test("text layout hangs CJK closing punctuation on the previous line", () => {
  const context = recordingContext();
  const runs = [{
    text: "设备、设施",
    fontFamily: "SimSun",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, { maxWidth: 10, lineHeight: 18 });
  const lines = Array.from({ length: layout.lineCount }, (_, line) => (
    layout.runs.filter((placement) => placement.line === line).map(({ text }) => text).join("")
  ));

  assert.deepEqual(lines, ["设备、", "设施"]);
});

test("text layout breaks an oversized Latin token at grapheme boundaries", () => {
  const context = recordingContext();
  const runs = [{
    text: "ABCD",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, { maxWidth: 10, lineHeight: 18 });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "A", x: 0, line: 0 },
      { text: "B", x: 5, line: 0 },
      { text: "C", x: 0, line: 1 },
      { text: "D", x: 5, line: 1 },
    ],
  );
  assert.equal(layout.lineCount, 2);
});

test("text layout keeps a fitting Latin word intact when latinLnBrk is disabled", () => {
  const context = recordingContext();
  const runs = [{
    text: "A ABCD",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];
  const paragraph = {
    align: "start",
    marginLeft: 0,
    marginRight: 0,
    firstLineIndent: 0,
    defaultTabStop: 36,
    latinLineBreak: false,
    hangingPunctuation: false,
  };

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 20,
    lineHeight: 18,
    paragraphs: [paragraph],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "A", x: 0, line: 0 },
      { text: " ", x: 5, line: 0 },
      { text: "ABCD", x: 0, line: 1 },
    ],
  );
  assert.equal(layout.lineCount, 2);
});

test("supplied font-scale keeps bullets beside emergency-wrapped body text", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("./fixtures/font-scale.pptx", import.meta.url)));
  try {
    const object = doc.scene.objects.find(object => object.name === "CustomShape 2");
    const { layout: authored, visual } = object.visual;
    const context = recordingContext();
    context.measureText = text => ({ width: [...text].reduce((width, char) => width + (char === "•" ? 8 : char === " " ? 6 : 22), 0) });
    const layout = layoutTextRuns(context, visual.runs, {
      ...authored, autoFit: "none", lineHeight: visual.lineHeight,
      maxWidth: object.bounds.width - authored.insetLeft - authored.insetRight,
    });
    const bullets = layout.runs.filter(run => run.text === "•");
    assert.equal(bullets.length, 3);
    for (const [index, bullet] of bullets.entries()) {
      const body = layout.runs.find(run => run.text.startsWith("ABC"[index]));
      assert.equal(bullet.line, body.line);
      assert.ok(Math.abs(body.x - authored.paragraphs[index].marginLeft) < 0.001);
    }
    assert.ok(layout.lineCount > 10);
    assert.equal(layout.runs.map(run => run.text).join(""), visual.runs.map(run => run.text).join("").replaceAll("\n", "").replaceAll("\u2028", ""));
  } finally { doc.close(); core.close(); }
});

test("text layout wraps an oversized URL when latinLnBrk is disabled", () => {
  const context = recordingContext();
  const runs = [{
    text: "http://office/14",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: true,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];
  const layout = layoutTextRuns(context, runs, {
    maxWidth: 10,
    lineHeight: 18,
    paragraphs: [{
      align: "start",
      marginLeft: 0,
      marginRight: 0,
      firstLineIndent: 0,
      defaultTabStop: 36,
      latinLineBreak: false,
      hangingPunctuation: false,
    }],
  });

  assert.ok(layout.lineCount > 1);
  assert.equal(layout.runs.map(({ text }) => text).join(""), "http://office/14");
});

test("text layout permits East Asian terminal punctuation to hang beyond the paragraph edge", () => {
  const context = recordingContext();
  const runs = [{
    text: "X AB。",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];
  const paragraph = {
    align: "start",
    marginLeft: 0,
    marginRight: 0,
    firstLineIndent: 0,
    defaultTabStop: 36,
    latinLineBreak: false,
    hangingPunctuation: true,
  };

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 20,
    lineHeight: 18,
    paragraphs: [paragraph],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "X", x: 0, line: 0 },
      { text: " ", x: 5, line: 0 },
      { text: "AB", x: 10, line: 0 },
      { text: "。", x: 20, line: 0 },
    ],
  );
  assert.equal(layout.lineCount, 1);
});

test("text layout keeps ASCII closing punctuation off the next line", () => {
  const context = recordingContext();
  const runs = [{
    text: "X AB)",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];
  const paragraph = {
    align: "start",
    marginLeft: 0,
    marginRight: 0,
    firstLineIndent: 0,
    defaultTabStop: 36,
    latinLineBreak: false,
    hangingPunctuation: true,
  };

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 20,
    lineHeight: 18,
    paragraphs: [paragraph],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "X", x: 0, line: 0 },
      { text: " ", x: 5, line: 0 },
      { text: "AB", x: 10, line: 0 },
      { text: ")", x: 20, line: 0 },
    ],
  );
});

test("text layout keeps a fitting Latin compound and its closing punctuation together", () => {
  const context = recordingContext();
  const runs = [{
    text: "AAAA/BBBB)",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];
  const paragraph = {
    align: "start",
    marginLeft: 0,
    marginRight: 0,
    firstLineIndent: 0,
    defaultTabStop: 36,
    latinLineBreak: false,
    hangingPunctuation: true,
  };

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 45,
    lineHeight: 18,
    paragraphs: [paragraph],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "AAAA/BBBB", x: 0, line: 0 },
      { text: ")", x: 45, line: 0 },
    ],
  );
});

test("text layout preserves common Latin compound punctuation when latinLnBrk is disabled", () => {
  const context = recordingContext();
  const paragraph = {
    align: "start",
    marginLeft: 0,
    marginRight: 0,
    firstLineIndent: 0,
    defaultTabStop: 36,
    latinLineBreak: false,
    hangingPunctuation: true,
  };
  for (const text of ["RF/bench", "pre-production", "03/2026", "SWE.1", "ECU↔module"]) {
    const layout = layoutTextRuns(context, [{
      text,
      fontFamily: "Arial",
      fontSize: 12,
      color: 0x000000ff,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      highlight: 0,
      baselineShift: 0,
      letterSpacing: 0,
    }], {
      maxWidth: 100,
      lineHeight: 18,
      paragraphs: [paragraph],
    });
    assert.deepEqual(layout.runs.map((run) => run.text), [text]);
  }
});

test("text layout wraps after a hyphen without splitting ordinary words", () => {
  const context = recordingContext();
  const runs = [{
    text: "heavy-duty",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, { maxWidth: 30, lineHeight: 18 });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "heavy-", x: 0, line: 0 },
      { text: "duty", x: 0, line: 1 },
    ],
  );
  assert.equal(layout.lineCount, 2);
});

test("text layout keeps a styled word together across run boundaries", () => {
  const run = (text, color) => ({
    text,
    fontFamily: "Arial",
    fontSize: 12,
    color,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  });
  const layout = layoutTextRuns(recordingContext(), [
    run("AAAA ", 0x000000ff),
    run("D", 0xffcc00ff),
    run("R", 0xff0000ff),
    run("O", 0x3333ccff),
  ], { maxWidth: 30, lineHeight: 18 });

  assert.deepEqual(
    layout.runs.map(({ text, line }) => ({ text, line })),
    [
      { text: "AAAA", line: 0 },
      { text: " ", line: 0 },
      { text: "D", line: 1 },
      { text: "R", line: 1 },
      { text: "O", line: 1 },
    ],
  );
});

test("text layout collapses an overflowing space at a soft-wrap boundary", () => {
  const context = recordingContext();
  const runs = [{
    text: "AAAA BBBB",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, { maxWidth: 22, lineHeight: 18 });

  assert.deepEqual(
    layout.runs.map(({ text, x, width, line }) => ({ text, x, width, line })),
    [
      { text: "AAAA", x: 0, width: 20, line: 0 },
      { text: " ", x: 20, width: 0, line: 0 },
      { text: "BBBB", x: 0, width: 20, line: 1 },
    ],
  );
  assert.equal(layout.lineCount, 2);
});

test("text layout applies an explicit list prefix and tab stops", () => {
  const context = recordingContext();
  const runs = [{
    text: "甲\t乙",
    fontFamily: "DengXian",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 40,
    lineHeight: 18,
    prefix: "• ",
    tabStops: [20, 30],
    defaultTabStop: 10,
  });

  assert.deepEqual(
    layout.runs.map(({ text, x }) => ({ text, x })),
    [
      { text: "• ", x: 0 },
      { text: "甲", x: 10 },
      { text: "\t", x: 15 },
      { text: "乙", x: 20 },
    ],
  );
});

test("text layout right-aligns tab content and preserves a dot leader", () => {
  const context = recordingContext();
  const runs = [{
    text: "Title\t7",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 100,
    tabStops: [{ position: 90, align: "end", leader: "dot" }],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, width, leader }) => ({ text, x, width, leader })),
    [
      { text: "Title", x: 0, width: 25, leader: undefined },
      { text: "\t", x: 25, width: 60, leader: "dot" },
      { text: "7", x: 85, width: 5, leader: undefined },
    ],
  );
});

test("text layout clamps an out-of-bounds right tab to the text edge", () => {
  const context = recordingContext();
  const layout = layoutTextRuns(context, [{
    text: "Title\t7",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }], {
    maxWidth: 100,
    tabStops: [{ position: 120, align: "end", leader: "none" }],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, width, line }) => ({ text, x, width, line })),
    [
      { text: "Title", x: 0, width: 25, line: 0 },
      { text: "\t", x: 25, width: 69, line: 0 },
      { text: "7", x: 94, width: 5, line: 0 },
    ],
  );
});

test("text layout preserves an underlined tab as a drawable advance", () => {
  const context = recordingContext();
  const layout = layoutTextRuns(context, [{
    text: "\t",
    fontFamily: "STSong",
    fontSize: 14,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: true,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }], {
    maxWidth: 100,
    lineHeight: 20,
    defaultTabStop: 28,
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, width, style }) => ({ text, x, width, underline: style.underline })),
    [{ text: "\t", x: 0, width: 28, underline: true }],
  );
});

test("supplied Pages TOC aligns page fields beyond hanging paragraph indents", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-original.pages", import.meta.url)));
  try {
    let count = 0;
    for (const unitIndex of [2, 3]) {
      document.loadUnit(unitIndex);
      for (const object of document.scene.objects.filter(object => object.unitIndex === unitIndex && object.id.includes("-toc-"))) {
        const { layout, visual } = object.visual;
        const rendered = layoutTextRuns(recordingContext(), visual.runs, {
          ...layout, maxWidth: object.bounds.width, lineHeight: visual.lineHeight,
        });
        const page = rendered.runs.filter(run => run.text.trim()).at(-1);
        assert.ok(Math.abs(page.x + page.width - layout.tabStops.at(-1).position) < .01,
          `page field must reach the authored right tab: ${object.text}`);
        count += 1;
      }
    }
    assert.equal(count, 61);
  } finally { document.close(); core.close(); }
});

test("supplied vehicle deployment PPTX never backs body text into a bullet", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/vehicle-side-deployment-options.pptx", import.meta.url)));
  try {
    const context = {
      font: "", letterSpacing: "0px",
      measureText(text) {
        const size = Number(this.font.match(/([\d.]+)px/u)[1]);
        return { width: [...text].length * size * .5 };
      },
    };
    let markers = 0;
    for (const object of document.scene.objects.filter(object => object.text?.includes("•\t"))) {
      const { visual, layout } = object.visual;
      const rendered = layoutTextRuns(context, visual.runs, {
        ...layout, lineHeight: visual.lineHeight, align: visual.align, maxWidth: object.bounds.width,
      });
      for (const [index, marker] of rendered.runs.entries()) {
        if (marker.text !== "•") continue;
        const body = rendered.runs.slice(index + 1).find(run => run.text.trim());
        assert.ok(body, "every bullet retains its body text");
        assert.equal(body.line, marker.line, "body remains beside its bullet");
        assert.ok(body.x > marker.x + marker.width,
          `overlapping bullet: ${body.text}, marker ends at ${marker.x + marker.width}, body starts at ${body.x}`);
        markers += 1;
      }
    }
    assert.ok(markers >= 6, "exercise the original configuration and other list paragraphs");
  } finally { document.close(); core.close(); }
});

test("text layout anchors a hanging paragraph's first tab at its body margin", () => {
  const context = {
    font: "",
    letterSpacing: "0px",
    measureText(text) {
      return { width: text === "■" ? 13 : text.length * 5 };
    },
  };
  const style = {
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  };
  const paragraph = {
    align: "start",
    marginLeft: 12,
    marginRight: 0,
    firstLineIndent: -12,
    defaultTabStop: 96,
  };

  const fitting = layoutTextRuns(context, [{ ...style, text: "•\tAssets" }], {
    maxWidth: 60, lineHeight: 18, paragraphs: [paragraph],
  });
  assert.equal(fitting.runs.find(run => run.text === "Assets").x, 12,
    "a fitting marker preserves the authored body margin");

  const layout = layoutTextRuns(context, [
    { ...style, text: "■\t" },
    { ...style, text: "Assets" },
  ], {
    maxWidth: 60,
    lineHeight: 18,
    paragraphs: [paragraph],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "■", x: 0, line: 0 },
      { text: "\t", x: 13, line: 0 },
      { text: "Assets", x: 18, line: 0 },
    ],
  );

  const continuation = layoutTextRuns(context, [{ ...style, text: "■\tA\nBB\tC" }], {
    maxWidth: 120,
    lineHeight: 18,
    paragraphs: [paragraph],
  });
  assert.deepEqual(
    continuation.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "■", x: 0, line: 0 },
      { text: "\t", x: 13, line: 0 },
      { text: "A", x: 18, line: 0 },
      { text: "BB", x: 12, line: 1 },
      { text: "\t", x: 22, line: 1 },
      { text: "C", x: 96, line: 1 },
    ],
  );

  const extremeIndent = layoutTextRuns(context, [{ ...style, text: "Visible" }], {
    maxWidth: 120,
    lineHeight: 18,
    paragraphs: [{ ...paragraph, marginLeft: 0, firstLineIndent: -5_376 }],
  });
  assert.deepEqual(
    extremeIndent.runs.map(({ text, x }) => ({ text, x })),
    [{ text: "Visible", x: 0 }],
  );
});

test("text layout places an authored paragraph rule after wrapping is resolved", () => {
  const layout = layoutTextRuns(recordingContext(), [{
    text: "Title\nSubject",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }], {
    maxWidth: 100,
    lineHeight: 20,
    paragraphs: [
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36 },
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36,
        ruleAbove: { color: 0x515151ff, strokeWidth: .5, offsetX: -3, offsetY: -3, width: 1 },
        ruleBelow: { color: 0x515151ff, strokeWidth: .5, offsetX: -3, offsetY: 3, width: 1 } },
    ],
  });

  assert.deepEqual(layout.rules, [
    { x: -3, y: 17, width: 106, color: 0x515151ff, strokeWidth: .5 },
    { x: -3, y: 43, width: 106, color: 0x515151ff, strokeWidth: .5 },
  ]);
});

test("supplied Pages title preserves both paragraph borders and their reserved padding", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/newletter3.pages", import.meta.url)));
  try {
    const title = document.scene.objects.find(object => object.text === "NEWSLETTER");
    const paragraph = title.visual.layout.paragraphs?.[0];
    assert.ok(paragraph?.ruleAbove, "the real title lost its top border");
    assert.ok(paragraph?.ruleBelow, "the real title lost its bottom border");
    assert.equal(paragraph.spaceBefore, 8);
    assert.equal(paragraph.spaceAfter, 8);
    assert.equal(paragraph.ruleAbove.offsetY, -8);
    assert.equal(paragraph.ruleBelow.offsetY, 8);
  } finally { document.close(); core.close(); }
});

test("text layout keeps a wide generated list marker beside its body text", () => {
  const context = {
    font: "",
    letterSpacing: "0px",
    measureText(text) { return { width: text === "❖" ? 44 : text.length * 5 }; },
  };
  const style = {
    fontFamily: "Arial",
    fontSize: 32,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  };

  const layout = layoutTextRuns(context, [
    { ...style, text: "❖\t" },
    { ...style, text: "Item" },
  ], {
    maxWidth: 160,
    lineHeight: 38,
    paragraphs: [{
      align: "start",
      marginLeft: 0,
      marginRight: 0,
      firstLineIndent: 0,
      defaultTabStop: 36,
    }],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x }) => ({ text, x })),
    [{ text: "❖", x: 0 }, { text: "Item", x: 44 }],
  );
});

test("text layout honors a right-to-left paragraph direction", () => {
  const context = recordingContext();
  const runs = [{
    text: "אב גד",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 30,
    lineHeight: 18,
    direction: "rtl",
    align: "start",
  });

  assert.deepEqual(
    layout.runs.map(({ text, x }) => ({ text, x })),
    [{ text: "אב", x: 20 }, { text: " ", x: 15 }, { text: "גד", x: 5 }],
  );
});

test("RTL hanging paragraphs stay inside the right text edge", () => {
  const layout = layoutTextRuns(recordingContext(), [{
    text: "אב", fontFamily: "Arial", fontSize: 12, color: 0x000000ff,
    bold: false, italic: false, underline: false, strikethrough: false,
    highlight: 0, baselineShift: 0, letterSpacing: 0,
  }], {
    maxWidth: 100, lineHeight: 18, direction: "rtl", align: "start",
    paragraphs: [{ align: "start", marginLeft: 56, marginRight: 0,
      firstLineIndent: -56, defaultTabStop: 96, lineHeight: 18,
      spaceBefore: 0, spaceAfter: 0 }],
  });
  assert.equal(Math.max(...layout.runs.map(run => run.x + run.width)), 100);
});

test("text layout keeps authored paragraph metrics across a manual line break", () => {
  const context = recordingContext();
  const layout = layoutTextRuns(context, [{
    text: "AA\u2028BB\nCC",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }], {
    maxWidth: 200,
    lineHeight: 18,
    paragraphs: [{
      align: "start",
      marginLeft: 10,
      marginRight: 4,
      firstLineIndent: -5,
      defaultTabStop: 36,
      spaceAfter: 7,
    }, { align: "start", marginLeft: 50, firstLineIndent: -10, spaceBefore: 9 }],
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, line }) => ({ text, x, line })),
    [
      { text: "AA", x: 5, line: 0 },
      { text: "BB", x: 10, line: 1 },
      { text: "CC", x: 40, line: 2 },
    ],
  );
  assert.equal(layout.runs[1].lineTop - layout.runs[0].lineTop, 18);
  assert.equal(layout.runs[2].lineTop - layout.runs[1].lineTop, 34);
});

test("text layout justifies wrapped lines but leaves the final line unchanged", () => {
  const context = recordingContext();
  const runs = [{
    text: "AA BB CC",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 32,
    lineHeight: 18,
    align: "justify",
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, width, line }) => ({ text, x, width, line })),
    [
      { text: "AA", x: 0, width: 10, line: 0 },
      { text: " ", x: 10, width: 12, line: 0 },
      { text: "BB", x: 22, width: 10, line: 0 },
      { text: " ", x: 32, width: 0, line: 0 },
      { text: "CC", x: 0, width: 10, line: 1 },
    ],
  );
  assert.equal(layout.width, 32);
});

test("distributed alignment uses inter-character spacing instead of ordinary word gaps", () => {
  const context = recordingContext();
  const run = {
    text: "AA BB CC",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  };
  const layout = layoutTextRuns(context, [run], {
    maxWidth: 32,
    lineHeight: 18,
    align: "distribute",
  });

  assert.deepEqual(
    layout.runs.filter(({ line }) => line === 0).map(
      ({ text, x, width, tracking }) => ({ text, x, width, tracking }),
    ),
    [
      { text: "AA", x: 0, width: 13.5, tracking: 1.75 },
      { text: " ", x: 13.5, width: 6.75, tracking: 1.75 },
      { text: "BB", x: 20.25, width: 11.75, tracking: 1.75 },
      { text: " ", x: 32, width: 0, tracking: 0 },
    ],
  );
  assert.deepEqual(
    layout.runs.filter(({ line }) => line === 1).map(
      ({ text, x, width, tracking }) => ({ text, x, width, tracking }),
    ),
    [{ text: "CC", x: 0, width: 32, tracking: 22 }],
  );

  const singlePlacement = layoutTextRuns(context, [{ ...run, text: "AAAA BBB" }], {
    maxWidth: 25,
    lineHeight: 18,
    align: "distribute",
  });
  assert.equal(singlePlacement.runs[0].width, 25);
  assert.equal(Math.abs(singlePlacement.runs[0].tracking - 5 / 3) < 0.0001, true);

  const singleLine = layoutTextRuns(context, [{ ...run, text: "AB" }], {
    maxWidth: 30,
    lineHeight: 18,
    align: "distribute",
  });
  assert.deepEqual(
    singleLine.runs.map(({ text, x, width, tracking }) => ({ text, x, width, tracking })),
    [{ text: "AB", x: 0, width: 30, tracking: 20 }],
  );

  const splitRuns = layoutTextRuns(context, [
    { ...run, text: "甲" },
    { ...run, text: "乙丙 丁", color: 0xff0000ff },
  ], {
    maxWidth: 20,
    lineHeight: 18,
    align: "distribute",
  });
  assert.deepEqual(
    splitRuns.runs
      .filter(({ line }) => line === 0)
      .map(({ text, x }) => ({ text, x })),
    [
      { text: "甲", x: 0 },
      { text: "乙", x: 7.5 },
      { text: "丙", x: 15 },
      { text: " ", x: 20 },
    ],
  );
});

test("Kashida and Thai distribution track only eligible script fragments", () => {
  const context = recordingContext();
  const run = {
    text: "",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  };
  const cases = [
    {
      align: "high-kashida",
      direction: "rtl",
      runs: [{ ...run, text: "ا" }, { ...run, text: "بAB ZZ" }],
      expected: [
        { text: "ا", x: 15, width: 10, tracking: 5 },
        { text: "ب", x: 10, width: 5, tracking: 5 },
        { text: "AB", x: 0, width: 10, tracking: 0 },
      ],
    },
    {
      align: "thai-distribute",
      direction: "ltr",
      runs: [{ ...run, text: "ก" }, { ...run, text: "ขAB ZZ" }],
      expected: [
        { text: "ก", x: 0, width: 10, tracking: 5 },
        { text: "ข", x: 10, width: 5, tracking: 5 },
        { text: "AB", x: 15, width: 10, tracking: 0 },
      ],
    },
  ];

  for (const { align, direction, runs, expected } of cases) {
    const layout = layoutTextRuns(context, runs, {
      maxWidth: 25,
      lineHeight: 18,
      align,
      direction,
    });
    assert.deepEqual(
      layout.runs
        .filter(({ line, text }) => line === 0 && text.trim() !== "")
        .map(({ text, x, width, tracking }) => ({ text, x, width, tracking })),
      expected,
    );
    assert.equal(layout.width, 25);
  }

  const ordinaryJustify = layoutTextRuns(context, [{ ...run, text: "ابABZZ" }], {
    maxWidth: 28,
    lineHeight: 18,
    align: "justify",
    direction: "rtl",
  });
  assert.deepEqual(
    ordinaryJustify.runs
      .filter(({ line }) => line === 0)
      .map(({ text, tracking }) => ({ text, tracking })),
    [
      { text: "ا", tracking: 3 },
      { text: "ب", tracking: 3 },
      { text: "A", tracking: 0 },
      { text: "B", tracking: 0 },
      { text: "Z", tracking: 0 },
    ],
  );
});

test("renderer retains wrapped trailing whitespace without painting its decorations past the edge", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "text:justify-decoration-edge",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 32, height: 40 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "ltr",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 0.1,
          insetLeft: 0,
          insetRight: 0,
          insetTop: 0,
          insetBottom: 0,
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "justify",
          lineHeight: 18,
          runs: [{
            text: "AA BB CC",
            fontFamily: "Arial",
            fontSize: 12,
            color: 0x000000ff,
            bold: false,
            italic: false,
            underline: true,
            strikethrough: false,
            highlight: 0xffee88ff,
            baselineShift: 0,
            letterSpacing: 0,
          }],
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 32, height: 40 },
      { unitIndex: 0 },
    );

    assert.equal(
      context.calls.filter((call) => call[0] === "fillText").map(([, text]) => text).join(""),
      "AA BB CC",
    );
    const highlights = context.calls.filter((call) => call[0] === "fillRect" && call[4] === 18);
    assert.equal(highlights.length, 4);
    assert.equal(highlights.every(([, x, , width]) => x + width <= 32), true);
    const underlineEnds = context.calls.filter((call) => call[0] === "lineTo");
    assert.equal(underlineEnds.length, 4);
    assert.equal(underlineEnds.every(([, x]) => x <= 32), true);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

for (const [script, text, direction, expected] of [
  [
    "CJK",
    "甲乙丙丁戊",
    "ltr",
    [
      { text: "甲", x: 0, line: 0 },
      { text: "乙", x: 6.5, line: 0 },
      { text: "丙", x: 13, line: 0 },
      { text: "丁", x: 0, line: 1 },
      { text: "戊", x: 5, line: 1 },
    ],
  ],
  [
    "Thai",
    "กขคงจ",
    "ltr",
    [
      { text: "ก", x: 0, line: 0 },
      { text: "ข", x: 6.5, line: 0 },
      { text: "ค", x: 13, line: 0 },
      { text: "ง", x: 0, line: 1 },
      { text: "จ", x: 5, line: 1 },
    ],
  ],
  [
    "Arabic RTL",
    "ابتثر",
    "rtl",
    [
      { text: "ا", x: 11.5, line: 0 },
      { text: "ب", x: 5, line: 0 },
      { text: "ت", x: 0, line: 0 },
      { text: "ث", x: 13, line: 1 },
      { text: "ر", x: 8, line: 1 },
    ],
  ],
]) {
  test(`text layout distributes a wrapped ${script} line but leaves the final line unchanged`, () => {
    const context = recordingContext();
    const runs = [{
      text,
      fontFamily: "Arial",
      fontSize: 12,
      color: 0x000000ff,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      highlight: 0,
      baselineShift: 0,
      letterSpacing: 0,
    }];

    const layout = layoutTextRuns(context, runs, {
      maxWidth: 18,
      lineHeight: 18,
      align: "justify",
      direction,
    });

    assert.deepEqual(
      layout.runs.map(({ text: runText, x, line }) => ({ text: runText, x, line })),
      expected,
    );
    assert.equal(layout.width, 18);
  });
}

test("renderer paints Arabic shaping runs instead of isolated graphemes", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const object = sceneObject({
      numericId: 1,
      id: "text:arabic-kashida",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 18, height: 40 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "rtl",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 0.1,
          insetLeft: 0,
          insetRight: 0,
          insetTop: 0,
          insetBottom: 0,
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "high-kashida",
          lineHeight: 18,
          runs: [{
            text: "ابتثر",
            fontFamily: "Arial",
            fontSize: 12,
            color: 0x000000ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baselineShift: 0,
            letterSpacing: 0,
          }],
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 18, height: 40 },
      { unitIndex: 0 },
    );

    assert.deepEqual(
      context.calls.filter((call) => call[0] === "fillText").map(([, text]) => text),
      ["ابت", "ثر"],
    );
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer keeps mixed RTL Kashida paint within the laid-out width", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = recordingContext();
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const baseRun = {
      fontFamily: "Arial",
      fontSize: 12,
      color: 0x000000ff,
      bold: false,
      italic: false,
      underline: true,
      strikethrough: false,
      highlight: 0xffee88ff,
      baselineShift: 0,
      letterSpacing: 0,
    };
    const object = sceneObject({
      numericId: 1,
      id: "text:mixed-arabic-kashida",
      type: "text-box",
      bounds: { x: 0, y: 0, width: 25, height: 40 },
      visual: {
        kind: "text-layout",
        layout: {
          direction: "rtl",
          orientation: "horizontal",
          autoFit: "none",
          verticalAlign: "top",
          defaultTabStop: 36,
          hangingIndent: 0,
          minScale: 0.1,
          insetLeft: 0,
          insetRight: 0,
          insetTop: 0,
          insetBottom: 0,
        },
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "high-kashida",
          lineHeight: 18,
          runs: [
            { ...baseRun, text: "ا" },
            { ...baseRun, text: "بAB ZZ" },
          ],
        },
      },
    });
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    await renderer.render(
      { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 25, height: 40 },
      { unitIndex: 0 },
    );

    assert.deepEqual(
      context.calls
        .filter((call) => call[0] === "fillText" && call[3] === 0 && call[1].trim() !== "")
        .map((call) => ({ text: call[1], x: call[2], letterSpacing: call[6] })),
      [
        { text: "اب", x: 10, letterSpacing: "5px" },
        { text: "AB", x: 0, letterSpacing: "0px" },
      ],
    );
    const firstLineHighlights = context.calls
      .filter((call) => call[0] === "fillRect" && call[2] === 0 && call[4] === 18)
      .map(([, x, , width]) => [x, width])
      .sort((left, right) => left[0] - right[0]);
    assert.deepEqual(firstLineHighlights, [[0, 10], [10, 15]]);
    assert.deepEqual(
      context.calls
        .filter((call) => call[0] === "lineTo" && call[2] < 18)
        .map(([, x]) => x)
        .sort((left, right) => left - right),
      [10, 25],
    );
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("text layout distributes around collapsible edge whitespace but not across a tab stop", () => {
  const context = recordingContext();
  const run = {
    text: "",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  };

  const leading = layoutTextRuns(context, [{ ...run, text: " 甲乙丙 丁" }], {
    maxWidth: 23,
    lineHeight: 18,
    align: "justify",
  });
  assert.deepEqual(
    leading.runs.filter(({ line }) => line === 0).map(({ text, x }) => ({ text, x })),
    [
      { text: " ", x: 0 },
      { text: "甲", x: 5 },
      { text: "乙", x: 11.5 },
      { text: "丙", x: 18 },
      { text: " ", x: 23 },
    ],
  );

  const trailing = layoutTextRuns(context, [{ ...run, text: "甲乙丙 丁" }], {
    maxWidth: 20,
    lineHeight: 18,
    align: "justify",
  });
  assert.deepEqual(
    trailing.runs.filter(({ line }) => line === 0).map(({ text, x, width }) => ({ text, x, width })),
    [
      { text: "甲", x: 0, width: 7.5 },
      { text: "乙", x: 7.5, width: 7.5 },
      { text: "丙", x: 15, width: 5 },
      { text: " ", x: 20, width: 0 },
    ],
  );

  const tabbed = layoutTextRuns(context, [{ ...run, text: "甲\t乙丙" }], {
    maxWidth: 18,
    lineHeight: 18,
    align: "justify",
    defaultTabStop: 10,
  });
  assert.deepEqual(
    tabbed.runs.filter(({ line }) => line === 0).map(({ text, x }) => ({ text, x })),
    [{ text: "甲", x: 0 }, { text: "\t", x: 5 }, { text: "乙", x: 10 }],
  );

  const tabbedWithSpace = layoutTextRuns(context, [{ ...run, text: "AA \tBB CC" }], {
    maxWidth: 32,
    lineHeight: 18,
    align: "justify",
    defaultTabStop: 20,
  });
  assert.deepEqual(
    tabbedWithSpace.runs
      .filter(({ line }) => line === 0)
      .map(({ text, x, width }) => ({ text, x, width })),
    [
      { text: "AA", x: 0, width: 10 },
      { text: " ", x: 10, width: 5 },
      { text: "\t", x: 15, width: 5 },
      { text: "BB", x: 20, width: 10 },
      { text: " ", x: 30, width: 0 },
    ],
  );

  const whitespaceOnly = layoutTextRuns(context, [{ ...run, text: "     X" }], {
    maxWidth: 20,
    lineHeight: 18,
    align: "justify",
  });
  assert.deepEqual(
    whitespaceOnly.runs.map(({ text, x, width, line }) => ({ text, x, width, line })),
    [
      { text: "     ", x: 0, width: 0, line: 0 },
      { text: "X", x: 0, width: 5, line: 1 },
    ],
  );
  assert.equal(whitespaceOnly.width, 5);
});

test("text layout shrink-to-fit returns and measures with the applied font scale", () => {
  const context = {
    font: "",
    letterSpacing: "0px",
    measureText(text) {
      const size = Number(/([\d.]+)px/u.exec(this.font)?.[1] ?? 0);
      return { width: text.length * size };
    },
  };
  const runs = [{
    text: "ABCD",
    fontFamily: "Arial",
    fontSize: 20,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 40,
    maxHeight: 20,
    lineHeight: 24,
    autoFit: "shrink",
    minScale: 0.5,
  });

  assert.equal(Math.abs(layout.scale - 0.5) < 0.0001, true);
  assert.equal(Math.abs(layout.runs[0].width - 40) < 0.001, true);
  assert.equal(Math.abs(layout.height - 12) < 0.001, true);
});

test("text layout fits frame text vertically without wrapping", () => {
  const context = {
    font: "",
    letterSpacing: "0px",
    measureText(text) {
      const size = Number(/([\d.]+)px/u.exec(this.font)?.[1] ?? 0);
      return { width: text.length * size };
    },
  };
  const layout = layoutTextRuns(context, [{
    text: "Fit frame text",
    fontFamily: "Arial",
    fontSize: 100,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }], {
    maxWidth: 400,
    maxHeight: 20,
    lineHeight: 120,
    autoFit: "fit-frame",
    wrap: false,
  });

  assert.equal(Math.abs(layout.scale - 1 / 6) < 0.0001, true);
  assert.equal(Math.abs(layout.height - 20) < 0.001, true);
  assert.equal(layout.lineCount, 1);
});

test("text layout places vertical CJK text in right-to-left columns", () => {
  const context = recordingContext();
  const runs = [{
    text: "甲乙丙",
    fontFamily: "DengXian",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 36,
    maxHeight: 36,
    lineHeight: 12,
    orientation: "vertical-rl",
  });

  assert.deepEqual(
    layout.runs.map(({ text, x, y, line }) => ({ text, x, y, line })),
    [
      { text: "甲", x: 24, y: 0, line: 0 },
      { text: "乙", x: 24, y: 12, line: 0 },
      { text: "丙", x: 24, y: 24, line: 0 },
    ],
  );
  assert.deepEqual({ width: layout.width, height: layout.height, lineCount: layout.lineCount }, {
    width: 12,
    height: 36,
    lineCount: 1,
  });
});

test("stacked WordArt starts each word in a left-to-right column", () => {
  const context = recordingContext();
  const runs = [{
    text: "Vertical Title",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 40,
    maxHeight: 120,
    lineHeight: 12,
    orientation: "stacked-lr",
  });

  assert.equal(layout.lineCount, 2);
  assert.deepEqual(
    layout.runs.filter(({ text }) => text === "V" || text === "T")
      .map(({ text, x, y, line }) => ({ text, x, y, line })),
    [
      { text: "V", x: 0, y: 0, line: 0 },
      { text: "T", x: 12, y: 0, line: 1 },
    ],
  );
});

test("text layout vertically anchors the complete text block", () => {
  const context = recordingContext();
  const runs = [{
    text: "A\nB",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 40,
    maxHeight: 60,
    lineHeight: 20,
    verticalAlign: "center",
  });

  assert.deepEqual(layout.runs.map(({ text, y }) => ({ text, y })), [
    { text: "A", y: 10 },
    { text: "B", y: 30 },
  ]);
  assert.equal(layout.height, 40);
});

test("text layout centers and bottom-aligns overflowing text blocks", () => {
  const context = recordingContext();
  const runs = [{
    text: "A\nB\nC",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];
  const positions = (verticalAlign) => layoutTextRuns(context, runs, {
    maxWidth: 40,
    maxHeight: 40,
    lineHeight: 20,
    verticalAlign,
  }).runs.map(({ text, y }) => ({ text, y }));

  assert.deepEqual(positions("center"), [
    { text: "A", y: -10 },
    { text: "B", y: 10 },
    { text: "C", y: 30 },
  ]);
  assert.deepEqual(positions("bottom"), [
    { text: "A", y: -20 },
    { text: "B", y: 0 },
    { text: "C", y: 20 },
  ]);
});

test("supplied PPTX empty paragraphs retain bottom and center overflow anchors", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/corpus-TextFittingTopBottomMiddle.pptx", import.meta.url)));
  try {
    for (const label of ["BOTTOM", "MIDDLE", "TOP"]) {
      const object = document.scene.objects.find(o => o.text?.includes(label) && o.text.includes("EMPTY PARAGRAPHS"));
      assert.ok(object, label);
      let visual = object.visual;
      while (visual.kind === "layer") visual = visual.visual;
      const options = visual.layout, rich = visual.visual;
      const layout = layoutTextRuns(recordingContext(), rich.runs, {
        ...options, lineHeight: rich.lineHeight, align: rich.align,
        maxWidth: object.bounds.width - options.insetLeft - options.insetRight,
        maxHeight: object.bounds.height - options.insetTop - options.insetBottom,
      });
      const text = layout.runs.find(r => r.text.includes(label));
      assert.ok(text, label);
      // Cached Office reference: bottom text above the frame, center within it,
      // top text below it; the two trailing empty paragraphs belong to the block.
      if (label === "BOTTOM") assert.ok(text.y < 0, `bottom text y=${text.y}`);
      if (label === "MIDDLE") assert.ok(text.y < object.bounds.height / 2, `center text y=${text.y}`);
      if (label === "TOP") assert.ok(text.y > object.bounds.height, `top text y=${text.y}`);
    }
  } finally { document.close(); core.close(); }
});

test("trailing empty paragraphs do not pull visible overflowing text above its box", () => {
  const context = recordingContext();
  const runs = [{
    text: "A\nB\nC\n",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 40,
    maxHeight: 40,
    lineHeight: 20,
    verticalAlign: "center",
  });

  assert.deepEqual(layout.runs.map(({ text, y }) => ({ text, y })), [
    { text: "A", y: 0 },
    { text: "B", y: 20 },
    { text: "C", y: 40 },
  ]);
});

test("a trailing DrawingML line break participates in vertical anchoring", () => {
  const context = recordingContext();
  const layout = layoutTextRuns(context, [{
    text: "A\nB\n",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }], {
    maxWidth: 40,
    maxHeight: 40,
    lineHeight: 20,
    verticalAlign: "center",
    paragraphs: [{
      align: "start",
      marginLeft: 0,
      marginRight: 0,
      firstLineIndent: 0,
      defaultTabStop: 36,
      lineHeight: 20,
    }],
  });

  assert.deepEqual(layout.runs.map(({ text, y }) => ({ text, y })), [
    { text: "A", y: -4.399999999999999 },
    { text: "B", y: 15.600000000000001 },
  ]);
  assert.equal(layout.height, 60);
});

test("vertical text centers a glyph whose line box exceeds the available height", () => {
  const context = recordingContext();
  const layout = layoutTextRuns(context, [{
    text: "甲",
    fontFamily: "DengXian",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }], {
    maxWidth: 24,
    maxHeight: 8,
    lineHeight: 12,
    orientation: "vertical-rl",
    verticalAlign: "center",
  });

  assert.deepEqual(layout.runs.map(({ text, x, y }) => ({ text, x, y })), [
    { text: "甲", x: 12, y: -2 },
  ]);
});

test("text layout applies paragraph-local line height and before and after spacing", () => {
  const context = recordingContext();
  const runs = [{
    text: "First\nSecond\nThird",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 200,
    lineHeight: 16,
    paragraphs: [
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, lineHeight: 20, spaceBefore: 4, spaceAfter: 6 },
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, lineHeight: 24, spaceBefore: 8, spaceAfter: 10 },
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, lineHeight: 18, spaceBefore: 2, spaceAfter: 0 },
    ],
  });

  assert.deepEqual(layout.runs.map(({ text }) => text), ["First", "Second", "Third"]);
  assert.ok(Math.abs(layout.runs[0].y - 9.6) < 0.001);
  assert.ok(Math.abs(layout.runs[1].y - 47.6) < 0.001);
  assert.ok(Math.abs(layout.runs[2].y - 77.6) < 0.001);
  assert.equal(layout.height, 92);
});

test("text layout preserves a small empty paragraph without drawing a list marker", () => {
  const context = recordingContext();
  const runs = [
    {
      text: "Heading\n\n",
      fontFamily: "Arial",
      fontSize: 56 / 3,
      color: 0x000000ff,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      highlight: 0,
      baselineShift: 0,
      letterSpacing: 0,
    },
    {
      text: "Body",
      fontFamily: "Arial",
      fontSize: 40 / 3,
      color: 0x000000ff,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      highlight: 0,
      baselineShift: 0,
      letterSpacing: 0,
    },
  ];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 200,
    paragraphs: [
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, lineHeight: 22.4 },
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, lineHeight: 1.6 },
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, lineHeight: 16 },
    ],
  });

  assert.deepEqual(layout.runs.map(({ text }) => text), ["Heading", "Body"]);
  assert.equal(layout.lineCount, 3);
  assert.ok(Math.abs(layout.runs[1].y - 24) < 0.001);
  assert.ok(Math.abs(layout.height - 40) < 0.001);
});

test("text layout applies autofit line-spacing reduction to paragraph-local line heights", () => {
  const context = recordingContext();
  const runs = [{
    text: "First\nSecond",
    fontFamily: "Arial",
    fontSize: 12,
    color: 0x000000ff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];

  const layout = layoutTextRuns(context, runs, {
    maxWidth: 200,
    lineHeight: 20,
    lineSpacingReduction: 0.25,
    paragraphs: [
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, lineHeight: 20, spaceBefore: 0, spaceAfter: 0 },
      { align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0, defaultTabStop: 36, lineHeight: 24, spaceBefore: 0, spaceAfter: 0 },
    ],
  });

  assert.deepEqual(layout.runs.map(({ text }) => text), ["First", "Second"]);
  assert.ok(Math.abs(layout.runs[0].y - 0.6) < 0.001);
  assert.ok(Math.abs(layout.runs[1].y - 18.6) < 0.001);
  assert.equal(layout.height, 33);
});

test("supplied Word 2007 form keeps rules below text and the seal after the blank", async () => {
  const { registerFonts } = await import("../dist/font.js");
  const { measureSceneFontMetrics, decodeFontMetricTable } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/word-2007-table.docx", import.meta.url));
  const initial = core.open(bytes);
  const form = initial.scene.objects.find(o => o.text?.startsWith("申请人名称："));
  const context = recordingContext();
  context.measureText = function (text) {
    const size = Number(this.font.match(/([\d.]+)px/u)?.[1] ?? 10);
    // Actual local STSong advances: Han = 1 em, space = .25 em.
    return { width: [...text].reduce((w, c) => w + (c === " " ? .25 : 1) * size + parseFloat(this.letterSpacing || "0"), 0),
      fontBoundingBoxAscent: size * (this.textBaseline === "top" ? -.04 : .86),
      fontBoundingBoxDescent: size * (this.textBaseline === "top" ? 1.04 : .14) };
  };
  class LocalFace {
    constructor(family, source) { this.family = family; this.source = source; }
    async load() { if (this.source !== 'local("STSong")') throw Error("unavailable"); return this; }
  }
  const fonts = await registerFonts([], 50, { FontFace: LocalFace, fontSet: { add() {}, delete() {}, check() { return true; } } },
    undefined, [], ["宋体"], { policy: "local-first" });
  const metrics = measureSceneFontMetrics([form], fonts, { createContext: () => context });
  const document = core.open(bytes, metrics.table);
  const object = document.scene.objects.find(o => o.text?.startsWith("申请人名称："));
  const rule = document.scene.objects.find(o => o.source.paragraphIndex === object.source.paragraphIndex && o.type === "shape");
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas { getContext() { return context; } transferToImageBitmap() { return { close() {} }; } }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const frame = await new SceneRenderer([object], DEFAULT_LIMITS, fonts).render(document.scene.info.units[object.unitIndex], { unitIndex: object.unitIndex });
    frame.bitmap.close();
    const draws = context.calls.filter(c => c[0] === "fillText");
    const label = draws.find(c => c[1].startsWith("申请人"));
    assert.ok(label[3] < rule.bounds.y, `baseline ${label[3]} must be above rule ${rule.bounds.y}`);
    const seal = draws.find(c => c[1].includes("盖章"));
    assert.ok(seal && seal[2] > rule.bounds.x + rule.bounds.width,
      `seal must follow the authored rule, got ${JSON.stringify(seal)}`);
    const face = decodeFontMetricTable(metrics.table).faces.find(f => f.family === "宋体" && !f.fontSize);
    assert.equal(face.metrics.find(m => m.codePoint === 32).advanceEm, .5, "pagination and paint retain SimSun's half-em spaces");
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    fonts.close(); document.close(); initial.close(); core.close();
  }
});

test("supplied Word 2006 table centers fixed-line glyphs inside centered cells", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-2006-table.docx", import.meta.url)));
  const context = recordingContext();
  context.measureText = function (text) {
    const size = Number(this.font.match(/([\d.]+)px/u)?.[1] ?? 10);
    return { width: [...text].length * size,
      fontBoundingBoxAscent: size * (this.textBaseline === "top" ? -0.04 : 0.86),
      fontBoundingBoxDescent: size * (this.textBaseline === "top" ? 1.04 : 0.14) };
  };
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    for (const text of ["姓名", "性别", "专科学历", "本科学历"]) {
      const object = document.scene.objects.find(o => o.type === "paragraph" && o.text === text);
      assert.ok(object, text);
      assert.equal(object.visual.layout.fixedLineHeight, true, "preserve the authored exact spacing through the protocol");
      const cell = document.scene.objects.find(o => o.numericId === object.parentNumericId);
      const center = cell.bounds.y + cell.bounds.height / 2;
      assert.ok(Math.abs(object.bounds.y + object.bounds.height / 2 - center) < .01,
        `${text}: the paragraph line box is already centered`);
      const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
      const frame = await renderer.render(document.scene.info.units[0], { unitIndex: 0 });
      frame.bitmap.close();
      const draw = context.calls.find(call => call[0] === "fillText" && call[1] === text);
      assert.ok(draw, text);
      const size = object.visual.visual.runs[0].fontSize;
      const glyphCenter = draw[3] - size * (.86 - .14) / 2;
      assert.ok(Math.abs(glyphCenter - center) < .01,
        `${text}: font box is ${glyphCenter - center}px below the cell center`);
    }
    // Exercise the shared renderer independently of DOCX and table ancestry:
    // two lines, mixed ascents/descents, oversized fonts, and explicit baselines.
    const originalName = document.scene.objects.find(o => o.type === "paragraph" && o.text === "姓名");
    const run = originalName.visual.visual.runs[0];
    context.measureText = function (text) {
      const small = this.font.includes("10px");
      const ascent = small ? 6 : 12;
      const descent = small ? 4 : 2;
      return { width: [...text].length * 5,
        fontBoundingBoxAscent: this.textBaseline === "top" ? -1 : ascent,
        fontBoundingBoxDescent: descent };
    };
    for (const [height, authored, expected, paragraphs] of [
      [24, 0, 16, []], [12, 0, 8, []], [24, 9, 9, []],
      [24, 0, 16, [{ align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0,
        defaultTabStop: 36, lineHeight: 24, spaceBefore: 0, spaceAfter: 0 }]],
      [24, 9, 9, [{ align: "start", marginLeft: 0, marginRight: 0, firstLineIndent: 0,
        defaultTabStop: 36, lineHeight: 24, spaceBefore: 0, spaceAfter: 0 }]],
    ]) {
      context.calls.length = 0;
      const object = { ...originalName, type: "text-box", source: { ...originalName.source, format: "pptx" },
        bounds: { x: 0, y: 0, width: 200, height: height * 2 },
        visual: { ...originalName.visual,
          layout: { ...originalName.visual.layout, textBaseline: authored, paragraphs,
            fixedLineHeight: authored > 0 ? false : originalName.visual.layout.fixedLineHeight },
          visual: { ...originalName.visual.visual, lineHeight: height, runs: [
            { ...run, text: "A", fontSize: 14 }, { ...run, text: "b\n", fontSize: 10 },
            { ...run, text: "C", fontSize: 14 }, { ...run, text: "d", fontSize: 10 },
          ] } } };
      const frame = await new SceneRenderer([object], DEFAULT_LIMITS)
        .render(document.scene.info.units[0], { unitIndex: 0 });
      frame.bitmap.close();
      const draws = context.calls.filter(call => call[0] === "fillText");
      assert.deepEqual(draws.map(call => call[1]), ["A", "b", "C", "d"]);
      assert.deepEqual(draws.map(call => call[3]), [expected, expected, height + expected, height + expected]);
    }
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    document.close(); core.close();
  }
});

test("supplied Word heading retains natural leading below the table rule", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/word-first-page-text.docx", import.meta.url));
  const initial = core.open(bytes);
  const headingText = "控制危险废物越境转移及其处置巴塞尔公约";
  const heading = initial.scene.objects.find((object) => object.text === headingText);
  const context = recordingContext();
  context.measureText = function (text) {
    const size = Number(this.font.match(/([\d.]+)px/u)?.[1] ?? 10);
    return {
      width: [...text].length * size,
      fontBoundingBoxAscent: size * (this.textBaseline === "top" ? -0.04 : 0.86),
      fontBoundingBoxDescent: size * (this.textBaseline === "top" ? 1.04 : 0.14),
    };
  };
  const metrics = measureSceneFontMetrics([heading], { resolve: (family) => ({ family, source: "browser" }) }, { createContext: () => context });
  const measured = core.open(bytes, metrics.table);
  const object = measured.scene.objects.find((candidate) => candidate.text === headingText);
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  try {
    const renderer = new SceneRenderer([object], DEFAULT_LIMITS);
    const frame = await renderer.render(measured.scene.info.units[0], { unitIndex: 0 });
    frame.bitmap.close();
    const draw = context.calls.find((call) => call[0] === "fillText" && call[1] === headingText);
    const size = object.visual.visual.runs[0].fontSize;
    const expected = object.bounds.y + object.bounds.height - size * 0.14;
    assert.ok(Math.abs(draw[3] - expected) < 0.01, `Word places natural leading above the glyph: ${draw[3]} versus ${expected}`);
  } finally {
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
    measured.close();
    initial.close();
    core.close();
  }
});

test("supplied Word tabbed page fragment does not paint into its footnote", async () => {
  const { measureSceneFontMetrics } = await import("../dist/font-metrics.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/word-footnote-page-break.docx", import.meta.url));
  const initial = core.open(bytes);
  const context = { font: "", letterSpacing: "0px", measureText(text) {
    const size = Number(this.font.match(/([\d.]+)px/u)[1]);
    return { width: [...text].reduce((sum, c) => sum + size * (/[\u2e80-\uffff]/u.test(c) ? 1 : .5) + (parseFloat(this.letterSpacing) || 0), 0),
      fontBoundingBoxAscent: .86 * size, fontBoundingBoxDescent: .14 * size };
  } };
  const metrics = measureSceneFontMetrics(initial.scene.objects, { resolve: family => ({ family, source: "browser" }) }, { createContext: () => context });
  const document = core.open(bytes, metrics.table);
  try {
    const german = document.scene.objects.find(o => o.source.paragraphId === "2A0EC19B");
    const germanLayout = layoutTextRuns(context, german.visual.visual.runs, {
      ...german.visual.layout, maxWidth: german.bounds.width, lineHeight: german.visual.visual.lineHeight,
    });
    assert.equal(germanLayout.lineCount, 2, "the shared no-line-start policy also applies to German-language CJK runs");
    assert.ok(Math.abs(german.bounds.height - 2 * german.visual.visual.lineHeight) < .01);
    const chineseLayout = layoutTextRuns(context, german.visual.visual.runs.map(run => ({ ...run, eastAsianLineBreaks: true })), {
      ...german.visual.layout, maxWidth: german.bounds.width, lineHeight: german.visual.visual.lineHeight,
    });
    assert.equal(chineseLayout.lineCount, 2, "Word's Chinese-language control keeps the closing punctuation on line two");

    const fragments = document.scene.objects.filter(o => o.id.startsWith("docx:paragraph:") && o.source.paragraphId === "2F5C5946");
    assert.ok(fragments.length > 0);
    for (const object of fragments) {
      const rendered = layoutTextRuns(context, object.visual.visual.runs, {
        ...object.visual.layout, maxWidth: object.bounds.width, align: object.visual.visual.align, lineHeight: object.visual.visual.lineHeight,
      });
      const height = rendered.lineCount * object.visual.visual.lineHeight;
      assert.ok(height <= object.bounds.height + .01, `Core allocated ${object.bounds.height}px but paints ${height}px: ${object.text}`);
      const separator = document.scene.objects.find(o => o.id === `docx:footnote:${object.unitIndex}:separator`);
      if (separator) assert.ok(object.bounds.y + height <= separator.bounds.y, "painted body clears the note separator");
    }
    const source = initial.scene.objects.filter(o => o.id.startsWith("docx:paragraph:") && o.source.paragraphId === "2F5C5946").map(o => o.text).join("");
    assert.equal(fragments.map(o => o.text).join(""), source, "pagination preserves every source character");
  } finally { document.close(); initial.close(); core.close(); }
});

test("supplied Word footnote separator preserves its paragraph indent and rule geometry", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/word-footnote-page-break.docx", import.meta.url)));
  try {
    const separator = document.scene.objects.find(o => o.id === "docx:footnote:0:separator");
    const note = document.scene.objects.find(o => o.id.startsWith("docx:footnote:") && o.type === "paragraph" && o.unitIndex === 0);
    // Actual Word's PDF rule: x=143.68, width=192, thickness=.64 CSS pixels.
    assert.ok(Math.abs(separator.bounds.x - 143.68) < .2);
    assert.ok(Math.abs(separator.bounds.width - 192) < .01);
    assert.ok(Math.abs(separator.bounds.height - .64) < .05);
    assert.ok(note.bounds.y - separator.bounds.y > 5 && note.bounds.y - separator.bounds.y < 7);
  } finally { document.close(); core.close(); }
});

test("supplied TextDistancesInsets keeps bottom text above an oversized bottom inset", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/text-distances-insets.pptx", import.meta.url)));
  try {
    document.loadUnit(0);
    const object = document.scene.objects.find(object => object.text === "BOTTOM"
      && object.visual.kind === "text-layout" && object.visual.layout.insetBottom > object.bounds.height);
    assert.ok(object, "original file has an oversized authored bottom inset");
    const context = recordingContext();
    const positions = [];
    context.fillText = (text, x, y) => positions.push({ text, y });
    class Canvas { getContext() { return context; } transferToImageBitmap() { return { close() {} }; } }
    Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: Canvas });
    const frame = await new SceneRenderer([object], DEFAULT_LIMITS).render(document.scene.info.units[0], { unitIndex: 0 });
    frame.bitmap.close();
    assert.ok(positions.length > 0);
    assert.ok(positions.every(position => position.y > 0 && position.y < object.bounds.y), "PowerPoint anchors this text above the shape, not inside it");
  } finally {
    document.close(); core.close();
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("supplied ODT thin double border keeps its gap narrower than its ink", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-odf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/odt-border-types.odt", import.meta.url)));
  try {
    const object = document.scene.objects.find(object => object.visual.kind === "stroke-style"
      && object.visual.style.compound === "double" && object.visual.visual.visual.strokeWidth < 1);
    assert.ok(object);
    const context = recordingContext();
    context.getTransform = () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
    const strokes = [];
    context.stroke = function () { strokes.push(this.lineWidth); };
    class Canvas { getContext() { return context; } transferToImageBitmap() { return { close() {} }; } }
    Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: Canvas });
    const frame = await new SceneRenderer([object], DEFAULT_LIMITS).render(document.scene.info.units[0], { unitIndex: 0 });
    frame.bitmap.close();
    assert.equal(strokes.length, 2);
    assert.ok(strokes[1] < strokes[0], "snapping the white gap to the same pixel width erases the border");
  } finally {
    document.close(); core.close();
    if (original) Object.defineProperty(globalThis, "OffscreenCanvas", original);
    else delete globalThis.OffscreenCanvas;
  }
});

test("supplied TextFittingComparison keeps the authored body anchor after a fitting bullet", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/corpus-TextFittingComparisonWithMSO_1.pptx", import.meta.url)));
  try {
    document.loadUnit(1);
    const object = document.scene.objects.find(o => o.unitIndex === 1 && o.text?.includes("ABC abc"));
    let visual = object.visual;
    while (visual.kind === "layer") visual = visual.visual;
    const context = recordingContext();
    context.measureText = text => ({width:[...text].reduce((n,ch)=>n+(ch==='•'?16.8:ch===' '?13.3:24),0)});
    const result = layoutTextRuns(context, visual.visual.runs, {maxWidth:850,lineHeight:52,paragraphs:visual.layout.paragraphs});
    const first = result.runs.find(r=>r.text==='ABC');
    assert.equal(first.x,visual.layout.paragraphs[0].marginLeft, "a bullet fitting before the tab must not push text beyond its authored indent");
  } finally { document.close(); core.close(); }
});

test("supplied XLS pie chart crosses the Wasm protocol with all ten outside labels", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-legacy-office.wasm", import.meta.url)), DEFAULT_LIMITS);
  let document;
  try {
    document = core.open(await readFile(new URL("./fixtures/piechart_outside.xls", import.meta.url)));
    const { objects, info } = document.scene;
    assert.equal(info.units.length, 1);
    assert.ok(collectFontRequests(objects).length > 0, "chart text must produce valid font requests");
    assert.equal(objects.filter(object => object.type === "cell").length, 0);
    assert.ok(objects.some(object => object.type === "group"));
    assert.ok(objects.every(object => object.source.format === "xls" && object.source.kind === "drawing"
      && object.source.stream === "Workbook" && object.source.recordOffset > 0));
    assert.ok(objects.some(object => object.text === "MEDALS per PARTICIPANT"));
    const sectors = objects.filter(object => object.visual.fill?.kind === "linear-gradient");
    assert.equal(sectors.length, 10, "authored gradient fills must survive palette resolution");
    const labels = objects.filter(object => object.text?.includes("%"));
    assert.equal(labels.length, 10);
    for (const [country, percent] of [["Russia", 8], ["United States", 7], ["Norway", 11], ["Canada", 7],
      ["Netherlands", 34], ["Germany", 7], ["Austria", 7], ["France", 7], ["Sweden", 8], ["Switzerland", 4]]) {
      assert.ok(labels.some(object => object.text.includes(country) && object.text.includes(`${percent}%`)), country);
    }
  } finally { document?.close(); core.close(); }
});


test("supplied ODT list prefix uses the authored tab and continuation indent", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-odf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/file-sample_1MB.odt", import.meta.url)));
  try {
    const bullet = document.scene.objects.find(o => o.text?.startsWith("Maecenas non lorem"));
    const {layout, visual} = bullet.visual;
    const result = layoutTextRuns(recordingContext(), visual.runs, { ...layout,
      lineHeight: visual.lineHeight, align: visual.align, maxWidth: bullet.bounds.width });
    const marker = result.runs.find(run => run.text === "•");
    const body = result.runs.find(run => run.text.startsWith("Maecenas"));
    assert.equal(marker.x, 24);
    assert.equal(body.x, 48);
  } finally { document.close(); core.close(); }
});
