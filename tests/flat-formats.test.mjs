import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";

function installCanvas() {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const calls = [];
  const styles = [];
  const context = new Proxy({ font: "", fillStyle: "", textAlign: "start" }, {
    get(target, property) {
      if (property === "measureText") return (text) => ({ width: [...text].length * 8 });
      if (property === "fillText") {
        return (text) => {
          calls.push(String(text));
          styles.push({ text: String(text), font: target.font, fillStyle: target.fillStyle });
        };
      }
      if (property in target) return target[property];
      return () => {};
    },
    set(target, property, value) {
      target[property] = value;
      return true;
    },
  });
  class RecordingCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", {
    configurable: true,
    value: RecordingCanvas,
  });
  return {
    calls,
    styles,
    restore() {
      if (descriptor) Object.defineProperty(globalThis, "OffscreenCanvas", descriptor);
      else delete globalThis.OffscreenCanvas;
    },
  };
}

async function createEngine(options = {}) {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  return createOfficeEngine({ wasm, execution: "inline", fontPolicy: "local-first", ...options });
}

test("keeps plain-text documents unsupported instead of treating them as TXT", async () => {
  const engine = await createEngine();
  try {
    const utf8Bom = new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode("Plain title\r\n第二行")]);
    const utf16Le = new Uint8Array([0xff, 0xfe, 0x50, 0x00, 0x6c, 0x00, 0x61, 0x00, 0x69, 0x00, 0x6e, 0x00]);
    const utf16Be = new Uint8Array([0xfe, 0xff, 0x00, 0x50, 0x00, 0x6c, 0x00, 0x61, 0x00, 0x69, 0x00, 0x6e]);
    for (const bytes of [
      new TextEncoder().encode("Plain title\nSecond line"),
      utf8Bom,
      utf16Le,
      utf16Be,
      new TextEncoder().encode("2 < 3 and 5 > 4"),
    ]) {
      await assert.rejects(
        engine.open(bytes),
        (error) => error?.code === "UNSUPPORTED_FORMAT",
      );
    }
  } finally {
    engine.close();
  }
});

test("opens RFC 4180 CSV bytes with quoted commas, quotes, and embedded newlines", async () => {
  const bytes = new TextEncoder().encode(
    'name,notes,value\r\nAda,"line 1\r\nline 2",42\r\n"Lin, Q","quote ""ok""",7\r\n',
  );
  const canvas = installCanvas();
  const engine = await createEngine();
  let document;
  let frame;
  try {
    document = await engine.open(bytes);
    assert.equal(document.info.format, "csv");
    assert.equal(document.info.kind, "spreadsheet");
    assert.equal(document.info.units.length, 1);
    assert.deepEqual(
      {
        type: document.info.units[0].type,
        name: document.info.units[0].name,
        rows: document.info.units[0].rows,
        columns: document.info.units[0].columns,
        frozenRows: document.info.units[0].frozenRows,
        frozenColumns: document.info.units[0].frozenColumns,
      },
      {
        type: "sheet",
        name: "Sheet1",
        rows: 3,
        columns: 3,
        frozenRows: 0,
        frozenColumns: 0,
      },
    );

    const cells = await document.listObjects({ types: ["cell"] });
    assert.equal(cells.length, 9);
    const multiline = cells.find(({ source }) => source.row === 1 && source.column === 1);
    const quoted = cells.find(({ source }) => source.row === 2 && source.column === 0);
    assert.equal(multiline?.text, "line 1\nline 2");
    assert.deepEqual(multiline?.source, {
      format: "csv",
      part: "input.csv",
      kind: "cell",
      row: 1,
      column: 1,
      textRange: [0, 13],
      mapping: "exact",
    });
    assert.equal(quoted?.text, "Lin, Q");

    const hits = await document.hitTest({
      unitIndex: 0,
      x: multiline.bounds.x + 2,
      y: multiline.bounds.y + 2,
    });
    assert.equal(hits[0]?.object.id, multiline.id);
    frame = await document.render({ unitIndex: 0 });
    assert.match(canvas.calls.join(""), /line 1/);
    assert.match(canvas.calls.join(""), /quote "ok"/);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    canvas.restore();
  }
});

test("keeps HTML documents unsupported instead of misidentifying markup as CSV", async () => {
  const engine = await createEngine();
  try {
    const fixtures = [
      "<!doctype html><html><body><p>Quarterly report</p></body></html>",
      "\uFEFF \n<HTML><HEAD><TITLE>Report</TITLE></HEAD><BODY>Body</BODY></HTML>",
      "<table><tr><td>a,b</td></tr>\n<tr><td>c,d</td></tr></table>",
    ];
    for (const fixture of fixtures) {
      await assert.rejects(
        engine.open(new TextEncoder().encode(fixture)),
        (error) => error?.code === "UNSUPPORTED_FORMAT",
      );
    }
  } finally {
    engine.close();
  }
});

test("opens and renders the supplied Cocoa RTF report with escaped paragraph breaks", async () => {
  const canvas = installCanvas();
  const engine = await createEngine();
  let document;
  let frame;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/report-cocoa.rtf", import.meta.url)));
    assert.equal(document.info.format, "rtf");
    assert.equal(document.info.units.length, 1);
    const paragraphs = await document.listObjects({ textOnly: true });
    assert.equal(paragraphs.length, 8);
    assert.deepEqual(paragraphs.slice(0, 2).map(({ text }) => text), [
      "Simple Home Styling", "Easy Decorating",
    ]);
    assert.match(paragraphs.at(-1).text, /turning off Document Body in the Document controls\.$/);
    frame = await document.render({ unitIndex: 0 });
    const paintedText = canvas.calls.join("");
    assert.match(paintedText, /Easy Decorating/);
    assert.match(paintedText, /Document controls\./);
    assert.ok(!document.diagnostics().some(({ severity }) => severity === "fatal"));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    canvas.restore();
  }
});

test("opens RTF paragraphs with Unicode and common character formatting", async () => {
  const bytes = new TextEncoder().encode(String.raw`{\rtf1\ansi\ansicpg1252\deff0
{\fonttbl{\f0 Arial;}{\f1 Times New Roman;}}
{\colortbl;\red192\green0\blue0;\red0\green0\blue255;}
\viewkind4\uc1
\pard\f0\fs32 Plain {\b bold} {\i italic}\par
\pard\cf1\fs28 Red \u20320?\u22909?\line next\par
}`);
  const canvas = installCanvas();
  const engine = await createEngine();
  let document;
  let frame;
  try {
    document = await engine.open(bytes);
    assert.equal(document.info.format, "rtf");
    assert.equal(document.info.kind, "text");
    const paragraphs = await document.listObjects({ textOnly: true });
    assert.deepEqual(paragraphs.map(({ text }) => text), [
      "Plain bold italic",
      "Red 你好\nnext",
    ]);
    assert.deepEqual(paragraphs[1].source, {
      format: "rtf",
      part: "input.rtf",
      kind: "paragraph",
      paragraphIndex: 1,
      textRange: [0, 11],
      mapping: "exact",
    });

    frame = await document.render({ unitIndex: 0 });
    const bold = canvas.styles.find(({ text }) => text === "bold");
    const italic = canvas.styles.find(({ text }) => text === "italic");
    const red = canvas.styles.find(({ text }) => text.includes("Red"));
    assert.match(bold?.font ?? "", /700/);
    assert.match(italic?.font ?? "", /italic/);
    assert.equal(red?.fillStyle, "rgba(192, 0, 0, 1)");
    assert.deepEqual(document.diagnostics().map(({ code }) => code), ["FONT_METRICS_APPLIED"]);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    canvas.restore();
  }
});

test("renders unresolved RTF compatibility fonts like Word without changing the authored family", async () => {
  const bytes = new TextEncoder().encode(String.raw`{\rtf1\ansi
{\fonttbl{\f0 DejaVu Sans;}}
\pard\f0\fs24 Compatibility\par}`);
  const canvas = installCanvas();
  const engine = await createEngine();
  let document;
  let frame;
  try {
    document = await engine.open(bytes);
    const [paragraph] = await document.listObjects({ textOnly: true });
    assert.deepEqual(paragraph.fontRuns, [{
      start: 0,
      end: 13,
      authoredFamily: "DejaVu Sans",
      renderedFamily: "Verdana",
      source: "fallback",
      fontSize: 16,
    }]);
    frame = await document.render({ unitIndex: 0 });
    assert.match(canvas.styles.find(({ text }) => text === "Compatibility")?.font ?? "", /Verdana/);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    canvas.restore();
  }
});

test("keeps preferred RTF pictures source-mapped and honors authored page geometry", async () => {
  const png = "89504e470d0a1a0a0000000d4948445200000001000000010804000000b51c0c020000000b4944415478da6364f80f00010501012718e3660000000049454e44ae426082";
  const bytes = new TextEncoder().encode(String.raw`{\rtf1\ansi
\paperw11906\paperh16838\margl1134\margr1134\margt1134\margb1134
\pard Before\par
{\*\shppict{\pict\picw1\pich1\picwgoal1500\pichgoal750\pngblip
${png}}}{\nonshppict{\pict\picw1\pich1\picwgoal1500\pichgoal750\wmetafile8
00}}\par
{\shp{\*\shpinst\shpleft150\shpright3150\shptop300\shpbottom1800
{\pict\picw1\pich1\picwgoal1500\pichgoal750\pngblip
${png}}}}\par
After\par}`);
  const engine = await createEngine();
  let document;
  try {
    document = await engine.open(bytes);
    assert.equal(document.info.format, "rtf");
    assert.ok(Math.abs(document.info.units[0].width - 11906 / 15) < 0.001);
    assert.ok(Math.abs(document.info.units[0].height - 16838 / 15) < 0.001);
    const images = await document.listObjects({ types: ["image"] });
    assert.equal(images.length, 2);
    assert.deepEqual(images.map(({ source }) => source), [
      {
        format: "rtf",
        part: "input.rtf",
        kind: "picture",
        pictureIndex: 0,
        mapping: "exact",
      },
      {
        format: "rtf",
        part: "input.rtf",
        kind: "picture",
        pictureIndex: 1,
        mapping: "exact",
      },
    ]);
    const expectedBounds = [
      { x: 1134 / 15, width: 1500 / 15, height: 750 / 15 },
      { x: (1134 + 150) / 15, width: (3150 - 150) / 15, height: (1800 - 300) / 15 },
    ];
    for (const [index, expected] of expectedBounds.entries()) {
      for (const field of ["x", "width", "height"]) {
        assert.ok(Math.abs(images[index].bounds[field] - expected[field]) < 0.001);
      }
    }
    assert.equal(
      document.diagnostics().some(({ code }) => code === "UNSUPPORTED_FEATURE"),
      false,
    );
  } finally {
    document?.close();
    engine.close();
  }
});

test("lays out RTF table rows as bounded cells instead of concatenated text", async () => {
  const bytes = new TextEncoder().encode(String.raw`{\rtf1\ansi
\paperw11906\paperh16838\margl1134\margr1134\margt1134\margb1134
\trowd\trleft60\trrh300\cellx600\cellx1800
\pard\intbl\fs24 Alpha\cell
\pard\intbl\fs24 Beta\cell\row
\pard After\par}`);
  const canvas = installCanvas();
  const engine = await createEngine();
  let document;
  let frame;
  try {
    document = await engine.open(bytes);
    const cells = await document.listObjects({ types: ["cell"] });
    assert.deepEqual(cells.map(({ text }) => text), ["Alpha", "Beta"]);
    assert.ok(Math.abs(cells[0].bounds.x - (1134 + 60) / 15) < 0.001);
    assert.ok(Math.abs(cells[0].bounds.width - (600 - 60) / 15) < 0.001);
    assert.ok(Math.abs(cells[1].bounds.width - (1800 - 600) / 15) < 0.001);
    assert.equal(cells[0].bounds.height, cells[1].bounds.height);
    assert.ok(cells[0].bounds.height >= 300 / 15);
    frame = await document.render({ unitIndex: 0 });
    assert.match(canvas.calls.join(""), /Alpha/);
    assert.match(canvas.calls.join(""), /Beta/);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    canvas.restore();
  }
});

test("keeps PDF and legacy compound Office bytes unsupported", async () => {
  const engine = await createEngine();
  try {
    const fixtures = [
      new TextEncoder().encode("%PDF-1.7\n1 0 obj\n<< /Type /Catalog >>\nendobj\n%%EOF"),
      Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0),
    ];
    for (const bytes of fixtures) {
      await assert.rejects(
        engine.open(bytes),
        (error) => error?.code === "UNSUPPORTED_FORMAT",
      );
    }
  } finally {
    engine.close();
  }
});

test("rejects malformed CSV and RTF instead of silently rendering damaged input", async () => {
  const engine = await createEngine();
  try {
    await assert.rejects(
      engine.open(new TextEncoder().encode('a,b\r\n"unterminated,value')),
      (error) => error?.code === "FORMAT_INVALID",
    );
    await assert.rejects(
      engine.open(new TextEncoder().encode(String.raw`{\rtf1\ansi unclosed`)),
      (error) => error?.code === "FORMAT_INVALID",
    );
  } finally {
    engine.close();
  }
});

test("enforces flat-format object and nesting limits", async () => {
  const fixtures = [
    {
      limits: { documentObjects: 4 },
      bytes: new TextEncoder().encode("a,b,c\r\n1,2,3\r\n"),
      code: "OBJECT_LIMIT",
    },
    {
      limits: { xmlDepth: 2 },
      bytes: new TextEncoder().encode(String.raw`{\rtf1{\b{\i too deep}}}`),
      code: "XML_DEPTH_LIMIT",
    },
  ];
  for (const fixture of fixtures) {
    const engine = await createEngine({ limits: fixture.limits });
    try {
      await assert.rejects(
        engine.open(fixture.bytes),
        (error) => error?.code === fixture.code,
      );
    } finally {
      engine.close();
    }
  }
});
