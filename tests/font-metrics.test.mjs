import assert from "node:assert/strict";
import test from "node:test";

import {
  decodeFontMetricTable,
  fontMetricLayoutObjects,
  measureSceneFontMetrics,
} from "../dist/font-metrics.js";
import { fontShorthand } from "../dist/font.js";

function richTextObject(runs) {
  return {
    id: "text-1",
    numericId: 1,
    type: "paragraph",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 100, height: 20 },
    z: 0,
    source: {
      format: "docx",
      part: "word/document.xml",
      kind: "paragraph",
      mapping: "derived",
    },
    visual: {
      kind: "rich-text",
      geometry: "rectangle",
      fill: { kind: "none" },
      stroke: { kind: "none" },
      strokeWidth: 0,
      align: "start",
      lineHeight: 16,
      runs,
    },
  };
}

function run(text, overrides = {}) {
  return {
    text,
    fontFamily: "Authored Sans",
    fontSize: 12,
    color: 0xff,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
    ...overrides,
  };
}

function sfntWithLineGap(lineGap, unitsPerEm = 1_000) {
  const bytes = new Uint8Array(12 + 3 * 16 + 54 + 10 + 74);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x0001_0000);
  view.setUint16(4, 3);
  let offset = 12 + 3 * 16;
  for (const [index, tag, length] of [
    [0, 0x6865_6164, 54],
    [1, 0x6868_6561, 10],
    [2, 0x4f53_2f32, 74],
  ]) {
    const record = 12 + index * 16;
    view.setUint32(record, tag);
    view.setUint32(record + 8, offset);
    view.setUint32(record + 12, length);
    if (tag === 0x6865_6164) view.setUint16(offset + 18, unitsPerEm);
    if (tag === 0x6868_6561) view.setInt16(offset + 8, lineGap);
    if (tag === 0x4f53_2f32) view.setInt16(offset + 72, lineGap);
    offset += length;
  }
  return bytes.buffer;
}

test("presentation font relayout measures table descendants but skips ordinary slide text", () => {
  const ordinary = { ...richTextObject([run("ordinary")]), numericId: 1, id: "ordinary", source: { format: "pptx" } };
  const table = { ...richTextObject([]), numericId: 2, id: "table", type: "table", source: { format: "pptx" } };
  const cell = { ...richTextObject([run("cell")]), numericId: 3, id: "cell", type: "cell", parentNumericId: 2, source: { format: "pptx" } };
  const paragraph = { ...richTextObject([run("nested")]), numericId: 4, id: "nested", parentNumericId: 3, source: { format: "pptx" } };
  const objects = [ordinary, table, cell, paragraph];

  assert.deepEqual(fontMetricLayoutObjects(objects, "pptx").map((object) => object.id), ["table", "cell", "nested"]);
  assert.deepEqual(fontMetricLayoutObjects(objects, "odp").map((object) => object.id), ["table", "cell", "nested"]);
  assert.equal(fontMetricLayoutObjects([ordinary], "pptx").length, 0);
  assert.equal(fontMetricLayoutObjects([ordinary], "odp").length, 0);
  assert.equal(fontMetricLayoutObjects(objects, "docx"), objects);
  assert.equal(fontMetricLayoutObjects(objects, "doc"), objects);
  assert.equal(fontMetricLayoutObjects(objects, "pages"), objects);
  assert.equal(fontMetricLayoutObjects(objects, "odt"), objects);
});

test("symbol metrics measure the same glyph as rendering for Unicode and legacy fonts", () => {
  for (const source of ["host", "fallback"]) {
    for (const legacy of [true, false]) {
      const context = {
        font: "",
        measureText(text) {
          const size = Number(this.font.match(/([\d.]+)px/u)[1]);
          const em = text === "•" || (legacy && source === "host" && text === "\uf0b7") ? 0.35 : 0.8;
          return { width: size * em };
        },
      };
      const result = measureSceneFontMetrics(
        [richTextObject([run("\uf0b7", { fontFamily: "Symbol" })])],
        { resolve() { return { family: "Resolved Symbol", source }; } },
        { createContext: () => context },
      );
      assert.deepEqual(decodeFontMetricTable(result.table).faces[0].metrics,
        [{ codePoint: 0xf0b7, advanceEm: Math.fround(0.35) }], `${source}, legacy=${legacy}`);
    }
  }
});

test("retains size-specific advances without changing scalable faces or vertical metrics", () => {
  const context = { font: "", measureText(text) {
    const size = Number(this.font.match(/([\d.]+)px/u)[1]);
    return { width: size * (text === 'W' && size === 12 ? .91 : .9),
      fontBoundingBoxAscent: size * .8, fontBoundingBoxDescent: size * .2 };
  } };
  const result = measureSceneFontMetrics([richTextObject([
    run('Wi', { fontSize: 12 }), run('Wi', { fontSize: 24 }),
  ])], { resolve: family => ({ family, source: 'browser' }) }, { createContext: () => context });
  const { faces } = decodeFontMetricTable(result.table);
  assert.equal(faces.length, 2, 'one scalable face and only the genuinely different 12px override');
  assert.equal(faces[0].fontSize, undefined);
  assert.equal(faces[0].verticalMetrics.ascentEm, Math.fround(.8));
  assert.equal(faces[1].fontSize, 12);
  assert.deepEqual(faces[1].metrics, [{ codePoint: 87, advanceEm: Math.fround(.91) }]);
  assert.equal(decodeFontMetricTable(result.table).version, 3);
  for (const size of [-1, Number.NaN, Number.POSITIVE_INFINITY]) {
    const corrupt = result.table.slice();
    new DataView(corrupt.buffer).setFloat32(20 + 28, size, true);
    assert.throws(() => decodeFontMetricTable(corrupt), { code: "INVALID_FONT_METRIC_TABLE" });
  }
  const limited = measureSceneFontMetrics([richTextObject([run('Wi')])],
    { resolve: family => ({ family, source: 'browser' }) },
    { createContext: () => context, maxBytes: 20 + 28 + 'Authored Sans'.length + 16 });
  assert.equal(decodeFontMetricTable(limited.table).version, 2, 'a tight budget retains the base metrics');
  assert.equal(limited.metricCount, 2);
  assert.ok(limited.diagnostics.some(diagnostic => diagnostic.code === 'FONT_METRICS_LIMIT'));
});

test("keeps scalable wide and narrow advances compact after checking the actual size", () => {
  const fonts = [];
  const context = {
    font: "",
    measureText(text) {
      fonts.push(this.font);
      return { width: Number(this.font.match(/([\d.]+)px/u)[1]) * (text === "W" ? .9 : .2) };
    },
  };
  const resolver = {
    resolve() {
      return { family: "Resolved Sans", source: "host" };
    },
  };

  const result = measureSceneFontMetrics(
    [richTextObject([run("Wi")])],
    resolver,
    { createContext: () => context },
  );
  const decoded = decodeFontMetricTable(result.table);

  assert.deepEqual(result.diagnostics, []);
  assert.equal(decoded.version, 2);
  assert.deepEqual(decoded.faces, [{
    family: "Authored Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    metrics: [
      { codePoint: 0x57, advanceEm: Math.fround(0.9) },
      { codePoint: 0x69, advanceEm: Math.fround(0.2) },
    ],
  }]);
  assert.deepEqual(fonts, [
    "400 1000px \"Resolved Sans\"",
    "400 1000px \"Resolved Sans\"",
    "400 12px \"Resolved Sans\"",
    "400 12px \"Resolved Sans\"",
  ]);
});

test("encodes browser vertical face metrics for DOCX line layout", () => {
  const result = measureSceneFontMetrics(
    [richTextObject([run("Hg国")])],
    {
      resolve: () => ({ family: "Resolved Sans", source: "host" }),
      imageCodecFonts: () => [{
        family: "Resolved Sans",
        bytes: sfntWithLineGap(500),
        style: "normal",
        weight: 400,
        stretch: "normal",
      }],
    },
    {
      createContext: () => ({
        font: "",
        measureText(text) {
          return {
            width: text.length * Number(this.font.match(/([\d.]+)px/u)[1]) * .5,
            fontBoundingBoxAscent: 920,
            fontBoundingBoxDescent: 260,
          };
        },
      }),
    },
  );

  const decoded = decodeFontMetricTable(result.table);
  assert.equal(decoded.version, 2);
  assert.deepEqual(decoded.faces[0].verticalMetrics, {
    ascentEm: Math.fround(0.92),
    descentEm: Math.fround(0.26),
    lineGapEm: Math.fround(0.5),
  });
});

test("measures mixed-script fallback scalars with their actual resolved faces", () => {
  const fonts = [];
  const result = measureSceneFontMetrics(
    [richTextObject([run("A中B", { fontFamily: "等线" })])],
    {
      resolve: () => ({ family: "Calibri", source: "fallback" }),
      resolveFace(face) {
        return {
          ...face,
          family: face.codePoints?.some((codePoint) => codePoint === 0x4e2d)
            ? "Hiragino Sans"
            : "Calibri",
          source: "fallback",
        };
      },
    },
    {
      createContext: () => ({
        font: "",
        measureText() {
          fonts.push(this.font);
          return { width: Number(this.font.match(/([\d.]+)px/u)[1]) * (/^400 [\d.]+px "Hiragino Sans"/u.test(this.font) ? 1 : .5) };
        },
      }),
    },
  );

  const face = decodeFontMetricTable(result.table).faces[0];
  assert.equal(face.fallback, true);
  assert.deepEqual(face.metrics, [
    { codePoint: 0x41, advanceEm: Math.fround(0.5) },
    { codePoint: 0x42, advanceEm: Math.fround(0.5) },
    { codePoint: 0x4e2d, advanceEm: Math.fround(1) },
  ]);
  assert.ok(fonts.some((font) => font.includes("\"Calibri\"")));
  assert.ok(fonts.some((font) => font.includes("\"Hiragino Sans\"")));
});

test("deduplicates Unicode scalars by authored face while measuring the resolved face", () => {
  const fonts = [];
  const context = {
    font: "",
    measureText(text) {
      fonts.push(this.font);
      return { width: Number(this.font.match(/([\d.]+)px/u)[1]) * (text.codePointAt(0) === 0x1f600 ? 1.25 : .5) };
    },
  };
  const resolver = {
    resolve() {
      throw new Error("resolveFace should be used when it is available");
    },
    resolveFace() {
      return {
        family: "Exact Face",
        style: "oblique",
        weight: 625,
        stretch: "condensed",
      };
    },
  };

  const result = measureSceneFontMetrics([
    richTextObject([run("😀A😀A", { bold: true, italic: true })]),
  ], resolver, { createContext: () => context });
  const decoded = decodeFontMetricTable(result.table);

  assert.equal(result.faceCount, 1);
  assert.equal(result.metricCount, 2);
  assert.deepEqual(decoded.faces, [{
    family: "Authored Sans",
    style: "italic",
    weight: 700,
    stretch: "normal",
    metrics: [
      { codePoint: 0x41, advanceEm: Math.fround(0.5) },
      { codePoint: 0x1f600, advanceEm: Math.fround(1.25) },
    ],
  }]);
  assert.deepEqual(fonts, [
    "oblique 625 condensed 1000px \"Exact Face\"",
    "oblique 625 condensed 1000px \"Exact Face\"",
    "oblique 625 condensed 12px \"Exact Face\"",
    "oblique 625 condensed 12px \"Exact Face\"",
  ]);
});

test("font metrics traverse DrawingML style, effect, and image wrappers", () => {
  const object = richTextObject([run("Wrapped", {
    fontFamily: "Wrapped Sans",
    bold: true,
    italic: true,
  })]);
  object.visual = {
    kind: "text-effects",
    effects: [],
    visual: {
      kind: "stroke-style",
      style: {
        cap: "flat",
        join: "miter",
        compound: "single",
        alignment: "center",
        miterLimit: 4,
        dash: [],
      },
      visual: {
        kind: "advanced-effect",
        visual: {
          kind: "image-color-change",
          from: 0xffffffff,
          to: 0x00000000,
          useAlpha: true,
          visual: object.visual,
        },
      },
    },
  };
  const result = measureSceneFontMetrics(
    [object],
    { resolve: () => ({ family: "Resolved Sans", source: "host" }) },
    { createContext: () => ({ font: "", measureText() { return { width: Number(this.font.match(/([\d.]+)px/u)[1]) * .5 }; } }) },
  );

  assert.deepEqual(decodeFontMetricTable(result.table).faces, [{
    family: "Wrapped Sans",
    style: "italic",
    weight: 700,
    stretch: "normal",
    metrics: [
      { codePoint: 0x57, advanceEm: Math.fround(0.5) },
      { codePoint: 0x61, advanceEm: Math.fround(0.5) },
      { codePoint: 0x64, advanceEm: Math.fround(0.5) },
      { codePoint: 0x65, advanceEm: Math.fround(0.5) },
      { codePoint: 0x70, advanceEm: Math.fround(0.5) },
      { codePoint: 0x72, advanceEm: Math.fround(0.5) },
    ],
  }]);
});

test("encodes the same compact table regardless of object and scalar order", () => {
  const context = {
    font: "",
    measureText(text) {
      return { width: text.codePointAt(0) * Number(this.font.match(/([\d.]+)px/u)[1]) / 1000 };
    },
  };
  const resolver = {
    resolve(family) {
      return { family: `${family} Resolved`, source: "host" };
    },
  };
  const first = richTextObject([run("zy", { fontFamily: "Zulu" })]);
  const second = {
    ...richTextObject([run("ba", { fontFamily: "Alpha", bold: true })]),
    id: "text-2",
    numericId: 2,
  };

  const forward = measureSceneFontMetrics([first, second], resolver, {
    createContext: () => context,
  });
  const reverse = measureSceneFontMetrics([
    { ...second, visual: { ...second.visual, runs: [run("ab", { fontFamily: "Alpha", bold: true })] } },
    { ...first, visual: { ...first.visual, runs: [run("yz", { fontFamily: "Zulu" })] } },
  ], resolver, { createContext: () => context });

  assert.deepEqual(forward.table, reverse.table);
  // 20-byte header + 2 v2 face records + UTF-8 names + 4 metrics.
  assert.equal(forward.table.byteLength, 117);
  assert.deepEqual(decodeFontMetricTable(forward.table).faces.map(({ family }) => family), [
    "Alpha",
    "Zulu",
  ]);
});

test("keeps distinct authored lookup keys when faces share one registered alias", () => {
  const fonts = [];
  const context = {
    font: "",
    measureText(text) {
      fonts.push(this.font);
      return { width: Number(this.font.match(/([\d.]+)px/u)[1]) * (text === "A" ? .6 : .7) };
    },
  };
  const result = measureSceneFontMetrics([
    richTextObject([
      run("A", { fontFamily: "First Authored" }),
      run("B", { fontFamily: "Second Authored" }),
    ]),
  ], {
    resolve() {
      return { family: "Shared Registered Alias", source: "host" };
    },
  }, { createContext: () => context });

  const faces = decodeFontMetricTable(result.table).faces;
  assert.deepEqual(faces.map(({ family, metrics }) => ({ family, codePoint: metrics[0].codePoint })), [
    { family: "First Authored", codePoint: 0x41 },
    { family: "Second Authored", codePoint: 0x42 },
  ]);
  assert.deepEqual(fonts, [
    "400 1000px \"Shared Registered Alias\"",
    "400 1000px \"Shared Registered Alias\"",
    "400 12px \"Shared Registered Alias\"",
    "400 12px \"Shared Registered Alias\"",
  ]);
});

test("layout metrics use the same Office-compatible fallback shorthand as paint", () => {
  const fonts = [];
  measureSceneFontMetrics(
    [richTextObject([run("AB")])],
    { resolve: () => ({ family: "Calibri", source: "fallback" }) },
    {
      createContext: () => ({
        font: "",
        measureText() {
          fonts.push(this.font);
          return { width: 500 };
        },
      }),
    },
  );

  assert.deepEqual(fonts, [
    fontShorthand("Calibri", 1_000, false, false),
    fontShorthand("Calibri", 1_000, false, false),
    fontShorthand("Calibri", 12, false, false),
    fontShorthand("Calibri", 12, false, false),
  ]);
});

test("merges authored family casing that maps to one Wasm lookup key", () => {
  const result = measureSceneFontMetrics([
    richTextObject([
      run("A", { fontFamily: "Case Sans" }),
      run("B", { fontFamily: "case sans" }),
    ]),
  ], {
    resolve: () => ({ family: "Shared Alias", source: "host" }),
  }, {
    createContext: () => ({ font: "", measureText() { return { width: Number(this.font.match(/([\d.]+)px/u)[1]) * .5 }; } }),
  });

  assert.deepEqual(decodeFontMetricTable(result.table).faces, [{
    family: "Case Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    metrics: [
      { codePoint: 0x41, advanceEm: Math.fround(0.5) },
      { codePoint: 0x42, advanceEm: Math.fround(0.5) },
    ],
  }]);
});

test("returns a valid empty table and approximate diagnostic without Canvas", () => {
  const result = measureSceneFontMetrics(
    [richTextObject([run("No canvas")])],
    { resolve: (family) => ({ family, source: "browser" }) },
    { createContext: () => undefined },
  );

  assert.equal(result.faceCount, 0);
  assert.equal(result.metricCount, 0);
  assert.deepEqual(decodeFontMetricTable(result.table).faces, []);
  assert.deepEqual(result.diagnostics, [{
    code: "FONT_METRICS_UNAVAILABLE",
    severity: "warning",
    fidelity: "approximate",
    phase: "layout",
    message: "OffscreenCanvas 2D metrics are unavailable; text layout remains approximate",
    details: { requestedFaceCount: 2, requestedMetricCount: 16 },
  }]);
});

test("omits invalid Canvas advances and strictly rejects a corrupted metric", () => {
  const context = {
    font: "",
    measureText(text) {
      if (text === "X") return { width: Number.NaN };
      if (text === "Y") return { width: -1 };
      return { width: 0 };
    },
  };
  const result = measureSceneFontMetrics(
    [richTextObject([run("XY\u0301")])],
    { resolve: (family) => ({ family, source: "host" }) },
    { createContext: () => context },
  );

  assert.deepEqual(decodeFontMetricTable(result.table).faces[0].metrics, [{
    codePoint: 0x301,
    advanceEm: 0,
  }]);
  assert.deepEqual(result.diagnostics[0], {
    code: "FONT_METRIC_INVALID",
    severity: "warning",
    fidelity: "approximate",
    phase: "layout",
    message: "Canvas returned invalid font advances; affected characters remain approximate",
    details: { invalidMetricCount: 2, requestedMetricCount: 6, metricCount: 1 },
  });

  const corrupted = result.table.slice();
  const view = new DataView(corrupted.buffer);
  const metricsOffset = view.getUint32(16, true);
  view.setFloat32(metricsOffset + 4, Number.NaN, true);
  assert.throws(
    () => decodeFontMetricTable(corrupted),
    (error) => error?.code === "INVALID_FONT_METRIC_TABLE",
  );
});

test("bounds metric materialization before measuring a hostile demand set", () => {
  let calls = 0;
  const result = measureSceneFontMetrics(
    [richTextObject([run("ABC")])],
    { resolve: (family) => ({ family, source: "host" }) },
    {
      createContext: () => ({
        font: "",
        measureText() {
          calls += 1;
          return { width: 500 };
        },
      }),
      maxBytes: 69,
    },
  );

  assert.equal(result.table.byteLength, 69);
  assert.equal(result.metricCount, 1);
  assert.equal(calls, 1);
  assert.equal(result.diagnostics[0].code, "FONT_METRICS_LIMIT");
  assert.deepEqual(decodeFontMetricTable(result.table).faces[0].metrics, [
    { codePoint: 0x41, advanceEm: Math.fround(0.5) },
  ]);
});

for (const source of ["browser", "fallback"]) test(`supplied Word date line fits with measured ${source} glyph advances`, async () => {
  const { readFile } = await import("node:fs/promises");
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/word-first-page-text.docx", import.meta.url));
  const initial = core.open(bytes);
  const date = initial.scene.objects.find((object) => object.text?.startsWith("2021"));
  // Word's Times New Roman digits/asterisk advance 0.5em; the date's CJK glyphs advance 1em.
  const metrics = measureSceneFontMetrics([date], { resolve: (family) => ({ family, source }) }, {
    createContext: () => ({
      font: "",
      measureText(text) {
        const size = Number(this.font.match(/([\d.]+)px/u)[1]);
        return { width: [...text].reduce((width, character) => width + (/^[0-9∗]$/u.test(character) ? 0.5 : 1), 0) * size };
      },
    }),
  });
  const measured = core.open(bytes, metrics.table);
  try {
    const result = measured.scene.objects.find((object) => object.text?.startsWith("2021"));
    assert.equal(result.text, "2021年7月26日至30日和2022年6月6日至17日，日内瓦∗");
    assert.ok(result.bounds.height < 20, `date must be one 10pt line, got ${result.bounds.height}px`);
  } finally {
    measured.close();
    initial.close();
    core.close();
  }
});

test("supplied Word body keeps its natural line box when local font leading is unknown", async () => {
  const { readFile } = await import("node:fs/promises");
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const bytes = await readFile(new URL("./fixtures/word-numbered-body.docx", import.meta.url));
  const initial = core.open(bytes);
  try {
    for (const source of ["browser", "fallback"]) for (const [lineGap, ascent, expected] of [[undefined, 860, 1.31], [0, 860, 1.15], [500, 860, 1.5], [undefined, 1460, 1.6]]) {
      const metrics = measureSceneFontMetrics(initial.scene.objects, {
        resolve: family => ({ family, source }),
        imageCodecFonts: () => lineGap === undefined ? [] : [...new Set(initial.scene.objects.flatMap(o => o.visual?.visual?.runs?.map(r => r.fontFamily) ?? []))].map(family => ({
          family, bytes: sfntWithLineGap(lineGap), style: "normal", weight: 400, stretch: "normal",
        })),
      }, { createContext: () => ({ font: "", measureText(text) {
        return { width: Number(this.font.match(/([\d.]+)px/u)[1]) * (/[\u2e80-\uffff]/u.test(text) ? 1 : .5), fontBoundingBoxAscent: ascent, fontBoundingBoxDescent: 140 };
      } }) });
      assert.equal(decodeFontMetricTable(metrics.table).faces[0].verticalMetrics.lineGapUnknown === true, lineGap === undefined);
      assert.equal(metrics.diagnostics.some(d => d.code === "FONT_METRIC_VERTICAL_APPROXIMATE"), lineGap === undefined);
      const measured = core.open(bytes, metrics.table);
      try {
        const paragraph = measured.scene.objects.find(o => o.id === "docx:paragraph:0:fragment:0");
        assert.ok(Math.abs(paragraph.visual.visual.lineHeight - 40 / 3 * expected) < .001,
          `original paragraph: source=${source}, gap=${lineGap}, ascent=${ascent}, height=${paragraph.visual.visual.lineHeight}`);
        assert.equal(paragraph.text, initial.scene.objects.find(o => o.id === paragraph.id).text);
      } finally { measured.close(); }
    }
  } finally { initial.close(); core.close(); }
});

test("ODT compatible fallback measures the same font used for painting", () => {
  const object = richTextObject([run("AB", {fontFamily: "DejaVu Sans"})]);
  object.source.format = "odt";
  for (const source of ["fallback", "host"]) {
    const fonts = [];
    measureSceneFontMetrics([object], {resolve: () => ({family: "Resolved Face", source})}, {
      createContext: () => ({font: "", measureText() { fonts.push(this.font); return {width: 500}; }}),
    });
    assert.ok(fonts.length > 0);
    assert.ok(fonts.every(font => font.includes(source === "fallback" ? "Verdana" : "Resolved Face")));
  }
});
