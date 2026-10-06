import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";
import { resolveLimits } from "../dist/core.js";
import { decodeSnapshot } from "../dist/protocol.js";
import { byteFingerprint, bytesEqual } from "../dist/bytes.js";

const encoder = new TextEncoder();
const workerSource = await readFile(new URL("../src/worker.ts", import.meta.url), "utf8");

test("shared binary resource keys preserve full-byte hashing and equality", () => {
  for (const length of [0, 1, 4095, 4096, 65537]) {
    const bytes = Uint8Array.from({ length: length + 3 }, (_, index) => index * 31).subarray(3);
    let first = 0x811c9dc5;
    let second = 0x9e3779b9;
    for (const byte of bytes) {
      first = Math.imul(first ^ byte, 0x01000193);
      second = Math.imul(second ^ byte, 0x85ebca6b);
    }
    assert.equal(byteFingerprint(bytes), `${length}:${first >>> 0}:${second >>> 0}`);
    assert.ok(bytesEqual(bytes, bytes.slice()));
    assert.equal(bytesEqual(bytes, new Uint8Array(length + 1)), false);
    for (const index of new Set([0, Math.floor(length / 2), length - 1])) {
      if (length === 0) continue;
      const different = bytes.slice();
      different[index] ^= 1;
      assert.equal(bytesEqual(bytes, different), false);
    }
  }
});

test("document queries and renders share one Worker operation scheduler", () => {
  assert.match(workerSource, /function enqueueDocumentOperation/);
  assert.match(workerSource, /await enqueueDocumentOperation\(async \(\) =>/);
  assert.match(workerSource, /void enqueueDocumentOperation\(\(\) => handle\(request\)\)/);
  assert.doesNotMatch(workerSource, /\n  void handle\(request\);/);
});

class SnapshotWriter {
  bytes = [];

  u8(value) {
    this.bytes.push(value & 0xff);
  }

  u16(value) {
    this.u8(value);
    this.u8(value >>> 8);
  }

  u32(value) {
    this.u16(value);
    this.u16(value >>> 16);
  }

  i32(value) {
    this.u32(value >>> 0);
  }

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

  byteArray(value) {
    this.u32(value.length);
    this.bytes.push(...value);
  }

  absentString() {
    this.u32(0xffff_ffff);
  }
}

test("font alternate metadata rejects unbounded counts and invalid names", () => {
  for (const [count, family, names] of [
    [1025, "Authored", ["Courier New"]],
    [1, "", ["Courier New"]],
    [1, "x".repeat(129), ["Courier New"]],
    [1, "Authored", []],
    [1, "Authored", Array(65).fill("Courier New")],
    [1, "Authored", ["bad\u0000name"]],
  ]) {
    const writer = new SnapshotWriter();
    writer.u32(0x3144564f); writer.u16(76);
    for (const value of [1, 0, 0, 0, 0, 0]) writer.u8(value);
    for (let index = 0; index < 4; index += 1) writer.u32(0);
    writer.u32(count); writer.string(family); writer.u32(names.length);
    for (const name of names) writer.string(name);
    writer.u32(0);
    assert.throws(() => decodeSnapshot(Uint8Array.from(writer.bytes)),
      error => error?.code === "CORE_PROTOCOL_INVALID");
  }
});

function repeatedImageSnapshot(bytes, crop = [0, 0, 0, 0]) {
  const writer = new SnapshotWriter();
  writer.u32(0x3144564f);
  writer.u16(3);
  writer.u8(0);
  writer.u8(1);
  writer.u8(1);
  writer.u8(0);
  writer.u8(0);
  writer.u8(0);
  writer.u32(0);
  writer.u32(1);
  writer.u8(1);
  writer.u8(0);
  writer.u8(0);
  writer.u8(0);
  writer.u32(0);
  writer.string("unit:0");
  writer.string("Slide 1");
  writer.f32(960);
  writer.f32(720);
  writer.u32(0);
  writer.u32(0);
  writer.u32(2);
  for (let index = 0; index < 2; index += 1) {
    writer.u32(index + 1);
    writer.i32(-1);
    writer.u32(0);
    writer.u8(4);
    writer.u8(3);
    writer.u8(0);
    writer.u8(0);
    writer.i32(index);
    writer.f32(index * 10);
    writer.f32(0);
    writer.f32(10);
    writer.f32(10);
    writer.string(`image:${index}`);
    writer.absentString();
    writer.absentString();
    writer.string("ppt/media/image.png");
    writer.string("shape");
    writer.u8(1);
    writer.string("shapeId");
    writer.u8(1);
    writer.u32(index + 1);
    for (const value of crop) writer.f32(value);
    writer.string("image/png");
    writer.byteArray(bytes);
  }
  return Uint8Array.from(writer.bytes);
}

test("snapshot decoding interns repeated large binary resources", () => {
  const bytes = new Uint8Array(4_096).fill(7);
  const scene = decodeSnapshot(repeatedImageSnapshot(bytes));
  assert.equal(scene.objects[0].visual.bytes, scene.objects[1].visual.bytes);
});

test("snapshot decoding accepts an f32 crop that removes every source pixel", () => {
  const scene = decodeSnapshot(repeatedImageSnapshot(Uint8Array.of(1), [0.6, 0, 0.4, 0]));

  assert.equal(Math.fround(scene.objects[0].visual.cropLeft + scene.objects[0].visual.cropRight), 1);
});

function sourceSnapshot(formatCode, documentKindCode, sourceKind, fields, unitType = 1) {
  const writer = new SnapshotWriter();
  writer.u32(0x3144564f);
  writer.u16(3);
  writer.u8(0);
  writer.u8(formatCode);
  writer.u8(documentKindCode);
  writer.u8(0);
  writer.u8(0);
  writer.u8(0);
  writer.u32(0);
  writer.u32(1);
  writer.u8(unitType);
  writer.u8(0);
  writer.u8(0);
  writer.u8(0);
  writer.u32(0);
  writer.string("unit:0");
  writer.string("Slide 1");
  writer.f32(960);
  writer.f32(720);
  writer.u32(0);
  writer.u32(0);
  writer.u32(1);

  writer.u32(1);
  writer.i32(-1);
  writer.u32(0);
  writer.u8(2);
  writer.u8(0);
  writer.u8(0);
  writer.u8(0);
  writer.i32(0);
  writer.f32(1);
  writer.f32(2);
  writer.f32(3);
  writer.f32(4);
  writer.string("object:1");
  writer.absentString();
  writer.absentString();
  writer.string("content.xml");
  writer.string(sourceKind);
  writer.u8(fields.length);
  for (const [key, value] of fields) {
    writer.string(key);
    if (typeof value === "string") {
      writer.u8(0);
      writer.string(value);
    } else {
      writer.u8(1);
      writer.u32(value);
    }
  }
  return Uint8Array.from(writer.bytes);
}

function pptxSceneSnapshot({
  kindCode = 1,
  units = [{ type: 1, index: 0, id: "unit:0", width: 960, height: 720 }],
  objects = [{ numericId: 1, id: "object:1", unitIndex: 0 }],
} = {}) {
  const writer = new SnapshotWriter();
  writer.u32(0x3144564f);
  writer.u16(3);
  writer.u8(0);
  writer.u8(1);
  writer.u8(kindCode);
  writer.u8(0);
  writer.u8(0);
  writer.u8(0);
  writer.u32(0);
  writer.u32(units.length);
  for (const unit of units) {
    writer.u8(unit.type);
    writer.u8(0);
    writer.u8(0);
    writer.u8(0);
    writer.u32(unit.index);
    writer.string(unit.id);
    writer.string("Unit");
    writer.f32(unit.width);
    writer.f32(unit.height);
    writer.u32(unit.rows ?? 0);
    writer.u32(unit.columns ?? 0);
  }
  writer.u32(objects.length);
  for (const object of objects) {
    writer.u32(object.numericId);
    writer.i32(object.parentNumericId ?? -1);
    writer.u32(object.unitIndex);
    writer.u8(2);
    writer.u8(0);
    writer.u8(0);
    writer.u8(0);
    writer.i32(0);
    writer.f32(object.x ?? 1);
    writer.f32(object.y ?? 2);
    writer.f32(object.width ?? 3);
    writer.f32(object.height ?? 4);
    writer.string(object.id);
    if (object.parentId === undefined) writer.absentString();
    else writer.string(object.parentId);
    writer.absentString();
    writer.string("ppt/slides/slide1.xml");
    writer.string("shape");
    writer.u8(1);
    writer.string("shapeId");
    writer.u8(1);
    writer.u32(object.numericId);
  }
  return Uint8Array.from(writer.bytes);
}

function uleb(value) {
  const bytes = [];
  do {
    let byte = value & 0x7f;
    value >>>= 7;
    if (value !== 0) byte |= 0x80;
    bytes.push(byte);
  } while (value !== 0);
  return bytes;
}

function wasmSection(id, payload) {
  return [id, ...uleb(payload.length), ...payload];
}

function wasmName(value) {
  const bytes = encoder.encode(value);
  return [...uleb(bytes.length), ...bytes];
}

function abiModule(abiVersion) {
  const names = [
    "ov_core_abi_version",
    "ov_alloc",
    "ov_free",
    "ov_document_open",
    "ov_document_open_with_font_metrics",
    "ov_document_snapshot",
    "ov_result_pointer",
    "ov_result_clear",
    "ov_document_hit_test",
    "ov_document_load_unit",
    "ov_document_load_unit_region",
    "ov_document_load_all",
    "ov_document_requires_calculation",
    "ov_document_apply_calculation",
    "ov_document_close",
  ];
  const exportPayload = [16, ...wasmName("memory"), 2, 0];
  for (let index = 0; index < names.length; index += 1) {
    exportPayload.push(...wasmName(names[index]), 0, ...uleb(index));
  }

  const constant = (value) => [0x41, value < 0 ? value & 0x7f : value];
  const bodies = [
    constant(abiVersion),
    constant(8),
    constant(0),
    constant(1),
    constant(1),
    constant(16),
    [0x23, 0],
    constant(0),
    constant(0),
    constant(0),
    constant(0),
    constant(0),
    constant(0),
    constant(0),
    [0x41, 0x78, 0x24, 0, 0x41, 0],
  ];
  const codePayload = [bodies.length];
  for (const instructions of bodies) {
    const body = [0, ...instructions, 0x0b];
    codePayload.push(...uleb(body.length), ...body);
  }

  return Uint8Array.from([
    0, 0x61, 0x73, 0x6d, 1, 0, 0, 0,
    ...wasmSection(1, [1, 0x60, 0, 1, 0x7f]),
    ...wasmSection(3, [15, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    ...wasmSection(5, [1, 0, 1]),
    ...wasmSection(6, [1, 0x7f, 1, 0x41, 0, 0x0b]),
    ...wasmSection(7, exportPayload),
    ...wasmSection(10, codePayload),
  ]);
}

test("PPTX and ODP source locators preserve row, column, and text ranges", () => {
  const pptx = decodeSnapshot(sourceSnapshot(1, 1, "shape", [
    ["shapeId", 7],
    ["row", 2],
    ["column", 3],
    ["rangeStart", 4],
    ["rangeEnd", 9],
  ]));
  assert.deepEqual(pptx.objects[0].source, {
    format: "pptx",
    part: "content.xml",
    kind: "shape",
    shapeId: 7,
    row: 2,
    column: 3,
    textRange: [4, 9],
    mapping: "exact",
  });

  const odp = decodeSnapshot(sourceSnapshot(2, 1, "element", [
    ["path", "/office:body/draw:page[1]/draw:frame[2]"],
    ["elementId", "frame-2"],
    ["row", 5],
    ["column", 6],
  ]));
  assert.deepEqual(odp.objects[0].source, {
    format: "odp",
    part: "content.xml",
    kind: "element",
    elementId: "frame-2",
    path: "/office:body/draw:page[1]/draw:frame[2]",
    row: 5,
    column: 6,
    mapping: "exact",
  });
});

test("optional native packs preserve legacy Office and iWork source locators", () => {
  const ppt = decodeSnapshot(sourceSnapshot(11, 1, "text", [
    ["stream", "PowerPoint Document"],
    ["recordOffset", 128],
    ["rangeStart", 0],
    ["rangeEnd", 12],
  ])).objects[0].source;
  assert.deepEqual(ppt, {
    format: "ppt",
    part: "content.xml",
    kind: "text",
    stream: "PowerPoint Document",
    recordOffset: 128,
    textRange: [0, 12],
    mapping: "exact",
  });

  const xls = decodeSnapshot(sourceSnapshot(12, 2, "cell", [
    ["stream", "Workbook"],
    ["recordOffset", 256],
    ["row", 3],
    ["column", 4],
  ], 2)).objects[0].source;
  assert.deepEqual(xls, {
    format: "xls",
    part: "content.xml",
    kind: "cell",
    stream: "Workbook",
    recordOffset: 256,
    row: 3,
    column: 4,
    mapping: "exact",
  });

  const doc = decodeSnapshot(sourceSnapshot(13, 3, "paragraph", [
    ["stream", "WordDocument"],
    ["recordOffset", 512],
    ["rangeStart", 10],
    ["rangeEnd", 20],
  ], 3)).objects[0].source;
  assert.deepEqual(doc, {
    format: "doc",
    part: "content.xml",
    kind: "paragraph",
    stream: "WordDocument",
    recordOffset: 512,
    textRange: [10, 20],
    mapping: "exact",
  });

  for (const [formatCode, format, kindCode, unitType] of [
    [14, "pages", 3, 3],
    [15, "numbers", 2, 2],
    [16, "keynote", 1, 1],
  ]) {
    const source = decodeSnapshot(sourceSnapshot(formatCode, kindCode, "preview", [
      ["component", "preview.jpg"],
    ], unitType)).objects[0].source;
    assert.deepEqual(source, {
      format,
      part: "content.xml",
      kind: "preview",
      component: "preview.jpg",
      mapping: "exact",
    });
  }

  const keynoteSlide = decodeSnapshot(sourceSnapshot(16, 1, "slide-preview", [
    ["component", "slide-node-42"],
  ])).objects[0].source;
  assert.deepEqual(keynoteSlide, {
    format: "keynote",
    part: "content.xml",
    kind: "slide-preview",
    component: "slide-node-42",
    mapping: "exact",
  });

  const numbersCell = decodeSnapshot(sourceSnapshot(15, 2, "cell", [
    ["component", "Index/Tile-12.iwa#B3"],
  ], 2)).objects[0].source;
  assert.deepEqual(numbersCell, {
    format: "numbers",
    part: "content.xml",
    kind: "cell",
    component: "Index/Tile-12.iwa#B3",
    mapping: "exact",
  });

  for (const kind of [
    "slide-background",
    "shape",
    "image",
    "media",
    "text-box",
    "table",
    "table-grid",
    "table-cell",
    "chart",
    "chart-series",
  ]) {
    const source = decodeSnapshot(sourceSnapshot(16, 1, kind, [
      ["component", "archive-42"],
    ])).objects[0].source;
    assert.deepEqual(source, {
      format: "keynote",
      part: "content.xml",
      kind,
      component: "archive-42",
      mapping: "exact",
    });
  }

  for (const kind of [
    "body",
    "inline-image",
    "table",
    "table-grid",
    "table-cell",
    "chart",
    "chart-series",
  ]) {
    const source = decodeSnapshot(sourceSnapshot(14, 3, kind, [
      ["component", "archive-42"],
    ], 3)).objects[0].source;
    assert.deepEqual(source, {
      format: "pages",
      part: "content.xml",
      kind,
      component: "archive-42",
      mapping: "exact",
    });
  }

  const pdf = decodeSnapshot(sourceSnapshot(17, 3, "text", [
    ["objectNumber", 42],
    ["byteOffset", 128],
  ], 3)).objects[0].source;
  assert.deepEqual(pdf, {
    format: "pdf",
    part: "content.xml",
    kind: "text",
    objectNumber: 42,
    byteOffset: 128,
    mapping: "exact",
  });

  const annotation = decodeSnapshot(sourceSnapshot(17, 3, "annotation", [
    ["objectNumber", 43],
  ], 3)).objects[0].source;
  assert.deepEqual(annotation, {
    format: "pdf",
    part: "content.xml",
    kind: "annotation",
    objectNumber: 43,
    mapping: "exact",
  });
});

test("optional native source locators reject cross-family kinds", () => {
  assert.throws(
    () => decodeSnapshot(sourceSnapshot(11, 1, "cell", [["stream", "PowerPoint Document"]])),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /ppt source locator cell/u.test(error.message),
  );
  assert.throws(
    () => decodeSnapshot(sourceSnapshot(14, 3, "paragraph", [["component", "preview.jpg"]], 3)),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /pages source locator paragraph/u.test(error.message),
  );
  assert.throws(
    () => decodeSnapshot(sourceSnapshot(13, 3, "paragraph", [["stream", "Workbook"]], 3)),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /doc source locator paragraph/u.test(error.message),
  );
});

test("legacy PowerPoint drawing locators preserve their stream offset", () => {
  const source = decodeSnapshot(sourceSnapshot(11, 1, "drawing", [
    ["stream", "PowerPoint Document"],
    ["recordOffset", 640],
  ])).objects[0].source;
  assert.deepEqual(source, {
    format: "ppt",
    part: "content.xml",
    kind: "drawing",
    stream: "PowerPoint Document",
    recordOffset: 640,
    mapping: "exact",
  });
});

test("retired HTML snapshot format code is rejected instead of being reused", () => {
  assert.throws(
    () => decodeSnapshot(sourceSnapshot(9, 3, "paragraph", [["index", 0]])),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /format code 9/u.test(error.message),
  );
});

test("retired TXT snapshot format code is rejected instead of being reused", () => {
  assert.throws(
    () => decodeSnapshot(sourceSnapshot(7, 3, "paragraph", [["index", 0]], 3)),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /format code 7/u.test(error.message),
  );
});

test("corrupt required source fields are rejected instead of defaulted", () => {
  assert.throws(
    () => decodeSnapshot(sourceSnapshot(1, 1, "shape", [])),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /shapeId/u.test(error.message),
  );
  assert.throws(
    () => decodeSnapshot(sourceSnapshot(2, 1, "element", [["elementId", "frame-2"]])),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /path/u.test(error.message),
  );
});

test("core snapshots reject invalid unit topology and geometry", () => {
  const invalid = [
    pptxSceneSnapshot({ units: [{ type: 1, index: 1, id: "unit:1", width: 960, height: 720 }] }),
    pptxSceneSnapshot({ units: [{ type: 2, index: 0, id: "unit:0", width: 960, height: 720, rows: 1, columns: 1 }] }),
    pptxSceneSnapshot({ units: [{ type: 1, index: 0, id: "unit:0", width: 0, height: 720 }] }),
    pptxSceneSnapshot({ units: [{ type: 1, index: 0, id: "unit:0", width: 960, height: Number.NaN }] }),
  ];
  for (const snapshot of invalid) {
    assert.throws(() => decodeSnapshot(snapshot), (error) => error?.code === "CORE_PROTOCOL_INVALID");
  }
});

test("core snapshots reject invalid object identity, bounds, and parent topology", () => {
  const validParent = { numericId: 1, id: "parent", unitIndex: 0 };
  const invalid = [
    pptxSceneSnapshot({ objects: [validParent, { numericId: 1, id: "child", unitIndex: 0 }] }),
    pptxSceneSnapshot({ objects: [validParent, { numericId: 2, id: "parent", unitIndex: 0 }] }),
    pptxSceneSnapshot({ objects: [{ numericId: 1, id: "missing-unit", unitIndex: 1 }] }),
    pptxSceneSnapshot({ objects: [{ numericId: 1, id: "negative-size", unitIndex: 0, width: -1 }] }),
    pptxSceneSnapshot({ objects: [{ numericId: 1, id: "infinite", unitIndex: 0, width: Number.POSITIVE_INFINITY }] }),
    pptxSceneSnapshot({ objects: [{ numericId: 1, id: "missing-parent", unitIndex: 0, parentNumericId: 99, parentId: "none" }] }),
    pptxSceneSnapshot({ objects: [validParent, { numericId: 2, id: "child", unitIndex: 0, parentNumericId: 1, parentId: "wrong" }] }),
    pptxSceneSnapshot({ objects: [
      { numericId: 1, id: "first", unitIndex: 0, parentNumericId: 2, parentId: "second" },
      { numericId: 2, id: "second", unitIndex: 0, parentNumericId: 1, parentId: "first" },
    ] }),
  ];
  for (const snapshot of invalid) {
    assert.throws(() => decodeSnapshot(snapshot), (error) => error?.code === "CORE_PROTOCOL_INVALID");
  }
});

test("core snapshots preserve objects positioned outside the unit origin", () => {
  const scene = decodeSnapshot(pptxSceneSnapshot({
    objects: [{ numericId: 1, id: "partially-off-canvas", unitIndex: 0, x: -12, y: -8 }],
  }));

  assert.deepEqual(scene.objects[0].bounds, { x: -12, y: -8, width: 3, height: 4 });
});

test("invalid UTF-8 in a core snapshot is a structured protocol error", () => {
  const snapshot = pptxSceneSnapshot();
  const marker = encoder.encode("unit:0");
  let offset = -1;
  for (let index = 0; index <= snapshot.length - marker.length; index += 1) {
    if (marker.every((byte, markerIndex) => snapshot[index + markerIndex] === byte)) {
      offset = index;
      break;
    }
  }
  assert.notEqual(offset, -1);
  snapshot[offset] = 0xff;

  assert.throws(
    () => decodeSnapshot(snapshot),
    (error) => error?.code === "CORE_PROTOCOL_INVALID" && /UTF-8/u.test(error.message),
  );
});

test("the JS adapter rejects an incompatible core ABI", async () => {
  await assert.rejects(
    createOfficeEngine({ execution: "inline", wasm: abiModule(1) }),
    (error) => error?.code === "CORE_ABI_MISMATCH",
  );
});

test("snapshot failure closes its handle and rejects out-of-range ABI pointers", async () => {
  const engine = await createOfficeEngine({ execution: "inline", wasm: abiModule(11) });
  try {
    await assert.rejects(
      engine.open(new Uint8Array([1])),
      (error) => error?.code === "CORE_PROTOCOL_INVALID" && /empty snapshot/u.test(error.message),
    );
    await assert.rejects(
      engine.open(new Uint8Array([1])),
      (error) => error?.code === "CORE_PROTOCOL_INVALID" && /invalid memory range/u.test(error.message),
    );
  } finally {
    engine.close();
  }
});

test("resource limit overrides reject unknown keys and return an immutable policy", () => {
  assert.throws(
    () => resolveLimits({ renderPixel: 1 }),
    (error) => error?.code === "INVALID_LIMITS" && /renderPixel/u.test(error.message),
  );

  const limits = resolveLimits({ renderPixels: 1_000 });
  assert.equal(limits.renderPixels, 1_000);
  assert.equal(Object.isFrozen(limits), true);
});

test("Wasm acquisition failures use the structured core load error", async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => {
    throw new TypeError("blocked by policy");
  };
  try {
    await assert.rejects(
      createOfficeEngine({ execution: "inline", wasm: "https://invalid.example/core.wasm" }),
      (error) => error?.code === "CORE_LOAD_FAILED"
        && error.cause instanceof TypeError
        && /load the Wasm core/u.test(error.message),
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});
