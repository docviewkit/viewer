import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { readZipEntries } from "../scripts/accuracy-metamorphic.mjs";
import { gzipSync } from "node:zlib";

import { decodeOfficeImagePayload } from "../dist/image-codecs.js";

const PNG_1X1 = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
  0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
]);

function pngHeader(width, height) {
  const bytes = PNG_1X1.slice();
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width, false);
  view.setUint32(20, height, false);
  return bytes;
}

function emfPlusRecord(type, flags, data = new Uint8Array()) {
  const size = (12 + data.length + 3) & ~3;
  const bytes = new Uint8Array(size);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, type, true);
  view.setUint16(2, flags, true);
  view.setUint32(4, size, true);
  view.setUint32(8, data.length, true);
  bytes.set(data, 12);
  return bytes;
}

function emfPlusTextRecords(text) {
  const family = new Uint8Array(Buffer.from("A", "utf16le"));
  const font = new Uint8Array(28);
  const fontView = new DataView(font.buffer);
  fontView.setFloat32(4, 12, true);
  fontView.setUint32(20, 1, true);
  font.set(family, 24);
  const encoded = new Uint8Array(Buffer.from(text, "utf16le"));
  const draw = new Uint8Array((28 + encoded.length + 3) & ~3);
  const drawView = new DataView(draw.buffer);
  drawView.setUint32(0, 0xff000000, true);
  drawView.setUint32(8, text.length, true);
  draw.set(encoded, 28);
  return [
    emfPlusRecord(0x4008, (6 << 8) | 1, font),
    emfPlusRecord(0x401c, 0x8000 | 1, draw),
  ];
}

function emfPlusRasterObject(id, image) {
  const data = new Uint8Array(28 + image.length);
  const view = new DataView(data.buffer);
  view.setUint32(4, 1, true); // ImageDataTypeBitmap
  view.setInt32(8, 1, true);
  view.setInt32(12, 1, true);
  view.setInt32(16, 4, true);
  view.setUint32(20, 0x0026_200a, true);
  view.setUint32(24, 2, true); // BitmapDataTypeCompressed
  data.set(image, 28);
  return emfPlusRecord(0x4008, (5 << 8) | id, data);
}

function emfPlusMetafileObject(id, metafile) {
  const data = new Uint8Array(16 + metafile.length);
  const view = new DataView(data.buffer);
  view.setUint32(4, 2, true); // ImageDataTypeMetafile
  view.setUint32(8, 3, true); // MetafileDataTypeEmfPlusDual
  view.setUint32(12, metafile.length, true);
  data.set(metafile, 16);
  return emfPlusRecord(0x4008, (5 << 8) | id, data);
}

function emfPlusContinuedMetafileObject(id, metafile) {
  const object = emfPlusMetafileObject(id, metafile).subarray(12);
  const first = new Uint8Array(9);
  new DataView(first.buffer).setUint32(0, object.length, true);
  first.set(object.subarray(0, 5), 4);
  return [
    emfPlusRecord(0x4008, 0x8000 | (5 << 8) | id, first),
    emfPlusRecord(0x4008, 0x8000 | (5 << 8) | id, object.subarray(5, 7)),
    emfPlusRecord(0x4008, (5 << 8) | id, object.subarray(7)),
  ];
}

function emfPlusDrawImage(id) {
  const data = new Uint8Array(40);
  const view = new DataView(data.buffer);
  view.setFloat32(24, 0, true);
  view.setFloat32(28, 0, true);
  view.setFloat32(32, 1, true);
  view.setFloat32(36, 1, true);
  return emfPlusRecord(0x401a, id, data);
}

function makeEmfRecords(records, bounds = [0, 0, 1, 1], frame = [0, 0, 26, 26]) {
  const recordsLength = records.reduce((total, record) => total + record.length, 0);
  const totalSize = 88 + recordsLength + 20;
  const bytes = new Uint8Array(totalSize);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 1, true);
  view.setUint32(4, 88, true);
  view.setInt32(8, bounds[0], true);
  view.setInt32(12, bounds[1], true);
  view.setInt32(16, bounds[2], true);
  view.setInt32(20, bounds[3], true);
  view.setInt32(24, frame[0], true);
  view.setInt32(28, frame[1], true);
  view.setInt32(32, frame[2], true);
  view.setInt32(36, frame[3], true);
  bytes.set([0x20, 0x45, 0x4d, 0x46], 40);
  view.setUint32(48, totalSize, true);
  view.setUint32(52, records.length + 2, true);
  let offset = 88;
  for (const record of records) {
    bytes.set(record, offset);
    offset += record.length;
  }
  view.setUint32(offset, 14, true); // EMR_EOF
  view.setUint32(offset + 4, 20, true);
  return bytes;
}

function emfCoordinateRecord(type, ...values) {
  const bytes = new Uint8Array(8 + values.length * 4);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, type, true);
  view.setUint32(4, bytes.length, true);
  values.forEach((value, index) => view.setInt32(8 + index * 4, value, true));
  return bytes;
}

function emfWorldTransformRecord(scale) {
  const bytes = new Uint8Array(36);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 36, true); // EMR_MODIFYWORLDTRANSFORM
  view.setUint32(4, bytes.length, true);
  view.setFloat32(8, scale, true);
  view.setFloat32(20, scale, true);
  view.setUint32(32, 2, true); // MWT_LEFTMULTIPLY
  return bytes;
}

function emfPenRecord(handle, width) {
  const bytes = new Uint8Array(28);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 38, true); // EMR_CREATEPEN
  view.setUint32(4, bytes.length, true);
  view.setUint32(8, handle, true);
  view.setInt32(16, width, true);
  return bytes;
}

function emfTextRecord(text, x = 50, y = 50, advances = undefined) {
  const encoded = new Uint8Array(Buffer.from(text, "utf16le"));
  const advancesOffset = (76 + encoded.length + 3) & ~3;
  const size = advances === undefined ? advancesOffset : advancesOffset + advances.length * 4;
  const bytes = new Uint8Array(size);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 84, true); // EMR_EXTTEXTOUTW
  view.setUint32(4, size, true);
  view.setInt32(36, x, true);
  view.setInt32(40, y, true);
  view.setUint32(44, text.length, true);
  view.setUint32(48, 76, true);
  bytes.set(encoded, 76);
  if (advances !== undefined) {
    view.setUint32(72, advancesOffset, true);
    advances.forEach((advance, index) => view.setInt32(advancesOffset + index * 4, advance, true));
  }
  return bytes;
}

function emfPlusComment(plusRecords) {
  const plusLength = plusRecords.reduce((total, record) => total + record.length, 0);
  const commentSize = 16 + plusLength;
  const bytes = new Uint8Array(commentSize);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 70, true); // EMR_COMMENT
  view.setUint32(4, commentSize, true);
  view.setUint32(8, 4 + plusLength, true);
  view.setUint32(12, 0x2b46_4d45, true); // EMF+
  let offset = 16;
  for (const record of plusRecords) {
    bytes.set(record, offset);
    offset += record.length;
  }
  return bytes;
}

function makeEmf(plusRecords = []) {
  return makeEmfRecords(plusRecords.length === 0 ? [] : [emfPlusComment(plusRecords)]);
}

function emfDibRecord(type, width, height, { invalidBitsOffset = false } = {}) {
  const fieldsOffset = type === 76 ? 84 : 48;
  const bmiOffset = Math.max(fieldsOffset + 16, 100);
  const bmiLength = 40;
  const bitsOffset = bmiOffset + bmiLength;
  const bitsLength = 4;
  const recordSize = bitsOffset + bitsLength;
  const bytes = new Uint8Array(recordSize);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, type, true);
  view.setUint32(4, recordSize, true);
  view.setUint32(fieldsOffset, bmiOffset, true);
  view.setUint32(fieldsOffset + 4, bmiLength, true);
  view.setUint32(fieldsOffset + 8, invalidBitsOffset ? recordSize + 4 : bitsOffset, true);
  view.setUint32(fieldsOffset + 12, bitsLength, true);
  view.setUint32(bmiOffset, 40, true);
  view.setInt32(bmiOffset + 4, width, true);
  view.setInt32(bmiOffset + 8, height, true);
  view.setUint16(bmiOffset + 12, 1, true);
  view.setUint16(bmiOffset + 14, 32, true);
  return bytes;
}

function emfOverlappingDibRecord() {
  const recordSize = 1_000;
  const resourceOffset = 100;
  const resourceLength = recordSize - resourceOffset;
  const bytes = new Uint8Array(recordSize);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 76, true);
  view.setUint32(4, recordSize, true);
  view.setUint32(84, resourceOffset, true);
  view.setUint32(88, resourceLength, true);
  view.setUint32(92, resourceOffset, true);
  view.setUint32(96, resourceLength, true);
  view.setUint32(resourceOffset, 40, true);
  view.setInt32(resourceOffset + 4, 1, true);
  view.setInt32(resourceOffset + 8, 1, true);
  view.setUint16(resourceOffset + 12, 1, true);
  view.setUint16(resourceOffset + 14, 32, true);
  return bytes;
}

function nestedEmf(depth) {
  let current = makeEmf();
  for (let level = 0; level < depth; level += 1) {
    current = makeEmf([emfPlusMetafileObject(0, current), emfPlusDrawImage(0)]);
  }
  return current;
}

function makeWmf() {
  const bytes = new Uint8Array(46);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x9ac6_cdd7, true);
  view.setInt16(10, 1, true);
  view.setInt16(12, 1, true);
  view.setUint16(14, 96, true);
  let checksum = 0;
  for (let offset = 0; offset < 20; offset += 2) checksum ^= view.getUint16(offset, true);
  view.setUint16(20, checksum, true);
  view.setUint16(22, 1, true);
  view.setUint16(24, 9, true);
  view.setUint16(26, 0x0300, true);
  view.setUint32(28, bytes.length / 2, true);
  view.setUint32(34, 3, true);
  view.setUint32(40, 3, true);
  return bytes;
}

function wmfRecord(type, ...values) {
  const bytes = new Uint8Array(6 + values.length * 2);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, bytes.length / 2, true);
  view.setUint16(4, type, true);
  values.forEach((value, index) => view.setInt16(6 + index * 2, value, true));
  return bytes;
}

function wmfFontRecord(family, height = 88) {
  const bytes = new Uint8Array(56);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, bytes.length / 2, true);
  view.setUint16(4, 0x02fb, true);
  view.setInt16(6, height, true);
  view.setInt16(14, 400, true);
  bytes.set(Buffer.from(family, "latin1").subarray(0, 31), 24);
  return bytes;
}

function wmfTextRecord(text, x, y) {
  const encoded = Buffer.from(text, "latin1");
  const bytes = new Uint8Array(12 + encoded.length + encoded.length % 2);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, bytes.length / 2, true);
  view.setUint16(4, 0x0521, true);
  view.setInt16(6, encoded.length, true);
  bytes.set(encoded, 8);
  const coordinateOffset = 8 + encoded.length + encoded.length % 2;
  view.setInt16(coordinateOffset, y, true);
  view.setInt16(coordinateOffset + 2, x, true);
  return bytes;
}

function wmfExtTextRecord(text, x, y) {
  const encoded = Buffer.from(text, "latin1");
  const bytes = new Uint8Array(14 + encoded.length + encoded.length % 2);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, bytes.length / 2, true);
  view.setUint16(4, 0x0a32, true);
  view.setInt16(6, y, true);
  view.setInt16(8, x, true);
  view.setInt16(10, encoded.length, true);
  bytes.set(encoded, 14);
  return bytes;
}

function makeStandardWmf(records) {
  const eof = wmfRecord(0);
  const recordsLength = records.reduce((total, record) => total + record.length, eof.length);
  const bytes = new Uint8Array(18 + recordsLength);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, 1, true);
  view.setUint16(2, 9, true);
  view.setUint16(4, 0x0300, true);
  view.setUint32(6, bytes.length / 2, true);
  view.setUint16(10, 0, true);
  view.setUint32(12, Math.max(3, ...records.map((record) => record.length / 2)), true);
  let offset = 18;
  for (const record of [...records, eof]) {
    bytes.set(record, offset);
    offset += record.length;
  }
  return bytes;
}

function makePlaceableWmfWithMissingHeight(records) {
  const standard = makeStandardWmf(records);
  const bytes = new Uint8Array(22 + standard.length);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x9ac6_cdd7, true);
  view.setInt16(10, 524, true);
  view.setUint16(14, 72, true);
  let checksum = 0;
  for (let offset = 0; offset < 20; offset += 2) checksum ^= view.getUint16(offset, true);
  view.setUint16(20, checksum, true);
  bytes.set(standard, 22);
  return bytes;
}

function wmfStretchDibRecord(rop = 0x00cc_0020, color = [0, 0, 0]) {
  const dib = new Uint8Array(44);
  const dibView = new DataView(dib.buffer);
  dibView.setUint32(0, 40, true);
  dibView.setInt32(4, 1, true);
  dibView.setInt32(8, 1, true);
  dibView.setUint16(12, 1, true);
  dibView.setUint16(14, 24, true);
  dibView.setUint32(20, 4, true);
  const record = new Uint8Array(28 + dib.length);
  const view = new DataView(record.buffer);
  view.setUint32(0, record.length / 2, true);
  view.setUint16(4, 0x0f43, true);
  view.setUint32(6, rop, true);
  for (const offset of [12, 14, 20, 22]) view.setInt16(offset, 1, true);
  dib.set([color[2], color[1], color[0]], 40);
  record.set(dib, 28);
  return record;
}

function wmfPatBltRecord(x, y, width, height) {
  const record = new Uint8Array(18);
  const view = new DataView(record.buffer);
  view.setUint32(0, record.length / 2, true);
  view.setUint16(4, 0x061d, true);
  view.setUint32(6, 0x00f0_0021, true);
  view.setInt16(10, height, true);
  view.setInt16(12, width, true);
  view.setInt16(14, y, true);
  view.setInt16(16, x, true);
  return record;
}

function wmfDibStretchBltRecord(color = [255, 0, 0], width = 1, height = 1) {
  const rowStride = Math.ceil(width * 3 / 4) * 4;
  const dib = new Uint8Array(40 + rowStride * height);
  const dibView = new DataView(dib.buffer);
  dibView.setUint32(0, 40, true);
  dibView.setInt32(4, width, true);
  dibView.setInt32(8, height, true);
  dibView.setUint16(12, 1, true);
  dibView.setUint16(14, 24, true);
  dibView.setUint32(20, rowStride * height, true);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) dib.set([color[2], color[1], color[0]], 40 + y * rowStride + x * 3);
  }
  const record = new Uint8Array(26 + dib.length);
  const view = new DataView(record.buffer);
  view.setUint32(0, record.length / 2, true);
  view.setUint16(4, 0x0b41, true);
  view.setUint32(6, 0x00cc_0020, true);
  view.setInt16(10, height, true);
  view.setInt16(12, width, true);
  view.setInt16(18, height, true);
  view.setInt16(20, width, true);
  record.set(dib, 26);
  return record;
}

function codecInput(bytes, overrides = {}) {
  return {
    format: "emf",
    mediaType: "image/x-emf",
    bytes,
    maxPixels: 1_000,
    maxBytes: 1_000_000,
    maxCompressionRatio: 200,
    ...overrides,
  };
}

function installCanvasEnvironment() {
  const names = ["OffscreenCanvas", "FileReader", "ImageData", "createImageBitmap"];
  const descriptors = new Map(names.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)]));
  let bitmapCalls = 0;
  const context = new Proxy({}, {
    get(target, property) {
      if (!(property in target)) target[property] = () => {};
      return target[property];
    },
    set(target, property, value) {
      target[property] = value;
      return true;
    },
  });
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
    }
    getContext() { return context; }
    async convertToBlob() {
      return new Blob([pngHeader(this.width, this.height)], { type: "image/png" });
    }
  }
  class FakeFileReader {
    result = null;
    error = null;
    onload = null;
    onerror = null;
    readAsDataURL(blob) {
      blob.arrayBuffer().then((buffer) => {
        this.result = `data:${blob.type};base64,${Buffer.from(buffer).toString("base64")}`;
        this.onload?.();
      }, (error) => {
        this.error = error;
        this.onerror?.();
      });
    }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "FileReader", { configurable: true, value: FakeFileReader });
  Object.defineProperty(globalThis, "ImageData", {
    configurable: true,
    value: class ImageData {
      constructor(data, width, height) {
        this.data = data;
        this.width = width;
        this.height = height;
      }
    },
  });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => {
      bitmapCalls += 1;
      return { width: 1, height: 1, close() {} };
    },
  });
  return {
    context,
    bitmapCalls: () => bitmapCalls,
    restore() {
      for (const [name, descriptor] of descriptors) {
        if (descriptor === undefined) delete globalThis[name];
        else Object.defineProperty(globalThis, name, descriptor);
      }
    },
  };
}

test("normalizes GDI mapping coordinates to non-zero EMF bounds", async () => {
  const environment = installCanvasEnvironment();
  const moves = [];
  environment.context.moveTo = (x, y) => moves.push([x, y]);
  try {
    const bytes = makeEmfRecords([
      emfCoordinateRecord(17, 8),
      emfCoordinateRecord(10, 0, 0),
      emfCoordinateRecord(9, 100, 100),
      emfCoordinateRecord(12, 0, 0),
      emfCoordinateRecord(11, 100, 100),
      emfCoordinateRecord(27, 150, 150),
      emfCoordinateRecord(54, 160, 160),
    ], [100, 100, 200, 200], [0, 0, 3_175, 3_175]);
    await decodeOfficeImagePayload(codecInput(bytes, { maxPixels: 20_000 }));

    assert.deepEqual(moves[0], [50, 50]);
  } finally {
    environment.restore();
  }
});

test("uses the EMF frame and device metrics for the full render canvas", async () => {
  const environment = installCanvasEnvironment();
  try {
    const bytes = makeEmfRecords([], [11, 14, 369, 276], [0, 0, 9_599, 6_491]);
    const view = new DataView(bytes.buffer);
    view.setInt32(72, 1_280, true);
    view.setInt32(76, 1_024, true);
    view.setInt32(80, 320, true);
    view.setInt32(84, 240, true);
    const output = await decodeOfficeImagePayload(codecInput(bytes, {
      maxPixels: 384 * 277,
      targetWidth: 384,
      targetHeight: 277,
    }));

    assert.equal(output.kind, "encoded");
    assert.deepEqual([output.width, output.height], [384, 277]);
  } finally {
    environment.restore();
  }
});

test("normalizes an implicit EMF viewport origin to non-zero header bounds", async () => {
  const environment = installCanvasEnvironment();
  const moves = [];
  environment.context.moveTo = (x, y) => moves.push([x, y]);
  try {
    const bytes = makeEmfRecords([
      emfCoordinateRecord(17, 8),
      emfCoordinateRecord(11, 384, 256),
      emfCoordinateRecord(9, 384, 256),
      emfCoordinateRecord(10, 0, 5_460),
      emfCoordinateRecord(9, 16_383, 10_923),
      emfCoordinateRecord(27, 1_536, 5_460),
      emfCoordinateRecord(54, 16_383, 14_311),
    ], [35, 0, 383, 207], [0, 0, 9_599, 5_999]);
    await decodeOfficeImagePayload(codecInput(bytes, {
      maxPixels: 452 * 269,
      targetWidth: 452,
      targetHeight: 269,
    }));

    assert.ok(Math.abs(moves[0][0] - 1.3) < 0.1, `implicit viewport origin was not normalized: ${JSON.stringify(moves[0])}`);
  } finally {
    environment.restore();
  }
});

test("maps geometric EMF pen widths through the active GDI window", async () => {
  const environment = installCanvasEnvironment();
  const lineWidths = [];
  Object.defineProperty(environment.context, "lineWidth", {
    configurable: true,
    set(value) { lineWidths.push(value); },
  });
  try {
    const bytes = makeEmfRecords([
      emfCoordinateRecord(17, 8),
      emfCoordinateRecord(11, 384, 256),
      emfCoordinateRecord(10, 0, 5_460),
      emfCoordinateRecord(9, 16_383, 10_923),
      emfPenRecord(1, 154),
      emfCoordinateRecord(37, 1),
      emfCoordinateRecord(27, 1_536, 5_460),
      emfCoordinateRecord(54, 16_383, 14_311),
    ], [35, 0, 383, 207], [0, 0, 9_599, 5_999]);
    await decodeOfficeImagePayload(codecInput(bytes, {
      maxPixels: 452 * 269,
      targetWidth: 452,
      targetHeight: 269,
    }));

    assert.ok(lineWidths.some((width) => width > 4 && width < 5), `geometric pen width was not mapped: ${JSON.stringify(lineWidths)}`);
  } finally {
    environment.restore();
  }
});

test("preserves EMF centered text alignment", async () => {
  const environment = installCanvasEnvironment();
  const alignments = [];
  environment.context.fillText = () => alignments.push(environment.context.textAlign);
  try {
    const bytes = makeEmfRecords([
      emfCoordinateRecord(22, 6), // TA_CENTER
      emfTextRecord("Centered"),
    ], [0, 0, 100, 100], [0, 0, 2_646, 2_646]);
    await decodeOfficeImagePayload(codecInput(bytes, { maxPixels: 10_000 }));

    assert.deepEqual(alignments, ["center"]);
  } finally {
    environment.restore();
  }
});

test("uses authored EMF character advances instead of browser font widths", async () => {
  const environment = installCanvasEnvironment();
  const text = [];
  environment.context.fillText = (value, x, y, maxWidth) => text.push([value, x, y, maxWidth]);
  try {
    const bytes = makeEmfRecords([
      emfTextRecord("File", 10, 20, [5, 5, 5, 0]),
    ], [0, 0, 100, 100], [0, 0, 2_646, 2_646]);
    await decodeOfficeImagePayload(codecInput(bytes, { maxPixels: 10_000 }));

    assert.deepEqual(text, [
      ["F", 10, 20, undefined],
      ["i", 15, 20, undefined],
      ["l", 20, 20, undefined],
      ["e", 25, 20, undefined],
    ]);
  } finally {
    environment.restore();
  }
});

test("honors the supplied photo workbook's GDI font cell height and width", async () => {
  const parts = readZipEntries(await readFile(new URL("./fixtures/Photo Formats - CGM-O12-XL-Pictures.xlsx", import.meta.url)));
  const environment = installCanvasEnvironment();
  const fonts = [];
  const scales = [];
  const glyphs = [];
  const stack = [];
  let origin = [0, 0];
  environment.context.save = () => stack.push([...origin]);
  environment.context.restore = () => { origin = stack.pop() ?? [0, 0]; };
  environment.context.translate = (x, y) => { origin = [origin[0] + x, origin[1] + y]; };
  environment.context.measureText = () => {
    const size = Number(environment.context.font.match(/([\d.]+)px/u)[1]);
    return { width: size / 2, fontBoundingBoxAscent: size, fontBoundingBoxDescent: size / 5 };
  };
  environment.context.fillText = (text, x, y) => {
    fonts.push(environment.context.font);
    glyphs.push({ text, x: origin[0] + x, y: origin[1] + y });
  };
  environment.context.scale = (x, y) => scales.push([x, y]);
  try {
    await decodeOfficeImagePayload(codecInput(parts.find(({ name }) => name === "xl/media/image1.emf").data, {
      maxPixels: 10_000_000, maxBytes: 16_000_000,
    }));
    // Original LOGFONT: Arial Italic, cell height 730, average width 287.
    // Original window 16383 x 11833; codec fits its viewport to 340 x 244.
    const emHeight = 730 * 244 / 11833 / 1.2;
    assert.ok(Math.abs(Number(fonts[0].match(/([\d.]+)px/u)[1]) - emHeight) < 0.001, fonts[0]);
    const widthScale = (287 * 340 / 16383) / (emHeight / 2);
    assert.ok(scales.some(([x, y]) => Math.abs(x - widthScale) < 0.001 && y === 1), JSON.stringify(scales[0]));
    assert.match(fonts[0], /^italic .* Arial$/u);
    // The subtitle intentionally remains at the exact recorded origins, even when
    // a native application displays wider natural spacing. Never reflow it.
    const subtitle = glyphs.slice(22, 47);
    assert.equal(subtitle.map(({ text }) => text).join(""), "Lease Maps vs. MIMIC Maps");
    const advances = [171, 170, 171, 171, 170, 86, 298, 171, 171, 170, 86, 170, 171, 85, 86, 298, 86, 298, 85, 256, 86, 298, 171, 171, 170];
    let x = 8190 - 4266 / 2;
    for (const [index, glyph] of subtitle.entries()) {
      assert.ok(Math.abs(glyph.x - x * 340 / 16383) < 0.001, `subtitle glyph ${index} x`);
      assert.ok(Math.abs(glyph.y - (5860 - 4550) * 244 / 11833) < 0.001, `subtitle glyph ${index} y`);
      x += advances[index];
    }
  } finally {
    environment.restore();
  }
});

test("shares signed GDI heights and explicit widths across EMF and both WMF text records", async () => {
  for (const height of [12, -12]) for (const width of [0, 3]) {
    const font = new Uint8Array(332);
    const view = new DataView(font.buffer);
    view.setUint32(0, 82, true);
    view.setUint32(4, font.length, true);
    view.setUint32(8, 1, true);
    view.setInt32(12, height, true);
    view.setInt32(16, width, true);
    font.set(Buffer.from("Arial", "utf16le"), 40);
    const inputs = [codecInput(makeEmfRecords([
      font, emfCoordinateRecord(37, 1), emfTextRecord("AB", 10, 20),
    ], [0, 0, 100, 100], [0, 0, 2_646, 2_646]), { maxPixels: 10_000 })];
    for (const record of [wmfTextRecord, wmfExtTextRecord]) {
      const wmfFont = wmfFontRecord("Arial", height);
      new DataView(wmfFont.buffer).setInt16(8, width, true);
      inputs.push(codecInput(makeStandardWmf([
        wmfRecord(524, 100, 100), wmfFont, wmfRecord(0x012d, 0), record("AB", 10, 20),
      ]), { format: "wmf", mediaType: "image/wmf", maxPixels: 10_000, targetWidth: 100, targetHeight: 100 }));
    }
    for (const input of inputs) {
      const environment = installCanvasEnvironment();
      const scales = [];
      const drawn = [];
      environment.context.measureText = () => {
        const size = Number(environment.context.font.match(/([\d.]+)px/u)[1]);
        return { width: size / 2, fontBoundingBoxAscent: size, fontBoundingBoxDescent: size / 5 };
      };
      environment.context.scale = (x, y) => scales.push([x, y]);
      environment.context.fillText = (text) => drawn.push([text, environment.context.font]);
      try {
        await decodeOfficeImagePayload(input);
        const size = height > 0 ? 10 : 12;
        assert.deepEqual(drawn, [["AB", `${size}px Arial`]], input.format);
        assert.deepEqual(scales, width === 0 ? [] : [[width / (size / 2), 1]], input.format);
      } finally {
        environment.restore();
      }
    }
  }
});

test("keeps the supplied PAAK EMF diagrams inside their scaled canvas", async () => {
  const parts = readZipEntries(await readFile(new URL("./fixtures/China PAAK EPC evaluation.xlsx", import.meta.url)));
  for (const { name, data } of parts.filter(({ name }) => /^xl\/media\/image\d+\.emf$/u.test(name))) {
    const environment = installCanvasEnvironment();
    const fills = [];
    const images = [];
    environment.context.drawImage = (...args) => images.push(args);
    let matrix = [1, 0, 0, 1, 0, 0];
    environment.context.setTransform = (...value) => { matrix = value; };
    environment.context.fillRect = (x, y, width, height) => {
      fills.push([matrix[0] * (x + width) + matrix[2] * (y + height) + matrix[4],
        matrix[1] * (x + width) + matrix[3] * (y + height) + matrix[5]]);
    };
    try {
      const output = await decodeOfficeImagePayload(codecInput(data, {
        maxPixels: 10_000_000, maxBytes: 16_000_000, targetWidth: 200, targetHeight: 200,
      }));
      if (name.endsWith("image1.emf")) assert.equal(images.length, 5, "card and key bitmap records must render");
      assert.ok(fills.length > 0, `${name} must draw its authored shapes`);
      for (const [right, bottom] of fills) {
        assert.ok(right <= output.width + 2 && bottom <= output.height + 2,
          `${name}: shape at ${right},${bottom} is cropped by ${output.width}x${output.height}`);
      }
    } finally {
      environment.restore();
    }
  }
});

test("preserves the supplied PAAK connector label angles and glyph widths", async () => {
  const parts = readZipEntries(await readFile(new URL("./fixtures/China PAAK EPC evaluation.xlsx", import.meta.url)));
  const environment = installCanvasEnvironment();
  const text = [];
  const stack = [];
  let angle = 0;
  environment.context.save = () => stack.push(angle);
  environment.context.restore = () => { angle = stack.pop() ?? 0; };
  environment.context.setTransform = (a, b) => { angle = Math.atan2(b, a); };
  environment.context.rotate = value => { angle += value; };
  environment.context.fillText = (value, x, y, maxWidth) => { if (value.trim()) text.push({value, angle, maxWidth}); };
  try {
    await decodeOfficeImagePayload(codecInput(parts.find(({name}) => name === "xl/media/image1.emf").data, {
      maxPixels: 10_000_000, maxBytes: 16_000_000, targetWidth: 1104, targetHeight: 561,
    }));
    for (const [tenths, label] of [[3371,"UWB"],[562,"UWB+BLE"],[1078,"UWB"],[801,"UWB"],
      [3318,"UWB"],[352,"UWB+BLE"],[3285,"UWB"],[3187,"UWB"],[679,"UWB"],[480,"UWB"],[756,"NFC"],[552,"NFC"]]) {
      const radians = -tenths * Math.PI / 1800;
      const glyphs = text.filter(({angle}) => Math.abs(Math.sin(angle-radians)) < 0.0001 && Math.cos(angle-radians) > 0);
      assert.equal(glyphs.map(({value}) => value).join(""), label, `${tenths / 10} degree label must rotate as authored`);
      assert.ok(glyphs.every(({maxWidth}) => maxWidth === undefined), "single glyph spacing must not squeeze the glyph");
    }
  } finally { environment.restore(); }
});

test("renders the supplied PAAK red path borders with their authored dash style", async () => {
  const parts = readZipEntries(await readFile(new URL("./fixtures/China PAAK EPC evaluation.xlsx", import.meta.url)));
  const environment = installCanvasEnvironment();
  const redStrokes = [];
  let lineDash = [];
  environment.context.setLineDash = value => { lineDash = [...value]; };
  environment.context.stroke = () => {
    if (environment.context.strokeStyle === "rgba(255,0,0,1.000)") redStrokes.push([...lineDash]);
  };
  try {
    await decodeOfficeImagePayload(codecInput(parts.find(({name}) => name === "xl/media/image1.emf").data, {
      maxPixels: 10_000_000, maxBytes: 16_000_000, targetWidth: 1104, targetHeight: 561,
    }));
    assert.ok(redStrokes.length > 0, "the supplied diagram must stroke its red path borders");
    assert.ok(redStrokes.every(dash => dash.length > 0), "red path borders must keep their authored dash style");
  } finally { environment.restore(); }
});

test("shares GDI font rotation across EMF and both WMF text records", async () => {
  const environment = installCanvasEnvironment();
  const rotations = [];
  const text = [];
  environment.context.rotate = angle => rotations.push(angle);
  environment.context.fillText = (value, x, y) => text.push([value, x, y]);
  try {
    for (const angle of [900, -450, 0]) {
      const font = new Uint8Array(332);
      const view = new DataView(font.buffer);
      view.setUint32(0, 82, true); view.setUint32(4, 332, true);
      view.setUint32(8, 1, true); view.setInt32(12, -13, true);
      view.setInt32(20, angle, true); view.setInt32(24, angle, true);
      font.set(Buffer.from("Arial", "utf16le"), 40);
      await decodeOfficeImagePayload(codecInput(makeEmfRecords([
        font, emfCoordinateRecord(37, 1), emfTextRecord("AB", 10, 20),
      ]), {maxPixels: 10_000}));
      for (const record of [wmfTextRecord, wmfExtTextRecord]) {
        const font = wmfFontRecord("Arial", 13);
        new DataView(font.buffer).setInt16(10, angle, true);
        await decodeOfficeImagePayload(codecInput(makeStandardWmf([
          wmfRecord(524, 100, 100), font, wmfRecord(0x012d, 0), record("AB", 10, 20),
        ]), {format: "wmf", mediaType: "image/x-wmf", maxPixels: 10_000}));
      }
    }
    assert.deepEqual(rotations, [-Math.PI/2, -Math.PI/2, -Math.PI/2, Math.PI/4, Math.PI/4, Math.PI/4]);
    assert.ok(text.slice(0, 6).every(([value, x, y]) => value === "AB" && x === 0 && y === 0));
    assert.ok(text.slice(6).every(([value, x, y]) => value === "AB" && x === 10 && y === 20));
  } finally { environment.restore(); }
});

test("uses dual EMF GDI bitmaps only when the corresponding EMF+ image is unavailable", async () => {
  const environment = installCanvasEnvironment();
  const images = [];
  environment.context.drawImage = (...args) => images.push(args);
  try {
    const bytes = makeEmfRecords([
      emfPlusComment([emfPlusRecord(0x4001, 1, new Uint8Array(16)), emfPlusDrawImage(0)]),
      emfDibRecord(81, 2, 2),
      emfPlusComment([emfPlusRasterObject(0, pngHeader(1, 1)), emfPlusDrawImage(0)]),
      emfDibRecord(81, 2, 2),
    ]);
    await decodeOfficeImagePayload(codecInput(bytes, { maxPixels: 10_000 }));
    assert.equal(images.length, 2, "one GDI bitmap and one EMF+ bitmap, without duplicate drawing");
    assert.equal(environment.bitmapCalls(), 1, "only the available EMF+ raster is decoded");
  } finally {
    environment.restore();
  }
});

test("uses classic text once when merging a dual EMF+ stream", async () => {
  const environment = installCanvasEnvironment();
  const text = [];
  environment.context.fillText = (value) => text.push(value);
  try {
    const dualHeader = emfPlusRecord(0x4001, 1, new Uint8Array(16));
    const bytes = makeEmfRecords([
      emfPlusComment([dualHeader, ...emfPlusTextRecords("EMF+ copy")]),
      emfTextRecord("Classic fallback"),
    ]);
    await decodeOfficeImagePayload(codecInput(bytes, { maxPixels: 10_000 }));

    assert.deepEqual(text, ["Classic fallback"]);
  } finally {
    environment.restore();
  }
});

test("applies the EMF world transform to text position and size", async () => {
  const environment = installCanvasEnvironment();
  const text = [];
  const transforms = [];
  environment.context.setTransform = (...matrix) => transforms.push(matrix);
  environment.context.fillText = (value, x, y) => text.push([value, x, y, environment.context.font]);
  try {
    const bytes = makeEmfRecords([
      emfWorldTransformRecord(0.5),
      emfTextRecord("Scaled", 50, 50),
    ], [0, 0, 100, 100], [0, 0, 2_646, 2_646]);
    await decodeOfficeImagePayload(codecInput(bytes, { maxPixels: 10_000 }));

    assert.ok(transforms.some((matrix) => matrix.join() === "0.5,0,0,0.5,0,0"));
    assert.deepEqual(text, [["Scaled", 50, 50, "12px sans-serif"]]);
  } finally {
    environment.restore();
  }
});

test("uses an explicit EMF window when tight header bounds would crop the drawing", async () => {
  const environment = installCanvasEnvironment();
  try {
    const bytes = makeEmfRecords([
      emfCoordinateRecord(10, 0, 0),
      emfCoordinateRecord(9, 631, 424),
    ], [45, 82, 621, 341], [0, 0, 15_775, 9_938]);
    const output = await decodeOfficeImagePayload(codecInput(bytes, {
      maxPixels: 1_000_000,
      targetWidth: 800,
      targetHeight: 800,
    }));

    assert.equal(output.width, 800);
    assert.ok(output.height > 500);
  } finally {
    environment.restore();
  }
});

test("rasterizes a standard WMF window at the requested display resolution", async () => {
  const environment = installCanvasEnvironment();
  try {
    const bytes = makeStandardWmf([
      wmfRecord(523, 0, 0),
      wmfRecord(524, 60, 80),
      wmfFontRecord("Arial"),
      wmfRecord(0x012d, 0),
      wmfTextRecord("Sharp", 10, 10),
    ]);
    const output = await decodeOfficeImagePayload({
      ...codecInput(bytes, {
        format: "wmf",
        mediaType: "image/x-wmf",
        maxPixels: 1_000_000,
        targetWidth: 800,
        targetHeight: 800,
      }),
    });

    assert.equal(output.width, 800);
    assert.equal(output.height, 600);
    assert.match(environment.context.font, /Arial/u);
  } finally {
    environment.restore();
  }
});

test("maps Equation Editor MT Extra WMF accents to portable glyphs", async () => {
  const environment = installCanvasEnvironment();
  const text = [];
  environment.context.fillText = (value) => text.push([value, environment.context.font]);
  try {
    const bytes = makeStandardWmf([
      wmfRecord(524, 576, 2_752),
      wmfFontRecord("MT Extra"),
      wmfRecord(0x012d, 0),
      wmfExtTextRecord("&", 326, 366),
    ]);
    await decodeOfficeImagePayload(codecInput(bytes, {
      format: "wmf",
      mediaType: "image/x-wmf",
      maxPixels: 10_000,
      targetWidth: 114,
      targetHeight: 24,
    }));

    assert.equal(text[0]?.[0], "˙");
    assert.match(text[0]?.[1] ?? "", /Times New Roman/u);
  } finally {
    environment.restore();
  }
});

test("repairs missing placeable WMF bounds from its drawing window", async () => {
  const environment = installCanvasEnvironment();
  try {
    const bytes = makePlaceableWmfWithMissingHeight([
      wmfRecord(523, 0, 0),
      wmfRecord(524, 136, 4_192),
      wmfFontRecord("Arial"),
      wmfRecord(0x012d, 0),
      wmfTextRecord("Microsoft", 3_500, 20),
    ]);
    const output = await decodeOfficeImagePayload(codecInput(bytes, {
      format: "wmf",
      mediaType: "image/x-wmf",
      maxPixels: 1_000_000,
      targetWidth: 640,
      targetHeight: 36,
    }));

    assert.ok(output.width > 0);
    assert.ok(output.height > 0);
  } finally {
    environment.restore();
  }
});

test("renders WMF PATBLT controls and positioned DIB previews", async () => {
  const environment = installCanvasEnvironment();
  const fills = [];
  const images = [];
  environment.context.fillRect = (...args) => fills.push([environment.context.fillStyle, ...args]);
  environment.context.drawImage = (...args) => images.push(args);
  try {
    const bytes = makeStandardWmf([
      wmfRecord(524, 10, 10),
      wmfRecord(0x02fc, 0, 255, 0, 0),
      wmfRecord(0x012d, 0),
      wmfPatBltRecord(1, 2, 3, 4),
      wmfDibStretchBltRecord(),
    ]);
    await decodeOfficeImagePayload(codecInput(bytes, {
      format: "wmf",
      mediaType: "image/x-wmf",
      maxPixels: 10_000,
      targetWidth: 100,
      targetHeight: 100,
    }));

    assert.deepEqual(fills, [["#ff0000", 10, 20, 30, 40]]);
    assert.equal(images.length, 1);
  } finally {
    environment.restore();
  }
});

test("preflights pixels embedded in WMF DIB stretch records", async () => {
  const bytes = makeStandardWmf([
    wmfRecord(524, 10, 10),
    wmfDibStretchBltRecord([255, 0, 0], 20, 20),
  ]);

  await assert.rejects(
    decodeOfficeImagePayload(codecInput(bytes, {
      format: "wmf",
      mediaType: "image/x-wmf",
      maxPixels: 100,
      targetWidth: 10,
      targetHeight: 10,
    })),
    (error) => error?.code === "IMAGE_DIMENSION_LIMIT",
  );
});

test("maps WMF pen widths through the active logical window", async () => {
  const environment = installCanvasEnvironment();
  const lineWidths = [];
  Object.defineProperty(environment.context, "lineWidth", {
    configurable: true,
    set(value) { lineWidths.push(value); },
  });
  try {
    const bytes = makeStandardWmf([
      wmfRecord(524, 600, 2_400), // META_SETWINDOWEXT
      wmfRecord(0x02fa, 0, 40, 0, 0, 0), // META_CREATEPENINDIRECT
      wmfRecord(0x012d, 0), // META_SELECTOBJECT
      wmfRecord(0x0214, 0, 0), // META_MOVETO
      wmfRecord(0x0213, 600, 2_400), // META_LINETO
    ]);
    await decodeOfficeImagePayload(codecInput(bytes, {
      format: "wmf",
      mediaType: "image/x-wmf",
      maxPixels: 240 * 60,
      targetWidth: 240,
      targetHeight: 60,
    }));

    assert.deepEqual(lineWidths, [4]);
  } finally {
    environment.restore();
  }
});

test("keeps WMF object indexes aligned when a logical palette is created", async () => {
  const environment = installCanvasEnvironment();
  const fills = [];
  environment.context.fillRect = () => fills.push(environment.context.fillStyle);
  try {
    const bytes = makeStandardWmf([
      wmfRecord(0x00f7, 0x0300, 1, 0, 0), // META_CREATEPALETTE occupies object slot 0
      wmfRecord(0x02fc, 0, -8_448, 0x02ff, 0), // solid #00dfff brush in slot 1
      wmfRecord(0x012d, 1), // META_SELECTOBJECT
      wmfRecord(0x041b, 100, 100, 0, 0), // META_RECTANGLE
    ]);
    await decodeOfficeImagePayload({
      ...codecInput(bytes, { format: "wmf", mediaType: "image/x-wmf" }),
    });

    assert.deepEqual(fills, ["#00dfff"]);
  } finally {
    environment.restore();
  }
});

test("uses a full-frame WMF StretchDIB screen preview", async () => {
  const environment = installCanvasEnvironment();
  try {
    const output = await decodeOfficeImagePayload({
      ...codecInput(makeStandardWmf([wmfStretchDibRecord()]), {
        format: "wmf",
        mediaType: "image/x-wmf",
      }),
    });

    assert.equal(output.kind, "rgba");
    assert.deepEqual([...output.data], [0, 0, 0, 255]);
  } finally {
    environment.restore();
  }
});

test("composites full-frame WMF DIB raster operations instead of returning the first mask", async () => {
  const output = await decodeOfficeImagePayload({
    ...codecInput(makeStandardWmf([
      wmfStretchDibRecord(0x00ee_0086, [0, 0, 0]), // SRCPAINT mask
      wmfStretchDibRecord(0x0088_00c6, [255, 0, 0]), // SRCAND color
    ]), {
      format: "wmf",
      mediaType: "image/x-wmf",
    }),
  });

  assert.equal(output.kind, "rgba");
  assert.deepEqual([...output.data], [255, 0, 0, 255]);
});

test("rejects EMF+ raster images when their cumulative decoded pixels exceed the image budget", async () => {
  const environment = installCanvasEnvironment();
  try {
    const image = pngHeader(35, 20);
    const bytes = makeEmf([
      emfPlusRasterObject(0, image),
      emfPlusDrawImage(0),
      emfPlusRasterObject(1, image),
      emfPlusDrawImage(1),
    ]);
    await assert.rejects(
      decodeOfficeImagePayload(codecInput(bytes)),
      (error) => error?.code === "IMAGE_DIMENSION_LIMIT" && /cumulative/i.test(error.message),
    );
    assert.equal(environment.bitmapCalls(), 0);
  } finally {
    environment.restore();
  }
});

test("skips recursively embedded EMF+ metafiles while preserving top-level metafile rendering", async () => {
  const environment = installCanvasEnvironment();
  try {
    const output = await decodeOfficeImagePayload(codecInput(nestedEmf(6)));
    assert.equal(output.kind, "encoded");
    assert.equal(output.mediaType, "image/png");
    assert.equal(environment.bitmapCalls(), 0);
  } finally {
    environment.restore();
  }
});

test("cannot bypass nested-metafile disabling with EMF+ object continuation records", async () => {
  const environment = installCanvasEnvironment();
  try {
    const records = emfPlusContinuedMetafileObject(0, nestedEmf(3));
    const output = await decodeOfficeImagePayload(codecInput(makeEmf([...records, emfPlusDrawImage(0)])));
    assert.equal(output.kind, "encoded");
    assert.equal(environment.bitmapCalls(), 0);
  } finally {
    environment.restore();
  }
});

test("preflights cumulative pixels for classic EMR_BITBLT and EMR_STRETCHDIBITS images", async () => {
  const bytes = makeEmfRecords([
    emfDibRecord(76, 30, 20),
    emfDibRecord(81, 30, 20),
  ]);
  await assert.rejects(
    decodeOfficeImagePayload(codecInput(bytes)),
    (error) => error?.code === "IMAGE_DIMENSION_LIMIT" && /cumulative/i.test(error.message),
  );
});

test("rejects classic EMF bitmap offsets that escape their record", async () => {
  const bytes = makeEmfRecords([emfDibRecord(76, 1, 1, { invalidBitsOffset: true })]);
  await assert.rejects(
    decodeOfficeImagePayload(codecInput(bytes)),
    (error) => error?.code === "IMAGE_HEADER_INVALID" && /outside its record/i.test(error.message),
  );
});

test("rejects a truncated EMR_BITBLT before converter fields can cross into the next record", async () => {
  const record = new Uint8Array(96);
  const view = new DataView(record.buffer);
  view.setUint32(0, 76, true);
  view.setUint32(4, record.length, true);
  await assert.rejects(
    decodeOfficeImagePayload(codecInput(makeEmfRecords([record]))),
    (error) => error?.code === "IMAGE_HEADER_INVALID" && /truncated/i.test(error.message),
  );
});

test("enforces the cumulative embedded-byte budget before classic EMF bitmap decode", async () => {
  const bytes = makeEmfRecords([emfOverlappingDibRecord()]);
  await assert.rejects(
    decodeOfficeImagePayload(codecInput(bytes, { maxBytes: bytes.length })),
    (error) => error?.code === "IMAGE_SIZE_LIMIT" && /cumulative/i.test(error.message),
  );
});

test("keeps top-level EMF, WMF, EMZ, and WMZ rendering enabled", async () => {
  const environment = installCanvasEnvironment();
  try {
    const emf = makeEmf();
    const wmf = makeWmf();
    const cases = [
      ["emf", "image/x-emf", emf],
      ["wmf", "image/x-wmf", wmf],
      ["emz", "image/x-emz", new Uint8Array(gzipSync(emf))],
      ["wmz", "image/x-wmz", new Uint8Array(gzipSync(wmf))],
    ];
    for (const [format, mediaType, bytes] of cases) {
      const output = await decodeOfficeImagePayload(codecInput(bytes, { format, mediaType }));
      assert.equal(output.kind, "encoded", format);
      assert.equal(output.mediaType, "image/png", format);
    }
  } finally {
    environment.restore();
  }
});

test("rasterizes EMF at the requested display resolution", async () => {
  const environment = installCanvasEnvironment();
  const moves = [];
  environment.context.moveTo = (x, y) => moves.push([x, y]);
  try {
    const output = await decodeOfficeImagePayload(codecInput(makeEmfRecords(
      [
        emfCoordinateRecord(17, 8),
        emfCoordinateRecord(10, 0, 0),
        emfCoordinateRecord(9, 16, 8),
        emfCoordinateRecord(12, 0, 0),
        emfCoordinateRecord(11, 16, 8),
        emfCoordinateRecord(27, 8, 4),
        emfCoordinateRecord(54, 16, 8),
      ],
      [0, 0, 16, 8],
      [0, 0, 423, 212],
    ), {
      maxPixels: 64 * 32,
      targetWidth: 64,
      targetHeight: 32,
    }));

    assert.equal(output.kind, "encoded");
    assert.equal(output.width, 64);
    assert.equal(output.height, 32);
    assert.deepEqual(moves[0], [32, 16]);
  } finally {
    environment.restore();
  }
});

test("renders the original tdf114488 formula Symbol characters as Unicode", async () => {
  const xml = await readFile(new URL("./fixtures/tdf114488.fodg", import.meta.url), "utf8");
  const bytes = new Uint8Array(Buffer.from(xml.match(/<office:binary-data>([\s\S]*?)<\/office:binary-data>/u)[1].replace(/\s/gu, ""), "base64"));
  const environment = installCanvasEnvironment();
  const text = [];
  environment.context.fillText = value => text.push(value);
  try {
    await decodeOfficeImagePayload(codecInput(bytes, {
      format: "wmf", mediaType: "image/x-wmf", maxPixels: 100_000,
      targetWidth: 708, targetHeight: 116,
    }));
    assert.deepEqual(text, ["r", "d", "r", "h", "r", "j", "h", "i", "j", "i", "3", "*", ")", "(", "ˆ", ")", "(", "]", "|", "ˆ", "|", "[", "φ", "φ", "∫", "="]);
  } finally { environment.restore(); }
});


test("shares GDI text alignment across WMF text records and EMF", async () => {
  const environment = installCanvasEnvironment();
  const text = [];
  environment.context.fillText = () => text.push([environment.context.textBaseline, environment.context.textAlign]);
  try {
    for (const record of [wmfTextRecord, wmfExtTextRecord]) {
      for (const [flags, baseline, horizontal] of [[0, "top", "left"], [8, "bottom", "left"], [24, "alphabetic", "left"], [6, "top", "center"], [2, "top", "right"]]) {
        await decodeOfficeImagePayload(codecInput(makeStandardWmf([
          wmfRecord(524, 100, 300), wmfRecord(0x012e, flags), record("Caption", 100, 42),
        ]), { format: "wmf", mediaType: "image/x-wmf", maxPixels: 100_000 }));
        assert.deepEqual(text.pop(), [baseline, horizontal]);
        await decodeOfficeImagePayload(codecInput(makeEmfRecords([
          emfCoordinateRecord(22, flags), emfTextRecord("Caption", 100, 42),
        ]), { maxPixels: 100_000 }));
        assert.deepEqual(text.pop(), [baseline, horizontal]);
      }
    }
  } finally { environment.restore(); }
});


test("replays the supplied toggle button monochrome pattern with current GDI colors", async () => {
  const entries = readZipEntries(await readFile(new URL("./fixtures/activex_togglebutton.pptx", import.meta.url)));
  const bytes = entries.find(entry => entry.name === "ppt/media/image2.wmf").data;
  const environment = installCanvasEnvironment();
  const fills = [];
  const pattern = {};
  environment.context.createPattern = () => pattern;
  environment.context.fillRect = (...rect) => fills.push([environment.context.fillStyle, ...rect]);
  try {
    await decodeOfficeImagePayload(codecInput(bytes, { format: "wmf", mediaType: "image/x-wmf", maxPixels: 100_000 }));
    assert.ok(fills.some(([color, , , w, h]) => color === "#ff80ff" && w === 8 && h === 8));
    assert.equal(fills.filter(([color, , , w, h]) => color === "#ffffff" && w === 1 && h === 1).length, 32);
    assert.ok(fills.some(([color]) => color === pattern));
  } finally { environment.restore(); }
});

test("charges WMF pattern brush bitmaps to the same image budget as stretch records", async () => {
  const dib = wmfDibStretchBltRecord([255, 0, 0], 20, 20).subarray(26);
  const record = new Uint8Array(10 + dib.length);
  const view = new DataView(record.buffer);
  view.setUint32(0, record.length / 2, true);
  view.setUint16(4, 0x0142, true);
  view.setUint16(6, 6, true);
  record.set(dib, 10);
  await assert.rejects(decodeOfficeImagePayload(codecInput(makeStandardWmf([
    wmfRecord(524, 10, 10), record,
  ]), { format: "wmf", mediaType: "image/x-wmf", maxPixels: 100, targetWidth: 10, targetHeight: 10 })),
  error => error?.code === "IMAGE_DIMENSION_LIMIT");
});


test("complex0 WMF control pictures replay text and sound bitmap records", async () => {
  const entries = readZipEntries(await readFile(new URL("./fixtures/complex0.docx", import.meta.url)));
  const environment = installCanvasEnvironment();
  const text = [];
  environment.context.fillText = value => text.push(value);
  try {
    for (const name of ["image7.wmf", "image8.wmf"]) {
      const bytes = entries.find(entry => entry.name === `word/media/${name}`).data;
      await decodeOfficeImagePayload(codecInput(bytes, { format: "wmf", mediaType: "image/x-wmf", maxPixels: 100_000, targetWidth: 576, targetHeight: 32 }));
    }
    assert.ok(text.includes("Scrolling Text"), "ETO_OPAQUE includes a rectangle before its text");
    const sound = await decodeOfficeImagePayload(codecInput(entries.find(entry => entry.name === "word/media/image8.wmf").data, {
      format: "wmf", mediaType: "image/x-wmf", maxPixels: 100_000,
    }));
    assert.equal(sound.kind, "rgba", "The sound picture composites SRCAND followed by SRCINVERT");
    assert.deepEqual([...sound.data.slice((10 * 32 + 10) * 4, (10 * 32 + 10) * 4 + 4)], [255, 255, 255, 255]);
  } finally { environment.restore(); }
});


test("complex0 Marlett checkbox layers render their authored bevel instead of Latin letters", async () => {
  const entries = readZipEntries(await readFile(new URL("./fixtures/complex0.docx", import.meta.url)));
  const environment = installCanvasEnvironment();
  const text = [], paths = [];
  environment.context.fillText = value => text.push(value);
  environment.context.fill = () => paths.push(environment.context.fillStyle);
  try {
    await decodeOfficeImagePayload(codecInput(entries.find(entry => entry.name === "word/media/image9.wmf").data, {
      format: "wmf", mediaType: "image/x-wmf", maxPixels: 100_000,
    }));
    assert.deepEqual(text, []);
    assert.equal(paths.length, 5, "all five Marlett checkbox layers retain their own authored colors");
    const font = new Uint8Array(332);
    const view = new DataView(font.buffer);
    view.setUint32(0, 82, true); view.setUint32(4, font.length, true);
    view.setUint32(8, 1, true); view.setInt32(12, 13, true);
    font.set(Buffer.from("Marlett", "utf16le"), 40);
    await decodeOfficeImagePayload(codecInput(makeEmfRecords([
      font, emfCoordinateRecord(37, 1), ...["g", "f", "e", "d", "c"].map(layer => emfTextRecord(layer, 1, 5)),
    ], [0, 0, 27, 23], [0, 0, 714, 609]), { maxPixels: 100_000 }));
    assert.deepEqual(text, [], "EMF uses the same geometric glyph replay");
    assert.equal(paths.length, 10);

  } finally { environment.restore(); }
});
