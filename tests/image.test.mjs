import assert from "node:assert/strict";
import test from "node:test";

import { identifyImage, inspectImage } from "../dist/image.js";
import { DEFAULT_LIMITS } from "../dist/core.js";
import { decodeOfficeImage } from "../dist/image-codec-client.js";
import { SceneRenderer } from "../dist/render.js";
import { secureMetafileForConversion } from "../dist/metafile-safety.js";

function pngHeader(width, height) {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  new DataView(bytes.buffer).setUint32(8, 13, false);
  bytes.set([0x49, 0x48, 0x44, 0x52], 12);
  new DataView(bytes.buffer).setUint32(16, width, false);
  new DataView(bytes.buffer).setUint32(20, height, false);
  return bytes;
}

function setAscii(bytes, offset, value) {
  bytes.set(new TextEncoder().encode(value), offset);
}

function jpegHeader(width, height) {
  const bytes = new Uint8Array(11);
  bytes.set([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x07, 0x08]);
  const view = new DataView(bytes.buffer);
  view.setUint16(7, height, false);
  view.setUint16(9, width, false);
  return bytes;
}

function gifHeader(width, height) {
  const bytes = new Uint8Array(10);
  setAscii(bytes, 0, "GIF89a");
  const view = new DataView(bytes.buffer);
  view.setUint16(6, width, true);
  view.setUint16(8, height, true);
  return bytes;
}

test("keeps scaled EMF viewport edges inside the exclusive Canvas boundary", () => {
  const bytes = new Uint8Array(104);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 1, true);
  view.setUint32(4, 88, true);
  // EMF header bounds are inclusive, while its viewport extent is a size.
  view.setInt32(16, 383, true);
  view.setInt32(20, 383, true);
  view.setUint32(88, 11, true);
  view.setUint32(92, 16, true);
  view.setInt32(96, 384, true);
  view.setInt32(100, 384, true);

  const secured = secureMetafileForConversion(bytes, "emf", 1024, 1_000_000, 331, 331);
  const securedView = new DataView(secured.buffer, secured.byteOffset, secured.byteLength);
  assert.equal(securedView.getInt32(96, true), 330);
  assert.equal(securedView.getInt32(100, true), 330);
});

function webpHeader(kind, width, height) {
  const bytes = new Uint8Array(30);
  setAscii(bytes, 0, "RIFF");
  setAscii(bytes, 8, "WEBP");
  setAscii(bytes, 12, kind);
  const view = new DataView(bytes.buffer);
  if (kind === "VP8X") {
    const encodedWidth = width - 1;
    const encodedHeight = height - 1;
    bytes.set([encodedWidth & 0xff, (encodedWidth >>> 8) & 0xff, (encodedWidth >>> 16) & 0xff], 24);
    bytes.set([encodedHeight & 0xff, (encodedHeight >>> 8) & 0xff, (encodedHeight >>> 16) & 0xff], 27);
  } else if (kind === "VP8L") {
    const encodedWidth = width - 1;
    const encodedHeight = height - 1;
    bytes[20] = 0x2f;
    bytes[21] = encodedWidth & 0xff;
    bytes[22] = ((encodedWidth >>> 8) & 0x3f) | ((encodedHeight & 0x03) << 6);
    bytes[23] = (encodedHeight >>> 2) & 0xff;
    bytes[24] = (encodedHeight >>> 10) & 0x0f;
  } else {
    bytes.set([0x9d, 0x01, 0x2a], 23);
    view.setUint16(26, width, true);
    view.setUint16(28, height, true);
  }
  return bytes;
}

function bitmapHeader(width, height, { fileHeader, compression = 0 } = {}) {
  const offset = fileHeader ? 14 : 0;
  const bytes = new Uint8Array(offset + 40);
  const view = new DataView(bytes.buffer);
  if (fileHeader) setAscii(bytes, 0, "BM");
  view.setUint32(offset, 40, true);
  view.setInt32(offset + 4, width, true);
  view.setInt32(offset + 8, height, true);
  view.setUint32(offset + 16, compression, true);
  return bytes;
}

function os2BitmapHeader(width, height, { planes = 1, bits = 24 } = {}) {
  const bytes = new Uint8Array(16);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 16, true);
  view.setInt32(4, width, true);
  view.setInt32(8, height, true);
  view.setUint16(12, planes, true);
  view.setUint16(14, bits, true);
  return bytes;
}

function svgHeader(width, height) {
  return new TextEncoder().encode(
    `<?xml version="1.0"?><!--Office--><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}"/>`,
  );
}

function tiffHeader(width, height, { orientation = 1 } = {}) {
  const bytes = new Uint8Array(50);
  const view = new DataView(bytes.buffer);
  setAscii(bytes, 0, "II");
  view.setUint16(2, 42, true);
  view.setUint32(4, 8, true);
  view.setUint16(8, 3, true);
  for (const [offset, tag, type, value] of [
    [10, 256, 4, width],
    [22, 257, 4, height],
    [34, 274, 3, orientation],
  ]) {
    view.setUint16(offset, tag, true);
    view.setUint16(offset + 2, type, true);
    view.setUint32(offset + 4, 1, true);
    if (type === 3) view.setUint16(offset + 8, value, true);
    else view.setUint32(offset + 8, value, true);
  }
  return bytes;
}

function jpegXrHeader(width, height) {
  const bytes = new Uint8Array(38);
  const view = new DataView(bytes.buffer);
  bytes.set([0x49, 0x49, 0xbc, 0x01]);
  view.setUint32(4, 8, true);
  view.setUint16(8, 2, true);
  for (const [offset, tag, value] of [[10, 0xbc80, width], [22, 0xbc81, height]]) {
    view.setUint16(offset, tag, true);
    view.setUint16(offset + 2, 4, true);
    view.setUint32(offset + 4, 1, true);
    view.setUint32(offset + 8, value, true);
  }
  return bytes;
}

function icoHeader(width, height) {
  const bytes = new Uint8Array(23);
  const view = new DataView(bytes.buffer);
  view.setUint16(2, 1, true);
  view.setUint16(4, 1, true);
  bytes[6] = width === 256 ? 0 : width;
  bytes[7] = height === 256 ? 0 : height;
  view.setUint16(12, 32, true);
  view.setUint32(14, 1, true);
  view.setUint32(18, 22, true);
  return bytes;
}

function pcxHeader(width, height) {
  const bytes = new Uint8Array(128);
  const view = new DataView(bytes.buffer);
  bytes.set([0x0a, 0x05, 0x01, 0x08]);
  view.setUint16(8, width - 1, true);
  view.setUint16(10, height - 1, true);
  return bytes;
}

function j2kHeader(width, height, { xOrigin = 0, yOrigin = 0, components = 3 } = {}) {
  const sizLength = 38 + components * 3;
  const bytes = new Uint8Array(4 + sizLength);
  const view = new DataView(bytes.buffer);
  bytes.set([0xff, 0x4f, 0xff, 0x51]);
  view.setUint16(4, sizLength, false);
  view.setUint32(8, width + xOrigin, false);
  view.setUint32(12, height + yOrigin, false);
  view.setUint32(16, xOrigin, false);
  view.setUint32(20, yOrigin, false);
  view.setUint32(24, Math.max(1, width + xOrigin), false);
  view.setUint32(28, Math.max(1, height + yOrigin), false);
  view.setUint32(32, 0, false);
  view.setUint32(36, 0, false);
  view.setUint16(40, components, false);
  for (let component = 0; component < components; component += 1) {
    bytes.set([7, 1, 1], 42 + component * 3);
  }
  return bytes;
}

function jp2Header(width, height, { headerWidth = width, xOrigin = 0, yOrigin = 0 } = {}) {
  const codestream = j2kHeader(width, height, { xOrigin, yOrigin });
  const bytes = new Uint8Array(12 + 30 + 8 + codestream.length);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 12, false);
  setAscii(bytes, 4, "jP  ");
  view.setUint32(8, 0x0d0a870a, false);
  view.setUint32(12, 30, false);
  setAscii(bytes, 16, "jp2h");
  view.setUint32(20, 22, false);
  setAscii(bytes, 24, "ihdr");
  view.setUint32(28, height, false);
  view.setUint32(32, headerWidth, false);
  view.setUint16(36, 3, false);
  bytes[38] = 7;
  view.setUint32(42, 8 + codestream.length, false);
  setAscii(bytes, 46, "jp2c");
  bytes.set(codestream, 50);
  return bytes;
}

function emfHeader(width, height, { logicalWidth = width, logicalHeight = height } = {}) {
  const bytes = new Uint8Array(88);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 1, true);
  view.setUint32(4, 88, true);
  view.setInt32(16, logicalWidth, true);
  view.setInt32(20, logicalHeight, true);
  view.setInt32(32, Math.round(width * 2_540 / 96), true);
  view.setInt32(36, Math.round(height * 2_540 / 96), true);
  setAscii(bytes, 40, " EMF");
  return bytes;
}

function placeableWmfHeader(width, height) {
  const bytes = new Uint8Array(40);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x9ac6cdd7, true);
  view.setInt16(10, width, true);
  view.setInt16(12, height, true);
  view.setUint16(14, 96, true);
  let checksum = 0;
  for (let offset = 0; offset < 20; offset += 2) checksum ^= view.getUint16(offset, true);
  view.setUint16(20, checksum, true);
  view.setUint16(22, 1, true);
  view.setUint16(24, 9, true);
  view.setUint16(26, 0x0300, true);
  return bytes;
}

function standardWmfHeader(width, height) {
  const bytes = new Uint8Array(28);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, 1, true);
  view.setUint16(2, 9, true);
  view.setUint16(4, 0x0300, true);
  view.setUint32(18, 5, true);
  view.setUint16(22, 0x020c, true);
  view.setInt16(24, height, true);
  view.setInt16(26, width, true);
  return bytes;
}

function gzipHeader() {
  return Uint8Array.from([0x1f, 0x8b, 0x08, 0x00, 0, 0, 0, 0, 0, 0]);
}

function pdfRasterHeader(width, height, components = 3, predictor = 1) {
  const bytes = new Uint8Array(21);
  setAscii(bytes, 0, "OVPDFR01");
  const view = new DataView(bytes.buffer);
  view.setUint32(8, width, false);
  view.setUint32(12, height, false);
  bytes[16] = components;
  bytes[17] = predictor;
  return bytes;
}

function controlledBitmapDecodes() {
  const pending = [];
  return {
    pending,
    decode() {
      return new Promise((complete) => {
        const entry = {
          settled: false,
          resolve(bitmap) {
            if (entry.settled) return;
            entry.settled = true;
            complete(bitmap);
          },
        };
        pending.push(entry);
      });
    },
    settleRemaining() {
      for (const [index, entry] of pending.entries()) {
        entry.resolve({ id: `cleanup:${index}`, width: 1, height: 1, close() {} });
      }
    },
  };
}

function prefetchImageObject(format, numericId, { bounds, bytes, z = numericId } = {}) {
  let source;
  if (format === "pptx") {
    source = {
      format,
      part: "ppt/slides/slide1.xml",
      kind: "shape",
      shapeId: numericId,
      mapping: "exact",
    };
  } else if (format === "xlsx") {
    source = {
      format,
      part: "xl/drawings/drawing1.xml",
      kind: "drawing",
      sheetName: "Sheet1",
      drawingId: numericId,
      mapping: "exact",
    };
  } else if (format === "docx") {
    source = {
      format,
      part: "word/document.xml",
      kind: "drawing",
      drawingId: numericId,
      mapping: "exact",
    };
  } else {
    throw new Error(`Unsupported prefetch test format: ${format}`);
  }
  return {
    numericId,
    id: `image:${format}:${numericId}`,
    type: "image",
    unitIndex: 0,
    bounds: bounds ?? { x: numericId * 10, y: 0, width: 8, height: 8 },
    source,
    z,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes: bytes ?? Uint8Array.from([...pngHeader(1, 1), numericId]),
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  };
}

test("validates raster signatures, media types, and decoded-pixel budgets before browser decode", () => {
  assert.deepEqual(inspectImage(pngHeader(2, 3), "image/png", 6), {
    mediaType: "image/png",
    width: 2,
    height: 3,
    pixels: 6,
  });

  assert.throws(
    () => inspectImage(pngHeader(2, 3), "image/jpeg", 6),
    (error) => error?.code === "IMAGE_MEDIA_TYPE_MISMATCH",
  );
  assert.throws(
    () => inspectImage(pngHeader(4096, 4096), "image/png", 1_000_000),
    (error) => error?.code === "IMAGE_DIMENSION_LIMIT",
  );
  assert.deepEqual(
    inspectImage(
      new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="2" height="3" viewBox="0 0 2 3"/>'),
      "image/svg+xml",
      6,
    ),
    { mediaType: "image/svg+xml", width: 2, height: 3, pixels: 6 },
  );
});

test("identifies every image family accepted from modern Office documents", () => {
  const cases = [
    ["PNG alias", pngHeader(2, 3), "image/x-png", "png", "image/png", 2, 3, true],
    ["JPEG alias", jpegHeader(2, 3), "image/jpg", "jpeg", "image/jpeg", 2, 3, true],
    ["GIF", gifHeader(2, 3), "image/gif", "gif", "image/gif", 2, 3, true],
    ["WebP extended", webpHeader("VP8X", 2, 3), "image/webp", "webp", "image/webp", 2, 3, true],
    ["WebP lossless", webpHeader("VP8L", 2, 3), "image/webp", "webp", "image/webp", 2, 3, true],
    ["WebP lossy", webpHeader("VP8 ", 2, 3), "image/webp", "webp", "image/webp", 2, 3, true],
    ["BMP RLE", bitmapHeader(2, 3, { fileHeader: true, compression: 1 }), "image/x-ms-bmp", "bmp", "image/bmp", 2, 3, true],
    ["raw DIB", bitmapHeader(2, -3, { fileHeader: false }), "image/bmp", "dib", "image/bmp", 2, 3, true],
    ["OS/2 2.x short DIB", os2BitmapHeader(2, 3), "image/bmp", "dib", "image/bmp", 2, 3, true],
    ["SVG alias", svgHeader(2, 3), "image/svg", "svg", "image/svg+xml", 2, 3, true],
    ["TIFF orientation", tiffHeader(2, 3, { orientation: 6 }), "image/tif", "tiff", "image/tiff", 3, 2, true],
    ["JPEG XR alias", jpegXrHeader(2, 3), "image/jxr", "jxr", "image/vnd.ms-photo", 2, 3, true],
    ["ICO alias", icoHeader(2, 3), "image/vnd.microsoft.icon", "ico", "image/x-icon", 2, 3, true],
    ["PCX alias", pcxHeader(2, 3), "image/vnd.zbrush.pcx", "pcx", "image/x-pcx", 2, 3, true],
    ["JP2 alias", jp2Header(2, 3), "image/jpx", "jp2", "image/jp2", 2, 3, true],
    ["J2K alias", j2kHeader(2, 3), "image/jpc", "j2k", "image/j2k", 2, 3, true],
    ["EMF alias", emfHeader(2, 3, { logicalWidth: 2_000_000_000, logicalHeight: 2_000_000_000 }), "image/emf", "emf", "image/x-emf", 2, 3, true],
    ["placeable WMF alias", placeableWmfHeader(2, 3), "image/wmf", "wmf", "image/x-wmf", 2, 3, true],
    ["standard WMF", standardWmfHeader(2, 3), "image/x-wmf", "wmf", "image/x-wmf", 1, 1, false],
    ["EMZ", gzipHeader(), "image/x-emz", "emz", "image/x-emz", 1, 1, false],
    ["WMZ", gzipHeader(), "image/x-wmz", "wmz", "image/x-wmz", 1, 1, false],
    ["compressed PDF raster", pdfRasterHeader(2, 3), "image/x-officeviewer-pdf-raster", "pdf-raster", "image/x-officeviewer-pdf-raster", 2, 3, true],
  ];

  for (const [label, bytes, declaredType, format, mediaType, width, height, dimensionsKnown] of cases) {
    assert.deepEqual(
      identifyImage(bytes, declaredType, 10_000),
      { format, mediaType, width, height, pixels: width * height, dimensionsKnown },
      label,
    );
  }
});

test("accepts a placeable WMF whose legacy bounds omit one dimension", () => {
  assert.deepEqual(
    identifyImage(placeableWmfHeader(524, 0), "image/x-wmf", 10_000),
    { format: "wmf", mediaType: "image/x-wmf", width: 1, height: 1, pixels: 1, dimensionsKnown: false },
  );
});

test("cross-checks JP2 image headers against codestream dimensions and honors non-zero origins", () => {
  assert.deepEqual(identifyImage(jp2Header(2, 3, { xOrigin: 7, yOrigin: 11 }), "image/jp2", 6), {
    format: "jp2",
    mediaType: "image/jp2",
    width: 2,
    height: 3,
    pixels: 6,
    dimensionsKnown: true,
  });
  assert.throws(
    () => identifyImage(jp2Header(2, 3, { headerWidth: 1 }), "image/jp2", 100),
    (error) => error?.code === "IMAGE_HEADER_INVALID",
  );
});

test("rejects malformed 16-byte OS/2 DIB layouts during identification", () => {
  assert.throws(
    () => identifyImage(os2BitmapHeader(2, 3, { planes: 0 }), "image/bmp", 100),
    (error) => error?.code === "IMAGE_HEADER_INVALID",
  );
});

test("Office image codec reuses at most eight healthy Workers", async () => {
  const workerDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const pending = [];
  const workers = [];
  let automatic = true;
  class FakeWorker {
    onmessage = null;
    onerror = null;
    onmessageerror = null;
    terminated = false;

    constructor() { workers.push(this); }
    postMessage(request) {
      const response = () => this.onmessage?.({ data: {
        id: request.id,
        ok: true,
        bitmap: { width: 1, height: 1, close() {} },
        approximate: false,
      } });
      if (automatic) queueMicrotask(response);
      else pending.push(response);
    }
    terminate() { this.terminated = true; }
  }
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: FakeWorker });
  const identified = {
    format: "pdf-raster",
    mediaType: "image/x-officeviewer-pdf-raster",
    width: 1,
    height: 1,
    pixels: 1,
    dimensionsKnown: true,
  };
  const decode = () => decodeOfficeImage(Uint8Array.of(1), identified, DEFAULT_LIMITS);
  try {
    (await decode()).bitmap.close();
    (await decode()).bitmap.close();
    assert.equal(workers.length, 1, "sequential decodes should reuse one Worker");

    automatic = false;
    const decoding = Array.from({ length: 12 }, decode);
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(workers.length, 8, "codec concurrency must stay bounded");
    while (pending.length > 0) {
      const batch = pending.splice(0);
      for (const respond of batch) respond();
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    const decoded = await Promise.all(decoding);
    for (const { bitmap } of decoded) bitmap.close();
    assert.equal(workers.length, 8);
  } finally {
    if (workerDescriptor) Object.defineProperty(globalThis, "Worker", workerDescriptor);
    else delete globalThis.Worker;
  }
});

test("rerasterizes vector images when the render scale increases", async () => {
  const workerDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const requests = [];
  const transforms = [{ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }];
  const context = {
    save() { transforms.push({ ...transforms.at(-1) }); },
    restore() { transforms.pop(); },
    scale(x, y) { transforms.at(-1).a *= x; transforms.at(-1).d *= y; },
    getTransform() { return transforms.at(-1); },
    translate() {}, fillRect() {}, drawImage() {},
    beginPath() {}, rect() {}, clip() {},
  };
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
    }
    getContext() { return context; }
    transferToImageBitmap() {
      return { width: this.width, height: this.height, close() {} };
    }
  }
  class FakeWorker {
    onmessage = null;
    onerror = null;
    onmessageerror = null;
    postMessage(request) {
      requests.push({
        format: request.format,
        width: request.targetWidth,
        height: request.targetHeight,
      });
      queueMicrotask(() => this.onmessage?.({ data: {
        id: request.id,
        ok: true,
        bitmap: {
          width: request.targetWidth ?? 2,
          height: request.targetHeight ?? 1,
          close() {},
        },
        approximate: true,
      } }));
    }
    terminate() {}
  }
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: FakeWorker });
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const renderers = [
    ["emf", "image/x-emf", emfHeader(100, 50)],
    ["svg", "image/svg+xml", svgHeader(2, 1)],
  ].map(([format, mediaType, bytes], index) => new SceneRenderer([{
    numericId: index + 1,
    id: `image:${index + 1}`,
    type: "image",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 100, height: 50 },
    source: {
      format: "doc",
      part: "WordDocument",
      kind: "image",
      mapping: "exact",
    },
    z: 1,
    visual: {
      kind: "image",
      mediaType,
      bytes,
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  }], DEFAULT_LIMITS));
  renderers.push(new SceneRenderer([{
    numericId: 3,
    id: "shape:3",
    type: "shape",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 100, height: 50 },
    source: { format: "xlsx", part: "xl/drawings/drawing1.xml", kind: "shape", mapping: "exact" },
    z: 1,
    visual: {
      kind: "painted-shape",
      geometry: "rectangle",
      fill: {
        kind: "image",
        mediaType: "image/x-emf",
        bytes: emfHeader(100, 50),
        cropLeft: 0,
        cropTop: 0,
        cropRight: 0,
        cropBottom: 0,
        tile: false,
      },
      stroke: { kind: "none" },
      strokeWidth: 0,
    },
  }], DEFAULT_LIMITS));
  renderers.push(new SceneRenderer([{
    numericId: 4,
    id: "group:4",
    type: "group",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 100, height: 50 },
    source: { format: "xlsx", part: "xl/drawings/drawing1.xml", kind: "group", mapping: "exact" },
    z: 1,
    visual: {
      kind: "group",
      children: [{
        bounds: { x: 0, y: 0, width: 100, height: 50 },
        visual: {
          kind: "painted-shape",
          geometry: "rectangle",
          fill: {
            kind: "image",
            mediaType: "image/x-emf",
            bytes: emfHeader(100, 50),
            cropLeft: 0.6,
            cropTop: 0,
            cropRight: 0,
            cropBottom: 0,
            tile: false,
          },
          stroke: { kind: "none" },
          strokeWidth: 0,
        },
      }],
    },
  }], DEFAULT_LIMITS));
  const unit = {
    type: "page",
    index: 0,
    id: "unit:0",
    name: "Page 1",
    width: 100,
    height: 50,
  };
  try {
    for (const renderer of renderers) {
      (await renderer.render(unit, { unitIndex: 0, scale: 1 })).bitmap.close();
      (await renderer.render(unit, { unitIndex: 0, scale: 3 })).bitmap.close();
    }
    assert.deepEqual(requests, [
      { format: "emf", width: 201, height: 101 },
      { format: "emf", width: 601, height: 301 },
      { format: "svg", width: 100, height: 50 },
      { format: "svg", width: 300, height: 150 },
      { format: "emf", width: 201, height: 101 },
      { format: "emf", width: 601, height: 301 },
      { format: "emf", width: 501, height: 101 },
      { format: "emf", width: 1501, height: 301 },
    ]);
  } finally {
    for (const renderer of renderers) renderer.close();
    if (workerDescriptor) Object.defineProperty(globalThis, "Worker", workerDescriptor);
    else delete globalThis.Worker;
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renderer enforces image limits before decode and keeps specific diagnostics", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  let decodeCalls = 0;
  let drawCalls = 0;
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage() { drawCalls += 1; },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => {
      decodeCalls += 1;
      return { width: 2, height: 3, close() {} };
    },
  });
  const object = (numericId, x, bytes, crop = {}) => ({
    numericId,
    id: `image:${numericId}`,
    type: "image",
    unitIndex: 0,
    bounds: { x, y: 0, width: 10, height: 10 },
    source: {
      format: "pptx",
      part: `ppt/media/image${numericId}.png`,
      kind: "shape",
      shapeId: numericId,
      mapping: "exact",
    },
    z: numericId,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes,
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
      ...crop,
    },
  });
  const unit = { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 100 };
  try {
    const pageRaster = new SceneRenderer(
      [object(0, 0, pngHeader(8_001, 4_501))],
      DEFAULT_LIMITS,
    );
    const pageRasterResult = await pageRaster.render(unit, { unitIndex: 0 });
    assert.deepEqual(pageRasterResult.diagnostics, []);
    assert.equal(decodeCalls, 1);
    assert.equal(drawCalls, 1);
    pageRasterResult.bitmap.close();
    pageRaster.close();

    const oversized = new SceneRenderer(
      [object(1, 0, pngHeader(20, 20))],
      { ...DEFAULT_LIMITS, imagePixels: 100 },
    );
    const oversizedResult = await oversized.render(unit, { unitIndex: 0 });
    assert.deepEqual(oversizedResult.diagnostics.map((item) => item.code), ["IMAGE_DIMENSION_LIMIT"]);
    assert.equal(decodeCalls, 1);

    const cumulative = new SceneRenderer(
      [
        object(1, 0, pngHeader(2, 3)),
        object(2, 20, Uint8Array.from([...pngHeader(2, 3), 1])),
      ],
      { ...DEFAULT_LIMITS, imagePixels: 6, totalImagePixels: 6 },
    );
    const cumulativeResult = await cumulative.render(unit, { unitIndex: 0 });
    assert.deepEqual(cumulativeResult.diagnostics.map((item) => item.code), ["IMAGE_TOTAL_PIXEL_LIMIT"]);
    assert.equal(decodeCalls, 2);
    cumulative.close();

    const fullyCropped = new SceneRenderer(
      [object(3, 0, pngHeader(2, 3), { cropLeft: 0.6, cropRight: 0.4 })],
      { ...DEFAULT_LIMITS, imagePixels: 6, totalImagePixels: 6 },
    );
    const drawCallsBefore = drawCalls;
    const fullyCroppedResult = await fullyCropped.render(unit, { unitIndex: 0 });
    assert.deepEqual(fullyCroppedResult.diagnostics, []);
    assert.equal(drawCalls, drawCallsBefore);
    fullyCroppedResult.bitmap.close();
    fullyCropped.close();
  } finally {
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer evicts reusable image cache before enforcing the cumulative pixel limit", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  let decodeCalls = 0;
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => {
      decodeCalls += 1;
      return { width: 2, height: 3, close() {} };
    },
  });
  const object = (numericId, unitIndex) => ({
    numericId,
    id: `image:${numericId}`,
    type: "image",
    unitIndex,
    bounds: { x: numericId * 10, y: 0, width: 10, height: 10 },
    source: {
      format: "pdf",
      part: "document.pdf",
      kind: "image",
      mapping: "exact",
    },
    z: numericId,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes: Uint8Array.from([...pngHeader(2, 3), numericId]),
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  });
  const unit = (index) => ({
    type: "page",
    index,
    id: `unit:${index}`,
    name: `Page ${index + 1}`,
    width: 100,
    height: 100,
  });
  const renderer = new SceneRenderer(
    [object(1, 0), object(2, 0), object(3, 1)],
    { ...DEFAULT_LIMITS, imagePixels: 6, totalImagePixels: 12 },
  );
  try {
    const first = await renderer.render(unit(0), { unitIndex: 0 });
    assert.deepEqual(first.diagnostics, []);
    assert.equal(first.renderedObjectCount, 2);
    first.bitmap.close();

    const second = await renderer.render(unit(1), { unitIndex: 1 });
    assert.deepEqual(second.diagnostics, []);
    assert.equal(second.renderedObjectCount, 1);
    assert.equal(decodeCalls, 3);
    second.bitmap.close();
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("closing a renderer cancels an in-flight multi-image render without leaking bitmaps", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = [];
  const closeCounts = [0, 0];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => new Promise((resolve) => decodes.push(resolve)),
  });
  const object = (numericId, x) => ({
    numericId,
    id: `image:${numericId}`,
    type: "image",
    unitIndex: 0,
    bounds: { x, y: 0, width: 10, height: 10 },
    source: {
      format: "pptx",
      part: `ppt/media/image${numericId}.png`,
      kind: "shape",
      shapeId: numericId,
      mapping: "exact",
    },
    z: numericId,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes: pngHeader(2, 3),
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  });
  const unit = { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 100 };
  const renderer = new SceneRenderer(
    [object(1, 0), object(2, 20)],
    { ...DEFAULT_LIMITS, imagePixels: 6, totalImagePixels: 12 },
  );
  try {
    const rendering = renderer.render(unit, { unitIndex: 0 });
    assert.equal(decodes.length, 1);
    renderer.close();
    decodes[0]({ width: 2, height: 3, close() { closeCounts[0] += 1; } });
    const outcome = await Promise.race([
      rendering.then(() => "resolved", (error) => error?.code),
      new Promise((resolve) => setTimeout(() => resolve("pending"), 20)),
    ]);
    if (decodes[1]) {
      decodes[1]({ width: 2, height: 3, close() { closeCounts[1] += 1; } });
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    assert.equal(outcome, "DOCUMENT_CLOSED");
    assert.equal(decodes.length, 1);
    assert.deepEqual(closeCounts, [1, 0]);
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer prefetches PDF images concurrently while preserving paint order", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = [];
  const draws = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(bitmap) { draws.push(bitmap.id); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => new Promise((resolve) => decodes.push(resolve)),
  });
  const object = (numericId) => ({
    numericId,
    id: `image:${numericId}`,
    type: "image",
    unitIndex: 0,
    bounds: { x: numericId * 10, y: 0, width: 8, height: 8 },
    source: {
      format: "pdf",
      part: "page-1",
      kind: "image",
      objectNumber: numericId,
      mapping: "exact",
    },
    z: numericId,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes: Uint8Array.from([...pngHeader(1, 1), numericId]),
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  });
  const renderer = new SceneRenderer([object(1), object(2), object(3)], DEFAULT_LIMITS);
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 100 };
  try {
    const rendering = renderer.render(unit, { unitIndex: 0 });
    assert.equal(decodes.length, 3);
    decodes[2]({ id: 3, width: 1, height: 1, close() {} });
    decodes[1]({ id: 2, width: 1, height: 1, close() {} });
    decodes[0]({ id: 1, width: 1, height: 1, close() {} });
    const result = await rendering;
    assert.deepEqual(draws, [1, 2, 3]);
    assert.equal(result.renderedObjectCount, 3);
    result.bitmap.close();
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("cancelled renders share one bounded image-prefetch window with their replacement", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = controlledBitmapDecodes();
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => decodes.decode(),
  });
  const objects = [0, 1].flatMap((unitIndex) => Array.from({ length: 9 }, (_, offset) => ({
    ...prefetchImageObject("docx", unitIndex * 100 + offset + 1, {
      bounds: { x: offset * 10, y: 0, width: 8, height: 8 },
    }),
    unitIndex,
  })));
  const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
  const unit = (index) => ({
    type: "page",
    index,
    id: `unit:${index}`,
    name: `Page ${index + 1}`,
    width: 1_000,
    height: 1_000,
  });
  let firstCancelled = false;
  let secondCancelled = false;
  let firstRendering;
  let secondRendering;
  try {
    firstRendering = renderer.render(unit(0), { unitIndex: 0 }, {
      isCancelled: () => firstCancelled,
    });
    assert.equal(decodes.pending.filter(({ settled }) => !settled).length, 8);

    firstCancelled = true;
    decodes.pending[0].resolve({ id: "cancelled:first", width: 1, height: 1, close() {} });
    await assert.rejects(firstRendering, (error) => error?.code === "OPERATION_ABORTED");

    secondRendering = renderer.render(unit(1), { unitIndex: 1 }, {
      isCancelled: () => secondCancelled,
    });
    assert.equal(
      decodes.pending.filter(({ settled }) => !settled).length,
      8,
      "replacement renders must reuse the global prefetch window instead of stacking another eight decodes",
    );
  } finally {
    secondCancelled = true;
    renderer.close();
    decodes.settleRemaining();
    await secondRendering?.catch(() => undefined);
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("demanded oversized images do not bypass the shared decode window", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = controlledBitmapDecodes();
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => decodes.decode(),
  });
  const blockingObjects = Array.from({ length: 8 }, (_, index) => prefetchImageObject("docx", index + 1, {
    bounds: { x: index * 10, y: 0, width: 8, height: 8 },
    bytes: Uint8Array.from([...pngHeader(1, 1), index]),
  }));
  const demandedObject = {
    ...prefetchImageObject("docx", 100, {
      bounds: { x: 0, y: 0, width: 8, height: 8 },
      bytes: pngHeader(5_000, 4_000),
    }),
    unitIndex: 1,
  };
  const renderer = new SceneRenderer([...blockingObjects, demandedObject], DEFAULT_LIMITS);
  const unit = (index) => ({
    type: "page",
    index,
    id: `unit:${index}`,
    name: `Page ${index + 1}`,
    width: 1_000,
    height: 1_000,
  });
  let blockingRender;
  let demandedRender;
  try {
    blockingRender = renderer.render(unit(0), { unitIndex: 0 });
    assert.equal(decodes.pending.filter(({ settled }) => !settled).length, 8);
    demandedRender = renderer.render(unit(1), { unitIndex: 1 });
    assert.equal(
      decodes.pending.filter(({ settled }) => !settled).length,
      8,
      "an oversized demanded image must wait for the existing shared decode window",
    );
    renderer.close();
    const outcome = await Promise.race([
      demandedRender.then(() => "resolved", (error) => error?.code),
      new Promise((resolve) => setTimeout(() => resolve("pending"), 20)),
    ]);
    assert.equal(outcome, "DOCUMENT_CLOSED", "closing must reject a render waiting for a decode slot");
  } finally {
    decodes.settleRemaining();
    await Promise.all([
      blockingRender?.catch(() => undefined),
      demandedRender?.catch(() => undefined),
    ]);
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("PDF soft masks share the same bounded browser-decode window as their images", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = controlledBitmapDecodes();
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => decodes.decode(),
  });
  const objects = Array.from({ length: 8 }, (_, index) => {
    const object = prefetchImageObject("docx", index + 1, {
      bounds: { x: index * 10, y: 0, width: 8, height: 8 },
    });
    return {
      ...object,
      visual: {
        ...object.visual,
        alphaMask: {
          width: 1,
          height: 1,
          mediaType: "image/png",
          bytes: Uint8Array.from([...pngHeader(1, 1), 100 + index]),
        },
      },
    };
  });
  const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 1_000, height: 1_000 };
  let rendering;
  try {
    rendering = renderer.render(unit, { unitIndex: 0 });
    assert.equal(
      decodes.pending.filter(({ settled }) => !settled).length,
      8,
      "source images and PDF soft masks must count as separate browser decodes",
    );
  } finally {
    renderer.close();
    decodes.settleRemaining();
    await rendering?.catch(() => undefined);
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("prefetched bitmaps remain open until their ordered draw consumes them", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = controlledBitmapDecodes();
  const draws = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(bitmap) {
      assert.equal(bitmap.closed, false, `bitmap ${bitmap.id} was evicted before it was drawn`);
      draws.push(bitmap.id);
    },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => decodes.decode(),
  });
  const objects = Array.from({ length: 10 }, (_, index) => prefetchImageObject("docx", index + 1, {
    bounds: { x: index * 10, y: 0, width: 8, height: 8 },
    bytes: Uint8Array.from([...pngHeader(2_000, 2_000), index]),
  }));
  const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 1_000, height: 1_000 };
  const bitmaps = Array.from({ length: objects.length }, (_, index) => ({
    id: index + 1,
    width: 2_000,
    height: 2_000,
    closed: false,
    close() { this.closed = true; },
  }));
  try {
    const rendering = renderer.render(unit, { unitIndex: 0 });
    assert.equal(decodes.pending.length, 4);
    for (let index = 1; index < 4; index += 1) {
      decodes.pending[index].resolve(bitmaps[index]);
    }
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(
      decodes.pending.length,
      4,
      "settled prefetches must retain their logical slots until ordered drawing consumes them",
    );
    decodes.pending[0].resolve(bitmaps[0]);
    for (let index = 4; index < bitmaps.length; index += 1) {
      while (decodes.pending[index] === undefined) {
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
      decodes.pending[index].resolve(bitmaps[index]);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    const result = await rendering;
    assert.deepEqual(draws, bitmaps.map(({ id }) => id));
    result.bitmap.close();
  } finally {
    decodes.settleRemaining();
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer prefetches PPTX images concurrently while preserving z-order paint", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = controlledBitmapDecodes();
  const draws = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(bitmap) { draws.push(bitmap.id); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => decodes.decode(),
  });
  const renderer = new SceneRenderer([
    prefetchImageObject("pptx", 1, { z: 30 }),
    prefetchImageObject("pptx", 2, { z: 10 }),
    prefetchImageObject("pptx", 3, { z: 20 }),
  ], DEFAULT_LIMITS);
  const unit = { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 100 };
  let rendering;
  try {
    rendering = renderer.render(unit, { unitIndex: 0 });
    assert.equal(decodes.pending.length, 3);
    decodes.pending[2].resolve({ id: "high", width: 1, height: 1, close() {} });
    decodes.pending[1].resolve({ id: "middle", width: 1, height: 1, close() {} });
    decodes.pending[0].resolve({ id: "low", width: 1, height: 1, close() {} });
    const result = await rendering;

    assert.deepEqual(draws, ["low", "middle", "high"]);
    assert.equal(result.renderedObjectCount, 3);
    result.bitmap.close();
  } finally {
    renderer.close();
    decodes.settleRemaining();
    await rendering?.catch(() => undefined);
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer prefetches only visible XLSX images", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = controlledBitmapDecodes();
  const draws = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(bitmap) { draws.push(bitmap.id); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => decodes.decode(),
  });
  const renderer = new SceneRenderer([
    prefetchImageObject("xlsx", 1, { bounds: { x: 10, y: 10, width: 8, height: 8 } }),
    prefetchImageObject("xlsx", 2, { bounds: { x: 30, y: 10, width: 8, height: 8 } }),
    prefetchImageObject("xlsx", 3, { bounds: { x: 150, y: 10, width: 8, height: 8 } }),
  ], DEFAULT_LIMITS);
  const unit = {
    type: "sheet",
    index: 0,
    id: "unit:0",
    name: "Sheet1",
    width: 200,
    height: 100,
    rows: 5,
    columns: 10,
    frozenRows: 0,
    frozenColumns: 0,
    frozenWidth: 0,
    frozenHeight: 0,
    rowAxis: { defaultSize: 20, spans: [] },
    columnAxis: { defaultSize: 20, spans: [] },
  };
  let rendering;
  try {
    rendering = renderer.render(unit, {
      unitIndex: 0,
      viewport: { x: 0, y: 0, width: 80, height: 80 },
    });
    assert.equal(decodes.pending.length, 2, "the off-viewport drawing must not be decoded");
    decodes.pending[1].resolve({ id: "visible:2", width: 1, height: 1, close() {} });
    decodes.pending[0].resolve({ id: "visible:1", width: 1, height: 1, close() {} });
    const result = await rendering;

    assert.deepEqual(draws, ["visible:1", "visible:2"]);
    assert.equal(result.renderedObjectCount, 2);
    result.bitmap.close();
  } finally {
    renderer.close();
    decodes.settleRemaining();
    await rendering?.catch(() => undefined);
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer uses the shared DOCX prefetch path and deduplicates image resources", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = controlledBitmapDecodes();
  const draws = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(bitmap) { draws.push(bitmap.id); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => decodes.decode(),
  });
  const shared = Uint8Array.from([...pngHeader(1, 1), 42]);
  const renderer = new SceneRenderer([
    prefetchImageObject("docx", 1, { bytes: shared.slice() }),
    prefetchImageObject("docx", 2, { bytes: shared.slice() }),
    prefetchImageObject("docx", 3),
  ], DEFAULT_LIMITS);
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 100 };
  let rendering;
  try {
    rendering = renderer.render(unit, { unitIndex: 0 });
    assert.equal(decodes.pending.length, 2, "content-identical DOCX images must share one decode");
    decodes.pending[1].resolve({ id: "unique", width: 1, height: 1, close() {} });
    decodes.pending[0].resolve({ id: "shared", width: 1, height: 1, close() {} });
    const result = await rendering;

    assert.deepEqual(draws, ["shared", "shared", "unique"]);
    assert.equal(result.renderedObjectCount, 3);
    result.bitmap.close();
  } finally {
    renderer.close();
    decodes.settleRemaining();
    await rendering?.catch(() => undefined);
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("failed PDF prefetch is reported without decoding the same image twice", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  let decodeCalls = 0;
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => {
      decodeCalls += 1;
      if (decodeCalls === 1) throw new Error("authored image decode failed");
      return { width: 1, height: 1, close() {} };
    },
  });
  const object = (numericId) => ({
    numericId,
    id: `image:${numericId}`,
    type: "image",
    unitIndex: 0,
    bounds: { x: numericId * 10, y: 0, width: 8, height: 8 },
    source: {
      format: "pdf",
      part: "page-1",
      kind: "image",
      objectNumber: numericId,
      mapping: "exact",
    },
    z: numericId,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes: Uint8Array.from([...pngHeader(1, 1), numericId]),
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  });
  const renderer = new SceneRenderer([object(1), object(2)], DEFAULT_LIMITS);
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 100 };
  try {
    const result = await renderer.render(unit, { unitIndex: 0 });
    assert.equal(decodeCalls, 2);
    assert.deepEqual(result.diagnostics.map(({ code }) => code), ["IMAGE_DECODE_FAILED"]);
    assert.equal(result.renderedObjectCount, 1);
    result.bitmap.close();
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer decodes a PDF image and soft mask in parallel and applies mask luminance", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const decodes = [];
  const sourcePixels = Uint8ClampedArray.from([
    10, 20, 30, 200,
    40, 50, 60, 100,
  ]);
  const maskPixels = Uint8ClampedArray.from([
    64, 64, 64, 255,
    192, 192, 192, 255,
  ]);
  let compositedPixels;
  let sourceCloses = 0;
  let maskCloses = 0;
  let compositedCloses = 0;
  const compositedBitmap = {
    width: 2,
    height: 1,
    close() { compositedCloses += 1; },
  };
  const finalContext = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  const sourceContext = {
    drawImage() {},
    getImageData() { return { data: sourcePixels.slice() }; },
    putImageData(image) { compositedPixels = image.data.slice(); },
  };
  const maskContext = {
    drawImage() {},
    getImageData() { return { data: maskPixels.slice() }; },
  };
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
    }
    getContext(_kind, options) {
      if (this.width !== 2 || this.height !== 1) return finalContext;
      return options?.alpha === false ? maskContext : sourceContext;
    }
    transferToImageBitmap() {
      return this.width === 2 && this.height === 1 ? compositedBitmap : { close() {} };
    }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: () => new Promise((resolve) => decodes.push(resolve)),
  });
  const renderer = new SceneRenderer([{
    numericId: 1,
    id: "image:1",
    type: "image",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 20, height: 10 },
    source: {
      format: "pdf",
      part: "page-1",
      kind: "image",
      objectNumber: 1,
      mapping: "exact",
    },
    z: 0,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes: pngHeader(2, 1),
      alphaMask: {
        width: 2,
        height: 1,
        mediaType: "image/png",
        bytes: Uint8Array.from([...pngHeader(2, 1), 1]),
      },
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  }], DEFAULT_LIMITS);
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 100, height: 100 };
  try {
    const rendering = renderer.render(unit, { unitIndex: 0 });
    assert.equal(decodes.length, 2);
    decodes[0]({ width: 2, height: 1, close() { sourceCloses += 1; } });
    decodes[1]({ width: 2, height: 1, close() { maskCloses += 1; } });
    const result = await rendering;
    assert.deepEqual([...compositedPixels], [
      10, 20, 30, 50,
      40, 50, 60, 75,
    ]);
    assert.equal(sourceCloses, 1);
    assert.equal(maskCloses, 1);
    assert.equal(compositedCloses, 0);
    result.bitmap.close();
    renderer.close();
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(compositedCloses, 1);
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer preserves native bitmap pixels across subpixel extent rounding", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const bitmap = { width: 4, height: 2, close() {} };
  const draws = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(...args) { draws.push(args); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", { configurable: true, value: async () => bitmap });
  const object = (numericId, x, width, height) => ({
    numericId,
    id: `image:${numericId}`,
    type: "image",
    unitIndex: 0,
    bounds: { x, y: 0, width, height },
    source: {
      format: "pptx",
      part: `ppt/media/image${numericId}.png`,
      kind: "shape",
      shapeId: numericId,
      mapping: "exact",
    },
    z: numericId,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes: pngHeader(4, 2),
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  });
  const renderer = new SceneRenderer([
    object(1, 0, 4.08, 2.08),
    object(2, 10, 4.2, 2.2),
    object(3, 20, 4.08, 2.2),
  ], DEFAULT_LIMITS);
  const unit = { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 100 };
  try {
    const result = await renderer.render(unit, { unitIndex: 0 });
    assert.deepEqual(draws.map((draw) => draw.slice(1)), [
      [0, 0, 4, 2],
      [10, 0, 4.2, 2.2],
      [20, 0, 4.08, 2.2],
    ]);
    result.bitmap.close();
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer forwards host-provided fonts for generated PDF XFA SVG pages", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const workerDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const fetchDescriptor = Object.getOwnPropertyDescriptor(globalThis, "fetch");
  const draws = [];
  const requests = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(...args) { draws.push(args); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 100, height: 100, close() {} }),
  });
  Object.defineProperty(globalThis, "fetch", {
    configurable: true,
    value: async () => ({ ok: true, arrayBuffer: async () => Uint8Array.of(1, 2, 3).buffer }),
  });
  Object.defineProperty(globalThis, "Worker", {
    configurable: true,
    value: class {
      postMessage(request) {
        requests.push(request);
        queueMicrotask(() => this.onmessage?.({ data: {
          id: request.id,
          ok: true,
          bitmap: { width: 100, height: 100, close() {} },
          approximate: true,
        } }));
      }
      terminate() {}
    },
  });
  const renderer = new SceneRenderer([{
    numericId: 0,
    id: "pdf:0:0",
    type: "image",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 100, height: 100 },
    source: { format: "pdf", part: "document.pdf", kind: "form", mapping: "derived" },
    z: 0,
    text: "Filled XFA",
    visual: {
      kind: "image",
      mediaType: "image/svg+xml",
      bytes: new TextEncoder().encode(
        '<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text>Filled XFA</text></svg>',
      ),
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  }], DEFAULT_LIMITS, undefined, [{
    family: "Noto Sans S Chinese",
    bytes: Uint8Array.of(1, 2, 3),
  }]);
  try {
    const result = await renderer.render(
      { type: "page", index: 0, id: "page:0", name: "Page 1", width: 100, height: 100 },
      { unitIndex: 0 },
    );
    assert.equal(draws.length, 1);
    assert.equal(requests.length, 1);
    assert.equal(requests[0].fonts[0]?.family, "Noto Sans S Chinese", JSON.stringify(requests));
    result.bitmap.close();
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
    if (workerDescriptor) Object.defineProperty(globalThis, "Worker", workerDescriptor);
    else delete globalThis.Worker;
    if (fetchDescriptor) Object.defineProperty(globalThis, "fetch", fetchDescriptor);
    else delete globalThis.fetch;
  }
});

test("renderer rounds axis-aligned PDF image bounds in device pixels", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const bitmap = { width: 4, height: 2, close() {} };
  const draws = [];
  const transforms = [];
  let transform = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const context = {
    save() {}, restore() {}, translate() {}, fillRect() {},
    scale(x, y) { transform = { ...transform, a: transform.a * x, d: transform.d * y }; },
    getTransform() { return transform; },
    setTransform(a, b, c, d, e, f) {
      transform = { a, b, c, d, e, f };
      transforms.push([a, b, c, d, e, f]);
    },
    drawImage(...args) { draws.push(args); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", { configurable: true, value: async () => bitmap });
  const renderer = new SceneRenderer([{
    numericId: 1,
    id: "pdf:image:1",
    type: "image",
    unitIndex: 0,
    bounds: { x: 1.4, y: 2.4, width: 4.4, height: 2.4 },
    source: {
      format: "pdf",
      part: "document.pdf",
      kind: "image",
      objectNumber: 1,
      mapping: "exact",
    },
    z: 0,
    visual: {
      kind: "image",
      mediaType: "image/png",
      bytes: pngHeader(4, 2),
      cropLeft: 0,
      cropTop: 0,
      cropRight: 0,
      cropBottom: 0,
    },
  }], DEFAULT_LIMITS);
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 20, height: 20.25 };
  try {
    const result = await renderer.render(unit, { unitIndex: 0, scale: 0.75 });
    assert.deepEqual(transforms, [
      [1, 0, 0, 1, 1, 2],
      [0.75, 0, 0, 0.75, 0, 0],
    ]);
    assert.deepEqual(draws[0].slice(1), [0, 0, 4, 2, 0, 0, 3, 2]);
    result.bitmap.close();
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer preserves DrawingML contrast while shifting image brightness", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const sourcePixels = Uint8ClampedArray.from([64, 128, 192, 255]);
  let processedPixels;
  const effectContext = {
    drawImage() {},
    getImageData() { return { data: sourcePixels.slice() }; },
    putImageData(image) { processedPixels = image.data.slice(); },
  };
  const finalContext = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; }
    getContext() { return this.width === 1 && this.height === 1 ? effectContext : finalContext; }
    transferToImageBitmap() { return { width: this.width, height: this.height, close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 1, height: 1, close() {} }),
  });
  const renderer = new SceneRenderer([{
    numericId: 1,
    id: "image:1",
    type: "image",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 20, height: 20 },
    source: {
      format: "xlsx",
      part: "xl/drawings/drawing1.xml",
      kind: "shape",
      shapeId: 1,
      mapping: "exact",
    },
    z: 0,
    visual: {
      kind: "image-adjustment",
      grayscale: false,
      brightness: 0.7,
      contrast: -0.7,
      visual: {
        kind: "image",
        mediaType: "image/png",
        bytes: pngHeader(1, 1),
        cropLeft: 0,
        cropTop: 0,
        cropRight: 0,
        cropBottom: 0,
      },
    },
  }], DEFAULT_LIMITS);
  try {
    const result = await renderer.render(
      { type: "page", index: 0, id: "unit:0", name: "Sheet1", width: 20, height: 20 },
      { unitIndex: 0 },
    );
    assert.deepEqual([...processedPixels], [224, 244, 255, 255]);
    result.bitmap.close();
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer applies DrawingML color replacement without source-alpha matching when useA is false", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const sourcePixels = Uint8ClampedArray.from([
    255, 255, 255, 200,
    254, 255, 255, 200,
  ]);
  let processedPixels;
  let decodedCloses = 0;
  let processedCloses = 0;
  const processedBitmap = { width: 2, height: 1, close() { processedCloses += 1; } };
  const finalDraws = [];
  const finalContext = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(...args) { finalDraws.push(args); },
  };
  const effectContext = {
    drawImage() {},
    getImageData() { return { data: sourcePixels.slice() }; },
    putImageData(image) { processedPixels = image.data.slice(); },
  };
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
    }
    getContext() { return this.width === 2 && this.height === 1 ? effectContext : finalContext; }
    transferToImageBitmap() {
      return this.width === 2 && this.height === 1 ? processedBitmap : { close() {} };
    }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 2, height: 1, close() { decodedCloses += 1; } }),
  });
  const renderer = new SceneRenderer([{
    numericId: 1,
    id: "image:1",
    type: "image",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 20, height: 10 },
    source: {
      format: "pptx",
      part: "ppt/media/image1.png",
      kind: "shape",
      shapeId: 1,
      mapping: "exact",
    },
    z: 0,
    visual: {
      kind: "image-color-change",
      from: 0xffff_ffff,
      to: 0x3366_9980,
      useAlpha: false,
      visual: {
        kind: "image",
        mediaType: "image/png",
        bytes: pngHeader(2, 1),
        cropLeft: 0,
        cropTop: 0,
        cropRight: 0,
        cropBottom: 0,
      },
    },
  }], { ...DEFAULT_LIMITS, imagePixels: 2, totalImagePixels: 6 });
  const unit = { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 100 };
  try {
    const result = await renderer.render(unit, { unitIndex: 0 });
    assert.deepEqual([...processedPixels], [
      51, 102, 153, 100,
      254, 255, 255, 200,
    ]);
    assert.equal(finalDraws.at(-1)?.[0], processedBitmap);
    assert.equal(decodedCloses, 1);
    assert.equal(result.diagnostics.length, 0);
    renderer.close();
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(processedCloses, 1);
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer includes source alpha in DrawingML color matching when useA is true", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const sourcePixels = Uint8ClampedArray.from([
    255, 255, 255, 255,
    255, 255, 255, 128,
  ]);
  let processedPixels;
  const finalContext = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage() {},
  };
  const effectContext = {
    drawImage() {},
    getImageData() { return { data: sourcePixels.slice() }; },
    putImageData(image) { processedPixels = image.data.slice(); },
  };
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
    }
    getContext() { return this.width === 2 && this.height === 1 ? effectContext : finalContext; }
    transferToImageBitmap() { return { width: this.width, height: this.height, close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 2, height: 1, close() {} }),
  });
  const renderer = new SceneRenderer([{
    numericId: 1,
    id: "image:1",
    type: "image",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 20, height: 10 },
    source: {
      format: "pptx",
      part: "ppt/media/image1.png",
      kind: "shape",
      shapeId: 1,
      mapping: "exact",
    },
    z: 0,
    visual: {
      kind: "image-color-change",
      from: 0xffff_ff80,
      to: 0x3366_9980,
      useAlpha: true,
      visual: {
        kind: "image",
        mediaType: "image/png",
        bytes: pngHeader(2, 1),
        cropLeft: 0,
        cropTop: 0,
        cropRight: 0,
        cropBottom: 0,
      },
    },
  }], { ...DEFAULT_LIMITS, imagePixels: 2, totalImagePixels: 6 });
  const unit = { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 100 };
  try {
    const result = await renderer.render(unit, { unitIndex: 0 });
    assert.deepEqual([...processedPixels], [
      255, 255, 255, 255,
      51, 102, 153, 64,
    ]);
    assert.equal(result.diagnostics.length, 0);
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer budgets the full DrawingML color-change working set and releases failed decodes", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  let decodeCalls = 0;
  let decodedCloses = 0;
  let effectCanvases = 0;
  let effectTransfers = 0;
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  const effectContext = {
    drawImage() {},
    getImageData() {
      return { data: Uint8ClampedArray.from([255, 255, 255, 255, 255, 255, 255, 255]) };
    },
    putImageData() {},
  };
  class FakeCanvas {
    constructor(width, height) {
      this.width = width;
      this.height = height;
      if (width === 2 && height === 1) effectCanvases += 1;
    }
    getContext() { return this.width === 2 && this.height === 1 ? effectContext : context; }
    transferToImageBitmap() {
      if (this.width === 2 && this.height === 1) effectTransfers += 1;
      return { close() {} };
    }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => {
      decodeCalls += 1;
      return { width: 2, height: 1, close() { decodedCloses += 1; } };
    },
  });
  const imageVisual = {
    kind: "image",
    mediaType: "image/png",
    bytes: pngHeader(2, 1),
    cropLeft: 0,
    cropTop: 0,
    cropRight: 0,
    cropBottom: 0,
  };
  const object = (numericId, visual) => ({
    numericId,
    id: `image:${numericId}`,
    type: "image",
    unitIndex: 0,
    bounds: { x: (numericId - 1) * 20, y: 0, width: 20, height: 10 },
    source: {
      format: "pptx",
      part: `ppt/media/image${numericId}.png`,
      kind: "shape",
      shapeId: numericId,
      mapping: "exact",
    },
    z: numericId,
    visual,
  });
  const renderer = new SceneRenderer([
    object(1, {
      kind: "image-color-change",
      from: 0xffff_ffff,
      to: 0xffff_ff00,
      useAlpha: true,
      visual: imageVisual,
    }),
    object(2, imageVisual),
  ], { ...DEFAULT_LIMITS, imagePixels: 2, totalImagePixels: 3 });
  const unit = { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 100, height: 100 };
  try {
    const result = await renderer.render(unit, { unitIndex: 0 });
    assert.deepEqual(result.diagnostics.map(({ code }) => code), ["IMAGE_TOTAL_PIXEL_LIMIT"]);
    assert.equal(result.renderedObjectCount, 1);
    assert.equal(decodeCalls, 2);
    assert.equal(decodedCloses, 1);
    assert.equal(effectCanvases, 0);

    renderer.close();
    const bitmapLimited = new SceneRenderer([object(3, {
      kind: "image-color-change",
      from: 0xffff_ffff,
      to: 0xffff_ff00,
      useAlpha: true,
      visual: imageVisual,
    })], { ...DEFAULT_LIMITS, imagePixels: 2, totalImagePixels: 5 });
    try {
      const bitmapLimitedResult = await bitmapLimited.render(unit, { unitIndex: 0 });
      assert.deepEqual(
        bitmapLimitedResult.diagnostics.map(({ code }) => code),
        ["IMAGE_TOTAL_PIXEL_LIMIT"],
      );
      assert.equal(bitmapLimitedResult.renderedObjectCount, 0);
      assert.equal(decodeCalls, 3);
      assert.equal(effectCanvases, 1);
      assert.equal(effectTransfers, 0);
    } finally {
      bitmapLimited.close();
    }
  } finally {
    renderer.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("renderer retains decoded sources across thumbnails, zoom and scene replacement", async () => {
  const descriptors = ["OffscreenCanvas", "createImageBitmap", "ImageBitmap", "structuredClone", "navigator"].map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)]);
  Object.defineProperty(globalThis, "navigator", { configurable: true, value: { userAgent: "Chrome/140.0" } });
  let decodes = 0;
  let liveSources = 0;
  class FakeBitmap {
    width = 5000; height = 5000; closed = false;
    constructor() { liveSources += 1; }
    close() { if (!this.closed) { this.closed = true; liveSources -= 1; } }
  }
  Object.defineProperty(globalThis, "ImageBitmap", { configurable: true, value: FakeBitmap });
  Object.defineProperty(globalThis, "structuredClone", { configurable: true, value: () => new FakeBitmap() });
  const context = { save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {} };
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; }
    getContext() { return context; }
    transferToImageBitmap() { return { width: this.width, height: this.height, close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", { configurable: true, value: async source => {
    if (!(source instanceof Blob)) return { width: source.width, height: source.height, close() {} };
    decodes += 1;
    return new FakeBitmap();
  } });
  const objects = [0, 1].map(index => ({
    numericId: index + 1, id: `image:${index}`, type: "image", unitIndex: index,
    bounds: { x: 0, y: 0, width: 1000, height: 1000 }, z: 0,
    source: { format: "pptx", part: `ppt/media/image${index}.png`, kind: "shape", shapeId: index + 1, mapping: "exact" },
    visual: { kind: "image", mediaType: "image/png", bytes: Uint8Array.from([...pngHeader(5000, 5000), index]), cropLeft: 0, cropTop: 0, cropRight: 0, cropBottom: 0 },
  }));
  const renderer = new SceneRenderer(objects, { ...DEFAULT_LIMITS, totalImagePixels: 80_000_000 });
  try {
    for (const scale of [0.1, 1, 2, 0.5]) {
      if (scale === 0.5) renderer.replaceObjects(objects.map(object => ({ ...object, bounds: { ...object.bounds, x: 1 } })), []);
      for (const index of [0, 1]) {
        const frame = await renderer.render({ type: "slide", index, id: `unit:${index}`, name: "Slide", width: 1000, height: 1000 }, { unitIndex: index, scale });
        assert.deepEqual(frame.diagnostics, []);
        frame.bitmap.close();
      }
      assert.equal(decodes, 2, `scale ${scale} must reuse both large decoded images`);
      assert.equal(liveSources, 2, "only the two cached sources survive each render; temporary copies are released");
    }
    renderer.close();
    await Promise.resolve();
    assert.equal(liveSources, 0, "closing the document releases both cached sources");
  } finally {
    renderer.close();
    for (const [name, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor);
      else delete globalThis[name];
    }
  }
});
