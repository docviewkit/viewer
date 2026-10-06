import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { deflateSync, gzipSync } from "node:zlib";
import test from "node:test";

import { decodeOfficeImagePayload } from "../dist/image-codecs.js";
import { decodeOfficeImage } from "../dist/image-codec-client.js";
import { DEFAULT_LIMITS } from "../dist/core.js";
import { identifyImage } from "../dist/image.js";

function codecInput(format, mediaType, bytes, overrides = {}) {
  return {
    format,
    mediaType,
    bytes,
    maxPixels: 1_000_000,
    maxBytes: 1_000_000,
    maxCompressionRatio: 200,
    ...overrides,
  };
}

function pdfRaster(width, height, components, predictor, samples, palette) {
  const compressed = new Uint8Array(deflateSync(samples));
  const paletteBytes = palette?.bytes ?? new Uint8Array();
  const headerSize = palette === undefined ? 20 : 28;
  const bytes = new Uint8Array(headerSize + compressed.length + paletteBytes.length);
  bytes.set(new TextEncoder().encode("OVPDFR01"));
  const view = new DataView(bytes.buffer);
  view.setUint32(8, width, false);
  view.setUint32(12, height, false);
  bytes[16] = components;
  bytes[17] = predictor;
  if (palette !== undefined) {
    bytes[18] = palette.components;
    bytes[19] = palette.high;
    view.setFloat32(20, palette.decode[0], false);
    view.setFloat32(24, palette.decode[1], false);
  }
  bytes.set(compressed, headerSize);
  bytes.set(paletteBytes, headerSize + compressed.length);
  return bytes;
}

function bmp24() {
  const bytes = new Uint8Array(62);
  const view = new DataView(bytes.buffer);
  bytes.set([0x42, 0x4d]);
  view.setUint32(2, bytes.length, true);
  view.setUint32(10, 54, true);
  view.setUint32(14, 40, true);
  view.setInt32(18, 2, true);
  view.setInt32(22, 1, true);
  view.setUint16(26, 1, true);
  view.setUint16(28, 24, true);
  view.setUint32(34, 8, true);
  bytes.set([0, 0, 255, 0, 255, 0, 0, 0], 54);
  return bytes;
}

function dibRle8() {
  const bytes = new Uint8Array(52);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 40, true);
  view.setInt32(4, 2, true);
  view.setInt32(8, 1, true);
  view.setUint16(12, 1, true);
  view.setUint16(14, 8, true);
  view.setUint32(16, 1, true);
  view.setUint32(20, 4, true);
  view.setUint32(32, 2, true);
  bytes.set([0, 0, 0, 0, 0, 0, 255, 0], 40);
  bytes.set([2, 1, 0, 1], 48);
  return bytes;
}

function dibCore1() {
  const bytes = new Uint8Array(22);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 12, true);
  view.setUint16(4, 1, true);
  view.setUint16(6, 1, true);
  view.setUint16(8, 1, true);
  view.setUint16(10, 1, true);
  bytes.set([0, 0, 0, 0, 0, 255], 12);
  bytes[18] = 0x80;
  return bytes;
}

function dibOs2Short1() {
  const bytes = new Uint8Array(28);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 16, true);
  view.setInt32(4, 1, true);
  view.setInt32(8, 1, true);
  view.setUint16(12, 1, true);
  view.setUint16(14, 1, true);
  bytes.set([0, 0, 0, 0, 0, 0, 255, 0], 16);
  bytes[24] = 0x80;
  return bytes;
}

function pcxIndexed() {
  const bytes = new Uint8Array(128 + 2 + 769);
  const view = new DataView(bytes.buffer);
  bytes.set([0x0a, 0x05, 0x01, 0x08]);
  view.setUint16(8, 1, true);
  bytes[65] = 1;
  view.setUint16(66, 2, true);
  view.setUint16(68, 1, true);
  bytes.set([1, 2], 128);
  bytes[130] = 0x0c;
  bytes.set([255, 0, 0], 131 + 3);
  bytes.set([0, 255, 0], 131 + 6);
  return bytes;
}

function tiffRgb() {
  const entryCount = 10;
  const bitsOffset = 8 + 2 + entryCount * 12 + 4;
  const pixelOffset = bitsOffset + 6;
  const bytes = new Uint8Array(pixelOffset + 6);
  const view = new DataView(bytes.buffer);
  bytes.set([0x49, 0x49]);
  view.setUint16(2, 42, true);
  view.setUint32(4, 8, true);
  view.setUint16(8, entryCount, true);
  const entries = [
    [256, 4, 1, 2],
    [257, 4, 1, 1],
    [258, 3, 3, bitsOffset],
    [259, 3, 1, 1],
    [262, 3, 1, 2],
    [273, 4, 1, pixelOffset],
    [277, 3, 1, 3],
    [278, 4, 1, 1],
    [279, 4, 1, 6],
    [284, 3, 1, 1],
  ];
  entries.forEach(([tag, type, count, value], index) => {
    const offset = 10 + index * 12;
    view.setUint16(offset, tag, true);
    view.setUint16(offset + 2, type, true);
    view.setUint32(offset + 4, count, true);
    if (type === 3 && count === 1) view.setUint16(offset + 8, value, true);
    else view.setUint32(offset + 8, value, true);
  });
  view.setUint16(bitsOffset, 8, true);
  view.setUint16(bitsOffset + 2, 8, true);
  view.setUint16(bitsOffset + 4, 8, true);
  bytes.set([255, 0, 0, 0, 255, 0], pixelOffset);
  return bytes;
}

function tiffPackBitsOriented() {
  const entryCount = 11;
  const bitsOffset = 8 + 2 + entryCount * 12 + 4;
  const pixelOffset = bitsOffset + 6;
  const bytes = new Uint8Array(pixelOffset + 7);
  const view = new DataView(bytes.buffer);
  bytes.set([0x49, 0x49]);
  view.setUint16(2, 42, true);
  view.setUint32(4, 8, true);
  view.setUint16(8, entryCount, true);
  const entries = [
    [256, 4, 1, 2],
    [257, 4, 1, 1],
    [258, 3, 3, bitsOffset],
    [259, 3, 1, 32773],
    [262, 3, 1, 2],
    [273, 4, 1, pixelOffset],
    [274, 3, 1, 6],
    [277, 3, 1, 3],
    [278, 4, 1, 1],
    [279, 4, 1, 7],
    [284, 3, 1, 1],
  ];
  entries.forEach(([tag, type, count, value], index) => {
    const offset = 10 + index * 12;
    view.setUint16(offset, tag, true);
    view.setUint16(offset + 2, type, true);
    view.setUint32(offset + 4, count, true);
    if (type === 3 && count === 1) view.setUint16(offset + 8, value, true);
    else view.setUint32(offset + 8, value, true);
  });
  view.setUint16(bitsOffset, 8, true);
  view.setUint16(bitsOffset + 2, 8, true);
  view.setUint16(bitsOffset + 4, 8, true);
  bytes.set([5, 255, 0, 0, 0, 255, 0], pixelOffset);
  return bytes;
}

function tiffCompressedClaim(width, height, byteCount) {
  const entryCount = 9;
  const pixelOffset = 8 + 2 + entryCount * 12 + 4;
  const bytes = new Uint8Array(pixelOffset + byteCount);
  const view = new DataView(bytes.buffer);
  bytes.set([0x49, 0x49]);
  view.setUint16(2, 42, true);
  view.setUint32(4, 8, true);
  view.setUint16(8, entryCount, true);
  const entries = [
    [256, 4, 1, width],
    [257, 4, 1, height],
    [258, 3, 1, 8],
    [259, 3, 1, 5],
    [262, 3, 1, 1],
    [273, 4, 1, pixelOffset],
    [277, 3, 1, 1],
    [278, 4, 1, height],
    [279, 4, 1, byteCount],
  ];
  entries.forEach(([tag, type, count, value], index) => {
    const offset = 10 + index * 12;
    view.setUint16(offset, tag, true);
    view.setUint16(offset + 2, type, true);
    view.setUint32(offset + 4, count, true);
    if (type === 3) view.setUint16(offset + 8, value, true);
    else view.setUint32(offset + 8, value, true);
  });
  return bytes;
}

function ico32() {
  const image = new Uint8Array(48);
  const imageView = new DataView(image.buffer);
  imageView.setUint32(0, 40, true);
  imageView.setInt32(4, 1, true);
  imageView.setInt32(8, 2, true);
  imageView.setUint16(12, 1, true);
  imageView.setUint16(14, 32, true);
  image.set([0, 0, 255, 255], 40);

  const bytes = new Uint8Array(22 + image.length);
  const view = new DataView(bytes.buffer);
  view.setUint16(2, 1, true);
  view.setUint16(4, 1, true);
  bytes[6] = 1;
  bytes[7] = 1;
  view.setUint16(10, 1, true);
  view.setUint16(12, 32, true);
  view.setUint32(14, image.length, true);
  view.setUint32(18, 22, true);
  bytes.set(image, 22);
  return bytes;
}

function deceptiveIcoPng(width, height) {
  const image = new Uint8Array(24);
  image.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const imageView = new DataView(image.buffer);
  imageView.setUint32(8, 13, false);
  image.set([0x49, 0x48, 0x44, 0x52], 12);
  imageView.setUint32(16, width, false);
  imageView.setUint32(20, height, false);
  const bytes = new Uint8Array(22 + image.length);
  const view = new DataView(bytes.buffer);
  view.setUint16(2, 1, true);
  view.setUint16(4, 1, true);
  bytes[6] = 1;
  bytes[7] = 1;
  view.setUint16(12, 32, true);
  view.setUint32(14, image.length, true);
  view.setUint32(18, 22, true);
  bytes.set(image, 22);
  return bytes;
}

function jpeg2000() {
  return Uint8Array.from(Buffer.from(
    "AAAADGpQICANCocKAAAAFGZ0eXBqcDIgAAAAAGpwMiAAAAAtanAyaAAAABZpaGRyAAAAAQAAAAIAAwgHAAAAAAAPY29scgEAAAAAABAAAAC+anAyY/9P/1EALwAAAAAAAgAAAAEAAAAAAAAAAAAAAQAAAAEAAAAAAAAAAAAAAwcBAQcBAQcBAf9SAAwAAAABAAYCAgAA/1wAKSJ/IH7gfuB+oHbwdvB2wG8AbwBu4GdQZ1BnaFAFUAVQR1fTV9NXYv9kABEAAUxhdmM2Mi4yOC4xMDH/kAAKAAAAAAA1AAH/k9/4MBgEox/f+DAYBziH3/gwGAc4hwAAAAAAAAAAAAAAAAAAAAAAAP/Z",
    "base64",
  ));
}

function jp2Codestream(bytes) {
  const marker = new TextEncoder().encode("jp2c");
  for (let offset = 4; offset + 4 <= bytes.length; offset += 1) {
    if (marker.every((value, index) => bytes[offset + index] === value)) {
      const boxStart = offset - 4;
      const boxLength = new DataView(bytes.buffer, bytes.byteOffset + boxStart, 4).getUint32(0, false);
      return bytes.subarray(offset + 4, boxStart + boxLength);
    }
  }
  throw new Error("JP2 fixture contains no codestream box");
}

function boxedJpxReduction(bytes, reducePower) {
  const marker = new TextEncoder().encode("OVPJPXRD");
  const output = new Uint8Array(bytes.length + 17);
  output.set(bytes);
  const view = new DataView(output.buffer);
  view.setUint32(bytes.length, 17, false);
  output.set(new TextEncoder().encode("free"), bytes.length + 4);
  output.set(marker, bytes.length + 8);
  output[bytes.length + 16] = reducePower;
  return output;
}

function pngHeaderDataUri(width, height) {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const view = new DataView(bytes.buffer);
  view.setUint32(8, 13, false);
  bytes.set([0x49, 0x48, 0x44, 0x52], 12);
  view.setUint32(16, width, false);
  view.setUint32(20, height, false);
  return `data:image/png;base64,${Buffer.from(bytes).toString("base64")}`;
}

test("decodes BMP, DIB/RLE, ICO, PCX, and TIFF pixels in the Office codec", async () => {
  const cases = [
    ["bmp", "image/bmp", bmp24(), [255, 0, 0, 255, 0, 255, 0, 255]],
    ["dib", "image/bmp", dibRle8(), [255, 0, 0, 255, 255, 0, 0, 255]],
    ["dib", "image/bmp", dibCore1(), [255, 0, 0, 255]],
    ["dib", "image/bmp", dibOs2Short1(), [255, 0, 0, 255]],
    ["ico", "image/x-icon", ico32(), [255, 0, 0, 255]],
    ["pcx", "image/x-pcx", pcxIndexed(), [255, 0, 0, 255, 0, 255, 0, 255]],
    ["tiff", "image/tiff", tiffRgb(), [255, 0, 0, 255, 0, 255, 0, 255]],
  ];
  for (const [format, mediaType, bytes, expected] of cases) {
    const output = await decodeOfficeImagePayload(codecInput(format, mediaType, bytes));
    assert.equal(output.kind, "rgba", format);
    assert.deepEqual([...output.data], expected, format);
    assert.equal(output.approximate, false, format);
  }
});

test("decodes PackBits TIFF pixels and applies orientation before returning RGBA", async () => {
  const output = await decodeOfficeImagePayload(codecInput("tiff", "image/tiff", tiffPackBitsOriented()));
  assert.equal(output.kind, "rgba");
  assert.equal(output.width, 1);
  assert.equal(output.height, 2);
  assert.deepEqual([...output.data], [255, 0, 0, 255, 0, 255, 0, 255]);
});

test("rejects TIFF strip claims that exceed the decode ratio before decompression", async () => {
  await assert.rejects(
    decodeOfficeImagePayload(codecInput("tiff", "image/tiff", tiffCompressedClaim(100, 100, 1), {
      maxPixels: 10_000,
      maxCompressionRatio: 2,
    })),
    (error) => error?.code === "IMAGE_DECOMPRESSION_LIMIT",
  );
});

test("validates the selected ICO payload dimensions before image decode", async () => {
  await assert.rejects(
    decodeOfficeImagePayload(codecInput("ico", "image/x-icon", deceptiveIcoPng(1_000, 1_000), { maxPixels: 10_000 })),
    (error) => error?.code === "IMAGE_DIMENSION_LIMIT",
  );
});

test("rejects oversized raster data URIs before SVG rasterization", async () => {
  const dataUri = pngHeaderDataUri(2_000, 2_000);
  for (const body of [
    `<image href="${dataUri}"/>`,
    `<filter id="f"><feImage href="${dataUri}"/></filter>`,
    `<style>.paint{fill:url('${dataUri}')}</style><rect class="paint" width="1" height="1"/>`,
  ]) {
    const bytes = new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1">${body}</svg>`);
    await assert.rejects(
      decodeOfficeImagePayload(codecInput("svg", "image/svg+xml", bytes, { maxPixels: 1_000_000 })),
      (error) => error?.code === "IMAGE_DIMENSION_LIMIT",
    );
  }
});

test("enforces the cumulative SVG data-image pixel budget", async () => {
  const dataUri = pngHeaderDataUri(800, 800);
  const source = `<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><image href="${dataUri}"/><image href="${dataUri}"/></svg>`;
  const bytes = new TextEncoder().encode(source);
  await assert.rejects(
    decodeOfficeImagePayload(codecInput("svg", "image/svg+xml", bytes, { maxPixels: 1_000_000 })),
    (error) => error?.code === "IMAGE_DIMENSION_LIMIT",
  );
});

test("rasterizes BMP and TIFF data images embedded in SVG", async () => {
  const nativeFetch = globalThis.fetch;
  globalThis.fetch = async (input, init) => input instanceof URL && input.protocol === "file:"
    ? new Response(await readFile(input), { headers: { "content-type": "application/wasm" } })
    : nativeFetch(input, init);
  try {
    for (const [mediaType, image] of [["image/bmp", bmp24()], ["image/tiff", tiffRgb()]]) {
      const dataUri = `data:${mediaType};base64,${Buffer.from(image).toString("base64")}`;
      const source = `<svg xmlns="http://www.w3.org/2000/svg" width="2" height="1"><image href="${dataUri}" width="2" height="1"/></svg>`;
      const output = await decodeOfficeImagePayload(codecInput(
        "svg",
        "image/svg+xml",
        new TextEncoder().encode(source),
      ));
      assert.equal(output.kind, "encoded");
      assert.deepEqual([...output.data.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
    }
  } finally {
    globalThis.fetch = nativeFetch;
  }
});

test("requires shared document or host font assets for SVG text", async () => {
  const bytes = new TextEncoder().encode(
    '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><text x="0" y="9">Text</text></svg>',
  );
  await assert.rejects(
    decodeOfficeImagePayload(codecInput("svg", "image/svg+xml", bytes)),
    (error) => error?.code === "IMAGE_SVG_FONT_UNAVAILABLE",
  );
});

test("blocks SVG font, external CSS, and unsupported data resources", async () => {
  for (const body of [
    '<style>@font-face{font-family:x;src:url(data:font/woff2;base64,AA==)}</style>',
    '<image href="data:text/plain;base64,QQ=="/>',
    '<style>.paint{fill:url(https://example.invalid/image.png)}</style>',
  ]) {
    const bytes = new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1">${body}</svg>`);
    await assert.rejects(
      decodeOfficeImagePayload(codecInput("svg", "image/svg+xml", bytes)),
      (error) => error?.code === "IMAGE_EXTERNAL_RESOURCE_BLOCKED",
    );
  }
});

test("always routes ICO decoding through the isolated codec Worker", async () => {
  const workerDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const bitmapDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  let workerCalls = 0;
  let bitmapCalls = 0;
  class FakeWorker {
    onmessage = null;
    onerror = null;
    onmessageerror = null;
    constructor() { workerCalls += 1; }
    postMessage() {
      queueMicrotask(() => this.onmessage?.({ data: { id: 1, ok: false, code: "TEST_STOP", message: "stop" } }));
    }
    terminate() {}
  }
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: FakeWorker });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => {
      bitmapCalls += 1;
      return { width: 1, height: 1, close() {} };
    },
  });
  try {
    const bytes = ico32();
    await assert.rejects(
      decodeOfficeImage(bytes, identifyImage(bytes, "image/x-icon", 10), DEFAULT_LIMITS),
      (error) => error?.code === "TEST_STOP",
    );
    assert.equal(workerCalls, 1);
    assert.equal(bitmapCalls, 0);
  } finally {
    if (workerDescriptor === undefined) delete globalThis.Worker;
    else Object.defineProperty(globalThis, "Worker", workerDescriptor);
    if (bitmapDescriptor === undefined) delete globalThis.createImageBitmap;
    else Object.defineProperty(globalThis, "createImageBitmap", bitmapDescriptor);
  }
});

test("stops compressed Office metafiles before their inflated bytes exceed the configured budget", async () => {
  const bytes = new Uint8Array(gzipSync(new Uint8Array(4_096)));
  await assert.rejects(
    decodeOfficeImagePayload(codecInput("emz", "image/x-emz", bytes, {
      maxBytes: 1_000,
      maxCompressionRatio: 2,
    })),
    (error) => error?.code === "IMAGE_DECOMPRESSION_LIMIT",
  );
});

test("decodes bounded PDF Flate rasters and reverses their row predictors", async () => {
  const bytes = pdfRaster(2, 1, 3, 2, Uint8Array.from([10, 20, 30, 30, 30, 30]));
  const output = await decodeOfficeImagePayload(codecInput(
    "pdf-raster",
    "image/x-officeviewer-pdf-raster",
    bytes,
  ));
  assert.equal(output.kind, "rgba");
  assert.equal(output.width, 2);
  assert.equal(output.height, 1);
  assert.deepEqual([...output.data], [10, 20, 30, 255, 40, 50, 60, 255]);
});

test("decodes compressed indexed PDF rasters without expanding them in the parser", async () => {
  const bytes = pdfRaster(2, 1, 1, 1, Uint8Array.from([0, 255]), {
    components: 3,
    high: 2,
    decode: [0, 255],
    bytes: Uint8Array.from([255, 0, 0, 0, 255, 0, 0, 0, 255]),
  });
  const output = await decodeOfficeImagePayload(codecInput(
    "pdf-raster",
    "image/x-officeviewer-pdf-raster",
    bytes,
  ));
  assert.deepEqual([...output.data], [255, 0, 0, 255, 0, 0, 255, 255]);
});

test("bounds highly compressed PDF rasters by their authenticated dimensions", async () => {
  const valid = pdfRaster(64, 64, 1, 1, new Uint8Array(4_096));
  const output = await decodeOfficeImagePayload(codecInput(
    "pdf-raster",
    "image/x-officeviewer-pdf-raster",
    valid,
    { maxBytes: 4_096, maxPixels: 4_096, maxCompressionRatio: 1 },
  ));
  assert.equal(output.kind, "rgba");
  assert.equal(output.data.length, 64 * 64 * 4);

  const falseDimensions = valid.slice();
  const view = new DataView(falseDimensions.buffer);
  view.setUint32(8, 1, false);
  view.setUint32(12, 1, false);
  await assert.rejects(
    decodeOfficeImagePayload(codecInput(
      "pdf-raster",
      "image/x-officeviewer-pdf-raster",
      falseDimensions,
      { maxBytes: 4_096, maxPixels: 4_096, maxCompressionRatio: 1 },
    )),
    (error) => error?.code === "IMAGE_DECOMPRESSION_LIMIT",
  );
});

test("decodes both boxed JP2 and raw J2K Office image parts", async () => {
  const jp2 = jpeg2000();
  for (const [format, mediaType, bytes] of [
    ["jp2", "image/jp2", jp2],
    ["j2k", "image/j2k", jp2Codestream(jp2)],
  ]) {
    const output = await decodeOfficeImagePayload(codecInput(format, mediaType, bytes));
    assert.equal(output.kind, "rgba");
    assert.equal(output.width, 2);
    assert.equal(output.height, 1);
    assert.deepEqual([...output.data], [252, 0, 0, 255, 252, 0, 0, 255]);
    assert.equal(output.approximate, format === "j2k");
  }
});

test("decodes oversized PDF JPX assets at their bounded reduction level", async () => {
  const output = await decodeOfficeImagePayload(codecInput(
    "jp2",
    "image/jp2",
    boxedJpxReduction(jpeg2000(), 1),
  ));
  assert.equal(output.kind, "rgba");
  assert.equal(output.width, 1);
  assert.equal(output.height, 1);
  assert.deepEqual([...output.data], [252, 0, 0, 255]);
  assert.equal(output.approximate, true);
});

test("normalizes JPEG 2000 tile coordinates against a non-zero image origin", async () => {
  const bytes = jpeg2000().slice();
  const siz = bytes.findIndex((value, index) => value === 0xff && bytes[index + 1] === 0x51);
  assert.ok(siz >= 0);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(siz + 6, 3, false);
  view.setUint32(siz + 14, 1, false);
  const output = await decodeOfficeImagePayload(codecInput("jp2", "image/jp2", bytes));
  assert.equal(output.kind, "rgba");
  assert.equal(output.width, 2);
  assert.equal(output.height, 1);
  assert.deepEqual([output.data[3], output.data[7]], [255, 255]);
});

test("rejects JP2 files whose ihdr dimensions disagree with the codestream SIZ", async () => {
  const bytes = jpeg2000().slice();
  const ihdr = new TextEncoder().encode("ihdr");
  const marker = bytes.findIndex((_, offset) => ihdr.every((value, index) => bytes[offset + index] === value));
  assert.ok(marker >= 0);
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint32(marker + 8, 99, false);
  await assert.rejects(
    decodeOfficeImagePayload(codecInput("jp2", "image/jp2", bytes)),
    (error) => error?.code === "IMAGE_HEADER_INVALID",
  );
});
