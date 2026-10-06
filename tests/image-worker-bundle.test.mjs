import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { inflateSync } from "node:zlib";

function pngPixels(bytes) {
  assert.deepEqual([...bytes.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const idat = [];
  let width;
  let height;
  let offset = 8;
  while (offset + 12 <= bytes.length) {
    const length = view.getUint32(offset, false);
    const type = new TextDecoder().decode(bytes.subarray(offset + 4, offset + 8));
    const data = bytes.subarray(offset + 8, offset + 8 + length);
    if (type === "IHDR") {
      width = view.getUint32(offset + 8, false);
      height = view.getUint32(offset + 12, false);
      assert.equal(data[8], 8, "test decoder expects 8-bit PNG output");
      assert.equal(data[9], 6, "test decoder expects RGBA PNG output");
    } else if (type === "IDAT") {
      idat.push(data);
    } else if (type === "IEND") {
      break;
    }
    offset += 12 + length;
  }
  assert.ok(width > 0 && height > 0);
  const packed = new Uint8Array(idat.reduce((total, chunk) => total + chunk.length, 0));
  let packedOffset = 0;
  for (const chunk of idat) {
    packed.set(chunk, packedOffset);
    packedOffset += chunk.length;
  }
  const inflated = new Uint8Array(inflateSync(packed));
  const stride = width * 4;
  const pixels = new Uint8Array(stride * height);
  const paeth = (left, above, upperLeft) => {
    const estimate = left + above - upperLeft;
    const leftDistance = Math.abs(estimate - left);
    const aboveDistance = Math.abs(estimate - above);
    const upperLeftDistance = Math.abs(estimate - upperLeft);
    return leftDistance <= aboveDistance && leftDistance <= upperLeftDistance
      ? left
      : aboveDistance <= upperLeftDistance ? above : upperLeft;
  };
  for (let y = 0; y < height; y += 1) {
    const source = y * (stride + 1);
    const filter = inflated[source];
    for (let x = 0; x < stride; x += 1) {
      const raw = inflated[source + 1 + x];
      const target = y * stride + x;
      const left = x >= 4 ? pixels[target - 4] : 0;
      const above = y > 0 ? pixels[target - stride] : 0;
      const upperLeft = y > 0 && x >= 4 ? pixels[target - stride - 4] : 0;
      const predictor = filter === 0 ? 0
        : filter === 1 ? left
          : filter === 2 ? above
            : filter === 3 ? Math.floor((left + above) / 2)
              : filter === 4 ? paeth(left, above, upperLeft) : -1;
      assert.notEqual(predictor, -1, `unsupported PNG filter ${filter}`);
      pixels[target] = (raw + predictor) & 0xff;
    }
  }
  return { width, height, pixels };
}

function restore(name, descriptor) {
  if (descriptor === undefined) delete globalThis[name];
  else Object.defineProperty(globalThis, name, descriptor);
}

test("production image Worker bundle renders SVG text and blocks relative resources", async () => {
  const descriptors = Object.fromEntries(
    ["onmessage", "postMessage", "createImageBitmap", "fetch"].map((name) => [
      name,
      Object.getOwnPropertyDescriptor(globalThis, name),
    ]),
  );
  let encodedPng;
  let resolveResponse;
  const fontBytes = await readFile(new URL(
    "../node_modules/@embedpdf/fonts-sc/fonts/NotoSansHans-Regular.otf",
    import.meta.url,
  ));
  try {
    globalThis.fetch = async (input) => {
      const url = input instanceof Request ? input.url : String(input);
      if (!url.startsWith("file:")) throw new Error(`unexpected Worker fetch: ${url}`);
      return new Response(await readFile(new URL(url)), {
        status: 200,
        headers: { "Content-Type": url.endsWith(".wasm") ? "application/wasm" : "font/otf" },
      });
    };
    globalThis.createImageBitmap = async (source) => {
      assert.ok(source instanceof Blob);
      encodedPng = new Uint8Array(await source.arrayBuffer());
      const decoded = pngPixels(encodedPng);
      return { width: decoded.width, height: decoded.height, close() {} };
    };
    globalThis.postMessage = (message) => resolveResponse(message);
    await import(`../dist/image-codec-worker.js?bundle-test=${Date.now()}`);

    const request = async (source) => {
      const bytes = new TextEncoder().encode(source);
      const response = new Promise((resolve, reject) => {
        resolveResponse = resolve;
        setTimeout(() => reject(new Error("image Worker test timed out")), 3_000).unref();
      });
      globalThis.onmessage({ data: {
        id: 1,
        format: "svg",
        mediaType: "image/svg+xml",
        bytes: bytes.buffer,
        maxPixels: 1_000_000,
        maxBytes: 1_000_000,
        maxCompressionRatio: 200,
        fonts: [{ family: "Noto Sans Hans", bytes: fontBytes.buffer.slice(fontBytes.byteOffset, fontBytes.byteOffset + fontBytes.byteLength) }],
      } });
      return response;
    };

    const rendered = await request(
      '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="30"><text x="1" y="22" font-size="20">Office中文</text></svg>',
    );
    assert.equal(rendered.ok, true);
    assert.equal(rendered.approximate, true);
    const decoded = pngPixels(encodedPng);
    assert.ok(decoded.pixels.some((value, index) => index % 4 === 3 && value !== 0), "SVG text must produce visible pixels");

    const blocked = await request(
      '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><image href="relative.png" width="1" height="1"/></svg>',
    );
    assert.deepEqual(
      { ok: blocked.ok, code: blocked.code },
      { ok: false, code: "IMAGE_EXTERNAL_RESOURCE_BLOCKED" },
    );

    const embedded = await request(
      '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><image width="1" height="1" href="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="/></svg>',
    );
    assert.equal(embedded.ok, true);

    let nested = '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>';
    for (let depth = 0; depth < 10; depth += 1) {
      nested = `<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><image href="data:image/svg+xml;base64,${Buffer.from(nested).toString("base64")}"/></svg>`;
    }
    const tooDeep = await request(nested);
    assert.deepEqual(
      { ok: tooDeep.ok, code: tooDeep.code },
      { ok: false, code: "IMAGE_SVG_RECURSION_LIMIT" },
    );
  } finally {
    for (const [name, descriptor] of Object.entries(descriptors)) restore(name, descriptor);
  }
});
