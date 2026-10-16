import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

// Exercise the distributed MPL source, including its private stream boundary.
const modified = await readFile(new URL("../dist/third-party-licenses/mtx-decompressor-1.6.0-modified.mjs", import.meta.url), "utf8");
const upstream = await readFile(new URL("../node_modules/mtx-decompressor/dist/index.mjs", import.meta.url), "utf8");
function streamClass(source, Bytes = Uint8Array) {
  return vm.runInNewContext(source.replace(/export \{[^}]+\};\s*$/u, "") + "\nStream;", { Uint8Array: Bytes });
}

test("MTX repeated reservations allocate linear storage and preserve font bytes", () => {
  let allocated = 0;
  const Stream = streamClass(modified, class extends Uint8Array {
    constructor(...args) {
      super(...args);
      if (typeof args[0] === "number") allocated += this.byteLength;
    }
  });
  const stream = new Stream(null, 0);
  for (let i = 0; i < 4096; i++) {
    stream.reserve(i + 1);
    stream.writeU8(i & 255);
  }
  assert.ok(allocated < 4096 * 4, `Repeated reservations allocated ${allocated} bytes`);
  assert.deepEqual(Buffer.from(stream.toUint8Array()), Buffer.from(Array.from({ length: 4096 }, (_, i) => i & 255)));
});

test("MTX spare physical capacity does not relax original copy and seek bounds", () => {
  const exercise = source => {
    const Stream = streamClass(source);
    const stream = new Stream(null, 0);
    stream.reserve(4);
    stream.writeU32(0x01020304);
    stream.reserve(5); // Modified storage has spare bytes, logical reservation remains five.
    const input = new Stream(Uint8Array.of(5, 6), 2);
    const errors = [];
    for (const action of [() => input.copyTo(stream, 2), () => stream.seekAbsoluteThroughReserve(6)]) {
      try { action(); assert.fail("Out-of-reservation access must fail"); }
      catch (error) { errors.push({ code: error.code, message: error.message }); }
    }
    stream.reserve(6);
    input.copyTo(stream, 2);
    stream.seekAbsolute(1);
    stream.writeU8(9);
    return { errors, pos: stream.pos, size: stream.size, reserved: stream.reserved,
      bytes: Buffer.from(stream.toUint8Array()) };
  };
  const actual = exercise(modified);
  assert.deepEqual(actual, exercise(upstream));
  assert.deepEqual(actual.errors.map(error => error.code), ["OUT_OF_RESERVED_SPACE", "SEEK_PAST_EOS"]);
  assert.deepEqual(actual.bytes, Buffer.from([1, 9, 3, 4, 5, 6]));
});

function huffmanClasses(source) {
  return vm.runInNewContext(source.replace(/export \{[^}]+\};\s*$/u, "") + "\n({ AHuff, BitIO });");
}

test("MTX adaptive Huffman storage preserves symbols, weights and truncated-input position", () => {
  const original = huffmanClasses(upstream), optimized = huffmanClasses(modified);
  for (const range of [2, 3, 8, 32, 255, 256, 258, 300, 511, 512]) {
    let seed = range;
    const bytes = Uint8Array.from({ length: 4096 }, () => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return seed >>> 24;
    });
    const beforeIO = new original.BitIO(bytes), afterIO = new optimized.BitIO(bytes);
    const before = new original.AHuff(beforeIO, range), after = new optimized.AHuff(afterIO, range);
    // More than 32-bit weights must retain the original floating point semantics.
    before.tree[1].weight += 2 ** 32;
    after.tree.weight[1] += 2 ** 32;
    for (;;) {
      let expected;
      try { expected = before.readSymbol(); }
      catch (error) { assert.throws(() => after.readSymbol(), { message: error.message }); break; }
      assert.equal(after.readSymbol(), expected, `range ${range}`);
    }
    assert.deepEqual([afterIO.index, afterIO.bitCount, afterIO.bitBuffer],
      [beforeIO.index, beforeIO.bitCount, beforeIO.bitBuffer]);
    for (const field of ["up", "left", "right", "code", "weight"]) {
      assert.deepEqual(Array.from(after.tree[field]), Array.from(before.tree, node => node[field]), field);
    }
    assert.deepEqual(Array.from(after.symbolIndex), Array.from(before.symbolIndex));
  }
});
