import { copyFile, mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { createHash } from "node:crypto";

import { build } from "esbuild";

const root = resolve(import.meta.dirname, "..");
const dist = resolve(root, "dist");
const manifest = JSON.parse(await readFile(resolve(root, "package.json"), "utf8"));

for (const entry of await readdir(dist)) {
  if (/^runtime-[A-Z0-9]+\.js$/u.test(entry)) await rm(resolve(dist, entry), { force: true });
}

await build({
  entryPoints: {
    index: resolve(root, "src/index.ts"),
    accuracy: resolve(root, "src/accuracy.ts"),
    viewer: resolve(root, "src/viewer.ts"),
    worker: resolve(root, "src/worker.ts"),
    "mtx-decompressor": resolve(root, "src/mtx-decompressor.ts"),
  },
  outdir: dist,
  bundle: true,
  splitting: true,
  chunkNames: "runtime-[hash]",
  format: "esm",
  platform: "browser",
  target: "es2022",
  minify: true,
  sourcemap: false,
  legalComments: "linked",
  plugins: [{
    name: "mtx-buffer-growth",
    setup(builder) {
      builder.onLoad({ filter: /mtx-decompressor[/\\]dist[/\\]index\.mjs$/ }, async ({ path }) => {
        const source = await readFile(path, "utf8");
        if (createHash("sha256").update(source).digest("hex") !==
          "12bfb0a420bea3b45a8c0ceac48fce0557b9f59313fb86d11d81cdf9978db4af") {
          throw new Error("mtx-decompressor 1.6.0 source changed; revalidate the performance patches");
        }
        const original = [
          "    const newBuf = new Uint8Array(n);",
          "    newBuf.set(this.buf.subarray(0, this.size));",
          "    this.buf = newBuf;",
        ].join("\n");
        if (source.split(original).length !== 2) {
          throw new Error("mtx-decompressor glyph allocator changed; revalidate the pinned performance patch");
        }
        // 1.6.0 copies the growing glyf table once per glyph. Grow physical
        // storage geometrically, but keep the original logical `reserved = n`
        // assignment so copyTo/seekThroughReserve reject the same invalid input.
        let contents = source.replace(original, [
          "    if (n > this.buf.length) {",
          "      const newBuf = new Uint8Array(Math.max(n, this.buf.length * 2));",
          "      newBuf.set(this.buf.subarray(0, this.size));",
          "      this.buf = newBuf;",
          "    }",
        ].join("\n"));
        // Preserve the adaptive Huffman algorithm while storing each node field
        // contiguously. Float64 weights retain JavaScript's original number range.
        const start = contents.indexOf("var AHuff =");
        const end = contents.indexOf("// src/bitio.ts", start);
        let huffman = contents.slice(start, end).replace(
          `this.tree = Array.from({ length: treeSize });
    for (let i = 0; i < treeSize; i++) {
      this.tree[i] = { up: 0, left: 0, right: 0, code: -1, weight: 0 };
    }`,
          `this.tree = {
      up: new Int32Array(treeSize), left: new Int32Array(treeSize),
      right: new Int32Array(treeSize), code: new Int32Array(treeSize).fill(-1),
      weight: new Float64Array(treeSize),
    };`,
        ).replace("this.symbolIndex = Array.from({ length: range });", "this.symbolIndex = new Int32Array(range);");
        const swapStart = huffman.indexOf("    const upa = tree[a].up;");
        const swapEnd = huffman.indexOf("    let code = tree[a].code;", swapStart);
        huffman = huffman.slice(0, swapStart) + [
          "    let value = tree.left[a]; tree.left[a] = tree.left[b]; tree.left[b] = value;",
          "    value = tree.right[a]; tree.right[a] = tree.right[b]; tree.right[b] = value;",
          "    value = tree.code[a]; tree.code[a] = tree.code[b]; tree.code[b] = value;",
          "    value = tree.weight[a]; tree.weight[a] = tree.weight[b]; tree.weight[b] = value;",
          "",
        ].join("\n") + huffman.slice(swapEnd);
        huffman = huffman.replace(`    const node = this.tree[a];
    if (node.code >= 0) {
      return node.weight;
    }
    node.weight = this.initWeight(node.left) + this.initWeight(node.right);
    return node.weight;`, `    const tree = this.tree;
    if (tree.code[a] >= 0) return tree.weight[a];
    tree.weight[a] = this.initWeight(tree.left[a]) + this.initWeight(tree.right[a]);
    return tree.weight[a];`);
        huffman = huffman.replace(/((?:this\.)?tree)\[([^\[\]]+)\]\.(up|left|right|code|weight)/gu, "$1.$3[$2]")
          .replace(/tree\[tree\.(left|right)\[([ab])\]\]\.up/gu, "tree.up[tree.$1[$2]]");
        huffman = huffman.replace(`    let symbol;
    do {
      a = this.bio.inputBit() ? this.tree.right[a] : this.tree.left[a];
      symbol = this.tree.code[a];
    } while (symbol < 0);`, `    let symbol;
    const bio = this.bio, tree = this.tree;
    let bitBuffer = bio.bitBuffer, bitCount = bio.bitCount, index = bio.index;
    do {
      if (bitCount === 0) {
        if (index >= bio.size) {
          bio.bitBuffer = bitBuffer; bio.bitCount = bitCount; bio.index = index;
          throw new Error("BitIO: end of data");
        }
        bitBuffer = bio.data[index++];
        bitCount = 8;
      }
      bitCount--;
      bitBuffer <<= 1;
      a = (bitBuffer & 256) !== 0 ? tree.right[a] : tree.left[a];
      symbol = tree.code[a];
    } while (symbol < 0);
    bio.bitBuffer = bitBuffer; bio.bitCount = bitCount; bio.index = index;`);
        contents = contents.slice(0, start) + huffman + contents.slice(end);
        const licenses = resolve(dist, "third-party-licenses");
        await mkdir(licenses, { recursive: true });
        await writeFile(resolve(licenses, "mtx-decompressor-1.6.0-modified.mjs"),
          "// MPL-2.0. mtx-decompressor 1.6.0; OfficeViewer modifies glyph buffer growth and adaptive Huffman storage/bit reads.\n" + contents);
        return { contents, loader: "js" };
      });
    },
  }],
});

for (const entry of await readdir(dist)) {
  if (entry.endsWith(".map")) await rm(resolve(dist, entry), { force: true });
}
await writeFile(
  resolve(dist, "version.json"),
  `${JSON.stringify({ name: "@docviewkit/viewer", version: manifest.version }, null, 2)}\n`,
);
for (const entry of ["LICENSE", "NOTICE"]) {
  await copyFile(resolve(root, entry), resolve(dist, entry));
}
await copyFile(
  resolve(root, "node_modules/@kittl/little-cms/dist/lcms.wasm"),
  resolve(dist, "lcms.wasm"),
);

for (const family of ["Caladea", "Carlito"]) {
  for (const variant of ["Regular", "Bold", "Italic", "BoldItalic"]) {
    await copyFile(resolve(root, `third_party/${family.toLowerCase()}/${family}-${variant}.ttf`),
      resolve(dist, `${family}-${variant}.ttf`));
  }
}
