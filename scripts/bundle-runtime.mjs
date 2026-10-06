import { copyFile, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

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
