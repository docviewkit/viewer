import { copyFile, mkdir, rm, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { build } from "esbuild";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const dist = resolve(root, "dist");
const licenses = resolve(dist, "third-party-licenses");

await mkdir(dist, { recursive: true });
await mkdir(licenses, { recursive: true });
await Promise.all([
  rm(resolve(dist, "noto-sans-hans.otf"), { force: true }),
  rm(resolve(licenses, "noto-sans-hans-OFL-1.1.txt"), { force: true }),
  rm(resolve(dist, "jbig2.wasm"), { force: true }),
  rm(resolve(dist, "jbig2_nowasm_fallback.js"), { force: true }),
  rm(resolve(dist, "openjpeg.wasm"), { force: true }),
  rm(resolve(dist, "jxr_dec.wasm"), { force: true }),
  rm(resolve(licenses, "pdfjs-dist-Apache-2.0.txt"), { force: true }),
  rm(resolve(licenses, "pdfjs-jbig2-LICENSE.txt"), { force: true }),
  rm(resolve(licenses, "pdfjs-jbig2-wrapper-LICENSE.txt"), { force: true }),
  rm(resolve(licenses, "pdfjs-dist@5.7.284-LICENSE"), { force: true }),
]);
const result = await build({
  entryPoints: [resolve(root, "src/image-codec-worker.ts")],
  outfile: resolve(dist, "image-codec-worker.js"),
  bundle: true,
  format: "esm",
  platform: "browser",
  target: "es2022",
  sourcemap: false,
  legalComments: "linked",
  define: {
    "process.env.NODE_ENV": '"production"',
  },
  metafile: true,
});
await writeFile(resolve(dist, ".image-codec-metafile.json"), `${JSON.stringify(result.metafile)}\n`);
await copyFile(
  resolve(root, "node_modules/@resvg/resvg-wasm/index_bg.wasm"),
  resolve(dist, "resvg.wasm"),
);
await copyFile(
  resolve(root, "node_modules/@discourse/jxr/codec/dec/jxr_dec.wasm"),
  resolve(dist, "jxr_dec.wasm"),
);
await copyFile(
  resolve(root, "third_party/openjpeg/openjpeg.wasm"),
  resolve(dist, "openjpeg.wasm"),
);
for (const [source, target] of [
  ["node_modules/emf-converter/LICENSE", "emf-converter-Apache-2.0.txt"],
  ["third_party/emf-converter/LICENSE", "apache-poi-preset-shapes-Apache-2.0.txt"],
  ["third_party/openjpeg/LICENSE_OPENJPEG", "openjpeg-BSD-2-Clause.txt"],
  ["third_party/openjpeg/LICENSE_PDFJS_OPENJPEG", "pdfjs-openjpeg-BSD-2-Clause.txt"],
  ["node_modules/utif2/LICENSE", "utif2-MIT.txt"],
  ["node_modules/@discourse/jxr/LICENSE", "discourse-jxr-Apache-2.0.txt"],
  ["node_modules/pako/LICENSE", "pako-MIT.txt"],
  ["node_modules/pako/lib/zlib/README", "pako-zlib.txt"],
  ["third_party/licenses/MPL-2.0.txt", "resvg-MPL-2.0.txt"],
]) {
  await copyFile(resolve(root, source), resolve(licenses, target));
}
