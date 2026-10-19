import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { Core, DEFAULT_LIMITS } from "../dist/core.js";

const coreWasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));

test("optimized core stays within the size budget", () => {
  // Shared binary references: 3,823,898 -> 3,827,003 bytes (+3,105).
  // The supplied DOCX's repeated 5.7 MB header previously exceeded the 256 MiB
  // transfer cap; its snapshot is now 8,704,908 bytes. The 47-page regression
  // and all real-document scene baselines below preserve the document content.
  // Page-local merged table cells add 1,381 bytes (3,827,079 -> 3,828,460).
  // The supplied table regression verifies continued borders on all three pages.
  // Reusing immutable resource preflight adds 2,266 bytes (-> 3,830,726).
  // Original DOCX first paint: Chromium 1339 -> 864 ms, WebKit 1132 -> 699 ms;
  // all 51 pages remain pixel-identical in both engines.
  // Shared exact byte search adds 7,194 bytes; real PDF performance/pixel proof
  // is recorded in docs/research/CONNECTED-CAR-PDF-PERFORMANCE.md.
  assert.ok(coreWasm.byteLength <= 3_839_000, `core Wasm is ${coreWasm.byteLength} bytes; budget is 3839000`);
});

test("calculator contains only its runtime ABI and stays within its size budget", async () => {
  const bytes = await readFile(new URL("../dist/office-viewer-calc.wasm", import.meta.url));
  const module = await WebAssembly.compile(bytes);
  assert.deepEqual(WebAssembly.Module.exports(module).map(({ name }) => name).sort(), [
    "memory", "ov_alloc", "ov_calc_abi_version", "ov_calculate", "ov_free",
    "ov_result_clear", "ov_result_pointer",
  ].sort());
  assert.equal(WebAssembly.Module.customSections(module, "__wasm_bindgen_unstable").length, 0);
  assert.ok(bytes.byteLength <= 2_040_000, `calculator Wasm is ${bytes.byteLength} bytes`);
});

test("PDF pack stays within the reviewed binary size budget", async () => {
  const bytes = await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url));
  // The same shared protocol adds 3,241 bytes: 3,141,589 -> 3,144,830.
  // Repeated images, masks and color profiles share exact bytes; fonts and
  // document rendering semantics remain unchanged.
  // Shared preflight reuse adds 2,181 bytes (3,144,830 -> 3,147,011).
  // Exact search and CMYK Decode tables: 3,147,815 -> 3,154,935 bytes.
  assert.ok(bytes.byteLength <= 3_156_000, `PDF Wasm is ${bytes.byteLength} bytes`);
});

test("binary Dingbats font preserves the original decoded font bytes", async () => {
  const bytes = await readFile(new URL("../third_party/pdf-standard-fonts/FoxitDingbats.cff", import.meta.url));
  assert.equal(bytes.byteLength, 29513);
  assert.equal(createHash("sha256").update(bytes).digest("hex"),
    "845c752392b6c914fb989c75a08b7792b88f542d2499042ef2889f8c814a16ed");
});

// Captured before callback type erasure. Include initial, incremental and full scenes,
// with a fixed field date/time; geometry, text, colors, media and diagnostics must match.
// Reviewed v0.2.71 changes: XLSX tab colors, ODT header/footer space, ODS optimal
// row heights, and ODP removal of the spurious terminal dash cap.
// Reviewed v0.2.73 scene changes are listed in OOXML-DISPLAY-GAP-AUDIT.md section 14:
// saved XLSX row metrics, chart labels, DOCX caps/borders/picture effects and SmartArt shadows.
// Reviewed v0.2.75 chart text/axes, DOCX cell-end markers, floating-object z order
// and VML gradient changes are recorded in the same audit, section 15.
// Reviewed v0.2.77 reading-color provenance, transparent unfilled cells and ODS
// default grid ownership are recorded in the same audit, section 16.
const scenes = {
  "core": {
    "chart-original.docx": "fff3cf84016fa0ed9d5c7ad28584e66a7333423aa77501d1c75bdf199e7db3f2",
    // Three numbering labels now share their paragraph's fixed leading policy;
    // page-local cells also remove a spurious merge-clipping warning caused by
    // float rounding in a nonmerged cell. All other fields match across 102 loads.
    "complex2005_12rtm.docx": "ad97735306afa0ca8fa145bd765bb4d5917ffa6795ee8250b7a1e3af25999345",
    "word-data-label-borders.docx": "5a7c82fdfb7bdbb33d656f763df9cb98e00d178c22a4264bf5cb49ea2cb2b3d3",
    "chart_pt_color_bg1.pptx": "c3a40615d11e2ad4057daa21265786dbdff5fd1c104862c090aa83fa7fd29597",
    "corpus-stacked-mix.pptx": "8078699c48e1abc96a8717f87e33602cbf2c44f15d41d73b75668317d58361a8",
    "funnel-pp1.pptx": "29df874d81433d53a150ca2f608ac75ef75596d7e1852ea53a8a5d6cc4983e8d",
    "chart-empty-title-original.xlsx": "a5c0e8891931d772f978da9e2080887f256d6f08b4119081daf54438c9c36530",
    "sunburst.xlsx": "e252ed9932529f328d7813643789e7ed8af8c4d9dc40b882fc21120fb93a514f",
    "color_funnel.xlsx": "e92466efaf2f65a46c026936ee1a33924c1e0a82061410052ac9cab5f76a0576",
    "smartart-chevron.pptx": "efad1c39b4b86fb13ae7efdac2a154dc543962daa62195eb34ee0e94b10b6497",
    "smartart-cycle.pptx": "b7d4b001d6a2f4896c5087dac5c66dca62aaf4dca9914cbe096ffb4fdcf220c9",
    "smartart-dir.pptx": "6ac256dce458e769b5f9bb0f236b807afda5df046fb5b26821c569113444df0f",
    "table-style-cascade.pptx": "c365b5b56d6d1824a41f5d2633417a6217bd74c300064c2af31a3421f353cee6",
    "customshape-bitmapfill-srcrect.pptx": "de55fdb1f0f400580a36541402750ca55ec480d0a2b37604a24663e310be9ccc",
    "background-effects.pptx": "e0d90358880538b8568c57a238eb38d11738ee6ca7296e0e0b583c35955f41d4",
    "background-theme-gradient.pptx": "5e423eb5a23aaf406d4fd55c2176e6c1be8fbd63c7a63456904f45dc2f6a2b18",
    "background-linear.docx": "2fff91a03563856c8f3f052b07437bafce3ee08b19b2d50f355a5ec10eec995f",
    "background-radial.docx": "838ab0d1938124b4bc39b4e646799605d8b6e6cdf8f35ba8c2f9b89329da585f",
    "background-worksheet.xlsx": "a7240cd9ba9cb42d0e0660997621237fe956a374466e17bae7e595fb53b7148b"
  },
  "odf": {
    "background-image.odt": "e813012a4ca332d64691cc5bfb3bae88e91af41ea26a6430c2db4ef73baf4f06",
    "toolkit-basic-sheet.ods": "f38cd9e0195531bcf5802a1fb1a6a503d7d80b9925e8158ef85aeb4d1bbcbee2",
    "odf-dash-linecaps.odp": "30000f3bfdacd475c2a84ad39cdb0292ffcc8ecbcd3545ff4ecc5854dd04b582"
  },
  "legacy-office": {
    "word-table-auto-height.doc": "d61cb045a6ed82ad2963667c205b742fdd7c6c308e63a0ca01ea11c610ccd3a3",
    "piechart_outside.xls": "cf2a4bc7e2c951173ac9c6e6dff103a3f285427fd82d0bfed99735f108ada17e",
    "ppt-render-values.ppt": "6b7bd9683536ecf3a74fb92b6268fb49194656a3ff7a8bbc1d21eebe6811eb54"
  },
  "pdf": {
    "word-original.pages": "9ec57779010c6bd3ee7f85f83c5594bacc8464ef281f333dde67e8013219d677",
    "keynote-first-line-indent.key": "d1b2ba39ae4a0356b99a0e043da10e03310aff094cf1ac8d2bf169b0dccf576e",
    "chart-original.numbers": "cfe0aa24ce73398bd6ba8c73316b6da3a768640f002713079d5b07984080a920"
  },
  "xps": {
    "ecma-388-cover.xps": "0b77e74e8931c0ed08a97a5587e488714c52d259e1b3d7c604a68f9553d628f6"
  },
  "ofd": {
    "ofdrw-intro.ofd": "9a1d5a0bacaf7315f52baf93ce3bcd3a7043619d4a8d16d58ad2ac42a3a9e372",
    "ofdrw-testImageNotFound.ofd": "8b4c73aa5b29267287f50ab29d24df9ed1560a688130d9977db7a10e0394a0e6"
  }
};

for (const [pack, fixtures] of Object.entries(scenes)) {
  const wasm = pack === "core" ? coreWasm
    : await readFile(new URL(`../dist/office-viewer-${pack}.wasm`, import.meta.url));
  for (const [file, expected] of Object.entries(fixtures)) {
    test(`shared parser size optimization preserves ${file}`, async () => {
      const core = await Core.create(wasm, DEFAULT_LIMITS);
      const doc = core.open(
        await readFile(new URL(`./fixtures/${file}`, import.meta.url)),
        undefined, undefined, [20260921, 1200],
      );
      try {
        const hash = createHash("sha256");
        // New alternate-name metadata has its own real-document regression;
        // retain this baseline for every pre-existing scene field.
        const hashScene = () => {
          const { fontAlternateNames, ...scene } = doc.scene;
          hash.update(JSON.stringify(scene));
        };
        hashScene();
        for (let index = 0; index < doc.scene.info.units.length; index += 1) {
          doc.loadUnit(index);
          hashScene();
        }
        doc.loadAll();
        assert.equal(doc.scene.fatal, false);
        assert.ok(doc.scene.objects.length > 0);
        hashScene();
        assert.equal(hash.digest("hex"), expected, "render scene changed from the reviewed baseline");
      } finally {
        doc.close();
        core.close();
      }
    });
  }
}
