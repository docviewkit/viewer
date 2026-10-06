import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { Core, DEFAULT_LIMITS } from "../dist/core.js";

// Office/OpenXML SDK and LibreOffice/OASIS files; PPTX reference variants and
// the XLSX picture relationship are derived from their real authored documents.
for (const [file, expected] of [
  ["background-theme-gradient.pptx", "linear-gradient"],
  ["background-rgb-reference.pptx", "solid"],
  ["background-effects.pptx", "effects"],
  ["background-worksheet.xlsx", "image"],
  ["background-linear.docx", "linear-gradient"],
  ["background-radial.docx", "shape-gradient"],
  ["background-stops.docx", "linear-gradient"],
  ["background-texture.docx", "image"],
  ["background-picture.docx", "image"],
  ["background-color.odt", "solid"],
  ["background-image.odt", "image"],
]) {
  test(`background regression: ${file}`, async () => {
    const wasm = file.endsWith(".odt") ? "office-viewer-odf.wasm" : "office-viewer-core.wasm";
    const core = await Core.create(await readFile(new URL(`../dist/${wasm}`, import.meta.url)), DEFAULT_LIMITS);
    const doc = core.open(await readFile(new URL(`./fixtures/${file}`, import.meta.url)));
    try {
      doc.loadUnit(0);
      const unit = doc.scene.info.units[0];
      const background = doc.scene.objects.find(o => o.unitIndex === 0 && o.bounds.x === 0 && o.bounds.y === 0
        && Math.abs(o.bounds.width - unit.width) < .1 && Math.abs(o.bounds.height - unit.height) < .1);
      assert.ok(background, "full-page background must exist");
      let visual = background.visual;
      if (expected === "effects") assert.equal(visual.kind, "advanced-effect");
      while (visual.visual) visual = visual.visual;
      if (expected !== "effects") assert.equal(visual.fill?.kind, expected);
      if (file === "background-rgb-reference.pptx") assert.equal(visual.fill.color, 0x00cc99ff);
      if (file === "background-color.odt") assert.equal(visual.fill.color, 0xeee8aaff);
      if (file === "background-stops.docx") assert.ok(visual.fill.stops.length >= 4);
      if (["background-worksheet.xlsx", "background-texture.docx", "background-image.odt"].includes(file)) {
        assert.equal(visual.fill.tile, true);
      }
      if (!["background-picture.docx", "background-image.odt"].includes(file)) assert.ok(doc.scene.objects.some(o => o.text?.trim()), "body content must survive background parsing");
    } finally { doc.close(); core.close(); }
  });
}
