import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

import { createOfficeEngine as createOfficeEngineBase } from "../dist/engine.js";
import { extendedFormatPack } from "../dist/extended-formats.js";
import {
  approximateFontFamily,
  collectFontRequests,
  compatibleFallbackFamily,
  documentFontRuns,
  dominantDocumentFontSize,
  fallbackBaselineShift,
  fallbackFontSizeFactor,
  fallbackSpaceAdvanceEm,
  FontProviderCache,
  fontShorthand,
  prepareEmbeddedFontAssets,
  prepareFontAssets,
  probeBrowserFontFace,
  registerFonts,
  semanticSymbolFontText,
} from "../dist/font.js";
import { Core, DEFAULT_LIMITS } from "../dist/core.js";
import { SceneRenderer } from "../dist/render.js";
import { createZip } from "./zip-fixture.mjs";

const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
const odfWasm = await readFile(new URL("../dist/office-viewer-odf.wasm", import.meta.url));

test("dominant text size aggregates character coverage and ignores invalid runs", () => {
  assert.equal(dominantDocumentFontSize([]), undefined);
  assert.equal(dominantDocumentFontSize([
    { start: 0, end: 4, fontSize: 32 },
    { start: 4, end: 7, fontSize: 16 },
    { start: 7, end: 10, fontSize: 16 },
    { start: 10, end: 100, fontSize: NaN },
    { start: 10, end: 100, fontSize: 0 },
    { start: 10, end: 100 },
    { start: 100, end: 10, fontSize: 64 },
  ]), 16);
});

function createOfficeEngine(options) {
  return createOfficeEngineBase({
    formatPack: async () => ({
      async load(candidate) {
        return readFile(await extendedFormatPack.load(candidate));
      },
    }),
    ...options,
  });
}

const odt = createZip({
  mimetype: "application/vnd.oasis.opendocument.text",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
      <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
      <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
    </manifest:manifest>`,
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
      <office:automatic-styles>
        <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin="1in"/></style:page-layout>
      </office:automatic-styles>
      <office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>
    </office:document-styles>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
      <office:body><office:text><text:p>Font lifecycle</text:p></office:text></office:body>
    </office:document-content>`,
}, { compress: false });

const embeddedFontKey = Uint8Array.of(
  0xae, 0x1e, 0x8e, 0x94, 0xa0, 0x18, 0xec, 0x90,
  0xd5, 0x4a, 0x60, 0xaa, 0xdc, 0x70, 0x1b, 0x00,
);
const embeddedFontBytes = Uint8Array.from({ length: 64 }, (_, index) => index);
const obfuscatedFontBytes = embeddedFontBytes.slice();
for (let index = 0; index < 32; index += 1) {
  obfuscatedFontBytes[index] ^= embeddedFontKey[index % embeddedFontKey.length];
}

function uncompressedEot(fontBytes) {
  const headerSize = 96;
  const bytes = new Uint8Array(headerSize + fontBytes.length);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, bytes.length, true);
  view.setUint32(4, fontBytes.length, true);
  view.setUint32(8, 0x00010000, true);
  view.setUint32(28, 400, true);
  view.setUint16(32, 0, true);
  view.setUint16(34, 0x504c, true);
  bytes.set(fontBytes, headerSize);
  return bytes;
}

function sfntWithNonzeroFormat12Language() {
  const bytes = new Uint8Array(140);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, 2);
  view.setUint32(12, 0x636d6170);
  view.setUint32(20, 44);
  view.setUint32(24, 40);
  view.setUint32(28, 0x68656164);
  view.setUint32(36, 84);
  view.setUint32(40, 54);
  view.setUint16(44, 0);
  view.setUint16(46, 1);
  view.setUint16(48, 3);
  view.setUint16(50, 10);
  view.setUint32(52, 12);
  view.setUint16(56, 12);
  view.setUint16(58, 0);
  view.setUint32(60, 28);
  view.setUint32(64, 634_208);
  view.setUint32(68, 1);
  view.setUint32(72, 0x4e00);
  view.setUint32(76, 0x4e00);
  view.setUint32(80, 1);
  return bytes;
}

test("uses compatible families for unresolved RTF and ODT authoring fonts", () => {
  assert.equal(compatibleFallbackFamily("odt", "DejaVu Sans"), "Verdana");
  assert.equal(compatibleFallbackFamily("odt", "Liberation Sans"), "Arial");
  assert.equal(compatibleFallbackFamily("rtf", "Liberation Sans"), "Arial");
  assert.equal(compatibleFallbackFamily("rtf", "DejaVu Sans"), "Verdana");
  assert.equal(compatibleFallbackFamily("docx", "DejaVu Sans"), undefined);
  assert.equal(compatibleFallbackFamily("rtf", "Open Sans"), undefined);
});

test("preserves legacy ShuSong fallback space metrics without widening other formats", () => {
  assert.equal(compatibleFallbackFamily("doc", "汉仪书宋二KW"), "Hiragino Mincho ProN");
  assert.equal(fallbackSpaceAdvanceEm("doc", "汉仪书宋二KW"), 0.5);
  assert.equal(fallbackSpaceAdvanceEm("docx", "汉仪书宋二KW"), undefined);
  assert.equal(fallbackSpaceAdvanceEm("doc", "宋体"), undefined);
});

test("extracts browser-compatible font data from PPTX EOT containers", () => {
  const [font] = prepareEmbeddedFontAssets([{
    family: "Presentation Embed",
    bytes: uncompressedEot(embeddedFontBytes),
    style: "normal",
    weight: 400,
  }]);
  assert.deepEqual([...new Uint8Array(font.bytes)], [...embeddedFontBytes]);
});

test("normalizes nonzero EOT cmap language values rejected by browser font sanitizers", () => {
  const [font] = prepareEmbeddedFontAssets([{
    family: "Presentation CJK Embed",
    bytes: uncompressedEot(sfntWithNonzeroFormat12Language()),
    style: "normal",
    weight: 400,
  }]);
  assert.equal(new DataView(font.bytes).getUint32(64), 0);
});

test("real OFD embedded fonts repair the lowercase OS/2 tag without changing glyphs", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-ofd.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/ofdrw-ano.ofd", import.meta.url)));
  const tables = (bytes) => {
    const buffer = Buffer.from(bytes);
    return Array.from({ length: buffer.readUInt16BE(4) }, (_, index) => {
      const record = 12 + index * 16;
      const offset = buffer.readUInt32BE(record + 8);
      return [buffer.toString("ascii", record, record + 4), buffer.subarray(offset, offset + buffer.readUInt32BE(record + 12))];
    });
  };
  try {
    const sourceFonts = document.scene.embeddedFonts;
    const prepared = prepareEmbeddedFontAssets(sourceFonts);
    let repaired = 0;
    for (let index = 0; index < sourceFonts.length; index += 1) {
      const before = new Map(tables(sourceFonts[index].bytes));
      const after = new Map(tables(prepared[index].bytes));
      if (!before.has("os/2")) continue;
      repaired += 1;
      assert.ok(after.has("OS/2"), `${sourceFonts[index].family}: browser requires OS/2`);
      assert.ok(!after.has("os/2"));
      assert.deepEqual([...after.keys()], [...after.keys()].sort());
      for (const tag of ["glyf", "loca", "hmtx"]) assert.deepEqual(after.get(tag), before.get(tag), tag);
      assert.deepEqual(after.get("OS/2"), before.get("os/2"));
      const expectedCmap = Buffer.from(before.get("cmap"));
      const subtable = expectedCmap.readUInt32BE(8);
      assert.equal(expectedCmap.readUInt16BE(subtable), 4);
      const segments = expectedCmap.readUInt16BE(subtable + 6) / 2;
      const sentinelDelta = subtable + 16 + segments * 4 + (segments - 1) * 2;
      assert.equal(expectedCmap.readUInt16BE(sentinelDelta), 0);
      expectedCmap.writeUInt16BE(1, sentinelDelta);
      assert.deepEqual(after.get("cmap"), expectedCmap, "only fix the invalid U+FFFF terminator");
      const bytes = Buffer.from(prepared[index].bytes);
      let checksum = 0;
      for (let offset = 0; offset < bytes.length; offset += 4) {
        const word = Buffer.alloc(4);
        bytes.copy(word, 0, offset, Math.min(offset + 4, bytes.length));
        checksum = (checksum + word.readUInt32BE()) >>> 0;
      }
      assert.equal(checksum, 0xb1b0afba);
      assert.ok(new Map(tables(sourceFonts[index].bytes)).has("os/2"), "do not mutate source font bytes");
    }
    assert.equal(repaired, 4);
  } finally { document.close(); core.close(); }
});

test("real Foxit OFD symbol subset drops stale post names without changing its bullet glyph", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-ofd.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/foxit-directory-first.ofd", import.meta.url)));
  const tables = (bytes) => {
    const buffer = Buffer.from(bytes);
    return new Map(Array.from({ length: buffer.readUInt16BE(4) }, (_, index) => {
      const record = 12 + index * 16;
      const offset = buffer.readUInt32BE(record + 8);
      return [buffer.toString("ascii", record, record + 4), buffer.subarray(offset, offset + buffer.readUInt32BE(record + 12))];
    }));
  };
  try {
    const source = document.scene.embeddedFonts.find(({ family }) => family === "Wingdings");
    assert.ok(source);
    const original = Buffer.from(source.bytes);
    const before = tables(original);
    assert.equal(before.get("maxp").readUInt16BE(4), 6);
    assert.equal(before.get("post").readUInt16BE(32), 226);
    const after = tables(prepareEmbeddedFontAssets([source])[0].bytes);
    assert.equal(after.get("post").readUInt32BE(0), 0x00030000);
    assert.equal(after.get("post").length, 32);
    assert.deepEqual([...after.keys()], [...after.keys()].sort());
    for (const [tag, bytes] of before) {
      if (tag !== "post" && tag !== "head") assert.deepEqual(after.get(tag), bytes, tag);
    }
    assert.deepEqual(source.bytes, new Uint8Array(original));
  } finally { document.close(); core.close(); }
});

test("normalizes nonzero raw SFNT cmap language values used by XPS fonts", () => {
  const [font] = prepareEmbeddedFontAssets([{
    family: "XPS CJK Embed",
    bytes: sfntWithNonzeroFormat12Language(),
    style: "normal",
    weight: 400,
  }]);
  assert.equal(new DataView(font.bytes).getUint32(64), 0);
});

test("uses weight-aware CJK fallbacks for unresolved Office presentation DengXian faces", () => {
  assert.equal(compatibleFallbackFamily("ppt", "等线 Light"), "PingFang SC Light");
  assert.equal(compatibleFallbackFamily("odp", "等线 Light"), "PingFang SC Light");
  assert.equal(compatibleFallbackFamily("ppt", "DengXian"), "PingFang SC Light");
  assert.equal(compatibleFallbackFamily("pptx", "等线 Light"), undefined);
  assert.equal(fallbackFontSizeFactor("ppt", "等线 Light"), 0.955);
  assert.equal(fallbackFontSizeFactor("odp", "等线 Light"), 0.955);
  assert.equal(fallbackFontSizeFactor("ppt", "等线"), undefined);
  assert.equal(fallbackFontSizeFactor("pptx", "等线 Light"), undefined);
  assert.ok(Math.abs(fallbackBaselineShift("ppt", "等线 Light", 80) - 7.2) < 1e-9);
  assert.ok(Math.abs(fallbackBaselineShift("odp", "等线 Light", 80) - 7.2) < 1e-9);
  assert.equal(fallbackBaselineShift("ppt", "等线", 80), undefined);
  assert.equal(
    fontShorthand("PingFang SC Light", 48, false, false),
    '300 48px "PingFang SC", "Microsoft YaHei UI", "Microsoft YaHei", "Hiragino Sans GB", sans-serif',
  );
});

const docxWithEmbeddedFont = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
      <Override PartName="/word/fonts/embedded.odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
  "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Embedded Sans"/></w:rPr><w:t>Embedded</w:t></w:r></w:p></w:body>
    </w:document>`,
  "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdFonts" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
    </Relationships>`,
  "word/fontTable.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:font w:name="Embedded Sans"><w:embedRegular r:id="rIdEmbedded" w:fontKey="{001B70DC-AA60-4AD5-90EC-18A0948E1EAE}"/></w:font>
    </w:fonts>`,
  "word/_rels/fontTable.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/embedded.odttf"/>
    </Relationships>`,
  "word/fonts/embedded.odttf": obfuscatedFontBytes,
}, { compress: false });

const pptxWithEmbeddedFont = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
      <Override PartName="/ppt/fonts/font1.fntdata" ContentType="application/x-fontdata"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    </Relationships>`,
  "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
      <p:sldSz cx="9144000" cy="6858000"/>
      <p:embeddedFontLst><p:embeddedFont><p:font typeface="Presentation Embed"/><p:regular r:id="rIdFont"/></p:embeddedFont></p:embeddedFontLst>
    </p:presentation>`,
  "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
      <Relationship Id="rIdFont" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.fntdata"/>
    </Relationships>`,
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:latin typeface="Presentation Embed"/></a:rPr><a:t>Embedded</a:t></a:r></a:p></p:txBody>
      </p:sp></p:spTree></p:cSld>
    </p:sld>`,
  "ppt/fonts/font1.fntdata": embeddedFontBytes,
}, { compress: false });

const odtWithEmbeddedFont = createZip({
  mimetype: "application/vnd.oasis.opendocument.text",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
      <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
    </manifest:manifest>`,
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
      <office:font-face-decls><style:font-face style:name="OdfEmbed" svg:font-family="Odf Embed"><svg:font-face-src><svg:font-face-uri xlink:href="Fonts/odf.ttf"/></svg:font-face-src></style:font-face></office:font-face-decls>
      <office:styles><style:style style:name="Body" style:family="paragraph"><style:text-properties style:font-name="OdfEmbed"/></style:style></office:styles>
      <office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin="1in"/></style:page-layout></office:automatic-styles>
      <office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>
    </office:document-styles>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
      <office:body><office:text><text:p text:style-name="Body">Embedded</text:p></office:text></office:body>
    </office:document-content>`,
  "Fonts/odf.ttf": embeddedFontBytes,
}, { compress: false });

class SliceTrap extends ArrayBuffer {
  slice() {
    throw new Error("font bytes were copied before the total limit was checked");
  }
}

class FakeFontSet {
  constructor() {
    this.faces = new Set();
    this.deleted = [];
  }

  add(face) {
    this.faces.add(face);
  }

  delete(face) {
    this.deleted.push(face);
    return this.faces.delete(face);
  }

  check() {
    return true;
  }
}

class LoadedFontFace {
  constructor(family, source, descriptors) {
    this.family = family;
    this.source = source;
    this.descriptors = descriptors;
  }

  async load() {
    if (typeof this.source === "string") throw new Error("local font unavailable");
    return this;
  }
}

function installFontRuntime(fontSet, FontFace = LoadedFontFace) {
  const fontFaceDescriptor = Object.getOwnPropertyDescriptor(globalThis, "FontFace");
  const fontsDescriptor = Object.getOwnPropertyDescriptor(globalThis, "fonts");
  Object.defineProperty(globalThis, "FontFace", { configurable: true, value: FontFace });
  Object.defineProperty(globalThis, "fonts", { configurable: true, value: fontSet });
  return () => {
    if (fontFaceDescriptor) Object.defineProperty(globalThis, "FontFace", fontFaceDescriptor);
    else delete globalThis.FontFace;
    if (fontsDescriptor) Object.defineProperty(globalThis, "fonts", fontsDescriptor);
    else delete globalThis.fonts;
  };
}

test("complex0 preserves declared font alternate names through the Core snapshot", async () => {
  const core = await Core.create(wasm, DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/complex0.docx", import.meta.url)));
  try {
    assert.deepEqual(document.scene.fontAlternateNames, [
      { family: "SimSun", names: ["宋体"] },
      { family: "Bremen Bd BT", names: ["Courier New"] },
    ]);
    const requests = collectFontRequests(document.scene.objects, document.scene.fontAlternateNames);
    const authored = requests.find(face => face.family === "Bremen Bd BT" && face.weight === 700);
    const alternate = requests.find(face => face.family === "Courier New" && face.weight === 700);
    assert.ok(authored && alternate);
    assert.deepEqual(alternate.codePoints, authored.codePoints);
  } finally { document.close(); core.close(); }
});

test("complex0 supplies declared alternate faces to an inline font provider", async () => {
  const restore = installFontRuntime(new FakeFontSet());
  const requested = [];
  const engine = await createOfficeEngine({ execution: "inline", wasm, fontPolicy: "deterministic",
    fontProvider(faces) {
      requested.push(...faces);
      return faces.filter(face => face.family === "Courier New")
        .map(face => ({ ...face, bytes: Uint8Array.of(1) }));
    },
  });
  let document;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/complex0.docx", import.meta.url)));
    const faces = requested.filter(face => face.family === "Courier New");
    assert.deepEqual(faces.map(face => [face.style, face.weight]).sort(), [
      ["italic", 400], ["italic", 700], ["normal", 400], ["normal", 700],
    ]);
    const objects = await document.listObjects({ unitIndex: 0, textOnly: true });
    const runs = objects.flatMap(object => object.fontRuns ?? []).filter(run => run.authoredFamily === "Bremen Bd BT");
    assert.ok(runs.length >= 4);
    assert.ok(runs.every(run => run.source === "fallback" && /provider/u.test(run.renderedFamily)));
  } finally { document?.close(); engine.close(); restore(); }
});

test("declared alternate fonts retain face traits and precede approximate fallback", async () => {
  const assets = prepareFontAssets([
    { family: "Second", bytes: Uint8Array.of(1), style: "italic", weight: 700 },
    { family: "First", bytes: Uint8Array.of(2), style: "italic", weight: 700 },
  ]);
  const face = { family: "Authored", style: "italic", weight: 700, stretch: "normal" };
  const options = { fontAlternateNames: [{ family: "Authored", names: ["Missing", "First", "Second"] }] };
  const registration = await registerFonts(assets, undefined,
    { FontFace: LoadedFontFace, fontSet: new FakeFontSet() }, undefined, [], [], options);
  try {
    assert.equal(registration.resolveFace(face).family, registration.resolveFace({ ...face, family: "First" }).family);
    assert.equal(registration.resolveFace(face).source, "fallback");
    assert.equal(registration.resolveFace(face).style, "italic");
    assert.equal(registration.resolveFace(face).weight, 700);
    assert.equal(registration.hasExactFace(face), false);
  } finally { registration.close(); }
});

test("declared local names preserve bold italic faces and exact authored fonts win", async () => {
  class SelectiveFace extends LoadedFontFace {
    async load() {
      if (this.source === 'local("Courier New")') return this;
      return super.load();
    }
  }
  const face = { family: "Bremen Bd BT", style: "italic", weight: 700, stretch: "normal" };
  const options = { policy: "local-first", requestedFaces: [face],
    fontAlternateNames: [{ family: face.family, names: ["Courier New"] }] };
  for (const source of ["fallback", "host", "provider"]) {
    const fontSet = new FakeFontSet();
    const assets = source === "fallback" ? [] : prepareFontAssets([{ family: face.family, bytes: Uint8Array.of(1), style: face.style, weight: face.weight }]);
    const registration = await registerFonts(source === "host" ? assets : [], undefined,
      { FontFace: SelectiveFace, fontSet }, undefined, [], [],
      { ...options, providerAssets: source === "provider" ? assets : [] });
    try {
      const resolved = registration.resolveFace(face);
      assert.equal(resolved.style, face.style);
      assert.equal(resolved.weight, face.weight);
      assert.equal(resolved.source, source);
      if (source === "fallback") {
        assert.equal(resolved.family, "Courier New");
        assert.ok([...fontSet.faces].some(font => font.source === 'local("Courier New")'
          && font.descriptors.style === 'italic' && font.descriptors.weight === '700'));
      }
    } finally { registration.close(); }
  }
});

test("declared names precede compatibility substitutions across local and host faces", async () => {
  class SelectiveFace extends LoadedFontFace {
    async load() {
      if (['local("Courier New")', 'local("Carlito")'].includes(this.source)) return this;
      return super.load();
    }
  }
  const face = { family: "Authored", style: "normal", weight: 400, stretch: "normal" };
  const options = { policy: "local-first", requestedFaces: [face, { ...face, family: "Calibri" }],
    fontAlternateNames: [{ family: face.family, names: ["Calibri", "Courier New"] }] };
  for (const supplied of [false, true]) {
    const assets = supplied ? prepareFontAssets([{ family: "Calibri", bytes: Uint8Array.of(1) }]) : [];
    const registration = await registerFonts(assets, undefined,
      { FontFace: SelectiveFace, fontSet: new FakeFontSet() }, undefined, [], [], options);
    try {
      assert.equal(registration.resolveFace(face).family, supplied
        ? registration.resolveFace({ ...face, family: "Calibri" }).family : "Courier New");
      assert.equal(registration.resolveFace(face).source, "fallback");
    } finally { registration.close(); }
  }
});

test("font assets normalize descriptors and copy only the selected byte view", () => {
  const input = Uint8Array.of(99, 1, 2, 3, 99).subarray(1, 4);
  const fonts = prepareFontAssets([{
    family: "  Local Sans  ",
    bytes: input,
    style: "italic",
    weight: 650,
    stretch: "condensed",
  }], 64);

  input[0] = 42;
  assert.equal(Object.isFrozen(fonts), true);
  assert.equal(Object.isFrozen(fonts[0]), true);
  assert.equal(fonts[0].family, "Local Sans");
  assert.equal(fonts[0].style, "italic");
  assert.equal(fonts[0].weight, 650);
  assert.equal(fonts[0].stretch, "condensed");
  assert.deepEqual([...new Uint8Array(fonts[0].bytes)], [1, 2, 3]);
});

test("font assets accept a byte view created in another JavaScript realm", () => {
  const foreign = vm.runInNewContext(
    "Uint8Array.from([99, 1, 2, 3, 99]).subarray(1, 4)",
  );
  assert.equal(foreign instanceof Uint8Array, false);

  const fonts = prepareFontAssets([{ family: "Realm Sans", bytes: foreign }], 64);
  assert.equal(fonts[0].bytes instanceof ArrayBuffer, true);
  assert.deepEqual([...new Uint8Array(fonts[0].bytes)], [1, 2, 3]);
});

test("font byte budget is checked before any asset is copied", () => {
  assert.throws(
    () => prepareFontAssets([{ family: "Too Large", bytes: new SliceTrap(8) }], 4),
    (error) => error?.code === "FONT_BYTES_LIMIT",
  );
});

test("font registration loads binary faces and close removes them", async () => {
  const fontSet = new FakeFontSet();
  const prepared = prepareFontAssets([{
    family: "Lifecycle Sans",
    bytes: Uint8Array.of(1, 2, 3),
    style: "oblique",
    weight: 500,
    stretch: "expanded",
  }], 64);
  const registration = await registerFonts(prepared, 50, {
    FontFace: LoadedFontFace,
    fontSet,
  });
  const [face] = [...fontSet.faces];

  assert.match(face.family, /^OfficeViewer \d+ host 0$/u);
  assert.deepEqual(registration.resolve("Lifecycle Sans"), {
    family: face.family,
    source: "host",
  });
  assert.deepEqual(registration.imageCodecFonts(), prepared);
  assert.deepEqual(face.descriptors, { style: "oblique", weight: "500", stretch: "expanded" });
  registration.close();
  registration.close();
  assert.equal(fontSet.faces.size, 0);
  assert.equal(fontSet.deleted.length, 1);
});

test("font registration can append a lazily materialized embedded face", async () => {
  const fontSet = new FakeFontSet();
  const registration = await registerFonts([], 50, {
    FontFace: LoadedFontFace,
    fontSet,
  });

  await registration.addEmbeddedAssets([{
    family: "Lazy PDF Sans",
    bytes: Uint8Array.of(1, 2, 3).buffer,
    style: "normal",
    weight: 400,
    stretch: "normal",
  }], 50);

  assert.equal(registration.resolve("Lazy PDF Sans").source, "embedded");
  assert.equal(fontSet.faces.size, 1);
  registration.close();
  assert.equal(fontSet.faces.size, 0);
});

test("PDF object aliases prefer their embedded face and recover the source family on failure", async () => {
  const fontSet = new FakeFontSet();
  const host = prepareFontAssets(
    [{ family: "Arial", bytes: Uint8Array.of(1) }],
    64,
  );
  const embedded = prepareFontAssets(
    [{ family: "ABCDEF+Arial PDF 5 0", bytes: Uint8Array.of(2) }],
    64,
  );
  const registration = await registerFonts(
    host,
    50,
    { FontFace: LoadedFontFace, fontSet },
    undefined,
    embedded,
  );

  assert.equal(registration.resolve("ABCDEF+Arial PDF 5 0").source, "embedded");
  registration.close();

  const fallbackSet = new FakeFontSet();
  const fallback = await registerFonts(host, 50, {
    FontFace: LoadedFontFace,
    fontSet: fallbackSet,
  });
  assert.equal(fallback.resolve("ABCDEF+Arial PDF 5 0").source, "fallback");
  assert.equal(
    fallback.resolve("ABCDEF+Arial PDF 5 0").family,
    fallback.resolve("Arial").family,
  );
  fallback.close();

  const browserFallback = await registerFonts([], 50, {
    FontFace: LoadedFontFace,
    fontSet: new FakeFontSet(),
  });
  assert.deepEqual(browserFallback.resolve("ABCDEF+Arial PDF 5 0"), {
    family: "Arial",
    source: "fallback",
  });
  browserFallback.close();
});

test("font provider receives the source family after unresolved PDF aliases pass preflight", async () => {
  const calls = [];
  const cache = new FontProviderCache(async (requests) => {
    calls.push(requests);
    return [{ family: "Arial", bytes: Uint8Array.of(1, 2, 3) }];
  }, {
    maxBytes: 64,
    timeoutMs: 50,
  });

  const result = await cache.resolve([
    {
      family: "ABCDEF+Arial PDF 5 0",
      style: "normal",
      weight: 400,
      stretch: "normal",
      codePoints: [0x41],
    },
    {
      family: "UVWXYZ+Arial PDF 8 0",
      style: "normal",
      weight: 400,
      stretch: "normal",
      codePoints: [0x42],
    },
  ]);

  assert.deepEqual(calls, [[{
    family: "Arial",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41, 0x42],
  }]]);
  assert.equal(result.assets[0].family, "Arial");
  assert.deepEqual(result.diagnostics, []);
  cache.close();
});

test("embedded font rejection is isolated and reported without hiding successful faces", async () => {
  class SelectiveFontFace extends LoadedFontFace {
    async load() {
      if (this.source.byteLength === 1) throw new Error("invalid embedded font");
      return this;
    }
  }
  const fontSet = new FakeFontSet();
  const embedded = prepareFontAssets([
    { family: "Good PDF Sans", bytes: Uint8Array.of(1, 2) },
    { family: "Bad PDF Sans", bytes: Uint8Array.of(1) },
  ], 64);
  const registration = await registerFonts(
    [],
    50,
    { FontFace: SelectiveFontFace, fontSet },
    undefined,
    embedded,
  );

  assert.equal(registration.resolve("Good PDF Sans").source, "embedded");
  assert.equal(registration.resolve("Bad PDF Sans").source, "fallback");
  assert.equal(fontSet.faces.size, 1);
  assert.deepEqual(registration.diagnostics(), [{
    code: "FONT_LOAD_FAILED",
    severity: "warning",
    fidelity: "approximate",
    phase: "render",
    message: "Embedded font \"Bad PDF Sans\" could not be loaded; this attempt was skipped",
    details: {
      source: "embedded",
      family: "Bad PDF Sans",
      style: "normal",
      weight: 400,
      stretch: "normal",
      bytes: 1,
      error: "invalid embedded font",
    },
  }]);
  registration.close();
});

test("lazy embedded font failures remain non-fatal and diagnostics are deduplicated", async () => {
  class FailedFontFace extends LoadedFontFace {
    async load() {
      throw new Error("lazy font rejected");
    }
  }
  const fontSet = new FakeFontSet();
  const registration = await registerFonts([], 50, {
    FontFace: FailedFontFace,
    fontSet,
  });
  const asset = {
    family: "Lazy Broken PDF Sans",
    bytes: Uint8Array.of(1).buffer,
    style: "normal",
    weight: 400,
    stretch: "normal",
  };

  await registration.addEmbeddedAssets([asset], 50);
  await registration.addEmbeddedAssets([asset], 50);

  assert.equal(fontSet.faces.size, 0);
  assert.equal(registration.diagnostics().length, 1);
  assert.equal(registration.diagnostics()[0].details.error, "lazy font rejected");
  registration.close();
});

test("inline lazy PDF font failures update document diagnostics without hiding the page", async () => {
  class FailedFontFace extends LoadedFontFace {
    async load() {
      throw new Error("browser rejected lazy PDF font");
    }
  }
  const pdf = new TextEncoder().encode(`%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 /MediaBox [0 0 100 100] >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 7 0 R >> >> /Contents 5 0 R >> endobj
4 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F2 8 0 R >> >> /Contents 6 0 R >> endobj
5 0 obj << >> stream
BT /F1 10 Tf 1 0 0 1 10 80 Tm (First) Tj ET
endstream endobj
6 0 obj << >> stream
BT /F2 10 Tf 1 0 0 1 10 80 Tm (A) Tj ET
endstream endobj
7 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
8 0 obj << /Type /Font /Subtype /TrueType /BaseFont /LazyBad /FirstChar 65 /LastChar 65 /Widths [600] /FontDescriptor 9 0 R >> endobj
9 0 obj << /Type /FontDescriptor /FontName /LazyBad /Flags 32 /Ascent 800 /Descent -200 /FontBBox [0 -200 1000 800] /FontFile2 10 0 R >> endobj
10 0 obj << /Length 4 >> stream
true
endstream endobj
trailer << /Root 1 0 R >>
%%EOF`);
  const fontSet = new FakeFontSet();
  const restore = installFontRuntime(fontSet, FailedFontFace);
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(pdf);
    assert.equal(
      document.diagnostics().some(({ code }) => code === "FONT_LOAD_FAILED"),
      false,
    );

    const objects = await document.listObjects({ unitIndex: 1, textOnly: true });

    assert.deepEqual(objects.map(({ text }) => text), ["A"]);
    assert.equal(
      document.diagnostics().some(({ code }) => code === "FONT_LOAD_FAILED"),
      true,
    );
    assert.equal(fontSet.faces.size, 0);
  } finally {
    document?.close();
    engine.close();
    restore();
  }
});

test("embedded font timeout remains non-fatal and observable", async () => {
  class StalledFontFace extends LoadedFontFace {
    load() {
      return new Promise(() => {});
    }
  }
  const fontSet = new FakeFontSet();
  const embedded = prepareFontAssets(
    [{ family: "Stalled PDF Sans", bytes: Uint8Array.of(1) }],
    64,
  );
  const registration = await registerFonts(
    [],
    5,
    { FontFace: StalledFontFace, fontSet },
    undefined,
    embedded,
  );

  assert.equal(fontSet.faces.size, 0);
  assert.equal(registration.diagnostics()[0].code, "FONT_LOAD_TIMEOUT");
  assert.equal(registration.diagnostics()[0].details.family, "Stalled PDF Sans");
  registration.close();
});

test("one stalled embedded font does not discard another face that already loaded", async () => {
  class SelectiveFontFace extends LoadedFontFace {
    load() {
      return this.source.byteLength === 1
        ? new Promise(() => {})
        : Promise.resolve(this);
    }
  }
  const fontSet = new FakeFontSet();
  const embedded = prepareFontAssets([
    { family: "Loaded PDF Sans", bytes: Uint8Array.of(1, 2) },
    { family: "Stalled PDF Sans", bytes: Uint8Array.of(1) },
  ], 64);
  const registration = await registerFonts(
    [],
    5,
    { FontFace: SelectiveFontFace, fontSet },
    undefined,
    embedded,
  );

  assert.equal(registration.resolve("Loaded PDF Sans").source, "embedded");
  assert.equal(registration.resolve("Stalled PDF Sans").source, "fallback");
  assert.equal(fontSet.faces.size, 1);
  assert.deepEqual(registration.diagnostics().map((diagnostic) => ({
    code: diagnostic.code,
    family: diagnostic.details.family,
  })), [{
    code: "FONT_LOAD_TIMEOUT",
    family: "Stalled PDF Sans",
  }]);
  registration.close();
});

test("one stalled local face does not discard another compatible face that already loaded", async () => {
  class SelectiveLocalFontFace extends LoadedFontFace {
    load() {
      if (this.source === 'local("宋体")') return new Promise(() => {});
      if (this.source === 'local("SimSun")') return Promise.resolve(this);
      return Promise.reject(new Error("local font unavailable"));
    }
  }
  const fontSet = new FakeFontSet();
  const registration = await registerFonts(
    [],
    5,
    { FontFace: SelectiveLocalFontFace, fontSet },
    undefined,
    [],
    [],
    {
      policy: "local-first",
      requestedFaces: [{
        family: "宋体",
        style: "normal",
        weight: 400,
        stretch: "normal",
        codePoints: [0x4e2d],
      }],
    },
  );

  assert.equal(registration.resolve("宋体").source, "browser");
  assert.equal(registration.resolve("宋体").spaceAdvanceEm, undefined, "SimSun already has the authored spaces");
  assert.equal(fontSet.faces.size, 1);
  assert.equal([...fontSet.faces][0].source, 'local("SimSun")');
  registration.close();
});

test("Song substitutes preserve half-em blanks without changing exact or unrelated fonts", async () => {
  for (const [authored, local, expected] of [
    ["宋体", "STSong", .5], ["SimSun", "Songti SC", .5], ["NSimSun", "STSong", .5],
    ["宋体", "SimSun", undefined], ["STSong", "STSong", undefined], ["Times New Roman", "Times New Roman", undefined],
  ]) {
    class LocalFace extends LoadedFontFace {
      async load() { if (this.source !== `local(${JSON.stringify(local)})`) throw Error("unavailable"); return this; }
    }
    const fonts = await registerFonts([], 50, { FontFace: LocalFace, fontSet: new FakeFontSet() },
      undefined, [], [authored], { policy: "local-first" });
    assert.equal(fonts.resolve(authored).spaceAdvanceEm, expected, `${authored} -> ${local}`);
    assert.equal(fonts.resolveFace({ family: authored, style: "italic", weight: 700, stretch: "normal" }).spaceAdvanceEm,
      expected, "style fallback retains the advance");
    fonts.close();
  }
});

test("local-first PDF object fonts probe the exact source family", async () => {
  class PdfLocalFontFace extends LoadedFontFace {
    async load() {
      if (this.source === 'local("Arial")') return this;
      throw new Error("local font unavailable");
    }
  }
  const fontSet = new FakeFontSet();
  const registration = await registerFonts(
    [],
    50,
    { FontFace: PdfLocalFontFace, fontSet },
    undefined,
    [],
    [],
    {
      policy: "local-first",
      requestedFaces: [{
        family: "Arial PDF 5 0",
        style: "normal",
        weight: 400,
        stretch: "normal",
        codePoints: [0x4e2d],
      }],
    },
  );

  assert.equal(registration.resolve("Arial PDF 5 0").source, "browser");
  assert.equal(fontSet.faces.size, 1);
  assert.equal([...fontSet.faces][0].source, 'local("Arial")');
  registration.close();
});

test("local-first PDF object fonts probe compatible PostScript face names", async () => {
  class PdfLocalFontFace extends LoadedFontFace {
    async load() {
      if (this.source === 'local("Futura-Medium")') return this;
      throw new Error("local font unavailable");
    }
  }
  class ProbeContext {
    font = "";

    measureText(text) {
      const advance = this.font.includes("OfficeViewer") || this.font.includes("Futura-Medium")
        ? 9
        : 8;
      return { width: text.length * advance };
    }
  }
  class ProbeCanvas {
    getContext(kind) { return kind === "2d" ? new ProbeContext() : null; }
  }
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: ProbeCanvas });
  const fontSet = new FakeFontSet();
  try {
    const registration = await registerFonts(
      [],
      50,
      { FontFace: PdfLocalFontFace, fontSet },
      undefined,
      [],
      [],
      {
        policy: "local-first",
        requestedFaces: [{
          family: "ABCDEF+Futura-Boo PDF 56 0",
          style: "normal",
          weight: 400,
          stretch: "normal",
          codePoints: [0x46, 0x75, 0x74, 0x72, 0x61],
        }],
      },
    );

    assert.equal(registration.resolve("ABCDEF+Futura-Boo PDF 56 0").source, "browser");
    assert.deepEqual(
      [...fontSet.faces].map((face) => face.source),
      ['local("Futura-Medium")'],
    );
    registration.close();
  } finally {
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("missing Windows CJK PDF fonts follow Preview's Helvetica fallback", async () => {
  const registration = await registerFonts([], 50, {
    FontFace: LoadedFontFace,
    fontSet: new FakeFontSet(),
  });

  for (const family of ["SimSun", "KaiTi", "SimHei", "DengXian"]) {
    assert.deepEqual(registration.resolve(`${family} PDF 5 0`), {
      family: "Helvetica",
      source: "fallback",
    });
  }
  registration.close();
});

test("font loading diagnostics stay bounded and make truncation visible", async () => {
  class FailedFontFace extends LoadedFontFace {
    async load() {
      throw new Error("font rejected");
    }
  }
  const fontSet = new FakeFontSet();
  const embedded = prepareFontAssets(
    Array.from({ length: 130 }, (_, index) => ({
      family: `Broken PDF Sans ${index}`,
      bytes: Uint8Array.of(index & 0xff),
    })),
    1_024,
  );
  const registration = await registerFonts(
    [],
    50,
    { FontFace: FailedFontFace, fontSet },
    undefined,
    embedded,
  );
  const diagnostics = registration.diagnostics();

  assert.equal(diagnostics.length, 128);
  assert.equal(diagnostics.filter(({ code }) => code === "FONT_LOAD_FAILED").length, 127);
  assert.deepEqual(diagnostics.at(-1), {
    code: "FONT_LOAD_DIAGNOSTICS_TRUNCATED",
    severity: "warning",
    fidelity: "approximate",
    phase: "render",
    message: "Additional font loading failures were omitted after the diagnostic limit was reached",
    details: { limit: 128, retained: 127 },
  });
  registration.close();
});

test("local-first font resolution prefers embedded assets, then exact local fonts, then explicit approximation", async () => {
  class SelectiveFontFace extends LoadedFontFace {
    async load() {
      if (typeof this.source !== "string") return this;
      if (this.source.includes("Browser Sans")) return this;
      throw new Error("local font unavailable");
    }
  }
  const fontSet = new FakeFontSet();
  const host = prepareFontAssets([{
    family: "Priority Sans",
    bytes: Uint8Array.of(1),
  }], 64);
  const embedded = [{
    family: "Priority Sans",
    bytes: Uint8Array.of(2).buffer,
    style: "normal",
    weight: 400,
    stretch: "normal",
  }];
  const registration = await registerFonts(
    host,
    50,
    { FontFace: SelectiveFontFace, fontSet },
    undefined,
    embedded,
    ["Priority Sans", "Browser Sans", "Missing Serif", "Missing Code"],
    { policy: "local-first" },
  );

  assert.equal(registration.resolve("Priority Sans").source, "embedded");
  assert.equal(registration.resolve("Browser Sans").source, "browser");
  assert.deepEqual(registration.resolve("Missing Serif"), { family: "serif", source: "fallback" });
  assert.deepEqual(registration.resolve("Missing Code"), { family: "monospace", source: "fallback" });
  assert.deepEqual(registration.imageCodecFonts().map(({ family, bytes }) => [family, [...new Uint8Array(bytes)]]), [
    ["Priority Sans", [2]],
  ]);
  registration.close();
});

test("local-first probing rejects a loaded face whose canvas metrics are only browser fallback", async () => {
  class FallbackMaskingFontFace extends LoadedFontFace {
    async load() { return this; }
  }
  class ProbeContext {
    font = "";

    measureText(text) {
      const advance = this.font.includes("Measured Sans") ? 9 : 8;
      return { width: text.length * advance };
    }
  }
  class ProbeCanvas {
    getContext(kind) { return kind === "2d" ? new ProbeContext() : null; }
  }
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: ProbeCanvas });
  const fontSet = new FakeFontSet();
  try {
    const registration = await registerFonts(
      [],
      50,
      { FontFace: FallbackMaskingFontFace, fontSet },
      undefined,
      [],
      ["Fallback Masked Sans", "Measured Sans"],
      { policy: "local-first" },
    );

    assert.equal(registration.resolve("Fallback Masked Sans").source, "fallback");
    assert.equal(registration.resolve("Measured Sans").source, "browser");
    assert.equal(fontSet.faces.size, 1);
    registration.close();
  } finally {
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("supplied Pages CSS family survives local() requiring a PostScript face name", async () => {
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("./fixtures/newletter2.pages", import.meta.url)));
  const requestedFaces = collectFontRequests(document.scene.objects).filter((face) => ["Superclarendon", "Charter-Roman"].includes(face.family));
  assert.ok(requestedFaces.some((face) => face.weight === 700));
  document.close();
  core.close();
  class MissingLocalFace extends LoadedFontFace {
    async load() {
      if (this.source === 'local("Charter-Roman")') return this;
      throw new Error("local face name unavailable");
    }
  }
  class ProbeCanvas {
    getContext() {
      return { font: "", measureText(text) {
        return { width: text.length * (this.font.includes('"Superclarendon"') || this.font.includes('"OfficeViewer ') ? 10 : 8) };
      } };
    }
  }
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: ProbeCanvas });
  const fontSet = new FakeFontSet();
  try {
    const registration = await registerFonts([], 50, { FontFace: MissingLocalFace, fontSet },
      undefined, [], ["Missing Serif"], { policy: "local-first", requestedFaces });
    assert.deepEqual(registration.resolve("Superclarendon"), { family: "Superclarendon", source: "browser" });
    assert.match(registration.resolve("Charter-Roman").family, /^OfficeViewer /,
      "an available PostScript face must render through its loaded alias when CSS cannot resolve the authored name");
    assert.equal(registration.resolve("Missing Serif").source, "fallback");
    assert.equal(fontSet.faces.size, 1);
    registration.close();
  } finally {
    if (descriptor) Object.defineProperty(globalThis, "OffscreenCanvas", descriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("local-first probing rejects a compatible local alias with different glyph metrics", async () => {
  class CompatibleAliasFace extends LoadedFontFace {
    async load() {
      if (this.source !== 'local("STSong")') throw new Error("local face is unavailable");
      return this;
    }
  }
  class ProbeContext {
    font = "";

    measureText(text) {
      const advance = this.font.includes("OfficeViewer")
        ? 10
        : this.font.includes("STSong")
          ? 9
          : 8;
      return { width: text.length * advance };
    }
  }
  class ProbeCanvas {
    getContext(kind) { return kind === "2d" ? new ProbeContext() : null; }
  }
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: ProbeCanvas });
  const fontSet = new FakeFontSet();
  try {
    const registration = await registerFonts(
      [],
      50,
      { FontFace: CompatibleAliasFace, fontSet },
      undefined,
      [],
      ["宋体"],
      { policy: "local-first" },
    );

    assert.equal(registration.resolve("宋体").source, "fallback");
    assert.equal(fontSet.faces.size, 0);
    registration.close();
  } finally {
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("missing Office sans text reuses a managed Calibri face before an environment fallback", async () => {
  const fontSet = new FakeFontSet();
  const calibri = prepareFontAssets([{
    family: "Calibri",
    bytes: Uint8Array.of(1, 2, 3),
  }], 64);
  const registration = await registerFonts(calibri, 50, {
    FontFace: LoadedFontFace,
    fontSet,
  });

  const managedCalibri = registration.resolve("Calibri");
  const substituted = registration.resolve("Uncatalogued Office Sans");
  assert.equal(managedCalibri.source, "host");
  assert.deepEqual(substituted, {
    family: managedCalibri.family,
    source: "fallback",
  });
  registration.close();
});

test("Calibri text reuses a managed Carlito face before an environment fallback", async () => {
  const fontSet = new FakeFontSet();
  const carlito = prepareFontAssets([{
    family: "Carlito",
    bytes: Uint8Array.of(1, 2, 3),
  }], 64);
  const registration = await registerFonts(carlito, 50, {
    FontFace: LoadedFontFace,
    fontSet,
  });

  const managedCarlito = registration.resolve("Carlito");
  assert.equal(managedCarlito.source, "host");
  assert.deepEqual(registration.resolve("Calibri"), {
    family: managedCarlito.family,
    source: "fallback",
  });
  registration.close();
});

test("missing Office sans fallback selects managed faces from the requested scripts", async () => {
  const fontSet = new FakeFontSet();
  const managed = prepareFontAssets([
    { family: "Calibri", bytes: Uint8Array.of(1) },
    { family: "Meiryo UI", bytes: Uint8Array.of(2) },
  ], 64);
  const requestedFaces = collectFontRequests([
    { text: "Latin", visual: { kind: "text", fontFamily: "Missing Latin Sans", italic: false, bold: false } },
    { text: "中文", visual: { kind: "text", fontFamily: "Missing CJK Sans", italic: false, bold: false } },
    { text: "A中", visual: { kind: "text", fontFamily: "Missing Mixed Sans", italic: false, bold: false } },
  ]);
  const registration = await registerFonts(
    managed,
    50,
    { FontFace: LoadedFontFace, fontSet },
    undefined,
    [],
    requestedFaces.map(({ family }) => family),
    { policy: "local-first", requestedFaces },
  );

  assert.deepEqual(registration.resolve("Missing Latin Sans"), {
    family: registration.resolve("Calibri").family,
    source: "fallback",
  });
  assert.deepEqual(registration.resolve("Missing CJK Sans"), {
    family: registration.resolve("Meiryo UI").family,
    source: "fallback",
  });
  assert.deepEqual(registration.resolve("Missing Mixed Sans"), {
    family: "Calibri",
    source: "fallback",
  });
  registration.close();
});

test("uses the native macOS Japanese Gothic fallback with a portable sans-serif tail", () => {
  assert.equal(approximateFontFamily("Aptos"), "Calibri");
  assert.equal(approximateFontFamily("Calibri Light"), "Calibri");
  assert.equal(approximateFontFamily("Segoe UI (Body)"), "Segoe UI");
  assert.equal(approximateFontFamily("SegoeUISymbol"), "Arial");
  assert.equal(approximateFontFamily("SegoeUISymbol", [0xf123]), "Segoe UI Symbol");
  assert.equal(approximateFontFamily("JRJDDM+ArialUnicodeMS"), "Arial");
  assert.equal(approximateFontFamily("ArialMT"), "Arial");
  assert.equal(approximateFontFamily("Liberation Sans"), "Arial", "tdf96206.odp and its Office ArialMT reference");
  assert.equal(approximateFontFamily("Arial-BoldMT"), "Arial");
  assert.equal(approximateFontFamily("Arial-ItalicMT"), "Arial");
  assert.equal(approximateFontFamily("LucidaSans-Demi"), "Arial");
  assert.equal(approximateFontFamily("CenturyGothic Bold"), "Arial");
  assert.equal(approximateFontFamily("BankGothic Lt BT"), "Copperplate");
  assert.equal(approximateFontFamily("Fira Code"), "monospace");
  assert.equal(approximateFontFamily("IYCZZB+CMR10"), "serif");
  assert.equal(approximateFontFamily("Book Antiqua"), "serif");
  assert.equal(approximateFontFamily("ABCDEF+JansonText-Bold"), "serif");
  assert.equal(approximateFontFamily("ABCDEF+NimbusRomNo9L-Medi"), "serif");
  assert.equal(approximateFontFamily("MMPHFK+ACaslon-Regular"), "serif");
  assert.equal(approximateFontFamily("NuptialScript"), "serif");
  assert.equal(approximateFontFamily("ABCDEF+CMTT10"), "monospace");
  assert.equal(approximateFontFamily("ABCDEF+CMSS10"), "Arial");
  assert.equal(approximateFontFamily("Uncatalogued Office Sans"), "Calibri");
  assert.equal(approximateFontFamily("Uncatalogued Office Sans", [0x416]), "Calibri");
  assert.equal(approximateFontFamily("Uncatalogued Office Sans", [0x4e2d]), "Hiragino Sans");
  assert.equal(approximateFontFamily("Avenir Next", [0x4e2d]), "SimSun");
  assert.equal(approximateFontFamily("Uncatalogued Office Sans", [0x627]), "Arial");
  assert.equal(approximateFontFamily("Uncatalogued Office Sans", [0xf123]), "Segoe UI Symbol");
  assert.equal(approximateFontFamily("Yu Gothic UI Semibold"), "Hiragino Sans");
  assert.equal(
    fontShorthand("Hiragino Sans", 48, false, false),
    '400 48px "Hiragino Sans", "Yu Gothic UI", "Yu Gothic", "Meiryo UI", "Meiryo", "Microsoft YaHei", sans-serif',
  );
  assert.equal(
    fontShorthand("Calibri", 48, false, false),
    '400 48px "Calibri", "Carlito", "Aptos", "Segoe UI", "Arial", "Arial Unicode MS", "Microsoft YaHei", "Yu Gothic UI", "Meiryo UI", "Hiragino Sans", "Noto Sans Arabic", "Noto Sans Hebrew", "Noto Sans Devanagari", "Noto Sans Thai", "Segoe UI Symbol", "Apple Symbols", sans-serif',
  );
  assert.equal(
    fontShorthand("Calibri", 48, false, true),
    '700 48px "Calibri", "Carlito", "Aptos", "Segoe UI", "Arial", "Arial Unicode MS", "Microsoft YaHei", "Yu Gothic UI", "Meiryo UI", "Hiragino Sans", "Noto Sans Arabic", "Noto Sans Hebrew", "Noto Sans Devanagari", "Noto Sans Thai", "Segoe UI Symbol", "Apple Symbols", sans-serif',
  );
  assert.equal(
    fontShorthand("Segoe UI", 48, false, false),
    '400 48px "Segoe UI", "Arial", "Arial Unicode MS", "Microsoft YaHei", "Yu Gothic UI", "Meiryo UI", "Hiragino Sans", "Noto Sans Arabic", "Noto Sans Hebrew", "Noto Sans Devanagari", "Noto Sans Thai", "Segoe UI Symbol", "Apple Symbols", sans-serif',
  );
});

test("browser font probing uses exact subset coverage instead of fixed samples", () => {
  const measured = [];
  const context = {
    font: "400 12px sans-serif",
    measureText(text) {
      measured.push(text);
      const exactSubset = this.font.includes("Private Icon Subset") && text.includes("\uf123");
      return { width: text.length * (exactSubset ? 9 : 8) };
    },
  };
  const available = probeBrowserFontFace({
    family: "Private Icon Subset",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0xf123],
  }, context);

  assert.equal(available, true);
  assert.equal(measured.every((text) => text === "\uf123"), true);
  assert.equal(context.font, "400 12px sans-serif");
});

test("font load failures and timeouts are explicit and clean registrations", async () => {
  class FailedFontFace extends LoadedFontFace {
    async load() {
      throw new Error("bad font");
    }
  }
  class StalledFontFace extends LoadedFontFace {
    load() {
      return new Promise(() => {});
    }
  }
  const prepared = prepareFontAssets([{ family: "Bad Sans", bytes: Uint8Array.of(1) }], 64);

  const failedSet = new FakeFontSet();
  await assert.rejects(
    registerFonts(prepared, 50, { FontFace: FailedFontFace, fontSet: failedSet }),
    (error) => error?.code === "FONT_LOAD_FAILED",
  );
  assert.equal(failedSet.faces.size, 0);

  const stalledSet = new FakeFontSet();
  await assert.rejects(
    registerFonts(prepared, 5, { FontFace: StalledFontFace, fontSet: stalledSet }),
    (error) => error?.code === "FONT_LOAD_TIMEOUT",
  );
  assert.equal(stalledSet.faces.size, 0);
});

test("aborting font registration removes faces and reports operation cancellation", async () => {
  class StalledFontFace extends LoadedFontFace {
    load() {
      return new Promise(() => {});
    }
  }
  const prepared = prepareFontAssets([{ family: "Abort Sans", bytes: Uint8Array.of(1) }], 64);
  const fontSet = new FakeFontSet();
  const controller = new AbortController();
  const registering = registerFonts(
    prepared,
    20,
    { FontFace: StalledFontFace, fontSet },
    controller.signal,
  );

  controller.abort("cancelled");

  await assert.rejects(registering, (error) => error?.code === "OPERATION_ABORTED");
  assert.equal(fontSet.faces.size, 0);
});

test("missing FontFace support is reported explicitly", async () => {
  const fontFaceDescriptor = Object.getOwnPropertyDescriptor(globalThis, "FontFace");
  const fontsDescriptor = Object.getOwnPropertyDescriptor(globalThis, "fonts");
  delete globalThis.FontFace;
  delete globalThis.fonts;
  try {
    const prepared = prepareFontAssets([{ family: "Local Sans", bytes: Uint8Array.of(1) }], 64);
    await assert.rejects(
      registerFonts(prepared),
      (error) => error?.code === "FONT_ENVIRONMENT_UNAVAILABLE",
    );
  } finally {
    if (fontFaceDescriptor) Object.defineProperty(globalThis, "FontFace", fontFaceDescriptor);
    if (fontsDescriptor) Object.defineProperty(globalThis, "fonts", fontsDescriptor);
  }
});

test("renderer rejects FontFaceSet fallback false positives without rejecting a measurable face", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = {
    save() {},
    restore() {},
    scale() {},
    translate() {},
    fillRect() {},
    beginPath() {},
    rect() {},
    clip() {},
    fill() {},
    stroke() {},
    font: "",
    measureText(text) {
      const advance = this.font.includes("Available Sans") ? 9 : 8;
      return { width: text.length * advance };
    },
    fillText() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const object = {
    numericId: 1,
    id: "text:1",
    type: "paragraph",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 200, height: 40 },
    text: "Hello",
    source: {
      format: "odt",
      part: "content.xml",
      kind: "text-range",
      path: "/office:document-content/office:body/office:text/text:p[1]",
      textRange: [0, 5],
      mapping: "exact",
    },
    z: 0,
    visual: {
      kind: "text",
      geometry: "rectangle",
      fill: 0,
      stroke: 0,
      strokeWidth: 0,
      fontFamily: "Unavailable Sans",
      fontSize: 12,
      color: 0x000000ff,
      bold: false,
      italic: false,
      align: "start",
    },
  };
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 200, height: 40 };
  try {
    const unavailable = new SceneRenderer([object], DEFAULT_LIMITS, {
      add() {},
      delete() { return true; },
      check() { return false; },
    });
    const unavailableResult = await unavailable.render(unit, { unitIndex: 0 });
    assert.deepEqual(unavailableResult.diagnostics.map((item) => item.code), ["FONT_SUBSTITUTED"]);
    assert.deepEqual(unavailableResult.diagnostics[0].details, {
      family: "Unavailable Sans",
      fallbackFamily: "Calibri",
    });

    const fallbackMasked = new SceneRenderer([object], DEFAULT_LIMITS, {
      add() {},
      delete() { return true; },
      check() { return true; },
    });
    const fallbackMaskedResult = await fallbackMasked.render(unit, { unitIndex: 0 });
    assert.deepEqual(fallbackMaskedResult.diagnostics.map((item) => item.code), ["FONT_SUBSTITUTED"]);

    const availableObject = {
      ...object,
      visual: { ...object.visual, fontFamily: "Available Sans" },
    };
    const usable = new SceneRenderer([availableObject], DEFAULT_LIMITS, {
      add() {},
      delete() { return true; },
      check() { return true; },
    });
    const usableResult = await usable.render(unit, { unitIndex: 0 });
    assert.deepEqual(usableResult.diagnostics, []);
  } finally {
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("normalizes complete Symbol and Wingdings source-byte families from one generated table", () => {
  assert.equal(semanticSymbolFontText("ABG@\\£³·", "Symbol"), "ΑΒΓ≅∴≤≥•");
  assert.equal(semanticSymbolFontText("ABG", "ABCDEF+Symbol PDF 5 0"), "ΑΒΓ");
  assert.equal(
    semanticSymbolFontText("\uf041\uf042\uf047\uf040\uf05c\uf0a3\uf0b3\uf0b7", "\"symbol\""),
    "ΑΒΓ≅∴≤≥•",
  );
  assert.equal(semanticSymbolFontText("!;Jqu§ü", "Wingdings"), "🖉🖴☺❑◆▪✓");
  assert.equal(semanticSymbolFontText("p", "Wingdings-Regular"), "□");
  assert.equal(semanticSymbolFontText("!;J", "ABCDEF+Wingdings PDF 8 0"), "🖉🖴☺");
  assert.equal(
    semanticSymbolFontText("\uf021\uf03b\uf04a\uf071\uf075\uf0a7\uf0fc", "'wingdings'"),
    "🖉🖴☺❑◆▪✓",
  );
  assert.equal(semanticSymbolFontText("\uf0d2\uf0d3\uf0d4", "Symbol"), "®©™");
  assert.equal(semanticSymbolFontText("\uf0e2\uf0e3\uf0e4", "Symbol"), "®©™");
  assert.equal(semanticSymbolFontText("\uf04e", "Webdings"), "👁");
  assert.equal(semanticSymbolFontText("ABq", "Arial"), "ABq");
});

test("PDF object aliases preserve symbol semantics for fallback classification", () => {
  const [request] = collectFontRequests([{
    text: "ABG",
    visual: {
      kind: "text",
      fontFamily: "ABCDEF+Symbol PDF 5 0",
      italic: false,
      bold: false,
    },
  }]);

  assert.deepEqual(request.codePoints, [0x391, 0x392, 0x393]);
});

test("private-use symbol glyphs probe the authored font, not its Unicode substitute", () => {
  const [request] = collectFontRequests([{
    text: "\uf0df\uf0e0",
    visual: { kind: "text", fontFamily: "Wingdings", italic: false, bold: false },
  }]);
  assert.deepEqual(request.codePoints, [0xf0df, 0xf0e0]);
  assert.equal(semanticSymbolFontText("\uf0df\uf0e0", "Wingdings"), "←→");
});

test("font requests traverse DrawingML text effects", () => {
  const [request] = collectFontRequests([{
    text: "Effect",
    visual: {
      kind: "text-effects",
      effects: [],
      visual: {
        kind: "text",
        fontFamily: "Effect Sans",
        italic: true,
        bold: true,
      },
    },
  }]);

  assert.deepEqual(request, {
    family: "Effect Sans",
    style: "italic",
    weight: 700,
    stretch: "normal",
    codePoints: [0x45, 0x63, 0x65, 0x66, 0x74],
  });
});

test("renderer preserves exact font glyphs and uses compatible fallback text", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const painted = [];
  const context = {
    globalAlpha: 1,
    letterSpacing: "",
    textBaseline: "",
    textAlign: "",
    direction: "ltr",
    font: "",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, rotate() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    measureText(text) { return { width: text.length * 8 }; },
    fillText(text) { painted.push({ text, font: this.font }); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const object = {
    numericId: 1,
    id: "symbol-bullet",
    type: "paragraph",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 200, height: 40 },
    text: "❑\tItem",
    source: {
      format: "pptx",
      part: "ppt/slides/slide1.xml",
      kind: "text-range",
      path: "/p:sld/p:cSld/p:spTree/p:sp[1]",
      textRange: [0, 6],
      mapping: "exact",
    },
    z: 0,
    visual: {
      kind: "rich-text",
      geometry: "rectangle",
      fill: { kind: "solid", color: 0x00000000 },
      stroke: { kind: "solid", color: 0x00000000 },
      strokeWidth: 0,
      align: "start",
      lineHeight: 0,
      runs: [
        {
          text: "q\t",
          fontFamily: "Wingdings",
          fontSize: 16,
          color: 0x000000ff,
          bold: false,
          italic: false,
          underline: false,
          strikethrough: false,
          highlight: 0,
          baselineShift: 0,
          letterSpacing: 0,
        },
        {
          text: "Item",
          fontFamily: "Main Sans",
          fontSize: 16,
          color: 0x000000ff,
          bold: false,
          italic: false,
          underline: false,
          strikethrough: false,
          highlight: 0,
          baselineShift: 0,
          letterSpacing: 0,
        },
      ],
    },
  };
  const exactFonts = {
    resolve(family) { return { family: `${family} Alias`, source: "host" }; },
    resolveFace(face) { return { ...face, family: `${face.family} Alias`, source: "host" }; },
  };
  const fallbackFonts = {
    resolve(family) {
      return family === "Wingdings" || family === "Symbol" || family === "Algerian"
        ? { family: "Calibri", source: "fallback" }
        : { family: `${family} Alias`, source: "host" };
    },
    resolveFace(face) {
      return face.family === "Wingdings" || face.family === "Symbol" || face.family === "Algerian"
        ? { ...face, family: "Calibri", source: "fallback" }
        : { ...face, family: `${face.family} Alias`, source: "host" };
    },
  };
  const unit = { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 200, height: 40 };

  try {
    assert.deepEqual(documentFontRuns(object, exactFonts), [
      { start: 0, end: 2, authoredFamily: "Wingdings", renderedFamily: "Wingdings Alias", source: "host" },
      { start: 2, end: 6, authoredFamily: "Main Sans", renderedFamily: "Main Sans Alias", source: "host" },
    ]);

    const exact = new SceneRenderer([object], DEFAULT_LIMITS, exactFonts);
    const exactResult = await exact.render(unit, { unitIndex: 0 });
    exactResult.bitmap.close();
    assert.deepEqual(painted.map(({ text }) => text), ["q", "Item"]);
    assert.match(painted[0].font, /"Wingdings Alias"/u);

    painted.length = 0;
    const fallback = new SceneRenderer([object], DEFAULT_LIMITS, fallbackFonts);
    const fallbackResult = await fallback.render(unit, { unitIndex: 0 });
    fallbackResult.bitmap.close();
    assert.deepEqual(painted.map(({ text }) => text), ["❑", "Item"]);
    assert.match(painted[0].font, /"Calibri"/u);

    object.visual.runs[0].text = "\uf0df\uf0e0\t";
    painted.length = 0;
    const arrows = new SceneRenderer([object], DEFAULT_LIMITS, exactFonts);
    (await arrows.render(unit, { unitIndex: 0 })).bitmap.close();
    assert.deepEqual(painted.map(({ text }) => text), ["\uf0df\uf0e0", "Item"]);
    assert.match(painted[0].font, /"Wingdings Alias"/u);

    painted.length = 0;
    const fallbackArrows = new SceneRenderer([object], DEFAULT_LIMITS, fallbackFonts);
    (await fallbackArrows.render(unit, { unitIndex: 0 })).bitmap.close();
    assert.deepEqual(painted.map(({ text }) => text), ["←→", "Item"]);

    object.source.format = "docx";
    object.text = "•\tItem";
    object.visual.runs[0].fontFamily = "Symbol";
    object.visual.runs[0].text = "\uf0b7\t";
    for (const [fonts, glyph] of [[exactFonts, "\uf0b7"], [fallbackFonts, "•"]]) {
      painted.length = 0;
      const renderer = new SceneRenderer([object], DEFAULT_LIMITS, fonts);
      (await renderer.render(unit, { unitIndex: 0 })).bitmap.close();
      assert.deepEqual(painted.map(({ text }) => text), [glyph, "Item"]);
      assert.match(painted[0].font, /16px/u);
      assert.match(painted[0].font, fonts === exactFonts ? /"Symbol Alias"/u : /"Calibri"/u);
    }
    // macOS Symbol exposes Unicode, while the DOCX stores Windows PUA codes.
    context.measureText = (text) => ({ width: text === "•" ? 25 : text.length * 40 });
    painted.length = 0;
    const unicodeSymbol = new SceneRenderer([object], DEFAULT_LIMITS, exactFonts);
    (await unicodeSymbol.render(unit, { unitIndex: 0 })).bitmap.close();
    assert.deepEqual(painted.map(({ text }) => text), ["•", "Item"]);
    assert.match(painted[0].font, /"Symbol Alias"/u);

    object.text = "Do the following";
    object.visual.runs = [{
      ...object.visual.runs[0],
      text: "Do the following",
      fontFamily: "Algerian",
    }];
    for (const [fonts, glyphs] of [[exactFonts, "Do the following"], [fallbackFonts, "DO THE FOLLOWING"]]) {
      painted.length = 0;
      const renderer = new SceneRenderer([object], DEFAULT_LIMITS, fonts);
      (await renderer.render(unit, { unitIndex: 0 })).bitmap.close();
      assert.equal(painted.map(({ text }) => text).join(""), glyphs);
    }
  } finally {
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("inline documents register local fonts and remove them on close", async () => {
  const fontSet = new FakeFontSet();
  const restore = installFontRuntime(fontSet);
  try {
    const engine = await createOfficeEngine({
      execution: "inline",
      wasm,
      fonts: [{ family: "Inline Sans", bytes: Uint8Array.of(1, 2, 3) }],
    });
    const document = await engine.open(odt);
    assert.equal(fontSet.faces.size, 1);
    document.close();
    assert.equal(fontSet.faces.size, 0);
    assert.equal(fontSet.deleted.length, 1);
    engine.close();
  } finally {
    restore();
  }
});

test("document font runs expose script-specific fallback faces for mixed text", () => {
  const object = {
    numericId: 1,
    id: "mixed-fallback",
    type: "paragraph",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 200, height: 40 },
    source: { format: "docx", part: "word/document.xml", kind: "paragraph", mapping: "derived" },
    z: 0,
    text: "1.控制器名称： BMS",
    visual: {
      kind: "rich-text",
      geometry: "rectangle",
      fill: { kind: "none" },
      stroke: { kind: "none" },
      strokeWidth: 0,
      align: "start",
      lineHeight: 16,
      runs: [{
        text: "1.控制器名称： BMS",
        fontFamily: "等线",
        fontSize: 12,
        color: 0xff,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        highlight: 0,
        baselineShift: 0,
        letterSpacing: 0,
      }],
    },
  };
  const fallback = {
    resolve: () => ({ family: "Calibri", source: "fallback" }),
    resolveFace(face) {
      const hasCjk = face.codePoints?.some((codePoint) => codePoint >= 0x3400 && codePoint <= 0x9fff);
      const hasLatin = face.codePoints?.some((codePoint) => codePoint >= 0x30 && codePoint <= 0x7a);
      return {
        ...face,
        family: hasCjk && !hasLatin ? "Hiragino Sans" : "Calibri",
        source: "fallback",
      };
    },
  };

  assert.deepEqual(documentFontRuns(object, fallback), [
    { start: 0, end: 2, authoredFamily: "等线", renderedFamily: "Calibri", source: "fallback" },
    { start: 2, end: 9, authoredFamily: "等线", renderedFamily: "Hiragino Sans", source: "fallback" },
    { start: 9, end: 12, authoredFamily: "等线", renderedFamily: "Calibri", source: "fallback" },
  ]);
});

test("DOCX embedded fonts are deobfuscated, registered, and document-scoped", async () => {
  const fontSet = new FakeFontSet();
  const restore = installFontRuntime(fontSet);
  try {
    const engine = await createOfficeEngine({ execution: "inline", wasm });
    const document = await engine.open(docxWithEmbeddedFont);
    const [face] = [...fontSet.faces];
    const object = (await document.listObjects({ textOnly: true }))
      .find((candidate) => candidate.text === "Embedded");

    assert.match(face.family, /^OfficeViewer \d+ embedded 0$/u);
    assert.deepEqual([...new Uint8Array(face.source)], [...embeddedFontBytes]);
    assert.deepEqual(object?.fontRuns, [{
      start: 0,
      end: 8,
      authoredFamily: "Embedded Sans",
      renderedFamily: face.family,
      source: "embedded",
    }]);
    assert.equal(Object.isFrozen(object?.fontRuns), true);
    assert.equal(Object.isFrozen(object?.fontRuns?.[0]), true);
    document.close();
    assert.equal(fontSet.faces.size, 0);
    engine.close();
  } finally {
    restore();
  }
});

test("PPTX and OpenDocument packaged fonts use the same embedded-first path", async () => {
  for (const bytes of [pptxWithEmbeddedFont, odtWithEmbeddedFont]) {
    const fontSet = new FakeFontSet();
    const restore = installFontRuntime(fontSet);
    try {
      const engine = await createOfficeEngine({ execution: "inline", wasm });
      const document = await engine.open(bytes);
      const embedded = [...fontSet.faces].find((face) => /^OfficeViewer \d+ embedded 0$/u.test(face.family));

      assert.ok(embedded);
      assert.deepEqual([...new Uint8Array(embedded.source)], [...embeddedFontBytes]);
      document.close();
      assert.equal(fontSet.faces.size, 0);
      engine.close();
    } finally {
      restore();
    }
  }
});

test("host and embedded fonts share one document font-byte budget", async () => {
  const fontSet = new FakeFontSet();
  const restore = installFontRuntime(fontSet);
  try {
    const engine = await createOfficeEngine({
      execution: "inline",
      wasm,
      limits: { fontBytes: embeddedFontBytes.length },
      fonts: [{ family: "Host Budget", bytes: Uint8Array.of(1) }],
    });
    const document = await engine.open(docxWithEmbeddedFont);

    assert.equal(document.diagnostics().some(({ code }) => code === "FONT_BYTES_LIMIT"), true);
    assert.equal([...fontSet.faces].some((face) => /^OfficeViewer \d+ embedded 0$/u.test(face.family)), false);
    document.close();
    engine.close();
  } finally {
    restore();
  }
});

test("closing an inline engine while fonts load closes the pending document and registration", async () => {
  let finishLoading;
  class PendingFontFace extends LoadedFontFace {
    load() {
      if (typeof this.source === "string") return Promise.reject(new Error("local font unavailable"));
      return new Promise((resolve) => {
        finishLoading = () => resolve(this);
      });
    }
  }
  const fontSet = new FakeFontSet();
  const restore = installFontRuntime(fontSet, PendingFontFace);
  try {
    const engine = await createOfficeEngine({
      execution: "inline",
      wasm: odfWasm,
      formatPack: false,
      fonts: [{ family: "Pending Sans", bytes: Uint8Array.of(1, 2, 3) }],
    });
    const opening = engine.open(odt);
    assert.equal(fontSet.faces.size, 1);

    engine.close();
    finishLoading();

    await assert.rejects(opening, (error) => error?.code === "ENGINE_CLOSED");
    assert.equal(fontSet.faces.size, 0);
  } finally {
    restore();
  }
});

class FakeWorker {
  constructor() {
    this.onmessage = null;
    this.onerror = null;
    this.onmessageerror = null;
    this.messages = [];
    this.transfers = [];
    this.terminated = false;
  }

  postMessage(message, transfer) {
    this.messages.push(message);
    this.transfers.push(transfer);
  }

  terminate() {
    this.terminated = true;
  }

  respondToOpen() {
    const request = this.messages.findLast((message) => message.type === "open");
    queueMicrotask(() => this.onmessage?.({
      data: {
        id: request.id,
        ok: true,
        type: "open",
        snapshot: {
          info: { format: "odt", kind: "text", units: [] },
          diagnostics: [],
          objects: [],
        },
      },
    }));
  }
}

test("worker execution transfers a distinct font copy to every document", async () => {
  const workers = [];
  const source = Uint8Array.of(7, 8, 9);
  const engine = await createOfficeEngine({
    execution: "worker",
    wasm,
    fonts: [{ family: "Worker Sans", bytes: source }],
    workerFactory: () => {
      const worker = new FakeWorker();
      workers.push(worker);
      return worker;
    },
  });

  const firstOpening = engine.open(Uint8Array.of(1));
  workers[0].respondToOpen();
  const first = await firstOpening;
  first.close();

  source[0] = 99;
  const secondOpening = engine.open(Uint8Array.of(1));
  workers[1].respondToOpen();
  const second = await secondOpening;

  const firstFont = workers[0].messages[0].fonts[0];
  const secondFont = workers[1].messages[0].fonts[0];
  assert.notEqual(firstFont.bytes, secondFont.bytes);
  assert.deepEqual([...new Uint8Array(firstFont.bytes)], [7, 8, 9]);
  assert.deepEqual([...new Uint8Array(secondFont.bytes)], [7, 8, 9]);
  assert.equal(workers[0].transfers[0].includes(firstFont.bytes), true);
  assert.equal(workers[1].transfers[0].includes(secondFont.bytes), true);
  assert.equal(workers[0].terminated, true);

  second.close();
  engine.close();
  assert.equal(workers[1].terminated, true);
});

test("supplied Word first-page Chinese fonts use loadable macOS face names", async () => {
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(await readFile(new URL("./fixtures/word-first-page-text.docx", import.meta.url)));
  const requests = (await document.listObjects({ unitIndex: 0 })).flatMap((object) =>
    (object.fontRuns ?? []).map((run) => ({
      family: run.authoredFamily, style: "normal", weight: 400, stretch: "normal",
      codePoints: [...new Set([...object.text.slice(run.start, run.end)].map((character) => character.codePointAt(0)))],
    })));
  const families = ["SimHei", "等线"];
  class MacFace extends LoadedFontFace {
    async load() {
      if (!['local("STHeitiSC-Medium")', 'local("PingFangSC-Regular")'].includes(this.source)) {
        throw new Error("macOS local() requires a face name, not this family alias");
      }
      return this;
    }
  }
  const registration = await registerFonts([], 50, { FontFace: MacFace, fontSet: new FakeFontSet() }, undefined, [], [], {
    policy: "local-first",
    requestedFaces: requests.filter((face) => families.includes(face.family)),
  });
  try {
    for (const family of families) assert.equal(registration.resolve(family).source, "browser", family);
  } finally {
    registration.close();
    await document.close();
  }
});

test("supplied PAAK Sheet2 uses a condensed local fallback for missing Roboto Condensed", async () => {
  class ArialNarrowOnly extends LoadedFontFace {
    async load() {
      if (this.source === 'local("Arial Narrow")') return this;
      throw new Error("local font unavailable");
    }
  }
  const fontSet = new FakeFontSet();
  const restore = installFontRuntime(fontSet, ArialNarrowOnly);
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/China PAAK EPC evaluation.xlsx", import.meta.url)));
    const cell = (await document.listObjects({ unitIndex: 1, textOnly: true }))
      .find((object) => object.source.address === "B3");
    assert.equal(cell?.text?.startsWith("RKE use-cases"), true);
    assert.equal(cell.fontRuns[0].authoredFamily, "Roboto Condensed");
    assert.equal(cell.fontRuns[0].source, "browser");
    assert.ok([...fontSet.faces].some((face) => face.source === 'local("Arial Narrow")'));
  } finally {
    document?.close();
    engine.close();
    restore();
  }
});

test("supplied Cambria DOCX uses pinned metric-compatible fallback instead of Georgia", async () => {
  const core = await Core.create(wasm, DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("fixtures/complex2005_12rtm.docx", import.meta.url)));
  const requests = collectFontRequests(doc.scene.objects).filter((face) => face.family === "Cambria");
  assert.ok(requests.length > 0);
  const originalFetch = globalThis.fetch;
  const urls = [];
  globalThis.fetch = async (url) => {
    urls.push(url);
    return new Response(await readFile(url));
  };
  class GeorgiaOnly extends LoadedFontFace {
    async load() {
      if (this.source === 'local("Georgia")') return this;
      return super.load();
    }
  }
  try {
    const fontSet = new FakeFontSet();
    const registration = await registerFonts([], 1000, { FontFace: GeorgiaOnly, fontSet },
      undefined, [], [], { policy: "local-first", requestedFaces: requests });
    for (const request of requests) {
      assert.equal(registration.resolveFace(request).source, "fallback");
      assert.match(registration.resolveFace(request).family, /Caladea/);
      assert.equal(registration.hasExactFace(request), false, "provider preflight must still request real Cambria");
    }
    const { createHash } = await import("node:crypto");
    const regular = [...fontSet.faces].find((face) => face.descriptors.weight === "400" && face.descriptors.style === "normal");
    assert.equal(createHash("sha256").update(new Uint8Array(regular.source)).digest("hex"),
      "d2f6cad33f191e65b68bd74e6d4f7708080a41b32db635866109df3090265d91",
      "newer Caladea fonts change the real-file line widths");
    assert.equal(urls.length, requests.length);
    assert.ok([...fontSet.faces].every((face) => face.source instanceof ArrayBuffer));
    assert.equal(registration.diagnostics().length, 0);
    registration.close();
    assert.equal(fontSet.faces.size, 0);
    urls.length = 0;
    const supplied = await registerFonts(prepareFontAssets([{ family: "Cambria", bytes: new Uint8Array([1, 2]) }]),
      1000, { FontFace: GeorgiaOnly, fontSet }, undefined, [], ["Cambria"]);
    assert.equal(supplied.resolve("Cambria").source, "host");
    assert.equal(urls.length, 0, "available authored fonts must not fetch fallback assets");
    supplied.close();
    const provided = await registerFonts([], 1000, { FontFace: GeorgiaOnly, fontSet },
      undefined, [], ["Cambria"], { providerAssets: prepareFontAssets([{ family: "Cambria", bytes: new Uint8Array([1, 2]) }]) });
    assert.equal(provided.resolve("Cambria").source, "provider");
    assert.equal(urls.length, 0);
    provided.close();
    globalThis.fetch = async () => { throw new Error("asset unavailable"); };
    const missing = await registerFonts([], 1000, { FontFace: GeorgiaOnly, fontSet },
      undefined, [], ["Cambria"]);
    assert.equal(missing.resolve("Cambria").source, "fallback");
    assert.equal(missing.diagnostics()[0].code, "FONT_LOAD_FAILED");
    missing.close();
    assert.equal(fontSet.faces.size, 0);

  } finally {
    globalThis.fetch = originalFetch;
    doc.close();
    core.close();
  }
});

test("supplied bnc910045 table uses Carlito when Calibri is missing", async () => {
  const core = await Core.create(wasm, DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("fixtures/bnc910045.pptx", import.meta.url)));
  const requests = collectFontRequests(doc.scene.objects).filter((face) => face.family === "Calibri");
  assert.ok(requests.length > 0);
  const originalFetch = globalThis.fetch;
  const urls = [];
  globalThis.fetch = async (url) => {
    urls.push(url);
    return new Response(await readFile(url));
  };
  class GeorgiaOnly extends LoadedFontFace {
    async load() {
      if (this.source === 'local("Georgia")') return this;
      return super.load();
    }
  }
  try {
    const fontSet = new FakeFontSet();
    const registration = await registerFonts([], 1000, { FontFace: GeorgiaOnly, fontSet },
      undefined, [], [], { policy: "local-first", requestedFaces: requests });
    for (const request of requests) {
      assert.equal(registration.resolveFace(request).source, "fallback");
      assert.match(registration.resolveFace(request).family, /Carlito/);
      assert.equal(registration.hasExactFace(request), false, "provider preflight must still request real Calibri");
    }
    assert.equal(urls.length, requests.length);
    assert.ok([...fontSet.faces].every((face) => face.source instanceof ArrayBuffer));
    assert.equal(registration.diagnostics().length, 0);
    registration.close();
    assert.equal(fontSet.faces.size, 0);
    urls.length = 0;
    const supplied = await registerFonts(prepareFontAssets([{ family: "Calibri", bytes: new Uint8Array([1, 2]) }]),
      1000, { FontFace: GeorgiaOnly, fontSet }, undefined, [], ["Calibri"]);
    assert.equal(supplied.resolve("Calibri").source, "host");
    assert.equal(urls.length, 0, "available authored fonts must not fetch fallback assets");
    supplied.close();
    const provided = await registerFonts([], 1000, { FontFace: GeorgiaOnly, fontSet },
      undefined, [], ["Calibri"], { providerAssets: prepareFontAssets([{ family: "Calibri", bytes: new Uint8Array([1, 2]) }]) });
    assert.equal(provided.resolve("Calibri").source, "provider");
    assert.equal(urls.length, 0);
    provided.close();
    globalThis.fetch = async () => { throw new Error("asset unavailable"); };
    const missing = await registerFonts([], 1000, { FontFace: GeorgiaOnly, fontSet },
      undefined, [], ["Calibri"]);
    assert.equal(missing.resolve("Calibri").source, "fallback");
    assert.equal(missing.diagnostics()[0].code, "FONT_LOAD_FAILED");
    missing.close();
    assert.equal(fontSet.faces.size, 0);

  } finally {
    globalThis.fetch = originalFetch;
    doc.close();
    core.close();
  }
});
