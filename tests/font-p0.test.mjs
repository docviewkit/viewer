import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";
import { createZip } from "./zip-fixture.mjs";

const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));

function docxWithText(text, family = "P0 Sans", { pageWidth = 12240, pageHeight = 15840, margin = 1440 } = {}) {
  return createZip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>`,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>
          <w:sectPr><w:pgSz w:w="${pageWidth}" w:h="${pageHeight}"/><w:pgMar w:top="${margin}" w:right="${margin}" w:bottom="${margin}" w:left="${margin}"/></w:sectPr>
        </w:body>
      </w:document>`,
  }, { compress: false });
}

function rtfWithText(text, family = "RTF Metrics Sans") {
  return new TextEncoder().encode(String.raw`{\rtf1\ansi\deff0
{\fonttbl{\f0 ${family};}}
\paperw2880\paperh15840\margl360\margr360\margt360\margb360
\pard\f0\fs24 ${text}\par}`);
}

function pptxWithTableText(text, family = "PPTX Metrics Sans") {
  return createZip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
        <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
      </Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
      </Relationships>`,
    "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
        <p:sldSz cx="9144000" cy="6858000"/>
      </p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
      </Relationships>`,
    "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="7"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
              <a:tblGrid><a:gridCol w="762000"/></a:tblGrid>
              <a:tr h="0"><a:tc><a:txBody><a:p><a:r>
                <a:rPr sz="1200"><a:latin typeface="${family}"/></a:rPr><a:t>${text}</a:t>
              </a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
              <a:tr h="0"><a:tc><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld>
      </p:sld>`,
  }, { compress: false });
}

test("open requests the exact document face and Unicode scalars from the host font provider", async () => {
  const calls = [];
  const engine = await createOfficeEngine({
    execution: "inline",
    wasm,
    fontPolicy: "deterministic",
    fontProvider: async (requests, context) => {
      calls.push({ requests, context });
      return [];
    },
  });

  const document = await engine.open(docxWithText("A中🙂A"));

  assert.equal(calls.length, 1);
  assert.equal(calls[0].context.policy, "deterministic");
  assert.deepEqual(calls[0].requests, [{
    family: "P0 Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [65, 0x4e2d, 0x1f642],
  }]);

  document.close();
  engine.close();
});

test("engine initialization rejects a static font whose manifest digest does not match", async () => {
  await assert.rejects(
    createOfficeEngine({
      execution: "inline",
      wasm,
      fonts: [{
        family: "Pinned Static Sans",
        bytes: Uint8Array.of(1, 2, 3),
        sha256: "cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc",
      }],
    }),
    (error) => error?.code === "FONT_INTEGRITY_MISMATCH",
  );
});

class TestFontSet {
  constructor() {
    this.faces = new Set();
  }

  add(face) {
    this.faces.add(face);
  }

  delete(face) {
    return this.faces.delete(face);
  }

  check() {
    return true;
  }

  forEach(callback) {
    this.faces.forEach(callback);
  }
}

class TestFontFace {
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

class TestMeasureContext {
  font = "400 16px sans-serif";
  letterSpacing = "0px";

  measureText(text) {
    const size = Number.parseFloat(/([0-9.]+)px/u.exec(this.font)?.[1] ?? "16");
    const advanceEm = this.font.includes("OfficeViewer") ? 1 : 0.55;
    return { width: Array.from(text).length * size * advanceEm };
  }
}

class TestOffscreenCanvas {
  getContext(kind) {
    return kind === "2d" ? new TestMeasureContext() : null;
  }
}

async function withFontRuntime(run, FontFace = TestFontFace) {
  const descriptors = ["FontFace", "fonts", "OffscreenCanvas"].map((name) => [
    name,
    Object.getOwnPropertyDescriptor(globalThis, name),
  ]);
  Object.defineProperties(globalThis, {
    FontFace: { configurable: true, value: FontFace },
    fonts: { configurable: true, value: new TestFontSet() },
    OffscreenCanvas: { configurable: true, value: TestOffscreenCanvas },
  });
  try {
    return await run();
  } finally {
    for (const [name, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor);
      else delete globalThis[name];
    }
  }
}

class AvailableLocalFontFace extends TestFontFace {
  async load() {
    return this;
  }
}

test("the default local-first policy avoids the provider when the exact browser face is available", async () => {
  await withFontRuntime(async () => {
    let providerCalls = 0;
    const engine = await createOfficeEngine({
      execution: "inline",
      wasm,
      fontProvider: async () => {
        providerCalls += 1;
        return [{ family: "Local First Sans", bytes: Uint8Array.of(0, 1, 0, 0, 1) }];
      },
    });
    const document = await engine.open(docxWithText("local", "Local First Sans"));
    const [paragraph] = await document.listObjects({ textOnly: true });

    assert.equal(providerCalls, 0);
    assert.equal(paragraph.fontRuns[0].source, "browser");

    document.close();
    engine.close();
  }, AvailableLocalFontFace);
});

test("deterministic policy prefers the provider over the same exact browser face", async () => {
  await withFontRuntime(async () => {
    let providerCalls = 0;
    const engine = await createOfficeEngine({
      execution: "inline",
      wasm,
      fontPolicy: "deterministic",
      fontProvider: async () => {
        providerCalls += 1;
        return [{ family: "Deterministic Sans", bytes: Uint8Array.of(0, 1, 0, 0, 1) }];
      },
    });
    const document = await engine.open(docxWithText("stable", "Deterministic Sans"));
    const [paragraph] = await document.listObjects({ textOnly: true });

    assert.equal(providerCalls, 1);
    assert.equal(paragraph.fontRuns[0].source, "provider");

    document.close();
    engine.close();
  }, AvailableLocalFontFace);
});

test("resolved browser font advances participate in final DOCX pagination", async () => {
  await withFontRuntime(async () => {
    const input = docxWithText("i".repeat(800), "Pagination Sans", {
      pageWidth: 2880,
      pageHeight: 2880,
      margin: 360,
    });
    const fallbackEngine = await createOfficeEngine({ execution: "inline", wasm });
    const fallback = await fallbackEngine.open(input);
    const fallbackPages = fallback.info.units.length;
    fallback.close();
    fallbackEngine.close();

    const fontEngine = await createOfficeEngine({
      execution: "inline",
      wasm,
      fontProvider: async () => [{
        family: "Pagination Sans",
        bytes: Uint8Array.of(0, 1, 0, 0, 1),
      }],
    });
    const measured = await fontEngine.open(input);

    assert.ok(measured.info.units.length > fallbackPages, `${measured.info.units.length} should exceed ${fallbackPages}`);
    assert.ok(measured.diagnostics().some(({ code }) => code === "FONT_METRICS_APPLIED"));

    measured.close();
    fontEngine.close();
  });
});

test("resolved browser font advances participate in final RTF paragraph layout", async () => {
  await withFontRuntime(async () => {
    const input = rtfWithText("i ".repeat(100));
    const fallbackEngine = await createOfficeEngine({ execution: "inline", wasm });
    const fallback = await fallbackEngine.open(input);
    const [fallbackParagraph] = await fallback.listObjects({ textOnly: true });
    const fallbackHeight = fallbackParagraph.bounds.height;
    fallback.close();
    fallbackEngine.close();

    const fontEngine = await createOfficeEngine({
      execution: "inline",
      wasm,
      fontProvider: async () => [{
        family: "RTF Metrics Sans",
        bytes: Uint8Array.of(0, 1, 0, 0, 1),
      }],
    });
    const measured = await fontEngine.open(input);
    const [measuredParagraph] = await measured.listObjects({ textOnly: true });

    assert.ok(measuredParagraph.bounds.height > fallbackHeight, `${measuredParagraph.bounds.height} should exceed ${fallbackHeight}`);
    assert.ok(measured.diagnostics().some(({ code }) => code === "FONT_METRICS_APPLIED"));

    measured.close();
    fontEngine.close();
  });
});

test("resolved browser font advances participate in final PPTX table layout", async () => {
  await withFontRuntime(async () => {
    const input = pptxWithTableText("ii ii ii");
    const fallbackEngine = await createOfficeEngine({ execution: "inline", wasm });
    const fallback = await fallbackEngine.open(input);
    const fallbackCells = await fallback.listObjects({ unitIndex: 0 });
    const fallbackHeight = fallbackCells.find(({ type }) => type === "cell")?.bounds.height;

    const fontEngine = await createOfficeEngine({
      execution: "inline",
      wasm,
      fontPolicy: "deterministic",
      fontProvider: async () => [{
        family: "PPTX Metrics Sans",
        bytes: Uint8Array.of(0, 1, 0, 0, 1),
      }],
    });
    const measured = await fontEngine.open(input);
    const measuredCells = await measured.listObjects({ unitIndex: 0 });
    const measuredHeight = measuredCells.find(({ type }) => type === "cell")?.bounds.height;

    assert.equal(typeof fallbackHeight, "number");
    assert.equal(typeof measuredHeight, "number");
    assert.ok(measuredHeight > fallbackHeight, `${measuredHeight} should exceed ${fallbackHeight}`);
    assert.ok(measured.diagnostics().some(({ code }) => code === "FONT_METRICS_APPLIED"));

    measured.close();
    fontEngine.close();
    fallback.close();
    fallbackEngine.close();
  });
});
