import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine as createOfficeEngineBase } from "../dist/engine.js";
import { extendedFormatPack } from "../dist/extended-formats.js";
import { createZip } from "./zip-fixture.mjs";
import { readZipEntries } from "../scripts/accuracy-metamorphic.mjs";

test("real 3D column chart projects joined faces and paints distant series first", async () => {
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("fixtures/chart-3d-column.pptx", import.meta.url)));
  try {
    doc.loadUnit(0);
    const faces = doc.scene.objects.filter(o => o.source.part === "ppt/charts/chart1.xml" &&
      o.source.row !== undefined && o.source.column !== undefined && o.type === "shape");
    assert.equal(faces.length, 36, "12 columns each have three visible faces");
    for (let row = 0; row < 3; row++) {
      for (let column = 0; column < 4; column++) {
        const group = faces.filter(o => o.source.row === row && o.source.column === column);
        assert.equal(group.length, 3);
        const vertices = group.map(o => {
          const geometry = o.visual.geometry;
          assert.equal(geometry.kind, "path", "front faces must also be projected");
          return geometry.commands.filter(c => "x" in c).map(c => [o.bounds.x + c.x, o.bounds.y + c.y]);
        });
        assert.ok(Math.abs(vertices[0][0][1] - vertices[0][1][1]) > 0.1, "floor slope");
        for (const side of vertices.slice(1)) {
          const shared = side.filter(p => vertices[0].some(q => Math.hypot(p[0]-q[0], p[1]-q[1]) < 0.001));
          assert.equal(shared.length, 2, "adjacent faces must share an exact edge");
        }
      }
    }
    assert.ok(Math.max(...faces.filter(o => o.source.row === 2).map(o => o.z)) <
      Math.min(...faces.filter(o => o.source.row === 1).map(o => o.z)));
    assert.ok(Math.max(...faces.filter(o => o.source.row === 1).map(o => o.z)) <
      Math.min(...faces.filter(o => o.source.row === 0).map(o => o.z)));
  } finally { doc.close(); core.close(); }
});

test("real PPTX OMML carrier renders a real DOCX fraction as stacked math", async () => {
  const source = await readFile(new URL("fixtures/ooxml-omml-slide.pptx", import.meta.url));
  const parts = Object.fromEntries(readZipEntries(source).map(({name, data}) => [name, data]));
  const xml = new TextDecoder().decode(parts["ppt/slides/slide1.xml"]);
  const start = xml.indexOf("<m:oMath xmlns:m=");
  const contentStart = xml.indexOf(">", start) + 1;
  const end = xml.indexOf("</m:oMath>", contentStart);
  assert.ok(start >= 0 && end > contentStart);
  // Exact m:f from LibreOffice's math-vertical_stacks.docx, carried by the real PPTX math shape.
  const fraction = '<m:f><m:num><m:r><m:t xml:space="preserve">a</m:t></m:r></m:num><m:den><m:r><m:t xml:space="preserve">b</m:t></m:r></m:den></m:f>';
  parts["ppt/slides/slide1.xml"] = xml.slice(0, contentStart) + fraction + xml.slice(end);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({wasm, execution: "inline"});
  const original = await engine.open(source);
  let document;
  try {
    assert.ok((await original.listObjects()).some(({source, text}) => source.shapeId === 4 && text === "𝜕"));
    document = await engine.open(createZip(parts));
    const objects = (await document.listObjects()).filter(({source}) => source.shapeId === 4);
    const numerator = objects.find(({text}) => text === "a");
    const denominator = objects.find(({text}) => text === "b");
    assert.ok(numerator && denominator, "fraction needs separate visible numerator and denominator");
    assert.ok(numerator.bounds.y < denominator.bounds.y, "numerator must be above denominator");
  } finally { original.close(); document?.close(); engine.close(); }
});

const drawingMlPresetShapes = Object.freeze([
  "accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3", "accentCallout1", "accentCallout2", "accentCallout3",
  "actionButtonBackPrevious", "actionButtonBeginning", "actionButtonBlank", "actionButtonDocument", "actionButtonEnd", "actionButtonForwardNext",
  "actionButtonHelp", "actionButtonHome", "actionButtonInformation", "actionButtonMovie", "actionButtonReturn", "actionButtonSound",
  "arc", "bentArrow", "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5",
  "bentUpArrow", "bevel", "blockArc", "borderCallout1", "borderCallout2", "borderCallout3",
  "bracePair", "bracketPair", "callout1", "callout2", "callout3", "can",
  "chartPlus", "chartStar", "chartX", "chevron", "chord", "circularArrow",
  "cloud", "cloudCallout", "corner", "cornerTabs", "cube", "curvedConnector2",
  "curvedConnector3", "curvedConnector4", "curvedConnector5", "curvedDownArrow", "curvedLeftArrow", "curvedRightArrow",
  "curvedUpArrow", "decagon", "diagStripe", "diamond", "dodecagon", "donut",
  "doubleWave", "downArrow", "downArrowCallout", "ellipse", "ellipseRibbon", "ellipseRibbon2",
  "flowChartAlternateProcess", "flowChartCollate", "flowChartConnector", "flowChartDecision", "flowChartDelay", "flowChartDisplay",
  "flowChartDocument", "flowChartExtract", "flowChartInputOutput", "flowChartInternalStorage", "flowChartMagneticDisk", "flowChartMagneticDrum",
  "flowChartMagneticTape", "flowChartManualInput", "flowChartManualOperation", "flowChartMerge", "flowChartMultidocument", "flowChartOfflineStorage",
  "flowChartOffpageConnector", "flowChartOnlineStorage", "flowChartOr", "flowChartPredefinedProcess", "flowChartPreparation", "flowChartProcess",
  "flowChartPunchedCard", "flowChartPunchedTape", "flowChartSort", "flowChartSummingJunction", "flowChartTerminator", "foldedCorner",
  "frame", "funnel", "gear6", "gear9", "halfFrame", "heart",
  "heptagon", "hexagon", "homePlate", "horizontalScroll", "irregularSeal1", "irregularSeal2",
  "leftArrow", "leftArrowCallout", "leftBrace", "leftBracket", "leftCircularArrow", "leftRightArrow",
  "leftRightArrowCallout", "leftRightCircularArrow", "leftRightRibbon", "leftRightUpArrow", "leftUpArrow", "lightningBolt",
  "line", "lineInv", "mathDivide", "mathEqual", "mathMinus", "mathMultiply",
  "mathNotEqual", "mathPlus", "moon", "nonIsoscelesTrapezoid", "noSmoking", "notchedRightArrow",
  "octagon", "parallelogram", "pentagon", "pie", "pieWedge", "plaque",
  "plaqueTabs", "plus", "quadArrow", "quadArrowCallout", "rect", "ribbon",
  "ribbon2", "rightArrow", "rightArrowCallout", "rightBrace", "rightBracket", "round1Rect",
  "round2DiagRect", "round2SameRect", "roundRect", "rtTriangle", "smileyFace", "snip1Rect",
  "snip2DiagRect", "snip2SameRect", "snipRoundRect", "squareTabs", "star10", "star12",
  "star16", "star24", "star32", "star4", "star5", "star6",
  "star7", "star8", "straightConnector1", "stripedRightArrow", "sun", "swooshArrow",
  "teardrop", "trapezoid", "triangle", "upArrowCallout", "upDownArrow", "upArrow",
  "upDownArrowCallout", "uturnArrow", "verticalScroll", "wave", "wedgeEllipseCallout", "wedgeRectCallout",
  "wedgeRoundRectCallout",
]);

function createOfficeEngine(options) {
  return createOfficeEngineBase({
    fontPolicy: "local-first",
    formatPack: async () => ({
      async load(candidate) {
        return readFile(await extendedFormatPack.load(candidate));
      },
    }),
    ...options,
  });
}

const onePixelPng = Uint8Array.of(
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2,
  0, 0, 0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1,
  5, 1, 1, 39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
);

function pptxWithSlide(slide, extraEntries = {}) {
  return createZip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
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
    "ppt/slides/slide1.xml": slide,
    ...extraEntries,
  });
}

function odpWithContent(automaticStyles, pageContent, masterContent = "", extraEntries = {}) {
  return createZip({
    mimetype: "application/vnd.oasis.opendocument.presentation",
    "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
        <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>
        <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
        <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
      </manifest:manifest>`,
    "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
        <office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/></style:page-layout></office:automatic-styles>
        <office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1">${masterContent}</style:master-page></office:master-styles>
      </office:document-styles>`,
    "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">
        <office:automatic-styles>${automaticStyles}</office:automatic-styles>
        <office:body><office:presentation><draw:page draw:name="Slide 1" draw:master-page-name="Default" draw:style-name="page1">${pageContent}</draw:page></office:presentation></office:body>
      </office:document-content>`,
    ...extraEntries,
  }, { compress: false });
}

test("renders a solid PPTX slide background as a source-mapped object", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld>
        <p:bg><p:bgPr><a:solidFill><a:srgbClr val="102030"/></a:solidFill></p:bgPr></p:bg>
        <p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree>
      </p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 10, y: 10 });

    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.type, "shape");
    assert.deepEqual(hits[0].object.bounds, { x: 0, y: 0, width: 960, height: 720 });
    assert.equal(hits[0].object.source.part, "ppt/slides/slide1.xml");
    assert.equal(hits[0].object.source.mapping, "derived");
    assert.equal(document.diagnostics().some(({ fidelity }) => fidelity === "not-rendered"), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("exposes PPTX slide and object metadata, actions, visibility, and unsupported features", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      show="0">
      <p:cSld name="Board Update"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Visible card"><a:hlinkHover action="ppaction://macro?name=RunReport"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
            <a:scene3d/>
          </p:spPr>
        </p:sp>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="3" name="Hidden action" title="Open details" descr="Accessible details" hidden="1">
              <a:hlinkClick r:id="rIdLink" tooltip="Open example"/>
              <a:hlinkHover action="ppaction://hlinkshowjump?jump=nextslide" tooltip="Next slide"/>
            </p:cNvPr>
            <p:cNvSpPr/><p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="2857500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
      <p:transition/>
      <p:timing/>
    </p:sld>`, {
    "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        firstSlideNum="5">
        <p:sldIdLst><p:sldId id="512" r:id="rId1"/></p:sldIdLst>
        <p:sldSz cx="9144000" cy="6858000"/>
      </p:presentation>`,
    "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
          Target="https://example.com/details" TargetMode="External"/>
        <Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
        <Relationship Id="rIdVideo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/video1.mp4"/>
        <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/object1.bin"/>
        <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>
      </Relationships>`,
    "ppt/notesSlides/notesSlide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree>
          <p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
            <p:txBody><a:bodyPr/><a:lstStyle/>
              <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="C62828"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>Opening point</a:t></a:r></a:p>
              <a:p><a:pPr marL="342900" indent="-171450"><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="1200" i="1" u="sng"/><a:t>第二条备注</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
          <p:sp><p:nvSpPr><p:cNvPr id="3"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>5</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld>
      </p:notes>`,
    "ppt/notesSlides/_rels/notesSlide1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdNotesMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
      </Relationships>`,
    "ppt/notesMasters/notesMaster1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:notesStyle><a:lvl1pPr><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="1565C0"/></a:solidFill><a:latin typeface="Courier New"/></a:defRPr></a:lvl1pPr></p:notesStyle>
      </p:notesMaster>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    set fillStyle(value) { fills.push(value); },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const [unit] = document.info.units;
    const objects = await document.listObjects();
    const hidden = objects.find(({ source }) => source.format === "pptx" && source.shapeId === 3);
    const hits = await document.hitTest({ unitIndex: 0, x: 350, y: 150 });
    frame = await document.render({ unitIndex: 0 });

    const { speakerNoteParagraphs, ...plainUnit } = unit;
    assert.deepEqual(plainUnit, {
      type: "slide",
      index: 0,
      id: "unit:0",
      name: "Board Update",
      width: 960,
      height: 720,
      sourceId: "512",
      sourcePart: "ppt/slides/slide1.xml",
      slideNumber: 5,
      hidden: true,
      speakerNotes: "Opening point\n第二条备注",
      speakerNotesPart: "ppt/notesSlides/notesSlide1.xml",
    });
    assert.equal(speakerNoteParagraphs?.length, 2);
    assert.equal(speakerNoteParagraphs?.[0]?.align, "center");
    assert.equal(speakerNoteParagraphs?.[0]?.runs[0]?.text, "Opening point");
    assert.equal(speakerNoteParagraphs?.[0]?.runs[0]?.fontFamily, "Aptos");
    assert.equal(speakerNoteParagraphs?.[0]?.runs[0]?.fontSize, 24);
    assert.equal(speakerNoteParagraphs?.[0]?.runs[0]?.color, 0xc62828ff);
    assert.equal(speakerNoteParagraphs?.[0]?.runs[0]?.bold, true);
    assert.equal(speakerNoteParagraphs?.[1]?.marginLeft, 36);
    assert.equal(speakerNoteParagraphs?.[1]?.firstLineIndent, -18);
    assert.equal(speakerNoteParagraphs?.[1]?.runs[0]?.text, "•\t");
    assert.equal(speakerNoteParagraphs?.[1]?.runs[1]?.fontFamily, "Courier New");
    assert.equal(speakerNoteParagraphs?.[1]?.runs[1]?.color, 0x1565c0ff);
    assert.equal(speakerNoteParagraphs?.[1]?.runs[1]?.italic, true);
    assert.equal(speakerNoteParagraphs?.[1]?.runs[1]?.underline, true);
    assert.equal(hidden?.name, "Hidden action");
    assert.equal(hidden?.title, "Open details");
    assert.equal(hidden?.description, "Accessible details");
    assert.equal(hidden?.hidden, true);
    assert.deepEqual(hidden?.actions, [
      {
        trigger: "click",
        kind: "hyperlink",
        target: "https://example.com/details",
        tooltip: "Open example",
      },
      {
        trigger: "hover",
        kind: "slide",
        action: "ppaction://hlinkshowjump?jump=nextslide",
        target: "next",
        tooltip: "Next slide",
      },
    ]);
    assert.equal(hits.some(({ object }) => object.source.shapeId === 3), false);
    assert.equal(frame.renderedObjectCount, 1);
    assert.equal(fills.includes("rgba(255, 0, 0, 1)"), false);
    assert.ok(document.diagnostics().some(({ code, details }) => (
      code === "UNSUPPORTED_FEATURE" && details?.feature === "slide-transition"
    )));
    assert.ok(document.diagnostics().some(({ code, details }) => (
      code === "UNSUPPORTED_FEATURE" && details?.feature === "animation-timing"
    )));
    for (const feature of [
      "media-playback",
      "comments",
    ]) {
      assert.ok(document.diagnostics().some(({ details }) => details?.feature === feature));
    }
    assert.equal(document.diagnostics().some(({ details }) => (
      details?.feature === "three-dimensional-effects"
    )), false);
    assert.ok(document.diagnostics().some(({ code, details, objectId }) => (
      code === "ACTIVE_CONTENT_BLOCKED"
      && details?.feature === "native-action-command"
      && typeof objectId === "string"
    )));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("applies a slide color-map override before resolving scheme colors", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
      <p:clrMapOvr>
        <a:overrideClrMapping accent1="accent2"/>
      </p:clrMapOvr>
    </p:sld>`, {
    "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
        <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
      </Relationships>`,
    "ppt/theme/theme1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Colors">
        <a:themeElements><a:clrScheme name="Colors">
          <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
          <a:accent2><a:srgbClr val="0000FF"/></a:accent2>
        </a:clrScheme></a:themeElements>
      </a:theme>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    set fillStyle(value) { fills.push(value); },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(fills.includes("rgba(0, 0, 255, 1)"));
    assert.equal(fills.includes("rgba(255, 0, 0, 1)"), false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("applies PPTX shape paint, rounded geometry, and rotation through public APIs", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Rotated card"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm rot="5400000"><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="476250"/></a:xfrm>
            <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fillStyles = [];
  const strokeStyles = [];
  const lineWidths = [];
  const transforms = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {},
    clip() {}, fillText() {}, measureText() { return { width: 0 }; },
    transform(...args) { transforms.push(args); },
    set fillStyle(value) { fillStyles.push(value); },
    set strokeStyle(value) { strokeStyles.push(value); },
    set lineWidth(value) { lineWidths.push(value); },
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 200, y: 50 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.source.shapeId, 2);
    assert.ok(fillStyles.includes("rgba(255, 0, 0, 1)"));
    assert.ok(strokeStyles.includes("rgba(0, 255, 0, 1)"));
    assert.ok(lineWidths.includes(2));
    assert.equal(transforms.length, 1);
    assert.equal(frame.renderedObjectCount, 1);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("keeps direct PPTX shape paint ahead of theme style references", async () => {
  const style = (line = true) => `<p:style>
    ${line ? '<a:lnRef idx="1"><a:schemeClr val="accent2"/></a:lnRef>' : ""}
    <a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>
  </p:style>`;
  const shape = (id, x, properties, shapeStyle) => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${x}" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      ${properties}
    </p:spPr>
    ${shapeStyle}
  </p:sp>`;
  const slide = `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        ${shape(2, 476250, '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln>', style())}
        ${shape(3, 2857500, '<a:solidFill><a:srgbClr val="034EA2"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>', style())}
        ${shape(4, 5238750, "", style(false))}
        ${shape(5, 7143750, '<a:pattFill prst="pct5"><a:fgClr><a:srgbClr val="102030"/></a:fgClr><a:bgClr><a:srgbClr val="F0F0F0"><a:alpha val="50000"/></a:srgbClr></a:bgClr></a:pattFill>', style(false))}
      </p:spTree></p:cSld>
    </p:sld>`;
  const theme = `<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Paint precedence"><a:themeElements>
      <a:clrScheme name="Paint precedence">
        <a:accent1><a:srgbClr val="F37021"/></a:accent1>
        <a:accent2><a:srgbClr val="8454F6"/></a:accent2>
      </a:clrScheme>
    </a:themeElements></a:theme>`;
  const bytes = pptxWithSlide(slide, {
    "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
        <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
      </Relationships>`,
    "ppt/theme/theme1.xml": theme,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const strokes = [];
  const lineWidths = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    createPattern() { return { setTransform() {} }; },
    set fillStyle(value) { fills.push(value); },
    set strokeStyle(value) { strokes.push(value); },
    set lineWidth(value) { lineWidths.push(value); },
  };
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; }
    getContext() {
      return this.width === 8 && this.height === 8
        ? { fillRect() {}, set fillStyle(value) { fills.push(value); } }
        : context;
    }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    const shapeFills = fills.filter((value) => String(value).startsWith("rgba("));
    assert.deepEqual(shapeFills, [
      "rgba(255, 255, 255, 1)",
      "rgba(3, 78, 162, 1)",
      "rgba(243, 112, 33, 1)",
      "rgba(240, 240, 240, 0.5019607843137255)",
      "rgba(16, 32, 48, 1)",
    ]);
    assert.deepEqual(strokes, ["rgba(0, 255, 0, 1)"]);
    assert.deepEqual(lineWidths, [2]);
    assert.equal(frame.renderedObjectCount, 4);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("resolves grpFill from the nearest explicitly filled PPTX group", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="10"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/><a:chOff x="0" y="0"/><a:chExt cx="3810000" cy="1905000"/></a:xfrm>
            <a:solidFill><a:srgbClr val="CC0000"/></a:solidFill>
          </p:grpSpPr>
          <p:sp><p:nvSpPr><p:cNvPr id="11"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="952500" y="476250"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/><a:grpFill/></p:spPr>
          </p:sp>
          <p:grpSp>
            <p:nvGrpSpPr><p:cNvPr id="12"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr>
              <a:xfrm><a:off x="1905000" y="0"/><a:ext cx="1905000" cy="1905000"/><a:chOff x="0" y="0"/><a:chExt cx="1905000" cy="1905000"/></a:xfrm>
              <a:noFill/>
            </p:grpSpPr>
            <p:sp><p:nvSpPr><p:cNvPr id="13"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="476250" y="476250"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/><a:grpFill/></p:spPr>
            </p:sp>
          </p:grpSp>
        </p:grpSp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    set fillStyle(value) { fills.push(value); },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(fills.filter((value) => String(value).startsWith("rgba(")), [
      "rgba(204, 0, 0, 1)",
    ]);
    assert.equal(frame.renderedObjectCount, 1);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders cropped and tiled PPTX shape image fills", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="38100" cy="38100"/></a:xfrm>
            <a:prstGeom prst="rect"/>
            <a:blipFill><a:blip r:embed="rIdImage"/><a:srcRect l="25000" r="25000"/><a:tile/></a:blipFill>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
      </Relationships>`,
    "ppt/media/image1.png": onePixelPng,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const drawCalls = [];
  const smoothing = [];
  const patterns = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    createPattern(_source, repeat) { return { setTransform(value) { patterns.push({ repeat, ...value }); } }; },
    drawImage(...args) { drawCalls.push(args); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set imageSmoothingEnabled(value) { smoothing.push(["enabled", value]); },
    set imageSmoothingQuality(value) { smoothing.push(["quality", value]); },
  };
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; }
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 4, height: 2, close() {} }),
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(drawCalls.length, 1, "crop the source once and repeat its brush without overlapping tile edges");
    assert.deepEqual(drawCalls[0].slice(1), [1, 0, 2, 2, 0, 0, 2, 2]);
    assert.deepEqual(patterns, [{ repeat: "repeat", a: 1, b: 0, c: 0, d: 1, e: 100, f: 100 }]);
    assert.deepEqual(smoothing, [["enabled", true], ["quality", "high"]]);
    assert.equal(frame.renderedObjectCount, 1);
    assert.equal(frame.diagnostics.length, 0);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("clips PPTX pictures to their authored shape and paints the outline", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="8"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdImage"/><a:srcRect r="25000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="38100" cy="38100"/></a:xfrm>
            <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
            <a:ln w="9525"><a:solidFill><a:srgbClr val="D7D743"/></a:solidFill><a:round/></a:ln>
          </p:spPr>
        </p:pic>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
      </Relationships>`,
    "ppt/media/image1.png": onePixelPng,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const drawCalls = [];
  let curves = 0;
  let clips = 0;
  let strokes = 0;
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {},
    bezierCurveTo() { curves += 1; }, quadraticCurveTo() {}, closePath() {}, fill() {},
    stroke() { strokes += 1; }, clip() { clips += 1; }, fillText() {},
    measureText() { return { width: 0 }; },
    drawImage(...args) { drawCalls.push(args); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set imageSmoothingEnabled(_value) {}, set imageSmoothingQuality(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 4, height: 4, close() {} }),
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(drawCalls.length, 1);
    assert.deepEqual(drawCalls[0].slice(1), [0, 0, 3, 4, 100, 100, 4, 4]);
    assert.ok(curves >= 8, `expected the oval to be traced for stroke and fill, got ${curves}`);
    assert.ok(clips >= 1);
    assert.equal(strokes, 1);
    assert.equal(frame.renderedObjectCount, 1);
    assert.equal(frame.diagnostics.length, 0);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
  }
});

test("opens PPTX files with duplicate presentation theme relationships", async () => {
  const slide = `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
    </p:sld>`;
  const theme = (name) => `<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}">
      <a:themeElements><a:clrScheme name="${name}">
        <a:dk1><a:srgbClr val="000000"/></a:dk1>
        <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      </a:clrScheme></a:themeElements>
    </a:theme>`;
  const bytes = pptxWithSlide(slide, {
    "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
        <Relationship Id="rIdTheme1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
        <Relationship Id="rIdTheme2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme2.xml"/>
      </Relationships>`,
    "ppt/theme/theme1.xml": theme("Primary"),
    "ppt/theme/theme2.xml": theme("Duplicate"),
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);

    assert.equal(document.info.format, "pptx");
    assert.equal(document.info.units.length, 1);
    assert.ok(document.diagnostics().some(({ code, message }) => (
      code === "FORMAT_INVALID" && message.includes("multiple presentation theme relationships")
    )));
  } finally {
    document?.close();
    engine.close();
  }
});

test("keeps PPTX slide content when a picture package part is missing", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Visible sibling</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdMissing"/></p:blipFill>
          <p:spPr><a:xfrm><a:off x="3810000" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>
        </p:pic>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdMissing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="NULL"/>
      </Relationships>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    assert.deepEqual(
      (await document.listObjects({ unitIndex: 0, textOnly: true })).map(({ text }) => text),
      ["Visible sibling"],
    );
    assert.ok(document.diagnostics().some(({ fidelity, message }) => (
      fidelity === "not-rendered" && message.includes("picture package part")
    )));
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders numeric PPTX custGeom paths instead of rectangle fallbacks", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>
          <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
          <a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:pathLst><a:path w="200" h="100">
            <a:moveTo><a:pt x="0" y="100"/></a:moveTo>
            <a:lnTo><a:pt x="100" y="0"/></a:lnTo>
            <a:lnTo><a:pt x="200" y="100"/></a:lnTo><a:close/>
          </a:path></a:pathLst></a:custGeom>
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        </p:spPr></p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const moves = [];
  const lines = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, rect() {}, bezierCurveTo() {}, quadraticCurveTo() {},
    closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    moveTo(x, y) { moves.push([x, y]); }, lineTo(x, y) { lines.push([x, y]); },
    measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(moves, [[100, 200]]);
    assert.deepEqual(lines, [[200, 100], [300, 200]]);
    assert.equal(document.diagnostics().some(({ message }) => message.includes("geometry uses a rectangle fallback")), false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("evaluates PPTX custom geometry guides and arcTo commands", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>
          <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
          <a:custGeom>
            <a:avLst/><a:gdLst>
              <a:gd name="startX" fmla="+- w 0 wd4"/>
              <a:gd name="startY" fmla="val vc"/>
              <a:gd name="radiusX" fmla="*/ w 1 4"/>
              <a:gd name="radiusY" fmla="*/ h 1 2"/>
              <a:gd name="sweep" fmla="val cd2"/>
            </a:gdLst><a:ahLst/><a:cxnLst/><a:pathLst><a:path w="200" h="100">
              <a:moveTo><a:pt x="startX" y="startY"/></a:moveTo>
              <a:arcTo wR="radiusX" hR="radiusY" stAng="0" swAng="sweep"/>
            </a:path></a:pathLst>
          </a:custGeom>
          <a:noFill/><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
        </p:spPr></p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const moves = [];
  const curves = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, rect() {}, lineTo() {}, quadraticCurveTo() {},
    closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    moveTo(x, y) { moves.push([x, y]); },
    bezierCurveTo(...args) { curves.push(args); },
    measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(moves[0], [250, 150]);
    assert.equal(curves.length, 2);
    assert.ok(Math.abs(curves.at(-1).at(-2) - 150) < 0.001);
    assert.ok(Math.abs(curves.at(-1).at(-1) - 150) < 0.001);
    assert.equal(document.diagnostics().some(({ message }) => message.includes("geometry uses a rectangle fallback")), false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders PPTX arc and rectangular callout presets without rectangle fallbacks", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Arc"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm>
            <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 12108080"/><a:gd name="adj2" fmla="val 9601012"/></a:avLst></a:prstGeom>
            <a:noFill/><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
          </p:spPr>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Callout"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="2857500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="wedgeRectCallout"><a:avLst><a:gd name="adj1" fmla="val -65780"/><a:gd name="adj2" fmla="val -243"/></a:avLst></a:prstGeom>
            <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const beziers = [];
  const lines = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, rect() {}, quadraticCurveTo() {},
    closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {}, moveTo() {},
    lineTo(x, y) { lines.push([x, y]); },
    bezierCurveTo(...values) { beziers.push(values); },
    measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(beziers.length >= 4);
    assert.ok(lines.some(([x]) => x > 250 && x < 300));
    assert.equal(document.diagnostics().some(({ message }) => message.includes("rectangle fallback")), false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("mirrors multi-path PPTX arcs once around their authored bounds", async () => {
  const arc = (id, x, flip = "") => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm${flip}><a:off x="${x}" y="952500"/><a:ext cx="1905000" cy="1943100"/></a:xfrm>
      <a:prstGeom prst="arc"><a:avLst>
        <a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 5486615"/>
      </a:avLst></a:prstGeom>
      <a:noFill/><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
    </p:spPr>
  </p:sp>`;
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        ${arc(2, 952500)}
        ${arc(3, 3810000, ' flipH="1"')}
        ${arc(4, 6667500, ' flipV="1"')}
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const matrices = [];
  const moves = [];
  const endpoints = [];
  let matrix = [1, 0, 0, 1, 0, 0];
  const stack = [];
  const multiply = ([a, b, c, d, e, f], [g, h, i, j, k, l]) => [
    a * g + c * h, b * g + d * h,
    a * i + c * j, b * i + d * j,
    a * k + c * l + e, b * k + d * l + f,
  ];
  const point = (x, y) => [
    matrix[0] * x + matrix[2] * y + matrix[4],
    matrix[1] * x + matrix[3] * y + matrix[5],
  ];
  const context = {
    globalAlpha: 1,
    save() { stack.push([...matrix]); },
    restore() { matrix = stack.pop() ?? [1, 0, 0, 1, 0, 0]; },
    scale(x, y) { matrix = multiply(matrix, [x, 0, 0, y, 0, 0]); },
    translate(x, y) { matrix = multiply(matrix, [1, 0, 0, 1, x, y]); },
    transform(...values) { matrices.push(values); matrix = multiply(matrix, values); },
    fillRect() {}, beginPath() {}, ellipse() {}, rect() {}, lineTo() {}, quadraticCurveTo() {},
    closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    moveTo(x, y) { moves.push(point(x, y)); },
    bezierCurveTo(_cp1x, _cp1y, _cp2x, _cp2y, x, y) { endpoints.push(point(x, y)); },
    measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set lineCap(_value) {}, set lineJoin(_value) {}, set miterLimit(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(matrices, [
      [-1, -0, -0, 1, 1000, 0],
      [1, 0, 0, -1, 0, 404],
    ]);
    assert.ok(endpoints.some(([x]) => x > 285 && x < 301), "unflipped arc must sweep through its right half");
    assert.ok(endpoints.some(([x]) => x > 399 && x < 415), "flipH arc must sweep through its left half");
    assert.ok(endpoints.some(([x]) => x > 885 && x < 901), "flipV must preserve the arc's horizontal half");
    assert.ok(moves.some(([x, y]) => Math.abs(x - 800) < 1 && Math.abs(y - 304) < 1),
      "flipV must mirror the top endpoint to the bottom around the bounds center");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders the PPTX star5 preset as a five-point star instead of a rectangle", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="17"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm>
            <a:prstGeom prst="star5"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
            <a:ln><a:noFill/></a:ln>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const traced = { rects: 0, moves: 0, lines: 0, closes: 0 };
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, bezierCurveTo() {}, quadraticCurveTo() {},
    moveTo() { traced.moves += 1; },
    lineTo() { traced.lines += 1; },
    rect() { traced.rects += 1; },
    closePath() { traced.closes += 1; },
    fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(traced.rects, 0);
    assert.equal(traced.moves, 1);
    assert.equal(traced.lines, 9);
    assert.equal(traced.closes, 1);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders all 187 DrawingML preset shapes without rectangle fallbacks", async () => {
  const presets = drawingMlPresetShapes;
  assert.equal(presets.length, 187);
  const shapeXml = presets.map((preset, index) => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${index + 2}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${(index % 17) * 533000}" y="${Math.floor(index / 17) * 600000}"/><a:ext cx="400000" cy="400000"/></a:xfrm>
      <a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
      <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
    </p:spPr>
  </p:sp>`).join("");
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${shapeXml}</p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const traced = { moves: 0, lines: 0, beziers: 0, closes: 0 };
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, rect() {}, quadraticCurveTo() {},
    moveTo() { traced.moves += 1; },
    lineTo() { traced.lines += 1; },
    bezierCurveTo() { traced.beziers += 1; },
    closePath() { traced.closes += 1; },
    fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(
      document.diagnostics()
        .filter(({ message }) => message.includes("fallback"))
        .map(({ message }) => message),
      [],
    );
    assert.ok(traced.moves >= presets.length);
    assert.ok(traced.lines >= 100);
    assert.ok(traced.beziers >= 10);
    assert.ok(traced.closes >= 15);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DrawingML bevel faces with Office path fill color modifiers", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp>
        <p:nvSpPr><p:cNvPr id="34"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="1714500"/></a:xfrm>
          <a:prstGeom prst="bevel"><a:avLst><a:gd name="adj" fmla="val 25198"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="00B0F0"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
        </p:spPr>
        <p:style>
          <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
          <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
          <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
        </p:style>
        <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p></p:txBody>
      </p:sp></p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  let currentFill = "";
  let strokes = 0;
  const context = {
    globalAlpha: 1,
    filter: undefined,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() { fills.push(currentFill); },
    stroke() { strokes += 1; }, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    set fillStyle(value) { currentFill = value; },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(fills, [
      "rgba(0, 176, 240, 1)",
      "rgba(50, 191, 243, 1)",
      "rgba(0, 141, 193, 1)",
      "rgba(102, 208, 246, 1)",
      "rgba(0, 106, 144, 1)",
    ]);
    assert.equal(strokes, 1);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("preserves DrawingML preset path fill and stroke semantics", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="1905000"/></a:xfrm>
            <a:prstGeom prst="actionButtonHelp"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="203864"/></a:solidFill></a:ln>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fillColors = [];
  let currentFill = "";
  let strokes = 0;
  const context = {
    globalAlpha: 1,
    filter: "none",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, rect() {}, moveTo() {}, lineTo() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() { fillColors.push(currentFill); }, stroke() { strokes += 1; },
    clip() {}, fillText() {}, measureText() { return { width: 0 }; },
    set fillStyle(value) { currentFill = value; }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(
      fillColors,
      ["rgba(68, 114, 196, 1)", "rgba(41, 68, 118, 1)"],
      "normal and darkened paths should use their Office-resolved colors",
    );
    assert.equal(strokes, 2, "only the two fill-none outline paths should be stroked");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("inherits visible PPTX master and layout shapes", async () => {
  const slide = `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
    </p:sld>`;
  const shape = (id, x, color) => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="${x}" y="952500"/><a:ext cx="476250" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></p:spPr>
  </p:sp>`;
  const bytes = pptxWithSlide(slide, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    </Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${shape(21, 1905000, "00FF00")}</p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${shape(11, 952500, "0000FF")}</p:spTree></p:cSld></p:sldMaster>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const masterHits = await document.hitTest({ unitIndex: 0, x: 110, y: 110 });
    const layoutHits = await document.hitTest({ unitIndex: 0, x: 210, y: 110 });

    assert.equal(masterHits[0].object.source.part, "ppt/slideMasters/slideMaster1.xml");
    assert.equal(masterHits[0].object.source.shapeId, 11);
    assert.equal(layoutHits[0].object.source.part, "ppt/slideLayouts/slideLayout1.xml");
    assert.equal(layoutHits[0].object.source.shapeId, 21);
  } finally {
    document?.close();
    engine.close();
  }
});

test("inherits theme color from a PPTX layout placeholder paragraph style", async () => {
  const slide = `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp>
      <p:nvSpPr><p:cNvPr id="31"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Layout blue</a:t></a:r></a:p></p:txBody>
    </p:sp></p:spTree></p:cSld>
  </p:sld>`;
  const placeholder = `<p:sp>
    <p:nvSpPr><p:cNvPr id="21"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="dk2"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody>
  </p:sp>`;
  const bytes = pptxWithSlide(slide, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${placeholder}</p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle></p:txStyles></p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/theme/theme1.xml": `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="2058F3"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2></a:clrScheme></a:themeElements></a:theme>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const textFills = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText() { textFills.push(this.fillStyle); }, measureText() { return { width: 20 }; },
    fillStyle: "", set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(textFills.includes("rgba(32, 88, 243, 1)"), JSON.stringify(textFills));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("inherits a PPTX master title color through a centered-title layout placeholder", async () => {
  const placeholder = (id, type, text = "") => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle>${type === "title" ? `<a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl1pPr>` : ""}</a:lstStyle><a:p>${text ? `<a:r><a:rPr sz="2400"/><a:t>${text}</a:t></a:r>` : ""}</a:p></p:txBody>
  </p:sp>`;
  const slide = `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${placeholder(31, "ctrTitle", "White title")}</p:spTree></p:cSld>
  </p:sld>`;
  const bytes = pptxWithSlide(slide, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${placeholder(21, "ctrTitle")}</p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${placeholder(11, "title")}</p:spTree></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle></p:txStyles></p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/theme/theme1.xml": `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme></a:themeElements></a:theme>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const textFills = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText() { textFills.push(this.fillStyle); }, measureText() { return { width: 20 }; },
    fillStyle: "", set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(textFills.includes("rgba(255, 255, 255, 1)"), JSON.stringify(textFills));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("resolves PPTX master theme colors, transforms, and major/minor fonts", async () => {
  const slide = `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
    </p:sld>`;
  const masterShape = `<p:sp>
    <p:nvSpPr><p:cNvPr id="61"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/>
      <a:solidFill><a:schemeClr val="accent1"><a:tint val="25000"/><a:satMod val="160000"/><a:alpha val="50000"/></a:schemeClr></a:solidFill>
      <a:ln w="19050"><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="50000"/><a:lumOff val="10000"/></a:schemeClr></a:solidFill></a:ln>
    </p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
      <a:r><a:rPr sz="1800"><a:solidFill><a:schemeClr val="hlink"><a:shade val="50000"/></a:schemeClr></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>Major</a:t></a:r>
      <a:r><a:rPr sz="1800"><a:latin typeface="+mn-lt"/></a:rPr><a:t>Minor</a:t></a:r>
    </a:p></p:txBody>
  </p:sp>`;
  const theme = `<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme"><a:themeElements>
      <a:clrScheme name="Custom">
        <a:dk1><a:sysClr val="windowText" lastClr="101010"/></a:dk1>
        <a:lt1><a:sysClr val="window" lastClr="F0F0F0"/></a:lt1>
        <a:dk2><a:srgbClr val="202020"/></a:dk2><a:lt2><a:srgbClr val="E0E0E0"/></a:lt2>
        <a:accent1><a:srgbClr val="204060"/></a:accent1><a:accent2><a:srgbClr val="808080"/></a:accent2>
        <a:accent3><a:srgbClr val="20A020"/></a:accent3><a:accent4><a:srgbClr val="2020A0"/></a:accent4>
        <a:accent5><a:srgbClr val="A0A020"/></a:accent5><a:accent6><a:srgbClr val="20A0A0"/></a:accent6>
        <a:hlink><a:srgbClr val="00AA00"/></a:hlink><a:folHlink><a:srgbClr val="AA00AA"/></a:folHlink>
      </a:clrScheme>
      <a:fontScheme name="Custom Fonts">
        <a:majorFont><a:latin typeface="Major Theme"/><a:ea typeface="Major East Asian"/><a:cs typeface="Major Complex"/></a:majorFont>
        <a:minorFont><a:latin typeface="Minor Theme"/><a:ea typeface="Minor East Asian"/><a:cs typeface="Minor Complex"/></a:minorFont>
      </a:fontScheme>
    </a:themeElements></a:theme>`;
  const bytes = pptxWithSlide(slide, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${masterShape}</p:spTree></p:cSld></p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/theme/theme1.xml": theme,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const fonts = [];
  const strokes = [];
  const context = {
    globalAlpha: 1,
    letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 20 }; },
    set font(value) { fonts.push(value); },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(value) { strokes.push(value); }, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 110, y: 110 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.source.part, "ppt/slideMasters/slideMaster1.xml");
    assert.ok(fills.includes("rgba(224, 226, 229, 0.5019607843137255)"));
    assert.ok(fills.includes("rgba(0, 124, 0, 1)"));
    assert.ok(strokes.includes("rgba(90, 90, 90, 1)"));
    assert.deepEqual(
      hits[0].object.fontRuns.map(({ authoredFamily, renderedFamily, source }) => ({
        authoredFamily,
        renderedFamily,
        source,
      })),
      [
        { authoredFamily: "Major Theme", renderedFamily: "Calibri", source: "fallback" },
        { authoredFamily: "Minor Theme", renderedFamily: "Calibri", source: "fallback" },
      ],
    );
    assert.ok(fonts.some((font) => font.includes("Calibri")));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders scaled DrawingML linear gradients in normalized fill space", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="62"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"/>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="FFFFFF"><a:shade val="30000"/></a:srgbClr></a:gs>
              <a:gs pos="50000"><a:srgbClr val="FFFFFF"><a:shade val="67500"/></a:srgbClr></a:gs>
              <a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>
            </a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const linearGradients = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    createLinearGradient(...args) {
      const stops = [];
      linearGradients.push({ args, stops });
      return { addColorStop(offset, color) { stops.push([offset, color]); } };
    },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(linearGradients.length, 1);
    const [{ args, stops }] = linearGradients;
    const [x0, y0, x1, y1] = args;
    const dx = x1 - x0;
    const dy = y1 - y0;
    const parameterAt = (x, y) => ((x - x0) * dx + (y - y0) * dy) / (dx * dx + dy * dy);
    assert.ok(Math.abs(parameterAt(100, 100)) < 0.0001);
    assert.ok(Math.abs(parameterAt(500, 300) - 1) < 0.0001);
    assert.ok(Math.abs(parameterAt(500, 100) - 0.5) < 0.0001);
    assert.ok(Math.abs(parameterAt(100, 300) - 0.5) < 0.0001);
    assert.deepEqual(stops, [
      [0, "rgba(149, 149, 149, 1)"],
      [0.5, "rgba(214, 214, 214, 1)"],
      [1, "rgba(255, 255, 255, 1)"],
    ]);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("preserves circular DrawingML path gradients and their focus rectangle", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="62"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
              <a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>
            </a:gsLst><a:path path="circle"><a:fillToRect r="100000" b="100000"/></a:path><a:tileRect l="-100000" t="-100000"/></a:gradFill>
          </p:spPr>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="63"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="3810000" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="7030A0"/></a:gs>
              <a:gs pos="100000"><a:srgbClr val="4472C4"/></a:gs>
            </a:gsLst><a:path path="circle"><a:fillToRect l="100000" t="100000"/></a:path><a:tileRect r="-100000" b="-100000"/></a:gradFill>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const rasterImages = [];
  const drawCalls = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    getTransform() { return { a: 1, b: 0, c: 0, d: 1 }; },
    createImageData(width, height) {
      return { width, height, data: new Uint8ClampedArray(width * height * 4) };
    },
    putImageData(image) { rasterImages.push(image); },
    drawImage(...args) { drawCalls.push(args); },
    createLinearGradient() { throw new Error("circular path gradient was flattened to a line"); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(rasterImages.length, 2);
    assert.deepEqual(rasterImages.map(({ width, height }) => [width, height]), [[200, 100], [200, 100]]);
    assert.deepEqual(drawCalls.map((call) => call.slice(1)), [
      [100, 100, 200, 100],
      [400, 100, 200, 100],
    ]);
    assert.equal(rasterImages.every(({ data }) => data.some((value) => value !== data[0])), true);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("applies PPTX group transforms and inherited 3D scenes while preserving ancestry", async () => {
  // Command recording checks affine composition; test-picture-transforms.mjs checks perspective pixels.
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="8"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="1905000"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm>
            <a:scene3d><a:camera prst="perspectiveAboveLeftFacing" zoom="60000" fov="0"><a:rot lat="300000" lon="-600000" rev="0"/></a:camera><a:lightRig rig="threePt" dir="tr"/></a:scene3d>
          </p:grpSpPr>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:sp3d extrusionH="38100" prstMaterial="plastic"/></p:spPr>
          </p:sp>
        </p:grpSp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const transforms = [];
  let fills = 0;
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, translate() {}, transform(...values) { transforms.push(values); }, fillRect() {},
    scale() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() { fills += 1; }, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {}, set filter(_value) {},
    set shadowColor(_value) {}, set shadowBlur(_value) {}, set shadowOffsetX(_value) {}, set shadowOffsetY(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 130, y: 130 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.source.shapeId, 9);
    assert.deepEqual(hits[0].object.bounds, { x: 120, y: 120, width: 40, height: 40 });
    assert.equal(hits[0].ancestors.length, 1);
    assert.equal(hits[0].ancestors[0].type, "group");
    assert.equal(hits[0].ancestors[0].source.shapeId, 8);
    assert.ok(transforms.some(([a, b, c, d]) => Math.abs(a * d - b * c) < 0.49));
    assert.ok(fills >= 5);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("keeps grouped PPTX text at its authored font size", async () => {
  const bytes = pptxWithSlide(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="8"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="9"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1600"/><a:t>Review</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:grpSp></p:spTree></p:cSld>
  </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const scales = [1];
  const painted = [];
  let currentFont = "";
  const context = {
    globalAlpha: 1,
    save() { scales.push(scales.at(-1)); }, restore() { scales.pop(); },
    scale(x, y) { scales[scales.length - 1] *= Math.sqrt(Math.abs(x * y)); },
    transform(a, b, c, d) { scales[scales.length - 1] *= Math.sqrt(Math.abs(a * d - b * c)); },
    translate() {}, rotate() {}, fillRect() {}, beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {},
    rect() {}, bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text) {
      const size = Number.parseFloat(/\s([\d.]+)px\s/u.exec(currentFont)?.[1] ?? "NaN");
      painted.push({ text, effectiveSize: size * scales.at(-1) });
    },
    measureText(text) { return { width: text.length * 8 }; },
    set font(value) { currentFont = value; }, get font() { return currentFont; },
    letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr", fillStyle: "",
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    const review = painted.find(({ text }) => text === "Review");
    assert.ok(Math.abs(review.effectiveSize - 64 / 3) < 0.01, JSON.stringify(painted));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("honors absolute PPTX bullet size independently of the text run", async () => {
  const bytes = pptxWithSlide(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp>
      <p:nvSpPr><p:cNvPr id="9"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buSzPts val="1100"/><a:buChar char="●"/></a:pPr><a:r><a:rPr sz="1467"/><a:t>Item</a:t></a:r></a:p></p:txBody>
    </p:sp></p:spTree></p:cSld>
  </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const painted = [];
  let currentFont = "";
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, transform() {}, translate() {}, rotate() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text) {
      painted.push({ text, size: Number.parseFloat(/\s([\d.]+)px\s/u.exec(currentFont)?.[1] ?? "NaN") });
    },
    measureText(text) { return { width: text.length * 8 }; },
    set font(value) { currentFont = value; }, get font() { return currentFont; },
    letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr", fillStyle: "",
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    const bullet = painted.find(({ text }) => text === "●");
    const item = painted.find(({ text }) => text === "Item");
    assert.ok(Math.abs(bullet.size - 44 / 3) < 0.01, JSON.stringify(painted));
    assert.ok(Math.abs(item.size - 1467 / 75) < 0.01, JSON.stringify(painted));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("resolves a slide placeholder against inherited layout geometry", async () => {
  const slidePlaceholder = (id, text, transform = "") => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr>
    <p:spPr>${transform}<a:prstGeom prst="rect"/></p:spPr>
    <p:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
  </p:sp>`;
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${slidePlaceholder(31, "Slide title retains inherited geometry even when the text is long")}</p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    </Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${slidePlaceholder(21, "Layout title", '<a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="476250"/></a:xfrm>')}</p:spTree></p:cSld></p:sldLayout>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 110, y: 110 });

    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.source.part, "ppt/slides/slide1.xml");
    assert.equal(hits[0].object.source.shapeId, 31);
    assert.equal(hits[0].object.text, "Slide title retains inherited geometry even when the text is long");
    assert.deepEqual(hits[0].object.bounds, { x: 100, y: 100, width: 200, height: 50 });
  } finally {
    document?.close();
    engine.close();
  }
});

test("inherits layout text defaults without rendering placeholder prompts", async () => {
  const placeholder = (id, index, text, body = "", transform = "", customPrompt = false, type = "body") => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr><p:ph${type === null ? "" : ` type="${type}"`} idx="${index}"${customPrompt ? ' hasCustomPrompt="1"' : ""}/></p:nvPr></p:nvSpPr>
    <p:spPr>${transform}<a:prstGeom prst="rect"/></p:spPr>
    <p:txBody><a:bodyPr/>${body}<a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
  </p:sp>`;
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${placeholder(31, 41, "Slide title")}</p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    </Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
      ${placeholder(21, 41, "Layout title", '<a:lstStyle><a:lvl1pPr><a:defRPr sz="3600"/></a:lvl1pPr></a:lstStyle>', '<a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm>', true)}
      ${placeholder(22, 11, "Presenter Name", "", '<a:xfrm><a:off x="952500" y="2857500"/><a:ext cx="1905000" cy="476250"/></a:xfrm>', true)}
      ${placeholder(23, 12, "Untyped object prompt", "", '<a:xfrm><a:off x="952500" y="3810000"/><a:ext cx="1905000" cy="476250"/></a:xfrm>', false, null)}
    </p:spTree></p:cSld></p:sldLayout>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fonts = [];
  const context = {
    globalAlpha: 1,
    letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 20 }; },
    set font(value) { fonts.push(value); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const objects = await document.listObjects({ unitIndex: 0 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(objects.some(({ text }) => text === "Presenter Name"), false);
    assert.equal(objects.some(({ text }) => text === "Untyped object prompt"), false);
    assert.ok(fonts.some((font) => font.includes("48px")), `expected inherited 36pt font, got ${fonts.join(", ")}`);
    assert.equal(frame.renderedObjectCount, 1);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("matches fixed PPTX placeholders by type and suppresses uninstantiated inherited content", async () => {
  const placeholder = (id, type, index, text, transform = "", fieldType = null) => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${index}"/></p:nvPr></p:nvSpPr>
    <p:spPr>${transform}<a:prstGeom prst="rect"/></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/><a:p>${fieldType
      ? `<a:fld id="{00000000-0000-0000-0000-000000000001}" type="${fieldType}"><a:rPr/><a:t>${text}</a:t></a:fld>`
      : `<a:r><a:t>${text}</a:t></a:r>`}</a:p></p:txBody>
  </p:sp>`;
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        ${placeholder(31, "sldNum", 11, "‹#›", "", "slidenum")}
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    </Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
      ${placeholder(21, "sldNum", 11, "‹#›", "", "slidenum")}
      ${placeholder(22, "dt", 10, "2026/2/2", "", "datetimeFigureOut")}
    </p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
      ${placeholder(2, "title", 0, "Master Title Formatting", '<a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="476250"/></a:xfrm>')}
      ${placeholder(3, "sldNum", 4, "‹#›", '<a:xfrm><a:off x="8572500" y="6191250"/><a:ext cx="476250" cy="285750"/></a:xfrm>', "slidenum")}
      ${placeholder(4, "dt", 10, "2026/2/2", '<a:xfrm><a:off x="952500" y="6191250"/><a:ext cx="1905000" cy="285750"/></a:xfrm>', "datetimeFigureOut")}
    </p:spTree></p:cSld></p:sldMaster>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const objects = await document.listObjects({ unitIndex: 0 });

    assert.equal(objects.some(({ text }) => text === "Master Title Formatting"), false);
    assert.equal(objects.some(({ text }) => text === "‹#›"), false);
    assert.equal(objects.some(({ text }) => text === "2026/2/2"), false);
    const slideNumber = objects.find(({ text }) => text === "1");
    assert.ok(slideNumber);
    assert.equal(slideNumber.source.part, "ppt/slides/slide1.xml");
    assert.equal(slideNumber.source.shapeId, 31);
    assert.deepEqual(slideNumber.bounds, { x: 900, y: 650, width: 50, height: 30 });
  } finally {
    document?.close();
    engine.close();
  }
});

test("inherits PPTX body bounds through an untyped object placeholder", async () => {
  const placeholder = (id, placeholder, transform = "", text = "Prompt") => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr><p:ph ${placeholder}/></p:nvPr></p:nvSpPr>
    <p:spPr>${transform}<a:prstGeom prst="rect"/></p:spPr>
    <p:txBody><a:bodyPr><a:normAutofit/><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="1800000" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d extrusionH="127000"><a:extrusionClr><a:schemeClr val="accent2"/></a:extrusionClr></a:sp3d></a:bodyPr><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
  </p:sp>`;
  const bytes = pptxWithSlide(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${placeholder(5, 'idx="1"', "", "29% of slides created in a typical PowerPoint presentation contain only text")}</p:spTree></p:cSld>
  </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${placeholder(3, 'idx="1"')}</p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${placeholder(3, 'type="body" idx="1"', '<a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm>')}</p:spTree></p:cSld></p:sldMaster>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const textDraws = [];
  const depthOffsets = [];
  let imageDraws = 0;
  class FakeCanvas {
    constructor(width = 960, height = 720) { this.width = width; this.height = height; }
    getContext() {
      const state = {
        canvas: this, globalAlpha: 1, globalCompositeOperation: "source-over", filter: "none",
        fillStyle: "", font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
        measureText(text) { return { width: text.length * 10 }; },
        fillText(text) { textDraws.push(text); },
        translate(x, y) {
          if (Math.abs(x) < 0.001 && y < 0 && y > -20) depthOffsets.push(y);
        },
        fillRect() {},
        drawImage() { imageDraws++; },
        getTransform() { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; },
      };
      return new Proxy(state, { get: (target, key) => key in target ? target[key] : () => {} });
    }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const body = (await document.listObjects({ unitIndex: 0 })).find(({ source }) => source.shapeId === 5);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(body.text, "29% of slides created in a typical PowerPoint presentation contain only text");
    assert.deepEqual(body.bounds, { x: 48, y: 168, width: 864, height: 475.1667175292969 });
    assert.equal(textDraws.filter((text) => text === "29%").length, 1, "text metrics are resolved once for all depth layers");
    assert.ok(imageDraws > 1 && depthOffsets.length > 1, "text body extrusion reuses the glyph surface at each depth");
    assert.ok(
      Math.abs(Math.min(...depthOffsets) + 20 / 3) < 0.01,
      `30-degree camera must project 127000 EMU extrusion to 6.67 px, got ${Math.min(...depthOffsets)}`,
    );
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("inherits PPTX automatic numbering, hanging indentation, and paragraph spacing", async () => {
  const bodyPlaceholder = (id, textBody, transform = "", listStyle = "") => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="18"/></p:nvPr></p:nvSpPr>
    <p:spPr>${transform}<a:prstGeom prst="rect"/></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle>${listStyle}</a:lstStyle>${textBody}</p:txBody>
  </p:sp>`;
  const paragraph = (text) => `<a:p><a:r><a:rPr sz="2000"/><a:t>${text}</a:t></a:r></a:p>`;
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        ${bodyPlaceholder(31, `${paragraph("First item")}${paragraph("Second item")}`)}
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    </Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
      ${bodyPlaceholder(21, paragraph("Prompt"), '<a:xfrm><a:off x="952500" y="952500"/><a:ext cx="5715000" cy="2857500"/></a:xfrm>', '<a:lvl1pPr marL="457200" indent="-457200"><a:lnSpc><a:spcPct val="114000"/></a:lnSpc><a:buAutoNum type="arabicPeriod"/><a:defRPr sz="2500"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="1800"/></a:lvl2pPr>')}
    </p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:txStyles><p:bodyStyle>
      <a:lvl1pPr><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:buNone/></a:lvl1pPr>
    </p:bodyStyle></p:txStyles></p:sldMaster>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const drawn = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    measureText(text) { return { width: text.length * 10 }; },
    fillText(text, x, y) { drawn.push({ text, x, y }); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const objects = await document.listObjects({ unitIndex: 0 });
    frame = await document.render({ unitIndex: 0 });
    const body = objects.find(({ source }) => source.shapeId === 31);
    const firstMarker = drawn.find(({ text }) => text === "1.");
    const secondMarker = drawn.find(({ text }) => text === "2.");
    const firstText = drawn.find(({ text }) => text === "First");

    assert.equal(body.text, "1.\tFirst item\n2.\tSecond item");
    assert.ok(firstMarker && secondMarker && firstText);
    assert.ok(Math.abs(firstMarker.x - 109.6) < 0.01);
    assert.ok(Math.abs(firstText.x - 157.6) < 0.01);
    assert.ok(Math.abs(secondMarker.y - firstMarker.y - 49.813) < 0.02);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("parses DrawingML connector shapes", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="7" name="Approval line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="476250"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></p:spPr>
        </p:cxnSp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const objects = await document.listObjects({ unitIndex: 0 });

    assert.equal(objects.length, 1);
    assert.equal(objects[0].type, "shape");
    assert.equal(objects[0].source.shapeId, 7);
    assert.deepEqual(objects[0].bounds, { x: 100, y: 100, width: 200, height: 50 });
  } finally {
    document?.close();
    engine.close();
  }
});

test("keeps PPTX group line ends and text legible under nonuniform child coordinates", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="7"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>
          <p:sp><p:nvSpPr><p:cNvPr id="8"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="500"/></a:xfrm><a:prstGeom prst="line"/><a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:headEnd type="triangle"/></a:ln></p:spPr>
          </p:sp>
          <p:sp><p:nvSpPr><p:cNvPr id="9"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="500"/><a:ext cx="1000" cy="500"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
            <p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Readable</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:grpSp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  let currentPath = [];
  const strokes = [];
  const scales = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    scale(x, y) { scales.push([x, y]); },
    beginPath() { currentPath = []; }, ellipse() {}, rect() {}, bezierCurveTo() {}, quadraticCurveTo() {},
    moveTo(x, y) { currentPath.push(["moveTo", x, y]); },
    lineTo(x, y) { currentPath.push(["lineTo", x, y]); }, closePath() { currentPath.push(["closePath"]); },
    fill() {}, stroke() { strokes.push(currentPath.slice()); }, clip() {}, fillText() {}, strokeText() {},
    measureText(text) { return { width: text.length * 8 }; }, setLineDash() {},
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {}, set lineCap(_value) {},
    set lineJoin(_value) {}, set miterLimit(_value) {}, set font(_value) {}, set fontKerning(_value) {},
    set letterSpacing(_value) {}, set textBaseline(_value) {}, set textAlign(_value) {}, set direction(_value) {},
    set shadowColor(_value) {}, set shadowBlur(_value) {}, set shadowOffsetX(_value) {}, set shadowOffsetY(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(strokes.some((path) => {
      const points = path.filter(([kind]) => kind === "moveTo" || kind === "lineTo");
      return points.length >= 2 && Math.hypot(points.at(-1)[1] - points[0][1], points.at(-1)[2] - points[0][2]) > 0.05;
    }), "the group-scaled arrow must retain a visible shaft");
    assert.ok(scales.some(([x, y]) => Math.abs(x - y) > 0.000_001), "text must cancel nonuniform group coordinate scaling per axis");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders unfilled PPTX rectangles with sysDash outlines as dashed paths", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="15"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></a:xfrm>
            <a:prstGeom prst="rect"/><a:noFill/>
            <a:ln w="28575"><a:solidFill><a:srgbClr val="1F2279"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const traced = { rects: 0, moves: 0, lines: 0, dashes: [], caps: [] };
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {},
    fill() {}, stroke() {}, clip() {}, fillText() {},
    moveTo() { traced.moves += 1; },
    lineTo() { traced.lines += 1; },
    rect() { traced.rects += 1; },
    setLineDash(value) { traced.dashes.push(value); },
    measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set lineCap(value) { traced.caps.push(value); },
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(traced.rects > 0);
    assert.deepEqual(traced.dashes, [[9, 3]]);
    assert.ok(traced.caps.includes("butt"), "flat DrawingML caps must stay flat in Canvas");
    assert.equal(traced.moves, 0);
    assert.equal(traced.lines, 0);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("preserves the long-dash and dot rhythm of PPTX dashDot outlines", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="16"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></a:xfrm>
            <a:prstGeom prst="rect"/><a:noFill/>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="DD73C9"/></a:solidFill><a:prstDash val="dashDot"/></a:ln>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const dashes = [];
  let lineWidth = 1;
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {},
    fill() {}, stroke() {}, clip() {}, fillText() {},
    moveTo() {},
    lineTo() {},
    rect() {},
    setLineDash(value) { dashes.push(value); },
    measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(dashes, [[8, 6, 2, 6]]);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders bent, dashed PPTX connectors with arrowheads and preset arrows", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="8"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="bentConnector2"/><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/><a:headEnd type="triangle"/></a:ln>
          </p:spPr>
        </p:cxnSp>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="10"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="2381250"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></a:ln>
          </p:spPr>
          <p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></p:style>
        </p:cxnSp>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="11"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="3333750" y="2381250"/><a:ext cx="952500" cy="952500"/></a:xfrm>
            <a:prstGeom prst="bentConnector5"/><a:ln w="19050"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></a:ln>
          </p:spPr>
          <p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></p:style>
        </p:cxnSp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="9"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm flipV="1"><a:off x="3810000" y="952500"/><a:ext cx="476250" cy="476250"/></a:xfrm>
            <a:prstGeom prst="upArrow"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const traced = { rects: 0, moves: 0, lines: 0, linePoints: [], closes: 0, fills: 0, dashes: [] };
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, bezierCurveTo() {}, quadraticCurveTo() {}, fill() { traced.fills += 1; }, stroke() {}, clip() {},
    setLineDash(value) { traced.dashes.push(value); },
    moveTo() { traced.moves += 1; },
    lineTo(x, y) { traced.lines += 1; traced.linePoints.push([x, y]); },
    rect() { traced.rects += 1; }, closePath() { traced.closes += 1; },
    fillText() {}, measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {}, set font(_value) {},
    set letterSpacing(_value) {}, set textBaseline(_value) {}, set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(traced.rects, 0);
    assert.ok(traced.dashes.some((value) => value[0] === 8 && value[1] === 6));
    assert.ok(traced.lines > 8, `expected connector and arrow paths, got ${traced.lines} lines`);
    assert.ok(traced.closes >= 2, `expected closed arrowheads, got ${traced.closes}`);
    assert.equal(traced.fills, 2, "the connector arrowhead and preset arrow should both be filled");
    assert.equal(
      traced.linePoints.some(([x, y]) => Math.abs(x - 108) < 0.01 && Math.abs(y - 104) < 0.01)
        && traced.linePoints.some(([x, y]) => Math.abs(x - 108) < 0.01 && Math.abs(y - 96) < 0.01),
      true,
      "the default medium triangle should retain PowerPoint's 6 pt hairline size floor",
    );
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("stops PPTX connector strokes at closed arrowhead bases", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="8"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="0"/></a:xfrm>
            <a:prstGeom prst="straightConnector1"/>
            <a:ln w="19050">
              <a:solidFill><a:srgbClr val="808080"/></a:solidFill>
              <a:headEnd type="triangle"/>
              <a:tailEnd type="triangle"/>
            </a:ln>
          </p:spPr>
        </p:cxnSp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  let currentPath = [];
  const strokes = [];
  const fills = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() { currentPath = []; },
    ellipse() {}, rect() {}, bezierCurveTo() {}, quadraticCurveTo() {},
    moveTo(x, y) { currentPath.push(["moveTo", x, y]); },
    lineTo(x, y) { currentPath.push(["lineTo", x, y]); },
    closePath() { currentPath.push(["closePath"]); },
    fill() { fills.push(currentPath.slice()); },
    stroke() { strokes.push(currentPath.slice()); },
    clip() {}, fillText() {}, measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {}, set font(_value) {},
    set letterSpacing(_value) {}, set textBaseline(_value) {}, set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(strokes, [[["moveTo", 108, 100], ["lineTo", 292, 100]]]);
    assert.equal(fills.length, 1);
    assert.equal(fills[0].filter(([kind]) => kind === "closePath").length, 2);
    assert.equal(
      fills[0].some(([kind, x, y]) => kind === "moveTo" && x === 100 && y === 100)
        && fills[0].some(([kind, x, y]) => kind === "moveTo" && x === 300 && y === 100),
      true,
      "filled arrow tips should end exactly at the connector endpoints",
    );
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("honors DrawingML line-end width and length", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="8"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="0"/></a:xfrm>
            <a:prstGeom prst="straightConnector1"/><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle" w="lg" len="sm"/></a:ln>
          </p:spPr>
        </p:cxnSp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const linePoints = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo(x, y) { linePoints.push([x, y]); }, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText() {}, measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {}, set font(_value) {},
    set letterSpacing(_value) {}, set textBaseline(_value) {}, set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.equal(
      linePoints.some(([x, y]) => Math.abs(x - (300 - 16 / 3)) < 0.01 && Math.abs(y - (100 - 20 / 3)) < 0.01)
        && linePoints.some(([x, y]) => Math.abs(x - (300 - 16 / 3)) < 0.01 && Math.abs(y - (100 + 20 / 3)) < 0.01),
      true,
      "large width and small length should independently scale the triangle",
    );
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders PPTX compound line styles and diamond and oval line ends", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"/><a:noFill/>
            <a:ln w="19050" cap="rnd" cmpd="tri" algn="in">
              <a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="dash"/><a:round/>
            </a:ln>
          </p:spPr>
        </p:sp>
        <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="8"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="1905000"/><a:ext cx="1905000" cy="0"/></a:xfrm>
            <a:prstGeom prst="straightConnector1"/><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="diamond"/></a:ln>
          </p:spPr>
        </p:cxnSp>
        <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="9"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="2381250"/><a:ext cx="1905000" cy="0"/></a:xfrm>
            <a:prstGeom prst="straightConnector1"/><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="oval"/></a:ln>
          </p:spPr>
        </p:cxnSp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const dashes = [];
  const caps = [];
  const joins = [];
  let clips = 0;
  let strokes = 0;
  let fills = 0;
  let closedPaths = 0;
  let curves = 0;
  let currentPathCurves = 0;
  const fillCurveCounts = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() { currentPathCurves = 0; }, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, quadraticCurveTo() {},
    closePath() { closedPaths += 1; }, fill() { fills += 1; fillCurveCounts.push(currentPathCurves); }, stroke() { strokes += 1; },
    clip() { clips += 1; }, fillText() {}, bezierCurveTo() { curves += 1; currentPathCurves += 1; },
    setLineDash(value) { dashes.push(value); },
    measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set lineCap(value) { caps.push(value); },
    set lineJoin(value) { joins.push(value); },
    set miterLimit(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(dashes.some((value) => value.length === 2 && value[0] === 8 && value[1] === 6));
    assert.ok(caps.includes("round"));
    assert.ok(joins.includes("round"));
    assert.ok(clips >= 1);
    assert.ok(strokes >= 7);
    assert.ok(fills >= 2, `expected filled diamond and oval line ends, got ${fills}`);
    assert.ok(fillCurveCounts.includes(4), `expected a four-curve oval fill, got ${fillCurveCounts}`);
    assert.ok(closedPaths >= 2);
    assert.ok(curves >= 4);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders filled stealth and open arrow DrawingML line ends", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="8"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="0"/></a:xfrm>
            <a:prstGeom prst="straightConnector1"/><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
          </p:spPr>
        </p:cxnSp>
        <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="9"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="1905000"/><a:ext cx="1905000" cy="0"/></a:xfrm>
            <a:prstGeom prst="straightConnector1"/><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="arrow"/></a:ln>
          </p:spPr>
        </p:cxnSp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const traced = { closes: 0, fills: 0 };
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {}, quadraticCurveTo() {},
    closePath() { traced.closes += 1; }, fill() { traced.fills += 1; }, stroke() {}, clip() {},
    fillText() {}, measureText() { return { width: 0 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {}, set font(_value) {},
    set letterSpacing(_value) {}, set textBaseline(_value) {}, set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.deepEqual(traced, { closes: 2, fills: 1 });
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders PPTX rich text runs with their own typography and colors", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
            <a:r><a:rPr sz="1200" b="1"><a:solidFill><a:prstClr val="red"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>Red</a:t></a:r>
            <a:r><a:rPr sz="2400" i="1"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:latin typeface="Georgia"/></a:rPr><a:t>Blue</a:t></a:r>
            <a:r><a:rPr sz="1200"><a:solidFill><a:prstClr val="white"/></a:solidFill><a:highlight><a:srgbClr val="FFFF00"/></a:highlight><a:latin typeface="Aptos"/></a:rPr><a:t>White</a:t></a:r>
          </a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const texts = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    measureText() { return { width: 12 }; },
    fillText(text) { texts.push(text); },
    set fillStyle(value) { fills.push(value); },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 110, y: 110 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.text, "RedBlueWhite");
    assert.deepEqual(texts, ["Red", "Blue", "White"]);
    assert.ok(fills.includes("rgba(255, 0, 0, 1)"));
    assert.ok(fills.includes("rgba(0, 0, 255, 1)"));
    assert.ok(fills.includes("rgba(255, 255, 255, 1)"));
    assert.ok(fills.includes("rgba(255, 255, 0, 1)"));
    assert.equal(frame.renderedObjectCount, 1);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders a DrawingML shape line break as a styled line inside one paragraph", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="5"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p>
            <a:pPr algn="ctr"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:spcAft><a:spcPts val="3000"/></a:spcAft></a:pPr>
            <a:r><a:rPr sz="1200" b="1"><a:latin typeface="Segoe UI"/></a:rPr><a:t>SWE.1</a:t></a:r>
            <a:br/>
            <a:r><a:rPr sz="1200" i="1"><a:latin typeface="Arial"/></a:rPr><a:t>Software Engineering</a:t></a:r>
          </a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const painted = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    measureText(text) {
      return {
        width: text.length * 6,
        fontBoundingBoxAscent: 12,
        fontBoundingBoxDescent: 4,
        actualBoundingBoxAscent: 10,
      };
    },
    fillText(text, x, y) { painted.push({ text, x, y, font: this.font }); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 110, y: 110 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.text, "SWE.1\nSoftware Engineering");
    const visible = painted.filter(({ text }) => text.trim() !== "");
    const lineYs = [...new Set(visible.map(({ y }) => y))];
    const [headingY, descriptionY] = lineYs;
    const heading = visible.filter(({ y }) => y === headingY);
    const description = visible.filter(({ y }) => y === descriptionY);
    assert.equal(lineYs.length, 2);
    assert.equal(heading.map(({ text }) => text).join(""), "SWE.1");
    assert.equal(description.map(({ text }) => text).join(""), "SoftwareEngineering");
    assert.ok(heading.every(({ font }) => /700 16px "Calibri"/.test(font)));
    assert.ok(description.every(({ font }) => /italic 400 16px "Arial"/.test(font)));
    assert.equal(description[0].y - heading[0].y, 24);
    for (const line of [heading, description]) {
      for (let index = 1; index < line.length; index += 1) {
        assert.ok(line[index].x >= line[index - 1].x + line[index - 1].text.length * 6);
      }
    }
    assert.equal(frame.renderedObjectCount, 1);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("keeps semantic DrawingML Wingdings bullets when the exact glyph font is unavailable", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
            <a:pPr><a:buClr><a:srgbClr val="E44E1F"/></a:buClr><a:buFont typeface="Wingdings"/><a:buChar char="q"/></a:pPr>
            <a:r><a:rPr lang="zh-CN" sz="2400"><a:latin typeface="Times New Roman"/><a:ea typeface="Kai Text"/></a:rPr><a:t>中文</a:t></a:r>
            <a:r><a:rPr lang="en-US" sz="1200"><a:latin typeface="Aptos"/><a:ea typeface="Kai Text"/></a:rPr><a:t>ABC</a:t></a:r>
          </a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const drawn = [];
  let currentFill = "";
  let currentFont = "";
  const context = {
    globalAlpha: 1,
    letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    measureText(text) { return { width: text.length * 12 }; },
    fillText(text) { drawn.push({ text, fill: currentFill, font: currentFont }); },
    set fillStyle(value) { currentFill = value; },
    set font(value) { currentFont = value; },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const [textBox] = await document.listObjects({ unitIndex: 0 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(textBox.text, "❑\t中文ABC");
    assert.deepEqual(
      textBox.fontRuns.map(({ start, end, authoredFamily }) => ({ start, end, authoredFamily })),
      [
        { start: 0, end: 2, authoredFamily: "Arial" },
        { start: 2, end: 4, authoredFamily: "Kai Text" },
        { start: 4, end: 7, authoredFamily: "Aptos" },
      ],
    );
    assert.deepEqual(drawn.map(({ text }) => text), ["❑", "中文", "ABC"]);
    assert.equal(drawn[0].fill, "rgba(228, 78, 31, 1)");
    assert.ok(drawn[0].font.includes("32px"), `expected bullet to inherit first-run size, got ${drawn[0].font}`);
    assert.ok(drawn[0].font.includes("Segoe UI Symbol"), `expected semantic symbol fallback, got ${drawn[0].font}`);
    assert.ok(!/\s"Wingdings"(?:,|$)/u.test(drawn[0].font), `semantic bullet must not be rendered through Wingdings: ${drawn[0].font}`);
    assert.ok(drawn[1].font.includes("Hiragino Sans"), `expected East Asian fallback, got ${drawn[1].font}`);
    assert.ok(drawn[2].font.includes("Calibri"), `expected Latin fallback, got ${drawn[2].font}`);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("selects DrawingML latin, East Asian, and complex-script fonts per character", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
            <a:r><a:rPr lang="en-US" sz="1200"><a:ea typeface="East Asian Face"/><a:cs typeface="Complex Script Face"/></a:rPr><a:t>Aé中ع</a:t></a:r>
          </a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const [textBox] = await document.listObjects({ unitIndex: 0, textOnly: true });

    assert.equal(textBox.text, "Aé中ع");
    assert.deepEqual(
      textBox.fontRuns.map(({ start, end, authoredFamily }) => ({ start, end, authoredFamily })),
      [
        { start: 0, end: 2, authoredFamily: "Arial" },
        { start: 2, end: 3, authoredFamily: "East Asian Face" },
        { start: 3, end: 4, authoredFamily: "Complex Script Face" },
      ],
    );
  } finally {
    document?.close();
    engine.close();
  }
});

test("keeps semantic DrawingML Symbol bullets when the exact glyph font is unavailable", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
            <a:pPr><a:buFont typeface="Symbol"/><a:buChar char="&#xF0B7;"/></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1800"><a:latin typeface="Segoe UI"/></a:rPr><a:t>Item</a:t></a:r>
          </a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const drawn = [];
  let currentFont = "";
  const context = {
    globalAlpha: 1,
    letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    measureText(text) { return { width: text.length * 12 }; },
    fillText(text) { drawn.push({ text, font: currentFont }); },
    set fillStyle(_value) {}, set font(value) { currentFont = value; }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const [textBox] = await document.listObjects({ unitIndex: 0 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(textBox.text, "•\tItem");
    assert.deepEqual(
      textBox.fontRuns.map(({ start, end, authoredFamily }) => ({ start, end, authoredFamily })),
      [
        { start: 0, end: 2, authoredFamily: "Arial" },
        { start: 2, end: 6, authoredFamily: "Segoe UI" },
      ],
    );
    assert.deepEqual(drawn.map(({ text }) => text), ["•", "Item"]);
    assert.ok(!/\s"Symbol"(?:,|$)/u.test(drawn[0].font), `semantic bullet must not be rendered through Symbol: ${drawn[0].font}`);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("preserves authored PPTX shape bounds for spAutoFit text", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="10"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1524000" cy="190500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr lIns="190500" rIns="190500" tIns="95250" bIns="95250"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
            <a:p><a:pPr marL="0"/><a:r><a:rPr sz="1200"/><a:t>First</a:t></a:r></a:p>
            <a:p><a:pPr marL="0"/><a:r><a:rPr sz="1200"/><a:t>Second</a:t></a:r></a:p>
            <a:p><a:pPr marL="0"/><a:r><a:rPr sz="1200"/><a:t>Third</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const drawn = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    measureText(text) { return { width: text.length * 8 }; },
    fillText(text, x, y) { drawn.push({ text, x, y }); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const [textBox] = await document.listObjects({ unitIndex: 0 });
    frame = await document.render({ unitIndex: 0 });

    assert.ok(
      Math.abs(textBox.bounds.height - 20) < 0.01,
      `expected authored 20px height, got ${textBox.bounds.height}`,
    );
    assert.deepEqual(drawn.map(({ text }) => text), ["First", "Second", "Third"]);
    assert.equal(drawn[0].x, 120);
    const paragraphGap = drawn[1].y - drawn[0].y;
    assert.ok(Math.abs(paragraphGap - 19.2) < 0.01, `expected 19.2px paragraph gap, got ${paragraphGap}`);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders PPTX RTL columns, normAutoFit, overflow, text rotation, and WordArt warp", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="14"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="381000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody>
            <a:bodyPr rtlCol="1" numCol="2" spcCol="95250" horzOverflow="clip" vertOverflow="ellipsis" wrap="square" rot="900000">
              <a:normAutofit fontScale="80000" lnSpcReduction="10000"/>
              <a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp>
            </a:bodyPr>
            <a:lstStyle/>
            <a:p><a:r><a:rPr sz="1200"/><a:t>א</a:t></a:r></a:p>
            <a:p><a:r><a:rPr sz="1200"/><a:t>ב</a:t></a:r></a:p>
            <a:p><a:r><a:rPr sz="1200"/><a:t>ג</a:t></a:r></a:p>
            <a:p><a:r><a:rPr sz="1200"/><a:t>ד</a:t></a:r></a:p>
            <a:p><a:r><a:rPr sz="1200"/><a:t>ה</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const directions = [];
  const rotations = [];
  const drawn = [];
  let clips = 0;
  let currentFont = "";
  const context = {
    globalAlpha: 1,
    letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    rotate(value) { rotations.push(value); },
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() { clips += 1; },
    fillText(text, x, y) { drawn.push({ text, x, y, font: currentFont }); },
    measureText(text) { return { width: text.length * 8 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(value) { currentFont = value; },
    get font() { return currentFont; },
    set direction(value) { directions.push(value); },
    get direction() { return directions.at(-1) ?? "ltr"; },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(directions.includes("rtl"));
    assert.ok(rotations.some((value) => Math.abs(value - Math.PI / 12) < 0.0001));
    assert.ok(rotations.length > 1, "expected WordArt to add per-run deformation");
    assert.ok(clips >= 1);
    assert.ok(drawn.some(({ font }) => {
      const size = Number.parseFloat(/\s([\d.]+)px\s/u.exec(font)?.[1] ?? "NaN");
      // PowerPoint rounds persisted AutoFit sizes to whole points: 12pt × 80% → 10pt.
      // The real 16-page TextFittingComparisonWithMSO fixture covers this rounding in Core.
      return Math.abs(size - 10 * 96 / 72) < 0.001;
    }), `expected rounded 10pt AutoFit text, got ${drawn.map(({ font }) => font)}`);
    assert.ok(Math.max(...drawn.map(({ x }) => x)) - Math.min(...drawn.map(({ x }) => x)) > 90);
    assert.ok(drawn.some(({ text }) => text.endsWith("…")));
    assert.ok(document.diagnostics().some(({ fidelity, details }) => (
      fidelity === "approximate" && details?.feature === "wordart-warp"
    )));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("preserves PPTX title, per-paragraph, and table-cell text indentation", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="20"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="476250"/><a:ext cx="3810000" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>Title</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="21"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="1428750"/><a:ext cx="3810000" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr marL="0" indent="0"/><a:r><a:rPr sz="1200"/><a:t>Heading</a:t></a:r></a:p>
            <a:p><a:pPr marL="342900" indent="-342900"><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="1200"/><a:t>Item</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="22"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="3810000"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
            <a:tr h="952500"><a:tc><a:txBody><a:p><a:r><a:rPr sz="1200"/><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const drawn = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    measureText(text) { return { width: text.length * 10 }; },
    fillText(text, x, y) { drawn.push({ text, x, y }); },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    const x = (text) => drawn.find((entry) => entry.text === text)?.x;

    assert.ok(Math.abs(x("Title") - 109.6) < 0.01, `expected title x=109.6, got ${x("Title")}`);
    assert.ok(Math.abs(x("Heading") - 109.6) < 0.01, `expected heading x=109.6, got ${x("Heading")}`);
    assert.ok(Math.abs(x("•") - 109.6) < 0.01, `expected bullet x=109.6, got ${x("•")}`);
    assert.ok(Math.abs(x("Item") - 145.6) < 0.01, `expected item x=145.6, got ${x("Item")}`);
    assert.ok(Math.abs(x("Cell") - 109.6) < 0.01, `expected cell x=109.6, got ${x("Cell")}`);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("inherits PPTX title typography and effects from the slide master through the layout", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Inherited title</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
      </Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
          <p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></a:lstStyle><a:p/></p:txBody>
          </p:sp>
        </p:spTree></p:cSld>
      </p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
      </Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
          <p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p/></p:txBody>
          </p:sp>
        </p:spTree></p:cSld>
        <p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" b="1">
          <a:effectLst><a:outerShdw blurRad="38100" dist="25500" dir="5400000"><a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr></a:outerShdw></a:effectLst>
        </a:defRPr></a:lvl1pPr></p:titleStyle></p:txStyles>
      </p:sldMaster>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const drawn = [];
  const filters = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text, _x, y) { drawn.push({ text, font: this.font, y }); },
    measureText(text) { return { width: text.length * 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set filter(value) { filters.push(value); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(drawn.some(({ text }) => text === "Inherited") && drawn.some(({ text }) => text === "title"));
    assert.ok(drawn.every(({ font }) => font.startsWith("700 ")), `expected inherited bold title, got ${JSON.stringify(drawn)}`);
    assert.ok(drawn.every(({ y }) => y > 115), `expected inherited centered title, got ${JSON.stringify(drawn)}`);
    assert.ok(filters.includes("blur(2px)"), `expected inherited title shadow, got ${JSON.stringify(filters)}`);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("vertically centers PPTX text when bodyPr anchor is ctr", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="6"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>Centered</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const textPositions = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text, x, y) { textPositions.push([text, x, y]); },
    measureText() { return { width: 40 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(textPositions[0][0], "Centered");
    assert.ok(textPositions[0][2] > 135 && textPositions[0][2] < 145);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("resolves PPTX percentage line spacing from the actual run size", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
          <p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></a:lstStyle>
            <a:p><a:pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef></a:pPr><a:r><a:rPr sz="900"/><a:t>First</a:t></a:r></a:p>
            <a:p><a:pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="900"/><a:t>Second</a:t></a:r></a:p>
            <a:p><a:r><a:rPr sz="900"/><a:t>Natural</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const positions = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text, _x, y) { positions.push({ text, y }); },
    measureText(text) { return { width: text.length * 8 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(positions.map(({ text }) => text), ["First", "Second", "Natural"]);
    assert.ok(Math.abs(positions[0].y - 114.4) < 0.001);
    assert.ok(Math.abs(positions[1].y - positions[0].y - 28.8) < 0.001);
    assert.ok(Math.abs(positions[2].y - positions[1].y - 14.4) < 0.001);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DrawingML vert270 text as one counterclockwise rotated line", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="8"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="381000" cy="1524000"/></a:xfrm>
            <a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln>
          </p:spPr>
          <p:txBody>
            <a:bodyPr vert="vert270" lIns="0" rIns="0" tIns="0" bIns="0"/>
            <a:lstStyle/>
            <a:p><a:r><a:rPr sz="1800"/><a:t>03/2026</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const rotations = [];
  const texts = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    rotate(value) { rotations.push(value); },
    beginPath() {}, ellipse() {}, rect() {}, moveTo() {}, lineTo() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text, x, y) { texts.push({ text, x, y }); },
    measureText(text) { return { width: text.length * 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(texts.map(({ text }) => text).join(""), "03/2026");
    assert.ok(texts.every(({ y }) => y === texts[0].y));
    for (let index = 1; index < texts.length; index += 1) {
      assert.equal(texts[index].x, texts[index - 1].x + texts[index - 1].text.length * 10);
    }
    assert.ok(rotations.some((value) => Math.abs(value + Math.PI / 2) < 0.0001));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DrawingML eaVert Latin text as one clockwise rotated line", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="8"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="1981200" y="838200"/><a:ext cx="1170898" cy="4191000"/></a:xfrm>
            <a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln>
          </p:spPr>
          <p:txBody>
            <a:bodyPr vert="eaVert" wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>
            <a:lstStyle/>
            <a:p><a:r><a:rPr lang="en-US"/><a:t>This is Vertical text</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const rotations = [];
  const texts = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    rotate(value) { rotations.push(value); },
    beginPath() {}, ellipse() {}, rect() {}, moveTo() {}, lineTo() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text) { texts.push(text); },
    measureText(text) { return { width: text.length * 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(texts, ["This is Vertical text"]);
    assert.ok(rotations.some((value) => Math.abs(value - Math.PI / 2) < 0.0001));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders PPTX table-cell vert270 text as one counterclockwise rotated line", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="9"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="381000" cy="1524000"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblGrid><a:gridCol w="381000"/></a:tblGrid>
            <a:tr h="1524000"><a:tc>
              <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Register &amp; Evaluate</a:t></a:r></a:p></a:txBody>
              <a:tcPr vert="vert270" anchor="ctr" marL="0" marR="0" marT="0" marB="0"/>
            </a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const rotations = [];
  const texts = [];
  const context = {
    globalAlpha: 1,
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    rotate(value) { rotations.push(value); },
    beginPath() {}, ellipse() {}, rect() {}, moveTo() {}, lineTo() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text) { texts.push(text); },
    measureText(text) { return { width: text.length * 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set font(_value) {}, set letterSpacing(_value) {}, set textBaseline(_value) {},
    set textAlign(_value) {}, set direction(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(texts.join(""), "Register & Evaluate");
    assert.ok(rotations.some((value) => Math.abs(value + Math.PI / 2) < 0.0001));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders PPTX outerShdw color alpha blur distance and direction", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst><a:outerShdw blurRad="19050" dist="95250" dir="2700000"><a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const shadowColors = [];
  const shadowBlurs = [];
  const shadowOffsetsX = [];
  const shadowOffsetsY = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set shadowColor(value) { shadowColors.push(value); },
    set shadowBlur(value) { shadowBlurs.push(value); },
    set shadowOffsetX(value) { shadowOffsetsX.push(value); },
    set shadowOffsetY(value) { shadowOffsetsY.push(value); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(shadowColors.includes("rgba(17, 34, 51, 0.5019607843137255)"));
    assert.ok(shadowBlurs.includes(2));
    assert.ok(shadowOffsetsX.some((value) => Math.abs(value - Math.sqrt(50)) < 0.001));
    assert.ok(shadowOffsetsY.some((value) => Math.abs(value - Math.sqrt(50)) < 0.001));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders PPTX inner shadow, glow, reflection, and soft edge together", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="476250"/></a:xfrm>
            <a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            <a:effectLst>
              <a:innerShdw blurRad="9525" dist="9525" dir="0"><a:srgbClr val="FF0000"/></a:innerShdw>
              <a:glow rad="19050"><a:srgbClr val="00FF00"/></a:glow>
              <a:reflection stA="50000" endA="300" endPos="90000" blurRad="9525" dist="19050" sx="100000" sy="-50000"/>
              <a:softEdge rad="9525"/>
            </a:effectLst>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const shadowColors = [];
  const fillColors = [];
  const shadowBlurs = [];
  const filters = [];
  const scales = [];
  const gradients = [];
  let fills = 0;
  const context = {
    globalAlpha: 1,
    globalCompositeOperation: "source-over",
    imageSmoothingEnabled: true,
    imageSmoothingQuality: "high",
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, translate() {}, transform() {}, fillRect() {}, drawImage() {},
    getTransform() { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; },
    setTransform() {},
    getImageData(_x, _y, width, height) { return { width, height, data: new Uint8ClampedArray(width * height * 4) }; },
    putImageData() {},
    scale(x, y) { scales.push([x, y]); },
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() { fills += 1; }, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    createLinearGradient(...args) {
      const stops = [];
      gradients.push({ args, stops });
      return { addColorStop(offset, color) { stops.push([offset, color]); } };
    },
    set fillStyle(value) { fillColors.push(value); }, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set shadowColor(value) { shadowColors.push(value); },
    set shadowBlur(value) { shadowBlurs.push(value); }, set shadowOffsetX(_value) {}, set shadowOffsetY(_value) {},
    set filter(value) { filters.push(value); },
  };
  class FakeCanvas {
    constructor(width = 960, height = 720) { this.width = width; this.height = height; }
    getContext() { context.canvas = this; return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(fillColors.includes("rgba(255, 0, 0, 1)"), "inner shadow colors the inverted alpha mask");
    assert.ok(shadowColors.includes("rgba(0, 255, 0, 1)"), "glow uses the shared alpha mask and native Canvas shadow");
    assert.ok(shadowColors.includes("black") && shadowBlurs.includes(2), "soft edge feathers an alpha mask without blurring the source colors");
    assert.ok(scales.some(([x, y]) => x === 1 && y === -0.5));
    assert.ok(gradients.some(({ stops }) => (
      stops.some(([offset, value]) => offset === 0 && value === "rgba(0, 0, 0, 0.5)")
      && stops.some(([offset, value]) => (
        Math.abs(offset - 0.9) < 0.0001 && value.startsWith("rgba(0, 0, 0, 0.003")
      ))
    )), `expected reflection opacity gradient, got ${JSON.stringify(gradients)}`);
    assert.ok(fills > 0, "base shape remains painted while effects use raster masks");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DrawingML camera, light, material, extrusion, contour, and bevel effects", async () => {
  // Command recording checks affine composition; test-picture-transforms.mjs checks perspective pixels.
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="73"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="roundRect"/><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
            <a:scene3d>
              <a:camera prst="perspectiveContrastingRightFacing" zoom="75000" fov="0"><a:rot lat="600000" lon="1200000" rev="300000"/></a:camera>
              <a:lightRig rig="balanced" dir="br"><a:rot lat="300000" lon="600000" rev="0"/></a:lightRig>
            </a:scene3d>
            <a:sp3d z="9525" extrusionH="76200" contourW="9525" prstMaterial="metal">
              <a:bevelT w="19050" h="28575" prst="angle"/>
              <a:bevelB w="9525" h="19050" prst="softRound"/>
              <a:extrusionClr><a:srgbClr val="203864"/></a:extrusionClr>
              <a:contourClr><a:srgbClr val="D9E2F3"/></a:contourClr>
            </a:sp3d>
          </p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const translations = [];
  const transforms = [];
  const filters = [];
  const shadowColors = [];
  const maskedColors = [];
  const gradients = [];
  let fills = 0;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, transform(...values) { transforms.push(values); }, fillRect() {},
    getTransform() { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; }, setTransform() {}, drawImage() {},
    translate(x, y) { translations.push([x, y]); },
    scale() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() { fills += 1; }, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    createLinearGradient(...args) {
      const stops = [];
      gradients.push({ args, stops });
      return { addColorStop(offset, color) { stops.push([offset, color]); } };
    },
    set fillStyle(value) { if (this.globalCompositeOperation === "source-in") maskedColors.push(value); },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
    set shadowColor(value) { shadowColors.push(value); },
    set shadowBlur(_value) {}, set shadowOffsetX(_value) {}, set shadowOffsetY(_value) {},
    set filter(value) { filters.push(value); },
  };
  class FakeCanvas {
    constructor(width = 960, height = 720) { this.width = width; this.height = height; }
    getContext() { context.canvas = this; return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(document.diagnostics().some(({ details }) => details?.feature === "three-dimensional-effects"), false);
    assert.ok(translations.some(([x, y]) => Math.abs(x) + Math.abs(y) > 0.1));
    assert.ok(transforms.some(([a, b, c, d]) => (
      Math.abs(a - 0.70596982) < 0.0001
        && Math.abs(b - (-0.01705083)) < 0.0001
        && Math.abs(c - 0.06437374) < 0.0001
        && Math.abs(d - 0.73579520) < 0.0001
    )));
    assert.ok(filters.some((value) => value.includes("brightness") && value.includes("contrast")));
    assert.ok(gradients.some(({ stops }) => (
      stops.some(([, value]) => value.startsWith("rgba(255, 255, 255,"))
        && stops.some(([, value]) => value.startsWith("rgba(0, 0, 0,"))
    )));
    assert.ok(maskedColors.includes("rgba(32, 56, 100, 1)"));
    assert.ok(shadowColors.includes("rgba(217, 226, 243, 1)"));
    assert.ok(fills >= 6);
    assert.equal(frame.renderedObjectCount, 1);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("keeps flat DrawingML text outside extrusion while applying the 3D backdrop to the shape", async () => {
  // Command recording checks affine composition; test-picture-transforms.mjs checks perspective pixels.
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="74"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="roundRect"/><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
            <a:scene3d><a:camera prst="perspectiveRight" fov="0"><a:rot lat="300000" lon="600000" rev="0"/></a:camera><a:lightRig rig="balanced" dir="br"/>
              <a:backdrop><a:anchor x="9525" y="19050" z="28575"/><a:norm dx="20000" dy="10000" dz="50000"/><a:up dx="5000" dy="-20000" dz="50000"/></a:backdrop>
            </a:scene3d>
            <a:sp3d extrusionH="76200" prstMaterial="plastic"/>
          </p:spPr>
          <p:txBody><a:bodyPr><a:flatTx z="38100"/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Once</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const texts = [];
  const transforms = [];
  const stateStack = [];
  let cameraTransformed = false;
  let textWasCameraTransformed = false;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() { stateStack.push(cameraTransformed); },
    restore() { cameraTransformed = stateStack.pop() ?? false; },
    scale(x, y) { if (Math.abs(x - y) > 0.001) cameraTransformed = true; }, translate() {},
    transform(...values) { transforms.push(values); cameraTransformed = true; },
    fillRect() {},
    beginPath() {}, ellipse() {}, rect() {}, moveTo() {}, lineTo() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text) { texts.push(text); textWasCameraTransformed ||= cameraTransformed; },
    measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {}, set filter(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.deepEqual(texts, ["Once"]);
    assert.ok(transforms.some(([, b, c]) => Math.abs(b) + Math.abs(c) > 0.001));
    assert.equal(textWasCameraTransformed, false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("opens a PPTX table as source-mapped table and cell objects", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="12"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="1905000"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblGrid><a:gridCol w="1905000"/><a:gridCol w="1905000"/></a:tblGrid>
            <a:tr h="952500">
              <a:tc><a:txBody><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill></a:tcPr></a:tc>
              <a:tc><a:txBody><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
            </a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 150, y: 220 });

    assert.equal(hits[0].object.type, "cell");
    assert.equal(hits[0].object.text, "A1");
    assert.equal(hits[0].object.source.shapeId, 12);
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, 0);
    assert.equal(hits[0].ancestors.length, 1);
    assert.equal(hits[0].ancestors[0].type, "table");
    assert.deepEqual(hits[0].object.bounds, { x: 100, y: 200, width: 200, height: 100 });
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders theme-driven fills from built-in PPTX table styles", async () => {
  const styleId = "{7DF18680-E054-41AD-8BC1-D1AEF772440D}";
  const cells = Array.from({ length: 9 }, (_, index) =>
    `<a:tc><a:txBody><a:p><a:r><a:rPr sz="1200"/><a:t>C${index + 1}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`);
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="16"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1428750"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblPr firstRow="1" firstCol="1" bandRow="1"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="1270000"/><a:gridCol w="1270000"/><a:gridCol w="1270000"/></a:tblGrid>
            <a:tr h="476250">${cells.slice(0, 3).join("")}</a:tr>
            <a:tr h="476250">${cells.slice(3, 6).join("")}</a:tr>
            <a:tr h="476250">${cells.slice(6, 9).join("")}</a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/theme/theme1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>
        <a:clrScheme name="Table Test">
          <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
          <a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
          <a:accent1><a:srgbClr val="111111"/></a:accent1><a:accent2><a:srgbClr val="222222"/></a:accent2>
          <a:accent3><a:srgbClr val="333333"/></a:accent3><a:accent4><a:srgbClr val="444444"/></a:accent4>
          <a:accent5><a:srgbClr val="19226D"/></a:accent5><a:accent6><a:srgbClr val="666666"/></a:accent6>
          <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Table Test"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
      </a:themeElements></a:theme>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const textColors = [];
  let currentFill = "";
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() { fills.push(currentFill); }, stroke() {}, clip() {},
    fillText() { textColors.push(currentFill); }, measureText() { return { width: 10 }; },
    set fillStyle(value) { currentFill = value; },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(fills, [
      "rgba(25, 34, 109, 1)", "rgba(25, 34, 109, 1)", "rgba(25, 34, 109, 1)",
      "rgba(25, 34, 109, 1)", "rgba(204, 204, 212, 1)", "rgba(204, 204, 212, 1)",
      "rgba(25, 34, 109, 1)", "rgba(231, 232, 235, 1)", "rgba(231, 232, 235, 1)",
    ]);
    // MS-OE376 Dark Style 1 Accent 5 replaces dk1 with accent5, including body text.
    assert.deepEqual(textColors, [
      "rgba(255, 255, 255, 1)", "rgba(255, 255, 255, 1)", "rgba(255, 255, 255, 1)",
      "rgba(255, 255, 255, 1)", "rgba(25, 34, 109, 1)", "rgba(25, 34, 109, 1)",
      "rgba(255, 255, 255, 1)", "rgba(25, 34, 109, 1)", "rgba(25, 34, 109, 1)",
    ]);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders package-defined PPTX table style fills and text colors", async () => {
  const styleId = "{11111111-2222-3333-4444-555555555555}";
  const cells = Array.from({ length: 9 }, (_, index) =>
    `<a:tc><a:txBody><a:p><a:r><a:rPr sz="1200"/><a:t>C${index + 1}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`);
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="31"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1428750"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblPr firstRow="1" firstCol="1" bandRow="1"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="1270000"/><a:gridCol w="1270000"/><a:gridCol w="1270000"/></a:tblGrid>
            <a:tr h="476250">${cells.slice(0, 3).join("")}</a:tr>
            <a:tr h="476250">${cells.slice(3, 6).join("")}</a:tr>
            <a:tr h="476250">${cells.slice(6, 9).join("")}</a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="32"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="2857500"/><a:ext cx="952500" cy="476250"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblPr><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="952500"/></a:tblGrid>
            <a:tr h="476250"><a:tc><a:txBody><a:p><a:r><a:rPr sz="1200"/><a:t>Whole</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/theme/theme1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>
        <a:clrScheme name="Table Style Test">
          <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
          <a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
          <a:accent1><a:srgbClr val="111111"/></a:accent1><a:accent2><a:srgbClr val="222222"/></a:accent2>
          <a:accent3><a:srgbClr val="333333"/></a:accent3><a:accent4><a:srgbClr val="445566"/></a:accent4>
          <a:accent5><a:srgbClr val="19226D"/></a:accent5><a:accent6><a:srgbClr val="666666"/></a:accent6>
          <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Table Style Test"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
      </a:themeElements></a:theme>`,
    "ppt/tableStyles.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:tblStyle styleId="${styleId}" styleName="Complete package style">
          <a:wholeTbl>
            <a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle>
            <a:tcStyle><a:tcBdr><a:left><a:ln w="19050"><a:solidFill><a:srgbClr val="034EA2"/></a:solidFill></a:ln></a:left></a:tcBdr>
              <a:fill><a:solidFill><a:schemeClr val="accent5"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill>
            </a:tcStyle>
          </a:wholeTbl>
          <a:band1H><a:tcTxStyle><a:schemeClr val="accent1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent5"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>
          <a:band2H><a:tcTxStyle><a:schemeClr val="accent2"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></a:fill></a:tcStyle></a:band2H>
          <a:firstRow><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent5"/></a:solidFill></a:fill></a:tcStyle></a:firstRow>
          <a:firstCol><a:tcTxStyle b="on"><a:schemeClr val="accent4"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent4"/></a:solidFill></a:fill></a:tcStyle></a:firstCol>
        </a:tblStyle>
      </a:tblStyleLst>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const textColors = [];
  let currentFill = "";
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() { fills.push(currentFill); }, stroke() {}, clip() {},
    fillText() { textColors.push(currentFill); }, measureText() { return { width: 10 }; },
    set fillStyle(value) { currentFill = value; },
    set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.deepEqual(fills, [
      "rgba(25, 34, 109, 1)", "rgba(25, 34, 109, 1)", "rgba(25, 34, 109, 1)",
      "rgba(68, 85, 102, 1)", "rgba(204, 204, 212, 1)", "rgba(204, 204, 212, 1)",
      "rgba(68, 85, 102, 1)", "rgba(170, 187, 204, 1)", "rgba(170, 187, 204, 1)",
      "rgba(231, 232, 235, 1)",
    ]);
    assert.deepEqual(textColors, [
      "rgba(255, 255, 255, 1)", "rgba(255, 255, 255, 1)", "rgba(255, 255, 255, 1)",
      "rgba(68, 85, 102, 1)", "rgba(17, 17, 17, 1)", "rgba(17, 17, 17, 1)",
      "rgba(68, 85, 102, 1)", "rgba(34, 34, 34, 1)", "rgba(34, 34, 34, 1)",
      "rgba(0, 0, 0, 1)",
    ]);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders PPTX table borders from the referenced table style", async () => {
  const styleId = "{0505E3EF-67EA-436B-97B2-0124C06EBD24}";
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="15"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblPr><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="1905000"/></a:tblGrid>
            <a:tr h="952500"><a:tc><a:txBody><a:p><a:r><a:t>Styled</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/tableStyles.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:tblStyle styleId="${styleId}" styleName="Blue border">
          <a:wholeTbl><a:tcStyle><a:tcBdr>
            <a:left><a:ln w="19050"><a:solidFill><a:srgbClr val="034EA2"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></a:left>
          </a:tcBdr></a:tcStyle></a:wholeTbl>
        </a:tblStyle>
      </a:tblStyleLst>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const strokes = [];
  const widths = [];
  const dashes = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    setLineDash(value) { dashes.push(value); },
    set fillStyle(_value) {},
    set strokeStyle(value) { strokes.push(value); },
    set lineWidth(value) { widths.push(value); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(strokes.includes("rgba(3, 78, 162, 1)"));
    assert.ok(widths.includes(2));
    assert.ok(dashes.some((dash) => dash.length === 2 && dash[0] === 6 && dash[1] === 2));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("keeps PPTX cell edge borders when diagonal borders are noFill", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="18"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblGrid><a:gridCol w="1905000"/></a:tblGrid>
            <a:tr h="952500"><a:tc><a:txBody><a:p><a:r><a:rPr sz="1200"/><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr>
              <a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill>
              <a:lnL w="9525"><a:solidFill><a:srgbClr val="1F2279"/></a:solidFill></a:lnL>
              <a:lnR w="9525"><a:solidFill><a:srgbClr val="1F2279"/></a:solidFill></a:lnR>
              <a:lnT w="9525"><a:solidFill><a:srgbClr val="1F2279"/></a:solidFill></a:lnT>
              <a:lnB w="9525"><a:solidFill><a:srgbClr val="1F2279"/></a:solidFill></a:lnB>
              <a:lnTlToBr><a:noFill/></a:lnTlToBr>
              <a:lnBlToTr><a:noFill/></a:lnBlToTr>
            </a:tcPr></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const strokes = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(value) { strokes.push(value); }, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(strokes.includes("rgba(31, 34, 121, 1)"));
    assert.ok(fills.includes("rgba(255, 242, 204, 1)"));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("treats authored PPTX table row heights as content-aware minimums", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="14"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="556260"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
            <a:tr h="190500"><a:tc><a:txBody><a:p><a:r><a:rPr sz="1200"/><a:t>Authored minimum</a:t></a:r></a:p><a:p><a:r><a:rPr sz="1200"/><a:t>second line</a:t></a:r></a:p></a:txBody><a:tcPr marT="0" marB="0"/></a:tc></a:tr>
            <a:tr h="190500"><a:tc><a:txBody><a:p/></a:txBody><a:tcPr marT="0" marB="0"/></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const objects = await document.listObjects({ unitIndex: 0 });
    const table = objects.find(({ type }) => type === "table");
    const cells = objects.filter(({ type }) => type === "cell");

    assert.ok(table);
    assert.equal(cells.length, 2);
    assert.ok(Math.abs(table.bounds.height - 58.4) < 0.01);
    assert.ok(Math.abs(cells[0].bounds.height - 38.4) < 0.01);
    assert.ok(Math.abs(cells[1].bounds.height - 20) < 0.01);
    assert.ok(Math.abs(cells[1].bounds.y - 138.4) < 0.01);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a merged PPTX table cell with authored border and text styling", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="13"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="1905000"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblGrid><a:gridCol w="1905000"/><a:gridCol w="1905000"/></a:tblGrid>
            <a:tr h="476250">
              <a:tc gridSpan="2" rowSpan="2"><a:txBody><a:p><a:r><a:rPr sz="1200" b="1" i="1"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>Merged</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:lnL w="19050"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:lnL></a:tcPr></a:tc>
              <a:tc hMerge="1" rowSpan="2"><a:txBody/><a:tcPr/></a:tc>
            </a:tr>
            <a:tr h="476250">
              <a:tc vMerge="1" gridSpan="2"><a:txBody/><a:tcPr/></a:tc>
              <a:tc hMerge="1" vMerge="1"><a:txBody/><a:tcPr/></a:tc>
            </a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const strokes = [];
  const widths = [];
  const fonts = [];
  const context = {
    globalAlpha: 1,
    letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); },
    set strokeStyle(value) { strokes.push(value); },
    set lineWidth(value) { widths.push(value); },
    set font(value) { fonts.push(value); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 450, y: 275 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.type, "cell");
    assert.equal(hits[0].object.text, "Merged");
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, 0);
    assert.deepEqual(hits[0].object.bounds, { x: 100, y: 200, width: 400, height: 100 });
    assert.ok(fills.includes("rgba(255, 0, 0, 1)"));
    assert.ok(fills.includes("rgba(0, 255, 0, 1)"));
    assert.ok(strokes.includes("rgba(0, 0, 255, 1)"));
    assert.ok(widths.includes(2));
    assert.ok(fonts.some((font) => font.includes("italic 700") && font.includes("Aptos")));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("keeps cells after a gridSpan and hMerge continuation in their authored columns", async () => {
  const styleId = "{7B3CAFF9-2D80-4630-9455-82A25E37F7C1}";
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="14"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="476250"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblPr><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="952500"/><a:gridCol w="952500"/><a:gridCol w="952500"/><a:gridCol w="952500"/></a:tblGrid>
            <a:tr h="476250">
              <a:tc gridSpan="2"><a:txBody><a:p><a:r><a:t>Merged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
              <a:tc hMerge="1"><a:txBody/><a:tcPr/></a:tc>
              <a:tc><a:txBody><a:p><a:r><a:t>Unit</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
              <a:tc><a:txBody><a:p><a:r><a:t>Remarks fit here</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
            </a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/tableStyles.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:tblStyle styleId="${styleId}" styleName="Merge column border test">
          <a:wholeTbl><a:tcStyle><a:tcBdr>
            <a:insideV><a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:insideV>
            <a:right><a:ln w="38100"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></a:right>
          </a:tcBdr></a:tcStyle></a:wholeTbl>
        </a:tblStyle>
      </a:tblStyleLst>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const strokes = [];
  let path = [];
  let strokeStyle = "";
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() { path = []; }, ellipse() {},
    moveTo(x, y) { path.push([x, y]); }, lineTo(x, y) { path.push([x, y]); },
    rect() {}, bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {},
    stroke() { strokes.push({ path: path.map((point) => [...point]), strokeStyle }); },
    clip() {}, fillText() {}, measureText() { return { width: 10 }; }, setLineDash() {},
    set fillStyle(_value) {}, set strokeStyle(value) { strokeStyle = value; }, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const cells = (await document.listObjects({ unitIndex: 0 })).filter(({ type }) => type === "cell");
    assert.deepEqual(cells.map(({ text, bounds, source }) => ({
      text,
      x: bounds.x,
      width: bounds.width,
      column: source.column,
    })), [
      { text: "Merged", x: 100, width: 200, column: 0 },
      { text: "Unit", x: 300, width: 100, column: 2 },
      { text: "Remarks fit here", x: 400, width: 100, column: 3 },
    ]);
    frame = await document.render({ unitIndex: 0 });
    const verticalBorderAt = (x) => strokes.find(({ path: points }) => (
      points.length >= 2 && points.every(([pointX]) => Math.abs(pointX - x) < 0.01)
    ));
    assert.equal(verticalBorderAt(400)?.strokeStyle, "rgba(255, 0, 0, 1)");
    assert.equal(verticalBorderAt(500)?.strokeStyle, "rgba(0, 0, 255, 1)");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("treats vertically merged PPTX cell content as a row-range constraint", async () => {
  const paragraphs = Array.from({ length: 5 }, (_, index) => (
    `<a:p><a:r><a:rPr sz="1000"/><a:t>Line ${index + 1}</a:t></a:r></a:p>`
  )).join("");
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="13"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblGrid><a:gridCol w="1905000"/><a:gridCol w="1905000"/></a:tblGrid>
            <a:tr h="762000"><a:tc rowSpan="2"><a:txBody>${paragraphs}</a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
            <a:tr h="95250"><a:tc vMerge="1"><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const cells = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ type }) => type === "cell");

    assert.equal(cells.length, 3);
    const firstRow = cells.find(({ source }) => source.row === 0 && source.column === 1);
    const secondRow = cells.find(({ source }) => source.row === 1 && source.column === 1);
    const merged = cells.find(({ source }) => source.row === 0 && source.column === 0);
    assert.ok(firstRow);
    assert.ok(secondRow);
    assert.ok(merged);
    assert.ok(Math.abs(firstRow.bounds.height - 85) < 0.01);
    assert.ok(Math.abs(secondRow.bounds.height - 15) < 0.01);
    assert.ok(Math.abs(merged.bounds.height - 100) < 0.01);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders DrawingML wavy underlines as waves", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" u="wavy"/><a:t>Wavy underline</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  let curves = 0;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {},
    bezierCurveTo() { curves += 1; }, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {},
    clip() {}, fillText() {}, measureText() {
      return { width: 120, actualBoundingBoxAscent: 18, actualBoundingBoxDescent: 5 };
    },
    setLineDash() {},
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(curves > 0, "wavy underline should contain curved segments");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("keeps DrawingML double wavy underlines light at large font sizes", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="5715000" cy="1905000"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="5400" u="wavyDbl"/><a:t>Underline</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  let lineWidth = 1;
  let maximumLineWidth = 0;
  let waveY = 0;
  let maximumAmplitude = 0;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo(_x, y) { waveY = y; }, lineTo() {}, rect() {},
    bezierCurveTo(_x1, y1, _x2, y2) {
      maximumAmplitude = Math.max(maximumAmplitude, Math.abs(y1 - waveY), Math.abs(y2 - waveY));
    },
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 360, actualBoundingBoxAscent: 48, actualBoundingBoxDescent: 12 }; },
    setLineDash() {},
    set fillStyle(_value) {}, set strokeStyle(_value) {},
    get lineWidth() { return lineWidth; },
    set lineWidth(value) { lineWidth = value; maximumLineWidth = Math.max(maximumLineWidth, value); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(maximumLineWidth <= 1.5, `wavy underline width ${maximumLineWidth} should stay light`);
    assert.ok(maximumAmplitude <= 1.5, `wavy underline amplitude ${maximumAmplitude} should stay compact`);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DrawingML dot-dash underlines with alternating dot and dash segments", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" u="dotDash"/><a:t>Dot dash underline</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const dashes = [];
  let lineWidth = 1;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {},
    clip() {}, fillText() {}, measureText() {
      return { width: 180, actualBoundingBoxAscent: 18, actualBoundingBoxDescent: 5 };
    },
    setLineDash(value) { dashes.push(value); },
    set fillStyle(_value) {}, set strokeStyle(_value) {},
    get lineWidth() { return lineWidth; }, set lineWidth(value) { lineWidth = value; },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(dashes.some((dash) => dash.length === 4
      && dash.every((value, index) => value === [8, 4, 2, 4][index])));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DrawingML dotted heavy underlines as heavy dots", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" u="dottedHeavy"/><a:t>Dotted heavy underline</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5"/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="1905000"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="3810000"/></a:tblGrid><a:tr h="952500"><a:tc>
              <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" u="dottedHeavy"/><a:t>Table underline</a:t></a:r></a:p></a:txBody><a:tcPr/>
            </a:tc></a:tr></a:tbl>
          </a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const dashes = [];
  const widths = [];
  let lineWidth = 1;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {},
    clip() {}, fillText() {}, measureText() {
      return { width: 120, actualBoundingBoxAscent: 18, actualBoundingBoxDescent: 5 };
    },
    setLineDash(value) { dashes.push(value); },
    set fillStyle(_value) {}, set strokeStyle(_value) {},
    get lineWidth() { return lineWidth; },
    set lineWidth(value) { lineWidth = value; widths.push(value); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(dashes.filter((dash) => dash.length > 0).length >= 2,
      "shape and table dotted underlines should install dash patterns");
    assert.ok(Math.max(...widths) >= 2, "heavy underline should remain twice the normal weight");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DrawingML double underlines as two lines", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" u="dbl"/><a:t>Double underline</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  let moveY;
  let lineWidth = 1;
  const horizontalYs = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo(_x, y) { moveY = y; },
    lineTo(_x, y) { if (y === moveY) horizontalYs.push(y); }, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {},
    stroke() {},
    clip() {}, fillText() {}, measureText() {
      return { width: 120, actualBoundingBoxAscent: 18, actualBoundingBoxDescent: 5 };
    },
    setLineDash() {},
    set fillStyle(_value) {}, set strokeStyle(_value) {},
    get lineWidth() { return lineWidth; }, set lineWidth(value) { lineWidth = value; },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.equal(new Set(horizontalYs).size, 2,
      "double underline should paint at two distinct vertical positions");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("centers DrawingML strikethrough on the glyph ink bounds", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" strike="sngStrike"/><a:t>strikethrough</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  let textY;
  let moveY;
  let strikeY;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo(_x, y) { moveY = y; },
    lineTo(_x, y) { if (textY !== undefined && y === moveY) strikeY = y; }, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {},
    clip() {}, fillText(_text, _x, y) { textY = y; }, measureText() {
      return {
        width: 120,
        actualBoundingBoxAscent: 18,
        actualBoundingBoxDescent: 6,
        fontBoundingBoxAscent: 20,
        fontBoundingBoxDescent: 5,
      };
    },
    setLineDash() {},
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.equal(strikeY, textY + (6 - 18) / 2);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DrawingML double strikethrough as two centered lines", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" strike="dblStrike"/><a:t>Double strikethrough</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const horizontalYs = [];
  let moveY;
  let textY;
  let lineWidth = 1;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo(_x, y) { moveY = y; },
    lineTo(_x, y) { if (y === moveY) horizontalYs.push(y); }, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {},
    clip() {}, fillText(_text, _x, y) { textY = y; }, measureText() {
      return {
        width: 180,
        actualBoundingBoxAscent: 18,
        actualBoundingBoxDescent: 6,
        fontBoundingBoxAscent: 20,
        fontBoundingBoxDescent: 5,
      };
    },
    setLineDash() {},
    set fillStyle(_value) {}, set strokeStyle(_value) {},
    get lineWidth() { return lineWidth; }, set lineWidth(value) { lineWidth = value; },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    const strikeYs = new Set(horizontalYs.filter((y) => y >= textY - 18 && y <= textY + 6));
    assert.equal(strikeYs.size, 2, "double strikethrough should paint two centered lines");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("distributes DrawingML paragraph text across the available width", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp><p:nvSpPr><p:cNvPr id="4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="dist"/><a:r><a:rPr sz="2400"/><a:t>Distributed Text</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const paintedLetterSpacing = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "0px", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText() { paintedLetterSpacing.push(Number.parseFloat(context.letterSpacing)); },
    measureText(text) {
      return {
        width: text.length * 10,
        actualBoundingBoxAscent: 18,
        actualBoundingBoxDescent: 5,
      };
    },
    setLineDash() {},
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(paintedLetterSpacing.some((spacing) => spacing > 0),
      "distributed text should expand grapheme spacing");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("uses PPTX table paragraph spacing and bullet layout when sizing rows", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="13"/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
          <a:graphic><a:graphicData><a:tbl>
            <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
            <a:tr h="0"><a:tc><a:txBody><a:p>
              <a:pPr marL="171450" indent="-171450"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>
              <a:r><a:rPr sz="1000"/><a:t>Spaced bullet</a:t></a:r>
            </a:p></a:txBody><a:tcPr/></a:tc></a:tr>
            <a:tr h="0"><a:tc><a:txBody><a:p><a:r><a:rPr sz="1000"/><a:t>Natural</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const cells = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ type }) => type === "cell")
      .sort((left, right) => left.source.row - right.source.row);

    assert.equal(cells.length, 2);
    assert.equal(cells[0].text, "•\tSpaced bullet");
    assert.ok(Math.abs(cells[0].bounds.height - 104) < 0.01);
    assert.ok(Math.abs(cells[1].bounds.height - 96) < 0.01);
  } finally {
    document?.close();
    engine.close();
  }
});

test("uses each PPTX table paragraph run size for its line box", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="14"/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData><a:tbl>
            <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
            <a:tr h="952500"><a:tc><a:txBody>
              <a:p><a:r><a:rPr sz="1400"/><a:t>Heading</a:t></a:r></a:p>
              <a:p><a:r><a:rPr sz="1000"/><a:t>Body one</a:t></a:r></a:p>
              <a:p><a:r><a:rPr sz="1000"/><a:t>Body two</a:t></a:r></a:p>
            </a:txBody><a:tcPr marL="0" marR="0" marT="0" marB="0"/></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const positions = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, rotate() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text, _x, y) { positions.push({ text, y }); },
    measureText(text) { return { width: text.length * 8 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    const lineStarts = positions.filter(({ text }) => text === "Heading" || text === "Body");
    assert.equal(lineStarts.length, 3);
    assert.ok(Math.abs(lineStarts[1].y - lineStarts[0].y - 22.4) < 0.001);
    assert.ok(Math.abs(lineStarts[2].y - lineStarts[1].y - 16) < 0.001);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("keeps unsupported PPTX charts visible as mapped static placeholders", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="14"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></p:xfrm><a:graphic><a:graphicData><c:chart/></a:graphicData></a:graphic></p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 110, y: 110 });

    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.type, "unknown");
    assert.equal(hits[0].object.text, "Chart");
    assert.equal(hits[0].object.source.shapeId, 14);
    assert.ok(document.diagnostics().some(({ message, fidelity }) => (
      message.includes("static placeholder") && fidelity === "approximate"
    )));
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a cached PPTX bar chart as source-mapped display objects", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="15"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="2857500"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
    "ppt/charts/chart1.xml": `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="47"/><c:chart><c:plotArea><c:barChart><c:ser>
        <c:tx><c:v>Revenue</c:v></c:tx>
        <c:spPr><a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill></c:spPr>
        <c:dLbls><c:dLbl><c:idx val="2"/><c:tx><c:rich><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Custom</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:noFill/><a:ln w="22860"><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:ln></c:spPr></c:dLbl></c:dLbls>
        <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>20</c:v></c:pt><c:pt idx="1"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const strokes = [];
  const texts = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText(text) { texts.push(text); },
    measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); },
    set strokeStyle(value) { strokes.push(value); }, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 200, y: 200 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.type, "shape");
    assert.equal(hits[0].object.source.part, "ppt/charts/chart1.xml");
    assert.equal(hits[0].object.source.shapeId, 15);
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, 0);
    assert.ok(fills.includes("rgba(63, 63, 63, 1)"));
    assert.ok(fills.includes("rgba(91, 155, 213, 1)"));
    assert.ok(fills.includes("rgba(255, 0, 0, 1)"));
    assert.ok(strokes.includes("rgba(255, 255, 0, 1)"));
    assert.ok(texts.includes("Custom"));
    assert.ok(frame.renderedObjectCount >= 4);
    assert.equal(document.diagnostics().some(({ message }) => message.includes("chart content uses a source-mapped static placeholder")), false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders a PPTX chart title gradient fill", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="16"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="2857500"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
    "ppt/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Title</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="D9E2F3"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></c:spPr></c:title><c:plotArea><c:barChart><c:ser><c:tx><c:v>Series</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea><c:legend><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="32CD32"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></c:spPr></c:legend></c:chart><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="4169E1"/></a:gs></a:gsLst><a:lin ang="3600000"/></a:gradFill></c:spPr></c:chartSpace>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const gradients = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    createLinearGradient() {
      const stops = [];
      gradients.push(stops);
      return { addColorStop(offset, color) { stops.push([offset, color]); } };
    },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const titleFill = (await document.listObjects({ unitIndex: 0 })).find(({ bounds, source, type }) => (
      type === "shape"
      && source.part === "ppt/charts/chart1.xml"
      && bounds.y === 108
      && bounds.height === 28
    ));
    frame = await document.render({ unitIndex: 0 });
    assert.ok(titleFill);
    assert.ok(titleFill.bounds.width < 200, `chart title fill is too wide: ${titleFill.bounds.width}`);
    assert.deepEqual(gradients, [
      [[0, "rgba(255, 255, 0, 1)"], [1, "rgba(65, 105, 225, 1)"]],
      // The classic chart outline also binds the area fill for stroke gaps.
      [[0, "rgba(255, 255, 0, 1)"], [1, "rgba(65, 105, 225, 1)"]],
      [[0, "rgba(255, 255, 255, 1)"], [1, "rgba(217, 226, 243, 1)"]],
      [[0, "rgba(50, 205, 50, 1)"], [1, "rgba(255, 255, 255, 1)"]],
    ]);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders a cached PPTX line chart with source-mapped series segments", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="17"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="2857500"/></p:xfrm><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>`,
    "ppt/charts/_rels/chart2.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" Target="../theme/themeOverride1.xml"/></Relationships>`,
    "ppt/theme/themeOverride1.xml": `<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:clrScheme name="Chart"><a:accent2><a:srgbClr val="ED7D31"/></a:accent2></a:clrScheme></a:themeOverride>`,
    "ppt/charts/chart2.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart><c:ser>
      <c:tx><c:v>Revenue</c:v></c:tx>
      <c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr>
      <c:cat><c:strLit><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strLit></c:cat>
      <c:val><c:numLit><c:pt idx="0"><c:v>20</c:v></c:pt><c:pt idx="1"><c:v>10</c:v></c:pt></c:numLit></c:val>
    </c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx><c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable></c:plotArea></c:chart></c:chartSpace>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const strokes = [];
  const linePoints = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    lineTo(x, y) { linePoints.push([x, y]); },
    measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(value) { strokes.push(value); }, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const chartObjects = await document.listObjects({ unitIndex: 0 });
    const chartText = chartObjects
      .filter(({ source }) => source.part === "ppt/charts/chart2.xml")
      .map(({ text }) => text)
      .filter(Boolean);
    const hits = await document.hitTest({ unitIndex: 0, x: 300, y: 180 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.type, "shape");
    assert.equal(hits[0].object.source.part, "ppt/charts/chart2.xml");
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, undefined, "a joined series path does not identify one data point");
    assert.ok(chartText.includes("Revenue"), JSON.stringify(chartText));
    assert.ok(chartText.includes("Q1") && chartText.includes("Q2"), JSON.stringify(chartText));
    assert.ok(!chartObjects.some(({ type, bounds }) => (
      type === "shape"
      && Math.abs(bounds.x - 112) < 0.01
      && Math.abs(bounds.y - 325) < 0.01
      && Math.abs(bounds.width - 56) < 0.01
      && bounds.height < 0.02
    )), "data-table outline must not render a series-label top border");
    assert.ok(chartObjects.some(({ type, bounds }) => (
      type === "shape"
      && Math.abs(bounds.x - 168) < 0.01
      && Math.abs(bounds.y - 325) < 0.01
      && Math.abs(bounds.width - 320) < 0.01
      && bounds.height < 0.02
    )), "data-table chart must retain its horizontal axis above the category header");
    assert.ok(strokes.includes("rgba(237, 125, 49, 1)"));
    assert.ok(linePoints.some(([x, y]) => x > 390 && y > 320 && y < 330), JSON.stringify(linePoints));
    assert.ok(frame.renderedObjectCount >= 4);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders mixed PPTX bar and line chart groups on their own axes", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="19"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="2857500"/></p:xfrm><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart4.xml"/></Relationships>`,
    "ppt/charts/chart4.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>
      <c:barChart><c:ser><c:spPr><a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>200</c:v></c:pt><c:pt idx="1"><c:v>400</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
      <c:lineChart><c:ser><c:spPr><a:ln w="38100"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>0.025</c:v></c:pt><c:pt idx="1"><c:v>0.05</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="3"/><c:axId val="4"/></c:lineChart>
    </c:plotArea></c:chart></c:chartSpace>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const chartShapes = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "ppt/charts/chart4.xml");
    const seriesRows = new Set(chartShapes.map(({ source }) => source.row).filter((row) => row !== undefined));

    assert.deepEqual([...seriesRows].sort(), [0, 1]);
    assert.ok(chartShapes.some(({ source, bounds }) => source.row === 0 && bounds.width > 0 && bounds.height > 0));
    assert.ok(chartShapes.some(({ source, bounds }) => source.row === 1 && bounds.width > 0 && bounds.height > 0));
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a cached PPTX pie chart as source-mapped slice paths", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="18"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="2857500"/></p:xfrm><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/></Relationships>`,
    "ppt/charts/chart3.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.1"/><c:w val="0.8"/><c:h val="0.8"/></c:manualLayout></c:layout><c:pieChart><c:ser>
      <c:dPt><c:idx val="1"/><c:explosion val="17"/></c:dPt>
      <c:dLbls><c:dLbl><c:idx val="0"/><c:layout><c:manualLayout><c:x val="-0.25"/><c:y val="0.002"/><c:w val="0.38"/><c:h val="0.077"/></c:manualLayout></c:layout></c:dLbl><c:dLbl><c:idx val="1"/><c:layout><c:manualLayout><c:x val="0.189"/><c:y val="-0.242"/><c:w val="0.367"/><c:h val="0.154"/></c:manualLayout></c:layout></c:dLbl><c:txPr><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="404040"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:dLblPos val="bestFit"/><c:showVal val="1"/><c:showCatName val="1"/></c:dLbls>
      <c:cat><c:strLit><c:pt idx="0"><c:v>Product</c:v></c:pt><c:pt idx="1"><c:v>Service</c:v></c:pt></c:strLit></c:cat>
      <c:val><c:numLit><c:pt idx="0"><c:v>75</c:v></c:pt><c:pt idx="1"><c:v>25</c:v></c:pt></c:numLit></c:val>
    </c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>`,
  });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const lines = [];
  const curves = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, rect() {}, bezierCurveTo(...values) { curves.push(values); },
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    lineTo(x, y) { lines.push([x, y]); },
    measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const objects = await document.listObjects({ unitIndex: 0 });
    const explodedSlice = objects.find(({ source }) =>
      source.part === "ppt/charts/chart3.xml" && source.row === 0 && source.column === 1);
    const productLabel = objects.find(({ text }) => text === "Product, 75");
    const serviceLabel = objects.find(({ text }) => text === "Service, 25");
    const firstLabelIndex = objects.findIndex(({ text }) => text === "Product, 75");
    const lastSliceIndex = objects.map(({ type, source }) => (
      type === "shape" && source.part === "ppt/charts/chart3.xml" && source.column !== undefined
    )).lastIndexOf(true);
    const hits = await document.hitTest({ unitIndex: 0, x: 350, y: 300 });
    const sliceHit = hits.find(({ object }) => object.type === "shape");
    frame = await document.render({ unitIndex: 0 });

    assert.equal(sliceHit.object.source.part, "ppt/charts/chart3.xml");
    assert.equal(sliceHit.object.source.row, 0);
    assert.equal(sliceHit.object.source.column, 0);
    assert.ok(explodedSlice.bounds.x < 180 && explodedSlice.bounds.y < 130);
    assert.ok(Math.abs(productLabel.bounds.width - 152) < 0.01);
    assert.ok(Math.abs(serviceLabel.bounds.width - 146.8) < 0.01);
    assert.ok(
      productLabel.bounds.x > serviceLabel.bounds.x + 40,
      JSON.stringify({ product: productLabel.bounds, service: serviceLabel.bounds }),
    );
    assert.ok(firstLabelIndex > lastSliceIndex, "pie labels must paint above every slice");
    assert.ok(fills.includes("rgba(68, 114, 196, 1)"));
    assert.ok(fills.includes("rgba(237, 125, 49, 1)"));
    assert.ok(lines.length >= 2);
    assert.ok(curves.length > 10);
    assert.ok(frame.renderedObjectCount >= 3);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders cached PPTX area, scatter, radar, and doughnut charts", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="28"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="952500"/><a:ext cx="5715000" cy="3810000"/></p:xfrm><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart5.xml"/></Relationships>`,
    "ppt/charts/chart5.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>
      <c:areaChart><c:ser><c:tx><c:v>Area</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser></c:areaChart>
      <c:scatterChart><c:ser><c:tx><c:v>Scatter</c:v></c:tx><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numLit></c:yVal><c:trendline><c:trendlineType val="linear"/><c:dispRSqr val="1"/><c:dispEq val="1"/></c:trendline></c:ser></c:scatterChart>
      <c:radarChart><c:ser><c:tx><c:v>Radar</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser></c:radarChart>
      <c:doughnutChart><c:ser><c:tx><c:v>Doughnut</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numLit></c:val></c:ser><c:holeSize val="55"/></c:doughnutChart>
      <c:valAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>X Axis</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>
      <c:valAx><c:axPos val="l"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Y Axis</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>
    </c:plotArea><c:legend/></c:chart></c:chartSpace>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const objects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "ppt/charts/chart5.xml");
    const series = new Set(objects.map(({ source }) => source.row).filter((row) => row !== undefined));

    assert.deepEqual([...series].sort(), [0, 1, 2, 3]);
    assert.ok(objects.some(({ text }) => text === "Area"), JSON.stringify(objects));
    assert.ok(objects.some(({ text }) => text === "Scatter"));
    assert.ok(objects.some(({ text }) => text === "Radar"));
    assert.ok(objects.some(({ text }) => text === "Doughnut"));
    assert.ok(objects.some(({ text }) => text === "X Axis"));
    assert.ok(objects.some(({ text }) => text === "Y Axis"));
    assert.ok(objects.some(({ text }) => text?.includes("R² =")));
    assert.equal(document.diagnostics().some(({ message }) => message.includes("static placeholder")), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders common PPTX OMML structures as readable mapped math text", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="16"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:p><m:oMath><m:f><m:num><m:r><m:t>x+1</m:t></m:r></m:num><m:den><m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:den></m:f></m:oMath></a:p></p:txBody></p:sp></p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 110, y: 110 });

    assert.equal(hits[0].object.type, "text-box");
    assert.equal(hits[0].object.text, "(x+1)/(y^(2))");
    assert.equal(hits[0].object.source.shapeId, 16);
    assert.equal(document.diagnostics().some(({ message }) => message.includes("static placeholder")), false);
    assert.equal(document.diagnostics().some(({ message, fidelity }) => (
      message.includes("formula") && fidelity === "approximate"
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders extended PPTX OMML delimiters, matrices, functions, accents, and limits", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="29"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="5715000" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:p><m:oMath>
        <m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e><m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr></m:m></m:e></m:d>
        <m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func>
        <m:acc><m:accPr><m:chr m:val="ˆ"/></m:accPr><m:e><m:r><m:t>y</m:t></m:r></m:e></m:acc>
        <m:limLow><m:e><m:r><m:t>lim</m:t></m:r></m:e><m:lim><m:r><m:t>n→∞</m:t></m:r></m:lim></m:limLow>
      </m:oMath></a:p></p:txBody></p:sp></p:spTree></p:cSld>
    </p:sld>`);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const [formula] = await document.listObjects({ unitIndex: 0, types: ["text-box"] });

    assert.equal(formula.text, "[[a, b]; [c, d]]sin(x)yˆlim_(n→∞)");
    assert.equal(document.diagnostics().some(({ message, fidelity }) => (
      message.includes("formula") && fidelity === "approximate"
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("uses a PPTX diagram drawing fallback as real source-mapped shapes", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="70"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
      </p:graphicFrame></p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>`,
    "ppt/diagrams/data1.xml": `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>`,
    "ppt/diagrams/_rels/data1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/></Relationships>`,
    "ppt/diagrams/drawing1.xml": `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:nvGrpSpPr/><dsp:grpSpPr/><dsp:sp><dsp:nvSpPr><dsp:cNvPr id="71"/><dsp:cNvSpPr/><dsp:nvPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="476250"/></a:xfrm><a:prstGeom prst="roundRect"/><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Rendered node</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 110, y: 110 });

    assert.equal(hits[0].object.type, "text-box");
    assert.equal(hits[0].object.text, "Rendered node");
    assert.equal(hits[0].object.source.part, "ppt/diagrams/drawing1.xml");
    assert.equal(hits[0].object.source.shapeId, 71);
    assert.equal(document.diagnostics().some(({ message }) => message.includes("static placeholder")), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("resolves a slide-scoped PPTX diagram drawing and offsets its local coordinates", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="72"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
      </p:graphicFrame></p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data2.xml"/>
      <Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing2.xml"/>
    </Relationships>`,
    "ppt/diagrams/data2.xml": `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt relId="rIdDrawing"/></a:ext></dgm:extLst></dgm:dataModel>`,
    "ppt/diagrams/drawing2.xml": `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:nvGrpSpPr/><dsp:grpSpPr/><dsp:sp><dsp:nvSpPr><dsp:cNvPr id="73"/><dsp:cNvSpPr/><dsp:nvPr/></dsp:nvSpPr><dsp:spPr><a:xfrm rot="2160000"><a:off x="476250" y="476250"/><a:ext cx="1905000" cy="476250"/></a:xfrm><a:prstGeom prst="roundRect"/><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Offset node</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 250, y: 175 });

    assert.equal(hits[0].object.text, "Offset node");
    assert.equal(hits[0].object.source.part, "ppt/diagrams/drawing2.xml");
    assert.deepEqual(hits[0].object.bounds, { x: 150, y: 150, width: 200, height: 50 });
  } finally {
    document?.close();
    engine.close();
  }
});

test("lays out PPTX diagram data when no drawing fallback exists", async () => {
  const bytes = pptxWithSlide(`<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="80"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="2857500"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
      </p:graphicFrame></p:spTree></p:cSld>
    </p:sld>`, {
    "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data2.xml"/></Relationships>`,
    "ppt/diagrams/data2.xml": `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>Root node</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="2"><dgm:t><a:p><a:r><a:t>Child node</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn type="parOf" srcId="1" destId="2"/></dgm:cxnLst></dgm:dataModel>`,
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const rootHits = await document.hitTest({ unitIndex: 0, x: 288, y: 135 });
    const childHits = await document.hitTest({ unitIndex: 0, x: 288, y: 260 });
    const objects = await document.listObjects({ unitIndex: 0 });

    assert.equal(rootHits[0].object.type, "text-box");
    assert.equal(rootHits[0].object.text, "Root node");
    assert.equal(rootHits[0].object.source.part, "ppt/diagrams/data2.xml");
    assert.equal(childHits[0].object.text, "Child node");
    assert.equal(childHits[0].object.source.part, "ppt/diagrams/data2.xml");
    assert.ok(objects.some(({ type, text, source }) => (
      type === "shape" && text === undefined && source.part === "ppt/diagrams/data2.xml"
    )));
    assert.equal(document.diagnostics().some(({ message }) => message.includes("static placeholder")), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("exposes transformed ODP object bounds in document coordinates", async () => {
  const bytes = odpWithContent("", `
    <draw:custom-shape draw:id="transformed" svg:width="2in" svg:height="1in"
      draw:transform="rotate (0.523598775598299) translate (4in 2in)"><text:p>Transformed</text:p></draw:custom-shape>
  `);
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(bytes);
    const object = await document.getObject("object:0");
    assert.ok(object);
    assert.notDeepEqual(object.bounds, { x: 0, y: 0, width: 192, height: 96 });
    const hits = await document.hitTest({
      unitIndex: 0,
      x: object.bounds.x + object.bounds.width / 2,
      y: object.bounds.y + object.bounds.height / 2,
    });
    assert.equal(hits.some((hit) => hit.object.id === object.id), true);
    assert.deepEqual(hits.find((hit) => hit.object.id === object.id)?.object.bounds, object.bounds);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a merged ODP table cell with standard cell border and text styling", async () => {
  const bytes = odpWithContent(`
    <style:style style:name="mergedCell" style:family="table-cell">
      <style:table-cell-properties fo:background-color="#ff0000" fo:border="2px solid #0000ff"/>
      <style:text-properties fo:color="#00ff00" fo:font-family="Aptos" fo:font-size="12pt" fo:font-weight="bold" fo:font-style="italic"/>
    </style:style>
  `, `
    <draw:frame draw:id="merged-table" svg:x="1in" svg:y="2in" svg:width="4in" svg:height="1in">
      <table:table table:name="MergedTable">
        <table:table-row>
          <table:table-cell xml:id="merged-a1" table:style-name="mergedCell" table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>Merged</text:p></table:table-cell>
          <table:covered-table-cell/>
        </table:table-row>
        <table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row>
      </table:table>
    </draw:frame>
  `);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const strokes = [];
  const widths = [];
  const fonts = [];
  const context = {
    globalAlpha: 1,
    letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); },
    set strokeStyle(value) { strokes.push(value); },
    set lineWidth(value) { widths.push(value); },
    set font(value) { fonts.push(value); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 450, y: 275 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.type, "cell");
    assert.equal(hits[0].object.text, "Merged");
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, 0);
    assert.deepEqual(hits[0].object.bounds, { x: 96, y: 192, width: 384, height: 96 });
    assert.ok(fills.includes("rgba(255, 0, 0, 1)"));
    assert.ok(fills.includes("rgba(0, 255, 0, 1)"));
    assert.ok(strokes.includes("rgba(0, 0, 255, 1)"));
    assert.ok(widths.includes(2));
    assert.ok(fonts.some((font) => font.includes("italic 700") && font.includes("Aptos")));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("inherits ODP default and forward-declared parent style colors and fonts", async () => {
  const bytes = odpWithContent(`
    <style:default-style style:family="graphic">
      <style:graphic-properties draw:fill="solid" draw:fill-color="#654321"/>
      <style:text-properties fo:font-family="'Default Font'" fo:color="#aa0000"/>
    </style:default-style>
    <style:style style:name="child" style:family="graphic" style:parent-style-name="parent">
      <style:text-properties fo:font-family="'Child Font'"/>
    </style:style>
    <style:style style:name="defaultOnly" style:family="graphic"/>
    <style:style style:name="parent" style:family="graphic">
      <style:graphic-properties draw:fill="solid" draw:fill-color="#123456"/>
      <style:text-properties fo:font-family="'Parent Font'" fo:color="#00aa00"/>
    </style:style>
  `, `
    <draw:rect draw:id="child-style" draw:style-name="child" svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in"><text:p>Child</text:p></draw:rect>
    <draw:rect draw:id="default-style" draw:style-name="defaultOnly" svg:x="4in" svg:y="1in" svg:width="2in" svg:height="1in"><text:p>Default</text:p></draw:rect>
  `);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const fonts = [];
  const context = {
    globalAlpha: 1,
    letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 20 }; },
    set font(value) { fonts.push(value); },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const childHits = await document.hitTest({ unitIndex: 0, x: 120, y: 120 });
    const defaultHits = await document.hitTest({ unitIndex: 0, x: 400, y: 120 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(childHits[0].object.source.elementId, "child-style");
    assert.equal(defaultHits[0].object.source.elementId, "default-style");
    assert.ok(fills.includes("rgba(18, 52, 86, 1)"));
    assert.ok(fills.includes("rgba(0, 170, 0, 1)"));
    assert.ok(fills.includes("rgba(101, 67, 33, 1)"));
    assert.ok(fills.includes("rgba(170, 0, 0, 1)"));
    assert.deepEqual(
      childHits[0].object.fontRuns.map(({ authoredFamily, renderedFamily, source }) => ({
        authoredFamily,
        renderedFamily,
        source,
      })),
      [{ authoredFamily: "Child Font", renderedFamily: "Calibri", source: "fallback" }],
    );
    assert.deepEqual(
      defaultHits[0].object.fontRuns.map(({ authoredFamily, renderedFamily, source }) => ({
        authoredFamily,
        renderedFamily,
        source,
      })),
      [{ authoredFamily: "Default Font", renderedFamily: "Calibri", source: "fallback" }],
    );
    assert.ok(fonts.some((font) => font.includes("Calibri")));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("vertically centers ODP text from its graphic style", async () => {
  const bytes = odpWithContent(`
    <style:style style:name="middleText" style:family="graphic">
      <style:graphic-properties draw:textarea-vertical-align="middle"/>
      <style:text-properties fo:font-size="12pt"/>
    </style:style>
  `, `
    <draw:frame draw:id="middle-text" draw:style-name="middleText" svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in">
      <draw:text-box><text:p>Centered</text:p></draw:text-box>
    </draw:frame>
  `);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const textPositions = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText(text, x, y) { textPositions.push([text, x, y]); },
    measureText() { return { width: 40 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(textPositions[0][0], "Centered");
    assert.ok(textPositions[0][2] > 125 && textPositions[0][2] < 135);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders ODP draw shadow color opacity and offsets", async () => {
  const bytes = odpWithContent(`
    <style:style style:name="shadowShape" style:family="graphic">
      <style:graphic-properties draw:fill="solid" draw:fill-color="#ffffff" draw:shadow="visible" draw:shadow-color="#112233" draw:shadow-opacity="50%" draw:shadow-offset-x="0.1in" draw:shadow-offset-y="0.2in"/>
    </style:style>
  `, `<draw:rect draw:id="shadow-rect" draw:style-name="shadowShape" svg:x="1in" svg:y="1in" svg:width="1in" svg:height="0.5in"/>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const shadowColors = [];
  const shadowBlurs = [];
  const shadowOffsetsX = [];
  const shadowOffsetsY = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(_value) {}, set lineWidth(_value) {},
    set shadowColor(value) { shadowColors.push(value); },
    set shadowBlur(value) { shadowBlurs.push(value); },
    set shadowOffsetX(value) { shadowOffsetsX.push(value); },
    set shadowOffsetY(value) { shadowOffsetsY.push(value); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    frame = await document.render({ unitIndex: 0 });

    assert.ok(shadowColors.includes("rgba(17, 34, 51, 0.5019607843137255)"));
    assert.ok(shadowBlurs.includes(0));
    assert.ok(shadowOffsetsX.some((value) => Math.abs(value - 9.6) < 0.001));
    assert.ok(shadowOffsetsY.some((value) => Math.abs(value - 19.2) < 0.001));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders an inline cached ODP bar chart as source-mapped display objects", async () => {
  const bytes = odpWithContent("", `
    <draw:frame draw:id="odp-bar" svg:x="1in" svg:y="1in" svg:width="4in" svg:height="3in">
      <chart:chart chart:class="chart:bar">
        <chart:plot-area>
          <chart:categories table:cell-range-address="local.A2:A3"/>
          <chart:series chart:values-cell-range-address="local.B2:B3" chart:label-cell-address="local.B1"/>
        </chart:plot-area>
        <table:table table:name="local">
          <table:table-row><table:table-cell/><table:table-cell office:value-type="string"><text:p>Revenue</text:p></table:table-cell></table:table-row>
          <table:table-row><table:table-cell office:value-type="string"><text:p>Q1</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="20"/></table:table-row>
          <table:table-row><table:table-cell office:value-type="string"><text:p>Q2</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="10"/></table:table-row>
        </table:table>
      </chart:chart>
    </draw:frame>
  `);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 200, y: 200 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.type, "shape");
    assert.equal(hits[0].object.source.part, "content.xml");
    assert.equal(hits[0].object.source.elementId, "odp-bar");
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, 0);
    assert.ok(fills.includes("rgba(68, 114, 196, 1)"));
    assert.ok(frame.renderedObjectCount >= 4);
    assert.equal(document.diagnostics().some(({ message }) => message.includes("chart content uses a source-mapped static placeholder")), false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders an embedded ODP chart object from its package part", async () => {
  const embeddedChart = `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
      <office:body><office:chart>
        <chart:chart chart:class="chart:bar"><chart:plot-area>
          <chart:categories table:cell-range-address="local.A2:A3"/>
          <chart:series chart:values-cell-range-address="local.B2:B3" chart:label-cell-address="local.B1"/>
        </chart:plot-area><table:table table:name="local">
          <table:table-row><table:table-cell/><table:table-cell><text:p>Revenue</text:p></table:table-cell></table:table-row>
          <table:table-row><table:table-cell><text:p>Q1</text:p></table:table-cell><table:table-cell office:value="20"/></table:table-row>
          <table:table-row><table:table-cell><text:p>Q2</text:p></table:table-cell><table:table-cell office:value="10"/></table:table-row>
        </table:table></chart:chart>
      </office:chart></office:body>
    </office:document-content>`;
  const bytes = odpWithContent("", `
    <draw:frame draw:id="embedded-chart" svg:x="1in" svg:y="1in" svg:width="4in" svg:height="3in">
      <draw:object xlink:href="./Object 1"/>
    </draw:frame>
  `, "", { "Object 1/content.xml": embeddedChart });
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 200, y: 200 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.type, "shape");
    assert.equal(hits[0].object.source.part, "Object 1/content.xml");
    assert.equal(hits[0].object.source.elementId, "embedded-chart");
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, 0);
    assert.ok(fills.includes("rgba(68, 114, 196, 1)"));
    assert.ok(frame.renderedObjectCount >= 4);
    assert.equal(document.diagnostics().some(({ message }) => message.includes("static placeholder")), false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders common embedded ODP MathML as readable source-mapped text", async () => {
  const math = `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML">
      <office:body><office:formula><math:math><math:msqrt><math:mfrac><math:mi>a</math:mi><math:mi>b</math:mi></math:mfrac></math:msqrt></math:math></office:formula></office:body>
    </office:document-content>`;
  const bytes = odpWithContent("", `
    <draw:frame draw:id="embedded-math" svg:x="1in" svg:y="1in" svg:width="4in" svg:height="1in">
      <draw:object xlink:href="./Object 2"/>
    </draw:frame>
  `, "", { "Object 2/content.xml": math });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 200, y: 120 });

    assert.equal(hits[0].object.type, "text-box");
    assert.equal(hits[0].object.text, "√((a)/(b))");
    assert.equal(hits[0].object.source.part, "Object 2/content.xml");
    assert.equal(hits[0].object.source.elementId, "embedded-math");
    assert.equal(document.diagnostics().some(({ message }) => message.includes("static placeholder")), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders ODP path and polygon coordinates from their SVG view boxes", async () => {
  const bytes = odpWithContent(`
    <style:style style:name="pathStyle" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#336699" draw:stroke="solid" svg:stroke-color="#112233" svg:stroke-width="1px"/></style:style>
  `, `
    <draw:path draw:id="triangle-path" draw:style-name="pathStyle" svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in" svg:viewBox="0 0 200 100" svg:d="M 0 100 L 100 0 L 200 100 Z"/>
    <draw:polygon draw:id="triangle-polygon" draw:style-name="pathStyle" svg:x="4in" svg:y="1in" svg:width="1in" svg:height="1in" svg:viewBox="0 0 100 100" draw:points="0,100 50,0 100,100"/>
  `);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const moves = [];
  const lines = [];
  const fills = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, rect() {}, bezierCurveTo() {}, quadraticCurveTo() {},
    closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    moveTo(x, y) { moves.push([x, y]); }, lineTo(x, y) { lines.push([x, y]); },
    measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const pathHits = await document.hitTest({ unitIndex: 0, x: 192, y: 140 });
    const polygonHits = await document.hitTest({ unitIndex: 0, x: 432, y: 140 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(pathHits[0].object.source.elementId, "triangle-path");
    assert.equal(polygonHits[0].object.source.elementId, "triangle-polygon");
    assert.ok(moves.some(([x, y]) => Math.abs(x - 96) < 0.001 && Math.abs(y - 192) < 0.001));
    assert.ok(lines.some(([x, y]) => Math.abs(x - 192) < 0.001 && Math.abs(y - 96) < 0.001));
    assert.ok(moves.some(([x, y]) => Math.abs(x - 384) < 0.001 && Math.abs(y - 192) < 0.001));
    assert.ok(lines.some(([x, y]) => Math.abs(x - 432) < 0.001 && Math.abs(y - 96) < 0.001));
    assert.ok(fills.includes("rgba(51, 102, 153, 1)"));
    assert.equal(document.diagnostics().some(({ message }) => message.includes("rectangle fallback")), false);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders an inline cached ODP line chart as source-mapped series segments", async () => {
  const bytes = odpWithContent("", `
    <draw:frame draw:id="odp-line" svg:x="1in" svg:y="1in" svg:width="4in" svg:height="3in">
      <chart:chart chart:class="chart:line"><chart:plot-area>
        <chart:categories table:cell-range-address="local.A2:A3"/>
        <chart:series chart:values-cell-range-address="local.B2:B3" chart:label-cell-address="local.B1"/>
      </chart:plot-area><table:table table:name="local">
        <table:table-row><table:table-cell/><table:table-cell><text:p>Revenue</text:p></table:table-cell></table:table-row>
        <table:table-row><table:table-cell><text:p>Q1</text:p></table:table-cell><table:table-cell office:value="20"/></table:table-row>
        <table:table-row><table:table-cell><text:p>Q2</text:p></table:table-cell><table:table-cell office:value="10"/></table:table-row>
      </table:table></chart:chart>
    </draw:frame>
  `);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const strokedPaths = [];
  let currentPath = [];
  let currentStrokeStyle = "";
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() { currentPath = []; }, ellipse() {},
    moveTo(x, y) { currentPath.push(["move", x, y]); }, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {},
    stroke() { strokedPaths.push({ color: currentStrokeStyle, path: currentPath }); },
    clip() {}, fillText() {},
    lineTo(x, y) { currentPath.push(["line", x, y]); }, measureText() { return { width: 10 }; },
    set fillStyle(_value) {}, set strokeStyle(value) { currentStrokeStyle = value; }, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 300, y: 180 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.type, "shape");
    assert.equal(hits[0].object.source.part, "content.xml");
    assert.equal(hits[0].object.source.elementId, "odp-line");
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, 0);
    const seriesPaths = strokedPaths.filter(({ color }) => color === "rgba(68, 114, 196, 1)");
    assert.equal(seriesPaths.length, 1);
    assert.equal(seriesPaths[0].path.length, 2);
    const [start, end] = seriesPaths[0].path;
    assert.equal(start[0], "move");
    assert.ok(Math.abs(start[1] - 126.72) < 0.01);
    assert.ok(Math.abs(start[2] - 133.702) < 0.01);
    assert.equal(end[0], "line");
    assert.ok(Math.abs(end[1] - 456.96) < 0.01);
    assert.ok(Math.abs(end[2] - 235.811) < 0.01);
    assert.ok(frame.renderedObjectCount >= 4);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders an inline cached ODP circle chart as source-mapped slice paths", async () => {
  const bytes = odpWithContent("", `
    <draw:frame draw:id="odp-pie" svg:x="1in" svg:y="1in" svg:width="4in" svg:height="3in">
      <chart:chart chart:class="chart:circle"><chart:plot-area>
        <chart:categories table:cell-range-address="local.A2:A3"/>
        <chart:series chart:values-cell-range-address="local.B2:B3"/>
      </chart:plot-area><table:table table:name="local">
        <table:table-row><table:table-cell/><table:table-cell/></table:table-row>
        <table:table-row><table:table-cell><text:p>Product</text:p></table:table-cell><table:table-cell office:value="75"/></table:table-row>
        <table:table-row><table:table-cell><text:p>Service</text:p></table:table-cell><table:table-cell office:value="25"/></table:table-row>
      </table:table></chart:chart>
    </draw:frame>
  `);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const lines = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    lineTo(x, y) { lines.push([x, y]); }, measureText() { return { width: 10 }; },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const hits = await document.hitTest({ unitIndex: 0, x: 350, y: 280 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(hits[0].object.type, "shape");
    assert.equal(hits[0].object.source.part, "content.xml");
    assert.equal(hits[0].object.source.elementId, "odp-pie");
    assert.equal(hits[0].object.source.row, 0);
    assert.equal(hits[0].object.source.column, 0);
    assert.ok(fills.includes("rgba(68, 114, 196, 1)"));
    assert.ok(fills.includes("rgba(237, 125, 49, 1)"));
    assert.ok(lines.length > 10);
    assert.ok(frame.renderedObjectCount >= 3);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders styled ODP backgrounds, grouped shapes, tables, and chart placeholders", async () => {
  const bytes = odpWithContent(`
    <draw:gradient draw:name="grad1" draw:start-color="#102030" draw:end-color="#405060" draw:angle="0"/>
    <style:style style:name="page1" style:family="drawing-page"><style:drawing-page-properties draw:fill="gradient" draw:fill-gradient-name="grad1"/></style:style>
    <style:style style:name="shape1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000" draw:stroke="solid" svg:stroke-color="#00ff00" svg:stroke-width="2px"/></style:style>
    <style:style style:name="masterShape" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#7030a0"/></style:style>
    <style:style style:name="cell1" style:family="table-cell"><style:graphic-properties draw:fill="solid" draw:fill-color="#fff2cc"/><style:text-properties fo:color="#0000ff" fo:font-size="14pt" fo:font-weight="bold"/></style:style>
    <style:style style:name="textRed" style:family="text"><style:text-properties fo:color="#ff0000" fo:font-size="12pt" fo:font-weight="bold"/></style:style>
    <style:style style:name="textBlue" style:family="text"><style:text-properties fo:color="#0000ff" fo:font-size="20pt" fo:font-style="italic"/></style:style>
  `, `
    <draw:g draw:id="group1" draw:transform="translate(1in 0in)">
      <draw:rect draw:id="rect1" draw:style-name="shape1" svg:x="1in" svg:y="1in" svg:width="1in" svg:height="1in"/>
    </draw:g>
    <draw:frame draw:id="rich-text" svg:x="3.5in" svg:y="1in" svg:width="2in" svg:height="1in"><draw:text-box><text:p><text:span text:style-name="textRed">Red</text:span><text:span text:style-name="textBlue">Blue</text:span></text:p></draw:text-box></draw:frame>
    <draw:frame draw:id="table-frame" svg:x="1in" svg:y="3in" svg:width="4in" svg:height="1in"><table:table table:name="Table1"><table:table-row><table:table-cell xml:id="a1" table:style-name="cell1"><text:p>A1</text:p></table:table-cell><table:table-cell xml:id="b1"><text:p>B1</text:p></table:table-cell></table:table-row></table:table></draw:frame>
    <draw:frame draw:id="chart1" svg:x="6in" svg:y="1in" svg:width="2in" svg:height="1in"><chart:chart/></draw:frame>
  `, `<draw:rect draw:id="master-mark" draw:style-name="masterShape" svg:x="8in" svg:y="6in" svg:width="1in" svg:height="0.5in"/>`);
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const gradientStops = [];
  const fills = [];
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "",
    save() {}, restore() {}, scale() {}, translate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, bezierCurveTo() {},
    quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {}, fillText() {},
    measureText() { return { width: 10 }; },
    createLinearGradient() { return { addColorStop(offset, color) { gradientStops.push([offset, color]); } }; },
    set fillStyle(value) { fills.push(value); }, set strokeStyle(_value) {}, set lineWidth(_value) {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(bytes);
    const shapeHits = await document.hitTest({ unitIndex: 0, x: 210, y: 110 });
    const tableHits = await document.hitTest({ unitIndex: 0, x: 120, y: 310 });
    const chartHits = await document.hitTest({ unitIndex: 0, x: 590, y: 110 });
    const masterHits = await document.hitTest({ unitIndex: 0, x: 780, y: 590 });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(shapeHits[0].object.source.elementId, "rect1");
    assert.equal(shapeHits[0].ancestors[0].source.elementId, "group1");
    assert.equal(tableHits[0].object.type, "cell");
    assert.equal(tableHits[0].object.text, "A1");
    assert.equal(tableHits[0].object.source.row, 0);
    assert.equal(tableHits[0].object.source.column, 0);
    assert.equal(chartHits[0].object.type, "unknown");
    assert.equal(chartHits[0].object.text, "Chart");
    assert.equal(masterHits[0].object.source.part, "styles.xml");
    assert.equal(masterHits[0].object.source.elementId, "master-mark");
    assert.ok(fills.includes("rgba(255, 0, 0, 1)"));
    assert.ok(fills.includes("rgba(0, 0, 255, 1)"));
    assert.deepEqual(gradientStops, [[0, "rgba(16, 32, 48, 1)"], [1, "rgba(64, 80, 96, 1)"]]);
    assert.ok(document.diagnostics().some(({ message, fidelity }) => (
      message.includes("static placeholder") && fidelity === "approximate"
    )));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("real orgChart1 preserves assistant branches, deep levels and authored blue fills", async () => {
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("fixtures/smartart-orgchart.pptx", import.meta.url)));
  try {
    document.loadUnit(0);
    const nodes = document.scene.objects.filter(object => object.text && object.source.part === "ppt/diagrams/data1.xml");
    assert.ok(document.scene.objects.some(o => o.visual.kind === "painted-shape"
      && Math.abs(o.bounds.x - 218) < 3 && Math.abs(o.bounds.y - 340) < 3),
      "assistant branches retain the central connector site");
    // PowerPoint's final static build, exported from the unmodified SDK corpus file.
    const expected = [
      ["Top", 326, 168], ["Mid-Left", 147, 269], ["Mid-Right", 583, 269],
      ["Bottom-Left", 240, 370], ["Bottom-Left", 240, 471],
      ["Bottom-Right", 497, 370], ["Bottom-Right", 669, 370],
      ["Temp", 497, 471], ["Assistant", 411, 572],
    ];
    assert.equal(nodes.length, expected.length);
    for (const [index, [text, x, y]] of expected.entries()) {
      const node = nodes[index];
      assert.equal(node.text, text);
      let visual = node.visual;
      while (visual.visual) visual = visual.visual;
      assert.deepEqual(visual.fill, { kind: "solid", color: 0x4f81bdff }, text);
      assert.equal(visual.geometry, "rectangle", text);
      for (const [key, value] of Object.entries({x, y, width: 142, height: 71})) {
        assert.ok(Math.abs(node.bounds[key] - value) < 3, `${text} ${key}: ${node.bounds[key]} vs ${value}`);
      }
    }
  } finally { document.close(); core.close(); }
});

test("real Chart_2D fills the chart frame and keeps category labels horizontal", async () => {
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("fixtures/chart-2d.pptx", import.meta.url)));
  try {
    document.loadUnit(0);
    const objects = document.scene.objects.filter(o => o.source.part === "ppt/charts/chart1.xml");
    const categories = objects.filter(o => o.text?.startsWith("Category"));
    assert.equal(categories.length, 4);
    for (const label of categories) {
      assert.equal(label.visual.layout.rotationDegrees, 0, label.text);
      assert.equal(label.visual.visual.runs[0].fontSize, 24);
    }
    const gridlines = objects.filter(o => o.visual.kind === "painted-shape" && o.visual.geometry === "line" && o.bounds.width > 100 && o.bounds.height < 1);
    // PowerPoint reference: plot approximately (91, 189, 688, 401) at 96 DPI.
    assert.ok(gridlines.length >= 7);
    for (const line of gridlines) {
      assert.ok(Math.abs(line.bounds.x - 91) < 6, `plot left: ${line.bounds.x}`);
      assert.ok(Math.abs(line.bounds.width - 688) < 12, `plot width: ${line.bounds.width}`);
    }
    const ys = gridlines.map(o => o.bounds.y);
    assert.ok(Math.abs(Math.min(...ys) - 189) < 5);
    assert.ok(Math.abs(Math.max(...ys) - 590) < 5);
    const legends = objects.filter(o => o.text?.startsWith("Series"));
    assert.equal(legends.length, 3);
    for (const label of legends) assert.equal(label.visual.visual.runs[0].fontSize, 24);
  } finally { document.close(); core.close(); }
});

test("real 3D orgChart1 retains bevel, material, lighting, theme gradient and shadow", async () => {
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const document = core.open(await readFile(new URL("fixtures/smartart-orgchart-3d.pptx", import.meta.url)));
  try {
    document.loadUnit(0);
    const nodes = document.scene.objects.filter(o => o.text && o.source.part === "ppt/diagrams/data1.xml");
    assert.equal(nodes.length, 9);
    assert.ok(document.scene.objects.some(o => o.visual.kind === "painted-shape"
      && Math.abs(o.bounds.x - 233) < 3 && Math.abs(o.bounds.y - 340) < 3),
      "ordinary hanging branches use the left connector site");
    const expected = [[365,168], [219,269], [254,370], [254,471], [511,269],
      [426,370], [426,471], [462,572], [598,370]];
    for (const [index, node] of nodes.entries()) {
      assert.ok(Math.abs(node.bounds.x - expected[index][0]) < 3 && Math.abs(node.bounds.y - expected[index][1]) < 3,
        `${node.text}: incorrect native layout ${JSON.stringify(node.bounds)}`);
      const layers = [];
      for (let v = node.visual; v; v = v.visual) layers.push(v);
      const effect = layers.find(v => v.threeD);
      assert.ok(effect, `${node.text}: missing SmartArt quickStyle 3D`);
      assert.equal(effect.threeD.material, "plastic");
      assert.equal(effect.threeD.bevelTop.preset, "relaxedInset");
      assert.ok(Math.abs(effect.threeD.bevelTop.width - 127000 / 9525) < .001);
      assert.ok(Math.abs(effect.threeD.bevelTop.height - 25400 / 9525) < .001);
      assert.equal(effect.threeD.lightRevolution, 125);
      assert.ok(layers.some(v => v.outerShadow?.color === 0x00000059));
      assert.equal(layers.at(-1).fill.kind, "linear-gradient");
      assert.equal(layers.at(-1).stroke.kind, "none");
    }
  } finally { document.close(); core.close(); }
});

test("real SmartArt keeps its nodes when an optional quick style is corrupt or unrelated", async () => {
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const source = await readFile(new URL("fixtures/smartart-orgchart-3d.pptx", import.meta.url));
  const part = "ppt/diagrams/quickStyle1.xml";
  try {
    for (const corrupt of [false, true]) {
      const parts = Object.fromEntries(readZipEntries(source).map(({ name, data }) => [name, data]));
      parts[part] = corrupt ? "<broken" : new TextDecoder().decode(parts[part]).replace('quickstyle/3d2', 'quickstyle/unrelated');
      const doc = core.open(createZip(parts));
      try {
        doc.loadUnit(0);
        const nodes = doc.scene.objects.filter(o => o.text && o.source.part === "ppt/diagrams/data1.xml");
        assert.equal(nodes.length, 9);
        assert.ok(nodes.every(o => o.visual.kind === "text-layout"));
        assert.equal(doc.scene.diagnostics.some(d => d.part === part && d.severity === "warning"), corrupt);
      } finally { doc.close(); }
    }
  } finally { core.close(); }
});

test("real text-effects PPTX resolves gradient coordinates against inherited placeholder bounds", async () => {
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("fixtures/Text_withEffects_100chars.pptx", import.meta.url)));
  try {
    doc.loadUnit(0);
    const body = doc.scene.objects.find(o => o.text?.includes("29% OF SLIDES"));
    assert.ok(body);
    let visual = body.visual;
    while (visual.visual) visual = visual.visual;
    const paint = visual.runs.find(run => run.text.startsWith("29%")).paint;
    assert.equal(paint.kind, "linear-gradient");
    assert.ok(Math.abs(paint.end.y - paint.start.y - body.bounds.height) < .01,
      "gradient must use the inherited slide-layout height, not an empty local extent");
  } finally { doc.close(); core.close(); }
});

test("real text-effects PPTX keeps title outline/glow and first-run bullet styling", async () => {
  const { Core, DEFAULT_LIMITS } = await import("../dist/core.js");
  const core = await Core.create(await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)), DEFAULT_LIMITS);
  const doc = core.open(await readFile(new URL("fixtures/Text_withEffects_100chars.pptx", import.meta.url)));
  try {
    doc.loadUnit(0);
    const effectsOf = object => { let v = object.visual; while (v && v.kind !== "text-effects") v = v.visual; return v; };
    const title = effectsOf(doc.scene.objects.find(o => o.text === "Text: 100 characters"));
    assert.ok(title?.effects[0].stroke, "white title requires its authored orange outline");
    assert.ok(title.effects[0].glow?.radius > 5, "title glow is a glyph effect");
    assert.ok(title.effects[0].strokeWidth > 1);
    const body = effectsOf(doc.scene.objects.find(o => o.text?.includes("29% OF SLIDES")));
    assert.deepEqual(body.visual.runs[0].paint, body.visual.runs[1].paint, "bullet follows first text fill");
    assert.deepEqual(body.effects[0].reflection, body.effects[1].reflection, "bullet follows first text effects");
  } finally { doc.close(); core.close(); }
});
