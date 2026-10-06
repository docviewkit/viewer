import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";
import { extendedFormatPack } from "../dist/extended-formats.js";
import { createZip } from "./zip-fixture.mjs";

const onePixelPng = Uint8Array.of(
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2,
  0, 0, 0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1,
  5, 1, 1, 39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
);

const docx = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
      <Override PartName="/word/headerFirst.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
      <Override PartName="/word/headerEven.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
      <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
      <Override PartName="/word/footerFirst.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
      <Override PartName="/word/footerEven.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
      <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
      <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
      <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
      <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
      <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
  "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
      <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
      <Relationship Id="rIdHeaderFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerFirst.xml"/>
      <Relationship Id="rIdHeaderEven" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerEven.xml"/>
      <Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
      <Relationship Id="rIdFooterFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerFirst.xml"/>
      <Relationship Id="rIdFooterEven" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerEven.xml"/>
      <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
      <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
      <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
      <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
      <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
      <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>`,
  "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body>
        <w:p w14:paraId="11111111"><w:pPr><w:jc w:val="center"/></w:pPr>
          <w:r><w:rPr><w:rFonts w:ascii="Aptos"/><w:sz w:val="28"/><w:color w:val="c00000"/><w:b/><w:u w:val="single"/><w:highlight w:val="yellow"/></w:rPr><w:t>Rich</w:t></w:r>
          <w:r><w:rPr><w:rFonts w:ascii="Courier New"/><w:i/></w:rPr><w:t> document</w:t></w:r>
          <w:commentRangeStart w:id="4"/><w:r><w:footnoteReference w:id="2"/><w:endnoteReference w:id="3"/><w:commentReference w:id="4"/></w:r><w:commentRangeEnd w:id="4"/>
          <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="7" name="Inline image"/><a:graphic><a:graphicData><a:blip r:embed="rIdImage"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
        </w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>List item</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> TOC \\h </w:instrText><w:fldChar w:fldCharType="separate"/>
          <w:fldChar w:fldCharType="begin"/><w:instrText> HYPERLINK \\l "field-target" </w:instrText><w:fldChar w:fldCharType="separate"/>
          <w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Field entry</w:t><w:fldChar w:fldCharType="end"/><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> TOC \\h </w:instrText><w:fldChar w:fldCharType="separate"/></w:r>
          <w:hyperlink w:anchor="element-target"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Element entry</w:t></w:r></w:hyperlink>
          <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:bookmarkStart w:id="1" w:name="field-target"/><w:r><w:t>Page two</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
        <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:bookmarkStart w:id="2" w:name="element-target"/><w:r><w:t>Page three</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>
        <w:sectPr><w:titlePg/><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdHeaderFirst"/><w:headerReference w:type="even" r:id="rIdHeaderEven"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFooterFirst"/><w:footerReference w:type="even" r:id="rIdFooterEven"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body>
    </w:document>`,
  "word/header1.xml": `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Header</w:t></w:r></w:p></w:hdr>`,
  "word/headerFirst.xml": `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:t>First Header</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="457200" cy="457200"/><wp:docPr id="17" name="Header image"/><a:graphic><a:graphicData><a:blip r:embed="rIdHeaderImage"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:hdr>`,
  "word/_rels/headerFirst.xml.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeaderImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`,
  "word/headerEven.xml": `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Even Header</w:t></w:r></w:p></w:hdr>`,
  "word/footer1.xml": `<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>`,
  "word/footerFirst.xml": `<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>First Footer</w:t></w:r></w:p></w:ftr>`,
  "word/footerEven.xml": `<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Even Footer</w:t></w:r></w:p></w:ftr>`,
  "word/numbering.xml": `<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
  "word/settings.xml": `<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:evenAndOddHeaders/></w:settings>`,
  "word/footnotes.xml": `<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:r><w:t>Footnote body</w:t></w:r></w:p></w:footnote></w:footnotes>`,
  "word/endnotes.xml": `<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="3"><w:p><w:r><w:t>Endnote body</w:t></w:r></w:p></w:endnote></w:endnotes>`,
  "word/comments.xml": `<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="4"><w:p><w:r><w:t>Comment body</w:t></w:r></w:p></w:comment></w:comments>`,
  "word/styles.xml": `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="character" w:styleId="Hyperlink"><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`,
  "word/media/image1.png": onePixelPng,
});

const odt = createZip({
  mimetype: "application/vnd.oasis.opendocument.text",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
      <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
      <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"/>
    </manifest:manifest>`,
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
      <office:styles>
        <style:style style:name="Body" style:family="paragraph"><style:paragraph-properties fo:text-align="center" fo:line-height="150%"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="12pt" fo:color="#123456" fo:font-weight="bold"/></style:style>
        <style:style style:name="Em" style:family="text"><style:text-properties style:font-name="Liberation Serif" fo:font-size="14pt" fo:color="#c00000" style:text-underline-style="solid" fo:background-color="#ffff00"/></style:style>
        <style:style style:name="Graphic" style:family="graphic"><style:graphic-properties style:wrap="run-through" draw:fill="solid" draw:fill-color="#d9eaf7" draw:stroke="solid" svg:stroke-color="#336699" svg:stroke-width="1pt"/></style:style>
        <text:list-style style:name="Numbered"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/><text:list-level-style-bullet text:level="2" text:bullet-char="•"/></text:list-style>
      </office:styles>
      <office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin="1in"/></style:page-layout></office:automatic-styles>
      <office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>ODT Header</text:p></style:header><style:footer><text:p>ODT Footer</text:p></style:footer></style:master-page></office:master-styles>
    </office:document-styles>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">
      <office:body><office:text>
        <text:p text:style-name="Body">ODT <text:span text:style-name="Em">rich text</text:span></text:p>
        <text:list text:style-name="Numbered"><text:list-item><text:p>First list item</text:p><text:list><text:list-item><text:p>Nested item</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>Second list item</text:p></text:list-item></text:list>
        <table:table xml:id="span-table"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell xml:id="span-cell" table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>Spanning cell</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row></table:table>
        <draw:frame xml:id="image-frame" text:anchor-type="as-char" svg:width="1in" svg:height="0.5in"><draw:image xlink:href="Pictures/image1.png"/></draw:frame>
        <draw:frame xml:id="text-frame" draw:style-name="Graphic" draw:transform="rotate(0.15)" svg:x="2in" svg:y="2in" svg:width="2in" svg:height="1in"><draw:text-box><text:p>Frame text</text:p></draw:text-box></draw:frame>
        <draw:rect xml:id="rect-shape" draw:style-name="Graphic" draw:transform="translate(4pt 0pt)" svg:x="1in" svg:y="4in" svg:width="2in" svg:height="1in"/>
        <draw:ellipse xml:id="ellipse-shape" draw:style-name="Graphic" svg:x="4in" svg:y="4in" svg:width="1in" svg:height="1in"/>
      </office:text></office:body>
    </office:document-content>`,
  "Pictures/image1.png": onePixelPng,
}, { compress: false });

function installCanvas() {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const bitmapDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const calls = { text: [], images: 0 };
  const values = { font: "", fillStyle: "", strokeStyle: "", letterSpacing: "0px", globalAlpha: 1 };
  const context = new Proxy(values, {
    get(target, property) {
      if (property === "measureText") {
        return (text) => ({ width: [...text].length * 8 });
      }
      if (property === "fillText") {
        return (text) => calls.text.push({ text, font: target.font, fillStyle: target.fillStyle });
      }
      if (property === "drawImage") {
        return () => { calls.images += 1; };
      }
      if (property === "createLinearGradient" || property === "createRadialGradient") {
        return () => ({ addColorStop() {} });
      }
      if (property in target) return target[property];
      return () => {};
    },
    set(target, property, value) {
      target[property] = value;
      return true;
    },
  });
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 1, height: 1, close() {} }),
  });
  return {
    calls,
    restore() {
      if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
      else delete globalThis.OffscreenCanvas;
      if (bitmapDescriptor) Object.defineProperty(globalThis, "createImageBitmap", bitmapDescriptor);
      else delete globalThis.createImageBitmap;
    },
  };
}

async function createEngine() {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  return createOfficeEngine({
    wasm,
    execution: "inline",
    fontPolicy: "local-first",
    formatPack: async () => ({
      async load(candidate) {
        return readFile(await extendedFormatPack.load(candidate));
      },
    }),
  });
}

test("supplied DOC renders source-mapped headers on both pages through the public engine", async () => {
  const canvas = installCanvas();
  const engine = await createEngine();
  let document;
  let frame;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/report-drop-cap.doc", import.meta.url)));
    for (const unitIndex of [0, 1]) {
      const header = await document.getObject(`doc:header:${unitIndex}:0`);
      assert.equal(header?.text, "Simple Home Styling");
      assert.equal(header?.source.kind, "paragraph");
      assert.deepEqual(header?.source.textRange, [1461, 1481]);
      assert.equal(header?.bounds.x, 96);
      assert.equal(header?.bounds.y, 40);
      const hits = await document.hitTest({ unitIndex, x: 100, y: 45 });
      assert.ok(hits.some(({ object }) => object.id === header.id));
      canvas.calls.text.length = 0;
      frame = await document.render({ unitIndex, viewport: { x: 90, y: 30, width: 610, height: 45 } });
      assert.match(canvas.calls.text.map(({ text }) => text).join(""), /Simple Home Styling/);
      frame.bitmap.close();
      frame = undefined;
    }
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    canvas.restore();
  }
});

test("DOCX rich text, numbering, stories, and inline images cross the public engine seam", async () => {
  const canvas = installCanvas();
  const engine = await createEngine();
  let document;
  let frame;
  try {
    document = await engine.open(docx);
    const image = await document.getObject("docx:drawing:0:7");
    const header = await document.getObject("docx:header:0:0");
    const footer = await document.getObject("docx:footer:0:0");
    const evenHeader = await document.getObject("docx:header:1:0");
    const defaultHeader = await document.getObject("docx:header:2:0");
    const headerImage = await document.getObject("docx:header-drawing:0:0:17");
    const numbering = await document.getObject("docx:numbering:1");
    const fieldTocResult = await document.getObject("docx:paragraph:2:fragment:0");
    const elementTocResult = await document.getObject("docx:paragraph:3:fragment:0");
    const footnote = await document.getObject("docx:footnote:2:paragraph:0");
    const endnote = await document.getObject("docx:endnote:3:paragraph:0");
    const comment = await document.getObject("docx:comment:4:paragraph:0");
    assert.equal(image?.type, "image");
    assert.equal(image?.source.kind, "drawing");
    assert.equal(image?.source.drawingId, 7);
    assert.equal(header?.text, "First Header");
    assert.equal(header?.source.part, "word/headerFirst.xml");
    assert.equal(footer?.text, "First Footer");
    assert.equal(evenHeader?.text, "Even Header");
    assert.equal(defaultHeader?.text, "Header");
    assert.equal(headerImage?.type, "image");
    assert.equal(headerImage?.source.part, "word/headerFirst.xml");
    assert.equal(numbering?.text, "1.");
    assert.deepEqual(fieldTocResult?.actions, [{
      trigger: "click",
      kind: "command",
      action: "reveal",
      target: "docx:paragraph:4:fragment:0",
    }]);
    assert.deepEqual(elementTocResult?.actions, [{
      trigger: "click",
      kind: "command",
      action: "reveal",
      target: "docx:paragraph:5:fragment:0",
    }]);
    assert.equal(footnote?.text, "Footnote body");
    assert.equal(footnote?.source.part, "word/footnotes.xml");
    assert.equal(endnote?.text, "Endnote body");
    assert.equal(comment?.text, "Comment body");
    assert.equal(comment?.type, "text-box");
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.images, 2);
    assert.match(canvas.calls.text.map(({ text }) => text).join(""), /Rich document/);
    assert.match(canvas.calls.text.map(({ text }) => text).join(""), /First Header/);
    assert.deepEqual(
      canvas.calls.text
        .filter(({ text }) => text === "Field" || text === "Element")
        .map(({ fillStyle }) => fillStyle),
      ["rgba(0, 0, 0, 1)", "rgba(0, 0, 0, 1)"],
    );
    const hits = await document.hitTest({ unitIndex: 0, x: image.bounds.x + 1, y: image.bounds.y + 1 });
    assert.ok(hits.some(({ object }) => object.id === image.id));
    frame.bitmap.close();
    frame = await document.render({ unitIndex: footnote.unitIndex });
    frame.bitmap.close();
    frame = await document.render({ unitIndex: endnote.unitIndex });
    const noteText = canvas.calls.text.map(({ text }) => text).join("");
    assert.match(noteText, /Footnote body/);
    assert.match(noteText, /Endnote body/);
    assert.doesNotMatch(noteText, /Comment body/);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    canvas.restore();
  }
});

test("ODT styles, master stories, images, and text frames cross the public engine seam", async () => {
  const canvas = installCanvas();
  const engine = await createEngine();
  let document;
  let frame;
  try {
    document = await engine.open(odt);
    const header = await document.getObject("odt:header:0:0");
    const image = await document.getObject("odt:frame:0:visual");
    const textFrame = await document.getObject("odt:frame:1:text");
    const transformedFrame = await document.getObject("odt:frame:1:visual");
    const spanningCell = await document.getObject("odt:table:0:row:0:column:0");
    const rectangle = await document.getObject("odt:shape:0");
    const ellipse = await document.getObject("odt:shape:1");
    assert.equal(header?.text, "ODT Header");
    assert.equal(header?.source.kind, "text-range");
    assert.equal(image?.type, "image");
    assert.equal(image?.source.elementId, "image-frame");
    assert.equal(textFrame?.type, "text-box");
    assert.equal(textFrame?.text, "Frame text");
    assert.equal(transformedFrame?.source.elementId, "text-frame");
    assert.equal(spanningCell?.text, "Spanning cell");
    assert.equal(spanningCell?.source.elementId, "span-cell");
    assert.ok(spanningCell.bounds.width > 500);
    assert.ok(spanningCell.bounds.height >= 48);
    assert.equal(rectangle?.type, "shape");
    assert.equal(rectangle?.source.elementId, "rect-shape");
    assert.equal(ellipse?.type, "shape");
    frame = await document.render({ unitIndex: 0 });
    assert.equal(canvas.calls.images, 1);
    const renderedText = canvas.calls.text.map(({ text }) => text).join("");
    assert.match(renderedText, /ODT rich text/);
    assert.match(renderedText, /1\./);
    assert.match(renderedText, /First list item/);
    assert.match(renderedText, /Frame text/);
    const cellHits = await document.hitTest({
      unitIndex: spanningCell.unitIndex,
      x: spanningCell.bounds.x + 2,
      y: spanningCell.bounds.y + 2,
    });
    assert.ok(cellHits.some(({ object }) => object.id === spanningCell.id));
    const shapeHits = await document.hitTest({
      unitIndex: rectangle.unitIndex,
      x: rectangle.bounds.x + 10,
      y: rectangle.bounds.y + 2,
    });
    assert.ok(shapeHits.some(({ object }) => object.id === rectangle.id));
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    canvas.restore();
  }
});

for (const fixture of ["ooxml-gutter-left.docx", "ooxml-gutter-top.docx", "ooxml-gutter-right.docx", "ooxml-mirror-margins.docx"]) {
  test(`real OOXML page margins ${fixture}`, async () => {
    const engine = await createEngine();
    const document = await engine.open(await readFile(new URL(`fixtures/${fixture}`, import.meta.url)));
    try {
      const objects = await document.listObjects();
      const paragraphs = objects.filter(object => object.type === "paragraph" && object.text);
      if (fixture.includes("mirror")) {
        for (const [page, x] of [[0,120], [1,280]]) {
          const body = paragraphs.find(object => object.unitIndex === page && object.text.includes("Lorem"));
          assert.ok(Math.abs(body.bounds.x - x) < .1, JSON.stringify(body.bounds));
        }
      } else {
        const body = paragraphs.find(object => /Half in gutter|He heard quiet|hello/u.test(object.text));
        if (fixture.includes("left")) assert.ok(Math.abs(body.bounds.x - 144) < .1);
        if (fixture.includes("top")) assert.ok(body.bounds.y >= 143.9);
        if (fixture.includes("right")) {
          assert.ok(Math.abs(body.bounds.x - 24) < .1);
          assert.ok(Math.abs(body.bounds.width - (document.info.units[0].width - 120)) < .1);
        }
      }
    } finally { document.close(); engine.close(); }
  });
}

test("real OOXML reversed categories keep bars paired with their labels", async () => {
  const engine = await createEngine();
  const document = await engine.open(await readFile(new URL("fixtures/ooxml-reversed-categories.docx", import.meta.url)));
  try {
    const objects = await document.listObjects();
    const first = objects.find(object => object.text === "Concursuri (online/ offline)");
    const last = objects.find(object => object.text === "Saloane de inventică");
    assert.ok(first.bounds.y < last.bounds.y, "authored reverse order starts at the top");
    const bars = objects.filter(object => object.id.includes(":bar-0-"));
    assert.ok(bars.length >= 2);
    const firstBar = bars.find(object => object.id.endsWith("bar-0-0"));
    const lastBar = bars.find(object => object.id.endsWith("bar-0-9"));
    assert.ok(firstBar.bounds.y < lastBar.bounds.y);
    assert.ok(firstBar.bounds.width > lastBar.bounds.width);
  } finally { document.close(); engine.close(); }
});
