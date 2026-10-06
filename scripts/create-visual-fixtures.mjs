import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";

import { createZip } from "../tests/zip-fixture.mjs";

const presentationOutput = resolve("tests/fixtures/visual-baseline.pptx");
const spreadsheetOutput = resolve("tests/fixtures/frozen-panes.xlsx");

const shape = ({
  id,
  name,
  x,
  y,
  width,
  height,
  geometry = "rect",
  rotation,
  fill = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
  line = '<a:ln><a:noFill/></a:ln>',
  text = "",
  textRunAttributes = 'lang="zh-CN" sz="1800"',
  textRunProperties = '<a:solidFill><a:srgbClr val="17324D"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Noto Sans S Chinese"/>',
  paragraphProperties = '<a:pPr algn="l"/>',
}) => `
  <p:sp>
    <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm${rotation === undefined ? "" : ` rot="${rotation}"`}><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
      <a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>
      ${fill}${line}
    </p:spPr>
    ${text ? `<p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/><a:p>${paragraphProperties}<a:r><a:rPr ${textRunAttributes}>${textRunProperties}</a:rPr>${text}</a:r><a:endParaRPr lang="zh-CN" sz="1800"/></a:p></p:txBody>` : ""}
  </p:sp>`;

const title = `
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="571500" y="381000"/><a:ext cx="8001000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
    <p:txBody><a:bodyPr wrap="none" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="zh-CN" sz="3000" b="1"><a:solidFill><a:srgbClr val="17324D"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Noto Sans S Chinese"/></a:rPr><a:t>OfficeViewer 视觉基线</a:t></a:r><a:endParaRPr lang="zh-CN" sz="3000"/></a:p></p:txBody>
  </p:sp>`;

const richText = `
  <p:sp>
    <p:nvSpPr><p:cNvPr id="7" name="Rich text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="4191000" y="1657350"/><a:ext cx="4000500" cy="1524000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
    <p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" wrap="square" anchor="t"/><a:lstStyle/>
      <a:p><a:pPr algn="l"/><a:r><a:rPr lang="zh-CN" sz="2200" b="1"><a:solidFill><a:srgbClr val="17324D"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Noto Sans S Chinese"/></a:rPr><a:t>真实文字 · </a:t></a:r><a:r><a:rPr lang="en-US" sz="2200" i="1"><a:solidFill><a:srgbClr val="2F6BFF"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Rich Text</a:t></a:r><a:endParaRPr lang="zh-CN" sz="2200"/></a:p>
      <a:p><a:pPr algn="l" marL="0"/><a:r><a:rPr lang="zh-CN" sz="1500"><a:solidFill><a:srgbClr val="5F7185"/></a:solidFill><a:ea typeface="Noto Sans S Chinese"/></a:rPr><a:t>用于核对字体、字重、颜色、换行和基线。</a:t></a:r><a:endParaRPr lang="zh-CN" sz="1500"/></a:p>
    </p:txBody>
  </p:sp>`;

const entries = {
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
      <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
      <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    </Relationships>`,
  "docProps/core.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>OfficeViewer visual baseline</dc:title></cp:coreProperties>`,
  "docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>OfficeViewer fixture generator</Application><Slides>1</Slides></Properties>`,
  "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
      <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>`,
  "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`,
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld name="Visual baseline">
        <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F4F7FC"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
        <p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
          ${title}
          ${shape({id: 3, name: "Gradient card", x: 571500, y: 1524000, width: 2857500, height: 2095500, geometry: "roundRect", fill: '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="2F6BFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="69A7FF"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill>', line: '<a:ln w="19050"><a:solidFill><a:srgbClr val="2857CE"/></a:solidFill></a:ln>', text: '<a:t>渐变圆角卡片</a:t>', paragraphProperties: '<a:pPr algn="ctr"/>', textRunAttributes: 'lang="zh-CN" sz="1800" b="1"', textRunProperties: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ea typeface="Noto Sans S Chinese"/>'})}
          ${shape({id: 4, name: "Circle", x: 857250, y: 2095500, width: 762000, height: 762000, geometry: "ellipse", fill: '<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="90000"/></a:srgbClr></a:solidFill>', text: '<a:t>01</a:t>', paragraphProperties: '<a:pPr algn="ctr"/>', textRunAttributes: 'lang="en-US" sz="1900" b="1"', textRunProperties: '<a:solidFill><a:srgbClr val="2F6BFF"/></a:solidFill><a:latin typeface="Arial"/>'})}
          ${richText}
          ${shape({id: 8, name: "Rotated accent", x: 6905625, y: 3238500, width: 952500, height: 428625, rotation: 600000, fill: '<a:solidFill><a:srgbClr val="FFB020"/></a:solidFill>', line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="E49100"/></a:solidFill></a:ln>'})}
          ${shape({id: 9, name: "Footer rule", x: 571500, y: 4191000, width: 8001000, height: 0, geometry: "line", fill: '<a:noFill/>', line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="C7D2E1"/></a:solidFill></a:ln>'})}
          ${shape({id: 10, name: "Footer", x: 571500, y: 4305300, width: 8001000, height: 381000, fill: '<a:noFill/>', text: '<a:t>Reviewed read-only oracle · 960 × 540</a:t>', textRunAttributes: 'lang="en-US" sz="1100"', textRunProperties: '<a:solidFill><a:srgbClr val="7C8DA1"/></a:solidFill><a:latin typeface="Arial"/>'})}
        </p:spTree>
      </p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
    </p:sld>`,
};

const spreadsheetRows = Array.from({ length: 240 }, (_, index) => {
  const row = index + 1;
  if (row === 1) {
    return `<row r="1" ht="28" customHeight="1"><c r="A1" s="2" t="inlineStr"><is><t>区域</t></is></c><c r="B1" s="2" t="inlineStr"><is><t>负责人</t></is></c><c r="C1" s="2" t="inlineStr"><is><t>一月</t></is></c><c r="D1" s="2" t="inlineStr"><is><t>二月</t></is></c><c r="E1" s="2" t="inlineStr"><is><t>三月</t></is></c></row>`;
  }
  if (row === 2) {
    return `<row r="2" ht="24" customHeight="1"><c r="A2" s="3" t="inlineStr"><is><t>固定列 A</t></is></c><c r="B2" s="3" t="inlineStr"><is><t>固定列 B</t></is></c><c r="C2" s="3" t="inlineStr"><is><t>滚动数据</t></is></c><c r="D2" s="3" t="inlineStr"><is><t>滚动数据</t></is></c><c r="E2" s="3" t="inlineStr"><is><t>滚动数据</t></is></c></row>`;
  }
  const region = `区域 ${String(row - 2).padStart(3, "0")}`;
  const owner = `成员 ${(row - 3) % 12 + 1}`;
  return `<row r="${row}"><c r="A${row}" s="1" t="inlineStr"><is><t>${region}</t></is></c><c r="B${row}" s="1" t="inlineStr"><is><t>${owner}</t></is></c><c r="C${row}" s="4"><v>${row * 3}</v></c><c r="D${row}" s="4"><v>${row * 5}</v></c><c r="E${row}" s="4"><v>${row * 7}</v></c></row>`;
}).join("");

const spreadsheetEntries = {
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
  "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="冻结窗格验证" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>`,
  "xl/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="3"><font><sz val="11"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font><font><b/><color rgb="FF17324D"/><sz val="10"/><name val="Arial"/></font></fonts>
      <fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2F6BFF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEAF1FF"/><bgColor indexed="64"/></patternFill></fill></fills>
      <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD7E0EC"/></left><right style="thin"><color rgb="FFD7E0EC"/></right><top style="thin"><color rgb="FFD7E0EC"/></top><bottom style="thin"><color rgb="FFD7E0EC"/></bottom><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs>
    </styleSheet>`,
  "xl/worksheets/sheet1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:ZZ240"/>
      <sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="C3" sqref="C3"/></sheetView></sheetViews>
      <sheetFormatPr defaultRowHeight="20" defaultColWidth="10"/>
      <cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/><col min="3" max="702" width="12" customWidth="1"/></cols>
      <sheetData>${spreadsheetRows}</sheetData>
    </worksheet>`,
};

await mkdir(dirname(presentationOutput), { recursive: true });
await Promise.all([
  writeFile(presentationOutput, createZip(entries)),
  writeFile(spreadsheetOutput, createZip(spreadsheetEntries)),
]);
console.log(presentationOutput);
console.log(spreadsheetOutput);
