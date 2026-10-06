import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { createOfficeEngine } from "../dist/engine.js";
import { createZip } from "./zip-fixture.mjs";

const onePixelPng = Uint8Array.of(
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2,
  0, 0, 0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1,
  5, 1, 1, 39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
);

const pptx = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    </Relationships>`,
  "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
      <p:sldSz cx="9144000" cy="6858000"/>
    </p:presentation>`,
  "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
    </Relationships>`,
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Greeting"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Hello local document</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="7" name="One pixel"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdImage"/><a:srcRect l="25%" r="25000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm><a:off x="5486400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:pic>
      </p:spTree></p:cSld>
    </p:sld>`,
  "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
    </Relationships>`,
  "ppt/media/image1.png": onePixelPng,
});

const pptxWithLayoutPictureEntries = {
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
      <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
      <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
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
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr/><p:grpSpPr/>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="31" name="Slide picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdSlideImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm><a:off x="1371600" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm></p:spPr>
        </p:pic>
      </p:spTree></p:cSld>
    </p:sld>`,
  "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
      <Relationship Id="rIdSlideImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/slide.png"/>
    </Relationships>`,
  "ppt/slideLayouts/slideLayout1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld>
        <p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBackground"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>
        <p:spTree>
        <p:nvGrpSpPr/><p:grpSpPr/>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="11" name="Layout picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
        </p:pic>
        </p:spTree>
      </p:cSld>
    </p:sldLayout>`,
  "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdBackground" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/background.png"/>
      <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout.png"/>
      <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>`,
  "ppt/slideMasters/slideMaster1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr/><p:grpSpPr/>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="21" name="Master picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdMasterImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm></p:spPr>
        </p:pic>
      </p:spTree></p:cSld>
    </p:sldMaster>`,
  "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdMasterImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/master.png"/>
    </Relationships>`,
  "ppt/media/background.png": onePixelPng,
  "ppt/media/layout.png": onePixelPng,
  "ppt/media/master.png": onePixelPng,
  "ppt/media/slide.png": onePixelPng,
};

const pptxWithLayoutPicture = createZip(pptxWithLayoutPictureEntries);
const pptxWithPictureGlow = createZip({
  ...pptxWithLayoutPictureEntries,
  "ppt/slides/slide1.xml": pptxWithLayoutPictureEntries["ppt/slides/slide1.xml"].replace(
    "</p:spPr>",
    '<a:effectLst><a:glow rad="19050"><a:srgbClr val="FF0000"/></a:glow></a:effectLst></p:spPr>',
  ),
});

const pptxWithThemeBackground = createZip({
  ...pptxWithLayoutPictureEntries,
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
    </p:sld>`,
  "ppt/slideLayouts/slideLayout1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" showMasterSp="0">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
    </p:sldLayout>`,
  "ppt/slideMasters/slideMaster1.xml": pptxWithLayoutPictureEntries["ppt/slideMasters/slideMaster1.xml"]
    .replace("<p:cSld>", '<p:cSld><p:bg><p:bgRef idx="1003"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'),
  "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdMasterImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/master.png"/>
      <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
    </Relationships>`,
  "ppt/theme/theme1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <a:themeElements><a:fmtScheme><a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill>
        <a:blipFill><a:blip r:embed="rIdBackground"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
      </a:bgFillStyleLst></a:fmtScheme></a:themeElements>
    </a:theme>`,
  "ppt/theme/_rels/theme1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdBackground" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/theme-background.png"/>
    </Relationships>`,
  "ppt/media/theme-background.png": onePixelPng,
});

const pptxWithInheritedPicturePlaceholderEntries = {
  ...pptxWithLayoutPictureEntries,
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="31" name="Assigned picture"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="13"/></p:nvPr></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdSlideImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr/>
        </p:pic>
      </p:spTree></p:cSld>
    </p:sld>`,
  "ppt/slideLayouts/slideLayout1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="13" name="Picture placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="13"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="1143000"/><a:ext cx="2857500" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sldLayout>`,
  "ppt/slideMasters/slideMaster1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>`,
};

const pptxWithInheritedPicturePlaceholder = createZip(pptxWithInheritedPicturePlaceholderEntries);

const pptxWithIncompletePictureTransform = createZip({
  ...pptxWithInheritedPicturePlaceholderEntries,
  "ppt/slides/slide1.xml": pptxWithInheritedPicturePlaceholderEntries["ppt/slides/slide1.xml"]
    .replace("<p:spPr/>", '<p:spPr><a:xfrm><a:off y="1143000"/></a:xfrm></p:spPr>'),
});

const pptxWithRotatedInheritedPicturePlaceholder = createZip({
  ...pptxWithInheritedPicturePlaceholderEntries,
  "ppt/slideLayouts/slideLayout1.xml": pptxWithInheritedPicturePlaceholderEntries["ppt/slideLayouts/slideLayout1.xml"]
    .replace("<a:xfrm>", '<a:xfrm rot="5400000">')
    .replace(
      "</p:spPr>",
      '<a:effectDag name="Placeholder effects"><a:xfrm sx="100000" sy="100000" tx="0" ty="0"/></a:effectDag></p:spPr>',
    ),
});

const pptxWithReorderedPicturePlaceholders = createZip({
  ...pptxWithInheritedPicturePlaceholderEntries,
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="35" name="First replacement"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="15"/></p:nvPr></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdSlideImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr/>
        </p:pic>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="33" name="Second replacement"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="13"/></p:nvPr></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdSlideImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr/>
        </p:pic>
      </p:spTree></p:cSld>
    </p:sld>`,
  "ppt/slideLayouts/slideLayout1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="13" name="First placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="13"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="1143000"/><a:ext cx="2857500" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="15" name="Second placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="15"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="1143000"/><a:ext cx="2857500" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sldLayout>`,
});

const pptxWithPictureEffectTransform = createZip({
  ...pptxWithLayoutPictureEntries,
  "ppt/slides/slide1.xml": pptxWithLayoutPictureEntries["ppt/slides/slide1.xml"].replace(
    "</p:spPr>",
    '<a:effectDag name="Picture effects"><a:xfrm sx="100000" sy="100000" tx="0" ty="0"/></a:effectDag></p:spPr>',
  ),
});

const pptxWithPicturePlaceholderReplacedByShape = createZip({
  ...pptxWithLayoutPictureEntries,
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="31" name="Replacement shape"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="13"/></p:nvPr></p:nvSpPr>
          <p:spPr/>
        </p:sp>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="32" name="Reused image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdReusedImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm><a:off x="3810000" y="1905000"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>
        </p:pic>
      </p:spTree></p:cSld>
    </p:sld>`,
  "ppt/slides/_rels/slide1.xml.rels": pptxWithLayoutPictureEntries["ppt/slides/_rels/slide1.xml.rels"]
    .replace(
      "</Relationships>",
      '<Relationship Id="rIdReusedImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout.png"/></Relationships>',
    ),
  "ppt/slideLayouts/slideLayout1.xml": pptxWithLayoutPictureEntries["ppt/slideLayouts/slideLayout1.xml"]
    .replace("<p:nvPr/>", '<p:nvPr><p:ph type="pic" idx="13"/></p:nvPr>'),
});

const pptxWithSvgPicture = createZip({
  ...pptxWithLayoutPictureEntries,
  "[Content_Types].xml": pptxWithLayoutPictureEntries["[Content_Types].xml"].replace(
    "</Types>",
    '<Default Extension="svg" ContentType="image/svg+xml"/></Types>',
  ),
  "ppt/slides/slide1.xml": pptxWithLayoutPictureEntries["ppt/slides/slide1.xml"]
    .replace(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"',
    )
    .replace(
      '<a:blip r:embed="rIdSlideImage"/>',
      '<a:blip r:embed="rIdSlideImage"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rIdPreferredSvg"/></a:ext></a:extLst></a:blip>',
    ),
  "ppt/slides/_rels/slide1.xml.rels": pptxWithLayoutPictureEntries["ppt/slides/_rels/slide1.xml.rels"]
    .replace(
      "</Relationships>",
      '<Relationship Id="rIdPreferredSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/slide.svg"/></Relationships>',
    ),
  "ppt/media/slide.svg": new TextEncoder().encode(
    '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="1" viewBox="0 0 2 1"><rect width="2" height="1" fill="#e53935"/></svg>',
  ),
});

const pptxWithUnsupportedSlideBackground = createZip({
  ...pptxWithLayoutPictureEntries,
  "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld>
        <p:bg><p:bgPr><a:pattFill prst="pct5"><a:fgClr><a:srgbClr val="102030"/></a:fgClr><a:bgClr><a:srgbClr val="F0F0F0"/></a:bgClr></a:pattFill></p:bgPr></p:bg>
        <p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree>
      </p:cSld>
    </p:sld>`,
});

const odp = createZip({
  mimetype: "application/vnd.oasis.opendocument.presentation",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
      <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>
      <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
    </manifest:manifest>`,
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
      <office:automatic-styles>
        <style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/></style:page-layout>
      </office:automatic-styles>
      <office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>
    </office:document-styles>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
      <office:body><office:presentation>
        <draw:page draw:name="Slide 1" draw:master-page-name="Default">
          <draw:frame draw:id="greeting" draw:name="Greeting" svg:x="1in" svg:y="1in" svg:width="4in" svg:height="1in">
            <draw:text-box><text:p>Hello open document</text:p></draw:text-box>
          </draw:frame>
        </draw:page>
      </office:presentation></office:body>
    </office:document-content>`,
}, { compress: false });

const fodp = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
  <office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    office:mimetype="application/vnd.oasis.opendocument.presentation">
    <office:automatic-styles>
      <style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/></style:page-layout>
    </office:automatic-styles>
    <office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>
    <office:body><office:presentation>
      <draw:page draw:name="Flat Slide" draw:master-page-name="Default">
        <draw:frame draw:id="flat-greeting" svg:x="1in" svg:y="1in" svg:width="4in" svg:height="1in">
          <draw:text-box><text:p>Hello flat open document</text:p></draw:text-box>
        </draw:frame>
      </draw:page>
    </office:presentation></office:body>
  </office:document>`);

const odpWithMeasuredTable = createZip({
  mimetype: "application/vnd.oasis.opendocument.presentation",
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
      <office:automatic-styles>
        <style:page-layout style:name="PageLayout">
          <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
        </style:page-layout>
        <style:style style:name="NarrowColumn" style:family="table-column">
          <style:table-column-properties style:column-width="0.35in"/>
        </style:style>
        <style:style style:name="ShortRow" style:family="table-row">
          <style:table-row-properties style:row-height="0.1in"/>
        </style:style>
        <style:style style:name="Cell" style:family="table-cell">
          <style:table-cell-properties fo:padding="0in"/>
          <style:text-properties fo:font-family="Arial" fo:font-size="10pt"/>
        </style:style>
      </office:automatic-styles>
      <office:master-styles>
        <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
      </office:master-styles>
    </office:document-styles>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:table="table" xmlns:text="text">
      <office:body><office:presentation><draw:page draw:master-page-name="Master">
        <draw:frame svg:x="1in" svg:y="1in" svg:width="0.35in" svg:height="2in">
          <table:table>
            <table:table-column table:style-name="NarrowColumn"/>
            <table:table-row table:style-name="ShortRow">
              <table:table-cell table:style-name="Cell"><text:p>WWWW</text:p></table:table-cell>
            </table:table-row>
          </table:table>
        </draw:frame>
      </draw:page></office:presentation></office:body>
    </office:document-content>`,
}, { compress: false });

const embeddedWav = Uint8Array.of(
  0x52, 0x49, 0x46, 0x46, 0x04, 0x00, 0x00, 0x00,
  0x57, 0x41, 0x56, 0x45,
);

const odpWithEmbeddedAudio = createZip({
  mimetype: "application/vnd.oasis.opendocument.presentation",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
      <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>
      <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="Media/sound.wav" manifest:media-type="audio/wav"/>
    </manifest:manifest>`,
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
      <office:automatic-styles>
        <style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/></style:page-layout>
      </office:automatic-styles>
      <office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>
    </office:document-styles>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">
      <office:body><office:presentation>
        <draw:page draw:name="Slide 1" draw:master-page-name="Default">
          <draw:frame draw:id="audio" svg:x="1in" svg:y="1in" svg:width="1in" svg:height="1in">
            <draw:plugin xlink:href="Media/sound.wav" draw:mime-type="audio/wav"/>
          </draw:frame>
        </draw:page>
      </office:presentation></office:body>
    </office:document-content>`,
  "Media/sound.wav": embeddedWav,
}, { compress: false });

const xlsx = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
  "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`,
  "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
    </Relationships>`,
  "xl/sharedStrings.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
      <si><t>Hello sheet</t></si>
    </sst>`,
  "xl/worksheets/sheet1.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:B2"/>
      <sheetData>
        <row r="1"><c r="A1" t="s"><v>0</v></c></row>
        <row r="2"><c r="B2"><v>42</v></c></row>
      </sheetData>
    </worksheet>`,
});

const odt = createZip({
  mimetype: "application/vnd.oasis.opendocument.text",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
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
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
      <office:body><office:text>
        <text:p text:style-name="Standard">Hello text document</text:p>
        <table:table xml:id="matrix">
          <table:table-row>
            <table:table-cell xml:id="a1"><text:p>A1</text:p></table:table-cell>
            <table:table-cell xml:id="b1"><text:p>B1</text:p></table:table-cell>
          </table:table-row>
          <table:table-row>
            <table:table-cell xml:id="a2"><text:p>A2</text:p></table:table-cell>
            <table:table-cell xml:id="b2"><text:p>B2</text:p></table:table-cell>
          </table:table-row>
        </table:table>
      </office:text></office:body>
    </office:document-content>`,
}, { compress: false });

const odtWithCharacterAnchoredChart = createZip({
  mimetype: "application/vnd.oasis.opendocument.text",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
      <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
      <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/>
      <manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>
    </manifest:manifest>`,
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
      <office:automatic-styles>
        <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin="1in"/></style:page-layout>
      </office:automatic-styles>
      <office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>
    </office:document-styles>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">
      <office:body><office:text>
        <text:p>Paragraph before the chart</text:p>
        <text:p><draw:frame draw:name="Object1" text:anchor-type="char" svg:x="1in" svg:y="10pt" svg:width="4in" svg:height="2in">
          <draw:object xlink:href="./Object 1"/>
        </draw:frame></text:p>
      </office:text></office:body>
    </office:document-content>`,
  "Object 1/content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
      <office:automatic-styles>
        <style:style style:name="series-blue" style:family="chart"><style:graphic-properties draw:fill-color="#004586"/></style:style>
        <style:style style:name="series-orange" style:family="chart"><style:graphic-properties draw:fill-color="#ff420e"/></style:style>
        <style:style style:name="series-yellow" style:family="chart"><style:graphic-properties draw:fill-color="#ffd320"/></style:style>
      </office:automatic-styles>
      <office:body><office:chart><chart:chart chart:class="chart:bar"><chart:plot-area chart:data-source-has-labels="both">
        <chart:axis chart:dimension="x"><chart:categories table:cell-range-address="local-table.$A$2:.$A$5"/></chart:axis>
        <chart:axis chart:dimension="y"><chart:grid chart:class="major"/></chart:axis>
        <chart:series chart:style-name="series-blue" chart:values-cell-range-address="local-table.$B$2:.$B$5" chart:label-cell-address="local-table.$B$1"/>
        <chart:series chart:style-name="series-orange" chart:values-cell-range-address="local-table.$C$2:.$C$5" chart:label-cell-address="local-table.$C$1"/>
        <chart:series chart:style-name="series-yellow" chart:values-cell-range-address="local-table.$D$2:.$D$5" chart:label-cell-address="local-table.$D$1"/>
        <table:table table:name="local-table">
          <table:table-row><table:table-cell/><table:table-cell office:value-type="string"><text:p>Column 1</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Column 2</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Column 3</text:p></table:table-cell></table:table-row>
          <table:table-row><table:table-cell office:value-type="string"><text:p>Row 1</text:p></table:table-cell><table:table-cell office:value="9.1"/><table:table-cell office:value="3.2"/><table:table-cell office:value="4.54"/></table:table-row>
          <table:table-row><table:table-cell office:value-type="string"><text:p>Row 2</text:p></table:table-cell><table:table-cell office:value="2.4"/><table:table-cell office:value="8.8"/><table:table-cell office:value="9.65"/></table:table-row>
          <table:table-row><table:table-cell office:value-type="string"><text:p>Row 3</text:p></table:table-cell><table:table-cell office:value="3.1"/><table:table-cell office:value="1.5"/><table:table-cell office:value="3.7"/></table:table-row>
          <table:table-row><table:table-cell office:value-type="string"><text:p>Row 4</text:p></table:table-cell><table:table-cell office:value="4.3"/><table:table-cell office:value="9.02"/><table:table-cell office:value="6.2"/></table:table-row>
        </table:table>
      </chart:plot-area></chart:chart></office:chart></office:body>
    </office:document-content>`,
}, { compress: false });

const odtWithLegacyMathDtd = createZip({
  mimetype: "application/vnd.oasis.opendocument.text",
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="office"/>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
      <office:body><office:text><text:p><draw:frame svg:width="1in" svg:height="1in">
        <draw:object xlink:href="./Object 1"/>
      </draw:frame></text:p></office:text></office:body>
    </office:document-content>`,
  "Object 1/content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <!DOCTYPE math:math PUBLIC "-//OpenOffice.org//DTD Modified W3C MathML 1.01//EN" "math.dtd">
    <math:math xmlns:math="http://www.w3.org/1998/Math/MathML"><math:mn>1</math:mn></math:math>`,
}, { compress: false });

const odtWithFormTextBox = createZip({
  mimetype: "application/vnd.oasis.opendocument.text",
  "styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-styles xmlns:office="office"/>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:form="form" xmlns:svg="svg">
      <office:automatic-styles><style:style xmlns:style="style" style:name="gr1" style:family="graphic"><style:graphic-properties/></style:style></office:automatic-styles>
      <office:body><office:text><office:forms><form:form><form:textarea xml:id="control1"/></form:form></office:forms>
        <text:p><draw:control draw:name="Text box 1" draw:style-name="gr1" draw:control="control1" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm"/></text:p>
      </office:text></office:body>
    </office:document-content>`,
}, { compress: false });

const ods = createZip({
  mimetype: "application/vnd.oasis.opendocument.spreadsheet",
  "META-INF/manifest.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
      <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
      <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
    </manifest:manifest>`,
  "content.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
      <office:body><office:spreadsheet><table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>Hello ODS</text:p></table:table-cell>
          <table:table-cell office:value-type="float" office:value="42.5"/>
        </table:table-row>
      </table:table></office:spreadsheet></office:body>
    </office:document-content>`,
}, { compress: false });

const fods = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
  <office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
    <office:body><office:spreadsheet><table:table table:name="Flat Sheet">
      <table:table-row><table:table-cell office:value-type="string"><text:p>Hello FODS</text:p></table:table-cell></table:table-row>
    </table:table></office:spreadsheet></office:body>
  </office:document>`);

async function createOdfEngine() {
  const wasm = await readFile(new URL("../dist/office-viewer-odf.wasm", import.meta.url));
  return createOfficeEngine({ wasm, execution: "inline" });
}

test("ODT contextual spacing keeps the supplied section document on one page", async () => {
  const engine = await createOdfEngine();
  const document = await engine.open(await readFile(new URL("./fixtures/oasis-contextual-spacing-section.odt", import.meta.url)));
  try {
    assert.equal(document.info.units.length, 1);
    assert.ok(Math.abs(document.info.units[0].width - 210.01 * 96 / 25.4) < 0.1);
  } finally {
    document.close();
    engine.close();
  }
});

const docx = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
  "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body>
        <w:p w14:paraId="0A1B2C3D"><w:pPr><w:spacing w:line="360"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Hello DOCX</w:t></w:r></w:p>
        <w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body>
    </w:document>`,
});

const docxWithCharacterScale = docxCompatibilityFixture(`
  <w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:body>
    <w:p><w:r><w:rPr><w:w w:val="33"/></w:rPr><w:t>x</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`);

const docxWithDrawingTextBox = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
  "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>`,
  "word/styles.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
      <w:style w:type="paragraph" w:styleId="ProfileHeading">
        <w:name w:val="Profile Heading"/><w:basedOn w:val="Normal"/>
        <w:pPr><w:spacing w:before="360" w:after="120" w:line="240"/></w:pPr>
        <w:rPr><w:sz w:val="40"/></w:rPr>
      </w:style>
    </w:styles>`,
  "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml">
      <w:body>
        <w:p w14:paraId="A1B2C3D4"><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>476250</wp:posOffset></wp:positionV><wp:extent cx="5715000" cy="1714500"/><wp:wrapTopAndBottom/><wp:docPr id="7" name="Profile"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>
          <w:p><w:pPr><w:pStyle w:val="ProfileHeading"/></w:pPr><w:r><w:t>Basic information</w:t></w:r></w:p>
          <w:tbl><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>Name:</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Ju Wang</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Gender:</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Male</w:t></w:r></w:p></w:tc></w:tr>
            <w:tr><w:tc><w:p><w:r><w:t>Year of birth:</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>1984</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Experience:</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>15 years</w:t></w:r></w:p></w:tc></w:tr>
          </w:tbl>
        </w:txbxContent></wps:txbx><wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/></wps:wsp></a:graphicData></a:graphic><wp:sizeRelH><wp:pctWidth>0</wp:pctWidth></wp:sizeRelH><wp:sizeRelV><wp:pctHeight>0</wp:pctHeight></wp:sizeRelV></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Fallback duplicate</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
        <w:p w14:paraId="D4C3B2A1"><w:pPr><w:pStyle w:val="ProfileHeading"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body>
    </w:document>`,
});

const docxWithPictureGlow = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="7" name="Picture"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill>
        <pic:spPr><a:prstGeom prst="rect"/><a:effectLst><a:glow rad="19050"><a:srgbClr val="FF0000"/></a:glow></a:effectLst></pic:spPr>
      </pic:pic></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`,
  `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/>`,
  { "word/media/image.png": onePixelPng },
);

const docxWithInvalidNestedParagraph = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
  "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:p><w:r><w:t>Invalid nesting</w:t></w:r></w:p></w:p>
    </w:body></w:document>`,
});

function docxCompatibilityFixture(documentXml, relationships = "", parts = {}) {
  return createZip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>`,
    "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>${documentXml}`,
    ...parts,
  });
}

const docxWithLockedCanvas = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Diagram"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"><lc:lockedCanvas>
        <a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr>
        <a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></a:grpSpPr>
        <a:sp><a:nvSpPr><a:cNvPr id="2" name="Box"/><a:cNvSpPr/></a:nvSpPr><a:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        </a:spPr><a:txSp><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>Box</a:t></a:r></a:p></a:txBody></a:txSp></a:sp>
      </lc:lockedCanvas></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p><w:sectPr/>
  </w:body></w:document>`,
);

const docxWithCachedBarChartKind = (chartKind, shape = "", plotSuffix = "") => docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="4099560" cy="2059940"/><wp:docPr id="1" name="Sales chart"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`, `
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`, {
  "word/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>${chartKind === "bar3DChart" ? '<c:view3D><c:perspective val="30"/></c:view3D>' : ""}<c:plotArea><c:${chartKind}><c:ser>
    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Column 1</c:v></c:pt></c:strCache></c:strRef></c:tx>
    <c:spPr><a:solidFill><a:srgbClr val="004586"/></a:solidFill></c:spPr>
    <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Row 1</c:v></c:pt><c:pt idx="1"><c:v>Row 2</c:v></c:pt></c:strCache></c:strRef></c:cat>
    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>9.1</c:v></c:pt><c:pt idx="1"><c:v>2.4</c:v></c:pt></c:numCache></c:numRef></c:val>
    <c:axId val="1"/><c:axId val="2"/>
  </c:ser>${shape ? `<c:shape val="${shape}"/>` : ""}</c:${chartKind}>${plotSuffix}</c:plotArea><c:legend/></c:chart></c:chartSpace>`,
});
const docxWithCachedBarChart = docxWithCachedBarChartKind("barChart");
const docxWithCached3dBarChart = docxWithCachedBarChartKind("bar3DChart");
const docxWithCached3dBarChartAndDataTable = docxWithCachedBarChartKind(
  "bar3DChart",
  "",
  '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable>',
);
const docxWithCachedCylinderChart = docxWithCachedBarChartKind("bar3DChart", "cylinder");

const docxWithCached3dPieChart = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="5055235" cy="3200400"/><wp:docPr id="5" name="Enrollment chart"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p><w:sectPr/>
  </w:body></w:document>`, `
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`, {
  "word/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart>
    <c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Grade 6 WL Enrollment 2015-16</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:view3D><c:rotX val="30"/><c:rotY val="0"/></c:view3D><c:plotArea><c:pie3DChart><c:varyColors val="1"/><c:ser><c:explosion val="25"/>
      <c:dLbls><c:dLblPos val="bestFit"/><c:showVal val="0"/><c:showCatName val="1"/><c:showPercent val="1"/><c:showLeaderLines val="1"/></c:dLbls>
      <c:cat><c:strLit>${["Immersion", "Transition Spanish", "Span for Flnt Spkrs", "Intro to [Lang]", "No WL"].map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:strLit></c:cat>
      <c:val><c:numLit>${[114, 338, 134, 471, 774].map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:numLit></c:val>
    </c:ser></c:pie3DChart></c:plotArea></c:chart></c:chartSpace>`,
});

const docxWithInheritedPiePointColor = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="5759450" cy="3239770"/><wp:docPr id="7" name="Inherited pie color"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p><w:sectPr/>
  </w:body></w:document>`, `
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`, {
  "word/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:pieChart><c:varyColors val="1"/><c:ser>
    <c:spPr><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></c:spPr>
    <c:dPt><c:idx val="0"/></c:dPt>
    ${[[1, "FF420E"], [2, "FFD320"], [3, "579D1C"]].map(([index, color]) => `<c:dPt><c:idx val="${index}"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:dPt>`).join("")}
    <c:cat><c:strLit>${[1, 2, 3, 4].map((value, index) => `<c:pt idx="${index}"><c:v>Row ${value}</c:v></c:pt>`).join("")}</c:strLit></c:cat>
    <c:val><c:numLit>${[9.1, 2.4, 3.1, 4.3].map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:numLit></c:val>
  </c:ser></c:pieChart></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`,
});

const docxWithCachedUpDownBarsChart = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="6" name="Up/down bars chart"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p><w:sectPr/>
  </w:body></w:document>`, `
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`, {
  "word/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart><c:grouping val="percentStacked"/>
    ${[
      ["Series 1", [4.3, 2.5, 3.5, 4.5]],
      ["Series 2", [2.4, 4.4, 1.8, 2.8]],
      ["Series 3", [2, 2, 3, 5]],
    ].map(([name, values]) => `<c:ser><c:tx><c:v>${name}</c:v></c:tx><c:marker><c:symbol val="none"/></c:marker>
      <c:cat><c:strLit>${[1, 2, 3, 4].map((value, index) => `<c:pt idx="${index}"><c:v>Category ${value}</c:v></c:pt>`).join("")}</c:strLit></c:cat>
      <c:val><c:numLit>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:numLit></c:val></c:ser>`).join("")}
    <c:upDownBars><c:gapWidth val="150"/><c:upBars><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:upBars><c:downBars/></c:upDownBars>
  </c:lineChart><c:catAx><c:majorTickMark val="out"/></c:catAx><c:valAx><c:numFmt formatCode="0%"/><c:majorGridlines/><c:majorTickMark val="out"/><c:crossBetween val="between"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`,
});

const docxWithXMarkerChart = (scatter) => docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="8" name="X marker chart"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p><w:sectPr/>
  </w:body></w:document>`, `
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`, {
  "word/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>
    <c:${scatter ? "scatterChart" : "lineChart"}>${scatter ? '<c:scatterStyle val="lineMarker"/>' : ""}<c:ser><c:tx><c:v>Series 1</c:v></c:tx>
      <c:marker><c:symbol val="x"/><c:size val="7"/></c:marker>
      ${scatter ? '<c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:xVal>' : '<c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>'}
      <c:${scatter ? "yVal" : "val"}><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:${scatter ? "yVal" : "val"}>
    </c:ser></c:${scatter ? "scatterChart" : "lineChart"}><c:valAx/><c:valAx/></c:plotArea></c:chart></c:chartSpace>`,
});

const docxWithCachedHorizontalConeChart = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Cone chart"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p><w:sectPr/>
  </w:body></w:document>`, `
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`, {
  "word/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:bar3DChart><c:barDir val="bar"/>
    <c:ser><c:tx><c:v>Series 1</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>Category 1</c:v></c:pt><c:pt idx="1"><c:v>Category 2</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>4.3</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt></c:numLit></c:val></c:ser>
    <c:shape val="cone"/></c:bar3DChart></c:plotArea><c:legend/></c:chart></c:chartSpace>`,
});

const docxWithCachedAreaChart = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Area chart"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p><w:sectPr/>
  </w:body></w:document>`, `
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`, {
  "word/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:areaChart><c:grouping val="percentStacked"/>
    <c:ser><c:tx><c:v>Series 1</c:v></c:tx><c:dLbls><c:showVal val="1"/></c:dLbls><c:cat><c:strLit><c:pt idx="0"><c:v>Day 1</c:v></c:pt><c:pt idx="1"><c:v>Day 2</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>32</c:v></c:pt><c:pt idx="1"><c:v>28</c:v></c:pt></c:numLit></c:val></c:ser>
    <c:ser><c:tx><c:v>Series 2</c:v></c:tx><c:dLbls><c:showVal val="1"/></c:dLbls><c:cat><c:strLit><c:pt idx="0"><c:v>Day 1</c:v></c:pt><c:pt idx="1"><c:v>Day 2</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>12</c:v></c:pt></c:numLit></c:val></c:ser>
  </c:areaChart></c:plotArea><c:legend/></c:chart></c:chartSpace>`,
});

const docxWithCachedRadarChart = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Radar chart"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p><w:sectPr/>
  </w:body></w:document>`, `
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`, {
  "word/charts/chart1.xml": `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title/><c:plotArea><c:radarChart>
    <c:ser><c:tx><c:v>Series 1</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>East</c:v></c:pt><c:pt idx="2"><c:v>South</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>32</c:v></c:pt><c:pt idx="1"><c:v>28</c:v></c:pt><c:pt idx="2"><c:v>12</c:v></c:pt></c:numLit></c:val></c:ser>
    <c:ser><c:tx><c:v>Series 2</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>East</c:v></c:pt><c:pt idx="2"><c:v>South</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>21</c:v></c:pt><c:pt idx="2"><c:v>28</c:v></c:pt></c:numLit></c:val></c:ser>
  </c:radarChart></c:plotArea><c:legend><c:legendPos val="t"/></c:legend></c:chart></c:chartSpace>`,
});

const docxWithChartExWaterfall = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><mc:AlternateContent><mc:Choice Requires="cx1"><w:drawing><wp:inline>
      <wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Waterfall chart"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart r:id="rIdChartEx"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></mc:Choice><mc:Fallback><w:drawing><wp:inline>
      <wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Waterfall preview"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdPreview"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
    </wp:inline></w:drawing></mc:Fallback></mc:AlternateContent></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`, `
  <Relationship Id="rIdChartEx" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="charts/chartEx1.xml"/>
  <Relationship Id="rIdPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/preview.png"/>`, {
  "word/charts/chartEx1.xml": `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="4"><cx:pt idx="0">100</cx:pt><cx:pt idx="1">-20</cx:pt><cx:pt idx="2">50</cx:pt><cx:pt idx="3">130</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t"/><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:tx><cx:txData><cx:v>Series1</cx:v></cx:txData></cx:tx><cx:dataLabels pos="inEnd"/><cx:dataId val="0"/><cx:layoutPr><cx:subtotals><cx:idx val="0"/><cx:idx val="3"/></cx:subtotals></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea><cx:legend pos="t"/></cx:chart></cx:chartSpace>`,
  "word/charts/_rels/chartEx1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="../style1.xml"/><Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="../colors1.xml"/></Relationships>`,
  "word/style1.xml": `<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="372"/>`,
  "word/colors1.xml": `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="10"><a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/></cs:colorStyle>`,
  "word/media/preview.png": onePixelPng,
});

const docxWithSectionFooterOverrides = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:pPr><w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr><w:r><w:t>First section</w:t></w:r></w:p>
    <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second section</w:t></w:r></w:p>
    <w:sectPr><w:footerReference w:type="default" r:id="rIdFooter2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`, `
  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>`, {
  "word/footer1.xml": `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>First footer</w:t></w:r></w:p></w:ftr>`,
  "word/footer2.xml": `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Second footer</w:t></w:r></w:p></w:ftr>`,
});

const docxWithNestedTable = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc>
      <w:p><w:r><w:t>Outer cell</w:t></w:r></w:p>
      <w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Nested cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      <w:p/>
    </w:tc></w:tr></w:tbl>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`);

function numberedDocx(levelAttributes, start) {
  return docxCompatibilityFixture(`
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Bullet item</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
    </w:body></w:document>`, `
    <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`, {
    "word/numbering.xml": `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl ${levelAttributes}><w:start w:val="${start}"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
  });
}

const docxWithZeroNumberingStart = numberedDocx(`w:ilvl="0"`, 0);
const docxWithImplicitNumberingLevel = numberedDocx("", 1);

const docxWithZeroHeightHeaderDrawing = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:r><w:t>Visible body</w:t></w:r></w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`, `
  <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`, {
  "word/header1.xml": `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="0"/><wp:docPr id="1" name="Hidden header drawing"/></wp:inline></w:drawing></w:r></w:p></w:hdr>`,
});

const docxWithLegacyVmlTextBox = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>
    <w:p><w:r><w:t>Visible body</w:t></w:r><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Legacy text box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`);

const docxWithEmptyLegacyVmlImageData = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>
    <w:p><w:r><w:t>Visible body</w:t></w:r><w:r><w:pict><v:shape><v:imagedata o:title=""/><v:textbox><w:txbxContent><w:p><w:r><w:t>Legacy text box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body></w:document>`);

const docxWithDefaultMainContentType = createZip({
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
  "word/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:t>Default content type</w:t></w:r></w:p>
    </w:body></w:document>`,
});

const docxWithWordCompatibilityMeasures = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="-720"/></w:tabs><w:ind w:start="36pt"/><w:spacing w:line="-12pt" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:sz w:val="20pt"/><w:color w:themeColor="dark1"/></w:rPr><w:t>Signed measures</w:t></w:r></w:p>
    <w:tbl><w:tblPr><w:tblCellMar><w:left w:w="-80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:fill="#e6e6e6"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="-23" w:right="8pt" w:bottom="-23" w:left="1133.8582677165355"/></w:sectPr>
  </w:body></w:document>`);

const docxWithNonDisplayMetadataText = docxCompatibilityFixture(`
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>
    <w:p><w:r><w:t>Visible text</w:t></w:r><w:r><w:object><o:FieldCodes>\\s</o:FieldCodes></w:object></w:r><w:fldData>opaque metadata</w:fldData></w:p>
  </w:body></w:document>`, `
  <Relationship Id="rIdBookmark" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="#bookmark"/>`);

test("opens compressed PPTX bytes with source-mapped, hittable objects", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(pptx);

  assert.equal(document.info.kind, "presentation");
  assert.equal(document.info.format, "pptx");
  assert.deepEqual(document.info.units, [
    {
      type: "slide",
      index: 0,
      id: "unit:0",
      name: "Slide 1",
      width: 960,
      height: 720,
      sourceId: "256",
      sourcePart: "ppt/slides/slide1.xml",
      slideNumber: 1,
      hidden: false,
    },
  ]);

  const hits = await document.hitTest({ unitIndex: 0, x: 120, y: 120 });
  assert.equal(hits.length, 1);
  assert.equal(hits[0].object.type, "text-box");
  assert.equal(hits[0].object.text, "Hello local document");
  assert.equal(hits[0].object.source.format, "pptx");
  assert.equal(hits[0].object.source.kind, "shape");
  assert.equal(hits[0].object.source.shapeId, 2);
  assert.equal(hits[0].object.source.mapping, "exact");
  assert.deepEqual(hits[0].object.bounds, { x: 96, y: 96, width: 384, height: 96 });
  assert.equal(Object.isFrozen(document.info), true);
  assert.equal(Object.isFrozen(document.info.units), true);
  assert.equal(Object.isFrozen(document.info.units[0]), true);
  assert.equal(Object.isFrozen(document.diagnostics()), true);
  assert.equal(Object.isFrozen(hits), true);
  assert.equal(Object.isFrozen(hits[0]), true);
  assert.equal(Object.isFrozen(hits[0].ancestors), true);
  assert.equal(Object.isFrozen(hits[0].object), true);
  assert.equal(Object.isFrozen(hits[0].object.bounds), true);
  assert.equal(Object.isFrozen(hits[0].object.source), true);

  document.close();
  engine.close();
});

test("opens an embedded PPTX raster image with exact shape mapping and hit testing", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptx);
    const hits = await document.hitTest({ unitIndex: 0, x: 600, y: 120 });

    assert.equal(hits.length, 1);
    const image = hits[0].object;
    assert.equal(image.type, "image");
    assert.deepEqual(image.bounds, { x: 576, y: 96, width: 96, height: 96 });
    assert.equal(image.source.format, "pptx");
    assert.equal(image.source.kind, "shape");
    assert.equal(image.source.shapeId, 7);
    assert.equal(image.source.mapping, "exact");
    assert.equal(await document.getObject(image.id), image);
  } finally {
    document?.close();
    engine.close();
  }
});

test("applies PPTX picture source cropping through the public render API", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const drawCalls = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    beginPath() {}, ellipse() {}, moveTo() {}, lineTo() {}, rect() {}, fill() {}, stroke() {},
    clip() {}, fillText() {}, measureText() { return { width: 0 }; },
    drawImage(...args) { drawCalls.push(args); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 200, height: 100, close() {} }),
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline", fontPolicy: "local-first" });
  let document;
  let frame;

  try {
    document = await engine.open(pptx);
    frame = await document.render({ unitIndex: 0 });
    const pictureCall = drawCalls.find((args) => args.includes(576) && args.includes(96));

    assert.ok(pictureCall);
    assert.deepEqual(pictureCall.slice(1), [50, 0, 100, 100, 576, 96, 96, 96]);
    assert.deepEqual(frame.diagnostics.map(({ code }) => code), ["FONT_SUBSTITUTED"]);
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

test("inherits embedded raster pictures from a PPTX slide layout", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithLayoutPicture);
    const hits = await document.hitTest({ unitIndex: 0, x: 120, y: 120, limit: 1 });

    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.type, "image");
    assert.deepEqual(hits[0].object.bounds, { x: 96, y: 96, width: 96, height: 96 });
    assert.equal(hits[0].object.source.part, "ppt/slideLayouts/slideLayout1.xml");
    assert.equal(hits[0].object.source.format, "pptx");
    assert.equal(hits[0].object.source.kind, "shape");
    assert.equal(hits[0].object.source.shapeId, 11);
    assert.equal(hits[0].object.source.mapping, "exact");
    assert.deepEqual(document.diagnostics(), []);
  } finally {
    document?.close();
    engine.close();
  }
});

test("replaces a layout placeholder with a slide picture using inherited bounds", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const drawCalls = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    drawImage(...args) { drawCalls.push(args); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 1, height: 1, close() {} }),
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({
    wasm,
    execution: "inline",
    limits: { documentObjects: 1 },
  });
  let document;
  let frame;

  try {
    document = await engine.open(pptxWithInheritedPicturePlaceholder);
    const objects = await document.listObjects({ unitIndex: 0 });
    const outsideObjects = await document.listObjects({
      unitIndex: 0,
      viewport: { x: 0, y: 0, width: 50, height: 50 },
    });
    frame = await document.render({ unitIndex: 0 });

    assert.equal(objects.length, 1);
    assert.equal(outsideObjects.length, 0);
    assert.equal(objects[0].type, "image");
    assert.deepEqual(objects[0].bounds, { x: 100, y: 120, width: 300, height: 200 });
    assert.equal(objects[0].source.part, "ppt/slides/slide1.xml");
    assert.equal(objects[0].source.shapeId, 31);
    assert.equal(frame.renderedObjectCount, 1);
    assert.deepEqual(drawCalls[0].slice(1), [100, 120, 300, 200]);
    assert.deepEqual(frame.diagnostics, []);
    assert.deepEqual(document.diagnostics(), []);
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

test("ignores effect transforms when reading PPTX picture geometry", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithPictureEffectTransform);
    const images = await document.listObjects({ unitIndex: 0, types: ["image"] });
    const slideImage = images.find(({ source }) => source.part === "ppt/slides/slide1.xml");

    assert.ok(slideImage);
    assert.deepEqual(slideImage.bounds, { x: 144, y: 96, width: 192, height: 96 });
    assert.equal(slideImage.source.shapeId, 31);
  } finally {
    document?.close();
    engine.close();
  }
});

test("inherits a layout placeholder transform for a slide picture", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithRotatedInheritedPicturePlaceholder);
    const hits = await document.hitTest({ unitIndex: 0, x: 250, y: 80 });

    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.type, "image");
    assert.equal(hits[0].object.source.part, "ppt/slides/slide1.xml");
    assert.equal(hits[0].object.source.shapeId, 31);
  } finally {
    document?.close();
    engine.close();
  }
});

test("uses slide order when multiple picture placeholders replace layout objects", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithReorderedPicturePlaceholders);
    const hits = await document.hitTest({ unitIndex: 0, x: 150, y: 150, limit: 2 });

    assert.deepEqual(hits.map(({ object }) => object.source.shapeId), [33, 35]);
  } finally {
    document?.close();
    engine.close();
  }
});

test("rejects incomplete direct PPTX picture transform coordinates", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });

  try {
    await assert.rejects(
      engine.open(pptxWithIncompletePictureTransform),
      (error) => {
        assert.equal(error.code, "FORMAT_INVALID");
        assert.equal(error.diagnostics[0].message, "picture transform is missing x");
        return true;
      },
    );
  } finally {
    engine.close();
  }
});

test("keeps image reuse valid after a shape replaces a picture placeholder", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithPicturePlaceholderReplacedByShape);
    const slideObjects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "ppt/slides/slide1.xml");

    assert.deepEqual(slideObjects.map(({ type }) => type), ["shape", "image"]);
    assert.deepEqual(slideObjects[1].bounds, { x: 400, y: 200, width: 200, height: 100 });
    assert.equal(slideObjects[1].source.shapeId, 32);
  } finally {
    document?.close();
    engine.close();
  }
});

test("keeps an embedded SVG picture renderable through the public PPTX document API", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithSvgPicture);
    const hits = await document.hitTest({ unitIndex: 0, x: 250, y: 120, limit: 1 });

    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.type, "image");
    assert.equal(hits[0].object.source.part, "ppt/slides/slide1.xml");
    assert.equal(hits[0].object.source.shapeId, 31);
    assert.deepEqual(document.diagnostics(), []);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders an embedded SVG through the isolated Office image codec", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const workerDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const workerMessages = [];
  let nativeDecodeCalls = 0;
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  class FakeCodecWorker {
    onmessage;
    onerror;
    onmessageerror;
    postMessage(message) {
      workerMessages.push(message);
      queueMicrotask(() => this.onmessage?.({
        data: {
          id: message.id,
          ok: true,
          bitmap: { width: 2, height: 1, close() {} },
          approximate: false,
        },
      }));
    }
    terminate() {}
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => {
      nativeDecodeCalls += 1;
      return { width: 1, height: 1, close() {} };
    },
  });
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: FakeCodecWorker });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(pptxWithSvgPicture);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(frame.renderedObjectCount, 4);
    assert.deepEqual(frame.diagnostics, []);
    assert.equal(nativeDecodeCalls, 1, "reused raster resources decode once");
    assert.equal(workerMessages.length, 1);
    assert.equal(workerMessages[0].format, "svg");
    assert.equal(workerMessages[0].mediaType, "image/svg+xml");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
    if (workerDescriptor) Object.defineProperty(globalThis, "Worker", workerDescriptor);
    else delete globalThis.Worker;
  }
});

test("renders the OOXML raster fallback when preferred SVG decoding fails", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const workerDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  let nativeDecodeCalls = 0;
  let nativeBitmapCloseCalls = 0;
  let svgDecodeCalls = 0;
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {}, drawImage() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  class FailingSvgWorker {
    onmessage;
    onerror;
    onmessageerror;
    postMessage(message) {
      svgDecodeCalls += 1;
      queueMicrotask(() => this.onmessage?.({
        data: {
          id: message.id,
          ok: false,
          code: "IMAGE_EXTERNAL_RESOURCE_BLOCKED",
          message: "The SVG references an external resource",
        },
      }));
    }
    terminate() {}
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => {
      nativeDecodeCalls += 1;
      return {
        width: 1,
        height: 1,
        close() { nativeBitmapCloseCalls += 1; },
      };
    },
  });
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: FailingSvgWorker });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({
    wasm,
    execution: "inline",
    limits: { imagePixels: 5, totalImagePixels: 5 },
  });
  let document;
  let firstFrame;
  let secondFrame;

  try {
    document = await engine.open(pptxWithSvgPicture);
    firstFrame = await document.render({ unitIndex: 0 });
    secondFrame = await document.render({ unitIndex: 0 });

    assert.equal(firstFrame.renderedObjectCount, 4, JSON.stringify(firstFrame.diagnostics));
    assert.deepEqual(firstFrame.diagnostics, [{
      code: "IMAGE_FALLBACK_USED",
      severity: "warning",
      fidelity: "approximate",
      phase: "render",
      message: "The preferred image/svg+xml image could not be decoded; its embedded image/png fallback was rendered",
      objectId: "object:3",
      part: "ppt/slides/slide1.xml",
      details: {
        preferredMediaType: "image/svg+xml",
        fallbackMediaType: "image/png",
      },
    }]);
    assert.deepEqual(secondFrame.diagnostics, firstFrame.diagnostics);
    assert.equal(svgDecodeCalls, 1);
    assert.equal(nativeDecodeCalls, 2, "fallback resources are reused across objects and renders");

    firstFrame.bitmap.close();
    firstFrame = undefined;
    secondFrame.bitmap.close();
    secondFrame = undefined;
    document.close();
    document = undefined;
    await Promise.resolve();
    assert.equal(nativeBitmapCloseCalls, 2);
  } finally {
    firstFrame?.bitmap.close();
    secondFrame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
    if (decodeDescriptor) Object.defineProperty(globalThis, "createImageBitmap", decodeDescriptor);
    else delete globalThis.createImageBitmap;
    if (workerDescriptor) Object.defineProperty(globalThis, "Worker", workerDescriptor);
    else delete globalThis.Worker;
  }
});

test("inherits an embedded raster background from a PPTX slide layout", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithLayoutPicture);
    const hits = await document.hitTest({ unitIndex: 0, x: 12, y: 12 });

    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.type, "image");
    assert.deepEqual(hits[0].object.bounds, { x: 0, y: 0, width: 960, height: 720 });
    assert.equal(hits[0].object.source.part, "ppt/slideLayouts/slideLayout1.xml");
    assert.equal(hits[0].object.source.format, "pptx");
    assert.equal(hits[0].object.source.kind, "shape");
    assert.equal(hits[0].object.source.mapping, "derived");
  } finally {
    document?.close();
    engine.close();
  }
});

test("resolves a theme image fill referenced by a PPTX master background", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const drawCalls = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    beginPath() {}, rect() {}, fill() {}, clip() {},
    drawImage(...args) { drawCalls.push(args); },
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 1, height: 1, close() {} }),
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;

  try {
    document = await engine.open(pptxWithThemeBackground);
    const objects = await document.listObjects({ unitIndex: 0 });
    const background = objects.find(({ source }) => (
      source.part === "ppt/slideMasters/slideMaster1.xml" && source.mapping === "derived"
    ));
    frame = await document.render({ unitIndex: 0 });

    assert.equal(background?.type, "shape");
    assert.deepEqual(background?.bounds, { x: 0, y: 0, width: 960, height: 720 });
    assert.equal(drawCalls.length, 1);
    assert.equal(objects.some(({ source }) => source.shapeId === 21), false);
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

test("renders a supported slide pattern background without falling back to the layout image", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithUnsupportedSlideBackground);
    const hits = await document.hitTest({ unitIndex: 0, x: 12, y: 12 });
    assert.equal(hits.length, 1);
    assert.equal(hits[0].object.type, "shape");
    assert.equal(hits[0].object.source.part, "ppt/slides/slide1.xml");
    assert.equal(hits[0].object.source.mapping, "derived");
    assert.equal(hits.some(({ object }) => object.type === "image"), false);
    assert.equal(document.diagnostics().some(({ code, part }) => (
      code === "UNSUPPORTED_FEATURE" && part === "ppt/slides/slide1.xml"
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("composes PPTX slide, layout, and master pictures in visible z-order", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;

  try {
    document = await engine.open(pptxWithLayoutPicture);
    const hits = await document.hitTest({ unitIndex: 0, x: 168, y: 120 });
    const pictures = hits.filter(({ object }) => object.source.mapping === "exact");

    assert.deepEqual(
      pictures.map(({ object }) => object.source.part),
      [
        "ppt/slides/slide1.xml",
        "ppt/slideLayouts/slideLayout1.xml",
        "ppt/slideMasters/slideMaster1.xml",
      ],
    );
    assert.deepEqual(
      pictures.map(({ object }) => object.source.shapeId),
      [31, 11, 21],
    );
    assert.deepEqual(
      pictures.map(({ object }) => object.type),
      ["image", "image", "image"],
    );
    assert.equal(hits.at(-1).object.source.part, "ppt/slideLayouts/slideLayout1.xml");
    assert.equal(hits.at(-1).object.source.mapping, "derived");
    assert.deepEqual(hits.at(-1).object.bounds, { x: 0, y: 0, width: 960, height: 720 });
  } finally {
    document?.close();
    engine.close();
  }
});

test("opens ODP bytes through its native ODF model with element mapping", async () => {
  const engine = await createOdfEngine();
  const document = await engine.open(odp);

  assert.equal(document.info.kind, "presentation");
  assert.equal(document.info.format, "odp");
  assert.deepEqual(document.info.units, [
    {
      type: "slide",
      index: 0,
      id: "unit:0",
      name: "Slide 1",
      width: 960,
      height: 720,
      sourcePart: "content.xml",
      slideNumber: 1,
      hidden: false,
    },
  ]);

  const hits = await document.hitTest({ unitIndex: 0, x: 120, y: 120 });
  assert.equal(hits.length, 1);
  assert.equal(hits[0].object.type, "text-box");
  assert.equal(hits[0].object.text, "Hello open document");
  assert.equal(hits[0].object.source.format, "odp");
  assert.equal(hits[0].object.source.kind, "element");
  assert.equal(hits[0].object.source.elementId, "greeting");
  assert.equal(hits[0].object.source.mapping, "exact");
  assert.deepEqual(hits[0].object.bounds, { x: 96, y: 96, width: 384, height: 96 });

  document.close();
  engine.close();
});

test("opens FODP flat XML through the shared ODP model", async () => {
  const engine = await createOdfEngine();
  const document = await engine.open(fodp);

  try {
    assert.equal(document.info.kind, "presentation");
    assert.equal(document.info.format, "odp");
    assert.equal(document.info.units[0].name, "Flat Slide");
    const hits = await document.hitTest({ unitIndex: 0, x: 120, y: 120 });
    assert.equal(hits[0].object.source.elementId, "flat-greeting");
  } finally {
    document.close();
    engine.close();
  }
});

test("opens FODS flat XML through the shared ODS model", async () => {
  const engine = await createOdfEngine();
  const document = await engine.open(fods);

  try {
    assert.equal(document.info.kind, "spreadsheet");
    assert.equal(document.info.format, "ods");
    assert.equal(document.info.units[0].name, "Flat Sheet");
    assert.equal((await document.listObjects({ unitIndex: 0, textOnly: true }))[0].text, "Hello FODS");
  } finally {
    document.close();
    engine.close();
  }
});

test("feeds browser font metrics back into ODP table layout", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  class MeasureCanvas {
    getContext() {
      return {
        font: "",
        measureText(text) {
          return {
            width: text.length * Number(this.font.match(/([\d.]+)px/u)[1]),
            fontBoundingBoxAscent: 800,
            fontBoundingBoxDescent: 200,
          };
        },
      };
    }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", {
    configurable: true,
    value: MeasureCanvas,
  });
  const engine = await createOdfEngine();
  let document;

  try {
    document = await engine.open(odpWithMeasuredTable);
    const objects = await document.listObjects({ unitIndex: 0 });
    const cell = objects.find(({ type }) => type === "cell");

    assert.ok(cell);
    assert.ok(cell.bounds.height > 25, `height=${cell.bounds.height}`);
  } finally {
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("returns embedded ODP audio through the public render API", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
  };
  class FakeCanvas {
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const engine = await createOdfEngine();
  let document;
  let frame;

  try {
    document = await engine.open(odpWithEmbeddedAudio);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(frame.media.length, 1);
    assert.deepEqual(frame.media[0], {
      objectId: "object:0",
      kind: "audio",
      mediaType: "audio/wav",
      bytes: embeddedWav,
      bounds: { x: 96, y: 96, width: 96, height: 96 },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    });
    assert.equal(Object.isFrozen(frame.media), true);
    assert.equal(Object.isFrozen(frame.media[0]), true);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("opens XLSX bytes with shared-string cells and A1 source mapping", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(xlsx);

  assert.equal(document.info.kind, "spreadsheet");
  assert.equal(document.info.format, "xlsx");
  assert.deepEqual(document.info.units, [
    { type: "sheet", index: 0, id: "unit:0", name: "Sheet1", width: 122, height: 48, rows: 2, columns: 2, frozenRows: 0, frozenColumns: 0, frozenWidth: 0, frozenHeight: 0, rowAxis: { defaultSize: 24, spans: [] }, columnAxis: { defaultSize: 61, spans: [] }, showGridLines: true },
  ]);

  const hits = await document.hitTest({ unitIndex: 0, x: 10, y: 10 });
  assert.equal(hits.length, 1);
  assert.equal(hits[0].object.type, "cell");
  assert.equal(hits[0].object.text, "Hello sheet");
  assert.equal(hits[0].object.source.format, "xlsx");
  assert.equal(hits[0].object.source.kind, "cell");
  assert.equal(hits[0].object.source.sheetName, "Sheet1");
  assert.equal(hits[0].object.source.address, "A1");
  assert.deepEqual(hits[0].object.bounds, { x: 0, y: 0, width: 61, height: 24 });

  document.close();
  engine.close();
});

test("opens ODT bytes with page layout and text-range source mapping", async () => {
  const engine = await createOdfEngine();
  const document = await engine.open(odt);

  assert.equal(document.info.kind, "text");
  assert.equal(document.info.format, "odt");
  assert.deepEqual(document.info.units, [
    { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 816, height: 1056 },
  ]);

  const hits = await document.hitTest({ unitIndex: 0, x: 100, y: 100 });
  assert.equal(hits.length, 1);
  assert.equal(hits[0].object.type, "paragraph");
  assert.equal(hits[0].object.text, "Hello text document");
  assert.equal(hits[0].object.source.format, "odt");
  assert.equal(hits[0].object.source.kind, "text-range");
  assert.deepEqual(hits[0].object.source.textRange, [0, 19]);
  assert.deepEqual(hits[0].object.bounds, { x: 96, y: 96, width: 624, height: Math.fround(16 * 1.2) });

  document.close();
  engine.close();
});

test("renders embedded ODT charts at their character-anchor flow position", async () => {
  const engine = await createOdfEngine();
  let document;
  try {
    document = await engine.open(odtWithCharacterAnchoredChart);
    const objects = await document.listObjects({ unitIndex: 0 });
    const paragraph = objects.find(({ text }) => text === "Paragraph before the chart");
    const frame = objects.find(({ id }) => id === "odt:frame:0:visual");
    const chartObjects = objects.filter(({ source }) => source.part === "Object 1/content.xml");

    assert.equal(chartObjects.filter(({ type }) => type === "shape").length, 28);
    assert.deepEqual(
      chartObjects.filter(({ type }) => type === "text-box").map(({ text }) => text),
      ["Row 1", "Row 2", "Row 3", "Row 4", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "Column 1", "Column 2", "Column 3"],
    );
    assert.ok(chartObjects.every(({ parentId }) => parentId === frame?.id));
    assert.ok(frame.bounds.y >= paragraph.bounds.y + paragraph.bounds.height);
  } finally {
    document?.close();
    engine.close();
  }
});

test("isolates a forbidden declaration in an embedded ODT object", async () => {
  const engine = await createOdfEngine();
  let document;
  try {
    document = await engine.open(odtWithLegacyMathDtd);
    assert.equal(document.info.format, "odt");
    assert.ok((await document.listObjects({ unitIndex: 0 })).length > 0);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders an ODT form text box control", async () => {
  const engine = await createOdfEngine();
  let document;
  try {
    document = await engine.open(odtWithFormTextBox);
    const control = (await document.listObjects({ unitIndex: 0 }))
      .find(({ id }) => id === "odt:frame:0:visual");
    assert.ok(control);
    assert.equal(control.type, "shape");
    assert.match(control.source.path, /draw:control/);
    assert.ok(control.bounds.width > 150 && control.bounds.height > 110);
  } finally {
    document?.close();
    engine.close();
  }
});

test("opens an ODT table with source-mapped cells and public ancestor traversal", async () => {
  const engine = await createOdfEngine();
  let document;

  try {
    document = await engine.open(odt);
    const hits = await document.hitTest({ unitIndex: 0, x: 500, y: 150 });

    assert.equal(hits.length, 2);
    const cell = hits[0].object;
    assert.equal(cell.type, "cell");
    assert.equal(cell.text, "B2");
    // A 19.2px line plus the existing 4px top/bottom padding must fit each row.
    const rowHeight = Math.fround(16 * 1.2 + 8);
    const tableY = Math.fround(96 + 16 * 1.2);
    assert.deepEqual(cell.bounds, { x: 408, y: Math.fround(tableY + rowHeight), width: 312, height: rowHeight });
    assert.equal(cell.source.format, "odt");
    assert.equal(cell.source.kind, "table-cell");
    assert.equal(cell.source.elementId, "b2");
    assert.equal(cell.source.path, "/office:document-content/office:body/office:text/table:table[1]/table:table-row[2]/table:table-cell[2]");
    assert.equal(cell.source.row, 1);
    assert.equal(cell.source.column, 1);
    assert.deepEqual(cell.source.textRange, [0, 2]);
    assert.equal(cell.source.mapping, "exact");

    assert.equal(hits[0].ancestors.length, 1);
    const table = hits[0].ancestors[0];
    assert.equal(table.type, "table");
    assert.equal(cell.parentId, table.id);
    assert.deepEqual(table.bounds, {
      x: 96, y: tableY, width: 624,
      height: Math.fround(Math.fround(tableY + rowHeight) + rowHeight) - tableY,
    });
    assert.equal(table.source.format, "odt");
    assert.equal(table.source.kind, "element");
    assert.equal(table.source.elementId, "matrix");
    assert.equal(table.source.path, "/office:document-content/office:body/office:text/table:table[1]");
    assert.equal(await document.getObject(table.id), table);
    assert.equal(hits[1].object, table);
  } finally {
    document?.close();
    engine.close();
  }
});

test("opens ODS bytes with native table coordinates and cell mapping", async () => {
  const engine = await createOdfEngine();
  const document = await engine.open(ods);

  assert.equal(document.info.kind, "spreadsheet");
  assert.equal(document.info.format, "ods");
  assert.deepEqual(document.info.units, [
    { type: "sheet", index: 0, id: "unit:0", name: "Sheet1", width: 192, height: 24, rows: 1, columns: 2, frozenRows: 0, frozenColumns: 0, frozenWidth: 0, frozenHeight: 0, rowAxis: { defaultSize: 24, spans: [] }, columnAxis: { defaultSize: 96, spans: [] }, showGridLines: true },
  ]);

  const hits = await document.hitTest({ unitIndex: 0, x: 10, y: 10 });
  assert.equal(hits.length, 1);
  assert.equal(hits[0].object.type, "cell");
  assert.equal(hits[0].object.text, "Hello ODS");
  assert.equal(hits[0].object.source.format, "ods");
  assert.equal(hits[0].object.source.tableName, "Sheet1");
  assert.equal(hits[0].object.source.row, 0);
  assert.equal(hits[0].object.source.column, 0);
  assert.deepEqual(hits[0].object.bounds, { x: 0, y: 0, width: 96, height: 24 });

  document.close();
  engine.close();
});

test("opens DOCX bytes with bounded pagination and paragraph source ranges", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(docx);

  assert.equal(document.info.kind, "text");
  assert.equal(document.info.format, "docx");
  assert.deepEqual(document.info.units, [
    { type: "page", index: 0, id: "unit:0", name: "Page 1", width: 816, height: 1056 },
  ]);

  const hits = await document.hitTest({ unitIndex: 0, x: 100, y: 100 });
  assert.equal(hits.length, 1);
  assert.equal(hits[0].object.type, "paragraph");
  assert.equal(hits[0].object.text, "Hello DOCX");
  assert.equal(hits[0].object.source.format, "docx");
  assert.equal(hits[0].object.source.kind, "paragraph");
  assert.equal(hits[0].object.source.paragraphId, "0A1B2C3D");
  assert.equal(hits[0].object.source.paragraphIndex, 0);
  assert.deepEqual(hits[0].object.source.textRange, [0, 10]);
  assert.deepEqual(hits[0].object.bounds, { x: 96, y: 96, width: 624, height: 24 });

  const table = await document.getObject("docx:table:0:fragment:0");
  const cell = await document.getObject("docx:table:0:row:0:column:0:fragment:0");
  assert.equal(table?.source.kind, "table");
  assert.equal(cell?.source.kind, "table-cell");

  const approximateLayouts = document.diagnostics().filter((diagnostic) => (
    diagnostic.phase === "layout" && diagnostic.fidelity === "approximate"
  ));
  assert.ok(approximateLayouts.length > 0);
  assert.ok(approximateLayouts.some((diagnostic) => diagnostic.code === "APPROXIMATE_LAYOUT"));
  assert.ok(approximateLayouts.every((diagnostic) => (
    diagnostic.code === "APPROXIMATE_LAYOUT" || diagnostic.code === "FONT_METRICS_UNAVAILABLE"
  )));

  document.close();
  engine.close();
});

test("opens the supplied DOCX whose preferred table width is 150%", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const fixture = await readFile(
    new URL("./fixtures/visual-captable-pct.docx.base64", import.meta.url),
    "utf8",
  );
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(Buffer.from(fixture.trim(), "base64"));

  assert.equal(document.info.format, "docx");
  assert.equal(document.info.units.length, 1);
  assert.ok((await document.listObjects()).some(({ type }) => type === "table"));

  document.close();
  engine.close();
});

test("renders DOCX character scaling as authored glyph width", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const scales = [];
  const fills = [];
  const context = {
    save() {}, restore() {}, translate() {}, rotate() {}, transform() {},
    scale(x, y) { scales.push([x, y]); },
    fillRect() {}, strokeRect() {}, clearRect() {},
    beginPath() {}, moveTo() {}, lineTo() {}, rect() {}, closePath() {}, clip() {}, fill() {}, stroke() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, ellipse() {}, setLineDash() {},
    fillText(text) { fills.push(text); }, strokeText() {},
    measureText(text) { return { width: text.length * 10 }; },
    letterSpacing: "0px",
  };
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; context.canvas = this; }
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline", fontPolicy: "local-first" });
  let document;
  let frame;
  try {
    document = await engine.open(docxWithCharacterScale);
    frame = await document.render({ unitIndex: 0 });
    assert.ok(fills.includes("x"));
    assert.ok(
      scales.some(([x, y]) => Math.abs(x - 0.33) < 0.000_001 && y === 1),
      JSON.stringify(scales),
    );
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders DOCX and PPTX picture effects through the shared DrawingML model", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const decodeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const drawCalls = [];
  const shadows = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, fillRect() {},
    beginPath() {}, moveTo() {}, lineTo() {}, rect() {}, closePath() {}, clip() {}, fill() {}, stroke() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, ellipse() {}, fillText() {},
    measureText() { return { width: 0 }; },
    drawImage(...args) { drawCalls.push(args); },
    getTransform() { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; },
    setTransform() {},
    getImageData() { return { width: 1, height: 1, data: new Uint8ClampedArray([0, 0, 0, 255]) }; },
    putImageData() {},
    set shadowBlur(value) { shadows.push({ color: this.shadowColor, blur: value }); },
  };
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; context.canvas = this; }
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 1, height: 1, close() {} }),
  });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline", fontPolicy: "local-first" });
  let document;
  let frame;
  try {
    for (const [format, bytes] of [["DOCX", docxWithPictureGlow], ["PPTX", pptxWithPictureGlow]]) {
      drawCalls.length = 0;
      shadows.length = 0;
      document = await engine.open(bytes);
      frame = await document.render({ unitIndex: 0 });
      assert.ok(drawCalls.length >= 2, `${format} missing glow image composition: ${drawCalls.length}`);
      assert.ok(
        shadows.some(({ color, blur }) => color === "rgba(255, 0, 0, 1)" && blur === 1),
        `${format} missing authored glow radius: ${JSON.stringify(shadows)}`,
      );
      frame.bitmap.close();
      frame = undefined;
      document.close();
      document = undefined;
    }
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

test("identifies a DOCX main part through an OPC default content type", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithDefaultMainContentType);
    assert.equal(document.info.format, "docx");
    assert.equal((await document.searchText({ query: "Default content type" })).length, 1);
  } finally {
    document?.close();
    engine.close();
  }
});

test("bounds Word-compatible signed and imprecise DOCX measures", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithWordCompatibilityMeasures);
    assert.equal((await document.searchText({ query: "Signed measures" })).length, 1);
    assert.equal(document.diagnostics().some(({ fidelity }) => fidelity === "approximate"), true);
  } finally {
    document?.close();
    engine.close();
  }
});

test("ignores bounded non-display DOCX metadata text", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithNonDisplayMetadataText);
    assert.equal((await document.searchText({ query: "Visible text" })).length, 1);
    assert.equal((await document.searchText({ query: "opaque metadata" })).length, 0);
    assert.equal(document.diagnostics().some(({ message }) => message === "unsupported DOCX text outside a run was ignored"), true);
  } finally {
    document?.close();
    engine.close();
  }
});

test("opens DOCX drawing text boxes without treating their paragraphs as invalid nesting", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  const document = await engine.open(docxWithDrawingTextBox);

  const textBox = await document.getObject("docx:drawing:0:7");
  assert.equal(textBox?.type, "text-box");
  assert.equal(textBox?.text, "Basic information\nName:\nJu Wang\nGender:\nMale\nYear of birth:\n1984\nExperience:\n15 years");
  assert.equal(textBox?.source.format, "docx");
  assert.equal(textBox?.source.kind, "drawing");
  assert.equal(textBox?.source.drawingId, 7);

  const objects = await document.listObjects({ unitIndex: 0 });
  const heading = objects.find((object) => object.text === "Basic information");
  const table = objects.find((object) => object.type === "table" && object.parentId === textBox.id);
  const overview = objects.find((object) => object.text === "Overview");
  assert.equal(heading?.parentId, textBox.id);
  assert.ok(heading.bounds.height >= 26);
  assert.ok(table);
  assert.ok(overview.bounds.y >= textBox.bounds.y + textBox.bounds.height);

  document.close();
  engine.close();
});

test("renders locked DOCX canvas shapes through shared DrawingML geometry", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithLockedCanvas);
    const objects = await document.listObjects({ unitIndex: 0 });

    assert.ok(
      objects.some(({ type, text }) => type === "text-box" && text === "Box"),
      JSON.stringify(objects.map(({ id, parentId, type, text }) => ({ id, parentId, type, text }))),
    );
    assert.equal(document.diagnostics().some(({ message }) => (
      message === "nested content in an unsupported DOCX drawing was omitted"
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders supplied DOCX chart titles and unsplit marker legends", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const fills = [];
  const context = {
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {},
    fillRect() {}, strokeRect() {}, clearRect() {}, beginPath() {}, moveTo() {}, lineTo() {},
    rect() {}, closePath() {}, clip() {}, fill() {}, stroke() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, ellipse() {}, setLineDash() {}, strokeText() {},
    fillText(text, x, y) { fills.push({ text, x, y }); },
    measureText(text) {
      const size = Number(/([\d.]+)px/u.exec(this.font)?.[1] ?? 12);
      return { width: [...text].reduce((width, character) => (
        width + size * (character.charCodeAt(0) < 128 ? 0.52 : 1)
      ), 0) };
    },
    letterSpacing: "0px",
  };
  class FakeCanvas {
    constructor(width, height) { this.width = width; this.height = height; context.canvas = this; }
    getContext() { return context; }
    transferToImageBitmap() { return { close() {} }; }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", { configurable: true, value: FakeCanvas });
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/chart-original.docx", import.meta.url)));
    const objects = await document.listObjects({ unitIndex: 0 });
    for (const text of ["折线统计图", "人数", "日期"]) {
      assert.ok(objects.some((object) => object.text === text), `missing ${text}`);
    }
    assert.equal(objects.filter((object) => object.id.includes(":chart:legend-marker-")).length, 3);
    frame = await document.render({ unitIndex: 0 });
    const legends = fills.flatMap((fill, index) => fill.text === "测试折线" ? [index] : []);
    assert.equal(legends.length, 3);
    for (const [series, index] of legends.entries()) {
      assert.equal(fills[index + 1].text, String(series + 1));
      assert.equal(fills[index + 1].y, fills[index].y, "legend digit must remain on the same line");
    }
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders a cached DOCX bar chart as source-mapped drawing children", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithCachedBarChart);
    const objects = await document.listObjects({ unitIndex: 0 });
    const chartObjects = objects.filter(({ source }) => source.part === "word/charts/chart1.xml");
    const chartParent = objects.find(({ id }) => id === chartObjects[0]?.parentId);

    assert.equal(chartParent?.type, "group");
    assert.ok(chartObjects.some(({ type, source }) => (
      type === "shape" && source.row === 0 && source.column === 0
    )));
    assert.ok(chartObjects.some(({ text }) => text === "Column 1"));
    assert.equal(document.diagnostics().some(({ message }) => (
      message.includes("drawing has no embedded image relationship")
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders each cached DOCX 3D bar as source-mapped front, top, and side faces", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithCached3dBarChart);
    const objects = await document.listObjects({ unitIndex: 0 });
    const facesByCategory = objects
      .filter(({ type, source }) => (
        type === "shape"
          && source.part === "word/charts/chart1.xml"
          && source.row === 0
          && source.column !== undefined
      ))
      .reduce((counts, { source }) => counts.set(
        source.column,
        (counts.get(source.column) ?? 0) + 1,
      ), new Map());

    assert.deepEqual([...facesByCategory], [[0, 3], [1, 3]]);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a DOCX 3D chart data table without flattening its bars", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithCached3dBarChartAndDataTable);
    const objects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "word/charts/chart1.xml");

    assert.ok(objects.some(({ id, text }) => id.includes(":chart:data-table-") && text === "9.1"));
    assert.ok(objects.some(({ id, text }) => id.includes(":chart:data-table-") && text === "2.4"));
    assert.ok(objects.some(({ id, text, bounds }) => (
      id.includes(":chart:data-table-") && text === "Column 1" && bounds.width >= 40
    )));
    assert.deepEqual(objects
      .filter(({ id, type }) => id.includes(":chart:bar-0-") && type === "shape")
      .map(({ source }) => source.column), [0, 0, 0, 1, 1, 1]);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders cached DOCX 3D pie slices, depth, title, and labels", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithCached3dPieChart);
    const chartObjects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "word/charts/chart1.xml");

    assert.ok(chartObjects.filter(({ type }) => type === "shape").length >= 10);
    assert.ok(["Grade 6 WL Enrollment 2015-16", "Immersion\n6%", "No WL\n42%"]
      .every((text) => chartObjects.some((object) => object.text === text)));
    assert.equal(document.diagnostics().some(({ message }) => (
      message.includes("unsupported DOCX chart type")
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("preserves the original DOCX Symbol bullet font in public font runs", async () => {
  const engine = await createOfficeEngine({
    wasm: await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)),
    execution: "inline",
  });
  let document;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/word-symbol-original.docx", import.meta.url)));
    const object = (await document.listObjects({ unitIndex: 0 }))
      .find(({ text }) => text?.startsWith("Bullet (Alt+0183)"));
    assert.equal(object.text, "Bullet (Alt+0183)\t•");
    const position = object.text.indexOf("•");
    assert.equal(object.fontRuns.find(({ start, end }) => start <= position && position < end).authoredFamily, "Symbol");
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a DOCX pie legend with inherited and overridden point colors", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithInheritedPiePointColor);
    const chartObjects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "word/charts/chart1.xml");

    assert.equal(chartObjects.filter(({ id }) => id.includes("pie-top-")).length, 4);
    assert.equal(chartObjects.filter(({ id }) => id.includes("pie-legend-key-")).length, 4);
    assert.ok(["Row 1", "Row 2", "Row 3", "Row 4"]
      .every((text) => chartObjects.some((object) => object.text === text)));
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders percent-stacked DOCX lines with up/down bars", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithCachedUpDownBarsChart);
    const chartObjects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "word/charts/chart1.xml");

    assert.equal(chartObjects.filter(({ id }) => id.includes("up-down-bar-")).length, 4);
    assert.ok(["Category 1", "Category 4", "Series 1", "Series 3"]
      .every((text) => chartObjects.some((object) => object.text === text)));
    assert.equal(document.diagnostics().some(({ message }) => (
      message.includes("unsupported DOCX chart type")
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders shared OOXML x markers in DOCX line and line-marker scatter charts", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  try {
    for (const [bytes, markerId, lineId] of [
      [docxWithXMarkerChart(false), ":chart:marker-", ":chart:line-"],
      [docxWithXMarkerChart(true), ":chart:scatter-marker-", ":chart:scatter-line-"],
    ]) {
      const document = await engine.open(bytes);
      const objects = await document.listObjects({ unitIndex: 0 });
      assert.equal(objects.filter(({ id }) => id.includes(markerId)).length, 3);
      assert.equal(objects.filter(({ id }) => id.includes(lineId)).length, 1, "one joined path per series");
      assert.equal(document.diagnostics().some(({ message }) => (
        message.includes("unsupported DOCX chart type")
      )), false);
      document.close();
    }
  } finally {
    engine.close();
  }
});

test("renders cached DOCX cylinder bars with elliptical caps", async () => {
  const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  let ellipses = 0;
  const context = {
    globalAlpha: 1,
    font: "", letterSpacing: "", textBaseline: "", textAlign: "", direction: "ltr",
    save() {}, restore() {}, scale() {}, translate() {}, rotate() {}, transform() {}, fillRect() {},
    beginPath() {}, ellipse() { ellipses += 1; }, moveTo() {}, lineTo() {}, rect() {},
    bezierCurveTo() {}, quadraticCurveTo() {}, closePath() {}, fill() {}, stroke() {}, clip() {},
    fillText() {}, measureText() { return { width: 10 }; }, setLineDash() {},
    createLinearGradient() { return { addColorStop() {} }; },
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
    document = await engine.open(docxWithCachedCylinderChart);
    frame = await document.render({ unitIndex: 0 });

    assert.equal(ellipses, 2);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    if (canvasDescriptor) Object.defineProperty(globalThis, "OffscreenCanvas", canvasDescriptor);
    else delete globalThis.OffscreenCanvas;
  }
});

test("renders cached DOCX horizontal cones along the value axis", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithCachedHorizontalConeChart);
    const fronts = (await document.listObjects({ unitIndex: 0 })).filter(({ id, type }) => (
      type === "shape" && /:chart:bar-\d+-\d+$/u.test(id)
    ));

    assert.equal(fronts.length, 2);
    assert.ok(fronts.every(({ bounds }) => bounds.width > bounds.height));
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a cached DOCX percent-stacked area chart with data labels", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithCachedAreaChart);
    const chartObjects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "word/charts/chart1.xml");

    assert.equal(chartObjects.filter(({ id, type }) => (
      type === "shape" && id.includes(":chart:area-")
    )).length, 2);
    assert.ok(["32", "28", "12", "Day 1", "Day 2", "Series 1", "Series 2"]
      .every((text) => chartObjects.some((object) => object.text === text)));
    assert.equal(document.diagnostics().some(({ message }) => (
      message.includes("unsupported DOCX chart type")
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a DOCX ChartEx waterfall from semantic data instead of its preview image", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithChartExWaterfall);
    const objects = await document.listObjects({ unitIndex: 0 });
    const chartObjects = objects.filter(({ source }) => source.part === "word/charts/chartEx1.xml");
    const chartParent = objects.find(({ id }) => id === chartObjects[0]?.parentId);

    assert.ok(chartObjects.filter(({ type }) => type === "shape").length >= 4);
    assert.ok(chartObjects.some(({ text }) => text === "-20"));
    assert.ok(["Chart Title", "Increase", "Decrease", "Total", "1", "2", "3", "4", "130"]
      .every((text) => chartObjects.some((object) => object.text === text)), JSON.stringify(chartObjects));
    assert.equal(chartParent?.bounds.y, 96);
    assert.equal(objects.some(({ source }) => source.part === "word/media/preview.png"), false);
    assert.equal(document.diagnostics().some(({ message }) => /placeholder/u.test(message)), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("renders a cached DOCX radar chart through the shared DrawingML geometry", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithCachedRadarChart);
    const chartObjects = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.part === "word/charts/chart1.xml");

    assert.ok(chartObjects.filter(({ type }) => type === "shape").length >= 2);
    assert.ok(["Series 1", "Series 2", "North", "East", "South"]
      .every((text) => chartObjects.some((object) => object.text === text)));
    assert.equal(document.diagnostics().some(({ message }) => (
      message.includes("unsupported DOCX chart type")
    )), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("still rejects nested DOCX body paragraphs outside drawing text boxes", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });

  try {
    await assert.rejects(
      engine.open(docxWithInvalidNestedParagraph),
      (error) => {
        assert.equal(error.code, "FORMAT_INVALID");
        assert.equal(error.diagnostics[0].message, "nested DOCX paragraphs are invalid");
        return true;
      },
    );
  } finally {
    engine.close();
  }
});

test("lets an explicit DOCX section footer override the inherited footer", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithSectionFooterOverrides);
    assert.equal((await document.searchText({ query: "First footer" })).length > 0, true);
    assert.equal((await document.searchText({ query: "Second footer" })).length > 0, true);
  } finally {
    document?.close();
    engine.close();
  }
});

test("preserves nested DOCX tables and searchable text", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithNestedTable);
    assert.equal((await document.searchText({ query: "Nested cell" })).length, 1);
    const tables = (await document.listObjects()).filter(({ type }) => type === "table");
    assert.equal(tables.length, 2);
    assert.equal(tables[1].bounds.width, 200);
    assert.equal(document.diagnostics().some(({ message }) => message === "nested DOCX tables are flattened into their parent cell flow"), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("preserves the original eight-level complex DOCX table hierarchy", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/word-complex-nested-table.docx", import.meta.url)));
    const objects = await document.listObjects();
    const tables = objects.filter(({ type }) => type === "table");
    assert.equal(tables.length, 8);
    assert.ok(tables[0].bounds.height < 400, "nested table end markers must not add visible empty lines");
    assert.equal(objects.filter(({ type }) => type === "cell").length, 32);
    for (const table of tables.slice(1)) {
      const parent = objects.find(({ id }) => id === table.parentId);
      assert.equal(parent?.type, "cell");
      assert.ok(table.bounds.x >= parent.bounds.x);
      assert.ok(table.bounds.y >= parent.bounds.y);
      assert.ok(table.bounds.x + table.bounds.width <= parent.bounds.x + parent.bounds.width + 0.01);
      assert.ok(table.bounds.y + table.bounds.height <= parent.bounds.y + parent.bounds.height + 0.01);
    }
    for (let number = 1; number <= 8; number++) {
      assert.equal((await document.searchText({ query: String(number) })).length, 1);
    }
  } finally {
    document?.close();
    engine.close();
  }
});

test("accepts zero as a DOCX bullet numbering start", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithZeroNumberingStart);
    assert.equal((await document.searchText({ query: "Bullet item" })).length, 1);
  } finally {
    document?.close();
    engine.close();
  }
});

test("defaults a missing DOCX numbering ilvl to level zero", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithImplicitNumberingLevel);
    assert.equal((await document.searchText({ query: "Bullet item" })).length, 1);
    assert.equal(document.diagnostics().some(({ message }) => message === "DOCX numbering level without ilvl was interpreted as level zero"), true);
  } finally {
    document?.close();
    engine.close();
  }
});

test("omits a zero-size DOCX drawing without rejecting the document", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithZeroHeightHeaderDrawing);
    assert.equal((await document.searchText({ query: "Visible body" })).length, 1);
    assert.equal(document.diagnostics().some(({ message }) => message === "zero-size DOCX drawings are omitted"), true);
  } finally {
    document?.close();
    engine.close();
  }
});

test("does not reject legacy VML text-box paragraphs nested in a body paragraph", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithLegacyVmlTextBox);
    assert.equal((await document.searchText({ query: "Visible body" })).length, 1);
    assert.equal(document.diagnostics().some(({ message }) => message === "legacy VML DOCX drawings use a source-mapped static placeholder"), false);
  } finally {
    document?.close();
    engine.close();
  }
});

test("does not reject empty legacy VML image data without a relationship ID", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  let document;
  try {
    document = await engine.open(docxWithEmptyLegacyVmlImageData);
    assert.equal((await document.searchText({ query: "Visible body" })).length, 1);
    assert.equal(document.diagnostics().some(({ message }) => message === "legacy VML DOCX drawings use a source-mapped static placeholder"), false);
  } finally {
    document?.close();
    engine.close();
  }
});
